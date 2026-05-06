using Dm; // 达梦驱动命名空间
using GB_NewCadPlus_IV.DMDatabaseReader;
using GB_NewCadPlus_IV.UniFiedStandards;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 达梦数据库认证与组织架构管理服务 (最终兼容模式版)
    /// 解决“无效的表或视图名”终极方案：全局 Set Schema，移除所有冗余表前缀和双引号保护
    /// </summary>
    public class DMAuthService
    {
        #region 基础字段与连接

        private readonly string _server;
        private readonly string _port;
        private readonly string _dbUser;
        private readonly string _dbPassword;
        private readonly string _database = "CAD_SW_LIBRARY"; // 达梦模式名/数据库名



        /// <summary>
        /// 构造服务实例
        /// </summary>
        public DMAuthService(string server, string port, string user = "SYSDBA", string pwd = "SYSDBA")
        {
            _server = string.IsNullOrEmpty(server) ? "127.0.0.1" : server;
            _port = string.IsNullOrEmpty(port) ? "5236" : port;
            _dbUser = (user ?? "SYSDBA").ToUpper();
            _dbPassword = pwd;
        }

        /// <summary>
        /// 根据用户名读取角色字段（ROLE），若不存在或未激活返回 null。
        /// 返回值不做枚举限制，调用方负责按需解析（如包含 "admin" 则视为管理员）。
        /// </summary>
        /// <param name="username">要查询的用户名（大小写不敏感）</param>
        /// <returns>角色字符串或 null</returns>
        public string? GetUserRole(string username)
        {
            if (string.IsNullOrWhiteSpace(username)) return null;
            try
            {
                using (var conn = CreateOpenConnection())
                using (var cmd = conn.CreateCommand())
                {
                    // 使用 UPPER 比较以避免大小写差异导致匹配失败
                    cmd.CommandText = "SELECT ROLE FROM USERS WHERE UPPER(USERNAME) = :u AND IS_ACTIVE = 1";
                    AddParam(cmd, "u", username.Trim().ToUpperInvariant());
                    var obj = cmd.ExecuteScalar();
                    if (obj == null || obj == DBNull.Value) return null;
                    return Convert.ToString(obj);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetUserRole 查询失败: {ex.Message}");
                return null;
            }
        }

        private string ConnString() => $"Server={_server};Port={_port};User Id={_dbUser};Password={_dbPassword};Schema={_database};";

        /// <summary>
        /// 创建并初始化连接，并切换模式上下文
        /// </summary>
        private DmConnection CreateOpenConnection()
        {
            var conn = new DmConnection(ConnString());
            conn.Open();
            // 在当前连接环境中强制指派操作范畴，防止漂移
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SET SCHEMA {_database}";
                try { cmd.ExecuteNonQuery(); } catch { }
            }
            return conn;
        }

        /// <summary>
        /// 添加参数的辅助方法，简化代码并确保参数化查询（防止SQL注入）
        /// </summary>
        /// <param name="cmd">要添加参数的命令对象</param>
        /// <param name="name">参数名称</param>
        /// <param name="value">参数值</param>
        private static void AddParam(IDbCommand cmd, string name, object value)
        {
            var p = cmd.CreateParameter();
            p.ParameterName = name;
            p.Value = value ?? DBNull.Value;
            cmd.Parameters.Add(p);
        }

        /// <summary>
        /// 优先解析真正有数据的目标表名。
        /// 先检查当前 Schema，再检查 ALL_TABLES 中同名表，优先返回有数据的那个。
        /// </summary>
        private string ResolvePreferredTableName(DmConnection conn, string tableName, string preferredOwner = null)
        {
            var normalizedTableName = (tableName ?? string.Empty).Trim().ToUpperInvariant(); // 达梦数据库默认大写，且不区分大小写，但为了保险起见，统一转换为大写进行比较
            var candidates = new List<Tuple<string, string, int>>();//候选列表：Owner, QualifiedTableName, RowCount

            // 先把当前 Schema 下的同名表加入候选。
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT COUNT(1) FROM USER_TABLES WHERE TABLE_NAME = :t";// 直接查询 USER_TABLES 以避免权限问题导致的 ALL_TABLES 无法访问
                AddParam(cmd, "t", normalizedTableName); // 注意：达梦数据库的 USER_TABLES 视图只包含当前 Schema 的表，因此这里不需要再加 OWNER 条件了
                var existsInCurrentSchema = Convert.ToInt32(cmd.ExecuteScalar() ?? 0) > 0; // 先确认当前 Schema 是否存在同名表，避免后续查询行数时因表不存在而抛异常
                if (existsInCurrentSchema) // 如果当前 Schema 下存在同名表，则尝试获取行数并加入候选列表；如果不存在，则不加入，直接等待后续 ALL_TABLES 的结果
                {
                    // 注意：这里直接使用 normalizedTableName 作为 QualifiedTableName，因为在当前 Schema 下访问时不需要 OWNER 前缀，且达梦数据库默认不区分大小写，所以不需要加双引号保护。
                    candidates.Add(Tuple.Create(string.Empty, normalizedTableName, TryGetTableRowCount(conn, normalizedTableName)));
                }
            }

            // 再从 ALL_TABLES 中查找其它 Schema 的同名表。
            try
            {
                using (var cmd = conn.CreateCommand()) // 注意：这里不加 OWNER 条件，直接查询所有同名表，后续在 C# 端进行过滤和优先级判断，以避免权限问题导致的查询失败，同时也能兼容用户在不同 Schema 下创建同名表的情况。
                {
                    // 注意：达梦数据库的 ALL_TABLES 视图包含了所有 Schema 的表信息，因此这里直接查询 TABLE_NAME 即可，后续在 C# 端进行 OWNER 的过滤和优先级判断。
                    cmd.CommandText = "SELECT OWNER FROM ALL_TABLES WHERE TABLE_NAME = :t ORDER BY OWNER";
                    AddParam(cmd, "t", normalizedTableName); // 注意：达梦数据库的 ALL_TABLES 视图中的 TABLE_NAME 字段默认是大写的，因此这里直接使用 normalizedTableName 进行比较，以避免大小写不一致导致的匹配失败。
                    using (var reader = cmd.ExecuteReader()) // 直接读取 OWNER 列，后续在 C# 端构建完整的 QualifiedTableName 并进行优先级判断，以避免权限问题导致的查询失败，同时也能兼容用户在不同 Schema 下创建同名表的情况。
                    {
                        while (reader.Read())// 注意：这里直接使用 OWNER 列的值来构建 QualifiedTableName，因为在 ALL_TABLES 视图中 OWNER 列已经包含了表所属的 Schema 信息，后续在 C# 端进行优先级判断时也会使用这个 OWNER 值来判断是否匹配首选 Schema。
                        {
                            // 注意：这里直接使用 OWNER 列的值来构建 QualifiedTableName，因为在 ALL_TABLES 视图中 OWNER 列已经包含了表所属的 Schema 信息，后续在 C# 端进行优先级判断时也会使用这个 OWNER 值来判断是否匹配首选 Schema。
                            var owner = reader.IsDBNull(0) ? string.Empty : Convert.ToString(reader.GetValue(0));
                            // 注意：达梦数据库的 ALL_TABLES 视图中的 OWNER 列默认是大写的，因此这里直接使用 reader.GetString(0) 来获取 OWNER 的值，并且不需要再进行 ToUpper 操作了，因为之前在查询时已经使用 normalizedTableName 进行了大写比较，确保了匹配的一致性。
                            if (string.IsNullOrWhiteSpace(owner))
                            {
                                continue;
                            }
                            // 注意：这里直接使用 normalizedTableName 作为 QualifiedTableName，因为在访问时不需要 OWNER 前缀，且达梦数据库默认不区分大小写，所以不需要加双引号保护。同时，在后续的优先级判断中会使用 OWNER 来判断是否匹配首选 Schema。
                            var qualifiedName = owner + "." + normalizedTableName;
                            if (candidates.Any(c => string.Equals(c.Item2, qualifiedName, StringComparison.OrdinalIgnoreCase)))
                            {
                                continue;
                            }

                            candidates.Add(Tuple.Create(owner, qualifiedName, TryGetTableRowCount(conn, qualifiedName)));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 注意：这里捕获所有异常并记录日志，而不是让异常冒泡中断流程，因为在某些环境下用户可能没有权限访问 ALL_TABLES 视图，如果直接抛异常会导致整个功能无法使用。通过捕获异常并记录日志，我们可以在没有 ALL_TABLES 访问权限的环境下仍然正常工作，只不过无法优先选择有数据的表了。
                LogManager.Instance.LogInfo($"ResolvePreferredTableName: 查询 ALL_TABLES 失败，table={normalizedTableName}，{ex.Message}");
            }
            // 注意：这里不直接抛异常，而是继续使用当前 Schema 下的表（如果存在）或者后续 ALL_TABLES 查询到的表（如果有）进行优先级判断和选择，以确保在没有 ALL_TABLES 访问权限的环境下仍然能够正常工作。
            if (candidates.Count == 0)
            {
                // 既然 ALL_TABLES 无法访问或者没有找到任何同名表，那么我们就退回到直接使用默认表名的方式，虽然可能会因为权限问题导致后续访问失败，但至少不会因为找不到表而抛异常中断整个流程。
                LogManager.Instance.LogInfo($"ResolvePreferredTableName: 未找到 {normalizedTableName}，将直接使用默认表名。\n");
                return normalizedTableName;// 直接返回原始表名，后续访问时如果权限不足导致找不到表会自然抛异常，这样至少不会因为找不到表而抛异常中断整个流程。
            }

            // 先优先使用指定 Owner 且有数据的表。
            if (!string.IsNullOrWhiteSpace(preferredOwner))
            {
                var preferred = candidates.FirstOrDefault(c =>
                    string.Equals(c.Item1, preferredOwner, StringComparison.OrdinalIgnoreCase) && c.Item3 > 0);// 注意：这里直接使用 OWNER 列的值来判断是否匹配首选 Schema，因为在 ALL_TABLES 视图中 OWNER 列已经包含了表所属的 Schema 信息。同时，优先选择行数大于0的表，以确保我们选择的是有数据的表，而不是空表。
                if (preferred != null)
                {
                    LogManager.Instance.LogInfo($"ResolvePreferredTableName: {normalizedTableName} 命中首选 Schema={preferred.Item1}，Rows={preferred.Item3}");
                    return preferred.Item2;// 注意：这里直接返回 qualifiedName，因为在访问时不需要 OWNER 前缀，且达梦数据库默认不区分大小写，所以不需要加双引号保护。同时，在后续的访问中也会使用这个 qualifiedName 来访问表，以确保访问的一致性。
                }
            }

            // 其次优先使用有数据的表。
            var nonEmpty = candidates.OrderByDescending(c => c.Item3).FirstOrDefault(c => c.Item3 > 0);
            if (nonEmpty != null)
            {
                LogManager.Instance.LogInfo($"ResolvePreferredTableName: {normalizedTableName} 选中有数据的表 {nonEmpty.Item2}，Rows={nonEmpty.Item3}");
                return nonEmpty.Item2;
            }

            // 最后退回第一个候选。
            var fallback = candidates[0];
            LogManager.Instance.LogInfo($"ResolvePreferredTableName: {normalizedTableName} 所有候选都为空，退回 {fallback.Item2}");
            return fallback.Item2;
        }

        /// <summary>
        /// 尝试读取表行数，失败时返回 -1，避免因单个 Schema 无权限而中断整个流程。
        /// </summary>
        private int TryGetTableRowCount(DmConnection conn, string qualifiedTableName)
        {
            try
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = $"SELECT COUNT(1) FROM {qualifiedTableName}";
                    return Convert.ToInt32(cmd.ExecuteScalar() ?? 0);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"TryGetTableRowCount: 读取 {qualifiedTableName} 行数失败：{ex.Message}");
                return -1;
            }
        }

        /// <summary>
        /// 从限定表名中提取 Owner，例如 SYSDBA.CAD_CATEGORIES -> SYSDBA。
        /// </summary>
        private static string GetOwnerFromQualifiedTableName(string qualifiedTableName)
        {
            if (string.IsNullOrWhiteSpace(qualifiedTableName))// 注意：这里直接使用 string.IsNullOrWhiteSpace 来判断输入是否合法，因为在某些环境下可能会出现空格或者其他不可见字符导致的表名异常，如果直接使用 string.IsNullOrEmpty 可能无法正确识别这些情况。通过使用 string.IsNullOrWhiteSpace，我们可以更全面地判断输入是否合法，避免因为表名异常而导致后续访问失败。
            {
                return string.Empty;
            }
            // 注意：这里直接使用 '.' 来分割 Owner 和 TableName，因为在达梦数据库中，表的限定名通常是以 "OWNER.TABLE_NAME" 的形式出现的。同时，在后续的访问中也会使用这个 qualifiedName 来访问表，以确保访问的一致性。
            var index = qualifiedTableName.IndexOf('.');
            if (index <= 0)// 注意：这里直接使用 index <= 0 来判断是否存在有效的 Owner，因为在达梦数据库中，表的限定名通常是以 "OWNER.TABLE_NAME" 的形式出现的，如果没有 '.' 或者 '.' 在开头，那么就说明没有有效的 Owner 信息。
            {
                return string.Empty;// 注意：这里直接返回空字符串来表示没有 Owner 信息，因为在达梦数据库中，如果表名没有 Owner 前缀，那么默认就是当前 Schema 下的表，我们可以通过全局 Set Schema 来确保访问的一致性，而不需要在表名前加上默认的 Schema 前缀。
            }
            // 注意：这里直接使用 Substring 来提取 Owner 部分，因为在达梦数据库中，表的限定名通常是以 "OWNER.TABLE_NAME" 的形式出现的，我们可以通过 Substring 来提取 Owner 部分，同时在后续的访问中也会使用这个 qualifiedName 来访问表，以确保访问的一致性。
            return qualifiedTableName.Substring(0, index);
        }

        #endregion

        #region 表结构与注释初始化

        /// <summary>
        ///  确保分类、部门、用户表的存在。每当发生业务时调用本方法，它会实时验证物理表
        /// </summary>
        public void EnsureAllTablesExist()
        {
            EnsureTable("CAD_CATEGORIES", @"
                CREATE TABLE CAD_CATEGORIES (
                    ID INT IDENTITY(1,1) PRIMARY KEY,
                    NAME VARCHAR(200) NOT NULL,
                    DISPLAY_NAME VARCHAR(200),
                    SORT_ORDER INT DEFAULT 0,
                    CREATED_AT DATETIME DEFAULT CURRENT_TIMESTAMP
                )");

            EnsureTable("DEPARTMENTS", @"
                CREATE TABLE DEPARTMENTS (
                    ID INT IDENTITY(1,1) PRIMARY KEY,
                    CAD_CATEGORY_ID INT NULL,
                    NAME VARCHAR(200) NOT NULL,
                    DISPLAY_NAME VARCHAR(200),
                    DESCRIPTION TEXT,
                    MANAGER_USER_ID INT NULL,
                    SORT_ORDER INT DEFAULT 0,
                    IS_ACTIVE TINYINT DEFAULT 1,
                    CREATED_AT DATETIME DEFAULT CURRENT_TIMESTAMP,
                    UPDATED_AT DATETIME
                )");

            EnsureTable("USERS", @"
                CREATE TABLE USERS (
                    ID INT IDENTITY(1,1) PRIMARY KEY,
                    USERNAME VARCHAR(100) NOT NULL UNIQUE,
                    PASSWORD_HASH VARCHAR(512) NOT NULL,
                    SALT VARCHAR(64) NOT NULL,
                    DISPLAY_NAME VARCHAR(200),
                    DEPARTMENT_ID INT DEFAULT 0,
                    DEPARTMENT_NAME VARCHAR(200),
                    ROLE VARCHAR(64),
                    IS_ACTIVE TINYINT DEFAULT 1,
                    CREATED_AT DATETIME DEFAULT CURRENT_TIMESTAMP
                )");

            SyncTableComments();// 确保表注释同步，提升数据库自描述能力
        }

        /// <summary>
        /// 通用表创建工具：避免使用 DECLARE 等达梦复合语句，采用更稳妥的 C# 原子操作以确保持久化
        /// </summary>
        private void EnsureTable(string tableName, string createSql)
        {
            using (var conn = CreateOpenConnection())
            {
                bool tableExists = true;
                try
                {
                    // [降维打击]：只要 Select 指令抛出任何奇奇怪怪的“找不到表”错误，就当作表没建好
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = $"SELECT 1 FROM {tableName} WHERE 1=0";
                        cmd.ExecuteNonQuery();
                    }
                }
                catch
                {
                    tableExists = false;
                }

                if (!tableExists)
                {
                    try
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = createSql;
                            cmd.ExecuteNonQuery();
                        }
                        // 创建完毕后强行 COMMIT 提交，防止因异常关闭造成达梦数据库物理未落盘（回滚）
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = "COMMIT";
                            cmd.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogError($"建表 {tableName} 失败：{ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// 同步表注释（兼容最终模式：无前缀，无双引号），并且增加了对工艺图元相关表的注释同步，确保所有核心表都有清晰的说明，提升数据库自描述能力
        /// </summary>
        private void SyncTableComments()
        {
            using (var conn = CreateOpenConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
                BEGIN
                    EXECUTE IMMEDIATE 'COMMENT ON TABLE CAD_CATEGORIES IS ''CAD分类信息表''';
                    EXECUTE IMMEDIATE 'COMMENT ON TABLE DEPARTMENTS IS ''部门组织架构表''';
                    EXECUTE IMMEDIATE 'COMMENT ON TABLE USERS IS ''系统用户信息表''';
                END;";
                try { cmd.ExecuteNonQuery(); } catch { }
            }
        }

        /// <summary>
        /// 确认 CAD_CATEGORIES 分类表存在，不存在则创建（兼容最终模式：无前缀，无双引号）
        /// </summary>
        public void EnsureCategoriesTableExists()
        {
            using (var conn = CreateOpenConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $@"
                DECLARE V_CNT INT;
                BEGIN
                    SELECT COUNT(*) INTO V_CNT FROM USER_TABLES WHERE TABLE_NAME = 'CAD_CATEGORIES';
                    IF V_CNT = 0 THEN
                        EXECUTE IMMEDIATE 'CREATE TABLE {_database}.CAD_CATEGORIES (
                            ID INT IDENTITY(1,1) PRIMARY KEY,
                            NAME VARCHAR(200) NOT NULL,
                            DISPLAY_NAME VARCHAR(200),
                            SORT_ORDER INT DEFAULT 0,
                            CREATED_AT DATETIME DEFAULT CURRENT_TIMESTAMP
                        )';
                    END IF;
                END;";
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 确认部门表存在，不存在则创建（兼容最终模式：无前缀，无双引号）
        /// </summary>

        public void EnsureDepartmentsTableExists()
        {
            using (var conn = CreateOpenConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $@"
                DECLARE V_CNT INT;
                BEGIN
                    SELECT COUNT(*) INTO V_CNT FROM USER_TABLES WHERE TABLE_NAME = 'DEPARTMENTS';
                    IF V_CNT = 0 THEN
                        EXECUTE IMMEDIATE 'CREATE TABLE {_database}.DEPARTMENTS (
                            ID INT IDENTITY(1,1) PRIMARY KEY,
                            CAD_CATEGORY_ID INT NULL,
                            NAME VARCHAR(200) NOT NULL,
                            DISPLAY_NAME VARCHAR(200),
                            DESCRIPTION TEXT,
                            MANAGER_USER_ID INT NULL,
                            SORT_ORDER INT DEFAULT 0,
                            IS_ACTIVE TINYINT DEFAULT 1,
                            CREATED_AT DATETIME DEFAULT CURRENT_TIMESTAMP,
                            UPDATED_AT DATETIME
                        )';
                    END IF;
                END;";
                cmd.ExecuteNonQuery();
            }
        }
        /// <summary>
        /// 确认用户表存在，不存在则创建（兼容最终模式：无前缀，无双引号）
        /// </summary>
        public void EnsureUserTableExists()
        {
            using (var conn = CreateOpenConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $@"
                DECLARE V_CNT INT;
                BEGIN
                    SELECT COUNT(*) INTO V_CNT FROM USER_TABLES WHERE TABLE_NAME = 'USERS';
                    IF V_CNT = 0 THEN
                        EXECUTE IMMEDIATE 'CREATE TABLE {_database}.USERS (
                            ID INT IDENTITY(1,1) PRIMARY KEY,
                            USERNAME VARCHAR(100) NOT NULL UNIQUE,
                            PASSWORD_HASH VARCHAR(512) NOT NULL,
                            SALT VARCHAR(64) NOT NULL,
                            DISPLAY_NAME VARCHAR(200),
                            DEPARTMENT_ID INT DEFAULT 0,
                            DEPARTMENT_NAME VARCHAR(200),
                            ROLE VARCHAR(64),
                            IS_ACTIVE TINYINT DEFAULT 1,
                            CREATED_AT DATETIME DEFAULT CURRENT_TIMESTAMP
                        )';
                        EXECUTE IMMEDIATE 'CREATE INDEX IDX_USERS_DEPT ON {_database}.USERS(DEPARTMENT_ID)';
                    END IF;
                END;";
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 同步表注释（兼容最终模式：无前缀，无双引号）
        /// </summary>
        //private void SyncTableComments()
        //{
        //    using (var conn = CreateOpenConnection())
        //    using (var cmd = conn.CreateCommand())
        //    {
        //        cmd.CommandText = @"
        //        BEGIN
        //            EXECUTE IMMEDIATE 'COMMENT ON TABLE SYSDBA.CAD_CATEGORIES IS ''CAD分类信息表''';
        //            EXECUTE IMMEDIATE 'COMMENT ON TABLE SYSDBA.DEPARTMENTS IS ''部门组织架构表''';
        //            EXECUTE IMMEDIATE 'COMMENT ON TABLE SYSDBA.USERS IS ''系统用户信息表''';

        //            EXECUTE IMMEDIATE 'COMMENT ON TABLE SYSDBA.SW_CATEGORIES IS ''工艺图元分类表''';
        //            EXECUTE IMMEDIATE 'COMMENT ON TABLE SYSDBA.SW_GRAPHICS IS ''工艺图元图形定义表''';
        //            EXECUTE IMMEDIATE 'COMMENT ON TABLE SYSDBA.SW_SUBCATEGORIES IS ''工艺图元子分类表''';
        //        END;";
        //        try { cmd.ExecuteNonQuery(); } catch { /* 忽略个别环境注释失败 */ }
        //    }
        //}

        #endregion

        #region 健壮版业务逻辑 （去除 MERGE / TOP 等易异常语法）

        /// <summary>
        /// 验证用户身份，返回该用户的角色字符串（如 "admin"）或 null（用户不存在、未激活或发生错误）。用户名比较不区分大小写。
        /// </summary>
        public bool AuthenticateUser(string username, string password)
        {
            EnsureAllTablesExist(); // 确保安全
            if (string.IsNullOrEmpty(username)) return false;
            try
            {
                using (var conn = CreateOpenConnection())
                using (var cmd = conn.CreateCommand())
                {
                    // 移除 TOP 1 解决达梦特性报错，采用最基础标准的查询，利用 Reader 截断结果
                    cmd.CommandText = "SELECT PASSWORD_HASH, SALT FROM USERS WHERE USERNAME = :u AND IS_ACTIVE = 1";
                    AddParam(cmd, "u", username);
                    using (var r = cmd.ExecuteReader())
                    {
                        if (r.Read())
                        {
                            string dbHash = r.GetString(0);// 数据库中存储的密码哈希
                            string dbSalt = r.GetString(1);// 数据库中存储的盐值
                            //string dbPassword = r.GetString(2);// 数据库中存储的明文密码（兼容旧版本，已废弃，不再使用）
                            string passwordWithSalt = ComputeHash(password, dbSalt);// 将输入密码与盐值组合
                            if(passwordWithSalt==dbHash)
                            return true;// 密码验证是否成功
                        }
                    }
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 注册新用户
        /// </summary>
        public bool RegisterUser(string username, string password, int departmentId = 0, string departmentName = "")
        {
            EnsureAllTablesExist();
            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password)) return false;
            var salt = GenerateSalt();
            var hash = ComputeHash(password, salt);
            try
            {
                using (var conn = CreateOpenConnection())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "INSERT INTO USERS (USERNAME, PASSWORD_HASH, SALT, DEPARTMENT_ID, DEPARTMENT_NAME) VALUES (:u, :h, :s, :d, :dn)";
                    AddParam(cmd, "u", username); AddParam(cmd, "h", hash); AddParam(cmd, "s", salt);
                    AddParam(cmd, "d", departmentId); AddParam(cmd, "dn", departmentName);
                    return cmd.ExecuteNonQuery() > 0;
                }
            }
            catch { return false; }
        }

        /// <summary>
        /// 获取所有部门及其用户数
        /// </summary>
        public List<DepartmentModel> GetDepartmentsWithCounts()
        {
            EnsureAllTablesExist();
            // 先按正式部门表读取；如果部门表为空，再自动从 CAD_CATEGORIES 回填一次，最后仍为空则直接用分类表兜底，避免界面无数据。
            var list = LoadDepartmentsWithCountsCore();
            if (list.Count > 0)
            {
                return list;
            }

            LogManager.Instance.LogInfo("GetDepartmentsWithCounts: DEPARTMENTS 当前为空，开始尝试从 CAD_CATEGORIES 自动同步。");
            try
            {
                SyncDepartmentsFromCadCategories();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetDepartmentsWithCounts: 自动同步部门失败: {ex.Message}");
            }

            list = LoadDepartmentsWithCountsCore();
            if (list.Count > 0)
            {
                return list;
            }

            LogManager.Instance.LogInfo("GetDepartmentsWithCounts: 同步后部门表仍为空，改为直接读取 CAD_CATEGORIES 作为下拉框兜底数据源。");
            return LoadDepartmentsFromCategoriesFallback();
        }

        /// <summary>
        /// 从 DEPARTMENTS 正式表读取部门及人数。
        /// </summary>
        private List<DepartmentModel> LoadDepartmentsWithCountsCore()
        {
            var list = new List<DepartmentModel>();
            using (var conn = CreateOpenConnection())
            {
                var departmentsTableName = ResolvePreferredTableName(conn, "DEPARTMENTS", _dbUser);
                var usersTableName = ResolvePreferredTableName(conn, "USERS", _dbUser);
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = $@"SELECT D.ID, D.CAD_CATEGORY_ID, D.NAME, D.DISPLAY_NAME, D.DESCRIPTION, D.SORT_ORDER, D.IS_ACTIVE,
                                   (SELECT COUNT(1) FROM {usersTableName} U WHERE U.DEPARTMENT_ID = D.ID) AS USER_COUNT
                                   FROM {departmentsTableName} D ORDER BY D.SORT_ORDER";
                    LogManager.Instance.LogInfo($"LoadDepartmentsWithCountsCore: departmentsTable={departmentsTableName}, usersTable={usersTableName}");
                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            list.Add(new DepartmentModel
                            {
                                Id = Convert.ToInt32(r.GetValue(0)),
                                CadCategoryId = r.IsDBNull(1) ? null : (int?)Convert.ToInt32(r.GetValue(1)),
                                Name = r.IsDBNull(2) ? string.Empty : Convert.ToString(r.GetValue(2)),
                                DisplayName = r.IsDBNull(3) ? (r.IsDBNull(2) ? string.Empty : Convert.ToString(r.GetValue(2))) : Convert.ToString(r.GetValue(3)),
                                Description = r.IsDBNull(4) ? string.Empty : Convert.ToString(r.GetValue(4)),
                                SortOrder = r.IsDBNull(5) ? 0 : Convert.ToInt32(r.GetValue(5)),
                                IsActive = !r.IsDBNull(6) && Convert.ToInt32(r.GetValue(6)) == 1,
                                UserCount = r.IsDBNull(7) ? 0 : Convert.ToInt32(r.GetValue(7))
                            });
                        }
                    }
                }
            }

            return list;
        }

        /// <summary>
        /// 当 DEPARTMENTS 暂时没有数据时，直接用 CAD_CATEGORIES 作为部门下拉框的兜底来源。
        /// </summary>
        private List<DepartmentModel> LoadDepartmentsFromCategoriesFallback()
        {
            var list = new List<DepartmentModel>();
            using (var conn = CreateOpenConnection())
            {
                var categoriesTableName = ResolvePreferredTableName(conn, "CAD_CATEGORIES", _dbUser);
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = $@"SELECT ID, NAME, DISPLAY_NAME, SORT_ORDER FROM {categoriesTableName} ORDER BY SORT_ORDER";
                    LogManager.Instance.LogInfo($"LoadDepartmentsFromCategoriesFallback: 使用数据源 {categoriesTableName}");
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        var id = Convert.ToInt32(r.GetValue(0));
                        var name = r.IsDBNull(1) ? string.Empty : Convert.ToString(r.GetValue(1));
                        var displayName = r.IsDBNull(2) ? name : Convert.ToString(r.GetValue(2));
                        var sortOrder = r.IsDBNull(3) ? 0 : Convert.ToInt32(r.GetValue(3));

                        list.Add(new DepartmentModel
                        {
                            Id = id,
                            CadCategoryId = id,
                            Name = name,
                            DisplayName = displayName,
                            Description = string.Empty,
                            SortOrder = sortOrder,
                            IsActive = true,
                            UserCount = 0
                        });
                    }
                }
                }
            }

            return list;
        }

        /// <summary>
        /// 添加部门（兼容最终模式：无前缀，无双引号）
        /// </summary>
        /// <param name="name">部门名称</param>
        /// <param name="displayName">显示名称</param>
        /// <param name="description">描述</param>
        /// <param name="managerUserId">部门经理用户ID</param>
        /// <param name="sortOrder">排序顺序</param>
        /// <returns>新添加的部门ID</returns>
        public int AddDepartment(string name, string displayName = null, string description = null, int? managerUserId = null, int sortOrder = 0)
        {
            EnsureAllTablesExist();
            if (string.IsNullOrWhiteSpace(name)) return 0;
            using (var conn = CreateOpenConnection())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "INSERT INTO DEPARTMENTS (NAME, DISPLAY_NAME, DESCRIPTION, MANAGER_USER_ID, SORT_ORDER, IS_ACTIVE, CREATED_AT) VALUES (:n, :dn, :desc, :mgr, :so, 1, SYSDATE)";
                    AddParam(cmd, "n", name.Trim()); AddParam(cmd, "dn", displayName ?? name);
                    AddParam(cmd, "desc", description); AddParam(cmd, "mgr", managerUserId); AddParam(cmd, "so", sortOrder);
                    cmd.ExecuteNonQuery();
                }
                using (var tid = conn.CreateCommand())
                {
                    // 利用基础方式获取最新 ID
                    tid.CommandText = "SELECT ID FROM DEPARTMENTS WHERE NAME = :n ORDER BY ID DESC";
                    AddParam(tid, "n", name.Trim());
                    using (var r = tid.ExecuteReader())
                    {
                        if (r.Read()) return r.GetInt32(0);
                    }
                }
                return 0;
            }
        }

        /// <summary>
        /// 更新部门
        /// </summary>
        /// <param name="id">部门ID</param>
        /// <param name="name">部门名称</param>
        /// <param name="displayName">显示名称</param>
        /// <param name="description">描述</param>
        /// <param name="sortOrder">排序顺序</param>
        /// <param name="managerUserId">部门经理用户ID</param>
        /// <param name="isActive">是否激活</param>
        /// <returns>是否更新成功</returns>
        public bool UpdateDepartment(int id, string name, string displayName = null, string description = null, int sortOrder = 0, int? managerUserId = null, bool? isActive = null)
        {
            EnsureAllTablesExist();
            if (id <= 0) return false;
            using (var conn = CreateOpenConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "UPDATE DEPARTMENTS SET NAME=:n, DISPLAY_NAME=:dn, DESCRIPTION=:desc, MANAGER_USER_ID=:mgr, SORT_ORDER=:so, IS_ACTIVE=:ia, UPDATED_AT=SYSDATE WHERE ID=:id";
                AddParam(cmd, "n", name); AddParam(cmd, "dn", displayName); AddParam(cmd, "desc", description);
                AddParam(cmd, "mgr", managerUserId); AddParam(cmd, "so", sortOrder); AddParam(cmd, "ia", (isActive ?? true) ? 1 : 0); AddParam(cmd, "id", id);
                return cmd.ExecuteNonQuery() > 0;
            }
        }

        /// <summary>
        /// 删除部门（同时将该部门下的用户的部门ID重置为0，部门名称重置为空字符串，兼容最终模式：无前缀，无双引号）
        /// </summary>
        /// <param name="id">部门ID</param>
        /// <returns>是否删除成功</returns>
        public bool DeleteDepartment(int id)
        {
            EnsureAllTablesExist();
            if (id <= 0) return false;
            using (var conn = CreateOpenConnection())
            {
                using (var c1 = conn.CreateCommand())
                {
                    c1.CommandText = "UPDATE USERS SET DEPARTMENT_ID = 0, DEPARTMENT_NAME = '' WHERE DEPARTMENT_ID = :id";
                    AddParam(c1, "id", id); c1.ExecuteNonQuery();
                }
                using (var c2 = conn.CreateCommand())
                {
                    c2.CommandText = "DELETE FROM DEPARTMENTS WHERE ID = :id";
                    AddParam(c2, "id", id); return c2.ExecuteNonQuery() > 0;
                }
            }
        }

        /// <summary>
        /// 获取指定部门ID下的所有用户（兼容最终模式：无前缀，无双引号）
        /// </summary>
        /// <param name="departmentId">部门ID</param>
        /// <returns>用户列表</returns>
        public List<UserModel> GetUsersByDepartmentId(int departmentId)
        {
            EnsureAllTablesExist();
            var list = new List<UserModel>();
            using (var conn = CreateOpenConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT ID, USERNAME, ROLE, IS_ACTIVE FROM USERS WHERE DEPARTMENT_ID = :d ORDER BY ID";
                AddParam(cmd, "d", departmentId);
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        list.Add(new UserModel
                        {
                            Id = r.GetInt32(0),
                            Username = r.GetString(1),
                            Role = r.IsDBNull(2) ? "" : r.GetString(2),
                            IsActive = r.GetInt16(3) == 1
                        });
                    }
                }
            }
            return list;
        }

        /// <summary>
        /// 根据用户名将用户分配到部门
        /// </summary>
        /// <param name="username">用户名</param>
        /// <param name="departmentId">部门ID</param>
        /// <returns>是否分配成功</returns>
        public bool AssignUserToDepartmentByUsername(string username, int departmentId)
        {
            EnsureAllTablesExist();
            if (string.IsNullOrEmpty(username)) return false;
            using (var conn = CreateOpenConnection())
            {
                string dname = "";
                using (var gdn = conn.CreateCommand())
                {
                    gdn.CommandText = "SELECT NAME FROM DEPARTMENTS WHERE ID = :id";
                    AddParam(gdn, "id", departmentId);
                    dname = gdn.ExecuteScalar()?.ToString() ?? "";
                }
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "UPDATE USERS SET DEPARTMENT_ID = :d, DEPARTMENT_NAME = :dn WHERE USERNAME = :u";
                    AddParam(cmd, "d", departmentId); AddParam(cmd, "dn", dname); AddParam(cmd, "u", username);
                    return cmd.ExecuteNonQuery() > 0;
                }
            }
        }

        /// <summary>
        /// 获取当前数据库中所有表的名称和注释（兼容最终模式：无前缀，无双引号）
        /// </summary>
        /// <returns></returns>
        public List<TableInfo> GetTables()
        {
            EnsureAllTablesExist();
            var list = new List<TableInfo>();
            using (var conn = CreateOpenConnection())
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT T.TABLE_NAME, C.COMMENTS FROM USER_TABLES T LEFT JOIN USER_TAB_COMMENTS C ON T.TABLE_NAME = C.TABLE_NAME ORDER BY T.TABLE_NAME";
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        list.Add(new TableInfo { TableName = r.GetString(0), TableComment = r.IsDBNull(1) ? "" : r.GetString(1) });
                    }
                }
            }
            return list;
        }

        /// <summary>
        /// 从CAD类别同步部门信息：将 CAD_CATEGORIES 表中的分类信息同步到 DEPARTMENTS 表中，已存在的记录则更新，不存在的记录则插入（兼容最终模式：无前缀，无双引号）  
        /// </summary>
        //public void SyncDepartmentsFromCadCategories()
        //{
        //    EnsureAllTablesExist();// 确保表存在，避免因表不存在导致的同步失败
        //    using (var conn = CreateOpenConnection())// 使用单一连接执行整个同步过程，确保事务一致性和性能优化
        //    {
        //        var categoriesTableName = ResolvePreferredTableName(conn, "CAD_CATEGORIES", _dbUser);
        //        var categoriesOwner = GetOwnerFromQualifiedTableName(categoriesTableName);
        //        var departmentsTableName = string.IsNullOrWhiteSpace(categoriesOwner) ? "DEPARTMENTS" : categoriesOwner + ".DEPARTMENTS";

        //        LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories: categoriesTable={categoriesTableName}, departmentsTable={departmentsTableName}");

        //        // 1. 读取基础表数据
        //        var list = new List<(int Id, string Name, string Display, int So)>();
        //        using (var sel = conn.CreateCommand())
        //        {
        //            // 明确指定列名，避免因表中存在额外字段（如您提到的逗号字符串或日期列）导致的索引混乱
        //            sel.CommandText = $"SELECT ID, NAME, DISPLAY_NAME, SORT_ORDER FROM {categoriesTableName}";
        //            using (var r = sel.ExecuteReader())
        //            {
        //                while (r.Read())
        //                {
        //                    // 增加严谨的 NULL 判断和显式转换，确保能读到您表中的数据
        //                    int id = r.GetInt32(0);
        //                    string name = r.IsDBNull(1) ? "未命名分类" : r.GetString(1);
        //                    string disp = r.IsDBNull(2) ? name : r.GetString(2);
        //                    int sort = r.IsDBNull(3) ? 0 : Convert.ToInt32(r.GetValue(3));

        //                    list.Add((id, name, disp, sort));
        //                }
        //            }
        //        }

        //        // 如果源表没数据，记录日志以便排查
        //        if (list.Count == 0)
        //        {
        //            LogManager.Instance.LogInfo("SyncDepartmentsFromCadCategories: CAD_CATEGORIES 表为空，无数据可同步。");
        //            return;
        //        }

        //        LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories: 已从 CAD_CATEGORIES 读取到 {list.Count} 条分类数据。");

        //        // 2. 执行同步逻辑
        //        // 先按 CAD_CATEGORY_ID 匹配；若未命中，再按 NAME 匹配，避免已有同名部门时重复插入触发唯一约束。
        //        var affectedRows = 0;
        //        foreach (var c in list)
        //        {
        //            int? matchedDepartmentId = null;
        //            string matchMode = string.Empty;

        //            using (var chkByCategory = conn.CreateCommand())
        //            {
        //                chkByCategory.CommandText = $"SELECT ID FROM {departmentsTableName} WHERE CAD_CATEGORY_ID = :id";
        //                AddParam(chkByCategory, "id", c.Id);
        //                var existingId = chkByCategory.ExecuteScalar();
        //                if (existingId != null && existingId != DBNull.Value)
        //                {
        //                    matchedDepartmentId = Convert.ToInt32(existingId);
        //                    matchMode = "CAD_CATEGORY_ID";
        //                }
        //            }

        //            if (!matchedDepartmentId.HasValue)
        //            {
        //                using (var chkByName = conn.CreateCommand())
        //                {
        //                    chkByName.CommandText = $"SELECT ID FROM {departmentsTableName} WHERE NAME = :name";
        //                    AddParam(chkByName, "name", c.Name);
        //                    var existingId = chkByName.ExecuteScalar();
        //                    if (existingId != null && existingId != DBNull.Value)
        //                    {
        //                        matchedDepartmentId = Convert.ToInt32(existingId);
        //                        matchMode = "NAME";
        //                    }
        //                }
        //            }

        //            using (var cmd = conn.CreateCommand())
        //            {
        //                if (matchedDepartmentId.HasValue)
        //                {
        //                    // 命中现有部门后统一按主键更新，同时回填 CAD_CATEGORY_ID，避免后续再次误判为不存在。
        //                    cmd.CommandText = $"UPDATE {departmentsTableName} SET CAD_CATEGORY_ID = :cid, NAME = :n, DISPLAY_NAME = :d, SORT_ORDER = :s, IS_ACTIVE = 1, UPDATED_AT = SYSDATE WHERE ID = :deptId";
        //                    AddParam(cmd, "deptId", matchedDepartmentId.Value);
        //                }
        //                else
        //                {
        //                    // 只有按分类ID和名称都未命中时，才执行插入。
        //                    cmd.CommandText = $"INSERT INTO {departmentsTableName} (CAD_CATEGORY_ID, NAME, DISPLAY_NAME, SORT_ORDER, IS_ACTIVE, CREATED_AT) VALUES (:cid, :n, :d, :s, 1, SYSDATE)";
        //                }

        //                AddParam(cmd, "cid", c.Id);
        //                AddParam(cmd, "n", c.Name);
        //                AddParam(cmd, "d", string.IsNullOrEmpty(c.Display) ? c.Name : c.Display);
        //                AddParam(cmd, "s", c.So);
        //                affectedRows += cmd.ExecuteNonQuery();
        //            }

        //            if (matchedDepartmentId.HasValue)
        //            {
        //                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories: 分类 {c.Id}-{c.Name} 通过 {matchMode} 命中部门 ID={matchedDepartmentId.Value}，已执行更新。");
        //            }
        //        }

        //        // 达梦当前连接下 DML 结果需要显式提交，否则后续新连接读取不到刚同步的数据。
        //        using (var commitCmd = conn.CreateCommand())
        //        {
        //            commitCmd.CommandText = "COMMIT";
        //            commitCmd.ExecuteNonQuery();
        //        }

        //        LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories: 同步完成，受影响行数={affectedRows}。");
        //    }
        //}


        // 同步 CAD_CATEGORIES 到 DEPARTMENTS，自动区分数据库类型

        
        /// <summary>
        /// 同步 CAD_CATEGORIES 到 DEPARTMENTS，自动区分数据库类型
        /// </summary>
        public void SyncDepartmentsFromCadCategories()
        {
            // 判断当前数据库类型，分别调用不同实现
            var dbType = VariableDictionary._databaseType?.ToUpper() ?? "DM";
            if (dbType == "MYSQL")
            {
                SyncDepartmentsFromCadCategories_MySQL();// MySQL 版本的同步方法，兼容 MySQL 的 SQL 语法和特性
            }
            else
            {
                SyncDepartmentsFromCadCategories_DM();// 达梦数据库版本的同步方法，兼容达梦数据库的 SQL 语法和特性
            }
        }

        /// <summary>
        /// 达梦数据库专用同步方法
        /// </summary>
        private void SyncDepartmentsFromCadCategories_DM()
        {
            EnsureAllTablesExist();// 确保表存在，避免因表不存在导致的同步失败
            using (var conn = CreateOpenConnection())// 使用单一连接执行整个同步过程，确保事务一致性和性能优化
            {
                // 解析表名和所属用户，构建正确的表访问路径，兼容不同环境下可能存在的前缀或用户差异
                var categoriesTableName = ResolvePreferredTableName(conn, "CAD_CATEGORIES", _dbUser);
                // 从解析到的表名中提取所属用户（如果有），以便构建部门表的访问路径
                var categoriesOwner = GetOwnerFromQualifiedTableName(categoriesTableName);
                // 根据分类表的所属用户动态构建部门表的访问路径，确保在同一用户下访问部门表，避免跨用户访问权限问题
                var departmentsTableName = string.IsNullOrWhiteSpace(categoriesOwner) ? "DEPARTMENTS" : categoriesOwner + ".DEPARTMENTS";
                // 记录解析结果以便排查日志，确保能正确识别表名和访问路径
                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories_DM: categoriesTable={categoriesTableName}, departmentsTable={departmentsTableName}");

                // 1. 读取基础表数据
                var list = new List<(int Id, string Name, string Display, int So)>();
                using (var sel = conn.CreateCommand())// 明确指定列名，避免因表中存在额外字段（如您提到的逗号字符串或日期列）导致的索引混乱
                {
                    // 兼容不同环境下可能存在的字段差异，增加严谨的 NULL 判断和显式转换，确保能读到您表中的数据
                    sel.CommandText = $"SELECT ID, NAME, DISPLAY_NAME, SORT_ORDER FROM {categoriesTableName}";
                    using (var r = sel.ExecuteReader())// 使用 ExecuteReader 逐行读取，避免一次性加载大量数据导致的内存问题，同时能更灵活地处理数据转换和异常
                    {
                        while (r.Read())// 逐行读取数据，增加异常处理和数据验证，确保能正确处理各种数据情况
                        {
                            int id = r.GetInt32(0);// ID 列必须存在且不能为空，否则无法同步
                            string name = r.IsDBNull(1) ? "未命名分类" : r.GetString(1);// NAME 列如果不存在或为空，使用默认值避免同步失败
                            string disp = r.IsDBNull(2) ? name : r.GetString(2);// DISPLAY_NAME 列如果不存在或为空，使用 NAME 作为显示名称，确保界面有值可显示
                            int sort = r.IsDBNull(3) ? 0 : Convert.ToInt32(r.GetValue(3));// SORT_ORDER 列如果不存在或为空，使用默认值 0，确保同步逻辑有排序依据
                            list.Add((id, name, disp, sort));// 将读取到的数据添加到列表中，后续进行同步处理
                        }
                    }
                }

                if (list.Count == 0)// 如果源表没数据，记录日志以便排查
                {
                    LogManager.Instance.LogInfo("SyncDepartmentsFromCadCategories_DM: CAD_CATEGORIES 表为空，无数据可同步。");
                    return;
                }
                // 记录成功读取到的数据量，确保能确认同步的基础数据情况
                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories_DM: 已从 CAD_CATEGORIES 读取到 {list.Count} 条分类数据。");

                // 2. 执行同步逻辑
                var affectedRows = 0;
                foreach (var c in list)// 先按 CAD_CATEGORY_ID 匹配；若未命中，再按 NAME 匹配，避免已有同名部门时重复插入触发唯一约束。
                {
                    int? matchedDepartmentId = null;// 先尝试通过 CAD_CATEGORY_ID 匹配，确保同一分类ID对应同一部门，避免重复插入
                    string matchMode = string.Empty;// 记录匹配方式，便于后续日志分析和问题排查
                    // 通过 CAD_CATEGORY_ID 匹配部门，确保分类与部门的一一对应关系，避免同一分类被多个部门重复引用
                    using (var chkByCategory = conn.CreateCommand())// 使用单一连接执行匹配查询，确保事务一致性和性能优化
                    {
                        chkByCategory.CommandText = $"SELECT ID FROM {departmentsTableName} WHERE CAD_CATEGORY_ID = :id";// 兼容不同环境下可能存在的字段差异，增加严谨的 NULL 判断和显式转换，确保能正确处理数据
                        AddParam(chkByCategory, "id", c.Id);// 使用参数化查询避免 SQL 注入风险，同时确保数据类型正确传递
                        var existingId = chkByCategory.ExecuteScalar();// 直接获取匹配的部门ID，如果存在且不为 NULL，则说明找到了对应的部门
                        if (existingId != null && existingId != DBNull.Value)// 如果查询结果不为 NULL，说明找到了匹配的部门，记录匹配的部门ID和匹配方式
                        {
                            // 将匹配到的部门ID转换为整数类型，并记录匹配方式为 CAD_CATEGORY_ID，便于后续日志分析和问题排查
                            matchedDepartmentId = Convert.ToInt32(existingId);
                            // 记录匹配方式为 CAD_CATEGORY_ID，便于后续日志分析和问题排查
                            matchMode = "CAD_CATEGORY_ID";
                        }
                    }

                    if (!matchedDepartmentId.HasValue)// 如果未通过 CAD_CATEGORY_ID 匹配到部门，再尝试通过 NAME 匹配，确保同一名称的分类能对应到同一部门，避免重复插入
                    {
                        // 通过 NAME 匹配部门，确保同一名称的分类能对应到同一部门，避免重复插入触发唯一约束，同时兼容已有数据中可能存在的 CAD_CATEGORY_ID 为空但名称匹配的情况
                        using (var chkByName = conn.CreateCommand())
                        {
                            // 使用单一连接执行匹配查询，确保事务一致性和性能优化，同时增加日志记录以便排查匹配过程中的问题
                            chkByName.CommandText = $"SELECT ID FROM {departmentsTableName} WHERE NAME = :name";
                            // 兼容不同环境下可能存在的字段差异，增加严谨的 NULL 判断和显式转换，确保能正确处理数据，同时使用参数化查询避免 SQL 注入风险
                            AddParam(chkByName, "name", c.Name);
                            // 直接获取匹配的部门ID，如果存在且不为 NULL，则说明找到了对应的部门，记录匹配的部门ID和匹配方式
                            var existingId = chkByName.ExecuteScalar();
                            // 如果查询结果不为 NULL，说明找到了匹配的部门，记录匹配的部门ID和匹配方式为 NAME，便于后续日志分析和问题排查
                            if (existingId != null && existingId != DBNull.Value)
                            {
                                // 将匹配到的部门ID转换为整数类型，并记录匹配方式为 NAME，便于后续日志分析和问题排查
                                matchedDepartmentId = Convert.ToInt32(existingId);
                                // 记录匹配方式为 NAME，便于后续日志分析和问题排查
                                matchMode = "NAME";
                            }
                        }
                    }
                    // 根据匹配结果执行更新或插入操作，如果匹配到现有部门则执行更新，回填 CAD_CATEGORY_ID 和其他字段；如果未匹配到则执行插入，创建新部门记录，确保同步后的部门表能完整反映分类表的数据，同时避免重复插入触发唯一约束
                    using (var cmd = conn.CreateCommand())
                    {
                        // 如果匹配到现有部门，则执行更新操作，回填 CAD_CATEGORY_ID 和其他字段，同时记录日志说明是通过哪种方式匹配到的部门，以便后续分析和排查
                        if (matchedDepartmentId.HasValue)
                        {
                            // 命中现有部门后统一按主键更新，同时回填 CAD_CATEGORY_ID，避免后续再次误判为不存在，同时记录日志说明是通过哪种方式匹配到的部门，以便后续分析和排查
                            cmd.CommandText = $"UPDATE {departmentsTableName} SET CAD_CATEGORY_ID = :cid, NAME = :n, DISPLAY_NAME = :d, SORT_ORDER = :s, IS_ACTIVE = 1, UPDATED_AT = SYSDATE WHERE ID = :deptId";
                            // 使用参数化查询避免 SQL 注入风险，同时确保数据类型正确传递，同时记录日志说明是通过哪种方式匹配到的部门，以便后续分析和排查
                            AddParam(cmd, "deptId", matchedDepartmentId.Value);
                        }
                        else
                        {
                            // 只有按分类ID和名称都未命中时，才执行插入，创建新部门记录，确保同步后的部门表能完整反映分类表的数据，同时避免重复插入触发唯一约束
                            cmd.CommandText = $"INSERT INTO {departmentsTableName} (CAD_CATEGORY_ID, NAME, DISPLAY_NAME, SORT_ORDER, IS_ACTIVE, CREATED_AT) VALUES (:cid, :n, :d, :s, 1, SYSDATE)";
                        }
                        AddParam(cmd, "cid", c.Id);// 使用参数化查询避免 SQL 注入风险，同时确保数据类型正确传递
                        AddParam(cmd, "n", c.Name);// 使用参数化查询避免 SQL 注入风险，同时确保数据类型正确传递
                        AddParam(cmd, "d", string.IsNullOrEmpty(c.Display) ? c.Name : c.Display);// 使用参数化查询避免 SQL 注入风险，同时确保数据类型正确传递
                        AddParam(cmd, "s", c.So);// 使用参数化查询避免 SQL 注入风险，同时确保数据类型正确传递
                        affectedRows += cmd.ExecuteNonQuery();// 执行更新或插入操作，并累加受影响的行数，便于后续日志记录和分析
                    }

                    if (matchedDepartmentId.HasValue)// 如果匹配到现有部门，记录日志说明是通过哪种方式匹配到的部门，以便后续分析和排查
                    {
                        LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories_DM: 分类 {c.Id}-{c.Name} 通过 {matchMode} 命中部门 ID={matchedDepartmentId.Value}，已执行更新。");
                    }
                }

                // 达梦当前连接下 DML 结果需要显式提交，否则后续新连接读取不到刚同步的数据。
                using (var commitCmd = conn.CreateCommand())
                {
                    commitCmd.CommandText = "COMMIT";// 执行 COMMIT 命令提交事务，确保同步后的数据能被其他连接读取到，同时记录日志说明已执行提交操作，以便后续分析和排查
                    commitCmd.ExecuteNonQuery();// 执行 COMMIT 命令提交事务，确保同步后的数据能被其他连接读取到，同时记录日志说明已执行提交操作，以便后续分析和排查
                }

                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories_DM: 同步完成，受影响行数={affectedRows}。");
            }
        }

        /// <summary>
        /// MySQL数据库专用同步方法
        /// </summary>
        private void SyncDepartmentsFromCadCategories_MySQL()
        {
            // 确保表存在
            var mySvc = new MySqlAuthService(
                VariableDictionary._serverIP,
                VariableDictionary._serverPort.ToString(),
                VariableDictionary._userName,
                VariableDictionary._passWord
            );
            mySvc.EnsureAllTablesExist();

            using (var conn = new MySql.Data.MySqlClient.MySqlConnection(
                $"Server={VariableDictionary._serverIP};Port={VariableDictionary._serverPort};Database=cad_sw_library;Uid={VariableDictionary._userName};Pwd={VariableDictionary._passWord};Allow User Variables=True;"))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                {
                    // 1. 读取分类表
                    var list = new List<(int Id, string Name, string Display, int So)>();
                    using (var sel = conn.CreateCommand())
                    {
                        sel.Transaction = tx;
                        sel.CommandText = "SELECT id, name, display_name, sort_order FROM cad_categories";
                        using (var r = sel.ExecuteReader())
                        {
                            while (r.Read())
                            {
                                int id = r.GetInt32("id");
                                string name = r.IsDBNull(r.GetOrdinal("name")) ? "未命名分类" : r.GetString("name");
                                string disp = r.IsDBNull(r.GetOrdinal("display_name")) ? name : r.GetString("display_name");
                                int sort = r.IsDBNull(r.GetOrdinal("sort_order")) ? 0 : r.GetInt32("sort_order");
                                list.Add((id, name, disp, sort));
                            }
                        }
                    }

                    if (list.Count == 0)
                    {
                        LogManager.Instance.LogInfo("SyncDepartmentsFromCadCategories_MySQL: cad_categories 表为空，无数据可同步。");
                        tx.Commit();
                        return;
                    }

                    LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories_MySQL: 已从 cad_categories 读取到 {list.Count} 条分类数据。");

                    // 2. 执行同步逻辑
                    var affectedRows = 0;
                    foreach (var c in list)
                    {
                        int? matchedDepartmentId = null;
                        string matchMode = string.Empty;

                        using (var chkByCategory = conn.CreateCommand())
                        {
                            chkByCategory.Transaction = tx;
                            chkByCategory.CommandText = "SELECT id FROM departments WHERE cad_category_id = @id";
                            chkByCategory.Parameters.AddWithValue("@id", c.Id);
                            var existingId = chkByCategory.ExecuteScalar();
                            if (existingId != null && existingId != DBNull.Value)
                            {
                                matchedDepartmentId = Convert.ToInt32(existingId);
                                matchMode = "CAD_CATEGORY_ID";
                            }
                        }

                        if (!matchedDepartmentId.HasValue)
                        {
                            using (var chkByName = conn.CreateCommand())
                            {
                                chkByName.Transaction = tx;
                                chkByName.CommandText = "SELECT id FROM departments WHERE name = @name";
                                chkByName.Parameters.AddWithValue("@name", c.Name);
                                var existingId = chkByName.ExecuteScalar();
                                if (existingId != null && existingId != DBNull.Value)
                                {
                                    matchedDepartmentId = Convert.ToInt32(existingId);
                                    matchMode = "NAME";
                                }
                            }
                        }

                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = tx;
                            if (matchedDepartmentId.HasValue)
                            {
                                cmd.CommandText = "UPDATE departments SET cad_category_id = @cid, name = @n, display_name = @d, sort_order = @s, is_active = 1, updated_at = CURRENT_TIMESTAMP WHERE id = @deptId";
                                cmd.Parameters.AddWithValue("@deptId", matchedDepartmentId.Value);
                            }
                            else
                            {
                                cmd.CommandText = "INSERT INTO departments (cad_category_id, name, display_name, sort_order, is_active, created_at) VALUES (@cid, @n, @d, @s, 1, CURRENT_TIMESTAMP)";
                            }
                            cmd.Parameters.AddWithValue("@cid", c.Id);
                            cmd.Parameters.AddWithValue("@n", c.Name);
                            cmd.Parameters.AddWithValue("@d", string.IsNullOrEmpty(c.Display) ? c.Name : c.Display);
                            cmd.Parameters.AddWithValue("@s", c.So);
                            affectedRows += cmd.ExecuteNonQuery();
                        }

                        if (matchedDepartmentId.HasValue)
                        {
                            LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories_MySQL: 分类 {c.Id}-{c.Name} 通过 {matchMode} 命中部门 ID={matchedDepartmentId.Value}，已执行更新。");
                        }
                    }

                    tx.Commit();
                    LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories_MySQL: 同步完成，受影响行数={affectedRows}。");
                }
            }
        }
        #endregion

        #region 加密套件 (SHA512)
        /// <summary>
        /// 生成一个随机盐值（32字节，Base64编码），用于密码哈希
        /// </summary>
        /// <returns></returns>

        private string GenerateSalt()
        {
            byte[] bytes = new byte[32];
            using (var rng = new RNGCryptoServiceProvider()) rng.GetBytes(bytes);
            return Convert.ToBase64String(bytes);
        }
        /// <summary>
        /// 计算哈希
        /// </summary>
        /// <param name="pwd">密码</param>
        /// <param name="salt">盐值</param>
        /// <returns>哈希值</returns>
        private string ComputeHash(string pwd, string salt)
        {
            using (var sha = SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(pwd + salt);
                return Convert.ToBase64String(sha.ComputeHash(bytes));
            }
        }

        #endregion
    }
}
