using Dm; // 达梦驱动命名空间
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using GB_NewCadPlus_IV.DMDatabaseReader;

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
            var normalizedTableName = (tableName ?? string.Empty).Trim().ToUpperInvariant();
            var candidates = new List<Tuple<string, string, int>>();

            // 先把当前 Schema 下的同名表加入候选。
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT COUNT(1) FROM USER_TABLES WHERE TABLE_NAME = :t";
                AddParam(cmd, "t", normalizedTableName);
                var existsInCurrentSchema = Convert.ToInt32(cmd.ExecuteScalar() ?? 0) > 0;
                if (existsInCurrentSchema)
                {
                    candidates.Add(Tuple.Create(string.Empty, normalizedTableName, TryGetTableRowCount(conn, normalizedTableName)));
                }
            }

            // 再从 ALL_TABLES 中查找其它 Schema 的同名表。
            try
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT OWNER FROM ALL_TABLES WHERE TABLE_NAME = :t ORDER BY OWNER";
                    AddParam(cmd, "t", normalizedTableName);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var owner = reader.IsDBNull(0) ? string.Empty : Convert.ToString(reader.GetValue(0));
                            if (string.IsNullOrWhiteSpace(owner))
                            {
                                continue;
                            }

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
                LogManager.Instance.LogInfo($"ResolvePreferredTableName: 查询 ALL_TABLES 失败，table={normalizedTableName}，{ex.Message}");
            }

            if (candidates.Count == 0)
            {
                LogManager.Instance.LogInfo($"ResolvePreferredTableName: 未找到 {normalizedTableName}，将直接使用默认表名。\n");
                return normalizedTableName;
            }

            // 先优先使用指定 Owner 且有数据的表。
            if (!string.IsNullOrWhiteSpace(preferredOwner))
            {
                var preferred = candidates.FirstOrDefault(c =>
                    string.Equals(c.Item1, preferredOwner, StringComparison.OrdinalIgnoreCase) && c.Item3 > 0);
                if (preferred != null)
                {
                    LogManager.Instance.LogInfo($"ResolvePreferredTableName: {normalizedTableName} 命中首选 Schema={preferred.Item1}，Rows={preferred.Item3}");
                    return preferred.Item2;
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
            if (string.IsNullOrWhiteSpace(qualifiedTableName))
            {
                return string.Empty;
            }

            var index = qualifiedTableName.IndexOf('.');
            if (index <= 0)
            {
                return string.Empty;
            }

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
        public void SyncDepartmentsFromCadCategories()
        {
            EnsureAllTablesExist();
            using (var conn = CreateOpenConnection())
            {
                var categoriesTableName = ResolvePreferredTableName(conn, "CAD_CATEGORIES", _dbUser);
                var categoriesOwner = GetOwnerFromQualifiedTableName(categoriesTableName);
                var departmentsTableName = string.IsNullOrWhiteSpace(categoriesOwner) ? "DEPARTMENTS" : categoriesOwner + ".DEPARTMENTS";

                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories: categoriesTable={categoriesTableName}, departmentsTable={departmentsTableName}");

                // 1. 读取基础表数据
                var list = new List<(int Id, string Name, string Display, int So)>();
                using (var sel = conn.CreateCommand())
                {
                    // 明确指定列名，避免因表中存在额外字段（如您提到的逗号字符串或日期列）导致的索引混乱
                    sel.CommandText = $"SELECT ID, NAME, DISPLAY_NAME, SORT_ORDER FROM {categoriesTableName}";
                    using (var r = sel.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            // 增加严谨的 NULL 判断和显式转换，确保能读到您表中的数据
                            int id = r.GetInt32(0);
                            string name = r.IsDBNull(1) ? "未命名分类" : r.GetString(1);
                            string disp = r.IsDBNull(2) ? name : r.GetString(2);
                            int sort = r.IsDBNull(3) ? 0 : Convert.ToInt32(r.GetValue(3));

                            list.Add((id, name, disp, sort));
                        }
                    }
                }

                // 如果源表没数据，记录日志以便排查
                if (list.Count == 0)
                {
                    LogManager.Instance.LogInfo("SyncDepartmentsFromCadCategories: CAD_CATEGORIES 表为空，无数据可同步。");
                    return;
                }

                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories: 已从 CAD_CATEGORIES 读取到 {list.Count} 条分类数据。");

                // 2. 执行同步逻辑
                // 先按 CAD_CATEGORY_ID 匹配；若未命中，再按 NAME 匹配，避免已有同名部门时重复插入触发唯一约束。
                var affectedRows = 0;
                foreach (var c in list)
                {
                    int? matchedDepartmentId = null;
                    string matchMode = string.Empty;

                    using (var chkByCategory = conn.CreateCommand())
                    {
                        chkByCategory.CommandText = $"SELECT ID FROM {departmentsTableName} WHERE CAD_CATEGORY_ID = :id";
                        AddParam(chkByCategory, "id", c.Id);
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
                            chkByName.CommandText = $"SELECT ID FROM {departmentsTableName} WHERE NAME = :name";
                            AddParam(chkByName, "name", c.Name);
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
                        if (matchedDepartmentId.HasValue)
                        {
                            // 命中现有部门后统一按主键更新，同时回填 CAD_CATEGORY_ID，避免后续再次误判为不存在。
                            cmd.CommandText = $"UPDATE {departmentsTableName} SET CAD_CATEGORY_ID = :cid, NAME = :n, DISPLAY_NAME = :d, SORT_ORDER = :s, IS_ACTIVE = 1, UPDATED_AT = SYSDATE WHERE ID = :deptId";
                            AddParam(cmd, "deptId", matchedDepartmentId.Value);
                        }
                        else
                        {
                            // 只有按分类ID和名称都未命中时，才执行插入。
                            cmd.CommandText = $"INSERT INTO {departmentsTableName} (CAD_CATEGORY_ID, NAME, DISPLAY_NAME, SORT_ORDER, IS_ACTIVE, CREATED_AT) VALUES (:cid, :n, :d, :s, 1, SYSDATE)";
                        }

                        AddParam(cmd, "cid", c.Id);
                        AddParam(cmd, "n", c.Name);
                        AddParam(cmd, "d", string.IsNullOrEmpty(c.Display) ? c.Name : c.Display);
                        AddParam(cmd, "s", c.So);
                        affectedRows += cmd.ExecuteNonQuery();
                    }

                    if (matchedDepartmentId.HasValue)
                    {
                        LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories: 分类 {c.Id}-{c.Name} 通过 {matchMode} 命中部门 ID={matchedDepartmentId.Value}，已执行更新。");
                    }
                }

                // 达梦当前连接下 DML 结果需要显式提交，否则后续新连接读取不到刚同步的数据。
                using (var commitCmd = conn.CreateCommand())
                {
                    commitCmd.CommandText = "COMMIT";
                    commitCmd.ExecuteNonQuery();
                }

                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories: 同步完成，受影响行数={affectedRows}。");
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
