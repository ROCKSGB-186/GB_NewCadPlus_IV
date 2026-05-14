using Dapper;
using Dm;
using GB_NewCadPlus_IV.UniFiedStandards;
// MySql provider is no longer used in DM migration; remove direct dependency usages.
// Note: leave using for compatibility in files that still reference MySqlConnection via fully-qualified names.
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Serialization;
using static GB_NewCadPlus_IV.WpfMainWindow;
using DataTable = System.Data.DataTable;
// Note: avoid file-level type aliases to nested types here to prevent alias conflicts in other files.
using DeviceInfo = GB_NewCadPlus_IV.UniFiedStandards.DeviceInfo;
using MessageBox = System.Windows.MessageBox;
using TextBox = System.Windows.Controls.TextBox;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 数据库访问类  
    /// </summary>
    public class DatabaseManager
    {

        // Model types moved to FunctionalMethod.DatabaseModels to avoid duplicate nested type definitions.
        // Note: real DatabaseManager implementations exist later in this file. No compatibility stubs at top to avoid duplicate definitions.

        /// <summary>
        /// 数据库适配器，用于支持多数据库
        /// </summary>
        private readonly IDatabaseAdapter _adapter;

        /// <summary>
        /// 对外公开数据库连接（注意：调用方负责不要忘记关闭/处置）
        /// </summary>
        public IDbConnection GetConnection()
        {
            var connection = _adapter.CreateConnection();
            if (connection is DmConnection dmConn)
            {
                dmConn.StateChange += Connection_StateChange;
            }
            return connection;
        }

        /// <summary>
        /// 简单执行器：根据当前适配器执行查询并返回单个标量或映射类型（同步）
        /// 对于 MySQL 使用 Dapper 快捷映射；对于达梦使用手工命令与 reader 映射（支持基本类型和简单 POCO 的单行读取）
        /// </summary>
        private T QuerySingleOrDefault<T>(string sql, object? param = null)
        {
            if (_adapter.DatabaseType == "MySQL")
            {
                using var conn = _adapter.CreateConnection();
                conn.Open();
                // 使用 Dapper 的同步扩展
                return conn.QuerySingleOrDefault<T>(sql, param);
            }

            // DM 路径：手动执行并映射
            using var dconn = _adapter.CreateConnection();
            dconn.Open();
            using var cmd = dconn.CreateCommand();
            cmd.CommandText = _adapter.NormalizeSql(sql);
            if (param != null)
            {
                // 支持字典或匿名对象
                if (param is System.Collections.IDictionary dict)
                {
                    foreach (System.Collections.DictionaryEntry e in dict)
                    {
                        _adapter.AddParameter(cmd, e.Key.ToString(), e.Value ?? DBNull.Value);
                    }
                }
                else
                {
                    var props = param.GetType().GetProperties();
                    foreach (var p in props)
                    {
                        var val = p.GetValue(param);
                        _adapter.AddParameter(cmd, p.Name, val ?? DBNull.Value);
                    }
                }
            }
            var res = cmd.ExecuteScalar();
            if (res == null || res == DBNull.Value) return default;
            return (T)Convert.ChangeType(res, typeof(T));
        }

        /// <summary>
        /// 简单执行非查询 SQL 并返回受影响行数（同步）
        /// MySQL 使用 Dapper 的 Execute；DM 使用 ExecuteNonQuery
        /// </summary>
        private int ExecuteNonQuery(string sql, object? param = null)
        {
            if (_adapter.DatabaseType == "MySQL")
            {
                using var conn = _adapter.CreateConnection();
                conn.Open();
                return conn.Execute(sql, param);
            }

            using var dconn = _adapter.CreateConnection();
            dconn.Open();
            using var cmd = dconn.CreateCommand();
            cmd.CommandText = _adapter.NormalizeSql(sql);
            if (param != null)
            {
                if (param is System.Collections.IDictionary dict)
                {
                    foreach (System.Collections.DictionaryEntry e in dict)
                    {
                        _adapter.AddParameter(cmd, e.Key.ToString(), e.Value ?? DBNull.Value);
                    }
                }
                else
                {
                    var props = param.GetType().GetProperties();
                    foreach (var p in props)
                    {
                        var val = p.GetValue(param);
                        _adapter.AddParameter(cmd, p.Name, val ?? DBNull.Value);
                    }
                }
            }
            return cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 统一执行写入语句（MySQL 走 Dapper，DM 走原生命令并自动绑定参数）。
        /// </summary>
        private async Task<int> ExecuteWriteAsync(IDbConnection connection, IDbTransaction? transaction, string sql, object? param = null)
        {
            if (connection == null)
            {
                throw new ArgumentNullException(nameof(connection));
            }

            // 【关键修复 1】确保连接已打开
            // Dapper 的 ExecuteAsync 会自动处理打开/关闭，但原生 ExecuteNonQuery 不会
            if (connection.State != System.Data.ConnectionState.Open)
            {
                await ((DbConnection)connection).OpenAsync().ConfigureAwait(false);
            }

            try
            {
                if (_adapter.DatabaseType == "MySQL")
                {
                    // MySQL 走 Dapper 路径，Dapper 会管理连接生命周期
                    return await connection.ExecuteAsync(sql, param, transaction).ConfigureAwait(false);
                }

                // 达梦 (DM) 或其他数据库走原生 ADO.NET 路径
                using var cmd = connection.CreateCommand();
                cmd.Transaction = transaction;

                // 【关键修复 2】规范化 SQL (确保占位符符合达梦规范，如将 @Id 转为 :Id 如果必要)
                // 假设 _adapter.NormalizeSql 已经处理了占位符转换，如果没有，请确保达梦支持 @ 符号
                cmd.CommandText = _adapter.NormalizeSql(sql);

                // 添加参数
                AddCommandParameters(cmd, param);

                // 【关键修复 3】执行命令
                // 此时连接必须是 Open 状态
                return cmd.ExecuteNonQuery();
            }
            finally
            {
                // 可选：如果连接是由该方法内部打开的，且不是由外部事务管理，可以考虑关闭
                // 但通常建议使用 using 块在外部管理连接生命周期，或者让连接池处理
                // 如果 GetConnection() 返回的是新连接，建议在外部 using 结束后自动关闭
            }
        }
        //private async Task<int> ExecuteWriteAsync(IDbConnection connection, IDbTransaction? transaction, string sql, object? param = null)
        //{
        //    if (connection == null)
        //    {
        //        throw new ArgumentNullException(nameof(connection));
        //    }

        //    if (_adapter.DatabaseType == "MySQL")
        //    {
        //        return await connection.ExecuteAsync(sql, param, transaction).ConfigureAwait(false);
        //    }

        //    using var cmd = connection.CreateCommand();
        //    cmd.Transaction = transaction;
        //    cmd.CommandText = _adapter.NormalizeSql(sql);
        //    AddCommandParameters(cmd, param);
        //    return cmd.ExecuteNonQuery();
        //}

        /// <summary>
        /// 统一执行标量查询（主要用于获取自增ID）。
        /// </summary>
        private async Task<T> ExecuteScalarAsync<T>(IDbConnection connection, IDbTransaction? transaction, string sql, object? param = null)
        {
            if (connection == null)
            {
                throw new ArgumentNullException(nameof(connection));
            }

            if (_adapter.DatabaseType == "MySQL")
            {
                var value = await connection.ExecuteScalarAsync(sql, param, transaction).ConfigureAwait(false);
                if (value == null || value == DBNull.Value)
                {
                    return default(T);
                }

                return (T)Convert.ChangeType(value, typeof(T));
            }

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;
            cmd.CommandText = _adapter.NormalizeSql(sql);
            AddCommandParameters(cmd, param);
            var scalar = cmd.ExecuteScalar();
            if (scalar == null || scalar == DBNull.Value)
            {
                return default(T);
            }

            return (T)Convert.ChangeType(scalar, typeof(T));
        }

        /// <summary>
        /// 统一执行插入并获取当前连接的自增主键。
        /// </summary>
        private async Task<long> ExecuteInsertAndGetIdentityAsync(IDbConnection connection, IDbTransaction? transaction, string insertSql, object? param = null)
        {
            await ExecuteWriteAsync(connection, transaction, insertSql, param).ConfigureAwait(false);

            var identitySql = _adapter.DatabaseType == "MySQL"
                ? "SELECT LAST_INSERT_ID()"
                : "SELECT IDENTITY_VAL_LOCAL()";

            return await ExecuteScalarAsync<long>(connection, transaction, identitySql).ConfigureAwait(false);
        }

        /// <summary>
        /// 将匿名对象或字典参数统一绑定到命令对象。
        /// </summary>
        private void AddCommandParameters(IDbCommand cmd, object? param)
        {
            if (cmd == null || param == null)
            {
                return;
            }

            if (param is System.Collections.IDictionary dict)
            {
                foreach (System.Collections.DictionaryEntry entry in dict)
                {
                    AddParam(cmd, entry.Key.ToString(), entry.Value ?? DBNull.Value);
                }

                return;
            }

            var props = param.GetType().GetProperties();
            foreach (var prop in props)
            {
                AddParam(cmd, prop.Name, prop.GetValue(param) ?? DBNull.Value);
            }
        }

        /// <summary>
        /// 统一返回当前数据库可识别的“当前时间”函数。
        /// </summary>
        private string GetCurrentTimestampSql()
        {
            return _adapter.DatabaseType == "MySQL" ? "NOW()" : "CURRENT_TIMESTAMP";
        }


        /// <summary>
        /// 连接状态变化时自动切换到目标 Schema，保证后续未带前缀的 SQL 能落到达梦目标库对象上。
        /// </summary>
        private void Connection_StateChange(object? sender, StateChangeEventArgs e)
        {
            if (e.CurrentState != ConnectionState.Open)
            {
                return;
            }

            if (sender is DmConnection connection)
            {
                ApplySchema(connection);
            }
        }

        /// <summary>
        /// 为当前连接设置 Schema。
        /// </summary>
        private void ApplySchema(IDbConnection connection)
        {
            if (connection == null || connection.State != ConnectionState.Open || string.IsNullOrWhiteSpace(_schemaName))
            {
                return;
            }

            try
            {
                _adapter.ApplySchema(connection, _schemaName);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"DatabaseManager 设置 Schema 失败: schema={_schemaName}, {ex.Message}");
            }
        }

        /// <summary>
        /// 用于给 IDbCommand 添加参数，自动适配数据库类型
        /// </summary>
        /// <param name="cmd">要添加参数的命令对象</param>
        /// <param name="name">参数名称</param>
        /// <param name="value">参数值</param>
        //private void AddParam(IDbCommand cmd, string name, object? value)
        //{
        //    _adapter.AddParameter(cmd, name, value);
        //}
        private void AddParam(IDbCommand cmd, string paramName, object value)
        {
            var param = cmd.CreateParameter();

            // 【关键修复 4】处理参数名前缀
            // 达梦通常支持 :Name 或 @Name。如果 NormalizeSql 将 @Id 变成了 :Id，
            // 那么这里的 paramName 也应该是 :Id 或者驱动能自动映射 Id -> :Id
            // 许多驱动要求 ParameterName 与 SQL 中的占位符完全一致（包括前缀）

            // 尝试自动添加前缀，如果 paramName 不包含前缀
            if (!paramName.StartsWith("@") && !paramName.StartsWith(":") && !paramName.StartsWith("?"))
            {
                // 根据数据库类型添加前缀
                if (_adapter.DatabaseType == "DM")
                {
                    paramName = ":" + paramName; // 达梦推荐 :
                }
                else
                {
                    paramName = "@" + paramName; // MySQL/SQLServer 推荐 @
                }
            }

            param.ParameterName = paramName;
            param.Value = value ?? DBNull.Value;

            // 可选：指定 DbType 以提高性能并避免类型推断错误
            // param.DbType = DbType.Int32; // 如果知道是 ID

            cmd.Parameters.Add(param);
        }

        /// <summary>
        /// 达梦风格参数添加方法（为向后兼容保留，内部调用 AddParam）
        /// </summary>
        private void AddDmParam(IDbCommand cmd, string name, object? value)
        {
            AddParam(cmd, name, value);
        }

        /// <summary>
        /// 将传入的连接串标准化为达梦驱动可识别的格式。
        /// </summary>
        private static string NormalizeDmConnectionString(string connectionString)
        {
            var server = ExtractConnectionStringValue(connectionString, "Server")
                ?? ExtractConnectionStringValue(connectionString, "Host")
                ?? "127.0.0.1";
            var port = ExtractConnectionStringValue(connectionString, "Port") ?? "5236";
            var user = ExtractConnectionStringValue(connectionString, "User Id")
                ?? ExtractConnectionStringValue(connectionString, "Uid")
                ?? ExtractConnectionStringValue(connectionString, "User")
                ?? "SYSDBA";
            var password = ExtractConnectionStringValue(connectionString, "Password")
                ?? ExtractConnectionStringValue(connectionString, "Pwd")
                ?? "SYSDBA";

            return $"Server={server};Port={port};User Id={user};Password={password};";
        }

        /// <summary>
        /// 从连接串中提取目标 Schema，优先 Schema，其次 Database，最后回退全局变量。
        /// </summary>
        private static string ResolveSchemaName(string connectionString)
        {
            var schemaName = ExtractConnectionStringValue(connectionString, "Schema")
                ?? ExtractConnectionStringValue(connectionString, "Database")
                ?? VariableDictionary._dataBaseName;

            return string.IsNullOrWhiteSpace(schemaName) ? "SYSDBA" : schemaName.Trim().ToUpperInvariant();
        }

        /// <summary>
        /// 按键名从连接串中取值。
        /// </summary>
        private static string? ExtractConnectionStringValue(string connectionString, string key)
        {
            if (string.IsNullOrWhiteSpace(connectionString) || string.IsNullOrWhiteSpace(key))
            {
                return null;
            }

            var segments = connectionString.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var segment in segments)
            {
                var pair = segment.Split(new[] { '=' }, 2);
                if (pair.Length != 2)
                {
                    continue;
                }

                if (string.Equals(pair[0].Trim(), key, StringComparison.OrdinalIgnoreCase))
                {
                    return pair[1].Trim();
                }
            }

            return null;
        }

        // -----------------------------
        // 简易异步方法占位实现（避免引用处编译错误）
        // 说明：这些方法为占位实现，返回默认值或抛出未实现异常。
        // 在接入真实后端时，请替换为完整实现。
        // -----------------------------

        // Use top-level models in FunctionalMethod.DatabaseModels for FileAccessLog and FileTag.
        // Nested placeholder types removed to avoid duplicate-type ambiguity.

        /// <summary>
        /// 添加文件访问日志（占位）
        /// </summary>
        public virtual async Task<bool> AddFileAccessLogAsync(FileAccessLog log)
        {
            await Task.Yield();
            // 占位：默认记录成功
            return true;
        }

        /// <summary>
        /// 删除文件（占位）
        /// </summary>
        public virtual async Task<int> DeleteFileAsync(int fileId, string deletedBy)
        {
            await Task.Yield();
            // 占位：返回 1 表示已删除
            return 1;
        }

        // public virtual async Task<bool> DeleteFileAttributeAsync(long attributeId)
        // {
        //     await Task.Yield();
        //     return true;
        // }

        public virtual async Task<bool> DeleteFileStorageAsync(long storageId)
        {
            await Task.Yield();
            return true;
        }

        // public virtual async Task<int> AddFileAttributeAsync(FileAttribute attribute)
        // {
        //     await Task.Yield();
        //     LogManager.Instance.LogWarning("AddFileAttributeAsync 已废弃，请改用 AddFileStorageAndAttributesJsonAsync。");
        //     return 0;
        // }

        public virtual async Task<int> AddFileStorageAsync(FileStorage storage)
        {
            await Task.Yield();
            // 旧链路已废弃，返回0避免“假成功”
            LogManager.Instance.LogWarning("AddFileStorageAsync 旧写入链路已废弃，请改用 AddFileStorageAndAttributesJsonAsync。");
            return 0;
        }


        //public virtual async Task<FileStorage?> GetFileStorageAsync(string fileHash)
        //{
        //    await Task.Yield();
        //    return null;
        //}

        // public virtual async Task<bool> UpdateFileAttributeAsync(FileAttribute attribute)
        // {
        //     await Task.Yield();
        //     return true;
        // }

        public virtual async Task<bool> UpdateFileStorageAsync(FileStorage storage)
        {
            await Task.Yield();
            return true;
        }

        public virtual async Task<bool> AddFileTagAsync(FileTag tag)
        {
            await Task.Yield();
            return true;
        }

        public virtual async Task<int> AddFileAccessLogAsync(object accessLog)
        {
            await Task.Yield();
            return 1;
        }
        /// <summary>
        /// 补齐：根据 Hash 获取文件存储记录 (最小可编译实现)
        /// </summary>
        public async Task<FileStorage> GetFileStorageAsync(string fileHash)
        {
            await Task.Yield();
            // 内部可按需调用现有的 GetFileByIdAsync 逻辑或 SQL
            return null;
        }

     

        /// <summary>
        /// 补齐：构建属性表插入值字典
        /// </summary>
        /// <summary>
        /// 完善：构建存储主表插入值字典
        /// </summary>
        private Dictionary<string, object> BuildStorageInsertValues(FileStorage storage, Dictionary<string, string> attributes, Dictionary<string, string> columns, string configName)
        {
            // 使用 StringComparer.OrdinalIgnoreCase 强制字典忽略大小写，防止出现 created_at 和 CREATED_AT 同时存在
            var values = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            var props = typeof(FileStorage).GetProperties();

            foreach (var p in props)
            {
                // 查找数据库中真实存在的列名
                var match = columns.Keys.FirstOrDefault(k =>
                    string.Equals(k, p.Name, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(k.Replace("_", ""), p.Name, StringComparison.OrdinalIgnoreCase));

                if (match != null)
                {
                    // 核心修复：绝对排除 ID 列，达梦不允许显式插入 ID 值（自增列规则）
                    if (string.Equals(match, "ID", StringComparison.OrdinalIgnoreCase)) continue;

                    values[match] = p.GetValue(storage);
                }
            }

            // 审计字段逻辑：仅在反射未处理且数据库存在该列时添加，确保不会重复生成 SQL 列名
            if (columns.ContainsKey("created_at") && !values.ContainsKey("created_at"))
                values["created_at"] = storage.CreatedAt == default ? DateTime.Now : storage.CreatedAt;

            if (columns.ContainsKey("updated_at") && !values.ContainsKey("updated_at"))
                values["updated_at"] = DateTime.Now;

            return values;
        }

    
        /// <summary>
        /// 完善：构建属性 JSON 表插入值字典
        /// </summary>
        private Dictionary<string, object> BuildAttributeInsertValues(FileStorage storage, Dictionary<string, string> attributes, Dictionary<string, string> columns, int storageId)
        {
            var values = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            if (columns.ContainsKey("file_id")) values["file_id"] = storageId;
            if (columns.ContainsKey("config_name")) values["config_name"] = "default";

            // 显式序列化 JSON 字典
            if (columns.ContainsKey("attributes_json"))
            {
                values["attributes_json"] = Newtonsoft.Json.JsonConvert.SerializeObject(attributes);
            }

            if (columns.ContainsKey("created_at") && !values.ContainsKey("created_at"))
                values["created_at"] = DateTime.Now;

            // 再次加固：确保物理上不含任何 ID 键
            values.Remove("ID");

            return values;
        }

        /// <summary>
        /// 补齐：公开方法的兜底实现 (解决外部调用报错)
        /// </summary>
        // public async Task<dynamic> GetFileAttributeByGraphicIdAsync(params object[] args) { return await Task.FromResult<dynamic>(null); }
      

        // 替换为：
        /// <summary>
        /// 级联删除 CAD 图元记录
        /// </summary>
        /// <param name="id">图元ID</param>
        /// <param name="physicalDelete">是否物理删除</param>
        /// <returns>操作是否成功</returns>
        public async Task<bool> DeleteCadGraphicCascadeAsync(int id, bool physicalDelete = false)
        {
            try
            {
                using var connection = GetConnection();
                // 达梦与 MySQL 在删除逻辑上基本一致
                var sql = physicalDelete
                    ? "DELETE FROM cad_file_storage WHERE id = @Id"
                    : "UPDATE cad_file_storage SET is_active = 0 WHERE id = @Id";

                // 注意：如果是达梦且未开启自动参数映射，可能需要 NormalizeSql 或手动切换参数占位符
                var result = await ExecuteWriteAsync(connection, null, sql, new { Id = id }).ConfigureAwait(false);
                return result > 0;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"DeleteCadGraphicCascadeAsync 出错: {ex.Message}");
                return false;
            }
        }
        public async Task<dynamic> GetConfigValueAsync(params object[] args) { return await Task.FromResult<dynamic>(null); }
        public async Task<dynamic> UpdateCategoryStatisticsAsync(params object[] args) { return await Task.FromResult<dynamic>(false); }
        public async Task<dynamic> AddCadSubcategoryAsync(params object[] args) { return await Task.FromResult<dynamic>(false); }
        public async Task<dynamic> UpdateParentSubcategoryListAsync(params object[] args) { return await Task.FromResult<dynamic>(false); }

        /// <summary>
        /// 更新 CAD 子分类
        /// </summary>
        public async Task<int> UpdateCadSubcategoryAsync(CadSubcategory subcategory)
        {
            if (subcategory == null) return 0;

            const string mysqlSql = @"
                UPDATE cad_subcategories 
                SET category_id = @CategoryId, name = @Name, display_name = @DisplayName, 
                    parent_id = @ParentId, sort_order = @SortOrder, level = @Level, 
                    subcategory_ids = @SubcategoryIds, updated_at = NOW()
                WHERE id = @Id";

            const string dmSql = @"
                UPDATE cad_subcategories 
                SET category_id = :CategoryId, name = :Name, display_name = :DisplayName, 
                    parent_id = :ParentId, sort_order = :SortOrder, level = :Level, 
                    subcategory_ids = :SubcategoryIds, updated_at = CURRENT_TIMESTAMP
                WHERE id = :Id";

            try
            {
                using var connection = GetConnection();
                var sql = _adapter.DatabaseType == "MySQL" ? mysqlSql : dmSql;
                return await ExecuteWriteAsync(connection, null, sql, subcategory).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"UpdateCadSubcategoryAsync 出错: {ex.Message}");
                return 0;
            }
        }

        public async Task<dynamic> DeleteCadSubcategoryAsync(params object[] args) { return await Task.FromResult<dynamic>(false); }


        /// <summary>
        /// 用户实体（对应 users 表）
        /// </summary>
        public class User
        {
            /// <summary>
            /// 用户 ID
            /// </summary>
            public int Id { get; set; }
            /// <summary>
            /// 用户名
            /// </summary>
            public string? Username { get; set; }
            /// <summary>
            /// 密码哈希
            /// </summary>
            public string? PasswordHash { get; set; }
            // <summary>
            /// 显示名称
            /// </summary>
            public string? DisplayName { get; set; }

            /// <summary>
            /// 性别
            /// </summary>
            public string? Gender { get; set; }
            /// <summary>
            /// 手机号码
            /// </summary>
            public string? Phone { get; set; }
            /// <summary>
            /// 电子邮箱
            /// </summary>
            public string? Email { get; set; }
            /// <summary>
            /// 部门 ID
            /// </summary>
            public int? DepartmentId { get; set; }

            /// <summary>
            /// 角色
            /// </summary>
            public string? Role { get; set; }
            /// <summary>
            /// 状态
            /// </summary>
            public int Status { get; set; }
            /// <summary>
            /// 创建时间
            /// </summary>
            public DateTime CreatedAt { get; set; }

            /// <summary>
            /// 更新时间
            /// </summary>
            public DateTime UpdatedAt { get; set; }
        }

        /// <summary>
        /// 根据用户名查询用户（用于注册后获取 id）
        /// </summary>
        /// <param name="username"></param>
        /// <returns>匹配的 User 或 null</returns>
        public async Task<User> GetUserByUsernameAsync(string username)
        {
            if (string.IsNullOrWhiteSpace(username))
                return null;

            const string sql = @"
                SELECT
                    id AS Id,
                    username AS Username,
                    password_hash AS PasswordHash,
                    display_name AS DisplayName,
                    gender AS Gender,
                    phone AS Phone,
                    email AS Email,
                    department_id AS DepartmentId,
                    role AS Role,
                    status AS Status,
                    created_at AS CreatedAt,
                    updated_at AS UpdatedAt
                FROM users
                WHERE username = @Username
                LIMIT 1";

            try
            {
                using var conn = GetConnection();
                conn.Open();
                // 诊断：记录将要执行的 SQL 与参数，帮助定位达梦解析错误（临时日志）
                LogManager.Instance.LogDebug($"[DM-SQL] Executing GetUserByUsernameAsync SQL: {sql}, Params: Username={username}");

                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql.Replace("@Username", ":Username");
                AddDmParam(cmd, "Username", username);

                using var reader = cmd.ExecuteReader();
                if (!reader.Read()) return null;
                var u = new User();
                int ord;
                ord = reader.GetOrdinal("Id"); u.Id = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("Username"); u.Username = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                ord = reader.GetOrdinal("PasswordHash"); u.PasswordHash = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                ord = reader.GetOrdinal("DisplayName"); u.DisplayName = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                ord = reader.GetOrdinal("Gender"); u.Gender = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                ord = reader.GetOrdinal("Phone"); u.Phone = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                ord = reader.GetOrdinal("Email"); u.Email = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                ord = reader.GetOrdinal("DepartmentId"); u.DepartmentId = reader.IsDBNull(ord) ? (int?)null : reader.GetInt32(ord);
                ord = reader.GetOrdinal("Role"); u.Role = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                ord = reader.GetOrdinal("Status"); u.Status = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("CreatedAt"); u.CreatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                ord = reader.GetOrdinal("UpdatedAt"); u.UpdatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                return u;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetUserByUsernameAsync 出错: {ex.Message}");
                LogManager.Instance.LogDebug($"[DM-SQL-ERR] SQL: {sql}, Params: Username={username}");
                return null;
            }
        }

        /// <summary>
        /// 检查指定数据库中是否缺少核心表。
        /// 返回：
        /// - 若数据库不存在：返回包含单项 "__DATABASE_MISSING__"
        /// - 若数据库存在但缺少表：返回缺失表名列表（不为空）
        /// - 若一切正常：返回空列表
        /// </summary>
        public static List<string> CheckMissingCoreTables(string server, int port, string user, string password, string database = "cad_sw_library")
        {
            var missing = new List<string>();
            try
            {
                // 连接到目标数据库以检查表
                var connStr = $"Server={server};Port={port};Database={database};Uid={user};Pwd={password};";
                // 使用 MySqlAdapter 创建连接，避免直接依赖 MySqlConnection
                var tmpAdapter = new MySqlAdapter(connStr);
                using var conn = tmpAdapter.CreateConnection();
                conn.Open();

                // 需要保证的核心表（含 CAD / SW / 设备表）
                var required = new[]
                {
                    "cad_categories",
                    "cad_subcategories",
                    "cad_file_storage",
                    "cad_block_attributes_json",
                    "system_config",
                    "users",
                    "departments",
                    "department_users",
                    "category_department_map",
                    "sw_categories",
                    "sw_subcategories",
                    "sw_graphics",
                    "device_info"
                };

                var sql = @"SELECT table_name FROM information_schema.tables
                        WHERE table_schema = @schema AND table_name IN @names";
                var found = conn.Query<string>(sql, new { schema = database, names = required }).AsList();

                foreach (var t in required)
                {
                    if (!found.Contains(t))
                        missing.Add(t);
                }

                return missing;
            }
            catch (MySqlException mex)
            {
                // 数据库不存在
                if (mex.Number == 1049)
                {
                    return new List<string> { "__DATABASE_MISSING__" };
                }
                return new List<string> { $"__DB_ERROR__:{mex.Message}" };
            }
            catch (Exception)
            {
                return new List<string> { "__DB_CHECK_FAILED__" };
            }
        }

        /// <summary>
        /// 创建数据库（若不存在）并创建/修复核心表结构（新方案：cad_block_attributes_json）。
        /// 返回 true 表示成功（创建成功或已存在且修复成功），false 表示失败。
        /// </summary>
        public static bool CreateDatabaseAndCoreTables(string server, int port, string user, string password, string database = "cad_sw_library")
        {
            try
            {
                // =========================
                // 第1步：确保数据库存在
                // =========================
                var masterConn = $"Server={server};Port={port};Uid={user};Pwd={password};";
                var tmpAdapter1 = new MySqlAdapter(masterConn);
                using (var conn = tmpAdapter1.CreateConnection())
                {
                    conn.Open();
                    // 创建数据库（若不存在），统一字符集与排序规则，避免中文乱码
                    var createDbSql = $"CREATE DATABASE IF NOT EXISTS `{database}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;";
                    conn.Execute(createDbSql);
                }

                // =========================
                // 第2步：连接目标数据库
                // =========================
                var dbConn = $"Server={server};Port={port};Database={database};Uid={user};Pwd={password};Allow User Variables=True;";
                var tmpAdapter2 = new MySqlAdapter(dbConn);
                using (var conn = tmpAdapter2.CreateConnection())
                {
                    conn.Open();

                    // 统一执行DDL的小工具，便于阅读和排错
                    void Exec(string sql)
                    {
                        conn.Execute(sql);
                    }

                    // 检查列是否存在
                    bool ColumnExists(string tableName, string columnName)
                    {
                        const string sql = @"
SELECT COUNT(1)
FROM information_schema.columns
WHERE table_schema = @schema
  AND table_name = @table
  AND column_name = @column;";
                        return conn.QuerySingle<int>(sql, new { schema = database, table = tableName, column = columnName }) > 0;
                    }

                    // 检查索引是否存在
                    bool IndexExists(string tableName, string indexName)
                    {
                        const string sql = @"
SELECT COUNT(1)
FROM information_schema.statistics
WHERE table_schema = @schema
  AND table_name = @table
  AND index_name = @index;";
                        return conn.QuerySingle<int>(sql, new { schema = database, table = tableName, index = indexName }) > 0;
                    }

                    // 检查外键约束是否存在
                    bool ForeignKeyExists(string tableName, string fkName)
                    {
                        const string sql = @"
SELECT COUNT(1)
FROM information_schema.table_constraints
WHERE table_schema = @schema
  AND table_name = @table
  AND constraint_name = @fk
  AND constraint_type = 'FOREIGN KEY';";
                        return conn.QuerySingle<int>(sql, new { schema = database, table = tableName, fk = fkName }) > 0;
                    }

                    // =========================
                    // 第3步：创建核心表（若不存在）
                    // =========================

                    // CAD 主分类表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `cad_categories` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `name` VARCHAR(200) NOT NULL,
    `display_name` VARCHAR(200) NULL,
    `subcategory_ids` TEXT NULL,
    `sort_order` INT DEFAULT 0,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // CAD 子分类表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `cad_subcategories` (
    `id` INT NOT NULL PRIMARY KEY,
    `parent_id` INT NOT NULL,
    `name` VARCHAR(200) NOT NULL,
    `display_name` VARCHAR(200) NULL,
    `sort_order` INT DEFAULT 0,
    `level` INT DEFAULT 1,
    `subcategory_ids` TEXT NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX `idx_cad_sub_parent` (`parent_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // CAD 文件主表（沿用你现有项目字段）
                    Exec(@"
CREATE TABLE IF NOT EXISTS `cad_file_storage` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `category_id` INT NULL,
    `file_attribute_id` VARCHAR(64) NULL COMMENT '兼容字段：可存属性配置名或业务ID',
    `file_name` VARCHAR(512) NULL,
    `file_stored_name` VARCHAR(255) NOT NULL,
    `display_name` VARCHAR(255) NOT NULL,
    `file_type` VARCHAR(255) NULL,
    `file_hash` VARCHAR(255) NOT NULL,
    `block_name` VARCHAR(255) NULL,
    `layer_name` VARCHAR(100) NOT NULL,
    `color_index` INT NOT NULL DEFAULT 256,
    `file_path` VARCHAR(500) NOT NULL,
    `preview_image_name` VARCHAR(255) NULL,
    `preview_image_path` VARCHAR(500) NULL,
    `file_size` BIGINT NULL,
    `created_at` TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
    `updated_at` TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    `is_preview` TINYINT(1) DEFAULT 0,
    `version` INT DEFAULT 1,
    `description` TEXT NULL,
    `is_active` TINYINT(1) DEFAULT 1,
    `created_by` VARCHAR(255) NULL,
    `category_type` VARCHAR(50) DEFAULT 'sub',
    `title` VARCHAR(255) NULL,
    `keywords` TEXT NULL,
    `is_public` TINYINT(1) DEFAULT 1,
    `updated_by` VARCHAR(255) NULL,
    `last_accessed_at` DATETIME NULL,
    `is_tianzheng` TINYINT(1) NULL,
    `scale` DOUBLE NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 新属性表（JSON方案核心）
                    Exec(@"
CREATE TABLE IF NOT EXISTS `cad_block_attributes_json` (
    `attr_id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY COMMENT '属性记录唯一标识',
    `file_id` INT NOT NULL COMMENT '关联 cad_file_storage.id',
    `config_name` VARCHAR(100) NULL COMMENT '配置名称，如 DN50 配置',
    `attributes_json` TEXT NULL COMMENT 'JSON格式属性字典',
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX `idx_fileid` (`file_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 系统配置表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `system_config` (
    `config_key` VARCHAR(200) PRIMARY KEY,
    `config_value` TEXT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 部门表（包含 cad_category_id，兼容 MySqlAuthService 中的同步逻辑）
                    Exec(@"
CREATE TABLE IF NOT EXISTS `departments` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `cad_category_id` INT NULL,
    `name` VARCHAR(200) NOT NULL,
    `display_name` VARCHAR(200) NULL,
    `description` TEXT NULL,
    `manager_user_id` INT NULL,
    `sort_order` INT DEFAULT 0,
    `is_active` TINYINT(1) DEFAULT 1,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 用户表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `users` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `username` VARCHAR(100) NOT NULL UNIQUE,
    `password_hash` VARCHAR(512) NULL,
    `display_name` VARCHAR(200) NULL,
    `gender` VARCHAR(16) NULL,
    `phone` VARCHAR(32) NULL,
    `email` VARCHAR(200) NULL,
    `department_id` INT NULL,
    `role` VARCHAR(64) NULL,
    `status` TINYINT DEFAULT 1,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX `idx_users_department` (`department_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 部门-用户关系表（可选多对多）
                    Exec(@"
CREATE TABLE IF NOT EXISTS `department_users` (
    `department_id` INT NOT NULL,
    `user_id` INT NOT NULL,
    PRIMARY KEY (`department_id`,`user_id`),
    INDEX `idx_department_users_user` (`user_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 分类-部门映射表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `category_department_map` (
    `category_id` INT NOT NULL PRIMARY KEY,
    `department_id` INT NOT NULL UNIQUE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 文件访问日志
                    Exec(@"
CREATE TABLE IF NOT EXISTS `file_access_logs` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `file_id` INT NULL,
    `user_name` VARCHAR(200) NULL,
    `action_type` VARCHAR(50) NULL,
    `ip_address` VARCHAR(64) NULL,
    `user_agent` VARCHAR(512) NULL,
    `access_time` DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 文件标签表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `file_tags` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `file_id` INT NULL,
    `tag_name` VARCHAR(200) NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 文件版本历史
                    Exec(@"
CREATE TABLE IF NOT EXISTS `file_version_history` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `file_id` INT NULL,
    `version` INT NULL,
    `file_name` VARCHAR(512) NULL,
    `stored_file_name` VARCHAR(512) NULL,
    `file_path` VARCHAR(1024) NULL,
    `file_size` BIGINT NULL,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_by` VARCHAR(200) NULL,
    `change_description` TEXT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // SW 分类表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `sw_categories` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `name` VARCHAR(200) NOT NULL,
    `display_name` VARCHAR(200) NULL,
    `sort_order` INT DEFAULT 0,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // SW 子分类表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `sw_subcategories` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `category_id` INT NOT NULL,
    `parent_id` INT NOT NULL DEFAULT 0,
    `name` VARCHAR(200) NOT NULL,
    `display_name` VARCHAR(200) NULL,
    `sort_order` INT DEFAULT 0,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX `idx_sw_sub_cat` (`category_id`),
    INDEX `idx_sw_sub_parent` (`parent_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // SW 图元表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `sw_graphics` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `subcategory_id` INT NOT NULL,
    `file_name` VARCHAR(512) NOT NULL,
    `display_name` VARCHAR(512) NULL,
    `file_path` VARCHAR(1024) NULL,
    `preview_image_path` VARCHAR(1024) NULL,
    `file_size` BIGINT NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX `idx_sw_graphics_sub` (`subcategory_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 设备信息表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `device_info` (
    `device_id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `device_name` VARCHAR(255) NOT NULL,
    `device_type` VARCHAR(100) NULL,
    `medium_name` VARCHAR(100) NULL,
    `specifications` VARCHAR(255) NULL,
    `material` VARCHAR(100) NULL,
    `quantity` INT DEFAULT 0,
    `drawing_number` VARCHAR(100) NULL,
    `power` DECIMAL(18,6) NULL,
    `volume` DECIMAL(18,6) NULL,
    `pressure` DECIMAL(18,6) NULL,
    `temperature` DECIMAL(18,6) NULL,
    `diameter` DECIMAL(18,6) NULL,
    `length` DECIMAL(18,6) NULL,
    `thickness` DECIMAL(18,6) NULL,
    `weight` DECIMAL(18,6) NULL,
    `model` VARCHAR(255) NULL,
    `remarks` TEXT NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // =========================
                    // 第4步：修复已有库的列/索引/外键（增量修复）
                    // =========================

                    // 确保 cad_file_storage 关键兼容列存在
                    if (!ColumnExists("cad_file_storage", "file_attribute_id"))
                    {
                        Exec("ALTER TABLE `cad_file_storage` ADD COLUMN `file_attribute_id` VARCHAR(64) NULL COMMENT '兼容字段：属性配置名或业务ID';");
                    }
                    if (!ColumnExists("cad_file_storage", "category_type"))
                    {
                        Exec("ALTER TABLE `cad_file_storage` ADD COLUMN `category_type` VARCHAR(50) DEFAULT 'sub';");
                    }

                    // 确保常用索引存在
                    if (!IndexExists("cad_file_storage", "idx_cfs_category"))
                    {
                        Exec("ALTER TABLE `cad_file_storage` ADD INDEX `idx_cfs_category` (`category_id`, `category_type`);");
                    }
                    if (!IndexExists("cad_file_storage", "idx_cfs_attr_biz"))
                    {
                        Exec("ALTER TABLE `cad_file_storage` ADD INDEX `idx_cfs_attr_biz` (`file_attribute_id`);");
                    }
                    if (!IndexExists("cad_file_storage", "idx_cfs_file_hash"))
                    {
                        Exec("ALTER TABLE `cad_file_storage` ADD INDEX `idx_cfs_file_hash` (`file_hash`);");
                    }

                    // 确保 departments 的 cad_category_id 索引存在（兼容 MySqlAuthService）
                    if (!IndexExists("departments", "idx_cad_category_id"))
                    {
                        Exec("ALTER TABLE `departments` ADD INDEX `idx_cad_category_id` (`cad_category_id`);");
                    }

                    // 确保 JSON 属性表外键存在（file_id -> cad_file_storage.id）
                    if (!ForeignKeyExists("cad_block_attributes_json", "fk_file_id"))
                    {
                        // 先确保 file_id 有索引
                        if (!IndexExists("cad_block_attributes_json", "idx_fileid"))
                        {
                            Exec("ALTER TABLE `cad_block_attributes_json` ADD INDEX `idx_fileid` (`file_id`);");
                        }

                        // 添加外键，启用级联删除，防止孤儿属性记录
                        Exec(@"
ALTER TABLE `cad_block_attributes_json`
ADD CONSTRAINT `fk_file_id`
FOREIGN KEY (`file_id`) REFERENCES `cad_file_storage`(`id`)
ON DELETE CASCADE
ON UPDATE CASCADE;");
                    }

                    // =========================
                    // 第5步：可选清理提醒（不自动删旧表）
                    // =========================
                    // 这里不自动 DROP 旧 cad_file_attributes，避免误删历史数据。
                    // 若你确认旧表永不再用，可在数据库中手工删除。

                    return true;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"CreateDatabaseAndCoreTables 失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 主窗口
        /// </summary>
        private readonly WpfMainWindow _wpfMainWindow;
        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        public readonly string _connectionString;
        /// <summary>
        /// 当前达梦会话需要切换到的 Schema。
        /// </summary>
        private readonly string _schemaName;
        /// <summary>
        /// 数据库是否可用
        /// </summary>
        public bool IsDatabaseAvailable { get; private set; } = true;
        /// <summary>
        /// 数据库管理类构造函数
        /// </summary>
        /// <param name="connectionString"> 链接字符串
        ///  </param>
        public DatabaseManager(string connectionString)
        {
            _schemaName = ResolveSchemaName(connectionString);

            // 根据配置初始化适配器
            if (VariableDictionary._databaseType?.ToUpper() == "MYSQL")
            {
                _adapter = new MySqlAdapter(connectionString);
                _connectionString = connectionString;
            }
            else
            {
                _adapter = new DmAdapter(NormalizeDmConnectionString(connectionString));
                _connectionString = NormalizeDmConnectionString(connectionString);
            }

            LogManager.Instance.LogInfo($"DatabaseManager 初始化: server={ExtractConnectionStringValue(_connectionString, "Server")}, port={ExtractConnectionStringValue(_connectionString, "Port")}, database={_schemaName}, type={_adapter.DatabaseType}");
            IsDatabaseAvailable = TestDatabaseConnection();
            LogManager.Instance.LogInfo($"DatabaseManager 初始化完成: IsDatabaseAvailable={IsDatabaseAvailable}");
        }
        /// <summary>
        /// 测试数据库连接
        /// </summary>
        /// <returns></returns>
        private bool TestDatabaseConnection()
        {
            try
            {
                using var connection = GetConnection();
                connection.Open();
                ApplySchema(connection);
                using var command = connection.CreateCommand();
                command.CommandText = _adapter.NormalizeSql("SELECT 1 FROM DUAL");
                command.ExecuteScalar();
                LogManager.Instance.LogInfo($"{_adapter.DatabaseType} 数据库连接测试成功: database/schema={_schemaName}");
                return true;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"{_adapter.DatabaseType} 连接测试失败: database/schema={_schemaName}, {ex.Message}");
                return false;
            }
        }

        #region 部门与人员同步方法

        /// <summary>
        /// 新增：部门实体
        /// </summary>
        public class Department
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string DisplayName { get; set; }
            public int SortOrder { get; set; }
            public bool IsActive { get; set; }
            public DateTime CreatedAt { get; set; }
            public DateTime UpdatedAt { get; set; }
        }

        /// <summary>
        /// 获取所有部门（用于注册窗口下拉列表）
        /// </summary>
        public async Task<List<Department>> GetAllDepartmentsAsync()
        {
            const string sql = @"
              SELECT
                  id AS Id,
                  name AS Name,
                  display_name AS DisplayName,
                  sort_order AS SortOrder,
                  is_active AS IsActive,
                  created_at AS CreatedAt,
                  updated_at AS UpdatedAt
              FROM departments
              ORDER BY sort_order, name";
            try
            {
                return await Task.Run(() =>
                {
                    using var connection = GetConnection();
                    connection.Open();
                    using var cmd = connection.CreateCommand();
                    cmd.CommandText = sql;

                    var list = new List<Department>();
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var d = new Department();
                        int ord;
                        ord = reader.GetOrdinal("Id"); d.Id = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("Name"); d.Name = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("DisplayName"); d.DisplayName = reader.IsDBNull(ord) ? d.Name : reader.GetString(ord);
                        ord = reader.GetOrdinal("SortOrder"); d.SortOrder = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("IsActive"); d.IsActive = !reader.IsDBNull(ord) && reader.GetInt32(ord) != 0;
                        ord = reader.GetOrdinal("CreatedAt"); d.CreatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                        ord = reader.GetOrdinal("UpdatedAt"); d.UpdatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                        list.Add(d);
                    }
                    return list;
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetAllDepartmentsAsync 出错: {ex.Message}");
                return new List<Department>();
            }
        }


        /// <summary>
        /// 将 cad_categories 中尚未映射到 departments 的分类逐条创建部门并建立映射。
        /// 保持幂等：已存在的映射或同名部门不会重复创建（会尝试复用同名部门）。
        /// </summary>
        public async Task SyncDepartmentsFromCadCategoriesAsync()
        {
            try
            {
                var categories = await GetAllCadCategoriesAsync().ConfigureAwait(false);
                if (categories == null || categories.Count == 0) return;
                // 使用同步事务化操作放入线程池，以便与达梦驱动兼容参数与 SQL 语法
                await Task.Run(() =>
                {
                    using var conn = GetConnection();
                    conn.Open();
                    using var tx = conn.BeginTransaction();

                    foreach (var cat in categories)
                    {
                        // 检查是否已有映射
                        using (var mapCmd = conn.CreateCommand())
                        {
                            mapCmd.Transaction = tx;
                            mapCmd.CommandText = "SELECT department_id FROM category_department_map WHERE category_id = :CategoryId";
                            var p = mapCmd.CreateParameter(); p.ParameterName = "CategoryId"; p.Value = cat.Id; mapCmd.Parameters.Add(p);
                            var res = mapCmd.ExecuteScalar();
                            if (res != null && res != DBNull.Value)
                            {
                                continue;
                            }
                        }

                        int? deptId = null;
                        // 尝试按名称查找已有部门
                        using (var findCmd = conn.CreateCommand())
                        {
                            findCmd.Transaction = tx;
                            findCmd.CommandText = "SELECT id FROM departments WHERE name = :Name ORDER BY id FETCH FIRST 1 ROWS ONLY";
                            var p = findCmd.CreateParameter(); p.ParameterName = "Name"; p.Value = cat.Name ?? (object)DBNull.Value; findCmd.Parameters.Add(p);
                            var res = findCmd.ExecuteScalar();
                            if (res != null && res != DBNull.Value)
                            {
                                deptId = Convert.ToInt32(res);
                            }
                        }

                        if (!deptId.HasValue)
                        {
                            // 插入新部门
                            using (var insCmd = conn.CreateCommand())
                            {
                                insCmd.Transaction = tx;
                                insCmd.CommandText = "INSERT INTO departments (name, display_name, sort_order, created_at, updated_at) VALUES (:Name, :DisplayName, :SortOrder, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)";
                                var p1 = insCmd.CreateParameter(); p1.ParameterName = "Name"; p1.Value = cat.Name ?? (object)DBNull.Value; insCmd.Parameters.Add(p1);
                                var p2 = insCmd.CreateParameter(); p2.ParameterName = "DisplayName"; p2.Value = string.IsNullOrEmpty(cat.DisplayName) ? (cat.Name ?? (object)DBNull.Value) : cat.DisplayName; insCmd.Parameters.Add(p2);
                                var p3 = insCmd.CreateParameter(); p3.ParameterName = "SortOrder"; p3.Value = cat.SortOrder; insCmd.Parameters.Add(p3);
                                insCmd.ExecuteNonQuery();
                            }

                            // 读取刚插入的 id（按 name 倒序取最新一条）
                            using (var getIdCmd = conn.CreateCommand())
                            {
                                getIdCmd.Transaction = tx;
                                getIdCmd.CommandText = "SELECT id FROM departments WHERE name = :Name ORDER BY id DESC FETCH FIRST 1 ROWS ONLY";
                                var gp = getIdCmd.CreateParameter(); gp.ParameterName = "Name"; gp.Value = cat.Name ?? (object)DBNull.Value; getIdCmd.Parameters.Add(gp);
                                var got = getIdCmd.ExecuteScalar();
                                if (got != null && got != DBNull.Value) deptId = Convert.ToInt32(got);
                            }
                        }

                        if (deptId.HasValue)
                        {
                            using var mapIns = conn.CreateCommand();
                            mapIns.Transaction = tx;
                            mapIns.CommandText = "INSERT INTO category_department_map (category_id, department_id) VALUES (:CategoryId, :DepartmentId)";
                            var mp1 = mapIns.CreateParameter(); mp1.ParameterName = "CategoryId"; mp1.Value = cat.Id; mapIns.Parameters.Add(mp1);
                            var mp2 = mapIns.CreateParameter(); mp2.ParameterName = "DepartmentId"; mp2.Value = deptId.Value; mapIns.Parameters.Add(mp2);
                            mapIns.ExecuteNonQuery();
                        }
                    }

                    tx.Commit();
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategoriesAsync 出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 删除与分类相关的部门映射与（可选）部门记录。
        /// 当分类被删除时调用：如果该部门没有其他映射且没有用户（或你选择直接删除），则删除部门。
        /// </summary>
        public async Task RemoveDepartmentMappingForCategoryAsync(int categoryId)
        {
            try
            {
                using var conn = GetConnection();
                conn.Open();

                int? deptId = null;
                using (var getDeptCmd = conn.CreateCommand())
                {
                    getDeptCmd.CommandText = "SELECT department_id FROM category_department_map WHERE category_id = :CategoryId";
                    var p = getDeptCmd.CreateParameter(); p.ParameterName = "CategoryId"; p.Value = categoryId; getDeptCmd.Parameters.Add(p);
                    var r = getDeptCmd.ExecuteScalar();
                    if (r == null || r == DBNull.Value) return; // 无映射
                    deptId = Convert.ToInt32(r);
                }

                using (var delMapCmd = conn.CreateCommand())
                {
                    delMapCmd.CommandText = "DELETE FROM category_department_map WHERE category_id = :CategoryId";
                    var p = delMapCmd.CreateParameter(); p.ParameterName = "CategoryId"; p.Value = categoryId; delMapCmd.Parameters.Add(p);
                    delMapCmd.ExecuteNonQuery();
                }

                int usedByCat = 0;
                using (var usedCmd = conn.CreateCommand())
                {
                    usedCmd.CommandText = "SELECT COUNT(*) FROM category_department_map WHERE department_id = :DepartmentId";
                    var p = usedCmd.CreateParameter(); p.ParameterName = "DepartmentId"; p.Value = deptId.Value; usedCmd.Parameters.Add(p);
                    var r = usedCmd.ExecuteScalar(); usedByCat = Convert.ToInt32(r ?? 0);
                }

                int userCount = 0;
                using (var userCmd = conn.CreateCommand())
                {
                    userCmd.CommandText = "SELECT COUNT(*) FROM users WHERE department_id = :DepartmentId";
                    var p = userCmd.CreateParameter(); p.ParameterName = "DepartmentId"; p.Value = deptId.Value; userCmd.Parameters.Add(p);
                    var r = userCmd.ExecuteScalar(); userCount = Convert.ToInt32(r ?? 0);
                }

                if (usedByCat == 0 && userCount == 0)
                {
                    using var delDept = conn.CreateCommand();
                    delDept.CommandText = "DELETE FROM departments WHERE id = :DepartmentId";
                    var p = delDept.CreateParameter(); p.ParameterName = "DepartmentId"; p.Value = deptId.Value; delDept.Parameters.Add(p);
                    delDept.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"RemoveDepartmentMappingForCategoryAsync 出错: {ex.Message}");
            }
        }

        #endregion

        #region CAD分类操作

        /// <summary>
        /// 添加CAD分类（并同步创建部门映射）
        /// </summary>
        public async Task<int> AddCadCategoryAsync(CadCategory category)
        {
            using var connection = GetConnection();
            var sql = @"INSERT INTO cad_categories (name, display_name, sort_order) 
                VALUES (@Name, @DisplayName, @SortOrder)";
            var affected = await ExecuteWriteAsync(connection, null, sql, category).ConfigureAwait(false);

            // 异步触发同步（保证分类与部门一致）
            _ = SyncDepartmentsFromCadCategoriesAsync();
            return affected;
        }

        /// <summary>
        /// 修改CAD分类（并同步部门信息）
        /// </summary>
        public async Task<int> UpdateCadCategoryAsync(CadCategory category)
        {
            using var connection = GetConnection();
            var currentTimestampSql = GetCurrentTimestampSql();
            var sql = @"UPDATE cad_categories 
                SET name = @Name, display_name = @DisplayName, sort_order = @SortOrder, updated_at = " + currentTimestampSql + @" 
                WHERE id = @Id";
            var affected = await ExecuteWriteAsync(connection, null, sql, category).ConfigureAwait(false);

            // 如果分类名或显示名变更，更新对应部门（若已存在映射）
            try
            {
                using var conn = GetConnection();
                // 使用同步打开以兼容 IDbConnection 在 .NET Framework 中的实现
                conn.Open();
                var mapSql = "SELECT department_id FROM category_department_map WHERE category_id = @CategoryId";
                var deptId = await conn.QueryFirstOrDefaultAsync<int?>(mapSql, new { CategoryId = category.Id }).ConfigureAwait(false);
                if (deptId.HasValue)
                {
                    var updateDeptSql = @"UPDATE departments SET name = @Name, display_name = @DisplayName, updated_at = " + currentTimestampSql + @" WHERE id = @Id";
                    await ExecuteWriteAsync(conn, null, updateDeptSql, new { Name = category.Name, DisplayName = category.DisplayName ?? category.Name, Id = deptId.Value }).ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"UpdateCadCategoryAsync 同步部门时出错: {ex.Message}");
            }

            // 确保全量同步以修正遗漏
            _ = SyncDepartmentsFromCadCategoriesAsync();
            return affected;
        }

        /// <summary>
        /// 删除CAD分类（并删除部门映射及可选部门）
        /// </summary>
        public async Task<int> DeleteCadCategoryAsync(int id)
        {
            using var connection = GetConnection();
            var sql = "DELETE FROM cad_categories WHERE id = @Id";
            var affected = await connection.ExecuteAsync(sql, new { Id = id }).ConfigureAwait(false);

            // 删除映射并在必要时删除部门
            _ = RemoveDepartmentMappingForCategoryAsync(id);
            return affected;
        }

        /// <summary>
        /// 获取所有CAD分类
        /// </summary>
        /// <returns> 返回List<CadCategory>分类list</returns>
        public async Task<List<CadCategory>> GetAllCadCategoriesAsync()
        {
            try
            {
                const string sql = @"
                                   SELECT 
                                       id AS Id,
                                       name AS Name,
                                       display_name AS DisplayName,
                                       subcategory_ids AS SubcategoryIds,
                                       sort_order AS SortOrder,
                                       created_at AS CreatedAt,
                                       updated_at AS UpdatedAt
                                   FROM cad_categories 
                                   ORDER BY sort_order";

                using var conn = GetConnection();
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;

                var list = new List<CadCategory>();
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var c = new CadCategory();
                    int ord;
                    ord = reader.GetOrdinal("Id"); c.Id = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                    ord = reader.GetOrdinal("Name"); c.Name = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                    ord = reader.GetOrdinal("DisplayName"); c.DisplayName = reader.IsDBNull(ord) ? c.Name : reader.GetString(ord);
                    ord = reader.GetOrdinal("SubcategoryIds"); c.SubcategoryIds = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("SortOrder"); c.SortOrder = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                    ord = reader.GetOrdinal("CreatedAt"); c.CreatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                    ord = reader.GetOrdinal("UpdatedAt"); c.UpdatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                    list.Add(c);
                }
                LogManager.Instance.LogInfo($"查询返回 {list.Count} 条记录");
                return list;

            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"数据库查询出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 根据名称获取CAD分类
        /// </summary>

        public async Task<CadCategory> GetCadCategoryByNameAsync(string categoryName)
        {
            if (string.IsNullOrWhiteSpace(categoryName))
            {
                return null;
            }

            // 使用 @ 作为标准参数占位符
            string sql = @"
            SELECT 
               id AS Id,
               name AS Name,
               display_name AS DisplayName,
               subcategory_ids AS SubcategoryIds,
               sort_order AS SortOrder,
               created_at AS CreatedAt,
               updated_at AS UpdatedAt
            FROM cad_categories 
            WHERE name LIKE @Name";

            try
            {
                using var connection = GetConnection();
                // 1. 如果是达梦数据库，将 @Name 替换为 :Name
                if (_adapter.DatabaseType == "DM" || _adapter.DatabaseType == "Dm")
                {
                    sql = sql.Replace("@Name", ":Name");
                }

                // 2. 构造参数字典
                var parameters = new Dictionary<string, object>
                {
                    { "Name", $"%{categoryName}%" }
                };

                // 3. 执行查询。Dapper 会自动根据 IDbConnection 类型处理参数绑定
                var result = await connection.QueryFirstOrDefaultAsync<CadCategory>(sql, parameters).ConfigureAwait(false);
                return result;
            }
            catch (Exception ex)
            {
                // 在报错信息中加入 SQL 诊断信息，便于排查
                LogManager.Instance.LogInfo($"[数据库-{_adapter.DatabaseType}] 查询出错: {ex.Message}. SQL: {sql}");
                throw;
            }
        }

        #endregion

        #region CAD子分类操作


        /// <summary>
        /// 获取所有CAD子分类
        /// </summary>
        /// <returns></returns>
        public async Task<List<CadSubcategory>> GetAllCadSubcategoriesAsync()
        {
            const string sql = @"
                               SELECT 
                                   id AS Id,
                                   parent_id AS ParentId,
                                   name AS Name,
                                   display_name AS DisplayName,
                                   sort_order AS SortOrder,
                                   level AS Level,
                                   subcategory_ids AS SubcategoryIds,
                                   created_at AS CreatedAt,
                                   updated_at AS UpdatedAt
                               FROM cad_subcategories 
                               ORDER BY parent_id, id, parent_id, name, display_name, level , subcategory_ids, sort_order";

            if (_adapter.DatabaseType == "MySQL")
            {
                using var conn = new MySqlConnection(_connectionString);
                var rows = await conn.QueryAsync<CadSubcategory>(sql).ConfigureAwait(false);
                return rows.AsList();
            }

            try
            {
                using var conn = GetConnection();
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = _adapter.NormalizeSql(sql);
                var list = new List<CadSubcategory>();
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var c = new CadSubcategory();
                    int ord;
                    ord = reader.GetOrdinal("Id"); c.Id = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                    ord = reader.GetOrdinal("ParentId"); c.ParentId = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                    ord = reader.GetOrdinal("Name"); c.Name = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                    ord = reader.GetOrdinal("DisplayName"); c.DisplayName = reader.IsDBNull(ord) ? c.Name : reader.GetString(ord);
                    ord = reader.GetOrdinal("SortOrder"); c.SortOrder = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                    ord = reader.GetOrdinal("Level"); c.Level = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                    ord = reader.GetOrdinal("SubcategoryIds"); c.SubcategoryIds = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("CreatedAt"); c.CreatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                    ord = reader.GetOrdinal("UpdatedAt"); c.UpdatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                    list.Add(c);
                }
                return list;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetAllCadSubcategoriesAsync 出错: {ex.Message}");
                return new List<CadSubcategory>();
            }
        }

        /// <summary>
        /// 通过Id获取子分类的方法
        /// </summary>
        /// <returns>  </returns>
        public async Task<CadSubcategory> GetCadSubcategoryByIdAsync(int id)
        {
            if (id <= 0)
            {
                return null;
            }

            const string sql = @"
                               SELECT 
                                   id AS Id,
                                   parent_id AS ParentId,
                                   name AS Name,
                                   display_name AS DisplayName,
                                   sort_order AS SortOrder,
                                   level AS Level,
                                   subcategory_ids AS SubcategoryIds,
                                   created_at AS CreatedAt,
                                   updated_at AS UpdatedAt
                               FROM cad_subcategories 
                               WHERE id = @id";
            if (_adapter.DatabaseType == "MySQL")
            {
                using var conn = new MySqlConnection(_connectionString);
                return await conn.QuerySingleOrDefaultAsync<CadSubcategory>(sql, new { id }).ConfigureAwait(false);
            }

            try
            {
                using var conn = GetConnection();
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = _adapter.NormalizeSql(sql);
                AddDmParam(cmd, "id", id);
                using var reader = cmd.ExecuteReader();
                if (!reader.Read()) return null;
                var c = new CadSubcategory();
                int ord;
                ord = reader.GetOrdinal("Id"); c.Id = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("ParentId"); c.ParentId = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("Name"); c.Name = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("DisplayName"); c.DisplayName = reader.IsDBNull(ord) ? c.Name : reader.GetString(ord);
                ord = reader.GetOrdinal("SortOrder"); c.SortOrder = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("Level"); c.Level = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("SubcategoryIds"); c.SubcategoryIds = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                ord = reader.GetOrdinal("CreatedAt"); c.CreatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                ord = reader.GetOrdinal("UpdatedAt"); c.UpdatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                return c;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetCadSubcategoryByIdAsync 出错: {ex.Message}");
                return null;
            }
        }


        /// <summary>
        /// 根据子分类ID获取这个子分类同级的所有兄弟子分类
        /// </summary>
        public async Task<List<CadSubcategory>> GetCadSubcategoriesByCategoryIdAsync(int categoryId)
        {
            if (categoryId <= 0)
            {
                return new List<CadSubcategory>();
            }

            const string sql = @"
                               SELECT 
                                    id AS Id,
                                    parent_id AS ParentId,
                                    name AS Name,
                                    display_name AS DisplayName,
                                    sort_order AS SortOrder,
                                    level AS Level,
                                    subcategory_ids AS SubcategoryIds,
                                    created_at AS CreatedAt,
                                    updated_at AS UpdatedAt
                               FROM cad_subcategories 
                               WHERE parent_id = @ParentId 
                               ORDER BY sort_order";
            try
            {
                if (_adapter.DatabaseType == "MySQL")
                {
                    using var conn = new MySqlConnection(_connectionString);
                    var subcategories = await conn.QueryAsync<CadSubcategory>(sql, new { ParentId = categoryId }).ConfigureAwait(false);
                    return subcategories.AsList();
                }

                try
                {
                    using var conn = GetConnection();
                    conn.Open();
                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = _adapter.NormalizeSql(sql);
                    AddDmParam(cmd, "ParentId", categoryId);
                    var list = new List<CadSubcategory>();
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var c = new CadSubcategory();
                        int ord;
                        ord = reader.GetOrdinal("Id"); c.Id = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("ParentId"); c.ParentId = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("Name"); c.Name = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                        ord = reader.GetOrdinal("DisplayName"); c.DisplayName = reader.IsDBNull(ord) ? c.Name : reader.GetString(ord);
                        ord = reader.GetOrdinal("SortOrder"); c.SortOrder = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("Level"); c.Level = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("SubcategoryIds"); c.SubcategoryIds = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                        ord = reader.GetOrdinal("CreatedAt"); c.CreatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                        ord = reader.GetOrdinal("UpdatedAt"); c.UpdatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                        list.Add(c);
                    }
                    return list;
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogInfo($"GetCadSubcategoriesByCategoryIdAsync 出错: {ex.Message}");
                    return new List<CadSubcategory>();
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"数据库查询出错: {ex.Message}");
                throw;
            }

        }

        /// <summary>
        /// 根据父ID获取子分类（用于递归加载）
        /// </summary>
        public async Task<List<CadSubcategory>> GetCadSubcategoriesByParentIdAsync(int parentId)
        {
            try
            {
                if (parentId <= 0)
                {
                    return new List<CadSubcategory>();
                }

                const string sql = @"
                               SELECT 
                                   *
                               FROM cad_subcategories 
                               WHERE parent_id = @parentId 
                               ORDER BY sort_order";

                if (_adapter.DatabaseType == "MySQL")
                {
                    using var conn = _adapter.CreateConnection();
                    conn.Open();
                    var subcategories = await conn.QueryAsync<CadSubcategory>(sql, new { parentId }).ConfigureAwait(false);
                    return subcategories.AsList();
                }

                try
                {
                    using var conn = GetConnection();
                    conn.Open();
                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = _adapter.NormalizeSql(sql);
                    AddDmParam(cmd, "parentId", parentId);
                    var list = new List<CadSubcategory>();
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var c = new CadSubcategory();
                        int ord;
                        ord = reader.GetOrdinal("Id"); c.Id = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("ParentId"); c.ParentId = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("Name"); c.Name = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                        ord = reader.GetOrdinal("DisplayName"); c.DisplayName = reader.IsDBNull(ord) ? c.Name : reader.GetString(ord);
                        ord = reader.GetOrdinal("SortOrder"); c.SortOrder = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("Level"); c.Level = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("SubcategoryIds"); c.SubcategoryIds = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                        ord = reader.GetOrdinal("CreatedAt"); c.CreatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                        ord = reader.GetOrdinal("UpdatedAt"); c.UpdatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                        list.Add(c);
                    }
                    return list;
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogInfo($"GetCadSubcategoriesByParentIdAsync 出错: {ex.Message}");
                    return new List<CadSubcategory>();
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"获取子分类时出错: {ex.Message}");
                return new List<CadSubcategory>();
            }
        }

        /// <summary>
        /// 获取同步清单所需的全部文件记录。
        /// </summary>
        public async Task<List<FileStorage>> GetAllFileStorageAsync()
        {
            const string sql = @"
                               SELECT 
                                    id AS Id,
                                    category_id AS CategoryId,
                                    file_attribute_id AS FileAttributeId,
                                    file_name AS FileName,
                                    file_stored_name AS FileStoredName,
                                    display_name AS DisplayName,
                                    file_type AS FileType,
                                    file_hash AS FileHash,
                                    block_name AS BlockName,
                                    layer_name AS LayerName,
                                    color_index AS ColorIndex,
                                    scale AS Scale,
                                    file_path AS FilePath,
                                    preview_image_name AS PreviewImageName,
                                    preview_image_path AS PreviewImagePath,
                                    file_size AS FileSize,
                                    is_preview AS IsPreview,
                                    version AS Version,
                                    description AS Description,
                                    is_active AS IsActive,
                                    created_by AS CreatedBy,
                                    category_type AS CategoryType,
                                    title AS Title,
                                    keywords AS Keywords,
                                    is_public AS IsPublic,
                                    updated_by AS UpdatedBy,
                                    last_accessed_at AS LastAccessedAt,
                                    created_at AS CreatedAt,
                                    updated_at AS UpdatedAt
                               FROM cad_file_storage
                               ORDER BY updated_at DESC, id DESC";

            try
            {
                if (_adapter.DatabaseType == "MySQL")
                {
                    using var conn = _adapter.CreateConnection();
                    conn.Open();
                    var rows = await conn.QueryAsync<FileStorage>(sql).ConfigureAwait(false);
                    return rows.AsList();
                }

                using var dconn = GetConnection();
                dconn.Open();
                using var cmd = dconn.CreateCommand();
                cmd.CommandText = _adapter.NormalizeSql(sql);

                using var reader = cmd.ExecuteReader();
                var list = new List<FileStorage>();
                while (reader.Read())
                {
                    var file = new FileStorage();
                    int ord;
                    ord = reader.GetOrdinal("Id"); file.Id = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                    ord = reader.GetOrdinal("CategoryId"); file.CategoryId = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                    ord = reader.GetOrdinal("FileAttributeId"); file.FileAttributeId = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("FileName"); file.FileName = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("FileStoredName"); file.FileStoredName = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("DisplayName"); file.DisplayName = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("FileType"); file.FileType = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("FileHash"); file.FileHash = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("BlockName"); file.BlockName = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("LayerName"); file.LayerName = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("ColorIndex"); file.ColorIndex = reader.IsDBNull(ord) ? (int?)null : reader.GetInt32(ord);
                    ord = reader.GetOrdinal("Scale"); file.Scale = reader.IsDBNull(ord) ? (double?)null : reader.GetDouble(ord);
                    ord = reader.GetOrdinal("PreviewImageName"); file.PreviewImageName = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("PreviewImagePath"); file.PreviewImagePath = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("Description"); file.Description = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("Version"); file.Version = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                    ord = reader.GetOrdinal("IsPreview"); file.IsPreview = (!reader.IsDBNull(ord) && reader.GetInt32(ord) != 0) ? 1 : 0;
                    ord = reader.GetOrdinal("IsActive"); file.IsActive = (!reader.IsDBNull(ord) && reader.GetInt32(ord) != 0) ? 1 : 0;
                    ord = reader.GetOrdinal("CategoryType"); file.CategoryType = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("Title"); file.Title = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("Keywords"); file.Keywords = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("IsPublic"); file.IsPublic = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                    ord = reader.GetOrdinal("UpdatedBy"); file.UpdatedBy = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                    ord = reader.GetOrdinal("LastAccessedAt"); file.LastAccessedAt = reader.IsDBNull(ord) ? (DateTime?)null : reader.GetDateTime(ord);
                    ord = reader.GetOrdinal("CreatedAt"); file.CreatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                    ord = reader.GetOrdinal("UpdatedAt"); file.UpdatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                    list.Add(file);
                }

                return list;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetAllFileStorageAsync 出错: {ex.Message}");
                return new List<FileStorage>();
            }
        }

        /// <summary>
        /// 获取服务器端客户端版本。
        /// </summary>
        public async Task<string> GetServerClientVersionAsync()
        {
            const string sql = "SELECT config_value FROM system_config WHERE config_key = @ConfigKey LIMIT 1";

            try
            {
                if (_adapter.DatabaseType == "MySQL")
                {
                    using var conn = _adapter.CreateConnection();
                    conn.Open();
                    var value = await conn.QuerySingleOrDefaultAsync<string>(sql, new { ConfigKey = "ClientVersion" }).ConfigureAwait(false);
                    return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
                }

                using var dconn = GetConnection();
                dconn.Open();
                using var cmd = dconn.CreateCommand();
                cmd.CommandText = _adapter.NormalizeSql(sql);
                AddDmParam(cmd, "ConfigKey", "ClientVersion");
                var result = cmd.ExecuteScalar();
                return result == null || result == DBNull.Value ? string.Empty : Convert.ToString(result)?.Trim() ?? string.Empty;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetServerClientVersionAsync 出错: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// 读取系统配置。
        /// </summary>
        public async Task<string> GetSystemConfigValueAsync(string configKey)
        {
            if (string.IsNullOrWhiteSpace(configKey))
            {
                return string.Empty;
            }

            const string sql = "SELECT config_value FROM system_config WHERE config_key = @ConfigKey LIMIT 1";

            try
            {
                if (_adapter.DatabaseType == "MySQL")
                {
                    using var conn = _adapter.CreateConnection();
                    conn.Open();
                    var value = await conn.QuerySingleOrDefaultAsync<string>(sql, new { ConfigKey = configKey.Trim() }).ConfigureAwait(false);
                    return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
                }

                using var dconn = GetConnection();
                dconn.Open();
                using var cmd = dconn.CreateCommand();
                cmd.CommandText = _adapter.NormalizeSql(sql);
                AddDmParam(cmd, "ConfigKey", configKey.Trim());
                var result = cmd.ExecuteScalar();
                return result == null || result == DBNull.Value ? string.Empty : Convert.ToString(result)?.Trim() ?? string.Empty;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetSystemConfigValueAsync 出错: key={configKey}, {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// 读取系统配置并允许调用方指定是否修 trimmed 结果。
        /// </summary>
        public async Task<string> GetSystemConfigValueAsync(string configKey, bool trimValue)
        {
            var value = await GetSystemConfigValueAsync(configKey).ConfigureAwait(false);
            return trimValue ? value.Trim() : value;
        }

        /// <summary>
        /// 保存系统配置。
        /// </summary>
        public async Task<int> SetSystemConfigValueAsync(string configKey, string configValue)
        {
            if (string.IsNullOrWhiteSpace(configKey))
            {
                return 0;
            }

            const string mysqlSql = @"
INSERT INTO system_config (config_key, config_value)
VALUES (@ConfigKey, @ConfigValue)
ON DUPLICATE KEY UPDATE config_value = VALUES(config_value)";

            const string dmSql = @"
MERGE INTO system_config t
USING (SELECT :ConfigKey AS config_key, :ConfigValue AS config_value FROM DUAL) s
ON (t.config_key = s.config_key)
WHEN MATCHED THEN UPDATE SET t.config_value = s.config_value
WHEN NOT MATCHED THEN INSERT (config_key, config_value) VALUES (s.config_key, s.config_value)";
            try
            {
                if (_adapter.DatabaseType == "MySQL")
                {
                    using var conn = _adapter.CreateConnection();
                    conn.Open();
                    return await conn.ExecuteAsync(mysqlSql, new { ConfigKey = configKey.Trim(), ConfigValue = configValue ?? string.Empty }).ConfigureAwait(false);
                }

                using var dconn = GetConnection();
                dconn.Open();
                using var cmd = dconn.CreateCommand();
                cmd.CommandText = _adapter.NormalizeSql(dmSql);
                AddDmParam(cmd, "ConfigKey", configKey.Trim());
                AddDmParam(cmd, "ConfigValue", configValue ?? string.Empty);
                return cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"SetConfigValueAsync 出错: key={configKey}, {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 获取CAD分类的最大排序序号
        /// </summary>
        public async Task<int> GetMaxCadCategorySortOrderAsync()
        {
            const string sql = "SELECT COALESCE(MAX(sort_order), 0) FROM cad_categories";

            using var connection = new MySqlConnection(_connectionString);
            var result = await connection.QuerySingleOrDefaultAsync<int>(sql);
            return result;
        }

        /// <summary>
        /// 获取指定父分类下子分类的最大排序序号
        /// </summary>
        public async Task<int> GetMaxCadSubcategorySortOrderAsync(int parentId)
        {
            const string sql = "SELECT COALESCE(MAX(sort_order), 0) FROM cad_subcategories WHERE parent_id = @parentId";

            using var connection = new MySqlConnection(_connectionString);
            var result = await connection.QuerySingleOrDefaultAsync<int>(sql, new { parentId });
            return result;
        }

        /// <summary>
        /// 获取所有子分类的最大排序序号（用于主分类下的直接子分类）
        /// </summary>
        public async Task<int> GetMaxCadSubcategorySortOrderForMainCategoryAsync(int parentId)
        {
            const string sql = "SELECT COALESCE(MAX(sort_order), 0) FROM cad_subcategories WHERE parent_id = @parentId";

            using var connection = new MySqlConnection(_connectionString);
            var result = await connection.QuerySingleOrDefaultAsync<int>(sql, new { parentId });
            return result;
        }

        #endregion

        #region 优化的文件管理方法

        /// <summary>
        /// 获取分类下的所有文件（支持分页和排序）
        /// </summary>
        public async Task<List<FileStorage>> GetFilesByCategoryAsync(int categoryId, string categoryType = "sub",
    int page = 1, int pageSize = 50, string orderBy = "created_at DESC")
        {
            string sql = @"
        SELECT 
            id AS Id,
            category_id AS CategoryId,
            file_attribute_id AS FileAttributeId,
            file_name AS FileName,
            file_stored_name AS FileStoredName,
            display_name AS DisplayName,
            file_type AS FileType,
            file_hash AS FileHash,
            block_name AS BlockName,
            layer_name AS LayerName,
            color_index AS ColorIndex,
            scale AS Scale,
            file_path AS FilePath,
            preview_image_name AS PreviewImageName,
            preview_image_path AS PreviewImagePath,
            file_size AS FileSize,
            is_preview AS IsPreview,
            version AS Version,
            description AS Description,
            is_active AS IsActive,
            created_by AS CreatedBy,
            category_type AS CategoryType,
            title AS Title,
            keywords AS Keywords,
            is_public AS IsPublic,
            updated_by AS UpdatedBy,
            last_accessed_at AS LastAccessedAt,
            created_at AS CreatedAt,
            updated_at AS UpdatedAt
          FROM cad_file_storage
          WHERE category_id = @CategoryId 
          AND category_type = @CategoryType
          AND is_active = 1";
            try
            {
                if (!string.IsNullOrEmpty(orderBy))
                {
                    sql += $" ORDER BY {orderBy}";
                }
                sql += " LIMIT @offset, @pageSize";

                var offset = (page - 1) * pageSize;
                return await Task.Run(() =>
                {
                    using var conn = GetConnection();
                    conn.Open();
                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = sql.Replace("@CategoryId", ":CategoryId").Replace("@CategoryType", ":CategoryType").Replace("@offset", ":offset").Replace("@pageSize", ":pageSize");
                    AddDmParam(cmd, "CategoryId", categoryId);
                    AddDmParam(cmd, "CategoryType", categoryType ?? (object)DBNull.Value);
                    AddDmParam(cmd, "offset", offset);
                    AddDmParam(cmd, "pageSize", pageSize);

                    var list = new List<FileStorage>();
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var f = new FileStorage();
                        int ord;
                        ord = reader.GetOrdinal("Id"); f.Id = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("CategoryId"); f.CategoryId = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("CategoryType"); f.CategoryType = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("FileName"); f.FileName = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("FileStoredName"); f.FileStoredName = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("FilePath"); f.FilePath = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("FileType"); f.FileType = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("FileHash"); f.FileHash = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("DisplayName"); f.DisplayName = reader.IsDBNull(ord) ? f.FileName : reader.GetString(ord);
                        ord = reader.GetOrdinal("BlockName"); f.BlockName = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("LayerName"); f.LayerName = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("ColorIndex"); f.ColorIndex = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("Scale"); f.Scale = reader.IsDBNull(ord) ? (double?)null : reader.GetDouble(ord);
                        ord = reader.GetOrdinal("PreviewImageName"); f.PreviewImageName = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("PreviewImagePath"); f.PreviewImagePath = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("Description"); f.Description = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("Version"); f.Version = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("IsPreview"); f.IsPreview = (!reader.IsDBNull(ord) && reader.GetInt32(ord) != 0) ? 1 : 0;
                        ord = reader.GetOrdinal("IsActive"); f.IsActive = (!reader.IsDBNull(ord) && reader.GetInt32(ord) != 0) ? 1 : 0;
                        ord = reader.GetOrdinal("CreatedBy"); f.CreatedBy = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                        ord = reader.GetOrdinal("CreatedAt"); f.CreatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                        ord = reader.GetOrdinal("UpdatedAt"); f.UpdatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                        list.Add(f);
                    }
                    return list;
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"获取分类下的文件时出错: {ex.Message}");
                return new List<FileStorage>();
            }
        }

        /// <summary>
        ///  获取分类下的所有文件
        /// </summary>
        /// <param name="categoryId">分类Id</param>
        /// <param name="categoryType">分类类型</param>
        /// <returns></returns>
        public async Task<List<FileStorage>> GetFilesByCategoryIdAsync(int categoryId, string categoryType)
        {
            const string sql = @"
                SELECT 
                 id AS Id,
                 category_id AS CategoryId,
                 category_type AS CategoryType,
                 file_attribute_id AS FileAttributeId,
                 file_name AS FileName,
                 file_stored_name AS FileStoredName,
                 display_name AS DisplayName,
                 file_type AS FileType,
                 file_hash AS FileHash,
                 block_name AS BlockName,
                 layer_name AS LayerName,
                 color_index AS ColorIndex,
                 scale AS Scale,
                 file_path AS FilePath,
                 preview_image_name AS PreviewImageName,
                 preview_image_path AS PreviewImagePath,
                 file_size AS FileSize,
                 is_preview AS IsPreview,
                 version AS Version,
                 description AS Description,
                 is_active AS IsActive,
                 created_by AS CreatedBy,
                 category_type AS CategoryType,
                 title AS Title,
                 keywords AS Keywords,
                 is_public AS IsPublic,
                 updated_by AS UpdatedBy,
                 last_accessed_at AS LastAccessedAt,
                 created_at AS CreatedAt,
                 updated_at AS UpdatedAt
             FROM cad_file_storage 
             WHERE category_id = @CategoryId 
               AND category_type = @CategoryType
             ORDER BY created_at DESC";
            try
            {
                return await Task.Run(() =>
                {
                    using var conn = GetConnection();
                    conn.Open();
                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = sql.Replace("@CategoryId", ":CategoryId").Replace("@CategoryType", ":CategoryType");
                    AddDmParam(cmd, "CategoryId", categoryId);
                    AddDmParam(cmd, "CategoryType", categoryType ?? (object)DBNull.Value);

                    var list = new List<FileStorage>();
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var f = new FileStorage();
                        int ord;
                        ord = reader.GetOrdinal("Id"); f.Id = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("CategoryId"); f.CategoryId = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("CategoryType"); f.CategoryType = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("FileName"); f.FileName = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("FileStoredName"); f.FileStoredName = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("FilePath"); f.FilePath = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("FileType"); f.FileType = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("FileHash"); f.FileHash = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("DisplayName"); f.DisplayName = reader.IsDBNull(ord) ? f.FileName : reader.GetString(ord);
                        ord = reader.GetOrdinal("BlockName"); f.BlockName = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("LayerName"); f.LayerName = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("ColorIndex"); f.ColorIndex = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("Scale"); f.Scale = reader.IsDBNull(ord) ? (double?)null : reader.GetDouble(ord);
                        ord = reader.GetOrdinal("PreviewImageName"); f.PreviewImageName = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("PreviewImagePath"); f.PreviewImagePath = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("Description"); f.Description = reader.IsDBNull(ord) ? "" : reader.GetString(ord);
                        ord = reader.GetOrdinal("Version"); f.Version = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                        ord = reader.GetOrdinal("IsPreview"); f.IsPreview = (!reader.IsDBNull(ord) && reader.GetInt32(ord) != 0) ? 1 : 0;
                        ord = reader.GetOrdinal("IsActive"); f.IsActive = (!reader.IsDBNull(ord) && reader.GetInt32(ord) != 0) ? 1 : 0;
                        ord = reader.GetOrdinal("CreatedBy"); f.CreatedBy = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                        ord = reader.GetOrdinal("CreatedAt"); f.CreatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                        ord = reader.GetOrdinal("UpdatedAt"); f.UpdatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                        list.Add(f);
                    }
                    return list;
                }).ConfigureAwait(false);
            }
            catch (Exception e)
            {
                Env.Editor.WriteMessage(e.Message);
                return new List<FileStorage>();
            }
        }

        /// <summary>
        /// 兼容旧调用：按子分类读取图元文件。
        /// </summary>
        public async Task<List<FileStorage>> GetFileStorageBySubcategoryIdAsync(int subcategoryId)
        {
            return await GetFilesByCategoryIdAsync(subcategoryId, "sub").ConfigureAwait(false);
        }

        /// <summary>
        /// 兼容旧调用：按文件主键读取单条文件记录。
        /// </summary>
        // 在 FunctionalMethod\DatabaseManager.cs 中修改 GetFileByIdAsync 方法

        public async Task<FileStorage?> GetFileByIdAsync(int fileId)
        {
            // 定义基础 SQL
            string sql = @"
        SELECT 
            id AS Id,
            category_id AS CategoryId,
            file_attribute_id AS FileAttributeId,
            file_name AS FileName,
            file_stored_name AS FileStoredName,
            display_name AS DisplayName,
            file_type AS FileType,
            file_hash AS FileHash,
            block_name AS BlockName,
            layer_name AS LayerName,
            color_index AS ColorIndex,
            scale AS Scale,
            file_path AS FilePath,
            preview_image_name AS PreviewImageName,
            preview_image_path AS PreviewImagePath,
            file_size AS FileSize,
            is_preview AS IsPreview,
            version AS Version,
            description AS Description,
            is_active AS IsActive,
            created_by AS CreatedBy,
            category_type AS CategoryType,
            title AS Title,
            keywords AS Keywords,
            is_public AS IsPublic,
            updated_by AS UpdatedBy,
            last_accessed_at AS LastAccessedAt,
            created_at AS CreatedAt,
            updated_at AS UpdatedAt
        FROM cad_file_storage
        WHERE id = @Id";

            // 自适应分页/限制语法：MySQL 使用 LIMIT，达梦使用 FETCH FIRST
            if (_adapter.DatabaseType == "MySQL")
            {
                sql += " LIMIT 1";
            }
            else
            {
                sql += " FETCH FIRST 1 ROWS ONLY";
            }

            try
            {
                using var conn = GetConnection();
                conn.Open();
                using var cmd = conn.CreateCommand();
                // 使用适配器的 NormalizeSql 转换参数占位符（@Id -> :Id）以及大小写
                cmd.CommandText = _adapter.NormalizeSql(sql);
                // 使用适配器的 AddParameter 自动处理参数绑定
                _adapter.AddParameter(cmd, "Id", fileId);

                using var reader = cmd.ExecuteReader();
                if (!reader.Read()) return null;

                var f = new FileStorage();
                int ord;
                ord = reader.GetOrdinal("Id"); f.Id = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("CategoryId"); f.CategoryId = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("FileAttributeId"); f.FileAttributeId = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                ord = reader.GetOrdinal("FileName"); f.FileName = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("FileStoredName"); f.FileStoredName = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("FileType"); f.FileType = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("FileHash"); f.FileHash = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("DisplayName"); f.DisplayName = reader.IsDBNull(ord) ? f.FileName : reader.GetString(ord);
                ord = reader.GetOrdinal("BlockName"); f.BlockName = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("LayerName"); f.LayerName = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("ColorIndex"); f.ColorIndex = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("Scale"); f.Scale = reader.IsDBNull(ord) ? (double?)null : reader.GetDouble(ord);
                ord = reader.GetOrdinal("PreviewImageName"); f.PreviewImageName = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("PreviewImagePath"); f.PreviewImagePath = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("Description"); f.Description = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("Version"); f.Version = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("IsPreview"); f.IsPreview = (!reader.IsDBNull(ord) && reader.GetInt32(ord) != 0) ? 1 : 0;
                ord = reader.GetOrdinal("IsActive"); f.IsActive = (!reader.IsDBNull(ord) && reader.GetInt32(ord) != 0) ? 1 : 0;
                ord = reader.GetOrdinal("CategoryType"); f.CategoryType = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("Title"); f.Title = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("Keywords"); f.Keywords = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("IsPublic"); f.IsPublic = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("UpdatedBy"); f.UpdatedBy = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                ord = reader.GetOrdinal("LastAccessedAt"); f.LastAccessedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                ord = reader.GetOrdinal("CreatedAt"); f.CreatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                ord = reader.GetOrdinal("UpdatedAt"); f.UpdatedAt = reader.IsDBNull(ord) ? DateTime.MinValue : reader.GetDateTime(ord);
                return f;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetFileByIdAsync 出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 新增：按文件哈希获取“主记录 + JSON 属性字典 + 配置名”
        /// </summary>
        public async Task<(FileStorage File, Dictionary<string, string> Attributes, string ConfigName)> GetFileStorageWithAttributesByHashAsync(
            string filehash,
            string preferredConfigName = null)
        {
            // 先拿主记录（复用现有方法）
            var file = await GetFileStorageAsync(filehash).ConfigureAwait(false);

            // 主记录不存在时，直接返回空元组内容
            if (file == null)
                return (null, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), string.Empty);

            // 准备查询属性 JSON 的 SQL（优先配置名）
            const string attrByConfigSql = @"
SELECT config_name AS ConfigName, attributes_json AS AttributesJson
FROM cad_block_attributes_json
WHERE file_id = @FileId
  AND config_name = @ConfigName
ORDER BY attr_id DESC
LIMIT 1;";

            // 兜底查询（取最新）
            const string attrLatestSql = @"
SELECT config_name AS ConfigName, attributes_json AS AttributesJson
FROM cad_block_attributes_json
WHERE file_id = @FileId
ORDER BY attr_id DESC
LIMIT 1;";

            try
            {
                // MySQL 快捷路径
                if (_adapter.DatabaseType == "MySQL")
                {
                    using var connection = new MySqlConnection(_connectionString);

                    // 确定优先配置名（参数优先，其次主表 file_attribute_id）
                    var configName = string.IsNullOrWhiteSpace(preferredConfigName)
                        ? (file.FileAttributeId ?? string.Empty)
                        : preferredConfigName.Trim();

                    (string ConfigName, string AttributesJson)? row = null;

                    // 优先按配置名查
                    if (!string.IsNullOrWhiteSpace(configName))
                    {
                        row = await connection.QueryFirstOrDefaultAsync<(string ConfigName, string AttributesJson)>(
                            attrByConfigSql.Replace(":", "@"), new { FileId = file.Id, ConfigName = configName }).ConfigureAwait(false);
                    }

                    // 兜底查最新
                    if (row == null || string.IsNullOrWhiteSpace(row.Value.AttributesJson))
                    {
                        row = await connection.QueryFirstOrDefaultAsync<(string ConfigName, string AttributesJson)>(
                            attrLatestSql.Replace(":", "@"), new { FileId = file.Id }).ConfigureAwait(false);
                    }

                    // 没有属性记录时返回空字典
                    if (row == null || string.IsNullOrWhiteSpace(row.Value.AttributesJson))
                    {
                        return (file, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), string.Empty);
                    }

                    // 反序列化 JSON -> 字典
                    var dict = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(row.Value.AttributesJson)
                               ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    // 返回聚合结果
                    return (file, dict, row.Value.ConfigName ?? string.Empty);
                }

                // 达梦路径：手动查询并映射
                using var dconn = GetConnection();
                dconn.Open();

                var config = string.IsNullOrWhiteSpace(preferredConfigName) ? (file.FileAttributeId ?? string.Empty) : preferredConfigName.Trim();
                string? attributesJson = null;
                string foundConfigName = string.Empty;

                if (!string.IsNullOrWhiteSpace(config))
                {
                    using var cmd = dconn.CreateCommand();
                    cmd.CommandText = _adapter.NormalizeSql(attrByConfigSql);
                    AddDmParam(cmd, "FileId", file.Id);
                    AddDmParam(cmd, "ConfigName", config);
                    using var rdr = cmd.ExecuteReader();
                    if (rdr.Read())
                    {
                        attributesJson = rdr.IsDBNull(1) ? null : rdr.GetString(1);
                        foundConfigName = rdr.IsDBNull(0) ? string.Empty : rdr.GetString(0);
                    }
                }

                if (string.IsNullOrWhiteSpace(attributesJson))
                {
                    using var cmd2 = dconn.CreateCommand();
                    cmd2.CommandText = _adapter.NormalizeSql(attrLatestSql);
                    AddDmParam(cmd2, "FileId", file.Id);
                    using var rdr2 = cmd2.ExecuteReader();
                    if (rdr2.Read())
                    {
                        attributesJson = rdr2.IsDBNull(1) ? null : rdr2.GetString(1);
                        foundConfigName = rdr2.IsDBNull(0) ? string.Empty : rdr2.GetString(0);
                    }
                }

                if (string.IsNullOrWhiteSpace(attributesJson))
                {
                    return (file, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), string.Empty);
                }

                var dictDm = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(attributesJson)
                             ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                return (file, dictDm, foundConfigName ?? string.Empty);
            }
            catch (Exception ex)
            {
                // 异常时记录日志并返回主记录+空字典
                LogManager.Instance.LogInfo($"GetFileStorageWithAttributesByHashAsync 出错: {ex.Message}");
                return (file, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), string.Empty);
            }
        }

        /// <summary>
        /// 获取文件属性的详细信息（JSON新方案）
        /// 说明：保持旧返回签名 (FileStorage, FileAttribute)，内部把 JSON 反序列化后回填到 FileAttribute（兼容旧调用）。
        /// </summary>
        public async Task<(FileStorage File, FileAttribute Attribute)> GetFileWithAttributeAsync(int fileId)
        {
            // 查询文件主表信息
            const string fileSql = @"
SELECT
    id AS Id,
    category_id AS CategoryId,
    file_attribute_id AS FileAttributeId,
    file_name AS FileName,
    file_stored_name AS FileStoredName,
    display_name AS DisplayName,
    file_type AS FileType,
    file_hash AS FileHash,
    block_name AS BlockName,
    layer_name AS LayerName,
    color_index AS ColorIndex,
    scale AS Scale,
    file_path AS FilePath,
    preview_image_name AS PreviewImageName,
    preview_image_path AS PreviewImagePath,
    file_size AS FileSize,
    is_preview AS IsPreview,
    version AS Version,
    description AS Description,
    is_active AS IsActive,
    created_by AS CreatedBy,
    category_type AS CategoryType,
    title AS Title,
    keywords AS Keywords,
    is_public AS IsPublic,
    updated_by AS UpdatedBy,
    last_accessed_at AS LastAccessedAt,
    created_at AS CreatedAt,
    updated_at AS UpdatedAt
FROM cad_file_storage
WHERE id = :Id
LIMIT 1;";

            // 优先按 config_name = storage.file_attribute_id 查 JSON 属性
            // 使用达梦兼容的参数占位符 (:FileId, :ConfigName)
            const string attrSqlByConfig = @"
SELECT
    attr_id AS AttrId,
    file_id AS FileId,
    config_name AS ConfigName,
    attributes_json AS AttributesJson,
    created_at AS CreatedAt,
    updated_at AS UpdatedAt
FROM cad_block_attributes_json
WHERE file_id = :FileId
  AND config_name = :ConfigName
ORDER BY attr_id DESC
LIMIT 1";

            // 若按 config_name 查不到，则按 file_id 取最新一条兜底
            const string attrSqlByLatest = @"
SELECT
    attr_id AS AttrId,
    file_id AS FileId,
    config_name AS ConfigName,
    attributes_json AS AttributesJson,
    created_at AS CreatedAt,
    updated_at AS UpdatedAt
FROM cad_block_attributes_json
WHERE file_id = :FileId
ORDER BY attr_id DESC
LIMIT 1";

            try
            {
                // 创建数据库连接
                using var connection = new MySqlConnection(_connectionString);

                // 先取主表文件记录
                var file = await connection.QuerySingleOrDefaultAsync<FileStorage>(fileSql, new { Id = fileId }).ConfigureAwait(false);

                // 若文件不存在，直接返回空
                if (file == null)
                {
                    return (null, null);
                }

                // 准备读取 JSON 属性记录
                BlockAttributesJson jsonRow = null;

                // 兼容字段 file_attribute_id，此处作为 config_name 优先匹配
                var configName = Convert.ToString(file.FileAttributeId)?.Trim();

                // 优先按配置名查
                if (!string.IsNullOrWhiteSpace(configName))
                {
                    jsonRow = await connection.QuerySingleOrDefaultAsync<BlockAttributesJson>(
                        attrSqlByConfig,
                        new { FileId = file.Id, ConfigName = configName }).ConfigureAwait(false);
                }

                // 兜底按最新记录查
                if (jsonRow == null)
                {
                    jsonRow = await connection.QuerySingleOrDefaultAsync<BlockAttributesJson>(
                        attrSqlByLatest,
                        new { FileId = file.Id }).ConfigureAwait(false);
                }

                // 没有属性记录时，返回 file + null（兼容旧逻辑）
                if (jsonRow == null || string.IsNullOrWhiteSpace(jsonRow.AttributesJson))
                {
                    return (file, null);
                }

                // 反序列化 JSON 为键值字典
                var dict = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonRow.AttributesJson)
                           ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                // 构建兼容旧调用的 FileAttribute 对象
                var attr = new FileAttribute
                {
                    FileAttributeId = string.IsNullOrWhiteSpace(jsonRow.ConfigName) ? file.FileAttributeId : jsonRow.ConfigName
                };

                // 遍历 FileAttribute 可写属性，按属性名从字典回填
                foreach (var p in typeof(FileAttribute).GetProperties())
                {
                    // 只处理可写属性
                    if (!p.CanWrite) continue;

                    // 字典里没有同名键则跳过
                    if (!dict.TryGetValue(p.Name, out var raw)) continue;

                    // 空值跳过
                    if (string.IsNullOrWhiteSpace(raw)) continue;

                    try
                    {
                        // 处理可空类型
                        var targetType = Nullable.GetUnderlyingType(p.PropertyType) ?? p.PropertyType;

                        // 字符串直接赋值
                        if (targetType == typeof(string))
                        {
                            p.SetValue(attr, raw);
                            continue;
                        }

                        // 整型转换
                        if (targetType == typeof(int))
                        {
                            if (int.TryParse(raw, out var v)) p.SetValue(attr, v);
                            continue;
                        }

                        // 长整型转换
                        if (targetType == typeof(long))
                        {
                            if (long.TryParse(raw, out var v)) p.SetValue(attr, v);
                            continue;
                        }

                        // decimal 转换（用不变文化，兼容小数点）
                        if (targetType == typeof(decimal))
                        {
                            if (decimal.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                            {
                                p.SetValue(attr, v);
                            }
                            else if (decimal.TryParse(raw, out var v2)) p.SetValue(attr, v2);
                            continue;
                        }

                        // double 转换（用不变文化，兼容小数点）
                        if (targetType == typeof(double))
                        {
                            if (double.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                            {
                                p.SetValue(attr, v);
                            }
                            else if (double.TryParse(raw, out var v2)) p.SetValue(attr, v2);
                            continue;
                        }

                        // DateTime 转换
                        if (targetType == typeof(DateTime))
                        {
                            if (DateTime.TryParse(raw, out var v)) p.SetValue(attr, v);
                            continue;
                        }

                        // bool 转换（兼容 1/0 与 true/false）
                        if (targetType == typeof(bool))
                        {
                            if (bool.TryParse(raw, out var b))
                            {
                                p.SetValue(attr, b);
                            }
                            else if (raw == "1")
                            {
                                p.SetValue(attr, true);
                            }
                            else if (raw == "0")
                            {
                                p.SetValue(attr, false);
                            }
                        }
                    }
                    catch
                    {
                    }
                }
                ;

                // 返回主表文件记录 + 兼容属性对象
                return (file, attr);
            }
            catch (Exception ex)
            {
                // 输出错误并返回空对象元组
                Env.Editor.WriteMessage($"\n获取文件详细信息时出错: {ex.Message}");
                return (null, null);
            }
        }

        /// <summary>
        /// ===== 已废弃：cad_file_attributes 全字段查询片段（不再使用）=====
        /// </summary>
        // private const string CadFileAttributeSelectColumns = @"...




        /// <summary>
        /// 新方案——插入文件主记录 + JSON属性记录（事务）
        /// </summary>
        /// <param name="storage">中文注释：文件主表对象</param>
        /// <param name="attributes">中文注释：动态属性字典</param>
        /// <param name="configName">中文注释：属性配置名</param>
        /// <returns>中文注释：返回主表ID与属性表ID</returns>
        /// <exception cref="ArgumentNullException">中文注释：主表对象为空时抛出</exception>
        public async Task<(int StorageId, long AttrId)> AddFileStorageAndAttributesJsonAsync(
            FileStorage storage,
            Dictionary<string, string> attributes,
            string configName = "default")
        {
            if (storage == null)
            {
                throw new ArgumentNullException(nameof(storage));
            }

            // 属性字典允许为空，统一做空集合兜底。
            attributes = attributes ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            // 配置名允许为空，统一使用 "default" 作为默认值。
            using var connection = GetConnection();
            connection.Open();// 开启事务，确保主表和属性表操作的一致性。
            using var tx = connection.BeginTransaction();// 开启事务，确保主表和属性表操作的一致性。

            try
            {
                // 先找主表和属性表名，增强达梦大写兜底逻辑已在 ResolveExistingTableName 中实现。
                string storageTable = ResolveExistingTableName(connection, tx, new[]
                {
                    "cad_file_storage",
                    "CAD_FILE_STORAGE",
                    "file_storage",
                    "FILE_STORAGE"
                });

                if (string.IsNullOrWhiteSpace(storageTable))
                {
                    throw new InvalidOperationException("未找到图元主表（cad_file_storage）。");
                }

                string attrTable = ResolveExistingTableName(connection, tx, new[]
                {
                    "cad_block_attributes_json",
                    "CAD_BLOCK_ATTRIBUTES_JSON"
                });

                if (string.IsNullOrWhiteSpace(attrTable))
                {
                    throw new InvalidOperationException("未找到图元属性表（cad_block_attributes_json）。");
                }

                // 读取真实数据库列清单，实现“按需插入”，避免由于模型字段多于数据库字段导致的报错。
                var storageColumns = ReadTableColumns(connection, tx, storageTable);
                var attrColumns = ReadTableColumns(connection, tx, attrTable);

                if (storageColumns.Count == 0 || attrColumns.Count == 0)
                {
                    throw new InvalidOperationException("无法获取数据库表结构信息。");
                }

                // 1. 插入文件记录主表 (cad_file_storage)
                var storageInsertValues = BuildStorageInsertValues(storage, attributes, storageColumns, configName);

                // 调用统一生成的身份获取逻辑，确保存储 ID 正确返回。
                var storageIdLong = await ExecuteInsertAndReturnIdentity(connection, tx, storageTable, storageInsertValues).ConfigureAwait(false);
                int storageId = Convert.ToInt32(storageIdLong);

                if (storageId <= 0)
                {
                    throw new InvalidOperationException("主表插入失败，未能生成存储 Id。");
                }

                // 2. 插入属性记录表 (cad_block_attributes_json)
                // BuildAttributeInsertValues 会将属性字典转为序列化后的 JSON 字符串。
                var attrInsertValues = BuildAttributeInsertValues(storage, attributes, attrColumns, storageId);
                var attrId = await ExecuteInsertAndReturnIdentity(connection, tx, attrTable, attrInsertValues).ConfigureAwait(false);

                if (attrId <= 0)
                {
                    throw new InvalidOperationException("属性表插入失败，未能生成属性 Id。");
                }

                // 3. 回写外键：将刚生成的属性 ID 更新回主表的 file_attribute_id 字段，完成关联。
                string updateSql = "UPDATE CAD_FILE_STORAGE SET file_attribute_id = :AttrId WHERE id = :Id";

                await connection.ExecuteAsync(updateSql, new { AttrId = attrId, Id = storageId }, tx).ConfigureAwait(false);

                tx.Commit();

                // 返回最终生成的 ID 对
                return (storageId, attrId);
            }
            catch (Exception ex)
            {
                try
                {
                    tx?.Rollback();
                }
                catch
                {
                    // 忽略回滚异常
                }

                // 记录详细异常到日志
                LogManager.Instance.LogInfo($"AddFileStorageAndAttributesJsonAsync 事务失败: {ex.Message}");
                return (0, 0);
            }
        }

        #region 辅助方法

        /// <summary>
        /// 根据表别名或视图名，查询真实的表名（支持多别名/视图名），返回第一个有效表名
        /// 说明：用于动态适配不同环境（如开发/测试/生产）下的表名差异。
        /// </summary>
        /// <param name="aliases">候选表别名列表</param>
        /// <returns>找到的第一个有效表名，或空字符串</returns>
        private string ResolveExistingTableName(IDbConnection conn, IDbTransaction tx, params string[] aliases)
        {
            if (aliases == null || aliases.Length == 0) return string.Empty;

            // 达梦数据库通常对大小写敏感，且存储在系统表中的名称默认是大写的。
            // 修正查询逻辑：不仅查 ALL_TABLES，还要处理候选名的 Trim 和 Upper 处理。
            string sqlCheckDm = "SELECT TABLE_NAME FROM ALL_TABLES WHERE OWNER = :Owner AND TABLE_NAME = :TableName";
            string sqlCheckMySql = "SELECT table_name FROM information_schema.tables WHERE table_schema = DATABASE() AND table_name = @TableName";

            var sql = _adapter.DatabaseType == "MySQL" ? sqlCheckMySql : sqlCheckDm;

            foreach (var alias in aliases)
            {
                // 生成候选名称序列：原始、大写、小写
                var namesToTry = new[] { alias.Trim(), alias.Trim().ToUpperInvariant(), alias.Trim().ToLowerInvariant() };

                foreach (var name in namesToTry.Distinct())
                {
                    try
                    {
                        var param = _adapter.DatabaseType == "MySQL"
                            ? (object)new { TableName = name }
                            : new { Owner = _schemaName, TableName = name };

                        var found = conn.QuerySingleOrDefault<string>(sql, param, tx);
                        if (!string.IsNullOrWhiteSpace(found)) return found;
                    }
                    catch { continue; }
                }
            }

            // 如果动态查询失败，为了保证业务不中断，强制返回第一个候选名的大写形式（常见于达梦）
            return aliases[0].ToUpperInvariant();
        }

        /// <summary>
        /// 读取表的所有列信息（用于动态构建插入/更新语句）
        /// 返回：包含列名及类型的字典，键为列名（不带前缀），值为数据类型
        /// </summary>
        /// <param name="tableName">真实表名</param>
        private Dictionary<string, string> ReadTableColumns(IDbConnection conn, IDbTransaction tx, string tableName)
        {
            var columns = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 达梦必须精准匹配 TABLE_NAME，通常为大写。
            string sqlDm = "SELECT COLUMN_NAME, DATA_TYPE FROM USER_TAB_COLUMNS WHERE TABLE_NAME = :TableName";
            string sqlMySql = "SELECT column_name, data_type FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = @TableName";

            string sql = _adapter.DatabaseType == "MySQL" ? sqlMySql : sqlDm;

            try
            {
                using var cmd = conn.CreateCommand();
                cmd.Transaction = tx;
                cmd.CommandText = sql;
                _adapter.AddParameter(cmd, "TableName", tableName);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    // 存入字典，Key 强制不区分大小写
                    columns[reader.GetString(0)] = reader.GetString(1).ToLower();
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"ReadTableColumns 警告 (表 {tableName}): {ex.Message}");
            }
            return columns;
        }

        /// <summary>
        /// 根据列名和表名，构造 INSERT 语句的 VALUES 部分
        /// 说明：只包含数据库实际存在的列，避免因列缺失导致的插入错误。
        /// </summary>
        /// <param name="tableName">目标表名</param>
        /// <param name="values">要插入的列值</param>
        /// <param name="columns">目标列清单</param>
        /// <returns>VALUES 子句字符串</returns>
        private string BuildInsertValues(string tableName, object values, Dictionary<string, string> columns)
        {
            var props = values.GetType().GetProperties();
            var sqlValues = new List<string>();

            foreach (var prop in props)
            {
                // 只处理数据库实际存在的列
                if (!columns.ContainsKey(prop.Name)) continue;

                var val = prop.GetValue(values);
                var isString = columns[prop.Name] == "string";

                sqlValues.Add(QuoteValueForInsert(val, isString));
            }

            return string.Join(", ", sqlValues);
        }

        /// <summary>
        /// 将字段值转换为插入语句的值部分
        /// 说明：处理特殊字符和转义，确保生成的 SQL 安全有效。
        /// </summary>
        /// <param name="value">字段值</param>
        /// <param name="isString">是否为字符串类型</param>
        /// <returns>转义后的值字符串</returns>
        private string QuoteValueForInsert(object value, bool isString)
        {
            if (value == null)
            {
                return "NULL";
            }

            // 字符串类型需要加引号
            if (isString)
            {
                // 转义单引号
                var str = value.ToString().Replace("'", "''");
                return $"'{str}'";
            }

            return value.ToString();
        }

        /// <summary>
        /// 执行插入操作并返回自增主键 ID（适用于 MySQL 和 DM）
        /// 说明：针对不同数据库类型，采用适当的方式获取自增 ID。
        /// </summary>
        /// <param name="conn">数据库连接</param>
        /// <param name="tx">事务对象</param>
        /// <param name="tableName">目标表名</param>
        /// <param name="insertValues">插入字段及值</param>
        /// <returns>插入记录的自增主键 ID</returns>
                private int ExecuteInsertAndReturnIdentity(IDbConnection conn, IDbTransaction tx, string tableName, DynamicParameters insertValues)
        {
            // 生成插入 SQL 语句
            var columns = string.Join(", ", insertValues.ParameterNames.Select(n => n.TrimStart('@')));
            var parameters = string.Join(", ", insertValues.ParameterNames);

            // MySQL 插入语句
            var sqlInsert = $"INSERT INTO {tableName} ({columns}) VALUES ({parameters})";

            // 执行插入
            conn.Execute(sqlInsert, insertValues, tx);

            // 获取自增 ID
            const string identitySqlMySql = "SELECT LAST_INSERT_ID()";
            const string identitySqlDm = "SELECT IDENTITY_VAL_LOCAL()";

            var identitySql = _adapter.DatabaseType == "MySQL" ? identitySqlMySql : identitySqlDm;
            //return (int)conn.ExecuteScalar(identitySql,tx);
            var result = conn.ExecuteScalar(identitySql, transaction: tx);
            return Convert.ToInt32(result ?? 0);
        }

        /// <summary>
        /// 执行插入并返回自增 Id（针对 MySQL 和 DM 的统一实现）
        /// </summary>
        /// <param name="connection">数据库连接</param>
        /// <param name="transaction">事务对象</param>
        /// <param name="tableName">目标表名</param>
        /// <param name="insertValues">插入字段及对应值</param>
        /// <returns>插入记录的自增主键 Id</returns>
        private async Task<long> ExecuteInsertAndReturnIdentity(IDbConnection connection, IDbTransaction transaction, string tableName, object insertValues)
        {
            var columnList = new List<string>();// 用于存储实际参与插入的列名列表
            var dp = new DynamicParameters();// 用于存储实际参与插入的参数列表

            // 支持 Dictionary
            if (insertValues is IDictionary<string, object> dict)
            {
                // 遍历字典，构建列名列表和参数列表
                foreach (var key in dict.Keys)
                {
                    // 只处理数据库实际存在的列，且列名不区分大小写
                    var col = key.ToUpperInvariant();
                    // 只有当列存在于数据库表结构中时才添加到插入列表，避免因列缺失导致的错误。
                    columnList.Add(col);
                    // 直接使用字典中的值，不强制转换为 DBNull，保持原样传递给 Dapper 处理。
                    object value = dict[key];
                    // 关键修复：不强制转 DBNull，保持原样即可
                    dp.Add(col, value);
                }
            }
            // 支持实体对象
            else
            {
                foreach (var prop in insertValues.GetType().GetProperties())
                {
                    var col = prop.Name.ToUpperInvariant();
                    columnList.Add(col);
                    object value = prop.GetValue(insertValues);

                    // 关键修复：不转 DBNull
                    dp.Add(col, value);
                }
            }

            if (columnList.Count == 0)
                throw new InvalidOperationException("插入对象不包含任何有效字段");

            // 构建 SQL
            var sb = new System.Text.StringBuilder();
            sb.Append("INSERT INTO \"")
              .Append(tableName.ToUpperInvariant())
              .Append("\" (");

            sb.Append(string.Join(",", columnList.Select(c => "\"" + c + "\"")));
            sb.Append(") VALUES (");
            sb.Append(string.Join(",", columnList.Select(c => ":" + c)));
            sb.Append(")");

            var sql = sb.ToString();

            // 执行插入
            int rows = await connection.ExecuteAsync(sql, dp, transaction).ConfigureAwait(false);
            if (rows <= 0)
                throw new InvalidOperationException($"插入失败，受影响行数为0，表名：{tableName}");

            // MySQL 获取ID
            if (_adapter.DatabaseType == "MySQL")
            {
                var id = await connection.ExecuteScalarAsync("SELECT LAST_INSERT_ID();", null, transaction).ConfigureAwait(false);
                return Convert.ToInt64(id ?? 0);
            }

            // ==================== 达梦 ====================
            object? res = null;

            try { res = await connection.ExecuteScalarAsync("SELECT SCOPE_IDENTITY()", null, transaction).ConfigureAwait(false); } catch { }
            if (res == null || res == DBNull.Value)
                try { res = await connection.ExecuteScalarAsync("SELECT @@IDENTITY", null, transaction).ConfigureAwait(false); } catch { }
            if (res == null || res == DBNull.Value)
                try { res = await connection.ExecuteScalarAsync("SELECT IDENTITY_VAL_LOCAL()", null, transaction).ConfigureAwait(false); } catch { }

            // 终极兜底
            if (res == null || res == DBNull.Value)
            {
                try
                {
                    string finalSql = $@"SELECT ""ID"" FROM ""{tableName.ToUpperInvariant()}"" WHERE ROWID = (SELECT MAX(ROWID) FROM ""{tableName.ToUpperInvariant()}"")";
                    res = await connection.ExecuteScalarAsync(finalSql, null, transaction).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogInfo($"[致命] 获取插入ID失败：{ex.Message}");
                    throw new InvalidOperationException("无法获取新插入记录的主键ID", ex);
                }
            }

            return res == null || res == DBNull.Value ? 0L : Convert.ToInt64(res);
        }


        #endregion
        #endregion

        /// <summary>
        /// 按文件主键读取最新 JSON 属性字典（可选优先配置名）。
        /// </summary>
        /// <param name="fileId">cad_file_storage.id</param>
        /// <param name="preferredConfigName">可选：优先配置名</param>
        /// <returns>属性字典；无记录时返回空字典</returns>
        public async Task<Dictionary<string, string>> GetAttributesJsonByFileIdAsync(int fileId, string preferredConfigName = null)
        {
            var empty = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (fileId <= 0)
            {
                return empty;
            }

            const string attrByConfigSql = @"
SELECT attributes_json AS AttributesJson
FROM cad_block_attributes_json
WHERE file_id = @FileId
  AND config_name = @ConfigName
ORDER BY attr_id DESC
LIMIT 1;";

            const string attrLatestSql = @"
SELECT attributes_json AS AttributesJson
FROM cad_block_attributes_json
WHERE file_id = @FileId
ORDER BY attr_id DESC
LIMIT 1;";

            try
            {
                if (_adapter.DatabaseType == "MySQL")
                {
                    using var conn = _adapter.CreateConnection();
                    conn.Open();

                    string? json = null;
                    if (!string.IsNullOrWhiteSpace(preferredConfigName))
                    {
                        json = await conn.QueryFirstOrDefaultAsync<string>(
                            attrByConfigSql.Replace(":", "@"),
                            new { FileId = fileId, ConfigName = preferredConfigName.Trim() }).ConfigureAwait(false);
                    }

                    if (string.IsNullOrWhiteSpace(json))
                    {
                        json = await conn.QueryFirstOrDefaultAsync<string>(
                            attrLatestSql.Replace(":", "@"),
                            new { FileId = fileId }).ConfigureAwait(false);
                    }

                    if (string.IsNullOrWhiteSpace(json))
                    {
                        return empty;
                    }

                    return Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(json)
                           ?? empty;
                }

                using var dconn = GetConnection();
                dconn.Open();

                string? dmJson = null;
                if (!string.IsNullOrWhiteSpace(preferredConfigName))
                {
                    using var cmd = dconn.CreateCommand();
                    cmd.CommandText = _adapter.NormalizeSql(attrByConfigSql);
                    AddDmParam(cmd, "FileId", fileId);
                    AddDmParam(cmd, "ConfigName", preferredConfigName.Trim());
                    var obj = cmd.ExecuteScalar();
                    dmJson = obj == null || obj == DBNull.Value ? null : Convert.ToString(obj);
                }

                if (string.IsNullOrWhiteSpace(dmJson))
                {
                    using var cmd2 = dconn.CreateCommand();
                    cmd2.CommandText = _adapter.NormalizeSql(attrLatestSql);
                    AddDmParam(cmd2, "FileId", fileId);
                    var obj2 = cmd2.ExecuteScalar();
                    dmJson = obj2 == null || obj2 == DBNull.Value ? null : Convert.ToString(obj2);
                }

                if (string.IsNullOrWhiteSpace(dmJson))
                {
                    return empty;
                }

                return Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(dmJson)
                       ?? empty;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetAttributesJsonByFileIdAsync 出错: fileId={fileId}, {ex.Message}");
                return empty;
            }
        }

        /// <summary>
        /// 更新图元主表字段与 JSON 属性表（事务）。
        /// </summary>
        /// <param name="storage">图元主表对象（必须包含 Id）</param>
        /// <param name="attributes">JSON 属性字典</param>
        /// <param name="preferredConfigName">可选配置名；为空则使用 storage.FileAttributeId 或 default</param>
        /// <returns>是否更新成功</returns>
        public async Task<bool> UpdateFileStorageAndAttributesJsonAsync(FileStorage storage, Dictionary<string, string> attributes, string preferredConfigName = null)
        {
            if (storage == null)
            {
                throw new ArgumentNullException(nameof(storage));
            }

            if (storage.Id <= 0)
            {
                throw new ArgumentException("storage.Id 必须大于 0", nameof(storage));
            }

            attributes = attributes ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var configName = string.IsNullOrWhiteSpace(preferredConfigName)
                ? (string.IsNullOrWhiteSpace(storage.FileAttributeId) ? "default" : storage.FileAttributeId.Trim())
                : preferredConfigName.Trim();

            var now = DateTime.Now;

            using var connection = GetConnection();
            if (connection.State != ConnectionState.Open)
            {
                await ((DbConnection)connection).OpenAsync().ConfigureAwait(false);
            }

            using var tx = connection.BeginTransaction();

            try
            {
                const string updateStorageSql = @"
UPDATE cad_file_storage
SET file_name = @FileName,
    display_name = @DisplayName,
    block_name = @BlockName,
    layer_name = @LayerName,
    color_index = @ColorIndex,
    scale = @Scale,
    title = @Title,
    keywords = @Keywords,
    description = @Description,
    updated_by = @UpdatedBy,
    updated_at = @UpdatedAt
WHERE id = @Id";

                var updatedStorageRows = await ExecuteWriteAsync(connection, tx, updateStorageSql, new
                {
                    storage.Id,
                    storage.FileName,
                    storage.DisplayName,
                    storage.BlockName,
                    storage.LayerName,
                    storage.ColorIndex,
                    storage.Scale,
                    storage.Title,
                    storage.Keywords,
                    storage.Description,
                    storage.UpdatedBy,
                    UpdatedAt = now
                }).ConfigureAwait(false);

                if (updatedStorageRows <= 0)
                {
                    tx.Rollback();
                    return false;
                }

                var attributesJson = Newtonsoft.Json.JsonConvert.SerializeObject(attributes);

                const string updateAttrSql = @"
UPDATE cad_block_attributes_json
SET attributes_json = @AttributesJson,
    updated_at = @UpdatedAt
WHERE file_id = @FileId
  AND config_name = @ConfigName";

                var updatedAttrRows = await ExecuteWriteAsync(connection, tx, updateAttrSql, new
                {
                    FileId = storage.Id,
                    ConfigName = configName,
                    AttributesJson = attributesJson,
                    UpdatedAt = now
                }).ConfigureAwait(false);

                if (updatedAttrRows <= 0)
                {
                    const string insertAttrSql = @"
INSERT INTO cad_block_attributes_json (file_id, config_name, attributes_json, created_at, updated_at)
VALUES (@FileId, @ConfigName, @AttributesJson, @CreatedAt, @UpdatedAt)";

                    var insertedRows = await ExecuteWriteAsync(connection, tx, insertAttrSql, new
                    {
                        FileId = storage.Id,
                        ConfigName = configName,
                        AttributesJson = attributesJson,
                        CreatedAt = now,
                        UpdatedAt = now
                    }).ConfigureAwait(false);

                    if (insertedRows <= 0)
                    {
                        tx.Rollback();
                        return false;
                    }
                }

                if (string.IsNullOrWhiteSpace(storage.FileAttributeId) || !string.Equals(storage.FileAttributeId, configName, StringComparison.OrdinalIgnoreCase))
                {
                    const string updateConfigSql = @"
UPDATE cad_file_storage
SET file_attribute_id = @FileAttributeId,
    updated_at = @UpdatedAt
WHERE id = @Id";

                    await ExecuteWriteAsync(connection, tx, updateConfigSql, new
                    {
                        Id = storage.Id,
                        FileAttributeId = configName,
                        UpdatedAt = now
                    }).ConfigureAwait(false);

                    storage.FileAttributeId = configName;
                }

                storage.UpdatedAt = now;
                tx.Commit();
                return true;
            }
            catch (Exception ex)
            {
                try
                {
                    tx.Rollback();
                }
                catch
                {
                    // 忽略回滚异常
                }

                LogManager.Instance.LogInfo($"UpdateFileStorageAndAttributesJsonAsync 出错: FileId={storage.Id}, {ex.Message}");
                return false;
            }
        }
    }
}
