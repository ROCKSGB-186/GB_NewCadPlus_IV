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
                // 必须从参数对象本身读取属性值，不能把 PropertyInfo 自己传给 GetValue。
                // 这一步对达梦原生命令尤其重要，因为达梦不会像 Dapper 一样替我们反射绑定对象。
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
                ?? VariableDictionary._serverIP
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

        /// <summary>
        /// 删除文件存储记录（最小可编译实现，实际应包含完整逻辑）
        /// </summary>
        /// <param name="storageId"></param>
        /// <returns></returns>
        //public virtual async Task<bool> DeleteFileStorageAsync(long storageId)
        //{
        //    await Task.Yield();
        //    return true;
        //}


        /// <summary>
        /// 上传或更新文件存储记录（最小可编译实现，实际应包含完整逻辑）
        /// </summary>
        /// <param name="storage"></param>
        /// <returns></returns>
        public virtual async Task<bool> UpdateFileStorageAsync(FileStorage storage)
        {
            await Task.Yield();
            return true;
        }

        /// <summary>
        /// 补齐：根据 Hash 获取文件存储记录 (最小可编译实现)
        /// </summary>
        public async Task<FileStorage> GetFileStorageAsync(string fileHash)
        {
            if (string.IsNullOrWhiteSpace(fileHash)) return null;

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
                  description AS Description,
                  version AS Version,
                  is_preview AS IsPreview,
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
              WHERE file_hash = :FileHash
              FETCH FIRST 1 ROWS ONLY";

            try
            {
                using var conn = GetConnection();
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = _adapter.DatabaseType == "MySQL"
                    ? sql.Replace(":FileHash", "@FileHash").Replace("FETCH FIRST 1 ROWS ONLY", "LIMIT 1")
                    : _adapter.NormalizeSql(sql);
                AddParam(cmd, "FileHash", fileHash);

                using var reader = cmd.ExecuteReader();
                if (!reader.Read()) return null;

                var f = new FileStorage();
                // ... 列映射省略，参照您已有的 GetFileByIdAsync 写法 ...
                int ord;
                ord = reader.GetOrdinal("Id"); f.Id = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("CategoryId"); f.CategoryId = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("FileAttributeId"); f.FileAttributeId = reader.IsDBNull(ord) ? null : reader.GetString(ord);
                ord = reader.GetOrdinal("FileName"); f.FileName = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("FileStoredName"); f.FileStoredName = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("DisplayName"); f.DisplayName = reader.IsDBNull(ord) ? f.FileName : reader.GetString(ord);
                ord = reader.GetOrdinal("FileType"); f.FileType = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("FileHash"); f.FileHash = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("BlockName"); f.BlockName = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("LayerName"); f.LayerName = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("ColorIndex"); f.ColorIndex = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("Scale"); f.Scale = reader.IsDBNull(ord) ? (double?)null : reader.GetDouble(ord);
                ord = reader.GetOrdinal("FilePath"); f.FilePath = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("PreviewImageName"); f.PreviewImageName = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("PreviewImagePath"); f.PreviewImagePath = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("Description"); f.Description = reader.IsDBNull(ord) ? string.Empty : reader.GetString(ord);
                ord = reader.GetOrdinal("Version"); f.Version = reader.IsDBNull(ord) ? 0 : reader.GetInt32(ord);
                ord = reader.GetOrdinal("IsPreview"); f.IsPreview = (!reader.IsDBNull(ord) && reader.GetInt32(ord) != 0) ? 1 : 0;
                ord = reader.GetOrdinal("IsActive"); f.IsActive = (!reader.IsDBNull(ord) && reader.GetInt32(ord) != 0) ? 1 : 0;
                ord = reader.GetOrdinal("CreatedBy"); f.CreatedBy = reader.IsDBNull(ord) ? null : reader.GetString(ord);
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
                LogManager.Instance.LogError($"GetFileStorageAsync 出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 获取所有有效的图元记录（is_active = 1）
        /// </summary>
        public async Task<List<FileStorage>> GetAllActiveFileStoragesAsync()
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
        WHERE is_active = 1
        ORDER BY updated_at DESC, id DESC";

            try
            {
                // ---- MySQL 分支：同步 Open，异步查询 ----
                if (_adapter.DatabaseType == "MySQL")
                {
                    using var conn = _adapter.CreateConnection();
                    conn.Open();                       // 同步打开，IDbConnection 支持
                    var rows = await conn.QueryAsync<FileStorage>(sql).ConfigureAwait(false);
                    return rows.AsList();
                }

                // ---- 达梦 (DM) 分支：同步 Open，同步读取，整体包裹在 Task.Run 中保持异步 ----
                return await Task.Run(() =>
                {
                    var list = new List<FileStorage>();
                    using var dconn = GetConnection();
                    dconn.Open();
                    using var cmd = dconn.CreateCommand();
                    cmd.CommandText = _adapter.NormalizeSql(sql);

                    using var reader = cmd.ExecuteReader();
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
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetAllActiveFileStoragesAsync 出错: {ex.Message}");
                return new List<FileStorage>();
            }
        }

        /// <summary>
        /// 补齐：构建属性表插入值字典
        /// </summary>
        /// <summary>
        /// 完善：构建存储主表插入值字典
        /// </summary>
        //private Dictionary<string, object> BuildStorageInsertValues(FileStorage storage, Dictionary<string, string> attributes, Dictionary<string, string> columns, string configName)
        //{
        //    // 使用 StringComparer.OrdinalIgnoreCase 强制字典忽略大小写，防止出现 created_at 和 CREATED_AT 同时存在
        //    var values = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        //    var props = typeof(FileStorage).GetProperties();

        //    foreach (var p in props)
        //    {
        //        // 查找数据库中真实存在的列名
        //        var match = columns.Keys.FirstOrDefault(k =>
        //            string.Equals(k, p.Name, StringComparison.OrdinalIgnoreCase) ||
        //            string.Equals(k.Replace("_", ""), p.Name, StringComparison.OrdinalIgnoreCase));

        //        if (match != null)
        //        {
        //            // 核心修复：绝对排除 ID 列，达梦不允许显式插入 ID 值（自增列规则）
        //            if (string.Equals(match, "ID", StringComparison.OrdinalIgnoreCase)) continue;

        //            values[match] = p.GetValue(storage);
        //        }
        //    }

        //    // 审计字段逻辑：仅在反射未处理且数据库存在该列时添加，确保不会重复生成 SQL 列名
        //    if (columns.ContainsKey("created_at") && !values.ContainsKey("created_at"))
        //        values["created_at"] = storage.CreatedAt == default ? DateTime.Now : storage.CreatedAt;

        //    if (columns.ContainsKey("updated_at") && !values.ContainsKey("updated_at"))
        //        values["updated_at"] = DateTime.Now;

        //    return values;
        //}

        /// <summary>
        /// 完善：构建属性 JSON 表插入值字典
        /// </summary>
        //private Dictionary<string, object> BuildAttributeInsertValues(FileStorage storage, Dictionary<string, string> attributes, Dictionary<string, string> columns, int storageId)
        //{
        //    var values = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        //    if (columns.ContainsKey("file_id")) values["file_id"] = storageId;
        //    if (columns.ContainsKey("config_name")) values["config_name"] = "default";

        //    // 显式序列化 JSON 字典
        //    if (columns.ContainsKey("attributes_json"))
        //    {
        //        values["attributes_json"] = Newtonsoft.Json.JsonConvert.SerializeObject(attributes);
        //    }

        //    if (columns.ContainsKey("created_at") && !values.ContainsKey("created_at"))
        //        values["created_at"] = DateTime.Now;

        //    // 再次加固：确保物理上不含任何 ID 键
        //    values.Remove("ID");

        //    return values;
        //}

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
                connection.Open();
                using var cmd = connection.CreateCommand();

                string sql = physicalDelete
                    ? "DELETE FROM cad_file_storage WHERE id = :Id"
                    : "UPDATE cad_file_storage SET is_active = 0 WHERE id = :Id";

                // MySQL 参数占位符适配
                if (_adapter.DatabaseType == "MySQL")
                    sql = sql.Replace(":Id", "@Id");

                cmd.CommandText = _adapter.NormalizeSql(sql);
                AddParam(cmd, "Id", id);   // 使用类型安全的 AddParam 方法

                int rows = cmd.ExecuteNonQuery();
                LogManager.Instance.LogInfo($"DeleteCadGraphicCascadeAsync: id={id}, rows={rows}, physical={physicalDelete}");
                return rows > 0;
            }
            catch (Exception ex)
            {
                // 更详细的错误日志
                LogManager.Instance.LogError($"DeleteCadGraphicCascadeAsync 出错: id={id}, type={_adapter.DatabaseType}, {ex.Message}");
                return false;
            }
        }
        /// <summary>
        /// 获取配置值（最小可编译实现，实际应包含完整逻辑）
        /// </summary>
        /// <param name="args">  </param>
        /// <returns></returns>
        public async Task<dynamic> GetConfigValueAsync(params object[] args) { return await Task.FromResult<dynamic>(null); }
        /// <summary>
        /// 更新分类统计数据（最小可编译实现，实际应包含完整逻辑）
        /// </summary>
        /// <param name="args"> </param>
        /// <returns></returns>
        //public async Task<dynamic> UpdateCategoryStatisticsAsync(params object[] args) { return await Task.FromResult<dynamic>(false); }
        /// <summary>
        /// 添加 CAD 子分类。
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        public async Task<int> AddCadSubcategoryAsync(params object[] args)
        {
            // CategoryManager 传入的是一个 CadSubcategory 对象，保留 params 形式兼容旧调用方。
            var subcategory = args?.FirstOrDefault() as CadSubcategory;
            if (subcategory == null || subcategory.Id <= 0 || subcategory.ParentId <= 0 || string.IsNullOrWhiteSpace(subcategory.Name))
            {
                return 0;
            }

            try
            {
                // 使用统一连接和写入方法，自动适配 MySQL 与达梦参数格式。
                using var connection = GetConnection();
                const string sql = @"
                    INSERT INTO cad_subcategories
                        (id, parent_id, name, display_name, sort_order, level, subcategory_ids)
                    VALUES
                        (@Id, @ParentId, @Name, @DisplayName, @SortOrder, @Level, @SubcategoryIds)";

                // 保存前统一处理字符串，避免数据库中出现无意义的首尾空格。
                subcategory.Name = subcategory.Name.Trim();
                subcategory.DisplayName = string.IsNullOrWhiteSpace(subcategory.DisplayName)
                    ? subcategory.Name
                    : subcategory.DisplayName.Trim();
                subcategory.SubcategoryIds = subcategory.SubcategoryIds ?? string.Empty;

                return await ExecuteWriteAsync(connection, null, sql, subcategory).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // 记录数据库类型，便于区分 MySQL 和达梦的表结构/主键问题。
                LogManager.Instance.LogInfo($"AddCadSubcategoryAsync({_adapter.DatabaseType}) 出错: {ex.Message}");
                return 0;
            }
        }
        /// <summary>
        /// 更新父分类的子分类列表。
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        public async Task<int> UpdateParentSubcategoryListAsync(params object[] args)
        {
            // 新增调用传入 parentId + 子分类ID；删除调用传入 parentId + 完整ID字符串。
            if (args == null || args.Length < 2 || !TryConvertToInt(args[0], out var parentId) || parentId <= 0)
            {
                return 0;
            }

            try
            {
                string subcategoryIds;
                if (args[1] is string suppliedIds)
                {
                    // 删除场景：调用方已经计算出移除后的完整列表。
                    subcategoryIds = suppliedIds ?? string.Empty;
                }
                else if (TryConvertToInt(args[1], out var childId))
                {
                    // 新增场景：先读出旧列表，再追加新 ID，且不重复追加。
                    var currentIds = await GetParentSubcategoryIdsAsync(parentId).ConfigureAwait(false);
                    var ids = ParseSubcategoryIds(currentIds);
                    if (!ids.Contains(childId.ToString()))
                    {
                        ids.Add(childId.ToString());
                    }
                    subcategoryIds = string.Join(",", ids);
                }
                else
                {
                    return 0;
                }

                // 主分类 ID 小于 10000，子分类父级 ID 从 10000 开始，这是现有项目约定。
                const string categorySql = "UPDATE cad_categories SET subcategory_ids = @SubcategoryIds WHERE id = @ParentId";
                const string subcategorySql = "UPDATE cad_subcategories SET subcategory_ids = @SubcategoryIds WHERE id = @ParentId";
                var sql = parentId >= 10000 ? subcategorySql : categorySql;

                using var connection = GetConnection();
                return await ExecuteWriteAsync(connection, null, sql, new { ParentId = parentId, SubcategoryIds = subcategoryIds }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"UpdateParentSubcategoryListAsync({_adapter.DatabaseType}) 出错: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 读取父级当前保存的子分类 ID 列表。
        /// </summary>
        private async Task<string> GetParentSubcategoryIdsAsync(int parentId)
        {
            if (parentId >= 10000)
            {
                var parentSubcategory = await GetCadSubcategoryByIdAsync(parentId).ConfigureAwait(false);
                return parentSubcategory?.SubcategoryIds ?? string.Empty;
            }

            var categories = await GetAllCadCategoriesAsync().ConfigureAwait(false);
            return categories.FirstOrDefault(c => c.Id == parentId)?.SubcategoryIds ?? string.Empty;
        }

        /// <summary>
        /// 将逗号分隔的 ID 字符串整理为去空格、不重复的列表。
        /// </summary>
        private static List<string> ParseSubcategoryIds(string ids)
        {
            return (ids ?? string.Empty)
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(id => id.Trim())
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Distinct(StringComparer.Ordinal)
                .ToList();
        }

        /// <summary>
        /// 安全转换旧 params 调用传入的整数参数。
        /// </summary>
        private static bool TryConvertToInt(object value, out int result)
        {
            result = 0;
            return value != null && int.TryParse(value.ToString(), out result);
        }

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
        /// <summary>
        /// 删除cad子分类（最小可编译实现，实际应包含完整逻辑）
        /// </summary>
        /// <param name="args"> </param>
        /// <returns></returns>
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
            /// <summary>
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
            /// <summary>
            /// 部门 ID
            /// </summary>
            public int Id { get; set; }
            /// <summary>
            /// 部门名称（唯一标识）
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            /// 部门显示名称（可重复，允许为空，默认同 Name）
            /// </summary>
            public string DisplayName { get; set; }
            /// <summary>
            /// 排序字段，整数，默认 0
            /// </summary>
            public int SortOrder { get; set; }
            /// <summary>
            /// 是否有效，布尔值，默认 true
            /// </summary>
            public bool IsActive { get; set; }
            /// <summary>
            /// 创建时间，DateTime，默认当前时间
            /// </summary>
            public DateTime CreatedAt { get; set; }
            /// <summary>
            /// 更新时间，DateTime，默认当前时间
            /// </summary>
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
            // 先做服务层校验，避免把空分类名称写入两个数据库。
            if (category == null || string.IsNullOrWhiteSpace(category.Name))
            {
                return 0;
            }

            try
            {
                // 使用统一连接入口，由适配器决定创建 MySQL 还是达梦连接。
                using var connection = GetConnection();

                // SQL 使用中性 @ 参数，ExecuteWriteAsync 会按数据库类型转换并绑定参数。
                const string sql = @"INSERT INTO cad_categories (name, display_name, sort_order)
                    VALUES (@Name, @DisplayName, @SortOrder)";

                // 显示名称为空时使用分类名称，保持历史 MySQL 数据行为一致。
                category.DisplayName = string.IsNullOrWhiteSpace(category.DisplayName)
                    ? category.Name.Trim()
                    : category.DisplayName.Trim();

                // 去除名称首尾空格，避免出现肉眼相同但实际不同的分类。
                category.Name = category.Name.Trim();

                // 统一执行写入，兼容 MySQL 的 Dapper 路径和达梦的原生 ADO.NET 路径。
                var affected = await ExecuteWriteAsync(connection, null, sql, category).ConfigureAwait(false);

                // 保留原有部门同步行为；分类记录成功后再异步修正映射。
                if (affected > 0)
                {
                    _ = SyncDepartmentsFromCadCategoriesAsync();
                }

                return affected;
            }
            catch (Exception ex)
            {
                // 记录数据库类型和异常，便于后续分别排查 MySQL/达梦差异。
                LogManager.Instance.LogInfo($"AddCadCategoryAsync({_adapter.DatabaseType}) 出错: {ex.Message}");
                return 0;
            }
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
            // COALESCE 同时被 MySQL 和达梦支持，统一返回空表时的 0。
            const string sql = "SELECT COALESCE(MAX(sort_order), 0) FROM cad_categories";

            try
            {
                // 不再直接 new MySqlConnection，避免达梦模式仍然访问 MySQL。
                using var connection = GetConnection();
                connection.Open();

                // 使用适配器规范化 SQL，保留同一套查询代码供两个数据库使用。
                using var command = connection.CreateCommand();
                command.CommandText = _adapter.NormalizeSql(sql);

                // MAX 结果可能是数据库特定的数值类型，因此统一通过 Convert 转换。
                var result = await Task.Run(() => command.ExecuteScalar()).ConfigureAwait(false);
                return result == null || result == DBNull.Value ? 0 : Convert.ToInt32(result);
            }
            catch (Exception ex)
            {
                // 排序号查询失败时不能静默使用不确定值，记录后抛出给分类管理器处理。
                LogManager.Instance.LogInfo($"GetMaxCadCategorySortOrderAsync({_adapter.DatabaseType}) 出错: {ex.Message}");
                throw;
            }
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

        #endregion

        #region 优化的文件管理方法

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

        public async Task<(FileStorage File, Dictionary<string, string> Attributes, string ConfigName)> GetFileStorageWithAttributesByHashAsync(
            string filehash,
            string preferredConfigName = null)
        {
            // 1. 获取主记录
            var fileStorage = await GetFileStorageAsync(filehash).ConfigureAwait(false);
            if (fileStorage == null)
                return (null, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), string.Empty);

            // 2. 构造 SQL（根据数据库类型使用正确占位符和限制子句）
            string GetAttrSql(bool byConfig)
            {
                string prefix = _adapter.DatabaseType == "MySQL" ? "@" : ":";
                string sql;
                if (byConfig)
                {
                    sql = $@"SELECT config_name AS ConfigName, attributes_json AS AttributesJson
FROM cad_block_attributes_json
WHERE file_id = {prefix}FileId
  AND config_name = {prefix}ConfigName
ORDER BY attr_id DESC";
                }
                else
                {
                    sql = $@"SELECT config_name AS ConfigName, attributes_json AS AttributesJson
FROM cad_block_attributes_json
WHERE file_id = {prefix}FileId
ORDER BY attr_id DESC";
                }
                sql += _adapter.DatabaseType == "MySQL" ? " LIMIT 1" : " FETCH FIRST 1 ROWS ONLY";
                return sql;
            }

            try
            {
                // 修正配置名取值：优先传入参数，否则使用 "default"
                string configName = !string.IsNullOrWhiteSpace(preferredConfigName)
                    ? preferredConfigName.Trim()
                    : "default";

                // ---- MySQL 分支 ----
                if (_adapter.DatabaseType == "MySQL")
                {
                    using var connection = new MySqlConnection(_connectionString);

                    (string ConfigName, string AttributesJson)? row = null;

                    if (!string.IsNullOrWhiteSpace(configName))
                    {
                        row = await connection.QueryFirstOrDefaultAsync<(string, string)>(
                            GetAttrSql(true), new { FileId = fileStorage.Id, ConfigName = configName }).ConfigureAwait(false);
                    }

                    if (row == null || string.IsNullOrWhiteSpace(row.Value.AttributesJson))
                    {
                        row = await connection.QueryFirstOrDefaultAsync<(string, string)>(
                            GetAttrSql(false), new { FileId = fileStorage.Id }).ConfigureAwait(false);
                    }

                    if (row == null || string.IsNullOrWhiteSpace(row.Value.AttributesJson))
                        return (fileStorage, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), string.Empty);

                    string rawJson = row.Value.AttributesJson;
                    LogManager.Instance.LogDebug($"从数据库读取到的属性JSON（MySQL）：{rawJson}");

                    var dict = JsonConvert.DeserializeObject<Dictionary<string, string>>(rawJson)
                               ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    return (fileStorage, dict, row.Value.ConfigName ?? string.Empty);
                }

                // ---- 达梦（DM）分支 ----
                using var dconn = GetConnection();
                dconn.Open();

                string? attributesJson = null;
                string foundConfigName = string.Empty;

                // 优先按配置名查询
                if (!string.IsNullOrWhiteSpace(configName))
                {
                    using var cmd = dconn.CreateCommand();
                    cmd.CommandText = GetAttrSql(true); // SQL 中已含正确的 :FileId 等
                    AddDmParam(cmd, "FileId", fileStorage.Id);
                    AddDmParam(cmd, "ConfigName", configName);
                    using var rdr = cmd.ExecuteReader();
                    if (rdr.Read())
                    {
                        attributesJson = rdr.IsDBNull(1) ? null : rdr.GetString(1);
                        foundConfigName = rdr.IsDBNull(0) ? string.Empty : rdr.GetString(0);
                    }
                }

                // 兜底查询（取最新一条）
                if (string.IsNullOrWhiteSpace(attributesJson))
                {
                    using var cmd2 = dconn.CreateCommand();
                    cmd2.CommandText = GetAttrSql(false);
                    AddDmParam(cmd2, "FileId", fileStorage.Id);
                    using var rdr2 = cmd2.ExecuteReader();
                    if (rdr2.Read())
                    {
                        attributesJson = rdr2.IsDBNull(1) ? null : rdr2.GetString(1);
                        foundConfigName = rdr2.IsDBNull(0) ? string.Empty : rdr2.GetString(0);
                    }
                }
                // 最终检查：如果仍未获取到有效的 JSON，则返回空字典和空配置名
                if (string.IsNullOrWhiteSpace(attributesJson))
                    return (fileStorage, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), string.Empty);
                // 解析 JSON 到字典（使用不区分大小写的比较器）
                var dictDm = JsonConvert.DeserializeObject<Dictionary<string, string>>(attributesJson)
                             ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                return (fileStorage, dictDm, foundConfigName ?? string.Empty);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"GetFileStorageWithAttributesByHashAsync 出错: {ex.Message}");
                return (fileStorage, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), string.Empty);
            }
        }

        #region 辅助方法

        ///// <summary>
        ///// 根据表别名或视图名，查询真实的表名（支持多别名/视图名），返回第一个有效表名
        ///// 说明：用于动态适配不同环境（如开发/测试/生产）下的表名差异。
        ///// </summary>
        ///// <param name="aliases">候选表别名列表</param>
        ///// <returns>找到的第一个有效表名，或空字符串</returns>
        //private string ResolveExistingTableName(IDbConnection conn, IDbTransaction tx, params string[] aliases)
        //{
        //    if (aliases == null || aliases.Length == 0) return string.Empty;

        //    // 达梦数据库通常对大小写敏感，且存储在系统表中的名称默认是大写的。
        //    // 修正查询逻辑：不仅查 ALL_TABLES，还要处理候选名的 Trim 和 Upper 处理。
        //    string sqlCheckDm = "SELECT TABLE_NAME FROM ALL_TABLES WHERE OWNER = :Owner AND TABLE_NAME = :TableName";
        //    string sqlCheckMySql = "SELECT table_name FROM information_schema.tables WHERE table_schema = DATABASE() AND table_name = @TableName";

        //    var sql = _adapter.DatabaseType == "MySQL" ? sqlCheckMySql : sqlCheckDm;

        //    foreach (var alias in aliases)
        //    {
        //        // 生成候选名称序列：原始、大写、小写
        //        var namesToTry = new[] { alias.Trim(), alias.Trim().ToUpperInvariant(), alias.Trim().ToLowerInvariant() };

        //        foreach (var name in namesToTry.Distinct())
        //        {
        //            try
        //            {
        //                var param = _adapter.DatabaseType == "MySQL"
        //                    ? (object)new { TableName = name }
        //                    : new { Owner = _schemaName, TableName = name };

        //                var found = conn.QuerySingleOrDefault<string>(sql, param, tx);
        //                if (!string.IsNullOrWhiteSpace(found)) return found;
        //            }
        //            catch { continue; }
        //        }
        //    }

        //    // 如果动态查询失败，为了保证业务不中断，强制返回第一个候选名的大写形式（常见于达梦）
        //    return aliases[0].ToUpperInvariant();
        //}

        ///// <summary>
        ///// 读取表的所有列信息（用于动态构建插入/更新语句）
        ///// 返回：包含列名及类型的字典，键为列名（不带前缀），值为数据类型
        ///// </summary>
        ///// <param name="tableName">真实表名</param>
        //private Dictionary<string, string> ReadTableColumns(IDbConnection conn, IDbTransaction tx, string tableName)
        //{
        //    var columns = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //    // 达梦必须精准匹配 TABLE_NAME，通常为大写。
        //    string sqlDm = "SELECT COLUMN_NAME, DATA_TYPE FROM USER_TAB_COLUMNS WHERE TABLE_NAME = :TableName";
        //    string sqlMySql = "SELECT column_name, data_type FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = @TableName";

        //    string sql = _adapter.DatabaseType == "MySQL" ? sqlMySql : sqlDm;

        //    try
        //    {
        //        using var cmd = conn.CreateCommand();
        //        cmd.Transaction = tx;
        //        cmd.CommandText = sql;
        //        _adapter.AddParameter(cmd, "TableName", tableName);

        //        using var reader = cmd.ExecuteReader();
        //        while (reader.Read())
        //        {
        //            // 存入字典，Key 强制不区分大小写
        //            columns[reader.GetString(0)] = reader.GetString(1).ToLower();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        LogManager.Instance.LogInfo($"ReadTableColumns 警告 (表 {tableName}): {ex.Message}");
        //    }
        //    return columns;
        //}

        ///// <summary>
        ///// 执行插入并返回自增 Id（针对 MySQL 和 DM 的统一实现）
        ///// </summary>
        ///// <param name="connection">数据库连接</param>
        ///// <param name="transaction">事务对象</param>
        ///// <param name="tableName">目标表名</param>
        ///// <param name="insertValues">插入字段及对应值</param>
        ///// <returns>插入记录的自增主键 Id</returns>
        //private async Task<long> ExecuteInsertAndReturnIdentity(IDbConnection connection, IDbTransaction transaction, string tableName, object insertValues)
        //{
        //    var columnList = new List<string>();// 用于存储实际参与插入的列名列表
        //    var dp = new DynamicParameters();// 用于存储实际参与插入的参数列表

        //    // 支持 Dictionary
        //    if (insertValues is IDictionary<string, object> dict)
        //    {
        //        // 遍历字典，构建列名列表和参数列表
        //        foreach (var key in dict.Keys)
        //        {
        //            // 只处理数据库实际存在的列，且列名不区分大小写
        //            var col = key.ToUpperInvariant();
        //            // 只有当列存在于数据库表结构中时才添加到插入列表，避免因列缺失导致的错误。
        //            columnList.Add(col);
        //            // 直接使用字典中的值，不强制转换为 DBNull，保持原样传递给 Dapper 处理。
        //            object value = dict[key];
        //            // 关键修复：不强制转 DBNull，保持原样即可
        //            dp.Add(col, value);
        //        }
        //    }
        //    // 支持实体对象
        //    else
        //    {
        //        foreach (var prop in insertValues.GetType().GetProperties())
        //        {
        //            var col = prop.Name.ToUpperInvariant();
        //            columnList.Add(col);
        //            object value = prop.GetValue(insertValues);

        //            // 关键修复：不转 DBNull
        //            dp.Add(col, value);
        //        }
        //    }

        //    if (columnList.Count == 0)
        //        throw new InvalidOperationException("插入对象不包含任何有效字段");

        //    // 构建 SQL
        //    var sb = new System.Text.StringBuilder();
        //    sb.Append("INSERT INTO \"")
        //      .Append(tableName.ToUpperInvariant())
        //      .Append("\" (");

        //    sb.Append(string.Join(",", columnList.Select(c => "\"" + c + "\"")));
        //    sb.Append(") VALUES (");
        //    sb.Append(string.Join(",", columnList.Select(c => ":" + c)));
        //    sb.Append(")");

        //    var sql = sb.ToString();

        //    // 执行插入
        //    int rows = await connection.ExecuteAsync(sql, dp, transaction).ConfigureAwait(false);
        //    if (rows <= 0)
        //        throw new InvalidOperationException($"插入失败，受影响行数为0，表名：{tableName}");

        //    // MySQL 获取ID
        //    if (_adapter.DatabaseType == "MySQL")
        //    {
        //        var id = await connection.ExecuteScalarAsync("SELECT LAST_INSERT_ID();", null, transaction).ConfigureAwait(false);
        //        return Convert.ToInt64(id ?? 0);
        //    }

        //    // ==================== 达梦 ====================
        //    object? res = null;

        //    try { res = await connection.ExecuteScalarAsync("SELECT SCOPE_IDENTITY()", null, transaction).ConfigureAwait(false); } catch { }
        //    if (res == null || res == DBNull.Value)
        //        try { res = await connection.ExecuteScalarAsync("SELECT @@IDENTITY", null, transaction).ConfigureAwait(false); } catch { }
        //    if (res == null || res == DBNull.Value)
        //        try { res = await connection.ExecuteScalarAsync("SELECT IDENTITY_VAL_LOCAL()", null, transaction).ConfigureAwait(false); } catch { }

        //    // 终极兜底
        //    if (res == null || res == DBNull.Value)
        //    {
        //        try
        //        {
        //            string finalSql = $@"SELECT ""ID"" FROM ""{tableName.ToUpperInvariant()}"" WHERE ROWID = (SELECT MAX(ROWID) FROM ""{tableName.ToUpperInvariant()}"")";
        //            res = await connection.ExecuteScalarAsync(finalSql, null, transaction).ConfigureAwait(false);
        //        }
        //        catch (Exception ex)
        //        {
        //            LogManager.Instance.LogInfo($"[致命] 获取插入ID失败：{ex.Message}");
        //            throw new InvalidOperationException("无法获取新插入记录的主键ID", ex);
        //        }
        //    }

        //    return res == null || res == DBNull.Value ? 0L : Convert.ToInt64(res);
        //}

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
