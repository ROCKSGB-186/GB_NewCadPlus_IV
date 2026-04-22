using Dapper; // Dapper 微ORM，用于简洁执行 SQL
using GB_NewCadPlus_IV.FunctionalMethod;
using MySql.Data.MySqlClient; // MySQL 连接
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Application = Autodesk.AutoCAD.ApplicationServices.Application; // JavaScriptSerializer（用于 JSON 序列化/反序列化，兼容 .NET Framework 4.8）

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 单条映射项：表示一个原图层名与解释名的对应关系（可有任意多项）
    /// </summary>
    public class LayerMapping
    {
        public string OriginalLayer { get; set; } = ""; // 原图层名称
        public string DicLayer { get; set; } = "";      // 解释后的图层名称
    }

    /// <summary>
    /// 图层字典实体（对应数据库 layer_dictionary 表的一行）
    /// 说明：MappingsJson 字段在数据库中以 JSON 字符串形式保存任意数量的映射对，
    ///       程序使用 Mappings 属性进行访问/修改（会自动序列化/反序列化）。
    /// </summary>
    public class LayerDictionaryHelper
    {
        public int Id { get; set; }                // 主键 id，自增
        public int Seq { get; set; }               // 序号（用于排序）
        public string Major { get; set; }          // 专业/类别（例如：暖通、电气等）
        public string Username { get; set; }       // 记录创建/所属用户名
        public int? UserId { get; set; }           // 可选的用户 id（若需关联 users 表）
        public string Source { get; set; }         // 来源：personal / standard 等

        // 下面使用 MappingsJson 存储任意数量的映射对（数据库列为 TEXT）
        public string MappingsJson { get; set; } = ""; // 存储 JSON 字符串（数据库列）

        // 运行时使用的映射集合（自动与 MappingsJson 同步）
        private List<LayerMapping> _mappings;
        public List<LayerMapping> Mappings
        {
            get
            {
                // 延迟解析：若已解析则直接返回
                if (_mappings != null) return _mappings;

                // 若 JSON 为空则返回空列表
                if (string.IsNullOrWhiteSpace(MappingsJson))
                {
                    _mappings = new List<LayerMapping>();
                    return _mappings;
                }

                try
                {
                    // 使用 JavaScriptSerializer 解析 JSON（与 .NET Framework 4.8 兼容）
                    var ser = new JavaScriptSerializer();
                    var list = ser.Deserialize<List<LayerMapping>>(MappingsJson) ?? new List<LayerMapping>();
                    _mappings = list;
                }
                catch
                {
                    // 若解析失败，退回为空列表以保证稳健性
                    _mappings = new List<LayerMapping>();
                }
                return _mappings;
            }
            set
            {
                // 设置内存集合并更新 JSON 文本
                _mappings = value ?? new List<LayerMapping>();
                try
                {
                    var ser = new JavaScriptSerializer();
                    MappingsJson = ser.Serialize(_mappings);
                }
                catch
                {
                    // 序列化失败则置为空字符串
                    MappingsJson = "";
                }
            }
        }

        public string CreatedBy { get; set; }      // 创建者用户名（记录）
        public DateTime? CreatedAt { get; set; }   // 创建时间
        public DateTime? UpdatedAt { get; set; }   // 更新时间
        public int IsActive { get; set; } = 1;     // 是否有效（软删除标志）

        /// <summary>
        /// 确保目标图层存在，如果不存在则创建并设置颜色
        /// </summary>
        /// <param name="tr">数据库事务对象</param>
        /// <param name="preferLayerName">首选图层名称</param>
        /// <param name="colorIndex">颜色索引</param>
        /// <returns>确保存在的图层名称</returns>
        public static string EnsureTargetLayer(DBTrans tr, string? preferLayerName, short colorIndex)
        {
            string layerName = string.IsNullOrWhiteSpace(preferLayerName) ? "0" : preferLayerName;

            if (!tr.LayerTable.Has(layerName))
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return "0";
                using (doc.LockDocument()) // 锁定文档以安全修改
                {
                    // 避免使用 Add(name, lambda) 重载，规避委托绑定异常
                    tr.LayerTable.Add(layerName, colorIndex);

                    var ltr = tr.LayerTable.GetRecord(layerName, OpenMode.ForWrite);
                    if (ltr != null)
                    {
                        ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
                        ltr.IsPlottable = true; // 设置图层为可打印
                    }
                }
            }
            AutoCadHelper.LogWithSafety($"创建新图层：{layerName}");
            return layerName;
        }

    }



    /// <summary>
    /// 为 DatabaseManager 增加图层字典相关的扩展方法（创建表、查询、保存、删除）
    /// 说明：表结构使用 mappings_json 列存储任意数量的映射对（JSON 格式）。
    /// </summary>
    public static class DatabaseManagerLayerDictionaryExtensions
    {
        /// <summary>
        /// 如果不存在则创建 layer_dictionary 表（幂等）
        /// 注意：新增列 mappings_json 用于保存任意数量的映射对（JSON）
        /// </summary>
        public static async Task<bool> CreateLayerDictionaryTableIfNotExistsAsync(this DatabaseManager db)
        {
            if (db == null) return false; // 参数保护
            try
            {
                using var conn = db.GetConnection(); // 获取连接
                await conn.OpenAsync().ConfigureAwait(false); // 打开连接

                // 建表 SQL（包含 mappings_json TEXT 列以存储任意数量的映射对）
                var sql = @"
CREATE TABLE IF NOT EXISTS `layer_dictionary` (
  `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  `seq` INT DEFAULT 0,
  `major` VARCHAR(200) DEFAULT NULL,
  `username` VARCHAR(128) DEFAULT NULL,
  `user_id` INT DEFAULT NULL,
  `source` VARCHAR(32) DEFAULT 'personal',
  `mappings_json` TEXT DEFAULT NULL, -- 存储任意数量的映射对，JSON 格式
  `created_by` VARCHAR(128) DEFAULT NULL,
  `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
  `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `is_active` TINYINT DEFAULT 1,
  INDEX (`username`),
  INDEX (`major`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";

                await conn.ExecuteAsync(sql).ConfigureAwait(false); // 执行建表
                return true; // 返回成功
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"CreateLayerDictionaryTableIfNotExistsAsync 失败: {ex.Message}"); // 记录日志
                return false; // 返回失败
            }
        }

        /// <summary>
        /// 根据用户名获取该用户的个人图层字典（按 seq 排序）
        /// 返回时会把 mappings_json 字段填到 MappingsJson，并可通过 Mappings 访问解析后的集合
        /// </summary>
        public static async Task<List<LayerDictionaryHelper>> GetLayerDictionaryByUsernameAsync(this DatabaseManager db, string username)
        {
            if (db == null) return new List<LayerDictionaryHelper>(); // 参数保护
            const string sql = @"
SELECT
  id AS Id,
  seq AS Seq,
  major AS Major,
  username AS Username,
  user_id AS UserId,
  source AS Source,
  mappings_json AS MappingsJson,
  created_by AS CreatedBy,
  created_at AS CreatedAt,
  updated_at AS UpdatedAt,
  is_active AS IsActive
FROM layer_dictionary
WHERE username = @Username AND is_active = 1
ORDER BY seq;"; // 查询 SQL（注意选取 mappings_json 列）

            try
            {
                using var conn = db.GetConnection(); // 获取连接
                var rows = await conn.QueryAsync<LayerDictionaryHelper>(sql, new { Username = username }).ConfigureAwait(false); // 执行查询
                return rows.AsList(); // 返回列表（MappingsJson 可用于延迟解析）
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetLayerDictionaryByUsernameAsync 出错: {ex.Message}"); // 记录日志
                return new List<LayerDictionaryHelper>(); // 返回空列表
            }
        }

        /// <summary>
        /// 获取“标准”字典（由管理员用户 sa/root/admin 发布），可按专业 major 过滤
        /// </summary>
        public static async Task<List<LayerDictionaryHelper>> GetStandardLayerDictionaryByMajorAsync(this DatabaseManager db, string major)
        {
            if (db == null) return new List<LayerDictionaryHelper>(); // 参数保护
            const string sql = @"
SELECT
  id AS Id,
  seq AS Seq,
  major AS Major,
  username AS Username,
  user_id AS UserId,
  source AS Source,
  mappings_json AS MappingsJson,
  created_by AS CreatedBy,
  created_at AS CreatedAt,
  updated_at AS UpdatedAt,
  is_active AS IsActive
FROM layer_dictionary
WHERE username IN ('sa','root','admin') AND (major = @Major OR @Major IS NULL OR @Major = '') AND is_active = 1
ORDER BY seq;"; // 管理员发布且按 major 可选过滤

            try
            {
                using var conn = db.GetConnection(); // 获取连接
                var rows = await conn.QueryAsync<LayerDictionaryHelper>(sql, new { Major = major }).ConfigureAwait(false); // 执行查询
                return rows.AsList(); // 返回结果
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetStandardLayerDictionaryByMajorAsync 出错: {ex.Message}"); // 记录日志
                return new List<LayerDictionaryHelper>(); // 返回空列表
            }
        }

        /// <summary>
        /// 为某个用户保存（覆盖）他的图层字典：先删除该用户名下已存在的数据，再批量插入新行（事务）
        /// 注意：entries 中的 Mappings 属性会被序列化为 mappings_json 存入数据库
        /// </summary>
        public static async Task<bool> SaveLayerDictionaryForUserAsync(this DatabaseManager db, string username, List<LayerDictionaryHelper> entries)
        {
            if (db == null) return false; // 参数保护
            try
            {
                using var conn = db.GetConnection(); // 获取连接
                await conn.OpenAsync().ConfigureAwait(false); // 打开连接
                using var tx = conn.BeginTransaction(); // 开启事务

                // 删除已存在的用户数据（覆盖策略）
                await conn.ExecuteAsync("DELETE FROM layer_dictionary WHERE username = @Username", new { Username = username }, tx).ConfigureAwait(false);

                // 若有要插入的条目，则依次插入
                if (entries != null && entries.Count > 0)
                {
                    const string insertSql = @"
INSERT INTO layer_dictionary
(seq, major, username, user_id, source, mappings_json, created_by, created_at, updated_at, is_active)
VALUES
(@Seq, @Major, @Username, @UserId, @Source, @MappingsJson, @CreatedBy, NOW(), NOW(), 1);";

                    foreach (var e in entries)
                    {
                        // 确保 MappingsJson 已被设置：如果调用方只设置了 Mappings 列表，此处负责序列化
                        string mappingsJson = e.MappingsJson;
                        if (string.IsNullOrWhiteSpace(mappingsJson))
                        {
                            try
                            {
                                var ser = new JavaScriptSerializer();
                                var list = e.Mappings ?? new List<LayerMapping>();
                                mappingsJson = ser.Serialize(list);
                            }
                            catch
                            {
                                mappingsJson = "";
                            }
                        }

                        var param = new
                        {
                            Seq = e.Seq, // 序号
                            Major = e.Major, // 专业
                            Username = username, // 所属用户名
                            UserId = e.UserId, // 可选用户 id
                            Source = string.IsNullOrEmpty(e.Source) ? "personal" : e.Source, // 来源
                            MappingsJson = mappingsJson, // JSON 映射文本
                            CreatedBy = string.IsNullOrEmpty(e.CreatedBy) ? username : e.CreatedBy // 创建者
                        };
                        await conn.ExecuteAsync(insertSql, param, tx).ConfigureAwait(false); // 插入单条记录
                    }
                }

                await tx.CommitAsync().ConfigureAwait(false); // 提交事务
                return true; // 返回成功
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"SaveLayerDictionaryForUserAsync 失败: {ex.Message}"); // 记录失败日志
                return false; // 返回失败
            }
        }

        /// <summary>
        /// 根据 id 列表删除图层字典行（直接删除）
        /// </summary>
        public static async Task<int> DeleteLayerDictionaryEntriesAsync(this DatabaseManager db, IEnumerable<int> ids)
        {
            if (db == null) return 0; // 参数保护
            var idList = ids?.ToArray() ?? Array.Empty<int>(); // 转数组
            if (idList.Length == 0) return 0; // 无需删除
            const string sql = "DELETE FROM layer_dictionary WHERE id IN @Ids"; // 删除语句
            try
            {
                using var conn = db.GetConnection(); // 获取连接
                return await conn.ExecuteAsync(sql, new { Ids = idList }).ConfigureAwait(false); // 执行删除并返回受影响行数
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"DeleteLayerDictionaryEntriesAsync 失败: {ex.Message}"); // 记录日志
                return 0; // 返回 0 表示删除失败或无影响
            }
        }
    }
}
