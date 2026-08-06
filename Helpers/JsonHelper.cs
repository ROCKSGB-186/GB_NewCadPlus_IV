using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using JsonSerializer = System.Text.Json.JsonSerializer;

namespace GB_NewCadPlus_IV.Helpers
{
    public static class JsonHelper
    {
        /// <summary>
        /// 保存Json文件到本地
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <param name="filePath"></param>
        public static void SaveToFile<T>(T obj, string filePath)
        {
            string json = JsonSerializer.Serialize(obj, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(filePath, json);
        }

        /// <summary>
        /// 加载本地Json文件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static T LoadFromFile<T>(string filePath) where T : new()
        {
            if (!File.Exists(filePath)) return new T();
            string json = File.ReadAllText(filePath);
            return JsonSerializer.Deserialize<T>(json) ?? new T();
        }

        /// <summary>
        /// 判断字段是否在黑名单中（支持归一化后判断）
        /// </summary>
        public static bool IsBlacklistedPropertyKey(string rawKey)
        {
            // 空键直接当黑名单处理
            if (string.IsNullOrWhiteSpace(rawKey)) return true;

            // 原始键去空白
            string key = rawKey.Trim();

            // 原始键命中黑名单
            if (InsertGraphicHelper._propertySyncBlacklist.Contains(key)) return true;

            // 归一化键命中黑名单
            string nKey = InsertGraphicHelper.NormalizePropertyKey(key);
            if (!string.IsNullOrWhiteSpace(nKey) && InsertGraphicHelper._propertySyncBlacklist.Contains(nKey)) return true;

            // 未命中黑名单
            return false;
        }

        /// <summary>
        /// 候选来源实体模型
        /// </summary>
        public sealed class OverlapCandidate
        {
            // 候选实体对象
            public Entity Entity { get; set; }

            // 候选评分（越高越优先）
            public double Score { get; set; }

            // 是否与插入对象同层
            public bool IsSameLayer { get; set; }

            // 实体标识字符串（日志用）
            public string Identity { get; set; } = string.Empty;
        }

        /// <summary>
        /// 获取“所有重叠候选”并按评分排序（第二版增强）
        /// </summary>
        public static List<OverlapCandidate> FindOverlappedCandidates(DBTrans tr, BlockReference insertingBr, int maxCandidates = 8)
        {
            // 准备结果集合
            var list = new List<OverlapCandidate>();
            var totalTimer = Stopwatch.StartNew();
            long openTicks = 0;
            long extentsTicks = 0;
            long overlapTicks = 0;
            long scoreTicks = 0;
            int enumeratedCount = 0;
            int openedCount = 0;
            int extentsRejectedCount = 0;
            int overlapRejectedCount = 0;

            // 参数校验
            if (tr == null || insertingBr == null) return list;

            // 读取插入对象图层
            string insertLayer = string.Empty;
            try { insertLayer = (insertingBr.Layer ?? string.Empty).Trim(); } catch { insertLayer = string.Empty; }

            // 插入实体包围盒在本次查找中保持不变，只读取一次
            if (!InsertGraphicHelper.TryGetEntityExtents(insertingBr, out var insertExtents)) return list;

            // 遍历当前空间全部实体
            foreach (ObjectId id in tr.CurrentSpace)
            {
                enumeratedCount++;
                // 跳过自身
                if (id == insertingBr.ObjectId) continue;

                // 读取候选实体
                var openTimer = Stopwatch.StartNew();
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                openTimer.Stop();
                openTicks += openTimer.ElapsedTicks;
                openedCount++;
                // 无效实体跳过
                if (ent == null || ent.IsErased) continue;

                // 包围盒是保守的粗筛选，排除不可能重叠的实体
                var extentsTimer = Stopwatch.StartNew();
                bool hasExtents = InsertGraphicHelper.TryGetEntityExtents(ent, out var entityExtents);
                bool extentsIntersect = hasExtents && InsertGraphicHelper.IsExtentsIntersect(insertExtents, entityExtents);
                extentsTimer.Stop();
                extentsTicks += extentsTimer.ElapsedTicks;
                if (!extentsIntersect)
                {
                    extentsRejectedCount++;
                    continue;
                }

                // 重叠判定，未重叠跳过
                var overlapTimer = Stopwatch.StartNew();
                bool isOverlap = InsertGraphicHelper.IsEntityOverlap(insertingBr, ent, insertExtents, entityExtents);
                overlapTimer.Stop();
                overlapTicks += overlapTimer.ElapsedTicks;
                if (!isOverlap)
                {
                    overlapRejectedCount++;
                    continue;
                }

                // 计算候选分数
                var scoreTimer = Stopwatch.StartNew();
                double score = InsertGraphicHelper.ComputeOverlapCandidateScore(insertingBr, ent, insertExtents, entityExtents);
                scoreTimer.Stop();
                scoreTicks += scoreTimer.ElapsedTicks;

                // 记录同层标识
                bool sameLayer = false;
                try
                {
                    string l2 = (ent.Layer ?? string.Empty).Trim();
                    sameLayer = !string.IsNullOrWhiteSpace(insertLayer) &&
                                !string.IsNullOrWhiteSpace(l2) &&
                                string.Equals(insertLayer, l2, StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    sameLayer = false;
                }

                // 组装候选对象
                var candidate = new OverlapCandidate
                {
                    Entity = ent,
                    Score = score,
                    IsSameLayer = sameLayer,
                    Identity = $"Id={ent.ObjectId},Type={ent.GetType().Name},Layer={ent.Layer}"
                };

                // 加入候选列表
                list.Add(candidate);
            }

            // 排序规则：优先同层（可配置）+ 再按评分降序
            IEnumerable<OverlapCandidate> ordered = list;
            if (InsertGraphicHelper._propertySyncPreferSameLayer)
            {
                ordered = ordered
                    .OrderByDescending(c => c.IsSameLayer)
                    .ThenByDescending(c => c.Score);
            }
            else
            {
                ordered = ordered.OrderByDescending(c => c.Score);
            }

            // 截断候选数量，控制性能
            var result = ordered.Take(Math.Max(1, maxCandidates)).ToList();

            // 输出候选日志
            try
            {
                LogManager.Instance.LogInfo($"\n重叠候选数量: 原始={list.Count}, 参与合并={result.Count}");
                LogManager.Instance.LogInfo($"\n重叠查找性能: 总耗时={totalTimer.Elapsed.TotalMilliseconds:F2}ms, 枚举={enumeratedCount}, 打开={openedCount}, 包围盒排除={extentsRejectedCount}, 几何排除={overlapRejectedCount}, 对象打开={openTicks * 1000.0 / Stopwatch.Frequency:F2}ms, 包围盒={extentsTicks * 1000.0 / Stopwatch.Frequency:F2}ms, 几何判定={overlapTicks * 1000.0 / Stopwatch.Frequency:F2}ms, 评分={scoreTicks * 1000.0 / Stopwatch.Frequency:F2}ms");
                for (int i = 0; i < result.Count; i++)
                {
                    var c = result[i];
                    LogManager.Instance.LogInfo($"\n候选[{i + 1}] Score={c.Score:F6}, SameLayer={c.IsSameLayer}, {c.Identity}");
                }
            }
            catch
            {
                // 日志异常不影响流程
            }

            // 返回候选集合
            return result;
        }

        /// <summary>
        /// 白名单模板配置（第三版增强）
        /// 说明：键使用“归一化字段名”（NormalizePropertyKey 后）
        /// </summary>
        // 是否启用白名单模板控制（启用后，仅允许白名单字段被继承）
        public static readonly bool _propertySyncUseWhitelistTemplate = false;

        /// <summary>
        /// 解析当前插入图元所属专业模板键
        /// </summary>
        public static string ResolvePropertySyncTemplateKey(BlockReference insertingBr)
        {
            // 优先从按钮名推断（你项目里按钮名语义最明确）
            string btnName = (VariableDictionary.btnFileName ?? string.Empty).Trim().ToUpperInvariant();
            // 其次从图层名推断
            string layer = string.Empty;
            try { layer = (insertingBr?.Layer ?? string.Empty).Trim().ToUpperInvariant(); } catch { layer = string.Empty; }

            // 拼接统一判断文本
            string text = $"{btnName}|{layer}";

            // 工艺
            if (text.Contains("GY") || text.Contains("工艺")) return "GY";
            // 暖通
            if (text.Contains("NT") || text.Contains("暖通")) return "NT";
            // 给排水
            if (text.Contains("GPS") || text.Contains("给排水") || text.Contains("水")) return "GPS";
            // 电气
            if (text.Contains("DQ") || text.Contains("电气")) return "DQ";
            // 自控
            if (text.Contains("ZK") || text.Contains("自控")) return "ZK";
            // 建筑
            if (text.Contains("JZ") || text.Contains("建筑")) return "JZ";
            // 结构
            if (text.Contains("JG") || text.Contains("结构")) return "JG";

            // 无法识别时走默认模板
            return "DEFAULT";
        }

        #region 第三版
        /// <summary>
        /// 白名单模板 JSON 路径（可手工编辑）
        /// </summary>
        public static readonly string _propertySyncWhitelistJsonPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "GB_NewCadPlus_IV",
                "property-sync-whitelist.json");

        /// <summary>
        /// 白名单模板缓存（热加载后存这里）
        /// </summary>
        public static Dictionary<string, HashSet<string>> _propertySyncWhitelistTemplatesRuntime =
            new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 白名单文件最近写入时间（用于热加载判断）
        /// </summary>
        public static DateTime _propertySyncWhitelistJsonLastWriteUtc = DateTime.MinValue;

        /// <summary>
        /// 白名单热加载锁，避免并发读写冲突
        /// </summary>
        public static readonly object _propertySyncWhitelistLock = new object();

        /// <summary>
        /// 构建内置默认白名单模板（JSON 不存在或解析失败时兜底）
        /// </summary>
        public static Dictionary<string, HashSet<string>> BuildDefaultWhitelistTemplates()
        {
            // 创建默认模板字典
            var dict = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            // 默认模板
            dict["DEFAULT"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","DISPLAYNAME","TYPE","MODEL","SPEC","MATERIAL",
        "DN","PN","QTY","UNIT","REMARK","CODE","NO"
    };

            // 工艺模板
            dict["GY"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","MODEL","SPEC","MATERIAL","DN","PN","QTY","UNIT","REMARK",
        "PIPEMATERIAL","PIPESPEC","VALVEMODEL","PUMPMODEL","WORKINGPRESSURE"
    };

            // 暖通模板
            dict["NT"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","MODEL","SPEC","MATERIAL","DN","QTY","UNIT","REMARK",
        "AIRVOLUME","WINDSPEED","PIPEMATERIAL","INSULATIONTHICKNESS"
    };

            // 给排水模板
            dict["GPS"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","MODEL","SPEC","MATERIAL","DN","QTY","UNIT","REMARK",
        "PRESSURE","PIPEMATERIAL","PIPELEVEL"
    };

            // 电气模板
            dict["DQ"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","MODEL","SPEC","QTY","UNIT","REMARK",
        "POWERRATING","VOLTAGE","CABLESPEC","CABLEMODEL"
    };

            // 自控模板
            dict["ZK"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","MODEL","SPEC","QTY","UNIT","REMARK",
        "SIGNALTYPE","IOPOINT","CONTROLMODE"
    };

            // 建筑模板
            dict["JZ"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","TYPE","SPEC","QTY","UNIT","REMARK","ROOMNO","LEVEL"
    };

            // 结构模板
            dict["JG"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","TYPE","SPEC","QTY","UNIT","REMARK","CONCRETEGRADE","STEELGRADE"
    };

            // 返回默认模板
            return dict;
        }

        /// <summary>
        /// 把模板字段做归一化（与属性匹配规则一致）
        /// </summary>
        public static HashSet<string> NormalizeWhitelistFields(IEnumerable<string> fields)
        {
            // 创建结果集合
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (fields == null) return set;

            // 逐个归一化后入集合
            foreach (var f in fields)
            {
                string n = InsertGraphicHelper.NormalizePropertyKey(f ?? string.Empty);
                if (!string.IsNullOrWhiteSpace(n))
                    set.Add(n);
            }

            return set;
        }

        /// <summary>
        /// 尝试把默认模板写出到 JSON（首次生成，方便用户编辑）
        /// </summary>
        public static void EnsureWhitelistSeedJson()
        {
            try
            {
                // 文件已存在则不覆盖
                if (System.IO.File.Exists(_propertySyncWhitelistJsonPath)) return;

                // 确保目录存在
                string dir = System.IO.Path.GetDirectoryName(_propertySyncWhitelistJsonPath) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(dir) && !System.IO.Directory.Exists(dir))
                    System.IO.Directory.CreateDirectory(dir);

                // 取默认模板（写出原始键，不做归一化，便于人读）
                var defaults = BuildDefaultWhitelistTemplates()
                    .ToDictionary(k => k.Key, v => v.Value.ToList(), StringComparer.OrdinalIgnoreCase);

                // 序列化为缩进 JSON
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(defaults, Newtonsoft.Json.Formatting.Indented);

                // 写文件（UTF-8）
                System.IO.File.WriteAllText(_propertySyncWhitelistJsonPath, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n白名单模板种子 JSON 生成失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 尝试从 JSON 读取白名单模板
        /// </summary>
        public static Dictionary<string, HashSet<string>> LoadWhitelistTemplatesFromJson()
        {
            // 准备返回字典
            var result = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            // 文件不存在直接返回空（上层会兜底）
            if (!System.IO.File.Exists(_propertySyncWhitelistJsonPath))
                return result;

            // 读取 JSON 文本
            string json = System.IO.File.ReadAllText(_propertySyncWhitelistJsonPath, Encoding.UTF8);

            // 反序列化为 字典<模板名, 字段列表>
            var raw = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(json);

            // 空配置直接返回空
            if (raw == null || raw.Count == 0) return result;

            // 逐模板归一化
            foreach (var kv in raw)
            {
                string key = (kv.Key ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(key)) continue;

                var normalizedSet = NormalizeWhitelistFields(kv.Value ?? new List<string>());
                if (normalizedSet.Count > 0)
                    result[key] = normalizedSet;
            }

            return result;
        }

        /// <summary>
        /// 确保白名单模板已热加载（文件变化则自动重载）
        /// </summary>
        public static void EnsureWhitelistTemplatesHotLoaded()
        {
            lock (_propertySyncWhitelistLock)
            {
                // 先确保有种子文件（首次）
                EnsureWhitelistSeedJson();

                // 获取当前文件写入时间
                DateTime lastWrite = DateTime.MinValue;
                if (System.IO.File.Exists(_propertySyncWhitelistJsonPath))
                    lastWrite = System.IO.File.GetLastWriteTimeUtc(_propertySyncWhitelistJsonPath);

                // 缓存为空或文件已更新时才重载
                bool needReload = _propertySyncWhitelistTemplatesRuntime.Count == 0 || lastWrite > _propertySyncWhitelistJsonLastWriteUtc;
                if (!needReload) return;

                try
                {
                    // 优先读 JSON
                    var loaded = LoadWhitelistTemplatesFromJson();

                    // JSON 无有效模板时使用默认模板
                    if (loaded.Count == 0)
                    {
                        var fallback = BuildDefaultWhitelistTemplates();
                        _propertySyncWhitelistTemplatesRuntime = fallback
                            .ToDictionary(k => k.Key, v => NormalizeWhitelistFields(v.Value), StringComparer.OrdinalIgnoreCase);
                    }
                    else
                    {
                        _propertySyncWhitelistTemplatesRuntime = loaded;
                        // 确保 DEFAULT 存在
                        if (!_propertySyncWhitelistTemplatesRuntime.ContainsKey("DEFAULT"))
                            _propertySyncWhitelistTemplatesRuntime["DEFAULT"] = NormalizeWhitelistFields(BuildDefaultWhitelistTemplates()["DEFAULT"]);
                    }

                    // 更新版本时间戳
                    _propertySyncWhitelistJsonLastWriteUtc = lastWrite;

                    LogManager.Instance.LogInfo($"\n白名单模板热加载成功: {_propertySyncWhitelistJsonPath}");
                }
                catch (Exception ex)
                {
                    // 异常时回退默认模板
                    var fallback = BuildDefaultWhitelistTemplates();
                    _propertySyncWhitelistTemplatesRuntime = fallback
                        .ToDictionary(k => k.Key, v => NormalizeWhitelistFields(v.Value), StringComparer.OrdinalIgnoreCase);

                    LogManager.Instance.LogInfo($"\n白名单模板热加载失败，已回退默认模板: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 获取当前激活白名单（热加载版）
        /// </summary>
        public static HashSet<string> GetActiveWhitelist(BlockReference insertingBr)
        {
            // 先执行热加载
            EnsureWhitelistTemplatesHotLoaded();

            // 按专业解析模板键（你已有 ResolvePropertySyncTemplateKey）
            string key = ResolvePropertySyncTemplateKey(insertingBr);

            // 命中专业模板优先
            if (_propertySyncWhitelistTemplatesRuntime.TryGetValue(key, out var set) && set != null && set.Count > 0)
                return set;

            // 兜底 DEFAULT
            if (_propertySyncWhitelistTemplatesRuntime.TryGetValue("DEFAULT", out var def) && def != null)
                return def;

            // 最终兜底空集合
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        #endregion
    }

}
