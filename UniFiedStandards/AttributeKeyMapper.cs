using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;

namespace GB_NewCadPlus_IV.UniFiedStandards
{
    /// <summary>
    /// 属性键名映射器
    /// 说明：
    /// - 把“历史中文/历史英文/别名”等映射到统一的 canonical key（英文）
    /// - 提供从实体属性字典到 canonical 字典的转换
    /// - 提供查找可写入块属性的辅助（优先写已有 alias）
    /// </summary>
    public static class AttributeKeyMapper
    {
        // 核心映射：canonical(英文) => 别名集合（包括中文、旧英文）
        // 需要根据你的业务继续补充或从配置文件加载
        private static readonly Dictionary<string, List<string>> _canonicalToAliases = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase)
        {
            // 基本示例（在此处添加完整字段映射）
            {"PIPELINETITLE", new List<string>{ "PIPELINETITLE","管道标题","管道提示标题","PIPELINE_TITLE","PIPE_TITLE" }},
            {"TAG_NO", new List<string>{ "TAG_NO","管段号","管段编号","Pipeline No","Pipe No","管段" }},
            {"NAME", new List<string>{ "NAME","名称","设备名称","位号","Tag" }},
            {"MODEL", new List<string>{ "MODEL","规格型号","规格","型号","Spec" }},
            {"DRAWINGNO_STANDARDNO", new List<string>{ "DRAWINGNO.STANDARDNO","DRAWINGNO_STANDARDNO","图号","设计标准","标准号" }},
            {"DN", new List<string>{ "DN","公称通径","管径" }},
            {"PIPE_OD", new List<string>{ "PIPE_OD","管道外径","外径" }},
            {"PIPE_ID", new List<string>{ "PIPE_ID","管道内径","内径" }},
            {"PIPE_THK", new List<string>{ "PIPE_THK","管道壁厚","壁厚" }},
            {"SCHEDULE", new List<string>{ "SCHEDULE","壁厚等级","SCH" }},
            {"PN", new List<string>{ "PN","公称压力" }},
            {"CLASS", new List<string>{ "CLASS","压力等级" }},
            {"WORK_PRESSURE", new List<string>{ "WORK_PRESSURE","工作压力" }},
            {"DESIGN_PRESSURE", new List<string>{ "DESIGN_PRESSURE","设计压力" }},
            {"DESIGN_TEMP", new List<string>{ "DESIGN_TEMP","设计温度" }},
            {"WORK_TEMP", new List<string>{ "WORK_TEMP","工作温度" }},
            {"TEMP_RANGE", new List<string>{ "TEMP_RANGE","适用温度范围" }},
            {"MEDIUM", new List<string>{ "MEDIUM","介质","适用介质" }},
            {"PIPE_TYPE", new List<string>{ "PIPE_TYPE","管道类型" }},
            {"PIPE_CLASS", new List<string>{ "PIPE_CLASS","管道等级" }},
            {"CONN_TYPE", new List<string>{ "CONN_TYPE","连接方式" }},
            {"PIPE_MATL", new List<string>{ "PIPE_MATL","管道材质","材质","材料" }},
            {"LINING_MATL", new List<string>{ "LINING_MATL","衬里材质","衬里" }},
            {"LINING_THK", new List<string>{ "LINING_THK","衬里厚度" }},
            {"PIPE_LENGTH", new List<string>{ "PIPE_LENGTH","管道长度","长度" }},
            {"FLOW_VEL", new List<string>{ "FLOW_VEL","设计流速","流速" }},
            {"FLOW_RATE", new List<string>{ "FLOW_RATE","介质流量","流量" }},
            {"HOT\\SOUND_ISOLACODE", new List<string>{ "HOT\\SOUND_ISOLACODE", "隔热隔声代号:", "隔热隔声代号" }},
            {"IS_ANTICORRO", new List<string>{ "IS_ANTICORRO", "是否防腐:", "是否防腐" }},
            {"START_POINT", new List<string>{ "START_POINT", "起始点","起始点" }},
            {"END_POINT", new List<string>{ "END_POINT", "终点","终点" }},
            {"QTY", new List<string>{ "QTY","数量","数量(个)" }},
            {"WEIGHT", new List<string>{ "WEIGHT","总重量","重量" }},
            {"REMARK", new List<string>{ "REMARK","备注" }},
            {"SW_MODEL", new List<string>{ "SW_MODEL","SW_MODEL","3D模型","对应3D模型文件名" }},
            {"SYSTEM", new List<string>{ "SYSTEM","系统","所属系统" }}
            // TODO: 继续把你提供的字段全部以相同形式加入
        };

        /// <summary>
        /// 反向索引：alias（包含中文） -> canonical
        /// </summary>
        private static readonly Dictionary<string, string> _aliasToCanonical;

        /// <summary>
        /// 属性键映射器
        /// </summary>
        static AttributeKeyMapper()
        {
            // 构建反向索引
            _aliasToCanonical = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in _canonicalToAliases) // 循环每个 canonical 键及其别名列表
            {
                // 过滤掉空白别名
                var canonical = kv.Key;
                // 遍历别名列表，建立 alias -> canonical 映射
                foreach (var alias in kv.Value.Where(x => !string.IsNullOrWhiteSpace(x)))
                {
                    if (!_aliasToCanonical.ContainsKey(alias)) // 避免重复覆盖
                        _aliasToCanonical[alias] = canonical;  // 记录别名对应的 canonical
                }
                // 也把 canonical 自身映射
                if (!_aliasToCanonical.ContainsKey(canonical)) // 避免重复覆盖
                    _aliasToCanonical[canonical] = canonical;  // canonical 自身映射到自己
            }
        }

        /// <summary>
        /// 获取 canonical（首选英文）对应的别名列表（包含 canonical 本身）
        /// </summary>
        public static IReadOnlyList<string> GetAliases(string canonical)
        {
            if (string.IsNullOrWhiteSpace(canonical)) return Array.Empty<string>(); // 空输入返回空列表
            if (_canonicalToAliases.TryGetValue(canonical, out var list)) return list; // 找到对应的别名列表
            return new[] { canonical }; // 未找到则返回仅包含 canonical 本身的列表
        }

        /// <summary>
        /// 尝试把任意键名（中文/旧英文/英文）映射到 canonical 英文键
        /// 若找不到，返回原字符串但会 Trim 并大写
        /// </summary>
        public static string ToCanonicalKey(string rawKey)
        {
            if (string.IsNullOrWhiteSpace(rawKey)) return string.Empty;
            var k = rawKey.Trim();
            if (_aliasToCanonical.TryGetValue(k, out var c)) return c;
            // 宽松匹配：去掉空白、点、下划线并比较
            var kNorm = NormalizeKeySimple(k);
            foreach (var alias in _aliasToCanonical.Keys)
            {
                if (NormalizeKeySimple(alias).Equals(kNorm, StringComparison.OrdinalIgnoreCase))
                    return _aliasToCanonical[alias];
            }
            // 兜底：把空格与特殊字符替换并返回上层作为 canonical（转换为大写）
            return k.Replace(" ", "_").Replace(".", "_").ToUpperInvariant();
        }

        /// <summary>
        /// 把原始属性字典（可能含中文键）映射为以 canonical 英文键为 Key 的字典（返回新字典）
        /// - 如果出现多种 alias 指向同一 canonical，优先取第一个非空值
        /// - 原始 map 不被修改
        /// </summary>
        public static Dictionary<string, string> MapToCanonical(IDictionary<string, string> rawMap)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (rawMap == null) return result;

            // 先把所有原键做映射并收集候选值（canonical -> list of (alias,value)）
            var candidates = new Dictionary<string, List<(string alias, string value)>>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in rawMap)
            {
                if (kv.Key == null) continue;
                var c = ToCanonicalKey(kv.Key);
                if (!candidates.ContainsKey(c)) candidates[c] = new List<(string, string)>();
                candidates[c].Add((kv.Key, kv.Value ?? string.Empty));
            }

            // 对每个 canonical 取首个非空（且非“0/0mm”等可跳过值按项目策略决定）的值
            foreach (var kv in candidates)
            {
                string chosen = string.Empty;
                foreach (var pair in kv.Value)
                {
                    var v = (pair.value ?? string.Empty).Trim();
                    if (!string.IsNullOrEmpty(v))
                    {
                        chosen = v;
                        break;
                    }
                }
                // 仍为空则记录空串
                result[kv.Key] = chosen ?? string.Empty;
            }

            return result;
        }

        /// <summary>
        /// 在属性字典中按照 canonicalKey 查找首个非空值（考虑所有 alias）
        /// </summary>
        public static bool TryFindFirstValue(IDictionary<string, string> rawMap, string canonicalKey, out string value)
        {
            value = string.Empty;
            if (rawMap == null || string.IsNullOrWhiteSpace(canonicalKey)) return false;

            // 1) 直接查找 canonicalKey 本身
            if (rawMap.TryGetValue(canonicalKey, out var v1) && !string.IsNullOrWhiteSpace(v1))
            {
                value = v1.Trim();
                return true;
            }

            // 2) 查找已知 alias
            var aliases = GetAliases(canonicalKey);
            foreach (var a in aliases)
            {
                if (rawMap.TryGetValue(a, out var v) && !string.IsNullOrWhiteSpace(v))
                {
                    value = v.Trim();
                    return true;
                }
            }

            // 3) 宽松匹配：键名包含匹配
            foreach (var kv in rawMap)
            {
                if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                foreach (var a in aliases)
                {
                    if (kv.Key.IndexOf(a, StringComparison.OrdinalIgnoreCase) >= 0 && !string.IsNullOrWhiteSpace(kv.Value))
                    {
                        value = kv.Value.Trim();
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// 写回属性到块参照：
        /// - 优先写入块上已存在的 alias 标签（不改变标签名）
        /// - 若未找到现有属性，则可选择是否写入新的 AttributeReference（默认 false，因为需要修改块定义）
        /// 使用示例： AttributeKeyMapper.WriteAttributeToBlock(tr, blockRef, "PIPELINETITLE", "文本值", allowCreate:false);
        /// </summary>
        public static void WriteAttributeToBlock(Transaction tr, BlockReference br, string canonicalKey, string newValue, bool allowCreate = false)
        {
            if (tr == null || br == null || string.IsNullOrWhiteSpace(canonicalKey)) return;

            // 优先搜索存在的 AttributeReference（按 alias 列表）
            var aliases = GetAliases(canonicalKey);

            foreach (ObjectId aid in br.AttributeCollection)
            {
                try
                {
                    var ar = tr.GetObject(aid, OpenMode.ForWrite) as AttributeReference;
                    if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue;
                    // 匹配任一 alias（严格等于或忽略大小写）
                    foreach (var a in aliases)
                    {
                        if (string.Equals(ar.Tag.Trim(), a.Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            ar.TextString = newValue ?? string.Empty;
                            try { ar.AdjustAlignment(br.Database); } catch { }
                            return;
                        }
                    }
                }
                catch { /* 忽略单个属性写失败 */ }
            }

            // 没找到可写入的属性
            if (!allowCreate) return;

            // 若允许创建：在简单场景可直接添加 AttributeReference 到块参照（但更规范应修改块定义）
            try
            {
                // 直接创建 AttributeReference 并附加到块参照（仅对非常简单的场景）
                // 需要从块定义中找到对应 AttributeDefinition；若无则直接跳过（建议不要在运行时随意新增）
                var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                if (btr == null) return;

                foreach (ObjectId eid in btr)
                {
                    try
                    {
                        var ent = tr.GetObject(eid, OpenMode.ForRead) as Entity;
                        if (ent is AttributeDefinition ad && !string.IsNullOrWhiteSpace(ad.Tag))
                        {
                            // 找到第一个匹配 alias 的定义，基于它创建 AttributeReference
                            if (aliases.Any(a => string.Equals(ad.Tag, a, StringComparison.OrdinalIgnoreCase)))
                            {
                                var ar = new AttributeReference();
                                ar.SetAttributeFromBlock(ad, br.BlockTransform);
                                ar.TextString = newValue ?? string.Empty;
                                ar.Position = ad.Position.TransformBy(br.BlockTransform);
                                ar.Tag = ad.Tag;
                                ar.Layer = br.Layer;
                                br.UpgradeOpen();
                                br.AttributeCollection.AppendAttribute(ar);
                                tr.AddNewlyCreatedDBObject(ar, true);
                                return;
                            }
                        }
                    }
                    catch { }
                }
                // 若仍未找到匹配的 AttributeDefinition，可以考虑创建一个新的 AttributeDefinition 到块定义（风险较高，需谨慎）
            }
            catch { /* 忽略创建失败 */ }
        }

        /// <summary>
        /// 简单规范化：去掉空白并小写，去掉点与下划线用于宽松比较
        /// </summary>
        /// <param name="k"> </param>
        /// <returns></returns>
        private static string NormalizeKeySimple(string k)
        {
            if (string.IsNullOrWhiteSpace(k)) return string.Empty;
            return new string(k.Where(ch => !char.IsWhiteSpace(ch) && ch != '.' && ch != '_').ToArray()).ToLowerInvariant();
        }
    }
}
