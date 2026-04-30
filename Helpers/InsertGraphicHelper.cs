using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_IV.Helpers
{
    public static class InsertGraphicHelper
    {
        /// <summary>
        /// 存储用户点击的点
        /// </summary>
        public static List<Point3d> pointS = new List<Point3d>();

        #region 将外部 DWG 文件插入到当前图纸鼠标指定点 可以插入天正 TCH 实体保留天正等自定义实体属性的快速方法（整图复制）


        //SyncCommonPropertiesToBlockReference   _propertySyncMaxCandidates

        #endregion


        #region  插入图元的核心方法（CopyDwgAllFast）和相关字段

        #region 读取属性和判定重叠的辅助方法

        /// <summary>
        /// 尝试安全获取实体包围盒（防止部分实体抛异常）
        /// </summary>
        public static bool TryGetEntityExtents(Entity entity, out Extents3d extents)
        {
            // 先给 out 参数一个默认值，避免未赋值异常
            extents = default;
            // 空实体直接返回失败
            if (entity == null) return false;
            // 已擦除实体直接返回失败
            if (entity.IsErased) return false;
            try
            {
                // 读取几何包围盒
                extents = entity.GeometricExtents;
                // 读取成功返回 true
                return true;
            }
            catch
            {
                // 读取失败返回 false
                return false;
            }
        }

        /// <summary>
        /// 判断两个包围盒是否相交（含接触）
        /// </summary>
        public static bool IsExtentsIntersect(Extents3d a, Extents3d b, double tol = 1e-6)
        {
            // X 轴左侧分离
            if (a.MaxPoint.X < b.MinPoint.X - tol) return false;
            // X 轴右侧分离
            if (a.MinPoint.X > b.MaxPoint.X + tol) return false;
            // Y 轴下侧分离
            if (a.MaxPoint.Y < b.MinPoint.Y - tol) return false;
            // Y 轴上侧分离
            if (a.MinPoint.Y > b.MaxPoint.Y + tol) return false;
            // Z 轴后侧分离
            if (a.MaxPoint.Z < b.MinPoint.Z - tol) return false;
            // Z 轴前侧分离
            if (a.MinPoint.Z > b.MaxPoint.Z + tol) return false;
            // 通过所有分离轴检测则视为相交
            return true;
        }

        /// <summary>
        /// 粗精结合的重叠判定：先包围盒，再尝试曲线求交
        /// </summary>
        public static bool IsEntityOverlap(Entity source, Entity target)
        {
            // 源实体包围盒获取失败，直接不重叠
            if (!TryGetEntityExtents(source, out var e1)) return false;
            // 目标实体包围盒获取失败，直接不重叠
            if (!TryGetEntityExtents(target, out var e2)) return false;
            // 包围盒不相交直接返回
            if (!IsExtentsIntersect(e1, e2)) return false;

            // 当两者都是曲线时，追加一次更精确的求交判定
            if (source is Curve c1 && target is Curve c2)
            {
                try
                {
                    // 用于接收交点集合
                    var pts = new Point3dCollection();
                    // 执行曲线求交
                    c1.IntersectWith(c2, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero);
                    // 有交点则确认重叠/相交
                    if (pts.Count > 0) return true;
                }
                catch
                {
                    // 曲线求交异常时，保留包围盒相交结论继续走
                }
            }

            // 非曲线场景，包围盒相交即视为重叠
            return true;
        }

        /// <summary>
        /// 计算两个包围盒交叠体积（用于候选评分）
        /// </summary>
        public static double CalcOverlapVolume(Extents3d a, Extents3d b)
        {
            // 计算 X 方向交叠长度
            double dx = Math.Min(a.MaxPoint.X, b.MaxPoint.X) - Math.Max(a.MinPoint.X, b.MinPoint.X);
            // 计算 Y 方向交叠长度
            double dy = Math.Min(a.MaxPoint.Y, b.MaxPoint.Y) - Math.Max(a.MinPoint.Y, b.MinPoint.Y);
            // 计算 Z 方向交叠长度
            double dz = Math.Min(a.MaxPoint.Z, b.MaxPoint.Z) - Math.Max(a.MinPoint.Z, b.MinPoint.Z);

            // 任一方向非正即无交叠体积
            if (dx <= 0 || dy <= 0 || dz <= 0) return 0.0;
            // 返回交叠体积
            return dx * dy * dz;
        }

        /// <summary>
        /// 计算包围盒中心点
        /// </summary>
        public static Point3d GetExtentsCenter(Extents3d e)
        {
            // 按最小点与最大点中点计算中心
            return new Point3d(
                (e.MinPoint.X + e.MaxPoint.X) * 0.5,
                (e.MinPoint.Y + e.MaxPoint.Y) * 0.5,
                (e.MinPoint.Z + e.MaxPoint.Z) * 0.5);
        }

        /// <summary>
        /// 规范化属性键名（用于同名匹配增强）
        /// </summary>
        public static string NormalizePropertyKey(string raw)
        {
            // 空值直接返回空串
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;

            // 转大写并去掉首尾空白
            string s = raw.Trim().ToUpperInvariant();

            // 只保留字母数字（中文字符会被 IsLetter 识别保留）
            var sb = new StringBuilder(s.Length);
            foreach (char ch in s)
            {
                // 保留字母与数字，丢弃空格/下划线/符号
                if (char.IsLetterOrDigit(ch))
                {
                    sb.Append(ch);
                }
            }

            // 返回归一化结果
            return sb.ToString();
        }

        /// <summary>
        /// 统一判定：该值是否应跳过继承（空值或数值零）
        /// </summary>
        public static bool ShouldSkipInheritedValue(string raw)
        {
            // null 转空，避免空引用
            string text = (raw ?? string.Empty).Trim();

            // 空串直接跳过
            if (string.IsNullOrWhiteSpace(text)) return true;

            // 兼容全角 ０
            text = text.Replace('０', '0');

            // 纯字符串 "0" 直接跳过
            if (string.Equals(text, "0", StringComparison.OrdinalIgnoreCase)) return true;

            // 数值可解析且等于 0（如 0.0、0.00、+0、-0）则跳过
            double numeric;
            if (double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out numeric))
            {
                if (Math.Abs(numeric) < 1e-12) return true;
            }
            else if (double.TryParse(text, out numeric))
            {
                if (Math.Abs(numeric) < 1e-12) return true;
            }

            // 其余值允许继承
            return false;
        }

        /// <summary>
        /// 读取实体“类型标识”用于评分（块优先取块名）
        /// </summary>
        public static string GetEntityTypeToken(Entity e)
        {
            // 空实体返回空串
            if (e == null) return string.Empty;

            // 块参照优先读取块名
            if (e is BlockReference br)
            {
                try
                {
                    // Name 通常可直接拿到块名
                    return (br.Name ?? string.Empty).Trim();
                }
                catch
                {
                    // 读取失败回退到类型名
                    return e.GetType().Name;
                }
            }

            // 非块参照返回类型名
            return e.GetType().Name;
        }

        /// <summary>
        /// 对候选重叠实体打分（分数越高越优先）
        /// </summary>
        public static double ComputeOverlapCandidateScore(BlockReference insertingBr, Entity candidate)
        {
            // 空对象直接最低分
            if (insertingBr == null || candidate == null) return double.MinValue;

            // 包围盒读取失败直接最低分
            if (!TryGetEntityExtents(insertingBr, out var srcExt)) return double.MinValue;
            if (!TryGetEntityExtents(candidate, out var dstExt)) return double.MinValue;

            // 必须先满足相交
            if (!IsExtentsIntersect(srcExt, dstExt)) return double.MinValue;

            // 交叠体积比例分（越大越好）
            double overlapVol = CalcOverlapVolume(srcExt, dstExt);
            double srcVol = Math.Max(
                (srcExt.MaxPoint.X - srcExt.MinPoint.X) *
                (srcExt.MaxPoint.Y - srcExt.MinPoint.Y) *
                (srcExt.MaxPoint.Z - srcExt.MinPoint.Z), 1e-9);
            double overlapRatio = overlapVol / srcVol;

            // 中心点距离分（越近越好）
            var c1 = GetExtentsCenter(srcExt);
            var c2 = GetExtentsCenter(dstExt);
            double distance = c1.DistanceTo(c2);
            double distanceScore = 1.0 / (1.0 + distance);

            // 同层加权（工程里同层通常更可能是正确来源）
            double layerScore = 0.0;
            try
            {
                string l1 = (insertingBr.Layer ?? string.Empty).Trim();
                string l2 = (candidate.Layer ?? string.Empty).Trim();
                if (!string.IsNullOrWhiteSpace(l1) && !string.IsNullOrWhiteSpace(l2))
                {
                    if (string.Equals(l1, l2, StringComparison.OrdinalIgnoreCase))
                        layerScore = 0.25;
                    else if (l1.IndexOf(l2, StringComparison.OrdinalIgnoreCase) >= 0 || l2.IndexOf(l1, StringComparison.OrdinalIgnoreCase) >= 0)
                        layerScore = 0.10;
                }
            }
            catch
            {
                // 层名读取失败不影响主流程
            }

            // 类型相似度（块名一致或类型一致）
            double typeScore = 0.0;
            string t1 = GetEntityTypeToken(insertingBr);
            string t2 = GetEntityTypeToken(candidate);
            if (!string.IsNullOrWhiteSpace(t1) && !string.IsNullOrWhiteSpace(t2))
            {
                if (string.Equals(t1, t2, StringComparison.OrdinalIgnoreCase))
                    typeScore = 0.20;
            }

            // 组合总分（可按项目实际继续调权重）
            double score = overlapRatio * 0.55 + distanceScore * 0.25 + layerScore + typeScore;

            // 返回最终评分
            return score;
        }

        /// <summary>
        /// 判断字段是否允许参与继承（白名单优先 + 黑名单兜底）
        /// </summary>
        public static bool IsPropertyKeyAllowed(string rawKey, HashSet<string> activeWhitelist)
        {
            // 空键直接不允许，避免脏数据进入
            if (string.IsNullOrWhiteSpace(rawKey)) return false;

            // 黑名单先拦截（绝对禁止）
            if (IsBlacklistedPropertyKey(rawKey)) return false;

            // 未启用白名单时，黑名单外都允许
            if (!_propertySyncUseWhitelistTemplate) return true;

            // 白名单为空时，不允许任何字段（安全兜底）
            if (activeWhitelist == null || activeWhitelist.Count == 0) return false;

            // 归一化后比对白名单
            string nKey = NormalizePropertyKey(rawKey);
            if (string.IsNullOrWhiteSpace(nKey)) return false;

            // 命中白名单才允许
            return activeWhitelist.Contains(nKey);
        }

        /// <summary>
        /// 读取实体属性映射（支持：块属性 + 扩展字典XRecord）
        /// </summary>
        public static Dictionary<string, string> ReadEntityPropertyMap(DBTrans tr, Entity entity)
        {
            // 创建不区分大小写字典，降低字段大小写差异影响
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 空参数直接返回空字典
            if (tr == null || entity == null) return map;

            // 1) 读取块属性（AttributeReference）
            if (entity is BlockReference br)
            {
                // 遍历块属性集合
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 读取属性引用
                    var ar = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                    // 无效属性跳过
                    if (ar == null) continue;

                    // 读取标签名并清洗
                    string tag = (ar.Tag ?? string.Empty).Trim();
                    // 空标签跳过
                    if (string.IsNullOrWhiteSpace(tag)) continue;

                    // 读取属性值
                    string value = ar.TextString ?? string.Empty;
                    // 值为 0 或空时不加入映射（核心修复）
                    if (ShouldSkipInheritedValue(value)) continue;

                    // 保存原键
                    map[tag] = value;

                    // 保存归一化键（增强同名匹配稳定性）
                    string nTag = NormalizePropertyKey(tag);
                    if (!string.IsNullOrWhiteSpace(nTag) && !map.ContainsKey(nTag))
                    {
                        map[nTag] = value;
                    }
                }
            }

            // 2) 读取扩展字典中的 XRecord（键名作为属性名）
            if (entity.ExtensionDictionary != ObjectId.Null)
            {
                // 打开扩展字典
                var dict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                // 字典有效才继续
                if (dict != null)
                {
                    // 遍历扩展字典项
                    foreach (DBDictionaryEntry entry in dict)
                    {
                        // 读取 XRecord
                        var xrec = tr.GetObject(entry.Value, OpenMode.ForRead) as Xrecord;
                        // 无数据跳过
                        if (xrec?.Data == null) continue;

                        // 读取 TypedValue 数组
                        var values = xrec.Data.AsArray();
                        // 空数组跳过
                        if (values == null || values.Length == 0) continue;

                        // 读取键名
                        string key = (entry.Key ?? string.Empty).Trim();
                        // 空键跳过
                        if (string.IsNullOrWhiteSpace(key)) continue;

                        // 优先取首值作为当前同步值
                        string val = values[0].Value?.ToString() ?? string.Empty;
                        // 值为 0 或空时不加入映射（核心修复）
                        if (ShouldSkipInheritedValue(val)) continue;

                        // 保存原键
                        map[key] = val;

                        // 保存归一化键
                        string nKey = NormalizePropertyKey(key);
                        if (!string.IsNullOrWhiteSpace(nKey) && !map.ContainsKey(nKey))
                        {
                            map[nKey] = val;
                        }
                    }
                }
            }

            // 返回属性映射结果
            return map;
        }

        /// <summary>
        /// 将源属性同步到目标块参照（仅同名属性，增强键名匹配）
        /// </summary>
        public static void SyncCommonPropertiesToBlockReference(DBTrans tr, BlockReference targetBr, Dictionary<string, string> sourceMap)
        {
            // 参数校验
            if (tr == null || targetBr == null || sourceMap == null || sourceMap.Count == 0) return;

            // 遍历目标块的属性引用
            foreach (ObjectId attId in targetBr.AttributeCollection)
            {
                // 以写模式打开属性
                var ar = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                // 无效属性跳过
                if (ar == null) continue;

                // 读取目标标签
                string tag = (ar.Tag ?? string.Empty).Trim();
                // 空标签跳过
                if (string.IsNullOrWhiteSpace(tag)) continue;

                // 优先按原键查找
                bool found = sourceMap.TryGetValue(tag, out string val);

                // 原键未命中时按归一化键兜底查找
                if (!found)
                {
                    string nTag = NormalizePropertyKey(tag);
                    if (!string.IsNullOrWhiteSpace(nTag))
                    {
                        found = sourceMap.TryGetValue(nTag, out val);
                    }
                }

                // 未命中同名字段则跳过
                if (!found) continue;

                // null 统一转空串
                string newValue = val ?? string.Empty;

                // 值为 0 或空时，不继承赋值（核心修复）
                if (ShouldSkipInheritedValue(newValue)) continue;

                // 值没变化就不写，减少无效写事务
                if (string.Equals(ar.TextString ?? string.Empty, newValue, StringComparison.Ordinal)) continue;

                // 执行覆盖写入
                ar.TextString = newValue;
            }
        }

        /// <summary>
        /// 根据目标 TypedValue 的类型码把字符串转换为对应对象（用于尽量保持 XRecord 原类型）
        /// </summary>
        public static object ConvertValueByTypeCode(int typeCode, string raw)
        {
            // 空值统一按空串处理
            string text = raw ?? string.Empty;

            // 整型类型码分支
            if (typeCode == (int)DxfCode.Int16 ||
                typeCode == (int)DxfCode.Int32 ||
                typeCode == (int)DxfCode.Int64 ||
                typeCode == (int)DxfCode.ExtendedDataInteger16 ||
                typeCode == (int)DxfCode.ExtendedDataInteger32)
            {
                // 尽量转整型，失败回退原字符串
                if (int.TryParse(text, out int iv)) return iv;
                return text;
            }

            // 实数类型码分支
            if (typeCode == (int)DxfCode.Real ||
                typeCode == (int)DxfCode.ExtendedDataReal)
            {
                // 先按不变文化解析
                if (double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double dv))
                    return dv;
                // 再按当前文化解析
                if (double.TryParse(text, out dv))
                    return dv;
                // 失败回退字符串
                return text;
            }

            // 默认按字符串写入
            return text;
        }

        /// <summary>
        /// 将源属性同步到目标实体扩展字典（仅同名键，增强：尽量保持原类型）
        /// </summary>
        public static void SyncCommonPropertiesToEntityXRecord(DBTrans tr, Entity targetEntity, Dictionary<string, string> sourceMap)
        {
            // 参数校验
            if (tr == null || targetEntity == null || sourceMap == null || sourceMap.Count == 0) return;
            // 目标无扩展字典直接返回
            if (targetEntity.ExtensionDictionary == ObjectId.Null) return;

            // 打开目标扩展字典
            var dict = tr.GetObject(targetEntity.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;
            // 打开失败返回
            if (dict == null) return;

            // 遍历目标扩展字典键
            foreach (DBDictionaryEntry entry in dict)
            {
                // 读取当前键名
                string key = (entry.Key ?? string.Empty).Trim();
                // 空键跳过
                if (string.IsNullOrWhiteSpace(key)) continue;

                // 先按原键匹配
                bool found = sourceMap.TryGetValue(key, out string val);

                // 原键未命中时按归一化键匹配
                if (!found)
                {
                    string nKey = NormalizePropertyKey(key);
                    if (!string.IsNullOrWhiteSpace(nKey))
                    {
                        found = sourceMap.TryGetValue(nKey, out val);
                    }
                }

                // 源中无同名字段跳过
                if (!found) continue;

                // 值为 0 或空时，不继承赋值（核心修复）
                if (ShouldSkipInheritedValue(val)) continue;

                // 打开目标 XRecord
                var xrec = tr.GetObject(entry.Value, OpenMode.ForWrite) as Xrecord;
                // 无效 XRecord 跳过
                if (xrec == null) continue;

                // 读取原始数据数组（用于保留类型）
                var oldArray = xrec.Data?.AsArray();

                // 如果原本没有数据，则创建一个文本 TypedValue
                if (oldArray == null || oldArray.Length == 0)
                {
                    xrec.Data = new ResultBuffer(new TypedValue((int)DxfCode.Text, val ?? string.Empty));
                    continue;
                }

                // 复制原数组用于构建新数组
                var newArray = oldArray.ToArray();

                // 取首项类型码并按类型转换新值
                int firstCode = newArray[0].TypeCode;
                object firstObj = ConvertValueByTypeCode(firstCode, val ?? string.Empty);

                // 若值未变化，直接跳过写入
                string oldText = newArray[0].Value?.ToString() ?? string.Empty;
                string newText = firstObj?.ToString() ?? string.Empty;
                if (string.Equals(oldText, newText, StringComparison.Ordinal)) continue;

                // 只替换首项值，保留其余 TypedValue，降低破坏性
                newArray[0] = new TypedValue(firstCode, firstObj);

                // 回写 XRecord 数据
                xrec.Data = new ResultBuffer(newArray);
            }
        }

        /// <summary>
        /// 对“分解后的新实体集合”执行同名属性同步
        /// </summary>
        public static void SyncCommonPropertiesToInsertedEntities(DBTrans tr, List<Entity> insertedEntities, Dictionary<string, string> sourceMap)
        {
            // 参数校验
            if (tr == null || insertedEntities == null || sourceMap == null || sourceMap.Count == 0) return;

            // 遍历每个新实体
            foreach (var ent in insertedEntities)
            {
                // 空实体跳过
                if (ent == null) continue;

                // 若是块参照，先同步块属性
                if (ent is BlockReference br)
                {
                    SyncCommonPropertiesToBlockReference(tr, br, sourceMap);
                }

                // 同步扩展字典同名键
                SyncCommonPropertiesToEntityXRecord(tr, ent, sourceMap);
            }
        }

        /// <summary>
        /// 从多个候选实体合并属性映射（第三版：白名单模板）
        /// 规则：
        /// 1) 候选已按优先级排序，先到先得
        /// 2) 黑名单字段绝对禁止
        /// 3) 白名单启用时，仅允许白名单字段
        /// 4) 同名（归一化后）字段只取首个来源
        /// 5) 值为 0 或空时，不参与合并
        /// </summary>
        public static Dictionary<string, string> BuildMergedPropertyMapFromCandidates(
            DBTrans tr,
            List<OverlapCandidate> candidates,
            HashSet<string> activeWhitelist)
        {
            // 返回字典（不区分大小写）
            var merged = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 记录已写入归一化键，避免后来源覆盖先来源
            var takenNormalizedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 参数校验
            if (tr == null || candidates == null || candidates.Count == 0) return merged;

            // 遍历候选来源（已排序）
            foreach (var candidate in candidates)
            {
                // 空候选跳过
                if (candidate?.Entity == null) continue;

                // 读取该候选的属性映射
                var sourceMap = ReadEntityPropertyMap(tr, candidate.Entity);
                // 无属性跳过
                if (sourceMap == null || sourceMap.Count == 0) continue;

                // 统计来源贡献字段数
                int acceptedCount = 0;

                // 逐字段合并
                foreach (var kv in sourceMap)
                {
                    // 原始键值
                    string rawKey = kv.Key ?? string.Empty;
                    string rawVal = kv.Value ?? string.Empty;

                    // 值为 0 或空时跳过（核心修复）
                    if (ShouldSkipInheritedValue(rawVal)) continue;

                    // 按白名单/黑名单统一判定
                    if (!IsPropertyKeyAllowed(rawKey, activeWhitelist)) continue;

                    // 归一化键用于去重
                    string nKey = NormalizePropertyKey(rawKey);
                    if (string.IsNullOrWhiteSpace(nKey)) continue;

                    // 更高优先级来源已经写过同键则跳过
                    if (takenNormalizedKeys.Contains(nKey)) continue;

                    // 写入合并结果（使用归一化键）
                    merged[nKey] = rawVal;

                    // 登记占用
                    takenNormalizedKeys.Add(nKey);

                    // 贡献计数
                    acceptedCount++;

                    // 达到上限提前结束
                    if (merged.Count >= _propertySyncMaxMergedFields)
                        break;
                }

                // 记录来源贡献日志
                try
                {
                    LogManager.Instance.LogInfo($"\n来源实体贡献字段: {acceptedCount}, {candidate.Identity}");
                }
                catch { }

                // 达到上限终止后续来源
                if (merged.Count >= _propertySyncMaxMergedFields)
                    break;
            }

            // 总结果日志
            try
            {
                LogManager.Instance.LogInfo($"\n多来源属性合并完成（白名单模式={_propertySyncUseWhitelistTemplate}），合并字段数={merged.Count}");
            }
            catch { }

            return merged;
        }

        #endregion

        //IsCopyDwgAllFastDragging

        #region

        /// <summary>
        /// 属性同步策略配置（第二版增强）
        /// </summary>
        // 是否优先同层来源（true 时同层会加权，且可额外筛选）
        public static readonly bool _propertySyncPreferSameLayer = true;

        // 最多参与合并的重叠候选数量（防止大图性能波动）
        public static readonly int _propertySyncMaxCandidates = 8;

        // 最多合并字段数量（防止异常图元导致字段爆炸）
        public static readonly int _propertySyncMaxMergedFields = 200;

        // 黑名单字段（这些字段不参与“交叠继承”）
        public static readonly HashSet<string> _propertySyncBlacklist = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {   "ID","GUID","UUID","OBJECTID","HANDLE",
            "CREATEDAT","UPDATEDAT","CREATETIME","UPDATETIME","TIMESTAMP",
            "CREATEDBY","UPDATEDBY","USER","USERNAME","OWNER",
            "VERSION","REVISION","REV",
            "FILENAME","FILEPATH","FILEHASH","PREVIEWIMAGEPATH","PREVIEWIMAGENAME",
            "BLOCKNAME","LAYERNAME",
            "NAME", // 新增：禁止继承“名称”字段
            "名称"  // 新增：禁止继承中文“名称”字段
        };

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
        /// 判断字段是否在黑名单中（支持归一化后判断）
        /// </summary>
        public static bool IsBlacklistedPropertyKey(string rawKey)
        {
            // 空键直接当黑名单处理
            if (string.IsNullOrWhiteSpace(rawKey)) return true;

            // 原始键去空白
            string key = rawKey.Trim();

            // 原始键命中黑名单
            if (_propertySyncBlacklist.Contains(key)) return true;

            // 归一化键命中黑名单
            string nKey = NormalizePropertyKey(key);
            if (!string.IsNullOrWhiteSpace(nKey) && _propertySyncBlacklist.Contains(nKey)) return true;

            // 未命中黑名单
            return false;
        }

        /// <summary>
        /// 获取“所有重叠候选”并按评分排序（第二版增强）
        /// </summary>
        public static List<OverlapCandidate> FindOverlappedCandidates(DBTrans tr, BlockReference insertingBr, int maxCandidates = 8)
        {
            // 准备结果集合
            var list = new List<OverlapCandidate>();

            // 参数校验
            if (tr == null || insertingBr == null) return list;

            // 读取插入对象图层
            string insertLayer = string.Empty;
            try { insertLayer = (insertingBr.Layer ?? string.Empty).Trim(); } catch { insertLayer = string.Empty; }

            // 遍历当前空间全部实体
            foreach (ObjectId id in tr.CurrentSpace)
            {
                // 跳过自身
                if (id == insertingBr.ObjectId) continue;

                // 读取候选实体
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                // 无效实体跳过
                if (ent == null || ent.IsErased) continue;

                // 重叠判定，未重叠跳过
                if (!IsEntityOverlap(insertingBr, ent)) continue;

                // 计算候选分数
                double score = ComputeOverlapCandidateScore(insertingBr, ent);

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
            if (_propertySyncPreferSameLayer)
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
                string n = NormalizePropertyKey(f ?? string.Empty);
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

        #endregion

       
        #endregion
    }
}
