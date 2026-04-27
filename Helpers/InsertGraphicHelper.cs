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
        private static List<Point3d> pointS = new List<Point3d>();

        #region 将外部 DWG 文件插入到当前图纸鼠标指定点 可以插入天正 TCH 实体保留天正等自定义实体属性的快速方法（整图复制）


        #region 再次点方向按键的重复插入逻辑

        /// 放在 CopyDwgAllFast 相关静态字段附近
        private static bool _isCopyDwgAllFastDragging;

        /// <summary>
        /// 新增：命令级互斥标志，防止 Drag 期间重复进入导致崩溃
        /// </summary>
        private static int _copyDwgAllFastBusyFlag = 0;

        /// <summary>
        /// 当前是否处于 COPYDWGALLFAST 的 Drag 交互中
        /// </summary>
        public static bool IsCopyDwgAllFastDragging => _isCopyDwgAllFastDragging;

        /// <summary>
        /// 当前是否处于 COPYDWGALLFAST 执行中（含 Drag）
        /// </summary>
        public static bool IsCopyDwgAllFastBusy => System.Threading.Volatile.Read(ref _copyDwgAllFastBusyFlag) == 1;

        /// <summary>
        /// COPYDWGALLFAST 执行完成事件：success=true 表示插入成功，error 为失败原因（可空）
        /// </summary>
        public static event Action<bool, string?>? CopyDwgAllFastCompleted;

        /// <summary>
        /// 尝试进入 COPYDWGALLFAST 临界区
        /// </summary>
        private static bool TryEnterCopyDwgAllFastBusy()
        {
            return System.Threading.Interlocked.CompareExchange(ref _copyDwgAllFastBusyFlag, 1, 0) == 0;
        }

        /// <summary>
        /// 退出 COPYDWGALLFAST 临界区
        /// </summary>
        private static void ExitCopyDwgAllFastBusy()
        {
            System.Threading.Interlocked.Exchange(ref _copyDwgAllFastBusyFlag, 0);
        }

        /// <summary>
        /// 是否存在可重复执行的上次整图插入
        /// </summary>
        public static bool HasRepeatableCopyDwgAllFast()
        {
            if (_lastCopyDwgBytes != null && _lastCopyDwgBytes.Length > 0) return true;
            return !string.IsNullOrWhiteSpace(_lastCopyDwgPath) && System.IO.File.Exists(_lastCopyDwgPath);
        }

        /// <summary>
        /// 供方向键调用：按当前角度重复一次上次整图插入
        /// </summary>
        public static void RepeatLastCopyDwgAllFastFromDirection()
        {
            if (!HasRepeatableCopyDwgAllFast()) return;
            Env.Document.SendStringToExecute("COPYDWGALLFASTLAST ", false, false, false);
        }

        /// <summary>
        /// 统一派发插入结果，避免事件回调影响主流程
        /// </summary>
        private static void RaiseCopyDwgAllFastCompleted(bool success, string? error = null)
        {
            try
            {
                CopyDwgAllFastCompleted?.Invoke(success, error);
            }
            catch
            {
                // 忽略事件回调异常，避免影响命令本身
            }
        }

        #endregion

        #endregion


        #region  插入图元的核心方法（CopyDwgAllFast）和相关字段

        #region 读取属性和判定重叠的辅助方法

        /// <summary>
        /// 尝试安全获取实体包围盒（防止部分实体抛异常）
        /// </summary>
        private static bool TryGetEntityExtents(Entity entity, out Extents3d extents)
        {
            // 中文注释：先给 out 参数一个默认值，避免未赋值异常
            extents = default;
            // 中文注释：空实体直接返回失败
            if (entity == null) return false;
            // 中文注释：已擦除实体直接返回失败
            if (entity.IsErased) return false;
            try
            {
                // 中文注释：读取几何包围盒
                extents = entity.GeometricExtents;
                // 中文注释：读取成功返回 true
                return true;
            }
            catch
            {
                // 中文注释：读取失败返回 false
                return false;
            }
        }

        /// <summary>
        /// 判断两个包围盒是否相交（含接触）
        /// </summary>
        private static bool IsExtentsIntersect(Extents3d a, Extents3d b, double tol = 1e-6)
        {
            // 中文注释：X 轴左侧分离
            if (a.MaxPoint.X < b.MinPoint.X - tol) return false;
            // 中文注释：X 轴右侧分离
            if (a.MinPoint.X > b.MaxPoint.X + tol) return false;
            // 中文注释：Y 轴下侧分离
            if (a.MaxPoint.Y < b.MinPoint.Y - tol) return false;
            // 中文注释：Y 轴上侧分离
            if (a.MinPoint.Y > b.MaxPoint.Y + tol) return false;
            // 中文注释：Z 轴后侧分离
            if (a.MaxPoint.Z < b.MinPoint.Z - tol) return false;
            // 中文注释：Z 轴前侧分离
            if (a.MinPoint.Z > b.MaxPoint.Z + tol) return false;
            // 中文注释：通过所有分离轴检测则视为相交
            return true;
        }

        /// <summary>
        /// 粗精结合的重叠判定：先包围盒，再尝试曲线求交
        /// </summary>
        private static bool IsEntityOverlap(Entity source, Entity target)
        {
            // 中文注释：源实体包围盒获取失败，直接不重叠
            if (!TryGetEntityExtents(source, out var e1)) return false;
            // 中文注释：目标实体包围盒获取失败，直接不重叠
            if (!TryGetEntityExtents(target, out var e2)) return false;
            // 中文注释：包围盒不相交直接返回
            if (!IsExtentsIntersect(e1, e2)) return false;

            // 中文注释：当两者都是曲线时，追加一次更精确的求交判定
            if (source is Curve c1 && target is Curve c2)
            {
                try
                {
                    // 中文注释：用于接收交点集合
                    var pts = new Point3dCollection();
                    // 中文注释：执行曲线求交
                    c1.IntersectWith(c2, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero);
                    // 中文注释：有交点则确认重叠/相交
                    if (pts.Count > 0) return true;
                }
                catch
                {
                    // 中文注释：曲线求交异常时，保留包围盒相交结论继续走
                }
            }

            // 中文注释：非曲线场景，包围盒相交即视为重叠
            return true;
        }

        /// <summary>
        /// 计算两个包围盒交叠体积（用于候选评分）
        /// </summary>
        private static double CalcOverlapVolume(Extents3d a, Extents3d b)
        {
            // 中文注释：计算 X 方向交叠长度
            double dx = Math.Min(a.MaxPoint.X, b.MaxPoint.X) - Math.Max(a.MinPoint.X, b.MinPoint.X);
            // 中文注释：计算 Y 方向交叠长度
            double dy = Math.Min(a.MaxPoint.Y, b.MaxPoint.Y) - Math.Max(a.MinPoint.Y, b.MinPoint.Y);
            // 中文注释：计算 Z 方向交叠长度
            double dz = Math.Min(a.MaxPoint.Z, b.MaxPoint.Z) - Math.Max(a.MinPoint.Z, b.MinPoint.Z);

            // 中文注释：任一方向非正即无交叠体积
            if (dx <= 0 || dy <= 0 || dz <= 0) return 0.0;
            // 中文注释：返回交叠体积
            return dx * dy * dz;
        }

        /// <summary>
        /// 计算包围盒中心点
        /// </summary>
        private static Point3d GetExtentsCenter(Extents3d e)
        {
            // 中文注释：按最小点与最大点中点计算中心
            return new Point3d(
                (e.MinPoint.X + e.MaxPoint.X) * 0.5,
                (e.MinPoint.Y + e.MaxPoint.Y) * 0.5,
                (e.MinPoint.Z + e.MaxPoint.Z) * 0.5);
        }

        /// <summary>
        /// 规范化属性键名（用于同名匹配增强）
        /// </summary>
        private static string NormalizePropertyKey(string raw)
        {
            // 中文注释：空值直接返回空串
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;

            // 中文注释：转大写并去掉首尾空白
            string s = raw.Trim().ToUpperInvariant();

            // 中文注释：只保留字母数字（中文字符会被 IsLetter 识别保留）
            var sb = new StringBuilder(s.Length);
            foreach (char ch in s)
            {
                // 中文注释：保留字母与数字，丢弃空格/下划线/符号
                if (char.IsLetterOrDigit(ch))
                {
                    sb.Append(ch);
                }
            }

            // 中文注释：返回归一化结果
            return sb.ToString();
        }

        /// <summary>
        /// 统一判定：该值是否应跳过继承（空值或数值零）
        /// </summary>
        private static bool ShouldSkipInheritedValue(string raw)
        {
            // 中文注释：null 转空，避免空引用
            string text = (raw ?? string.Empty).Trim();

            // 中文注释：空串直接跳过
            if (string.IsNullOrWhiteSpace(text)) return true;

            // 中文注释：兼容全角 ０
            text = text.Replace('０', '0');

            // 中文注释：纯字符串 "0" 直接跳过
            if (string.Equals(text, "0", StringComparison.OrdinalIgnoreCase)) return true;

            // 中文注释：数值可解析且等于 0（如 0.0、0.00、+0、-0）则跳过
            double numeric;
            if (double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out numeric))
            {
                if (Math.Abs(numeric) < 1e-12) return true;
            }
            else if (double.TryParse(text, out numeric))
            {
                if (Math.Abs(numeric) < 1e-12) return true;
            }

            // 中文注释：其余值允许继承
            return false;
        }

        /// <summary>
        /// 读取实体“类型标识”用于评分（块优先取块名）
        /// </summary>
        private static string GetEntityTypeToken(Entity e)
        {
            // 中文注释：空实体返回空串
            if (e == null) return string.Empty;

            // 中文注释：块参照优先读取块名
            if (e is BlockReference br)
            {
                try
                {
                    // 中文注释：Name 通常可直接拿到块名
                    return (br.Name ?? string.Empty).Trim();
                }
                catch
                {
                    // 中文注释：读取失败回退到类型名
                    return e.GetType().Name;
                }
            }

            // 中文注释：非块参照返回类型名
            return e.GetType().Name;
        }

        /// <summary>
        /// 对候选重叠实体打分（分数越高越优先）
        /// </summary>
        private static double ComputeOverlapCandidateScore(BlockReference insertingBr, Entity candidate)
        {
            // 中文注释：空对象直接最低分
            if (insertingBr == null || candidate == null) return double.MinValue;

            // 中文注释：包围盒读取失败直接最低分
            if (!TryGetEntityExtents(insertingBr, out var srcExt)) return double.MinValue;
            if (!TryGetEntityExtents(candidate, out var dstExt)) return double.MinValue;

            // 中文注释：必须先满足相交
            if (!IsExtentsIntersect(srcExt, dstExt)) return double.MinValue;

            // 中文注释：交叠体积比例分（越大越好）
            double overlapVol = CalcOverlapVolume(srcExt, dstExt);
            double srcVol = Math.Max(
                (srcExt.MaxPoint.X - srcExt.MinPoint.X) *
                (srcExt.MaxPoint.Y - srcExt.MinPoint.Y) *
                (srcExt.MaxPoint.Z - srcExt.MinPoint.Z), 1e-9);
            double overlapRatio = overlapVol / srcVol;

            // 中文注释：中心点距离分（越近越好）
            var c1 = GetExtentsCenter(srcExt);
            var c2 = GetExtentsCenter(dstExt);
            double distance = c1.DistanceTo(c2);
            double distanceScore = 1.0 / (1.0 + distance);

            // 中文注释：同层加权（工程里同层通常更可能是正确来源）
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
                // 中文注释：层名读取失败不影响主流程
            }

            // 中文注释：类型相似度（块名一致或类型一致）
            double typeScore = 0.0;
            string t1 = GetEntityTypeToken(insertingBr);
            string t2 = GetEntityTypeToken(candidate);
            if (!string.IsNullOrWhiteSpace(t1) && !string.IsNullOrWhiteSpace(t2))
            {
                if (string.Equals(t1, t2, StringComparison.OrdinalIgnoreCase))
                    typeScore = 0.20;
            }

            // 中文注释：组合总分（可按项目实际继续调权重）
            double score = overlapRatio * 0.55 + distanceScore * 0.25 + layerScore + typeScore;

            // 中文注释：返回最终评分
            return score;
        }

        /// <summary>
        /// 判断字段是否允许参与继承（白名单优先 + 黑名单兜底）
        /// </summary>
        private static bool IsPropertyKeyAllowed(string rawKey, HashSet<string> activeWhitelist)
        {
            // 中文注释：空键直接不允许，避免脏数据进入
            if (string.IsNullOrWhiteSpace(rawKey)) return false;

            // 中文注释：黑名单先拦截（绝对禁止）
            if (IsBlacklistedPropertyKey(rawKey)) return false;

            // 中文注释：未启用白名单时，黑名单外都允许
            if (!_propertySyncUseWhitelistTemplate) return true;

            // 中文注释：白名单为空时，不允许任何字段（安全兜底）
            if (activeWhitelist == null || activeWhitelist.Count == 0) return false;

            // 中文注释：归一化后比对白名单
            string nKey = NormalizePropertyKey(rawKey);
            if (string.IsNullOrWhiteSpace(nKey)) return false;

            // 中文注释：命中白名单才允许
            return activeWhitelist.Contains(nKey);
        }

        /// <summary>
        /// 读取实体属性映射（支持：块属性 + 扩展字典XRecord）
        /// </summary>
        private static Dictionary<string, string> ReadEntityPropertyMap(DBTrans tr, Entity entity)
        {
            // 中文注释：创建不区分大小写字典，降低字段大小写差异影响
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 中文注释：空参数直接返回空字典
            if (tr == null || entity == null) return map;

            // 1) 读取块属性（AttributeReference）
            if (entity is BlockReference br)
            {
                // 中文注释：遍历块属性集合
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 中文注释：读取属性引用
                    var ar = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                    // 中文注释：无效属性跳过
                    if (ar == null) continue;

                    // 中文注释：读取标签名并清洗
                    string tag = (ar.Tag ?? string.Empty).Trim();
                    // 中文注释：空标签跳过
                    if (string.IsNullOrWhiteSpace(tag)) continue;

                    // 中文注释：读取属性值
                    string value = ar.TextString ?? string.Empty;
                    // 中文注释：值为 0 或空时不加入映射（核心修复）
                    if (ShouldSkipInheritedValue(value)) continue;

                    // 中文注释：保存原键
                    map[tag] = value;

                    // 中文注释：保存归一化键（增强同名匹配稳定性）
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
                // 中文注释：打开扩展字典
                var dict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                // 中文注释：字典有效才继续
                if (dict != null)
                {
                    // 中文注释：遍历扩展字典项
                    foreach (DBDictionaryEntry entry in dict)
                    {
                        // 中文注释：读取 XRecord
                        var xrec = tr.GetObject(entry.Value, OpenMode.ForRead) as Xrecord;
                        // 中文注释：无数据跳过
                        if (xrec?.Data == null) continue;

                        // 中文注释：读取 TypedValue 数组
                        var values = xrec.Data.AsArray();
                        // 中文注释：空数组跳过
                        if (values == null || values.Length == 0) continue;

                        // 中文注释：读取键名
                        string key = (entry.Key ?? string.Empty).Trim();
                        // 中文注释：空键跳过
                        if (string.IsNullOrWhiteSpace(key)) continue;

                        // 中文注释：优先取首值作为当前同步值
                        string val = values[0].Value?.ToString() ?? string.Empty;
                        // 中文注释：值为 0 或空时不加入映射（核心修复）
                        if (ShouldSkipInheritedValue(val)) continue;

                        // 中文注释：保存原键
                        map[key] = val;

                        // 中文注释：保存归一化键
                        string nKey = NormalizePropertyKey(key);
                        if (!string.IsNullOrWhiteSpace(nKey) && !map.ContainsKey(nKey))
                        {
                            map[nKey] = val;
                        }
                    }
                }
            }

            // 中文注释：返回属性映射结果
            return map;
        }

        /// <summary>
        /// 将源属性同步到目标块参照（仅同名属性，增强键名匹配）
        /// </summary>
        private static void SyncCommonPropertiesToBlockReference(DBTrans tr, BlockReference targetBr, Dictionary<string, string> sourceMap)
        {
            // 中文注释：参数校验
            if (tr == null || targetBr == null || sourceMap == null || sourceMap.Count == 0) return;

            // 中文注释：遍历目标块的属性引用
            foreach (ObjectId attId in targetBr.AttributeCollection)
            {
                // 中文注释：以写模式打开属性
                var ar = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                // 中文注释：无效属性跳过
                if (ar == null) continue;

                // 中文注释：读取目标标签
                string tag = (ar.Tag ?? string.Empty).Trim();
                // 中文注释：空标签跳过
                if (string.IsNullOrWhiteSpace(tag)) continue;

                // 中文注释：优先按原键查找
                bool found = sourceMap.TryGetValue(tag, out string val);

                // 中文注释：原键未命中时按归一化键兜底查找
                if (!found)
                {
                    string nTag = NormalizePropertyKey(tag);
                    if (!string.IsNullOrWhiteSpace(nTag))
                    {
                        found = sourceMap.TryGetValue(nTag, out val);
                    }
                }

                // 中文注释：未命中同名字段则跳过
                if (!found) continue;

                // 中文注释：null 统一转空串
                string newValue = val ?? string.Empty;

                // 中文注释：值为 0 或空时，不继承赋值（核心修复）
                if (ShouldSkipInheritedValue(newValue)) continue;

                // 中文注释：值没变化就不写，减少无效写事务
                if (string.Equals(ar.TextString ?? string.Empty, newValue, StringComparison.Ordinal)) continue;

                // 中文注释：执行覆盖写入
                ar.TextString = newValue;
            }
        }

        /// <summary>
        /// 根据目标 TypedValue 的类型码把字符串转换为对应对象（用于尽量保持 XRecord 原类型）
        /// </summary>
        private static object ConvertValueByTypeCode(int typeCode, string raw)
        {
            // 中文注释：空值统一按空串处理
            string text = raw ?? string.Empty;

            // 中文注释：整型类型码分支
            if (typeCode == (int)DxfCode.Int16 ||
                typeCode == (int)DxfCode.Int32 ||
                typeCode == (int)DxfCode.Int64 ||
                typeCode == (int)DxfCode.ExtendedDataInteger16 ||
                typeCode == (int)DxfCode.ExtendedDataInteger32)
            {
                // 中文注释：尽量转整型，失败回退原字符串
                if (int.TryParse(text, out int iv)) return iv;
                return text;
            }

            // 中文注释：实数类型码分支
            if (typeCode == (int)DxfCode.Real ||
                typeCode == (int)DxfCode.ExtendedDataReal)
            {
                // 中文注释：先按不变文化解析
                if (double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double dv))
                    return dv;
                // 中文注释：再按当前文化解析
                if (double.TryParse(text, out dv))
                    return dv;
                // 中文注释：失败回退字符串
                return text;
            }

            // 中文注释：默认按字符串写入
            return text;
        }

        /// <summary>
        /// 将源属性同步到目标实体扩展字典（仅同名键，增强：尽量保持原类型）
        /// </summary>
        private static void SyncCommonPropertiesToEntityXRecord(DBTrans tr, Entity targetEntity, Dictionary<string, string> sourceMap)
        {
            // 中文注释：参数校验
            if (tr == null || targetEntity == null || sourceMap == null || sourceMap.Count == 0) return;
            // 中文注释：目标无扩展字典直接返回
            if (targetEntity.ExtensionDictionary == ObjectId.Null) return;

            // 中文注释：打开目标扩展字典
            var dict = tr.GetObject(targetEntity.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;
            // 中文注释：打开失败返回
            if (dict == null) return;

            // 中文注释：遍历目标扩展字典键
            foreach (DBDictionaryEntry entry in dict)
            {
                // 中文注释：读取当前键名
                string key = (entry.Key ?? string.Empty).Trim();
                // 中文注释：空键跳过
                if (string.IsNullOrWhiteSpace(key)) continue;

                // 中文注释：先按原键匹配
                bool found = sourceMap.TryGetValue(key, out string val);

                // 中文注释：原键未命中时按归一化键匹配
                if (!found)
                {
                    string nKey = NormalizePropertyKey(key);
                    if (!string.IsNullOrWhiteSpace(nKey))
                    {
                        found = sourceMap.TryGetValue(nKey, out val);
                    }
                }

                // 中文注释：源中无同名字段跳过
                if (!found) continue;

                // 中文注释：值为 0 或空时，不继承赋值（核心修复）
                if (ShouldSkipInheritedValue(val)) continue;

                // 中文注释：打开目标 XRecord
                var xrec = tr.GetObject(entry.Value, OpenMode.ForWrite) as Xrecord;
                // 中文注释：无效 XRecord 跳过
                if (xrec == null) continue;

                // 中文注释：读取原始数据数组（用于保留类型）
                var oldArray = xrec.Data?.AsArray();

                // 中文注释：如果原本没有数据，则创建一个文本 TypedValue
                if (oldArray == null || oldArray.Length == 0)
                {
                    xrec.Data = new ResultBuffer(new TypedValue((int)DxfCode.Text, val ?? string.Empty));
                    continue;
                }

                // 中文注释：复制原数组用于构建新数组
                var newArray = oldArray.ToArray();

                // 中文注释：取首项类型码并按类型转换新值
                int firstCode = newArray[0].TypeCode;
                object firstObj = ConvertValueByTypeCode(firstCode, val ?? string.Empty);

                // 中文注释：若值未变化，直接跳过写入
                string oldText = newArray[0].Value?.ToString() ?? string.Empty;
                string newText = firstObj?.ToString() ?? string.Empty;
                if (string.Equals(oldText, newText, StringComparison.Ordinal)) continue;

                // 中文注释：只替换首项值，保留其余 TypedValue，降低破坏性
                newArray[0] = new TypedValue(firstCode, firstObj);

                // 中文注释：回写 XRecord 数据
                xrec.Data = new ResultBuffer(newArray);
            }
        }

        /// <summary>
        /// 对“分解后的新实体集合”执行同名属性同步
        /// </summary>
        private static void SyncCommonPropertiesToInsertedEntities(DBTrans tr, List<Entity> insertedEntities, Dictionary<string, string> sourceMap)
        {
            // 中文注释：参数校验
            if (tr == null || insertedEntities == null || sourceMap == null || sourceMap.Count == 0) return;

            // 中文注释：遍历每个新实体
            foreach (var ent in insertedEntities)
            {
                // 中文注释：空实体跳过
                if (ent == null) continue;

                // 中文注释：若是块参照，先同步块属性
                if (ent is BlockReference br)
                {
                    SyncCommonPropertiesToBlockReference(tr, br, sourceMap);
                }

                // 中文注释：同步扩展字典同名键
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
        private static Dictionary<string, string> BuildMergedPropertyMapFromCandidates(
            DBTrans tr,
            List<OverlapCandidate> candidates,
            HashSet<string> activeWhitelist)
        {
            // 中文注释：返回字典（不区分大小写）
            var merged = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 中文注释：记录已写入归一化键，避免后来源覆盖先来源
            var takenNormalizedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 中文注释：参数校验
            if (tr == null || candidates == null || candidates.Count == 0) return merged;

            // 中文注释：遍历候选来源（已排序）
            foreach (var candidate in candidates)
            {
                // 中文注释：空候选跳过
                if (candidate?.Entity == null) continue;

                // 中文注释：读取该候选的属性映射
                var sourceMap = ReadEntityPropertyMap(tr, candidate.Entity);
                // 中文注释：无属性跳过
                if (sourceMap == null || sourceMap.Count == 0) continue;

                // 中文注释：统计来源贡献字段数
                int acceptedCount = 0;

                // 中文注释：逐字段合并
                foreach (var kv in sourceMap)
                {
                    // 中文注释：原始键值
                    string rawKey = kv.Key ?? string.Empty;
                    string rawVal = kv.Value ?? string.Empty;

                    // 中文注释：值为 0 或空时跳过（核心修复）
                    if (ShouldSkipInheritedValue(rawVal)) continue;

                    // 中文注释：按白名单/黑名单统一判定
                    if (!IsPropertyKeyAllowed(rawKey, activeWhitelist)) continue;

                    // 中文注释：归一化键用于去重
                    string nKey = NormalizePropertyKey(rawKey);
                    if (string.IsNullOrWhiteSpace(nKey)) continue;

                    // 中文注释：更高优先级来源已经写过同键则跳过
                    if (takenNormalizedKeys.Contains(nKey)) continue;

                    // 中文注释：写入合并结果（使用归一化键）
                    merged[nKey] = rawVal;

                    // 中文注释：登记占用
                    takenNormalizedKeys.Add(nKey);

                    // 中文注释：贡献计数
                    acceptedCount++;

                    // 中文注释：达到上限提前结束
                    if (merged.Count >= _propertySyncMaxMergedFields)
                        break;
                }

                // 中文注释：记录来源贡献日志
                try
                {
                    LogManager.Instance.LogInfo($"\n来源实体贡献字段: {acceptedCount}, {candidate.Identity}");
                }
                catch { }

                // 中文注释：达到上限终止后续来源
                if (merged.Count >= _propertySyncMaxMergedFields)
                    break;
            }

            // 中文注释：总结果日志
            try
            {
                LogManager.Instance.LogInfo($"\n多来源属性合并完成（白名单模式={_propertySyncUseWhitelistTemplate}），合并字段数={merged.Count}");
            }
            catch { }

            return merged;
        }

        #endregion

        #region

        /// <summary>
        /// 属性同步策略配置（第二版增强）
        /// </summary>
        // 中文注释：是否优先同层来源（true 时同层会加权，且可额外筛选）
        private static readonly bool _propertySyncPreferSameLayer = true;

        // 中文注释：最多参与合并的重叠候选数量（防止大图性能波动）
        private static readonly int _propertySyncMaxCandidates = 8;

        // 中文注释：最多合并字段数量（防止异常图元导致字段爆炸）
        private static readonly int _propertySyncMaxMergedFields = 200;

        // 中文注释：黑名单字段（这些字段不参与“交叠继承”）
        private static readonly HashSet<string> _propertySyncBlacklist = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {   "ID","GUID","UUID","OBJECTID","HANDLE",
            "CREATEDAT","UPDATEDAT","CREATETIME","UPDATETIME","TIMESTAMP",
            "CREATEDBY","UPDATEDBY","USER","USERNAME","OWNER",
            "VERSION","REVISION","REV",
            "FILENAME","FILEPATH","FILEHASH","PREVIEWIMAGEPATH","PREVIEWIMAGENAME",
            "BLOCKNAME","LAYERNAME",
            "NAME", // 中文注释：新增：禁止继承“名称”字段
            "名称"  // 中文注释：新增：禁止继承中文“名称”字段
        };

        /// <summary>
        /// 候选来源实体模型
        /// </summary>
        private sealed class OverlapCandidate
        {
            // 中文注释：候选实体对象
            public Entity Entity { get; set; }

            // 中文注释：候选评分（越高越优先）
            public double Score { get; set; }

            // 中文注释：是否与插入对象同层
            public bool IsSameLayer { get; set; }

            // 中文注释：实体标识字符串（日志用）
            public string Identity { get; set; } = string.Empty;
        }

        /// <summary>
        /// 判断字段是否在黑名单中（支持归一化后判断）
        /// </summary>
        private static bool IsBlacklistedPropertyKey(string rawKey)
        {
            // 中文注释：空键直接当黑名单处理
            if (string.IsNullOrWhiteSpace(rawKey)) return true;

            // 中文注释：原始键去空白
            string key = rawKey.Trim();

            // 中文注释：原始键命中黑名单
            if (_propertySyncBlacklist.Contains(key)) return true;

            // 中文注释：归一化键命中黑名单
            string nKey = NormalizePropertyKey(key);
            if (!string.IsNullOrWhiteSpace(nKey) && _propertySyncBlacklist.Contains(nKey)) return true;

            // 中文注释：未命中黑名单
            return false;
        }

        /// <summary>
        /// 获取“所有重叠候选”并按评分排序（第二版增强）
        /// </summary>
        private static List<OverlapCandidate> FindOverlappedCandidates(DBTrans tr, BlockReference insertingBr, int maxCandidates = 8)
        {
            // 中文注释：准备结果集合
            var list = new List<OverlapCandidate>();

            // 中文注释：参数校验
            if (tr == null || insertingBr == null) return list;

            // 中文注释：读取插入对象图层
            string insertLayer = string.Empty;
            try { insertLayer = (insertingBr.Layer ?? string.Empty).Trim(); } catch { insertLayer = string.Empty; }

            // 中文注释：遍历当前空间全部实体
            foreach (ObjectId id in tr.CurrentSpace)
            {
                // 中文注释：跳过自身
                if (id == insertingBr.ObjectId) continue;

                // 中文注释：读取候选实体
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                // 中文注释：无效实体跳过
                if (ent == null || ent.IsErased) continue;

                // 中文注释：重叠判定，未重叠跳过
                if (!IsEntityOverlap(insertingBr, ent)) continue;

                // 中文注释：计算候选分数
                double score = ComputeOverlapCandidateScore(insertingBr, ent);

                // 中文注释：记录同层标识
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

                // 中文注释：组装候选对象
                var candidate = new OverlapCandidate
                {
                    Entity = ent,
                    Score = score,
                    IsSameLayer = sameLayer,
                    Identity = $"Id={ent.ObjectId},Type={ent.GetType().Name},Layer={ent.Layer}"
                };

                // 中文注释：加入候选列表
                list.Add(candidate);
            }

            // 中文注释：排序规则：优先同层（可配置）+ 再按评分降序
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

            // 中文注释：截断候选数量，控制性能
            var result = ordered.Take(Math.Max(1, maxCandidates)).ToList();

            // 中文注释：输出候选日志
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
                // 中文注释：日志异常不影响流程
            }

            // 中文注释：返回候选集合
            return result;
        }

        /// <summary>
        /// 白名单模板配置（第三版增强）
        /// 说明：键使用“归一化字段名”（NormalizePropertyKey 后）
        /// </summary>
        // 中文注释：是否启用白名单模板控制（启用后，仅允许白名单字段被继承）
        private static readonly bool _propertySyncUseWhitelistTemplate = false;

        /// <summary>
        /// 解析当前插入图元所属专业模板键
        /// </summary>
        private static string ResolvePropertySyncTemplateKey(BlockReference insertingBr)
        {
            // 中文注释：优先从按钮名推断（你项目里按钮名语义最明确）
            string btnName = (VariableDictionary.btnFileName ?? string.Empty).Trim().ToUpperInvariant();
            // 中文注释：其次从图层名推断
            string layer = string.Empty;
            try { layer = (insertingBr?.Layer ?? string.Empty).Trim().ToUpperInvariant(); } catch { layer = string.Empty; }

            // 中文注释：拼接统一判断文本
            string text = $"{btnName}|{layer}";

            // 中文注释：工艺
            if (text.Contains("GY") || text.Contains("工艺")) return "GY";
            // 中文注释：暖通
            if (text.Contains("NT") || text.Contains("暖通")) return "NT";
            // 中文注释：给排水
            if (text.Contains("GPS") || text.Contains("给排水") || text.Contains("水")) return "GPS";
            // 中文注释：电气
            if (text.Contains("DQ") || text.Contains("电气")) return "DQ";
            // 中文注释：自控
            if (text.Contains("ZK") || text.Contains("自控")) return "ZK";
            // 中文注释：建筑
            if (text.Contains("JZ") || text.Contains("建筑")) return "JZ";
            // 中文注释：结构
            if (text.Contains("JG") || text.Contains("结构")) return "JG";

            // 中文注释：无法识别时走默认模板
            return "DEFAULT";
        }

        #region 第三版
        /// <summary>
        /// 中文注释：白名单模板 JSON 路径（可手工编辑）
        /// </summary>
        private static readonly string _propertySyncWhitelistJsonPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "GB_NewCadPlus_IV",
                "property-sync-whitelist.json");

        /// <summary>
        /// 中文注释：白名单模板缓存（热加载后存这里）
        /// </summary>
        private static Dictionary<string, HashSet<string>> _propertySyncWhitelistTemplatesRuntime =
            new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 中文注释：白名单文件最近写入时间（用于热加载判断）
        /// </summary>
        private static DateTime _propertySyncWhitelistJsonLastWriteUtc = DateTime.MinValue;

        /// <summary>
        /// 中文注释：白名单热加载锁，避免并发读写冲突
        /// </summary>
        private static readonly object _propertySyncWhitelistLock = new object();

        /// <summary>
        /// 构建内置默认白名单模板（JSON 不存在或解析失败时兜底）
        /// </summary>
        private static Dictionary<string, HashSet<string>> BuildDefaultWhitelistTemplates()
        {
            // 中文注释：创建默认模板字典
            var dict = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            // 中文注释：默认模板
            dict["DEFAULT"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","DISPLAYNAME","TYPE","MODEL","SPEC","MATERIAL",
        "DN","PN","QTY","UNIT","REMARK","CODE","NO"
    };

            // 中文注释：工艺模板
            dict["GY"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","MODEL","SPEC","MATERIAL","DN","PN","QTY","UNIT","REMARK",
        "PIPEMATERIAL","PIPESPEC","VALVEMODEL","PUMPMODEL","WORKINGPRESSURE"
    };

            // 中文注释：暖通模板
            dict["NT"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","MODEL","SPEC","MATERIAL","DN","QTY","UNIT","REMARK",
        "AIRVOLUME","WINDSPEED","PIPEMATERIAL","INSULATIONTHICKNESS"
    };

            // 中文注释：给排水模板
            dict["GPS"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","MODEL","SPEC","MATERIAL","DN","QTY","UNIT","REMARK",
        "PRESSURE","PIPEMATERIAL","PIPELEVEL"
    };

            // 中文注释：电气模板
            dict["DQ"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","MODEL","SPEC","QTY","UNIT","REMARK",
        "POWERRATING","VOLTAGE","CABLESPEC","CABLEMODEL"
    };

            // 中文注释：自控模板
            dict["ZK"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","MODEL","SPEC","QTY","UNIT","REMARK",
        "SIGNALTYPE","IOPOINT","CONTROLMODE"
    };

            // 中文注释：建筑模板
            dict["JZ"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","TYPE","SPEC","QTY","UNIT","REMARK","ROOMNO","LEVEL"
    };

            // 中文注释：结构模板
            dict["JG"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "TAG","NAME","TYPE","SPEC","QTY","UNIT","REMARK","CONCRETEGRADE","STEELGRADE"
    };

            // 中文注释：返回默认模板
            return dict;
        }

        /// <summary>
        /// 把模板字段做归一化（与属性匹配规则一致）
        /// </summary>
        private static HashSet<string> NormalizeWhitelistFields(IEnumerable<string> fields)
        {
            // 中文注释：创建结果集合
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (fields == null) return set;

            // 中文注释：逐个归一化后入集合
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
        private static void EnsureWhitelistSeedJson()
        {
            try
            {
                // 中文注释：文件已存在则不覆盖
                if (System.IO.File.Exists(_propertySyncWhitelistJsonPath)) return;

                // 中文注释：确保目录存在
                string dir = System.IO.Path.GetDirectoryName(_propertySyncWhitelistJsonPath) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(dir) && !System.IO.Directory.Exists(dir))
                    System.IO.Directory.CreateDirectory(dir);

                // 中文注释：取默认模板（写出原始键，不做归一化，便于人读）
                var defaults = BuildDefaultWhitelistTemplates()
                    .ToDictionary(k => k.Key, v => v.Value.ToList(), StringComparer.OrdinalIgnoreCase);

                // 中文注释：序列化为缩进 JSON
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(defaults, Newtonsoft.Json.Formatting.Indented);

                // 中文注释：写文件（UTF-8）
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
        private static Dictionary<string, HashSet<string>> LoadWhitelistTemplatesFromJson()
        {
            // 中文注释：准备返回字典
            var result = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            // 中文注释：文件不存在直接返回空（上层会兜底）
            if (!System.IO.File.Exists(_propertySyncWhitelistJsonPath))
                return result;

            // 中文注释：读取 JSON 文本
            string json = System.IO.File.ReadAllText(_propertySyncWhitelistJsonPath, Encoding.UTF8);

            // 中文注释：反序列化为 字典<模板名, 字段列表>
            var raw = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(json);

            // 中文注释：空配置直接返回空
            if (raw == null || raw.Count == 0) return result;

            // 中文注释：逐模板归一化
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
        private static void EnsureWhitelistTemplatesHotLoaded()
        {
            lock (_propertySyncWhitelistLock)
            {
                // 中文注释：先确保有种子文件（首次）
                EnsureWhitelistSeedJson();

                // 中文注释：获取当前文件写入时间
                DateTime lastWrite = DateTime.MinValue;
                if (System.IO.File.Exists(_propertySyncWhitelistJsonPath))
                    lastWrite = System.IO.File.GetLastWriteTimeUtc(_propertySyncWhitelistJsonPath);

                // 中文注释：缓存为空或文件已更新时才重载
                bool needReload = _propertySyncWhitelistTemplatesRuntime.Count == 0 || lastWrite > _propertySyncWhitelistJsonLastWriteUtc;
                if (!needReload) return;

                try
                {
                    // 中文注释：优先读 JSON
                    var loaded = LoadWhitelistTemplatesFromJson();

                    // 中文注释：JSON 无有效模板时使用默认模板
                    if (loaded.Count == 0)
                    {
                        var fallback = BuildDefaultWhitelistTemplates();
                        _propertySyncWhitelistTemplatesRuntime = fallback
                            .ToDictionary(k => k.Key, v => NormalizeWhitelistFields(v.Value), StringComparer.OrdinalIgnoreCase);
                    }
                    else
                    {
                        _propertySyncWhitelistTemplatesRuntime = loaded;
                        // 中文注释：确保 DEFAULT 存在
                        if (!_propertySyncWhitelistTemplatesRuntime.ContainsKey("DEFAULT"))
                            _propertySyncWhitelistTemplatesRuntime["DEFAULT"] = NormalizeWhitelistFields(BuildDefaultWhitelistTemplates()["DEFAULT"]);
                    }

                    // 中文注释：更新版本时间戳
                    _propertySyncWhitelistJsonLastWriteUtc = lastWrite;

                    LogManager.Instance.LogInfo($"\n白名单模板热加载成功: {_propertySyncWhitelistJsonPath}");
                }
                catch (Exception ex)
                {
                    // 中文注释：异常时回退默认模板
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
        private static HashSet<string> GetActiveWhitelist(BlockReference insertingBr)
        {
            // 中文注释：先执行热加载
            EnsureWhitelistTemplatesHotLoaded();

            // 中文注释：按专业解析模板键（你已有 ResolvePropertySyncTemplateKey）
            string key = ResolvePropertySyncTemplateKey(insertingBr);

            // 中文注释：命中专业模板优先
            if (_propertySyncWhitelistTemplatesRuntime.TryGetValue(key, out var set) && set != null && set.Count > 0)
                return set;

            // 中文注释：兜底 DEFAULT
            if (_propertySyncWhitelistTemplatesRuntime.TryGetValue("DEFAULT", out var def) && def != null)
                return def;

            // 中文注释：最终兜底空集合
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        #endregion

        #endregion

        // 上一次“整图插入”缓存
        private static byte[]? _lastCopyDwgBytes;
        private static string? _lastCopyDwgFileNameBase;
        private static string? _lastCopyDwgPath;

        /// <summary>
        /// 按钮侧统一调用：注册“上次插入参数”，并通过命令行触发，保证空格可重复
        /// </summary>
        public static void ExecuteCopyDwgAllFastWithRepeat(string sourceFilePath)
        {
            // 新增：Drag/执行中禁止再次触发，避免命令重入导致 CAD 崩溃
            if (IsCopyDwgAllFastDragging || IsCopyDwgAllFastBusy)
            {
                Env.Editor?.WriteMessage("\n当前图元正在跟随插入，请先左键落点或按 Esc 结束后再触发。");
                return;
            }

            try
            {
                // 优先缓存资源字节（最稳，避免临时文件被删后无法重复）
                if (VariableDictionary.resourcesFile != null && VariableDictionary.resourcesFile.Length > 0)
                {
                    _lastCopyDwgBytes = (byte[])VariableDictionary.resourcesFile.Clone();// 克隆一份字节数组，避免后续被修改
                    // 仅缓存文件名的基础部分，去掉路径和扩展名，避免重复执行时文件名过长或包含非法字符
                    _lastCopyDwgFileNameBase = string.IsNullOrWhiteSpace(VariableDictionary.btnFileName)
                        ? "GB_CopyDwgAllFast"
                        : VariableDictionary.btnFileName;
                    _lastCopyDwgPath = null;// 已缓存字节后路径不可靠，置空避免误用
                }
                else
                {
                    _lastCopyDwgBytes = null;
                    _lastCopyDwgFileNameBase = Path.GetFileNameWithoutExtension(sourceFilePath);
                    _lastCopyDwgPath = sourceFilePath;
                }
            }
            catch
            {
                // 缓存失败不阻断主流程
            }

            // 关键：通过命令执行，这样空格会重复这个命令
            Env.Document.SendStringToExecute("COPYDWGALLFASTLAST ", false, false, false);
        }

        /// <summary>
        /// 可被空格重复的命令：再次执行上一次插入
        /// </summary>
        [CommandMethod("COPYDWGALLFASTLAST")]
        public static void CopyDwgAllFastLast()
        {
            // 新增：执行中直接拦截，防止重入
            if (IsCopyDwgAllFastDragging || IsCopyDwgAllFastBusy)
            {
                Env.Editor?.WriteMessage("\n当前图元正在跟随插入，请先完成当前插入。");
                return;
            }

            try
            {
                string? runPath = null;

                // 如果有字节缓存，则每次重复都新建一个临时文件
                if (_lastCopyDwgBytes != null && _lastCopyDwgBytes.Length > 0)
                {
                    // 生成临时文件路径，使用基础文件名加上 GUID，避免重复执行时文件名过长或包含非法字符
                    string baseName = string.IsNullOrWhiteSpace(_lastCopyDwgFileNameBase)
                        ? "GB_CopyDwgAllFast"
                        : _lastCopyDwgFileNameBase;
                    // 确保基础文件名不包含非法字符
                    foreach (var c in Path.GetInvalidFileNameChars())
                    {
                        baseName = baseName.Replace(c, '_');
                    }
                    // 生成临时文件路径
                    runPath = Path.Combine(Path.GetTempPath(), $"{baseName}_{Guid.NewGuid():N}.dwg");
                    System.IO.File.WriteAllBytes(runPath, _lastCopyDwgBytes);// 写入临时文件
                }
                else
                {
                    runPath = _lastCopyDwgPath;// 没有字节缓存则使用上次的路径（可能是原文件路径，存在被删除风险）
                }
                // 最后再次验证路径有效性，避免误用已被删除的临时文件路径
                if (string.IsNullOrWhiteSpace(runPath) || !System.IO.File.Exists(runPath))
                {
                    Env.Editor.WriteMessage("\n没有可重复的上一次插入命令。");
                    return;
                }

                if (runPath != null)
                    //插入源文件中的图元到当前图纸
                    CopyDwgAllFast(runPath);// 直接调用插入方法，传入路径
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n重复执行上次插入失败: {ex.Message}");
            }
        }


        /// <summary>
        /// 以“整图复制”的方式，将指定 DWG 文件的模型空间全部图元一次性克隆到当前图纸空间，
        /// 等效于在源图里全选 Ctrl+C，再在当前图 Ctrl+V，能够保留天正等自定义实体的附加属性。
        /// 新增功能：如果插入点与现有图元重叠，自动继承重叠图元的业务属性（如压力、介质等）。
        /// </summary>
        [CommandMethod("COPYDWGALLFAST")] // 注册 CAD 命令名，允许命令行调用
        public static void CopyDwgAllFast(string sourceFilePath) // 整图插入主方法，参数为源DWG文件路径
        {
            // 中文注释：获取当前活动的 AutoCAD 文档对象
            var doc = Application.DocumentManager.MdiActiveDocument;
            // 中文注释：如果未找到活动文档，提示错误并退出
            if (doc == null)
            {
                RaiseCopyDwgAllFastCompleted(false, "未找到活动文档。"); // 回调通知上层调用者失败
                return; // 结束方法执行
            }

            // 中文注释：获取编辑器对象，用于向命令行发送调试信息
            var ed = doc.Editor;

            // 中文注释：尝试进入互斥锁，防止命令重入导致崩溃
            if (!TryEnterCopyDwgAllFastBusy())
            {
                ed.WriteMessage("\n当前插入命令正在执行中，请先完成当前插入（左键落点或 Esc）。"); // 命令行提示用户
                RaiseCopyDwgAllFastCompleted(false, "当前命令忙碌中。"); // 回调通知上层调用者失败
                return; // 结束方法执行
            }

            // 中文注释：获取当前文档的数据库对象（目标数据库）
            var destDb = doc.Database;
            // 中文注释：标记是否需要在成功后删除源文件（通常用于临时文件）
            bool deleteAfter = false;
            // 中文注释：标记插入操作是否成功
            bool insertSuccess = false;
            // 中文注释：记录失败时的具体原因
            string? failReason = null;

            try // 主流程异常捕获块
            {
                // 中文注释：判断源文件是否为临时文件，如果是则标记后续删除
                try
                {
                    if (!string.IsNullOrEmpty(sourceFilePath))
                    {
                        var tempDir = Path.GetTempPath();
                        if (!string.IsNullOrEmpty(tempDir) && sourceFilePath.StartsWith(tempDir, StringComparison.OrdinalIgnoreCase))
                        {
                            deleteAfter = true; // 标记为需要删除
                        }
                    }
                }
                catch
                {
                    deleteAfter = false;
                }

                Point3d targetPoint = Point3d.Origin;

                using (doc.LockDocument())
                using (var sourceDb = new Autodesk.AutoCAD.DatabaseServices.Database(false, true))
                {
                    sourceDb.ReadDwgFile(sourceFilePath, FileShare.Read, true, null);
                    sourceDb.CloseInput(true);

                    using (var tr = new DBTrans())
                    {
                        var blkDefId = destDb.Insert("*U", sourceDb, false);
                        var br = new BlockReference(targetPoint, blkDefId);

                        double scale = 1.0;
                        if (VariableDictionary.winForm_Status)
                        {
                            try { scale = VariableDictionary.textBoxScale; }
                            catch { scale = 1.0; }
                        }
                        else
                        {
                            try { scale = VariableDictionary.wpfTextBoxScale; }
                            catch { scale = 1.0; }
                        }
                        if (double.IsNaN(scale) || scale <= 0) scale = 1.0;

                        br.ScaleFactors = new Scale3d(scale);
                        var entityObjectId = tr.CurrentSpace.AddEntity(br);
                        var fileEntity = (BlockReference)tr.GetObject(entityObjectId, OpenMode.ForWrite);

                        double tempAngle = VariableDictionary.entityRotateAngle;
                        double tempScale = scale;

                        if (Math.Abs(tempAngle) > 1e-12)
                        {
                            fileEntity.TransformBy(Matrix3d.Rotation(tempAngle, Vector3d.ZAxis, targetPoint));
                        }

                        var entityBlock = new JigEx((mpw, _) =>
                        {
                            fileEntity.Move(targetPoint, mpw);
                            targetPoint = mpw;

                            if (VariableDictionary.entityRotateAngle != tempAngle)
                            {
                                fileEntity.TransformBy(Matrix3d.Rotation(-tempAngle, Vector3d.ZAxis, targetPoint));
                                tempAngle = VariableDictionary.entityRotateAngle;
                                fileEntity.TransformBy(Matrix3d.Rotation(tempAngle, Vector3d.ZAxis, targetPoint));
                            }

                            double currentUiScale = scale;
                            if (VariableDictionary.winForm_Status)
                            {
                                try { currentUiScale = VariableDictionary.textBoxScale; }
                                catch { currentUiScale = tempScale; }
                            }
                            else
                            {
                                try { currentUiScale = VariableDictionary.wpfTextBoxScale; }
                                catch { currentUiScale = tempScale; }
                            }
                            if (double.IsNaN(currentUiScale) || currentUiScale <= 0) currentUiScale = 1.0;

                            if (Math.Abs(currentUiScale - tempScale) > 1e-9)
                            {
                                fileEntity.ScaleFactors = new Scale3d(currentUiScale);
                                tempScale = currentUiScale;
                            }
                        });

                        entityBlock.DatabaseEntityDraw(wd => wd.Geometry.Draw(fileEntity));
                        entityBlock.SetOptions(msg: "\n指定插入点");
                        EnsureDwgViewFocus();

                        PromptResult endPoint;
                        _isCopyDwgAllFastDragging = true;
                        try
                        {
                            endPoint = Env.Editor.Drag(entityBlock);
                        }
                        finally
                        {
                            _isCopyDwgAllFastDragging = false;
                        }

                        if (endPoint.Status != PromptStatus.OK)
                        {
                            failReason = "用户取消插入。";
                            tr.Abort();
                            return;
                        }

                        // ================== 核心新功能：重叠检测与属性继承（带命令行调试） ==================

                        ed.WriteMessage("\n[调试] >>> 开始执行重叠检测与属性继承逻辑...");

                        // 中文注释：初始化用于存储从重叠图元读取到的属性映射表
                        var overlapSourcePropertyMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                        // 中文注释：查找当前插入点附近所有可能重叠的候选图元
                        var overlapCandidates = FindOverlappedCandidates(tr, fileEntity, _propertySyncMaxCandidates);

                        // 中文注释：在命令行输出找到的候选数量
                        ed.WriteMessage($"\n[调试] 找到重叠候选图元数量: {overlapCandidates.Count}");

                        if (overlapCandidates.Count > 0)
                        {
                            // 中文注释：根据当前图元的专业类型，获取对应的属性白名单
                            var activeWhitelist = GetActiveWhitelist(fileEntity);

                            // 中文注释：从所有候选图元中合并属性，并应用白名单/黑名单过滤
                            overlapSourcePropertyMap = BuildMergedPropertyMapFromCandidates(tr, overlapCandidates, activeWhitelist);

                            // 中文注释：在命令行输出过滤后的属性数量
                            ed.WriteMessage($"\n[调试] 过滤后待同步属性数量: {overlapSourcePropertyMap.Count}");

                            // 中文注释：如果有属性，打印前3个属性的键值对，方便排查
                            if (overlapSourcePropertyMap.Count > 0)
                            {
                                int debugCount = 0;
                                foreach (var kv in overlapSourcePropertyMap)
                                {
                                    ed.WriteMessage($"\n  [调试] 属性名: {kv.Key}, 属性值: {kv.Value}");
                                    debugCount++;
                                    if (debugCount >= 3) break; // 只打印前3个，避免刷屏
                                }
                            }
                            else
                            {
                                // 中文注释：如果数量为0，提示可能是白名单过滤掉了
                                ed.WriteMessage("\n[调试] 警告：未找到可同步属性。请检查白名单配置或原图元是否有扩展数据/块属性。");
                            }
                        }
                        else
                        {
                            // 中文注释：如果没有重叠，提示用户检查插入位置
                            ed.WriteMessage("\n[调试] 未检测到重叠图元。请确保新图元与旧图元有几何交集。");
                        }

                        // 中文注释：如果成功读取到有效的重叠属性
                        if (overlapSourcePropertyMap.Count > 0)
                        {
                            // 中文注释：第一步：将属性同步到当前的块参照（BlockReference）本身
                            SyncCommonPropertiesToBlockReference(tr, fileEntity, overlapSourcePropertyMap);
                            SyncCommonPropertiesToEntityXRecord(tr, fileEntity, overlapSourcePropertyMap);

                            ed.WriteMessage($"\n[调试] 已向块参照同步 {overlapSourcePropertyMap.Count} 个属性。");
                        }

                        // ================== 结束核心新功能 ==================

                        // 中文注释：创建集合用于存储分解后产生的所有新实体
                        var newIds = new DBObjectCollection();

                        // 中文注释：执行分解操作
                        fileEntity.Explode(newIds);

                        // 中文注释：删除原始的块参照实体
                        fileEntity.Erase();

                        // 中文注释：创建列表用于跟踪所有新加入到图纸中的实体
                        var insertedEntities = new List<Entity>();
                        foreach (Entity ent in newIds)
                        {
                            tr.CurrentSpace.AddEntity(ent);
                            insertedEntities.Add(ent);
                        }

                        // 中文注释：如果之前读取到了重叠属性，需要将这些属性进一步同步到分解后的子实体上
                        if (overlapSourcePropertyMap.Count > 0)
                        {
                            ed.WriteMessage($"\n[调试] 正在向 {insertedEntities.Count} 个分解后的实体同步属性...");

                            // 中文注释：遍历分解后的实体，将属性同步到它们的扩展数据或嵌套块中
                            SyncCommonPropertiesToInsertedEntities(tr, insertedEntities, overlapSourcePropertyMap);

                            ed.WriteMessage("\n[调试] 属性同步流程结束。请选中新图元输入 LIST 命令查看扩展数据。");
                        }

                        // 中文注释：检查是否有待创建的标注文本
                        if (VariableDictionary.dimString != null)
                        {
                            try
                            {
                                Command.DDimLinear(tr, entityBlock.MousePointWcsLast, VariableDictionary.dimString);
                                VariableDictionary.dimString = null;
                            }
                            catch (Exception ex)
                            {
                                LogManager.Instance.LogInfo($"\n设置标注样式失败: {ex.Message}");
                            }
                        }

                        // 中文注释：提交事务
                        tr.Commit();
                        Env.Editor.Redraw();
                        insertSuccess = true;
                    }
                }

                if (deleteAfter && insertSuccess && !string.IsNullOrEmpty(sourceFilePath))
                {
                    try
                    {
                        if (System.IO.File.Exists(sourceFilePath))
                        {
                            System.IO.File.Delete(sourceFilePath);
                            LogManager.Instance.LogInfo($"\n已删除临时 DWG 文件: {sourceFilePath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogInfo($"\n尝试删除临时 DWG 失败: {ex.Message}");
                    }
                }

                VariableDictionary.winForm_Status = false;
            }
            catch (Exception ex)
            {
                failReason = ex.Message;
                LogManager.Instance.LogInfo($"\nCOPYDWGALLFAST 执行失败: {ex.Message}");
                LogManager.Instance.LogInfo($"\n堆栈跟踪: {ex.StackTrace}");
            }
            finally
            {
                _isCopyDwgAllFastDragging = false;
                ExitCopyDwgAllFastBusy();
                RaiseCopyDwgAllFastCompleted(insertSuccess, insertSuccess ? null : failReason);
            }
        }

        /// <summary>
        /// 焦点切换：尝试将焦点切换回 AutoCAD 的绘图区域，确保用户在插入块后能够直接看到并操作新插入的图元。
        /// </summary>
        private static void EnsureDwgViewFocus()
        {
            try
            {
                // 优先尝试 AutoCAD 内部焦点切换（用反射避免强依赖）
                var t1 = Type.GetType("Autodesk.AutoCAD.Internal.Utils, AcMgd", false);
                var m1 = t1?.GetMethod("SetFocusToDwgView", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                if (m1 != null)
                {
                    m1.Invoke(null, null);
                }
                else
                {
                    // 兼容部分版本程序集名
                    var t2 = Type.GetType("Autodesk.AutoCAD.Internal.Utils, AcCoreMgd", false);
                    var m2 = t2?.GetMethod("SetFocusToDwgView", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                    m2?.Invoke(null, null);
                }
            }
            catch
            {
                // 忽略焦点切换异常，避免中断主流程
            }

            try
            {
                System.Windows.Forms.Application.DoEvents();
            }
            catch { }

            try
            {
                Env.Editor.UpdateScreen();
            }
            catch { }
        }


        /// <summary>
        /// 插入外部条件图元
        /// </summary>
        [CommandMethod(nameof(GB_InsertBlock_Ptj))]
        public static void GB_InsertBlock_Ptj()
        {
            #region 方法2：
            try
            {
                Directory.CreateDirectory(GetPath.referenceFile);  //获取到本工具的系统目录；
                if (VariableDictionary.btnFileName == null) return; //判断点现的按键名是不是空；
                if (VariableDictionary.resourcesFile == null) return; //判断点现的原文件是不是空；
                using var tr = new DBTrans();
                var referenceFileObId = tr.BlockTable.GetBlockFormA(VariableDictionary.resourcesFile, VariableDictionary.btnFileName, true);//拿到本工具的系统目录下的按键名的原文件的objectid；
                var refFileRec = tr.GetObject(referenceFileObId, OpenMode.ForRead) as BlockTableRecord;//拿到原文件的块表记录；
                int isNo = -1;
                if (refFileRec != null)
                    foreach (var fileId in refFileRec)
                    {
                        isNo++;
                        if (VariableDictionary.btnFileName == "PTJ_给水点") isNo = -1;
                        else if (isNo != VariableDictionary.TCH_Ptj_No) continue;
                        //判读是不是为0，0就是天正元素
                        //if (fileId.ObjectClass.DxfName == "INSERT") continue;
                        var fileEntity = tr.GetObject(fileId, OpenMode.ForRead) as Entity;
                        //LogManager.Instance.LogInfo("PTJ元素！");
                        if (fileEntity == null) continue;
                        var fileEntityCopy = fileEntity.Clone() as Entity;
                        //tr.CurrentSpace.DeepCloneEx(fileEntity,)//深度克隆，可以复制天正图元；
                        //(vlax-dump-object (vlax-ename->vla-object (car (entsel )))T)//这个是在cad命令里能读出天正属性的lisp命令；
                        if (fileEntityCopy == null) continue;
                        var dxfName = fileId.ObjectClass.DxfName;//抓取图元的DXFName,判断是不是天正的图元
                        var fileType = dxfName.Split('_');//截取‘_’字符
                        if (fileType[0] == "TCH")
                        {
                            LogManager.Instance.LogInfo("TCH！");
                            var fileEntityCopyObId = tr.CurrentSpace.AddEntity(fileEntityCopy);
                            double tempAngle = 0;
                            var startPoint = new Point3d(0, 0, 0);
                            var entityBlock = new JigEx((mpw, _) =>
                            {
                                fileEntityCopy.Move(startPoint, mpw);
                                startPoint = mpw;
                                if (VariableDictionary.entityRotateAngle == tempAngle)
                                {
                                    return;
                                }
                                else if (VariableDictionary.entityRotateAngle != tempAngle)
                                {
                                    fileEntityCopy.Rotation(center: mpw, 0);
                                    tempAngle = VariableDictionary.entityRotateAngle;
                                    fileEntityCopy.Rotation(center: mpw, tempAngle);
                                }
                            });
                            entityBlock.DatabaseEntityDraw(wd => wd.Geometry.Draw(fileEntityCopy));
                            entityBlock.SetOptions(msg: "\n指定插入点");
                            //entityBlock.SetOptions(startPoint, msg: "\n指定插入点");这个startpoint，是有个参考线在里面，用于托拽时的辅助；
                            var endPoint = Env.Editor.Drag(entityBlock);
                            if (endPoint.Status != PromptStatus.OK) return;
                            tr.BlockTable.Remove(referenceFileObId);
                        }
                        else if (fileEntityCopy is BlockReference)
                        {
                            LogManager.Instance.LogInfo("PTJ-块表记录！");
                            //if (fileEntityCopy.ColorIndex.ToString() != "130") return;
                            var fileEntityCopyObId = tr.CurrentSpace.AddEntity(fileEntityCopy);//在当前图纸空间中加入这个实体并获取它的ObjoectId
                            double tempAngle = 0;
                            var startPoint = new Point3d(0, 0, 0);
                            var entityBlock = new JigEx((mpw, _) =>
                            {
                                fileEntityCopy.Move(startPoint, mpw);
                                startPoint = mpw;
                                if (VariableDictionary.entityRotateAngle == tempAngle)
                                {
                                    return;
                                }
                                else if (VariableDictionary.entityRotateAngle != tempAngle)
                                {
                                    fileEntityCopy.Rotation(center: mpw, 0);
                                    tempAngle = VariableDictionary.entityRotateAngle;
                                    fileEntityCopy.Rotation(center: mpw, tempAngle);
                                }
                            });
                            entityBlock.DatabaseEntityDraw(wd => wd.Geometry.Draw(fileEntityCopy));
                            entityBlock.SetOptions(msg: "\n指定插入点");
                            //entityBlock.SetOptions(startPoint, msg: "\n指定插入点");这个startpoint，是有个参考线在里面，用于托拽时的辅助；
                            var endPoint = Env.Editor.Drag(entityBlock);
                            if (endPoint.Status != PromptStatus.OK) return;
                            tr.BlockTable.Remove(referenceFileObId);
                        }
                        //else
                        //{
                        //    LogManager.Instance.LogInfo("PTJ-块！");
                        //    var referenceFileBlock = tr.CurrentSpace.InsertBlock(Point3d.Origin, referenceFileObId);
                        //    //tr.BlockTable.Remove(referenceFileObId);
                        //    if (tr.GetObject(referenceFileBlock) is not Entity referenceFileEntity) return;
                        //    double tempAngle = 0;
                        //    var startPoint = new Point3d(0, 0, 0);
                        //    var entityBlock = new JigEx((mpw, _) =>
                        //    {
                        //        referenceFileEntity.Move(startPoint, mpw);
                        //        startPoint = mpw;
                        //        if (VariableDictionary.entityRotateAngle == tempAngle)
                        //        {
                        //            return;
                        //        }
                        //        else if (VariableDictionary.entityRotateAngle != tempAngle)
                        //        {
                        //            referenceFileEntity.Rotation(center: mpw, 0);
                        //            tempAngle = VariableDictionary.entityRotateAngle;
                        //            referenceFileEntity.Rotation(center: mpw, tempAngle);
                        //        }
                        //    });
                        //    entityBlock.DatabaseEntityDraw(wd => wd.Geometry.Draw(referenceFileEntity));
                        //    entityBlock.SetOptions(msg: "\n指定插入点");
                        //    var endPoint = Env.Editor.Drag(entityBlock);
                        //    if (endPoint.Status != PromptStatus.OK) return;
                        //    referenceFileEntity.Layer = VariableDictionary.btnBlockLayer;
                        //    break;
                        //}
                    }

                tr.Commit();
                Env.Editor.Redraw();

            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("插入图元失败！");
                LogManager.Instance.LogInfo("错误信息: " + ex.Message);
            }
            #endregion
        }

        /// <summary>
        /// 一个块反复插入图中
        /// </summary>
        [CommandMethod(nameof(GB_InsertBlock_5))]
        public static void GB_InsertBlock_5()
        {
            #region 方法1：  
            try
            {
                pointS.Clear();
                Directory.CreateDirectory(GetPath.referenceFile);
                if (VariableDictionary.btnFileName == null) return;

                if (VariableDictionary.resourcesFile == null) return; //判断点现的原文件是不是空；
                using var tr = new DBTrans();

                // 获取对应块的 ObjectId  
                var referenceFileObId = tr.BlockTable.GetBlockFormA(
                    VariableDictionary.resourcesFile,
                    VariableDictionary.btnFileName,
                    VariableDictionary.btnFileName_blockName,
                    true);

                var refFileRec = tr.GetObject(referenceFileObId, OpenMode.ForRead) as BlockTableRecord;
                if (refFileRec == null)
                {
                    LogManager.Instance.LogInfo("未找到块记录！");
                    return;
                }

                LogManager.Instance.LogInfo("块！");
                while (true)
                {
                    // 把块插入到当前空间  
                    var referenceFileBlock = tr.CurrentSpace.InsertBlock(Point3d.Origin, referenceFileObId);

                    // 检查是否为实体  
                    if (tr.GetObject(referenceFileBlock) is not Entity referenceFileEntity)
                        return;

                    // 设置图层和颜色等属性  
                    referenceFileEntity.Layer = VariableDictionary.btnBlockLayer;
                    referenceFileEntity.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                    referenceFileEntity.Scale(new Point3d(0, 0, 0), VariableDictionary.blockScale);

                    //double tempAngle = 0; // 原始角度  
                    var startPoint = new Point3d(0, 0, 0);

                    var jigBlock = new JigEx((mpw, _) =>
                    {
                        // 先移动  
                        referenceFileEntity.Move(startPoint, mpw);
                        startPoint = mpw;
                    });
                    jigBlock.DatabaseEntityDraw(wd => wd.Geometry.Draw(referenceFileEntity));
                    jigBlock.SetOptions(msg: "\n指定插入点");

                    // 拖拽  
                    var endPoint = Env.Editor.Drag(jigBlock);
                    if (endPoint.Status != PromptStatus.OK)
                    {
                        // 用户取消插入，则删除已插入的块  
                        tr.GetObject(referenceFileBlock, OpenMode.ForWrite);
                        referenceFileBlock.Erase();
                        break;
                    }

                    // 存储插入点坐标（WCS）  
                    var UcsEndPoint = jigBlock.MousePointWcsLast;
                    pointS.Add(UcsEndPoint);


                    Env.Editor.Redraw();
                }

                // ======================  
                // 在此处根据插入数量绘图  
                // ======================  
                int count = pointS.Count;
                LogManager.Instance.LogInfo($"\n已插入 {count} 个块，开始绘制外围图形...");

                if (count == 3)
                { // 三点生成外接圆，圆心与3点等距  
                    Point3d p1 = pointS[0];
                    Point3d p2 = pointS[1];
                    Point3d p3 = pointS[2];

                    // 计算三角形外接圆圆心（与三点等距的点）  
                    Point3d circleCenter = GetCircumcenter(p1, p2, p3);

                    // 计算圆心到三个点的距离，取最大值，然后加上150作为新圆的半径  
                    double radius = p1.DistanceTo(circleCenter) + 150.0;

                    // 检查计算出的圆心是否与三点等距  
                    double dist1 = circleCenter.DistanceTo(p1);
                    double dist2 = circleCenter.DistanceTo(p2);
                    double dist3 = circleCenter.DistanceTo(p3);

                    // 记录到日志，以验证计算正确性  
                    LogManager.Instance.LogInfo($"\n圆心到三点的距离: {dist1:F4}, {dist2:F4}, {dist3:F4}");

                    // 正确创建圆：使用外接圆圆心和半径  
                    var circle = new Circle(circleCenter, Vector3d.ZAxis, radius);
                    circle.Layer = VariableDictionary.btnBlockLayer;
                    circle.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                    tr.CurrentSpace.AddEntity(circle);
                    Env.Editor.Redraw();
                    LogManager.Instance.LogInfo("\n已创建外围圆形，与三点等距并向外扩展150。");
                }
                else if (count == 4)
                {
                    // 计算中心点  
                    var center = new Point3d(
                        pointS.Average(p => p.X),
                        pointS.Average(p => p.Y),
                        pointS.Average(p => p.Z)
                    );

                    // 绘制矩形 - 使用原始4点作为矩形顶点，向外扩展150  
                    List<Point2d> expandedPoints = new List<Point2d>();

                    foreach (var point in pointS)
                    {
                        // 计算从中心到点的方向向量  
                        Vector3d dirVector = point - center;
                        dirVector = dirVector.GetNormal(); // 单位化向量  

                        // 创建新点：沿着方向向量延伸150的距离  
                        Point3d expandedPoint3d = point + dirVector * 150.0;
                        Point2d expandedPoint = new Point2d(expandedPoint3d.X, expandedPoint3d.Y);
                        expandedPoints.Add(expandedPoint);
                    }

                    // 确保点按顺时针或逆时针排序  
                    // 对顶点按角度排序  
                    var sortedPoints = expandedPoints.Select((p, index) => new
                    {
                        Point = p,
                        Angle = Math.Atan2(p.Y - center.Y, p.X - center.X)
                    })
                    .OrderBy(item => item.Angle)
                    .Select(item => item.Point)
                    .ToList();

                    // 创建Polyline并添加扩展后的顶点  
                    var pl = new Polyline();
                    for (int i = 0; i < sortedPoints.Count; i++)
                    {
                        pl.AddVertexAt(i, sortedPoints[i], 0, 30, 30);
                    }

                    // 闭合  
                    pl.Closed = true;
                    pl.Layer = VariableDictionary.btnBlockLayer;
                    pl.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                    tr.CurrentSpace.AddEntity(pl);
                    Env.Editor.Redraw();
                    LogManager.Instance.LogInfo("\n已创建外围矩形，向外扩展150。");
                }
                else if (count > 4)
                {
                    // 计算中心点  
                    var center = new Point3d(
                        pointS.Average(p => p.X),
                        pointS.Average(p => p.Y),
                        pointS.Average(p => p.Z)
                    );

                    // 创建多边形 - 使用原始点作为多边形顶点，向外扩展150  
                    List<Point2d> expandedPoints = new List<Point2d>();
                    foreach (var point in pointS)
                    {
                        // 计算从中心到点的方向向量  
                        Vector3d dirVector = point - center;
                        dirVector = dirVector.GetNormal(); // 单位化向量  
                        // 创建新点：沿着方向向量延伸150的距离  
                        Point3d expandedPoint3d = point + dirVector * 150.0;
                        Point2d expandedPoint = new Point2d(expandedPoint3d.X, expandedPoint3d.Y);
                        expandedPoints.Add(expandedPoint);
                    }
                    // 对顶点按角度排序，确保多边形正确  
                    var sortedPoints = expandedPoints.Select((p, index) => new
                    {
                        Point = p,
                        Angle = Math.Atan2(p.Y - center.Y, p.X - center.X)
                    })
                    .OrderBy(item => item.Angle)
                    .Select(item => item.Point)
                    .ToList();

                    // 创建多边形  
                    var polygon = new Polyline();
                    for (int i = 0; i < sortedPoints.Count; i++)
                    {
                        polygon.AddVertexAt(i, sortedPoints[i], 0, 30, 30);
                    }
                    // 闭合多边形  
                    polygon.Closed = true;
                    polygon.Layer = VariableDictionary.btnBlockLayer;
                    polygon.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                    tr.CurrentSpace.AddEntity(polygon);
                    Env.Editor.Redraw();
                    LogManager.Instance.LogInfo($"\n已创建{count}边形外围，向外扩展150。");
                }
                else if (count > 0)
                {
                    LogManager.Instance.LogInfo($"\n已插入{count}个块，但数量不满足绘制外围图形的条件（需要至少3个点）。");
                }
                //加标注
                // DDimLinear("总重:" + VariableDictionary.dimString + "kg" + "\n" + $"{count}点着地", Convert.ToInt16(pointS.Count));
                var dimColorLine = VariableDictionary.layerColorIndex;
                if (VariableDictionary.btnFileName.Contains("结构"))
                {
                    dimColorLine = 3;
                }
                if (VariableDictionary.dimString != null)
                    Command.DDimLinear(tr, VariableDictionary.dimString, count.ToString(), Convert.ToInt16(dimColorLine));
                tr.Commit();
                Env.Editor.Redraw();
                LogManager.Instance.LogInfo("\n操作完成。");
                pointS.Clear();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n插入图元失败：{ex.Message}");
                // 可以添加更详细的错误信息记录  
                LogManager.Instance.LogInfo($"\n错误详情：{ex.StackTrace}");
            }
            #endregion
        }

        /// <summary>
        /// 计算三角形外接圆圆心，确保圆心与三个点等距 
        /// </summary>
        /// <param name="A">A</param>
        /// <param name="B">B</param>
        /// <param name="C">C</param>
        /// <returns></returns>
        private static Point3d GetCircumcenter(Point3d A, Point3d B, Point3d C)
        {
            // 处理共线情况：如果三点共线，则返回三点的平均点  
            if (ArePointsCollinear(A, B, C))
            {
                return new Point3d(
                    (A.X + B.X + C.X) / 3.0,
                    (A.Y + B.Y + C.Y) / 3.0,
                    (A.Z + B.Z + C.Z) / 3.0
                );
            }
            // 计算分母 d  
            double d = 2 * (A.X * (B.Y - C.Y) + B.X * (C.Y - A.Y) + C.X * (A.Y - B.Y));
            if (Math.Abs(d) < 1e-10)
            {
                // 保护性返回：当分母太小时则视为共线，返回平均值  
                return new Point3d(
                    (A.X + B.X + C.X) / 3.0,
                    (A.Y + B.Y + C.Y) / 3.0,
                    (A.Z + B.Z + C.Z) / 3.0
                );
            }
            // 分别计算各点的 (x^2 + y^2)  
            double Asq = A.X * A.X + A.Y * A.Y;
            double Bsq = B.X * B.X + B.Y * B.Y;
            double Csq = C.X * C.X + C.Y * C.Y;
            // 使用标准公式计算圆心 X/Y 坐标  
            double centerX = (Asq * (B.Y - C.Y) + Bsq * (C.Y - A.Y) + Csq * (A.Y - B.Y)) / d;
            double centerY = (Asq * (C.X - B.X) + Bsq * (A.X - C.X) + Csq * (B.X - A.X)) / d;
            double centerZ = (A.Z + B.Z + C.Z) / 3.0; // Z 坐标取平均值  
            return new Point3d(centerX, centerY, centerZ);
        }

        /// <summary>
        /// 检查三点是否共线
        /// </summary>
        /// <param name="A">A</param>
        /// <param name="B">B</param>
        /// <param name="C">C</param>
        /// <returns></returns>

        private static bool ArePointsCollinear(Point3d A, Point3d B, Point3d C)
        {
            Vector3d v1 = B - A;
            Vector3d v2 = C - A;
            Vector3d crossProduct = v1.CrossProduct(v2);
            return crossProduct.Length < 1e-8;
        }
        #endregion
    }
}
