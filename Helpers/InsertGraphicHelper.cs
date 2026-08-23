using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Models;
using GB_NewCadPlus_IV.UniFiedStandards;
using GB_NewCadPlus_IV.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Interop;
using static GB_NewCadPlus_IV.Helpers.JsonHelper;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 插入图元的辅助方法（支持整图复制、属性继承、重叠判定等）
    /// </summary>
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
        /// 在图元正式炸开前显示属性编辑窗口，并返回用户确认后的属性。
        /// </summary>
        private static bool TryEditPropertiesBeforeInsert(
            DBTrans tr,
            BlockReference blockReference,
            string? title,
            LogManager logger,
            bool hasFlangeStandardMatch,
            FlangeStandardMatchResponse? flangeStandardResponse,
            IDictionary<string, string>? inheritedProperties,
            out Dictionary<string, string> editedProperties)
        {
            editedProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var propertyMap = ReadInsertEditablePropertyMap(tr, blockReference, logger);

            // 规范返回的字段可能在源图元中不存在，先补入编辑页面，保证 BOLT_HOLES 等字段可以参与计算。
            if (flangeStandardResponse?.Success == true && flangeStandardResponse.Attributes != null)
            {
                foreach (KeyValuePair<string, string> standardProperty in flangeStandardResponse.Attributes)
                {
                    if (string.IsNullOrWhiteSpace(standardProperty.Key)) continue;

                    // 法兰规范命中后，按归一化 Tag 覆盖编辑窗口中的同名旧值，避免 DWG 默认值再次覆盖服务器规范值。
                    string? existingKey = FindPropertyKeyByNormalizedKey(propertyMap, standardProperty.Key);
                    if (existingKey != null)
                    {
                        propertyMap[existingKey] = standardProperty.Value ?? string.Empty;
                        logger.LogInfo($"[规范属性合并][插入前窗口] Tag={standardProperty.Key}, 覆盖原键={existingKey}, NewValue={standardProperty.Value ?? string.Empty}");
                    }
                    else
                    {
                        propertyMap[standardProperty.Key] = standardProperty.Value ?? string.Empty;
                        logger.LogInfo($"[规范属性合并][插入前窗口] Tag={standardProperty.Key}, 原属性不存在，新增值={standardProperty.Value ?? string.Empty}");
                    }
                }

                logger.LogInfo($"规范返回属性已补充到插入前编辑页面：属性数量={flangeStandardResponse.Attributes.Count}");
            }

            // 只有本次确实命中法兰/连接规范时，才加入新增的三个业务属性。
            if (hasFlangeStandardMatch)
            {
                string boltHoles = FindProperty(propertyMap, "BOLT_HOLES") ?? string.Empty;
                // 当前块属性中可能没有连接方式，使用已确认的重叠继承属性作为计算兜底来源。
                string connectionType = FindProperty(propertyMap, "CONN_TYPE", "DNCONN_TYPE", "连接方式", "连接形式")
                    ?? FindProperty(inheritedProperties ?? new Dictionary<string, string>(), "CONN_TYPE", "DNCONN_TYPE", "连接方式", "连接形式")
                    ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(connectionType))
                {
                    // 将继承得到的连接方式补回编辑字典，确保页面显示和后续属性写回使用同一个值。
                    SetPropertyValueByNormalizedKey(propertyMap, "CONN_TYPE", connectionType);
                }
                int flangeQuantity = GetInsertFlangeQuantity(connectionType);
                // 螺栓数量按螺栓孔数量显示；法兰数量只单独记录在 FLG_QTY，不再参与 BOLT_QTY 计算。
                int boltQuantity = ParseIntegerOrZeroForInsert(boltHoles);
                SetPropertyValueByNormalizedKey(propertyMap, "FLG_QTY", flangeQuantity.ToString());
                SetPropertyValueByNormalizedKey(propertyMap, "BOLT_QTY", boltQuantity.ToString());
                SetPropertyValueByNormalizedKey(propertyMap, "BOLT_LENGTH", "0");
                logger.LogInfo($"插入前法兰扩展属性已加入：连接方式={connectionType}, FLG_QTY={FindProperty(propertyMap, "FLG_QTY") ?? string.Empty}, BOLT_HOLES={boltHoles}, BOLT_QTY={FindProperty(propertyMap, "BOLT_QTY") ?? string.Empty}, BOLT_LENGTH={FindProperty(propertyMap, "BOLT_LENGTH") ?? string.Empty}");
            }

            logger.LogInfo($"插入前属性编辑窗口准备打开：属性数量={propertyMap.Count}");
            try
            {
                var window = new InsertGraphicPropertyWindow(propertyMap, title);
                window.SourceInitialized += (_, _) =>
                {
                    try
                    {
                        new WindowInteropHelper(window) { Owner = Application.MainWindow.Handle };
                    }
                    catch (Exception ownerEx)
                    {
                        logger.LogWarning($"设置插入属性窗口宿主失败，将继续显示窗口：{ownerEx.Message}");
                    }
                };

                if (window.ShowDialog() != true)
                {
                    logger.LogInfo("用户在插入前属性编辑窗口中取消了插入。");
                    return false;
                }

                editedProperties = window.GetEditedProperties();
                logger.LogInfo($"用户确认插入图元：编辑后属性数量={editedProperties.Count}");
                return true;
            }
            catch (Exception ex)
            {
                logger.LogError($"显示插入前属性编辑窗口失败：{ex.Message}");
                return false;
            }
        }

        private static int ParseIntegerOrZeroForInsert(string value)
        {
            if (int.TryParse(value?.Trim(), out int result)) return Math.Max(0, result);
            if (double.TryParse(value?.Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double number))
            {
                return Math.Max(0, (int)Math.Round(number));
            }
            return 0;
        }

        /// <summary>
        /// 按归一化 Tag 查找属性字典中的原始键，兼容 CONN_TYPE/CONNTYPE、FLG_STD/FLGSTD 等写法。
        /// </summary>
        private static string? FindPropertyKeyByNormalizedKey(
            IDictionary<string, string> properties,
            string key)
        {
            string normalizedKey = NormalizePropertyKey(key);
            if (string.IsNullOrWhiteSpace(normalizedKey)) return null;

            foreach (string existingKey in properties.Keys)
            {
                if (string.Equals(NormalizePropertyKey(existingKey), normalizedKey, StringComparison.OrdinalIgnoreCase))
                    return existingKey;
            }

            return null;
        }

        /// <summary>
        /// 按归一化 Tag 更新已有属性，找不到时才新增规范键，避免产生同名重复字段。
        /// </summary>
        private static void SetPropertyValueByNormalizedKey(
            IDictionary<string, string> properties,
            string key,
            string value)
        {
            string? existingKey = FindPropertyKeyByNormalizedKey(properties, key);
            properties[existingKey ?? key] = value ?? string.Empty;
        }

        /// <summary>
        /// 根据插入前连接方式计算法兰数量。
        /// </summary>
        private static int GetInsertFlangeQuantity(string connectionType)
        {
            // 法兰类图元插入时默认按一个法兰计数，用户可在插入前属性页面中修改该值。
            string value = connectionType?.Trim() ?? string.Empty;
            if (value.Contains("法兰", StringComparison.OrdinalIgnoreCase) ||
                value.Contains("对夹", StringComparison.OrdinalIgnoreCase)) return 1;
            return 0;
        }

        /// <summary>
        /// 将用户编辑后的属性写回块属性和 XRecord，允许用户将属性清空。
        /// </summary>
        private static void ApplyEditedPropertiesToEntity(
            DBTrans tr,
            Entity entity,
            IDictionary<string, string> editedProperties,
            LogManager logger)
        {
            // 参数无效时不执行任何数据库写入。
            if (tr == null || entity == null || editedProperties == null || editedProperties.Count == 0) return;

            // 块属性必须按实际 AttributeReference.Tag 匹配，不能只按字典键名匹配。
            if (entity is BlockReference blockReference)
            {
                foreach (ObjectId attributeId in blockReference.AttributeCollection)
                {
                    if (tr.GetObject(attributeId, OpenMode.ForWrite) is not AttributeReference attribute) continue;

                    string tag = (attribute.Tag ?? string.Empty).Trim();
                    string decodedTag = PipelineCadPropertyKeyHelper.Decode(tag);
                    if (!TryGetEditedValue(editedProperties, tag, decodedTag, out string newValue)) continue;

                    string oldValue = attribute.TextString ?? string.Empty;
                    if (string.Equals(oldValue, newValue, StringComparison.Ordinal)) continue;

                    attribute.TextString = newValue;
                    logger.LogInfo($"[插入前属性编辑赋值][AttributeReference] Tag={tag}, OldValue={oldValue}, NewValue={newValue}, TargetObjectId={entity.ObjectId}");
                }
            }

            // 先创建不存在的属性记录，确保新增的 FLG_QTY、BOLT_QTY、BOLT_LENGTH 不会因源图元没有同名属性而丢失。
            EnsureEditedPropertiesInXRecord(tr, entity, editedProperties, logger);

            // XRecord 继续更新已有记录，保留原有字段类型和存储结构。
            if (entity.ExtensionDictionary == ObjectId.Null) return;
            if (tr.GetObject(entity.ExtensionDictionary, OpenMode.ForWrite) is not DBDictionary dictionary) return;

            foreach (DBDictionaryEntry entry in dictionary)
            {
                if (tr.GetObject(entry.Value, OpenMode.ForWrite) is not Xrecord record) continue;
                TypedValue[] values = record.Data?.AsArray() ?? Array.Empty<TypedValue>();
                if (values.Length == 0) continue;

                string key = PipelineCadPropertyKeyHelper.Decode((entry.Key ?? string.Empty).Trim());
                if (string.Equals(entry.Key, PipelineCadPropertyKeyHelper.StorageKey, StringComparison.OrdinalIgnoreCase))
                {
                    // 管道专用 XRecord 按“键、值”交替存储，偶数项为键，奇数项为值。
                    var updatedValues = values.ToArray();
                    bool changed = false;
                    for (int index = 0; index + 1 < updatedValues.Length; index += 2)
                    {
                        string itemKey = updatedValues[index].Value?.ToString()?.Trim() ?? string.Empty;
                        if (!TryGetEditedValue(editedProperties, itemKey, PipelineCadPropertyKeyHelper.Decode(itemKey), out string newValue)) continue;
                        string oldValue = updatedValues[index + 1].Value?.ToString() ?? string.Empty;
                        if (string.Equals(oldValue, newValue, StringComparison.Ordinal)) continue;
                        updatedValues[index + 1] = new TypedValue(updatedValues[index + 1].TypeCode, ConvertValueByTypeCode(updatedValues[index + 1].TypeCode, newValue));
                        changed = true;
                        logger.LogInfo($"[插入前属性编辑赋值][XRecord] Tag={itemKey}, OldValue={oldValue}, NewValue={newValue}, TargetObjectId={entity.ObjectId}");
                    }
                    if (changed) record.Data = new ResultBuffer(updatedValues);
                    continue;
                }

                if (!TryGetEditedValue(editedProperties, key, entry.Key, out string editedValue)) continue;
                string oldText = values[0].Value?.ToString() ?? string.Empty;
                if (string.Equals(oldText, editedValue, StringComparison.Ordinal)) continue;
                // 重新组合首项和其余 TypedValue，保留原 XRecord 的附加数据项。
                var rewrittenValues = new[]
                {
                    new TypedValue(values[0].TypeCode, ConvertValueByTypeCode(values[0].TypeCode, editedValue))
                }
                .Concat(values.Skip(1))
                .ToArray();
                record.Data = new ResultBuffer(rewrittenValues);
                logger.LogInfo($"[插入前属性编辑赋值][XRecord] Tag={key}, OldValue={oldText}, NewValue={editedValue}, TargetObjectId={entity.ObjectId}");
            }
        }

        /// <summary>
        /// 按原始键、解码键及归一化键查找用户编辑值。
        /// </summary>
        private static bool TryGetEditedValue(
            IDictionary<string, string> properties,
            string firstKey,
            string secondKey,
            out string value)
        {
            if (properties.TryGetValue(firstKey, out value!)) return true;
            if (!string.IsNullOrWhiteSpace(secondKey) && properties.TryGetValue(secondKey, out value!)) return true;
            string normalized = NormalizePropertyKey(firstKey);
            if (!string.IsNullOrWhiteSpace(normalized) && properties.TryGetValue(normalized, out value!)) return true;

            // 编辑窗口只保留一个规范 Tag，因此需要按归一化结果反向匹配原始 Tag。
            foreach (var property in properties)
            {
                if (string.Equals(NormalizePropertyKey(property.Key), normalized, StringComparison.OrdinalIgnoreCase))
                {
                    value = property.Value ?? string.Empty;
                    return true;
                }
            }

            value = string.Empty;
            return false;
        }

        // 为实体扩展字典补充编辑窗口中的缺失属性，确保新增字段不会因源图元没有同名属性而丢失。
        private static void EnsureEditedPropertiesInXRecord(
            DBTrans tr,
            Entity entity,
            IDictionary<string, string> editedProperties,
            LogManager logger)
        {
            if (tr == null || entity == null || editedProperties == null || editedProperties.Count == 0) return;

            var existingKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var attributeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (entity is BlockReference blockReference)
            {
                foreach (ObjectId attributeId in blockReference.AttributeCollection)
                {
                    if (tr.GetObject(attributeId, OpenMode.ForRead) is AttributeReference attribute)
                    {
                        string tag = PipelineCadPropertyKeyHelper.Decode((attribute.Tag ?? string.Empty).Trim());
                        if (!string.IsNullOrWhiteSpace(tag))
                        {
                            string normalizedTag = NormalizePropertyKey(tag);
                            attributeKeys.Add(normalizedTag);
                            existingKeys.Add(normalizedTag);
                        }
                    }
                }
            }

            if (entity.ExtensionDictionary == ObjectId.Null) entity.CreateExtensionDictionary();
            if (tr.GetObject(entity.ExtensionDictionary, OpenMode.ForWrite) is not DBDictionary dictionary) return;

            Xrecord? storageRecord = null;
            foreach (DBDictionaryEntry entry in dictionary)
            {
                if (string.Equals(entry.Key, PipelineCadPropertyKeyHelper.StorageKey, StringComparison.OrdinalIgnoreCase))
                {
                    storageRecord = tr.GetObject(entry.Value, OpenMode.ForWrite) as Xrecord;
                    continue;
                }

                string key = PipelineCadPropertyKeyHelper.Decode((entry.Key ?? string.Empty).Trim());
                if (!string.IsNullOrWhiteSpace(key)) existingKeys.Add(NormalizePropertyKey(key));
            }

            var storageValues = storageRecord?.Data?.AsArray()?.ToList() ?? new List<TypedValue>();
            for (int index = 0; index + 1 < storageValues.Count; index += 2)
            {
                string key = storageValues[index].Value?.ToString()?.Trim() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(key)) existingKeys.Add(NormalizePropertyKey(key));
            }

            foreach (KeyValuePair<string, string> property in editedProperties)
            {
                string key = PipelineCadPropertyKeyHelper.Decode((property.Key ?? string.Empty).Trim());
                if (string.IsNullOrWhiteSpace(key) || !existingKeys.Add(NormalizePropertyKey(key))) continue;

                storageValues.Add(new TypedValue((int)DxfCode.Text, key));
                storageValues.Add(new TypedValue((int)DxfCode.Text, property.Value ?? string.Empty));
                logger.LogInfo($"[插入前属性编辑新增赋值][XRecord] Tag={key}, OldValue=<不存在>, NewValue={property.Value ?? string.Empty}, TargetObjectId={entity.ObjectId}");
            }

            // 统一存储记录保存全部确认后的字段，包含空值和 0，避免业务字段被过滤掉。
            if (storageValues.Count > 0 && storageRecord == null)
            {
                storageRecord = new Xrecord { Data = new ResultBuffer(storageValues.ToArray()) };
                dictionary.SetAt(PipelineCadPropertyKeyHelper.StorageKey, storageRecord);
                tr.Transaction.AddNewlyCreatedDBObject(storageRecord, true);
            }
            else if (storageRecord != null && storageValues.Count > 0)
            {
                storageRecord.Data = new ResultBuffer(storageValues.ToArray());
            }

            // 为没有 AttributeReference 的字段额外建立独立 XRecord，兼容只按字典键读取属性的旧逻辑。
            foreach (KeyValuePair<string, string> property in editedProperties)
            {
                string key = PipelineCadPropertyKeyHelper.Decode((property.Key ?? string.Empty).Trim());
                string normalizedKey = NormalizePropertyKey(key);
                if (string.IsNullOrWhiteSpace(key) ||
                    string.Equals(key, PipelineCadPropertyKeyHelper.StorageKey, StringComparison.OrdinalIgnoreCase) ||
                    attributeKeys.Contains(normalizedKey)) continue;

                string encodedKey = PipelineCadPropertyKeyHelper.Encode(key);
                if (dictionary.Contains(encodedKey)) continue;

                var propertyRecord = new Xrecord
                {
                    Data = new ResultBuffer(new TypedValue(
                        (int)DxfCode.Text,
                        property.Value ?? string.Empty))
                };
                dictionary.SetAt(encodedKey, propertyRecord);
                tr.Transaction.AddNewlyCreatedDBObject(propertyRecord, true);
                logger.LogInfo(
                    $"[插入前属性编辑新增赋值][独立XRecord] Tag={key}, OldValue=<不存在>, NewValue={property.Value ?? string.Empty}, TargetObjectId={entity.ObjectId}");
            }
        }

        // 按 XRecord 原字段类型转换编辑后的文本，避免修改属性时破坏 TypedValue 类型。
        private static object ConvertValueByTypeCode(int typeCode, string value)
        {
            string text = value ?? string.Empty;
            if (typeCode == 70 || typeCode == 71 || typeCode == 72 || typeCode == 73)
            {
                return short.TryParse(text, out short shortValue) ? shortValue : (short)0;
            }
            if (typeCode == 90 || typeCode == 91 || typeCode == 92 || typeCode == 93)
            {
                return int.TryParse(text, out int intValue) ? intValue : 0;
            }
            if (typeCode == 40 || typeCode == 41 || typeCode == 42 || typeCode == 43)
            {
                return double.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double doubleValue)
                    ? doubleValue
                    : 0.0;
            }
            return text;
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
        internal static bool TryGetEntityExtents(Entity entity, out Extents3d extents)
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
        internal static bool IsExtentsIntersect(Extents3d a, Extents3d b, double tol = 1e-6)
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

            return IsEntityOverlap(source, target, e1, e2);
        }

        /// <summary>
        /// 使用已缓存包围盒判断两个实体是否重叠，避免重复读取几何包围盒
        /// </summary>
        internal static bool IsEntityOverlap(Entity source, Entity target, Extents3d sourceExtents, Extents3d targetExtents)
        {
            // 包围盒不相交直接返回
            if (!IsExtentsIntersect(sourceExtents, targetExtents)) return false;

            // 插入块与管道曲线不能只使用包围盒判断，否则旋转块的外接矩形会覆盖附近无关管道。
            if (source is BlockReference sourceBlock && target is Curve targetCurve)
            {
                // 只有明确计算出插入块实际几何与管道的关系后，才认定二者重叠。
                if (TryGetBlockCurveRelation(sourceBlock, targetCurve, out _, out bool hasCurveGeometry))
                    return true;

                // 炸解后存在曲线但没有相交时，必须排除该候选，不能退回包围盒结论。
                if (hasCurveGeometry) return false;
            }

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
        /// 计算插入块实际几何与候选管道曲线的关系。
        /// </summary>
        /// <param name="block">当前已经按最终位置、旋转和比例变换的块参照</param>
        /// <param name="pipeCurve">当前候选管道曲线</param>
        /// <param name="distance">实际几何到管道的最小近似距离</param>
        /// <param name="hasCurveGeometry">块炸解结果中是否存在可计算曲线</param>
        /// <returns>实际几何与管道相交或在允许精度内接近时返回 true</returns>
        internal static bool TryGetBlockCurveRelation(
            BlockReference block,
            Curve pipeCurve,
            out double distance,
            out bool hasCurveGeometry)
        {
            // 初始化输出值，避免代理实体异常时返回未定义结果。
            distance = double.MaxValue;
            hasCurveGeometry = false;

            // 参数无效时无法进行精确判断。
            if (block == null || pipeCurve == null) return false;

            DBObjectCollection exploded = new DBObjectCollection();
            try
            {
                // Explode 使用块当前的 BlockTransform，得到最终落图位置的实际几何。
                block.Explode(exploded);

                // 逐个检查块内实体，避免使用包含旋转误差的整体外接包围盒。
                foreach (DBObject dbObject in exploded)
                {
                    if (!(dbObject is Entity entity)) continue;

                    // 嵌套块继续炸解一层，兼容管道附件中套块的情况。
                    if (entity is BlockReference nestedBlock)
                    {
                        DBObjectCollection nestedExploded = new DBObjectCollection();
                        try
                        {
                            nestedBlock.Explode(nestedExploded);
                            foreach (DBObject nestedObject in nestedExploded)
                            {
                                if (nestedObject is Curve nestedCurve)
                                {
                                    hasCurveGeometry = true;
                                    if (TryMeasureCurveRelation(nestedCurve, pipeCurve, ref distance))
                                        return true;
                                }
                                else
                                {
                                    nestedObject.Dispose();
                                }
                            }
                        }
                        finally
                        {
                            nestedBlock.Dispose();
                        }

                        continue;
                    }

                    // 管道附件通常由 Line/Polyline/Arc 等曲线组成。
                    if (entity is Curve curve)
                    {
                        hasCurveGeometry = true;
                        if (TryMeasureCurveRelation(curve, pipeCurve, ref distance))
                            return true;
                    }
                    else
                    {
                        // 非曲线实体不参与精确曲线判定，但仍然释放临时对象。
                        entity.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                // 代理实体可能不支持炸解，记录原因并由调用方决定是否使用安全回退。
                LogManager.Instance.LogWarning(
                    $"插入块精确交叠判定失败：BlockObjectId={block.ObjectId}, PipeObjectId={pipeCurve.ObjectId}, 原因={ex.Message}");
            }
            finally
            {
                // 清理所有尚未提前返回的临时炸解对象。
                foreach (DBObject dbObject in exploded)
                {
                    try { dbObject.Dispose(); } catch { }
                }
            }

            return false;
        }

        /// <summary>
        /// 计算两条曲线是否相交，并用端点最近距离作为未相交时的安全近似值。
        /// </summary>
        private static bool TryMeasureCurveRelation(Curve sourceCurve, Curve targetCurve, ref double distance)
        {
            try
            {
                // 真实相交优先级最高，不受曲线方向和角度影响。
                var points = new Point3dCollection();
                sourceCurve.IntersectWith(targetCurve, Intersect.OnBothOperands, points, IntPtr.Zero, IntPtr.Zero);
                if (points.Count > 0)
                {
                    distance = 0.0;
                    return true;
                }

                // 未相交时从两条曲线端点互相求最近点，避免使用块基点代替实际图形位置。
                double current = double.MaxValue;
                foreach (Point3d point in new[] { sourceCurve.StartPoint, sourceCurve.EndPoint })
                {
                    current = Math.Min(current, targetCurve.GetClosestPointTo(point, false).DistanceTo(point));
                }

                foreach (Point3d point in new[] { targetCurve.StartPoint, targetCurve.EndPoint })
                {
                    current = Math.Min(current, sourceCurve.GetClosestPointTo(point, false).DistanceTo(point));
                }

                distance = Math.Min(distance, current);
            }
            catch
            {
                // 单个异常曲线不影响其他块内曲线继续判定。
            }

            return false;
        }

        /// <summary>
        /// 计算两个包围盒交叠体积（用于候选评分）
        /// </summary>
        private static double CalcOverlapVolume(Extents3d a, Extents3d b)
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
        private static Point3d GetExtentsCenter(Extents3d e)
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
        private static bool ShouldSkipInheritedValue(string raw)
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
        private static string GetEntityTypeToken(Entity e)
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

            return ComputeOverlapCandidateScore(insertingBr, candidate, srcExt, dstExt);
        }

        /// <summary>
        /// 使用已缓存包围盒计算候选分数，避免重复读取几何包围盒
        /// </summary>
        internal static double ComputeOverlapCandidateScore(BlockReference insertingBr, Entity candidate, Extents3d srcExt, Extents3d dstExt)
        {
            // 空对象直接最低分
            if (insertingBr == null || candidate == null) return double.MinValue;

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

            // 优先使用插入块实际几何到候选管道的距离，而不是使用块基点到管道的距离。
            double referenceDistance = GetDistanceToReferencePoint(candidate, insertingBr.Position);
            if (candidate is Curve candidateCurve &&
                TryGetBlockCurveRelation(insertingBr, candidateCurve, out double geometryDistance, out bool hasCurveGeometry) &&
                hasCurveGeometry)
            {
                // 精确相交时该距离为 0，保证真正被插入图元压住的管道优先级最高。
                referenceDistance = geometryDistance;
            }
            double referencePointScore = 1.0 / (1.0 + referenceDistance);

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

            // 组合总分：捕捉点接近度优先，包围盒和中心距离作为辅助，避免多条管道时误选。
            double score = referencePointScore * 0.60 +
                           overlapRatio * 0.15 +
                           distanceScore * 0.10 +
                           layerScore * 0.10 +
                           typeScore * 0.05;

            // 返回最终评分
            return score;
        }

        /// <summary>
        /// 计算候选实体到插入基点的距离。
        /// </summary>
        public static double GetDistanceToReferencePoint(Entity candidate, Point3d referencePoint)
        {
            // 无效候选返回最大距离，避免被误判为最佳候选。
            if (candidate == null || candidate.IsErased) return double.MaxValue;

            try
            {
                // 管道通常是 Line、Polyline 或其他 Curve，使用曲线最近点进行精確距离判断。
                if (candidate is Curve curve)
                {
                    Point3d closestPoint = curve.GetClosestPointTo(referencePoint, false);
                    return closestPoint.DistanceTo(referencePoint);
                }

                // 非曲线实体使用包围盒中心作为兜底距离。
                if (TryGetEntityExtents(candidate, out Extents3d extents))
                {
                    return GetExtentsCenter(extents).DistanceTo(referencePoint);
                }
            }
            catch
            {
                // 个别代理实体可能无法求最近点，继续使用最大距离兜底。
            }

            return double.MaxValue;
        }

        /// <summary>
        /// 判断字段是否允许参与继承（白名单优先 + 黑名单兜底）
        /// </summary>
        private static bool IsPropertyKeyAllowed(string rawKey, HashSet<string> activeWhitelist)
        {
            // 空键直接不允许，避免脏数据进入
            if (string.IsNullOrWhiteSpace(rawKey)) return false;

            // 黑名单先拦截（绝对禁止）
            if (JsonHelper.IsBlacklistedPropertyKey(rawKey)) return false;

            // 法兰匹配所需的关键字段必须跨专业白名单保留，否则部分蝶阀只能继承到普通属性而无法查询规范。
            if (IsFlangeQueryPropertyKey(rawKey)) return true;

            // 未启用白名单时，黑名单外都允许
            if (!JsonHelper._propertySyncUseWhitelistTemplate) return true;

            // 白名单为空时，不允许任何字段（安全兜底）
            if (activeWhitelist == null || activeWhitelist.Count == 0) return false;

            // 归一化后比对白名单
            string nKey = NormalizePropertyKey(rawKey);
            if (string.IsNullOrWhiteSpace(nKey)) return false;

            // 命中白名单才允许
            return activeWhitelist.Contains(nKey);
        }

        /// <summary>
        /// 判断属性是否是法兰规范查询必须继承的关键字段。
        /// </summary>
        private static bool IsFlangeQueryPropertyKey(string rawKey)
        {
            string normalizedKey = NormalizePropertyKey(rawKey);
            return normalizedKey == NormalizePropertyKey("DN") ||
                   normalizedKey == NormalizePropertyKey("PN") ||
                   normalizedKey == NormalizePropertyKey("DNCONN_TYPE") ||
                   normalizedKey == NormalizePropertyKey("CONN_TYPE") ||
                   normalizedKey == NormalizePropertyKey("CONNECTION_MODE") ||
                   normalizedKey == NormalizePropertyKey("连接方式") ||
                   normalizedKey == NormalizePropertyKey("连接形式") ||
                   normalizedKey == NormalizePropertyKey("FLG_STD") ||
                   normalizedKey == NormalizePropertyKey("DRAWINGNO.STANDARDNO") ||
                   normalizedKey == NormalizePropertyKey("SERIES") ||
                   normalizedKey == NormalizePropertyKey("FLG_TYPE") ||
                   normalizedKey == NormalizePropertyKey("FACE_TYPE");
        }

        /// <summary>
        /// 读取实体属性映射（支持：块属性 + 扩展字典XRecord）
        /// </summary>
        private static Dictionary<string, string> ReadEntityPropertyMap(
            DBTrans tr,
            Entity entity,
            bool skipEmptyOrZero = true)
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

                    // 读取标签名并清洗；新管道属性载体使用编码后的 CAD Tag，需要还原业务字段名
                    string tag = PipelineCadPropertyKeyHelper.Decode((ar.Tag ?? string.Empty).Trim());
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

                        // 新管道将所有属性按“键、值”交替写入 GBPIPE_DATA，不能把首项当作整条记录的值。
                        if (string.Equals(
                            entry.Key,
                            PipelineCadPropertyKeyHelper.StorageKey,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            for (int index = 0; index + 1 < values.Length; index += 2)
                            {
                                string key = values[index].Value?.ToString()?.Trim() ?? string.Empty;
                                string val = values[index + 1].Value?.ToString() ?? string.Empty;
                                AddPropertyToMap(map, key, val, skipEmptyOrZero);
                            }
                        }
                        else
                        {
                            // 历史/普通 XRecord 仍使用“字典键作为属性名、首项作为属性值”。
                            string key = PipelineCadPropertyKeyHelper.Decode((entry.Key ?? string.Empty).Trim());
                            string val = values[0].Value?.ToString() ?? string.Empty;
                            AddPropertyToMap(map, key, val, skipEmptyOrZero);
                        }
                    }
                }
            }

            // 返回属性映射结果
            return map;
        }

        /// <summary>
        /// 读取插入前编辑窗口需要显示的完整属性。
        /// </summary>
        private static Dictionary<string, string> ReadInsertEditablePropertyMap(
            DBTrans tr,
            Entity entity,
            LogManager logger)
        {
            // 先保留原有读取结果，保证重叠继承和规范回写产生的属性仍然显示。
            // 插入前编辑页面必须保留 0 和空值，否则用户确认后的完整字段会再次从页面中消失。
            var map = ReadEntityPropertyMap(tr, entity, skipEmptyOrZero: false);
            var visitedBlockDefinitions = new HashSet<ObjectId>();

            // 递归扫描当前块、块定义和嵌套块，补充源 DWG 中实际保存的属性。
            CollectInsertEditableProperties(tr, entity, map, visitedBlockDefinitions, 0);
            // 窗口只显示一个业务 Tag，避免 TAG_NO/TAGNO 这类匹配别名重复出现。
            map = CollapseInsertPropertyAliases(map);
            logger.LogInfo($"插入前属性完整收集完成：属性数量={map.Count}");
            return map;
        }

        /// <summary>
        /// 折叠插入编辑窗口中的归一化别名，只保留一个最有业务含义的 Tag。
        /// </summary>
        private static Dictionary<string, string> CollapseInsertPropertyAliases(
            IDictionary<string, string> source)
        {
            // 使用不区分大小写的字典，保证大小写不同的 Tag 也不会重复显示。
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (source == null || source.Count == 0) return result;

            // 按归一化键分组，例如 TAG_NO 和 TAGNO 会进入同一组。
            foreach (var group in source
                .Where(item => !string.IsNullOrWhiteSpace(item.Key))
                .GroupBy(item => NormalizePropertyKey(item.Key), StringComparer.OrdinalIgnoreCase))
            {
                // 优先保留下划线、点号等业务分隔符更完整的原始 Tag。
                var selected = group
                    .OrderByDescending(item => CountPropertySeparators(item.Key))
                    .ThenBy(item => item.Key.Length)
                    .First();

                string selectedValue = selected.Value ?? string.Empty;

                // 归一化键对应的值如果只是空值或 0，则优先采用同组中更有实际内容的值。
                if (IsEmptyOrZeroPropertyValue(selectedValue))
                {
                    var meaningful = group.FirstOrDefault(item => !IsEmptyOrZeroPropertyValue(item.Value));
                    if (!string.IsNullOrWhiteSpace(meaningful.Key))
                    {
                        selectedValue = meaningful.Value ?? string.Empty;
                    }
                }

                result[selected.Key.Trim()] = selectedValue;
            }

            return result;
        }

        /// <summary>
        /// 统计 Tag 中用于区分业务字段的分隔符数量。
        /// </summary>
        private static int CountPropertySeparators(string key)
        {
            // 下划线和点号是当前项目中最常见的业务 Tag 分隔符。
            return (key ?? string.Empty).Count(character => character == '_' || character == '.');
        }

        /// <summary>
        /// 判断属性值是否为空或只是占位数值 0。
        /// </summary>
        private static bool IsEmptyOrZeroPropertyValue(string value)
        {
            // 只把空字符串和纯数字 0 当作占位值，不影响正常文本属性。
            return string.IsNullOrWhiteSpace(value) || string.Equals(value.Trim(), "0", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 递归收集块参照、块定义及嵌套块中的属性。
        /// </summary>
        private static void CollectInsertEditableProperties(
            DBTrans tr,
            Entity entity,
            Dictionary<string, string> map,
            HashSet<ObjectId> visitedBlockDefinitions,
            int depth)
        {
            // 防止异常 DWG 的循环块引用造成无限递归。
            if (tr == null || entity == null || map == null || visitedBlockDefinitions == null || depth > 20) return;

            // 先读取当前实体直接拥有的 AttributeReference，即使值为空也要显示给用户编辑。
            if (entity is BlockReference blockReference)
            {
                foreach (ObjectId attributeId in blockReference.AttributeCollection)
                {
                    if (tr.GetObject(attributeId, OpenMode.ForRead) is not AttributeReference attribute) continue;
                    AddEditableProperty(map, attribute.Tag, attribute.TextString);
                }

                // 动态块属性也属于用户可编辑属性，读取其名称和值。
                try
                {
                    foreach (DynamicBlockReferenceProperty property in blockReference.DynamicBlockReferencePropertyCollection)
                    {
                        if (property == null || string.IsNullOrWhiteSpace(property.PropertyName)) continue;
                        AddEditableProperty(map, property.PropertyName, property.Value?.ToString());
                    }
                }
                catch
                {
                    // 部分普通块不支持动态属性，读取失败不影响普通属性显示。
                }

                // 打开当前块定义并递归检查其 AttributeDefinition、嵌套块和扩展属性。
                ObjectId blockDefinitionId = blockReference.BlockTableRecord;
                if (!visitedBlockDefinitions.Add(blockDefinitionId)) return;
                if (tr.GetObject(blockDefinitionId, OpenMode.ForRead) is not BlockTableRecord blockDefinition) return;

                foreach (ObjectId childId in blockDefinition)
                {
                    if (tr.GetObject(childId, OpenMode.ForRead) is AttributeDefinition definition)
                    {
                        AddEditableProperty(map, definition.Tag, definition.TextString);
                        continue;
                    }

                    if (tr.GetObject(childId, OpenMode.ForRead) is Entity childEntity)
                    {
                        CollectInsertEditableProperties(tr, childEntity, map, visitedBlockDefinitions, depth + 1);
                    }
                }
            }
        }

        /// <summary>
        /// 将属性加入编辑字典，不过滤空字符串，以便窗口能够显示空属性字段。
        /// </summary>
        private static void AddEditableProperty(
            Dictionary<string, string> map,
            string? rawKey,
            string? rawValue)
        {
            // 清理 CAD Tag 并还原管道属性编码。
            string key = PipelineCadPropertyKeyHelper.Decode((rawKey ?? string.Empty).Trim());
            if (string.IsNullOrWhiteSpace(key)) return;

            // 保存原始业务键和值。
            map[key] = rawValue ?? string.Empty;

            // 同时保存归一化键，兼容后续回写的大小写和特殊字符匹配。
            string normalizedKey = NormalizePropertyKey(key);
            if (!string.IsNullOrWhiteSpace(normalizedKey) && !map.ContainsKey(normalizedKey))
            {
                map[normalizedKey] = rawValue ?? string.Empty;
            }
        }

        /// <summary>
        /// 将属性加入映射，同时保留原始键和归一化键，统一处理空值过滤。
        /// </summary>
        private static void AddPropertyToMap(
            Dictionary<string, string> map,
            string key,
            string value,
            bool skipEmptyOrZero = true)
        {
            if (map == null || string.IsNullOrWhiteSpace(key)) return;
            if (skipEmptyOrZero && ShouldSkipInheritedValue(value)) return;

            map[key] = value ?? string.Empty;
            string normalizedKey = NormalizePropertyKey(key);
            if (!string.IsNullOrWhiteSpace(normalizedKey) && !map.ContainsKey(normalizedKey))
            {
                map[normalizedKey] = value ?? string.Empty;
            }
        }

        /// <summary>
        /// 将源属性同步到目标块参照（仅同名属性，增强键名匹配）
        /// </summary>
        private static void SyncCommonPropertiesToBlockReference(DBTrans tr, BlockReference targetBr, Dictionary<string, string> sourceMap)
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
                string oldValue = ar.TextString ?? string.Empty;
                if (string.Equals(oldValue, newValue, StringComparison.Ordinal)) continue;

                // 执行覆盖写入
                ar.TextString = newValue;

                // 只记录实际发生的属性赋值，便于核对插入图元的 Tag 和最终值。
                LogManager.Instance.LogInfo(
                    $"[属性继承赋值][AttributeReference] Tag={tag}, OldValue={oldValue}, NewValue={newValue}, TargetObjectId={targetBr.ObjectId}");
            }
        }

        /// <summary>
        /// 将源属性同步到目标实体扩展字典（仅同名键，增强：尽量保持原类型）
        /// </summary>
        private static void SyncCommonPropertiesToEntityXRecord(DBTrans tr, Entity targetEntity, Dictionary<string, string> sourceMap)
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

                    // 记录首次创建 XRecord 数据时的实际赋值。
                    LogManager.Instance.LogInfo(
                        $"[属性继承赋值][XRecord] Tag={key}, OldValue=, NewValue={val ?? string.Empty}, TargetObjectId={targetEntity.ObjectId}");
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

                // 记录 XRecord 实际发生的标题和值变更。
                LogManager.Instance.LogInfo(
                    $"[属性继承赋值][XRecord] Tag={key}, OldValue={oldText}, NewValue={newText}, TargetObjectId={targetEntity.ObjectId}");
            }
        }

        /// <summary>
        /// 对“分解后的新实体集合”执行同名属性同步
        /// </summary>
        private static void SyncCommonPropertiesToInsertedEntities(DBTrans tr, List<Entity> insertedEntities, Dictionary<string, string> sourceMap)
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
        private static Dictionary<string, string> BuildMergedPropertyMapFromCandidates(
            DBTrans tr,
            List<JsonHelper.OverlapCandidate> candidates,
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
                LogManager.Instance.LogInfo($"\n多来源属性合并完成（白名单模式={JsonHelper._propertySyncUseWhitelistTemplate}），合并字段数={merged.Count}");
            }
            catch { }

            return merged;
        }

        #endregion

        #region

        /// <summary>
        /// 属性同步策略配置（第二版增强）
        /// </summary>
        // 是否优先同层来源（true 时同层会加权，且可额外筛选）
        public static readonly bool _propertySyncPreferSameLayer = true;

        // 属性继承只允许一个最高可信来源，避免相邻管道的字段被拼接到同一个新图元。
        public static readonly int _propertySyncMaxCandidates = 1;

        // 最多合并字段数量（防止异常图元导致字段爆炸）
        public static readonly int _propertySyncMaxMergedFields = 200;

        // 黑名单字段（这些字段不参与“交叠继承”）
        public static readonly HashSet<string> _propertySyncBlacklist = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {   "ID","GUID","UUID","OBJECTID","HANDLE", "CREATEDAT","UPDATEDAT","CREATETIME","UPDATETIME","TIMESTAMP",
            "CREATEDBY","UPDATEDBY","USER","USERNAME","OWNER", "VERSION","REVISION","REV", "FILENAME","FILEPATH",
            "FILEHASH","PREVIEWIMAGEPATH","PREVIEWIMAGENAME", "BLOCKNAME","LAYERNAME", "NAME", "名称"
        };
 



        #endregion

        // 上一次“整图插入”缓存
        private static byte[]? _lastCopyDwgBytes;// 最后一次复制的 DWG 文件字节内容，优先使用字节缓存以避免临时文件被删后无法重复
        private static string? _lastCopyDwgFileNameBase; // 最后一次复制的 DWG 文件基础名称（不含路径和扩展名，用于生成临时文件名，避免重复执行时文件名过长或包含非法字符）
        private static string? _lastCopyDwgPath; // 最后一次复制的 DWG 文件路径（仅在没有字节缓存时使用，存在被删除风险）
        private static GraphicInsertContext? _lastGraphicInsertContext; // 与上一次 DWG 缓存配套的分类上下文，确保重复插入时文件和类型一致

        /// <summary>
        /// 执行“整图复制”命令，并缓存相关信息以支持重复执行（空格键再次插入同一图元）
        /// </summary>
        /// <param name="sourceFilePath">源文件路径</param>
        /// <param name="insertContext">当前图元的分类插入上下文</param>
        public static void ExecuteCopyDwgAllFastWithRepeat(
            string sourceFilePath,
            GraphicInsertContext? insertContext = null)
        {
            // 新增：Drag/执行中禁止再次触发，避免命令重入导致 CAD 崩溃
            if (IsCopyDwgAllFastDragging || IsCopyDwgAllFastBusy)
            {
                Env.Editor?.WriteMessage("\n当前图元正在跟随插入，请先左键落点或按 Esc 结束后再触发。");
                return;
            }

            try
            {
                // 先缓存上下文，再发出 AutoCAD 命令，保证 COPYDWGALLFASTLAST 能使用同一图元类型。
                _lastGraphicInsertContext = insertContext;

                //判断源文件路径有效性
                if (VariableDictionary.resourcesFile != null && VariableDictionary.resourcesFile.Length > 0)
                {
                    _lastCopyDwgBytes = (byte[])VariableDictionary.resourcesFile.Clone();// 把用户按键指定的文件克隆一份字节数组，避免后续被修改

                    _lastCopyDwgFileNameBase = string.IsNullOrWhiteSpace(VariableDictionary.btnFileName)
                        ? "GB_CopyDwgAllFast"
                        : VariableDictionary.btnFileName; // 使用用户指定的按钮名作为基础文件名，避免重复执行时文件名过长或包含非法字符
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
                string? tempFilePath = null;
                //VariableDictionary.winformTextBoxScale = AutoCadHelper.GetScale();// 同步最新图纸比例，避免用户忘了更新导致插入图元过大过小
                //VariableDictionary.wpfTextBoxScale = AutoCadHelper.GetScale();
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
                    tempFilePath = Path.Combine(Path.GetTempPath(), $"{baseName}_{Guid.NewGuid():N}.dwg");
                    // 写入临时文件
                    System.IO.File.WriteAllBytes(tempFilePath, _lastCopyDwgBytes);
                }
                else
                {
                    tempFilePath = _lastCopyDwgPath;// 没有字节缓存则使用上次的路径（可能是原文件路径，存在被删除风险）
                }
                // 最后再次验证路径有效性，避免误用已被删除的临时文件路径
                if (string.IsNullOrWhiteSpace(tempFilePath) || !System.IO.File.Exists(tempFilePath))
                {
                    Env.Editor.WriteMessage("\n没有可重复的上一次插入命令。");
                    return;
                }

                if (tempFilePath != null)
                    //插入源文件中的图元到当前图纸
                    CopyDwgAllFast(tempFilePath, _lastGraphicInsertContext);// 直接调用插入方法，传入路径和分类上下文
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
        [CommandMethod("COPYDWGALLFAST")] // 注册 CAD 命令名，允许在命令行输入 COPYDWGALLFAST 调用
        public static void CopyDwgAllFast(
            string sourceFilePath,
            GraphicInsertContext? insertContext = null) // 整图插入主方法，同时接收可选的分类上下文
        {
            // 获取 LogManager 的单例实例，用于记录日志
            var logger = LogManager.Instance;
            // 在日志文件中记录命令开始执行
            logger.LogInfo(
                $">>> 开始执行 COPYDWGALLFAST 命令：CategoryPath={insertContext?.CategoryPath ?? "未知"}, EntityType={insertContext?.EntityType.ToString() ?? "Unknown"}");

            // 获取当前活动的 AutoCAD 文档对象
            var doc = Application.DocumentManager.MdiActiveDocument;
            // 如果未找到活动文档，说明 CAD 环境异常，记录错误并退出
            if (doc == null)
            {
                // 记录严重错误日志
                logger.LogError("错误：未找到活动文档。");
                // 回调通知上层调用者失败
                RaiseCopyDwgAllFastCompleted(false, "未找到活动文档。");
                // 结束方法执行
                return;
            }

            // 获取编辑器对象 ed，用于向 CAD 命令行发送提示消息
            var ed = doc.Editor;

            // 尝试进入互斥锁，防止命令重入（即防止在上一次没结束时再次触发）导致崩溃
            if (!TryEnterCopyDwgAllFastBusy())
            {
                // 记录警告日志
                logger.LogWarning("警告：当前命令忙碌中。");
                // 在 CAD 命令行提示用户
                ed.WriteMessage("\n当前插入命令正在执行中，请先完成当前插入（左键落点或 Esc）。");
                // 回调通知上层调用者失败
                RaiseCopyDwgAllFastCompleted(false, "当前命令忙碌中。");
                // 结束方法执行
                return;
            }

            // 获取当前文档的数据库对象 destDb，这是我们要插入图元的目标数据库
            var destDb = doc.Database;
            // 标记 deleteAfter，判断是否需要在成功后删除源文件（通常用于临时文件）
            bool deleteAfter = false;
            // 标记 insertSuccess，记录本次插入操作最终是否成功
            bool insertSuccess = false;
            // 变量 failReason，用于记录如果失败时的具体原因字符串
            string? failReason = null;

            try // 主流程异常捕获块，包裹整个插入逻辑
            {
                // 内部 try-catch，专门用于判断源文件是否为临时文件
                try
                {
                    // 如果传入了有效的源文件路径
                    if (!string.IsNullOrEmpty(sourceFilePath))
                    {
                        // 获取系统临时文件夹的路径
                        var tempDir = Path.GetTempPath();
                        // 如果源文件路径以临时目录开头，则视为临时文件
                        if (!string.IsNullOrEmpty(tempDir) && sourceFilePath.StartsWith(tempDir, StringComparison.OrdinalIgnoreCase))
                        {
                            // 标记为需要删除
                            deleteAfter = true;
                            // 记录日志，说明识别到了临时文件
                            logger.LogInfo($"源文件识别为临时文件: {sourceFilePath}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 如果判断路径时发生异常，记录错误日志
                    logger.LogError($"判断临时文件异常: {ex.Message}");
                    // 发生异常时保守处理，不删除文件
                    deleteAfter = false;
                }

                // 初始化插入点 targetPoint 为原点 (0,0,0)，后续拖拽会更新此值
                Point3d targetPoint = Point3d.Origin;

                // 锁定文档以确保线程安全的数据库写操作
                using (doc.LockDocument())
                // 创建一个新的数据库对象 sourceDb 用于读取源 DWG 文件
                using (var sourceDb = new Autodesk.AutoCAD.DatabaseServices.Database(false, true))
                {
                    // 从磁盘读取源 DWG 文件内容到内存数据库 sourceDb
                    sourceDb.ReadDwgFile(sourceFilePath, FileShare.Read, true, null);
                    logger.LogInfo($"[插入流程][步骤1] 本地 DWG 读取完成：文件名={Path.GetFileName(sourceFilePath)}");
                    // 关闭输入流，释放文件占用，允许其他进程访问该文件
                    sourceDb.CloseInput(true);

                    // 开启针对目标数据库的事务处理 tr
                    using (var tr = new DBTrans())
                    {
                        // 将源数据库作为匿名块定义插入到目标数据库 destDb 中
                        var blkDefId = destDb.Insert("*U", sourceDb, false);
                        logger.LogInfo($"[插入流程][步骤2] DWG 已插入为临时块定义：BlockTableRecord={blkDefId}");
                        // 基于插入的块定义创建一个新的块参照实体 br
                        var br = new BlockReference(targetPoint, blkDefId);
                        logger.LogInfo("[插入流程][步骤2] 临时块参照已创建。");

                        // 初始化缩放比例 scale 为 1.0
                        double scale = AutoCadHelper.GetScale();
                        // 根据当前界面状态（WinForm 或 WPF）获取用户设置的缩放比例
                        if (VariableDictionary.winForm_Status) // 如果是 WinForm 模式
                        {
                            try { scale = VariableDictionary.winformTextBoxScale; } // 尝试读取 WinForm 的比例值
                            catch { scale = 1.0; } // 读取失败则使用默认值 1.0
                        }
                        else // 如果是 WPF 模式
                        {
                            try { scale = VariableDictionary.wpfTextBoxScale; } // 尝试读取 WPF 的比例值
                            catch { scale = 1.0; } // 读取失败则使用默认值 1.0
                        }
                        // 如果比例值非法（NaN 或小于等于 0），强制重置为 1.0
                        if (double.IsNaN(scale) || scale <= 0) scale = 1.0;
                        logger.LogInfo($"[插入流程][步骤3] 当前图纸及界面比例已确定：Scale={scale:F6}, WinForm={VariableDictionary.winForm_Status}");

                        // 设置块参照 br 的初始缩放比例
                        br.ScaleFactors = new Scale3d(scale);
                        // 将块参照 br 添加到当前空间的模型空间中
                        var entityObjectId = tr.CurrentSpace.AddEntity(br);
                        logger.LogInfo($"[插入流程][步骤4] 临时块参照已加入当前空间：ObjectId={entityObjectId}");
                        // 以写模式打开刚刚添加的块参照 fileEntity，以便后续修改属性或变换
                        var fileEntity = (BlockReference)tr.GetObject(entityObjectId, OpenMode.ForWrite);
                        // 整图插入路径不会自动创建属性引用，这里先按块定义补齐属性引用
                        EnsureBlockAttributeReferences(tr, fileEntity, logger);
                        logger.LogInfo($"[插入流程][步骤5] 块属性引用补齐完成：ObjectId={fileEntity.ObjectId}, AttributeCount={fileEntity.AttributeCollection.Count}");
                        // 记录当前的旋转角度 tempAngle，用于拖拽过程中的增量计算
                        double tempAngle = VariableDictionary.entityRotateAngle;
                        // 记录当前的缩放比例 tempScale，用于拖拽过程中的增量计算
                        double tempScale = scale;
                        // 如果初始旋转角度不为 0，则预先应用旋转，确保预览方向正确
                        if (Math.Abs(tempAngle) > 1e-12)
                        {
                            // 绕 Z 轴旋转块参照 fileEntity
                            fileEntity.TransformBy(Matrix3d.Rotation(tempAngle, Vector3d.ZAxis, targetPoint));
                        }
                        // 创建自定义拖拽类 JigEx，用于实现鼠标跟随效果
                        var entityBlock = new JigEx((mpw, _) =>
                        {
                            // 移动块参照从旧位置到新鼠标位置 mpw
                            fileEntity.Move(targetPoint, mpw);
                            // 更新全局记录的插入点 targetPoint
                            targetPoint = mpw;

                            // 检查外部设置的旋转角度是否发生变化
                            if (VariableDictionary.entityRotateAngle != tempAngle)
                            {
                                // 先逆向旋转抵消之前的角度
                                fileEntity.TransformBy(Matrix3d.Rotation(-tempAngle, Vector3d.ZAxis, targetPoint));
                                // 更新缓存的角度值 tempAngle
                                tempAngle = VariableDictionary.entityRotateAngle;
                                // 应用新的旋转角度
                                fileEntity.TransformBy(Matrix3d.Rotation(tempAngle, Vector3d.ZAxis, targetPoint));
                            }

                            // 获取当前界面实时设置的缩放比例 currentUiScale
                            double currentUiScale = scale;
                            if (VariableDictionary.winForm_Status) // WinForm 模式
                            {
                                try { currentUiScale = VariableDictionary.winformTextBoxScale; } // 读取实时比例
                                catch { currentUiScale = tempScale; } // 失败则沿用上次有效值
                            }
                            else // WPF 模式
                            {
                                try { currentUiScale = VariableDictionary.wpfTextBoxScale; } // 读取实时比例
                                catch { currentUiScale = tempScale; } // 失败则沿用上次有效值
                            }
                            // 校验比例值合法性
                            if (double.IsNaN(currentUiScale) || currentUiScale <= 0) currentUiScale = 1.0;

                            // 如果比例发生变化，则更新块参照的缩放
                            if (Math.Abs(currentUiScale - tempScale) > 1e-9)
                            {
                                // 设置新比例
                                fileEntity.ScaleFactors = new Scale3d(currentUiScale);
                                // 更新缓存比例 tempScale
                                tempScale = currentUiScale;
                            }
                        });

                        // 执行图形绘制，显示拖拽预览
                        entityBlock.DatabaseEntityDraw(wd => wd.Geometry.Draw(fileEntity));
                        // 设置命令行提示文字
                        entityBlock.SetOptions(msg: "\n指定插入点");
                        // 确保绘图区域获得焦点，以便接收鼠标事件
                        EnsureDwgViewFocus();

                        // 声明变量 endPoint 用于接收拖拽结果
                        PromptResult endPoint;
                        // 标记当前处于拖拽交互状态
                        _isCopyDwgAllFastDragging = true;
                        try
                        {
                            // 启动编辑器拖拽交互，等待用户点击确认或取消
                            endPoint = Env.Editor.Drag(entityBlock);
                        }
                        finally
                        {
                            // 无论拖拽成功与否，都清除拖拽状态标记
                            _isCopyDwgAllFastDragging = false;
                        }

                        // 如果用户按 Esc 取消或拖拽失败
                        if (endPoint.Status != PromptStatus.OK)
                        {
                            // 记录失败原因
                            failReason = "用户取消插入。";
                            // 记录日志
                            logger.LogInfo("用户取消了插入操作。");
                            // 回滚事务，不保存任何更改
                            tr.Abort();
                            // 退出方法
                            return;
                        }

                        // 记录日志，表示拖拽结束，开始核心逻辑
                        logger.LogInfo($"拖拽结束：插入点=({targetPoint.X:F3},{targetPoint.Y:F3},{targetPoint.Z:F3}), ObjectId={fileEntity.ObjectId}, 图块名={fileEntity.Name}");
                        logger.LogInfo($"[插入流程][步骤7] 用户已确认插入点：ObjectId={fileEntity.ObjectId}");
                        logger.LogInfo($"开始重叠检测：候选上限={_propertySyncMaxCandidates}, 当前图层={fileEntity.Layer}");

                        // ================== 核心新功能：重叠检测与属性继承（带 LogManager 日志） ==================

                        // 初始化字典 overlapSourcePropertyMap，用于存储从重叠图元读取到的属性
                        var overlapSourcePropertyMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        logger.LogInfo("[插入流程][步骤8] 开始执行重叠图元检测。");

                        List<OverlapCandidate> overlapCandidates = new List<OverlapCandidate>();

                        if (!VariableDictionary.winForm_Status)//如果不是winform状态，则查找当前插入点附近所有可能重叠的候选图元

                            // 调用 FindOverlappedCandidates 查找当前插入点附近所有可能重叠的候选图元
                            overlapCandidates = JsonHelper.FindOverlappedCandidates(tr, fileEntity, _propertySyncMaxCandidates);

                        // 记录日志，输出找到的候选图元数量
                        logger.LogInfo($"重叠检测完成：候选数量={overlapCandidates.Count}");
                        for (int candidateIndex = 0; candidateIndex < overlapCandidates.Count; candidateIndex++)
                        {
                            var candidate = overlapCandidates[candidateIndex];
                            logger.LogInfo($"重叠候选[{candidateIndex + 1}]：ObjectId={candidate.Entity.ObjectId}, 类型={candidate.Entity.GetType().Name}, 图层={candidate.Entity.Layer}, 评分={candidate.Score:F6}, 同层={candidate.IsSameLayer}");
                        }

                        // 如果找到了重叠候选图元
                        if (overlapCandidates.Count > 0)
                        {
                            // 调用 GetActiveWhitelist 根据当前图元的专业类型获取对应的属性白名单
                            var activeWhitelist = JsonHelper.GetActiveWhitelist(fileEntity);

                            // 调用 BuildMergedPropertyMapFromCandidates 从候选图元中合并属性，并应用白名单/黑名单过滤
                            overlapSourcePropertyMap = BuildMergedPropertyMapFromCandidates(tr, overlapCandidates, activeWhitelist);

                            // 记录日志，输出过滤后待同步的属性数量
                            logger.LogInfo($"入口管道属性合并完成：属性数量={overlapSourcePropertyMap.Count}");

                            // 如果有属性，打印前 5 个属性的键值对，方便排查
                            if (overlapSourcePropertyMap.Count > 0)
                            {
                                int debugCount = 0; // 计数器
                                foreach (var kv in overlapSourcePropertyMap) // 遍历属性字典
                                {
                                    // 记录每个属性的名称和值
                                    logger.LogInfo($"入口管道属性[{debugCount + 1}]：{kv.Key}={kv.Value}");
                                    debugCount++; // 计数器加 1
                                    if (debugCount >= 5) break; // 只打印前 5 个，避免日志过多
                                }
                            }
                            else
                            {
                                // 如果数量为 0，记录警告日志，提示可能是白名单过滤掉了
                                logger.LogWarning("未找到可同步属性：请检查入口管道块属性、扩展字典或白名单配置。");
                            }
                        }
                        else
                        {
                            // 如果没有重叠，记录警告日志，提示用户检查插入位置
                            logger.LogWarning("未检测到重叠图元：无法从入口管道读取 DN/PN。");
                        }

                        // 如果成功读取到有效的重叠属性
                        if (overlapSourcePropertyMap.Count > 0)
                        {
                            // 第一步：调用 SyncCommonPropertiesToBlockReference 将属性同步到当前的块参照本身
                            SyncCommonPropertiesToBlockReference(tr, fileEntity, overlapSourcePropertyMap);
                            // 第二步：调用 SyncCommonPropertiesToEntityXRecord 将属性同步到块参照的扩展字典
                            SyncCommonPropertiesToEntityXRecord(tr, fileEntity, overlapSourcePropertyMap);

                            // 记录日志，说明已向块参照同步了多少个属性
                            logger.LogInfo($"已向新图块同步入口管道属性：属性数量={overlapSourcePropertyMap.Count}");
                        }
                        logger.LogInfo($"[插入流程][步骤9] 重叠属性读取及继承准备完成：继承属性数量={overlapSourcePropertyMap.Count}");

                        logger.LogInfo($"准备执行规范匹配：块ObjectId={fileEntity.ObjectId}");
                        logger.LogInfo($"[插入流程][步骤10] 开始执行连接规范查询：图元类型={insertContext?.EntityType.ToString() ?? "Unknown"}, 分类路径={insertContext?.CategoryPath ?? "未知"}");

                        // 所有插入图元都在炸开前完成连接规范查询，确保规范属性可以写入原始块的 AttributeReference。
                        FlangeStandardMatchResponse? flangeStandardResponse =
                            ApplyFlangeStandardAttributes(tr, fileEntity, overlapSourcePropertyMap, logger);

                        // 在炸开原始图元块之前先回写规范属性，确保原始块自身的规范属性不会丢失。
                        if (flangeStandardResponse?.Success == true)
                        {
                            int blockUpdatedCount = new StandardPropertySyncService()
                                .ApplyToBlockReference(
                                    tr.Transaction,
                                    fileEntity,
                                    flangeStandardResponse,
                                    createMissingAttributes: false);
                            logger.LogInfo(
                                $"插入块炸开前连接规范回写完成：ObjectId={fileEntity.ObjectId}, 写入数量={blockUpdatedCount}");
                        }
                        else
                        {
                            logger.LogWarning("[插入流程][步骤10] 规范查询未返回成功结果，炸开前没有可回写的规范属性。");
                        }

                        // ================== 结束核心新功能 ==================

                        // 规范匹配及规范属性回写完成后，先让用户确认并编辑最终要插入的图元属性。
                        if (!TryEditPropertiesBeforeInsert(
                             tr,
                             fileEntity,
                             insertContext?.CategoryPath,
                             logger,
                             flangeStandardResponse?.Success == true,
                             flangeStandardResponse,
                             overlapSourcePropertyMap,
                             out var editedProperties))
                        {
                            // 用户取消或窗口异常时必须在炸开前退出，事务不会把临时插入保存到当前图纸。
                            failReason = "用户取消插入或属性编辑窗口打开失败。";
                            tr.Abort();
                            return;
                        }
                        logger.LogInfo($"[插入流程][步骤11-12] 插入前属性窗口已确认：最终属性数量={editedProperties.Count}");

                        // 将用户确认后的值写回原始块，随后炸开时这些值会随属性引用进入最终图元。
                        ApplyEditedPropertiesToEntity(tr, fileEntity, editedProperties, logger);
                        logger.LogInfo($"[插入流程][步骤13] 用户确认属性已写回临时块：ObjectId={fileEntity.ObjectId}, 属性数量={editedProperties.Count}");

                        // 创建集合 newIds 用于存储分解后产生的所有新实体
                        var newIds = new DBObjectCollection();

                        // 执行分解操作，将块参照 fileEntity 炸开为独立的图元
                        fileEntity.Explode(newIds);
                        logger.LogInfo($"[插入流程][步骤14] 临时块炸开完成：原始ObjectId={fileEntity.ObjectId}, 生成实体数量={newIds.Count}");

                        // 删除原始的块参照实体 fileEntity，因为已经被分解替代
                        fileEntity.Erase();

                        // 创建列表 insertedEntities 用于跟踪所有新加入到图纸中的实体
                        var insertedEntities = new List<Entity>();
                        // 遍历分解产生的每一个实体 ent
                        foreach (Entity ent in newIds)
                        {
                            // 将实体 ent 添加到当前空间，使其在图纸中可见
                            tr.CurrentSpace.AddEntity(ent);
                            // 将实体 ent 加入跟踪列表 insertedEntities
                            insertedEntities.Add(ent);
                        }
                        logger.LogInfo($"[插入流程][步骤15] 炸开实体已加入当前空间：实体数量={insertedEntities.Count}");

                        // 如果之前读取到了重叠属性，需要将这些属性进一步同步到分解后的子实体上
                        if (overlapSourcePropertyMap.Count > 0)
                        {
                            // 记录日志，说明正在向多少个分解后的实体同步属性
                            logger.LogInfo($"图元炸开完成：新实体数量={insertedEntities.Count}，开始同步继承属性");

                            // 调用 SyncCommonPropertiesToInsertedEntities 遍历分解后的实体，将属性同步到它们的扩展数据或嵌套块中
                            SyncCommonPropertiesToInsertedEntities(tr, insertedEntities, overlapSourcePropertyMap);

                            // 记录日志，说明属性同步流程结束
                            logger.LogInfo("炸开后属性同步完成。");
                        }

                        // 规范属性位于炸开后的嵌套块中，必须在入口管道属性同步完成后再写入，避免被继承值覆盖。
                        ApplyStandardResponseToInsertedEntities(
                            tr,
                            insertedEntities,
                            flangeStandardResponse,
                            logger);
                        logger.LogInfo($"[插入流程][步骤16] 继承属性、规范属性和用户属性同步流程已完成：实体数量={insertedEntities.Count}, 规范查询成功={flangeStandardResponse?.Success == true}");

                        // 规范同步完成后，为确认窗口中缺失的字段创建隐藏 AttributeReference。
                        // 这样字段不仅保存在 XRecord 中，也会出现在最终图块属性集合中。
                        foreach (Entity insertedEntity in insertedEntities)
                        {
                            if (insertedEntity is BlockReference insertedBlockReference)
                            {
                                int createdAttributeCount = new StandardPropertySyncService()
                                    .EnsureEditedAttributes(
                                        tr.Transaction,
                                        insertedBlockReference,
                                        editedProperties);
                                logger.LogInfo(
                                    $"炸开后插入属性定义补充完成：ObjectId={insertedBlockReference.ObjectId}, 新增AttributeReference数量={createdAttributeCount}");
                            }
                        }

                        // 规范同步完成后再次应用用户编辑值，确保用户修改优先于默认规范值。
                        foreach (Entity insertedEntity in insertedEntities)
                        {
                            ApplyEditedPropertiesToEntity(tr, insertedEntity, editedProperties, logger);
                        }

                        // 检查是否有待创建的标注文本 dimString
                        if (VariableDictionary.dimString != null)
                        {
                            try
                            {
                                // 调用 Command.DDimLinear 创建线性标注
                                Command.DDimLinear(tr, entityBlock.MousePointWcsLast, VariableDictionary.dimString);
                                // 清空标注缓存，防止重复创建
                                VariableDictionary.dimString = null;
                            }
                            catch (Exception ex)
                            {
                                // 如果标注创建失败，记录错误日志，但不影响主流程
                                logger.LogError($"设置标注样式失败: {ex.Message}");
                            }
                        }

                        // 提交事务 tr，将所有数据库更改永久保存
                        tr.Commit();
                        // 刷新屏幕显示，让用户立即看到结果
                        Env.Editor.Redraw();
                        // 标记操作成功
                        insertSuccess = true;
                        // 记录日志，说明事务提交成功
                        logger.LogInfo($"事务提交成功：插入完成，炸开实体数量={insertedEntities.Count}");
                    }
                }

                // 如果标记为临时文件且插入成功，则清理临时文件
                if (deleteAfter && insertSuccess && !string.IsNullOrEmpty(sourceFilePath))
                {
                    try
                    {
                        // 再次确认文件存在
                        if (System.IO.File.Exists(sourceFilePath))
                        {
                            // 删除临时 DWG 文件，释放磁盘空间
                            System.IO.File.Delete(sourceFilePath);
                            // 记录日志，说明已删除临时文件
                            logger.LogInfo($"已删除临时 DWG 文件: {sourceFilePath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        // 如果删除失败，记录错误日志，不影响主流程
                        logger.LogError($"尝试删除临时 DWG 失败: {ex.Message}");
                    }
                }

                // 重置 WinForm 状态标志
                VariableDictionary.winForm_Status = false;

            }
            catch (Exception ex) // 捕获整个执行过程中的任何未预期异常
            {
                // 记录异常消息到 failReason
                failReason = ex.Message;
                // 记录错误日志
                logger.LogError($"COPYDWGALLFAST 执行失败: {ex.Message}");
                // 记录堆栈跟踪日志，方便排查深层错误
                logger.LogError($"堆栈跟踪: {ex.StackTrace}");
            }
            finally // 无论成功失败都会执行的收尾工作
            {
                // 确保拖拽状态被清除
                _isCopyDwgAllFastDragging = false;
                // 释放互斥锁，允许下次执行
                ExitCopyDwgAllFastBusy();
                // 通知上层调用者最终结果
                RaiseCopyDwgAllFastCompleted(insertSuccess, insertSuccess ? null : failReason);
                // 记录命令执行结束日志
                logger.LogInfo("<<< 命令执行结束");
            }
        }

        /// <summary>
        /// 根据块定义中的 AttributeDefinition 补齐块参照的 AttributeReference。
        /// </summary>
        private static void EnsureBlockAttributeReferences(DBTrans tr, BlockReference blockReference, LogManager logger)
        {
            // 参数无效时不处理
            if (tr == null || blockReference == null) return;

            // 先收集当前块参照已有的 Tag，避免重复创建属性引用
            var existingTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ObjectId attributeId in blockReference.AttributeCollection)
            {
                if (tr.GetObject(attributeId, OpenMode.ForRead) is AttributeReference attribute)
                {
                    string tag = PipelineCadPropertyKeyHelper.Decode((attribute.Tag ?? string.Empty).Trim());
                    if (!string.IsNullOrWhiteSpace(tag)) existingTags.Add(NormalizePropertyKey(tag));
                }
            }

            // 读取块定义中的属性定义
            var blockDefinition = tr.GetObject(blockReference.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            if (blockDefinition == null || !blockDefinition.HasAttributeDefinitions) return;

            // 遍历 AttributeDefinition，同时检查 Tag 和 Prompt
            foreach (ObjectId entityId in blockDefinition)
            {
                if (!(tr.GetObject(entityId, OpenMode.ForRead) is AttributeDefinition attributeDefinition) ||
                    attributeDefinition.Constant ||
                    string.IsNullOrWhiteSpace(attributeDefinition.Tag))
                {
                    continue;
                }

                // 同一 Tag 已经有引用时不重复添加
                string normalizedTag = NormalizePropertyKey(attributeDefinition.Tag);
                if (existingTags.Contains(normalizedTag)) continue;

                // 按块变换创建属性引用并保留定义中的默认值
                var attributeReference = new AttributeReference();
                attributeReference.SetAttributeFromBlock(attributeDefinition, blockReference.BlockTransform);
                blockReference.AttributeCollection.AppendAttribute(attributeReference);
                tr.Transaction.AddNewlyCreatedDBObject(attributeReference, true);
                existingTags.Add(normalizedTag);
            }
        }

        /// <summary>
        /// 根据插入图元自身或重叠图元继承的属性构建连接规范查询请求，调用 API 匹配规范并返回结果。
        /// </summary>
        /// <param name="transaction">数据库事务，用于可能的数据读写。</param>
        /// <param name="flangeBlock">当前处理的图元块参照，用于获取图元信息。</param>
        /// <param name="inheritedProperties">从重叠图元继承的属性字典（键为属性名，值为属性值）。</param>
        /// <param name="logger">日志管理器，记录操作和诊断信息。</param>
        /// <returns>连接规范匹配响应对象；若不属于目标连接方式、缺少关键属性或 API 调用失败则返回 null。</returns>
        private static FlangeStandardMatchResponse? ApplyFlangeStandardAttributes(
            DBTrans transaction,
            BlockReference flangeBlock,
            IDictionary<string, string> inheritedProperties,
            LogManager logger)
        {
            // 参数有效性检查
            if (transaction == null || flangeBlock == null || inheritedProperties == null)
            {
                logger?.LogWarning("参数无效：transaction、flangeBlock 或 inheritedProperties 为空。");
                return null;
            }

            // 记录开始处理信息
            logger.LogInfo($"开始连接规范查询：块ObjectId={flangeBlock.ObjectId}, 图块名={flangeBlock.Name}, 继承属性数量={inheritedProperties.Count}");
            logger.LogInfo($"可用属性键：{string.Join(",", inheritedProperties.Keys)}");

            // 合并入口图元属性和当前插入块已有属性，兼容连接方式位于不同属性来源的情况。
            var queryProperties = new Dictionary<string, string>(inheritedProperties, StringComparer.OrdinalIgnoreCase);
            var insertedBlockProperties = ReadEntityPropertyMap(transaction, flangeBlock);
            logger.LogInfo($"[规范查询][属性来源] 当前临时块属性数量={insertedBlockProperties.Count}，继承属性数量={inheritedProperties.Count}");
            logger.LogInfo($"[规范查询][关键字段] 继承CONN_TYPE={FindProperty(inheritedProperties, "CONN_TYPE", "DNCONN_TYPE", "连接方式", "连接形式") ?? "<空>"}，块内CONN_TYPE={FindProperty(insertedBlockProperties, "CONN_TYPE", "DNCONN_TYPE", "连接方式", "连接形式") ?? "<空>"}");
            logger.LogInfo($"[规范查询][关键字段] 继承DN={FindProperty(inheritedProperties, "DN", "公称通径", "通径", "管径", "公称直径") ?? "<空>"}，块内DN={FindProperty(insertedBlockProperties, "DN", "公称通径", "通径", "管径", "公称直径") ?? "<空>"}");
            logger.LogInfo($"[规范查询][关键字段] 继承PN={FindProperty(inheritedProperties, "PN", "公称压力", "压力等级") ?? "<空>"}，块内PN={FindProperty(insertedBlockProperties, "PN", "公称压力", "压力等级") ?? "<空>"}");
            logger.LogInfo($"[规范查询][关键字段] 继承FLG_STD={FindProperty(inheritedProperties, "FLG_STD", "法兰标准", "标准号") ?? "<空>"}，块内FLG_STD={FindProperty(insertedBlockProperties, "FLG_STD", "法兰标准", "标准号") ?? "<空>"}");
            foreach (KeyValuePair<string, string> property in insertedBlockProperties)
            {
                if (!queryProperties.ContainsKey(property.Key))
                {
                    queryProperties[property.Key] = property.Value;
                }
            }

            // 连接方式必须来自当前蝶阀的实际属性值，不能仅凭块定义的 Prompt 或默认值触发查询。
            // 这样可以避免普通阀门或连接方式尚未同步完成时误调用法兰规范接口。
            string connectionMode = FindProperty(
                queryProperties,
                "连接方式",
                "CONN_TYPE",
                "DNCONN_TYPE",
                "CONNECTION_MODE",
                "CONNECTIONTYPE",
                "连接形式",
                "连接型式",
                "连接类别") ?? string.Empty;
            if (!ShouldQueryFlangeStandard(queryProperties))
            {
                logger.LogWarning($"[规范查询][跳过] 当前图元连接方式不属于法兰/对夹：CONN_TYPE={connectionMode}，查询属性总数={queryProperties.Count}");
                return null;
            }
            logger.LogInfo($"当前图元连接方式确认，开始查询法兰/对夹连接规范：CONN_TYPE={connectionMode}");

            // 从继承属性中提取关键信息：DN、PN，使用多个可能的键名进行查找（不区分大小写）
            string dn = FindProperty(queryProperties, "DN", "公称通径", "通径", "管径", "公称直径");
            string pn = FindProperty(queryProperties, "PN", "公称压力", "压力等级");

            // 必须同时具有 DN 和 PN 才能进行查询
            if (string.IsNullOrWhiteSpace(dn) || string.IsNullOrWhiteSpace(pn))
            {
                logger.LogWarning($"[规范查询][跳过] 缺少 DN/PN：DN={dn ?? "<空>"}, PN={pn ?? "<空>"}，CONN_TYPE={connectionMode}");
                return null;
            }

            // 构建查询请求对象
            var request = new FlangeStandardMatchRequest
            {
                // 如果图元自身携带规范库编码，则使用图元配置；没有配置时交由服务器按规范号等条件选择候选系列。
                FamilyCode = FindProperty(queryProperties, "FAMILY_CODE", "规范大类编码") ?? "FLANGE",
                SeriesCode = FindProperty(queryProperties, "SERIES_CODE", "规范系列编码") ?? string.Empty,
                // 标准化 DN、PN、系列
                DN = NormalizeDn(dn),
                PN = NormalizePn(pn),
                Series = NormalizeSeries(FindProperty(queryProperties, "SERIES", "钢管系列", "管道系列")),
                // 优先使用图元已有的 FLG_STD/法兰标准作为服务器筛选条件
                StandardNumber = FindProperty(queryProperties, "FLG_STD", "法兰标准", "标准号") ?? string.Empty,
                TableNumber = FindProperty(queryProperties, "TABLE_NUMBER", "标准表号", "表号") ?? string.Empty,
                ConnectionMode = connectionMode,
                // 连接方式与法兰类型是不同业务字段；未提供法兰类型时不限制候选记录。
                FlangeType = FindProperty(queryProperties, "FLG_TYPE", "法兰类型") ?? string.Empty,
                FaceType = FindProperty(queryProperties, "FACE_TYPE", "密封面形式", "密封面型式") ?? string.Empty
            };

            // 记录标准号来源信息，便于排查
            string inheritedStandard = FindProperty(queryProperties, "FLG_STD", "DRAWINGNO.STANDARDNO", "法兰标准", "标准号") ?? string.Empty;
            logger.LogInfo($"规范标准号来源：使用标准库配置值={request.StandardNumber}，入口管道属性中的候选值={inheritedStandard}");
            logger.LogInfo($"请求参数：FamilyCode={request.FamilyCode}, SeriesCode={request.SeriesCode}, StandardNumber={request.StandardNumber}, TableNumber={request.TableNumber}, DN={request.DN}, PN={request.PN}, Series={request.Series}, ConnectionMode={request.ConnectionMode}, FlangeType={request.FlangeType}, FaceType={request.FaceType}");

            try
            {
                // 调用 API 服务进行匹配（同步等待异步结果）
                var standardApiService = new StandardApiService();
                FlangeStandardMatchResponse response = standardApiService
                    .MatchFlangeAsync(request)
                    .GetAwaiter()
                    .GetResult();

                if (response == null)
                {
                    logger.LogError("API 返回 null。");
                    return null;
                }

                // 记录 API 返回的详细信息
                logger.LogInfo($"API 返回：Success={response.Success}, MatchCount={response.MatchCount}, IsUniqueMatch={response.IsUniqueMatch}, Message={response.Message}, 属性数量={response.Attributes?.Count ?? 0}");
                if (response.Attributes != null)
                {
                    foreach (KeyValuePair<string, string> attribute in response.Attributes)
                    {
                        logger.LogInfo($"规范返回属性：{attribute.Key}={attribute.Value}");
                    }
                }

                // 若匹配失败，记录警告但仍返回响应对象（调用方可根据 Success 判断）
                if (!response.Success)
                {
                    logger.LogWarning($"按指定标准号查询未命中：StandardNumber={request.StandardNumber}, DN={request.DN}, PN={request.PN}, 消息={response.Message}");

                    // 入口管道的 FLG_STD 可能是管道/阀门标准（例如 HG/T 20592），不一定是服务器法兰系列标准。
                    // 第一次按指定标准号未命中时，清空标准号重新按系列、DN、PN 等关键条件查询。
                    if (!string.IsNullOrWhiteSpace(request.StandardNumber))
                    {
                        var fallbackRequest = new FlangeStandardMatchRequest
                        {
                            FamilyCode = request.FamilyCode,
                            SeriesCode = request.SeriesCode,
                            StandardNumber = string.Empty,
                            TableNumber = string.Empty,
                            PN = request.PN,
                            DN = request.DN,
                            Series = request.Series,
                            FlangeType = string.Empty,
                            FaceType = string.Empty
                        };

                        logger.LogInfo(
                            $"指定标准号未命中，开始按法兰系列回退查询：SeriesCode={fallbackRequest.SeriesCode}, DN={fallbackRequest.DN}, PN={fallbackRequest.PN}, Series={fallbackRequest.Series}");

                        response = standardApiService
                            .MatchFlangeAsync(fallbackRequest)
                            .GetAwaiter()
                            .GetResult();

                        logger.LogInfo(
                            $"法兰标准回退查询结果：Success={response?.Success}, MatchCount={response?.MatchCount}, Message={response?.Message}");
                    }

                    if (!response.Success)
                    {
                        logger.LogWarning($"法兰标准查询最终未命中：DN={request.DN}, PN={request.PN}, 消息={response.Message}");
                        return response;
                    }
                }

                // 匹配成功，记录成功信息并返回响应
                logger.LogInfo($"查询成功，等待炸开后写入：DN={request.DN}, PN={request.PN}, 返回属性={response.Attributes.Count}");
                return response;
            }
            catch (Exception ex)
            {
                // 规范库不可用时保留已完成的图元插入和管道属性继承，不让网络故障破坏 CAD 事务。
                // 记录警告而非抛出异常，保证事务继续进行
                logger.LogWarning($"查询或回写失败：DN={request.DN}, PN={request.PN}, 错误={ex.Message}");
                return null;
            }
        }
        /// <summary>
        /// 将规范匹配结果应用到新插入的图块实体上，同步其属性值。
        /// </summary>
        /// <param name="transaction">数据库事务，用于读写实体。</param>
        /// <param name="insertedEntities">新插入的实体列表（通常为炸开后的块参照）。</param>
        /// <param name="response">规范匹配响应，包含要写入的属性键值对。</param>
        /// <param name="logger">日志管理器，用于记录操作日志。</param>
        private static void ApplyStandardResponseToInsertedEntities(
            DBTrans transaction,
            IList<Entity> insertedEntities,
            FlangeStandardMatchResponse? response,
            LogManager logger)
        {
            // 若响应无效或表示失败，则跳过属性写入并记录日志
            if (response == null || !response.Success)
            {
                logger.LogInfo("没有成功的规范查询结果，跳过炸开后规范属性写入。");
                return;
            }

            int blockCount = 0;          // 统计处理的块参照数量
            int updatedCount = 0;        // 累计所有块参照写入的属性总数

            // 遍历插入的每个实体
            foreach (Entity entity in insertedEntities)
            {
                // 只处理块参照（BlockReference），其他类型跳过并记录
                if (!(entity is BlockReference blockReference))
                {
                    logger.LogInfo($"炸开实体不包含属性：ObjectId={entity.ObjectId}, 类型={entity.GetType().Name}");
                    continue;
                }

                blockCount++;

                // 创建属性同步服务实例，并应用到当前块参照
                int entityUpdatedCount = new StandardPropertySyncService()
                    .ApplyToBlockReference(
                        transaction,
                        blockReference,
                        response,
                        createMissingAttributes: true);
                updatedCount += entityUpdatedCount;

                // 记录本次同步结果
                logger.LogInfo($"炸开后规范属性写入：ObjectId={blockReference.ObjectId}, 图块名={blockReference.Name}, 写入数量={entityUpdatedCount}");

                // 额外检查并记录响应中未能在该块参照中找到对应Tag的属性（用于诊断）
                LogUnmatchedStandardAttributes(transaction, blockReference, response, logger);
            }

            // 输出汇总信息
            logger.LogInfo($"炸开后规范属性写入完成：BlockReference数量={blockCount}, 返回属性数量={response.Attributes.Count}, 总写入数量={updatedCount}");
        }

        /// <summary>
        /// 检查规范匹配响应中的属性键是否在当前块参照的Attribute集合中存在对应的Tag，
        /// 并记录未命中的项，用于诊断数据不一致问题。
        /// </summary>
        /// <param name="transaction">数据库事务。</param>
        /// <param name="flangeBlock">当前处理的块参照。</param>
        /// <param name="response">规范匹配响应。</param>
        /// <param name="logger">日志管理器。</param>
        private static void LogUnmatchedStandardAttributes(
            DBTrans transaction,
            BlockReference flangeBlock,
            FlangeStandardMatchResponse response,
            LogManager logger)
        {
            // 若响应或属性字典为空，则无需检查
            if (response?.Attributes == null) return;

            // 收集当前块参照中所有属性的Tag（经过标准化处理），存入哈希集合以便快速查找
            var targetTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ObjectId attributeId in flangeBlock.AttributeCollection)
            {
                // 获取属性参照并读取其Tag
                if (transaction.GetObject(attributeId, OpenMode.ForRead) is AttributeReference attribute &&
                    !string.IsNullOrWhiteSpace(attribute.Tag))
                {
                    targetTags.Add(NormalizePropertyKey(attribute.Tag)); // 标准化Tag，如转大写或去除空格
                }
            }

            // 遍历响应中的每个属性键，检查是否在块参照的Tag集合中存在
            int unmatchedCount = 0;
            foreach (string responseKey in response.Attributes.Keys)
            {
                // 若标准化后的键未在集合中找到，则记录警告
                if (!targetTags.Contains(NormalizePropertyKey(responseKey)))
                {
                    unmatchedCount++;
                    logger.LogWarning($"返回属性未命中法兰DWG Tag：{responseKey}");
                }
            }

            // 输出统计信息
            logger.LogInfo($"法兰DWG属性检查：目标Tag数量={targetTags.Count}, 未命中规范属性数量={unmatchedCount}");
        }

        /// <summary>
        /// 判断属性是否表示“法兰，对夹”连接方式。
        /// </summary>
        private static bool ShouldQueryFlangeStandard(IDictionary<string, string> properties)
        {
            // 没有属性时不能确认图元类型
            if (properties == null || properties.Count == 0) return false;

            // 兼容中文标题、英文 Tag 和历史字段名称
            string connectionMode = FindProperty(
                properties,
                "连接方式",
                "CONN_TYPE",
                "DNCONN_TYPE",
                "CONNECTION_MODE",
                "CONNECTIONTYPE",
                "连接形式",
                "连接型式",
                "连接类别");

            // 去除常见分隔符后，所有图元统一接受业务明确允许的四种连接方式。
            string normalizedMode = NormalizeConnectionMode(connectionMode);

            // 所有图元只有在这四种连接方式下才需要查询法兰连接规范。
            return string.Equals(normalizedMode, "法兰", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalizedMode, "法兰连接", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalizedMode, "对夹", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalizedMode, "对夹连接", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 检查块定义的 Tag 或 Prompt 是否声明了“法兰、对夹”连接方式。
        /// </summary>
        private static bool HasFlangeConnectionDefinition(DBTrans tr, BlockReference blockReference)
        {
            // 参数无效时无法判断
            if (tr == null || blockReference == null) return false;

            // 获取块定义
            var blockDefinition = tr.GetObject(blockReference.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            if (blockDefinition == null || !blockDefinition.HasAttributeDefinitions) return false;

            // 遍历 AttributeDefinition，同时检查 Tag 和 Prompt
            foreach (ObjectId entityId in blockDefinition)
            {
                if (!(tr.GetObject(entityId, OpenMode.ForRead) is AttributeDefinition attributeDefinition)) continue;

                string tag = attributeDefinition.Tag ?? string.Empty;
                string prompt = attributeDefinition.Prompt ?? string.Empty;

                // 连接方式字段可能使用 CONN_TYPE、DNCONN_TYPE 或中文 Tag。
                bool isConnectionField = string.Equals(
                    NormalizePropertyKey(tag),
                    NormalizePropertyKey("CONN_TYPE"),
                    StringComparison.Ordinal) ||
                    string.Equals(
                        NormalizePropertyKey(tag),
                        NormalizePropertyKey("DNCONN_TYPE"),
                        StringComparison.Ordinal) ||
                    string.Equals(
                        NormalizePropertyKey(tag),
                        NormalizePropertyKey("连接方式"),
                        StringComparison.Ordinal);
                bool hasFlangeStandardField = string.Equals(
                    NormalizePropertyKey(tag),
                    NormalizePropertyKey("FLG_STD"),
                    StringComparison.Ordinal) ||
                    string.Equals(
                        NormalizePropertyKey(prompt),
                        NormalizePropertyKey("法兰标准"),
                        StringComparison.Ordinal);

                string normalizedPrompt = NormalizeConnectionMode(prompt);
                string normalizedDefaultValue = NormalizeConnectionMode(attributeDefinition.TextString);
                bool isFlangeClampMode = normalizedPrompt.IndexOf("法兰", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                         normalizedPrompt.IndexOf("对夹", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                         normalizedDefaultValue.IndexOf("法兰", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                         normalizedDefaultValue.IndexOf("对夹", StringComparison.OrdinalIgnoreCase) >= 0;

                if ((isConnectionField || hasFlangeStandardField) && isFlangeClampMode)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 统一清理连接方式中的空白和常见分隔符。
        /// </summary>
        private static string NormalizeConnectionMode(string value)
        {
            return (value ?? string.Empty)
                .Replace(" ", string.Empty)
                .Replace("　", string.Empty)
                .Replace(",", string.Empty)
                .Replace("，", string.Empty)
                .Replace("、", string.Empty)
                .Replace("；", string.Empty)
                .Replace(";", string.Empty)
                .Replace(":", string.Empty)
                .Replace("：", string.Empty)
                .Replace("/", string.Empty)
                .Replace("\\", string.Empty);
        }

        /// <summary>
        /// 在属性字典中按指定的键名列表查找第一个非空属性值。
        /// 支持键名标准化匹配（忽略大小写、去除空格等），提高查找灵活性。
        /// </summary>
        /// <param name="properties">源属性字典，键为属性名称，值为属性值。</param>
        /// <param name="keys">要查找的键名列表，按顺序匹配；返回第一个匹配成功的非空值。</param>
        /// <returns>匹配到的属性值（已去除首尾空格）；若未找到任何匹配或值均为空，则返回 null。</returns>
        private static string FindProperty(IDictionary<string, string> properties, params string[] keys)
        {
            // 遍历所有候选键名（按传入顺序）
            foreach (string key in keys)
            {
                // 将当前候选键标准化（例如转大写、去除空格），以便进行不区分大小写的比较
                string normalizedKey = NormalizePropertyKey(key);

                // 遍历字典中的每个属性条目
                foreach (KeyValuePair<string, string> property in properties)
                {
                    // 若字典键经标准化后与候选键相同，且对应的值不为空或空白
                    if (string.Equals(NormalizePropertyKey(property.Key), normalizedKey, StringComparison.Ordinal) &&
                        !string.IsNullOrWhiteSpace(property.Value))
                    {
                        // 返回去除首尾空格的属性值
                        return property.Value.Trim();
                    }
                }
            }

            // 所有键均未匹配到有效值，返回 null
            return null;
        }
        /// <summary>
        /// 标准化公称通径（DN）字符串，统一格式为 "DN" 后接数字（如 "DN100"）。
        /// </summary>
        /// <param name="value">原始通径字符串，可能包含空格、大小写混合或纯数字。</param>
        /// <returns>标准化后的 DN 字符串；若输入为纯数字则转为 "DN{数字}"，否则保留原始格式（去除空格并转大写）。</returns>
        private static string NormalizeDn(string value)
        {
            // 去除首尾空格、转大写、移除所有空格
            string normalized = (value ?? string.Empty).Trim().ToUpperInvariant().Replace(" ", string.Empty);

            // 若已以 "DN" 开头，则直接返回（已标准化）
            if (normalized.StartsWith("DN", StringComparison.Ordinal)) return normalized;

            // 若输入为纯数字（如 "100"），则转换为 "DN100"
            return int.TryParse(normalized, out int number) ? $"DN{number}" : normalized;
        }
        /// <summary>
        /// 标准化公称压力（PN）字符串，统一格式为 "PN" 后接数字（如 "PN16"），保留小数形式（如 "PN2.5"）。
        /// </summary>
        /// <param name="value">原始压力字符串，可能包含空格、大小写混合或纯数字。</param>
        /// <returns>标准化后的 PN 字符串；若输入为数字则转为 "PN{数字}"（使用不变文化保留小数点），否则保留原始格式（去除空格并转大写）。</returns>
        private static string NormalizePn(string value)
        {
            // 去除首尾空格、转大写、移除所有空格
            string normalized = (value ?? string.Empty).Trim().ToUpperInvariant().Replace(" ", string.Empty);

            // 若已以 "PN" 开头，则直接返回
            if (normalized.StartsWith("PN", StringComparison.Ordinal)) return normalized;

            // 若输入为数字（支持小数），则转换为 "PN{数字}"，使用不变文化确保小数点符号统一
            return decimal.TryParse(normalized, out decimal number)
                ? $"PN{number.ToString(System.Globalization.CultureInfo.InvariantCulture)}"
                : normalized;
        }
        /// <summary>
        /// 标准化钢管系列标识，统一为 "Ⅰ系列" 或 "Ⅱ系列"。
        /// </summary>
        /// <param name="value">原始系列字符串，可能包含 "II"、"Ⅱ"、"2" 等变体。</param>
        /// <returns>标准化后的系列字符串："Ⅱ系列"（若输入为2系列）或 "Ⅰ系列"（默认）。</returns>
        private static string NormalizeSeries(string value)
        {
            // 去除首尾空格
            string normalized = (value ?? string.Empty).Trim();

            // 若包含 "Ⅱ"、"II" 或等于 "2"，则视为Ⅱ系列，否则默认Ⅰ系列
            return normalized.Contains("Ⅱ") || normalized.Contains("II") || normalized == "2"
                ? "Ⅱ系列"
                : "Ⅰ系列";
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

        [CommandMethod(nameof(GB_InsertBlock_5))]
        public static void GB_InsertBlock_5()
        {
            try // 整个方法主 try，捕获并记录异常
            {
                pointS.Clear(); // 清空坐标集合，准备存储新插入块的坐标
                var uiScale = AutoCadHelper.GetScale(); // 获取当前图纸的比例，作为块的默认插入比例
                var plScale = 0.3 * uiScale; // 多段线宽度缩放比例，基于 UI 比例计算
                Directory.CreateDirectory(GetPath.referenceFile); // 确保参考文件目录存在
                if (VariableDictionary.btnFileName == null) return; // 判断点现的按键名是不是空；
                if (VariableDictionary.resourcesFile == null) return; // 判断点现的原文件是不是空；

                using var tr = new DBTrans(); // 使用事务包装对图形数据库的修改

                // 获取对应块的 ObjectId（从外部资源 DWG 中）
                var referenceFileObId = tr.BlockTable.GetBlockFormA(
                    VariableDictionary.resourcesFile,
                    VariableDictionary.btnFileName,
                    VariableDictionary.btnFileName_blockName,
                    true);

                var refFileRec = tr.GetObject(referenceFileObId, OpenMode.ForRead) as BlockTableRecord; // 读取块表记录
                if (refFileRec == null)
                {
                    LogManager.Instance.LogInfo("未找到块记录！"); // 日志：未找到块记录
                    return; // 退出
                }

                LogManager.Instance.LogInfo("块！"); // 日志：进入块插入循环
                while (true) // 循环插入，直到用户取消
                {
                    // 把块插入到当前空间
                    var referenceFileBlock = tr.CurrentSpace.InsertBlock(Point3d.Origin, referenceFileObId);

                    // 检查是否为实体
                    if (tr.GetObject(referenceFileBlock) is not Entity referenceFileEntity)
                        return; // 若不是实体则退出

                    // 设置图层和颜色等属性
                    referenceFileEntity.Layer = VariableDictionary.btnBlockLayer; // 指定图层
                    referenceFileEntity.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex); // 指定颜色索引
                    referenceFileEntity.Scale(new Point3d(0, 0, 0), uiScale / 100); // 按 UI 比例缩放块

                    var startPoint = new Point3d(0, 0, 0); // 拖拽起始点（本地变量）

                    var jigBlock = new JigEx((mpw, _) =>
                    {
                        referenceFileEntity.Move(startPoint, mpw); // 将实体从上次位置移动到当前鼠标位置
                        startPoint = mpw; // 更新起始位置为当前点，便于下次增量移动
                    });

                    jigBlock.DatabaseEntityDraw(wd => wd.Geometry.Draw(referenceFileEntity)); // 绘制拖拽预览
                    jigBlock.SetOptions(msg: "\n指定插入点"); // 设置提示信息

                    // 执行拖拽交互
                    var endPoint = Env.Editor.Drag(jigBlock);
                    if (endPoint.Status != PromptStatus.OK)
                    {
                        // 用户取消插入，则删除已插入的块并退出循环
                        tr.GetObject(referenceFileBlock, OpenMode.ForWrite);
                        referenceFileBlock.Erase();
                        break;
                    }

                    // 存储插入点坐标（WCS）
                    var UcsEndPoint = jigBlock.MousePointWcsLast; // 获取最后一次鼠标 WCS 点
                    pointS.Add(UcsEndPoint); // 添加到点集合
                    Env.Editor.Redraw(); // 刷新视图
                }

                // ======================
                // 在此处根据插入数量绘图
                // ======================
                int count = pointS.Count; // 获取插入点数量
                LogManager.Instance.LogInfo($"\n已插入 {count} 个块，开始绘制外围图形..."); // 日志

                if (count == 3)
                { // 三点生成外接“圆”的多段线形式（使用带 bulge 的 polyline 生成平滑圆弧）
                    Point3d p1 = pointS[0]; // 第一点
                    Point3d p2 = pointS[1]; // 第二点
                    Point3d p3 = pointS[2]; // 第三点

                    // 计算三角形外接圆圆心（与三点等距的点）
                    Point3d circleCenter = GetCircumcenter(p1, p2, p3); // 复用已有方法计算圆心

                    // 计算圆心到某一点的距离作为半径，并向外扩展 1.5 * uiScale（与旧逻辑保持一致）
                    double radius = p1.DistanceTo(circleCenter) + 1.5 * uiScale; // 半径计算

                    // 记录到日志，以验证计算正确性
                    double dist1 = circleCenter.DistanceTo(p1); // 距离1
                    double dist2 = circleCenter.DistanceTo(p2); // 距离2
                    double dist3 = circleCenter.DistanceTo(p3); // 距离3
                    LogManager.Instance.LogInfo($"\n圆心到三点的距离: {dist1:F4}, {dist2:F4}, {dist3:F4}"); // 打印距离

                    // 使用 Polyline（2D）并用 bulge 值创建若干弧段以近似平滑圆
                    var plCircle = new Polyline(); // 新建 Polyline 对象（2D）
                    int segments = 72; // 分段数（越大越光滑，性能开销也越大）
                    double delta = 2.0 * Math.PI / segments; // 每段对应的角度增量
                    double bulge = Math.Tan(delta / 4.0); // 对应弧段的 bulge 值（tan(Δ/4)）

                    // 起始角度可以由第一个点方向决定，也可以固定为 0；这里以 0 开始，生成完整闭合圆
                    double startAngle = 0.0; // 起始角度

                    // 循环添加顶点并指定 bulge，使得每个段为圆弧
                    for (int i = 0; i < segments; i++)
                    {
                        double angle = startAngle + i * delta; // 当前角度
                        double x = circleCenter.X + radius * Math.Cos(angle); // 计算顶点 X
                        double y = circleCenter.Y + radius * Math.Sin(angle); // 计算顶点 Y
                        var pt2d = new Point2d(x, y); // 构造 Point2d 坐标
                        plCircle.AddVertexAt(i, pt2d, bulge, plScale, plScale); // 添加顶点并设置 bulge 与宽度
                    }

                    plCircle.Closed = true; // 闭合多段线（形成完整环）
                    plCircle.Layer = VariableDictionary.btnBlockLayer; // 设置图层
                    plCircle.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex); // 设置颜色索引

                    // 将多段线添加到当前空间（使用事务 tr）
                    tr.CurrentSpace.AddEntity(plCircle); // 添加实体到图纸
                    Env.Editor.Redraw(); // 刷新显示
                    LogManager.Instance.LogInfo("\n已创建外围多段线圆，使用 bulge 生成平滑弧段。"); // 日志说明
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
                        Point3d expandedPoint3d = point + dirVector * 1.5 * uiScale;
                        Point2d expandedPoint = new Point2d(expandedPoint3d.X, expandedPoint3d.Y);
                        expandedPoints.Add(expandedPoint);
                    }

                    // 确保点按顺时针或逆时针排序
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
                        pl.AddVertexAt(i, sortedPoints[i], 0, plScale, plScale);
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
                        Point3d expandedPoint3d = point + dirVector * 1.5 * uiScale;
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
                        polygon.AddVertexAt(i, sortedPoints[i], 0, plScale, plScale);
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

                // 添加标注（若有）
                var dimColorLine = VariableDictionary.layerColorIndex;
                if (VariableDictionary.btnFileName.Contains("结构"))
                {
                    dimColorLine = 3;
                }
                if (VariableDictionary.dimString != null)
                    Command.DDimLinear(tr, VariableDictionary.dimString, count.ToString(), Convert.ToInt16(dimColorLine)); // 创建标注

                tr.Commit(); // 提交事务
                Env.Editor.Redraw(); // 刷新显示
                LogManager.Instance.LogInfo("\n操作完成。"); // 日志：完成
                pointS.Clear(); // 清理点集合
            }
            catch (Exception ex) // 捕获方法级别异常并记录
            {
                LogManager.Instance.LogInfo($"\n插入图元失败：{ex.Message}"); // 日志错误信息
                LogManager.Instance.LogInfo($"\n错误详情：{ex.StackTrace}"); // 日志堆栈信息
            }
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
