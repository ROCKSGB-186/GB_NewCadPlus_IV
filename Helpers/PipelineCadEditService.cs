using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 管道 CAD 属性回写服务。
    /// </summary>
    public static class PipelineCadEditService
    {
        /// <summary>
        /// 将参数页面确认后的属性写回管道主体、属性载体和标题对象。
        /// </summary>
        public static IReadOnlyList<ObjectId> UpdatePipeline(
            Database database,
            ObjectId pipelineObjectId,
            IDictionary<string, string> editedAttributes,
            IDictionary<string, string> changedAttributes)
        {
            if (database == null)
            {
                throw new ArgumentNullException(nameof(database));
            }

            if (pipelineObjectId == ObjectId.Null)
            {
                throw new ArgumentException("管道对象 ID 无效。", nameof(pipelineObjectId));
            }

            Dictionary<string, string> attributes = new Dictionary<string, string>(
                editedAttributes ?? new Dictionary<string, string>(),
                StringComparer.OrdinalIgnoreCase);

            var flangeComponentIds = new List<ObjectId>();
            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                Polyline pipeline = transaction.GetObject(
                    pipelineObjectId,
                    OpenMode.ForWrite) as Polyline;
                if (pipeline == null)
                {
                    throw new InvalidOperationException("选择的对象不是管道 Polyline。\n");
                }

                string pipeId = GetAttribute(attributes, "PIPEID");
                if (string.IsNullOrWhiteSpace(pipeId))
                {
                    pipeId = GetAttribute(PipelineEndpointPropertyHelper.ReadEntityProperties(transaction, pipeline), "PIPEID");
                }

                if (string.IsNullOrWhiteSpace(pipeId))
                {
                    throw new InvalidOperationException("管道缺少 PIPEID，无法建立关联对象。\n");
                }

                string pipeRole = GetAttribute(attributes, "PIPE_ROLE");
                if (string.IsNullOrWhiteSpace(pipeRole))
                {
                    pipeRole = GetAttribute(PipelineEndpointPropertyHelper.ReadEntityProperties(transaction, pipeline), "PIPE_ROLE");
                }

                attributes["PIPEID"] = pipeId;
                attributes["PIPE_ROLE"] = pipeRole;
                double currentLength = GetPolylineLength(pipeline);
                attributes["PIPE_LENGTH"] = currentLength
                    .ToString("0.###", CultureInfo.InvariantCulture);
                LogManager.Instance.LogInfo(
                    $"[管道编辑][长度重算] PipeId={pipeId}, CurrentLength={attributes["PIPE_LENGTH"]}, VertexCount={pipeline.NumberOfVertices}");
                attributes["PIPELINETITLE"] = PipelineTitleBuilder.Build(attributes);
                WriteXRecords(transaction, pipeline, attributes);
                string tagNo = GetTagNo(attributes);

                BlockTableRecord currentSpace = transaction.GetObject(
                    database.CurrentSpaceId,
                    OpenMode.ForRead) as BlockTableRecord;
                if (currentSpace != null)
                {
                    foreach (ObjectId objectId in currentSpace)
                    {
                        Entity entity = transaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                        if (entity == null || entity.ObjectId == pipelineObjectId)
                        {
                            continue;
                        }

                        Dictionary<string, string> linkProperties =
                            PipelineEndpointPropertyHelper.ReadEntityProperties(transaction, entity);
                        if (!string.Equals(GetAttribute(linkProperties, "PIPEID"), pipeId, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        string objectRole = GetAttribute(linkProperties, "OBJECT_ROLE");
                        if (string.Equals(objectRole, "TITLE", StringComparison.OrdinalIgnoreCase) && entity is DBText title)
                        {
                            title.TextString = attributes["PIPELINETITLE"];
                            title.Height = 3.5 * GetDisplayScale();
                            title.Color = Color.FromColorIndex(
                                ColorMethod.ByAci,
                                GetTitleColor(pipeRole));
                            LogManager.Instance.LogInfo(
                                $"[管道标题同步] PipeId={pipeId}, NewTitle={title.TextString}, Height={title.Height}, TargetObjectId={title.ObjectId}");
                        }
                        else if (string.Equals(objectRole, "FLOW_DIRECTION", StringComparison.OrdinalIgnoreCase))
                        {
                            entity.Color = Color.FromColorIndex(
                                ColorMethod.ByAci,
                                GetRoleColor(pipeRole));
                        }
                        else if (entity is BlockReference carrier)
                        {
                            UpdateAttributeCarrier(transaction, carrier, attributes, pipeId);
                        }
                    }

                    int synchronizedComponentCount = SynchronizeSameTagNoComponents(
                        transaction,
                        currentSpace,
                        pipelineObjectId,
                        tagNo,
                        changedAttributes,
                        flangeComponentIds);
                    LogManager.Instance.LogInfo(
                        $"[管段属性批量同步][完成] PipeId={pipeId}, TagNo={tagNo}, ChangedFieldCount={changedAttributes?.Count ?? 0}, ComponentCount={synchronizedComponentCount}");
                }
                transaction.Commit();
            }

            return flangeComponentIds;
        }

        /// <summary>
        /// 按唯一管段号同步当前空间内的非管道部件；只更新本次实际修改且部件已存在的字段。
        /// </summary>
        private static int SynchronizeSameTagNoComponents(
            Transaction transaction,
            BlockTableRecord currentSpace,
            ObjectId pipelineObjectId,
            string tagNo,
            IDictionary<string, string> changedAttributes,
            ICollection<ObjectId> flangeComponentIds)
        {
            if (string.IsNullOrWhiteSpace(tagNo) || changedAttributes == null || changedAttributes.Count == 0)
            {
                LogManager.Instance.LogInfo($"[管段属性批量同步][跳过] TagNo={tagNo}, 原因=管段号为空或本次没有实际修改字段。");
                return 0;
            }

            int componentCount = 0;
            foreach (ObjectId objectId in currentSpace)
            {
                if (objectId == pipelineObjectId)
                {
                    continue;
                }

                Entity entity = transaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                // 所有 Polyline 均为管道或图线，按规则不参与部件属性同步。
                if (entity == null || entity.IsErased || entity is Polyline)
                {
                    continue;
                }

                Dictionary<string, string> componentProperties =
                    PipelineEndpointPropertyHelper.ReadEntityProperties(transaction, entity);
                if (!string.Equals(GetTagNo(componentProperties), tagNo, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                int updatedCount = UpdateExistingComponentProperties(
                    transaction,
                    entity,
                    changedAttributes,
                    tagNo);
                if (updatedCount > 0)
                {
                    componentCount++;
                    LogManager.Instance.LogInfo(
                        $"[管段属性批量同步][部件完成] TagNo={tagNo}, TargetObjectId={entity.ObjectId}, EntityType={entity.GetType().Name}, UpdatedCount={updatedCount}");
                }

                if (updatedCount > 0 &&
                    HasFlangeSpecificationChange(changedAttributes) &&
                    entity is BlockReference flangeBlock &&
                    IsFlangeSpecificationComponent(componentProperties, flangeBlock))
                {
                    flangeComponentIds?.Add(entity.ObjectId);
                    LogManager.Instance.LogInfo(
                        $"[管段法兰规范刷新][收集] TagNo={tagNo}, TargetObjectId={entity.ObjectId}");
                }
            }

            return componentCount;
        }

        /// <summary>
        /// 更新部件已有的块属性和统一属性 Xrecord 键，不新增任何字段。
        /// </summary>
        private static int UpdateExistingComponentProperties(
            Transaction transaction,
            Entity entity,
            IDictionary<string, string> changedAttributes,
            string tagNo)
        {
            int updatedCount = 0;
            if (entity is BlockReference blockReference)
            {
                foreach (ObjectId attributeId in blockReference.AttributeCollection)
                {
                    AttributeReference attribute = transaction.GetObject(attributeId, OpenMode.ForWrite) as AttributeReference;
                    if (attribute == null || !TryGetChangedValue(changedAttributes, PipelineCadPropertyKeyHelper.Decode(attribute.Tag), out string newValue))
                    {
                        continue;
                    }

                    string oldValue = attribute.TextString ?? string.Empty;
                    if (string.Equals(oldValue, newValue, StringComparison.Ordinal)) continue;
                    attribute.TextString = newValue;
                    updatedCount++;
                    LogManager.Instance.LogInfo($"[管段属性批量同步][AttributeReference] TagNo={tagNo}, Tag={attribute.Tag}, OldValue={oldValue}, NewValue={newValue}, TargetObjectId={entity.ObjectId}");
                }
            }

            if (entity.ExtensionDictionary == ObjectId.Null)
            {
                return updatedCount;
            }

            DBDictionary dictionary = transaction.GetObject(entity.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
            if (dictionary == null || !dictionary.Contains(PipelineCadPropertyKeyHelper.StorageKey))
            {
                return updatedCount;
            }

            Xrecord record = transaction.GetObject(dictionary.GetAt(PipelineCadPropertyKeyHelper.StorageKey), OpenMode.ForWrite) as Xrecord;
            TypedValue[] values = record?.Data?.AsArray();
            if (values == null) return updatedCount;
            bool recordChanged = false;
            for (int index = 0; index + 1 < values.Length; index += 2)
            {
                string businessTag = values[index].Value?.ToString() ?? string.Empty;
                if (!TryGetChangedValue(changedAttributes, businessTag, out string newValue)) continue;
                string oldValue = values[index + 1].Value?.ToString() ?? string.Empty;
                if (string.Equals(oldValue, newValue, StringComparison.Ordinal)) continue;
                values[index + 1] = new TypedValue((int)DxfCode.Text, newValue);
                updatedCount++;
                recordChanged = true;
                LogManager.Instance.LogInfo($"[管段属性批量同步][Xrecord] TagNo={tagNo}, Tag={businessTag}, OldValue={oldValue}, NewValue={newValue}, TargetObjectId={entity.ObjectId}");
            }
            if (recordChanged)
            {
                record.Data = new ResultBuffer(values);
            }

            return updatedCount;
        }

        /// <summary>
        /// 更新隐藏属性块中的 AttributeReference，并记录每个实际赋值的 Tag。
        /// </summary>
        private static void UpdateAttributeCarrier(
            Transaction transaction,
            BlockReference carrier,
            IDictionary<string, string> attributes,
            string pipeId)
        {
            foreach (ObjectId attributeId in carrier.AttributeCollection)
            {
                AttributeReference attribute = transaction.GetObject(
                    attributeId,
                    OpenMode.ForWrite) as AttributeReference;
                string businessTag = attribute == null
                    ? string.Empty
                    : PipelineCadPropertyKeyHelper.Decode(attribute.Tag);
                if (attribute == null || !attributes.TryGetValue(businessTag, out string newValue))
                {
                    continue;
                }

                string oldValue = attribute.TextString ?? string.Empty;
                newValue = newValue ?? string.Empty;
                if (string.Equals(oldValue, newValue, StringComparison.Ordinal))
                {
                    continue;
                }

                attribute.TextString = newValue;
                LogManager.Instance.LogInfo(
                    $"[管道属性同步][AttributeReference] PipeId={pipeId}, Tag={businessTag}, OldValue={oldValue}, NewValue={newValue}, TargetObjectId={carrier.ObjectId}");
            }
        }

        /// <summary>
        /// 将属性字典保存到实体扩展字典。
        /// </summary>
        private static void WriteXRecords(
            Transaction transaction,
            Entity entity,
            IDictionary<string, string> attributes)
        {
            if (entity.ExtensionDictionary == ObjectId.Null)
            {
                entity.CreateExtensionDictionary();
            }

            DBDictionary dictionary = transaction.GetObject(
                entity.ExtensionDictionary,
                OpenMode.ForWrite) as DBDictionary;
            if (dictionary == null)
            {
                return;
            }

            List<TypedValue> values = new List<TypedValue>();
            foreach (KeyValuePair<string, string> attribute in attributes)
            {
                if (string.IsNullOrWhiteSpace(attribute.Key))
                {
                    continue;
                }

                values.Add(new TypedValue((int)DxfCode.Text, attribute.Key.Trim()));
                values.Add(new TypedValue((int)DxfCode.Text, attribute.Value ?? string.Empty));
            }

            LogManager.Instance.LogInfo(
                $"[管道编辑][属性字典准备] ObjectId={entity.ObjectId}, StorageKey={PipelineCadPropertyKeyHelper.StorageKey}, PairCount={values.Count / 2}");
            if (dictionary.Contains(PipelineCadPropertyKeyHelper.StorageKey))
            {
                dictionary.Remove(PipelineCadPropertyKeyHelper.StorageKey);
            }

            Xrecord record = new Xrecord
            {
                Data = new ResultBuffer(values.ToArray())
            };
            dictionary.SetAt(PipelineCadPropertyKeyHelper.StorageKey, record);
            transaction.AddNewlyCreatedDBObject(record, true);
        }

        private static bool HasFlangeSpecificationChange(IDictionary<string, string> changedAttributes)
        {
            return TryGetChangedValue(changedAttributes, "DN", out _) ||
                   TryGetChangedValue(changedAttributes, "PN", out _);
        }

        private static bool IsFlangeConnection(IDictionary<string, string> attributes)
        {
            string connectionMode = FindProperty(
                attributes,
                "CONN_TYPE",
                "CONNTYPE",
                "DNCONN_TYPE",
                "CONNECTION_MODE",
                "连接方式",
                "连接形式");

            string normalizedMode = NormalizeBusinessTag(connectionMode);
            return normalizedMode == "法兰" || normalizedMode == "法兰连接" ||
                   normalizedMode == "对夹" || normalizedMode == "对夹连接";
        }

        /// <summary>
        /// 识别需要刷新法兰规范的部件：连接方式优先，其次判断已有法兰规范 Tag 或块名特征。
        /// 独立法兰和管端盲板通常没有 CONN_TYPE，但会保留法兰尺寸/标准属性。
        /// </summary>
        private static bool IsFlangeSpecificationComponent(
            IDictionary<string, string> attributes,
            BlockReference blockReference)
        {
            if (IsFlangeConnection(attributes))
            {
                return true;
            }

            bool hasStandard = !string.IsNullOrWhiteSpace(FindProperty(
                attributes,
                "FLG_STD",
                "DRAWINGNO.STANDARDNO",
                "法兰标准"));
            bool hasDimensions = !string.IsNullOrWhiteSpace(FindProperty(
                attributes,
                "FLG_OD",
                "BOLT_PCD",
                "FLG_THK",
                "BOLT_SPEC"));
            if (hasStandard && hasDimensions)
            {
                return true;
            }

            string blockName = blockReference?.Name ?? string.Empty;
            return blockName.IndexOf("法兰", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   blockName.IndexOf("盲板", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   blockName.IndexOf("FLANGE", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   blockName.IndexOf("BLIND", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   blockName.IndexOf("_FL_", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// 对管道修改后受影响的法兰块重新查询规范库，并写入完整尺寸规范属性。
        /// 网络查询在单独事务外执行，避免长时间占用 AutoCAD 数据库事务。
        /// </summary>
        public static void RefreshFlangeStandards(Database database, IEnumerable<ObjectId> flangeComponentIds)
        {
            if (database == null || flangeComponentIds == null)
            {
                return;
            }

            foreach (ObjectId flangeObjectId in flangeComponentIds.Distinct())
            {
                Dictionary<string, string> properties;
                using (Transaction transaction = database.TransactionManager.StartTransaction())
                {
                    BlockReference flange = transaction.GetObject(flangeObjectId, OpenMode.ForRead) as BlockReference;
                    if (flange == null || flange.IsErased)
                    {
                        continue;
                    }

                    properties = PipelineEndpointPropertyHelper.ReadEntityProperties(transaction, flange);
                    transaction.Commit();
                }

                string dn = FindProperty(properties, "DN", "公称通径", "通径");
                string pn = FindProperty(properties, "PN", "公称压力", "压力等级");
                string connectionMode = FindProperty(properties, "CONN_TYPE", "CONNTYPE", "连接方式", "连接形式");
                if (string.IsNullOrWhiteSpace(dn) || string.IsNullOrWhiteSpace(pn))
                {
                    LogManager.Instance.LogWarning($"[管段法兰规范刷新][跳过] TargetObjectId={flangeObjectId}, DN={dn}, PN={pn}, ConnectionMode={connectionMode}");
                    continue;
                }

                if (string.IsNullOrWhiteSpace(connectionMode))
                {
                    connectionMode = "法兰连接";
                }

                string familyCode = FindProperty(properties, "FAMILY_CODE", "规范大类编码", "FAMILY");
                string series = FindProperty(properties, "SERIES", "钢管系列", "管道系列");
                var request = new FlangeStandardMatchRequest
                {
                    FamilyCode = string.IsNullOrWhiteSpace(familyCode) ? "FLANGE" : familyCode,
                    SeriesCode = FindProperty(properties, "SERIES_CODE", "规范系列编码"),
                    StandardNumber = FindProperty(properties, "FLG_STD", "DRAWINGNO.STANDARDNO", "法兰标准", "标准号"),
                    TableNumber = FindProperty(properties, "TABLE_NUMBER", "标准表号", "表号"),
                    DN = dn,
                    PN = pn,
                    Series = string.IsNullOrWhiteSpace(series) ? "Ⅰ系列" : series,
                    ConnectionMode = connectionMode,
                    // 动态规范系列未维护类型/密封面元数据时，PL/RF 可能只是服务器上次回写的默认显示值。
                    // 刷新时不将其作为强制筛选条件，避免把有效 DN/PN 行排除。
                    FlangeType = string.Empty,
                    FaceType = string.Empty
                };

                LogManager.Instance.LogInfo(
                    $"[管段法兰规范刷新][请求] TargetObjectId={flangeObjectId}, FamilyCode={request.FamilyCode}, SeriesCode={request.SeriesCode}, StandardNumber={request.StandardNumber}, TableNumber={request.TableNumber}, DN={request.DN}, PN={request.PN}, Series={request.Series}, ConnectionMode={request.ConnectionMode}, FlangeType={request.FlangeType}, FaceType={request.FaceType}");

                try
                {
                    FlangeStandardMatchResponse response = new StandardApiService()
                        .MatchFlangeAsync(request)
                        .GetAwaiter()
                        .GetResult();
                    if (response == null || !response.Success)
                    {
                        LogManager.Instance.LogWarning($"[管段法兰规范刷新][未命中] TargetObjectId={flangeObjectId}, DN={request.DN}, PN={request.PN}, Message={response?.Message}");
                        continue;
                    }

                    using (Transaction transaction = database.TransactionManager.StartTransaction())
                    {
                        BlockReference flange = transaction.GetObject(flangeObjectId, OpenMode.ForWrite) as BlockReference;
                        if (flange == null || flange.IsErased)
                        {
                            continue;
                        }

                        int updatedCount = new StandardPropertySyncService()
                            .ApplyToBlockReference(transaction, flange, response, createMissingAttributes: false);
                        transaction.Commit();
                        LogManager.Instance.LogInfo($"[管段法兰规范刷新][完成] TargetObjectId={flangeObjectId}, DN={request.DN}, PN={request.PN}, UpdatedCount={updatedCount}");
                    }
                }
                catch (Exception exception)
                {
                    LogManager.Instance.LogWarning($"[管段法兰规范刷新][失败] TargetObjectId={flangeObjectId}, DN={request.DN}, PN={request.PN}, Error={exception.Message}");
                }
            }
        }

        /// <summary>
        /// 计算当前管道 Polyline 的长度。
        /// </summary>
        public static double GetPolylineLength(Polyline pipeline)
        {
            if (pipeline == null)
            {
                return 0;
            }

            double length = 0;
            for (int index = 1; index < pipeline.NumberOfVertices; index++)
            {
                length += pipeline.GetPoint3dAt(index - 1).DistanceTo(pipeline.GetPoint3dAt(index));
            }

            return length;
        }

        /// <summary>
        /// 获取项目统一显示比例，避免比例无效时生成零高度标题。
        /// </summary>
        private static double GetDisplayScale()
        {
            double scale = AutoCadHelper.GetScale(true);
            return double.IsNaN(scale) || double.IsInfinity(scale) || scale <= 0
                ? 1.0
                : scale;
        }

        /// <summary>
        /// 按不区分大小写的 Tag 读取属性。
        /// </summary>
        private static string GetAttribute(
            IDictionary<string, string> attributes,
            string tag)
        {
            return attributes != null && attributes.TryGetValue(tag, out string value)
                ? value ?? string.Empty
                : string.Empty;
        }

        /// <summary>
        /// 按多个候选业务字段名读取属性，兼容下划线、空格与大小写差异。
        /// </summary>
        private static string FindProperty(IDictionary<string, string> attributes, params string[] tags)
        {
            if (attributes == null || tags == null) return string.Empty;
            foreach (string tag in tags)
            {
                string normalizedTag = NormalizeBusinessTag(tag);
                foreach (KeyValuePair<string, string> attribute in attributes)
                {
                    if (NormalizeBusinessTag(attribute.Key) == normalizedTag)
                    {
                        return attribute.Value ?? string.Empty;
                    }
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// 读取管段号，兼容 TAG_NO、TAGNO 和中文管段号字段。
        /// </summary>
        private static string GetTagNo(IDictionary<string, string> attributes)
        {
            if (attributes == null) return string.Empty;
            foreach (KeyValuePair<string, string> attribute in attributes)
            {
                string normalizedTag = NormalizeBusinessTag(attribute.Key);
                if (normalizedTag == "TAGNO" || normalizedTag == "管段号")
                {
                    return (attribute.Value ?? string.Empty).Trim();
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// 按业务字段名获取本次变更值，兼容下划线、空格及编码 Tag 差异。
        /// </summary>
        private static bool TryGetChangedValue(
            IDictionary<string, string> changedAttributes,
            string businessTag,
            out string value)
        {
            value = string.Empty;
            if (changedAttributes == null || string.IsNullOrWhiteSpace(businessTag)) return false;
            string normalizedTarget = NormalizeBusinessTag(PipelineCadPropertyKeyHelper.Decode(businessTag));
            foreach (KeyValuePair<string, string> changedAttribute in changedAttributes)
            {
                if (NormalizeBusinessTag(changedAttribute.Key) != normalizedTarget) continue;
                value = changedAttribute.Value ?? string.Empty;
                return true;
            }

            return false;
        }

        /// <summary>
        /// 统一业务字段名用于比较，保留中文字符并忽略下划线、空格和大小写差异。
        /// </summary>
        private static string NormalizeBusinessTag(string tag)
        {
            return (tag ?? string.Empty)
                .Replace("_", string.Empty)
                .Replace(" ", string.Empty)
                .Trim()
                .ToUpperInvariant();
        }

        /// <summary>
        /// 根据角色返回标题和流向颜色。
        /// </summary>
        private static short GetRoleColor(string pipeRole)
        {
            return string.Equals(pipeRole, "EXPORT", StringComparison.OrdinalIgnoreCase)
                ? (short)2
                : (short)3;
        }

        /// <summary>
        /// 获取管道标题颜色，必须与管道初次落图时保持一致。
        /// </summary>
        private static short GetTitleColor(string pipeRole)
        {
            return string.Equals(pipeRole, "EXPORT", StringComparison.OrdinalIgnoreCase)
                ? (short)2
                : (short)1;
        }
    }
}
