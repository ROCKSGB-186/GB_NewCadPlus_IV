using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using GB_NewCadPlus_IV.FunctionalMethod;
using System;
using System.Collections.Generic;
using System.Globalization;

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
        public static void UpdatePipeline(
            Database database,
            ObjectId pipelineObjectId,
            IDictionary<string, string> editedAttributes)
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
                }

                transaction.Commit();
            }
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
