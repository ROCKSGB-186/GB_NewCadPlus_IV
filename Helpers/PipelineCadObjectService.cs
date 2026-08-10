using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using GB_NewCadPlus_IV.FunctionalMethod;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 管道 CAD 落图结果。
    /// </summary>
    public sealed class PipelineCadPlacementResult
    {
        /// <summary>业务管道 ID。</summary>
        public string PipeId { get; set; } = string.Empty;

        /// <summary>管道主体 Polyline。</summary>
        public ObjectId PipelineObjectId { get; set; } = ObjectId.Null;

        /// <summary>隐藏属性块。</summary>
        public ObjectId AttributeCarrierObjectId { get; set; } = ObjectId.Null;

        /// <summary>标题文字对象。</summary>
        public ObjectId TitleObjectId { get; set; } = ObjectId.Null;

        /// <summary>流向符号对象。</summary>
        public ObjectId FlowDirectionObjectId { get; set; } = ObjectId.Null;
    }

    /// <summary>
    /// 管道 Polyline、属性、标题和流向符号落图服务。
    /// </summary>
    public static class PipelineCadObjectService
    {
        private sealed class PipelineSegment
        {
            public int Index { get; set; }
            public Point3d Start { get; set; }
            public Point3d End { get; set; }
        }

        /// <summary>
        /// 将一条管道草稿作为正式 CAD 对象写入当前空间。
        /// </summary>
        public static PipelineCadPlacementResult PlacePipeline(
            Database database,
            IList<Point3d> points,
            IDictionary<string, string> sourceAttributes,
            string pipeRole)
        {
            if (database == null)
            {
                throw new ArgumentNullException(nameof(database));
            }

            if (points == null || points.Count < 2)
            {
                throw new ArgumentException("管道至少需要两个点。", nameof(points));
            }

            Dictionary<string, string> attributes = new Dictionary<string, string>(
                sourceAttributes ?? new Dictionary<string, string>(),
                StringComparer.OrdinalIgnoreCase);
            string normalizedRole = NormalizeRole(pipeRole);
            string pipeId = Guid.NewGuid().ToString("N");
            attributes["PIPEID"] = pipeId;
            attributes["PIPE_ROLE"] = normalizedRole;
            attributes["PIPELINETITLE"] = PipelineTitleBuilder.Build(attributes);
            attributes["PIPE_LENGTH"] = CalculateLength(points).ToString("0.###", CultureInfo.InvariantCulture);

            Stopwatch stopwatch = Stopwatch.StartNew();
            LogManager.Instance.LogInfo(
                $"[管道落图][开始] PipeId={pipeId}, Role={normalizedRole}, PointCount={points.Count}, AttributeCount={attributes.Count}, Title={attributes["PIPELINETITLE"]}");
            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                LogManager.Instance.LogInfo("[管道落图][事务开始] 正在打开当前空间。");
                BlockTableRecord currentSpace = transaction.GetObject(
                    database.CurrentSpaceId,
                    OpenMode.ForWrite) as BlockTableRecord;
                if (currentSpace == null)
                {
                    throw new InvalidOperationException("当前图纸空间不可写。\n");
                }

                Polyline pipeline = CreatePipelinePolyline(points, normalizedRole);
                currentSpace.AppendEntity(pipeline);
                transaction.AddNewlyCreatedDBObject(pipeline, true);
                LogManager.Instance.LogInfo(
                    $"[管道落图][Polyline完成] ObjectId={pipeline.ObjectId}, VertexCount={pipeline.NumberOfVertices}");

                WriteAttributesToExtensionDictionary(transaction, pipeline, attributes);
                LogManager.Instance.LogInfo(
                    $"[管道落图][主体属性完成] ObjectId={pipeline.ObjectId}, AttributeCount={attributes.Count}");

                ObjectId attributeCarrierId = CreateHiddenAttributeCarrier(
                    database,
                    transaction,
                    currentSpace,
                    points[0],
                    attributes,
                    pipeId);
                attributes["ATTRIBUTE_CARRIER_HANDLE"] = GetHandle(attributeCarrierId);
                WriteAttributesToExtensionDictionary(transaction, pipeline, attributes);
                LogManager.Instance.LogInfo(
                    $"[管道落图][属性载体完成] ObjectId={attributeCarrierId}, Handle={attributes["ATTRIBUTE_CARRIER_HANDLE"]}");

                List<PipelineSegment> displaySegments = GetDisplaySegments(points);
                List<ObjectId> titleIds = new List<ObjectId>();
                List<ObjectId> flowIds = new List<ObjectId>();
                if (displaySegments.Count == 0)
                {
                    LogManager.Instance.LogInfo(
                        $"[管道落图][标注跳过] PipeId={pipeId}, Reason=没有长度大于50×比例的有效线段。");
                }
                foreach (PipelineSegment segment in displaySegments)
                {
                    DBText title = CreateTitleText(
                        transaction,
                        database,
                        segment.Start,
                        segment.End,
                        attributes["PIPELINETITLE"],
                        normalizedRole);
                    currentSpace.AppendEntity(title);
                    transaction.AddNewlyCreatedDBObject(title, true);
                    WriteLinkRecord(transaction, title, pipeId, "TITLE");
                    titleIds.Add(title.ObjectId);
                    LogManager.Instance.LogInfo(
                        $"[管道落图][标题完成] SegmentIndex={segment.Index}, ObjectId={title.ObjectId}, Text={title.TextString}, Position={title.Position}, Height={title.Height}, Layer={title.Layer}, ColorIndex={title.ColorIndex}");

                    Solid flowFill;
                    Polyline flowDirection = CreateFlowDirectionSymbol(segment.Start, segment.End, normalizedRole, out flowFill);
                    currentSpace.AppendEntity(flowDirection);
                    transaction.AddNewlyCreatedDBObject(flowDirection, true);
                    WriteLinkRecord(transaction, flowDirection, pipeId, "FLOW_DIRECTION");
                    flowIds.Add(flowDirection.ObjectId);
                    if (flowFill != null)
                    {
                        currentSpace.AppendEntity(flowFill);
                        transaction.AddNewlyCreatedDBObject(flowFill, true);
                        WriteLinkRecord(transaction, flowFill, pipeId, "FLOW_DIRECTION_FILL");
                    }
                    LogManager.Instance.LogInfo(
                        $"[管道落图][流向完成] SegmentIndex={segment.Index}, ObjectId={flowDirection.ObjectId}, VertexCount={flowDirection.NumberOfVertices}, Layer={flowDirection.Layer}, ColorIndex={flowDirection.ColorIndex}");
                }

                LogManager.Instance.LogInfo(
                    $"[管道落图][标注完成] SegmentCount={displaySegments.Count}, TitleCount={titleIds.Count}, FlowCount={flowIds.Count}");

                LogManager.Instance.LogInfo("[管道落图][事务提交开始]");
                transaction.Commit();
                stopwatch.Stop();

                LogManager.Instance.LogInfo(
                    $"[管道落图][事务提交成功] PipeId={pipeId}, Role={normalizedRole}, PipelineObjectId={pipeline.ObjectId}, AttributeCarrierObjectId={attributeCarrierId}, TitleCount={titleIds.Count}, FirstTitleObjectId={(titleIds.Count == 0 ? ObjectId.Null : titleIds[0])}, FlowCount={flowIds.Count}, FirstFlowObjectId={(flowIds.Count == 0 ? ObjectId.Null : flowIds[0])}, ElapsedMs={stopwatch.ElapsedMilliseconds}");

                return new PipelineCadPlacementResult
                {
                    PipeId = pipeId,
                    PipelineObjectId = pipeline.ObjectId,
                    AttributeCarrierObjectId = attributeCarrierId,
                    TitleObjectId = titleIds.Count == 0 ? ObjectId.Null : titleIds[0],
                    FlowDirectionObjectId = flowIds.Count == 0 ? ObjectId.Null : flowIds[0]
                };
            }
        }

        /// <summary>
        /// 创建支持夹点编辑的二维多段线主体。
        /// </summary>
        private static Polyline CreatePipelinePolyline(IList<Point3d> points, string pipeRole)
        {
            Polyline pipeline = new Polyline();
            pipeline.SetDatabaseDefaults();
            double scale = GetDisplayScale();
            for (int index = 0; index < points.Count; index++)
            {
                Point3d point = points[index];
                pipeline.AddVertexAt(index, new Point2d(point.X, point.Y), 0, 0.3 * scale, 0.3 * scale);
            }

            pipeline.ConstantWidth = 0.3 * scale;
            pipeline.Color = Color.FromColorIndex(ColorMethod.ByAci, GetPipeColor(pipeRole));

            return pipeline;
        }

        /// <summary>
        /// 创建隐藏属性块，保留现有 AttributeReference 读取兼容性。
        /// </summary>
        private static ObjectId CreateHiddenAttributeCarrier(
            Database database,
            Transaction transaction,
            BlockTableRecord currentSpace,
            Point3d position,
            IDictionary<string, string> attributes,
            string pipeId)
        {
            BlockTable blockTable = transaction.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;
            if (blockTable == null)
            {
                return ObjectId.Null;
            }

            string blockName = $"GB_PIPE_ATTR_{pipeId}";
            blockTable.UpgradeOpen();
            BlockTableRecord definition = new BlockTableRecord
            {
                Name = blockName
            };
            ObjectId definitionId = blockTable.Add(definition);
            transaction.AddNewlyCreatedDBObject(definition, true);

            foreach (KeyValuePair<string, string> attribute in attributes)
            {
                if (string.IsNullOrWhiteSpace(attribute.Key))
                {
                    continue;
                }

                AttributeDefinition definitionAttribute = new AttributeDefinition
                {
                    Position = Point3d.Origin,
                    Tag = PipelineCadPropertyKeyHelper.Encode(attribute.Key.Trim()),
                    Prompt = string.Empty,
                    TextString = attribute.Value ?? string.Empty,
                    Height = 0.1,
                    Invisible = true
                };
                definitionAttribute.SetDatabaseDefaults();
                definition.AppendEntity(definitionAttribute);
                transaction.AddNewlyCreatedDBObject(definitionAttribute, true);
            }

            BlockReference carrier = new BlockReference(position, definitionId);
            currentSpace.AppendEntity(carrier);
            transaction.AddNewlyCreatedDBObject(carrier, true);

            foreach (ObjectId entityId in definition)
            {
                AttributeDefinition definitionAttribute = transaction.GetObject(
                    entityId,
                    OpenMode.ForRead) as AttributeDefinition;
                if (definitionAttribute == null || definitionAttribute.Constant)
                {
                    continue;
                }

                AttributeReference attributeReference = new AttributeReference();
                attributeReference.SetAttributeFromBlock(definitionAttribute, carrier.BlockTransform);
                attributeReference.TextString = definitionAttribute.TextString;
                carrier.AttributeCollection.AppendAttribute(attributeReference);
                transaction.AddNewlyCreatedDBObject(attributeReference, true);
            }

            return carrier.ObjectId;
        }

        /// <summary>
        /// 创建图面标题文字。
        /// </summary>
        private static DBText CreateTitleText(
            Transaction transaction,
            Database database,
            Point3d start,
            Point3d end,
            string title,
            string pipeRole)
        {
            Vector3d direction = end - start;
            if (direction.Length <= 1e-8)
            {
                direction = Vector3d.XAxis;
            }

            direction = direction.GetNormal();
            Vector3d perpendicular = new Vector3d(-direction.Y, direction.X, 0).GetNormal();
            if (perpendicular.DotProduct(Vector3d.YAxis) < 0)
            {
                perpendicular = -perpendicular;
            }

            double scale = GetDisplayScale();
            double titleHeight = 3.5 * scale;
            Point3d segmentMiddle = new Point3d(
                (start.X + end.X) / 2.0,
                (start.Y + end.Y) / 2.0,
                (start.Z + end.Z) / 2.0);
            double titleOffset = 4.0 * scale + titleHeight * 0.75;
            Point3d position = segmentMiddle + perpendicular * titleOffset;

            double textRotation = Math.Atan2(direction.Y, direction.X);
            if (Math.Cos(textRotation) < 0)
            {
                textRotation += Math.PI;
            }

            DBText text = new DBText();
            text.SetDatabaseDefaults();
            text.Position = position;
            text.TextString = title ?? string.Empty;
            text.Height = titleHeight;
            text.Rotation = textRotation;
            text.Justify = AttachmentPoint.MiddleCenter;
            text.AlignmentPoint = position;
            text.HorizontalMode = TextHorizontalMode.TextCenter;
            text.VerticalMode = TextVerticalMode.TextVerticalMid;
            text.Color = Color.FromColorIndex(ColorMethod.ByAci, GetTitleColor(pipeRole));
            try
            {
                using (DBTrans styleTransaction = new DBTrans())
                {
                    TextFontsStyleHelper.ApplyTitleToDBText(styleTransaction, text, scale);
                    text.AdjustAlignment(database);
                    styleTransaction.Commit();
                }
            }
            catch
            {
                // 保留基础文字属性，避免字体样式异常影响管道落图。
            }
            LogManager.Instance.LogInfo(
                $"[管道落图][标题尺寸] Scale={scale}, BaseHeight=3.5, ActualHeight={text.Height}");
            return text;
        }

        /// <summary>
        /// 创建简单三角形流向符号，方向与管道首尾方向一致。
        /// </summary>
        private static Polyline CreateFlowDirectionSymbol(
            Point3d start,
            Point3d end,
            string pipeRole,
            out Solid fill)
        {
            fill = null;
            Vector3d direction = end - start;
            if (direction.Length <= 1e-8)
            {
                direction = Vector3d.XAxis;
            }

            direction = direction.GetNormal();
            Vector3d perpendicular = new Vector3d(-direction.Y, direction.X, 0).GetNormal();
            Point3d center = new Point3d(
                (start.X + end.X) / 2.0,
                (start.Y + end.Y) / 2.0,
                (start.Z + end.Z) / 2.0);
            double scale = GetDisplayScale();
            double length = 10.0 * scale;
            double height = 2.0 * scale;
            Point3d tip = center + direction * (length / 2.0);
            Point3d left = center - direction * (length / 2.0) - perpendicular * (height / 2.0);
            Point3d right = center - direction * (length / 2.0) + perpendicular * (height / 2.0);

            Polyline arrow = new Polyline();
            arrow.SetDatabaseDefaults();
            arrow.AddVertexAt(0, new Point2d(tip.X, tip.Y), 0, 0, 0);
            arrow.AddVertexAt(1, new Point2d(left.X, left.Y), 0, 0, 0);
            arrow.AddVertexAt(2, new Point2d(right.X, right.Y), 0, 0, 0);
            arrow.Closed = true;
            arrow.Color = Color.FromColorIndex(ColorMethod.ByAci, GetPipeColor(pipeRole));
            arrow.LineWeight = LineWeight.LineWeight025;

            fill = new Solid(
                tip,
                left,
                right,
                right)
            {
                Color = Color.FromColorIndex(ColorMethod.ByAci, GetArrowColor(pipeRole)),
                LineWeight = LineWeight.LineWeight025
            };
            LogManager.Instance.LogInfo(
                $"[管道落图][流向尺寸] Scale={scale}, BaseLength=10, BaseHeight=2, ActualLength={length}, ActualHeight={height}");
            return arrow;
        }

        private static List<PipelineSegment> GetDisplaySegments(IList<Point3d> points)
        {
            double minimumSegmentLength = 50.0 * GetDisplayScale();
            List<PipelineSegment> segments = new List<PipelineSegment>();
            for (int index = 0; index < points.Count - 1; index++)
            {
                if (points[index].DistanceTo(points[index + 1]) > minimumSegmentLength)
                {
                    segments.Add(new PipelineSegment
                    {
                        Index = index,
                        Start = points[index],
                        End = points[index + 1]
                    });
                }
            }

            return segments;
        }

        /// <summary>
        /// 获取项目统一 CAD 显示比例，异常或无效值使用 1，避免生成零尺寸图元。
        /// </summary>
        private static double GetDisplayScale()
        {
            double scale = AutoCadHelper.GetScale(true);
            return double.IsNaN(scale) || double.IsInfinity(scale) || scale <= 0
                ? 1.0
                : scale;
        }

        /// <summary>
        /// 向实体扩展字典写入单个字符串属性。
        /// </summary>
        private static void WriteAttributesToExtensionDictionary(
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
                $"[管道落图][属性字典准备] ObjectId={entity.ObjectId}, StorageKey={PipelineCadPropertyKeyHelper.StorageKey}, PairCount={values.Count / 2}");
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
        /// 给标题或流向符号写入 PipeId 和对象用途关联。
        /// </summary>
        private static void WriteLinkRecord(
            Transaction transaction,
            Entity entity,
            string pipeId,
            string objectRole)
        {
            WriteAttributesToExtensionDictionary(
                transaction,
                entity,
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["PIPEID"] = pipeId,
                    ["OBJECT_ROLE"] = objectRole
                });
        }

        /// <summary>
        /// 计算多段线总长度。
        /// </summary>
        private static double CalculateLength(IList<Point3d> points)
        {
            double length = 0;
            for (int index = 1; index < points.Count; index++)
            {
                length += points[index - 1].DistanceTo(points[index]);
            }

            return length;
        }

        /// <summary>
        /// 规范化管道角色。
        /// </summary>
        private static string NormalizeRole(string pipeRole)
        {
            return string.Equals(pipeRole, "EXPORT", StringComparison.OrdinalIgnoreCase)
                ? "EXPORT"
                : "IMPORT";
        }

        /// <summary>
        /// 根据管道返回标题和流向符号颜色。
        /// </summary>      

        private static short GetPipeColor(string pipeRole)
        {
            return 3;
        }
        /// <summary>
        /// 获取标题文字颜色，EXPORT 为红色，IMPORT 为蓝色。
        /// </summary>
        /// <param name="pipeRole"> 管道角色 </param>
        /// <returns> 标题文字颜色 </returns>
        private static short GetTitleColor(string pipeRole)
        {
            return NormalizeRole(pipeRole) == "EXPORT" ? (short)2 : (short)1;
        }

        private static short GetArrowColor(string pipeRole)
        {
            return 6;
        }

        /// <summary>
        /// 获取对象句柄文本。
        /// </summary>
        private static string GetHandle(ObjectId objectId)
        {
            return objectId == ObjectId.Null ? string.Empty : objectId.Handle.ToString();
        }
    }
}
