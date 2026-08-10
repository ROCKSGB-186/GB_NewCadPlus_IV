using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using GB_NewCadPlus_IV.FunctionalMethod;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 管道端点附近图元及其属性读取结果。
    /// </summary>
    public sealed class PipelineEndpointSource
    {
        /// <summary>命中的图元 ObjectId。</summary>
        public ObjectId ObjectId { get; set; } = ObjectId.Null;

        /// <summary>命中的图元句柄。</summary>
        public string Handle { get; set; } = string.Empty;

        /// <summary>命中的图元类型。</summary>
        public string EntityType { get; set; } = string.Empty;

        /// <summary>用于业务显示的图元名称。</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>命中位置。</summary>
        public Point3d HitPoint { get; set; } = Point3d.Origin;

        /// <summary>点到图元的距离。</summary>
        public double Distance { get; set; } = double.MaxValue;

        /// <summary>图元 AttributeReference 和 XRecord 属性。</summary>
        public Dictionary<string, string> Attributes { get; set; }
            = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 起点和终点属性合并结果。
    /// </summary>
    public sealed class PipelineEndpointMergeResult
    {
        /// <summary>合并后的属性。</summary>
        public Dictionary<string, string> Attributes { get; set; }
            = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>起点属性冲突字段。</summary>
        public List<string> StartEndConflicts { get; set; } = new List<string>();
    }

    /// <summary>
    /// 管道端点图元识别和属性读取辅助类。
    /// </summary>
    public static class PipelineEndpointPropertyHelper
    {
        /// <summary>
        /// 查找当前图纸空间中距离指定点最近且在容差内的图元。
        /// </summary>
        public static PipelineEndpointSource FindNearestEntity(
            Database database,
            Point3d point,
            double tolerance)
        {
            PipelineEndpointSource result = new PipelineEndpointSource();
            if (database == null || tolerance <= 0)
            {
                return result;
            }

            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                BlockTableRecord currentSpace = transaction.GetObject(
                    database.CurrentSpaceId,
                    OpenMode.ForRead) as BlockTableRecord;
                if (currentSpace == null)
                {
                    return result;
                }

                foreach (ObjectId objectId in currentSpace)
                {
                    Entity entity = transaction.GetObject(objectId, OpenMode.ForRead) as Entity;
                    if (entity == null || entity.IsErased)
                    {
                        continue;
                    }

                    if (!TryGetClosestPoint(entity, point, tolerance, out Point3d hitPoint, out double distance))
                    {
                        continue;
                    }

                    if (distance >= result.Distance)
                    {
                        continue;
                    }

                    result = ReadEndpointSource(transaction, entity, hitPoint, distance);
                }

                transaction.Commit();
            }

            return result.ObjectId == ObjectId.Null ? null : result;
        }

        /// <summary>
        /// 从实体读取块属性和 XRecord 属性。
        /// </summary>
        public static Dictionary<string, string> ReadEntityProperties(
            Transaction transaction,
            Entity entity)
        {
            Dictionary<string, string> properties =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (transaction == null || entity == null)
            {
                return properties;
            }

            if (entity is BlockReference blockReference)
            {
                foreach (ObjectId attributeId in blockReference.AttributeCollection)
                {
                    AttributeReference attribute = transaction.GetObject(
                        attributeId,
                        OpenMode.ForRead) as AttributeReference;
                    if (attribute == null || string.IsNullOrWhiteSpace(attribute.Tag))
                    {
                        continue;
                    }

                    string businessTag = PipelineCadPropertyKeyHelper.Decode(attribute.Tag.Trim());
                    properties[businessTag] = attribute.TextString ?? string.Empty;
                }
            }

            if (entity.ExtensionDictionary != ObjectId.Null)
            {
                DBDictionary extensionDictionary = transaction.GetObject(
                    entity.ExtensionDictionary,
                    OpenMode.ForRead) as DBDictionary;
                if (extensionDictionary != null)
                {
                    foreach (DBDictionaryEntry entry in extensionDictionary)
                    {
                        Xrecord record = transaction.GetObject(
                            entry.Value,
                            OpenMode.ForRead) as Xrecord;
                        if (record?.Data == null)
                        {
                            continue;
                        }

                        TypedValue[] values = record.Data.AsArray();
                        if (string.Equals(
                            entry.Key,
                            PipelineCadPropertyKeyHelper.StorageKey,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            for (int index = 0; index + 1 < values.Length; index += 2)
                            {
                                string businessTag = values[index].Value?.ToString() ?? string.Empty;
                                if (!string.IsNullOrWhiteSpace(businessTag))
                                {
                                    properties[businessTag.Trim()] = values[index + 1].Value?.ToString() ?? string.Empty;
                                }
                            }
                        }
                        else if (values.Length > 0 && !string.IsNullOrWhiteSpace(entry.Key))
                        {
                            string businessTag = PipelineCadPropertyKeyHelper.Decode(entry.Key.Trim());
                            properties[businessTag] = values[0].Value?.ToString() ?? string.Empty;
                        }
                    }
                }
            }

            return properties;
        }

        /// <summary>
        /// 合并默认值、起点属性和终点属性。
        /// 起点和终点只写入 START_POINT/END_POINT，不覆盖管道 NAME。
        /// </summary>
        public static PipelineEndpointMergeResult MergeEndpointProperties(
            IDictionary<string, string> defaults,
            PipelineEndpointSource startSource,
            PipelineEndpointSource endSource)
        {
            PipelineEndpointMergeResult result = new PipelineEndpointMergeResult();
            CopyProperties(result.Attributes, defaults);
            MergeSourceProperties(result.Attributes, startSource?.Attributes, result.StartEndConflicts, "起点");
            MergeSourceProperties(result.Attributes, endSource?.Attributes, result.StartEndConflicts, "终点");

            if (startSource != null && !string.IsNullOrWhiteSpace(startSource.Name))
            {
                result.Attributes["START_POINT"] = startSource.Name;
            }

            if (endSource != null && !string.IsNullOrWhiteSpace(endSource.Name))
            {
                result.Attributes["END_POINT"] = endSource.Name;
            }

            return result;
        }

        /// <summary>
        /// 判断实体与指定点的最近距离。
        /// </summary>
        private static bool TryGetClosestPoint(
            Entity entity,
            Point3d point,
            double tolerance,
            out Point3d closestPoint,
            out double distance)
        {
            closestPoint = Point3d.Origin;
            distance = double.MaxValue;

            try
            {
                if (entity is Curve curve)
                {
                    closestPoint = curve.GetClosestPointTo(point, false);
                    distance = closestPoint.DistanceTo(point);
                    return distance <= tolerance;
                }

                Extents3d extents = entity.GeometricExtents;
                if (point.X < extents.MinPoint.X - tolerance || point.X > extents.MaxPoint.X + tolerance ||
                    point.Y < extents.MinPoint.Y - tolerance || point.Y > extents.MaxPoint.Y + tolerance ||
                    point.Z < extents.MinPoint.Z - tolerance || point.Z > extents.MaxPoint.Z + tolerance)
                {
                    return false;
                }

                closestPoint = point;
                distance = 0;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 将实体和属性转换为端点来源对象。
        /// </summary>
        private static PipelineEndpointSource ReadEndpointSource(
            Transaction transaction,
            Entity entity,
            Point3d hitPoint,
            double distance)
        {
            Dictionary<string, string> properties = ReadEntityProperties(transaction, entity);
            string name = FindProperty(properties, "NAME", "名称");

            return new PipelineEndpointSource
            {
                ObjectId = entity.ObjectId,
                Handle = entity.Handle.ToString(),
                EntityType = entity.GetRXClass()?.Name ?? entity.GetType().Name,
                Name = name,
                HitPoint = hitPoint,
                Distance = distance,
                Attributes = properties
            };
        }

        /// <summary>
        /// 合并一个端点来源，空值不覆盖已有值，非空冲突进入提示列表。
        /// </summary>
        private static void MergeSourceProperties(
            IDictionary<string, string> target,
            IDictionary<string, string> source,
            ICollection<string> conflicts,
            string sourceName)
        {
            if (source == null)
            {
                return;
            }

            foreach (KeyValuePair<string, string> pair in source)
            {
                if (string.Equals(pair.Key, "NAME", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(pair.Key, "名称", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string sourceValue = pair.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(sourceValue))
                {
                    continue;
                }

                if (!target.TryGetValue(pair.Key, out string targetValue) || string.IsNullOrWhiteSpace(targetValue))
                {
                    target[pair.Key] = sourceValue;
                    continue;
                }

                if (!string.Equals(targetValue, sourceValue, StringComparison.OrdinalIgnoreCase))
                {
                    conflicts.Add($"{sourceName}:{pair.Key}={sourceValue} 与已有值 {targetValue} 不一致");
                }
            }
        }

        /// <summary>
        /// 复制默认属性并统一使用不区分大小写字典。
        /// </summary>
        private static void CopyProperties(
            IDictionary<string, string> target,
            IDictionary<string, string> source)
        {
            if (source == null)
            {
                return;
            }

            foreach (KeyValuePair<string, string> pair in source)
            {
                if (!string.IsNullOrWhiteSpace(pair.Key))
                {
                    target[pair.Key.Trim()] = pair.Value ?? string.Empty;
                }
            }
        }

        /// <summary>
        /// 按多个候选 Tag 读取名称。
        /// </summary>
        private static string FindProperty(
            IDictionary<string, string> properties,
            params string[] keys)
        {
            if (properties == null)
            {
                return string.Empty;
            }

            foreach (string key in keys)
            {
                if (properties.TryGetValue(key, out string value) && !string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }

            return string.Empty;
        }
    }
}
