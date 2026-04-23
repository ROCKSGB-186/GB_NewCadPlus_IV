using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_IV.UniFiedStandards
{
    /// <summary>
    /// 兼容补丁：补齐被截断文件中缺失的方法，先保证工程可编译与主流程可用。
    /// </summary>
    public partial class UnifiedTableGenerator
    {
        /// <summary>
        /// 兼容补丁：设备表生成（简化实现）。
        /// </summary>
        private void CreateDeviceTable(Database db, List<DeviceInfo> deviceList, double scaleDenominator = 0.0)
        {
            // 中文注释：防御空数据，避免空引用。
            if (db == null || deviceList == null || deviceList.Count == 0) return;

            // 中文注释：直接复用带类型标题的方法，统一逻辑。
            CreateDeviceTableWithType(db, deviceList, "设备", scaleDenominator);
        }

        /// <summary>
        /// 兼容补丁：按类型生成设备表（简化实现）。
        /// </summary>
        private void CreateDeviceTableWithType(Database db, List<DeviceInfo> deviceList, string typeTitle, double scaleDenominator = 0.0)
        {
            if (db == null || deviceList == null || deviceList.Count == 0) return;

            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            // 中文注释：让用户指定插入点。
            var ppr = ed.GetPoint($"\n'{typeTitle}'表：指定插入位置:");
            if (ppr.Status != PromptStatus.OK) return;

            // 中文注释：动态收集字段列，优先常用字段。
            var preferred = new List<string> { "名称", "规格", "材质", "数量", "压力等级", "公称直径DN", "备注" };
            var allKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in deviceList)
            {
                if (d?.Attributes == null) continue;
                foreach (var k in d.Attributes.Keys)
                {
                    if (!string.IsNullOrWhiteSpace(k)) allKeys.Add(k.Trim());
                }
            }
            var cols = preferred.Where(k => allKeys.Contains(k)).ToList();
            cols.AddRange(allKeys.Where(k => !cols.Contains(k)).OrderBy(k => k, StringComparer.OrdinalIgnoreCase));
            if (!cols.Contains("名称")) cols.Insert(0, "名称");
            if (!cols.Contains("数量")) cols.Add("数量");

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // 中文注释：标题行+表头行+数据行。
                int rows = 2 + deviceList.Count;
                int columns = Math.Max(1, cols.Count);
                var table = new Table();
                table.SetSize(rows, columns);
                table.Position = ppr.Value;

                // 中文注释：标题行。
                table.Cells[0, 0].TextString = $"{typeTitle} - 材料明细表";
                if (columns > 1)
                {
                    table.MergeCells(CellRange.Create(table, 0, 0, 0, columns - 1));
                }

                // 中文注释：表头行。
                for (int c = 0; c < columns; c++)
                {
                    var k = cols[c];
                    table.Cells[1, c].TextString = k;
                }

                // 中文注释：数据行。
                for (int r = 0; r < deviceList.Count; r++)
                {
                    var item = deviceList[r];
                    for (int c = 0; c < columns; c++)
                    {
                        var key = cols[c];
                        string val = string.Empty;

                        if (string.Equals(key, "名称", StringComparison.OrdinalIgnoreCase))
                        {
                            val = item?.Name ?? string.Empty;
                        }
                        else if (string.Equals(key, "数量", StringComparison.OrdinalIgnoreCase))
                        {
                            if (item?.Attributes != null && item.Attributes.TryGetValue("数量", out var q) && !string.IsNullOrWhiteSpace(q))
                                val = q;
                            else
                                val = (item?.Count ?? 0).ToString();
                        }
                        else
                        {
                            if (item?.Attributes != null && item.Attributes.TryGetValue(key, out var raw))
                                val = raw ?? string.Empty;
                        }

                        table.Cells[r + 2, c].TextString = val;
                    }
                }

                // 中文注释：简单列宽自适应，避免表格太挤。
                for (int c = 0; c < columns; c++)
                {
                    table.SetColumnWidth(c, 18.0);
                }

                currentSpace.AppendEntity(table);
                tr.AddNewlyCreatedDBObject(table, true);
                tr.Commit();
            }
        }

        /// <summary>
        /// 兼容补丁：清洗属性文本（去掉换行和首尾空格）。
        /// </summary>
        private string CleanAttributeText(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            return value.Replace("\r", " ").Replace("\n", " ").Trim();
        }

        /// <summary>
        /// 兼容补丁：提供比例分母读取，避免旧调用报错。
        /// </summary>
        private double GetScaleDenominatorForDatabase(Database db, bool roundToCommon = false)
        {
            try
            {
                // 中文注释：优先使用当前系统比例计算。
                return AutoCadHelper.GetScale();
            }
            catch
            {
                return 1.0;
            }
        }

        /// <summary>
        /// 兼容补丁：分析示例块并提取模板。
        /// </summary>
        private SamplePipeInfo AnalyzeSampleBlock(Transaction tr, BlockReference blockRef)
        {
            var info = new SamplePipeInfo();
            if (tr == null || blockRef == null) return info;

            var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            if (btr == null) return info;
            info.BasePoint = btr.Origin;

            var polylines = new List<Polyline>();
            foreach (ObjectId id in btr)
            {
                var obj = tr.GetObject(id, OpenMode.ForRead);
                if (obj is Polyline pl)
                {
                    var c = pl.Clone() as Polyline;
                    if (c != null) polylines.Add(c);
                }
                else if (obj is AttributeDefinition ad)
                {
                    var c = ad.Clone() as AttributeDefinition;
                    if (c != null) info.AttributeDefinitions.Add(c);
                }
                else if (obj is Solid solid)
                {
                    // 中文注释：优先记录模板中实体填充色，用于箭头填充色同步。
                    if (info.DirectionArrowFillColor == null)
                    {
                        info.DirectionArrowFillColor = solid.Color;
                    }
                }
                else if (obj is Hatch hatch)
                {
                    // 中文注释：若存在 Hatch 填充，也可作为模板填充颜色来源。
                    if (info.DirectionArrowFillColor == null)
                    {
                        info.DirectionArrowFillColor = hatch.Color;
                    }
                }
            }

            if (polylines.Count > 0)
            {
                polylines = polylines.OrderByDescending(p => p.Length).ToList();
                info.PipeBodyTemplate = polylines[0];
                info.DirectionArrowTemplate = polylines.FirstOrDefault(p => p.Closed && p.NumberOfVertices == 3);
            }

            return info;
        }

        /// <summary>
        /// 兼容补丁：收集线段。
        /// </summary>
        private List<LineSegmentInfo> CollectLineSegments(Transaction tr, List<ObjectId> ids)
        {
            var segments = new List<LineSegmentInfo>();
            foreach (var id in ids)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead);
                if (ent is Line line)
                {
                    segments.Add(ProcessLine(line, tr));
                }
                else if (ent is Polyline pl)
                {
                    for (int i = 0; i < pl.NumberOfVertices - 1; i++)
                    {
                        if (pl.GetSegmentType(i) == SegmentType.Line)
                        {
                            var p1 = pl.GetPoint3dAt(i);
                            var p2 = pl.GetPoint3dAt(i + 1);
                            var vec = p2 - p1;
                            segments.Add(new LineSegmentInfo
                            {
                                StartPoint = p1,
                                EndPoint = p2,
                                Length = vec.Length,
                                Angle = vec.GetAngleTo(Vector3d.XAxis, Vector3d.ZAxis),
                                Layer = pl.Layer,
                                ColorIndex = pl.ColorIndex,
                                LinetypeScale = pl.LinetypeScale,
                                EntityType = "POLYLINE_SEGMENT"
                            });
                        }
                    }
                }
            }
            return segments;
        }
        //private List<LineSegmentInfo> CollectLineSegments(Transaction tr, List<ObjectId> ids)
        //{
        //    var segs = new List<LineSegmentInfo>();
        //    if (tr == null || ids == null) return segs;

        //    foreach (var id in ids)
        //    {
        //        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
        //        if (ent is Line ln)
        //        {
        //            var v = ln.EndPoint - ln.StartPoint;
        //            segs.Add(new LineSegmentInfo
        //            {
        //                Id = id,
        //                StartPoint = ln.StartPoint,
        //                EndPoint = ln.EndPoint,
        //                Length = v.Length,
        //                Angle = v.GetAngleTo(Vector3d.XAxis, Vector3d.ZAxis),
        //                Layer = ln.Layer,
        //                ColorIndex = ln.ColorIndex,
        //                LinetypeScale = ln.LinetypeScale,
        //                EntityType = "LINE"
        //            });
        //        }
        //        else if (ent is Polyline pl)
        //        {
        //            for (int i = 0; i < pl.NumberOfVertices - 1; i++)
        //            {
        //                if (pl.GetSegmentType(i) != SegmentType.Line) continue;
        //                var p1 = pl.GetPoint3dAt(i);
        //                var p2 = pl.GetPoint3dAt(i + 1);
        //                var v = p2 - p1;
        //                segs.Add(new LineSegmentInfo
        //                {
        //                    Id = id,
        //                    StartPoint = p1,
        //                    EndPoint = p2,
        //                    Length = v.Length,
        //                    Angle = v.GetAngleTo(Vector3d.XAxis, Vector3d.ZAxis),
        //                    Layer = pl.Layer,
        //                    ColorIndex = pl.ColorIndex,
        //                    LinetypeScale = pl.LinetypeScale,
        //                    EntityType = "POLYLINE_SEGMENT"
        //                });
        //            }
        //        }
        //    }
        //    return segs;
        //}

        /// <summary>
        /// 新增：根据首尾相连的线段集合，按连通顺序构建连续顶点列表（起点、每个连接点、终点）
        /// </summary>
        /// <param name="segments">线段集合</param>
        /// <param name="tol">容差</param>
        /// <returns></returns>

        private List<Point3d> BuildOrderedVerticesFromSegments(List<LineSegmentInfo> segments, double tol = 1e-6)
        {
            var result = new List<Point3d>();// 结果顶点列表
            if (segments == null || segments.Count == 0) return result;

            // 比较两点是否相等（使用容差）
            static bool PointsEqual(Point3d a, Point3d b, double tol)
            {
                return Math.Abs(a.X - b.X) <= tol && Math.Abs(a.Y - b.Y) <= tol && Math.Abs(a.Z - b.Z) <= tol;
            }
            // 构建唯一点列表并统计度数（出现次数）
            var uniquePoints = new List<Point3d>();
            Func<Point3d, int> getIndex = p =>
            {
                for (int i = 0; i < uniquePoints.Count; i++)
                {
                    if (PointsEqual(uniquePoints[i], p, tol)) return i;
                }
                uniquePoints.Add(p);
                return uniquePoints.Count - 1;
            };
            // 构建索引列表
            var counts = new List<int>();
            var segPairs = new List<(int s, int e)>();
            foreach (var seg in segments)
            {
                var si = getIndex(seg.StartPoint);
                var ei = getIndex(seg.EndPoint);
                segPairs.Add((si, ei));

                // ensure counts capacity
                while (counts.Count < uniquePoints.Count) counts.Add(0);
                counts[si]++;
                counts[ei]++;
            }
            // 找到链的端点：度为1的点（非闭合链）
            int startPointIndex = -1;
            for (int i = 0; i < counts.Count; i++)
            {
                if (counts[i] == 1)
                {
                    startPointIndex = i;
                    break;
                }
            }
            // 若都是度 >=2（闭合回路或多分支），退回到第一个段的起点
            if (startPointIndex == -1)
            {
                startPointIndex = segPairs.Count > 0 ? segPairs[0].s : 0;
            }
            // 从 startPointIndex 开始按链遍历段
            var visited = new bool[segPairs.Count];
            Point3d current = uniquePoints[startPointIndex];
            result.Add(current);
            bool progressed;
            do
            {
                progressed = false;
                for (int i = 0; i < segPairs.Count; i++)
                {
                    if (visited[i]) continue;
                    var (si, ei) = segPairs[i];
                    if (PointsEqual(uniquePoints[si], current, tol))
                    {
                        // forward
                        var next = uniquePoints[ei];
                        if (!PointsEqual(next, result.Last(), tol))
                            result.Add(next);
                        current = next;
                        visited[i] = true;
                        progressed = true;
                        break;
                    }
                    else if (PointsEqual(uniquePoints[ei], current, tol))
                    {
                        // reverse
                        var next = uniquePoints[si];
                        if (!PointsEqual(next, result.Last(), tol))
                            result.Add(next);
                        current = next;
                        visited[i] = true;
                        progressed = true;
                        break;
                    }
                }
            } while (progressed);

            // 新增校验：确保最终的方向与线段聚合方向一致
            try
            {
                if (result.Count >= 2)
                {
                    var overallVec = result.Last() - result.First();
                    if (!overallVec.IsZeroLength())
                    {
                        var agg = ComputeAggregateSegmentDirection(segments);
                        if (!agg.IsZeroLength())
                        {
                            // 如果总体向量与聚合向量点积为负，则反转顶点顺序
                            if (overallVec.DotProduct(agg) < 0)
                            {
                                result.Reverse();
                            }
                        }
                    }
                }
            }
            catch
            {
                // 容错：若聚合计算失败，不影响已有顺序
            }
            return result;
        }


        //private List<Point3d> BuildOrderedVerticesFromSegments(List<LineSegmentInfo> segments, double tol = 1e-6)
        //{
        //    var result = new List<Point3d>();
        //    if (segments == null || segments.Count == 0) return result;

        //    // 中文注释：容差比较两点是否相等，避免浮点误差导致断链。
        //    static bool PointsEqual(Point3d a, Point3d b, double tolerance)
        //    {
        //        return Math.Abs(a.X - b.X) <= tolerance &&
        //               Math.Abs(a.Y - b.Y) <= tolerance &&
        //               Math.Abs(a.Z - b.Z) <= tolerance;
        //    }

        //    // 中文注释：把几何点离散成“唯一点索引”，用于做图拓扑统计。
        //    var uniquePoints = new List<Point3d>();
        //    int GetOrAddPointIndex(Point3d p)
        //    {
        //        for (int i = 0; i < uniquePoints.Count; i++)
        //        {
        //            if (PointsEqual(uniquePoints[i], p, tol)) return i;
        //        }
        //        uniquePoints.Add(p);
        //        return uniquePoints.Count - 1;
        //    }

        //    // 中文注释：统计每个点的度数，并记录每条线段的起止索引。
        //    var degree = new List<int>();
        //    var segPairs = new List<(int s, int e)>();
        //    foreach (var seg in segments)
        //    {
        //        int si = GetOrAddPointIndex(seg.StartPoint);
        //        int ei = GetOrAddPointIndex(seg.EndPoint);
        //        segPairs.Add((si, ei));

        //        while (degree.Count < uniquePoints.Count) degree.Add(0);
        //        degree[si]++;
        //        degree[ei]++;
        //    }

        //    // 中文注释：优先选“度为1”的端点作为起点（非闭环链）。
        //    int startIndex = -1;
        //    for (int i = 0; i < degree.Count; i++)
        //    {
        //        if (degree[i] == 1)
        //        {
        //            startIndex = i;
        //            break;
        //        }
        //    }

        //    // 中文注释：若没有端点（闭环或复杂网），退回第一条线段起点。
        //    if (startIndex == -1)
        //    {
        //        startIndex = segPairs.Count > 0 ? segPairs[0].s : 0;
        //    }

        //    // 中文注释：沿连通关系逐段行走，构建有序顶点。
        //    var visited = new bool[segPairs.Count];
        //    var current = uniquePoints[startIndex];
        //    result.Add(current);

        //    bool progressed;
        //    do
        //    {
        //        progressed = false;
        //        for (int i = 0; i < segPairs.Count; i++)
        //        {
        //            if (visited[i]) continue;

        //            var (si, ei) = segPairs[i];
        //            if (PointsEqual(uniquePoints[si], current, tol))
        //            {
        //                var next = uniquePoints[ei];
        //                if (!PointsEqual(next, result.Last(), tol)) result.Add(next);
        //                current = next;
        //                visited[i] = true;
        //                progressed = true;
        //                break;
        //            }

        //            if (PointsEqual(uniquePoints[ei], current, tol))
        //            {
        //                var next = uniquePoints[si];
        //                if (!PointsEqual(next, result.Last(), tol)) result.Add(next);
        //                current = next;
        //                visited[i] = true;
        //                progressed = true;
        //                break;
        //            }
        //        }
        //    } while (progressed);

        //    // 中文注释：方向一致性校验，避免结果方向与线段聚合方向相反。
        //    try
        //    {
        //        if (result.Count >= 2)
        //        {
        //            var overallVec = result.Last() - result.First();
        //            if (!overallVec.IsZeroLength())
        //            {
        //                var aggregateDir = ComputeAggregateSegmentDirection(segments);
        //                if (!aggregateDir.IsZeroLength() && overallVec.DotProduct(aggregateDir) < 0)
        //                {
        //                    result.Reverse();
        //                }
        //            }
        //        }
        //    }
        //    catch
        //    {
        //        // 中文注释：方向校验失败不影响主流程，返回已有排序结果。
        //    }

        //    return result;
        //}


        /// <summary>
        /// 获取选择的线段信息
        /// </summary>
        /// <param name="orderedVertices">有序顶点列表</param>
        /// <param name="totalLength">总长度</param>
        /// <returns></returns>
        private (Point3d midPoint, double midAngle) ComputeMidPointAndAngle(List<Point3d> orderedVertices, double totalLength)
        {
            double halfLen = totalLength / 2.0;
            double acc = 0.0;
            Point3d midPoint = orderedVertices[0];
            double midAngle = 0.0;

            for (int i = 0; i < orderedVertices.Count - 1; i++)
            {
                var p1 = orderedVertices[i];
                var p2 = orderedVertices[i + 1];
                double segLen = p1.DistanceTo(p2);
                if (acc + segLen >= halfLen)
                {
                    double t = (halfLen - acc) / segLen;
                    midPoint = new Point3d(
                        p1.X + (p2.X - p1.X) * t,
                        p1.Y + (p2.Y - p1.Y) * t,
                        p1.Z + (p2.Z - p1.Z) * t
                    );
                    midAngle = ComputeSegmentAngleUcs(p1, p2);
                    break;
                }
                acc += segLen;
            }
            return (midPoint, midAngle);
        }

        /// <summary>
        /// 计算线段角度
        /// </summary>
        /// <param name="p1">起点</param>
        /// <param name="p2">终点</param>
        /// <returns>线段在UCS中的角度</returns>
        private static double ComputeSegmentAngleUcs(Point3d p1, Point3d p2)
        {
            // 当前UCS的XY平面，保证与AutoCAD旋转角同一参考
            var plane = new Plane(Point3d.Origin, Vector3d.ZAxis);
            Vector3d dir = (p2 - p1).GetNormal();
            double angle = dir.AngleOnPlane(plane); // 以正X为0，逆时针为正
                                                    // 归一化到 [0, 2π)
            if (angle < 0) angle += 2.0 * Math.PI;
            return angle;
        }

        /// <summary>
        /// 兼容补丁：计算中点与中点段角度。
        /// </summary>
        //private (Point3d midPoint, double midAngle) ComputeMidPointAndAngle(List<Point3d> orderedVertices, double totalLength)
        //{
        //    if (orderedVertices == null || orderedVertices.Count < 2) return (Point3d.Origin, 0.0);
        //    if (totalLength <= 0.0)
        //    {
        //        var p0 = orderedVertices[0];
        //        var p1 = orderedVertices[1];
        //        var mid = new Point3d((p0.X + p1.X) / 2.0, (p0.Y + p1.Y) / 2.0, (p0.Z + p1.Z) / 2.0);
        //        return (mid, (p1 - p0).AngleOnPlane(new Plane(Point3d.Origin, Vector3d.ZAxis)));
        //    }

        //    double half = totalLength / 2.0;
        //    double acc = 0.0;
        //    for (int i = 0; i < orderedVertices.Count - 1; i++)
        //    {
        //        var a = orderedVertices[i];
        //        var b = orderedVertices[i + 1];
        //        double len = a.DistanceTo(b);
        //        if (acc + len >= half)
        //        {
        //            double t = (half - acc) / Math.Max(len, 1e-9);
        //            var mid = new Point3d(a.X + (b.X - a.X) * t, a.Y + (b.Y - a.Y) * t, a.Z + (b.Z - a.Z) * t);
        //            double ang = (b - a).AngleOnPlane(new Plane(Point3d.Origin, Vector3d.ZAxis));
        //            return (mid, ang);
        //        }
        //        acc += len;
        //    }

        //    var s0 = orderedVertices[orderedVertices.Count - 2];
        //    var s1 = orderedVertices[orderedVertices.Count - 1];
        //    return (s1, (s1 - s0).AngleOnPlane(new Plane(Point3d.Origin, Vector3d.ZAxis)));
        //}


        /// <summary>
        /// 计算某点附近的方向向量（优先使用与 referencePoint 最近的线段）
        /// </summary>
        private static Vector3d ComputeDirectionAtPoint(List<Point3d> orderedVertices, Point3d referencePoint, double tol = 1e-6)
        {
            if (orderedVertices == null || orderedVertices.Count < 2)
                return Vector3d.XAxis;

            Vector3d fallbackDir = ComputePathDirectionVector(orderedVertices, tol);
            double bestDist = double.MaxValue;
            Vector3d bestDir = fallbackDir.IsZeroLength() ? Vector3d.XAxis : fallbackDir;

            for (int i = 0; i < orderedVertices.Count - 1; i++)
            {
                Point3d start = orderedVertices[i];
                Point3d end = orderedVertices[i + 1];
                Vector3d segment = end - start;
                if (segment.IsZeroLength())
                    continue;

                Point3d projected = ProjectPointToSegment(referencePoint, start, end);
                double dist = referencePoint.DistanceTo(projected);
                if (dist + tol < bestDist)
                {
                    bestDist = dist;
                    bestDir = segment.GetNormal();
                }
            }

            if (!bestDir.IsZeroLength() && !fallbackDir.IsZeroLength() && bestDir.DotProduct(fallbackDir) < 0)
            {
                bestDir = -bestDir;
            }

            return bestDir.IsZeroLength() ? fallbackDir : bestDir;
        }

        /// <summary>
        /// 兼容补丁：计算点位处方向。
        /// </summary>
        //private Vector3d ComputeDirectionAtPoint(List<Point3d> orderedVertices, Point3d targetPoint, double tol = 1e-6)
        //{
        //    if (orderedVertices == null || orderedVertices.Count < 2) return Vector3d.XAxis;
        //    double best = double.MaxValue;
        //    Vector3d bestDir = Vector3d.XAxis;
        //    for (int i = 0; i < orderedVertices.Count - 1; i++)
        //    {
        //        var a = orderedVertices[i];
        //        var b = orderedVertices[i + 1];
        //        var v = b - a;
        //        if (v.IsZeroLength()) continue;
        //        var mid = new Point3d((a.X + b.X) / 2.0, (a.Y + b.Y) / 2.0, (a.Z + b.Z) / 2.0);
        //        var d = mid.DistanceTo(targetPoint);
        //        if (d < best - tol)
        //        {
        //            best = d;
        //            bestDir = v.GetNormal();
        //        }
        //    }
        //    return bestDir;
        //}

        /// <summary>
        /// 计算整条路径的总体方向向量（UCS，Z=+）
        /// </summary>
        private static Vector3d ComputePathDirectionVector(List<Point3d> orderedVertices, double tol = 1e-6)
        {
            if (orderedVertices == null || orderedVertices.Count < 2)
                return Vector3d.XAxis;

            // 直接用整体起点→终点的向量，保证箭头指向终点（流向）
            Vector3d overall = orderedVertices.Last() - orderedVertices.First();
            if (overall.Length > tol)
                return overall.GetNormal();

            // 回退：选择最长段方向
            double maxLen = 0.0;
            Vector3d longestDir = Vector3d.XAxis;
            for (int i = 0; i < orderedVertices.Count - 1; i++)
            {
                Vector3d v = orderedVertices[i + 1] - orderedVertices[i];
                if (v.Length > maxLen)
                {
                    maxLen = v.Length;
                    longestDir = v.GetNormal();
                }
            }
            return longestDir;
        }

        /// <summary>
        /// 将点投影到指定线段上
        /// </summary>
        private static Point3d ProjectPointToSegment(Point3d point, Point3d segmentStart, Point3d segmentEnd)
        {
            Vector3d segment = segmentEnd - segmentStart;
            if (segment.IsZeroLength())
                return segmentStart;

            Vector3d toPoint = point - segmentStart;
            double t = toPoint.DotProduct(segment) / segment.DotProduct(segment);
            t = Math.Max(0.0, Math.Min(1.0, t));
            return segmentStart + segment * t;
        }



        /// <summary>
        /// 兼容补丁：聚合线段方向。
        /// </summary>
        //private Vector3d ComputeAggregateSegmentDirection(List<LineSegmentInfo> segments)
        //{
        //    if (segments == null || segments.Count == 0) return Vector3d.XAxis;
        //    Vector3d sum = Vector3d.XAxis * 0.0;
        //    foreach (var s in segments)
        //    {
        //        var v = s.EndPoint - s.StartPoint;
        //        if (!v.IsZeroLength()) sum = sum + v.GetNormal() * Math.Max(s.Length, 1.0);
        //    }
        //    return sum.IsZeroLength() ? Vector3d.XAxis : sum.GetNormal();
        //}

        /// <summary>
        /// 获取箭头
        /// </summary>
        /// <param name="segments"></param>
        /// <returns></returns>
        private static Vector3d ComputeAggregateSegmentDirection(List<LineSegmentInfo> segments)
        {
            if (segments == null || segments.Count == 0)
                return new Vector3d(0, 0, 0);

            Vector3d sum = new Vector3d(0, 0, 0);
            foreach (var seg in segments)
            {
                Vector3d dir = seg.EndPoint - seg.StartPoint;
                if (!dir.IsZeroLength())
                    sum += dir.GetNormal();
            }

            return sum.IsZeroLength() ? new Vector3d(0, 0, 0) : sum.GetNormal();
        }

        /// <summary>
        /// 兼容补丁：构建局部管道 Polyline。
        /// </summary>
        //private Polyline BuildPipePolylineLocal(Polyline template, List<Point3d> verticesWorld, Point3d midPointWorld)
        //{
        //    var pl = new Polyline();
        //    if (template == null || verticesWorld == null) return pl;

        //    double w = template.ConstantWidth;
        //    for (int i = 0; i < verticesWorld.Count; i++)
        //    {
        //        var p = verticesWorld[i];
        //        var local = new Point2d(p.X - midPointWorld.X, p.Y - midPointWorld.Y);
        //        pl.AddVertexAt(i, local, 0.0, w, w);
        //    }

        //    pl.Layer = template.Layer;
        //    pl.Color = template.Color;
        //    pl.LineWeight = template.LineWeight;
        //    pl.Linetype = template.Linetype;
        //    pl.LinetypeScale = template.LinetypeScale;
        //    pl.Closed = false;
        //    return pl;
        //}

        /// <summary>
        /// 构建局部坐标的管线 Polyline
        /// </summary>
        /// <param name="template">模板 Polyline</param>
        /// <param name="verticesWorld">全局坐标系下的顶点列表</param>
        /// <param name="midPointWorld">全局坐标系下的中点</param>
        /// <returns>局部坐标系下的管线 Polyline</returns>
        private Polyline BuildPipePolylineLocal(Polyline template, List<Point3d> verticesWorld, Point3d midPointWorld)
        {
            var pl = new Polyline();
            double lineWeightScale = VariableDictionary.textBoxScale;
            for (int i = 0; i < verticesWorld.Count; i++)
            {
                var local = new Point2d(verticesWorld[i].X - midPointWorld.X, verticesWorld[i].Y - midPointWorld.Y);
                //var local = new Point2d(verticesWorld[i].X, verticesWorld[i].Y);
                pl.AddVertexAt(i, local, 0,
                    template.ConstantWidth * lineWeightScale,
                    template.ConstantWidth * lineWeightScale);
            }

            pl.Layer = template.Layer;
            pl.Color = template.Color;
            pl.LineWeight = template.LineWeight;
            pl.Linetype = template.Linetype;
            pl.LinetypeScale = template.LinetypeScale;
            pl.Elevation = 0;
            pl.Normal = Vector3d.ZAxis;
            pl.Closed = false;
            return pl;
        }


        /// <summary>
        /// 兼容补丁：克隆属性定义并转换到局部坐标。
        /// </summary>
        //private List<AttributeDefinition> CloneAttributeDefinitionsLocal(List<AttributeDefinition> defs, Point3d midPointWorld, double finalRotation, double pipelineLength, string titleFallback)
        //{
        //    var result = new List<AttributeDefinition>();
        //    if (defs == null) defs = new List<AttributeDefinition>();

        //    foreach (var d in defs)
        //    {
        //        var c = d?.Clone() as AttributeDefinition;
        //        if (c == null) continue;
        //        c.Position = new Point3d(d.Position.X - midPointWorld.X, d.Position.Y - midPointWorld.Y, 0.0);
        //        c.Rotation = d.Rotation;
        //        result.Add(c);
        //    }

        //    if (!result.Any(x => string.Equals(x.Tag, "管道标题", StringComparison.OrdinalIgnoreCase)))
        //    {
        //        result.Add(new AttributeDefinition
        //        {
        //            Tag = "管道标题",
        //            Position = Point3d.Origin,
        //            Rotation = finalRotation,
        //            TextString = string.IsNullOrWhiteSpace(titleFallback) ? "管道" : titleFallback,
        //            Height = defs.Count > 0 ? defs[0].Height : 2.5,
        //            Invisible = false,
        //            Constant = false
        //        });
        //    }

        //    return result;
        //}

        /// <summary>
        /// 创建属性定义
        /// </summary>
        /// <param name="defs">属性定义列表</param>
        /// <param name="midPointWorld">中点位置（世界坐标系）</param>
        /// <param name="finalRotation">最终旋转角度</param>
        /// <param name="pipelineLength">管道长度</param>
        /// <param name="titleFallback">管道标题后备值</param>
        /// <returns>属性定义列表</returns>
        private List<AttributeDefinition> CloneAttributeDefinitionsLocal(List<AttributeDefinition> defs, Point3d midPointWorld, double finalRotation, double pipelineLength, string titleFallback)
        {
            var result = new List<AttributeDefinition>();
            bool hasTitle = false;

            foreach (var def in defs)
            {
                var cloned = def.Clone() as AttributeDefinition;
                if (cloned == null) continue;

                // 转为局部坐标（相对中点）
                var localPos = new Point3d(def.Position.X - midPointWorld.X, def.Position.Y - midPointWorld.Y, 0);
                cloned.Position = localPos;
                cloned.Rotation = def.Rotation;
                cloned.Invisible = def.Invisible;
                cloned.Constant = def.Constant;
                cloned.Tag = def.Tag;
                cloned.TextString = def.TextString;
                cloned.Height = def.Height;

                if (!string.IsNullOrWhiteSpace(cloned.Tag))
                {
                    var tagLower = cloned.Tag.ToLowerInvariant();
                    if (tagLower.Contains("长度") || tagLower.Contains("length"))
                    {
                        double baseValue = 0.0;
                        if (double.TryParse(cloned.TextString, out double parsed)) baseValue = parsed;
                        cloned.TextString = (baseValue + pipelineLength).ToString("0.###");
                    }
                    if (string.Equals(cloned.Tag, "管道标题", StringComparison.OrdinalIgnoreCase))
                    {
                        hasTitle = true;
                        cloned.Position = Point3d.Origin;
                        cloned.Rotation = finalRotation;
                        cloned.Invisible = false;
                        if (string.IsNullOrWhiteSpace(cloned.TextString))
                            cloned.TextString = titleFallback ?? "管道";
                    }
                }

                result.Add(cloned);
            }

            if (!hasTitle)
            {
                result.Add(new AttributeDefinition
                {
                    Tag = "管道标题",
                    Position = Point3d.Origin,
                    Rotation = finalRotation,
                    TextString = string.IsNullOrWhiteSpace(titleFallback) ? "管道" : titleFallback,
                    Height = defs != null && defs.Count > 0 ? defs[0].Height : 2.5,
                    Invisible = false,
                    Constant = false
                });
            }

            return result;
        }





        /// <summary>
        /// 兼容补丁：创建箭头与标题（增强版：保证有三角符号，并在其上方添加管道标题）。
        /// </summary>
        //private List<Entity> CreateDirectionalArrowsAndTitles(DBTrans tr, SamplePipeInfo sampleInfo, List<Point3d> orderedVertices, Point3d midPoint, string pipeTitle, string sampleBlockName)
        //{
        //    var list = new List<Entity>();
        //    if (orderedVertices == null || orderedVertices.Count < 2) return list;

        //    // 中文注释：优先使用样例块中的箭头模板；若模板不是闭合三角，则回退为默认闭合三角，确保不会只显示两条线。
        //    Polyline arrowTemplate = null;
        //    if (sampleInfo?.DirectionArrowTemplate != null && sampleInfo.DirectionArrowTemplate.NumberOfVertices >= 3)
        //    {
        //        var tmp = sampleInfo.DirectionArrowTemplate.Clone() as Polyline;
        //        if (tmp != null)
        //        {
        //            // 中文注释：模板顶点足够时强制闭合，避免出现“只有两条线”的箭头。
        //            tmp.Closed = true;
        //            arrowTemplate = tmp;
        //        }
        //    }

        //    // 中文注释：若样例中没有可用三角模板，退回默认三角。默认尺寸接近原有逻辑。
        //    if (arrowTemplate == null)
        //    {
        //        arrowTemplate = new Polyline();
        //        arrowTemplate.AddVertexAt(0, new Point2d(5.0, 0.0), 0, 0, 0);
        //        arrowTemplate.AddVertexAt(1, new Point2d(-5.0, -2.0), 0, 0, 0);
        //        arrowTemplate.AddVertexAt(2, new Point2d(-5.0, 2.0), 0, 0, 0);
        //        arrowTemplate.Closed = true;
        //        if (sampleInfo?.PipeBodyTemplate != null)
        //        {
        //            arrowTemplate.Layer = sampleInfo.PipeBodyTemplate.Layer;
        //            arrowTemplate.Color = sampleInfo.PipeBodyTemplate.Color;
        //            arrowTemplate.LineWeight = sampleInfo.PipeBodyTemplate.LineWeight;
        //            arrowTemplate.Linetype = sampleInfo.PipeBodyTemplate.Linetype;
        //            arrowTemplate.LinetypeScale = sampleInfo.PipeBodyTemplate.LinetypeScale;
        //        }
        //    }

        //    // 中文注释：从模板属性定义中提取“管道标题”样式（颜色/高度/旋转/文字样式）。
        //    AttributeDefinition titleTemplate = null;
        //    if (sampleInfo?.AttributeDefinitions != null)
        //    {
        //        titleTemplate = sampleInfo.AttributeDefinitions.FirstOrDefault(a =>
        //            a != null && string.Equals(a.Tag, "管道标题", StringComparison.OrdinalIgnoreCase));
        //    }

        //    // 中文注释：先生成世界坐标中的箭头，再转为块局部坐标。
        //    var worldArrows = PipeArrowPlacer.CreateDirectionalArrows(orderedVertices, arrowTemplate, 50.0);
        //    foreach (var e in worldArrows)
        //    {
        //        // 中文注释：按需求固定流向箭头颜色为 ACI 6。
        //        e.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
        //            Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 6);

        //        e.TransformBy(Matrix3d.Displacement(Point3d.Origin - midPoint));
        //        list.Add(e);

        //        // 中文注释：若箭头是闭合三角Polyline，则追加Solid填充，颜色固定为 ACI 6。
        //        try
        //        {
        //            if (e is Polyline arrowPl && arrowPl.NumberOfVertices >= 3 && arrowPl.Closed)
        //            {
        //                var p0 = arrowPl.GetPoint3dAt(0);
        //                var p1 = arrowPl.GetPoint3dAt(1);
        //                var p2 = arrowPl.GetPoint3dAt(2);

        //                var fill = new Solid(p0, p1, p2, p2);
        //                fill.Layer = arrowPl.Layer;
        //                fill.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
        //                    Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 6);
        //                fill.LineWeight = arrowPl.LineWeight;
        //                list.Add(fill);
        //            }
        //        }
        //        catch
        //        {
        //            // 中文注释：填充失败不影响主流程，至少保留轮廓箭头。
        //        }
        //    }

        //    // 中文注释：为每一段在线段中点上方添加管道标题文本（与线段方向一致）。
        //    var title = string.IsNullOrWhiteSpace(pipeTitle) ? (sampleBlockName ?? "管道") : pipeTitle;
        //    for (int i = 0; i < orderedVertices.Count - 1; i++)
        //    {
        //        var p0 = orderedVertices[i];
        //        var p1 = orderedVertices[i + 1];
        //        var segVec = p1 - p0;
        //        if (segVec.IsZeroLength()) continue;

        //        var segLen = p0.DistanceTo(p1);
        //        bool isFirst = (i == 0);
        //        bool isLast = (i == orderedVertices.Count - 2);
        //        bool shouldPlace = isFirst || isLast || segLen >= 50.0;
        //        if (!shouldPlace) continue;

        //        var dir = segVec.GetNormal();

        //        // 中文注释：先取垂直于管段的法向，再统一到“视觉上方”（优先 +Y 方向）。
        //        var normal = new Vector3d(-dir.Y, dir.X, 0.0);
        //        if (normal.IsZeroLength()) normal = Vector3d.YAxis;
        //        normal = normal.GetNormal();
        //        if (normal.Y < 0 || (Math.Abs(normal.Y) < 1e-6 && normal.X < 0))
        //        {
        //            normal = -normal;
        //        }

        //        var mid = new Point3d((p0.X + p1.X) / 2.0, (p0.Y + p1.Y) / 2.0, (p0.Z + p1.Z) / 2.0);
        //        var textPosWorld = mid + normal * 6.0;
        //        var textPosLocal = textPosWorld.TransformBy(Matrix3d.Displacement(Point3d.Origin - midPoint));

        //        var text = new DBText
        //        {
        //            TextString = title,
        //            Position = textPosLocal,
        //            Height = (titleTemplate != null && titleTemplate.Height > 0) ? titleTemplate.Height : 3.5,
        //            Rotation = dir.AngleOnPlane(new Plane(Point3d.Origin, Vector3d.ZAxis)),
        //            HorizontalMode = TextHorizontalMode.TextCenter,
        //            VerticalMode = TextVerticalMode.TextBottom,
        //            AlignmentPoint = textPosLocal
        //        };

        //        // 中文注释：标题样式优先读取模板“管道标题”属性定义。
        //        if (titleTemplate != null)
        //        {
        //            text.Layer = titleTemplate.Layer;
        //            text.Color = titleTemplate.Color;
        //            text.TextStyleId = titleTemplate.TextStyleId;
        //            text.WidthFactor = titleTemplate.WidthFactor;
        //            text.Oblique = titleTemplate.Oblique;
        //        }
        //        else if (sampleInfo?.PipeBodyTemplate != null)
        //        {
        //            text.Layer = sampleInfo.PipeBodyTemplate.Layer;
        //            text.Color = sampleInfo.PipeBodyTemplate.Color;
        //        }

        //        list.Add(text);
        //    }

        //    return list;
        //}


        private List<Entity> CreateDirectionalArrowsAndTitles(DBTrans tr, SamplePipeInfo sampleInfo, List<Point3d> verticesWorld, Point3d midPointWorld, string pipeTitle, string sampleBlockName)
        {
            var overlay = new List<Entity>();
            if (sampleInfo == null || verticesWorld == null || verticesWorld.Count < 2) return overlay;

            // 优先使用用户在TextBox_绘图比例中设置的比例值
            var scaleDenom = VariableDictionary.textBoxScale;
            if (scaleDenom <= 0) // 如果获取失败，使用原有逻辑
            {
                AutoCadHelper.GetAndApplyActiveDrawingScale();//获取当前绘图比例
                scaleDenom = VariableDictionary.blockScale;
            }

            // 计算缩放因子（相对于100的比例）
            //double scaleFactor = scaleDenom / 100.0;
            double scaleFactor = scaleDenom;

            // 箭头模板与填充准备：若无模板则用默认三角
            Polyline arrowTemplate = sampleInfo.DirectionArrowTemplate;
            Solid? fillTemplate = null;
            double explicitArrowLength = 8;  // 基础长度
            double explicitArrowHeight = 2.0;   // 基础高度
            if (arrowTemplate == null)
            {
                // 根据名称确定箭头样式
                var (colorIdx, length, height) = DetermineArrowStyleByName(sampleBlockName);
                explicitArrowLength = length * scaleFactor;  // 应用比例
                explicitArrowHeight = height * scaleFactor;  // 应用比例
                // 创建箭头
                var (outline, fill) = CreateArrowTriangleFilled(explicitArrowLength, explicitArrowHeight, colorIdx, sampleInfo.PipeBodyTemplate);
                arrowTemplate = outline;
                fillTemplate = fill;
            }
            else
            {
                // 如果有模板箭头，也按比例缩放
                try
                {
                    if (scaleFactor != 1.0)
                    {
                        // 克隆模板并按比例缩放
                        arrowTemplate = (Polyline)arrowTemplate.Clone();
                        Matrix3d scaleMatrix = Matrix3d.Scaling(scaleFactor, Point3d.Origin);
                        arrowTemplate.TransformBy(scaleMatrix);

                        if (fillTemplate != null)
                        {
                            fillTemplate = (Solid)fillTemplate.Clone();
                            fillTemplate.TransformBy(scaleMatrix);
                        }
                    }
                }
                catch
                {
                    // 如果缩放失败，使用原始模板
                }
            }

            // 标题最终高度：基准 3.5 * 比例分母（与表格一致）
            double finalTitleHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, scaleDenom);

            // 遍历每一段，生成箭头并在箭头"上方"放置居中对齐的标题文字
            for (int i = 0; i < verticesWorld.Count - 1; i++)
            {
                var p1 = verticesWorld[i];
                var p2 = verticesWorld[i + 1];
                var seg = p2 - p1;
                if (seg.IsZeroLength()) continue;

                var dir = seg.GetNormal();
                var mid = new Point3d((p1.X + p2.X) / 2.0, (p1.Y + p2.Y) / 2.0, (p1.Z + p2.Z) / 2.0);

                Polyline? outlineAligned = null;
                Solid? fillAligned = null;
                try
                {
                    // 箭头模板对齐（注意：这里只做旋转，缩放已在上面处理）
                    (outlineAligned, fillAligned) = AlignArrowToDirection(arrowTemplate, fillTemplate, dir);

                    // 箭头模板平移
                    var localDisp = mid - midPointWorld;
                    if (outlineAligned != null)
                    {
                        // 箭头模板平移
                        outlineAligned.TransformBy(Matrix3d.Displacement(new Vector3d(localDisp.X, localDisp.Y, localDisp.Z)));
                        // 箭头模板设置图层
                        outlineAligned.Layer = sampleInfo.PipeBodyTemplate.Layer;
                        // 箭头模板添加到 overlay
                        overlay.Add(outlineAligned);
                    }
                    if (fillAligned != null)//填充
                    {
                        // 填充模板平移
                        fillAligned.TransformBy(Matrix3d.Displacement(new Vector3d(localDisp.X, localDisp.Y, localDisp.Z)));
                        // 填充模板设置图层
                        fillAligned.Layer = sampleInfo.PipeBodyTemplate.Layer;
                        overlay.Add(fillAligned);
                    }
                }
                catch
                {
                    // 忽略箭头生成异常，继续生成标题
                }

                try
                {
                    // 计算文字放置方向：取段法线的+90度方向作为"上方"
                    var perp = new Vector3d(-dir.Y, dir.X, 0.0);
                    if (perp.IsZeroLength())
                        perp = Vector3d.YAxis;
                    else
                        perp = perp.GetNormal();

                    // 确保 perp 指向图纸上侧（全局 +Y）
                    if (perp.DotProduct(Vector3d.YAxis) < 0)
                        perp = -perp;

                    // 估算箭头半高以确定文字偏移，优先使用已对齐实体的几何包围盒
                    double arrowHalfHeight = explicitArrowHeight / 2.0;
                    try
                    {
                        // 获取实体尺寸
                        Entity sizeEntity = (Entity?)outlineAligned ?? (Entity?)fillAligned;
                        if (sizeEntity != null)
                        {
                            var ext = sizeEntity.GeometricExtents;// 获取实体尺寸
                            arrowHalfHeight = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y) / 2.0;// 计算箭头半高
                            if (arrowHalfHeight < 1e-6) arrowHalfHeight = explicitArrowHeight / 2.0;// 如果获取尺寸失败，使用默认值
                        }
                    }
                    catch { arrowHalfHeight = explicitArrowHeight / 2.0; }// 如果获取尺寸失败，使用默认值

                    // 文字偏移：箭头上方 + 与文字高度相关的间距（按比例调整）
                    //double offset = (arrowHalfHeight + finalTitleHeight * 0.8) * scaleFactor; // 应用比例
                    double offset = (arrowHalfHeight + finalTitleHeight * 0.8); // 应用比例
                    var worldTextPos = mid + perp * offset;
                    var localTextPos = new Point3d(worldTextPos.X - midPointWorld.X, worldTextPos.Y - midPointWorld.Y, worldTextPos.Z - midPointWorld.Z);

                    // 文字方向：沿段方向，保证可读（不倒置）
                    double segAngle = ComputeSegmentAngleUcs(p1, p2);
                    double textRot = segAngle;
                    if (Math.Cos(textRot) < 0) textRot += Math.PI;
                    if (textRot > Math.PI) textRot -= 2.0 * Math.PI;
                    if (textRot <= -Math.PI) textRot += 2.0 * Math.PI;

                    // 创建 DBText 并设置为居中对齐
                    var dbText = new DBText
                    {
                        Position = localTextPos,
                        Height = finalTitleHeight,
                        TextString = string.IsNullOrWhiteSpace(pipeTitle) ? sampleBlockName ?? "管道" : pipeTitle,
                        Rotation = textRot,
                        Layer = sampleInfo.PipeBodyTemplate.Layer,
                        Normal = Vector3d.ZAxis,
                        Oblique = 0.0
                    };

                    // 设置对齐点并置中（水平 + 垂直）
                    try
                    {
                        dbText.AlignmentPoint = localTextPos;
                        dbText.HorizontalMode = TextHorizontalMode.TextCenter;
                        dbText.VerticalMode = TextVerticalMode.TextVerticalMid;
                    }
                    catch
                    {
                        // 某些 API/版本对这些属性有限制，忽略异常
                    }

                    // 应用样式并保证高度按当前比例（FontsStyleHelper 内部也会确保 TextStyle 存在）
                    try
                    {
                        TextFontsStyleHelper.ApplyTitleToDBText(tr, dbText, scaleDenom);
                    }
                    catch
                    {
                        // 若样式应用失败，仍使用 dbText 的 Height
                    }

                    overlay.Add(dbText);
                }
                catch
                {
                    // 忽略该段文字生成异常
                }
            }

            return overlay;
        }

        /// <summary>
        /// 根据名称确定箭头样式
        /// </summary>
        /// <param name="blockName">块名称</param>
        /// <returns>箭头样式元组</returns>
        private (short colorIndex, double length, double height) DetermineArrowStyleByName(string blockName)
        {
            string nameLower = (blockName ?? string.Empty).ToLowerInvariant();
            bool isOutlet = nameLower.Contains("出口") || nameLower.Contains("outlet");
            bool isInlet = nameLower.Contains("入口") || nameLower.Contains("inlet");

            // 出口=黄色(ACI 2)，入口=绿色(ACI 3)，默认黄色
            short colorIndex = isInlet ? (short)3 : (short)2;
            if (!isInlet && !isOutlet)
            {
                colorIndex = 2;
            }

            return (colorIndex, 10.0, 3.0);
        }

        /// <summary>
        /// 新增：创建方向箭头（轮廓 + 填充）
        /// </summary>
        /// <param name="arrowLength">箭头长度</param>
        /// <param name="arrowHeight">箭头高度</param>
        /// <param name="colorIndex">颜色索引</param>
        /// <param name="pipeTemplate">管道模板</param>
        /// <returns>轮廓和填充的元组</returns>
        private (Polyline outline, Solid fill) CreateArrowTriangleFilled(double arrowLength, double arrowHeight, short colorIndex, Polyline pipeTemplate)
        {
            // 三角顶点（局部坐标，尖端朝 +X）
            var tip = new Point2d(arrowLength / 2.0, 0.0);
            var leftBottom = new Point2d(-arrowLength / 2.0, -arrowHeight / 2.0);
            var leftTop = new Point2d(-arrowLength / 2.0, arrowHeight / 2.0);

            // 轮廓
            var arrow = new Polyline();
            arrow.AddVertexAt(0, tip, 0, 0, 0);
            arrow.AddVertexAt(1, leftBottom, 0, 0, 0);
            arrow.AddVertexAt(2, leftTop, 0, 0, 0);
            arrow.Closed = true;
            arrow.Layer = pipeTemplate.Layer;
            //arrow.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex);
            //arrow.Linetype = pipeTemplate.Linetype;
            arrow.LinetypeScale = pipeTemplate.LinetypeScale;
            arrow.LineWeight = pipeTemplate.LineWeight;
            arrow.Elevation = 0;
            arrow.Normal = Vector3d.ZAxis;

            // 填充（二维实心三角形）
            var solid = new Solid(
                new Point3d(tip.X, tip.Y, 0),
                new Point3d(leftBottom.X, leftBottom.Y, 0),
                new Point3d(leftTop.X, leftTop.Y, 0),
                new Point3d(leftTop.X, leftTop.Y, 0) // 三角形第四点与第三点相同
            );
            solid.Layer = pipeTemplate.Layer;
            solid.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex);
            solid.LineWeight = pipeTemplate.LineWeight;
            solid.Normal = Vector3d.ZAxis;

            return (arrow, solid);
        }

        /// <summary>
        /// 将箭头几何按照指定方向对齐
        /// </summary>
        private (Polyline outline, Solid? fill) AlignArrowToDirection(Polyline arrowTemplate, Solid? fillTemplate, Vector3d direction)
        {
            // 计算模板主方向
            Vector3d dir = direction.IsZeroLength() ? Vector3d.XAxis : direction.GetNormal();
            // 计算模板侧向
            Vector3d yAxis = Vector3d.ZAxis.CrossProduct(dir);
            if (yAxis.IsZeroLength())// 如果主向和侧向平行，则侧向为 Y 轴
                yAxis = Vector3d.YAxis;// 侧向为 Z 轴
            else
                yAxis = yAxis.GetNormal();// 计算侧向
            // 计算对齐矩阵
            Matrix3d alignMatrix = Matrix3d.AlignCoordinateSystem(
                Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                Point3d.Origin, dir, yAxis, Vector3d.ZAxis
            );
            // 对齐模板
            var outline = (Polyline)arrowTemplate.Clone();
            outline.TransformBy(alignMatrix);// 对齐
            // 对齐填充
            Solid? fill = null;
            if (fillTemplate != null)
            {
                // 对齐填充
                fill = (Solid)fillTemplate.Clone();
                fill.TransformBy(alignMatrix);// 对齐
            }
            return (outline, fill);
        }

        /// <summary>
        /// 兼容补丁：获取下一个管段号。
        /// </summary>
        private int GetNextPipeSegmentNumber(Database db)
        {
            int max = 0;
            if (db == null) return 1;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                foreach (ObjectId id in ms)
                {
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;
                    foreach (ObjectId aid in br.AttributeCollection)
                    {
                        var ar = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                        if (ar == null) continue;
                        if (!string.Equals(ar.Tag, "管段号", StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(ar.Tag, "管段编号", StringComparison.OrdinalIgnoreCase)) continue;

                        var m = Regex.Match(ar.TextString ?? string.Empty, @"\d+");
                        if (m.Success && int.TryParse(m.Value, out var n) && n > max) max = n;
                    }
                }
                tr.Commit();
            }
            return max + 1;
        }

        /// <summary>
        /// 兼容补丁：构建块定义并返回块名。
        /// </summary>
        private string BuildPipeBlockDefinition(DBTrans tr, string desiredName, Polyline pipeLocal, List<Entity> overlayEntities, List<AttributeDefinition> attDefsLocal)
        {
            string finalName = string.IsNullOrWhiteSpace(desiredName) ? "PIPE_BLOCK" : desiredName;
            int suf = 1;
            while (tr.BlockTable.Has(finalName)) finalName = (desiredName ?? "PIPE_BLOCK") + "_PIPEGEN_" + suf++;

            tr.BlockTable.Add(
                finalName,
                btr => { btr.Origin = Point3d.Origin; },
                () =>
                {
                    var entities = new List<Entity>();
                    if (pipeLocal != null) entities.Add((Polyline)pipeLocal.Clone());
                    if (overlayEntities != null)
                    {
                        foreach (var e in overlayEntities)
                        {
                            if (e == null) continue;
                            var c = e.Clone() as Entity;
                            if (c != null) entities.Add(c);
                        }
                    }
                    return entities;
                },
                () => attDefsLocal ?? new List<AttributeDefinition>()
            );

            return finalName;
        }

        /// <summary>
        /// 兼容补丁：插入块并写入属性。
        /// </summary>
        private ObjectId InsertPipeBlockWithAttributes(DBTrans tr, Point3d insertPointWorld, string blockName, double rotation, Dictionary<string, string> attValues)
        {
            ObjectId btrId = tr.BlockTable[blockName];
            return tr.CurrentSpace.InsertBlock(insertPointWorld, btrId, rotation: rotation, atts: attValues);
        }
    }

    /// <summary>
    /// 管道属性编辑器（带字段排序与中文别名展示）。
    /// </summary>
    public class PipeAttributeEditorForm : Form
    {
        private readonly DataGridView _dataGridView;
        private readonly Button _btnOk;
        private readonly Button _btnCancel;
        private Dictionary<string, string> _attributes;

        /// <summary>
        /// 编辑后属性结果。
        /// </summary>
        public Dictionary<string, string> Attributes => new Dictionary<string, string>(_attributes, StringComparer.OrdinalIgnoreCase);

        public PipeAttributeEditorForm(Dictionary<string, string> initialAttributes)
        {
            _attributes = new Dictionary<string, string>(initialAttributes ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase);

            Text = "示例管道属性编辑";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(760, 520);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;

            _dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false
            };
            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn { Name = "Key", HeaderText = "字段", ReadOnly = true });
            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn { Name = "Value", HeaderText = "值", ReadOnly = false });

            _btnOk = new Button { Text = "完成", DialogResult = DialogResult.OK, Width = 90, Height = 30 };
            _btnCancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Width = 90, Height = 30 };
            _btnOk.Click += BtnOk_Click;
            _btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 52,
                FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                Padding = new Padding(8),
                WrapContents = false
            };
            panel.Controls.Add(_btnCancel);
            panel.Controls.Add(_btnOk);

            Controls.Add(_dataGridView);
            Controls.Add(panel);

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;

            LoadAttributesToGrid();
        }

        /// <summary>
        /// 字段排序与别名展示：规则来自 DictionaryHelper，可配置。
        /// </summary>
        private void LoadAttributesToGrid()
        {
            _dataGridView.Rows.Clear();

            var ordered = _attributes
                .OrderBy(kv => DictionaryHelper.GetPipeAttributeSortPriority(kv.Key))
                .ThenBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var kv in ordered)
            {
                _dataGridView.Rows.Add(DictionaryHelper.GetPipeAttributeDisplayAlias(kv.Key), kv.Value);
                _dataGridView.Rows[_dataGridView.Rows.Count - 1].Tag = kv.Key;
            }

            if (_dataGridView.Rows.Count > 0)
                _dataGridView.CurrentCell = _dataGridView.Rows[0].Cells[1];
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            try
            {
                var newDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < _dataGridView.Rows.Count; i++)
                {
                    var row = _dataGridView.Rows[i];
                    if (row.IsNewRow) continue;

                    var key = (row.Tag as string) ?? row.Cells["Key"].Value?.ToString() ?? string.Empty;
                    var val = row.Cells["Value"].Value?.ToString() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(key)) continue;

                    newDict[key] = val;
                }

                _attributes = newDict;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存属性失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
