using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>    
    ///  - 如果箭头位于第1段或最后一段 -> 强制添加；
    ///  - 否则仅当该段长度 >= 50 时才添加。
    /// 返回的实体为 Polyline（与 PipeDirectionHelper.AlignArrowToDirection 兼容）。
    /// </summary>
    public static class PipeArrowPlacer
    {
        /// <summary>
        /// 为路径（顶点序列）生成箭头实体集合（仅创建箭头，不创建文字标题）。
        /// verticesWorld: 按路径顺序的世界坐标点（至少 2 个）
        /// sampleInfo.DirectionArrowTemplate: 箭头模板（Polyline），将被对齐并复制到段中点
        /// </summary>
        public static List<Entity> CreateDirectionalArrows(List<Point3d> verticesWorld, Polyline arrowTemplate, double minLengthForArrow = 50.0)
        {
            var result = new List<Entity>();

            if (verticesWorld == null || verticesWorld.Count < 2 || arrowTemplate == null)
                return result;

            // 计算每一段并判断是否需要箭头
            int segCount = verticesWorld.Count - 1;
            for (int i = 0; i < segCount; i++)
            {
                var p0 = verticesWorld[i];
                var p1 = verticesWorld[i + 1];
                double segLength = p0.DistanceTo(p1);

                bool isFirst = (i == 0);
                bool isLast = (i == segCount - 1);

                bool placeArrow = false;
                if (isFirst || isLast)
                {
                    placeArrow = true;
                }
                else
                {
                    placeArrow = segLength >= minLengthForArrow;
                }

                if (!placeArrow) continue;

                // 计算段方向向量与中点
                var dir = (p1 - p0);
                if (dir.IsZeroLength()) continue;
                var dirNorm = dir.GetNormal();
                var mid = new Point3d((p0.X + p1.X) / 2.0, (p0.Y + p1.Y) / 2.0, (p0.Z + p1.Z) / 2.0);

                // 使用 PipeDirectionHelper.AlignArrowToDirection 将模板旋转到与段方向近似一致
                var aligned = PipeDirectionHelper.AlignArrowToDirection(arrowTemplate, dirNorm);

                // 进一步校正：保证模板短边中点到尖端的轴线与段方向严格共线，
                // 并把短边中点对齐到段中点 mid，这样尖端也会落在管道轴线上。
                try
                {
                    // 计算质心（反射调用私有方法，回退到第一个顶点）
                    var centroid = typeof(PipeDirectionHelper)
                        .GetMethod("ComputePolylineCentroid", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
                        ?.Invoke(null, new object[] { aligned }) as Point3d? ?? aligned.GetPoint3dAt(0);

                    // 找出尖端：在沿段方向上的最大投影点（相对质心投影）
                    double maxProj = double.NegativeInfinity;
                    int tipIdx = 0;
                    for (int vi = 0; vi < aligned.NumberOfVertices; vi++)
                    {
                        var vp = aligned.GetPoint3dAt(vi);
                        var rel = vp - centroid;
                        double proj = rel.DotProduct(dirNorm);
                        if (proj > maxProj)
                        {
                            maxProj = proj;
                            tipIdx = vi;
                        }
                    }

                    // 其余顶点中找最短边对作为短边
                    var others = new List<Point3d>();
                    for (int vi = 0; vi < aligned.NumberOfVertices; vi++)
                    {
                        if (vi == tipIdx) continue;
                        others.Add(aligned.GetPoint3dAt(vi));
                    }

                    Point3d baseMid;
                    if (others.Count >= 2)
                    {
                        double minPair = double.PositiveInfinity;
                        Point3d a = others[0], b = others[1];
                        for (int i2 = 0; i2 < others.Count; i2++)
                        {
                            for (int j2 = i2 + 1; j2 < others.Count; j2++)
                            {
                                double dd = others[i2].DistanceTo(others[j2]);
                                if (dd < minPair)
                                {
                                    minPair = dd;
                                    a = others[i2];
                                    b = others[j2];
                                }
                            }
                        }
                        baseMid = new Point3d((a.X + b.X) / 2.0, (a.Y + b.Y) / 2.0, (a.Z + b.Z) / 2.0);
                    }
                    else
                    {
                        baseMid = centroid;
                    }

                    // 旋转对齐：计算当前轴（tip - baseMid）与期望轴 dirNorm 的夹角，绕 baseMid 旋转该角
                    var tipPt = aligned.GetPoint3dAt(tipIdx);
                    var axis = tipPt - baseMid;
                    if (!axis.IsZeroLength())
                    {
                        double axisAngle = Math.Atan2(axis.Y, axis.X);
                        double pathAngle = Math.Atan2(dirNorm.Y, dirNorm.X);
                        double delta = pathAngle - axisAngle;
                        // 规范化到 [-PI,PI]
                        while (delta > Math.PI) delta -= 2 * Math.PI;
                        while (delta < -Math.PI) delta += 2 * Math.PI;
                        if (Math.Abs(delta) > 1e-9)
                        {
                            var rot = Matrix3d.Rotation(delta, Vector3d.ZAxis, baseMid);
                            aligned.TransformBy(rot);
                            // 更新 tipPt 与 baseMid（旋转后坐标已改变）
                            tipPt = aligned.GetPoint3dAt(tipIdx);
                            // recompute baseMid by rotating a and b if needed is unnecessary because rotation applied to aligned
                        }
                    }

                    // 将短边中点 baseMid 对齐到段中点 mid
                    var translation = Matrix3d.Displacement(mid - baseMid);
                    aligned.TransformBy(translation);
                }
                catch
                {
                    // 兜底：若校正失败，回退到以质心为基准的对齐
                    var refPt = aligned.GetPoint3dAt(0);
                    var translation = Matrix3d.Displacement(mid - refPt);
                    aligned.TransformBy(translation);
                }

                // 复制到结果（注意：aligned 来自 AlignArrowToDirection，已经为新实例）
                result.Add(aligned);
            }

            return result;
        }
    }
}
