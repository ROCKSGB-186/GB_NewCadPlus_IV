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

                // 使用 PipeDirectionHelper.AlignArrowToDirection 将模板对齐到段方向
                var aligned = PipeDirectionHelper.AlignArrowToDirection(arrowTemplate, dirNorm);

                // 把 aligned 平移到 mid 的位置。这里使用“短边中点”对齐到线段中点，确保短边中心落在管道线上
                Point3d refPt;
                try
                {
                    // 先找尖端：沿流向投影最大顶点
                    double maxProj = double.NegativeInfinity;
                    Point3d tip = aligned.GetPoint3dAt(0);
                    for (int vi = 0; vi < aligned.NumberOfVertices; vi++)
                    {
                        var vp = aligned.GetPoint3dAt(vi);
                        double proj = vp.X * dirNorm.X + vp.Y * dirNorm.Y + vp.Z * dirNorm.Z;
                        if (proj > maxProj)
                        {
                            maxProj = proj;
                            tip = vp;
                        }
                    }

                    // 再找短边：排除尖端后，两点距离最短的一对作为短边
                    var others = new List<Point3d>();
                    for (int vi = 0; vi < aligned.NumberOfVertices; vi++)
                    {
                        var vp = aligned.GetPoint3dAt(vi);
                        if (!vp.IsEqualTo(tip)) others.Add(vp);
                    }

                    if (others.Count >= 2)
                    {
                        double minDist = double.PositiveInfinity;
                        Point3d a = others[0], b = others[1];
                        for (int i2 = 0; i2 < others.Count; i2++)
                        {
                            for (int j2 = i2 + 1; j2 < others.Count; j2++)
                            {
                                double d = others[i2].DistanceTo(others[j2]);
                                if (d < minDist)
                                {
                                    minDist = d;
                                    a = others[i2];
                                    b = others[j2];
                                }
                            }
                        }
                        // 短边中点作为对齐基准
                        refPt = new Point3d((a.X + b.X) / 2.0, (a.Y + b.Y) / 2.0, (a.Z + b.Z) / 2.0);
                    }
                    else
                    {
                        // 兜底：若无法识别短边，回退尖端
                        refPt = tip;
                    }
                }
                catch
                {
                    // 兜底：若计算失败，使用第一个顶点
                    refPt = aligned.GetPoint3dAt(0);
                }

                var translation = Matrix3d.Displacement(mid - refPt);
                aligned.TransformBy(translation);

                // 复制到结果（注意：aligned 来自 AlignArrowToDirection，已经为新实例）
                result.Add(aligned);
            }

            return result;
        }
    }
}
