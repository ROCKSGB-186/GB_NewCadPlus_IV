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

                // 把 aligned 平移到 mid 的位置。计算 aligned 质心
                var centroid = typeof(PipeDirectionHelper)
                    .GetMethod("ComputePolylineCentroid", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
                    ?.Invoke(null, new object[] { aligned }) as Point3d?;

                // 如果获取不到非公开方法，改用简单的第一个顶点作为参考
                Point3d refPt = centroid ?? aligned.GetPoint3dAt(0);

                var translation = Matrix3d.Displacement(mid - refPt);
                aligned.TransformBy(translation);

                // 复制到结果（注意：aligned 来自 AlignArrowToDirection，已经为新实例）
                result.Add(aligned);
            }

            return result;
        }
    }
}
