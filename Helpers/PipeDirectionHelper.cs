using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace GB_NewCadPlus_LM.Helpers
{
    /// <summary>
    /// 管线方向与箭头对齐辅助方法
    /// 目的：无论路径顶点顺序如何（左->右或右->左），都能正确计算箭头朝向并在必要时翻转箭头模板。
    /// 使用说明：
    /// - 调用 AlignArrowToDirection(arrowTemplate, direction) 会返回一个新的 Polyline 实例（已旋转/翻转），原模板保持不变。
    /// - 在 ElementAndTable.cs 中将原有 AlignArrowToDirection(...) 实现替换为本工具（或直接改为调用本工具）。
    /// </summary>
    public static class PipeDirectionHelper
    {
        /// <summary>
        /// 计算路径方向向量（从第一个非重复点指向最后一个非重复点），并归一化。
        /// 如果所有点相同，返回 Z 轴零向量。
        /// </summary>
        ///
        public static Vector3d ComputePathDirectionVector(List<Point3d> orderedVertices, double tol = 1e-8)
        {
            if (orderedVertices == null || orderedVertices.Count == 0)
                return Vector3d.ZAxis; // 兜底

            // 找第一个与最后一个不重复的点（容差）
            Point3d first = orderedVertices.First();
            Point3d last = orderedVertices.Last();

            // 如果首尾相近则尝试找到真正的第一个和最后一个不重合点
            int i = 0;
            while (i < orderedVertices.Count - 1 && orderedVertices[i].DistanceTo(orderedVertices[i + 1]) <= tol) i++;
            int j = orderedVertices.Count - 1;
            while (j > 0 && orderedVertices[j].DistanceTo(orderedVertices[j - 1]) <= tol) j--;

            first = orderedVertices[i];
            last = orderedVertices[j];

            var vec = last - first;
            //if (vec.LengthSquared <= tol * tol)
            //    return Vector3d.ZAxis; // 无效向量

            return vec.GetNormal();
        }

        /// <summary>
        /// 将箭头模板对齐到给定方向并返回新 Polyline（模板不变）。
        /// 算法：
        ///  1) 计算模板的“主方向”（以模板顶点中 minX->maxX 构造向量）。
        ///  2) 计算路径方向（已规范化）。
        ///  3) 若两向量点积小于0，则需要额外翻转 180°（确保箭头朝向路径正向）。
        ///  4) 计算旋转角并绕模板质心旋转，返回新实例。
        /// 注意：对于 template 没有明显左右结构的情况，主方向由模板顶点的 minX/maxX 决定，通常适用于三角箭头或等长箭头模板。
        /// </summary>
        public static Polyline AlignArrowToDirection(Polyline arrowTemplate, Vector3d pathDirection)
        {
            if (arrowTemplate == null) throw new ArgumentNullException(nameof(arrowTemplate));

            // 复制模板（浅复制顶点数据到新 Polyline）
            var result = new Polyline();
            int vn = arrowTemplate.NumberOfVertices;
            for (int k = 0; k < vn; k++)
            {
                var p2d = arrowTemplate.GetPoint2dAt(k);
                result.AddVertexAt(k, p2d, arrowTemplate.GetBulgeAt(k), arrowTemplate.GetStartWidthAt(k), arrowTemplate.GetEndWidthAt(k));
            }

            // 计算模板质心（2D 平面近似）
            Point3d centroid = ComputePolylineCentroid(arrowTemplate);

            // 计算模板主方向：在模板顶点中查找 minX 和 maxX 点
            Point3d? minXPt = null;
            Point3d? maxXPt = null;
            double minX = double.PositiveInfinity, maxX = double.NegativeInfinity;
            for (int k = 0; k < vn; k++)
            {
                var p = arrowTemplate.GetPoint3dAt(k);
                if (p.X < minX) { minX = p.X; minXPt = p; }
                if (p.X > maxX) { maxX = p.X; maxXPt = p; }
            }

            // 如果无法得到主方向则退回使用 (1,0,0)
            Vector3d templateDir = Vector3d.XAxis;
            if (minXPt.HasValue && maxXPt.HasValue && !minXPt.Value.IsEqualTo(maxXPt.Value))
            {
                templateDir = (maxXPt.Value - minXPt.Value).GetNormal();
            }

            // 规范化路径方向（若非法则不旋转）
            Vector3d pd = pathDirection;
            if (pd.IsZeroLength())
            {
                // 不做旋转，直接返回复制品
                return result;
            }
            pd = pd.GetNormal();

            // 决定是否需要翻转（点积 < 0 表示模板朝向与路径方向相反）
            double dot = templateDir.DotProduct(pd);
            bool needFlip = dot < 0;

            // 计算当前角度与目标角度（以 XY 平面为准）
            double templateAngle = Math.Atan2(templateDir.Y, templateDir.X);
            double pathAngle = Math.Atan2(pd.Y, pd.X);

            // 基本旋转量
            double rotateBy = pathAngle - templateAngle;

            // 如果需要翻转，额外加 PI（180 度）
            if (needFlip)
            {
                rotateBy += Math.PI;
            }

            // 在质心处做旋转
            var m = Matrix3d.Displacement(centroid.GetAsVector().Negate()) // 移到原点
                    * Matrix3d.Rotation(rotateBy, Vector3d.ZAxis, Point3d.Origin)
                    * Matrix3d.Displacement(centroid.GetAsVector()); // 移回

            result.TransformBy(m);

            // 复制颜色/层等基本属性（视需要扩展）
            result.Layer = arrowTemplate.Layer;
            result.Color = arrowTemplate.Color;
            result.Linetype = arrowTemplate.Linetype;
            result.LineWeight = arrowTemplate.LineWeight;

            return result;
        }

        /// <summary>
        /// 计算 Polyline 的近似质心（基于顶点平均值，适用于简单箭头）
        /// </summary>
        private static Point3d ComputePolylineCentroid(Polyline pl)
        {
            if (pl == null) return Point3d.Origin;
            int n = pl.NumberOfVertices;
            if (n == 0) return Point3d.Origin;

            double sx = 0, sy = 0;
            for (int i = 0; i < n; i++)
            {
                var p = pl.GetPoint3dAt(i);
                sx += p.X; sy += p.Y;
            }
            return new Point3d(sx / n, sy / n, pl.GetPoint3dAt(0).Z);
        }

        /// <summary>
        /// 安全判断 Vector3d 是否长度为 0
        /// </summary>
        //private static bool IsZeroLength(this Vector3d v)
        //{
        //    return v.LengthSquared <= 1e-12;
        //}
    }
}
