using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using GB_NewCadPlus_IV.FunctionalMethod;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

// 兼容 DatabaseManager 内部嵌套类型：给类型起别名，避免全文件逐个替换
using FileStorage = GB_NewCadPlus_IV.FunctionalMethod.DatabaseManager.FileStorage;
using FileAttribute = GB_NewCadPlus_IV.FunctionalMethod.DatabaseManager.FileAttribute;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 用于从CAD选择导入时传输数据的DTO
    /// </summary>
    public class ImportEntityDto
    {
        /// <summary>
        /// 文件存储信息（对应 cad_file_storage 表）
        /// </summary>
        public FileStorage FileStorage { get; set; } = new FileStorage();

        /// <summary>
        /// 文件属性信息（旧模型，过渡期保留，不再作为上传写库主入口）
        /// </summary>
        public FileAttribute FileAttribute { get; set; } = new FileAttribute();

        /// <summary>
        /// 属性业务ID（兼容字段）
        /// </summary>
        public string? FileAttributeId { get; set; }

        /// <summary>
        /// 预览图路径
        /// </summary>
        public string? PreviewImagePath { get; set; }

        /// <summary>
        /// JSON属性字典（新上传入口使用）
        /// </summary>
        public Dictionary<string, string> AttributesJson { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 负责从当前CAD图形中选择实体并提取信息的辅助类
    /// </summary>
    public static class SelectionImportHelper
    {
        /// <summary>
        /// 交互式地在CAD中选择一个实体，并返回包含其信息的DTO。
        /// 此方法应在后台线程或UI线程中调用，它内部会处理DocumentLock。
        /// </summary>
        public static ImportEntityDto PickAndReadEntity()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                throw new InvalidOperationException("没有活动的CAD文档。");
            }

            ImportEntityDto dto = null;

            // 必须在文档锁定的情况下与CAD数据库交互
            using (doc.LockDocument())
            {
                var ed = doc.Editor;
                var db = doc.Database;

                var peo = new PromptEntityOptions("\n请选择要导入的图元（块、属性块等）：");
                var per = ed.GetEntity(peo);

                if (per.Status != PromptStatus.OK)
                {
                    return null; // 用户取消
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var entity = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                    if (entity == null)
                    {
                        tr.Commit();
                        return null;
                    }

                    dto = new ImportEntityDto();
                    var fs = dto.FileStorage;
                    var fa = dto.FileAttribute;

                    // 填充通用信息
                    fs.FilePath = doc.Name; // 当前DWG文件路径
                    fs.FileType = ".dwg";
                    fs.CreatedAt = DateTime.Now;
                    fs.UpdatedAt = DateTime.Now;
                    fs.IsActive = 1;
                    fs.IsPublic = 1;
                    fs.CreatedBy = Environment.UserName;
                    fa.CreatedAt = DateTime.Now;
                    fa.UpdatedAt = DateTime.Now;

                    // 针对不同实体类型提取信息
                    if (entity is BlockReference br)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                        fs.DisplayName = btr.Name;
                        fs.BlockName = btr.Name;
                        fs.LayerName = br.Layer;
                        fs.ColorIndex = br.Color.ColorIndex;
                        fs.Scale = br.ScaleFactors.X; // 简化，只取X向缩放

                        fa.Angle = (decimal)br.Rotation;
                        fa.BasePointX = (decimal)br.Position.X;
                        fa.BasePointY = (decimal)br.Position.Y;
                        fa.BasePointZ = (decimal)br.Position.Z;

                        // 提取属性
                        if (br.AttributeCollection.Count > 0)
                        {
                            var attributesText = new List<string>();
                            foreach (ObjectId attId in br.AttributeCollection)
                            {
                                var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                if (attRef != null)
                                {
                                    attributesText.Add($"{attRef.Tag}: {attRef.TextString}");
                                }
                            }
                            fa.Remarks = string.Join("\n", attributesText);
                        }
                    }
                    else
                    {
                        fs.DisplayName = entity.GetRXClass().Name;
                        fs.LayerName = entity.Layer;
                        fs.ColorIndex = entity.Color.ColorIndex;
                    }

                    fs.FileName = fs.DisplayName; // 默认文件名等于显示名
                    fa.FileName = fs.DisplayName;

                    // 估算尺寸
                    try
                    {
                        var ext = entity.GeometricExtents;
                        fa.Length = (decimal)(ext.MaxPoint.X - ext.MinPoint.X);
                        fa.Width = (decimal)(ext.MaxPoint.Y - ext.MinPoint.Y);
                        fa.Height = (decimal)(ext.MaxPoint.Z - ext.MinPoint.Z);
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogWarning($"估算尺寸失败: {ex.Message}");
                    }

                    // 生成预览图片
                    string previewPath = PreparePreviewPath();

                    try
                    {
                        // 保存当前视图
                        ViewTableRecord originalView = ed.GetCurrentView();

                        // --- 获取实体在世界坐标下的包围盒 ---
                        Extents3d? worldExt = null;
                        try
                        {
                            if (TryGetEntityWorldExtents(tr, entity, out var extWorld))
                                worldExt = extWorld;
                        }
                        catch (Exception exExt)
                        {
                            LogManager.Instance.LogWarning($"TryGetEntityWorldExtents 失败: {exExt.Message}");
                            worldExt = null;
                        }

                        // 如果能拿到包围盒，设置合适视图并多次重试截图；否则尝试临时插入回退再全图截图
                        bool captured = false;
                        if (worldExt != null)
                        {
                            // 记录包围盒信息用于调试
                            var min = worldExt.Value.MinPoint;
                            var max = worldExt.Value.MaxPoint;
                            double extWidth = max.X - min.X;
                            double extHeight = max.Y - min.Y;
                            LogManager.Instance.LogInfo($"实体包围盒: Min({min.X:F2},{min.Y:F2}), Max({max.X:F2},{max.Y:F2}), 尺寸({extWidth:F2},{extHeight:F2})");

                            captured = TryCaptureUsingView(doc, ed, worldExt.Value, previewPath);
                            if (captured)
                            {
                                dto.PreviewImagePath = previewPath;
                                fs.PreviewImagePath = previewPath;
                                fs.PreviewImageName = Path.GetFileName(previewPath);
                                LogManager.Instance.LogInfo($"生成按实体包围盒的预览图: {previewPath}");
                            }
                            else
                            {
                                LogManager.Instance.LogWarning("按实体包围盒捕获预览图失败，尝试临时插入回退。");
                            }
                        }

                        // 回退方案：临时插入实体副本到 ModelSpace，移动到原点并 ZoomExtents 后截图
                        if (!captured)
                        {
                            try
                            {
                                if (TryCaptureByTemporaryInsertion(doc, ed, db, entity, worldExt, previewPath))
                                {
                                    dto.PreviewImagePath = previewPath;
                                    fs.PreviewImagePath = previewPath;
                                    fs.PreviewImageName = Path.GetFileName(previewPath);
                                    captured = true;
                                    LogManager.Instance.LogInfo($"通过临时插入生成预览图: {previewPath}");
                                }
                                else
                                {
                                    LogManager.Instance.LogWarning("临时插入回退方案未能生成预览图，尝试全图截图回退。");
                                }
                            }
                            catch (Exception exTemp)
                            {
                                LogManager.Instance.LogWarning($"临时插入回退异常: {exTemp.Message}");
                            }
                        }

                        // 最后退路：全图截图
                        if (!captured)
                        {
                            Bitmap previewBmp = null;
                            for (int i = 0; i < 3 && previewBmp == null; i++)
                            {
                                try
                                {
                                    previewBmp = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.CapturePreviewImage(doc, 800, 600);
                                }
                                catch (Exception ex)
                                {
                                    LogManager.Instance.LogWarning($"全图 CapturePreviewImage 尝试 {i + 1} 失败: {ex.Message}");
                                    previewBmp = null;
                                }

                                if (previewBmp == null)
                                {
                                    System.Threading.Thread.Sleep(150);
                                    try { Application.UpdateScreen(); } catch { }
                                }
                            }

                            if (previewBmp != null)
                            {
                                try
                                {
                                    previewBmp.Save(previewPath, System.Drawing.Imaging.ImageFormat.Png);
                                    dto.PreviewImagePath = previewPath;
                                    fs.PreviewImagePath = previewPath;
                                    fs.PreviewImageName = Path.GetFileName(previewPath);
                                    LogManager.Instance.LogInfo($"无包围盒或其它回退失败，已生成全图预览: {previewPath}");
                                }
                                finally
                                {
                                    previewBmp.Dispose();
                                }
                            }
                            else
                            {
                                LogManager.Instance.LogWarning("生成预览图失败（全图捕获也失败）");
                            }
                        }

                        // 恢复原视图
                        try { ed.SetCurrentView(originalView); } catch (Exception ex) { LogManager.Instance.LogWarning($"恢复原视图失败: {ex.Message}"); }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogWarning($"生成预览图异常: {ex.Message}");
                    }

                    tr.Commit();
                }
            }
            return dto;
        }
        /// <summary>
        /// 新增的辅助校验方法（放在类内）
        /// </summary>
        /// <param name="ext"></param>
        /// <returns></returns>
        private static bool IsExtentsReasonable(Extents3d ext)
        {
            try
            {
                var min = ext.MinPoint;
                var max = ext.MaxPoint;
                if (double.IsNaN(min.X) || double.IsNaN(min.Y) || double.IsNaN(max.X) || double.IsNaN(max.Y)) return false;
                double w = max.X - min.X;
                double h = max.Y - min.Y;
                if (!(w > 0 && h > 0)) return false;
                // 合理阈值：避免极小或极大导致异常视图（可按需调整）
                const double minSize = 1e-4;
                const double maxSize = 1e7;
                if (w < minSize || h < minSize) return false;
                if (w > maxSize || h > maxSize) return false;
                // 防止过分长/瘦或过分扁的范围（避免宽高比极端）
                double aspect = Math.Max(w / h, h / w);
                if (aspect > 1000.0) return false;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 辅助：尝试获取实体在世界坐标下的包围盒（支持 BlockReference 的特殊计算）
        /// </summary>
        /// <param name="tr">事务</param>
        /// <param name="ent">实体</param>
        /// <param name="worldExt">输出的世界坐标包围盒</param>
        /// <returns>是否成功获取包围盒</returns>
        private static bool TryGetEntityWorldExtents(Transaction tr, Entity ent, out Extents3d worldExt)
        {
            worldExt = new Extents3d();
            try
            {
                // 优先尝试实体自身的 GeometricExtents（多数情况下最可靠）
                //try
                //{
                //    var eExt = ent.GeometricExtents;
                //    if (IsExtentsReasonable(eExt))
                //    {
                //        worldExt = eExt;
                //        return true;
                //    }
                //    else
                //    {
                //        LogManager.Instance.LogWarning($"实体自身 GeometricExtents 非常规：{eExt.MinPoint} - {eExt.MaxPoint}，将尝试其他方法。");
                //    }
                //}
                //catch (Exception ex)
                //{
                //    LogManager.Instance.LogWarning($"常规获取几何边界失败: {ex.Message}");
                //}

                // 如果是块参照，尝试基于块定义逐个实体变换出世界包围盒（更鲁棒）
                if (ent is BlockReference br)
                {
                    if (TryComputeBlockReferenceWorldExtents(tr, br, out var extFromBtr) && IsExtentsReasonable(extFromBtr))
                    {
                        worldExt = extFromBtr;
                        return true;
                    }
                    else
                    {
                        LogManager.Instance.LogWarning("基于块定义计算的 worldExt 非常规或计算失败，回退其它方案。");
                        return false;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"TryGetEntityWorldExtents 处理失败: {ex.Message}");
                worldExt = new Extents3d();
                return false;
            }
        }

        /// <summary>
        /// 辅助：计算 BlockReference 在世界坐标下的包围盒（遍历 BlockTableRecord 内部实体并变换）
        /// </summary>
        /// <param name="tr">事务</param>
        /// <param name="br">块参照</param>
        /// <param name="worldExt">输出的世界坐标包围盒</param>
        /// <returns>是否成功计算包围盒</returns>
        private static bool TryComputeBlockReferenceWorldExtents(Transaction tr, BlockReference br, out Extents3d worldExt)
        {
            // 顶层入口：把 br.BlockTransform 作为初始累积变换传入实际处理方法
            worldExt = new Extents3d();
            try
            {
                return TryComputeBlockReferenceWorldExtents_Internal(tr, br, br.BlockTransform, out worldExt);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"TryComputeBlockReferenceWorldExtents 入口异常: {ex.Message}");
                worldExt = new Extents3d();
                return false;
            }
        }
        // 内部实现：传入累积变换 accumulatedTransform（相对于最终世界）
        private static bool TryComputeBlockReferenceWorldExtents_Internal(Transaction tr, BlockReference br, Matrix3d accumulatedTransform, out Extents3d worldExt)
        {
            worldExt = new Extents3d();
            try
            {
                var btrId = br.BlockTableRecord;
                if (btrId.IsNull) return false;
                var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                if (btr == null) return false;

                var allPts = new List<Point3d>();

                Point3d[] Get8Points(Extents3d e)
                {
                    return new[]
                    {
                new Point3d(e.MinPoint.X, e.MinPoint.Y, e.MinPoint.Z),
                new Point3d(e.MinPoint.X, e.MinPoint.Y, e.MaxPoint.Z),
                new Point3d(e.MinPoint.X, e.MaxPoint.Y, e.MinPoint.Z),
                new Point3d(e.MinPoint.X, e.MaxPoint.Y, e.MaxPoint.Z),
                new Point3d(e.MaxPoint.X, e.MinPoint.Y, e.MinPoint.Z),
                new Point3d(e.MaxPoint.X, e.MinPoint.Y, e.MaxPoint.Z),
                new Point3d(e.MaxPoint.X, e.MaxPoint.Y, e.MinPoint.Z),
                new Point3d(e.MaxPoint.X, e.MaxPoint.Y, e.MaxPoint.Z),
            };
                }

                foreach (ObjectId id in btr)
                {
                    try
                    {
                        if (id.IsNull) continue;
                        var subObj = tr.GetObject(id, OpenMode.ForRead);
                        if (subObj == null) continue;

                        // 如果子对象是一个块参照，需要把它的 BlockTransform 与当前累积变换合成
                        if (subObj is BlockReference nestedBr)
                        {
                            // 修复点：变换合成顺序应为 nested.BlockTransform * parentAccumulatedTransform
                            // （先将子块定义坐标变换到父定义坐标，再由父的累积变换到最终世界）
                            var nestedAccum = nestedBr.BlockTransform * accumulatedTransform;

                            if (TryComputeBlockReferenceWorldExtents_Internal(tr, nestedBr, nestedAccum, out var nestedExt))
                            {
                                // 递归返回的 nestedExt 已是相对于 nestedAccum（即已在顶层坐标系），直接收集角点
                                var pts = Get8Points(nestedExt);
                                foreach (var p in pts)
                                {
                                    if (!double.IsNaN(p.X) && !double.IsInfinity(p.X))
                                        allPts.Add(p);
                                }
                            }
                            else
                            {
                                // 递归失败则尝试读取 nestedBr 的 GeometricExtents 并按 nestedAccum 转换
                                try
                                {
                                    var nExt = nestedBr.GeometricExtents;
                                    var pts = Get8Points(nExt);
                                    for (int i = 0; i < pts.Length; i++)
                                        pts[i] = pts[i].TransformBy(nestedAccum);
                                    foreach (var tp in pts)
                                    {
                                        if (!double.IsNaN(tp.X) && !double.IsInfinity(tp.X))
                                            allPts.Add(tp);
                                    }
                                }
                                catch
                                {
                                    LogManager.Instance.LogWarning("嵌套 BlockReference 无法获取 extents，已跳过该子项。");
                                    continue;
                                }
                            }
                        }
                        else if (subObj is Entity subEnt)
                        {
                            // 普通子实体：其 GeometricExtents 在定义坐标系，按 accumulatedTransform 变换到顶层世界
                            try
                            {
                                var subExt = subEnt.GeometricExtents;
                                var pts = Get8Points(subExt);
                                for (int i = 0; i < pts.Length; i++)
                                {
                                    var tp = pts[i].TransformBy(accumulatedTransform);
                                    if (double.IsNaN(tp.X) || double.IsInfinity(tp.X)) continue;
                                    const double outlierThreshold = 1e8;
                                    if (Math.Abs(tp.X) > outlierThreshold || Math.Abs(tp.Y) > outlierThreshold || Math.Abs(tp.Z) > outlierThreshold)
                                    {
                                        LogManager.Instance.LogWarning($"忽略可能的离群点: ({tp.X:F2},{tp.Y:F2},{tp.Z:F2})");
                                        continue;
                                    }
                                    allPts.Add(tp);
                                }
                            }
                            catch (Exception ex)
                            {
                                LogManager.Instance.LogWarning($"获取子实体几何边界失败: {ex.Message} (类型: {subEnt.GetType().Name})，跳过该子实体。");
                                continue;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogWarning($"处理子实体时出错: {ex.Message}，继续处理其它子实体。");
                    }
                }

                if (allPts.Count == 0) return false;

                // 计算 min/max
                double minX = double.PositiveInfinity, minY = double.PositiveInfinity, minZ = double.PositiveInfinity;
                double maxX = double.NegativeInfinity, maxY = double.NegativeInfinity, maxZ = double.NegativeInfinity;
                foreach (var p in allPts)
                {
                    if (p.X < minX) minX = p.X;
                    if (p.Y < minY) minY = p.Y;
                    if (p.Z < minZ) minZ = p.Z;
                    if (p.X > maxX) maxX = p.X;
                    if (p.Y > maxY) maxY = p.Y;
                    if (p.Z > maxZ) maxZ = p.Z;
                }
                worldExt = new Extents3d(new Point3d(minX, minY, minZ), new Point3d(maxX, maxY, maxZ));

                return true;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"TryComputeBlockReferenceWorldExtents_Internal 失败: {ex.Message}");
                worldExt = new Extents3d();
                return false;
            }
        }

        /// <summary>
        /// 在已知包围盒的情况下设置视图并尝试截图（返回是否成功）
        /// </summary>
        /// <param name="doc">文档</param>
        /// <param name="ed">编辑器</param>
        /// <param name="ext">包围盒</param>
        /// <param name="outPreviewPath">输出预览图路径</param>
        /// <returns>是否截图成功</returns>
        private static bool TryCaptureUsingView(Document doc, Editor ed, Extents3d ext, string outPreviewPath)
        {
            try
            {
                //1:合法性检查:先用 IsExtentsReasonable(Extents3d) 过滤异常/极端的包围盒（NaN、极小/极大、极端长宽比等），不合理则放弃按 ext 截图并返回 false。
                if (!IsExtentsReasonable(ext))
                {
                    LogManager.Instance.LogWarning($"TryCaptureUsingView: 提供的 ext 不合理，放弃按 ext 截图: Min{ext.MinPoint} Max{ext.MaxPoint}");
                    return false;
                }

                //2. 计算包围盒与内容世界尺寸 
                var min = ext.MinPoint;
                var max = ext.MaxPoint;

                // 原始内容世界尺寸（不含 margin）•	防止宽或高为 0 导致后续除零或视图尺寸为零的问题，最小取 1.0（世界单位）。
                double contentWidthWorld = Math.Max(max.X - min.X, 1.0);
                double contentHeightWorld = Math.Max(max.Y - min.Y, 1.0);

                // 3.自适应 world margin  按实体较大边的百分比，•	margin = 较大边的 5%（相对）但至少 0.5，最多 10。目的是既保证边距又避免 margin 过大或过小。
                double worldMargin = Math.Min(10.0, Math.Max(Math.Max(contentWidthWorld, contentHeightWorld) * 0.05, 0.5));

                // 4.计算最终视图的世界尺寸（包含 margin）视图尺寸 = 内容 + margin*2 
                double contentWidth = contentWidthWorld + worldMargin * 2;
                double contentHeight = contentHeightWorld + worldMargin * 2;

                // 5.计算视图中心（世界坐标，Z 设为 0）•	注意：这里 Z 直接设为 0，意味着截屏视图是顶视并以 Z=0 平面作为投影中心。
                var centerWCS = new Point3d((min.X + max.X) / 2.0, (min.Y + max.Y) / 2.0, 0);

                //6. 备份并克隆当前视图 •	克隆后在 newView 上修改，方便后续恢复 originalView。
                var originalView = ed.GetCurrentView();
                var newView = (ViewTableRecord)originalView.Clone();

                //7. 配置新视图为顶视（默认朝向），置中并设置宽高（世界单位）
                newView.ViewDirection = Vector3d.ZAxis;
                newView.ViewTwist = 0.0;
                newView.CenterPoint = new Point2d(centerWCS.X, centerWCS.Y);
                newView.Width = contentWidth;
                newView.Height = contentHeight;

                //8. 应用视图并等待渲染 •	给 AutoCAD 一点时间刷新并重绘画面，之后代码会去捕获截图。
                ed.SetCurrentView(newView);
                System.Threading.Thread.Sleep(200);
                try { Application.UpdateScreen(); } catch { }

                const int baseSize = 600;
                uint outputWidth, outputHeight;
                double aspectRatio = contentWidth / contentHeight;

                if (aspectRatio >= 1.0)
                {
                    outputWidth = baseSize;
                    outputHeight = (uint)(baseSize / aspectRatio);
                }
                else
                {
                    outputHeight = baseSize;
                    outputWidth = (uint)(baseSize * aspectRatio);
                }

                outputWidth = Math.Max(outputWidth, 300u);
                outputHeight = Math.Max(outputHeight, 300u);

                Bitmap previewBmp = CaptureImage(doc, outputWidth, outputHeight);
                if (previewBmp != null)
                {
                    try
                    {
                        // 这里传入未包含 margin 的世界宽高（contentWidthWorld/contentHeightWorld）
                        using (var cropped = CropPreviewToContent(previewBmp, worldMargin, contentWidthWorld, contentHeightWorld, outputWidth, outputHeight))
                        {
                            cropped.Save(outPreviewPath, System.Drawing.Imaging.ImageFormat.Png);
                        }
                        LogManager.Instance.LogInfo($"截图成功: 中心WCS({centerWCS.X:F2},{centerWCS.Y:F2}), 视图尺寸({contentWidth:F2}×{contentHeight:F2}), 图像尺寸({outputWidth}×{outputHeight}), worldMargin({worldMargin:F2})");
                    }
                    finally
                    {
                        previewBmp.Dispose();
                    }
                    return true;
                }
                else
                {
                    LogManager.Instance.LogWarning($"截图失败: 中心WCS({centerWCS.X:F2},{centerWCS.Y:F2}), 视图尺寸({contentWidth:F2}×{contentHeight:F2})");
                }

                try { ed.SetCurrentView(originalView); } catch (Exception ex) { LogManager.Instance.LogWarning($"恢复原视图失败: {ex.Message}"); }

                return false;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"TryCaptureUsingView 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 回退：在当前图形中临时插入实体副本（移动到原点）、ZoomExtents 并截图，最后删除副本
        /// </summary>
        /// <param name="doc">文档</param>
        /// <param name="ed">编辑器</param>
        /// <param name="db">数据库</param>
        /// <param name="sourceEntity">源实体</param>
        /// <param name="sourceWorldExt">源实体的包围盒</param>
        /// <param name="previewPath">预览图路径</param>
        /// <returns>是否截图成功</returns>
        private static bool TryCaptureByTemporaryInsertion(Document doc, Editor ed, Database db, Entity sourceEntity, Extents3d? sourceWorldExt, string previewPath)
        {
            bool success = false;
            ObjectId cloneId = ObjectId.Null;
            ViewTableRecord originalView = null;

            try
            {
                Extents3d worldExt;
                if (sourceWorldExt != null)
                    worldExt = sourceWorldExt.Value;
                else
                {
                    try { worldExt = sourceEntity.GeometricExtents; }
                    catch (Exception ex) { LogManager.Instance.LogWarning($"读取 GeometricExtents 失败: {ex.Message}"); return false; }
                }

                var min = worldExt.MinPoint;
                var max = worldExt.MaxPoint;
                var center3d = new Point3d((min.X + max.X) / 2.0, (min.Y + max.Y) / 2.0, (min.Z + max.Z) / 2.0);
                var displacement = Matrix3d.Displacement(Point3d.Origin - center3d);

                using (var wtr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)wtr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var ms = (BlockTableRecord)wtr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    var cloned = sourceEntity.Clone() as Entity;
                    if (cloned == null) { wtr.Abort(); return false; }

                    try { cloned.TransformBy(displacement); }
                    catch (Exception ex) { LogManager.Instance.LogWarning($"变换实体失败: {ex.Message}"); }

                    cloneId = ms.AppendEntity(cloned);
                    wtr.AddNewlyCreatedDBObject(cloned, true);
                    wtr.Commit();
                }

                if (cloneId.IsNull) return false;

                // 读取克隆实体实际 extents（WCS）
                Extents3d cloneExt;
                using (var tr2 = db.TransactionManager.StartTransaction())
                {
                    var ent = tr2.GetObject(cloneId, OpenMode.ForRead) as Entity;
                    if (ent == null) { tr2.Commit(); return false; }
                    try { cloneExt = ent.GeometricExtents; }
                    catch (Exception ex) { LogManager.Instance.LogWarning($"读取克clone实体 GeometricExtents 失败: {ex.Message}"); tr2.Commit(); return false; }
                    tr2.Commit();
                }

                var cmin = cloneExt.MinPoint;
                var cmax = cloneExt.MaxPoint;
                double entityWidthWorld = Math.Max(cmax.X - cmin.X, 1.0);
                double entityHeightWorld = Math.Max(cmax.Y - cmin.Y, 1.0);

                double worldMargin = Math.Min(10.0, Math.Max(Math.Max(entityWidthWorld, entityHeightWorld) * 0.05, 0.5));
                double contentWidth = entityWidthWorld + worldMargin * 2;
                double contentHeight = entityHeightWorld + worldMargin * 2;

                originalView = ed.GetCurrentView();
                var newView = (ViewTableRecord)originalView.Clone();

                newView.ViewDirection = Vector3d.ZAxis;
                newView.ViewTwist = 0.0;
                var center = new Point3d((cmin.X + cmax.X) / 2.0, (cmin.Y + cmax.Y) / 2.0, 0);
                newView.CenterPoint = new Point2d(center.X, center.Y);
                newView.Width = contentWidth;
                newView.Height = contentHeight;

                ed.SetCurrentView(newView);
                System.Threading.Thread.Sleep(200);
                try { Application.UpdateScreen(); } catch { }

                const int baseSize = 600;
                uint outputWidth, outputHeight;
                double aspectRatio = contentWidth / contentHeight;
                if (aspectRatio >= 1.0) { outputWidth = baseSize; outputHeight = (uint)(baseSize / aspectRatio); }
                else { outputHeight = baseSize; outputWidth = (uint)(baseSize * aspectRatio); }
                outputWidth = Math.Max(outputWidth, 300u);
                outputHeight = Math.Max(outputHeight, 300u);

                Bitmap previewBmp = CaptureImage(doc, outputWidth, outputHeight);
                if (previewBmp != null)
                {
                    try
                    {
                        // 这里也传入未包含 margin 的世界宽高（entityWidthWorld/entityHeightWorld）
                        using (var cropped = CropPreviewToContent(previewBmp, worldMargin, entityWidthWorld, entityHeightWorld, outputWidth, outputHeight))
                        {
                            cropped.Save(previewPath, System.Drawing.Imaging.ImageFormat.Png);
                        }
                        LogManager.Instance.LogInfo($"临时插入截图成功: 视图中心WCS({center.X:F2},{center.Y:F2}), 视图尺寸({contentWidth:F2}×{contentHeight:F2}), 图像尺寸({outputWidth}×{outputHeight}), worldMargin({worldMargin:F2})");
                        success = true;
                    }
                    finally { previewBmp.Dispose(); }
                }
                else
                {
                    LogManager.Instance.LogWarning($"临时插入截图失败: 视图中心WCS({center.X:F2},{center.Y:F2}), 视图尺寸({contentWidth:F2}×{contentHeight:F2})");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"TryCaptureByTemporaryInsertion 异常: {ex.Message}");
                success = false;
            }
            finally
            {
                if (!cloneId.IsNull)
                {
                    try
                    {
                        using (var wtr2 = db.TransactionManager.StartTransaction())
                        {
                            var ent = wtr2.GetObject(cloneId, OpenMode.ForWrite) as Entity;
                            if (ent != null) ent.Erase();
                            wtr2.Commit();
                        }
                    }
                    catch (Exception ex) { LogManager.Instance.LogWarning($"删除临时克隆失败: {ex.Message}"); }
                }

                try { if (originalView != null) ed.SetCurrentView(originalView); }
                catch (Exception ex) { LogManager.Instance.LogWarning($"恢复原视图失败: {ex.Message}"); }
            }

            return success;
        }

        /// <summary>
        /// 尝试将业务属性ID回写到 FileStorage.FileAttributeId（兼容 int / string 类型）
        /// </summary>
        /// <param name="fileStorage">文件存储对象</param>
        /// <param name="bizId">业务属性ID</param>
        private static void TrySetFileStorageAttributeId(FileStorage fileStorage, string bizId)
        {
            // 防御性校验，避免空对象或空业务ID
            if (fileStorage == null || string.IsNullOrWhiteSpace(bizId))
            {
                return;
            }

            // 当前 FileStorage.FileAttributeId 已是 string，直接赋值最稳妥
            fileStorage.FileAttributeId = bizId.Trim();
        }

        /// <summary>
        /// 准备预览图路径
        /// </summary>
        /// <returns>预览图路径</returns>
        private static string PreparePreviewPath()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "GB_NewCadPlus_IV_Previews");
            Directory.CreateDirectory(tempDir);
            return Path.Combine(tempDir, $"preview_{Guid.NewGuid()}.png");
        }

        /// <summary>
        /// 捕获图像
        /// </summary>
        /// <param name="doc">文档</param>
        /// <param name="width">图像宽度</param>
        /// <param name="height">图像高度</param>
        /// <returns>捕获的图像</returns>
        private static Bitmap CaptureImage(Document doc, uint width, uint height)
        {
            Bitmap previewBmp = null;
            for (int i = 0; i < 3 && previewBmp == null; i++)
            {
                try
                {
                    previewBmp = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.CapturePreviewImage(doc, width, height);
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogWarning($"CapturePreviewImage 尝试 {i + 1} 失败: {ex.Message}");
                    previewBmp = null;
                }

                if (previewBmp == null)
                {
                    System.Threading.Thread.Sleep(150);
                    try { Application.UpdateScreen(); } catch { }
                }
            }
            return previewBmp;
        }

        /// <summary>
        /// 对捕获到的 Bitmap 进行基于像素的紧缩裁剪，然后按世界单位的 margin 转为像素追加边距并返回新的 Bitmap。
        /// 采用图像四角像素投票决定背景色，使用欧氏颜色距离并允许安全像素，防止抗锯齿与单点异常导致裁剪过度。
        /// </summary>
        private static Bitmap CropPreviewToContent(Bitmap src, double worldMargin, double contentWidthWorld, double contentHeightWorld, uint outputWidth, uint outputHeight)
        {
            
            if (src == null) throw new ArgumentNullException(nameof(src));

            int w = src.Width;
            int h = src.Height;

            #region 计算非背景像素的最小包围盒
            // --- 背景颜色判定：从四个角采样并取多数（fallback 为第一个） ---
            System.Drawing.Color[] cornerSamples = new System.Drawing.Color[4];
            cornerSamples[0] = src.GetPixel(0, 0);
            cornerSamples[1] = src.GetPixel(Math.Max(0, w - 1), 0);
            cornerSamples[2] = src.GetPixel(0, Math.Max(0, h - 1));
            cornerSamples[3] = src.GetPixel(Math.Max(0, w - 1), Math.Max(0, h - 1));

            // 找到出现次数最多的颜色（简单多数投票）
            System.Drawing.Color bg = cornerSamples[0];
            var counts = new Dictionary<int, int>();
            for (int i = 0; i < cornerSamples.Length; i++)
            {
                int key = (cornerSamples[i].ToArgb());
                if (!counts.ContainsKey(key)) counts[key] = 0;
                counts[key]++;
            }
            int maxCount = 0;
            foreach (var kv in counts)
            {
                if (kv.Value > maxCount)
                {
                    maxCount = kv.Value;
                    bg = System.Drawing.Color.FromArgb(kv.Key);
                }
            }

            // 容差（颜色欧氏距离阈值），可按需调整
            const int colorTol = 20;
            bool IsBg(System.Drawing.Color c)
            {
                int dr = c.R - bg.R;
                int dg = c.G - bg.G;
                int db = c.B - bg.B;
                // 欧氏距离（不开平方），比较时平方阈值
                int dist2 = dr * dr + dg * dg + db * db;
                return dist2 <= colorTol * colorTol;
            }

            // 扫描找出非背景像素的最小包围像素框
            int minX = w, minY = h, maxX = -1, maxY = -1;
            for (int y = 0; y < h; y++)
            {
                for (int x = 0; x < w; x++)
                {
                    var px = src.GetPixel(x, y); // 返回 System.Drawing.Color
                    if (!IsBg(px))
                    {
                        if (x < minX) minX = x;
                        if (y < minY) minY = y;
                        if (x > maxX) maxX = x;
                        if (y > maxY) maxY = y;
                    }
                }
            }

            #endregion

            // 如果没有检测到任何非背景像素，则直接返回原图的副本
            if (maxX < minX || maxY < minY)
            {
                return new Bitmap(src);
            }

            // 关键修正：基于包含 margin 的世界尺寸计算像素比例（与截图时使用的视图尺寸一致）
            double contentWidth = contentWidthWorld + worldMargin * 2.0;
            double contentHeight = contentHeightWorld + worldMargin * 2.0;

            double scaleX = (double)outputWidth / contentWidth;
            double scaleY = (double)outputHeight / contentHeight;
            // 仍然使用 min 以防极端纵横比
            double scale = Math.Min(scaleX, scaleY);

            int pixelMargin = (int)Math.Ceiling(worldMargin * scale);
            // 上限保护：避免对小实体产生极大的像素 margin（导致主体被挤到一侧或图像看似偏移）
            int maxAllowedMargin = Math.Max((int)Math.Min(outputWidth, outputHeight) / 4, 8); // 最多占图像最小边的 1/4
            if (pixelMargin > maxAllowedMargin) pixelMargin = maxAllowedMargin;
            if (pixelMargin < 1) pixelMargin = 1;
            // 额外安全像素，防止抗锯齿被裁掉
            const int safetyPx = 3;

            // 期望裁剪区域（包含 world margin）
            int desiredX = Math.Max(0, minX - pixelMargin - safetyPx);
            int desiredY = Math.Max(0, minY - pixelMargin - safetyPx);
            int desiredW = Math.Min(w - desiredX, (maxX - minX + 1) + (pixelMargin + safetyPx) * 2);
            int desiredH = Math.Min(h - desiredY, (maxY - minY + 1) + (pixelMargin + safetyPx) * 2);

            // 额外检查：如果因为某些比例问题导致最终像素 margin 太小（小于 pixelMargin），则强制扩展
            // 计算当前像素 margin 实测（相对于原 contentWidthWorld 所计算的像素比例）
            // 这里我们试图保证在横纵方向至少满足 pixelMargin
            int leftMarginPx = minX - desiredX;
            int rightMarginPx = (desiredX + desiredW - 1) - maxX;
            int topMarginPx = minY - desiredY;
            int bottomMarginPx = (desiredY + desiredH - 1) - maxY;

            int needExpandLeft = Math.Max(0, pixelMargin - leftMarginPx);
            int needExpandRight = Math.Max(0, pixelMargin - rightMarginPx);
            int needExpandTop = Math.Max(0, pixelMargin - topMarginPx);
            int needExpandBottom = Math.Max(0, pixelMargin - bottomMarginPx);

            // 计算最终裁切框并 clamp 到图片范围
            int finalX = Math.Max(0, desiredX - needExpandLeft);
            int finalY = Math.Max(0, desiredY - needExpandTop);
            int finalW = desiredW + needExpandLeft + needExpandRight;
            int finalH = desiredH + needExpandTop + needExpandBottom;

            if (finalX + finalW > w) finalW = w - finalX;
            if (finalY + finalH > h) finalH = h - finalY;
            if (finalW <= 0 || finalH <= 0)
            {
                // 兜底：返回原图副本
                return new Bitmap(src);
            }

            var bmp = new Bitmap(finalW, finalH);
            using (var g = Graphics.FromImage(bmp))
            {
                g.DrawImage(src, new Rectangle(0, 0, finalW, finalH), new Rectangle(finalX, finalY, finalW, finalH), GraphicsUnit.Pixel);
            }

            LogManager.Instance.LogInfo($"CropPreviewToContent: src({w}x{h}) contentBox([{minX},{minY}]-[{maxX},{maxY}]) finalCrop([{finalX},{finalY},{finalW},{finalH}]) pixelMargin({pixelMargin})");

            return bmp;
        }
    }
}
