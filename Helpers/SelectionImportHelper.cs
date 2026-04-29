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
            // 获取当前活动CAD文档
            var doc = Application.DocumentManager.MdiActiveDocument;

            // 如果没有活动文档，直接抛异常
            if (doc == null)
            {
                throw new InvalidOperationException("没有活动的CAD文档。");
            }

            // 初始化返回对象
            ImportEntityDto dto = null;

            // 锁定当前文档，避免在访问CAD数据库时出现冲突
            using (doc.LockDocument())
            {
                // 获取编辑器和数据库对象
                var ed = doc.Editor;
                var db = doc.Database;

                // 弹出实体选择提示
                var peo = new PromptEntityOptions("\n请选择要导入的图元（块、属性块等）：");
                var per = ed.GetEntity(peo);

                // 如果用户取消选择，则返回空
                if (per.Status != PromptStatus.OK)
                {
                    return null;
                }

                // 开启事务读取实体信息
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // 读取用户选中的实体
                    var entity = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;

                    // 如果实体为空，则直接返回空
                    if (entity == null)
                    {
                        tr.Commit();
                        return null;
                    }

                    // 初始化DTO、文件主信息对象、旧属性对象、JSON属性字典
                    dto = new ImportEntityDto();
                    var fs = dto.FileStorage;
                    var fa = dto.FileAttribute;
                    var attrs = dto.AttributesJson;

                    // 局部函数，安全写入文本属性到JSON字典
                    void AddTextAttr(string key, string value)
                    {
                        // 属性名或属性值为空时不写入
                        if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(value))
                        {
                            return;
                        }

                        // 去除首尾空格后写入
                        attrs[key.Trim()] = value.Trim();
                    }

                    // 局部函数，安全写入数值属性到JSON字典，统一使用英文小数点
                    void AddNumberAttr(string key, double value)
                    {
                        // 属性名为空时不写入
                        if (string.IsNullOrWhiteSpace(key))
                        {
                            return;
                        }

                        // 将数字转成InvariantCulture字符串，避免小数点格式混乱
                        attrs[key.Trim()] = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }

                    // 初始化文件主表信息
                    fs.FilePath = doc.Name;
                    fs.FileType = ".dwg";
                    fs.CreatedAt = DateTime.Now;
                    fs.UpdatedAt = DateTime.Now;
                    fs.IsActive = 1;
                    fs.IsPublic = 1;
                    fs.CreatedBy = Environment.UserName;
                    fs.Version = 1;
                    fs.Scale = 1.0;

                    // 初始化旧属性对象，保持兼容
                    fa.CreatedAt = DateTime.Now;
                    fa.UpdatedAt = DateTime.Now;

                    // 记录基础元数据到JSON
                    AddTextAttr("CreatedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    AddTextAttr("UpdatedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    AddTextAttr("SourceDrawing", doc.Name);

                    // 如果当前实体是块参照，则提取块相关信息
                    if (entity is BlockReference br)
                    {
                        // 读取块定义
                        var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                        // 块定义存在时，优先使用块名
                        if (btr != null)
                        {
                            fs.DisplayName = btr.Name;
                            fs.BlockName = btr.Name;
                        }
                        else
                        {
                            // 兜底处理，避免块定义读取失败时名称为空
                            fs.DisplayName = "BlockReference";
                            fs.BlockName = "BlockReference";
                        }

                        // 写入块的主信息
                        fs.LayerName = br.Layer;
                        fs.ColorIndex = br.Color.ColorIndex;
                        fs.Scale = br.ScaleFactors.X;

                        // 同步写入旧属性对象，兼容旧界面和旧逻辑
                        fa.Angle = (decimal)br.Rotation;
                        fa.BasePointX = (decimal)br.Position.X;
                        fa.BasePointY = (decimal)br.Position.Y;
                        fa.BasePointZ = (decimal)br.Position.Z;

                        // 写入JSON动态属性
                        AddNumberAttr("Angle", br.Rotation);
                        AddNumberAttr("BasePointX", br.Position.X);
                        AddNumberAttr("BasePointY", br.Position.Y);
                        AddNumberAttr("BasePointZ", br.Position.Z);

                        // 提取块属性集合
                        if (br.AttributeCollection != null && br.AttributeCollection.Count > 0)
                        {
                            // 用于拼接备注文本
                            var attributesText = new List<string>();

                            // 遍历块属性
                            foreach (ObjectId attId in br.AttributeCollection)
                            {
                                // 读取属性引用对象
                                var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;

                                // 空对象跳过
                                if (attRef == null)
                                {
                                    continue;
                                }

                                // 把块属性标签和值写入JSON
                                AddTextAttr(attRef.Tag, attRef.TextString ?? string.Empty);

                                // 同时汇总成备注，兼容旧显示方式
                                attributesText.Add($"{attRef.Tag}: {attRef.TextString}");
                            }

                            // 如果有属性文本，则写入旧属性对象和JSON备注
                            if (attributesText.Count > 0)
                            {
                                fa.Remarks = string.Join("\n", attributesText);
                                AddTextAttr("Remarks", fa.Remarks);
                            }
                        }
                    }
                    else
                    {
                        // 普通实体的名称使用RXClass名称
                        fs.DisplayName = entity.GetRXClass().Name;
                        fs.LayerName = entity.Layer;
                        fs.ColorIndex = entity.Color.ColorIndex;
                    }

                    // 默认文件名使用显示名称
                    fs.FileName = fs.DisplayName;
                    fa.FileName = fs.DisplayName;

                    // 把主表关键字段写入JSON，方便测试核对
                    AddTextAttr("FileName", fs.FileName ?? string.Empty);
                    AddTextAttr("DisplayName", fs.DisplayName ?? string.Empty);
                    AddTextAttr("BlockName", fs.BlockName ?? string.Empty);
                    AddTextAttr("LayerName", fs.LayerName ?? string.Empty);
                    AddTextAttr("ColorIndex", fs.ColorIndex.HasValue ? fs.ColorIndex.Value.ToString() : string.Empty);
                    AddTextAttr("Scale", fs.Scale.HasValue ? fs.Scale.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : string.Empty);

                    // 尝试估算实体尺寸
                    try
                    {
                        // 读取几何边界
                        var ext = entity.GeometricExtents;

                        // 计算长宽高
                        double length = ext.MaxPoint.X - ext.MinPoint.X;
                        double width = ext.MaxPoint.Y - ext.MinPoint.Y;
                        double height = ext.MaxPoint.Z - ext.MinPoint.Z;

                        // 同步写入旧属性对象
                        fa.Length = (decimal)length;
                        fa.Width = (decimal)width;
                        fa.Height = (decimal)height;

                        // 写入JSON动态属性
                        AddNumberAttr("Length", length);
                        AddNumberAttr("Width", width);
                        AddNumberAttr("Height", height);
                    }
                    catch (Exception ex)
                    {
                        // 尺寸估算失败仅记日志，不中断流程
                        LogManager.Instance.LogWarning($"估算尺寸失败: {ex.Message}");
                    }

                    // 准备预览图输出路径
                    string previewPath = PreparePreviewPath();

                    try
                    {
                        // 保存当前视图，方便截图后恢复
                        ViewTableRecord originalView = ed.GetCurrentView();

                        // 先尝试获取实体世界坐标包围盒
                        Extents3d? worldExt = null;
                        try
                        {
                            if (TryGetEntityWorldExtents(tr, entity, out var extWorld))
                            {
                                worldExt = extWorld;
                            }
                        }
                        catch (Exception exExt)
                        {
                            LogManager.Instance.LogWarning($"TryGetEntityWorldExtents 失败: {exExt.Message}");
                            worldExt = null;
                        }

                        // 截图成功标记
                        bool captured = false;

                        // 优先按实体包围盒截图
                        if (worldExt != null)
                        {
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

                        // 如果包围盒截图失败，则尝试临时插入方式回退截图
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

                        // 如果还失败，则退回到全图截图
                        if (!captured)
                        {
                            Bitmap previewBmp = null;

                            // 最多尝试3次获取截图
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

                            // 截图成功则保存为PNG
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

                        // 恢复原始视图
                        try
                        {
                            ed.SetCurrentView(originalView);
                        }
                        catch (Exception ex)
                        {
                            LogManager.Instance.LogWarning($"恢复原视图失败: {ex.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        // 预览图失败不影响主流程
                        LogManager.Instance.LogWarning($"生成预览图异常: {ex.Message}");
                    }

                    // 提交事务
                    tr.Commit();
                }
            }

            // 返回最终DTO
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
        /// 辅助：尝试获取实体在世界坐标下的包围盒（支持普通实体与 BlockReference）
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
                // 先尝试使用实体自身的 GeometricExtents，这对普通实体最直接有效
                try
                {
                    var eExt = ent.GeometricExtents;
                    if (IsExtentsReasonable(eExt))
                    {
                        worldExt = eExt;
                        return true;
                    }
                    else
                    {
                        LogManager.Instance.LogWarning($"实体自身 GeometricExtents 非常规：{eExt.MinPoint} - {eExt.MaxPoint}，将尝试其它方案。");
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogWarning($"常规获取几何边界失败: {ex.Message}");
                }

                // 如果是块参照，则进一步尝试基于块定义递归计算包围盒
                if (ent is BlockReference br)
                {
                    if (TryComputeBlockReferenceWorldExtents(tr, br, out var extFromBtr) && IsExtentsReasonable(extFromBtr))
                    {
                        worldExt = extFromBtr;
                        return true;
                    }

                    LogManager.Instance.LogWarning("基于块定义计算的 worldExt 非常规或计算失败。");
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
            ViewTableRecord originalView = null;

            try
            {
                // 先检查包围盒是否合理，不合理则直接放弃
                if (!IsExtentsReasonable(ext))
                {
                    LogManager.Instance.LogWarning($"TryCaptureUsingView: 提供的 ext 不合理，放弃按 ext 截图: Min{ext.MinPoint} Max{ext.MaxPoint}");
                    return false;
                }

                // 计算包围盒尺寸
                var min = ext.MinPoint;
                var max = ext.MaxPoint;

                double contentWidthWorld = Math.Max(max.X - min.X, 1.0);
                double contentHeightWorld = Math.Max(max.Y - min.Y, 1.0);

                // 计算适度边距，避免截图太贴边
                double worldMargin = Math.Min(10.0, Math.Max(Math.Max(contentWidthWorld, contentHeightWorld) * 0.05, 0.5));

                // 最终视图宽高 = 内容尺寸 + 边距
                double contentWidth = contentWidthWorld + worldMargin * 2;
                double contentHeight = contentHeightWorld + worldMargin * 2;

                // 计算视图中心
                var centerWCS = new Point3d((min.X + max.X) / 2.0, (min.Y + max.Y) / 2.0, 0);

                // 备份当前视图
                originalView = ed.GetCurrentView();
                var newView = (ViewTableRecord)originalView.Clone();

                // 设置顶视图和中心点
                newView.ViewDirection = Vector3d.ZAxis;
                newView.ViewTwist = 0.0;
                newView.CenterPoint = new Point2d(centerWCS.X, centerWCS.Y);
                newView.Width = contentWidth;
                newView.Height = contentHeight;

                // 应用新视图并等待屏幕刷新
                ed.SetCurrentView(newView);
                System.Threading.Thread.Sleep(300);
                try { Application.UpdateScreen(); } catch { }

                const int baseSize = 600;
                uint outputWidth;
                uint outputHeight;
                double aspectRatio = contentWidth / contentHeight;

                if (aspectRatio >= 1.0)
                {
                    outputWidth = (uint)baseSize;
                    outputHeight = (uint)(baseSize / aspectRatio);
                }
                else
                {
                    outputHeight = (uint)baseSize;
                    outputWidth = (uint)(baseSize * aspectRatio);
                }

                outputWidth = Math.Max(outputWidth, 300u);
                outputHeight = Math.Max(outputHeight, 300u);

                // 执行截图
                Bitmap previewBmp = CaptureImage(doc, outputWidth, outputHeight);
                if (previewBmp == null)
                {
                    LogManager.Instance.LogWarning($"截图失败: 中心WCS({centerWCS.X:F2},{centerWCS.Y:F2}), 视图尺寸({contentWidth:F2}×{contentHeight:F2})");
                    return false;
                }

                try
                {
                    // 裁剪后保存PNG
                    using (var cropped = CropPreviewToContent(previewBmp, worldMargin, contentWidthWorld, contentHeightWorld, outputWidth, outputHeight))
                    {
                        cropped.Save(outPreviewPath, System.Drawing.Imaging.ImageFormat.Png);
                    }
                }
                finally
                {
                    previewBmp.Dispose();
                }

                // 最终检查文件是否真实生成成功
                if (!System.IO.File.Exists(outPreviewPath))
                {
                    LogManager.Instance.LogWarning($"截图流程结束但文件未生成: {outPreviewPath}");
                    return false;
                }

                var fileInfo = new System.IO.FileInfo(outPreviewPath);
                if (fileInfo.Length <= 0)
                {
                    LogManager.Instance.LogWarning($"截图流程结束但文件为空: {outPreviewPath}");
                    return false;
                }

                LogManager.Instance.LogInfo($"截图成功: 中心WCS({centerWCS.X:F2},{centerWCS.Y:F2}), 视图尺寸({contentWidth:F2}×{contentHeight:F2}), 图像尺寸({outputWidth}×{outputHeight}), worldMargin({worldMargin:F2})");
                return true;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"TryCaptureUsingView 异常: {ex.Message}");
                return false;
            }
            finally
            {
                // 无论成功失败都恢复原视图，避免影响后续截图和界面状态
                if (originalView != null)
                {
                    try
                    {
                        ed.SetCurrentView(originalView);
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogWarning($"恢复原视图失败: {ex.Message}");
                    }
                }
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
