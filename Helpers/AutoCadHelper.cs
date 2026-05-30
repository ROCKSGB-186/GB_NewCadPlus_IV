using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// AutoCAD操作工具类
    /// </summary>
    public static class AutoCadHelper
    {
        /// <summary>
        /// 安全文件名
        /// </summary>
        /// <param name="name">要处理的文件名</param>
        /// <returns></returns>
        private static string SanitizeFileName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
        
        /// <summary>
        /// 从外部 DWG 导入指定块定义到当前文档并在目标点插入一个 BlockReference（包含属性）
        /// 返回插入的 BlockReference 的 ObjectId，失败返回 ObjectId.Null。
        /// （保留原有实现，已做健壮性和注释增强）
        /// </summary>
        public static ObjectId InsertBlockFromExternalDwg(string dwgPath, string blockName, Point3d insertPoint)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return ObjectId.Null;

            // 保证在 AutoCAD 主线程并加文档锁
            using (doc.LockDocument())
            {
                try
                {
                    // 读取外部 DWG 到临时 Database
                    using (var sourceDb = new Database(false, true))
                    {
                        sourceDb.ReadDwgFile(dwgPath, System.IO.FileShare.Read, true, null);

                        using (var sourceTr = sourceDb.TransactionManager.StartTransaction())
                        {
                            var sourceBt = (BlockTable)sourceTr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
                            if (!sourceBt.Has(blockName))
                                return ObjectId.Null;

                            ObjectId sourceBtrId = sourceBt[blockName];

                            // 克隆到当前文档数据库（一次性把块定义导入目标 DB）
                            IdMapping mapping = new IdMapping();
                            sourceDb.WblockCloneObjects(new ObjectIdCollection { sourceBtrId },
                                                        doc.Database.BlockTableId,
                                                        mapping,
                                                        DuplicateRecordCloning.Replace,
                                                        false);

                            sourceTr.Commit();

                            if (!mapping.Contains(sourceBtrId))
                                return ObjectId.Null;

                            ObjectId newBtrId = mapping[sourceBtrId].Value;

                            // 在当前文档开启事务并插入 BlockReference（并正确处理属性）
                            using (var tr = doc.Database.TransactionManager.StartTransaction())
                            {
                                // 获取目标模型空间（写模式）
                                var ms = (BlockTableRecord)tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);
                                // 创建 BlockReference 引用新块定义
                                var blockRef = new BlockReference(insertPoint, newBtrId);
                                // 把 BlockReference 加入模型空间并注册
                                ms.AppendEntity(blockRef);
                                tr.AddNewlyCreatedDBObject(blockRef, true);
                                // 读取目标数据库中新克隆的块表记录（以只读方式）
                                var btr = (BlockTableRecord)tr.GetObject(newBtrId, OpenMode.ForRead);
                                // 如果块定义包含属性定义，逐一创建 AttributeReference 并追加到 blockRef
                                if (btr.HasAttributeDefinitions)
                                {
                                    foreach (ObjectId id in btr)
                                    {
                                        var dbObj = tr.GetObject(id, OpenMode.ForRead);
                                        if (dbObj is AttributeDefinition attDef && !attDef.Constant)
                                        {
                                            // 创建属性引用，并从定义设置默认值（相对于块）
                                            var attRef = new AttributeReference();
                                            attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                                            // 必须在把属性附加到 BlockReference 后调用 AddNewlyCreatedDBObject
                                            blockRef.AttributeCollection.AppendAttribute(attRef);
                                            tr.AddNewlyCreatedDBObject(attRef, true);
                                        }
                                    }
                                }
                                tr.Commit();
                                return blockRef.ObjectId;
                            }
                        }
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n插入块失败: {ex.Message}");
                    return ObjectId.Null;
                }
                catch (Exception ex)
                {
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n未知错误: {ex.Message}");
                    return ObjectId.Null;
                }
            }
        }

        /// <summary>
        /// 在活动文档中执行事务
        /// </summary>
        /// <param name="action">要执行的操作</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void ExecuteInDocumentTransaction(Action<Document, Transaction> action)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) throw new InvalidOperationException("当前没有活动文档。");
            using (doc.LockDocument())
            {
                var db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    action(doc, tr);
                    tr.Commit();
                }
            }
        }
        
        /// <summary>
        /// 缓存锁
        /// </summary>
        private static readonly object _cacheLock = new object();

        /// <summary>
        /// 缓存比例
        /// </summary>
        private static double _cachedScale = double.NaN;

        /// <summary>
        /// 缓存时间
        /// </summary>
        private static DateTime _cacheTime = DateTime.MinValue;

        /// <summary>
        /// 缓存时间跨度
        /// </summary>
        private static readonly TimeSpan _cacheTtl = TimeSpan.FromSeconds(2);

        /// <summary>
        /// 获取当前绘图比例（优先使用用户在WPF界面输入的值）
        /// </summary>
        /// <param name="useCache">是否使用缓存</param>
        /// <returns>绘图比例</returns>
        public static double GetScale(bool useCache = true)
        {
            // 默认先给一个无效占位值，后续按界面来源覆盖
            double userScale = 0.0;

            // 如果当前是 WinForm 状态，则优先读取 WinForm 比例缓存
            if (VariableDictionary.winForm_Status)
            {
                // WinForm 模式下直接取 textBoxScale
                userScale = VariableDictionary.textBoxScale;
            }
            else
            {
                // WPF 模式优先使用全局缓存的 wpfTextBoxScale（你要求的优先级）
                if (VariableDictionary.wpfTextBoxScale > 0.0)
                {
                    userScale = VariableDictionary.wpfTextBoxScale;
                }
                else
                {
                    // 若 wpfTextBoxScale 无效，再尝试实时从 WPF 文本框读取
                    userScale = GetDrawingScaleFromWpf();
                }

                // 兜底再尝试 textBoxScale，避免某些旧流程仅写入 textBoxScale
                if (userScale <= 0.0 && VariableDictionary.textBoxScale > 0.0)
                {
                    userScale = VariableDictionary.textBoxScale;
                }
            }

            // 只要拿到有效用户比例，直接返回，不走CAD视口计算
            if (userScale > 0.0)
            {
                return userScale;
            }

            // 如果界面比例不可用，则走原有缓存逻辑
            if (useCache)
            {
                lock (_cacheLock)
                {
                    if (!double.IsNaN(_cachedScale) && (DateTime.UtcNow - _cacheTime) < _cacheTtl)
                        return _cachedScale;
                }
            }

            // 计算当前图纸/视口比例作为最终回退
            double scale = ComputeActiveDrawingScale();

            // 写入缓存，减少频繁计算
            lock (_cacheLock)
            {
                _cachedScale = scale;
                _cacheTime = DateTime.UtcNow;
            }

            return scale;
        }

        /// <summary>
        /// 获取当前视图比例的方法
        /// </summary>
        /// <returns></returns>
        private static double ComputeActiveDrawingScale()
        {
            const double defaultScale = 1.0;

            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return defaultScale;

                using (doc.LockDocument())
                {
                    var db = doc.Database;
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            // 若在模型空间，返回 1.0
                            try { if (db.TileMode) return 1.0; } catch { }

                            // 尝试获取当前视图
                            Autodesk.AutoCAD.DatabaseServices.ViewTableRecord currentView = null;
                            try { currentView = doc.Editor.GetCurrentView(); } catch { currentView = null; }

                            // 遍历布局里实体，找 Viewport（使用反射以兼容不同 API）
                            var lm = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
                            string layoutName = null;
                            try { layoutName = lm.CurrentLayout; } catch { layoutName = null; }

                            double bestScore = double.MaxValue;
                            double candidateScale = double.NaN;
                            bool found = false;

                            if (!string.IsNullOrEmpty(layoutName))
                            {
                                try
                                {
                                    ObjectId layoutId = lm.GetLayoutId(layoutName);
                                    var layout = (Autodesk.AutoCAD.DatabaseServices.Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                                    if (layout != null)
                                    {
                                        var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                                        foreach (ObjectId entId in btr)
                                        {
                                            try
                                            {
                                                var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                                                if (ent == null) continue;

                                                var etype = ent.GetType();
                                                if (!string.Equals(etype.Name, "Viewport", StringComparison.OrdinalIgnoreCase))
                                                    continue;

                                                double? customScaleRaw = null;
                                                double? viewHeight = null;
                                                object viewCenterObj = null;
                                                object centerPointObj = null;

                                                try
                                                {
                                                    var p = etype.GetProperty("CustomScale");
                                                    if (p != null) { var v = p.GetValue(ent); if (v != null) customScaleRaw = Convert.ToDouble(v); }
                                                }
                                                catch { }

                                                try
                                                {
                                                    var p = etype.GetProperty("ViewHeight");
                                                    if (p != null) { var v = p.GetValue(ent); if (v != null) viewHeight = Convert.ToDouble(v); }
                                                }
                                                catch { }

                                                try { var p = etype.GetProperty("ViewCenter"); if (p != null) viewCenterObj = p.GetValue(ent); } catch { }
                                                try { var p = etype.GetProperty("CenterPoint"); if (p != null) centerPointObj = p.GetValue(ent); } catch { }

                                                double score = 0.0;
                                                if (currentView != null)
                                                {
                                                    try
                                                    {
                                                        double vx = double.NaN, vy = double.NaN;
                                                        if (viewCenterObj != null)
                                                        {
                                                            var tc = viewCenterObj.GetType();
                                                            var px = tc.GetProperty("X")?.GetValue(viewCenterObj);
                                                            var py = tc.GetProperty("Y")?.GetValue(viewCenterObj);
                                                            vx = Convert.ToDouble(px);
                                                            vy = Convert.ToDouble(py);
                                                        }
                                                        else if (centerPointObj != null)
                                                        {
                                                            var tc = centerPointObj.GetType();
                                                            var px = tc.GetProperty("X")?.GetValue(centerPointObj);
                                                            var py = tc.GetProperty("Y")?.GetValue(centerPointObj);
                                                            vx = Convert.ToDouble(px);
                                                            vy = Convert.ToDouble(py);
                                                        }
                                                        else
                                                        {
                                                            score = 1e6;
                                                        }

                                                        if (!double.IsNaN(vx) && !double.IsNaN(vy))
                                                        {
                                                            var cur = currentView.CenterPoint;
                                                            score = Math.Abs(vx - cur.X) + Math.Abs(vy - cur.Y);
                                                        }
                                                    }
                                                    catch { score = 1e6; }
                                                }
                                                else
                                                {
                                                    score = 1e5;
                                                }

                                                if (customScaleRaw.HasValue && customScaleRaw.Value > 0.0)
                                                {
                                                    double normalized = customScaleRaw.Value >= 1.0 ? 1.0 / customScaleRaw.Value : customScaleRaw.Value;
                                                    if (score < bestScore)
                                                    {
                                                        bestScore = score;
                                                        candidateScale = normalized;
                                                        found = true;
                                                        if (score <= 1e-6) break;
                                                    }
                                                }
                                                else if (viewHeight.HasValue && viewHeight.Value > 1e-12 && currentView != null)
                                                {
                                                    try
                                                    {
                                                        double normalized = currentView.Height / viewHeight.Value;
                                                        if (score < bestScore)
                                                        {
                                                            bestScore = score;
                                                            candidateScale = normalized;
                                                            found = true;
                                                        }
                                                    }
                                                    catch { }
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                                catch { }
                            }

                            if (found && !double.IsNaN(candidateScale) && candidateScale > 0.0)
                            {
                                return candidateScale;
                            }

                            return default(double);
                        }
                        catch { return default(double); }
                    }
                }
            }
            catch { return default(double); }
        }
        
        /// <summary>
        /// 从WPF界面获取用户输入的绘图比例
        /// </summary>
        /// <returns>用户输入的比例值，如果获取失败返回0</returns>
        private static double GetDrawingScaleFromWpf()
        {
            // 局部函数，统一解析字符串到正数比例
            static double ParsePositiveScale(string raw)
            {
                // 空字符串直接返回0，表示无有效输入
                if (string.IsNullOrWhiteSpace(raw)) return 0.0;
                // 先去掉首尾空白
                raw = raw.Trim();

                // 先按 InvariantCulture 解析（支持标准小数点）
                if (double.TryParse(raw, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out double v1) && v1 > 0.0)
                    return v1;

                // 兼容中文环境下用逗号作小数分隔符
                string alt = raw.Replace(',', '.');
                if (double.TryParse(alt, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out double v2) && v2 > 0.0)
                    return v2;

                // 最后兜底用当前区域解析
                if (double.TryParse(raw, out double v3) && v3 > 0.0)
                    return v3;

                // 解析失败返回0
                return 1;
            }

            try
            {
                // 优先拿到WPF主界面实例
                var inst = GB_NewCadPlus_IV.WpfMainWindow.Instance;
                if (inst != null)
                {
                    // 用于承接从UI线程读取到的文本
                    string textFromUi = string.Empty;
                    // 用于承接Tag默认值（例如XAML里 Tag="100"）
                    string tagFromUi = string.Empty;

                    // 必须在WPF Dispatcher线程访问TextBox，避免跨线程异常
                    if (inst.Dispatcher != null)
                    {
                        // 若当前就在UI线程，直接读取
                        if (inst.Dispatcher.CheckAccess())
                        {
                            // 读取TextBox文本
                            textFromUi = inst.TextBox绘图比例?.Text ?? string.Empty;
                            // 读取Tag作为默认比例兜底
                            tagFromUi = inst.TextBox绘图比例?.Tag?.ToString() ?? string.Empty;
                        }
                        else
                        {
                            // 不在UI线程时切回UI线程读取，避免抛跨线程异常
                            inst.Dispatcher.Invoke(() =>
                            {
                                // 读取Text
                                textFromUi = inst.TextBox绘图比例?.Text ?? string.Empty;
                                // 读取Tag
                                tagFromUi = inst.TextBox绘图比例?.Tag?.ToString() ?? string.Empty;
                            });
                        }
                    }

                    // 优先解析用户输入的Text
                    double v = ParsePositiveScale(textFromUi);
                    if (v > 0.0) return v;

                    // Text无效时尝试Tag默认值（你当前XAML里是100）
                    v = ParsePositiveScale(tagFromUi);
                    if (v > 0.0) return v;
                }
            }
            catch
            {
                // WPF读取失败时继续走变量兜底
            }

            // 兜底1，读取WPF侧缓存值（由WPF代码维护）
            if (VariableDictionary.wpfTextBoxScale > 0.0)
                return VariableDictionary.wpfTextBoxScale;

            // 兜底2，读取通用缓存值
            if (VariableDictionary.textBoxScale > 0.0)
                return VariableDictionary.textBoxScale;

            // 最终失败返回0，让上层走原有回退逻辑
            return 1;
        }

        /// <summary>
        /// 获取WPF主窗口实例
        /// </summary>
        /// <returns>WPF主窗口实例</returns>
        public static object GetWpfWindow()
        {
            try
            {
                // 优先返回静态实例（若已初始化）
                var inst = GB_NewCadPlus_IV.WpfMainWindow.Instance;
                if (inst != null) return inst;

                // 兜底：尝试遍历 Application.Windows 查找包含 WpfMainWindow 的 Window 并返回其 Content
                var app = System.Windows.Application.Current;
                if (app != null)
                {
                    foreach (System.Windows.Window w in app.Windows)
                    {
                        try
                        {
                            // 若 Window 的 Content 或视觉树中包含 WpfMainWindow，返回它
                            if (w.Content is GB_NewCadPlus_IV.WpfMainWindow wc) return wc;

                            // 遍历视觉树查找 UserControl
                            var found = FindChildInVisualTree<GB_NewCadPlus_IV.WpfMainWindow>(w);
                            if (found != null) return found;
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 辅助：在视觉树中查找指定类型的子元素（递归）
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="parent"></param>
        /// <returns></returns>
        private static T FindChildInVisualTree<T>(System.Windows.DependencyObject parent) where T : System.Windows.DependencyObject
        {
            if (parent == null) return null;
            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T typed) return typed;
                var res = FindChildInVisualTree<T>(child);
                if (res != null) return res;
            }
            return null;
        }              

        /// <summary>
        /// 安全的日志记录方法，防止并发访问问题
        /// </summary>
        /// <param name="message">日志消息</param>
        public static void LogWithSafety(string message)
        {
            try
            {
                LogManager.Instance.LogInfo(message);
            }
            catch (System.Exception ex)
            {
                // 如果日志记录失败，至少在命令行显示
                Env.Editor.WriteMessage($"\n日志记录失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 安全读取系统变量并返回 double，异常或不存在时返回 defaultValue（中文注释）
        /// </summary>
        public static double SafeGetSystemVariableDouble(string varName, double defaultValue = 1.0)
        {
            try
            {
                var v = Application.GetSystemVariable(varName);
                if (v == null) return defaultValue;
                return Convert.ToDouble(v);
            }
            catch
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// 计算全局线型缩放因子：LTSCALE * CELTSCALE * PSLTSCALE（容错）
        /// 这个因子用于把“期望的图上断线段长度”映射为实体的 LinetypeScale
        /// </summary>
        public static double ComputeGlobalLinetypeScaleFactor()
        {
            double ltscale = SafeGetSystemVariableDouble("LTSCALE", 1.0);
            double celtscale = SafeGetSystemVariableDouble("CELTSCALE", 1.0);
            double psltscale = SafeGetSystemVariableDouble("PSLTSCALE", 1.0);
            double factor = ltscale * celtscale * psltscale;
            if (double.IsNaN(factor) || factor <= 0) factor = 1.0;
            return factor;
        }

        /// <summary>
        /// 通用：安全提示用户输入一个正的 double（带默认值、回车使用默认）
        /// 说明：避免直接在每个命令中使用不存在的 LowerLimit 属性，统一校验逻辑放在这里。
        /// 返回：始终返回一个有效的正数（<=0 会回退为 defaultValue）
        /// </summary>
        public static double PromptForPositiveDouble(string message, double defaultValue = 100.0)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return defaultValue;
                var ed = doc.Editor;
                var pdo = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions(message)
                {
                    DefaultValue = defaultValue,
                    AllowNone = true
                };
                var pdr = ed.GetDouble(pdo);
                if (pdr.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.None)
                {
                    // 用户直接回车，使用默认值
                    return defaultValue;
                }
                if (pdr.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK && pdr.Value > 0.0)
                {
                    return pdr.Value;
                }
                // 非 OK 或者非法值，回退并提示（但不抛出）
                ed.WriteMessage($"\n输入无效，已使用默认值 {defaultValue}。");
                return defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }

    
    }
}
