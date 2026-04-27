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
        /// 把二进制 DWG 写入到临时文件并返回路径，调用者负责删除（如果需要）
        /// 中文注释：用于把从服务器下载到内存的 DWG（二进制）写入磁盘，供 AutoCAD API 读取。
        /// </summary>
        public static string SaveBytesToTempDwg(byte[] dwgBytes, string? hintName = null)
        {
            // 使用用户临时目录，避免与程序资源目录冲突
            string tempDir = Path.Combine(Path.GetTempPath(), "GB_CADTools", "TempDwg");
            Directory.CreateDirectory(tempDir);
            // 安全文件名
            string fileName = string.IsNullOrWhiteSpace(hintName) ? $"tmp_{DateTime.Now:yyyyMMddHHmmssfff}.dwg" : $"{SanitizeFileName(hintName)}_{DateTime.Now:yyyyMMddHHmmssfff}.dwg";
            string fullPath = Path.Combine(tempDir, fileName);
            File.WriteAllBytes(fullPath, dwgBytes);
            return fullPath;
        }

        /// <summary>
        /// 清理 AutoCadHelper 产生的临时 DWG 文件
        /// olderThanMinutes > 0 时仅清理早于该分钟数的文件；否则清理全部
        /// </summary>
        public static void CleanupTempDwgFiles(int olderThanMinutes = 0)
        {
            try
            {
                string tempDir = Path.Combine(Path.GetTempPath(), "GB_CADTools", "TempDwg");
                if (!Directory.Exists(tempDir)) return;

                DateTime now = DateTime.Now;
                foreach (var file in Directory.GetFiles(tempDir, "*.dwg", SearchOption.TopDirectoryOnly))
                {
                    try
                    {
                        if (olderThanMinutes > 0)
                        {
                            var fi = new FileInfo(file);
                            if ((now - fi.LastWriteTime).TotalMinutes < olderThanMinutes)
                                continue;
                        }

                        File.Delete(file);
                    }
                    catch
                    {
                        // 单文件清理失败忽略，继续清理其他文件
                    }
                }
            }
            catch
            {
                // 目录级异常忽略
            }
        }

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
        /// 从字节数组导入块定义到当前数据库（通过临时文件实现）。
        /// 返回导入后的块定义 ObjectId 或 ObjectId.Null（失败）。
        /// 注：使用时需要在外层事务/DBTrans 中调用以避免重复导入/线程问题。
        /// </summary>
        public static ObjectId ImportBlockDefinitionFromBytes(byte[] dwgBytes, string blockName, DBTrans dbTr)
        {
            // 写临时文件
            string tmpFile = SaveBytesToTempDwg(dwgBytes, blockName);
            try
            {
                return ImportBlockDefinitionToCurrentDatabase(tmpFile, blockName, dbTr);
            }
            finally
            {
                // 尝试删除临时文件（失败不抛）
                try { if (File.Exists(tmpFile)) File.Delete(tmpFile); } catch { }
            }
        }

        /// <summary>
        /// 将 byte[] DWG 临时写盘后在当前文档中插入一个 BlockReference（包含属性），并返回插入的 ObjectId。
        /// 这个方法是对已有 InsertBlockFromExternalDwg 的补充：支持从内存直接插入。
        /// </summary>
        public static ObjectId InsertBlockFromExternalDwg(byte[] dwgBytes, string blockName, Point3d insertPoint)
        {
            string tmpFile = SaveBytesToTempDwg(dwgBytes, blockName);
            try
            {
                return InsertBlockFromExternalDwg(tmpFile, blockName, insertPoint);
            }
            finally
            {
                try { if (File.Exists(tmpFile)) File.Delete(tmpFile); } catch { }
            }
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
        /// 从源数据库中导入指定块
        /// </summary>
        /// <param name="dwgPath">DWG 文件路径</param>
        /// <param name="blockName">块名称</param>
        /// <param name="dbTr">目标数据库事务</param>
        /// <returns>目标数据库中新块记录的 ObjectId（失败返回 ObjectId.Null）</returns>
        public static ObjectId ImportBlockDefinitionToCurrentDatabase(string dwgPath, string blockName, DBTrans dbTr)
        {
            if (!File.Exists(dwgPath)) return ObjectId.Null;// 文件不存在 
            using (var sourceDb = new Database(false, true))// 创建源数据库 只读打开
            {
                sourceDb.ReadDwgFile(dwgPath, FileShare.ReadWrite, true, null);// 打开源数据库 读取 DWG 文件
                using (var sourceTr = sourceDb.TransactionManager.StartTransaction())// 开始事务 启动源数据库事务
                {
                    var sourceBt = (BlockTable)sourceTr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);// 获取块表 获取源数据库块表
                    if (!sourceBt.Has(blockName)) return ObjectId.Null;// 源数据库没有该块 块定义不存在
                    ObjectId sourceBtrId = sourceBt[blockName];// 获取块记录 获取源块记录 ID
                    // 直接从 sourceDb 克隆到 targetTr.Database（目标数据库）
                    var ids = new ObjectIdCollection { sourceBtrId };// 创建 ObjectId 集合 包含源块记录 ID
                    var mapping = new IdMapping();
                    sourceDb.WblockCloneObjects(ids, dbTr.Database.BlockTableId, mapping, DuplicateRecordCloning.Replace, false);// 用 WblockCloneObjects 克隆块记录 到目标数据库

                    sourceTr.Commit();// 提交源数据库事务
                    return mapping.Contains(sourceBtrId) ? mapping[sourceBtrId].Value : ObjectId.Null;// 返回目标数据库中新块记录的 ObjectId
                }
            }
        }

        /// <summary>
        /// 把一个来自其它数据库的实体（已从源读取）以安全方式复制到当前事务所在数据库并返回新实体的 ObjectId。
        /// 适用于单个实体：先在源上用 WblockCloneObjects 克隆其所属块记录，或直接创建 GetTransformedCopy 然后在 targetTr 中 Add。
        /// </summary>
        public static ObjectId CloneEntityIntoCurrentDatabase(Entity sourceEntity, DBTrans dbTr)
        {
            // 更稳妥的方式是使用 GetTransformedCopy（会返回一个新的实体实例）
            var clone = sourceEntity.GetTransformedCopy(Matrix3d.Identity);// 获取实体的变换副本（无变换）
            var id = dbTr.CurrentSpace.AddEntity(clone);// 在当前空间中添加实体 将克隆实体添加到当前空间 并获取新实体的 ObjectId
            return id;
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
        /// 在活动文档中执行事务
        /// </summary>
        /// <typeparam name="T"> </typeparam>
        /// <param name="func"> </param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public static T ExecuteInDocumentTransaction<T>(Func<Document, Transaction, T> func)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) throw new InvalidOperationException("当前没有活动文档。");
            using (doc.LockDocument())
            {
                var db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    T result = func(doc, tr);
                    tr.Commit();
                    return result;
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
            // 中文注释：默认先给一个无效占位值，后续按界面来源覆盖
            double userScale = 0.0;

            // 中文注释：如果当前是 WinForm 状态，则优先读取 WinForm 比例缓存
            if (VariableDictionary.winForm_Status)
            {
                // 中文注释：WinForm 模式下直接取 textBoxScale
                userScale = VariableDictionary.textBoxScale;
            }
            else
            {
                // 中文注释：WPF 模式优先使用全局缓存的 wpfTextBoxScale（你要求的优先级）
                if (VariableDictionary.wpfTextBoxScale > 0.0)
                {
                    userScale = VariableDictionary.wpfTextBoxScale;
                }
                else
                {
                    // 中文注释：若 wpfTextBoxScale 无效，再尝试实时从 WPF 文本框读取
                    userScale = GetDrawingScaleFromWpf();
                }

                // 中文注释：兜底再尝试 textBoxScale，避免某些旧流程仅写入 textBoxScale
                if (userScale <= 0.0 && VariableDictionary.textBoxScale > 0.0)
                {
                    userScale = VariableDictionary.textBoxScale;
                }
            }

            // 中文注释：只要拿到有效用户比例，直接返回，不走CAD视口计算
            if (userScale > 0.0)
            {
                return userScale;
            }

            // 中文注释：如果界面比例不可用，则走原有缓存逻辑
            if (useCache)
            {
                lock (_cacheLock)
                {
                    if (!double.IsNaN(_cachedScale) && (DateTime.UtcNow - _cacheTime) < _cacheTtl)
                        return _cachedScale;
                }
            }

            // 中文注释：计算当前图纸/视口比例作为最终回退
            double scale = ComputeActiveDrawingScale();

            // 中文注释：写入缓存，减少频繁计算
            lock (_cacheLock)
            {
                _cachedScale = scale;
                _cacheTime = DateTime.UtcNow;
            }

            return scale;
        }

        /// <summary>
        /// 清除比例缓存
        /// </summary>
        public static void Invalidate()
        {
            lock (_cacheLock)
            {
                _cachedScale = double.NaN;
                _cacheTime = DateTime.MinValue;
            }
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
        /// 计算当前绘图比例并写入 VariableDictionary.blockScale（调用 DrawingScaleService 实现具体检测）
        /// </summary>
        public static double GetAndApplyActiveDrawingScale()
        {
            double scale = GetScale(false);
            //VariableDictionary.blockScale = scale;
            VariableDictionary.wpfTextBoxScale = scale;
            return scale;
        }
        /// <summary>
        /// 新增：基于所选 ObjectId 列表推断比例分母（例如 1 -> 1:1, 100 -> 1:100）
        /// 优化：不仅检查Viewport，还检查图元本身的缩放信息
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="selIds">所选对象的 ObjectId 列表</param>
        /// <param name="roundToCommon">是否四舍五入到常见比例</param>
        /// <returns>推断出的比例分母</returns>
        public static double GetScaleDenominatorForSelection(Database db, ObjectId[] selIds, bool roundToCommon = false)
        {
            try
            {
                // 无选择项时回退到当前活动视口/数据库检测
                if (selIds == null || selIds.Length == 0 || db == null)
                    return TextFontsStyleHelper.DetermineScaleDenominator(GetScale(true), null, roundToCommon);

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // 优先检查选中的Viewport，因为Viewport直接包含了比例信息
                    foreach (var id in selIds)
                    {
                        try
                        {
                            var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (ent == null) continue;

                            // 如果选中的是Viewport，直接获取其比例
                            if (string.Equals(ent.GetType().Name, "Viewport", StringComparison.OrdinalIgnoreCase))
                            {
                                var pi = ent.GetType().GetProperty("CustomScale");
                                if (pi != null)
                                {
                                    var raw = pi.GetValue(ent);
                                    if (raw != null)
                                    {
                                        double customScale = Convert.ToDouble(raw);
                                        if (customScale > 0.0)
                                        {
                                            double denom = TextFontsStyleHelper.DetermineScaleDenominator(customScale, null, roundToCommon);
                                            return denom;
                                        }
                                    }
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    // 检查选中图元的比例信息
                    var scaleValues = new List<double>();

                    foreach (var id in selIds)
                    {
                        try
                        {
                            var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (ent == null) continue;

                            double entityScale = 1.0;

                            // 检查不同类型的图元比例信息
                            if (ent is BlockReference br)
                            {
                                // 块参照的比例 - 取X、Y、Z三个方向的比例的平均值
                                entityScale = (br.ScaleFactors.X + br.ScaleFactors.Y + br.ScaleFactors.Z) / 3.0;
                            }
                            else if (ent is DBText txt)
                            {
                                // 文字比例 - 使用文字高度作为比例参考
                                entityScale = txt.Height / 2.5; // 假设标准文字高度为2.5
                            }
                            else if (ent is MText mtxt)
                            {
                                // 多行文字比例 - 使用文字高度作为比例参考
                                entityScale = mtxt.Height / 2.5;
                            }
                            else if (ent is Line line)
                            {
                                // 直线 - 使用长度作为比例参考（需要与其他实体类型进行比较）
                                entityScale = line.Length / 1000.0; // 假设1000为标准长度
                            }
                            else if (ent is Polyline pl)
                            {
                                // 多段线 - 使用长度作为比例参考
                                entityScale = pl.Length / 1000.0;
                            }
                            else
                            {
                                // 对于其他类型的实体，使用几何尺寸估算比例
                                try
                                {
                                    var extents = ent.GeometricExtents;
                                    var width = Math.Abs(extents.MaxPoint.X - extents.MinPoint.X);
                                    var height = Math.Abs(extents.MaxPoint.Y - extents.MinPoint.Y);
                                    var avgSize = (width + height) / 2.0;
                                    // 假设正常图元在1:1比例下平均尺寸为1000
                                    entityScale = avgSize / 1000.0;
                                }
                                catch
                                {
                                    entityScale = 1.0; // 默认比例
                                }
                            }

                            if (entityScale > 0.0001) // 过滤掉极小的值
                            {
                                scaleValues.Add(entityScale);
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    // 如果找到了有效的比例值，返回平均值
                    if (scaleValues.Count > 0)
                    {
                        double avgScale = scaleValues.Average();
                        double denom = TextFontsStyleHelper.DetermineScaleDenominator(avgScale, null, roundToCommon);
                        return denom;
                    }

                    // 检查图元所属的布局中的Viewport
                    foreach (var id in selIds)
                    {
                        try
                        {
                            var obj = tr.GetObject(id, OpenMode.ForRead) as DBObject;
                            if (obj == null) continue;

                            // 获取图元所属的BlockTableRecord
                            ObjectId ownerId = obj.OwnerId;
                            if (ownerId == ObjectId.Null) continue;

                            var ownerBtr = tr.GetObject(ownerId, OpenMode.ForRead) as BlockTableRecord;
                            if (ownerBtr == null) continue;

                            // 如果所属的是布局（Layout），在该布局中查找Viewport
                            if (ownerBtr.IsLayout)
                            {
                                foreach (ObjectId entId in ownerBtr)
                                {
                                    try
                                    {
                                        var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                                        if (ent == null) continue;

                                        // 检查是否为Viewport
                                        if (string.Equals(ent.GetType().Name, "Viewport", StringComparison.OrdinalIgnoreCase))
                                        {
                                            var pi = ent.GetType().GetProperty("CustomScale");
                                            if (pi != null)
                                            {
                                                var raw = pi.GetValue(ent);
                                                if (raw != null)
                                                {
                                                    double customScale = Convert.ToDouble(raw);
                                                    if (customScale > 0.0)
                                                    {
                                                        double denom = TextFontsStyleHelper.DetermineScaleDenominator(customScale, null, roundToCommon);
                                                        return denom;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    tr.Commit();
                }
            }
            catch
            {
                // 忽略并回退
            }

            // 回退到全局/当前视口检测
            return TextFontsStyleHelper.DetermineScaleDenominator(GetScale(true), null, roundToCommon);
        }

        /// <summary>
        /// 从WPF界面获取用户输入的绘图比例
        /// </summary>
        /// <returns>用户输入的比例值，如果获取失败返回0</returns>
        private static double GetDrawingScaleFromWpf()
        {
            // 中文注释：局部函数，统一解析字符串到正数比例
            static double ParsePositiveScale(string raw)
            {
                // 中文注释：空字符串直接返回0，表示无有效输入
                if (string.IsNullOrWhiteSpace(raw)) return 0.0;
                // 中文注释：先去掉首尾空白
                raw = raw.Trim();

                // 中文注释：先按 InvariantCulture 解析（支持标准小数点）
                if (double.TryParse(raw, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out double v1) && v1 > 0.0)
                    return v1;

                // 中文注释：兼容中文环境下用逗号作小数分隔符
                string alt = raw.Replace(',', '.');
                if (double.TryParse(alt, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out double v2) && v2 > 0.0)
                    return v2;

                // 中文注释：最后兜底用当前区域解析
                if (double.TryParse(raw, out double v3) && v3 > 0.0)
                    return v3;

                // 中文注释：解析失败返回0
                return 1;
            }

            try
            {
                // 中文注释：优先拿到WPF主界面实例
                var inst = GB_NewCadPlus_IV.WpfMainWindow.Instance;
                if (inst != null)
                {
                    // 中文注释：用于承接从UI线程读取到的文本
                    string textFromUi = string.Empty;
                    // 中文注释：用于承接Tag默认值（例如XAML里 Tag="100"）
                    string tagFromUi = string.Empty;

                    // 中文注释：必须在WPF Dispatcher线程访问TextBox，避免跨线程异常
                    if (inst.Dispatcher != null)
                    {
                        // 中文注释：若当前就在UI线程，直接读取
                        if (inst.Dispatcher.CheckAccess())
                        {
                            // 中文注释：读取TextBox文本
                            textFromUi = inst.TextBox_绘图比例?.Text ?? string.Empty;
                            // 中文注释：读取Tag作为默认比例兜底
                            tagFromUi = inst.TextBox_绘图比例?.Tag?.ToString() ?? string.Empty;
                        }
                        else
                        {
                            // 中文注释：不在UI线程时切回UI线程读取，避免抛跨线程异常
                            inst.Dispatcher.Invoke(() =>
                            {
                                // 中文注释：读取Text
                                textFromUi = inst.TextBox_绘图比例?.Text ?? string.Empty;
                                // 中文注释：读取Tag
                                tagFromUi = inst.TextBox_绘图比例?.Tag?.ToString() ?? string.Empty;
                            });
                        }
                    }

                    // 中文注释：优先解析用户输入的Text
                    double v = ParsePositiveScale(textFromUi);
                    if (v > 0.0) return v;

                    // 中文注释：Text无效时尝试Tag默认值（你当前XAML里是100）
                    v = ParsePositiveScale(tagFromUi);
                    if (v > 0.0) return v;
                }
            }
            catch
            {
                // 中文注释：WPF读取失败时继续走变量兜底
            }

            // 中文注释：兜底1，读取WPF侧缓存值（由WPF代码维护）
            if (VariableDictionary.wpfTextBoxScale > 0.0)
                return VariableDictionary.wpfTextBoxScale;

            // 中文注释：兜底2，读取通用缓存值
            if (VariableDictionary.textBoxScale > 0.0)
                return VariableDictionary.textBoxScale;

            // 中文注释：最终失败返回0，让上层走原有回退逻辑
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

        // 新增：安全读取系统变量与安全提示正数输入等公共方法（中文注释）
        // 注意：把这些方法粘贴到 AutoCadHelper 类内部，靠近其它公用工具方法处。

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

        /// <summary>
        /// 确认 DimStyle 是否真的写入了当前图纸、是否有重名/不可见字符、以及 EnsureOrCreateDimStyle(DBTrans, string, short, double, ObjectId) 返回的 ObjectId 是否为回退值（即失败）
        /// </summary>
        [CommandMethod("DEBUG_DIMSTYLE")]
        public static void DEBUG_DIMSTYLE()
        {
            try
            {
                // 使用常用 DBTrans 开启事务（与你工程中其它命令一致）
                using var tr = new DBTrans();

                // 1) 列出当前 DimStyle 表中的所有样式名（调试用）
                var before = DimStyleHelper.ListDimStyleNames(tr) ?? new List<string>();
                string beforeList = before.Count == 0 ? "<empty>" : string.Join("\n", before);
                Env.Editor.WriteMessage($"\n[DEBUG] 当前 DimStyle 列表（创建前）:\n{beforeList}");

                // 2) 检查是否存在精确或规范化（去不可见/大小写忽略）的目标名
                string target = "JLPDI-定位";
                bool existsExact = before.Contains(target);
                bool existsNormalized = before.Any(n => string.Equals((n ?? string.Empty).Trim().Replace('\u00A0', ' '), target, StringComparison.OrdinalIgnoreCase));
                Env.Editor.WriteMessage($"\n[DEBUG] 精确存在: {existsExact}, 规范化存在(忽略大小写/NBSP): {existsNormalized}");

                // 3) 打印当前事务数据库与活动文档数据库是否一致（常见问题：事务写入了错误的数据库）
                var activeDb = Application.DocumentManager.MdiActiveDocument?.Database;
                var trDb = tr.Database;
                Env.Editor.WriteMessage($"\n[DEBUG] ActiveDocument.Database == tr.Database ? {ReferenceEquals(activeDb, trDb)}");

                // 4) 尝试调用 EnsureOrCreateDimStyle 创建/更新目标样式，并打印返回的 ObjectId
                //    注意：这里使用一个合理的默认参数（颜色索引 3、文字样式 tJText、scale 从 AutoCadHelper 读取）
                double uiScale = 1.0;
                try { uiScale = AutoCadHelper.GetScale(true); } catch { uiScale = 1.0; }
                ObjectId textStyleId = ObjectId.Null;
                try { textStyleId = tr.TextStyleTable["tJText"]; } catch { textStyleId = tr.Database.Textstyle; }

                ObjectId createdId = DimStyleHelper.EnsureOrCreateDimStyle(tr, target, 3, uiScale, textStyleId);

                // 5) 提交事务（EnsureOrCreateDimStyle 在内部可能已修改表，但要确认提交）
                tr.Commit();

                // 6) 比较返回值是否等同数据库默认 dimstyle（这是 EnsureOrCreateDimStyle 的失败回退）
                var db = activeDb;
                bool returnedIsDefault = (createdId == db?.Dimstyle);
                Env.Editor.WriteMessage($"\n[DEBUG] EnsureOrCreateDimStyle 返回: {createdId}  (是否回退到 db.Dimstyle: {returnedIsDefault})");

                // 7) 再次打开新事务列出样式，确认是否可见
                using var tr2 = new DBTrans();
                var after = DimStyleHelper.ListDimStyleNames(tr2) ?? new List<string>();
                string afterList = after.Count == 0 ? "<empty>" : string.Join("\n", after);
                Env.Editor.WriteMessage($"\n[DEBUG] 当前 DimStyle 列表（创建后）:\n{afterList}");

                // 8) 列出所有近似匹配（包含目标字串 / 忽略不可见字符），方便排查隐藏字符或类似名称
                var similars = after.Where(n =>
                    (n ?? string.Empty).IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    string.Equals((n ?? string.Empty).Trim().Replace('\u00A0', ' '), target, StringComparison.OrdinalIgnoreCase)
                ).ToList();
                if (similars.Count > 0)
                    Env.Editor.WriteMessage($"\n[DEBUG] 检测到近似名（包含或规范化匹配）:\n{string.Join("\n", similars)}");
                else
                    Env.Editor.WriteMessage($"\n[DEBUG] 未检测到近似名。");

                tr2.Commit();
            }
            catch (Exception ex)
            {
                AutoCadHelper.LogWithSafety($"\nDEBUG_DIMSTYLE_JLPDI 异常: {ex.Message}");
                Env.Editor.WriteMessage($"\nDEBUG_DIMSTYLE_JLPDI 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 端点接触容差（例如 1e-3）
        /// 中文注释：安全读取可能存在的系统变量作为容差，若不存在或读取失败则返回合理的默认值。
        /// </summary>
        public static double GetPipeEndpointTolerance()
        {
            // 尝试从系统变量读取（变量名为示例，可根据项目实际变量名调整）
            // 如果系统变量不存在或读取失败，SafeGetSystemVariableDouble 会返回第二个参数作为默认值
            double tol = SafeGetSystemVariableDouble("PIPE_ENDPOINT_TOL", 0.001);
            // 容错：确保返回为正数，避免后续计算异常
            if (!(tol > 0.0))
            {
                tol = 0.001; // 默认 1mm（可根据项目调整）
            }
            return tol;
        }

        /// <summary>
        /// 共线角度容差（例如 1.0 度转弧度）
        /// 中文注释：从系统变量读取角度容差（单位默认度），返回值为弧度。
        /// </summary>
        public static double GetPipeCollinearAngleToleranceRad()
        {
            // 尝试从系统变量读取角度容差（以度为单位），若不可用则使用 1.0 度为默认值
            double angleDeg = SafeGetSystemVariableDouble("PIPE_COLLINEAR_ANGLE_DEG", 1.0);
            // 容错：确保角度为非负合理值
            if (double.IsNaN(angleDeg) || angleDeg <= 0.0)
            {
                angleDeg = 1.0; // 默认 1 度
            }
            // 将度转换为弧度返回
            double angleRad = angleDeg * (Math.PI / 180.0);
            return angleRad;
        }
    }
}
