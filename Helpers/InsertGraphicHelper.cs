using GB_NewCadPlus_LM.FunctionalMethod;
using GB_NewCadPlus_LM.UniFiedStandards;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_LM.Helpers
{
    public static class InsertGraphicHelper
    {
        /// <summary>
        /// 存储用户点击的点
        /// </summary>
        private static List<Point3d> pointS = new List<Point3d>();

        #region 将外部 DWG 文件插入到当前图纸鼠标指定点 可以插入天正 TCH 实体保留天正等自定义实体属性的快速方法（整图复制）


        #region 再次点方向按键的重复插入逻辑

        /// 放在 CopyDwgAllFast 相关静态字段附近
        private static bool _isCopyDwgAllFastDragging;

        /// <summary>
        /// 新增：命令级互斥标志，防止 Drag 期间重复进入导致崩溃
        /// </summary>
        private static int _copyDwgAllFastBusyFlag = 0;

        /// <summary>
        /// 当前是否处于 COPYDWGALLFAST 的 Drag 交互中
        /// </summary>
        public static bool IsCopyDwgAllFastDragging => _isCopyDwgAllFastDragging;

        /// <summary>
        /// 当前是否处于 COPYDWGALLFAST 执行中（含 Drag）
        /// </summary>
        public static bool IsCopyDwgAllFastBusy => System.Threading.Volatile.Read(ref _copyDwgAllFastBusyFlag) == 1;

        /// <summary>
        /// COPYDWGALLFAST 执行完成事件：success=true 表示插入成功，error 为失败原因（可空）
        /// </summary>
        public static event Action<bool, string?>? CopyDwgAllFastCompleted;

        /// <summary>
        /// 尝试进入 COPYDWGALLFAST 临界区
        /// </summary>
        private static bool TryEnterCopyDwgAllFastBusy()
        {
            return System.Threading.Interlocked.CompareExchange(ref _copyDwgAllFastBusyFlag, 1, 0) == 0;
        }

        /// <summary>
        /// 退出 COPYDWGALLFAST 临界区
        /// </summary>
        private static void ExitCopyDwgAllFastBusy()
        {
            System.Threading.Interlocked.Exchange(ref _copyDwgAllFastBusyFlag, 0);
        }

        /// <summary>
        /// 是否存在可重复执行的上次整图插入
        /// </summary>
        public static bool HasRepeatableCopyDwgAllFast()
        {
            if (_lastCopyDwgBytes != null && _lastCopyDwgBytes.Length > 0) return true;
            return !string.IsNullOrWhiteSpace(_lastCopyDwgPath) && System.IO.File.Exists(_lastCopyDwgPath);
        }

        /// <summary>
        /// 供方向键调用：按当前角度重复一次上次整图插入
        /// </summary>
        public static void RepeatLastCopyDwgAllFastFromDirection()
        {
            if (!HasRepeatableCopyDwgAllFast()) return;
            Env.Document.SendStringToExecute("COPYDWGALLFASTLAST ", false, false, false);
        }

        /// <summary>
        /// 统一派发插入结果，避免事件回调影响主流程
        /// </summary>
        private static void RaiseCopyDwgAllFastCompleted(bool success, string? error = null)
        {
            try
            {
                CopyDwgAllFastCompleted?.Invoke(success, error);
            }
            catch
            {
                // 忽略事件回调异常，避免影响命令本身
            }
        }

        #endregion

        #endregion


        #region  插入图元的核心方法（CopyDwgAllFast）和相关字段

        #region 读取属性和判定重叠的辅助方法

        /// <summary>
        /// 尝试安全获取实体包围盒（防止部分实体抛异常）
        /// </summary>
        private static bool TryGetEntityExtents(Entity entity, out Extents3d extents)
        {
            extents = default; // 初始化包围盒
            if (entity == null) return false; // 空实体直接失败
            if (entity.IsErased) return false; // 已删除实体直接失败
            try
            {
                extents = entity.GeometricExtents; // 读取几何包围盒
                return true; // 成功返回
            }
            catch
            {
                return false; // 读取失败返回 false
            }
        }

        /// <summary>
        /// 判断两个包围盒是否相交（含接触）
        /// </summary>
        private static bool IsExtentsIntersect(Extents3d a, Extents3d b, double tol = 1e-6)
        {
            if (a.MaxPoint.X < b.MinPoint.X - tol) return false; // X 轴分离
            if (a.MinPoint.X > b.MaxPoint.X + tol) return false; // X 轴分离
            if (a.MaxPoint.Y < b.MinPoint.Y - tol) return false; // Y 轴分离
            if (a.MinPoint.Y > b.MaxPoint.Y + tol) return false; // Y 轴分离
            if (a.MaxPoint.Z < b.MinPoint.Z - tol) return false; // Z 轴分离
            if (a.MinPoint.Z > b.MaxPoint.Z + tol) return false; // Z 轴分离
            return true; // 包围盒存在交集
        }

        /// <summary>
        /// 粗精结合的重叠判定：先包围盒，再尝试曲线求交
        /// </summary>
        private static bool IsEntityOverlap(Entity source, Entity target)
        {
            if (!TryGetEntityExtents(source, out var e1)) return false; // 取源包围盒
            if (!TryGetEntityExtents(target, out var e2)) return false; // 取目标包围盒
            if (!IsExtentsIntersect(e1, e2)) return false; // 包围盒不相交直接返回

            // 仅当两者都是 Curve 时做一次更精确求交，减少误判
            if (source is Curve c1 && target is Curve c2)
            {
                try
                {
                    var pts = new Point3dCollection(); // 交点集合
                    c1.IntersectWith(c2, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero); // 求交
                    if (pts.Count > 0) return true; // 有交点即重叠/交叉
                }
                catch
                {
                    // 曲线求交失败时走包围盒结果
                }
            }

            return true; // 非 Curve 或求交失败，包围盒相交即视为重叠
        }

        /// <summary>
        /// 在当前空间里查找第一个与待插入块参照重叠的实体
        /// </summary>
        private static Entity? FindFirstOverlappedEntity(DBTrans tr, BlockReference insertingBr)
        {
            if (tr == null) return null; // 事务判空
            if (insertingBr == null) return null; // 待插入块判空

            foreach (ObjectId id in tr.CurrentSpace) // 遍历当前空间实体
            {
                if (id == insertingBr.ObjectId) continue; // 跳过自身
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity; // 读取实体
                if (ent == null) continue; // 过滤空实体
                if (ent.IsErased) continue; // 过滤已删除实体

                if (IsEntityOverlap(insertingBr, ent)) // 判定是否重叠
                {
                    return ent; // 找到后立即返回
                }
            }

            return null; // 未找到重叠实体
        }

        /// <summary>
        /// 读取实体属性映射（支持：块属性 + 扩展字典XRecord）
        /// </summary>
        private static Dictionary<string, string> ReadEntityPropertyMap(DBTrans tr, Entity entity)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // 使用不区分大小写字典
            if (tr == null || entity == null) return map; // 判空返回

            // 1) 读取块属性（AttributeReference）
            if (entity is BlockReference br)
            {
                foreach (ObjectId attId in br.AttributeCollection) // 遍历块属性
                {
                    var ar = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference; // 读取属性引用
                    if (ar == null) continue; // 空则跳过
                    string tag = (ar.Tag ?? string.Empty).Trim(); // 取属性名
                    if (string.IsNullOrWhiteSpace(tag)) continue; // 空名跳过
                    map[tag] = ar.TextString ?? string.Empty; // 保存键值
                }
            }

            // 2) 读取扩展字典中的 XRecord（键名作为属性名）
            if (entity.ExtensionDictionary != ObjectId.Null) // 判断是否有扩展字典
            {
                var dict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead) as DBDictionary; // 打开扩展字典
                if (dict != null) // 字典有效
                {
                    foreach (DBDictionaryEntry entry in dict) // 遍历字典项
                    {
                        var xrec = tr.GetObject(entry.Value, OpenMode.ForRead) as Xrecord; // 打开 XRecord
                        if (xrec?.Data == null) continue; // 无数据跳过
                        var values = xrec.Data.AsArray(); // 读取 TypedValue 数组
                        if (values == null || values.Length == 0) continue; // 空值跳过
                        string key = (entry.Key ?? string.Empty).Trim(); // 字典键名
                        if (string.IsNullOrWhiteSpace(key)) continue; // 空键跳过
                        string val = values[0].Value?.ToString() ?? string.Empty; // 取第一个值作为字符串
                        map[key] = val; // 保存键值
                    }
                }
            }

            return map; // 返回属性映射
        }

        /// <summary>
        /// 将源属性同步到目标块参照（仅同名属性）
        /// </summary>
        private static void SyncCommonPropertiesToBlockReference(DBTrans tr, BlockReference targetBr, Dictionary<string, string> sourceMap)
        {
            if (tr == null || targetBr == null || sourceMap == null || sourceMap.Count == 0) return; // 判空

            foreach (ObjectId attId in targetBr.AttributeCollection) // 遍历目标块属性
            {
                var ar = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference; // 以写模式打开
                if (ar == null) continue; // 空跳过
                string tag = (ar.Tag ?? string.Empty).Trim(); // 目标属性名
                if (string.IsNullOrWhiteSpace(tag)) continue; // 空名跳过
                if (!sourceMap.TryGetValue(tag, out string val)) continue; // 源里无同名则跳过
                ar.TextString = val ?? string.Empty; // 同名覆盖
            }
        }

        /// <summary>
        /// 将源属性同步到目标实体扩展字典（仅同名键）
        /// </summary>
        private static void SyncCommonPropertiesToEntityXRecord(DBTrans tr, Entity targetEntity, Dictionary<string, string> sourceMap)
        {
            if (tr == null || targetEntity == null || sourceMap == null || sourceMap.Count == 0) return; // 判空
            if (targetEntity.ExtensionDictionary == ObjectId.Null) return; // 无扩展字典直接返回

            var dict = tr.GetObject(targetEntity.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary; // 打开扩展字典
            if (dict == null) return; // 打开失败返回

            foreach (DBDictionaryEntry entry in dict) // 遍历目标键
            {
                string key = (entry.Key ?? string.Empty).Trim(); // 当前键
                if (string.IsNullOrWhiteSpace(key)) continue; // 空键跳过
                if (!sourceMap.TryGetValue(key, out string val)) continue; // 源里无同名跳过

                var xrec = tr.GetObject(entry.Value, OpenMode.ForWrite) as Xrecord; // 打开 XRecord
                if (xrec == null) continue; // 失败跳过

                xrec.Data = new ResultBuffer( // 覆盖 XRecord 数据
                    new TypedValue((int)DxfCode.Text, val ?? string.Empty)); // 写入字符串值
            }
        }

        /// <summary>
        /// 对“分解后的新实体集合”执行同名属性同步
        /// </summary>
        private static void SyncCommonPropertiesToInsertedEntities(DBTrans tr, List<Entity> insertedEntities, Dictionary<string, string> sourceMap)
        {
            if (tr == null || insertedEntities == null || sourceMap == null || sourceMap.Count == 0) return; // 判空

            foreach (var ent in insertedEntities) // 遍历每个新实体
            {
                if (ent == null) continue; // 空跳过

                if (ent is BlockReference br) // 若是块参照，先同步块属性
                {
                    SyncCommonPropertiesToBlockReference(tr, br, sourceMap); // 同步同名块属性
                }

                SyncCommonPropertiesToEntityXRecord(tr, ent, sourceMap); // 再同步扩展字典同名键
            }
        }

        #endregion


        // 上一次“整图插入”缓存
        private static byte[]? _lastCopyDwgBytes;
        private static string? _lastCopyDwgFileNameBase;
        private static string? _lastCopyDwgPath;

        /// <summary>
        /// 按钮侧统一调用：注册“上次插入参数”，并通过命令行触发，保证空格可重复
        /// </summary>
        public static void ExecuteCopyDwgAllFastWithRepeat(string sourceFilePath)
        {
            // 新增：Drag/执行中禁止再次触发，避免命令重入导致 CAD 崩溃
            if (IsCopyDwgAllFastDragging || IsCopyDwgAllFastBusy)
            {
                Env.Editor?.WriteMessage("\n当前图元正在跟随插入，请先左键落点或按 Esc 结束后再触发。");
                return;
            }

            try
            {
                // 优先缓存资源字节（最稳，避免临时文件被删后无法重复）
                if (VariableDictionary.resourcesFile != null && VariableDictionary.resourcesFile.Length > 0)
                {
                    _lastCopyDwgBytes = (byte[])VariableDictionary.resourcesFile.Clone();
                    _lastCopyDwgFileNameBase = string.IsNullOrWhiteSpace(VariableDictionary.btnFileName)
                        ? "GB_CopyDwgAllFast"
                        : VariableDictionary.btnFileName;
                    _lastCopyDwgPath = null;
                }
                else
                {
                    _lastCopyDwgBytes = null;
                    _lastCopyDwgFileNameBase = Path.GetFileNameWithoutExtension(sourceFilePath);
                    _lastCopyDwgPath = sourceFilePath;
                }
            }
            catch
            {
                // 缓存失败不阻断主流程
            }

            // 关键：通过命令执行，这样空格会重复这个命令
            Env.Document.SendStringToExecute("COPYDWGALLFASTLAST ", false, false, false);
        }

        /// <summary>
        /// 可被空格重复的命令：再次执行上一次插入
        /// </summary>
        [CommandMethod("COPYDWGALLFASTLAST")]
        public static void CopyDwgAllFastLast()
        {
            // 新增：执行中直接拦截，防止重入
            if (IsCopyDwgAllFastDragging || IsCopyDwgAllFastBusy)
            {
                Env.Editor?.WriteMessage("\n当前图元正在跟随插入，请先完成当前插入。");
                return;
            }

            try
            {
                string? runPath = null;

                // 如果有字节缓存，则每次重复都新建一个临时文件
                if (_lastCopyDwgBytes != null && _lastCopyDwgBytes.Length > 0)
                {
                    string baseName = string.IsNullOrWhiteSpace(_lastCopyDwgFileNameBase)
                        ? "GB_CopyDwgAllFast"
                        : _lastCopyDwgFileNameBase;

                    runPath = Path.Combine(Path.GetTempPath(), $"{baseName}_{Guid.NewGuid():N}.dwg");
                    System.IO.File.WriteAllBytes(runPath, _lastCopyDwgBytes);
                }
                else
                {
                    runPath = _lastCopyDwgPath;
                }

                if (string.IsNullOrWhiteSpace(runPath) || !System.IO.File.Exists(runPath))
                {
                    Env.Editor.WriteMessage("\n没有可重复的上一次插入命令。");
                    return;
                }

                if (runPath != null)
                    //插入源文件中的图元到当前图纸
                    CopyDwgAllFast(runPath);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n重复执行上次插入失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 以“整图复制”的方式，将指定 DWG 文件的模型空间全部图元一次性克隆到当前图纸空间，
        /// 等效于在源图里全选 Ctrl+C，再在当前图 Ctrl+V，能够保留天正等自定义实体的附加属性。
        /// </summary>        
        [CommandMethod("COPYDWGALLFAST")] // 注册 CAD 命令名
        public static void CopyDwgAllFast(string sourceFilePath) // 整图插入主方法
        {
            var doc = Application.DocumentManager.MdiActiveDocument; // 获取当前活动文档
            if (doc == null) // 如果没有活动文档
            {
                RaiseCopyDwgAllFastCompleted(false, "未找到活动文档。"); // 回调失败状态
                return; // 结束方法
            }

            if (!TryEnterCopyDwgAllFastBusy()) // 尝试进入互斥区，防止重入
            {
                Env.Editor?.WriteMessage("\n当前插入命令正在执行中，请先完成当前插入（左键落点或 Esc）。"); // 命令行提示
                RaiseCopyDwgAllFastCompleted(false, "当前命令忙碌中。"); // 回调失败状态
                return; // 结束方法
            }

            var destDb = doc.Database; // 获取目标数据库
            bool deleteAfter = false; // 是否在成功后删除源文件（临时文件场景）
            bool insertSuccess = false; // 记录插入是否成功
            string? failReason = null; // 记录失败原因

            try // 主流程异常保护
            {
                try // 判断是否临时文件并设置删除标记
                {
                    if (!string.IsNullOrEmpty(sourceFilePath)) // 如果传入了源文件路径
                    {
                        var tempDir = Path.GetTempPath(); // 获取系统临时目录
                        if (!string.IsNullOrEmpty(tempDir) && sourceFilePath.StartsWith(tempDir, StringComparison.OrdinalIgnoreCase)) // 如果源文件在临时目录下
                        {
                            deleteAfter = true; // 标记成功后删除
                        }
                    }
                }
                catch // 临时文件判断异常时兜底
                {
                    deleteAfter = false; // 异常时不删文件
                }

                Point3d targetPoint = Point3d.Origin; // 初始化插入点（拖拽过程中更新）
                using (doc.LockDocument()) // 锁定文档，确保线程安全写图
                using (var sourceDb = new Autodesk.AutoCAD.DatabaseServices.Database(false, true)) // 打开源数据库
                {
                    sourceDb.ReadDwgFile(sourceFilePath, FileShare.Read, true, null); // 读取源 DWG
                    sourceDb.CloseInput(true); // 关闭输入流，减少文件占用

                    using (var tr = new DBTrans()) // 开启目标库事务
                    {
                        var blkDefId = destDb.Insert("*U", sourceDb, false); // 将整张源图插入为匿名块定义
                        var br = new BlockReference(targetPoint, blkDefId); // 创建块参照实体

                        double scale = 1.0; // 初始化比例
                        if (VariableDictionary.winForm_Status) // 若当前由 WinForm 发起
                        {
                            try // 安全读取 WinForm 比例
                            {
                                scale = VariableDictionary.textBoxScale; // 使用 WinForm 比例
                            }
                            catch // 读取异常
                            {
                                scale = 1.0; // 异常回退默认比例
                            }
                        }
                        else // 当前由 WPF 发起
                        {
                            try // 安全读取 WPF 比例
                            {
                                scale = VariableDictionary.wpfTextBoxScale; // 使用 WPF 比例
                            }
                            catch // 读取异常
                            {
                                scale = 1.0; // 异常回退默认比例
                            }
                        }
                        if (double.IsNaN(scale) || scale <= 0) scale = 1.0; // 比例非法时强制为 1

                        br.ScaleFactors = new Scale3d(scale); // 设置初始比例
                        var entityObjectId = tr.CurrentSpace.AddEntity(br); // 将块参照加入当前空间
                        var fileEntity = (BlockReference)tr.GetObject(entityObjectId, OpenMode.ForWrite); // 以可写模式打开块参照

                        double tempAngle = VariableDictionary.entityRotateAngle; // 记录当前旋转角度
                        double tempScale = scale; // 记录当前比例

                        if (Math.Abs(tempAngle) > 1e-12) // 若初始角度非零
                        {
                            fileEntity.TransformBy(Matrix3d.Rotation(tempAngle, Vector3d.ZAxis, targetPoint)); // 预先旋转，确保预览方向正确
                        }

                        var entityBlock = new JigEx((mpw, _) => // 创建拖拽 Jig
                        {
                            fileEntity.Move(targetPoint, mpw); // 将块从旧点移动到新点
                            targetPoint = mpw; // 更新当前插入点

                            if (VariableDictionary.entityRotateAngle != tempAngle) // 若外部旋转角度发生变化
                            {
                                fileEntity.TransformBy(Matrix3d.Rotation(-tempAngle, Vector3d.ZAxis, targetPoint)); // 先撤销旧角度
                                tempAngle = VariableDictionary.entityRotateAngle; // 更新角度缓存
                                fileEntity.TransformBy(Matrix3d.Rotation(tempAngle, Vector3d.ZAxis, targetPoint)); // 应用新角度
                            }

                            double currentUiScale = scale; // 默认沿用初始比例
                            if (VariableDictionary.winForm_Status) // WinForm 模式读取实时比例
                            {
                                try { currentUiScale = VariableDictionary.textBoxScale; } catch { currentUiScale = tempScale; } // 读取失败沿用旧值
                            }
                            else // WPF 模式读取实时比例
                            {
                                try { currentUiScale = VariableDictionary.wpfTextBoxScale; } catch { currentUiScale = tempScale; } // 读取失败沿用旧值
                            }
                            if (double.IsNaN(currentUiScale) || currentUiScale <= 0) currentUiScale = 1.0; // 非法比例兜底

                            if (Math.Abs(currentUiScale - tempScale) > 1e-9) // 比例有变化才更新
                            {
                                fileEntity.ScaleFactors = new Scale3d(currentUiScale); // 设置新比例
                                tempScale = currentUiScale; // 更新比例缓存
                            }
                        });

                        entityBlock.DatabaseEntityDraw(wd => wd.Geometry.Draw(fileEntity)); // 绘制拖拽预览
                        entityBlock.SetOptions(msg: "\n指定插入点"); // 设置交互提示
                        EnsureDwgViewFocus(); // 尝试切回绘图区焦点

                        PromptResult endPoint; // 接收拖拽结果
                        _isCopyDwgAllFastDragging = true; // 标记拖拽中
                        try // 执行拖拽
                        {
                            endPoint = Env.Editor.Drag(entityBlock); // 启动交互拖拽
                        }
                        finally // 无论成功失败都清除拖拽标记
                        {
                            _isCopyDwgAllFastDragging = false; // 清理拖拽状态
                        }

                        if (endPoint.Status != PromptStatus.OK) // 如果用户取消拖拽
                        {
                            failReason = "用户取消插入。"; // 记录失败原因
                            tr.Abort(); // 回滚事务
                            return; // 结束方法
                        }

                        var overlapSourcePropertyMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // 准备重叠源属性字典
                        var overlappedEntity = FindFirstOverlappedEntity(tr, fileEntity); // 查找与待插入块参照重叠的已有实体

                        if (overlappedEntity != null) // 如果存在重叠实体
                        {
                            overlapSourcePropertyMap = ReadEntityPropertyMap(tr, overlappedEntity); // 读取重叠实体属性
                            if (overlapSourcePropertyMap.Count > 0) // 只有读取到属性才同步
                            {
                                SyncCommonPropertiesToBlockReference(tr, fileEntity, overlapSourcePropertyMap); // 先同步到块属性同名项
                                SyncCommonPropertiesToEntityXRecord(tr, fileEntity, overlapSourcePropertyMap); // 再同步扩展字典同名项
                                LogManager.Instance.LogInfo($"\n检测到重叠图元，已预同步属性：{overlapSourcePropertyMap.Count} 项。"); // 记录同步日志
                            }
                        }

                        var newIds = new DBObjectCollection(); // 创建分解结果容器
                        fileEntity.Explode(newIds); // 分解块参照
                        fileEntity.Erase(); // 删除原块参照（保留分解实体）

                        var insertedEntities = new List<Entity>(); // 记录新加入空间的实体
                        foreach (Entity ent in newIds) // 遍历分解后的实体
                        {
                            tr.CurrentSpace.AddEntity(ent); // 添加到当前空间
                            insertedEntities.Add(ent); // 收集到列表
                        }

                        if (overlapSourcePropertyMap.Count > 0) // 若存在可同步属性
                        {
                            SyncCommonPropertiesToInsertedEntities(tr, insertedEntities, overlapSourcePropertyMap); // 同步到分解后实体
                        }

                        if (VariableDictionary.dimString != null) // 如果有待创建的标注文本
                        {
                            try // 尝试创建标注
                            {
                                Command.DDimLinear(tr, entityBlock.MousePointWcsLast, VariableDictionary.dimString); // 创建标注
                                VariableDictionary.dimString = null; // 清空标注缓存
                            }
                            catch (Exception ex) // 标注失败不影响插入
                            {
                                LogManager.Instance.LogInfo($"\n设置标注样式失败: {ex.Message}"); // 记录日志
                            }
                        }

                        tr.Commit(); // 提交事务
                        Env.Editor.Redraw(); // 刷新界面
                        insertSuccess = true; // 标记成功
                    }
                }

                if (deleteAfter && insertSuccess && !string.IsNullOrEmpty(sourceFilePath)) // 如果是临时文件且插入成功
                {
                    try // 尝试删除临时文件
                    {
                        if (System.IO.File.Exists(sourceFilePath)) // 文件存在才删除
                        {
                            System.IO.File.Delete(sourceFilePath); // 删除文件
                            LogManager.Instance.LogInfo($"\n已删除临时 DWG 文件: {sourceFilePath}"); // 记录删除日志
                        }
                    }
                    catch (Exception ex) // 删除失败仅记录不抛出
                    {
                        LogManager.Instance.LogInfo($"\n尝试删除临时 DWG 失败: {ex.Message}"); // 记录失败日志
                    }
                }

                VariableDictionary.winForm_Status = false; // 重置 WinForm 状态
            }
            catch (Exception ex) // 捕获主流程异常
            {
                failReason = ex.Message; // 保存失败信息
                LogManager.Instance.LogInfo($"\nCOPYDWGALLFAST 执行失败: {ex.Message}"); // 写日志
            }
            finally // 统一收尾
            {
                _isCopyDwgAllFastDragging = false; // 兜底清除拖拽状态
                ExitCopyDwgAllFastBusy(); // 释放互斥锁
                RaiseCopyDwgAllFastCompleted(insertSuccess, insertSuccess ? null : failReason); // 回调最终执行结果
            }
        }


        /// <summary>
        /// 焦点切换：尝试将焦点切换回 AutoCAD 的绘图区域，确保用户在插入块后能够直接看到并操作新插入的图元。
        /// </summary>
        private static void EnsureDwgViewFocus()
        {
            try
            {
                // 优先尝试 AutoCAD 内部焦点切换（用反射避免强依赖）
                var t1 = Type.GetType("Autodesk.AutoCAD.Internal.Utils, AcMgd", false);
                var m1 = t1?.GetMethod("SetFocusToDwgView", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                if (m1 != null)
                {
                    m1.Invoke(null, null);
                }
                else
                {
                    // 兼容部分版本程序集名
                    var t2 = Type.GetType("Autodesk.AutoCAD.Internal.Utils, AcCoreMgd", false);
                    var m2 = t2?.GetMethod("SetFocusToDwgView", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                    m2?.Invoke(null, null);
                }
            }
            catch
            {
                // 忽略焦点切换异常，避免中断主流程
            }

            try
            {
                System.Windows.Forms.Application.DoEvents();
            }
            catch { }

            try
            {
                Env.Editor.UpdateScreen();
            }
            catch { }
        }

        /// <summary>
        /// 插入外部条件图元
        /// </summary>
        [CommandMethod(nameof(GB_InsertBlock_Ptj))]
        public static void GB_InsertBlock_Ptj()
        {
            #region 方法2：
            try
            {
                Directory.CreateDirectory(GetPath.referenceFile);  //获取到本工具的系统目录；
                if (VariableDictionary.btnFileName == null) return; //判断点现的按键名是不是空；
                if (VariableDictionary.resourcesFile == null) return; //判断点现的原文件是不是空；
                using var tr = new DBTrans();
                var referenceFileObId = tr.BlockTable.GetBlockFormA(VariableDictionary.resourcesFile, VariableDictionary.btnFileName, true);//拿到本工具的系统目录下的按键名的原文件的objectid；
                var refFileRec = tr.GetObject(referenceFileObId, OpenMode.ForRead) as BlockTableRecord;//拿到原文件的块表记录；
                int isNo = -1;
                if (refFileRec != null)
                    foreach (var fileId in refFileRec)
                    {
                        isNo++;
                        if (VariableDictionary.btnFileName == "PTJ_给水点") isNo = -1;
                        else if (isNo != VariableDictionary.TCH_Ptj_No) continue;
                        //判读是不是为0，0就是天正元素
                        //if (fileId.ObjectClass.DxfName == "INSERT") continue;
                        var fileEntity = tr.GetObject(fileId, OpenMode.ForRead) as Entity;
                        //LogManager.Instance.LogInfo("PTJ元素！");
                        if (fileEntity == null) continue;
                        var fileEntityCopy = fileEntity.Clone() as Entity;
                        //tr.CurrentSpace.DeepCloneEx(fileEntity,)//深度克隆，可以复制天正图元；
                        //(vlax-dump-object (vlax-ename->vla-object (car (entsel )))T)//这个是在cad命令里能读出天正属性的lisp命令；
                        if (fileEntityCopy == null) continue;
                        var dxfName = fileId.ObjectClass.DxfName;//抓取图元的DXFName,判断是不是天正的图元
                        var fileType = dxfName.Split('_');//截取‘_’字符
                        if (fileType[0] == "TCH")
                        {
                            LogManager.Instance.LogInfo("TCH！");
                            var fileEntityCopyObId = tr.CurrentSpace.AddEntity(fileEntityCopy);
                            double tempAngle = 0;
                            var startPoint = new Point3d(0, 0, 0);
                            var entityBlock = new JigEx((mpw, _) =>
                            {
                                fileEntityCopy.Move(startPoint, mpw);
                                startPoint = mpw;
                                if (VariableDictionary.entityRotateAngle == tempAngle)
                                {
                                    return;
                                }
                                else if (VariableDictionary.entityRotateAngle != tempAngle)
                                {
                                    fileEntityCopy.Rotation(center: mpw, 0);
                                    tempAngle = VariableDictionary.entityRotateAngle;
                                    fileEntityCopy.Rotation(center: mpw, tempAngle);
                                }
                            });
                            entityBlock.DatabaseEntityDraw(wd => wd.Geometry.Draw(fileEntityCopy));
                            entityBlock.SetOptions(msg: "\n指定插入点");
                            //entityBlock.SetOptions(startPoint, msg: "\n指定插入点");这个startpoint，是有个参考线在里面，用于托拽时的辅助；
                            var endPoint = Env.Editor.Drag(entityBlock);
                            if (endPoint.Status != PromptStatus.OK) return;
                            tr.BlockTable.Remove(referenceFileObId);
                        }
                        else if (fileEntityCopy is BlockReference)
                        {
                            LogManager.Instance.LogInfo("PTJ-块表记录！");
                            //if (fileEntityCopy.ColorIndex.ToString() != "130") return;
                            var fileEntityCopyObId = tr.CurrentSpace.AddEntity(fileEntityCopy);//在当前图纸空间中加入这个实体并获取它的ObjoectId
                            double tempAngle = 0;
                            var startPoint = new Point3d(0, 0, 0);
                            var entityBlock = new JigEx((mpw, _) =>
                            {
                                fileEntityCopy.Move(startPoint, mpw);
                                startPoint = mpw;
                                if (VariableDictionary.entityRotateAngle == tempAngle)
                                {
                                    return;
                                }
                                else if (VariableDictionary.entityRotateAngle != tempAngle)
                                {
                                    fileEntityCopy.Rotation(center: mpw, 0);
                                    tempAngle = VariableDictionary.entityRotateAngle;
                                    fileEntityCopy.Rotation(center: mpw, tempAngle);
                                }
                            });
                            entityBlock.DatabaseEntityDraw(wd => wd.Geometry.Draw(fileEntityCopy));
                            entityBlock.SetOptions(msg: "\n指定插入点");
                            //entityBlock.SetOptions(startPoint, msg: "\n指定插入点");这个startpoint，是有个参考线在里面，用于托拽时的辅助；
                            var endPoint = Env.Editor.Drag(entityBlock);
                            if (endPoint.Status != PromptStatus.OK) return;
                            tr.BlockTable.Remove(referenceFileObId);
                        }
                        //else
                        //{
                        //    LogManager.Instance.LogInfo("PTJ-块！");
                        //    var referenceFileBlock = tr.CurrentSpace.InsertBlock(Point3d.Origin, referenceFileObId);
                        //    //tr.BlockTable.Remove(referenceFileObId);
                        //    if (tr.GetObject(referenceFileBlock) is not Entity referenceFileEntity) return;
                        //    double tempAngle = 0;
                        //    var startPoint = new Point3d(0, 0, 0);
                        //    var entityBlock = new JigEx((mpw, _) =>
                        //    {
                        //        referenceFileEntity.Move(startPoint, mpw);
                        //        startPoint = mpw;
                        //        if (VariableDictionary.entityRotateAngle == tempAngle)
                        //        {
                        //            return;
                        //        }
                        //        else if (VariableDictionary.entityRotateAngle != tempAngle)
                        //        {
                        //            referenceFileEntity.Rotation(center: mpw, 0);
                        //            tempAngle = VariableDictionary.entityRotateAngle;
                        //            referenceFileEntity.Rotation(center: mpw, tempAngle);
                        //        }
                        //    });
                        //    entityBlock.DatabaseEntityDraw(wd => wd.Geometry.Draw(referenceFileEntity));
                        //    entityBlock.SetOptions(msg: "\n指定插入点");
                        //    var endPoint = Env.Editor.Drag(entityBlock);
                        //    if (endPoint.Status != PromptStatus.OK) return;
                        //    referenceFileEntity.Layer = VariableDictionary.btnBlockLayer;
                        //    break;
                        //}
                    }

                tr.Commit();
                Env.Editor.Redraw();

            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("插入图元失败！");
                LogManager.Instance.LogInfo("错误信息: " + ex.Message);
            }
            #endregion
        }


        /// <summary>
        /// 一个块反复插入图中
        /// </summary>
        [CommandMethod(nameof(GB_InsertBlock_5))]
        public static void GB_InsertBlock_5()
        {
            #region 方法1：  
            try
            {
                pointS.Clear();
                Directory.CreateDirectory(GetPath.referenceFile);
                if (VariableDictionary.btnFileName == null) return;

                if (VariableDictionary.resourcesFile == null) return; //判断点现的原文件是不是空；
                using var tr = new DBTrans();

                // 获取对应块的 ObjectId  
                var referenceFileObId = tr.BlockTable.GetBlockFormA(
                    VariableDictionary.resourcesFile,
                    VariableDictionary.btnFileName,
                    VariableDictionary.btnFileName_blockName,
                    true);

                var refFileRec = tr.GetObject(referenceFileObId, OpenMode.ForRead) as BlockTableRecord;
                if (refFileRec == null)
                {
                    LogManager.Instance.LogInfo("未找到块记录！");
                    return;
                }

                LogManager.Instance.LogInfo("块！");
                while (true)
                {
                    // 把块插入到当前空间  
                    var referenceFileBlock = tr.CurrentSpace.InsertBlock(Point3d.Origin, referenceFileObId);

                    // 检查是否为实体  
                    if (tr.GetObject(referenceFileBlock) is not Entity referenceFileEntity)
                        return;

                    // 设置图层和颜色等属性  
                    referenceFileEntity.Layer = VariableDictionary.btnBlockLayer;
                    referenceFileEntity.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                    referenceFileEntity.Scale(new Point3d(0, 0, 0), VariableDictionary.blockScale);

                    //double tempAngle = 0; // 原始角度  
                    var startPoint = new Point3d(0, 0, 0);

                    var jigBlock = new JigEx((mpw, _) =>
                    {
                        // 先移动  
                        referenceFileEntity.Move(startPoint, mpw);
                        startPoint = mpw;
                    });
                    jigBlock.DatabaseEntityDraw(wd => wd.Geometry.Draw(referenceFileEntity));
                    jigBlock.SetOptions(msg: "\n指定插入点");

                    // 拖拽  
                    var endPoint = Env.Editor.Drag(jigBlock);
                    if (endPoint.Status != PromptStatus.OK)
                    {
                        // 用户取消插入，则删除已插入的块  
                        tr.GetObject(referenceFileBlock, OpenMode.ForWrite);
                        referenceFileBlock.Erase();
                        break;
                    }

                    // 存储插入点坐标（WCS）  
                    var UcsEndPoint = jigBlock.MousePointWcsLast;
                    pointS.Add(UcsEndPoint);


                    Env.Editor.Redraw();
                }

                // ======================  
                // 在此处根据插入数量绘图  
                // ======================  
                int count = pointS.Count;
                LogManager.Instance.LogInfo($"\n已插入 {count} 个块，开始绘制外围图形...");

                if (count == 3)
                { // 三点生成外接圆，圆心与3点等距  
                    Point3d p1 = pointS[0];
                    Point3d p2 = pointS[1];
                    Point3d p3 = pointS[2];

                    // 计算三角形外接圆圆心（与三点等距的点）  
                    Point3d circleCenter = GetCircumcenter(p1, p2, p3);

                    // 计算圆心到三个点的距离，取最大值，然后加上150作为新圆的半径  
                    double radius = p1.DistanceTo(circleCenter) + 150.0;

                    // 检查计算出的圆心是否与三点等距  
                    double dist1 = circleCenter.DistanceTo(p1);
                    double dist2 = circleCenter.DistanceTo(p2);
                    double dist3 = circleCenter.DistanceTo(p3);

                    // 记录到日志，以验证计算正确性  
                    LogManager.Instance.LogInfo($"\n圆心到三点的距离: {dist1:F4}, {dist2:F4}, {dist3:F4}");

                    // 正确创建圆：使用外接圆圆心和半径  
                    var circle = new Circle(circleCenter, Vector3d.ZAxis, radius);
                    circle.Layer = VariableDictionary.btnBlockLayer;
                    circle.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                    tr.CurrentSpace.AddEntity(circle);
                    Env.Editor.Redraw();
                    LogManager.Instance.LogInfo("\n已创建外围圆形，与三点等距并向外扩展150。");
                }
                else if (count == 4)
                {
                    // 计算中心点  
                    var center = new Point3d(
                        pointS.Average(p => p.X),
                        pointS.Average(p => p.Y),
                        pointS.Average(p => p.Z)
                    );

                    // 绘制矩形 - 使用原始4点作为矩形顶点，向外扩展150  
                    List<Point2d> expandedPoints = new List<Point2d>();

                    foreach (var point in pointS)
                    {
                        // 计算从中心到点的方向向量  
                        Vector3d dirVector = point - center;
                        dirVector = dirVector.GetNormal(); // 单位化向量  

                        // 创建新点：沿着方向向量延伸150的距离  
                        Point3d expandedPoint3d = point + dirVector * 150.0;
                        Point2d expandedPoint = new Point2d(expandedPoint3d.X, expandedPoint3d.Y);
                        expandedPoints.Add(expandedPoint);
                    }

                    // 确保点按顺时针或逆时针排序  
                    // 对顶点按角度排序  
                    var sortedPoints = expandedPoints.Select((p, index) => new
                    {
                        Point = p,
                        Angle = Math.Atan2(p.Y - center.Y, p.X - center.X)
                    })
                    .OrderBy(item => item.Angle)
                    .Select(item => item.Point)
                    .ToList();

                    // 创建Polyline并添加扩展后的顶点  
                    var pl = new Polyline();
                    for (int i = 0; i < sortedPoints.Count; i++)
                    {
                        pl.AddVertexAt(i, sortedPoints[i], 0, 30, 30);
                    }

                    // 闭合  
                    pl.Closed = true;
                    pl.Layer = VariableDictionary.btnBlockLayer;
                    pl.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                    tr.CurrentSpace.AddEntity(pl);
                    Env.Editor.Redraw();
                    LogManager.Instance.LogInfo("\n已创建外围矩形，向外扩展150。");
                }
                else if (count > 4)
                {
                    // 计算中心点  
                    var center = new Point3d(
                        pointS.Average(p => p.X),
                        pointS.Average(p => p.Y),
                        pointS.Average(p => p.Z)
                    );

                    // 创建多边形 - 使用原始点作为多边形顶点，向外扩展150  
                    List<Point2d> expandedPoints = new List<Point2d>();
                    foreach (var point in pointS)
                    {
                        // 计算从中心到点的方向向量  
                        Vector3d dirVector = point - center;
                        dirVector = dirVector.GetNormal(); // 单位化向量  
                        // 创建新点：沿着方向向量延伸150的距离  
                        Point3d expandedPoint3d = point + dirVector * 150.0;
                        Point2d expandedPoint = new Point2d(expandedPoint3d.X, expandedPoint3d.Y);
                        expandedPoints.Add(expandedPoint);
                    }
                    // 对顶点按角度排序，确保多边形正确  
                    var sortedPoints = expandedPoints.Select((p, index) => new
                    {
                        Point = p,
                        Angle = Math.Atan2(p.Y - center.Y, p.X - center.X)
                    })
                    .OrderBy(item => item.Angle)
                    .Select(item => item.Point)
                    .ToList();

                    // 创建多边形  
                    var polygon = new Polyline();
                    for (int i = 0; i < sortedPoints.Count; i++)
                    {
                        polygon.AddVertexAt(i, sortedPoints[i], 0, 30, 30);
                    }
                    // 闭合多边形  
                    polygon.Closed = true;
                    polygon.Layer = VariableDictionary.btnBlockLayer;
                    polygon.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                    tr.CurrentSpace.AddEntity(polygon);
                    Env.Editor.Redraw();
                    LogManager.Instance.LogInfo($"\n已创建{count}边形外围，向外扩展150。");
                }
                else if (count > 0)
                {
                    LogManager.Instance.LogInfo($"\n已插入{count}个块，但数量不满足绘制外围图形的条件（需要至少3个点）。");
                }
                //加标注
                // DDimLinear("总重:" + VariableDictionary.dimString + "kg" + "\n" + $"{count}点着地", Convert.ToInt16(pointS.Count));
                var dimColorLine = VariableDictionary.layerColorIndex;
                if (VariableDictionary.btnFileName.Contains("结构"))
                {
                    dimColorLine = 3;
                }
                if (VariableDictionary.dimString != null)
                    Command.DDimLinear(tr, VariableDictionary.dimString, count.ToString(), Convert.ToInt16(dimColorLine));
                tr.Commit();
                Env.Editor.Redraw();
                LogManager.Instance.LogInfo("\n操作完成。");
                pointS.Clear();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n插入图元失败：{ex.Message}");
                // 可以添加更详细的错误信息记录  
                LogManager.Instance.LogInfo($"\n错误详情：{ex.StackTrace}");
            }
            #endregion
        }


        /// <summary>
        /// 计算三角形外接圆圆心，确保圆心与三个点等距 
        /// </summary>
        /// <param name="A">A</param>
        /// <param name="B">B</param>
        /// <param name="C">C</param>
        /// <returns></returns>
        private static Point3d GetCircumcenter(Point3d A, Point3d B, Point3d C)
        {
            // 处理共线情况：如果三点共线，则返回三点的平均点  
            if (ArePointsCollinear(A, B, C))
            {
                return new Point3d(
                    (A.X + B.X + C.X) / 3.0,
                    (A.Y + B.Y + C.Y) / 3.0,
                    (A.Z + B.Z + C.Z) / 3.0
                );
            }
            // 计算分母 d  
            double d = 2 * (A.X * (B.Y - C.Y) + B.X * (C.Y - A.Y) + C.X * (A.Y - B.Y));
            if (Math.Abs(d) < 1e-10)
            {
                // 保护性返回：当分母太小时则视为共线，返回平均值  
                return new Point3d(
                    (A.X + B.X + C.X) / 3.0,
                    (A.Y + B.Y + C.Y) / 3.0,
                    (A.Z + B.Z + C.Z) / 3.0
                );
            }
            // 分别计算各点的 (x^2 + y^2)  
            double Asq = A.X * A.X + A.Y * A.Y;
            double Bsq = B.X * B.X + B.Y * B.Y;
            double Csq = C.X * C.X + C.Y * C.Y;
            // 使用标准公式计算圆心 X/Y 坐标  
            double centerX = (Asq * (B.Y - C.Y) + Bsq * (C.Y - A.Y) + Csq * (A.Y - B.Y)) / d;
            double centerY = (Asq * (C.X - B.X) + Bsq * (A.X - C.X) + Csq * (B.X - A.X)) / d;
            double centerZ = (A.Z + B.Z + C.Z) / 3.0; // Z 坐标取平均值  
            return new Point3d(centerX, centerY, centerZ);
        }

        /// <summary>
        /// 检查三点是否共线
        /// </summary>
        /// <param name="A">A</param>
        /// <param name="B">B</param>
        /// <param name="C">C</param>
        /// <returns></returns>

        private static bool ArePointsCollinear(Point3d A, Point3d B, Point3d C)
        {
            Vector3d v1 = B - A;
            Vector3d v2 = C - A;
            Vector3d crossProduct = v1.CrossProduct(v2);
            return crossProduct.Length < 1e-8;
        }
        #endregion
    }
}
