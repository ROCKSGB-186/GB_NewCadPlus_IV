using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// AutoCAD操作工具类
    /// </summary>
    public static class AutoCadHelper
    {
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
        /// 缓存比例
        /// </summary>
        private static double _cachedScale = 1.0;

        /// <summary>
        /// 获取绘图比例
        /// </summary>
        /// <param name="useCache">是否优先使用缓存值</param>
        /// <returns>返回用户设置的比例值，如果未设置则返回 1.0</returns>
        public static double GetScale(bool useCache = true)
        {
            try
            {
                // 2. 从界面读取（按优先级：WinForm > WPF > 默认值）
                double scale;
                // 优先 WinForm
                if (VariableDictionary.winForm_Status)
                {
                     scale = GetDrawingScaleFrom_Winform();
                   
                }
                else
                {
                    scale = GetDrawingScaleFrom_Wpf();
                }
                // 4. 所有方法都失败，返回默认值
                LogManager.Instance.LogInfo($"使用比例 {scale}");
                return scale;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"GetScale 异常: {ex.Message}");
                return 1.0;
            }
        }
        
        /// <summary>
        /// 从 WinForm 界面读取比例值
        /// </summary>
        private static double GetDrawingScaleFrom_Winform()
        {
            try
            {
                // 从窗体实例读取
                var form = System.Windows.Forms.Application.OpenForms
                    .OfType<FormMain>()
                    .FirstOrDefault();

                if (form != null)
                {
                    string text = GetWinFormControlText(form, "textBox_Scale_比例");
                    if (double.TryParse(text, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double result))
                    {
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"读取 WinForm 比例失败: {ex.Message}");
            }
            return 0;
        }
       
        /// <summary>
        /// 从 WPF 界面读取比例值
        /// </summary>
        private static double GetDrawingScaleFrom_Wpf()
        {
            try
            {
                var instance = WpfMainWindow.Instance;      // 获取当前 WPF 界面实例
                if (instance == null)
                    return 100.0;                           // 无实例时的默认值

                var textBox = instance.FindName("TextBox绘图比例") as System.Windows.Controls.TextBox;
                if (textBox == null)
                    return 100.0;

                // 优先取输入文本
                string text = textBox.Text;
                if (!string.IsNullOrWhiteSpace(text) && double.TryParse(text, out double result))
                {
                    VariableDictionary.wpfTextBoxScale = result; // 保持原有缓存
                    return result;
                }

                // 文本为空或非法时，尝试取 Tag 值
                if (textBox.Tag is string tagValue && double.TryParse(tagValue, out double tagResult))
                {
                    VariableDictionary.wpfTextBoxScale = tagResult;
                    return tagResult;
                }

                return 100.0; // 最终默认值
            }
            catch (Exception ex)
            {
                // 静默记录日志
                LogManager.Instance.LogWarning($"GetDrawingScaleFrom_Wpf 异常: {ex.Message}");
                return 100.0;
            }
        }
        
        /// <summary>
        /// 获取 WinForm 控件的文本
        /// </summary>
        private static string GetWinFormControlText(System.Windows.Forms.Form form, string controlName)
        {
            var control = form.Controls.Find(controlName, true).FirstOrDefault();
            if (control is System.Windows.Forms.TextBox textBox)
                return textBox.Text;
            if (control is System.Windows.Forms.NumericUpDown numericUpDown)
                return numericUpDown.Value.ToString();
            return string.Empty;
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
