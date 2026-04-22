using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Colors;
using IFoxCAD.Cad; // 如果 DBTrans 在其它命名空间，请按项目实际修改
using GB_NewCadPlus_LM.Helpers;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_LM.Helpers
{
    /// <summary>
    /// 标注样式（DimStyle）相关帮助方法
    /// 责任：确保样式存在、按项目规范创建/更新、并提供调试列表方法
    /// </summary>
    public static class DimStyleHelper
    {
        private const double MinTextSize = 0.1; // 最小文字/箭头等尺寸，避免设置为 0 导致异常

        /// <summary>
        /// 确保标注样式存在；若不存在则创建，存在则按当前参数更新。
        /// 返回样式 ObjectId。
        /// 注意：使用 DBTrans（你工程常用的事务包装），函数内做了输入校验并保证 Dimlfac 不为 0。
        /// </summary>
       
        public static ObjectId EnsureOrCreateDimStyle(
            DBTrans tr,
            string styleName,
            short layerColorIndex,
            double textBoxScale,
            ObjectId textStyleId)
        {
            if (tr == null) throw new ArgumentNullException(nameof(tr));

            // 规范化样式名（去首尾空格与 NBSP）
            styleName = (styleName ?? string.Empty).Trim();
            styleName = styleName.Replace('\u00A0', ' ').Trim();
            if (string.IsNullOrEmpty(styleName))
                throw new ArgumentException("样式名不能为空", nameof(styleName));

            if (double.IsNaN(textBoxScale) || textBoxScale <= 0) textBoxScale = 1.0;

            // 如果 DBTrans 指向的数据库不是当前活动文档的数据库，
            // 为避免样式写入到错误的库，改为在活动文档数据库中创建/更新样式并返回其 ObjectId。
            var activeDoc = Application.DocumentManager.MdiActiveDocument;
            var activeDb = activeDoc?.Database;
            if (activeDb != null && !ReferenceEquals(tr.Database, activeDb))
            {
                // 记录诊断日志，方便排查 DBTrans 构造问题
                AutoCadHelper.LogWithSafety($"DimStyleHelper: 检测到 tr.Database != ActiveDocument.Database，改为在活动文档创建/更新样式 '{styleName}'");

                // 在活动文档数据库内执行创建/更新逻辑（独立事务）
                return EnsureOrCreateDimStyleInActiveDatabase(activeDb, styleName, layerColorIndex, textBoxScale, textStyleId);
            }

            // 原始路径：在传入的 tr（DBTrans）所在数据库上操作（保持原实现）
            var db = tr.Database;
            if (textStyleId == ObjectId.Null)
                textStyleId = db.Textstyle;

            var dimStyleTable = tr.GetObject(db.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
            if (dimStyleTable == null) return db.Dimstyle;

            if (dimStyleTable.Has(styleName))
            {
                var existId = dimStyleTable[styleName];
                var existRec = tr.GetObject(existId, OpenMode.ForWrite) as DimStyleTableRecord;
                if (existRec != null)
                {
                    ApplyDimStyleParameters(existRec, layerColorIndex, textBoxScale, textStyleId);
                    return existId;
                }
            }

            foreach (ObjectId id in dimStyleTable)
            {
                try
                {
                    var rec = tr.GetObject(id, OpenMode.ForRead) as DimStyleTableRecord;
                    if (rec == null) continue;
                    var recName = (rec.Name ?? string.Empty).Trim().Replace('\u00A0', ' ');
                    if (string.Equals(recName, styleName, StringComparison.OrdinalIgnoreCase))
                    {
                        var recW = tr.GetObject(id, OpenMode.ForWrite) as DimStyleTableRecord;
                        if (recW != null)
                        {
                            ApplyDimStyleParameters(recW, layerColorIndex, textBoxScale, textStyleId);
                            AutoCadHelper.LogWithSafety($"DimStyleHelper: 找到近似样式 '{rec.Name}' 并已更新（请求名: '{styleName}'）。");
                            return id;
                        }
                    }
                }
                catch { /* 容错 */ }
            }

            try { dimStyleTable.UpgradeOpen(); }
            catch
            {
                dimStyleTable = tr.GetObject(db.DimStyleTableId, OpenMode.ForWrite) as DimStyleTable;
                if (dimStyleTable == null) return db.Dimstyle;
            }

            var newRec = new DimStyleTableRecord { Name = styleName };
            ApplyDimStyleParameters(newRec, layerColorIndex, textBoxScale, textStyleId);

            ObjectId newId = ObjectId.Null;
            try
            {
                newId = dimStyleTable.Add(newRec);
            }
            catch
            {
                // 保守回退：如果无法添加则返回当前数据库默认 dimstyle（原有行为）
                newId = db.Dimstyle;
            }

            return newId;
        }

        // 新增辅助：在活动文档的 Database 上创建或更新 DimStyle（使用标准 Transaction）
        // 这样可以保证样式写入到 AutoCAD 主窗口所显示的数据库。
        private static ObjectId EnsureOrCreateDimStyleInActiveDatabase(
            Database activeDb,
            string styleName,
            short layerColorIndex,
            double textBoxScale,
            ObjectId textStyleId)
        {
            if (activeDb == null) return ObjectId.Null;

            // 使用标准事务在活动数据库上操作
            using (var t = activeDb.TransactionManager.StartTransaction())
            {
                // 规范化样式名（再次保证）
                styleName = (styleName ?? string.Empty).Trim();
                styleName = styleName.Replace('\u00A0', ' ').Trim();
                if (string.IsNullOrEmpty(styleName)) return activeDb.Dimstyle;

                if (double.IsNaN(textBoxScale) || textBoxScale <= 0) textBoxScale = 1.0;
                if (textStyleId == ObjectId.Null) textStyleId = activeDb.Textstyle;

                var dimStyleTable = t.GetObject(activeDb.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
                if (dimStyleTable == null) return activeDb.Dimstyle;

                // 精确匹配
                if (dimStyleTable.Has(styleName))
                {
                    var existId = dimStyleTable[styleName];
                    var existRec = t.GetObject(existId, OpenMode.ForWrite) as DimStyleTableRecord;
                    if (existRec != null)
                    {
                        ApplyDimStyleParameters(existRec, layerColorIndex, textBoxScale, textStyleId);
                        t.Commit();
                        return existId;
                    }
                }

                // 大小写/不可见字符容错匹配
                foreach (ObjectId id in dimStyleTable)
                {
                    try
                    {
                        var rec = t.GetObject(id, OpenMode.ForRead) as DimStyleTableRecord;
                        if (rec == null) continue;
                        var recName = (rec.Name ?? string.Empty).Trim().Replace('\u00A0', ' ');
                        if (string.Equals(recName, styleName, StringComparison.OrdinalIgnoreCase))
                        {
                            var recW = t.GetObject(id, OpenMode.ForWrite) as DimStyleTableRecord;
                            if (recW != null)
                            {
                                ApplyDimStyleParameters(recW, layerColorIndex, textBoxScale, textStyleId);
                                AutoCadHelper.LogWithSafety($"DimStyleHelper: 在活动文档找到近似样式 '{rec.Name}' 并已更新（请求名: '{styleName}'）。");
                                t.Commit();
                                return id;
                            }
                        }
                    }
                    catch { /* 容错 */ }
                }

                // 创建新记录
                dimStyleTable = t.GetObject(activeDb.DimStyleTableId, OpenMode.ForWrite) as DimStyleTable;
                var newRec = new DimStyleTableRecord { Name = styleName };
                ApplyDimStyleParameters(newRec, layerColorIndex, textBoxScale, textStyleId);

                ObjectId newId = ObjectId.Null;
                try
                {
                    newId = dimStyleTable.Add(newRec);
                    t.AddNewlyCreatedDBObject(newRec, true); // 注册新对象到事务
                    t.Commit();
                    return newId;
                }
                catch (Exception ex)
                {
                    AutoCadHelper.LogWithSafety($"DimStyleHelper: 在活动文档添加样式 '{styleName}' 失败: {ex.Message}");
                    // 回退返回活动数据库当前 dimstyle
                    t.Abort();
                    return activeDb.Dimstyle;
                }
            }
        }

        /// <summary>
        /// 把关键的 DimStyle 参数统一应用到记录（复用函数，避免重复代码）
        /// 说明：本函数不会 Commit 或改变事务，仅在传入 DimStyleTableRecord 的写模式下调用。
        /// </summary>
        private static void ApplyDimStyleParameters(DimStyleTableRecord dimStyleRec, short layerColorIndex, double textBoxScale, ObjectId textStyleId)
        {
            if (dimStyleRec == null) return;

            // 基准真实显示高度（你要求的基准高度）
            const double baseTextHeight = 3.5;

            // 防护：确保 uiScale 合法
            if (double.IsNaN(textBoxScale) || textBoxScale <= 0) textBoxScale = 1.0;

            // 将 Dimtxt 存为：baseHeight * uiScale （在样式管理器中将显示“放大值”，例如 3.5 * 100 = 350）
            dimStyleRec.Dimtxt = Math.Max(MinTextSize, baseTextHeight * textBoxScale);

            // 文字/箭头/偏移等按比例写入样式（存为放大值）
            dimStyleRec.Dimasz = Math.Max(MinTextSize, 2.0 * textBoxScale);     // 箭头（放大值）
            dimStyleRec.Dimexo = Math.Max(0.1, 3.0 * textBoxScale);             // 界线偏移（放大值）
            dimStyleRec.Dimgap = Math.Max(0.1, 2.0 * textBoxScale);             // 文字与尺寸线间隙（放大值）

            // Dimlfac 保存注释缩放因子，使得实际显示高度 = Dimtxt * Dimlfac = baseTextHeight
            double dimlfac;
            try
            {
                dimlfac = (textBoxScale > 0.0) ? (1.0 / textBoxScale) : 1.0;
            }
            catch
            {
                dimlfac = 1.0;
            }
            if (double.IsNaN(dimlfac) || dimlfac <= 0.0) dimlfac = 1.0;
            dimStyleRec.Dimlfac = dimlfac;

            // 颜色设置（使用 ByAci 索引）
            try
            {
                dimStyleRec.Dimclrt = Color.FromColorIndex(ColorMethod.ByAci, layerColorIndex); // 文字颜色
                dimStyleRec.Dimclrd = Color.FromColorIndex(ColorMethod.ByAci, layerColorIndex); // 尺寸线颜色
                dimStyleRec.Dimclre = Color.FromColorIndex(ColorMethod.ByAci, layerColorIndex); // 界线颜色
            }
            catch
            {
                // 容错，不影响主要行为
            }

            // 文字样式与常用行为
            if (textStyleId != ObjectId.Null) dimStyleRec.Dimtxsty = textStyleId;
            dimStyleRec.Dimtad = 1;          // 文本垂直位置（常用）
            dimStyleRec.Dimtih = false;      // 水平标注时不强制水平
            dimStyleRec.Dimtoh = false;      // 竖直标注时不强制水平
            dimStyleRec.Dimzin = 0;          // 显示前后零
            dimStyleRec.Dimdec = 0;          // 小数位数

            // 保持标注为注释类型（如果 API 支持）
            try
            {
                dimStyleRec.Annotative = AnnotativeStates.True;
            }
            catch
            {
                // 某些 API 版本中可能不存在该属性，容错处理
            }
        }

        /// <summary>
        /// 辅助方法：在事务上下文中列出当前 DimStyle 表中所有样式名（用于调试，方便在命令行查看）
        /// </summary>
        public static List<string> ListDimStyleNames(DBTrans tr)
        {
            var list = new List<string>();
            if (tr == null) return list;

            try
            {
                var db = tr.Database;
                var dimStyleTable = tr.GetObject(db.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
                if (dimStyleTable == null) return list;

                foreach (ObjectId id in dimStyleTable)
                {
                    try
                    {
                        var rec = tr.GetObject(id, OpenMode.ForRead) as DimStyleTableRecord;
                        if (rec != null) list.Add(rec.Name ?? string.Empty);
                    }
                    catch { }
                }
            }
            catch { }

            return list;
        }

        // 新增：将数值向上取整到指定步长（例如步长50：101->150, 150->150, 151->200）
        // 放置位置：Command 类的私有辅助方法区域（靠近 TryGetDouble / EnsurePromptOk 等方法）
        public static long RoundUpToStep(double value, double step)
        {
            // 防御性处理：当 step <= 0 时回退为四舍五入的整数
            if (step <= 0) return Convert.ToInt64(Math.Round(value));
            // 使用 Math.Ceiling 做向上取整到最接近 step 的倍数
            return (long)(Math.Ceiling(value / step) * step);
        }

    }
}
