using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using System;
using System.Collections.Generic;
using GB_NewCadPlus_IV.Models;
using System.Linq;


namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 表格生成帮助类
    /// </summary>
    public static class CadTableHelper
    {
        /// <summary>
        /// 插入表格到当前图纸空间中方法
        /// </summary>
        /// <param name="doc">文档数据库</param>
        /// <param name="rows">计算表的所有行数据</param>
        /// <param name="insertionPoint">插入点</param>
        public static void InsertTable(Document doc, List<ParameterRow> rows, Point3d insertionPoint)
        {
            if (rows == null || rows.Count == 0) return;

            // 1. 获取绘图比例
            double scale = AutoCadHelper.GetScale();
            if (scale <= 0) scale = 1.0; // 防止无效比例

            // 打开文件锁
            using (DocumentLock docLock = doc.LockDocument())
            {
                Database db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    Table table = new Table(); // 新建表
                    int totalRows = rows.Count + 2; // 表行数再加2 (标题+表头+数据)
                    table.SetSize(totalRows, 5); // 设置表格大小
                    table.Position = insertionPoint; // 设置插入点
                    table.TableStyle = db.Tablestyle; // 设置表格样式

                    // 2. 计算并应用带比例的列宽
                    double[] colWidths = CalculateColumnWidths(rows, scale);
                    for (int i = 0; i < 5; i++)
                        table.SetColumnWidth(i, colWidths[i] * 0.8);    // 设置列宽

                    // 3. 应用比例到行高
                    // 原始基准高度: 标题14, 表头10, 数据7
                    table.SetRowHeight(0, 8 * scale);     //设置标题高度
                    table.SetRowHeight(1, 6 * scale);     // 设置表头高度
                    for (int i = 2; i < totalRows; i++)          // 从第2行开始设置行高
                        table.SetRowHeight(i, 5 * scale);  // 应用比例

                    // ---- 标题 ----
                    string title = $"{rows[0].Group}: 计算表";
                    for (int i = 0; i < 5; i++)
                    {
                        table.Cells[0, i].TextString = title;                                                          // 设置标题文本
                        table.Cells[0, i].TextHeight = 4 * scale;                                                      // 文字高度
                        table.Cells[0, i].Alignment = CellAlignment.MiddleCenter;                                      // 居中
                        table.Cells[0, i].BackgroundColor = Color.FromColorIndex(ColorMethod.ByColor, 140);  // 深灰色
                    }

                    // ---- 表头 ----
                    string[] headers = { "设备名称", "部件名称", "参数值", "单位", "备注" };
                    for (int i = 0; i < 5; i++)
                    {
                        table.Cells[1, i].TextString = headers[i]; // 设置表头文本
                        table.Cells[1, i].TextHeight = 3 * scale; // 应用比例
                        table.Cells[1, i].Alignment = CellAlignment.MiddleCenter; // 居中
                        table.Cells[1, i].BackgroundColor = Color.FromColorIndex(ColorMethod.ByColor, 253); // 浅灰色
                    }

                    // ---- 数据 ----
                    for (int i = 0; i < rows.Count; i++)
                    {
                        int rowIdx = i + 2;
                        var row = rows[i];

                        table.Cells[rowIdx, 0].TextString = row.Group ?? "";                // 设备名称
                        table.Cells[rowIdx, 1].TextString = row.Name ?? "";                 // 部件名称
                        table.Cells[rowIdx, 2].TextString = row.Value ?? "";                // 参数值
                        table.Cells[rowIdx, 3].TextString = row.Unit ?? "";                 // 单位
                        table.Cells[rowIdx, 4].TextString = row.Remark ?? "";               // 备注

                        table.Cells[rowIdx, 0].Alignment = CellAlignment.MiddleCenter;        // 居中
                        table.Cells[rowIdx, 1].Alignment = CellAlignment.MiddleCenter;        // 居中
                        table.Cells[rowIdx, 2].Alignment = CellAlignment.MiddleCenter;         // 右对齐
                        table.Cells[rowIdx, 3].Alignment = CellAlignment.MiddleLeft;          // 左对齐 
                        table.Cells[rowIdx, 4].Alignment = CellAlignment.MiddleCenter;        // 居中

                        for (int col = 0; col < 5; col++)
                            table.Cells[rowIdx, col].TextHeight = 2.5 * scale; // 文字高度

                        // 如果需要背景色逻辑，可以取消注释以下部分
                        // Color bgColor = GetBackgroundColor(row.Remark);
                        // for (int col = 0; col < 5; col++)
                        //    table.Cells[rowIdx, col].BackgroundColor = bgColor;
                    }

                    // ---- 视觉合并（清空非首行文本 + 隐藏边框） ----
                    // 注意：MergeFirstColumnVisually 内部可能也需要感知比例，如果它涉及具体的像素/单位操作
                    // 目前它只处理文本和背景色，受 TextHeight 影响较小，但保持一致性较好
                    MergeFirstColumnVisually(table, rows, scale);

                    btr.AppendEntity(table);
                    tr.AddNewlyCreatedDBObject(table, true);
                    Env.Editor.Regen();
                    tr.Commit();
                }
            }
        }


        /// <summary>
        /// 计算每列宽度（按内容最大等效字符数，并应用比例）
        /// 优化点：区分中英文宽度，汉字按2个字符宽度计算，更贴合AutoCAD显示效果
        /// </summary>
        public static double[] CalculateColumnWidths(List<ParameterRow> rows, double scale)
        {
            string[] headers = { "设备名称", "部件名称", "参数值", "单位", "备注" };

            // maxVisualChars 存储的是“视觉等效字符数” (汉字=2, 英文=1)
            double[] maxVisualChars = new double[5];

            // 1. 初始化表头的视觉宽度
            for (int i = 0; i < 5; i++)
            {
                maxVisualChars[i] = GetVisualCharCount(headers[i]);// 获取表头视觉字符数
            }

            // 2. 遍历所有数据行，找出每列的最大视觉宽度
            foreach (var row in rows)
            {
                maxVisualChars[0] = Math.Max(maxVisualChars[0], GetVisualCharCount(row.Group));          // 设备名称
                maxVisualChars[1] = Math.Max(maxVisualChars[1], GetVisualCharCount(row.Name));           // 部件名称
                maxVisualChars[2] = Math.Max(maxVisualChars[2], GetVisualCharCount(row.Value));          // 参数值
                maxVisualChars[3] = Math.Max(maxVisualChars[3], GetVisualCharCount(row.Unit));           // 单位
                maxVisualChars[4] = Math.Max(maxVisualChars[4], GetVisualCharCount(row.Remark));         // 备注
            }

            // 3. 将视觉字符数转换为实际宽度
            // baseCharWidth: 每个“半角字符”的基础宽度 (例如 4.0)
            // 如果一个汉字算2个视觉字符，那么它的宽度就是 4.0 * 2 = 8.0
            double baseCharWidth = 2.5;

            double[] widths = new double[5];
            for (int i = 0; i < 5; i++)
            {
                // 计算基础宽度：视觉字符数 * 单字符宽 + 固定内边距(10)
                double width = (maxVisualChars[i] * baseCharWidth) * scale;
                // 应用最小/最大限制 (随比例缩放)
                double minWidth = 15 * scale;
                double maxWidth = 500 * scale;
                if (width < minWidth)
                {
                    width = minWidth;
                    // 调试用：如果发现宽度被最小值限制，可以取消注释下面这行查看
                    // System.Diagnostics.Debug.WriteLine($"Col {i} width clamped to min: {minWidth}");
                }
                if (width > maxWidth)
                {
                    width = maxWidth;
                    // 调试用：如果发现宽度被最大值限制，可以取消注释下面这行查看
                    // System.Diagnostics.Debug.WriteLine($"Col {i} width clamped to max: {maxWidth}, original: {maxVisualChars[i] * baseCharWidth + 10}");
                }
                widths[i] = width;
            }
            // 【调试输出】可以在控制台看到每列计算出的最大视觉字符数，帮助排查问题
            // System.Diagnostics.Debug.WriteLine($"Max Visual Chars: [{string.Join(", ", maxVisualChars)}]");

            return widths;
        }

        /// <summary>
        /// 获取给定数据行所需的表格总宽度（包括列宽总和 + 边距）
        /// </summary>
        public static double GetTableWidth(List<ParameterRow> rows, double scale = 1.0)
        {
            if (rows == null || rows.Count == 0)
                return 0;

            double[] colWidths = CalculateColumnWidths(rows, scale); // 计算列宽
            double totalWidth = 0; // 总宽度
            foreach (double w in colWidths)
                totalWidth += w; // 累加列宽

            // 增加少量边距（与表格样式匹配，也需缩放）
            return totalWidth + (5 * scale);
        }

        /// <summary>
        /// 计算字符串的“视觉等效字符数”
        /// 汉字/全角字符计为 2，英文/数字/半角字符计为 1
        /// </summary>
        private static double GetVisualCharCount(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0;

            double count = 0;
            foreach (char c in text)
            {
                // Unicode 范围判断：
                // 基本汉字范围：\u4e00 - \u9fa5
                // 全角字符范围：\uff00 - \uffef (包含全角字母、数字、标点)
                if ((c >= '\u4e00' && c <= '\u9fa5') ||
                    (c >= '\uff00' && c <= '\uffef'))
                {
                    count += 2;
                }
                else
                {
                    count += 1;
                }
            }
            return count;
        }
        /// <summary>
        /// 视觉合并：清空非首行的 Group 文本，并将该单元格的边框颜色设为背景色（模拟隐藏边框）
        /// </summary>
        private static void MergeFirstColumnVisually(Table table, List<ParameterRow> rows, double scale)
        {
            if (rows == null || rows.Count == 0) return;

            int i = 0;
            while (i < rows.Count) // 遍历所有行
            {
                string currentGroup = rows[i].Group ?? "";                               // 当前行的 Group
                int startRow = i + 2;                                                    // 当前行在表格中的索引
                int j = i;                                                               // 下一行的索引
                while (j + 1 < rows.Count && (rows[j + 1].Group ?? "") == currentGroup)
                    j++;                                                                 // 找到当前 Group 的所有行
                int endRow = j + 2;                                                      // 当前 Group 最后一行的索引

                // 如果有多个连续行
                if (startRow < endRow)                                                   // 如果有多个连续行
                {
                    // 非首行的第一列清空文本
                    for (int k = startRow + 1; k <= endRow; k++)                         // 遍历当前 Group 的所有行
                    {
                        table.Cells[k, 0].TextString = "";                               // 清空文本

                        // 统一背景色以模拟合并效果
                        Color groupBgColor = table.Cells[startRow, 0].BackgroundColor;
                        table.Cells[k, 0].BackgroundColor = groupBgColor;                // 统一背景色以模拟合并效果
                    }
                }
                i = j + 1;
            }
        }

        private static Color GetBackgroundColor(string remark)
        {
            if (string.IsNullOrEmpty(remark)) return Color.FromColorIndex(ColorMethod.ByColor, 255);
            if (remark.Contains("输入")) return Color.FromColorIndex(ColorMethod.ByColor, 255);
            if (remark.Contains("计算")) return Color.FromColorIndex(ColorMethod.ByColor, 253);
            if (remark.Contains("推荐")) return Color.FromColorIndex(ColorMethod.ByColor, 230);
            return Color.FromColorIndex(ColorMethod.ByColor, 255);
        }


        /// <summary>
        /// 【新增】计算下一个表格的插入点
        /// 用于在多个表格横向排列时，避免重叠
        /// </summary>
        /// <param name="currentInsertionPoint">当前表格的插入点</param>
        /// <param name="rows">当前表格的数据行</param>
        /// <param name="spacing">额外的间距（模型空间单位），默认为0，建议传入如 10*scale</param>
        /// <returns>下一个表格的插入点</returns>
        public static Point3d GetNextInsertionPoint(Point3d currentInsertionPoint, List<ParameterRow> rows, double spacing = 0)
        {
            if (rows == null || rows.Count == 0)
                return currentInsertionPoint;

            double scale = 1.0;
            try
            {
                scale = AutoCadHelper.GetScale();
                if (scale <= 0) scale = 1.0;
            }
            catch { }

            // 获取当前表格的实际宽度
            double currentTableWidth = GetTableWidth(rows, scale);

            // 如果没有指定额外间距，使用默认间距系数
            if (spacing == 0)
            {
                spacing = 10 * scale; // 默认间距为 10 * 比例
            }

            // 新的 X 坐标 = 当前 X + 表格宽度 + 间距
            double nextX = currentInsertionPoint.X + currentTableWidth + spacing;

            return new Point3d(nextX, currentInsertionPoint.Y, currentInsertionPoint.Z);
        }
    }
}


//namespace GB_NewCadPlus_IV.Helpers
//{
//    public static class CadTableHelper
//    {
//        /// <summary>
//        /// 插入表格到当前图纸空间中方法
//        /// </summary>
//        /// <param name="doc">文档数据库</param>
//        /// <param name="rows">行</param>
//        /// <param name="insertionPoint">插入点</param>
//        public static void InsertTable(Document doc, List<ParameterRow> rows, Point3d insertionPoint)
//        {
//            if (rows == null || rows.Count == 0) return;

//            var scale= AutoCadHelper.GetScale();
//            // 打开文件锁
//            using (DocumentLock docLock = doc.LockDocument())
//            {
//                Database db = doc.Database;
//                using (Transaction tr = db.TransactionManager.StartTransaction())
//                {
//                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
//                    BlockTableRecord btr =
//                        tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

//                    Table table = new Table(); // 新建表
//                    int totalRows = rows.Count + 2; // 表行数再加2
//                    table.SetSize(totalRows, 5); // 设置表格大小
//                    table.Position = insertionPoint; // 设置插入点
//                    table.TableStyle = db.Tablestyle; // 设置表格样式

//                    // 列宽
//                    double[] colWidths = CalculateColumnWidths(rows);
//                    for (int i = 0; i < 5; i++)
//                        table.SetColumnWidth(i, colWidths[i]);

//                    // 行高
//                    table.SetRowHeight(0, 14);
//                    table.SetRowHeight(1, 10);
//                    for (int i = 2; i < totalRows; i++)
//                        table.SetRowHeight(i, 7);

//                    // ---- 标题 ----
//                    //string title = $"设备计算表  生成日期: {DateTime.Now:yyyy-MM-dd HH:mm}";
//                    string title = "设备计算表";
//                    for (int i = 0; i < 5; i++)
//                    {
//                        table.Cells[0, i].TextString = title;
//                        table.Cells[0, i].TextHeight = 7;
//                        table.Cells[0, i].Alignment = CellAlignment.MiddleCenter;
//                        table.Cells[0, i].BackgroundColor = Color.FromColorIndex(ColorMethod.ByColor, 140);
//                    }

//                    // ---- 表头 ----
//                    string[] headers = { "设备名称", "部件名称", "参数值", "单位", "备注" };
//                    for (int i = 0; i < 5; i++)
//                    {
//                        table.Cells[1, i].TextString = headers[i];
//                        table.Cells[1, i].TextHeight = 5;
//                        table.Cells[1, i].Alignment = CellAlignment.MiddleCenter;
//                        table.Cells[1, i].BackgroundColor = Color.FromColorIndex(ColorMethod.ByColor, 253);
//                    }

//                    // ---- 数据 ----
//                    for (int i = 0; i < rows.Count; i++)
//                    {
//                        int rowIdx = i + 2;
//                        var row = rows[i];

//                        table.Cells[rowIdx, 0].TextString = row.Group ?? "";
//                        table.Cells[rowIdx, 1].TextString = row.Name ?? "";
//                        table.Cells[rowIdx, 2].TextString = row.Value ?? "";
//                        table.Cells[rowIdx, 3].TextString = row.Unit ?? "";
//                        table.Cells[rowIdx, 4].TextString = row.Remark ?? "";

//                        table.Cells[rowIdx, 0].Alignment = CellAlignment.MiddleLeft;
//                        table.Cells[rowIdx, 1].Alignment = CellAlignment.MiddleLeft;
//                        table.Cells[rowIdx, 2].Alignment = CellAlignment.MiddleRight;
//                        table.Cells[rowIdx, 3].Alignment = CellAlignment.MiddleRight;
//                        table.Cells[rowIdx, 4].Alignment = CellAlignment.MiddleCenter;

//                        for (int col = 0; col < 5; col++)
//                            table.Cells[rowIdx, col].TextHeight = 4.5;

//                        //Color bgColor = GetBackgroundColor(row.Remark);
//                        //for (int col = 0; col < 5; col++)
//                        //    table.Cells[rowIdx, col].BackgroundColor = bgColor;
//                    }

//                    // ---- 视觉合并（清空非首行文本 + 隐藏边框） ----
//                    MergeFirstColumnVisually(table, rows);

//                    btr.AppendEntity(table);
//                    tr.AddNewlyCreatedDBObject(table, true);
//                    tr.Commit();
//                }
//            }
//        }

//        /// <summary>
//        /// 视觉合并：清空非首行的 Group 文本，并将该单元格的边框颜色设为背景色（模拟隐藏边框）
//        /// </summary>
//        private static void MergeFirstColumnVisually(Table table, List<ParameterRow> rows)
//        {
//            if (rows == null || rows.Count == 0) return;

//            int i = 0;
//            while (i < rows.Count)
//            {
//                string currentGroup = rows[i].Group ?? "";
//                int startRow = i + 2;
//                int j = i;
//                while (j + 1 < rows.Count && (rows[j + 1].Group ?? "") == currentGroup)
//                    j++;
//                int endRow = j + 2;

//                // 如果有多个连续行
//                if (startRow < endRow)
//                {
//                    // 非首行的第一列清空文本
//                    for (int k = startRow + 1; k <= endRow; k++)
//                    {
//                        table.Cells[k, 0].TextString = "";
//                        // ★ 将边框颜色设为背景色，隐藏分隔线（仅支持设置背景色，无法单独设置边框）
//                        // 注意：这里无法单独设置单元格边框，只能通过设置单元格背景色和表格整体边框来弥补
//                        // 为了更接近合并效果，我们让非首行第一列的背景色与首行保持一致
//                        // 背景色已经在数据填充时根据备注设置，但可能同一Group下不同行的备注不同，导致颜色不同
//                        // 这里强制统一为第一行的背景色
//                        Color groupBgColor = table.Cells[startRow, 0].BackgroundColor;
//                        table.Cells[k, 0].BackgroundColor = groupBgColor;
//                    }
//                }
//                i = j + 1;
//            }
//        }

//        private static Color GetBackgroundColor(string remark)
//        {
//            if (string.IsNullOrEmpty(remark)) return Color.FromColorIndex(ColorMethod.ByColor, 255);
//            if (remark.Contains("输入")) return Color.FromColorIndex(ColorMethod.ByColor, 255);
//            if (remark.Contains("计算")) return Color.FromColorIndex(ColorMethod.ByColor, 253);
//            if (remark.Contains("推荐")) return Color.FromColorIndex(ColorMethod.ByColor, 230);
//            return Color.FromColorIndex(ColorMethod.ByColor, 255);
//        }

//        /// <summary>
//        /// 计算每列宽度（按内容最大字符数）
//        /// </summary>
//        public static double[] CalculateColumnWidths(List<ParameterRow> rows)
//        {
//            string[] headers = { "设备名称", "部件名称", "参数值", "单位", "备注" };
//            double[] maxChars = new double[5];
//            for (int i = 0; i < 5; i++)
//                maxChars[i] = headers[i].Length;

//            foreach (var row in rows)
//            {
//                maxChars[0] = Math.Max(maxChars[0], row.Group?.Length ?? 0);
//                maxChars[1] = Math.Max(maxChars[1], row.Name?.Length ?? 0);
//                maxChars[2] = Math.Max(maxChars[2], row.Value?.Length ?? 0);
//                maxChars[3] = Math.Max(maxChars[3], row.Unit?.Length ?? 0);
//                maxChars[4] = Math.Max(maxChars[4], row.Remark?.Length ?? 0);
//            }

//            double charWidth = 4.0;
//            double[] widths = new double[5];
//            for (int i = 0; i < 5; i++)
//            {
//                widths[i] = maxChars[i] * charWidth + 10;
//                if (widths[i] < 35) widths[i] = 35;
//                if (widths[i] > 200) widths[i] = 200;
//            }
//            return widths;
//        }
//        /// <summary>
//        /// 获取给定数据行所需的表格总宽度（包括列宽总和 + 边距）
//        /// </summary>
//        public static double GetTableWidth(List<ParameterRow> rows)
//        {
//            if (rows == null || rows.Count == 0)
//                return 0;

//            double[] colWidths = CalculateColumnWidths(rows);
//            double totalWidth = 0;
//            foreach (double w in colWidths)
//                totalWidth += w;

//            // 增加少量边距（与表格样式匹配）
//            return totalWidth + 5;
//        }
//    }
//}

//namespace GB_NewCadPlus_IV.Helpers
//{
//    public static class CadTableHelper
//    {
//        /// <summary>
//        /// 插入计算表到当前图纸的模型空间（兼容所有AutoCAD版本）
//        /// </summary>
//        public static void InsertTable(Document doc, List<ParameterRow> rows, Point3d insertionPoint)
//        {
//            if (rows == null || rows.Count == 0) return;
//            using (DocumentLock docLock = doc.LockDocument())
//            {
//                Database db = doc.Database;
//                using (Transaction tr = db.TransactionManager.StartTransaction())
//                {
//                    // 获取模型空间
//                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
//                    BlockTableRecord btr =
//                        tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

//                    // 创建表格
//                    Table table = new Table();
//                    int totalRows = rows.Count + 2; // 标题行 + 表头行 + 数据行
//                    table.SetSize(totalRows, 5);
//                    table.Position = insertionPoint;

//                    // 使用当前数据库的表格样式（确保有默认边框）
//                    table.TableStyle = db.Tablestyle;

//                    // ---- 设置列宽 ----
//                    double[] colWidths = CalculateColumnWidths(rows);
//                    for (int i = 0; i < 5; i++)
//                        table.SetColumnWidth(i, colWidths[i]);

//                    // ---- 设置行高 ----
//                    table.SetRowHeight(0, 14); // 标题行
//                    table.SetRowHeight(1, 10); // 表头行
//                    for (int i = 2; i < totalRows; i++)
//                        table.SetRowHeight(i, 7);

//                    // ============================================================
//                    //  1. 标题行（不合并，直接填充所有列）
//                    // ============================================================
//                    string title = $"设备计算表  生成日期: {DateTime.Now:yyyy-MM-dd HH:mm}";
//                    for (int i = 0; i < 5; i++)
//                    {
//                        table.Cells[0, i].TextString = title;
//                        table.Cells[0, i].TextHeight = 7;
//                        table.Cells[0, i].Alignment = CellAlignment.MiddleCenter;
//                        table.Cells[0, i].BackgroundColor = Color.FromColorIndex(ColorMethod.ByColor, 140); // 浅蓝色
//                    }

//                    // ============================================================
//                    //  2. 表头行（第1行）
//                    // ============================================================
//                    string[] headers = { "设备名称", "部件名称", "参数值", "单位", "备注" };
//                    for (int i = 0; i < 5; i++)
//                    {
//                        table.Cells[1, i].TextString = headers[i];
//                        table.Cells[1, i].TextHeight = 5;
//                        table.Cells[1, i].Alignment = CellAlignment.MiddleCenter;
//                        table.Cells[1, i].BackgroundColor = Color.FromColorIndex(ColorMethod.ByColor, 253); // 浅灰色
//                    }

//                    // ============================================================
//                    //  3. 数据行（从第2行开始）
//                    // ============================================================
//                    for (int i = 0; i < rows.Count; i++)
//                    {
//                        int rowIdx = i + 2;
//                        var row = rows[i];

//                        // 填充数据
//                        table.Cells[rowIdx, 0].TextString = row.Group ?? "";
//                        table.Cells[rowIdx, 1].TextString = row.Name ?? "";
//                        table.Cells[rowIdx, 2].TextString = row.Value ?? "";
//                        table.Cells[rowIdx, 3].TextString = row.Unit ?? "";
//                        table.Cells[rowIdx, 4].TextString = row.Remark ?? "";

//                        // 设置对齐方式
//                        table.Cells[rowIdx, 0].Alignment = CellAlignment.MiddleLeft;
//                        table.Cells[rowIdx, 1].Alignment = CellAlignment.MiddleLeft;
//                        table.Cells[rowIdx, 2].Alignment = CellAlignment.MiddleRight;
//                        table.Cells[rowIdx, 3].Alignment = CellAlignment.MiddleRight;
//                        table.Cells[rowIdx, 4].Alignment = CellAlignment.MiddleCenter;

//                        // 设置字体大小
//                        for (int col = 0; col < 5; col++)
//                            table.Cells[rowIdx, col].TextHeight = 4.5;

//                        // 根据备注设置背景色
//                        //Color bgColor = GetBackgroundColor(row.Remark);
//                        //for (int col = 0; col < 5; col++)
//                        //    table.Cells[rowIdx, col].BackgroundColor = bgColor;
//                    }

//                    // ---- 将表格添加到模型空间 ----
//                    btr.AppendEntity(table);
//                    tr.AddNewlyCreatedDBObject(table, true);
//                    tr.Commit();
//                }
//            }
//        }

//        /// <summary>
//        /// 根据备注获取对应的背景色
//        /// </summary>
//        private static Color GetBackgroundColor(string remark)
//        {
//            if (string.IsNullOrEmpty(remark))
//                return Color.FromColorIndex(ColorMethod.ByColor, 7); // 白色

//            if (remark.Contains("输入"))
//                return Color.FromColorIndex(ColorMethod.ByColor, 7); // 白色

//            if (remark.Contains("计算"))
//                return Color.FromColorIndex(ColorMethod.ByColor, 253); // 浅灰色

//            if (remark.Contains("推荐"))
//                return Color.FromColorIndex(ColorMethod.ByColor, 230); // 浅黄色

//            return Color.FromColorIndex(ColorMethod.ByColor, 7); // 默认白色
//        }

//        /// <summary>
//        /// 计算每列宽度（根据内容长度自适应）
//        /// </summary>
//        private static double[] CalculateColumnWidths(List<ParameterRow> rows)
//        {
//            // 初始化最大字符数（包含表头）
//            double[] maxChars = { 8, 10, 12, 6, 6 };
//            string[] headers = { "设备名称", "部件名称", "参数值", "单位", "备注" };
//            for (int i = 0; i < headers.Length; i++)
//                maxChars[i] = Math.Max(maxChars[i], headers[i].Length);

//            // 统计每列最大字符数
//            foreach (var row in rows)
//            {
//                maxChars[0] = Math.Max(maxChars[0], row.Group?.Length ?? 0);
//                maxChars[1] = Math.Max(maxChars[1], row.Name?.Length ?? 0);
//                maxChars[2] = Math.Max(maxChars[2], row.Value?.Length ?? 0);
//                maxChars[3] = Math.Max(maxChars[3], row.Unit?.Length ?? 0);
//                maxChars[4] = Math.Max(maxChars[4], row.Remark?.Length ?? 0);
//            }

//            // 每个字符宽度约 4mm（根据字体大小调整）
//            double charWidth = 4.0;
//            double[] widths = new double[5];
//            for (int i = 0; i < 5; i++)
//            {
//                widths[i] = maxChars[i] * charWidth + 8; // 加内边距
//                // 设置最小宽度
//                if (widths[i] < 30) widths[i] = 30;
//                // 设置最大宽度（防止过宽）
//                if (widths[i] > 180) widths[i] = 180;
//            }

//            return widths;
//        }
//    }
//}
