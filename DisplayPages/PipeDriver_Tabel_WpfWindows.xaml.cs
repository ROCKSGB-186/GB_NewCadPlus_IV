using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Data;
using Autodesk.AutoCAD.ApplicationServices;
using GB_NewCadPlus_IV.UniFiedStandards; // 引用 DeviceInfo
using Microsoft.Win32; // 用于文件对话框
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.IO;
using DataTable = System.Data.DataTable;
using DataColumn = System.Data.DataColumn;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using BorderStyle = NPOI.SS.UserModel.BorderStyle;
using Path = System.IO.Path; // 引用 EPPlus

namespace GB_NewCadPlus_IV.DisplayPages
{
    /// <summary>
    /// PipeDriver_Tabel_WpfWindows.xaml 的交互逻辑
    /// </summary>
    public partial class PipeDriver_Tabel_WpfWindows : Window
    {
        /// <summary>
        /// 内部存储原始数据，用于转换
        /// </summary>
        private List<DeviceInfo> _sourceData;
        /// <summary>
        /// DataTable 用于绑定 DataGrid，方便编辑和动态列
        /// </summary>
        private DataTable _dataTable;
        /// <summary>
        /// 存储当前表格数据，用于同步
        /// </summary>
        private DataTable _currentTableData;
       
        public PipeDriver_Tabel_WpfWindows()
        {
            InitializeComponent();
            
        }
        /// <summary>
        /// 初始化窗口，传入数据
        /// </summary>
        public void InitializeWithData(List<DeviceInfo> deviceList, string title)
        {
            this.Title = $"{title} - 编辑器";

            // 【新增】更新界面显示的大标题
            if (TextBlock_TableTitle != null)
            {
                TextBlock_TableTitle.Text = title;
            }

            _sourceData = deviceList;
            ConvertToDataTable();
            DataGrid_PipeDriver.ItemsSource = _dataTable.DefaultView;
            StatusText.Text = $"已加载 {_sourceData.Count} 条记录";
        }


        private void Btn_加载管道设备表_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            openFileDialog.Title = "选择要加载的管道/设备表";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    IWorkbook workbook;
                    using (FileStream fs = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        workbook = new XSSFWorkbook(fs);   // 读取 .xlsx
                    }

                    if (workbook.NumberOfSheets == 0)
                    {
                        MessageBox.Show("Excel 文件中没有工作表。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    ISheet sheet = workbook.GetSheetAt(0);   // 第一个 Sheet
                    if (sheet == null || sheet.LastRowNum < 0)
                    {
                        MessageBox.Show("工作表为空。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // 构建 DataTable，以第一行作为列名
                    DataTable dataTable = new DataTable();
                    DataFormatter formatter = new DataFormatter();

                    IRow headerRow = sheet.GetRow(0);
                    if (headerRow == null)
                    {
                        MessageBox.Show("表头行不存在。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // 创建列，以首行内容为 ColumnName
                    int colCount = headerRow.LastCellNum; // 1‑based
                    for (int c = 0; c < colCount; c++)
                    {
                        ICell cell = headerRow.GetCell(c);
                        string colName = cell != null ? formatter.FormatCellValue(cell).Trim() : $"Column{c + 1}";
                        if (string.IsNullOrWhiteSpace(colName)) colName = $"Column{c + 1}";
                        dataTable.Columns.Add(colName);
                    }

                    // 填充数据行（从第2行开始）
                    int rowCount = sheet.LastRowNum; // 0‑based
                    for (int r = 1; r <= rowCount; r++)
                    {
                        IRow row = sheet.GetRow(r);
                        if (row == null) continue;

                        DataRow dr = dataTable.NewRow();
                        for (int c = 0; c < colCount; c++)
                        {
                            ICell cell = row.GetCell(c);
                            string cellValue = cell != null ? formatter.FormatCellValue(cell) : string.Empty;
                            dr[c] = cellValue;
                        }
                        dataTable.Rows.Add(dr);
                    }

                    // 绑定到 UI
                    _currentTableData = dataTable;
                    DataGrid_PipeDriver.ItemsSource = _currentTableData.DefaultView;

                    StatusText.Text = $"已加载: {System.IO.Path.GetFileName(openFileDialog.FileName)} ({dataTable.Rows.Count} 行)";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"加载失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        /// <summary>
        /// 同步到 CAD 图元的按钮点击事件处理程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_同步到cad图元_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTableData == null || _currentTableData.Rows.Count == 0)
            {
                MessageBox.Show("表格中没有数据，请先加载 Excel。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("未找到活动的 CAD 文档。请确保 CAD 已启动。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // 1. 将 DataTable 转换为 List<DeviceInfo>
                var deviceList = ConvertDataTableToDeviceInfoList(_currentTableData);

                if (deviceList.Count == 0)
                {
                    MessageBox.Show("未能从表格中提取有效数据。请检查表头是否包含'管段号'或'Name'列。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 2. 【关键修改】将数据暂存到新的静态变量中
                UnifiedTableGenerator.TempExcelSyncData = deviceList;

                // 3. 隐藏当前 WPF 窗口
                this.Hide();

                // 4. 【关键修改】激活 CAD 窗口并执行新的同步命令
                doc.SendStringToExecute("_.SyncExcelDataToCad ", true, false, true);

                StatusText.Text = "已切换到 CAD，请在命令行提示下选择要更新的图元...";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"同步准备失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Show(); // 出错则重新显示窗口
            }
        }
        /// <summary>
        /// 管道设备表插入到cad中
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_管道设备表插入到cad_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. 获取编辑后的数据
                var editedData = ConvertFromDataTable();

                // 2. 调用 CAD 插入逻辑
                // 注意：这里需要在 UI 线程外执行 CAD 操作，或者使用 CommandMethod 触发
                // 为了简单，我们触发一个命令，或者直接在当前上下文调用（如果允许）
                // 假设我们在 WPF 中，需要通过 Application.DocumentManager 获取当前文档

                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("未找到活动的 CAD 文档。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 获取比例分母 (可以从主窗口传递过来，或者重新获取)
                double scaleDenom = VariableDictionary.wpfTextBoxScale;

                // 实例化生成器
                var generator = new UnifiedTableGenerator();

                // 由于 CreateDeviceTableWithType 是 void 且需要用户指定插入点，它会阻塞 UI 直到用户点击
                // 这在 WPF 中可能会导致界面假死，建议改为异步或提示用户
                doc.Editor.WriteMessage("\n请在 CAD 中指定表格插入位置...");

                // 调用插入方法
                generator.CreateDeviceTableWithType(doc.Database, editedData, this.Title.Replace(" - 编辑器", ""), scaleDenom);

                StatusText.Text = "表格已插入 CAD";
                MessageBox.Show("表格已成功插入到 CAD 中！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// 导出 Excel 的按钮点击事件处理程序
        /// 注意：这里需要重建 CAD 表格的复杂表头格式
        /// </summary>
        private void Btn_导出EXCEL_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_dataTable == null || _dataTable.Columns.Count == 0)
                {
                    MessageBox.Show("当前没有数据可以导出。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excel 文件 (*.xlsx)|*.xlsx",
                    FileName = $"管道明细表_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveDialog.ShowDialog() != true) return;

                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Sheet1");

                // ---------- 清除任何可能存在的初始合并（新 sheet 通常没有，但以防万一） ----------
                RemoveAllMergedRegions(sheet);

                // ---------- 创建样式 (省略部分重复代码，只保留关键结构) ----------
                // 标题样式
                ICellStyle titleStyle = workbook.CreateCellStyle();
                titleStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                titleStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                IFont titleFont = workbook.CreateFont();
                titleFont.IsBold = true;
                titleFont.FontHeightInPoints = 14;
                titleStyle.SetFont(titleFont);
                SetBorderThin(titleStyle);

                // 中文表头样式
                ICellStyle headerCnStyle = workbook.CreateCellStyle();
                headerCnStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                headerCnStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                IFont headerCnFont = workbook.CreateFont();
                headerCnFont.IsBold = true;
                headerCnFont.FontHeightInPoints = 11;
                headerCnStyle.SetFont(headerCnFont);
                SetBorderThin(headerCnStyle);

                // 英文表头样式
                ICellStyle headerEnStyle = workbook.CreateCellStyle();
                headerEnStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                headerEnStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                IFont headerEnFont = workbook.CreateFont();
                headerEnFont.IsBold = true;
                headerEnFont.FontHeightInPoints = 10;
                headerEnFont.IsItalic = true;
                headerEnStyle.SetFont(headerEnFont);
                SetBorderThin(headerEnStyle);

                // 数据样式
                ICellStyle dataStyle = workbook.CreateCellStyle();
                dataStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                dataStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                SetBorderThin(dataStyle);

                string title = TextBlock_TableTitle?.Text ?? "管道明细表";
                int totalCols = _dataTable.Columns.Count;

                // ---------- 第1行：标题（先写数据，再尝试合并） ----------
                IRow titleRow = sheet.CreateRow(0);
                for (int c = 0; c < totalCols; c++)
                {
                    ICell cell = titleRow.CreateCell(c);
                    cell.CellStyle = titleStyle;
                    if (c == 0) cell.SetCellValue(title);
                }

                // 尝试合并标题行，如果失败则保留未合并状态（至少文字居中且边框完整）
                if (totalCols > 1)
                {
                    try
                    {
                        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, totalCols - 1));
                    }
                    catch (Exception exMerge)
                    {
                        // 记录合并失败，但不影响后续数据输出
                        System.Diagnostics.Debug.WriteLine($"标题合并失败: {exMerge.Message}");
                    }
                }

                // ---------- 第2行：中文表头 ----------
                IRow headerCnRow = sheet.CreateRow(1);
                // ---------- 第3行：英文表头 ----------
                IRow headerEnRow = sheet.CreateRow(2);

                var headerMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "管道标题", "Pipe Title" },
            { "管段号", "Pipe No." },
            { "起点", "Start Point" },
            { "终点", "End Point" },
            { "管道等级", "Pipe Class" },
            { "介质名称", "Medium" },
            { "操作温度", "Op. Temp" },
            { "操作压力", "Op. Press" },
            { "隔热隔声代号", "Insulation Code" },
            { "是否防腐", "Anti-Corrosion" },
            { "名称", "Name" },
            { "材料", "Material" },
            { "图号或标准号", "DWG/STD No." },
            { "数量", "Qty" },
            { "泵前/后", "Pump Side" }
        };

                for (int col = 0; col < totalCols; col++)
                {
                    string cnName = _dataTable.Columns[col].ColumnName;
                    string enName = headerMap.TryGetValue(cnName, out var mapped) ? mapped : cnName;

                    ICell cnCell = headerCnRow.CreateCell(col);
                    cnCell.SetCellValue(cnName);
                    cnCell.CellStyle = headerCnStyle;

                    ICell enCell = headerEnRow.CreateCell(col);
                    enCell.SetCellValue(enName);
                    enCell.CellStyle = headerEnStyle;
                }

                // ---------- 从第4行开始填充数据 ----------
                for (int i = 0; i < _dataTable.Rows.Count; i++)
                {
                    IRow dataRow = sheet.CreateRow(3 + i);
                    for (int col = 0; col < totalCols; col++)
                    {
                        ICell cell = dataRow.CreateCell(col);
                        string value = _dataTable.Rows[i][col]?.ToString() ?? string.Empty;
                        cell.SetCellValue(value);
                        cell.CellStyle = dataStyle;
                    }
                }

                // ---------- 列宽 ----------
                for (int col = 0; col < totalCols; col++)
                {
                    double maxLen = GetVisualCharWidth(title);
                    string cnH = _dataTable.Columns[col].ColumnName;
                    string enH = headerMap.TryGetValue(cnH, out var m) ? m : cnH;
                    maxLen = Math.Max(maxLen, GetVisualCharWidth(cnH));
                    maxLen = Math.Max(maxLen, GetVisualCharWidth(enH));
                    int checkRows = Math.Min(200, _dataTable.Rows.Count);
                    for (int r = 0; r < checkRows; r++)
                    {
                        string txt = _dataTable.Rows[r][col]?.ToString() ?? string.Empty;
                        maxLen = Math.Max(maxLen, GetVisualCharWidth(txt));
                    }
                    int colWidth = (int)(maxLen * 256 * 1.15) + 512;
                    colWidth = Math.Max(2048, Math.Min(colWidth, 255 * 256));
                    sheet.SetColumnWidth(col, colWidth);
                }

                // ---------- 保存 ----------
                string dir = Path.GetDirectoryName(saveDialog.FileName);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                using (FileStream fs = new FileStream(saveDialog.FileName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs, false);
                }

                StatusText.Text = $"已导出到: {saveDialog.FileName}";
                MessageBox.Show("导出成功！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}\n\n堆栈:\n{ex.StackTrace}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 辅助：彻底删除所有合并区域
        /// </summary>
        /// <param name="sheet"></param>
        private void RemoveAllMergedRegions(ISheet sheet)
        {
            while (sheet.NumMergedRegions > 0)
            {
                sheet.RemoveMergedRegion(sheet.NumMergedRegions - 1); // 倒序删除，索引0-based
            }
        }

        /// <summary>
        /// 辅助：设置细边框
        /// </summary>
        /// <param name="style"></param>
        private void SetBorderThin(ICellStyle style)
        {
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
        }

        /// <summary>
        /// 辅助：文字视觉宽度
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private double GetVisualCharWidth(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0;
            double w = 0;
            foreach (char c in text)
            {
                if (c >= 0x4e00 && c <= 0x9fff || c >= 0x3000 && c <= 0x303f || c >= 0xff00 && c <= 0xffef)
                    w += 2.0;
                else
                    w += 1.0;
            }
            return w;
        }

        /// <summary>
        /// 将 List<DeviceInfo> 转换为具有固定列顺序的 DataTable
        /// 此顺序与 CAD 中 CreateDeviceTableWithType 生成的表格顺序保持一致
        /// </summary>
        private void ConvertToDataTable()
        {
            _dataTable = new DataTable();

            if (_sourceData == null || _sourceData.Count == 0) return;

            // ================== 第一步：定义严格的列顺序 ==================

            // 1. 固定列 (对应 CAD 表格的前10列)
            var fixedColumns = new List<string>
            {
                "管道标题",     // Index 0
                "管段号",       // Index 1
                "起点",         // Index 2
                "终点",         // Index 3
                "管道等级",     // Index 4
                "介质名称",     // Index 5
                "操作温度",     // Index 6
                "操作压力",     // Index 7
                "隔热隔声代号", // Index 8
                "是否防腐"      // Index 9
            };

            // 2. 首选动态列 (对应 CAD 表格中紧随其后的特定属性)
            var preferredDynamicColumns = new List<string>
            {
                "名称",
                "材料",
                "图号或标准号",
                "数量",
                "泵前/后"
            };

            // 3. 收集其余所有动态列
            var otherDynamicColumnsSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var device in _sourceData)
            {
                if (device.Attributes != null)
                {
                    foreach (var key in device.Attributes.Keys)
                    {
                        // 排除已经定义在固定列和首选列中的键
                        if (!fixedColumns.Contains(key) && !preferredDynamicColumns.Contains(key))
                        {
                            otherDynamicColumnsSet.Add(key);
                        }
                    }
                }
            }
            // 将其余列按字母排序，保证每次打开界面顺序一致
            var otherDynamicColumns = otherDynamicColumnsSet.OrderBy(c => c).ToList();

            // 4. 合并最终列顺序
            var finalColumnOrder = new List<string>();
            finalColumnOrder.AddRange(fixedColumns);
            finalColumnOrder.AddRange(preferredDynamicColumns);
            finalColumnOrder.AddRange(otherDynamicColumns);

            // ================== 第二步：创建 DataTable 列 ==================

            foreach (var colName in finalColumnOrder)
            {
                // 防止重复添加（虽然逻辑上不会重复，但为了安全）
                if (!_dataTable.Columns.Contains(colName))
                {
                    _dataTable.Columns.Add(colName, typeof(string));
                }
            }

            // ================== 第三步：填充数据行 ==================

            foreach (var device in _sourceData)
            {
                DataRow row = _dataTable.NewRow();

                // 遍历定义好的列顺序进行赋值
                foreach (var colName in finalColumnOrder)
                {
                    string value = string.Empty;

                    // --- 特殊逻辑处理 ---

                    // 1. 管道标题 (Index 0)
                    if (colName == "管道标题")
                    {
                        // 优先取属性
                        if (device.Attributes != null && device.Attributes.TryGetValue("管道标题", out var titleVal))
                        {
                            value = titleVal;
                        }
                        // 其次从 Name 截取 (模拟 CAD 逻辑: PIPE_123 -> 123)
                        else if (!string.IsNullOrWhiteSpace(device.Name))
                        {
                            int pos = device.Name.LastIndexOf('_');
                            value = (pos >= 0 && pos < device.Name.Length - 1) ? device.Name.Substring(pos + 1) : device.Name;
                        }
                    }
                    // 2. 数量 (如果 Attributes 里没有，使用 DeviceInfo.Count)
                    else if (colName == "数量")
                    {
                        if (device.Attributes != null && device.Attributes.TryGetValue("数量", out var qVal))
                        {
                            value = qVal;
                        }
                        else
                        {
                            value = device.Count.ToString();
                        }
                    }
                    // 3. 普通属性列
                    else
                    {
                        if (device.Attributes != null && device.Attributes.TryGetValue(colName, out var attrVal))
                        {
                            value = attrVal;
                        }
                    }

                    // 赋值给 DataRow
                    row[colName] = value;
                }

                _dataTable.Rows.Add(row);
            }
        }

        /// <summary>
        /// 从 DataTable 还原为 List<DeviceInfo>
        /// </summary>
        private List<DeviceInfo> ConvertFromDataTable()
        {
            var newList = new List<DeviceInfo>();
            if (_dataTable == null) return newList;

            foreach (DataRow row in _dataTable.Rows)
            {
                // 使用完整命名空间，避免与 DatabaseManager.DeviceInfo 冲突
                var device = new GB_NewCadPlus_IV.UniFiedStandards.DeviceInfo();
                device.Attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                // 遍历 DataTable 的所有列
                foreach (DataColumn col in _dataTable.Columns)
                {
                    var val = row[col.ColumnName]?.ToString();

                    // 跳过空值
                    if (string.IsNullOrWhiteSpace(val)) continue;

                    string columnName = col.ColumnName;

                    // --- 特殊列处理 ---

                    if (columnName == "管道标题")
                    {
                        device.Attributes["管道标题"] = val;
                        // 可选：如果需要，也可以更新 Name，但通常同步只关注 Attributes
                    }
                    else if (columnName == "管段号")
                    {
                        device.Attributes["管段号"] = val;
                        // 管段号通常作为匹配键，同时也存入 Name 以便兼容
                        if (string.IsNullOrWhiteSpace(device.Name))
                            device.Name = val;
                    }
                    else if (columnName == "数量")
                    {
                        if (int.TryParse(val, out int count))
                            device.Count = count;

                        device.Attributes["数量"] = val;
                    }
                    else if (columnName == "名称")
                    {
                        // 如果表格里有"名称"列，且 Name 为空，则赋值
                        if (string.IsNullOrWhiteSpace(device.Name))
                            device.Name = val;

                        device.Attributes["名称"] = val;
                    }
                    else
                    {
                        // 其他所有列都放入 Attributes
                        // 这样 CAD 同步命令 SyncExcelDataToCadEntities 就能通过 Tag 匹配到
                        device.Attributes[columnName] = val;
                    }
                }

                // 兜底：如果 Name 还是空的，尝试用管段号
                if (string.IsNullOrWhiteSpace(device.Name))
                {
                    if (device.Attributes.TryGetValue("管段号", out var pn))
                        device.Name = pn;
                    else
                        device.Name = "Unknown_Pipe";
                }

                // 默认类型
                if (string.IsNullOrWhiteSpace(device.Type))
                    device.Type = "管道";

                newList.Add(device);
            }
            return newList;
        }

        /// <summary>
        /// 辅助方法：DataTable 转 List<DeviceInfo>
        /// </summary>
        private List<DeviceInfo> ConvertDataTableToDeviceInfoList(DataTable dt)
        {
            var newList = new List<DeviceInfo>();
            if (_dataTable == null) return newList;

            foreach (DataRow row in _dataTable.Rows)
            {
                var device = new GB_NewCadPlus_IV.UniFiedStandards.DeviceInfo(); // 使用完整命名空间避免冲突
                device.Attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                // 遍历 DataTable 的所有列
                foreach (DataColumn col in _dataTable.Columns)
                {
                    var val = row[col.ColumnName]?.ToString();

                    // 跳过空值
                    if (string.IsNullOrWhiteSpace(val)) continue;

                    // 特殊列处理
                    if (col.ColumnName == "管道标题")
                    {
                        // 可以选择存入 Attributes["管道标题"]，或者更新 Name
                        device.Attributes["管道标题"] = val;
                        // 如果需要，也可以更新 Name 以便后续匹配
                        // device.Name = val; 
                    }
                    else if (col.ColumnName == "管段号")
                    {
                        device.Attributes["管段号"] = val;
                        // 管段号通常也作为 Name 用于匹配
                        if (string.IsNullOrWhiteSpace(device.Name)) device.Name = val;
                    }
                    else if (col.ColumnName == "数量")
                    {
                        if (int.TryParse(val, out int count))
                            device.Count = count;
                        device.Attributes["数量"] = val;
                    }
                    else
                    {
                        // 其他所有列都放入 Attributes
                        device.Attributes[col.ColumnName] = val;
                    }
                }

                // 如果 Name 还是空的，尝试用管段号或第一列填充
                if (string.IsNullOrWhiteSpace(device.Name))
                {
                    if (device.Attributes.TryGetValue("管段号", out var pn)) device.Name = pn;
                    else device.Name = "Unknown";
                }

                newList.Add(device);
            }
            return newList;
        }
    }
}
