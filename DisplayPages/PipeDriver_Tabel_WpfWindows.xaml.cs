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
using OfficeOpenXml;
using DataTable = System.Data.DataTable;
using DataColumn = System.Data.DataColumn;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog; // 引用 EPPlus

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
            // 设置 EPPlus 许可证（如果尚未在全局设置）
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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
                    using (var package = new ExcelPackage(new System.IO.FileInfo(openFileDialog.FileName)))
                    {
                        if (package.Workbook.Worksheets.Count == 0)
                        {
                            MessageBox.Show("Excel 文件中没有工作表。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }

                        var worksheet = package.Workbook.Worksheets[0]; // 读取第一个 Sheet
                        var dataTable = new DataTable();

                        // 加载 Excel 到 DataTable (第一行作为表头)
                        // 注意：ToDataTable 是 EPPlus 的扩展方法，确保引用了正确的命名空间
                        dataTable = worksheet.Cells[worksheet.Dimension.Address].ToDataTable();

                        _currentTableData = dataTable;
                        DataGrid_PipeDriver.ItemsSource = _currentTableData.DefaultView;

                        StatusText.Text = $"已加载: {System.IO.Path.GetFileName(openFileDialog.FileName)} ({dataTable.Rows.Count} 行)";
                    }
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
                double scaleDenom = GB_NewCadPlus_IV.Helpers.AutoCadHelper.GetScale();

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
                var saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                saveDialog.FileName = $"管道明细表_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                if (saveDialog.ShowDialog() == true)
                {                    
                    using (var package = new ExcelPackage())
                    {
                        var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                        // --- 第1行：大标题 (合并所有列) ---
                        string title = TextBlock_TableTitle.Text; // 获取界面标题
                        int totalCols = _dataTable.Columns.Count;
                        worksheet.Cells[1, 1, 1, totalCols].Merge = true;
                        worksheet.Cells[1, 1].Value = title;
                        worksheet.Cells[1, 1].Style.Font.Bold = true;
                        worksheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        // --- 第2-3行：表头 (中文/英文) ---
                        // 定义固定列的中英文对照 (根据 CreateDeviceTableWithType 逻辑)
                        var headerMap = new Dictionary<string, string>
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

                        for (int col = 0; col < _dataTable.Columns.Count; col++)
                        {
                            string cnName = _dataTable.Columns[col].ColumnName;
                            string enName = headerMap.ContainsKey(cnName) ? headerMap[cnName] : cnName; // 如果没有映射，就用原名

                            // 第2行：中文
                            worksheet.Cells[2, col + 1].Value = cnName;
                            // 第3行：英文
                            worksheet.Cells[3, col + 1].Value = enName;

                            // 设置表头样式
                            worksheet.Cells[2, col + 1].Style.Font.Bold = true;
                            worksheet.Cells[3, col + 1].Style.Font.Bold = true;
                            worksheet.Cells[2, col + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheet.Cells[3, col + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        }

                        // --- 第4行起：数据 ---
                        // LoadFromDataTable 从指定单元格开始加载，不覆盖表头
                        worksheet.Cells[4, 1].LoadFromDataTable(_dataTable, false); // false 表示不加载列头，因为我们要自己画

                        // 自动调整列宽
                        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                        // 添加边框
                        using (var range = worksheet.Cells[1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns])
                        {
                            range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }

                        package.SaveAs(new System.IO.FileInfo(saveDialog.FileName));
                    }
                    StatusText.Text = $"已导出到: {saveDialog.FileName}";
                    MessageBox.Show("导出成功！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
