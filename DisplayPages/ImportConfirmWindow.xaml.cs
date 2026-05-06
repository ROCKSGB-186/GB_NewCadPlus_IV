using GB_NewCadPlus_IV.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
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
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Clipboard = System.Windows.Clipboard;
using Path = System.IO.Path;
using MessageBox = System.Windows.MessageBox;
using DataTable = Autodesk.AutoCAD.DatabaseServices.DataTable;
using Cursors = System.Windows.Input.Cursors;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.FunctionalMethod;

namespace GB_NewCadPlus_IV.Views
{

    /// <summary>
    /// 导入确认窗口交互逻辑
    /// </summary>
    public partial class ImportConfirmWindow : Window
    {
        // P/Invoke: 释放 GDI 对象（供 WinForms Bitmap -> BitmapSource 转换后释放 HBITMAP）
        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool DeleteObject(IntPtr hObject);
        private readonly ImportEntityDto _dto;// 导入的实体数据
        private readonly WpfMainWindow _mainWindow; // 引用主窗口以访问方法

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="dto">导入实体数据传输对象</param>
        /// <param name="mainWindow">主窗口</param>
        /// <exception cref="ArgumentNullException"></exception>
        public ImportConfirmWindow(ImportEntityDto dto, WpfMainWindow mainWindow)
        {
            InitializeComponent();
            _dto = dto ?? throw new ArgumentNullException(nameof(dto));// 参数不能为空 主窗口引用
            // 初始化控件 绑定数据源引用 
            _mainWindow = mainWindow ?? throw new ArgumentNullException(nameof(mainWindow));// 参数不能为空
            // 加载预览图片 窗口加载时绑定数据
            this.Loaded += (s, e) =>
            {
                // 加载预览图片 加载预览图
                LoadPreviewImage();
                // 绑定数据源 绑定属性到数据网格 绑定属性到数据网格
                BindPropertiesToGrid();
            };

            BtnConfirm.Click += BtnConfirm_Click;//确认按钮点击事件
            BtnCancel.Click += (s, e) => this.DialogResult = false;//取消按钮点击事件
            BtnPastePreview.Click += BtnPastePreview_Click;//粘贴预览按钮点击事件
            BtnExportTemplate.Click += BtnExportTemplate_Click;//导出模板按钮点击事件
        }

        /// <summary>
        /// 加载预览图片
        /// </summary>
        private void LoadPreviewImage()
        {
            try
            {
                // 统一走候选路径解析，避免只依赖单一字段导致预览不显示
                string previewPath = ResolvePreviewImagePath();

                // 如果最终路径为空或文件不存在，则清空图片并记录日志
                if (string.IsNullOrWhiteSpace(previewPath))
                {
                    PreviewImage.Source = null;
                    LogManager.Instance.LogWarning($"预览图未找到。DTO路径: {_dto?.PreviewImagePath}，FileStorage路径: {_dto?.FileStorage?.PreviewImagePath}，File路径: {_dto?.FileStorage?.FilePath}");
                    return;
                }

                if (!System.IO.File.Exists(previewPath))
                {
                    PreviewImage.Source = null;
                    LogManager.Instance.LogWarning($"预览图文件不存在: {previewPath}");
                    return;
                }

                // 使用文件流只读打开并允许共享读取，避免图片文件被占用时无法显示
                using (var fs = new FileStream(previewPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.StreamSource = fs;
                    bitmap.EndInit();
                    bitmap.Freeze();
                    PreviewImage.Source = bitmap;
                }

                // 确保图片控件可见并刷新布局
                PreviewImage.Visibility = System.Windows.Visibility.Visible;
                PreviewImage.InvalidateVisual();
                PreviewImage.UpdateLayout();

                LogManager.Instance.LogInfo($"已加载并显示预览图: {previewPath}");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"加载预览图失败: {ex.Message}");
                PreviewImage.Source = null;
            }
        }

        /// <summary>
        /// 解析预览图路径，按 DTO 顶层字段、FileStorage 字段、文件同目录命名顺序兜底。
        /// </summary>
        private string ResolvePreviewImagePath()
        {
            string[] candidates =
            {
                _dto?.PreviewImagePath,
                _dto?.FileStorage?.PreviewImagePath,
                BuildPreviewPathFromFileStorage(_dto?.FileStorage)
            };

            foreach (var candidate in candidates)
            {
                if (!string.IsNullOrWhiteSpace(candidate) && System.IO.File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// 当数据库只保存文件名而非完整预览路径时，尝试基于主文件目录拼出预览图路径。
        /// </summary>
        private static string BuildPreviewPathFromFileStorage(FileStorage fileStorage)
        {
            if (fileStorage == null)
            {
                return string.Empty;
            }

            if (string.IsNullOrWhiteSpace(fileStorage.PreviewImageName))
            {
                return string.Empty;
            }

            var baseDir = string.Empty;
            if (!string.IsNullOrWhiteSpace(fileStorage.PreviewImagePath))
            {
                baseDir = Path.GetDirectoryName(fileStorage.PreviewImagePath) ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(baseDir) && !string.IsNullOrWhiteSpace(fileStorage.FilePath))
            {
                baseDir = Path.GetDirectoryName(fileStorage.FilePath) ?? string.Empty;
            }

            return string.IsNullOrWhiteSpace(baseDir)
                ? string.Empty
                : Path.Combine(baseDir, fileStorage.PreviewImageName);
        }

        /// <summary>
        /// 绑定属性到数据网格
        /// </summary>
        private void BindPropertiesToGrid()
        {
            // 优先使用 DTO 中的 JSON 属性字典来生成展示数据
            var displayData = _mainWindow.PrepareFileDisplayData(_dto.FileStorage, _dto.AttributesJson);
            PropertiesGrid.ItemsSource = displayData;
        }

        /// <summary>
        /// 粘贴预览剪贴板图片
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnPastePreview_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "选择预览图片",
                    Filter = "图片文件 (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif|所有文件|*.*",
                    Multiselect = false
                };

                if (dlg.ShowDialog() != true)
                {
                    return;
                }

                string selectedFile = dlg.FileName;
                if (!File.Exists(selectedFile))
                {
                    MessageBox.Show("选定的文件不存在。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                try
                {
                    string tempDir = Path.Combine(Path.GetTempPath(), "GB_NewCadPlus_IV_Previews");
                    Directory.CreateDirectory(tempDir);
                    string ext = Path.GetExtension(selectedFile);
                    string newPreviewPath = Path.Combine(tempDir, $"preview_uploaded_{Guid.NewGuid()}{ext}");

                    // 复制文件到临时位置，避免后续文件被移动/删除导致丢失
                    File.Copy(selectedFile, newPreviewPath);

                    _dto.PreviewImagePath = newPreviewPath;
                    LogManager.Instance.LogInfo($"预览图已从文件选择保存到: {newPreviewPath}");

                    // 使用已存在的安全加载方法刷新预览
                    Dispatcher.BeginInvoke((Action)(() => LoadPreviewImage()), System.Windows.Threading.DispatcherPriority.Render);
                }
                catch (Exception exSave)
                {
                    LogManager.Instance.LogError($"保存上传预览图失败: {exSave.Message}");
                    MessageBox.Show($"保存预览图片失败: {exSave.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"上传预览图处理失败: {ex.Message}");
                MessageBox.Show($"操作失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 粘贴剪贴板图片
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnExportTemplate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 先把界面编辑值回写到 DTO
                UpdateDtoFromGrid();

                // 生成模板表
                var dt = _mainWindow.CreateTemplateDataTable();
                dt.Rows.Clear();
                DataRow newRow = dt.NewRow();

                // 先填充 FileStorage 固定字段
                foreach (PropertyInfo prop in typeof(FileStorage).GetProperties())
                {
                    if (dt.Columns.Contains(prop.Name))
                    {
                        newRow[prop.Name] = prop.GetValue(_dto.FileStorage) ?? DBNull.Value;
                    }
                }

                // 构建 JSON 属性字典（优先 DTO 里的 Attributes 字典；没有则从旧 FileAttribute 桥接）
                var attrDict = BuildExportAttributesDictionary();

                // 动态列导出——模板里没有的字段自动补列
                foreach (var kv in attrDict)
                {
                    if (string.IsNullOrWhiteSpace(kv.Key)) continue;

                    // 如果模板不存在该列，则动态新增
                    if (!dt.Columns.Contains(kv.Key))
                    {
                        dt.Columns.Add(kv.Key, typeof(string));
                    }

                    // 写入字段值
                    newRow[kv.Key] = kv.Value ?? string.Empty;
                }

                // 加入一行
                dt.Rows.Add(newRow);

                // 导出文件
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel 文件 (*.xlsx)|*.xlsx",
                    FileName = $"图元_{_dto.FileStorage.DisplayName}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    if (_mainWindow.ExportDataTableToExcel(dt, saveFileDialog.FileName))
                    {
                        MessageBox.Show("模板导出成功！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出模板失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 构建导出用属性字典（JSON化）
        /// </summary>
        private Dictionary<string, string> BuildExportAttributesDictionary()
        {
            // 优先返回 DTO 中最新的 JSON 属性字典副本
            if (_dto != null && _dto.AttributesJson != null && _dto.AttributesJson.Count > 0)
            {
                return new Dictionary<string, string>(_dto.AttributesJson, StringComparer.OrdinalIgnoreCase);
            }

            // 兜底返回空字典，避免空引用
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 在 ImportConfirmWindow 类中新增字段（靠近其它私有字段）
        /// </summary>
        private bool _isConfirmProcessing = false;

        /// <summary>
        /// 确认按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnConfirm_Click(object sender, RoutedEventArgs e)
        {
            if (_isConfirmProcessing) // 防止重复触发
                return;

            _isConfirmProcessing = true;
            BtnConfirm.IsEnabled = false;
            var prevCursor = Mouse.OverrideCursor;
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                // 1. 从UI更新DTO
                UpdateDtoFromGrid(); // 确保先保存用户在表格中修改的属性

                // 2. 提示用户是否关闭当前文件
                var result = MessageBox.Show("是否关闭当前文件？\n关闭后可正常上传，不关闭可能导致文件被占用。", "关闭文件提示", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        if (doc != null)
                        {
                            // 调用 AutoCAD API 需小心：捕获可能的异常
                            doc.CloseAndSave(doc.Name);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"关闭文件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        // 失败也可以继续，让用户决定
                    }
                }
                else
                {
                    var giveUp = MessageBox.Show("是否放弃本次识别并退出？", "放弃识别", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    if (giveUp == MessageBoxResult.Yes)
                    {
                        this.DialogResult = false;
                    }
                    return;
                }

                // 3. 再次从UI更新DTO，确保所有更改都已保存
                UpdateDtoFromGrid();

                // 4. 执行导入（将 DTO 注册到主窗口并执行上传）
                _mainWindow.SetSelectedFileForImport(_dto);

                try
                {
                    await _mainWindow.UploadFileAndSaveToDatabase(_dto);
                    // 5. 关闭窗口
                    this.DialogResult = true;
                }
                catch (Exception exUp)
                {
                    MessageBox.Show($"导入失败: {exUp.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    this.DialogResult = false;
                }
            }
            finally
            {
                _isConfirmProcessing = false;
                BtnConfirm.IsEnabled = true;
                Mouse.OverrideCursor = prevCursor;
            }
        }

        /// <summary>
        /// 从UI的DataGrid中读取修改后的值，并更新回DTO对象
        /// </summary>
        private void UpdateDtoFromGrid()
        {
            var items = PropertiesGrid.ItemsSource as List<CategoryPropertyEditModel>;
            if (items == null) return;

            // 确保 JSON 属性字典已初始化
            _dto.AttributesJson ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 每次重新采集前先清空，避免旧值残留
            _dto.AttributesJson.Clear();

            // 把界面显示名统一映射为系统内部标准键名，避免后续 JSON 键混乱
            string NormalizeKey(string key)
            {
                if (string.IsNullOrWhiteSpace(key)) return string.Empty;

                switch (key.Trim())
                {
                    case "文件名": return "FileName";
                    case "显示名称": return "DisplayName";
                    case "元素块名": return "BlockName";
                    case "图层名称": return "LayerName";
                    case "层名": return "LayerName";
                    case "颜色索引": return "ColorIndex";
                    case "比例": return "Scale";
                    case "长度": return "Length";
                    case "宽度": return "Width";
                    case "高度": return "Height";
                    case "角度": return "Angle";
                    case "基点X": return "BasePointX";
                    case "基点Y": return "BasePointY";
                    case "基点Z": return "BasePointZ";
                    case "介质": return "MediumName";
                    case "规格": return "Specifications";
                    case "材质": return "Material";
                    case "标准号": return "StandardNumber";
                    case "功率": return "Power";
                    case "容积": return "Volume";
                    case "压力": return "Pressure";
                    case "温度": return "Temperature";
                    case "直径": return "Diameter";
                    case "外径": return "OuterDiameter";
                    case "内径": return "InnerDiameter";
                    case "厚度": return "Thickness";
                    case "重量": return "Weight";
                    case "型号": return "Model";
                    case "备注": return "Remarks";
                    case "自定义1": return "Customize1";
                    case "自定义2": return "Customize2";
                    case "自定义3": return "Customize3";
                    default: return key.Trim();
                }
            }

            // 局部函数，安全写入 JSON 属性字典
            void AddAttr(string key, string value)
            {
                if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(value)) return;

                // 统一键名后再写入
                var normalizedKey = NormalizeKey(key);
                if (string.IsNullOrWhiteSpace(normalizedKey)) return;

                _dto.AttributesJson[normalizedKey] = value.Trim();
            }

            foreach (var item in items)
            {
                // 先回写 FileStorage 固定字段
                _mainWindow.SetFileStorageProperty(_dto.FileStorage, item.PropertyName1, item.PropertyValue1);
                _mainWindow.SetFileStorageProperty(_dto.FileStorage, item.PropertyName2, item.PropertyValue2);

                // 再把两列属性写回 JSON 字典
                AddAttr(item.PropertyName1, item.PropertyValue1);
                AddAttr(item.PropertyName2, item.PropertyValue2);
            }

            // 补充主表关键字段，保证 JSON 中也有一份稳定数据
            AddAttr("FileName", _dto.FileStorage.FileName ?? string.Empty);
            AddAttr("DisplayName", _dto.FileStorage.DisplayName ?? string.Empty);
            AddAttr("BlockName", _dto.FileStorage.BlockName ?? string.Empty);
            AddAttr("LayerName", _dto.FileStorage.LayerName ?? string.Empty);
            AddAttr("ColorIndex", _dto.FileStorage.ColorIndex?.ToString() ?? string.Empty);
            AddAttr("Scale", _dto.FileStorage.Scale?.ToString() ?? string.Empty);
            AddAttr("UpdatedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        }

    }
}
