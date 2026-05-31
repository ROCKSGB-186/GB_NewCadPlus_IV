using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Helpers;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
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
using System.Windows.Threading;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Clipboard = System.Windows.Clipboard;
using Cursor = System.Windows.Input.Cursor;
using Cursors = System.Windows.Input.Cursors;
using DataTable = Autodesk.AutoCAD.DatabaseServices.DataTable;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Path = System.IO.Path;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;

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

            this.BtnConfirm.Click += new RoutedEventHandler(this.BtnConfirm_Click); // 确认按钮点击事件
            this.BtnCancel.Click += (RoutedEventHandler)((s, e) => this.CloseDialogWithResult(false)); // 取消按钮直接关闭窗口并返回 false
            this.BtnPastePreview.Click += new RoutedEventHandler(this.BtnPastePreview_Click); // 粘贴预览按钮点击事件
            this.BtnExportTemplate.Click += new RoutedEventHandler(this.BtnExportTemplate_Click); // 导出模板按钮点击事件
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
            if (fileStorage == null || string.IsNullOrWhiteSpace(fileStorage.PreviewImageName)) // 没有预览图文件名，无法构建路径
                return string.Empty;
            string path1 = string.Empty; // 优先使用预览图路径的目录部分，如果没有再使用主文件路径的目录部分
            if (!string.IsNullOrWhiteSpace(fileStorage.PreviewImagePath)) // 如果预览图路径存在，使用其目录部分
                path1 = Path.GetDirectoryName(fileStorage.PreviewImagePath) ?? string.Empty; // 如果预览图路径没有目录部分，则尝试使用主文件路径的目录部分
            if (string.IsNullOrWhiteSpace(path1) && !string.IsNullOrWhiteSpace(fileStorage.FilePath)) // 如果预览图路径没有目录部分且主文件路径存在，使用主文件路径的目录部分
                path1 = Path.GetDirectoryName(fileStorage.FilePath) ?? string.Empty; // 最终如果 path1 仍然为空，则无法构建预览图路径，返回空字符串
            return string.IsNullOrWhiteSpace(path1) ? string.Empty : Path.Combine(path1, fileStorage.PreviewImageName); // 拼接目录部分和预览图文件名，得到完整的预览图路径
        }

        /// <summary>
        /// 绑定属性到数据网格
        /// </summary>
        private void BindPropertiesToGrid()
        {
            // 诊断日志
            if (_dto.AttributesJson == null)
                LogManager.Instance.LogWarning("BindPropertiesToGrid: _dto.AttributesJson 为 null");
            else
                LogManager.Instance.LogInfo($"BindPropertiesToGrid: AttributesJson 条目数 = {_dto.AttributesJson.Count}");
            // 准备显示数据时捕获异常并记录日志，避免因单条数据问题导致整个绑定失败
            var displayData = _mainWindow.PrepareFileDisplayData(_dto.FileStorage, _dto.AttributesJson);
            if (displayData == null || !displayData.Any())
                LogManager.Instance.LogWarning("PrepareFileDisplayData 返回空集合");

            PropertiesGrid.ItemsSource = null; // 先清空绑定，确保 UI 能正确刷新，避免因旧数据残留导致的显示异常
            
            PropertiesGrid.ItemsSource = displayData;// 绑定数据源到数据网格，确保 UI 能正确显示属性列表，避免因绑定问题导致的显示异常
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
                OpenFileDialog openFileDialog1 = new OpenFileDialog();// 创建文件选择对话框
                openFileDialog1.Title = "选择预览图片"; // 设置对话框标题
                openFileDialog1.Filter = "图片文件 (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif|所有文件|*.*"; // 设置文件过滤器，限制只能选择图片文件
                openFileDialog1.Multiselect = false; // 禁止多选，确保一次只能选择一个文件
                OpenFileDialog openFileDialog2 = openFileDialog1; // 复制一份对话框实例，避免直接使用 openFileDialog1 导致潜在的状态问题
                if (!openFileDialog2.ShowDialog().GetValueOrDefault()) // 显示对话框并检查用户是否选择了文件，如果没有选择则直接返回
                    return;
                string fileName = openFileDialog2.FileName; // 获取用户选择的文件路径
                if (!File.Exists(fileName)) // 再次验证文件是否存在，避免用户选择后文件被删除或移动导致的错误
                {
                    // 文件不存在，显示错误消息并记录日志
                    int num1 = (int)MessageBox.Show("选定的文件不存在。", "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
                }
                else
                {
                    try
                    {
                        string str1 = Path.Combine(Path.GetTempPath(), "GB_NewCadPlus_IV_Previews"); // 构建临时目录路径，专门用于存放预览图，避免与其他临时文件混淆
                        Directory.CreateDirectory(str1); // 确保临时目录存在，如果已存在则不会覆盖，避免潜在的权限问题或数据丢失
                        string str2 = Path.GetExtension(fileName); // 获取原文件的扩展名，保持预览图的格式一致，避免因格式不支持导致无法显示
                        string destFileName = Path.Combine(str1, $"preview_uploaded_{Guid.NewGuid()}{str2}"); // 构建目标文件路径，使用 GUID 确保文件名唯一，避免多次上传导致的文件覆盖问题
                        File.Copy(fileName, destFileName); // 复制文件到目标路径，使用 File.Copy 而非 File.Move 保持原文件不变，避免用户误操作导致数据丢失
                        this._dto.PreviewImagePath = destFileName; // 更新 DTO 中的预览图路径，确保后续加载预览图时能正确找到新文件
                        this._dto.PreviewImageName = fileName; // 同时更新预览图文件名字段，保持 DTO 数据的一致性，避免后续使用时因字段不一致导致的问题
                        this._dto.FileStorage.PreviewImagePath = destFileName; // 更新 DTO 中的预览图路径，确保后续加载预览图时能正确找到新文件
                        this._dto.FileStorage.PreviewImageName = fileName; // 同时更新预览图文件名字段，保持 DTO 数据的一致性，避免后续使用时因字段不一致导致的问题
                        LogManager.Instance.LogInfo("预览图已从文件选择保存到: " + destFileName); // 成功保存预览图后，立即加载显示新预览图，提供即时反馈
                        this.Dispatcher.BeginInvoke((Delegate)(() => this.LoadPreviewImage()), DispatcherPriority.Render); // 使用 Dispatcher 调度加载预览图，确保 UI 线程安全地更新图片显示，避免因直接调用 LoadPreviewImage 导致的线程问题
                        BindPropertiesToGrid();// 绑定数据源 绑定属性到数据网格 绑定属性到数据网格
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogError("保存上传预览图失败: " + ex.Message);
                        int num2 = (int)MessageBox.Show("保存预览图片失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError("上传预览图处理失败: " + ex.Message);
                int num = (int)MessageBox.Show("操作失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
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
                this.UpdateDtoFromGrid(); // 确保先保存用户在表格中修改的属性，避免导出模板时数据不完整
                var templateDataTable = this._mainWindow.CreateTemplateDataTable(); // 从主窗口获取预定义的模板 DataTable 结构，确保列定义一致
                templateDataTable.Rows.Clear(); // 清空任何现有数据，确保导出时只有当前图元的数据，避免混入旧数据
                DataRow row = templateDataTable.NewRow(); // 创建新行用于填充当前图元的数据
                foreach (PropertyInfo property in typeof(FileStorage).GetProperties()) // 反射获取 FileStorage 类的所有属性，动态填充数据，避免硬编码字段导致的维护问题
                {
                    if (templateDataTable.Columns.Contains(property.Name)) // 仅填充模板中定义的列，避免因 DTO 中新增字段导致的列不匹配问题
                        row[property.Name] = property.GetValue((object)this._dto.FileStorage) ?? (object)DBNull.Value; // 获取属性值，如果为 null 则使用 DBNull.Value 填充，确保 Excel 中显示为空而非 "null" 字符串
                }
                // 额外添加 DTO 中 JSON 属性字典的键值对，允许用户在模板中看到并修改这些动态属性，增强模板的灵活性和适用性
                foreach (KeyValuePair<string, string> exportAttributes in this.BuildExportAttributesDictionary())
                {
                    if (!string.IsNullOrWhiteSpace(exportAttributes.Key)) // 仅添加键非空的属性，避免因空键导致的列名问题
                    {
                        if (!templateDataTable.Columns.Contains(exportAttributes.Key)) // 如果模板中没有定义该列，则动态添加列，允许用户在模板中看到并修改这些动态属性，增强模板的灵活性和适用性
                            templateDataTable.Columns.Add(exportAttributes.Key, typeof(string)); // 添加新列时默认类型为 string，确保 Excel 中显示正确，避免因类型不匹配导致的显示问题
                        row[exportAttributes.Key] = (object)(exportAttributes.Value ?? string.Empty); // 填充属性值，如果为 null 则使用空字符串填充，确保 Excel 中显示为空而非 "null" 字符串
                    }
                }
                // 将填充好的行添加到模板 DataTable 中，准备导出，确保导出的 Excel 文件包含当前图元的完整数据，满足用户的定制化需求
                templateDataTable.Rows.Add(row);
                // 创建并配置保存文件对话框，允许用户选择导出文件的保存位置和名称，增强用户体验，确保导出的文件符合用户的期望
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Excel 文件 (*.xlsx)|*.xlsx";
                saveFileDialog1.FileName = $"图元_{this._dto.FileStorage.DisplayName}.xlsx";
                SaveFileDialog saveFileDialog2 = saveFileDialog1; // 复制一份对话框实例，避免直接使用 saveFileDialog1 导致潜在的状态问题
                if (!saveFileDialog2.ShowDialog().GetValueOrDefault() || !this._mainWindow.ExportDataTableToExcel(templateDataTable, saveFileDialog2.FileName))
                    return;
                int num = (int)MessageBox.Show("模板导出成功！", "成功", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("导出模板失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        /// <summary>
        /// 构建导出用属性字典（JSON化）
        /// </summary>
        private Dictionary<string, string> BuildExportAttributesDictionary()
        {
            // 构建一个新的字典，优先使用 DTO 中的 JSON 属性字典，如果 DTO 中没有则返回一个空字典，确保导出模板时有一个稳定的数据来源，避免因 DTO 中缺失属性导致的导出失败问题
            return this._dto != null && this._dto.AttributesJson != null && this._dto.AttributesJson.Count > 0 ? new Dictionary<string, string>((IDictionary<string, string>)this._dto.AttributesJson, (IEqualityComparer<string>)StringComparer.OrdinalIgnoreCase) : new Dictionary<string, string>((IEqualityComparer<string>)StringComparer.OrdinalIgnoreCase);

        }

        /// <summary>
        /// 在 ImportConfirmWindow 类中新增字段（靠近其它私有字段）
        /// </summary>
        private bool _isConfirmProcessing = false;

        /// <summary>
        /// 安全关闭对话框并返回结果。
        /// </summary>
        private void CloseDialogWithResult(bool result)
        {
            try
            {
                this.DialogResult = new bool?(result);// 设置 DialogResult 以便主窗口能正确接收结果
            }
            catch (InvalidOperationException ex)
            {
                this.Close(); // 如果设置 DialogResult 失败（如窗口未以 ShowDialog 方式打开），则直接关闭窗口，确保用户操作得到响应
            }
        }

        /// <summary>
        /// 确认按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnConfirm_Click(object sender, RoutedEventArgs e)
        {
            Cursor prevCursor; // 保存先前光标，操作完成后恢复
            if (this._isConfirmProcessing)
            {
                prevCursor = (Cursor)null; // 如果正在处理，则避免重复执行
            }
            else
            {
                this._isConfirmProcessing = true; // 标记正在处理，防止重复点击
                this.BtnConfirm.IsEnabled = false; // 禁用按钮以防多次触发
                prevCursor = Mouse.OverrideCursor; // 记录当前光标
                Mouse.OverrideCursor = Cursors.Wait; // 显示等待光标，提示用户操作中
                try
                {
                    this.UpdateDtoFromGrid(); // 先把 UI 修改回写到 DTO
                    MessageBoxResult result = MessageBox.Show("是否关闭当前文件？\n关闭后可正常上传，不关闭可能导致文件被占用。", "关闭文件提示", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                    {
                        try
                        {
                            // 获取当前活动文档
                            Document doc = Application.DocumentManager.MdiActiveDocument;
                            // 直接使用 null 检查，避免显式调用 operator 方法导致编译错误
                            if (doc != null)
                            {
                                // 关闭并保存当前文档，确保上传时文件不被占用
                                DocumentExtension.CloseAndSave(doc, doc.Name);
                            }
                            // 将引用置空，避免后续误用已关闭的文档对象
                            doc = (Document)null;
                        }
                        catch (Exception ex)
                        {
                            // 关闭失败也要给出明确提示
                            int num = (int)MessageBox.Show("关闭文件失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
                        }
                        //this.UpdateDtoFromGrid(); // 关闭后再次同步 UI 到 DTO，保证上传数据一致
                        this._mainWindow.SetSelectedFileForImport(this._dto); // 设置要导入的文件信息到主窗口
                        try
                        {
                            // 执行上传与保存到数据库的操作
                            await this._mainWindow.UploadFileAndSaveToDatabase(this._dto);
                            this.CloseDialogWithResult(true); // 成功则关闭对话并返回 true
                            prevCursor = (Cursor)null;
                        }
                        catch (Exception ex)
                        {
                            // 上传失败提示并返回 false
                            int num = (int)MessageBox.Show("导入失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
                            this.CloseDialogWithResult(false);
                            prevCursor = (Cursor)null;
                        }
                    }
                    else
                    {
                        // 用户选择不关闭当前文件，则询问是否放弃识别
                        MessageBoxResult giveUp = MessageBox.Show("是否放弃本次识别并退出？", "放弃识别", MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
                        if (giveUp != MessageBoxResult.Yes)
                        {
                            prevCursor = (Cursor)null;
                        }
                        else
                        {
                            this.CloseDialogWithResult(false);
                            prevCursor = (Cursor)null;
                        }
                    }
                }
                finally
                {
                    // 无论如何都要恢复状态
                    this._isConfirmProcessing = false;// 重置处理标志，允许再次点击
                    this.BtnConfirm.IsEnabled = true; // 恢复按钮可用状态
                    Mouse.OverrideCursor = prevCursor;// 恢复先前光标，确保 UI 状态一致
                }
            }
        }

        /// <summary>
        /// 从UI的DataGrid中读取修改后的值，并更新回DTO对象
        /// </summary>
        private void UpdateDtoFromGrid()
        {
            // 从数据网格的 ItemsSource 中获取当前显示的属性列表
            var itemsSource = PropertiesGrid.ItemsSource as List<CategoryPropertyEditModel>;
            if (itemsSource == null) return;
            // 每次重新采集前先清空，避免旧值残留
            _dto.AttributesJson.Clear();
            // 确保 JSON 属性字典已初始化
            _dto.AttributesJson ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
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
                    case "创建时间": return "CREATED_AT";
                    case "更新时间": return "UPDATED_AT";
                    default: return key.Trim();
                }
            }
            foreach (var item in itemsSource)// 遍历每一行数据，回写到 DTO 中，确保用户修改的属性能够正确保存并上传，避免因未同步 UI 修改导致的数据不一致问题
            {
                // 先回写 FileStorage 固定字段
                _mainWindow.SetFileStorageProperty(_dto.FileStorage, item.PropertyName1, item.PropertyValue1);
                _mainWindow.SetFileStorageProperty(_dto.FileStorage, item.PropertyName2, item.PropertyValue2);

                // 再把两列属性写回 JSON 字典
                AddAttr(item.PropertyName1, item.PropertyValue1);
                AddAttr(item.PropertyName2, item.PropertyValue2);
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

            // 补充主表关键字段，保证 JSON 中也有一份稳定数据
            AddAttr("FileName", _dto.FileStorage.FileName ?? string.Empty);
            AddAttr("DisplayName", _dto.FileStorage.DisplayName ?? string.Empty);
            AddAttr("BlockName", _dto.FileStorage.BlockName ?? string.Empty);
            AddAttr("LayerName", _dto.FileStorage.LayerName ?? string.Empty);
            AddAttr("ColorIndex", _dto.FileStorage.ColorIndex?.ToString() ?? string.Empty);
            AddAttr("Scale", _dto.FileStorage.Scale?.ToString() ?? string.Empty);
            AddAttr("UpdatedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            AddAttr("FilePath", _dto.FileStorage.FilePath ?? string.Empty);

            _dto.FilePath = _dto.FileStorage.FilePath ?? string.Empty;
            _dto.DisplayName = _dto.FileStorage.DisplayName ?? string.Empty;
            _dto.BlockName = _dto.FileStorage.BlockName ?? string.Empty;
            _dto.LayerName = _dto.FileStorage.LayerName ?? string.Empty;
            _dto.ColorIndex = _dto.FileStorage.ColorIndex;
            _dto.Scale = _dto.FileStorage.Scale;
            _dto.CategoryId = _dto.FileStorage.CategoryId;
            _dto.CategoryType = _dto.FileStorage.CategoryType;
            _dto.CreatedBy = _dto.FileStorage.CreatedBy;
            _dto.Description = _dto.FileStorage.Description;
            _dto.PreviewImageName = _dto.FileStorage.PreviewImageName;
            _dto.PreviewImagePath = _dto.FileStorage.PreviewImagePath;

        }

    }
}
