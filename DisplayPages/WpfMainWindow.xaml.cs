using GB_NewCadPlus_LM.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Binding = System.Windows.Data.Binding;
using Border = System.Windows.Controls.Border;
using Brushes = System.Windows.Media.Brushes;
using Button = System.Windows.Controls.Button;
using ComboBox = System.Windows.Controls.ComboBox;
using DataGrid = System.Windows.Controls.DataGrid;
using DataTable = System.Data.DataTable;
using FontFamily = System.Windows.Media.FontFamily;
using Image = System.Windows.Controls.Image;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;
using Panel = System.Windows.Controls.Panel;
using Pen = System.Windows.Media.Pen;
using Point = System.Windows.Point;
using TabControl = System.Windows.Controls.TabControl;
using TextBox = System.Windows.Controls.TextBox;
using UserControl = System.Windows.Controls.UserControl;
using System.Text.RegularExpressions;
using GB_NewCadPlus_LM.FunctionalMethod;
using GB_NewCadPlus_LM.UniFiedStandards;
using Autodesk.AutoCAD.Internal.Calculator;
using FileStorage = GB_NewCadPlus_LM.FunctionalMethod.DatabaseManager.FileStorage;
using FileAttribute = GB_NewCadPlus_LM.FunctionalMethod.DatabaseManager.FileAttribute;
using CadCategory = GB_NewCadPlus_LM.FunctionalMethod.DatabaseManager.CadCategory;
using CadSubcategory = GB_NewCadPlus_LM.FunctionalMethod.DatabaseManager.CadSubcategory;

namespace GB_NewCadPlus_LM
{
    /// <summary>
    /// WpfMainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class WpfMainWindow : UserControl
    {
        #region  私有字段和属性

        /// <summary>
        /// 选中的预览图片路径
        /// </summary>
        private string? _selectedPreviewImagePath;
        /// <summary>
        /// 当前文件存储信息
        /// </summary>
        private FileStorage? _currentFileStorage;
        /// <summary>
        /// 当前文件属性信息
        /// </summary>
        private FileAttribute? _currentFileAttribute;
        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        private string? _connectionString;
        /// <summary>
        /// 图片缓存
        /// </summary>
        private readonly Dictionary<string, BitmapImage> _imageCache = new Dictionary<string, BitmapImage>();
        /// <summary>
        /// 预览图片缓存路径
        /// </summary>
        private readonly string _previewCachePath;
        /// <summary>
        /// 文件路径
        /// </summary>
        private string? _selectedFilePath;
        /// <summary>
        /// 文件属性信息
        /// </summary>
        private FileAttribute? _selectedFileAttribute;
        /// <summary>
        /// 文件管理器
        /// </summary>
        private FileManager? _fileManager;
        /// <summary>
        /// 新增字段：记住上一次显示到 PropertiesDataGrid 的键/值快照（不分大小写）
        /// </summary>
        private Dictionary<string, string> _propertiesSnapshotForInsert = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>
        /// 分类管理器
        /// </summary>
        private CategoryManager? _categoryManager;
        /// <summary>
        /// 在WpfMainWindow类中添加以下字段和属性
        /// </summary>
        private ManagementOperationType _currentOperation = ManagementOperationType.None;
        /// <summary>
        /// 创建结构树节点
        /// </summary>
        private List<CategoryTreeNode> _categoryTreeNodes = new List<CategoryTreeNode>();
        /// <summary>
        /// 添加数据库管理器
        /// </summary>
        private DatabaseManager? _databaseManager;
        /// <summary>
        /// 在WpfMainWindow类中添加字段
        /// </summary>
        private CategoryTreeNode? _selectedCategoryNode; // 在分类架构树的当前选中的分类节点
        /// <summary>
        /// 添加枚举类型
        /// </summary>
        public enum ManagementOperationType
        {
            None,
            AddCategory,
            AddSubcategory
        }
        /// <summary>
        /// 是否使用数据库模式
        /// </summary>
        private bool _useDatabaseMode = true;
        /// <summary>
        /// 当前选中的数据库类型（CAD或SW）
        /// </summary>
        private string _currentDatabaseType = "";
        /// <summary>
        /// 当前选中的节点ID
        /// </summary>
        private int _currentNodeId = 0;
        /// <summary>
        /// CAD文件存储路径
        /// </summary>
        private string _cadStoragePath = "";
        /// <summary>
        /// SW文件存储路径
        /// </summary>
        private string _swStoragePath = "";
        /// <summary>
        /// 用于显示分类树的TreeView控件
        /// </summary>
        private System.Windows.Controls.TreeView? _categoryTreeView;
        /// <summary>
        /// 添加预览图片显示的Viewbox引用
        /// </summary>
        private Viewbox? previewViewbox;
        /// <summary>
        /// 拿到本app的local的路径，并创建GB_CADPLUS文件夹
        /// </summary>
        public static string AppPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GB_CADPLUS");
        /// <summary>
        /// 文件路径与名称  resourcesFile
        /// </summary>
        public static string? filePathAndName = null;
        /// <summary>
        /// 引用文件referenceFile文件夹  
        /// </summary>
        public static string referenceFile = System.IO.Path.Combine(AppPath, "ReferenceFile");
        /// <summary>
        /// 层管理器
        /// </summary>
        private LayerManager? _layerManager;
        /// <summary>
        /// 层数据源
        /// </summary>
        private ObservableCollection<LayerInfo>? _layerData;

        #endregion

        /// <summary>
        ///（ public partial class WpfMainWindow : UserControl { 之后）添加静态实例声明
        /// </summary>
        public static WpfMainWindow? Instance { get; private set; }

        /// <summary>
        /// 从 LoginWindow 读取服务器配置
        /// </summary>
        private void LoadServerConfigFromLogin()
        {
            try
            {
                // 从登录窗口中读取服务器配置
                var configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GB_NewCadPlus_LM", "login_config.json");
                if (!File.Exists(configPath))// 如果文件不存在，则返回
                    return;

                var json = File.ReadAllText(configPath);//读取配置文件内容
                var ser = new JavaScriptSerializer();//创建JavaScript序列化器
                var dict = ser.Deserialize<Dictionary<string, object>>(json);//反序列化JSON
                if (dict == null)
                    return;

                if (dict.TryGetValue("ServerIP", out var sip) && sip != null)// 尝试获取 ServerIP
                {
                    // 将登录窗口中的服务器 IP 同步回设置界面文本框，保证 UI 与实际使用一致
                    TextBox_Set_ServiceIP.Text = sip.ToString();
                }

                if (dict.TryGetValue("ServerPort", out var sport) && sport != null)// 尝试获取 ServerPort
                {
                    TextBox_Set_ServicePort.Text = sport.ToString();//同步回设置界面文本框
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载登录窗口配置失败: {ex.Message}");
                LogManager.Instance.LogInfo($"加载登录窗口配置失败: {ex.Message}");
                // 不要抛异常，失败则继续使用界面中已有的值
            }
        }

        /// <summary>
        /// WpfMainWindow主界面
        /// </summary>
#pragma warning disable CS8618 // 在退出构造函数时，不可为 null 的字段必须包含非 null 值。请考虑添加 "required" 修饰符或声明为可为 null。
        public WpfMainWindow()
#pragma warning restore CS8618 // 在退出构造函数时，不可为 null 的字段必须包含非 null 值。请考虑添加 "required" 修饰符或声明为可为 null。
        {
            InitializeComponent();//初始化界面

            // 设置全局静态实例，供非 UI 代码读取 TextBox_绘图比例
            Instance = this;
            UnifiedUIManager.SetWpfInstance(this); // 注册到统一管理器
            LogManager.Instance.LogInfo("WPF实例已注册到UnifiedUIManager"); // 调试输出，确认注册成功
            NewTjLayer();//初始化图层
            Loaded += WpfMainWindow_Loaded;//加载按钮
            // 初始化预览图片缓存路径
            _previewCachePath = Path.Combine(AppPath, "PreviewCache");
            //_fileManager = new FileManager(_databaseManager);
            //_categoryManager = new CategoryManager(_databaseManager);
            _fileManager = null;
            _categoryManager = null;
            if (!Directory.Exists(_previewCachePath))
            {
                Directory.CreateDirectory(_previewCachePath);
            }

            _layerManager = new LayerManager();
            _layerData = new ObservableCollection<LayerInfo>();
            // 初始化DataGrid
            InitializeLayerDataGrid();
        }
        /// <summary>
        /// 初始化图层数据DataGrid表
        /// </summary>
        private void InitializeLayerDataGrid()
        {
            // 如果您使用的是DataGrid，可以这样设置
            LayerDataGrid.AutoGenerateColumns = false;
            LayerDataGrid.ItemsSource = _layerData;
        }

        /// <summary>
        /// 初始化 LayerDictionary DataGrid 的数据源（在窗口初始化时调用）LayerDataGrid 加载本地图层
        /// </summary>
        private async Task InitializeLayerDictionaryDataGridSource()
        {
            // 绑定集合到 DataGrid
            LayerDictionary_DataGrid.ItemsSource = _layerDictionaryRows;

            // 强制设置 ScrollViewer 行为（防止 XAML 未生效）
            LayerDictionary_DataGrid.SetValue(ScrollViewer.VerticalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
            LayerDictionary_DataGrid.SetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
            LayerDictionary_DataGrid.SetValue(ScrollViewer.CanContentScrollProperty, false); // 像素滚动更平滑

            // 绑定编辑事件以便捕获 ComboBox 并处理 SelectionChanged
            LayerDictionary_DataGrid.PreparingCellForEdit -= LayerDictionary_DataGrid_PreparingCellForEdit;
            LayerDictionary_DataGrid.PreparingCellForEdit += LayerDictionary_DataGrid_PreparingCellForEdit;
            LayerDictionary_DataGrid.CellEditEnding -= LayerDictionary_DataGrid_CellEditEnding;
            LayerDictionary_DataGrid.CellEditEnding += LayerDictionary_DataGrid_CellEditEnding;

            // 先尝试加载分类名
            try
            {
                if ((_categoryNames == null || _categoryNames.Count == 0) && _databaseManager != null && _databaseManager.IsDatabaseAvailable)
                {
                    await LoadCategoryNamesAsync();
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"初始化 CategoryNames 时出错: {ex.Message}");
            }

            // DataGridComboBoxColumn 在视觉树外，XAML 绑定有时无效 -> 在代码中分配 ItemsSource
            try
            {
                var comboCol = LayerDictionary_DataGrid.Columns
                    .OfType<DataGridComboBoxColumn>()
                    .FirstOrDefault(c => (c.Header?.ToString() ?? string.Empty).IndexOf("专业", StringComparison.OrdinalIgnoreCase) >= 0);

                if (comboCol != null)
                {
                    comboCol.ItemsSource = _categoryNames;
                    comboCol.SelectedItemBinding = new System.Windows.Data.Binding("Major")
                    {
                        Mode = BindingMode.TwoWay,
                        UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                    };
                }
                else
                {
                    if (LayerDictionary_DataGrid.Columns.Count > 1 && !(LayerDictionary_DataGrid.Columns[1] is DataGridComboBoxColumn))
                    {
                        var newCombo = new DataGridComboBoxColumn
                        {
                            Header = "专业",
                            Width = 75,
                            ItemsSource = _categoryNames,
                            SelectedItemBinding = new Binding("Major") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }
                        };
                        LayerDictionary_DataGrid.Columns[1] = newCombo;
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"为 LayerDictionary_DataGrid 设置下拉数据源失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 窗口初始化时运行加载项
        /// </summary>
        /// <param Name="sender"></param>
        /// <param Name="e"></param>
        private async void WpfMainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // 尝试从 login_config 或界面同步服务器设置
                LoadServerConfigFromLogin();
                // 加载绘图配置
                LoadDrawingConfig();
                // 尝试连接数据库（若失败会弹出登录窗口让用户修正）
                bool connected = await EnsureDatabaseConnectedOrShowLoginAsync();
                if (!connected)
                {
                    _useDatabaseMode = false;
                    MessageBox.Show("未能连接到数据库，已进入离线模式。部分功能将不可用。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    _useDatabaseMode = true;
                    // 初始化管理器（如果尚未初始化）
                    if (_databaseManager != null)
                    {
                        _fileManager = new FileManager(_databaseManager);
                        _categoryManager = new CategoryManager(_databaseManager);
                    }
                    // 后续数据库相关初始化（例如加载分类树）
                    ReinitializeDatabase();
                    // 根据当前登录用户决定是否显示管理员/部门模块
                    UpdateAdminTabsVisibility();
                }
                // 继续其它 UI 初始化（保持原有逻辑）
                AddContextMenuToTreeView(CategoryTreeView);
                PropertiesDataGrid = FindVisualChild<DataGrid>(this, "PropertiesDataGrid");
                // 初始化图层字典数据源（等待 CategoryNames 可用后绑定列的 ItemsSource）
                await InitializeLayerDictionaryDataGridSource();
                Loaded += DepartmentAdminControl_Loaded;//注册加载完成事件
                                                        //Load();
                                                        // 替换原有的版本号显示代码，修复 CS0120 错误
                                                        //TextBox_PluginVersion.Text = $"{System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}" +
                                                        //    $"\n{System.Reflection.Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion}";
                                                        //显示版本号
                TextBox_PluginVersion.Text = $"{System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}";


            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"WpfMainWindow_Loaded 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 根据当前用户（VariableDictionary._userName 或界面用户名）显示/隐藏“管理员模块”与“部门\\人员模块”
        /// 仅当用户名为 "sa" 或 "admin"（不区分大小写）时才显示；否则折叠（Collapsed）
        ///</summary>
        private void UpdateAdminTabsVisibility()
        {
            try
            {
                // 获取当前用户名（优先使用全局变量，再退回到设置界面输入）
                var userName = (VariableDictionary._userName ?? TextBox_Set_Username.Text ?? string.Empty).Trim().ToLowerInvariant();
                bool isAdmin = userName == "sa" || userName == "admin" || userName == "root";

                if (MainTabControl == null)
                {
                    LogManager.Instance.LogInfo("UpdateAdminTabsVisibility: MainTabControl 为 null，跳过");
                    return;
                }
                if (!isAdmin)
                    // 遍历 TabControl 的 TabItem，根据 Header 文本判断并设置 Visibility
                    foreach (var item in MainTabControl.Items)
                    {
                        if (item is TabItem tab)
                        {
                            string header = tab.Header?.ToString() ?? string.Empty;

                            // 匹配包含“管理员”关键词的 TabItem
                            if (header.IndexOf("管理员", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                tab.Visibility = System.Windows.Visibility.Collapsed;
                                continue;
                            }

                            // 匹配“部门”或“人员”关键词（常见组合为 “部门/人员”、“部门\人员” 等）
                            if (header.IndexOf("部门", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                header.IndexOf("人员", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // 进一步避免误伤（例如 “人员列表” 也算人员模块）
                                // 这里按需求统一控制：只要包含部门或人员就受权限控制
                                tab.Visibility = System.Windows.Visibility.Collapsed;
                            }
                        }
                    }

                // 如果当前选中项被隐藏，切换到第一个可见的 TabItem
                if (MainTabControl.SelectedItem is TabItem selected && selected.Visibility != System.Windows.Visibility.Visible)
                {
                    foreach (var item in MainTabControl.Items)
                    {
                        if (item is TabItem t && t.Visibility == System.Windows.Visibility.Visible)
                        {
                            MainTabControl.SelectedItem = t;
                            break;
                        }
                    }
                }

                LogManager.Instance.LogInfo($"UpdateAdminTabsVisibility: user='{userName}', isAdmin={isAdmin}");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"UpdateAdminTabsVisibility 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 绘图配置文件路径
        /// </summary>
        private string DrawingConfigPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "GB_NewCadPlus_LM",
            "drawing_config.json"
        );

        /// <summary>
        /// 绘图配置类
        /// </summary>
        public class DrawingConfig
        {
            /// <summary>
            /// 绘图比例
            /// </summary>
            public double DrawingScale { get; set; } = 100.0; // 默认1:100

            /// <summary>
            /// 表格行高比例
            /// </summary>
            public double TableRowHeightScale { get; set; } = 1.0;

            /// <summary>
            /// 文字高度比例
            /// </summary>
            public double TextHeightScale { get; set; } = 1.0;

            /// <summary>
            /// 箭头大小比例
            /// </summary>
            public double ArrowSizeScale { get; set; } = 1.0;

            /// <summary>
            /// 其他配置项（保留扩展性）
            /// </summary>
            public Dictionary<string, object> AdditionalSettings { get; set; } = new Dictionary<string, object>();
        }

        /// <summary>
        /// 加载绘图配置
        /// </summary>
        private void LoadDrawingConfig()
        {
            try
            {
                double loadedScale = 0.0;

                // 1) 尝试从本地 json 配置读取
                if (File.Exists(DrawingConfigPath))
                {
                    try
                    {
                        string jsonContent = File.ReadAllText(DrawingConfigPath);
                        var config = Newtonsoft.Json.JsonConvert.DeserializeObject<DrawingConfig>(jsonContent);
                        if (config != null && config.DrawingScale > 0.0)
                        {
                            loadedScale = config.DrawingScale;
                        }
                    }
                    catch (Exception exInner)
                    {
                        LogManager.Instance.LogWarning($"解析本地绘图配置失败，回退到其它来源: {exInner.Message}");
                    }
                }

                // 2) 若本地 json 未提供有效值，则尝试优先使用界面输入（若已初始化）
                if (loadedScale <= 0.0)
                {
                    try
                    {
                        var uiText = TextBox_绘图比例?.Text ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(uiText))
                        {
                            var s = uiText.Trim();
                            if (!double.TryParse(s, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out double parsed))
                            {
                                // 兼容逗号作为小数分隔符
                                if (!double.TryParse(s.Replace(',', '.'), System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out parsed))
                                    parsed = 0.0;
                            }
                            if (parsed > 0.0) loadedScale = parsed;
                        }
                    }
                    catch { loadedScale = 0.0; }
                }

                // 3) 若仍无有效值，回退到 AutoCadHelper 检测（视口或数据库检测）
                if (loadedScale <= 0.0)
                {
                    try
                    {
                        // AutoCadHelper.GetScale(true) 可能返回规范化因子或分母相关值，使用 FontsStyleHelper 判定分母
                        double factor = AutoCadHelper.GetScale(true);
                        loadedScale = TextFontsStyleHelper.DetermineScaleDenominator(factor, null, false);
                    }
                    catch { loadedScale = 100.0; }
                }

                // 4) 最终兜底
                if (loadedScale <= 0.0 || double.IsNaN(loadedScale) || double.IsInfinity(loadedScale))
                    loadedScale = 100.0;

                // 把值写回 UI（使用 InvariantCulture 显示）与全局变量
                if (TextBox_绘图比例 != null)
                {
                    try
                    {
                        TextBox_绘图比例.Text = loadedScale.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }
                    catch
                    {
                        TextBox_绘图比例.Text = loadedScale.ToString();
                    }
                }

                // 更新全局并刷新 AutoCAD 端缓存（确保其它模块读取到最新值）
                VariableDictionary.blockScale = loadedScale;
                try
                {
                    AutoCadHelper.Invalidate(); // 清缓存
                    AutoCadHelper.GetAndApplyActiveDrawingScale(); // 触发一次应用（会再次读取界面优先级）
                }
                catch { /* 忽略AutoCAD端刷新失败 */ }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"加载绘图配置失败: {ex.Message}");
                // 出错时使用默认值
                if (TextBox_绘图比例 != null)
                {
                    TextBox_绘图比例.Text = "100";
                }
                VariableDictionary.blockScale = 100.0;
            }
        }

        /// <summary>
        /// 保存绘图配置
        /// </summary>
        private void SaveDrawingConfig()
        {
            try
            {
                var config = new DrawingConfig
                {
                    DrawingScale = 100.0 // 默认值
                };

                // 从输入框获取绘图比例（兼容逗号小数分隔符）
                if (TextBox_绘图比例 != null)
                {
                    var raw = (TextBox_绘图比例.Text ?? string.Empty).Trim();
                    if (!string.IsNullOrEmpty(raw))
                    {
                        double scale = 0.0;
                        // 先使用 InvariantCulture 解析
                        if (!double.TryParse(raw, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out scale))
                        {
                            // 兼容逗号作为小数分隔符
                            var alt = raw.Replace(',', '.');
                            double.TryParse(alt, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out scale);
                        }

                        if (scale > 0.0)
                            config.DrawingScale = scale;
                    }
                }

                // 更新全局变量，保证立即生效
                VariableDictionary.blockScale = config.DrawingScale;

                // 刷新 AutoCadHelper 的内部缓存（使 CAD 端能及时读取到新比例）
                try
                {
                    AutoCadHelper.Invalidate();
                    // 再次应用以确保 VariableDictionary.blockScale 与 AutoCadHelper 的检测逻辑保持一致
                    AutoCadHelper.GetAndApplyActiveDrawingScale();
                }
                catch { /* 忽略 */ }

                // 确保配置目录存在并保存 legacy drawing_config.json（保持向后兼容）
                string configDir = Path.GetDirectoryName(DrawingConfigPath);
                if (!string.IsNullOrEmpty(configDir) && !Directory.Exists(configDir))
                {
                    Directory.CreateDirectory(configDir);
                }

                // 序列化并保存配置（使用 InvariantCulture 以便数值稳定）
                var serSettings = new Newtonsoft.Json.JsonSerializerSettings
                {
                    Formatting = Newtonsoft.Json.Formatting.Indented,
                    Culture = System.Globalization.CultureInfo.InvariantCulture
                };
                string jsonContent = Newtonsoft.Json.JsonConvert.SerializeObject(config, serSettings);
                File.WriteAllText(DrawingConfigPath, jsonContent);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"保存绘图配置失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 从TextBox_绘图比例获取用户设置的比例值（更鲁棒的解析：支持千位/小数、逗号/点）
        /// </summary>
        /// <returns>用户设置的比例值，如果获取失败返回0</returns>
        public double GetDrawingScaleFromTextBox()
        {
            try
            {
                if (TextBox_绘图比例 == null) return 0;

                var raw = (TextBox_绘图比例.Text ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(raw)) return 0;

                // 尝试直接使用不变文化解析
                if (double.TryParse(raw, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out double v) && v > 0)
                    return v;

                // 兼容逗号作为小数分隔符
                var alt = raw.Replace(',', '.');
                if (double.TryParse(alt, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out v) && v > 0)
                    return v;

                // 兼容当前区域设置解析（最后兜底）
                if (double.TryParse(raw, out v) && v > 0) return v;

            }
            catch
            {
                // 忽略并返回0
            }

            return 0;
        }

        /// <summary>
        /// 尝试使用当前界面/配置连接数据库；失败时弹出 LoginWindow 让用户登录并返回 login.CreatedDatabaseManager（若创建成功）。
        /// 成功时设置 _databaseManager 并返回 true；失败返回 false。
        /// </summary>
        private async Task<bool> EnsureDatabaseConnectedOrShowLoginAsync()
        {
            try
            {
                // 读取当前界面配置（优先）
                VariableDictionary._serverIP = TextBox_Set_ServiceIP.Text?.Trim();
                if (string.IsNullOrWhiteSpace(VariableDictionary._serverIP))
                {
                    // 如果界面无值，尝试从 login_config 中加载（LoadServerConfigFromLogin 已被调用）
                    VariableDictionary._serverIP = TextBox_Set_ServiceIP.Text?.Trim() ?? "127.0.0.1";

                }
                /// 端口
                VariableDictionary._serverPort = int.TryParse(TextBox_Set_ServicePort.Text, out var p) ? p : 3306;


                // 快速 TCP 检测
                bool tcpOk = await Task.Run(() => LoginWindow.TestNetworkConnection(VariableDictionary._serverIP, VariableDictionary._serverPort));
                if (tcpOk)
                {
                    // 尝试创建 DatabaseManager（使用项目约定的默认参数）
                    string conn = $"Server={VariableDictionary._serverIP};Port={VariableDictionary._serverPort};Database=cad_sw_library;Uid=root;Pwd=root;";
                    try
                    {
                        var db = new DatabaseManager(conn);
                        if (db.IsDatabaseAvailable)
                        {
                            _databaseManager = db;
                            _useDatabaseMode = true;
                            LogManager.Instance.LogInfo($"直接连接数据库成功：{VariableDictionary._serverIP}:{VariableDictionary._serverPort}");
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogInfo($"直接构造 DatabaseManager 失败: {ex.Message}");
                    }
                }
                else
                {
                    LogManager.Instance.LogInfo("TCP 层不可达，准备弹出登录窗口让用户更新配置。");
                }

                // 弹出登录窗口（模态），由 LoginWindow 尝试创建 DatabaseManager
                var login = new LoginWindow();
                try
                {
                    var owner = Window.GetWindow(this);
                    if (owner != null) login.Owner = owner;
                }
                catch { /* 忽略 */ }

                var dlg = login.ShowDialog();
                if (dlg == true)
                {
                    // 如果 LoginWindow 创建了 DatabaseManager，直接使用
                    if (login.CreatedDatabaseManager != null && login.CreatedDatabaseManager.IsDatabaseAvailable)
                    {
                        _databaseManager = login.CreatedDatabaseManager;
                        _useDatabaseMode = true;
                        LogManager.Instance.LogInfo("使用 LoginWindow 返回的 DatabaseManager 成功连接数据库。");
                        return true;
                    }

                    // 否则尝试使用登录窗口保存的界面值重试一次（LoadServerConfigFromLogin 会把配置同步到界面）
                    LoadServerConfigFromLogin();
                    VariableDictionary._serverIP = TextBox_Set_ServiceIP.Text?.Trim() ?? VariableDictionary._serverIP;
                    VariableDictionary._serverPort = int.TryParse(TextBox_Set_ServicePort.Text, out var pp) ? pp : VariableDictionary._serverPort;
                    string conn2 = $"Server={VariableDictionary._serverIP};Port={VariableDictionary._serverPort};Database=cad_sw_library;Uid=root;Pwd=root;SslMode=None;";
                    try
                    {
                        var db2 = new DatabaseManager(conn2);
                        if (db2.IsDatabaseAvailable)
                        {
                            _databaseManager = db2;
                            _useDatabaseMode = true;
                            LogManager.Instance.LogInfo("使用登录后配置创建 DatabaseManager 成功。");
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogInfo($"使用登录后配置创建 DatabaseManager 失败: {ex.Message}");
                    }
                }

                // 所有尝试均失败
                return false;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"EnsureDatabaseConnectedOrShowLoginAsync 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 重新初始化数据库连接
        /// </summary>
        private async void ReinitializeDatabase()
        {
            try
            {
                // 先加载分类名用于下拉绑定
                await LoadCategoryNamesAsync();
                await _categoryManager.RefreshCategoryTreeAsync(_selectedCategoryNode, _categoryTreeView, _categoryTreeNodes, _databaseManager);

                LogManager.Instance.LogInfo("数据库连接已重新初始化");
                // 在树刷新后，主动刷新主分类面板
                await RefreshAllCategoryPanelsAsync();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"重新初始化数据库时出错: {ex.Message}");
                MessageBox.Show($"重新初始化数据库失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// 刷新所有分类面板
        /// </summary>
        /// <returns></returns>
        private async Task RefreshAllCategoryPanelsAsync()
        {
            try
            {
                if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
                {
                    LogManager.Instance.LogInfo("RefreshAllCategoryPanelsAsync：数据库不可用，跳过面板刷新");
                    return;
                }

                var majorCategories = new[]
                { "工艺","建筑","结构","电气","给排水","暖通","自控","总图","公共图" };

                foreach (var majorItem in majorCategories)// 遍历所有主分类
                {
                    try
                    {
                        var majorPanel = GetPanelByFolderName(majorItem);// 根据分类名称获取面板
                        if (majorPanel == null)
                        {
                            LogManager.Instance.LogInfo($"RefreshAllCategoryPanelsAsync：未找到面板 {majorItem}，跳过");
                            continue;
                        }

                        // 调用已有方法加载（方法内部已处理回退到 Resources）
                        await LoadButtonsFromDatabase(majorItem, majorPanel);

                        // 稍作延迟以减少短时间内并发压力（可根据需要调整或删除）
                        await Task.Delay(60);
                    }
                    catch (Exception exInner)
                    {
                        LogManager.Instance.LogInfo($"RefreshAllCategoryPanelsAsync: 加载分类 {majorItem} 时出错: {exInner.Message}");
                    }
                }

                LogManager.Instance.LogInfo("RefreshAllCategoryPanelsAsync: 所有主分类面板刷新完成");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"RefreshAllCategoryPanelsAsync 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加端口输入验证（可选）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBox_Set_ServicePort_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // 只允许输入数字
            e.Handled = !IsTextAllowed(e.Text);
        }

        /// <summary>
        /// 初始化分类属性编辑网格
        /// </summary>
        private void InitializeCategoryPropertyGrid()
        {
            var initialRows = new List<CategoryPropertyEditModel>
                {
                    new CategoryPropertyEditModel(),
                    new CategoryPropertyEditModel(),
                    new CategoryPropertyEditModel()
                };

            CategoryPropertiesDataGrid.ItemsSource = initialRows;
            LogManager.Instance.LogInfo("初始化分类属性编辑网格成功:InitializeCategoryPropertyGrid()");
        }

        /// <summary>
        /// 初始化属性编辑网格
        /// </summary>
        private void AddFileInitializeFilePropertiesGrid()
        {
            try
            {
                var properties = new List<CategoryPropertyEditModel>
                  {
                  // 文件存储表(cad_file_storage)相关属性
                  new CategoryPropertyEditModel { PropertyName1 = "显示名称", PropertyValue1 = Path.GetFileNameWithoutExtension(_selectedFilePath), PropertyName2 = "元素块名", PropertyValue2 = "" },
                  new CategoryPropertyEditModel { PropertyName1 = "层名", PropertyValue1 = "TJ(  专业  )", PropertyName2 = "颜色索引", PropertyValue2 = "40" },
                  new CategoryPropertyEditModel { PropertyName1 = "描述", PropertyValue1 = "", PropertyName2 = "版本", PropertyValue2 = "1" },
                  new CategoryPropertyEditModel { PropertyName1 = "是否公开", PropertyValue1 = "是", PropertyName2 = "创建者", PropertyValue2 = Environment.UserName },
                  new CategoryPropertyEditModel { PropertyName1 = "是否天正", PropertyValue1 = "否" },
                  // 文件属性表(cad_file_attributes)相关属性
                  new CategoryPropertyEditModel { PropertyName1 = "长度", PropertyValue1 = "", PropertyName2 = "宽度", PropertyValue2 = "" },
                  new CategoryPropertyEditModel { PropertyName1 = "高度", PropertyValue1 = "", PropertyName2 = "角度", PropertyValue2 = "0" },
                  new CategoryPropertyEditModel { PropertyName1 = "基点X", PropertyValue1 = "0", PropertyName2 = "基点Y", PropertyValue2 = "0" },
                  new CategoryPropertyEditModel { PropertyName1 = "基点Z", PropertyValue1 = "0", PropertyName2 = "介质", PropertyValue2 = "" },
                  new CategoryPropertyEditModel { PropertyName1 = "规格", PropertyValue1 = "", PropertyName2 = "材质", PropertyValue2 = "" },
                  new CategoryPropertyEditModel { PropertyName1 = "标准号", PropertyValue1 = "", PropertyName2 = "功率", PropertyValue2 = "" },
                  new CategoryPropertyEditModel { PropertyName1 = "容积", PropertyValue1 = "", PropertyName2 = "压力", PropertyValue2 = "" },
                  new CategoryPropertyEditModel { PropertyName1 = "温度", PropertyValue1 = "", PropertyName2 = "直径", PropertyValue2 = "" },
                  new CategoryPropertyEditModel { PropertyName1 = "外径", PropertyValue1 = "", PropertyName2 = "内径", PropertyValue2 = "" },
                  new CategoryPropertyEditModel { PropertyName1 = "厚度", PropertyValue1 = "", PropertyName2 = "重量", PropertyValue2 = "" },
                  new CategoryPropertyEditModel { PropertyName1 = "型号", PropertyValue1 = "", PropertyName2 = "备注", PropertyValue2 = "" },
                  
                  // 文件标签表(file_tags)相关属性（可以添加多个标签）
                  new CategoryPropertyEditModel { PropertyName1 = "标签1", PropertyValue1 = "", PropertyName2 = "标签2", PropertyValue2 = "" },
                  new CategoryPropertyEditModel { PropertyName1 = "标签3", PropertyValue1 = "", PropertyName2 = "", PropertyValue2 = "" }
                  };

                CategoryPropertiesDataGrid.ItemsSource = properties;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"初始化属性编辑失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取默认预览图片
        /// </summary>
        private BitmapImage GetDefaultPreviewImage()
        {
            try
            {
                // 首先尝试从资源加载默认图片
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri("pack://application:,,,/GB_NewCadPlus_LM;component/Resources/default_preview.png");
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();
                return bitmap;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载默认预览图片失败: {ex.Message}");

                // 如果资源图片不存在，创建一个纯色图片
                try
                {
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri("pack://application:,,,/GB_NewCadPlus_LM;component/Resources/no_preview.png");
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    bitmap.Freeze();
                    return bitmap;
                }
                catch
                {
                    // 如果都失败了，创建一个空白图片
                    return CreatePlaceholderImage();
                }
            }
        }

        /// <summary>
        /// 创建占位符图片
        /// </summary>
        private BitmapImage CreatePlaceholderImage()
        {
            try
            {
                // 创建一个简单的占位符图片
                var bitmap = new RenderTargetBitmap(80, 60, 96, 96, PixelFormats.Pbgra32);

                var drawingVisual = new DrawingVisual();
                using (var drawingContext = drawingVisual.RenderOpen())
                {
                    // 绘制灰色背景
                    drawingContext.DrawRectangle(Brushes.LightGray, new Pen(Brushes.Gray, 1), new System.Windows.Rect(0, 0, 80, 60));

                    // 绘制"No Preview"文本
                    var text = new FormattedText(
                        "无预览",
                        System.Globalization.CultureInfo.CurrentCulture,
                        FlowDirection,
                        new Typeface("Arial"),
                        12,
                        Brushes.Gray);

                    drawingContext.DrawText(text, new Point(20, 20));
                }

                bitmap.Render(drawingVisual);

                // 转换为BitmapImage
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmap));

                using (var stream = new MemoryStream())
                {
                    encoder.Save(stream);
                    stream.Position = 0;

                    var result = new BitmapImage();
                    result.BeginInit();
                    result.StreamSource = stream;
                    result.CacheOption = BitmapCacheOption.OnLoad;
                    result.EndInit();
                    result.Freeze();
                    return result;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"创建占位符图片失败: {ex.Message}");

                // 最后的备选方案：返回空的BitmapImage
                return new BitmapImage();
            }
        }

        /// <summary>
        /// 清理无效的图片缓存
        /// </summary>
        private void CleanupInvalidImageCache()
        {
            try
            {
                var invalidKeys = new List<string>();

                foreach (var kvp in _imageCache)
                {
                    try
                    {
                        // 检查图片是否仍然有效
                        if (kvp.Value == null || kvp.Value.Width <= 0 || kvp.Value.Height <= 0)
                        {
                            invalidKeys.Add(kvp.Key);
                        }
                    }
                    catch
                    {
                        invalidKeys.Add(kvp.Key);
                    }
                }

                // 移除无效的缓存项
                foreach (string key in invalidKeys)
                {
                    _imageCache.Remove(key);
                }

                LogManager.Instance.LogInfo($"清理了 {invalidKeys.Count} 个无效的图片缓存项");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"清理图片缓存时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 从服务器获取预览图片并缓存
        /// </summary>
        private async Task<BitmapImage> GetPreviewImageAsync(FileStorage fileStorage)
        {
            try
            {
                // 检查文件存储对象是否有效
                if (fileStorage == null)
                {
                    LogManager.Instance.LogInfo("文件存储对象为空");
                    return GetDefaultPreviewImage();
                }

                // 检查内存缓存
                if (_imageCache.ContainsKey(fileStorage.FilePath ?? fileStorage.Id.ToString()))
                {
                    return _imageCache[fileStorage.FilePath ?? fileStorage.Id.ToString()];
                }

                // 检查是否有预览图片路径
                string previewImagePath = fileStorage.PreviewImagePath ?? fileStorage.FilePath;
                if (string.IsNullOrEmpty(previewImagePath))
                {
                    LogManager.Instance.LogInfo("预览图片路径为空");
                    return GetDefaultPreviewImage();
                }

                // 检查本地缓存文件
                string cacheFileName = $"{fileStorage.Id}_{Path.GetFileName(previewImagePath)}.png";
                string cacheFilePath = Path.Combine(_previewCachePath, cacheFileName);

                // 如果本地缓存存在且有效，直接加载
                if (File.Exists(cacheFilePath))
                {
                    try
                    {
                        var bitmap = LoadImageFromFile(cacheFilePath);
                        if (bitmap != null)
                        {
                            // 添加到内存缓存
                            _imageCache[fileStorage.FilePath ?? fileStorage.Id.ToString()] = bitmap;
                            return bitmap;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogInfo($"从本地缓存加载图片失败: {ex.Message}");
                        // 删除损坏的缓存文件
                        try { File.Delete(cacheFilePath); } catch { }
                    }
                }

                // 尝试从原始路径加载图片
                if (File.Exists(previewImagePath))
                {
                    try
                    {
                        var bitmap = LoadImageFromFile(previewImagePath);
                        if (bitmap != null)
                        {
                            // 保存到本地缓存
                            try
                            {
                                using (var fileStream = File.Create(cacheFilePath))
                                {
                                    var encoder = new PngBitmapEncoder();
                                    encoder.Frames.Add(BitmapFrame.Create(bitmap));
                                    encoder.Save(fileStream);
                                }
                            }
                            catch (Exception cacheEx)
                            {
                                LogManager.Instance.LogInfo($"保存到缓存失败: {cacheEx.Message}");
                            }

                            // 添加到内存缓存
                            _imageCache[fileStorage.FilePath ?? fileStorage.Id.ToString()] = bitmap;
                            return bitmap;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogInfo($"从原始路径加载图片失败: {ex.Message}");
                    }
                }

                // 如果所有方法都失败，返回默认图片
                return GetDefaultPreviewImage();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"获取预览图片时出错: {ex.Message}");
                return GetDefaultPreviewImage();
            }
        }

        /// <summary>
        /// 从文件加载图片（带错误处理）
        /// </summary>
        private BitmapImage LoadImageFromFile(string imagePath)
        {
            try
            {
                if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath))
                    return null;

                // 获取文件扩展名
                string extension = Path.GetExtension(imagePath)?.ToLower();

                // 根据文件类型使用不同的加载方法
                switch (extension)
                {
                    case ".png":
                    case ".jpg":
                    case ".jpeg":
                    case ".bmp":
                    case ".gif":
                    case ".tif":
                    case ".tiff":
                        return LoadStandardImage(imagePath);
                    default:
                        // 尝试使用默认方式加载
                        return LoadStandardImage(imagePath);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载图片文件失败 {imagePath}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 加载标准图片格式
        /// </summary>
        private BitmapImage LoadStandardImage(string imagePath)
        {
            try
            {
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(imagePath, UriKind.Absolute);
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                bitmap.EndInit();
                bitmap.Freeze(); // 冻结以提高性能
                return bitmap;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载标准图片失败 {imagePath}: {ex.Message}");

                // 尝试使用流方式加载
                try
                {
                    using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                    {
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.StreamSource = stream;
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        bitmap.Freeze();
                        return bitmap;
                    }
                }
                catch (Exception streamEx)
                {
                    LogManager.Instance.LogInfo($"流方式加载图片也失败 {imagePath}: {streamEx.Message}");
                    return null;
                }
            }
        }

        /// <summary>
        /// 在窗口关闭时清理缓存
        /// </summary>
        /// <param name="e"></param>
        //protected void OnClosing(CancelEventArgs e)
        //{
        //    try
        //    {
        //        // 停止同步
        //        //_serverSyncManager?.StopSync();

        //        // 清理图片缓存
        //        CleanupInvalidImageCache();

        //        OnClosing(e);
        //    }
        //    catch (Exception ex)
        //    {
        //        LogManager.Instance.LogInfo($"关闭窗口时清理缓存失败: {ex.Message}");
        //    }
        //}

        // 替换原来的 OnClosing 方法，避免无限递归。
        // 在宿主 Window 的 Closing/Unloaded 事件中调用 CleanupOnClosing(new CancelEventArgs())

        /// <summary>
        /// 来执行清理操作。
        /// </summary>
        protected void CleanupOnClosing(CancelEventArgs e)
        {
            try
            {
                // 停止同步或其它需要清理的服务（如果有）
                //_serverSyncManager?.StopSync();

                // 清理图片缓存
                CleanupInvalidImageCache();

                // 如果有其它需要释放的资源或取消的操作，在这里处理
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"关闭窗口时清理缓存失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加手动清理缓存按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 清理缓存按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                CleanupInvalidImageCache();
                MessageBox.Show("缓存已清理", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"清理缓存失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// 获取缩放比例
        /// </summary>
        /// <param name="fileStorage"></param>
        /// <returns></returns>
        private double GetScaleFromFileStorage(FileStorage fileStorage)
        {
            // 尝试从 FileStorage 以兼容的方式读取 Scale 字段（支持 "Scale" 或 "scale"，并兼容多种数值类型）
            try
            {
                if (fileStorage == null) return VariableDictionary.blockScale;

                var t = fileStorage.GetType();
                var prop = t.GetProperty("Scale") ?? t.GetProperty("scale");
                if (prop != null)
                {
                    var val = prop.GetValue(fileStorage);
                    if (val == null) return VariableDictionary.blockScale;

                    if (val is double d) return d;
                    if (val is float f) return Convert.ToDouble(f);
                    if (val is decimal dec) return Convert.ToDouble(dec);
                    if (val is int iv) return Convert.ToDouble(iv);
                    if (double.TryParse(val.ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double parsed))
                        return parsed;
                }
            }
            catch
            {
                // 忽略，返回默认
            }
            return VariableDictionary.blockScale;
        }

        #region 文件按钮点击与拖拽处理

        // 新增字段：用于按钮拖拽检测
        private System.Windows.Point _buttonDragStartPoint;
        private bool _isButtonMouseDown = false;
        private bool _isButtonDragging = false;
        private Button? _dragSourceButton = null;
        // 添加到 WpfMainWindow 类的字段区
        private System.Windows.Controls.Button? _lastSelectedDynamicButton = null;
        private readonly Dictionary<System.Windows.Controls.Button, System.Windows.Media.Brush> _originalButtonBackgrounds
            = new Dictionary<System.Windows.Controls.Button, System.Windows.Media.Brush>();


        /// <summary>
        /// 创建文件按钮
        /// </summary>
        private Button CreateFileButton(FileStorage file)
        {
            // 从 FileStorage 获取数据库中的"显示名称"字段，回退到空字符串以避免 null 引发异常
            string buttonText = file?.DisplayName ?? string.Empty;
           
            Button btn = new Button
            {
                Content = buttonText,
                Width = 88,
                Height = 22,
                Margin = new Thickness(0, 0, 5, 0),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                VerticalAlignment = System.Windows.VerticalAlignment.Top,
                Tag = new ButtonTagCommandInfo
                {
                    Type = "FileStorage",
                    ButtonName = buttonText,
                    fileStorage = file
                },
                // 默认背景设为 LightBlue（确保视觉一致）
                Background = System.Windows.Media.Brushes.Azure
            };

            btn.Click += DynamicButton_Click;
            // 双击检测（保留）
            btn.PreviewMouseLeftButtonDown += DynamicButton_PreviewMouseLeftButtonDown;
            // 增加拖拽事件
            btn.PreviewMouseMove += DynamicButton_PreviewMouseMove;
            // 结束拖拽检测
            btn.PreviewMouseLeftButtonUp += DynamicButton_PreviewMouseLeftButtonUp;
            
            return btn;
        }

        /// <summary>
        /// 修改 DynamicButton_Click：在数据库模式下将文件缓存到本地并确保使用真实文件名
        /// </summary>
        private void DynamicButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 检查发送者是否为按钮
                if (sender is Button btn)
                {
                    // 处理按钮高亮显示（仅限数据库模式的动态按钮）
                    if (sender is System.Windows.Controls.Button clickedBtn)
                    {
                        // 如果有上一个选中且不是当前，恢复它的原始背景
                        if (_lastSelectedDynamicButton != null && _lastSelectedDynamicButton != clickedBtn)
                        {
                            if (_originalButtonBackgrounds.TryGetValue(_lastSelectedDynamicButton, out var prevBrush))
                            {
                                _lastSelectedDynamicButton.Background = prevBrush;
                            }
                            else
                            {
                                // 没有备份则设置为样式默认（清空以让模板/样式生效）
                                _lastSelectedDynamicButton.ClearValue(System.Windows.Controls.Control.BackgroundProperty);
                            }
                        }

                        // 记录当前按钮原始背景（第一次点击时）
                        if (!_originalButtonBackgrounds.ContainsKey(clickedBtn))
                        {
                            var current = clickedBtn.Background;
                            //_originalButtonBackgrounds[clickedBtn] = current ?? System.Windows.SystemColors.ControlBrush;
                            _originalButtonBackgrounds[clickedBtn] = current ?? System.Windows.SystemColors.ControlBrush;
                        }

                        // 设置当前按钮高亮
                        clickedBtn.Background = System.Windows.Media.Brushes.LightGoldenrodYellow;

                        // 更新上一次引用
                        _lastSelectedDynamicButton = clickedBtn;
                    }
                    FileStorage fileStorage = null; // 获取按钮的 Tag 属性 用于后续处理的文件存储对象
                    LogManager.Instance.LogInfo($"单次点击了按钮: {btn.Content} ");
                    if (_useDatabaseMode && btn.Tag is ButtonTagCommandInfo tagInfo)// 数据库模式
                    {
                        // 数据库模式：处理数据库图元
                        fileStorage = tagInfo.fileStorage;
                        LogManager.Instance.LogInfo($"点击了数据库图元按钮: {tagInfo.ButtonName}");

                        // 缓存到本地：使用真实文件名 fileStorage.FileName（带扩展）
                        try
                        {
                            if (fileStorage != null)
                            {
                                string localDir = _cadStoragePath;// 本地缓存目录
                                if (string.IsNullOrEmpty(localDir))
                                {
                                    localDir = Path.Combine(AppPath, "CadFiles");// 默认缓存目录
                                    _cadStoragePath = localDir;// 保存到默认缓存目录
                                }
                                if (!Directory.Exists(localDir)) Directory.CreateDirectory(localDir);// 创建目录 确保目录存在

                                string sourcePath = fileStorage.FilePath ?? string.Empty;// 源文件路径
                                string sourceExt = Path.GetExtension(sourcePath);// 源文件扩展名
                                if (string.IsNullOrEmpty(sourceExt)) sourceExt = ".dwg";// 源文件没有扩展名则默认为 .dwg

                                string realName = fileStorage.FileName ?? Path.GetFileNameWithoutExtension(sourcePath);// 获取真实文件名
                                if (!realName.EndsWith(sourceExt, StringComparison.OrdinalIgnoreCase))// 真实文件名没有扩展名则添加源文件扩展名
                                {
                                    realName = Path.GetFileNameWithoutExtension(realName) + sourceExt;// 添加源文件扩展名
                                }

                                string localPath = Path.Combine(localDir, realName);// 本地缓存文件路径

                                if (!File.Exists(localPath))// 本地缓存文件不存在 本地缓存不存在则复制文件到本地
                                {
                                    if (!string.IsNullOrEmpty(sourcePath) && File.Exists(sourcePath))// 源文件存在 源文件存在 则复制文件到本地缓存
                                    {
                                        File.Copy(sourcePath, localPath, true);// 复制文件 复制文件到本地缓存
                                        LogManager.Instance.LogInfo($"已将图元缓存到本地: {localPath}");// 记录日志
                                    }
                                    else
                                    {
                                        LogManager.Instance.LogWarning($"无法找到源文件，未能缓存: {sourcePath}");
                                    }
                                }
                                else
                                {
                                    LogManager.Instance.LogInfo($"本地缓存已存在: {localPath}");
                                }

                                // 更新路径，后续预览/插入使用本地缓存
                                if (File.Exists(localPath))
                                {
                                    fileStorage.FilePath = localPath;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            LogManager.Instance.LogWarning($"缓存文件到本地失败: {ex.Message}");
                        }

                        // 记录为当前选中文件，供“应用图元”使用
                        _currentFileStorage = fileStorage;
                        _selectedFileStorage = fileStorage;

                        // 显示预览图片（使用可能已更新的本地路径）
                        ShowFilePreview(fileStorage);

                        // 显示文件详细属性（使用PropertiesDataGrid）
                        DisplayFilePropertiesInDataGridAsync(fileStorage);
                    }
                    else if (!_useDatabaseMode && btn.Tag is string filePath)
                    {
                        // Resources模式：处理文件路径
                        LogManager.Instance.LogInfo($"点击了Resources图元按钮: {filePath}");

                        // 从文件路径提取按钮名称
                        string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
                        string buttonName = fileNameWithoutExt;
                        if (fileNameWithoutExt.Contains("_"))
                        {
                            buttonName = fileNameWithoutExt.Substring(fileNameWithoutExt.IndexOf("_") + 1);
                        }

                        // 构造临时 FileStorage 并设置为当前选中
                        fileStorage = new FileStorage
                        {
                            FilePath = filePath,
                            FileName = Path.GetFileNameWithoutExtension(filePath),
                            DisplayName = Path.GetFileName(filePath)
                        };
                        _currentFileStorage = fileStorage;
                        _selectedFileStorage = fileStorage;

                        // 显示预览图片
                        ShowPreviewImage(filePath, buttonName);

                        // 清空属性显示
                        ClearFilePropertiesInDataGrid();

                        // 执行原有操作
                        ExecuteDynamicButtonActionFromResources(buttonName, filePath);
                    }
                    else if (btn.Tag is FileStorage directFileStorage)
                    {
                        // 直接的FileStorage对象
                        fileStorage = directFileStorage;
                        LogManager.Instance.LogInfo($"点击了文件按钮: {fileStorage.DisplayName}");

                        // 记录为当前选中文件，供“应用图元”使用
                        _currentFileStorage = fileStorage;
                        _selectedFileStorage = fileStorage;

                        // 尝试缓存到本地（同上逻辑）
                        try
                        {
                            if (fileStorage != null)
                            {
                                string localDir = _cadStoragePath;
                                if (string.IsNullOrEmpty(localDir))
                                {
                                    localDir = Path.Combine(AppPath, "CadFiles");
                                    _cadStoragePath = localDir;
                                }
                                if (!Directory.Exists(localDir)) Directory.CreateDirectory(localDir);

                                string sourcePath = fileStorage.FilePath ?? string.Empty;
                                string sourceExt = Path.GetExtension(sourcePath);
                                if (string.IsNullOrEmpty(sourceExt)) sourceExt = ".dwg";

                                string realName = fileStorage.FileName ?? Path.GetFileNameWithoutExtension(sourcePath);
                                if (!realName.EndsWith(sourceExt, StringComparison.OrdinalIgnoreCase))
                                {
                                    realName = Path.GetFileNameWithoutExtension(realName) + sourceExt;
                                }

                                string localPath = Path.Combine(localDir, realName);

                                if (!File.Exists(localPath))
                                {
                                    if (!string.IsNullOrEmpty(sourcePath) && File.Exists(sourcePath))
                                    {
                                        File.Copy(sourcePath, localPath, true);
                                        LogManager.Instance.LogInfo($"已将图元缓存到本地: {localPath}");
                                    }
                                }

                                if (File.Exists(localPath))
                                {
                                    fileStorage.FilePath = localPath;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            LogManager.Instance.LogWarning($"缓存文件到本地失败: {ex.Message}");
                        }

                        // 显示预览图片
                        ShowFilePreview(fileStorage);

                        // 显示文件详细属性（使用PropertiesDataGrid）
                        _ = DisplayFilePropertiesInDataGridAsync(fileStorage);
                    }
                    else
                    {
                        LogManager.Instance.LogWarning("按钮点击事件处理失败：无法识别的数据类型");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"处理按钮点击事件时出错: {ex.Message}");
                System.Windows.MessageBox.Show($"处理按钮点击事件时出错: {ex.Message}");
            }


            //try
            //{
            //    // 检查发送者是否为按钮
            //    if (sender is Button btn)
            //    {
            //        FileStorage fileStorage = null;// 获取按钮的 Tag 属性 用于后续处理的文件存储对象
            //        LogManager.Instance.LogInfo($"单次点击了按钮: {btn.Content} ");
            //        if (_useDatabaseMode && btn.Tag is ButtonTagCommandInfo tagInfo)// 数据库模式
            //        {
            //            // 数据库模式：处理数据库图元
            //            fileStorage = tagInfo.fileStorage;
            //            LogManager.Instance.LogInfo($"点击了数据库图元按钮: {tagInfo.ButtonName}");

            //            // 缓存到本地：使用真实文件名 fileStorage.FileName（带扩展）
            //            try
            //            {
            //                if (fileStorage != null)
            //                {
            //                    string localDir = _cadStoragePath;// 本地缓存目录
            //                    if (string.IsNullOrEmpty(localDir))
            //                    {
            //                        localDir = Path.Combine(AppPath, "CadFiles");// 默认缓存目录
            //                        _cadStoragePath = localDir;// 保存到默认缓存目录
            //                    }
            //                    if (!Directory.Exists(localDir)) Directory.CreateDirectory(localDir);// 创建目录 确保目录存在

            //                    string sourcePath = fileStorage.FilePath ?? string.Empty;// 源文件路径
            //                    string sourceExt = Path.GetExtension(sourcePath);// 源文件扩展名
            //                    if (string.IsNullOrEmpty(sourceExt)) sourceExt = ".dwg";// 源文件没有扩展名则默认为 .dwg

            //                    string realName = fileStorage.FileName ?? Path.GetFileNameWithoutExtension(sourcePath);// 获取真实文件名
            //                    if (!realName.EndsWith(sourceExt, StringComparison.OrdinalIgnoreCase))// 真实文件名没有扩展名则添加源文件扩展名
            //                    {
            //                        realName = Path.GetFileNameWithoutExtension(realName) + sourceExt;// 添加源文件扩展名
            //                    }

            //                    string localPath = Path.Combine(localDir, realName);// 本地缓存文件路径

            //                    if (!File.Exists(localPath))// 本地缓存文件不存在 本地缓存不存在则复制文件到本地
            //                    {
            //                        if (!string.IsNullOrEmpty(sourcePath) && File.Exists(sourcePath))// 源文件存在 源文件存在 则复制文件到本地缓存
            //                        {
            //                            File.Copy(sourcePath, localPath, true);// 复制文件 复制文件到本地缓存
            //                            LogManager.Instance.LogInfo($"已将图元缓存到本地: {localPath}");// 记录日志
            //                        }
            //                        else
            //                        {
            //                            LogManager.Instance.LogWarning($"无法找到源文件，未能缓存: {sourcePath}");
            //                        }
            //                    }
            //                    else
            //                    {
            //                        LogManager.Instance.LogInfo($"本地缓存已存在: {localPath}");
            //                    }

            //                    // 更新路径，后续预览/插入使用本地缓存
            //                    if (File.Exists(localPath))
            //                    {
            //                        fileStorage.FilePath = localPath;
            //                    }
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                LogManager.Instance.LogWarning($"缓存文件到本地失败: {ex.Message}");
            //            }

            //            // 显示预览图片（使用可能已更新的本地路径）
            //            ShowFilePreview(fileStorage);

            //            // 显示文件详细属性（使用PropertiesDataGrid）
            //            DisplayFilePropertiesInDataGridAsync(fileStorage);

            //        }
            //        else if (!_useDatabaseMode && btn.Tag is string filePath)
            //        {
            //            // Resources模式：处理文件路径
            //            LogManager.Instance.LogInfo($"点击了Resources图元按钮: {filePath}");

            //            // 从文件路径提取按钮名称
            //            string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
            //            string buttonName = fileNameWithoutExt;
            //            if (fileNameWithoutExt.Contains("_"))
            //            {
            //                buttonName = fileNameWithoutExt.Substring(fileNameWithoutExt.IndexOf("_") + 1);
            //            }

            //            // 显示预览图片
            //            ShowPreviewImage(filePath, buttonName);

            //            // 清空属性显示
            //            ClearFilePropertiesInDataGrid();

            //            // 执行原有操作
            //            ExecuteDynamicButtonActionFromResources(buttonName, filePath);
            //        }
            //        else if (btn.Tag is FileStorage directFileStorage)
            //        {
            //            // 直接的FileStorage对象
            //            fileStorage = directFileStorage;
            //            LogManager.Instance.LogInfo($"点击了文件按钮: {fileStorage.DisplayName}");

            //            // 尝试缓存到本地（同上逻辑）
            //            try
            //            {
            //                if (fileStorage != null)
            //                {
            //                    string localDir = _cadStoragePath;
            //                    if (string.IsNullOrEmpty(localDir))
            //                    {
            //                        localDir = Path.Combine(AppPath, "CadFiles");
            //                        _cadStoragePath = localDir;
            //                    }
            //                    if (!Directory.Exists(localDir)) Directory.CreateDirectory(localDir);

            //                    string sourcePath = fileStorage.FilePath ?? string.Empty;
            //                    string sourceExt = Path.GetExtension(sourcePath);
            //                    if (string.IsNullOrEmpty(sourceExt)) sourceExt = ".dwg";

            //                    string realName = fileStorage.FileName ?? Path.GetFileNameWithoutExtension(sourcePath);
            //                    if (!realName.EndsWith(sourceExt, StringComparison.OrdinalIgnoreCase))
            //                    {
            //                        realName = Path.GetFileNameWithoutExtension(realName) + sourceExt;
            //                    }

            //                    string localPath = Path.Combine(localDir, realName);

            //                    if (!File.Exists(localPath))
            //                    {
            //                        if (!string.IsNullOrEmpty(sourcePath) && File.Exists(sourcePath))
            //                        {
            //                            File.Copy(sourcePath, localPath, true);
            //                            LogManager.Instance.LogInfo($"已将图元缓存到本地: {localPath}");
            //                        }
            //                    }

            //                    if (File.Exists(localPath))
            //                    {
            //                        fileStorage.FilePath = localPath;
            //                    }
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                LogManager.Instance.LogWarning($"缓存文件到本地失败: {ex.Message}");
            //            }

            //            // 显示预览图片
            //            ShowFilePreview(fileStorage);

            //            // 显示文件详细属性（使用PropertiesDataGrid）
            //            _ = DisplayFilePropertiesInDataGridAsync(fileStorage);

            //        }
            //        else
            //        {
            //            LogManager.Instance.LogWarning("按钮点击事件处理失败：无法识别的数据类型");
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    LogManager.Instance.LogError($"处理按钮点击事件时出错: {ex.Message}");
            //    System.Windows.MessageBox.Show($"处理按钮点击事件时出错: {ex.Message}");
            //}
        }

        /// <summary>
        /// PreviewMouseLeftButtonDown 事件的统一处理器，用于检测双击（ClickCount==2）
        /// </summary>
        private void DynamicButton_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                // 记录按下状态（用于拖拽判断）
                _isButtonMouseDown = true;
                _isButtonDragging = false;
                _buttonDragStartPoint = e.GetPosition(this);
                _dragSourceButton = sender as Button;

                // 双击逻辑仍然保留
                if (e.ClickCount == 2)
                {
                    DynamicButton_MouseDoubleClick(sender, e);
                    e.Handled = true;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"PreviewMouseLeftButtonDown 处理失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 统一处理动态生成按钮的双击（使用 PreviewMouseLeftButtonDown 判断双击）
        /// 双击时仅显示详情（预览 + 属性），不直接执行插入命令
        /// </summary>
        private async void DynamicButton_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (sender is Button btn)
                {
                    // 获取当前绘图比例（优先使用用户在 TextBox_绘图比例 中设置的值）
                    VariableDictionary.textBoxScale = GetDrawingScaleFromTextBox();
                    // 记录双击日志
                    LogManager.Instance.LogInfo($"双击了按钮: {btn.Content}");

                    // 数据库模式：ButtonTagCommandInfo 中带 fileStorage
                    if (_useDatabaseMode && btn.Tag is ButtonTagCommandInfo tagInfo && tagInfo.fileStorage != null)
                    {
                        // 获取文件存储对象
                        var fileStorage = tagInfo.fileStorage;
                        LogManager.Instance.LogInfo($"双击了数据库图元按钮: {tagInfo.ButtonName}");
                        // 在双击处理数据库图元时，准备参数并调用统一插入方法
                        // 假设 fileStorage 包含 LocalPath、BlockName、LayerName、Scale 等属性
                        VariableDictionary.entityRotateAngle = 0;
                        string localDwg = fileStorage.FilePath; // 确保存在，或先把数据库的 bytes 写入到此路径
                        VariableDictionary.btnFileName = fileStorage.FileName ?? fileStorage.DisplayName ?? System.IO.Path.GetFileNameWithoutExtension(localDwg);
                        VariableDictionary.btnFileName_blockName = fileStorage.BlockName; // 可选
                        VariableDictionary.btnBlockLayer = fileStorage.LayerName;
                        VariableDictionary.layerColorIndex = Convert.ToInt32(fileStorage.ColorIndex);
                        var isTianzheng = fileStorage.IsTianZheng ; // 可选，视具体需求而定
                        // 优先使用用户在TextBox_绘图比例中设置的比例值
                        VariableDictionary.textBoxScale = VariableDictionary.textBoxScale;
                        if (VariableDictionary.textBoxScale <= 0) // 如果获取失败，使用原有逻辑
                        {
                            AutoCadHelper.GetAndApplyActiveDrawingScale();//获取当前绘图比例
                            VariableDictionary.textBoxScale = VariableDictionary.blockScale;
                        }
                        //string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
                        //System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
                        // 调用统一插入方法（交互放置）
                        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        if (doc != null)
                        {
                            // 在 document lock 内调用 CAD API，保证线程安全和锁定语义
                            using (doc.LockDocument())
                            {
                                //Command.InsertBlockFromResource(localDwg, fileName, blockRecordName, layer, layerColorIndex, scale, interactive: true,isTianzheng);                                
                                InsertGraphicHelper.CopyDwgAllFast(localDwg);
                            }
                        }
                        else
                        {
                            LogManager.Instance.LogWarning("未找到活动文档，无法插入图块。");
                        }
                    }
                    // 资源模式或文件路径存储在 ButtonTagCommandInfo.FilePath
                    else if (btn.Tag is ButtonTagCommandInfo tagWithPath && !string.IsNullOrEmpty(tagWithPath.FilePath))
                    {
                        string filePath = tagWithPath.FilePath;
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                        string buttonName = fileNameWithoutExt.Contains("_")
                            ? fileNameWithoutExt.Substring(fileNameWithoutExt.IndexOf("_") + 1)
                            : fileNameWithoutExt;

                        LogManager.Instance.LogInfo($"双击了Resources图元按钮: {filePath}");
                        ShowPreviewImage(filePath, buttonName);
                        DisplayFileInfo(filePath);
                        ClearFilePropertiesInDataGrid();
                    }
                    // 资源模式直接把路径存在 Tag 为 string
                    else if (btn.Tag is string filePathString)
                    {
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePathString);
                        string buttonName = fileNameWithoutExt.Contains("_")
                            ? fileNameWithoutExt.Substring(fileNameWithoutExt.IndexOf("_") + 1)
                            : fileNameWithoutExt;

                        LogManager.Instance.LogInfo($"双击了Resources图元按钮 (string tag): {filePathString}");
                        ShowPreviewImage(filePathString, buttonName);
                        DisplayFileInfo(filePathString);
                        ClearFilePropertiesInDataGrid();
                    }
                    else
                    {
                        LogManager.Instance.LogWarning("双击事件：无法识别按钮的 Tag 类型");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"处理按钮双击时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 鼠标移动：判定为拖拽并触发插入操作（交互式）
        /// </summary>
        private void DynamicButton_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            // 拖拽判断（保持现有逻辑，仅在真正开始拖拽时触发插入）
            var btn = sender as Button;
            if (btn == null) return;
            if (!_isButtonMouseDown) return;

            var currentPos = e.GetPosition(null);
            if (!_isButtonDragging)
            {
                if (Math.Abs(currentPos.X - _buttonDragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance &&
                    Math.Abs(currentPos.Y - _buttonDragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
                {
                    return;
                }
                _isButtonDragging = true;
                _dragSourceButton = btn;
            }

            // 一旦拖拽被认定为开始，尝试从按钮 Tag 中获取要插入的 DWG 路径或字节并调用 CopyDwgAllFast
            try
            {
                var tagInfo = btn.Tag as ButtonTagCommandInfo;
                string? tempPath = null;

                // 优先使用已有的文件路径
                if (tagInfo != null)
                {
                    if (!string.IsNullOrEmpty(tagInfo.FilePath) && System.IO.File.Exists(tagInfo.FilePath))
                    {
                        tempPath = tagInfo.FilePath;
                    }
                    else if (tagInfo.fileStorage != null)
                    {
                        // 如果 FileStorage 已包含物理路径，直接使用
                        if (!string.IsNullOrEmpty(tagInfo.fileStorage.FilePath) && System.IO.File.Exists(tagInfo.fileStorage.FilePath))
                        {
                            tempPath = tagInfo.fileStorage.FilePath;
                        }
                        // 否则若有字节流（常见场景：资源以字节形式存储），写入临时文件
                        else if (tagInfo.fileStorage is var fs && fs != null)
                        {
                            // 约定 FileStorage 可能包含字段/属性名为 FileBytes 或 Content（根据项目实际调整）
                            // 下面尝试从 Parameters 或 fileStorage 的常见成员读取字节数组
                            if (tagInfo.Parameters != null && tagInfo.Parameters.TryGetValue("resourcesFile", out object? obj) && obj is byte[] bytesFromParam)
                            {
                                tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{tagInfo.FileName ?? "tmp"}_{Guid.NewGuid():N}.dwg");
                                System.IO.File.WriteAllBytes(tempPath, bytesFromParam);
                            }
                            else
                            {
                                // 兜底尝试：如果 FileStorage 有可访问的字节属性，尝试反射读取常见字段名
                                var fileStorageType = fs.GetType();
                                var prop = fileStorageType.GetProperty("FileBytes") ?? fileStorageType.GetProperty("Content") ?? fileStorageType.GetProperty("Bytes");
                                if (prop != null)
                                {
                                    var val = prop.GetValue(fs) as byte[];
                                    if (val != null && val.Length > 0)
                                    {
                                        tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{tagInfo.FileName ?? "tmp"}_{Guid.NewGuid():N}.dwg");
                                        System.IO.File.WriteAllBytes(tempPath, val);
                                    }
                                }
                            }
                        }
                    }
                    else if (tagInfo.Parameters != null && tagInfo.Parameters.TryGetValue("resourcesFile", out object? resObj) && resObj is byte[] resBytes)
                    {
                        tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{tagInfo.FileName ?? "tmp"}_{Guid.NewGuid():N}.dwg");
                        System.IO.File.WriteAllBytes(tempPath, resBytes);
                    }
                }

                if (string.IsNullOrEmpty(tempPath))
                {
                    // 无法获取临时 DWG 文件，直接返回，不抛异常
                    return;
                }

                // 调用新的插入方法（此方法内部处理事务与文档加锁）
                InsertGraphicHelper.CopyDwgAllFast(tempPath);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"插入图元失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // 拖拽完成后不自动重置鼠标按下状态，保持原有行为；若需要可在此处重置 _isButtonDragging/_isButtonMouseDown 等
                _isButtonDragging = false;
                _dragSourceButton = null;
            }
        }

        /// <summary>
        /// 鼠标抬起：清理拖拽标记
        /// </summary>
        private void DynamicButton_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                _isButtonMouseDown = false;
                _isButtonDragging = false;
                _dragSourceButton = null;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"PreviewMouseLeftButtonUp 处理失败: {ex.Message}");
            }
        }
               
        #endregion


        #region 图元tabItem

        /// <summary>
        /// TabControl选择改变事件
        /// </summary>
        /// <param Name="sender"></param>
        /// <param Name="e"></param>
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                LogManager.Instance.LogInfo("TabControl选择改变事件触发");

                // 获取当前选中的TabItem
                if (e.AddedItems.Count > 0 && e.AddedItems[0] is TabItem selectedTab)
                {
                    string header = selectedTab.Header.ToString().Trim();
                    LogManager.Instance.LogInfo($"选中的TabItem: {header}");

                    // 处理主分类TabItem（工艺、建筑、结构等）
                    if (header == "工艺" || header == "建筑" || header == "结构" ||
                        header == "电气" || header == "给排水" || header == "暖通" ||
                        header == "自控" || header == "总图" || header == "公共图")
                    {
                        LogManager.Instance.LogInfo($"处理主分类TabItem: {header}");
                        LoadButtonsForMainCategoryTab(selectedTab, header);

                        // 特殊处理工艺分类
                        if (header == "工艺")
                        {
                            LoadConditionButtons();
                        }
                    }
                    // 处理嵌套的TabItem（图元集、图层管理等）
                    else if (header.Contains("图元集") || header.Contains("图层管理"))
                    {
                        LogManager.Instance.LogInfo($"处理嵌套TabItem: {header}");
                        TabItem parentTabItem = FindParentTabItem(selectedTab);
                        if (parentTabItem != null)
                        {
                            string parentHeader = parentTabItem.Header.ToString().Trim();
                            LogManager.Instance.LogInfo($"父级TabItem: {parentHeader}");
                            LoadButtonsForMainCategoryTab(parentTabItem, parentHeader);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"处理TabControl选择改变时出错: {ex.Message}");
                //LogManager.Instance.LogError($"堆栈跟踪: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 为指定的主分类TabItem加载按钮
        /// </summary>
        private void LoadButtonsForMainCategoryTab(TabItem tabItem, string categoryName)
        {
            try
            {
                LogManager.Instance.LogInfo($"开始为分类 {categoryName} 加载按钮");

                // 查找对应的面板
                WrapPanel panel = GetPanelByFolderName(categoryName);
                if (panel == null)
                {
                    LogManager.Instance.LogInfo($"未找到 {categoryName} 对应的面板");
                    return;
                }

                // 清空面板内容
                panel.Children.Clear();

                // 检查数据库是否可用
                if (_databaseManager != null && _databaseManager.IsDatabaseAvailable)
                {
                    LogManager.Instance.LogInfo($"使用数据库模式加载 {categoryName}");
                    LoadButtonsFromDatabaseForCategory(categoryName, panel);
                }
                else
                {
                    LogManager.Instance.LogInfo($"使用Resources文件夹模式加载 {categoryName}");
                    LoadButtonsFromResources(categoryName, panel);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"为分类 {categoryName} 加载按钮时出错: {ex.Message}");
                LogManager.Instance.LogInfo($"堆栈跟踪: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 从数据库为指定分类加载按钮
        /// </summary>
        private async void LoadButtonsFromDatabaseForCategory(string categoryName, WrapPanel panel)
        {
            try
            {
                LogManager.Instance.LogInfo($"=== 开始从数据库加载分类 {categoryName} ===");

                if (_databaseManager == null)
                {
                    LogManager.Instance.LogInfo("数据库管理器为空");
                    LoadButtonsFromResources(categoryName, panel);
                    return;
                }

                if (!_databaseManager.IsDatabaseAvailable)
                {
                    LogManager.Instance.LogInfo("数据库连接不可用");
                    LoadButtonsFromResources(categoryName, panel);
                    return;
                }

                // 获取主分类
                var category = await _databaseManager.GetCadCategoryByNameAsync(categoryName);
                if (category == null)
                {
                    LogManager.Instance.LogInfo($"数据库中未找到分类: {categoryName}");
                    LoadButtonsFromResources(categoryName, panel);
                    return;
                }

                LogManager.Instance.LogInfo($"找到主分类: {category.DisplayName} (ID: {category.Id})");

                // 获取子分类
                var subcategories = await _databaseManager.GetCadSubcategoriesByCategoryIdAsync(category.Id);
                LogManager.Instance.LogInfo($"找到 {subcategories.Count} 个子分类");

                // 清空面板
                panel.Children.Clear();

                if (subcategories.Count == 0)
                {
                    // 没有子分类，直接加载该分类下的文件
                    await LoadFilesDirectlyForCategory(category, panel);
                }
                else
                {
                    // 有子分类，按子分类组织文件
                    await LoadFilesBySubcategories(category, subcategories, panel);
                }

                LogManager.Instance.LogInfo($"=== 完成加载分类 {categoryName} ===");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"从数据库加载分类 {categoryName} 时出错: {ex.Message}");
                LogManager.Instance.LogInfo($"堆栈跟踪: {ex.StackTrace}");
                LoadButtonsFromResources(categoryName, panel);
            }
        }

        /// <summary>
        /// 直接为分类加载文件（无子分类情况）
        /// </summary>
        private async Task LoadFilesDirectlyForCategory(CadCategory category, WrapPanel panel)
        {
            try
            {
                LogManager.Instance.LogInfo($"直接加载分类 {category.DisplayName} 下的文件");

                // 获取该分类下的所有文件
                var files = await _databaseManager.GetFilesByCategoryIdAsync(category.Id, "main");
                LogManager.Instance.LogInfo($"在分类 {category.DisplayName} 中找到 {files.Count} 个文件");

                if (files.Count > 0)
                {
                    // 按显示名称排序
                    var sortedFiles = files.OrderBy(f => f.DisplayName).ToList();

                    // 创建文件显示区域
                    CreateFileButtonsForPanel(sortedFiles, panel, category.DisplayName);
                }
                else
                {
                    ShowNoFilesMessage(panel, "暂无文件");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"直接加载分类文件时出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 按子分类加载文件
        /// </summary>
        private async Task LoadFilesBySubcategories(CadCategory category,
            List<CadSubcategory> subcategories, WrapPanel panel)
        {
            try
            {
                LogManager.Instance.LogInfo($"按子分类加载分类 {category.DisplayName} 下的文件");

                // 定义背景色
                List<System.Windows.Media.Color> backgroundColors = new List<System.Windows.Media.Color>
                {
                    Colors.FloralWhite, Colors.Azure, Colors.FloralWhite, Colors.Azure
                };

                int colorIndex = 0;

                // 遍历子分类
                foreach (var subcategory in subcategories.OrderBy(s => s.SortOrder))
                {
                    LogManager.Instance.LogInfo($"处理子分类: {subcategory.DisplayName} (ID: {subcategory.Id})");

                    // 获取子分类下的文件
                    var files = await _databaseManager.GetFilesByCategoryIdAsync(subcategory.Id, "sub");
                    LogManager.Instance.LogInfo($"在子分类 {subcategory.DisplayName} 中找到 {files.Count} 个文件");

                    // 创建子分类区域
                    Border sectionBorder = CreateSubcategorySection(
                        subcategory.DisplayName,
                        backgroundColors[colorIndex % backgroundColors.Count]);

                    StackPanel sectionPanel = sectionBorder.Child as StackPanel;

                    if (files.Count > 0)
                    {
                        // 按显示名称排序
                        var sortedFiles = files.OrderBy(fileStorage => fileStorage.DisplayName).ToList();
                        CreateFileButtonsForPanel(sortedFiles, sectionPanel, subcategory.DisplayName);
                    }
                    else
                    {
                        ShowNoFilesMessage(sectionPanel, "暂无文件");
                    }

                    panel.Children.Add(sectionBorder);
                    colorIndex++;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"按子分类加载文件时出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 创建子分类区域
        /// </summary>
        private Border CreateSubcategorySection(string title, System.Windows.Media.Color backgroundColor)
        {
            Border sectionBorder = new Border
            {
                BorderBrush = new SolidColorBrush(Colors.Gray),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(5),
                Margin = new Thickness(0, 2, 0, 2),
                Width = 300,
                Background = new SolidColorBrush(backgroundColor),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left
            };

            StackPanel sectionPanel = new StackPanel
            {
                Margin = new Thickness(3)
            };

            TextBlock sectionHeader = new TextBlock
            {
                Text = title,
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 2),
                Foreground = new SolidColorBrush(Colors.DarkBlue)
            };

            sectionPanel.Children.Add(sectionHeader);
            sectionBorder.Child = sectionPanel;

            return sectionBorder;
        }

        /// <summary>
        /// 找到分类文件后为面板创建文件按钮
        /// </summary>
        private void CreateFileButtonsForPanel(List<FileStorage> files, Panel targetPanel, string sectionName)
        {
            try
            {
                LogManager.Instance.LogInfo($"为 {sectionName} 创建 {files.Count} 个文件按钮");

                // 按3列分组
                int columns = 3;
                for (int i = 0; i < files.Count; i += columns)
                {
                    StackPanel rowPanel = new StackPanel
                    {
                        Orientation = System.Windows.Controls.Orientation.Horizontal,
                        Margin = new Thickness(0, 0, 0, 2)
                    };

                    // 创建一行按钮（最多3个）
                    for (int j = 0; j < columns && (i + j) < files.Count; j++)
                    {
                        var file = files[i + j];
                        Button btn = CreateFileButton(file);//创建文件按钮
                        rowPanel.Children.Add(btn);//添加按钮
                    }

                    targetPanel.Children.Add(rowPanel);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"创建文件按钮时出错: {ex.Message}");
            }
        }

        private void 还原初始值_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (PropertiesDataGrid == null)
                {
                    MessageBox.Show("未找到属性表。", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var items = PropertiesDataGrid.ItemsSource as List<CategoryPropertyEditModel>;
                if (items == null || items.Count == 0)
                {
                    MessageBox.Show("当前没有可还原的属性行。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (_propertiesSnapshotForInsert == null || _propertiesSnapshotForInsert.Count == 0)
                {
                    MessageBox.Show("未检测到初始快照，无法还原。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // 把每一行的 PropertyName 映射到 snapshot 中的值恢复回去
                foreach (var row in items)
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(row.PropertyName1))
                        {
                            var key1 = NormalizePropertyDisplayName(row.PropertyName1);
                            if (_propertiesSnapshotForInsert.TryGetValue(key1, out var val1))
                                row.PropertyValue1 = val1;
                        }
                        if (!string.IsNullOrWhiteSpace(row.PropertyName2))
                        {
                            var key2 = NormalizePropertyDisplayName(row.PropertyName2);
                            if (_propertiesSnapshotForInsert.TryGetValue(key2, out var val2))
                                row.PropertyValue2 = val2;
                        }
                    }
                    catch { /* 单行还原异常忽略，继续其它行 */ }
                }

                // 刷新 UI
                PropertiesDataGrid.Items.Refresh();
                MessageBox.Show("属性已还原为初始快照值。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"还原初始值失败: {ex.Message}");
                MessageBox.Show($"还原初始值失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void 应用图元_Click(object sender, RoutedEventArgs e)
        {
            // 将“应用图元”按钮的行为统一为调用 CopyDwgAllFast，
            // 根据当前选中文件优先选择 _selectedFileStorage，然后是 _currentFileStorage，再是 _selectedFilePath
            try
            {
                string? tempPath = null;

                // 优先：WPF 窗口中选中的 FileStorage（视实现而定）
                var fs = _selectedFileStorage ?? _currentFileStorage;
                if (fs != null)
                {
                    // 尝试直接使用已有物理路径
                    var propPath = fs.GetType().GetProperty("FilePath")?.GetValue(fs) as string;
                    if (!string.IsNullOrEmpty(propPath) && System.IO.File.Exists(propPath))
                    {
                        tempPath = propPath;
                    }
                    else
                    {
                        // 尝试从对象内提取字节并写临时文件（常见属性名）
                        var propBytes = fs.GetType().GetProperty("FileBytes")?.GetValue(fs) as byte[];
                        if (propBytes != null && propBytes.Length > 0)
                        {
                            tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{fs.GetType().Name}_{Guid.NewGuid():N}.dwg");
                            System.IO.File.WriteAllBytes(tempPath, propBytes);
                        }
                    }
                }

                // 其次：已记录的选中文件路径
                if (string.IsNullOrEmpty(tempPath) && !string.IsNullOrEmpty(_selectedFilePath) && System.IO.File.Exists(_selectedFilePath))
                {
                    tempPath = _selectedFilePath;
                }

                if (string.IsNullOrEmpty(tempPath))
                {
                    MessageBox.Show("未找到可插入的 DWG 文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // 调用统一插入方法
                InsertGraphicHelper.CopyDwgAllFast(tempPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用图元时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// 新增：为动态按钮统一附加鼠标事件和 click 处理器，避免重复绑定预定义按钮
        /// </summary>
        private static readonly DependencyProperty HandlersAttachedProperty = DependencyProperty.RegisterAttached("HandlersAttached", typeof(bool), typeof(WpfMainWindow), new PropertyMetadata(false));

        /// <summary>
        /// 为动态按钮附加鼠标事件和 click 处理器
        /// </summary>
        /// <param name="btn"></param>
        private void AttachDynamicButtonHandlers(Button btn)
        {
            if (btn == null) return;

            bool attached = false;
            try
            {
                attached = (bool)btn.GetValue(HandlersAttachedProperty);
            }
            catch { attached = false; }

            if (attached) return;

            // 如果该按钮是预定义按钮（Type=="Predefined"），不要替换它已有的 Click 处理器
            var tagInfo = btn.Tag as ButtonTagCommandInfo;
            bool isPredefined = tagInfo != null && string.Equals(tagInfo.Type, "Predefined", StringComparison.OrdinalIgnoreCase);

            // 只有非预定义按钮才绑定 DynamicButton_Click（资源模式很多按钮最初没绑定）
            if (!isPredefined)
            {
                // 仅在尚未绑定时附加 Click（保守策略）
                btn.Click += DynamicButton_Click;
                // 默认把动态生成的按钮背景设置为 Azure（仅在没有显式背景时）
                try
                {
                    var current = btn.Background;
                    bool isTransparentOrNull =
                        current == null ||
                        (current is System.Windows.Media.SolidColorBrush scb && scb.Color == System.Windows.Media.Colors.Transparent);

                    if (isTransparentOrNull)
                    {
                        btn.Background = System.Windows.Media.Brushes.Azure;
                    }
                    btn.Background = System.Windows.Media.Brushes.Azure;
                }
                catch
                {
                    // 忽略任何设置背景时的异常，不中断流程
                }
            }

            // 始终绑定预检鼠标事件以支持双击与拖拽
            btn.PreviewMouseLeftButtonDown += DynamicButton_PreviewMouseLeftButtonDown;
            btn.PreviewMouseMove += DynamicButton_PreviewMouseMove;
            btn.PreviewMouseLeftButtonUp += DynamicButton_PreviewMouseLeftButtonUp;
            btn.SetValue(HandlersAttachedProperty, true);
        }

        /// <summary>
        /// 显示文件信息
        /// </summary>
        /// <param name="filePath"></param>
        private void DisplayFileInfo(string filePath)
        {
            try
            {
                var fileInfo = new FileInfo(filePath);

                // 显示文件信息
                file_Path.Text = filePath;
                File_Name.Text = fileInfo.Name;
                File_Size.Text = $"{fileInfo.Length / 1024.0:F2} KB";
                File_Type.Text = fileInfo.Extension.ToLower();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"显示文件信息失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 添加显示文件存储信息的方法
        /// </summary>
        /// <param name="fileStorage"></param>
        public void DisplayFileStorageInfo(FileStorage fileStorage)
        {
            try
            {
                // 显示文件信息
                file_Path.Text = fileStorage.FilePath ?? "";
                File_Name.Text = fileStorage.DisplayName ?? fileStorage.FileName ?? "";
                File_Name.Text = FormatFileNameForDisplay(fileStorage.DisplayName ?? fileStorage.FileName ?? "");
                File_Size.Text = fileStorage.FileSize > 0 ? $"{fileStorage.FileSize / 1024.0:F2} KB" : "";
                File_Type.Text = fileStorage.FileType ?? "";
                view_File_Path.Text = fileStorage.PreviewImagePath ?? "无预览图片";
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"显示文件信息失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 显示预览图片
        /// </summary>
        /// <param name="imagePath"></param>
        private void DisplayPreviewImage(string imagePath)
        {
            try
            {
                view_File_Path.Text = imagePath;

                // 显示预览图片
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(imagePath);
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                ViewImage.Source = bitmap;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"显示预览图片失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 显示无文件消息
        /// </summary>
        private void ShowNoFilesMessage(Panel panel, string message)
        {
            TextBlock noFilesText = new TextBlock
            {
                Text = message,
                FontSize = 12,
                Margin = new Thickness(5, 0, 0, 3),
                Foreground = new SolidColorBrush(Colors.Gray)
            };
            panel.Children.Add(noFilesText);
        }

        /// <summary>
        /// 清空文件上传界面
        /// </summary>
        private void ClearFileUploadInterface()
        {
            // 清空所有输入框和显示
            file_Path.Text = "";
            File_Name.Text = "";
            File_Size.Text = "";
            File_Type.Text = "";
            view_File_Path.Text = "";
            ViewImage.Source = null;

            // 清空属性编辑网格
            CategoryPropertiesDataGrid.ItemsSource = null;

            // 重置字段
            _selectedFilePath = null;
            _selectedPreviewImagePath = null;
            _currentFileStorage = null;
            _currentFileAttribute = null;
            _selectedCategoryNode = null;
        }

        /// <summary>
        /// 查找TabItem的父级TabItem
        /// </summary>
        /// <param Name="tabItem"></param>
        /// <returns></returns>
        private TabItem FindParentTabItem(TabItem tabItem)
        {
            DependencyObject parent = VisualTreeHelper.GetParent(tabItem);//获取父级
            while (parent != null)
            {
                if (parent is TabItem parentTabItem)
                {
                    return parentTabItem;
                }
                parent = VisualTreeHelper.GetParent(parent);
            }
            return null;
        }

        /// <summary>
        /// 通过文件夹名称获取对应的面板引用
        /// </summary>
        /// <param Name="folderName"></param>
        /// <returns></returns>
        private WrapPanel GetPanelByFolderName(string folderName)
        {
            LogManager.Instance.LogInfo($"查找面板: {folderName}");

            WrapPanel panel = null;

            switch (folderName)
            {
                case "公用工具":
                    panel = PublicButtonsPanel;
                    break;
                case "工艺":
                    panel = CraftButtonsPanel;
                    break;
                case "建筑":
                    panel = ArchitectureButtonsPanel;
                    break;
                case "结构":
                    panel = StructureButtonsPanel;
                    break;
                case "电气":
                    panel = ElectricalButtonsPanel;
                    break;
                case "给排水":
                    panel = PlumbingButtonsPanel;
                    break;
                case "暖通":
                    panel = HVACButtonsPanel;
                    break;
                case "自控":
                    panel = ControlButtonsPanel;
                    break;
                case "总图":
                    panel = GeneralButtonsPanel;
                    break;
            }

            LogManager.Instance.LogInfo($"面板查找结果: {panel != null}");
            return panel;
        }

        /// <summary>
        /// 从数据库加载按钮（新方法）
        /// </summary>
        /// <param Name="folderName">分类名称</param>
        /// <param Name="panel">目标面板</param>
        private async Task LoadButtonsFromDatabase(string folderName, WrapPanel panel)
        {
            try
            {
                if (_databaseManager == null)
                {
                    LogManager.Instance.LogInfo("数据库管理器未初始化");
                    return;
                }

                LogManager.Instance.LogInfo($"开始从数据库加载分类 {folderName} 的按钮");

                // 从数据库获取主分类信息
                var category = await _databaseManager.GetCadCategoryByNameAsync(folderName);
                if (category == null)
                {
                    LogManager.Instance.LogInfo($"未找到分类: {folderName}");
                    return;
                }

                // 获取该分类下的所有子分类
                var subcategories = await _databaseManager.GetCadSubcategoriesByCategoryIdAsync(category.Id);
                LogManager.Instance.LogInfo($"找到 {subcategories.Count} 个子分类");

                // 定义背景色列表，用于区分不同区域
                List<System.Windows.Media.Color> backgroundColors = new List<System.Windows.Media.Color>
                {
                    Colors.FloralWhite,
                    Colors.Azure,
                    Colors.FloralWhite,
                    Colors.Azure,
                    Colors.FloralWhite,
                    Colors.Azure,
                    Colors.FloralWhite,
                    Colors.Azure,
                    Colors.FloralWhite,
                    Colors.Azure,
                    Colors.FloralWhite,
                };

                int colorIndex = 0;

                // 遍历所有子分类
                foreach (var subcategory in subcategories)
                {
                    LogManager.Instance.LogInfo($"处理子分类: {subcategory.DisplayName}");

                    // 为每个子分类创建一个带边框和背景色的区域
                    Border sectionBorder = new Border
                    {
                        BorderBrush = new SolidColorBrush(Colors.Gray),
                        BorderThickness = new Thickness(1),
                        CornerRadius = new CornerRadius(5),
                        Margin = new Thickness(0, 2, 0, 2),
                        Width = 300,
                        Background = new SolidColorBrush(backgroundColors[colorIndex % backgroundColors.Count]),
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Left
                    };

                    // 创建区域内容的StackPanel
                    StackPanel sectionPanel = new StackPanel
                    {
                        Margin = new Thickness(3)
                    };

                    // 添加区域标题
                    TextBlock sectionHeader = new TextBlock
                    {
                        Text = subcategory.DisplayName,
                        FontSize = 14,
                        FontWeight = FontWeights.Bold,
                        Margin = new Thickness(0, 0, 0, 2),
                        Foreground = new SolidColorBrush(Colors.DarkBlue)
                    };
                    sectionPanel.Children.Add(sectionHeader);

                    // 从数据库获取该子分类下的所有图元文件
                    var graphics = await _databaseManager.GetFileStorageBySubcategoryIdAsync(subcategory.Id);
                    LogManager.Instance.LogInfo($"在 {subcategory.DisplayName} 中找到 {graphics.Count} 个图元文件");


                    if (graphics.Count > 0)
                    {
                        // 按显示名称排序
                        graphics.Sort((x, y) => x.DisplayName.CompareTo(y.DisplayName));

                        // 按3列分组处理
                        int columns = 3;
                        for (int i = 0; i < graphics.Count; i += columns)
                        {
                            // 创建水平StackPanel用于放置一行按钮
                            StackPanel rowPanel = new StackPanel
                            {
                                Orientation = System.Windows.Controls.Orientation.Horizontal,
                                Margin = new Thickness(0, 0, 0, 2)
                            };

                            // 添加该行的按钮（最多3个）
                            for (int j = 0; j < columns && (i + j) < graphics.Count; j++)
                            {
                                var graphic = graphics[i + j];

                                // 检查是否是预定义的按钮
                                //var commandInfo = ButtonCommandMapper.GetCommandInfo(graphic.DisplayName);

                                // 仅显示最后一个下划线后的名称，例如 "DQTJ_EQUIP_潮湿插座" -> "潮湿插座"

                                string buttonName = graphic.DisplayName;
                                if (!string.IsNullOrWhiteSpace(buttonName))
                                {
                                    int lastUnderscore = buttonName.LastIndexOf('_');
                                    if (lastUnderscore >= 0 && lastUnderscore + 1 < buttonName.Length)
                                    {
                                        buttonName = buttonName.Substring(lastUnderscore + 1).Trim();
                                    }
                                    else
                                    {
                                        buttonName = buttonName.Trim();
                                    }
                                }
                                // 创建按钮
                                Button btn = new Button
                                {
                                    Content = buttonName,
                                    Width = 88,
                                    Height = 22,
                                    Margin = new Thickness(0, 0, 5, 0),
                                    HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                                    VerticalAlignment = System.Windows.VerticalAlignment.Top,
                                    Tag = graphic // 存储完整的图元信息
                                };

                                // 检查是否是预定义的按钮
                                if (UnifiedCommandManager.IsPredefinedCommand(buttonName))
                                {
                                    // 如果是预定义按钮
                                    btn.Tag = new ButtonTagCommandInfo
                                    {
                                        Type = "Predefined",
                                        ButtonName = buttonName,
                                        fileStorage = graphic
                                    };
                                    btn.Click += PredefinedButton_Click;
                                }
                                else
                                {
                                    // 如果是普通图元按钮，存储图元信息
                                    btn.Tag = new ButtonTagCommandInfo
                                    {
                                        Type = "FileStorage",
                                        ButtonName = buttonName,
                                        fileStorage = graphic
                                    };
                                    //btn.Click += DynamicButton_Click;
                                }
                                // 新增：统一附加预检/拖拽事件（并避免重复绑定）
                                AttachDynamicButtonHandlers(btn);

                                // 添加按钮到行面板
                                rowPanel.Children.Add(btn);

                            }

                            // 添加行面板到区域面板
                            sectionPanel.Children.Add(rowPanel);
                        }
                    }
                    else
                    {
                        // 如果该子分类没有文件，显示提示信息
                        TextBlock noFilesText = new TextBlock
                        {
                            Text = "暂无文件",
                            FontSize = 12,
                            Margin = new Thickness(5, 0, 0, 3),
                            Foreground = new SolidColorBrush(Colors.Gray)
                        };
                        sectionPanel.Children.Add(noFilesText);
                    }

                    // 将区域面板添加到边框中
                    sectionBorder.Child = sectionPanel;

                    // 将边框添加到主面板
                    panel.Children.Add(sectionBorder);

                    // 切换到下一个背景色
                    colorIndex++;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"从数据库加载按钮时出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 从Resources文件夹加载按钮（支持命令映射）
        /// </summary>
        /// <param Name="folderName">文件夹名称</param>
        /// <param Name="panel">目标面板</param>
        private void LoadButtonsFromResources(string folderName, WrapPanel panel)
        {
            try
            {
                // 显示调试信息
                string appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                LogManager.Instance.LogInfo($"应用程序路径: {appPath}");//显示调试信息没找到资源文件夹

                string resourcePath = System.IO.Path.Combine(appPath, "Resources", folderName);//返回本程序的资源文件夹路径；
                LogManager.Instance.LogInfo($"资源文件夹路径: {resourcePath}");//显示调试信息没找到资源文件夹
                LogManager.Instance.LogInfo($"资源文件夹是否存在: {System.IO.Directory.Exists(resourcePath)}");//显示调试信息没找到资源文件夹


                // 定义背景色列表，用于区分不同区域
                List<System.Windows.Media.Color> backgroundColors = new List<System.Windows.Media.Color>
                {
                    Colors.FloralWhite,
                    Colors.Azure,
                    Colors.FloralWhite,
                    Colors.Azure,
                    Colors.FloralWhite,
                    Colors.Azure,
                    Colors.FloralWhite,
                };

                int colorIndex = 0;

                // 检查一级文件夹是否存在
                if (System.IO.Directory.Exists(resourcePath))
                {
                    // 获取所有二级文件夹
                    string[] subDirectories = System.IO.Directory.GetDirectories(resourcePath);
                    LogManager.Instance.LogInfo($"找到 {subDirectories.Length} 个二级文件夹");

                    // 遍历所有二级文件夹
                    foreach (string subDir in subDirectories)
                    {
                        string subDirName = System.IO.Path.GetFileName(subDir);
                        LogManager.Instance.LogInfo($"处理二级文件夹: {subDirName}");

                        // 为每个二级文件夹创建一个带边框和背景色的区域
                        Border sectionBorder = new Border
                        {
                            BorderBrush = new SolidColorBrush(Colors.Gray),//边框颜色
                            BorderThickness = new Thickness(1),//边框宽度
                            CornerRadius = new CornerRadius(5),//圆角
                            Margin = new Thickness(0, 2, 0, 3),//间隔
                            Width = 282,
                            Background = new SolidColorBrush(backgroundColors[colorIndex % backgroundColors.Count]),//背景色
                            HorizontalAlignment = System.Windows.HorizontalAlignment.Left  // 左对齐
                        };

                        // 创建区域内容的StackPanel
                        StackPanel sectionPanel = new StackPanel
                        {
                            Margin = new Thickness(5)//间隔
                        };

                        // 添加区域标题
                        TextBlock sectionHeader = new TextBlock
                        {
                            Text = subDirName,
                            FontSize = 12,
                            FontWeight = FontWeights.Bold,
                            Margin = new Thickness(0, 5, 0, 5),
                            Foreground = new SolidColorBrush(Colors.DarkBlue)
                        };
                        sectionPanel.Children.Add(sectionHeader);//区域标题

                        // 获取该二级文件夹下的所有dwg文件
                        string[] files = System.IO.Directory.GetFiles(subDir, "*.dwg");
                        LogManager.Instance.LogInfo($"在 {subDirName} 中找到 {files.Length} 个dwg文件");//显示文件数量

                        if (files.Length > 0) //创建行面板
                        {
                            // 过滤并处理文件名
                            var buttonInfoList = new List<Tuple<string, string>>(); // (按钮名称, 完整文件路径)
                                                                                    // 遍历所有dwg文件
                            foreach (string file in files)
                            {
                                //调试文件
                                LogManager.Instance.LogInfo($"处理文件: {file}");
                                // 获取不带扩展名的文件名
                                string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(file);
                                // 去除_前的字符，获取按钮显示名称
                                string buttonName = fileNameWithoutExt;
                                if (fileNameWithoutExt.Contains("_"))
                                {
                                    // 去除_前的字符
                                    buttonName = fileNameWithoutExt.Substring(fileNameWithoutExt.IndexOf("_") + 1);
                                }
                                // 只保留中文字符，去除所有符号与英文字母
                                string chineseOnlyName = ExtractChineseCharacters(buttonName);
                                // 如果提取后没有中文字符，则使用原名称
                                if (string.IsNullOrEmpty(chineseOnlyName))
                                {
                                    chineseOnlyName = buttonName;
                                }
                                // 添加到列表 (按钮名称, 完整文件路径)
                                buttonInfoList.Add(new Tuple<string, string>(chineseOnlyName, file));
                            }

                            // 按钮名称排序
                            buttonInfoList.Sort((x, y) => x.Item1.CompareTo(y.Item1));

                            // 按3列分组处理
                            int columns = 3;
                            for (int i = 0; i < buttonInfoList.Count; i += columns)//3列
                            {
                                // 创建水平StackPanel用于放置一行按钮
                                StackPanel rowPanel = new StackPanel
                                {
                                    Orientation = System.Windows.Controls.Orientation.Horizontal,//水平
                                    Margin = new Thickness(0, 0, 0, 5) // 每行底部间隔5
                                };

                                // 添加该行的按钮（最多3个）
                                for (int j = 0; j < columns && (i + j) < buttonInfoList.Count; j++)
                                {
                                    var buttonInfo = buttonInfoList[i + j];//按钮信息
                                    string buttonName = buttonInfo.Item1;//按钮名称
                                    string fullPath = buttonInfo.Item2;//完整文件路径
                                    //string buttonName = graphic.DisplayName;
                                    if (!string.IsNullOrWhiteSpace(buttonName))
                                    {
                                        int lastUnderscore = buttonName.LastIndexOf('_');
                                        if (lastUnderscore >= 0 && lastUnderscore + 1 < buttonName.Length)
                                        {
                                            buttonName = buttonName.Substring(lastUnderscore + 1).Trim();
                                        }
                                        else
                                        {
                                            buttonName = buttonName.Trim();
                                        }
                                    }
                                    Button btn = new Button
                                    {
                                        Content = buttonName,//按钮内容
                                        Width = 88,//按钮宽度
                                        Height = 20,//按钮高度
                                        FontSize = 12,
                                        FontFamily = new System.Windows.Media.FontFamily("微软雅黑"),
                                        Margin = new Thickness(0, 0, 3, 0), // 按钮右侧间隔5
                                        HorizontalAlignment = System.Windows.HorizontalAlignment.Left,//水平居左
                                        VerticalAlignment = System.Windows.VerticalAlignment.Top,//垂直居上
                                        Tag = fullPath // 将完整路径存储在Tag属性中
                                    };
                                    btn.FontWeight = FontWeights.Normal;
                                    // 检查是否是预定义的按钮
                                    if (UnifiedCommandManager.IsPredefinedCommand(buttonName))
                                    {
                                        // 如果是预定义按钮
                                        btn.Tag = new ButtonTagCommandInfo
                                        {
                                            Type = "Predefined",
                                            ButtonName = buttonName,
                                            FilePath = fullPath
                                        };
                                        btn.Click += PredefinedButton_Click;
                                    }
                                    else
                                    {
                                        // 如果是普通图元按钮，存储文件路径
                                        btn.Tag = new ButtonTagCommandInfo
                                        {
                                            Type = "File",
                                            ButtonName = buttonName,
                                            FilePath = fullPath
                                        };
                                        //btn.Click += DynamicButton_Click;
                                    }
                                    // 新增：统一附加预检/拖拽事件（并避免重复绑定）
                                    AttachDynamicButtonHandlers(btn);

                                    // 添加按钮到行面板
                                    rowPanel.Children.Add(btn);
                                }
                                // 添加行面板到区域面板
                                sectionPanel.Children.Add(rowPanel);
                            }
                        }
                        else
                        {
                            // 如果该文件夹没有文件，显示提示信息
                            TextBlock noFilesText = new TextBlock
                            {
                                Text = "暂无文件",
                                FontSize = 12,
                                Margin = new Thickness(5, 0, 0, 5),
                                Foreground = new SolidColorBrush(Colors.Gray)
                            };
                            sectionPanel.Children.Add(noFilesText);
                        }

                        // 将区域面板添加到边框中
                        sectionBorder.Child = sectionPanel;

                        // 将边框添加到主面板
                        panel.Children.Add(sectionBorder);

                        // 切换到下一个背景色
                        colorIndex++;
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show($"找不到资源文件夹: {resourcePath}\n请检查Resources文件夹中的'{folderName}'文件夹是否存在");
                }
            }
            catch (Exception ex)
            {
                // 处理异常
                System.Windows.MessageBox.Show($"加载按钮时出错: {ex.Message}\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// 加载条件图元按钮
        /// </summary>
        private async void LoadConditionButtons()
        {
            try
            {
                LogManager.Instance.LogInfo("开始加载条件图元按钮...");

                // 清空现有按钮
                ClearConditionButtons();

                // 加载各专业条件按钮
                await LoadSpecializedConditionButtons("电气", 电气条件按钮面板);
                await LoadSpecializedConditionButtons("给排水", 给排水条件按钮面板);
                await LoadSpecializedConditionButtons("自控", 自控条件按钮面板);
                await LoadSpecializedConditionButtons("建筑", 结构条件按钮面板);
                await LoadSpecializedConditionButtons("结构", 结构条件按钮面板);
                await LoadSpecializedConditionButtons("暖通", 暖通条件按钮面板);

                LogManager.Instance.LogInfo("条件图元按钮加载完成");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载条件图元按钮时出错: {ex.Message}");
                System.Windows.MessageBox.Show($"加载条件图元按钮时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 加载指定专业的条件按钮
        /// </summary>
        private async Task LoadSpecializedConditionButtons(string 专业名称, WrapPanel targetPanel)
        {
            try
            {
                LogManager.Instance.LogInfo($"开始加载{专业名称}条件按钮...");

                if (targetPanel == null)
                {
                    LogManager.Instance.LogInfo($"目标面板 {专业名称} 为空");
                    return;
                }

                // 从数据库或资源文件夹中获取指定专业的条件文件
                var conditionFiles = await GetConditionFilesForSpecialty(专业名称);
                LogManager.Instance.LogInfo($"找到 {conditionFiles.Count} 个{专业名称}条件文件");

                if (conditionFiles.Count == 0)
                {
                    // 添加"暂无文件"提示
                    AddNoFilesLabel(targetPanel, $"暂无{专业名称}条件文件");
                    return;
                }

                // 按3列排列按钮
                int columns = 3;
                for (int i = 0; i < conditionFiles.Count; i += columns)
                {
                    StackPanel rowPanel = new StackPanel
                    {
                        Orientation = System.Windows.Controls.Orientation.Horizontal,
                        Margin = new Thickness(0, 0, 0, 5)
                    };

                    for (int j = 0; j < columns && (i + j) < conditionFiles.Count; j++)
                    {
                        var file = conditionFiles[i + j];
                        Button btn = CreateConditionButton(file);//创建条件按钮
                        rowPanel.Children.Add(btn);
                    }

                    targetPanel.Children.Add(rowPanel);
                }

                LogManager.Instance.LogInfo($"{专业名称}条件按钮加载完成");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载{专业名称}条件按钮时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取指定专业的条件文件
        /// </summary>
        private async Task<List<ConditionFileInfo>> GetConditionFilesForSpecialty(string specialtyName)
        {
            var conditionFiles = new List<ConditionFileInfo>();

            try
            {
                // 这里根据您的实际数据源来实现
                // 可以是从数据库、资源文件夹或其他地方获取

                // 示例实现（您需要根据实际情况修改）：
                if (_databaseManager != null)
                {
                    // 从数据库获取条件文件
                    // conditionFiles = await _databaseManager.GetConditionFilesBySpecialtyAsync(specialtyName);
                }
                else
                {
                    // 从资源文件夹获取条件文件
                    conditionFiles = GetConditionFilesFromResources(specialtyName);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"获取{specialtyName}条件文件时出错: {ex.Message}");
            }

            return conditionFiles;
        }

        /// <summary>
        /// 从资源文件夹获取条件文件
        /// </summary>
        private List<ConditionFileInfo> GetConditionFilesFromResources(string specialtyName)
        {
            var conditionFiles = new List<ConditionFileInfo>();

            try
            {
                // 根据专业名称确定资源路径
                string resourcePath = $"pack://application:,,,/Resources/Conditions/{specialtyName}/";

                // 这里需要根据您的实际资源结构来实现
                // 示例数据：
                switch (specialtyName)
                {
                    case "电气":
                        conditionFiles.AddRange(new[]
                        {
                    new ConditionFileInfo { Name = "电气条件1", DisplayName = "电气条件1", FilePath = $"{resourcePath}电气条件1.dwg" },
                    new ConditionFileInfo { Name = "电气条件2", DisplayName = "电气条件2", FilePath = $"{resourcePath}电气条件2.dwg" },
                    new ConditionFileInfo { Name = "电气条件3", DisplayName = "电气条件3", FilePath = $"{resourcePath}电气条件3.dwg" }
                });
                        break;

                    case "自控":
                        conditionFiles.AddRange(new[]
                        {
                    new ConditionFileInfo { Name = "自控条件1", DisplayName = "自控条件1", FilePath = $"{resourcePath}自控条件1.dwg" },
                    new ConditionFileInfo { Name = "自控条件2", DisplayName = "自控条件2", FilePath = $"{resourcePath}自控条件2.dwg" }
                });
                        break;

                    case "给排水":
                        conditionFiles.AddRange(new[]
                        {
                    new ConditionFileInfo { Name = "给排水条件1", DisplayName = "给排水条件1", FilePath = $"{resourcePath}给排水条件1.dwg" },
                    new ConditionFileInfo { Name = "给排水条件2", DisplayName = "给排水条件2", FilePath = $"{resourcePath}给排水条件2.dwg" },
                    new ConditionFileInfo { Name = "给排水条件3", DisplayName = "给排水条件3", FilePath = $"{resourcePath}给排水条件3.dwg" }
                });
                        break;

                    case "暖通":
                        conditionFiles.AddRange(new[]
                        {
                    new ConditionFileInfo { Name = "暖通条件1", DisplayName = "暖通条件1", FilePath = $"{resourcePath}暖通条件1.dwg" }
                });
                        break;

                    case "结构":
                        conditionFiles.AddRange(new[]
                        {
                    new ConditionFileInfo { Name = "结构条件1", DisplayName = "结构条件1", FilePath = $"{resourcePath}结构条件1.dwg" },
                    new ConditionFileInfo { Name = "结构条件2", DisplayName = "结构条件2", FilePath = $"{resourcePath}结构条件2.dwg" }
                });
                        break;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"从资源获取{specialtyName}条件文件时出错: {ex.Message}");
            }

            return conditionFiles;
        }

        /// <summary>
        /// 创建条件按钮
        /// </summary>
        private Button CreateConditionButton(ConditionFileInfo fileInfo)
        {
            // 仅显示最后一个下划线后的名称，例如 "DQTJ_EQUIP_潮湿插座" -> "潮湿插座"
            string buttonText = fileInfo?.DisplayName ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(buttonText))
            {
                int lastUnderscore = buttonText.LastIndexOf('_');
                if (lastUnderscore >= 0 && lastUnderscore + 1 < buttonText.Length)
                {
                    buttonText = buttonText.Substring(lastUnderscore + 1).Trim();
                }
                else
                {
                    buttonText = buttonText.Trim();
                }
            }
            Button btn = new Button
            {
                Content = buttonText,
                Width = 85,
                Height = 20,
                Margin = new Thickness(5, 1, 1, 1),
                Tag = fileInfo, // 存储文件信息
                FontFamily = new FontFamily("Microsoft YaHei UI"),
                FontWeight = FontWeights.Normal
            };

            // 应用统一的按钮样式
            btn.Style = (Style)FindResource("ButtonStyle"); // 如果您有自定义按钮样式

            // 添加点击事件
            btn.Click += ConditionButton_Click;

            return btn;
        }

        /// <summary>
        /// 条件按钮点击事件
        /// </summary>
        private void ConditionButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button btn && btn.Tag is ConditionFileInfo fileInfo)
                {
                    LogManager.Instance.LogInfo($"点击条件按钮: {fileInfo.DisplayName}");

                    // 执行条件插入操作
                    ExecuteConditionInsert(fileInfo);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"执行条件插入时出错: {ex.Message}");
                MessageBox.Show($"执行条件插入时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 执行条件插入操作
        /// </summary>
        private void ExecuteConditionInsert(ConditionFileInfo fileInfo)
        {
            try
            {
                // 设置全局变量
                VariableDictionary.btnFileName = fileInfo.Name;
                VariableDictionary.btnBlockLayer = "TJ(条件图元)";
                VariableDictionary.layerColorIndex = 7; // 默认颜色

                // 执行插入命令
                Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);

                LogManager.Instance.LogInfo($"成功插入条件: {fileInfo.DisplayName}");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"插入条件失败: {ex.Message}");
                MessageBox.Show($"插入条件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 清空条件按钮
        /// </summary>
        private void ClearConditionButtons()
        {
            if (电气条件按钮面板 != null) 电气条件按钮面板.Children.Clear();
            if (自控条件按钮面板 != null) 自控条件按钮面板.Children.Clear();
            if (给排水条件按钮面板 != null) 给排水条件按钮面板.Children.Clear();
            if (暖通条件按钮面板 != null) 暖通条件按钮面板.Children.Clear();
            if (结构条件按钮面板 != null) 结构条件按钮面板.Children.Clear();
        }

        /// <summary>
        /// 添加"暂无文件"提示
        /// </summary>
        private void AddNoFilesLabel(WrapPanel panel, string message)
        {
            TextBlock noFilesText = new TextBlock
            {
                Text = message,
                FontStyle = FontStyles.Italic,
                Foreground = new SolidColorBrush(Colors.Gray),
                Margin = new Thickness(10, 5, 0, 5)
            };
            panel.Children.Add(noFilesText);
        }

        /// <summary>
        /// 条件文件信息类
        /// </summary>
        public class ConditionFileInfo
        {
            public string? Name { get; set; }
            public string? DisplayName { get; set; }
            public string? FilePath { get; set; }
            public string? Specialty { get; set; } // 专业类别
            public DateTime CreatedTime { get; set; }
        }

        /// <summary>
        /// 提取字符串中的中文字符，去除所有符号与英文字母
        /// </summary>
        /// <param Name="input">输入字符串</param>
        /// <returns>只包含中文字符的字符串</returns>
        private string ExtractChineseCharacters(string input)
        {
            //使用正则表达式 提取中文字符
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            //匹配中文字符 正则表达式匹配中文字符范围
            var matches = Regex.Matches(input!, @"[\u4e00-\u9fff]+");
            //获取匹配的字符串如果没有匹配到中文字符，返回空字符串
            if (matches.Count == 0) return string.Empty;
            //获取匹配的字符串并拼接为一个字符串连接所有匹配到的中文字符
            return string.Concat(matches.Cast<Match>().Select(m => m.Value)).Trim();
        }

        /// <summary>
        /// 在DataGrid中显示文件属性（用于CAD图元界面）
        /// </summary>
        private async Task DisplayFilePropertiesInDataGridAsync(FileStorage fileStorage)
        {
            try
            {
                LogManager.Instance.LogInfo($"在PropertiesDataGrid中显示文件 {fileStorage.DisplayName} 的属性");

                if (PropertiesDataGrid == null)
                {
                    LogManager.Instance.LogWarning("PropertiesDataGrid控件为空");
                    return;
                }

                if (_databaseManager == null)
                {
                    LogManager.Instance.LogWarning("数据库管理器为空");
                    PropertiesDataGrid.ItemsSource = null;
                    return;
                }

                // 获取文件属性
                var fileAttribute = await _databaseManager.GetFileAttributeByGraphicIdAsync(fileStorage.Id);

                // 准备显示数据
                //var displayData = PrepareFileDisplayDataForDataGrid(fileStorage, fileAttribute);
                var displayData = PrepareFileDisplayData(fileStorage, fileAttribute);
                PropertiesDataGrid.ItemsSource = displayData;
                // 记录当前显示的属性快照，便于后续判断哪些字段被修改（用于插入后应用属性）
                CapturePropertiesSnapshot(displayData);
                LogManager.Instance.LogInfo("文件属性在PropertiesDataGrid中显示完成");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"在PropertiesDataGrid中显示文件属性时出错: {ex.Message}");
                if (PropertiesDataGrid != null)
                {
                    PropertiesDataGrid.ItemsSource = null;
                }
                MessageBox.Show($"显示文件属性时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 显示文件预览图片
        /// </summary>
        private async void ShowFilePreview(FileStorage fileStorage)
        {
            try
            {
                if (预览 == null)
                {
                    LogManager.Instance.LogInfo("预览图片控件为空");
                    return;
                }

                // 清空现有预览
                预览.Source = null;

                if (fileStorage == null)
                {
                    LogManager.Instance.LogInfo("文件存储对象为空");
                    return;
                }

                LogManager.Instance.LogInfo($"显示文件预览: {fileStorage.DisplayName}");

                // 获取预览图片
                var previewImage = await GetPreviewImageAsync(fileStorage);

                if (previewImage != null)
                {
                    预览.Source = previewImage;
                    LogManager.Instance.LogInfo("预览图片显示成功");
                }
                else
                {
                    LogManager.Instance.LogWarning("无法加载预览图片");
                    // 显示默认图片或提示
                    预览.Source = GetDefaultPreviewImage();
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"显示文件预览时出错: {ex.Message}");
                // 显示错误图片
                预览.Source = GetDefaultPreviewImage();
            }
        }

        /// <summary>
        /// 显示文件属性
        /// </summary>
        private async Task DisplayFilePropertiesAsync(FileStorage fileStorage)
        {
            try
            {
                LogManager.Instance.LogInfo($"显示文件属性: {fileStorage.DisplayName}");

                if (_databaseManager == null)
                {
                    LogManager.Instance.LogWarning("数据库管理器为空");
                    return;
                }

                // 获取文件属性
                var fileAttribute = await _databaseManager.GetFileAttributeByGraphicIdAsync(fileStorage.Id);

                // 准备显示数据
                var displayData = PrepareFileDisplayData(fileStorage, fileAttribute);

                // 更新显示（CAD图元界面）
                if (PropertiesDataGrid != null)
                {
                    PropertiesDataGrid.ItemsSource = displayData;
                }

                LogManager.Instance.LogInfo("文件属性显示完成");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"显示文件属性时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行Resources图元按钮点击后的操作
        /// </summary>
        /// <param Name="buttonName">按钮名称</param>
        /// <param Name="filePath">文件路径</param>
        private void ExecuteDynamicButtonActionFromResources(string buttonName, string filePath)
        {
            try
            {
                // 1. 显示预览图
                ShowPreviewImage(filePath, buttonName);

                // 2. 调用AutoCAD命令
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.SendStringToExecute($"DBTextLabel\n", true, false, false);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"执行Resources按钮操作时出错: {ex.Message}");
                System.Windows.MessageBox.Show($"执行Resources按钮操作时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 显示预览图片
        /// </summary>
        /// <param Name="dwgFilePath">dwg文件路径</param>
        /// <param Name="buttonName">按钮名称</param>
        private void ShowPreviewImage(string dwgFilePath, string buttonName)
        {
            try
            {
                // 如果没有预览Viewbox，直接返回
                if (previewViewbox == null) return;

                // 清空现有的预览内容
                previewViewbox.Child = null;

                // 获取文件所在的文件夹路径
                string folderPath = System.IO.Path.GetDirectoryName(dwgFilePath);

                // 构造png文件路径 (与dwg文件同名)
                string pngFilePath = System.IO.Path.Combine(folderPath,
                    System.IO.Path.GetFileNameWithoutExtension(dwgFilePath) + ".png");

                // 检查png文件是否存在
                if (System.IO.File.Exists(pngFilePath))
                {
                    // 创建Image控件显示预览图
                    System.Windows.Controls.Image previewImage = new System.Windows.Controls.Image
                    {
                        Stretch = Stretch.Uniform,
                        Margin = new Thickness(5)
                    };

                    // 创建BitmapImage并加载png文件
                    BitmapImage bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(pngFilePath);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();

                    // 设置图片源
                    previewImage.Source = bitmap;

                    // 将图片添加到Viewbox中
                    previewViewbox.Child = previewImage;
                }
                else
                {
                    // 如果没有找到png文件，显示提示文字
                    TextBlock noPreviewText = new TextBlock
                    {
                        Text = "无预览图",
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                        VerticalAlignment = System.Windows.VerticalAlignment.Center,
                        Foreground = new SolidColorBrush(Colors.Gray)
                    };

                    previewViewbox.Child = noPreviewText;
                }
            }
            catch (Exception ex)
            {
                // 处理预览图加载异常
                System.Windows.MessageBox.Show($"加载预览图时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 查找可视化树中的子元素
        /// </summary>
        /// <typeparam Name="T">要查找的元素类型</typeparam>
        /// <param Name="parent">父元素</param>
        /// <param Name="childName">子元素名称</param>
        /// <returns>找到的子元素或null</returns>
        private T FindVisualChild<T>(DependencyObject parent, string childName) where T : DependencyObject
        {
            // 遍历所有子元素
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                // 检查是否是目标类型
                if (child != null && child is T typedChild)
                {
                    // 检查名称是否匹配
                    if (child is FrameworkElement frameworkElement && frameworkElement.Name == childName)
                    {
                        return typedChild;
                    }
                }

                // 递归查找子元素
                var childOfChild = FindVisualChild<T>(child, childName);
                if (childOfChild != null)
                {
                    return childOfChild;
                }
            }

            return null;
        }

        /// <summary>
        /// 初始化条件图图层
        /// </summary>
        public static void NewTjLayer()
        {
            while (true)
            {
                foreach (var item in VariableDictionary.GGtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.GYtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.AtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.StjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.PtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.NtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.EtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.ZKtjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                foreach (var item in VariableDictionary.tjtBtn)
                {
                    if (!VariableDictionary.allTjtLayer.Contains(item))
                        VariableDictionary.allTjtLayer.Add(item);
                }
                break;
            }
        }

        /// <summary>
        /// 读取本地设置路径下的配置文件
        /// </summary>
        private void Load()
        {
            string[]? lines = null;
            try
            {
                lines = System.IO.File.ReadAllLines(GetPath.referenceFile);//按每一行为一个DWG文件读进来； 
                GetPath.ListDwgFile.AddRange(lines);//把本程序下添加的文件都显示在列表里；
            }
            catch
            {
            }
        }

        /// <summary>
        /// 保存添加的图库文件与写入配置文件中
        /// </summary>
        public void Save()
        {
            try
            {
                using (var sr = new StreamWriter(GetPath.referenceFile)) //useing调用后主动释放文件
                {
                    foreach (var item in GetPath.ListDwgFile)
                    {
                        sr.WriteLine(item);
                    }
                }
            }
            catch (System.Exception)
            {
            }
        }

        #endregion

        #region 方向按钮事件处理方法...

        private void 上_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("上");
            command?.Invoke();
        }

        private void 右上_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("右上");
            command?.Invoke();
        }

        private void 右_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("右");
            command?.Invoke();
        }

        private void 右下_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("右下");
            command?.Invoke();

        }

        private void 下_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("下");
            command?.Invoke();
        }

        private void 左下_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("左下");
            command?.Invoke();
        }

        private void 左_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("左");
            command?.Invoke();
        }

        private void 左上_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("左上");
            command?.Invoke();
        }
        #endregion

        #region 功能区按键处理方法...

        private void 查找_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 功能1_Btn_Click(object sender, RoutedEventArgs e)
        {
            Env.Document.SendStringToExecute("pu ", false, false, false);
        }

        private void 功能2_Btn_Click(object sender, RoutedEventArgs e)
        {
            Env.Document.SendStringToExecute("audit\n y ", false, false, false);
        }

        private void 功能3_Btn_Click(object sender, RoutedEventArgs e)
        {
            Env.Document.SendStringToExecute("DRAWINGRECOVERY ", false, false, false);
        }

        /// <summary>
        /// 当绘图比例输入框内容改变时保存配置
        /// </summary>
        private void TextBox_绘图比例_TextChanged(object sender, TextChangedEventArgs e)
        {
            SaveDrawingConfig();
        }
        #endregion

        #region CAD\SW 管理员数据库操作

        #region CAD\SW 分类树

        #region 架构树新方法

        /// <summary>
        /// 初始化架构树
        /// </summary>
        /// <returns></returns>
        private async Task InitializeCategoryTreeAsync()
        {
            try
            {
                await _categoryManager.LoadCategoryTreeAsync(_categoryTreeNodes, _databaseManager);
                _categoryTreeView = CategoryTreeView;//赋值给全局变量
                _categoryManager.DisplayCategoryTree(_categoryTreeView, _categoryTreeNodes);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"初始化架构树失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取TreeViewItem的辅助方法（增强版）
        /// </summary>
        /// <param name="container"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        private TreeViewItem GetTreeViewItem(ItemsControl container, object item)
        {
            if (container == null) return null;

            // 首先尝试直接获取
            var directlyFound = container.ItemContainerGenerator.ContainerFromItem(item) as TreeViewItem;
            if (directlyFound != null)
                return directlyFound;

            // 如果直接获取失败，遍历所有子项
            if (container.Items != null)
            {
                foreach (var containerItem in container.Items)
                {
                    var treeViewItem = container.ItemContainerGenerator.ContainerFromItem(containerItem) as TreeViewItem;
                    if (treeViewItem != null)
                    {
                        if (treeViewItem.DataContext == item)
                        {
                            return treeViewItem;
                        }

                        // 递归查找子项
                        var child = GetTreeViewItem(treeViewItem, item);
                        if (child != null)
                        {
                            return child;
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 调整DataGrid行高以适应换行文本
        /// </summary>
        private void AdjustDataGridRowHeight()
        {
            try
            {
                if (StroageFileDataGrid != null)
                {
                    // 设置行高为自动调整
                    StroageFileDataGrid.RowHeight = Double.NaN; // 自动行高

                    // 或者设置一个最小行高
                    // StroageFileDataGrid.MinRowHeight = 60;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"调整DataGrid行高时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在DataGrid数据源改变时调整行高
        /// </summary>
        private void StroageFileDataGrid_TargetUpdated(object sender, System.Windows.Data.DataTransferEventArgs e)
        {
            try
            {
                // 延迟调整行高，确保数据已加载
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    AdjustDataGridRowHeight();
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"DataGrid数据更新时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 在DataGrid加载完成后调整行高
        /// </summary>
        //private void StroageFileDataGrid_Loaded(object sender, RoutedEventArgs e)
        //{
        //    try
        //    {
        //        AdjustDataGridRowHeight();
        //    }
        //    catch (Exception ex)
        //    {
        //        LogManager.Instance.LogInfo($"DataGrid加载时出错: {ex.Message}");
        //    }
        //}
        private void StroageFileDataGrid_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                var dg = sender as System.Windows.Controls.DataGrid;
                if (dg == null) return;

                // 防止重复订阅
                dg.LoadingRow -= DataGrid_LoadingRow;
                dg.LoadingRow += DataGrid_LoadingRow;

                // 对已存在行也附加一次（designer/热重载场景）
                foreach (var item in dg.Items)
                {
                    var row = dg.ItemContainerGenerator.ContainerFromItem(item) as System.Windows.Controls.DataGridRow;
                    if (row != null)
                    {
                        EnsureRowContextMenu(row);
                    }
                }
            }
            catch
            {
                // 忽略 UI 初始化中的异常，避免二次失败
            }
        }

        // 每行加载时调用
        private void DataGrid_LoadingRow(object? sender, System.Windows.Controls.DataGridRowEventArgs e)
        {
            try
            {
                EnsureRowContextMenu(e.Row);
            }
            catch { }
        }

        // 为单个 DataGridRow 动态创建 ContextMenu（只在管理员模块可见由 Replace 事件处理里做权限校验）
        private void EnsureRowContextMenu(System.Windows.Controls.DataGridRow row)
        {
            if (row == null) return;

            // 如果已存在并且是我们创建的则跳过
            if (row.ContextMenu != null) return;

            var cm = new System.Windows.Controls.ContextMenu();

            var miReplace = new System.Windows.Controls.MenuItem
            {
                Header = "替换图元"
            };

            // 把当前行的数据对象直接设置为 CommandParameter，事件中优先使用
            miReplace.CommandParameter = row.Item;
            miReplace.Click += ReplaceFileMenuItem_Click;

            cm.Items.Add(miReplace);
            row.ContextMenu = cm;
        }

        // 替换菜单点击：优先用 CommandParameter 获取行对象，回退使用 DataContext
        private async void ReplaceFileMenuItem_Click(object? sender, RoutedEventArgs e)
        {
            try
            {
                var mi = sender as System.Windows.Controls.MenuItem;
                if (mi == null) return;

                // 优先使用 CommandParameter（在 EnsureRowContextMenu 中设定）
                object? storageObj = mi.CommandParameter;

                // 回退：使用 MenuItem.DataContext 或 ContextMenu.PlacementTarget 的 DataContext
                if (storageObj == null)
                {
                    storageObj = mi.DataContext;
                    if (storageObj == null)
                    {
                        var cm = mi.Parent as System.Windows.Controls.ContextMenu;
                        if (cm == null)
                        {
                            // 更稳健的向上查找 ContextMenu
                            var parent = System.Windows.Media.VisualTreeHelper.GetParent(mi as System.Windows.DependencyObject);
                            while (parent != null && !(parent is System.Windows.Controls.ContextMenu))
                                parent = System.Windows.Media.VisualTreeHelper.GetParent(parent);
                            cm = parent as System.Windows.Controls.ContextMenu;
                        }
                        var row = cm?.PlacementTarget as System.Windows.Controls.DataGridRow;
                        storageObj = row?.DataContext;
                    }
                }

                if (storageObj == null)
                {
                    System.Windows.MessageBox.Show("未能识别要替换的文件记录。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 权限校验（示例：只允许管理员）
                var userName = (VariableDictionary._userName ?? TextBox_Set_Username.Text ?? string.Empty).Trim();
                if (!IsAdminUser(userName))
                {
                    System.Windows.MessageBox.Show("仅管理员用户可以执行替换操作。", "权限不足", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 选择本地文件（WinForms OpenFileDialog）
                using (var ofd = new System.Windows.Forms.OpenFileDialog())
                {
                    ofd.Filter = "DWG 文件 (*.dwg)|*.dwg|所有文件 (*.*)|*.*";
                    ofd.Title = "选择要上传并替换的文件";
                    if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
                    var localPath = ofd.FileName;
                    if (!System.IO.File.Exists(localPath))
                    {
                        System.Windows.MessageBox.Show("所选文件不存在。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    var confirm = System.Windows.MessageBox.Show($"确认将本地文件\n{System.IO.Path.GetFileName(localPath)}\n覆盖服务器上此条记录对应的文件（保留原始文件名/位置）？",
                                                               "确认替换", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                    if (confirm != MessageBoxResult.OK) return;

                    // 调用你现有的替换逻辑（TryInvokeReplaceApisAsync）完成上传/覆盖
                    var (success, error) = await TryInvokeReplaceApisAsync(storageObj, localPath);
                    if (!success)
                    {
                        System.Windows.MessageBox.Show($"替换失败: {error}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // 清理缓存与刷新
                    try
                    {
                        var previewKey = GetStoragePreviewKey(storageObj);
                        if (!string.IsNullOrWhiteSpace(previewKey))
                        {
                            var cached = System.IO.Path.Combine(_previewCachePath ?? string.Empty, previewKey + ".png");
                            if (System.IO.File.Exists(cached)) System.IO.File.Delete(cached);
                        }
                    }
                    catch { }

                    try { await RefreshCurrentCategoryDisplayAsync(_selectedCategoryNode); } catch { }

                    System.Windows.MessageBox.Show("替换成功。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"替换过程中发生异常: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>
        /// 添加文件名处理方法
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="maxLength"></param>
        /// <returns></returns>
        private string FormatFileNameForDisplay(string fileName, int maxLength = 50)
        {
            if (string.IsNullOrEmpty(fileName))
                return string.Empty;

            if (fileName.Length <= maxLength)
                return fileName;

            // 截断过长的文件名并添加省略号
            return fileName.Substring(0, maxLength - 3) + "...";
        }

        /// <summary>
        /// 递归展开所有子节点
        /// </summary>
        /// <param name="item"></param>
        private void ExpandAllChildren(TreeViewItem item)
        {
            if (item == null) return;

            item.IsExpanded = true;
            foreach (var child in item.Items)
            {
                var childItem = GetTreeViewItem(item, child);
                if (childItem != null)
                {
                    ExpandAllChildren(childItem);
                }
            }
        }

        /// <summary>
        /// 架构树选中项改变事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void CategoryTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            try
            {
                if (e.NewValue is CategoryTreeNode selectedNode)
                {
                    _selectedCategoryNode = selectedNode;
                    //LogManager.Instance.LogInfo($"选中分类节点: {selectedNode.DisplayText} (ID: {selectedNode.Id}, Level: {selectedNode.Level})");
                    LogManager.Instance.LogInfo($"选中分类节点: {selectedNode.DisplayText} (ID: {selectedNode.Id}, Level: {selectedNode.Level})");
                    // 根据选中的节点类型显示相应的属性编辑界面
                    DisplayNodePropertiesForEditing(selectedNode);

                    // 加载该分类下的文件
                    await LoadFilesForCategoryAsync(selectedNode);
                }
                else
                {
                    LogManager.Instance.LogInfo("选中的节点为空或类型不正确");
                    StroageFileDataGrid.ItemsSource = null;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"处理架构树选中项改变失败: {ex.Message}");
                MessageBox.Show($"处理分类选择失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 显示节点属性用于编辑
        /// </summary>
        /// <param name="node"></param>
        private void DisplayNodePropertiesForEditing(CategoryTreeNode node)
        {
            try
            {
                var propertyRows = new List<CategoryPropertyEditModel>();

                if (node.Level == 0 && node.Data is CadCategory category)
                {
                    // 主分类
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "ID",
                        PropertyValue1 = category.Id.ToString(),
                        PropertyName2 = "名称",
                        PropertyValue2 = category.Name
                    });
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "显示名称",
                        PropertyValue1 = category.DisplayName,
                        PropertyName2 = "排序序号",
                        PropertyValue2 = category.SortOrder.ToString()
                    });
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "子分类数",
                        PropertyValue1 = GetSubcategoryCount(category).ToString(),
                        PropertyName2 = "",
                        PropertyValue2 = ""
                    });
                }
                else if (node.Data is CadSubcategory subcategory)
                {
                    // 子分类
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "ID",
                        PropertyValue1 = subcategory.Id.ToString(),
                        PropertyName2 = "父ID",
                        PropertyValue2 = subcategory.ParentId.ToString()
                    });
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "名称",
                        PropertyValue1 = subcategory.Name,
                        PropertyName2 = "显示名称",
                        PropertyValue2 = subcategory.DisplayName
                    });
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "排序序号",
                        PropertyValue1 = subcategory.SortOrder.ToString(),
                        PropertyName2 = "层级",
                        PropertyValue2 = subcategory.Level.ToString()
                    });
                    propertyRows.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = "子分类数",
                        PropertyValue1 = GetSubcategoryCount(subcategory).ToString(),
                        PropertyName2 = "",
                        PropertyValue2 = ""
                    });
                }

                // 添加空行用于编辑
                propertyRows.Add(new CategoryPropertyEditModel());
                propertyRows.Add(new CategoryPropertyEditModel());

                CategoryPropertiesDataGrid.ItemsSource = propertyRows;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"显示节点属性失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取分类数量
        /// </summary>
        /// <param name="category"></param>
        /// <returns></returns>
        private int GetSubcategoryCount(CadCategory category)
        {
            if (string.IsNullOrEmpty(category.SubcategoryIds))
                return 0;

            return category.SubcategoryIds.Split(',').Length;
        }

        /// <summary>
        /// 获取子分类数量
        /// </summary>
        /// <param name="subcategory"></param>
        /// <returns></returns>
        private int GetSubcategoryCount(CadSubcategory subcategory)
        {
            if (string.IsNullOrEmpty(subcategory.SubcategoryIds))
                return 0;

            return subcategory.SubcategoryIds.Split(',').Length;
        }

        /// <summary>
        /// 初始化子分类属性编辑界面
        /// </summary>
        /// <param name="parentNode"></param>
        private void InitializeSubcategoryPropertiesForEditing(CategoryTreeNode parentNode)
        {
            var subcategoryProperties = new List<CategoryPropertyEditModel>
            {
                new CategoryPropertyEditModel { PropertyName1 = "父分类ID", PropertyValue1 = parentNode.Id.ToString(), PropertyName2 = "名称", PropertyValue2 = "" },
                new CategoryPropertyEditModel { PropertyName1 = "显示名称", PropertyValue1 = "", PropertyName2 = "排序序号", PropertyValue2 = "自动生成" } // 留空，表示自动生成
            };

            // 添加参考信息
            subcategoryProperties.Add(new CategoryPropertyEditModel
            {
                PropertyName1 = "父级名称",
                PropertyValue1 = parentNode.DisplayText,
                PropertyName2 = "",
                PropertyValue2 = ""
            });

            // 添加空行用于用户输入
            subcategoryProperties.Add(new CategoryPropertyEditModel());
            subcategoryProperties.Add(new CategoryPropertyEditModel());

            CategoryPropertiesDataGrid.ItemsSource = subcategoryProperties;
        }

        /// <summary>
        /// 初始化主分类属性编辑界面
        /// </summary>
        private void InitializeCategoryPropertiesForCategory()
        {
            var categoryProperties = new List<CategoryPropertyEditModel>
            {
                new CategoryPropertyEditModel { PropertyName1 = "名称", PropertyValue1 = "", PropertyName2 = "显示名称", PropertyValue2 = "" },
                new CategoryPropertyEditModel { PropertyName1 = "排序序号", PropertyValue1 = "自动生成", PropertyName2 = "", PropertyValue2 = "" } // 留空，表示自动生成
            };

            // 添加空行用于用户输入
            categoryProperties.Add(new CategoryPropertyEditModel());
            categoryProperties.Add(new CategoryPropertyEditModel());

            CategoryPropertiesDataGrid.ItemsSource = categoryProperties;
        }

        /// <summary>
        /// 加载分类下的文件
        /// </summary>
        /// <param name="categoryNode"></param>
        /// <returns></returns>
        private async Task LoadFilesForCategoryAsync(CategoryTreeNode categoryNode)
        {
            try
            {
                if (_databaseManager == null)
                {
                    LogManager.Instance.LogInfo("数据库管理器未初始化");
                    return;
                }

                List<FileStorage> files = new List<FileStorage>();

                //LogManager.Instance.LogInfo($"开始加载分类 {categoryNode.Id} ({categoryNode.DisplayText}) 的文件");
                LogManager.Instance.LogInfo($"开始加载分类 {categoryNode.Id} ({categoryNode.DisplayText}) 的文件");

                if (categoryNode.Level == 0 && categoryNode.Data is CadCategory category)
                {
                    // 主分类
                    LogManager.Instance.LogInfo($"加载主分类 {category.Name} (ID: {category.Id}) 的文件");
                    files = await _databaseManager.GetFilesByCategoryIdAsync(category.Id, "main");
                }
                else if (categoryNode.Data is CadSubcategory subcategory)
                {
                    // 子分类

                    LogManager.Instance.LogInfo($"加载子分类 {subcategory.Name} (ID: {subcategory.Id}) 的文件");
                    files = await _databaseManager.GetFilesByCategoryIdAsync(subcategory.Id, "sub");
                }
                else
                {
                    LogManager.Instance.LogInfo("未知的节点类型");
                    return;
                }

                LogManager.Instance.LogInfo($"从数据库查询到 {files.Count} 个文件");

                // 调试输出文件信息
                DebugFileData(files);

                // 确保在UI线程更新DataGrid
                Dispatcher.Invoke(() =>
                {
                    StroageFileDataGrid.ItemsSource = files;
                    LogManager.Instance.LogInfo($"DataGrid已更新，显示 {files.Count} 个文件");
                });

                // 如果没有文件，显示提示
                if (files.Count == 0)
                {
                    LogManager.Instance.LogInfo($"分类 '{categoryNode.DisplayText}' 下没有文件");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载文件列表失败: {ex.Message}");
                LogManager.Instance.LogInfo($"堆栈跟踪: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 添加调试方法来检查加载的文件数据
        /// </summary>
        /// <param name="files"></param>
        private void DebugFileData(List<FileStorage> files)
        {
            LogManager.Instance.LogInfo($"=== 文件数据调试信息 ===");
            LogManager.Instance.LogInfo($"文件总数: {files.Count}");

            foreach (var file in files)
            {
                LogManager.Instance.LogInfo($"文件ID: {file.Id}");
                LogManager.Instance.LogInfo($"  名称: {file.DisplayName ?? file.FileName}");
                LogManager.Instance.LogInfo($"  路径: {file.FilePath}");
                LogManager.Instance.LogInfo($"  预览图: {file.PreviewImagePath}");
                LogManager.Instance.LogInfo($"  分类ID: {file.CategoryId}");
                LogManager.Instance.LogInfo($"  分类类型: {file.CategoryType}");
                LogManager.Instance.LogInfo("---");
            }
        }

        /// <summary>
        /// 文件列表双击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void StroageFileDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (StroageFileDataGrid.SelectedItem is FileStorage selectedFile)
                {
                    // 显示选中文件的详细信息
                    DisplayFileStorageInfo(selectedFile);

                    // 显示预览图片
                    var previewBitmap = await GetPreviewImageAsync(selectedFile);

                    System.Diagnostics.Debug.WriteLine($"选中文件: {selectedFile.DisplayName}\n文件ID: {selectedFile.Id}",
                        "文件信息", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"处理文件选择失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 添加预览图片加载事件处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void PreviewImage_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                var image = sender as Image;
                if (image?.Tag is FileStorage fileStorage)
                {
                    try
                    {
                        // 异步加载图片
                        var bitmap = await GetPreviewImageAsync(fileStorage);

                        // 在 PreviewImage_Loaded 的异步 UI 更新段内，将错误的赋值改为具体的 Visibility 枚举值
                        await Dispatcher.InvokeAsync(() =>
                        {
                            if (image != null)
                            {
                                image.Source = bitmap;

                                // 隐藏加载文本（使用合适的枚举值）
                                var parentGrid = image.Parent as Grid;
                                if (parentGrid != null)
                                {
                                    var loadingText = parentGrid.Children.OfType<TextBlock>().FirstOrDefault();
                                    if (loadingText != null)
                                    {
                                        loadingText.Visibility = System.Windows.Visibility.Collapsed;
                                    }
                                }
                            }
                        });

                        //// 在UI线程更新图片
                        //await Dispatcher.InvokeAsync(() =>
                        //{
                        //    if (image != null)
                        //    {
                        //        image.Source = bitmap;

                        //        // 隐藏加载文本
                        //        var parentGrid = image.Parent as Grid;
                        //        if (parentGrid != null)
                        //        {
                        //            var loadingText = parentGrid.Children.OfType<TextBlock>().FirstOrDefault();
                        //            if (loadingText != null)
                        //            {
                        //                loadingText.Visibility = Visibility;
                        //            }
                        //        }
                        //    }
                        //});
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogInfo($"设置图片源时出错: {ex.Message}");

                        // 显示错误信息
                        await Dispatcher.InvokeAsync(() =>
                        {
                            if (image != null)
                            {
                                image.Source = GetDefaultPreviewImage();

                                // 显示错误文本
                                var parentGrid = image.Parent as Grid;
                                if (parentGrid != null)
                                {
                                    var loadingText = parentGrid.Children.OfType<TextBlock>().FirstOrDefault();
                                    if (loadingText != null)
                                    {
                                        loadingText.Text = "加载失败";
                                        loadingText.Foreground = new SolidColorBrush(Colors.Red);
                                    }
                                }
                            }
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"预览图片加载事件处理失败: {ex.Message}");
            }
        }

        private string? GetStoragePreviewKey(object? storage)
        {
            if (storage == null) return null;
            var t = storage.GetType();
            var candidates = new[] { "PreviewCacheKey", "PreviewKey", "CacheKey", "Id", "FileId", "FileName", "FilePath", "Path" };
            foreach (var name in candidates)
            {
                var prop = t.GetProperty(name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (prop != null)
                {
                    try
                    {
                        var val = prop.GetValue(storage);
                        if (val != null) return val.ToString();
                    }
                    catch { }
                }
            }
            return storage.ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="storageObj"></param>
        /// <param name="localPath"></param>
        /// <returns></returns>
        private async Task<(bool success, string error)> TryInvokeReplaceApisAsync(object storageObj, string localPath)
        {
            if (storageObj == null) return (false, "storageObj 为 null。");
            if (string.IsNullOrWhiteSpace(localPath) || !System.IO.File.Exists(localPath))
                return (false, "本地文件不存在或路径无效。");

            try
            {
                // 优先使用 FileManager（若存在）
                if (_fileManager != null)
                {
                    try
                    {
                        // 调用扩展（扩展会优先反射到实际 API）
                        _fileManager.ReplaceFileContent(storageObj, localPath);
                        return (true, string.Empty);
                    }
                    catch (Exception ex)
                    {
                        // 若 FileManager 路径方式失败，继续尝试 DatabaseManager
                        var msg = $"FileManager 替换失败: {ex.Message}";
                        // 继续向下尝试 DatabaseManager
                    }
                }

                // 尝试 DatabaseManager（以二进制替换为主）
                if (_databaseManager != null)
                {
                    byte[] bytes = await Task.Run(() => System.IO.File.ReadAllBytes(localPath)).ConfigureAwait(false);
                    try
                    {
                        _databaseManager.ReplaceFileBinary(storageObj, bytes);
                        return (true, string.Empty);
                    }
                    catch (Exception ex)
                    {
                        return (false, $"DatabaseManager 替换失败: {ex.Message}");
                    }
                }

                // 退回：若 storageObj 本身实现替换方法则尝试调用
                {
                    var storageType = storageObj.GetType();
                    var pathProp = storageType.GetProperty("FilePath", BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase) ??
                                   storageType.GetProperty("Path", BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                    if (pathProp != null)
                    {
                        var dest = pathProp.GetValue(storageObj) as string;
                        if (!string.IsNullOrWhiteSpace(dest))
                        {
                            try
                            {
                                // 备份与复制
                                var dir = System.IO.Path.GetDirectoryName(dest);
                                if (!string.IsNullOrWhiteSpace(dir) && !System.IO.Directory.Exists(dir))
                                    System.IO.Directory.CreateDirectory(dir);
                                if (System.IO.File.Exists(dest))
                                    System.IO.File.Copy(dest, dest + ".bak", overwrite: true);
                                System.IO.File.Copy(localPath, dest, overwrite: true);
                                return (true, string.Empty);
                            }
                            catch (Exception ex)
                            {
                                return (false, $"直接写入 storage 路径失败: {ex.Message}");
                            }
                        }
                    }
                }

                return (false, "未在 FileManager/DatabaseManager/storage 对象中找到可调用的替换方法。请在后端或 FileManager/DatabaseManager 中实现替换 API（例如：FileManager.ReplaceFileContent(FileStorage, string) 或 DatabaseManager.ReplaceFileBinary(FileStorage, byte[])）。");
            }
            catch (Exception ex)
            {
                return (false, $"替换过程中发生异常: {ex.Message}");
            }
        }

       
        #endregion

        /// <summary>
        /// 架构树节点类
        /// </summary>
        public class CategoryTreeNode
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string DisplayName { get; set; }
            public int Level { get; set; } // 0=主分类, 1=二级子分类, 2=三级子分类...
            public int ParentId { get; set; }
            public object Data { get; set; } // 存储原始数据对象
            public List<CategoryTreeNode> Children { get; set; } = new List<CategoryTreeNode>();
            public string DisplayText { get; set; }
            //public string DisplayText => string.IsNullOrEmpty(DisplayName) ? Name : DisplayName;

            public CategoryTreeNode(int id, string name, string displayName, int level, int parentId, object data)
            {
                Id = id;
                Name = name;
                DisplayText = displayName;
                Level = level;
                ParentId = parentId;
                Data = data;
                Children = new List<CategoryTreeNode>();
            }
        }

        /// <summary>
        /// 通过Row索引查找Grid
        /// </summary>
        private Grid FindGridByRow(DependencyObject parent, int targetRow)
        {
            LogManager.Instance.LogInfo($"开始查找Row={targetRow}的Grid");

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                // 检查是否是Grid且有Grid.Row属性
                if (child is Grid grid)
                {
                    var row = Grid.GetRow(grid);
                    LogManager.Instance.LogInfo($"找到Grid，Row={row}");
                    if (row == targetRow)
                    {
                        LogManager.Instance.LogInfo($"找到目标Grid，Row={targetRow}");
                        return grid;
                    }
                }

                // 递归查找子元素
                var result = FindGridByRow(child, targetRow);
                if (result != null)
                {
                    return result;
                }
            }

            return null;
        }

        /// <summary>
        /// 加载CAD子分类（递归加载多级子分类）
        /// </summary>
        private async Task LoadCadSubcategoriesAsync(int parentId, TreeViewItem parentItem, int level)
        {
            try
            {
                var subcategories = await _databaseManager.GetCadSubcategoriesByParentIdAsync(parentId);// 获取指定父级ID的子分类
                foreach (var subcategory in subcategories)// 遍历子分类
                {
                    // 创建子分类节点
                    string indent = new string(' ', level * 2); // 根据层级添加缩进
                    TreeViewItem subcategoryItem = new TreeViewItem// 创建子分类节点
                    {
                        Header = $"{indent}{subcategory.DisplayName}",// 显示子分类名称
                        Tag = new { Type = "Subcategory", Id = subcategory.Id, Object = subcategory }// 设置Tag属性
                    };
                    await LoadCadGraphicsAsync(subcategory.Id, subcategoryItem); // 加载图元
                    await LoadCadSubcategoriesAsync(subcategory.Id, subcategoryItem, level + 1); // 递归加载子子分类
                    parentItem.Items.Add(subcategoryItem);// 添加子分类节点到父分类节点
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载CAD子分类时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 加载CAD图元
        /// </summary>
        private async Task LoadCadGraphicsAsync(int subcategoryId, TreeViewItem parentItem)
        {
            try
            {
                var files = await _databaseManager.GetFileStorageBySubcategoryIdAsync(subcategoryId);
                foreach (var file in files)
                {
                    TreeViewItem fileItem = new TreeViewItem
                    {
                        Header = $"    {file.LayerName}",
                        Tag = new { Type = "Graphic", Id = file.Id, Object = file }
                    };
                    parentItem.Items.Add(fileItem);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载CAD图元时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 加载SW子分类（递归加载多级子分类）
        /// </summary>
        private async Task LoadSwSubcategoriesAsync(int parentId, TreeViewItem parentItem, int level)
        {
            try
            {
                var subcategories = await _databaseManager.GetSwSubcategoriesByParentIdAsync(parentId);
                foreach (var subcategory in subcategories)
                {
                    // 创建子分类节点
                    string indent = new string(' ', level * 2); // 根据层级添加缩进
                    TreeViewItem subcategoryItem = new TreeViewItem
                    {
                        Header = $"{indent}{subcategory.DisplayName}",
                        Tag = new { Type = "Subcategory", Id = subcategory.Id, Object = subcategory }
                    };

                    // 加载图元
                    await LoadSwGraphicsAsync(subcategory.Id, subcategoryItem);

                    // 递归加载子子分类
                    await LoadSwSubcategoriesAsync(subcategory.Id, subcategoryItem, level + 1);

                    parentItem.Items.Add(subcategoryItem);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载SW子分类时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 加载SW图元
        /// </summary>
        private async Task LoadSwGraphicsAsync(int subcategoryId, TreeViewItem parentItem)
        {
            try
            {
                var graphics = await _databaseManager.GetSwGraphicsBySubcategoryIdAsync(subcategoryId);
                foreach (var graphic in graphics)
                {
                    TreeViewItem graphicItem = new TreeViewItem
                    {
                        Header = $"    {graphic.FileName}",
                        Tag = new { Type = "Graphic", Id = graphic.Id, Object = graphic }
                    };
                    parentItem.Items.Add(graphicItem);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载SW图元时出错: {ex.Message}");
            }
        }


        #endregion

        #region 树节点选中与右键操作


        /// <summary>
        /// 为TreeView添加右键菜单
        /// </summary>
        private void AddContextMenuToTreeView(System.Windows.Controls.TreeView treeView)
        {
            try
            {
                var contextMenu = new System.Windows.Controls.ContextMenu();

                // 新建分类菜单项
                var newItem = new System.Windows.Controls.MenuItem { Header = "新建分类" };
                newItem.Click += NewCategory_MenuItem_Click;
                contextMenu.Items.Add(newItem);

                // 添加子分类菜单项
                var addSubItem = new System.Windows.Controls.MenuItem { Header = "添加子分类" };
                addSubItem.Click += AddSubcategory_MenuItem_Click;
                contextMenu.Items.Add(addSubItem);

                // 修改菜单项
                var editItem = new System.Windows.Controls.MenuItem { Header = "修改" };
                editItem.Click += Edit_MenuItem_Click;
                contextMenu.Items.Add(editItem);

                // 删除菜单项
                var deleteItem = new System.Windows.Controls.MenuItem { Header = "删除" };
                deleteItem.Click += Delete_MenuItem_Click;
                contextMenu.Items.Add(deleteItem);

                // 刷新菜单项
                var RefreshItem = new System.Windows.Controls.MenuItem { Header = "刷新" };
                deleteItem.Click += 刷新文件列表按钮_Click;
                contextMenu.Items.Add(RefreshItem);

                treeView.ContextMenu = contextMenu;

                LogManager.Instance.LogInfo("右键菜单添加成功");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"添加右键菜单时出错: {ex.Message}");
                MessageBox.Show($"添加右键菜单失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>
        /// 解析更新的分类属性
        /// </summary>
        /// <param name="properties"></param>
        /// <returns></returns>
        public static (string Name, string DisplayName, int SortOrder) ParseUpdatedCategoryProperties(List<CategoryPropertyEditModel> properties)
        {
            string name = "";
            string displayName = "";
            int sortOrder = 0;

            foreach (var property in properties)
            {
                CategoryManager.ProcessCategoryProperty(property.PropertyName1, property.PropertyValue1, ref name, ref displayName, ref sortOrder);
                CategoryManager.ProcessCategoryProperty(property.PropertyName2, property.PropertyValue2, ref name, ref displayName, ref sortOrder);
            }

            return (name, displayName, sortOrder);
        }

        /// <summary>
        /// 添加手动刷新文件列表的方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void 刷新文件列表按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_selectedCategoryNode != null)
                {
                    LogManager.Instance.LogInfo("手动刷新文件列表");
                    await LoadFilesForCategoryAsync(_selectedCategoryNode);
                    MessageBox.Show("文件列表已刷新", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("请先选择一个分类", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"刷新文件列表失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion
        #endregion

        #region 按键点击操作

        /// <summary>
        /// 按钮信息类，用于管理按钮与命令的关联
        /// </summary>
        public class ButtonTagCommandInfo
        {
            /// <summary>
            /// cad图元
            /// </summary>
            public FileStorage? fileStorage { get; set; }
            /// <summary>
            /// 命令信息
            /// </summary>
            public ButtonTagCommandInfo? CommandInfo { get; set; }
            /// <summary>
            /// 按钮类型
            /// </summary>
            public string? Type { get; set; }
            /// <summary>
            /// 按钮名称
            /// </summary>
            public string? ButtonName { get; set; }
            /// <summary>
            /// 输入名称
            /// </summary>
            public string? BtnInputText { get; set; }
            /// <summary>
            /// 旋转角度
            /// </summary>
            public double RotateAngle { get; set; }
            /// <summary>
            /// 层颜色索引
            /// </summary>
            public int LayerColorIndex { get; set; }
            /// <summary>
            /// 文件名（不包含路径和扩展名）
            /// </summary>
            public string? FileName { get; set; }
            /// <summary>
            /// 块名称（从文件名中提取）
            /// </summary>
            public string? BlockName { get; set; }
            /// <summary>
            /// 文件完整路径
            /// </summary>
            public string? FilePath { get; set; }
            /// <summary>
            /// 对应的Command方法名
            /// </summary>
            public string? CommandMethodName { get; set; }
            /// <summary>
            /// 按钮显示文本
            /// </summary>
            public string? DisplayText { get; set; }
            /// <summary>
            /// 按钮类型（自定义按钮、工艺按钮等）
            /// </summary>
            public string? ButtonType { get; set; }
            /// <summary>
            /// 相关参数
            /// </summary>
            public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();
            /// <summary>
            /// 所属分类（二级文件夹名称）
            /// </summary>
            public string? Category { get; set; }
            /// <summary>
            /// 其他Object对象
            /// </summary>
            public object? Object { get; set; } // 用于存储其他对象
        }

        /// <summary>
        /// 预定义按钮点击事件处理
        /// </summary>
        /// <param Name="sender"></param>
        /// <param Name="e"></param>
        private void PredefinedButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button btn && btn.Tag is ButtonTagCommandInfo tagInfo)
                {
                    LogManager.Instance.LogInfo($"点击预定义按钮: {tagInfo.ButtonName}");

                    // 通过统一管理器获取并执行对应的命令
                    var command = UnifiedCommandManager.GetCommand(tagInfo.ButtonName);
                    if (command != null)
                    {
                        try
                        {
                            command.Invoke();
                            LogManager.Instance.LogInfo($"成功执行按钮命令: {tagInfo.ButtonName}");
                        }
                        catch (Exception invokeEx)
                        {
                            LogManager.Instance.LogInfo($"执行按钮命令时出错: {invokeEx.Message}");
                            System.Windows.MessageBox.Show($"执行命令 '{tagInfo.ButtonName}' 时出错: {invokeEx.Message}");
                        }
                    }
                    else
                    {
                        LogManager.Instance.LogInfo($"未找到按钮 '{tagInfo.ButtonName}' 对应的命令");
                        System.Windows.MessageBox.Show($"未找到按钮 '{tagInfo.ButtonName}' 对应的命令");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"处理预定义按钮点击事件时出错: {ex.Message}");
                System.Windows.MessageBox.Show($"处理按钮点击事件时出错: {ex.Message}");
            }
        }

        #region 共用图按键处理方法...

        private void 共用条件_Btn_Click(object sender, RoutedEventArgs e)
        {
            //MigrateResourcesToDatabaseAsync();//数据迁移
        }

        private void 所有条件开关_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 设备开关_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        #endregion

        #region 工艺按键处理方法...
        private void 纯化水_Btn_Clic(object sender, RoutedEventArgs e)
        {
            /// 获取按钮的命令
            var command = UnifiedCommandManager.GetCommand("纯化水");
            command?.Invoke();//执行命令
        }

        private void 纯蒸汽_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("纯蒸汽");
            command?.Invoke();
        }

        private void 注射用水_Btn_Click(object sender, RoutedEventArgs e)
        {

            var command = UnifiedCommandManager.GetCommand("注射用水");
            command?.Invoke();
        }

        private void 凝结回水_Btn_Click(object sender, RoutedEventArgs e)
        {

            var command = UnifiedCommandManager.GetCommand("凝结回水");
            command?.Invoke();
        }

        private void 氧气_Btn_Click(object sender, RoutedEventArgs e)
        {

            var command = UnifiedCommandManager.GetCommand("氧气");
            command?.Invoke();
        }

        private void 氮气_Btn_Click(object sender, RoutedEventArgs e)
        {

            var command = UnifiedCommandManager.GetCommand("氮气");
            command?.Invoke();
        }

        private void 二氧化碳_Btn_Click(object sender, RoutedEventArgs e)
        {

            var command = UnifiedCommandManager.GetCommand("二氧化碳");
            command?.Invoke();
        }

        private void 无菌压缩空气_Btn_Click(object sender, RoutedEventArgs e)
        {

            var command = UnifiedCommandManager.GetCommand("无菌压缩空气");
            command?.Invoke();
        }

        private void 仪表压缩空气_Btn_Click(object sender, RoutedEventArgs e)
        {

            var command = UnifiedCommandManager.GetCommand("仪表压缩空气");
            command?.Invoke();
        }

        private void 低压蒸汽_Btn_Click(object sender, RoutedEventArgs e)
        {

            var command = UnifiedCommandManager.GetCommand("低压蒸汽");
            command?.Invoke();
        }

        private void 低温循环上水_Btn_Click(object sender, RoutedEventArgs e)
        {

            var command = UnifiedCommandManager.GetCommand("低温循环上水");
            command?.Invoke();
        }

        private void 常温循环上水_Btn_Click(object sender, RoutedEventArgs e)
        {

            var command = UnifiedCommandManager.GetCommand("常温循环上水");
            command?.Invoke();
        }

        private void 设备表导入_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 设备表导出_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 区域开关_Btn_Click(object sender, RoutedEventArgs e)
        {

        }
        #endregion

        #region 管理员按键处理方法...

        /// <summary>
        /// 加载CAD数据库按钮点击事件
        /// </summary>
        private async void LoadCadDatabase_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 设置当前数据库类型
                _currentDatabaseType = "CAD";
                LogManager.Instance.LogInfo("设置数据库类型为: " + _currentDatabaseType);
                LogManager.Instance.LogInfo("=== 开始加载CAD数据库 ===");
                if (!_useDatabaseMode || _databaseManager == null || !_databaseManager.IsDatabaseAvailable)
                {
                    System.Windows.MessageBox.Show("数据库不可用，请检查数据库连接配置");
                    LogManager.Instance.LogInfo("数据库不可用，请检查数据库连接配置");
                    return;
                }
                _cadStoragePath = await _databaseManager.GetConfigValueAsync("cad_storage_path");  // 获取CAD存储路径
                if (string.IsNullOrEmpty(_cadStoragePath))
                {
                    _cadStoragePath = System.IO.Path.Combine(AppPath, "CadFiles");
                }
                System.IO.Directory.CreateDirectory(_cadStoragePath); // 确保存储路径存在

                // 加载并显示CAD分类树
                await InitializeCategoryTreeAsync();

            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"加载CAD数据库时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 加载SW数据库按钮点击事件
        /// </summary>
        private async void LoadSwDatabase_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_databaseManager == null)
                {
                    System.Windows.MessageBox.Show("数据库未初始化");
                    return;
                }

                // 设置当前数据库类型
                _currentDatabaseType = "SW";

                // 获取SW存储路径
                _swStoragePath = await _databaseManager.GetConfigValueAsync("sw_storage_path");
                if (string.IsNullOrEmpty(_swStoragePath))
                {
                    _swStoragePath = System.IO.Path.Combine(AppPath, "SwFiles");
                }

                // 确保存储路径存在
                System.IO.Directory.CreateDirectory(_swStoragePath);

                // 加载并显示SW分类树
                //await LoadAndDisplayCategoryTreeAsync();

                System.Windows.MessageBox.Show("SW数据库加载成功");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"加载SW数据库时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 新建分类菜单项点击事件
        /// </summary>
        private async void NewCategory_MenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_currentDatabaseType))
                {
                    System.Windows.MessageBox.Show("请先加载数据库");
                    return;
                }
                if (_currentDatabaseType == "CAD")
                {
                    _currentOperation = ManagementOperationType.AddCategory;
                    InitializeCategoryPropertiesForCategory();//初始化新建分类界面
                    _selectedCategoryNode = null; // 清除选中节点，表示添加主分类

                    //ShowNewCategoryTips();// 显示提示信息
                    LogManager.Instance.LogInfo("初始化新建主分类界面");
                }
                else if (_currentDatabaseType == "SW")
                {
                    _currentOperation = ManagementOperationType.AddCategory; // 创建新的SW分类
                    InitializeCategoryPropertiesForCategory();
                }

                MessageBox.Show("请在表格中填写分类属性，然后点击'应用属性'按钮", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                //System.Windows.MessageBox.Show("分类创建成功");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"新建分类时出错: {ex.Message}");
                MessageBox.Show($"初始化分类添加失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 添加子分类菜单项点击事件
        /// </summary>
        private async void AddSubcategory_MenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_currentDatabaseType))
                {
                    System.Windows.MessageBox.Show("请先加载数据库");
                    return;
                }

                if (_currentDatabaseType == "CAD")
                {
                    if (_selectedCategoryNode != null)
                    {
                        _currentOperation = ManagementOperationType.AddSubcategory;
                        // 初始化子分类属性编辑界面，预填父分类ID
                        InitializeSubcategoryPropertiesForEditing(_selectedCategoryNode);
                        // 显示提示信息
                        //ShowNewSubcategoryTips(_selectedCategoryNode);

                        LogManager.Instance.LogInfo($"初始化添加子分类界面，父节点: {_selectedCategoryNode.DisplayText}");
                    }
                    else
                    {
                        MessageBox.Show("请先选择一个分类或子分类", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加子分类时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 修改菜单项点击事件
        /// </summary>
        private async void Edit_MenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (_currentNodeId <= 0)
            {
                System.Windows.MessageBox.Show("请先选择一个项目");
                return;
            }
            try
            {
                if (_selectedCategoryNode != null)
                {
                    // 显示当前选中节点的属性用于编辑
                    DisplayNodePropertiesForEditing(_selectedCategoryNode);
                    _currentOperation = ManagementOperationType.None; // 设置为编辑模式

                    LogManager.Instance.LogInfo($"初始化编辑分类界面: {_selectedCategoryNode.DisplayText}");
                }
                else
                {
                    MessageBox.Show("请先选择一个分类或子分类", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化编辑分类失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        /// <summary>
        /// 删除菜单项点击事件
        /// </summary>
        private async void Delete_MenuItem_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                if (_selectedCategoryNode != null)
                {
                    string nodeName = _selectedCategoryNode.DisplayText;

                    if (MessageBox.Show($"确定要删除分类 '{nodeName}' 吗？\n注意：删除操作不可恢复！",
                                      "确认删除", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        await DeleteCategoryNodeAsync(_selectedCategoryNode);
                    }
                }
                else
                {
                    MessageBox.Show("请先选择一个分类或子分类", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除分类失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 删除分类节点方法 
        /// </summary>
        /// <param name="nodeToDelete"></param>
        /// <returns></returns>
        private async Task DeleteCategoryNodeAsync(CategoryTreeNode nodeToDelete)
        {
            try
            {
                if (nodeToDelete == null) return;

                bool success = false;

                // 根据节点层级执行不同的删除操作
                if (nodeToDelete.Level == 0)
                {
                    // 删除主分类
                    success = await _categoryManager.DeleteMainCategoryAsync(nodeToDelete);
                }
                else
                {
                    // 删除子分类
                    success = await _categoryManager.DeleteSubcategoryAsync(nodeToDelete);
                }

                if (success)
                {
                    // 刷新架构树
                    await _categoryManager.RefreshCategoryTreeAsync(_selectedCategoryNode, CategoryTreeView, _categoryTreeNodes, _databaseManager);
                    _selectedCategoryNode = null; // 清除选中节点
                    InitializeCategoryPropertyGrid(); // 清空属性编辑区

                    MessageBox.Show("分类删除成功", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    throw new Exception("删除操作失败");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除分类失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 刷新架构树按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void 刷新架构树按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await _categoryManager.RefreshCategoryTreeAsync(_selectedCategoryNode, CategoryTreeView, _categoryTreeNodes, _databaseManager);
                MessageBox.Show("架构树刷新成功", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"刷新架构树失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 展开折叠架构按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 展开_折叠架构_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (CategoryTreeView == null || CategoryTreeView.Items == null)
                    return;

                // 检测当前是否存在已展开的顶层节点
                bool anyExpanded = false;
                foreach (var item in CategoryTreeView.Items)
                {
                    var tvi = GetTreeViewItem(CategoryTreeView, item);
                    if (tvi != null && tvi.IsExpanded)
                    {
                        anyExpanded = true;
                        break;
                    }
                }

                // 如果有已展开的节点 -> 折叠所有；否则展开所有
                if (anyExpanded)
                {
                    foreach (var item in CategoryTreeView.Items)
                    {
                        var tvi = GetTreeViewItem(CategoryTreeView, item);
                        if (tvi != null)
                            tvi.IsExpanded = false;
                    }

                    // 更新按钮显示提示（如果按钮存在）
                    if (sender is Button btn)
                        btn.Content = "展开/折叠架构";
                    LogManager.Instance.LogInfo("已折叠架构树所有节点");
                }
                else
                {
                    foreach (var item in CategoryTreeView.Items)
                    {
                        var tvi = GetTreeViewItem(CategoryTreeView, item);
                        if (tvi != null)
                        {
                            tvi.IsExpanded = true;
                            // 递归展开所有子节点
                            ExpandAllChildren(tvi);
                        }
                    }

                    if (sender is Button btn)
                        btn.Content = "折叠架构";
                    LogManager.Instance.LogInfo("已展开架构树所有节点");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"切换展开/折叠架构失败: {ex.Message}");
                MessageBox.Show($"操作失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 选择文件按钮点击事件
        /// </summary>
        private async void SelectFile_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_selectedCategoryNode == null)
                {
                    MessageBox.Show("请先在架构树中选择一个分类", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 初始化文件上传界面
                InitializeFileUploadInterface();
                var openFileDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "选择要上传的文件",
                    Filter = "所有文件 (*.*)|*.*|DWG文件 (*.dwg)|*.dwg|图片文件 (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp|文档文件 (*.pdf;*.doc;*.docx)|*.pdf;*.doc;*.docx",
                    Multiselect = false
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    _selectedFilePath = openFileDialog.FileName;

                    // 显示文件信息
                    DisplayFileInfo(_selectedFilePath);

                    // 初始化属性编辑界面
                    AddFileInitializeFilePropertiesGrid();
                }
                MessageBox.Show("文件已上传到服务器，可以编辑属性后点击'完成添加'", "成功",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择文件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 添加预览图按钮点击事件
        /// </summary>
        private async void SelectViewImage_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "选择预览图片",
                    Filter = "图片文件 (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif",
                    Multiselect = false
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    _selectedPreviewImagePath = openFileDialog.FileName;

                    // 显示预览图片
                    DisplayPreviewImage(_selectedPreviewImagePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择预览图片失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 删除图元按钮点击事件
        /// </summary>
        private async void DeleteGraphic_Btn_Click(object sender, RoutedEventArgs e)
        {
            var selected = StroageFileDataGrid.SelectedItem as FileStorage;// 获取选中的行 获取选中的图元
            if (selected == null)
            {
                MessageBox.Show("未选中要删除的图元。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var confirm = MessageBox.Show(
                $"确定删除图元：{selected.DisplayName ?? selected.FileName} ?\n该操作将删除所有关联数据且不可恢复。",
                "删除确认",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirm != MessageBoxResult.Yes) return;

            DeleteGraphic_Btn.IsEnabled = false;

            try
            {
                if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
                {
                    MessageBox.Show("数据库不可用，无法删除。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                bool ok = await _databaseManager.DeleteCadGraphicCascadeAsync(selected.Id, physicalDelete: true);
                if (ok)
                {
                    // 尝试从 ItemsSource 移除并刷新
                    var src = StroageFileDataGrid.ItemsSource;

                    if (src is IList<FileStorage> list)
                    {
                        list.Remove(selected);
                    }
                    else if (src is System.Collections.IList nonGenericList)
                    {
                        nonGenericList.Remove(selected);
                    }
                    else if (src is System.Windows.Data.CollectionView view && view.SourceCollection is System.Collections.IList viewList)
                    {
                        viewList.Remove(selected);
                        view.Refresh();
                    }
                    else
                    {
                        // 退化策略：直接刷新当前分类文件列表
                        await RefreshFilesForCurrentCategoryAsync();
                    }

                    StroageFileDataGrid.Items.Refresh();
                    CategoryPropertiesDataGrid.ItemsSource = null;

                    MessageBox.Show("删除成功。", "信息", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("删除失败，请查看日志。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除过程中发生错误：{ex.Message}", "异常", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                DeleteGraphic_Btn.IsEnabled = true;
            }
        }

        /// <summary>
        /// 还原初始值按钮点击事件
        /// </summary>
        private void ResetToInitial_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _currentOperation = ManagementOperationType.None;
                InitializeCategoryPropertyGrid();
                ClearFileUploadInterface();
                MessageBox.Show("操作已取消", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"取消操作失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 应用属性按钮点击事件
        /// </summary>
        private async void ApplyProperties_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                bool success = false;

                switch (_currentOperation)
                {
                    case ManagementOperationType.AddCategory:
                        success = await _categoryManager.ApplyCategoryPropertiesAsync(CategoryPropertiesDataGrid);
                        break;
                    case ManagementOperationType.AddSubcategory:
                        success = await _categoryManager.ApplySubcategoryPropertiesAsync(CategoryPropertiesDataGrid);
                        break;
                    case ManagementOperationType.None:
                        // 如果没有明确的操作类型，可能是编辑操作
                        if (_selectedCategoryNode != null)
                        {
                            success = await _categoryManager.UpdateCategoryPropertiesAsync(CategoryPropertiesDataGrid, _selectedCategoryNode);
                        }
                        else
                        {
                            MessageBox.Show("请先选择要操作的分类或子分类", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                        break;
                    default:
                        MessageBox.Show("未知操作类型", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                }

                if (success)
                {
                    // 重置操作状态
                    _currentOperation = ManagementOperationType.None;
                    InitializeCategoryPropertyGrid();

                    // 刷新架构树显示
                    await _categoryManager.RefreshCategoryTreeAsync(_selectedCategoryNode, CategoryTreeView, _categoryTreeNodes, _databaseManager);

                    MessageBox.Show("操作成功，架构树已更新", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("操作失败，请检查输入数据", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 添加文件按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddFile_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _ = UploadFileAndSaveToDatabase();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"完成添加失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>
        /// 将管理上下文菜单附加到指定按钮，启用与文件存储管理相关的操作。
        /// </summary>
        /// <remarks>如果按钮为空，则该方法不执行任何操作。在执行过程中遇到的任何错误附加上下文菜单将被自动忽略，以防止UI呈现问题。</remarks>
        /// <param name="btn">将附加上下文菜单的按钮控件。不能为null。</param>
        /// <param name="storage">一个可选的FileStorage实例，为菜单操作提供上下文。如果没有存储，则可能为空</param>
        private void AttachAdminContextMenu(System.Windows.Controls.Button btn, FileStorage? storage)
        {
            if (btn == null) return;
            try
            {
                var contextMenu = new System.Windows.Controls.ContextMenu();
                var miReplace = new System.Windows.Controls.MenuItem { Header = "替换图元" };
                // 存储按钮引用，事件处理时从 btn.Tag 获取 FileStorage/Info
                miReplace.Tag = btn;
                miReplace.Click += ReplaceFileMenuItem_Click;
                contextMenu.Items.Add(miReplace);
                btn.ContextMenu = contextMenu;
            }
            catch
            {
                // 忽略 UI 附加错误，避免影响面板渲染
            }
        }



        public async Task UploadFileAndSaveToDatabase()
        {
            // 从 UI 状态构造 DTO，然后委托到核心方法
            try
            {
                if (_selectedCategoryNode == null)
                {
                    MessageBox.Show("请先在分类树中选择目标分类。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(_selectedFilePath) || !File.Exists(_selectedFilePath))
                {
                    MessageBox.Show("请选择并缓存要上传的文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 构造基础 DTO（只包含必要字段，保留 UI 可能编辑的字段）
                var dto = new ImportEntityDto
                {
                    FileStorage = new FileStorage
                    {
                        CategoryId = _selectedCategoryNode.Id,
                        CategoryType = _selectedCategoryNode.Level == 0 ? "main" : "sub",
                        FilePath = _selectedFilePath,
                        FileName = Path.GetFileName(_selectedFilePath),
                        DisplayName = _currentFileStorage?.DisplayName ?? Path.GetFileNameWithoutExtension(_selectedFilePath),
                        Description = _currentFileStorage?.Description,
                        CreatedBy = Environment.UserName
                    },
                    FileAttribute = _currentFileAttribute ?? new FileAttribute()
                };

                // 把用户在属性网格中填写的属性应用到 dto.FileAttribute / dto.FileStorage
                var gridProps = CategoryPropertiesDataGrid?.ItemsSource as List<CategoryPropertyEditModel>;
                if (gridProps != null)
                {
                    foreach (var p in gridProps)
                    {
                        // 两列都尝试设置
                        SetFileAttributeProperty(dto.FileAttribute, p.PropertyName1 ?? string.Empty, p.PropertyValue1 ?? string.Empty);
                        SetFileAttributeProperty(dto.FileAttribute, p.PropertyName2 ?? string.Empty, p.PropertyValue2 ?? string.Empty);
                        SetFileStorageProperty(dto.FileStorage, p.PropertyName1 ?? string.Empty, p.PropertyValue1 ?? string.Empty);
                        SetFileStorageProperty(dto.FileStorage, p.PropertyName2 ?? string.Empty, p.PropertyValue2 ?? string.Empty);
                    }
                }

                await UploadFileAndSaveToDatabase(dto).ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"UploadFileAndSaveToDatabase() 无参入口失败: {ex.Message}");
                MessageBox.Show($"上传失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }



        /// <summary>
        /// 上传文件并保存到数据库
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        //public async Task UploadFileAndSaveToDatabase()
        //{
        //    List<string> uploadedFiles = new List<string>(); // 记录已上传的文件路径，用于回滚
        //    FileStorage savedFileStorage = null; // 记录已保存的文件记录
        //    FileAttribute savedFileAttribute = null; // 记录已保存的属性记录
        //    bool transactionSuccess = false;

        //    try
        //    {
        //        if (string.IsNullOrEmpty(_selectedFilePath) || _selectedCategoryNode == null)
        //        {
        //            throw new Exception("文件路径或分类节点为空");
        //        }

        //        // 1. 获取文件信息
        //        var fileInfo = new FileInfo(_selectedFilePath);
        //        string fileName = fileInfo.Name;
        //        string displayName = Path.GetFileNameWithoutExtension(fileName);
        //        string description = $"上传文件: {fileName}";
        //        var fileStorage = new FileStorage();
        //        // 2. 使用FileManager上传主文件到服务器指定路径
        //        using (var fileStream = File.OpenRead(_selectedFilePath))
        //        {
        //            fileStorage = await _fileManager.UploadFileAsync(_databaseManager,
        //                _selectedCategoryNode.Id,
        //                _selectedCategoryNode.Level == 0 ? "main" : "sub",
        //                fileName,
        //                fileStream,
        //                description,
        //                Environment.UserName
        //            );

        //            // 保存上传后的文件信息
        //            _currentFileStorage = fileStorage;
        //            savedFileStorage = fileStorage;
        //            uploadedFiles.Add(fileStorage.FilePath); // 记录已上传的文件路径
        //        }

        //        // 3. 如果有预览图片，上传预览图片
        //        string previewStoredPath = null;
        //        if (!string.IsNullOrEmpty(_selectedPreviewImagePath) && File.Exists(_selectedPreviewImagePath))
        //        {
        //            var previewInfo = new FileInfo(_selectedPreviewImagePath);
        //            string previewFileName = $"{Path.GetFileNameWithoutExtension(_selectedPreviewImagePath)}_preview{previewInfo.Extension}";

        //            using (var previewStream = File.OpenRead(_selectedPreviewImagePath))
        //            {
        //                // 生成预览文件存储路径
        //                string previewStoredName = $"{Guid.NewGuid()}{previewInfo.Extension}";
        //                previewStoredPath = Path.Combine(
        //                    Path.GetDirectoryName(_currentFileStorage.FilePath),
        //                    previewStoredName);

        //                // 复制预览图片到同一目录
        //                File.Copy(_selectedPreviewImagePath, previewStoredPath, true);

        //                _currentFileStorage.PreviewImageName = previewStoredName;
        //                _currentFileStorage.PreviewImagePath = previewStoredPath;
        //                uploadedFiles.Add(previewStoredPath); // 记录预览文件路径
        //            }
        //        }

        //        // 4. 创建文件属性对象
        //        _currentFileAttribute = new FileAttribute
        //        {
        //            //FileStorageId = _currentFileStorage.Id,
        //            FileName = _currentFileStorage.FileName,
        //            CreatedAt = DateTime.Now,
        //            UpdatedAt = DateTime.Now
        //        };

        //        // 5. 从属性编辑网格中获取属性值
        //        var gridProperties = CategoryPropertiesDataGrid.ItemsSource as List<CategoryPropertyEditModel>;
        //        if (gridProperties != null)
        //        {
        //            foreach (var property in gridProperties)
        //            {
        //                // 使用实例调用
        //                SetFileAttributeProperty(_currentFileAttribute, property.PropertyName1, property.PropertyValue1);
        //                SetFileAttributeProperty(_currentFileAttribute, property.PropertyName2, property.PropertyValue2);
        //            }
        //        }
        //        if (_currentFileAttribute.FileName == null)
        //        {
        //            MessageBox.Show("请填写文件名称", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        //            return;
        //        }

        //        // 6. 保存文件属性到数据库
        //        int attributeResult = await _databaseManager.AddFileAttributeAsync(_currentFileAttribute);
        //        if (attributeResult <= 0)
        //        {
        //            LogManager.Instance.LogInfo("保存文件属性失败");

        //        }
        //        else
        //        {
        //            LogManager.Instance.LogInfo("保存文件属性到数据库:成功");
        //        }
        //        //获取文件属性ID
        //        _currentFileAttribute = await _databaseManager.GetFileAttributeAsync(_currentFileStorage.DisplayName);
        //        if (_currentFileAttribute == null || _currentFileAttribute.Id == null)
        //        {
        //            LogManager.Instance.LogInfo("获取文件属性ID失败");
        //            // 发生异常，需要回滚操作
        //            await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, _currentFileStorage, _currentFileAttribute);
        //            return;
        //        }
        //        _currentFileStorage.FileAttributeId = _currentFileAttribute.Id;

        //        //新加文件到数据库中
        //        var fileResult = await _databaseManager.AddFileStorageAsync(_currentFileStorage);
        //        if (fileResult == 0)
        //        {
        //            LogManager.Instance.LogInfo("保存文件记录到数据库:失败");
        //            // 发生异常，需要回滚操作
        //            await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, _currentFileStorage, _currentFileAttribute);
        //            return;
        //        }
        //        else
        //        {
        //            LogManager.Instance.LogInfo("保存文件记录到数据库:成功");
        //        }
        //        ;
        //        _currentFileStorage = await _databaseManager.GetFileStorageAsync(_currentFileStorage.FileHash);//获取文件的基本信息
        //        _currentFileAttribute.FileStorageId = _currentFileStorage.Id;//文件属性ID

        //        await _databaseManager.UpdateFileAttributeAsync(_currentFileAttribute);//更新文件属性
        //        // 8. 处理标签信息
        //        // await FileManager.ProcessFileTags(_currentFileStorage.Id, properties);

        //        // 9. 更新分类统计
        //        var updateBool = await _databaseManager.UpdateCategoryStatisticsAsync(
        //            _currentFileStorage.CategoryId,
        //            _currentFileStorage.CategoryType);

        //        // 如果所有操作都成功，标记事务成功
        //        transactionSuccess = true;
        //        // 11. 刷新分类树和界面显示
        //        // 替换为：
        //        //await RefreshCurrentCategoryFilesAsync();
        //        await RefreshCurrentCategoryDisplayAsync(_selectedCategoryNode);
        //        MessageBox.Show($"文件已成功上传并保存到服务器指定路径\n文件路径: {_currentFileStorage.FilePath}",
        //            "成功", MessageBoxButton.OK, MessageBoxImage.Information);
        //    }
        //    catch (Exception ex)
        //    {
        //        // 发生异常，需要回滚操作
        //        await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, _currentFileStorage, _currentFileAttribute);
        //        throw new Exception($"文件上传和数据库保存失败: {ex.Message}", ex);
        //    }
        //    finally
        //    {
        //        // 如果事务失败，执行回滚
        //        if (!transactionSuccess)
        //        {
        //            await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, _currentFileStorage, _currentFileAttribute);
        //        }
        //    }
        //}

        public async Task UploadFileAndSaveToDatabase(ImportEntityDto dto)
        {
            if (dto == null) throw new ArgumentNullException(nameof(dto));
            if (dto.FileStorage == null) throw new ArgumentException("dto.FileStorage 不能为空", nameof(dto));
            if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
            {
                MessageBox.Show("数据库未连接或不可用，无法保存。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (_fileManager == null)
            {
                MessageBox.Show("文件管理器未初始化，无法上传文件。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            await _uploadSemaphore.WaitAsync().ConfigureAwait(false);
            var uploadedFiles = new List<string>();
            bool success = false;

            try
            {
                LogManager.Instance.LogInfo("开始上传文件并保存到数据库");

                // 1) 上传主文件（FileManager 负责物理存储并返回带完整信息的 FileStorage）
                var sourcePath = dto.FileStorage.FilePath;
                if (string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath))
                    throw new FileNotFoundException("要上传的文件不存在", sourcePath);

                using (var fs = File.OpenRead(sourcePath))
                {
                    var uploaded = await _fileManager.UploadFileAsync(
                        _databaseManager,
                        dto.FileStorage.CategoryId,
                        dto.FileStorage.CategoryType ?? "sub",
                        Path.GetFileName(sourcePath),
                        fs,
                        dto.FileStorage.Description ?? string.Empty,
                        dto.FileStorage.CreatedBy ?? Environment.UserName
                    ).ConfigureAwait(false);

                    if (uploaded == null)
                        throw new Exception("文件上传失败，返回信息为空。");

                    // 用上传后的元数据更新 DTO（保留用户在 DTO/界面上设置的 DisplayName/FileName）
                    dto.FileStorage.FileStoredName = uploaded.FileStoredName ?? uploaded.FileStoredName;
                    dto.FileStorage.FilePath = uploaded.FilePath;
                    dto.FileStorage.FileHash = uploaded.FileHash;
                    dto.FileStorage.FileSize = uploaded.FileSize;
                    dto.FileStorage.FileType = uploaded.FileType;
                    uploadedFiles.Add(dto.FileStorage.FilePath);
                    LogManager.Instance.LogInfo($"文件上传成功: {dto.FileStorage.FilePath}");
                }

                // 2) 上传/复制预览图片（如存在）
                if (!string.IsNullOrEmpty(dto.PreviewImagePath) && File.Exists(dto.PreviewImagePath))
                {
                    try
                    {
                        var previewInfo = new FileInfo(dto.PreviewImagePath);
                        string previewStoredName = $"{Guid.NewGuid()}{previewInfo.Extension}";
                        string previewStoredPath = Path.Combine(Path.GetDirectoryName(dto.FileStorage.FilePath) ?? Path.GetTempPath(), previewStoredName);

                        File.Copy(dto.PreviewImagePath, previewStoredPath, true);
                        dto.FileStorage.PreviewImageName = previewStoredName;
                        dto.FileStorage.PreviewImagePath = previewStoredPath;
                        uploadedFiles.Add(previewStoredPath);
                        LogManager.Instance.LogInfo($"预览图片复制成功: {previewStoredPath}");
                    }
                    catch (Exception exPreview)
                    {
                        LogManager.Instance.LogWarning($"复制预览图片失败（继续）：{exPreview.Message}");
                    }
                }
               
                // 3) 准备 FileAttribute（确保时间与必要字段）
                dto.FileAttribute.FileName = dto.FileStorage.FileName ?? dto.FileStorage.DisplayName ?? dto.FileStorage.FileStoredName;
                dto.FileAttribute.CreatedAt = dto.FileAttribute.CreatedAt == default ? DateTime.Now : dto.FileAttribute.CreatedAt;
                dto.FileAttribute.UpdatedAt = dto.FileAttribute.UpdatedAt == default ? DateTime.Now : dto.FileAttribute.UpdatedAt;
                // 4) 原子写入数据库：storage + attribute（使用 DatabaseManager 的原子方法）
                var (storageId, attributeId) = await _databaseManager.AddFileStorageAndAttributeAsync(dto.FileStorage, dto.FileAttribute).ConfigureAwait(false);

                if (storageId <= 0 || attributeId <= 0)
                    throw new Exception("数据库写入失败：未返回有效的存储/属性 Id");

                // 5) 更新本地 DTO 与界面状态
                dto.FileStorage.Id = storageId;
                dto.FileStorage.FileAttributeId = attributeId.ToString();
                dto.FileAttribute.Id = attributeId;
                dto.FileAttribute.FileStorageId = storageId;

                // 6) 更新分类统计（异步，可不阻塞主流程）
                try
                {
                    await _databaseManager.UpdateCategoryStatisticsAsync(dto.FileStorage.CategoryId, dto.FileStorage.CategoryType ?? "sub").ConfigureAwait(false);
                }
                catch (Exception exStat)
                {
                    LogManager.Instance.LogWarning($"更新分类统计失败（非致命）：{exStat.Message}");
                }

                success = true;
                LogManager.Instance.LogInfo($"文件及属性已成功保存到数据库 (StorageId={storageId}, AttributeId={attributeId})");

                // 7) 刷新 UI（在 UI 线程）
                await Dispatcher.InvokeAsync(async () =>
                {
                    try
                    {
                        // 刷新当前分类列表
                        if (_selectedCategoryNode != null)
                            await RefreshCurrentCategoryDisplayAsync(_selectedCategoryNode);
                    }
                    catch (Exception exUi)
                    {
                        LogManager.Instance.LogWarning($"刷新界面失败: {exUi.Message}");
                    }
                });

                MessageBox.Show($"文件上传并保存成功。\n路径: {dto.FileStorage.FilePath}", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"UploadFileAndSaveToDatabase(dto) 失败: {ex.Message}");
                // 回滚已经上传的文件与数据库残留（FileManager.RollbackFileUpload 需能处理部分已完成步骤）
                try
                {
                    await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, dto.FileStorage, dto.FileAttribute).ConfigureAwait(false);
                    LogManager.Instance.LogInfo("已执行上传回滚 (文件/数据库回退)。");
                }
                catch (Exception exRb)
                {
                    LogManager.Instance.LogError($"回滚时失败: {exRb.Message}");
                }

                // 把错误抛回或显示友好消息
                MessageBox.Show($"上传/保存失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _uploadSemaphore.Release();
            }
        }


        /// <summary>
        /// 上传文件并保存到数据库
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        //public async Task UploadFileAndSaveToDatabase(ImportEntityDto dto)
        //{

        //    await _uploadSemaphore.WaitAsync();
        //    List<string> uploadedFiles = new List<string>(); // 上传成功的文件列表
        //    bool transactionSuccess = false; // 事务成功标志

        //    try
        //    {
        //        // 基本校验
        //        if (dto == null || dto.FileStorage == null || string.IsNullOrEmpty(dto.FileStorage.FilePath))
        //            throw new ArgumentException("ImportEntityDto 参数无效或文件路径为空");
        //        if (_databaseManager == null)
        //            throw new InvalidOperationException("数据库管理器未初始化");

        //        // 1. 上传主文件
        //        var fileInfo = new FileInfo(dto.FileStorage.FilePath);
        //        string fileName = fileInfo.Name;
        //        string description = dto.FileStorage.Description ?? $"上传文件: {fileName}";
        //        using (var fileStream = File.OpenRead(dto.FileStorage.FilePath))
        //        {
        //            var uploadedStorage = await _fileManager.UploadFileAsync(
        //                _databaseManager,
        //                dto.FileStorage.CategoryId,
        //                dto.FileStorage.CategoryType,
        //                fileName,
        //                fileStream,
        //                description,
        //                dto.FileStorage.CreatedBy ?? Environment.UserName
        //            );

        //            if (uploadedStorage == null)
        //            {
        //                await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, dto.FileStorage, dto.FileAttribute);
        //                throw new Exception("上传文件失败：返回的上传信息为空。");
        //            }

        //            // 用上传后的路径和信息覆盖，但保留用户修改的 DisplayName 和 FileName
        //            dto.FileStorage.FileStoredName = uploadedStorage.FileStoredName;
        //            dto.FileStorage.FilePath = uploadedStorage.FilePath;
        //            dto.FileStorage.FileSize = uploadedStorage.FileSize;
        //            dto.FileStorage.FileType = uploadedStorage.FileType;
        //            dto.FileStorage.FileHash = uploadedStorage.FileHash;
        //            // 保留用户在界面上修改的 DisplayName 和 FileName，不要覆盖
        //            // dto.FileStorage.DisplayName = Path.GetFileNameWithoutExtension(dto.FileStorage.FileName) ?? dto.FileStorage.DisplayName; // 已移除覆盖逻辑
        //            // dto.FileStorage.FileName = fileName; // 保留原有的 FileName
        //            uploadedFiles.Add(dto.FileStorage.FilePath);
        //        }

        //        // 2. 上传预览图片（如果有）
        //        if (!string.IsNullOrEmpty(dto.PreviewImagePath) && File.Exists(dto.PreviewImagePath))
        //        {
        //            var previewInfo = new FileInfo(dto.PreviewImagePath);
        //            string previewStoredName = $"{Guid.NewGuid()}{previewInfo.Extension}";
        //            string previewStoredPath = Path.Combine(
        //                Path.GetDirectoryName(dto.FileStorage.FilePath),
        //                previewStoredName);

        //            File.Copy(dto.PreviewImagePath, previewStoredPath, true);
        //            dto.FileStorage.PreviewImageName = previewStoredName;
        //            dto.FileStorage.PreviewImagePath = previewStoredPath;
        //            uploadedFiles.Add(previewStoredPath);
        //        }

        //        // 3. 保存文件属性到数据库
        //        dto.FileAttribute.FileName = dto.FileStorage.FileName;
        //        dto.FileAttribute.CreatedAt = DateTime.Now;
        //        dto.FileAttribute.UpdatedAt = DateTime.Now;
        //        // 文件属性的 FileStorageId 需要在文件记录保存后才能关联，所以先保存属性获取 ID，再更新属性关联 FileStorageId
        //        int attributeResult = await _databaseManager.AddFileAttributeAsync(dto.FileAttribute);
        //        if (attributeResult <= 0)
        //        {
        //            LogManager.Instance.LogInfo("保存文件属性失败");
        //            // 发生异常，需要回滚操作
        //            await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, dto.FileStorage, dto.FileAttribute);
        //            return;
        //        }
        //        LogManager.Instance.LogInfo("保存文件属性到数据库:成功");

        //        // 获取属性ID并关联（尽量使用返回的 ID，如果 DatabaseManager 提供返回ID的重载建议改为使用）
        //        var dbAttribute = await _databaseManager.GetFileAttributeAsync(dto.FileStorage.DisplayName);
        //        if (dbAttribute == null || dbAttribute.Id == null)
        //        {
        //            LogManager.Instance.LogInfo("获取文件属性ID失败");
        //            await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, dto.FileStorage, dto.FileAttribute);
        //            return;
        //        }
        //        dto.FileAttribute.Id = dbAttribute.Id;
        //        dto.FileStorage.FileAttributeId = dbAttribute.Id;

        //        // 4. 保存文件记录到数据库
        //        var fileResult = await _databaseManager.AddFileStorageAsync(dto.FileStorage);
        //        if (fileResult == 0)
        //        {
        //            LogManager.Instance.LogInfo("保存文件记录到数据库:失败");
        //            await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, dto.FileStorage, dto.FileAttribute);
        //            return;
        //        }
        //        LogManager.Instance.LogInfo("保存文件记录到数据库:成功");

        //        // 5. 更新属性的FileStorageId (通过哈希查找有风险，但保持现有逻辑)
        //        var dbStorage = await _databaseManager.GetFileStorageAsync(dto.FileStorage.FileHash);
        //        if (dbStorage == null || dbStorage.Id == 0)
        //        {
        //            LogManager.Instance.LogInfo("获取存储记录失败");
        //            await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, dto.FileStorage, dto.FileAttribute);
        //            return;
        //        }
        //        dto.FileAttribute.FileStorageId = dbStorage.Id;
        //        await _databaseManager.UpdateFileAttributeAsync(dto.FileAttribute);

        //        // 6. 更新分类统计
        //        await _databaseManager.UpdateCategoryStatisticsAsync(dto.FileStorage.CategoryId, dto.FileStorage.CategoryType);

        //        transactionSuccess = true;

        //        // 7. 刷新界面（确保在 UI 线程）
        //        await Dispatcher.InvokeAsync(async () =>
        //        {
        //            try
        //            {
        //                await RefreshCurrentCategoryDisplayAsync(_selectedCategoryNode);
        //            }
        //            catch (Exception ex)
        //            {
        //                LogManager.Instance.LogError($"刷新分类显示失败: {ex.Message}");
        //            }
        //        });

        //        MessageBox.Show($"文件已成功上传并保存到服务器指定路径\n文件路径: {dto.FileStorage.FilePath}",
        //            "成功", MessageBoxButton.OK, MessageBoxImage.Information);
        //    }
        //    catch (Exception ex)
        //    {
        //        await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, dto.FileStorage, dto.FileAttribute);
        //        LogManager.Instance.LogError($"UploadFileAndSaveToDatabase 错误: {ex.Message}");
        //        throw;
        //    }
        //    finally
        //    {
        //        if (!transactionSuccess)
        //        {
        //            await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, dto.FileStorage, dto.FileAttribute);
        //        }
        //        _uploadSemaphore.Release();
        //    }
        //}

        /// <summary>
        /// 在 PropertiesDataGrid 显示完毕后调用，保存一次快照用于后续比较
        /// 参数 displayRows 通常来自 PrepareFileDisplayData(...) 返回的 List<CategoryPropertyEditModel>
        /// </summary>
        private void CapturePropertiesSnapshot(List<CategoryPropertyEditModel> displayRows)
        {
            try
            {
                _propertiesSnapshotForInsert.Clear();
                if (displayRows == null) return;

                foreach (var r in displayRows)
                {
                    if (!string.IsNullOrWhiteSpace(r.PropertyName1))
                    {
                        var key = NormalizePropertyDisplayName(r.PropertyName1);
                        if (!_propertiesSnapshotForInsert.ContainsKey(key))
                            _propertiesSnapshotForInsert[key] = r.PropertyValue1 ?? string.Empty;
                    }
                    if (!string.IsNullOrWhiteSpace(r.PropertyName2))
                    {
                        var key = NormalizePropertyDisplayName(r.PropertyName2);
                        if (!_propertiesSnapshotForInsert.ContainsKey(key))
                            _propertiesSnapshotForInsert[key] = r.PropertyValue2 ?? string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"CapturePropertiesSnapshot 失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 从当前 PropertiesDataGrid 读取键/值（用于比较或作为插入属性）
        /// </summary>
        private Dictionary<string, string> CollectCurrentPropertiesFromGrid()
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var items = PropertiesDataGrid?.ItemsSource as List<CategoryPropertyEditModel>;
                if (items == null) return result;

                foreach (var r in items)
                {
                    if (!string.IsNullOrWhiteSpace(r.PropertyName1))
                    {
                        var key = NormalizePropertyDisplayName(r.PropertyName1);
                        result[key] = r.PropertyValue1 ?? string.Empty;
                    }
                    if (!string.IsNullOrWhiteSpace(r.PropertyName2))
                    {
                        var key = NormalizePropertyDisplayName(r.PropertyName2);
                        result[key] = r.PropertyValue2 ?? string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"CollectCurrentPropertiesFromGrid 失败: {ex.Message}");
            }
            return result;
        }

        /// <summary>
        /// 比较快照与当前值，返回新增或修改（只要不同就返回）的键/值集合
        /// </summary>
        private Dictionary<string, string> GetModifiedPropertiesForInsert()
        {
            var current = CollectCurrentPropertiesFromGrid();
            var modified = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var kv in current)
            {
                if (!_propertiesSnapshotForInsert.TryGetValue(kv.Key, out var oldVal) || !string.Equals(oldVal ?? string.Empty, kv.Value ?? string.Empty, StringComparison.Ordinal))
                {
                    // 新增或修改（包括 snapshot 中没有的）
                    modified[kv.Key] = kv.Value ?? string.Empty;
                }
            }
            return modified;
        }

        /// <summary>
        /// 简单规范化 DisplayName -> 内部 key（用于比较与作为属性 Tag 的候选）
        /// - 去除前后空格，折叠多个空格
        /// - 保留原样以便对照 _propertyDisplayNameMap 反向查找（UI 显示名）
        /// 你可以把它拓展为更复杂的 Tag 生成规则（英文、去中文、截断等）
        private string NormalizePropertyDisplayName(string displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName)) return string.Empty;
            var s = Regex.Replace(displayName.Trim(), @"\s+", " ");
            return s;
        }

        /// <summary>
        /// 插入块后调用：把 modifiedProperties 应用到刚插入的 BlockReference
        /// 说明：这里仅提供方法签名与日志/步骤说明；具体 AutoCAD 事务内实现需要使用 Insert/Clone BlockTableRecord / AttributeDefinition/AttributeReference API。
        /// </summary>
        private void ApplyPropertiesToInsertedBlock(Autodesk.AutoCAD.DatabaseServices.ObjectId insertedBlockRefId, Dictionary<string, string> modifiedProperties)
        {
            try
            {
                if (modifiedProperties == null || modifiedProperties.Count == 0)
                {
                    LogManager.Instance.LogInfo("没有检测到修改属性，跳过属性应用。");
                    return;
                }

                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    LogManager.Instance.LogWarning("ApplyPropertiesToInsertedBlock：未找到活动文档");
                    return;
                }

                // 如果没有传入 ObjectId，提示用户选择刚插入的块
                if (insertedBlockRefId == Autodesk.AutoCAD.DatabaseServices.ObjectId.Null)
                {
                    var ed = doc.Editor;
                    var peo = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\n请选择刚插入的块参照（单击该块）：");
                    peo.SetRejectMessage("\n请选择块参照（BlockReference）。");
                    peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.BlockReference), true);
                    var per = ed.GetEntity(peo);
                    if (per.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                    {
                        LogManager.Instance.LogInfo("未选择块参照，取消属性应用。");
                        return;
                    }
                    insertedBlockRefId = per.ObjectId;
                }

                LogManager.Instance.LogInfo($"准备把 {modifiedProperties.Count} 个修改属性应用到插入的块 {insertedBlockRefId}（方案 B：写入 ExtensionDictionary/XRecord，若存在 AttributeReference 同步更新）。");

                using (doc.LockDocument())
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var br = tr.GetObject(insertedBlockRefId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.BlockReference;
                    if (br == null)
                    {
                        LogManager.Instance.LogWarning("目标对象不是 BlockReference，取消处理。");
                        return;
                    }

                    // 1) 尝试更新已有的 AttributeReference（如果 Tag 匹配）
                    try
                    {
                        foreach (ObjectId attId in br.AttributeCollection)
                        {
                            try
                            {
                                var ar = tr.GetObject(attId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.AttributeReference;
                                if (ar == null) continue;

                                string arTag = (ar.Tag ?? string.Empty).Trim();
                                if (string.IsNullOrEmpty(arTag)) continue;

                                // 在 modifiedProperties 中查找匹配项（优先尝试 MapDisplayNameToTag, 然后直接键比较）
                                var match = modifiedProperties.FirstOrDefault(kv =>
                                    string.Equals(MapDisplayNameToTag(kv.Key), arTag, StringComparison.OrdinalIgnoreCase)
                                    || string.Equals(NormalizePropertyDisplayName(kv.Key), arTag, StringComparison.OrdinalIgnoreCase)
                                    || string.Equals(kv.Key, arTag, StringComparison.OrdinalIgnoreCase)
                                );

                                if (!string.IsNullOrEmpty(match.Key))
                                {
                                    ar.TextString = match.Value ?? string.Empty;
                                }
                            }
                            catch
                            {
                                // 单个属性更新失败不影响整体流程
                            }
                        }
                    }
                    catch (Exception exUpdate)
                    {
                        LogManager.Instance.LogWarning($"更新 AttributeReference 值时遇到问题: {exUpdate.Message}");
                    }

                    // 2) 写入 ExtensionDictionary 的 XRecord（以便保存所有修改的键值对，不改变块定义）
                    try
                    {
                        // 如果没有扩展字典，创建
                        if (br.ExtensionDictionary == Autodesk.AutoCAD.DatabaseServices.ObjectId.Null)
                        {
                            br.CreateExtensionDictionary();
                        }

                        var extDict = tr.GetObject(br.ExtensionDictionary, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.DBDictionary;
                        if (extDict == null)
                        {
                            LogManager.Instance.LogWarning("无法获取或创建 BlockReference 的 ExtensionDictionary，跳过 XRecord 写入。");
                        }
                        else
                        {
                            // 对于每个修改项，使用规范化的 tag 作为字典键（保证稳定），内容保存在 XRecord 的单个文本 TypedValue 中
                            foreach (var kv in modifiedProperties)
                            {
                                try
                                {
                                    string displayKey = kv.Key ?? string.Empty;
                                    string value = kv.Value ?? string.Empty;
                                    string tag = MapDisplayNameToTag(displayKey);
                                    if (string.IsNullOrEmpty(tag))
                                        tag = ToValidAttributeTag(displayKey);

                                    // 如果已经存在同名条目，删除后重建（覆盖）
                                    if (extDict.Contains(tag))
                                    {
                                        var existingId = extDict.GetAt(tag);
                                        try
                                        {
                                            var dbObj = tr.GetObject(existingId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                            dbObj.Erase(true);
                                        }
                                        catch { /* 忽略删除失败 */ }
                                        // Remove from dictionary (SetAt will replace, but to be safe remove)
                                        try { extDict.Remove(tag); } catch { }
                                    }

                                    var xrec = new Autodesk.AutoCAD.DatabaseServices.Xrecord();
                                    xrec.Data = new Autodesk.AutoCAD.DatabaseServices.ResultBuffer(
                                        new Autodesk.AutoCAD.DatabaseServices.TypedValue((int)Autodesk.AutoCAD.DatabaseServices.DxfCode.Text, value ?? string.Empty)
                                    );

                                    extDict.SetAt(tag, xrec);
                                    tr.AddNewlyCreatedDBObject(xrec, true);
                                }
                                catch (Exception exX)
                                {
                                    LogManager.Instance.LogWarning($"为键 '{kv.Key}' 写入 XRecord 时失败: {exX.Message}");
                                }
                            }
                        }
                    }
                    catch (Exception exExt)
                    {
                        LogManager.Instance.LogWarning($"写入 ExtensionDictionary/XRecord 时出错: {exExt.Message}");
                    }

                    tr.Commit();
                    LogManager.Instance.LogInfo("ApplyPropertiesToInsertedBlock：属性已写入（XRecord）并尝试更新 AttributeReference。");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"ApplyPropertiesToInsertedBlock 失败: {ex.Message}");
                MessageBox.Show($"应用图元属性失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 把 UI 显示名映射为合适的 Attribute Tag（优先反查 _propertyDisplayNameMap）
        /// </summary>
        private string MapDisplayNameToTag(string displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName)) return string.Empty;

            try
            {
                // 反向查找：显示名 -> 属性名（例如 中文显示名 -> PropertyName）
                var kv = DictionaryHelper._propertyDisplayNameMap.FirstOrDefault(p => string.Equals(p.Value, displayName, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(kv.Key))
                {
                    return ToValidAttributeTag(kv.Key);
                }
            }
            catch { /* 忽略反查异常 */ }

            // 兜底：直接规范化 displayName（去空格、保留字母数字、转大写）
            return ToValidAttributeTag(displayName);
        }

        /// <summary>
        /// 生成合法 Attribute Tag（保守策略）
        /// - 仅保留字母数字和下划线，替换空格与常见符号为下划线
        /// - 转为大写并截断到 31 字符以内以兼容 AutoCAD Tag 限制
        /// </summary>
        private string ToValidAttributeTag(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            var sb = new System.Text.StringBuilder();
            foreach (var ch in raw)
            {
                if (char.IsLetterOrDigit(ch) || ch == '_')
                    sb.Append(ch);
                else if (char.IsWhiteSpace(ch) || ch == '-' || ch == '.' || ch == ':' || ch == '/')
                    sb.Append('_');
                // 其他字符忽略
            }
            var t = sb.ToString().Trim('_');
            if (t.Length == 0) t = "ATTR";
            if (t.Length > 31) t = t.Substring(0, 31);
            return t.ToUpperInvariant();
        }

        /// <summary>
        /// 在 WpfMainWindow 类中添加并发保护字段（类顶部私有字段区）
        /// </summary>
        private readonly System.Threading.SemaphoreSlim _uploadSemaphore = new System.Threading.SemaphoreSlim(1, 1);


        #endregion

        #region 文件处理

        /// <summary>
        /// 辅助方法
        /// </summary>
        private void InitializeFileUploadInterface()
        {
            // 清空所有输入框
            file_Path.Text = "";
            File_Name.Text = "";
            File_Size.Text = "";
            view_File_Path.Text = "";
            ViewImage.Source = null;

            // 清空属性编辑网格
            CategoryPropertiesDataGrid.ItemsSource = null;

            // 重置字段
            _selectedFilePath = null;
            _selectedPreviewImagePath = null;
            _currentFileStorage = null;
            _currentFileAttribute = null;
        }

        /// <summary>
        /// 设置文件属性
        /// </summary>
        /// <param name="attribute"></param>
        /// <param name="propertyName"></param>
        /// <param name="propertyValue"></param>
        /// <returns> </returns>

        public void SetFileAttributeProperty(FileAttribute attribute, string propertyName, string propertyValue)
        {
            if (string.IsNullOrEmpty(propertyName) || string.IsNullOrEmpty(propertyValue))
                return;
            bool boolValue = false;
            _currentFileStorage = _currentFileStorage ?? new FileStorage();
            if (propertyValue == "是") boolValue = true;
            try
            {
                switch (propertyName.ToLower())
                {
                    case "长度":
                        if (decimal.TryParse(propertyValue, out decimal length))
                            attribute.Length = length;
                        break;
                    case "宽度":
                        if (decimal.TryParse(propertyValue, out decimal width))
                            attribute.Width = width;
                        break;
                    case "高度":
                        if (decimal.TryParse(propertyValue, out decimal height))
                            attribute.Height = height;
                        break;
                    case "角度":
                        if (decimal.TryParse(propertyValue, out decimal angle))
                            attribute.Angle = angle;
                        break;
                    case "基点x":
                        if (decimal.TryParse(propertyValue, out decimal baseX))
                            attribute.BasePointX = baseX;
                        break;
                    case "基点y":
                        if (decimal.TryParse(propertyValue, out decimal baseY))
                            attribute.BasePointY = baseY;
                        break;
                    case "基点z":
                        if (decimal.TryParse(propertyValue, out decimal baseZ))
                            attribute.BasePointZ = baseZ;
                        break;
                    case "介质":
                        attribute.MediumName = propertyValue;
                        break;
                    case "规格":
                        attribute.Specifications = propertyValue;
                        break;
                    case "材质":
                        attribute.Material = propertyValue;
                        break;
                    case "标准号":
                        attribute.StandardNumber = propertyValue;
                        break;
                    case "功率":
                        attribute.Power = propertyValue;
                        break;
                    case "容积":
                        attribute.Volume = propertyValue;
                        break;
                    case "压力":
                        attribute.Pressure = propertyValue;
                        break;
                    case "温度":
                        attribute.Temperature = propertyValue;
                        break;
                    case "直径":
                        attribute.Diameter = propertyValue;
                        break;
                    case "外径":
                        attribute.OuterDiameter = propertyValue;
                        break;
                    case "内径":
                        attribute.InnerDiameter = propertyValue;
                        break;
                    case "厚度":
                        attribute.Thickness = propertyValue;
                        break;
                    case "重量":
                        attribute.Weight = propertyValue;
                        break;
                    case "型号":
                        attribute.Model = propertyValue;
                        break;
                    case "备注":
                        attribute.Remarks = propertyValue;
                        break;
                    case "名称":
                        attribute.FileName = propertyValue;
                        break;
                    case "元素块名":
                        _currentFileStorage.BlockName = propertyValue;
                        break;
                    case "层名":
                        _currentFileStorage.LayerName = propertyValue;
                        break;
                    case "颜色索引":
                        _currentFileStorage.ColorIndex = Convert.ToInt32(propertyValue);
                        break;
                    case "是否公开":
                        _currentFileStorage.IsPublic = Convert.ToInt32(propertyValue);
                        break;
                    case "描述":
                        _currentFileStorage.Description = propertyValue;
                        break;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"设置属性 {propertyName} 时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置文件存储属性
        /// </summary>
        /// <param name="storage"></param>
        /// <param name="propertyName"></param>
        /// <param name="propertyValue"></param>
        public void SetFileStorageProperty(FileStorage storage, string propertyName, string propertyValue)
        {
            if (string.IsNullOrEmpty(propertyName) || string.IsNullOrEmpty(propertyValue) || storage == null)
                return;
            try
            {
                switch (propertyName.Trim().ToLower())
                {
                    case "文件名":
                    case "名称":
                        storage.FileName = propertyValue;
                        break;
                    case "显示名称":
                        storage.DisplayName = propertyValue;
                        break;
                    case "文件路径":
                        storage.FilePath = propertyValue;
                        break;
                    case "文件类型":
                        storage.FileType = propertyValue;
                        break;
                    case "文件大小":
                        if (long.TryParse(propertyValue, out long size))
                            storage.FileSize = size;
                        break;
                    case "元素块名":
                        storage.BlockName = propertyValue;
                        break;
                    case "图层名称":
                    case "层名":
                        storage.LayerName = propertyValue;
                        break;
                    case "颜色索引":
                        if (int.TryParse(propertyValue, out int colorIdx))
                            storage.ColorIndex = colorIdx;
                        break;
                    case "预览图片名称":
                        storage.PreviewImageName = propertyValue;
                        break;
                    case "预览图片路径":
                        storage.PreviewImagePath = propertyValue;
                        break;
                    case "是否预览":
                        storage.IsPreview = propertyValue == "是" ? 1 : 0;
                        break;
                    case "创建者":
                        storage.CreatedBy = propertyValue;
                        break;
                    case "标题":
                        storage.Title = propertyValue;
                        break;
                    case "关键字":
                        storage.Keywords = propertyValue;
                        break;
                    case "更新者":
                        storage.UpdatedBy = propertyValue;
                        break;
                    case "版本号":
                        if (int.TryParse(propertyValue, out int ver))
                            storage.Version = ver;
                        break;
                    case "是否激活":
                        storage.IsActive = propertyValue == "是" ? 1 : 0;
                        break;
                    case "是否公开":
                        storage.IsPublic = propertyValue == "是" ? 1 : 0;
                        break;
                    case "描述":
                        storage.Description = propertyValue;
                        break;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"设置文件信息属性 {propertyName} 时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加文件选择相关字段
        /// </summary>
        private FileStorage _selectedFileStorage;

        /// <summary>
        /// DataGrid选中事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void StroageFileDataGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            try
            {
                if (StroageFileDataGrid.SelectedItem is FileStorage selectedFile)
                {
                    LogManager.Instance.LogInfo($"选中文件: {selectedFile.DisplayName} (ID: {selectedFile.Id})");
                    _selectedFileStorage = selectedFile;

                    // 显示文件基本信息
                    DisplayFileBasicInfo(selectedFile);
                    // 加载并显示文件属性
                    await LoadAndDisplayFileAttributesAsync(selectedFile);

                    // 加载预览图片
                    var previewImage = await GetPreviewImageAsync(selectedFile);
                    // 预览图片会在PreviewImage_Loaded事件中处理
                    // 初始化属性编辑界面
                    //InitializeFilePropertiesGrid();
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"处理文件选择时出错: {ex.Message}");
                MessageBox.Show($"处理文件选择时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 显示文件基本信息
        /// </summary>
        /// <param name="fileStorage"></param>
        private void DisplayFileBasicInfo(FileStorage fileStorage)
        {
            try
            {
                file_Path.Text = fileStorage.FilePath ?? "";
                File_Name.Text = fileStorage.DisplayName ?? fileStorage.FileName ?? "";
                File_Size.Text = fileStorage.FileSize > 0 ? $"{fileStorage.FileSize / 1024.0:F2} KB" : "";
                File_Type.Text = fileStorage.FileType ?? "";
                view_File_Path.Text = fileStorage.PreviewImagePath ?? "无预览图片";
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"显示文件基本信息时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 加载并显示文件属性（修改现有方法）
        /// </summary>
        /// <param name="fileStorage">数据库储存的文件</param>
        /// <returns></returns>
        private async Task LoadAndDisplayFileAttributesAsync(FileStorage fileStorage)
        {
            try
            {
                LogManager.Instance.LogInfo($"开始加载文件 {fileStorage.DisplayName} 的属性");

                if (_databaseManager == null)
                {
                    LogManager.Instance.LogWarning("数据库管理器为空");
                    return;
                }

                // 获取文件属性
                var fileAttribute = await _databaseManager.GetFileAttributeByGraphicIdAsync(fileStorage.Id);
                _selectedFileAttribute = fileAttribute;

                // 准备显示数据
                var displayData = PrepareFileDisplayData(fileStorage, fileAttribute);
                CategoryPropertiesDataGrid.ItemsSource = displayData;

                LogManager.Instance.LogInfo("文件属性加载完成");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"加载文件属性时出错: {ex.Message}");
                MessageBox.Show($"加载文件属性时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 准备文件显示数据
        /// </summary>
        /// <param name="fileStorage"></param>
        /// <param name="fileAttribute"></param>
        /// <returns></returns>
        public List<CategoryPropertyEditModel> PrepareFileDisplayData(FileStorage fileStorage, FileAttribute fileAttribute)
        {

            var propertyRows = new List<CategoryPropertyEditModel>();

            try
            {
                LogManager.Instance.LogDebug("准备文件显示数据（限制 FileStorage 字段）");

                // 使用有序集合按显示顺序收集要显示的属性键/值
                var allProperties = new List<KeyValuePair<string, string>>();

                // 1) 仅加入 cad_file_storage 指定的字段（按要求：file_name, display_name, element_block_name, layer_name, color_index, scale）
                if (fileStorage != null)
                {
                    // 添加文件名称和显示名称（新增）
                    string fileName = fileStorage.FileName ?? string.Empty;
                    string displayName = fileStorage.DisplayName ?? string.Empty;

                    // 注意：Category 前缀与 AddObjectProperties 保持一致，便于 GetPropertyDisplayName 做映射
                    string eb = fileStorage.BlockName ?? string.Empty;
                    string ln = fileStorage.LayerName ?? string.Empty;
                    string ci = fileStorage.ColorIndex.HasValue ? fileStorage.ColorIndex.Value.ToString() : (fileStorage.ColorIndex != null ? fileStorage.ColorIndex.ToString() : string.Empty);

                    // 使用兼容读取器获取比例（兼容 Scale/scale 字段以及多种数值类型）
                    string sc = string.Empty;
                    try
                    {
                        double scaleVal = GetScaleFromFileStorage(fileStorage);
                        sc = double.IsNaN(scaleVal) ? string.Empty : scaleVal.ToString("G");
                    }
                    catch
                    {
                        sc = string.Empty;
                    }

                    // 按顺序添加属性：文件名称、显示名称、块名、层名、颜色索引、比例
                    allProperties.Add(new KeyValuePair<string, string>("文件信息.FileName", fileName));
                    allProperties.Add(new KeyValuePair<string, string>("文件信息.DisplayName", displayName));
                    allProperties.Add(new KeyValuePair<string, string>("文件信息.ElementBlockName", eb));
                    allProperties.Add(new KeyValuePair<string, string>("文件信息.LayerName", ln));
                    allProperties.Add(new KeyValuePair<string, string>("文件信息.ColorIndex", ci));
                    allProperties.Add(new KeyValuePair<string, string>("文件信息.Scale", sc));
                }

                // 2) 追加 cad_file_attributes 中的所有可展示属性（使用现有方法统一处理命名与格式化）
                if (fileAttribute != null)
                {
                    AddObjectProperties(allProperties, fileAttribute, "属性信息");
                }

                // 3) 把列表转换为两列显示格式（与以前行为一致）
                for (int i = 0; i < allProperties.Count; i += 2)
                {
                    var row = new CategoryPropertyEditModel();

                    // 第一列
                    var prop1 = allProperties[i];
                    row.PropertyName1 = GetPropertyDisplayName(prop1.Key);
                    row.PropertyValue1 = prop1.Value ?? string.Empty;

                    // 第二列（如果有）
                    if (i + 1 < allProperties.Count)
                    {
                        var prop2 = allProperties[i + 1];
                        row.PropertyName2 = GetPropertyDisplayName(prop2.Key);
                        row.PropertyValue2 = prop2.Value ?? string.Empty;
                    }

                    propertyRows.Add(row);
                }

                // 4) 保证至少有若干空行用于编辑
                while (propertyRows.Count < 5)
                {
                    propertyRows.Add(new CategoryPropertyEditModel());
                }

                LogManager.Instance.LogDebug($"准备完成 {propertyRows.Count} 行属性数据（仅限指定 FileStorage 字段 + FileAttribute）");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"准备文件显示数据时出错: {ex.Message}");
            }

            return propertyRows;

            #region 添加属性属性显示名称映射
            //var propertyRows = new List<CategoryPropertyEditModel>();

            //try
            //{
            //    LogManager.Instance.LogDebug("准备文件显示数据");

            //    // 收集所有属性
            //    var allProperties = new List<KeyValuePair<string, string>>();

            //    // 添加FileStorage属性
            //    if (fileStorage != null)
            //    {
            //        AddObjectProperties(allProperties, fileStorage, "文件信息");
            //    }

            //    // 添加FileAttribute属性
            //    if (fileAttribute != null)
            //    {
            //        AddObjectProperties(allProperties, fileAttribute, "属性信息");
            //    }

            //    // 转换为两列显示格式
            //    for (int i = 0; i < allProperties.Count; i += 2)
            //    {
            //        var row = new CategoryPropertyEditModel();

            //        // 第一列
            //        var prop1 = allProperties[i];
            //        row.PropertyName1 = GetPropertyDisplayName(prop1.Key);
            //        row.PropertyValue1 = prop1.Value ?? "";

            //        // 第二列（如果有）
            //        if (i + 1 < allProperties.Count)
            //        {
            //            var prop2 = allProperties[i + 1];
            //            row.PropertyName2 = GetPropertyDisplayName(prop2.Key);
            //            row.PropertyValue2 = prop2.Value ?? "";
            //        }

            //        propertyRows.Add(row);
            //    }

            //    // 确保至少有几行空行用于编辑
            //    while (propertyRows.Count < 5)
            //    {
            //        propertyRows.Add(new CategoryPropertyEditModel());
            //    }

            //    LogManager.Instance.LogDebug($"准备完成 {propertyRows.Count} 行属性数据");
            //}
            //catch (Exception ex)
            //{
            //    LogManager.Instance.LogError($"准备文件显示数据时出错: {ex.Message}");
            //}

            //return propertyRows;

            #endregion

        }

        /// <summary>
        /// 添加对象属性到列表
        /// </summary>
        /// <param name="properties"></param>
        /// <param name="obj"></param>
        /// <param name="category"></param>
        private void AddObjectProperties(List<KeyValuePair<string, string>> properties, object obj, string category)
        {
            try
            {
                if (obj == null) return;

                var objectType = obj.GetType();
                var objectProperties = objectType.GetProperties();

                foreach (var prop in objectProperties)
                {
                    try
                    {
                        // 跳过一些不需要显示的属性
                        if (ShouldSkipProperty(prop.Name))
                            continue;

                        var value = prop.GetValue(obj);
                        string displayValue = value?.ToString() ?? "";

                        // 特殊处理某些属性
                        if (prop.Name == "FileSize" && value is long fileSize)
                        {
                            displayValue = FileManager.FormatFileSize(fileSize);
                        }
                        else if (prop.Name.EndsWith("At") && value is DateTime dateTime)
                        {
                            displayValue = dateTime.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        else if (prop.Name == "IsActive" || prop.Name == "IsPublic" || prop.Name == "IsPreview")
                        {
                            displayValue = (value?.ToString() == "True") ? "是" : "否";
                        }

                        properties.Add(new KeyValuePair<string, string>($"{category}.{prop.Name}", displayValue));
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogDebug($"获取属性 {prop.Name} 值时出错: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"添加对象属性时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 判断是否应该跳过属性
        /// </summary>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        private bool ShouldSkipProperty(string propertyName)
        {
            if (string.IsNullOrEmpty(propertyName)) return true;

            // 不区分大小写的跳过集合：
            // - 二进制/大对象字段
            // - cad_file_attributes 中不想显示的字段：Id, CategoryId, FileStorageId, FileName
            var skipProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "FileData",
                "PreviewImageData",
                "Id",
                "CategoryId",
                "FileStorageId",
                "FileName",
                "BasePointX",
                "BasePointY",
                "BasePointZ",
                "CreatedAt",
                "UpdatedAt"
            };

            return skipProperties.Contains(propertyName);
        }

        /// <summary>
        /// 获取属性显示名称
        /// </summary>
        /// <param name="fullPropertyName"></param>
        /// <returns></returns>
        private string GetPropertyDisplayName(string fullPropertyName)
        {
            try
            {
                // 分离分类和属性名
                if (fullPropertyName.Contains("."))
                {
                    var parts = fullPropertyName.Split('.');
                    var category = parts[0];
                    var propertyName = parts[1];

                    // 获取映射名称
                    if (DictionaryHelper._propertyDisplayNameMap.TryGetValue(propertyName, out string displayName))
                    {
                        return displayName;
                    }
                }
                else
                {
                    // 直接属性名
                    if (DictionaryHelper._propertyDisplayNameMap.TryGetValue(fullPropertyName, out string displayName))
                    {
                        return displayName;
                    }
                }

                // 如果没有映射，返回原始名称
                return fullPropertyName.Contains(".") ? fullPropertyName.Split('.')[1] : fullPropertyName;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogDebug($"获取属性显示名称时出错: {ex.Message}");
                return fullPropertyName;
            }
        }

        /// <summary>
        /// 清空DataGrid中的文件属性显示
        /// </summary>
        private void ClearFilePropertiesInDataGrid()
        {
            try
            {
                if (PropertiesDataGrid != null)
                {
                    PropertiesDataGrid.ItemsSource = null;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"清空PropertiesDataGrid时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 刷新当前分类的显示
        /// </summary>
        /// <param name="categoryNode">分类节点</param>
        public async Task RefreshCurrentCategoryDisplayAsync(CategoryTreeNode categoryNode)
        {
            try
            {
                // 根据当前选中的分类节点，刷新对应的界面显示
                if (categoryNode.Level == 0 && categoryNode.Data is CadCategory)
                {
                    // 主分类，刷新主分类下的内容显示
                    await RefreshMainCategoryDisplayAsync(categoryNode);
                }
                else if (categoryNode.Data is CadSubcategory)
                {
                    // 子分类，刷新子分类下的内容显示
                    await RefreshSubcategoryDisplayAsync(categoryNode);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"刷新当前分类显示时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 刷新主分类显示
        /// </summary>
        /// <param name="categoryNode">主分类节点</param>
        private async Task RefreshMainCategoryDisplayAsync(CategoryTreeNode categoryNode)
        {
            try
            {
                // 这里可以根据需要刷新主分类的显示
                // 例如：刷新主分类下的文件列表等
                LogManager.Instance.LogInfo($"刷新主分类显示: {categoryNode.DisplayText}");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"刷新主分类显示时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 刷新子分类显示
        /// </summary>
        /// <param name="categoryNode">子分类节点</param>
        private async Task RefreshSubcategoryDisplayAsync(CategoryTreeNode categoryNode)
        {
            try
            {
                // 刷新子分类下的文件显示
                await RefreshSubcategoryFilesDisplayAsync(categoryNode);
                LogManager.Instance.LogInfo($"刷新子分类显示: {categoryNode.DisplayText}");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"刷新子分类显示时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 刷新子分类文件显示
        /// </summary>
        /// <param name="subcategoryNode">子分类节点</param>
        private async Task RefreshSubcategoryFilesDisplayAsync(CategoryTreeNode subcategoryNode)
        {
            try
            {
                // 根据当前TabItem刷新对应的文件显示
                // 这里需要根据当前选中的TabItem来确定刷新哪个面板
                string currentTabHeader = GetCurrentSelectedTabHeader();

                if (!string.IsNullOrEmpty(currentTabHeader))
                {
                    // 找到对应的面板并刷新
                    WrapPanel targetPanel = GetPanelByFolderName(currentTabHeader);
                    if (targetPanel != null)
                    {
                        // 重新加载该分类下的按钮
                        await LoadButtonsFromDatabase(currentTabHeader, targetPanel);
                        LogManager.Instance.LogInfo($"刷新了 {currentTabHeader} 面板的文件显示");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"刷新子分类文件显示时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取当前选中的TabItem标题  _selectedCategoryNode  RefreshCategoryTreeAsync
        /// </summary>
        /// <returns>TabItem标题</returns>
        private string GetCurrentSelectedTabHeader()
        {
            try
            {
                // 根据当前选中的分类节点确定对应的TabItem
                // 这里需要根据您的具体界面结构来实现
                if (_selectedCategoryNode != null)
                {
                    // 可以根据分类节点的名称来确定对应的TabItem
                    // 例如：如果分类节点名称为"工艺"，则对应的TabItem为"工艺"
                    return _selectedCategoryNode.DisplayText;
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"获取当前选中TabItem标题时出错: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// 刷新右侧图元文件列表（根据当前选中的分类节点）
        /// </summary>
        private async Task RefreshFilesForCurrentCategoryAsync()
        {
            try
            {
                if (_selectedCategoryNode == null || _databaseManager == null || !_databaseManager.IsDatabaseAvailable)
                    return;

                var nodeData = _selectedCategoryNode.Data;
                List<FileStorage> files = null;

                if (nodeData is CadCategory main)
                {
                    files = await _databaseManager.GetFilesByCategoryIdAsync(main.Id, "main");
                }
                else if (nodeData is CadSubcategory sub)
                {
                    files = await _databaseManager.GetFilesByCategoryIdAsync(sub.Id, "sub");
                }

                if (files != null)
                {
                    StroageFileDataGrid.ItemsSource = files;
                    StroageFileDataGrid.Items.Refresh();
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"刷新文件列表失败: {ex.Message}");
            }
        }

        #endregion

        #region 电气按键
        private void 横墙电开建筑洞_Btn_Clic(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("横墙电开建洞");
            command?.Invoke();
        }

        private void 纵墙电开建筑洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 矩形电开结构洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 直径电开结构洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 半径电开结构洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }
        #endregion

        #region 暖通按键
        private void 横墙暖开建筑洞_Btn_Clic(object sender, RoutedEventArgs e)
        {

        }

        private void 纵墙暖开建筑洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 矩形暖开结构洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 直径暖开结构洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 半径暖开结构洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }
        #endregion

        #region 自控按键
        private void 横墙自控开建筑洞_Btn_Clic(object sender, RoutedEventArgs e)
        {

        }

        private void 纵墙自控开建筑洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 矩形自控开结构洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 直径自控开结构洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 半径自控开结构洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }
        #endregion

        #region 建筑按键
        private void 吊顶_Btn_Clic(object sender, RoutedEventArgs e)
        {
            VariableDictionary.winForm_Status = false;
            var command = UnifiedCommandManager.GetCommand("吊顶");
            command?.Invoke();
        }

        private void 不吊顶_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("不吊顶");
            command?.Invoke();
        }

        private void 防撞护板_Btn_Clic(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("防撞护板");
            command?.Invoke();
        }

        private void 房间编号_Btn_Clic(object sender, RoutedEventArgs e)
        {
            VariableDictionary.winForm_Status = false;

            var command = UnifiedCommandManager.GetCommand("房间编号");
            command?.Invoke();
        }

        private void 编号检查_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("编号检查");
            command?.Invoke();
        }

        private void 冷藏库降板_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("冷藏库降板");
            command?.Invoke();
        }

        private void 冷冻库降板_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("冷冻库降板");
            command?.Invoke();
        }

        private void 特殊地面做法要求_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("特殊地面做法要求");
            command?.Invoke();
        }

        private void 排水沟_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("排水沟");
            command?.Invoke();
        }

        private void 横墙建筑开洞_Btn_Clic(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("横墙建筑开洞");
            command?.Invoke();
        }

        private void 纵墙建筑开洞_Btn_Click(object sender, RoutedEventArgs e)
        {
            var command = UnifiedCommandManager.GetCommand("纵墙建筑开洞");
            command?.Invoke();
        }
        #endregion

        #region 结构按键
        private void 结开建筑洞_Btn_Clic(object sender, RoutedEventArgs e)
        {

        }

        private void 纵墙结开建筑洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 矩形结开洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 直径结开洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 半径结开洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }
        #endregion

        #region 水按键
        private void 横墙水开建筑洞_Btn_Clic(object sender, RoutedEventArgs e)
        {

        }

        private void 纵墙水开建筑洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 矩形水开结构洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 直径水开结构洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 半径水开结构洞_Btn_Click(object sender, RoutedEventArgs e)
        {

        }
        #endregion

        #endregion

        private void 保存设置按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveSettings();
                MessageBox.Show("设置已保存", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void 测试连接按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取当前输入的设置
                string serverIP = TextBox_Set_ServiceIP.Text.Trim();
                string serverPort = TextBox_Set_ServicePort.Text.Trim();
                string databaseName = TextBox_Set_DatabaseName.Text.Trim();
                string username = TextBox_Set_Username.Text.Trim();
                string password = PasswordBox_Set_Password.Text.Trim();

                if (string.IsNullOrEmpty(serverIP) || string.IsNullOrEmpty(serverPort) ||
                    string.IsNullOrEmpty(databaseName) || string.IsNullOrEmpty(username) ||
                    string.IsNullOrEmpty(password))
                {
                    MessageBox.Show("请填写完整的数据库连接信息", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!int.TryParse(serverPort, out int port) || port <= 0 || port > 65535)
                {
                    MessageBox.Show("请输入有效的端口号（1-65535）", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 首先测试网络连接
                MessageBox.Show("正在测试网络连接...", "提示", MessageBoxButton.OK, MessageBoxImage.Information);

                try
                {
                    // 构建连接字符串
                    string testConnectionString = $"Server={serverIP};Port={port};Database={databaseName};Uid={username};Pwd={password};Connection Timeout=5;";

                    LogManager.Instance.LogInfo($"尝试连接到 {serverIP}:{port}");

                    // 测试连接
                    var testDatabaseManager = new DatabaseManager(testConnectionString);
                    if (testDatabaseManager.IsDatabaseAvailable)
                    {
                        MessageBox.Show($"数据库连接测试成功\n服务器: {serverIP}:{port}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }
                    else
                    {
                        MessageBox.Show($"数据库连接测试失败\n服务器: {serverIP}:{port}", "失败", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"数据库连接测试失败\n错误信息: {ex.Message}\n服务器: {serverIP}:{port}",
                        "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"测试连接失败: {ex.Message}\n{ex.StackTrace}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void 应用设置按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveSettings();
                ReinitializeDatabase();
                // 立即根据可能修改的用户名/权限更新 Tab 可见性
                UpdateAdminTabsVisibility();
                MessageBox.Show("设置已应用，数据库连接已更新", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用设置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #region 服务器设置方法

        /// <summary>
        /// 保存设置到配置文件
        /// </summary>
        private void SaveSettings()
        {
            try
            {
                // 更新字段值
                VariableDictionary._serverIP = TextBox_Set_ServiceIP.Text.Trim();
                VariableDictionary._serverPort = int.TryParse(TextBox_Set_ServicePort.Text.Trim(), out int port) ? port : 3306;
                VariableDictionary._dataBaseName = TextBox_Set_DatabaseName.Text.Trim();
                //VariableDictionary._userName = TextBox_Set_Username.Text.Trim();
                //VariableDictionary._passWord = PasswordBox_Set_Password.Text.Trim();
                VariableDictionary._storagePath = TextBox_Set_StoragePath.Text.Trim();
                VariableDictionary._useDPath = CheckBox_UseDPath.IsChecked ?? true;
                VariableDictionary._autoSync = CheckBox_AutoSync.IsChecked ?? true;
                VariableDictionary._syncInterval = int.TryParse(TextBox_SyncInterval.Text, out int interval) ? interval : 30;

                // 保存到配置文件
                Properties.Settings.Default.ServerIP = VariableDictionary._serverIP;
                Properties.Settings.Default.ServerPort = VariableDictionary._serverPort;
                Properties.Settings.Default.DatabaseName = VariableDictionary._dataBaseName;
                Properties.Settings.Default.Username = VariableDictionary._userName;
                Properties.Settings.Default.Password = VariableDictionary._passWord;
                Properties.Settings.Default.StoragePath = VariableDictionary._storagePath;
                Properties.Settings.Default.UseDPath = VariableDictionary._useDPath;
                Properties.Settings.Default.AutoSync = VariableDictionary._autoSync;
                Properties.Settings.Default.SyncInterval = VariableDictionary._syncInterval;
                Properties.Settings.Default.Save();

                LogManager.Instance.LogInfo("设置已保存到配置文件");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"保存设置时出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 检测输入的字符是否合法
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private static bool IsTextAllowed(string text)
        {
            return text.All(char.IsDigit);
        }

        /// <summary>
        /// 处理粘贴操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBox_Set_ServicePort_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(string)))
            {
                string text = (string)e.DataObject.GetData(typeof(string));
                if (!IsTextAllowed(text))
                {
                    e.CancelCommand();
                }
            }
            else
            {
                e.CancelCommand();
            }
        }

        /// <summary>
        /// 在管理员模块中添加查看日志按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 查看日志按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string logFilePath = LogManager.Instance.LogFilePath;
                if (File.Exists(logFilePath))
                {
                    // 使用默认程序打开日志文件
                    System.Diagnostics.Process.Start(logFilePath);
                }
                else
                {
                    MessageBox.Show("日志文件不存在", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Env.Editor.WriteMessage($"打开日志文件失败: {ex.Message}");
                MessageBox.Show($"打开日志文件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 或者添加一个显示最新日志的方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 显示最新日志按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string logFilePath = LogManager.Instance.LogFilePath;
                if (File.Exists(logFilePath))
                {
                    var lines = File.ReadAllLines(logFilePath);
                    // 显示最后100行日志
                    var recentLogs = lines.Skip(Math.Max(0, lines.Length - 100));

                    // 可以显示在弹出窗口或TextBox中
                    string logContent = string.Join(Environment.NewLine, recentLogs);

                    // 创建一个简单的日志查看窗口
                    var logWindow = new Window
                    {
                        Title = "应用程序日志",
                        Width = 800,
                        Height = 600,
                        WindowStartupLocation = WindowStartupLocation.CenterOwner
                    };

                    var textBox = new TextBox
                    {
                        Text = logContent,
                        IsReadOnly = true,
                        VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                        HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                        FontFamily = new FontFamily("Consolas"),
                        FontSize = 12
                    };

                    logWindow.Content = textBox;
                    logWindow.Show();
                }
            }
            catch (Exception ex)
            {
                Env.Editor.WriteMessage($"显示日志失败: {ex.Message}");
            }
        }

        #endregion

        #region 批量添加文件

        private void 导出模板_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogManager.Instance.LogInfo("开始导出模板");

                // 创建模板DataTable
                DataTable templateTable = CreateTemplateDataTable();

                // 选择保存路径
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Title = "保存模板文件",
                    Filter = "Excel文件 (*.xlsx)|*.xlsx|Excel 97-2003文件 (*.xls)|*.xls",
                    FileName = "图元批量添加模板.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;

                    // 导出到Excel
                    if (ExportDataTableToExcel(templateTable, filePath))
                    {
                        MessageBox.Show($"模板已成功导出到:\n{filePath}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                        LogManager.Instance.LogInfo($"模板导出成功: {filePath}");
                    }
                    else
                    {
                        MessageBox.Show("模板导出失败", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        LogManager.Instance.LogError("模板导出失败");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"导出模板时出错: {ex.Message}");
                MessageBox.Show($"导出模板时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void 导入模板_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "选择Excel文件",
                    Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls",
                    Multiselect = false
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    file_Path.Text = openFileDialog.FileName;
                    LogManager.Instance.LogInfo($"选择Excel文件: {openFileDialog.FileName}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"选择Excel文件时出错: {ex.Message}");
                MessageBox.Show($"选择Excel文件时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void 批量添加图元_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(file_Path.Text))
                {
                    MessageBox.Show("请先选择Excel文件", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!File.Exists(file_Path.Text))
                {
                    MessageBox.Show("选择的Excel文件不存在", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 显示确认对话框
                var result = MessageBox.Show("确定要批量添加图元吗？这将导入Excel文件中的所有数据。",
                    "确认", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.No)
                    return;

                // 开始批量导入
                _fileManager.BatchImportGraphicsAsync(file_Path.Text, _selectedCategoryNode, CategoryPropertiesDataGrid, _categoryTreeNodes);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"批量添加图元时出错: {ex.Message}");
                MessageBox.Show($"批量添加图元时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 创建模板DataTable
        /// </summary>
        public DataTable CreateTemplateDataTable()
        {
            DataTable dt = new DataTable("图元批量添加模板");
            #region FileStorage
            // 添加列（基于FileStorage和FileAttribute的所有字段）
            dt.Columns.Add("分类ID", typeof(int));
            dt.Columns.Add("分类类型", typeof(string));
            dt.Columns.Add("文件名", typeof(string));
            dt.Columns.Add("显示名称", typeof(string));
            dt.Columns.Add("文件路径", typeof(string));
            dt.Columns.Add("文件类型", typeof(string));
            dt.Columns.Add("文件大小", typeof(long));
            dt.Columns.Add("元素块名", typeof(string));
            dt.Columns.Add("图层名称", typeof(string));
            dt.Columns.Add("颜色索引", typeof(int));
            dt.Columns.Add("预览图片名称", typeof(string));
            dt.Columns.Add("预览图片路径", typeof(string));
            dt.Columns.Add("是否预览", typeof(int));
            dt.Columns.Add("创建者", typeof(string));
            dt.Columns.Add("标题", typeof(string));
            dt.Columns.Add("关键字", typeof(string));
            dt.Columns.Add("更新者", typeof(string));
            dt.Columns.Add("版本号", typeof(int));
            dt.Columns.Add("是否激活", typeof(int));
            dt.Columns.Add("是否公开", typeof(int));
            dt.Columns.Add("描述", typeof(string));
            #endregion
            #region FileAttribute
            // FileAttribute字段
            dt.Columns.Add("存储文件ID", typeof(string));
            dt.Columns.Add("文件名称", typeof(string));
            dt.Columns.Add("长度", typeof(double));
            dt.Columns.Add("宽度", typeof(double));
            dt.Columns.Add("高度", typeof(double));
            dt.Columns.Add("角度", typeof(string));
            dt.Columns.Add("介质", typeof(string));
            dt.Columns.Add("材质", typeof(string));
            dt.Columns.Add("规格", typeof(string));
            dt.Columns.Add("标准号", typeof(string));
            dt.Columns.Add("功率", typeof(string));
            dt.Columns.Add("容积", typeof(string));
            dt.Columns.Add("压力", typeof(string));
            dt.Columns.Add("温度", typeof(string));
            dt.Columns.Add("直径", typeof(string));
            dt.Columns.Add("外径", typeof(string));
            dt.Columns.Add("内径", typeof(string));
            dt.Columns.Add("厚度", typeof(string));
            dt.Columns.Add("重量", typeof(string));
            dt.Columns.Add("型号", typeof(string));
            dt.Columns.Add("备注", typeof(string));
            dt.Columns.Add("自定义1", typeof(string));
            dt.Columns.Add("自定义2", typeof(string));
            dt.Columns.Add("自定义3", typeof(string));
            #endregion
            #region FileStorage 示例
            // 添加示例行
            DataRow sampleRow = dt.NewRow();
            sampleRow["分类ID"] = 1;
            sampleRow["分类类型"] = "sub";
            sampleRow["文件名"] = "示例文件.dwg";
            sampleRow["显示名称"] = "示例图元";
            sampleRow["文件路径"] = "C:\\示例路径\\示例文件.dwg";
            sampleRow["文件类型"] = ".dwg";
            sampleRow["文件大小"] = 102400;
            sampleRow["元素块名"] = "220V插座";
            sampleRow["图层名称"] = "TJ(电气专业D)";
            sampleRow["颜色索引"] = "142";
            sampleRow["预览图片名称"] = "示例图片.png";
            sampleRow["预览图片路径"] = "C:\\示例路径\\示例文件.png";
            sampleRow["是否预览"] = 0;
            sampleRow["创建者"] = "张三";
            sampleRow["标题"] = "220V电源插座";
            sampleRow["描述"] = "220V电源插座";
            sampleRow["关键字"] = "220V、电源插座";
            sampleRow["版本号"] = 1;
            sampleRow["是否激活"] = 1;
            sampleRow["是否公开"] = 1;
            #endregion
            #region FileAttribute示例数据
            sampleRow["长度"] = 200.0;
            sampleRow["宽度"] = 100.0;
            sampleRow["高度"] = 50.0;
            sampleRow["角度"] = 90.0;
            sampleRow["介质"] = "水";
            sampleRow["材质"] = "316不锈钢";
            sampleRow["规格"] = "Standard";
            sampleRow["标准编号"] = "2.5";
            sampleRow["功率"] = "10KW";
            sampleRow["容积"] = "100L";
            sampleRow["压力"] = "5MPa";
            sampleRow["温度"] = "100℃";
            sampleRow["直径"] = "100mm";
            sampleRow["外径"] = "10mm";
            sampleRow["内径"] = "90mm";
            sampleRow["厚度"] = "10mm";
            sampleRow["重量"] = "10Kg";
            sampleRow["型号"] = "A4";
            sampleRow["备注"] = "备注";
            sampleRow["自定义1"] = "自定义1";
            sampleRow["自定义2"] = "自定义2";
            sampleRow["自定义3"] = "自定义3";
            #endregion
            dt.Rows.Add(sampleRow);

            return dt;
        }

        /// <summary>
        /// 导出DataTable到Excel
        /// </summary>
        public bool ExportDataTableToExcel(DataTable dataTable, string filePath)
        {
            try
            {
                // 使用EPPlus库导出Excel（推荐方式）
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("图元批量添加模板");

                    // 添加标题行
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                        worksheet.Cells[1, i + 1].Style.Font.Bold = true;

                        // 修复ExcelFillPatternType的引用问题
                        worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                    }

                    // 添加数据行
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                        }
                    }

                    // 自动调整列宽
                    worksheet.Cells.AutoFitColumns();

                    // 保存文件
                    var fileInfo = new FileInfo(filePath);
                    package.SaveAs(fileInfo);
                }

                return true;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"导出Excel时出错: {ex.Message}");
                return false;
            }
        }

        #endregion

        /// <summary>
        /// 添加当前图形入库“从当前图形拾取”按钮的点击事件处理器
        /// </summary>
        private async void ImportFromSelection_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
                {
                    MessageBox.Show("数据库未连接，无法执行导入操作。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (_selectedCategoryNode == null)
                {
                    MessageBox.Show("请先在左侧的分类树中选择一个要导入的目标分类。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 重要：交互式CAD选择必须在主线程（AutoCAD/UI线程）执行，避免 Task.Run 导致跨线程访问控件或API
                ImportEntityDto dto = null;
                try
                {
                    // 直接调用选择方法，这个方法内部应该处理好与CAD的交互，确保在正确的线程上执行
                    dto = Helpers.SelectionImportHelper.PickAndReadEntity();
                }
                catch (Exception exSel)
                {
                    LogManager.Instance.LogError($"CAD 选择或读取图元时失败: {exSel.Message}");
                    MessageBox.Show($"从当前图形读取图元失败: {exSel.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (dto == null)
                {
                    LogManager.Instance.LogInfo("用户取消了实体选择或未选择有效实体。");
                    return; // 用户取消或无效选择
                }

                // 把当前选中的分类信息设定到 DTO
                dto.FileStorage.CategoryId = _selectedCategoryNode.Id;
                dto.FileStorage.CategoryType = _selectedCategoryNode.Level == 0 ? "main" : "sub";
                // 属性文本
                ApplyAttributeTextToDto(dto);
                // 在 UI 线程显示确认窗口（安全）
                var owner = Window.GetWindow(this);
                var confirmWindow = new Views.ImportConfirmWindow(dto, this);
                if (owner != null) confirmWindow.Owner = owner;

                bool? dialogResult = null;
                try
                {
                    dialogResult = confirmWindow.ShowDialog();
                }
                catch (Exception exShow)
                {
                    LogManager.Instance.LogError($"显示导入确认窗口失败: {exShow.Message}");
                    MessageBox.Show($"打开确认界面失败: {exShow.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (dialogResult == true)
                {
                    try
                    {
                        // 导入后刷新当前分类文件列表（在 UI 线程上安全等待）
                        await LoadFilesForCategoryAsync(_selectedCategoryNode);
                        MessageBox.Show("图元导入成功！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception exRefresh)
                    {
                        LogManager.Instance.LogError($"导入后刷新列表失败: {exRefresh.Message}");
                        MessageBox.Show($"导入完成，但刷新列表失败: {exRefresh.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"从CAD拾取导入失败: {ex.Message}");
                MessageBox.Show($"操作失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 为ImportConfirmWindow提供一个设置内部状态的公共方法
        /// </summary>
        public void SetSelectedFileForImport(ImportEntityDto dto)
        {
            // 这个方法让确认窗口能把最终确认的数据传递回来
            _selectedFilePath = dto.FileStorage.FilePath; // 虽然是虚拟路径，但流程需要
            _selectedPreviewImagePath = dto.PreviewImagePath;
            _currentFileStorage = dto.FileStorage;
            _currentFileAttribute = dto.FileAttribute;
        }
        /// <summary>
        /// 确认导入
        /// </summary>
        /// <param name="dto"></param>
        private void ApplyAttributeTextToDto(ImportEntityDto dto)
        {
            if (dto == null || dto.FileAttribute == null || dto.FileStorage == null) return;

            // 1) 从 Remarks（或其他字符串）提取 key/value 对
            var raw = dto.FileAttribute.Remarks ?? string.Empty;
            var pairs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 支持多种分隔：冒号、等号、空格；每行一个
            foreach (var line in raw.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var l = line.Trim();
                if (string.IsNullOrEmpty(l)) continue;

                string key = null, val = null;
                int idx;
                if ((idx = l.IndexOf(':')) >= 0)
                {
                    key = l.Substring(0, idx).Trim();
                    val = l.Substring(idx + 1).Trim();
                }
                else if ((idx = l.IndexOf('=')) >= 0)
                {
                    key = l.Substring(0, idx).Trim();
                    val = l.Substring(idx + 1).Trim();
                }
                else
                {
                    // 尝试以第一个空白为分隔
                    var parts = l.Split(new[] { ' ' }, 2, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length == 2)
                    {
                        key = parts[0].Trim();
                        val = parts[1].Trim();
                    }
                    else
                    {
                        // 如果无法解析，放进 Remarks 合并字段备用
                        continue;
                    }
                }

                if (!string.IsNullOrEmpty(key))
                {
                    if (!pairs.ContainsKey(key))
                        pairs[key] = val ?? string.Empty;
                }
            }

            if (pairs.Count == 0) return;

            // 2) 准备候选属性字典（FileStorage 与 FileAttribute 的 PropertyInfo）
            var fsType = typeof(FileStorage);
            var faType = typeof(FileAttribute);
            var fsProps = fsType.GetProperties().Where(p => p.CanWrite).ToList();
            var faProps = faType.GetProperties().Where(p => p.CanWrite).ToList();

            // 3) 映射表（优先使用属性名、其次使用显示名映射 _propertyDisplayNameMap）
            string Normalize(string s)
            {
                if (string.IsNullOrEmpty(s)) return string.Empty;
                var sb = new StringBuilder();
                foreach (var ch in s.ToLowerInvariant())
                {
                    if (char.IsLetterOrDigit(ch) || ch == '.') sb.Append(ch);
                }
                return sb.ToString();
            }

            // 构建 displayName -> propertyName 映射（反查）
            var displayToProp = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in DictionaryHelper._propertyDisplayNameMap)
            {
                // kv.Key 是属性名，kv.Value 是显示名（中文）
                displayToProp[kv.Value] = kv.Key;
            }

            // 4) 尝试匹配并设置属性
            foreach (var kv in pairs)
            {
                var tag = kv.Key;
                var val = kv.Value;
                if (string.IsNullOrWhiteSpace(tag)) continue;

                var nTag = Normalize(tag);

                bool matched = false;

                // 尝试直接匹配 FileAttribute 属性名
                foreach (var prop in faProps)
                {
                    var pName = prop.Name;
                    if (Normalize(pName) == nTag)
                    {
                        TrySetPropertyValue(dto.FileAttribute, prop, val);
                        matched = true;
                        break;
                    }
                    // 尝试匹配显示名（_propertyDisplayNameMap）
                    if (DictionaryHelper._propertyDisplayNameMap.TryGetValue(pName, out var dispName))
                    {
                        if (Normalize(dispName) == nTag || Normalize(dispName).Contains(nTag) || nTag.Contains(Normalize(dispName)))
                        {
                            TrySetPropertyValue(dto.FileAttribute, prop, val);
                            matched = true;
                            break;
                        }
                    }
                }
                if (matched) continue;

                // 尝试直接匹配 FileStorage 属性名
                foreach (var prop in fsProps)
                {
                    var pName = prop.Name;
                    if (Normalize(pName) == nTag)
                    {
                        TrySetPropertyValue(dto.FileStorage, prop, val);
                        matched = true;
                        break;
                    }
                    if (DictionaryHelper._propertyDisplayNameMap.TryGetValue(pName, out var dispName))
                    {
                        if (Normalize(dispName) == nTag || Normalize(dispName).Contains(nTag) || nTag.Contains(Normalize(dispName)))
                        {
                            TrySetPropertyValue(dto.FileStorage, prop, val);
                            matched = true;
                            break;
                        }
                    }
                }
                if (matched) continue;

                // 5) 若仍未匹配，尝试容错匹配：按包含关系匹配显示名或属性名
                foreach (var prop in faProps.Concat(fsProps))
                {
                    var pName = prop.Name;
                    if (DictionaryHelper._propertyDisplayNameMap.TryGetValue(pName, out var dispName))
                    {
                        if (Normalize(dispName).IndexOf(nTag, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            nTag.IndexOf(Normalize(dispName), StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // decide target object
                            var target = prop.DeclaringType == faType ? (object)dto.FileAttribute : (object)dto.FileStorage;
                            TrySetPropertyValue(target, prop, val);
                            matched = true;
                            break;
                        }
                    }
                    else
                    {
                        if (Normalize(pName).IndexOf(nTag, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            nTag.IndexOf(Normalize(pName), StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            var target = prop.DeclaringType == faType ? (object)dto.FileAttribute : (object)dto.FileStorage;
                            TrySetPropertyValue(target, prop, val);
                            matched = true;
                            break;
                        }
                    }
                }
            }

            // 辅助：若 Remarks 中包含属性图片或特殊字段，也可复制到描述
            if (string.IsNullOrEmpty(dto.FileStorage.Description) && !string.IsNullOrEmpty(raw))
            {
                dto.FileStorage.Description = raw.Length > 500 ? raw.Substring(0, 500) : raw;
            }
        }

        /// <summary>
        /// 读取属性图片
        /// </summary>
        /// <param name="target"></param>
        /// <param name="prop"></param>
        /// <param name="rawValue"></param>
        private void TrySetPropertyValue(object target, System.Reflection.PropertyInfo prop, string rawValue)
        {
            if (target == null || prop == null || rawValue == null) return;

            try
            {
                var targetType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;

                if (targetType == typeof(string))
                {
                    prop.SetValue(target, rawValue);
                    return;
                }

                if (string.IsNullOrWhiteSpace(rawValue))
                    return;

                if (targetType == typeof(int))
                {
                    if (int.TryParse(rawValue, out var iv)) prop.SetValue(target, iv);
                    else if (double.TryParse(rawValue, out var dv)) prop.SetValue(target, Convert.ToInt32(dv));
                    return;
                }

                if (targetType == typeof(long))
                {
                    if (long.TryParse(rawValue, out var lv)) prop.SetValue(target, lv);
                    else if (double.TryParse(rawValue, out var dv)) prop.SetValue(target, Convert.ToInt64(dv));
                    return;
                }

                if (targetType == typeof(decimal))
                {
                    if (decimal.TryParse(rawValue, out var dec)) prop.SetValue(target, dec);
                    return;
                }

                if (targetType == typeof(double))
                {
                    if (double.TryParse(rawValue, out var d)) prop.SetValue(target, d);
                    return;
                }

                if (targetType == typeof(bool))
                {
                    var lower = rawValue.Trim().ToLowerInvariant();
                    if (lower == "是" || lower == "true" || lower == "1") prop.SetValue(target, true);
                    else if (lower == "否" || lower == "false" || lower == "0") prop.SetValue(target, false);
                    return;
                }

                if (targetType == typeof(DateTime))
                {
                    if (DateTime.TryParse(rawValue, out var dt)) prop.SetValue(target, dt);
                    return;
                }

                // 退化策略：对非复杂类型尝试 Convert.ChangeType
                if (targetType.IsPrimitive)
                {
                    var converted = Convert.ChangeType(rawValue, targetType);
                    prop.SetValue(target, converted);
                    return;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"TrySetPropertyValue 失败: prop={prop.Name}, val='{rawValue}', err={ex.Message}");
            }
        }

        #region 管道相关操作
        private void 加载Excel表_Click(object sender, RoutedEventArgs e)
        {
            Env.Document.SendStringToExecute("ImportTableFromExcel ", false, false, false);
            
        }

        private void 属性同步_Click(object sender, RoutedEventArgs e)
        {
            Env.Document.SendStringToExecute("SyncPipeProperties ", false, false, false);

            //Env.Document.SendStringToExecute("PreviewPipeGeometry ", false, false, false);
        }

        private void 表格同步_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("未找到活动的 AutoCAD 文档。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 在 AutoCAD 文档锁内进行交互选择并判断表类型
                using (doc.LockDocument())
                {
                    var ed = doc.Editor;

                    // 让用户选择一个 Table 对象
                    var peo = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\n请选择要同步的表格（Table）：");
                    peo.SetRejectMessage("\n请选择一个表格对象。");
                    peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Table), true);

                    var per = ed.GetEntity(peo);
                    if (per.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        return;

                    using (var tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        var table = tr.GetObject(per.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Table;
                        if (table == null)
                        {
                            MessageBox.Show("选中的对象不是表格。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                            return;
                        }

                        // 读取前几行作为表头用于判断（最多读取前3行）
                        int headerRowsToCheck = Math.Min(3, table.Rows.Count);
                        // 创建一个 StringBuilder 用于存储表头文本 拼接表头文本
                        var headerTextSb = new System.Text.StringBuilder();
                        // 遍历每一行 遍历表格的前几行和所有列
                        for (int r = 0; r < headerRowsToCheck; r++)
                        {
                            // 遍历每一列
                            for (int c = 0; c < table.Columns.Count; c++)
                            {
                                try
                                {
                                    // 尝试获取单元格文本 读取单元格文本
                                    var cellText = (table.Cells[r, c].TextString ?? string.Empty).Trim();
                                    // 检查单元格文本是否非空
                                    if (!string.IsNullOrEmpty(cellText))
                                    {
                                        // 将单元格文本添加到表头文本中 拼接单元格文本
                                        headerTextSb.Append(cellText).Append(" ");
                                    }
                                }
                                catch
                                {
                                    // 忽略单元格读取错误
                                }
                            }
                        }
                        // 将表头文本转换为字符串 拼接后的表头文本
                        string headerText = headerTextSb.ToString();

                        // 关键词判断：若包含任一管道关键词则判定为管道表；包含设备关键词则判定为设备表
                        var pipeKeywords = new[] { "管道", "直径", "外径", "内径", "厚度", "单位", "数量", "米", "长度", "规格" };
                        var equipKeywords = new[] { "部件ID", "部件编号", "序号", "名称", "型号", "材质", "数量", "规格", "部件" };
                        // 判断表头文本是否包含管道关键词 判断表头文本中是否包含关键词
                        bool looksLikePipe = pipeKeywords.Any(k => !string.IsNullOrWhiteSpace(k) && headerText.IndexOf(k, System.StringComparison.OrdinalIgnoreCase) >= 0);
                        // 设备表 判断表头文本是否包含设备关键词
                        bool looksLikeEquip = equipKeywords.Any(k => !string.IsNullOrWhiteSpace(k) && headerText.IndexOf(k, System.StringComparison.OrdinalIgnoreCase) >= 0);
                        // 选择命令 根据判断结果选择要执行的命令
                        string chosenCommand = null;
                        // 管道表 根据判断结果选择命令
                        if (looksLikePipe && !looksLikeEquip)
                        {
                            // 管道表 管道表同步
                            chosenCommand = "SyncTableToEntities ";
                        }
                        else if (looksLikeEquip && !looksLikePipe)
                        {
                            // 设备表 设备表同步
                            chosenCommand = "SyncDeviceTableToBlocks ";
                        }
                        else
                        {
                            // 无法明确判断或两者同时匹配，询问用户确认
                            var result = MessageBox.Show("未能自动判断表类型或同时匹配到管道/设备特征。\n请选择要执行的同步类型：\n\n[是] - 设备表同步（表 -> 图元）\n[否] - 管道表同步（表 -> 管道实体）\n[取消] - 取消操作",
                                                         "请选择表类型",
                                                         MessageBoxButton.YesNoCancel,
                                                         MessageBoxImage.Question);

                            if (result == MessageBoxResult.Cancel)
                            {
                                return;
                            }
                            else if (result == MessageBoxResult.Yes)
                            {
                                // 设备表同步
                                chosenCommand = "SyncDeviceTableToBlocks ";
                            }
                            else
                            {
                                // 管道表同步
                                chosenCommand = "SyncTableToEntities ";
                            }
                        }

                        tr.Commit();

                        // 通过 Env.Document.SendStringToExecute 调用对应的命令（命令内部会再次提示选择或使用当前上下文）
                        if (!string.IsNullOrEmpty(chosenCommand))
                        {
                            // 在发送命令前提示用户（可省略）
                            ed.WriteMessage($"\n将执行: {chosenCommand.Trim()}");
                            Env.Document.SendStringToExecute(chosenCommand, false, false, false);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"表格同步失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private void 导出表格_Click(object sender, RoutedEventArgs e)
        {
            Env.Document.SendStringToExecute("ExportTableToExcel ", false, false, false);

        }

        private void 生成设备表_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Env.Document.SendStringToExecute("GenerateDeviceTable ", false, false, false);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行生成设备表失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private void 绘制进口管道_Click(object sender, RoutedEventArgs e)
        {
            VariableDictionary.textBoxScale = GetDrawingScaleFromTextBox();
            Env.Document.SendStringToExecute("DrawInletPipeByClicks ", false, false, false);
            //Env.Document.SendStringToExecute("Draw_GD_PipeLine_DynamicBlock ", false, false, false);
        }

        private void 绘制出口管道_Click(object sender, RoutedEventArgs e)
        {
            VariableDictionary.textBoxScale = GetDrawingScaleFromTextBox();
            Env.Document.SendStringToExecute("DrawOutletPipeByClicks ", false, false, false);
            //Env.Document.SendStringToExecute("Draw_GD_PipeLine_DynamicBlock ", false, false, false);
        }

        #endregion

        #region 图层管理器相关
        /// <summary>
        /// [加载当前图层]按键事件 - 完善版
        /// </summary>
        private void 加载当前图层_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var layers = _layerManager.LoadCurrentLayers();//层管理器加载当前图层
                _layerData.Clear();

                for (int i = 0; i < layers.Count; i++)
                {
                    layers[i].DisplayIndex = i + 1;
                    _layerData.Add(layers[i]);
                }

                Env.Editor?.WriteMessage($"\n成功加载 {layers.Count} 个图层");
            }
            catch (Exception ex)
            {
                Env.Editor?.WriteMessage($"\n加载图层时出错: {ex.Message}");
            }
        }
        /// <summary>
        /// [保存图层配置]按键事件 - 完善版
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 保存图层配置_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                bool success = _layerManager.SaveLayersToExcel();
                if (success)
                {
                    Env.Editor?.WriteMessage("\n图层配置已保存");
                    MessageBox.Show("图层配置已保存", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    Env.Editor?.WriteMessage("\n保存图层配置失败");
                    MessageBox.Show("保存图层配置失败", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                Env.Editor?.WriteMessage($"\n保存配置时出错: {ex.Message}");
                MessageBox.Show($"保存配置时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// [加载本地图层]按键事件 - 完善版
        /// </summary>
        private void 加载本地图层_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var layers = _layerManager.LoadLayersFromExcel();
                _layerData.Clear();

                for (int i = 0; i < layers.Count; i++)
                {
                    layers[i].DisplayIndex = i + 1;
                    _layerData.Add(layers[i]);
                }

                Env.Editor?.WriteMessage($"\n成功加载 {layers.Count} 个图层配置");
                MessageBox.Show($"成功加载 {layers.Count} 个图层配置", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Env.Editor?.WriteMessage($"\n加载本地配置时出错: {ex.Message}");
                MessageBox.Show($"加载本地配置时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// [应用图层]按键事件 - 完善版
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 应用图层_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _layerManager.CreateBeforeApplySnapshot();
                bool success = _layerManager.ApplyLayers(_layerData);

                if (success)
                {
                    Env.Editor?.WriteMessage("\n图层配置已应用");
                    MessageBox.Show("图层配置已应用", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    Env.Editor?.WriteMessage("\n应用图层配置失败");
                    MessageBox.Show("应用图层配置失败", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                Env.Editor?.WriteMessage($"\n应用图层时出错: {ex.Message}");
                MessageBox.Show($"应用图层时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// [还原图层]按键事件 - 完善版
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 还原图层_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                bool success = _layerManager.RestorePreviousState();
                if (success)
                {
                    Env.Editor?.WriteMessage("\n已还原到应用前的图层状态");
                    加载当前图层_Btn_Click(null, null); // 重新加载当前图层
                    MessageBox.Show("已还原到应用前的图层状态", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    Env.Editor?.WriteMessage("\n还原图层状态失败");
                    MessageBox.Show("还原图层状态失败", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                Env.Editor?.WriteMessage($"\n还原图层时出错: {ex.Message}");
                MessageBox.Show($"还原图层时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>
        /// 图层管理器
        /// </summary>
        public class LayerManager
        {
            /// <summary>
            /// 图层信息
            /// </summary>
            private List<LayerInfo> _layerInfo;
            /// <summary>
            /// 图层状态快照
            /// </summary>
            private LayerStateSnapshot _beforeApplySnapshot;
            /// <summary>
            /// 工作目录
            /// </summary>
            private string _workingDirectory;
            /// <summary>
            /// 构造函数
            /// </summary>
            public LayerManager()
            {
                _layerInfo = new List<LayerInfo>();
                _workingDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                               "GB_NewCadPlus_LM");
                if (!Directory.Exists(_workingDirectory))
                    Directory.CreateDirectory(_workingDirectory);
            }

            /// <summary>
            /// 加载当前图纸的所有图层
            /// </summary>
            public List<LayerInfo> LoadCurrentLayers()
            {
                _layerInfo.Clear();
                var layers = new List<LayerInfo>();

                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return layers;

                Database db = doc.Database;
                Editor ed = doc.Editor;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    int index = 1;

                    foreach (ObjectId layerId in lt)
                    {
                        LayerTableRecord layer = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
                        if (layer != null)
                        {
                            var layerInfo = new LayerInfo
                            {
                                Index = index++,
                                LayerName = layer.Name,
                                IsOn = !layer.IsOff,
                                IsFrozen = layer.IsFrozen,
                                ColorIndex = layer.Color.ColorIndex,
                                Color = layer.Color,
                                IsDelete = false
                            };
                            layers.Add(layerInfo);
                        }
                    }
                    tr.Commit();
                }

                // 按名称排序
                layers.Sort((x, y) => string.Compare(x.LayerName, y.LayerName, StringComparison.OrdinalIgnoreCase));
                _layerInfo = new List<LayerInfo>(layers);
                return layers;
            }

            /// <summary>
            /// 保存图层配置到Excel
            /// </summary>
            public bool SaveLayersToExcel(string fileName = "layer_config.xlsx")
            {
                try
                {
                    string filePath = Path.Combine(_workingDirectory, fileName);

                    using (var package = new ExcelPackage())
                    {
                        var worksheet = package.Workbook.Worksheets.Add("图层配置");

                        worksheet.Cells[1, 1].Value = "序号";
                        worksheet.Cells[1, 2].Value = "图层名称";
                        worksheet.Cells[1, 3].Value = "开关";
                        worksheet.Cells[1, 4].Value = "冻结";
                        worksheet.Cells[1, 5].Value = "颜色索引";
                        worksheet.Cells[1, 6].Value = "删除";

                        for (int i = 0; i < _layerInfo.Count; i++)
                        {
                            var layer = _layerInfo[i];
                            int row = i + 2;

                            worksheet.Cells[row, 1].Value = layer.Index;
                            worksheet.Cells[row, 2].Value = layer.LayerName;
                            worksheet.Cells[row, 3].Value = layer.IsOn;
                            worksheet.Cells[row, 4].Value = layer.IsFrozen;
                            worksheet.Cells[row, 5].Value = layer.ColorIndex;
                            worksheet.Cells[row, 6].Value = layer.IsDelete;
                        }

                        FileInfo fileInfo = new FileInfo(filePath);
                        package.SaveAs(fileInfo);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Editor ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                    ed?.WriteMessage($"\n保存图层配置时出错: {ex.Message}");
                    return false;
                }
            }

            /// <summary>
            /// 从Excel加载图层配置
            /// </summary>
            public List<LayerInfo> LoadLayersFromExcel(string fileName = "layer_config.xlsx")
            {
                try
                {
                    string filePath = Path.Combine(_workingDirectory, fileName);
                    if (!File.Exists(filePath))
                    {
                        Editor ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                        ed?.WriteMessage("\n配置文件不存在");
                        return new List<LayerInfo>();
                    }

                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        var layers = new List<LayerInfo>();

                        int rowCount = worksheet.Dimension?.Rows ?? 0;
                        if (rowCount <= 1) return new List<LayerInfo>();

                        for (int row = 2; row <= rowCount; row++)
                        {
                            try
                            {
                                var index = Convert.ToInt32(worksheet.Cells[row, 1].Value ?? 0);
                                var name = worksheet.Cells[row, 2].Value?.ToString() ?? "";
                                var isOn = Convert.ToBoolean(worksheet.Cells[row, 3].Value ?? true);
                                var isFrozen = Convert.ToBoolean(worksheet.Cells[row, 4].Value ?? false);
                                var colorIndex = Convert.ToInt16(worksheet.Cells[row, 5].Value ?? 7);
                                var toDelete = Convert.ToBoolean(worksheet.Cells[row, 6].Value ?? false);

                                var layerInfo = new LayerInfo
                                {
                                    Index = index,
                                    LayerName = name,
                                    IsOn = isOn,
                                    IsFrozen = isFrozen,
                                    ColorIndex = colorIndex,
                                    IsDelete = toDelete,
                                    Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                                        Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex)
                                };
                                layers.Add(layerInfo);
                            }
                            catch (Exception rowEx)
                            {
                                Editor ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                                ed?.WriteMessage($"\n读取第{row}行数据时出错: {rowEx.Message}");
                            }
                        }

                        // 按名称排序
                        layers.Sort((x, y) => string.Compare(x.LayerName, y.LayerName, StringComparison.OrdinalIgnoreCase));
                        _layerInfo = new List<LayerInfo>(layers);
                        return layers;
                    }
                }
                catch (Exception ex)
                {
                    Editor ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                    ed?.WriteMessage($"\n加载图层配置时出错: {ex.Message}");
                    return new List<LayerInfo>();
                }
            }

            /// <summary>
            /// 创建应用前的状态快照
            /// </summary>
            public void CreateBeforeApplySnapshot()
            {
                _beforeApplySnapshot = new LayerStateSnapshot();
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                Database db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    foreach (ObjectId layerId in lt)
                    {
                        LayerTableRecord layer = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
                        if (layer != null)
                        {
                            var layerInfo = new LayerInfo
                            {
                                LayerName = layer.Name,
                                IsOn = !layer.IsOff,
                                IsFrozen = layer.IsFrozen,
                                ColorIndex = layer.Color.ColorIndex,
                                IsDelete = false
                            };
                            _beforeApplySnapshot.Layers[layer.Name] = layerInfo;
                        }
                    }
                    tr.Commit();
                }
            }

            /// <summary>
            /// 应用图层配置
            /// </summary>
            public bool ApplyLayers(ObservableCollection<LayerInfo> layerData)
            {
                try
                {
                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    if (doc == null) return false;

                    Database db = doc.Database;
                    Editor ed = doc.Editor;
                    using (doc.LockDocument())
                    using (var tr = new DBTrans())
                    {
                        foreach (var layerInfo in layerData)
                        {
                            if (tr.LayerTable.Has(layerInfo.LayerName))
                            {
                                var layerObjectId = tr.LayerTable[layerInfo.LayerName];

                                if (layerInfo.IsDelete)
                                {
                                    DeleteAllEntitiesOnLayer(tr, db, layerObjectId, layerInfo.LayerName);
                                    tr.LayerTable.Remove(layerObjectId);
                                    continue;
                                }
                            }
                        }
                        tr.Commit();
                        Env.Editor.Redraw();
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Editor ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                    ed?.WriteMessage($"\n应用图层配置时出错: {ex.Message}");
                    return false;
                }
            }

            /// <summary>
            /// 删除图层上的所有实体
            /// </summary>
            private int DeleteAllEntitiesOnLayer(DBTrans tr, Database db, ObjectId layerId, string layerName)
            {
                int deletedCount = 0;
                try
                {
                    foreach (ObjectId blockTableId in tr.BlockTable)
                    {
                        BlockTableRecord blockTableRecord = tr.GetObject(blockTableId, OpenMode.ForWrite) as BlockTableRecord;
                        var entitiesToDelete = new List<ObjectId>();

                        foreach (ObjectId entId in blockTableRecord)
                        {
                            Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                            if (ent != null && ent.LayerId == layerId)
                            {
                                entitiesToDelete.Add(entId);
                            }
                        }

                        foreach (ObjectId entId in entitiesToDelete)
                        {
                            try
                            {
                                Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                                if (ent != null)
                                {
                                    ent.Erase(true);
                                    deletedCount++;
                                }
                            }
                            catch (Exception entDeleteEx)
                            {
                                Env.Editor?.WriteMessage($"\n删除实体时出错: {entDeleteEx.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Env.Editor?.WriteMessage($"\n遍历实体时出错: {ex.Message}");
                }
                return deletedCount;
            }


            /// <summary>
            /// 还原到应用前的状态
            /// </summary>
            public bool RestorePreviousState()
            {
                if (_beforeApplySnapshot == null || _beforeApplySnapshot.Layers.Count == 0)
                {
                    Editor ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                    ed?.WriteMessage("\n没有可还原的图层状态");
                    return false;
                }

                try
                {
                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    if (doc == null) return false;

                    Database db = doc.Database;
                    Editor ed = doc.Editor;

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable;

                        foreach (var kvp in _beforeApplySnapshot.Layers)
                        {
                            string layerName = kvp.Key;
                            LayerInfo originalState = kvp.Value;

                            if (lt.Has(layerName))
                            {
                                LayerTableRecord layer = tr.GetObject(lt[layerName], OpenMode.ForWrite) as LayerTableRecord;
                                layer.IsOff = !originalState.IsOn;
                                layer.IsFrozen = originalState.IsFrozen;
                                layer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                                    Autodesk.AutoCAD.Colors.ColorMethod.ByAci, originalState.ColorIndex);
                            }
                            else
                            {
                                using (var newLayer = new LayerTableRecord())
                                {
                                    newLayer.Name = layerName;
                                    newLayer.IsOff = !originalState.IsOn;
                                    newLayer.IsFrozen = originalState.IsFrozen;
                                    newLayer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                                        Autodesk.AutoCAD.Colors.ColorMethod.ByAci, originalState.ColorIndex);

                                    lt.Add(newLayer);
                                    tr.AddNewlyCreatedDBObject(newLayer, true);
                                }
                            }
                        }
                        tr.Commit();
                        ed.Regen();
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Editor ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                    ed?.WriteMessage($"\n还原图层状态时出错: {ex.Message}");
                    return false;
                }
            }

        }
        #endregion

        #region 管道表生成器
        /// <summary>
        /// 新增：保存最近一次生成的临时表文件路径（以便后续插入）
        /// </summary>
        private string _lastSavedPipeTablePath;

        /// <summary>
        /// 新增：记录上次生成表所涉及的图层列表（供插入时选择）
        /// </summary>
        private List<string> _lastSavedPipeTableLayers = new List<string>();

        /// <summary>
        /// 生成管道表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void 生成管道表_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 在调用选择/统计前确保比例被刷新到缓存，全局代码（例如 SavePipeTableToTempDwg）可读取该比例
                try
                {
                    AutoCadHelper.GetAndApplyActiveDrawingScale();
                }
                catch (Exception exScale)
                {
                    LogManager.Instance.LogWarning($"读取/应用当前绘图比例失败: {exScale.Message}");
                }

                // 触发在 ElementAndTable 中实现的命令（交互式选择将在 CAD 端执行）
                Env.Document.SendStringToExecute("GeneratePipeTableFromSelection ", false, false, false);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"触发生成管道表命令失败: {ex.Message}");
                MessageBox.Show($"生成管道表失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 插入管道表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 插入管道表_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_lastSavedPipeTablePath) || !File.Exists(_lastSavedPipeTablePath))
                {
                    MessageBox.Show("未找到已保存的临时表文件，请先点击【生成管道表】并完成保存。", "未找到文件", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("未找到活动的 AutoCAD 文档。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                using (doc.LockDocument())
                {
                    var ed = doc.Editor;

                    // 提示用户切换视口并确认
                    var pko = new Autodesk.AutoCAD.EditorInput.PromptKeywordOptions("\n请将鼠标移动到目标视口（布局请先双击进入视口），然后选择：\n[继续] 继续 / [取消] 取消")
                    {
                        AllowNone = true
                    };
                    pko.Keywords.Add("继续");
                    pko.Keywords.Add("取消");
                    try { pko.Keywords.Default = "继续"; } catch { }

                    var pkr = ed.GetKeywords(pko);
                    if (pkr.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK || string.Equals(pkr.StringResult, "取消", StringComparison.OrdinalIgnoreCase))
                        return;

                    // 拾取插入点
                    var ppo = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\n请选择表格插入点（请在目标视口内拾取）：");
                    var ppr = ed.GetPoint(ppo);
                    if (ppr.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                    {
                        MessageBox.Show("未选择插入点，操作已取消。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    // 如果是在布局中插入，提示用户输入视口比例分母以便缩放
                    double insertionScale = 1.0;
                    var scalePrompt = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n如果在布局/纸空间插入，请输入视口比例的分母（例如 100 表示 1:100），回车表示 1：")
                    {
                        AllowNone = true,
                        DefaultValue = 1.0,
                        UseDefaultValue = true,
                        AllowZero = false,
                        AllowNegative = false
                    };
                    var scaleRes = ed.GetDouble(scalePrompt);
                    if (scaleRes.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK && scaleRes.Value > 0.0)
                        insertionScale = 1.0 / scaleRes.Value; // 实际缩放因子

                    // 插入块（返回 BlockReference 的 ObjectId）
                    var insertedBrId = AutoCadHelper.InsertBlockFromExternalDwg(_lastSavedPipeTablePath, Path.GetFileNameWithoutExtension(_lastSavedPipeTablePath), ppr.Value);

                    if (insertedBrId == Autodesk.AutoCAD.DatabaseServices.ObjectId.Null)
                    {
                        MessageBox.Show("插入失败：未能将临时 DWG 中的块导入并插入。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // 如果需要缩放（例如布局视口按比例显示），调整已插入的 BlockReference 的 ScaleFactors
                    if (Math.Abs(insertionScale - 1.0) > 1e-9)
                    {
                        try
                        {
                            using (var tr = doc.Database.TransactionManager.StartTransaction())
                            {
                                var br = tr.GetObject(insertedBrId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.BlockReference;
                                if (br != null)
                                {
                                    br.ScaleFactors = new Autodesk.AutoCAD.Geometry.Scale3d(insertionScale, insertionScale, insertionScale);
                                }
                                tr.Commit();
                            }
                        }
                        catch (Exception exScale)
                        {
                            LogManager.Instance.LogWarning($"调整插入比例时出错: {exScale.Message}");
                        }
                    }

                    // 新增：将插入的表设置到“选定图元图层”中，并确保块定义内实体也使用该图层
                    try
                    {
                        string targetLayer = null;
                        if (_lastSavedPipeTableLayers != null && _lastSavedPipeTableLayers.Count > 0)
                        {
                            if (_lastSavedPipeTableLayers.Count == 1)
                            {
                                targetLayer = _lastSavedPipeTableLayers[0];
                            }
                            else
                            {
                                // 列出可选图层并让用户通过索引选择，避免输入图层名错误
                                var listText = new StringBuilder();
                                for (int i = 0; i < _lastSavedPipeTableLayers.Count; i++)
                                {
                                    listText.AppendLine($"{i + 1}. {_lastSavedPipeTableLayers[i]}");
                                }
                                ed.WriteMessage($"\n检测到生成表时涉及以下图层：\n{listText}\n请输入要用于表的图层序号（1 - {_lastSavedPipeTableLayers.Count}），回车使用默认第1项：");

                                var pio = new Autodesk.AutoCAD.EditorInput.PromptIntegerOptions("\n请输入序号：")
                                {
                                    AllowNone = true,
                                    DefaultValue = 1,
                                    LowerLimit = 1,
                                    UpperLimit = _lastSavedPipeTableLayers.Count,
                                    UseDefaultValue = true
                                };
                                var pir = ed.GetInteger(pio);
                                if (pir.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    int idx = Math.Max(1, Math.Min(_lastSavedPipeTableLayers.Count, pir.Value));
                                    targetLayer = _lastSavedPipeTableLayers[idx - 1];
                                }
                                else
                                {
                                    targetLayer = _lastSavedPipeTableLayers[0];
                                }
                            }
                        }

                        if (!string.IsNullOrEmpty(targetLayer))
                        {
                            using (var tr = doc.Database.TransactionManager.StartTransaction())
                            {
                                // 确保目标图层存在，否则创建
                                var lt = (Autodesk.AutoCAD.DatabaseServices.LayerTable)tr.GetObject(doc.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                if (!lt.Has(targetLayer))
                                {
                                    lt.UpgradeOpen();
                                    var ltr = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord
                                    {
                                        Name = targetLayer
                                    };
                                    lt.Add(ltr);
                                    tr.AddNewlyCreatedDBObject(ltr, true);
                                }

                                // 设置 BlockReference 的图层，并把块定义中所有实体的 Layer 设置为目标图层
                                var br = tr.GetObject(insertedBrId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.BlockReference;
                                if (br != null)
                                {
                                    // 设置 BlockReference 层
                                    br.Layer = targetLayer;

                                    // 获取块定义（BlockTableRecord）并把其子实体层设置为目标图层
                                    var btrId = br.BlockTableRecord;
                                    if (btrId != Autodesk.AutoCAD.DatabaseServices.ObjectId.Null)
                                    {
                                        var btr = tr.GetObject(btrId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.BlockTableRecord;
                                        if (btr != null)
                                        {
                                            foreach (ObjectId entId in btr)
                                            {
                                                try
                                                {
                                                    // 可能包含匿名/非实体项，安全转换
                                                    var ent = tr.GetObject(entId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.Entity;
                                                    if (ent != null)
                                                    {
                                                        // 跳过属性定义（AttributeDefinition）以免影响属性行为
                                                        if (ent is Autodesk.AutoCAD.DatabaseServices.AttributeDefinition) continue;

                                                        // 设置实体层为目标图层
                                                        ent.Layer = targetLayer;
                                                    }
                                                }
                                                catch
                                                {
                                                    // 忽略个别实体修改失败
                                                }
                                            }
                                        }
                                    }
                                }

                                tr.Commit();
                            }
                        }
                    }
                    catch (Exception exLayer)
                    {
                        LogManager.Instance.LogWarning($"设置插入表图层时出错: {exLayer.Message}");
                    }

                    MessageBox.Show("临时管道表已插入当前空间。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"插入管道表失败: {ex.Message}");
                MessageBox.Show($"插入管道表失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 表示管道汇总行
        /// </summary>
        private class PipeSummary
        {
            /// <summary>
            /// 序号
            /// </summary>
            public int Seq { get; set; }
            /// <summary>
            /// 名称
            /// </summary>
            public string? Name { get; set; }      // 示例可设为 "管道"
            /// <summary>
            /// 宽度规格
            /// </summary>
            public int WidthSpec { get; set; }          // y（整型，单位：毫米整数）
            /// <summary>
            /// 厚度规格
            /// </summary>
            public int ThicknessSpec { get; set; }      // z（整型）
            /// <summary>
            /// 累计长度（米）
            /// </summary>      
            public double QuantityMeters { get; set; }  // x    累加后以米为单位（double）
            /// <summary>
            /// 备注
            /// </summary>
            public string Remark { get; set; }

            /// <summary>
            /// 改为可写属性，兼容原先代码对 SpecString 的赋值需求
            /// </summary>
            private string _specString;
            public string SpecString
            {
                get
                {
                    // 如果没有显式设置，则返回基于宽厚的默认格式
                    if (string.IsNullOrEmpty(_specString))
                        return $"{WidthSpec} X {ThicknessSpec}";
                    return _specString;
                }
                set => _specString = value;
            }
            /// <summary>
            /// 新增：属性字典（属性名 -> 值）
            /// </summary>
            public Dictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 将管道汇总表生成到一个临时 DWG 文件（ModelSpace，表放在原点）。
        /// 文件名格式：{layerPart}_{yyyyMMdd_HHmmss}.dwg，保存在 %TEMP%\GB_NewCadPlus_LM 下。ParseLengthValueFromAttribute
        /// 生成后会把路径写入字段 `_lastSavedPipeTablePath` 并提示用户下一步操作。
        /// 列宽会根据表头与单元格内容自动计算（近似字符宽度）。
        /// </summary>
        private void SavePipeTableToTempDwg(List<PipeSummary> summaries, string material, double scaleDenom, List<string> attributesColumns = null)
        {
            try
            {
                if (summaries == null || summaries.Count == 0)
                {
                    MessageBox.Show("没有要保存的表数据。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                try
                {
                    if (!string.IsNullOrEmpty(_lastSavedPipeTablePath) && File.Exists(_lastSavedPipeTablePath))
                    {
                        try { File.Delete(_lastSavedPipeTablePath); } catch { /* 忽略删除失败 */ }
                    }
                }
                catch { }

                string buttonFolderName = "生成管道表";
                string tempDir = Path.Combine(Path.GetTempPath(), "GB_NewCadPlus_LM", buttonFolderName);
                if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);

                string firstLayer = summaries.FirstOrDefault()?.Name ?? "PipeTable";
                string safeLayer = string.Concat(firstLayer.Split(Path.GetInvalidFileNameChars())).Trim();
                if (string.IsNullOrWhiteSpace(safeLayer)) safeLayer = "PipeTable";

                string fileName = $"{safeLayer}_{DateTime.Now:yyyyMMdd_HHmmss}.dwg";
                string fullPath = Path.Combine(tempDir, fileName);
                string blockName = Path.GetFileNameWithoutExtension(fileName);

                // 处理比例分母：如果调用方传入 <= 0，则尝试自动从当前视口读取并转换为分母
                double denom = 1.0;
                if (scaleDenom > 0.0)
                {
                    denom = scaleDenom;
                }
                else
                {
                    try
                    {
                        // AutoCadHelper.GetScale(true) 可能返回规范化视口尺度因子（如 0.01 表示 1:100）或直接分母
                        double viewportScaleFactor = AutoCadHelper.GetScale(true);
                        denom = TextFontsStyleHelper.DetermineScaleDenominator(viewportScaleFactor, null, false);
                    }
                    catch
                    {
                        denom = 1.0;
                    }
                }

                // 防御性：确保 denom 合理
                if (double.IsNaN(denom) || double.IsInfinity(denom) || denom <= 0.0) denom = 1.0;

                using (var tempDb = new Autodesk.AutoCAD.DatabaseServices.Database(true, true))
                {
                    using (var tr = tempDb.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            // 确保目标图层存在
                            var ltForLayer = (Autodesk.AutoCAD.DatabaseServices.LayerTable)tr.GetObject(tempDb.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            if (!ltForLayer.Has(safeLayer))
                            {
                                var ltr = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord { Name = safeLayer };
                                ltForLayer.Add(ltr);
                                tr.AddNewlyCreatedDBObject(ltr, true);
                            }

                            var bt = (Autodesk.AutoCAD.DatabaseServices.BlockTable)tr.GetObject(tempDb.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            var btr = new Autodesk.AutoCAD.DatabaseServices.BlockTableRecord { Name = blockName };
                            bt.Add(btr);
                            tr.AddNewlyCreatedDBObject(btr, true);

                            // 构建列头列表：序号 | 名称 | (属性列...) | 单位 | 数量 | 备注
                            var attributeCols = attributesColumns ?? new List<string>();
                            int fixedTailCols = 3; // 单位, 数量, 备注
                            int cols = 2 + attributeCols.Count + fixedTailCols; // 序号 + 名称 + 属性列 + 尾部列
                            int rows = summaries.Count + 1;

                            var table = new Autodesk.AutoCAD.DatabaseServices.Table();
                            table.SetSize(rows, cols);
                            try { table.SetDatabaseDefaults(tempDb); } catch { }

                            // 设置表图层
                            try { table.Layer = safeLayer; } catch { }

                            table.Position = new Autodesk.AutoCAD.Geometry.Point3d(0, 0, 0);

                            // 依据 denom 计算文本高度与行高（保持原有经验常量，但按视口分母调整）
                            // 说明：denom 表示比例的分母，例如 1:100 -> denom=100
                            double baseTextHeight = 200.0; // 原经验值（保留）
                            double textHeight = baseTextHeight / denom; // 保持此前代码的计算方式，但 denom 现在来自视口或调用方
                            try { table.SetRowHeight(300.0 / denom); } catch { }

                            // 帮助函数：估算字符串“宽度”（以字符为单位，中文略宽）
                            double EstimateTextUnitLength(string s)
                            {
                                if (string.IsNullOrEmpty(s)) return 0.0;
                                double len = 0.0;
                                foreach (var ch in s)
                                {
                                    // 中文字符或全角字符视为 1.2 单位，拉丁字母与数字为 1.0 单位
                                    if (ch >= 0x4E00 && ch <= 0x9FFF) len += 1.2;
                                    else if (char.IsLetterOrDigit(ch) || char.IsPunctuation(ch) || char.IsWhiteSpace(ch)) len += 1.0;
                                    else len += 1.0;
                                }
                                return len;
                            }

                            // 先收集每列的最大字符“宽度”
                            var maxCharUnits = new double[cols];
                            for (int c = 0; c < cols; c++) maxCharUnits[c] = 0.0;

                            // 头部
                            var headers = new List<string> { "序号", "名称" };
                            headers.AddRange(attributeCols);
                            headers.AddRange(new[] { "单位", "数量", "备注" });

                            for (int c = 0; c < headers.Count && c < cols; c++)
                            {
                                var h = headers[c] ?? string.Empty;
                                maxCharUnits[c] = Math.Max(maxCharUnits[c], EstimateTextUnitLength(h));
                                table.Cells[0, c].TextString = h;
                                try { table.Cells[0, c].Alignment = Autodesk.AutoCAD.DatabaseServices.CellAlignment.MiddleCenter; } catch { }
                            }

                            // 数据写入同时更新最大宽度需求
                            for (int r = 0; r < summaries.Count; r++)
                            {
                                var s = summaries[r];
                                int rowIndex = r + 1;
                                // 序号
                                var seqStr = s.Seq.ToString();
                                maxCharUnits[0] = Math.Max(maxCharUnits[0], EstimateTextUnitLength(seqStr));
                                table.Cells[rowIndex, 0].TextString = seqStr;

                                // 名称
                                var nameStr = s.Name ?? "管道";
                                maxCharUnits[1] = Math.Max(maxCharUnits[1], EstimateTextUnitLength(nameStr));
                                table.Cells[rowIndex, 1].TextString = nameStr;

                                // 属性列
                                for (int ai = 0; ai < attributeCols.Count; ai++)
                                {
                                    var colName = attributeCols[ai];
                                    string value = "";
                                    if (s.Attributes != null)
                                    {
                                        var foundKey = s.Attributes.Keys.FirstOrDefault(k => k.Equals(colName, StringComparison.OrdinalIgnoreCase) || k.IndexOf(colName, StringComparison.OrdinalIgnoreCase) >= 0);
                                        if (foundKey != null) value = s.Attributes[foundKey] ?? string.Empty;
                                    }
                                    table.Cells[rowIndex, 2 + ai].TextString = value;
                                    maxCharUnits[2 + ai] = Math.Max(maxCharUnits[2 + ai], EstimateTextUnitLength(value));
                                }

                                // 单位/数量/备注
                                int unitCol = 2 + attributeCols.Count;
                                var unitStr = "米";
                                var qtyStr = s.QuantityMeters.ToString("F3");
                                var remarkStr = string.IsNullOrEmpty(s.Remark) ? (string.IsNullOrEmpty(material) ? "请填写材料" : material) : s.Remark;

                                table.Cells[rowIndex, unitCol].TextString = unitStr;
                                table.Cells[rowIndex, unitCol + 1].TextString = qtyStr;
                                table.Cells[rowIndex, unitCol + 2].TextString = remarkStr;

                                maxCharUnits[unitCol] = Math.Max(maxCharUnits[unitCol], EstimateTextUnitLength(unitStr));
                                maxCharUnits[unitCol + 1] = Math.Max(maxCharUnits[unitCol + 1], EstimateTextUnitLength(qtyStr));
                                maxCharUnits[unitCol + 2] = Math.Max(maxCharUnits[unitCol + 2], EstimateTextUnitLength(remarkStr));

                                for (int c = 0; c < cols; c++)
                                {
                                    try { table.Cells[rowIndex, c].Alignment = Autodesk.AutoCAD.DatabaseServices.CellAlignment.MiddleCenter; } catch { }
                                    try { table.Cells[rowIndex, c].TextHeight = textHeight; } catch { }
                                }
                            }

                            // 为所有单元设置文本高度（含表头）
                            for (int r = 0; r < rows; r++)
                                for (int c = 0; c < cols; c++)
                                    try { table.Cells[r, c].TextHeight = textHeight; } catch { }

                            // 根据每列需要的最大字符数计算列宽（经验系数）
                            double charWidthFactor = 0.6; // 经验值：字符宽约为文本高度的0.6倍（可微调）
                            double minColWidth = 600.0 / denom; // 最小列宽（防止过窄）
                            double maxColWidth = 8000.0 / denom; // 最大列宽限制

                            for (int c = 0; c < cols; c++)
                            {
                                try
                                {
                                    double estimatedWidth = Math.Ceiling(maxCharUnits[c] * textHeight * charWidthFactor);
                                    // 对于序号列设为较小最小宽度
                                    if (c == 0) estimatedWidth = Math.Max(estimatedWidth, 500.0 / denom);
                                    // 名称与备注列默认更宽一些
                                    if (c == 1 || (c == cols - 1)) estimatedWidth = Math.Max(estimatedWidth, 1200.0 / denom);

                                    double finalWidth = Math.Max(minColWidth, Math.Min(maxColWidth, estimatedWidth));
                                    table.SetColumnWidth(c, finalWidth);
                                }
                                catch
                                {
                                    // 忽略单列设置失败，继续下一列
                                }
                            }

                            // 把 table 加入块定义
                            btr.AppendEntity(table);
                            tr.AddNewlyCreatedDBObject(table, true);

                            // 确保块定义内实体使用 safeLayer
                            try
                            {
                                foreach (ObjectId entId in btr)
                                {
                                    try
                                    {
                                        var ent = tr.GetObject(entId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.Entity;
                                        if (ent != null && !(ent is Autodesk.AutoCAD.DatabaseServices.AttributeDefinition))
                                        {
                                            ent.Layer = safeLayer;
                                        }
                                    }
                                    catch { }
                                }
                            }
                            catch { }

                            tr.Commit();
                        }
                        catch (Exception exInner)
                        {
                            try { tr.Abort(); } catch { }
                            LogManager.Instance.LogWarning($"在临时数据库中创建表时出错: {exInner.Message}");
                            throw;
                        }
                    }

                    try
                    {
                        tempDb.SaveAs(fullPath, Autodesk.AutoCAD.DatabaseServices.DwgVersion.Current);
                    }
                    catch (Exception exSave)
                    {
                        LogManager.Instance.LogError($"保存临时 DWG 时出错: {exSave.Message}");
                        throw;
                    }
                }

                _lastSavedPipeTablePath = fullPath;
                MessageBox.Show($"管道表已生成并保存到临时文件：\n{fullPath}\n\n下一步：切换到目标视口（布局视口请双击进入），点击“插入管道表”并拾取插入点。", "已保存", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"生成管道表时出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 帮助：从实体中提取属性（AttributeReference / ExtensionDictionary Xrecord / RegApp XData）
        /// </summary>
        /// <param name="tr"></param>
        /// <param name="ent"></param>
        /// <returns></returns>
        private Dictionary<string, string> GetEntityAttributeMap(Autodesk.AutoCAD.DatabaseServices.Transaction tr, Autodesk.AutoCAD.DatabaseServices.Entity ent)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (ent == null) return map;

                // 1) 若是块参照，读取 AttributeReference（标签 -> 值）
                if (ent is Autodesk.AutoCAD.DatabaseServices.BlockReference br)
                {
                    try
                    {
                        var attCol = br.AttributeCollection;
                        foreach (ObjectId attId in attCol)
                        {
                            try
                            {
                                var ar = tr.GetObject(attId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.AttributeReference;
                                if (ar != null)
                                {
                                    var tag = (ar.Tag ?? string.Empty).Trim();
                                    var val = (ar.TextString ?? string.Empty).Trim();
                                    if (!string.IsNullOrEmpty(tag))
                                    {
                                        if (!map.ContainsKey(tag)) map[tag] = val;
                                    }
                                }
                            }
                            catch { /* 忽略单个属性读取失败 */ }
                        }
                    }
                    catch { /* 忽略块参照属性读取异常 */ }
                }

                // 2) ExtensionDictionary 中的 Xrecord（键名 -> 以 | 分隔的值）
                try
                {
                    if (ent.ExtensionDictionary != Autodesk.AutoCAD.DatabaseServices.ObjectId.Null)
                    {
                        var extDict = tr.GetObject(ent.ExtensionDictionary, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.DBDictionary;
                        if (extDict != null)
                        {
                            foreach (var entry in extDict)
                            {
                                try
                                {
                                    var xrec = tr.GetObject(entry.Value, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Xrecord;
                                    if (xrec != null)
                                    {
                                        var vals = new List<string>();
                                        if (xrec.Data != null)
                                        {
                                            foreach (Autodesk.AutoCAD.DatabaseServices.TypedValue tv in xrec.Data)
                                            {
                                                if (tv.Value != null) vals.Add(tv.Value.ToString());
                                            }
                                        }
                                        var key = entry.Key ?? string.Empty;
                                        if (!string.IsNullOrWhiteSpace(key))
                                        {
                                            var value = string.Join("|", vals);
                                            if (!map.ContainsKey(key)) map[key] = value;
                                        }
                                    }
                                }
                                catch { /* 忽略单个 Xrecord 读取失败 */ }
                            }
                        }
                    }
                }
                catch { /* 忽略 ExtensionDictionary 读取错误 */ }

                // 3) 遍历注册应用（RegAppTable），读取 XData（AppName::TypedValues）
                try
                {
                    var db = ent.Database;
                    var rat = (Autodesk.AutoCAD.DatabaseServices.RegAppTable)tr.GetObject(db.RegAppTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    foreach (ObjectId appId in rat)
                    {
                        try
                        {
                            var app = tr.GetObject(appId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.RegAppTableRecord;
                            if (app == null) continue;
                            var appName = app.Name;
                            // GetXDataForApplication 返回 ResultBuffer（可能为 null）
                            var rb = ent.GetXDataForApplication(appName);
                            if (rb != null)
                            {
                                var vals = new List<string>();
                                foreach (var tv in rb)
                                {
                                    if (tv.Value != null) vals.Add(tv.Value.ToString());
                                }
                                if (vals.Count > 0)
                                {
                                    var key = $"XDATA:{appName}";
                                    var value = string.Join("|", vals);
                                    if (!map.ContainsKey(key)) map[key] = value;
                                }
                            }
                        }
                        catch { /* 忽略某个 RegApp 读取失败 */ }
                    }
                }
                catch { /* 忽略 RegApp 读取失败 */ }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetEntityAttributeMap 异常: {ex.Message}");
            }
            return map;
        }

        /// <summary>
        /// 帮助：基于已知关键字优先级构建分组键（属性存在时）
        /// </summary>
        /// <param name="attrMap"></param>
        /// <returns></returns>
        private string BuildAttributeGroupKey(Dictionary<string, string> attrMap)
        {
            if (attrMap == null || attrMap.Count == 0) return string.Empty;

            // 关注的属性顺序（优先级）：可以按需扩展
            var priorityKeys = new[]
            {
            "直径", "外径", "内径", "厚度", "规格", "型号",
            "宽度", "高度",
            "材料", "介质",
            "标准号", "标准",
            "功率", "容积", "压力", "温度",
            "材质" // 兼容不同命名
        };

            var parts = new List<string>();
            foreach (var pk in priorityKeys)
            {
                // 找到 attrMap 中包含 pk 的键（不区分大小写，包含匹配）
                var found = attrMap.Keys.FirstOrDefault(k => k.IndexOf(pk, StringComparison.OrdinalIgnoreCase) >= 0);
                if (!string.IsNullOrEmpty(found))
                {
                    var v = (attrMap[found] ?? string.Empty).Trim();
                    parts.Add($"{pk}:{v}");
                }
            }

            // 若没有匹配到优先字段，则把所有属性按键排序并拼接，保证分组稳定性
            if (parts.Count == 0)
            {
                foreach (var kv in attrMap.OrderBy(k => k.Key))
                {
                    parts.Add($"{kv.Key}:{(kv.Value ?? string.Empty).Trim()}");
                }
            }

            return string.Join("|", parts);
        }

        /// <summary>
        /// 帮助：解析属性字符串中的长度值
        /// </summary>
        /// <param name="rawValue"></param>
        /// <returns></returns>
        private double ParseLengthValueFromAttribute(string rawValue)
        {
            // 解析属性字符串中的数值并返回以米为单位的长度。
            // 支持带单位的字符串："2000mm", "2000 mm", "2.5m", "2.5 米", "2500" 等。
            // 规则（启发式）：
            // - 如果字符串中包含 "mm" 或 "毫米" -> 视为毫米，除以1000 返回米。
            // - 如果包含 "m" 或 "米"（且不包含 mm/毫米） -> 视为米。
            // - 否则，若解析出的数值 >= 1000 则假定为毫米并除以1000；否则假定为米。
            if (string.IsNullOrWhiteSpace(rawValue))
                return double.NaN;

            try
            {
                var s = rawValue.Trim();
                var lower = s.ToLowerInvariant();

                bool containsMm = lower.Contains("mm") || lower.Contains("毫米");
                bool containsM = (lower.Contains("m") && !containsMm) || lower.Contains("米");

                // 提取第一个数值（支持小数）
                var m = System.Text.RegularExpressions.Regex.Match(lower, @"[-+]?[0-9]*\.?[0-9]+");
                if (!m.Success)
                    return double.NaN;

                if (!double.TryParse(m.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double value))
                    return double.NaN;

                if (containsMm)
                    return value / 1000.0;

                if (containsM)
                    return value;

                // 无明显单位时根据数值启发式判断：大于等于1000视作mm
                if (value >= 1000.0)
                    return value / 1000.0;

                return value;
            }
            catch
            {
                return double.NaN;
            }
        }

        #endregion


        #region 图层字典
        /// <summary>
        /// 在类成员区域添加 CategoryNames 集合与加载方法
        /// </summary>
        private readonly ObservableCollection<string> _categoryNames = new ObservableCollection<string>();

        /// <summary>
        /// UI 绑定集合：用于绑定到 LayerDictionary_DataGrid.ItemsSource
        /// </summary>
        private ObservableCollection<LayerDictionaryRow> _layerDictionaryRows = new ObservableCollection<LayerDictionaryRow>();

        /// <summary>
        /// 从数据库加载 cad_categories 表的 name 列并填充 CategoryNames
        /// </summary>
        private async Task LoadCategoryNamesAsync()
        {
            try
            {
                if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
                    return;

                var cats = await _databaseManager.GetAllCadCategoriesAsync();
                // 保持 UI 线程安全更新集合
                await Dispatcher.InvokeAsync(() =>
                {
                    _categoryNames.Clear();
                    if (cats == null) return;
                    foreach (var c in cats.OrderBy(x => x.Name ?? x.DisplayName))
                    {
                        // 优先使用 Name 列（题述即为分类），若为空使用 DisplayName 作为回退
                        var name = string.IsNullOrWhiteSpace(c.Name) ? (c.DisplayName ?? string.Empty) : c.Name;
                        if (!string.IsNullOrWhiteSpace(name) && !_categoryNames.Contains(name))
                            _categoryNames.Add(name);
                    }
                });
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"LoadCategoryNamesAsync 失败: {ex.Message}");
            }
        }

        private void 图层字典_Btn_Click(object sender, RoutedEventArgs e)
        {

        }

        private async void 加载标准图层字典_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_databaseManager == null) // 未连接数据库时提示
                {
                    MessageBox.Show("未连接数据库，无法加载图层字典。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string currentUser = VariableDictionary._userName ?? ""; // 当前登录用户名
                // 如果是管理员用户，直接加载该管理员的个人数据（sa/root/admin）
                if (IsAdminUser(currentUser))
                {
                    var list = await _databaseManager.GetLayerDictionaryByUsernameAsync(currentUser).ConfigureAwait(false); // 获取数据
                    Dispatcher.Invoke(() => PopulateLayerDictionaryRowsFromEntries(list)); // 在 UI 线程填充
                    return;
                }

                // 非管理员：尝试按用户所属部门/专业筛选管理员发布的标准字典
                var user = await _databaseManager.GetUserByUsernameAsync(currentUser).ConfigureAwait(false); // 获取用户信息
                string deptName = null; // 用于作为 major 过滤
                if (user != null && user.DepartmentId.HasValue && user.DepartmentId.Value > 0)
                {
                    var depts = await _databaseManager.GetAllDepartmentsAsync().ConfigureAwait(false); // 读取所有部门
                    var dept = depts?.FirstOrDefault(d => d.Id == user.DepartmentId.Value); // 寻找当前用户部门
                    deptName = dept?.Name ?? dept?.DisplayName; // 使用部门名或显示名作为专业标识
                }

                var standards = await _databaseManager.GetStandardLayerDictionaryByMajorAsync(deptName).ConfigureAwait(false); // 获取标准字典
                Dispatcher.Invoke(() => PopulateLayerDictionaryRowsFromEntries(standards)); // 填充到 Grid
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"加载标准图层字典 出错: {ex.Message}"); // 日志记录
                MessageBox.Show($"加载标准图层字典失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); // 提示用户
            }
        }

        private async void 加载个人图层字典_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_databaseManager == null) // 检查 DB
                {
                    MessageBox.Show("未连接数据库，无法加载图层字典。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                string currentUser = VariableDictionary._userName ?? ""; // 当前用户名
                var list = await _databaseManager.GetLayerDictionaryByUsernameAsync(currentUser).ConfigureAwait(false); // 查询个人字典
                Dispatcher.Invoke(() => PopulateLayerDictionaryRowsFromEntries(list)); // 更新 UI
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"加载个人图层字典 出错: {ex.Message}"); // 日志
                MessageBox.Show($"加载个人图层字典失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); // 提示
            }
        }

        private void 当前图纸图层_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_layerManager == null) // 若未实现图层管理器，提示
                {
                    MessageBox.Show("未找到图层管理器，无法读取当前图层。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var layers = _layerManager.LoadCurrentLayers(); // 假定返回 List<LayerInfo>（包含 Name 属性）
                if (layers == null || layers.Count == 0)
                {
                    MessageBox.Show("当前图纸未检测到图层。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // 按名称排序、去重后加入 DataGrid
                var ordered = layers.Select(l => l.LayerName).Where(n => !string.IsNullOrEmpty(n)).Distinct().OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();

                foreach (var name in ordered)
                {
                    _layerDictionaryRows.Add(new LayerDictionaryRow
                    {
                        Id = 0, // 新行
                        Major = "", // 用户可手动填写专业
                        LayerName = name, // 填充原图层名
                        DicLayerName = "" // 解释名留空，供用户编辑
                    });
                }

                // 重新编号从 1 开始
                ReindexLayerDictionaryRows();

                if (LayerDictionary_DataGrid.ItemsSource == null) // 绑定检查
                    LayerDictionary_DataGrid.ItemsSource = _layerDictionaryRows;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"当前图纸图层 加载失败: {ex.Message}"); // 记录日志
                MessageBox.Show($"加载当前图纸图层失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); // 提示
            }
        }

        private async void 添加一行_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 确保 CategoryNames 已加载，以便新行的专业下拉能立即显示内容
                if ((_categoryNames == null || _categoryNames.Count == 0) && _databaseManager != null && _databaseManager.IsDatabaseAvailable)
                {
                    await LoadCategoryNamesAsync();
                    // 如果 DataGrid 的 ComboColumn 是通过代码设置的，确保重新应用一次（防止先前为空）
                    try
                    {
                        var comboCol = LayerDictionary_DataGrid.Columns
                            .OfType<DataGridComboBoxColumn>()
                            .FirstOrDefault(c => (c.Header?.ToString() ?? string.Empty).IndexOf("专业", StringComparison.OrdinalIgnoreCase) >= 0);
                        if (comboCol != null)
                            comboCol.ItemsSource = _categoryNames;
                    }
                    catch { /* 忽略 */ }
                }

                var newRow = new LayerDictionaryRow
                {
                    Id = 0, // 新行未持久化
                    Major = _categoryNames.FirstOrDefault() ?? "", // 默认选中第一个分类（若存在），提升 UX
                    LayerName = "", // 默认空
                    DicLayerName = "" // 默认空
                };
                _layerDictionaryRows.Add(newRow); // 添加到集合，UI 自动更新

                // 重新生成并刷新序号（从 1 开始）
                ReindexLayerDictionaryRows();

                // 如果 ItemsSource 尚未绑定，则重新绑定（通常已在初始化时绑定）
                if (LayerDictionary_DataGrid.ItemsSource == null)
                    LayerDictionary_DataGrid.ItemsSource = _layerDictionaryRows;

                // 将新行滚动至可见并启用编辑第一可编辑单元格（Major）
                try
                {
                    var rowIndex = _layerDictionaryRows.IndexOf(newRow);
                    if (rowIndex >= 0)
                    {
                        LayerDictionary_DataGrid.ScrollIntoView(newRow);
                        LayerDictionary_DataGrid.UpdateLayout();
                        var row = (DataGridRow)LayerDictionary_DataGrid.ItemContainerGenerator.ContainerFromIndex(rowIndex);
                        if (row != null)
                        {
                            LayerDictionary_DataGrid.SelectedItem = newRow;
                            LayerDictionary_DataGrid.CurrentCell = new DataGridCellInfo(newRow, LayerDictionary_DataGrid.Columns[1]);
                            LayerDictionary_DataGrid.BeginEdit();
                        }
                    }
                }
                catch { /* 忽略编辑辅助失败 */ }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"添加一行 出错: {ex.Message}"); // 记录异常
            }
        }

        private async void 删除选中行_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selected = LayerDictionary_DataGrid.SelectedItems.Cast<LayerDictionaryRow>().ToList(); // 获取选中项
                if (selected == null || selected.Count == 0)
                {
                    MessageBox.Show("请先选择要删除的行。", "提示", MessageBoxButton.OK, MessageBoxImage.Information); // 无选中时提示
                    return;
                }

                // 从 UI 集合中移除选中行
                foreach (var s in selected)
                {
                    _layerDictionaryRows.Remove(s);
                }

                // 重新编号使序号从1开始连续
                ReindexLayerDictionaryRows();

                // 若选中行已经存在数据库 id，则删除数据库中对应记录
                var idsToDelete = selected.Where(x => x.Id > 0).Select(x => x.Id).ToList();
                if (idsToDelete.Count > 0 && _databaseManager != null)
                {
                    await _databaseManager.DeleteLayerDictionaryEntriesAsync(idsToDelete).ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"删除选中行 出错: {ex.Message}"); // 日志
                MessageBox.Show($"删除选中行失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); // 提示
            }
        }

        private async void 保存图层字典_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                if (_databaseManager == null) // 检查 DB
                {
                    MessageBox.Show("未连接数据库，无法保存图层字典。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string username = VariableDictionary._userName ?? ""; // 当前用户名
                var rows = _layerDictionaryRows.ToList(); // 从 ObservableCollection 获取快照

                // 将 UI 行转换为数据库实体 LayerDictionaryEntry
                // 说明：LayerDictionaryEntry 使用 Mappings（List<LayerMapping>） / MappingsJson 存储任意数量的映射对
                var entries = new List<LayerDictionaryHelper>();
                int seq = 1;

                // 逐行保存：每行生成一条数据库记录，记录中 mappings 包含单个映射对
                foreach (var r in rows)
                {
                    // 若行中既没有原图层也没有解释名，则跳过（避免写入空记录）
                    if (string.IsNullOrWhiteSpace(r.LayerName) && string.IsNullOrWhiteSpace(r.DicLayerName))
                        continue;

                    var entry = new LayerDictionaryHelper
                    {
                        Seq = seq++, // 自动分配序号
                        Major = r.Major, // 专业字段
                        Username = username, // 所属用户
                        UserId = null, // 可选，不填
                        Source = r.Source ?? "personal", // 来源
                        CreatedBy = username // 创建者记录
                    };

                    // 使用运行时列表 Mappings，LayerDictionaryEntry 的 setter 会序列化为 MappingsJson
                    entry.Mappings = new List<LayerMapping>
                    {
                        new LayerMapping
                        {
                            OriginalLayer = r.LayerName ?? string.Empty,
                            DicLayer = r.DicLayerName ?? string.Empty
                        }
                    };

                    entries.Add(entry); // 加入待保存列表
                }

                if (entries.Count == 0)
                {
                    Dispatcher.Invoke(() => MessageBox.Show("没有可保存的映射行，请先添加或填写映射。", "提示", MessageBoxButton.OK, MessageBoxImage.Information));
                    return;
                }

                // 调用数据库扩展方法保存（覆盖当前用户所有记录）
                var ok = await _databaseManager.SaveLayerDictionaryForUserAsync(username, entries).ConfigureAwait(false);
                if (ok)
                {
                    Dispatcher.Invoke(() => MessageBox.Show("保存图层字典成功。", "完成", MessageBoxButton.OK, MessageBoxImage.Information)); // 成功提示
                }
                else
                {
                    Dispatcher.Invoke(() => MessageBox.Show("保存图层字典失败，请查看日志。", "失败", MessageBoxButton.OK, MessageBoxImage.Error)); // 失败提示
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"保存图层字典 出错: {ex.Message}"); // 记录日志
                MessageBox.Show($"保存图层字典失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); // 提示用户
            }
        }

        /// <summary>
        /// 将数据库实体列表转换并填充到 _layerDictionaryRows（用于显示）
        /// </summary>
        private void PopulateLayerDictionaryRowsFromEntries(List<LayerDictionaryHelper> entries)
        {
            _layerDictionaryRows.Clear(); // 先清空集合
            if (entries == null || entries.Count == 0)
            {
                // 确保 DataGrid 刷新显示空集合
                LayerDictionary_DataGrid.ItemsSource = _layerDictionaryRows;
                LayerDictionary_DataGrid.Items.Refresh();
                return; // 若无数据则返回
            }

            // 仍按原逻辑展平 Mappings -> 多行
            foreach (var e in entries)
            {
                var mappings = e.Mappings ?? new List<LayerMapping>();

                if ((mappings == null || mappings.Count == 0) && e != null)
                {
                    string GetPropAsString(object obj, string propName)
                    {
                        try
                        {
                            var pi = obj.GetType().GetProperty(propName);
                            if (pi == null) return string.Empty;
                            var v = pi.GetValue(obj);
                            return v?.ToString() ?? string.Empty;
                        }
                        catch
                        {
                            return string.Empty;
                        }
                    }

                    var o1 = GetPropAsString(e, "OriginalLayer1");
                    var d1 = GetPropAsString(e, "DicLayer1");
                    if (!string.IsNullOrEmpty(o1) || !string.IsNullOrEmpty(d1))
                    {
                        mappings = new List<LayerMapping> { new LayerMapping { OriginalLayer = o1, DicLayer = d1 } };
                    }
                    else
                    {
                        var o2 = GetPropAsString(e, "Original");
                        var d2 = GetPropAsString(e, "Dic");
                        if (!string.IsNullOrEmpty(o2) || !string.IsNullOrEmpty(d2))
                            mappings = new List<LayerMapping> { new LayerMapping { OriginalLayer = o2, DicLayer = d2 } };
                    }
                }

                if (mappings == null || mappings.Count == 0)
                {
                    _layerDictionaryRows.Add(new LayerDictionaryRow
                    {
                        Id = e.Id,
                        Major = e.Major ?? string.Empty,
                        LayerName = string.Empty,
                        DicLayerName = string.Empty,
                        Source = e.Source ?? "personal"
                    });
                    continue;
                }

                foreach (var m in mappings)
                {
                    _layerDictionaryRows.Add(new LayerDictionaryRow
                    {
                        Id = e.Id,
                        Major = e.Major ?? string.Empty,
                        LayerName = m?.OriginalLayer ?? string.Empty,
                        DicLayerName = m?.DicLayer ?? string.Empty,
                        Source = e.Source ?? "personal"
                    });
                }
            }

            // 绑定并重新编号（从1开始）
            LayerDictionary_DataGrid.ItemsSource = _layerDictionaryRows;
            ReindexLayerDictionaryRows();
        }

        /// <summary>
        /// 判断是否为管理员用户（sa/root/admin）
        /// </summary>
        private static bool IsAdminUser(string username)
        {
            if (string.IsNullOrWhiteSpace(username)) return false; // 空用户名不是管理员
            var low = username.Trim().ToLowerInvariant(); // 规范为小写比较
            return low == "sa" || low == "root" || low == "admin"; // 三个默认管理员用户名
        }

        /// <summary>
        /// 重新生成 _layerDictionaryRows 的显示序号，从 1 开始连续编号并刷新 DataGrid 显示
        /// 在添加/删除/填充后都应调用此方法以保证序号连续且从 1 开始
        /// </summary>
        private void ReindexLayerDictionaryRows()
        {
            try
            {
                if (_layerDictionaryRows == null) return;
                int idx = 1;
                foreach (var row in _layerDictionaryRows)
                {
                    row.DisplayIndex = idx++;
                }

                // 确保 DataGrid 绑定并刷新显示
                if (LayerDictionary_DataGrid.ItemsSource == null)
                    LayerDictionary_DataGrid.ItemsSource = _layerDictionaryRows;

                LayerDictionary_DataGrid.Items.Refresh();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"ReindexLayerDictionaryRows 出错: {ex.Message}");
            }
        }

        /// <summary>
        /// PreparingCellForEdit：当开始编辑单元格时，如果是 ComboBoxColumn，查找 ComboBox 并订阅 SelectionChanged
        /// </summary>
        private void LayerDictionary_DataGrid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            try
            {
                if (e.Column is DataGridComboBoxColumn)
                {
                    // EditingElement 有时是 ComboBox 或 ContentPresenter，尝试直接转换或在视觉树中查找
                    ComboBox combo = e.EditingElement as ComboBox ?? FindVisualChildByType<ComboBox>(e.EditingElement);
                    if (combo != null)
                    {
                        // 避免重复订阅
                        combo.SelectionChanged -= LayerDictionaryComboBox_SelectionChanged;
                        combo.SelectionChanged += LayerDictionaryComboBox_SelectionChanged;
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"PreparingCellForEdit 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// CellEditEnding：编辑结束时从编辑元素上解绑事件（防止内存泄漏或重复处理）
        /// </summary>
        private void LayerDictionary_DataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            try
            {
                if (!(e.Column is DataGridComboBoxColumn)) return;

                // 尝试从编辑元素读取新值（编辑元素通常是 ComboBox 或 ContentPresenter）
                object newValue = null;
                ComboBox combo = e.EditingElement as ComboBox ?? FindVisualChildByType<ComboBox>(e.EditingElement);
                if (combo != null)
                {
                    newValue = combo.SelectedItem ?? combo.SelectedValue ?? combo.Text;
                }
                else
                {
                    // 兜底：从行绑定对象读取（如果绑定尚未更新，这里可能仍是旧值，但我们还是尝试）
                    if (e.Row?.Item is LayerDictionaryRow row)
                        newValue = row.Major;
                }

                if (newValue == null) return;

                // 复制选中项列表，避免在遍历时集合改变
                var selectedItems = LayerDictionary_DataGrid.SelectedItems.Cast<object>().ToList();

                // 延迟应用，确保 DataGrid 自身完成数据绑定/提交
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        var text = newValue.ToString();
                        bool changed = false;
                        foreach (var it in selectedItems)
                        {
                            if (it is LayerDictionaryRow r)
                            {
                                if (!string.Equals(r.Major ?? string.Empty, text, StringComparison.Ordinal))
                                {
                                    r.Major = text;
                                    changed = true;
                                }
                            }
                        }

                        if (changed)
                        {
                            // 刷新 DataGrid 显示并结束任何编辑状态
                            LayerDictionary_DataGrid.CommitEdit(DataGridEditingUnit.Row, true);
                            LayerDictionary_DataGrid.Items.Refresh();
                        }
                    }
                    catch (Exception exInner)
                    {
                        LogManager.Instance.LogInfo($"批量更新 Major 失败: {exInner.Message}");
                    }
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"CellEditEnding 处理失败: {ex.Message}");
            }
            finally
            {
                // 尝试解绑编辑元素上的事件，避免重复订阅（防御性）
                try
                {
                    if (e.EditingElement != null)
                    {
                        var cb = e.EditingElement as ComboBox ?? FindVisualChildByType<ComboBox>(e.EditingElement);
                        if (cb != null)
                            cb.SelectionChanged -= LayerDictionaryComboBox_SelectionChanged;
                    }
                }
                catch { /* 忽略解绑异常 */ }
            }
        }

        /// <summary>
        /// ComboBox SelectionChanged：将所选值应用到所有当前选中的行（批量修改 Major）
        /// </summary>
        private void LayerDictionaryComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var combo = sender as ComboBox;
                if (combo == null) return;

                // 获取选择值（SelectedItem 优先）
                var selectedValue = combo.SelectedItem ?? combo.SelectedValue;
                if (selectedValue == null) return;

                // 如果没有选中任何行，则只更新当前编辑行（保留默认行为）
                if (LayerDictionary_DataGrid.SelectedItems == null || LayerDictionary_DataGrid.SelectedItems.Count == 0)
                {
                    // 若编辑单元格的 DataContext 可用，则更新该行
                    var rowContext = combo.DataContext as LayerDictionaryRow;
                    if (rowContext != null)
                    {
                        rowContext.Major = selectedValue.ToString();
                        LayerDictionary_DataGrid.Items.Refresh();
                    }
                    return;
                }

                // 将选择值写入所有被选中的行（批量修改）
                var changed = false;
                foreach (var it in LayerDictionary_DataGrid.SelectedItems)
                {
                    if (it is LayerDictionaryRow row)
                    {
                        // 仅当值不同才赋值
                        var newVal = selectedValue.ToString();
                        if (!string.Equals(row.Major ?? string.Empty, newVal, StringComparison.Ordinal))
                        {
                            row.Major = newVal;
                            changed = true;
                        }
                    }
                }

                if (changed)
                {
                    // 刷新显示
                    LayerDictionary_DataGrid.Items.Refresh();
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"LayerDictionaryComboBox_SelectionChanged 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 通用 VisualTree 查找指定类型子元素（递归）
        /// </summary>
        private T FindVisualChildByType<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;
            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typed) return typed;
                var result = FindVisualChildByType<T>(child);
                if (result != null) return result;
            }
            return null;
        }

        /// <summary>
        /// 构建图层字典映射表（优先级：UI 编辑的行 > 个人字典 > 标准字典）
        /// 返回字典：Original -> Dic（不区分大小写）
        /// 设计要点：
        /// - 不在关键点使用 ConfigureAwait(false)；
        /// - 只用 Dispatcher 在开始时同步读取 UI （_layerDictionaryRows），避免在后续数据库读取时触发跨线程问题；
        /// - 该方法可以在进入 AutoCAD document lock 之前安全 await 执行（不会访问 AutoCAD API）。
        /// </summary>
        private async Task<Dictionary<string, string>> BuildLayerDictionaryMapAsync()
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // 1) 同步从 UI 读取当前表格（在 UI 线程上执行）
                try
                {
                    // 使用同步 Dispatcher.Invoke，确保立即获得最新 UI 数据（不会产生 await 导致线程切换）
                    Dispatcher.Invoke(() =>
                    {
                        if (_layerDictionaryRows != null)
                        {
                            foreach (var row in _layerDictionaryRows)
                            {
                                if (row == null) continue;
                                var orig = (row.LayerName ?? string.Empty).Trim();
                                var dic = (row.DicLayerName ?? string.Empty).Trim();
                                if (string.IsNullOrEmpty(orig)) continue;
                                // UI 编辑优先，直接写入/覆盖映射
                                map[orig] = string.IsNullOrEmpty(dic) ? orig : dic;
                            }
                        }
                    }, System.Windows.Threading.DispatcherPriority.Send);
                }
                catch (Exception uiEx)
                {
                    LogManager.Instance.LogInfo($"BuildLayerDictionaryMapAsync: 读取 UI 数据失败: {uiEx.Message}");
                    // 继续尝试从数据库加载
                }

                // 如果没有可用的数据库，直接返回基于 UI 的 map
                if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
                    return map;

                // 当前用户名（可能为空）
                string username = VariableDictionary._userName ?? string.Empty;

                // 2) 个人字典（覆盖标准字典）
                try
                {
                    // 注意：这里不要使用 ConfigureAwait(false)，需要保持异常与上下文的自然传播与记录。
                    var personal = await _databaseManager.GetLayerDictionaryByUsernameAsync(username);
                    if (personal != null)
                    {
                        foreach (var entry in personal)
                        {
                            var mappings = entry.Mappings ?? new List<LayerMapping>();
                            // 兼容旧结构：若 mappings 为空，可以尝试从其它属性回退解析（省略复杂回退）
                            foreach (var m in mappings)
                            {
                                if (m == null) continue;
                                var origLayer = (m.OriginalLayer ?? string.Empty).Trim();
                                var dicLayer = (m.DicLayer ?? string.Empty).Trim();
                                if (string.IsNullOrEmpty(origLayer)) continue;
                                // 个人映射优先，直接覆盖
                                map[origLayer] = string.IsNullOrEmpty(dicLayer) ? origLayer : dicLayer;
                            }
                        }
                    }
                }
                catch (Exception exPersonal)
                {
                    LogManager.Instance.LogInfo($"BuildLayerDictionaryMapAsync: 读取个人图层字典失败: {exPersonal.Message}");
                }

                // 3) 标准字典（仅当键不存在时写入，不覆盖 UI/个人）
                try
                {
                    string deptName = null;
                    try
                    {
                        var user = await _databaseManager.GetUserByUsernameAsync(username);
                        if (user != null && user.DepartmentId.HasValue && user.DepartmentId.Value > 0)
                        {
                            var depts = await _databaseManager.GetAllDepartmentsAsync();
                            var dept = depts?.FirstOrDefault(d => d.Id == user.DepartmentId.Value);
                            deptName = dept?.Name ?? dept?.DisplayName;
                        }
                    }
                    catch
                    {
                        // 忽略部门解析错误，继续使用 null
                    }

                    var standards = await _databaseManager.GetStandardLayerDictionaryByMajorAsync(deptName);
                    if (standards != null)
                    {
                        foreach (var entry in standards)
                        {
                            var mappings = entry.Mappings ?? new List<LayerMapping>();
                            foreach (var m in mappings)
                            {
                                if (m == null) continue;
                                var origLayer = (m.OriginalLayer ?? string.Empty).Trim();
                                var dicLayer = (m.DicLayer ?? string.Empty).Trim();
                                if (string.IsNullOrEmpty(origLayer)) continue;
                                // 仅在不存在时写入，避免覆盖 UI/个人映射
                                if (!map.ContainsKey(origLayer))
                                    map[origLayer] = string.IsNullOrEmpty(dicLayer) ? origLayer : dicLayer;
                            }
                        }
                    }
                }
                catch (Exception exStd)
                {
                    LogManager.Instance.LogInfo($"BuildLayerDictionaryMapAsync: 读取标准图层字典失败: {exStd.Message}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"BuildLayerDictionaryMapAsync 异常: {ex.Message}");
            }

            return map;
        }

        #endregion

        #region 生成表\插入表格相关
        private async void 生成暖通管道表_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("未找到活动的 AutoCAD 文档。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 在进入锁前准备图层映射（异步完成）
                Dictionary<string, string> layerMap;
                try
                {
                    layerMap = await BuildLayerDictionaryMapAsync();
                    if (layerMap == null) layerMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                }
                catch (Exception exMap)
                {
                    LogManager.Instance.LogWarning($"构建图层映射失败（继续执行）：{exMap.Message}");
                    layerMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                }

                using (doc.LockDocument())
                {
                    var ed = doc.Editor;

                    var hint = "\n开始生成管道表：\n1) 在视口内选择管道图元（选择完成后按 Enter/空格或右键结束选择）";
                    ed.WriteMessage(hint);

                    var pso = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions
                    {
                        MessageForAdding = "\n请选择要统计的管道图元：",
                        AllowDuplicates = false,
                        RejectObjectsFromNonCurrentSpace = false
                    };
                    var psr = ed.GetSelection(pso);
                    if (psr.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK || psr.Value == null)
                    {
                        ed.WriteMessage("\n未选择实体或已取消。");
                        return;
                    }

                    var selIds = psr.Value.GetObjectIds();
                    if (selIds == null || selIds.Length == 0)
                    {
                        MessageBox.Show("未选择任何实体。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    double unitToMeters = 1000.0;
                    var pdo = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n请输入图形单位到米的换算分母（例如：图形单位为毫米请输入1000，回车使用默认1000）：")
                    {
                        AllowNone = true,
                        DefaultValue = unitToMeters,
                        UseDefaultValue = true,
                        AllowZero = false,
                        AllowNegative = false
                    };
                    var pdr = ed.GetDouble(pdo);
                    if (pdr.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK && pdr.Value > 0.0)
                        unitToMeters = pdr.Value;

                    // 先在事务中遍历实体、提取属性并收集要统计的字段
                    var globalAttributeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var summariesMap = new Dictionary<string, PipeSummary>(StringComparer.OrdinalIgnoreCase);
                    bool anyAttributesFound = false;

                    using (var tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        foreach (var id in selIds)
                        {
                            try
                            {
                                var ent = tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                                if (ent == null) continue;

                                dynamic ext;
                                try
                                {
                                    ext = ent.GeometricExtents;
                                }
                                catch
                                {
                                    continue;
                                }

                                double sizeX = Math.Abs((double)ext.MaxPoint.X - (double)ext.MinPoint.X);
                                double sizeY = Math.Abs((double)ext.MaxPoint.Y - (double)ext.MinPoint.Y);
                                double sizeZ = Math.Abs((double)ext.MaxPoint.Z - (double)ext.MinPoint.Z);
                                if (sizeX <= 0.0) continue;

                                double sizeY_m = sizeY / unitToMeters;
                                double sizeZ_m = sizeZ / unitToMeters;
                                int roundedY = (int)Math.Round(sizeY_m * 1000.0);
                                int roundedZ = (int)Math.Round(sizeZ_m * 1000.0);

                                // 原始图层名
                                string originalLayer = string.IsNullOrEmpty(ent.Layer) ? "无图层" : ent.Layer;
                                string mappedLayer = originalLayer;
                                if (layerMap != null && layerMap.TryGetValue(originalLayer.Trim(), out var val) && !string.IsNullOrWhiteSpace(val))
                                    mappedLayer = val;
                                else if (layerMap != null && layerMap.Count > 0)
                                {
                                    foreach (var kv in layerMap)
                                    {
                                        if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                                        if (originalLayer.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            mappedLayer = string.IsNullOrWhiteSpace(kv.Value) ? kv.Key : kv.Value;
                                            break;
                                        }
                                    }
                                }

                                // 读取实体属性
                                var attrMap = GetEntityAttributeMap(tr, ent);
                                if (attrMap != null && attrMap.Count > 0)
                                {
                                    anyAttributesFound = true;
                                    foreach (var k in attrMap.Keys) globalAttributeKeys.Add(k);
                                }

                                // 生成分组键：如果有属性则按属性组合，否则按尺寸组合（保留原逻辑）
                                string specKey;
                                if (anyAttributesFound && attrMap != null && attrMap.Count > 0)
                                {
                                    // 使用属性组合键
                                    specKey = BuildAttributeGroupKey(attrMap);
                                }
                                else
                                {
                                    specKey = $"{roundedY}_{roundedZ}";
                                }

                                // 计算实体长度：优先使用属性中包含 "长度" 的字段（若能解析出有效数值），否则使用几何尺寸
                                double length_m = double.NaN;
                                bool usedAttrLength = false;
                                if (attrMap != null && attrMap.Count > 0)
                                {
                                    // 查找属性键中包含 "长度" 的项（不区分大小写）
                                    foreach (var key in attrMap.Keys)
                                    {
                                        if (key != null && key.IndexOf("长度", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            var raw = attrMap[key];
                                            var parsed = ParseLengthValueFromAttribute(raw);
                                            if (!double.IsNaN(parsed) && parsed > 0.0)
                                            {
                                                length_m = parsed; // 已是米单位（Parse 方法保证）
                                                usedAttrLength = true;
                                                break;
                                            }
                                        }
                                    }
                                }

                                if (!usedAttrLength)
                                {
                                    // 回退到几何长度（以 X 方向为长度）
                                    length_m = sizeX / unitToMeters;
                                }

                                if (!summariesMap.TryGetValue(specKey, out PipeSummary summary))
                                {
                                    summary = new PipeSummary
                                    {
                                        Seq = 0,
                                        Name = mappedLayer,
                                        WidthSpec = roundedY,
                                        ThicknessSpec = roundedZ,
                                        QuantityMeters = 0.0,
                                        Remark = "",
                                        SpecString = anyAttributesFound ? specKey : $"{roundedY} X {roundedZ}"
                                    };
                                    // 若实体有属性则保存属性字典（优先权：属性键原样保存）
                                    if (attrMap != null && attrMap.Count > 0)
                                    {
                                        foreach (var kv in attrMap)
                                        {
                                            if (!summary.Attributes.ContainsKey(kv.Key))
                                                summary.Attributes[kv.Key] = kv.Value;
                                        }
                                    }
                                    summariesMap[specKey] = summary;
                                }

                                // 累加长度（若属性提供长度则使用属性长度，否则使用几何长度）
                                if (!double.IsNaN(length_m) && length_m > 0.0)
                                {
                                    summary.QuantityMeters += length_m;
                                }
                            }
                            catch (Exception exEnt)
                            {
                                LogManager.Instance.LogWarning($"处理实体 {id} 时出错: {exEnt.Message}");
                            }
                        }

                        tr.Commit();
                    }

                    if (summariesMap.Count == 0)
                    {
                        MessageBox.Show("所选实体未能提取到尺寸信息，无法生成表格。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    var summaries = summariesMap.Values
                        .OrderBy(s => s.WidthSpec)
                        .ThenBy(s => s.ThicknessSpec)
                        .ToList();

                    for (int i = 0; i < summaries.Count; i++) summaries[i].Seq = i + 1;

                    // 属性列列表（按稳定顺序：优先常见字段，其余按字典序）
                    List<string> attributeColumns = null;
                    if (anyAttributesFound && globalAttributeKeys.Count > 0)
                    {
                        var priority = new[] { "直径", "外径", "内径", "厚度", "规格", "型号", "宽度", "高度", "材料", "介质", "标准号", "功率", "容积", "压力", "温度", "材质" };
                        attributeColumns = new List<string>();
                        foreach (var p in priority)
                        {
                            var match = globalAttributeKeys.FirstOrDefault(k => k.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0);
                            if (match != null)
                            {
                                attributeColumns.Add(match);
                            }
                        }
                        // 添加剩余未包含的属性（按字母顺序）
                        var remaining = globalAttributeKeys.Except(attributeColumns, StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase);
                        attributeColumns.AddRange(remaining);
                    }

                    // 保存已映射的图层列表
                    _lastSavedPipeTableLayers = summaries
                        .Select(s => s.Name ?? "无图层")
                        .Where(n => !string.IsNullOrWhiteSpace(n))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    var promptMat = new Autodesk.AutoCAD.EditorInput.PromptStringOptions("\n请输入材料（将作为备注默认值，可回车留空）：")
                    {
                        AllowSpaces = true,
                        DefaultValue = ""
                    };
                    var presMat = ed.GetString(promptMat);
                    string material = (presMat.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK) ? presMat.StringResult : string.Empty;
                    foreach (var s in summaries) s.Remark = string.IsNullOrEmpty(material) ? "请填写材料" : material;

                    // 调用保存（支持动态属性列）
                    SavePipeTableToTempDwg(summaries, material, 1.0, attributeColumns);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"生成管道表失败: {ex.Message}");
                MessageBox.Show($"生成管道表时发生错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void 生成暖通设备表_Click(object sender, RoutedEventArgs e)
        {

        }

        private void 插入暖通管道表_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_lastSavedPipeTablePath) || !File.Exists(_lastSavedPipeTablePath))
                {
                    MessageBox.Show("未找到已保存的临时表文件，请先点击【生成管道表】并完成保存。", "未找到文件", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("未找到活动的 AutoCAD 文档。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                using (doc.LockDocument())
                {
                    var ed = doc.Editor;

                    // 提示用户切换视口并确认
                    var pko = new Autodesk.AutoCAD.EditorInput.PromptKeywordOptions("\n请将鼠标移动到目标视口（布局请先双击进入视口），然后选择：\n[继续] 继续 / [取消] 取消")
                    {
                        AllowNone = true
                    };
                    pko.Keywords.Add("继续");
                    pko.Keywords.Add("取消");
                    try { pko.Keywords.Default = "继续"; } catch { }

                    var pkr = ed.GetKeywords(pko);
                    if (pkr.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK || string.Equals(pkr.StringResult, "取消", StringComparison.OrdinalIgnoreCase))
                        return;

                    // 拾取插入点
                    var ppo = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\n请选择表格插入点（请在目标视口内拾取）：");
                    var ppr = ed.GetPoint(ppo);
                    if (ppr.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                    {
                        MessageBox.Show("未选择插入点，操作已取消。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    // 如果是在布局中插入，提示用户输入视口比例分母以便缩放
                    double insertionScale = 1.0;
                    var scalePrompt = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n如果在布局/纸空间插入，请输入视口比例的分母（例如 100 表示 1:100），回车表示 1：")
                    {
                        AllowNone = true,
                        DefaultValue = 1.0,
                        UseDefaultValue = true,
                        AllowZero = false,
                        AllowNegative = false
                    };
                    var scaleRes = ed.GetDouble(scalePrompt);
                    if (scaleRes.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK && scaleRes.Value > 0.0)
                        insertionScale = 1.0 / scaleRes.Value; // 实际缩放因子

                    // 插入块（返回 BlockReference 的 ObjectId）
                    var insertedBrId = AutoCadHelper.InsertBlockFromExternalDwg(_lastSavedPipeTablePath, Path.GetFileNameWithoutExtension(_lastSavedPipeTablePath), ppr.Value);

                    if (insertedBrId == Autodesk.AutoCAD.DatabaseServices.ObjectId.Null)
                    {
                        MessageBox.Show("插入失败：未能将临时 DWG 中的块导入并插入。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // 如果需要缩放（例如布局视口按比例显示），调整已插入的 BlockReference 的 ScaleFactors
                    if (Math.Abs(insertionScale - 1.0) > 1e-9)
                    {
                        try
                        {
                            using (var tr = doc.Database.TransactionManager.StartTransaction())
                            {
                                var br = tr.GetObject(insertedBrId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.BlockReference;
                                if (br != null)
                                {
                                    br.ScaleFactors = new Autodesk.AutoCAD.Geometry.Scale3d(insertionScale, insertionScale, insertionScale);
                                }
                                tr.Commit();
                            }
                        }
                        catch (Exception exScale)
                        {
                            LogManager.Instance.LogWarning($"调整插入比例时出错: {exScale.Message}");
                        }
                    }

                    // 新增：将插入的表设置到“选定图元图层”中，并确保块定义内实体也使用该图层
                    try
                    {
                        string targetLayer = null;
                        if (_lastSavedPipeTableLayers != null && _lastSavedPipeTableLayers.Count > 0)
                        {
                            if (_lastSavedPipeTableLayers.Count == 1)
                            {
                                targetLayer = _lastSavedPipeTableLayers[0];
                            }
                            else
                            {
                                // 列出可选图层并让用户通过索引选择，避免输入图层名错误
                                var listText = new StringBuilder();
                                for (int i = 0; i < _lastSavedPipeTableLayers.Count; i++)
                                {
                                    listText.AppendLine($"{i + 1}. {_lastSavedPipeTableLayers[i]}");
                                }
                                ed.WriteMessage($"\n检测到生成表时涉及以下图层：\n{listText}\n请输入要用于表的图层序号（1 - {_lastSavedPipeTableLayers.Count}），回车使用默认第1项：");

                                var pio = new Autodesk.AutoCAD.EditorInput.PromptIntegerOptions("\n请输入序号：")
                                {
                                    AllowNone = true,
                                    DefaultValue = 1,
                                    LowerLimit = 1,
                                    UpperLimit = _lastSavedPipeTableLayers.Count,
                                    UseDefaultValue = true
                                };
                                var pir = ed.GetInteger(pio);
                                if (pir.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    int idx = Math.Max(1, Math.Min(_lastSavedPipeTableLayers.Count, pir.Value));
                                    targetLayer = _lastSavedPipeTableLayers[idx - 1];
                                }
                                else
                                {
                                    targetLayer = _lastSavedPipeTableLayers[0];
                                }
                            }
                        }

                        if (!string.IsNullOrEmpty(targetLayer))
                        {
                            using (var tr = doc.Database.TransactionManager.StartTransaction())
                            {
                                // 确保目标图层存在，否则创建
                                var lt = (Autodesk.AutoCAD.DatabaseServices.LayerTable)tr.GetObject(doc.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                if (!lt.Has(targetLayer))
                                {
                                    lt.UpgradeOpen();
                                    var ltr = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord
                                    {
                                        Name = targetLayer
                                    };
                                    lt.Add(ltr);
                                    tr.AddNewlyCreatedDBObject(ltr, true);
                                }

                                // 设置 BlockReference 的图层，并把块定义中所有实体的 Layer 设置为目标图层
                                var br = tr.GetObject(insertedBrId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.BlockReference;
                                if (br != null)
                                {
                                    // 设置 BlockReference 层
                                    br.Layer = targetLayer;

                                    // 获取块定义（BlockTableRecord）并把其子实体层设置为目标图层
                                    var btrId = br.BlockTableRecord;
                                    if (btrId != Autodesk.AutoCAD.DatabaseServices.ObjectId.Null)
                                    {
                                        var btr = tr.GetObject(btrId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.BlockTableRecord;
                                        if (btr != null)
                                        {
                                            foreach (ObjectId entId in btr)
                                            {
                                                try
                                                {
                                                    // 可能包含匿名/非实体项，安全转换
                                                    var ent = tr.GetObject(entId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.Entity;
                                                    if (ent != null)
                                                    {
                                                        // 跳过属性定义（AttributeDefinition）以免影响属性行为
                                                        if (ent is Autodesk.AutoCAD.DatabaseServices.AttributeDefinition) continue;

                                                        // 设置实体层为目标图层
                                                        ent.Layer = targetLayer;
                                                    }
                                                }
                                                catch
                                                {
                                                    // 忽略个别实体修改失败
                                                }
                                            }
                                        }
                                    }
                                }

                                tr.Commit();
                            }
                        }
                    }
                    catch (Exception exLayer)
                    {
                        LogManager.Instance.LogWarning($"设置插入表图层时出错: {exLayer.Message}");
                    }

                    MessageBox.Show("临时管道表已插入当前空间。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"插入管道表失败: {ex.Message}");
                MessageBox.Show($"插入管道表失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void 导出暖通表格_Click(object sender, RoutedEventArgs e)
        {

        }
        #endregion


        /// <summary>
        /// 分类属性数据网格选择项改变时触发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void CategoryPropertiesDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        #region 部门与人员管理相关代码
        /// <summary>
        /// 加载完成
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DepartmentAdminControl_Loaded(object sender, RoutedEventArgs e)
        {
            // 从 login 配置优先读取服务器/端口（与之前主界面约定一致）
            try
            {
                /// 从 login 配置优先读取服务器/端口（与之前主界面约定一致）
                var cfgPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GB_NewCadPlus_LM", "login_config.json");
                string host = "127.0.0.1";
                string port = "3306";
                if (System.IO.File.Exists(cfgPath))
                {
                    var json = System.IO.File.ReadAllText(cfgPath);//读取配置文件读取配置文件
                    var ser = new System.Web.Script.Serialization.JavaScriptSerializer();//创建JSON序列化器 创建序列化器
                    var cfg = ser.Deserialize<LoginConfig>(json);//反序列化JSON反序列化为配置对象
                    if (cfg != null)
                    {
                        if (!string.IsNullOrWhiteSpace(cfg.ServerIP)) host = cfg.ServerIP;//设置服务器地址
                        if (!string.IsNullOrWhiteSpace(cfg.ServerPort)) port = cfg.ServerPort;//设置端口设置服务器端口
                    }
                }

                _svc = new MySqlAuthService(host, port);
                RefreshDepartmentsAsync();
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "初始化失败：" + ex.Message;
            }
        }
        /// <summary>
        /// 数据库服务
        /// </summary>
        private MySqlAuthService _svc;

        /// <summary>
        /// 刷新部门
        /// </summary>
        private async void RefreshDepartmentsAsync()
        {
            // 先确保 _svc 已初始化，避免 NullReferenceException
            if (!EnsureSvcInitialized())
            {
                TxtStatus.Text = "未能初始化部门服务，请检查配置。";
                MessageBox.Show("未能初始化部门服务（_svc）。请检查登录配置或网络，并重试。", "初始化失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            TxtStatus.Text = "正在刷新部门...";
            await Task.Run(() =>
            {
                try
                {
                    _svc.EnsureCategoriesTableExists();
                    _svc.EnsureDepartmentsTableExists();
                }
                catch { }
            });

            await Task.Delay(50);
            try
            {
                var depts = await Task.Run(() => _svc.GetDepartmentsWithCounts());
                DepartmentsGrid.ItemsSource = depts;
                TxtStatus.Text = $"加载完成，共 {depts.Count} 个部门。";
                UsersGrid.ItemsSource = null;
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "刷新部门失败：" + ex.Message;
                //TxtStatus.Text = "刷新部门失败：" + ex.Message;
            }
        }
        /// <summary>
        /// 按分类同步部门
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnSync_Click(object sender, RoutedEventArgs e)
        {
            TxtStatus.Text = "正在按分类同步部门...";
            try
            {
                await Task.Run(() =>
                {
                    _svc.EnsureCategoriesTableExists();
                    _svc.EnsureDepartmentsTableExists();
                    _svc.SyncDepartmentsFromCadCategories();
                });
                RefreshDepartmentsAsync();
                MessageBox.Show("同步完成。", "信息", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "同步失败：" + ex.Message;
            }
        }
        /// <summary>
        /// 新增部门
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnAddDept_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ShowDepartmentEditor(null, out var result))
                {
                    int? mgrId = (result.ManagerUserId.HasValue && result.ManagerUserId.Value > 0)
                        ? result.ManagerUserId
                        : null;

                    var id = _svc.AddDepartment(result.Name, result.DisplayName, result.Description, mgrId, result.SortOrder);
                    if (id > 0)
                    {
                        RefreshDepartmentsAsync();
                        System.Windows.MessageBox.Show("新增部门成功。", "成功", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("新增部门失败。", "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("新增部门异常：" + ex.Message, "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// 修改部门
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnEditDept_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sel = DepartmentsGrid.SelectedItem as DepartmentModel;
                if (sel == null)
                {
                    System.Windows.MessageBox.Show("请先选择一个部门", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }

                if (ShowDepartmentEditor(sel, out var edited))
                {
                    int? mgrId = (edited.ManagerUserId.HasValue && edited.ManagerUserId.Value > 0)
                        ? edited.ManagerUserId
                        : null;

                    var ok = _svc.UpdateDepartment(sel.Id, edited.Name, edited.DisplayName, edited.Description, edited.SortOrder, mgrId, edited.IsActive);
                    if (ok)
                    {
                        RefreshDepartmentsAsync();
                        System.Windows.MessageBox.Show("修改成功。", "成功", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("修改失败。", "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("修改部门异常：" + ex.Message, "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// 删除部门
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDeleteDept_Click(object sender, RoutedEventArgs e)
        {
            var sel = DepartmentsGrid.SelectedItem as DepartmentModel;
            if (sel == null) { MessageBox.Show("请先选择一个部门", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            if (MessageBox.Show($"确认删除部门：{sel.DisplayName} ?\n删除后该部门下用户将被置为未分配。", "确认删除", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;
            var ok = _svc.DeleteDepartment(sel.Id);
            if (ok)
            {
                RefreshDepartmentsAsync();
                MessageBox.Show("删除成功。", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else MessageBox.Show("删除失败。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        /// <summary>
        /// 部门选择变更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DepartmentsGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var sel = DepartmentsGrid.SelectedItem as DepartmentModel;
            if (sel == null) { UsersGrid.ItemsSource = null; return; }
            LoadUsersForDepartment(sel.Id);
        }
        /// <summary>
        /// 加载部门用户
        /// </summary>
        /// <param name="departmentId"></param>
        private async void LoadUsersForDepartment(int departmentId)
        {
            TxtStatus.Text = "正在加载用户...";
            try
            {
                var users = await Task.Run(() => _svc.GetUsersByDepartmentId(departmentId));
                UsersGrid.ItemsSource = users;
                TxtStatus.Text = $"部门用户：{users.Count} 个";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "加载用户失败：" + ex.Message;
            }
        }
        /// <summary>
        /// 在运行时创建并显示一个简洁的部门编辑模态窗口（Add / Edit 共用）
        /// 如果传入的 initial 为 null 表示新增；否则以 initial 填充默认值用于编辑。
        /// 返回 true 表示用户点击“确定”，并通过 out 返回填好的 DepartmentModel。
        /// </summary>
        private bool ShowDepartmentEditor(DepartmentModel initial, out DepartmentModel result)
        {
            result = null;

            var model = new DepartmentModel
            {
                Id = initial?.Id ?? 0,
                Name = initial?.Name ?? string.Empty,
                DisplayName = initial?.DisplayName ?? string.Empty,
                Description = initial?.Description ?? string.Empty,
                SortOrder = initial?.SortOrder ?? 0,
                ManagerUserId = initial?.ManagerUserId,
                IsActive = initial?.IsActive ?? true
            };

            var win = new System.Windows.Window
            {
                Title = initial == null ? "新增部门" : "编辑部门",
                Owner = System.Windows.Window.GetWindow(this),
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner,
                SizeToContent = System.Windows.SizeToContent.WidthAndHeight,
                ResizeMode = System.Windows.ResizeMode.NoResize,
                WindowStyle = System.Windows.WindowStyle.ToolWindow,
                MinWidth = 420
            };

            var grid = new System.Windows.Controls.Grid { Margin = new System.Windows.Thickness(10) };
            for (int i = 0; i < 7; i++) grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new System.Windows.GridLength(100) });
            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });

            int r = 0;

            var lblName = new System.Windows.Controls.TextBlock { Text = "名称:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblName, r); Grid.SetColumn(lblName, 0); grid.Children.Add(lblName);
            var tbName = new System.Windows.Controls.TextBox { Text = model.Name, Margin = new System.Windows.Thickness(4) };
            Grid.SetRow(tbName, r); Grid.SetColumn(tbName, 1); grid.Children.Add(tbName);
            r++;

            var lblDisplay = new System.Windows.Controls.TextBlock { Text = "显示名称:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblDisplay, r); Grid.SetColumn(lblDisplay, 0); grid.Children.Add(lblDisplay);
            var tbDisplay = new System.Windows.Controls.TextBox { Text = model.DisplayName, Margin = new System.Windows.Thickness(4) };
            Grid.SetRow(tbDisplay, r); Grid.SetColumn(tbDisplay, 1); grid.Children.Add(tbDisplay);
            r++;

            var lblDesc = new System.Windows.Controls.TextBlock { Text = "描述:", VerticalAlignment = System.Windows.VerticalAlignment.Top, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblDesc, r); Grid.SetColumn(lblDesc, 0); grid.Children.Add(lblDesc);
            var tbDesc = new System.Windows.Controls.TextBox { Text = model.Description, Margin = new System.Windows.Thickness(4), AcceptsReturn = true, Height = 80, TextWrapping = System.Windows.TextWrapping.Wrap };
            Grid.SetRow(tbDesc, r); Grid.SetColumn(tbDesc, 1); grid.Children.Add(tbDesc);
            r++;

            var lblSort = new System.Windows.Controls.TextBlock { Text = "排序序号:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblSort, r); Grid.SetColumn(lblSort, 0); grid.Children.Add(lblSort);
            var tbSort = new System.Windows.Controls.TextBox { Text = model.SortOrder.ToString(), Margin = new System.Windows.Thickness(4) };
            Grid.SetRow(tbSort, r); Grid.SetColumn(tbSort, 1); grid.Children.Add(tbSort);
            r++;

            var lblMgr = new System.Windows.Controls.TextBlock { Text = "负责人ID:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblMgr, r); Grid.SetColumn(lblMgr, 0); grid.Children.Add(lblMgr);
            var tbMgr = new System.Windows.Controls.TextBox { Text = model.ManagerUserId?.ToString() ?? string.Empty, Margin = new System.Windows.Thickness(4) };
            Grid.SetRow(tbMgr, r); Grid.SetColumn(tbMgr, 1); grid.Children.Add(tbMgr);
            r++;

            var lblActive = new System.Windows.Controls.TextBlock { Text = "是否启用:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblActive, r); Grid.SetColumn(lblActive, 0); grid.Children.Add(lblActive);
            var cbActive = new System.Windows.Controls.CheckBox { IsChecked = model.IsActive, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(4) };
            Grid.SetRow(cbActive, r); Grid.SetColumn(cbActive, 1); grid.Children.Add(cbActive);
            r++;

            var panelBtns = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment = System.Windows.HorizontalAlignment.Right, Margin = new System.Windows.Thickness(0, 10, 0, 0) };
            var btnOk = new System.Windows.Controls.Button { Content = "确定", Width = 80, Margin = new System.Windows.Thickness(4) };
            var btnCancel = new System.Windows.Controls.Button { Content = "取消", Width = 80, Margin = new System.Windows.Thickness(4) };
            panelBtns.Children.Add(btnOk); panelBtns.Children.Add(btnCancel);
            Grid.SetRow(panelBtns, r); Grid.SetColumn(panelBtns, 0); Grid.SetColumnSpan(panelBtns, 2);
            grid.Children.Add(panelBtns);

            win.Content = grid;

            btnCancel.Click += (s, e) => win.DialogResult = false;

            btnOk.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(tbName.Text))
                {
                    System.Windows.MessageBox.Show("请填写部门名称。", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    tbName.Focus();
                    return;
                }

                int sort = 0;
                if (!int.TryParse(tbSort.Text.Trim(), out sort)) sort = 0;

                int? mgr = null;
                if (int.TryParse(tbMgr.Text.Trim(), out var mid) && mid > 0) mgr = mid;

                model.Name = tbName.Text.Trim();
                model.DisplayName = tbDisplay.Text.Trim();
                model.Description = tbDesc.Text;
                model.SortOrder = sort;
                model.ManagerUserId = mgr;
                model.IsActive = cbActive.IsChecked ?? true;

                win.DialogResult = true;
            };

            var res = win.ShowDialog();
            if (res == true)
            {
                result = model;
                return true;
            }

            return false;
        }
        /// <summary>
        /// 分配用户
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnAssignUser_Click(object sender, RoutedEventArgs e)
        {
            var sel = DepartmentsGrid.SelectedItem as DepartmentModel;
            if (sel == null) { MessageBox.Show("请先选择部门", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            var username = (TxtSearchUser.Text ?? "").Trim();
            if (string.IsNullOrEmpty(username)) { MessageBox.Show("请输入要分配的用户名", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); return; }

            var ok = _svc.AssignUserToDepartmentByUsername(username, sel.Id);
            if (ok)
            {
                LoadUsersForDepartment(sel.Id);
                MessageBox.Show($"用户 {username} 已分配到 {sel.DisplayName}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show($"分配失败，检查用户名是否存在或数据库状态。", "失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// 刷新部门
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            RefreshDepartmentsAsync();
        }
        /// <summary>
        /// 刷新用户
        /// </summary>
        /// <returns></returns>
        private bool EnsureSvcInitialized()
        {
            if (_svc != null) return true;

            try
            {
                // 读取 login_config.json 或使用默认值初始化 _svc
                var cfgPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GB_NewCadPlus_LM", "login_config.json");
                string host = "127.0.0.1";
                string port = "3306";

                if (File.Exists(cfgPath))
                {
                    try
                    {
                        var json = File.ReadAllText(cfgPath);
                        var ser = new JavaScriptSerializer();
                        var dict = ser.Deserialize<Dictionary<string, object>>(json);
                        if (dict != null)
                        {
                            if (dict.TryGetValue("ServerIP", out var sip) && sip != null) host = sip.ToString();
                            if (dict.TryGetValue("ServerPort", out var sport) && sport != null) port = sport.ToString();
                        }
                    }
                    catch
                    {
                        // 忽略解析错误，使用默认 host/port
                    }
                }

                _svc = new MySqlAuthService(host, port);
                LogManager.Instance.LogInfo($"MySqlAuthService 已初始化: {host}:{port}");
                return true;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"EnsureSvcInitialized 初始化失败: {ex.Message}");
                _svc = null;
                return false;
            }
        }

        #endregion

        /// <summary>
        /// 清理参数输入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CleanupParameter_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                var textBox = sender as TextBox;
                if (textBox != null)
                {
                    // 获取文本框的值
                    string text = textBox.Text.Trim();

                    // 如果文本框为空，使用默认值7
                    if (string.IsNullOrEmpty(text))
                    {
                        VariableDictionary.cleanupParameter = 7.0;
                        return;
                    }

                    // 尝试解析数值
                    if (double.TryParse(text, out double value))
                    {
                        // 确保值在合理范围内（例如0.1到100之间）
                        if (value >= 0.1 && value <= 100)
                        {
                            VariableDictionary.cleanupParameter = value;
                        }
                        else
                        {
                            // 如果超出范围，使用默认值
                            VariableDictionary.cleanupParameter = 7.0;
                        }
                    }
                    else
                    {
                        // 如果解析失败，使用默认值
                        VariableDictionary.cleanupParameter = 7.0;
                    }
                }
            }
            catch (Exception ex)
            {
                // 如果出现异常，使用默认值
                VariableDictionary.cleanupParameter = 7.0;
                // 可选：记录错误日志
                // Console.WriteLine($"设置清理参数时出错: {ex.Message}");
            }
        }
        private void 分解块_Btn_Click(object sender, RoutedEventArgs e)
        {
            Env.Document.SendStringToExecute("ExplodeNestedBlock ", false, false, false);
            //Env.Document.SendStringToExecute("EXPLODE_AND_REBLOCK ", false, false, false);

        }

        private void 分解图层块_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    // 在文档锁定状态下执行命令
                    using (doc.LockDocument())
                    {
                        // 执行交互式分解图层块命令
                        doc.SendStringToExecute("ExplodeBlocksInLayerInteractive ", false, false, false);
                    }
                }
                else
                {
                    MessageBox.Show("无法获取当前CAD文档！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动分解图层块时发生错误：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                LogManager.Instance.LogError($"启动分解图层块时出错: {ex.Message}");
            }
        }

        private void 计算暖通房间面积_Click(object sender, RoutedEventArgs e)
        {
            VariableDictionary.btnBlockLayer = "暖通房间面积";//设置为被插入的图层名
            VariableDictionary.buttonText = "暖通房间面积";
            VariableDictionary.layerColorIndex = 6;//设置图层颜色

            Env.Document.SendStringToExecute("AreaByPoints ", false, false, false);
        }
    }
    /// <summary>
    /// DataGrid 绑定使用的行模型（用于 LayerDictionary_DataGrid）
    /// </summary>
    public class LayerDictionaryRow
    {
        /// <summary>
        /// 数据库 id，0 表示新行（未持久化）
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// 显示序号
        /// </summary>
        public int DisplayIndex { get; set; }
        /// <summary>
        /// 专业列（可填写或来自部门）
        /// </summary>
        public string Major { get; set; } = "";
        /// <summary>
        /// 原图层名（只读显示）
        /// </summary>
        public string LayerName { get; set; } = "";
        /// <summary>
        /// 解释图层名（可编辑）
        /// </summary>
        public string DicLayerName { get; set; } = "";
        /// <summary>
        /// 来源（personal/standard）
        /// </summary>
        public string Source { get; set; } = "personal";
    }

    /// <summary>
    /// 图层信息类
    /// </summary>
    public class LayerInfo : INotifyPropertyChanged
    {
        /// <summary>
        /// 原始序号
        /// </summary>
        private int _index;              // 原始序号
        /// <summary>
        /// 显示序号
        /// </summary>
        private int _displayIndex;       // 显示序号
        /// <summary>
        /// 图层名
        /// </summary>
        private string _layerName;
        /// <summary>
        /// 是否可见
        /// </summary>
        private bool _isOn;
        /// <summary>
        /// 是否冻结
        /// </summary>
        private bool _isFrozen;
        /// <summary>
        /// 颜色索引
        /// </summary>
        private short _colorIndex;
        /// <summary>
        /// 是否删除
        /// </summary>
        private bool _isDelete;
        /// <summary>
        /// 颜色
        /// </summary>
        private Autodesk.AutoCAD.Colors.Color _autoCadColor;
        /// <summary>
        /// 序号
        /// </summary>
        public int Index
        {
            get => _index;
            set { _index = value; OnPropertyChanged(); }
        }
        /// <summary>
        /// 显示序号
        /// </summary>
        public int DisplayIndex
        {
            get => _displayIndex;
            set { _displayIndex = value; OnPropertyChanged(); }
        }
        /// <summary>
        /// 图层名
        /// </summary>
        public string LayerName
        {
            get => _layerName;
            set { _layerName = value; OnPropertyChanged(); }
        }
        /// <summary>
        /// 是否可见
        /// </summary>
        public bool IsOn
        {
            get => _isOn;
            set { _isOn = value; OnPropertyChanged(); }
        }
        /// <summary>
        /// 是否冻结
        /// </summary>
        public bool IsFrozen
        {
            get => _isFrozen;
            set { _isFrozen = value; OnPropertyChanged(); }
        }
        /// <summary>
        /// 颜色索引
        /// </summary>
        public short ColorIndex
        {
            get => _colorIndex;
            set { _colorIndex = value; OnPropertyChanged(); }
        }
        /// <summary>
        /// 是否删除
        /// </summary>
        public bool IsDelete
        {
            get => _isDelete;
            set { _isDelete = value; OnPropertyChanged(); }
        }
        /// <summary>
        /// CAD_Color颜色
        /// </summary>
        public Autodesk.AutoCAD.Colors.Color Color
        {
            get => _autoCadColor;
            set { _autoCadColor = value; OnPropertyChanged(); }
        }
        /// <summary>
        /// 属性变更通知
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;
        /// <summary>
        /// 属性变更通知方法
        /// </summary>
        /// <param name="propertyName"></param>
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// 图层状态快照类（用于还原功能）
    /// </summary>
    public class LayerStateSnapshot
    {
        /// <summary>
        /// 图层信息字典
        /// </summary>
        public Dictionary<string, LayerInfo> Layers { get; set; }
        /// <summary>
        /// 时间戳
        /// </summary>
        public DateTime Timestamp { get; set; }
        /// <summary>
        /// 构造函数
        /// </summary>
        public LayerStateSnapshot()
        {
            Layers = new Dictionary<string, LayerInfo>();
            Timestamp = DateTime.Now;
        }
    }

    /// <summary>
    /// 分类属性编辑模型
    /// </summary>
    public class CategoryPropertyEditModel : INotifyPropertyChanged
    {
        /// <summary>
        /// 属性名称1
        /// </summary>
        private string? _propertyName1;
        /// <summary>
        /// 属性值1
        /// </summary>
        private string? _propertyValue1;
        /// <summary>
        /// 属性名称2
        /// </summary>
        private string? _propertyName2;
        /// <summary>
        /// 属性值2
        /// </summary>
        private string? _propertyValue2;
        /// <summary>
        /// 属性名称1
        /// </summary>
        public string PropertyName1
        {
            get => _propertyName1;
            set
            {
                if (_propertyName1 != value)
                {
                    _propertyName1 = value;
                    OnPropertyChanged();
                }
            }
        }
        /// <summary>
        /// 属性值1
        /// </summary>
        public string PropertyValue1
        {
            get => _propertyValue1;
            set
            {
                if (_propertyValue1 != value)
                {
                    _propertyValue1 = value;
                    OnPropertyChanged();
                }
            }
        }
        /// <summary>
        /// 属性名称2
        /// </summary>
        public string PropertyName2
        {
            get => _propertyName2;
            set
            {
                if (_propertyName2 != value)
                {
                    _propertyName2 = value;
                    OnPropertyChanged();
                }
            }
        }
        /// <summary>
        /// 属性值2
        /// </summary>
        public string PropertyValue2
        {
            get => _propertyValue2;
            set
            {
                if (_propertyValue2 != value)
                {
                    _propertyValue2 = value;
                    OnPropertyChanged();
                }
            }
        }
        /// <summary>
        /// 属性变更通知方法
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;
        /// <summary>
        /// 属性变更通知
        /// </summary>
        /// <param name="propertyName"></param>
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
