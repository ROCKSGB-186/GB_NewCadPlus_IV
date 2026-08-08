using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Ribbon;
using Autodesk.AutoCAD.Runtime;
using Dm;
using GB_NewCadPlus_IV.DisplayPages;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.Models;
using GB_NewCadPlus_IV.UniFiedStandards;
using GB_NewCadPlus_IV.ViewModels;
using GB_NewCadPlus_IV.Views;
using IFoxCAD.Cad;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.SS.Util;          // 用于 CellRangeAddress 等辅助类
using NPOI.XSSF.UserModel;  // 仅用于创建 .xlsx 工作簿
using Org.BouncyCastle.Asn1.Cms;
using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using static GB_NewCadPlus_IV.FunctionalMethod.DatabaseManager;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Binding = System.Windows.Data.Binding;
using Border = System.Windows.Controls.Border;
using Brushes = System.Windows.Media.Brushes;
using Button = System.Windows.Controls.Button;
using CellType = NPOI.SS.UserModel.CellType;
using ComboBox = System.Windows.Controls.ComboBox;
using ContextMenu = System.Windows.Controls.ContextMenu;
using Control = System.Windows.Controls.Control;
using DataGrid = System.Windows.Controls.DataGrid;
using DataTable = System.Data.DataTable;
using FontFamily = System.Windows.Media.FontFamily;
using HorizontalAlignment = System.Windows.HorizontalAlignment;
using Image = System.Windows.Controls.Image;
using MenuItem = System.Windows.Controls.MenuItem;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using Orientation = System.Windows.Controls.Orientation;
using Panel = System.Windows.Controls.Panel;
using Pen = System.Windows.Media.Pen;
using Point = System.Windows.Point;
using SystemColors = System.Windows.SystemColors;
using Table = Autodesk.AutoCAD.DatabaseServices.Table;
using TextBox = System.Windows.Controls.TextBox;
using UserControl = System.Windows.Controls.UserControl;

namespace GB_NewCadPlus_IV
{
    /// <summary>
    /// WpfMainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class WpfMainWindow : UserControl
    {
        #region 私有字段与属性（翻译自反编译，保留原语义）
        /// <summary>
        /// 选中的预览图片路径
        /// </summary>
        private string? _selectedPreviewImagePath;
        /// <summary>
        /// 当前文件存储信息
        /// </summary>
        private FileStorage? _currentFileStorage;
        /// <summary>
        /// 数据库连接字符串（可能由外部注入）
        /// </summary>
        private string? _connectionString;
        /// <summary>
        /// 本地选中的 DWG 文件路径（上传/插入）
        /// </summary>
        private string? _selectedFilePath;
        /// <summary>
        /// 文件管理器（延后初始化）
        /// </summary>
        private FileManager? _fileManager;
        /// <summary>
        /// 预览图片内存缓存，避免重复磁盘加载，键为图片路径
        /// </summary>
        private readonly Dictionary<string, BitmapImage> _imageCache = new Dictionary<string, BitmapImage>(); // 缓存预览图片，避免重复加载
        /// <summary>
        /// 动态认证服务实例（兼容 DM 与 MySQL 实现，使用 dynamic 以兼容不同实现的不同方法签名）
        /// </summary>
        private dynamic? _authServiceDynamic; // 运行时绑定认证服务方法，编译时不再报找不到方法的错误
        /// <summary>
        /// 插入时的属性快照（不区分大小写的键）
        /// </summary>
        private Dictionary<string, string> _propertiesSnapshotForInsert = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>
        /// 分类管理器
        /// </summary>
        private CategoryManager? _categoryManager;
        /// <summary>
        /// 当前操作类型（管理分类时使用）
        /// </summary>
        private ManagementOperationType _currentOperation = ManagementOperationType.None;
        /// <summary>
        /// 分类树节点缓存
        /// </summary>
        private List<CategoryTreeNode> _categoryTreeNodes = new List<CategoryTreeNode>();
        /// <summary>
        /// 数据库管理器实例（外部注入或内部创建）
        /// </summary>
        private DatabaseManager? _databaseManager;
        private readonly GraphicApiService _graphicApiService = new GraphicApiService();
        private readonly AuthUserDepartmentApiService _apiService = new AuthUserDepartmentApiService();
        private readonly DepartmentApiService _departmentApiService = new DepartmentApiService();
        /// <summary>
        /// 规范管理 API 服务，负责目录查询和文件预览请求。
        /// </summary>
        private readonly StandardApiService _standardManagementApiService = new StandardApiService();
        /// <summary>
        /// 规范管理树数据源，避免与图元分类树缓存混用。
        /// </summary>
        private readonly List<CategoryTreeNode> _standardTreeNodes = new List<CategoryTreeNode>();
        /// <summary>
        /// 当前选中的分类节点
        /// </summary>
        private CategoryTreeNode? _selectedCategoryNode;
        /// <summary>
        /// 上一次同步清单（用于同步流程回溯/展示）
        /// </summary>
        private SyncManifest? _lastSyncManifest;
        /// <summary>
        /// 同步流程的并发控制信号量（避免并发执行）
        /// </summary>
        private readonly SemaphoreSlim _syncSemaphore = new SemaphoreSlim(1, 1);
        /// <summary>
        /// 同步取消令牌源
        /// </summary>
        private CancellationTokenSource? _syncCancellationSource;
        /// <summary>
        /// 同步源根路径（服务端共享根）
        /// </summary>
        private string? _syncSourceRoot;
        /// <summary>
        /// 本机映射的本地根路径（用于回源时替换盘符）
        /// </summary>
        private string? _syncLocalRoot;
        /// <summary>
        /// 是否使用数据库模式（在线）还是离线资源模式
        /// </summary>
        private bool _useDatabaseMode = true;
        /// <summary>
        /// 当前选中的数据库类型（"DM" 或 "MYSQL"）
        /// </summary>
        private string _currentDatabaseType = string.Empty;
        /// <summary>
        /// 当前选中的节点 ID（用于面板定位）
        /// </summary>
        private int _currentNodeId = 0;
        /// <summary>
        /// CAD 与 SW 存储路径（配置项）
        /// </summary>
        private string _cadStoragePath = string.Empty;
        /// <summary>
        /// SW 存储路径（配置项）
        /// </summary>
        private string _swStoragePath = string.Empty;
        /// <summary>
        /// 分类树视图引用（在初始化时赋值）
        /// </summary>
        private System.Windows.Controls.TreeView? _categoryTreeView;
        /// <summary>
        /// 预览容器（如果 XAML 中存在）
        /// </summary>
        private Viewbox? previewViewbox;
        /// <summary>
        /// 应用本地数据路径（LocalAppData\GB_CADPLUS）
        /// </summary>
        public static string AppPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GB_CADPLUS");
        /// <summary>
        /// 文件路径和名称（可能由外部设置，用于插入或其他操作）
        /// </summary>
        public static string? filePathAndName = null;
        /// <summary>
        /// 刷新文件列表的标志（可能由外部设置，触发界面刷新）
        /// </summary>
        public static string referenceFile = Path.Combine(AppPath, "ReferenceFile");
        /// <summary>
        /// 层管理器与层数据源
        /// </summary>
        private LayerManager? _layerManager;
        /// <summary>
        /// 图层信息数据源（绑定到 LayerDataGrid，展示当前图层列表）
        /// </summary>
        private ObservableCollection<LayerInfo>? _layerData;
        /// <summary>
        /// 标识主窗口是否打开
        /// </summary>
        public static bool wpfMainWindowsIsOpenClose = false;
        /// <summary>
        /// 限制同时下载的文件数量，避免对服务器造成过大压力
        /// </summary>
        public static readonly SemaphoreSlim _cacheDownloadSemaphore = new SemaphoreSlim(5);
        /// <summary>
        /// 原有下载客户端（保持不变）
        /// </summary>
        private static readonly HttpClient _downloadHttpClient = new HttpClient
        {
            Timeout = TimeSpan.FromMinutes(5)   // 下载大文件需要较长时间
        };
        /// <summary>
        /// 新增上传客户端（直接在声明时初始化）
        /// </summary>
        private static readonly HttpClient _uploadHttpClient = new HttpClient
        {
            Timeout = TimeSpan.FromMinutes(5)   // 上传预览图或大文件需要较长时间
        };
       
        #endregion

        #region 公开静态实例（方便外部访问）
        // 设置全局静态实例，供非 UI 代码读取 TextBox_绘图比例 等
        public static WpfMainWindow? Instance { get; private set; }
        #endregion

        #region 枚举与内部类型
        // 管理操作类型：添加主分类 / 添加子分类 等
        public enum ManagementOperationType
        {
            None,
            AddCategory,
            AddSubcategory
        }
        #endregion

        #region 构造函数与初始化
#pragma warning disable CS8618 // 在退出构造函数时，不可为 null 的字段必须包含非 null 值。请考虑添加 "required" 修饰符或声明为可为 null。
        public WpfMainWindow()
#pragma warning restore CS8618 // 在退出构造函数时，不可为 null 的字段必须包含非 null 值。请考虑添加 "required" 修饰符或声明为可为 null。

        {
            InitializeComponent(); // WPF 生成的控件初始化（XAML -> 对象树）
            Instance = this; // 注册静态实例
            UnifiedUIManager.SetWpfInstance(this); // 注册到统一 UI 管理器（便于其它模块访问）
            LogManager.Instance.LogInfo("WPF实例已注册到UnifiedUIManager");

            // 初始化基准图层集合（静态方法，将常用图层填充到 VariableDictionary.allTjtLayer）
            NewTjLayer();
            
            // 窗口加载完成事件，用于延迟加载数据
            Loaded += WpfMainWindow_Loaded;
            
            _fileManager = null;
            _categoryManager = null;
            this.DataContext = new PipeCalcViewModel();// 初始化计算表

            // 层管理器与层数据源初始化（仅对象创建，实际数据在 InitializeLayerDataGrid 中绑定）
            _layerManager = new LayerManager();
            _layerData = new ObservableCollection<LayerInfo>();

            // 初始化界面相关的数据网格/控件绑定
            InitializeLayerDataGrid();

            // 初始化计算 CSV 表结构（方法实现可能在后续段）
            //InitializeCalcCsvTables();
        }
        #endregion

        #region 核心初始化方法（小而明确的职责）
        /// <summary>
        /// 初始化 LayerDataGrid 的列与数据绑定（不自动生成列）
        /// </summary>
        private void InitializeLayerDataGrid()
        {
            try
            {
                if (LayerDataGrid == null) return; // 检查对象是否为空
                LayerDataGrid.AutoGenerateColumns = false;
                LayerDataGrid.ItemsSource = _layerData;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("InitializeLayerDataGrid 发生异常: " + ex.Message);
            }
        }

        /// <summary>
        /// 异步初始化 LayerDictionary_DataGrid 的数据源与事件订阅（使用时调用）
        /// </summary>
        /// <returns></returns>
        private async Task InitializeLayerDictionaryDataGridSource()
        {
            // 把行集合绑定到 DataGrid
            LayerDictionary_DataGrid.ItemsSource = _layerDictionaryRows;
            // 保证滚动条行为可用
            LayerDictionary_DataGrid.SetValue(ScrollViewer.VerticalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
            LayerDictionary_DataGrid.SetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
            LayerDictionary_DataGrid.SetValue(ScrollViewer.CanContentScrollProperty, false);

            // 先移除再添加事件，避免重复订阅
            LayerDictionary_DataGrid.PreparingCellForEdit -= LayerDictionary_DataGrid_PreparingCellForEdit;
            LayerDictionary_DataGrid.PreparingCellForEdit += LayerDictionary_DataGrid_PreparingCellForEdit;
            LayerDictionary_DataGrid.CellEditEnding -= LayerDictionary_DataGrid_CellEditEnding;
            LayerDictionary_DataGrid.CellEditEnding += LayerDictionary_DataGrid_CellEditEnding;

            try
            {
                // 如果分类名集合为空且数据库可用，则尝试加载分类名
                if ((_categoryNames == null || _categoryNames.Count == 0) && _databaseManager != null && _databaseManager.IsDatabaseAvailable)
                    await LoadCategoryNamesAsync();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("初始化 CategoryNames 时出错: " + ex.Message);
            }

            // 为“专业”列设置下拉数据源（如果存在 DataGridComboBoxColumn）
            try
            {
                var comboCol = LayerDictionary_DataGrid.Columns.OfType<DataGridComboBoxColumn>().FirstOrDefault(c => (c.Header?.ToString() ?? string.Empty).IndexOf("专业", StringComparison.OrdinalIgnoreCase) >= 0);
                if (comboCol != null)
                {
                    comboCol.ItemsSource = _categoryNames;
                    comboCol.SelectedItemBinding = new System.Windows.Data.Binding("Major")
                    {
                        Mode = System.Windows.Data.BindingMode.TwoWay,
                        UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged
                    };
                }
                else if (LayerDictionary_DataGrid.Columns.Count > 1 && !(LayerDictionary_DataGrid.Columns[1] is DataGridComboBoxColumn))
                {
                    var newCombo = new DataGridComboBoxColumn
                    {
                        Header = "专业",
                        Width = 75,
                        ItemsSource = _categoryNames,
                        SelectedItemBinding = new System.Windows.Data.Binding("Major")
                        {
                            Mode = System.Windows.Data.BindingMode.TwoWay,
                            UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged
                        }
                    };
                    LayerDictionary_DataGrid.Columns[1] = newCombo;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("为 LayerDictionary_DataGrid 设置下拉数据源失败: " + ex.Message);
            }
        }

        /// <summary>
        /// 窗体 Loaded 事件处理器：延迟执行 UI/数据库初始化任务
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void WpfMainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                VariableDictionary.wpfWindows_Status = true; // 标记 WPF 窗口已加载
                if (!Directory.Exists(GetPath.PreviewCachePath))
                    Directory.CreateDirectory(GetPath.PreviewCachePath);
                if (!Directory.Exists(GetPath.DwgCachePath))
                    Directory.CreateDirectory(GetPath.DwgCachePath);
                // 显示客户端版本（EntryAssembly 可能为 null，如果为插件则使用执行程序集版本）
                var entryAsm = System.Reflection.Assembly.GetEntryAssembly();
                var execAsm = System.Reflection.Assembly.GetExecutingAssembly();
                var clientVersion = entryAsm?.GetName().Version?.ToString() ?? execAsm?.GetName().Version?.ToString() ?? "未知";
                if (FindName("TextBox客户端版本") is TextBox clientVersionBox)
                    clientVersionBox.Text = clientVersion;
                // 读取服务器端版本号（通过 DatabaseManager）
                var serverVersionText = "未连接";
                try
                {
                    if (_databaseManager != null)
                    {
                        string svr = await _databaseManager.GetSystemConfigValueAsync("Version").ConfigureAwait(true);
                        serverVersionText = string.IsNullOrWhiteSpace(svr) ? "服务器未设置版本" : svr;
                    }
                    else
                    {
                        serverVersionText = "本地未初始化数据库管理器";
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogInfo("读取服务器版本失败: " + ex.Message);
                    serverVersionText = "读取失败";
                }

                if (FindName("ServerVer") is TextBlock serverVerBlock)
                    serverVerBlock.Text = serverVersionText;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("WpfMainWindow_Loaded 出错: " + ex.Message);
            }

            try
            {
                bool connected = false;// 是否能数据库连接
                try
                {
                    connected = await EnsureDatabaseConnectedSingleEntryAsync();// 确保数据库连接（单入口方法，内部会处理连接字符串和实例创建）
                    wpfMainWindowsIsOpenClose = true;// 标记主窗口已打开（可能在 EnsureDatabaseConnectedSingleEntryAsync 内部也会设置，视具体实现而定）
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogWarning("主窗口加载连接检测异常: " + ex.Message);
                    connected = false;
                }

                if (connected && _databaseManager != null)
                {
                    _useDatabaseMode = true; // 成功连接数据库，启用在线模式
                    _fileManager = new FileManager(_databaseManager); // 初始化文件管理器实例
                    _categoryManager = new CategoryManager(_databaseManager); // 初始化分类管理器实例
                    try { ReinitializeDatabase(); } catch { } // 尝试重新初始化数据库连接（如果之前的连接参数不完整或需要刷新）
                    try { UpdateAdminTabsVisibility(); } catch { } // 尝试更新管理员标签页的可见性（基于数据库连接状态和用户角色）
                }
                else
                {
                    _useDatabaseMode = false; // 无法连接数据库，启用离线模式
                    LogManager.Instance.LogInfo("主窗口启动：由于没有可用数据库连接或外部未注入成功，当前处于离线模式。");
                }

                try { UpdateAdminTabsVisibility(); } catch (Exception ex) { LogManager.Instance.LogWarning("管理员标签页初始化异常: " + ex.Message); }

                AddContextMenuToTreeView(CategoryTreeView); // 为分类树视图添加右键菜单（如果 XAML 中存在）
                PropertiesDataGrid = FindVisualChild<System.Windows.Controls.DataGrid>(this, "PropertiesDataGrid"); // 查找属性数据网格（如果 XAML 中存在）
                await InitializeLayerDictionaryDataGridSource(); // 初始化 LayerDictionary_DataGrid 的数据源与事件订阅（使用时调用）
                Loaded += DepartmentAdminControl_Loaded; // 部门管理员控件加载事件（如果存在部门管理员控件）
                TextBox插件版本.Text = $"{System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}";
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning("WpfMainWindow_Loaded 异常: " + ex.Message);
                MessageBox.Show("主界面加载异常: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        #endregion
        
        /// <summary>
        /// 根据当前用户（VariableDictionary._userName 或界面用户名）显示/隐藏“管理员模块”与“部门\\人员模块”
        /// 仅当用户名为 "sa"、"admin"、"root" 或 达梦管理员 "SYSDBA"（不区分大小写）时才显示；否则折叠（Collapsed）
        ///</summary>
        private void UpdateAdminTabsVisibility()
        {
            try
            {
                // 获取当前用户名（优先使用全局变量，再退回到设置界面输入）
                // 优先获取原始用户名（保留大小写以便后续数据库查询），再构造小写用于简单白名单判断
                var userNameRaw = VariableDictionary._userName;
                var userName = userNameRaw.ToLowerInvariant();
                // 白名单匹配：系统设置的3个常规管理员和达梦/MySQL的两个超级管理员
                bool isAdmin = userName == "sa" || userName == "admin" || userName == "root" || userName == "sysdba";

                // 若白名单未命中且数据库可用，则以 USERS 表中的 ROLE 字段为准进行角色判定（更灵活且可由管理员在 DB 中维护）
                if (!isAdmin && _databaseManager != null && _databaseManager.IsDatabaseAvailable && !string.IsNullOrWhiteSpace(userNameRaw))
                {
                    try
                    {
                        using (var conn = _databaseManager.GetConnection())
                        {
                            // 打开连接并确保 Schema 已设置（Connection_StateChange 会在打开时调用 ApplySchema）
                            conn.Open();
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.CommandText = "SELECT ROLE FROM USERS WHERE UPPER(USERNAME) = :u AND IS_ACTIVE = 1";
                                var p = cmd.CreateParameter();
                                p.ParameterName = "u";
                                p.Value = userNameRaw.ToUpperInvariant();
                                cmd.Parameters.Add(p);
                                var obj = cmd.ExecuteScalar();
                                if (obj != null && obj != DBNull.Value)
                                {
                                    var role = Convert.ToString(obj) ?? string.Empty;
                                    var rl = role.ToLowerInvariant();
                                    // 常见的管理员角色标识：包含 admin、sysdba、或中文“管理员”关键词
                                    if (rl.Contains("admin") || rl.Contains("sysdba") || rl.Contains("管理员"))
                                    {
                                        isAdmin = true;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogInfo($"UpdateAdminTabsVisibility: 查询用户角色失败: {ex.Message}");
                        // 在出现数据库错误时保留白名单判断结果（即不提升权限）
                    }
                }

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
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "GB_CADPLUS",
            "drawing_config.json"
        );
       
        #region 绘图配置（读取/保存/解析绘图比例）
        
        /// <summary>
        /// 保存当前绘图配置到本地 JSON 文件
        /// </summary>
        private void SaveDrawingConfig()
        {
            try
            { var scale = Convert.ToDouble(TextBox绘图比例.Text);
                VariableDictionary.wpfTextBoxScale = scale != 0 ? scale : Convert.ToDouble(TextBox绘图比例.Tag);// 先从 WPF 界面获取绘图比例并更新全局变量
                // 确保目录存在
                var dir = Path.GetDirectoryName(DrawingConfigPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
                // 使用 Newtonsoft.Json 序列化绘图配置对象为 JSON 字符串，并写入文件
                var settings = new JsonSerializerSettings { Formatting = Formatting.Indented, Culture = CultureInfo.InvariantCulture };
                File.WriteAllText(DrawingConfigPath, JsonConvert.SerializeObject(VariableDictionary.wpfTextBoxScale, settings));
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning("保存绘图配置失败: " + ex.Message);
            }
        }

        /// <summary>
        /// 从本地配置文件加载绘图比例并写入 VariableDictionary.wpfTextBoxScale
        /// 返回是否成功加载（true 表示成功并已设置变量）
        /// </summary>
        public double LoadDrawingConfig()
        {
            try
            {
                // 如果 DrawingConfigPath 未定义或为空，则直接返回失败
                if (string.IsNullOrWhiteSpace(DrawingConfigPath)) // 检查路径变量
                {
                    LogManager.Instance.LogWarning("LoadDrawingConfig: DrawingConfigPath 检查路径变量未设置"); // 记录日志
                    return 0; // 返回失败
                }

                // 如果配置文件不存在，则直接返回失败（不抛异常）
                if (!File.Exists(DrawingConfigPath)) // 检查文件是否存在
                {
                    LogManager.Instance.LogWarning("LoadDrawingConfig: DrawingConfigPath 配置文件不存在"); // 记录日志
                    return 0; // 返回失败
                }

                // 读取文件全部文本（使用 UTF8 编码）
                string json = File.ReadAllText(DrawingConfigPath, System.Text.Encoding.UTF8); // 读取文件内容

                // 如果文件为空或全是空白，认为加载失败
                if (string.IsNullOrWhiteSpace(json)) // 内容为空检查
                {
                    LogManager.Instance.LogWarning("LoadDrawingConfig: DrawingConfigPath 文件为空或全是空白"); // 记录日志
                    return 0; // 返回失败
                }

                // 尝试用 JSON 反序列化为 double（因为保存时直接序列化了一个 double 值）
                double parsedValue; // 用于保存解析结果
                try
                {
                    // 优先用 JsonConvert 反序列化（兼容数字或 JSON number）
                    parsedValue = JsonConvert.DeserializeObject<double>(json); // 反序列化为 double
                }
                catch
                {
                    // 如果 JSON 反序列化失败，尝试直接按文本解析为数字（宽松处理）
                    string raw = json.Trim(); // 去除空白
                                              // 去掉可能的引号包裹（以防意外）
                    if ((raw.StartsWith("\"") && raw.EndsWith("\"")) || (raw.StartsWith("'") && raw.EndsWith("'")))
                        raw = raw.Substring(1, raw.Length - 2); // 去掉引号
                                                                // 规范化一些常见全角字符
                    raw = raw.Replace('，', ',').Replace('．', '.').Replace('\u3000', ' '); // 替换全角符号
                                                                                          // 尝试不变文化解析（小数点为 '.' 的情况）
                    if (!double.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out parsedValue))
                    {
                        // 回退到当前文化解析
                        if (!double.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out parsedValue))
                        {
                            // 若仍无法解析，记录日志并返回失败
                            LogManager.Instance.LogWarning($"LoadDrawingConfig: 无法解析绘图比例，原文='{json}'"); // 记录无法解析的原文
                            return 0; // 返回失败
                        }
                    }
                }

                // 解析成功后，把值写入全局变量
                VariableDictionary.wpfTextBoxScale = parsedValue; // 设置全局变量
                
                // 成功加载并设置变量
                return VariableDictionary.wpfTextBoxScale; // 返回成功
            }
            catch (Exception ex)
            {
                // 捕获所有意外异常并记录日志，避免程序中断
                LogManager.Instance.LogWarning("加载绘图配置失败: " + ex.Message); // 记录异常
                return 0; // 返回失败
            }
        }

        #endregion

        #region 数据库初始化/检测/重置（与 DatabaseManager 协作）

        /// 外部注入 DatabaseManager（例如在程序外部创建后调用）
        public void SetInitialDatabase(DatabaseManager db)
        {
            if (db == null || !db.IsDatabaseAvailable) return;
            _databaseManager = db;
            _useDatabaseMode = true;
            wpfMainWindowsIsOpenClose = true;
            
            LogManager.Instance.LogInfo("WpfMainWindow: 已从外部注入数据库管理器并锁定联机状态。");
        }

        /// <summary>
        /// 尝试静默检测并建立数据库连接（不弹 UI 提示）
        /// </summary>
        /// <returns></returns>
        private async Task<bool> EnsureDatabaseConnectedSilentAsync()
        {
            try
            {
                if (_databaseManager != null && _databaseManager.IsDatabaseAvailable) return true;
                string dbtype = (VariableDictionary._databaseType ?? "DM").ToUpper().Trim();// 默认为达梦（DM），也支持 MYSQL
                string defaultUser = dbtype == "MYSQL" ? "root" : "SYSDBA"; // MySQL 默认管理员是 root，达梦默认管理员是 SYSDBA
                string defaultPwd = dbtype == "MYSQL" ? "123456" : "675756SGBsgb"; // MySQL 默认密码（强烈建议修改），达梦默认密码（强烈建议修改）
                int defaultPort = dbtype == "MYSQL" ? 3308 : 3306; // MySQL 默认端口 3306 的变体，达梦默认端口 5236
                string dbUser = string.IsNullOrWhiteSpace(VariableDictionary._dbUserName) ? defaultUser : VariableDictionary._dbUserName.Trim(); // 用户名如果未设置则使用默认管理员用户名
                string dbPassword = string.IsNullOrWhiteSpace(VariableDictionary._dbPassWord) ? defaultPwd : VariableDictionary._dbPassWord; // 密码如果未设置则使用默认管理员密码
                int dbPort = VariableDictionary._dataBaseServerPort > 0 ? VariableDictionary._dataBaseServerPort : defaultPort; // 端口如果未设置或无效则使用默认端口
                string schemaName = string.IsNullOrWhiteSpace(VariableDictionary._dataBaseName) ? "CAD_SW_LIBRARY" : VariableDictionary._dataBaseName.Trim(); // 数据库名如果未设置则使用默认数据库名（达梦默认是 CAD_SW_LIBRARY，MySQL 也使用同名数据库）
                if (string.IsNullOrEmpty(VariableDictionary._serverIP)) return false; // 没有服务器 IP 配置，无法连接数据库

                // 先测试网络连通性（TCP）
                if (!await Task.Run(() => LoginWindow.TestNetworkConnection(VariableDictionary._serverIP, dbPort)))
                    return false;
                string conn; // 根据数据库类型构建连接字符串（达梦和 MySQL 的连接字符串格式不同）
                if (dbtype == "MYSQL")
                    conn = $"Server={VariableDictionary._serverIP};Port={dbPort};Database={schemaName};Uid={dbUser};Pwd={dbPassword};Allow User Variables=True;"; // MySQL 连接字符串，包含允许用户变量（如果需要）
                else
                    conn = $"Server={VariableDictionary._serverIP};Port={dbPort};Schema={schemaName};User Id={dbUser};Password={dbPassword};";// 达梦连接字符串
                // 尝试创建 DatabaseManager 实例并测试数据库可用性
                var db = new DatabaseManager(conn);
                if (db.IsDatabaseAvailable) // 连接成功且数据库可用，设置实例并返回成功
                {
                    _databaseManager = db; // 设置数据库管理器实例
                    return true;
                }
            }
            catch { /* 静默失败 */ }

            return false;
        }

        /// <summary>
        /// 确保数据库连接（单入口方法，内部会处理连接字符串和实例创建）
        /// </summary>
        /// <returns></returns>
        private async Task<bool> EnsureDatabaseConnectedSingleEntryAsync()
        {
            if (_databaseManager != null && _databaseManager.IsDatabaseAvailable)
            {
                _useDatabaseMode = true;// 已经有可用的数据库连接实例，直接使用在线模式
                return true;
            }
            // 首次尝试连接数据库（可能在 WpfMainWindow_Loaded 中调用，或外部调用 SetInitialDatabase 后调用）
            bool connected = await EnsureDatabaseConnectedSilentAsync();
            _useDatabaseMode = connected;// 根据连接结果设置在线/离线模式标志
            return connected;
        }

        /// <summary>
        /// 重新初始化数据库相关缓存与 UI（在连接成功后调用）
        /// </summary>
        private async void ReinitializeDatabase()
        {
            try
            {
                await LoadCategoryNamesAsync();// 重新加载分类名称集合（供下拉列表使用）
                if (_categoryManager != null)
                    await _categoryManager.RefreshCategoryTreeAsync(_selectedCategoryNode, _categoryTreeView, _categoryTreeNodes, _databaseManager); // 刷新分类树（保持当前选中节点）
                LogManager.Instance.LogInfo("数据库连接已重新初始化");
                await RefreshAllCategoryPanelsAsync();// 刷新所有主分类面板（按主分类名称依次触发加载）
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("重新初始化数据库时出错: " + ex.Message);
                MessageBox.Show("重新初始化数据库失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        /// <summary>
        /// 刷新所有主分类面板（按主分类名称依次触发加载）
        /// </summary>
        /// <returns></returns>
        /// <summary>
        /// 异步刷新所有主分类面板（工艺、建筑、结构等），从数据库重新加载按钮
        /// </summary>
        private async Task RefreshAllCategoryPanelsAsync()
        {
            try
            {
                // 检查数据库管理器是否可用，若不可用则跳过刷新并记录日志
                if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
                {
                    LogManager.Instance.LogInfo("RefreshAllCategoryPanelsAsync：数据库不可用，跳过面板刷新");
                    return;
                }

                // 定义所有需要刷新的主分类名称（与界面 Tab 标题一致）
                string[] majorCategories = new string[]
                {
                    "工艺","建筑","结构","电气","给排水","暖通","自控","总图","公共图"
                };

                // 遍历每个主分类，逐一刷新其面板
                foreach (var majorItem in majorCategories)
                {
                    try
                    {
                        // 根据分类名称获取对应的 WrapPanel 面板
                        var majorPanel = GetPanelByFolderName(majorItem);
                        if (majorPanel == null)
                        {
                            // 若未找到面板，记录日志并跳过该分类
                            LogManager.Instance.LogInfo($"RefreshAllCategoryPanelsAsync：未找到面板 {majorItem}，跳过");
                            continue;
                        }

                        // 从数据库异步加载该分类下的按钮数据并填充到面板中
                        await LoadButtonsFromDatabase(majorItem, majorPanel);

                        // 短暂延迟 60ms，避免连续大量操作导致 UI 卡顿，同时给界面更新留出时间
                        await Task.Delay(60);
                    }
                    catch (Exception ex)
                    {
                        // 捕获单个分类加载时的异常，记录日志后继续处理后续分类（不中断整体流程）
                        LogManager.Instance.LogInfo($"RefreshAllCategoryPanelsAsync: 加载分类 {majorItem} 时出错: {ex.Message}");
                    }
                }

                // 记录所有主分类面板刷新完成
                LogManager.Instance.LogInfo("RefreshAllCategoryPanelsAsync: 所有主分类面板刷新完成");
            }
            catch (Exception ex)
            {
                // 捕获整个刷新流程中的未预期异常，记录日志
                LogManager.Instance.LogInfo("RefreshAllCategoryPanelsAsync 异常: " + ex.Message);
            }
        }
        #endregion


        #region // 第二阶段：预览与图片缓存相关方法（可直接替换到 WpfMainWindow.xaml.cs 的相应区域）;说明：包含预览路径解析、本地缓存保证、图片加载与内存缓存等方法，均带中文注释便于理解与维护。

        /// <summary>
        /// 生成默认预览图片（优先从嵌入资源加载，失败则生成占位图）
        /// </summary>
        private BitmapImage GetDefaultPreviewImage()
        {
            try
            {
                // 尝试从程序集资源加载常见的默认预览图
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.UriSource = new Uri("pack://application:,,,/GB_NewCadPlus_IV;component/Resources/default_preview.png", UriKind.Absolute);
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                bmp.EndInit();
                bmp.Freeze();
                return bmp;
            }
            catch
            {
                try
                {
                    // 兜底尝试另一个资源名
                    var bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.UriSource = new Uri("pack://application:,,,/GB_NewCadPlus_IV;component/Resources/no_preview.png", UriKind.Absolute);
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                    bmp.EndInit();
                    bmp.Freeze();
                    return bmp;
                }
                catch
                {
                    // 都失败则返回程序生成的占位图
                    return CreatePlaceholderImage();
                }
            }
        }

        /// <summary>
        /// 创建简单的占位符图片（在无法加载资源时使用）
        /// </summary>
        private BitmapImage CreatePlaceholderImage()
        {
            try
            {
                // 创建内存位图并绘制简单文本作为占位符
                var rt = new RenderTargetBitmap(160, 120, 96, 96, PixelFormats.Pbgra32);
                var dv = new DrawingVisual();
                using (var dc = dv.RenderOpen())
                {
                    dc.DrawRectangle(Brushes.LightGray, new Pen(Brushes.Gray, 1), new System.Windows.Rect(0, 0, 160, 120));
                    var ft = new FormattedText("无预览",
                        System.Globalization.CultureInfo.CurrentCulture,
                        FlowDirection,
                        new Typeface("Arial"), 14, Brushes.Gray,
                        VisualTreeHelper.GetDpi(this).PixelsPerDip);
                    dc.DrawText(ft, new System.Windows.Point(40, 50));
                }
                rt.Render(dv);

                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(rt));
                using (var ms = new MemoryStream())
                {
                    encoder.Save(ms);
                    ms.Position = 0;
                    var bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.StreamSource = ms;
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.EndInit();
                    bmp.Freeze();
                    return bmp;
                }
            }
            catch
            {
                // 最后兜底：返回空实例（调用方按需处理 null/空）
                return new BitmapImage();
            }
        }

        /// <summary>
        /// 将指定图元列表的 DWG 和预览图缓存到本地。
        /// </summary>
        private async Task CacheGraphicFilesAsync(List<FileStorage> files)
        {
            if (files == null || files.Count == 0) return;

            var tasks = files.Select(async file =>
            {
                await _cacheDownloadSemaphore.WaitAsync();
                try
                {
                    // 内部已有"已缓存则跳过"逻辑 CacheGraphicFilesAsync
                    await ServerFileService.EnsureDwgCacheAsync(file, GetPath.DwgCachePath);
                    await ServerFileService.EnsurePreviewCacheAsync(file, GetPath.PreviewCachePath);
                    LogManager.Instance.LogInfo($"缓存完成: {file.DisplayName}");
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogInfo($"缓存失败 {file.DisplayName}: {ex.Message}");
                }
                finally
                {
                    _cacheDownloadSemaphore.Release();
                }
            });

            await Task.WhenAll(tasks);
        }

        /// <summary>
        /// 清理内存中的无效图片缓存项
        /// </summary>
        private void CleanupInvalidImageCache()
        {
            try
            {
                var toRemove = new List<string>();
                foreach (var kv in _imageCache)
                {
                    try
                    {
                        if (kv.Value == null || kv.Value.Width <= 0 || kv.Value.Height <= 0)
                            toRemove.Add(kv.Key);
                    }
                    catch
                    {
                        toRemove.Add(kv.Key);
                    }
                }
                foreach (var k in toRemove) _imageCache.Remove(k);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"清理图片缓存时出错: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 从指定路径加载图片（根据扩展名选择合理的加载方式）
        /// </summary>
        private BitmapImage LoadImageFromFile(string imagePath)
        {
            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                return null;

            try
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.UriSource = new Uri(imagePath, UriKind.Absolute);
                bmp.CacheOption = BitmapCacheOption.OnLoad;   // 加载后立即释放文件，避免锁定
                bmp.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                bmp.EndInit();
                bmp.Freeze();   // 允许跨线程访问
                return bmp;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"图片加载失败 [{imagePath}]: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// 显示文件预览到 UI（会设置预览控件的 Source）
        /// </summary>
        private async Task<bool> ShowFilePreviewAsync(FileStorage fileStorage)
        {
            try
            {
                if (预览 == null) return false; // 界面控件不存在

                预览.Source = null;

                if (fileStorage == null) return false;

                // 从缓存或回源获取预览图（本方法返回 BitmapImage）
                var bmp = await GetPreviewImageAsync(fileStorage).ConfigureAwait(true);
                if (bmp != null)
                {
                    预览.Source = bmp;
                    return true;
                }

                // 回退占位图
                预览.Source = GetDefaultPreviewImage();
                return false;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"显示文件预览时出错: {ex.Message}");
                try { 预览.Source = GetDefaultPreviewImage(); } catch { }
                return false;
            }
        }

        /// <summary>
        /// 尝试显示预览，若首次失败则再 EnsureLocalPreviewCacheAsync 后重试一次
        /// </summary>
        private async Task<bool> TryShowFilePreviewWithRetryAsync(FileStorage fileStorage)
        {
            if (fileStorage == null) return false;

            if (await ShowFilePreviewAsync(fileStorage).ConfigureAwait(true)) return true;

            try
            {
                // 若首次显示失败，尝试确保本地缓存后再次显示
                await EnsureLocalPreviewCacheAsync(fileStorage).ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"预览图重试前缓存失败: {ex.Message}");
            }

            return await ShowFilePreviewAsync(fileStorage).ConfigureAwait(true);
        }

        /// <summary>
        /// 确保 DWG 文件在本地有缓存副本并返回本地路径（用于插入/拖拽等）
        /// 逻辑：
        ///  - 如果 FilePath 存在并可访问 => 返回
        ///  - 否则若 FileBytes 有内容 => 写入临时文件并返回
        ///  - 否则如有 DatabaseManager 且 Id>0 => 从 DB 获取最新记录并重试
        /// </summary>
        private async Task<string> EnsureLocalCachedFilePathAsync(FileStorage fileStorage)
        {
            // 内部方法已包含“有则返回，无则下载”的逻辑
            return await ServerFileService.EnsureDwgCacheAsync(fileStorage, GetPath.DwgCachePath);
        }
        /// <summary>
        /// 确保预览图的本地缓存存在并返回本地路径
        /// - 如果原始路径是 UNC/可访问文件 -> 复制到 _previewCachePath 并返回本地缓存路径（按 FileId_ 前缀避免冲突）
        /// - 如果本地缓存已存在则直接返回
        /// - 如果未命中且数据库可用，会从数据库回源后再次尝试解析
        /// </summary>
        private async Task<string?> EnsureLocalPreviewCacheAsync(FileStorage fileStorage)
        {
            // 内部方法已包含“有则返回，无则下载”的逻辑
            return await ServerFileService.EnsurePreviewCacheAsync(fileStorage, GetPath.PreviewCachePath);
        }

        private System.Windows.VerticalAlignment GetVerticalAlignment()
        {
            return VerticalAlignment;
        }
        #endregion

        #region 第三阶段：按钮与面板加载、动态按钮交互、分类/面板加载相关方法（带中文注释，直接替换相应区域）

        /// <summary>
        /// 为单个 FileStorage 创建按钮（显示名称、绑定 Tag、并注册拖拽/点击事件）
        /// </summary>
        private Button CreateFileButton(FileStorage file)
        {
            // 中文注释：创建用于在 WrapPanel/StackPanel 中显示的按钮
            string caption = file?.DisplayName ?? string.Empty;
            var btn = new Button
            {
                Content = caption,
                Width = 88,
                Height = 22,
                Margin = new Thickness(0, 0, 5, 0),
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = System.Windows.VerticalAlignment.Top,   // ✅ 使用完全限定名
                Tag = new ButtonTagCommandInfo
                {
                    Type = "FileStorage",
                    ButtonName = caption,
                    fileStorage = file
                },
                Background = Brushes.Azure
            };

            // 事件：单击显示属性/预览；拖拽双用途（单击+拖动或双击）
            btn.Click += DynamicButton_Click;
            btn.PreviewMouseLeftButtonDown += DynamicButton_PreviewMouseLeftButtonDown;
            btn.PreviewMouseMove += DynamicButton_PreviewMouseMove;
            btn.PreviewMouseLeftButtonUp += DynamicButton_PreviewMouseLeftButtonUp;

            return btn;
        }

        /// <summary>
        /// 动态按钮单击处理：优先使用本地缓存；缺失时再通过服务器下载。
        /// </summary>
        private async void DynamicButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                VariableDictionary.wpfWindows_Status = true;
                if (!(sender is Button btn)) return; // 仅处理 Button 类型的点击事件

                // 1. 视觉选中样式管理
                if (_lastSelectedDynamicButton != null && _lastSelectedDynamicButton != btn)
                {
                    if (_originalButtonBackgrounds.TryGetValue(_lastSelectedDynamicButton, out var prev)) // 恢复上一个按钮的原始背景色
                        _lastSelectedDynamicButton.Background = prev; // 恢复原始背景色
                    else
                        _lastSelectedDynamicButton.ClearValue(Control.BackgroundProperty); // 如果没有原始背景记录，则清除设置（回退到默认样式）
                }
                if (!_originalButtonBackgrounds.ContainsKey(btn)) // 记录当前按钮的原始背景色（如果尚未记录过）
                {
                    _originalButtonBackgrounds[btn] = btn.Background ?? SystemColors.ControlBrush;
                }
                btn.Background = Brushes.LightGoldenrodYellow;// 设置当前按钮的背景色以示选中,颜色为亮黄色，确保与默认的浅灰色有明显区分
                _lastSelectedDynamicButton = btn;

                // 2. 解析 FileStorage 对象
                var file = ResolveFileStorageFromTag(btn.Tag);
                if (file == null)
                {
                    LogManager.Instance.LogWarning("无法解析图元信息，操作中止");
                    return;
                } else if (file != null)
                {
                    LogManager.Instance.LogInfo($"[点击] 解析文件 文件名={file.DisplayName}, Id={file.Id}");
                }
                LogManager.Instance.LogInfo($"[点击] 处理图元: {file.FileName}, 原始路径: {file.FilePath}");

                // 3. 在线模式：优先使用本地缓存
                if (_useDatabaseMode && _databaseManager != null)
                {

                    // 3.1 获取 DWG 本地缓存路径（内部已包含“有则返回，无则下载”）
                    string dwgCache = await EnsureLocalCachedFilePathAsync(file);
                    if (!string.IsNullOrEmpty(dwgCache))
                    {
                        LogManager.Instance.LogInfo($"DWG 本地缓存就绪: {dwgCache},文件名: {file.FileName}");
                        _currentFileStorage = file;// 设置当前文件储存对象（供属性显示等使用）
                        _selectedFileStorage = file;// 设置选中文件储存对象（供属性显示等使用）
                    }
                    else
                    {
                        LogManager.Instance.LogWarning("无法获取 DWG 本地缓存，可能网络异常或文件不存在");
                    }
                    // 3.2 显示预览
                    var previewShown = await TryShowFilePreviewWithRetryAsync(file);
                    if (!previewShown)
                    {
                        // 额外检查：可能预览控件已显示占位图，但不影响使用
                        if (预览.Source != null)
                            LogManager.Instance.LogInfo("预览图已加载（可能为占位图）");
                        else
                            LogManager.Instance.LogWarning("预览图加载失败且无占位图");
                    }
                    // 3.3 显示属性
                    await DisplayFilePropertiesInDataGridAsync(file).ConfigureAwait(true);
                }
                else
                {
                    // 4. 离线模式：仅尝试显示预览并使用原始路径
                    await TryShowFilePreviewWithRetryAsync(file);
                    _currentFileStorage = file;
                    _selectedFileStorage = file;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError("DynamicButton_Click 异常: " + ex.Message);
            }
        }

        /// <summary>
        /// 处理按钮按下（用于拖拽判定或双击触发插入）
        /// </summary>
        private void DynamicButton_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (e.ClickCount == 2) //如果是双击事件，直接触发双击处理逻辑
                {
                    // 双击：触发双击处理逻辑并阻止后续拖拽
                    _isButtonMouseDown = false;
                    _isButtonDragging = false;
                    _dragSourceButton = null;
                    DynamicButton_MouseDoubleClick(sender, e);// 执行双击处理方法
                    e.Handled = true;
                }
                else
                {
                    // 单次按下：记录起始点以便拖拽检测
                    _isButtonMouseDown = true;
                    _isButtonDragging = false;
                    _buttonDragStartPoint = e.GetPosition(null);
                    _dragSourceButton = sender as Button;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError("PreviewMouseLeftButtonDown 处理失败: " + ex.Message);
            }
        }

        /// <summary>
        /// 双击按钮处理：确保本地缓存并执行插入命令
        /// </summary>
        private async void DynamicButton_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                VariableDictionary.wpfWindows_Status = true;
                // 解析文件储存信息
                var file = ResolveFileStorageFromTag((sender as Button)?.Tag);
                if (file == null) return;
                // 确保 DWG 文件在本地有缓存副本并获取有效路径（内部已包含“有则返回，无则下载”的逻辑）
                var validPath = await EnsureLocalCachedFilePathAsync(file);
                if (string.IsNullOrEmpty(validPath))
                {
                    LogManager.Instance.LogWarning("无法获取有效的本地图元文件路径，中止操作");
                    return;
                }

                VariableDictionary.btnFileName = file.FileName;
                //GetPath._cacheStoragePath = validPath;
                var (ok, err) = await ExecuteInsertAndWaitResultAsync(validPath);// 内部已包含插入命令执行和结果等待逻辑
                if (!ok) LogManager.Instance.LogWarning("插入失败: " + err);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError("双击处理出错: " + ex.Message);
            }
        }

        /// <summary>
        /// 鼠标移动事件：判断为拖拽（超过最小拖拽距离）时触发插入流程
        /// </summary>
        private async void DynamicButton_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            try
            {
                VariableDictionary.wpfWindows_Status = true;
                if (!_isButtonMouseDown || _isButtonDragging) return; // 只有在鼠标左键按下且尚未判定为拖拽时才进行拖拽检测
                var diff = e.GetPosition(null) - _buttonDragStartPoint; // 计算当前鼠标位置与起始点的差值
                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)  // 如果水平或垂直方向的移动都超过系统定义的最小拖拽距离，则判定为拖拽
                {
                    _isButtonDragging = true; // 设置拖拽状态，避免重复触发
                    if (sender is Button btn) // 确保 sender 是 Button 类型
                    {
                        var file = ResolveFileStorageFromTag(btn.Tag); // 解析按钮关联的 FileStorage 对象
                        if (file != null) // 如果成功解析出文件信息，则继续处理拖拽插入逻辑
                        {
                            var localPath = await EnsureLocalCachedFilePathAsync(file); // 确保 DWG 文件在本地有缓存副本并获取路径（内部已包含“有则返回，无则下载”的逻辑）
                            if (!string.IsNullOrWhiteSpace(localPath)) // 如果成功获取到有效的本地文件路径，则执行插入命令并等待结果（内部已包含插入命令执行和结果等待逻辑）
                            {
                                var (ok, err) = await ExecuteInsertAndWaitResultAsync(localPath); // 内部已包含插入命令执行和结果等待逻辑
                                if (!ok) LogManager.Instance.LogWarning("插入失败: " + err);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError("拖拽处理出错: " + ex.Message);
            }
        }

        /// <summary>
        /// 鼠标左键抬起：重置拖拽状态
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
                LogManager.Instance.LogError("PreviewMouseLeftButtonUp 处理失败: " + ex.Message);
            }
        }

        /// <summary>
        /// TabControl 选择改变处理（打开分类面板或加载 CSV 表）
        /// </summary>
        // 异步处理 TabControl 选择变更事件
        private async void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                // 记录事件触发日志
                LogManager.Instance.LogInfo("TabControl选择改变事件触发");

                // 如果没有新增的选中项，则直接返回
                if (e.AddedItems.Count == 0) return;

                // 获取新增选中项并尝试转换为 TabItem
                var added = e.AddedItems[0] as TabItem;
                if (added == null) return;

                // 获取当前选中 Tab 的标题（去除首尾空白）
                var header = (added.Header?.ToString() ?? string.Empty).Trim();
                LogManager.Instance.LogInfo("选中的TabItem: " + header);

                // 如果选中的是“计算数据表”选项卡
                if (string.Equals(header, "计算数据表", StringComparison.OrdinalIgnoreCase))
                {
                    // 若动态宿主容器存在且尚无子元素，则重新加载计算数据表（不强制刷新）
                    if (CalcDynamicHost != null && CalcDynamicHost.Children.Count == 0)
                    {
                        try { ReloadCalcCsvTables(false); }
                        catch (Exception ex) { LogManager.Instance.LogWarning("打开计算数据表时加载失败: " + ex.Message); }
                    }
                }
                // 如果选中的是主分类选项卡（工艺、建筑、结构、电气、给排水、暖通、自控、总图、公共图）
                else if (new[] { "工艺", "建筑", "结构", "电气", "给排水", "暖通", "自控", "总图", "公共图" }.Contains(header))
                {
                    // 异步等待该主分类下的动态按钮加载完成，避免异步任务尚未执行就显示空面板
                    await LoadButtonsForMainCategoryTabAsync(added, header);

                    // 如果是“工艺”页，还需要单独等待条件图元面板的按钮加载完成
                    if (header == "工艺")
                    {
                        await LoadConditionButtonsAsync();
                    }
                }
                // 如果选中的是“图元集”或“图层管理”等子选项卡（通常位于某个主分类下）
                else if (header.Contains("图元集") || header.Contains("图层管理"))
                {
                    // 查找所属的父级 TabItem（即主分类选项卡）
                    var parent = FindParentTabItem(added);
                    if (parent != null)
                    {
                        // 获取父级标题作为分类名称
                        var categoryName = (parent.Header?.ToString() ?? string.Empty).Trim();
                        // 同步加载该主分类下的按钮（因为此处可能是子选项卡切换，不需要异步等待）
                        LoadButtonsForMainCategoryTab(parent, categoryName);
                    }
                }
                // 其他情况不做特殊处理
            }
            catch (Exception ex)
            {
                // 捕获任何异常并记录错误日志
                LogManager.Instance.LogError("处理TabControl选择改变时出错: " + ex.Message);
            }
        }

        /// <summary>
        /// 异步加载主分类按钮，保证分类数据准备完成后再创建界面按钮。
        /// </summary>
        private async Task LoadButtonsForMainCategoryTabAsync(TabItem tabItem, string categoryName)
        {
            try
            {
                // 分类树尚未完成初始化时，先从服务器重新加载一次分类数据。
                if (_categoryTreeNodes == null || _categoryTreeNodes.Count == 0)
                    await InitializeCategoryTreeAsync();

                var panel = GetPanelByFolderName(categoryName);// 根据分类名称获取对应的 WrapPanel
                if (panel == null)
                {
                    LogManager.Instance.LogWarning($"未找到分类 {categoryName} 对应的按钮面板。");
                    return;
                }

                // 清除旧按钮，避免重复加载和旧分类按钮残留。
                panel.Children.Clear();
                await LoadButtonsFromDatabaseForCategory(categoryName, panel);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"异步加载分类 {categoryName} 按钮失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 为主分类 Tab 加载按钮（根据数据库/资源选择加载源）
        /// </summary>
        private void LoadButtonsForMainCategoryTab(TabItem tabItem, string categoryName)
        {
            try
            {
                LogManager.Instance.LogInfo($"开始为分类 {categoryName} 加载按钮");
                var panel = GetPanelByFolderName(categoryName);
                if (panel == null)
                {
                    LogManager.Instance.LogInfo($"未找到 {categoryName} 对应的面板");
                    return;
                }
                panel.Children.Clear();

                LogManager.Instance.LogInfo("使用服务器模式加载 " + categoryName);
                _ = LoadButtonsFromDatabaseForCategory(categoryName, panel);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"为分类 {categoryName} 加载按钮时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 从数据库加载指定分类的按钮到 WrapPanel 面板（包含处理子分类与直达文件）
        /// </summary>
        /// <param name="categoryName">分类名称（如“工艺”、“建筑”等）</param>
        /// <param name="panel">承载按钮的 WrapPanel 面板</param>
        private async Task LoadButtonsFromDatabaseForCategory(string categoryName, WrapPanel panel)
        {
            try
            {
                // 记录开始加载日志
                LogManager.Instance.LogInfo($"=== 开始从数据库加载分类 {categoryName} ===");

                // 在缓存的分类树节点中查找一级节点（Level == 0），匹配名称或显示文本（忽略大小写）
                var categoryNode = FindMainCategoryNode(categoryName);

                // 获取节点中存储的业务数据对象（CadCategory）
                var category = categoryNode?.Data as CadCategory;

                // 若未找到该分类，则回退使用本地资源加载按钮
                if (category == null)
                {
                    LogManager.Instance.LogInfo("服务器分类树中未找到分类: " + categoryName);
                    LoadButtonsFromResources(categoryName, panel);
                    return;
                }

                // 获取该分类下的所有子分类（子节点数据为 CadSubcategory）
                var subcategories = categoryNode.Children
                    .Select(n => n.Data)
                    .OfType<CadSubcategory>()
                    .ToList();

                // 清空面板现有内容
                panel.Children.Clear();

                // 如果没有子分类，则直接加载该分类下的文件（按钮）
                if (subcategories.Count == 0)
                {
                    await LoadFilesDirectlyForCategory(category, panel);
                }
                else
                {
                    // 否则按子分类分组加载文件
                    await LoadFilesBySubcategories(category, subcategories, panel);
                }

                // 记录加载完成日志
                LogManager.Instance.LogInfo($"=== 完成加载分类 {categoryName} ===");
            }
            catch (Exception ex)
            {
                // 若发生异常，记录错误并回退使用本地资源加载
                LogManager.Instance.LogInfo($"从数据库加载分类 {categoryName} 时出错: {ex.Message}");
                LoadButtonsFromResources(categoryName, panel);
            }
        }

        /// <summary>
        /// 直接加载主分类下未分子分类的文件
        /// </summary>
        private async Task LoadFilesDirectlyForCategory(CadCategory category, WrapPanel panel)
        {
            try
            {
                
                 var files = await _graphicApiService.GetFilesByCategoryIdAsync(category.Id, "main");
               
                if (files.Count > 0)
                {
                    var sorted = files.OrderBy(f => f.DisplayName).ToList();
                    CreateFileButtonsForPanel(sorted, panel, category.DisplayName);
                }
                else
                {
                    ShowNoFilesMessage(panel, "暂无文件");
                }
                // ★ 新增：后台缓存该分类所有图元
                _ = CacheGraphicFilesAsync(files);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("直接加载分类文件时出错: " + ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 按子分类加载文件（每个子分类创建 section Border）
        /// </summary>
        private async Task LoadFilesBySubcategories(CadCategory category, List<CadSubcategory> subcategories, WrapPanel panel)
        {
            try
            {
                var allFiles = new List<FileStorage>();   // 收集所有文件
                var bgColors = new List<System.Windows.Media.Color> { Colors.FloralWhite, Colors.Azure, Colors.FloralWhite, Colors.Azure };
                int colorIndex = 0;
                foreach (var sub in subcategories.OrderBy(s => s.SortOrder))
                {
                    var files = await _graphicApiService.GetFilesByCategoryIdAsync(sub.Id, "sub");
                    var section = CreateSubcategorySection(sub.DisplayName, bgColors[colorIndex % bgColors.Count]);
                    var host = section.Child as StackPanel;
                    if (files.Count > 0)
                    {
                        var sorted = files.OrderBy(f => f.DisplayName).ToList();
                        CreateFileButtonsForPanel(sorted, host, sub.DisplayName);
                    }
                    else
                    {
                        ShowNoFilesMessage(host, "暂无文件");
                    }
                    panel.Children.Add(section);
                    colorIndex++;
                    // ... 创建 UI section ...
                    allFiles.AddRange(files);
                }
                // ★ 新增：后台缓存该分类所有图元
                _ = CacheGraphicFilesAsync(allFiles);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("按子分类加载文件时出错: " + ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 创建子分类分组区域（带标题）
        /// </summary>
        private Border CreateSubcategorySection(string title, System.Windows.Media.Color backgroundColor)
        {
            var border = new Border
            {
                BorderBrush = new SolidColorBrush(Colors.Gray),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(5),
                Margin = new Thickness(0, 2, 0, 2),
                Width = 300,
                Background = new SolidColorBrush(backgroundColor),
                HorizontalAlignment = HorizontalAlignment.Left
            };

            var sp = new StackPanel { Margin = new Thickness(3) };
            var tb = new TextBlock
            {
                Text = title,
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 2),
                Foreground = new SolidColorBrush(Colors.DarkBlue)
            };
            sp.Children.Add(tb);
            border.Child = sp;
            return border;
        }

        /// <summary>
        /// 为目标面板创建文件按钮（按每行3列布局）
        /// </summary>
        private void CreateFileButtonsForPanel(List<FileStorage> files, Panel targetPanel, string sectionName)
        {
            try
            {
                int perRow = 3;
                for (int i = 0; i < files.Count; i += perRow)
                {
                    var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 2) };
                    for (int j = 0; j < perRow && i + j < files.Count; j++)
                    {
                        var btn = CreateFileButton(files[i + j]);   // 不再传 GetVerticalAlignment()
                        row.Children.Add(btn);
                    }
                    targetPanel.Children.Add(row);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("创建文件按钮时出错: " + ex.Message);
            }
        }

        /// <summary>
        /// 为按钮绑定常用事件并设置样式（避免重复绑定）
        /// </summary>
        private void AttachDynamicButtonHandlers(Button btn)
        {
            if (btn == null) return;
            bool attached = false;
            try { attached = (bool)btn.GetValue(HandlersAttachedProperty); } catch { attached = false; }
            if (attached) return;

            if (!(btn.Tag is ButtonTagCommandInfo tag) || !string.Equals(tag.Type, "Predefined", StringComparison.OrdinalIgnoreCase))
            {
                btn.Click += DynamicButton_Click;
                try
                {
                    var bg = btn.Background;
                    if (bg == null) btn.Background = Brushes.Azure;
                }
                catch { }
            }
            btn.PreviewMouseLeftButtonDown += DynamicButton_PreviewMouseLeftButtonDown;
            btn.PreviewMouseMove += DynamicButton_PreviewMouseMove;
            btn.PreviewMouseLeftButtonUp += DynamicButton_PreviewMouseLeftButtonUp;
            btn.SetValue(HandlersAttachedProperty, true);
        }

        /// <summary>
        /// 根据按钮 Tag 解析 FileStorage 对象（支持三种形态：FileStorage，string 路径，ButtonTagCommandInfo）
        /// </summary>
        private FileStorage ResolveFileStorageFromTag(object tag)
        {
            VariableDictionary.entityRotateAngle = 0;// 每次解析图元信息时重置旋转角度，避免旧值干扰新操作
            switch (tag)// 根据 Tag 的实际类型进行解析，优先支持直接存储 FileStorage 对象的方式，兼容旧版可能使用字符串路径的方式，以及封装在 ButtonTagCommandInfo 中的方式
            {
                // 封装在 ButtonTagCommandInfo 中的方式（兼容性强，包含类型区分和额外信息）
                case ButtonTagCommandInfo info:
                    if (info.fileStorage != null) return info.fileStorage; // 优先使用已封装的 FileStorage 对象
                    break;
            }
            return null;
        }

        /// <summary>
        /// 根据页面主分类标题查找分类树节点，兼容数据库分类名称中的数字编号前缀。
        /// </summary>
        private CategoryTreeNode? FindMainCategoryNode(string categoryName)
        {
            if (_categoryTreeNodes == null || string.IsNullOrWhiteSpace(categoryName))
                return null;

            string NormalizeCategoryName(string value)
            {
                var normalized = (value ?? string.Empty).Trim();
                normalized = normalized.Replace('－', '-').Replace('—', '-').Replace('‐', '-');
                normalized = Regex.Replace(normalized, @"^[0-9０-９]+\s*[-_\.、\s]+", string.Empty);
                return normalized.Trim();
            }

            string requestedName = NormalizeCategoryName(categoryName);
            return _categoryTreeNodes.FirstOrDefault(n =>
                n.Level == 0 &&
                (string.Equals(n.Name?.Trim(), categoryName.Trim(), StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(n.DisplayText?.Trim(), categoryName.Trim(), StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(NormalizeCategoryName(n.Name ?? string.Empty), requestedName, StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(NormalizeCategoryName(n.DisplayText ?? string.Empty), requestedName, StringComparison.OrdinalIgnoreCase)));
        }

        /// <summary>
        /// 根据主分类名字返回对应的 WrapPanel 引用（从 XAML 成员中查找）
        /// </summary>
        private WrapPanel GetPanelByFolderName(string folderName)
        {
            LogManager.Instance.LogInfo("查找面板: " + folderName);
            switch (folderName)
            {
                case "公用工具": return PublicButtonsPanel;
                case "工艺": return CraftButtonsPanel;
                case "建筑": return ArchitectureButtonsPanel;
                case "总图": return GeneralButtonsPanel;
                case "暖通": return HVACButtonsPanel;
                case "电气": return ElectricalButtonsPanel;
                case "结构": return StructureButtonsPanel;
                case "给排水": return PlumbingButtonsPanel;
                case "自控": return ControlButtonsPanel;
                default: return null;
            }
        }

        /// <summary>
        /// 从数据库加载分类下的按钮（更通用的批量接口，已在其它方法中调用）
        /// </summary>
        /// <summary>
        /// 从服务器数据库异步加载指定文件夹（分类）下的按钮，并按子分类分组展示到 WrapPanel 面板中
        /// </summary>
        /// <param name="folderName">分类名称（如“工艺”、“建筑”等）</param>
        /// <param name="panel">承载按钮的 WrapPanel 面板</param>
        private async Task LoadButtonsFromDatabase(string folderName, WrapPanel panel)
        {
            try
            {
                // 记录开始加载日志
                LogManager.Instance.LogInfo($"开始从服务器加载分类 {folderName} 的按钮");

                // 在缓存的分类树节点中查找一级节点（Level == 0），匹配名称或显示文本（忽略大小写）
                var categoryNode = FindMainCategoryNode(folderName);

                // 获取节点中存储的业务数据对象（CadCategory）
                var category = categoryNode?.Data as CadCategory;

                // 若未找到该分类，则记录日志并返回（不做任何加载）
                if (category == null)
                {
                    LogManager.Instance.LogInfo("服务器分类树中未找到分类: " + folderName);
                    return;
                }

                // 获取该分类下的所有子分类（子节点数据为 CadSubcategory）
                var subcategories = categoryNode.Children.Select(n => n.Data).OfType<CadSubcategory>().ToList();

                // 定义一组背景颜色，用于交替显示子分类区块
                var bgColors = new List<System.Windows.Media.Color> { Colors.FloralWhite, Colors.Azure, Colors.FloralWhite };
                int colorIndex = 0;

                // 遍历每个子分类，为每个子分类生成一个独立的区块（Border）
                foreach (var sub in subcategories)
                {
                    // 创建边框容器，作为子分类的卡片样式
                    var border = new Border
                    {
                        BorderBrush = new SolidColorBrush(Colors.Gray),
                        BorderThickness = new Thickness(1),
                        CornerRadius = new CornerRadius(5),
                        Margin = new Thickness(0, 2, 0, 2),
                        Width = 300,
                        Background = new SolidColorBrush(bgColors[colorIndex % bgColors.Count]),
                        HorizontalAlignment = HorizontalAlignment.Left
                    };

                    // 内部垂直面板，用于放置标题和按钮行
                    var sectionPanel = new StackPanel { Margin = new Thickness(3) };

                    // 子分类标题（显示 DisplayName）
                    var header = new TextBlock
                    {
                        Text = sub.DisplayName,
                        FontSize = 14,
                        FontWeight = FontWeights.Bold,
                        Margin = new Thickness(0, 0, 0, 2),
                        Foreground = new SolidColorBrush(Colors.DarkBlue)
                    };
                    sectionPanel.Children.Add(header);

                    // 异步调用 API 获取该子分类下的所有图形文件
                    var graphics = await _graphicApiService.GetFilesByCategoryIdAsync(sub.Id, "sub");

                    // 如果存在文件，则按显示名称排序，并分列显示按钮
                    if (graphics.Count > 0)
                    {
                        // 按显示名称排序
                        graphics.Sort((x, y) => x.DisplayName.CompareTo(y.DisplayName));

                        // 每行显示3个按钮
                        int cols = 3;
                        for (int i = 0; i < graphics.Count; i += cols)
                        {
                            // 创建水平行面板
                            var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 2) };

                            // 在当前行中生成最多 cols 个按钮
                            for (int j = 0; j < cols && i + j < graphics.Count; ++j)
                            {
                                var g = graphics[i + j];
                                string btnName = g.DisplayName;

                                // 处理按钮显示名称：若包含下划线，则截取下划线后的部分作为显示名
                                if (!string.IsNullOrWhiteSpace(btnName))
                                {
                                    int idx = btnName.LastIndexOf('_');
                                    btnName = idx < 0 || idx + 1 >= btnName.Length ? btnName.Trim() : btnName.Substring(idx + 1).Trim();
                                }

                                // 创建按钮，设置宽度、高度、边距，并将文件对象存入 Tag 以备后续使用
                                var button = new Button
                                {
                                    Content = btnName,
                                    Width = 88,
                                    Height = 22,
                                    Margin = new Thickness(0, 0, 5, 0),
                                    Tag = new ButtonTagCommandInfo { Type = "FileStorage", ButtonName = btnName, fileStorage = g }
                                };

                                // 附加动态按钮的点击事件等处理逻辑
                                AttachDynamicButtonHandlers(button);

                                // 将按钮添加到当前行
                                row.Children.Add(button);
                            }

                            // 将该行添加到子分类面板
                            sectionPanel.Children.Add(row);
                        }
                    }
                    else
                    {
                        // 若无文件，则显示“暂无文件”提示
                        var no = new TextBlock
                        {
                            Text = "暂无文件",
                            FontSize = 12,
                            Margin = new Thickness(5, 0, 0, 3),
                            Foreground = new SolidColorBrush(Colors.Gray)
                        };
                        sectionPanel.Children.Add(no);
                    }

                    // 将组装好的面板放入边框，并将边框添加到主面板中
                    border.Child = sectionPanel;
                    panel.Children.Add(border);

                    // 更新背景颜色索引，使不同子分类区块交替颜色
                    colorIndex++;
                }
            }
            catch (Exception ex)
            {
                // 记录异常并重新抛出
                LogManager.Instance.LogInfo("从数据库加载按钮时出错: " + ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 从 Resources 文件夹加载按钮（当数据库不可用时回退）
        /// </summary>
        private void LoadButtonsFromResources(string folderName, WrapPanel panel)
        {
            try
            {
                var baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var path = Path.Combine(baseDir, "Resources", folderName);
                if (!Directory.Exists(path))
                {
                    MessageBox.Show($"找不到资源文件夹: {path}\n请检查Resources文件夹中的'{folderName}'文件夹是否存在");
                    return;
                }

                var colorList = new List<System.Windows.Media.Color> { Colors.FloralWhite, Colors.Azure, Colors.FloralWhite };
                int idx = 0;
                var directories = Directory.GetDirectories(path);
                foreach (var dir in directories)
                {
                    var name = Path.GetFileName(dir);
                    var border = new Border { BorderBrush = new SolidColorBrush(Colors.Gray), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(5), Margin = new Thickness(0, 2, 0, 3), Width = 282, Background = new SolidColorBrush(colorList[idx % colorList.Count]), HorizontalAlignment = HorizontalAlignment.Left };
                    var sp = new StackPanel { Margin = new Thickness(5) };
                    var header = new TextBlock { Text = name, FontSize = 12, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 5, 0, 5), Foreground = new SolidColorBrush(Colors.DarkBlue) };
                    sp.Children.Add(header);

                    var files = Directory.GetFiles(dir, "*.dwg");
                    if (files.Length > 0)
                    {
                        var tuples = new List<Tuple<string, string>>();
                        foreach (var f in files)
                        {
                            var withoutExt = Path.GetFileNameWithoutExtension(f);
                            var label = withoutExt.Contains("_") ? withoutExt.Substring(withoutExt.IndexOf("_") + 1) : withoutExt;
                            var chinese = ExtractChineseCharacters(label);
                            var display = string.IsNullOrEmpty(chinese) ? label : chinese;
                            tuples.Add(Tuple.Create(display, f));
                        }
                        tuples.Sort((x, y) => x.Item1.CompareTo(y.Item1));

                        int cols = 3;
                        for (int i = 0; i < tuples.Count; i += cols)
                        {
                            var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 5) };
                            for (int j = 0; j < cols && i + j < tuples.Count; ++j)
                            {
                                var t = tuples[i + j];
                                string cmdName = t.Item1;
                                var filePath = t.Item2;
                                if (!string.IsNullOrWhiteSpace(cmdName))
                                {
                                    int last = cmdName.LastIndexOf('_');
                                    cmdName = last < 0 || last + 1 >= cmdName.Length ? cmdName.Trim() : cmdName.Substring(last + 1).Trim();
                                }
                                var btn = new Button { Content = cmdName, Width = 88, Height = 20, FontSize = 12, FontFamily = new FontFamily("微软雅黑"), Margin = new Thickness(0, 0, 3, 0), Tag = new ButtonTagCommandInfo { Type = "File", ButtonName = cmdName, FilePath = filePath } };
                                AttachDynamicButtonHandlers(btn);
                                row.Children.Add(btn);
                            }
                            sp.Children.Add(row);
                        }
                    }
                    else
                    {
                        var no = new TextBlock { Text = "暂无文件", FontSize = 12, Margin = new Thickness(5, 0, 0, 5), Foreground = new SolidColorBrush(Colors.Gray) };
                        sp.Children.Add(no);
                    }

                    border.Child = sp;
                    panel.Children.Add(border);
                    idx++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载按钮时出错: {ex.Message}\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// 加载“条件图元”页签的按钮（工艺专用）
        /// </summary>
        private async Task LoadConditionButtonsAsync()
        {
            try
            {
                LogManager.Instance.LogInfo("开始加载条件图元按钮...");
                ClearConditionButtons();
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
                LogManager.Instance.LogInfo("加载条件图元按钮时出错: " + ex.Message);
                MessageBox.Show("加载条件图元按钮时出错: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        /// <summary>
        /// 兼容旧代码调用条件图元按钮加载方法。
        /// </summary>
        private void LoadConditionButtons()
        {
            _ = LoadConditionButtonsAsync();
        }

        /// <summary>
        /// 为指定专业加载条件文件按钮（可从 DB 或资源获取）
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

                var conditionFiles = await GetConditionFilesForSpecialty(专业名称);
                if (conditionFiles.Count == 0)
                {
                    AddNoFilesLabel(targetPanel, $"暂无{专业名称}条件文件");
                    return;
                }

                int cols = 3;
                for (int i = 0; i < conditionFiles.Count; i += cols)
                {
                    var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 5) };
                    for (int j = 0; j < cols && i + j < conditionFiles.Count; ++j)
                    {
                        var file = conditionFiles[i + j];
                        var btn = CreateConditionButton(file);
                        row.Children.Add(btn);
                    }
                    targetPanel.Children.Add(row);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"加载{专业名称}条件按钮时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取某专业的条件文件列表（优先从 DB，否则回退到资源）
        /// </summary>
        private async Task<List<ConditionFileInfo>> GetConditionFilesForSpecialty(string specialtyName)
        {
            var result = new List<ConditionFileInfo>();
            try
            {
                if (_databaseManager == null) return GetConditionFilesFromResources(specialtyName);
                // TODO: 若数据库中存储专用条件文件，可在这里实现 DB 查询逻辑
                return GetConditionFilesFromResources(specialtyName);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"获取{specialtyName}条件文件时出错: {ex.Message}");
                return result;
            }
        }

        /// <summary>
        /// 从 Resources 创建条件文件的占位信息（当 DB 不可用时使用）
        /// </summary>
        private List<ConditionFileInfo> GetConditionFilesFromResources(string specialtyName)
        {
            var list = new List<ConditionFileInfo>();
            try
            {
                var basePack = $"pack://application:,,,/Resources/Conditions/{specialtyName}/";
                switch (specialtyName)
                {
                    case "电气":
                        list.Add(new ConditionFileInfo { Name = "电气条件1", DisplayName = "电气条件1", FilePath = basePack + "电气条件1.dwg" });
                        list.Add(new ConditionFileInfo { Name = "电气条件2", DisplayName = "电气条件2", FilePath = basePack + "电气条件2.dwg" });
                        list.Add(new ConditionFileInfo { Name = "电气条件3", DisplayName = "电气条件3", FilePath = basePack + "电气条件3.dwg" });
                        break;
                    case "自控":
                        list.Add(new ConditionFileInfo { Name = "自控条件1", DisplayName = "自控条件1", FilePath = basePack + "自控条件1.dwg" });
                        list.Add(new ConditionFileInfo { Name = "自控条件2", DisplayName = "自控条件2", FilePath = basePack + "自控条件2.dwg" });
                        break;
                    case "给排水":
                        list.Add(new ConditionFileInfo { Name = "给排水条件1", DisplayName = "给排水条件1", FilePath = basePack + "给排水条件1.dwg" });
                        list.Add(new ConditionFileInfo { Name = "给排水条件2", DisplayName = "给排水条件2", FilePath = basePack + "给排水条件2.dwg" });
                        list.Add(new ConditionFileInfo { Name = "给排水条件3", DisplayName = "给排水条件3", FilePath = basePack + "给排水条件3.dwg" });
                        break;
                    case "暖通":
                        list.Add(new ConditionFileInfo { Name = "暖通条件1", DisplayName = "暖通条件1", FilePath = basePack + "暖通条件1.dwg" });
                        break;
                    case "结构":
                        list.Add(new ConditionFileInfo { Name = "结构条件1", DisplayName = "结构条件1", FilePath = basePack + "结构条件1.dwg" });
                        list.Add(new ConditionFileInfo { Name = "结构条件2", DisplayName = "结构条件2", FilePath = basePack + "结构条件2.dwg" });
                        break;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"从资源获取{specialtyName}条件文件时出错: {ex.Message}");
            }
            return list;
        }

        /// <summary>
        /// 为条件文件创建按钮
        /// </summary>
        private Button CreateConditionButton(ConditionFileInfo fileInfo)
        {
            var label = fileInfo?.DisplayName ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(label))
            {
                int idx = label.LastIndexOf('_');
                label = idx < 0 || idx + 1 >= label.Length ? label.Trim() : label.Substring(idx + 1).Trim();
            }

            var btn = new Button
            {
                Content = label,
                Width = 85,
                Height = 20,
                Margin = new Thickness(5, 1, 1, 1),
                Tag = fileInfo,
                FontFamily = new FontFamily("Microsoft YaHei UI"),
                FontWeight = FontWeights.Normal,
            };
            try
            {
                btn.Style = (Style)FindResource("ButtonStyle");
            }
            catch { /* 资源缺失时使用默认样式 */ }
            btn.Click += ConditionButton_Click;
            return btn;
        }

        /// <summary>
        /// 条件按钮点击：直接在 CAD 中执行对应插入命令
        /// </summary>
        private void ConditionButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!(sender is Button btn) || !(btn.Tag is ConditionFileInfo info)) return;
                LogManager.Instance.LogInfo("点击条件按钮: " + info.DisplayName);
                ExecuteConditionInsert(info);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("执行条件插入时出错: " + ex.Message);
                MessageBox.Show("执行条件插入时出错: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        /// <summary>
        /// 执行条件插入（设置变量并发送命令到 CAD）
        /// </summary>
        private void ExecuteConditionInsert(ConditionFileInfo fileInfo)
        {
            try
            {
                VariableDictionary.btnFileName = fileInfo.Name;
                VariableDictionary.btnBlockLayer = "TJ(条件图元)";
                VariableDictionary.layerColorIndex = 7;
                Env.Document.SendStringToExecute("GB_InsertBlock ", false, false, false);
                LogManager.Instance.LogInfo("成功插入条件: " + fileInfo.DisplayName);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("插入条件失败: " + ex.Message);
                MessageBox.Show("插入条件失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        #endregion

        #region 第四阶段：用户/部门管理、导出/表格与剩余杂项方法（带中文注释）.说明：此段包含用户管理按钮处理、用户编辑对话、部门/用户刷新等逻辑，直接替换到类的相应区域。

        /// <summary>
        /// 确保认证服务 (_authServiceDynamic) 已初始化。
        /// </summary>
        /// <param name="host">服务器地址</param>
        /// <param name="port">端口</param>
        /// <param name="dbType">数据库类型 (MYSQL 或 DM)</param>
        /// <param name="dbUser">用户名</param>
        /// <param name="dbPwd">密码</param>
        /// <returns>如果服务初始化成功返回 true，否则返回 false</returns>
        private bool EnsureSvcInitialized(string host, int port, string dbType, string dbUser, string dbPwd)
        {
            return _apiService != null;
        }
        
        /// <summary>
        /// 加载某部门的用户列表并展示到 UsersGrid（异步可改为同步）
        /// </summary>
        private void LoadUsersForDepartment(int departmentId)
        {
            try
            {
                if (_apiService == null) return;
                UsersGrid.ItemsSource = _apiService.GetUsersAsync(departmentId).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"LoadUsersForDepartment 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 刷新部门树或列表（示例：重新读取 DepartmentsGrid 数据源）
        /// </summary>
        private async Task RefreshDepartmentsAsync()
        {
            try
            {
                if (_departmentApiService == null) return;
                var departments = await _departmentApiService.GetDepartmentsWithCountsAsync();

                // 3. 绑定到 Grid
                if (DepartmentsGrid != null)
                {
                    DepartmentsGrid.ItemsSource = departments;
                    DepartmentsGrid.Items.Refresh();
                }

                LogManager.Instance.LogInfo($"RefreshDepartmentsAsync: 成功加载 {departments.Count} 个部门");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"RefreshDepartmentsAsync 异常: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 额外补充：尝试删除临时文件（用于清理生成产物）
        /// </summary>
        /// <param name="path">要删除的临时文件路径</param>
        private void TryDeleteTempFile(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
                File.Delete(path);
                LogManager.Instance.LogInfo("已删除临时文件: " + path);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"删除临时文件失败: {path}, {ex.Message}");
            }
        }

        #endregion

        #region  第5阶段：剩余方法（属性显示、插入流程、DataGrid 支持、帮助方法等）

        /// <summary>
        /// 在 PropertiesDataGrid 中显示文件属性（按 Hash 优先，从 DB 回退到 FileId）
        /// </summary>
        private async Task DisplayFilePropertiesInDataGridAsync(FileStorage fileStorage)
        {
            try
            {
                // ---- 前置校验 ----
                if (fileStorage == null)
                {
                    LogManager.Instance.LogWarning("DisplayFilePropertiesInDataGridAsync: fileStorage 为空");
                    if (PropertiesDataGrid != null) PropertiesDataGrid.ItemsSource = null;
                    return;
                }

                LogManager.Instance.LogInfo($"在PropertiesDataGrid中显示文件 {fileStorage.DisplayName} 的属性 (FileId={fileStorage.Id}, FileHash={fileStorage.FileHash ?? "空"})");

                if (PropertiesDataGrid == null)
                {
                    LogManager.Instance.LogWarning("PropertiesDataGrid 控件为空");
                    return;
                }

                if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
                {
                    LogManager.Instance.LogWarning("数据库管理器不可用");
                    PropertiesDataGrid.ItemsSource = null;
                    return;
                }

                Dictionary<string, string> attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                // ---- 1. 优先按 FileHash 查询 ----
                if (!string.IsNullOrWhiteSpace(fileStorage.FileHash))
                {
                    try
                    {
                        // GetFileStorageWithAttributesByHashAsync 返回一个 Tuple<FileStorage, Dictionary<string, string>>，我们取其中的属性字典
                        var fileStorageWithAttributes = await _databaseManager.GetFileStorageWithAttributesByHashAsync(fileStorage.FileHash);
                        if (fileStorageWithAttributes.Item2 != null)
                            attributes = fileStorageWithAttributes.Item2;

                        LogManager.Instance.LogInfo(attributes.Count > 0
                            ? $"按 Hash={fileStorage.FileHash} 加载属性成功，条数={attributes.Count}"
                            : $"按 Hash={fileStorage.FileHash} 未获取到属性");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogWarning($"按 Hash 查询属性失败: {ex.Message}");
                    }
                }
                else
                {
                    LogManager.Instance.LogInfo("FileHash 为空，跳过基于 Hash 的查询");
                }

                // ---- 2. Hash 未命中或 FileHash 为空时，用 FileId 兜底 ----
                if (attributes.Count == 0 && fileStorage.Id > 0)
                {
                    try
                    {
                        var fallback = await _databaseManager.GetAttributesJsonByFileIdAsync(fileStorage.Id, fileStorage.FileAttributeId);
                        if (fallback != null && fallback.Count > 0)
                        {
                            attributes = fallback;
                            LogManager.Instance.LogInfo($"按 FileId={fileStorage.Id} 兜底加载属性成功，条数={attributes.Count}");
                        }
                        else
                        {
                            LogManager.Instance.LogInfo($"按 FileId={fileStorage.Id} 也未获取到属性");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogWarning($"按 FileId 查询属性失败: {ex.Message}");
                    }
                }

                // ---- 3. 最终兜底：如果 FileStorage 对象中已经携带了 AttributesJson 字典（例如从列表查询时已附带），直接使用 ----
                if (attributes.Count == 0 && fileStorage.AttributesJson is Dictionary<string, string> attachedDict && attachedDict.Count > 0)
                {
                    attributes = attachedDict;
                    LogManager.Instance.LogInfo($"使用 FileStorage 对象自带的属性字典，条数={attributes.Count}");
                }

                // ---- 4. 生成显示数据并绑定 ----
                var displayData = PrepareFileDisplayData(fileStorage, attributes);
                PropertiesDataGrid.ItemsSource = displayData;

                // 快照（如果需要的话）
                try { CapturePropertiesSnapshot(displayData); } catch { }

                if (displayData != null && displayData.Count > 0)
                    LogManager.Instance.LogInfo($"文件属性显示完成，共 {displayData.Count} 行");
                else
                    LogManager.Instance.LogInfo("文件属性显示完成，但数据为空");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError("在PropertiesDataGrid中显示文件属性时出错: " + ex.Message);
                if (PropertiesDataGrid != null) PropertiesDataGrid.ItemsSource = null;
            }
        }

        /// <summary>
        /// 根据 FileStorage 与属性字典构建 DataGrid 显示模型集合
        /// 返回 List<CategoryPropertyEditModel>
        /// </summary>
        //public List<CategoryPropertyEditModel> PrepareFileDisplayData(FileStorage fileStorage, Dictionary<string, string> attributes)
        //{
        //    var result = new List<CategoryPropertyEditModel>();

        //    // 辅助方法：安全获取字典值，若缺失则返回默认值
        //    string GetAttr(string key, string defaultValue = "")
        //    {
        //        if (attributes != null && attributes.TryGetValue(key, out var v) && !string.IsNullOrWhiteSpace(v))
        //            return v;
        //        return defaultValue;
        //    }

        //    // 标记哪些键已经被固定行占用，后续遍历时跳过它们
        //    var usedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        //    // 辅助方法：添加一行到结果列表，并标记使用的键
        //    void AddRow(string name1, string key1, string def1,
        //                string name2, string key2, string def2)
        //    {
        //        usedKeys.Add(key1);
        //        usedKeys.Add(key2);
        //        result.Add(new CategoryPropertyEditModel
        //        {
        //            PropertyName1 = name1,
        //            PropertyValue1 = GetAttr(key1, def1),
        //            PropertyName2 = name2,
        //            PropertyValue2 = GetAttr(key2, def2)
        //        });
        //    }

        //    try
        //    {
        //        // ---- 固定字段（从字典或直接赋默认值） ----
        //        string displayName = fileStorage?.DisplayName ?? string.Empty;
        //        AddRow("文件名", "FileName", displayName,
        //               "显示名称", "DisplayName", displayName);
        //        AddRow("元素块名", "BlockName", "",
        //               "图层名称", "LayerName", "");
        //        AddRow("颜色索引", "ColorIndex", "1",
        //               "比例", "Scale", "1");
        //        //AddRow("描述", "Description", "",
        //        //       "版本", "Version", "1");
        //        AddRow("创建者", "CreatedBy", Environment.UserName,
        //               "是否公开", "IsPublic", "是");
        //        AddRow("是否天正", "IsTianZheng", "否",
        //               "", "", "");
        //        AddRow("预览图名", "PreviewImageName", fileStorage.PreviewImageName,
        //            "预览图地址", "PreviewImageAddress", fileStorage.PreviewImagePath);
        //        // 几何信息
        //        AddRow("长度", "Length", "",
        //               "宽度", "Width", "");
        //        AddRow("高度", "Height", "",
        //               "角度", "Angle", "0");
        //        if (attributes != null && attributes.Count > 0)
        //        {
        //            //---- 处理剩余的非空属性，按两列显示 ----
        //            //var pairs = attributes
        //            //    .Where(kv => !string.IsNullOrWhiteSpace(kv.Value) && !usedKeys.Contains(kv.Key))
        //            //    .ToList();
        //            var pairs = attributes
        //                .Where(kv => !usedKeys.Contains(kv.Key))
        //                .ToList();
        //            for (int i = 0; i < pairs.Count; i += 2)
        //            {
        //                var p1 = pairs[i];
        //                var p2 = (i + 1 < pairs.Count) ? pairs[i + 1] : default;
        //                result.Add(new CategoryPropertyEditModel
        //                {
        //                    PropertyName1 = p1.Key,
        //                    PropertyValue1 = p1.Value,
        //                    PropertyName2 = p2.Key ?? string.Empty,
        //                    PropertyValue2 = p2.Value ?? string.Empty
        //                });
        //            }
        //        }
        //        return result;
        //    }
        //    catch (Exception ex)
        //    {
        //        LogManager.Instance.LogInfo("PrepareFileDisplayData 异常: " + ex.Message);
        //    }

        //    return result;
        //}


        public List<CategoryPropertyEditModel> PrepareFileDisplayData(FileStorage fileStorage, Dictionary<string, string> attributes)
        {
            var result = new List<CategoryPropertyEditModel>();

            string GetAttr(string key, string defaultValue = "")
            {
                if (attributes != null && attributes.TryGetValue(key, out var v) && !string.IsNullOrWhiteSpace(v))
                    return v;
                return defaultValue;
            }

            var usedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AddRow(string name1, string key1, string def1, string name2, string key2, string def2)
            {
                usedKeys.Add(key1);
                usedKeys.Add(key2);
                result.Add(new CategoryPropertyEditModel
                {
                    PropertyName1 = name1,
                    PropertyValue1 = GetAttr(key1, def1),
                    PropertyName2 = name2,
                    PropertyValue2 = GetAttr(key2, def2)
                });
            }

            try
            {
                // ---- 固定字段 ----
                string displayName = fileStorage?.DisplayName ?? string.Empty;
                AddRow("文件名", "FileName", displayName, "显示名称", "DisplayName", displayName);
                AddRow("元素块名", "BlockName", "", "图层名称", "LayerName", "");
                AddRow("颜色索引", "ColorIndex", "1", "比例", "Scale", "1");
                AddRow("创建者", "CreatedBy", Environment.UserName, "是否公开", "IsPublic", "是");
                AddRow("是否天正", "IsTianZheng", "否", "", "", "");
                AddRow("预览图名", "PreviewImageName", fileStorage.PreviewImageName, "预览图地址", "PreviewImageAddress", fileStorage.PreviewImagePath);
                AddRow("长度", "Length", "", "宽度", "Width", "");
                AddRow("高度", "Height", "", "角度", "Angle", "0");

                // ---- 动态属性（预设 + 剩余排序） ----
                if (attributes != null && attributes.Count > 0)
                {
                    // 1. 预设顺序键列表（不包含 REMARK）
                    var predefinedKeys = new List<string>
                    {
                        "PIPELINETITLE", "TAG_NO", "NAME", "MODEL",
                        "DRAWINGNO.STANDARDNO", "STRUCT_LEN_STD", "START_POINT", "END_POINT",
                        "FLG_STD", "TEST_STD", "DNCONN_TYPE", "QTY",
                        "WEIGHT", "SW_MODEL", "SYSTEM","REMARK"
                        // REMARK 单独处理，放在最后
                    };
                    var predefinedKeysSet = new HashSet<string>(predefinedKeys, StringComparer.OrdinalIgnoreCase);

                    // 2. 按顺序提取预设键（不含 REMARK），并记录已用的键
                    var predefinedPairs = new List<KeyValuePair<string, string>>();
                    foreach (var key in predefinedKeys)
                    {
                        if (!usedKeys.Contains(key) && attributes.TryGetValue(key, out var value))
                        {
                            predefinedPairs.Add(new KeyValuePair<string, string>(key, value));
                        }
                    }

                    // 3. 提取剩余键（排除 usedKeys、所有预设键及 REMARK），并按字母排序
                    var allPredefinedSet = new HashSet<string>(predefinedKeys, StringComparer.OrdinalIgnoreCase);
                    allPredefinedSet.Add("Remarks"); // 也排除 REMARK，因为要单独处理
                    var remainingPairs = attributes
                        .Where(kv => !usedKeys.Contains(kv.Key) && !allPredefinedSet.Contains(kv.Key))
                        .OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    // 4. 合并：预设 + 剩余排序
                    var allPairs = predefinedPairs.Concat(remainingPairs).ToList();

                    // 5. 单独处理 REMARK（如果存在且未被 usedKeys 占用）
                    if (!usedKeys.Contains("Remarks") && attributes.TryGetValue("Remarks", out var remarkValue))
                    {
                        allPairs.Add(new KeyValuePair<string, string>("Remarks", remarkValue ?? string.Empty));
                    }

                    // 6. 两两配对添加到结果
                    for (int i = 0; i < allPairs.Count; i += 2)
                    {
                        var p1 = allPairs[i];
                        var p2 = (i + 1 < allPairs.Count) ? allPairs[i + 1] : default;
                        result.Add(new CategoryPropertyEditModel
                        {
                            PropertyName1 = p1.Key,
                            PropertyValue1 = p1.Value ?? string.Empty,
                            PropertyName2 = p2.Key ?? string.Empty,
                            PropertyValue2 = p2.Value ?? string.Empty
                        });
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("PrepareFileDisplayData 异常: " + ex.Message);
            }
            return result;
        }


        /// <summary>
        /// 抓取当前 DataGrid 的键/值快照到 _propertiesSnapshotForInsert（用于“还原初始值”功能）
        /// </summary>
        private void CapturePropertiesSnapshot(List<CategoryPropertyEditModel> displayData)
        {
            try
            {
                _propertiesSnapshotForInsert = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (displayData == null) return;
                foreach (var row in displayData)
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(row.PropertyName1))
                            _propertiesSnapshotForInsert[NormalizePropertyDisplayName(row.PropertyName1)] = row.PropertyValue1 ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(row.PropertyName2))
                            _propertiesSnapshotForInsert[NormalizePropertyDisplayName(row.PropertyName2)] = row.PropertyValue2 ?? string.Empty;
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("CapturePropertiesSnapshot 异常: " + ex.Message);
            }
        }

        /// <summary>
        /// 规范化属性展示名（去掉空格并小写）
        /// </summary>
        /// <param name="name"> 属性展示名 </param>
        /// <returns> 规范化后的属性展示名 </returns>
        private string NormalizePropertyDisplayName(string name)
        {
            return (name ?? string.Empty).Trim().ToLowerInvariant();
        }

        /// <summary>
        /// 执行 DWG 插入并等待结果的异步包装（返回 (成功, 错误信息)）
        /// 实现说明：尝试使用已有的辅助方法插入（InsertGraphicHelper），尽量采用同步到 CAD 的方式并且避免阻塞 UI 线程
        /// </summary>
        private Task<(bool, string)> ExecuteInsertAndWaitResultAsync(string localDwgPath)
        {
            return Task.Run(() =>
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(localDwgPath) || !File.Exists(localDwgPath))
                        return (false, "本地 DWG 文件不存在");

                    // 优先使用 InsertGraphicHelper 的快速插入，如果存在该工具
                    try
                    {
                        InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(localDwgPath);
                        return (true, string.Empty);
                    }
                    catch (Exception ex)
                    {
                        // 若 InsertGraphicHelper 不可用，则尝试发送字符串到 CAD（兼容性兜底）
                        try
                        {
                            var doc = Application.DocumentManager.MdiActiveDocument;
                            if (doc != null)
                            {
                                doc.SendStringToExecute($"_.-INSERT \"{localDwgPath}\" 0,0 1 1 0 ", true, false, false);
                                return (true, string.Empty);
                            }
                            return (false, "未找到活动文档执行插入");
                        }
                        catch (Exception ex2)
                        {
                            return (false, $"插入异常: {ex2.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    return (false, ex.Message);
                }
            });
        }

        /// <summary>
        /// DataGrid 加载行时事件（确保行右键菜单）
        /// </summary>
        private void DataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            try
            {
                EnsureRowContextMenu(e.Row);
            }
            catch { }
        }

        /// <summary>
        /// StorageFile DataGrid 的 Loaded 事件：为已存在行添加上下文菜单并订阅 LoadingRow
        /// </summary>
        private void StroageFileDataGrid_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!(sender is DataGrid dg)) return;
                dg.LoadingRow -= DataGrid_LoadingRow;
                dg.LoadingRow += DataGrid_LoadingRow;

                foreach (var item in dg.Items)
                {
                    if (dg.ItemContainerGenerator.ContainerFromItem(item) is DataGridRow row)
                        EnsureRowContextMenu(row);
                }
            }
            catch { }
        }

        /// <summary>
        /// 为单行添加右键上下文菜单（更新图元、更新预览图等）
        /// </summary>
        private void EnsureRowContextMenu(DataGridRow row)
        {
            if (row == null || row.ContextMenu != null) return;

            var cm = new System.Windows.Controls.ContextMenu();

            var miUpdate = new System.Windows.Controls.MenuItem { Header = "更新图元" };
            miUpdate.CommandParameter = row.Item;
            miUpdate.Click += ReplaceFileMenuItem_Click;// 更新图元按钮事件
            cm.Items.Add(miUpdate);// 更新图元

            var miUpdatePreview = new System.Windows.Controls.MenuItem { Header = "更新预览图" };
            miUpdatePreview.CommandParameter = row.Item;
            miUpdatePreview.Click += ReplacePreviewMenuItem_Click;
            cm.Items.Add(miUpdatePreview);

            var miDelete = new System.Windows.Controls.MenuItem { Header = "删除图元" };
            miDelete.CommandParameter = row.Item;
            miDelete.Click += DeleteGraphic_Btn_Click;
            //miDelete.Click += DeleteRowGraphicMenuItem_Click;
            cm.Items.Add(miDelete);
            row.ContextMenu = cm;
        }

        /// <summary>
        /// 在视觉树中查找命名的子元素（泛型实现）
        /// </summary>
        private T FindVisualChild<T>(DependencyObject parent, string childName) where T : DependencyObject
        {
            if (parent == null) return default(T);
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typed && child is FrameworkElement fe && fe.Name == childName)
                    return typed;
                var found = FindVisualChild<T>(child, childName);
                if (found != null) return found;
            }
            return default(T);
        }

        /// <summary>
        /// 从字符串中提取连续的中文字符（用于资源文件名解析）
        /// </summary>
        private string ExtractChineseCharacters(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            var mc = Regex.Matches(input, "[\\u4e00-\\u9fff]+");
            if (mc.Count == 0) return string.Empty;
            return string.Concat(mc.Cast<Match>().Select(m => m.Value)).Trim();
        }

        /// <summary>
        /// 窗口关闭时进行必要的清理（释放缓存、删除临时文件）
        /// </summary>
        protected void OnClosing(CancelEventArgs e)
        {
            try
            {
                CleanupInvalidImageCache();
                // 若基类有 OnClosing，请在子类中调用（这里使用原方法名以兼容旧代码）
                // base.OnClosing(e); // UserControl 无此方法，按需在宿主窗口调用
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("关闭窗口时清理缓存失败: " + ex.Message);
            }
        }

        #endregion

        #region 对图元右键的点击事件
        /// <summary>
        /// 更新图元（替换文件）的右键菜单项点击事件处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void ReplaceFileMenuItem_Click(object? sender, RoutedEventArgs e)
        {
            try
            {
                // 获取被点击菜单项的数据上下文（要替换的文件记录对象）
                if (!(sender is MenuItem menuItem))
                    return;
                // 获取菜单项的数据上下文
                object? target = menuItem.CommandParameter ?? menuItem.DataContext;
                if (target == null)
                {
                    // 尝试通过父级 ContextMenu 和 PlacementTarget（如 DataGridRow）获取
                    var parent = VisualTreeHelper.GetParent(menuItem);// 获取右键选项的父级
                    while (!(parent is ContextMenu) && parent != null)
                        parent = VisualTreeHelper.GetParent(parent);//
                    //如果父级是 ContextMenu，则尝试获取 PlacementTarget
                    if (parent is ContextMenu cm && cm.PlacementTarget is DataGridRow row)
                        target = row.DataContext;// 获取 DataGridRow 的数据上下文
                }
                // 如果DataGridRow 的数据上下文为空，则无法获取要替换的文件记录
                if (target == null)
                {
                    MessageBox.Show("未能识别要替换的文件记录。", "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
                    return;
                }

                // 权限检查：仅管理员可操作
                string userName = VariableDictionary._userName;
                if (!WpfMainWindow.IsAdminUser(userName))
                {
                    MessageBox.Show("仅管理员用户可以执行替换操作。", "权限不足", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }

                // 选择本地文件
                using var ofd = new OpenFileDialog();
                ofd.Filter = "DWG 文件 (*.dwg)|*.dwg|所有文件 (*.*)|*.*";
                ofd.Title = "选择要上传并替换的文件";
                if (ofd.ShowDialog() != DialogResult.OK)
                    return;

                string localPath = ofd.FileName;// 获取选择的文件路径
                if (!File.Exists(localPath))
                {
                    MessageBox.Show("所选文件不存在。", "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
                    return;
                }

                // 二次确认
                if (MessageBox.Show($"确认将本地文件\n{Path.GetFileName(localPath)}\n覆盖服务器上此条记录对应的文件（保留原始文件名/位置）？",
                        "确认替换", MessageBoxButton.OKCancel, MessageBoxImage.Question) != MessageBoxResult.OK)
                    return;

                // 调用 API 执行替换
                (bool success, string error) = await TryInvokeReplaceApisAsync(target, localPath);
                if (!success)
                {
                    MessageBox.Show("替换失败: " + error, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
                    return;
                }

                // 清理缓存（预览图、本地CAD缓存）
                try
                {
                    string previewKey = GetStoragePreviewKey(target);
                    if (!string.IsNullOrWhiteSpace(previewKey))
                    {
                        string cachedPath = Path.Combine(GetPath.PreviewCachePath ?? string.Empty, previewKey + ".png");
                        if (File.Exists(cachedPath))
                            File.Delete(cachedPath);
                    }

                    await InvalidateLocalCadCacheAfterReplaceAsync(target, localPath);
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogWarning("替换后清理缓存失败: " + ex.Message);
                }

                // 刷新界面显示
                try
                {
                    await RefreshCurrentCategoryDisplayAsync(_selectedCategoryNode);
                    await ReloadButtonsDataSourceAfterReplaceAsync();
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogWarning("替换后刷新数据源失败: " + ex.Message);
                }

                MessageBox.Show("替换成功。", "完成", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            catch (Exception ex)
            {
                MessageBox.Show("替换过程中发生异常: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }
        
        /// <summary>
        /// 替换选中图元的预览图（右键菜单触发）
        /// 流程：获取记录 → 选图 → 上传到服务器现有 API → 清理缓存 → 刷新UI
        /// </summary>
        private async void ReplacePreviewMenuItem_Click(object? sender, RoutedEventArgs e)
        {
            // 用于日志追踪本次操作
            string traceId = Guid.NewGuid().ToString("N").Substring(0, 8);

            try
            {
                // ==================== 步骤1：获取 FileStorage ====================
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤1: 开始获取选中的图元记录");
                FileStorage? storage = GetFileStorageFromMenuItem(sender as MenuItem);
                if (storage == null)
                {
                    LogManager.Instance.LogWarning($"[ReplacePreview|{traceId}] 步骤1 失败: 未获取到图元对象");
                    MessageBox.Show("未能识别要替换预览图的记录。", "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
                    return;
                }
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤1 完成: 获取到图元 ID={storage.Id}, DisplayName={storage.DisplayName}");

                // ==================== 步骤2：管理员权限检查 ====================
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤2: 检查管理员权限");
                string userName = VariableDictionary._userName;
                if (!IsAdminUser(userName))
                {
                    LogManager.Instance.LogWarning($"[ReplacePreview|{traceId}] 步骤2 失败: 用户 '{userName}' 无管理员权限");
                    MessageBox.Show("仅管理员用户可以执行替换预览图操作。", "权限不足", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤2 通过: 用户 '{userName}' 具有管理员权限");

                // ==================== 步骤3：数据库可用性检查 ====================
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤3: 检查数据库可用性");
                if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
                {
                    LogManager.Instance.LogWarning($"[ReplacePreview|{traceId}] 步骤3 失败: 数据库不可用");
                    MessageBox.Show("数据库不可用，无法替换预览图。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤3 通过: 数据库连接可用");

                // ==================== 步骤4：从数据库获取最新记录 ====================
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤4: 从数据库获取最新记录 (ID={storage.Id})");
                FileStorage? latest = await _databaseManager.GetFileByIdAsync(storage.Id);
                if (latest == null)
                {
                    LogManager.Instance.LogWarning($"[ReplacePreview|{traceId}] 步骤4 失败: 数据库中未找到 ID={storage.Id} 的记录");
                    MessageBox.Show("数据库中未找到该图元记录，可能已被删除。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤4 完成: 获取到最新记录, 原预览图路径={latest.PreviewImagePath ?? "无"}");

                // ==================== 步骤5：选择新预览图 ====================
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤5: 打开文件选择对话框");
                string? newPreviewLocalPath;
                using (var ofd = new OpenFileDialog())
                {
                    ofd.Filter = "图片文件 (*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff)|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|所有文件 (*.*)|*.*";
                    ofd.Title = "选择要替换的预览图";
                    if (ofd.ShowDialog() != DialogResult.OK)
                    {
                        LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤5 取消: 用户关闭了文件选择对话框");
                        return;
                    }

                    newPreviewLocalPath = ofd.FileName;
                    if (!File.Exists(newPreviewLocalPath))
                    {
                        LogManager.Instance.LogWarning($"[ReplacePreview|{traceId}] 步骤5 失败: 所选文件不存在 -> {newPreviewLocalPath}");
                        MessageBox.Show("所选图片不存在。", "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
                        return;
                    }
                }
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤5 完成: 用户选择了 -> {newPreviewLocalPath} (大小={new FileInfo(newPreviewLocalPath).Length} 字节)");

                // ==================== 步骤6：确认操作 ====================
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤6: 弹出确认对话框");
                MessageBoxResult confirm = MessageBox.Show(
                    $"确认将图片\n{Path.GetFileName(newPreviewLocalPath)}\n替换为该图元的预览图吗？",
                    "确认替换预览图",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Question);
                if (confirm != MessageBoxResult.OK)
                {
                    LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤6 取消: 用户取消了替换操作");
                    return;
                }
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤6 通过: 用户确认替换");

                // ==================== 步骤7：上传到服务器 ====================
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤7: 开始上传预览图到服务器");
                bool uploadSuccess = await ReplacePreviewViaUploadApiAsync(traceId, latest, newPreviewLocalPath); //添加当前图形入库的上传接口，使用 replacePreviewId 字段标识替换预览图
                if (!uploadSuccess)
                {
                    LogManager.Instance.LogError($"[ReplacePreview|{traceId}] 步骤7 失败: 服务器上传失败");
                    MessageBox.Show("替换预览图失败，请检查服务器连接或查看日志。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤7 完成: 服务器上传成功");

                // ==================== 步骤8：清理客户端缓存 ====================
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤8: 清理客户端缓存");
                ClearPreviewCache(latest);
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤8 完成: 缓存已清理");

                // ==================== 步骤9：从数据库重新获取最新记录（服务器已更新） ====================
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤9: 从数据库重新获取最新记录（验证更新结果）");
                FileStorage? refreshed = await _databaseManager.GetFileByIdAsync(storage.Id);
                if (refreshed != null)
                {
                    LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤9 完成: 新预览图路径={refreshed.PreviewImagePath}, 新预览图名={refreshed.PreviewImageName}");
                }
                else
                {
                    LogManager.Instance.LogWarning($"[ReplacePreview|{traceId}] 步骤9: 无法获取更新后的记录");
                    refreshed = latest; // 使用旧记录作为回退
                }

                // ==================== 步骤10：刷新 UI ====================
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤10: 刷新 UI 列表");
                await RefreshFilesForCurrentCategoryAsync();
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤10-1: 刷新按钮面板");
                await ReloadButtonsDataSourceAfterReplaceAsync();
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 步骤10-2: 刷新预览显示");
                RefreshPreviewDisplay(refreshed);

                // ==================== 完成 ====================
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] ===== 全部步骤完成，替换预览图成功 =====");
                MessageBox.Show("替换预览图成功。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"[ReplacePreview|{traceId}] 未捕获异常: {ex.GetType().Name} - {ex.Message}\n堆栈: {ex.StackTrace}");
                MessageBox.Show($"替换预览图过程中发生异常: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        /// <summary>
        /// 通过服务器的上传接口替换预览图。
        /// 参考"添加当前图形入库"的实现，使用 POST /api/graphics/upload，
        /// 并通过 replacePreviewId 字段标识这是替换预览图操作。
        /// </summary>
        /// <param name="traceId">日志追踪ID</param>
        /// <param name="storage">要替换的图元记录</param>
        /// <param name="localPreviewPath">本地新预览图路径</param>
        /// <returns>上传是否成功</returns>
        private async Task<bool> ReplacePreviewViaUploadApiAsync(string traceId, FileStorage storage, string localPreviewPath)
        {
            try
            {
                // 构建服务器 URL
                string serverIp = VariableDictionary._serverIP?.Trim() ?? "127.0.0.1";
                int serverPort = VariableDictionary._apiPort > 0 ? VariableDictionary._apiPort : 10010;
                string apiUrl = $"http://{serverIp}:{serverPort}/api/graphics/upload";

                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 上传目标 URL: {apiUrl}");
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 上传参数: storageId={storage.Id}, 文件名={Path.GetFileName(localPreviewPath)}");

                // 检查文件是否存在且可读
                if (!File.Exists(localPreviewPath))
                {
                    LogManager.Instance.LogError($"[ReplacePreview|{traceId}] 上传前检查失败: 文件不存在 -> {localPreviewPath}");
                    return false;
                }
                long fileSize = new FileInfo(localPreviewPath).Length;
                if (fileSize == 0)
                {
                    LogManager.Instance.LogError($"[ReplacePreview|{traceId}] 上传前检查失败: 文件大小为 0 -> {localPreviewPath}");
                    return false;
                }
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 上传文件大小: {fileSize} 字节 ({fileSize / 1024.0:F2} KB)");

                // 构建 multipart/form-data 内容（与"添加当前图形入库"保持一致）
                using var content = new MultipartFormDataContent();

                // 1. 添加预览图文件
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 准备读取文件流...");
                var fileStream = new FileStream(localPreviewPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                var streamContent = new StreamContent(fileStream);

                // 根据扩展名设置 MIME 类型
                string ext = Path.GetExtension(localPreviewPath).ToLowerInvariant();
                string mimeType = ext switch
                {
                    ".png" => "image/png",
                    ".jpg" or ".jpeg" => "image/jpeg",
                    ".gif" => "image/gif",
                    ".bmp" => "image/bmp",
                    ".tif" or ".tiff" => "image/tiff",
                    _ => "application/octet-stream"
                };
                streamContent.Headers.ContentType = new MediaTypeHeaderValue(mimeType);
                content.Add(streamContent, "previewFile", Path.GetFileName(localPreviewPath));
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 已添加预览图文件 (MIME: {mimeType})");

                // 2. 关键字段：告诉服务器这是替换预览图操作，只传预览图，不传 DWG
                content.Add(new StringContent(storage.Id.ToString()), "replacePreviewId");
                content.Add(new StringContent("true"), "replacePreviewOnly");
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 已添加 replacePreviewId={storage.Id}, replacePreviewOnly=true");

                // 3. 可选：传递原始信息便于服务器校验
                content.Add(new StringContent(storage.DisplayName ?? storage.FileName ?? ""), "displayName");
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 已添加 displayName={storage.DisplayName}");

                // 4. 发送 POST 请求
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 开始发送 HTTP POST 请求...");
                var startTime = DateTime.UtcNow;

                using var response = await _uploadHttpClient.PostAsync(apiUrl, content);

                var elapsed = (DateTime.UtcNow - startTime).TotalSeconds;
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 收到服务器响应: StatusCode={response.StatusCode} (耗时 {elapsed:F2}s)");

                // 5. 读取响应体
                string responseBody = await response.Content.ReadAsStringAsync();
                LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 服务器响应内容: {responseBody}");

                if (response.IsSuccessStatusCode)
                {
                    LogManager.Instance.LogInfo($"[ReplacePreview|{traceId}] 上传成功 (HTTP {(int)response.StatusCode})");
                    return true;
                }
                else
                {
                    LogManager.Instance.LogError($"[ReplacePreview|{traceId}] 上传失败 (HTTP {(int)response.StatusCode}): {responseBody}");

                    // 根据状态码给出更具体的错误信息
                    string errorDetail = response.StatusCode switch
                    {
                        System.Net.HttpStatusCode.MethodNotAllowed => "服务器不允许此请求方法，请检查 API 路由是否正确",
                        System.Net.HttpStatusCode.NotFound => "服务器未找到上传接口，请检查服务器是否运行",
                        System.Net.HttpStatusCode.RequestEntityTooLarge => "文件过大，超出服务器限制",
                        System.Net.HttpStatusCode.Unauthorized => "未授权，请检查登录状态",
                        System.Net.HttpStatusCode.InternalServerError => "服务器内部错误",
                        _ => $"未知错误 ({(int)response.StatusCode})"
                    };
                    LogManager.Instance.LogError($"[ReplacePreview|{traceId}] 错误详情: {errorDetail}");
                    return false;
                }
            }
            catch (HttpRequestException ex)
            {
                LogManager.Instance.LogError($"[ReplacePreview|{traceId}] HTTP 请求异常: {ex.GetType().Name} - {ex.Message}");
                LogManager.Instance.LogError($"[ReplacePreview|{traceId}] 异常堆栈: {ex.StackTrace}");
                return false;
            }
            catch (TaskCanceledException ex)
            {
                LogManager.Instance.LogError($"[ReplacePreview|{traceId}] 请求超时: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"[ReplacePreview|{traceId}] 上传过程未预期异常: {ex.GetType().Name} - {ex.Message}");
                LogManager.Instance.LogError($"[ReplacePreview|{traceId}] 异常堆栈: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// 从右键菜单的 MenuItem 中安全提取 FileStorage 对象
        /// </summary>
        private FileStorage? GetFileStorageFromMenuItem(MenuItem? menuItem)
        {
            if (menuItem == null) return null;

            // 优先级1: CommandParameter
            if (menuItem.CommandParameter is FileStorage fs)
                return fs;

            // 优先级2: DataContext
            if (menuItem.DataContext is FileStorage dcFs)
                return dcFs;

            // 优先级3: 从父级 ContextMenu 的 PlacementTarget 获取
            var cm = menuItem.Parent as ContextMenu;
            if (cm == null)
            {
                DependencyObject parent = VisualTreeHelper.GetParent(menuItem);
                while (parent != null && !(parent is ContextMenu))
                    parent = VisualTreeHelper.GetParent(parent);
                cm = parent as ContextMenu;
            }
            if (cm?.PlacementTarget is DataGridRow row && row.DataContext is FileStorage rowFs)
                return rowFs;

            // 优先级4: StroageFileDataGrid 当前选中项（兜底）
            if (StroageFileDataGrid?.SelectedItem is FileStorage selected)
                return selected;

            return null;
        }

        /// <summary>
        /// 清理指定图元的所有客户端缓存
        /// </summary>
        private void ClearPreviewCache(FileStorage storage)
        {
            try
            {
                int removedCount = 0;

                // 清理内存缓存
                var keysToRemove = new List<string>();
                foreach (var kv in _imageCache)
                {
                    if (kv.Key.Contains(storage.Id.ToString()) ||
                        (!string.IsNullOrWhiteSpace(storage.PreviewImagePath) && kv.Key.Contains(storage.PreviewImagePath)) ||
                        (!string.IsNullOrWhiteSpace(storage.FilePath) && kv.Key.Contains(storage.FilePath)))
                    {
                        keysToRemove.Add(kv.Key);
                    }
                }
                foreach (var key in keysToRemove)
                {
                    _imageCache.Remove(key);
                    removedCount++;
                }

                // 清理本地缓存文件
                if (!string.IsNullOrWhiteSpace(GetPath.PreviewCachePath) && Directory.Exists(GetPath.PreviewCachePath))
                {
                    var pattern = $"{storage.Id}_*.*";
                    var files = Directory.GetFiles(GetPath.PreviewCachePath, pattern);
                    foreach (var file in files)
                    {
                        try { File.Delete(file); removedCount++; }
                        catch { /* 忽略 */ }
                    }
                }

                LogManager.Instance.LogInfo($"ClearPreviewCache: 已清理图元 {storage.Id} 的缓存 (移除 {removedCount} 项)");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"ClearPreviewCache 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 刷新预览控件显示
        /// </summary>
        private async void RefreshPreviewDisplay(FileStorage storage)
        {
            try
            {
                var bmp = await GetPreviewImageAsync(storage);
                if (预览 != null)
                    预览.Source = bmp;
                if (ViewImage != null)
                    ViewImage.Source = bmp;
                LogManager.Instance.LogInfo($"RefreshPreviewDisplay: 已刷新图元 {storage.Id} 的预览显示");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"RefreshPreviewDisplay 失败: {ex.Message}");
            }
        }

        #endregion
     
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
        /// 添加文件时初始化属性编辑网格，包含管道相关的标准属性
        /// </summary>
        private async void AddFileInitializeFilePropertiesGrid()
        {
            try
            {
                // 准备：canonical 列表（应与 AttributeKeyMapper 中 canonical 对齐）
                var canonicalFields = new[]
                {
            "PIPELINETITLE","TAG_NO","NAME","MODEL","DRAWINGNO_STANDARDNO","STRUCT_LEN_STD",
            "START_POINT","END_POINT","FLG_STD","TEST_STD","DN","PIPE_OD","PIPE_ID","PIPE_THK",
            "SCHEDULE","PN","CLASS","WORK_PRESSURE","DESIGN_PRESSURE","DESIGN_TEMP","WORK_TEMP",
            "TEMP_RANGE","MEDIUM","PIPE_TYPE","PIPE_CLASS","CONN_TYPE","PIPE_MATL","LINING_MATL",
            "LINING_THK","LINING_PROC","PIPE_LENGTH","FLOW_VEL","FLOW_RATE","PIPE_WEIGHT","PIPE_COATING",
            "COATING_THK","INSUL_MATL","INSUL_THK","INSUL_TYPE","HEAT_TRACE","HEAT_TRACE_POWER",
            "NDT_RATIO","NDT_TYPE","PRESSURE_LOSS","ROUGHNESS","MAX_BENDING","EXPANSION","CLEANING_REQ",
            "QTY","WEIGHT","SW_MODEL","SYSTEM","REMARK"
        };

                // 显示名映射（中文），建议把这张表放到 AttributeKeyMapper/GetDisplayName 中统一维护
                var displayMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            {"PIPELINETITLE","管道提示标题"},
            {"TAG_NO","管段号"},
            {"NAME","名称"},
            {"MODEL","规格型号"},
            {"DRAWINGNO_STANDARDNO","图号/标准号"},
            {"STRUCT_LEN_STD","结构长度标准"},
            {"START_POINT","起点"},
            {"END_POINT","终点"},
            {"FLG_STD","法兰标准"},
            {"TEST_STD","试验标准"},
            {"DN","公称通径"},
            {"PIPE_OD","外径"},
            {"PIPE_ID","内径"},
            {"PIPE_THK","壁厚"},
            {"SCHEDULE","壁厚等级"},
            {"PN","公称压力"},
            {"CLASS","压力等级"},
            {"WORK_PRESSURE","工作压力"},
            {"DESIGN_PRESSURE","设计压力"},
            {"DESIGN_TEMP","设计温度"},
            {"WORK_TEMP","工作温度"},
            {"TEMP_RANGE","适用温度范围"},
            {"MEDIUM","介质"},
            {"PIPE_TYPE","管道类型"},
            {"PIPE_CLASS","管道等级"},
            {"CONN_TYPE","连接方式"},
            {"PIPE_MATL","管道材质"},
            {"LINING_MATL","衬里材质"},
            {"LINING_THK","衬里厚度"},
            {"LINING_PROC","衬里工艺"},
            {"PIPE_LENGTH","管道长度"},
            {"FLOW_VEL","设计流速"},
            {"FLOW_RATE","介质流量"},
            {"PIPE_WEIGHT","管道计算重量"},
            {"PIPE_COATING","防腐涂层"},
            {"COATING_THK","涂层厚度"},
            {"INSUL_MATL","保温材料"},
            {"INSUL_THK","保温厚度"},
            {"INSUL_TYPE","保温方式"},
            {"HEAT_TRACE","伴热类型"},
            {"HEAT_TRACE_POWER","伴热功率"},
            {"NDT_RATIO","无损检测比例"},
            {"NDT_TYPE","无损检测方法"},
            {"PRESSURE_LOSS","允许压损"},
            {"ROUGHNESS","内壁粗糙度"},
            {"MAX_BENDING","允许弯曲半径"},
            {"EXPANSION","热膨胀量"},
            {"CLEANING_REQ","清洁度要求"},
            {"QTY","数量"},
            {"WEIGHT","总重量"},
            {"SW_MODEL","3D模型文件名"},
            {"SYSTEM","所属系统"},
            {"REMARK","备注"}
        };

                // 构建 UI 用的行（两列形式：显示名 / 值；第二列临时存放 canonical 以便保存时读取）
                var properties = new List<CategoryPropertyEditModel>();

                foreach (var canonical in canonicalFields)
                {
                    string display = displayMap.ContainsKey(canonical) ? displayMap[canonical] : canonical;
                    // PropertyName1: 显示名（中文）， PropertyValue1: 初始值（空）
                    // PropertyName2: 我临时放 canonical key（便于后续保存）， PropertyValue2: 可留空或存默认
                    properties.Add(new CategoryPropertyEditModel
                    {
                        PropertyName1 = display,
                        PropertyValue1 = string.Empty,
                        PropertyName2 = canonical,      // 暂用第二列存 canonical key，或改 model 加字段 CanonicalKey 更好
                        PropertyValue2 = string.Empty
                    });
                }

                // 你原有的通用/文件元数据放在表头前面（可按需合并）
                // 可把显示名称、层名等插入到 properties 开头
                properties.Insert(0, new CategoryPropertyEditModel { PropertyName1 = "显示名称", PropertyValue1 = Path.GetFileNameWithoutExtension(_selectedFilePath), PropertyName2 = "元素块名", PropertyValue2 = "" });
                properties.Insert(1, new CategoryPropertyEditModel { PropertyName1 = "层名", PropertyValue1 = "TJ(  专业  )", PropertyName2 = "颜色索引", PropertyValue2 = "40" });
                properties.Insert(2, new CategoryPropertyEditModel { PropertyName1 = "描述", PropertyValue1 = "", PropertyName2 = "版本", PropertyValue2 = "1" });
                properties.Insert(3, new CategoryPropertyEditModel { PropertyName1 = "是否公开", PropertyValue1 = "是", PropertyName2 = "创建者", PropertyValue2 = Environment.UserName });

                CategoryPropertiesDataGrid.ItemsSource = properties;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"初始化属性编辑失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 从服务器获取预览图片并缓存
        /// </summary>
        private async Task<BitmapImage?> GetPreviewImageAsync(FileStorage fileStorage)
        {
            if (fileStorage == null) return null;// 参数检查传进来的文件存储对象是否为 null

            try
            {
                string? cachedPreview = await EnsureLocalPreviewCacheAsync(fileStorage);
                if (!string.IsNullOrWhiteSpace(cachedPreview))
                {
                    LogManager.Instance.LogInfo($"预览缓存路径: {cachedPreview}");
                    var bmp = LoadImageFromFile(cachedPreview);
                    if (bmp != null)
                    {
                        LogManager.Instance.LogInfo($"预览图加载成功 [路径与文件名={cachedPreview}]");
                        return bmp;
                    }
                    else
                    {
                        LogManager.Instance.LogWarning($"预览文件存在但加载失败: {cachedPreview}");
                    }
                }
                else
                {
                    LogManager.Instance.LogWarning($"未获取到预览缓存路径 [Hash={fileStorage.FilePath}], 文件名: {fileStorage.FileName}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"获取预览图异常: {ex.Message}");
            }
            return null;
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

        #endregion


        #region 图元tabItem 操作相关

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

        private async void 应用图元_Click(object sender, RoutedEventArgs e)
        {
            // 将“应用图元”按钮的行为统一为调用 CopyDwgAllFast，
            // 根据当前选中文件优先选择 _selectedFileStorage，然后是 _currentFileStorage，再是 _selectedFilePath
            try
            {
                // 新增：Drag期间禁止再次触发，避免命令重入导致崩溃
                if (GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.IsCopyDwgAllFastDragging ||
                    GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.IsCopyDwgAllFastBusy)
                {
                    MessageBox.Show("当前图元正在跟随插入，请先左键落点或按 Esc 结束当前插入。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                // 获取当前绘图比例（优先使用用户在 TextBox_绘图比例 中设置的值）
                //VariableDictionary.wpfTextBoxScale = AutoCadHelper.GetScale();
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
                var (ok, err) = await ExecuteInsertAndWaitResultAsync(tempPath);
                if (ok)
                {
                    await CleanupLocalCadCacheAfterInsertAsync(tempPath);
                }
                else
                {
                    MessageBox.Show($"插入失败，已保留缓存用于排查：{err}", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                if (string.IsNullOrEmpty(tempPath))
                {
                    MessageBox.Show("未找到可插入的 DWG 文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // 调用统一插入方法
                GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
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
        /// 统一写入文件信息面板（避免重复代码）
        /// </summary>
        private void UpdateFileInfoPanel(
            string? filePath,// 优先显示物理路径，如果需要也可以改为显示逻辑路径或其他标识
            string? displayName,// 显示名称优先，如果没有则显示文件名，是否格式化显示名称（例如去除前缀、下划线等）根据实际需求调整
            string? fileType,// 优先显示文件类型，如果没有则可以考虑从路径提取扩展名
            long? fileSizeBytes,// 显示文件大小（字节），如果没有则显示为 null
            string? previewPath,// 显示预览图片路径，如果没有则显示为 "无预览图片"
            bool formatDisplayName)
        {
            // 201229：新增对 null 值的处理，避免直接赋值给 Text 属性时抛异常
            this.filePath.Text = filePath ?? string.Empty;
            // 201229：新增对 displayName 的 null 或空白处理，避免直接赋值给 Text 属性时抛异常，同时根据 formatDisplayName 决定是否格式化显示名称
            var finalName = displayName ?? string.Empty;
            // 201229：新增对格式化显示名称的处理，避免直接赋值给 Text 属性时抛异常，同时根据 formatDisplayName 决定是否格式化显示名称
            if (formatDisplayName && !string.IsNullOrWhiteSpace(finalName))
            {
                finalName = FormatFileNameForDisplay(finalName);// 根据实际需求实现格式化逻辑，例如去除前缀、下划线等    
            }
            FileName.Text = finalName;// 201229：新增对 null 值的处理，避免直接赋值给 Text 属性时抛异常
            // 201229：新增对 fileType 的 null 或空白处理，避免直接赋值给 Text 属性时抛异常，同时如果 fileType 为空则尝试从 filePath 提取扩展名作为文件类型显示
            FileSize.Text = (fileSizeBytes.HasValue && fileSizeBytes.Value > 0)
                ? $"{fileSizeBytes.Value / 1024.0:F2} KB"
                : string.Empty;
            // 201229：新增对 fileType 的 null 或空白处理，避免直接赋值给 Text 属性时抛异常，同时如果 fileType 为空则尝试从 filePath 提取扩展名作为文件类型显示
            ClientVersion.Text = fileType ?? string.Empty;
            // 201229：新增对 ClientVersion 的显示逻辑，如果 fileType 为空则尝试从 filePath 提取扩展名作为文件类型显示，避免直接赋值给 Text 属性时抛异常
            if (string.IsNullOrWhiteSpace(ClientVersion.Text) && !string.IsNullOrWhiteSpace(filePath))
            {
                // 201229：新增对 ClientVersion 的显示逻辑，如果 fileType 为空则尝试从 filePath 提取扩展名作为文件类型显示，避免直接赋值给 Text 属性时抛异常
                ClientVersion.Text = Path.GetExtension(filePath).ToLowerInvariant();
            }
            // 201229：新增对 previewPath 的 null 或空白处理，避免直接赋值给 Text 属性时抛异常，同时如果 previewPath 为空则显示为 "无预览图片"
            viewFilePath.Text = string.IsNullOrWhiteSpace(previewPath) ? "无预览图片" : previewPath;
        }

        /// <summary>
        /// 显示文件信息（路径模式）
        /// </summary>
        private void DisplayFileInfo(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath))// 201229：新增对空路径的处理，避免 FileInfo 抛异常
                {
                    UpdateFileInfoPanel(string.Empty, string.Empty, string.Empty, null, string.Empty, false);// 清空面板显示
                    return;
                }
                // 201229：新增对文件不存在的处理，避免 FileInfo 抛异常
                var fileInfo = new FileInfo(filePath);
                UpdateFileInfoPanel(
                    filePath: filePath,// 优先显示物理路径，如果需要也可以改为显示逻辑路径或其他标识
                    displayName: fileInfo.Name,// 显示文件名（含扩展名），如果需要也可以改为显示更友好的名称
                    fileType: fileInfo.Extension.ToLowerInvariant(),// 优先显示文件类型（扩展名），如果没有则可以考虑显示为 "未知类型"
                    fileSizeBytes: fileInfo.Exists ? fileInfo.Length : null,// 显示文件大小（字节），如果文件不存在则显示为 null
                    previewPath: viewFilePath.Text,// 保持现有预览图片路径显示不变，如果需要也可以改为 "无预览图片"
                    formatDisplayName: false);// 是否对显示名称进行格式化（例如去除前缀、下划线等），根据实际需求调整
            }
            catch (Exception ex)
            {
                MessageBox.Show($"显示文件信息失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 显示文件信息（FileStorage模式）
        /// </summary>
        /// <summary>
        /// 显示文件存储信息到用户界面面板。
        /// </summary>
        /// <param name="fileStorage">包含文件详细信息的 FileStorage 对象。如果为 null，则清空面板显示。</param>
        public void DisplayFileStorageInfo(FileStorage fileStorage)
        {
            try
            {
                // 处理空对象情况，重置文件信息面板
                if (fileStorage == null)
                {
                    UpdateFileInfoPanel(string.Empty, string.Empty, string.Empty, null, string.Empty, false);
                    return;
                }

                // 更新面板以显示具体的文件详细信息
                UpdateFileInfoPanel(
                    filePath: fileStorage.FilePath,// 优先显示物理路径，如果需要也可以改为显示逻辑路径或其他标识
                    displayName: fileStorage.DisplayName ?? fileStorage.FileName,// 显示名称优先，如果没有则显示文件名
                    fileType: fileStorage.FileType,// 优先显示文件类型，如果没有则可以考虑从路径提取扩展名
                    fileSizeBytes: fileStorage.FileSize,// 显示文件大小（字节），如果没有则显示为 null
                    previewPath: fileStorage.PreviewImagePath,// 显示预览图片路径，如果没有则显示为 "无预览图片"
                    formatDisplayName: true);// 是否对显示名称进行格式化（例如去除前缀、下划线等），根据实际需求调整
            }
            catch (Exception ex)
            {
                // 记录显示文件信息过程中的异常
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
                viewFilePath.Text = imagePath;

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
            /// <summary>
            /// 名称
            /// </summary>
            public string? Name { get; set; }
            /// <summary>
            /// 显示名称（可以与文件名不同，用于界面显示）
            /// </summary>
            public string? DisplayName { get; set; }
            /// <summary>
            /// 文件路径（物理路径或逻辑路径，根据实际需求决定显示哪一个）
            /// </summary>
            public string? FilePath { get; set; }
            /// <summary>
            /// 专业
            /// </summary>
            public string? Majors { get; set; } // 专业类别
            /// <summary>
            /// 创建时间（可以根据实际需求选择显示创建时间、修改时间或其他时间属性）
            /// </summary>
            public DateTime CreatedTime { get; set; }
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
        
        #endregion
        
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
        
        #endregion

        #region 图元替换/删除等 事件处理方法

        /// <summary>
        /// 管理员模块中对分类下图元 DataGrid 行右键菜单：删除图元
        /// </summary>
        private async void DeleteRowGraphicMenuItem_Click(object? sender, RoutedEventArgs e)
        {
            try
            {
                // 获取菜单项对象
                var mi = sender as System.Windows.Controls.MenuItem;
                // 优先从 CommandParameter 获取当前行对象
                object? storageObj = mi?.CommandParameter;
                // 兜底回退到 DataContext / PlacementTarget
                if (storageObj == null && mi != null)
                {
                    // 第一层回退：MenuItem.DataContext
                    storageObj = mi.DataContext;
                    if (storageObj == null)// 第二层回退：ContextMenu.PlacementTarget.DataContext
                    {
                        var cm = mi.Parent as System.Windows.Controls.ContextMenu;// 先尝试直接父级
                        var row = cm?.PlacementTarget as System.Windows.Controls.DataGridRow;// 如果父级不是 ContextMenu，则向上查找
                        storageObj = row?.DataContext;// 最终回退到行的 DataContext
                    }
                }
                // 强类型转换为 FileStorage
                var selected = storageObj as FileStorage;
                await DeleteGraphicCoreAsync(selected, needAdminCheck: true, entryName: "右键菜单");// 执行删除核心逻辑（包含权限校验、用户确认、数据库删除、预览缓存清理、界面刷新等）
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"右键删除图元异常: {ex.Message}");
                MessageBox.Show($"删除图元时发生异常：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>
        /// 核心删除逻辑：两个入口（按钮/右键）统一调用
        /// </summary>
        /// <param name="selected">待删除的图元对象</param>
        /// <param name="needAdminCheck">是否需要管理员权限检查</param>
        /// <param name="entryName">调用入口名称（用于日志区分）</param>
        /// <returns>删除是否成功</returns>
        private async Task<bool> DeleteGraphicCoreAsync(FileStorage selected, bool needAdminCheck, string entryName)
        {
            // ============ 1. 空对象保护 ============
            if (selected == null)
            {
                MessageBox.Show("未选中要删除的图元。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            // ============ 2. 管理员权限检查（右键入口可选） ============
            if (needAdminCheck)
            {
                var userName = VariableDictionary._userName;
                if (!IsAdminUser(userName))
                {
                    MessageBox.Show("仅管理员用户可以执行删除图元操作。", "权限不足", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }

            // ============ 3. 数据库可用性检查 ============
            if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
            {
                MessageBox.Show("数据库不可用，无法删除图元。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            // ============ 4. ID 有效性检查 ============
            if (selected.Id <= 0)
            {
                LogManager.Instance.LogWarning($"[{entryName}] 图元 ID 无效: {selected.Id}");
                MessageBox.Show("图元记录 ID 无效，无法执行删除。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            // ============ 5. 获取服务器端最新记录（保证路径等数据正确） ============
            FileStorage fullRecord = selected;
            try
            {
                var latest = await _databaseManager.GetFileByIdAsync(selected.Id).ConfigureAwait(true);
                if (latest != null)
                {
                    fullRecord = latest;
                    LogManager.Instance.LogInfo($"[{entryName}] 已获取最新数据库记录: {fullRecord.DisplayName} (ID={fullRecord.Id})");
                }
                else
                {
                    // 数据库已查不到记录，但仍可能残留在本地，继续尝试清理
                    LogManager.Instance.LogWarning($"[{entryName}] 数据库中未找到该记录 (ID={selected.Id})，将仅清理客户端缓存");
                    fullRecord = selected; // 使用传入对象继续后续清理
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"[{entryName}] 获取最新记录失败，使用当前数据继续: {ex.Message}");
                fullRecord = selected;
            }

            // ============ 6. 确认提示 ============
            string selectedName = fullRecord.DisplayName ?? fullRecord.FileName ?? $"ID={fullRecord.Id}";
            var confirm = MessageBox.Show(
                $"确定删除图元：{selectedName} ?\n该操作将删除主记录、属性JSON、标签/日志及关联物理文件，且不可恢复。",
                "删除确认",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirm != MessageBoxResult.Yes)
                return false;

            LogManager.Instance.LogInfo($"[{entryName}] 用户确认删除: {selectedName} (ID={fullRecord.Id})");

            // ============ 7. 删除物理文件（容错，失败不中断流程） ============
            bool physicalDeleted = false;
            if (_fileManager != null)
            {
                try
                {
                    await _fileManager.DeletePhysicalFilesAsync(_databaseManager, fullRecord, deleteBackupFiles: true);
                    physicalDeleted = true;
                    LogManager.Instance.LogInfo($"[{entryName}] 物理文件删除成功: {fullRecord.FilePath}");
                }
                catch (Exception exDeleteFile)
                {
                    LogManager.Instance.LogWarning($"[{entryName}] 删除服务器物理文件失败（继续执行数据库删除）: {exDeleteFile.Message}");
                }
            }

            // ============ 8. 级联删除数据库记录 ============
            bool dbDeleted = false;
            try
            {
                // physicalDelete=false 时走软删除（is_active=0），为 true 则物理删除
                dbDeleted = await _databaseManager.DeleteCadGraphicCascadeAsync(fullRecord.Id, physicalDelete: true);
                if (!dbDeleted)
                {
                    LogManager.Instance.LogError($"[{entryName}] 数据库级联删除失败: {selectedName} (ID={fullRecord.Id})");
                }
            }
            catch (Exception exDb)
            {
                LogManager.Instance.LogError($"[{entryName}] 数据库删除异常: {exDb.Message}");
            }

            // ============ 9. 如果没有成功删除任何东西，报错 ============
            if (!dbDeleted && !physicalDeleted)
            {
                MessageBox.Show("删除失败：物理文件和数据库记录均无法删除，请查看日志。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            else if (!dbDeleted)
            {
                // 物理已删但数据库失败
                MessageBox.Show("物理文件已删除，但数据库记录删除失败，请查看日志并手动处理。", "部分成功", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            // ============ 10. 清理内存缓存 ============
            try
            {
                // 按路径键清理
                var keysToRemove = _imageCache
                    .Where(kv => kv.Key.Contains(fullRecord.Id.ToString()) ||
                                 (!string.IsNullOrWhiteSpace(fullRecord.FilePath) && kv.Key.Contains(fullRecord.FilePath)))
                    .Select(kv => kv.Key)
                    .ToList();

                foreach (var key in keysToRemove)
                    _imageCache.Remove(key);

                LogManager.Instance.LogInfo($"[{entryName}] 已清理 {keysToRemove.Count} 条图片缓存");
            }
            catch (Exception exCache)
            {
                LogManager.Instance.LogWarning($"[{entryName}] 清理预览缓存失败: {exCache.Message}");
            }

            // ============ 11. 清理选择状态 ============
            if ((_selectedFileStorage != null && _selectedFileStorage.Id == fullRecord.Id) ||
                (_currentFileStorage != null && _currentFileStorage.Id == fullRecord.Id))
            {
                _selectedFileStorage = null;
                _currentFileStorage = null;
                _selectedFilePath = null;
                _selectedPreviewImagePath = null;
            }

            // ============ 12. 清空详情区 ============
            CategoryPropertiesDataGrid.ItemsSource = null;
            PropertiesDataGrid.ItemsSource = null;
            if (预览 != null) 预览.Source = null;
            if (ViewImage != null) ViewImage.Source = null;
            filePath.Text = string.Empty;
            FileName.Text = string.Empty;
            FileSize.Text = string.Empty;
            ClientVersion.Text = string.Empty;
            viewFilePath.Text = string.Empty;

            // ============ 13. 刷新 UI 列表与按钮面板 ============
            try
            {
                await RefreshFilesForCurrentCategoryAsync();
                if (_useDatabaseMode && _databaseManager.IsDatabaseAvailable)
                {
                    await RefreshAllCategoryPanelsAsync();
                }
                StroageFileDataGrid.Items.Refresh();
            }
            catch (Exception exRefresh)
            {
                LogManager.Instance.LogWarning($"[{entryName}] 刷新界面失败: {exRefresh.Message}");
            }

            // ============ 14. 成功返回 ============
            LogManager.Instance.LogInfo($"[{entryName}] 删除完成: {selectedName} (ID={fullRecord.Id})");
            if (dbDeleted)
            {
                MessageBox.Show("删除成功。", "信息", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            return dbDeleted;
        }
       

        #endregion

        #region 清理本地 CAD 文件缓存的方法

        /// <summary>
        /// 仅清理“本次替换对应”的本地 CadFiles 缓存文件（按文件名精确定位）
        /// 不做全量清理，避免影响其它按钮预热缓存。
        /// </summary>
        private async Task InvalidateLocalCadCacheAfterReplaceAsync(object storageObj, string localPath)
        {
            await Task.Run(() =>
            {
                try
                {
                    // CadFiles 根目录（与 DynamicButton_Click 缓存规则保持一致）
                    string cacheRoot = string.IsNullOrWhiteSpace(_cadStoragePath)
                        ? System.IO.Path.Combine(AppPath, "CadFiles")
                        : _cadStoragePath;

                    if (!System.IO.Directory.Exists(cacheRoot))
                        return;

                    // 优先按 storage 计算缓存文件名（最精确）
                    string storageFileName = GetStorageStringProperty(storageObj, "FileName");
                    string storageFilePath = GetStorageStringProperty(storageObj, "FilePath");

                    string? cachePath = BuildExactCadCachePath(cacheRoot, storageFileName, storageFilePath, localPath);

                    if (string.IsNullOrWhiteSpace(cachePath))
                    {
                        LogManager.Instance.LogInfo("未能解析本次替换对应的本地缓存路径，跳过清理。");
                        return;
                    }

                    if (System.IO.File.Exists(cachePath))
                    {
                        System.IO.File.Delete(cachePath);
                        LogManager.Instance.LogInfo($"已精确删除本地缓存文件: {cachePath}");
                    }
                    else
                    {
                        LogManager.Instance.LogInfo($"本次替换对应缓存文件不存在，无需删除: {cachePath}");
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogWarning($"InvalidateLocalCadCacheAfterReplaceAsync 异常: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 按项目现有缓存命名规则构造“本次替换对应”的 CadFiles 缓存路径
        /// 规则：realName = storage.FileName（无扩展则补扩展），扩展优先 FilePath，再回退 localPath。
        /// </summary>
        private string? BuildExactCadCachePath(string cacheRoot, string storageFileName, string storageFilePath, string localPath)
        {
            try
            {
                // 优先从 storage.FilePath 取扩展名，其次用本地替换文件扩展名
                string ext = System.IO.Path.GetExtension(storageFilePath);
                if (string.IsNullOrWhiteSpace(ext))
                    ext = System.IO.Path.GetExtension(localPath);

                // 文件名优先 storage.FileName；若缺失则回退 localPath 文件名
                string realName = storageFileName;
                if (string.IsNullOrWhiteSpace(realName))
                    realName = System.IO.Path.GetFileNameWithoutExtension(localPath);

                if (string.IsNullOrWhiteSpace(realName))
                    return null;

                // 若 FileName 没有扩展名，按规则补上
                if (string.IsNullOrWhiteSpace(System.IO.Path.GetExtension(realName)) && !string.IsNullOrWhiteSpace(ext))
                {
                    realName = System.IO.Path.GetFileNameWithoutExtension(realName) + ext;
                }

                return System.IO.Path.Combine(cacheRoot, realName);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 替换成功后刷新按钮数据源，确保按钮 Tag 使用最新文件记录
        /// </summary>
        private async Task ReloadButtonsDataSourceAfterReplaceAsync()
        {
            try
            {
                // 刷新当前分类右侧文件列表
                await RefreshFilesForCurrentCategoryAsync();

                // 刷新九大主分类按钮面板（数据库模式）
                if (_useDatabaseMode && _databaseManager != null && _databaseManager.IsDatabaseAvailable)
                {
                    await RefreshAllCategoryPanelsAsync();
                }

                // 若当前正停在主分类页签，强制重载当前页签按钮
                var selectedTab = MainTabControl?.SelectedItem as TabItem;
                if (selectedTab != null)
                {
                    string header = (selectedTab.Header?.ToString() ?? string.Empty).Trim();
                    if (header == "工艺" || header == "建筑" || header == "结构" ||
                        header == "电气" || header == "给排水" || header == "暖通" ||
                        header == "自控" || header == "总图" || header == "公共图")
                    {
                        LoadButtonsForMainCategoryTab(selectedTab, header);
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"ReloadButtonsDataSourceAfterReplaceAsync 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 反射读取 storage 上的字符串属性
        /// </summary>
        private string GetStorageStringProperty(object storageObj, string propName)
        {
            if (storageObj == null || string.IsNullOrWhiteSpace(propName))
                return string.Empty;

            try
            {
                var t = storageObj.GetType();
                var p = t.GetProperty(propName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (p == null) return string.Empty;
                var v = p.GetValue(storageObj);
                return v?.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
        
        /// <summary>
        /// 仅在插入成功后清理本次 CadFiles 缓存
        /// </summary>
        private async Task CleanupLocalCadCacheAfterInsertAsync(string? insertedPath)
        {
            if (string.IsNullOrWhiteSpace(insertedPath))
                return;

            try
            {
                string cacheRoot = string.IsNullOrWhiteSpace(_cadStoragePath)
                    ? Path.Combine(AppPath, "CadFiles")
                    : _cadStoragePath;

                if (!Directory.Exists(cacheRoot))
                    return;

                string fullPath = Path.GetFullPath(insertedPath);
                string fullCache = Path.GetFullPath(cacheRoot).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;

                // 只删 CadFiles 目录内文件，避免误删服务器源文件
                if (!fullPath.StartsWith(fullCache, StringComparison.OrdinalIgnoreCase))
                    return;

                const int maxRetry = 8;
                for (int i = 0; i < maxRetry; i++)
                {
                    try
                    {
                        if (File.Exists(fullPath))
                        {
                            File.Delete(fullPath);
                            LogManager.Instance.LogInfo($"插入成功后已清理本地缓存: {fullPath}");
                        }
                        return;
                    }
                    catch (IOException)
                    {
                        await Task.Delay(200);
                    }
                    catch (UnauthorizedAccessException)
                    {
                        await Task.Delay(200);
                    }
                }

                LogManager.Instance.LogWarning($"插入后缓存清理失败（可能仍被占用）: {fullPath}");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"CleanupLocalCadCacheAfterInsertAsync 异常: {ex.Message}");
            }
        }
        #endregion

        /// <summary>
        /// 尝试异步调用替换API
        /// </summary>
        /// <param name="storageObj"></param>
        /// <param name="localPath"></param>
        /// <returns></returns>
        private async Task<(bool success, string error)> TryInvokeReplaceApisAsync(object storageObj, string localPath)
        {
            if (storageObj == null) return (false, "storageObj 为 null。");
            if (string.IsNullOrWhiteSpace(localPath) || !System.IO.File.Exists(localPath))
                return (false, "本地文件不存在或路径无效。");

            var storage = storageObj as FileStorage;
            if (storage == null)
                return (false, "无法识别文件记录类型（需要 FileStorage）。");

            if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
                return (false, "数据库管理器未初始化或数据库不可用。");

            if (_fileManager == null)
                return (false, "文件管理服务未初始化。");

            try
            {
                // 统一调用 FileManager 的服务器侧替换逻辑，避免回退到本地缓存路径
                await _fileManager.ReplaceGraphicFileAsync(_databaseManager, storage, localPath);

                // 替换成功后统一回写数据库记录
                bool dbOk = await _databaseManager.UpdateFileStorageAsync(storage);
                if (!dbOk)
                    return (false, "数据库更新失败（UpdateFileStorageAsync 返回 false）。");

                return (true, string.Empty);
            }
            catch (Exception ex)
            {
                return (false, $"替换过程中发生异常: {ex.Message}");
            }
        }

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

        #region 管理员按键处理方法...

        #region 树节点选中与右键操作

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
            /// <summary>
            /// 控制规范树节点是否展开。
            /// </summary>
            public bool IsExpanded { get; set; }
            //public string DisplayText => string.IsNullOrEmpty(DisplayName) ? Name : DisplayName;


            /// <summary>
            /// 构造函数
            /// </summary>
            /// <param name="id"> ID</param>
            /// <param name="name">名称</param>
            /// <param name="displayName">显示名称</param>
            /// <param name="level">层级</param>
            /// <param name="parentId">父Id</param>
            /// <param name="data">Object数据</param>
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
                //deleteItem.Click += 刷新文件列表按钮_Click;
                RefreshItem.Click += 刷新文件列表按钮_Click;
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
                var openFileDialog = new global::Microsoft.Win32.OpenFileDialog
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
            var selected = StroageFileDataGrid.SelectedItem as FileStorage;
            DeleteGraphicBtn.IsEnabled = false;
            try
            {
                await DeleteGraphicCoreAsync(selected, needAdminCheck: false, entryName: "按钮");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"按钮删除图元异常: {ex.Message}");
                MessageBox.Show($"删除过程中发生错误：{ex.Message}", "异常", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                DeleteGraphicBtn.IsEnabled = true;
            }
        }

        /// <summary>
        /// 还原初始值按钮点击事件
        /// </summary>
        private void ResetToInitial_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _currentOperation = ManagementOperationType.None;// 重置当前操作状态
                InitializeCategoryPropertyGrid();// 重置分类属性编辑界面
                InitializeFileUploadInterface();// 重置文件上传界面
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
                bool success = false;// 操作成功标志
                bool isFilePropertyUpdate = false;// 是否为文件属性更新

                switch (_currentOperation)// 根据当前操作类型(主分类、子分类)执行相应的操作
                {
                    case ManagementOperationType.AddCategory:
                        success = await _categoryManager.ApplyCategoryPropertiesAsync(CategoryPropertiesDataGrid);
                        break;
                    case ManagementOperationType.AddSubcategory:
                        success = await _categoryManager.ApplySubcategoryPropertiesAsync(CategoryPropertiesDataGrid);
                        break;
                    case ManagementOperationType.None:
                        // 中文注释：编辑模式下，若当前选中了图元文件，则优先执行图元属性更新
                        if (_selectedFileStorage != null && StroageFileDataGrid?.SelectedItem is FileStorage)
                        {
                            success = await ApplySelectedFilePropertiesAsync();
                            isFilePropertyUpdate = true;
                        }
                        // 中文注释：否则按原有逻辑更新分类/子分类属性
                        else if (_selectedCategoryNode != null)
                        {
                            success = await _categoryManager.UpdateCategoryPropertiesAsync(CategoryPropertiesDataGrid, _selectedCategoryNode);
                        }
                        else
                        {
                            MessageBox.Show("请先选择要操作的分类、子分类或图元", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                        break;
                    default:
                        MessageBox.Show("未知操作类型", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                }

                if (success)
                {
                    if (isFilePropertyUpdate)
                    {
                        // 中文注释：图元属性更新成功后，刷新当前分类下的图元列表，保证界面显示最新数据
                        await RefreshFilesForCurrentCategoryAsync();
                        MessageBox.Show("图元属性已更新到数据库，并已刷新当前分类", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        _currentOperation = ManagementOperationType.None;
                        InitializeCategoryPropertyGrid();
                        await _categoryManager.RefreshCategoryTreeAsync(_selectedCategoryNode, CategoryTreeView, _categoryTreeNodes, _databaseManager);
                        MessageBox.Show("操作成功，架构树已更新", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
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
        /// 上传当前选中的图形文件到服务器（无参便捷方法）
        /// 内部自动从 UI 控件和字段中收集上传所需的数据
        /// </summary>
        public async Task UploadFileAndSaveToDatabase()
        {
            // ===== 1. 前置校验 =====
            int categoryId = _selectedCategoryNode?.Id ?? 0;
            if (categoryId <= 0)
            {
                LogManager.Instance.LogWarning("UploadFileAndSaveToDatabase: 未选中有效分类，操作中止");
                MessageBox.Show("请先在左侧分类树中选择一个目标分类", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(_selectedFilePath) || !File.Exists(_selectedFilePath))
            {
                LogManager.Instance.LogWarning("UploadFileAndSaveToDatabase: 未选择有效的 DWG 文件");
                MessageBox.Show("请先选择一个 DWG 文件，例如通过“浏览”或“添加当前图形”。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // ===== 2. 从 UI 提取信息 =====
            string displayName = (FindName("TxtDisplayName") as TextBox)?.Text;
            if (string.IsNullOrWhiteSpace(displayName))
                displayName = Path.GetFileNameWithoutExtension(_selectedFilePath);

            string description = (FindName("TxtDescription") as TextBox)?.Text ?? string.Empty;
            string blockName = (FindName("txtBlockName") as TextBox)?.Text ?? string.Empty;
            string layerName = (FindName("txtLayerName") as TextBox)?.Text ?? string.Empty;

            int colorIndex = 1;
            double scale = 1.0;
            try
            {
                var txtColorIndex = FindName("TxtColorIndex") as TextBox;
                if (txtColorIndex != null && int.TryParse(txtColorIndex.Text, out int ci) && ci > 0)
                    colorIndex = ci;

                var txtScale = FindName("TxtScale") as TextBox;
                if (txtScale != null && double.TryParse(txtScale.Text, out double sc) && sc > 0)
                    scale = sc;
            }
            catch { }

            string createdBy = VariableDictionary._userName ?? "System";
            string? previewPath = _selectedPreviewImagePath;

            // ===== 3. ★ 关键修改：建立属性字典并填充所有提取到的字段 ★ =====
            var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 填充所有已提取的属性（键名需与 ImportConfirmWindow.UpdateDtoFromGrid 中的 NormalizeKey 映射一致）
            void AddIfNotEmpty(string key, string value)
            {
                if (!string.IsNullOrWhiteSpace(value))
                    attributes[key] = value;
            }

            AddIfNotEmpty("FileName", Path.GetFileName(_selectedFilePath));
            AddIfNotEmpty("DisplayName", displayName);
            AddIfNotEmpty("BlockName", blockName);
            AddIfNotEmpty("LayerName", layerName);
            AddIfNotEmpty("ColorIndex", colorIndex.ToString());
            AddIfNotEmpty("Scale", scale.ToString());
            AddIfNotEmpty("Description", description);
            AddIfNotEmpty("CreatedBy", createdBy);
            if (!string.IsNullOrWhiteSpace(previewPath))
                AddIfNotEmpty("PreviewImagePath", previewPath);

            // 如果有其他从 CAD 实体提取的属性，也在这里添加
            // 例如：从 _propertiesSnapshotForInsert 合并
            if (_propertiesSnapshotForInsert != null && _propertiesSnapshotForInsert.Count > 0)
            {
                foreach (var kvp in _propertiesSnapshotForInsert)
                {
                    if (!attributes.ContainsKey(kvp.Key))
                        attributes[kvp.Key] = kvp.Value;
                }
            }

            // ===== 4. 构建 DTO 并调用带参上传方法 =====
            var dto = new ImportEntityDto
            {
                FilePath = _selectedFilePath,
                PreviewImagePath = previewPath,
                CategoryId = categoryId,
                CategoryType = "sub",
                BlockName = blockName,
                LayerName = layerName,
                ColorIndex = colorIndex,
                Scale = scale,
                DisplayName = displayName,
                Description = description,
                CreatedBy = createdBy,
                AttributesJson = attributes   // ★ 现在字典里面有东西了
            };
            dto.AttributesJson = attributes;
            var (success, message, storageId) = await UploadFileAndSaveToDatabase(dto);

            if (success)
                LogManager.Instance.LogInfo($"[上传成功] StorageId={storageId}, {message}");
            else
                LogManager.Instance.LogWarning($"[上传失败] {message}");
        }
        
        /// <summary>
        /// 上传文件到服务器并保存元数据
        /// </summary>
        /// <param name="dto">包含文件路径和分类信息的 DTO</param>
        /// <returns>上传结果（成功/失败 + 消息）</returns>
        public async Task<(bool Success, string Message, int StorageId)> UploadFileAndSaveToDatabase(ImportEntityDto dto)
        {
            // ========== 1. 参数校验 ==========
            if (dto == null)
            {
                LogManager.Instance.LogInfo("DTO 参数不能为空;");
                return (false, "DTO 参数不能为空", 0);
            }

            if (string.IsNullOrWhiteSpace(dto.FileStorage?.FilePath) || !File.Exists(dto.FileStorage?.FilePath))
            {
                LogManager.Instance.LogInfo("DWG 文件不存在，请检查路径;");
                return (false, "DWG 文件不存在，请检查路径", 0);
            }

            if (dto.FileStorage?.CategoryId <= 0)
            {
                LogManager.Instance.LogInfo("categoryId 必须为有效的正整数;");
                return (false, "categoryId 必须为有效的正整数", 0);
            }

            // ========== 2. 构建服务端 URL ==========
            string serverIp = VariableDictionary._serverIP ?? "127.0.0.1";
            int serverPort = VariableDictionary._apiPort > 0 ? VariableDictionary._apiPort : 10010;
            string baseUrl = $"http://{serverIp}:{serverPort}";
            string uploadUrl = $"{baseUrl}/api/graphics/upload";

            LogManager.Instance.LogInfo($"[Upload] 目标地址: {uploadUrl}");

            // ========== 3. 构建 MultipartFormDataContent ========== string attributesJsonString = "{}";
            try
            {
                using var multipartFormDataform = new MultipartFormDataContent();

                // 3.1 添加 DWG 主文件（字段名 "dwgFile"，必填）
                var dwgStream = new FileStream(dto.FileStorage?.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                var dwgContent = new StreamContent(dwgStream);
                dwgContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
                multipartFormDataform.Add(dwgContent, "dwgFile", Path.GetFileName(dto.FileStorage?.FilePath));

                // 3.2 添加预览图（字段名 "previewFile"，可选）
                FileStream? previewStream = null;
                if (!string.IsNullOrWhiteSpace(dto.PreviewImagePath) && File.Exists(dto.PreviewImagePath))
                {
                    previewStream = new FileStream(dto.PreviewImagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    var previewContent = new StreamContent(previewStream);
                    string previewExt = Path.GetExtension(dto.PreviewImagePath).ToLowerInvariant();
                    string mime = previewExt switch
                    {
                        ".png" => "image/png",
                        ".jpg" or ".jpeg" => "image/jpeg",
                        ".bmp" => "image/bmp",
                        ".gif" => "image/gif",
                        _ => "application/octet-stream"
                    };
                    previewContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(mime);
                    multipartFormDataform.Add(previewContent, "previewFile", Path.GetFileName(dto.PreviewImagePath));
                }

                // 3.3 添加表单字段（全部使用 StringContent，Key 必须与服务端 Request.Form 一致）
                multipartFormDataform.Add(new StringContent(dto.CategoryId.ToString()), "categoryId");

                if (!string.IsNullOrWhiteSpace(dto.CategoryType))
                    multipartFormDataform.Add(new StringContent(dto.CategoryType), "categoryType");
                else
                    multipartFormDataform.Add(new StringContent("sub"), "categoryType");

                if (!string.IsNullOrWhiteSpace(dto.BlockName))
                    multipartFormDataform.Add(new StringContent(dto.BlockName), "blockName");

                if (!string.IsNullOrWhiteSpace(dto.LayerName))
                    multipartFormDataform.Add(new StringContent(dto.LayerName), "layerName");

                if (dto.ColorIndex.HasValue)
                    multipartFormDataform.Add(new StringContent(dto.ColorIndex.Value.ToString()), "colorIndex");

                if (dto.Scale.HasValue && dto.Scale.Value > 0)
                    multipartFormDataform.Add(new StringContent(dto.Scale.Value.ToString()), "scale");
                else
                    multipartFormDataform.Add(new StringContent("1.0"), "scale");

                if (!string.IsNullOrWhiteSpace(dto.DisplayName))
                    multipartFormDataform.Add(new StringContent(dto.DisplayName), "displayName");
                else
                    multipartFormDataform.Add(new StringContent(Path.GetFileNameWithoutExtension(dto.FileStorage.FilePath)), "displayName");

                if (!string.IsNullOrWhiteSpace(dto.Description))
                    multipartFormDataform.Add(new StringContent(dto.Description), "description");

                if (!string.IsNullOrWhiteSpace(dto.CreatedBy))
                    multipartFormDataform.Add(new StringContent(dto.CreatedBy), "createdBy");
                else
                    multipartFormDataform.Add(new StringContent(VariableDictionary._userName ?? "System"), "createdBy");

                var json = JsonConvert.SerializeObject(dto.AttributesJson ?? new Dictionary<string, string>());
                multipartFormDataform.Add(new StringContent(json), "attributesJson");

                // ========== 4. 发送请求 ==========
                LogManager.Instance.LogInfo($"[Upload] 开始上传: {Path.GetFileName(dto.FilePath)}, 分类ID={dto.CategoryId}");

                var response = await _uploadHttpClient.PostAsync(uploadUrl, multipartFormDataform);
                var responseBody = await response.Content.ReadAsStringAsync();

                // ========== 5. 解析响应 ==========
                if (response.IsSuccessStatusCode)
                {
                    // 服务端成功返回格式: { success: true, message: "上传成功", storageId: 123, ... }
                    try
                    {
                        var result = JObject.Parse(responseBody);
                        int storageId = result["storageId"]?.Value<int>() ?? 0;
                        string message = result["message"]?.Value<string>() ?? "上传成功";
                        LogManager.Instance.LogInfo($"[Upload] 成功: StorageId={storageId}, {message}");
                        return (true, message, storageId);
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogWarning($"[Upload] 解析成功响应失败: {ex.Message}, 原始返回: {responseBody}");
                        return (true, "上传成功（解析响应异常）", 0);
                    }
                }
                else
                {
                    // 服务端返回错误
                    string errorMsg;
                    try
                    {
                        var errResult = JObject.Parse(responseBody);
                        errorMsg = errResult["message"]?.Value<string>()
                                   ?? errResult["title"]?.Value<string>()
                                   ?? $"HTTP {(int)response.StatusCode}";
                    }
                    catch
                    {
                        errorMsg = $"服务器返回错误 ({(int)response.StatusCode}): {responseBody}";
                    }

                    LogManager.Instance.LogWarning($"[Upload] 失败: {errorMsg}");
                    return (false, errorMsg, 0);
                }
            }
            catch (TaskCanceledException)
            {
                LogManager.Instance.LogError("[Upload] 请求超时");
                return (false, "上传超时，请检查网络或减小文件大小", 0);
            }
            catch (HttpRequestException ex)
            {
                LogManager.Instance.LogError($"[Upload] 网络异常: {ex.Message}");
                return (false, $"无法连接到服务器 ({serverIp}:{serverPort})，请检查网络", 0);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"[Upload] 未知异常: {ex.Message}");
                return (false, $"上传异常: {ex.Message}", 0);
            }
            finally
            {
                // 清理预留给 GC 处理（FileStream 由 using 管理或用 try-finally 关闭）
                // 注意：dwgStream 和 previewStream 会被 MultipartFormDataContent 在 Dispose 时处理
            }
        }
        #endregion


        #region 各专业按键

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
            //SaveDrawingConfig();
        }
        #endregion

        #region 共用图按键处理方法...

        private void 共用条件_Btn_Click(object sender, RoutedEventArgs e)
        {
            
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
        

        #region 分类\文件列表\文件属性管理相关

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
                List<FileStorage> files = new List<FileStorage>();

                //LogManager.Instance.LogInfo($"开始加载分类 {categoryNode.Id} ({categoryNode.DisplayText}) 的文件");
                LogManager.Instance.LogInfo($"开始加载分类 {categoryNode.Id} ({categoryNode.DisplayText}) 的文件");

                if (categoryNode.Level == 0 && categoryNode.Data is CadCategory category)
                {
                    // 主分类
                    LogManager.Instance.LogInfo($"加载主分类 {category.Name} (ID: {category.Id}) 的文件");
                    files = await _graphicApiService.GetFilesByCategoryIdAsync(category.Id, "main");
                }
                else if (categoryNode.Data is CadSubcategory subcategory)
                {
                    // 子分类

                    LogManager.Instance.LogInfo($"加载子分类 {subcategory.Name} (ID: {subcategory.Id}) 的文件");
                    files = await _graphicApiService.GetFilesByCategoryIdAsync(subcategory.Id, "sub");
                }
                else
                {
                    LogManager.Instance.LogInfo("未知的节点类型");
                    return;
                }

                LogManager.Instance.LogInfo($"从服务器查询到 {files.Count} 个文件");

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

        /// <summary>
        /// 分类储存预览键的通用方法，按优先级尝试多个属性，最后回退 ToString()
        /// </summary>
        /// <param name="storage"></param>
        /// <returns></returns>
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
        /// 辅助方法
        /// </summary>
        private void InitializeFileUploadInterface()
        {
            // 清空所有输入框
            filePath.Text = "";
            FileName.Text = "";
            FileSize.Text = "";
            viewFilePath.Text = "";
            ViewImage.Source = null;

            // 清空属性编辑网格
            CategoryPropertiesDataGrid.ItemsSource = null;

            // 重置字段
            _selectedFilePath = null;
            _selectedPreviewImagePath = null;
            _currentFileStorage = null;
            // （待清理）旧代码： _currentFileAttribute = null;
            _selectedCategoryNode = null;
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
                // 确保选中项是 FileStorage 类型
                if (StroageFileDataGrid.SelectedItem is FileStorage selectedFile)
                {
                    LogManager.Instance.LogInfo($"选中文件: {selectedFile.DisplayName} (ID: {selectedFile.Id})");
                    _selectedFileStorage = selectedFile;// 更新当前选中文件存储对象，供后续属性编辑使用

                    // 显示文件基本信息
                    DisplayFileStorageInfo(selectedFile);
                    // 加载并显示文件属性
                    await LoadAndDisplayFileAttributesAsync(selectedFile);
                    // 加载预览图片
                    var previewImage = await GetPreviewImageAsync(selectedFile);

                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"处理文件选择时出错: {ex.Message}");
                MessageBox.Show($"处理文件选择时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
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

                // 获取文件属性（新架构：优先按哈希获取主对象+JSON属性）
                var result = await _databaseManager.GetFileStorageWithAttributesByHashAsync(fileStorage.FileHash);

                // 若按 hash 未取到 JSON 属性，则按 file_id 兜底（兼容历史/替换后 hash 不一致场景）
                var attributes = result.Attributes ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (attributes.Count == 0 && fileStorage.Id > 0)
                {
                    var fallbackById = await _databaseManager.GetAttributesJsonByFileIdAsync(fileStorage.Id, fileStorage.FileAttributeId);
                    if (fallbackById.Count > 0)
                    {
                        attributes = fallbackById;
                        LogManager.Instance.LogInfo($"CategoryPropertiesDataGrid 按Hash未命中属性，已按FileId兜底加载属性: FileId={fileStorage.Id}, Count={attributes.Count}");
                    }
                }

                // 准备显示数据
                var displayData = PrepareFileDisplayData(fileStorage, attributes);
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
        /// 应用当前图元属性编辑结果到数据库（主表字段 + JSON属性）。
        /// </summary>
        /// <returns>更新是否成功</returns>
        private async Task<bool> ApplySelectedFilePropertiesAsync()
        {
            try
            {
                if (_databaseManager == null)
                {
                    MessageBox.Show("数据库管理器未初始化", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }

                if (_selectedFileStorage == null || _selectedFileStorage.Id <= 0)
                {
                    MessageBox.Show("请先在图元列表中选择一个图元", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }

                var gridRows = CategoryPropertiesDataGrid?.ItemsSource as List<CategoryPropertyEditModel>;
                if (gridRows == null || gridRows.Count == 0)
                {
                    MessageBox.Show("没有可应用的属性数据", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }

                var storage = await _databaseManager.GetFileByIdAsync(_selectedFileStorage.Id).ConfigureAwait(true) ?? _selectedFileStorage;
                var existingAttributes = await _databaseManager.GetAttributesJsonByFileIdAsync(storage.Id, storage.FileAttributeId).ConfigureAwait(true);
                var attributes = new Dictionary<string, string>(existingAttributes, StringComparer.OrdinalIgnoreCase);

                void ApplyOne(string propertyDisplayName, string propertyValue)
                {
                    if (string.IsNullOrWhiteSpace(propertyDisplayName)) return;
                    if (propertyValue == null) return;

                    string key = ResolveInternalPropertyName(propertyDisplayName.Trim());
                    string val = propertyValue.Trim();

                    if (string.IsNullOrWhiteSpace(key)) return;

                    // 中文注释：主表可识别字段同步到 FileStorage 对象
                    SetFileStorageProperty(storage, propertyDisplayName.Trim(), val);

                    // 中文注释：所有编辑项都回写到 JSON 属性字典，保证动态属性不丢失
                    attributes[key] = val;
                }

                foreach (var row in gridRows)
                {
                    ApplyOne(row.PropertyName1, row.PropertyValue1);
                    ApplyOne(row.PropertyName2, row.PropertyValue2);
                }

                storage.UpdatedBy = string.IsNullOrWhiteSpace(storage.UpdatedBy) ? Environment.UserName : storage.UpdatedBy;
                storage.UpdatedAt = DateTime.Now;

                var success = await _databaseManager
                    .UpdateFileStorageAndAttributesJsonAsync(storage, attributes, storage.FileAttributeId)
                    .ConfigureAwait(true);

                if (!success)
                {
                    return false;
                }

                _selectedFileStorage = storage;
                await LoadAndDisplayFileAttributesAsync(storage).ConfigureAwait(true);
                return true;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"应用图元属性到数据库时出错: {ex.Message}");
                MessageBox.Show($"应用图元属性失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        /// <summary>
        /// 将界面显示属性名反向映射为 JSON 内部属性键。
        /// </summary>
        /// <param name="displayName">界面显示名</param>
        /// <returns>内部属性键</returns>
        private string ResolveInternalPropertyName(string displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName))
            {
                return string.Empty;
            }

            switch (displayName.Trim())
            {
                case "文件名": return "FileName";
                case "显示名称": return "DisplayName";
                case "元素块名": return "ElementBlockName";
                case "图层名称":
                case "层名": return "LayerName";
                case "颜色索引": return "ColorIndex";
                case "比例": return "Scale";
                case "标题": return "Title";
                case "关键字": return "Keywords";
                case "描述": return "Description";
            }

            foreach (var item in DictionaryHelper._propertyDisplayNameMap)
            {
                if (string.Equals(item.Value, displayName, StringComparison.OrdinalIgnoreCase))
                {
                    return item.Key;
                }
            }

            return displayName;
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
                if (_selectedCategoryNode == null)
                    return;

                var nodeData = _selectedCategoryNode.Data;
                List<FileStorage> files = null;

                if (nodeData is CadCategory main)
                {
                    files = await _graphicApiService.GetFilesByCategoryIdAsync(main.Id, "main");
                }
                else if (nodeData is CadSubcategory sub)
                {
                    files = await _graphicApiService.GetFilesByCategoryIdAsync(sub.Id, "sub");
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


        #region 设置页面操作方法

        private void 应用设置按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // ──── 1. 根据 UI 状态确定最终存储路径 ────
                if (CheckBoxUseDPath.IsChecked == true)
                {
                    // 使用默认路径（程序智能选择 D/C 盘）
                    GetPath._cacheStoragePath = null;
                    TextBoxSetStoragePath.Text = string.Empty;
                    LogManager.Instance.LogInfo("[应用设置] 使用默认存储路径");
                }
                else
                {
                    // 使用手动指定的路径
                    string manualPath = TextBoxSetStoragePath.Text?.Trim();
                    if (string.IsNullOrWhiteSpace(manualPath))
                    {
                        MessageBox.Show("未指定存储路径！请先点击“指定路径…”选择文件夹，或勾选“使用默认路径”。",
                                        "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // 规范化处理
                    manualPath = Path.GetFullPath(manualPath); // 确保为绝对路径
                    GetPath._cacheStoragePath = manualPath;
                    LogManager.Instance.LogInfo($"[应用设置] 使用手动存储路径: {manualPath}");
                }
                
                // ──── 关键：设置日志目录 ────
                LogManager.Instance.SetLogRoot(GetPath._cacheStoragePath);

                // 3. 创建日志目录并重新指向日志管理器
                string logRoot = string.IsNullOrWhiteSpace(GetPath._cacheStoragePath)
                    ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GB_CADPLUS", "Logs")
                    : Path.Combine(GetPath._cacheStoragePath, "Logs");
                LogManager.Instance.LogInfo(logRoot);  // 假设添加了此方法

                // ──── 3. 创建必要的缓存子目录 ────
                try
                {
                    Directory.CreateDirectory(GetPath.PreviewCachePath);
                    Directory.CreateDirectory(GetPath.DwgCachePath);
                    LogManager.Instance.LogInfo($"[应用设置] 缓存目录已就绪: PreviewCache={GetPath.PreviewCachePath}, DwgCache={GetPath.DwgCachePath}");
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogError($"创建缓存目录失败: {ex.Message}");
                    MessageBox.Show("无法在指定目录下创建缓存文件夹，请检查权限或换用默认路径。",
                                    "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // ──── 4. 原有保存设置及重新初始化数据库 ────
                SaveSettings(); // 保存其他设置（如果有）
                ReinitializeDatabase(); // 重新初始化数据库连接，确保新路径生效
                UpdateAdminTabsVisibility(); // 更新管理员选项卡的可见性（如果权限相关）

                MessageBox.Show("设置已应用，数据库连接已更新。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用设置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void 查看日志文件夹按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. 安全获取当前日志文件路径
                if (LogManager.Instance is null)
                {
                    MessageBox.Show("日志管理器未初始化，无法获取日志路径。", "错误",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string logPath = LogManager.Instance.GetCurrentLogFilePath();

                // 2. 判断文件与目录的存在情况，决定打开方式
                bool fileExists = !string.IsNullOrWhiteSpace(logPath) && File.Exists(logPath);
                string? folderPath = string.IsNullOrWhiteSpace(logPath)
                    ? null
                    : Path.GetDirectoryName(logPath);

                bool folderExists = !string.IsNullOrWhiteSpace(folderPath) && Directory.Exists(folderPath);

                if (!fileExists && !folderExists)
                {
                    MessageBox.Show("日志文件及所在目录均不存在。", "提示",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 3. 构造启动参数
                string arguments;
                if (fileExists && folderPath != null)
                {
                    // 文件存在 → 打开文件夹并选中文件
                    // 注意：explorer 的 /select 参数要求路径用双引号包裹以支持含空格的路径
                    arguments = $"/select,\"{logPath}\"";
                }
                else
                {
                    // 文件不存在但目录存在 → 仅打开目录
                    arguments = folderPath!;
                }

                // 4. 启动资源管理器
                using (var process = new System.Diagnostics.Process())
                {
                    process.StartInfo = new System.Diagnostics.ProcessStartInfo
                    {
                        UseShellExecute = true,
                        FileName = "explorer.exe",
                        Arguments = arguments
                    };
                    process.Start();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开日志文件夹失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void 打开最新日志按钮_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                string logPath = LogManager.Instance.GetCurrentLogFilePath();

                if (string.IsNullOrWhiteSpace(logPath) || !File.Exists(logPath))
                {
                    MessageBox.Show("日志文件不存在或路径无效。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 使用系统默认程序打开文件（通常是记事本或其他关联 .log 的应用）
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = logPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开日志文件: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 手动加载并显示当前日志文件内容
        /// </summary>
        private async void 加载显示日志按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取当前日志文件路径
                string logPath = LogManager.Instance.GetCurrentLogFilePath();

                if (File.Exists(logPath))
                {
                    // 异步读取文件内容，避免阻塞 UI
                    string content = await Task.Run(() =>
                    {
                        return File.ReadAllText(logPath, Encoding.UTF8);
                    });

                    LogContentTextBlock.Text = content;
                }
                else
                {
                    LogContentTextBlock.Text = "日志文件尚未生成。";
                }

                // 自动滚动到底部
                LogScrollViewer.ScrollToBottom();
            }
            catch (Exception ex)
            {
                LogContentTextBlock.Text = $"无法读取日志: {ex.Message}";
            }
        }
        
        /// <summary>
        /// 保存设置到配置文件
        /// </summary>
        private void SaveSettings()
        {
            try
            {
                // 更新字段值

                GetPath._cacheStoragePath = TextBoxSetStoragePath.Text.Trim();                                    //缓存文件路径
                VariableDictionary._useDPath = CheckBoxUseDPath.IsChecked ?? true;                                           //用户自定义数据文件路径
                VariableDictionary._autoSync = CheckBoxAutoSync.IsChecked ?? true;                                           //自动同步判断
                VariableDictionary._syncInterval = int.TryParse(TextBoxSyncInterval.Text, out int interval) ? interval : 30; //同步间隔
                Properties.Settings.Default.ServerIP = VariableDictionary._serverIP;                                         // 服务器ip地址
                Properties.Settings.Default.ServerPort = VariableDictionary._dataBaseServerPort;                             // 服务器端口
                Properties.Settings.Default.DatabaseName = VariableDictionary._dataBaseName;                                 // 数据库名称
                Properties.Settings.Default.Username = VariableDictionary._userName;                                         // 用户名
                Properties.Settings.Default.Password = VariableDictionary._passWord;                                         // 密码
                Properties.Settings.Default.StoragePath = GetPath._cacheStoragePath;                              // 缓存路径
                Properties.Settings.Default.UseDPath = VariableDictionary._useDPath;                                         //用户自定义数据文件路径
                Properties.Settings.Default.AutoSync = VariableDictionary._autoSync;                                         // 自动同步判断
                Properties.Settings.Default.SyncInterval = VariableDictionary._syncInterval;                                 // 同步间隔
                Properties.Settings.Default.Save();                                                                          // 保存到用户配置文件

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
                    filePath.Text = openFileDialog.FileName;
                    LogManager.Instance.LogInfo($"选择Excel文件: {openFileDialog.FileName}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"选择Excel文件时出错: {ex.Message}");
                MessageBox.Show($"选择Excel文件时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
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
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("图元批量添加模板");

                // 创建标题样式：加粗、水平居中、浅蓝背景
                ICellStyle headerStyle = workbook.CreateCellStyle();
                IFont headerFont = workbook.CreateFont();
                headerFont.IsBold = true;
                headerStyle.SetFont(headerFont);
                headerStyle.Alignment = (NPOI.SS.UserModel.HorizontalAlignment)HorizontalAlignment.Center;
                // 设置背景色为浅蓝 (NPOI 使用索引色或自定义颜色)
                headerStyle.FillForegroundColor = IndexedColors.LightBlue.Index;
                headerStyle.FillPattern = FillPattern.SolidForeground;

                // 写入标题行
                IRow headerRow = sheet.CreateRow(0);
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    ICell cell = headerRow.CreateCell(i);
                    cell.SetCellValue(dataTable.Columns[i].ColumnName);
                    cell.CellStyle = headerStyle;
                }

                // 写入数据行
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    IRow dataRow = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        object value = dataTable.Rows[i][j];
                        ICell cell = dataRow.CreateCell(j);
                        // 根据数据类型写入，避免异常
                        if (value == null || value == DBNull.Value)
                        {
                            cell.SetCellValue(string.Empty);
                        }
                        else if (value is double || value is float || value is decimal)
                        {
                            cell.SetCellValue(Convert.ToDouble(value));
                        }
                        else if (value is int || value is long || value is short)
                        {
                            cell.SetCellValue(Convert.ToDouble(value)); // NPOI 也支持 SetCellValue(double)
                        }
                        else if (value is DateTime dt)
                        {
                            cell.SetCellValue(dt);
                        }
                        else
                        {
                            cell.SetCellValue(value.ToString());
                        }
                    }
                }

                // 自动调整列宽
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    sheet.AutoSizeColumn(col);
                }

                // 保存文件
                using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
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


        #region 导入图元按键区

        /// <summary>
        /// 添加当前图形入库“从当前图形拾取”按钮的点击事件处理器
        /// </summary>
        private async void ImportFromSelection_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 检查数据库连接
                if (this._databaseManager == null || !this._databaseManager.IsDatabaseAvailable)
                {
                    int num1 = (int)MessageBox.Show("数据库未连接，无法执行导入操作。", "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
                }
                else if (this._selectedCategoryNode == null)
                {
                    int num2 = (int)MessageBox.Show("请先在左侧的分类树中选择一个要导入的目标分类。", "提示", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
                else
                {
                    ImportEntityDto dto = (ImportEntityDto)null; // 定义一个变量，用于存储从CAD选择和读取的图元信息
                    try
                    {
                        dto = SelectionImportHelper.PickAndReadEntity(); // 调用一个辅助方法来处理CAD的选择和读取逻辑，这个方法需要你根据实际情况来实现
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogError("CAD 选择或读取图元时失败: " + ex.Message);
                        int num3 = (int)MessageBox.Show("从当前图形读取图元失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
                        return;
                    }
                    if (dto == null)
                    {
                        LogManager.Instance.LogInfo("用户取消了实体选择或未选择有效实体。");
                    }
                    else
                    {
                        dto.FileStorage.CategoryId = this._selectedCategoryNode.Id; // 根据当前选中的分类节点设置CategoryId
                        dto.FileStorage.CategoryType = this._selectedCategoryNode.Level == 0 ? "main" : "sub"; // 根据节点层级设置CategoryType
                        this.ApplyAttributeTextToDto(dto); // 将FileAttribute中的备注文本解析并应用到FileStorage的属性中，确保DTO包含完整的信息
                        Window owner = Window.GetWindow((DependencyObject)this); // 获取当前窗口作为导入确认窗口的Owner，确保模态显示
                        ImportConfirmWindow confirmWindow = new ImportConfirmWindow(dto, this); //  创建导入确认窗口，并传入DTO和当前窗口的引用，以便在确认后回写数据
                        if (owner != null)
                            confirmWindow.Owner = owner;// 设置Owner属性，确保窗口模态显示在当前窗口之上
                        bool? dialogResult = new bool?();// 显示导入确认窗口，并捕获可能的异常，确保即使窗口显示失败也能给用户反馈
                        try
                        {
                            dialogResult = confirmWindow.ShowDialog(); // 显示窗口并等待用户操作，结果将决定是否继续执行导入后的刷新逻辑
                        }
                        catch (Exception ex)
                        {
                            LogManager.Instance.LogError("显示导入确认窗口失败: " + ex.Message);
                            int num4 = (int)MessageBox.Show("打开确认界面失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
                            return;
                        }
                        if (dialogResult.GetValueOrDefault()) // 如果用户在确认窗口中点击了确认（OK），则继续执行刷新逻辑，否则取消
                        {
                            try
                            {
                                await this.LoadFilesForCategoryAsync(this._selectedCategoryNode); // 刷新当前分类下的文件列表
                            }
                            catch (Exception ex)
                            {
                                LogManager.Instance.LogError("导入后刷新列表失败: " + ex.Message);
                                int num5 = (int)MessageBox.Show("导入完成，但刷新列表失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                            }
                        }
                        dto = (ImportEntityDto)null; // 释放变量以释放对象
                        owner = (Window)null; // 释放变量以释放对象
                        confirmWindow = (ImportConfirmWindow)null; // 释放变量以释放对象
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError("从CAD拾取导入失败: " + ex.Message);
                int num = (int)MessageBox.Show("操作失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        /// <summary>
        /// 为ImportConfirmWindow提供一个设置内部状态的公共方法
        /// </summary>
        public void SetSelectedFileForImport(ImportEntityDto dto)
        {
            // 把确认后的导入对象回写到主窗口状态
            _selectedFilePath = dto.FileStorage.FilePath;
            _selectedPreviewImagePath = dto.PreviewImagePath;
            _currentFileStorage = dto.FileStorage;
        }

        /// <summary>
        /// 确认导入
        /// </summary>
        /// <param name="dto"></param>
        private void ApplyAttributeTextToDto(ImportEntityDto dto)
        {
            if (dto == null || dto.FileStorage == null) return;

            // 1) 从 Remarks（或其他字符串）提取 key/value 对
            string raw = "";
            dto.AttributesJson.TryGetValue("Remarks", out raw);
            if (string.IsNullOrWhiteSpace(raw)) raw = string.Empty;

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

            // 2) 准备候选属性字典（FileStorage 的 PropertyInfo）
            var fsType = typeof(FileStorage);
            var fsProps = fsType.GetProperties().Where(p => p.CanWrite).ToList();

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

                // 将值写入 JSON 字典
                dto.AttributesJson[tag] = val;
                matched = true;

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
                foreach (var prop in fsProps)
                {
                    var pName = prop.Name;
                    if (DictionaryHelper._propertyDisplayNameMap.TryGetValue(pName, out var dispName))
                    {
                        if (Normalize(dispName).IndexOf(nTag, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            nTag.IndexOf(Normalize(dispName), StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            TrySetPropertyValue(dto.FileStorage, prop, val);
                            matched = true;
                            break;
                        }
                    }
                    else
                    {
                        if (Normalize(pName).IndexOf(nTag, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            nTag.IndexOf(Normalize(pName), StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            TrySetPropertyValue(dto.FileStorage, prop, val);
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
        /// <param name="target">目标对象</param>
        /// <param name="prop">属性信息</param>
        /// <param name="rawValue">原始值</param>        
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

        #endregion

        
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

        private async void 表格同步_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("未找到活动的 AutoCAD 文档。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                using (doc.LockDocument())
                {
                    var ed = doc.Editor;

                    var peo = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\n请选择要同步的表格（Table）：");
                    peo.SetRejectMessage("\n请选择一个表格对象。");
                    peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Table), true);

                    var per = ed.GetEntity(peo); // 把选中的对象获取为实体对象
                    if (per.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        return;
                    //开始事务
                    using (var tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        // 获取表格对象
                        var table = tr.GetObject(per.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Table; 
                        if (table == null)
                        {
                            MessageBox.Show("选中的对象不是表格。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                            return;
                        }

                        // 读取前几行作为表头用于判断（最多读取前3行）
                        int headerRowsToCheck = Math.Min(3, table.Rows.Count);
                        var headerTextSb = new System.Text.StringBuilder();  // 用于存储表头文本
                        for (int r = 0; r < headerRowsToCheck; r++)
                        {
                            for (int c = 0; c < table.Columns.Count; c++)
                            {
                                try
                                {
                                    var cellText = (table.Cells[r, c].TextString ?? string.Empty).Trim();
                                    if (!string.IsNullOrEmpty(cellText))
                                    {
                                        headerTextSb.Append(cellText).Append(" ");
                                    }
                                }
                                catch
                                {
                                    // 忽略单元格读取错误
                                }
                            }
                        }
                        string headerText = headerTextSb.ToString();

                        // -----------------------------------------------------------------
                        // 关键词定义（可根据实际表头增加/调整）
                        // -----------------------------------------------------------------
                        var pipeKeywords = new[] {
                            "PIPELINETITLE", "PIPE_OD", "PIPE_ID", "PIPE_THK",
                            "SCHEDULE", "PIPE_TYPE", "PIPE_CLASS", "PIPE_MATL"
                        };
                        var equipKeywords = new[] {
                             "阀门", "法兰", "电机", "泵", "DRIVE_MODE"
                         };
                        // ★ 新增：计算表关键词（涵盖脱硫系统计算、泵计算、管道计算等）
                        var calcKeywords = new[] {
                             "脱硫系统", "入口SO2", "出口SO2", "大气压", "出口温度",
                             "标况烟气量", "工况烟气量", "烟囱出口风速", "烟囱出口直径",
                             "浆液循环总量", "循环泵选型", "氧化区有效容积", "氧化区直径",
                             "循环泵流量", "扬程", "电机系数", "计算功率", "推荐电机功率",
                             "进口管道尺寸", "出口管道尺寸", "介质", "使用条件", "流速",
                             "管道尺寸", "核算流速", "氧化风机选型", "氧化风量", "氧化风压",
                             "石灰石浆液泵", "石灰石纯度", "浆液密度", "系统数量",
                             "运行时长", "泵台数", "效率", "电机功率选型"
                         };

                        // 判断表类型
                        bool looksLikePipe = pipeKeywords.Any(k => !string.IsNullOrWhiteSpace(k) &&
                            headerText.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0);
                        bool looksLikeEquip = equipKeywords.Any(k => !string.IsNullOrWhiteSpace(k) &&
                            headerText.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0);
                        bool looksLikeCalc = calcKeywords.Any(k => !string.IsNullOrWhiteSpace(k) &&
                            headerText.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0);

                        string chosenCommand = null;

                        // ★ 优先处理计算表（因为计算表可能同时匹配管道/设备关键词）
                        if (looksLikeCalc)
                        {
                            // 收集参数：忽略设备名列，只取参数名和值
                            var calcProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            int startRow = 0;   // 如果第一行是标题，可改为 1
                            for (int r = startRow; r < table.Rows.Count; r++)
                            {
                                if (table.Columns.Count < 2) continue;

                                // 根据列数决定列索引：
                                // 3列：参数名在列1，值在列2
                                // 2列：参数名在列0，值在列1
                                int nameCol = (table.Columns.Count >= 3) ? 1 : 0;
                                int valueCol = (table.Columns.Count >= 3) ? 2 : 1;
                                string paramName = (table.Cells[r, nameCol].TextString ?? string.Empty).Trim();
                                string paramValue = (table.Cells[r, valueCol].TextString ?? string.Empty).Trim();

                                if (!string.IsNullOrEmpty(paramName))
                                    calcProps[paramName] = paramValue;   // 重复参数时后面的覆盖前面的
                            }

                            tr.Commit();

                            if (calcProps.Count == 0)
                            {
                                ed.WriteMessage("\n表格中没有有效的参数。");
                                return;
                            }

                            ed.WriteMessage($"\n共提取 {calcProps.Count} 个参数。");
                            await SyncCalcPropsToBlocksAsync(calcProps);
                            return;
                        }
                        else if (looksLikePipe && !looksLikeEquip)
                        {
                            chosenCommand = "SyncTableToEntities ";
                        }
                        else if (looksLikeEquip && !looksLikePipe)
                        {
                            chosenCommand = "SyncDeviceTableToBlocks ";
                        }
                        else if (looksLikePipe && looksLikeEquip)
                        {
                            // 同时匹配管道和设备特征，询问用户
                            var result = MessageBox.Show(
                                "未能自动判断表类型或同时匹配到管道/设备特征。\n请选择要执行的同步类型：\n\n[是] - 设备表同步（表 -> 图元）\n[否] - 管道表同步（表 -> 管道实体）\n[取消] - 取消操作",
                                "请选择表类型",
                                MessageBoxButton.YesNoCancel,
                                MessageBoxImage.Question);

                            if (result == MessageBoxResult.Cancel)
                            {
                                return;
                            }
                            else if (result == MessageBoxResult.Yes)
                            {
                                chosenCommand = "SyncDeviceTableToBlocks ";
                            }
                            else
                            {
                                chosenCommand = "SyncTableToEntities ";
                            }
                        }
                        else
                        {
                            // 没有任何关键词匹配，提示用户选择所有可能的同步方式
                            var result = MessageBox.Show(
                                "未能自动判断表类型。\n请选择要执行的同步类型：\n\n[是] - 设备表同步\n[否] - 管道表同步\n[取消] - 取消操作\n（若为计算表，暂无法处理，请检查命令）",
                                "请选择表类型",
                                MessageBoxButton.YesNoCancel,
                                MessageBoxImage.Question);

                            if (result == MessageBoxResult.Cancel) return;
                            if (result == MessageBoxResult.Yes)
                                chosenCommand = "SyncDeviceTableToBlocks ";
                            else
                                chosenCommand = "SyncTableToEntities ";
                        }

                        tr.Commit();

                        // 发送命令到 AutoCAD 命令行
                        if (!string.IsNullOrEmpty(chosenCommand))
                        {
                            ed.WriteMessage($"\n将执行: {chosenCommand.Trim()}");
                            Env.Document.SendStringToExecute(chosenCommand, false, false, false);
                        }
                        else
                        {
                            ed.WriteMessage("\n未识别可用的同步命令，操作取消。");
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"表格同步失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 同步计算表
        /// </summary>
        /// <param name="props"> </param>
        /// <returns></returns>
        private async Task SyncCalcPropsToBlocksAsync(Dictionary<string, string> props)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            // 选择要同步的块参照
            var selOpts = new PromptSelectionOptions
            {
                RejectPaperspaceViewport = true,
                MessageForAdding = "\n请选择要同步属性的图元（块参照）："
            };
            var selRes = ed.GetSelection(selOpts);
            if (selRes.Status != PromptStatus.OK || selRes.Value.Count == 0)
            {
                ed.WriteMessage("\n未选择任何图元，操作取消。");
                return;
            }

            // 预先将中文参数名翻译为英文 Tag，构建 Tag -> 值 的字典（加速匹配）
            var tagToValue = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in props)
            {
                string targetTag = FindCanonicalTagByChineseName(kvp.Key); // 尝试获取英文标签
                if (!string.IsNullOrEmpty(targetTag) && !tagToValue.ContainsKey(targetTag))
                {
                    // 同时应用单位换算
                    string finalValue = PerformUnitConversionIfNeeded(targetTag, kvp.Value);
                    tagToValue[targetTag] = finalValue;
                }
            }

            if (tagToValue.Count == 0)
            {
                ed.WriteMessage("\n没有参数能在字典中找到对应的英文标签，请检查 DictionaryHelper._canonicalToAliases 配置。");
                return;
            }

            int matchedBlocks = 0;
            int assignedAttributes = 0;

            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in selRes.Value.GetObjectIds())
                {
                    if (!(tr.GetObject(id, OpenMode.ForWrite) is BlockReference br)) continue;

                    BlockTableRecord btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (btr == null || !btr.HasAttributeDefinitions) continue;

                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        if (!attId.IsValid || attId.IsErased) continue;
                        var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                        if (attRef == null) continue;

                        if (tagToValue.TryGetValue(attRef.Tag, out string newValue))
                        {
                            attRef.TextString = newValue;
                            assignedAttributes++;
                        }
                    }
                    matchedBlocks++;
                }

                tr.Commit();
            }

            ed.WriteMessage($"\n同步完成：处理 {matchedBlocks} 个图元，共赋值 {assignedAttributes} 个属性。");
        }
        

        /// <summary>
        /// 根据中文列名从 DictionaryHelper._canonicalToAliases 中反向查找英文属性标签
        /// </summary>
        /// <param name="chineseColumnName">表格中的中文列名</param>
        /// <returns>英文属性 Tag，若未找到则返回 null</returns>
        private string FindCanonicalTagByChineseName(string chineseColumnName)
        {
            // 精确匹配 + 忽略大小写
            foreach (var kvp in DictionaryHelper._canonicalToAliases)
            {
                foreach (var alias in kvp.Value)
                {
                    if (string.Equals(alias, chineseColumnName, StringComparison.OrdinalIgnoreCase))
                        return kvp.Key;
                }
            }

            // 如果精确匹配失败，尝试包含匹配（宽松模式）
            string trimmed = chineseColumnName.Trim();
            foreach (var kvp in DictionaryHelper._canonicalToAliases)
            {
                foreach (var alias in kvp.Value)
                {
                    if (alias.IndexOf(trimmed, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        trimmed.IndexOf(alias, StringComparison.OrdinalIgnoreCase) >= 0)
                        return kvp.Key;
                }
            }
            return null;
            
        }
        /// <summary>
        /// 根据属性标签判断是否需要进行单位换算
        /// </summary>
        /// <param name="tag">标题</param>
        /// <param name="rawValue">行值</param>
        /// <returns></returns>
        private string PerformUnitConversionIfNeeded(string tag, string rawValue)
        {
            if (tag.Equals("ATM_PRESSURE", StringComparison.OrdinalIgnoreCase) && double.TryParse(rawValue, out double pa))
                return (pa / 1000.0).ToString("F2");  // Pa → kPa
            if (tag.Equals("STACK_EXIT_DIA", StringComparison.OrdinalIgnoreCase) && double.TryParse(rawValue, out double mm))
                return (mm / 1000.0).ToString("F3");  // mm → m
            return rawValue;
        }
      
        private void 导出表格_Click(object sender, RoutedEventArgs e)
        {
            Env.Document.SendStringToExecute("ExportTableToExcel ", false, false, false);
            //Env.Document.SendStringToExecute("ExportTableToExcelFile ", false, false, false);
            
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
            // 设置绘图比例到全局变量，供命令使用
            Env.Document.SendStringToExecute("DrawInletPipeByClicks ", false, false, false);
        }

        private void 绘制出口管道_Click(object sender, RoutedEventArgs e)
        {
            Env.Document.SendStringToExecute("DrawOutletPipeByClicks ", false, false, false);
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
                _workingDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                               "GB_Tools");
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

                    // 创建工作簿和工作表
                    IWorkbook workbook = new XSSFWorkbook();
                    ISheet sheet = workbook.CreateSheet("图层配置");

                    // 写入表头
                    IRow headerRow = sheet.CreateRow(0);
                    headerRow.CreateCell(0).SetCellValue("序号");
                    headerRow.CreateCell(1).SetCellValue("图层名称");
                    headerRow.CreateCell(2).SetCellValue("开关");
                    headerRow.CreateCell(3).SetCellValue("冻结");
                    headerRow.CreateCell(4).SetCellValue("颜色索引");
                    headerRow.CreateCell(5).SetCellValue("删除");

                    // 写入数据行
                    for (int i = 0; i < _layerInfo.Count; i++)
                    {
                        var layer = _layerInfo[i];
                        IRow dataRow = sheet.CreateRow(i + 1);   // 第 i+1 行 (因为第0行是表头)

                        dataRow.CreateCell(0).SetCellValue(layer.Index);
                        dataRow.CreateCell(1).SetCellValue(layer.LayerName);
                        dataRow.CreateCell(2).SetCellValue(layer.IsOn);
                        dataRow.CreateCell(3).SetCellValue(layer.IsFrozen);
                        dataRow.CreateCell(4).SetCellValue(layer.ColorIndex);
                        dataRow.CreateCell(5).SetCellValue(layer.IsDelete);
                    }

                    // 保存文件
                    using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(fs);
                    }

                    return true;
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

                    // 使用 NPOI 打开 Excel 文件
                    IWorkbook workbook;
                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        workbook = new XSSFWorkbook(fs);
                    }

                    // 获取第一个工作表
                    ISheet sheet = workbook.GetSheetAt(0);
                    if (sheet == null)
                        return new List<LayerInfo>();

                    var layers = new List<LayerInfo>();

                    // 从第1行开始（索引0是表头，所以从索引1开始）
                    int rowCount = sheet.LastRowNum; // 最后一行的索引
                    if (rowCount < 1) return new List<LayerInfo>(); // 至少有一行数据

                    for (int rowIdx = 1; rowIdx <= rowCount; rowIdx++) // rowIdx 从1开始，0为表头
                    {
                        IRow row = sheet.GetRow(rowIdx);
                        if (row == null) continue; // 跳过空行

                        try
                        {
                            // 注意：NPOI 的列索引从0开始，对应原 EPPlus 的 row, col (col从1开始)，所以要减1
                            // 原EPPlus: worksheet.Cells[row, 1] -> NPOI: row.GetCell(0)
                            int index = Convert.ToInt32(GetCellValue(row, 0) ?? 0);
                            string name = GetCellValue(row, 1)?.ToString() ?? "";
                            bool isOn = Convert.ToBoolean(GetCellValue(row, 2) ?? true);
                            bool isFrozen = Convert.ToBoolean(GetCellValue(row, 3) ?? false);
                            short colorIndex = Convert.ToInt16(GetCellValue(row, 4) ?? 7);
                            bool toDelete = Convert.ToBoolean(GetCellValue(row, 5) ?? false);

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
                            ed?.WriteMessage($"\n读取第{rowIdx + 1}行数据时出错: {rowEx.Message}");
                        }
                    }

                    // 按名称排序
                    layers.Sort((x, y) => string.Compare(x.LayerName, y.LayerName, StringComparison.OrdinalIgnoreCase));
                    _layerInfo = new List<LayerInfo>(layers);
                    return layers;
                }
                catch (Exception ex)
                {
                    Editor ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                    ed?.WriteMessage($"\n加载图层配置时出错: {ex.Message}");
                    return new List<LayerInfo>();
                }
            }

            // 辅助方法：安全获取单元格值（兼容 null 单元格）
            private object GetCellValue(IRow row, int colIndex)
            {
                ICell cell = row.GetCell(colIndex);
                if (cell == null) return null;
                switch (cell.CellType)
                {
                    case CellType.Numeric:
                        return cell.NumericCellValue;
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Boolean:
                        return cell.BooleanCellValue;
                    case CellType.Formula:
                        // 公式单元格可根据需要返回计算值或公式字符串，这里返回其字符串表示
                        return cell.ToString();
                    default:
                        return cell.ToString();
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
                    if (VariableDictionary.wpfTextBoxScale == 0)
                    {
                        VariableDictionary.wpfTextBoxScale = 1;
                    }
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
                    // 插入成功后立即删除该次临时文件，避免累积
                    TryDeleteTempFile(_lastSavedPipeTablePath);
                    _lastSavedPipeTablePath = string.Empty;
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
        /// 文件名格式：{layerPart}_{yyyyMMdd_HHmmss}.dwg，保存在 %TEMP%\GB_NewCadPlus_IV 下。ParseLengthValueFromAttribute
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
                string tempDir = Path.Combine(Path.GetTempPath(), "GB_NewCadPlus_IV", buttonFolderName);
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
        /// 判断是否为管理员用户（sa/SYSDBA/admin）
        /// </summary>
        private static bool IsAdminUser(string username)
        {
            if (string.IsNullOrWhiteSpace(username)) return false; // 空用户名不是管理员
            var low = username.Trim().ToLowerInvariant(); // 规范为小写比较
            return low == "sa" || low == "sysdba" || low == "admin"; // 当前阶段的三个默认管理员用户名
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
                    // 插入成功后立即删除该次临时文件，避免累积
                    TryDeleteTempFile(_lastSavedPipeTablePath);
                    _lastSavedPipeTablePath = string.Empty;
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
                var cfgPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GB_NewCadPlus_IV", "login_config.json");
                string host = "127.0.0.1";
                string port = "5236";
                if (System.IO.File.Exists(cfgPath))
                {
                    var json = System.IO.File.ReadAllText(cfgPath);
                    var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                    var cfg = ser.Deserialize<LoginConfig>(json);
                    if (cfg != null)
                    {
                        if (!string.IsNullOrWhiteSpace(cfg.ServerIP)) host = cfg.ServerIP;
                        if (!string.IsNullOrWhiteSpace(cfg.DataBaseserverPort)) port = cfg.DataBaseserverPort;
                    }
                }

                // 中文注释：部门/人员模块必须使用数据库物理连接账号，不能使用应用登录用户名。
                var dbType = (VariableDictionary._databaseType ?? "DM").ToUpperInvariant();
                var dbUser = string.IsNullOrWhiteSpace(VariableDictionary._dbUserName)
                    ? (dbType == "MYSQL" ? "root" : "SYSDBA")
                    : VariableDictionary._dbUserName.Trim();
                var dbPwd = string.IsNullOrWhiteSpace(VariableDictionary._dbPassWord)
                    ? (dbType == "MYSQL" ? "123456" : "675756SGBsgb")
                    : VariableDictionary._dbPassWord;

                LogManager.Instance.LogInfo($"部门服务使用服务器 API: host={host}, apiPort={VariableDictionary._apiPort}");
                VariableDictionary._serverIP = host;
                if (VariableDictionary._apiPort <= 0)
                    VariableDictionary._apiPort = 10010;
                RefreshDepartmentsAsync();
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "初始化失败：" + ex.Message;
            }
        }
                
        /// <summary>
        /// 从服务器获取部门及部门用户数量
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnSync_Click(object sender, RoutedEventArgs e)
        {
            TxtStatus.Text = "正在获取部门...";
            if (sender is Button button)
                button.IsEnabled = false;

            try
            {
                if (string.IsNullOrWhiteSpace(VariableDictionary._serverIP))
                    VariableDictionary._serverIP = "127.0.0.1";
                if (VariableDictionary._apiPort <= 0)
                    VariableDictionary._apiPort = 10010;

                LogManager.Instance.LogInfo($"开始获取部门：API={VariableDictionary._serverIP}:{VariableDictionary._apiPort}");
                var departments = await _departmentApiService.GetDepartmentsWithCountsAsync();
                DepartmentsGrid.ItemsSource = departments;
                DepartmentsGrid.Items.Refresh();

                string message = $"已获取 {departments.Count} 个部门。";
                LogManager.Instance.LogInfo($"获取部门完成：{message}");
                TxtStatus.Text = message;
                MessageBox.Show(message, "获取完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "获取部门失败：" + ex.Message;
                LogManager.Instance.LogInfo("获取部门失败：" + ex);
                MessageBox.Show(TxtStatus.Text, "获取失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (sender is Button currentButton)
                    currentButton.IsEnabled = true;
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

                    var response = _apiService.AddDepartmentAsync(result.Name, result.RealName, result.Description, result.SortOrder, mgrId).GetAwaiter().GetResult();
                    if (response.Success)
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

                    var response = _apiService.UpdateDepartmentAsync(sel.Id, edited.Name, edited.RealName, edited.Description, edited.SortOrder, mgrId, edited.IsActive).GetAwaiter().GetResult();
                    if (response.Success)
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
            if (MessageBox.Show($"确认删除部门：{sel.RealName} ?\n删除后该部门下用户将被置为未分配。", "确认删除", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;
            var response = _apiService.DeleteDepartmentAsync(sel.Id).GetAwaiter().GetResult();
            if (response.Success)
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
                RealName = initial?.RealName ?? string.Empty,
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
            var tbDisplay = new System.Windows.Controls.TextBox { Text = model.RealName, Margin = new System.Windows.Thickness(4) };
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
                model.RealName = tbDisplay.Text.Trim();
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
        /// 选定用户后分配用户所在部门
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnAssignUser_Click(object sender, RoutedEventArgs e)
        {
            var sel = DepartmentsGrid.SelectedItem as DepartmentModel;
            if (sel == null) { MessageBox.Show("请先选择部门", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            var username = (TxtSearchUser.Text ?? "").Trim();
            if (string.IsNullOrEmpty(username)) { MessageBox.Show("请输入要分配的用户名", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); return; }

            var response = _apiService.AssignUserToDepartmentAsync(username, sel.Id).GetAwaiter().GetResult();
            if (response.Success)
            {
                LoadUsersForDepartment(sel.Id);
                MessageBox.Show($"用户 {username} 已分配到 {sel.RealName}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show($"分配失败，检查用户名是否存在或数据库状态。", "失败", MessageBoxButton.OK, MessageBoxImage.Error);
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


        #region 总表计算模式

        #region CSV表转换工具
        /// <summary>
        /// 公式合法性批量校验（增强版）：
        /// - 归一化 OCR 噪声 token（IC16/lC16/|C16 -> C16）
        /// - 归一化全角符号/异常引号
        /// - 无法识别的脏 token 直接忽略，减少误报
        /// </summary>
        private List<string> ValidateCalcSectionsFormulaReferences(List<ExcelCalcSection> sections)
        {
            var errors = new List<string>();
            if (sections == null || sections.Count == 0) return errors;

            // Grid -> 地址集合
            var cellMap = sections
                .GroupBy(s => s.GridName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    g => g.Key,
                    g => new HashSet<string>(
                        g.SelectMany(x => x.Rows)
                         .Where(r => !string.IsNullOrWhiteSpace(r.Address))
                         .Select(r => (r.Address ?? string.Empty).Trim().ToUpperInvariant()),
                        StringComparer.OrdinalIgnoreCase),
                    StringComparer.OrdinalIgnoreCase);

            // 地址 -> 所属 Grid 列表（用于“本表找不到时，尝试全局唯一归属”）
            var addressOwners = sections
                .SelectMany(s => s.Rows.Select(r => new { s.GridName, Address = (r.Address ?? string.Empty).Trim().ToUpperInvariant() }))
                .Where(x => !string.IsNullOrWhiteSpace(x.Address))
                .GroupBy(x => x.Address, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    g => g.Key,
                    g => g.Select(x => x.GridName).Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                    StringComparer.OrdinalIgnoreCase);

            foreach (var section in sections)
            {
                foreach (var row in section.Rows.Where(r => !string.IsNullOrWhiteSpace(r.Formula)))
                {
                    string formula = NormalizeFormulaForValidation((row.Formula ?? string.Empty).Trim());
                    if (formula.StartsWith("="))
                        formula = formula.Substring(1);

                    string rowTag = $"{row.GridName}!{row.Address} ({row.ParameterName})";

                    // 1) 校验跨表引用：Grid!Addr
                    var crossRefs = Regex.Matches(
                        formula,
                        @"(?<grid>[^\s\+\-\*/\(\),!]+)\s*!\s*(?<addr>[A-Z\|IL]*C\d+|[A-Z]+\d+)",
                        RegexOptions.IgnoreCase);

                    foreach (Match m in crossRefs)
                    {
                        string gridRaw = (m.Groups["grid"].Value ?? string.Empty).Trim().Trim('\'', '"', '“', '”', '‘', '’');
                        string gridToken = gridRaw.Replace(" ", "_");

                        if (gridToken.StartsWith("DataGrid", StringComparison.OrdinalIgnoreCase) &&
                            !gridToken.StartsWith("DataGrid_", StringComparison.OrdinalIgnoreCase))
                        {
                            gridToken = "DataGrid_" + gridToken.Substring("DataGrid".Length).TrimStart('_');
                        }

                        string grid = NormalizeCalcGridNameToken(gridToken);
                        if (string.IsNullOrWhiteSpace(grid)) grid = gridToken;

                        string rawAddr = (m.Groups["addr"].Value ?? string.Empty).Trim();
                        if (!TryNormalizeCellAddressToken(rawAddr, out string addr))
                            continue; // 脏 token 忽略

                        if (!cellMap.TryGetValue(grid, out var addrSet))
                        {
                            errors.Add($"[{rowTag}] 引用了不存在的表：{gridRaw} -> {grid}");
                            continue;
                        }

                        if (!addrSet.Contains(addr))
                        {
                            errors.Add($"[{rowTag}] 引用了不存在的单元格：{grid}!{addr}");
                        }
                    }

                    // 2) 本表引用校验：先去掉跨表引用，再检查本地地址
                    string localExpr = Regex.Replace(
                        formula,
                        @"(?<grid>[^\s\+\-\*/\(\),!]+)\s*!\s*(?<addr>[A-Z\|IL]*C\d+|[A-Z]+\d+)",
                        "0",
                        RegexOptions.IgnoreCase);

                    var localRawRefs = Regex.Matches(localExpr, @"\b([A-Z\|IL]*C\d+|[A-Z]+\d+)\b", RegexOptions.IgnoreCase)
                        .Cast<Match>()
                        .Select(x => x.Value)
                        .Distinct(StringComparer.OrdinalIgnoreCase);

                    var localRefs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var raw in localRawRefs)
                    {
                        if (TryNormalizeCellAddressToken(raw, out string normalized))
                            localRefs.Add(normalized);
                    }

                    if (!cellMap.TryGetValue(row.GridName, out var currentGridCells))
                    {
                        errors.Add($"[{rowTag}] 当前表不存在：{row.GridName}");
                        continue;
                    }

                    foreach (var addr in localRefs)
                    {
                        // 当前表存在，OK
                        if (currentGridCells.Contains(addr))
                            continue;

                        // 当前表不存在，尝试全局唯一归属（同一 Excel 按行号拆段时常见）
                        if (addressOwners.TryGetValue(addr, out var owners) && owners.Count == 1)
                            continue;

                        if (addressOwners.TryGetValue(addr, out owners) && owners.Count > 1)
                        {
                            errors.Add($"[{rowTag}] 本表引用地址歧义：{addr}（出现在多个表：{string.Join(" / ", owners)}）");
                            continue;
                        }

                        errors.Add($"[{rowTag}] 本表引用不存在单元格：{row.GridName}!{addr}");
                    }
                }
            }

            return errors
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        #region 优化后的CSV导出（增强版）：
        private bool FuzzyMatchCell(string missingRefId, HashSet<string> availableCells)
        {
            int idx = missingRefId.IndexOf('!');
            if (idx < 0) return false;

            string missingSheet = missingRefId.Substring(0, idx);
            string missingCell = missingRefId.Substring(idx + 1);

            // 生成极简 Key：只保留字母和数字
            string cleanMissingSheet = Regex.Replace(missingSheet, @"[^a-zA-Z0-9\u4e00-\u9fff]", "");
            string cleanMissingCell = missingCell.Replace("$", ""); // 去除 $

            foreach (var avail in availableCells)
            {
                int availIdx = avail.IndexOf('!');
                if (availIdx < 0) continue;

                string availSheet = avail.Substring(0, availIdx);
                string availCell = avail.Substring(availIdx + 1);

                string cleanAvailSheet = Regex.Replace(availSheet, @"[^a-zA-Z0-9\u4e00-\u9fff]", "");
                string cleanAvailCell = availCell.Replace("$", "");

                // 如果 Sheet 和 Cell 的纯字母数字部分都一致，则认为匹配
                if (cleanMissingCell.Equals(cleanAvailCell, StringComparison.OrdinalIgnoreCase) &&
                    cleanMissingSheet.Equals(cleanAvailSheet, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        #endregion
        


        /// <summary>
        /// 归一化公式字符串（用于校验，不改原公式）
        ///— 主要处理全角符号、中文引号、OCR 产生的异形符号
        /// </summary>
        private static string NormalizeFormulaForValidation(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;

            string s = text
                .Replace('（', '(').Replace('）', ')')
                .Replace('【', '[').Replace('】', ']')
                .Replace('〔', '[').Replace('〕', ']')
                .Replace('，', ',').Replace('。', '.')
                .Replace('！', '!')
                .Replace('“', '"').Replace('”', '"')
                .Replace('‘', '\'').Replace('’', '\'')
                .Trim();

            // 去空白，避免 OCR 把 token 打断
            s = Regex.Replace(s, @"\s+", string.Empty);

            // DataGridxxxIC16 / DataGrid动态3IC96 -> DataGridxxx!IC16
            s = Regex.Replace(
                s,
                @"(?<grid>DataGrid[^+\-*/\(\),!]+?)(?<addr>[IL\|]*C\d+)(?=$|[+\-*/\),])",
                m => $"{m.Groups["grid"].Value}!{m.Groups["addr"].Value}",
                RegexOptions.IgnoreCase);

            return s;
        }
        /// <summary>
        /// 归一化地址 token：IC16/lC16/|C16 -> C16
        /// 仅接受最终形态 C+数字；无法识别返回 false（表示脏 token，忽略）
        /// </summary>
        private static bool TryNormalizeCellAddressToken(string raw, out string normalized)
        {
            normalized = string.Empty;
            if (string.IsNullOrWhiteSpace(raw)) return false;

            string s = raw.Trim().ToUpperInvariant();
            s = s.Trim('\'', '"', '[', ']', '(', ')', '{', '}');

            // OCR: IC16 / LC16 / |C16 / IIC16 -> C16
            s = Regex.Replace(s, @"^[IL\|]+C", "C", RegexOptions.IgnoreCase);

            if (Regex.IsMatch(s, @"^C\d+$", RegexOptions.IgnoreCase))
            {
                normalized = s;
                return true;
            }

            return false;
        }

        /// <summary>
        /// 存储计算单元地址标准化
        /// </summary>
        /// <param name="rawAddress"> </param>
        /// <returns></returns>
        private static string NormalizeCalcCellAddressForStorage(string? rawAddress)
        {
            string raw = (rawAddress ?? string.Empty).Trim();
            if (TryNormalizeCellAddressToken(raw, out string normalized))
                return normalized;
            return raw.ToUpperInvariant();
        }

        /// <summary>
        /// CSV 字段转义：包含逗号/引号/换行时自动加引号并转义
        /// </summary>
        private static string EscapeCsv(string text)
        {
            if (text == null) return string.Empty;

            bool needQuote = text.Contains(",") || text.Contains("\"") || text.Contains("\r") || text.Contains("\n");
            if (!needQuote) return text;

            return "\"" + text.Replace("\"", "\"\"") + "\"";
        }
        #endregion

        #region

        /// <summary>
        /// excel 分段信息：以空行分隔，第一列设备名可合并
        /// </summary>
        private sealed class ExcelCalcSection
        {
            public string DeviceName { get; set; } = string.Empty;
            public string GridName { get; set; } = string.Empty;
            public string GroupHeader { get; set; } = string.Empty;
            public List<CalcCsvTableRow> Rows { get; set; } = new List<CalcCsvTableRow>();
        }

        /// <summary>
        /// 从总表 Excel 动态加载页面，并导出总 CSV
        /// </summary>
        private void LoadFromMasterExcelAndBuildUi(string excelPath, bool exportCsv = true)
        {
            var sections = ParseMasterExcelSections(excelPath);// 解析 Excel 分段

            var formulaErrors = ValidateCalcSectionsFormulaReferences(sections);// 公式合法性批量校验
            if (formulaErrors.Count > 0)// 如果有错误，弹窗显示前40条，并提示修正 Excel
            {
                string preview = string.Join("\n", formulaErrors.Take(40));// 40条预览限制，避免过长
                if (formulaErrors.Count > 40)// 超过40条时提示省略
                    preview += $"\n... 其余 {formulaErrors.Count - 40} 条省略";// 提示剩余数量

                MessageBox.Show(
                    "公式校验未通过，请先修正 Excel：\n\n" + preview,
                    "公式校验失败",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return;
            }

            BuildDynamicCalcGrids(sections);// 根据分段构建动态页面
                                            // 导出总 CSV（如果需要）
            if (exportCsv)
            {
                string csvPath = ExportSectionsToMasterCsv(sections);// 导出总 CSV
                LogManager.Instance.LogInfo($"Excel 转总CSV完成: {csvPath}");// 可选：弹窗提示 CSV 已生成，并提供路径
            }
            // 重新计算所有计算表格
            RecalculateAllCalcCsvTables();
        }

        /// <summary>
        /// 解析 Excel 分段：空行分隔，第一列设备名可合并
        /// 列定义：A设备 B参数名称 C参数 D单位 E值类型
        /// </summary>
        private List<ExcelCalcSection> ParseMasterExcelSections(string excelPath)
        {
            if (!File.Exists(excelPath))
                throw new FileNotFoundException("Excel 文件不存在", excelPath);

            var result = new List<ExcelCalcSection>();

            IWorkbook workbook;
            using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fs);
            }

            ISheet ws = workbook.GetSheetAt(0);
            if (ws == null || ws.LastRowNum < 0)
                return result;

            int headerRow = 1;          // Excel 第1行为标题（1‑based）
            int rowStart = 2;          // 数据从第2行开始（1‑based）
            int rowEnd = ws.LastRowNum + 1;  // 最大行号（1‑based）

            // FindExcelColumn 返回 1‑based 列索引（前面已修改为接受 ISheet）
            int colDevice = FindExcelColumn(ws, headerRow, "设备", 1);
            int colParamName = FindExcelColumn(ws, headerRow, "参数名称", 2);
            int colValue = FindExcelColumn(ws, headerRow, "参数", 3);
            int colUnit = FindExcelColumn(ws, headerRow, "单位", 4);
            int colValueType = FindExcelColumn(ws, headerRow, "值类型", 5);
            int colFormula = FindExcelColumn(ws, headerRow, "公式", 6);

            string currentDevice = string.Empty;
            ExcelCalcSection? currentSection = null;
            int dynamicIndex = 1;

            DataFormatter formatter = new DataFormatter(); // 用于获取与 Excel 显示一致的文本

            for (int r = rowStart; r <= rowEnd; r++)
            {
                // 转换为 0‑based 行号
                IRow row = ws.GetRow(r - 1);
                if (row == null) continue;

                // 获取各列单元格
                ICell cellDevice = row.GetCell(colDevice - 1);
                ICell cellParamName = row.GetCell(colParamName - 1);
                ICell cellValue = row.GetCell(colValue - 1);
                ICell cellUnit = row.GetCell(colUnit - 1);
                ICell cellValueType = row.GetCell(colValueType - 1);
                ICell cellFormula = row.GetCell(colFormula - 1);

                // 获取显示文本
                string cDevice = GetCellDisplayText(cellDevice, formatter).Trim();
                string cParamName = GetCellDisplayText(cellParamName, formatter).Trim();
                string cValue = GetCellDisplayText(cellValue, formatter).Trim();
                string cUnit = GetCellDisplayText(cellUnit, formatter).Trim();
                string cValueType = GetCellDisplayText(cellValueType, formatter).Trim();
                string cFormula = GetCellDisplayText(cellFormula, formatter).Trim();

                // EPPlus 的 .Formula 返回不带等号的公式字符串，NPOI 返回带等号，需要还原
                string cValueExcelFormula = GetCellFormulaWithoutEquals(cellValue).Trim();

                // 空白行判断
                bool isBlankRow =
                    string.IsNullOrWhiteSpace(cDevice) &&
                    string.IsNullOrWhiteSpace(cParamName) &&
                    string.IsNullOrWhiteSpace(cValue) &&
                    string.IsNullOrWhiteSpace(cUnit) &&
                    string.IsNullOrWhiteSpace(cValueType) &&
                    string.IsNullOrWhiteSpace(cFormula) &&
                    string.IsNullOrWhiteSpace(cValueExcelFormula);

                if (isBlankRow)
                {
                    if (currentSection != null && currentSection.Rows.Count > 0)
                    {
                        result.Add(currentSection);
                        currentSection = null;
                    }
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(cDevice))
                    currentDevice = cDevice;

                if (string.IsNullOrWhiteSpace(currentDevice))
                    continue;

                if (currentSection == null)
                {
                    string gridName = NormalizeCalcGridNameToken(currentDevice);
                    if (string.IsNullOrWhiteSpace(gridName) || !gridName.StartsWith("DataGrid_", StringComparison.OrdinalIgnoreCase))
                        gridName = $"DataGrid_动态_{dynamicIndex++}";

                    currentSection = new ExcelCalcSection
                    {
                        DeviceName = currentDevice,
                        GridName = gridName,
                        GroupHeader = currentDevice
                    };
                }

                if (string.IsNullOrWhiteSpace(cParamName))
                    continue;

                // 地址/序号使用 Excel 原始行号（保持原来的逻辑）
                int seq = r;               // 仍是 1‑based 的 Excel 行号
                string address = "C" + r;

                string valueType = NormalizeValueType(cValueType);

                // 公式优先：公式列 > 参数列自身公式
                string formulaRaw = !string.IsNullOrWhiteSpace(cFormula)
                    ? cFormula
                    : (!string.IsNullOrWhiteSpace(cValueExcelFormula)
                        ? "=" + cValueExcelFormula   // NPOI 返回的公式已去掉等号，这里加回
                        : string.Empty);

                string normalizedFormula = NormalizeExcelFormulaToInternal(formulaRaw, currentSection.GridName);

                if (!string.IsNullOrWhiteSpace(normalizedFormula))
                    valueType = "Calc";

                // MultiSelect 选项
                var options = ParseMultiSelectOptions(cValue, valueType);
                string valueText = cValue;

                if (string.Equals(valueType, "MultiSelect", StringComparison.OrdinalIgnoreCase))
                {
                    if (options.Count > 0 && !options.Contains(valueText, StringComparer.OrdinalIgnoreCase))
                        valueText = options[0];
                    normalizedFormula = string.Empty;
                }

                currentSection.Rows.Add(new CalcCsvTableRow
                {
                    GridName = currentSection.GridName,
                    Address = address,
                    Sequence = seq,
                    ParameterName = cParamName,
                    ValueType = valueType,
                    Formula = normalizedFormula,
                    ValueText = valueText,
                    Unit = cUnit,
                    Description = "Excel导入",
                    OptionValues = new ObservableCollection<string>(options)
                });
            }

            if (currentSection != null && currentSection.Rows.Count > 0)
                result.Add(currentSection);

            return result;
        }

        /// <summary>
        /// 获取与 Excel 显示一致的单元格文本（模拟 EPPlus 的 .Text）
        /// </summary>
        private static string GetCellDisplayText(ICell cell, DataFormatter formatter)
        {
            if (cell == null) return string.Empty;
            // DataFormatter 会按照单元格格式（日期、数字等）返回显示值
            return formatter.FormatCellValue(cell);
        }

        /// <summary>
        /// 获取公式单元格的公式字符串（不含前面的等号），与 EPPlus 的 .Formula 行为一致
        /// </summary>
        private static string GetCellFormulaWithoutEquals(ICell cell)
        {
            if (cell != null && cell.CellType == CellType.Formula)
            {
                string formula = cell.CellFormula;
                if (formula != null && formula.StartsWith("="))
                    return formula.Substring(1);
                return formula ?? string.Empty;
            }
            return string.Empty;
        }

        /// <summary>
        /// excel列查找：根据列标题查找列号，未找到时返回 fallback（默认按顺序查找，兼容旧模板）
        /// </summary>
        /// <param name="ws">Excel 工作表对象</param>
        /// <param name="headerRow">标题行号</param>
        /// <param name="headerName">列标题名称</param>
        /// <param name="fallback">未找到时的默认列号</param>
        /// <returns>列号</returns>
        private static int FindExcelColumn(ISheet ws, int headerRow, string headerName, int fallback)
        {
            // 原 EPPlus 的 headerRow 是 1-based；NPOI 行索引从 0 开始，此处转换为 0-based
            int rowIndex = headerRow - 1;
            if (rowIndex < 0) return fallback;

            IRow row = ws.GetRow(rowIndex);
            if (row == null) return fallback;

            int endCol = row.LastCellNum; // 实际存在的最大列索引+1（即列数）
            for (int c = 0; c < endCol; c++)
            {
                ICell cell = row.GetCell(c);
                // 获取单元格的文本内容（优先取字符串，否则用 ToString）
                string text = "";
                if (cell != null)
                {
                    // 尝试以字符串类型读取，若单元格不是字符串也尽量获取其显示值
                    text = cell.CellType == CellType.String ? cell.StringCellValue : cell.ToString();
                    text = text ?? "";
                }
                text = text.Trim();
                if (string.Equals(text, headerName, StringComparison.OrdinalIgnoreCase))
                    return c + 1; // 转换为 1-based 列号，与原函数返回值一致
            }
            return fallback;
        }

        /// <summary>
        /// 标准化值类型 输入，兼容用户输入的各种变体：
        /// </summary>
        /// <param name="raw">原始值类型字符串</param>
        /// <returns>标准化后的值类型字符串</returns>
        private static string NormalizeValueType(string raw)
        {
            string s = (raw ?? string.Empty).Trim();
            if (s.Equals("Calc", StringComparison.OrdinalIgnoreCase)) return "Calc";
            if (s.Equals("MultiSelect", StringComparison.OrdinalIgnoreCase) ||
                s.Equals("MultiplEselect", StringComparison.OrdinalIgnoreCase)) return "MultiSelect";
            return "Input";
        }

        /// <summary>
        /// excel 公式识别增强版：除了等号前缀，还支持常见函数、运算符、单元格引用等特征，减少误判
        /// </summary>
        /// <param name="raw">原始公式字符串</param>
        /// <returns>是否可能是 Excel 公式</returns>
        private static bool IsLikelyExcelFormula(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return false;
            string s = raw.Trim();

            if (s == "输入值" || s == "计算值" || s == "可选值")
                return false;

            if (s.StartsWith("=")) return true;
            if (Regex.IsMatch(s, @"\b(SQRT|IF|ROUND|MIN|MAX|SUM)\s*\(", RegexOptions.IgnoreCase)) return true;
            if (Regex.IsMatch(s, @"[+\-*/^()]")) return true;
            if (Regex.IsMatch(s, @"[A-Za-z]+\d+", RegexOptions.IgnoreCase)) return true;

            return false;
        }

        /// <summary>
        /// 下拉选项解析：支持斜杠分隔的多选（泵前\泵后\其它）和百分比范围（10%~100%）
        /// </summary>
        /// <param name="rawValue">原始下拉选项字符串</param>
        /// <param name="valueType">值类型</param>
        /// <returns>解析后的下拉选项列表</returns>
        private static List<string> ParseMultiSelectOptions(string rawValue, string valueType)
        {
            var list = new List<string>();
            if (!string.Equals(valueType, "MultiSelect", StringComparison.OrdinalIgnoreCase))
                return list;

            string s = (rawValue ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(s)) return list;

            // 泵前\泵后\其它
            if (s.Contains("\\"))
            {
                return s.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(x => x.Trim())
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();
            }

            // 90% -> 10%~100%
            if (Regex.IsMatch(s, @"^\d{1,3}\s*%$"))
            {
                for (int i = 10; i <= 100; i += 10)
                    list.Add(i + "%");
                return list;
            }

            list.Add(s);
            return list;
        }

        /// <summary>
        /// 将 Excel 公式转换为内部公式格式：
        /// 1) 统一补 '=' 前缀
        /// 2) 跨表引用中的“设备名/组名”转换为 DataGrid 名称
        /// 3) 清理引号与全角符号
        /// </summary>
        private string NormalizeExcelFormulaToInternal(string rawFormula, string currentGridName)
        {
            if (!IsLikelyExcelFormula(rawFormula))
                return string.Empty;

            string expr = rawFormula.Trim();

            expr = expr.Replace('，', ',')
                       .Replace('（', '(')
                       .Replace('）', ')')
                       .Replace('！', '!');

            if (!expr.StartsWith("="))
                expr = "=" + expr;

            expr = Regex.Replace(
                expr,
                @"'?(?<grid>[^'!]+)'?\s*!\s*(?<addr>[A-Za-z]+\d+)",
                m =>
                {
                    string gridToken = (m.Groups["grid"].Value ?? string.Empty).Trim();
                    string addr = (m.Groups["addr"].Value ?? string.Empty).Trim().ToUpperInvariant();
                    string normalizedGrid = NormalizeCalcGridNameToken(gridToken);
                    if (string.IsNullOrWhiteSpace(normalizedGrid))
                        normalizedGrid = currentGridName;
                    return $"{normalizedGrid}!{addr}";
                },
                RegexOptions.IgnoreCase);

            return expr;
        }

        /// <summary>
        /// 根据解析结果动态生成 GroupBox + DataGrid，并绑定数据源
        /// </summary>
        private void BuildDynamicCalcGrids(List<ExcelCalcSection> sections)
        {
            if (CalcDynamicHost == null)
                throw new InvalidOperationException("未找到动态容器 CalcDynamicHost。");

            CalcDynamicHost.Children.Clear();
            _calcCsvTableSources.Clear();

            int index = 1;
            foreach (var s in sections)
            {
                var group = new System.Windows.Controls.GroupBox
                {
                    Header = $"{index},{s.GroupHeader}",
                    Margin = new Thickness(0, 3, 0, 0),
                    Padding = new Thickness(0, 3, 0, 0),
                    FontFamily = new FontFamily("Microsoft YaHei UI")
                };

                var dg = CreateCalcCsvDataGrid(s.GridName);
                var source = new ObservableCollection<CalcCsvTableRow>(s.Rows.OrderBy(x => x.Sequence));

                _calcCsvTableSources[s.GridName] = source;
                dg.ItemsSource = source;

                group.Content = dg;
                CalcDynamicHost.Children.Add(group);

                index++;
            }
        }

        /// <summary>
        /// 导出为总 CSV（GridName 模板）
        /// </summary>
        private string ExportSectionsToMasterCsv(List<ExcelCalcSection> sections)
        {
            string outputPath = Path.Combine(GetCalcCsvExternalDirectory(), CalcCsvMasterFileName); // 

            var allRows = sections
                .SelectMany(s => s.Rows)
                .OrderBy(r => r.GridName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.Sequence)
                .ToList();

            var sb = new StringBuilder();
            sb.AppendLine("GridName,Address,Sequence,ParameterName,ValueType,Formula,Value,Unit,Description,Options");

            foreach (var r in allRows)
            {
                sb.AppendLine(string.Join(",",
                    EscapeCsv(r.GridName),
                    EscapeCsv(r.Address),
                    EscapeCsv(r.Sequence.ToString()),
                    EscapeCsv(r.ParameterName),
                    EscapeCsv(string.IsNullOrWhiteSpace(r.ValueType) ? "Input" : r.ValueType),
                    EscapeCsv(r.Formula ?? string.Empty),
                    EscapeCsv(r.ValueText ?? string.Empty),
                    EscapeCsv(r.Unit ?? string.Empty),
                    EscapeCsv(r.Description ?? string.Empty),
                    EscapeCsv(string.Join("|", r.OptionValues ?? new ObservableCollection<string>()))
                ));
            }

            File.WriteAllText(outputPath, sb.ToString(), new UTF8Encoding(true));
            return outputPath;
        }

        #endregion


        /// <summary>
        /// 将总表中的 GridName/段名 归一化到真正的 DataGrid 名称
        /// 支持：
        /// 1) 直接写 DataGrid_XXX
        /// 2) 写 GroupBox 名（如“循环泵选型”）
        /// 3) 写带序号或括号的段名（如“1,脱硫系统计算(单台)”）
        /// </summary>
        private string NormalizeCalcGridNameToken(string rawToken)
        {
            string token = (rawToken ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(token))
                return string.Empty;

            // 去掉前缀序号，如 "1,xxx" / "1.xxx" / "1、xxx"
            token = Regex.Replace(token, @"^\s*\d+\s*[,，.、]\s*", string.Empty);

            // 去掉括号内容，如 "(单台)" / "（单台）"
            token = Regex.Replace(token, @"[\(（].*?[\)）]", string.Empty).Trim();

            // 1) 已经是 DataGrid 名称
            if (_calcCsvTableDefinitions.Values.Any(d =>
                string.Equals(d.DataGridName, token, StringComparison.OrdinalIgnoreCase)))
            {
                return _calcCsvTableDefinitions.Values
                    .First(d => string.Equals(d.DataGridName, token, StringComparison.OrdinalIgnoreCase))
                    .DataGridName;
            }

            // 2) GroupBox 名匹配
            var byGroup = _calcCsvTableDefinitions.Values.FirstOrDefault(d =>
                string.Equals(d.GroupBoxName, token, StringComparison.OrdinalIgnoreCase));
            if (byGroup != null)
                return byGroup.DataGridName;

            // 3) 包含匹配（更宽松）
            var fuzzy = _calcCsvTableDefinitions.Values.FirstOrDefault(d =>
                d.GroupBoxName.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0 ||
                token.IndexOf(d.GroupBoxName, StringComparison.OrdinalIgnoreCase) >= 0);
            if (fuzzy != null)
                return fuzzy.DataGridName;

            return rawToken.Trim();
        }

        // ====== 放在 CalcCsv 相关字段区域 ======
        private const string CalcCsvMasterFileName = "_00_计算总表_公式导出.csv";

        /// <summary>
        /// 构建用于存储的地址：GridName|Address
        /// </summary>
        /// <param name="gridName"></param>
        /// <param name="address"></param>
        /// <returns></returns>
        private static string BuildScopedAddress(string gridName, string address)
        {
            return $"{(gridName ?? string.Empty).Trim()}|{(address ?? string.Empty).Trim().ToUpperInvariant()}";
        }

        /// <summary>
        /// 重载所有 CSV 表（优先总表）
        /// </summary>
        private void ReloadCalcCsvTables(bool showMessage)
        {
            if (_calcCsvIsReloading)
                return;

            _calcCsvIsReloading = true;

            try
            {
                string masterPath = ResolveCalcCsvPath(CalcCsvMasterFileName); // 

                if (File.Exists(masterPath))
                {
                    ReloadFromMasterCalcCsv(masterPath);
                    LogManager.Instance.LogInfo($"已按总表模式加载计算数据：{masterPath}");
                }
                else
                {
                    foreach (var definition in _calcCsvTableDefinitions.Values)
                    {
                        ReloadSingleCalcCsvTable(definition);
                    }
                    LogManager.Instance.LogInfo("未找到总表，已按旧模式逐文件加载。");
                }

                // 统一做一次全局重算（支持跨表联动）
                RecalculateAllCalcCsvTables();

                if (showMessage)
                {
                    MessageBox.Show(
                        $"CSV 已重载。\n外部附件目录：{GetCalcCsvExternalDirectory()}",
                        "提示",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            finally
            {
                _calcCsvIsReloading = false;
            }
        }

        /// <summary>
        /// 从总表 CSV 重新加载计算数据（总表模式）
        /// </summary>
        /// <param name="masterPath"></param>
        private void ReloadFromMasterCalcCsv(string masterPath)
        {
            var allRows = LoadCalcCsvRows(masterPath, string.Empty);

            var grouped = allRows
                .Where(r => !string.IsNullOrWhiteSpace(r.GridName))
                .GroupBy(r => r.GridName, StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => g.Min(x => x.Sequence))
                .ToList();

            var sections = grouped.Select(g => new ExcelCalcSection
            {
                DeviceName = "计算数据表",
                GridName = g.Key,
                GroupHeader = GetCalcGroupHeaderByGridName(g.Key),
                Rows = g.OrderBy(x => x.Sequence).ToList()
            }).ToList();

            // 关键：总表模式使用动态UI，不再走旧的 GroupBox 查找逻辑
            BuildDynamicCalcGrids(sections);
        }

        /// <summary>
        /// 根据 GridName 获取 GroupBox Header（用于总表模式）
        /// </summary>
        /// <param name="gridName"> Grid名</param>
        /// <returns></returns>
        private static string GetCalcGroupHeaderByGridName(string gridName)
        {
            string name = (gridName ?? string.Empty).Trim();
            if (name.StartsWith("DataGrid_", StringComparison.OrdinalIgnoreCase))
                name = name.Substring("DataGrid_".Length);
            else if (name.StartsWith("DataGrid", StringComparison.OrdinalIgnoreCase))
                name = name.Substring("DataGrid".Length).TrimStart('_');

            return string.IsNullOrWhiteSpace(name) ? gridName : name.Replace("_", " ");
        }

        /// <summary>
        /// 重载单个 CSV 表（旧模式）
        /// </summary>
        private void ReloadSingleCalcCsvTable(CalcCsvTableDefinition definition)
        {
            string csvPath = ResolveCalcCsvPath(definition.FileName);
            LogManager.Instance.LogInfo($"重载 CSV: {definition.FileName} -> {csvPath}");

            if (!File.Exists(csvPath))
            {
                LogManager.Instance.LogInfo($"未找到 CSV 文件: {csvPath}");
                return;
            }

            var dataGrid = EnsureCalcCsvDataGrid(definition);
            var rows = LoadCalcCsvRows(csvPath, definition.DataGridName);

            if (!_calcCsvTableSources.ContainsKey(definition.DataGridName))
            {
                _calcCsvTableSources[definition.DataGridName] = new ObservableCollection<CalcCsvTableRow>();
            }

            _calcCsvTableSources[definition.DataGridName].Clear();
            foreach (var row in rows.OrderBy(r => r.Sequence))
            {
                _calcCsvTableSources[definition.DataGridName].Add(row);
            }

            dataGrid.ItemsSource = _calcCsvTableSources[definition.DataGridName];
            // 注意：这里不再单表重算，统一由 ReloadCalcCsvTables 末尾全局重算
        }

        /// <summary>
        /// 读取新模板格式 CSV
        /// targetGridName 为空时：读取全部 GridName（总表模式）
        /// targetGridName 非空时：只读取目标 GridName（旧模式）
        /// </summary>
        private List<CalcCsvTableRow> LoadCalcCsvRowsFromTemplate(string[] lines, string targetGridName)
        {
            var rows = new List<CalcCsvTableRow>();
            bool hasTarget = !string.IsNullOrWhiteSpace(targetGridName);

            for (int i = 1; i < lines.Length; i++)
            {
                string line = lines[i];
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var cols = SplitCsvLine(line);
                if (cols.Count < 9)
                    continue;

                string gridName = NormalizeCalcGridNameToken(cols[0].Trim());
                string addressRaw = cols[1].Trim();
                string address = NormalizeCalcCellAddressForStorage(addressRaw); // 关键
                string sequenceText = cols[2].Trim();
                string parameterName = cols[3].Trim();
                string valueType = cols[4].Trim();
                string formula = cols[5].Trim();
                string value = cols[6].Trim();
                string unit = cols[7].Trim();
                string description = cols[8].Trim();
                string options = cols.Count > 9 ? cols[9].Trim() : string.Empty;

                if (hasTarget)
                {
                    if (!string.IsNullOrWhiteSpace(gridName) &&
                        !string.Equals(gridName, targetGridName, StringComparison.OrdinalIgnoreCase))
                        continue;
                    gridName = targetGridName;
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(gridName))
                        continue;
                }

                if (string.IsNullOrWhiteSpace(address) || string.IsNullOrWhiteSpace(parameterName))
                    continue;

                int sequence = 0;
                int.TryParse(sequenceText, out sequence);
                if (sequence <= 0)
                    sequence = GetRowNumberFromAddress(address);

                var optionValues = new ObservableCollection<string>(
                    (options ?? string.Empty)
                    .Split(new[] { '|', ';', '\\' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(x => x.Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x)));

                // ===== 新增：针对“纯度 C64”补齐百分比下拉并默认 90% =====
                if (string.Equals(valueType, "MultiSelect", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(gridName, "DataGrid_石灰石浆液泵选型", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(address, "C64", StringComparison.OrdinalIgnoreCase))
                {
                    // 固定下拉：10%~100%
                    optionValues = new ObservableCollection<string>(
                        Enumerable.Range(1, 10).Select(i => (i * 10).ToString() + "%"));

                    // 默认值固定为 90%
                    value = "90%";
                }

                rows.Add(new CalcCsvTableRow
                {
                    GridName = gridName,
                    Address = address, // 用规范化地址
                    Sequence = sequence,
                    ParameterName = parameterName,
                    ValueType = string.IsNullOrWhiteSpace(valueType) ? "Input" : valueType,
                    Formula = formula,
                    ValueText = value,
                    Unit = unit,
                    Description = description,
                    OptionValues = optionValues
                });
            }

            return rows;
        }

        /// <summary>
        /// 编辑结束后执行全局联动重算
        /// </summary>
        private void DataGrid_计算表自动联动重算_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (!(e.Row.Item is CalcCsvTableRow row))
                return;

            if (!row.IsEditable)
                return;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                // 1) 下拉类型不做数字校验，直接重算
                if (row.IsMultiSelect)
                {
                    RecalculateAllCalcCsvTables();
                    return;
                }

                // 2) TemplateColumn 场景下，以绑定后的 ValueText 为准
                string committedText = (row.ValueText ?? string.Empty).Trim();
                string oldText = (_calcCsvEditingOldValue ?? string.Empty).Trim();

                // 3) 未改值：不提示，直接退出（避免“没输入也报错”）
                if (string.Equals(committedText, oldText, StringComparison.Ordinal))
                    return;

                // 4) 允许空输入回退旧值，不弹错误
                if (string.IsNullOrWhiteSpace(committedText))
                {
                    row.ValueText = _calcCsvEditingOldValue;
                    return;
                }

                // 5) Input 类型必须是数字
                if (!TryParseCalcCsvDouble(committedText, out double parsedValue))
                {
                    row.ValueText = _calcCsvEditingOldValue;
                    MessageBox.Show(
                        $"参数“{row.ParameterName}”请输入有效数字。",
                        "提示",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                row.ValueText = FormatCalcCsvNumber(parsedValue);
                try
                {
                    RecalculateAllCalcCsvTables();
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogWarning($"联动重算失败: {ex.Message}");
                    MessageBox.Show($"联动重算失败：{ex.Message}", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }), System.Windows.Threading.DispatcherPriority.Background);
        }
        
        /// <summary>
        /// 
        /// </summary>
        private readonly Dictionary<string, double> _calcCsvLastKnownInputValues =
        new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);


        /// <summary>
        /// 全局重算：支持跨表联动
        /// 跨表引用格式：DataGrid_脱硫系统!C12
        /// 表内引用格式：C12（默认当前 Grid）
        /// </summary>
        private void RecalculateAllCalcCsvTables()
        {
            var allRows = _calcCsvTableSources.Values.SelectMany(v => v).ToList();
            if (allRows.Count == 0)
                return;

            var cellValues = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in allRows)
            {
                if (string.IsNullOrWhiteSpace(row.Address) || string.IsNullOrWhiteSpace(row.GridName))
                    continue;

                string normalizedAddress = NormalizeCalcCellAddressForStorage(row.Address);
                string key = BuildScopedAddress(row.GridName, normalizedAddress);

                // 1) 先按当前值解析
                if (TryParseCalcCsvDouble(row.ValueText, out double parsed))
                {
                    cellValues[key] = parsed;
                    if (row.IsEditable)
                        _calcCsvLastKnownInputValues[key] = parsed;
                    continue;
                }

                // 2) MultiSelect 兜底：若当前值不可解析，尝试候选项首个可解析值（例如 90.0%）
                if (row.IsEditable && row.IsMultiSelect && row.OptionValues != null)
                {
                    var fallbackText = row.OptionValues.FirstOrDefault(v => TryParseCalcCsvDouble(v, out _));
                    if (!string.IsNullOrWhiteSpace(fallbackText) && TryParseCalcCsvDouble(fallbackText, out double parsedFallback))
                    {
                        row.ValueText = fallbackText; // 回写，保证界面与计算一致
                        cellValues[key] = parsedFallback;
                        _calcCsvLastKnownInputValues[key] = parsedFallback;
                        continue;
                    }
                }

                // 3) 可编辑项继续使用“最后一次有效值”
                if (row.IsEditable && _calcCsvLastKnownInputValues.TryGetValue(key, out double last))
                {
                    cellValues[key] = last;
                }
            }

            // 下面其余代码保持原样...
            var calcRows = allRows
                .Where(r => !r.IsEditable && !string.IsNullOrWhiteSpace(r.Formula))
                .OrderBy(r => r.Sequence)
                .ToList();

            var pendingRows = new List<CalcCsvTableRow>(calcRows);
            Exception? lastException = null;

            for (int pass = 0; pass < calcRows.Count * 2; pass++)
            {
                if (pendingRows.Count == 0)
                    return;

                bool anyProgress = false;
                var nextPendingRows = new List<CalcCsvTableRow>();

                foreach (var row in pendingRows)
                {
                    try
                    {
                        double value = EvaluateCalcCsvFormula(row.Formula, row.GridName, cellValues);

                        if (double.IsNaN(value) || double.IsInfinity(value))
                            throw new InvalidOperationException("结果不是有效数字。");

                        row.ValueText = FormatCalcCsvNumber(value);

                        string normalizedAddress = NormalizeCalcCellAddressForStorage(row.Address);
                        cellValues[BuildScopedAddress(row.GridName, normalizedAddress)] = value;

                        anyProgress = true;
                    }
                    catch (Exception ex)
                    {
                        lastException = ex;
                        nextPendingRows.Add(row);
                    }
                }

                if (!anyProgress)
                {
                    var details = nextPendingRows
                        .Take(8)
                        .Select(r => $"{r.GridName}!{r.Address} {r.ParameterName} -> {r.Formula}");
                    throw new InvalidOperationException("计算表存在未解依赖或非法公式：\n" + string.Join("\n", details), lastException);
                }

                pendingRows = nextPendingRows;
            }

            if (pendingRows.Count > 0)
            {
                var failed = pendingRows.First();
                throw new InvalidOperationException(
                    $"计算参数“{failed.ParameterName}”失败，公式：{failed.Formula}",
                    lastException);
            }
        }

        /// <summary>
        /// 计算公式（支持跨表）
        /// 跨表：GridName!C7
        /// 本表：C7
        /// </summary>
        private double EvaluateCalcCsvFormula(string formula, string currentGridName, IDictionary<string, double> cellValues)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return 0.0;

            string expression = NormalizeFormulaForValidation(formula.Trim());

            // 兼容 "V=(...)" / "-V=(...)" / "xxx=..."
            if (!expression.StartsWith("=") && expression.Contains("="))
            {
                int eq = expression.LastIndexOf('=');
                if (eq >= 0 && eq < expression.Length - 1)
                    expression = expression.Substring(eq + 1);
            }

            if (expression.StartsWith("="))
                expression = expression.Substring(1);

            // 跨表引用
            expression = Regex.Replace(
                expression,
                @"(?<grid>[^\s\+\-\*/\(\),!]+)\s*!\s*(?<addr>[A-Z\|IL]*C\d+|[A-Z]+\d+)",
                m =>
                {
                    string gridRaw = (m.Groups["grid"].Value ?? string.Empty).Trim().Trim('\'', '"', '“', '”', '‘', '’');
                    string gridToken = gridRaw.Replace(" ", "_");

                    if (gridToken.StartsWith("DataGrid", StringComparison.OrdinalIgnoreCase) &&
                        !gridToken.StartsWith("DataGrid_", StringComparison.OrdinalIgnoreCase))
                    {
                        gridToken = "DataGrid_" + gridToken.Substring("DataGrid".Length).TrimStart('_');
                    }

                    string grid = NormalizeCalcGridNameToken(gridToken);
                    if (string.IsNullOrWhiteSpace(grid)) grid = gridToken;

                    if (!TryNormalizeCellAddressToken(m.Groups["addr"].Value, out string addr))
                        throw new InvalidOperationException($"公式地址无法识别：{gridRaw}!{m.Groups["addr"].Value}");

                    string key = BuildScopedAddress(grid, addr);
                    if (!cellValues.TryGetValue(key, out double v))
                        throw new InvalidOperationException($"公式引用未计算完成：{grid}!{addr}");

                    return ToCalcCsvDoubleLiteral(v);
                },
                RegexOptions.IgnoreCase);

            // 本地地址
            expression = Regex.Replace(
                expression,
                @"\b([A-Z\|IL]*C\d+|[A-Z]+\d+)\b",
                m =>
                {
                    if (!TryNormalizeCellAddressToken(m.Value, out string addr))
                        return m.Value;

                    if (!TryResolveScopedAddressKey(currentGridName, addr, cellValues, out string resolvedKey))
                        throw new InvalidOperationException($"公式引用未计算完成：{currentGridName}!{addr}");

                    return ToCalcCsvDoubleLiteral(cellValues[resolvedKey]);
                },
                RegexOptions.IgnoreCase);

            expression = NormalizeCalcCsvExpression(expression);
            expression = ResolveCalcCsvExcelFunctions(expression); // 新增：Excel函数
            expression = ResolveCalcCsvSqrt(expression);

            return EvaluatePlainExpression(expression);
        }
        /// <summary>
        /// 解析计算 Csv Excel函数
        /// </summary>
        /// <param name="expression">要解析的表达式</param>
        /// <returns>解析后的表达式</returns>
        /// <exception cref="InvalidOperationException"></exception>
        private string ResolveCalcCsvExcelFunctions(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
                return "0.0";

            while (true)
            {
                int fnIndex = expression.IndexOf("CEILING(", StringComparison.OrdinalIgnoreCase);
                if (fnIndex < 0) break;

                int openIndex = fnIndex + "CEILING".Length;
                int closeIndex = FindCalcCsvMatchingParenthesis(expression, openIndex);

                string inner = expression.Substring(openIndex + 1, closeIndex - openIndex - 1);
                var args = SplitCalcCsvFunctionArgs(inner);
                if (args.Count == 0)
                    throw new InvalidOperationException("CEILING 参数为空。");

                double number = EvaluatePlainExpression(NormalizeCalcCsvExpression(args[0]));
                double significance = 1.0;
                if (args.Count > 1)
                    significance = EvaluatePlainExpression(NormalizeCalcCsvExpression(args[1]));

                if (Math.Abs(significance) < 1e-12)
                    throw new InvalidOperationException("CEILING 的 significance 不能为 0。");

                double result = Math.Ceiling(number / significance) * significance;
                string replacement = ToCalcCsvDoubleLiteral(result);

                expression = expression.Substring(0, fnIndex) + replacement + expression.Substring(closeIndex + 1);
            }

            return expression;
        }
        /// <summary>
        ///  拆分计算CSV函数参数
        /// </summary>
        /// <param name="text">函数参数字符串</param>
        /// <returns>参数列表</returns>
        private static List<string> SplitCalcCsvFunctionArgs(string text)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(text)) return list;

            int depth = 0;
            var sb = new StringBuilder();

            foreach (char ch in text)
            {
                if (ch == '(') { depth++; sb.Append(ch); continue; }
                if (ch == ')') { depth--; sb.Append(ch); continue; }

                if (ch == ',' && depth == 0)
                {
                    list.Add(sb.ToString().Trim());
                    sb.Clear();
                    continue;
                }

                sb.Append(ch);
            }

            if (sb.Length > 0) list.Add(sb.ToString().Trim());
            return list;
        }
        // 新增辅助方法：放在 EvaluateCalcCsvFormula 附近
        private bool TryResolveScopedAddressKey(string currentGridName, string address, IDictionary<string, double> cellValues, out string resolvedKey)
        {
            resolvedKey = BuildScopedAddress(currentGridName, address);

            // 1) 先找当前表
            if (cellValues.ContainsKey(resolvedKey))
                return true;

            // 2) 当前表找不到，尝试全局唯一地址（同一 Excel 按行号拆段常见）
            string suffix = "|" + address.ToUpperInvariant();
            var matches = cellValues.Keys
                .Where(k => k.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (matches.Count == 1)
            {
                resolvedKey = matches[0];
                return true;
            }

            return false;
        }
        #endregion
        

        #region CSV驱动计算表

        /// <summary>
        /// 单个 CSV 文件与界面区域的映射定义
        /// </summary>
        private sealed class CalcCsvTableDefinition
        {
            /// <summary>
            /// CSV 文件名
            /// </summary>
            public string FileName { get; set; } = string.Empty;

            /// <summary>
            /// 目标 GroupBox 名称
            /// </summary>
            public string GroupBoxName { get; set; } = string.Empty;

            /// <summary>
            /// 目标 DataGrid 名称
            /// </summary>
            public string DataGridName { get; set; } = string.Empty;
        }

        /// <summary>
        /// 计算表行模型
        /// </summary>
        private sealed class CalcCsvTableRow : INotifyPropertyChanged
        {
            private string _valueText = string.Empty;

            /// <summary>
            /// 序号
            /// </summary>
            public int Sequence { get; set; }

            /// <summary>
            /// 参数名称
            /// </summary>
            public string ParameterName { get; set; } = string.Empty;

            /// <summary>
            /// 公式单元格地址，例如 C7
            /// </summary>
            public string Address { get; set; } = string.Empty;

            /// <summary>
            /// 对应 DataGrid 名称
            /// </summary>
            public string GridName { get; set; } = string.Empty;

            /// <summary>
            /// 值类型：Input / Calc
            /// </summary>
            public string ValueType { get; set; } = "Input";

            /// <summary>
            /// 公式
            /// </summary>
            public string Formula { get; set; } = string.Empty;

            /// <summary>
            /// 单位
            /// </summary>
            public string Unit { get; set; } = string.Empty;

            /// <summary>
            /// 说明
            /// </summary>
            public string Description { get; set; } = string.Empty;

            // 新增：下拉候选值（MultiSelect 专用）
            public ObservableCollection<string> OptionValues { get; set; } = new ObservableCollection<string>();

            public bool IsMultiSelect =>
                string.Equals(ValueType, "MultiSelect", StringComparison.OrdinalIgnoreCase);

            /// <summary>
            /// 当前显示值
            /// </summary>
            public string ValueText
            {
                get => _valueText;
                set
                {
                    if (_valueText == value) return;
                    _valueText = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ValueText)));
                }
            }

            /// <summary>
            /// 是否可编辑
            /// </summary>
            // Input / MultiSelect 可编辑；Calc 不可编辑
            public bool IsEditable =>
                string.Equals(ValueType, "Input", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(ValueType, "MultiSelect", StringComparison.OrdinalIgnoreCase);

            public event PropertyChangedEventHandler? PropertyChanged;
        }

        /// <summary>
        /// 文件名与目标表的映射表
        /// 后续新增 CSV，只需要在这里补一条即可
        /// </summary>
        private readonly Dictionary<string, CalcCsvTableDefinition> _calcCsvTableDefinitions =
            new Dictionary<string, CalcCsvTableDefinition>(StringComparer.OrdinalIgnoreCase)
            {
        {
            "_1_脱硫系统计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_1_脱硫系统计算_公式导出.csv",
                GroupBoxName = "脱硫系统计算",
                DataGridName = "DataGrid_脱硫系统"
            }
        },
        {
            "_2_循环泵选型计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_2_循环泵选型计算_公式导出.csv",
                GroupBoxName = "循环泵选型",
                DataGridName = "DataGrid_循环泵选型计算"
            }
        },
        {
            "_3_进_口管道尺寸计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_3_进_口管道尺寸计算_公式导出.csv",
                GroupBoxName = "进口管道尺寸",
                DataGridName = "DataGrid_进口管道尺寸"
            }
        },
        {
            "_4_出_口管道尺寸计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_4_出_口管道尺寸计算_公式导出.csv",
                GroupBoxName = "出口管道尺寸",
                DataGridName = "DataGrid_出口管道尺寸"
            }
        },
        {
            "_5_氧化风机选型计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_5_氧化风机选型计算_公式导出.csv",
                GroupBoxName = "氧化风机选型",
                DataGridName = "DataGrid_氧化风机选型"
            }
        },
        {
            "_6_氧化风机后出口管道尺寸计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_6_氧化风机后出口管道尺寸计算_公式导出.csv",
                GroupBoxName = "氧化风机后出口管道尺寸",
                DataGridName = "DataGrid_氧化风机后出口管道尺寸"
            }
        },
        {
            "_7_石灰石浆液泵选1远型计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_7_石灰石浆液泵选1远型计算_公式导出.csv",
                GroupBoxName = "石灰石浆液泵选型",
                DataGridName = "DataGrid_石灰石浆液泵选型"
            }
        },
        {
            "_8_石灰石浆液泵选2近型计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_8_石灰石浆液泵选2近型计算_公式导出.csv",
                GroupBoxName = "石灰石浆液泵选2近型计算",
                DataGridName = "DataGrid_石灰石浆液泵选2近型计算"
            }
        },
        {
            "_9_石膏排出泵选1远型计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_9_石膏排出泵选1远型计算_公式导出.csv",
                GroupBoxName = "石膏排出泵选1远型计算",
                DataGridName = "DataGrid_石膏排出泵选1远型计算"
            }
        },
        {
            "_10_石膏排出泵选2近型计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_10_石膏排出泵选2近型计算_公式导出.csv",
                GroupBoxName = "石膏排出泵选2近型计算",
                DataGridName = "DataGrid_石膏排出泵选2近型计算"
            }
        },
        {
            "_11_缓冲泵选型计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_11_缓冲泵选型计算_公式导出.csv",
                GroupBoxName = "缓冲泵选型计算",
                DataGridName = "DataGrid_缓冲泵选型计算"
            }
        },
        {
            "_12_滤液水泵选1远型计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_12_滤液水泵选1远型计算_公式导出.csv",
                GroupBoxName = "滤液水泵选1远型计算",
                DataGridName = "DataGrid_滤液水泵选1远型计算"
            }
        },
        {
            "_13_滤液水泵选2近型计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_13_滤液水泵选2近型计算_公式导出.csv",
                GroupBoxName = "滤液水泵选2近型计算",
                DataGridName = "DataGrid_滤液水泵选2近型计算"
            }
        },
        {
            "_14_事故泵选1远型计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_14_事故泵选1远型计算_公式导出.csv",
                GroupBoxName = "事故泵选1远型计算",
                DataGridName = "DataGrid_事故泵选1远型计算"
            }
        },
        {
            "_15_事故泵选2近型计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_15_事故泵选2近型计算_公式导出.csv",
                GroupBoxName = "事故泵选2近型计算",
                DataGridName = "DataGrid_事故泵选2近型计算"
            }
        },
        {
            "_16_地坑泵选型计算_公式导出.csv",
            new CalcCsvTableDefinition
            {
                FileName = "_16_地坑泵选型计算_公式导出.csv",
                GroupBoxName = "地坑泵选型计算",
                DataGridName = "DataGrid_地坑泵选型计算"
            }
        }
            };



        /// <summary>
        /// 每个 DataGrid 对应一份数据源
        /// Key = DataGrid 名称
        /// </summary>
        private readonly Dictionary<string, ObservableCollection<CalcCsvTableRow>> _calcCsvTableSources =
            new Dictionary<string, ObservableCollection<CalcCsvTableRow>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 编辑前的旧值，用于非法输入回退
        /// </summary>
        private string _calcCsvEditingOldValue = string.Empty;

        /// <summary>
        /// 防止重算时递归触发
        /// </summary>
        private bool _calcCsvIsReloading = false;

        /// <summary>
        /// 初始化所有 CSV 驱动表
        /// </summary>
        private void InitializeCalcCsvTables()
        {
            try
            {
                ReloadCalcCsvTables(false); // 重载所有 CSV 表，false 表示不弹提示
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"初始化 CSV 计算表失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取外部 CSV 附件目录
        /// 这里放在本地用户目录，便于线下替换
        /// </summary>
        private string GetCalcCsvExternalDirectory()
        {
            string directory = System.IO.Path.Combine(AppPath, "CalcCsvTemplates");
            if (!System.IO.Directory.Exists(directory))
            {
                System.IO.Directory.CreateDirectory(directory);
            }
            return directory;
        }

        /// <summary>
        /// 获取最终加载的 CSV 文件路径
        /// 优先级：外部附件目录 > 程序 Resources 目录 > 开发目录 Resources
        /// </summary>
        private string ResolveCalcCsvPath(string fileName)
        {
            string externalPath = System.IO.Path.Combine(GetCalcCsvExternalDirectory(), fileName);

            // 1) 本地已有，直接用本地
            if (System.IO.File.Exists(externalPath))
                return externalPath;

            // 2) 运行目录 Resources 作为种子
            string runtimeResourcePath = System.IO.Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Resources",
                fileName);

            if (TrySeedCalcCsvToExternal(runtimeResourcePath, externalPath, out string resolvedFromRuntime))
                return resolvedFromRuntime;

            // 3) 开发目录 Resources 作为种子（调试兜底）
            string devResourcePath = System.IO.Path.Combine(
                @"C:\Users\ShiGu\source\repos\GB_NewCadPlus_IV\Resources",
                fileName);

            if (TrySeedCalcCsvToExternal(devResourcePath, externalPath, out string resolvedFromDev))
                return resolvedFromDev;

            // 4) 最后尝试从 Resources.resx 导出到本地外部目录
            if (TryExportCalcCsvFromResx(fileName, externalPath, out string exportedPath))
                return exportedPath;

            // 找不到时返回本地目标路径，便于日志提示
            return externalPath;
        }
        /// <summary>
        /// 用种子文件初始化本地外部 CSV。
        /// 成功时优先返回本地 externalPath；若复制失败则回退返回 seedPath（保证仍可读取）
        /// </summary>
        private bool TrySeedCalcCsvToExternal(string seedPath, string externalPath, out string resolvedPath)
        {
            resolvedPath = string.Empty;

            if (string.IsNullOrWhiteSpace(seedPath) || !System.IO.File.Exists(seedPath))
                return false;

            try
            {
                string? directory = System.IO.Path.GetDirectoryName(externalPath);
                if (!string.IsNullOrWhiteSpace(directory) && !System.IO.Directory.Exists(directory))
                {
                    System.IO.Directory.CreateDirectory(directory);
                }

                if (!System.IO.File.Exists(externalPath))
                {
                    System.IO.File.Copy(seedPath, externalPath, false);
                    LogManager.Instance.LogInfo($"已用种子公式文件初始化本地模板：{externalPath}");
                }

                resolvedPath = externalPath;
                return true;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogWarning($"种子公式文件复制到本地失败，改为直接加载种子文件：{ex.Message}");
                resolvedPath = seedPath;
                return true;
            }
        }
        /// <summary>
        /// 尝试从 Resources.resx 中导出 CSV 到外部附件目录
        /// </summary>
        private bool TryExportCalcCsvFromResx(string fileName, string targetPath, out string exportedPath)
        {
            exportedPath = string.Empty;

            foreach (string resourceKey in GetCalcCsvResourceCandidateKeys(fileName))
            {
                try
                {
                    object? resourceObject = global::GB_NewCadPlus_IV.Resources.ResourceManager.GetObject(
                        resourceKey,
                        global::GB_NewCadPlus_IV.Resources.Culture);

                    if (!(resourceObject is byte[] bytes) || bytes.Length == 0)
                        continue;

                    string? directory = System.IO.Path.GetDirectoryName(targetPath);
                    if (!string.IsNullOrWhiteSpace(directory) && !System.IO.Directory.Exists(directory))
                    {
                        System.IO.Directory.CreateDirectory(directory);
                    }

                    System.IO.File.WriteAllBytes(targetPath, bytes);
                    exportedPath = targetPath;

                    LogManager.Instance.LogInfo(
                        $"已从 Resources.resx 导出 CSV：资源键={resourceKey}，目标文件={targetPath}");

                    return true;
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogWarning(
                        $"从 Resources.resx 导出 CSV 失败：资源键={resourceKey}，错误={ex.Message}");
                }
            }

            return false;
        }

        /// <summary>
        /// 根据 CSV 文件名生成可能的 Resources.resx 资源键
        /// 兼容：
        /// 1. 文件名去掉 .csv 后直接匹配
        /// 2. 远型 / 近型 在 resx 中被转成 _远_型 / _近_型 的情况
        /// 3. 中英文括号被转成下划线的情况
        /// </summary>
        private IEnumerable<string> GetCalcCsvResourceCandidateKeys(string fileName)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AddKey(string? key)
            {
                if (!string.IsNullOrWhiteSpace(key))
                    keys.Add(key.Trim());
            }

            string baseName = System.IO.Path.GetFileNameWithoutExtension(fileName).Trim();

            AddKey(baseName);
            AddKey(baseName.Replace("远型", "_远_型").Replace("近型", "_近_型"));
            AddKey(baseName.Replace("(远)", "_远_").Replace("(近)", "_近_"));
            AddKey(baseName.Replace("（远）", "_远_").Replace("（近）", "_近_"));
            AddKey(
                baseName
                    .Replace("(远)", "_远_")
                    .Replace("(近)", "_近_")
                    .Replace("（远）", "_远_")
                    .Replace("（近）", "_近_")
                    .Replace("远型", "_远_型")
                    .Replace("近型", "_近_型"));

            return keys;
        }

        /// <summary>
        /// 重载所有 CSV 表
        /// </summary>
        //private void ReloadCalcCsvTables(bool showMessage)
        //{
        //    if (_calcCsvIsReloading)
        //        return;

        //    _calcCsvIsReloading = true;

        //    try
        //    {
        //        foreach (var definition in _calcCsvTableDefinitions.Values)
        //        {
        //            ReloadSingleCalcCsvTable(definition);
        //        }

        //        if (showMessage)
        //        {
        //            MessageBox.Show(
        //                $"CSV 已重载。\n外部附件目录：{GetCalcCsvExternalDirectory()}",
        //                "提示",
        //                MessageBoxButton.OK,
        //                MessageBoxImage.Information);
        //        }
        //    }
        //    finally
        //    {
        //        _calcCsvIsReloading = false;
        //    }
        //}

        /// <summary>
        /// 重载单个 CSV 表
        /// </summary>
        //private void ReloadSingleCalcCsvTable(CalcCsvTableDefinition definition)
        //{
        //    string csvPath = ResolveCalcCsvPath(definition.FileName);
        //    LogManager.Instance.LogInfo($"重载 CSV: {definition.FileName} -> {csvPath}");

        //    if (!System.IO.File.Exists(csvPath))
        //    {
        //        LogManager.Instance.LogInfo($"未找到 CSV 文件: {csvPath}");
        //        return;
        //    }

        //    var dataGrid = EnsureCalcCsvDataGrid(definition);
        //    var rows = LoadCalcCsvRows(csvPath, definition.DataGridName);

        //    if (!_calcCsvTableSources.ContainsKey(definition.DataGridName))
        //    {
        //        _calcCsvTableSources[definition.DataGridName] = new ObservableCollection<CalcCsvTableRow>();
        //    }

        //    _calcCsvTableSources[definition.DataGridName].Clear();
        //    foreach (var row in rows.OrderBy(r => r.Sequence))
        //    {
        //        _calcCsvTableSources[definition.DataGridName].Add(row);
        //    }

        //    dataGrid.ItemsSource = _calcCsvTableSources[definition.DataGridName];
        //    RecalculateCalcCsvTable(definition.DataGridName);
        //}

        /// <summary>
        /// 确保目标 GroupBox 中存在 DataGrid
        /// 已有则直接用；没有则动态创建
        /// </summary>
        private DataGrid EnsureCalcCsvDataGrid(CalcCsvTableDefinition definition)
        {
            // 先尝试找 XAML 中已存在的 DataGrid
            var existedDataGrid = FindName(definition.DataGridName) as DataGrid;
            if (existedDataGrid != null)
                return existedDataGrid;

            var groupBox = FindName(definition.GroupBoxName) as System.Windows.Controls.GroupBox;
            if (groupBox == null)
                throw new InvalidOperationException($"未找到 GroupBox：{definition.GroupBoxName}");

            var dataGrid = CreateCalcCsvDataGrid(definition.DataGridName);

            // 如果 GroupBox 内容本来就是一个 Grid，则清空后放入 DataGrid
            if (groupBox.Content is Grid hostGrid)
            {
                hostGrid.Children.Clear();
                hostGrid.Children.Add(dataGrid);
            }
            else
            {
                // 否则直接替换 GroupBox 内容
                groupBox.Content = dataGrid;
            }

            try
            {
                RegisterName(definition.DataGridName, dataGrid);
            }
            catch
            {
                // 运行时重复注册名称时忽略即可
            }

            return dataGrid;
        }

        /// <summary>
        /// 运行时动态创建计算表 DataGrid
        /// </summary>
        private DataGrid CreateCalcCsvDataGrid(string dataGridName)
        {
            var dataGrid = new DataGrid
            {
                Name = dataGridName,
                Margin = new Thickness(0),
                AutoGenerateColumns = false,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                GridLinesVisibility = DataGridGridLinesVisibility.All,
                HeadersVisibility = DataGridHeadersVisibility.Column,
                IsReadOnly = false
            };

            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Width = 38,
                Header = "序号",
                Binding = new Binding("Sequence"),
                IsReadOnly = true
            });

            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Width = 130,
                Header = "参数名称",
                Binding = new Binding("ParameterName"),
                IsReadOnly = true
            });

            dataGrid.Columns.Add(new DataGridTemplateColumn
            {
                Width = 60,
                Header = "值",
                CellTemplate = BuildCalcValueCellTemplate(),
                CellEditingTemplate = BuildCalcValueEditingTemplate()
            });

            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Width = 45,
                Header = "单位",
                Binding = new Binding("Unit"),
                IsReadOnly = true
            });

            var rowStyle = new Style(typeof(DataGridRow));
            var trigger = new DataTrigger
            {
                Binding = new Binding("IsEditable"),
                Value = false
            };
            trigger.Setters.Add(new Setter(
                DataGridRow.BackgroundProperty,
                new SolidColorBrush(global::System.Windows.Media.Color.FromRgb(244, 244, 244))));
            trigger.Setters.Add(new Setter(DataGridRow.ToolTipProperty, new Binding("Formula")));
            rowStyle.Triggers.Add(trigger);
            dataGrid.RowStyle = rowStyle;

            dataGrid.BeginningEdit += DataGrid_计算表输入项_BeginningEdit;
            dataGrid.CellEditEnding += DataGrid_计算表自动联动重算_CellEditEnding;

            return dataGrid;
        }
        /// <summary>
        ///  构建计算值单元格模板
        /// </summary>
        /// <returns> </returns>
        private DataTemplate BuildCalcValueCellTemplate()
        {
            var template = new DataTemplate();
            var grid = new FrameworkElementFactory(typeof(Grid));

            var txt = new FrameworkElementFactory(typeof(TextBlock));
            txt.SetBinding(TextBlock.TextProperty, new Binding("ValueText"));
            txt.SetValue(TextBlock.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);

            var txtStyle = new Style(typeof(TextBlock));
            txtStyle.Setters.Add(new Setter(UIElement.VisibilityProperty, global::System.Windows.Visibility.Visible));
            txtStyle.Triggers.Add(new DataTrigger
            {
                Binding = new Binding("IsMultiSelect"),
                Value = true,
                Setters = { new Setter(UIElement.VisibilityProperty, global::System.Windows.Visibility.Collapsed) }
            });
            txt.SetValue(TextBlock.StyleProperty, txtStyle);

            var combo = new FrameworkElementFactory(typeof(ComboBox));
            combo.SetBinding(ComboBox.ItemsSourceProperty, new Binding("OptionValues"));
            combo.SetBinding(ComboBox.SelectedItemProperty, new Binding("ValueText")
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            });
            combo.SetValue(ComboBox.IsEditableProperty, false);
            combo.SetValue(FrameworkElement.MarginProperty, new Thickness(0));

            var comboStyle = new Style(typeof(ComboBox));
            comboStyle.Setters.Add(new Setter(UIElement.VisibilityProperty, global::System.Windows.Visibility.Collapsed));
            comboStyle.Triggers.Add(new DataTrigger
            {
                Binding = new Binding("IsMultiSelect"),
                Value = true,
                Setters = { new Setter(UIElement.VisibilityProperty, global::System.Windows.Visibility.Visible) }
            });
            combo.SetValue(ComboBox.StyleProperty, comboStyle);

            grid.AppendChild(txt);
            grid.AppendChild(combo);
            template.VisualTree = grid;
            return template;
        }
        /// <summary>
        /// 构建计算值编辑模板
        /// </summary>
        /// <returns></returns>
        private DataTemplate BuildCalcValueEditingTemplate()
        {
            var template = new DataTemplate();
            var grid = new FrameworkElementFactory(typeof(Grid));

            var tb = new FrameworkElementFactory(typeof(TextBox));
            tb.SetBinding(TextBox.TextProperty, new Binding("ValueText")
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            });

            var tbStyle = new Style(typeof(TextBox));
            tbStyle.Setters.Add(new Setter(UIElement.VisibilityProperty, global::System.Windows.Visibility.Visible));
            tbStyle.Triggers.Add(new DataTrigger
            {
                Binding = new Binding("IsMultiSelect"),
                Value = true,
                Setters = { new Setter(UIElement.VisibilityProperty, global::System.Windows.Visibility.Collapsed) }
            });
            tb.SetValue(TextBox.StyleProperty, tbStyle);

            var combo = new FrameworkElementFactory(typeof(ComboBox));
            combo.SetBinding(ComboBox.ItemsSourceProperty, new Binding("OptionValues"));
            combo.SetBinding(ComboBox.SelectedItemProperty, new Binding("ValueText")
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            });
            combo.SetValue(ComboBox.IsEditableProperty, false);

            var comboStyle = new Style(typeof(ComboBox));
            comboStyle.Setters.Add(new Setter(UIElement.VisibilityProperty, global::System.Windows.Visibility.Collapsed));
            comboStyle.Triggers.Add(new DataTrigger
            {
                Binding = new Binding("IsMultiSelect"),
                Value = true,
                Setters = { new Setter(UIElement.VisibilityProperty, global::System.Windows.Visibility.Visible) }
            });
            combo.SetValue(ComboBox.StyleProperty, comboStyle);

            grid.AppendChild(tb);
            grid.AppendChild(combo);
            template.VisualTree = grid;
            return template;
        }

        /// <summary>
        /// 读取 CSV 行
        /// 同时兼容：
        /// 1. 新模板 CSV
        /// 2. 旧版 Address/Formula/Value 结构 CSV
        /// </summary>
        private List<CalcCsvTableRow> LoadCalcCsvRows(string csvPath, string targetGridName)
        {
            string[] lines = System.IO.File.ReadAllLines(csvPath, Encoding.UTF8);
            if (lines.Length == 0)
                return new List<CalcCsvTableRow>();

            string header = (lines[0] ?? string.Empty).Trim();

            // 新模板格式
            if (header.IndexOf("GridName", StringComparison.OrdinalIgnoreCase) >= 0 &&
                header.IndexOf("ParameterName", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return LoadCalcCsvRowsFromTemplate(lines, targetGridName);
            }

            // 兼容旧格式：Address,Formula,Value
            if (header.StartsWith("Address", StringComparison.OrdinalIgnoreCase))
            {
                return LoadCalcCsvRowsFromLegacy(lines, targetGridName);
            }

            throw new InvalidOperationException($"无法识别的 CSV 格式：{csvPath}");
        }

        /// <summary>
        /// 读取新模板格式 CSV
        /// </summary>
        //private List<CalcCsvTableRow> LoadCalcCsvRowsFromTemplate(string[] lines, string targetGridName)
        //{
        //    var rows = new List<CalcCsvTableRow>();

        //    for (int i = 1; i < lines.Length; i++)
        //    {
        //        string line = lines[i];
        //        if (string.IsNullOrWhiteSpace(line))
        //            continue;

        //        var cols = SplitCsvLine(line);
        //        if (cols.Count < 9)
        //            continue;

        //        string gridName = cols[0].Trim();
        //        string address = cols[1].Trim();
        //        string sequenceText = cols[2].Trim();
        //        string parameterName = cols[3].Trim();
        //        string valueType = cols[4].Trim();
        //        string formula = cols[5].Trim();
        //        string value = cols[6].Trim();
        //        string unit = cols[7].Trim();
        //        string description = cols[8].Trim();

        //        // 文件名已决定目标表，这里 GridName 主要用于自说明
        //        // 若模板里写了别的 GridName，则跳过
        //        if (!string.IsNullOrWhiteSpace(gridName) &&
        //            !string.Equals(gridName, targetGridName, StringComparison.OrdinalIgnoreCase))
        //        {
        //            continue;
        //        }

        //        if (string.IsNullOrWhiteSpace(address) || string.IsNullOrWhiteSpace(parameterName))
        //            continue;

        //        int sequence = 0;
        //        int.TryParse(sequenceText, out sequence);
        //        if (sequence <= 0)
        //            sequence = GetRowNumberFromAddress(address);

        //        rows.Add(new CalcCsvTableRow
        //        {
        //            GridName = targetGridName,
        //            Address = address,
        //            Sequence = sequence,
        //            ParameterName = parameterName,
        //            ValueType = string.IsNullOrWhiteSpace(valueType) ? "Input" : valueType,
        //            Formula = formula,
        //            ValueText = value,
        //            Unit = unit,
        //            Description = description
        //        });
        //    }

        //    return rows;
        //}

        /// <summary>
        /// 兼容旧版 CSV：
        /// Address,Formula,Value
        /// A/B/C/D 四列拆开的结构
        /// </summary>
        private List<CalcCsvTableRow> LoadCalcCsvRowsFromLegacy(string[] lines, string targetGridName)
        {
            var cellMap = new Dictionary<string, LegacyCsvCell>(StringComparer.OrdinalIgnoreCase);

            for (int i = 1; i < lines.Length; i++)
            {
                string line = lines[i];
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var cols = SplitCsvLine(line);
                if (cols.Count < 3)
                    continue;

                string address = cols[0].Trim();
                if (string.IsNullOrWhiteSpace(address))
                    continue;

                cellMap[address] = new LegacyCsvCell
                {
                    Formula = cols.Count > 1 ? cols[1].Trim() : string.Empty,
                    Value = cols.Count > 2 ? cols[2].Trim() : string.Empty
                };
            }

            var rows = new List<CalcCsvTableRow>();

            // 旧格式里真正参与计算和值显示的核心列是 C 列
            // 因此按 C* 提取行号，而不是按 B*
            var rowNumbers = cellMap.Keys
                .Where(k => k.StartsWith("C", StringComparison.OrdinalIgnoreCase))
                .Select(GetRowNumberFromAddress)
                .Where(n => n >= 2)
                .Distinct()
                .OrderBy(n => n)
                .ToList();

            foreach (int rowNumber in rowNumbers)
            {
                string address = "C" + rowNumber;
                string formula = GetLegacyCellFormula(cellMap, address);
                string value = GetLegacyCellValue(cellMap, address);
                string unit = GetLegacyCellValue(cellMap, "D" + rowNumber);

                // 参数名优先取 B 列；若为空，再回退 A 列；还为空则用地址兜底
                string parameterName = GetLegacyCellValue(cellMap, "B" + rowNumber);
                if (string.IsNullOrWhiteSpace(parameterName))
                    parameterName = GetLegacyCellValue(cellMap, "A" + rowNumber);
                if (string.IsNullOrWhiteSpace(parameterName))
                    parameterName = address;

                // 若该行既没有值也没有公式也没有单位/标题，则跳过纯空行
                if (string.IsNullOrWhiteSpace(parameterName) &&
                    string.IsNullOrWhiteSpace(formula) &&
                    string.IsNullOrWhiteSpace(value) &&
                    string.IsNullOrWhiteSpace(unit))
                {
                    continue;
                }

                rows.Add(new CalcCsvTableRow
                {
                    GridName = targetGridName,
                    Address = address,
                    Sequence = rowNumber,
                    ParameterName = parameterName,
                    ValueType = string.IsNullOrWhiteSpace(formula) ? "Input" : "Calc",
                    Formula = formula,
                    ValueText = value,
                    Unit = unit,
                    Description = string.IsNullOrWhiteSpace(formula) ? "输入值" : "计算值"
                });
            }

            return rows;
        }

        /// <summary>
        /// 处理 SQRT(...) 函数
        /// </summary>
        private string ResolveCalcCsvSqrt(string expression)
        {
            while (true)
            {
                int sqrtIndex = expression.IndexOf("SQRT(", StringComparison.OrdinalIgnoreCase);
                if (sqrtIndex < 0)
                    return expression;

                int openIndex = sqrtIndex + 4;
                int closeIndex = FindCalcCsvMatchingParenthesis(expression, openIndex);

                string innerExpression = expression.Substring(openIndex + 1, closeIndex - openIndex - 1);
                double innerValue = EvaluatePlainExpression(innerExpression);
                double sqrtValue = Math.Sqrt(innerValue);

                string replacement = ToCalcCsvDoubleLiteral(sqrtValue);
                expression = expression.Substring(0, sqrtIndex) + replacement + expression.Substring(closeIndex + 1);
            }
        }

        /// <summary>
        /// 计算普通表达式
        /// </summary>
        private static double EvaluatePlainExpression(string expression)
        {
            string normalizedExpression = NormalizeCalcCsvExpression(expression);

            var table = new DataTable
            {
                Locale = System.Globalization.CultureInfo.InvariantCulture
            };

            object result = table.Compute(normalizedExpression, string.Empty);
            return Convert.ToDouble(result, System.Globalization.CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 把表达式中的整数字面量统一转换为双精度字面量，避免 DataTable.Compute 走 Int32 计算
        /// 例如：101325 -> 101325.0
        /// 注意：这里只处理“独立的整数”，不会破坏已有小数，也不会处理单元格引用
        /// </summary>
        private static string NormalizeCalcCsvExpression(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
                return "0.0";

            return Regex.Replace(
                expression,
                @"(?<![\w.])\d+(?![\w.])",
                m => m.Value + ".0");
        }
        /// <summary>
        /// 把 double 转成适合公式计算的文本
        /// 要求：必须带小数点，避免被 DataTable.Compute 当成 Int32
        /// </summary>
        private static string ToCalcCsvDoubleLiteral(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
                return "0.0";

            string text = value.ToString("0.###############", System.Globalization.CultureInfo.InvariantCulture);
            return text.Contains(".") ? text : text + ".0";
        }

        /// <summary>
        /// 找到匹配的右括号
        /// </summary>
        private static int FindCalcCsvMatchingParenthesis(string text, int openIndex)
        {
            int depth = 0;

            for (int i = openIndex; i < text.Length; i++)
            {
                if (text[i] == '(')
                    depth++;
                else if (text[i] == ')')
                {
                    depth--;
                    if (depth == 0)
                        return i;
                }
            }

            throw new InvalidOperationException("公式括号不匹配。");
        }

        /// <summary>
        /// 计算表编辑开始事件
        /// 只允许编辑输入项
        /// </summary>
        private void DataGrid_计算表输入项_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            if (!(e.Row.Item is CalcCsvTableRow row))
                return;

            if (!row.IsEditable)
            {
                e.Cancel = true;
                return;
            }

            _calcCsvEditingOldValue = row.ValueText;
        }

        /// <summary>
        /// 计算表编辑结束事件
        /// 输入新值后，自动联动重算
        /// </summary>
        //private void DataGrid_计算表自动联动重算_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        //{
        //    if (!(e.Row.Item is CalcCsvTableRow row))
        //        return;

        //    if (!row.IsEditable)
        //        return;

        //    var textBox = e.EditingElement as TextBox;
        //    string newText = textBox?.Text?.Trim() ?? string.Empty;
        //    string dataGridName = (sender as DataGrid)?.Name ?? row.GridName;

        //    Dispatcher.BeginInvoke(new Action(() =>
        //    {
        //        if (!TryParseCalcCsvDouble(newText, out double parsedValue))
        //        {
        //            row.ValueText = _calcCsvEditingOldValue;
        //            MessageBox.Show(
        //                $"参数“{row.ParameterName}”请输入有效数字。",
        //                "提示",
        //                MessageBoxButton.OK,
        //                MessageBoxImage.Warning);
        //            return;
        //        }

        //        row.ValueText = FormatCalcCsvNumber(parsedValue);
        //        RecalculateCalcCsvTable(dataGridName);
        //    }), System.Windows.Threading.DispatcherPriority.Background);
        //}

        

        /// <summary>
        /// 尝试解析数值
        /// </summary>
        private static bool TryParseCalcCsvDouble(string? text, out double value)
        {
            value = 0.0;
            string raw = (text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw))
                return false;

            // 统一全角符号/空格
            raw = raw.Replace('％', '%')
                     .Replace('，', ',')
                     .Replace('\u3000', ' ')
                     .Trim();

            // 百分比优先处理：90% -> 0.9
            bool isPercent = raw.EndsWith("%", StringComparison.Ordinal);
            if (isPercent)
            {
                raw = raw.Substring(0, raw.Length - 1).Trim();
            }

            // 兼容多种小数格式
            bool ok =
                double.TryParse(raw, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out value) ||
                double.TryParse(raw, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.CurrentCulture, out value) ||
                double.TryParse(raw.Replace(",", "."), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out value);

            if (!ok)
                return false;

            if (isPercent)
                value /= 100.0;

            return true;
        }

        /// <summary>
        /// 统一格式化显示
        /// </summary>
        private static string FormatCalcCsvNumber(double value)
        {
            return value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 解析 CSV 行，兼容双引号
        /// </summary>
        private static List<string> SplitCsvLine(string line)
        {
            var result = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;

            foreach (char ch in line)
            {
                if (ch == '"')
                {
                    inQuotes = !inQuotes;
                    continue;
                }

                if (ch == ',' && !inQuotes)
                {
                    result.Add(current.ToString());
                    current.Clear();
                    continue;
                }

                current.Append(ch);
            }

            result.Add(current.ToString());
            return result;
        }

        /// <summary>
        /// 从地址中提取行号，例如 C21 -> 21
        /// </summary>
        private static int GetRowNumberFromAddress(string address)
        {
            var match = Regex.Match(address ?? string.Empty, @"\d+");
            return match.Success ? int.Parse(match.Value) : 0;
        }

        /// <summary>
        /// 旧格式 CSV 单元格模型
        /// </summary>
        private sealed class LegacyCsvCell
        {
            public string Formula { get; set; } = string.Empty;
            public string Value { get; set; } = string.Empty;
        }

        /// <summary>
        /// 读取旧格式单元格值
        /// </summary>
        private static string GetLegacyCellValue(Dictionary<string, LegacyCsvCell> cellMap, string address)
        {
            return cellMap.TryGetValue(address, out var cell) ? (cell.Value ?? string.Empty).Trim() : string.Empty;
        }

        /// <summary>
        /// 读取旧格式单元格公式
        /// </summary>
        private static string GetLegacyCellFormula(Dictionary<string, LegacyCsvCell> cellMap, string address)
        {
            return cellMap.TryGetValue(address, out var cell) ? (cell.Formula ?? string.Empty).Trim() : string.Empty;
        }

        #endregion
        

        #region 计算表插入到图纸空间中的相关代码


        /// <summary>
        /// 构建计算表/设备表的临时 DWG 文件
        /// </summary>
        /// <returns>元组：(临时DWG文件路径, 生成的块名称)</returns>
        private (string dwgPath, string blockName) BuildCalcTablesTempDwg()
        {
            // ================== 第一步：收集界面数据 ==================

            // 定义一个列表，用于存储各个分表的数据结构：(Grid名称, 表头标题, 行数据列表)
            var sections = new List<(string GridName, string Header, List<CalcCsvTableRow> Rows)>();

            // 获取当前 CAD 的绘图比例，用于后续计算文字高度和行列尺寸
            var scale = VariableDictionary.wpfTextBoxScale; 

            try
            {
                // 优先从界面的动态容器 CalcDynamicHost 中收集数据，保证顺序与用户看到的界面一致
                if (CalcDynamicHost != null && CalcDynamicHost.Children.Count > 0)
                {
                    // 遍历容器中的每一个子控件
                    foreach (var child in CalcDynamicHost.Children)
                    {
                        // 判断子控件是否是 GroupBox（分组框），且其内容是 DataGrid（数据网格）
                        if (child is System.Windows.Controls.GroupBox gb && gb.Content is DataGrid dg && dg.ItemsSource is IEnumerable<CalcCsvTableRow> src)
                        {
                            // 提取分组框的标题作为表格的大标题
                            string header = (gb.Header?.ToString() ?? string.Empty).Trim();
                            // 提取 DataGrid 的名称作为内部标识
                            string gridName = (dg.Name ?? string.Empty).Trim();

                            // 将数据源转换为列表，并按 Sequence（序号）排序，保证行顺序正确
                            var rows = src.OrderBy(r => r.Sequence).ToList();

                            // 只有当 Grid 有名称且包含数据时，才加入收集列表
                            if (!string.IsNullOrWhiteSpace(gridName) && rows.Count > 0)
                            {
                                sections.Add((gridName, header, rows));
                            }
                        }
                    }
                }

                // 如果界面容器为空（例如还没加载完），则回退到内存中的数据源 _calcCsvTableSources
                if (sections.Count == 0 && _calcCsvTableSources != null)
                {
                    // 遍历内存中的数据源，按 Key 排序以保证稳定性
                    foreach (var kv in _calcCsvTableSources.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
                    {
                        // 获取行数据并排序
                        var rows = kv.Value?.OrderBy(r => r.Sequence).ToList() ?? new List<CalcCsvTableRow>();
                        if (rows.Count == 0) continue; // 跳过空表

                        // 根据 GridName 获取对应的显示标题
                        string header = GetCalcGroupHeaderByGridName(kv.Key);
                        sections.Add((kv.Key, header, rows));
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录收集数据过程中的异常，防止程序崩溃
                LogManager.Instance.LogWarning($"收集计算表数据失败: {ex.Message}");
            }

            // 如果没有收集到任何数据，抛出异常，终止后续操作
            if (sections.Count == 0)
                throw new InvalidOperationException("当前没有可插入的计算表数据。");

            // ================== 第二步：准备临时文件路径 ==================

            // 定义临时文件夹路径：系统临时目录 + 插件名称 + CalcTables
            string tempDir = Path.Combine(Path.GetTempPath(), "GB_NewCadPlus_IV", "CalcTables");
            // 如果目录不存在，则创建
            if (!Directory.Exists(tempDir))
                Directory.CreateDirectory(tempDir);

            // 生成唯一的文件名，包含时间戳，避免冲突
            string fileName = $"CalcTables_{DateTime.Now:yyyyMMdd_HHmmss}.dwg";
            string fullPath = Path.Combine(tempDir, fileName);

            // 块名称通常与文件名一致（不含扩展名）
            string blockName = Path.GetFileNameWithoutExtension(fileName);

            // ================== 第三步：在内存数据库中构建表格 ==================

            // 创建一个新的、空的 AutoCAD 数据库对象 (true: 无文档关联, true: 使用默认单位)
            using (var tempDb = new Autodesk.AutoCAD.DatabaseServices.Database(true, true))
            {
                // 开启事务，用于操作这个临时数据库
                using (var tr = tempDb.TransactionManager.StartTransaction())
                {
                    // --- 3.1 确保图层存在 ---
                    const string calcLayer = "TJ(计算表)"; // 定义图层名称

                    // 获取图层表，以写模式打开
                    var lt = (Autodesk.AutoCAD.DatabaseServices.LayerTable)tr.GetObject(
                        tempDb.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    // 如果图层不存在，则创建新图层
                    if (!lt.Has(calcLayer))
                    {
                        var ltr = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord { Name = calcLayer };
                        lt.Add(ltr);
                        tr.AddNewlyCreatedDBObject(ltr, true); // 注册到事务
                    }

                    // --- 3.2 创建块定义 ---
                    // 获取块表，以写模式打开
                    var bt = (Autodesk.AutoCAD.DatabaseServices.BlockTable)tr.GetObject(
                        tempDb.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    // 创建新的块定义记录
                    var btr = new Autodesk.AutoCAD.DatabaseServices.BlockTableRecord
                    {
                        Name = blockName // 设置块名
                    };
                    bt.Add(btr); // 将块定义添加到块表
                    tr.AddNewlyCreatedDBObject(btr, true); // 注册到事务

                    // --- 3.3 定义表格样式参数 ---

                    // 【关键需求】设置文字高度
                    // 基础打印高度设为 2.5mm，乘以比例 scale 得到模型空间中的实际高度
                    // 例如：比例 1:100，则 textHeight = 2.5 * 100 = 250.0
                    double baseTextHeight = 2.5;
                    double textHeight = baseTextHeight * scale;

                    // 行高：通常比文字高度稍大，设为文字高度的 1.5~2 倍，这里沿用原逻辑 5 * scale
                    double rowHeight = 6 * scale;

                    // 表格之间的垂直间距
                    double gap = 6 * scale;

                    // 列宽定义：{序号, 参数名称, 值, 单位}，单位均为绘图单位
                    double[] colWidths = { 12 * scale, 55 * scale, 28 * scale, 18 * scale };

                    // Y 方向累计偏移量，用于让多个表格纵向堆叠，不重叠
                    double yOffset = 0.0;

                    // --- 3.4 遍历每个分表，生成 AutoCAD Table 对象 ---
                    foreach (var sec in sections)
                    {
                        int dataCount = sec.Rows.Count;
                        if (dataCount <= 0) continue; // 跳过无数据的表

                        // 计算总行数：1行标题 + 1行列头 + N行数据
                        int rows = dataCount + 2;
                        int cols = 4; // 固定4列

                        // 创建新的 Table 对象
                        var table = new Autodesk.AutoCAD.DatabaseServices.Table();

                        // 应用数据库默认设置（字体、颜色等）
                        table.SetDatabaseDefaults(tempDb);

                        // 设置表格所属图层
                        table.Layer = calcLayer;

                        // 设置表格的行数和列数
                        table.SetSize(rows, cols);

                        // 设置表格插入点位置：X=0, Y=-yOffset (向下排列), Z=0
                        table.Position = new Autodesk.AutoCAD.Geometry.Point3d(0.0, -yOffset, 0.0);

                        // 【关键修改】更稳健地设置文字高度
                        // 旧方法 table.SetTextHeight 已过时或有时不生效，建议直接设置 Cells 的属性
                        // 这里我们遍历所有单元格进行设置，确保万无一失
                        for (int r = 0; r < rows; r++)
                        {
                            for (int c = 0; c < cols; c++)
                            {
                                // 设置单元格文字高度
                                table.Cells[r, c].TextHeight = textHeight;

                                // 设置单元格对齐方式：居中
                                table.Cells[r, c].Alignment = Autodesk.AutoCAD.DatabaseServices.CellAlignment.MiddleCenter;
                            }
                        }

                        // 可选：如果你想让标题行文字更大，可以单独设置第0行
                        // for (int c = 0; c < cols; c++) { table.Cells[0, c].TextHeight = textHeight * 1.2; }

                        // 设置行高
                        table.SetRowHeight(rowHeight);

                        // 设置每列的宽度
                        for (int c = 0; c < cols; c++)
                            table.SetColumnWidth(c, colWidths[c]);

                        // --- 填充内容 ---

                        // 1. 标题行 (Row 0)
                        // 如果有自定义 Header 则用 Header，否则用 GridName
                        string titleText = string.IsNullOrWhiteSpace(sec.Header) ? sec.GridName : sec.Header;
                        table.Cells[0, 0].TextString = titleText;
                        // 标题行通常合并单元格，但为了兼容性和简单起见，这里只填第一个格，或者你可以选择合并
                        // table.MergeCells(CellRange.Create(table, 0, 0, 0, 3)); // 如果需要合并取消注释

                        // 2. 列头行 (Row 1)
                        table.Cells[1, 0].TextString = "序号";
                        table.Cells[1, 1].TextString = "参数名称";
                        table.Cells[1, 2].TextString = "值";
                        table.Cells[1, 3].TextString = "单位";

                        // 3. 数据行 (Row 2 开始)
                        for (int i = 0; i < sec.Rows.Count; i++)
                        {
                            int r = i + 2; // 数据行索引从2开始
                            var row = sec.Rows[i]; // 获取当前行数据对象

                            table.Cells[r, 0].TextString = row.Sequence.ToString();       // 序号
                            table.Cells[r, 1].TextString = row.ParameterName ?? string.Empty; // 参数名
                            table.Cells[r, 2].TextString = row.ValueText ?? string.Empty;     // 值
                            table.Cells[r, 3].TextString = row.Unit ?? string.Empty;          // 单位
                        }

                        // --- 将表格添加到块定义中 ---
                        btr.AppendEntity(table); // 附加实体
                        tr.AddNewlyCreatedDBObject(table, true); // 注册到事务

                        // --- 更新下一个表格的偏移量 ---
                        // 当前表格高度 + 间距
                        yOffset += rows * rowHeight + gap;
                    }

                    // 提交事务，保存对临时数据库的修改
                    tr.Commit();
                }

                // 将内存中的临时数据库保存为磁盘上的 .dwg 文件
                tempDb.SaveAs(fullPath, Autodesk.AutoCAD.DatabaseServices.DwgVersion.Current);
            }

            // 返回文件路径和块名，供调用者使用
            return (fullPath, blockName);
        }
        
        #endregion

        #region 新管道设备表WPF页面相关代码

        private void 管道设备表开关_Btn_Click(object sender, RoutedEventArgs e)
        {
            // 直接调用统一命令入口，避免编译期依赖额外提取方法
            var generator = new UnifiedTableGenerator();
            generator.GeneratePipeTableFromSelection();
        }

        #endregion

        private async void 上传客户端_Btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await UploadClientPackageAsync();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"上传客户端失败: {ex.Message}");
                MessageBox.Show($"上传客户端失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void 下载最新版客户端_Click(object sender, RoutedEventArgs e)
        {
            _ = CheckClientVersionAndPromptUpdateAsync();
        }

        private void 同步图元_Click(object sender, RoutedEventArgs e)
        {
            _ = BuildAndRunSyncAsync();
        }
        
        /// <summary>
        /// 同步所有服务器图元及预览图到本地缓存（通过 HTTP API 下载）。
        /// 替代原先直接访问共享文件夹的同步逻辑。
        /// </summary>
        private async Task BuildAndRunSyncAsync()
        {
            // 并发控制：避免同时多次同步
            await _syncSemaphore.WaitAsync();
            try
            {
                // 1. 检查数据库可用性
                if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
                {
                    LogManager.Instance.LogWarning("同步失败：数据库不可用。");
                    MessageBox.Show("数据库未连接，无法执行同步。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 2. 获取所有需要同步的图元（可根据业务筛选，这里以“所有未删除的图元”为例）
                List<FileStorage> allFiles;
                try
                {
                    allFiles = await _databaseManager.GetAllActiveFileStoragesAsync(); // 需要自行实现或使用现有类似方法
                    if (allFiles == null || allFiles.Count == 0)
                    {
                        LogManager.Instance.LogInfo("同步结束：没有需要同步的图元。");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogError($"获取图元列表失败: {ex.Message}");
                    MessageBox.Show("无法从数据库获取图元列表，请稍后重试。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                LogManager.Instance.LogInfo($"开始同步 {allFiles.Count} 个图元...");

                int successCount = 0;
                int failCount = 0;

                // 3. 逐一下载每个图元的 DWG 和预览图
                for (int i = 0; i < allFiles.Count; i++)
                {
                    var file = allFiles[i];
                    try
                    {
                        // 3.1 下载 DWG 文件到本地缓存
                        string dwgLocalPath = await ServerFileService.EnsureDwgCacheAsync(file,GetPath.DwgCachePath);
                        if (string.IsNullOrWhiteSpace(dwgLocalPath))
                        {
                            LogManager.Instance.LogWarning($"同步：图元 {file.DisplayName} (ID:{file.Id}) DWG 下载失败，跳过预览图。");
                            failCount++;
                            continue;
                        }

                        // 3.2 下载预览图到本地缓存
                        string previewLocalPath = await ServerFileService.EnsurePreviewCacheAsync(file, GetPath.PreviewCachePath);
                        if (string.IsNullOrWhiteSpace(previewLocalPath))
                        {
                            LogManager.Instance.LogWarning($"同步：图元 {file.DisplayName} (ID:{file.Id}) 预览图下载失败，但 DWG 已下载。");
                            // 不算完全失败，但可记录
                        }

                        successCount++;
                        LogManager.Instance.LogInfo($"同步进度：{i + 1}/{allFiles.Count} - {file.DisplayName} 完成");
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        LogManager.Instance.LogError($"同步图元 {file.DisplayName} (ID:{file.Id}) 时发生异常: {ex.Message}");
                    }
                }

                // 4. 完成报告
                string msg = $"同步完成：成功 {successCount} 个，失败 {failCount} 个。";
                LogManager.Instance.LogInfo(msg);
                MessageBox.Show(msg, "同步结果", MessageBoxButton.OK, MessageBoxImage.Information);

                // 5. （可选）刷新 UI 列表
                await RefreshFilesForCurrentCategoryAsync();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"BuildAndRunSyncAsync 发生未处理异常: {ex.Message}");
                MessageBox.Show("同步过程中发生严重错误，请查看日志。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _syncSemaphore.Release();
            }
        }
        
        /// <summary>
        /// 检查服务器端客户端版本，并提示是否需要更新。
        /// </summary>
        private async Task CheckClientVersionAndPromptUpdateAsync()
        {
            if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
            {
                MessageBox.Show("当前数据库未连接，无法检查客户端版本。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var serverVersion = await _databaseManager.GetServerClientVersionAsync();
                if (string.IsNullOrWhiteSpace(serverVersion))
                {
                    MessageBox.Show("服务器端未配置客户端版本。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var localVersion = TextBox客户端版本?.Text?.Trim() ?? string.Empty;
                if (IsRemoteVersionNewer(serverVersion, localVersion))
                {
                    var downloadedPath = await DownloadClientPackageAsync(serverVersion);
                    if (!string.IsNullOrWhiteSpace(downloadedPath))
                    {
                        MessageBox.Show($"已下载新版本客户端包：{downloadedPath}", "发现新版本", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show($"检测到新版本：服务器 {serverVersion}，本地 {localVersion}。客户端包尚未配置。", "发现新版本", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                else
                {
                    MessageBox.Show($"当前已是最新版本：{localVersion}。", "版本检查", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"检查客户端版本失败: {ex.Message}");
                MessageBox.Show($"检查客户端版本失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// 获取当前使用的本地存储根目录。
        /// </summary>
        private string GetConfiguredStorageRoot()
        {
            var configured = TextBoxSetStoragePath?.Text?.Trim();
            if (string.IsNullOrWhiteSpace(configured))
            {
                configured = @"D:\GB_Tools\Cad_Sw_Library";
            }

            return configured;
        }
        
        /// <summary>
        /// 规范化候选路径。
        /// </summary>
        private static string? NormalizePathCandidate(string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return null;
            }

            return path.Replace('/', '\\').Trim();
        }
        
        /// <summary>
        /// 选择并上传客户端包，同时保存版本与包路径。
        /// </summary>
        private async Task UploadClientPackageAsync()
        {
            if (_databaseManager == null || !_databaseManager.IsDatabaseAvailable)
            {
                MessageBox.Show("当前数据库未连接，无法上传客户端包。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var version = TextBox插件版本?.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(version))
            {
                MessageBox.Show("请先填写服务器版本号。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择客户端安装包",
                Filter = "客户端包|*.zip;*.7z;*.rar;*.exe;*.msi|所有文件|*.*",
                CheckFileExists = true,
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() != true)
            {
                return;
            }

            var sourcePath = openFileDialog.FileName;
            var targetDir = Path.Combine(GetConfiguredStorageRoot(), "ClientPackages");
            if (!Directory.Exists(targetDir))
            {
                Directory.CreateDirectory(targetDir);
            }

            var targetPath = Path.Combine(targetDir, $"Client_{version}{Path.GetExtension(sourcePath)}");
            await Task.Run(() => File.Copy(sourcePath, targetPath, true));

            await _databaseManager.SetSystemConfigValueAsync("ClientVersion", version);
            await _databaseManager.SetSystemConfigValueAsync("ClientPackagePath", targetPath);

            TextBox客户端版本.Text = version;
            MessageBox.Show($"客户端包已保存：{targetPath}\n版本已写入数据库：{version}", "上传完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        
        /// <summary>
        /// 下载客户端包到本地缓存目录。
        /// </summary>
        private async Task<string> DownloadClientPackageAsync(string serverVersion)
        {
            var packagePath = await _databaseManager!.GetSystemConfigValueAsync("ClientPackagePath");
            if (string.IsNullOrWhiteSpace(packagePath))
            {
                return string.Empty;
            }

            if (!File.Exists(packagePath))
            {
                return string.Empty;
            }

            var downloadDir = Path.Combine(GetConfiguredStorageRoot(), "ClientPackages");
            if (!Directory.Exists(downloadDir))
            {
                Directory.CreateDirectory(downloadDir);
            }

            var targetPath = Path.Combine(downloadDir, $"Client_{serverVersion}{Path.GetExtension(packagePath)}");
            await Task.Run(() => File.Copy(packagePath, targetPath, true));
            return targetPath;
        }

        /// <summary>
        /// 判断服务器版本是否高于本地版本。
        /// </summary>
        private static bool IsRemoteVersionNewer(string serverVersion, string localVersion)
        {
            if (string.IsNullOrWhiteSpace(serverVersion))
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(localVersion))
            {
                return true;
            }

            var serverParts = serverVersion.Split('.');
            var localParts = localVersion.Split('.');
            var length = Math.Max(serverParts.Length, localParts.Length);

            for (var i = 0; i < length; i++)
            {
                var serverPart = i < serverParts.Length && int.TryParse(serverParts[i], out var s) ? s : 0;
                var localPart = i < localParts.Length && int.TryParse(localParts[i], out var l) ? l : 0;

                if (serverPart > localPart)
                {
                    return true;
                }

                if (serverPart < localPart)
                {
                    return false;
                }
            }

            return false;
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
        /// 点击按钮，打开文件夹选择对话框，将选定的路径保存到 TextBoxSetStoragePath 中
        /// </summary>
        private void 指定路径按键_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "请选择客户端的临时文件交换根目录";
                    dialog.RootFolder = Environment.SpecialFolder.Desktop;
                    if (!string.IsNullOrWhiteSpace(TextBoxSetStoragePath.Text))
                        dialog.SelectedPath = TextBoxSetStoragePath.Text;

                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string selectedPath = dialog.SelectedPath;
                        if (string.IsNullOrWhiteSpace(selectedPath))
                        {
                            MessageBox.Show("未选择有效文件夹。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        // 设置 DWG 缓存路径（注意：不要包含文件名，必须是目录）
                        GetPath.DwgCachePath = selectedPath;
                        // 预览图通常放在同一个上级目录下，也可以单独设置
                        // VariableDictionary.PreviewCachePath = Path.Combine(selected, "PreviewCache");
                        // 如果希望预览图和 DWG 在同一目录，也可以
                        GetPath.PreviewCachePath = selectedPath;


                        // 更新界面和全局变量
                        TextBoxSetStoragePath.Text = selectedPath;
                        GetPath._cacheStoragePath = selectedPath;

                        // 自动取消“使用默认路径”复选框
                        CheckBoxUseDPath.IsChecked = false;

                        // 创建必要的子目录
                        try
                        {
                            Directory.CreateDirectory(GetPath.PreviewCachePath);
                            Directory.CreateDirectory(GetPath.DwgCachePath);
                            LogManager.Instance.LogInfo($"存储根路径设置为: {selectedPath}");
                            LogManager.Instance.LogInfo($"DwgCachePath 已更改为: {GetPath.DwgCachePath}");
                            LogManager.Instance.LogInfo($"PreviewCachePath 已更改为: {GetPath.PreviewCachePath}");
                        }
                        catch (Exception ex)
                        {
                            LogManager.Instance.LogError($"创建缓存目录失败: {ex.Message}");
                            MessageBox.Show("无法在指定目录下创建缓存文件夹，请检查权限。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"选择文件夹失败: {ex.Message}");
                MessageBox.Show("打开文件夹选择对话框时发生错误: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        /// <summary>
        /// 当用户勾选“使用默认路径”时，清空手动设置的路径，恢复默认智能选择。
        /// </summary>
        private void CheckBoxUseDPath_Checked(object sender, RoutedEventArgs e)
        {
            // 清空手动路径
            TextBoxSetStoragePath.Text = string.Empty;
            GetPath._cacheStoragePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "GB_CADPLUS"); // 或 string.Empty，视 GetBaseStoragePath 的判空逻辑而定

            LogManager.Instance.LogInfo("已切换为使用默认存储路径");
        }

        /// <summary>
        /// 当用户取消勾选“使用默认路径”时，不做特殊操作。
        /// 用户可以后续通过“指定路径…”按钮重新选择路径。
        /// </summary>
        private void CheckBoxUseDPath_Unchecked(object sender, RoutedEventArgs e)
        {
            // 这里可以不写，或者记录日志
            LogManager.Instance.LogInfo("已取消使用默认存储路径，将允许手动指定路径");
        }

        private void 重载计算表_Click(object sender, RoutedEventArgs e)
        {
            var pipeCalcViewModel = new PipeCalcViewModel();// 创建一个 PipeCalcViewModel 实例
                pipeCalcViewModel.LoadData(); // 加载数据
        }

        private async void 添加主规范库_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdminUser(VariableDictionary._userName ?? string.Empty))
            {
                MessageBox.Show("只有管理员可以添加规范库。", "权限不足", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            await CreateStandardCategoryAsync(null);
        }

        private async void 添加子规范库_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdminUser(VariableDictionary._userName ?? string.Empty))
            {
                MessageBox.Show("只有管理员可以添加规范库。", "权限不足", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!(SpecificationTreeView.SelectedItem is CategoryTreeNode selectedNode)
                || !(selectedNode.Data is StandardManagementCategoryClient selectedCategory)
                || selectedCategory.ParentId.HasValue)
            {
                MessageBox.Show("请先在左侧选择一个主规范库，再添加子规范库。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            await CreateStandardCategoryAsync(selectedCategory.Id);
        }

        /// <summary>
        /// 创建主分类或子分类。
        /// </summary>
        private async Task CreateStandardCategoryAsync(long? parentId)
        {
            StandardCategoryInputResult? input = ShowStandardCategoryInputDialog(parentId.HasValue);
            if (input == null) return;

            try
            {
                StandardCategoryCommandClientRequest request = new StandardCategoryCommandClientRequest
                {
                    ParentId = parentId,
                    Code = input.Code,
                    Name = input.Name,
                    Description = input.Description,
                    SortOrder = 0
                };
                StandardManagementOperationClientResponse response;
                try
                {
                    response = await _standardManagementApiService
                        .CreateManagementCategoryAsync(request, VariableDictionary._userName ?? string.Empty)
                        .ConfigureAwait(true);
                }
                catch (StandardCategoryConflictException conflict)
                {
                    string duplicateText = string.Join("\n", conflict.Duplicates.Select(item =>
                        $"ID={item.Id}，名称={item.Name}，CODE={item.Code}，原因={item.DuplicateReason}"));
                    MessageBoxResult updateResult = MessageBox.Show(
                        $"发现同层级重复规范库：\n{duplicateText}\n\n是否使用当前填写的信息更新第一条重复规范库？",
                        "规范库重复提示",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);
                    if (updateResult != MessageBoxResult.Yes || conflict.Duplicates.Count == 0)
                        return;

                    long duplicateId = conflict.Duplicates[0].Id;
                    response = await _standardManagementApiService
                        .UpdateManagementCategoryAsync(duplicateId, request, VariableDictionary._userName ?? string.Empty)
                        .ConfigureAwait(true);
                }

                MessageBox.Show(response.Message,
                    response.Success ? "规范库保存成功" : "规范库保存失败",
                    MessageBoxButton.OK,
                    response.Success ? MessageBoxImage.Information : MessageBoxImage.Warning);

                if (response.Success)
                    await RefreshStandardTreeAsync().ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"创建规范分类失败：{ex.Message}");
                MessageBox.Show($"创建规范分类失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 在一个窗口中集中填写规范库名称、编码和说明。
        /// </summary>
        private StandardCategoryInputResult? ShowStandardCategoryInputDialog(
            bool isChildCategory,
            StandardManagementCategoryClient? existingCategory = null)
        {
            var dialog = new Window
            {
                Title = existingCategory == null
                    ? (isChildCategory ? "添加子规范库" : "添加主规范库")
                    : (isChildCategory ? "修改子规范库" : "修改主规范库"),
                Width = 430,
                Height = 300,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false
            };

            var root = new Grid { Margin = new Thickness(18) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var nameLabel = new TextBlock { Text = "规范库名称：", Margin = new Thickness(0, 0, 0, 4) };
            Grid.SetRow(nameLabel, 0);
            root.Children.Add(nameLabel);

            var nameBox = new System.Windows.Controls.TextBox
            {
                Text = existingCategory?.Name ?? string.Empty,
                Margin = new Thickness(0, 0, 0, 10),
                MinHeight = 26
            };
            Grid.SetRow(nameBox, 1);
            root.Children.Add(nameBox);

            var codeLabel = new TextBlock { Text = "规范库编码：", Margin = new Thickness(0, 0, 0, 4) };
            Grid.SetRow(codeLabel, 2);
            root.Children.Add(codeLabel);

            var codeBox = new System.Windows.Controls.TextBox
            {
                Text = existingCategory?.Code ?? string.Empty,
                Margin = new Thickness(0, 0, 0, 10),
                MinHeight = 26
            };
            Grid.SetRow(codeBox, 3);
            root.Children.Add(codeBox);

            var descriptionPanel = new StackPanel { Margin = new Thickness(0, 0, 0, 10) };
            var descriptionLabel = new TextBlock { Text = "规范库说明（可选）：", Margin = new Thickness(0, 0, 0, 4) };
            var descriptionBox = new System.Windows.Controls.TextBox
            {
                Text = existingCategory?.Description ?? string.Empty,
                MinHeight = 26
            };
            descriptionPanel.Children.Add(descriptionLabel);
            descriptionPanel.Children.Add(descriptionBox);
            Grid.SetRow(descriptionPanel, 4);
            root.Children.Add(descriptionPanel);

            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            var cancelButton = new Button { Content = "取消", Width = 75, Margin = new Thickness(0, 0, 8, 0) };
            var confirmButton = new Button { Content = "确定", Width = 75 };
            cancelButton.Click += (_, _) => dialog.DialogResult = false;
            confirmButton.Click += (_, _) =>
            {
                if (string.IsNullOrWhiteSpace(nameBox.Text))
                {
                    MessageBox.Show(dialog, "请输入规范库名称。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    nameBox.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(codeBox.Text))
                {
                    MessageBox.Show(dialog, "请输入规范库编码。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    codeBox.Focus();
                    return;
                }

                dialog.DialogResult = true;
            };
            buttonPanel.Children.Add(cancelButton);
            buttonPanel.Children.Add(confirmButton);
            Grid.SetRow(buttonPanel, 5);
            root.Children.Add(buttonPanel);

            dialog.Content = root;
            dialog.Loaded += (_, _) => nameBox.Focus();
            bool? result = dialog.ShowDialog();
            if (result != true) return null;

            return new StandardCategoryInputResult
            {
                Name = nameBox.Text.Trim(),
                Code = codeBox.Text.Trim().ToUpperInvariant(),
                Description = descriptionBox.Text.Trim()
            };
        }

        private sealed class StandardCategoryInputResult
        {
            public string Name { get; set; } = string.Empty;
            public string Code { get; set; } = string.Empty;
            public string Description { get; set; } = string.Empty;
        }

        private async void 修改规范库_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdminUser(VariableDictionary._userName ?? string.Empty))
            {
                MessageBox.Show("只有管理员可以修改规范库。", "权限不足", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!(SpecificationTreeView.SelectedItem is CategoryTreeNode selectedNode)
                || !(selectedNode.Data is StandardManagementCategoryClient selectedCategory))
            {
                MessageBox.Show("请先在左侧选择要修改的主规范库或子规范库。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            StandardCategoryInputResult? input = ShowStandardCategoryInputDialog(
                selectedCategory.ParentId.HasValue,
                selectedCategory);
            if (input == null) return;

            try
            {
                StandardManagementOperationClientResponse response = await _standardManagementApiService
                    .UpdateManagementCategoryAsync(selectedCategory.Id, new StandardCategoryCommandClientRequest
                    {
                        ParentId = selectedCategory.ParentId,
                        Code = input.Code,
                        Name = input.Name,
                        Description = input.Description,
                        SortOrder = selectedCategory.SortOrder
                    }, VariableDictionary._userName ?? string.Empty)
                    .ConfigureAwait(true);

                MessageBox.Show(response.Message,
                    response.Success ? "规范库修改成功" : "规范库修改失败",
                    MessageBoxButton.OK,
                    response.Success ? MessageBoxImage.Information : MessageBoxImage.Warning);
                if (response.Success)
                    await RefreshStandardTreeAsync().ConfigureAwait(true);
            }
            catch (StandardCategoryConflictException conflict)
            {
                string duplicateText = string.Join("\n", conflict.Duplicates.Select(item =>
                    $"ID={item.Id}，名称={item.Name}，CODE={item.Code}，原因={item.DuplicateReason}"));
                MessageBox.Show($"修改后的名称或 CODE 与以下规范库重复：\n{duplicateText}",
                    "规范库重复提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"修改规范分类失败：{ex.Message}");
                MessageBox.Show($"修改规范分类失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task CreateAndUploadStandardVersionAsync()
        {
            if (!(SpecificationTreeView.SelectedItem is CategoryTreeNode selectedNode)
                || !(selectedNode.Data is StandardManagementSeriesClient series))
            {
                MessageBox.Show("请先在左侧选择一个规范系列。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string versionNo = Microsoft.VisualBasic.Interaction.InputBox("请输入规范版本号：", "新建规范版本", DateTime.Now.ToString("yyyyMMdd"));
            if (string.IsNullOrWhiteSpace(versionNo)) return;

            string changeSummary = Microsoft.VisualBasic.Interaction.InputBox("请输入本次变更说明（可选）：", "新建规范版本", string.Empty);
            using var dialog = new OpenFileDialog
            {
                Filter = "规范附件 (*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.json)|*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.json|所有文件 (*.*)|*.*",
                Title = "选择规范版本附件",
                Multiselect = false
            };
            if (dialog.ShowDialog() != DialogResult.OK) return;

            try
            {
                string operatorName = VariableDictionary._userName ?? string.Empty;
                StandardManagementOperationClientResponse created = await _standardManagementApiService
                    .CreateManagementVersionAsync(new StandardVersionCreateClientRequest
                    {
                        SeriesId = series.Id,
                        VersionNo = versionNo.Trim(),
                        VersionLabel = versionNo.Trim(),
                        ChangeSummary = changeSummary?.Trim() ?? string.Empty,
                        SourceType = "DOCUMENT"
                    }, operatorName)
                    .ConfigureAwait(true);

                if (!created.Success)
                {
                    MessageBox.Show(created.Message, "创建规范版本失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                StandardFileUploadClientResponse uploaded = await _standardManagementApiService
                    .UploadManagementFileAsync(created.Id, dialog.FileName, operatorName)
                    .ConfigureAwait(true);
                MessageBox.Show(uploaded.Message, uploaded.Success ? "规范版本创建成功" : "附件上传失败", MessageBoxButton.OK,
                    uploaded.Success ? MessageBoxImage.Information : MessageBoxImage.Warning);

                if (uploaded.Success)
                    await RefreshStandardTreeAsync().ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"创建规范版本失败：{ex.Message}");
                MessageBox.Show($"创建规范版本失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task RefreshStandardTreeAsync()
        {
            StandardManagementTreeClientResponse response = await _standardManagementApiService
                .GetManagementTreeAsync()
                .ConfigureAwait(true);
            if (response.Success)
                BuildStandardTree(response);
        }

        private async void 删除规范库_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdminUser(VariableDictionary._userName ?? string.Empty))
            {
                MessageBox.Show("只有管理员可以删除规范库。", "权限不足", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!(SpecificationTreeView.SelectedItem is CategoryTreeNode selectedNode))
            {
                MessageBox.Show("请先选择要删除的规范库。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                if (selectedNode.Data is StandardManagementCategoryClient category)
                {
                    MessageBoxResult categoryConfirm = MessageBox.Show(
                        $"确定删除规范库“{category.Name}（{category.Code}）”吗？\n删除后该分类及其下级内容将不再显示。",
                        "确认删除规范库", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    if (categoryConfirm != MessageBoxResult.Yes) return;

                    StandardManagementOperationClientResponse categoryResponse = await _standardManagementApiService
                        .DeleteManagementCategoryAsync(category.Id, VariableDictionary._userName ?? string.Empty)
                        .ConfigureAwait(true);
                    MessageBox.Show(categoryResponse.Message,
                        categoryResponse.Success ? "删除完成" : "删除失败",
                        MessageBoxButton.OK,
                        categoryResponse.Success ? MessageBoxImage.Information : MessageBoxImage.Warning);
                    if (categoryResponse.Success)
                        await RefreshStandardTreeAsync().ConfigureAwait(true);
                    return;
                }

                if (!(selectedNode.Data is StandardManagementSeriesClient series))
                {
                    MessageBox.Show("当前节点不支持删除。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                List<StandardDocumentVersionClient> versions = await _standardManagementApiService
                    .GetManagementVersionsAsync(series.Id)
                    .ConfigureAwait(true);
                StandardDocumentVersionClient? current = versions.FirstOrDefault(item => item.IsCurrent);
                if (current == null)
                {
                    MessageBox.Show("当前规范系列没有可删除的有效版本。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                MessageBoxResult confirm = MessageBox.Show(
                    $"确定软删除当前版本“{current.VersionNo}”吗？历史数据仍会保留。",
                    "确认删除规范版本", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (confirm != MessageBoxResult.Yes) return;

                StandardManagementOperationClientResponse response = await _standardManagementApiService
                    .DeleteManagementVersionAsync(current.Id, VariableDictionary._userName ?? string.Empty)
                    .ConfigureAwait(true);
                MessageBox.Show(response.Message, response.Success ? "删除完成" : "删除失败", MessageBoxButton.OK,
                    response.Success ? MessageBoxImage.Information : MessageBoxImage.Warning);
                if (response.Success) await RefreshStandardTreeAsync().ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"删除规范版本失败：{ex.Message}");
                MessageBox.Show($"删除规范版本失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 右键点击树节点时，先把该节点设为当前选中节点，避免菜单操作作用于旧节点。
        /// </summary>
        private void 规范树右键按下(object sender, MouseButtonEventArgs e)
        {
            DependencyObject? source = e.OriginalSource as DependencyObject;
            while (source != null && !(source is TreeViewItem))
                source = VisualTreeHelper.GetParent(source);

            if (source is TreeViewItem item)
            {
                item.IsSelected = true;
                item.Focus();
                LogManager.Instance.LogInfo("已通过右键选中规范树节点。");
            }
        }

        /// <summary>
        /// 打开右键菜单时，根据当前节点类型设置菜单可用状态。
        /// </summary>
        private void 规范树右键菜单_Opened(object sender, RoutedEventArgs e)
        {
            if (!(sender is ContextMenu menu)) return;
            bool isCategory = SpecificationTreeView.SelectedItem is CategoryTreeNode node
                && node.Data is StandardManagementCategoryClient;
            bool isSeries = SpecificationTreeView.SelectedItem is CategoryTreeNode seriesNode
                && seriesNode.Data is StandardManagementSeriesClient;
            bool isUncategorizedSeries = SpecificationTreeView.SelectedItem is CategoryTreeNode oldSeriesNode
                && oldSeriesNode.Data is StandardManagementSeriesClient oldSeries
                && !oldSeries.CategoryId.HasValue;

            SetContextMenuItemEnabled(menu, "添加子规范库", isCategory && IsMainCategorySelected());
            SetContextMenuItemEnabled(menu, "修改规范库", isCategory);
            SetContextMenuItemEnabled(menu, "移动规范位置", isCategory || isUncategorizedSeries);
            SetContextMenuItemEnabled(menu, "删除规范库", isCategory || isSeries);
        }

        private bool IsMainCategorySelected()
        {
            return SpecificationTreeView.SelectedItem is CategoryTreeNode node
                && node.Data is StandardManagementCategoryClient category
                && !category.ParentId.HasValue;
        }

        private static void SetContextMenuItemEnabled(ContextMenu menu, string header, bool isEnabled)
        {
            MenuItem? item = menu.Items.OfType<MenuItem>().FirstOrDefault(menuItem =>
                string.Equals(menuItem.Header?.ToString(), header, StringComparison.Ordinal));
            if (item != null) item.IsEnabled = isEnabled;
        }

        private void 右键添加主规范库_Click(object sender, RoutedEventArgs e)
        {
            添加主规范库_Click(sender, e);
        }

        private void 右键添加子规范库_Click(object sender, RoutedEventArgs e)
        {
            添加子规范库_Click(sender, e);
        }

        private void 右键修改规范库_Click(object sender, RoutedEventArgs e)
        {
            修改规范库_Click(sender, e);
        }

        private async void 右键移动规范位置_Click(object sender, RoutedEventArgs e)
        {
            await MoveSelectedStandardCategoryAsync().ConfigureAwait(true);
        }

        private async void 右键删除规范库_Click(object sender, RoutedEventArgs e)
        {
            删除规范库_Click(sender, e);
        }

        private void 右键展开折叠规范_Click(object sender, RoutedEventArgs e)
        {
            展开_折叠规范_Click(sender, e);
        }

        private void 右键刷新架构_Click(object sender, RoutedEventArgs e)
        {
            刷新规范_Click(sender, e);
        }

        /// <summary>
        /// 选择新的父分类并通过服务器移动当前规范分类。
        /// </summary>
        private async Task MoveSelectedStandardCategoryAsync()
        {
            if (!IsAdminUser(VariableDictionary._userName ?? string.Empty))
            {
                MessageBox.Show("只有管理员可以移动规范库。", "权限不足", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!(SpecificationTreeView.SelectedItem is CategoryTreeNode selectedNode))
            {
                MessageBox.Show("请先选择要移动的主规范库或子规范库。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (selectedNode.Data is StandardManagementSeriesClient oldSeries && !oldSeries.CategoryId.HasValue)
            {
                await MoveUncategorizedSeriesAsync(oldSeries).ConfigureAwait(true);
                return;
            }

            if (!(selectedNode.Data is StandardManagementCategoryClient selectedCategory))
            {
                MessageBox.Show("当前规范系列已经归类，不能通过此菜单再次移动。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            List<CategoryTreeNode> excludedNodes = new List<CategoryTreeNode>();
            CollectStandardTreeNodes(selectedNode, excludedNodes);
            List<CategoryTreeNode> targets = new List<CategoryTreeNode>();
            foreach (CategoryTreeNode root in _standardTreeNodes)
                CollectStandardCategoryTargets(root, excludedNodes, targets);

            var targetItems = new List<StandardMoveTargetItem>
            {
                new StandardMoveTargetItem { ParentId = null, DisplayText = "（根级分类）" }
            };
            targetItems.AddRange(targets
                .Where(node => node.Data is StandardManagementCategoryClient)
                .Select(node => new StandardMoveTargetItem
                {
                    ParentId = ((StandardManagementCategoryClient)node.Data).Id,
                    DisplayText = new string('　', Math.Max(0, node.Level - 1)) + node.DisplayText
                }));

            StandardMoveTargetItem? target = ShowStandardMoveTargetDialog(selectedCategory, targetItems);
            if (target == null) return;

            try
            {
                StandardManagementOperationClientResponse response = await _standardManagementApiService
                    .MoveManagementCategoryAsync(
                        selectedCategory.Id,
                        target.ParentId,
                        VariableDictionary._userName ?? string.Empty)
                    .ConfigureAwait(true);
                MessageBox.Show(response.Message,
                    response.Success ? "移动成功" : "移动失败",
                    MessageBoxButton.OK,
                    response.Success ? MessageBoxImage.Information : MessageBoxImage.Warning);
                if (response.Success)
                    await RefreshStandardTreeAsync().ConfigureAwait(true);
            }
            catch (StandardCategoryConflictException conflict)
            {
                string duplicateText = string.Join("\n", conflict.Duplicates.Select(item =>
                    $"ID={item.Id}，名称={item.Name}，CODE={item.Code}，原因={item.DuplicateReason}"));
                MessageBox.Show($"目标位置存在重复规范库：\n{duplicateText}", "移动失败", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"移动规范分类失败：{ex.Message}");
                MessageBox.Show($"移动规范分类失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task MoveUncategorizedSeriesAsync(StandardManagementSeriesClient series)
        {
            List<StandardMoveTargetItem> targetItems = GetStandardCategoryTargetItems();
            StandardMoveTargetItem? target = ShowStandardMoveTargetDialogForSeries(series, targetItems);
            if (target?.ParentId == null) return;

            try
            {
                StandardManagementOperationClientResponse response = await _standardManagementApiService
                    .MoveManagementSeriesAsync(series.Id, target.ParentId.Value, VariableDictionary._userName ?? string.Empty)
                    .ConfigureAwait(true);
                MessageBox.Show(response.Message,
                    response.Success ? "移动成功" : "移动失败",
                    MessageBoxButton.OK,
                    response.Success ? MessageBoxImage.Information : MessageBoxImage.Warning);
                if (response.Success) await RefreshStandardTreeAsync().ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"移动旧规范失败：{ex.Message}");
                MessageBox.Show($"移动旧规范失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<StandardMoveTargetItem> GetStandardCategoryTargetItems()
        {
            List<StandardMoveTargetItem> targetItems = new List<StandardMoveTargetItem>();
            foreach (CategoryTreeNode root in _standardTreeNodes)
                CollectStandardCategoryTargetItems(root, targetItems);
            return targetItems;
        }

        private static void CollectStandardCategoryTargetItems(CategoryTreeNode node, List<StandardMoveTargetItem> result)
        {
            if (node.Data is StandardManagementCategoryClient category)
            {
                result.Add(new StandardMoveTargetItem
                {
                    ParentId = category.Id,
                    DisplayText = new string('　', Math.Max(0, node.Level)) + node.DisplayText
                });
            }
            foreach (CategoryTreeNode child in node.Children)
                CollectStandardCategoryTargetItems(child, result);
        }

        private static void CollectStandardTreeNodes(CategoryTreeNode node, List<CategoryTreeNode> result)
        {
            result.Add(node);
            foreach (CategoryTreeNode child in node.Children)
                CollectStandardTreeNodes(child, result);
        }

        private static void CollectStandardCategoryTargets(
            CategoryTreeNode node,
            List<CategoryTreeNode> excludedNodes,
            List<CategoryTreeNode> result)
        {
            if (excludedNodes.Contains(node)) return;
            if (node.Data is StandardManagementCategoryClient)
                result.Add(node);
            foreach (CategoryTreeNode child in node.Children)
                CollectStandardCategoryTargets(child, excludedNodes, result);
        }

        private StandardMoveTargetItem? ShowStandardMoveTargetDialog(
            StandardManagementCategoryClient selectedCategory,
            List<StandardMoveTargetItem> targetItems)
        {
            var dialog = new Window
            {
                Title = "移动规范位置",
                Width = 460,
                Height = 190,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false
            };
            var root = new Grid { Margin = new Thickness(18) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var hint = new TextBlock
            {
                Text = $"当前规范库：{selectedCategory.Name}（{selectedCategory.Code}）\n请选择新的上级位置：",
                Margin = new Thickness(0, 0, 0, 8)
            };
            Grid.SetRow(hint, 0);
            root.Children.Add(hint);

            var targetBox = new ComboBox
            {
                ItemsSource = targetItems,
                DisplayMemberPath = nameof(StandardMoveTargetItem.DisplayText),
                SelectedIndex = 0,
                MinHeight = 28,
                Margin = new Thickness(0, 0, 0, 12)
            };
            Grid.SetRow(targetBox, 1);
            root.Children.Add(targetBox);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            var cancel = new Button { Content = "取消", Width = 75, Margin = new Thickness(0, 0, 8, 0) };
            var confirm = new Button { Content = "确定", Width = 75 };
            cancel.Click += (_, _) => dialog.DialogResult = false;
            confirm.Click += (_, _) => dialog.DialogResult = true;
            buttons.Children.Add(cancel);
            buttons.Children.Add(confirm);
            Grid.SetRow(buttons, 2);
            root.Children.Add(buttons);

            dialog.Content = root;
            bool? result = dialog.ShowDialog();
            return result == true ? targetBox.SelectedItem as StandardMoveTargetItem : null;
        }

        private StandardMoveTargetItem? ShowStandardMoveTargetDialogForSeries(
            StandardManagementSeriesClient series,
            List<StandardMoveTargetItem> targetItems)
        {
            var dialog = new Window
            {
                Title = "移动旧规范位置",
                Width = 460,
                Height = 190,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false
            };
            var root = new Grid { Margin = new Thickness(18) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var hint = new TextBlock
            {
                Text = $"旧规范：{series.SeriesName}（{series.SeriesCode}）\n请选择归属规范库：",
                Margin = new Thickness(0, 0, 0, 8)
            };
            Grid.SetRow(hint, 0);
            root.Children.Add(hint);
            var targetBox = new ComboBox
            {
                ItemsSource = targetItems,
                DisplayMemberPath = nameof(StandardMoveTargetItem.DisplayText),
                SelectedIndex = targetItems.Count > 0 ? 0 : -1,
                MinHeight = 28,
                Margin = new Thickness(0, 0, 0, 12)
            };
            Grid.SetRow(targetBox, 1);
            root.Children.Add(targetBox);
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var cancel = new Button { Content = "取消", Width = 75, Margin = new Thickness(0, 0, 8, 0) };
            var confirm = new Button { Content = "确定", Width = 75, IsEnabled = targetItems.Count > 0 };
            cancel.Click += (_, _) => dialog.DialogResult = false;
            confirm.Click += (_, _) => dialog.DialogResult = true;
            buttons.Children.Add(cancel);
            buttons.Children.Add(confirm);
            Grid.SetRow(buttons, 2);
            root.Children.Add(buttons);
            dialog.Content = root;
            return dialog.ShowDialog() == true ? targetBox.SelectedItem as StandardMoveTargetItem : null;
        }

        private sealed class StandardMoveTargetItem
        {
            public long? ParentId { get; set; }
            public string DisplayText { get; set; } = string.Empty;
        }

        private void 展开_折叠规范_Click(object sender, RoutedEventArgs e)
        {
            bool shouldExpand = _standardTreeNodes.Any(node => !node.IsExpanded || node.Children.Any(child => !child.IsExpanded));
            foreach (CategoryTreeNode node in _standardTreeNodes)
                SetStandardTreeExpanded(node, shouldExpand);

            SpecificationTreeView.Items.Refresh();
            LogManager.Instance.LogInfo($"规范架构树已{(shouldExpand ? "展开" : "折叠")}。");
        }

        private async void 刷新规范_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                StandardManagementTreeClientResponse response = await _standardManagementApiService
                    .GetManagementTreeAsync()
                    .ConfigureAwait(true);

                if (!response.Success)
                {
                    MessageBox.Show(response.Message, "规范目录查询失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                BuildStandardTree(response);
                MessageBox.Show($"规范架构树刷新完成：专业/类别 {response.Categories.Count} 个，规范系列 {response.Series.Count} 个。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"刷新规范架构树失败：{ex.Message}");
                MessageBox.Show($"刷新规范架构树失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void 导入规范_Btn_Click(object sender, RoutedEventArgs e)
        {
            if (!IsAdminUser(VariableDictionary._userName ?? string.Empty))
            {
                MessageBox.Show("只有管理员可以导入规范。", "权限不足", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            using var dialog = new OpenFileDialog
            {
                Filter = "规范文件 (*.xlsx;*.json)|*.xlsx;*.json|Excel 文件 (*.xlsx)|*.xlsx|JSON 文件 (*.json)|*.json",
                Title = "选择要预览的规范文件",
                Multiselect = false
            };

            if (dialog.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                StandardImportPreviewClientResponse response = await _standardManagementApiService
                    .PreviewImportAsync(dialog.FileName, VariableDictionary._userName ?? string.Empty)
                    .ConfigureAwait(true);
                string batchText = string.IsNullOrWhiteSpace(response.BatchId) ? "未生成" : response.BatchId;
                MessageBoxResult confirm = MessageBox.Show(
                    $"文件预览完成。\n{response.Message}\n错误：{response.ErrorCount}\n警告：{response.WarningCount}\n批次：{batchText}\n\n是否确认导入？",
                    response.Success ? "规范预览成功" : "规范预览存在问题",
                    response.Success ? MessageBoxButton.YesNo : MessageBoxButton.OK,
                    response.Success ? MessageBoxImage.Question : MessageBoxImage.Warning);
                if (response.Success && confirm == MessageBoxResult.Yes)
                {
                    StandardImportCommitClientResponse commit = await _standardManagementApiService
                        .CommitImportAsync(response.BatchId, response.WarningCount > 0, VariableDictionary._userName ?? string.Empty)
                        .ConfigureAwait(true);
                    MessageBox.Show(
                        $"{commit.Message}\n导入数量：{commit.ImportedCount}\n警告：{commit.WarningCount}",
                        commit.Success ? "规范导入成功" : "规范导入失败",
                        MessageBoxButton.OK,
                        commit.Success ? MessageBoxImage.Information : MessageBoxImage.Warning);
                    if (commit.Success)
                        await RefreshStandardTreeAsync().ConfigureAwait(true);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"规范文件预览失败：{ex.Message}");
                MessageBox.Show($"规范文件预览失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void 导出规范_Btn_Click(object sender, RoutedEventArgs e)
        {
            if (!(SpecificationTreeView.SelectedItem is CategoryTreeNode selectedNode)
                || !(selectedNode.Data is StandardManagementSeriesClient series))
            {
                MessageBox.Show("请先选择一个规范系列。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                List<StandardDocumentVersionClient> versions = await _standardManagementApiService
                    .GetManagementVersionsAsync(series.Id)
                    .ConfigureAwait(true);
                StandardDocumentVersionClient? current = versions.FirstOrDefault(item => item.IsCurrent);
                if (current == null)
                {
                    MessageBox.Show("当前规范系列没有可导出的有效版本。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                List<StandardDocumentFileClient> files = await _standardManagementApiService
                    .GetManagementFilesAsync(current.Id)
                    .ConfigureAwait(true);
                StandardDocumentFileClient? file = files.FirstOrDefault();
                if (file == null)
                {
                    MessageBox.Show("当前规范版本没有可导出的附件。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                using var saveDialog = new System.Windows.Forms.SaveFileDialog
                {
                    FileName = file.OriginalFileName,
                    Filter = "所有文件 (*.*)|*.*",
                    Title = "导出规范附件"
                };
                if (saveDialog.ShowDialog() != DialogResult.OK) return;

                await _standardManagementApiService
                    .DownloadManagementFileAsync(file.Id, saveDialog.FileName)
                    .ConfigureAwait(true);
                MessageBox.Show("规范附件导出成功。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"查询规范导出版本失败：{ex.Message}");
                MessageBox.Show($"查询规范导出版本失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 根据服务器返回的专业/类别和规范系列构造规范树。
        /// </summary>
        private void BuildStandardTree(StandardManagementTreeClientResponse response)
        {
            _standardTreeNodes.Clear();
            var nodeMap = new Dictionary<long, CategoryTreeNode>();

            foreach (StandardManagementCategoryClient category in response.Categories.OrderBy(item => item.SortOrder).ThenBy(item => item.Id))
            {
                nodeMap[category.Id] = new CategoryTreeNode(
                    unchecked((int)category.Id),
                    category.Code,
                    string.IsNullOrWhiteSpace(category.Name) ? category.Code : category.Name,
                    category.ParentId.HasValue ? 1 : 0,
                    category.ParentId.HasValue ? unchecked((int)category.ParentId.Value) : 0,
                    category)
                {
                    IsExpanded = false
                };
            }

            foreach (CategoryTreeNode node in nodeMap.Values)
            {
                StandardManagementCategoryClient category = (StandardManagementCategoryClient)node.Data;
                if (category.ParentId.HasValue && nodeMap.TryGetValue(category.ParentId.Value, out CategoryTreeNode? parent))
                    parent.Children.Add(node);
                else
                    _standardTreeNodes.Add(node);
            }

            foreach (StandardManagementSeriesClient series in response.Series.OrderBy(item => item.SeriesName).ThenBy(item => item.Id))
            {
                var seriesNode = new CategoryTreeNode(
                    unchecked((int)series.Id),
                    series.SeriesCode,
                    $"{series.SeriesName} [{series.StandardNumber}]",
                    2,
                    series.CategoryId.HasValue ? unchecked((int)series.CategoryId.Value) : 0,
                    series);

                if (series.CategoryId.HasValue && nodeMap.TryGetValue(series.CategoryId.Value, out CategoryTreeNode? parent))
                    parent.Children.Add(seriesNode);
                else
                    _standardTreeNodes.Add(seriesNode);
            }

            SpecificationTreeView.ItemsSource = null;
            SpecificationTreeView.ItemsSource = _standardTreeNodes;
        }

        /// <summary>
        /// 递归设置规范树所有节点的展开状态。
        /// </summary>
        private static void SetStandardTreeExpanded(CategoryTreeNode node, bool isExpanded)
        {
            node.IsExpanded = isExpanded;
            foreach (CategoryTreeNode child in node.Children)
                SetStandardTreeExpanded(child, isExpanded);
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
