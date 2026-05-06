using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.UniFiedStandards;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Dapper;
using MySql.Data.MySqlClient;
using static Autodesk.AutoCAD.DatabaseServices.TextEditor;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;
using UserControl = System.Windows.Controls.UserControl;

namespace GB_NewCadPlus_IV
{
    /// <summary>
    /// LoginWindow.xaml 的交互逻辑
    /// </summary>
    public partial class LoginWindow : Window
    {
        /// <summary>
        /// 登录配置
        /// </summary>
        private readonly string _configPath;

        /// <summary>
        /// 如果登录成功且可以连接数据库，此属性由 LoginWindow 构造并返回给调用方（可能为 null 表示未能连接 DB）
        /// </summary>
        public DatabaseManager CreatedDatabaseManager { get; private set; }
        /// <summary>
        /// 登录窗口
        /// </summary>
        public LoginWindow()
        {
            InitializeComponent();
            _configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GB_NewCadPlus_IV", "login_config.json");// 配置文件路径
            LoadConfig();//加载配置
            Loaded += LoginWindow_Loaded;//注册窗口加载事件处理程序
        }
        /// <summary>
        /// 窗口加载事件处理程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void LoginWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // 如果没有输入服务器，则填默认（先不强制覆盖，下面会根据 DB 类型再次设置）
            if (string.IsNullOrWhiteSpace(TxtServerIP.Text))
            {
                TxtServerIP.Text = "127.0.0.1"; // 默认本机 IP
            }

            // 先尝试从 UI 或配置读取数据库类型，优先 UI 选择
            string selectedDb = "DM"; // 默认使用达梦
            try
            {
                if (CmbDatabaseType != null && CmbDatabaseType.SelectedItem is ComboBoxItem cbi && cbi.Content is string s)// 尝试从 UI 读取数据库类型
                    selectedDb = s.ToUpper().Trim();// 标准化为大写并去除空白
                else if (!string.IsNullOrWhiteSpace(VariableDictionary._databaseType))// 如果 UI 没有选择但全局变量中有配置，则使用全局变量（例如从配置文件加载时）
                    selectedDb = VariableDictionary._databaseType.ToUpper().Trim(); // 标准化为大写并去除空白
            }
            catch { /* 忽略读取失败 */ }

            // 根据数据库类型设置端口与默认用户名（只在用户未填写时才覆盖）
            if (selectedDb == "MYSQL")
            {
                // MySQL 常用端口，只有在端口输入为空或为达梦默认时才替换为 MySQL 默认
                if (string.IsNullOrWhiteSpace(TxtServerPort.Text))
                    TxtServerPort.Text = "3306"; // MySQL 常见端口，项目中曾用 3308 可按需修改
                if (string.IsNullOrWhiteSpace(TxtUsername.Text))
                    TxtUsername.Text = "root"; // MySQL 管理用户，生产请替换为低权限用户
                if (string.IsNullOrWhiteSpace(PwdBox.Password))
                    PwdBox.Password="123456"; // MySQL 管理用户默认密码，生产请替换为实际密码或使用安全输入方式
            }
            else
            {
                // 达梦默认端口与用户名
                if (string.IsNullOrWhiteSpace(TxtServerPort.Text))
                    TxtServerPort.Text = "5236";
                if (string.IsNullOrWhiteSpace(TxtUsername.Text))
                    TxtUsername.Text = "SYSDBA";
                if (string.IsNullOrWhiteSpace(PwdBox.Password))
                    PwdBox.Password = "675756SGBsgb";
            }

            // 更新 UI 状态提示包含数据库类型信息，帮助排查
            TxtStatus.Text = $"正在检测服务器连接... (数据库类型: {selectedDb})";

            // 把选择写入全局变量，后续异步任务会读取这些值
            VariableDictionary._databaseType = selectedDb;
            VariableDictionary._serverIP = TxtServerIP.Text.Trim();
            VariableDictionary._serverPort = int.TryParse(TxtServerPort.Text.Trim(), out int port) ? port : (selectedDb == "MYSQL" ? 3306 : 5236);
            VariableDictionary._userName = TxtUsername.Text.Trim(); // UI 上的用户名（可能是管理账号用于初始化）
            VariableDictionary._passWord = PwdBox.Password; // UI 上的密码

            // 快速 TCP 层连通性检测，使用当前全局端口
            bool tcpOk = await Task.Run(() => TestNetworkConnection(VariableDictionary._serverIP, VariableDictionary._serverPort));
            if (!tcpOk)
            {
                // 首次尝试失败：提示用户填写有效服务器IP/端口（保留端口提示）
                TxtStatus.Text = $"无法连接到服务器 {VariableDictionary._serverIP}:{VariableDictionary._serverPort}，请在上方输入正确的服务器IP/端口后点击“保存服务器\\端口”。";
                TxtServerPort.Text = VariableDictionary._databaseType == "MYSQL" ? "3306" : "5236";
                TxtServerIP.Focus();
                CmbDepartments.ItemsSource = null;
                return;
            }

            // TCP 可达后再尝试从对应数据库读取部门（TryLoadDepartmentsAsync 已支持 DM 与 MySQL）
            var loaded = await TryLoadDepartmentsAsync(VariableDictionary._serverIP, VariableDictionary._serverPort);
            if (!loaded)
            {
                // 将失败原因显示在状态栏，提示用户检查 DB 类型/凭据
                TxtStatus.Text = $"服务器可达，但从 {VariableDictionary._databaseType} 数据库读取或初始化部门失败，请检查数据库或凭据，或在设置中修改服务器信息。";
                CmbDepartments.ItemsSource = null;
            }
            else
            {
                // 成功加载部门后提示并确保 UI 显示同步的 DB 类型
                TxtStatus.Text = $"已连接 {VariableDictionary._databaseType} 并加载部门。";
            }
        }

        /// <summary>
        /// 尝试使用 DMAuthService 读取部门并填充下拉框，返回是否成功
        /// </summary>
        private async Task<bool> TryLoadDepartmentsAsync(string host, int port)
        {
            try
            {
                // 在切换到后台线程之前，先读取需要的 UI 值，避免后台线程直接访问 UI 控件引发跨线程异常
                // 读取用户名和密码并做默认处理（在 UI 线程读取安全）
                //var uiUser = string.IsNullOrWhiteSpace(TxtUsername.Text) ? "SYSDBA" : TxtUsername.Text.Trim(); // 用户名（在后台使用）
                //var uiPwd = string.IsNullOrWhiteSpace(PwdBox.Password) ? "SYSDBA" : PwdBox.Password; // 密码（在后台使用）


                var uiUser = "SYSDBA" ; // 用户名（在后台使用）
                var uiPwd = "675756SGBsgb" ; // 密码（在后台使用）

                try
                {
                    // 使用在 UI 线程捕获的用户名/密码创建服务实例，避免直接访问控件
                    // 根据当前选择的数据库类型决定使用 DM 或 MySQL 的 AuthService
                    var selectedDb = VariableDictionary._databaseType ?? "DM";
                    try
                    {
                        if (CmbDatabaseType != null && CmbDatabaseType.SelectedItem is ComboBoxItem cbi && cbi.Content is string s)
                            selectedDb = s.ToUpper().Trim();
                    }
                    catch { }

                    List<DepartmentModel> depts = null;
                    if (selectedDb == "MYSQL")
                    {
                        // 使用 MySqlAuthService，并传入 UI 提供的用户名/密码（例如 sa）
                        var mySvc = new MySqlAuthService(host, port.ToString(), uiUser, uiPwd);
                        try
                        {
                            mySvc.EnsureAllTablesExist();
                            try { mySvc.SyncDepartmentsFromCadCategories(); } catch (Exception exSync) { LogManager.Instance.LogInfo($"MySql SyncDepartmentsFromCadCategories 失败: {exSync.Message}"); }
                            depts = mySvc.GetDepartmentsWithCounts();
                        }
                        catch (Exception exMyAuth)
                        {
                            LogManager.Instance.LogInfo($"使用 MySqlAuthService 读取部门失败 host={host}, port={port}, user={uiUser}: {exMyAuth}");
                            throw;
                        }
                    }
                    else
                    {
                        var svc = new DMAuthService(host, port.ToString(), uiUser, uiPwd);// 使用 DMAuthService，并传入 UI 提供的用户名/密码（例如 SYSDBA）

                        svc.EnsureAllTablesExist();// 确保表存在（如果连接成功但表不存在会抛异常）
                        try
                        {
                            svc.SyncDepartmentsFromCadCategories();// 尝试同步部门数据（如果 CAD_CATEGORIES 表存在但 DEPARTMENTS 表未正确同步可能会抛异常）
                        }
                        catch (Exception exSync)
                        {
                            LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories 失败: {exSync.Message}");
                        }

                        depts = svc.GetDepartmentsWithCounts();
                    }

                    LogManager.Instance.LogInfo($"读取到部门数量: {depts?.Count ?? 0}");
                    LogManager.Instance.LogInfo($"读取到部门数量: {depts?.Count ?? 0}");
                    // 记录获取到的部门数及其名称，便于排查为何 UI 未显示
                    LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync: retrieved departments count={depts?.Count ?? 0} from {host}:{port} user={uiUser}");
                    // 仅记录数量，避免在此处使用匿名类型或可能的类型不一致导致编译问题
                    LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync: departments list exists={depts != null}, count={depts?.Count ?? 0}");

                    // 将结果回到 UI 线程更新控件（直接使用 DepartmentModel 列表，避免匿名类型导致的问题）
                    // 将结果回到 UI 线程更新控件
                    Dispatcher.Invoke(() =>
                    {
                        if (depts != null && depts.Count > 0)
                        {
                            // 绑定整个对象列表
                            CmbDepartments.ItemsSource = depts;
                            // 关键：必须与 DepartmentModel 的属性名一致，且与 XAML 中的 DisplayMemberPath 一致
                            CmbDepartments.DisplayMemberPath = "DisplayName";
                            CmbDepartments.SelectedValuePath = "Id";
                            CmbDepartments.SelectedIndex = 0;
                            LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync: UI绑定完成，ItemsSourceCount={depts.Count}，ComboBoxItemsCount={CmbDepartments.Items.Count}，SelectedIndex={CmbDepartments.SelectedIndex}");

                            if (CmbDepartments.SelectedItem is DepartmentModel selectedDept)
                            {
                                LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync: 当前选中部门 Id={selectedDept.Id}, Name={selectedDept.Name}, DisplayName={selectedDept.DisplayName}");
                            }
                            else
                            {
                                LogManager.Instance.LogInfo("TryLoadDepartmentsAsync: 当前 SelectedItem 不是 DepartmentModel 或为空。");
                            }
                        }
                        else
                        {
                            CmbDepartments.ItemsSource = null;
                            LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync: UI绑定完成但无数据，ComboBoxItemsCount={CmbDepartments.Items.Count}");
                            LogManager.Instance.LogInfo("警告：部门列表为空，请检查后台 CAD_CATEGORIES 数据是否已成功同步至 DEPARTMENTS。");
                        }
                    });
                    return true;
                }
                catch (Exception ex)
                {
                    // 这里不能访问 TxtUsername.Text（可能导致跨线程异常），改用在外层捕获的 uiUser 变量记录日志
                    LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync 失败 host={host}, port={port}, user={uiUser}: {ex}");
                    Dispatcher.Invoke(() =>
                    {
                        TxtStatus.Text = "读取或初始化部门失败：" + ex.Message;
                        CmbDepartments.ItemsSource = null;
                    });
                    return false;
                }

                //return await Task.Run(() =>
                //{
                //    try
                //    {
                //        // 使用在 UI 线程捕获的用户名/密码创建服务实例，避免直接访问控件
                //        // 根据当前选择的数据库类型决定使用 DM 或 MySQL 的 AuthService
                //        var selectedDb = VariableDictionary._databaseType ?? "DM";
                //        try
                //        {
                //            if (CmbDatabaseType != null && CmbDatabaseType.SelectedItem is ComboBoxItem cbi && cbi.Content is string s)
                //                selectedDb = s.ToUpper().Trim();
                //        }
                //        catch { }

                //        List<DepartmentModel> depts = null;
                //        if (selectedDb == "MYSQL")
                //        {
                //            // 使用 MySqlAuthService，并传入 UI 提供的用户名/密码（例如 sa）
                //            var mySvc = new MySqlAuthService(host, port.ToString(), uiUser, uiPwd);
                //            try
                //            {
                //                mySvc.EnsureAllTablesExist();
                //                try { mySvc.SyncDepartmentsFromCadCategories(); } catch (Exception exSync) { LogManager.Instance.LogInfo($"MySql SyncDepartmentsFromCadCategories 失败: {exSync.Message}"); }
                //                depts = mySvc.GetDepartmentsWithCounts();
                //            }
                //            catch (Exception exMyAuth)
                //            {
                //                LogManager.Instance.LogInfo($"使用 MySqlAuthService 读取部门失败 host={host}, port={port}, user={uiUser}: {exMyAuth}");
                //                throw;
                //            }
                //        }
                //        else
                //        {
                //            var svc = new DMAuthService(host, port.ToString(), uiUser, uiPwd);// 使用 DMAuthService，并传入 UI 提供的用户名/密码（例如 SYSDBA）

                //            svc.EnsureAllTablesExist();// 确保表存在（如果连接成功但表不存在会抛异常）
                //            try
                //            {
                //                svc.SyncDepartmentsFromCadCategories();// 尝试同步部门数据（如果 CAD_CATEGORIES 表存在但 DEPARTMENTS 表未正确同步可能会抛异常）
                //            }
                //            catch (Exception exSync)
                //            {
                //                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories 失败: {exSync.Message}");
                //            }

                //            depts = svc.GetDepartmentsWithCounts();
                //        }

                //        LogManager.Instance.LogInfo($"读取到部门数量: {depts?.Count ?? 0}");
                //        LogManager.Instance.LogInfo($"读取到部门数量: {depts?.Count ?? 0}");
                //        // 记录获取到的部门数及其名称，便于排查为何 UI 未显示
                //        LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync: retrieved departments count={depts?.Count ?? 0} from {host}:{port} user={uiUser}");
                //        // 仅记录数量，避免在此处使用匿名类型或可能的类型不一致导致编译问题
                //        LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync: departments list exists={depts != null}, count={depts?.Count ?? 0}");

                //        // 将结果回到 UI 线程更新控件（直接使用 DepartmentModel 列表，避免匿名类型导致的问题）
                //        // 将结果回到 UI 线程更新控件
                //        Dispatcher.Invoke(() =>
                //        {
                //            if (depts != null && depts.Count > 0)
                //            {
                //                // 绑定整个对象列表
                //                CmbDepartments.ItemsSource = depts;
                //                // 关键：必须与 DepartmentModel 的属性名一致，且与 XAML 中的 DisplayMemberPath 一致
                //                CmbDepartments.DisplayMemberPath = "DisplayName";
                //                CmbDepartments.SelectedValuePath = "Id";
                //                CmbDepartments.SelectedIndex = 0;
                //                LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync: UI绑定完成，ItemsSourceCount={depts.Count}，ComboBoxItemsCount={CmbDepartments.Items.Count}，SelectedIndex={CmbDepartments.SelectedIndex}");

                //                if (CmbDepartments.SelectedItem is DepartmentModel selectedDept)
                //                {
                //                    LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync: 当前选中部门 Id={selectedDept.Id}, Name={selectedDept.Name}, DisplayName={selectedDept.DisplayName}");
                //                }
                //                else
                //                {
                //                    LogManager.Instance.LogInfo("TryLoadDepartmentsAsync: 当前 SelectedItem 不是 DepartmentModel 或为空。");
                //                }
                //            }
                //            else
                //            {
                //                CmbDepartments.ItemsSource = null;
                //                LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync: UI绑定完成但无数据，ComboBoxItemsCount={CmbDepartments.Items.Count}");
                //                LogManager.Instance.LogInfo("警告：部门列表为空，请检查后台 CAD_CATEGORIES 数据是否已成功同步至 DEPARTMENTS。");
                //            }
                //        });
                //        return true;
                //    }
                //    catch (Exception ex)
                //    {
                //        // 这里不能访问 TxtUsername.Text（可能导致跨线程异常），改用在外层捕获的 uiUser 变量记录日志
                //        LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync 失败 host={host}, port={port}, user={uiUser}: {ex}");
                //        Dispatcher.Invoke(() =>
                //        {
                //            TxtStatus.Text = "读取或初始化部门失败：" + ex.Message;
                //            CmbDepartments.ItemsSource = null;
                //        });
                //        return false;
                //    }
                //});
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync 外层异常 host={host}, port={port}: {ex}");
                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = "加载部门时发生错误：" + ex.Message;
                    CmbDepartments.ItemsSource = null;
                });
                return false;
            }
        }

        /// <summary>
        /// 加载登录配置
        /// </summary>
        private void LoadConfig()
        {
            try
            {
                if (!File.Exists(_configPath)) return;// 配置文件不存在则返回
                var json = File.ReadAllText(_configPath);//读取配置文件内容
                var ser = new JavaScriptSerializer();//创建JSON序列化器 创建序列化器
                var cfg = ser.Deserialize<LoginConfig>(json);//反序列化JSON反序列化为LoginConfig对象
                if (cfg == null) return;//配置文件为空则返回 配置为空则返回
                TxtServerIP.Text = cfg.ServerIP ?? "";//若配置存在则填入（不要覆盖为127.0.0.1，这里让 Loaded 处理默认）
                TxtServerPort.Text = cfg.ServerPort ?? "";//同上
                TxtUsername.Text = cfg.Username ?? "";//设置用户名
                if (cfg.EncryptedPassword != null && cfg.SavePassword)//保存密码如果保存了密码
                {
                    var pwd = ProtectedData.Unprotect(Convert.FromBase64String(cfg.EncryptedPassword), null, DataProtectionScope.CurrentUser);//解密密码
                    PwdBox.Password = Encoding.UTF8.GetString(pwd);//设置密码
                    ChkSavePassword.IsChecked = true;//勾选保存密码
                }
                VariableDictionary._serverIP = TxtServerIP.Text.Trim();
                VariableDictionary._serverPort = int.TryParse(TxtServerPort.Text.Trim(), out int port) ? port : 5236;
                VariableDictionary._userName = TxtUsername.Text.Trim();
                VariableDictionary._passWord = PwdBox.Password.Trim();
                // 恢复配置中的数据库类型（如果存在）并同步到 UI
                try
                {
                    if (cfg != null && !string.IsNullOrWhiteSpace(cfg.DatabaseType))
                    {
                        VariableDictionary._databaseType = cfg.DatabaseType.ToUpper().Trim();
                        if (CmbDatabaseType != null)
                        {
                            foreach (var item in CmbDatabaseType.Items)
                            {
                                if (item is ComboBoxItem cbi && cbi.Content is string s && s.Equals(VariableDictionary._databaseType, StringComparison.OrdinalIgnoreCase))
                                {
                                    CmbDatabaseType.SelectedItem = item;
                                    break;
                                }
                            }
                        }
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "加载配置失败：" + ex.Message;
                MessageBox.Show("加载配置失败：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// 保存登录配置
        /// </summary>
        /// <param name="savePassword"></param>
        private void SaveConfig(bool savePassword)
        {
            try
            {
                var dir = Path.GetDirectoryName(_configPath);//获取配置文件目录
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);//创建目录
                //创建登录配置对象
                var cfg = new LoginConfig
                {
                    ServerIP = TxtServerIP.Text,//设置服务器IP
                    ServerPort = TxtServerPort.Text,//设置服务端口
                    Username = TxtUsername.Text,//设置用户名
                    SavePassword = savePassword//保存密码
                };

                // 如果 UI 中存在数据库类型选择，则保存到配置
                try
                {
                    if (CmbDatabaseType != null && CmbDatabaseType.SelectedItem is ComboBoxItem sel)
                    {
                        var selStr = sel.Content as string;
                        if (!string.IsNullOrWhiteSpace(selStr))
                        {
                            // 使用反射设置 LoginConfig 上可能不存在的属性，保持向后兼容
                            var dbTypeProp = cfg.GetType().GetProperty("DatabaseType");
                            if (dbTypeProp != null)
                            {
                                dbTypeProp.SetValue(cfg, selStr.ToUpper());
                            }
                            else
                            {
                                // 如果 LoginConfig 没有该属性，则通过序列化前替换 JSON 文本方式持久化
                                // 这里先将 VariableDictionary 设置好，并在写文件后手动追加字段
                                VariableDictionary._databaseType = selStr.ToUpper();
                            }
                        }
                    }
                }
                catch { }

                if (savePassword)//保存密码
                {
                    var bytes = Encoding.UTF8.GetBytes(PwdBox.Password ?? "");//获取密码字节数组
                    var protectedBytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);//保护密码
                    cfg.EncryptedPassword = Convert.ToBase64String(protectedBytes);//保存加密后的密码
                }
                VariableDictionary._serverIP = TxtServerIP.Text.Trim();
                VariableDictionary._serverPort = int.TryParse(TxtServerPort.Text.Trim(), out int port) ? port : 5236;
                VariableDictionary._userName = TxtUsername.Text.Trim();
                VariableDictionary._passWord = PwdBox.Password.Trim();
                var ser = new JavaScriptSerializer();//创建JSON序列化器
                var json = ser.Serialize(cfg);
                // 如果 LoginConfig 类型没有 DatabaseType 字段，但 VariableDictionary._databaseType 已设置，则追加该字段到 JSON（向后兼容）
                try
                {
                    var dbTypeProp = cfg.GetType().GetProperty("DatabaseType");
                    if (dbTypeProp == null && !string.IsNullOrWhiteSpace(VariableDictionary._databaseType))
                    {
                        // 简单方式：在结尾前插入字段（假设 ser.Serialize 产出一个对象 JSON）
                        if (json.TrimEnd().EndsWith("}"))
                        {
                            var insert = $",\n  \"DatabaseType\": \"{VariableDictionary._databaseType}\"\n";
                            json = json.TrimEnd();
                            json = json.Substring(0, json.Length - 1) + insert + "}";
                        }
                    }
                }
                catch { }

                File.WriteAllText(_configPath, json);//保存配置序列化并写入配置文件
                TxtStatus.Text = "配置已保存。";//显示保存成功消息
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "保存配置失败：" + ex.Message;//显示保存失败消息
            }
        }
        /// <summary>
        /// 登录按钮点击事件处理程序
        /// 在登录成功后：1) 保存登录配置；2) 尝试创建 DatabaseManager 并赋值 CreatedDatabaseManager；3) 关闭窗口返回 DialogResult=true
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            // 在开始登录前，先确认为哪个数据库类型进行认证
            try
            {
                if (CmbDatabaseType != null && CmbDatabaseType.SelectedItem is ComboBoxItem selType)
                {
                    var selStr = selType.Content as string;
                    if (!string.IsNullOrWhiteSpace(selStr))
                    {
                        VariableDictionary._databaseType = selStr.ToUpper().Trim(); // 更新全局数据库类型标识
                    }
                }
            }
            catch { }
            VariableDictionary._serverIP = TxtServerIP.Text.Trim();
            VariableDictionary._serverPort = int.TryParse(TxtServerPort.Text.Trim(), out int port) ? port : 5236;
            VariableDictionary._userName = TxtUsername.Text.Trim();
            VariableDictionary._passWord = PwdBox.Password.Trim();

            BtnLogin.IsEnabled = false;
            TxtStatus.Text = "正在连接并验证用户...";
            // 1) 先做快速 TCP 连通性检测；失败则直接退回 FormMain
            bool tcpOk = await Task.Run(() => TestNetworkConnection(VariableDictionary._serverIP, VariableDictionary._serverPort));
            if (!tcpOk)
            {
                TxtStatus.Text = "无法连接服务器，将进入本地工具界面。";
                try
                {
                    Command.gfff(); // 显示 FormMain
                }
                finally
                {
                    DialogResult = false;
                    Close();
                    BtnLogin.IsEnabled = true;
                }
                return;
            }
            try
            {
                var authOk = await Task.Run(() =>
                {
                    try
                    {
                        // 判定是否为 MySQL 模式
                        bool isMySql = VariableDictionary._databaseType == "MYSQL";

                        if (isMySql)
                        {
                            // --- MySQL 分支 ---
                            // 使用数据库管理凭据（root）初始化服务，确保具有建表和查询系统表的权限
                            // 注意：如果您环境中的 MySQL root 密码不同，请调整此处或从配置读取
                            var svc = new MySqlAuthService(
                                VariableDictionary._serverIP,
                                VariableDictionary._serverPort.ToString(),
                                "root",
                                "123456"
                            );
                            svc.EnsureAllTablesExist(); // 确保业务表结构完整
                            // 最终使用用户输入的账号密码在 USERS 表中进行业务身份认证
                            return svc.AuthenticateUser(VariableDictionary._userName, VariableDictionary._passWord);
                        }
                        else
                        {
                            // --- 达梦 (DM) 分支 ---
                            // 修正 root cause：使用 SYSDBA 管理员账号建立物理连接，解决 6001 用户名错误
                            var svc = new DMAuthService(
                                VariableDictionary._serverIP,
                                VariableDictionary._serverPort.ToString(),
                                "SYSDBA",
                                "675756SGBsgb"
                            );
                            svc.EnsureAllTablesExist(); // 强制初始化或检查表结构
                            // 在物理连接成功的基础上，通过 SQL 查询验证应用用户身份
                            return svc.AuthenticateUser(VariableDictionary._userName, VariableDictionary._passWord);
                        }
                    }
                    catch (Exception exInner)
                    {
                        LogManager.Instance.LogInfo($"后台认证出错: {exInner.Message}");
                        throw;
                    }
                });

                if (authOk)
                {
                    // 在用户确认登录前，从 UI 读取数据库类型选择并写入全局变量，确保 DatabaseManager 在构造时使用正确适配器
                    try
                    {
                        if (CmbDatabaseType != null && CmbDatabaseType.SelectedItem is ComboBoxItem sel)
                        {
                            var selStr = sel.Content as string;
                            if (!string.IsNullOrWhiteSpace(selStr))
                            {
                                VariableDictionary._databaseType = selStr.ToUpper().Trim();
                            }
                        }
                    }
                    catch { }

                    if (ChkSavePassword.IsChecked == true)
                        SaveConfig(true);
                    else
                        SaveConfig(false);

                    TxtStatus.Text = "登录成功。";

                    try
                    {
                        var db = await Task.Run(() =>
                        {
                            try
                            {
                                string dbPart = string.IsNullOrWhiteSpace(VariableDictionary._dataBaseName)
                                    ? string.Empty
                                    : $"Schema={VariableDictionary._dataBaseName};";
                                VariableDictionary._newConnectionString =
                                    $"Server={VariableDictionary._serverIP};Port={VariableDictionary._serverPort};{dbPart}User Id={VariableDictionary._userName};Password={VariableDictionary._passWord};";

                                return new DatabaseManager(VariableDictionary._newConnectionString);
                            }
                            catch (Exception exDbCreate)
                            {
                                LogManager.Instance.LogInfo($"后台构造 DatabaseManager 失败: {exDbCreate.Message}");
                                return null;
                            }
                        });

                        if (db != null && db.IsDatabaseAvailable)
                        {
                            var ensureOk = await db.CreateLayerDictionaryTableIfNotExistsAsync();
                            if (!ensureOk)
                                LogManager.Instance.LogInfo("确保 layer_dictionary 表失败（但已继续登录）。");

                            CreatedDatabaseManager = db;
                            TxtStatus.Text += " 已连接数据库。";
                        }
                        else
                        {
                            CreatedDatabaseManager = null;
                            TxtStatus.Text += " 但未能连接数据库（请在设置中检查数据库凭据）。";
                        }
                    }
                    catch (Exception exDb)
                    {
                        CreatedDatabaseManager = null;
                        TxtStatus.Text += " 创建 DatabaseManager 时出错：" + exDb.Message;
                        LogManager.Instance.LogInfo($"创建 DatabaseManager 出错: {exDb.Message}");
                    }

                    // 登录成功后关闭窗口（在 UI 线程）
                    DialogResult = true;
                    Close();
                }
                else
                {
                    // 认证失败：在 UI 线程交互
                    var res = MessageBox.Show("用户不存在或密码错误。是否注册新用户？", "登录失败", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (res == MessageBoxResult.Yes)
                    {
                        var deptList = new List<(int Id, string Name)>();
                        try
                        {
                            if (CmbDepartments.ItemsSource != null)
                            {
                                foreach (var item in CmbDepartments.ItemsSource)
                                {
                                    var t = item.GetType();
                                    var pid = t.GetProperty("Id");
                                    var pname = t.GetProperty("Name");
                                    if (pid != null && pname != null)
                                    {
                                        int id = Convert.ToInt32(pid.GetValue(item));
                                        string name = Convert.ToString(pname.GetValue(item));
                                        deptList.Add((id, name));
                                    }
                                }
                            }
                        }
                        catch { }

                        var regWin = new RegisterUserWindow(VariableDictionary._serverIP, VariableDictionary._serverPort, deptList) { Owner = this };
                        var regRes = regWin.ShowDialog();
                        if (regRes == true && regWin.RegistrationSucceeded)
                            TxtStatus.Text = "注册成功，请使用新用户登录。";
                        else
                            TxtStatus.Text = "未注册。";
                    }
                }
            }
            catch (Exception ex)
            {
                // 捕获后台任务抛出的异常（包括跨线程访问异常）
                TxtStatus.Text = "登录过程中出现异常：" + ex.Message;
                LogManager.Instance.LogInfo($"BtnLogin_Click 异常: {ex.Message}");
            }
            finally
            {
                BtnLogin.IsEnabled = true;
            }
        }
        /// <summary>
        /// 取消按钮点击事件处理程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;//设置对话框结果为 false
            Close();//关闭登录窗口
        }
        /// <summary>
        /// 保存服务器按钮点击事件处理程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnSaveServer_Click(object sender, RoutedEventArgs e)
        {
            // 在保存配置并尝试连接前，确保把 UI 上的数据库类型选择写入全局变量
            try
            {
                if (CmbDatabaseType != null && CmbDatabaseType.SelectedItem is ComboBoxItem sel && sel.Content is string s)
                {
                    VariableDictionary._databaseType = s.ToUpper().Trim();
                }
            }
            catch { }

            SaveConfig(ChkSavePassword.IsChecked == true);//保存登录配置

            // 尝试用新配置连接并加载部门
            VariableDictionary._serverIP = TxtServerIP.Text.Trim();
            if (string.IsNullOrWhiteSpace(VariableDictionary._serverIP))
            {
                MessageBox.Show("请填写服务器IP地址。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                TxtServerIP.Focus();
                return;
            }

            if (!int.TryParse(TxtServerPort.Text.Trim(), out int port))
            {
                port = 5236;
                TxtServerPort.Text = "5236";
            }
            VariableDictionary._serverPort = port;
            TxtStatus.Text = "正在连接服务器...";

            // 先做 TCP 层检测，快速反馈
            bool tcpOk = await Task.Run(() => TestNetworkConnection(VariableDictionary._serverIP, VariableDictionary._serverPort));
            if (!tcpOk)
            {
                MessageBox.Show($"无法连接到服务器 {VariableDictionary._serverIP}:{VariableDictionary._serverPort}，请检查IP、端口或网络。", "连接失败", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtStatus.Text = "连接失败，请检查服务器IP或端口。";
                return;
            }

            // TCP 成功后尝试读取并初始化部门（TryLoadDepartmentsAsync 已包含初始化与同步）
            var loaded = await TryLoadDepartmentsAsync(VariableDictionary._serverIP, VariableDictionary._serverPort);
            if (loaded)
            {
                TxtStatus.Text = "服务器连接成功，部门已加载。";
                MessageBox.Show("服务器连接成功并加载部门。", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                var detail = TxtStatus.Text;
                if (string.IsNullOrWhiteSpace(detail))
                {
                    detail = "连接服务器成功，但读取部门失败，请检查数据库配置或凭据。";
                }
                TxtStatus.Text = detail;
                MessageBox.Show($"连接服务器成功，但读取部门失败。\n详细信息：{detail}", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

        }
        /// <summary>
        /// 忘记密码按钮点击事件处理程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ForgetPassword_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("请联系管理员重置密码，或在注册界面重新创建账号。", "忘记密码", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        /// <summary>
        /// 端口输入框预处理事件处理程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtServerPort_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // 只允许数字
            foreach (var ch in e.Text)
            {
                if (!char.IsDigit(ch)) { e.Handled = true; break; }
            }
        }

        /// <summary>
        /// 当数据库类型切换时，自动为 MySQL 填充常见默认值（不覆盖用户已有输入）
        /// 保持达梦相关逻辑不变
        /// </summary>
        private void CmbDatabaseType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (CmbDatabaseType.SelectedItem is ComboBoxItem cbi && cbi.Content is string s)
                {
                    var sel = s.ToUpper().Trim();
                    // 仅在用户还没有填写用户名/端口时才自动填充
                    if (sel == "MYSQL")
                    {
                        if (string.IsNullOrWhiteSpace(TxtServerPort.Text))
                            TxtServerPort.Text = "3308"; // 你的 MySQL 端口示例
                        if (string.IsNullOrWhiteSpace(TxtUsername.Text))
                            TxtUsername.Text = "sa"; // 常见测试用户（请在生产中替换）
                        // 不自动设置密码，避免写入明文
                    }
                    else
                    {
                        // DM 默认端口为 5236
                        if (string.IsNullOrWhiteSpace(TxtServerPort.Text))
                            TxtServerPort.Text = "5236";
                        if (string.IsNullOrWhiteSpace(TxtUsername.Text))
                            TxtUsername.Text = "SYSDBA";
                    }
                    // 将选择保存到全局变量，供后续构造 DatabaseManager 使用
                    VariableDictionary._databaseType = sel;
                }
            }
            catch { }
        }

        /// <summary>
        /// 快速 TCP 连通检测（保留原有实现）TestNetworkConnection
        /// </summary>
        /// <param name="host"></param>
        /// <param name="port"></param>
        /// <param name="timeoutMs"></param>   
        /// <returns></returns>
        public static bool TestNetworkConnection(string host, int port, int timeoutMs = 5000)
        {
            try
            {
                using (var client = new System.Net.Sockets.TcpClient())//创建TcpClient对象
                {
                    var result = client.BeginConnect(host, port, null, null);//开始异步连接
                    var success = result.AsyncWaitHandle.WaitOne(TimeSpan.FromMilliseconds(timeoutMs));//等待连接完成或超时
                    client.EndConnect(result);//结束异步连接
                    return success;//返回连接成功或失败
                }
            }
            catch
            {
                return false;
            }
        }

        private async void BtnTestServer测试服务器_Click(object sender, RoutedEventArgs e)
        {
            // 禁用按钮，避免重复点击
            BtnTestServer测试服务器.IsEnabled = false;
            TxtStatus.Text = "正在测试服务器连接并读取数据库架构...";
            //var user = "";
            //var pwd = "";
            try
            {
                // 根据所选数据库类型执行不同测试
                string selectedDb = VariableDictionary._databaseType ?? "DM";
                try
                {
                    if (CmbDatabaseType != null && CmbDatabaseType.SelectedItem is ComboBoxItem cbi && cbi.Content is string s)
                        selectedDb = s.ToUpper().Trim();
                }
                catch { }

                if (selectedDb == "MYSQL")
                {
                    // 测试 MySQL 连接
                    var server = TxtServerIP.Text.Trim();
                    var port = TxtServerPort.Text.Trim();
                    // 优先使用 UI 中填写的用户名，否则回退到 VariableDictionary 中可能已保存的用户名
                    //var user = string.IsNullOrWhiteSpace(TxtUsername.Text) ? (VariableDictionary._userName ?? string.Empty) : TxtUsername.Text.Trim();
                    //var pwd = PwdBox.Password.Trim();
                    // 记录用于测试的目标信息（不记录明文密码）
                    LogManager.Instance.LogInfo($"测试 MySQL 连接: {server}:{port} user = root ");
                    string dbPart = string.IsNullOrWhiteSpace(VariableDictionary._dataBaseName) ? string.Empty : $"Database={VariableDictionary._dataBaseName};";
                    var connStr = $"Server={server};Port={port};{dbPart}User Id=root;Password=123456;";
                    try
                    {
                        using var conn = new MySqlConnection(connStr);
                        conn.Open();
                        // 简单查询服务器版本以验证可用性
                        using var cmd = conn.CreateCommand();
                        cmd.CommandText = "SELECT VERSION();";
                        var ver = cmd.ExecuteScalar();
                        TxtStatus.Text = $"MySQL 连接成功，版本: {ver}";
                    }
                    catch (MySql.Data.MySqlClient.MySqlException exMy)
                    {
                        // 常见情况：认证失败（Access denied）或网络/端口不可达
                        LogManager.Instance.LogError($"测试 MySQL 连接失败: {exMy}");
                        if (exMy.Message != null && exMy.Message.Contains("Access denied"))
                        {
                            TxtStatus.Text = "认证失败：请检查用户名/密码或用户权限（Access denied）。";
                            MessageBox.Show("MySQL 认证失败：请确认用户名和密码正确，且用户在目标主机/端口上具有登录权限。若是 localhost/127.0.0.1，请确保为对应主机创建了用户（例如 'sa'@'localhost' 与 'sa'@'127.0.0.1'）。", "连接失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        else
                        {
                            TxtStatus.Text = $"操作失败：{exMy.Message}";
                            MessageBox.Show($"测试 MySQL 连接失败：{exMy.Message}", "连接失败", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    catch (Exception exMy)
                    {
                        LogManager.Instance.LogError($"测试 MySQL 连接失败(通用异常): {exMy}");
                        TxtStatus.Text = $"操作失败：{exMy.Message}";
                    }
                }
                else
                {
                    // 构建参数数组用于 DMDatabaseReader
                    string[] args = new string[]
                    {
                        TxtServerIP.Text.Trim(), // 服务器地址
                        TxtServerPort.Text.Trim(), // 服务器端口
                        //TxtUsername.Text.Trim(), // 用户名
                        //PwdBox.Password.Trim() // 密码
                        "SYSDBA",
                        "675756SGBsgb"
                    };
                    
                    // 调用 DMDatabaseReaderMethod 方法
                    GB_NewCadPlus_IV.DMDatabaseReader.DMDatabaseReader.DMDatabaseReaderMethod(args);

                    TxtStatus.Text = "数据库架构读取完成，详细信息请查看控制台输出。";
                }
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"操作失败：{ex.Message}";
                Console.WriteLine($"错误：{ex.Message}");
            }
            finally
            {
                BtnTestServer测试服务器.IsEnabled = true; // 恢复按钮可用状态
            }
        }
     
    }
    /// <summary>
    /// 登录配置类
    /// </summary>
    internal class LoginConfig
    {
        public string ServerIP { get; set; }
        public string ServerPort { get; set; }
        public string Username { get; set; }
        public bool SavePassword { get; set; }
        public string EncryptedPassword { get; set; }
        // 可选：持久化数据库类型（"DM" 或 "MYSQL"），用于下次启动时恢复选项
        public string DatabaseType { get; set; }
    }
}
