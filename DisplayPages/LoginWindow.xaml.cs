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
            // 如果没有输入服务器，则填默认
            if (string.IsNullOrWhiteSpace(TxtServerIP.Text))
            {
                TxtServerIP.Text = "127.0.0.1";
            }
            if (string.IsNullOrWhiteSpace(TxtServerPort.Text))
            {
                TxtServerPort.Text = "3306";
            }

            TxtStatus.Text = "正在检测服务器连接...";

            VariableDictionary._serverIP = TxtServerIP.Text.Trim();
            VariableDictionary._serverPort = int.TryParse(TxtServerPort.Text.Trim(), out int port) ? port : 3306;
            VariableDictionary._userName = TxtUsername.Text;// 获取新的用户名
            VariableDictionary._passWord = PwdBox.Password;// 获取新的密码

            bool tcpOk = await Task.Run(() => TestNetworkConnection(VariableDictionary._serverIP, VariableDictionary._serverPort));
            if (!tcpOk)
            {
                // 首次尝试失败：提示用户填写有效服务器IP（端口默认保留为3306）
                TxtStatus.Text = "无法连接到默认服务器，请在上方输入正确的服务器IP后点击“保存服务器\\端口”。端口默认 3306。";
                TxtServerPort.Text = "3306";
                TxtServerIP.Focus();
                CmbDepartments.ItemsSource = null;
                return;
            }

            // TCP 可达后再尝试从数据库读取部门（GetDepartments），包含表初始化与同步
            var loaded = await TryLoadDepartmentsAsync(VariableDictionary._serverIP, VariableDictionary._serverPort);
            if (!loaded)
            {
                TxtStatus.Text = "服务器可达，但从数据库读取或初始化部门失败，请检查数据库或凭据，或在设置中修改服务器信息。";
                CmbDepartments.ItemsSource = null;
            }
        }

        /// <summary>
        /// 尝试使用 MySqlAuthService 读取部门并填充下拉框，返回是否成功  TryLoadDepartmentsAsync
        /// （同时执行必要的表存在性检查与同步）
        /// </summary>
        private async Task<bool> TryLoadDepartmentsAsync(string host, int port)
        {
            try
            {
                return await Task.Run(() =>
                {
                    try
                    {
                        string defaultDbUser = "root";
                        string defaultDbPwd = "root";
                        string dbName = "cad_sw_library";

                        // 1) 检查数据库/核心表缺失
                        var missing = DatabaseManager.CheckMissingCoreTables(host, port, defaultDbUser, defaultDbPwd, dbName);

                        if (missing != null && missing.Count > 0)
                        {
                            // 如果数据库不存在或表缺失，在 UI 线程提示用户是否初始化
                            var missingIndicator = string.Join(", ", missing);
                            bool shouldInit = false;
                            Dispatcher.Invoke(() =>
                            {
                                if (missing.Count == 1 && missing[0] == "__DATABASE_MISSING__")
                                {
                                    var res = MessageBox.Show($"目标数据库 `{dbName}` 不存在。是否在服务器 {host}:{port} 上创建数据库并初始化所需表？", "初始化数据库", MessageBoxButton.YesNo, MessageBoxImage.Question);
                                    shouldInit = (res == MessageBoxResult.Yes);
                                }
                                else
                                {
                                    var res = MessageBox.Show($"检测到缺失数据库表：{missingIndicator}。是否自动创建缺失表并初始化数据库结构？", "初始化表结构", MessageBoxButton.YesNo, MessageBoxImage.Question);
                                    shouldInit = (res == MessageBoxResult.Yes);
                                }
                            });

                            if (!shouldInit)
                            {
                                // 用户拒绝初始化，视为失败
                                Dispatcher.Invoke(() =>
                                {
                                    TxtStatus.Text = "未初始化数据库，无法加载部门。";
                                    CmbDepartments.ItemsSource = null;
                                });
                                return false;
                            }

                            // 用户同意初始化：尝试创建数据库与核心表
                            var created = DatabaseManager.CreateDatabaseAndCoreTables(host, port, defaultDbUser, defaultDbPwd, dbName);
                            if (!created)
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    TxtStatus.Text = "初始化数据库失败，请检查数据库服务器及凭据。";
                                    CmbDepartments.ItemsSource = null;
                                });
                                return false;
                            }
                        }

                        // 2) 数据库与表已准备好，连接并读取部门（使用 DatabaseManager）
                        var conn = $"Server={host};Port={port};Database={dbName};Uid={defaultDbUser};Pwd={defaultDbPwd};";
                        var svc = new MySqlAuthService(host, port.ToString());// 使用 MySqlAuthService
                        // 确保表存在并同步（幂等）
                        svc.EnsureDepartmentsTableExists();
                        svc.EnsureCategoriesTableExists();
                        try
                        {
                            svc.SyncDepartmentsFromCadCategories();
                        }
                        catch (Exception exSync)
                        {
                            LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories 失败: {exSync.Message}");
                        }

                        // 读取 departments 用于下拉（显示部门）
                        var depts = svc.GetDepartmentsWithCounts();
                        // 在 UI 线程更新下拉
                        Dispatcher.Invoke(() =>
                        {
                            CmbDepartments.ItemsSource = depts?.Select(d => new { Id = d.Id, Name = string.IsNullOrEmpty(d.DisplayName) ? d.Name : d.DisplayName }).ToList();
                            if (CmbDepartments.Items.Count > 0) CmbDepartments.SelectedIndex = 0;
                        });
                        return true;
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() =>// 在 UI 线程更新
                        {
                            TxtStatus.Text = "读取或初始化部门失败：" + ex.Message;// 更新状态文本
                            CmbDepartments.ItemsSource = null;// 清空部门列表
                        });
                        return false;
                    }
                });
            }
            catch (Exception ex)
            {
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
                VariableDictionary._serverPort = int.TryParse(TxtServerPort.Text.Trim(), out int port) ? port : 3306;
                VariableDictionary._userName = TxtUsername.Text.Trim();
                VariableDictionary._passWord = PwdBox.Password.Trim();
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

                if (savePassword)//保存密码
                {
                    var bytes = Encoding.UTF8.GetBytes(PwdBox.Password ?? "");//获取密码字节数组
                    var protectedBytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);//保护密码
                    cfg.EncryptedPassword = Convert.ToBase64String(protectedBytes);//保存加密后的密码
                }
                VariableDictionary._serverIP = TxtServerIP.Text.Trim();
                VariableDictionary._serverPort = int.TryParse(TxtServerPort.Text.Trim(), out int port) ? port : 3306;
                VariableDictionary._userName = TxtUsername.Text.Trim();
                VariableDictionary._passWord = PwdBox.Password.Trim();
                var ser = new JavaScriptSerializer();//创建JSON序列化器
                File.WriteAllText(_configPath, ser.Serialize(cfg));//保存配置序列化并写入配置文件
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
            VariableDictionary._serverIP = TxtServerIP.Text.Trim();
            VariableDictionary._serverPort = int.TryParse(TxtServerPort.Text.Trim(), out int port) ? port : 3306;
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
                // 在后台线程执行认证和与 DB 相关的耗时工作，避免在 UI 线程中直接调用可能会访问 UI 的库或阻塞 UI。
                var authOk = await Task.Run(() =>
                {
                    try
                    {
                        var svc = new MySqlAuthService(VariableDictionary._serverIP, VariableDictionary._serverPort.ToString());
                        // EnsureUserTableExists 也可能会做 I/O 操作，放到后台执行更安全
                        svc.EnsureUserTableExists();
                        return svc.AuthenticateUser(VariableDictionary._userName, VariableDictionary._passWord);
                    }
                    catch (Exception exInner)
                    {
                        LogManager.Instance.LogInfo($"后台认证出错: {exInner.Message}");
                        throw;
                    }
                });

                if (authOk)
                {
                    // 认证成功：保存配置（UI 操作/文件写入在 UI 线程执行以保证安全）
                    if (ChkSavePassword.IsChecked == true)
                        SaveConfig(true);
                    else
                        SaveConfig(false);

                    TxtStatus.Text = "登录成功。";

                    try
                    {
                        // 创建 DatabaseManager 可能也需要一定时间，放到后台执行，但不要用 ConfigureAwait(false) 以便 await 后在 UI 线程继续运行并访问控件。
                        var db = await Task.Run(() =>
                        {
                            try
                            {
                                VariableDictionary._newConnectionString = $"Server={VariableDictionary._serverIP};Port={VariableDictionary._serverPort};Database=cad_sw_library;Uid=root;Pwd=root;";
                                
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
                            // 注意：这里直接 await，不使用 ConfigureAwait(false)，保证后续对 UI 的访问安全
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
                port = 3306;
                TxtServerPort.Text = "3306";
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
                TxtStatus.Text = "连接服务器成功，但读取部门失败，请检查数据库配置或凭据。";
                MessageBox.Show("连接服务器成功，但读取部门失败，请检查数据库或查看日志。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
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
    }
}
