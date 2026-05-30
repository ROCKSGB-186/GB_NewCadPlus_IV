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
            //初始化配置文件路径，放在用户的应用数据目录下，确保有写权限且不同用户之间隔离
            //_configPath = Path.Combine(
            //    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            //    "GB_NewCadPlus_IV",
            //    "login_config.json");//C:\Users\ShiGu\AppData\Roaming\GB_NewCadPlus_IV\login_config.json

            // 同时设置全局变量中的缓存存储路径，供 LogManager 和其他需要存储数据的组件使用
            GetPath._cacheStoragePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "GB_CADPLUS",
                "CacheStorage");
            // 确保日志管理器使用新的存储路径（如果之前已初始化，则会切换路径）
            _configPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "GB_CADPLUS",
                "login_config.json");//C:\Users\ShiGu\AppData\Local\GB_CADPLUS\Logs

            LoadConfig();//加载配置
            Loaded += LoginWindow_Loaded;//注册窗口加载事件处理程序
        }

        /// <summary>
        /// 窗口加载事件：填充默认值、同步全局变量、测试网络并加载部门。
        /// </summary>
        private async void LoginWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. 确保服务器 IP 有默认值
                if (string.IsNullOrWhiteSpace(Login_ServerIP.Text))
                    Login_ServerIP.Text = "127.0.0.1";

                // 2. 确定当前数据库类型（UI > VariableDictionary > 默认 "DM"）
                string selectedDb = ResolveSelectedDatabaseType();
                VariableDictionary._databaseType = selectedDb;

                // 3. 根据数据库类型填充空的端口/用户名/密码
                ApplyDatabaseTypeDefaults(selectedDb);

                // 4. 将 UI 状态同步到全局变量（统一操作，消除分散赋值）
                SyncUiToGlobalVariables();

                // 5. TCP 连通性检测
                TxtStatus.Text = $"正在检测服务器连接... (数据库类型: {selectedDb})";
                bool tcpOk = await Task.Run(() =>
                    TestNetworkConnection(VariableDictionary._serverIP,
                        VariableDictionary._dataBaseServerPort, 3000));

                if (!tcpOk)
                {
                    TxtStatus.Text = $"无法连接到服务器 {VariableDictionary._serverIP}:{VariableDictionary._dataBaseServerPort}，请检查IP/端口。";
                    CmbDepartments.ItemsSource = null;
                    return;
                }

                // 6. 尝试加载部门列表
                bool loaded = await TryLoadDepartmentsAsync(VariableDictionary._serverIP,
                    VariableDictionary._dataBaseServerPort);
                if (loaded)
                {
                    TxtStatus.Text = $"已连接 {selectedDb} 并加载部门。";
                }
                else
                {
                    // 具体失败原因已由 TryLoadDepartmentsAsync 写入 TxtStatus
                    CmbDepartments.ItemsSource = null;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"LoginWindow_Loaded 异常: {ex}");
                TxtStatus.Text = "初始化窗口时发生错误。";
            }
        }

        /// <summary>
        /// 从配置文件恢复 UI 控件值，不修改任何全局变量。
        /// </summary>
        private void LoadConfig()
        {
            try
            {
                if (!File.Exists(_configPath)) // 配置文件不存在时直接返回，保持 UI 默认值
                    return;

                string json = File.ReadAllText(_configPath); // 读取配置文件内容
                var serializer = new JavaScriptSerializer(); // 使用 JavaScriptSerializer 反序列化 JSON 到 LoginConfig 对象
                var cfg = serializer.Deserialize<LoginConfig>(json); // 反序列化后的对象可能为 null，需检查
                if (cfg == null) // 反序列化失败时保持默认值并返回
                    return;

                // 恢复基本连接信息
                Login_ServerIP.Text = cfg.ServerIP ?? string.Empty; // 登录服务器 IP
                Login_DataBaseserverPort.Text = cfg.DataBaseserverPort ?? string.Empty; // 登录服务器端口
                Login_Username.Text = cfg.Username ?? string.Empty; // 登录用户名

                // 恢复密码（若保存了加密凭证）
                if (cfg.SavePassword && !string.IsNullOrWhiteSpace(cfg.EncryptedPassword))
                {
                    try
                    {
                        byte[] encryptedBytes = Convert.FromBase64String(cfg.EncryptedPassword); // 从 Base64 字符串转换回字节数组
                        byte[] decryptedBytes = ProtectedData.Unprotect(encryptedBytes, null,
                            DataProtectionScope.CurrentUser);  // 使用 DPAPI 解密，作用范围为当前用户
                        Login_Password.Password = Encoding.UTF8.GetString(decryptedBytes); // 将解密后的字节数组转换回字符串并设置到密码框
                        ChkSavePassword.IsChecked = true; // 恢复保存密码的勾选状态
                    }
                    catch (Exception ex)
                    {
                        // 解密失败（如在不同的 Windows 账户下）时，清空密码并取消勾选
                        LogManager.Instance.LogWarning("恢复已保存密码失败，可能需要重新输入: " + ex.Message);
                        Login_Password.Password = string.Empty; // 清空密码框
                        ChkSavePassword.IsChecked = false; // 取消保存密码的勾选状态
                    }
                }

                // 恢复数据库类型选择（触发 SelectionChanged 会自动填充相应默认值）
                if (!string.IsNullOrWhiteSpace(cfg.DatabaseType) && CmbDatabaseType != null)
                {
                    string targetDbType = cfg.DatabaseType.ToUpperInvariant(); // 标准化为大写以便比较
                    foreach (var item in CmbDatabaseType.Items) // 遍历 ComboBox 的选项，寻找匹配的数据库类型
                    {
                        if (item is ComboBoxItem cbi &&
                            string.Equals(cbi.Content as string, targetDbType, StringComparison.OrdinalIgnoreCase)) // 比较内容是否与目标数据库类型匹配（忽略大小写）
                        {
                            CmbDatabaseType.SelectedItem = item; // 设置选中项为匹配的数据库类型
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 配置损坏时不应阻止窗口显示，仅记录日志
                LogManager.Instance.LogError($"加载登录配置失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 将当前 UI 上的连接配置保存到本地 JSON 文件。
        /// 当 savePassword 为 false 时，会清除之前保存的密码。
        /// </summary>
        private void SaveConfig(bool savePassword)
        {
            try
            {
                // 确保配置文件目录存在
                string? configDir = Path.GetDirectoryName(_configPath);
                if (!string.IsNullOrEmpty(configDir) && !Directory.Exists(configDir))
                    Directory.CreateDirectory(configDir); // 创建目录

                // 从 UI 控件读取当前值
                var loginConfig = new LoginConfig
                {
                    ServerIP = Login_ServerIP.Text.Trim(), // 登录服务器 IP
                    DataBaseserverPort = Login_DataBaseserverPort.Text.Trim(), // 登录服务器端口
                    Username = Login_Username.Text.Trim(), // 登录用户名
                    SavePassword = savePassword, // 是否保存密码
                    // 数据库类型也一并持久化
                    DatabaseType = (CmbDatabaseType?.SelectedItem as ComboBoxItem)?.Content as string
                                   ?? VariableDictionary._databaseType
                                   ?? "DM",
                    configPath = configDir
                };

                // 处理密码加密
                if (savePassword)
                {
                    string plainPwd = Login_Password.Password ?? string.Empty; // 从密码框获取明文密码
                    byte[] plainBytes = Encoding.UTF8.GetBytes(plainPwd); // 将明文密码转换为字节数组
                    byte[] protectedBytes = ProtectedData.Protect(plainBytes, null, // 使用 DPAPI 加密，作用范围为当前用户
                        DataProtectionScope.CurrentUser); //    将加密后的字节数组转换为 Base64 字符串以便存储
                    loginConfig.EncryptedPassword = Convert.ToBase64String(protectedBytes); // 将加密后的密码存储在配置对象中
                }
                else
                {
                    // 不保存密码时，显式清除已保存的加密字段，防止下次加载时恢复旧密码
                    loginConfig.EncryptedPassword = null;
                }

                // 序列化并写入文件
                var serializer = new JavaScriptSerializer(); // 使用 JavaScriptSerializer 将 LoginConfig 对象序列化为 JSON 字符串
                string loginJson = serializer.Serialize(loginConfig);    // 将 JSON 字符串写入配置文件
                File.WriteAllText(_configPath, loginJson); // 写入文件

                TxtStatus.Text = "配置已保存。";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "保存配置失败：" + ex.Message;
                LogManager.Instance.LogError($"保存登录配置失败: {ex.Message}");
            }
        }

        #region 辅助方法

        /// <summary>
        /// 从 UI 或 VariableDictionary 解析当前数据库类型，优先级：UI选择 > VariableDictionary > 默认"DM"
        /// </summary>
        private string ResolveSelectedDatabaseType()
        {
            // 1) 从 UI 获取
            if (CmbDatabaseType?.SelectedItem is ComboBoxItem cbi && cbi.Content is string uiType)
                return uiType.ToUpperInvariant();

            // 2) 从 VariableDictionary 获取
            if (!string.IsNullOrWhiteSpace(VariableDictionary._databaseType))
                return VariableDictionary._databaseType.ToUpperInvariant();

            // 3) 默认
            return "DM";
        }

        /// <summary>
        /// 根据数据库类型为空白输入框填充默认值（不覆盖已有输入）。
        /// </summary>
        private void ApplyDatabaseTypeDefaults(string dbType)
        {
            // 端口默认值
            if (string.IsNullOrWhiteSpace(Login_DataBaseserverPort.Text))
            {
                Login_DataBaseserverPort.Text = (dbType == "MYSQL") ? "3308" : "5236";
            }

            // 应用用户默认凭据（注意：这是数据库管理员凭据，生产环境应通过服务端管理）
            if (string.IsNullOrWhiteSpace(Login_Username.Text))
            {
                Login_Username.Text = (dbType == "MYSQL") ? "root" : "SYSDBA";
            }
            if (string.IsNullOrWhiteSpace(Login_Password.Password))
            {
                Login_Password.Password = (dbType == "MYSQL") ? "123456" : "675756SGBsgb";
            }
        }

        /// <summary>
        /// 将当前 UI 控件中的连接信息同步到 VariableDictionary 全局变量。
        /// </summary>
        private void SyncUiToGlobalVariables()
        {
            // 服务器 IP
            VariableDictionary._serverIP = (Login_ServerIP.Text ?? string.Empty).Trim();

            // 端口（转换失败时回退默认）
            if (int.TryParse(Login_DataBaseserverPort.Text.Trim(), out int port) && port > 0)
                VariableDictionary._dataBaseServerPort = port;
            else
                VariableDictionary._dataBaseServerPort = 5236;

            // 应用用户凭据
            VariableDictionary._userName = Login_Username.Text.Trim();
            VariableDictionary._passWord = Login_Password.Password.Trim();

            // 数据库管理员凭据（当前仍为硬编码，后续可改为服务端管理）
            string dbType = VariableDictionary._databaseType ?? "DM";
            VariableDictionary._dbUserName = (dbType == "MYSQL") ? "root" : "SYSDBA";
            VariableDictionary._dbPassWord = (dbType == "MYSQL") ? "123456" : "675756SGBsgb";

            // API 端口（固定值）
            VariableDictionary._apiPort = 10010;
        }

        #endregion




        /// <summary>
        /// 尝试使用 DMAuthService 读取部门并填充下拉框，返回是否成功
        /// </summary>
        private async Task<bool> TryLoadDepartmentsAsync(string host, int port)
        {
            try
            {
                var selectedDb = (VariableDictionary._databaseType ?? "DM").ToUpperInvariant();
                var uiUser = string.IsNullOrWhiteSpace(VariableDictionary._dbUserName)
                    ? (selectedDb == "MYSQL" ? "root" : "SYSDBA")
                    : VariableDictionary._dbUserName.Trim();
                var uiPwd = string.IsNullOrWhiteSpace(VariableDictionary._dbPassWord)
                    ? (selectedDb == "MYSQL" ? "123456" : "675756SGBsgb")
                    : VariableDictionary._dbPassWord;

                // 确保使用正确的 DB 类型（优先 UI 选择）
                try
                {
                    if (CmbDatabaseType?.SelectedItem is ComboBoxItem cbi && cbi.Content is string s)
                        selectedDb = s.ToUpper().Trim();
                }
                catch { }

                List<DepartmentModel> depts;
                if (selectedDb == "MYSQL")
                {
                    var mySvc = new MySqlAuthService(host, port.ToString(), uiUser, uiPwd);
                    mySvc.EnsureAllTablesExist();
                    try { mySvc.SyncDepartmentsFromCadCategories(); } catch { }
                    depts = mySvc.GetDepartmentsWithCounts();
                }
                else
                {
                    var svc = new DMAuthService(host, port.ToString(), uiUser, uiPwd);
                    svc.EnsureAllTablesExist();
                    try { svc.SyncDepartmentsFromCadCategories(); } catch { }
                    depts = svc.GetDepartmentsWithCounts();
                }

                // 调试日志：检查 DisplayName 是否有值
                if (depts != null && depts.Count > 0)
                {
                    foreach (var d in depts)
                        LogManager.Instance.LogInfo($"部门: Id={d.Id}, Name={d.Name}, DisplayName={d.DisplayName}");
                }

                return await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        if (depts != null && depts.Count > 0)
                        {
                            CmbDepartments.ItemsSource = depts;
                            CmbDepartments.DisplayMemberPath = "DisplayName";
                            CmbDepartments.SelectedValuePath = "Id";
                            CmbDepartments.SelectedIndex = 0;
                        }
                        else
                        {
                            CmbDepartments.ItemsSource = null;
                            LogManager.Instance.LogInfo("部门列表为空，无法绑定。");
                        }
                    });
                    return true;
                });
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"TryLoadDepartmentsAsync 失败: {ex.Message}");
                Dispatcher.Invoke(() =>
                {
                    TxtStatus.Text = "读取部门失败：" + ex.Message;
                    CmbDepartments.ItemsSource = null;
                });
                return false;
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
            VariableDictionary._databaseType = CmbDatabaseType.SelectedItem?.ToString(); // 尝试从 UI 读取数据库类型，后续认证逻辑会根据这个值分支处理
            VariableDictionary._serverIP = Login_ServerIP.Text.Trim(); // 更新全局服务器 IP
            VariableDictionary._dataBaseServerPort = int.TryParse(Login_DataBaseserverPort.Text.Trim(), out int port) ? port : 5236; // 更新全局数据库端口
            VariableDictionary._apiPort = 10010; // API 端口固定为 10010，后续可改为 UI 可配置
            VariableDictionary._userName = Login_Username.Text.Trim(); // 更新全局用户名
            VariableDictionary._passWord = Login_Password.Password.Trim(); // 更新全局密码

            BtnLogin.IsEnabled = false; // 禁用登录按钮，防止重复点击
            TxtStatus.Text = "正在连接并验证用户..."; // 更新状态提示
            // 1) 先做快速 TCP 连通性检测；失败则直接退回 FormMain
            bool tcpOk = await Task.Run(() => TestNetworkConnection(VariableDictionary._serverIP, VariableDictionary._dataBaseServerPort));
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

                        if (VariableDictionary._databaseType == "MYSQL")
                        {
                            // --- MySQL 分支 ---
                            // 使用数据库管理凭据（root）初始化服务，确保具有建表和查询系统表的权限
                            // 注意：如果您环境中的 MySQL root 密码不同，请调整此处或从配置读取
                            var svc = new MySqlAuthService(
                                VariableDictionary._serverIP,
                                VariableDictionary._dataBaseServerPort.ToString(),
                                VariableDictionary._dbUserName,
                                VariableDictionary._dbPassWord
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
                                VariableDictionary._dataBaseServerPort.ToString(),
                                VariableDictionary._dbUserName,
                                VariableDictionary._dbPassWord
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
                            var selStr = sel.Content as string; // 尝试从 UI 读取数据库类型，优先 UI 选择
                            if (!string.IsNullOrWhiteSpace(selStr)) // 标准化为大写并去除空白
                            {
                                VariableDictionary._databaseType = selStr.ToUpper().Trim(); // 更新全局数据库类型标识
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
                                // 统一入口：根据数据库类型组装连接串，避免后续主窗口静默连接与登录连接不一致
                                string dbType = (VariableDictionary._databaseType ?? "DM").ToUpperInvariant();
                                if (dbType == "MYSQL")
                                {
                                    string dbPart = string.IsNullOrWhiteSpace(VariableDictionary._dataBaseName)
                                        ? "Database=cad_sw_library;"
                                        : $"Database={VariableDictionary._dataBaseName};";
                                    VariableDictionary._newConnectionString =
                                        $"Server={VariableDictionary._serverIP};Port={VariableDictionary._dataBaseServerPort};{dbPart}Uid={VariableDictionary._dbUserName};Pwd={VariableDictionary._dbPassWord};Allow User Variables=True;";
                                }
                                else
                                {
                                    string dbPart = string.IsNullOrWhiteSpace(VariableDictionary._dataBaseName)
                                        ? "Schema=CAD_SW_LIBRARY;"
                                        : $"Schema={VariableDictionary._dataBaseName};";
                                    VariableDictionary._newConnectionString =
                                        $"Server={VariableDictionary._serverIP};Port={VariableDictionary._dataBaseServerPort};{dbPart}User Id={VariableDictionary._dbUserName};Password={VariableDictionary._dbPassWord};";
                                }

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

                        var regWin = new RegisterUserWindow(VariableDictionary._serverIP, VariableDictionary._dataBaseServerPort, deptList) { Owner = this };
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
            VariableDictionary._serverIP = Login_ServerIP.Text.Trim();
            if (string.IsNullOrWhiteSpace(VariableDictionary._serverIP))
            {
                MessageBox.Show("请填写服务器IP地址。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                Login_ServerIP.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(Login_DataBaseserverPort.Text))
            {
                MessageBox.Show("请填写服务器端口。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                Login_DataBaseserverPort.Focus();
                return;
            }
            else
            {
                VariableDictionary._dataBaseServerPort = Convert.ToInt32(Login_DataBaseserverPort.Text);
            }

            TxtStatus.Text = "正在连接服务器...";

            // 先做 TCP 层检测，快速反馈
            bool tcpOk = await Task.Run(() => TestNetworkConnection(VariableDictionary._serverIP, VariableDictionary._dataBaseServerPort));
            if (!tcpOk)
            {
                MessageBox.Show($"无法连接到服务器 {VariableDictionary._serverIP}:{VariableDictionary._dataBaseServerPort}，请检查IP、端口或网络。", "连接失败", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtStatus.Text = "连接失败，请检查服务器IP或端口。";
                return;
            }

            // TCP 成功后尝试读取并初始化部门（TryLoadDepartmentsAsync 已包含初始化与同步）
            var loaded = await TryLoadDepartmentsAsync(VariableDictionary._serverIP, VariableDictionary._dataBaseServerPort);
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
                        if (string.IsNullOrWhiteSpace(Login_DataBaseserverPort.Text))
                            Login_DataBaseserverPort.Text = "3308"; // 你的 MySQL 端口示例
                        if (string.IsNullOrWhiteSpace(Login_Username.Text))
                            Login_Username.Text = "root"; // MySQL 管理连接默认账号
                        // 不自动设置密码，避免写入明文
                    }
                    else
                    {
                        // DM 默认端口为 5236
                        if (string.IsNullOrWhiteSpace(Login_DataBaseserverPort.Text))
                            Login_DataBaseserverPort.Text = "5236";
                        if (string.IsNullOrWhiteSpace(Login_Username.Text))
                            Login_Username.Text = "SYSDBA";
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
                    var dataBaseserver = Login_ServerIP.Text.Trim();
                    VariableDictionary._serverIP = dataBaseserver;
                    var dataBaseServerPort = Login_DataBaseserverPort.Text.Trim();
                    VariableDictionary._dataBaseServerPort = int.TryParse(dataBaseServerPort, out int port) ? port : 5236;
                    // 优先使用 UI 中填写的用户名，否则回退到 VariableDictionary 中可能已保存的用户名
                    //var user = string.IsNullOrWhiteSpace(Login_Username.Text) ? (VariableDictionary._userName ?? string.Empty) : Login_Username.Text.Trim();
                    //var pwd = Login_Password.Password.Trim();
                    // 记录用于测试的目标信息（不记录明文密码）
                    LogManager.Instance.LogInfo($"测试 MySQL 连接: {dataBaseserver}:{dataBaseServerPort} user = root ");
                    string dbPart = string.IsNullOrWhiteSpace(VariableDictionary._dataBaseName) ? string.Empty : $"Database={VariableDictionary._dataBaseName};";
                    var connStr = $"Server={dataBaseserver};Port={dataBaseServerPort};{dbPart}User Id=root;Password=123456;";
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
                        Login_ServerIP.Text.Trim(), // 服务器地址
                        Login_DataBaseserverPort.Text.Trim(), // 服务器端口
                        //Login_Username.Text.Trim(), // 用户名
                        //Login_Password.Password.Trim() // 密码
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
        public string DataBaseserverPort { get; set; }
        public string Username { get; set; }
        public bool SavePassword { get; set; }
        public string EncryptedPassword { get; set; }
        // 可选：持久化数据库类型（"DM" 或 "MYSQL"），用于下次启动时恢复选项
        public string DatabaseType { get; set; }
        public string configPath { get; set; }
    }
}
