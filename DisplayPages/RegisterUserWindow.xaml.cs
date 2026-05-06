using Dapper;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace GB_NewCadPlus_IV
{
    /// <summary>
    /// 注册用户窗口
    /// </summary>
    public partial class RegisterUserWindow : Window
    {
        private readonly string _host;// 主机名 数据库主机
        private readonly int _port;// 端口号 端口号 数据库端口
        private readonly List<(int Id, string Name)> _departments;// 部门列表

        /// <summary>
        /// 是否注册成功
        /// </summary>
        public bool RegistrationSucceeded { get; private set; } = false;
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="host">主机名</param>
        /// <param name="port">端口号</param>
        /// <param name="departments">部门列表</param>
        public RegisterUserWindow(string host, int port, List<(int Id, string Name)> departments = null)
        {
            InitializeComponent();
            _host = host;
            _port = port;
            _departments = departments ?? new List<(int, string)>();
            Loaded += RegisterUserWindow_Loaded;
        }
        /// <summary>
        /// 窗口加载时填充部门下拉
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RegisterUserWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // 如果构造时传入了部门，则直接使用；否则从服务器拉取 departments（保证已创建表与同步）
                if (_departments != null && _departments.Count > 0)
                {
                    CmbDept.ItemsSource = _departments.Select(d => new { d.Id, d.Name }).ToList();
                    CmbDept.SelectedIndex = 0;
                    return;
                }

                // 否则从服务读取部门
                var svc = new DMAuthService(_host, _port.ToString(),
                    string.IsNullOrWhiteSpace(VariableDictionary._userName) ? "SYSDBA" : VariableDictionary._userName,
                    string.IsNullOrWhiteSpace(VariableDictionary._passWord) ? "SYSDBA" : VariableDictionary._passWord);
                svc.EnsureAllTablesExist();
                // 先尝试同步分类到部门（幂等）
                try { svc.SyncDepartmentsFromCadCategories(); } catch { /* 忽略同步异常 */ }

                var depts = svc.GetDepartmentsWithCounts();
                if (depts != null && depts.Count > 0)
                {
                    CmbDept.ItemsSource = depts.Select(d => new { Id = d.Id, Name = d.Name }).ToList();
                    CmbDept.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"RegisterUserWindow_Loaded 出错: {ex.Message}");
            }
        }
        /// <summary>
        /// 取消注册
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        /// <summary>
        /// 注册
        /// </summary>
        private async void BtnRegister_Click(object sender, RoutedEventArgs e)
        {
            TxtStatus.Text = "";
            var username = TxtUsername.Text?.Trim();
            var pwd = PwdPassword.Password ?? "";
            var confirm = PwdConfirm.Password ?? "";

            if (string.IsNullOrWhiteSpace(username))
            {
                TxtStatus.Text = "请输入用户名。";
                return;
            }
            if (pwd.Length < 1)
            {
                TxtStatus.Text = "请输入密码。";
                return;
            }
            if (pwd != confirm)
            {
                TxtStatus.Text = "两次密码不一致。";
                return;
            }

            int deptId = 0;
            try
            {
                if (CmbDept.SelectedItem != null)
                {
                    var propId = CmbDept.SelectedItem.GetType().GetProperty("Id");
                    if (propId != null) deptId = Convert.ToInt32(propId.GetValue(CmbDept.SelectedItem));
                }
            }
            catch { }

            // 注册
            BtnRegister.IsEnabled = false;
            TxtStatus.Text = "正在注册...";

            var svc = new DMAuthService(_host, _port.ToString(),
                string.IsNullOrWhiteSpace(VariableDictionary._userName) ? "SYSDBA" : VariableDictionary._userName,
                string.IsNullOrWhiteSpace(VariableDictionary._passWord) ? "SYSDBA" : VariableDictionary._passWord);
            svc.EnsureAllTablesExist();
            try { svc.SyncDepartmentsFromCadCategories(); } catch { }

            await Task.Run(() =>
            {
                try
                {
                    // DMAuthService 当前签名：RegisterUser(string username, string password, int departmentId = 0, string departmentName = "")
                    var created = svc.RegisterUser(username, pwd, departmentId: deptId, departmentName: string.Empty);

                    Dispatcher.Invoke(() =>
                    {
                        BtnRegister.IsEnabled = true;
                        if (created)
                        {
                            RegistrationSucceeded = true;
                            MessageBox.Show("注册成功，请使用新账号登录。", "注册成功", MessageBoxButton.OK, MessageBoxImage.Information);
                            DialogResult = true;
                            Close();
                        }
                        else
                        {
                            TxtStatus.Text = "注册失败，请检查用户名是否已存在或联系管理员。";
                        }
                    });
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        BtnRegister.IsEnabled = true;
                        TxtStatus.Text = "注册异常：" + ex.Message;
                    });
                }
            });
        }
    }
}
