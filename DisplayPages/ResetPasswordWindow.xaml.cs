using System;
using System.Windows;
using GB_NewCadPlus_IV.Helpers;
using MessageBox = System.Windows.MessageBox;

namespace GB_NewCadPlus_IV
{
    public partial class ResetPasswordWindow : Window
    {
        public ResetPasswordWindow()
        {
            InitializeComponent();
        }

        public void SetUsername(string username)
        {
            TxtUsername.Text = username ?? string.Empty;
        }

        private async void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            // 读取用户输入，并在发送请求前完成基本校验。
            string username = TxtUsername.Text.Trim();
            string phone = TxtPhone.Text.Trim();
            string email = TxtEmail.Text.Trim();
            string newPassword = PwdNewPassword.Password;
            string confirmPassword = PwdConfirmPassword.Password;

            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(phone) || string.IsNullOrWhiteSpace(email))
            {
                // 三项身份信息必须同时填写，避免发送无效请求。
                TxtStatus.Text = "请完整填写登录账号、手机号和邮箱。";
                return;
            }
            if (newPassword.Length < 6)
            {
                TxtStatus.Text = "新密码长度不能少于6位。";
                return;
            }
            if (!string.Equals(newPassword, confirmPassword, StringComparison.Ordinal))
            {
                TxtStatus.Text = "两次输入的新密码不一致。";
                return;
            }

            BtnReset.IsEnabled = false;
            TxtStatus.Text = "正在验证身份并修改密码...";
            try
            {
                // 由服务器核验账号、手机号和邮箱，客户端不接触数据库。
                AuthUserDepartmentApiService.MutationResponse response = await new AuthUserDepartmentApiService()
                    .ResetPasswordAsync(username, phone, email, newPassword);
                if (!response.Success)
                {
                    TxtStatus.Text = response.Message;
                    return;
                }

                MessageBox.Show(response.Message, "修改密码", MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = true;
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "修改密码失败：" + ex.Message;
            }
            finally
            {
                BtnReset.IsEnabled = true;
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
