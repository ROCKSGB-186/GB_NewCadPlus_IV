using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using System;
using System.Windows;
using System.Windows.Controls;
using Button = System.Windows.Controls.Button;
using CheckBox = System.Windows.Controls.CheckBox;
using ComboBox = System.Windows.Controls.ComboBox;
using HorizontalAlignment = System.Windows.HorizontalAlignment;
using MessageBox = System.Windows.MessageBox;
using Orientation = System.Windows.Controls.Orientation;
using TextBox = System.Windows.Controls.TextBox;

namespace GB_NewCadPlus_IV
{
    public partial class WpfMainWindow
    {
        string host = VariableDictionary._serverIP; // 或者从配置读取
        int port = VariableDictionary._dataBaseServerPort;  // 或者从配置读取
        string dbType = VariableDictionary._databaseType; // "MYSQL" 或 "DM"
        string user = VariableDictionary._dbUserName;
        string pwd = VariableDictionary._dbPassWord;


        /// <summary>
        /// 新增用户按钮点击处理（界面事件绑定）
        /// 逻辑：检查部门选择 -> 调用用户编辑对话 -> 调用服务器 API -> 刷新用户列表
        /// </summary>
        private void BtnAddUserManaged_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 确认已选择部门
                if (!(DepartmentsGrid?.SelectedItem is DepartmentModel selectedDept))
                {
                    MessageBox.Show("请先选择一个部门。", "提示", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }

                // 确保用户服务已初始化
                // 打开用户编辑对话（新增）
                if (!ShowUserEditorDialog(null, "新增用户", out var result)) return;

                // 执行新增
                string deptName = string.IsNullOrWhiteSpace(selectedDept.Name) ? (selectedDept.RealName ?? string.Empty) : selectedDept.Name;
                var response = _apiService.AddUserAsync(result.Username, result.Password, selectedDept.Id, deptName, result.Role, result.IsActive, result.RealName, result.Gender, result.Phone, result.Email).GetAwaiter().GetResult();
                if (!response.Success)
                {
                    MessageBox.Show("新增用户失败，可能是用户名已存在。", "失败", MessageBoxButton.OK, MessageBoxImage.Hand);
                    return;
                }

                // 刷新用户列表与部门
                LoadUsersForDepartment(selectedDept.Id);
                _ = RefreshDepartmentsAsync();
                TxtSearchUser.Text = result.Username;
                TxtStatus.Text = "已新增用户：" + result.Username;
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增用户异常：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        /// <summary>
        /// 编辑用户按钮点击处理（界面事件绑定）
        /// 逻辑：检查选中用户 -> 打开编辑对话 -> 调用服务器 API -> 刷新用户列表
        /// </summary>
        private void BtnEditUserManaged_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 检查是否有选中的用户
                if (!(UsersGrid?.SelectedItem is UserModel selectedUser))
                {
                    MessageBox.Show("请先选择要编辑的用户。", "提示", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }

                // 确保服务已初始化
                // 打开编辑对话并获取结果
                if (!ShowUserEditorDialog(selectedUser, "编辑用户", out var result)) return;

                // 先显式声明并赋值 selDept，以避免使用未赋值的局部变量（修复 CS0165）
                DepartmentModel selDept = DepartmentsGrid?.SelectedItem as DepartmentModel; // 尝试从部门列表取出选中项并转换
                // 计算部门 ID（如果未选中则为 null）
                int? deptId = selDept != null ? selDept.Id : (int?)null;
                // 计算部门显示名（优先 Name，Name 为空则用 RealName，否则为空字符串）
                string deptName = selDept == null ? string.Empty : (string.IsNullOrWhiteSpace(selDept.Name) ? selDept.RealName ?? string.Empty : selDept.Name);

                // 调用服务更新用户（如果密码为空则不修改密码）
                var response = _apiService.UpdateUserAsync(
                    selectedUser.Id,
                    result.Username,
                    result.Role,
                    result.IsActive,
                    deptId,
                    deptName,
                    string.IsNullOrWhiteSpace(result.Password) ? null : result.Password,
                    result.RealName,
                    result.Gender,
                    result.Phone,
                    result.Email).GetAwaiter().GetResult();

                if (!response.Success)
                {
                    MessageBox.Show("编辑用户失败。", "失败", MessageBoxButton.OK, MessageBoxImage.Hand);
                    return;
                }

                // 刷新当前部门的用户列表（如果存在已选部门）
                if (selDept != null) LoadUsersForDepartment(selDept.Id);
                // 更新搜索框与状态栏
                TxtSearchUser.Text = result.Username;
                TxtStatus.Text = "已更新用户：" + result.Username;
            }
            catch (Exception ex)
            {
                MessageBox.Show("编辑用户异常：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        /// <summary>
        /// 删除用户按钮点击处理
        /// 逻辑：确认 -> 调用服务器 API -> 刷新列表与部门
        /// </summary>
        private void BtnDeleteUserManaged_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!(UsersGrid?.SelectedItem is UserModel selectedUser))
                {
                    System.Windows.MessageBox.Show("请先选择要删除的用户。", "提示", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }

                if (MessageBox.Show($"确认删除用户：{selectedUser.Username} ?", "确认删除", MessageBoxButton.YesNo, MessageBoxImage.Exclamation) != MessageBoxResult.Yes)
                    return;

                if (!_apiService.DeleteUserAsync(selectedUser.Id).GetAwaiter().GetResult().Success)
                {
                    MessageBox.Show("删除用户失败。", "失败", MessageBoxButton.OK, MessageBoxImage.Hand);
                    return;
                }

                if (DepartmentsGrid?.SelectedItem is DepartmentModel dept) LoadUsersForDepartment(dept.Id);
                _ = RefreshDepartmentsAsync();
                TxtStatus.Text = "已删除用户：" + selectedUser.Username;
            }
            catch (Exception ex)
            {
                MessageBox.Show("删除用户异常：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        private sealed class UserEditorDialogResult
        {
            public string Username { get; set; }
            public string Password { get; set; }
            public string RealName { get; set; }
            public string Gender { get; set; }
            public string Phone { get; set; }
            public string Email { get; set; }
            public string Role { get; set; }
            public bool IsActive { get; set; }
        }

        /// <summary>
        /// 显示用户编辑对话（通用方法，既用于新增也用于编辑）
        /// 参数：
        ///   initial - 传入非 null 则为编辑模式，null 为新增
        ///   title - 对话框标题
        /// out result - 返回用户输入的数据结构
        /// 返回：true 表示用户点击确定并通过校验
        /// </summary>
        private bool ShowUserEditorDialog(UserModel initial, string title, out UserEditorDialogResult result)
        {
            // 1. 初始化 out 参数
            var tempResult = new UserEditorDialogResult();
            bool isEdit = initial != null;

            // 2. 构造对话窗口
            var win = new Window
            {
                Title = title,
                Owner = Window.GetWindow(this),
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                SizeToContent = SizeToContent.WidthAndHeight,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.ToolWindow,
                MinWidth = 480
            };

            // 3. 创建布局 Grid
            var grid = new Grid { Margin = new Thickness(10) };
            for (int i = 0; i < 10; i++) grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            int row = 0;

            // --- 用户名 ---
            var lblUser = new TextBlock { Text = "用户名:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblUser, row); Grid.SetColumn(lblUser, 0); grid.Children.Add(lblUser);
            var tbUser = new TextBox { Margin = new Thickness(4), Text = initial?.Username ?? (TxtSearchUser?.Text ?? string.Empty).Trim() };
            Grid.SetRow(tbUser, row); Grid.SetColumn(tbUser, 1); grid.Children.Add(tbUser);
            row++;

            // --- 真实姓名 ---
            var lblReal = new TextBlock { Text = "真实姓名:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblReal, row); Grid.SetColumn(lblReal, 0); grid.Children.Add(lblReal);
            var tbReal = new TextBox { Margin = new Thickness(4), Text = initial?.RealName ?? string.Empty };
            Grid.SetRow(tbReal, row); Grid.SetColumn(tbReal, 1); grid.Children.Add(tbReal);
            row++;

            // --- 性别 ---
            var lblGender = new TextBlock { Text = "性别:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblGender, row); Grid.SetColumn(lblGender, 0); grid.Children.Add(lblGender);
            var cmbGender = new ComboBox { Margin = new Thickness(4), IsEditable = false };
            cmbGender.Items.Add("无信息"); cmbGender.Items.Add("男"); cmbGender.Items.Add("女");
            string g = string.IsNullOrWhiteSpace(initial?.Gender) ? "无信息" : initial.Gender.Trim();
            cmbGender.SelectedItem = (g == "男" || g == "女") ? (object)g : "无信息";
            Grid.SetRow(cmbGender, row); Grid.SetColumn(cmbGender, 1); grid.Children.Add(cmbGender);
            row++;

            // --- 电话 ---
            var lblPhone = new TextBlock { Text = "电话:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblPhone, row); Grid.SetColumn(lblPhone, 0); grid.Children.Add(lblPhone);
            var tbPhone = new TextBox { Margin = new Thickness(4), Text = initial?.Phone ?? string.Empty };
            Grid.SetRow(tbPhone, row); Grid.SetColumn(tbPhone, 1); grid.Children.Add(tbPhone);
            row++;

            // --- Email ---
            var lblEmail = new TextBlock { Text = "Email:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblEmail, row); Grid.SetColumn(lblEmail, 0); grid.Children.Add(lblEmail);
            var tbEmail = new TextBox { Margin = new Thickness(4), Text = initial?.Email ?? string.Empty };
            Grid.SetRow(tbEmail, row); Grid.SetColumn(tbEmail, 1); grid.Children.Add(tbEmail);
            row++;

            // --- 密码 ---
            var lblPwd = new TextBlock { Text = isEdit ? "新密码(留空不修改):" : "密码:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblPwd, row); Grid.SetColumn(lblPwd, 0); grid.Children.Add(lblPwd);
            var pbPwd = new PasswordBox { Margin = new Thickness(4) };
            Grid.SetRow(pbPwd, row); Grid.SetColumn(pbPwd, 1); grid.Children.Add(pbPwd);
            row++;

            // --- 确认密码 ---
            var lblConfirm = new TextBlock { Text = isEdit ? "确认新密码:" : "确认密码:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblConfirm, row); Grid.SetColumn(lblConfirm, 0); grid.Children.Add(lblConfirm);
            var pbConfirm = new PasswordBox { Margin = new Thickness(4) };
            Grid.SetRow(pbConfirm, row); Grid.SetColumn(pbConfirm, 1); grid.Children.Add(pbConfirm);
            row++;

            // --- 角色 ---
            var lblRole = new TextBlock { Text = "角色:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblRole, row); Grid.SetColumn(lblRole, 0); grid.Children.Add(lblRole);
            var tbRole = new TextBox { Margin = new Thickness(4), Text = string.IsNullOrWhiteSpace(initial?.Role) ? "user" : initial.Role };
            Grid.SetRow(tbRole, row); Grid.SetColumn(tbRole, 1); grid.Children.Add(tbRole);
            row++;

            // --- 是否启用 ---
            var lblActive = new TextBlock { Text = "是否启用:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6, 0, 6) };
            Grid.SetRow(lblActive, row); Grid.SetColumn(lblActive, 0); grid.Children.Add(lblActive);
            var cbActive = new CheckBox { Margin = new Thickness(4), IsChecked = initial == null || initial.IsActive, VerticalAlignment = VerticalAlignment.Center };
            Grid.SetRow(cbActive, row); Grid.SetColumn(cbActive, 1); grid.Children.Add(cbActive);
            row++;

            // --- 确认/取消按钮 ---
            var spButtons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
            var btnOk = new Button { Content = "确定", Width = 80, Margin = new Thickness(4) };
            var btnCancel = new Button { Content = "取消", Width = 80, Margin = new Thickness(4) };
            spButtons.Children.Add(btnOk); spButtons.Children.Add(btnCancel);
            Grid.SetRow(spButtons, row); Grid.SetColumn(spButtons, 0); Grid.SetColumnSpan(spButtons, 2);
            grid.Children.Add(spButtons);

            // 4. 事件绑定
            btnCancel.Click += (s, e) => win.DialogResult = false;

            btnOk.Click += (s, e) =>
            {
                // 获取输入值
                var username = (tbUser.Text ?? string.Empty).Trim();
                var real = (tbReal.Text ?? string.Empty).Trim();
                var gender = cmbGender.SelectedItem as string ?? "无信息";
                var phone = (tbPhone.Text ?? string.Empty).Trim();
                var email = (tbEmail.Text ?? string.Empty).Trim();
                var pwd = pbPwd.Password ?? string.Empty;
                var confirm = pbConfirm.Password ?? string.Empty;
                var role = (tbRole.Text ?? string.Empty).Trim();

                // 校验逻辑
                if (string.IsNullOrWhiteSpace(username))
                {
                    MessageBox.Show("请输入用户名。", "提示", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    tbUser.Focus(); return;
                }
                if (string.IsNullOrWhiteSpace(real))
                {
                    MessageBox.Show("请输入真实姓名。", "提示", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    tbReal.Focus(); return;
                }
                if (!string.IsNullOrWhiteSpace(email) && !email.Contains("@"))
                {
                    MessageBox.Show("Email 格式不正确。", "提示", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    tbEmail.Focus(); return;
                }
                if (!isEdit && string.IsNullOrWhiteSpace(pwd))
                {
                    MessageBox.Show("请输入密码。", "提示", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }
                if (!string.IsNullOrEmpty(pwd) && pwd != confirm)
                {
                    MessageBox.Show("两次密码不一致。", "提示", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }

                // ✅ 修复核心：使用局部变量构建结果对象
                tempResult = new UserEditorDialogResult
                {
                    Username = username,
                    Password = pwd,
                    RealName = real,
                    Gender = gender,
                    Phone = phone,
                    Email = email,
                    Role = string.IsNullOrWhiteSpace(role) ? "user" : role,
                    IsActive = cbActive.IsChecked ?? true
                };

                // 关闭对话框
                win.DialogResult = true;
            };

            win.Content = grid;
            result= tempResult;
            // 显示对话框
            bool? dialogResult = win.ShowDialog();

            // 如果用户点击了确定 (true) 且 result 不为 null，则返回 true
            if (dialogResult == true && tempResult != null)
            {
                return true;
            }

            // 否则返回 false，并确保 result 是一个空对象（避免调用方拿到未初始化的对象）
            result = new UserEditorDialogResult();
            return false;
        }
    }
}
