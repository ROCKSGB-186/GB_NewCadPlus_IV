using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using System;

namespace GB_NewCadPlus_IV
{
    public partial class WpfMainWindow
    {
        private void BtnAddUserManaged_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                // 中文注释：新增用户前必须先选中部门，确保用户归属明确。
                var selDept = DepartmentsGrid.SelectedItem as DepartmentModel;
                if (selDept == null)
                {
                    System.Windows.MessageBox.Show("请先选择一个部门。", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }

                if (!EnsureSvcInitialized())
                {
                    System.Windows.MessageBox.Show("用户服务未初始化，请先检查数据库连接。", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }

                // 中文注释：新增时传入 null，弹窗要求填写密码。
                if (!ShowUserEditorDialog(null, "新增用户", out var editorResult))
                {
                    return;
                }

                var deptName = string.IsNullOrWhiteSpace(selDept.Name)
                    ? (selDept.RealName ?? string.Empty)
                    : selDept.Name;

                var ok = _svc.AddUser(
                    editorResult.Username,
                    editorResult.Password,
                    selDept.Id,
                    deptName,
                    editorResult.Role,
                    editorResult.IsActive,
                    editorResult.RealName,
                    editorResult.Gender,
                    editorResult.Phone,
                    editorResult.Email);

                if (!ok)
                {
                    System.Windows.MessageBox.Show("新增用户失败，可能是用户名已存在。", "失败", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    return;
                }

                LoadUsersForDepartment(selDept.Id);
                RefreshDepartmentsAsync();
                TxtSearchUser.Text = editorResult.Username;
                TxtStatus.Text = $"已新增用户：{editorResult.Username}";
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("新增用户异常：" + ex.Message, "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        private void BtnEditUserManaged_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                var selUser = UsersGrid.SelectedItem as UserModel;
                if (selUser == null)
                {
                    System.Windows.MessageBox.Show("请先选择要编辑的用户。", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }

                if (!EnsureSvcInitialized())
                {
                    System.Windows.MessageBox.Show("用户服务未初始化，请先检查数据库连接。", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }

                // 中文注释：编辑时允许密码留空，留空表示不修改密码。
                if (!ShowUserEditorDialog(selUser, "编辑用户", out var editorResult))
                {
                    return;
                }

                var selDept = DepartmentsGrid.SelectedItem as DepartmentModel;
                var deptId = selDept?.Id;
                var deptName = selDept == null
                    ? string.Empty
                    : (string.IsNullOrWhiteSpace(selDept.Name) ? (selDept.RealName ?? string.Empty) : selDept.Name);

                var ok = _svc.UpdateUser(
                    selUser.Id,
                    editorResult.Username,
                    editorResult.Role,
                    editorResult.IsActive,
                    deptId,
                    deptName,
                    string.IsNullOrWhiteSpace(editorResult.Password) ? null : editorResult.Password,
                    editorResult.RealName,
                    editorResult.Gender,
                    editorResult.Phone,
                    editorResult.Email);

                if (!ok)
                {
                    System.Windows.MessageBox.Show("编辑用户失败。", "失败", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    return;
                }

                if (selDept != null)
                {
                    LoadUsersForDepartment(selDept.Id);
                }

                TxtSearchUser.Text = editorResult.Username;
                TxtStatus.Text = $"已更新用户：{editorResult.Username}";
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("编辑用户异常：" + ex.Message, "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        private void BtnDeleteUserManaged_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                var selUser = UsersGrid.SelectedItem as UserModel;
                if (selUser == null)
                {
                    System.Windows.MessageBox.Show("请先选择要删除的用户。", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }

                var confirm = System.Windows.MessageBox.Show(
                    $"确认删除用户：{selUser.Username} ?",
                    "确认删除",
                    System.Windows.MessageBoxButton.YesNo,
                    System.Windows.MessageBoxImage.Warning);
                if (confirm != System.Windows.MessageBoxResult.Yes)
                {
                    return;
                }

                if (!EnsureSvcInitialized())
                {
                    System.Windows.MessageBox.Show("用户服务未初始化，请先检查数据库连接。", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }

                var ok = _svc.DeleteUser(selUser.Id);
                if (!ok)
                {
                    System.Windows.MessageBox.Show("删除用户失败。", "失败", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    return;
                }

                var selDept = DepartmentsGrid.SelectedItem as DepartmentModel;
                if (selDept != null)
                {
                    LoadUsersForDepartment(selDept.Id);
                }

                RefreshDepartmentsAsync();
                TxtStatus.Text = $"已删除用户：{selUser.Username}";
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("删除用户异常：" + ex.Message, "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
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

        private bool ShowUserEditorDialog(UserModel initial, string title, out UserEditorDialogResult result)
        {
            result = null;
            var isEdit = initial != null;

            var win = new System.Windows.Window
            {
                Title = title,
                Owner = System.Windows.Window.GetWindow(this),
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner,
                SizeToContent = System.Windows.SizeToContent.WidthAndHeight,
                ResizeMode = System.Windows.ResizeMode.NoResize,
                WindowStyle = System.Windows.WindowStyle.ToolWindow,
                MinWidth = 480
            };

            var grid = new System.Windows.Controls.Grid { Margin = new System.Windows.Thickness(10) };
            for (int i = 0; i < 10; i++)
            {
                grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
            }

            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new System.Windows.GridLength(140) });
            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });

            var row = 0;

            var lblUser = new System.Windows.Controls.TextBlock { Text = "用户名:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            System.Windows.Controls.Grid.SetRow(lblUser, row);
            System.Windows.Controls.Grid.SetColumn(lblUser, 0);
            grid.Children.Add(lblUser);

            var tbUser = new System.Windows.Controls.TextBox { Margin = new System.Windows.Thickness(4), Text = initial?.Username ?? (TxtSearchUser.Text ?? string.Empty).Trim() };
            System.Windows.Controls.Grid.SetRow(tbUser, row);
            System.Windows.Controls.Grid.SetColumn(tbUser, 1);
            grid.Children.Add(tbUser);
            row++;

            var lblRealName = new System.Windows.Controls.TextBlock { Text = "真实姓名:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            System.Windows.Controls.Grid.SetRow(lblRealName, row);
            System.Windows.Controls.Grid.SetColumn(lblRealName, 0);
            grid.Children.Add(lblRealName);

            var tbRealName = new System.Windows.Controls.TextBox { Margin = new System.Windows.Thickness(4), Text = initial?.RealName ?? string.Empty };
            System.Windows.Controls.Grid.SetRow(tbRealName, row);
            System.Windows.Controls.Grid.SetColumn(tbRealName, 1);
            grid.Children.Add(tbRealName);
            row++;

            var lblGender = new System.Windows.Controls.TextBlock { Text = "性别:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            System.Windows.Controls.Grid.SetRow(lblGender, row);
            System.Windows.Controls.Grid.SetColumn(lblGender, 0);
            grid.Children.Add(lblGender);

            var cmbGender = new System.Windows.Controls.ComboBox { Margin = new System.Windows.Thickness(4), IsEditable = false };
            cmbGender.Items.Add("无信息");
            cmbGender.Items.Add("男");
            cmbGender.Items.Add("女");
            var initialGender = string.IsNullOrWhiteSpace(initial?.Gender) ? "无信息" : initial.Gender.Trim();
            cmbGender.SelectedItem = initialGender == "男" || initialGender == "女" ? initialGender : "无信息";
            System.Windows.Controls.Grid.SetRow(cmbGender, row);
            System.Windows.Controls.Grid.SetColumn(cmbGender, 1);
            grid.Children.Add(cmbGender);
            row++;

            var lblPhone = new System.Windows.Controls.TextBlock { Text = "电话:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            System.Windows.Controls.Grid.SetRow(lblPhone, row);
            System.Windows.Controls.Grid.SetColumn(lblPhone, 0);
            grid.Children.Add(lblPhone);

            var tbPhone = new System.Windows.Controls.TextBox { Margin = new System.Windows.Thickness(4), Text = initial?.Phone ?? string.Empty };
            System.Windows.Controls.Grid.SetRow(tbPhone, row);
            System.Windows.Controls.Grid.SetColumn(tbPhone, 1);
            grid.Children.Add(tbPhone);
            row++;

            var lblEmail = new System.Windows.Controls.TextBlock { Text = "Email:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            System.Windows.Controls.Grid.SetRow(lblEmail, row);
            System.Windows.Controls.Grid.SetColumn(lblEmail, 0);
            grid.Children.Add(lblEmail);

            var tbEmail = new System.Windows.Controls.TextBox { Margin = new System.Windows.Thickness(4), Text = initial?.Email ?? string.Empty };
            System.Windows.Controls.Grid.SetRow(tbEmail, row);
            System.Windows.Controls.Grid.SetColumn(tbEmail, 1);
            grid.Children.Add(tbEmail);
            row++;

            var lblPwd = new System.Windows.Controls.TextBlock { Text = isEdit ? "新密码(留空不修改):" : "密码:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            System.Windows.Controls.Grid.SetRow(lblPwd, row);
            System.Windows.Controls.Grid.SetColumn(lblPwd, 0);
            grid.Children.Add(lblPwd);

            var pbPwd = new System.Windows.Controls.PasswordBox { Margin = new System.Windows.Thickness(4) };
            System.Windows.Controls.Grid.SetRow(pbPwd, row);
            System.Windows.Controls.Grid.SetColumn(pbPwd, 1);
            grid.Children.Add(pbPwd);
            row++;

            var lblConfirm = new System.Windows.Controls.TextBlock { Text = isEdit ? "确认新密码:" : "确认密码:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            System.Windows.Controls.Grid.SetRow(lblConfirm, row);
            System.Windows.Controls.Grid.SetColumn(lblConfirm, 0);
            grid.Children.Add(lblConfirm);

            var pbConfirm = new System.Windows.Controls.PasswordBox { Margin = new System.Windows.Thickness(4) };
            System.Windows.Controls.Grid.SetRow(pbConfirm, row);
            System.Windows.Controls.Grid.SetColumn(pbConfirm, 1);
            grid.Children.Add(pbConfirm);
            row++;

            var lblRole = new System.Windows.Controls.TextBlock { Text = "角色:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            System.Windows.Controls.Grid.SetRow(lblRole, row);
            System.Windows.Controls.Grid.SetColumn(lblRole, 0);
            grid.Children.Add(lblRole);

            var tbRole = new System.Windows.Controls.TextBox { Margin = new System.Windows.Thickness(4), Text = string.IsNullOrWhiteSpace(initial?.Role) ? "user" : initial.Role };
            System.Windows.Controls.Grid.SetRow(tbRole, row);
            System.Windows.Controls.Grid.SetColumn(tbRole, 1);
            grid.Children.Add(tbRole);
            row++;

            var lblActive = new System.Windows.Controls.TextBlock { Text = "是否启用:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 6, 0, 6) };
            System.Windows.Controls.Grid.SetRow(lblActive, row);
            System.Windows.Controls.Grid.SetColumn(lblActive, 0);
            grid.Children.Add(lblActive);

            var cbActive = new System.Windows.Controls.CheckBox { Margin = new System.Windows.Thickness(4), IsChecked = initial?.IsActive ?? true, VerticalAlignment = System.Windows.VerticalAlignment.Center };
            System.Windows.Controls.Grid.SetRow(cbActive, row);
            System.Windows.Controls.Grid.SetColumn(cbActive, 1);
            grid.Children.Add(cbActive);
            row++;

            var panelBtns = new System.Windows.Controls.StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                Margin = new System.Windows.Thickness(0, 10, 0, 0)
            };
            var btnOk = new System.Windows.Controls.Button { Content = "确定", Width = 80, Margin = new System.Windows.Thickness(4) };
            var btnCancel = new System.Windows.Controls.Button { Content = "取消", Width = 80, Margin = new System.Windows.Thickness(4) };
            panelBtns.Children.Add(btnOk);
            panelBtns.Children.Add(btnCancel);
            System.Windows.Controls.Grid.SetRow(panelBtns, row);
            System.Windows.Controls.Grid.SetColumn(panelBtns, 0);
            System.Windows.Controls.Grid.SetColumnSpan(panelBtns, 2);
            grid.Children.Add(panelBtns);

            win.Content = grid;

            UserEditorDialogResult localResult = null;

            btnCancel.Click += (s, e) => win.DialogResult = false;
            btnOk.Click += (s, e) =>
            {
                var username = (tbUser.Text ?? string.Empty).Trim();
                var realName = (tbRealName.Text ?? string.Empty).Trim();
                var gender = (cmbGender.SelectedItem as string) ?? "无信息";
                var phone = (tbPhone.Text ?? string.Empty).Trim();
                var email = (tbEmail.Text ?? string.Empty).Trim();
                var pwd = pbPwd.Password ?? string.Empty;
                var confirmPwd = pbConfirm.Password ?? string.Empty;
                var role = (tbRole.Text ?? string.Empty).Trim();

                if (string.IsNullOrWhiteSpace(username))
                {
                    System.Windows.MessageBox.Show("请输入用户名。", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    tbUser.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(realName))
                {
                    System.Windows.MessageBox.Show("请输入真实姓名。", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    tbRealName.Focus();
                    return;
                }

                if (!string.IsNullOrWhiteSpace(email) && !email.Contains("@"))
                {
                    System.Windows.MessageBox.Show("Email 格式不正确。", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    tbEmail.Focus();
                    return;
                }

                if (!isEdit && string.IsNullOrWhiteSpace(pwd))
                {
                    System.Windows.MessageBox.Show("请输入密码。", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }

                if (!string.IsNullOrEmpty(pwd) && pwd != confirmPwd)
                {
                    System.Windows.MessageBox.Show("两次密码不一致。", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }

                localResult = new UserEditorDialogResult
                {
                    Username = username,
                    Password = pwd,
                    RealName = realName,
                    Gender = gender,
                    Phone = phone,
                    Email = email,
                    Role = string.IsNullOrWhiteSpace(role) ? "user" : role,
                    IsActive = cbActive.IsChecked ?? true
                };

                win.DialogResult = true;
            };

            var dialogOk = win.ShowDialog() == true;
            if (!dialogOk || localResult == null)
            {
                return false;
            }

            result = localResult;
            return true;
        }
    }
}
