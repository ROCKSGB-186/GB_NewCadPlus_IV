using System;
using System.Windows;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;

namespace GB_NewCadPlus_IV
{
    public partial class WpfMainWindow
    {
        private void 转换CSV按钮_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("转换CSV功能正在完善中。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void 重载CSV按钮_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("重载CSV功能正在完善中。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void 插入计算表_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var command = UnifiedCommandManager.GetCommand("插入计算表");
                if (command != null)
                {
                    command.Invoke();
                    return;
                }

                System.Windows.MessageBox.Show("未找到“插入计算表”命令。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"插入计算表_Click 执行失败: {ex.Message}");
                System.Windows.MessageBox.Show($"插入计算表失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
