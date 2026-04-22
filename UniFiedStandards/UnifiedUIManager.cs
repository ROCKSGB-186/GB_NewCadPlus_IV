using System;
using System.Linq;

namespace GB_NewCadPlus_IV.UniFiedStandards
{
    /// <summary>
    /// 统一界面管理器 - 管理WPF和WinForm界面的统一访问
    /// </summary>
    public static class UnifiedUIManager
    {
        private static FormMain? _winFormInstance;
        private static WpfMainWindow? _wpfInstance;

        /// <summary>
        /// 设置WinForm实例
        /// </summary>
        public static void SetWinFormInstance(FormMain instance)
        {
            _winFormInstance = instance;
        }

        /// <summary>
        /// 设置WPF实例
        /// </summary>
        public static void SetWpfInstance(WpfMainWindow instance)
        {
            _wpfInstance = instance;
        }

        /// <summary>
        /// 获取TextBox的值（自动从当前活动界面获取）
        /// </summary>
        public static string GetTextBoxValue(string textBoxName, string defaultValue = "")
        {
            System.Diagnostics.Debug.WriteLine($"尝试获取TextBox值: {textBoxName}");

            // 优先从WPF界面获取
            if (_wpfInstance != null)
            {
                string wpfValue = GetWpfTextBoxValue(textBoxName);//获取WPF界面TextBox值
                if (wpfValue != null)
                {
                    System.Diagnostics.Debug.WriteLine($"从WPF获取到值: {wpfValue}");
                    return string.IsNullOrEmpty(wpfValue) ? defaultValue : wpfValue;//返回TextBox值
                }
            }

            // 如果WPF界面没有或为空，从WinForm界面获取
            if (_winFormInstance != null)
            {
                string winFormValue = VariableDictionary.btnFileName;
                if (winFormValue != null)
                {
                    System.Diagnostics.Debug.WriteLine($"从WinForm获取到值: {winFormValue}");
                    return string.IsNullOrEmpty(winFormValue) ? defaultValue : winFormValue;
                }
            }

            System.Diagnostics.Debug.WriteLine($"未找到TextBox，返回默认值: {defaultValue}");
            return defaultValue;
        }

        /// <summary>
        /// 获取WPF界面TextBox值
        /// </summary>
        private static string GetWpfTextBoxValue(string textBoxName)
        {
            try
            {
                if (_wpfInstance == null)
                {
                    System.Diagnostics.Debug.WriteLine("WPF实例为空");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"在WPF中查找TextBox: {textBoxName}");

                // 首先尝试使用FindName方法（推荐方式）
                var textBoxByName = _wpfInstance.FindName(textBoxName);
                if (textBoxByName != null)
                {
                    // 检查找到的控件是否是TextBox类型
                    if (textBoxByName is System.Windows.Controls.TextBox textBox)
                    {
                        System.Diagnostics.Debug.WriteLine($"通过FindName找到TextBox: {textBoxName}");
                        return TextBoxValueHelper.GetTextBoxValue(textBox);
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"找到的控件不是TextBox类型，实际类型: {textBoxByName.GetType().Name}");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"在WPF中未找到TextBox: {textBoxName}");
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取WPF TextBox值时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 设置WPF界面TextBox值
        /// </summary>
        public static void SetWpfTextBoxValue(string textBoxName, string value)
        {
            try
            {
                if (_wpfInstance == null) return;

                // 首先尝试使用FindName方法
                var textBoxByName = _wpfInstance.FindName(textBoxName) as System.Windows.Controls.TextBox;
                if (textBoxByName != null)
                {
                    // 使用WPF主窗口的Dispatcher来更新UI
                    if (_wpfInstance.Dispatcher.CheckAccess())
                    {
                        // 当前线程是UI线程，直接更新
                        textBoxByName.Text = value;
                    }
                    else
                    {
                        // 当前线程不是UI线程，使用Dispatcher更新
                        _wpfInstance.Dispatcher.Invoke(() =>
                        {
                            textBoxByName.Text = value;
                        });
                    }
                    return;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"设置WPF界面TextBox值时出错: {ex.Message}");
            }
        }
    }

}
/// <summary>
/// 辅助类，用于处理TextBox的值
/// </summary>
public static class TextBoxValueHelper
{
    /// <summary>
    /// 获取TextBox的值，如果为空则返回Tag属性值作为默认值
    /// </summary>
    public static string GetTextBoxValue(System.Windows.Controls.TextBox textBox)
    {
        if (textBox == null)
            return string.Empty;


        if (!string.IsNullOrWhiteSpace(textBox.Text))
            return textBox.Text.Trim();

        return textBox.Tag?.ToString() ?? string.Empty;
    }

    /// <summary>
    /// 获取TextBox的数值
    /// </summary>
    public static double GetNumericValueOrDefault(System.Windows.Controls.TextBox textBox, double defaultValue = 0)
    {
        string value = GetTextBoxValue(textBox);
        if (double.TryParse(value, out double result))
        {
            return result;
        }
        return defaultValue;
    }

    /// <summary>
    /// 获取TextBox的整数值  
    /// </summary>
    public static int GetIntTextBoxValue(System.Windows.Controls.TextBox textBox, int defaultValue = 0)
    {
        string value = GetTextBoxValue(textBox);
        if (int.TryParse(value, out int result))
        {
            return result;
        }
        return defaultValue;
    }

}


