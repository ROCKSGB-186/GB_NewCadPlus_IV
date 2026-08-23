using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using GB_NewCadPlus_IV.Helpers;

namespace GB_NewCadPlus_IV.Views
{
    /// <summary>
    /// 插入图元前的属性编辑窗口。
    /// </summary>
    public partial class InsertGraphicPropertyWindow : Window
    {
        /// <summary>
        /// 动态属性行集合，供 DataGrid 显示和编辑。
        /// </summary>
        public ObservableCollection<InsertGraphicPropertyRow> Properties { get; }
            = new ObservableCollection<InsertGraphicPropertyRow>();

        /// <summary>
        /// 构造属性编辑窗口。
        /// </summary>
        public InsertGraphicPropertyWindow(
            IDictionary<string, string> properties,
            string? title = null)
        {
            InitializeComponent();
            DataContext = this;
            Title = string.IsNullOrWhiteSpace(title) ? "编辑图元属性" : $"编辑图元属性 - {title}";

            // 同一个属性可能同时存在原始键和归一化键，只显示一次，避免用户重复编辑。
            var usedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var property in properties ?? new Dictionary<string, string>())
            {
                string name = property.Key?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(name) || !usedKeys.Add(name)) continue;

                // 根据属性 Tag 标记连接方式行，XAML 将该行显示为下拉框。
                var row = new InsertGraphicPropertyRow(name, property.Value ?? string.Empty)
                {
                    IsConnectionType = IsConnectionTypeName(name)
                };
                row.PropertyChanged += PropertyRow_PropertyChanged;
                Properties.Add(row);
            }

            // 页面首次打开时也按照当前连接方式初始化法兰数量和螺栓数量。
            UpdateFlangeAndBoltQuantities();
        }

        /// <summary>
        /// 返回用户最终确认的属性值。
        /// </summary>
        public Dictionary<string, string> GetEditedProperties()
        {
            // 强制提交当前正在编辑的单元格，避免用户未按回车时最后一次修改丢失。
            PropertiesGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
            PropertiesGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);

            return Properties
                .Where(item => !string.IsNullOrWhiteSpace(item.Name))
                .GroupBy(item => item.Name.Trim(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Last().Value ?? string.Empty, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 连接方式或螺栓孔数量改变时，实时更新法兰数量和螺栓数量。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PropertyRow_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (sender is not InsertGraphicPropertyRow row || e.PropertyName != nameof(InsertGraphicPropertyRow.Value)) return;

            // 连接方式和螺栓孔数量都是计算结果的输入项，任一项变化都要重新计算。
            if (!row.IsConnectionType && !string.Equals(row.Name, "BOLT_HOLES", StringComparison.OrdinalIgnoreCase)) return;

            UpdateFlangeAndBoltQuantities();
        }

        /// <summary>
        /// 按连接方式和螺栓孔数量计算法兰数量及螺栓数量。
        /// </summary>
        private void UpdateFlangeAndBoltQuantities()
        {
            // 读取当前页面中的连接方式和螺栓孔数量。
            string connectionType = FindConnectionTypeRow()?.Value?.Trim() ?? string.Empty;
            int boltHoles = ParseIntegerOrZero(FindRow("BOLT_HOLES")?.Value);

            // 法兰和对夹为两侧法兰，单侧法兰为一侧法兰，其余连接方式没有法兰。
            int flangeQuantity = connectionType switch
            {
                "法兰" => 2,
                "对夹" => 2,
                "单侧法兰" => 1,
                _ => 0
            };

            // 螺栓总数统一按照法兰数量乘以每个法兰的螺栓孔数量计算。
            int boltQuantity = flangeQuantity * boltHoles;

            // 更新页面中的计算结果行；缺少行时保持现有属性结构不变。
            InsertGraphicPropertyRow? flangeQuantityRow = FindRow("FLG_QTY");
            InsertGraphicPropertyRow? boltQuantityRow = FindRow("BOLT_QTY");
            if (flangeQuantityRow != null) flangeQuantityRow.Value = flangeQuantity.ToString();
            if (boltQuantityRow != null) boltQuantityRow.Value = boltQuantity.ToString();
        }

        /// <summary>
        /// 查找当前页面中的连接方式属性行。
        /// </summary>
        private InsertGraphicPropertyRow? FindConnectionTypeRow()
        {
            return Properties.FirstOrDefault(row => IsConnectionTypeName(row.Name));
        }
        /// <summary>
        /// 根据属性名查找对应的属性行。
        /// </summary>
        /// <param name="name">属性名</param>
        /// <returns>对应的属性行，如果未找到则返回 null</returns>
        private InsertGraphicPropertyRow? FindRow(string name)
        {
            return Properties.FirstOrDefault(row => string.Equals(row.Name, name, StringComparison.OrdinalIgnoreCase));
        }
        /// <summary>
        /// 尝试将字符串解析为整数，如果解析失败则返回 0。
        /// </summary>
        /// <param name="value">要解析的字符串值</param>
        /// <returns>解析得到的整数值，如果解析失败则返回 0</returns>
        private static int ParseIntegerOrZero(string? value)
        {
            if (int.TryParse(value?.Trim(), out int result)) return Math.Max(0, result);
            if (double.TryParse(value?.Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double number))
            {
                return Math.Max(0, (int)Math.Round(number));
            }
            return 0;
        }
        /// <summary>
        /// 判断属性名是否表示连接方式。
        /// </summary>
        /// <param name="name">属性名</param>
        /// <returns>如果属性名表示连接方式则返回 true，否则返回 false</returns>
        private static bool IsConnectionTypeName(string name)
        {
            return string.Equals(name, "CONN_TYPE", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "DNCONN_TYPE", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "连接方式", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "连接形式", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 取消本次插入并关闭窗口。
        /// </summary>
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        /// <summary>
        /// 确认属性并继续插入。
        /// </summary>
        private void BtnInsert_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }

    /// <summary>
    /// 属性编辑表格的一行。
    /// </summary>
    public sealed class InsertGraphicPropertyRow : INotifyPropertyChanged
    {
        private string _value;

        public InsertGraphicPropertyRow(string name, string value)
        {
            Name = name;
            // 创建属性行时立即根据 Tag 生成中文注释，确保 XAML 绑定时已有实际内容。
            Comment = PropertyTagCommentDictionary.GetComment(name);
            _value = value;
        }
        /// <summary>
        /// 属性名，通常为大写字母和下划线组成的标识符。
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Tag 对应的中文业务注释，仅用于页面显示，不参与属性回写。
        /// </summary>
        public string Comment { get; }
        /// <summary>
        /// 属性值，用户可以在 DataGrid 中编辑。
        /// </summary>
        public bool IsConnectionType { get; set; }
        /// <summary>
        /// 连接方式的可选值列表，供下拉框选择。
        /// </summary>
        public IReadOnlyList<string> Options { get; } = new[] { "单侧法兰", "法兰", "对夹", "螺纹", "焊接", "其它" };
        /// <summary>
        /// 属性值，用户可以在 DataGrid 中编辑。
        /// </summary>
        public string Value
        {
            get => _value;
            set
            {
                if (string.Equals(_value, value, StringComparison.Ordinal)) return;
                _value = value ?? string.Empty;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Value)));
            }
        }
        /// <summary>
        /// 当属性值发生变化时触发的事件，供外部监听以更新相关逻辑。
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;
    }
}
