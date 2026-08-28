using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Models;

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

        // 保存本次窗口使用的 S/L 螺栓规范结果，连接方式切换时不重复请求服务器。
        private readonly Dictionary<string, BoltStandardMatchResponse> _boltStandardResponses;
        private readonly bool _isStandaloneFlangeOrBlindPlate;
        private int _boltRequestVersion;
        private int _flangeRequestVersion;

        /// <summary>
        /// 构造属性编辑窗口。
        /// </summary>
        public InsertGraphicPropertyWindow(
            IDictionary<string, string> properties,
            string? title = null,
            IReadOnlyDictionary<string, BoltStandardMatchResponse>? boltStandardResponses = null,
            bool isStandaloneFlangeOrBlindPlate = false)
        {
            InitializeComponent();
            DataContext = this;
            Title = string.IsNullOrWhiteSpace(title) ? "编辑图元属性" : $"编辑图元属性 - {title}";
            _isStandaloneFlangeOrBlindPlate = isStandaloneFlangeOrBlindPlate;
            _boltStandardResponses = new Dictionary<string, BoltStandardMatchResponse>(StringComparer.OrdinalIgnoreCase);
            if (boltStandardResponses != null)
            {
                foreach (KeyValuePair<string, BoltStandardMatchResponse> item in boltStandardResponses)
                    _boltStandardResponses[item.Key] = item.Value;
            }

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

            // 页面首次打开时按照当前连接方式初始化法兰、螺栓数量和螺栓长度。
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
            if (!row.IsConnectionType && !IsBoltHolesName(row.Name)) return;

            UpdateFlangeAndBoltQuantities();

            if (row.IsConnectionType)
                _ = RefreshBoltStandardAsync();
        }

        private async void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            LogManager.Instance.LogInfo("[插入前属性][更新点击] 用户点击 UPDATE，开始提交当前编辑值。");
            PropertiesGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
            PropertiesGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);

            string connectionType = FindConnectionTypeRow()?.Value?.Trim() ?? string.Empty;
            string currentDn = FindRow("DN", "公称通径", "通径", "管径", "公称直径")?.Value?.Trim() ?? string.Empty;
            string currentPn = FindRow("PN", "公称压力", "压力等级")?.Value?.Trim() ?? string.Empty;
            LogManager.Instance.LogInfo(
                $"[插入前属性][更新参数] DN={currentDn}, PN={currentPn}, CONN_TYPE={connectionType}, " +
                $"FLG_STD={FindRow("FLG_STD", "DRAWINGNO.STANDARDNO")?.Value ?? string.Empty}, " +
                $"PropertyCount={Properties.Count}");
            if (connectionType.IndexOf("法兰", StringComparison.OrdinalIgnoreCase) < 0 &&
                connectionType.IndexOf("对夹", StringComparison.OrdinalIgnoreCase) < 0)
            {
                LogManager.Instance.LogWarning($"[插入前属性][更新跳过] 连接方式不是法兰/对夹：CONN_TYPE={connectionType}");
                System.Windows.MessageBox.Show(this, "当前连接方式不是法兰或对夹，无需查询法兰规范。", "更新规范", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return;
            }

            string dn = currentDn;
            string pn = currentPn;
            if (string.IsNullOrWhiteSpace(dn) || string.IsNullOrWhiteSpace(pn))
            {
                System.Windows.MessageBox.Show(this, "请先填写 DN 和 PN，再点击更新。", "更新规范", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            UPDATE.IsEnabled = false;
            try
            {
                // DN 已变化时，S/L 螺栓规范必须按当前 DN 重新查询，不能继续使用打开窗口时的缓存。
                _boltStandardResponses.Clear();
                bool refreshed = await RefreshFlangeStandardAsync();
                LogManager.Instance.LogInfo($"[插入前属性][更新完成] DN={dn}, PN={pn}, Success={refreshed}");
                if (!refreshed)
                {
                    System.Windows.MessageBox.Show(this, "法兰规范未匹配成功，请检查 DN、PN 或法兰标准。", "更新规范", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                }
            }
            finally
            {
                UPDATE.IsEnabled = true;
            }
        }

        private async Task<bool> RefreshFlangeStandardAsync()
        {
            string connectionType = FindConnectionTypeRow()?.Value?.Trim() ?? string.Empty;
            if (connectionType.IndexOf("法兰", StringComparison.OrdinalIgnoreCase) < 0 &&
                connectionType.IndexOf("对夹", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }

            string dn = FindRow("DN", "公称通径", "通径", "管径", "公称直径")?.Value?.Trim() ?? string.Empty;
            string pn = FindRow("PN", "公称压力", "压力等级")?.Value?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(dn) || string.IsNullOrWhiteSpace(pn)) return false;

            int requestVersion = Interlocked.Increment(ref _flangeRequestVersion);
            try
            {
                StandardApiService api = new StandardApiService();
                FlangeStandardMatchRequest request = new FlangeStandardMatchRequest
                {
                    FamilyCode = FindRow("FAMILY_CODE")?.Value ?? "FLANGE",
                    SeriesCode = FindRow("SERIES_CODE")?.Value ?? string.Empty,
                    StandardNumber = FindRow("FLG_STD", "DRAWINGNO.STANDARDNO")?.Value ?? string.Empty,
                    TableNumber = FindRow("TABLE_NUMBER")?.Value ?? string.Empty,
                    DN = dn,
                    PN = pn,
                    Series = FindRow("SERIES", "钢管系列")?.Value ?? "Ⅰ系列",
                    ConnectionMode = connectionType,
                    FlangeType = FindRow("FLG_TYPE")?.Value ?? string.Empty,
                    FaceType = FindRow("FACE_TYPE")?.Value ?? string.Empty
                };
                LogManager.Instance.LogInfo(
                    $"[插入前属性][法兰请求] DN={request.DN}, PN={request.PN}, StandardNumber={request.StandardNumber}, " +
                    $"SeriesCode={request.SeriesCode}, Series={request.Series}, ConnectionMode={request.ConnectionMode}");
                FlangeStandardMatchResponse response = await api.MatchFlangeAsync(request).ConfigureAwait(true);

                if ((response == null || !response.Success) && !string.IsNullOrWhiteSpace(request.StandardNumber))
                {
                    FlangeStandardMatchRequest fallbackRequest = new FlangeStandardMatchRequest
                    {
                        FamilyCode = request.FamilyCode,
                        SeriesCode = request.SeriesCode,
                        StandardNumber = string.Empty,
                        TableNumber = string.Empty,
                        DN = request.DN,
                        PN = request.PN,
                        Series = request.Series,
                        ConnectionMode = request.ConnectionMode,
                        FlangeType = string.Empty,
                        FaceType = string.Empty
                    };
                    LogManager.Instance.LogInfo(
                        $"[插入前属性][法兰回退请求] DN={fallbackRequest.DN}, PN={fallbackRequest.PN}, Series={fallbackRequest.Series}");
                    response = await api.MatchFlangeAsync(fallbackRequest).ConfigureAwait(true);
                    LogManager.Instance.LogInfo(
                        $"[插入前属性][法兰回退结果] Success={response?.Success}, MatchCount={response?.MatchCount}, Message={response?.Message}");
                }

                if (requestVersion != _flangeRequestVersion || response?.Success != true)
                {
                    LogManager.Instance.LogWarning(
                        $"[插入前属性][法兰更新失败] DN={dn}, PN={pn}, Success={response?.Success}, Message={response?.Message}");
                    return false;
                }

                int updatedCount = 0;
                foreach (KeyValuePair<string, string> item in response.Attributes ?? new Dictionary<string, string>())
                {
                    InsertGraphicPropertyRow? target = FindRow(item.Key);
                    if (target != null)
                    {
                        string oldValue = target.Value;
                        target.Value = item.Value ?? string.Empty;
                        updatedCount++;
                        LogManager.Instance.LogInfo(
                            $"[插入前属性][字段回写] Tag={target.Name}, OldValue={oldValue}, NewValue={target.Value}");
                    }
                    else
                    {
                        InsertGraphicPropertyRow newRow = new InsertGraphicPropertyRow(item.Key, item.Value ?? string.Empty)
                        {
                            IsConnectionType = IsConnectionTypeName(item.Key)
                        };
                        newRow.PropertyChanged += PropertyRow_PropertyChanged;
                        Properties.Add(newRow);
                        updatedCount++;
                        LogManager.Instance.LogInfo(
                            $"[插入前属性][字段新增] Tag={item.Key}, NewValue={item.Value ?? string.Empty}");
                    }
                }

                UpdateFlangeAndBoltQuantities();
                await RefreshBoltStandardAsync();
                LogManager.Instance.LogInfo($"插入前窗口法兰规范已刷新：DN={dn}，PN={pn}，服务器属性数量={response.Attributes?.Count ?? 0}，页面更新数量={updatedCount}");
                return true;
            }
            catch (Exception exception)
            {
                LogManager.Instance.LogWarning($"插入前窗口法兰规范刷新失败：DN={dn}，PN={pn}，错误={exception.Message}");
                return false;
            }
        }

        private async Task RefreshBoltStandardAsync()
        {
            string connectionType = FindConnectionTypeRow()?.Value?.Trim() ?? string.Empty;
            string shortCode = connectionType.IndexOf("对夹", StringComparison.OrdinalIgnoreCase) >= 0
                ? "L"
                : connectionType.IndexOf("法兰", StringComparison.OrdinalIgnoreCase) >= 0
                    ? "S"
                    : string.Empty;
            InsertGraphicPropertyRow? lengthRow = FindRow("BOLT_LENGTH", "BOLTLENGTH");
            if (lengthRow == null) return;

            int requestVersion = Interlocked.Increment(ref _boltRequestVersion);
            if (string.IsNullOrWhiteSpace(shortCode))
            {
                lengthRow.Value = string.Empty;
                return;
            }

            if (!_boltStandardResponses.TryGetValue(shortCode, out BoltStandardMatchResponse? response) || response?.Success != true)
            {
                string dn = FindRow("DN")?.Value?.Trim() ?? string.Empty;
                string pn = FindRow("PN")?.Value?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(dn) || string.IsNullOrWhiteSpace(pn)) return;

                string standardNumber = shortCode == "S"
                    ? "LSLM-S 9124.1-2019"
                    : "LSLM-L 9124.1-2019";
                StandardApiService api = new StandardApiService();
                response = await api.MatchBoltAsync(new BoltStandardMatchRequest
                {
                    FamilyCode = "FLANGE",
                    SeriesCode = "PLATE_WELD",
                    StandardNumber = standardNumber,
                    DN = dn,
                    PN = pn,
                    Short = shortCode
                }).ConfigureAwait(true);
                _boltStandardResponses[shortCode] = response ?? new BoltStandardMatchResponse();
            }

            if (requestVersion != _boltRequestVersion) return;
            string length = FindBoltValue(response?.Attributes ?? new Dictionary<string, string>(),
                "LENGTH", "Length", "长度", "BOLT_LENGTH", "BOLTLENGTH", "螺栓长度");
            lengthRow.Value = length;
            LogManager.Instance.LogInfo($"插入前窗口螺栓长度已刷新：CONN_TYPE={connectionType}，SHORT={shortCode}，LENGTH={length}");
        }

        /// <summary>
        /// 按连接方式和螺栓孔数量计算法兰数量及螺栓数量。
        /// </summary>
        private void UpdateFlangeAndBoltQuantities()
        {
            // 读取当前页面中的连接方式和螺栓孔数量。
            string connectionType = FindConnectionTypeRow()?.Value?.Trim() ?? string.Empty;
            int boltHoles = ParseIntegerOrZero(FindRow("BOLT_HOLES", "BOLTHOLES", "螺栓孔数", "螺栓孔数量")?.Value);

            // “法兰”表示两侧各有一个法兰；“单侧法兰”和“对夹”均只统计一个法兰。
            int flangeQuantity = GetFlangeQuantity(connectionType, _isStandaloneFlangeOrBlindPlate);

            // 螺栓数量按法兰数量计算：法兰为两侧法兰，单侧法兰和对夹为单侧法兰。
            int boltQuantity = boltHoles * flangeQuantity;

            // 更新页面中的计算结果行；缺少行时保持现有属性结构不变。
            InsertGraphicPropertyRow? flangeQuantityRow = FindRow("FLG_QTY", "FLGQTY");
            InsertGraphicPropertyRow? boltQuantityRow = FindRow("BOLT_QTY", "BOLTQTY");
            if (flangeQuantityRow != null) flangeQuantityRow.Value = flangeQuantity.ToString();
            if (boltQuantityRow != null) boltQuantityRow.Value = boltQuantity.ToString();

            UpdateBoltStandardValues(connectionType, boltHoles);
        }

        /// <summary>
        /// 按连接方式计算法兰数量。
        /// </summary>
        private static int GetFlangeQuantity(string connectionType, bool isStandaloneFlangeOrBlindPlate)
        {
            if (isStandaloneFlangeOrBlindPlate) return 1;

            string value = (connectionType ?? string.Empty).Replace(" ", string.Empty).Replace("　", string.Empty);
            if (value.Equals("法兰", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("法兰连接", StringComparison.OrdinalIgnoreCase)) return 2;
            if (value.IndexOf("单侧法兰", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("对夹", StringComparison.OrdinalIgnoreCase) >= 0) return 1;
            return 0;
        }

        /// <summary>
        /// 判断属性行是否为螺栓孔数量，兼容历史 Tag 写法。
        /// </summary>
        private static bool IsBoltHolesName(string name)
        {
            return string.Equals(name, "BOLT_HOLES", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "BOLTHOLES", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "螺栓孔数", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "螺栓孔数量", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 根据连接方式从本次窗口缓存中选择 S/L 螺栓规范并更新长度、数量。
        /// </summary>
        private void UpdateBoltStandardValues(string connectionType, int boltHoles)
        {
            string shortCode = connectionType.Contains("对夹", StringComparison.OrdinalIgnoreCase)
                ? "L"
                : connectionType.Contains("法兰", StringComparison.OrdinalIgnoreCase)
                    ? "S"
                    : string.Empty;

            InsertGraphicPropertyRow? lengthRow = FindRow("BOLT_LENGTH", "BOLTLENGTH");
            InsertGraphicPropertyRow? quantityRow = FindRow("BOLT_QTY", "BOLTQTY");
            if (string.IsNullOrWhiteSpace(shortCode)) return;

            if (!_boltStandardResponses.TryGetValue(shortCode, out BoltStandardMatchResponse? response) ||
                response?.Success != true)
                return;

            string length = FindBoltValue(response.Attributes,
                "LENGTH", "Length", "长度", "BOLT_LENGTH", "BOLTLENGTH", "螺栓长度");
            string quantity = FindBoltValue(response.Attributes,
                "QUANTITY", "Quantity", "数量", "BOLT_QTY", "BOLTQTY", "螺栓数量");
            if (lengthRow != null) lengthRow.Value = length;

            if (string.IsNullOrWhiteSpace(length))
            {
                LogManager.Instance.LogWarning(
                    $"插入前窗口螺栓规范已命中但未返回 LENGTH：CONN_TYPE={connectionType}，SHORT={shortCode}，属性={string.Join(",", response.Attributes.Keys)}");
            }

            LogManager.Instance.LogInfo(
                $"插入前窗口螺栓规范切换：CONN_TYPE={connectionType}，SHORT={shortCode}，LENGTH={length}，规范QUANTITY={quantity}，BOLT_HOLES={boltHoles}，BOLT_QTY={boltHoles * GetFlangeQuantity(connectionType, _isStandaloneFlangeOrBlindPlate)}");
        }

        /// <summary>
        /// 从服务器返回的动态字段中读取指定属性。
        /// </summary>
        private static string FindBoltValue(IDictionary<string, string> attributes, params string[] names)
        {
            foreach (string name in names)
            {
                foreach (KeyValuePair<string, string> item in attributes)
                {
                    if (string.Equals(item.Key, name, StringComparison.OrdinalIgnoreCase))
                        return item.Value?.Trim() ?? string.Empty;
                }
            }

            return string.Empty;
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
        private InsertGraphicPropertyRow? FindRow(params string[] names)
        {
            return Properties.FirstOrDefault(row => names.Any(name =>
                string.Equals(NormalizePropertyName(row.Name), NormalizePropertyName(name), StringComparison.OrdinalIgnoreCase)));
        }

        private static string NormalizePropertyName(string name)
        {
            string decoded = PipelineCadPropertyKeyHelper.Decode(name ?? string.Empty);
            return new string(decoded
                .Trim()
                .ToUpperInvariant()
                .Where(char.IsLetterOrDigit)
                .ToArray());
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
                || string.Equals(name, "CONNTYPE", StringComparison.OrdinalIgnoreCase)
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
