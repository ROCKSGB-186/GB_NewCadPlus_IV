using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using MessageBox = System.Windows.MessageBox;
using WindowVisibility = System.Windows.Visibility;
using WpfBinding = System.Windows.Data.Binding;

namespace GB_NewCadPlus_IV.DisplayPages
{
    /// <summary>
    /// 统一规范导入预览窗口。
    /// 固定 JSON 和动态 Excel 都使用本窗口，动态 Excel 的列由服务器模板生成。
    /// </summary>
    public partial class StandardImportPreviewWindow : Window, INotifyPropertyChanged
    {
        private int _selectedFilterIndex;
        private bool _isDynamicPreview;
        private DynamicStandardPreviewClientResponse _dynamicPreview;
        private StandardApiService _dynamicApiService;
        private string _dynamicOperatorName = string.Empty;
        private long? _dynamicSeriesId;
        private List<UnifiedPreviewRowViewModel> _unifiedRows = new List<UnifiedPreviewRowViewModel>();
        private StandardImportSeriesClient _dynamicSeries = new StandardImportSeriesClient();
        private bool _useMergeStrategy;
        private long? _dynamicCategoryId;
        private readonly ObservableCollection<StandardDocumentClient> _baseStandardCandidates = new ObservableCollection<StandardDocumentClient>();
        private long? _dynamicStandardDocumentId;
        private bool _baseSeriesCandidatesLoaded;
        private System.Windows.Controls.TextBox _baseStandardTextBox;

        /// <summary>构造固定 JSON 预览窗口。</summary>
        public StandardImportPreviewWindow(StandardImportPreviewClientResponse preview)
        {
            Preview = preview ?? throw new ArgumentNullException(nameof(preview));
            Preview.Series = Preview.Series ?? new StandardImportSeriesClient();
            InitializeComponent();
            InitializeFixedPreview();
        }

        /// <summary>构造动态 Excel 预览窗口，并统一使用 StandardImportPreviewWindow 页面。</summary>
        public StandardImportPreviewWindow(
            DynamicStandardPreviewClientResponse preview,
            string sourceFileName,
            StandardApiService apiService,
            string operatorName,
            long? seriesId = null,
            long? categoryId = null)
        {
            _dynamicPreview = preview ?? throw new ArgumentNullException(nameof(preview));
            _dynamicApiService = apiService ?? throw new ArgumentNullException(nameof(apiService));
            _dynamicOperatorName = operatorName ?? string.Empty;
            _dynamicSeriesId = seriesId;
            _dynamicCategoryId = categoryId;
            DynamicSourceFileName = sourceFileName ?? string.Empty;
            // 根据上传文件名自动分解规范注释名称、规范代号、表号和型号。
            _dynamicSeries = ParseSeriesFromFileName(DynamicSourceFileName);
            _isDynamicPreview = true;
            InitializeComponent();
            InitializeDynamicPreview();
        }

        /// <summary>固定预览响应。</summary>
        public StandardImportPreviewClientResponse Preview { get; }

        /// <summary>动态预览响应。</summary>
        public DynamicStandardPreviewClientResponse DynamicPreview => _dynamicPreview;

        /// <summary>固定预览的规范系列信息。</summary>
        public StandardImportSeriesClient Series => _isDynamicPreview
            ? _dynamicSeries
            : Preview?.Series ?? new StandardImportSeriesClient();

        /// <summary>动态预览的原始文件名。</summary>
        public string DynamicSourceFileName { get; }

        /// <summary>基础规范下拉候选项。</summary>
        public ObservableCollection<StandardDocumentClient> BaseStandardCandidates => _baseStandardCandidates;

        /// <summary>统一显示原始文件名。</summary>
        public string SourceFileName => _isDynamicPreview
            ? DynamicSourceFileName
            : Preview?.SourceFileName ?? string.Empty;

        /// <summary>显示固定预览的组合规范名称。</summary>
        public string DisplayName => string.Join("_", new[]
            {
                Series.SeriesName,
                Series.StandardNumber,
                Series.TableNumber,
                Series.PressureRating
            }.Where(value => !string.IsNullOrWhiteSpace(value)));

        /// <summary>统一数据行筛选视图。</summary>
        public ICollectionView RowsView { get; private set; }

        /// <summary>错误、警告和总行数摘要。</summary>
        public string SummaryText => _isDynamicPreview
            ? $"共 {DynamicPreview?.Rows?.Count ?? 0} 行；错误 {DynamicPreview?.ErrorCount ?? 0} 行；警告 {DynamicPreview?.WarningCount ?? 0} 行"
            : $"共 {Preview?.RowCount ?? 0} 行；错误 {Preview?.ErrorCount ?? 0} 行；警告 {Preview?.WarningCount ?? 0} 行";

        /// <summary>动态模板状态文字。</summary>
        public string DynamicTemplateStatus => !_isDynamicPreview
            ? string.Empty
            : DynamicPreview.IsTemplateMatched ? "已匹配规范模板" : "未匹配规范模板";

        /// <summary>动态模板详细信息。</summary>
        public string DynamicTemplateDetail => !_isDynamicPreview
            ? string.Empty
            : DynamicPreview.Template == null
                ? DynamicPreview.Message ?? string.Empty
                : $"{DynamicPreview.Template.TemplateName}（{DynamicPreview.Template.TemplateCode}），版本 {DynamicPreview.Template.Version}，字段 {DynamicPreview.Columns?.Count ?? 0} 个";

        /// <summary>动态导入目标规范系列状态。</summary>
        public string DynamicTargetSeriesStatus => !_isDynamicPreview
            ? string.Empty
            : _dynamicSeriesId.GetValueOrDefault() > 0
                ? $"目标规范系列：{Series.SeriesName}_{Series.StandardNumber}（SeriesId={_dynamicSeriesId.Value}）；本次细分：{Series.TableNumber}_{Series.PressureRating}"
                : "目标规范系列：未指定，请返回并选择目标规范系列。";

        /// <summary>动态预览未映射的 Excel 表头。</summary>
        public string DynamicUnmappedHeaders => !_isDynamicPreview || DynamicPreview.UnmappedHeaders == null || DynamicPreview.UnmappedHeaders.Count == 0
            ? string.Empty
            : "未映射表头：" + string.Join("、", DynamicPreview.UnmappedHeaders);

        /// <summary>动态模板状态区域可见性。</summary>
        public WindowVisibility DynamicTemplateVisibility => _isDynamicPreview ? WindowVisibility.Visible : WindowVisibility.Collapsed;

        /// <summary>动态模板草稿/更新策略提示。</summary>
        public string DynamicUpdateDescription => !_isDynamicPreview
            ? string.Empty
            : !DynamicPreview.IsTemplateMatched
                ? "首次上传：系统已根据 Excel 表头生成模板草稿，确认导入后保存为模板。字段默认 TEXT、非必填。"
                : "再次上传：可选择新版替换，或选择唯一键后融合更新。旧版本不会删除。";

        /// <summary>是否显示融合选项。</summary>
        public WindowVisibility DynamicMergeVisibility => _isDynamicPreview && DynamicPreview.IsTemplateMatched
            ? WindowVisibility.Visible
            : WindowVisibility.Collapsed;

        /// <summary>当前动态更新策略。</summary>
        public string DynamicUpdateStrategy => _useMergeStrategy ? "MERGE" : "REPLACE";
        public bool IsReplaceStrategy
        {
            get => !_useMergeStrategy;
            set { if (value) _useMergeStrategy = false; }
        }
        public bool IsMergeStrategy
        {
            get => _useMergeStrategy;
            set { if (value) _useMergeStrategy = true; }
        }

        private void UpdateStrategyRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            RefreshWindowState();
        }

        /// <summary>
        /// 获取可编辑 ComboBox 内部的 TextBox，并监听用户手工输入的基础规范号。
        /// </summary>
        private void 所在基础规范ComBox_Loaded(object sender, RoutedEventArgs e)
        {
            _baseStandardTextBox = 所在基础规范ComBox.Template.FindName("PART_EditableTextBox", 所在基础规范ComBox) as System.Windows.Controls.TextBox;
            if (_baseStandardTextBox != null)
            {
                _baseStandardTextBox.TextChanged -= 基础规范号TextBox_TextChanged;
                _baseStandardTextBox.TextChanged += 基础规范号TextBox_TextChanged;
            }
        }

        /// <summary>
        /// 用户输入基础规范号时，清除已有基础规范 ID，并立即刷新确认按钮状态。
        /// </summary>
        private void 基础规范号TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (所在基础规范ComBox.SelectedItem != null
                && string.Equals(所在基础规范ComBox.Text?.Trim(), ((StandardDocumentClient)所在基础规范ComBox.SelectedItem).DisplayText, StringComparison.OrdinalIgnoreCase))
                return;

            _dynamicStandardDocumentId = null;
            _dynamicSeries.StandardNumber = 所在基础规范ComBox.Text?.Trim() ?? string.Empty;
            RefreshWindowState();
        }

        /// <summary>固定预览重名提示可见性。</summary>
        public WindowVisibility DuplicateVisibility => _isDynamicPreview || Preview?.DuplicateSeries == null ? WindowVisibility.Collapsed : WindowVisibility.Visible;

        /// <summary>固定预览重名描述。</summary>
        public string DuplicateDescription => Preview?.DuplicateSeries == null
            ? string.Empty
            : $"已有规范：{Preview.DuplicateSeries.SeriesName}-{Preview.DuplicateSeries.StandardNumber}-{Preview.DuplicateSeries.TableNumber}-{Preview.DuplicateSeries.PressureRating}；已有有效记录 {Preview.DuplicateSeries.RecordCount} 条。";

        /// <summary>警告确认区域可见性。</summary>
        public WindowVisibility WarningConfirmationVisibility => GetWarningCount() > 0 ? WindowVisibility.Visible : WindowVisibility.Collapsed;

        /// <summary>是否允许提交警告数据。</summary>
        public bool AllowWarnings => GetWarningCount() > 0 && WarningConfirmedCheckBox != null && WarningConfirmedCheckBox.IsChecked == true;

        /// <summary>固定导入确认请求。</summary>
        public StandardImportCommitClientRequest ConfirmedRequest => new StandardImportCommitClientRequest
        {
            BatchId = Preview?.BatchId ?? string.Empty,
            AllowWarnings = AllowWarnings,
            Series = Series,
            DuplicateStrategy = StandardImportDuplicateStrategyClient.NewImport
        };

        /// <summary>页面校验提示。</summary>
        public string ValidationMessage { get; private set; } = string.Empty;

        /// <summary>窗口属性变化通知。</summary>
        public event PropertyChangedEventHandler PropertyChanged;

        private void InitializeFixedPreview()
        {
            _unifiedRows = (Preview.Rows ?? new List<StandardImportPreviewRowClient>())
                .Select(row => UnifiedPreviewRowViewModel.FromFixed(row))
                .ToList();
            BuildFixedColumns();
            BindRows();
            DataContext = this;
            LoadBaseSeriesCandidates();
            RefreshWindowState();
            LogManager.Instance.LogInfo($"打开统一规范导入预览窗口：固定模式，文件={SourceFileName}，批次={Preview.BatchId}，行数={Preview.RowCount}。");
        }

        /// <summary>
        /// 从已经加载的规范管理树中筛选基础系列，避免客户端直接访问数据库。
        /// </summary>
        private void LoadBaseSeriesCandidates()
        {
            if (!_isDynamicPreview || _dynamicApiService == null)
                return;

            // 候选数据由异步初始化事件加载，避免在窗口构造函数中阻塞 UI 线程。
            Loaded += async (_, _) =>
            {
                try
                {
                    if (_baseSeriesCandidatesLoaded)
                        return;

                    StandardManagementTreeClientResponse response = await _dynamicApiService
                        .GetManagementTreeAsync()
                        .ConfigureAwait(true);
                    IEnumerable<StandardDocumentClient> candidates = (response.Documents ?? new List<StandardDocumentClient>())
                        .OrderBy(document => document.StandardNumber)
                        .ThenBy(document => document.StandardName);
                    LogManager.Instance.LogInfo($"基础规范候选加载完成：Documents={response.Documents?.Count ?? 0}，Categories={response.Categories?.Count ?? 0}，Series={response.Series?.Count ?? 0}。");
                    _baseStandardCandidates.Clear();
                    foreach (StandardDocumentClient candidate in candidates)
                        _baseStandardCandidates.Add(candidate);
                    _baseSeriesCandidatesLoaded = true;

                    所在基础规范ComBox.ItemsSource = _baseStandardCandidates;
                    所在基础规范ComBox.DisplayMemberPath = nameof(StandardDocumentClient.DisplayText);
                    所在基础规范ComBox.SelectedValuePath = nameof(StandardDocumentClient.Id);
                    StandardDocumentClient? selected = _baseStandardCandidates.FirstOrDefault(document =>
                        document.Id == _dynamicStandardDocumentId
                        || string.Equals(document.StandardNumber, Series.StandardNumber, StringComparison.OrdinalIgnoreCase));
                    if (selected != null)
                    {
                        所在基础规范ComBox.SelectedItem = selected;
                        _dynamicStandardDocumentId = selected.Id;
                    }
                    else
                        所在基础规范ComBox.Text = Series.StandardNumber;
                    RefreshWindowState();
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogError($"加载基础规范候选失败：{ex}");
                }
            };
        }

        /// <summary>选择已有基础规范号时同步基础规范 ID，不覆盖当前上传的细分规范名称。</summary>
        private void 所在基础规范ComBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (所在基础规范ComBox.SelectedItem is StandardDocumentClient selected)
            {
                _dynamicStandardDocumentId = selected.Id;
                _dynamicSeries.StandardNumber = selected.StandardNumber;
                _dynamicSeries.FamilyCode = selected.FamilyCode;
                RefreshWindowState();
                LogManager.Instance.LogInfo($"已选择基础规范号：StandardDocumentId={selected.Id}，规范号={selected.StandardNumber}。");
            }
        }

        /// <summary>
        /// 用户手工输入基础规范时，清除旧的系列 ID，交由服务器按基础身份创建或查找。
        /// </summary>
        private void 所在基础规范ComBox_DropDownClosed(object sender, EventArgs e)
        {
            if (所在基础规范ComBox.SelectedItem == null)
            {
                _dynamicStandardDocumentId = null;
                _dynamicSeries.StandardNumber = 所在基础规范ComBox.Text?.Trim() ?? string.Empty;
            }
            RefreshWindowState();
        }

        private void InitializeDynamicPreview()
        {
            _unifiedRows = (DynamicPreview.Rows ?? new List<DynamicStandardPreviewRowClient>())
                .Select(row => UnifiedPreviewRowViewModel.FromDynamic(row))
                .ToList();
            BuildDynamicColumns();
            BindRows();
            DataContext = this;
            LoadBaseSeriesCandidates();
            RefreshWindowState();
            LogManager.Instance.LogInfo($"打开统一规范导入预览窗口：动态模式，文件={SourceFileName}，模板匹配={DynamicPreview.IsTemplateMatched}，行数={DynamicPreview.Rows?.Count ?? 0}。");
        }

        /// <summary>
        /// 按“规范注释名称_规范代号_表号_型号.xlsx”解析动态规范元数据。
        /// </summary>
        private static StandardImportSeriesClient ParseSeriesFromFileName(string sourceFileName)
        {
            string name = Path.GetFileNameWithoutExtension(sourceFileName ?? string.Empty);
            string[] parts = name.Split(new[] { '_' }, StringSplitOptions.None);
            return new StandardImportSeriesClient
            {
                SeriesName = parts.Length > 0 ? parts[0].Trim() : string.Empty,
                StandardNumber = parts.Length > 1 ? parts[1].Trim() : string.Empty,
                TableNumber = parts.Length > 2 ? parts[2].Trim() : string.Empty,
                PressureRating = parts.Length > 3 ? string.Join("_", parts.Skip(3)).Trim() : string.Empty
            };
        }
        /// <summary>
        /// 构建固定规范的列定义。
        /// </summary>
        private void BuildFixedColumns()
        {
            AddTextColumn("行号", "RowNumber", 55);
            AddTextColumn("DN", "Values[DN]", 55);
            AddTextColumn("PN", "Values[PN]", 65);
            AddTextColumn("管外径 I", "Values[PIPE_OUTER_DIAMETER_I]", 100);
            AddTextColumn("管外径 II", "Values[PIPE_OUTER_DIAMETER_II]", 100);
            AddTextColumn("法兰外径", "Values[FLANGE_OUTER_DIAMETER]", 75);
            AddTextColumn("螺栓中心圆", "Values[BOLT_CIRCLE_DIAMETER]", 85);
            AddTextColumn("孔径", "Values[BOLT_HOLE_DIAMETER]", 70);
            AddTextColumn("螺栓数", "Values[BOLT_COUNT]", 60);
            AddTextColumn("螺栓规格", "Values[BOLT_SPECIFICATION]", 75);
            AddTextColumn("法兰厚度", "Values[FLANGE_THICKNESS]", 70);
            AddTextColumn("状态", "ValidationStatus", 60);
            AddTextColumn("错误", "ErrorText", 80);
            AddTextColumn("警告", "WarningText", 80);
        }
        /// <summary>
        /// 构建动态规范的列定义。
        /// </summary>
        private void BuildDynamicColumns()
        {
            AddTextColumn("行号", "RowNumber", 80);
            foreach (DynamicStandardPreviewColumnClient column in DynamicPreview.Columns ?? new List<DynamicStandardPreviewColumnClient>())
            {
                string key = string.IsNullOrWhiteSpace(column.FieldCode) ? column.Header : column.FieldCode;
                string name = string.IsNullOrWhiteSpace(column.FieldName) ? column.Header : column.FieldName;
                AddDynamicTextColumn(name, key, column, 105);
            }
            AddTextColumn("校验信息", "ValidationMessage", 150);
        }

        /// <summary>
        /// 创建动态列：标题只显示字段名称，字段状态通过颜色表达。
        /// </summary>
        private void AddDynamicTextColumn(string header, string key, DynamicStandardPreviewColumnClient column, double width)
        {
            SolidColorBrush headerBackground = column.IsMapped
                ? (column.IsRequired ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(221, 235, 247)) : new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 242, 204)))
                : new SolidColorBrush(System.Windows.Media.Color.FromRgb(230, 230, 250));
            SolidColorBrush headerForeground = column.IsMapped ? Brushes.DarkSlateGray : Brushes.DarkViolet;

            FrameworkElementFactory headerBorder = new FrameworkElementFactory(typeof(Border));
            headerBorder.SetValue(Border.BackgroundProperty, headerBackground);
            headerBorder.SetValue(Border.PaddingProperty, new Thickness(6, 3, 6, 3));

            FrameworkElementFactory headerText = new FrameworkElementFactory(typeof(TextBlock));
            headerText.SetValue(TextBlock.TextProperty, header);
            headerText.SetValue(TextBlock.ForegroundProperty, headerForeground);
            headerText.SetValue(TextBlock.FontWeightProperty, FontWeights.Bold);
            headerText.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
            headerBorder.AppendChild(headerText);

            FrameworkElementFactory cellText = new FrameworkElementFactory(typeof(TextBlock));
            cellText.SetBinding(TextBlock.TextProperty, new WpfBinding("Values[" + key + "]"));
            cellText.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            cellText.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);

            FrameworkElementFactory cellBorder = new FrameworkElementFactory(typeof(Border));
            cellBorder.SetValue(Border.PaddingProperty, new Thickness(6, 5, 6, 5));
            cellBorder.AppendChild(cellText);

            PreviewDataGrid.Columns.Add(new DataGridTemplateColumn
            {
                Header = header,
                HeaderTemplate = new DataTemplate { VisualTree = headerBorder },
                CellTemplate = new DataTemplate { VisualTree = cellBorder },
                Width = width
            });
        }
        /// <summary>
        /// 创建固定列：标题和单元格都显示文本。
        /// </summary>
        /// <param name="header">列标题</param>
        /// <param name="bindingPath">绑定路径</param>
        /// <param name="width">列宽</param>
        private void AddTextColumn(string header, string bindingPath, double width)
        {
            PreviewDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = header,
                Binding = new WpfBinding(bindingPath),
                Width = width
            });
        }
        /// <summary>
        /// 绑定统一数据行视图到 DataGrid，并设置筛选器。
        /// </summary>
        private void BindRows()
        {
            RowsView = CollectionViewSource.GetDefaultView(_unifiedRows);
            RowsView.Filter = FilterRow;
            PreviewDataGrid.ItemsSource = RowsView;
        }
        /// <summary>
        /// 获取当前预览的警告行数，动态和固定预览统一使用本方法。
        /// </summary>
        /// <returns>警告行数</returns>
        private int GetWarningCount() => _isDynamicPreview ? DynamicPreview?.WarningCount ?? 0 : Preview?.WarningCount ?? 0;
        /// <summary>
        /// 获取当前预览的错误行数，动态和固定预览统一使用本方法。
        /// </summary>
        /// <returns>错误行数</returns>
        private int GetErrorCount() => _isDynamicPreview ? DynamicPreview?.ErrorCount ?? 0 : Preview?.ErrorCount ?? 0;
        /// <summary>
        /// 当元数据文本框内容变化时，刷新窗口状态，包括显示名称、摘要、警告确认可见性和校验提示。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MetadataTextBox_TextChanged(object sender, TextChangedEventArgs e) => RefreshWindowState();
        /// <summary>
        /// 当行筛选组合框选择变化时，更新筛选器并刷新数据行视图，同时记录日志。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RowFilterComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (RowFilterComboBox == null || RowsView == null) return;
            _selectedFilterIndex = RowFilterComboBox.SelectedIndex;
            RowsView.Refresh();
            LogManager.Instance.LogInfo($"统一规范导入预览筛选变化：文件={SourceFileName}，筛选={_selectedFilterIndex}。");
        }
        /// <summary>
        /// 当警告确认复选框状态变化时，刷新窗口状态，包括校验提示和确认按钮可用性。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WarningConfirmedCheckBox_Changed(object sender, RoutedEventArgs e) => RefreshWindowState();
        /// <summary>
        /// 当用户点击确认按钮时，先检查是否允许确认，如果是动态预览则调用异步确认方法，否则直接设置 DialogResult 并关闭窗口。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isDynamicPreview)
            {
                await ConfirmDynamicAsync();
                return;
            }

            if (!CanConfirm())
            {
                MessageBox.Show(ValidationMessage, "无法确认导入", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            LogManager.Instance.LogInfo($"用户确认固定规范导入：文件={SourceFileName}，批次={Preview.BatchId}，规范名={DisplayName}。");
            DialogResult = true;
            Close();
        }
        
        /// <summary>
        /// 当用户确认动态规范导入时，先检查是否允许确认，如果允许则调用异步确认方法。
        /// </summary>
        /// <returns></returns>
        private async System.Threading.Tasks.Task ConfirmDynamicAsync()
        {
            if (!CanConfirm())
            {
                MessageBox.Show(ValidationMessage, "无法确认导入", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ConfirmButton.IsEnabled = false;
            try
            {
                DynamicStandardImportCommitClientResponse response = await _dynamicApiService.CommitDynamicImportAsync(
                    new DynamicStandardImportCommitClientRequest
                    {
                        BatchId = Guid.NewGuid().ToString("N"),
                        SeriesId = _dynamicSeriesId.GetValueOrDefault(),
                        StandardDocumentId = _dynamicStandardDocumentId,
                        SeriesName = Series.SeriesName,
                        BaseStandardNumber = string.IsNullOrWhiteSpace(所在基础规范ComBox.Text)
                            ? Series.StandardNumber
                            : 所在基础规范ComBox.Text.Trim(),
                        BaseStandardName = string.Empty,
                        StandardNumber = Series.StandardNumber,
                        SeriesCode = string.IsNullOrWhiteSpace(Series.SeriesCode) ? "PLATE_WELD" : Series.SeriesCode,
                        TableNumber = Series.TableNumber,
                        PressureRating = Series.PressureRating,
                        FamilyCode = string.IsNullOrWhiteSpace(Series.FamilyCode) ? "FLANGE" : Series.FamilyCode,
                        TemplateId = DynamicPreview.Template?.Id,
                        CategoryId = _dynamicCategoryId,
                        SourceFileName = SourceFileName,
                        AllowWarnings = AllowWarnings,
                        UpdateStrategy = DynamicUpdateStrategy,
                        ConfirmTemplateCreation = !DynamicPreview.IsTemplateMatched,
                        TemplateDraft = DynamicPreview.TemplateDraft,
                        UniqueKeyFields = DynamicPreview.CandidateUniqueKeyFields ?? new List<string>(),
                        Rows = DynamicPreview.Rows ?? new List<DynamicStandardPreviewRowClient>()
                    }, _dynamicOperatorName).ConfigureAwait(true);
                LogManager.Instance.LogInfo($"统一动态规范导入完成：批次={response.BatchId}，保存行数={response.SavedRowCount}，状态={response.Status}。");
                MessageBox.Show($"{response.Message}\n保存行数：{response.SavedRowCount}", "动态规范导入", MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = response.Success;
                Close();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"统一动态规范导入确认失败：{ex}");
                MessageBox.Show($"动态规范导入确认失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                ConfirmButton.IsEnabled = CanConfirm();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            LogManager.Instance.LogInfo($"用户取消统一规范导入：文件={SourceFileName}。");
            DialogResult = false;
            Close();
        }

        private bool FilterRow(object item)
        {
            UnifiedPreviewRowViewModel row = item as UnifiedPreviewRowViewModel;
            if (row == null) return false;
            switch (_selectedFilterIndex)
            {
                case 1: return !row.HasErrors && !row.HasWarnings;
                case 2: return row.HasWarnings;
                case 3: return row.HasErrors;
                default: return true;
            }
        }

        private void RefreshWindowState()
        {
            OnPropertyChanged(nameof(DisplayName));
            OnPropertyChanged(nameof(SummaryText));
            OnPropertyChanged(nameof(DynamicUpdateDescription));
            OnPropertyChanged(nameof(DynamicUpdateStrategy));
            OnPropertyChanged(nameof(WarningConfirmationVisibility));
            ValidationMessage = GetValidationMessage();
            OnPropertyChanged(nameof(ValidationMessage));
            if (ConfirmButton != null) ConfirmButton.IsEnabled = CanConfirm();
        }

        private bool CanConfirm() => string.IsNullOrWhiteSpace(GetValidationMessage());

        private string GetValidationMessage()
        {
            if (_isDynamicPreview)
            {
                if (!DynamicPreview.Success)
                    return DynamicPreview.Message ?? "动态规范预览失败。";
                string baseStandardNumber = 所在基础规范ComBox?.Text?.Trim() ?? Series.StandardNumber;
                if (string.IsNullOrWhiteSpace(baseStandardNumber))
                    return "请输入所在基础规范号，例如 GB/T 9124.1-2019。";
                if (string.IsNullOrWhiteSpace(Series.SeriesName))
                    return "细分规范名称不能为空。";
                if (GetErrorCount() > 0)
                    return "预览中存在错误行，请修正源文件后重新预览。";
                if (GetWarningCount() > 0 && !AllowWarnings)
                    return "预览中存在警告，请勾选确认后继续。";
                if (_useMergeStrategy && (DynamicPreview.CandidateUniqueKeyFields == null || DynamicPreview.CandidateUniqueKeyFields.Count == 0))
                    return "融合更新必须先选择至少一个唯一键字段。";
                return string.Empty;
            }

            if (string.IsNullOrWhiteSpace(Series.SeriesName) ||
                string.IsNullOrWhiteSpace(Series.StandardNumber) ||
                string.IsNullOrWhiteSpace(Series.TableNumber) ||
                string.IsNullOrWhiteSpace(Series.PressureRating))
                return "规范注释名称、规范代号、表号和型号不能为空。";
            if (GetErrorCount() > 0) return "预览中存在错误行，请修正源文件后重新选择并预览。";
            if (Preview.DuplicateSeries != null) return "当前系统正在补齐法兰记录的版本快照和恢复能力；为保护已有数据，暂不允许提交同名规范。";
            if (GetWarningCount() > 0 && !AllowWarnings) return "预览中存在警告，请核对后勾选确认。";
            return string.Empty;
        }

        private void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    /// <summary>统一固定/动态预览行，避免 XAML 绑定两套数据类型。</summary>
    internal sealed class UnifiedPreviewRowViewModel
    {
        public int RowNumber { get; private set; }
        public Dictionary<string, string> Values { get; private set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string ValidationStatus { get; private set; } = string.Empty;
        public string ErrorText { get; private set; } = string.Empty;
        public string WarningText { get; private set; } = string.Empty;
        public string ValidationMessage { get; private set; } = string.Empty;
        public bool HasErrors { get; private set; }
        public bool HasWarnings { get; private set; }

        public static UnifiedPreviewRowViewModel FromFixed(StandardImportPreviewRowClient row)
        {
            return new UnifiedPreviewRowViewModel
            {
                // 预览页面的行号从 0 开始显示，服务器返回的业务行号从 1 开始。
                RowNumber = Math.Max(0, row.RowNumber - 1),
                Values = row.Record == null ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : FixedRecordValues(row.Record),
                ValidationStatus = row.ValidationStatus,
                ErrorText = row.ErrorText,
                WarningText = row.WarningText,
                ValidationMessage = $"{row.ErrorText} {row.WarningText}".Trim(),
                HasErrors = row.HasErrors,
                HasWarnings = row.HasWarnings
            };
        }

        public static UnifiedPreviewRowViewModel FromDynamic(DynamicStandardPreviewRowClient row)
        {
            List<string> errors = row.Errors ?? new List<string>();
            List<string> warnings = row.Warnings ?? new List<string>();
            return new UnifiedPreviewRowViewModel
            {
                // 预览页面的行号从 0 开始显示，服务器返回的业务行号从 1 开始。
                RowNumber = Math.Max(0, row.RowNumber - 1),
                Values = row.Values ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
                ErrorText = string.Join("；", errors),
                WarningText = string.Join("；", warnings),
                ValidationMessage = string.Join("；", errors.Concat(warnings)),
                ValidationStatus = errors.Count > 0 ? "错误" : warnings.Count > 0 ? "警告" : "正常",
                HasErrors = errors.Count > 0,
                HasWarnings = warnings.Count > 0
            };
        }

        private static Dictionary<string, string> FixedRecordValues(StandardImportRecordClient record)
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["DN"] = record.DN ?? string.Empty,
                ["PN"] = record.PN ?? string.Empty,
                ["PIPE_OUTER_DIAMETER_I"] = record.PipeOuterDiameterSeriesI?.ToString() ?? string.Empty,
                ["PIPE_OUTER_DIAMETER_II"] = record.PipeOuterDiameterSeriesII?.ToString() ?? string.Empty,
                ["FLANGE_OUTER_DIAMETER"] = record.FlangeOuterDiameter?.ToString() ?? string.Empty,
                ["BOLT_CIRCLE_DIAMETER"] = record.BoltCircleDiameter?.ToString() ?? string.Empty,
                ["BOLT_HOLE_DIAMETER"] = record.BoltHoleDiameter?.ToString() ?? string.Empty,
                ["BOLT_COUNT"] = record.BoltCount?.ToString() ?? string.Empty,
                ["BOLT_SPECIFICATION"] = record.BoltSpecification ?? string.Empty,
                ["FLANGE_THICKNESS"] = record.FlangeThickness?.ToString() ?? string.Empty
            };
        }
    }
}
