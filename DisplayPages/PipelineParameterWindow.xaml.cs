using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace GB_NewCadPlus_IV.DisplayPages
{
    /// <summary>
    /// 管道参数确认和编辑窗口。
    /// </summary>
    public partial class PipelineParameterWindow : Window
    {
        /// <summary>窗口绑定的参数集合。</summary>
        public ObservableCollection<PipelineParameterRow> Parameters { get; }
            = new ObservableCollection<PipelineParameterRow>();

        /// <summary>当前管道角色。</summary>
        public string PipeRole { get; private set; } = PipelineRoles.Import;

        /// <summary>用户完成参数确认后返回的属性。</summary>
        public Dictionary<string, string> ConfirmedAttributes { get; private set; }
            = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>构造参数窗口。</summary>
        public PipelineParameterWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        /// <summary>
        /// 初始化服务器字段、默认值、端点继承值和冲突提示。
        /// </summary>
        public void Initialize(
            string pipeRole,
            IEnumerable<PipelineFieldDefinitionClient> fields,
            IDictionary<string, string> attributes,
            IEnumerable<string> conflicts)
        {
            PipeRole = string.Equals(pipeRole, PipelineRoles.Export, StringComparison.OrdinalIgnoreCase)
                ? PipelineRoles.Export
                : PipelineRoles.Import;

            PipeRoleTextBlock.Text = PipeRole == PipelineRoles.Export
                ? "出口管道通用参数"
                : "进口管道通用参数";

            Parameters.Clear();
            Dictionary<string, string> sourceAttributes = attributes != null
                ? new Dictionary<string, string>(attributes, StringComparer.OrdinalIgnoreCase)
                : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (PipelineFieldDefinitionClient field in (fields ?? Enumerable.Empty<PipelineFieldDefinitionClient>())
                .OrderBy(item => item.DisplayOrder))
            {
                sourceAttributes.TryGetValue(field.Tag, out string value);
                PipelineParameterRow row = new PipelineParameterRow
                {
                    Tag = field.Tag,
                    Prompt = field.Prompt,
                    Value = value ?? field.DefaultValue ?? string.Empty,
                    Group = field.Group,
                    Editable = field.Editable
                };
                row.PropertyChanged += ParameterRow_PropertyChanged;
                Parameters.Add(row);
            }

            if (Parameters.Count == 0)
            {
                foreach (KeyValuePair<string, string> pair in sourceAttributes)
                {
                    PipelineParameterRow row = new PipelineParameterRow
                    {
                        Tag = pair.Key,
                        Prompt = pair.Key,
                        Value = pair.Value ?? string.Empty,
                        Group = "未分类",
                        Editable = true
                    };
                    row.PropertyChanged += ParameterRow_PropertyChanged;
                    Parameters.Add(row);
                }
            }

            string conflictText = string.Join(Environment.NewLine,
                conflicts ?? Enumerable.Empty<string>());
            ConflictTextBlock.Text = string.IsNullOrWhiteSpace(conflictText)
                ? string.Empty
                : "端点属性存在差异，请确认：" + Environment.NewLine + conflictText;

            RefreshTitlePreview();
            StatusTextBlock.Text = $"已加载 {Parameters.Count} 个管道参数。";
        }

        /// <summary>
        /// 参数值变化时实时刷新标题预览。
        /// </summary>
        private void ParameterRow_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (string.Equals(e.PropertyName, nameof(PipelineParameterRow.Value), StringComparison.Ordinal))
            {
                RefreshTitlePreview();
            }
        }

        /// <summary>
        /// 根据当前页面值刷新 PIPELINETITLE。
        /// </summary>
        private void RefreshTitlePreview()
        {
            Dictionary<string, string> attributes = ReadAttributes();
            PipelineTitleTextBlock.Text = PipelineTitleBuilder.Build(attributes);
        }

        /// <summary>
        /// 完成按钮：校验并返回当前参数。
        /// </summary>
        private void CompleteButton_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, string> attributes = ReadAttributes();
            List<string> requiredFields = Parameters
                .Where(row => row.Editable && string.IsNullOrWhiteSpace(row.Value))
                .Select(row => row.Tag)
                .Where(tag => tag == "DN" || tag == "PN" || tag == "TAG_NO")
                .ToList();

            if (requiredFields.Count > 0)
            {
                MessageBox.Show(
                    "以下必填字段不能为空：" + string.Join("、", requiredFields),
                    "管道参数校验",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            attributes["PIPELINETITLE"] = PipelineTitleBuilder.Build(attributes);
            ConfirmedAttributes = attributes;
            DialogResult = true;
            Close();
        }

        /// <summary>
        /// 取消按钮：不返回正式管道属性。
        /// </summary>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        /// <summary>
        /// 读取页面上的全部属性值。
        /// </summary>
        private Dictionary<string, string> ReadAttributes()
        {
            Dictionary<string, string> attributes =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (PipelineParameterRow row in Parameters)
            {
                if (!string.IsNullOrWhiteSpace(row.Tag))
                {
                    attributes[row.Tag] = row.Value ?? string.Empty;
                }
            }

            return attributes;
        }
    }

    /// <summary>
    /// 管道参数页面的一行数据。
    /// </summary>
    public sealed class PipelineParameterRow : INotifyPropertyChanged
    {
        private string _value = string.Empty;

        /// <summary>Tag。</summary>
        public string Tag { get; set; } = string.Empty;

        /// <summary>提示文本。</summary>
        public string Prompt { get; set; } = string.Empty;

        /// <summary>参数分组。</summary>
        public string Group { get; set; } = string.Empty;

        /// <summary>是否允许编辑。</summary>
        public bool Editable { get; set; } = true;

        /// <summary>参数值。</summary>
        public string Value
        {
            get => _value;
            set
            {
                if (string.Equals(_value, value, StringComparison.Ordinal))
                {
                    return;
                }

                _value = value ?? string.Empty;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Value)));
            }
        }

        /// <summary>属性变化事件。</summary>
        public event PropertyChangedEventHandler PropertyChanged;
    }
}
