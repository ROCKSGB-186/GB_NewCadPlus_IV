using System;
using System.Collections.Generic;

namespace GB_NewCadPlus_IV.Models
{
    /// <summary>
    /// 管道角色编码，进口和出口共用同一套参数字段。
    /// </summary>
    public static class PipelineRoles
    {
        /// <summary>进口管道角色。</summary>
        public const string Import = "IMPORT";

        /// <summary>出口管道角色。</summary>
        public const string Export = "EXPORT";
    }

    /// <summary>
    /// 管道字段定义的数据类型名称。
    /// 使用字符串避免 net48 与 net8 枚举序列化不一致。
    /// </summary>
    public static class PipelineFieldDataTypes
    {
        public const string Text = "text";
        public const string Number = "number";
        public const string Boolean = "boolean";
        public const string Select = "select";
        public const string MultiLine = "multiline";
    }

    /// <summary>
    /// 客户端使用的管道通用字段定义。
    /// </summary>
    public sealed class PipelineFieldDefinitionClient
    {
        /// <summary>AutoCAD 属性 Tag。</summary>
        public string Tag { get; set; } = string.Empty;

        /// <summary>参数页面提示文本。</summary>
        public string Prompt { get; set; } = string.Empty;

        /// <summary>默认值。</summary>
        public string DefaultValue { get; set; } = string.Empty;

        /// <summary>参数控件类型。</summary>
        public string DataType { get; set; } = PipelineFieldDataTypes.Text;

        /// <summary>是否必填。</summary>
        public bool Required { get; set; }

        /// <summary>是否允许编辑。</summary>
        public bool Editable { get; set; } = true;

        /// <summary>字段分组。</summary>
        public string Group { get; set; } = string.Empty;

        /// <summary>字段显示顺序。</summary>
        public int DisplayOrder { get; set; }

        /// <summary>下拉选项。</summary>
        public List<string> Options { get; set; } = new List<string>();
    }

    /// <summary>
    /// 客户端使用的进口/出口角色样式。
    /// </summary>
    public sealed class PipelineRoleStyleClient
    {
        /// <summary>角色编码。</summary>
        public string PipeRole { get; set; } = string.Empty;

        /// <summary>角色显示名称。</summary>
        public string DisplayName { get; set; } = string.Empty;

        /// <summary>标题文字颜色的 AutoCAD ACI 索引。</summary>
        public short TitleColorIndex { get; set; }

        /// <summary>流向符号颜色的 AutoCAD ACI 索引。</summary>
        public short FlowDirectionColorIndex { get; set; }

        /// <summary>流向符号资源编码。</summary>
        public string FlowDirectionSymbol { get; set; } = string.Empty;
    }

    /// <summary>
    /// 管道字段目录响应。
    /// </summary>
    public sealed class PipelineFieldCatalogResponseClient
    {
        /// <summary>响应是否成功。</summary>
        public bool Success { get; set; }

        /// <summary>服务器消息。</summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>通用字段定义。</summary>
        public List<PipelineFieldDefinitionClient> Fields { get; set; } = new List<PipelineFieldDefinitionClient>();

        /// <summary>进口/出口样式。</summary>
        public List<PipelineRoleStyleClient> RoleStyles { get; set; } = new List<PipelineRoleStyleClient>();
    }

    /// <summary>
    /// 管道默认值响应。
    /// </summary>
    public sealed class PipelineDefaultsResponseClient
    {
        /// <summary>响应是否成功。</summary>
        public bool Success { get; set; }

        /// <summary>服务器消息。</summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>按 Tag 返回默认值。</summary>
        public Dictionary<string, string> Attributes { get; set; }
            = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 管道设计规范匹配请求。
    /// </summary>
    public sealed class PipelineDesignStandardMatchRequestClient
    {
        /// <summary>设计标准号。</summary>
        public string DrawingStandardNo { get; set; } = string.Empty;

        /// <summary>公称通径。</summary>
        public string DN { get; set; } = string.Empty;

        /// <summary>公称压力。</summary>
        public string PN { get; set; } = string.Empty;

        /// <summary>壁厚等级。</summary>
        public string Schedule { get; set; } = string.Empty;

        /// <summary>管道材质。</summary>
        public string PipeMaterial { get; set; } = string.Empty;

        /// <summary>介质或介质编码。</summary>
        public string Medium { get; set; } = string.Empty;
    }

    /// <summary>
    /// 管道设计规范匹配响应。
    /// </summary>
    public sealed class PipelineDesignStandardMatchResponseClient
    {
        /// <summary>是否匹配成功。</summary>
        public bool Success { get; set; }

        /// <summary>服务器消息。</summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>匹配数量。</summary>
        public int MatchCount { get; set; }

        /// <summary>是否唯一匹配。</summary>
        public bool IsUniqueMatch { get; set; }

        /// <summary>服务器返回的规范属性。</summary>
        public Dictionary<string, string> Attributes { get; set; }
            = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>实际匹配到的标准号。</summary>
        public string StandardNumber { get; set; } = string.Empty;
    }
}
