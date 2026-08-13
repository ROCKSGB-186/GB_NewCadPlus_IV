using System;
using System.Collections.Generic;

namespace GB_NewCadPlus_IV.Models
{
    /// <summary>
    /// 规范记录的公共字段。专业字段由服务器按规范类型返回。
    /// </summary>
    public sealed class StandardItemClient
    {
        public long Id { get; set; }
        public long SeriesId { get; set; }
        public long VersionId { get; set; }
        public string FamilyCode { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemName { get; set; } = string.Empty;
        public string DN { get; set; } = string.Empty;
        public int? DNValue { get; set; }
        public string PN { get; set; } = string.Empty;
        public string Material { get; set; } = string.Empty;
        public string ConnectionType { get; set; } = string.Empty;
        public int SourceRowNumber { get; set; }
        public Dictionary<string, string> RawValues { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> ExtraAttributes { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public List<string> Warnings { get; set; } = new List<string>();
    }

    /// <summary>
    /// 通用规范匹配结果。旧法兰匹配模型继续用于兼容现有调用方。
    /// </summary>
    public sealed class StandardItemMatchClientResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public int MatchCount { get; set; }
        public bool IsUniqueMatch { get; set; }
        public StandardItemClient? Item { get; set; }
        public Dictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }
}
