using System;
using System.Collections.Generic;

namespace GB_NewCadPlus_IV.Models
{
    /// <summary>
    /// 客户端发送给服务器的法兰规范查询条件。
    /// </summary>
    public sealed class FlangeStandardMatchRequest
    {
        /// <summary>规范大类编码。</summary>
        public string FamilyCode { get; set; } = "FLANGE";

        /// <summary>规范系列编码。</summary>
        public string SeriesCode { get; set; } = "PLATE_WELD";

        /// <summary>标准号。</summary>
        public string StandardNumber { get; set; } = string.Empty;

        /// <summary>标准表号。</summary>
        public string TableNumber { get; set; } = string.Empty;

        /// <summary>PN 筛选条件。</summary>
        public string PN { get; set; } = "PN10";

        /// <summary>公称尺寸，保持 DN50 形式。</summary>
        public string DN { get; set; } = string.Empty;

        /// <summary>钢管外径系列。</summary>
        public string Series { get; set; } = "Ⅰ系列";

        /// <summary>法兰类型。</summary>
        public string FlangeType { get; set; } = "PL";

        /// <summary>密封面型式。</summary>
        public string FaceType { get; set; } = "RF";
    }

    /// <summary>
    /// 服务器返回的法兰规范匹配结果。
    /// </summary>
    public sealed class FlangeStandardMatchResponse
    {
        /// <summary>是否匹配成功。</summary>
        public bool Success { get; set; }

        /// <summary>服务器提示。</summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>匹配数量。</summary>
        public int MatchCount { get; set; }

        /// <summary>是否唯一匹配。</summary>
        public bool IsUniqueMatch { get; set; }

        /// <summary>可直接写入 AutoCAD 属性块的属性。</summary>
        public Dictionary<string, string> Attributes { get; set; }
            = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>完整规范记录。</summary>
        public FlangeStandardRecordClient? Record { get; set; }
    }

    /// <summary>
    /// 客户端使用的法兰规范完整记录。
    /// </summary>
    public sealed class FlangeStandardRecordClient
    {
        public long Id { get; set; }
        public long SeriesId { get; set; }
        public int SourceRowNumber { get; set; }
        public string DN { get; set; } = string.Empty;
        public int DNValue { get; set; }
        public string PN { get; set; } = string.Empty;
        public decimal? PipeOuterDiameterSeriesI { get; set; }
        public decimal? PipeOuterDiameterSeriesII { get; set; }
        public decimal? FlangeOuterDiameter { get; set; }
        public decimal? BoltCircleDiameter { get; set; }
        public decimal? BoltHoleDiameter { get; set; }
        public int? BoltCount { get; set; }
        public string BoltSpecification { get; set; } = string.Empty;
        public string? BoltRawSuffix { get; set; }
        public decimal? FlangeThickness { get; set; }
        public decimal? RaisedFaceHeight { get; set; }
        public decimal? FlangeInnerDiameterSeriesI { get; set; }
        public decimal? FlangeInnerDiameterSeriesII { get; set; }
        public Dictionary<string, string> RawValues { get; set; }
            = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public List<string> Warnings { get; set; } = new List<string>();
    }
}
