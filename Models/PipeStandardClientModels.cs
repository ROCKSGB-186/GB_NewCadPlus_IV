using System;

namespace GB_NewCadPlus_IV.Models
{
    /// <summary>
    /// 管子规范专业扩展。
    /// </summary>
    public sealed class PipeStandardExtensionClient
    {
        public decimal? OuterDiameter { get; set; }
        public decimal? WallThickness { get; set; }
        public string Schedule { get; set; } = string.Empty;
        public decimal? UnitWeight { get; set; }
        public decimal? StandardLength { get; set; }
    }
}
