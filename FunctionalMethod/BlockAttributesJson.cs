using System;
using System.Collections.Generic;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 图元属性 JSON 的最小可编译模型。
    /// </summary>
    public class BlockAttributesJson
    {
        /// <summary>
        /// 原始属性字典。
        /// </summary>
        public Dictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 原始属性 JSON 文本。
        /// </summary>
        public string AttributesJson { get; set; }

        /// <summary>
        /// 配置名称。
        /// </summary>
        public string ConfigName { get; set; }

        /// <summary>
        /// 创建时间。
        /// </summary>
        public DateTime? CreatedAt { get; set; }

        /// <summary>
        /// 更新时间。
        /// </summary>
        public DateTime? UpdatedAt { get; set; }

        /// <summary>
        /// 备注信息。
        /// </summary>
        public string Remarks { get; set; }
    }
}
