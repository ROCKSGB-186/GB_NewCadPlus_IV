using GB_NewCadPlus_IV.Models;
using System;
using System.Collections.Generic;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 将通用规范记录转换为 CAD 属性。专业字段通过 ExtraAttributes 传递，
    /// 具体图块 Tag 规则由专业映射配置或调用方补充。
    /// </summary>
    public static class StandardCadAttributeMapper
    {
        public static Dictionary<string, string> ToAttributes(StandardItemClient item)
        {
            if (item == null) throw new ArgumentNullException(nameof(item));

            var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["FAMILY_CODE"] = item.FamilyCode ?? string.Empty,
                ["ITEM_CODE"] = item.ItemCode ?? string.Empty,
                ["ITEM_NAME"] = item.ItemName ?? string.Empty,
                ["DN"] = item.DN ?? string.Empty,
                ["PN"] = item.PN ?? string.Empty,
                ["MATERIAL"] = item.Material ?? string.Empty,
                ["CONNECTION_TYPE"] = item.ConnectionType ?? string.Empty
            };

            if (item.ExtraAttributes != null)
            {
                foreach (KeyValuePair<string, string> attribute in item.ExtraAttributes)
                {
                    if (string.IsNullOrWhiteSpace(attribute.Key)) continue;
                    attributes[attribute.Key.Trim()] = attribute.Value ?? string.Empty;
                }
            }

            return attributes;
        }
    }
}
