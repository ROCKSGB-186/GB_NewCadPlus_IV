using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 管道标题生成器。
    /// 标题只由统一属性字典生成，避免多个页面分别拼接导致规则不一致。
    /// </summary>
    public static class PipelineTitleBuilder
    {
        /// <summary>
        /// 根据管道属性生成 PIPELINETITLE。
        /// </summary>
        public static string Build(IDictionary<string, string> attributes)
        {
            string dn = GetValue(attributes, "DN");
            string mediumCode = GetValue(attributes, "MEDIUM_CODE");
            if (string.IsNullOrWhiteSpace(mediumCode))
            {
                mediumCode = ToMediumCode(GetValue(attributes, "MEDIUM"));
            }

            string tagNo = NormalizeTagNo(GetValue(attributes, "TAG_NO"));
            string pnValue = ToPressureTitleValue(GetValue(attributes, "PN"));
            string isolationCode = GetValue(attributes, "HOT\\SOUND_ISOLACODE");

            return $"{RemovePrefix(dn, "DN")}-{mediumCode}-{tagNo}-{pnValue}{isolationCode}";
        }

        /// <summary>
        /// 去掉 DN 前缀，保留标题中的数字部分。
        /// </summary>
        private static string RemovePrefix(string value, string prefix)
        {
            string normalized = (value ?? string.Empty).Trim();
            return normalized.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                ? normalized.Substring(prefix.Length)
                : normalized;
        }

        /// <summary>
        /// 将 TAG_NO 规范为四位数字；非数字值原样保留，避免丢失业务编号。
        /// </summary>
        private static string NormalizeTagNo(string value)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out int number))
            {
                return number.ToString("D4", CultureInfo.InvariantCulture);
            }

            return normalized;
        }

        /// <summary>
        /// 将 PN10 转换为标题中的 1.0。
        /// </summary>
        private static string ToPressureTitleValue(string value)
        {
            string normalized = RemovePrefix(value, "PN");
            if (decimal.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal number))
            {
                return (number / 10m).ToString("0.0", CultureInfo.InvariantCulture);
            }

            return normalized;
        }

        /// <summary>
        /// 介质已经是工程编码时直接使用；中文介质暂按已确认的常用编码转换。
        /// </summary>
        private static string ToMediumCode(string value)
        {
            string normalized = (value ?? string.Empty).Trim();
            Dictionary<string, string> codes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["石灰石浆液"] = "AR",
                ["石膏浆液"] = "GY",
                ["原烟气"] = "FG",
                ["净烟气"] = "NG",
                ["工艺水"] = "PW",
                ["压缩空气"] = "CA",
                ["蒸汽"] = "ST",
                ["酸碱介质"] = "AB"
            };

            return codes.TryGetValue(normalized, out string code) ? code : normalized;
        }

        /// <summary>
        /// 读取不区分大小写的属性。
        /// </summary>
        private static string GetValue(IDictionary<string, string> attributes, string key)
        {
            return attributes != null && attributes.TryGetValue(key, out string value)
                ? value ?? string.Empty
                : string.Empty;
        }
    }
}
