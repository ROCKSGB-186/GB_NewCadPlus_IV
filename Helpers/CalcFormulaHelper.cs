using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace GB_NewCadPlus_IV.Helpers
{
    public static class CalcFormulaHelper
    {
        /// <summary>
        /// 标准化 Sheet 名称
        /// </summary>
        public static string NormalizeSheetName(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) return string.Empty;
            string clean = sheetName.Trim();
            // 将所有非字母、数字、中文的字符替换为下划线
            clean = Regex.Replace(clean, @"[^a-zA-Z0-9\u4e00-\u9fff]", "_");
            clean = Regex.Replace(clean, @"_+", "_");
            clean = clean.Trim('_');
            return clean.ToLowerInvariant();
        }

        /// <summary>
        /// 【核心修复】尝试修复公式中的常见错误，并提取引用
        /// 1. 修复漏写的 '!' (例如 SheetA1 -> Sheet!A1)
        /// 2. 修复错误的分隔符 (例如 SheetI1 -> Sheet!1)
        /// </summary>
        public static List<string> RepairAndExtractReferences(string formula, string currentSheetName, List<string> allKnownSheetNames)
        {
            var references = new List<string>();
            if (string.IsNullOrWhiteSpace(formula) || !formula.Trim().StartsWith("="))
                return references;

            string cleanFormula = formula.Trim().Substring(1);

            // 【关键修复步骤】预处理公式，尝试插入缺失的 '!'
            // 逻辑：如果公式中出现 "已知Sheet名" 紧接着 "单元格地址" 但没有 '!'，则强制插入 '!'
            // 例如: "DataGrid_循环泵选型计算C16" -> "DataGrid_循环泵选型计算!C16"

            foreach (var sheet in allKnownSheetNames)
            {
                // 转义 Sheet 名中的特殊字符用于正则
                string escapedSheet = Regex.Escape(sheet);

                // 匹配模式: SheetName 后面紧跟 $?[A-Z]+$?[0-9]+，且中间没有 !
                // 使用负向先行断言确保前面没有 !
                string pattern = $@"(?<!\!)({escapedSheet})(\$?[A-Z]+\$?[0-9]+)";

                // 替换为: Sheet!Cell
                cleanFormula = Regex.Replace(cleanFormula, pattern, "$1!$2", RegexOptions.IgnoreCase);
            }

            // 现在使用标准的正则提取
            string extractPattern = @"(?:(?<Sheet>'[^']+'|[a-zA-Z0-9_\u4e00-\u9fff]+)!)?(?<Cell>\$?[A-Z]+\$?[0-9]+)";
            var matches = Regex.Matches(cleanFormula, extractPattern, RegexOptions.IgnoreCase);

            string normalizedCurrentSheet = NormalizeSheetName(currentSheetName);

            foreach (Match match in matches)
            {
                string cellAddrRaw = match.Groups["Cell"].Value;
                string sheetRaw = match.Groups["Sheet"].Value;

                // 验证单元格地址合法性
                if (!Regex.IsMatch(cellAddrRaw, @"^[A-Z]+\d+$", RegexOptions.IgnoreCase))
                    continue;

                string cellAddr = cellAddrRaw.Replace("$", "").ToUpperInvariant();
                string finalSheet;

                if (string.IsNullOrEmpty(sheetRaw))
                {
                    finalSheet = normalizedCurrentSheet;
                }
                else
                {
                    string rawSheetName = sheetRaw.Trim('\'').TrimEnd('!');
                    finalSheet = NormalizeSheetName(rawSheetName);
                }

                string refId = $"{finalSheet}!{cellAddr}";
                if (!references.Contains(refId, StringComparer.OrdinalIgnoreCase))
                {
                    references.Add(refId);
                }
            }

            return references;
        }
    }
}
