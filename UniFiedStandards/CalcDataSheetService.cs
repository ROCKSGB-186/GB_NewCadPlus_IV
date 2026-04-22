using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GB_NewCadPlus_LM.FunctionalMethod.Calculation
{
    public sealed class CalcDataSheetInput
    {
        // 页面直接输入（兼容现有 UI）
        public double 标况烟气量 { get; set; }   // 对应 C6
        public double 工况系数 { get; set; }     // 预留（页面可用）
        public double 液气比 { get; set; }       // 对应 C16
        public int 循环泵台数 { get; set; }      // 可覆盖 C23
        public double 进口流速 { get; set; }     // 可覆盖 C33
        public double 出口流速 { get; set; }     // 可覆盖 C40
        public double 泵扬程 { get; set; }       // 可覆盖 F27
        public double 泵效率 { get; set; }       // 可覆盖循环泵效率(0.8)

        /// <summary>
        /// 公式一致性模式：true 时按 Excel 单元格公式链计算
        /// </summary>
        public bool FormulaConsistencyMode { get; set; } = true;

        /// <summary>
        /// 单元格覆盖（例如 C23, C33, C40, C35, C42 ...）
        /// </summary>
        public Dictionary<string, double> CellOverrides { get; } =
            new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class CalcDataSheetResult
    {
        // 兼容现有页面展示
        public double 工况烟气量 { get; set; }    // C7
        public double 循环总流量 { get; set; }    // C18
        public double 单泵流量 { get; set; }      // C25
        public double 进口管径 { get; set; }      // C34
        public double 出口管径 { get; set; }      // C41
        public double 轴功率 { get; set; }        // C28（循环泵计算功率）

        // 扩展输出
        public double 氧化风量 { get; set; }      // C46
        public double 氧化风压 { get; set; }      // C47
        public double 氧化风机功率 { get; set; }  // C49

        public Dictionary<string, double> Cells { get; } =
            new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

        public List<string> Warnings { get; } = new List<string>();
    }

    public sealed class CalcDataSheetService
    {
        public bool TryCalculate(CalcDataSheetInput input, out CalcDataSheetResult result, out string error)
        {
            result = null;
            error = string.Empty;

            if (input == null) { error = "输入为空"; return false; }

            try
            {
                var r = new CalcDataSheetResult();

                double O(string cell, double fallback)
                {
                    if (input.CellOverrides.TryGetValue(cell, out var v)) return v;
                    return fallback;
                }

                void S(string cell, double value)
                {
                    r.Cells[cell] = value;
                }

                // ===== 常量/输入（来自 CSV 默认值，可被 CellOverrides 覆盖） =====
                var C2 = O("C2", 210d);
                var C3 = O("C3", 35d);
                var C4 = O("C4", 93090d);
                var C5 = O("C5", 30d);
                var C6 = O("C6", input.标况烟气量 > 0 ? input.标况烟气量 : 1735000d);
                var C8 = O("C8", 15.4d);
                var C10 = O("C10", 7000d);
                var C11 = O("C11", 2d);

                var C14 = O("C14", 1.03d);
                var C16 = O("C16", input.液气比 > 0 ? input.液气比 : 2.2d);
                var C17 = O("C17", 1d);
                var C19 = O("C19", 4.35d);

                // 循环泵段
                var C27 = O("C27", 1950d);
                var F27 = O("F27", input.泵扬程 > 0 ? input.泵扬程 : 20d);
                var F28 = O("F28", 1.15d);

                // 管道流速
                var C33 = O("C33", input.进口流速 > 0 ? input.进口流速 : 1.5d);
                var C35 = O("C35", 350d); // DN取值
                var C40 = O("C40", input.出口流速 > 0 ? input.出口流速 : 1.8d);
                var C42 = O("C42", 300d); // DN取值

                // 氧化风机段
                var C48 = O("C48", 16d);
                var F48 = O("F48", 80d);
                var F49 = O("F49", 1.15d);

                // 石灰石浆液泵段
                var C59 = O("C59", 90d);
                var C62 = O("C62", 1200d);
                var C64 = O("C64", 2d);
                var C65 = O("C65", 8d);
                var C66 = O("C66", 1d);
                var C69 = O("C69", 15d);
                var F69 = O("F69", 35d);
                var F70 = O("F70", 1.3d);
                var F71 = O("F71", 0.5d);

                var C73 = O("C73", 1d);
                var C74 = O("C74", 8d);
                var C75 = O("C75", 1d);
                var C78 = O("C78", 10d);
                var F78 = O("F78", 15d);
                var F79 = O("F79", 1.3d);
                var F80 = O("F80", 0.5d);

                // 石膏排出泵段
                var C84 = O("C84", 1130d);
                var C88 = O("C88", 8d);
                var C89 = O("C89", 1d);
                var C92 = O("C92", 15d);
                var F92 = O("F92", 20d);
                var F93 = O("F93", 1.3d);
                var F94 = O("F94", 0.5d);

                var C97 = O("C97", 8d);
                var C98 = O("C98", 1d);
                var C101 = O("C101", 15d);
                var F101 = O("F101", 10d);
                var F102 = O("F102", 1.3d);
                var F103 = O("F103", 0.5d);

                // 缓冲泵段
                var C106 = O("C106", 5.7d);
                var C107 = O("C107", 3d);
                var C109 = O("C109", 1d);
                var C112 = O("C112", 55d);
                var F112 = O("F112", 40d);
                var F113 = O("F113", 1.3d);
                var F114 = O("F114", 0.5d);

                // 滤液水泵段
                var C115 = O("C115", 1000d);
                var C122 = O("C122", 50d);
                var F122 = O("F122", 35d);
                var F123 = O("F123", 1.3d);
                var F124 = O("F124", 0.5d);

                var C131 = O("C131", 50d);
                var F131 = O("F131", 15d);
                var F132 = O("F132", 1.3d);
                var F133 = O("F133", 0.5d);

                // 事故泵/地坑泵段
                var C134 = O("C134", 1150d);
                var C138 = O("C138", 40d);
                var F138 = O("F138", 35d);
                var F139 = O("F139", 1.3d);
                var F140 = O("F140", 0.5d);

                var C143 = O("C143", 40d);
                var F143 = O("F143", 15d);
                var F144 = O("F144", 1.3d);
                var F145 = O("F145", 0.5d);

                var C148 = O("C148", 10d);
                var F148 = O("F148", 15d);
                var F149 = O("F149", 1.3d);
                var F150 = O("F150", 0.4d);

                // ===== 公式链（按 Excel 公式） =====
                var C7 = C6 * 101325d / C4 * (273d + C5) / 273d;
                var C9 = 0.0188d * Math.Sqrt(C7 / C8) * 1000d;
                var C12 = C6 / C11;
                var C13 = C6 * (C2 - C3) / 64d / 1000000d;

                var C18 = C6 * C16 / 1000d;
                var C20 = C18 / 60d * C19;
                var C21 = C10 / 1000d;
                var C22 = C20 / 3.14d / C21 / C21 * 4d;

                var C23Base = C17 * C11;
                var C23 = O("C23", input.循环泵台数 > 0 ? input.循环泵台数 : C23Base);
                var C25 = C18 / C23;

                var C28 = C27 * F27 * 9.81d * 1150d / 3600d / 0.8d / 0.98d / 1000d * F28;

                var C34 = Math.Sqrt(C25 / 3600d / C33 / 3.14d * 4d) * 1000d;
                var C36 = C25 / 3600d / (C35 / 1000d * C35 / 1000d * 3.14d / 4d);

                var C41 = Math.Sqrt(C25 / 3600d / C40 / 3.14d * 4d) * 1000d;
                var C43 = C25 / 3600d / (C42 / 1000d * C42 / 1000d * 3.14d / 4d);

                var C46 = 0.5d * 22.4d * C13 * 0.9d / 0.21d / 0.3d / 60d;
                var C47 = 9.8d * C22;
                var C49 = (C48 * 60d * F48 * 1000d * F49) / (3600d * 102d * 0.9d * 0.98d * 9.8d);

                var C55 = Math.Sqrt(C48 * 60d / 3600d / O("C54", 8d) / 3.14d * 4d) * 1000d;
                var C57 = C48 * 60d / 3600d / (O("C56", 200d) / 1000d * O("C56", 200d) / 1000d * 3.14d / 4d);

                var C61 = C13 * C14 * 100d / C59;
                var C68 = C64 * C61 / 0.2d / 1200d / C66 * 24d / C65;
                var C70 = C69 * F69 * 9.81d * C62 / 3600d / F71 / 0.98d / 1000d * F70;

                var C77 = O("C75", 1d) == 0 ? 0 : (C73 * C61 / 0.2d / 1200d / C75 * 24d / C74);
                var C79 = C78 * F78 * 9.81d * C62 / 3600d / F80 / 0.98d / 1000d * F79;

                var C81 = (C61 * (1d - C59) + C13 * (C14 - 1d) * 100d + C13 * 172d) / 0.9d;
                var C82 = C81 * 0.1d + C13 * 36d;
                var C85 = C81 / 0.2d / C84;

                var C91 = C85 * 24d / C88 / C89;
                var C93 = C92 * F92 * 9.81d * C84 / 3600d / F94 / 0.98d / 1000d * F93;

                var C100 = C85 * 24d / C97 / C98;
                var C102 = C101 * F101 * 9.81d * C84 / 3600d / F103 / 0.98d / 1000d * F102;

                var C111 = C85 * C107 * 24d / C106 / C109;
                var C113 = C112 * F112 * 9.81d * C84 / 3600d / F114 / 0.98d / 1000d * F113;

                var C121 = (C111 * C84 / 1000d * 0.8d * 0.9d + 5d);
                var C123 = C122 * F122 * 9.81d * C115 / 3600d / F124 / 0.98d / 1000d * F123;

                var C130 = (C111 * C84 / 1000d * 0.8d * 0.9d + 5d);
                var C132 = C131 * F131 * 9.81d * C115 / 3600d / F133 / 0.98d / 1000d * F132;

                var C135 = C20 / 8d;
                var C139 = C138 * F138 * 9.81d * C134 / 3600d / F140 / 0.98d / 1000d * F139;
                var C144 = C143 * F143 * 9.81d * C134 / 3600d / F145 / 0.98d / 1000d * F144;

                var C146 = 3d * 3d * 3d / 8d;
                var C149 = C148 * F148 * 9.81d * C134 / 3600d / F150 / 0.98d / 1000d * F149;

                // ===== 回填结果（兼容 UI） =====
                r.工况烟气量 = C7;
                r.循环总流量 = C18;
                r.单泵流量 = C25;
                r.进口管径 = C34;
                r.出口管径 = C41;
                r.轴功率 = C28;
                r.氧化风量 = C46;
                r.氧化风压 = C47;
                r.氧化风机功率 = C49;

                // ===== 单元格映射输出（用于“公式一致性核对”） =====
                S("C4", C4); S("C5", C5); S("C6", C6); S("C7", C7); S("C8", C8); S("C9", C9); S("C10", C10);
                S("C11", C11); S("C12", C12); S("C13", C13); S("C14", C14); S("C16", C16); S("C17", C17);
                S("C18", C18); S("C19", C19); S("C20", C20); S("C21", C21); S("C22", C22); S("C23", C23);
                S("C25", C25); S("C27", C27); S("C28", C28); S("C33", C33); S("C34", C34); S("C35", C35);
                S("C36", C36); S("C40", C40); S("C41", C41); S("C42", C42); S("C43", C43);
                S("C46", C46); S("C47", C47); S("C48", C48); S("C49", C49);
                S("C55", C55); S("C57", C57);
                S("C61", C61); S("C68", C68); S("C70", C70); S("C77", C77); S("C79", C79);
                S("C81", C81); S("C82", C82); S("C85", C85); S("C91", C91); S("C93", C93);
                S("C100", C100); S("C102", C102); S("C111", C111); S("C113", C113);
                S("C121", C121); S("C123", C123); S("C130", C130); S("C132", C132);
                S("C135", C135); S("C139", C139); S("C144", C144); S("C146", C146); S("C149", C149);

                S("F27", F27); S("F28", F28); S("F48", F48); S("F49", F49);
                S("F69", F69); S("F70", F70); S("F71", F71); S("F78", F78); S("F79", F79); S("F80", F80);
                S("F92", F92); S("F93", F93); S("F94", F94); S("F101", F101); S("F102", F102); S("F103", F103);
                S("F112", F112); S("F113", F113); S("F114", F114); S("F122", F122); S("F123", F123); S("F124", F124);
                S("F131", F131); S("F132", F132); S("F133", F133); S("F138", F138); S("F139", F139); S("F140", F140);
                S("F143", F143); S("F144", F144); S("F145", F145); S("F148", F148); S("F149", F149); S("F150", F150);

                // ===== 风险提示 =====
                if (C36 < 1.2d || C36 > 2.0d) r.Warnings.Add($"进口核算流速 C36={C36:0.###} 超建议范围(1.2~2.0)");
                if (C43 < 1.5d || C43 > 2.5d) r.Warnings.Add($"出口核算流速 C43={C43:0.###} 超建议范围(1.5~2.5)");
                if (C57 < 8d || C57 > 15d) r.Warnings.Add($"氧化风出口核算流速 C57={C57:0.###} 超建议范围(8~15)");

                result = r;
                return true;
            }
            catch (Exception ex)
            {
                error = $"公式一致性计算失败: {ex.Message}";
                return false;
            }
        }
    }
}
