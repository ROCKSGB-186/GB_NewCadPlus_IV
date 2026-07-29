using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 标准电机功率帮助类
    /// </summary>
    public static class PowerStandardHelper
    {
        /// <summary>
        /// 标准电机功率（kW），按从小到大排列
        /// </summary>
        private static readonly double[] StandardPowers = {
            0.18, 0.25, 0.37, 0.55, 0.75, 1.1, 1.5, 2.2, 3, 4, 5.5, 7.5, 11, 15,
            18.5, 22, 30, 37, 45, 55, 75, 90, 110, 132, 160, 185, 200, 220, 250,
            280, 315, 355, 400, 450, 500, 560, 630, 710, 800, 900, 1000
        };

        /// <summary>
        /// 向上取整到标准功率
        /// </summary>
        public static double GetStandardPower(double calcPower)
        {
            if (calcPower <= 0) return 0;
            foreach (var std in StandardPowers)
                if (std >= calcPower)
                    return std;
            return calcPower; // 超过最大标准则返回原值
        }
    }
}
