using GB_NewCadPlus_LM.FunctionalMethod;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_LM.Helpers
{
    public static class LayerInfoHelper
    {
        public static double SafeGetSystemVariableDouble(string varName, double defaultValue = 1.0)
        {
            // 安全读取系统变量并转换为 double，任何异常都返回默认值
            try
            {
                var val = Application.GetSystemVariable(varName);
                if (val == null) return defaultValue;
                return Convert.ToDouble(val);
            }
            catch
            {
                return defaultValue;
            }
        }   

    }
}
