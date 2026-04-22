using GB_NewCadPlus_IV.FunctionalMethod;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.Helpers
{
    public class TianZhengHelper
    {
        /// <summary>
        /// 天丄接口数据
        /// </summary>
        public static string hvacR4 = "0";
        /// <summary>
        /// 天正接口数据
        /// </summary>
        public static string hvacR3 = "0";
        /// <summary>
        /// 临时存储接口数据
        /// </summary>
        public static string strHvacStart = "0";
        /// <summary>
        /// 获取天正数据
        /// </summary>
        [CommandMethod(nameof(tzData))]
        public static void tzData()
        {
            var sEper = Env.Editor.GetEntity("\n选择要标注的实体");
            if (sEper.Status != PromptStatus.OK)
                return;
            using var Tr = new DBTrans();
            try
            {
                //判断是不是曲线实体
                if (Tr.GetObject(sEper.ObjectId) is not Curve sEperObi)
                    return;
                //获取曲线实体的AcadObject对象
                var aCadSeperOb = sEperObi.AcadObject;
                if (aCadSeperOb != null)
                {
                    //获取到宽
                    hvacR4 = AddMenus.GetProperty(aCadSeperOb, "Hvac_R4").ToString();
                    //获取到高（厚）
                    hvacR3 = AddMenus.GetProperty(aCadSeperOb, "Hvac_R3").ToString();
                    //获取距地值，返回的是object[]数组；
                    object HvacStart = AddMenus.GetProperty(aCadSeperOb, "Hvac_Start");
                    //var havcR4 = Convert.ToString(aCadSeperOb.GetType().InvokeMember("Hvac_R4", BindingFlags.GetProperty, null, aCadSeperOb, null));
                    double[] doubles = new double[3] { 0, 0, 0 };
                    doubles = (double[])HvacStart;
                    strHvacStart = Convert.ToString(doubles[2]);
                    LogManager.Instance.LogInfo("\nhvacR4:" + hvacR4);
                    LogManager.Instance.LogInfo("\nhvacR3:" + hvacR3);
                    LogManager.Instance.LogInfo("\nhvacStart:" + strHvacStart);
                }
            }
            catch
            {
                //LogManager.Instance.LogInfo("您选定的图无不为天正图无，不能读出宽厚参数！");//在下面的历史记录框里显示一样文字
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("您选定的图元没有天正元素，不能读出宽厚等参数！");//弹出一个带有声音的消息框；
                return;
            }
        }

    }
}
