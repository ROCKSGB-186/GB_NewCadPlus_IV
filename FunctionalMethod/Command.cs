using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Windows;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.UniFiedStandards;
using IFoxCAD.Cad;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Org.BouncyCastle.Utilities;
using System.Windows;
using System.Windows.Forms.Integration;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Line = Autodesk.AutoCAD.DatabaseServices.Line;
using Path = System.IO.Path;

/// <summary>
/// CAD远行时自动运行 在菜单内加入工具
/// </summary>
public class AutodeskRun : IExtensionApplication
{
    /// <summary>
    /// 启动加载的命令扩展
    /// </summary>
    public void Initialize()
    {
        AddMenus.AddMenu();
    }
    /// <summary>
    /// 卸载
    /// </summary>
    public void Terminate()
    {

    }
}
/// <summary>
/// 命令参数类
/// </summary>
public class PsetArgs
{
    /// <summary>
    /// 构造函数
    /// </summary>
    public PsetArgs() { }

}
namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// Point3d 扩展方法：提供根据极坐标计算点位置的功能，简化在 CAD 中基于角度和距离定位点的操作。通过 PolarPoint 方法，用户可以直接从一个中心点出发，指定一个角度和距离，快速得到目标点的位置。这对于需要在特定方向上放置对象或进行测量的场景非常有用，提升了代码的可读性和开发效率。
    /// </summary>
    public static class Point3dExtensions
    {
        /// <summary>
        /// 扩展方法：根据极坐标计算点的位置 ffff
        /// </summary>
        /// <param name="center">中心点</param>
        /// <param name="angle">角度</param>
        /// <param name="distance"></param>
        /// <returns></returns>
        public static Point3d PolarPoint(this Point3d center, double angle, double distance)
        {
            double radians = angle * Math.PI / 180; // 将角度转换为弧度
            // 计算极坐标下的点
            double x = center.X + distance * Math.Cos(radians);
            // 计算极坐标下的点
            double y = center.Y + distance * Math.Sin(radians);
            return new Point3d(x, y, center.Z);
        }
    }

    /// <summary>
    /// 委托：在要发送内容的类里，建立一个委托，再实例化这个委托。同时给实例化的委托sendSum传值（sendSum?.invoke(传递值)）；在接收类里建立一个赋值方法，这个方法是这个值给到接收文本框显示的值，再在接收页面初始化方法里把这个委托值给到赋值方法即可；
    /// </summary>
    /// <param name="text">传递的text</param>
    public delegate void sendText(string text);

    /// <summary>
    /// 主命令类 
    /// </summary>
    public class Command
    {
        /// <summary>
        /// 静态变量，用于保存图库管理窗体
        /// </summary>
        private static PaletteSet? Wpf_Cad_PaletteSet;

        /// <summary>
        /// 方向变更事件：外部（WPF按钮/UnifiedCommandManager）通过 NotifyDirectionChanged 广播当前角度（弧度）
        /// EnhancedBlockPlacementJig 在创建时订阅此事件以实现拖拽期间的即时预览旋转。
        /// </summary>
        public static event Action<double>? DirectionChanged;

        /// <summary>
        /// 外部调用以广播方向变化（弧度）
        /// 同步 VariableDictionary.entityRotateAngle 并通知订阅者。
        /// </summary>
        /// <param name="angle">弧度</param>
        public static void NotifyDirectionChanged(double angle)
        {
            try
            {
                //VariableDictionary.entityRotateAngle = angle;
                //触发方向变更事件 广播事件通知所有订阅者
                DirectionChanged?.Invoke(angle);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\nNotifyDirectionChanged 错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 存储用户点击的点
        /// </summary>
        private static List<Point3d> pointS = new List<Point3d>();

        /// <summary>
        /// 当前图纸空间的ObjectId
        /// </summary>
        private static List<ObjectId>? currentSpaceObjectId = new List<ObjectId>();

        #region 接口、实体、天正数据、标注等


        /// <summary>
        /// 发送纯文本
        /// </summary>
        public static sendText? sendSum;
        /// <summary>
        /// 图层状态列表
        /// </summary>
        public static List<LayerState> layerOriStates = new List<LayerState>();
        /// <summary>
        /// 图层是否已被删除
        /// </summary>
        public static bool _isLayersDeleted = false;
        /// <summary>
        /// 图层信息字典
        /// </summary>
        public static Dictionary<string, LayerState> _originalLayerInfos;
        /// <summary>
        ///  图层开关状态字典
        /// </summary>
        public static Dictionary<string, bool> layerOnOffDic = new Dictionary<string, bool>();
        /// <summary>
        /// 读取图层开关状态
        /// </summary>
        public static bool readLayerONOFFState = false;

        #region 辅助方法,检查图层,标注文字内容构建,标注样式设置,自动孵化等


        /// <summary>
        /// 尝试获取双精度浮点数
        /// </summary>
        /// <param name="text">要解析的文本</param>
        /// <param name="defaultValue">默认值</param>
        /// <param name="value">解析结果</param>
        /// <returns>解析是否成功</returns>
        private static bool TryGetDouble(string? text, double defaultValue, out double value)
        {
            if (double.TryParse(text, out value))
                return true;

            value = defaultValue;
            return false;
        }
        /// <summary>
        /// 确保提示确认
        /// </summary>
        /// <param name="result">提示结果</param>
        /// <param name="tr">数据库事务对象</param>
        /// <returns>是否确认</returns>
        private static bool EnsurePromptOk(PromptResult result, DBTrans tr)
        {
            if (result.Status == PromptStatus.OK) return true;
            tr.Abort();
            return false;
        }
        /// <summary>
        /// 自动孵化核心
        /// </summary>
        /// <param name="tr">数据库事务对象</param>
        /// <param name="layerName">图层名称</param>
        /// <param name="hatchColorIndex">孵化颜色索引</param>
        /// <param name="hatchPatternScale">孵化图案缩放</param>
        /// <param name="patternName">图案名称</param>
        /// <param name="boundaryId">边界对象ID</param>
        /// <returns></returns>
        private static ObjectId AutoHatchCore(
            DBTrans tr,
            string layerName,
            int hatchColorIndex,
            int hatchPatternScale,
            string patternName,
            ObjectId boundaryId)
        {
            var hatch = new Hatch();
            hatch.SetHatchPattern(HatchPatternType.PreDefined, patternName);
            hatch.PatternScale = hatchPatternScale;
            hatch.Layer = layerName;
            hatch.ColorIndex = hatchColorIndex;
            hatch.PatternAngle = 0;
            hatch.Normal = Vector3d.ZAxis;

            var loops = new ObjectIdCollection { boundaryId };
            hatch.AppendLoop(HatchLoopTypes.Outermost, loops);
            hatch.EvaluateHatch(true);

            return tr.CurrentSpace.AddEntity(hatch);
        }



        /// <summary>
        /// 统一创建并交互放置 MLeader
        /// </summary>
        private static bool TryCreateLeaderByDrag(
            DBTrans tr,
            string layerName,
            short textColor,
            Point3d firstPoint,
            string contents,
            double textHeight,
            double arrowSize,
            TextAttachmentType textAttachmentType = TextAttachmentType.AttachmentBottomOfTopLine)
        {
            MText mt = new MText();
            MLeader mld = new MLeader();

            if (layerName == "TJ(设备位号)")
            {
                // 应用文本样式和图层信息到 MText 和 MLeader
                TextFontsStyleHelper.TextStyleAndLayerInfo(tr, layerName, textColor, "HZFS", ref mt, ref mld);
                mld.TextStyleId = tr.TextStyleTable["HZFS"];// 设置标注文本样式
            }
            else
            {
                // 应用文本样式和图层信息到 MText 和 MLeader
                TextFontsStyleHelper.TextStyleAndLayerInfo(tr, layerName, textColor, "tJText", ref mt, ref mld);
                mld.TextStyleId = tr.TextStyleTable["tJText"];// 设置标注文本样式
            }
            ;
            mt.Contents = contents;// 设置标注文本内容
            mt.ColorIndex = textColor; // 设置标注文本颜色
            mt.Attachment = AttachmentPoint.MiddleCenter;// 设置标注文本的附着点为中间
            mt.Height = textHeight;// 设置标注文本高度            
            mld.Layer = layerName;// 设置标注图层
            mld.ColorIndex = textColor;// 设置标注颜色
            mld.LeaderLineColor = Color.FromColorIndex(ColorMethod.ByAci, textColor);// 设置引线颜色
            mld.TextAttachmentType = textAttachmentType;// 设置文本附着类型
            //mld.TextHeight = textHeight;// 设置文本高度
            //mld.ArrowSize = arrowSize;// 设置箭头大小
            mld.MText = mt;// 将 MText 赋值给 MLeader 的 MText 属性            

            int ldNum = mld.AddLeader();// 添加一个引线并获取引线编号
            int lnNum = mld.AddLeaderLine(ldNum);// 在指定引线上添加一条引线并获取引线编号
            // 注意：AddFirstVertex 必须在设置 MLeader 属性之后调用，以确保正确应用样式和属性。
            mld.AddFirstVertex(lnNum, firstPoint);// 在指定引线上添加第一个顶点，位置为用户指定的第一点
            using var jig = new JigEx((mpw, _) =>
            {
                mld.TextLocation = mpw.Z20();
                if (layerName == "TJ(设备位号)")
                {
                    mld.TextStyleId = tr.TextStyleTable["HZFS"];// 设置标注文本样式
                }
                else
                {
                    mld.TextStyleId = tr.TextStyleTable["tJText"];// 设置标注文本样式
                }
                ;
                mld.TextHeight = textHeight;// 设置文本高度
                mld.ArrowSize = arrowSize;// 设置箭头大小
            });

            // 保持你当前代码行为
            mld.AddLastVertex(lnNum, jig.MousePointWcsLast);
            jig.DatabaseEntityDraw(wb => wb.Geometry.Draw(mld));
            jig.SetOptions(firstPoint, msg: "\n点选标注第二点");
            var dragRes = Env.Editor.Drag(jig);
            if (dragRes.Status != PromptStatus.OK) return false;
            //mld.MText = mt;// 将 MText 赋值给 MLeader 的 MText 属性
            if (layerName == "TJ(设备位号)")
            {

                mld.TextStyleId = tr.TextStyleTable["HZFS"];// 设置标注文本样式
            }
            else
            {
                mld.TextStyleId = tr.TextStyleTable["tJText"];// 设置标注文本样式
            }
            ;
            //mld.TextStyleId = tr.TextStyleTable["HZFS"];
            mld.TextHeight = textHeight;// 设置文本高度
            mld.ArrowSize = arrowSize;// 设置箭头大小
            mld.Layer = layerName;
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return false;
            using (doc.LockDocument()) // 锁定文档以安全修改
                tr.CurrentSpace.AddEntity(mld);
            return true;
        }

        #endregion

        #endregion

        /// <summary>
        /// 显示主窗体
        /// </summary>
        [CommandMethod(nameof(gfff))]
        public static void gfff()
        {
            try
            {
                DateTime setDate = new DateTime(2026, 12, 30);
                if (DateTime.Now < setDate)
                {
                    FormMain.GB_CadToolsForm.ShowToolsPanel();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("试用时间过期！");
                    LogManager.Instance.LogInfo("\n试用时间已过期。");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"显示主窗体时出错: {ex.Message}");
                LogManager.Instance.LogError($"错误堆栈: {ex.StackTrace}");
            }
        }
        [CommandMethod(nameof(ffff))]
        public static void ffff()
        {
            try
            {
                DateTime setDate = new DateTime(2026, 12, 30);
                if (DateTime.Now < setDate)
                {
                    LogManager.Instance.LogInfo("\n开始显示主窗体...");

                    var login = new LoginWindow();
                    if (login.ShowDialog() != true)
                    {
                        LogManager.Instance.LogInfo("\n登录窗口取消或关闭。");
                        return;
                    }

                    if (Wpf_Cad_PaletteSet is null)
                    {
                        try
                        {
                            //创建窗体容器
                            Wpf_Cad_PaletteSet = new PaletteSet("GB_CADTools");  //初始化窗体容器；
                            Wpf_Cad_PaletteSet.MinimumSize = new System.Drawing.Size(350, 800);//初始化窗体容器最小的尺寸

                            var wpfWindows = new WpfMainWindow();//初始化这个图库管理窗体；
                            var host = new ElementHost()//初始化子面板
                            {
                                AutoSize = true,//设置子面板自动大小
                                Dock = DockStyle.Fill,//子面板整体覆盖
                                Child = wpfWindows//设置子面板的子项为wpfWindows
                            };
                            Wpf_Cad_PaletteSet.Add("GB_CADTools", host);//添加子面板
                            Wpf_Cad_PaletteSet.Visible = true;//显示窗体容器
                            Wpf_Cad_PaletteSet.Dock = DockSides.Left;//窗体容器的停靠位置
                            //FormMain.GB_CadToolsForm.ShowToolsPanel();
                            LogManager.Instance.LogInfo("\n主窗体已成功创建并显示。");
                            return;
                        }
                        catch (Exception ex)
                        {
                            LogManager.Instance.LogError($"创建主窗体时出错: {ex.Message}");
                            LogManager.Instance.LogError($"错误堆栈: {ex.StackTrace}");
                        }
                    }
                    else
                    {
                        Wpf_Cad_PaletteSet.Visible = !Wpf_Cad_PaletteSet.Visible;
                        LogManager.Instance.LogInfo($"\n主窗体可见性已切换为: {Wpf_Cad_PaletteSet.Visible}");
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("试用时间过期！");
                    LogManager.Instance.LogInfo("\n试用时间已过期。");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"显示主窗体时出错: {ex.Message}");
                LogManager.Instance.LogError($"错误堆栈: {ex.StackTrace}");
            }
        }

        #region 填充方面

        /// <summary>
        /// 填充入口：解析图层名 + 解析颜色 + 确保图层存在 + 创建填充
        /// </summary>
        /// <param name="tr">数据库事务对象</param>
        /// <param name="layerName">图层名称</param>
        /// <param name="hatchColorIndex">填充颜色索引</param>
        /// <param name="hatchPatternScale">填充图案比例</param>
        /// <param name="patternName">填充图案名称</param>
        /// <param name="entityObjId">实体对象ID</param>
        public static void autoHatch(DBTrans tr, string layerName, int hatchColorIndex, int hatchPatternScale, string patternName, ObjectId entityObjId)
        {
            try
            {
                _ = AutoHatchCore(tr, layerName, hatchColorIndex, hatchPatternScale, patternName, entityObjId);
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n创建填充时出错: {ex.Message}");
            }
        }
        /// <summary>
        /// 填充入口：解析图层名 + 解析颜色 + 确保图层存在 + 创建填充，并返回填充对象ID
        /// </summary>
        /// <param name="tr">数据库事务对象</param>
        /// <param name="layerName">图层名称</param>
        /// <param name="hatchColorIndex">填充颜色索引</param>
        /// <param name="hatchPatternScale">填充图案比例</param>
        /// <param name="patternName">填充图案名称</param>
        /// <param name="entityObjId">实体对象ID</param>
        /// <param name="hatchId"></param>
        public static void autoHatch(DBTrans tr, string layerName, int hatchColorIndex, int hatchPatternScale, string patternName, ObjectId entityObjId, ref ObjectId hatchId)
        {
            try
            {
                hatchId = AutoHatchCore(tr, layerName, hatchColorIndex, hatchPatternScale, patternName, entityObjId);
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n创建填充时出错: {ex.Message}");
            }
        }

        #endregion


        #region 标注方面

        /// <summary>
        /// 统一拼接标注文字（按 buttonText 场景）
        /// </summary>
        /// <param name="dimString1"> 标注内容1</param>
        /// <param name="dimString2"> 标注内容2</param>
        /// <param name="jztjUseMeter"> 是否使用米作为单位</param>
        /// <param name="includeDeviceCode"> 是否包含设备编号</param>
        /// <returns>拼接后的标注文字</returns>
        private static string BuildContextualDimText(
            string dimString1,
            string? dimString2,
            bool jztjUseMeter = false,
            bool includeDeviceCode = false)
        {
            string buttonText = VariableDictionary.buttonText ?? string.Empty;

            if (buttonText.Contains("受力点"))
            {
                string pointCount = string.IsNullOrWhiteSpace(dimString2) ? "1" : dimString2;
                return $"总重:{dimString1}kg\n{pointCount}点受力";
            }

            if (buttonText.Contains("矩形开洞"))
            {
                return $"矩形洞口\n{dimString1}x{dimString2}";
            }

            if (buttonText.Contains("圆形开洞"))
            {
                return $"圆形洞口\n直径:{dimString1}";
            }

            if (buttonText.Contains("JZTJ"))
            {
                string unit = jztjUseMeter ? "m" : "mm";
                return $"排水沟\n宽:{dimString2}{unit}，深:{dimString1}{unit}";
            }

            if (includeDeviceCode && buttonText.Contains("设备位号"))
            {
                return $"{dimString1}-{dimString2}";
            }

            return string.IsNullOrWhiteSpace(dimString2) ? dimString1 : $"{dimString1}-{dimString2}";
        }

        /// <summary>
        /// 点线标注 TJ(结构专业JG):框着地,面着地,点受力,水平荷载
        /// </summary>
        /// <param name="dimString">标注文字内容</param>
        public static void DDimLinear(DBTrans tr, string dimString)
        {
            try
            {
                // 文字颜色
                short textColor = LayerControlHelper.ResolveLayerColor();
                // 特例：如果是结构专业的标注，强制使用红色（颜色索引3），以符合常见的行业标准。
                if (string.Equals(VariableDictionary.btnFileName, "TJ(结构专业JG)", StringComparison.OrdinalIgnoreCase))
                {
                    textColor = 3;
                }
                // 获取或创建目标图层，优先使用 btnBlockLayer 作为图层名称，如果没有则使用 layerName，确保图层存在并设置颜色。
                string targetLayer = LayerControlHelper.GetOrCreateTargetLayer(
                    tr,
                    VariableDictionary.btnBlockLayer ?? VariableDictionary.layerName,
                    textColor);
                // 根据 buttonText 场景构建适合的标注文本内容，特别处理结构专业的不同受力类型。
                string buttonText = VariableDictionary.buttonText ?? string.Empty;
                // 默认标注内容为 dimString，但如果 buttonText 包含特定关键词，则构建更详细的内容，包括总重和受力类型。
                string content = dimString;
                if (buttonText.Contains("框着地"))
                    content = $"总重:{dimString}kg\n框着地";
                else if (buttonText.Contains("面着地"))
                    content = $"总重:{dimString}kg\n面着地";
                else if (buttonText.Contains("点受力"))
                    content = $"总重:{dimString}kg\n点受力";
                else if (buttonText.Contains("水平荷载"))
                    content = $"总重:{dimString}kg\n水平荷载";
                // 根据用户输入的标注内容和当前场景，构建适合的标注文本，并调用统一的拖拽放置方法创建标注。
                var userPoint1 = Env.Editor.GetPoint("\n请指定标注第一点");
                if (userPoint1.Status != PromptStatus.OK) return;
                var ucsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();
                // 计算标注的文字高度和箭头大小，基于用户界面缩放比例，确保在不同缩放级别下标注具有合适的大小。
                double uiScale = VariableDictionary.textBoxScale;
                if (double.IsNaN(uiScale) || uiScale <= 0)
                {
                    try { uiScale = AutoCadHelper.GetScale(true); } catch { uiScale = 1.0; }
                }
                if (uiScale <= 0) uiScale = 1.0;

                double textHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, uiScale);
                double arrowSize = TextFontsStyleHelper.ComputeScaledHeight(2, uiScale);
                // 根据用户输入的标注内容和当前场景，构建适合的标注文本，并调用统一的拖拽放置方法创建标注。
                bool ok = TryCreateLeaderByDrag(
                    tr,
                    targetLayer,
                    textColor,
                    ucsUserPoint1,
                    content,
                    textHeight: textHeight,
                    arrowSize: arrowSize,
                    textAttachmentType: TextAttachmentType.AttachmentBottomOfTopLine);

                if (!ok)
                {
                    tr.Abort();
                    return;
                }

                //tr.Commit();
                Env.Editor.Redraw();
                VariableDictionary.dimString = null;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n做标注失败！错误信息: {ex.Message}");
                LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}");
            }
        }


        /// <summary>
        /// 将普通字符串转为 MText 安全文本（转义格式控制字符）
        /// </summary>
        /// <param name="text">原始文本</param>
        /// <returns>可安全写入 MText.Contents 的文本</returns>
        private static string EscapeMTextLiteral(string? text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty; // 空文本直接返回空字符串
            string s = text; // 复制原文本
            s = s.Replace("\\", "\\\\"); // 转义反斜杠，避免被当成格式码
            s = s.Replace("{", "\\{"); // 转义左花括号
            s = s.Replace("}", "\\}"); // 转义右花括号
            return s; // 返回转义后的文本
        }

        /// <summary>
        /// 构建 MLeader 文本：第一行正常字高，第二行半字高
        /// </summary>
        /// <param name="dimString1">主文本（第一行）</param>
        /// <param name="dimString2">副文本（第二行，可空）</param>
        /// <returns>MText 内容字符串</returns>
        private static string BuildLeaderContentMainAndSecondary(string dimString1, string? dimString2)
        {
            // 转义主文本，避免被 MText 格式码误解析
            string main = EscapeMTextLiteral(dimString1);

            // 副文本为空时，只返回主文本
            if (string.IsNullOrWhiteSpace(dimString2)) return main;

            // 转义副文本
            string sub = EscapeMTextLiteral(dimString2);

            // 同一行输出：主文本-副文本
            // 仅副文本使用 0.5 倍字高，并在末尾恢复为 1 倍字高
            return $"{main}\\H0.5x;-{sub}\\H1x;"; // “-”与dimString2同为0.5倍字高
        }

        /// <summary>
        /// 线性标注
        /// </summary>
        /// <param name="dimString1">标注文字1</param>
        /// <param name="dimString2">标注文字2</param>
        public static void DDimLinear(DBTrans tr, string dimString1, string dimString2 = null)
        {
            try
            {
                short textColor = LayerControlHelper.ResolveLayerColor(); // 解析标注颜色

                // 确保目标图层存在
                string targetLayer = LayerControlHelper.GetOrCreateTargetLayer(
                    tr,
                    LayerControlHelper.ResolveTargetLayerName(VariableDictionary.layerName ?? VariableDictionary.btnBlockLayer),
                    textColor);

                string content = BuildLeaderContentMainAndSecondary(dimString1, dimString2); // 生成“第二行半字高”的 MText 内容

                var userPoint1 = Env.Editor.GetPoint("\n请指定标注第一点"); // 获取引线第一点
                if (userPoint1.Status != PromptStatus.OK) return; // 用户取消则退出
                var ucsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20(); // 转换到 UCS

                double uiScale = AutoCadHelper.GetScale(true); // 读取当前界面比例

                double textHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, uiScale); // 文字高度
                double arrowSize = TextFontsStyleHelper.ComputeScaledHeight(2.0, uiScale); // 箭头尺寸

                bool ok = TryCreateLeaderByDrag( // 调用统一创建逻辑
                    tr,
                    targetLayer,
                    textColor,
                    ucsUserPoint1,
                    content,
                    textHeight: textHeight,
                    arrowSize: arrowSize,
                    textAttachmentType: TextAttachmentType.AttachmentBottomOfTopLine);

                if (!ok) // 如果创建失败或用户取消
                {
                    tr.Abort(); // 回滚事务
                    return; // 结束
                }

                Env.Editor.Redraw(); // 刷新显示
                VariableDictionary.dimString = null; // 清空缓存标注内容
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n做标注失败！错误信息: {ex.Message}"); // 记录错误日志
                LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}"); // 记录堆栈
            }
        }


        /// <summary>
        /// 线性标注 受力点\矩形开洞\圆形开洞\排水沟\设备位号
        /// </summary>
        public static void DDimLinear(DBTrans tr, string dimString, string dimString2, int layerColorIndex)
        {
            try
            {
                //using var tr = new DBTrans();

                short textColor = layerColorIndex > 0 ? Convert.ToInt16(layerColorIndex) : LayerControlHelper.ResolveLayerColor();
                string targetLayer = LayerControlHelper.GetOrCreateTargetLayer(
                    tr,
                    VariableDictionary.layerName ?? VariableDictionary.btnBlockLayer,
                    textColor);

                string content = BuildContextualDimText(dimString, dimString2, jztjUseMeter: false, includeDeviceCode: true);

                var userPoint1 = Env.Editor.GetPoint("\n请指定标注第一点");
                if (userPoint1.Status != PromptStatus.OK) return;
                var ucsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();

                double uiScale = AutoCadHelper.GetScale(true); // 读取当前界面比例

                double textHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, uiScale); // 文字高度
                double arrowSize = TextFontsStyleHelper.ComputeScaledHeight(2.0, uiScale); // 箭头尺寸
                // 根据用户输入的标注内容和当前场景，构建适合的标注文本，并调用统一的拖拽放置方法创建标注。
                bool ok = TryCreateLeaderByDrag(
                    tr,
                    targetLayer,
                    textColor,
                    ucsUserPoint1,
                    content,
                    textHeight: textHeight,
                    arrowSize: arrowSize,
                    textAttachmentType: TextAttachmentType.AttachmentBottomOfTopLine);

                if (!ok)
                {
                    tr.Abort();
                    return;
                }

                //tr.Commit();
                Env.Editor.Redraw();
                VariableDictionary.dimString = null;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n做标注失败！错误信息: {ex.Message}");
                LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 线性标注（给定首点）
        /// </summary>
        public static void DDimLinear(DBTrans tr, Point3d point3D, string dimString1, string dimString2 = null)
        {
            try
            {
                short textColor = LayerControlHelper.ResolveLayerColor();
                string targetLayer = LayerControlHelper.GetOrCreateTargetLayer(tr, VariableDictionary.btnBlockLayer ?? VariableDictionary.layerName, textColor);

                string content = BuildContextualDimText(dimString1, dimString2);

                double uiScale = AutoCadHelper.GetScale(true); // 读取当前界面比例

                double textHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, uiScale); // 文字高度
                double arrowSize = TextFontsStyleHelper.ComputeScaledHeight(2.0, uiScale); // 箭头尺寸
                // 根据用户输入的标注内容和当前场景，构建适合的标注文本，并调用统一的拖拽放置方法创建标注。
                bool ok = TryCreateLeaderByDrag(
                    tr,
                    targetLayer,
                    textColor,
                    point3D,
                    content,
                    textHeight: textHeight,
                    arrowSize: arrowSize,
                    textAttachmentType: TextAttachmentType.AttachmentBottomOfTopLine);
                if (!ok)
                {
                    tr.Abort();
                    return;
                }

                //tr.Commit();
                Env.Editor.Redraw();
                VariableDictionary.dimString = null;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n做标注失败！错误信息: {ex.Message}");
                LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 线性标注（给定首点+指定颜色）
        /// </summary>
        public static void DDimLinear(DBTrans tr, string dimString1, string dimString2, Int16 textLayerColorIndex, Point3d point3D)
        {
            try
            {
                //using var tr = new DBTrans();
                // 如果用户指定了颜色索引且有效，则使用它；否则解析默认颜色。
                short textColor = textLayerColorIndex > 0 ? textLayerColorIndex : LayerControlHelper.ResolveLayerColor();
                // 确保目标图层存在，并获取图层名称。
                string targetLayer = LayerControlHelper.GetOrCreateTargetLayer(tr, VariableDictionary.btnBlockLayer ?? VariableDictionary.layerName, textColor);
                // 根据用户输入的标注内容和当前场景，构建适合的标注文本，并调用统一的拖拽放置方法创建标注。
                string content = BuildContextualDimText(dimString1, dimString2, jztjUseMeter: true);

                double uiScale = AutoCadHelper.GetScale(true); // 读取当前界面比例

                double textHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, uiScale); // 文字高度
                double arrowSize = TextFontsStyleHelper.ComputeScaledHeight(2.0, uiScale); // 箭头尺寸
                // 根据用户输入的标注内容和当前场景，构建适合的标注文本，并调用统一的拖拽放置方法创建标注。
                bool ok = TryCreateLeaderByDrag(
                    tr,
                    targetLayer,
                    textColor,
                    point3D,
                    content,
                    textHeight: textHeight,
                    arrowSize: arrowSize,
                    textAttachmentType: TextAttachmentType.AttachmentBottomOfTopLine);
                if (!ok)
                {
                    tr.Abort();
                    return;
                }

                //tr.Commit();
                Env.Editor.Redraw();
                VariableDictionary.dimString = null;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n做标注失败！错误信息: {ex.Message}");
                LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 创建标注线 标注入口：解析图层名 + 解析颜色 + 确保图层存在 + 创建标注
        /// </summary>
        /// <param name="DimString">标注内容字符串</param>
        /// <param name="layerColorIndex">图层颜色索引</param>
        /// <param name="layerName">图层名称</param>
        /// <param name="userPoint">用户点击的点</param>
        public static void DDimLinear(DBTrans tr, string DimString, Int16 layerColorIndex, string layerName, Point3d userPoint)
        {
            try
            {
                //using var tr = new DBTrans();

                string targetLayer = LayerControlHelper.GetOrCreateTargetLayer(tr, layerName, layerColorIndex);
                short textColor = layerColorIndex > 0 ? layerColorIndex : LayerControlHelper.ResolveLayerColor();

                string content = "设备名称" + "\n" + $"{DimString}" + " ";

                double uiScale = AutoCadHelper.GetScale(true); // 读取当前界面比例

                double textHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, uiScale); // 文字高度
                double arrowSize = TextFontsStyleHelper.ComputeScaledHeight(2.0, uiScale); // 箭头尺寸
                // 根据用户输入的标注内容和当前场景，构建适合的标注文本，并调用统一的拖拽放置方法创建标注。
                bool ok = TryCreateLeaderByDrag(
                    tr,
                    targetLayer,
                    textColor,
                    userPoint,
                    content,
                    textHeight: textHeight,
                    arrowSize: arrowSize,
                    textAttachmentType: TextAttachmentType.AttachmentBottomOfTopLine);
                if (!ok)
                {
                    tr.Abort();
                    return;
                }

                //tr.Commit();
                Env.Editor.Redraw();
                VariableDictionary.dimString = null;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n做标注失败！错误信息: {ex.Message}");
                LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 创建水平、垂直或任意角度的线性标注对象
        /// </summary>
        /// <param name="pt1">第一点</param>
        /// <param name="layer">图层名</param>
        /// <param name="VariableDictionary.layerColorIndex">图层颜色</param>
        public static void DDimLinear(DBTrans tr, string layerName, Int16 layerColorIndex, Point3d pt1)
        {
            try
            {
                //开启事务
                //var tr = new DBTrans();
                TextFontsStyleHelper.TextStyleAndLayerInfo(tr, layerName, layerColorIndex, "tJText");
                var textBoxScale = AutoCadHelper.GetScale(true);//文本框缩放比例，基于界面比例进行调整，确保在不同缩放级别下标注具有合适的大小。  
                // 先确保“洞口标注”样式存在并更新参数
                ObjectId dimStyleId = DimStyleHelper.EnsureOrCreateDimStyle(
                    tr,
                    "JLPDI-定位",
                    layerColorIndex,
                    textBoxScale,
                    tr.TextStyleTable["tJText"]);

                // 再创建 RotatedDimension，并直接套样式
                RotatedDimension rDim = new RotatedDimension();
                rDim.DimensionStyle = dimStyleId;                 // 关键：用样式Id
                //rDim.DimensionStyleName = "洞口标注";              // 保留名称便于排查
                rDim.DimensionText = null;                        // 先不设置文本，让动态标注时自动生成
                rDim.Dimclrt = Color.FromColorIndex(ColorMethod.ByColor, layerColorIndex); // 通过工厂方法生成一个红色的 Color 对象（1 = 红色索引）                                                                                           
                rDim.Annotative = AnnotativeStates.True;// 启用注释缩放（让对象在布局里按注释比例自适应）              
                rDim.LinetypeScale = 1 * textBoxScale;
                rDim.Layer = layerName;
                rDim.TextStyleId = tr.TextStyleTable["tJText"];
                rDim.XLine1Point = pt1;
                // 打开正交模式
                //Env.OrthoMode = true;
                // 关闭正交模式
                //Env.OrthoMode = false;
                //动态显示标注并提示指定第二点；
                using var dimPoint2 = new JigEx((mpw, Queue) =>
                {
                    var pt3 = mpw.Z20(); // 当前鼠标点（第二个点或标注放置点）
                                         // 计算原始距离（以当前图纸单位为准）
                    double rawDistance = pt1.DistanceTo(pt3);

                    // 用自定义逻辑向上取整到步长 50（例如：101->150, 151->200）
                    // 若需要其它规则，可把 50 替换为变量或从配置读取
                    long roundedValue = DimStyleHelper.RoundUpToStep(rawDistance, 50);

                    // 设置标注显示文本（整数）
                    rDim.DimensionText = roundedValue.ToString();

                    // 其余保持原行为
                    rDim.XLine2Point = pt3;
                    rDim.DimLinePoint = pt3;
                    rDim.Annotative = AnnotativeStates.True; // 保持注释缩放设置
                    rDim.LinetypeScale = 1 * textBoxScale;
                    rDim.Dimasz = 2 * textBoxScale; // 控制引线箭头的大小
                    rDim.Dimtxt = 3.5 * textBoxScale; // 标注文字高度（样式 TextSize=0 时生效）
                    rDim.Dimexo = 3 * textBoxScale; // 尺寸界线偏移
                    rDim.Dimgap = 2 * textBoxScale; // 标注文字偏移量
                    rDim.TextStyleId = tr.TextStyleTable["tJText"];
                });
                dimPoint2.DatabaseEntityDraw(WorldDraw => WorldDraw.Geometry.Draw(rDim));
                dimPoint2.SetOptions(pt1, msg: "\n请指定方形洞口第二点");
                var userPoint2 = dimPoint2.Drag();//拿到的第二点；
                if (userPoint2.Status != PromptStatus.OK) return;
                // 计算旋转角度
                rDim.Rotation = pt1.GetVectorTo(dimPoint2.MousePointWcsLast).GetAngleTo(Vector3d.XAxis);
                SetDimensionRotationToNearest90Degrees(rDim, dimPoint2.MousePointWcsLast);
                // 提示用户指定标注位置
                using var dimTextPoint = new JigEx((mpw, Queue) =>
                {
                    var pt3 = mpw.Z20();
                    rDim.DimLinePoint = pt3;//标注点
                    rDim.Annotative = AnnotativeStates.True;// 启用注释缩放（让对象在布局里按注释比例自适应）
                    rDim.LinetypeScale = 1 * textBoxScale;
                    rDim.Dimasz = 2 * textBoxScale;//控制引线箭头的大小                                            
                    rDim.Dimtxt = 3.5 * textBoxScale;// Dimtxt 指定标注文字的高度，除非当前文字样式具有固定的高度
                    rDim.Dimexo = 3 * textBoxScale;//尺寸界线偏移
                    rDim.Dimgap = 2 * textBoxScale;// 标注文字偏移量
                    rDim.TextStyleId = tr.TextStyleTable["tJText"];
                });
                dimTextPoint.DatabaseEntityDraw(WorldDraw => WorldDraw.Geometry.Draw(rDim));
                dimTextPoint.SetOptions(dimPoint2.MousePointWcsLast, msg: "\n请选择标注文字位置");
                var userPoint3 = dimTextPoint.Drag();//拿到的标注文字点；
                if (userPoint3.Status != PromptStatus.OK) return;
                //double angle = pt1.GetVectorTo(dimPoint2.MousePointWcsLast).GetAngleTo(Vector3d.XAxis);
                //拿到鼠标位置
                rDim.TextStyleId = tr.TextStyleTable["tJText"];
                rDim.DimLinePoint = dimTextPoint.MousePointWcsLast;
                //rDim.Annotative = AnnotativeStates.True;// 启用注释缩放（让对象在布局里按注释比例自适应）    
                // 添加标注对象到模型空间
                tr.CurrentSpace.AddEntity(rDim);
                //正交关闭
                Env.OrthoMode = false;
                Env.Editor.Redraw();
                VariableDictionary.dimString = null;//清空标注内容，避免重复标注
            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("给出一点与图层名，做标注失败！");
                LogManager.Instance.LogInfo($"{ex.Message}");
            }
        }


        /// <summary>
        /// com接口(设置)对象属性值，类似VisualLisp的vlax-put-property函数
        /// </summary>
        /// <param name="obj">对象</param>
        /// <param name="key">属性名称</param>
        /// <param name="value">属性值</param>
        /// <summary>
        /// 给出一点与图层名，做标注
        /// </summary>
        /// <param name="pt1">标注的另一点</param>
        /// <param name="layer">图层名</param>
        [CommandMethod("DDimLinearP")]
        public static void DDimLinearP()
        {
            try
            {
                #region 创建文字属性

                using var tr = new DBTrans();
                TextFontsStyleHelper.TextStyleAndLayerInfo(tr, VariableDictionary.btnBlockLayer, Convert.ToInt16(VariableDictionary.layerColorIndex), "tJText");
                var mld = new MLeader
                {
                    Layer = VariableDictionary.btnBlockLayer,//设置多重引线的图层
                    //TextAttachmentType = TextAttachmentType.AttachmentBottomOfTopLine,//设置多重引线的标注文字下是不是有引线；
                    TextAttachmentType = TextAttachmentType.AttachmentBottomLine,//设置多重引线的标注文字下是不是有引线；
                    ContentType = ContentType.MTextContent,//内容类型
                    ColorIndex = Convert.ToInt32(VariableDictionary.layerColorIndex),
                    // 例如索引3通常代表绿色
                    LeaderLineColor = Color.FromColorIndex(ColorMethod.ByAci, Convert.ToInt16(VariableDictionary.layerColorIndex)),
                    //LeaderLineTypeId = Env.Database.LinetypeTableId, // 使用默认线型
                    //LeaderLineWeight = LineWeight.LineWeight030, // 设置引线线宽
                    //Scale = 1.0,// 设置多重引线的比例   
                };
                var userPoint1 = Env.Editor.GetPoint("\n请指定标注第一点");
                if (userPoint1.Status != PromptStatus.OK) return;
                var UcsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();
                //标注样式
                MText mt = new MText();
                //标注文字获取
                mt.Contents = VariableDictionary.btnFileName;
                mt.Attachment = AttachmentPoint.MiddleCenter; // 设置标注文字居中对齐  
                mld.MText = mt;//赋值标注文字样式
                mld.TextHeight = 300;//设置多重引线标注文字的高度
                mld.TextStyleId = tr.TextStyleTable["tJText"];
                // 设置引线颜色为 7  
                mld.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                // 设置箭头大小为 300  
                mld.ArrowSize = 250;
                // 添加引线和引线段  
                int ldNum = mld.AddLeader();
                int lnNum = mld.AddLeaderLine(ldNum);
                mld.AddFirstVertex(lnNum, UcsUserPoint1);  // 引线起始点（UCS 坐标）  
                using var mleaderjig = new JigEx((mpw, Queue) =>
                {
                    var pt2Ucs = mpw.Z20();


                    mld.TextLocation = pt2Ucs;  // 标注文字显示位置  
                });
                var UcsUserPoint2 = mleaderjig.MousePointWcsLast;
                mld.AddLastVertex(lnNum, UcsUserPoint2);// 引线结束点（UCS 坐标）  

                mleaderjig.DatabaseEntityDraw(wb => wb.Geometry.Draw(mld));
                mleaderjig.SetOptions(UcsUserPoint1, msg: "点选标注第二点");
                var pt2 = mleaderjig.Drag();
                if (pt2.Status != PromptStatus.OK) return;

                tr.CurrentSpace.AddEntity(mld);
                #endregion
                tr.Commit();
                Env.Editor.Redraw();//重新刷新
                VariableDictionary.dimString = null;//清空标注内容，避免重复标注
            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo($"\n做标注失败！错误信息: {ex.Message}");
                LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 点标注
        /// </summary>
        [CommandMethod(nameof(PointDim))]
        public static void PointDim(DBTrans tr, Point3d UcsUserPoint1, string dimX, string dimY, string dimFl, string layerName, Int16 layerColorIndex)
        {
            try
            {
                short textColor = layerColorIndex > 0 ? layerColorIndex : LayerControlHelper.ResolveLayerColor();
                string targetLayer = LayerControlHelper.GetOrCreateTargetLayer(tr, layerName, textColor);
                string content = $"{dimX}{dimY}{dimFl}";
                double uiScale = AutoCadHelper.GetScale(true); // 读取当前界面比例

                double textHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, uiScale); // 文字高度
                double arrowSize = TextFontsStyleHelper.ComputeScaledHeight(2.0, uiScale); // 箭头尺寸
                // 根据用户输入的标注内容和当前场景，构建适合的标注文本，并调用统一的拖拽放置方法创建标注。
                bool ok = TryCreateLeaderByDrag(
                    tr,
                    targetLayer,
                    textColor,
                    UcsUserPoint1,
                    content,
                    textHeight: textHeight,
                    arrowSize: arrowSize,
                    textAttachmentType: TextAttachmentType.AttachmentBottomOfTopLine);
                if (!ok)
                {
                    tr.Abort();
                    return;
                }

                //tr.Commit();
                Env.Editor.Redraw();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"标注失败！错误信息: {ex.Message}");
            }
        }



        /// <summary>
        /// 标注
        /// </summary>
        /// <param name="pt1Ucs">第一点坐标</param>
        /// <param name="dimX">X坐标</param>
        /// <param name="dimY">Y坐标</param>
        /// <param name="dimFl">标注文字</param>
        /// <param name="layerName">图层名</param>
        /// <param name="VariableDictionary.layerColorIndex">图层颜色</param>
        //[CommandMethod(nameof(PointDim))]
        //public static void PointDim(DBTrans tr, Point3d UcsUserPoint1, string dimX, string dimY, string dimFl, string layerName, Int16 layerColorIndex)
        //{
        //    try
        //    {
        //        //using var tr = new DBTrans();
        //        TextStyleAndLayerInfo(tr,layerName, Convert.ToInt16(VariableDictionary.layerColorIndex), "tJText");
        //        var mld = new MLeader
        //        {
        //            Layer = layerName,//设置多重引线的图层
        //            ColorIndex = VariableDictionary.layerColorIndex,
        //            TextAttachmentType = TextAttachmentType.AttachmentBottomOfTopLine,//设置多重引线的标注文字下是不是有引线；
        //            ContentType = ContentType.MTextContent,//内容类型
        //            LeaderLineColor = Color.FromColorIndex(ColorMethod.ByAci, Convert.ToInt16(VariableDictionary.layerColorIndex)),// 例如索引3通常代表绿色
        //        };
        //        //标注样式
        //        MText mt = new MText();
        //        TextStyleAndLayerInfo(tr, VariableDictionary.btnBlockLayer, Convert.ToInt16(VariableDictionary.layerColorIndex), "tJText", ref mt, ref mld);
        //        mt.Attachment = AttachmentPoint.MiddleCenter; // 设置标注文字居中对齐  
        //                                                      // 添加引线和引线段  
        //        int ldNum = mld.AddLeader();
        //        int lnNum = mld.AddLeaderLine(ldNum);
        //        mld.AddFirstVertex(lnNum, UcsUserPoint1);  // 引线起始点（UCS 坐标）
        //        var mpwUcs = new Point3d(0, 0, 0);
        //        using var mleaderjig = new JigEx((mpw, _) =>
        //        {
        //            // 引线结束点（UCS 坐标）
        //            mpwUcs = mpw.Z20();
        //            // 标注文字显示位置        
        //            mld.TextLocation = mpwUcs;
        //        });
        //        var UcsUserPoint2 = mleaderjig.MousePointWcsLast;
        //        //标注文字
        //        mt.Contents = dimX + dimY + dimFl;
        //        //标注文字高度
        //        mt.Height = 300;
        //        mt.ColorIndex = VariableDictionary.layerColorIndex;
        //        mld.AddLastVertex(lnNum, UcsUserPoint2);
        //        mld.MText = mt;
        //        mld.TextHeight = 300;
        //        mld.TextStyleId = tr.TextStyleTable["tJText"];
        //        mleaderjig.DatabaseEntityDraw(wb => wb.Geometry.Draw(mld));
        //        mleaderjig.SetOptions(UcsUserPoint1, msg: "\n标注文字的位置");
        //        var userPoint2 = Env.Editor.Drag(mleaderjig);
        //        if (userPoint2.Status != PromptStatus.OK) return;

        //        tr.CurrentSpace.AddEntity(mld);
        //        //tr.Commit();
        //        Env.Editor.Redraw();
        //    }
        //    catch (Exception ex)
        //    {
        //        // 记录错误日志  
        //        LogManager.Instance.LogInfo($"标注失败！错误信息: {ex.Message}"); // 输出错误信息  
        //    }
        //}

        #endregion



        /// <summary>
        /// 块统计
        /// </summary>
        public void BlockCountStatistics()
        {
            try
            {
                using var tr = new DBTrans();
                var i = tr.CurrentSpace
                        .GetEntities<BlockReference>()
                        .Where(brf => brf.GetBlockName() == "自定义块")
                        .Count();
                Env.Print(i);
                LogManager.Instance.LogInfo($"\n块统计完成，找到 {i} 个自定义块。");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"块统计时出错: {ex.Message}");
                LogManager.Instance.LogError($"错误堆栈: {ex.StackTrace}");
            }
        }

        #region 选择外参并从中选图元复制到当前空间内

        #region 选中外部参照中的图元
        [CommandMethod("SELECTXREFENTITY")]
        public void SelectXrefEntity()
        {
            //try
            //{
            //    FormMain.ReferenceEntity.Items.Clear();
            //    FormMain.SelectEntity.Items.Clear();
            //    FormMain.Reference.Items.Clear();
            //    // 第一步：创建嵌套实体选择选项
            //    PromptNestedEntityOptions options = new PromptNestedEntityOptions("\n请点击外部参照中的图元: ");
            //    // 允许用户选择任意层级的嵌套实体（外部参照中的实体）
            //    options.AllowNone = false;
            //    // 第二步：获取用户选择的嵌套实体
            //    PromptNestedEntityResult result = Env.Editor.GetNestedEntity(options);
            //    if (result.Status != PromptStatus.OK) return;

            //    using (var tr = new DBTrans())
            //    {
            //        // 第三步：获取选中的嵌套图元ObjectId
            //        ObjectId nestedId = result.ObjectId;
            //        // 获取外部参照的变换矩阵（包含位置/旋转/缩放信息）
            //        Matrix3d transform = result.Transform;
            //        // 获取鼠标点击位置（WCS坐标）
            //        Point3d pickPoint = result.PickedPoint;
            //        // 第四步：打开嵌套图元
            //        Entity nestedEntity = .GetObject(nestedId, OpenMode.ForRead) as Entity;
            //        if (nestedEntity == null)
            //        {
            //            LogManager.Instance.LogInfo("\n错误：选中的对象不是图元。");
            //            return;
            //        }

            //        using var tr = new DBTrans();

            //        // 获取外部参照中的图元
            //        xrefEntities = Command.GetXrefEntities(res.ObjectId);
            //        // 第五步：获取外部参照名称
            //        string xrefName = Command.getXrefName(tr, res.ObjectId);
            //        Reference.Items.Add(xrefName);
            //        // 添加图元到左侧列表
            //        foreach (ObjectId entityId in xrefEntities)
            //        {
            //            Entity entity = Command.GetEntity(entityId);
            //            if (entity is not null)
            //            {
            //                selectedEntities.Add(entityId);
            //                ReferenceEntity.Items.Add(Command.getXrefName(tr, entityId));
            //            }
            //        }

            //        // 第七步：提交事务
            //        tr.Commit();
            //        第八步：刷新界面
            //        Env.Editor.Redraw();

            //        LogManager.Instance.LogInfo("\n成功复制图元到当前位置！");
            //    }
            //     LogManager.Instance.LogInfo("\n外部参照图元选择完成。");
            //    }
            //catch (Exception ex)
            //{
            //    LogManager.Instance.LogError($"选择外部参照图元时出错: {ex.Message}");
            //    LogManager.Instance.LogError($"错误堆栈: {ex.StackTrace}");
            //}
        }
        #endregion


        #region 选图层分解块

        /// <summary>
        /// 交互式分解指定图层中的所有块参照
        /// </summary>
        [CommandMethod("ExplodeBlocksInLayerInteractive")]
        public static void ExplodeBlocksInLayerInteractive(double clearthreshold = 20)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                using (doc.LockDocument())
                using (var tr = new DBTrans())
                {
                    // 让用户通过选实体来确定目标图层
                    var peo = new PromptEntityOptions("\n请选择一个实体以确定要分解的图层: ");
                    peo.SetRejectMessage("\n请选择一个实体！");
                    var per = Env.Editor.GetEntity(peo);

                    if (per.Status != PromptStatus.OK)
                    {
                        Env.Editor.WriteMessage("\n未选择任何实体或操作已取消。");
                        return;
                    }

                    var selectedEntity = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                    if (selectedEntity == null)
                    {
                        Env.Editor.WriteMessage("\n选择的不是一个有效的实体！");
                        return;
                    }

                    string selectedLayer = selectedEntity.Layer;
                    Env.Editor.WriteMessage($"\n选择的图层: {selectedLayer}");

                    // 获取该图层内所有块参照
                    var blockRefs = GetBlockReferencesInLayer(tr, selectedLayer);
                    if (blockRefs.Count == 0)
                    {
                        Env.Editor.WriteMessage($"\n图层 '{selectedLayer}' 中没有找到任何块参照！");
                        LogManager.Instance.LogInfo($"图层 '{selectedLayer}' 中没有找到任何块参照！");
                        return;
                    }

                    // 与单块逻辑一致：询问是否设备块（这里只询问一次，应用到本次批处理）
                    string targetLayer = selectedLayer;
                    var dialogResult = System.Windows.Forms.MessageBox.Show(
                        "要分解的块是不是设备块:",
                        "分解确认",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Question);

                    if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                    {
                        if (string.Equals(selectedLayer, "设备", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(selectedLayer, "SB", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(selectedLayer, "设备名称", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(selectedLayer, "S_设备", StringComparison.OrdinalIgnoreCase))
                        {
                            targetLayer = "S_设备";

                            // 确保图层存在
                            LayerDictionaryHelper.EnsureTargetLayer(tr, targetLayer, 140);
                            Env.Editor.WriteMessage($"\n识别为设备块批处理，图层由 '{selectedLayer}' 调整为 '{targetLayer}'。");
                        }
                    }

                    // 确认提示
                    var options = new PromptKeywordOptions($"\n确定要处理图层 '{selectedLayer}' 中的 {blockRefs.Count} 个块吗？[Yes/No]");
                    options.Keywords.Add("Yes");
                    options.Keywords.Add("No");
                    options.Keywords.Default = "Yes";

                    var confirm = Env.Editor.GetKeywords(options);
                    if (confirm.Status == PromptStatus.OK && confirm.StringResult == "No")
                    {
                        Env.Editor.WriteMessage("\n用户取消了分解操作。");
                        return;
                    }

                    if (!VariableDictionary.winForm_Status)
                    {
                        clearthreshold = GetCleanupThreshold();
                    }

                    int processedCount = 0;
                    int failedCount = 0;

                    // 关键：循环调用与单块相同的递归方法
                    foreach (var blockRefId in blockRefs)
                    {
                        try
                        {
                            var blockRef = tr.GetObject(blockRefId, OpenMode.ForRead) as BlockReference;
                            if (blockRef == null || blockRef.IsErased || blockRef.IsDisposed)
                                continue;

                            var ids = RecursivelyExplodeBlocks(tr, blockRef, targetLayer, clearthreshold);
                            if (ids.Count > 0)
                            {
                                processedCount++;
                            }
                            else
                            {
                                failedCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogManager.Instance.LogError($"批量处理块时出错 (ID: {blockRefId}): {ex.Message}");
                            failedCount++;
                        }
                    }

                    tr.Commit();

                    Env.Editor.WriteMessage($"\n批量处理完成！成功: {processedCount}，失败: {failedCount}，图层: {selectedLayer} -> {targetLayer}");
                    LogManager.Instance.LogInfo($"批量处理完成！成功: {processedCount}，失败: {failedCount}，图层: {selectedLayer} -> {targetLayer}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"交互式分解图层块时出错: {ex.Message}");
                Env.Editor.WriteMessage($"\n错误: {ex.Message}");
            }
        }


        //[CommandMethod("ExplodeBlocksInLayerInteractive")]
        //public static void ExplodeBlocksInLayerInteractive(double clearthreshold = 20)
        //{
        //    try
        //    {
        //        using var tr = new DBTrans();

        //        // 获取当前文档中的所有图层
        //        var layers = GetLayersFromCurrentDocument(tr);

        //        if (layers.Count == 0)
        //        {
        //            Env.Editor.WriteMessage("\n当前文档中没有找到任何图层！");
        //            LogManager.Instance.LogInfo("当前文档中没有找到任何图层！");
        //            return;
        //        }

        //        // 让用户选择图层
        //        var peo = new PromptEntityOptions("\n请选择一个实体以确定要分解的图层: ");
        //        peo.SetRejectMessage("\n请选择一个实体！");

        //        var per = Env.Editor.GetEntity(peo);
        //        if (per.Status != PromptStatus.OK)
        //        {
        //            Env.Editor.WriteMessage("\n未选择任何实体或操作已取消。");
        //            return;
        //        }

        //        // 获取选择的实体所在的图层
        //        var selectedEntity = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
        //        if (selectedEntity == null)
        //        {
        //            Env.Editor.WriteMessage("\n选择的不是一个有效的实体！");
        //            return;
        //        }

        //        var layerName = selectedEntity.Layer;
        //        Env.Editor.WriteMessage($"\n选择的图层: {layerName}");

        //        // 获取指定图层中的所有块参照
        //        var blockRefs = GetBlockReferencesInLayer(tr, layerName);

        //        if (blockRefs.Count == 0)
        //        {
        //            Env.Editor.WriteMessage($"\n图层 '{layerName}' 中没有找到任何块参照！");
        //            LogManager.Instance.LogInfo($"图层 '{layerName}' 中没有找到任何块参照！");
        //            return;
        //        }

        //        // 询问用户是否确认分解
        //        var options = new PromptKeywordOptions($"\n确定要分解图层 '{layerName}' 中的 {blockRefs.Count} 个块吗？[Yes/No]");
        //        options.Keywords.Add("Yes");
        //        options.Keywords.Add("No");
        //        options.Keywords.Default = "Yes";

        //        var result = Env.Editor.GetKeywords(options);
        //        if (result.Status == PromptStatus.OK && result.StringResult == "No")
        //        {
        //            Env.Editor.WriteMessage("\n用户取消了分解操作。");
        //            return;
        //        }

        //        if (!VariableDictionary.winForm_Status)
        //        {
        //            // 获取清理参数阈值
        //            clearthreshold = GetCleanupThreshold();
        //        }

        //        // 获取清理参数阈值
        //        //var threshold = GetCleanupThreshold();

        //        int processedCount = 0;
        //        int failedCount = 0;

        //        foreach (var blockRefId in blockRefs)
        //        {
        //            try
        //            {
        //                var blockRef = tr.GetObject(blockRefId, OpenMode.ForWrite) as BlockReference;
        //                if (blockRef != null)
        //                {
        //                    // 使用DBObjectCollection来接收分解后的实体
        //                    var entitySet = new DBObjectCollection();
        //                    blockRef.Explode(entitySet);

        //                    // 将分解后的实体添加到当前空间（应用过滤）
        //                    foreach (Entity entity in entitySet)
        //                    {
        //                        // 应用过滤逻辑
        //                        if (ShouldKeepEntity(entity, clearthreshold))
        //                        {
        //                            // 设置图层
        //                            entity.Layer = blockRef.Layer;
        //                            // 使用IFoxCAD风格添加实体到当前空间
        //                            var newId = tr.CurrentSpace.AddEntity(entity);
        //                        }
        //                        // 如果不满足过滤条件，则不添加到图纸中（相当于过滤掉）
        //                    }

        //                    // 删除原始的块参照
        //                    blockRef.Erase();

        //                    processedCount++;
        //                }
        //            }
        //            catch (System.Exception ex)
        //            {
        //                LogManager.Instance.LogError($"分解块时出错 (ID: {blockRefId}): {ex.Message}");
        //                failedCount++;
        //            }
        //        }

        //        tr.Commit();

        //        Env.Editor.WriteMessage($"\n分解完成！成功处理: {processedCount} 个块，失败: {failedCount} 个块，图层: {layerName}");
        //        LogManager.Instance.LogInfo($"分解完成！成功处理: {processedCount} 个块，失败: {failedCount} 个块，图层: {layerName}");
        //    }
        //    catch (System.Exception ex)
        //    {
        //        LogManager.Instance.LogError($"交互式分解图层块时出错: {ex.Message}");
        //        Env.Editor.WriteMessage($"\n错误: {ex.Message}");
        //    }
        //}

        /// <summary>
        /// 分解指定图层中的所有块参照
        /// </summary>
        /// <param name="layerName">要分解的图层名称</param>
        [CommandMethod("ExplodeBlocksInLayer")]
        public static void ExplodeBlocksInLayer(string layerName)
        {
            try
            {
                using var tr = new DBTrans();

                // 获取指定图层中的所有块参照
                var blockRefs = GetBlockReferencesInLayer(tr, layerName);

                if (blockRefs.Count == 0)
                {
                    Env.Editor.WriteMessage($"\n图层 '{layerName}' 中没有找到任何块参照！");
                    LogManager.Instance.LogInfo($"图层 '{layerName}' 中没有找到任何块参照！");
                    return;
                }

                // 获取清理参数阈值
                var threshold = GetCleanupThreshold();

                int processedCount = 0;
                int failedCount = 0;

                foreach (var blockRefId in blockRefs)
                {
                    try
                    {
                        var blockRef = tr.GetObject(blockRefId, OpenMode.ForWrite) as BlockReference;
                        if (blockRef != null)
                        {
                            // 使用DBObjectCollection来接收分解后的实体
                            var entitySet = new DBObjectCollection();
                            blockRef.Explode(entitySet);

                            // 将分解后的实体添加到当前空间（应用过滤）
                            foreach (Entity entity in entitySet)
                            {
                                // 应用过滤逻辑
                                if (ShouldKeepEntity(entity, threshold))
                                {
                                    // 设置图层
                                    entity.Layer = blockRef.Layer;
                                    // 使用IFoxCAD风格添加实体到当前空间
                                    var newId = tr.CurrentSpace.AddEntity(entity);
                                }
                                // 如果不满足过滤条件，则不添加到图纸中（相当于过滤掉）
                            }

                            // 删除原始的块参照
                            blockRef.Erase();

                            processedCount++;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        LogManager.Instance.LogError($"分解块时出错 (ID: {blockRefId}): {ex.Message}");
                        failedCount++;
                    }
                }

                tr.Commit();

                Env.Editor.WriteMessage($"\n分解完成！成功处理: {processedCount} 个块，失败: {failedCount} 个块，图层: {layerName}");
                LogManager.Instance.LogInfo($"分解完成！成功处理: {processedCount} 个块，失败: {failedCount} 个块，图层: {layerName}");
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogError($"分解图层块时出错: {ex.Message}");
                Env.Editor.WriteMessage($"\n错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 从数据库事务中获取所有图层名称
        /// </summary>
        /// <param name="tr">数据库事务</param>
        /// <returns>图层名称列表</returns>
        private static List<string> GetLayersFromCurrentDocument(DBTrans tr)
        {
            var layers = new List<string>();

            try
            {
                var layerTable = tr.GetObject(tr.Database.LayerTableId, OpenMode.ForRead) as LayerTable;

                if (layerTable != null)
                {
                    foreach (ObjectId layerId in layerTable)
                    {
                        var layerRecord = layerId.GetObject(OpenMode.ForRead) as LayerTableRecord;
                        if (layerRecord != null && !layerRecord.IsErased)
                        {
                            layers.Add(layerRecord.Name);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogError($"获取图层列表时出错: {ex.Message}");
            }

            return layers;
        }

        /// <summary>
        /// 获取指定图层中的所有块参照
        /// </summary>
        /// <param name="tr">数据库事务</param>
        /// <param name="layerName">图层名称</param>
        /// <returns>块参照ObjectId列表</returns>
        private static List<ObjectId> GetBlockReferencesInLayer(DBTrans tr, string layerName)
        {
            var blockRefs = new List<ObjectId>();

            try
            {
                // 使用IFoxCAD框架的扩展方法遍历模型空间中的所有实体
                foreach (var entId in tr.CurrentSpace)
                {
                    var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                    if (ent != null && ent.Layer == layerName)
                    {
                        // 如果是块参照，添加到列表
                        if (ent is BlockReference blockRef)
                        {
                            blockRefs.Add(entId);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogError($"获取图层 '{layerName}' 中的块参照时出错: {ex.Message}");
            }

            return blockRefs;
        }

        /// <summary>
        /// 判断实体是否应该包含在分解结果中（过滤小图元，但保留文本）
        /// </summary>
        /// <param name="entity">要检查的实体</param>
        /// <param name="threshold">清理阈值</param>
        /// <returns>如果应该包含则返回true，否则返回false</returns>
        private static bool ShouldIncludeEntity(Entity entity, double threshold)
        {
            // 文本类实体直接保留（不进行大小过滤）
            if (entity is DBText || entity is MText)
            {
                return true;
            }

            switch (entity)
            {
                case Line line:
                    // 检查线段长度
                    var length = line.StartPoint.DistanceTo(line.EndPoint);
                    return length >= threshold;

                case Arc arc:
                    // 检查弧长和半径
                    var arcLength = arc.Radius * System.Math.Abs(arc.EndAngle - arc.StartAngle);
                    return arcLength >= threshold && arc.Radius >= threshold;

                case Circle circle:
                    // 检查圆的半径
                    return circle.Radius >= threshold;

                case Ellipse ellipse:
                    // 检查椭圆的大小
                    var majorRadius = ellipse.MajorAxis.Length * ellipse.RadiusRatio;
                    var minorRadius = ellipse.MajorAxis.Length;
                    return System.Math.Max(majorRadius, minorRadius) >= threshold;

                case Polyline polyline:
                    // 检查多段线的整体尺寸
                    if (polyline.NumberOfVertices <= 1) return false;

                    // 检查相邻顶点间的距离
                    for (int i = 0; i < polyline.NumberOfVertices - 1; i++)
                    {
                        var pt1 = polyline.GetPointAtParameter(i);
                        var pt2 = polyline.GetPointAtParameter(i + 1);
                        if (pt1.DistanceTo(pt2) >= threshold)
                        {
                            return true; // 如果有任何一段大于阈值，则保留整个多段线
                        }
                    }
                    return false;

                case Spline spline:
                    // 对于样条曲线，检查其几何范围
                    try
                    {
                        var extents = spline.GeometricExtents;
                        if (extents == null) return false;

                        var width = extents.MaxPoint.X - extents.MinPoint.X;
                        var height = extents.MaxPoint.Y - extents.MinPoint.Y;
                        var maxLength = System.Math.Max(width, height);

                        return maxLength >= threshold;
                    }
                    catch
                    {
                        // 如果无法获取几何范围，根据其他方法判断
                        return true; // 默认保留
                    }

                default:
                    // 对于其他类型的实体，通常保留
                    return true;
            }
        }

        /// <summary>
        /// 判断实体是否应该保留（基于清理阈值）
        /// </summary>
        /// <param name="entity">要检查的实体</param>
        /// <param name="threshold">清理阈值</param>
        /// <returns>如果应该保留则返回true，否则返回false</returns>
        private static bool ShouldKeepEntity(Entity entity, double threshold)
        {
            // 文本类实体直接保留（不进行大小过滤）
            if (entity is DBText || entity is MText)
            {
                return true;
            }

            switch (entity)
            {
                case Line line:
                    // 检查线段长度
                    var length = line.StartPoint.DistanceTo(line.EndPoint);
                    return length >= threshold;

                case Arc arc:
                    // 检查弧长和半径
                    var arcLength = arc.Radius * System.Math.Abs(arc.EndAngle - arc.StartAngle);
                    return arcLength >= threshold && arc.Radius >= threshold;

                case Circle circle:
                    // 检查圆的半径
                    return circle.Radius >= threshold;

                case Ellipse ellipse:
                    // 检查椭圆的大小
                    var majorRadius = ellipse.MajorAxis.Length * ellipse.RadiusRatio;
                    var minorRadius = ellipse.MajorAxis.Length;
                    return System.Math.Max(majorRadius, minorRadius) >= threshold;

                case Polyline polyline:
                    // 检查多段线的整体尺寸
                    if (polyline.NumberOfVertices <= 1) return false;

                    // 检查相邻顶点间的距离
                    for (int i = 0; i < polyline.NumberOfVertices - 1; i++)
                    {
                        var pt1 = polyline.GetPointAtParameter(i);
                        var pt2 = polyline.GetPointAtParameter(i + 1);
                        if (pt1.DistanceTo(pt2) >= threshold)
                        {
                            return true; // 如果有任何一段大于阈值，则保留整个多段线
                        }
                    }
                    return false;

                case Spline spline:
                    // 对于样条曲线，检查其几何范围
                    try
                    {
                        var extents = spline.GeometricExtents;
                        if (extents == null) return false;

                        var width = extents.MaxPoint.X - extents.MinPoint.X;
                        var height = extents.MaxPoint.Y - extents.MinPoint.Y;
                        var maxLength = System.Math.Max(width, height);

                        return maxLength >= threshold;
                    }
                    catch
                    {
                        // 如果无法获取几何范围，根据其他方法判断
                        return true; // 默认保留
                    }

                default:
                    // 对于其他类型的实体，通常保留
                    return true;
            }
        }



        #endregion


        #region 删除选定图层中的所有块参照

        /// <summary>
        /// 删除用户选定图元所在图层中的所有图元（仅当前空间），但不删除图层本身
        /// </summary>
        [CommandMethod("DeleteSelectLayerContent")]
        public static void DeleteSelectLayerContent()
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                using (doc.LockDocument())
                using (var tr = new DBTrans())
                {
                    // 1) 选择一个图元，用于确定目标图层
                    var peo = new PromptEntityOptions("\n请选择一个图元（将删除其所在图层在当前空间中的全部图元）: ");
                    peo.SetRejectMessage("\n请选择有效图元。");
                    var per = Env.Editor.GetEntity(peo);
                    if (per.Status != PromptStatus.OK)
                    {
                        Env.Editor.WriteMessage("\n未选择图元，操作取消。");
                        return;
                    }

                    var selectedEnt = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                    if (selectedEnt == null || string.IsNullOrWhiteSpace(selectedEnt.Layer))
                    {
                        Env.Editor.WriteMessage("\n无法获取所选图元图层，操作取消。");
                        return;
                    }

                    string targetLayer = selectedEnt.Layer;

                    // 2) 二次确认（默认 No，避免误删）
                    var kopt = new PromptKeywordOptions($"\n确认删除当前空间中图层“{targetLayer}”的所有图元吗？[Yes/No] <No>: ");
                    kopt.Keywords.Add("Yes");
                    kopt.Keywords.Add("No");
                    kopt.Keywords.Default = "Yes";

                    var kret = Env.Editor.GetKeywords(kopt);
                    if (kret.Status != PromptStatus.OK || !string.Equals(kret.StringResult, "Yes", StringComparison.OrdinalIgnoreCase))
                    {
                        Env.Editor.WriteMessage("\n已取消。");
                        return;
                    }

                    // 3) 仅遍历当前空间，删除该图层上的实体
                    int deleted = 0;
                    int failed = 0;

                    foreach (ObjectId id in tr.CurrentSpace)
                    {
                        try
                        {
                            var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (ent == null || ent.IsErased) continue;
                            if (!string.Equals(ent.Layer, targetLayer, StringComparison.OrdinalIgnoreCase)) continue;

                            ent.UpgradeOpen();
                            ent.Erase();
                            deleted++;
                        }
                        catch
                        {
                            // 单个实体失败不影响整体
                            failed++;
                        }
                    }

                    // 4) 提交事务：这里只删除实体，不会删除图层定义
                    tr.Commit();
                    Env.Editor.Redraw();

                    Env.Editor.WriteMessage($"\n完成：已删除图层“{targetLayer}”图元 {deleted} 个，失败 {failed} 个。图层本身未删除。");
                    LogManager.Instance.LogInfo($"DeleteSelectLayerContent: Layer={targetLayer}, Deleted={deleted}, Failed={failed}, LayerKept=True");
                }
            }
            catch (Exception ex)
            {
                AutoCadHelper.LogWithSafety($"DeleteSelectLayerContent 执行失败: {ex.Message}");
                Env.Editor.WriteMessage($"\n错误: {ex.Message}");
            }
        }

        #endregion


        #region 分解块方法二

        /// <summary>
        /// 分解嵌套块（优化版本 - 过滤小图元以减小文件大小）
        /// </summary>
        [CommandMethod("ExplodeNestedBlock")]
        public static void ExplodeNestedBlock(double clearthreshold = 20)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;
                using (doc.LockDocument())
                {
                    using var tr = new DBTrans(); // 开启事务管理器

                    // 提示用户选择块参照
                    var peo = new PromptEntityOptions("\n请选择要分解的块: ");
                    peo.SetRejectMessage("\n选择的不是块参照，请重新选择！");
                    peo.AddAllowedClass(typeof(BlockReference), true);

                    var per = Env.Editor.GetEntity(peo);
                    if (per.Status != PromptStatus.OK)
                    {
                        Env.Editor.WriteMessage("\n未选择任何块参照或操作已取消。");
                        return;
                    }

                    // 获取选中的块参照
                    var blockRef = tr.GetObject(per.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (blockRef == null)
                    {
                        Env.Editor.WriteMessage("\n选择的对象不是有效的块参照！");
                        return;
                    }

                    if (!VariableDictionary.winForm_Status)
                    {
                        // 获取清理参数阈值
                        clearthreshold = GetCleanupThreshold();
                    }

                    // 记录原始信息
                    var originalLayer = blockRef.Layer;
                    var originalPosition = blockRef.Position;
                    var originalBlockName = blockRef.Name;
                    // 默认使用原图层
                    string targetLayer = originalLayer;

                    // 弹出消息框：是否设备块
                    var dialogResult = System.Windows.Forms.MessageBox.Show(
                        "要分解的块是不是设备块:",
                        "分解确认",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Question);

                    if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                    {
                        // 仅当当前图层是“设备”或“设备名称”时，改为 Devices
                        if (string.Equals(originalLayer, "设备", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(originalLayer, "设备名称", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(originalLayer, "SB", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(originalLayer, "S_设备", StringComparison.OrdinalIgnoreCase))
                        {
                            targetLayer = "S_设备";
                            LayerDictionaryHelper.EnsureTargetLayer(tr, targetLayer, 140);
                            Env.Editor.WriteMessage($"\n已识别设备块，图层由 '{originalLayer}' 调整为 '{targetLayer}'。");
                        }
                    }
                    else
                    {
                        // 选“否”不处理
                        Env.Editor.WriteMessage("\n非设备块，保持原图层不变。");
                        targetLayer = originalLayer;
                    }

                    Env.Editor.WriteMessage($"\n开始处理块：{originalBlockName}，图层：{originalLayer}，清理阈值：{clearthreshold}");
                    AutoCadHelper.LogWithSafety($"开始处理块：{originalBlockName}，图层：{originalLayer}，位置：{originalPosition}，清理阈值：{clearthreshold}");

                    // 一、执行分解操作（带过滤）
                    var explodedEntities = RecursivelyExplodeBlocks(tr, blockRef, targetLayer, clearthreshold);

                    Env.Editor.WriteMessage($"\n分解完成，获得 {explodedEntities.Count} 个图元");
                    AutoCadHelper.LogWithSafety($"分解完成，获得 {explodedEntities.Count} 个图元");

                    tr.Commit(); // 提交事务
                }

            }
            catch (System.Exception ex)
            {
                AutoCadHelper.LogWithSafety($"分解嵌套块时出错: {ex.Message}");
                Env.Editor.WriteMessage($"\n错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 递归分解嵌套块，收集所有分解后的图元（优化版本 - 过滤小图元）
        /// 这版重点变化只有一处：**不再手工拼 `BlockTransform`，改为原生 `Explode` 递归**，镜像块分解错位问题通常会直接消失。
        /// </summary>
        /// <param name="tr">数据库事务</param>
        /// <param name="blockRef">要分解的块参照</param>
        /// <param name="targetLayer">目标图层名</param>
        /// <param name="clearthreshold">清理阈值</param>
        /// <returns>分解后得到的所有实体ID列表</returns>
        private static List<ObjectId> RecursivelyExplodeBlocks(DBTrans tr, BlockReference blockRef, string targetLayer = null, double clearthreshold = 20.0)
        {
            var resultEntities = new List<ObjectId>();
            var worldExplodedEntities = new List<Entity>(); // 收集最终图元（已是世界坐标下）
            var blocksToProcess = new Queue<(BlockReference br, bool isTransient)>(); // isTransient=true 表示临时对象，需要手动Dispose

            int totalProcessed = 0; // 处理块计数（含根块和嵌套块）
            int blockCount = 0;     // 嵌套块计数
            int entityCount = 0;    // 基础图元计数
            int filteredCount = 0;  // 过滤掉的图元数量

            // 根块是数据库对象，不应在本方法中Dispose
            blocksToProcess.Enqueue((blockRef, false));

            string layerForName = !string.IsNullOrEmpty(targetLayer) ? targetLayer : blockRef.Layer;

            // 确保目标图层存在
            if (!string.IsNullOrEmpty(targetLayer))
            {
                LayerDictionaryHelper.EnsureTargetLayer(tr, targetLayer, 140);
            }

            try
            {
                // 核心：使用AutoCAD原生 Explode 递归分解，避免手工矩阵链导致镜像/位置错误
                while (blocksToProcess.Count > 0)
                {
                    var (currentBr, isTransient) = blocksToProcess.Dequeue();
                    totalProcessed++;

                    try
                    {
                        var exploded = new DBObjectCollection();
                        currentBr.Explode(exploded); // 原生分解：自动处理旋转/镜像/缩放/位移

                        foreach (DBObject obj in exploded)
                        {
                            if (obj is BlockReference nestedBr)
                            {
                                blockCount++;
                                // 嵌套块继续递归分解（此对象为临时对象）
                                blocksToProcess.Enqueue((nestedBr, true));
                                continue;
                            }

                            if (obj is Entity ent)
                            {
                                // 过滤逻辑保持不变
                                if (ShouldIncludeEntity(ent, clearthreshold))
                                {
                                    ApplyExplodedEntityStyle(ent, targetLayer); // 图层/颜色规则保持不变
                                    worldExplodedEntities.Add(ent);
                                    entityCount++;

                                    if (entityCount % 5000 == 0)
                                    {
                                        Env.Editor.WriteMessage($"\r已处理 {entityCount} 个实体...");
                                    }
                                }
                                else
                                {
                                    filteredCount++;
                                    ent.Dispose(); // 被过滤的临时对象及时释放
                                }
                            }
                            else
                            {
                                obj.Dispose();
                            }
                        }
                    }
                    finally
                    {
                        // 仅释放临时块对象；根块为数据库对象不释放
                        if (isTransient && currentBr != null && !currentBr.IsDisposed)
                        {
                            currentBr.Dispose();
                        }
                    }
                }

                // 后续“创建新块+插入+删除临时图元”流程保持原样
                if (worldExplodedEntities.Count > 0)
                {
                    Point3d basePoint = GetRightMostPoint(worldExplodedEntities);

                    string timeStr = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string baseBlockName = $"{layerForName}_{timeStr}";
                    string newBlockName = baseBlockName;
                    int dupeCounter = 1;

                    var trBlockTable = tr.BlockTable;
                    while (trBlockTable.Has(newBlockName))
                    {
                        newBlockName = $"{baseBlockName}_{dupeCounter++}";
                    }

                    var newBtr = new BlockTableRecord { Name = newBlockName };
                    trBlockTable.Add(newBtr);

                    var explodedIdsToErase = new List<ObjectId>();

                    foreach (var ent in worldExplodedEntities)
                    {
                        // 先写入当前空间（用于保持你原来的流程）
                        var id = tr.CurrentSpace.AddEntity(ent);
                        explodedIdsToErase.Add(id);

                        // 再克隆一份写入新块定义
                        var blockEnt = ent.Clone() as Entity;
                        if (blockEnt != null)
                        {
                            blockEnt.TransformBy(Matrix3d.Displacement(basePoint.GetVectorTo(Point3d.Origin)));
                            newBtr.AddEntity(blockEnt);
                        }
                    }

                    var newBlockRef = new BlockReference(basePoint, newBtr.Id) { Layer = layerForName };
                    tr.CurrentSpace.AddEntity(newBlockRef);

                    // 删除临时写入的分解图元
                    foreach (var id in explodedIdsToErase)
                    {
                        var e = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                        if (e != null && !e.IsErased && !e.IsDisposed)
                        {
                            e.Erase();
                        }
                    }

                    // 删除原始块参照
                    if (!blockRef.IsDisposed)
                    {
                        blockRef.UpgradeOpen();
                        blockRef.Erase();
                    }

                    resultEntities.Clear();
                    resultEntities.Add(newBlockRef.Id);

                    Env.Editor.WriteMessage($"\n成功创建新块: {newBlockName}，基点(最右点): ({basePoint.X:0.###},{basePoint.Y:0.###},{basePoint.Z:0.###})");
                }
                else
                {
                    Env.Editor.WriteMessage("\n未提取到有效图元。");
                }

                Env.Editor.WriteMessage($"\n总共处理了 {totalProcessed} 个块，其中 {blockCount} 个嵌套块，{entityCount} 个基础图元，过滤掉 {filteredCount} 个小图元");
                AutoCadHelper.LogWithSafety($"处理完成：{totalProcessed} 个块，{blockCount} 个嵌套块，{entityCount} 个基础图元，过滤掉 {filteredCount} 个小图元");

                return resultEntities;
            }
            catch (Exception ex)
            {
                // 异常时释放未入库临时实体
                foreach (var ent in worldExplodedEntities)
                {
                    if (!ent.IsDisposed && ent.Database == null)
                    {
                        ent.Dispose();
                    }
                }

                AutoCadHelper.LogWithSafety($"分解嵌套块时出错: {ex.Message}");
                return resultEntities;
            }
        }

        /// <summary>
        /// 获取一组实体中最右边的点（如果有多个点X坐标相同，则取Y坐标最大的那个）
        /// </summary>
        /// <param name="entities"></param>
        /// <returns></returns>
        private static Point3d GetRightMostPoint(List<Entity> entities)
        {
            bool found = false;
            Point3d rightMost = Point3d.Origin;
            const double eps = 1e-6;

            foreach (var ent in entities)
            {
                try
                {
                    var ext = ent.GeometricExtents;
                    var p = ext.MaxPoint;

                    if (!found ||
                        p.X > rightMost.X + eps ||
                        (Math.Abs(p.X - rightMost.X) <= eps && p.Y > rightMost.Y))
                    {
                        rightMost = p;
                        found = true;
                    }
                }
                catch
                {
                    // 某些对象可能没有有效范围，忽略
                }
            }

            return found ? rightMost : Point3d.Origin;
        }


        /// <summary>
        /// 获取清理参数阈值
        /// </summary>
        /// <returns>清理参数阈值，默认为7</returns>
        private static double GetCleanupThreshold()
        {
            try
            {
                // 尝试从VariableDictionary获取清理参数值
                // 需要确保VariableDictionary类中已定义cleanupParameter字段
                var cleanupParam = VariableDictionary.cleanupParameter;

                // 如果值有效则使用，否则使用默认值
                if (cleanupParam > 0)
                {
                    return cleanupParam;
                }
            }
            catch
            {
                // 如果获取失败，使用默认值
            }

            // 默认值为7
            return 7.0;
        }

        /// <summary>
        /// 应用分解后图元的图层与颜色规则：
        /// 1) 图层统一到 targetLayer（如果有）
        /// 2) 非 Hatch 图元颜色改为 ByLayer
        /// 3) Hatch 保留原颜色
        /// </summary>
        private static void ApplyExplodedEntityStyle(Entity entity, string? targetLayer)
        {
            if (entity == null) return;

            if (!string.IsNullOrEmpty(targetLayer))
            {
                entity.Layer = targetLayer;
            }

            // 填充保持原颜色，不修改
            if (entity is Hatch)
            {
                return;
            }

            // 非填充图元统一改为随层颜色
            entity.ColorIndex = 256; // ByLayer
        }

        #endregion


        #region   通过拿到外参中的图元id后台打开外参中的图元再插入到当前文档中

        /// <summary>
        /// 命令入口：复制外参图元（统一走核心逻辑）
        /// </summary>
        [CommandMethod("CopyXrefEntity")]
        public void CopyXrefEntity()
        {
            // 调用统一核心方法，避免与 SELECTXREFENTITY 逻辑分叉
            CopyXrefEntityCore("\n请点击外部参照中的图元: ");
        }

        /// <summary>
        /// 外参图元复制核心方法（仅单图元复制，不做整图兜底）
        /// 说明：
        /// 1) 支持 eInvalidOwner/eInvalidOwnerObject 自动重选一次；
        /// 2) 保留天正代理对象提示；
        /// 3) 严格禁止整图复制兜底（按你的要求）。
        /// </summary>
        /// <param name="firstPrompt">首次提示文本</param>
        private void CopyXrefEntityCore(string firstPrompt)
        {
            // 获取当前活动文档
            Document doc = Application.DocumentManager.MdiActiveDocument;
            // 判空保护
            if (doc == null) return;

            // 获取编辑器
            Editor ed = doc.Editor;

            try
            {
                // 记录开始日志
                LogManager.Instance.LogInfo("\n开始复制外部参照图元（仅单图元模式）...");

                // 锁定文档，避免并发写冲突
                using (doc.LockDocument())
                {
                    // 最多尝试两次：首次 + 失败后重选一次
                    const int maxAttempts = 2;

                    // 重试循环
                    for (int attempt = 1; attempt <= maxAttempts; attempt++)
                    {
                        try
                        {
                            // 每次尝试独立事务，避免异常污染下一次
                            using (var tr = new DBTrans())
                            {
                                // 第二次给重选提示
                                string prompt = attempt == 1
                                    ? (string.IsNullOrWhiteSpace(firstPrompt) ? "\n请点击外部参照中的图元: " : firstPrompt)
                                    : "\n请重选外部参照中的块/主实体: ";

                                // 创建嵌套选择选项
                                PromptNestedEntityOptions options = new PromptNestedEntityOptions(prompt);
                                // 不允许空选
                                options.AllowNone = false;

                                // 获取嵌套选择结果
                                PromptNestedEntityResult per = Env.Editor.GetNestedEntity(options);

                                // 用户取消则直接退出
                                if (per.Status != PromptStatus.OK)
                                {
                                    LogManager.Instance.LogInfo("\n用户取消了外部参照图元选择。");
                                    return;
                                }

                                // 打开选中对象
                                Entity selectedEnt = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;

                                // 非实体直接退出
                                if (selectedEnt == null)
                                {
                                    LogManager.Instance.LogInfo("\n选中的对象不是有效实体，操作取消。");
                                    return;
                                }

                                // 判断是否疑似天正代理/子对象
                                string dxfName = selectedEnt.ObjectId.ObjectClass?.DxfName ?? string.Empty;
                                bool isTzProxy =
                                    selectedEnt is ProxyEntity ||
                                    (!string.IsNullOrWhiteSpace(dxfName) && dxfName.StartsWith("TCH", StringComparison.OrdinalIgnoreCase)) ||
                                    (selectedEnt.GetType().FullName?.IndexOf("Proxy", StringComparison.OrdinalIgnoreCase) >= 0);

                                // 给出更明确提示（不中断流程）
                                if (isTzProxy)
                                {
                                    ed.WriteMessage("\n提示：当前对象疑似天正代理/子对象，建议优先选择块参照本体或主实体。");
                                }

                                // 尝试单图元克隆到当前空间
                                ObjectId clonedId;
                                bool cloned = TryCloneEntityToCurrentSpace(tr, per.ObjectId, out clonedId);

                                // 克隆成功则应用嵌套变换，保持原位
                                if (cloned && !clonedId.IsNull)
                                {
                                    // 打开克隆对象用于写入变换
                                    Entity clonedEnt = tr.GetObject(clonedId, OpenMode.ForWrite) as Entity;
                                    if (clonedEnt != null)
                                    {
                                        // 应用选择时返回的嵌套变换矩阵
                                        clonedEnt.TransformBy(per.Transform);
                                    }

                                    // 提交事务并刷新
                                    tr.Commit();
                                    Env.Editor.Redraw();

                                    // 成功日志
                                    LogManager.Instance.LogInfo($"\n已成功复制图元到当前图纸（原位置）：{selectedEnt.GetType().Name}");
                                    return;
                                }

                                // 单图元复制失败：不做整图兜底（按需求）
                                if (isTzProxy)
                                {
                                    ed.WriteMessage("\n复制失败：天正代理对象无法直接单体复制，请改选块参照本体。");
                                }
                                else
                                {
                                    ed.WriteMessage("\n复制失败：无法克隆该图元，请改选外参中的主实体/块参照。");
                                }

                                // 记录失败日志
                                LogManager.Instance.LogInfo("\n复制失败：仅允许单图元复制，未启用整图兜底。");
                                return;
                            }
                        }
                        // 捕获 AutoCAD 异常，处理不可直接复制子对象场景
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            // 兼容不同版本枚举名
                            string statusName = ex.ErrorStatus.ToString();
                            bool isInvalidOwner =
                                string.Equals(statusName, "eInvalidOwnerObject", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(statusName, "eInvalidOwner", StringComparison.OrdinalIgnoreCase);

                            // 典型子对象复制失败
                            if (isInvalidOwner)
                            {
                                // 错误日志
                                LogManager.Instance.LogError($"复制外部参照图元失败（{statusName}）: {ex.Message}");
                                LogManager.Instance.LogError($"错误堆栈: {ex.StackTrace}");

                                // 还有重试机会就继续
                                if (attempt < maxAttempts)
                                {
                                    ed.WriteMessage("\n当前选中的是不可直接复制子对象，请改选块/主实体，系统将自动重选一次。");
                                    continue;
                                }

                                // 重试后仍失败
                                ed.WriteMessage("\n重选后仍失败：请优先选择块参照本体。");
                                return;
                            }

                            // 非预期 AutoCAD 异常
                            ed.WriteMessage($"\n错误: {ex.Message}");
                            LogManager.Instance.LogError($"复制外部参照图元时出错: {ex.Message}");
                            LogManager.Instance.LogError($"错误堆栈: {ex.StackTrace}");
                            return;
                        }
                    }
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                // 外层 AutoCAD 异常
                ed.WriteMessage($"\n错误: {ex.Message}");
                LogManager.Instance.LogError($"复制外部参照图元时出错: {ex.Message}");
                LogManager.Instance.LogError($"错误堆栈: {ex.StackTrace}");
            }
            catch (Exception ex)
            {
                // 外层通用异常
                LogManager.Instance.LogError($"复制外部参照图元时出错: {ex.Message}");
                LogManager.Instance.LogError($"错误堆栈: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 将源对象（可能来自当前库或外参库）克隆到当前空间
        /// 优先尝试标准克隆（DeepClone/WblockClone），遇到 eInvalidOwnerObject 时回退为“单实体手动克隆”
        /// 说明：本方法只做“单图元复制”，不会触发整图复制兜底
        /// </summary>
        /// <param name="tr">当前事务</param>
        /// <param name="sourceId">源对象Id</param>
        /// <param name="clonedId">返回克隆后的目标对象Id</param>
        /// <returns>是否克隆成功</returns>
        private static bool TryCloneEntityToCurrentSpace(DBTrans tr, ObjectId sourceId, out ObjectId clonedId)
        {
            // 默认返回空Id，表示失败
            clonedId = ObjectId.Null;

            // 基础保护：事务为空直接失败
            if (tr == null) return false;

            // 基础保护：源Id为空直接失败
            if (sourceId.IsNull) return false;

            // 本地函数：判断是否为“Owner无效”异常（兼容不同版本枚举名）
            static bool IsInvalidOwnerError(Autodesk.AutoCAD.Runtime.Exception ex)
            {
                string status = ex.ErrorStatus.ToString();
                return string.Equals(status, "eInvalidOwnerObject", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(status, "eInvalidOwner", StringComparison.OrdinalIgnoreCase);
            }

            // 本地函数：同库手动克隆回退（仅单实体，不跨库）
            bool TryManualCloneSameDb(out ObjectId newId)
            {
                // 初始化输出
                newId = ObjectId.Null;

                try
                {
                    // 读取源实体（同库可直接用当前事务读取）
                    var srcEnt = tr.GetObject(sourceId, OpenMode.ForRead) as Entity;
                    if (srcEnt == null) return false;

                    // 克隆实体副本
                    var clone = srcEnt.Clone() as Entity;
                    if (clone == null) return false;

                    // 加入当前空间
                    newId = tr.CurrentSpace.AddEntity(clone);
                    return !newId.IsNull;
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogInfo($"\nTryManualCloneSameDb 失败: {ex.Message}");
                    newId = ObjectId.Null;
                    return false;
                }
            }

            // 本地函数：跨库手动克隆回退（仅单实体，不整图）
            bool TryManualCloneCrossDb(Database sourceDb, out ObjectId newId)
            {
                // 初始化输出
                newId = ObjectId.Null;

                try
                {
                    // 在源库开启只读事务，读取源实体
                    using (var srcTr = sourceDb.TransactionManager.StartTransaction())
                    {
                        // 从源库读取实体
                        var srcEnt = srcTr.GetObject(sourceId, OpenMode.ForRead) as Entity;
                        if (srcEnt == null) return false;

                        // 克隆实体副本
                        var clone = srcEnt.Clone() as Entity;
                        if (clone == null) return false;

                        // 将克隆副本加入当前空间
                        newId = tr.CurrentSpace.AddEntity(clone);

                        // 源事务提交（规范收尾）
                        srcTr.Commit();
                    }

                    // 返回是否成功
                    return !newId.IsNull;
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogInfo($"\nTryManualCloneCrossDb 失败: {ex.Message}");
                    newId = ObjectId.Null;
                    return false;
                }
            }

            try
            {
                // 读取源数据库与目标数据库
                Database sourceDb = sourceId.Database;
                Database destDb = tr.Database;

                // 判空保护
                if (sourceDb == null || destDb == null) return false;

                // 情况A：同库对象
                if (ReferenceEquals(sourceDb, destDb))
                {
                    try
                    {
                        // 先确认是实体，避免非实体/子对象异常
                        var srcEnt = tr.GetObject(sourceId, OpenMode.ForRead) as Entity;
                        if (srcEnt == null) return false;

                        // 标准路径：DeepClone 到当前空间
                        var ids = new ObjectIdCollection { sourceId };
                        var idMap = new IdMapping();
                        destDb.DeepCloneObjects(ids, tr.CurrentSpace.ObjectId, idMap, false);

                        // 读取映射结果
                        if (!idMap.Contains(sourceId)) return false;
                        clonedId = idMap[sourceId].Value;
                        return !clonedId.IsNull;
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex) when (IsInvalidOwnerError(ex))
                    {
                        // 命中 owner 无效：回退到同库手动克隆
                        LogManager.Instance.LogInfo($"\nDeepCloneObjects 命中 {ex.ErrorStatus}，改走同库手动克隆回退。");

                        // 用本地输出接收，避免直接在本地函数里写外层 out 参数
                        if (TryManualCloneSameDb(out ObjectId fallbackId))
                        {
                            clonedId = fallbackId;
                            return true;
                        }

                        return false;
                    }
                }
                // 情况B：跨库对象（典型：外参库）
                try
                {
                    // 标准路径：WblockClone 单对象到当前空间
                    var xrefIds = new ObjectIdCollection { sourceId };
                    var xrefMap = new IdMapping();
                    sourceDb.WblockCloneObjects(
                        xrefIds,
                        tr.CurrentSpace.ObjectId,
                        xrefMap,
                        DuplicateRecordCloning.Ignore,
                        false);

                    // 读取映射结果
                    if (!xrefMap.Contains(sourceId)) return false;
                    clonedId = xrefMap[sourceId].Value;
                    return !clonedId.IsNull;
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex) when (IsInvalidOwnerError(ex))
                {
                    // 命中 owner 无效：回退到跨库手动克隆（仍然只克隆单实体）
                    LogManager.Instance.LogInfo($"\nWblockCloneObjects 命中 {ex.ErrorStatus}，改走跨库手动克隆回退。");

                    // 用本地输出接收，避免直接在本地函数里写外层 out 参数
                    if (TryManualCloneCrossDb(sourceDb, out ObjectId fallbackId))
                    {
                        clonedId = fallbackId;
                        return true;
                    }
                    return false;
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                // AutoCAD 异常兜底
                LogManager.Instance.LogInfo($"\nTryCloneEntityToCurrentSpace 克隆失败（AutoCAD异常）: {ex.Message}");
                clonedId = ObjectId.Null;
                return false;
            }
            catch (Exception ex)
            {
                // 通用异常兜底
                LogManager.Instance.LogInfo($"\nTryCloneEntityToCurrentSpace 克隆失败（通用异常）: {ex.Message}");
                clonedId = ObjectId.Null;
                return false;
            }
        }


        /// <summary>
        /// 复制外部参照中的所有图元
        /// </summary>
        //[CommandMethod("CopyXrefAllEntity")]
        //public void CopyXrefAllEntity()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    //Database db = doc.Database;
        //    Editor ed = doc.Editor;
        //    using (var tr = new DBTrans())
        //    {
        //        try
        //        {
        //            #region 拿到外参图元信息
        //            // 第一步：创建嵌套实体选择选项
        //            PromptNestedEntityOptions options = new PromptNestedEntityOptions("\n请点击外部参照中的图元: ");
        //            // 允许用户选择任意层级的嵌套实体（外部参照中的实体）
        //            options.AllowNone = false;
        //            // 第二步：获取用户选择的嵌套实体
        //            PromptNestedEntityResult per = Env.Editor.GetNestedEntity(options);
        //            if (per.Status != PromptStatus.OK) return;
        //            // 2. 获取选中的图元对象
        //            Entity selectedEnt = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
        //            #endregion
        //            #region 获取外参路径与块名
        //            // 获取嵌套容器（父块参照链）
        //            ObjectId[] containers = per.GetContainers();
        //            if (containers.Length == 0)
        //            {
        //                LogManager.Instance.LogInfo("\n未找到父块参照。");
        //                return;
        //            }
        //            // 一般我们取最后一个或倒数第二个
        //            ObjectId parentBlockRefId = containers.Last(); // 最外层块
        //            BlockReference parentBlockRef = tr.GetObject(parentBlockRefId, OpenMode.ForRead) as BlockReference;
        //            if (parentBlockRef == null)
        //            {
        //                LogManager.Instance.LogInfo("\n父块参照无效。");
        //                return;
        //            }
        //            // 获取父块参照（文件）的名称
        //            string parentBlockName = parentBlockRef.Name;
        //            LogManager.Instance.LogInfo($"\n父块参照名称: {parentBlockName}");
        //            // 获取父级块表记录
        //            BlockTableRecord btr = tr.GetObject(parentBlockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
        //            if (btr == null)
        //            {
        //                LogManager.Instance.LogInfo("\n块表记录无效。");
        //                return;
        //            }
        //            // 判断是否为外部参照
        //            if (btr.IsFromExternalReference)
        //            {
        //                // 获取外部参照信息
        //                var xrefDb = btr.GetXrefDatabase(true);
        //                if (xrefDb != null)
        //                {
        //                    // 获取外部参照文件路径
        //                    LogManager.Instance.LogInfo($"\n外部参照文件路径: {xrefDb.Filename}");
        //                }
        //                else
        //                {
        //                    LogManager.Instance.LogInfo("\n无法获取外部参照数据库。");
        //                }
        //            }
        //            else
        //            {
        //                LogManager.Instance.LogInfo("\n该块不是外部参照。");
        //            }
        //            #endregion
        //            // 获取外部参照路径信息
        //            string xrefPath = btr.PathName;
        //            string fileName = System.IO.Path.GetFileName(xrefPath);
        //            // 获取块名称
        //            var blockName = selectedEnt.BlockName;
        //            // 去除 |前缀
        //            blockName = blockName.Split('|').Last();
        //            if (GB_XrefInsertAllBlock(xrefPath, blockName))
        //            {
        //                LogManager.Instance.LogInfo("\n已复制外部参照：");
        //            }
        //            else
        //            {
        //                LogManager.Instance.LogInfo("\n复制外部参照失败：");
        //            }
        //            tr.Commit();
        //        }
        //        catch (Autodesk.AutoCAD.Runtime.Exception ex)
        //        {
        //            LogManager.Instance.LogInfo($"\n错误: {ex.Message}");
        //        }
        //    }
        //}

        /// <summary>
        /// 从指定DWG文件中复制符合ClassID和边界条件的实体到当前图纸
        /// </summary>
        /// <param name="sourceFilePath">源DWG文件路径</param>
        /// <param name="classId">目标实体的ClassID（用于筛选特定类型对象）</param>
        /// <param name="bounds">目标实体的边界范围（格式如"0,0,0;100,100,0"，用于筛选位置）</param>
        //public static bool CopyEntityByClassIdFromDwg(string sourceFilePath, string classId, string bounds)
        //{
        //    // 获取当前AutoCAD应用程序的当前文档和数据库
        //    var curDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        //    var curDb = curDoc.Database; // 当前图纸的数据库（用于写入复制后的实体）

        //    // 创建源DWG文件的数据库对象（不自动打开事务，不保留事务日志）
        //    using (var sourceDb = new Autodesk.AutoCAD.DatabaseServices.Database(false, true))
        //    {
        //        /* 步骤1：读取源DWG文件到临时数据库 */
        //        // 从指定路径读取DWG文件到源数据库（FileShare.ReadWrite允许其他进程读写，true表示加载外部参照）
        //        sourceDb.ReadDwgFile(sourceFilePath, System.IO.FileShare.ReadWrite, true, null);

        //        /* 步骤2：启动源数据库的事务，用于访问其数据 */
        //        using (var sourceTr = sourceDb.TransactionManager.StartTransaction())
        //        {
        //            // 获取源数据库的块表（存储所有块定义，如模型空间、图纸空间）
        //            var sourceBT = (BlockTable)sourceTr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
        //            // 获取模型空间的块表记录（模型空间是默认的绘图区域）
        //            var sourceBTR = (BlockTableRecord)sourceTr.GetObject(sourceBT[BlockTableRecord.ModelSpace], OpenMode.ForRead);

        //            /* 步骤3：遍历模型空间的所有实体，查找符合条件的目标实体 */
        //            Entity? foundEntity = null; // 存储找到的目标实体（可空类型）
        //            foreach (ObjectId entityObjId in sourceBTR) // 遍历模型空间中的每个实体ID
        //            {
        //                // 从事务中获取实体对象（只读模式）
        //                var sourceEntity = sourceTr.GetObject(entityObjId, OpenMode.ForRead) as Entity;
        //                // 筛选条件：实体存在、ClassID匹配、边界范围匹配
        //                if (sourceEntity != null
        //                    && sourceEntity.Bounds.ToString() == bounds)   // 边界范围匹配（筛选位置）
        //                {
        //                    foundEntity = sourceEntity; // 找到符合条件的实体，保存到变量
        //                    break; // 找到后退出循环
        //                }
        //            }

        //            // 检查是否找到目标实体
        //            if (foundEntity == null)
        //            {
        //                // 未找到时弹出提示对话框
        //                //Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("未找到图元。");
        //                LogManager.Instance.LogInfo("\n未找到天正图元。");
        //                return false; // 结束方法
        //            }

        //            /* 步骤4：通过临时数据库中转，复制实体到当前图纸 */
        //            // 创建临时数据库（用于中转克隆对象，避免直接操作源数据库或当前数据库）
        //            using (var tempDb = new Autodesk.AutoCAD.DatabaseServices.Database(true, true))
        //            {
        //                // 准备要克隆的对象ID集合（仅包含找到的目标实体）
        //                ObjectIdCollection ids = new ObjectIdCollection { foundEntity.ObjectId };
        //                // ID映射表（记录源对象ID到目标对象ID的映射关系，用于处理重复ID）
        //                IdMapping mapping = new IdMapping();

        //                /* 子步骤4.1：将目标实体从源数据库克隆到临时数据库的当前空间 */
        //                // WblockCloneObjects：将对象从源数据库克隆到目标数据库
        //                // 参数说明：源对象ID集合、目标空间ID、ID映射表、重复记录处理方式（替换）、是否保留颜色/线型
        //                sourceDb.WblockCloneObjects(ids, tempDb.CurrentSpaceId, mapping, DuplicateRecordCloning.Replace, false);

        //                /* 子步骤4.2：将临时数据库中的实体克隆到当前图纸的当前空间 */
        //                // 启动当前数据库的事务（用于写入克隆后的实体）
        //                using (var curTrans = curDb.TransactionManager.StartTransaction())
        //                {
        //                    // 获取当前图纸的模型空间块表记录（可写模式）
        //                    var curMs = (BlockTableRecord)curTrans.GetObject(
        //                        curDb.CurrentSpaceId, OpenMode.ForWrite);

        //                    // 准备临时数据库中需要克隆的实体ID集合（从临时数据库的当前空间获取）
        //                    ObjectIdCollection tmpIds = new ObjectIdCollection();
        //                    // 启动临时数据库的事务（用于读取克隆后的实体）
        //                    using (var tmpTrans = tempDb.TransactionManager.StartTransaction())
        //                    {
        //                        // 获取临时数据库的当前空间块表记录（只读模式）
        //                        var tmpMs = (BlockTableRecord)tmpTrans.GetObject(
        //                            tempDb.CurrentSpaceId, OpenMode.ForRead);
        //                        // 遍历临时空间中的所有实体ID，添加到tmpIds集合
        //                        foreach (ObjectId id in tmpMs)
        //                        {
        //                            tmpIds.Add(id);
        //                        }
        //                        tmpTrans.Commit(); // 提交临时数据库事务（释放资源）
        //                    }

        //                    // ID映射表（记录临时数据库对象ID到当前数据库对象ID的映射）
        //                    IdMapping curMapping = new IdMapping();
        //                    // 将临时数据库中的实体克隆到当前图纸的当前空间
        //                    // 参数说明：临时对象ID集合、当前空间ID、ID映射表、重复记录处理方式（替换）、是否保留颜色/线型
        //                    tempDb.WblockCloneObjects(tmpIds, curMs.ObjectId, curMapping, DuplicateRecordCloning.Replace, false);

        //                    curTrans.Commit(); // 提交当前数据库事务（保存克隆的实体）
        //                }
        //            }
        //            sourceTr.Commit(); // 提交源数据库事务（释放资源）
        //        }
        //    }
        //    return true;
        //}

        #endregion

        /// <summary>
        /// 复制外参1Line2Polyline
        /// </summary>
        //[CommandMethod("CopyAndSync1")]
        //public void CopyAndSync1()
        //{
        //    try
        //    {

        //        //选择的外部参照
        //        PromptSelectionResult getselection = Env.Editor.GetSelection(new PromptSelectionOptions() { MessageForAdding = "请选择待处理图形:\n" });
        //        if (getselection.Status == PromptStatus.OK)
        //        {
        //            using (var tr = new DBTrans())
        //            {
        //                //ojectid集合
        //                List<ObjectId> needids = new List<ObjectId>();
        //                //当前文件的块表
        //                BlockTable bt = (BlockTable)tr.GetObject(Env.Database.BlockTableId, OpenMode.ForWrite);
        //                //循环选择的外参中的每个元素的objectid
        //                foreach (ObjectId oneid in getselection.Value.GetObjectIds())
        //                {
        //                    //每个元素的dbobject
        //                    DBObject getdbo = tr.GetObject(oneid, OpenMode.ForWrite);
        //                    //如果这个元素是参照块
        //                    if (getdbo is BlockReference)
        //                    {
        //                        ObjectId newid = ObjectId.Null;
        //                        //判断是不是动态块
        //                        if ((getdbo as BlockReference).IsDynamicBlock)
        //                            //newid赋值动态块表记录
        //                            newid = (getdbo as BlockReference).DynamicBlockTableRecord;
        //                        else
        //                        {
        //                            //newid赋值匿名块表记录
        //                            newid = (getdbo as BlockReference).AnonymousBlockTableRecord;

        //                        }//newid是不是为空

        //                        if (newid.IsNull)
        //                            //newid赋值为块表记录
        //                            newid = (getdbo as BlockReference).BlockTableRecord;
        //                        else if (!newid.IsNull)
        //                        {
        //                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(newid, OpenMode.ForWrite);
        //                            //检查btr块表记录是不是外部参照
        //                            if (btr.IsFromExternalReference)
        //                                needids.Add(newid);
        //                        }
        //                    }
        //                }
        //                //把外参绑定进当前图里
        //                Env.Database.BindXrefs(new ObjectIdCollection(needids.ToArray()), false);

        //                foreach (ObjectId oneid in needids)
        //                {
        //                    //块表记录btr
        //                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(oneid, OpenMode.ForWrite);
        //                    //获取块表记录里的所有块引用
        //                    ObjectIdCollection findids = btr.GetBlockReferenceIds(true, false);
        //                    //获取块表记录里的所有匿名块引用
        //                    ObjectIdCollection findids1 = btr.GetAnonymousBlockIds();
        //                    foreach (ObjectId newid in findids)
        //                    {
        //                        //获取块引用
        //                        BlockReference newblk = (BlockReference)tr.GetObject(newid, OpenMode.ForWrite);
        //                        // 炸开这个块引用
        //                        newblk.ExplodeToOwnerSpace();

        //                    }
        //                    foreach (ObjectId newid in findids1)
        //                    {
        //                        BlockReference newblk = (BlockReference)tr.GetObject(newid, OpenMode.ForWrite);
        //                        newblk.ExplodeToOwnerSpace();
        //                    }
        //                    //btr.Erase();
        //                }
        //                tr.Commit();
        //                Env.Editor.Redraw();
        //            }
        //        }
        //        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("处理完成!");
        //    }
        //    catch (Exception ex)
        //    {
        //        // 记录错误日志  
        //        LogManager.Instance.LogInfo("选择图元复制到当前空间内失败！");
        //        LogManager.Instance.LogInfo($"\n错误: {ex.Message}");
        //    }
        //    try
        //    {
        //        var iFoxTr = new DBTrans();
        //        // 提示用户选择外部参照的图
        //        PromptEntityOptions entityOptions = new PromptEntityOptions("请选择外部参照的图");
        //        entityOptions.SetRejectMessage("请选择一个外部参照的图");
        //        entityOptions.AddAllowedClass(typeof(BlockReference), true);//设定选定的文件为外部参照块；
        //        PromptEntityResult entityResult = iFoxTr.Editor.GetEntity(entityOptions);//获取外部引用文件的实体；
        //        if (entityResult.Status != PromptStatus.OK) return;

        //        #region 处理文件新
        //        //Document thisdoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        //        List<ObjectId> needids = new List<ObjectId>();
        //        do
        //        {
        //            #region
        //            using (var tr = new DBTrans())
        //            {
        //                ////获取当前文档的块表
        //                BlockTable bt = (BlockTable)tr.GetObject(Env.Database.BlockTableId, OpenMode.ForRead);

        //                //循环块表
        //                foreach (ObjectId oneid in bt)
        //                {
        //                    //获取块表记录
        //                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(oneid, OpenMode.ForRead);
        //                    //检查btr块表记录是不是外部参照
        //                    if (btr.IsFromExternalReference)
        //                        needids.Add(oneid);
        //                }
        //                tr.Commit();
        //            }
        //            ///去重
        //            needids = needids.Distinct().ToList();
        //            #endregion
        //            #region
        //            if (needids.Count > 0)
        //            {
        //                using (var tr = new DBTrans())
        //                {
        //                    //获取当前文档的块表
        //                    BlockTable bt = (BlockTable)tr.GetObject(Env.Database.BlockTableId, OpenMode.ForRead);
        //                    //绑定外参
        //                    Env.Database.BindXrefs(new ObjectIdCollection(needids.ToArray()), false);
        //                    //循环块表
        //                    foreach (ObjectId oneid in needids)
        //                    {
        //                        ///获取块表记录
        //                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(oneid, OpenMode.ForRead);
        //                        ///获取块表记录里的所有块引用
        //                        ObjectIdCollection findids = btr.GetBlockReferenceIds(true, false);
        //                        ///获取块表记录里的所有匿名块引用
        //                        ObjectIdCollection findids1 = btr.GetAnonymousBlockIds();
        //                        ///循环块引用
        //                        foreach (ObjectId newid in findids)
        //                        {
        //                            ///获取块引用
        //                            BlockReference newblk = (BlockReference)tr.GetObject(newid, OpenMode.ForWrite);
        //                            ///炸开这个块引用
        //                            newblk.ExplodeToOwnerSpace();
        //                        }
        //                        foreach (ObjectId newid in findids1)
        //                        {
        //                            BlockReference newblk = (BlockReference)tr.GetObject(newid, OpenMode.ForWrite);
        //                            newblk.ExplodeToOwnerSpace();
        //                        }
        //                        ///删除块表记录
        //                        btr.Erase();
        //                    }

        //                    tr.Commit();
        //                }
        //                //
        //                needids.Clear();
        //            }
        //            #endregion

        //            #region
        //            using (Transaction tr = Env.Database.TransactionManager.StartTransaction())
        //            {
        //                BlockTable bt = (BlockTable)tr.GetObject(Env.Database.BlockTableId, OpenMode.ForWrite);
        //                #region
        //                foreach (ObjectId oneid in bt)
        //                {
        //                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(oneid, OpenMode.ForWrite);
        //                    if (btr.IsFromExternalReference)
        //                        needids.Add(oneid);
        //                }
        //                #endregion

        //                tr.Commit();
        //            }
        //            #endregion

        //        } while (needids.Count > 0);

        //        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("处理完成!");
        //        #endregion
        //        iFoxTr.Commit();
        //        Env.Editor.Redraw();
        //    }
        //    catch (Exception ex)
        //    {
        //        // 记录错误日志  
        //        LogManager.Instance.LogInfo("选择图元复制到当前空间内失败！");
        //    }
        //}

        /// <summary>
        /// 外参复制
        /// </summary>
        //[CommandMethod("ReferenceCopy")]
        //public void ReferenceCopy()
        //{
        //    try
        //    {
        //        using var tr = new DBTrans();
        //        foreach (var objectId in FormMain.selectedEntities)
        //        {
        //            //Entity entity = Command.GetEntity(objectId);
        //            Entity entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
        //            if (entity.Handle.ToString() == FormMain.selectItem)
        //                tr.CurrentSpace.AddEntity(entity);
        //        }
        //        tr.Commit();
        //        Env.Editor.Redraw();
        //    }
        //    catch (Exception ex)
        //    {
        //        LogManager.Instance.LogInfo(ex.Message);
        //    }
        //}

        /// <summary>
        /// 判断选择图元与当前空间中的图无是不相同，不同时复制到当前空间内
        /// </summary>
        //[CommandMethod("CompareAndReplace")]
        //public void CompareAndReplace()
        //{
        //    try
        //    {
        //        // 选择复制进来的图元
        //        PromptSelectionResult selectionResult = Env.Editor.GetSelection();
        //        if (selectionResult.Status != PromptStatus.OK) return;

        //        using (var tr = new DBTrans())
        //        {
        //            // 获取当前图中的所有图元
        //            List<Entity> currentEntities = new List<Entity>();
        //            foreach (ObjectId bTOBId in tr.BlockTable)
        //            {
        //                Entity? entity = tr.GetObject(bTOBId, OpenMode.ForRead) as Entity;
        //                if (entity != null)
        //                    currentEntities.Add(entity);
        //            }
        //            // 获取选择集中的图元
        //            SelectionSet selectionSet = selectionResult.Value;
        //            foreach (SelectedObject selectedObject in selectionSet)
        //            {
        //                Entity? copiedEntity = tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead) as Entity;
        //                // 判断复制进来的图元是否与当前图中的任何一个图元相同
        //                bool isSame = false;
        //                if (copiedEntity != null)
        //                    foreach (Entity currentEntity in currentEntities)
        //                    {
        //                        if (IsSameEntity(copiedEntity, currentEntity))
        //                        {
        //                            isSame = true;
        //                            break;
        //                        }
        //                    }
        //                if (!isSame)
        //                {
        //                    if (copiedEntity != null)
        //                    {
        //                        // 复制进来的图元与当前图中的图元不同，将其添加到当前图中
        //                        Entity? copy = copiedEntity.Clone() as Entity;
        //                        if (copy != null)
        //                            tr.CurrentSpace.AddEntity(copy);
        //                    }
        //                }
        //            }
        //            // 提交事务
        //            tr.Commit();
        //            Env.Editor.Redraw();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        // 记录错误日志  
        //        LogManager.Instance.LogInfo("选择图元复制到当前空间内失败！");
        //        LogManager.Instance.LogInfo(ex.Message);
        //    }
        //}

        /// <summary>
        /// 判断两个实体是否相同
        /// </summary>
        /// <param name="entity1">实体一</param>
        /// <param name="entity2">实体二</param>
        /// <returns></returns>
        //private bool IsSameEntity(Entity entity1, Entity entity2)
        //{

        //    // 比较图层
        //    if (entity1.Layer != entity2.Layer)
        //        return false;

        //    // 比较位置
        //    if (!entity1.GeometricExtents.MinPoint.IsEqualTo(entity2.GeometricExtents.MinPoint, Tolerance.Global))
        //        return false;

        //    // 比较大小
        //    if (!entity1.GeometricExtents.MaxPoint.IsEqualTo(entity2.GeometricExtents.MaxPoint, Tolerance.Global))
        //        return false;

        //    // 比较颜色
        //    if (entity1.Color != entity2.Color)
        //        return false;

        //    // 其他属性比较...

        //    return true;
        //}

        #endregion




        #region 结构画方、园、多边形
        /// <summary>
        /// 结构-用户指定原点后半径画圆；
        /// </summary>
        [CommandMethod(nameof(CirRadius))]
        public static void CirRadius()
        {
            try
            {
                double width = 0;
                //获取图层名称
                string? layerName = VariableDictionary.layerName;//图层名
                var layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);//图层颜色
                double cirPlus = Convert.ToDouble(VariableDictionary.textbox_CirPlus_Text) * 2;//拿到圆的扩展值；
                if (VariableDictionary.textbox_CirPlus_Text == null) cirPlus = 0;//如果没有搌值，那就是0；
                using var tr = new DBTrans();//开启事务
                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；
                var userPoint1 = Env.Editor.GetPoint("\n请指定圆形洞口的起点");//指定圆的第一点
                if (userPoint1.Status != PromptStatus.OK) return;
                var UcsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();//把指定的点转成Wcs2坐标；
                // 创建polyline
                Polyline polylineHatch = new Polyline();
                Point3d center = new Point3d(0, 0, 0);//圆心
                //拖动实现
                using var cir = new JigEx((mpw, queue) =>
                {
                    var userPoint2 = mpw.Z20();//mpw为鼠标移动变量；
                    var userCir = new Circle(UcsUserPoint1, Vector3d.ZAxis, (userPoint2.DistanceTo(UcsUserPoint1)) + cirPlus);//创建半径动态圆；
                    userCir.Layer = layerName;//动态圆设置图层；
                    userCir.ColorIndex = layerColorIndex;
                    center = userCir.Center;//圆心
                    double radius = userCir.Radius;//圆半径
                    width = radius * 2;
                    Point3d polyline1Start = new Point3d(center.X - radius * Math.Cos(Math.PI / 4), center.Y - radius * Math.Sin(Math.PI / 4), 0);
                    Point3d polylineCenter = new Point3d(center.X - radius * Math.Cos(Math.PI / 4) / 2, center.Y + radius * Math.Sin(Math.PI / 4) / 2, 0);
                    Point3d polyline2End = new Point3d(center.X + radius * Math.Cos(Math.PI / 4), center.Y + radius * Math.Sin(Math.PI / 4), 0);
                    Point3d pointOnArc = new Point3d(center.X - radius * Math.Cos(Math.PI / 4), center.Y - radius * Math.Sin(Math.PI / 4), 0);

                    var polyline1 = new Polyline();
                    double bulge = CalculateBulge(polyline1Start, pointOnArc, polyline2End, center, radius);
                    polyline1.AddVertexAt(0, new Point2d(polyline1Start.X, polyline1Start.Y), -bulge, 0, 0);
                    polyline1.AddVertexAt(1, new Point2d(polyline2End.X, polyline2End.Y), 0, 0, 0);
                    polyline1.AddVertexAt(2, new Point2d(polylineCenter.X, polylineCenter.Y), 0, 0, 0);
                    polyline1.AddVertexAt(3, new Point2d(polyline1Start.X, polyline1Start.Y), 0, 0, 0);
                    //polyline1.Closed = true;
                    polyline1.Layer = layerName;
                    polyline1.ColorIndex = layerColorIndex;
                    queue.Enqueue(polyline1);
                    queue.Enqueue(userCir);
                    polylineHatch = polyline1;
                });
                cir.SetOptions(UcsUserPoint1, msg: "\n请指定圆形洞口的终点");//提示用户输入第二点；
                var r2 = Env.Editor.Drag(cir);//拿到第二点
                if (r2.Status != PromptStatus.OK) tr.Abort();
                tr.CurrentSpace.AddEntity(cir.Entities);//把圆的实体写入当前的空间；
                if (layerName != null && layerName.Contains("结构"))
                    //调用填充
                    autoHatch(tr, layerName, layerColorIndex, 50, "DOTS", polylineHatch.ObjectId);
                Env.Editor.Redraw();
                DDimLinear(tr, (width).ToString("0" + "mm"), (0).ToString("0" + "mm"), Convert.ToInt16(3), center);
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n结构-用户指定两点为直径画圆失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 结构-用户指定两点为半径画圆；
        /// </summary>
        [CommandMethod(nameof(CirRadius_2))]
        public static void CirRadius_2()
        {
            try
            {
                string? layerName = VariableDictionary.layerName;
                var layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);//图层颜色
                double cirPlus = Convert.ToDouble(VariableDictionary.textbox_CirPlus_Text);//拿到圆的扩展值；
                double cirRedius = 1;//圆的半径
                if (VariableDictionary.textbox_S_Cirradius != null) cirRedius = VariableDictionary.textbox_S_Cirradius.Value;//圆的半径
                if (VariableDictionary.textbox_CirPlus_Text == null) cirPlus = 0;//如果没有搌值，那就是0；
                using var tr = new DBTrans();//开启事务             

                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；
                var cirCenter = new Point3d(0, 0, 0);
                var polylineHatch = new Polyline();
                Point3d center = new Point3d(0, 0, 0);//圆的中心
                //拖动圆
                using var cir = new JigEx((mpw, queue) =>
                {
                    var userPoint2 = mpw.Z20();//mpw为鼠标移动变量；
                    var userCir = new Circle(mpw, Vector3d.ZAxis, (cirRedius) + cirPlus);//创建指定直径的圆；
                    userCir.Layer = layerName;//圆设置图层；
                    userCir.ColorIndex = layerColorIndex;
                    center = userCir.Center;//圆的中心
                    double radius = userCir.Radius;//圆的半径
                    // 计算两个polyline的点
                    Point3d polyline1Start = new Point3d(center.X - radius * Math.Cos(Math.PI / 4), center.Y - radius * Math.Sin(Math.PI / 4), 0);
                    Point3d polylineCenter = new Point3d(center.X - radius * Math.Cos(Math.PI / 4) / 2, center.Y + radius * Math.Sin(Math.PI / 4) / 2, 0);
                    Point3d polyline2End = new Point3d(center.X + radius * Math.Cos(Math.PI / 4), center.Y + radius * Math.Sin(Math.PI / 4), 0);
                    Point3d pointOnArc = new Point3d(center.X - radius * Math.Cos(Math.PI / 4), center.Y - radius * Math.Sin(Math.PI / 4), 0);

                    var polyline1 = new Polyline();
                    double bulge = CalculateBulge(polyline1Start, pointOnArc, polyline2End, center, radius);
                    polyline1.AddVertexAt(0, new Point2d(polyline1Start.X, polyline1Start.Y), -bulge, 0, 0);
                    polyline1.AddVertexAt(1, new Point2d(polyline2End.X, polyline2End.Y), 0, 0, 0);
                    polyline1.AddVertexAt(2, new Point2d(polylineCenter.X, polylineCenter.Y), 0, 0, 0);
                    polyline1.AddVertexAt(3, new Point2d(polyline1Start.X, polyline1Start.Y), 0, 0, 0);
                    //polyline1.Closed = true;
                    polyline1.Layer = layerName;
                    polyline1.ColorIndex = layerColorIndex;
                    queue.Enqueue(polyline1);
                    queue.Enqueue(userCir);
                    polylineHatch = polyline1;
                });
                cir.SetOptions(msg: "\n请指定圆形洞口的终点");//提示用户输入第二点；
                var r2 = Env.Editor.Drag(cir);//拿到第二点
                if (r2.Status != PromptStatus.OK) tr.Abort();
                tr.CurrentSpace.AddEntity(cir.Entities);//把圆的实体写入当前的空间；
                if (layerName != null && layerName.Contains("结构"))
                    //调用填充
                    autoHatch(tr, layerName, layerColorIndex, 50, "DOTS", polylineHatch.ObjectId);
                Env.Editor.Redraw();
                DDimLinear(tr, (cirRedius * 2 + cirPlus * 2).ToString("0" + "mm"), (0).ToString("0" + "mm"), Convert.ToInt16(3), center);
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n结构-用户指定两点为直径画圆失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 结构-用户指定数值为直径画圆；
        /// </summary>
        [CommandMethod(nameof(CirDiameter))]
        public static void CirDiameter()
        {
            try
            {
                double width = 0;
                string? layerName = VariableDictionary.btnBlockLayer;
                var layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);//图层颜色
                double cirPlus = Convert.ToDouble(VariableDictionary.textbox_CirPlus_Text);//拿到圆的扩展值；
                if (VariableDictionary.textbox_CirPlus_Text == null) cirPlus = 0;//如果没有搌值，那就是0；
                using var tr = new DBTrans();//开启事务

                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；
                var userPoint1 = Env.Editor.GetPoint("\n请指定圆形洞口的起点");//指定圆的第一点
                if (userPoint1.Status != PromptStatus.OK) return; //指定成功
                var UcsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();//把指定的点转成Wcs2坐标；
                // 声明一个变量保存填充边界的多段线（用于 SOLID 填充）  
                Polyline polylineHatch = new Polyline();
                Point3d center = new Point3d(0, 0, 0);
                //拖动实现
                using var cir = new JigEx((mpw, queue) =>
                {
                    var userPoint2 = mpw.Z20();//mpw为鼠标移动变量；
                    var userCir = new Circle(UcsUserPoint1.GetMidPointTo(pt2: userPoint2), Vector3d.ZAxis, (userPoint2.DistanceTo(UcsUserPoint1) / 2) + cirPlus);//创建动态圆；
                    userCir.Layer = layerName;//动态圆设置图层；
                    userCir.ColorIndex = layerColorIndex;
                    queue.Enqueue(userCir);//动态的圆，跟随鼠标变化；
                    center = userCir.Center;
                    double radius = userCir.Radius;
                    width = radius * 2;
                    // 计算两个polyline的点
                    Point3d polyline1Start = new Point3d(center.X - radius * Math.Cos(Math.PI / 4), center.Y - radius * Math.Sin(Math.PI / 4), 0);
                    Point3d polylineCenter = new Point3d(center.X - radius * Math.Cos(Math.PI / 4) / 2, center.Y + radius * Math.Sin(Math.PI / 4) / 2, 0);
                    Point3d polyline2End = new Point3d(center.X + radius * Math.Cos(Math.PI / 4), center.Y + radius * Math.Sin(Math.PI / 4), 0);
                    Point3d pointOnArc = new Point3d(center.X - radius * Math.Cos(Math.PI / 4), center.Y - radius * Math.Sin(Math.PI / 4), 0);
                    // 创建第一个polyline
                    var polyline1 = new Polyline();
                    double bulge = CalculateBulge(polyline1Start, pointOnArc, polyline2End, center, radius);
                    polyline1.AddVertexAt(0, new Point2d(polyline1Start.X, polyline1Start.Y), -bulge, 0, 0);
                    polyline1.AddVertexAt(1, new Point2d(polyline2End.X, polyline2End.Y), 0, 0, 0);
                    polyline1.AddVertexAt(2, new Point2d(polylineCenter.X, polylineCenter.Y), 0, 0, 0);
                    polyline1.AddVertexAt(3, new Point2d(polyline1Start.X, polyline1Start.Y), 0, 0, 0);
                    queue.Enqueue(polyline1);
                    polyline1.Layer = layerName;
                    polyline1.ColorIndex = layerColorIndex;
                    polylineHatch = polyline1;
                });
                cir.SetOptions(UcsUserPoint1, msg: "\n请指定圆形洞口的终点");//提示用户输入第二点；
                var r2 = Env.Editor.Drag(cir);//拿到第二点
                if (r2.Status != PromptStatus.OK) tr.Abort();
                var cirEntity = tr.CurrentSpace.AddEntity(cir.Entities);//把圆的实体写入当前的空间；
                if (layerName != null && layerName.Contains("结构"))
                    //调用填充
                    autoHatch(tr, layerName, layerColorIndex, 50, "DOTS", polylineHatch.ObjectId);
                Env.Editor.Redraw();
                DDimLinear(tr, (width).ToString("0" + "mm"), (0).ToString("0" + "mm"), Convert.ToInt16(3), center);
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n结构-用户指定两点为直径画圆失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 结构-用户指定两点为直径画圆；
        /// </summary>
        [CommandMethod(nameof(CirDiameter_2))]
        public static void CirDiameter_2()
        {
            try
            {
                string? layerName = VariableDictionary.btnBlockLayer;
                var layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                double cirPlus = Convert.ToDouble(VariableDictionary.textbox_CirPlus_Text) * 2;//拿到圆的扩展值；
                double cirDiameter = 1;
                if (VariableDictionary.textBox_S_CirDiameter != null) cirDiameter = VariableDictionary.textBox_S_CirDiameter.Value;
                if (VariableDictionary.textbox_CirPlus_Text == null) cirPlus = 0;//如果没有搌值，那就是0；

                using var tr = new DBTrans();//开启事务

                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；
                var cirCenter = new Point3d(0, 0, 0);
                // 创建第一个polyline
                Polyline polylineHatch = new Polyline();
                Point3d center = new Point3d(0, 0, 0);//圆的中心
                //拖动圆
                using var cir = new JigEx((mpw, queue) =>
                {
                    var userCir = new Circle(mpw, Vector3d.ZAxis, (cirDiameter / 2) + cirPlus / 2);//创建指定直径的圆；
                    userCir.Layer = layerName;//圆设置图层；
                    userCir.ColorIndex = layerColorIndex;
                    center = userCir.Center;//圆的中心
                    double radius = userCir.Radius;//圆的半径
                    // 计算两个polyline的点
                    var polyline1Start = new Point3d(center.X - radius * Math.Cos(Math.PI / 4), center.Y - radius * Math.Sin(Math.PI / 4), 0);
                    var polylineCenter = new Point3d(center.X - radius * Math.Cos(Math.PI / 4) / 2, center.Y + radius * Math.Sin(Math.PI / 4) / 2, 0);
                    var polyline2End = new Point3d(center.X + radius * Math.Cos(Math.PI / 4), center.Y + radius * Math.Sin(Math.PI / 4), 0);
                    //拿到在开始点到结束点中间的一个点
                    var pointOnArc = new Point3d(center.X - radius * Math.Cos(Math.PI / 4), center.Y + radius * Math.Sin(Math.PI / 4), 0);

                    var polyline1 = new Polyline();
                    double bulge = CalculateBulge(polyline1Start, pointOnArc, polyline2End, center, radius);
                    polyline1.AddVertexAt(0, new Point2d(polyline1Start.X, polyline1Start.Y), bulge, 0, 0);
                    polyline1.AddVertexAt(1, new Point2d(polyline2End.X, polyline2End.Y), 0, 0, 0);
                    polyline1.AddVertexAt(2, new Point2d(polylineCenter.X, polylineCenter.Y), 0, 0, 0);
                    polyline1.AddVertexAt(3, new Point2d(polyline1Start.X, polyline1Start.Y), 0, 0, 0);
                    //polyline1.Closed = true;
                    polyline1.Layer = layerName;
                    polyline1.ColorIndex = layerColorIndex;
                    queue.Enqueue(polyline1);
                    queue.Enqueue(userCir);
                    polylineHatch = polyline1;
                });
                cir.SetOptions(msg: "\n请指定圆形洞口的终点");//提示用户输入第二点；
                var r2 = Env.Editor.Drag(cir);//拿到第二点
                if (r2.Status != PromptStatus.OK) tr.Abort();//如果不是ok，那就撤消
                tr.CurrentSpace.AddEntity(cir.Entities);//把圆的实体写入当前的空间；
                if (layerName != null && layerName.Contains("结构"))
                    //调用填充
                    autoHatch(tr, layerName, layerColorIndex, 50, "DOTS", polylineHatch.ObjectId);
                Env.Editor.Redraw();

                DDimLinear(tr, (cirDiameter + cirPlus).ToString("0" + "mm"), (0).ToString("0" + "mm"), Convert.ToInt16(3), center);

                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n结构-用户指定两点为直径画圆失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 计算bulge
        /// </summary>
        /// <param name="polyline1Start">pl线开始点</param>
        /// <param name="pointOnArc">圆弧上的点</param>
        /// <param name="polyline2End">pl线结束点</param>
        /// <param name="center">中心点</param>
        /// <param name="radius">半径</param>
        /// <returns></returns>
        private static double CalculateBulge(Point3d polyline1Start, Point3d pointOnArc, Point3d polyline2End, Point3d center, double radius)
        {
            //求得圆心到开始点的向量
            var cs = center.GetVectorTo(polyline1Start);
            //求得圆心到结束点的向量
            var ce = center.GetVectorTo(polyline2End);
            //X方向向量
            Vector3d xvector = new Vector3d(1, 0, 0);
            //绘制一个临时圆弧
            CircularArc3d cirArc = new CircularArc3d(polyline1Start, pointOnArc, polyline2End);
            ////计算圆弧的开如角度
            double startAngle = cs.Y > 0 ? xvector.GetAngleTo(cs) : -xvector.GetAngleTo(cs);
            ////计算圆弧的结束角度
            double endAngle = ce.Y > 0 ? xvector.GetAngleTo(ce) : -xvector.GetAngleTo(ce);
            //绘制一个圆弧

            Arc arc = new Arc(center, radius, startAngle, endAngle);

            // 计算圆弧的夹角
            //double angle = v1.GetAngleTo(v2);
            // 计算 bulge 值（tan(角度/4)）
            double bulge = Math.Tan(startAngle);
            if (((pointOnArc.X - polyline1Start.X) * (polyline2End.Y - polyline1Start.Y) - (pointOnArc.Y - polyline1Start.Y) * (polyline2End.X - polyline1Start.X)) < 0)
                bulge = -bulge;
            return bulge;
        }

        /// <summary>
        /// 结构用户用鼠标画矩形
        /// </summary>
        /// <param name="layerName"></param>
        [CommandMethod(nameof(Rec2PolyLine))]
        public static void Rec2PolyLine()
        {
            try
            {
                double width = 0;
                double height = 0;
                // 图层名称和颜色
                var layerName = VariableDictionary.layerName != null ? VariableDictionary.layerName : VariableDictionary.btnBlockLayer;
                var layerColor = VariableDictionary.layerColorIndex;
                double recPlus = Convert.ToDouble(VariableDictionary.textbox_RecPlus_Text); // 指定的偏移量  
                using var tr = new DBTrans();
                // 检查并创建图层，返回图层ID
                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, Convert.ToInt16(layerColor));
                // 获取方形洞口的第一点  
                var userPoint1 = Env.Editor.GetPoint("\n请指定方形洞口第一点");
                if (userPoint1.Status != PromptStatus.OK) return;
                var UcsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();
                // 声明一个变量保存填充边界的多段线（用于 SOLID 填充）  
                Polyline hatchBoundary = new Polyline();
                var newRectPointMin = new Point3d(0, 0, 0);
                var newRectPointMax = new Point3d(0, 0, 0);

                // 使用 JigEx 动态预览，绘制矩形以及两条辅助线  
                using var rec = new JigEx((mpw, queue) =>
                {
                    var UcsUserPoint2 = mpw.Z20();
                    // 动态绘制原始矩形（从用户第一点到第二点）  
                    Polyline polylineRec = new Polyline();
                    polylineRec.AddVertexAt(0, new Point2d(UcsUserPoint1.X, UcsUserPoint1.Y), 0, 0, 0);
                    polylineRec.AddVertexAt(1, new Point2d(UcsUserPoint2.X, UcsUserPoint1.Y), 0, 0, 0);
                    polylineRec.AddVertexAt(2, new Point2d(UcsUserPoint2.X, UcsUserPoint2.Y), 0, 0, 0);
                    polylineRec.AddVertexAt(3, new Point2d(UcsUserPoint1.X, UcsUserPoint2.Y), 0, 0, 0);
                    polylineRec.Closed = true; // 闭合成方形  
                    Extents3d polyLineRecExt = new Extents3d();
                    // 获取原始矩形的边界  
                    if (polylineRec.Bounds != null) polyLineRecExt = (Extents3d)polylineRec.Bounds;
                    // 计算扩大后的矩形边界（偏移 recPlus）  
                    Extents3d newRectBounds = new Extents3d(
                        new Point3d(polyLineRecExt.MinPoint.X - recPlus, polyLineRecExt.MinPoint.Y - recPlus, 0),
                        new Point3d(polyLineRecExt.MaxPoint.X + recPlus, polyLineRecExt.MaxPoint.Y + recPlus, 0));
                    width = newRectBounds.MaxPoint.X - newRectBounds.MinPoint.X;//拿到宽
                    height = newRectBounds.MaxPoint.Y - newRectBounds.MinPoint.Y;//拿到高
                    // 辅助计算：取扩大矩形的左上与右下点，计算二者间 1/4 点（作为两条线交合点）  
                    var leftUp = new Point3d(newRectBounds.MinPoint.X, newRectBounds.MaxPoint.Y, 0);
                    var rightDown = new Point3d(newRectBounds.MaxPoint.X, newRectBounds.MinPoint.Y, 0);
                    double targetX = leftUp.X + (rightDown.X - leftUp.X) * 1.0 / 4;
                    double targetY = leftUp.Y + (rightDown.Y - leftUp.Y) * 1.0 / 4;
                    double targetZ = 0; // Z 为 0  
                    Point3d targetPoint = new Point3d(targetX, targetY, targetZ);

                    // 创建扩大后的矩形  
                    Polyline newRect = new Polyline();
                    newRect.AddVertexAt(0, new Point2d(newRectBounds.MinPoint.X, newRectBounds.MinPoint.Y), 0, 3, 3);
                    newRect.AddVertexAt(1, new Point2d(newRectBounds.MinPoint.X, newRectBounds.MaxPoint.Y), 0, 3, 3);
                    newRect.AddVertexAt(2, new Point2d(newRectBounds.MaxPoint.X, newRectBounds.MaxPoint.Y), 0, 3, 3);
                    newRect.AddVertexAt(3, new Point2d(newRectBounds.MaxPoint.X, newRectBounds.MinPoint.Y), 0, 3, 3);
                    newRect.Closed = true;
                    newRect.Layer = layerName;
                    newRect.ColorIndex = layerColor;
                    newRectPointMin = new Point3d(newRectBounds.MinPoint.X, newRectBounds.MinPoint.Y, 0);
                    newRectPointMax = new Point3d(newRectBounds.MaxPoint.X, newRectBounds.MaxPoint.Y, 0);
                    queue.Enqueue(newRect); // 放大后的矩形  
                    // 绘制用来辅助生成填充边界的两条线（或辅助多段线）  
                    // 此处构造的填充边界区域：由扩大矩形的左下角、左上角、右上角，以及计算得到的交合点构成  
                    if (layerName.Contains("结构"))
                    {
                        hatchBoundary = new Polyline();
                        hatchBoundary.AddVertexAt(0, new Point2d(newRectBounds.MinPoint.X, newRectBounds.MinPoint.Y), 0, 3, 3);
                        hatchBoundary.AddVertexAt(1, new Point2d(newRectBounds.MinPoint.X, newRectBounds.MaxPoint.Y), 0, 3, 3);
                        hatchBoundary.AddVertexAt(2, new Point2d(newRectBounds.MaxPoint.X, newRectBounds.MaxPoint.Y), 0, 3, 3);
                        hatchBoundary.AddVertexAt(3, new Point2d(targetPoint.X, targetPoint.Y), 0, 3, 3);
                        hatchBoundary.Closed = true;
                        hatchBoundary.Layer = layerName;
                        hatchBoundary.ColorIndex = layerColor;
                        queue.Enqueue(hatchBoundary);
                    }
                });
                rec.SetOptions(UcsUserPoint1, msg: "\n请指定方形洞口第二点");
                var userPoint2 = Env.Editor.Drag(rec);
                if (userPoint2.Status != PromptStatus.OK)
                    tr.Abort();
                // 将动态绘制得到的所有实体添加到当前空间  
                tr.CurrentSpace.AddEntity(rec.Entities);
                Env.Editor.Redraw();
                if (layerName.Contains("结构"))
                    //调用填充
                    autoHatch(tr, layerName, layerColor, 50, "DOTS", hatchBoundary.ObjectId);
                Env.Editor.Redraw();
                //计算洞口的中心点
                var centerPoint = new Point3d(((newRectPointMin.X + newRectPointMax.X) / 2), (newRectPointMin.Y + newRectPointMax.Y) / 2, 0);
                var dimTextColor = Convert.ToInt16(VariableDictionary.textColorIndex);//标注文字颜色，绿色
                DDimLinear(tr, (width).ToString("0") + "mm", (height).ToString("0") + "mm", 3, centerPoint);
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n结构指定数值生成矩形失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 暖通用户指定两点为对角线画方Rec2PolyLine_N
        /// </summary>
        /// <param name="layerName"></param>
        //[CommandMethod(nameof(Rec2PolyLine_N))]
        //public static void Rec2PolyLine_N()
        //{
        //    try
        //    {
        //        // 获取当前编辑器
        //        string? layerName = VariableDictionary.layerName;
        //        Point3d centerPoint;//初始化起始点
        //        var textScale = VariableDictionary.textBoxScale;//获取标注比例
        //        if (layerName == null)
        //            return;
        //        double recLRPlus = Convert.ToDouble(VariableDictionary.textbox_Width);// 指定左右偏移量
        //        double recUDPlus = Convert.ToDouble(VariableDictionary.textbox_Height);// 指定左右偏移量
        //        short layerColor = Convert.ToInt16(VariableDictionary.layerColorIndex);//设置图层颜色
        //        using var tr = new DBTrans();
        //        // 检查图层是否存在，如果不存在则创建图层
        //        string targetLayer = EnsureTargetLayer(tr, VariableDictionary.btnBlockLayer ?? VariableDictionary.layerName, layerColor);
        //        ////查找图层，没有即创建；
        //        //if (!tr.LayerTable.Has(layerName))
        //        //    tr.LayerTable.Add(layerName, ltt =>
        //        //    {
        //        //        ltt.Color = Color.FromColorIndex(ColorMethod.ByColor, Convert.ToInt16(VariableDictionary.layerColorIndex));
        //        //        ltt.LineWeight = LineWeight.LineWeight030;
        //        //        ltt.IsPlottable = true;
        //        //    });
        //        //// 先记住原始的全局线型比例
        //        //double oldLtscale = Convert.ToDouble(Application.GetSystemVariable("LTSCALE"));
        //        //Application.SetSystemVariable("LTSCALE", oldLtscale);  // 或直接设 1.0

        //        //查找线型，没有即创建；
        //        if (!tr.LinetypeTable.Has("DASH"))
        //        {
        //            tr.LinetypeTable.Add("DASH", ltr =>
        //            {
        //                //设置线型的图层

        //                ltr.Name = "DASH";
        //                ltr.AsciiDescription = " - - - - - ";//线型描述
        //                ltr.PatternLength = 1; //线型总长度
        //                ltr.NumDashes = 2;//组成线型的笔画数目
        //                ltr.SetDashLengthAt(0, 0.6);//0.5个单位的画线
        //                ltr.SetDashLengthAt(1, -0.4);//0.3个单位的空格
        //                ltr.SetShapeStyleAt(1, tr.TextStyleTable["tJText"]);//设置文字的样式
        //                //ltr.SetShapeNumberAt(1, 0);//设置空格处包含的图案图形
        //                ltr.SetShapeOffsetAt(1, new Vector2d(-0.1, -0.05));//图形在X轴方向上左移0.1个单位，在Y轴方向上下移0.05个单位
        //                ltr.SetShapeScaleAt(1, 0.1 * textScale);//图形的缩放比例
        //                ltr.SetShapeIsUcsOrientedAt(1, false);
        //                Application.SetSystemVariable("LTSCALE", 100.0);
        //                //ltr.SetShapeRotationAt(1, 0);//ltr.SetTextAt(1, "测绘");//文字内容
        //                //ltr.SetDashLengthAt(2, -0.2);//0.2个单位的空格
        //            });
        //        }
        //        var userPoint1 = Env.Editor.GetPoint("\n请指定方形洞口第一点");
        //        if (userPoint1.Status != PromptStatus.OK) return;
        //        var UcsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();
        //        // 计算扩大后的矩形坐标
        //        var targetPoint = new Point3d(0, 0, 0);
        //        using var rec = new JigEx((mpw, queue) =>
        //        {
        //            var UcsUserPoint2 = mpw.Z20();//鼠标所在的动态位置；
        //            Polyline polylineRec = new Polyline();
        //            polylineRec.AddVertexAt(0, new Point2d(UcsUserPoint1.X, UcsUserPoint1.Y), 0, 0, 0);
        //            polylineRec.AddVertexAt(1, new Point2d(UcsUserPoint2.X, UcsUserPoint1.Y), 0, 0, 0);
        //            polylineRec.AddVertexAt(2, new Point2d(UcsUserPoint2.X, UcsUserPoint2.Y), 0, 0, 0);
        //            polylineRec.AddVertexAt(3, new Point2d(UcsUserPoint1.X, UcsUserPoint2.Y), 0, 0, 0);
        //            polylineRec.Closed = true;//闭合成方形

        //            centerPoint = new Point3d((UcsUserPoint1.X + UcsUserPoint2.X) / 2, (UcsUserPoint1.Y + UcsUserPoint2.Y) / 2, 0);
        //            Extents3d polyLineRecExt = new Extents3d();
        //            // 获取原始矩形的边界  
        //            if (polylineRec.Bounds != null) polyLineRecExt = (Extents3d)polylineRec.Bounds;
        //            Extents3d newRectBounds = new Extents3d();
        //            if (recLRPlus > 0)
        //            {
        //                // 计算扩大矩形的边界框
        //                newRectBounds = new Extents3d(
        //                    new Point3d(polyLineRecExt.MinPoint.X - recLRPlus, polyLineRecExt.MinPoint.Y - recUDPlus, 0),
        //                    new Point3d(polyLineRecExt.MaxPoint.X + recLRPlus, polyLineRecExt.MaxPoint.Y + recUDPlus, 0));
        //            }
        //            else
        //            {
        //                // 计算扩大矩形的边界框
        //                newRectBounds = new Extents3d(
        //                    new Point3d(polyLineRecExt.MinPoint.X - recLRPlus, polyLineRecExt.MinPoint.Y - recUDPlus, 0),
        //                    new Point3d(polyLineRecExt.MaxPoint.X + recLRPlus, polyLineRecExt.MaxPoint.Y + recUDPlus, 0));
        //            }

        //            // 创建扩大矩形
        //            Polyline newRect = new Polyline();
        //            newRect.LinetypeId = tr.LinetypeTable["DASH"];//为矩形设置线型
        //            newRect.Layer = layerName;
        //            newRect.ColorIndex = VariableDictionary.layerColorIndex;
        //            newRect.AddVertexAt(0, new Point2d(newRectBounds.MinPoint.X, newRectBounds.MinPoint.Y), 0, 0, 0);
        //            newRect.AddVertexAt(1, new Point2d(newRectBounds.MinPoint.X, newRectBounds.MaxPoint.Y), 0, 0, 0);
        //            newRect.AddVertexAt(2, new Point2d(newRectBounds.MaxPoint.X, newRectBounds.MaxPoint.Y), 0, 0, 0);
        //            newRect.AddVertexAt(3, new Point2d(newRectBounds.MaxPoint.X, newRectBounds.MinPoint.Y), 0, 0, 0);
        //            newRect.Closed = true;
        //            newRect.ConstantWidth = 30;
        //            queue.Enqueue(newRect);//放大后的矩形
        //        });
        //        rec.SetOptions(UcsUserPoint1, msg: "\n请指定方形洞口第二点");
        //        var userPoint2 = Env.Editor.Drag(rec);
        //        if (userPoint2.Status != PromptStatus.OK) return;
        //        var polyLineEntityObj = tr.CurrentSpace.AddEntity(rec.Entities);
        //        Env.Editor.Redraw();
        //        //MousePointWcsLast为鼠标最后的点击点；
        //        var UcsUserPoint2 = rec.MousePointWcsLast;
        //        //计算洞口的中心点
        //        centerPoint = new Point3d((UcsUserPoint1.X + UcsUserPoint2.X) / 2, (UcsUserPoint1.Y + UcsUserPoint2.Y) / 2, 0);
        //        //调用标注的方法，给定第一点坐标注与图层名；
        //        DDimLinear(tr, layerName, Convert.ToInt16(VariableDictionary.layerColorIndex), centerPoint);
        //        //调用读取天正数据的方法
        //        tzData();
        //        Env.Editor.Redraw();
        //        //if (!VariableDictionary.btnBlockLayer.Contains("给排水"))
        //        //如果厚度为0时
        //        if (hvacR3 == "0")
        //        {
        //            int plusX = Convert.ToInt32(Math.Abs(UcsUserPoint1.X - UcsUserPoint2.X)) + Convert.ToInt32(VariableDictionary.textbox_Width) * 2;
        //            int plusY = Convert.ToInt32(Math.Abs(UcsUserPoint1.Y - UcsUserPoint2.Y)) + Convert.ToInt32(VariableDictionary.textbox_Height) * 2;
        //            PointDim(tr,centerPoint, "洞：" + plusX.ToString(), "x" + plusY.ToString(), " ", layerName, Convert.ToInt16(VariableDictionary.layerColorIndex));

        //        }
        //        else
        //        {
        //            int plusX = Convert.ToInt32(hvacR4) + Convert.ToInt32(VariableDictionary.textbox_Width) * 2;
        //            int plusY = Convert.ToInt32(hvacR3) + Convert.ToInt32(VariableDictionary.textbox_Height) * 2;
        //            PointDim(tr,centerPoint, "洞：" + plusX.ToString(), " x " + plusY.ToString(), "\n距地：" + strHvacStart, layerName, Convert.ToInt16(VariableDictionary.layerColorIndex));
        //        }

        //        tr.Commit();
        //        Env.Editor.Redraw();
        //    }
        //    catch (Exception ex)
        //    {
        //        // 记录错误日志  
        //        LogManager.Instance.LogInfo($"\n暖通用户指定两点为对角线画方失败！错误信息: {ex.Message}");
        //        LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}");

        //    }
        //}

        /// <summary>
        /// 暖通用户指定两点为对角线画方Rec2PolyLine_N
        /// </summary>
        /// <param name="layerName"></param>
        [CommandMethod(nameof(Rec2PolyLine_N))]
        public static void Rec2PolyLine_N()
        {
            try
            {
                Point3d centerPoint = Point3d.Origin;
                Point3d ucsUserPoint1 = Point3d.Origin;
                Point3d ucsUserPoint2 = Point3d.Origin;

                double textScale = AutoCadHelper.GetScale();//获取标注比例
                double recLRPlus = Convert.ToDouble(VariableDictionary.textbox_Width);// 指定左右偏加值量
                double recUDPlus = Convert.ToDouble(VariableDictionary.textbox_Height);// 指定上下偏加值量
                short layerColor = Convert.ToInt16(VariableDictionary.layerColorIndex);// 设置图层颜色

                using var tr = new DBTrans();
                // 检查图层是否存在，如果不存在则创建图层，并返回图层名称
                string targetLayer = LayerDictionaryHelper.EnsureTargetLayer(tr, VariableDictionary.btnBlockLayer ?? VariableDictionary.layerName, layerColor);

                // 确保文字样式存在（DASH 里会引用）              
                TextFontsStyleHelper.EnsureTextStyle(tr, "tJText");
                // 确保 DASH 线型，并统一线型比例/系统变量               
                ObjectId dashLinetypeId = LineTypeStyleHelper.EnsureDashLinetype(
                    tr,
                    uiScale: textScale,
                    textStyleName: "tJText",
                    syncSystemVariables: true,
                    ltscaleOverride: Math.Max(0.1, textScale),
                    celtscaleOverride: 1.0);

                var userPoint1 = Env.Editor.GetPoint("\n请指定方形洞口第一点");
                if (userPoint1.Status != PromptStatus.OK) return;

                ucsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();

                using var rec = new JigEx((mpw, queue) =>
                {
                    var p2 = mpw.Z20();

                    var polylineRec = new Polyline();
                    polylineRec.AddVertexAt(0, new Point2d(ucsUserPoint1.X, ucsUserPoint1.Y), 0, 0, 0);
                    polylineRec.AddVertexAt(1, new Point2d(p2.X, ucsUserPoint1.Y), 0, 0, 0);
                    polylineRec.AddVertexAt(2, new Point2d(p2.X, p2.Y), 0, 0, 0);
                    polylineRec.AddVertexAt(3, new Point2d(ucsUserPoint1.X, p2.Y), 0, 0, 0);
                    polylineRec.Closed = true;

                    centerPoint = new Point3d((ucsUserPoint1.X + p2.X) / 2, (ucsUserPoint1.Y + p2.Y) / 2, 0);

                    Extents3d ext = polylineRec.Bounds.HasValue ? polylineRec.Bounds.Value : new Extents3d();
                    Extents3d newRectBounds = new Extents3d(
                        new Point3d(ext.MinPoint.X - recLRPlus, ext.MinPoint.Y - recUDPlus, 0),
                        new Point3d(ext.MaxPoint.X + recLRPlus, ext.MaxPoint.Y + recUDPlus, 0));

                    var newRect = new Polyline();
                    newRect.LinetypeId = dashLinetypeId;
                    newRect.LinetypeScale = 3 / textScale;
                    newRect.Layer = targetLayer;
                    newRect.ColorIndex = layerColor;
                    newRect.AddVertexAt(0, new Point2d(newRectBounds.MinPoint.X, newRectBounds.MinPoint.Y), 0, 0, 0);
                    newRect.AddVertexAt(1, new Point2d(newRectBounds.MinPoint.X, newRectBounds.MaxPoint.Y), 0, 0, 0);
                    newRect.AddVertexAt(2, new Point2d(newRectBounds.MaxPoint.X, newRectBounds.MaxPoint.Y), 0, 0, 0);
                    newRect.AddVertexAt(3, new Point2d(newRectBounds.MaxPoint.X, newRectBounds.MinPoint.Y), 0, 0, 0);
                    newRect.Closed = true;
                    newRect.ConstantWidth = 0.3 * textScale;
                    queue.Enqueue(newRect);
                    Env.Editor.Redraw();
                });
                rec.SetOptions(ucsUserPoint1, msg: "\n请指定方形洞口第二点");
                var dragRes = rec.Drag();
                if (dragRes.Status != PromptStatus.OK)
                {
                    tr.Abort();
                    return;
                }

                tr.CurrentSpace.AddEntity(rec.Entities);

                ucsUserPoint2 = rec.MousePointWcsLast;
                centerPoint = new Point3d((ucsUserPoint1.X + ucsUserPoint2.X) / 2, (ucsUserPoint1.Y + ucsUserPoint2.Y) / 2, 0);

                Env.Editor.Redraw();
                // 下面方法内部各自开事务：放在外层事务之外
                //调用标注的方法，给定第一点坐标注与图层名；
                DDimLinear(tr, targetLayer, Convert.ToInt16(VariableDictionary.layerColorIndex), centerPoint);
                TianZhengHelper.tzData();

                if (TianZhengHelper.hvacR3 == "0")
                {
                    int plusX = Convert.ToInt32(Math.Abs(ucsUserPoint1.X - ucsUserPoint2.X)) + Convert.ToInt32(VariableDictionary.textbox_Width) * 2;
                    int plusY = Convert.ToInt32(Math.Abs(ucsUserPoint1.Y - ucsUserPoint2.Y)) + Convert.ToInt32(VariableDictionary.textbox_Height) * 2;
                    PointDim(tr, centerPoint, "洞：" + plusX, "x" + plusY, "\n距地：", VariableDictionary.btnBlockLayer ?? VariableDictionary.layerName ?? "0", layerColor);
                }
                else
                {
                    int plusX = Convert.ToInt32(TianZhengHelper.hvacR4) + Convert.ToInt32(VariableDictionary.textbox_Width) * 2;
                    int plusY = Convert.ToInt32(TianZhengHelper.hvacR3) + Convert.ToInt32(VariableDictionary.textbox_Height) * 2;
                    PointDim(tr, centerPoint, "洞：" + plusX, "x" + plusY, "\n距地：" + TianZhengHelper.strHvacStart, VariableDictionary.btnBlockLayer ?? VariableDictionary.layerName ?? "0", layerColor);
                }
                // 先提交外层事务，避免后续方法再开事务冲突
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n暖通用户指定两点为对角线画方失败！错误信息: {ex.Message}");
                LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 可选：将 LTSCALE/CELTSCALE 恢复到常见的“默认”值（1.0）。仅在确实需要恢复时调用。
        /// 注意：恢复系统变量会影响整个会话，请在确认后调用。
        /// </summary>
        public static void RestoreDefaultLtscaleCeltscale()
        {
            try { Application.SetSystemVariable("LTSCALE", 1.0); } catch { }
            try { Application.SetSystemVariable("CELTSCALE", 1.0); } catch { }
            AutoCadHelper.LogWithSafety("已尝试将 LTSCALE/CELTSCALE 恢复为 1.0（如权限允许）。");
        }

        /// <summary>
        /// 结构、用户输入长宽后画矩形；
        /// </summary>
        [CommandMethod(nameof(DrawRec))]
        public static void DrawRec()
        {
            #region 方法一
            try
            {
                #region 原始矩形
                double width = Convert.ToDouble(VariableDictionary.textbox_Width); ;
                double height = Convert.ToDouble(VariableDictionary.textbox_Height);
                double recPlus = Convert.ToDouble(VariableDictionary.textbox_RecPlus_Text);// 指定的扩大偏移量
                var layerName = VariableDictionary.btnBlockLayer;
                var layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                var leftUp = new Point3d(0, 0, 0);
                var rightDown = new Point3d(leftUp.X + width, leftUp.Y + height, 0);
                var mouseEndPoint = new Point3d(0, 0, 0);
                using var tr = new DBTrans();

                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；
                var polyLineRec = new Polyline();
                polyLineRec.AddVertexAt(0, new(leftUp.X, leftUp.Y), 0, 0, 0);
                polyLineRec.AddVertexAt(1, new(leftUp.X, rightDown.Y), 0, 0, 0);
                polyLineRec.AddVertexAt(2, new(rightDown.X, rightDown.Y), 0, 0, 0);
                polyLineRec.AddVertexAt(3, new(rightDown.X, leftUp.Y), 0, 0, 0);
                polyLineRec.Closed = true;
                polyLineRec.Layer = layerName;
                polyLineRec.ColorIndex = layerColorIndex;

                #endregion
                #region 计算扩大后的矩形
                //拿到动态绘制矩形的边界
                Extents3d polyLineRecExt = (Extents3d)polyLineRec.Bounds;
                // 计算扩大矩形的边界框
                Extents3d newRectBounds = new Extents3d(
                    new Point3d(polyLineRecExt.MinPoint.X - recPlus, polyLineRecExt.MinPoint.Y - recPlus, 0),
                    new Point3d(polyLineRecExt.MaxPoint.X + recPlus, polyLineRecExt.MaxPoint.Y + recPlus, 0));
                //在生成的矩形上的两个临时点，一个左上，一下右下
                var recLeftUp = new Point3d(newRectBounds.MinPoint.X, newRectBounds.MaxPoint.Y, 0);
                var recRightDowm = new Point3d(newRectBounds.MaxPoint.X, newRectBounds.MinPoint.Y, 0);
                double totalLength = recLeftUp.DistanceTo(recRightDowm);//计算出左上与右下的辅助点的长度；
                double targetLength = totalLength * 1 / 4;//计算出1/4的点
                double targetX = recLeftUp.X + (recRightDowm.X - recLeftUp.X) * 1 / 4;
                double targetY = recLeftUp.Y + (recRightDowm.Y - recLeftUp.Y) * 1 / 4;
                double targetZ = recLeftUp.Z + (recRightDowm.Z - recLeftUp.Z) * 1 / 4;
                Point3d targetPoint = new Point3d(targetX, targetY, targetZ);//找到1/4点的坐标
                // 创建扩大矩形
                Polyline newRect = new Polyline();
                newRect.AddVertexAt(0, new Point2d(newRectBounds.MinPoint.X, newRectBounds.MinPoint.Y), 0, 0, 0);
                newRect.AddVertexAt(1, new Point2d(newRectBounds.MinPoint.X, newRectBounds.MaxPoint.Y), 0, 0, 0);
                newRect.AddVertexAt(2, new Point2d(newRectBounds.MaxPoint.X, newRectBounds.MaxPoint.Y), 0, 0, 0);
                newRect.AddVertexAt(3, new Point2d(newRectBounds.MaxPoint.X, newRectBounds.MinPoint.Y), 0, 0, 0);
                newRect.Closed = true;
                newRect.Layer = layerName;
                newRect.ColorIndex = layerColorIndex;

                if (layerName.Contains("结构"))
                {
                    // 绘制用来辅助生成填充边界的两条线（或辅助多段线）  
                    // 此处构造的填充边界区域：由扩大矩形的左下角、左上角、右上角，以及计算得到的交合点构成  
                    var hatchBoundary = new Polyline();
                    hatchBoundary.AddVertexAt(0, new Point2d(newRectBounds.MinPoint.X, newRectBounds.MinPoint.Y), 0, 0, 0);
                    hatchBoundary.AddVertexAt(1, new Point2d(newRectBounds.MinPoint.X, newRectBounds.MaxPoint.Y), 0, 0, 0);
                    hatchBoundary.AddVertexAt(2, new Point2d(newRectBounds.MaxPoint.X, newRectBounds.MaxPoint.Y), 0, 0, 0);
                    hatchBoundary.AddVertexAt(3, new Point2d(targetPoint.X, targetPoint.Y), 0, 0, 0);

                    hatchBoundary.Closed = true;
                    hatchBoundary.Layer = layerName;
                    hatchBoundary.ColorIndex = layerColorIndex;
                    #endregion
                    var newRectObjectId = tr.CurrentSpace.AddEntity(newRect);
                    var hatchBoundaryObjectId = tr.CurrentSpace.AddEntity(hatchBoundary);
                    var hatchObectId = new ObjectId();
                    //调用填充方法
                    autoHatch(tr, layerName, layerColorIndex, 50, "DOTS", hatchBoundary.ObjectId, ref hatchObectId);
                    var hatchEntity = tr.GetObject(hatchObectId, OpenMode.ForRead) as Entity;
                    var polylineIds = new ObjectId[2];
                    polylineIds[0] = newRectObjectId;
                    polylineIds[1] = hatchBoundaryObjectId;
                    using var polyLineRecMove = new JigEx((mpw, _) =>
                    {
                        newRect.Move(leftUp, mpw);
                        hatchBoundary.Move(leftUp, mpw);
                        if (hatchEntity != null)
                            hatchEntity.Move(leftUp, mpw);
                        leftUp = mpw;
                    });
                    polyLineRecMove.DatabaseEntityDraw(wd => wd.Geometry.Draw(newRect));
                    polyLineRecMove.SetOptions(msg: "指定插入点：");
                    var endPoint = Env.Editor.Drag(polyLineRecMove);
                    if (endPoint.Status != PromptStatus.OK) tr.Abort();
                    var mpwl = polyLineRecMove.MousePointWcsLast;
                    //计算洞口的中心点
                    mouseEndPoint = new Point3d(mpwl.X + (newRectBounds.MinPoint.X + newRectBounds.MaxPoint.X) / 2, mpwl.Y + (newRectBounds.MinPoint.Y + newRectBounds.MaxPoint.Y) / 2, 0);
                }
                else
                {
                    tr.CurrentSpace.AddEntity(newRect);
                    using var polyLineRecMove = new JigEx((mpw, _) =>
                    {
                        newRect.Move(leftUp, mpw);
                        leftUp = mpw;
                    });
                    polyLineRecMove.DatabaseEntityDraw(wd => wd.Geometry.Draw(newRect));
                    polyLineRecMove.SetOptions(msg: "指定插入点：");
                    var endPoint = Env.Editor.Drag(polyLineRecMove);
                    if (endPoint.Status != PromptStatus.OK) tr.Abort();
                    var mpwl = polyLineRecMove.MousePointWcsLast;
                    //计算洞口的中心点
                    mouseEndPoint = new Point3d(mpwl.X + (newRectBounds.MinPoint.X + newRectBounds.MaxPoint.X) / 2, mpwl.Y + (newRectBounds.MinPoint.Y + newRectBounds.MaxPoint.Y) / 2, 0);
                }
                var textLayerColorIndex = Convert.ToInt16(VariableDictionary.textColorIndex);
                Env.Editor.Redraw();
                DDimLinear(tr, (width + recPlus).ToString("0" + "mm"), (height + recPlus).ToString("0" + "mm"), textLayerColorIndex, mouseEndPoint);
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n结构指定数值生成矩形失败: {ex.Message}");
            }
            #endregion
        }

        /// <summary>
        /// 后并多图元返回实体objectid
        /// </summary>
        /// <param name="tr">tr事件</param>
        /// <param name="polylineIds">多个PL线Id</param>
        /// <param name="hatchId">填充Id</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>

        public static ObjectId CombineEntitiesToRegion(DBTrans tr, ObjectId[] polylineIds, ObjectId hatchId)
        {
            // 创建区域
            DBObjectCollection polylineCollection = new DBObjectCollection();
            foreach (ObjectId polylineId in polylineIds)
            {
                // 获取多段线对象
                Entity polyline = tr.GetObject<Entity>(polylineId, OpenMode.ForWrite);
                // 将多段线添加到集合中
                polylineCollection.Add(polyline);
            }

            // 从多段线集合创建区域
            DBObjectCollection regionCollection = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(polylineCollection);

            // 检查是否成功创建了一个区域
            if (regionCollection.Count != 1)
            {
                throw new InvalidOperationException("无法从多段线创建单个区域");
            }
            // 获取创建的区域
            Autodesk.AutoCAD.DatabaseServices.Region region = regionCollection[0] as Autodesk.AutoCAD.DatabaseServices.Region;
            // 将填充应用到区域
            Hatch hatch = tr.GetObject<Hatch>(hatchId, OpenMode.ForWrite);
            // 设置填充为关联填充
            hatch.Associative = true;
            // 将区域作为填充的外环
            hatch.AppendLoop(HatchLoopTypes.Outermost, new ObjectIdCollection { region.ObjectId });
            // 将区域添加到当前空间
            ObjectId regionId = tr.CurrentSpace.AddEntity(region);
            // 返回区域的 ObjectId
            return regionId;
        }

        /// <summary>
        /// 多点画不规则图形并填充-面着地
        /// </summary>
        [CommandMethod(nameof(NLinePolyline))]
        public static void NLinePolyline()
        {
            try
            {
                var layerName = VariableDictionary.btnBlockLayer;
                var layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                using var tr = new DBTrans();

                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；
                var userPoint1 = Env.Editor.GetPoint("\n指定多边形的第一个点：");
                if (userPoint1.Status != PromptStatus.OK) return;
                var UcsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();
                pointS.Add(UcsUserPoint1);
                while (true)
                {
                    using var polyLine = new JigEx((mpw, queue) =>
                    {
                        var UcsUserPoint2 = mpw.Z20();
                        Polyline polyline1 = new Polyline()
                        {
                            Layer = layerName,
                            ColorIndex = 231
                        };
                        for (int i = 0; i < pointS.Count; i++)
                        {
                            polyline1.AddVertexAt(i, new Point2d(pointS[i].X, pointS[i].Y), 0, 0, 0);
                        }
                        polyline1.AddVertexAt(pointS.Count, new Point2d(UcsUserPoint2.X, UcsUserPoint2.Y), 0, 0, 0);
                        if (pointS.Count >= 2)
                        {
                            polyline1.Closed = true;
                        }
                        queue.Enqueue(polyline1);
                    });

                    polyLine.SetOptions(UcsUserPoint1, msg: "\n指定多边形的下一个点（右键结束）：");
                    var userPoint2 = Env.Editor.Drag(polyLine);
                    if (userPoint2.Status != PromptStatus.OK) break;
                    var UcsUserPoint2 = polyLine.MousePointWcsLast;
                    pointS.Add(UcsUserPoint2);
                    UcsUserPoint1 = UcsUserPoint2;
                }

                if (pointS.Count >= 3)
                {
                    // 创建最终的多段线  
                    Polyline finalPolyline = new Polyline();
                    for (int i = 0; i < pointS.Count; i++)
                    {
                        finalPolyline.AddVertexAt(i, new Point2d(pointS[i].X, pointS[i].Y), 0, 0, 0);
                        finalPolyline.SetStartWidthAt(i, 30);
                        finalPolyline.SetEndWidthAt(i, 30);
                    }
                    finalPolyline.Closed = true;
                    finalPolyline.Layer = layerName;
                    finalPolyline.ColorIndex = 231;

                    // 添加多段线并获取其ID  
                    var polylineId = tr.CurrentSpace.AddEntity(finalPolyline);

                    if (layerName != null)
                        //调用填充方法
                        autoHatch(tr, layerName, 231, 200, "ANSI38", polylineId);
                    Env.Editor.Redraw();
                }
                else
                {
                    LogManager.Instance.LogInfo("\n至少需要三个点才能绘制闭合多边形。");
                }
                pointS.Clear();
                if (VariableDictionary.dimString != null)
                    DDimLinear(tr, VariableDictionary.dimString);
                tr.Commit();
                Env.Editor.Redraw();// 强制刷新视图
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n多点画不规则图形并填充失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 多点画不规则图形不填充-框着地
        /// </summary>
        [CommandMethod(nameof(NLinePolyline_Not))]
        public static void NLinePolyline_Not()
        {
            try
            {
                pointS.Clear();
                var layerName = VariableDictionary.btnBlockLayer; // 设置图层名称
                Int16 layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex == null ? VariableDictionary.layerColorIndex : 231); // 设置图层颜色索引
                using var tr = new DBTrans();//开启事务
                                             // 检查图层是否存在，如果不存在则创建

                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；
                // 获取第一个点
                var userPoint1 = Env.Editor.GetPoint("\n指定多边形的第一个点：");
                if (userPoint1.Status != PromptStatus.OK) return;
                // 将第一个点转换为 UCS 坐标并存储
                var UcsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();
                pointS.Add(UcsUserPoint1);
                while (true)
                {
                    // 使用 JigEx 动态绘制多段线
                    using var polyLine = new JigEx((mpw, queue) =>
                    {
                        var UcsUserPoint2 = mpw.Z20(); // 将当前鼠标点转换为 UCS 坐标
                                                       // 创建多段线
                        Polyline polyline1 = new Polyline()
                        {
                            Layer = layerName,  // 设置图层
                            ColorIndex = 231
                        };
                        for (int i = 0; i < pointS.Count; i++)
                        {
                            polyline1.AddVertexAt(i, new Point2d(pointS[i].X, pointS[i].Y), 0, 0, 0);
                        }
                        // 添加当前鼠标点作为临时顶点
                        polyline1.AddVertexAt(pointS.Count, new Point2d(UcsUserPoint2.X, UcsUserPoint2.Y), 0, 0, 0);
                        // 如果点数大于 2，则闭合多边形
                        if (pointS.Count >= 2)
                        {
                            polyline1.Closed = true;
                        }
                        queue.Enqueue(polyline1); // 将多段线加入绘制队列
                    });
                    // 设置 JigEx 的起始点和提示信息
                    polyLine.SetOptions(UcsUserPoint1, msg: "\n指定多边形的下一个点（右键结束）：");
                    // 获取用户输入的下一个点
                    var userPoint2 = Env.Editor.Drag(polyLine);
                    if (userPoint2.Status != PromptStatus.OK) break;
                    // 将用户输入的点转换为 UCS 坐标并存储
                    var UcsUserPoint2 = polyLine.MousePointWcsLast;
                    pointS.Add(UcsUserPoint2);
                    // 更新起始点为当前点
                    UcsUserPoint1 = UcsUserPoint2;
                    // 刷新视图
                    Env.Editor.Redraw();
                }
                // 如果点数大于 2，则闭合多边形并添加到模型空间
                if (pointS.Count >= 3)
                {
                    // 创建最终的多段线
                    Polyline finalPolyline = new Polyline();
                    for (int i = 0; i < pointS.Count; i++)
                    {
                        finalPolyline.AddVertexAt(i, new Point2d(pointS[i].X, pointS[i].Y), 0, 0, 0);
                        // 设置线宽为 30
                        finalPolyline.SetStartWidthAt(i, 30);
                        finalPolyline.SetEndWidthAt(i, 30);
                    }
                    finalPolyline.Closed = true; // 闭合多边形
                    finalPolyline.Layer = layerName; // 设置图层
                    finalPolyline.ColorIndex = 231;
                    var polylineId = tr.CurrentSpace.AddEntity(finalPolyline);// 将多段线添加到模型空间
                    Env.Editor.Redraw();// 强制刷新视图
                    LogManager.Instance.LogInfo("\n多边形绘制完成并闭合，填充图案已添加。");
                }
                else
                {
                    LogManager.Instance.LogInfo("\n至少需要三个点才能绘制闭合多边形。");
                }
                pointS.Clear();  // 清空点列表

                while (true)
                {
                    try
                    {
                        // 创建 PromptPointOptions，并允许空输入（即右键取消时不产生错误提示）  
                        PromptPointOptions ppo = new PromptPointOptions("\n请指定框着地内线第一点");
                        ppo.AllowNone = true;

                        // 获取第一个点（左键点击有效，右键取消则返回 None）  
                        var userPointX1 = Env.Editor.GetPoint(ppo);
                        if (userPointX1.Status != PromptStatus.OK)
                        {
                            // 如果状态非 OK，则退出循环  
                            break;
                        }

                        var UcsUserPointX1 = userPointX1.Value.Wcs2Ucs().Z20(); // 转换为 UCS 坐标  

                        using var polylineX = new JigEx((mpw, queue) =>
                        {
                            var UcsUserPointX2 = mpw.Z20();
                            // 定义第一条直线  
                            Polyline polylineX1 = new Polyline();
                            polylineX1.AddVertexAt(0, new Point2d(UcsUserPointX1.X, UcsUserPointX1.Y), 0, 0, 0);
                            polylineX1.AddVertexAt(1, new Point2d(UcsUserPointX2.X, UcsUserPointX2.Y), 0, 0, 0);
                            polylineX1.Closed = false;
                            polylineX1.Layer = layerName; // 设置线条图层  
                            polylineX1.ColorIndex = 231;  // 设置线条颜色  
                            polylineX1.SetStartWidthAt(0, 30);
                            polylineX1.SetEndWidthAt(0, 30);
                            queue.Enqueue(polylineX1);

                            // 定义闭合的方形  
                            Polyline polylineX2 = new Polyline();
                            polylineX2.AddVertexAt(0, new Point2d(UcsUserPointX2.X, UcsUserPointX1.Y), 0, 0, 0);
                            polylineX2.AddVertexAt(1, new Point2d(UcsUserPointX1.X, UcsUserPointX2.Y), 0, 0, 0);
                            polylineX2.Closed = true; // 闭合成方形  
                            polylineX2.Layer = layerName; // 设置线条图层  
                            polylineX2.ColorIndex = 231;
                            polylineX2.SetStartWidthAt(0, 30);
                            polylineX2.SetEndWidthAt(0, 30);
                        });

                        polylineX.SetOptions(UcsUserPointX1, msg: "\n请指定框着地内线第二点");
                        // 获取拖曳过程中的第二个点  
                        var userPointX2 = Env.Editor.Drag(polylineX);
                        if (userPointX2.Status == PromptStatus.OK)
                        {
                            // 如果拖曳成功，则将生成的图元加入当前空间并刷新界面  
                            tr.CurrentSpace.AddEntity(polylineX.Entities);
                            Env.Editor.Redraw();
                        }
                        else
                        {
                            // 如果拖曳过程中右键取消，同样退出循环  
                            break;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        // 记录错误日志  
                        LogManager.Instance.LogInfo($"\n多点画不规则图形不填充-框着地失败！错误信息: {ex.Message}");
                        LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}");
                    }
                }
                if (VariableDictionary.dimString != null)
                    DDimLinear(tr, VariableDictionary.dimString);
                tr.Commit(); // 提交事务
                Env.Editor.Redraw();
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n多点画不规则图形不填充-框着地失败: {ex.Message}");
                LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 结构、画多边形-水平荷载
        /// </summary>
        [CommandMethod(nameof(NLinePolyline_N))]
        public static void NLinePolyline_N()
        {
            try
            {
                pointS.Clear();
                var layerName = VariableDictionary.btnBlockLayer; // 设置图层名称
                Int16 layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex == null ? VariableDictionary.layerColorIndex : 231); // 设置图层颜色索引
                using var tr = new DBTrans();//开启事务
                // 检查图层是否存在，如果不存在则创建
                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；
                                                                                        // 获取第一个点
                var userPoint1 = Env.Editor.GetPoint("\n指定多边形的第一个点：");
                if (userPoint1.Status != PromptStatus.OK) return;
                // 将第一个点转换为 UCS 坐标并存储
                var UcsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();
                pointS.Add(UcsUserPoint1);
                while (true)
                {
                    // 使用 JigEx 动态绘制多段线
                    using var polyLine = new JigEx((mpw, queue) =>
                    {
                        var UcsUserPoint2 = mpw.Z20(); // 将当前鼠标点转换为 UCS 坐标
                                                       // 创建多段线
                        Polyline polyline1 = new Polyline()
                        {
                            Layer = layerName,  // 设置图层
                            ColorIndex = 231
                        };
                        for (int i = 0; i < pointS.Count; i++)
                        {
                            polyline1.AddVertexAt(i, new Point2d(pointS[i].X, pointS[i].Y), 0, 0, 0);
                        }
                        // 添加当前鼠标点作为临时顶点
                        polyline1.AddVertexAt(pointS.Count, new Point2d(UcsUserPoint2.X, UcsUserPoint2.Y), 0, 0, 0);

                        if (pointS.Count >= 2)// 如果点数大于 2，则闭合多边形
                        {
                            polyline1.Closed = true;
                        }
                        queue.Enqueue(polyline1); // 将多段线加入绘制队列
                    });
                    // 设置 JigEx 的起始点和提示信息
                    polyLine.SetOptions(UcsUserPoint1, msg: "\n指定多边形的下一个点（右键结束）");
                    var userPoint2 = Env.Editor.Drag(polyLine); // 获取用户输入的下一个点
                    if (userPoint2.Status != PromptStatus.OK) break;
                    var UcsUserPoint2 = polyLine.MousePointWcsLast;// 将用户输入的点转换为 UCS 坐标并存储
                    pointS.Add(UcsUserPoint2);
                    UcsUserPoint1 = UcsUserPoint2;// 更新起始点为当前点
                    Env.Editor.Redraw();// 刷新视图
                }
                // 如果点数大于 2，则闭合多边形并添加到模型空间
                if (pointS.Count >= 3)
                {
                    // 创建最终的多段线
                    Polyline finalPolyline = new Polyline();
                    for (int i = 0; i < pointS.Count; i++)
                    {
                        finalPolyline.AddVertexAt(i, new Point2d(pointS[i].X, pointS[i].Y), 0, 0, 0);
                        // 设置线宽为 30
                        finalPolyline.SetStartWidthAt(i, 30);
                        finalPolyline.SetEndWidthAt(i, 30);
                    }
                    finalPolyline.Closed = true; // 闭合多边形
                    finalPolyline.Layer = layerName; // 设置图层
                    finalPolyline.ColorIndex = 231;
                    var polylineId = tr.CurrentSpace.AddEntity(finalPolyline);// 将多段线添加到模型空间
                    if (layerName != null)
                        DrawArrows(finalPolyline, layerName);
                    Env.Editor.Redraw();// 强制刷新视图
                    LogManager.Instance.LogInfo("\n多边形绘制完成并闭合，填充图案已添加。");
                    pointS.Clear();  // 清空点列表

                }
                else
                {
                    LogManager.Instance.LogInfo("\n至少需要三个点才能绘制闭合多边形。");
                }
                if (VariableDictionary.dimString != null)
                    DDimLinear(tr, VariableDictionary.dimString);
                tr.Commit();
                Env.Editor.Redraw();// 强制刷新视图

            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n结构、画多边形失败: {ex.Message}");
                // 记录错误日志  
                LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 生成箭头
        /// </summary>
        /// <param name="existingPolygon">Polyline多边形</param>
        /// <param name="layerName"></param>
        public static void DrawArrows(Polyline existingPolygon, string layerName)
        {
            try
            {
                using var tr = new DBTrans();

                // 获取当前图比例（优先使用已缓存/全局值）
                double uiScale = VariableDictionary.blockScale;
                if (double.IsNaN(uiScale) || uiScale <= 0)
                {
                    try { uiScale = AutoCadHelper.GetScale(true); } catch { uiScale = 1.0; }
                }
                if (uiScale <= 0) uiScale = 1.0;

                // 获取多边形的边界  
                Extents3d bounds = existingPolygon.GeometricExtents;
                double width = bounds.MaxPoint.X - bounds.MinPoint.X;
                double height = bounds.MaxPoint.Y - bounds.MinPoint.Y;

                // 计算箭头的基本尺寸（按图面尺寸与当前比例缩放）
                double arrowWidth = width * 0.8 * uiScale; // 箭头总长度
                double arrowHeight = Math.Max(1.0, height * 0.3 * uiScale); // 箭头高度（最小保护）
                double arrowHeadWidth = Math.Max(1.0, width * 0.2 * uiScale); // 箭头底坐宽度
                double margin = Math.Max(1.0, width * 0.1 * uiScale); // 左右边距

                // 计算两个箭头的垂直位置  
                double firstArrowY = bounds.MinPoint.Y + height * 0.35;
                double secondArrowY = bounds.MinPoint.Y + height * 0.65;

                // 创建第一个箭头（向右）  
                Polyline arrow1 = new Polyline();
                double x1 = bounds.MinPoint.X + margin;
                arrow1.AddVertexAt(0, new Point2d(x1, firstArrowY), 0, 0, 0);  // 起点  
                arrow1.AddVertexAt(1, new Point2d(x1 + arrowWidth - arrowHeadWidth, firstArrowY), 0, 0, 0);  // 线条终点  
                arrow1.AddVertexAt(2, new Point2d(x1 + arrowWidth - arrowHeadWidth, firstArrowY - arrowHeight / 2), 0, 0, 0);  // 箭头底部  
                arrow1.AddVertexAt(3, new Point2d(x1 + arrowWidth, firstArrowY + arrowHeight / 8), 0, 0, 0); // 箭头尖端  
                arrow1.AddVertexAt(4, new Point2d(x1 + arrowWidth - arrowHeadWidth, firstArrowY + arrowHeight / 1.4), 0, 0, 0);  // 箭头顶部  
                arrow1.AddVertexAt(5, new Point2d(x1 + arrowWidth - arrowHeadWidth, firstArrowY + arrowHeight / 4), 0, 0, 0);  // 回到线条  
                arrow1.AddVertexAt(6, new Point2d(x1, firstArrowY + arrowHeight / 4), 0, 0, 0);  // 线条起点上边  
                arrow1.Closed = true;
                arrow1.Layer = layerName;
                arrow1.ColorIndex = 231;
                tr.CurrentSpace.AddEntity(arrow1);

                // 创建第二个箭头（向左）  
                Polyline arrow2 = new Polyline();
                double x2 = bounds.MaxPoint.X - margin;
                arrow2.AddVertexAt(0, new Point2d(x2, secondArrowY), 0, 0, 0);  // 起点  
                arrow2.AddVertexAt(1, new Point2d(x2 - arrowWidth + arrowHeadWidth, secondArrowY), 0, 0, 0);  // 线条终点  
                arrow2.AddVertexAt(2, new Point2d(x2 - arrowWidth + arrowHeadWidth, secondArrowY - arrowHeight / 2), 0, 0, 0);  // 箭头底部  
                arrow2.AddVertexAt(3, new Point2d(x2 - arrowWidth, secondArrowY + arrowHeight / 8), 0, 0, 0); // 箭头尖端  
                arrow2.AddVertexAt(4, new Point2d(x2 - arrowWidth + arrowHeadWidth, secondArrowY + arrowHeight / 1.4), 0, 0, 0);  // 箭头顶部  
                arrow2.AddVertexAt(5, new Point2d(x2 - arrowWidth + arrowHeadWidth, secondArrowY + arrowHeight / 4), 0, 0, 0);  // 回到线条  
                arrow2.AddVertexAt(6, new Point2d(x2, secondArrowY + arrowHeight / 4), 0, 0, 0);  // 线条起点上边  
                arrow2.Closed = true;
                arrow2.Layer = layerName;
                arrow2.ColorIndex = 231;
                tr.CurrentSpace.AddEntity(arrow2);

                // 为箭头添加填充（SOLID），并让填充的显示比例随 uiScale 调整（保护性值）
                using (Hatch hatch1 = new Hatch())
                {
                    hatch1.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                    hatch1.AppendLoop(HatchLoopTypes.Outermost, new ObjectIdCollection { arrow1.ObjectId });
                    hatch1.PatternScale = Math.Max(1.0, 100 * uiScale);
                    hatch1.Layer = layerName;
                    hatch1.ColorIndex = 231;
                    hatch1.EvaluateHatch(true);
                    tr.CurrentSpace.AddEntity(hatch1);
                }
                using (Hatch hatch2 = new Hatch())
                {
                    hatch2.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                    hatch2.AppendLoop(HatchLoopTypes.Outermost, new ObjectIdCollection { arrow2.ObjectId });
                    hatch2.PatternScale = Math.Max(1.0, 100 * uiScale);
                    hatch2.Layer = layerName;
                    hatch2.ColorIndex = 231;
                    hatch2.EvaluateHatch(true);
                    tr.CurrentSpace.AddEntity(hatch2);
                }

                tr.Commit();
                Env.Editor.Redraw();  // 强制刷新视图
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n结构生成箭头失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 多边形PL线
        /// </summary>
        [CommandMethod(nameof(NLinePolyline_K))]
        public static void NLinePolyline_K()
        {
            try
            {
                pointS.Clear();
                var layerName = VariableDictionary.btnBlockLayer; // 设置图层名称
                Int16 layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex == null ? VariableDictionary.layerColorIndex : 231); // 设置图层颜色索引
                using var tr = new DBTrans();//开启事务
                // 检查图层是否存在，如果不存在则创建
                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；

                // 获取第一个点
                var userPoint1 = Env.Editor.GetPoint("\n指定多边形的第一个点：");
                if (userPoint1.Status != PromptStatus.OK) return;

                // 将第一个点转换为 UCS 坐标并存储
                var UcsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();
                pointS.Add(UcsUserPoint1);
                while (true)
                {
                    // 使用 JigEx 动态绘制多段线
                    using var polyLine = new JigEx((mpw, queue) =>
                    {
                        var UcsUserPoint2 = mpw.Z20(); // 将当前鼠标点转换为 UCS 坐标
                        Polyline polyline1 = new Polyline() // 创建多段线
                        {
                            Layer = layerName,  // 设置图层
                            ColorIndex = 231
                        };
                        for (int i = 0; i < pointS.Count; i++)
                        {
                            polyline1.AddVertexAt(i, new Point2d(pointS[i].X, pointS[i].Y), 0, 0, 0);
                        }
                        // 添加当前鼠标点作为临时顶点
                        polyline1.AddVertexAt(pointS.Count, new Point2d(UcsUserPoint2.X, UcsUserPoint2.Y), 0, 0, 0);
                        if (pointS.Count >= 2) // 如果点数大于 2，则闭合多边形
                        {
                            polyline1.Closed = true;
                        }
                        queue.Enqueue(polyline1); // 将多段线加入绘制队列
                    });
                    polyLine.SetOptions(UcsUserPoint1, msg: "\n指定多边形的下一个点（右键结束）：");  // 设置 JigEx 的起始点和提示信息
                    var userPoint2 = Env.Editor.Drag(polyLine);  // 获取用户输入的下一个点
                    if (userPoint2.Status != PromptStatus.OK) break;
                    var UcsUserPoint2 = polyLine.MousePointWcsLast; // 将用户输入的点转换为 UCS 坐标并存储
                    pointS.Add(UcsUserPoint2);
                    UcsUserPoint1 = UcsUserPoint2; // 更新起始点为当前点
                    Env.Editor.Redraw();// 刷新视图
                }

                if (pointS.Count >= 3)// 如果点数大于 2，则闭合多边形并添加到模型空间
                {
                    Polyline finalPolyline = new Polyline();  // 创建最终的多段线
                    for (int i = 0; i < pointS.Count; i++)
                    {
                        finalPolyline.AddVertexAt(i, new Point2d(pointS[i].X, pointS[i].Y), 0, 0, 0);
                        // 设置线宽为 30
                        finalPolyline.SetStartWidthAt(i, 30);
                        finalPolyline.SetEndWidthAt(i, 30);
                    }
                    finalPolyline.Closed = true; // 闭合多边形
                    finalPolyline.Layer = layerName; // 设置图层
                    finalPolyline.ColorIndex = 231;
                    var polylineId = tr.CurrentSpace.AddEntity(finalPolyline); // 将多段线添加到模型空间
                    Extents3d polygonBounds = finalPolyline.Bounds.Value; // 计算多边形的边界
                    double arrowHeight = 50; // 定义箭头的高度
                    double polygonWidth = polygonBounds.MaxPoint.X - polygonBounds.MinPoint.X; // 多边形的宽度
                                                                                               // 计算箭头的中心点
                    double centerYTop = polygonBounds.MaxPoint.Y - arrowHeight / 2; // 上箭头的中心点 Y 坐标
                    double centerYBottom = polygonBounds.MinPoint.Y + arrowHeight / 2; // 下箭头的中心点 Y 坐标
                    double centerX = (polygonBounds.MinPoint.X + polygonBounds.MaxPoint.X) / 2; // 箭头的中心点 X 坐标
                    if (layerName != null)
                    {
                        // 绘制上箭头（向左）
                        DrawArrow(tr, new Point3d(centerX, centerYTop, 0), -90, polygonWidth, arrowHeight, layerName);
                        // 绘制下箭头（向右）
                        DrawArrow(tr, new Point3d(centerX, centerYBottom, 0), 90, polygonWidth, arrowHeight, layerName);
                    }

                    if (VariableDictionary.dimString != null)
                        DDimLinear(tr, VariableDictionary.dimString);
                    Env.Editor.Redraw(); // 强制刷新视图
                    LogManager.Instance.LogInfo("\n多边形绘制完成并闭合，箭头已添加。");
                }
                else
                {
                    LogManager.Instance.LogInfo("\n至少需要三个点才能绘制闭合多边形。");
                }
                pointS.Clear(); // 清空点列表
                                //if (VariableDictionary.dimString != null)
                                //    DDimLinear(VariableDictionary.dimString);
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (System.Exception ex)
            {

                // 记录错误日志  
                LogManager.Instance.LogInfo($"\n结构生成箭头失败！错误信息: {ex.Message}");
                LogManager.Instance.LogInfo($"\n错误堆栈: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 绘制箭头
        /// </summary>
        /// <param name="tr">事务</param>
        /// <param name="center">中心</param>
        /// <param name="angle">角度</param>
        /// <param name="arrowLength">箭头长</param>
        /// <param name="arrowHeight">箭头高</param>
        /// <param name="layerName">图层名</param>
        private static void DrawArrow(DBTrans tr, Point3d center, double angle, double arrowLength, double arrowHeight, string layerName)
        {
            // 计算箭头的三个顶点
            Point3d tip = center.PolarPoint(angle, arrowLength / 2); // 箭头尖端
            Point3d left = center.PolarPoint(angle - 90, arrowHeight / 2); // 左端点
            Point3d right = center.PolarPoint(angle + 90, arrowHeight / 2); // 右端点

            // 创建箭头的多段线
            Polyline arrow = new Polyline();
            arrow.AddVertexAt(0, new Point2d(tip.X, tip.Y), 0, 0, 0);
            arrow.AddVertexAt(1, new Point2d(left.X, left.Y), 0, 0, 0);
            arrow.AddVertexAt(2, new Point2d(right.X, right.Y), 0, 0, 0);
            arrow.Closed = true; // 闭合箭头
            arrow.Layer = layerName; // 设置图层
            arrow.ColorIndex = 231;
            tr.CurrentSpace.AddEntity(arrow);// 将箭头添加到模型空间
        }

        #endregion

        #region 接口、实体、天正数据、标注等、正交角度等工具方法


        /// <summary>
        /// 给体id返回实体对像
        /// </summary>
        /// <param name="entityId">输入要返回实体对像的Object</param>
        /// <returns>返回实体对像</returns>
        public static Entity GetEntity(ObjectId entityId)
        {
            Entity? entity = null;
            try
            {
                using (Transaction tr = entityId.Database.TransactionManager.StartTransaction())
                {
                    entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                    tr.Commit();
                }

            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("给体id返回实体对像失败！");
                LogManager.Instance.LogInfo(ex.Message);
            }
            return entity;
        }

        /// <summary>
        /// 获取引用实体名称
        /// </summary>
        /// <param name="tr">开启事务</param>
        /// <param name="objectId">objectId</param>
        /// <returns></returns>
        public static string getXrefName(DBTrans tr, ObjectId objectId)
        {
            string? xrefName;
            // 第三步：打开选中的外部参照
            BlockReference xrefEntity = tr.GetObject(objectId, OpenMode.ForRead) as BlockReference;

            if (xrefEntity == null)
            {
                LogManager.Instance.LogInfo("\n错误：选中的对象不是外部参照。");
                return xrefName = "";
            }

            // 第四步：获取外部参照的块表记录
            BlockTableRecord btr = tr.GetObject(xrefEntity.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

            if (btr == null)
            {
                LogManager.Instance.LogInfo("\n错误：无法获取块表记录。");
                return xrefName = "";
            }
            return xrefName = btr.Name;
        }

        /// <summary>
        /// 调用正交角度
        /// </summary>
        /// <param name="rDim">标注</param>
        /// <param name="dimPoint2">坐标点</param>
        public static void SetDimensionRotationToNearest90Degrees(RotatedDimension rDim, Point3d dimPoint2)
        {
            // 获取当前旋转角度
            double currentAngle = rDim.Rotation;

            // 计算向量
            //Vector3d vectorToDimPoint = rDim.Position.GetVectorTo(dimPoint2);

            var vectorToDimPoint = rDim.TextPosition.GetVectorTo(dimPoint2);

            // 计算向量与X轴的角度
            double angleToXAxis = vectorToDimPoint.GetAngleTo(Vector3d.XAxis);

            // 将角度转换为0-360度范围
            angleToXAxis = angleToXAxis * (180 / Math.PI);
            if (angleToXAxis < 0)
            {
                angleToXAxis += 360;
            }

            // 找到最接近的0、90、180、270度
            double nearestAngle = FindNearestAngle(angleToXAxis);

            // 设置旋转角度
            rDim.Rotation = nearestAngle * (Math.PI / 180);
        }

        /// <summary>
        /// 计算正交角度
        /// </summary>
        /// <param name="angle">任意角度</param>
        /// <returns></returns>
        private static double FindNearestAngle(double angle)
        {
            double[] targetAngles = { 0, 90, 180, 270 };
            double nearestAngle = targetAngles[0];
            double minDifference = Math.Abs(angle - targetAngles[0]);

            for (int i = 1; i < targetAngles.Length; i++)
            {
                double difference = Math.Abs(angle - targetAngles[i]);
                if (difference < minDifference)
                {
                    minDifference = difference;
                    nearestAngle = targetAngles[i];
                }
            }

            return nearestAngle;
        }

        /// <summary>
        /// 选择实体
        /// </summary>
        [CommandMethod("SelectEntities")]
        public void SelectEntities()
        {
            try
            {
                var tr = new DBTrans();
                double kwSum = 0;
                //int i = 0;
                // 提示用户选择图元
                PromptSelectionResult selectionResult = tr.Editor.GetSelection();
                if (selectionResult.Status == PromptStatus.OK)
                {
                    SelectionSet selectionSet = selectionResult.Value;

                    foreach (SelectedObject selectedObject in selectionSet)
                    {
                        if (selectedObject.ObjectId.ObjectClass.DxfName == "TEXT")
                        {
                            var kwTextString = tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead) as DBText;
                            if (kwTextString != null)
                            {

                                Match kwTextMatch = Regex.Match(kwTextString.TextString.ToLower(), @"\d+(\.\d+)?");
                                //i++;
                                if (kwTextMatch.Success)
                                {
                                    // 将匹配到的数字转换为double类型，并加到总和中
                                    double number = double.Parse(kwTextMatch.Value);
                                    kwSum += number;
                                }
                            }

                        }
                    }
                    var kwSumString = kwSum.ToString();
                    sendSum?.Invoke(kwSumString + "kw");//与下面的表达式相同，判断是不是为空，真时就调用传值；
                }
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("选择实体失败！");
            }
        }

        /// <summary>
        /// 获取外参实体的ObjectId的List
        /// </summary>
        /// <param name="xrefId">要获取的外参实体的ObjectId</param>
        /// <returns>返回外参实体的ObjectId的List</returns>
        public static List<ObjectId> GetXrefEntities(ObjectId xrefId)
        {
            List<ObjectId> entities = new List<ObjectId>();
            try
            {
                using (var tr = new DBTrans())
                {
                    BlockReference xref = tr.GetObject(xrefId, OpenMode.ForRead) as BlockReference;
                    if (xref != null)
                    {
                        BlockTableRecord xrefBtr = tr.GetObject(xref.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        if (xrefBtr is not null)
                            foreach (ObjectId entityId in xrefBtr)
                            {
                                entities.Add(entityId);
                            }
                    }
                    tr.Commit();
                }
            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("插入图元失败！");
                LogManager.Instance.LogInfo("错误信息: " + ex.Message);
            }
            return entities;
        }

        #endregion

        #region 建筑绘图

        /// <summary>
        /// 建筑、用户指定两点吊顶区Line2Polyline
        /// </summary>
        /// <param name="layerName"></param>
        [CommandMethod(nameof(Line2Polyline))]
        public static void Line2Polyline()
        {
            try
            {
                var layerName = VariableDictionary.btnBlockLayer;
                Int16 layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex == null ? VariableDictionary.layerColorIndex : 0); // 设置图层颜色索引
                var scale = VariableDictionary.textBoxScale;
                using var tr = new DBTrans();//开启事务
                // 检查图层是否存在，如果不存在则创建
                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；
                // 确保文字样式存在（DASH 里会引用）              
                TextFontsStyleHelper.EnsureTextStyle(tr, "tJText");
                var userPoint1 = Env.Editor.GetPoint("\n请指定第一点");
                if (userPoint1.Status != PromptStatus.OK) return;
                var UcsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20();//转换为UCS坐标  

                using var polyLine = new JigEx((mpw, queue) =>
                {
                    var UcsUserPoint2 = mpw.Z20();
                    Polyline polyline1 = new Polyline();
                    polyline1.AddVertexAt(0, new Point2d(UcsUserPoint1.X, UcsUserPoint1.Y), 0, 0, 0);
                    polyline1.AddVertexAt(1, new Point2d(UcsUserPoint2.X, UcsUserPoint2.Y), 0, 0, 0);
                    polyline1.Closed = false;
                    polyline1.Layer = layerName;
                    polyline1.ColorIndex = VariableDictionary.layerColorIndex;
                    queue.Enqueue(polyline1);

                    Polyline polyline2 = new Polyline();
                    polyline2.AddVertexAt(0, new Point2d(UcsUserPoint2.X, UcsUserPoint1.Y), 0, 0, 0);
                    polyline2.AddVertexAt(1, new Point2d(UcsUserPoint1.X, UcsUserPoint2.Y), 0, 0, 0);
                    polyline2.Closed = false;
                    polyline2.Layer = layerName;
                    polyline2.ColorIndex = VariableDictionary.layerColorIndex;
                    // 计算线的角度（弧度）  
                    double angle = Math.Atan2(UcsUserPoint2.Y - UcsUserPoint1.Y, UcsUserPoint2.X - UcsUserPoint1.X);

                    // 计算线的中点  
                    Point3d midPoint = new Point3d(
                        (UcsUserPoint1.X + UcsUserPoint2.X) / 2,
                        (UcsUserPoint1.Y + UcsUserPoint2.Y) / 2,
                        0
                    );
                    // 文字偏移距离  
                    double offsetDistance = 2 * scale; // 可以根据需要调整  
                                                       // 计算垂直偏移向量  
                    Vector3d offsetVector = new Vector3d(-Math.Sin(angle), Math.Cos(angle), 0) * offsetDistance;
                    // 计算文字插入点（线上方）  
                    Point3d textPoint = midPoint + offsetVector;
                    // 将角度转换为度  
                    double angleDegrees = angle * 180 / Math.PI;
                    // 确保文字方向正确（不会上下颠倒）  
                    if (angleDegrees > 90 || angleDegrees < -90)
                    {
                        angleDegrees += 180;
                        angle += Math.PI;
                        textPoint = midPoint - offsetVector;
                    }

                    if (VariableDictionary.btnFileName == "JZTJ_不吊顶")
                    {
                        queue.Enqueue(polyline2);
                        DBText text = new DBText()
                        {
                            TextStyleId = tr.TextStyleTable["tJText"],
                            TextString = "不吊顶",
                            Height = 3.50 * scale,
                            WidthFactor = 0.7,
                            ColorIndex = VariableDictionary.layerColorIndex,
                            Layer = layerName,
                            Position = textPoint,
                            //Rotation = angle,
                            HorizontalMode = TextHorizontalMode.TextCenter,
                            VerticalMode = TextVerticalMode.TextVerticalMid,
                            AlignmentPoint = textPoint
                        };
                        queue.Enqueue(text);
                    }
                    else
                    {
                        string diaoDingHeight = "2.8";
                        if (VariableDictionary.winForm_Status) diaoDingHeight = VariableDictionary.winFormDiaoDingHeight;
                        else diaoDingHeight = VariableDictionary.wpfDiaoDingHeight;

                        DBText text = new DBText()
                        {
                            TextStyleId = tr.TextStyleTable["tJText"],
                            TextString = "吊顶高度:" + diaoDingHeight + "米",
                            Height = 3.50 * scale,
                            WidthFactor = 0.7,
                            ColorIndex = VariableDictionary.layerColorIndex,
                            Layer = layerName,
                            Position = textPoint,
                            Rotation = angle,
                            //HorizontalMode = TextHorizontalMode.TextCenter,
                            HorizontalMode = TextHorizontalMode.TextCenter,
                            //VerticalMode = TextVerticalMode.TextVerticalMid,
                            VerticalMode = TextVerticalMode.TextVerticalMid,
                            AlignmentPoint = textPoint
                        };
                        queue.Enqueue(text);
                    }
                });

                polyLine.SetOptions(UcsUserPoint1, msg: "\n请指定第二点");
                var userPoint2 = Env.Editor.Drag(polyLine);
                if (userPoint2.Status != PromptStatus.OK) return;
                var polyLineEntityObj = tr.CurrentSpace.AddEntity(polyLine.Entities);
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("建筑、用户指定两点吊顶区Line2Polyline失败！");
                LogManager.Instance.LogInfo(ex.Message);
            }
        }

        /// <summary>
        /// 建筑房间号文字
        /// </summary>
        /// <param name="layerName"></param>
        [CommandMethod(nameof(DBTextLabel_JZ))]
        public static void DBTextLabel_JZ()
        {
            try
            {
                using var tr = new DBTrans();
                TextFontsStyleHelper.TextStyleAndLayerInfo(tr, VariableDictionary.btnBlockLayer, Convert.ToInt16(VariableDictionary.layerColorIndex), "tJText");
                //创建文字与文字属性
                DBText text = new DBText()
                {
                    TextStyleId = tr.TextStyleTable["tJText"],
                    TextString = VariableDictionary.btnFileName,//字体的内容
                    Height = 350,//字体的高度
                    WidthFactor = 0.7,//字体的宽度因子
                    ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex),//字体的颜色
                    Layer = VariableDictionary.btnBlockLayer,//字体的图层
                };

                var dbTextEntityObj = tr.CurrentSpace.AddEntity(text);//写入当前空间
                var startPoint = new Point3d(0, 0, 0);
                double tempAngle = 0;//角度
                var entityBBText = new JigEx((mpw, _) =>
                {
                    text.Move(startPoint, mpw);
                    startPoint = mpw;
                    if (VariableDictionary.entityRotateAngle == tempAngle)
                    {
                        return;
                    }
                    else if (VariableDictionary.entityRotateAngle != tempAngle)
                    {
                        text.Rotation(center: mpw, 0);
                        tempAngle = VariableDictionary.entityRotateAngle;
                        text.Rotation(center: mpw, tempAngle);
                    }
                });
                entityBBText.DatabaseEntityDraw(wd => wd.Geometry.Draw(text));
                entityBBText.SetOptions(msg: "\n指定插入点");
                //entityBlock.SetOptions(startPoint, msg: "\n指定插入点");这个startpoint，是有个参考线在里面，用于托拽时的辅助；
                var endPoint = Env.Editor.Drag(entityBBText);
                if (endPoint.Status != PromptStatus.OK)
                    tr.Abort();
                Env.Editor.Redraw();//重新刷新
                tr.Commit();
            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("建筑房间号文字失败！");
                LogManager.Instance.LogInfo(ex.Message);
            }
        }

        /// <summary>
        /// 建筑专业用鼠标画矩形
        /// </summary>
        /// <param name="layerName"></param>
        [CommandMethod(nameof(Rec2PolyLine_2))]
        public static void Rec2PolyLine_2()
        {
            try
            {
                var layerName = VariableDictionary.btnBlockLayer;
                Int16 layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex == null ? VariableDictionary.layerColorIndex : 0); // 设置图层颜色索引
                using var tr = new DBTrans();//开启事务
                // 检查图层是否存在，如果不存在则创建
                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；

                // 获取矩形左下角点（起始点）  
                var userPoint1 = Env.Editor.GetPoint("\n指定矩形的左下角点：");
                if (userPoint1.Status != PromptStatus.OK)
                    return;
                var basePoint = userPoint1.Value.Wcs2Ucs().Z20();

                // 通过鼠标拖动确定矩形宽度（X 方向增量），高度固定为300  
                using var jig = new JigEx((mpw, queue) =>
                {
                    // 获取当前动态点  
                    var dynamicPoint = mpw.Z20();
                    // 计算宽度（可为负值，表示向左延伸）  
                    double width = dynamicPoint.X - basePoint.X;
                    // 构造矩形的多段线  
                    Polyline polyline1 = new Polyline()
                    {
                        Layer = layerName,
                        ColorIndex = layerColorIndex
                    };
                    // 四个顶点顺序：左下、右下、右上、左上  
                    polyline1.AddVertexAt(0, new Point2d(basePoint.X, basePoint.Y), 0, 0, 0);
                    polyline1.AddVertexAt(1, new Point2d(basePoint.X + width, basePoint.Y), 0, 0, 0);
                    polyline1.AddVertexAt(2, new Point2d(basePoint.X + width, basePoint.Y + 300), 0, 0, 0);
                    polyline1.AddVertexAt(3, new Point2d(basePoint.X, basePoint.Y + 300), 0, 0, 0);
                    polyline1.Closed = true;

                    queue.Enqueue(polyline1);
                });

                jig.SetOptions(basePoint, msg: "\n指定矩形宽度的第二个点（右键结束）：");
                var userPoint2 = Env.Editor.Drag(jig);
                if (userPoint2.Status != PromptStatus.OK)
                {
                    LogManager.Instance.LogInfo("\n未指定矩形宽度，操作取消。");
                    return;
                }
                // 通过最后一个动态点确定最终宽度  
                double finalWidth = jig.MousePointWcsLast.X - basePoint.X;

                // 创建最终的矩形多段线  
                Polyline finalPolyline = new Polyline();
                finalPolyline.AddVertexAt(0, new Point2d(basePoint.X, basePoint.Y), 0, 0, 0);
                finalPolyline.AddVertexAt(1, new Point2d(basePoint.X + finalWidth, basePoint.Y), 0, 0, 0);
                finalPolyline.AddVertexAt(2, new Point2d(basePoint.X + finalWidth, basePoint.Y + 300), 0, 0, 0);
                finalPolyline.AddVertexAt(3, new Point2d(basePoint.X, basePoint.Y + 300), 0, 0, 0);
                finalPolyline.Closed = true;
                finalPolyline.Layer = layerName;
                finalPolyline.ColorIndex = 231;

                // 添加矩形并获取其ID  
                var polylineId = tr.CurrentSpace.AddEntity(finalPolyline);

                try
                {
                    // 创建填充图案  
                    Hatch hatch = new Hatch();
                    hatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI31"); // 设置填充图案为 ANSI31
                    hatch.PatternScale = 200; // 设置填充图案比例为 200  
                    hatch.Layer = layerName; // 设置图层  
                    hatch.ColorIndex = 231; // 设置填充色号  
                    hatch.PatternAngle = 0; // 设置填充角度  
                    hatch.Normal = Vector3d.ZAxis;
                    ObjectIdCollection boundaryIds = new ObjectIdCollection(); // 创建边界集合  
                    boundaryIds.Add(polylineId);
                    hatch.AppendLoop(HatchLoopTypes.External, boundaryIds); // 添加外部环  
                    hatch.EvaluateHatch(true); // 强制计算填充图案  
                    var hatchId = tr.CurrentSpace.AddEntity(hatch); // 将填充添加到模型空间  

                    Env.Editor.Redraw();  // 强制刷新视图  
                    LogManager.Instance.LogInfo("\n矩形绘制完成并闭合，填充图案已添加。");
                }
                catch (System.Exception ex)
                {
                    LogManager.Instance.LogInfo($"\n创建填充时出错: {ex.Message}");
                }

                Env.Editor.Redraw(); // 强制刷新视图  
                if (VariableDictionary.dimString_JZ_宽 != null && VariableDictionary.dimString_JZ_深 != null)
                    DDimLinear(tr, Convert.ToString(VariableDictionary.dimString_JZ_深), Convert.ToString(VariableDictionary.dimString_JZ_宽), Convert.ToInt16(VariableDictionary.layerColorIndex));
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n矩形绘制失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 建筑专业用鼠标画矩形
        /// </summary>
        /// <param name="layerName"></param>
        [CommandMethod(nameof(Rec2PolyLine_3))]
        public static void Rec2PolyLine_3()
        {
            try
            {
                var layerName = VariableDictionary.btnBlockLayer;
                var textBoxScale = VariableDictionary.textBoxScale == null ? VariableDictionary.textBoxScale : 1;
                // 计算矢量差（拖动时基于参考点的偏移量）  
                var delta = new Vector3d(0, 0, 0);
                Int16 layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex == null ? VariableDictionary.layerColorIndex : 0); // 设置图层颜色索引
                using var tr = new DBTrans();//开启事务
                // 检查图层是否存在，如果不存在则创建
                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；

                // 获取参考点（可以视为左下角，但后续根据方向调整）  
                var userPoint1 = Env.Editor.GetPoint("\n指定矩形的参考点：");
                if (userPoint1.Status != PromptStatus.OK)
                    return;
                var basePoint = userPoint1.Value.Wcs2Ucs().Z20();
                var dimWide = Convert.ToDouble(VariableDictionary.dimString_JZ_宽);
                double? userDistance = null; // 保存用户通过命令行输入的距离数值  

                // 使用 JigEx 动态预览矩形，同时允许用户输入数值作为距离值  
                using var jig = new JigEx((mpw, queue) =>
                {
                    // 获取当前动态点  
                    var dynamicPoint = mpw.Z20();
                    // 计算矢量差（拖动时基于参考点的偏移量）  
                    delta = dynamicPoint - basePoint;
                    // 判断以哪个方向为主：若 Y 差绝对值大于 X 差，视为竖直模式，否则为水平模式  
                    bool verticalMode = (System.Math.Abs(delta.Y) > System.Math.Abs(delta.X));

                    // 如果用户已输入数值，则采用该数值作为距离；  
                    // 否则直接以鼠标相对于参考点的偏移作为距离  
                    double distanceValue;
                    if (userDistance.HasValue)
                    {
                        if (verticalMode)
                        {
                            // 当处于竖直模式时，同时判断鼠标在参考点上方或下方  
                            distanceValue = (delta.Y >= 0 ? userDistance.Value : -userDistance.Value);
                        }
                        else
                        {
                            // 水平模式：判断鼠标是在参考点右侧还是左侧  
                            distanceValue = (delta.X >= 0 ? userDistance.Value : -userDistance.Value);
                        }
                    }
                    else
                    {
                        distanceValue = verticalMode ? delta.Y : delta.X;
                    }

                    // 构造预览矩形多段线  
                    Polyline polyline1 = new Polyline()
                    {
                        Layer = layerName,
                        ColorIndex = layerColorIndex
                    };

                    if (verticalMode)
                    {
                        // 竖直模式：宽固定 300，矩形高度由 distanceValue 决定，  
                        // 但这里还要判断鼠标在水平方向（相对于参考点）是否位于右侧或左侧，  
                        // 如果在左侧，则矩形宽度应反向（向左延伸）  
                        if (delta.X >= 0)
                        {
                            // 鼠标在右侧：宽向右，顶点顺序：参考点、向右延伸、上/下延伸、垂直延伸回到参考点的X  
                            polyline1.AddVertexAt(0, new Point2d(basePoint.X, basePoint.Y), 0, 0, 0);
                            polyline1.AddVertexAt(1, new Point2d(basePoint.X + dimWide, basePoint.Y), 0, 0, 0);
                            polyline1.AddVertexAt(2, new Point2d(basePoint.X + dimWide, basePoint.Y + distanceValue), 0, 0, 0);
                            polyline1.AddVertexAt(3, new Point2d(basePoint.X, basePoint.Y + distanceValue), 0, 0, 0);
                        }
                        else
                        {
                            // 鼠标在左侧：宽向左，顶点顺序：参考点、向左延伸、上/下延伸、垂直返回  
                            polyline1.AddVertexAt(0, new Point2d(basePoint.X, basePoint.Y), 0, 0, 0);
                            polyline1.AddVertexAt(1, new Point2d(basePoint.X - dimWide, basePoint.Y), 0, 0, 0);
                            polyline1.AddVertexAt(2, new Point2d(basePoint.X - dimWide, basePoint.Y + distanceValue), 0, 0, 0);
                            polyline1.AddVertexAt(3, new Point2d(basePoint.X, basePoint.Y + distanceValue), 0, 0, 0);
                        }
                    }
                    else
                    {
                        // 水平模式：高固定 300，由 distanceValue 决定矩形宽度方向，  
                        // 同时判断鼠标在垂直方向上是位于参考点的上方还是下方  
                        if (delta.Y >= 0)
                        {
                            // 鼠标在上方：高向上  
                            polyline1.AddVertexAt(0, new Point2d(basePoint.X, basePoint.Y), 0, 0, 0);
                            polyline1.AddVertexAt(1, new Point2d(basePoint.X + distanceValue, basePoint.Y), 0, 0, 0);
                            polyline1.AddVertexAt(2, new Point2d(basePoint.X + distanceValue, basePoint.Y + dimWide), 0, 0, 0);
                            polyline1.AddVertexAt(3, new Point2d(basePoint.X, basePoint.Y + dimWide), 0, 0, 0);
                        }
                        else
                        {
                            // 鼠标在下方：高向下，顶点顺序调整后确保矩形延伸方向正确  
                            polyline1.AddVertexAt(0, new Point2d(basePoint.X, basePoint.Y), 0, 0, 0);
                            polyline1.AddVertexAt(1, new Point2d(basePoint.X + distanceValue, basePoint.Y), 0, 0, 0);
                            polyline1.AddVertexAt(2, new Point2d(basePoint.X + distanceValue, basePoint.Y - dimWide), 0, 0, 0);
                            polyline1.AddVertexAt(3, new Point2d(basePoint.X, basePoint.Y - dimWide), 0, 0, 0);
                        }
                    }
                    polyline1.Closed = true;
                    queue.Enqueue(polyline1);
                });

                // 提示信息：用户拖动或输入数值后直接回车  
                jig.SetOptions(basePoint, msg: "\n指定矩形第二点（或输入距离值后回车）：");
                var userResponse = Env.Editor.Drag(jig);
                if (userResponse.Status != PromptStatus.OK)
                {
                    LogManager.Instance.LogInfo("\n未指定矩形尺寸，操作取消。");
                    return;
                }

                // 判断是否有用户输入的数值（假设 jig.DistanceEntered 属性可得输入距离，非 0 表示输入了数值）  
                //if (jig.DistanceEntered != 0)
                //    userDistance = jig.DistanceEntered;

                // 获取最终动态点（如果未输入数值则直接取鼠标位置）  
                var dynamicPointFinal = jig.MousePointWcsLast;
                var deltaFinal = dynamicPointFinal - basePoint;
                bool isVertical = (System.Math.Abs(deltaFinal.Y) > System.Math.Abs(deltaFinal.X));

                double finalDistance;
                if (userDistance.HasValue)
                {
                    finalDistance = isVertical ?
                                    (deltaFinal.Y >= 0 ? userDistance.Value : -userDistance.Value) :
                                    (deltaFinal.X >= 0 ? userDistance.Value : -userDistance.Value);
                }
                else
                {
                    finalDistance = isVertical ? deltaFinal.Y : deltaFinal.X;
                }

                // 根据鼠标最终位置判断方向，并计算矩形四个顶点  
                Point2d pt0, pt1, pt2, pt3;
                if (isVertical)
                {
                    if (delta.X >= 0)
                    {
                        // 鼠标在参考点右侧  
                        pt0 = new Point2d(basePoint.X, basePoint.Y);
                        pt1 = new Point2d(basePoint.X + dimWide, basePoint.Y);
                        pt2 = new Point2d(basePoint.X + dimWide, basePoint.Y + finalDistance);
                        pt3 = new Point2d(basePoint.X, basePoint.Y + finalDistance);
                    }
                    else
                    {
                        // 鼠标在参考点左侧  
                        pt0 = new Point2d(basePoint.X, basePoint.Y);
                        pt1 = new Point2d(basePoint.X - dimWide, basePoint.Y);
                        pt2 = new Point2d(basePoint.X - dimWide, basePoint.Y + finalDistance);
                        pt3 = new Point2d(basePoint.X, basePoint.Y + finalDistance);
                    }
                }
                else
                {
                    if (delta.Y >= 0)
                    {
                        // 鼠标在参考点上方  
                        pt0 = new Point2d(basePoint.X, basePoint.Y);
                        pt1 = new Point2d(basePoint.X + finalDistance, basePoint.Y);
                        pt2 = new Point2d(basePoint.X + finalDistance, basePoint.Y + dimWide);
                        pt3 = new Point2d(basePoint.X, basePoint.Y + dimWide);
                    }
                    else
                    {
                        // 鼠标在参考点下方  
                        pt0 = new Point2d(basePoint.X, basePoint.Y);
                        pt1 = new Point2d(basePoint.X + finalDistance, basePoint.Y);
                        pt2 = new Point2d(basePoint.X + finalDistance, basePoint.Y - dimWide);
                        pt3 = new Point2d(basePoint.X, basePoint.Y - dimWide);
                    }
                }

                // 创建最终矩形闭合多段线  
                Polyline finalPolyline = new Polyline();
                finalPolyline.AddVertexAt(0, new Point2d(pt0.X, pt0.Y), 0, 0, 0);
                finalPolyline.AddVertexAt(1, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                finalPolyline.AddVertexAt(2, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                finalPolyline.AddVertexAt(3, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                finalPolyline.Closed = true;
                finalPolyline.Layer = layerName;
                finalPolyline.ColorIndex = VariableDictionary.layerColorIndex;

                // 添加矩形并获取其 ID  
                var polylineId = tr.CurrentSpace.AddEntity(finalPolyline);

                try
                {
                    // 创建填充图案  
                    Hatch hatch = new Hatch();
                    hatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI31");
                    hatch.PatternScale = 100;
                    hatch.Layer = layerName;
                    hatch.ColorIndex = VariableDictionary.layerColorIndex;
                    hatch.PatternAngle = 0;
                    hatch.Normal = Vector3d.ZAxis;
                    ObjectIdCollection boundaryIds = new ObjectIdCollection();
                    boundaryIds.Add(polylineId);
                    hatch.AppendLoop(HatchLoopTypes.External, boundaryIds);
                    hatch.EvaluateHatch(true);
                    var hatchId = tr.CurrentSpace.AddEntity(hatch);

                    Env.Editor.Redraw();
                    LogManager.Instance.LogInfo("\n矩形绘制完成并闭合，填充图案已添加。");
                }
                catch (System.Exception ex)
                {
                    LogManager.Instance.LogInfo($"\n创建填充时出错: {ex.Message}");
                }
                if (VariableDictionary.dimString_JZ_宽 != null && VariableDictionary.dimString_JZ_深 != null)
                    DDimLinear(tr, Convert.ToString(VariableDictionary.dimString_JZ_深), Convert.ToString(VariableDictionary.dimString_JZ_宽), Convert.ToInt16(VariableDictionary.layerColorIndex));

                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogInfo($"\n矩形绘制失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 建筑画防撞板
        /// </summary>
        [CommandMethod(nameof(ParallelLines))]
        public static void ParallelLines()
        {
            try
            {
                var layerName = VariableDictionary.btnBlockLayer; // 图层
                var move = VariableDictionary.textbox_Gap; // 获取用户输入的移动距离

                Int16 layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex == null ? VariableDictionary.layerColorIndex : 64); // 设置图层颜色索引
                using var tr = new DBTrans();//开启事务
                // 检查图层是否存在，如果不存在则创建
                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);//添加图层；

                // 获取用户输入的第一点和第二点
                var userPoint1 = Env.Editor.GetPoint("\n请指定第一点");
                if (userPoint1.Status != PromptStatus.OK) return;
                var UcsUserPoint1 = userPoint1.Value.Wcs2Ucs().Z20(); // 转换为UCS坐标

                var userPoint2 = Env.Editor.GetPoint("\n请指定第二点（确定平行线距离）");
                if (userPoint2.Status != PromptStatus.OK) return;
                var UcsUserPoint2 = userPoint2.Value.Wcs2Ucs().Z20(); // 转换为UCS坐标

                // 计算两条平行线之间的偏移向量
                var offset = UcsUserPoint2 - UcsUserPoint1;

                // 计算移动后的偏移量
                var offsetDirection = offset.GetNormal(); // 获取偏移向量的单位方向向量
                var moveOffset = offsetDirection * move; // 计算移动的距离向量

                using var parallelLine = new JigEx((mpw, queue) =>
                {
                    var UcsUserPointEnd = mpw.Z20();

                    // 绘制第一条线
                    Polyline polyline1 = new Polyline();
                    polyline1.AddVertexAt(0, new Point2d(UcsUserPoint1.X + moveOffset.Value.X, UcsUserPoint1.Y + moveOffset.Value.Y), 0, 0, 0); // 起点向第二条线移动
                    polyline1.AddVertexAt(1, new Point2d(UcsUserPointEnd.X + moveOffset.Value.X, UcsUserPointEnd.Y + moveOffset.Value.Y), 0, 0, 0); // 终点向第二条线移动
                    polyline1.Closed = false;
                    polyline1.Layer = layerName; // 设置线条图层
                    polyline1.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                    polyline1.SetStartWidthAt(0, 50);
                    polyline1.SetEndWidthAt(0, 50);
                    queue.Enqueue(polyline1);

                    // 绘制第二条平行线
                    Polyline polyline2 = new Polyline();
                    polyline2.AddVertexAt(0, new Point2d(UcsUserPoint2.X - moveOffset.Value.X, UcsUserPoint2.Y - moveOffset.Value.Y), 0, 0, 0); // 起点向第一条线移动
                    polyline2.AddVertexAt(1, new Point2d(UcsUserPointEnd.X + offset.X - moveOffset.Value.X, UcsUserPointEnd.Y + offset.Y - moveOffset.Value.Y), 0, 0, 0); // 终点向第一条线移动
                    polyline2.Closed = false;
                    polyline2.Layer = layerName; // 设置线条图层
                    polyline2.ColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                    polyline2.SetStartWidthAt(0, 50);
                    polyline2.SetEndWidthAt(0, 50);
                    queue.Enqueue(polyline2);
                });

                parallelLine.SetOptions(UcsUserPoint1, msg: "\n请指定终点");
                var userPointEnd = Env.Editor.Drag(parallelLine);
                if (userPointEnd.Status != PromptStatus.OK) return;
                var polyLineEntityObj = tr.CurrentSpace.AddEntity(parallelLine.Entities);
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("建筑画防撞板失败！");
                LogManager.Instance.LogInfo(ex.Message);
            }
        }
        #endregion

        #region 计算面积方法


        /// <summary>
        /// 指定点计算点内面积（支持动态预览 + Z 撤销）
        /// </summary>
        [CommandMethod("AreaByPoints")]
        public void AreaByPoints()
        {
            try
            {
                // 1) 设置结果图层名称和颜色
                string layerName = "房屋面积";
                short layerColorIndex = 2;
                var textBoxScale = VariableDictionary.textBoxScale == null ? VariableDictionary.textBoxScale : 100;

                // 2) 用于存储用户依次确认的点（二维点，统一按 UCS 保存）
                List<Point2d> points = new List<Point2d>();

                // 3) 开启事务
                using var tr = new DBTrans();

                // 4) 确保结果图层存在，不存在则创建
                LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, layerColorIndex);

                // 5) 获取第一个点（首点不支持撤销）
                var firstPointRes = Env.Editor.GetPoint("\n指定多边形的第一个点：");
                if (firstPointRes.Status != PromptStatus.OK)
                {
                    LogManager.Instance.LogInfo("\n未指定第一个点，操作取消。");
                    return;
                }
                bool oldOrtho = false; // 保存原正交状态
                short oldDucs = 0; // 保存原 DUCS 状态
                try // 主流程
                {
                    oldOrtho = Env.OrthoMode; // 记录进入前正交状态
                    try { oldDucs = Convert.ToInt16(Application.GetSystemVariable("DUCS")); } catch { oldDucs = 0; } // 记录进入前 DUCS
                    Env.OrthoMode = true; // 选点期间强制正交，仅允许 X/Y 方向
                    try { Application.SetSystemVariable("DUCS", 0); } catch { } // 关闭动态 UCS，避免干扰正交方向

                    // ===== 你原有 AreaByPoints 逻辑保持不变（这里省略） =====
                    // 重点是 while 循环里继续使用 AreaPolylinePreviewJig 即可
                }
                finally // 无论成功失败都恢复环境
                {
                    try { Env.OrthoMode = oldOrtho; } catch { } // 恢复原正交状态
                    try { Application.SetSystemVariable("DUCS", oldDucs); } catch { } // 恢复原 DUCS 状态
                }
                //Env.OrthoMode = true; // 开启正交模式，方便画矩形等规则图形

                // 6) 将首点转为 UCS 后存入点集
                var firstUcs = firstPointRes.Value.Wcs2Ucs().Z20();
                points.Add(new Point2d(firstUcs.X, firstUcs.Y));

                // 7) 循环取点：支持动态预览、Z 撤销、右键结束
                while (true)
                {
                    // 7.1) 创建预览 Jig（每轮使用当前已确认点）
                    var jig = new AreaPolylinePreviewJig(points, layerName, layerColorIndex);

                    // 7.2) 启动拖拽交互
                    var dragRes = Env.Editor.Drag(jig);

                    // 7.3) 正常取点：将当前点加入点集
                    if (dragRes.Status == PromptStatus.OK)
                    {
                        var ucsPt = jig.CurrentPointWcs.Wcs2Ucs().Z20();
                        points.Add(new Point2d(ucsPt.X, ucsPt.Y));
                        Env.Editor.WriteMessage($"\n已添加点，当前点数：{points.Count}");
                        Env.Editor.Redraw();
                        continue;
                    }

                    // 7.4) 关键字撤销：
                    //     - 某些环境会返回 PromptStatus.Keyword
                    //     - 某些环境会返回 Cancel，但 jig.LastKeyword 仍是 "Z"
                    bool isUndo =
                        (dragRes.Status == PromptStatus.Keyword && string.Equals(dragRes.StringResult, "Z", StringComparison.OrdinalIgnoreCase)) ||
                        (dragRes.Status == PromptStatus.Cancel && string.Equals(jig.LastKeyword, "Z", StringComparison.OrdinalIgnoreCase));

                    if (isUndo)
                    {
                        if (points.Count > 1)
                        {
                            points.RemoveAt(points.Count - 1);
                            Env.Editor.WriteMessage($"\n已撤销上一个点，当前点数：{points.Count}");
                        }
                        else
                        {
                            Env.Editor.WriteMessage("\n当前只有首点，无法继续撤销。");
                        }

                        Env.Editor.Redraw();
                        continue;
                    }

                    // 7.5) 右键/回车/Esc/其他状态：结束选点
                    break;
                }

                // 8) 点数不足 3 个无法形成闭合多边形
                if (points.Count < 3)
                {
                    LogManager.Instance.LogInfo("\n至少需要3个点来计算面积！");
                    return;
                }

                // 9) 计算多边形面积
                double area = CalculatePolygonArea(points);

                // 10) 计算多边形质心（用于放置面积文字）
                Point2d centroid = CalculateCentroid(points);

                // 11) 创建面积文字（默认放在质心）
                var text = new DBText();
                TextFontsStyleHelper.TextStyleAndLayerInfo(tr, layerName, layerColorIndex, "tJText");
                text.TextString = $"{area:F2}";
                text.Height = 3 * textBoxScale;
                text.TextStyleId = tr.TextStyleTable["tJText"];
                text.Layer = layerName;
                text.ColorIndex = layerColorIndex;
                text.Position = new Point3d(centroid.X, centroid.Y, 0);


                // 12) 允许用户拖拽文字到目标位置
                using var moveText = new JigEx((mpw, _) =>
                {
                    text.Position = mpw.Z20();
                });
                moveText.DatabaseEntityDraw(wd => wd.Geometry.Draw(text));
                moveText.SetOptions(msg: "\n请指定面积文字位置：");
                var dragTextRes = Env.Editor.Drag(moveText);

                // 13) 用户确认拖拽后更新文字位置
                if (dragTextRes.Status == PromptStatus.OK)
                {
                    text.Position = moveText.MousePointWcsLast;
                }

                // 14) 写入文字并提交事务
                tr.CurrentSpace.AddEntity(text);
                tr.Commit();
                Env.Editor.Redraw();

                // 15) 输出计算结果
                LogManager.Instance.LogInfo($"\n多边形面积为: {area:F2}");
            }
            catch (Exception ex)
            {
                // 异常日志
                LogManager.Instance.LogInfo($"\n面积计算失败：{ex.Message}");
            }
        }


        /// <summary>
        /// 计算多边形面积（Shoelace公式）
        /// </summary>
        /// <param name="pts"></param>
        /// <returns></returns>
        private double CalculatePolygonArea(List<Point2d> pts)
        {
            double area = 0;
            int n = pts.Count;
            for (int i = 0; i < n; i++)
            {
                Point2d p1 = pts[i];
                Point2d p2 = pts[(i + 1) % n];
                area += (p1.X * p2.Y - p2.X * p1.Y);
            }
            return System.Math.Abs(area) / 2.0 / 100 / 100 / 100;
        }

        /// <summary>
        /// 计算多边形质心
        /// </summary>
        /// <param name="pts"></param>
        /// <returns></returns>
        private Point2d CalculateCentroid(List<Point2d> pts)
        {
            double cx = 0, cy = 0;
            double area = 0;
            int n = pts.Count;
            for (int i = 0; i < n; i++)
            {
                Point2d p0 = pts[i];
                Point2d p1 = pts[(i + 1) % n];
                double cross = p0.X * p1.Y - p1.X * p0.Y;
                cx += (p0.X + p1.X) * cross;
                cy += (p0.Y + p1.Y) * cross;
                area += cross;
            }
            area = area / 2.0;
            cx = cx / (6 * area);
            cy = cy / (6 * area);
            return new Point2d(cx, cy);
        }

        #endregion
        #region 工艺相关方法
        /// <summary>
        /// 工艺与暖通标注文字
        /// </summary>
        /// <param name="layerName"></param>
        [CommandMethod(nameof(DBTextLabel))]
        public static void DBTextLabel()
        {
            try
            {
                using var tr = new DBTrans();
                TextFontsStyleHelper.TextStyleAndLayerInfo(tr, VariableDictionary.btnBlockLayer, Convert.ToInt16(VariableDictionary.layerColorIndex), "tJText");
                var layerColorIndex = Convert.ToInt16(VariableDictionary.layerColorIndex);
                double uiScale = VariableDictionary.textBoxScale;
                if (double.IsNaN(uiScale) || uiScale <= 0)
                {
                    try { uiScale = AutoCadHelper.GetScale(true); } catch { uiScale = 1.0; }
                }
                if (uiScale <= 0) uiScale = 1.0;
                //double textHeight = ResolveLeaderTextHeight(3.0);
                //double arrowSize = ResolveLeaderArrowSize(2.0);

                string targetLayer = LayerControlHelper.GetOrCreateTargetLayer(
                    tr,
                    VariableDictionary.layerName ?? VariableDictionary.btnBlockLayer,
                    layerColorIndex);
                //创建文字与文字属性
                DBText text = new DBText()
                {
                    TextStyleId = tr.TextStyleTable["tJText"],
                    TextString = VariableDictionary.btnFileName,//字体的内容
                    Height = 3.5 * uiScale,//字体的高度
                    WidthFactor = 0.8,//字体的宽度因子
                    ColorIndex = layerColorIndex,//字体的颜色
                    Layer = targetLayer,//字体的图层
                };
                if (VariableDictionary.btnBlockLayer.Contains("工艺"))
                {
                    text.Height = 2.5 * uiScale;
                }
                var dbTextEntityObj = tr.CurrentSpace.AddEntity(text);//写入当前空间
                var startPoint = new Point3d(0, 0, 0);
                double tempAngle = 0;//角度
                var entityBBText = new JigEx((mpw, _) =>
                {
                    text.Move(startPoint, mpw);
                    text.Layer = targetLayer;
                    startPoint = mpw;
                    if (VariableDictionary.entityRotateAngle == tempAngle)
                    {
                        return;
                    }
                    else if (VariableDictionary.entityRotateAngle != tempAngle)
                    {
                        text.Rotation(center: mpw, 0);
                        tempAngle = VariableDictionary.entityRotateAngle;
                        text.Rotation(center: mpw, tempAngle);
                    }
                });
                entityBBText.DatabaseEntityDraw(wd => wd.Geometry.Draw(text));
                entityBBText.SetOptions(msg: "\n指定插入点");

                //entityBlock.SetOptions(startPoint, msg: "\n指定插入点");这个startpoint，是有个参考线在里面，用于托拽时的辅助；
                var endPoint = Env.Editor.Drag(entityBBText);
                if (endPoint.Status != PromptStatus.OK)
                    tr.Abort();
                tr.Commit();
                Env.Editor.Redraw();//重新刷新
            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("工艺与暖通标注文字失败！");
                LogManager.Instance.LogInfo($"\n错误信息：{ex.Message}");
            }
        }


        /// <summary>
        /// 创建传递窗
        /// </summary>
        /// <param name="doc"> </param>
        /// <param name="coreLen"> </param>
        /// <param name="coreWid"></param>
        /// <param name="coreHeightVal"></param>
        /// <param name="typeVal"></param>
        /// <param name="paramVal"></param>
        /// <param name="selectedRadioNames"></param>
        /// <exception cref="ArgumentNullException"></exception>

        public static void GenerateAndInsertTransferWindow(Document doc, double coreLen, double coreWid, double coreHeightVal, string typeVal, string paramVal, IEnumerable<string> selectedRadioNames)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            var selected = (selectedRadioNames ?? Enumerable.Empty<string>()).ToList();

            // 解析主侧逻辑：通常取倒数第二段（例如 XXX_上_左 -> 上）
            IEnumerable<string> GetPrimarySides(IEnumerable<string> names)
            {
                foreach (var n in names)
                {
                    if (string.IsNullOrWhiteSpace(n)) continue;
                    var parts = n.Split('_');
                    if (parts.Length >= 2)
                        yield return parts[parts.Length - 2];
                }
            }

            // 获取所有选中的主方向（用于计算外框偏移）
            var primarySides = GetPrimarySides(selected).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            bool topSel = primarySides.Contains("上");
            bool bottomSel = primarySides.Contains("下");
            bool leftSel = primarySides.Contains("左");
            bool rightSel = primarySides.Contains("右");

            // 根据是否有选中该侧来决定外扩距离（有门的一侧外扩30，无门一侧外扩100）
            double topOff = topSel ? 30.0 : 100.0;
            double bottomOff = bottomSel ? 30.0 : 100.0;
            double leftOff = leftSel ? 30.0 : 100.0;
            double rightOff = rightSel ? 30.0 : 100.0;

            // 计算外框总尺寸
            double outerW = coreLen + leftOff + rightOff;
            double outerH = coreWid + topOff + bottomOff;

            // 构造标签文本
            string code = typeVal?.Contains("PB", StringComparison.OrdinalIgnoreCase) == true ? "PB" :
                          typeVal?.Contains("JC", StringComparison.OrdinalIgnoreCase) == true ? "JC" : typeVal ?? string.Empty;
            string label = paramVal == "对开式" ? code : $"{code}\n{paramVal}";
            if (VariableDictionary.radioButton)
            {
                paramVal = "双层";
                label = $"{code}\n{paramVal}";
            }
            try
            {
                using (doc.LockDocument())
                using (var tr = new DBTrans())
                {
                    // 确保图层和文字样式存在
                    string targetLayer = LayerControlHelper.GetOrCreateTargetLayer(tr, "WALL-PARAPET", 3);

                    // 确保文字样式存在（DASH 里会引用）              
                    TextFontsStyleHelper.EnsureTextStyle(tr, "tJText");
                    // 1. 绘制外框和核心矩形
                    var outerPl = CreateRectPolyline(outerW, outerH, 0, 0); // 外框 (插入点为0,0)
                    outerPl.Layer = targetLayer;
                    outerPl.ColorIndex = VariableDictionary.layerColorIndex;
                    var outerId = tr.CurrentSpace.AddEntity(outerPl);

                    var corePl = CreateRectPolyline(coreLen, coreWid, leftOff, -topOff); // 内框 (考虑偏移)
                    corePl.Layer = targetLayer;
                    corePl.ColorIndex = VariableDictionary.layerColorIndex;
                    var coreId = tr.CurrentSpace.AddEntity(corePl);

                    // 1.5 添加新的尺寸文字 [长:L,宽:W,高:H] 到核心内框左上角
                    var dimTextStr = $"内径: 长:{coreLen},宽:{coreWid},高:{coreHeightVal}";
                    var dimText = new DBText
                    {
                        TextString = dimTextStr,
                        Height = 30,
                        Position = new Point3d(leftOff, -topOff, 0),
                        Layer = targetLayer,
                        ColorIndex = VariableDictionary.layerColorIndex,
                        TextStyleId = tr.TextStyleTable["tJText"]
                    };
                    var dimTextId = tr.CurrentSpace.AddEntity(dimText);

                    // 收集所有创建的实体ID，用于后续 Jig 整体拖拽
                    var createdIds = new List<ObjectId> { outerId, coreId, dimTextId };

                    // 2. 解析用户选中的所有位置代码（格式如 "上_左"）
                    var activeCodes = new HashSet<string>();
                    foreach (var name in selected)
                    {
                        if (string.IsNullOrWhiteSpace(name)) continue;
                        var parts = name.Split('_');
                        if (parts.Length >= 2)
                        {
                            string a = parts[parts.Length - 2].Trim();
                            string b = parts[parts.Length - 1].Trim();
                            activeCodes.Add($"{a}_{b}");
                        }
                    }

                    // 3. 【核心逻辑】根据尺寸自动补全双开门指示线
                    // 复制一份作为最终绘制列表
                    var codesToDraw = new HashSet<string>(activeCodes);

                    // 规则A：如果长度 coreLen >= 1000，上/下侧自动变成双开（互补左右）
                    if (coreLen >= 1000)
                    {
                        if (activeCodes.Contains("上_左")) codesToDraw.Add("上_右");
                        if (activeCodes.Contains("上_右")) codesToDraw.Add("上_左");

                        if (activeCodes.Contains("下_左")) codesToDraw.Add("下_右");
                        if (activeCodes.Contains("下_右")) codesToDraw.Add("下_左");
                    }

                    // 规则B：如果宽度 coreWid >= 1000，左/右侧自动变成双开（互补上下）
                    if (coreWid >= 1000)
                    {
                        if (activeCodes.Contains("左_上")) codesToDraw.Add("左_下");
                        if (activeCodes.Contains("左_下")) codesToDraw.Add("左_上");

                        if (activeCodes.Contains("右_上")) codesToDraw.Add("右_下");
                        if (activeCodes.Contains("右_下")) codesToDraw.Add("右_上");
                    }

                    // 4. 遍历所有代码生成指示线
                    foreach (var codePair in codesToDraw)
                    {
                        Point3d basePt = Point3d.Origin;
                        double angleDeg = 0;
                        double guideLen = 0;

                        // 外框角点
                        var outerLeftTop = new Point3d(0, 0, 0);
                        var outerRightTop = new Point3d(outerW, 0, 0);
                        var outerLeftBottom = new Point3d(0, -outerH, 0);
                        var outerRightBottom = new Point3d(outerW, -outerH, 0);

                        // 内框角点
                        var innerLeftTop = new Point3d(leftOff, -topOff, 0);
                        var innerRightTop = new Point3d(leftOff + coreLen, -topOff, 0);
                        var innerLeftBottom = new Point3d(leftOff, -(topOff + coreWid), 0);
                        var innerRightBottom = new Point3d(leftOff + coreLen, -(topOff + coreWid), 0);

                        switch (codePair)
                        {
                            // 上侧：Y 用外框上角，X 取内框对应上角
                            case "上_左":
                                basePt = new Point3d(innerLeftTop.X, outerLeftTop.Y, 0);
                                angleDeg = 30.0;
                                guideLen = outerW * 2.0 / 3.0;
                                break;

                            case "上_右":
                                basePt = new Point3d(innerRightTop.X, outerRightTop.Y, 0);
                                angleDeg = 150.0;
                                guideLen = outerW * 2.0 / 3.0;
                                break;

                            // 下侧：Y 用外框下角，X 取内框对应下角
                            case "下_左":
                                basePt = new Point3d(innerLeftBottom.X, outerLeftBottom.Y, 0);
                                angleDeg = 330.0;
                                guideLen = outerW * 2.0 / 3.0;
                                break;

                            case "下_右":
                                basePt = new Point3d(innerRightBottom.X, outerRightBottom.Y, 0);
                                angleDeg = 210.0;
                                guideLen = outerW * 2.0 / 3.0;
                                break;

                            // 左侧：X 用外框左角，Y 取内框对应左角
                            case "左_上":
                                basePt = new Point3d(outerLeftTop.X, innerLeftTop.Y, 0);
                                angleDeg = 240.0;
                                guideLen = outerH * 2.0 / 3.0;
                                break;

                            case "左_下":
                                basePt = new Point3d(outerLeftBottom.X, innerLeftBottom.Y, 0);
                                angleDeg = 120.0;
                                guideLen = outerH * 2.0 / 3.0;
                                break;

                            // 右侧：X 用外框右角，Y 取内框对应右角
                            case "右_上":
                                basePt = new Point3d(outerRightTop.X, innerRightTop.Y, 0);
                                angleDeg = 300.0;
                                guideLen = outerH * 2.0 / 3.0;
                                break;

                            case "右_下":
                                basePt = new Point3d(outerRightBottom.X, innerRightBottom.Y, 0);
                                angleDeg = 60.0;
                                guideLen = outerH * 2.0 / 3.0;
                                break;

                            default:
                                continue;
                        }
                        // 根据基点、角度和长度计算指示线终点
                        var pEnd = basePt.PolarPoint(angleDeg, guideLen);
                        // 绘制指示线
                        var line = new Line(basePt, pEnd) { Layer = targetLayer };
                        line.ColorIndex = VariableDictionary.layerColorIndex;
                        // 设置线宽以增强可见性
                        var lineId = tr.CurrentSpace.AddEntity(line);
                        createdIds.Add(lineId);// 添加到创建的实体列表
                    }

                    // 5. 绘制文字 (居中)
                    var centerPt = new Point3d(leftOff + coreLen / 2.0, -(topOff + coreWid / 2.0), 0);
                    var mtext = new MText
                    {
                        TextStyleId = tr.TextStyleTable["tJText"],
                        Layer = targetLayer,
                        ColorIndex = VariableDictionary.layerColorIndex,
                        Contents = label,
                        Location = centerPt,
                        //TextHeight = Math.Max(50, Math.Min(coreLen, coreWid) / 3.0),
                        TextHeight = 2 * VariableDictionary.textBoxScale,
                        Attachment = AttachmentPoint.MiddleCenter
                    };
                    var mtextId = tr.CurrentSpace.AddEntity(mtext);
                    createdIds.Add(mtextId);

                    // 6. 启动 Jig 交互拖拽
                    DoJigDrag(tr, createdIds);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"生成传递窗出错：{ex.Message}", "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            // 局部函数：创建矩形多段线
            static Polyline CreateRectPolyline(double w, double h, double x0, double y0)
            {
                var pl = new Polyline(4);
                pl.AddVertexAt(0, new Point2d(x0, y0), 0, 0, 0);
                pl.AddVertexAt(1, new Point2d(x0 + w, y0), 0, 0, 0);
                pl.AddVertexAt(2, new Point2d(x0 + w, y0 - h), 0, 0, 0);
                pl.AddVertexAt(3, new Point2d(x0, y0 - h), 0, 0, 0);
                pl.Closed = true;

                return pl;
            }
            // 局部函数：执行整体拖拽
            static void DoJigDrag(DBTrans tr, List<ObjectId> entityIds)
            {
                try
                {
                    // 获取实体引用
                    var entities = entityIds.Select(id => tr.GetObject(id, OpenMode.ForWrite) as Entity)
                                            .Where(e => e != null).ToList();
                    if (entities.Count == 0) { tr.Commit(); return; }

                    var preview = entities.First();
                    Point3d lastPt = Point3d.Origin;
                    double lastAngle = VariableDictionary.entityRotateAngle;
                    double lastScale = VariableDictionary.textBoxScale > 0 ? VariableDictionary.textBoxScale : 1.0;

                    var jig = new JigEx((currPt, _) =>
                    {
                        // 1. 平移
                        var moveVec = currPt - lastPt;
                        if (moveVec.Length > 1e-9)
                        {
                            var mat = Matrix3d.Displacement(moveVec);
                            foreach (var e in entities) e.TransformBy(mat);
                            lastPt = currPt;
                        }

                        // 2. 旋转
                        if (Math.Abs(VariableDictionary.entityRotateAngle - lastAngle) > 1e-9)
                        {
                            double diff = VariableDictionary.entityRotateAngle - lastAngle;
                            var mat = Matrix3d.Rotation(diff, Vector3d.ZAxis, currPt);
                            foreach (var e in entities) e.TransformBy(mat);
                            lastAngle = VariableDictionary.entityRotateAngle;
                        }

                        // 3. 缩放
                        double currScale = VariableDictionary.textBoxScale > 0 ? VariableDictionary.textBoxScale : 1.0;
                        if (Math.Abs(currScale - lastScale) > 1e-9)
                        {
                            double factor = currScale / lastScale;
                            var mat = Matrix3d.Scaling(factor, currPt);
                            foreach (var e in entities) e.TransformBy(mat);
                            lastScale = currScale;
                        }
                    });

                    jig.DatabaseEntityDraw(wd => wd.Geometry.Draw(preview));
                    jig.SetOptions(msg: "\n指定插入点");

                    var res = Env.Editor.Drag(jig);
                    if (res.Status == PromptStatus.OK)
                    {
                        tr.Commit();
                    }
                    else
                    {
                        tr.Abort();
                    }
                    Env.Editor.Redraw();
                }
                catch
                {
                    tr.Abort();
                }
            }
        }

        #endregion

        #region 生成外轮廓方法


        #region 方法一：
        /// <summary>
        /// 智能轮廓
        /// </summary>
        [CommandMethod("SMARTOUTLINE")]
        public void SmartOutline()
        {
            // 获取CAD当前文档、数据库和命令行  
            Document doc = Application.DocumentManager.MdiActiveDocument; // 当前文档  
            Database db = doc.Database;                                   // 当前数据库  
            Editor ed = doc.Editor;                                       // 获取命令行交互对象  

            // 让用户选择对象  
            PromptSelectionOptions selOpt = new PromptSelectionOptions(); // 定义选择框选参数  
            selOpt.MessageForAdding = "\n请选择要生成轮廓的图元：";    // 提示消息  
            PromptSelectionResult selRes = ed.GetSelection(selOpt);      // 获取用户选择的对象  
            if (selRes.Status != PromptStatus.OK) return;                // 如果用户没有选择，退出方法  

            using (Transaction tr = db.TransactionManager.StartTransaction()) // 开始数据库事务  
            {
                // 打开模型空间用于添加新对象  
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable; // 获取块表  
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord; // 获取模型空间记录  

                // 1. 收集用户所选的全部曲线并爆炸块参照  
                List<Curve> allCurves = new List<Curve>(); // 存储所有选择的曲线对象  
                foreach (SelectedObject selectItem in selRes.Value) // 遍历所有选中的对象  
                {
                    if (selectItem == null) continue; // 判断是否为空选择  
                    Entity entity = tr.GetObject(selectItem.ObjectId, OpenMode.ForRead) as Entity; // 获取实体对象  
                    if (entity == null) continue; // 如果无法获取实体，则跳过  
                    CollectAllCurves(entity, tr, allCurves, 0); // 递归收集当前实体中的所有曲线  
                }
                if (allCurves.Count == 0) // 如果没有找到有效曲线  
                {
                    ed.WriteMessage("\n未找到可以处理的曲线或块对象。"); // 提示消息  
                    return; // 退出  
                }

                // 2. 分组。不连续的图元单独分组（通过包围盒判断是否相交或几乎相交）  
                List<List<Curve>> curveGroups = GroupConnectedCurves(allCurves, 0.001); // 将曲线按连接性分组  

                // 依次为每组生成外轮廓  
                int groupIndex = 0; // 定义组索引  
                foreach (var curveGroup in curveGroups) // 遍历所有分组  
                {
                    groupIndex++; // 递增组索引  
                    if (curveGroup.Count == 0) continue; // 如果当前组没有曲线，跳过  

                    // 3. 该分组的联合包围盒最小点minP  
                    Extents3d curveGroupExt = GetBoundingExtentsFromCurves(curveGroup); // 获取分组的包围盒  
                    Point3d minP = curveGroupExt.MinPoint; // 获取包围盒的最小点  

                    // 4. 沿最小点顺时针扫描一圈，采样最外侧交点  
                    List<Point2d> edgePoints = new List<Point2d>(); // 存储外边缘点  
                    double step = 2 * Math.PI / 365; // 计算每个采样点的角度  
                    for (int i = 0; i < 365; i++) // 循环进行采样  
                    {
                        double ang = i * step; // 当前采样的角度  
                        Vector2d dir = new Vector2d(Math.Cos(ang), Math.Sin(ang)); // 计算此方向的单位向量  
                        Point2d rayStart = new Point2d(minP.X, minP.Y); // 射线起点为最小点  
                        Point2d rayEnd = new Point2d(minP.X + dir.X * 1e4, minP.Y + dir.Y * 1e4); // 射线终点非常远（假设足够长）  
                        Point2d? pt = FindOuterBoundaryPoint(curveGroup, rayStart, rayEnd); // 查找此方向的外边界点  
                        if (pt != null) edgePoints.Add(pt.Value); // 如果找到有效点，则将其添加到列表  
                    }

                    // 若有效点不足3 个不能生成多段线  
                    if (edgePoints.Count < 3) // 如果边缘点低于3个  
                    {
                        ed.WriteMessage($"\n分组{groupIndex}点数不足无法生成多段线。"); // 提示信息  
                        continue; // 跳过当前分组  
                    }

                    // 5. 用所有顺时针点组成闭合多段线  
                    Polyline pl = new Polyline(); // 创建多段线对象  
                    for (int j = 0; j < edgePoints.Count; j++) // 遍历所有的外边缘点  
                        pl.AddVertexAt(j, edgePoints[j], 0, 0, 0); // 将点添加为多段线的顶点  
                    pl.Closed = true; // 闭合多段线  
                    pl.ColorIndex = groupIndex % 7 + 1; // 给不同组不同颜色（循环使用1-7的颜色）  
                    ms.AppendEntity(pl); // 将多段线添加到模型空间  
                    tr.AddNewlyCreatedDBObject(pl, true); // 将新对象注册到当前事务  

                    ed.WriteMessage($"\n第{groupIndex}组轮廓已生成。"); // 输出生成信息  
                }
                // 释放克隆的曲线资源  
                foreach (var c in allCurves) c.Dispose(); // 释放内存  

                tr.Commit(); // 提交事务  
                ed.WriteMessage("\n全部轮廓生成完成！"); // 输出完成信息  
            }
        }

        //--------↓ 曲线分组、几何方法（每行详细注释） ------------  

        /// <summary>  
        /// 分组：把空间相连/重叠/紧挨的曲线归为一组  
        /// </summary>  
        /// <param name="curves">所有曲线列表</param>  
        /// <param name="threshold">连接阈值</param>  
        /// <returns>分组后的曲线列表</returns>  
        private List<List<Curve>> GroupConnectedCurves(List<Curve> curves, double threshold)
        {
            List<List<Curve>> groups = new List<List<Curve>>(); // 保存分组  
            HashSet<Curve> processed = new HashSet<Curve>(); // 跟踪已处理的曲线  
            foreach (Curve curve in curves) // 遍历所有曲线  
            {
                if (processed.Contains(curve)) continue; // 已处理则跳过  
                Queue<Curve> queue = new Queue<Curve>(); // 用队列进行分组  
                queue.Enqueue(curve); // 将当前曲线入队  
                processed.Add(curve); // 将当前曲线标记为已处理  
                List<Curve> group = new List<Curve>(); // 创建当前组的曲线列表  
                while (queue.Count > 0) // 处理队列中的曲线  
                {
                    Curve current = queue.Dequeue(); // 出队当前曲线  
                    group.Add(current); // 将曲线添加到当前组  
                    foreach (Curve other in curves) // 遍历所有曲线  
                    {
                        if (processed.Contains(other)) continue; // 已处理则跳过  
                        if (AreConnected(current, other, threshold)) // 判断是否连接  
                        {
                            queue.Enqueue(other); // 入队连接的曲线  
                            processed.Add(other); // 标记为已处理  
                        }
                    }
                }
                groups.Add(group); // 添加当前分组  
            }
            return groups; // 返回分组列表  
        }

        /// <summary>
        /// 判定两曲线包围盒是否碰到或足够接近  
        /// </summary>
        /// <param name="c1">曲线1</param>
        /// <param name="c2">曲线2</param>
        /// <param name="threshold">阈值</param>
        /// <returns></returns>
        private bool AreConnected(Curve c1, Curve c2, double threshold)
        {
            try
            {
                if (!c1.Bounds.HasValue || !c2.Bounds.HasValue) return false; // 如果无包围盒则返回false  

                Extents3d ext1 = c1.Bounds.Value; // 获取曲线1的包围盒  
                Extents3d ext2 = c2.Bounds.Value; // 获取曲线2的包围盒  

                // 扩展后的包围盒1（加上阈值）  
                Point3d ext1Min = new Point3d(
                    ext1.MinPoint.X - threshold, // 扩展最小点  
                    ext1.MinPoint.Y - threshold,
                    ext1.MinPoint.Z - threshold);
                Point3d ext1Max = new Point3d(
                    ext1.MaxPoint.X + threshold, // 扩展最大点  
                    ext1.MaxPoint.Y + threshold,
                    ext1.MaxPoint.Z + threshold);

                // 判断点是否在扩展包围盒内  
                bool contains1to2Min = IsPointInBox(ext2.MinPoint, ext1Min, ext1Max); // 判断曲线2最小点是否在扩展包围盒1中  
                bool contains1to2Max = IsPointInBox(ext2.MaxPoint, ext1Min, ext1Max); // 判断曲线2最大点是否在扩展包围盒1中  
                bool contains2to1Min = IsPointInBox(ext1Min, ext2.MinPoint, ext2.MaxPoint); // 判断扩展包围盒1最小点是否在包围盒2中  
                bool contains2to1Max = IsPointInBox(ext1Max, ext2.MinPoint, ext2.MaxPoint); // 判断扩展包围盒1最大点是否在包围盒2中  

                // 或者检查包围盒相交（任一极点被包含，或包围盒交叉）  
                return contains1to2Min || contains1to2Max || contains2to1Min || contains2to1Max ||
                       DoBoxesIntersect(ext1Min, ext1Max, ext2.MinPoint, ext2.MaxPoint); // 检查包围盒相交  
            }
            catch { return false; } // 出现异常时返回false  
        }

        /// <summary>
        /// 判断点是否在包围盒内  
        /// </summary>
        /// <param name="pt">点</param>
        /// <param name="boxMin">包围盒最小点</param>
        /// <param name="boxMax">包围盒最大点</param>
        /// <returns>  </returns>
        private bool IsPointInBox(Point3d pt, Point3d boxMin, Point3d boxMax)
        {
            // 检查点的坐标是否在包围盒的范围内  
            return (pt.X >= boxMin.X && pt.X <= boxMax.X &&
                    pt.Y >= boxMin.Y && pt.Y <= boxMax.Y &&
                    pt.Z >= boxMin.Z && pt.Z <= boxMax.Z);
        }

        // 判断两个包围盒是否相交  
        private bool DoBoxesIntersect(Point3d box1Min, Point3d box1Max, Point3d box2Min, Point3d box2Max)
        {
            // 只要有一个轴向不相交，整体就不相交  
            return !(box1Max.X < box2Min.X || box1Min.X > box2Max.X ||
                     box1Max.Y < box2Min.Y || box1Min.Y > box2Max.Y ||
                     box1Max.Z < box2Min.Z || box1Min.Z > box2Max.Z);
        }

        /// <summary>  
        /// 收集实体（含块内容）所有曲线，递归爆炸，深度防止栈溢出  
        /// </summary>  
        /// <param name="entity">实体对象</param>  
        /// <param name="tr">开启事务</param>  
        /// <param name="curveList">曲线列队</param>  
        /// <param name="depth">递归次数</param>  
        private void CollectAllCurves(Entity entity, Transaction tr, List<Curve> curveList, int depth = 0)
        {
            if (depth > 10) return; // 最多递归10层，防止栈溢出  
            if (entity is Curve curve)
                curveList.Add((Curve)curve.Clone()); // 克隆当前曲线并添加到列表  
            else if (entity is BlockReference br)
            {
                using (var subObjs = new DBObjectCollection()) // 用于存储块爆炸后的对象  
                {
                    br.Explode(subObjs); // 爆炸块参照，获取内部对象  
                    foreach (DBObject dBObject in subObjs) // 遍历爆炸后的对象  
                    {
                        if (dBObject is Entity se)
                            CollectAllCurves(se, tr, curveList, depth + 1); // 递归收集曲线   
                        dBObject.Dispose(); // 释临时对象  
                    }
                }
            }
        }

        /// <summary>
        /// 获得曲线列表包围盒，取全部包围  
        /// </summary>
        /// <param name="curves">曲线列表</param>
        /// <returns>包围盒</returns>
        private Extents3d GetBoundingExtentsFromCurves(List<Curve> curves)
        {
            bool first = true; // 初始状态标记  
            Extents3d ext = new Extents3d(); // 初始化包围盒  
            foreach (Curve c in curves)
            {
                try
                {
                    if (!c.Bounds.HasValue) continue; // 如果无包围盒则跳过  
                    if (first)
                    {
                        ext = c.Bounds.Value; // 第一个曲线的包围盒赋值  
                        first = false; // 更新状态  
                    }
                    else
                    {
                        ext.AddExtents(c.Bounds.Value); // 将后续的包围盒与已有的包围盒合并  
                    }
                }
                catch { } // 捕获异常并继续  
            }
            return ext; // 返回组合后的包围盒  
        }

        /// <summary>
        /// 在所有曲线上查找射线最近点  
        /// </summary>
        /// <param name="curves">曲线列表</param>
        /// <param name="rayStart">射线起点</param>
        /// <param name="rayEnd">射线终点</param>
        /// <returns>找到的最近点</returns>
        private Point2d? FindOuterBoundaryPoint(List<Curve> curves, Point2d rayStart, Point2d rayEnd)
        {
            Point2d? bestPt = null; // 存储找到的最佳点  
            double maxDist = double.MinValue; // 初始化最大距离  
            foreach (Curve cv in curves) // 遍历所有曲线  
            {
                try
                {
                    Point3dCollection pts = new Point3dCollection(); // 用于存储交点  
                    cv.IntersectWith( // 查找交点  
                        new Line(new Point3d(rayStart.X, rayStart.Y, 0), new Point3d(rayEnd.X, rayEnd.Y, 0)),
                        Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero);
                    foreach (Point3d p in pts) // 遍历所有交点  
                    {
                        double dist = (p.X - rayStart.X) * (p.X - rayStart.X) + // 计算二次距离  
                                      (p.Y - rayStart.Y) * (p.Y - rayStart.Y);
                        if (dist > maxDist) // 找到了更远的点  
                        {
                            maxDist = dist; // 更新最大距离  
                            bestPt = new Point2d(p.X, p.Y); // 更新最佳点  
                        }
                    }
                }
                catch { } // 某些实体不支持交点会报错，直接忽略  
            }
            return bestPt; // 返回最佳点  
        }
        #endregion

        #endregion

        ///GetCircumcenter  GeneratePipeTableFromSelection getscale

        #region Excel相关 


        /// <summary>
        /// 选择Excel文件对话框 DBTextLabel
        /// </summary>
        /// <returns></returns>
        private string SelectExcelFile()
        {
            using (var dlg = new System.Windows.Forms.OpenFileDialog())
            {
                dlg.Filter = "Excel文件 (*.xlsx)|*.xlsx";
                return dlg.ShowDialog() == DialogResult.OK ? dlg.FileName : null;
            }
        }

        ////////////////////////////////////
        /// <summary>
        /// 插入Excel表格数据到CAD
        /// </summary>
        [CommandMethod("InsertExcelTableToCAD")] // 在CAD中执行命令：EXCELTABLE
        public void InsertExcelTableToCAD()
        {
            try
            {
                // 1) 选择文件
                string excelPath = SelectExcelFile();
                if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
                {
                    LogManager.Instance.LogInfo("\n未选择有效的Excel文件，操作已取消。");
                    return;
                }

                // 2) 定义目标字段（按业务顺序）
                string keyField = "分类号";
                string[] targetFields = { "分类号", "设备位号", "设备名称", "主要技术规格型号", "电压", "功率", "数量", "单重" };

                // 保存：第一行表头 + 后续数据
                List<List<string>> dataList = new List<List<string>>();

                // 3) 读取Excel
                using (var package = new ExcelPackage(new FileInfo(excelPath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null || worksheet.Dimension == null)
                    {
                        Application.ShowAlertDialog("Excel工作表为空或无有效数据。");
                        return;
                    }

                    int rowCount = worksheet.Dimension.End.Row;
                    int colCount = worksheet.Dimension.End.Column;

                    // 只在常见表头区间查找（保持你原有5~7行思路，同时做边界保护）
                    int headerStart = Math.Max(1, 5);
                    int headerEnd = Math.Min(7, rowCount);

                    // 字段 -> 列索引
                    Dictionary<string, int> fieldColumnMap = new Dictionary<string, int>();

                    for (int r = headerStart; r <= headerEnd; r++)
                    {
                        for (int c = 1; c <= colCount; c++)
                        {
                            string raw = worksheet.Cells[r, c].Text ?? string.Empty;
                            string colName = raw.Replace("\r", "").Replace("\n", "").Trim();
                            if (string.IsNullOrWhiteSpace(colName)) continue;

                            // 按 targetFields 顺序匹配，避免字典顺序不稳定
                            foreach (var target in targetFields)
                            {
                                if (!fieldColumnMap.ContainsKey(target) && colName.Contains(target))
                                {
                                    fieldColumnMap[target] = c;
                                    break;
                                }
                            }
                        }
                    }

                    if (!fieldColumnMap.ContainsKey(keyField))
                    {
                        Application.ShowAlertDialog("未找到“分类号”列，请检查Excel模板。");
                        return;
                    }

                    // 按目标字段顺序输出，仅保留找到的列
                    var orderedFields = targetFields.Where(f => fieldColumnMap.ContainsKey(f)).ToList();
                    if (orderedFields.Count == 0)
                    {
                        Application.ShowAlertDialog("未识别到可导入列，请检查表头。");
                        return;
                    }

                    // 表头
                    dataList.Add(new List<string>(orderedFields));

                    // 数据区起始行（沿用你原逻辑从第8行开始）
                    int dataStartRow = 8;
                    int keyCol = fieldColumnMap[keyField];

                    for (int row = dataStartRow; row <= rowCount; row++)
                    {
                        string classifyValue = (worksheet.Cells[row, keyCol].Text ?? string.Empty).Trim();

                        // 保持你现有行为：分类号非空即导入
                        if (string.IsNullOrWhiteSpace(classifyValue)) continue;

                        List<string> rowData = new List<string>(orderedFields.Count);
                        foreach (var field in orderedFields)
                        {
                            int col = fieldColumnMap[field];
                            rowData.Add((worksheet.Cells[row, col].Text ?? string.Empty).Trim());
                        }
                        dataList.Add(rowData);
                    }
                }

                if (dataList.Count <= 1)
                {
                    Application.ShowAlertDialog("没有可导入的数据（分类号为空）。");
                    return;
                }

                // 4) 插入到CAD
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                using (doc.LockDocument())
                using (DBTrans tr = new())
                {
                    TextFontsStyleHelper.TextStyleAndLayerInfo(tr, "设备层", 1, "tJText");

                    // 保持你原有展示顺序：数据倒序（表头会变到最后一行）
                    dataList.Reverse();

                    PromptPointResult ppr = Env.Editor.GetPoint("\n请在CAD中指定表格插入点：");
                    if (ppr.Status != PromptStatus.OK) return;

                    int rows = dataList.Count;
                    int cols = dataList[0].Count;

                    var table = new Autodesk.AutoCAD.DatabaseServices.Table
                    {
                        Position = ppr.Value,
                        Layer = "设备层"
                    };
                    table.SetSize(rows, cols);

                    // 统一单元格填充
                    for (int r = 0; r < rows; r++)
                    {
                        bool isHeaderRow = (r == rows - 1); // 因为上面 Reverse 了
                        for (int c = 0; c < cols; c++)
                        {
                            var cell = table.Cells[r, c];
                            cell.TextString = dataList[r][c];
                            cell.TextStyleId = tr.TextStyleTable["tJText"];
                            cell.TextHeight = isHeaderRow ? 350 : 300;
                            cell.Alignment = CellAlignment.MiddleCenter;
                        }
                    }

                    // 行高
                    for (int r = 0; r < rows; r++)
                    {
                        table.Rows[r].Height = (r == rows - 1) ? 500 : 450;
                    }

                    // 列宽（按文本长度估算，并做上限保护）
                    for (int c = 0; c < cols; c++)
                    {
                        int maxLen = 1;
                        for (int r = 0; r < rows; r++)
                        {
                            var txt = dataList[r][c] ?? string.Empty;
                            if (txt.Length > maxLen) maxLen = txt.Length;
                        }

                        // 经验值：每字符约180，最小600，最大6000，避免异常宽列
                        double width = Math.Min(6000.0, Math.Max(600.0, maxLen * 180.0));
                        table.Columns[c].Width = width;
                    }

                    table.GenerateLayout();
                    table.RecomputeTableBlock(true);

                    // 添加表格
                    tr.CurrentSpace.AddEntity(table);

                    // 保留你原逻辑：插入后分解为普通图元
                    DBObjectCollection explodedObjects = new DBObjectCollection();
                    table.Explode(explodedObjects);
                    foreach (DBObject obj in explodedObjects)
                    {
                        if (obj is Entity ent)
                        {
                            tr.CurrentSpace.AddEntity(ent);
                        }
                        else
                        {
                            obj.Dispose();
                        }
                    }
                    table.Erase();

                    tr.Commit();
                    Env.Editor.Redraw();
                }

                LogManager.Instance.LogInfo("\n表格已经插入到CAD。");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n导入Excel失败：{ex.Message}");
            }
        }

        #region 导出Excel方法一；

        /// <summary>
        /// 导出CAD表格到Excel的命令
        /// </summary>
        [CommandMethod("ExportCADTable")]
        public void ExportCADTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 获取用户选择的对象
            PromptSelectionResult psr = ed.GetSelection();
            if (psr.Status != PromptStatus.OK)
                return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 收集选择集中的所有实体
                    SelectionSet selectionSet = psr.Value;
                    //保存全部实体变量
                    List<Entity> allEntities = new List<Entity>();
                    //循环所有选择实体
                    foreach (SelectedObject selectedObject in selectionSet)
                    {
                        if (selectedObject != null)
                        {
                            Entity ent = tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead) as Entity;
                            if (ent != null)
                            {
                                // 处理块引用
                                if (ent is BlockReference blockRef)
                                {
                                    allEntities.AddRange(ExplodeBlockReference(blockRef, tr));
                                }
                                else
                                {
                                    allEntities.Add(ent);
                                }
                            }
                        }
                    }

                    // 分析表格结构并提取数据
                    var tableData = AnalyzeTable(allEntities);
                    if (tableData == null || tableData.Count == 0)
                    {
                        ed.WriteMessage("\n未检测到有效的表格结构！");
                        return;
                    }

                    // 获取保存文件路径
                    string filePath = GetSaveFilePathWithDialog();
                    //如果保存路径为空，返回
                    if (string.IsNullOrEmpty(filePath))
                        return;

                    // 导出到Excel
                    ExportToExcel(tableData, filePath);

                    ed.WriteMessage($"\n表格数据已成功导出到: {filePath}");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n错误: {ex.Message}");
                }
                finally
                {
                    tr.Commit();
                }
            }
        }

        /// <summary>
        /// 处理块引用，提取其中的实体
        /// </summary>
        /// <param name="blockRef">返回块表记录</param>
        /// <param name="tr">进程</param>
        /// <returns></returns>
        private List<Entity> ExplodeBlockReference(BlockReference blockRef, Transaction tr)
        {
            //储存块引用的实体
            List<Entity> entities = new List<Entity>();
            // 获取块表记录
            BlockTableRecord btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

            // 添加属性值作为文本
            if (blockRef.AttributeCollection != null)
            {
                foreach (ObjectId attId in blockRef.AttributeCollection)
                {
                    AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                    if (attRef != null && !attRef.IsConstant)
                    {
                        DBText text = new DBText();
                        text.TextString = attRef.TextString;
                        text.Position = attRef.Position;
                        text.Height = attRef.Height;
                        text.Rotation = attRef.Rotation;
                        text.AlignmentPoint = attRef.AlignmentPoint;
                        text.HorizontalMode = attRef.HorizontalMode;
                        text.VerticalMode = attRef.VerticalMode;
                        entities.Add(text);
                    }
                }
            }

            // 添加块中的实体
            foreach (ObjectId objId in btr)
            {
                Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                if (ent != null)
                {
                    // 递归处理嵌套块
                    if (ent is BlockReference nestedBlockRef)
                    {
                        entities.AddRange(ExplodeBlockReference(nestedBlockRef, tr));
                    }
                    else if (!(ent is AttributeDefinition)) // 排除属性定义
                    {
                        // 转换实体坐标
                        Entity transformedEnt = ent.GetTransformedCopy(blockRef.BlockTransform);
                        entities.Add(transformedEnt);
                    }
                }
            }

            return entities;
        }

        /// <summary>
        /// 分析表格结构并提取数据
        /// </summary>
        /// <param name="entities">实体集合</param>
        /// <returns>返回表数据</returns>
        private List<List<TableCell>> AnalyzeTable(List<Entity> entities)
        {
            /// 分离线条和文字
            List<Line> lines = entities.OfType<Line>().ToList();
            /// 分离DBText
            List<DBText> texts = entities.OfType<DBText>().ToList();
            /// 分离MText
            List<MText> mtexts = entities.OfType<MText>().ToList();

            /// 合并普通文本和多行文本
            List<TextEntity> allTexts = new List<TextEntity>();
            foreach (var text in texts)
            {
                allTexts.Add(new TextEntity
                {
                    Text = text.TextString,
                    Position = text.Position,
                    Height = text.Height
                });
            }

            foreach (var mtext in mtexts)
            {
                allTexts.Add(new TextEntity
                {
                    Text = mtext.Contents,
                    Position = mtext.Location,
                    Height = mtext.TextHeight
                });
            }

            // 提取表格的横线和竖线
            var horizontalLines = lines.Where(l => IsHorizontalLine(l)).ToList();
            var verticalLines = lines.Where(l => IsVerticalLine(l)).ToList();

            // 按Y坐标排序横线，按X坐标排序竖线
            horizontalLines = horizontalLines.OrderByDescending(l => l.StartPoint.Y).ToList();
            verticalLines = verticalLines.OrderBy(l => l.StartPoint.X).ToList();

            // 如果没有足够的线条来形成表格，返回null
            if (horizontalLines.Count < 2 || verticalLines.Count < 2)
                return null;

            // 创建表格结构
            int rowCount = horizontalLines.Count - 1;
            int colCount = verticalLines.Count - 1;
            List<List<TableCell>> tableData = new List<List<TableCell>>();

            // 初始化表格数据结构
            for (int i = 0; i < rowCount; i++)
            {
                tableData.Add(new List<TableCell>());
                for (int j = 0; j < colCount; j++)
                {
                    tableData[i].Add(null);
                }
            }

            // 分析合并单元格情况
            var mergedCells = AnalyzeMergedCells(horizontalLines, verticalLines);

            // 将文本分配到对应的单元格
            foreach (var text in allTexts)
            {
                // 查找文本所在的单元格
                int rowIndex = FindRowIndex(text.Position, horizontalLines);
                int colIndex = FindColumnIndex(text.Position, verticalLines);

                if (rowIndex >= 0 && rowIndex < rowCount && colIndex >= 0 && colIndex < colCount)
                {
                    // 检查是否属于合并单元格
                    var mergedCell = mergedCells.FirstOrDefault(mc =>
                        mc.Row == rowIndex && mc.Column == colIndex);

                    TableCell cell = new TableCell
                    {
                        Text = text.Text,
                        Position = text.Position,
                        Width = verticalLines[colIndex + 1].StartPoint.X - verticalLines[colIndex].StartPoint.X,
                        Height = horizontalLines[rowIndex].StartPoint.Y - horizontalLines[rowIndex + 1].StartPoint.Y
                    };

                    if (mergedCell != null)
                    {
                        cell.IsMerged = true;
                        cell.MergeAcross = mergedCell.ColumnSpan - 1;
                        cell.MergeDown = mergedCell.RowSpan - 1;
                    }

                    tableData[rowIndex][colIndex] = cell;
                }
            }

            return tableData;
        }

        /// <summary>
        /// 分析合并单元格
        /// </summary>
        /// <param name="horizontalLines"></param>
        /// <param name="verticalLines"></param>
        /// <returns></returns>
        private List<MergedCellInfo> AnalyzeMergedCells(List<Line> horizontalLines, List<Line> verticalLines)
        {
            var mergedCells = new List<MergedCellInfo>();
            double tolerance = 0.001;

            // 分析水平方向合并单元格
            for (int row = 0; row < horizontalLines.Count - 1; row++)
            {
                for (int col = 0; col < verticalLines.Count - 1; col++)
                {
                    // 检查当前单元格是否已被合并
                    if (mergedCells.Any(mc => mc.Row <= row && mc.Row + mc.RowSpan > row &&
                                              mc.Column <= col && mc.Column + mc.ColumnSpan > col))
                    {
                        continue;
                    }

                    // 计算当前单元格的边界
                    double left = verticalLines[col].StartPoint.X;
                    double right = verticalLines[col + 1].StartPoint.X;
                    double top = horizontalLines[row].StartPoint.Y;
                    double bottom = horizontalLines[row + 1].StartPoint.Y;

                    // 查找可能的水平合并
                    int colSpan = 1;
                    while (col + colSpan < verticalLines.Count - 1)
                    {
                        double nextRight = verticalLines[col + colSpan + 1].StartPoint.X;

                        // 检查是否缺少垂直分隔线
                        bool hasVerticalLine = false;
                        foreach (var line in verticalLines)
                        {
                            if (Math.Abs(line.StartPoint.X - verticalLines[col + colSpan].StartPoint.X) < tolerance &&
                                line.StartPoint.Y >= bottom - tolerance && line.EndPoint.Y <= top + tolerance)
                            {
                                hasVerticalLine = true;
                                break;
                            }
                        }

                        if (hasVerticalLine)
                            break;

                        colSpan++;
                    }

                    // 查找可能的垂直合并
                    int rowSpan = 1;
                    while (row + rowSpan < horizontalLines.Count - 1)
                    {
                        double nextBottom = horizontalLines[row + rowSpan + 1].StartPoint.Y;

                        // 检查是否缺少水平分隔线
                        bool hasHorizontalLine = false;
                        foreach (var line in horizontalLines)
                        {
                            if (Math.Abs(line.StartPoint.Y - horizontalLines[row + rowSpan].StartPoint.Y) < tolerance &&
                                line.StartPoint.X >= left - tolerance && line.EndPoint.X <= right - tolerance)
                            {
                                hasHorizontalLine = true;
                                break;
                            }
                        }

                        if (hasHorizontalLine)
                            break;

                        rowSpan++;
                    }

                    // 如果有合并，记录合并单元格信息
                    if (colSpan > 1 || rowSpan > 1)
                    {
                        mergedCells.Add(new MergedCellInfo
                        {
                            Row = row,
                            Column = col,
                            RowSpan = rowSpan,
                            ColumnSpan = colSpan
                        });
                    }
                }
            }

            return mergedCells;
        }

        /// <summary>
        /// 导出到Excel
        /// </summary>
        /// <param name="tableData"></param>
        /// <param name="filePath"></param>
        private void ExportToExcel(List<List<TableCell>> tableData, string filePath)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("CAD表格数据");

                // 获取表格的行列数
                int rowCount = tableData.Count;
                int colCount = rowCount > 0 ? tableData[0].Count : 0;

                // 填充Excel表格并设置边框
                for (int row = 0; row < rowCount; row++)
                {
                    for (int col = 0; col < colCount; col++)
                    {
                        var cell = tableData[row][col];

                        if (cell != null && cell.IsMerged && (cell.MergeAcross > 0 || cell.MergeDown > 0))
                        {
                            // 合并单元格
                            worksheet.Cells[row + 1, col + 1,
                                           row + 1 + cell.MergeDown,
                                           col + 1 + cell.MergeAcross].Merge = true;
                        }

                        // 设置单元格值
                        if (cell != null)
                        {
                            worksheet.Cells[row + 1, col + 1].Value = cell.Text;
                        }

                        // 设置单元格边框
                        SetCellBorder(worksheet.Cells[row + 1, col + 1]);
                    }
                }

                // 自动调整列宽
                worksheet.Cells.AutoFitColumns();

                // 保存Excel文件
                package.Save();
            }
        }

        /// <summary>
        /// 设置单元格边框
        /// </summary>
        /// <param name="cell"></param>
        private void SetCellBorder(ExcelRange cell)
        {
            cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }

        /// <summary>
        /// 判断是否为横线
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        private bool IsHorizontalLine(Line line)
        {
            double tolerance = 0.001;
            return Math.Abs(line.StartPoint.Y - line.EndPoint.Y) < tolerance;
        }

        /// <summary>
        /// 判断是否为竖线
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        private bool IsVerticalLine(Line line)
        {
            double tolerance = 0.001;
            return Math.Abs(line.StartPoint.X - line.EndPoint.X) < tolerance;
        }

        /// <summary>
        /// 查找文本所在的行索引
        /// </summary>
        /// <param name="position"></param>
        /// <param name="horizontalLines"></param>
        /// <returns></returns>
        private int FindRowIndex(Point3d position, List<Line> horizontalLines)
        {
            for (int i = 0; i < horizontalLines.Count - 1; i++)
            {
                if (position.Y <= horizontalLines[i].StartPoint.Y &&
                    position.Y >= horizontalLines[i + 1].StartPoint.Y)
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// 查找文本所在的列索引
        /// </summary>
        /// <param name="position"></param>
        /// <param name="verticalLines"></param>
        /// <returns></returns>
        private int FindColumnIndex(Point3d position, List<Line> verticalLines)
        {
            for (int i = 0; i < verticalLines.Count - 1; i++)
            {
                if (position.X >= verticalLines[i].StartPoint.X &&
                    position.X <= verticalLines[i + 1].StartPoint.X)
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// 使用AutoCAD文件对话框获取保存文件路径
        /// </summary>
        /// <returns></returns>
        private string GetSaveFilePathWithDialog()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // 创建文件保存对话框
            System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "Excel文件 (*.xlsx)|*.xlsx";
            sfd.Title = "保存表格数据到Excel文件";
            sfd.DefaultExt = "xlsx";

            // 获取当前CAD文档路径作为默认路径
            string currentDocPath = doc.Name;
            if (!string.IsNullOrEmpty(currentDocPath))
            {
                string currentDir = Path.GetDirectoryName(currentDocPath);
                sfd.InitialDirectory = currentDir;
            }

            // 显示对话框
            System.Windows.Forms.DialogResult result = sfd.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                return sfd.FileName;
            }

            return null;
        }

        #endregion

        #endregion

        #region 清理、检查等命令
        /// <summary>
        /// 分解选定的块
        /// </summary>
        [CommandMethod("ExplodeBlockToNewBlock")]
        public void ExplodeBlockToNewBlock()
        {
            // 让用户选择一个块参照
            PromptEntityOptions peo = new PromptEntityOptions("\n请选择一个块参照: ");
            peo.SetRejectMessage("\n请选择块参照对象！");
            peo.AddAllowedClass(typeof(BlockReference), true);
            PromptEntityResult per = Env.Editor.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (var tr = new DBTrans())
            {
                // 获取用户选择的块参照对象
                BlockReference? blkRef = tr.GetObject(per.ObjectId, OpenMode.ForRead) as BlockReference;

                // 递归分解块，收集所有基本图元
                List<Entity> allEntities = new List<Entity>();
                if (blkRef != null)
                    ExplodeBlockRecursive(blkRef, tr, allEntities);

                if (allEntities.Count == 0)
                {
                    LogManager.Instance.LogInfo("\n未找到可用于新块的图元。");
                    return;
                }

                // 创建新块定义
                BlockTable bt = (BlockTable)tr.GetObject(tr.Database.BlockTableId, OpenMode.ForWrite);
                string newBlockName = "SB_" + System.DateTime.Now.Ticks;
                BlockTableRecord newBtr = new BlockTableRecord();
                newBtr.Name = newBlockName;


                // 将所有分解后的图元加入新块
                foreach (var ent in allEntities)
                {
                    newBtr.AppendEntity(ent);
                }
                bt.Add(newBtr);

                // 在模型空间插入新块参照
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                BlockReference newBlkRef = new BlockReference(Point3d.Origin, newBtr.ObjectId);
                ms.AppendEntity(newBlkRef);

                LogManager.Instance.LogInfo($"\n新块 \"{newBlockName}\" 已创建并插入。");

                tr.Commit();
            }
        }

        /// <summary>
        /// 递归分解块，将所有基本图元加入列表
        /// </summary>
        /// <param name="blkRef">块表记录</param>
        /// <param name="tr">事务</param>
        /// <param name="entityList">实体例表</param>
        private void ExplodeBlockRecursive(BlockReference blkRef, DBTrans tr, List<Entity> entityList)
        {
            // 分解当前块参照
            DBObjectCollection explodedObjs = new DBObjectCollection();
            blkRef.Explode(explodedObjs);

            foreach (DBObject obj in explodedObjs)
            {
                if (obj is BlockReference nestedBlkRef)
                {
                    // 如果是嵌套块，递归分解
                    ExplodeBlockRecursive(nestedBlkRef, tr, entityList);
                    obj.Dispose(); // 释放临时对象
                }
                else if (obj is Entity ent)
                {
                    // 复制实体，避免事务结束后被释放
                    Entity? clone = ent.Clone() as Entity;
                    if (clone != null)
                        entityList.Add(clone);
                    ent.Dispose(); // 释放临时对象
                }
                else
                {
                    obj.Dispose(); // 释放非实体对象
                }
            }
        }

        /// <summary>  
        /// DICTS命令的实现  
        /// 功能：列出所有字典及其计数，并允许删除选定的字典  
        /// </summary>  
        [CommandMethod("DICTS")]
        public void ListAndManageDictionaries()
        {

            bool continueCommand = true;

            while (continueCommand)
            {
                try
                {
                    using var tr = new DBTrans();
                    // 获取命名对象字典  
                    DBDictionary nod = (DBDictionary)tr.GetObject(
                        tr.NamedObjectsDict.Id,
                        OpenMode.ForRead);

                    // 开始撤销标记  
                    //doc.TransactionManager.StartUndoMark();

                    // 初始化计数器和集合  
                    int index = 1;
                    var dictNames = new List<string>();
                    var dictIds = new List<ObjectId>();

                    // 遍历所有字典  
                    foreach (DBDictionaryEntry entry in nod)
                    {
                        string count = GetDictionaryCount(tr, entry.Value);
                        LogManager.Instance.LogInfo($"\n{index}. \"{entry.Key}\"  {count}");

                        dictNames.Add(entry.Key);
                        dictIds.Add(entry.Value);
                        index++;
                    }

                    LogManager.Instance.LogInfo($"\nActiveDocument.Dictionaries.Count={dictNames.Count}\n");

                    // 设置用户输入选项  
                    PromptIntegerOptions opts = new PromptIntegerOptions(
                        "\nWhich one to REMOVE by index above? <Enter to exit>: ")
                    {
                        AllowNegative = false,
                        AllowZero = false,
                        UpperLimit = dictNames.Count,
                        AllowNone = true // 允许直接按回车  
                    };

                    // 获取用户输入  
                    PromptIntegerResult result = Env.Editor.GetInteger(opts);

                    // 如果用户按回车，退出命令  
                    if (result.Status == PromptStatus.None)
                    {
                        continueCommand = false;
                        LogManager.Instance.LogInfo("\nYou can type command DICTS to go again.");
                    }
                    // 如果用户输入了有效的索引号  
                    else if (result.Status == PromptStatus.OK)
                    {
                        int selectedIndex = result.Value - 1;
                        if (selectedIndex >= 0 && selectedIndex < dictIds.Count)
                        {
                            // 打开字典进行写操作并删除选定项  
                            DBDictionary dict = (DBDictionary)tr.GetObject(
                                nod.ObjectId,
                                OpenMode.ForWrite);
                            dict.Remove(dictNames[selectedIndex]);

                            // 提交事务  
                            tr.Commit();
                            LogManager.Instance.LogInfo($"\nDictionary \"{dictNames[selectedIndex]}\" has been removed.");
                        }
                    }

                    // 结束撤销标记  
                    //doc.TransactionManager.EndUndoMark();

                }
                catch (System.Exception ex)
                {
                    LogManager.Instance.LogInfo($"\nError: {ex.Message}");
                    continueCommand = false;
                }
            }
        }

        /// <summary>  
        /// 获取字典的计数  
        /// </summary>  
        private string GetDictionaryCount(Transaction trans, ObjectId dictId)
        {
            try
            {
                DBDictionary dict = (DBDictionary)trans.GetObject(dictId, OpenMode.ForRead);
                return dict.Count.ToString();
            }
            catch
            {
                return "#n/a";
            }
        }

        /// <summary>  
        /// 执行清理操作的命令  
        /// </summary>  
        [CommandMethod("CLEANUPDWG")]
        public void CleanupDrawing()
        {

        }
        #endregion



        #region 结尾：增强的块放置 Jig（已修改以订阅方向事件）
        /// <summary>
        /// 增强的块放置拖拽类，完整支持组合块预览并响应方向按键变化用于旋转预览
        /// </summary>
        private class EnhancedBlockPlacementJig : EntityJig, IDisposable
        {
            /// <summary>
            /// 块位置 插入点
            /// </summary>
            private Point3d _position;//块位置 插入点
            /// <summary>
            /// 块旋转角度（弧度）
            /// </summary>
            private double _rotation;
            /// <summary>
            /// 块缩放比例
            /// </summary>
            private Scale3d _scale;
            /// <summary>
            /// 块ID
            /// </summary>
            private ObjectId _blockId;
            /// <summary>
            /// 属性定义列表
            /// </summary>
            private List<AttributeDefinition> _attDefs;
            /// <summary>
            /// 块图层
            /// </summary>
            private string _blockLayer;
            /// <summary>
            ///  订阅方向事件
            /// </summary>
            private bool _subscribed = false;
            /// <summary>
            /// 释放资源
            /// </summary>
            private bool _disposed = false;

            // 属性
            /// <summary>
            /// 块位置 插入点
            /// </summary>
            public Point3d Position => _position;
            /// <summary>
            /// 块旋转角度
            /// </summary>
            public double Rotation => _rotation;
            /// <summary>
            /// 块缩放比例
            /// </summary>
            public Scale3d Scale => _scale;//块缩放比例

            /// <summary>
            /// 拖拽类的 - 动态增强的块放置类
            /// </summary>
            /// <param name="blockId">块ID</param>
            /// <param name="attDefs">属性定义列表</param>
            /// <param name="blockLayer"> 块图层</param>
            /// <param name="scale">比例</param>
            public EnhancedBlockPlacementJig(ObjectId blockId, List<AttributeDefinition> attDefs, string? blockLayer = null, Scale3d? scale = null)
                : base(new BlockReference(Point3d.Origin, blockId))
            {
                _blockId = blockId;//块ID
                _position = Point3d.Origin;//初始位置设为原点
                _rotation = VariableDictionary.entityRotateAngle; // 使用预设旋转角度
                _attDefs = attDefs ?? new List<AttributeDefinition>();//属性定义列表
                _blockLayer = blockLayer ?? "0";//块图层，默认"0"图层
                _scale = scale ?? new Scale3d(1, 1, 1);//块缩放比例，默认1:1比例

                // 配置预览实体
                if (Entity is BlockReference blockRef)
                {
                    blockRef.Layer = _blockLayer;//设置块图层
                    blockRef.ScaleFactors = _scale;//设置块缩放比例
                    blockRef.Rotation = _rotation;
                }

                // 订阅 Command 方向变更通知（允许外部按钮即时驱动旋转）
                if (!_subscribed)
                {
                    Command.DirectionChanged += OnDirectionChanged;
                    _subscribed = true;
                }
            }

            // 方向变更时直接修改预览实体并刷新编辑器显示（使拖拽期间实时可见）
            private void OnDirectionChanged(double angle)
            {
                try
                {
                    // 更新内部角度
                    _rotation = angle;

                    // 直接修改正在预览的实体（Entity）旋转值
                    if (Entity is BlockReference br)
                    {
                        br.Rotation = _rotation;
                    }

                    // 刷新当前文档视图，让预览立即反映
                    var doc = Application.DocumentManager.MdiActiveDocument;
                    doc?.Editor.UpdateScreen();
                }
                catch
                {
                    // 忽略以确保不会打断主流程
                }
            }

            /// <summary>
            /// 获取预览实体
            /// </summary>
            /// <param name="prompts">提示</param>
            /// <returns></returns>
            protected override SamplerStatus Sampler(JigPrompts prompts)
            {
                // 获取插入点
                JigPromptPointOptions pointOpts = new JigPromptPointOptions("\n指定组合块插入点（右键确认）:");
                pointOpts.UserInputControls = UserInputControls.Accept3dCoordinates;
                pointOpts.UseBasePoint = false;

                PromptPointResult pointResult = prompts.AcquirePoint(pointOpts);

                // 如果用户取消
                if (pointResult.Status != PromptStatus.OK)
                    return SamplerStatus.Cancel;

                // 如果位置和角度都没变，返回NoChange
                if (_position.DistanceTo(pointResult.Value) < 0.001 && Math.Abs(_rotation - VariableDictionary.entityRotateAngle) < 1e-9)
                    return SamplerStatus.NoChange;

                // 更新位置
                _position = pointResult.Value;

                // 始终使用 VariableDictionary.entityRotateAngle 作为旋转角度基准（外部按钮会通过事件另外通知）
                _rotation = VariableDictionary.entityRotateAngle;

                return SamplerStatus.OK;
            }

            /// <summary>
            ///  更新数据
            /// </summary>
            /// <returns></returns>
            protected override bool Update()
            {
                try
                {
                    // 更新块引用的位置、旋转和比例
                    if (Entity is BlockReference blockRef)
                    {
                        blockRef.Position = _position;
                        blockRef.Rotation = _rotation;
                        blockRef.ScaleFactors = _scale;
                    }

                    return true;
                }
                catch (Exception ex)
                {
                    // 记录错误但不抛出异常
                    LogManager.Instance.LogInfo($"\n块更新时发生错误: {ex.Message}");
                    return false;
                }
            }

            /// <summary>
            /// 显式释放：取消订阅避免泄露
            /// </summary>
            public void Dispose()
            {
                if (_disposed) return;
                if (_subscribed)
                {
                    try
                    {
                        Command.DirectionChanged -= OnDirectionChanged;
                    }
                    catch { }
                    _subscribed = false;
                }
                _disposed = true;
                GC.SuppressFinalize(this);
            }

            ~EnhancedBlockPlacementJig()
            {
                Dispose();
            }
        }
        #endregion
        #region 其它内部类（LayerState、TableCell、TextEntity、MergedCellInfo、BlockReferenceInfo等）

        /// <summary>  
        /// 图层状态信息类，用于存储单个图层的详细状态  
        /// </summary>  
        public class LayerState
        {
            public ObjectId ObjectId { get; set; }
            /// <summary>
            /// 图层名称
            /// </summary>
            public string? Name { get; set; }            // 图层名称
            /// <summary>
            /// 图层是否关闭
            /// </summary>
            public bool IsOff { get; set; }             // 图层是否关闭
            /// <summary>
            /// 图层是否冻结
            /// </summary>
            public bool IsFrozen { get; set; }          // 图层是否冻结
            /// <summary>
            /// 图层是否锁定
            /// </summary>
            public bool IsLocked { get; set; }          // 图层是否锁定
            /// <summary>
            /// 图层是否可打印
            /// </summary>
            public bool IsPlottable { get; set; }       // 图层是否可打印
            /// <summary>
            /// 图层颜色索引
            /// </summary>
            public short ColorIndex { get; set; }       // 图层颜色索引
            /// <summary>
            /// 图层线宽
            /// </summary>
            public LineWeight LineWeight { get; set; }  // 图层线宽
            /// <summary>
            /// 图层打印样式名称
            /// </summary>
            public string? PlotStyleName { get; set; }   // 图层打印样式名称
            /// <summary>
            /// 图层绘制数据
            /// </summary>
            public DrawStream? DrawStream { get; set; } // 图层绘制数据
            /// <summary>
            /// 图层是否隐藏
            /// </summary>
            public bool IsHidden { get; set; } // 图层是否隐藏
            /// <summary>
            /// 图层是否同步
            /// </summary>
            public bool IsReconciled { get; set; } // 图层是否同步
            /// <summary>
            /// 图层线型对象ID
            /// </summary>
            public ObjectId LinetypeObjectId { get; set; } // 图层线型对象ID
            /// <summary>
            /// 图层材质ID
            /// </summary>
            public ObjectId MaterialId { get; set; } // 图层材质ID
            /// <summary>
            /// 图层合并样式
            /// </summary>
            public DuplicateRecordCloning MergeStyle { get; set; } // 图层合并样式
            /// <summary>
            /// 图层所有者ID
            /// </summary>
            public ObjectId OwnerId { get; set; } // 图层所有者ID
            /// <summary>
            /// 图层打印样式名称ID
            /// </summary>
            public ObjectId PlotStyleNameId { get; set; } // 图层打印样式名称ID
            /// <summary>
            /// 图层透明度
            /// </summary>
            public Transparency Transparency { get; set; } // 图层透明度
            /// <summary>
            /// 图层视图可见性默认值
            /// </summary>
            public bool ViewportVisibilityDefault { get; set; } // 图层视图可见性默认值
            /// <summary>
            /// 图层X数据
            /// </summary>
            public ResultBuffer? XData { get; set; } // 图层X数据
            /// <summary>
            /// 图层注释
            /// </summary>
            public AnnotativeStates Annotative { get; set; } // 图层注释
            /// <summary>
            /// 图层描述
            /// </summary>
            public string? Description { get; set; } // 图层描述
            /// <summary>
            /// 图层保存版本覆盖
            /// </summary>
            public bool HasSaveVersionOverride { get; set; } // 图层保存版本覆盖
        }

        /// <summary>
        /// 表格单元格类
        /// </summary>
        public class TableCell
        {
            /// <summary>
            /// 表文字
            /// </summary>
            public string? Text { get; set; }
            /// <summary>
            /// 表坐标
            /// </summary>
            public Point3d Position { get; set; }
            /// <summary>
            /// 表宽
            /// </summary>
            public double Width { get; set; }
            /// <summary>
            /// 表高
            /// </summary>
            public double Height { get; set; }
            /// <summary>
            /// 合并格
            /// </summary>
            public bool IsMerged { get; set; } = false;
            /// <summary>
            /// 合并跨越
            /// </summary>
            public int MergeAcross { get; set; } = 0;
            /// <summary>
            /// 向下合并
            /// </summary>
            public int MergeDown { get; set; } = 0;
        }

        /// <summary>
        /// 辅助类：用于统一处理DBText和MText
        /// </summary>
        private class TextEntity
        {
            /// <summary>
            /// 文本
            /// </summary>
            public string? Text { get; set; }
            /// <summary>
            /// 位置
            /// </summary>
            public Point3d Position { get; set; }
            /// <summary>
            /// 高度
            /// </summary>
            public double Height { get; set; }
        }

        /// <summary>
        /// 辅助类：用于表示合并单元格信息
        /// </summary>
        private class MergedCellInfo
        {
            /// <summary>
            /// 行索引
            /// </summary>
            public int Row { get; set; }
            /// <summary>
            /// 列索引
            /// </summary>
            public int Column { get; set; }
            /// <summary>
            /// 行跨度
            /// </summary>
            public int RowSpan { get; set; }
            /// <summary>
            /// 列跨度
            /// </summary>
            public int ColumnSpan { get; set; }
        }

        /// <summary>
        /// 用于保存块引用信息的辅助类
        /// </summary>
        private class BlockReferenceInfo
        {
            /// <summary>
            /// 名称
            /// </summary>
            public string? Name { get; set; }
            /// <summary>
            /// 项目id
            /// </summary>
            public ObjectId Id { get; set; }
            /// <summary>
            /// 3D坐标点
            /// </summary>
            public Point3d Position { get; set; }
            /// <summary>
            /// 比例因子
            /// </summary>
            public Scale3d ScaleFactors { get; set; }
            /// <summary>
            /// 角度
            /// </summary>
            public double Rotation { get; set; }
            /// <summary>
            /// 比例
            /// </summary>
            public Scale3d Scale { get; set; }
            /// <summary>
            /// 图层名
            /// </summary>
            public string? Layer { get; set; }
            /// <summary>
            /// 线类型 objectid
            /// </summary>
            public ObjectId Linetype { get; set; }
            /// <summary>
            /// 线型比例
            /// </summary>
            public double LinetypeScale { get; set; }
            /// <summary>
            /// 线宽
            /// </summary>
            public Autodesk.AutoCAD.DatabaseServices.LineWeight Lineweight { get; set; }
            /// <summary>
            /// 线色号
            /// </summary>
            public Autodesk.AutoCAD.Colors.Color? Color { get; set; }
            /// <summary>
            /// 线色号
            /// </summary>
            public int ColorIndex { get; set; }
            /// <summary>
            /// 属性
            /// </summary>
            public Dictionary<string, object> AttributeValues { get; set; } = new Dictionary<string, object>();
            /// <summary>
            /// 透明度
            /// </summary>
            public bool Visibility { get; set; }
            /// <summary>
            /// 块法向量
            /// </summary>
            public Vector3d Normal { get; set; }
        }

        #endregion

        #region 生成excel与表格分析相关的内部方法（ExtractTableData、AnalyzeMergedCells、ExportToExcel、SetCellBorder、IsHorizontalLine、IsVerticalLine、FindRowIndex、FindColumnIndex、GetSaveFilePathWithDialog等）

        [CommandMethod("ExportSheetFormulaCsv")]
        public void ExportSheetFormulaCsv()
        {
            try
            {
                // 1) 选择 Excel
                string excelPath = SelectExcelFile();
                if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
                {
                    LogManager.Instance.LogInfo("\n未选择有效的Excel文件，操作已取消。");
                    return;
                }

                // 2) 读取工作簿
                using var package = new ExcelPackage(new FileInfo(excelPath));
                var workbook = package.Workbook;
                if (workbook == null || workbook.Worksheets.Count == 0)
                {
                    Application.ShowAlertDialog("Excel中没有可用工作表。");
                    return;
                }

                // 3) 输入 Sheet 名（默认：设备选型、管径计算）
                string defaultSheetName = "设备选型、管径计算";
                var pso = new PromptStringOptions($"\n请输入Sheet名称 <{defaultSheetName}>: ")
                {
                    AllowSpaces = true
                };
                var psr = Env.Editor.GetString(pso);

                string sheetName = defaultSheetName;
                if (psr.Status == PromptStatus.OK && !string.IsNullOrWhiteSpace(psr.StringResult))
                {
                    sheetName = psr.StringResult.Trim();
                }

                var ws = workbook.Worksheets[sheetName];
                if (ws == null)
                {
                    Application.ShowAlertDialog($"未找到工作表：{sheetName}");
                    return;
                }

                if (ws.Dimension == null)
                {
                    Application.ShowAlertDialog($"工作表“{sheetName}”没有有效数据。");
                    return;
                }

                // 4) 选择导出 CSV 路径
                string csvPath = SelectSaveCsvPath(excelPath, sheetName);
                if (string.IsNullOrWhiteSpace(csvPath))
                {
                    LogManager.Instance.LogInfo("\n未选择CSV保存路径，操作已取消。");
                    return;
                }

                // 5) 导出：Address, Formula, Value
                int rowStart = ws.Dimension.Start.Row;
                int rowEnd = ws.Dimension.End.Row;
                int colStart = ws.Dimension.Start.Column;
                int colEnd = ws.Dimension.End.Column;

                var sb = new System.Text.StringBuilder(1024 * 32);
                sb.AppendLine("Address,Formula,Value");

                int exportCount = 0;
                for (int r = rowStart; r <= rowEnd; r++)
                {
                    for (int c = colStart; c <= colEnd; c++)
                    {
                        var cell = ws.Cells[r, c];
                        if (cell == null) continue;

                        string formula = cell.Formula ?? string.Empty;
                        string value = cell.Text ?? string.Empty;

                        // 只导出有值或有公式的单元格，减少噪音
                        if (string.IsNullOrWhiteSpace(formula) && string.IsNullOrWhiteSpace(value))
                            continue;

                        if (!string.IsNullOrWhiteSpace(formula))
                        {
                            formula = "=" + formula;
                        }

                        sb.Append(CsvEscape(cell.Address)).Append(',')
                          .Append(CsvEscape(formula)).Append(',')
                          .Append(CsvEscape(value)).AppendLine();

                        exportCount++;
                    }
                }

                File.WriteAllText(csvPath, sb.ToString(), new System.Text.UTF8Encoding(true));
                LogManager.Instance.LogInfo($"\n导出完成：{csvPath}，共 {exportCount} 条。");
                Application.ShowAlertDialog($"导出完成：{csvPath}\n共 {exportCount} 条。");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n导出公式CSV失败：{ex.Message}");
                Application.ShowAlertDialog($"导出失败：{ex.Message}");
            }
        }

        private static string SelectSaveCsvPath(string excelPath, string sheetName)
        {
            using var sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.Filter = "CSV 文件 (*.csv)|*.csv";
            sfd.Title = "保存公式导出文件";
            sfd.DefaultExt = "csv";
            sfd.AddExtension = true;
            sfd.FileName = $"{Path.GetFileNameWithoutExtension(excelPath)}_{sheetName}_公式导出.csv";
            sfd.InitialDirectory = Path.GetDirectoryName(excelPath);

            return sfd.ShowDialog() == DialogResult.OK ? sfd.FileName : string.Empty;
        }

        private static string CsvEscape(string? input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            string s = input.Replace("\"", "\"\"");
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\r") || s.Contains("\n"))
            {
                return $"\"{s}\"";
            }
            return s;
        }

        #endregion

        /// <summary>
        /// 面积选点专用预览 Jig：
        /// 1) 动态显示“已确认点 + 当前鼠标点”形成的多段线预览；
        /// 2) 支持关键字 Z 撤销（LastKeyword 兜底，兼容不同返回状态）
        /// </summary>
        private sealed class AreaPolylinePreviewJig : DrawJig
        {
            /// <summary>
            /// 已确认点列表（UCS）
            /// </summary>
            private readonly List<Point2d> _fixedPoints;

            /// <summary>
            /// 预览图层
            /// </summary>
            private readonly string _layerName;

            /// <summary>
            /// 预览颜色
            /// </summary>
            private readonly short _colorIndex;

            /// <summary>
            /// 当前鼠标点（WCS）
            /// </summary>
            public Point3d CurrentPointWcs { get; private set; } = Point3d.Origin;

            /// <summary>
            /// 记录最后一次关键字（用于某些环境 Cancel + Keyword 的兜底识别）
            /// </summary>
            public string? LastKeyword { get; private set; }

            /// <summary>
            /// 当前是否已有有效鼠标点
            /// </summary>
            private bool _hasCurrentPoint = false;

            /// <summary>
            /// 构造
            /// </summary>
            public AreaPolylinePreviewJig(List<Point2d> fixedPoints, string layerName, short colorIndex)
            {
                _fixedPoints = fixedPoints ?? new List<Point2d>();
                _layerName = layerName;
                _colorIndex = colorIndex;
            }

            /// <summary>
            /// 采样输入（点输入 + 关键字输入）
            /// </summary>
            protected override SamplerStatus Sampler(JigPrompts prompts) // 采样输入
            {
                JigPromptPointOptions ppo = new JigPromptPointOptions("\n指定下一个点 [Z=撤销上一步]，右键结束："); // 创建点输入选项

                ppo.UserInputControls = // 配置输入控制
                    UserInputControls.Accept3dCoordinates | // 接受三维坐标输入
                    UserInputControls.GovernedByOrthoMode | // 受正交模式约束
                    UserInputControls.NullResponseAccepted; // 允许回车/右键结束

                if (_fixedPoints.Count > 0) // 如果已有确认点
                {
                    var lastUcs = _fixedPoints[_fixedPoints.Count - 1]; // 取最后一个确认点（UCS）
                    var baseWcs = new Point3d(lastUcs.X, lastUcs.Y, 0).Ucs2Wcs(); // 转成 WCS 基点
                    ppo.UseBasePoint = true; // 启用基点约束
                    ppo.BasePoint = baseWcs; // 设置正交参考基点
                }

                ppo.Keywords.Add("Z"); // 添加撤销关键字 Z
                ppo.AppendKeywordsToMessage = true; // 在提示中显示关键字

                PromptPointResult ppr = prompts.AcquirePoint(ppo); // 获取输入结果

                if (ppr.Status == PromptStatus.Keyword) // 如果输入关键字
                {
                    LastKeyword = ppr.StringResult; // 记录关键字
                    return SamplerStatus.Cancel; // 交由外层处理撤销
                }

                if (ppr.Status != PromptStatus.OK) // 非正常点输入
                {
                    return SamplerStatus.Cancel; // 结束本轮
                }

                if (_hasCurrentPoint && ppr.Value.DistanceTo(CurrentPointWcs) < 1e-6) // 点未变化
                {
                    return SamplerStatus.NoChange; // 不重绘
                }

                CurrentPointWcs = ppr.Value; // 更新当前点
                LastKeyword = null; // 清空关键字
                _hasCurrentPoint = true; // 标记已有当前点
                return SamplerStatus.OK; // 通知刷新预览
            }

            /// <summary>
            /// 绘制预览图形
            /// </summary>
            protected override bool WorldDraw(WorldDraw draw)
            {
                // 没有已确认点就不绘制
                if (_fixedPoints.Count == 0)
                    return true;

                // 创建预览多段线（临时对象，仅用于预览）
                using Polyline pl = new Polyline
                {
                    Layer = _layerName,
                    ColorIndex = _colorIndex
                };

                // 先加入所有已确认点（UCS）
                for (int i = 0; i < _fixedPoints.Count; i++)
                {
                    pl.AddVertexAt(i, _fixedPoints[i], 0, 50, 50);
                }

                // 再加入当前鼠标点（WCS -> UCS），形成动态末端
                if (_hasCurrentPoint)
                {
                    var curUcs = CurrentPointWcs.Wcs2Ucs().Z20();
                    pl.AddVertexAt(_fixedPoints.Count, new Point2d(curUcs.X, curUcs.Y), 0, 50, 50);
                }

                // 点数足够时闭合显示预览区域
                if (_fixedPoints.Count >= 2)
                {
                    pl.Closed = true;
                }

                // 绘制到屏幕
                draw.Geometry.Draw(pl);
                return true;
            }
        }

    }


}

