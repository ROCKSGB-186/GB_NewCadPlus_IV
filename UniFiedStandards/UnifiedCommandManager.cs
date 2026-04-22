using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Helpers;
using System;
using System.Linq;

namespace GB_NewCadPlus_IV.UniFiedStandards
{
    /// <summary>
    /// 统一命令管理器 - 整合所有专业命令
    /// </summary>
    public static class UnifiedCommandManager
    {
        /// <summary>
        /// 命令映射字典
        /// </summary>
        private static readonly Dictionary<string, Action> _commandMap = new Dictionary<string, Action>
    {
        #region 方向按键
        { "上", () => ExecuteDirectionCommand("上", Math.PI * 1.5) },
        { "右上", () => ExecuteDirectionCommand("右上", Math.PI * 1.25) },
        { "右", () => ExecuteDirectionCommand("右", Math.PI * 1) },
        { "右下", () => ExecuteDirectionCommand("右下", Math.PI * 0.75) },
        { "下", () => ExecuteDirectionCommand("下", Math.PI * 0.5) },
        { "左下", () => ExecuteDirectionCommand("左下", Math.PI * 0.25) },
        { "左", () => ExecuteDirectionCommand("左", 0) },
        { "左上", () => ExecuteDirectionCommand("左上", Math.PI * 1.75) },
        #endregion

        #region 功能按键
        { "查找", () => ExecuteFunctionCommand("查找") },
        { "PU", () => ExecuteFunctionCommand("PU") },
        { "Audit", () => ExecuteFunctionCommand("Audit") },
        { "恢复", () => ExecuteFunctionCommand("恢复") },
        #endregion

        #region 公共图按键
        { "共用条件", () => ExecuteCommonCommand("共用条件", "共用条件说明", "TJ(共用条件)", 161) },
        { "所有条件开关", () => ExecuteCommonCommand("所有条件开关", "所有条件开关", "TJ(共用条件)", 161) },
        { "设备开关", () => ExecuteCommonCommand("设备开关", "设备开关", "TJ(共用条件)", 161) },
        #endregion

        #region 工艺专业
        { "纯化水", () => ExecuteProcessCommand("纯化水", "PW,DN??,??L/h", "TJ(工艺专业GY)", 40, 0) },
        { "纯蒸汽", () => ExecuteProcessCommand("纯蒸汽", "LS,DN??,??MPa,??kg/h,", "TJ(工艺专业GY)", 40, 0) },
        { "注射用水", () => ExecuteProcessCommand("注射用水", "WFI,DN??,??℃,??L/h,使用量??L/h", "TJ(工艺专业GY)", 40, 0) },
        { "氧气", () => ExecuteProcessCommand("氧气", "O2,DN??,??MPa,??L/min", "TJ(工艺专业GY)", 40, 0) },
        { "氮气", () => ExecuteProcessCommand("氮气", "N2,DN??,??MPa,??L/min", "TJ(工艺专业GY)", 40, 0) },
        { "二氧化碳", () => ExecuteProcessCommand("二氧化碳", "CO2,DN??,??MPa,??L/min", "TJ(工艺专业GY)", 40, 0) },
        { "无菌压缩空气", () => ExecuteProcessCommand("无菌压缩空气", "CA,DN??,??MPa,??L/min", "TJ(工艺专业GY)", 40, 0) },
        { "仪表压缩空气", () => ExecuteProcessCommand("仪表压缩空气", "IA,DN??,??MPa,??L/min", "TJ(工艺专业GY)", 40, 0) },
        { "低压蒸汽", () => ExecuteProcessCommand("低压蒸汽", "LS,DN??,??MPa,??kg/h,", "TJ(工艺专业GY)", 40, 0) },
        { "低温循环上水", () => ExecuteProcessCommand("低温循环上水", "RWS,DN??,??m³/h", "TJ(工艺专业GY)", 40, 0) },
        { "常温循环上水", () => ExecuteProcessCommand("常温循环上水", "CWS,DN??,??m³/h", "TJ(工艺专业GY)", 40, 0) },
        { "凝结回水", () => ExecuteProcessCommand("凝结回水", "SC,DN??", "TJ(工艺专业GY)", 40, 0) },
        #endregion

        #region 建筑专业
        { "吊顶", () => ExecuteBuildingCommand("TextBox_吊顶高度", "JZTJ_吊顶", "TJ(建筑吊顶)", 30) },
        { "不吊顶", () => ExecuteBuildingCommand("不吊顶", "JZTJ_不吊顶", "TJ(建筑吊顶)", 30) },
        { "防撞护板", () => ExecuteBuildingCommand("TextBox_防撞护板", "JZTJ_防撞护板", "TJ(建筑专业J)", 30) },
        { "房间编号", () => ExecuteRoomNumberCommand() },
        { "编号检查", () => ExecuteBuildingCommand("编号检查", "JZTJ_编号检查", "TJ(建筑专业J)", 30) },
        { "冷藏库降板", () => ExecuteBuildingCommand("冷藏库降板", "冷藏库降板（270）", "TJ(建筑专业J)", 30) },
        { "冷冻库降板", () => ExecuteBuildingCommand("冷冻库降板", "冷冻库降板（390）", "TJ(建筑专业J)", 30) },
        { "特殊地面做法要求", () => ExecuteBuildingCommand("特殊地面做法要求", "JZTJ_特殊地面做法要求", "TJ(建筑专业J)", 30) },
        { "排水沟", () => ExecuteDrainageCommand() },
        { "横墙建筑开洞", () => ExecuteBuildingWallHoleCommand("横墙") },
        { "纵墙建筑开洞", () => ExecuteBuildingWallHoleCommand("纵墙") },
        #endregion

        #region 结构专业
        { "结构受力点", () => ExecuteStructureCommand("结构受力点") },
        { "水平荷载", () => ExecuteStructureCommand("水平荷载") },
        { "面着地", () => ExecuteStructureCommand("面着地") },
        { "框着地", () => ExecuteStructureCommand("框着地") },
        { "圆形开洞", () => ExecuteStructureCommand("圆形开洞") },
        { "半径开圆洞", () => ExecuteStructureCommand("半径开圆洞") },
        { "矩形开洞", () => ExecuteStructureCommand("矩形开洞") },
        { "横墙结开建筑洞", () => ExecuteStructureCommand("横墙结开建筑洞") },
        { "纵墙结开建筑洞", () => ExecuteStructureCommand("纵墙结开建筑洞") },
        #endregion

        #region 给排水专业
        { "洗眼器", () => ExecutePlumbingCommand("洗眼器") },
        { "不给饮用水", () => ExecutePlumbingCommand("不给饮用水") },
        { "小便器给水", () => ExecutePlumbingCommand("小便器给水") },
        { "大便器给水", () => ExecutePlumbingCommand("大便器给水") },
        { "洗涤盆", () => ExecutePlumbingCommand("洗涤盆") },
        { "水池给水", () => ExecutePlumbingCommand("水池给水") },
        { "横墙水开建筑洞", () => ExecutePlumbingCommand("横墙水开建筑洞") },
        { "纵墙水开建筑洞", () => ExecutePlumbingCommand("纵墙水开建筑洞") },
        #endregion

        #region 暖通专业
        { "排潮", () => ExecuteHVACCommand("排潮") },
        { "排尘", () => ExecuteHVACCommand("排尘") },
        { "排热", () => ExecuteHVACCommand("排热") },
        { "直排", () => ExecuteHVACCommand("直排") },
        { "除味", () => ExecuteHVACCommand("除味") },
        { "A级高度", () => ExecuteHVACCommand("A级高度") },
        { "设备取风量", () => ExecuteHVACCommand("设备取风量") },
        { "设备排风量", () => ExecuteHVACCommand("设备排风量") },
        { "排风百分比", () => ExecuteHVACCommand("排风百分比") },
        { "温度", () => ExecuteHVACCommand("温度") },
        { "湿度", () => ExecuteHVACCommand("湿度") },
        { "横墙暖开建洞", () => ExecuteHVACCommand("横墙暖开建洞") },
        { "纵墙暖开建洞", () => ExecuteHVACCommand("纵墙暖开建洞") },
        #endregion

        #region 电气专业
        { "无线AP", () => ExecuteElectricalCommand("无线AP") },
        { "电话插座", () => ExecuteElectricalCommand("电话插座") },
        { "网络插座", () => ExecuteElectricalCommand("网络插座") },
        { "电话网络插座", () => ExecuteElectricalCommand("电话网络插座") },
        { "安防监控", () => ExecuteElectricalCommand("安防监控") },
        { "横墙电开建洞", () => ExecuteElectricalCommand("横墙电开建洞") },
        { "纵墙电开建洞", () => ExecuteElectricalCommand("纵墙电开建洞") },
        #endregion

        #region 自控专业
        { "外线电话插座", () => ExecuteControlCommand("外线电话插座") },
        { "网络交换机", () => ExecuteControlCommand("网络交换机") },
        { "室外彩色摄像机", () => ExecuteControlCommand("室外彩色摄像机") },
        { "人像识别器", () => ExecuteControlCommand("人像识别器") },
        { "内线电话插座", () => ExecuteControlCommand("内线电话插座") },
        { "门磁开关", () => ExecuteControlCommand("门磁开关") },
        { "局域网插座", () => ExecuteControlCommand("局域网插座") },
        { "门禁控制器", () => ExecuteControlCommand("门禁控制器") },
        { "读卡器", () => ExecuteControlCommand("读卡器") },
        { "带扬声器电话机", () => ExecuteControlCommand("带扬声器电话机") },
        { "互联网插座", () => ExecuteControlCommand("互联网插座") },
        { "广角彩色摄像机", () => ExecuteControlCommand("广角彩色摄像机") },
        { "防爆型网络摄像机", () => ExecuteControlCommand("防爆型网络摄像机") },
        { "防爆型电话机", () => ExecuteControlCommand("防爆型电话机") },
        { "半球彩色摄像机", () => ExecuteControlCommand("半球彩色摄像机") },
        { "电锁按键", () => ExecuteControlCommand("电锁按键") },
        { "电控锁", () => ExecuteControlCommand("电控锁") },
        { "监控文字", () => ExecuteControlCommand("监控文字") },
        { "横墙自控开建洞", () => ExecuteControlCommand("横墙自控开建洞") },
        { "纵墙自控开建洞", () => ExecuteControlCommand("纵墙自控开建洞") },
        #endregion
    };

        #region 命令执行方法

        /// <summary>
        /// 执行方向命令
        /// 修改说明：
        ///  - 保持原有根据图元类型调整角度逻辑
        ///  - 在设置 VariableDictionary.entityRotateAngle 后，立即广播给 Command.NotifyDirectionChanged(angle)
        ///    以便任何正在运行的 EnhancedBlockPlacementJig 能即时响应并更新预览旋转
        ///  - 最后再发送插入命令启动交互插入流程
        /// </summary>
        private static void ExecuteDirectionCommand(string direction, double angle)
        {
            try
            {
                // 根据图元类型和方向设置角度
                if (VariableDictionary.btnFileName != null && VariableDictionary.btnFileName.Contains("摄像机"))
                {
                    // 摄像机 旋转
                    VariableDictionary.entityRotateAngle = DirectionButtonManager.GetAdjustedAngleForDirection(direction, true);
                }
                else
                {
                    // 其他图元 旋转
                    VariableDictionary.entityRotateAngle = DirectionButtonManager.GetAdjustedAngleForDirection(direction, false);
                }

                // 立即广播方向变化，通知可能存在的 Jig 进行实时预览旋转
                try
                {
                    // Command.NotifyDirectionChanged 会同步 VariableDictionary 并触发事件
                    Command.NotifyDirectionChanged(VariableDictionary.entityRotateAngle);
                }
                catch (Exception evEx)
                {
                    // 广播失败不应中断插入流程，记录日志并继续
                    LogManager.Instance.LogInfo($"\nNotifyDirectionChanged 触发失败: {evEx.Message}");
                }
                //string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{VariableDictionary.btnFileName}_{Guid.NewGuid():N}.dwg");
                //System.IO.File.WriteAllBytes(tempPath, VariableDictionary.resourcesFile);
                //GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.ExecuteCopyDwgAllFastWithRepeat(tempPath);
                #region 再次点方向按键的重复插入逻辑
                try
                {
                    // 非拖拽阶段：方向键 = 调角度 + 立即重复上次图元插入
                    if (!GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.IsCopyDwgAllFastDragging)
                    {
                        GB_NewCadPlus_IV.Helpers.InsertGraphicHelper.RepeatLastCopyDwgAllFastFromDirection();
                    }
                }
                catch (Exception repeatEx)
                {
                    LogManager.Instance.LogInfo($"\n方向键触发重复插入失败: {repeatEx.Message}");
                }

                #endregion
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行方向命令时出错: {ex.Message}");
            }
        }
        /// <summary>
        /// 方向按键管理器
        /// </summary>
        public static class DirectionButtonManager
        {
            /// <summary>
            /// 方向与旋转角度的映射
            /// </summary>
            public static readonly Dictionary<string, double> DirectionAngles = new Dictionary<string, double>
            {
                { "上", Math.PI * 1.5 },      // 逆时针90度
                { "右上", Math.PI * 1.25 },   // 顺时针135度
                { "右", Math.PI * 1 },        // 180度
                { "右下", Math.PI * 0.75 },   // 逆时针135度
                { "下", Math.PI * 0.5 },      // 逆时针90度
                { "左下", Math.PI * 0.25 },   // 逆时针45度
                { "左", 0 },                  // 原位置
                { "左上", Math.PI * 1.75 }    // 顺时针45度
            };

            /// <summary>
            /// 获取指定方向的旋转角度
            /// </summary>
            public static double GetAngleForDirection(string direction)
            {
                return DirectionAngles.ContainsKey(direction) ? DirectionAngles[direction] : 0;
            }

            /// <summary>
            /// 根据图元类型和方向获取旋转角度
            /// </summary>
            public static double GetAdjustedAngleForDirection(string direction, bool isCamera = false)
            {
                if (isCamera)
                {
                    // 摄像机的特殊角度处理
                    switch (direction)
                    {
                        case "上": return Math.PI * 1.5;
                        case "右上": return Math.PI * 1.25;
                        case "右": return Math.PI * 1;
                        case "右下": return Math.PI * 0.75;
                        case "下": return Math.PI * 0.5;
                        case "左下": return Math.PI * 0.25;
                        case "左": return 0;
                        case "左上": return Math.PI * 1.75;
                        default: return 0;
                    }
                }
                else
                {
                    switch (direction)
                    {
                        case "上": return Math.PI * 0;
                        case "右上": return Math.PI * 1.75;
                        case "右": return Math.PI * 1.5;
                        case "右下": return Math.PI * 1.25;
                        case "下": return Math.PI * 1;
                        case "左下": return Math.PI * 0.75;
                        case "左": return Math.PI * 0.5;
                        case "左上": return Math.PI * 0.25;
                        default: return 0;
                    }
                    //return GetAngleForDirection(direction);
                }
            }
        }

        /// <summary>
        /// 执行功能命令
        /// </summary>
        private static void ExecuteFunctionCommand(string functionName)
        {
            try
            {
                switch (functionName)
                {
                    case "查找":
                        // 实现查找功能
                        break;
                    case "PU":
                        Env.Document.SendStringToExecute("pu ", false, false, false);
                        break;
                    case "Audit":
                        Env.Document.SendStringToExecute("audit\n y ", false, false, false);
                        break;
                    case "恢复":
                        Env.Document.SendStringToExecute("DRAWINGRECOVERY ", false, false, false);
                        break;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行功能命令时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行公共命令
        /// </summary>
        private static void ExecuteCommonCommand(string commandName, string fileName, string layerName, int colorIndex)
        {
            try
            {
                VariableDictionary.entityRotateAngle = 0;
                VariableDictionary.btnFileName = fileName;
                VariableDictionary.btnBlockLayer = layerName;
                VariableDictionary.layerColorIndex = colorIndex;
                Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行公共命令时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行工艺命令
        /// </summary>
        private static void ExecuteProcessCommand(string commandName, string fileName, string layerName, int colorIndex, double angle)
        {
            try
            {
                VariableDictionary.entityRotateAngle = angle;
                VariableDictionary.btnFileName = fileName;
                VariableDictionary.btnBlockLayer = layerName;
                VariableDictionary.layerColorIndex = colorIndex;
                Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行工艺命令时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行建筑吊顶命令
        /// </summary>
        private static void ExecuteBuildingCommand(string commandName, string fileName, string layerName, int colorIndex)
        {
            try
            {
                VariableDictionary.entityRotateAngle = 0;
                VariableDictionary.btnFileName = fileName;
                VariableDictionary.btnBlockLayer = layerName;
                VariableDictionary.layerColorIndex = colorIndex;
                //VariableDictionary.winFormDiaoDingHeight = textBox_吊顶高文字.Text;

                if (commandName == "TextBox_吊顶高度" || commandName == "不吊顶")
                {
                    if (!VariableDictionary.winForm_Status)
                    {
                        VariableDictionary.wpfDiaoDingHeight = UnifiedUIManager.GetTextBoxValue(commandName);
                    }

                    Env.Document.SendStringToExecute("Line2Polyline ", false, false, false);
                }
                else if (commandName == "TextBox_防撞护板")
                {
                    VariableDictionary.textbox_Gap = Convert.ToDouble(UnifiedUIManager.GetTextBoxValue(commandName));
                    Env.Document.SendStringToExecute("ParallelLines ", false, false, false);
                }
                else
                {
                    Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行建筑命令时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行房间编号命令
        /// </summary>
        private static void ExecuteRoomNumberCommand()
        {
            try
            {
                if (!VariableDictionary.winForm_Status)
                {
                    VariableDictionary.entityRotateAngle = 0;
                    VariableDictionary.btnFileName = null;
                    // 从统一界面管理器获取房间编号信息
                    string floorNo = UnifiedUIManager.GetTextBoxValue("TextBox_楼层");
                    string cleanArea = UnifiedUIManager.GetTextBoxValue("TextBox_洁净区");
                    string systemArea = UnifiedUIManager.GetTextBoxValue("TextBox_系统区");
                    string roomSubNo = UnifiedUIManager.GetTextBoxValue("TextBox_房间号");
                    VariableDictionary.btnFileName = $"{floorNo}-{cleanArea}{systemArea}{roomSubNo}";
                    VariableDictionary.btnBlockLayer = "TJ(房间编号)";

                    if (floorNo == "1" && cleanArea == "1" && systemArea == "1")
                    {
                        VariableDictionary.layerColorIndex = 64;
                        VariableDictionary.jjqInt = 1;
                        VariableDictionary.xtqInt = 1;
                    }
                    else if (Convert.ToInt32(cleanArea) != VariableDictionary.jjqInt)
                    {
                        var layerColorTest = VariableDictionary.jjqLayerColorIndex[Convert.ToInt32(cleanArea)];
                        VariableDictionary.layerColorIndex = Convert.ToInt16(layerColorTest);//设置为被插入的图层颜色
                        VariableDictionary.jjqInt = Convert.ToInt32(cleanArea);
                    }
                    else if (Convert.ToInt32(systemArea) != VariableDictionary.xtqInt)
                    {
                        var layerColorTest = VariableDictionary.xtqLayerColorIndex[Convert.ToInt32(systemArea)];
                        VariableDictionary.layerColorIndex = Convert.ToInt16(layerColorTest);//设置为被插入的图层颜色
                        VariableDictionary.xtqInt = Convert.ToInt32(systemArea);
                    }
                    Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);

                    // 更新房间号
                    UpdateRoomNumber(roomSubNo);
                }
                Env.Document.SendStringToExecute("DBTextLabel ", false, false, false);
                VariableDictionary.winForm_Status = false;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行房间编号命令时出错: {ex.Message}");
            }
        }
        /// <summary>
        /// 更新房间号
        /// </summary>
        private static void UpdateRoomNumber(string currentNumber)
        {
            try
            {
                if (int.TryParse(currentNumber, out int number))
                {
                    string newRoomNumber;
                    if (number < 9)
                    {
                        newRoomNumber = $"0{number + 1}";
                    }
                    else
                    {
                        newRoomNumber = (number + 1).ToString();
                    }

                    // 通过统一界面管理器更新房间号TextBox的值
                    UnifiedUIManager.SetWpfTextBoxValue("TextBox_房间号", newRoomNumber);

                }
            }
            catch
            {
                // 忽略转换错误
            }
        }
        /// <summary>
        /// 执行排水沟命令
        /// </summary>
        private static void ExecuteDrainageCommand()
        {
            try
            {
                // 从统一界面管理器获取排水沟尺寸
                if (!VariableDictionary.winForm_Status)
                {
                    VariableDictionary.dimString_JZ_宽 = Convert.ToDouble(UnifiedUIManager.GetTextBoxValue("textBox_排水沟_宽.Text"));
                    VariableDictionary.dimString_JZ_深 = Convert.ToDouble(UnifiedUIManager.GetTextBoxValue("textBox_排水沟_深.Text"));
                }
                VariableDictionary.entityRotateAngle = 0;
                VariableDictionary.btnFileName = "JZTJ_排水沟";
                VariableDictionary.buttonText = "JZTJ_排水沟";
                //VariableDictionary.btnFileName_blockName = "$TWTSYS$00000508";
                VariableDictionary.btnBlockLayer = "TJ(建筑专业J)";
                VariableDictionary.layerColorIndex = 30;//设置为被插入的图层颜色
                VariableDictionary.layerName = VariableDictionary.btnBlockLayer;
                Env.Document.SendStringToExecute("Rec2PolyLine_3 ", false, false, false);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行排水沟命令时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行建筑墙体开洞命令
        /// </summary>
        private static void ExecuteBuildingWallHoleCommand(string wallType)
        {
            try
            {

                VariableDictionary.btnBlockLayer = "TJ(建筑专业J)";                

                VariableDictionary.buttonText = $"JZTJ_{wallType}开洞";
                if (wallType == "横墙")
                {
                    // 从统一界面管理器获取加宽值
                    VariableDictionary.textbox_Width = Convert.ToDouble(UnifiedUIManager.GetTextBoxValue("TextBox_TJ建筑孔洞左右宽"));
                    VariableDictionary.textbox_Height = 15;
                }
                else if (wallType == "纵墙")
                {
                    VariableDictionary.textbox_Width = 15;
                    VariableDictionary.textbox_Height = Convert.ToDouble(UnifiedUIManager.GetTextBoxValue("TextBox_TJ建筑孔洞上下高"));
                }

                VariableDictionary.layerColorIndex = 64;
                VariableDictionary.btnFileName = "建筑洞口：";
                Env.Document.SendStringToExecute("Rec2PolyLine_N ", false, false, false);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行建筑墙体开洞命令时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行结构命令
        /// </summary>
        private static void ExecuteStructureCommand(string commandName)
        {
            // 实现结构专业命令
            try
            {
                // 根据命令名称执行相应操作
                switch (commandName)
                {
                    case "结构受力点":
                        // 实现结构受力点逻辑
                        break;
                    case "水平荷载":
                        // 实现水平荷载逻辑
                        break;
                        // ... 其他结构命令
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行结构命令时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行给排水命令
        /// </summary>
        private static void ExecutePlumbingCommand(string commandName)
        {
            // 实现给排水专业命令
            try
            {
                // 根据命令名称执行相应操作
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行给排水命令时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行暖通命令
        /// </summary>
        private static void ExecuteHVACCommand(string commandName)
        {
            // 实现暖通专业命令
            try
            {
                // 根据命令名称执行相应操作
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行暖通命令时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行电气命令
        /// </summary>
        private static void ExecuteElectricalCommand(string commandName)
        {
            // 实现电气专业命令
            try
            {
                // 根据命令名称执行相应操作
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行电气命令时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行自控命令
        /// </summary>
        private static void ExecuteControlCommand(string commandName)
        {
            // 实现自控专业命令
            try
            {
                // 根据命令名称执行相应操作
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"执行自控命令时出错: {ex.Message}");
            }
        }

        #endregion

        /// <summary>
        /// 获取命令
        /// </summary>
        public static Action GetCommand(string commandName)
        {
            return _commandMap.ContainsKey(commandName) ? _commandMap[commandName] : null;
        }

        /// <summary>
        /// 检查是否是预定义命令
        /// </summary>
        public static bool IsPredefinedCommand(string commandName)
        {
            return _commandMap.ContainsKey(commandName);
        }

        /// <summary>
        /// 添加自定义命令
        /// </summary>
        public static void AddCustomCommand(string commandName, Action command)
        {
            if (!_commandMap.ContainsKey(commandName))
            {
                _commandMap[commandName] = command;
            }
        }


    }
}
