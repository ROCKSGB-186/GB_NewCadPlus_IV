using System;
using System.Collections.Generic;
using System.Linq;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 插入前属性页面使用的 Tag 中文注释字典。
    /// </summary>
    public static class PropertyTagCommentDictionary
    {
        // 使用不区分大小写的字典，兼容 DWG 中 Tag 大小写差异。
        private static readonly IReadOnlyDictionary<string, string> Comments =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // 基础识别属性。
                ["NAME"] = "名称",
                ["TYPE_CODE"] = "类型代码",
                ["TAG_NO"] = "位号",
                ["TAGNO"] = "位号",
                ["QTY"] = "数量",
                ["MODEL"] = "型号",
                ["SW_MODEL"] = "三维模型型号",
                ["SYSTEM"] = "所属系统",
                ["REMARK"] = "备注",

                // 管道和连接属性。
                ["PIPELINETITLE"] = "管道标题",
                ["PIPE_LENGTH"] = "管道长度",
                ["PIPELENGTH"] = "管道长度",
                ["PIPE_OD"] = "管道外径",
                ["PIPEID"] = "管道内径",
                ["PIPE_THK"] = "管道壁厚",
                ["PIPETHK"] = "管道壁厚",
                ["PIPE_MATL"] = "管道材质",
                ["PIPEMATL"] = "管道材质",
                ["PIPE_TYPE"] = "管道类型",
                ["PIPETYPE"] = "管道类型",
                ["PIPE_CLASS"] = "管道等级",
                ["PIPECLASS"] = "管道等级",
                ["PIPE_ROLE"] = "管道角色",
                ["PIPEROLE"] = "管道角色",
                ["CONN_TYPE"] = "连接方式",
                ["DNCONN_TYPE"] = "公称直径连接方式",
                ["CONNECTION_MODE"] = "连接形式",
                ["STARTPOINT"] = "起点",
                ["ENDPOINT"] = "终点",
                ["DN"] = "公称直径",
                ["PN"] = "公称压力",
                ["CLASS"] = "压力等级",
                ["MEDIUM"] = "介质",
                ["SCHEDULE"] = "管道壁厚等级",
                ["WORKPRESSURE"] = "工作压力",
                ["DESIGNPRESSURE"] = "设计压力",
                ["WORKTEMP"] = "工作温度",
                ["DESIGNTEMP"] = "设计温度",
                ["TEMP_RANGE"] = "温度范围",
                ["FLOWRATE"] = "流量",
                ["FLOWVEL"] = "流速",
                ["MAX_FLOW_VEL"] = "最大流速",

                // 标准和法兰属性。
                ["FLG_STD"] = "法兰标准",
                ["FLG_TYPE"] = "法兰类型",
                ["FACE_TYPE"] = "密封面形式",
                ["SERIES"] = "系列",
                ["TABLE_NUMBER"] = "表号",
                ["DRAWINGNO.STANDARDNO"] = "图号或标准号",
                ["TEST_STD"] = "试验标准",
                ["STRUCT_LEN_STD"] = "结构长度标准",
                ["FLG_OD"] = "法兰外径",
                ["FLG_ID"] = "法兰内径",
                ["FLG_THK"] = "法兰厚度",
                ["FLG_QTY"] = "法兰数量",
                ["BOLT_QTY"] = "螺栓数量",
                ["BOLT_LENGTH"] = "螺栓长度",
                ["BOLT_HOLES"] = "螺栓孔数量",
                ["BOLT_HOLE_DIA"] = "螺栓孔直径",
                ["BOLT_PCD"] = "螺栓孔中心圆直径",
                ["BOLT_SPEC"] = "螺栓规格",
                ["RAISED_FACE_HGT"] = "密封面高度",

                // 阀门和部件属性。
                ["DRIVE_MODE"] = "驱动方式",
                ["HANDWHEEL"] = "手轮",
                ["STRUCT_TYPE"] = "结构形式",
                ["SEAL_TYPE"] = "密封形式",
                ["BODY_MATL"] = "阀体材质",
                ["DISC_MATL"] = "阀板材质",
                ["STEM_MATL"] = "阀杆材质",
                ["SEAL_MATL"] = "密封材料",
                ["LINING_MATL"] = "衬里材料",
                ["LINING_THK"] = "衬里厚度",
                ["LINING_PROC"] = "衬里工艺",
                ["LEAK_CLASS"] = "泄漏等级",
                ["WEIGHT"] = "重量",
                ["MATERIAL"] = "材质",

                // 设计和附加属性。
                ["PIPE_SPEC"] = "管道规格",
                ["PIPE_SPECIFICATION"] = "管道设计规范",
                ["PRESSURE_LOSS"] = "压力损失",
                ["ROUGHNESS"] = "粗糙度",
                ["INSUL_MATL"] = "保温材料",
                ["INSUL_THK"] = "保温厚度",
                ["INSUL_TYPE"] = "保温类型",
                ["HEATTRACE"] = "伴热",
                ["HEATTRACEPOWER"] = "伴热功率",

                // 塔器和脱硫系统属性。
                ["TOWER_DIA"] = "塔体直径（mm）",
                ["TOWER_HGT"] = "塔体总高度（mm）",
                ["TOWER_MATL"] = "塔体材质",
                ["LINING_TYPE"] = "内防腐层类型",
                ["SPRAY_LAYER"] = "喷淋层数",
                ["WORK_PRESSURE"] = "工作压力（MPaG）",
                ["WORK_TEMP"] = "工作温度（℃）",
                ["GAS_FLOW"] = "单台脱硫装置标况烟气量（Nm³/h）",
                ["GAS_FLOW_WET"] = "单套净化系统出口工况烟气量（m³/h）",
                ["GAS_FLOW_DRY"] = "单套净化系统出口标况烟气量（Nm³/h）",
                ["SO2_IN"] = "入口烟气 SO₂ 浓度（mg/Nm³）",
                ["SO2_OUT"] = "出口烟气 SO₂ 浓度（mg/Nm³）",
                ["DESULF_EFF"] = "脱硫效率（%）",
                ["EXHAUST_TEMP"] = "脱硫系统出口温度（℃）",
                ["ATM_PRESSURE"] = "当地大气压（kPa）",
                ["STACK_EXIT_VEL"] = "烟囱出口风速（m/s）",
                ["STACK_EXIT_DIA"] = "烟囱出口直径（m）",
                ["CA_S_RATIO"] = "钙硫比 Ca/S（摩尔比）",
                ["L_G_RATIO"] = "液气比（L/Nm³）",
                ["SO2_MOL"] = "单套净化系统 SO₂ 摩尔数（kmol/h）",
                ["INNER_STRUCT"] = "塔内结构型式",
                ["INTERNALS_MATL"] = "内构件材质",
                ["OUTER_COATING"] = "外部防腐涂层",
                ["COATING_THK"] = "涂层厚度（μm）",
                ["NOZZ_TYPE"] = "喷嘴类型",
                ["NOZZ_MATL"] = "喷嘴材质",
                ["NOZZ_FLOW"] = "喷嘴流量（L/min@2bar）",
                ["NOZZ_ANGLE"] = "喷嘴雾化角（°）",
                ["SPRAY_COVERAGE"] = "单层喷淋覆盖率（%）",
                ["DEMISTER_TYPE"] = "除雾器类型",
                ["DEMISTER_MATL"] = "除雾器材质",
                ["DEMISTER_STAGES"] = "除雾器级数",
                ["DEMISTER_OUT"] = "除雾器出口雾滴含量（mg/Nm³）",
                ["CIRC_PUMP_CFG"] = "循环泵配置",
                ["TOWER_DP"] = "塔压降（Pa）",

                // 图纸和通用设计属性。
                ["DRAWINGNO"] = "图号",
                ["DESIGN_PRESSURE"] = "设计压力",
                ["DESIGN_TEMP"] = "设计温度（℃）",
                ["DESIGN_PRESS"] = "设计压力（Pa）",
                ["DESIGN_STD"] = "设计标准",
                ["MATING_FLG"] = "配对法兰",
                ["GASKET_MATL"] = "垫片材质",
                ["GASKET_THK"] = "垫片厚度（mm）",
                ["BOLT_MATL"] = "螺栓材质",
                ["NUT_MATL"] = "螺母材质",
                ["FLG_MATL"] = "法兰材质",

                // 搅拌器属性。
                ["IMPELLER_TYPE"] = "叶轮型式",
                ["IMPELLER_LAYER"] = "叶轮层数",
                ["ROTATION_SPEED"] = "搅拌转速（r/min）",
                ["FLOW_DIR"] = "介质流动方向",
                ["REDUCER_TYPE"] = "减速机类型",
                ["BEARING_TYPE"] = "轴承型式",
                ["COUPLING_TYPE"] = "联轴器型式",
                ["INSTALL_ANGLE"] = "安装角度",
                ["IMMERSION_DEPTH"] = "伸入长度/浸没深度（mm）",
                ["SEAL_LIFE"] = "机械密封使用寿命（年）",
                ["MEDIUM_DENSITY"] = "介质比重（kg/m³）",
                ["SOLID_CONTENT"] = "介质含固量（%）",
                ["MEDIUM_TEMP"] = "介质温度（℃）",
                ["CHLORIDE_CONTENT"] = "氯离子含量（ppm）",
                ["MOTOR_IP"] = "电机防护等级",
                ["MOTOR_INSUL"] = "电机绝缘等级",
                ["MOTOR_EFF"] = "电机能效等级",
                ["POWER_VOLTAGE"] = "电源电压",
                ["ONLINE_MAINT"] = "在线维护功能",
                ["FLUSH_PORT"] = "冲洗管口",

                // 烟囱、烟道和建筑安全属性。
                ["STACK_TYPE"] = "烟囱型式",
                ["INNER_TYPE"] = "内筒型式",
                ["GAS_VEL"] = "烟囱流速（m/s）",
                ["FOUNDATION_TYPE"] = "基础形式",
                ["LIGHTNING"] = "避雷针",
                ["AVIATION_LAMP"] = "航标灯",
                ["PLATFORM_SPAN"] = "检修平台间距（m）",
                ["SAMPLING_HOLE"] = "烟气取样孔",
                ["MANHOLE_DIA"] = "人孔直径（mm）",
                ["CORROSION_CLASS"] = "腐蚀性等级",
                ["DUCT_SHAPE"] = "烟道截面形状",
                ["DUCT_AREA"] = "烟道截面积（m²）",
                ["HEAT_TRACE"] = "伴热措施",
                ["REINF_TYPE"] = "加固肋型式",
                ["CORROSION_ALLOW"] = "腐蚀裕量（mm）",
                ["HOT\\SOUND_ISOLACODE"] = "隔热隔声代号",
                ["HOTSOUND_ISOLACODE"] = "隔热隔声代号",
                ["IS_ANTICORRO"] = "是否防腐",

                // 泵及泵组属性。
                ["PUMP_TYPE"] = "泵型式",
                ["PUMP_STRUCT"] = "泵结构形式",
                ["IMPELLER_DIA"] = "叶轮直径（mm）",
                ["PUMP_SPEED"] = "泵转速（r/min）",
                ["NPSHR"] = "汽蚀余量（m）",
                ["PUMP_EFF"] = "泵效率（%）",
                ["PUMP_BEFORE_AFTER"] = "泵前、后",

                // 放空、排放及安全附件属性。
                ["VENT_HEIGHT"] = "放空管口高度（m）",
                ["EXIT_VEL"] = "排放口流速（马赫数 Ma）",
                ["DESIGN_FLOW"] = "设计放空量（m³/h）",
                ["VENT_PRESSURE"] = "放空操作压力（MPa）",
                ["OPER_TEMP"] = "放空操作温度（℃）",
                ["BIRD_SCREEN"] = "防鸟网",
                ["LOW_TEMP_COND"] = "低温低应力工况",
                ["INNER_COATING"] = "内防腐涂层",
                ["DRAIN_HOLE"] = "排液孔",
                ["PLATFORM_CLEAR"] = "平台/建筑物安全间距（m）",

                // 管道扩展属性。
                ["START_POINT"] = "起始点",
                ["END_POINT"] = "终点",
                ["FLOW_VEL"] = "设计流速（m/s）",
                ["FLOW_RATE"] = "介质流量（m³/h）",
                ["PIPE_WEIGHT"] = "管道计算重量（kg/m）",
                ["PIPE_COATING"] = "外部防腐涂层",
                ["NDT_RATIO"] = "无损检测比例（%）",
                ["NDT_TYPE"] = "无损检测方法",
                ["MAX_BENDING"] = "允许弯曲半径（m）",
                ["EXPANSION"] = "热膨胀量（mm/100m）",
                ["CLEANING_REQ"] = "清洁度要求",

                // 软接头、阀门和安全阀扩展属性。
                ["LENGTH_L"] = "产品长度 L（mm）",
                ["RUBBER_TYPE"] = "橡胶类型",
                ["BALL_STRUCT"] = "球体结构",
                ["REINF_LAYER"] = "增强层材质",
                ["PRESSURE_RING"] = "压力环材质",
                ["AXIAL_EXTEND"] = "轴向伸长量（mm）",
                ["AXIAL_COMPRESS"] = "轴向压缩量（mm）",
                ["LATERAL_DISP"] = "横向位移量（mm）",
                ["ANGLE_OFFSET"] = "偏转角度（°）",
                ["FLOW_DIA"] = "流道直径（mm）",
                ["SET_PRESSURE"] = "整定压力（MPa）",
                ["SPRING_MATL"] = "弹簧材质",
                ["BACK_PRESSURE"] = "背压类型"
            };

        /// <summary>
        /// 按 Tag 获取中文注释；未知 Tag 返回 Tag 本身，保证动态规范字段仍可识别。
        /// </summary>
        public static string GetComment(string tag)
        {
            // 清理空白并兼容空 Tag。
            string value = tag?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;

            // 先按原始 Tag 查找，保留点号、下划线等业务分隔符。
            if (Comments.TryGetValue(value, out string comment)) return comment;

            // 再按去除分隔符的结果查找，例如 PIPE_LENGTH 和 PIPELENGTH 可互相兼容。
            string normalized = Normalize(value);
            KeyValuePair<string, string> match = Comments.FirstOrDefault(
                item => string.Equals(Normalize(item.Key), normalized, StringComparison.OrdinalIgnoreCase));
            return string.IsNullOrWhiteSpace(match.Key) ? value : match.Value;
        }

        // 归一化 Tag，仅用于翻译匹配，不改变实际回写 Tag。
        private static string Normalize(string value)
        {
            return new string((value ?? string.Empty)
                .Where(character => char.IsLetterOrDigit(character))
                .ToArray());
        }
    }
}
