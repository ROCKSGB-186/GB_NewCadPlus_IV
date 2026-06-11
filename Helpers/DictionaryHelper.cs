using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.Helpers
{
    internal static class DictionaryHelper
    {
        /// <summary>
        /// 预定义的英文对照表
        /// </summary>
        public static readonly Dictionary<string, string> ChineseToEnglish = new Dictionary<string, string>
           {
            // 基础物资/材料（带单位）
            { "管子", "Pipe(m)" },
            { "阀门", "Valve(Pcs.)" },
            { "法兰", "Flange(Pcs.)" },
            { "垫片", "Gasket(Pcs.)" },
            { "螺栓", "Bolt(Pcs.)" },
            { "螺母", "Nut(Pcs.)" },
            { "名称", "Name" },
            { "介质名称", "Medium Name" },
            { "规格", "Specifications" },
            { "材料", "Material" },
            { "数量", "Quan." },
            { "图号或标准号", "DWG.No./ STD.No." },
            { "功率", "Power" },
            { "容积", "Volume" },
            { "压力", "Pressure" },
            { "温度", "Temperature" },
            { "直径", "Diameter" },
            { "长度", "Length" },
            { "厚度", "Thickness" },
            { "重量", "Weight" },
            { "型号", "Model" },
            { "隔热隔声代号", "Code" },
            { "是否防腐", "Antisepsis" },
            { "操作压力", "OperatingPressure" },
            { "备注", "Remarks" },
          
            // 文件/通用属性
            { "文件ID", "Id" },
            { "文件名", "FileName" },
            { "分类ID", "CategoryId" },
            { "属性业务ID", "FileAttributeId" },
            { "存储文件名", "FileStoredName" },
            { "显示名称", "DisplayName" },
            { "文件类型", "FileType" },
            { "文件哈希", "FileHash" },
            { "元素块名", "BlockName" },
            { "图层名称", "LayerName" },
            { "比例", "Scale" },
            { "颜色索引", "ColorIndex" },
            { "文件路径", "FilePath" },
            { "预览图片名", "PreviewImageName" },
            { "预览图片路径", "PreviewImagePath" },
            { "文件大小", "FileSize" },
            { "是否预览", "IsPreview" },
            { "版本号", "Version" },
            { "描述", "Description" },
            { "创建时间", "CreatedAt" },
            { "更新时间", "UpdatedAt" },
            { "分类类型", "CategoryType" },
            { "创建者", "CreatedBy" },
            { "是否激活", "IsActive" },
            { "标题", "Title" },
            { "关键字", "Keywords" },
            { "是否公开", "IsPublic" },
            { "更新者", "UpdatedBy" },
          
            // FileAttribute 属性
            { "存储文件ID", "FileStorageId" },
            { "宽度", "Width" },
            { "高度", "Height" },
            { "角度", "Angle" },
            { "基点X", "BasePointX" },
            { "基点Y", "BasePointY" },
            { "基点Z", "BasePointZ" },
            { "介质", "MediumName" },
            { "材质", "Material" },
            { "标准号", "StandardNumber" },
            { "外径", "OuterDiameter" },
            { "内径", "InnerDiameter" },
            { "自定义1", "Customize1" },
            { "自定义2", "Customize2" },
            { "自定义3", "Customize3" },
          
            // 新增或替换的属性
            { "属性分组", "AttributeGroup" },
            { "压力等级", "PressureRating" },
            { "操作温度", "OperatingTemperature" },
            { "公称直径DN", "NominalDiameter" },
            { "密度", "Density" },
            { "流量", "Flow" },
            { "流速", "Velocity" },
            { "扬程", "Lift" },
            { "电压", "Voltage" },
            { "电流", "Current" },
            { "频率", "Frequency" },
            { "电导率", "Conductivity" },
            { "含湿量", "Moisture" },
            { "湿度", "Humidity" },
            { "真空度", "Vacuum" },
            { "辐射量", "Radiation" },
          
            // 管道相关
            { "管道规格", "PipeSpec" },
            { "管道公称直径", "PipeNominalDiameter" },
            { "管道壁厚", "PipeWallThickness" },
            { "管道压力等级", "PipePressureClass" },
            { "连接方式", "ConnectionType" },
            { "管道坡度", "PipeSlope" },
            { "防腐处理", "AnticorrosionTreatment" },
          
            // 阀门详细
            { "阀门型号", "ValveModel" },
            { "阀体材质", "ValveBodyMaterial" },
            { "阀板材质", "ValveDiscMaterial" },
            { "球体材质", "ValveBallMaterial" },
            { "密封材质", "SealMaterial" },
            { "传动方式", "DriveType" },
            { "开启方式", "OpenMode" },
            { "适用介质", "ApplicableMedium" },
          
            // 法兰详细
            { "法兰型号", "FlangeModel" },
            { "法兰类型", "FlangeType" },
            { "密封面形式", "FlangeFaceType" },
            { "法兰标准", "FlangeStandard" },
            { "螺栓规格", "BoltSpec" },
          
            // 泵
            { "泵型号", "PumpModel" },
            { "泵流量", "PumpFlow" },
            { "泵扬程", "PumpHead" },
            { "泵体材质", "PumpBodyMaterial" },
            { "电机功率", "MotorPower" },
            { "进出口直径", "InletOutletDiameter" },
            { "额定转速", "RatedSpeed" },
            { "防护等级", "ProtectionLevel" },
          
            // 烟气/环保
            { "处理烟气量", "FlueGasCapacity" },
            { "脱硫效率", "DesulfurizationEfficiency" },
            { "液滴粒径", "DropletSize" },
            { "喷淋层数", "SprayLayerCount" },
            { "烟囱规格", "ChimneySpec" },
            { "烟囱直径", "ChimneyDiameter" },
            { "烟囱高度", "ChimneyHeight" },
            { "烟囱材质", "ChimneyMaterial" },
            { "烟囱壁厚", "ChimneyThickness" },
            { "出口风速", "OutletWindSpeed" },
            { "保温层厚度", "InsulationThickness" },
            { "支撑方式", "SupportType" },
            { "烟气温度", "FlueGasTemperature" },
          
            // 仪表
            { "压力表型号", "PressureGaugeModel" },
            { "温度计型号", "ThermometerModel" },
            { "过滤器型号", "FilterModel" },
            { "止回阀型号", "CheckValveModel" },
            { "喷淋头型号", "SprinklerModel" },
            { "流量计型号", "FlowMeterModel" },
            { "安全阀型号", "SafetyValveModel" },
            { "柔性接头型号", "FlexibleJointModel" },
          
            // ========== 以下为补充的常见词汇 ==========
            // 阀门类型
            { "闸阀", "GateValve" },
            { "截止阀", "GlobeValve" },
            { "球阀", "BallValve" },
            { "蝶阀", "ButterflyValve" },
            { "止回阀", "CheckValve" },
            { "安全阀", "SafetyValve" },
            { "减压阀", "PressureReducingValve" },
            { "调节阀", "ControlValve" },
            { "隔膜阀", "DiaphragmValve" },
            { "旋塞阀", "PlugValve" },
            { "针型阀", "NeedleValve" },
          
            // 法兰类型
            { "平焊法兰", "SlipOnFlange" },
            { "对焊法兰", "WeldNeckFlange" },
            { "松套法兰", "LapJointFlange" },
            { "螺纹法兰", "ThreadedFlange" },
            { "盲板法兰", "BlindFlange" },
            { "承插焊法兰", "SocketWeldFlange" },
            { "整体法兰", "IntegralFlange" },
            { "法兰盖", "BlankFlange" },
          
            // 管件
            { "弯头", "Elbow" },
            { "长半径弯头", "LongRadiusElbow" },
            { "短半径弯头", "ShortRadiusElbow" },
            { "45度弯头", "45DegreeElbow" },
            { "90度弯头", "90DegreeElbow" },
            { "三通", "Tee" },
            { "等径三通", "StraightTee" },
            { "异径三通", "ReducingTee" },
            { "四通", "Cross" },
            { "同心异径管", "ConcentricReducer" },
            { "偏心异径管", "EccentricReducer" },
            { "管帽", "Cap" },
            { "管堵", "Plug" },
            { "活接头", "Union" },
            { "补芯", "Bushing" },
            { "短节", "Nipple" },
            { "大小头", "Reducer" },
          
            // 设备
            { "泵", "Pump" },
            { "离心泵", "CentrifugalPump" },
            { "往复泵", "ReciprocatingPump" },
            { "齿轮泵", "GearPump" },
            { "螺杆泵", "ScrewPump" },
            { "风机", "Fan" },
            { "鼓风机", "Blower" },
            { "压缩机", "Compressor" },
            { "换热器", "HeatExchanger" },
            { "管壳式换热器", "ShellAndTubeHeatExchanger" },
            { "板式换热器", "PlateHeatExchanger" },
            { "塔器", "Tower" },
            { "精馏塔", "DistillationTower" },
            { "吸收塔", "AbsorptionTower" },
            { "洗涤塔", "Scrubber" },
            { "反应器", "Reactor" },
            { "储罐", "StorageTank" },
            { "立式储罐", "VerticalTank" },
            { "卧式储罐", "HorizontalTank" },
            { "球罐", "SphericalTank" },
          
            // 仪表与电气
            { "压力表", "PressureGauge" },
            { "温度计", "Thermometer" },
            { "双金属温度计", "BimetallicThermometer" },
            { "热电偶", "Thermocouple" },
            { "热电阻", "RTD" },
            { "流量计", "FlowMeter" },
            { "电磁流量计", "ElectromagneticFlowMeter" },
            { "涡街流量计", "VortexFlowMeter" },
            { "质量流量计", "MassFlowMeter" },
            { "液位计", "LevelGauge" },
            { "磁翻板液位计", "MagneticLevelGauge" },
            { "雷达液位计", "RadarLevelGauge" },
            { "变送器", "Transmitter" },
            { "压力变送器", "PressureTransmitter" },
            { "差压变送器", "DifferentialPressureTransmitter" },
            { "温度变送器", "TemperatureTransmitter" },
            { "执行器", "Actuator" },
            { "电动执行器", "ElectricActuator" },
            { "气动执行器", "PneumaticActuator" },
          
            // 标准与材料
            { "国标", "GBStandard" },
            { "美标", "ANSIStandard" },
            { "德标", "DINStandard" },
            { "日标", "JISStandard" },
            { "碳钢", "CarbonSteel" },
            { "不锈钢", "StainlessSteel" },
            { "304不锈钢", "StainlessSteel304" },
            { "316L不锈钢", "StainlessSteel316L" },
            { "铸铁", "CastIron" },
            { "铸钢", "CastSteel" },
            { "青铜", "Bronze" },
            { "黄铜", "Brass" },
            { "聚四氟乙烯", "PTFE" },
            { "橡胶", "Rubber" },
          
            // 紧固件/密封
            { "垫圈", "Washer" },
            { "金属垫片", "MetalGasket" },
            { "非金属垫片", "NonMetallicGasket" },
            { "缠绕垫片", "SpiralWoundGasket" },
            { "密封圈", "SealRing" },
            { "卡箍", "Clamp" },
          
            // 工艺参数
            { "公称压力", "NominalPressure" },
            { "工作压力", "WorkingPressure" },
            { "设计压力", "DesignPressure" },
            { "试验压力", "TestPressure" },
            { "工作温度", "WorkingTemperature" },
            { "设计温度", "DesignTemperature" },
            { "介质流向", "FlowDirection" },
            { "吹扫", "Purge" },
            { "保温", "Insulation" },
            { "伴热", "HeatTracing" }
         };

        /// <summary>
        /// 属性同义词映射与过滤规则
        /// - AttributeSynonyms: 将常见属性名映射为标准列名，便于把诸如 "阀体材料" 写入 "材料" 列
        /// - ExcludedAttributeSubstrings: 出现在属性名中的子串如果匹配到则排除该属性（不生成列）
        /// </summary>
        public static readonly Dictionary<string, string> AttributeSynonyms = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // 材料类
            { "阀体材料", "材料" },
            { "阀体材质", "材料" },
            { "材质", "材料" },

            // 规格/型号
            { "规格型号", "规格" },
            { "型号", "规格" },
            { "规格", "规格" },

            // 长度类
            { "长度(m)", "管道长度(mm)" },
            { "管道长度(mm)", "管道长度(mm)" },
            { "长度", "管道长度(mm)" },

            // 数量
            { "数量", "数量" },

            // 其它常见映射，可按需扩展
            { "介质", "介质名称" },
            { "介质名称", "介质名称" }
        };

        /// <summary>
        /// 添加属性名称映射字典（如果还没有的话）
        /// </summary>
        public static readonly Dictionary<string, string> _propertyDisplayNameMap = new Dictionary<string, string>
        {
            // FileStorage 属性映射
            { "Id", "文件ID" },
            { "FileName", "文件名" },
            { "CategoryId", "分类ID" },
            { "FileAttributeId", "属性业务ID" },
            { "FileStoredName", "存储文件名" },
            { "DisplayName", "显示名称" },
            { "FileType", "文件类型" },
            { "FileHash", "文件哈希" },
            { "BlockName", "元素块名" },
            { "LayerName", "图层名称" },
            { "Scale", "比例" },
            { "ColorIndex", "颜色索引" },
            { "FilePath", "文件路径" },
            { "PreviewImageName", "预览图片名" },
            { "PreviewImagePath", "预览图片路径" },
            { "FileSize", "文件大小" },
            { "IsPreview", "是否预览" },
            { "Version", "版本号" },
            { "Description", "描述" },
            { "CreatedAt", "创建时间" },
            { "UpdatedAt", "更新时间" },
            { "CategoryType", "分类类型" },
            { "CreatedBy", "创建者" },
            { "IsActive", "是否激活" },
            { "Title", "标题" },
            { "Keywords", "关键字" },
            { "IsPublic", "是否公开" },
            { "UpdatedBy", "更新者" },
             //FileAttribute 属性映射
            { "FileStorageId", "存储文件ID" },
            { "Length", "长度" },
            { "Width", "宽度" },
            { "Height", "高度" },
            { "Angle", "角度" },
            { "BasePointX", "基点X" },
            { "BasePointY", "基点Y" },
            { "BasePointZ", "基点Z" },
            { "MediumName", "介质" },
            { "Specifications", "规格" },
            { "Material", "材质" },
            { "StandardNumber", "标准号" },
            { "Power", "功率" },
            { "Volume", "容积" },
            { "Pressure", "压力" },
            { "Temperature", "温度" },
            { "Diameter", "直径" },
            { "OuterDiameter", "外径" },
            { "InnerDiameter", "内径" },
            { "Thickness", "厚度" },
            { "Weight", "重量" },
            { "Model", "型号" },
            { "Remarks", "备注" },
            { "Customize1", "自定义1" },
            { "Customize2", "自定义2" },
            { "Customize3", "自定义3" },

            // 新增或替换的属性映射
            { "AttributeGroup", "属性分组" },
            { "PressureRating", "压力等级" },
            { "OperatingPressure", "操作压力" },
            { "OperatingTemperature", "操作温度" },
            { "NominalDiameter", "公称直径DN" },
            { "Density", "密度" },
            { "Flow", "流量" },
            { "Velocity", "流速" },
            { "Lift", "扬程" },
            { "Voltage", "电压" },
            { "Current", "电流" },
            { "Frequency", "频率" },
            { "Conductivity", "电导率" },
            { "Moisture", "含湿量" },
            { "Humidity", "湿度" },
            { "Vacuum", "真空度" },
            { "Radiation", "辐射量" },

            { "PipeSpec", "管道规格" },
            { "PipeNominalDiameter", "管道公称直径" },
            { "PipeWallThickness", "管道壁厚" },
            { "PipePressureClass", "管道压力等级" },
            { "ConnectionType", "连接方式" },
            { "PipeSlope", "管道坡度" },
            { "AnticorrosionTreatment", "防腐处理" },

            { "ValveModel", "阀门型号" },
            { "ValveBodyMaterial", "阀体材质" },
            { "ValveDiscMaterial", "阀板材质" },
            { "ValveBallMaterial", "球体材质" },
            { "SealMaterial", "密封材质" },
            { "DriveType", "传动方式" },
            { "OpenMode", "开启方式" },
            { "ApplicableMedium", "适用介质" },

            { "FlangeModel", "法兰型号" },
            { "FlangeType", "法兰类型" },
            { "FlangeFaceType", "密封面形式" },
            { "FlangeStandard", "法兰标准" },
            { "BoltSpec", "螺栓规格" },

            { "PumpModel", "泵型号" },
            { "PumpFlow", "泵流量" },
            { "PumpHead", "泵扬程" },
            { "PumpBodyMaterial", "泵体材质" },
            { "MotorPower", "电机功率" },
            { "InletOutletDiameter", "进出口直径" },
            { "RatedSpeed", "额定转速" },
            { "ProtectionLevel", "防护等级" },

            { "FlueGasCapacity", "处理烟气量" },
            { "DesulfurizationEfficiency", "脱硫效率" },
            { "DropletSize", "液滴粒径" },
            { "SprayLayerCount", "喷淋层数" },
            { "ChimneySpec", "烟囱规格" },
            { "ChimneyDiameter", "烟囱直径" },
            { "ChimneyHeight", "烟囱高度" },
            { "ChimneyMaterial", "烟囱材质" },
            { "ChimneyThickness", "烟囱壁厚" },
            { "OutletWindSpeed", "出口风速" },
            { "InsulationThickness", "保温层厚度" },
            { "SupportType", "支撑方式" },
            { "FlueGasTemperature", "烟气温度" },

            { "PressureGaugeModel", "压力表型号" },
            { "ThermometerModel", "温度计型号" },
            { "FilterModel", "过滤器型号" },
            { "CheckValveModel", "止回阀型号" },
            { "SprinklerModel", "喷淋头型号" },
            { "FlowMeterModel", "流量计型号" },
            { "SafetyValveModel", "安全阀型号" },
            { "FlexibleJointModel", "柔性接头型号" },
        };

        // 常用字段的优先级映射（数值越小越靠前显示）
        private static readonly Dictionary<string, int> _priorityOverrides = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            // 常用设备表列优先级定义（越小越靠前）
            { "名称", 10 }, // 名称第一
            { "设备名称", 10 }, // 设备名称同等优先
            { "规格", 20 }, // 规格靠前
            { "规格型号", 20 }, // 规格型号同等
            { "材料", 30 }, // 材料靠前
            { "材质", 30 }, // 材质同义词
            { "数量", 40 }, // 数量靠前
            { "数", 40 }, // 简写保护
            { "图号或标准号", 50 }, // 图号/标准号靠前
            { "图号", 50 }, // 图号同义
            { "标准号", 50 }, // 标准号同义
            { "DWG.No./STD.No.", 50 }, // 英文样式
            { "管段号", 60 }, // 管段号
            { "管道标题", 60 }, // 管道标题
            { "介质", 70 }, // 介质
            { "介质名称", 70 }, // 介质名称
            { "起点", 80 }, // 起点/终点靠后
            { "终点", 80 }, // 终点
            { "长度", 90 }, // 长度类字段靠后
            { "Length", 90 } // 英文长度
        }; // end _priorityOverrides

        // 常用字段的显示别名映射（对输入键做友好显示）
        private static readonly Dictionary<string, string> _displayAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "名称", "名称" }, // 原样显示
            { "设备名称", "名称" }, // 设备名称显示为“名称”
            { "规格", "规格" }, // 规格
            { "规格型号", "规格" }, // 规格型号显示为“规格”
            { "材料", "材料" }, // 材料
            { "材质", "材料" }, // 材质别名也显示为“材料”
            { "数量", "数量" }, // 数量
            { "图号或标准号", "图号或标准号" }, // 图号或标准号
            { "图号", "图号或标准号" }, // 图号映射到“图号或标准号”
            { "标准号", "图号或标准号" }, // 标准号映射到“图号或标准号”
            { "DWG.No./STD.No.", "图号或标准号" }, // 英文映射
            { "管段号", "管段号" }, // 管段号
            { "管道标题", "管道标题" }, // 管道标题
            { "介质", "介质" }, // 介质
            { "介质名称", "介质" }, // 介质名称
            { "起点", "起点" }, // 起点
            { "终点", "终点" }, // 终点
            { "长度", "长度" }, // 长度
            { "Length", "长度" } // 英文长度显示为“长度”
        }; // end _displayAliases

        /// <summary>
        /// 获取管道属性的排序优先级（数值越小越靠前）
        /// </summary>
        /// <param name="key">属性键名（可能为中文或英文）</param>
        /// <returns>返回一个整数优先级，默认较大的值表示靠后</returns>
        public static int GetPipeAttributeSortPriority(string key)
        {
            // 防御性检查：空键给出最大优先级（靠后显示）
            if (string.IsNullOrWhiteSpace(key)) return int.MaxValue / 2;

            // 标准化键：去前后空白并小写以便匹配
            string k = key.Trim(); // 保留原字符串但去空白

            // 1）精确匹配优先级字典
            if (_priorityOverrides.TryGetValue(k, out int pExact))
            {
                return pExact; // 命中直接返回预定义的优先级
            }

            // 2）宽松包含匹配（支持“材质/材料”、“图号/标准”等多种写法）
            string lower = k.ToLowerInvariant(); // 小写用于包含判断
            if (lower.Contains("名称") || lower.Contains("name") || lower.Contains("tag")) return 10; // 名称类靠前
            if (lower.Contains("规格") || lower.Contains("型号") || lower.Contains("spec")) return 20; // 规格类
            if (lower.Contains("材质") || lower.Contains("材料") || lower.Contains("material")) return 30; // 材料类
            if (lower.Contains("数") && !lower.Contains("标准")) return 40; // 数量类（排除“标准”包含“数”的特殊）
            if (lower.Contains("图号") || lower.Contains("标准") || lower.Contains("dwg") || lower.Contains("std")) return 50; // 图号/标准
            if (lower.Contains("管段") || lower.Contains("pipe") || lower.Contains("pipeline")) return 60; // 管段/管道号类
            if (lower.Contains("介质")) return 70; // 介质
            if (lower.Contains("起点") || lower.Contains("终点") || lower.Contains("from") || lower.Contains("to")) return 80; // 起终点
            if (lower.Contains("长") || lower.Contains("长度") || lower.Contains("length")) return 90; // 长度类

            // 3）兜底：未识别的字段给一个稳定且靠后的优先级（以长度与首字符保证稳定性）
            int basePriority = 500; // 基础靠后数值
            int stableOffset = (k.Length % 100); // 长度作为偏移，保证排序稳定但可分散
            return basePriority + stableOffset; // 返回稳定的默认优先级
        } // end GetPipeAttributeSortPriority

        /// <summary>
        /// 获取属性在界面中显示的别名（若有映射则返回映射值，否则返回原键或简化后的显示）
        /// </summary>
        /// <param name="key">原始属性键</param>
        /// <returns>返回用于界面显示的友好名称</returns>
        public static string GetPipeAttributeDisplayAlias(string key)
        {
            // 防御性检查：空键直接返回空字符串
            if (string.IsNullOrWhiteSpace(key)) return string.Empty;

            // 去除多余空白
            string k = key.Trim();

            // 1）优先精确映射
            if (_displayAliases.TryGetValue(k, out string aliasExact))
            {
                return aliasExact; // 命中直接返回别名
            }

            // 2）宽松包含匹配，兼容各种写法
            string lower = k.ToLowerInvariant();
            if (lower.Contains("名称") || lower.Contains("name") || lower.Contains("tag")) return "名称";
            if (lower.Contains("规格") || lower.Contains("型号") || lower.Contains("spec")) return "规格";
            if (lower.Contains("材质") || lower.Contains("材料") || lower.Contains("material")) return "材料";
            if (lower.Contains("数量") || lower == "qty" || lower == "qun" || lower.Contains("数")) return "数量";
            if (lower.Contains("图号") || lower.Contains("标准") || lower.Contains("dwg") || lower.Contains("std")) return "图号或标准号";
            if (lower.Contains("管段") || lower.Contains("pipe") || lower.Contains("pipeline")) return "管段号";
            if (lower.Contains("介质")) return "介质";
            if (lower.Contains("起点") || lower.Contains("from")) return "起点";
            if (lower.Contains("终点") || lower.Contains("to")) return "终点";
            if (lower.Contains("长度") || lower.Contains("length")) return "长度";

            // 3）兜底：如键含英文可尝试做大小写友好化，否则原样返回（尽量保持原键以免丢信息）
            //    如果键包含下划线或点，替换为空格以便显示更友好
            string normalized = k.Replace('_', ' ').Replace('.', ' ').Trim();
            return normalized; // 返回处理后的原始键作为显示名称
        } // end GetPipeAttributeDisplayAlias
    }
}
