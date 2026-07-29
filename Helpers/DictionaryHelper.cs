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


        /// <summary>
        /// 核心映射：canonical(英文) => 别名集合（包括中文、旧英文）
        /// </summary>
        public static readonly Dictionary<string, List<string>> _canonicalToAliases = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase)
        {
            // ======================== 管道/阀门/法兰等（原有） ========================
            {"PIPELINETITLE", new List<string>{ "PIPELINETITLE","管道标题","管道提示标题","PIPELINE_TITLE","PIPE_TITLE" }},
            {"TAG_NO", new List<string>{ "TAG_NO","管段号","管段编号","Pipeline No","Pipe No","管段","位号","位号（例：P-1001）" }},
            {"NAME", new List<string>{ "NAME","名称","设备名称","位号","Tag" }},
            {"MODEL", new List<string>{ "MODEL","规格型号","规格","型号","Spec" }},
            {"DRAWINGNO.STANDARDNO", new List<string>{ "DRAWINGNO.STANDARDNO","DRAWINGNO_STANDARDNO","图号","设计标准","标准号" }},
            {"DN", new List<string>{ "DN","公称通径","管径" }},
            {"PIPE_OD", new List<string>{ "PIPE_OD","管道外径","外径" }},
            {"PIPE_ID", new List<string>{ "PIPE_ID","管道内径","内径" }},
            {"PIPE_THK", new List<string>{ "PIPE_THK","管道壁厚","壁厚" }},
            {"SCHEDULE", new List<string>{ "SCHEDULE","壁厚等级","SCH" }},
            {"PN", new List<string>{ "PN","公称压力" }},
            {"CLASS", new List<string>{ "CLASS","压力等级" }},
            {"WORK_PRESSURE", new List<string>{ "WORK_PRESSURE","工作压力" }},
            {"DESIGN_PRESSURE", new List<string>{ "DESIGN_PRESSURE","设计压力" }},
            {"DESIGN_TEMP", new List<string>{ "DESIGN_TEMP","设计温度" }},
            {"WORK_TEMP", new List<string>{ "WORK_TEMP","工作温度" }},
            {"TEMP_RANGE", new List<string>{ "TEMP_RANGE","适用温度范围" }},
            {"MEDIUM", new List<string>{ "MEDIUM","介质","适用介质" }},
            {"PIPE_TYPE", new List<string>{ "PIPE_TYPE","管道类型" }},
            {"PIPE_CLASS", new List<string>{ "PIPE_CLASS","管道等级" }},
            {"CONN_TYPE", new List<string>{ "CONN_TYPE","连接方式" }},
            {"PIPE_MATL", new List<string>{ "PIPE_MATL","管道材质","材质","材料" }},
            {"LINING_MATL", new List<string>{ "LINING_MATL","衬里材质","衬里" }},
            {"LINING_THK", new List<string>{ "LINING_THK","衬里厚度" }},
            {"PIPE_LENGTH", new List<string>{ "PIPE_LENGTH","管道长度","长度" }},
            // ★ 合并后的 FLOW_RATE（包含管道和泵的别名）
            {"FLOW_RATE", new List<string>{ "FLOW_RATE","介质流量","流量","流量（选型,m³/h）","选型流量" }},
            {"HOT\\SOUND_ISOLACODE", new List<string>{ "HOT\\SOUND_ISOLACODE", "隔热隔声代号:", "隔热隔声代号" }},
            {"IS_ANTICORRO", new List<string>{ "IS_ANTICORRO", "是否防腐:", "是否防腐" }},
            {"START_POINT", new List<string>{ "START_POINT", "起始点","起始点" }},
            {"END_POINT", new List<string>{ "END_POINT", "终点","终点" }},
            {"QTY", new List<string>{ "QTY","数量","数量(个)","配置台数","泵台数","单套净化系统设备数量" }},
            {"WEIGHT", new List<string>{ "WEIGHT","总重量","重量","设备总重量（kg）","设备总重量" }},
            {"REMARK", new List<string>{ "REMARK","备注" }},
            {"SW_MODEL", new List<string>{ "SW_MODEL","SW_MODEL","3D模型","对应3D模型文件名" }},
            {"SYSTEM", new List<string>{ "SYSTEM","系统","所属系统","所属系统（吸收系统,浆液制备,石膏脱水,废水系统）" }},
        
            // ======================== 脱硫系统（塔）工艺计算参数 ========================
            {"SO2_IN", new List<string>{ "SO2_IN","入口烟气SO2浓度","入口SO2","SO2入口浓度" }},
            {"SO2_OUT", new List<string>{ "SO2_OUT","出口烟气SO2浓度","出口SO2","SO2出口浓度" }},
            {"ATM_PRESSURE", new List<string>{ "ATM_PRESSURE","当地大气压","大气压","ATM_PRESSURE" }},
            {"EXHAUST_TEMP", new List<string>{ "EXHAUST_TEMP","脱硫系统出口温度","出口温度" }},
            {"GAS_FLOW_DRY", new List<string>{ "GAS_FLOW_DRY","单套净化系统出口标况烟气量","标况烟气量" }},
            {"GAS_FLOW_WET", new List<string>{ "GAS_FLOW_WET","单套净化系统出口工况烟气量","工况烟气量" }},
            {"STACK_EXIT_VEL", new List<string>{ "STACK_EXIT_VEL","烟囱出口风速","出口风速" }},
            {"STACK_EXIT_DIA", new List<string>{ "STACK_EXIT_DIA","烟囱出口直径","烟囱直径","烟囱出口直径取值" }},
            {"GAS_FLOW", new List<string>{ "GAS_FLOW","单台脱硫装置标况烟气量","单台烟气量" }},
            {"SO2_MOL", new List<string>{ "SO2_MOL","单套净化系统SO2摩尔数","SO2摩尔数" }},
            {"CA_S_RATIO", new List<string>{ "CA_S_RATIO","钙硫比","CalS","Ca/S" }},
            {"L_G_RATIO", new List<string>{ "L_G_RATIO","液气比" }},
            {"SPRAY_LAYER", new List<string>{ "SPRAY_LAYER","喷淋层数" }},
            {"GYPSUM_OUTPUT", new List<string>{ "GYPSUM_OUTPUT","单套系统石膏产量","石膏产量" }},
            {"GYPSUM_WATER", new List<string>{ "GYPSUM_WATER","单套系统石膏携带水量","石膏携带水量" }},
            {"GYPSUM_BULK_DENSITY", new List<string>{ "GYPSUM_BULK_DENSITY","石膏堆积密度","石膏堆积密度（kg/m³）" }},
            {"GYPSUM_SLURRY_DENSITY", new List<string>{ "GYPSUM_SLURRY_DENSITY","石膏浆液平均密度","浆液平均密度","石膏浆液密度（kg/m³）" }},
            {"GYPSUM_SLURRY_OUTPUT", new List<string>{ "GYPSUM_SLURRY_OUTPUT","单套系统石膏浆液产量","石膏浆液产量" }},
            {"ACCIDENT_SLURRY_DENSITY", new List<string>{ "ACCIDENT_SLURRY_DENSITY","事故浆液密度","事故浆液密度（kg/m³）" }},
            {"FILTRATE_DENSITY", new List<string>{ "FILTRATE_DENSITY","滤液水密度","滤液水密度（kg/m³）" }},
        
            // ======================== 泵同步核心字段（L1） ========================
            {"PUMP_NAME", new List<string>{ "PUMP_NAME","泵名称","泵名称（循环泵,石灰石浆液泵,石膏排出泵,缓冲泵,滤液水泵,事故泵,地坑泵）","名称" }},
            {"PUMP_MODEL", new List<string>{ "PUMP_MODEL","泵型号","泵型号（例：800DT-A90, 150DT-A40）","型号" }},
            {"PUMP_EFF", new List<string>{ "PUMP_EFF", "泵效率（%）", "泵效率","效率"}},
            {"PUMP_TYPE", new List<string>{ "PUMP_TYPE","泵型式","泵型式（卧式离心渣浆泵,衬胶泵,氟塑料泵,液下泵）" }},
            {"PUMP_STRUCT", new List<string>{ "PUMP_STRUCT", "结构形式（悬臂式OH1,两端支承式BB,立式悬吊式VS）", "结构形式" , "泵结构形式", "泵结构" }},
            {"FLOW_VEL", new List<string>{ "FLOW_VEL","设计流速","流速" }},
            {"FLOW_RATE_CAL", new List<string>{ "FLOW_RATE_CAL", "计算单台介质泵流量（m³/h）", "计算单台介质泵流量","单台介质泵流量" ,"介质泵流量"}},
            // 注：FLOW_RATE 已在上方合并，此处不再重复
            {"PUMP_HEAD", new List<string>{ "PUMP_HEAD","扬程（m）","扬程" }},
            {"WET_MATL", new List<string>{ "WET_MATL","过流部件材质","过流部件材质（高铬合金Cr26,双相钢2205,A49,衬胶,氟塑料,UHMWPE）" }},
            {"SHAFT_SEAL", new List<string>{ "SHAFT_SEAL","轴封形式","轴封形式（集装式机械密封,双端面机械密封,填料密封,副叶轮动力密封）" }},
            {"QTY_CONFIG", new List<string>{ "QTY_CONFIG","配置台数（台）","配置台数" }},
        
            // ======================== 泵导出扩展字段（L2） ========================
            {"PUMP_SPEED", new List<string>{ "PUMP_SPEED","泵转速","泵转速（r/min:300～2900）" }},
            {"IMPELLER_DIA", new List<string>{ "IMPELLER_DIA","叶轮直径","叶轮直径（mm）" }},
            {"NPSHR", new List<string>{ "NPSHR","汽蚀余量","汽蚀余量（m）" }},
        
            // ======================== 物料与介质参数 ========================
            {"MEDIUM_DENSITY", new List<string>{ "MEDIUM_DENSITY","介质密度","浆液密度（kg/m³）","浆液密度","介质密度","密度比重" }},
            {"MEDIUM_CONTENT", new List<string>{ "MEDIUM_CONTENT", "含固量","浆液含固量","介质含固量（%）","介质含固量" }},
            {"MEDIUM_PH_VALUE", new List<string>{ "MEDIUM_PH_VALUE", "介质pH值", "pH值" }},
            {"MEDIUM_TEMP", new List<string>{ "MEDIUM_TEMP","介质温度","介质温度（℃:-20～120）" }},
            {"MEDIUM_TOTAL", new List<string>{ "MEDIUM_TOTAL", "介质循环总量（m³/h）", "循环总量" ,"介质循环总量"}},
            {"PH_VALUE", new List<string>{ "PH_VALUE","pH值","浆液pH值（2.5～13）" }},
            {"CHLORIDE_CONTENT", new List<string>{ "CHLORIDE_CONTENT","氯离子含量","氯离子含量（ppm:≤60000）" }},
        
            // ======================== 电机与效率参数 ========================
            {"POWER_CAL", new List<string>{ "POWER_CAL","计算功率（kW）","计算功率" }},
            {"MOTOR_POWER_REC", new List<string>{ "MOTOR_POWER_REC","推荐电机功率（kW）","推荐电机功率","推荐功率" }},
            {"MOTOR_POWER", new List<string>{ "MOTOR_POWER","电机功率（kW）","电机功率","电机功率选型","配套电机功率" }},
            {"MOTOR_COEFF", new List<string>{ "MOTOR_COEFF","电机系数","电机系数（安全裕量:1.1～1.2）" }},
            {"POWER_VOLTAGE", new List<string>{ "POWER_VOLTAGE","电源电压"}},
            {"MOTOR_IP", new List<string>{ "MOTOR_IP","电机防护等级","电机防护等级（IP54,IP55,IP65）" }},
            {"MOTOR_INSUL", new List<string>{ "MOTOR_INSUL","电机绝缘等级","电机绝缘等级（F级,H级）" }},
        
            // ======================== 循环系统工艺计算参数（循环泵专用） ========================
            {"SLURRY_TOTAL", new List<string>{ "SLURRY_TOTAL","浆液循环总量（m³/h）","浆液循环总量" }},
            {"RESIDENCE_TIME", new List<string>{ "RESIDENCE_TIME","停留时间","停留时间（min:石灰石溶解>4.3,石膏结晶≥15h）" }},
            {"OXIDATION_VOL", new List<string>{ "OXIDATION_VOL","氧化区有效容积","循环氧化区有效容积（m³）" }},
            {"OXIDATION_DIA", new List<string>{ "OXIDATION_DIA","氧化区直径","循环氧化区直径（m）" }},
            {"OXIDATION_HGT", new List<string>{ "OXIDATION_HGT","氧化区有效高度","循环氧化区有效高度（m）" }},
            {"PUMP_QTY_RUN", new List<string>{ "PUMP_QTY_RUN","运行泵台数","配置循环泵台数（连续运行）" }},
            {"PUMP_QTY_STANDBY", new List<string>{ "PUMP_QTY_STANDBY","备用泵台数","备用（台）" }},
            {"FLOW_SINGLE", new List<string>{ "FLOW_SINGLE","单台循环泵流量","单台循环泵流量（m³/h）","单台流量" }},
        
            // ======================== 运行与配置参数 ========================
            {"OPERATION_HOURS", new List<string>{ "OPERATION_HOURS","运行时长","运行时长（h/年）" }},
            {"REMOTE_SYS_QTY", new List<string>{ "REMOTE_SYS_QTY","远处系统数量","远处系统数量（套）","系统数量(远)","远处系统" }},
            {"NEAR_SYS_QTY", new List<string>{ "NEAR_SYS_QTY","近处系统数量","近处系统数量（套）","系统数量(近)","近处系统" }},
            {"SUPPLY_SYS_QTY", new List<string>{ "SUPPLY_SYS_QTY","供给系统数量","供给系统数量（套）","系统数量" }},
        };

        /// <summary>
        /// 计算表格中文列名 → 图元属性英文标记 映射字典
        /// 覆盖：脱硫系统（塔）、泵、管道 三大类
        /// </summary>
        public static readonly Dictionary<string, string> CalcColumnToAttributeMap = new Dictionary<string, string>
        {
            // ======================== 脱硫系统（塔）工艺计算参数 ========================
            ["入口烟气SO2浓度"] = "SO2_IN",
            ["出口烟气SO2浓度"] = "SO2_OUT",
            ["当地大气压"] = "ATM_PRESSURE",
            ["脱硫系统出口温度"] = "EXHAUST_TEMP",
            ["单套净化系统出口标况烟气量"] = "GAS_FLOW_DRY",
            ["单套净化系统出口工况烟气量"] = "GAS_FLOW_WET",
            ["烟囱出口风速"] = "STACK_EXIT_VEL",
            ["烟囱出口直径"] = "STACK_EXIT_DIA",
            ["烟囱出口直径取值"] = "STACK_EXIT_DIA",
            ["单套净化系统设备数量"] = "QTY",
            ["单台脱硫装置标况烟气量"] = "GAS_FLOW",
            ["单套净化系统SO2摩尔数"] = "SO2_MOL",
            ["CalS"] = "CA_S_RATIO",
            ["Ca/S"] = "CA_S_RATIO",
            ["液气比"] = "L_G_RATIO",
            ["喷淋层数"] = "SPRAY_LAYER",
            ["单套系统石膏产量"] = "GYPSUM_OUTPUT",
            ["单套系统石膏携带水量"] = "GYPSUM_WATER",
            ["石膏堆积密度"] = "GYPSUM_BULK_DENSITY",
            ["石膏浆液平均密度"] = "GYPSUM_SLURRY_DENSITY",
            ["单套系统石膏浆液产量"] = "GYPSUM_SLURRY_OUTPUT",
            ["事故浆液密度"] = "ACCIDENT_SLURRY_DENSITY",
            ["滤液水密度"] = "FILTRATE_DENSITY",

            // ======================== 泵同步核心字段（L1） ========================
            ["位号"] = "TAG_NO",
            ["位号（例：P-1001）"] = "TAG_NO",
            ["TAG_NO"] = "TAG_NO",
            ["泵名称"] = "PUMP_NAME",
            ["泵名称（循环泵,石灰石浆液泵,石膏排出泵,缓冲泵,滤液水泵,事故泵,地坑泵）"] = "PUMP_NAME",
            ["泵型号"] = "PUMP_MODEL",
            ["泵型号（例：800DT-A90, 150DT-A40）"] = "PUMP_MODEL",
            ["泵型式"] = "PUMP_TYPE",
            ["泵型式（卧式离心渣浆泵,衬胶泵,氟塑料泵,液下泵）"] = "PUMP_TYPE",
            ["计算单台浆液泵流量"] = "FLOW_RATE_CAL",
            ["计算单台浆液泵流量（m³/h）"] = "FLOW_RATE_CAL",
            ["计算流量"] = "FLOW_RATE_CAL",
            ["流量（选型）"] = "FLOW_RATE",
            ["流量（选型,m³/h）"] = "FLOW_RATE",
            ["选型流量"] = "FLOW_RATE",
            ["扬程"] = "PUMP_HEAD",
            ["扬程（m）"] = "PUMP_HEAD",
            ["过流部件材质"] = "WET_MATL",
            ["过流部件材质（高铬合金Cr26,双相钢2205,A49,衬胶,氟塑料,UHMWPE）"] = "WET_MATL",
            ["轴封形式"] = "SHAFT_SEAL",
            ["轴封形式（集装式机械密封,双端面机械密封,填料密封,副叶轮动力密封）"] = "SHAFT_SEAL",
            ["配置台数"] = "QTY_CONFIG",
            ["配置台数（台）"] = "QTY_CONFIG",
            ["对应3D模型文件名"] = "SW_MODEL",
            ["SW_MODEL"] = "SW_MODEL",

            // ======================== 泵导出扩展字段（L2） ========================
            ["图号"] = "DRAWINGNO",
            ["设计标准"] = "DESIGN_STD",
            ["图号、设计标准"] = "DRAWINGNO.STANDARDNO",
            ["泵转速"] = "PUMP_SPEED",
            ["泵转速（r/min:300～2900）"] = "PUMP_SPEED",
            ["叶轮直径"] = "IMPELLER_DIA",
            ["叶轮直径（mm）"] = "IMPELLER_DIA",
            ["汽蚀余量"] = "NPSHR",
            ["汽蚀余量（m）"] = "NPSHR",

            // ======================== 物料与介质参数 ========================
            ["输送介质"] = "MEDIUM",
            ["输送介质（石灰石浆液,石膏浆液,滤液水,事故浆液,工艺水）"] = "MEDIUM",
            ["介质"] = "MEDIUM",
            ["介质密度"] = "MEDIUM_DENSITY",
            ["浆液密度"] = "MEDIUM_DENSITY",
            ["浆液密度（kg/m³）"] = "MEDIUM_DENSITY",
            ["含固量"] = "SOLID_CONTENT",
            ["浆液含固量（%:灰浆≤45,矿浆≤60）"] = "SOLID_CONTENT",
            ["介质温度"] = "MEDIUM_TEMP",
            ["介质温度（℃:-20～120）"] = "MEDIUM_TEMP",
            ["pH值"] = "PH_VALUE",
            ["浆液pH值（2.5～13）"] = "PH_VALUE",
            ["氯离子含量"] = "CHLORIDE_CONTENT",
            ["氯离子含量（ppm:≤60000）"] = "CHLORIDE_CONTENT",

            // ======================== 电机与效率参数 ========================
            ["泵效率"] = "PUMP_EFF",
            ["泵效率（%:大型泵85～88,小型泵60～80）"] = "PUMP_EFF",
            ["效率"] = "PUMP_EFF",
            ["计算功率"] = "POWER_CAL",
            ["计算功率（kW）"] = "POWER_CAL",
            ["推荐电机功率"] = "MOTOR_POWER_REC",
            ["推荐电机功率（kW）"] = "MOTOR_POWER_REC",
            ["推荐功率"] = "MOTOR_POWER_REC",
            ["电机功率"] = "MOTOR_POWER",
            ["电机功率（kW）"] = "MOTOR_POWER",
            ["电机功率选型"] = "MOTOR_POWER",
            ["电机系数"] = "MOTOR_COEFF",
            ["电机系数（安全裕量:1.1～1.2）"] = "MOTOR_COEFF",
            ["电源电压"] = "POWER_VOLTAGE",
            ["电源电压（380V/50Hz,6kV/50Hz,10kV/50Hz）"] = "POWER_VOLTAGE",
            ["电机防护等级"] = "MOTOR_IP",
            ["电机防护等级（IP54,IP55,IP65）"] = "MOTOR_IP",
            ["电机绝缘等级"] = "MOTOR_INSUL",
            ["电机绝缘等级（F级,H级）"] = "MOTOR_INSUL",

            // ======================== 循环系统工艺计算参数（循环泵专用） ========================
            ["浆液循环总量"] = "SLURRY_TOTAL",
            ["浆液循环总量（m³/h）"] = "SLURRY_TOTAL",
            ["停留时间"] = "RESIDENCE_TIME",
            ["停留时间（min:石灰石溶解>4.3,石膏结晶≥15h）"] = "RESIDENCE_TIME",
            ["氧化区有效容积"] = "OXIDATION_VOL",
            ["循环氧化区有效容积（m³）"] = "OXIDATION_VOL",
            ["氧化区直径"] = "OXIDATION_DIA",
            ["循环氧化区直径（m）"] = "OXIDATION_DIA",
            ["氧化区有效高度"] = "OXIDATION_HGT",
            ["循环氧化区有效高度（m）"] = "OXIDATION_HGT",
            ["运行泵台数"] = "PUMP_QTY_RUN",
            ["配置循环泵台数（连续运行）"] = "PUMP_QTY_RUN",
            ["备用泵台数"] = "PUMP_QTY_STANDBY",
            ["备用（台）"] = "PUMP_QTY_STANDBY",
            ["单台循环泵流量"] = "FLOW_SINGLE",
            ["单台循环泵流量（m³/h）"] = "FLOW_SINGLE",
            ["单台流量"] = "FLOW_SINGLE",

            // ======================== 运行与配置参数 ========================
            ["运行时长"] = "OPERATION_HOURS",
            ["运行时长（h/年）"] = "OPERATION_HOURS",
            ["远处系统数量"] = "REMOTE_SYS_QTY",
            ["远处系统数量（套）"] = "REMOTE_SYS_QTY",
            ["系统数量(远)"] = "REMOTE_SYS_QTY",
            ["近处系统数量"] = "NEAR_SYS_QTY",
            ["近处系统数量（套）"] = "NEAR_SYS_QTY",
            ["系统数量(近)"] = "NEAR_SYS_QTY",
            ["供给系统数量"] = "SUPPLY_SYS_QTY",
            ["供给系统数量（套）"] = "SUPPLY_SYS_QTY",
            ["系统数量"] = "SUPPLY_SYS_QTY",
            ["设备总重量"] = "WEIGHT",
            ["设备总重量（kg）"] = "WEIGHT",
            ["总重量"] = "WEIGHT",
            ["重量"] = "WEIGHT",
            ["所属系统"] = "SYSTEM",
            ["所属系统（吸收系统,浆液制备,石膏脱水,废水系统）"] = "SYSTEM",
            ["备注"] = "REMARK",

            // ======================== 管道属性（原有，保持不变） ========================
            //["介质"] = "MEDIUM",
            ["使用条件"] = "",
            ["流速"] = "FLOW_VEL",
            ["管道尺寸"] = "DN",
            ["管道尺寸（计算）"] = "DN",
            ["取值 DN"] = "DN",
            ["核算流速"] = "",
            ["管道规格"] = "MODEL",
            ["材质"] = "PIPE_MATL",
            ["管道材质"] = "PIPE_MATL",
            ["管道外径"] = "PIPE_OD",
            ["管道内径"] = "PIPE_ID",
            ["管道壁厚"] = "PIPE_THK",
            ["壁厚等级"] = "SCHEDULE",
            ["公称通径"] = "DN",
            ["公称压力"] = "PN",
            ["工作压力"] = "WORK_PRESSURE",
            ["设计压力"] = "DESIGN_PRESSURE",
            ["设计温度"] = "DESIGN_TEMP",
            ["工作温度"] = "WORK_TEMP",
            ["连接方式"] = "CONN_TYPE",
            ["管道类型"] = "PIPE_TYPE",
            ["管道等级"] = "PIPE_CLASS",
            ["衬里材质"] = "LINING_MATL",
            ["衬里厚度"] = "LINING_THK",
            ["衬里工艺"] = "LINING_PROC",
            ["管道长度"] = "PIPE_LENGTH",
            ["设计流速"] = "FLOW_VEL",
            ["介质流量"] = "FLOW_RATE",
            ["外部防腐涂层"] = "PIPE_COATING",
            ["涂层厚度"] = "COATING_THK",
            ["保温材料"] = "INSUL_MATL",
            ["保温层厚度"] = "INSUL_THK",
            ["保温方式"] = "INSUL_TYPE",
            ["伴热类型"] = "HEAT_TRACE",
            ["电伴热功率"] = "HEAT_TRACE_POWER",
            ["无损检测比例"] = "NDT_RATIO",
            ["无损检测方法"] = "NDT_TYPE",
            ["允许压力损失"] = "PRESSURE_LOSS",
            ["管道内壁粗糙度"] = "ROUGHNESS",
            ["清洁度要求"] = "CLEANING_REQ",
            //["所属系统"] = "SYSTEM",
            ["数量"] = "QTY",
            //["总重量"] = "WEIGHT",
            //["备注"] = "REMARK",
            // 注意：管道部分原有“石灰石纯度”等可能属于泵，已移除
        };

        /// <summary>
        /// 计算表格中文列名 → 图元属性英文标记 映射字典
        /// 覆盖：脱硫系统（塔）、泵、管道 三大类
        /// </summary>
        public static readonly Dictionary<string, string> _CalcColumnToAttributeMap = new Dictionary<string, string>
        {
            // ======================================================
            // 一、脱硫系统（塔）属性
            // ======================================================
            ["入口烟气SO2浓度"] = "SO2_IN",
            ["出口烟气SO2浓度"] = "SO2_OUT",
            ["当地大气压"] = "ATM_PRESSURE",
            ["脱硫系统出口温度"] = "EXHAUST_TEMP",
            ["单套净化系统出口标况烟气量"] = "GAS_FLOW_DRY",
            ["单套净化系统出口工况烟气量"] = "GAS_FLOW_WET",
            ["烟囱出口风速"] = "STACK_EXIT_VEL",
            ["烟囱出口直径"] = "STACK_EXIT_DIA",
            ["烟囱出口直径取值"] = "STACK_EXIT_DIA",
            ["单套净化系统设备数量"] = "QTY",
            ["单台脱硫装置标况烟气量"] = "GAS_FLOW",
            ["单套净化系统SO2摩尔数"] = "SO2_MOL",
            ["CalS"] = "CA_S_RATIO",
            ["Ca/S"] = "CA_S_RATIO",
            ["液气比"] = "L_G_RATIO",
            ["喷淋层数"] = "SPRAY_LAYER",
            ["单套系统石膏产量"] = "GYPSUM_OUTPUT",
            ["单套系统石膏携带水量"] = "GYPSUM_WATER",
            ["石膏堆积密度"] = "GYPSUM_BULK_DENSITY",
            ["石膏浆液平均密度"] = "GYPSUM_SLURRY_DENSITY",
            ["单套系统石膏浆液产量"] = "GYPSUM_SLURRY_OUTPUT",
            ["事故浆液密度"] = "ACCIDENT_SLURRY_DENSITY",
            ["滤液水密度"] = "FILTRATE_DENSITY",

            // ======================================================
            // 二、泵属性
            // ======================================================
            ["远处系统数量"] = "REMOTE_SYS_QTY",
            ["近处系统数量"] = "NEAR_SYS_QTY",
            ["供给系统数量"] = "SUPPLY_SYS_QTY",
            ["系统数量(远)"] = "REMOTE_SYS_QTY",
            ["系统数量(近)"] = "NEAR_SYS_QTY",
            ["系统数量"] = "SUPPLY_SYS_QTY",
            ["运行时长"] = "OPERATION_HOURS",
            ["配置台数"] = "QTY",
            ["泵台数"] = "QTY",
            ["计算单台浆液泵流量"] = "FLOW_RATE_CAL",
            ["计算流量"] = "FLOW_RATE_CAL",
            ["流量（选型）"] = "FLOW_RATE",
            ["选型流量"] = "FLOW_RATE",
            ["扬程"] = "PUMP_HEAD",
            ["电机系数"] = "MOTOR_COEFF",
            ["效率"] = "MOTOR_EFF",
            ["计算功率"] = "POWER_CAL",
            ["推荐电机功率"] = "MOTOR_POWER_REC",
            ["推荐功率"] = "MOTOR_POWER_REC",
            ["电机功率"] = "MOTOR_POWER",
            ["电机功率选型"] = "MOTOR_POWER",
            ["石灰石纯度"] = "",
            ["石灰石消耗量"] = "",
            ["浆液密度"] = "MEDIUM_DENSITY",
            ["介质密度"] = "MEDIUM_DENSITY",

            // ======================================================
            // 三、管道属性
            // ======================================================
            ["介质"] = "MEDIUM",
            ["使用条件"] = "",
            ["流速"] = "FLOW_VEL",
            ["管道尺寸"] = "DN",
            ["管道尺寸（计算）"] = "DN",
            ["取值 DN"] = "DN",
            ["核算流速"] = "",
            ["管道规格"] = "MODEL",
            ["材质"] = "PIPE_MATL",
            ["管道材质"] = "PIPE_MATL",
            ["管道外径"] = "PIPE_OD",
            ["管道内径"] = "PIPE_ID",
            ["管道壁厚"] = "PIPE_THK",
            ["壁厚等级"] = "SCHEDULE",
            ["公称通径"] = "DN",
            ["公称压力"] = "PN",
            ["工作压力"] = "WORK_PRESSURE",
            ["设计压力"] = "DESIGN_PRESSURE",
            ["设计温度"] = "DESIGN_TEMP",
            ["工作温度"] = "WORK_TEMP",
            ["连接方式"] = "CONN_TYPE",
            ["管道类型"] = "PIPE_TYPE",
            ["管道等级"] = "PIPE_CLASS",
            ["衬里材质"] = "LINING_MATL",
            ["衬里厚度"] = "LINING_THK",
            ["衬里工艺"] = "LINING_PROC",
            ["管道长度"] = "PIPE_LENGTH",
            ["设计流速"] = "FLOW_VEL",
            ["介质流量"] = "FLOW_RATE",
            ["外部防腐涂层"] = "PIPE_COATING",
            ["涂层厚度"] = "COATING_THK",
            ["保温材料"] = "INSUL_MATL",
            ["保温层厚度"] = "INSUL_THK",
            ["保温方式"] = "INSUL_TYPE",
            ["伴热类型"] = "HEAT_TRACE",
            ["电伴热功率"] = "HEAT_TRACE_POWER",
            ["无损检测比例"] = "NDT_RATIO",
            ["无损检测方法"] = "NDT_TYPE",
            ["允许压力损失"] = "PRESSURE_LOSS",
            ["管道内壁粗糙度"] = "ROUGHNESS",
            ["清洁度要求"] = "CLEANING_REQ",
            ["所属系统"] = "SYSTEM",
            ["数量"] = "QTY",
            ["总重量"] = "WEIGHT",
            ["备注"] = "REMARK",
        };
        /// <summary>
        /// 计算表格中文列名 → 图元属性英文标记 映射字典
        /// 覆盖：脱硫系统（塔）、泵、管道 三大类
        /// </summary>
        private static readonly Dictionary<string, string> CalcTableColumnToAttributeMap = new Dictionary<string, string>
        {
            // ======================================================
            // 一、脱硫系统（塔）属性
            // ======================================================
            ["入口烟气SO2浓度"] = "SO2_IN",
            ["出口烟气SO2浓度"] = "SO2_OUT",
            ["当地大气压"] = "ATM_PRESSURE",
            ["脱硫系统出口温度"] = "EXHAUST_TEMP",
            ["单套净化系统出口标况烟气量"] = "GAS_FLOW_DRY",
            ["单套净化系统出口工况烟气量"] = "GAS_FLOW_WET",
            ["烟囱出口风速"] = "STACK_EXIT_VEL",
            ["烟囱出口直径"] = "STACK_EXIT_DIA",
            ["烟囱出口直径取值"] = "STACK_EXIT_DIA",
            ["单套净化系统设备数量"] = "QTY",
            ["单台脱硫装置标况烟气量"] = "GAS_FLOW",
            ["单套净化系统SO2摩尔数"] = "SO2_MOL",
            ["CalS"] = "CA_S_RATIO",
            ["Ca/S"] = "CA_S_RATIO",
            ["液气比"] = "L_G_RATIO",
            ["喷淋层数"] = "SPRAY_LAYER",
            ["单套系统石膏产量"] = "GYPSUM_OUTPUT",
            ["单套系统石膏携带水量"] = "GYPSUM_WATER",
            ["石膏堆积密度"] = "GYPSUM_BULK_DENSITY",
            ["石膏浆液平均密度"] = "GYPSUM_SLURRY_DENSITY",
            ["单套系统石膏浆液产量"] = "GYPSUM_SLURRY_OUTPUT",
            ["事故浆液密度"] = "ACCIDENT_SLURRY_DENSITY",
            ["滤液水密度"] = "FILTRATE_DENSITY",

            // ======================================================
            // 二、泵属性
            // ======================================================
            ["远处系统数量"] = "REMOTE_SYS_QTY",
            ["近处系统数量"] = "NEAR_SYS_QTY",
            ["供给系统数量"] = "SUPPLY_SYS_QTY",
            ["系统数量(远)"] = "REMOTE_SYS_QTY",
            ["系统数量(近)"] = "NEAR_SYS_QTY",
            ["系统数量"] = "SUPPLY_SYS_QTY",
            ["运行时长"] = "OPERATION_HOURS",
            ["配置台数"] = "QTY",
            ["泵台数"] = "QTY",
            ["计算单台浆液泵流量"] = "FLOW_RATE_CAL",
            ["计算流量"] = "FLOW_RATE_CAL",
            ["流量（选型）"] = "FLOW_RATE",
            ["选型流量"] = "FLOW_RATE",
            ["扬程"] = "PUMP_HEAD",
            ["电机系数"] = "MOTOR_COEFF",
            ["效率"] = "MOTOR_EFF",
            ["计算功率"] = "POWER_CAL",
            ["推荐电机功率"] = "MOTOR_POWER_REC",
            ["推荐功率"] = "MOTOR_POWER_REC",
            ["电机功率"] = "MOTOR_POWER",
            ["电机功率选型"] = "MOTOR_POWER",
            ["石灰石纯度"] = "",
            ["石灰石消耗量"] = "",
            ["浆液密度"] = "MEDIUM_DENSITY",
            ["介质密度"] = "MEDIUM_DENSITY",

            // ======================================================
            // 三、管道属性
            // ======================================================
            ["介质"] = "MEDIUM",
            ["使用条件"] = "",
            ["流速"] = "FLOW_VEL",
            ["管道尺寸"] = "DN",
            ["管道尺寸（计算）"] = "DN",
            ["取值 DN"] = "DN",
            ["核算流速"] = "",
            ["管道规格"] = "MODEL",
            ["材质"] = "PIPE_MATL",
            ["管道材质"] = "PIPE_MATL",
            ["管道外径"] = "PIPE_OD",
            ["管道内径"] = "PIPE_ID",
            ["管道壁厚"] = "PIPE_THK",
            ["壁厚等级"] = "SCHEDULE",
            ["公称通径"] = "DN",
            ["公称压力"] = "PN",
            ["工作压力"] = "WORK_PRESSURE",
            ["设计压力"] = "DESIGN_PRESSURE",
            ["设计温度"] = "DESIGN_TEMP",
            ["工作温度"] = "WORK_TEMP",
            ["连接方式"] = "CONN_TYPE",
            ["管道类型"] = "PIPE_TYPE",
            ["管道等级"] = "PIPE_CLASS",
            ["衬里材质"] = "LINING_MATL",
            ["衬里厚度"] = "LINING_THK",
            ["衬里工艺"] = "LINING_PROC",
            ["管道长度"] = "PIPE_LENGTH",
            ["设计流速"] = "FLOW_VEL",
            ["介质流量"] = "FLOW_RATE",
            ["外部防腐涂层"] = "PIPE_COATING",
            ["涂层厚度"] = "COATING_THK",
            ["保温材料"] = "INSUL_MATL",
            ["保温层厚度"] = "INSUL_THK",
            ["保温方式"] = "INSUL_TYPE",
            ["伴热类型"] = "HEAT_TRACE",
            ["电伴热功率"] = "HEAT_TRACE_POWER",
            ["无损检测比例"] = "NDT_RATIO",
            ["无损检测方法"] = "NDT_TYPE",
            ["允许压力损失"] = "PRESSURE_LOSS",
            ["管道内壁粗糙度"] = "ROUGHNESS",
            ["清洁度要求"] = "CLEANING_REQ",
            ["所属系统"] = "SYSTEM",
            ["数量"] = "QTY",
            ["总重量"] = "WEIGHT",
            ["备注"] = "REMARK",
        };

    }
}
