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
                { "管子", "Pipe(m)" },
                { "阀门", "Valve(Pcs.)" },
                { "法兰", "Flange(Pcs.)" },
                { "垫片", "Gasket(Pcs.)" },
                { "螺栓", "Bolt(Pcs.)" },
                { "螺母", "Nut(Pcs.)" },
                { "名称", "Name" },
                { "介质名称", "Medium Name" },
                { "规格", "Specs." },
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
                { "操作压力", "MPaG" },
                { "备注", "Remark" }
            };

        /// <summary>
        /// 属性同义词映射与过滤规则
        /// - AttributeSynonyms: 将常见属性名映射为标准列名，便于把诸如 "阀体材料" 写入 "材料" 列
        /// - ExcludedAttributeSubstrings: 出现在属性名中的子串如果匹配到则排除该属性（不生成列）
        /// </summary>
        public static readonly Dictionary<string, string> AttributeSynonyms = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // 材料类
            { "阀体材料", "材质" },
            { "阀体材质", "材质" },
            { "材料", "材质" },

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
    }
}
