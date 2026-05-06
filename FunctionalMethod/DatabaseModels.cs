using System;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    // 顶级模型集合，替代对 DatabaseManager 内嵌类型的直接引用
    public class FileStorage
    {
        public int Id { get; set; }
        public int CategoryId { get; set; }
        public string? CategoryType { get; set; }
        public string? FileAttributeId { get; set; }
        public string? FileName { get; set; }
        public string? FileStoredName { get; set; }
        public string? DisplayName { get; set; }
        public string? FileType { get; set; }
        public string? FileHash { get; set; }
        public string? BlockName { get; set; }
        public string? LayerName { get; set; }
        public int? ColorIndex { get; set; }
        public double? Scale { get; set; }
        public string? FilePath { get; set; }
        public string? PreviewImageName { get; set; }
        public string? PreviewImagePath { get; set; }
        public long? FileSize { get; set; }
        public int IsPreview { get; set; }
        public int Version { get; set; }
        public string? Description { get; set; }
        public int IsActive { get; set; }
        public string? CreatedBy { get; set; }
        public string? UpdatedBy { get; set; }
        public string? Title { get; set; }
        public string? Keywords { get; set; }
        public int IsPublic { get; set; }
        public DateTime? LastAccessedAt { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
        // 是否天正（兼容旧代码使用的字段）
        public int? IsTianZheng { get; set; }
    }

    public class SyncManifest
    {
        public string StorageRoot { get; set; } = string.Empty;
        public string SourceRoot { get; set; } = string.Empty;
        public string ServerClientVersion { get; set; } = string.Empty;
        public DateTime GeneratedAt { get; set; }
        public List<SyncManifestItem> Items { get; set; } = new List<SyncManifestItem>();
    }

    public class SyncManifestItem
    {
        public int FileId { get; set; }
        public string? FileName { get; set; }
        public string? FileStoredName { get; set; }
        public string? PreviewImageName { get; set; }
        public string? FilePath { get; set; }
        public string? PreviewImagePath { get; set; }
        public string? FileHash { get; set; }
        public string? PreviewImageHash { get; set; }
        public int Version { get; set; }
        public DateTime UpdatedAt { get; set; }
        public string? LocalFileHash { get; set; }
        public string? LocalPreviewHash { get; set; }
        public string? LocalFilePath { get; set; }
        public string? LocalPreviewPath { get; set; }
        public bool ServerFileExists { get; set; }
        public bool ServerPreviewExists { get; set; }
        public bool LocalFileExists { get; set; }
        public bool LocalPreviewExists { get; set; }
        public bool FileHashDifferent { get; set; }
        public bool PreviewHashDifferent { get; set; }
        public string? DifferenceSummary { get; set; }
        public bool NeedsFileSync { get; set; }
        public bool NeedsPreviewSync { get; set; }
    }

    public class FileAttribute
    {
        // 使用 long 以兼容数据库中 BIGINT 类型
        public long Id { get; set; }
        public long? CategoryId { get; set; }
        public long FileStorageId { get; set; }
        public string? FileName { get; set; }
        public string? FileAttributeId { get; set; }
        public string? Description { get; set; }
        public string? AttributeGroup { get; set; }
        public string? Remarks { get; set; }
        public string? Customize1 { get; set; }
        public string? Customize2 { get; set; }
        public string? Customize3 { get; set; }

        // 使用 decimal? 与数据库小数类型更兼容，避免隐式转换错误
        public decimal? Length { get; set; }
        public decimal? Width { get; set; }
        public decimal? Height { get; set; }
        public decimal? Angle { get; set; }
        public decimal? BasePointX { get; set; }
        public decimal? BasePointY { get; set; }
        public decimal? BasePointZ { get; set; }

        public string? Model { get; set; }
        public string? Specifications { get; set; }
        public string? Material { get; set; }
        public string? MediumName { get; set; }
        public string? StandardNumber { get; set; }

        public decimal? Power { get; set; }
        public decimal? Volume { get; set; }
        public decimal? Flow { get; set; }
        public decimal? Velocity { get; set; }
        public decimal? Lift { get; set; }
        public decimal? Pressure { get; set; }
        public decimal? Temperature { get; set; }
        public string? PressureRating { get; set; }
        public decimal? OperatingPressure { get; set; }
        public decimal? OperatingTemperature { get; set; }

        public decimal? Diameter { get; set; }
        public decimal? OuterDiameter { get; set; }
        public decimal? InnerDiameter { get; set; }
        public string? NominalDiameter { get; set; }
        public decimal? Thickness { get; set; }
        public decimal? Weight { get; set; }
        public decimal? Density { get; set; }

        // 电气/机械特性
        public decimal? Voltage { get; set; }
        public decimal? Current { get; set; }
        public decimal? Frequency { get; set; }
        public decimal? Conductivity { get; set; }
        public decimal? Moisture { get; set; }
        public decimal? Humidity { get; set; }
        public decimal? Vacuum { get; set; }
        public decimal? Radiation { get; set; }

        // 管道/法兰/连接相关
        public string? PipeSpec { get; set; }
        public string? PipeNominalDiameter { get; set; }
        public decimal? PipeWallThickness { get; set; }
        public string? PipePressureClass { get; set; }
        public string? ConnectionType { get; set; }
        public string? PipeSlope { get; set; }
        public string? AnticorrosionTreatment { get; set; }

        // 阀门及材料
        public string? ValveModel { get; set; }
        public string? ValveBodyMaterial { get; set; }
        public string? ValveDiscMaterial { get; set; }
        public string? ValveBallMaterial { get; set; }
        public string? SealMaterial { get; set; }
        public string? DriveType { get; set; }
        public string? OpenMode { get; set; }
        public string? ApplicableMedium { get; set; }

        // 法兰/螺栓
        public string? FlangeModel { get; set; }
        public string? FlangeType { get; set; }
        public string? FlangeFaceType { get; set; }
        public string? FlangeStandard { get; set; }
        public string? BoltSpec { get; set; }

        // 减径器
        public string? ReducerSpec { get; set; }
        public string? ReducerLargeDn { get; set; }
        public string? ReducerSmallDn { get; set; }
        public decimal? ReducerWallThicknessLarge { get; set; }
        public decimal? ReducerWallThicknessSmall { get; set; }
        public string? ReducerConnectionType { get; set; }
        public string? ReducerConicity { get; set; }
        public string? ReducerEccentricDirection { get; set; }
        public string? ReducerApplicableMedium { get; set; }
        public string? ReducerAnticorrosion { get; set; }

        // 泵
        public string? PumpModel { get; set; }
        public decimal? PumpFlow { get; set; }
        public decimal? PumpHead { get; set; }
        public string? PumpBodyMaterial { get; set; }
        public decimal? MotorPower { get; set; }
        public string? InletOutletDiameter { get; set; }
        public decimal? RatedSpeed { get; set; }
        public string? PumpApplicableMedium { get; set; }
        public decimal? WorkingPressure { get; set; }
        public string? ProtectionLevel { get; set; }

        // 膨胀节
        public string? ExpansionJointModel { get; set; }
        public string? BellowsMaterial { get; set; }
        public string? FlangeOrNozzleMaterial { get; set; }
        public decimal? CompensationAmount { get; set; }
        public string? ExpansionJointConnectionType { get; set; }
        public string? ExpansionJointMedium { get; set; }
        public decimal? ExpansionJointWorkingTemp { get; set; }

        // 烟气/脱硫/风管
        public decimal? FlueGasCapacity { get; set; }
        public decimal? DesulfurizationEfficiency { get; set; }
        public decimal? DropletSize { get; set; }
        public int? SprayLayerCount { get; set; }
        public string? ChimneySpec { get; set; }
        public decimal? ChimneyDiameter { get; set; }
        public decimal? ChimneyHeight { get; set; }
        public string? ChimneyMaterial { get; set; }
        public decimal? ChimneyThickness { get; set; }
        public decimal? OutletWindSpeed { get; set; }
        public decimal? InsulationThickness { get; set; }
        public string? SupportType { get; set; }
        public decimal? FlueGasTemperature { get; set; }

        // 仪表/检测设备
        public string? PressureGaugeModel { get; set; }
        public string? ThermometerModel { get; set; }
        public string? FilterModel { get; set; }
        public string? CheckValveModel { get; set; }
        public string? SprinklerModel { get; set; }
        public string? FlowMeterModel { get; set; }
        public string? SafetyValveModel { get; set; }
        public string? FlexibleJointModel { get; set; }

        public DateTime? CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }
    }

    // 兼容之前 DatabaseManager 中使用的 SW 相关模型
    public class SwCategory
    {
        public int Id { get; set; }
        // 软件类别名，可空以兼容旧数据
        public string? Name { get; set; }
        // 显示名称，可空
        public string? DisplayName { get; set; }
        public int SortOrder { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
    }

    public class SwSubcategory
    {
        public int Id { get; set; }
        public int ParentId { get; set; }
        public int CategoryId { get; set; }
        // 子类别名，可空
        public string? Name { get; set; }
        // 子类别显示名，可空
        public string? DisplayName { get; set; }
        public int SortOrder { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
    }

    public class SwGraphic
    {
        public int Id { get; set; }
        public int SubcategoryId { get; set; }
        public string? FileName { get; set; }
        public string? DisplayName { get; set; }
        public string? FilePath { get; set; }
        public string? PreviewImagePath { get; set; }
        public long FileSize { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
    }

    public class CadCategory
    {
        public int Id { get; set; }
        // 分类名，允许为空以兼容反序列化/数据库可空列
        public string? Name { get; set; }
        // 显示名称，允许为空
        public string? DisplayName { get; set; }
        // 子分类ID 列表（逗号分隔），允许为空
        public string? SubcategoryIds { get; set; }
        public int SortOrder { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
    }

    public class CadSubcategory
    {
        public int Id { get; set; }
        public int CategoryId { get; set; }
        // 子分类名称，允许为空以防导入数据缺省
        public string? Name { get; set; }
        // 子分类显示名称，允许为空
        public string? DisplayName { get; set; }
        public int ParentId { get; set; }
        public int SortOrder { get; set; }
        public int Level { get; set; }
        // 子分类ID 列表（逗号分隔），允许为空
        public string? SubcategoryIds { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
    }

    public class FileTag
    {
        public int Id { get; set; }
        public int FileId { get; set; }
        // 标签文本，允许为空以兼容历史数据
        public string? Tag { get; set; }
        public DateTime CreatedAt { get; set; }
    }

    public class FileAccessLog
    {
        public int FileId { get; set; }
        // 访问者用户名，允许为空
        public string? UserName { get; set; }
        // 操作类型（Download/View/...），允许为空
        public string? ActionType { get; set; }
        public DateTime AccessTime { get; set; }
        // 访问IP，允许为空
        public string? IpAddress { get; set; }
    }
}
