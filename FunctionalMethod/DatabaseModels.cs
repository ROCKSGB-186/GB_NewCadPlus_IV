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

    /// <summary>
    /// [已废弃] 图元属性模型。之前使用 cad_file_attributes 表存放各种预定义列，
    /// 现已被字典/JSON字符串 (cad_block_attributes_json) 完全取代。
    /// 仅为了防止其它残留类报错暂时保留空壳，或可直接删除。
    /// </summary>
    public class FileAttribute
    {
        // 仅保留Id，让残留能编译
        public long Id { get; set; }
        public string? FileAttributeId { get; set; }
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
