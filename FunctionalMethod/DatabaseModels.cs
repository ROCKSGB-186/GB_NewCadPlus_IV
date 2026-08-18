using System;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 定义导入图元的 DTO 模型，包含文件路径、预览图路径、分类信息、图块和图层信息、颜色索引、比例、显示名称、描述、创建人和属性 JSON 等字段，用于在导入过程中传递图元相关数据。兼容旧数据结构，允许部分字段为空以适应历史数据导入。
    /// </summary>
    public class ImportEntityDto
    {
        /// <summary>
        /// 包含文件存储相关信息的对象
        /// </summary>
        public FileStorage FileStorage { get; set; } = new FileStorage();
        /// <summary>
        /// 本地 DWG 文件路径
        /// </summary>
        public string FilePath { get; set; }
        /// <summary>
        /// 本地预览图路径（可选）
        /// </summary>
        public string? PreviewImagePath { get; set; }
        /// <summary>
        /// 分类 ID（必填，> 0）
        /// </summary>
        public int CategoryId { get; set; }
        /// <summary>
        /// 分类类型，默认 "sub"，允许为空以兼容旧数据
        /// </summary>
        public string? CategoryType { get; set; }
        /// <summary>
        /// 图块名
        /// </summary>
        public string? BlockName { get; set; }
        /// <summary>
        /// 图层名
        /// </summary>
        public string? LayerName { get; set; }
        /// <summary>
        /// 颜色索引
        /// </summary>
        public int? ColorIndex { get; set; }
        /// <summary>
        /// 比例
        /// </summary>
        public double? Scale { get; set; }
        /// <summary>
        /// 显示名称
        /// </summary>
        public string? DisplayName { get; set; }
        /// <summary>
        /// 描述
        /// </summary>
        public string? Description { get; set; }
        /// <summary>
        /// 创建人
        /// </summary>
        public string? CreatedBy { get; set; }
        /// <summary>
        /// JSON属性字典（新上传入口使用）
        /// </summary>
        public Dictionary<string, string> AttributesJson { get; set; }
            = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>
        /// 属性业务ID（兼容字段）
        /// </summary>
        public string? FileAttributeId { get; set; }
        /// <summary>
        /// 预览图文件名
        /// </summary>
        public string? PreviewImageName { get; set; }

    }
    /// <summary>
    /// 顶级模型集合，替代对 DatabaseManager 内嵌类型的直接引用
    /// </summary>
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
        public byte[]? FileBytes { get; set; }
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
        /// <summary>
        /// JSON属性字典（新上传入口使用）
        /// </summary>
        public Dictionary<string, string> AttributesJson { get; set; }
            = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 插入图元的业务类型。
    /// 类型由分类树映射得到，Unknown 表示本地旧资源或暂时无法识别的分类。
    /// </summary>
    public enum GraphicEntityType
    {
        Unknown,
        Pipeline,
        Flange,
        Valve,
        Fitting
    }

    /// <summary>
    /// 从图元分类树解析出的插入上下文。
    /// 该对象会随图元按钮一路传递到 AutoCAD 插入命令，避免核心方法依赖文件名猜测类型。
    /// </summary>
    public sealed class GraphicInsertContext
    {
        /// <summary>图元文件 ID。</summary>
        public int FileId { get; set; }

        /// <summary>图元所属分类 ID。</summary>
        public int CategoryId { get; set; }

        /// <summary>图元分类类型，通常为 main 或 sub。</summary>
        public string CategoryType { get; set; } = string.Empty;

        /// <summary>主专业分类 ID。</summary>
        public int MainCategoryId { get; set; }

        /// <summary>主专业分类名称。</summary>
        public string MainCategoryName { get; set; } = string.Empty;

        /// <summary>直接所属分类 ID。</summary>
        public int SubcategoryId { get; set; }

        /// <summary>直接所属分类名称。</summary>
        public string SubcategoryName { get; set; } = string.Empty;

        /// <summary>完整分类路径，例如“工艺/管道”。</summary>
        public string CategoryPath { get; set; } = string.Empty;

        /// <summary>根据分类映射出的业务图元类型。</summary>
        public GraphicEntityType EntityType { get; set; }
    }
    /// <summary>
    /// 同步清单模型，用于描述服务器和客户端文件状态的对比结果，辅助同步决策。
    /// </summary>
    public class SyncManifest
    {
        public string StorageRoot { get; set; } = string.Empty;
        public string SourceRoot { get; set; } = string.Empty;
        public string ServerClientVersion { get; set; } = string.Empty;
        public DateTime GeneratedAt { get; set; }
        public List<SyncManifestItem> Items { get; set; } = new List<SyncManifestItem>();
    }
    /// <summary>
    /// 同步清单项模型，描述单个文件在服务器和客户端的状态对比，包括存在性、哈希值、路径等信息，以及是否需要同步的决策结果。
    /// </summary>
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
    /// <summary>
    /// 兼容之前 DatabaseManager 中使用的 SW 相关模型
    /// </summary>
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
    /// <summary>
    /// SW 子类别模型，包含父类别ID以建立层级关系，兼容旧数据结构，允许名称和显示名称为空以适应历史数据导入。
    /// </summary>
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
    /// <summary>
    /// SW 图形模型，描述软件中的图形元素信息。
    /// </summary>
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
    /// <summary>
    /// cad分类模型，包含分类ID、名称、显示名称、子分类ID列表等信息，兼容旧数据结构，允许名称和显示名称为空以适应历史数据导入。
    /// </summary>
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
    /// <summary>
    /// cad子分类模型，包含子分类ID、父分类ID、名称、显示名称、子分类ID列表等信息，兼容旧数据结构，允许名称和显示名称为空以适应历史数据导入。
    /// </summary>
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
    /// <summary>
    /// 部门信息模型
    /// </summary>
    public class DepartmentModel
    {
        /// <summary>
        ///  id 
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// cad分类id，关联cad_category表的id字段，表示该部门所属的CAD分类。允许为空以兼容旧数据结构和历史数据导入。通过这个字段可以将部门与特定的CAD分类关联起来，便于在系统中进行分类管理和权限控制。
        /// </summary>
        public int? CadCategoryId { get; set; }
        /// <summary>
        /// 部门名称，允许为空以兼容旧数据结构和历史数据导入。这个字段用于存储部门的名称信息，在系统中显示和识别部门。由于历史数据可能存在缺失或不完整的部门名称，因此允许该字段为空以确保数据兼容性和系统稳定性。
        /// </summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>
        /// 真实名称
        /// </summary>
        public string RealName { get; set; } = string.Empty;
        /// <summary>
        /// 显示名称
        /// </summary>
        public string DisplayName { get; set; } = string.Empty;
        /// <summary>
        /// 描述
        /// </summary>
        public string Description { get; set; } = string.Empty;
        /// <summary>
        /// 排序顺序
        /// </summary>
        public int SortOrder { get; set; }
        /// <summary>
        /// 管理员用户id
        /// </summary>
        public int? ManagerUserId { get; set; }        // ★ 新增：部门经理用户ID
        /// <summary>
        /// 是否启用，默认为true。
        /// </summary>
        public bool IsActive { get; set; } = true;
        /// <summary>
        /// 用户数量，表示该部门下的用户总数。
        /// </summary>
        public int UserCount { get; set; }
    }

    /// <summary>
    /// 用户信息模型
    /// </summary>
    public class UserModel
    {
        public int Id { get; set; }
        public string Username { get; set; } = string.Empty;
        public string RealName { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public string Gender { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Phone { get; set; } = string.Empty;
        public string Role { get; set; } = string.Empty;
        public string DepartmentName { get; set; } = string.Empty;
        public bool IsActive { get; set; }
    }
    /// <summary>
    /// 表信息模型（用于数据库浏览）
    /// </summary>
    public class TableInfo
    {
        public string TableName { get; set; } = string.Empty;
        public string TableComment { get; set; } = string.Empty;
    }
}
