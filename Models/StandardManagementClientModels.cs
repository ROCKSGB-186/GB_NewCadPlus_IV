using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Linq;

namespace GB_NewCadPlus_IV.Models
{
    /// <summary>
    /// 规范管理目录节点。
    /// </summary>
    public sealed class StandardManagementCategoryClient
    {
        /// <summary>
        /// 规范分类的唯一标识符。
        /// </summary>
        public long Id { get; set; }
        /// <summary>
        /// 规范分类的父分类的唯一标识符，如果该分类是顶级分类，则为 null。
        /// </summary>
        public long? ParentId { get; set; }
        /// <summary>
        /// 规范分类的编码，用于唯一标识该分类。
        /// </summary>
        public string Code { get; set; } = string.Empty;
        /// <summary>
        /// 规范分类的名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>
        /// 规范分类的描述。
        /// </summary>
        public string Description { get; set; } = string.Empty;
        /// <summary>
        /// 规范分类的排序顺序。
        /// </summary>
        public int SortOrder { get; set; }
    }

    /// <summary>
    /// 修改动态规范细分显示名称的请求。
    /// </summary>
    public sealed class StandardVersionRenameClientRequest
    {
        public string Name { get; set; } = string.Empty;
    }

    /// <summary>动态模板预览响应。</summary>
    public sealed class DynamicStandardPreviewClientResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public bool IsTemplateMatched { get; set; }
        public StandardTemplateClient Template { get; set; }
        public List<DynamicStandardPreviewColumnClient> Columns { get; set; } = new List<DynamicStandardPreviewColumnClient>();
        public List<DynamicStandardPreviewRowClient> Rows { get; set; } = new List<DynamicStandardPreviewRowClient>();
        public List<string> UnmappedHeaders { get; set; } = new List<string>();
        public int ErrorCount { get; set; }
        public int WarningCount { get; set; }
        public StandardTemplateDraftClient TemplateDraft { get; set; }
        public DynamicStandardDifferenceClient Difference { get; set; }
        public bool HasExistingVersion { get; set; }
        public long? ExistingVersionId { get; set; }
        public List<string> CandidateUniqueKeyFields { get; set; } = new List<string>();
    }

    /// <summary>动态模板信息。</summary>
    public sealed class StandardTemplateClient
    {
        public long Id { get; set; }
        public string TemplateCode { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public string FamilyCode { get; set; } = string.Empty;
        public string FileType { get; set; } = string.Empty;
        public int Version { get; set; }
    }

    /// <summary>动态预览列定义。</summary>
    public sealed class DynamicStandardPreviewColumnClient
    {
        public string Header { get; set; } = string.Empty;
        public string FieldCode { get; set; } = string.Empty;
        public string FieldName { get; set; } = string.Empty;
        public string DataType { get; set; } = "TEXT";
        public string Unit { get; set; } = string.Empty;
        public bool IsRequired { get; set; }
        public bool IsMapped { get; set; }
    }

    /// <summary>动态预览行。</summary>
    public sealed class DynamicStandardPreviewRowClient
    {
        public int RowNumber { get; set; }
        public Dictionary<string, string> Values { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public List<string> Errors { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
    }

    /// <summary>动态规范导入确认请求。</summary>
    public sealed class DynamicStandardImportCommitClientRequest
    {
        public string BatchId { get; set; } = string.Empty;
        public long SeriesId { get; set; }
        public long? StandardDocumentId { get; set; }
        public long? CategoryId { get; set; }
        public string BaseStandardNumber { get; set; } = string.Empty;
        public string BaseStandardName { get; set; } = string.Empty;
        /// <summary>规范所属专业编码，例如 FLANGE。</summary>
        public string FamilyCode { get; set; } = string.Empty;
        public string SeriesName { get; set; } = string.Empty;
        public string StandardNumber { get; set; } = string.Empty;
        public string SeriesCode { get; set; } = string.Empty;
        public string TableNumber { get; set; } = string.Empty;
        public string PressureRating { get; set; } = string.Empty;
        public long? VersionId { get; set; }
        public long? TemplateId { get; set; }
        public string SourceFileName { get; set; } = string.Empty;
        public string SourceFileSha256 { get; set; } = string.Empty;
        public bool AllowWarnings { get; set; }
        public string UpdateStrategy { get; set; } = "REPLACE";
        public bool ConfirmTemplateCreation { get; set; }
        public StandardTemplateDraftClient TemplateDraft { get; set; }
        public List<string> UniqueKeyFields { get; set; } = new List<string>();
        public Dictionary<string, string> ConflictDecisions { get; set; } = new Dictionary<string, string>();
        public List<DynamicStandardPreviewRowClient> Rows { get; set; } = new List<DynamicStandardPreviewRowClient>();
    }

    /// <summary>首次上传生成的动态模板草稿。</summary>
    public sealed class StandardTemplateDraftClient
    {
        public string TemplateCode { get; set; } = string.Empty;
        public string TemplateName { get; set; } = string.Empty;
        public string FamilyCode { get; set; } = string.Empty;
        public string FileType { get; set; } = "XLSX";
        public List<StandardTemplateDraftColumnClient> Columns { get; set; } = new List<StandardTemplateDraftColumnClient>();
    }

    /// <summary>模板字段草稿。</summary>
    public sealed class StandardTemplateDraftColumnClient
    {
        public string Header { get; set; } = string.Empty;
        public string FieldCode { get; set; } = string.Empty;
        public string FieldName { get; set; } = string.Empty;
        public string DataType { get; set; } = "TEXT";
        public bool IsRequired { get; set; }
        public int SortOrder { get; set; }
    }

    /// <summary>动态规范差异摘要。</summary>
    public sealed class DynamicStandardDifferenceClient
    {
        public List<string> AddedHeaders { get; set; } = new List<string>();
        public List<string> RemovedHeaders { get; set; } = new List<string>();
        public List<string> ChangedHeaders { get; set; } = new List<string>();
        public int AddedRows { get; set; }
        public int RemovedRows { get; set; }
        public int ChangedRows { get; set; }
        public List<string> ConflictRows { get; set; } = new List<string>();
    }

    /// <summary>动态规范导入确认响应。</summary>
    public sealed class DynamicStandardImportCommitClientResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string BatchId { get; set; } = string.Empty;
        public int SavedRowCount { get; set; }
        public string Status { get; set; } = string.Empty;
    }
    /// <summary>
    /// 规范分类到新的父分类的请求参数。
    /// </summary>
    public sealed class StandardCategoryMoveClientRequest
    {
        /// <summary>
        /// 要移动的规范分类的唯一标识符。
        /// </summary>
        public long? ParentId { get; set; }
    }

    /// <summary>
    /// 调整规范分类显示顺序的请求参数。
    /// Direction=-1 表示上移，Direction=1 表示下移。
    /// </summary>
    public sealed class StandardCategoryReorderClientRequest
    {
        /// <summary>
        /// 要调整顺序的规范分类的唯一标识符。
        /// </summary>
        public int Direction { get; set; }
    }
    /// <summary>
    /// 规范分类重复信息。
    /// </summary>
    public sealed class StandardCategoryDuplicateClient
    {
        /// <summary>
        /// 要移动的规范分类的唯一标识符。
        /// </summary>
        public long Id { get; set; }
        /// <summary>
        /// 要移动的规范分类的父分类的唯一标识符。
        /// </summary>
        public long? ParentId { get; set; }
        /// <summary>
        /// 要移动的规范分类的编码。
        /// </summary>
        public string Code { get; set; } = string.Empty;
        /// <summary>
        /// 要移动的规范分类的名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>
        /// 要移动的规范分类的描述。
        /// </summary>
        public string Description { get; set; } = string.Empty;
        /// <summary>
        /// 要移动的规范分类的排序顺序。
        /// </summary>
        public int SortOrder { get; set; }
        /// <summary>
        /// 规范分类重复的原因。
        /// </summary>
        public string DuplicateReason { get; set; } = string.Empty;
    }
    /// <summary>
    /// 规范分类冲突响应。
    /// </summary>
    public sealed class StandardCategoryConflictClientResponse
    {
        /// <summary>
        /// 表示操作是否成功。如果为 true，则表示操作成功；如果为 false，则表示操作失败，并且可能存在冲突。
        /// </summary>
        public bool Success { get; set; }
        /// <summary>
        /// 表示操作的结果消息。如果 Success 为 false，则该消息可能包含冲突的详细信息。
        /// </summary>
        public string Message { get; set; } = string.Empty;
        /// <summary>
        /// 表示冲突的规范分类列表。如果 Success 为 false，则该列表可能包含导致冲突的规范分类信息。
        /// </summary>
        public List<StandardCategoryDuplicateClient> Duplicates { get; set; } = new List<StandardCategoryDuplicateClient>();
    }
    /// <summary>
    /// 规范分类冲突异常。
    /// </summary>
    public sealed class StandardCategoryConflictException : HttpRequestException
    {
        /// <summary>
        /// 初始化一个新的 <see cref="StandardCategoryConflictException"/> 实例。
        /// </summary>
        /// <param name="message">异常消息。</param>
        /// <param name="duplicates">导致冲突的规范分类列表。</param>
        public StandardCategoryConflictException(string message, List<StandardCategoryDuplicateClient> duplicates)
            : base(message)
        {
            Duplicates = duplicates;
        }
        /// <summary>
        /// 导致冲突的规范分类列表。
        /// </summary>
        public List<StandardCategoryDuplicateClient> Duplicates { get; }
    }

    /// <summary>
    /// 规范系列信息。
    /// </summary>
    public sealed class StandardManagementSeriesClient
    {
        /// <summary>
        /// 规范系列的唯一标识符。
        /// </summary>
        public long Id { get; set; }
        /// <summary>
        /// 规范系列所属的分类的唯一标识符。
        /// </summary>
        public long? CategoryId { get; set; }
        /// <summary>所属基础规范号记录的 ID。</summary>
        public long? StandardDocumentId { get; set; }
        /// <summary>
        /// 规范所属专业编码，例如 FLANGE。
        /// </summary>
        public string FamilyCode { get; set; } = string.Empty;
        /// <summary>
        /// 规范系列的编码。
        /// </summary>
        public string SeriesCode { get; set; } = string.Empty;
        /// <summary>
        /// 规范系列的名称。
        /// </summary>
        public string SeriesName { get; set; } = string.Empty;
        /// <summary>
        /// 规范系列的标准编号。
        /// </summary>
        public string StandardNumber { get; set; } = string.Empty;
        /// <summary>
        /// 规范系列的表编号。
        /// </summary>
        public string TableNumber { get; set; } = string.Empty;
        /// <summary>
        /// 规范系列的压力等级。
        /// </summary>
        public string PressureRating { get; set; } = string.Empty;
        /// <summary>
        /// 规范系列的法兰类型。
        /// </summary>
        public string FlangeType { get; set; } = string.Empty;
        /// <summary>
        /// 规范系列的面类型。
        /// </summary>
        public string FaceType { get; set; } = string.Empty;

        /// <summary>
        /// 下拉菜单显示文本，不参与数据库匹配。
        /// </summary>
        public string DisplayText => string.Join(" / ", new[] { SeriesName, StandardNumber }
            .Where(value => !string.IsNullOrWhiteSpace(value)));
    }

    /// <summary>
    /// 规范目录查询结果。
    /// </summary>
    public sealed class StandardManagementTreeClientResponse
    {
        /// <summary>
        /// 表示操作是否成功。如果为 true，则表示操作成功；如果为 false，则表示操作失败。
        /// </summary>
        public bool Success { get; set; }
        /// <summary>
        /// 表示操作的结果消息。
        /// </summary>
        public string Message { get; set; } = string.Empty;
        /// <summary>
        /// 表示规范分类列表。
        /// </summary>
        public List<StandardManagementCategoryClient> Categories { get; set; } = new List<StandardManagementCategoryClient>();
        /// <summary>基础规范号列表。</summary>
        public List<StandardDocumentClient> Documents { get; set; } = new List<StandardDocumentClient>();
        /// <summary>
        /// 表示规范系列列表。
        /// </summary>
        public List<StandardManagementSeriesClient> Series { get; set; } = new List<StandardManagementSeriesClient>();
    }

    /// <summary>客户端基础规范号模型。</summary>
    public sealed class StandardDocumentClient
    {
        public long Id { get; set; }
        public string FamilyCode { get; set; } = string.Empty;
        public long? CategoryId { get; set; }
        public string StandardNumber { get; set; } = string.Empty;
        public string StandardName { get; set; } = string.Empty;
        public bool IsActive { get; set; }
        public string DisplayText => string.IsNullOrWhiteSpace(StandardName)
            || string.Equals(StandardName.Trim(), StandardNumber.Trim(), StringComparison.OrdinalIgnoreCase)
            ? StandardNumber
            : $"{StandardNumber} / {StandardName}";
    }

    /// <summary>
    /// 新建或修改规范主分类、子分类的请求参数。
    /// </summary>
    public sealed class StandardCategoryCommandClientRequest
    {
        /// <summary>
        /// 要新建或修改的规范分类的唯一标识符。如果为 null，则表示新建分类；如果不为 null，则表示修改已有分类。
        /// </summary>
        public long? ParentId { get; set; }
        /// <summary>
        /// 规范分类的编码。
        /// </summary>
        public string Code { get; set; } = string.Empty;
        /// <summary>
        /// 规范分类的名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>
        /// 规范分类的描述。
        /// </summary>
        public string Description { get; set; } = string.Empty;
        /// <summary>
        /// 规范分类的排序顺序。
        /// </summary>
        public int SortOrder { get; set; }
    }
    /// <summary>
    /// 将规范系列移动到新的分类的请求参数。
    /// </summary>
    public sealed class StandardSeriesMoveClientRequest
    {
        /// <summary>
        /// 要移动到的目标分类的唯一标识符。
        /// </summary>
        public long CategoryId { get; set; }
    }

    /// <summary>
    /// 修改规范系列显示名称的请求参数。
    /// </summary>
    public sealed class StandardSeriesRenameClientRequest
    {
        /// <summary>
        /// 要修改的规范系列的名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;
    }

    /// <summary>
    /// 规范导入预览中的规范元数据。
    /// </summary>
    public sealed class StandardImportSeriesClient
    {
        /// <summary>规范大类编码。</summary>
        public string FamilyCode { get; set; } = string.Empty;
        /// <summary>规范大类名称。</summary>
        public string FamilyName { get; set; } = string.Empty;
        /// <summary>规范系列编码。</summary>
        public string SeriesCode { get; set; } = string.Empty;
        /// <summary>规范注释名称。</summary>
        public string SeriesName { get; set; } = string.Empty;
        /// <summary>规范代号。</summary>
        public string StandardNumber { get; set; } = string.Empty;
        /// <summary>表号。</summary>
        public string TableNumber { get; set; } = string.Empty;
        /// <summary>型号或压力等级。</summary>
        public string PressureRating { get; set; } = string.Empty;
    }

    /// <summary>
    /// 规范导入预览中的一条完整法兰规范记录。
    /// </summary>
    public sealed class StandardImportRecordClient
    {
        /// <summary>公称尺寸。</summary>
        public string DN { get; set; } = string.Empty;
        /// <summary>公称尺寸数值。</summary>
        public int DNValue { get; set; }
        /// <summary>压力等级。</summary>
        public string PN { get; set; } = string.Empty;
        /// <summary>钢管外径，系列 I。</summary>
        public decimal? PipeOuterDiameterSeriesI { get; set; }
        /// <summary>钢管外径，系列 II。</summary>
        public decimal? PipeOuterDiameterSeriesII { get; set; }
        /// <summary>法兰外径。</summary>
        public decimal? FlangeOuterDiameter { get; set; }
        /// <summary>螺栓孔中心圆直径。</summary>
        public decimal? BoltCircleDiameter { get; set; }
        /// <summary>螺栓孔直径。</summary>
        public decimal? BoltHoleDiameter { get; set; }
        /// <summary>螺栓数量。</summary>
        public int? BoltCount { get; set; }
        /// <summary>螺栓规格。</summary>
        public string BoltSpecification { get; set; } = string.Empty;
        /// <summary>法兰厚度。</summary>
        public decimal? FlangeThickness { get; set; }
        /// <summary>突面高度。</summary>
        public decimal? RaisedFaceHeight { get; set; }
    }

    /// <summary>
    /// 预览阶段命中的已有规范系列。
    /// </summary>
    public sealed class StandardImportDuplicateSeriesClient
    {
        /// <summary>已存在规范系列 ID。</summary>
        public long SeriesId { get; set; }
        /// <summary>已存在规范显示名称。</summary>
        public string SeriesName { get; set; } = string.Empty;
        /// <summary>已存在规范代号。</summary>
        public string StandardNumber { get; set; } = string.Empty;
        /// <summary>已存在规范表号。</summary>
        public string TableNumber { get; set; } = string.Empty;
        /// <summary>已存在规范型号。</summary>
        public string PressureRating { get; set; } = string.Empty;
        /// <summary>已有有效记录数量。</summary>
        public int RecordCount { get; set; }
    }

    /// <summary>
    /// 规范导入预览中的单行数据和校验结果。
    /// </summary>
    public sealed class StandardImportPreviewRowClient
    {
        /// <summary>
        /// 表示该行数据在导入文件中的行号。
        /// </summary>
        public int RowNumber { get; set; }
        /// <summary>该行解析出的完整规范记录。</summary>
        public StandardImportRecordClient? Record { get; set; }
        /// <summary>
        /// 表示该行数据在导入文件中的错误信息列表。
        /// </summary>
        public List<string> Errors { get; set; } = new List<string>();
        /// <summary>
        /// 表示该行数据在导入文件中的警告信息列表。
        /// </summary>
        public List<string> Warnings { get; set; } = new List<string>();

        /// <summary>该行是否存在错误。</summary>
        public bool HasErrors => Errors.Count > 0;
        /// <summary>该行是否存在警告。</summary>
        public bool HasWarnings => Warnings.Count > 0;
        /// <summary>该行的校验状态文本。</summary>
        public string ValidationStatus => HasErrors ? "错误" : HasWarnings ? "警告" : "正常";
        /// <summary>该行错误的显示文本。</summary>
        public string ErrorText => string.Join("；", Errors);
        /// <summary>该行警告的显示文本。</summary>
        public string WarningText => string.Join("；", Warnings);
    }

    /// <summary>
    /// 规范导入预览结果。
    /// </summary>
    public sealed class StandardImportPreviewClientResponse
    {
        /// <summary>
        /// 表示操作是否成功。如果为 true，则表示操作成功；如果为 false，则表示操作失败。
        /// </summary>
        public bool Success { get; set; }
        /// <summary>
        /// 表示操作的结果消息。
        /// </summary>
        public string Message { get; set; } = string.Empty;
        /// <summary>
        /// 表示导入批次的唯一标识符。
        /// </summary>
        public string BatchId { get; set; } = string.Empty;
        /// <summary>用户选择的原始规范文件名。</summary>
        public string SourceFileName { get; set; } = string.Empty;
        /// <summary>本次导入的规范元数据。</summary>
        public StandardImportSeriesClient Series { get; set; } = new StandardImportSeriesClient();
        /// <summary>同一业务标识下已存在的规范系列；未命中时为 null。</summary>
        public StandardImportDuplicateSeriesClient? DuplicateSeries { get; set; }
        /// <summary>
        /// 表示导入过程中发生的错误数量。
        /// </summary>
        public int ErrorCount { get; set; }
        /// <summary>
        /// 表示导入过程中发生的警告数量。
        /// </summary>
        public int WarningCount { get; set; }
        /// <summary>
        /// 表示导入文件中的行数。
        /// </summary>
        public int RowCount { get; set; }
        /// <summary>
        /// 表示导入文件中的行数据列表。
        /// </summary>
        public List<StandardImportPreviewRowClient> Rows { get; set; } = new List<StandardImportPreviewRowClient>();
    }
    /// <summary>
    /// 规范导入提交请求。
    /// </summary>
    public sealed class StandardImportCommitClientRequest
    {
        /// <summary>
        /// 表示要提交的导入批次的唯一标识符。
        /// </summary>
        public string BatchId { get; set; } = string.Empty;
        /// <summary>
        /// 表示是否允许提交包含警告的导入批次。
        /// </summary>
        public bool AllowWarnings { get; set; }
        /// <summary>用户最终确认的规范元数据。</summary>
        public StandardImportSeriesClient Series { get; set; } = new StandardImportSeriesClient();
        /// <summary>命中已有规范时的处理策略。</summary>
        public StandardImportDuplicateStrategyClient DuplicateStrategy { get; set; } = StandardImportDuplicateStrategyClient.NewImport;
    }

    /// <summary>
    /// 客户端向服务器声明的同名规范处理策略。
    /// </summary>
    public enum StandardImportDuplicateStrategyClient
    {
        /// <summary>仅导入未命中旧规范的新规范。</summary>
        NewImport = 0,
        /// <summary>请求创建新版本，当前版本快照完成前不可用。</summary>
        CreateVersion = 1,
        /// <summary>请求覆盖当前数据，当前版本快照完成前不可用。</summary>
        OverwriteCurrent = 2
    }
    /// <summary>
    /// 规范导入提交响应。
    /// </summary>
    public sealed class StandardImportCommitClientResponse
    {
        /// <summary>
        /// 表示操作是否成功。如果为 true，则表示操作成功；如果为 false，则表示操作失败。
        /// </summary>
        public bool Success { get; set; }
        /// <summary>
        /// 表示操作的结果消息。
        /// </summary>
        public string Message { get; set; } = string.Empty;
        /// <summary>
        /// 表示导入批次的唯一标识符。
        /// </summary>
        public string BatchId { get; set; } = string.Empty;
        /// <summary>
        /// 表示导入的记录数量。
        /// </summary>
        public int ImportedCount { get; set; }
        /// <summary>
        /// 表示导入过程中发生的警告数量。
        /// </summary>
        public int WarningCount { get; set; }
    }

    /// <summary>
    /// 规范版本信息。
    /// </summary>
    public sealed class StandardDocumentVersionClient
    {
        /// <summary>
        /// 规范版本的唯一标识符ID。
        /// </summary>
        public long Id { get; set; }
        /// <summary>
        /// 系列Id
        /// </summary>  
        public long SeriesId { get; set; }
        /// <summary>
        /// 规范版本号
        /// </summary>
        public string VersionNo { get; set; } = string.Empty;
        /// <summary>
        /// 规范版本标签
        /// </summary>
        public string VersionLabel { get; set; } = string.Empty;
        /// <summary>
        /// 变更摘要
        /// </summary>
        public string ChangeSummary { get; set; } = string.Empty;
        /// <summary>
        /// 来源类型
        /// </summary>
        public string SourceType { get; set; } = string.Empty;
        /// <summary>
        /// 状态
        /// </summary>
        public string Status { get; set; } = string.Empty;
        /// <summary>
        /// 是否为当前版本
        /// </summary>
        public bool IsCurrent { get; set; }
        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime? CreatedAt { get; set; }
        /// <summary>
        /// 创建人
        /// </summary>
        public string CreatedBy { get; set; } = string.Empty;
    }

    /// <summary>
    /// 新建规范版本请求。
    /// </summary>
    public sealed class StandardVersionCreateClientRequest
    {
        /// <summary>
        /// 规范系列的唯一标识符ID。
        /// </summary>
        public long SeriesId { get; set; }
        /// <summary>
        /// 规范版本号
        /// </summary>
        public string VersionNo { get; set; } = string.Empty;
        /// <summary>
        /// 规范版本标签
        /// </summary>
        public string VersionLabel { get; set; } = string.Empty;
        /// <summary>
        /// 变更摘要
        /// </summary>
        public string ChangeSummary { get; set; } = string.Empty;
        /// <summary>
        /// 来源类型
        /// </summary>
        public string SourceType { get; set; } = "DOCUMENT";
    }

    /// <summary>
    /// 规范管理操作响应。
    /// </summary>
    public sealed class StandardManagementOperationClientResponse
    {
        /// <summary>
        /// 表示操作是否成功。如果为 true，则表示操作成功；如果为 false，则表示操作失败。
        /// </summary>
        public bool Success { get; set; }
        /// <summary>
        /// 表示操作的结果消息。
        /// </summary>
        public string Message { get; set; } = string.Empty;
        /// <summary>
        /// 操作的唯一标识符ID。
        /// </summary>
        public long Id { get; set; }
    }

    /// <summary>
    /// 规范附件上传响应。
    /// </summary>
    public sealed class StandardFileUploadClientResponse
    {
        /// <summary>
        /// 表示操作是否成功。如果为 true，则表示操作成功；如果为 false，则表示操作失败。
        /// </summary>
        public bool Success { get; set; }
        /// <summary>
        /// 表示操作的结果消息。
        /// </summary>
        public string Message { get; set; } = string.Empty;
        /// <summary>
        /// 规范版本的唯一标识符ID。
        /// </summary>
        public long VersionId { get; set; }
        /// <summary>
        /// 规范附件的唯一标识符ID。
        /// </summary>
        public long FileId { get; set; }
        /// <summary>
        /// 规范附件的文件名。
        /// </summary>
        public string FileName { get; set; } = string.Empty;
    }
    /// <summary>
    /// 规范附件信息。
    /// </summary>
    public sealed class StandardDocumentFileClient
    {
        /// <summary>
        /// 规范附件的唯一标识符ID。
        /// </summary>
        public long Id { get; set; }
        /// <summary>
        /// 规范版本的唯一标识符ID。
        /// </summary>
        public long VersionId { get; set; }
        /// <summary>
        /// 规范附件的角色。
        /// </summary>
        public string FileRole { get; set; } = string.Empty;
        /// <summary>
        /// 规范附件的原始文件名。
        /// </summary>  
        public string OriginalFileName { get; set; } = string.Empty;
        /// <summary>
        /// 规范附件的扩展名。
        /// </summary>  
        public string Extension { get; set; } = string.Empty;
        /// <summary>
        /// 规范附件的内容类型（MIME 类型）。
        /// </summary>
        public string ContentType { get; set; } = string.Empty;
        /// <summary>
        /// 规范附件的文件大小（以字节为单位）。
        /// </summary>
        public long FileSize { get; set; }
        /// <summary>
        /// 规范附件的上传时间。
        /// </summary>
        public DateTime UploadedAt { get; set; }
        /// <summary>
        /// 规范附件的描述。
        /// </summary>
        public string Description { get; set; } = string.Empty;
    }
}
