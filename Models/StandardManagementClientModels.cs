using System;
using System.Collections.Generic;
using System.Net.Http;

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
        /// <summary>
        /// 表示规范系列列表。
        /// </summary>
        public List<StandardManagementSeriesClient> Series { get; set; } = new List<StandardManagementSeriesClient>();
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
    /// 规范导入预览中的单行错误明细。
    /// </summary>
    public sealed class StandardImportPreviewRowClient
    {
        /// <summary>
        /// 表示该行数据在导入文件中的行号。
        /// </summary>
        public int RowNumber { get; set; }
        /// <summary>
        /// 表示该行数据在导入文件中的错误信息列表。
        /// </summary>
        public List<string> Errors { get; set; } = new List<string>();
        /// <summary>
        /// 表示该行数据在导入文件中的警告信息列表。
        /// </summary>
        public List<string> Warnings { get; set; } = new List<string>();
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
