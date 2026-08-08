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
        public long Id { get; set; }
        public long? ParentId { get; set; }
        public string Code { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int SortOrder { get; set; }
    }

    public sealed class StandardCategoryMoveClientRequest
    {
        public long? ParentId { get; set; }
    }

    public sealed class StandardCategoryDuplicateClient
    {
        public long Id { get; set; }
        public long? ParentId { get; set; }
        public string Code { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int SortOrder { get; set; }
        public string DuplicateReason { get; set; } = string.Empty;
    }

    public sealed class StandardCategoryConflictClientResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public List<StandardCategoryDuplicateClient> Duplicates { get; set; } = new List<StandardCategoryDuplicateClient>();
    }

    public sealed class StandardCategoryConflictException : HttpRequestException
    {
        public StandardCategoryConflictException(string message, List<StandardCategoryDuplicateClient> duplicates)
            : base(message)
        {
            Duplicates = duplicates;
        }

        public List<StandardCategoryDuplicateClient> Duplicates { get; }
    }

    /// <summary>
    /// 规范系列。
    /// </summary>
    public sealed class StandardManagementSeriesClient
    {
        public long Id { get; set; }
        public long? CategoryId { get; set; }
        public string SeriesCode { get; set; } = string.Empty;
        public string SeriesName { get; set; } = string.Empty;
        public string StandardNumber { get; set; } = string.Empty;
        public string TableNumber { get; set; } = string.Empty;
        public string PressureRating { get; set; } = string.Empty;
        public string FlangeType { get; set; } = string.Empty;
        public string FaceType { get; set; } = string.Empty;
    }

    /// <summary>
    /// 规范目录查询结果。
    /// </summary>
    public sealed class StandardManagementTreeClientResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public List<StandardManagementCategoryClient> Categories { get; set; } = new List<StandardManagementCategoryClient>();
        public List<StandardManagementSeriesClient> Series { get; set; } = new List<StandardManagementSeriesClient>();
    }

    /// <summary>
    /// 新建或修改规范主分类、子分类的请求参数。
    /// </summary>
    public sealed class StandardCategoryCommandClientRequest
    {
        public long? ParentId { get; set; }
        public string Code { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public int SortOrder { get; set; }
    }

    public sealed class StandardSeriesMoveClientRequest
    {
        public long CategoryId { get; set; }
    }

    /// <summary>
    /// 规范导入预览结果的客户端摘要。
    /// </summary>
    public sealed class StandardImportPreviewClientResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string BatchId { get; set; } = string.Empty;
        public int ErrorCount { get; set; }
        public int WarningCount { get; set; }
        public int RowCount { get; set; }
    }

    public sealed class StandardImportCommitClientRequest
    {
        public string BatchId { get; set; } = string.Empty;
        public bool AllowWarnings { get; set; }
    }

    public sealed class StandardImportCommitClientResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string BatchId { get; set; } = string.Empty;
        public int ImportedCount { get; set; }
        public int WarningCount { get; set; }
    }

    /// <summary>
    /// 规范版本信息。
    /// </summary>
    public sealed class StandardDocumentVersionClient
    {
        public long Id { get; set; }
        public long SeriesId { get; set; }
        public string VersionNo { get; set; } = string.Empty;
        public string VersionLabel { get; set; } = string.Empty;
        public string ChangeSummary { get; set; } = string.Empty;
        public string SourceType { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public bool IsCurrent { get; set; }
        public DateTime? CreatedAt { get; set; }
        public string CreatedBy { get; set; } = string.Empty;
    }

    /// <summary>
    /// 新建规范版本请求。
    /// </summary>
    public sealed class StandardVersionCreateClientRequest
    {
        public long SeriesId { get; set; }
        public string VersionNo { get; set; } = string.Empty;
        public string VersionLabel { get; set; } = string.Empty;
        public string ChangeSummary { get; set; } = string.Empty;
        public string SourceType { get; set; } = "DOCUMENT";
    }

    /// <summary>
    /// 规范管理操作响应。
    /// </summary>
    public sealed class StandardManagementOperationClientResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public long Id { get; set; }
    }

    /// <summary>
    /// 规范附件上传响应。
    /// </summary>
    public sealed class StandardFileUploadClientResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public long VersionId { get; set; }
        public long FileId { get; set; }
        public string FileName { get; set; } = string.Empty;
    }

    public sealed class StandardDocumentFileClient
    {
        public long Id { get; set; }
        public long VersionId { get; set; }
        public string FileRole { get; set; } = string.Empty;
        public string OriginalFileName { get; set; } = string.Empty;
        public string Extension { get; set; } = string.Empty;
        public string ContentType { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public string Description { get; set; } = string.Empty;
    }
}
