using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.Helpers
{
    public sealed class GraphicApiService
    {
        private static readonly HttpClient HttpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(60)
        };

        public async Task<List<FileStorage>> GetFilesByCategoryIdAsync(
            int categoryId,
            string categoryType,
            CancellationToken cancellationToken = default)
        {
            if (categoryId <= 0)
                throw new ArgumentOutOfRangeException(nameof(categoryId));

            string type;
            if (string.Equals(categoryType, "main", StringComparison.OrdinalIgnoreCase))
            {
                type = "main";
            }
            else if (string.Equals(categoryType, "sub", StringComparison.OrdinalIgnoreCase))
            {
                type = "sub";
            }
            else
            {
                throw new ArgumentException("categoryType 必须是 main 或 sub。", nameof(categoryType));
            }
            string requestUrl = ApiEndpoint.Build($"api/graphics/category/{categoryId}?categoryType={type}");

            using (HttpResponseMessage response = await HttpClient.GetAsync(requestUrl, cancellationToken).ConfigureAwait(false))
            {
                string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new HttpRequestException($"文件接口请求失败，HTTP {(int)response.StatusCode}，响应：{responseBody}");

                GraphicListApiResponse? result = JsonConvert.DeserializeObject<GraphicListApiResponse>(responseBody);
                if (result == null || !result.Success)
                    throw new InvalidOperationException(result?.Message ?? "服务器文件查询失败。");

                result.Files ??= new List<GraphicApiDto>();
                var files = new List<FileStorage>(result.Files.Count);
                foreach (GraphicApiDto item in result.Files)
                {
                    files.Add(new FileStorage
                    {
                        Id = item.Id,
                        CategoryId = item.CategoryId,
                        CategoryType = item.CategoryType,
                        FileAttributeId = item.FileAttributeId,
                        FileName = item.FileName,
                        FileStoredName = item.FileStoredName,
                        DisplayName = string.IsNullOrWhiteSpace(item.DisplayName) ? item.FileName : item.DisplayName,
                        FileType = item.FileType,
                        FileHash = item.FileHash,
                        BlockName = item.BlockName,
                        LayerName = item.LayerName,
                        ColorIndex = item.ColorIndex,
                        Scale = item.Scale,
                        FilePath = item.FilePath,
                        PreviewImageName = item.PreviewImageName,
                        PreviewImagePath = item.PreviewImagePath,
                        FileSize = item.FileSize,
                        IsPreview = item.IsPreview,
                        Version = item.Version,
                        Description = item.Description,
                        IsActive = item.IsActive,
                        CreatedBy = item.CreatedBy,
                        Title = item.Title,
                        Keywords = item.Keywords,
                        IsPublic = item.IsPublic,
                        UpdatedBy = item.UpdatedBy,
                        LastAccessedAt = item.LastAccessedAt,
                        CreatedAt = item.CreatedAt ?? DateTime.MinValue,
                        UpdatedAt = item.UpdatedAt ?? DateTime.MinValue
                    });
                }

                LogManager.Instance.LogInfo($"通过服务器加载文件成功：地址={requestUrl}，文件数={files.Count}");
                return files;
            }
        }

        private sealed class GraphicListApiResponse
        {
            public bool Success { get; set; }
            public string Message { get; set; } = string.Empty;
            public List<GraphicApiDto> Files { get; set; } = new List<GraphicApiDto>();
        }

        private sealed class GraphicApiDto
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
            public string? Title { get; set; }
            public string? Keywords { get; set; }
            public int IsPublic { get; set; }
            public string? UpdatedBy { get; set; }
            public DateTime? LastAccessedAt { get; set; }
            public DateTime? CreatedAt { get; set; }
            public DateTime? UpdatedAt { get; set; }
        }
    }
}
