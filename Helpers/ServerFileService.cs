using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using GB_NewCadPlus_IV.FunctionalMethod; // 引入 FileStorage 等模型
using GB_NewCadPlus_IV.UniFiedStandards; // 引入 VariableDictionary

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 提供与服务器文件交互的统一方法（上传、下载、缓存管理）
    /// </summary>
    public static class ServerFileService
    {
        // 使用单例 HttpClient 避免端口耗尽，并设置合理的超时
        private static readonly HttpClient _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromMinutes(5)
        };

        /// <summary>
        /// 获取服务器基地址
        /// </summary>
        private static string GetServerBaseUrl()
        {
            string ip = VariableDictionary._serverIP ?? "127.0.0.1";
            int apiPort = VariableDictionary._apiPort > 0 ? VariableDictionary._apiPort : 10010;
            return $"http://{ip}:{apiPort}";
        }

        /// <summary>
        /// 构建下载文件的请求 URL
        /// </summary>
        /// <param name="storageId">存储记录 ID</param>
        /// <param name="type">"file" 或 "preview"</param>
        public static string BuildDownloadUrl(int storageId, string type)
        {
            string baseUrl = GetServerBaseUrl();
            string endpoint = type switch
            {
                "file" => $"/api/graphics/{storageId}/file",
                "preview" => $"/api/graphics/{storageId}/preview",
                _ => throw new ArgumentException("无效的下载类型，必须是 'file' 或 'preview'")
            };
            return baseUrl + endpoint;
        }

        /// <summary>
        /// 从服务器下载文件到指定本地路径
        /// </summary>
        /// <param name="storageId">存储记录 ID</param>
        /// <param name="type">"file" 或 "preview"</param>
        /// <param name="localPath">保存的完整本地路径</param>
        /// <returns>是否成功</returns>
        public static async Task<bool> DownloadFileToLocalAsync(int storageId, string type, string localPath)
        {
            try
            {
                string url = BuildDownloadUrl(storageId, type);
                var response = await _httpClient.GetAsync(url);
                response.EnsureSuccessStatusCode();

                // 确保目录存在
                string? dir = Path.GetDirectoryName(localPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                // 使用普通的 using，因为 Stream 未实现 IAsyncDisposable
                using var stream = await response.Content.ReadAsStreamAsync();
                using var fileStream = new FileStream(localPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await stream.CopyToAsync(fileStream);

                return true;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"[ServerFileService] 下载失败 (ID={storageId}, type={type}): {ex.Message}");
                try { if (File.Exists(localPath)) File.Delete(localPath); } catch { }
                return false;
            }
        }

        /// <summary>
        /// 确保预览图在本地有缓存，必要时从服务器下载
        /// </summary>
        /// <param name="fileStorage">文件记录</param>
        /// <param name="previewCacheDir">预览图本地缓存目录</param>
        /// <returns>本地缓存路径，失败返回 null</returns>
        public static async Task<string?> EnsurePreviewCacheAsync(FileStorage fileStorage, string previewCacheDir)
        {
            if (fileStorage == null) return null;

            // 1. 已有路径且存在则直接用
            if (!string.IsNullOrWhiteSpace(fileStorage.PreviewImagePath) && File.Exists(fileStorage.PreviewImagePath))
                return fileStorage.PreviewImagePath;

            // 2. 尝试通过 PreviewImageName 在缓存目录寻找
            if (!string.IsNullOrWhiteSpace(fileStorage.PreviewImageName))
            {
                string cachedPath = Path.Combine(previewCacheDir, $"{fileStorage.Id}_{fileStorage.PreviewImageName}");
                if (File.Exists(cachedPath))
                    return cachedPath;
            }

            // 3. 从服务器下载
            if (fileStorage.Id > 0)
            {
                // 确定扩展名
                string ext = ".png";
                if (!string.IsNullOrWhiteSpace(fileStorage.PreviewImageName))
                {
                    ext = Path.GetExtension(fileStorage.PreviewImageName);
                    if (string.IsNullOrEmpty(ext)) ext = ".png";
                }
                string localName = $"{fileStorage.Id}_preview{ext}";
                string localPath = Path.Combine(previewCacheDir, localName);

                bool ok = await DownloadFileToLocalAsync(fileStorage.Id, "preview", localPath);
                if (ok && File.Exists(localPath))
                    return localPath;
            }

            return null;
        }

        /// <summary>
        /// 确保 DWG 文件在本地有缓存，必要时从服务器下载
        /// </summary>
        /// <param name="fileStorage">文件记录</param>
        /// <param name="dwgCacheDir">DWG 缓存目录</param>
        /// <returns>本地缓存路径，失败返回 null</returns>
        public static async Task<string?> EnsureDwgCacheAsync(FileStorage fileStorage, string dwgCacheDir)
        {
            if (fileStorage == null) return null;

            // 1. 已有路径且存在则直接用
            if (!string.IsNullOrWhiteSpace(fileStorage.FilePath) && File.Exists(fileStorage.FilePath))
                return fileStorage.FilePath;

            // 2. 缓存目录中可能已有，尝试拼合
            if (!string.IsNullOrWhiteSpace(fileStorage.FileName))
            {
                string cachedPath = Path.Combine(dwgCacheDir, $"{fileStorage.Id}_{fileStorage.FileName}");
                if (File.Exists(cachedPath))
                    return cachedPath;
            }
            // 3. 从服务器下载
            if (fileStorage.Id > 0)
            {
                string ext = ".dwg";
                if (!string.IsNullOrWhiteSpace(fileStorage.FileType))
                {
                    ext = fileStorage.FileType;
                    if (!ext.StartsWith(".")) ext = "." + ext;
                }
                string localName = $"{fileStorage.Id}{ext}";
                string localPath = Path.Combine(dwgCacheDir, localName);

                bool ok = await DownloadFileToLocalAsync(fileStorage.Id, "file", localPath);
                if (ok && File.Exists(localPath))
                    return localPath;
            }

            return null;
        }
    }
}
