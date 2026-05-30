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
                // 绝对防御：如果 localPath 是一个现有文件夹，先删除
                EnsureIsNotDirectory(localPath);

                string url = BuildDownloadUrl(storageId, type);
                var response = await _httpClient.GetAsync(url);
                response.EnsureSuccessStatusCode();

                string? dir = Path.GetDirectoryName(localPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

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
            
            // 2. 基于 FileHash 的最终路径
            string basePath = BuildCacheFilePath(fileStorage, previewCacheDir, "_preview", ""); // 不带扩展名先
            // 先检查是否已有任何扩展名的文件存在
            foreach (var ext in new[] { ".png", ".jpg", ".jpeg", ".gif", ".bmp" })
            {
                string candidate = basePath + ext;
                if (File.Exists(candidate) && new FileInfo(candidate).Length > 0)
                    return candidate;
            }
            
            // 4. 下载
            // 使用 .png 作为临时下载扩展名，后续再根据文件头修正
            string finalPath = basePath + ".png";
            EnsureIsNotDirectory(finalPath);
            if (File.Exists(finalPath)) File.Delete(finalPath);

            string tempPath = finalPath + ".tmp";
            bool ok = await DownloadFileToLocalAsync(fileStorage.Id, "preview", tempPath);
            if (ok && File.Exists(tempPath) && new FileInfo(tempPath).Length > 0)
            {
                // 根据实际图片类型调整扩展名
                string realExt = GetImageExtension(tempPath);
                string realPath = basePath + realExt;
                SafeReplaceFile(tempPath, realPath);
                return realPath;
            }
            else
            {
                try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
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
            
            // 2. 基于 FileHash 的最终路径
            string finalPath = BuildCacheFilePath(fileStorage, dwgCacheDir, "", ".dwg");

            // 3. 如果本地已有有效文件，直接返回（不下载）
            if (File.Exists(finalPath) && new FileInfo(finalPath).Length > 0)
            {
                LogManager.Instance.LogInfo($"[DWG] 使用已有缓存: {finalPath},文件名: {fileStorage.FileName}");
                return finalPath;
            }
            return null;
        }


        #region 辅助方法
        /// <summary>
        /// 通过文件头判断图片扩展名
        /// </summary>
        private static string GetImageExtension(string filePath)
        {
            try
            {
                byte[] header = new byte[8];
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    fs.Read(header, 0, header.Length);
                if (header[0] == 0xFF && header[1] == 0xD8) return ".jpg";
                if (header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47) return ".png";
                if (header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46) return ".gif";
                if (header[0] == 0x42 && header[1] == 0x4D) return ".bmp";
            }
            catch { }
            return ".png";
        }
       
        /// <summary>
        /// 确保 path 表示一个文件（而非目录）。
        /// 如果 path 是一个已存在的目录，则递归删除该目录。
        /// </summary>
        private static void EnsureIsNotDirectory(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            try
            {
                if (Directory.Exists(path))
                {
                    LogManager.Instance.LogWarning($"[ServerFileService] 发现同名文件夹，正在删除: {path}");
                    Directory.Delete(path, true);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"[ServerFileService] 删除文件夹失败: {path}, {ex.Message}");
            }
        }

        /// <summary>
        /// 安全地将临时文件重命名为最终文件，若目标已是文件夹则先删除。
        /// </summary>
        private static void SafeReplaceFile(string tempPath, string finalPath)
        {
            EnsureIsNotDirectory(finalPath);
            try
            {
                if (File.Exists(finalPath))
                    File.Delete(finalPath);
                File.Move(tempPath, finalPath);
            }
            catch
            {
                // 回退复制 + 删除
                try
                {
                    File.Copy(tempPath, finalPath, true);
                    File.Delete(tempPath);
                }
                catch { }
            }
        }

        /// <summary>
        ///  缓存文件路径构建器：基于 FileHash 生成唯一且稳定的文件名，避免 ID 冲突和旧文件夹问题。
        /// </summary>
        /// <param name="file"> 文件存储信息 </param>
        /// <param name="cacheDir"> 缓存目录 </param>
        /// <param name="suffix"> 后缀 </param>
        /// <param name="extension"> 扩展名 </param>
        /// <returns> 缓存文件路径 </returns>
        private static string BuildCacheFilePath(FileStorage file, string cacheDir, string suffix, string extension)
        {
            string hash = file.FileHash;
            if (string.IsNullOrWhiteSpace(hash))
            {
                // 严重错误，回退使用 ID，但记录日志
                LogManager.Instance.LogError($"[ServerFileService] FileHash 为空！ID={file.Id}，将使用临时文件名。");
                hash = $"id_{file.Id}_nohash";
            }
            string fileName = $"{hash}{suffix}{extension}";
            string fullPath = Path.Combine(cacheDir, fileName);
            LogManager.Instance.LogDebug($"[缓存路径] {fullPath} (服务器中文件名={file.FileName})");
            LogManager.Instance.LogDebug($"[预览缓存路径] {cacheDir} (文件名={fileName})(预览图片名={file.PreviewImageName})");
            return fullPath;
        }
       
        #endregion
    }
}
