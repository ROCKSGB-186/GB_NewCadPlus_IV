using GB_NewCadPlus_IV.Models;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 客户端规范库 API 服务。
    /// 只通过 HTTP 查询服务器规范库，不直接访问数据库。
    /// </summary>
    public sealed class StandardApiService
    {
        /// <summary>
        /// 全局复用 HttpClient，避免每次查询都创建新的连接。
        /// </summary>
        private static readonly HttpClient HttpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(30)
        };

        /// <summary>
        /// 查询规范管理目录树。
        /// </summary>
        public async Task<StandardManagementTreeClientResponse> GetManagementTreeAsync(
            CancellationToken cancellationToken = default(CancellationToken))
        {
            string requestUrl = BuildServerUrl("/api/standards/management/tree");
            try
            {
                using (HttpResponseMessage response = await HttpClient.GetAsync(requestUrl, cancellationToken).ConfigureAwait(false))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                        throw new HttpRequestException($"规范目录查询失败，HTTP {(int)response.StatusCode}，响应：{body}");

                    StandardManagementTreeClientResponse? result = JsonConvert.DeserializeObject<StandardManagementTreeClientResponse>(body);
                    return result ?? throw new InvalidOperationException("服务器返回的规范目录为空。");
                }

            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"规范目录查询异常：地址={requestUrl}，错误={ex.Message}");
                throw;
            }
        }

        public async Task<DynamicStandardContentClientResponse?> GetDynamicContentAsync(
            long seriesId,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (seriesId <= 0)
                throw new ArgumentException("规范系列 ID 必须大于 0。", nameof(seriesId));

            string requestUrl = BuildServerUrl($"/api/standards/dynamic/series/{seriesId}/content");
            using (HttpResponseMessage response = await HttpClient.GetAsync(requestUrl, cancellationToken).ConfigureAwait(false))
            {
                string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if ((int)response.StatusCode == 404)
                    return null;
                if (!response.IsSuccessStatusCode)
                    throw new HttpRequestException($"动态规范内容查询失败，HTTP {(int)response.StatusCode}，响应：{responseBody}");
                return JsonConvert.DeserializeObject<DynamicStandardContentClientResponse>(responseBody);
            }
        }

        public async Task<StandardManagementOperationClientResponse> RenameManagementVersionAsync(
            long versionId,
            string name,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (versionId <= 0) throw new ArgumentException("规范版本 ID 必须大于 0。", nameof(versionId));
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("规范版本名称不能为空。", nameof(name));

            string requestJson = JsonConvert.SerializeObject(new StandardVersionRenameClientRequest { Name = name.Trim() });
            using var content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            using HttpRequestMessage message = new HttpRequestMessage(
                HttpMethod.Put,
                BuildServerUrl($"/api/standards/management/versions/{versionId}/name"))
            {
                Content = content
            };
            AddOperatorHeader(message, operatorName);
            return await SendManagementRequestAsync<StandardManagementOperationClientResponse>(
                message, "重命名动态规范细分", cancellationToken).ConfigureAwait(false);
        }

        public async Task<DynamicStandardContentClientResponse?> GetDynamicContentByVersionAsync(
            long versionId,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (versionId <= 0)
                throw new ArgumentException("规范版本 ID 必须大于 0。", nameof(versionId));

            string requestUrl = BuildServerUrl($"/api/standards/dynamic/versions/{versionId}/content");
            using (HttpResponseMessage response = await HttpClient.GetAsync(requestUrl, cancellationToken).ConfigureAwait(false))
            {
                string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if ((int)response.StatusCode == 404)
                    return null;
                if (!response.IsSuccessStatusCode)
                    throw new HttpRequestException($"动态规范版本内容查询失败，HTTP {(int)response.StatusCode}，响应：{responseBody}");
                return JsonConvert.DeserializeObject<DynamicStandardContentClientResponse>(responseBody);
            }
        }

        /// <summary>
        /// 通过服务器调整规范分类在同层级中的显示顺序。
        /// </summary>
        public async Task<StandardManagementOperationClientResponse> ReorderManagementCategoryAsync(
            long categoryId,
            int direction,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (categoryId <= 0) throw new ArgumentException("分类 ID 必须大于 0。", nameof(categoryId));
            if (direction != -1 && direction != 1) throw new ArgumentException("排序方向必须是 -1 或 1。", nameof(direction));

            string requestJson = JsonConvert.SerializeObject(new StandardCategoryReorderClientRequest
            {
                Direction = direction
            });
            using var content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            using HttpRequestMessage message = new HttpRequestMessage(
                HttpMethod.Post,
                BuildServerUrl($"/api/standards/management/categories/{categoryId}/reorder"))
            {
                Content = content
            };
            AddOperatorHeader(message, operatorName);
            return await SendManagementRequestAsync<StandardManagementOperationClientResponse>(
                message, direction < 0 ? "上移规范分类" : "下移规范分类", cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// 创建规范主分类或子分类。
        /// </summary>
        public async Task<StandardManagementOperationClientResponse> CreateManagementCategoryAsync(
            StandardCategoryCommandClientRequest request,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            string requestUrl = BuildServerUrl("/api/standards/management/categories");
            string requestJson = JsonConvert.SerializeObject(request);
            using var content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            using HttpRequestMessage message = new HttpRequestMessage(HttpMethod.Post, requestUrl)
            {
                Content = content
            };
            AddOperatorHeader(message, operatorName);
            using (HttpResponseMessage response = await HttpClient.SendAsync(message, cancellationToken).ConfigureAwait(false))
            {
                string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                {
                    if ((int)response.StatusCode == 409)
                    {
                        StandardCategoryConflictClientResponse? conflict = JsonConvert.DeserializeObject<StandardCategoryConflictClientResponse>(body);
                        if (conflict != null)
                            throw new StandardCategoryConflictException(conflict.Message, conflict.Duplicates);
                    }

                    throw new HttpRequestException($"创建规范分类失败，HTTP {(int)response.StatusCode}，响应：{body}");
                }

                return JsonConvert.DeserializeObject<StandardManagementOperationClientResponse>(body)
                    ?? throw new InvalidOperationException("服务器返回的创建规范分类响应为空。");
            }
        }

        public async Task<StandardManagementOperationClientResponse> UpdateManagementCategoryAsync(
            long categoryId,
            StandardCategoryCommandClientRequest request,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (categoryId <= 0) throw new ArgumentException("分类 ID 必须大于 0。", nameof(categoryId));
            string requestUrl = BuildServerUrl($"/api/standards/management/categories/{categoryId}");
            using var content = new StringContent(JsonConvert.SerializeObject(request), Encoding.UTF8, "application/json");
            using HttpRequestMessage message = new HttpRequestMessage(HttpMethod.Put, requestUrl) { Content = content };
            AddOperatorHeader(message, operatorName);
            using (HttpResponseMessage response = await HttpClient.SendAsync(message, cancellationToken).ConfigureAwait(false))
            {
                string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                {
                    if ((int)response.StatusCode == 409)
                    {
                        StandardCategoryConflictClientResponse? conflict = JsonConvert.DeserializeObject<StandardCategoryConflictClientResponse>(body);
                        if (conflict != null)
                            throw new StandardCategoryConflictException(conflict.Message, conflict.Duplicates);
                    }

                    throw new HttpRequestException($"修改规范分类失败，HTTP {(int)response.StatusCode}，响应：{body}");
                }

                return JsonConvert.DeserializeObject<StandardManagementOperationClientResponse>(body)
                    ?? throw new InvalidOperationException("服务器返回的修改规范分类响应为空。");
            }
        }

        /// <summary>
        /// 通过服务器软删除规范分类，不在客户端直接操作数据库。
        /// </summary>
        public async Task<StandardManagementOperationClientResponse> DeleteManagementCategoryAsync(
            long categoryId,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (categoryId <= 0) throw new ArgumentException("分类 ID 必须大于 0。", nameof(categoryId));
            using HttpRequestMessage message = new HttpRequestMessage(
                HttpMethod.Delete,
                BuildServerUrl($"/api/standards/management/categories/{categoryId}"));
            AddOperatorHeader(message, operatorName);
            return await SendManagementRequestAsync<StandardManagementOperationClientResponse>(
                message, "删除规范分类", cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// 通过服务器移动规范分类到新的父分类；ParentId 为空表示移动到根级。
        /// </summary>
        public async Task<StandardManagementOperationClientResponse> MoveManagementCategoryAsync(
            long categoryId,
            long? parentId,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (categoryId <= 0) throw new ArgumentException("分类 ID 必须大于 0。", nameof(categoryId));
            string requestJson = JsonConvert.SerializeObject(new StandardCategoryMoveClientRequest { ParentId = parentId });
            using var content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            using HttpRequestMessage message = new HttpRequestMessage(
                HttpMethod.Post,
                BuildServerUrl($"/api/standards/management/categories/{categoryId}/move"))
            {
                Content = content
            };
            AddOperatorHeader(message, operatorName);

            using (HttpResponseMessage response = await HttpClient.SendAsync(message, cancellationToken).ConfigureAwait(false))
            {
                string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                {
                    if ((int)response.StatusCode == 409)
                    {
                        StandardCategoryConflictClientResponse? conflict = JsonConvert.DeserializeObject<StandardCategoryConflictClientResponse>(body);
                        if (conflict != null)
                            throw new StandardCategoryConflictException(conflict.Message, conflict.Duplicates);
                    }

                    throw new HttpRequestException($"移动规范分类失败，HTTP {(int)response.StatusCode}，响应：{body}");
                }

                return JsonConvert.DeserializeObject<StandardManagementOperationClientResponse>(body)
                    ?? throw new InvalidOperationException("服务器返回的移动规范分类响应为空。");
            }
        }

        public async Task<StandardManagementOperationClientResponse> MoveManagementSeriesAsync(
            long seriesId,
            long categoryId,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (seriesId <= 0) throw new ArgumentException("规范系列 ID 必须大于 0。", nameof(seriesId));
            if (categoryId <= 0) throw new ArgumentException("目标分类 ID 必须大于 0。", nameof(categoryId));
            string requestJson = JsonConvert.SerializeObject(new StandardSeriesMoveClientRequest { CategoryId = categoryId });
            using var content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            using HttpRequestMessage message = new HttpRequestMessage(
                HttpMethod.Post,
                BuildServerUrl($"/api/standards/management/series/{seriesId}/move"))
            {
                Content = content
            };
            AddOperatorHeader(message, operatorName);
            return await SendManagementRequestAsync<StandardManagementOperationClientResponse>(
                message, "移动旧规范", cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// 通过服务器修改规范系列名称。
        /// </summary>
        public async Task<StandardManagementOperationClientResponse> RenameManagementSeriesAsync(
            long seriesId,
            string name,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (seriesId <= 0) throw new ArgumentException("规范系列 ID 必须大于 0。", nameof(seriesId));
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("规范名称不能为空。", nameof(name));

            string requestJson = JsonConvert.SerializeObject(new StandardSeriesRenameClientRequest
            {
                Name = name.Trim()
            });
            using var content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            using HttpRequestMessage message = new HttpRequestMessage(
                HttpMethod.Put,
                BuildServerUrl($"/api/standards/management/series/{seriesId}/name"))
            {
                Content = content
            };
            AddOperatorHeader(message, operatorName);
            return await SendManagementRequestAsync<StandardManagementOperationClientResponse>(
                message, "重命名规范系列", cancellationToken).ConfigureAwait(false);
        }

        public async Task<System.Collections.Generic.List<StandardDocumentFileClient>> GetManagementFilesAsync(
            long versionId,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            string requestUrl = BuildServerUrl($"/api/standards/management/versions/{versionId}/files");
            using (HttpResponseMessage response = await HttpClient.GetAsync(requestUrl, cancellationToken).ConfigureAwait(false))
            {
                string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new HttpRequestException($"规范附件查询失败，HTTP {(int)response.StatusCode}，响应：{body}");
                return JsonConvert.DeserializeObject<System.Collections.Generic.List<StandardDocumentFileClient>>(body)
                    ?? new System.Collections.Generic.List<StandardDocumentFileClient>();
            }
        }

        public async Task<System.Collections.Generic.List<StandardDocumentVersionClient>> GetManagementVersionsAsync(
            long seriesId,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            string requestUrl = BuildServerUrl($"/api/standards/management/series/{seriesId}/versions");
            using (HttpResponseMessage response = await HttpClient.GetAsync(requestUrl, cancellationToken).ConfigureAwait(false))
            {
                string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new HttpRequestException($"规范版本查询失败，HTTP {(int)response.StatusCode}，响应：{body}");
                return JsonConvert.DeserializeObject<System.Collections.Generic.List<StandardDocumentVersionClient>>(body)
                    ?? new System.Collections.Generic.List<StandardDocumentVersionClient>();
            }
        }

        public async Task<StandardManagementOperationClientResponse> CreateManagementVersionAsync(
            StandardVersionCreateClientRequest request,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            string requestUrl = BuildServerUrl("/api/standards/management/versions");
            string requestJson = JsonConvert.SerializeObject(request);
            using var content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            using HttpRequestMessage message = new HttpRequestMessage(HttpMethod.Post, requestUrl) { Content = content };
            AddOperatorHeader(message, operatorName);
            return await SendManagementRequestAsync<StandardManagementOperationClientResponse>(message, "创建规范版本", cancellationToken).ConfigureAwait(false);
        }

        public async Task<StandardFileUploadClientResponse> UploadManagementFileAsync(
            long versionId,
            string filePath,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            LogManager.Instance.LogInfo($"动态预览步骤 1/4：开始准备上传文件。文件={Path.GetFileName(filePath)}");
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("规范附件不存在。", filePath);

            string requestUrl = BuildServerUrl($"/api/standards/management/versions/{versionId}/files");
            using var form = new MultipartFormDataContent();
            using var stream = File.OpenRead(filePath);
            using var content = new StreamContent(stream);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            form.Add(content, "file", Path.GetFileName(filePath));
            LogManager.Instance.LogInfo($"动态预览步骤 2/4：已创建 multipart 请求。地址={requestUrl}，大小={stream.Length}");
            using HttpRequestMessage message = new HttpRequestMessage(HttpMethod.Post, requestUrl) { Content = form };
            AddOperatorHeader(message, operatorName);
            return await SendManagementRequestAsync<StandardFileUploadClientResponse>(message, "上传规范附件", cancellationToken).ConfigureAwait(false);
        }

        public async Task<StandardManagementOperationClientResponse> DeleteManagementVersionAsync(
            long versionId,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            using HttpRequestMessage message = new HttpRequestMessage(HttpMethod.Delete, BuildServerUrl($"/api/standards/management/versions/{versionId}"));
            AddOperatorHeader(message, operatorName);
            return await SendManagementRequestAsync<StandardManagementOperationClientResponse>(message, "删除规范版本", cancellationToken).ConfigureAwait(false);
        }

        public async Task<StandardManagementOperationClientResponse> RestoreManagementVersionAsync(
            long versionId,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            using HttpRequestMessage message = new HttpRequestMessage(HttpMethod.Post, BuildServerUrl($"/api/standards/management/versions/{versionId}/restore"));
            AddOperatorHeader(message, operatorName);
            return await SendManagementRequestAsync<StandardManagementOperationClientResponse>(message, "恢复规范版本", cancellationToken).ConfigureAwait(false);
        }

        public async Task DownloadManagementFileAsync(
            long fileId,
            string targetPath,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            using HttpResponseMessage response = await HttpClient.GetAsync(
                BuildServerUrl($"/api/standards/management/files/{fileId}/download"), cancellationToken).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                throw new HttpRequestException($"下载规范附件失败，HTTP {(int)response.StatusCode}，响应：{body}");
            }

            using Stream source = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
            using FileStream target = File.Create(targetPath);
            await source.CopyToAsync(target, 81920, cancellationToken).ConfigureAwait(false);
        }

        private static void AddOperatorHeader(HttpRequestMessage message, string operatorName)
        {
            message.Headers.TryAddWithoutValidation("X-Operator-Name", operatorName ?? string.Empty);
        }

        private static async Task<T> SendManagementRequestAsync<T>(
            HttpRequestMessage message,
            string operation,
            CancellationToken cancellationToken)
        {
            using HttpResponseMessage response = await HttpClient.SendAsync(message, cancellationToken).ConfigureAwait(false);
            string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            LogManager.Instance.LogInfo($"动态预览步骤 3/4：服务器已返回响应。HTTP={(int)response.StatusCode}，响应长度={body.Length}");
            if (!response.IsSuccessStatusCode)
                throw new HttpRequestException($"{operation}失败，HTTP {(int)response.StatusCode}，响应：{body}");
            return JsonConvert.DeserializeObject<T>(body) ?? throw new InvalidOperationException($"服务器返回的{operation}响应为空。");
        }

        /// <summary>
        /// 上传 Excel 或 JSON 规范文件，只执行服务器预览校验，不写入数据库。
        /// </summary>
        public async Task<StandardImportPreviewClientResponse> PreviewImportAsync(
            string filePath,
            string operatorName,
            long? categoryId = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("规范文件不存在。", filePath);

            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            string endpoint = extension == ".json"
                ? "/api/standards/import/preview-json"
                : "/api/standards/import/preview";
            string requestUrl = BuildServerUrl(endpoint);

            using var form = new MultipartFormDataContent();
            using var stream = File.OpenRead(filePath);
            using var content = new StreamContent(stream);
            content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(
                extension == ".json" ? "application/json" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            form.Add(content, "file", Path.GetFileName(filePath));

            // Excel 的系列元数据来自文件名，避免所有文件都被误导入到表52/PN10系列。
            if (extension != ".json")
            {
                StandardImportFileMetadata metadata = ParseStandardImportFileName(filePath);
                form.Add(new StringContent("FLANGE"), "familyCode");
                form.Add(new StringContent("法兰"), "familyName");
                form.Add(new StringContent("PLATE_WELD"), "seriesCode");
                form.Add(new StringContent(metadata.SeriesName), "seriesName");
                form.Add(new StringContent(metadata.StandardNumber), "standardNumber");
                form.Add(new StringContent(metadata.TableNumber), "tableNumber");
                form.Add(new StringContent(metadata.PressureRating), "pressureRating");
                form.Add(new StringContent("PL"), "flangeType");
                form.Add(new StringContent("RF"), "faceType");
                if (categoryId.HasValue)
                    form.Add(new StringContent(categoryId.Value.ToString()), "categoryId");
                LogManager.Instance.LogInfo($"规范 Excel 元数据解析：文件={Path.GetFileName(filePath)}，系列={metadata.SeriesName}，标准号={metadata.StandardNumber}，表号={metadata.TableNumber}，压力等级={metadata.PressureRating}");
            }

            using (HttpRequestMessage message = new HttpRequestMessage(HttpMethod.Post, requestUrl) { Content = form })
            {
                AddOperatorHeader(message, operatorName);
                using (HttpResponseMessage response = await HttpClient.SendAsync(message, cancellationToken).ConfigureAwait(false))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                        throw new HttpRequestException($"规范文件预览失败，HTTP {(int)response.StatusCode}，响应：{body}");

                    StandardImportPreviewClientResponse? result = JsonConvert.DeserializeObject<StandardImportPreviewClientResponse>(body);
                    return result ?? throw new InvalidOperationException("服务器返回的规范预览为空。");
                }
            }
        }

        /// <summary>
        /// 按服务器模板预览任意 Excel 表头，不写入规范数据。
        /// </summary>
        public async Task<DynamicStandardPreviewClientResponse> PreviewDynamicImportAsync(
            string filePath,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("规范文件不存在。", filePath);
            if (!string.Equals(Path.GetExtension(filePath), ".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("动态规范预览当前只支持 .xlsx 文件。");

            string requestUrl = BuildServerUrl("/api/standards/import/dynamic-preview");
            using var form = new MultipartFormDataContent();
            using var stream = File.OpenRead(filePath);
            using var content = new StreamContent(stream);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            form.Add(content, "file", Path.GetFileName(filePath));

            using HttpRequestMessage message = new HttpRequestMessage(HttpMethod.Post, requestUrl) { Content = form };
            AddOperatorHeader(message, operatorName);
            using HttpResponseMessage response = await HttpClient.SendAsync(message, cancellationToken).ConfigureAwait(false);
            string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                LogManager.Instance.LogError($"动态预览步骤 3/4：服务器返回失败。HTTP={(int)response.StatusCode}，响应={body}");
                throw new HttpRequestException($"动态规范预览失败，HTTP {(int)response.StatusCode}，响应：{body}");
            }

            DynamicStandardPreviewClientResponse result = JsonConvert.DeserializeObject<DynamicStandardPreviewClientResponse>(body)
                ?? throw new InvalidOperationException("服务器返回的动态规范预览为空。");
            LogManager.Instance.LogInfo($"动态预览步骤 4/4：响应解析完成。模板匹配={result.IsTemplateMatched}，行数={result.Rows?.Count ?? 0}，错误={result.ErrorCount}，警告={result.WarningCount}");
            return result;
        }

        /// <summary>
        /// 确认动态预览并保存为服务器导入批次，不直接发布到业务规范表。
        /// </summary>
        public async Task<DynamicStandardImportCommitClientResponse> CommitDynamicImportAsync(
            DynamicStandardImportCommitClientRequest request,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            LogManager.Instance.LogInfo($"动态确认步骤 1/3：开始校验提交请求。批次={request?.BatchId}，系列={request?.SeriesId}，行数={request?.Rows?.Count ?? 0}");
            if (request == null)
                throw new ArgumentNullException(nameof(request));
            if (string.IsNullOrWhiteSpace(request.BatchId))
                throw new ArgumentException("动态导入批次号不能为空。", nameof(request));
            // SeriesId 大于 0 表示使用已有基础规范；SeriesId 等于 0 表示由服务器在确认事务中创建基础规范。
            if (request.SeriesId < 0)
                throw new ArgumentException("目标规范系列 ID 不能小于 0。", nameof(request));
            // 新建基础规范时，客户端必须先提交完整身份信息，避免服务器进入事务后才发现字段缺失。
            if (request.SeriesId == 0)
            {
                if (string.IsNullOrWhiteSpace(request.SeriesName))
                    throw new ArgumentException("新建基础规范时，基础规范名称不能为空。", nameof(request));
                if (string.IsNullOrWhiteSpace(request.StandardNumber))
                    throw new ArgumentException("新建基础规范时，标准号不能为空。", nameof(request));
                if (string.IsNullOrWhiteSpace(request.SeriesCode))
                    throw new ArgumentException("新建基础规范时，系列编码不能为空。", nameof(request));
                if (string.IsNullOrWhiteSpace(request.FamilyCode))
                    throw new ArgumentException("新建基础规范时，专业编码不能为空。", nameof(request));
            }

            string requestJson = JsonConvert.SerializeObject(request);
            using var content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            using HttpRequestMessage message = new HttpRequestMessage(
                HttpMethod.Post,
                BuildServerUrl("/api/standards/import/dynamic-commit"))
            {
                Content = content
            };
            AddOperatorHeader(message, operatorName);
            LogManager.Instance.LogInfo("动态确认步骤 2/3：已创建 JSON 提交请求。");
            DynamicStandardImportCommitClientResponse result = await SendManagementRequestAsync<DynamicStandardImportCommitClientResponse>(
                message, "确认动态规范导入", cancellationToken).ConfigureAwait(false);
            LogManager.Instance.LogInfo($"动态确认步骤 3/3：服务器响应解析完成。成功={result.Success}，批次={result.BatchId}，保存行数={result.SavedRowCount}");
            return result;
        }

        private static StandardImportFileMetadata ParseStandardImportFileName(string filePath)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath).Trim();
            string[] parts = fileName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 4)
                throw new InvalidDataException("规范 Excel 文件名格式不正确，应为：规范名称_标准号_表号_压力等级.xlsx。");

            string seriesName = parts[0].Trim();
            string standardNumber = Regex.Replace(parts[1].Trim(), "^GB-T\\s*", "GB/T ", RegexOptions.IgnoreCase);
            string tableNumber = parts[2].Trim();
            string pressureRating = parts[3].Trim().ToUpperInvariant();
            if (!Regex.IsMatch(standardNumber, @"^GB/T\s+\S+$", RegexOptions.IgnoreCase)
                || !Regex.IsMatch(tableNumber, @"^表\s*\d+$", RegexOptions.IgnoreCase)
                || !Regex.IsMatch(pressureRating, @"^PN\s*\d+$", RegexOptions.IgnoreCase))
            {
                throw new InvalidDataException("规范 Excel 文件名中的标准号、表号或压力等级格式不正确，应为：GB-T xxxx_表51_PN6。");
            }

            return new StandardImportFileMetadata(seriesName, standardNumber, tableNumber, pressureRating.Replace(" ", string.Empty));
        }

        private sealed class StandardImportFileMetadata
        {
            public StandardImportFileMetadata(string seriesName, string standardNumber, string tableNumber, string pressureRating)
            {
                SeriesName = seriesName;
                StandardNumber = standardNumber;
                TableNumber = tableNumber;
                PressureRating = pressureRating;
            }

            public string SeriesName { get; }
            public string StandardNumber { get; }
            public string TableNumber { get; }
            public string PressureRating { get; }
        }

        public async Task<StandardImportCommitClientResponse> CommitImportAsync(
            StandardImportCommitClientRequest request,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));
            if (string.IsNullOrWhiteSpace(request.BatchId))
                throw new ArgumentException("导入批次号不能为空。", nameof(request));

            string requestJson = JsonConvert.SerializeObject(request);
            using var content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            using HttpRequestMessage message = new HttpRequestMessage(
                HttpMethod.Post,
                BuildServerUrl("/api/standards/import/commit"))
            {
                Content = content
            };
            AddOperatorHeader(message, operatorName);
            return await SendManagementRequestAsync<StandardImportCommitClientResponse>(
                message, "确认导入规范", cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// 查询规范系列下的全部实际法兰规范记录。
        /// </summary>
        public async Task<List<FlangeStandardRecordClient>> GetFlangeRecordsAsync(
            long seriesId,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (seriesId <= 0)
                throw new ArgumentException("规范系列 ID 必须大于 0。", nameof(seriesId));

            string requestUrl = BuildServerUrl($"/api/standards/flanges/series/{seriesId}/records");
            using (HttpResponseMessage response = await HttpClient.GetAsync(requestUrl, cancellationToken).ConfigureAwait(false))
            {
                string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new HttpRequestException($"规范内容查询失败，HTTP {(int)response.StatusCode}，响应：{responseBody}");

                return JsonConvert.DeserializeObject<List<FlangeStandardRecordClient>>(responseBody)
                    ?? new List<FlangeStandardRecordClient>();
            }
        }

        /// <summary>
        /// 查询管道通用字段目录和进口/出口图面样式。
        /// </summary>
        public async Task<PipelineFieldCatalogResponseClient> GetPipelineFieldCatalogAsync(
            CancellationToken cancellationToken = default(CancellationToken))
        {
            string requestUrl = BuildServerUrl("/api/pipelines/fields");
            try
            {
                using (HttpResponseMessage response = await HttpClient
                    .GetAsync(requestUrl, cancellationToken)
                    .ConfigureAwait(false))
                {
                    string responseBody = await response.Content
                        .ReadAsStringAsync()
                        .ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException(
                            $"管道字段目录查询失败，HTTP {(int)response.StatusCode}，地址：{requestUrl}，响应：{responseBody}");
                    }

                    PipelineFieldCatalogResponseClient? result = JsonConvert
                        .DeserializeObject<PipelineFieldCatalogResponseClient>(responseBody);
                    if (result == null)
                    {
                        throw new InvalidOperationException("管道字段目录查询返回为空。");
                    }

                    result.Fields ??= new List<PipelineFieldDefinitionClient>();
                    result.RoleStyles ??= new List<PipelineRoleStyleClient>();
                    LogManager.Instance.LogInfo(
                        $"管道字段目录查询完成：地址={requestUrl}，字段数={result.Fields.Count}，角色样式数={result.RoleStyles.Count}");
                    return result;
                }
            }
            catch (Exception exception)
            {
                LogManager.Instance.LogInfo(
                    $"管道字段目录查询异常：地址={requestUrl}，错误={exception.Message}");
                throw;
            }
        }

        /// <summary>
        /// 查询管道通用参数默认值。
        /// </summary>
        public async Task<PipelineDefaultsResponseClient> GetPipelineDefaultsAsync(
            CancellationToken cancellationToken = default(CancellationToken))
        {
            string requestUrl = BuildServerUrl("/api/pipelines/defaults");
            try
            {
                using (HttpResponseMessage response = await HttpClient
                    .GetAsync(requestUrl, cancellationToken)
                    .ConfigureAwait(false))
                {
                    string responseBody = await response.Content
                        .ReadAsStringAsync()
                        .ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException(
                            $"管道默认值查询失败，HTTP {(int)response.StatusCode}，地址：{requestUrl}，响应：{responseBody}");
                    }

                    PipelineDefaultsResponseClient? result = JsonConvert
                        .DeserializeObject<PipelineDefaultsResponseClient>(responseBody);
                    if (result == null)
                    {
                        throw new InvalidOperationException("管道默认值查询返回为空。");
                    }

                    result.Attributes ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    LogManager.Instance.LogInfo(
                        $"管道默认值查询完成：地址={requestUrl}，属性数={result.Attributes.Count}");
                    return result;
                }
            }
            catch (Exception exception)
            {
                LogManager.Instance.LogInfo(
                    $"管道默认值查询异常：地址={requestUrl}，错误={exception.Message}");
                throw;
            }
        }

        /// <summary>
        /// 请求服务器匹配管道 GB 设计规范。
        /// </summary>
        public async Task<PipelineDesignStandardMatchResponseClient> MatchPipelineDesignStandardAsync(
            PipelineDesignStandardMatchRequestClient request,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            string requestUrl = BuildServerUrl("/api/pipelines/design-standard/match");
            string requestJson = JsonConvert.SerializeObject(request);
            try
            {
                using (StringContent content = new StringContent(requestJson, Encoding.UTF8, "application/json"))
                using (HttpResponseMessage response = await HttpClient
                    .PostAsync(requestUrl, content, cancellationToken)
                    .ConfigureAwait(false))
                {
                    string responseBody = await response.Content
                        .ReadAsStringAsync()
                        .ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException(
                            $"管道 GB 设计规范查询失败，HTTP {(int)response.StatusCode}，地址：{requestUrl}，响应：{responseBody}");
                    }

                    PipelineDesignStandardMatchResponseClient? result = JsonConvert
                        .DeserializeObject<PipelineDesignStandardMatchResponseClient>(responseBody);
                    if (result == null)
                    {
                        throw new InvalidOperationException("管道 GB 设计规范查询返回为空。");
                    }

                    result.Attributes ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    LogManager.Instance.LogInfo(
                        $"管道 GB 设计规范查询完成：地址={requestUrl}，标准={request.DrawingStandardNo}，DN={request.DN}，PN={request.PN}，成功={result.Success}，匹配数={result.MatchCount}，消息={result.Message}");
                    return result;
                }
            }
            catch (Exception exception)
            {
                LogManager.Instance.LogInfo(
                    $"管道 GB 设计规范查询异常：地址={requestUrl}，标准={request.DrawingStandardNo}，DN={request.DN}，PN={request.PN}，错误={exception.Message}");
                throw;
            }
        }

        /// <summary>
        /// 查询法兰规范。
        /// </summary>
        /// <param name="request">法兰标准、DN、PN 和系列等条件。</param>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <returns>服务器返回的规范属性和完整记录。</returns>
        public async Task<FlangeStandardMatchResponse> MatchFlangeAsync(
            FlangeStandardMatchRequest request,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            // 对外统一保存 DN50 形式，避免客户端传入纯数字导致服务器无法识别。
            request.DN = NormalizeDn(request.DN);

            // 读取登录窗口保存的服务器地址；开发环境没有配置时回退到本机。
            string serverIp = string.IsNullOrWhiteSpace(VariableDictionary._serverIP)
                ? "127.0.0.1"
                : VariableDictionary._serverIP.Trim();

            // 与现有分类 API 保持相同端口读取规则。
            int apiPort = VariableDictionary._apiPort > 0
                ? VariableDictionary._apiPort
                : 10010;

            // 组合服务器规范查询接口地址。
            string requestUrl = $"http://{serverIp}:{apiPort}/api/standards/flanges/match";

            try
            {
                // 将查询条件序列化为服务器接口约定的 JSON。
                string requestJson = JsonConvert.SerializeObject(request);
                using (var content = new StringContent(requestJson, Encoding.UTF8, "application/json"))
                using (HttpResponseMessage response = await HttpClient
                    .PostAsync(requestUrl, content, cancellationToken)
                    .ConfigureAwait(false))
                {
                    // 先读取完整响应，便于服务器返回错误时保留具体提示。
                    string responseBody = await response.Content
                        .ReadAsStringAsync()
                        .ConfigureAwait(false);

                    // 非 2xx 响应直接抛出接口地址和服务器消息。
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException(
                            $"法兰规范查询失败，HTTP {(int)response.StatusCode}，地址：{requestUrl}，响应：{responseBody}");
                    }

                    // 反序列化服务器匹配结果。
                    FlangeStandardMatchResponse? result = JsonConvert
                        .DeserializeObject<FlangeStandardMatchResponse>(responseBody);
                    if (result == null)
                    {
                        throw new InvalidOperationException("法兰规范查询返回为空。");
                    }

                    // 兼容服务器返回 null 字典的情况。
                    result.Attributes ??= new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    // 记录查询结果，便于后续定位规范匹配问题。
                    LogManager.Instance.LogInfo(
                        $"法兰规范查询完成：地址={requestUrl}，DN={request.DN}，PN={request.PN}，成功={result.Success}，消息={result.Message}");

                    return result;
                }
            }
            catch (Exception ex)
            {
                // 只记录接口地址和业务条件，不记录数据库连接信息。
                LogManager.Instance.LogInfo(
                    $"法兰规范查询异常：地址={requestUrl}，DN={request.DN}，PN={request.PN}，错误={ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 第一阶段验证入口：查询 GB/T 9124.1-2019 表52 PN10 的 DN50 法兰数据。
        /// </summary>
        public Task<FlangeStandardMatchResponse> QueryDn50Async(
            CancellationToken cancellationToken = default(CancellationToken))
        {
            // 使用模型中的默认标准条件，仅补充本阶段验证所需的 DN50。
            var request = new FlangeStandardMatchRequest
            {
                StandardNumber = "GB/T 9124.1-2019",
                TableNumber = "表52",
                DN = "DN50"
            };

            // 复用正式查询方法，确保验证入口与后续业务使用完全相同的请求链路。
            return MatchFlangeAsync(request, cancellationToken);
        }

        /// <summary>
        /// 将 DN50 或 50 统一转换为 DN50。
        /// </summary>
        private static string NormalizeDn(string value)
        {
            string normalized = (value ?? string.Empty).Trim().ToUpperInvariant().Replace(" ", string.Empty);
            if (normalized.StartsWith("DN", StringComparison.Ordinal))
            {
                return normalized;
            }

            return int.TryParse(normalized, out int dnValue)
                ? $"DN{dnValue}"
                : normalized;
        }

        /// <summary>
        /// 按登录窗口的服务器配置构造 API 地址。
        /// </summary>
        private static string BuildServerUrl(string path)
        {
            string serverIp = string.IsNullOrWhiteSpace(VariableDictionary._serverIP)
                ? "127.0.0.1"
                : VariableDictionary._serverIP.Trim();
            int apiPort = VariableDictionary._apiPort > 0 ? VariableDictionary._apiPort : 10010;
            return $"http://{serverIp}:{apiPort}{path}";
        }
    }
}
