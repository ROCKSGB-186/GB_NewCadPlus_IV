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
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("规范附件不存在。", filePath);

            string requestUrl = BuildServerUrl($"/api/standards/management/versions/{versionId}/files");
            using var form = new MultipartFormDataContent();
            using var stream = File.OpenRead(filePath);
            using var content = new StreamContent(stream);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            form.Add(content, "file", Path.GetFileName(filePath));
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

            // Excel 预览接口需要系列元数据；第一阶段先提供法兰默认值，后续由新增规范库窗口填写。
            if (extension != ".json")
            {
                form.Add(new StringContent("FLANGE"), "familyCode");
                form.Add(new StringContent("法兰"), "familyName");
                form.Add(new StringContent("PLATE_WELD"), "seriesCode");
                form.Add(new StringContent("板式平焊钢制管法兰"), "seriesName");
                form.Add(new StringContent("GB/T 9124.1-2019"), "standardNumber");
                form.Add(new StringContent("表52"), "tableNumber");
                form.Add(new StringContent("PN10"), "pressureRating");
                form.Add(new StringContent("PL"), "flangeType");
                form.Add(new StringContent("RF"), "faceType");
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

        public async Task<StandardImportCommitClientResponse> CommitImportAsync(
            string batchId,
            bool allowWarnings,
            string operatorName,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (string.IsNullOrWhiteSpace(batchId))
                throw new ArgumentException("导入批次号不能为空。", nameof(batchId));
            using HttpRequestMessage message = new HttpRequestMessage(
                HttpMethod.Post,
                BuildServerUrl($"/api/standards/import/commit?batchId={Uri.EscapeDataString(batchId.Trim())}&allowWarnings={allowWarnings.ToString().ToLowerInvariant()}"));
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
