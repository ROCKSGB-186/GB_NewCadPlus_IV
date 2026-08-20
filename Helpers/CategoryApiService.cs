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
    /// <summary>
    /// 客户端分类接口服务。
    /// 这个类只负责通过 HTTP 调用服务器，不直接访问 MySQL 或达梦数据库。
    /// </summary>
    public sealed class CategoryApiService
    {
        /// <summary>
        /// 全局复用 HttpClient，避免频繁创建连接造成端口耗尽。
        /// </summary>
        private static readonly HttpClient HttpClient = new HttpClient
        {
            // 分类数据量通常不大，30 秒足够完成一次查询。
            Timeout = TimeSpan.FromSeconds(30)
        };

        /// <summary>
        /// 从服务器获取主分类和子分类。
        /// </summary>
        /// <param name="cancellationToken">取消请求的令牌。</param>
        /// <returns>服务器返回的分类数据。</returns>
        public async Task<CategoryTreeApiResponse> GetCategoryTreeAsync(CancellationToken cancellationToken = default)
        {
            // 读取登录窗口保存的服务器 IP；没有配置时使用本机地址作为开发环境兜底。
            string requestUrl = ApiEndpoint.Build("api/categories/tree");

            try
            {
                // 向服务器发送 GET 请求；此处不携带数据库账号密码。
                using (HttpResponseMessage response = await HttpClient
                    .GetAsync(requestUrl, cancellationToken)
                    .ConfigureAwait(false))
                {
                    // 先读取服务器返回的 JSON 文本，便于错误时记录服务器具体消息。
                    string responseBody = await response.Content
                        .ReadAsStringAsync()
                        .ConfigureAwait(false);

                    // HTTP 状态码不是 2xx 时，直接抛出包含地址和响应内容的异常。
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException(
                            $"分类接口请求失败，HTTP {(int)response.StatusCode}，地址：{requestUrl}，响应：{responseBody}");
                    }

                    // 将服务器 JSON 转换为客户端 DTO。
                    CategoryTreeApiResponse? result = JsonConvert
                        .DeserializeObject<CategoryTreeApiResponse>(responseBody);

                    // 防止服务器返回空 JSON 导致后续树构建出现空引用。
                    if (result == null)
                    {
                        throw new InvalidOperationException("分类接口返回为空，无法构建分类树。");
                    }

                    // 服务端 success=false 时，把服务端提示传递给上层。
                    if (!result.Success)
                    {
                        throw new InvalidOperationException(
                            string.IsNullOrWhiteSpace(result.Message)
                                ? "服务器分类查询失败。"
                                : result.Message);
                    }

                    // 确保两个列表不为 null，兼容历史服务器返回 null 的情况。
                    result.Categories ??= new List<CategoryApiDto>();
                    result.Subcategories ??= new List<SubcategoryApiDto>();

                    // 记录成功请求，明确说明本次分类数据来自服务器接口，而不是客户端数据库。
                    LogManager.Instance.LogInfo(
                        $"通过服务器加载分类树成功：地址={requestUrl}，主分类={result.Categories.Count}，子分类={result.Subcategories.Count}");

                    // 返回已经解析完成的分类数据。
                    return result;
                }
            }
            catch (Exception ex)
            {
                // 记录服务器地址和异常，便于新手根据日志判断是网络、HTTP 还是 JSON 问题。
                LogManager.Instance.LogInfo(
                    $"通过服务器加载分类树失败：地址={requestUrl}，错误={ex.Message}");

                // 继续抛出异常，让界面层显示明确的失败提示。
                throw;
            }
        }

        /// <summary>
        /// 通过服务器删除没有子分类的主分类。
        /// </summary>
        /// <param name="categoryId">要删除的主分类 ID。</param>
        /// <param name="cancellationToken">取消请求的令牌。</param>
        /// <returns>服务器删除结果。</returns>
        public async Task<bool> DeleteCategoryAsync(
            int categoryId,
            CancellationToken cancellationToken = default)
        {
            // 主分类 ID 必须小于 10000，防止误调用子分类删除接口。
            if (categoryId <= 0 || categoryId >= 10000)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(categoryId),
                    "主分类ID必须大于0且小于10000。");
            }

            string requestUrl = ApiEndpoint.Build($"api/categories/{categoryId}");

            try
            {
                // 发送 DELETE 请求，客户端不直接访问分类数据库。
                using (HttpResponseMessage response = await HttpClient
                    .DeleteAsync(requestUrl, cancellationToken)
                    .ConfigureAwait(false))
                {
                    // 读取响应内容，便于记录服务器业务错误。
                    string responseBody = await response.Content
                        .ReadAsStringAsync()
                        .ConfigureAwait(false);

                    // 非成功状态码直接抛出异常。
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException(
                            $"删除主分类请求失败，HTTP {(int)response.StatusCode}，地址：{requestUrl}，响应：{responseBody}");
                    }

                    // 解析服务器统一删除响应。
                    CategoryDeleteApiResponse? result = JsonConvert
                        .DeserializeObject<CategoryDeleteApiResponse>(responseBody);
                    if (result == null)
                    {
                        throw new InvalidOperationException("删除主分类接口返回为空。");
                    }

                    // 服务器业务失败时向上层传递服务器提示。
                    if (!result.Success)
                    {
                        throw new InvalidOperationException(
                            string.IsNullOrWhiteSpace(result.Message)
                                ? "服务器删除主分类失败。"
                                : result.Message);
                    }

                    // 记录服务器删除成功日志。
                    LogManager.Instance.LogInfo(
                        $"通过服务器删除主分类成功：地址={requestUrl}，Id={result.DeletedId}");
                    return result.DeletedId == categoryId;
                }
            }
            catch (Exception ex)
            {
                // 记录接口地址和错误信息，不记录数据库密码。
                LogManager.Instance.LogInfo(
                    $"通过服务器删除主分类失败：地址={requestUrl}，错误={ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 通过服务器新增主分类。
        /// </summary>
        /// <param name="name">主分类名称。</param>
        /// <param name="displayName">主分类显示名称。</param>
        /// <param name="sortOrder">排序号；小于等于 0 时由服务器自动生成。</param>
        /// <param name="cancellationToken">取消请求的令牌。</param>
        /// <returns>服务器新增后的主分类。</returns>
        public async Task<CategoryApiDto> AddCategoryAsync(
            string name,
            string displayName,
            int sortOrder,
            CancellationToken cancellationToken = default)
        {
            // 读取登录窗口保存的服务器 IP。
            string requestUrl = ApiEndpoint.Build("api/categories");

            // 使用服务器 DTO 发送 JSON，客户端不再直接访问数据库。
            var request = new AddCategoryApiRequest
            {
                Name = name?.Trim() ?? string.Empty,
                DisplayName = string.IsNullOrWhiteSpace(displayName)
                    ? null
                    : displayName.Trim(),
                SortOrder = sortOrder > 0 ? sortOrder : (int?)null
            };

            try
            {
                // 将请求对象序列化为服务器接口需要的 JSON。
                string requestJson = JsonConvert.SerializeObject(request);
                using var content = new StringContent(
                    requestJson,
                    System.Text.Encoding.UTF8,
                    "application/json");

                // 向服务器发送 POST 请求。
                using (HttpResponseMessage response = await HttpClient
                    .PostAsync(requestUrl, content, cancellationToken)
                    .ConfigureAwait(false))
                {
                    // 读取完整响应，便于记录服务器错误信息。
                    string responseBody = await response.Content
                        .ReadAsStringAsync()
                        .ConfigureAwait(false);

                    // 非成功状态码直接抛出包含响应内容的异常。
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException(
                            $"新增主分类请求失败，HTTP {(int)response.StatusCode}，地址：{requestUrl}，响应：{responseBody}");
                    }

                    // 解析服务器统一响应对象。
                    CategoryMutationApiResponse? result = JsonConvert
                        .DeserializeObject<CategoryMutationApiResponse>(responseBody);

                    // 防止服务器返回空响应。
                    if (result == null)
                    {
                        throw new InvalidOperationException("新增主分类接口返回为空。");
                    }

                    // 服务器业务失败时向上层传递服务器提示。
                    if (!result.Success || result.Category == null)
                    {
                        throw new InvalidOperationException(
                            string.IsNullOrWhiteSpace(result.Message)
                                ? "服务器新增主分类失败。"
                                : result.Message);
                    }

                    // 记录明确日志，证明新增操作由服务器完成。
                    LogManager.Instance.LogInfo(
                        $"通过服务器新增主分类成功：地址={requestUrl}，Id={result.Category.Id}，名称={result.Category.Name}");

                    // 返回服务器生成的完整分类对象。
                    return result.Category;
                }
            }
            catch (Exception ex)
            {
                // 记录请求地址和错误，但不记录数据库账号密码。
                LogManager.Instance.LogInfo(
                    $"通过服务器新增主分类失败：地址={requestUrl}，错误={ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 通过服务器新增子分类。
        /// </summary>
        /// <param name="parentId">父分类 ID。</param>
        /// <param name="name">子分类名称。</param>
        /// <param name="displayName">子分类显示名称。</param>
        /// <param name="sortOrder">排序号；小于等于 0 时由服务器自动生成。</param>
        /// <param name="cancellationToken">取消请求的令牌。</param>
        /// <returns>服务器新增后的子分类。</returns>
        public async Task<SubcategoryApiDto> AddSubcategoryAsync(
            int parentId,
            string name,
            string displayName,
            int sortOrder,
            CancellationToken cancellationToken = default)
        {
            // 校验父级 ID，避免组合出无效接口地址。
            if (parentId <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(parentId), "父分类ID必须大于0。");
            }

            // 读取服务器地址配置。
            string requestUrl = ApiEndpoint.Build($"api/categories/{parentId}/subcategories");

            // 使用服务器 DTO 发送请求，客户端不再生成 ID、层级或更新父级列表。
            var request = new AddSubcategoryApiRequest
            {
                Name = name?.Trim() ?? string.Empty,
                DisplayName = string.IsNullOrWhiteSpace(displayName)
                    ? null
                    : displayName.Trim(),
                SortOrder = sortOrder > 0 ? sortOrder : (int?)null
            };

            try
            {
                // 将请求对象序列化为 JSON。
                string requestJson = JsonConvert.SerializeObject(request);
                using var content = new StringContent(
                    requestJson,
                    System.Text.Encoding.UTF8,
                    "application/json");

                // 向服务器发送新增子分类请求。
                using (HttpResponseMessage response = await HttpClient
                    .PostAsync(requestUrl, content, cancellationToken)
                    .ConfigureAwait(false))
                {
                    // 读取服务器响应文本，便于错误日志定位。
                    string responseBody = await response.Content
                        .ReadAsStringAsync()
                        .ConfigureAwait(false);

                    // 非 2xx 状态码直接抛出异常。
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException(
                            $"新增子分类请求失败，HTTP {(int)response.StatusCode}，地址：{requestUrl}，响应：{responseBody}");
                    }

                    // 解析服务器统一响应。
                    SubcategoryMutationApiResponse? result = JsonConvert
                        .DeserializeObject<SubcategoryMutationApiResponse>(responseBody);

                    // 防止服务器返回空响应。
                    if (result == null)
                    {
                        throw new InvalidOperationException("新增子分类接口返回为空。");
                    }

                    // 服务器业务失败或没有返回子分类时视为失败。
                    if (!result.Success || result.Subcategory == null)
                    {
                        throw new InvalidOperationException(
                            string.IsNullOrWhiteSpace(result.Message)
                                ? "服务器新增子分类失败。"
                                : result.Message);
                    }

                    // 记录服务器写入成功日志，便于和旧的本地数据库日志区分。
                    LogManager.Instance.LogInfo(
                        $"通过服务器新增子分类成功：地址={requestUrl}，Id={result.Subcategory.Id}，名称={result.Subcategory.Name}");

                    // 返回服务器生成的 ID、排序号和层级。
                    return result.Subcategory;
                }
            }
            catch (Exception ex)
            {
                // 记录接口地址和错误信息，不记录数据库密码。
                LogManager.Instance.LogInfo(
                    $"通过服务器新增子分类失败：地址={requestUrl}，错误={ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 通过服务器删除子分类。
        /// </summary>
        /// <param name="subcategoryId">要删除的子分类 ID。</param>
        /// <param name="cancellationToken">取消请求的令牌。</param>
        /// <returns>服务器删除结果。</returns>
        public async Task<bool> DeleteSubcategoryAsync(
            int subcategoryId,
            CancellationToken cancellationToken = default)
        {
            // 校验子分类 ID，防止误调用主分类删除路径。
            if (subcategoryId < 10000)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(subcategoryId),
                    "子分类ID必须大于等于10000。");
            }

            // 读取服务器地址配置。
            string requestUrl = ApiEndpoint.Build($"api/categories/subcategories/{subcategoryId}");

            try
            {
                // 发送 DELETE 请求；客户端不直接操作数据库。
                using (HttpResponseMessage response = await HttpClient
                    .DeleteAsync(requestUrl, cancellationToken)
                    .ConfigureAwait(false))
                {
                    // 读取响应内容，便于记录服务器错误信息。
                    string responseBody = await response.Content
                        .ReadAsStringAsync()
                        .ConfigureAwait(false);

                    // 非成功状态码直接抛出异常。
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException(
                            $"删除子分类请求失败，HTTP {(int)response.StatusCode}，地址：{requestUrl}，响应：{responseBody}");
                    }

                    // 解析服务器统一删除响应。
                    CategoryDeleteApiResponse? result = JsonConvert
                        .DeserializeObject<CategoryDeleteApiResponse>(responseBody);

                    // 防止服务器返回空响应。
                    if (result == null)
                    {
                        throw new InvalidOperationException("删除子分类接口返回为空。");
                    }

                    // 服务器业务失败时向上层传递服务器提示。
                    if (!result.Success)
                    {
                        throw new InvalidOperationException(
                            string.IsNullOrWhiteSpace(result.Message)
                                ? "服务器删除子分类失败。"
                                : result.Message);
                    }

                    // 记录服务器删除成功日志。
                    LogManager.Instance.LogInfo(
                        $"通过服务器删除子分类成功：地址={requestUrl}，Id={result.DeletedId}");

                    // 返回服务器确认的删除结果。
                    return result.DeletedId == subcategoryId;
                }
            }
            catch (Exception ex)
            {
                // 记录接口地址和错误信息，不记录数据库密码。
                LogManager.Instance.LogInfo(
                    $"通过服务器删除子分类失败：地址={requestUrl}，错误={ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 通过服务器更新主分类或子分类。
        /// </summary>
        /// <param name="categoryId">要更新的分类 ID。</param>
        /// <param name="name">分类名称。</param>
        /// <param name="displayName">分类显示名称。</param>
        /// <param name="sortOrder">分类排序号。</param>
        /// <param name="cancellationToken">取消请求的令牌。</param>
        /// <returns>服务器返回的更新结果。</returns>
        public async Task<CategoryUpdateApiResponse> UpdateCategoryAsync(
            int categoryId,
            string name,
            string displayName,
            int sortOrder,
            CancellationToken cancellationToken = default)
        {
            // 校验分类 ID，避免向服务器发送无效请求。
            if (categoryId <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(categoryId), "分类ID必须大于0。");
            }

            // 组合分类更新接口地址。
            string requestUrl = ApiEndpoint.Build($"api/categories/{categoryId}");

            // 按服务端 DTO 结构构造请求，客户端不直接拼接 SQL。
            var request = new CategoryUpdateApiRequest
            {
                Name = name?.Trim() ?? string.Empty,
                DisplayName = string.IsNullOrWhiteSpace(displayName)
                    ? null
                    : displayName.Trim(),
                SortOrder = sortOrder > 0 ? sortOrder : (int?)null
            };

            try
            {
                // 序列化更新请求。
                string requestJson = JsonConvert.SerializeObject(request);
                using var content = new StringContent(
                    requestJson,
                    System.Text.Encoding.UTF8,
                    "application/json");

                // 发送 PUT 请求，由服务器执行数据库事务。
                using (HttpResponseMessage response = await HttpClient
                    .PutAsync(requestUrl, content, cancellationToken)
                    .ConfigureAwait(false))
                {
                    // 读取服务器响应，便于记录具体失败原因。
                    string responseBody = await response.Content
                        .ReadAsStringAsync()
                        .ConfigureAwait(false);

                    // 非成功状态码直接抛出异常。
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException(
                            $"更新分类请求失败，HTTP {(int)response.StatusCode}，地址：{requestUrl}，响应：{responseBody}");
                    }

                    // 解析服务器统一响应。
                    CategoryUpdateApiResponse? result = JsonConvert
                        .DeserializeObject<CategoryUpdateApiResponse>(responseBody);
                    if (result == null)
                    {
                        throw new InvalidOperationException("更新分类接口返回为空。");
                    }

                    // 服务器业务失败时保留服务器提示。
                    if (!result.Success)
                    {
                        throw new InvalidOperationException(
                            string.IsNullOrWhiteSpace(result.Message)
                                ? "服务器更新分类失败。"
                                : result.Message);
                    }

                    // 记录服务器更新成功日志。
                    LogManager.Instance.LogInfo(
                        $"通过服务器更新分类成功：地址={requestUrl}，Id={result.UpdatedId}，名称={result.Name}");
                    return result;
                }
            }
            catch (Exception ex)
            {
                // 记录接口错误，不记录数据库连接信息。
                LogManager.Instance.LogInfo(
                    $"通过服务器更新分类失败：地址={requestUrl}，错误={ex.Message}");
                throw;
            }
        }
    }

    /// <summary>
    /// 新增主分类请求 DTO。
    /// </summary>
    public sealed class AddCategoryApiRequest
    {
        /// <summary>
        /// 分类名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 分类显示名称。
        /// </summary>
        public string? DisplayName { get; set; }

        /// <summary>
        /// 分类排序号。
        /// </summary>
        public int? SortOrder { get; set; }
    }

    /// <summary>
    /// 新增主分类响应 DTO。
    /// </summary>
    public sealed class CategoryMutationApiResponse
    {
        /// <summary>
        /// 操作是否成功。
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 服务器提示信息。
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 新增后的主分类。
        /// </summary>
        public CategoryApiDto? Category { get; set; }
    }

    /// <summary>
    /// 新增子分类请求 DTO。
    /// </summary>
    public sealed class AddSubcategoryApiRequest
    {
        /// <summary>
        /// 子分类名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 子分类显示名称。
        /// </summary>
        public string? DisplayName { get; set; }

        /// <summary>
        /// 子分类排序号。
        /// </summary>
        public int? SortOrder { get; set; }
    }

    /// <summary>
    /// 新增子分类响应 DTO。
    /// </summary>
    public sealed class SubcategoryMutationApiResponse
    {
        /// <summary>
        /// 操作是否成功。
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 服务器提示信息。
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 新增后的子分类。
        /// </summary>
        public SubcategoryApiDto? Subcategory { get; set; }
    }

    /// <summary>
    /// 删除子分类响应 DTO。
    /// </summary>
    public sealed class CategoryDeleteApiResponse
    {
        /// <summary>
        /// 操作是否成功。
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 服务器提示信息。
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 实际删除的分类 ID。
        /// </summary>
        public int DeletedId { get; set; }
    }

    /// <summary>
    /// 更新分类请求 DTO。
    /// </summary>
    public sealed class CategoryUpdateApiRequest
    {
        /// <summary>
        /// 分类名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 分类显示名称。
        /// </summary>
        public string? DisplayName { get; set; }

        /// <summary>
        /// 分类排序号。
        /// </summary>
        public int? SortOrder { get; set; }
    }

    /// <summary>
    /// 更新分类响应 DTO。
    /// </summary>
    public sealed class CategoryUpdateApiResponse
    {
        /// <summary>
        /// 操作是否成功。
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 服务器提示信息。
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 更新后的分类 ID。
        /// </summary>
        public int UpdatedId { get; set; }

        /// <summary>
        /// 更新后的分类名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 更新后的显示名称。
        /// </summary>
        public string DisplayName { get; set; } = string.Empty;

        /// <summary>
        /// 更新后的排序号。
        /// </summary>
        public int SortOrder { get; set; }
    }

    /// <summary>
    /// 服务器分类树接口的返回模型。
    /// </summary>
    public sealed class CategoryTreeApiResponse
    {
        /// <summary>
        /// 服务端是否处理成功。
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 服务端提示信息。
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 主分类列表。
        /// </summary>
        public List<CategoryApiDto> Categories { get; set; } = new List<CategoryApiDto>();

        /// <summary>
        /// 子分类列表。
        /// </summary>
        public List<SubcategoryApiDto> Subcategories { get; set; } = new List<SubcategoryApiDto>();
    }

    /// <summary>
    /// 服务器返回的主分类 DTO。
    /// </summary>
    public sealed class CategoryApiDto
    {
        /// <summary>
        /// 主分类 ID。
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// 分类内部名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 分类显示名称。
        /// </summary>
        public string DisplayName { get; set; } = string.Empty;

        /// <summary>
        /// 子分类 ID 逗号分隔列表。
        /// </summary>
        public string SubcategoryIds { get; set; } = string.Empty;

        /// <summary>
        /// 排序号。
        /// </summary>
        public int SortOrder { get; set; }

        /// <summary>
        /// 创建时间。
        /// </summary>
        public DateTime? CreatedAt { get; set; }

        /// <summary>
        /// 更新时间。
        /// </summary>
        public DateTime? UpdatedAt { get; set; }
    }

    /// <summary>
    /// 服务器返回的子分类 DTO。
    /// </summary>
    public sealed class SubcategoryApiDto
    {
        /// <summary>
        /// 子分类 ID。
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// 父级 ID。
        /// </summary>
        public int ParentId { get; set; }

        /// <summary>
        /// 子分类内部名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 子分类显示名称。
        /// </summary>
        public string DisplayName { get; set; } = string.Empty;

        /// <summary>
        /// 排序号。
        /// </summary>
        public int SortOrder { get; set; }

        /// <summary>
        /// 分类层级。
        /// </summary>
        public int Level { get; set; }

        /// <summary>
        /// 下级子分类 ID 逗号分隔列表。
        /// </summary>
        public string SubcategoryIds { get; set; } = string.Empty;

        /// <summary>
        /// 创建时间。
        /// </summary>
        public DateTime? CreatedAt { get; set; }

        /// <summary>
        /// 更新时间。
        /// </summary>
        public DateTime? UpdatedAt { get; set; }
    }
}
