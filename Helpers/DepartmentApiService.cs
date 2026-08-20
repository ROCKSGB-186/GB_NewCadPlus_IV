using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.Helpers
{
    public sealed class DepartmentApiService
    {
        private static readonly HttpClient HttpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(30)
        };

        public async Task<List<DepartmentModel>> GetDepartmentsWithCountsAsync(CancellationToken cancellationToken = default)
        {
            string requestUrl = ApiEndpoint.Build("api/departments");
            Stopwatch stopwatch = Stopwatch.StartNew();
            LogManager.Instance.LogInfo($"部门接口请求开始：GET {requestUrl}");

            try
            {
                using (HttpResponseMessage response = await HttpClient.GetAsync(requestUrl, cancellationToken).ConfigureAwait(false))
                {
                    string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    string contentType = response.Content.Headers.ContentType?.MediaType ?? "未声明";
                    LogManager.Instance.LogInfo(
                        $"部门接口响应：HTTP {(int)response.StatusCode} {response.ReasonPhrase}, ContentType={contentType}, BodyLength={responseBody.Length}, ElapsedMs={stopwatch.ElapsedMilliseconds}");

                    if (!response.IsSuccessStatusCode)
                    {
                        string errorMessage = TryReadErrorMessage(responseBody);
                        LogManager.Instance.LogWarning(
                            $"部门接口返回失败：HTTP {(int)response.StatusCode}, SafeMessage={errorMessage}, ElapsedMs={stopwatch.ElapsedMilliseconds}");
                        throw new HttpRequestException($"部门接口请求失败，HTTP {(int)response.StatusCode}：{errorMessage}");
                    }

                    DepartmentListApiResponse? result = JsonConvert.DeserializeObject<DepartmentListApiResponse>(responseBody);
                    if (result == null || !result.Success)
                        throw new InvalidOperationException(result?.Message ?? "服务器部门查询失败。");

                    result.Departments ??= new List<DepartmentModel>();
                    LogManager.Instance.LogInfo($"通过服务器加载部门成功：地址={requestUrl}，部门数={result.Departments.Count}，ElapsedMs={stopwatch.ElapsedMilliseconds}");
                    return result.Departments;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError(
                    $"部门接口请求异常：URL={requestUrl}, ExceptionType={ex.GetType().FullName}, Message={ex.Message}, ElapsedMs={stopwatch.ElapsedMilliseconds}");
                throw;
            }
            finally
            {
                stopwatch.Stop();
            }
        }

        private static string TryReadErrorMessage(string responseBody)
        {
            if (string.IsNullOrWhiteSpace(responseBody))
                return "服务器未返回错误详情。";

            try
            {
                DepartmentErrorResponse? error = JsonConvert.DeserializeObject<DepartmentErrorResponse>(responseBody);
                if (!string.IsNullOrWhiteSpace(error?.Message))
                    return error.Message.Trim();
            }
            catch (JsonException)
            {
                // 非 JSON 响应（例如旧版本服务器返回的 HTML）不继续输出原始内容。
            }

            return responseBody.TrimStart().StartsWith("<", StringComparison.Ordinal)
                ? "服务器返回了无效的错误响应。"
                : "服务器返回了错误响应。";
        }

        private sealed class DepartmentListApiResponse
        {
            public bool Success { get; set; }
            public string Message { get; set; } = string.Empty;
            public List<DepartmentModel> Departments { get; set; } = new List<DepartmentModel>();
        }

        private sealed class DepartmentErrorResponse
        {
            public string Message { get; set; } = string.Empty;
        }
    }
}
