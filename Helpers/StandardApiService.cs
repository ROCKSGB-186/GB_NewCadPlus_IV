using GB_NewCadPlus_IV.Models;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using Newtonsoft.Json;
using System;
using System.Net.Http;
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
    }
}
