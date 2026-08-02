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
    public sealed class DepartmentApiService
    {
        private static readonly HttpClient HttpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(30)
        };

        public async Task<List<DepartmentModel>> GetDepartmentsWithCountsAsync(CancellationToken cancellationToken = default)
        {
            string serverIp = string.IsNullOrWhiteSpace(VariableDictionary._serverIP)
                ? "127.0.0.1"
                : VariableDictionary._serverIP.Trim();
            int apiPort = VariableDictionary._apiPort > 0 ? VariableDictionary._apiPort : 10010;
            string requestUrl = $"http://{serverIp}:{apiPort}/api/departments";

            using (HttpResponseMessage response = await HttpClient.GetAsync(requestUrl, cancellationToken).ConfigureAwait(false))
            {
                string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new HttpRequestException($"部门接口请求失败，HTTP {(int)response.StatusCode}，响应：{responseBody}");

                DepartmentListApiResponse? result = JsonConvert.DeserializeObject<DepartmentListApiResponse>(responseBody);
                if (result == null || !result.Success)
                    throw new InvalidOperationException(result?.Message ?? "服务器部门查询失败。");

                result.Departments ??= new List<DepartmentModel>();
                LogManager.Instance.LogInfo($"通过服务器加载部门成功：地址={requestUrl}，部门数={result.Departments.Count}，部门={string.Join("；", result.Departments.ConvertAll(d => $"{d.Id}:{d.Name},用户={d.UserCount}"))}");
                return result.Departments;
            }
        }

        private sealed class DepartmentListApiResponse
        {
            public bool Success { get; set; }
            public string Message { get; set; } = string.Empty;
            public List<DepartmentModel> Departments { get; set; } = new List<DepartmentModel>();
        }
    }
}
