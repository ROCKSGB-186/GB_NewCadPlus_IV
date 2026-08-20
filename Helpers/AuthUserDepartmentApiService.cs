using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.Helpers
{
    public sealed class AuthUserDepartmentApiService
    {
        private static readonly HttpClient HttpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };

        private sealed class LoginRequest { public string Username { get; set; } = ""; public string Password { get; set; } = ""; }
        // 修改密码请求：服务器使用账号、手机号和邮箱进行身份核验。
        private sealed class ResetPasswordRequest { public string Username { get; set; } = ""; public string Phone { get; set; } = ""; public string Email { get; set; } = ""; public string NewPassword { get; set; } = ""; }
        private sealed class LoginResponse { public bool Success { get; set; } public string Message { get; set; } = ""; public UserModel User { get; set; } }
        private sealed class RegisterRequest { public string Username { get; set; } = ""; public string Password { get; set; } = ""; public int DepartmentId { get; set; } public string DepartmentName { get; set; } public string RealName { get; set; } public string Gender { get; set; } public string Phone { get; set; } public string Email { get; set; } public string Role { get; set; } }
        private sealed class UserRequest { public string Username { get; set; } = ""; public string Password { get; set; } public int? DepartmentId { get; set; } public string DepartmentName { get; set; } public string RealName { get; set; } public string Gender { get; set; } public string Phone { get; set; } public string Email { get; set; } public string Role { get; set; } public bool IsActive { get; set; } }
        private sealed class DepartmentRequest { public string Name { get; set; } = ""; public string DisplayName { get; set; } public string Description { get; set; } public int SortOrder { get; set; } public int? ManagerUserId { get; set; } public bool IsActive { get; set; } = true; }
        public sealed class MutationResponse { public bool Success { get; set; } public string Message { get; set; } = ""; public int Id { get; set; } public List<DepartmentModel> Departments { get; set; } = new List<DepartmentModel>(); }

        public async Task<bool> AuthenticateAsync(string username, string password, CancellationToken cancellationToken = default)
        {
            var result = await SendAsync<LoginResponse>(HttpMethod.Post, "api/auth/login", new LoginRequest { Username = username, Password = password }, cancellationToken).ConfigureAwait(false);
            if (!result.Success) LogManager.Instance.LogInfo("服务器登录失败：" + result.Message);
            return result.Success;
        }

        public async Task<MutationResponse> RegisterAsync(string username, string password, int departmentId, string departmentName = "", string realName = null, string gender = null, string phone = null, string email = null, string role = "user", CancellationToken cancellationToken = default)
        {
            return await SendAsync<MutationResponse>(HttpMethod.Post, "api/auth/register", new RegisterRequest { Username = username, Password = password, DepartmentId = departmentId, DepartmentName = departmentName, RealName = realName, Gender = gender, Phone = phone, Email = email, Role = role }, cancellationToken).ConfigureAwait(false);
        }

        // 调用服务器修改密码接口，客户端不直接访问用户数据库。
        public async Task<MutationResponse> ResetPasswordAsync(string username, string phone, string email, string newPassword, CancellationToken cancellationToken = default)
        {
            return await SendAsync<MutationResponse>(HttpMethod.Post, "api/auth/reset-password", new ResetPasswordRequest { Username = username, Phone = phone, Email = email, NewPassword = newPassword }, cancellationToken).ConfigureAwait(false);
        }

        public async Task<List<UserModel>> GetUsersAsync(int departmentId, CancellationToken cancellationToken = default)
        {
            return await SendAsync<List<UserModel>>(HttpMethod.Get, "api/users?departmentId=" + departmentId, null, cancellationToken).ConfigureAwait(false);
        }

        public async Task<MutationResponse> AddUserAsync(string username, string password, int departmentId, string departmentName, string role, bool isActive, string realName, string gender, string phone, string email, CancellationToken cancellationToken = default)
        {
            return await SendAsync<MutationResponse>(HttpMethod.Post, "api/users", new UserRequest { Username = username, Password = password, DepartmentId = departmentId, DepartmentName = departmentName, Role = role, IsActive = isActive, RealName = realName, Gender = gender, Phone = phone, Email = email }, cancellationToken).ConfigureAwait(false);
        }

        public async Task<MutationResponse> UpdateUserAsync(int id, string username, string role, bool isActive, int? departmentId, string departmentName, string password, string realName, string gender, string phone, string email, CancellationToken cancellationToken = default)
        {
            return await SendAsync<MutationResponse>(HttpMethod.Put, "api/users/" + id, new UserRequest { Username = username, Password = password, DepartmentId = departmentId, DepartmentName = departmentName, Role = role, IsActive = isActive, RealName = realName, Gender = gender, Phone = phone, Email = email }, cancellationToken).ConfigureAwait(false);
        }

        public async Task<MutationResponse> DeleteUserAsync(int id, CancellationToken cancellationToken = default) => await SendAsync<MutationResponse>(HttpMethod.Delete, "api/users/" + id, null, cancellationToken).ConfigureAwait(false);

        public async Task<MutationResponse> AssignUserToDepartmentAsync(string username, int departmentId, CancellationToken cancellationToken = default)
            => await SendAsync<MutationResponse>(HttpMethod.Post, "api/users/assign", new UserRequest { Username = username, DepartmentId = departmentId }, cancellationToken).ConfigureAwait(false);

        public async Task<MutationResponse> AddDepartmentAsync(string name, string displayName, string description, int sortOrder, int? managerUserId, CancellationToken cancellationToken = default)
            => await SendAsync<MutationResponse>(HttpMethod.Post, "api/departments", new DepartmentRequest { Name = name, DisplayName = displayName, Description = description, SortOrder = sortOrder, ManagerUserId = managerUserId }, cancellationToken).ConfigureAwait(false);

        public async Task<MutationResponse> UpdateDepartmentAsync(int id, string name, string displayName, string description, int sortOrder, int? managerUserId, bool isActive, CancellationToken cancellationToken = default)
            => await SendAsync<MutationResponse>(HttpMethod.Put, "api/departments/" + id, new DepartmentRequest { Name = name, DisplayName = displayName, Description = description, SortOrder = sortOrder, ManagerUserId = managerUserId, IsActive = isActive }, cancellationToken).ConfigureAwait(false);

        public async Task<MutationResponse> DeleteDepartmentAsync(int id, CancellationToken cancellationToken = default) => await SendAsync<MutationResponse>(HttpMethod.Delete, "api/departments/" + id, null, cancellationToken).ConfigureAwait(false);

        public async Task<MutationResponse> SyncDepartmentsFromCategoriesAsync(CancellationToken cancellationToken = default)
            => await SendAsync<MutationResponse>(HttpMethod.Post, "api/departments/sync-from-categories", null, cancellationToken).ConfigureAwait(false);

        private static async Task<T> SendAsync<T>(HttpMethod method, string path, object body, CancellationToken cancellationToken)
        {
            using (var request = new HttpRequestMessage(method, ApiEndpoint.Build(path)))
            {
                if (body != null) request.Content = new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json");
                using (var response = await HttpClient.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    string text = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode) throw new HttpRequestException("服务器接口请求失败，HTTP " + (int)response.StatusCode + "：" + text);
                    var value = JsonConvert.DeserializeObject<T>(text);
                    if (value == null) throw new InvalidOperationException("服务器返回空响应。");
                    return value;
                }
            }
        }
    }
}
