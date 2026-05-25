using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 用户字段规范化工具类，提供统一的字符串兜底逻辑
    /// </summary>
    public static class UserFieldNormalizer
    {
        public static string NormalizeGender(string? gender)
        {
            var g = (gender ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(g)) return "无信息";
            if (string.Equals(g, "男", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(g, "male", StringComparison.OrdinalIgnoreCase) ||
                g == "M") return "男";
            if (string.Equals(g, "女", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(g, "female", StringComparison.OrdinalIgnoreCase) ||
                g == "F") return "女";
            return "无信息";
        }

        public static string NormalizeRole(string? role)
        {
            return string.IsNullOrWhiteSpace(role) ? "user" : role.Trim();
        }

        public static string NormalizeDepartmentName(string? departmentName)
        {
            var value = (departmentName ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(value) ? "未分配" : value;
        }

        public static string NormalizeRealName(string? realName, string? username)
        {
            var value = (realName ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(value)) return value;

            var user = (username ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(user) ? "未命名用户" : user;
        }

        public static string NormalizePhone(string? phone)
        {
            var value = (phone ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(value) ? "未填写" : value;
        }

        public static string NormalizeEmail(string? email)
        {
            var value = (email ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(value) ? "未填写" : value;
        }
    }
}
