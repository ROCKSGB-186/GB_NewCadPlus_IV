using GB_NewCadPlus_IV.UniFiedStandards;
using System;
using System.Configuration;

namespace GB_NewCadPlus_IV.Helpers
{
    internal static class ApiEndpoint
    {
        private const int FallbackPort = 10010;

        public static void LoadDefaults()
        {
            if (string.IsNullOrWhiteSpace(VariableDictionary._serverIP))
                VariableDictionary._serverIP = GetAppSetting("ApiServerHost", string.Empty);

            if (VariableDictionary._apiPort <= 0)
                VariableDictionary._apiPort = GetPortSetting();
        }

        public static string Build(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
                throw new ArgumentException("API 路径不能为空。", nameof(relativePath));

            LoadDefaults();
            string scheme = GetAppSetting("ApiServerScheme", "http");
            string host = VariableDictionary._serverIP?.Trim() ?? string.Empty;
            int port = VariableDictionary._apiPort > 0 ? VariableDictionary._apiPort : FallbackPort;

            if (string.IsNullOrWhiteSpace(host))
                throw new InvalidOperationException("请先在登录窗口中填写服务器 IP 地址。也可以在 App.config 的 ApiServerHost 中配置。" );

            if (!Uri.CheckHostName(host).Equals(UriHostNameType.Dns) &&
                !Uri.CheckHostName(host).Equals(UriHostNameType.IPv4) &&
                !Uri.CheckHostName(host).Equals(UriHostNameType.IPv6))
                throw new InvalidOperationException("API 服务器地址无效。请检查登录窗口中的服务器 IP 或主机名。");

            return new Uri(new Uri($"{scheme}://{host}:{port}/"), relativePath.TrimStart('/')).ToString();
        }

        public static string GetBaseUrl()
        {
            LoadDefaults();
            string scheme = GetAppSetting("ApiServerScheme", "http");
            string host = VariableDictionary._serverIP?.Trim() ?? string.Empty;
            int port = VariableDictionary._apiPort > 0 ? VariableDictionary._apiPort : FallbackPort;
            if (string.IsNullOrWhiteSpace(host))
                throw new InvalidOperationException("请先在登录窗口中填写服务器 IP 地址。也可以在 App.config 的 ApiServerHost 中配置。" );
            return new Uri($"{scheme}://{host}:{port}/").GetLeftPart(UriPartial.Authority);
        }

        public static string GetDisplayAddress()
        {
            LoadDefaults();
            return $"{VariableDictionary._serverIP}:{VariableDictionary._apiPort}";
        }

        private static int GetPortSetting()
        {
            return int.TryParse(GetAppSetting("ApiServerPort", FallbackPort.ToString()), out int port) && port > 0
                ? port
                : FallbackPort;
        }

        private static string GetAppSetting(string key, string fallback)
        {
            string value = ConfigurationManager.AppSettings[key];
            return string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
        }
    }
}
