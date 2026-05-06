using System;
using System.Data;
using Dm;
using GB_NewCadPlus_IV.UniFiedStandards;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 数据库连接工厂与方言辅助
    /// </summary>
    public static class DbProvider
    {
        /// <summary>
        /// 创建达梦连接（统一入口）
        /// </summary>
        public static IDbConnection GetConnection()
        {
            var host = string.IsNullOrWhiteSpace(VariableDictionary._serverIP) ? "127.0.0.1" : VariableDictionary._serverIP.Trim();
            var port = VariableDictionary._serverPort > 0 ? VariableDictionary._serverPort : 5236;
            var user = string.IsNullOrWhiteSpace(VariableDictionary._userName) ? "SYSDBA" : VariableDictionary._userName.Trim();
            var pwd = VariableDictionary._passWord ?? "SYSDBA";

            // 达梦连接字符串中 Database/Schema 可选，若未配置则由登录用户默认模式决定
            string databasePart = string.IsNullOrWhiteSpace(VariableDictionary._dataBaseName)
                ? string.Empty
                : $"Database={VariableDictionary._dataBaseName.Trim()};";

            VariableDictionary._newConnectionString = $"Server={host};Port={port};{databasePart}User Id={user};Password={pwd};";
            return new DmConnection(VariableDictionary._newConnectionString);
        }

        /// <summary>
        /// 参数前缀
        /// </summary>
        public static string ParamPrefix => ":";
    }
}
