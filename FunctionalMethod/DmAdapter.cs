using System;
using System.Data;
using Dm;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 达梦 (DM) 数据库适配器实现
    /// </summary>
    public class DmAdapter : IDatabaseAdapter
    {
        private readonly string _connectionString;

        public string DatabaseType => "DM";

        public DmAdapter(string connectionString)
        {
            _connectionString = connectionString;
        }

        public IDbConnection CreateConnection()
        {
            return new DmConnection(_connectionString);
        }

        public string NormalizeSql(string sql)
        {
            // 达梦建议使用 : 作为占位符，将 SQL 中的 @ 替换为 :
            // 注意：此简单替换假设 @ 仅用于占位符且不出现在字符串字面量中
            return sql.Replace("@", ":");
        }

        public void AddParameter(IDbCommand cmd, string name, object value)
        {
            var p = cmd.CreateParameter();
            // 达梦参数名在命令中用 :name 引用，但 ParameterName 属性通常不含前缀或包含适配的前缀
            p.ParameterName = name.StartsWith(":") ? name.Substring(1) : name;
            p.Value = value ?? DBNull.Value;
            cmd.Parameters.Add(p);
        }

        public void ApplySchema(IDbConnection connection, string schemaName)
        {
            if (connection.State == ConnectionState.Open && !string.IsNullOrWhiteSpace(schemaName))
            {
                using var cmd = connection.CreateCommand();
                cmd.CommandText = $"SET SCHEMA \"{schemaName.ToUpperInvariant()}\"";
                cmd.ExecuteNonQuery();
            }
        }
    }
}
