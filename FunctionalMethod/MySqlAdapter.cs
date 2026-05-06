using System;
using System.Data;
using MySql.Data.MySqlClient;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// MySQL 数据库适配器实现
    /// </summary>
    public class MySqlAdapter : IDatabaseAdapter
    {
        private readonly string _connectionString;

        public string DatabaseType => "MySQL";

        public MySqlAdapter(string connectionString)
        {
            _connectionString = connectionString;
        }

        public IDbConnection CreateConnection()
        {
            return new MySqlConnection(_connectionString);
        }

        public string NormalizeSql(string sql)
        {
            // MySQL 默认使用 @ 作为占位符，通常不需要转换
            return sql;
        }

        public void AddParameter(IDbCommand cmd, string name, object value)
        {
            var p = cmd.CreateParameter();
            p.ParameterName = name.StartsWith("@") ? name : "@" + name;
            p.Value = value ?? DBNull.Value;
            cmd.Parameters.Add(p);
        }

        public void ApplySchema(IDbConnection connection, string schemaName)
        {
            // MySQL 通常通过连接字符串中的 Database 或 USE [db] 切换，此处可留空或执行 USE
            if (connection.State == ConnectionState.Open && !string.IsNullOrWhiteSpace(schemaName))
            {
                using var cmd = connection.CreateCommand();
                cmd.CommandText = $"USE `{schemaName}`";
                cmd.ExecuteNonQuery();
            }
        }
    }
}
