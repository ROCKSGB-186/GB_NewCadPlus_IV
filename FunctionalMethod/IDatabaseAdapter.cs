using System;
using System.Data;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 数据库适配器接口：定义统一的数据库操作行为
    /// 用于同时支持 MySQL 与 达梦 (DM) 数据库
    /// </summary>
    public interface IDatabaseAdapter
    {
        /// <summary>
        /// 数据库类型标识
        /// </summary>
        string DatabaseType { get; }

        /// <summary>
        /// 创建并返回一个未打开的 IDbConnection 实例
        /// </summary>
        /// <returns>数据库连接对象</returns>
        IDbConnection CreateConnection();

        /// <summary>
        /// 规范化 SQL 语句，将中性占位符转换为目标数据库的占位符（如 @ -> :）
        /// </summary>
        /// <param name="sql">原始 SQL 语句</param>
        /// <returns>转换后的 SQL 语句</returns>
        string NormalizeSql(string sql);

        /// <summary>
        /// 为 IDbCommand 添加参数，自动处理不同数据库的参数前缀和类型映射
        /// </summary>
        /// <param name="cmd">数据库命令对象</param>
        /// <param name="name">参数名称（不含前缀）</param>
        /// <param name="value">参数值</param>
        void AddParameter(IDbCommand cmd, string name, object value);

        /// <summary>
        /// 应用 Schema（主要针对达梦数据库）
        /// </summary>
        /// <param name="connection">已打开的连接</param>
        /// <param name="schemaName">Schema 名称</param>
        void ApplySchema(IDbConnection connection, string schemaName);
    }
}
