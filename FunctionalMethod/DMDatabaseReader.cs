using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using Dapper;
using GB_NewCadPlus_IV.FunctionalMethod;
using Newtonsoft.Json;

namespace GB_NewCadPlus_IV.DMDatabaseReader
{
    // 修复方法：将原先嵌套在 DMDatabaseReader 静态类内部的 DTO 类移动到命名空间级别。
    // 这样在其他地方引用时就不会因为“类名与命名空间同名”而产生歧义。

    // 表信息类：定义在命名空间下，方便引用
    public class TableInfo
    {
        public string TableName { get; set; }
        public string TableComment { get; set; }
    }

    // 列信息类
    public class ColumnInfo
    {
        public string ColumnName { get; set; }
        public string DataType { get; set; }
        public int DataLength { get; set; }
        public string Nullable { get; set; }
        public string ColumnComment { get; set; }
    }

    // 外键信息类
    public class ForeignKeyInfo
    {
        public string ConstraintName { get; set; }
        public string ColumnName { get; set; }
        public string ReferencedTable { get; set; }
        public string ReferencedColumn { get; set; }
    }

    // 表的详细信息类
    public class TableDetail
    {
        public string TableName { get; set; }
        public string TableComment { get; set; }
        public List<ColumnInfo> Columns { get; set; }
        public List<string> PrimaryKeys { get; set; }
        public List<ForeignKeyInfo> ForeignKeys { get; set; }
    }

    /// <summary>
    /// 达梦数据库读取工具类
    /// </summary>
    public static class DMDatabaseReader
    {
        /// <summary>
        /// 读取达梦数据库的整体架构信息。
        /// </summary>
        /// <param name="args">数据库连接信息</param>
        public static void DMDatabaseReaderMethod(string[] args)
        {
            if (args == null || args.Length < 4)
            {
                LogManager.Instance.LogError("参数不足。");
                return;
            }

            string server = args[0];
            string port = args[1];
            string username = args[2];
            string password = args[3];

            try
            {
                var authService = new DMAuthService(server, port, username, password);
                authService.EnsureAllTablesExist();

                LogManager.Instance.LogInfo("成功连接到达梦数据库！");

                // 示例：引用已移至命名空间级别的 TableInfo
                var tables = authService.GetTables();
                foreach (var table in tables)
                {
                    LogManager.Instance.LogInfo($"表名: {table.TableName}, 注释: {table.TableComment}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"发生错误：{ex.Message}");
            }
        }
    }
}
