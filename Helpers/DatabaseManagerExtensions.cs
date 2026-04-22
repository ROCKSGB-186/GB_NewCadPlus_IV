using GB_NewCadPlus_IV.FunctionalMethod;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// DatabaseManager 的扩展占位：按记录替换二进制内容（适用于将文件保存在 DB BLOB 的场景）。
    /// 实现时请使用参数化 SQL、事务和备份策略，避免数据损坏。
    /// </summary>
    public static class DatabaseManagerExtensions
    {
        /// <summary>
        /// 优先反射调用 DatabaseManager 中已有的二进制替换方法（例如 ReplaceFileBinary/UpdateFileBinary 等）。
        /// 回退：若 storage 包含可写路径，则写入文件系统（适用于部分混合存储场景）。
        /// 否则抛出 NotImplementedException，提示实现数据库替换逻辑。
        /// </summary>
        public static void ReplaceFileBinary(this object dbManager, object storage, byte[] bytes)
        {
            if (dbManager == null) throw new ArgumentNullException(nameof(dbManager));
            if (storage == null) throw new ArgumentNullException(nameof(storage));
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));

            var dbType = dbManager.GetType();
            var storageType = storage.GetType();
            var candidateNames = new[] { "ReplaceFileBinary", "ReplaceFileContent", "UpdateFileBinary", "SaveFileBinary", "SetFileBinary" };

            // 尝试反射调用： (storage, byte[])
            foreach (var name in candidateNames)
            {
                var m = dbType.GetMethod(name, new Type[] { storageType, typeof(byte[]) }) ??
                        dbType.GetMethod(name, new Type[] { typeof(object), typeof(byte[]) });
                if (m != null)
                {
                    try
                    {
                        m.Invoke(dbManager, new object[] { storage, bytes });
                        return;
                    }
                    catch (TargetInvocationException tie)
                    {
                        throw new InvalidOperationException($"调用 {dbType.Name}.{name} 时发生异常: {tie.InnerException?.Message ?? tie.Message}", tie);
                    }
                }
            }

            // 回退：尝试找到 storage 的本地路径，然后把 bytes 写入（仅在混合场景可用）
            var candidates = new[] { "FilePath", "Path", "StoragePath", "ServerPath", "LocalPath" };
            string? destPath = null;
            foreach (var propName in candidates)
            {
                var prop = storageType.GetProperty(propName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (prop != null)
                {
                    var val = prop.GetValue(storage) as string;
                    if (!string.IsNullOrWhiteSpace(val))
                    {
                        destPath = val;
                        break;
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(destPath))
            {
                try
                {
                    var dir = Path.GetDirectoryName(destPath);
                    if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                        Directory.CreateDirectory(dir);

                    if (File.Exists(destPath))
                    {
                        var bak = destPath + ".bak";
                        File.Copy(destPath, bak, overwrite: true);
                    }

                    File.WriteAllBytes(destPath, bytes);
                    return;
                }
                catch (Exception ex)
                {
                    throw new IOException($"将二进制写入目标路径失败: {ex.Message}", ex);
                }
            }

            throw new NotImplementedException("未在 DatabaseManager 中找到用于替换二进制的实现（例如 ReplaceFileBinary(FileStorage, byte[])）。请在后端/DatabaseManager 中实现该接口，或更新此扩展以匹配你的实际签名。");
        }
    }
}
