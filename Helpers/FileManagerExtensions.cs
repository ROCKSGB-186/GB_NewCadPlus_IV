using GB_NewCadPlus_LM.FunctionalMethod;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
namespace GB_NewCadPlus_LM.Helpers
{
    /// <summary>
    /// FileManager 的扩展占位，用于覆盖已有存储项的内容。
    /// 请在此处实现实际的上传/覆盖逻辑（调用服务端覆盖 API 或替换物理文件），
    /// 目前抛出 NotImplementedException 以提示后续实现细节并解决编译错误。
    /// </summary>
    public static class FileManagerExtensions
    {
        /// <summary>
        /// 优先反射调用 FileManager 中已有的覆盖方法（例如 ReplaceFileContent/ReplaceFile 等）。
        /// 回退：尝试从 storage 寻找常见路径属性并直接把本地文件复制覆盖到该路径（含备份）。
        /// 若都不可行，会抛出 NotImplementedException，提示实现后端 API。
        /// </summary>
        public static void ReplaceFileContent(this object fileManager, object storage, string localPath)
        {
            if (fileManager == null) throw new ArgumentNullException(nameof(fileManager));
            if (storage == null) throw new ArgumentNullException(nameof(storage));
            if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
                throw new FileNotFoundException("本地替换文件不存在。", localPath);

            var fmType = fileManager.GetType();
            var storageType = storage.GetType();
            var candidateNames = new[] { "ReplaceFileContent", "ReplaceFile", "UploadAndReplace", "ReplaceFileByPath", "OverwriteFile" };

            // 尝试反射调用： (storage, string)
            foreach (var name in candidateNames)
            {
                var m = fmType.GetMethod(name, new Type[] { storageType, typeof(string) }) ??
                        fmType.GetMethod(name, new Type[] { typeof(object), typeof(string) });
                if (m != null)
                {
                    try
                    {
                        m.Invoke(fileManager, new object[] { storage, localPath });
                        return;
                    }
                    catch (TargetInvocationException tie)
                    {
                        throw new InvalidOperationException($"调用 {fmType.Name}.{name} 时发生异常: {tie.InnerException?.Message ?? tie.Message}", tie);
                    }
                }
            }

            // 回退：直接从 storage 找到可写路径并复制（FilePath / Path / StoragePath / ServerPath）
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

            if (string.IsNullOrWhiteSpace(destPath))
            {
                throw new NotImplementedException("未找到 FileManager 覆盖方法，且从 storage 未能推断出可写本地路径。" +
                    "若文件保存在远端或数据库，请在 FileManager/DatabaseManager 中实现替换 API（例如：ReplaceFileContent(FileStorage, string)）。");
            }

            // 做备份并覆盖
            try
            {
                var dir = System.IO.Path.GetDirectoryName(destPath);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                if (File.Exists(destPath))
                {
                    var bak = destPath + ".bak";
                    File.Copy(destPath, bak, overwrite: true);
                }
                File.Copy(localPath, destPath, overwrite: true);
            }
            catch (Exception ex)
            {
                throw new IOException($"直接用本地文件覆盖目标路径失败: {ex.Message}", ex);
            }
        }
    }
}
