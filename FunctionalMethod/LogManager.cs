using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using GB_NewCadPlus_IV.UniFiedStandards;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 日志管理器（单例），支持动态路径切换、按小时存档、异步安全写入。
    /// </summary>
    public class LogManager
    {
        #region 单例
        private static readonly Lazy<LogManager> _instance = new(() => new LogManager());
        public static LogManager Instance => _instance.Value;
        #endregion

        #region 字段
        public volatile string _logRootDirectory; // 日志根目录（不含小时子目录）
        private readonly SemaphoreSlim _writeSemaphore = new SemaphoreSlim(1, 1);
        private bool _isInitialized = false;
        #endregion

        private LogManager()
        {
            // 1. 设置默认根目录（AppData 下）
            _logRootDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "GB_CADPLUS",
                "Logs");

            // 2. 如果有用户指定的存储路径，则优先使用
            if (!string.IsNullOrWhiteSpace(GetPath._cacheStoragePath))
            {
                _logRootDirectory = Path.Combine(GetPath._cacheStoragePath, "Logs");
            }

            // 3. 确保目录存在
            EnsureDirectoryExists(_logRootDirectory);
            _isInitialized = true;

            // 4. 写一条初始化日志（此时路径已确定）
            LogInfo("日志管理器初始化完成");
        }

        /// <summary>
        /// 动态切换日志根目录（通常在用户改变存储路径时调用）。
        /// 传入 null 或空字符串则恢复为 AppData 默认路径。
        /// </summary>
        public void SetLogRoot(string storagePath)
        {
            if (string.IsNullOrWhiteSpace(storagePath))
            {
                _logRootDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "GB_CADPLUS",
                    "Logs");
            }
            else
            {
                _logRootDirectory = Path.Combine(storagePath, "Logs");
            }

            EnsureDirectoryExists(_logRootDirectory);
        }

        #region 公开日志方法
        public void LogInfo(string message) => EnqueueWrite("INFO", message);
        public void LogWarning(string message) => EnqueueWrite("WARN", message);
        public void LogError(string message) => EnqueueWrite("ERROR", message);
        public void LogDebug(string message) => EnqueueWrite("DEBUG", message);
        #endregion

        #region 内部实现

        /// <summary>
        /// 将日志放入后台任务异步写入，不阻塞调用线程。
        /// </summary>
        private void EnqueueWrite(string level, string message)
        {
            if (!_isInitialized) return;

            // 捕获当前根目录快照，避免 SetLogRoot 中途改变路径导致混乱
            string rootDir = _logRootDirectory;

            // 异步写日志（丢弃 Task，异常由内部处理）
            _ = Task.Run(async () =>
            {
                try
                {
                    await WriteLogAsync(rootDir, level, message);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"日志写入失败: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 异步安全地将一行日志写入文件（按小时生成文件）。
        /// </summary>
        private async Task WriteLogAsync(string rootDir, string level, string message)
        {
            // ★ 改动：按小时生成文件名，格式 log_yyyyMMddHH.txt
            string hourSuffix = DateTime.Now.ToString("yyyyMMddHH");
            string fileName = $"log_{hourSuffix}.txt";
            string fullPath = Path.Combine(rootDir, fileName);

            // 确保目录存在（防止中途被手动删除）
            string? dir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            string logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}";

            // 异步等待写锁，保证写入顺序和文件完整性
            await _writeSemaphore.WaitAsync();
            try
            {
                // 使用 Task.Run 包装同步写入（.NET Framework 不支持真正的异步 AppendAllTextAsync）
                await Task.Run(() => File.AppendAllText(fullPath, logLine + Environment.NewLine));
            }
            finally
            {
                _writeSemaphore.Release();
            }

            // 同时输出到调试窗口
            System.Diagnostics.Debug.WriteLine(logLine);
        }

        /// <summary>
        /// 创建目录（不存在时），忽略异常。
        /// </summary>
        private static void EnsureDirectoryExists(string path)
        {
            try
            {
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
            }
            catch { /* 忽略目录创建失败 */ }
        }

        /// <summary>
        /// 获取当前小时对应的日志文件完整路径。
        /// 例如：2025年3月15日10点 -> log_2025031510.txt
        /// </summary>
        public string GetCurrentLogFilePath()
        {
            // ★ 改动：同 WriteLogAsync 保持一致，按小时返回
            string hourSuffix = DateTime.Now.ToString("yyyyMMddHH");
            string fileName = $"log_{hourSuffix}.txt";
            return Path.Combine(_logRootDirectory, fileName);
        }

        #endregion
    }
}
