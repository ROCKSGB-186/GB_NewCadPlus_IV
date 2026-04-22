using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 日志管理器类
    /// </summary>
    public class LogManager
    {
        private static LogManager _instance;
        private static readonly object _lock = new object();
        private string _logFilePath;
        private bool _isInitialized = false;

        private LogManager()
        {
            Initialize();
        }
        /// <summary>
        ///  获取日志管理器的实例
        /// </summary>
        public static LogManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new LogManager();
                        }
                    }
                }
                return _instance;
            }
        }
        /// <summary>
        ///  初始化日志管理器
        /// </summary>
        private void Initialize()
        {
            try
            {
                // 创建日志目录
                string logDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "GB_CADPLUS", "Logs");

                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                // 创建日志文件名（按天命名）
                string fileName = $"GB_NewCadPlus_IV_{DateTime.Now:yyyyMMdd}.log";
                _logFilePath = Path.Combine(logDirectory, fileName);

                // 确保日志文件存在
                if (!File.Exists(_logFilePath))
                {
                    File.WriteAllText(_logFilePath, "");
                }

                _isInitialized = true;

                // 记录初始化日志
                LogInfo("日志管理器初始化完成");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"日志管理器初始化失败: {ex.Message}");
            }
        }
        /// <summary>
        ///  记录信息
        /// </summary>
        /// <param name="message"></param>
        public void LogInfo(string message)
        {
            WriteLog("INFO", message);
        }
        /// <summary>
        ///  记录警告
        /// </summary>
        /// <param name="message"></param>
        public void LogWarning(string message)
        {
            WriteLog("WARN", message);
        }
        /// <summary>
        ///   记录错误
        /// </summary>
        /// <param name="message"></param>
        public void LogError(string message)
        {
            WriteLog("ERROR", message);
        }
        /// <summary>
        ///   记录调试信息
        /// </summary>
        /// <param name="message"></param>
        public void LogDebug(string message)
        {
            WriteLog("DEBUG", message);
        }
        /// <summary>
        ///  写入日志
        /// </summary>
        /// <param name="level"></param>
        /// <param name="message"></param>
        private void WriteLog(string level, string message)
        {
            if (!_isInitialized)
                return;

            try
            {
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}";

                // 同时输出到调试窗口和日志文件
                System.Diagnostics.Debug.WriteLine(logEntry);

                // 异步写入日志文件
                Task.Run(() =>
                {
                    try
                    {
                        File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"写入日志文件失败: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"记录日志失败: {ex.Message}");
            }
        }
        /// <summary>
        ///  获取日志文件路径
        /// </summary>
        public string LogFilePath => _logFilePath;
    }
}
