using System.Drawing.Drawing2D;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.UniFiedStandards;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Serialization;
using static GB_NewCadPlus_IV.WpfMainWindow;
using DataTable = System.Data.DataTable;
using MessageBox = System.Windows.MessageBox;
using TextBox = System.Windows.Controls.TextBox;


namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 创建新的文件管理服务类
    /// </summary>
    public class FileManager
    {
        /// <summary>
        /// 数据库管理器
        /// </summary>
        private readonly DatabaseManager _databaseManager ;

        /// <summary>
        /// 分类管理器
        /// </summary>
        private CategoryManager _categoryManager;

        /// <summary>
        /// 分类管理器
        /// </summary>
        /// <param name="databaseManager">数据库管理器</param>
        /// <param name="baseStoragePath">基础存储路径</param>
        /// <param name="useDPath">是否使用D盘</param>
        public FileManager(DatabaseManager databaseManager)
        {
            _databaseManager = databaseManager;// 数据库管理器
        }
        
        /// <summary>
        /// 规范化配置路径（去空白、去包裹引号、展开环境变量）
        /// </summary>
        /// <param name="rawPath">原始路径字符串</param>
        /// <returns>规范化后的路径；无效时返回空字符串</returns>
        private static string NormalizeConfiguredPath(string rawPath)
        {
            // 空值直接返回空，调用方统一处理
            if (string.IsNullOrWhiteSpace(rawPath))
            {
                return string.Empty;
            }

            // 去首尾空白并去除首尾引号（常见于手工录入配置）
            string normalized = rawPath.Trim().Trim('"').Trim();

            // 再次判空，防止出现仅引号/空白的配置值
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }

            // 展开环境变量，支持如 %ProgramData% 这类配置写法
            normalized = Environment.ExpandEnvironmentVariables(normalized);

            return normalized;
        }

        /// <summary>
        /// 判断是否为盘符绝对路径（例如 D:\\xx）
        /// </summary>
        /// <param name="path">待判断路径</param>
        /// <returns>是盘符路径返回 true</returns>
        private static bool IsDriveAbsolutePath(string path)
        {
            // 空值直接否定
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            // 统一去空白后判断格式：字母 + 冒号 + 反斜杠
            string text = path.Trim();
            return text.Length >= 3
                   && char.IsLetter(text[0])
                   && text[1] == ':'
                   && (text[2] == '\\' || text[2] == '/');
        }

        /// <summary>
        /// 检查路径是否更像“文件路径”而不是“目录路径”。
        /// </summary>
        private static bool LooksLikeFilePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            var fileName = Path.GetFileName(path.Trim());
            return !string.IsNullOrWhiteSpace(fileName) && fileName.Contains(".");
        }

        /// <summary>
        /// 若传入的是文件路径，则自动转换为目录路径。
        /// </summary>
        private static string EnsureDirectoryPath(string path, string operationName)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            var normalized = path.Trim();
            if (!LooksLikeFilePath(normalized))
            {
                return normalized;
            }

            var dir = Path.GetDirectoryName(normalized);
            if (string.IsNullOrWhiteSpace(dir))
            {
                return normalized;
            }

            LogManager.Instance.LogWarning($"[{operationName}] 检测到配置值疑似文件路径，自动回退到目录路径: {dir}");
            return dir;
        }

        /// <summary>
        /// 解析 UNC 路径中的服务器名与共享名。
        /// </summary>
        private static bool TryParseUncPath(string uncPath, out string server, out string share, out string tail)
        {
            server = string.Empty;
            share = string.Empty;
            tail = string.Empty;

            if (string.IsNullOrWhiteSpace(uncPath) || !uncPath.StartsWith("\\\\", StringComparison.Ordinal))
            {
                return false;
            }

            var body = uncPath.TrimStart('\\');
            var parts = body.Split(new[] { '\\' }, StringSplitOptions.None);
            if (parts.Length < 2)
            {
                return false;
            }

            server = parts[0];
            share = parts[1];
            tail = parts.Length > 2 ? string.Join("\\", parts.Skip(2)) : string.Empty;
            return !string.IsNullOrWhiteSpace(server) && !string.IsNullOrWhiteSpace(share);
        }

        /// <summary>
        /// 仅探测共享根是否可访问。
        /// </summary>
        private static bool IsShareReachable(string server, string share)
        {
            try
            {
                var shareRoot = $"\\\\{server}\\{share}";
                return Directory.Exists(shareRoot);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 将 UNC 路径在 share 与 share$ 之间自动切换，优先返回可访问路径。
        /// </summary>
        private static string ResolveReachableUncPath(string uncPath, string operationName)
        {
            if (!TryParseUncPath(uncPath, out var server, out var share, out var tail))
            {
                return uncPath;
            }

            var candidates = new List<string>();
            candidates.Add(share);

            if (share.EndsWith("$", StringComparison.Ordinal))
            {
                candidates.Add(share.TrimEnd('$'));
            }
            else if (share.Length == 1 && char.IsLetter(share[0]))
            {
                candidates.Add(share + "$");
            }

            foreach (var candidateShare in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!IsShareReachable(server, candidateShare))
                {
                    continue;
                }

                var resolved = string.IsNullOrWhiteSpace(tail)
                    ? $"\\\\{server}\\{candidateShare}"
                    : $"\\\\{server}\\{candidateShare}\\{tail}";

                if (!string.Equals(resolved, uncPath, StringComparison.OrdinalIgnoreCase))
                {
                    LogManager.Instance.LogInfo($"[{operationName}] UNC 共享自动切换成功: {uncPath} -> {resolved}");
                }

                return resolved;
            }

            return uncPath;
        }

        /// <summary>
        /// 将配置路径解析为“服务器可访问路径”。
        /// 规则：
        /// 1) UNC 路径优先直接使用，并在 share/share$ 间自动尝试；
        /// 2) 盘符路径（D:\\...）结合服务器 IP 转为 \\IP\\d\\... 并自动尝试 \\IP\\d$\\...;
        /// 3) 其余路径原样返回。
        /// </summary>
        /// <param name="configuredPath">配置中的路径</param>
        /// <param name="serverIp">登录页服务器 IP</param>
        /// <param name="operationName">操作名（用于日志）</param>
        /// <returns>服务器可访问路径</returns>
        /// <exception cref="InvalidOperationException">盘符路径但服务器 IP 为空时抛出</exception>
        private static string ResolveServerStoragePath(string configuredPath, string serverIp, string operationName)
        {
            // 先做规范化，统一处理空白、引号、环境变量
            string normalized = NormalizeConfiguredPath(configuredPath);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }

            // 已是 UNC 路径：尝试 share/share$ 自动切换
            if (normalized.StartsWith("\\\\", StringComparison.Ordinal))
            {
                return ResolveReachableUncPath(normalized, operationName);
            }

            // 盘符路径按“服务器共享目录”规则转换
            if (IsDriveAbsolutePath(normalized))
            {
                string ip = NormalizeConfiguredPath(serverIp);
                if (string.IsNullOrWhiteSpace(ip))
                {
                    throw new InvalidOperationException($"检测到盘符路径 {normalized}，但服务器IP为空，无法转换为服务器共享路径。请先在登录页配置服务器IP。");
                }

                // 取盘符字母（如 D）并转为小写共享名（如 d）
                string driveShareName = char.ToLowerInvariant(normalized[0]).ToString();
                // 去掉 "D:\\" 前缀，保留剩余子路径
                string tailPath = normalized.Length > 3 ? normalized.Substring(3) : string.Empty;

                // 先按普通共享拼接，再交给 share/share$ 自动切换
                var unc = string.IsNullOrWhiteSpace(tailPath)
                    ? $"\\\\{ip}\\{driveShareName}"
                    : $"\\\\{ip}\\{driveShareName}\\{tailPath}";

                return ResolveReachableUncPath(unc, operationName);
            }

            // 非 UNC/非盘符路径原样返回
            return normalized;
        }

        /// <summary>
        /// 解析当前上传操作应使用的存储根路径
        /// </summary>
        /// <param name="databaseManager">数据库管理器</param>
        /// <param name="operationName">操作名称（用于日志定位）</param>
        /// <returns>最终可用的存储根路径</returns>
        private async Task<string> ResolveStorageRootPathAsync(DatabaseManager databaseManager, string operationName)
        {
            // 本地兜底路径（仅数据库不可用时才允许）
            string fallbackPath = GetPath.AppDataPath;

            // 读取当前登录服务器 IP（用于将 D:\\... 转换为 \\IP\\D$\\...）
            string serverIp = NormalizeConfiguredPath(VariableDictionary._serverIP);

            LogManager.Instance.LogInfo($"[{operationName}] 开始解析存储根路径。serverIp={serverIp}, fallbackPath={fallbackPath}");

            // 数据库不可用时，保留旧行为：允许回退本地路径
            if (databaseManager == null || !databaseManager.IsDatabaseAvailable)
            {
                LogManager.Instance.LogWarning($"[{operationName}] 数据库不可用，使用本地回退路径: {fallbackPath}");
                return EnsureDirectoryPath(fallbackPath, operationName);
            }

            // 1) 优先读取系统配置 SourceRoot
            string sourceRootRaw = await databaseManager.GetSystemConfigValueAsync("SourceRoot").ConfigureAwait(false);
            LogManager.Instance.LogInfo($"[{operationName}] 读取 SourceRoot 原始值: {sourceRootRaw}");
            // 转成服务器路径
            string sourceRoot = EnsureDirectoryPath(ResolveServerStoragePath(sourceRootRaw, serverIp, operationName), operationName);
            if (!string.IsNullOrWhiteSpace(sourceRoot))
            {
                LogManager.Instance.LogInfo($"[{operationName}] 解析到 SourceRoot(服务器路径): {sourceRoot}");
                return sourceRoot;
            }

            // 2) 兼容旧链路：读取运行时变量中的存储路径（即 TextBoxSetStoragePath）
            LogManager.Instance.LogInfo($"[{operationName}] SourceRoot 为空，读取运行时存储路径 VariableDictionary._cacheStoragePath: {GetPath._cacheStoragePath}");
            string runtimeStoragePath = EnsureDirectoryPath(ResolveServerStoragePath(GetPath._cacheStoragePath, serverIp, operationName), operationName);
            if (!string.IsNullOrWhiteSpace(runtimeStoragePath))
            {
                LogManager.Instance.LogInfo($"[{operationName}] SourceRoot 为空，回退到运行时存储路径(服务器路径): {runtimeStoragePath}");
                return runtimeStoragePath;
            }

            // 3) 继续兜底：读取本地设置文件中的 StoragePath（避免运行时变量尚未同步）
            LogManager.Instance.LogInfo($"[{operationName}] 运行时路径为空，读取 Settings.StoragePath: {Properties.Settings.Default.StoragePath}");
            string settingsStoragePath = EnsureDirectoryPath(ResolveServerStoragePath(Properties.Settings.Default.StoragePath, serverIp, operationName), operationName);
            if (!string.IsNullOrWhiteSpace(settingsStoragePath))
            {
                LogManager.Instance.LogInfo($"[{operationName}] 运行时存储路径为空，回退到设置 StoragePath(服务器路径): {settingsStoragePath}");
                return settingsStoragePath;
            }

            // 4) 最后兜底：使用约定默认路径并转换为服务器共享路径（不是本地写入）
            string defaultServerStoragePath = EnsureDirectoryPath(ResolveServerStoragePath(@"D:\GB_Tools\Cad_Sw_Library", serverIp, operationName), operationName);
            if (!string.IsNullOrWhiteSpace(defaultServerStoragePath))
            {
                LogManager.Instance.LogWarning($"[{operationName}] SourceRoot/运行时路径/设置路径均为空，使用默认服务器路径: {defaultServerStoragePath}");
                return defaultServerStoragePath;
            }

            // 5) 数据库可用但无可用服务器路径配置时，阻止上传
            throw new InvalidOperationException("系统配置 SourceRoot 与运行时存储路径均为空，无法确定服务器存储路径。请先在设置中配置存储路径，并确认登录页服务器IP正确。");
        }
        
        /// <summary>
        /// 获取管道属性保存路径
        /// </summary>
        /// <param name="isOutlet"></param>
        /// <returns></returns>
        public static string GetPipeAttrSavePath(bool isOutlet)
        {
            try
            {
                //缓存路径优先级：1) 用户手动设置的路径（GetPath._cacheStoragePath）；2) 默认本地应用数据路径
                string path = !string.IsNullOrWhiteSpace(GetPath._cacheStoragePath)
                    ? GetPath._cacheStoragePath
                    : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GB_CADPLUS");
                var pipeAttrsDir = Path.Combine(path, "PipeAttrs");// 管道属性文件夹
                if (!Directory.Exists(pipeAttrsDir)) Directory.CreateDirectory(pipeAttrsDir);// 确保目录存在
                return Path.Combine(pipeAttrsDir, isOutlet ? "LastPipeAttrs_Outlet.xml" : "LastPipeAttrs_Inlet.xml");// 返回管道属性文件路径

            }
            catch
            {
                return Path.GetTempFileName(); // 兜底
            }
        }
        
        /// <summary>
        /// 可序列化的键值项，保证序列化输出始终包含 <Key> 与 <Value>
        /// </summary>
        [XmlType("KeyValuePairOfStringString")]
        public class PipeAttrEntry
        {
            /// <summary>
            /// 键
            /// </summary>
            [XmlElement("Key")]
            public string Key { get; set; } = string.Empty;
            /// <summary>
            /// 值
            /// </summary>
            [XmlElement("Value")]
            public string Value { get; set; } = string.Empty;
        }

        /// <summary>
        /// 读取管道属性
        /// </summary>
        /// <param name="isOutlet"></param>
        /// <returns></returns>
        public static Dictionary<string, string> LoadLastPipeAttributes(bool isOutlet)
        {
            // 获取入口/出口属性文件路径
            var path = GetPipeAttrSavePath(isOutlet);

            // 若文件不存在，则自动创建一个“空结构”的XML文件，保证后续有物理文件
            if (!File.Exists(path))
            {
                try
                {
                    // 确保目录存在
                    var dir = Path.GetDirectoryName(path);
                    if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    // 创建空列表并序列化到文件（不是空文本，避免反序列化失败）
                    var xsCreate = new XmlSerializer(typeof(List<PipeAttrEntry>));
                    var settings = new System.Xml.XmlWriterSettings
                    {
                        Indent = true,
                        Encoding = new System.Text.UTF8Encoding(true),
                        NewLineChars = "\r\n"
                    };

                    using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                    using (var xw = System.Xml.XmlWriter.Create(fs, settings))
                    {
                        xsCreate.Serialize(xw, new List<PipeAttrEntry>());
                        xw.Flush();
                        fs.Flush(true);
                    }

                    System.Diagnostics.Debug.WriteLine($"[FileManager] Auto-created empty pipe attrs file: {path}");
                }
                catch (Exception exCreate)
                {
                    System.Diagnostics.Debug.WriteLine($"[FileManager] Auto-create pipe attrs file failed: {exCreate.Message}");
                }

                // 首次创建后返回空字典
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            try
            {
                // 按既有格式反序列化
                var xs = new XmlSerializer(typeof(List<PipeAttrEntry>));
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var obj = xs.Deserialize(fs) as List<PipeAttrEntry>;
                    if (obj == null) return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    // 构建不区分大小写字典
                    var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var e in obj)
                    {
                        if (e == null) continue;
                        var k = (e.Key ?? string.Empty).Trim();
                        if (string.IsNullOrWhiteSpace(k)) continue;
                        var v = (e.Value ?? string.Empty).Trim();
                        dict[k] = v;
                    }

                    System.Diagnostics.Debug.WriteLine($"[FileManager] Loaded {dict.Count} pipe attrs from {path}");
                    return dict;
                }
            }
            catch (Exception ex)
            {
                // 读取失败时返回空字典，避免影响主流程
                System.Diagnostics.Debug.WriteLine($"[FileManager] LoadLastPipeAttributes failed: {ex.Message}");
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// 保存管道属性
        /// </summary>
        /// <param name="isOutlet"></param>
        /// <param name="attrs"></param>
        public static void SaveLastPipeAttributes(bool isOutlet, Dictionary<string, string> attrs)
        {
            // 即使 attrs 为 null，也按“空字典”处理，确保文件可被重建
            if (attrs == null)
            {
                attrs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            // 获取保存路径
            var path = GetPipeAttrSavePath(isOutlet);

            try
            {
                // 确保目录存在
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                // 规范化键值，去空白并合并同名键
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var kv in attrs)
                {
                    var k = (kv.Key ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(k)) continue;
                    var v = (kv.Value ?? string.Empty).Trim();
                    dict[k] = v;
                }

                // 即使没有有效项，也写入“空列表XML”，而不是删除文件
                var list = dict.Select(kv => new PipeAttrEntry { Key = kv.Key, Value = kv.Value }).ToList();

                // 序列化器
                var xs = new XmlSerializer(typeof(List<PipeAttrEntry>));

                // 临时文件路径（先写临时，再替换，避免中间态损坏）
                var tmpFile = path + ".tmp";

                // 写入设置（UTF8+BOM）
                var settings = new System.Xml.XmlWriterSettings
                {
                    Indent = true,
                    Encoding = new System.Text.UTF8Encoding(true),
                    NewLineChars = "\r\n"
                };

                // 写入临时文件
                using (var ofs = new FileStream(tmpFile, FileMode.Create, FileAccess.Write, FileShare.None))
                using (var xw = System.Xml.XmlWriter.Create(ofs, settings))
                {
                    xs.Serialize(xw, list);
                    xw.Flush();
                    ofs.Flush(true);
                }

                // 原子替换目标文件
                try
                {
                    if (File.Exists(path))
                        File.Replace(tmpFile, path, null);
                    else
                        File.Move(tmpFile, path);
                }
                catch
                {
                    // 回退覆盖策略
                    if (File.Exists(tmpFile))
                    {
                        File.Copy(tmpFile, path, true);
                        File.Delete(tmpFile);
                    }
                }

                System.Diagnostics.Debug.WriteLine($"[FileManager] Saved {dict.Count} pipe attrs to {path}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[FileManager] SaveLastPipeAttributes failed: {ex.Message}");
            }
        }
      
        /// <summary>
        /// 解析图元主文件在服务器侧的权威路径。
        /// </summary>
        public async Task<string> ResolveServerGraphicPathAsync(DatabaseManager databaseManager, FileStorage storage, string operationName)
        {
            if (storage == null)
            {
                throw new ArgumentNullException(nameof(storage));
            }
            // 获取图元存储根路径
            string root = await ResolveStorageRootPathAsync(databaseManager, operationName).ConfigureAwait(false);
            string rootFull = Path.GetFullPath(root);

            string sourcePath = storage.FilePath ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(sourcePath))
            {
                try
                {
                    string full = Path.GetFullPath(sourcePath);

                    // 1) 若旧记录路径就在当前根目录下，直接使用
                    if (IsPathUnderRoot(full, rootFull))
                    {
                        return full;
                    }

                    // 2) 若旧记录路径不在当前根目录，但其目录可达（配置变更/历史数据场景），也优先复用
                    // 这样可避免“强制切换到新根目录”导致网络共享名不可达（找不到网络名）
                    string existingDir = Path.GetDirectoryName(full) ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(existingDir) && Directory.Exists(existingDir))
                    {
                        return full;
                    }
                }
                catch (Exception exPath)
                {
                    // 路径探测失败时继续走拼接规则，避免直接中断
                    LogManager.Instance.LogWarning($"[{operationName}] 探测历史路径失败，转为按当前根目录拼接: {exPath.Message}");
                }
            }

            if (storage.CategoryId <= 0 || string.IsNullOrWhiteSpace(storage.CategoryType))
            {
                throw new InvalidOperationException("图元缺少分类信息，无法解析服务器存储路径。");
            }
            // 按分类类型、分类ID拼接目录
            string categoryDir = Path.Combine(rootFull, storage.CategoryType, storage.CategoryId.ToString());
            // 按文件名拼接文件名
            string targetName = storage.FileStoredName;
            if (string.IsNullOrWhiteSpace(targetName))
            {
                // 使用原始文件名
                targetName = !string.IsNullOrWhiteSpace(sourcePath)
                    ? Path.GetFileName(sourcePath)
                    : storage.FileName;
            }
            // 文件名不可为空
            if (string.IsNullOrWhiteSpace(targetName))
            {
                throw new InvalidOperationException("图元缺少文件名信息，无法解析服务器文件路径。");
            }
            // 按文件名拼接文件路径
            return Path.Combine(categoryDir, targetName);
        }

        /// <summary>
        /// 通过 HTTP API 替换服务器上的主 DWG 文件
        /// </summary>
        public async Task<FileStorage> ReplaceGraphicFileAsync(DatabaseManager databaseManager, FileStorage storage, string localPath)
        {
            if (storage == null) throw new ArgumentNullException(nameof(storage));
            if (!File.Exists(localPath)) throw new FileNotFoundException("本地替换文件不存在", localPath);

            // 1. 构建 URL
            string serverIp = VariableDictionary._serverIP ?? "127.0.0.1";
            int port = VariableDictionary._apiPort > 0 ? VariableDictionary._apiPort : 10010;
            string url = $"http://{serverIp}:{port}/api/graphics/{storage.Id}/file";

            using var httpClient = new HttpClient();
            using var form = new MultipartFormDataContent();

            // 2. 添加文件流
            var fileStream = new FileStream(localPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            var fileContent = new StreamContent(fileStream);
            fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
            form.Add(fileContent, "dwgFile", Path.GetFileName(localPath));

            // 3. 发送 PUT 请求
            var response = await httpClient.PutAsync(url, form);
            var body = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new HttpRequestException($"替换失败 ({(int)response.StatusCode}): {body}");
            }

            // 4. 解析返回结果，更新本地 FileStorage 对象
            var result = JObject.Parse(body);
            storage.FilePath = result["filePath"]?.Value<string>();
            storage.FileStoredName = result["fileStoredName"]?.Value<string>();
            storage.FileHash = result["fileHash"]?.Value<string>();
            storage.FileSize = result["fileSize"]?.Value<long>() ?? 0;
            storage.FileType = Path.GetExtension(localPath).ToLowerInvariant();
            storage.UpdatedAt = DateTime.Now;

            return storage;
        }

        /// <summary>
        /// 删除服务器侧图元主文件和预览图（可选包含备份文件）。
        /// </summary>
        public async Task<bool> DeletePhysicalFilesAsync(DatabaseManager databaseManager, FileStorage storage, bool deleteBackupFiles)
        {
            if (storage == null)
            {
                throw new ArgumentNullException(nameof(storage));
            }

            string root = await ResolveStorageRootPathAsync(databaseManager, "DeletePhysicalFilesAsync").ConfigureAwait(false);
            string rootFull = Path.GetFullPath(root);

            string graphicPath = await ResolveServerGraphicPathAsync(databaseManager, storage, "DeletePhysicalFilesAsync").ConfigureAwait(false);

            var candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(graphicPath))
            {
                candidates.Add(graphicPath);
            }

            if (!string.IsNullOrWhiteSpace(storage.PreviewImagePath))
            {
                try
                {
                    string previewPath = Path.GetFullPath(storage.PreviewImagePath);
                    if (IsPathUnderRoot(previewPath, rootFull))
                    {
                        candidates.Add(previewPath);
                    }
                }
                catch
                {
                    // 忽略非法路径
                }
            }

            if (!string.IsNullOrWhiteSpace(storage.PreviewImageName))
            {
                string? graphicDir = Path.GetDirectoryName(graphicPath);
                if (!string.IsNullOrWhiteSpace(graphicDir))
                {
                    candidates.Add(Path.Combine(graphicDir, storage.PreviewImageName));
                }
            }

            foreach (string path in candidates)
            {
                if (!File.Exists(path))
                {
                    continue;
                }

                File.Delete(path);

                if (deleteBackupFiles)
                {
                    string bak = path + ".bak";
                    if (File.Exists(bak))
                    {
                        File.Delete(bak);
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// 判断指定路径是否位于指定根目录下。
        /// </summary>
        private static bool IsPathUnderRoot(string fullPath, string rootFullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath) || string.IsNullOrWhiteSpace(rootFullPath))
            {
                return false;
            }

            string normalizedRoot = rootFullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            return fullPath.StartsWith(normalizedRoot, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 计算文件的 SHA-256 哈希值。
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static string ComputeSha256(Stream stream)
        {
            if (stream.CanSeek) stream.Seek(0, SeekOrigin.Begin);
            using var sha256 = SHA256.Create();
            byte[] hash = sha256.ComputeHash(stream);
            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }
    }
}
