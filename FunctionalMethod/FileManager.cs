using System.Data;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Serialization;
using GB_NewCadPlus_IV.UniFiedStandards;
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
        private readonly DatabaseManager _databaseManager;

        /// <summary>
        /// 基础存储路径
        /// </summary>
        private readonly string _baseStoragePath;

        /// <summary>
        /// 是否使用D盘
        /// </summary>
        private readonly bool _useDPath;

        /// <summary>
        /// 新方案主字段——当前选中的 JSON 属性字典
        /// </summary>
        private Dictionary<string, string> _selectedAttributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 兼容字段（旧代码还可能用到），逐步退役
        /// </summary>
        [Obsolete("过渡字段：请逐步改用 _selectedAttributes")]
        private FileAttribute _selectedFileAttribute;
               
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
        public FileManager(DatabaseManager databaseManager, string baseStoragePath = null, bool useDPath = true)
        {
            _databaseManager = databaseManager;/// 数据库管理器
            _useDPath = useDPath;/// 是否使用D盘

            if (!string.IsNullOrEmpty(baseStoragePath))/// 如果提供了基础存储路径，则使用它
            {
                _baseStoragePath = baseStoragePath;/// 使用指定的基础存储路径
            }
            else
            {
                _baseStoragePath = GetBaseStoragePath();/// 否则，智能选择C盘或D盘作为基础存储路径
            }
        }

        /// <summary>
        /// 获取基础存储路径（智能选择C盘或D盘）
        /// </summary>
        private string GetBaseStoragePath()
        {
            // 如果启用D盘优先且D盘存在且可写
            if (_useDPath && Directory.Exists("D:\\") && IsDirectoryWritable("D:\\"))
            {
                return "D:\\GB_Tools\\Cad_Sw_Library";
            }
            else
            {
                // 使用C盘作为备选
                return "C:\\GB_Tools\\Cad_Sw_Library";
            }
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
        /// 将配置路径解析为“服务器可访问路径”。
        /// 规则：
        /// 1) UNC 路径（\\server\share）保持不变；
        /// 2) 盘符路径（D:\\...）会结合服务器 IP 转为 \\IP\d\...（使用普通共享名，不使用 D$）；
        /// 3) 其余路径原样返回。
        /// </summary>
        /// <param name="configuredPath">配置中的路径</param>
        /// <param name="serverIp">登录页服务器 IP</param>
        /// <returns>服务器可访问路径</returns>
        /// <exception cref="InvalidOperationException">盘符路径但服务器 IP 为空时抛出</exception>
        private static string ResolveServerStoragePath(string configuredPath, string serverIp)
        {
            // 先做规范化，统一处理空白、引号、环境变量
            string normalized = NormalizeConfiguredPath(configuredPath);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }

            // 已是 UNC 路径，直接返回
            if (normalized.StartsWith("\\\\", StringComparison.Ordinal))
            {
                return normalized;
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

                // 组合为普通共享 UNC：\\IP\\d\\...
                if (string.IsNullOrWhiteSpace(tailPath))
                {
                    return $"\\\\{ip}\\{driveShareName}";
                }

                return $"\\\\{ip}\\{driveShareName}\\{tailPath}";
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
            string fallbackPath = _baseStoragePath;

            // 读取当前登录服务器 IP（用于将 D:\\... 转换为 \\IP\\D$\\...）
            string serverIp = NormalizeConfiguredPath(VariableDictionary._serverIP);

            // 数据库不可用时，保留旧行为：允许回退本地路径
            if (databaseManager == null || !databaseManager.IsDatabaseAvailable)
            {
                LogManager.Instance.LogWarning($"[{operationName}] 数据库不可用，使用本地回退路径: {fallbackPath}");
                return fallbackPath;
            }

            // 1) 优先读取系统配置 SourceRoot
            string sourceRootRaw = await databaseManager.GetSystemConfigValueAsync("SourceRoot").ConfigureAwait(false);
            string sourceRoot = ResolveServerStoragePath(sourceRootRaw, serverIp);
            if (!string.IsNullOrWhiteSpace(sourceRoot))
            {
                LogManager.Instance.LogInfo($"[{operationName}] 解析到 SourceRoot(服务器路径): {sourceRoot}");
                return sourceRoot;
            }

            // 2) 兼容旧链路：读取运行时变量中的存储路径（即 TextBoxSetStoragePath）
            string runtimeStoragePath = ResolveServerStoragePath(VariableDictionary._storagePath, serverIp);
            if (!string.IsNullOrWhiteSpace(runtimeStoragePath))
            {
                LogManager.Instance.LogInfo($"[{operationName}] SourceRoot 为空，回退到运行时存储路径(服务器路径): {runtimeStoragePath}");
                return runtimeStoragePath;
            }

            // 3) 继续兜底：读取本地设置文件中的 StoragePath（避免运行时变量尚未同步）
            string settingsStoragePath = ResolveServerStoragePath(Properties.Settings.Default.StoragePath, serverIp);
            if (!string.IsNullOrWhiteSpace(settingsStoragePath))
            {
                LogManager.Instance.LogInfo($"[{operationName}] 运行时存储路径为空，回退到设置 StoragePath(服务器路径): {settingsStoragePath}");
                return settingsStoragePath;
            }

            // 4) 最后兜底：使用约定默认路径并转换为服务器共享路径（不是本地写入）
            string defaultServerStoragePath = ResolveServerStoragePath(@"D:\GB_Tools\Cad_Sw_Library", serverIp);
            if (!string.IsNullOrWhiteSpace(defaultServerStoragePath))
            {
                LogManager.Instance.LogWarning($"[{operationName}] SourceRoot/运行时路径/设置路径均为空，使用默认服务器路径: {defaultServerStoragePath}");
                return defaultServerStoragePath;
            }

            // 5) 数据库可用但无可用服务器路径配置时，阻止上传
            throw new InvalidOperationException("系统配置 SourceRoot 与运行时存储路径均为空，无法确定服务器存储路径。请先在设置中配置存储路径，并确认登录页服务器IP正确。");
        }

        /// <summary>
        /// 检查目录是否可写
        /// </summary>
        private bool IsDirectoryWritable(string directoryPath)
        {
            try
            {
                string testFilePath = Path.Combine(directoryPath, "test_write_permission.tmp");/// 测试文件路径
                File.WriteAllText(testFilePath, "test");/// 创建测试文件/ 尝试写入测试文件
                File.Delete(testFilePath);/// 删除测试文件
                return true;
            }
            catch
            {
                return false;/// 如果发生异常，则返回false/ 如果写入失败，返回不可写
            }
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
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GB_NewCadPlus_IV");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                return Path.Combine(dir, isOutlet ? "LastPipeAttrs_Outlet.xml" : "LastPipeAttrs_Inlet.xml");
            }
            catch
            {
                return Path.GetTempFileName(); // 兜底
            }
        }

        /// <summary>
        /// 格式化文件大小
        /// </summary>
        /// <param name="fileSize"></param>
        /// <returns></returns>
        public static string FormatFileSize(long fileSize)
        {
            try
            {
                if (fileSize < 1024)
                    return $"{fileSize} B";
                else if (fileSize < 1024 * 1024)
                    return $"{fileSize / 1024.0:F2} KB";
                else if (fileSize < 1024 * 1024 * 1024)
                    return $"{fileSize / (1024.0 * 1024.0):F2} MB";
                else
                    return $"{fileSize / (1024.0 * 1024.0 * 1024.0):F2} GB";
            }
            catch
            {
                return fileSize.ToString();
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
        /// 计算文件的哈希值
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <returns>文件的哈希值</returns>
        public static async Task<string> CalculateFileHashAsync(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            bool canSeek = stream.CanSeek;
            long originalPosition = canSeek ? stream.Position : 0;
            if (canSeek)
            {
                stream.Position = 0;
            }

            try
            {
                // 在后台线程执行 ComputeHash（同步读取），避免在 UI 线程执行耗时的 CPU/IO 操作
                byte[] hash = await Task.Run(() =>
                {
                    using (var sha = SHA256.Create())
                    {
                        // ComputeHash 会从当前流位置读取到末尾
                        return sha.ComputeHash(stream);
                    }
                }).ConfigureAwait(false);

                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }
            finally
            {
                // 恢复流原始位置（如果流可寻址）
                if (canSeek)
                {
                    stream.Position = originalPosition;
                }
            }
        }

        /// <summary>
        /// 判断文件扩展名是否为预览文件
        /// </summary>
        /// <param name="fileExtension">文件扩展名</param>
        /// <returns>1: 是预览文件, 0: 不是预览文件</returns>
        public static int IsPreviewFile(string fileExtension)
        {
            var previewExtensions = new[] { ".png", ".jpg", ".jpeg", ".bmp", ".gif" };// 预览文件扩展名列表
            if(previewExtensions.Contains(fileExtension.ToLower())) { return 1; }else { return 0; }// 判断文件扩展名是否为预览文件

        }
       
        /// <summary>
        /// 删除文件夹（如果为空）
        /// </summary>
        /// <param name="folderPath">文件夹路径</param>
        /// <returns>是否删除成功</returns>
        public static bool DeleteEmptyFolder(string folderPath)
        {
            try
            {
                if (Directory.Exists(folderPath) && Directory.GetFileSystemEntries(folderPath).Length == 0)
                {
                    Directory.Delete(folderPath);
                }
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"删除空文件夹失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 回滚文件上传操作
        /// </summary>
        /// <param name="uploadedFiles">已上传的文件路径列表</param>
        /// <param name="fileStorage">已保存的文件记录</param>
        /// <param name="fileAttribute">已保存的属性记录</param>
        public static async Task RollbackFileUpload(DatabaseManager databaseManager, List<string> uploadedFiles, FileStorage fileStorage, FileAttribute fileAttribute)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("开始回滚文件上传操作...");

                // 1. 删除已上传的文件
                foreach (string filePath in uploadedFiles)
                {
                    if (File.Exists(filePath))
                    {
                        try
                        {
                            File.Delete(filePath);//删除文件
                            System.Diagnostics.Debug.WriteLine($"已删除文件: {filePath}");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"删除文件失败 {filePath}: {ex.Message}");
                        }
                    }
                }

                // 2. 删除空的文件夹
                if (fileStorage != null)
                {
                    string categoryPath = Path.GetDirectoryName(fileStorage.FilePath);// 获取文件所在的分类文件夹路径
                    if (Directory.Exists(categoryPath))
                    {
                        try
                        {
                            // 尝试删除分类文件夹（如果为空）
                            DeleteEmptyFolder(categoryPath);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"删除文件夹失败: {ex.Message}");
                        }
                    }
                }

                // 3. 如果数据库记录已创建，删除数据库记录
                // 已废弃 DeleteFileAttributeAsync
                if (fileAttribute != null)
                {
                    System.Diagnostics.Debug.WriteLine("[Rollback] 略过旧表属性记录删除");
                }
            

                if (fileStorage != null )
                {
                    try
                    {
                        // 删除文件记录
                        await databaseManager.DeleteFileStorageAsync(fileStorage.Id);
                        System.Diagnostics.Debug.WriteLine("已删除文件存储记录");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"删除文件记录失败: {ex.Message}");
                    }
                }

                System.Diagnostics.Debug.WriteLine("文件上传回滚操作完成");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"回滚操作失败: {ex.Message}");
            }
        }
       
        /// <summary>
        /// 上传文件到服务器指定路径
        /// </summary>
        /// <param name="categoryId">分类ID</param>
        /// <param name="categoryType">分类类型</param>
        /// <param name="originalFileName">原始文件名</param>
        /// <param name="fileStream">文件流</param>
        /// <param name="description">文件描述</param>
        /// <param name="createdBy">创建者</param>
        /// <returns>文件存储信息</returns>
        public async Task<FileStorage> UploadFileAsync(DatabaseManager databaseManager, int categoryId, string categoryType,
            string originalFileName, Stream fileStream,
            string description, string createdBy)
        {
            try
            {
                // 统一解析上传根路径：数据库可用时必须命中服务器配置，避免静默回落本地路径
                string actualStoragePath = await ResolveStorageRootPathAsync(databaseManager, "UploadFileAsync").ConfigureAwait(false);

                // 确保存储路径存在
                if (!Directory.Exists(actualStoragePath))
                {
                    Directory.CreateDirectory(actualStoragePath);
                }

                // 生成唯一的存储文件名
                string fileExtension = Path.GetExtension(originalFileName);
                string storedFileName = $"{Guid.NewGuid()}{fileExtension}";

                // 确定存储路径（按分类类型和ID组织文件夹）
                string categoryPath = Path.Combine(actualStoragePath, categoryType, categoryId.ToString());
                if (!Directory.Exists(categoryPath))
                {
                    Directory.CreateDirectory(categoryPath);
                }

                string fullPath = Path.Combine(categoryPath, storedFileName);

                // 计算文件哈希值（用于去重）
                string fileHash = await CalculateFileHashAsync(fileStream);

                // 保存文件到磁盘
                fileStream.Position = 0;
                using (var fileStreamOutput = File.Create(fullPath))
                {
                    await fileStream.CopyToAsync(fileStreamOutput);
                }

                // 获取文件大小
                long fileSize = new FileInfo(fullPath).Length;

                // 创建文件记录
                var fileRecord = new FileStorage
                {
                    CategoryId = categoryId,
                    CategoryType = categoryType,
                    FileName = originalFileName,
                    FileStoredName = storedFileName,
                    FilePath = fullPath,  // 存储完整路径
                    FileType = fileExtension.ToLower(),
                    FileSize = fileSize,
                    FileHash = fileHash,
                    DisplayName = Path.GetFileNameWithoutExtension(originalFileName),
                    Description = description,
                    Version = 1,
                    IsPreview = IsPreviewFile(fileExtension),
                    CreatedAt = DateTime.Now,
                    UpdatedAt = DateTime.Now,
                    CreatedBy = createdBy,
                    IsActive = 1,
                    IsPublic = 1,
                    Scale = 1.0
                };

                return fileRecord;
            }
            catch (Exception ex)
            {
                throw new Exception($"上传文件失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 获取当前应使用的服务器存储根路径（统一入口）。
        /// </summary>
        public async Task<string> GetServerStorageRootPathAsync(DatabaseManager databaseManager, string operationName)
        {
            return await ResolveStorageRootPathAsync(databaseManager, operationName).ConfigureAwait(false);
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

            string categoryDir = Path.Combine(rootFull, storage.CategoryType, storage.CategoryId.ToString());

            string targetName = storage.FileStoredName;
            if (string.IsNullOrWhiteSpace(targetName))
            {
                targetName = !string.IsNullOrWhiteSpace(sourcePath)
                    ? Path.GetFileName(sourcePath)
                    : storage.FileName;
            }

            if (string.IsNullOrWhiteSpace(targetName))
            {
                throw new InvalidOperationException("图元缺少文件名信息，无法解析服务器文件路径。");
            }

            return Path.Combine(categoryDir, targetName);
        }

        /// <summary>
        /// 替换服务器图元主文件，并回写 FileStorage 关键字段。
        /// </summary>
        public async Task<FileStorage> ReplaceGraphicFileAsync(DatabaseManager databaseManager, FileStorage storage, string localPath)
        {
            if (storage == null)
            {
                throw new ArgumentNullException(nameof(storage));
            }

            if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
            {
                throw new FileNotFoundException("本地替换文件不存在。", localPath);
            }

            string serverPath = await ResolveServerGraphicPathAsync(databaseManager, storage, "ReplaceGraphicFileAsync").ConfigureAwait(false);
            string? serverDir = Path.GetDirectoryName(serverPath);
            if (string.IsNullOrWhiteSpace(serverDir))
            {
                throw new InvalidOperationException("无法解析服务器目标目录。");
            }

            if (!Directory.Exists(serverDir))
            {
                Directory.CreateDirectory(serverDir);
            }

            if (File.Exists(serverPath))
            {
                File.Copy(serverPath, serverPath + ".bak", true);
            }

            File.Copy(localPath, serverPath, true);

            using (FileStream fs = File.OpenRead(serverPath))
            {
                storage.FileHash = await CalculateFileHashAsync(fs).ConfigureAwait(false);
            }

            FileInfo fi = new FileInfo(serverPath);
            storage.FilePath = serverPath;
            storage.FileStoredName = Path.GetFileName(serverPath);
            storage.FileType = Path.GetExtension(serverPath).ToLowerInvariant();
            storage.FileSize = fi.Length;
            storage.UpdatedAt = DateTime.Now;

            return storage;
        }

        /// <summary>
        /// 替换服务器预览图，并回写 FileStorage 关键字段。
        /// </summary>
        public async Task<FileStorage> ReplacePreviewFileAsync(DatabaseManager databaseManager, FileStorage storage, string localPreviewPath)
        {
            if (storage == null)
            {
                throw new ArgumentNullException(nameof(storage));
            }

            if (string.IsNullOrWhiteSpace(localPreviewPath) || !File.Exists(localPreviewPath))
            {
                throw new FileNotFoundException("本地预览图不存在。", localPreviewPath);
            }

            string serverGraphicPath = await ResolveServerGraphicPathAsync(databaseManager, storage, "ReplacePreviewFileAsync").ConfigureAwait(false);
            string? targetDir = Path.GetDirectoryName(serverGraphicPath);
            if (string.IsNullOrWhiteSpace(targetDir))
            {
                throw new InvalidOperationException("无法解析服务器预览图目录。");
            }

            if (!Directory.Exists(targetDir))
            {
                Directory.CreateDirectory(targetDir);
            }

            string ext = Path.GetExtension(localPreviewPath);
            if (string.IsNullOrWhiteSpace(ext))
            {
                ext = ".png";
            }

            string previewName = storage.PreviewImageName;
            if (string.IsNullOrWhiteSpace(previewName))
            {
                previewName = Path.GetFileNameWithoutExtension(storage.FileStoredName ?? storage.FileName ?? Guid.NewGuid().ToString("N")) + ext;
            }
            else if (string.IsNullOrWhiteSpace(Path.GetExtension(previewName)))
            {
                previewName = Path.GetFileNameWithoutExtension(previewName) + ext;
            }

            string targetPreviewPath = Path.Combine(targetDir, previewName);

            if (File.Exists(targetPreviewPath))
            {
                File.Copy(targetPreviewPath, targetPreviewPath + ".bak", true);
            }

            File.Copy(localPreviewPath, targetPreviewPath, true);

            storage.PreviewImagePath = targetPreviewPath;
            storage.PreviewImageName = Path.GetFileName(targetPreviewPath);
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
    }
}
