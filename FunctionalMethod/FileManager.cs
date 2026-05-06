using System.Data;
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
        private readonly DatabaseManager _databaseManager;
        /// <summary>
        /// 基础存储路径
        /// </summary>
        private readonly string _baseStoragePath;
        /// <summary>
        /// 是否使用D盘
        /// </summary>
        private readonly bool _useDPath;
        // 新方案主字段——当前选中的 JSON 属性字典
        private Dictionary<string, string> _selectedAttributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // 兼容字段（旧代码还可能用到），逐步退役
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
        /// 下载文件
        /// </summary>
        public async Task<Stream> DownloadFileAsync(int fileId, string userName, string ipAddress)
        {
            try
            {
                // 获取文件信息
                var file = await _databaseManager.GetFileByIdAsync(fileId);
                if (file == null)
                {
                    throw new Exception("文件不存在或已被删除");
                }

                // 检查文件是否存在
                if (!File.Exists(file.FilePath))
                {
                    throw new Exception("文件在磁盘上不存在");
                }

                // 记录访问日志
                var accessLog = new FileAccessLog
                {
                    FileId = fileId,
                    UserName = userName,
                    ActionType = "Download",
                    AccessTime = DateTime.Now,
                    IpAddress = ipAddress
                };
                await _databaseManager.AddFileAccessLogAsync(accessLog);

                // 返回文件流
                return File.OpenRead(file.FilePath);
            }
            catch (Exception ex)
            {
                throw new Exception($"下载文件失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 获取分类下的所有文件
        /// </summary>
        /// <param name="categoryId">分类ID</param>
        /// <param name="categoryType">分类类型</param>
        /// <returns></returns>
        public async Task<List<FileStorage>> GetFilesByCategoryAsync(int categoryId, string categoryType)
        {
            return await _databaseManager.GetFilesByCategoryIdAsync(categoryId, categoryType);
        }

       /// <summary>
       /// 删除文件
       /// </summary>
       /// <param name="fileId">文件ID</param>
       /// <param name="deletedBy">删除者</param>
       /// <returns>是否删除成功</returns>
       /// <exception cref="Exception"></exception>
        public async Task<bool> DeleteFileAsync(int fileId, string deletedBy)
        {
            try
            {
                // 获取文件信息
                var file = await _databaseManager.GetFileByIdAsync(fileId);
                if (file == null)
                {
                    return false;
                }

                // 从数据库中软删除
                int result = await _databaseManager.DeleteFileAsync(fileId, deletedBy);

                // 可选：从磁盘删除文件（根据业务需求决定）
                // if (File.Exists(file.FilePath))
                // {
                //     File.Delete(file.FilePath);
                // }

                return result > 0;
            }
            catch (Exception ex)
            {
                throw new Exception($"删除文件失败: {ex.Message}", ex);
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
        /// 判断字符串是否为空
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static bool IsConsideredEmpty(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return true;
            var t = s.Trim();
            if (t == "-" || t == "—") return true;
            var tl = t.ToLowerInvariant();
            if (tl == "n/a" || tl == "na" || tl == "无" || tl == "0") return true;
            return false;
        }
        /// <summary>
        /// 合并管道属性
        /// </summary>
        /// <param name="sampleAttrMap">示例属性映射</param>
        /// <param name="savedAttrs">保存的属性</param>
        /// <returns></returns>
        public static Dictionary<string, string> MergeSavedPipeAttributes(Dictionary<string, string>? sampleAttrMap, Dictionary<string, string>? savedAttrs)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            sampleAttrMap = sampleAttrMap ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            savedAttrs = savedAttrs ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 合并键集合（保持不区分大小写）
            var keys = new HashSet<string>(sampleAttrMap.Keys, StringComparer.OrdinalIgnoreCase);
            foreach (var k in savedAttrs.Keys) keys.Add(k);

            foreach (var key in keys)
            {
                sampleAttrMap.TryGetValue(key, out var sampleVal);
                savedAttrs.TryGetValue(key, out var savedVal);

                // 优先使用示例中已有且不是占位/空的值；否则使用保存值（若有），否则空字符串
                if (!IsConsideredEmpty(sampleVal))
                    result[key] = sampleVal!;
                else
                    result[key] = savedVal ?? string.Empty;
            }

            return result;
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
        /// 删除已上传的文件（用于回滚操作）
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>是否删除成功</returns>
        public bool DeleteUploadedFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"删除文件失败: {ex.Message}");
                return false;
            }
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
                if (fileAttribute != null)
                {
                    try
                    {
                        // 删除属性记录
                        await databaseManager.DeleteFileAttributeAsync(fileAttribute.Id);
                        System.Diagnostics.Debug.WriteLine("已删除文件属性记录");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"删除属性记录失败: {ex.Message}");
                    }
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
                // 确保存储路径存在
                EnsureBaseStoragePathExists();

                // 生成唯一的存储文件名
                string fileExtension = Path.GetExtension(originalFileName);
                string storedFileName = $"{Guid.NewGuid()}{fileExtension}";

                // 确定存储路径（按分类类型和ID组织文件夹）
                string categoryPath = Path.Combine(_baseStoragePath, categoryType, categoryId.ToString());
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
        /// 确保存储基础路径存在
        /// </summary>
        private void EnsureBaseStoragePathExists()
        {
            if (!Directory.Exists(_baseStoragePath))
            {
                Directory.CreateDirectory(_baseStoragePath);
            }
        }


        /// <summary>
        /// 更新选中文件
        /// </summary>
        /// <returns></returns>
        public async Task UpdateSelectedFileAsync(FileStorage _selectedFileStorage, System.Windows.Controls.TextBox view_File_Path, TextBox file_Path)
        {
            try
            {
                if (_selectedFileStorage == null || _databaseManager == null)
                    return;

                bool fileUpdated = false;
                bool previewUpdated = false;

                // 更新文件
                if (!string.IsNullOrEmpty(file_Path.Text) && File.Exists(file_Path.Text))
                {
                    // 复制新文件到存储位置
                    string newStoredFileName = $"{Guid.NewGuid()}{Path.GetExtension(file_Path.Text)}";
                    string newStoredFilePath = Path.Combine(
                        Path.GetDirectoryName(_selectedFileStorage.FilePath),
                        newStoredFileName);

                    File.Copy(file_Path.Text, newStoredFilePath, true);

                    // 更新数据库记录
                    _selectedFileStorage.FilePath = newStoredFilePath;
                    _selectedFileStorage.FileName = Path.GetFileName(file_Path.Text);
                    _selectedFileStorage.FileSize = new FileInfo(newStoredFilePath).Length;
                    _selectedFileStorage.Version += 1; // 增加版本号
                    _selectedFileStorage.UpdatedAt = DateTime.Now;

                    fileUpdated = true;
                }

                // 更新预览图片
                if (!string.IsNullOrEmpty(view_File_Path.Text) && File.Exists(view_File_Path.Text))
                {
                    // 复制新预览图片到存储位置
                    string newPreviewFileName = $"{Guid.NewGuid()}{Path.GetExtension(view_File_Path.Text)}";
                    string newPreviewFilePath = Path.Combine(
                        Path.GetDirectoryName(_selectedFileStorage.PreviewImagePath ?? _selectedFileStorage.FilePath),
                        newPreviewFileName);

                    File.Copy(view_File_Path.Text, newPreviewFilePath, true);

                    // 更新数据库记录
                    _selectedFileStorage.PreviewImagePath = newPreviewFilePath;
                    _selectedFileStorage.PreviewImageName = newPreviewFileName;

                    previewUpdated = true;
                }

                // 保存到数据库
                if (fileUpdated || previewUpdated)
                {
                    await _databaseManager.UpdateFileStorageAsync(_selectedFileStorage);
                }

                // 清空输入框
                //new_File_Path.Text = "";
                //new_Preview_Path.Text = "";
                //version_Description.Text = "";

                LogManager.Instance.LogInfo($"文件更新完成: 文件={fileUpdated}, 预览={previewUpdated}");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"更新文件时出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 更新文件属性
        /// </summary>
        /// <param name="fileAttribute"></param>
        /// <param name="propertyName"></param>
        /// <param name="propertyValue"></param>
        private void UpdateFileAttributeProperty(FileAttribute fileAttribute, string propertyName, string propertyValue)
        {
            if (string.IsNullOrEmpty(propertyName) || fileAttribute == null)
                return;

            try
            {
                var property = fileAttribute.GetType().GetProperty(propertyName);
                if (property != null && property.CanWrite)
                {
                    // 根据属性类型进行转换
                    if (property.PropertyType == typeof(string))
                    {
                        property.SetValue(fileAttribute, propertyValue);
                    }
                    else if (property.PropertyType == typeof(int?) || property.PropertyType == typeof(int))
                    {
                        if (int.TryParse(propertyValue, out int intValue))
                        {
                            property.SetValue(fileAttribute, intValue);
                        }
                    }
                    else if (property.PropertyType == typeof(double?) || property.PropertyType == typeof(double))
                    {
                        if (double.TryParse(propertyValue, out double doubleValue))
                        {
                            property.SetValue(fileAttribute, doubleValue);
                        }
                    }
                    else if (property.PropertyType == typeof(DateTime?) || property.PropertyType == typeof(DateTime))
                    {
                        if (DateTime.TryParse(propertyValue, out DateTime dateTimeValue))
                        {
                            property.SetValue(fileAttribute, dateTimeValue);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"更新属性 {propertyName} 时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理文件标签
        /// </summary>
        /// <param name="fileId"></param>
        /// <param name="properties"></param>
        /// <returns></returns>
        public async Task ProcessFileTags(int fileId, List<CategoryPropertyEditModel> properties)
        {
            try
            {
                // 查找标签属性并添加到数据库
                foreach (var property in properties)
                {
                    // 处理标签1
                    if (property.PropertyName1?.StartsWith("标签") == true && !string.IsNullOrEmpty(property.PropertyValue1))
                    {
                        var tag = new FileTag
                        {
                            FileId = fileId,
                            Tag = property.PropertyValue1,
                            CreatedAt = DateTime.Now
                        };
                        // 调用 DatabaseManager 的 AddFileTagAsync，传入顶级模型 FileTag
                        var addFileTagBool = await _databaseManager.AddFileTagAsync(tag);
                        if (addFileTagBool)
                        {
                            LogManager.Instance.LogInfo($"添加标签 {tag.Tag} 成功");
                        }

                    }

                    // 处理标签2
                    if (property.PropertyName2?.StartsWith("标签") == true && !string.IsNullOrEmpty(property.PropertyValue2))
                    {
                        var tag = new FileTag
                        {
                            FileId = fileId,
                            Tag = property.PropertyValue2,
                            CreatedAt = DateTime.Now
                        };
                        var addFileTagBool = await _databaseManager.AddFileTagAsync(tag);
                        if (addFileTagBool)
                        {
                            LogManager.Instance.LogInfo($"添加标签 {tag.Tag} 成功");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"处理文件标签时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 上传文件并保存到数据库（JSON属性新方案）
        /// </summary>
        public async Task UploadFileAndSaveToDatabase(
            string _selectedFilePath,
            FileStorage _currentFileStorage,
            FileAttribute _currentFileAttribute,
            int categoryId,
            string categoryType,
            string filePath,
            string previewImagePath,
            CategoryTreeNode _selectedCategoryNode,
            string _selectedPreviewImagePath,
            List<CategoryPropertyEditModel> properties,
            string createdBy,
            ItemsControl categoryPropertiesDataGrid,
            WpfMainWindow wpfMainWindow
        )
        {
            // 记录已上传文件路径，用于异常时回滚物理文件
            List<string> uploadedFiles = new List<string>();
            // 记录新插入的主表ID，用于异常时回滚数据库
            int savedStorageId = 0;
            // 标记是否全流程成功
            bool transactionSuccess = false;

            try
            {
                // 参数校验，避免空路径或空分类导致后续异常
                if (string.IsNullOrWhiteSpace(_selectedFilePath) || _selectedCategoryNode == null)
                    throw new Exception("文件路径或分类节点为空");

                // 上传主文件到目标目录并生成 FileStorage 基础对象（仅内存）
                var fileInfo = new FileInfo(_selectedFilePath);
                string fileName = fileInfo.Name;
                string description = $"上传文件: {fileName}";

                using (var fileStream = File.OpenRead(_selectedFilePath))
                {
                    _currentFileStorage = await UploadFileAsync(
                        _databaseManager,
                        categoryId,
                        _selectedCategoryNode.Level == 0 ? "main" : "sub",
                        fileName,
                        fileStream,
                        description,
                        Environment.UserName
                    );

                    // 记录已上传主文件，供回滚使用
                    if (!string.IsNullOrWhiteSpace(_currentFileStorage?.FilePath))
                        uploadedFiles.Add(_currentFileStorage.FilePath);
                }

                // 如果用户选择了预览图，则复制到与主文件同目录
                if (!string.IsNullOrWhiteSpace(_selectedPreviewImagePath) && File.Exists(_selectedPreviewImagePath))
                {
                    var previewInfo = new FileInfo(_selectedPreviewImagePath);
                    string previewStoredName = $"{Guid.NewGuid()}{previewInfo.Extension}";
                    string previewStoredPath = Path.Combine(
                        Path.GetDirectoryName(_currentFileStorage.FilePath) ?? string.Empty,
                        previewStoredName);

                    File.Copy(_selectedPreviewImagePath, previewStoredPath, true);

                    _currentFileStorage.PreviewImageName = previewStoredName;
                    _currentFileStorage.PreviewImagePath = previewStoredPath;

                    // 记录已上传预览图，供回滚使用
                    uploadedFiles.Add(previewStoredPath);
                }

                // 从属性编辑网格构建 JSON 属性字典
                var attrs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                // 局部函数：属性名和值都非空时才写入字典
                void AddAttr(string key, string value)
                {
                    if (string.IsNullOrWhiteSpace(key)) return;
                    if (string.IsNullOrWhiteSpace(value)) return;
                    attrs[key.Trim()] = value.Trim();
                }

                // 把分类属性编辑器中的双列属性写入 JSON 字典
                var gridProperties = categoryPropertiesDataGrid.ItemsSource as List<CategoryPropertyEditModel>;
                if (gridProperties != null)
                {
                    foreach (var p in gridProperties)
                    {
                        AddAttr(p.PropertyName1, p.PropertyValue1);
                        AddAttr(p.PropertyName2, p.PropertyValue2);
                    }
                }

                // 补充基础元数据，便于后续查询和排查
                AddAttr("FileName", _currentFileStorage.FileName);
                AddAttr("DisplayName", _currentFileStorage.DisplayName);
                AddAttr("BlockName", _currentFileStorage.BlockName);
                AddAttr("LayerName", _currentFileStorage.LayerName);
                AddAttr("CreatedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                AddAttr("UpdatedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                // 调用新方法，一次性写入主表 + JSON属性表（事务）
                var (storageId, attrId) = await _databaseManager.AddFileStorageAndAttributesJsonAsync(
                    _currentFileStorage,
                    attrs,
                    "default");

                // 校验插入结果
                if (storageId <= 0 || attrId <= 0)
                    throw new Exception("保存文件与JSON属性失败");

                // 记录已保存主键，后续失败可级联回滚
                savedStorageId = storageId;
                _currentFileStorage.Id = storageId;

                // 回读主记录，确保界面层拿到数据库最新值
                var dbStorage = await _databaseManager.GetFileStorageAsync(_currentFileStorage.FileHash);
                if (dbStorage != null)
                    _currentFileStorage = dbStorage;

                // 处理标签数据（沿用原有逻辑）
                await ProcessFileTags(_currentFileStorage.Id, properties);

                // 刷新分类统计（沿用原有逻辑）
                await _databaseManager.UpdateCategoryStatisticsAsync(
                    _currentFileStorage.CategoryId,
                    _currentFileStorage.CategoryType);

                // 刷新界面显示
                await wpfMainWindow.RefreshCurrentCategoryDisplayAsync(_selectedCategoryNode);

                // 标记成功
                transactionSuccess = true;

                MessageBox.Show(
                    $"文件已成功上传并保存\n文件路径: {_currentFileStorage.FilePath}",
                    "成功",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                // 如果数据库记录已落地，则优先级联回滚数据库+文件
                if (savedStorageId > 0)
                {
                    try
                    {
                        await _databaseManager.DeleteCadGraphicCascadeAsync(savedStorageId, true);
                    }
                    catch (Exception rollbackEx)
                    {
                        LogManager.Instance.LogError($"数据库级联回滚失败: {rollbackEx.Message}");
                    }
                }
                else
                {
                    // 若数据库尚未落地，则仅删除已上传的物理文件
                    foreach (var f in uploadedFiles)
                    {
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(f) && File.Exists(f))
                                File.Delete(f);
                        }
                        catch (Exception delEx)
                        {
                            LogManager.Instance.LogError($"回滚删除文件失败: {delEx.Message}");
                        }
                    }
                }

                throw new Exception($"文件上传和数据库保存失败: {ex.Message}", ex);
            }
            finally
            {
                // 兜底日志，便于排查流程状态
                if (!transactionSuccess)
                    LogManager.Instance.LogWarning("UploadFileAndSaveToDatabase 未成功完成，已执行回滚逻辑。");
            }
        }

        /// <summary>
        /// 批量导入图元（JSON属性新链路）
        /// </summary>
        public async Task BatchImportGraphicsAsync(string excelFilePath, CategoryTreeNode _selectedCategoryNode, ItemsControl _categoryTreeView, List<CategoryTreeNode> _categoryTreeNodes)
        {
            try
            {
                // 记录开始日志
                LogManager.Instance.LogInfo($"开始批量导入图元: {excelFilePath}");

                // 读取Excel数据
                DataTable dataTable = ReadExcelToDataTable(excelFilePath);

                // 空数据直接返回
                if (dataTable == null || dataTable.Rows.Count == 0)
                {
                    MessageBox.Show("Excel文件中没有数据", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 成功失败计数
                int successCount = 0;
                int failCount = 0;

                // 逐行导入
                foreach (DataRow row in dataTable.Rows)
                {
                    try
                    {
                        // 创建主表对象
                        var fileStorage = CreateFileStorageFromRow(row);
                        if (fileStorage == null)
                        {
                            failCount++;
                            LogManager.Instance.LogWarning("创建FileStorage对象失败");
                            continue;
                        }

                        // 创建JSON属性字典
                        var attrs = CreateFileAttributeFromRow(row, 0) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                        // 事务写入主表+JSON属性表
                        var (storageId, attrId) = await _databaseManager.AddFileStorageAndAttributesJsonAsync(fileStorage, attrs, "default");

                        // 结果判断
                        if (storageId > 0 && attrId > 0)
                        {
                            successCount++;
                            LogManager.Instance.LogInfo($"成功导入图元: {fileStorage.DisplayName} (StorageId={storageId}, AttrId={attrId})");
                        }
                        else
                        {
                            failCount++;
                            LogManager.Instance.LogWarning($"导入图元失败: {fileStorage.DisplayName}");
                        }
                    }
                    catch (Exception exRow)
                    {
                        // 单行失败不影响整体
                        failCount++;
                        LogManager.Instance.LogError($"导入单个图元时出错: {exRow.Message}");
                    }
                }

                // 显示导入结果
                MessageBox.Show($"批量导入完成\n成功: {successCount} 个\n失败: {failCount} 个",
                    "完成", MessageBoxButton.OK, MessageBoxImage.Information);

                LogManager.Instance.LogInfo($"批量导入完成 - 成功: {successCount}, 失败: {failCount}");

                // 刷新分类树
                await _categoryManager.RefreshCategoryTreeAsync(_selectedCategoryNode, _categoryTreeView, _categoryTreeNodes, _databaseManager);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"批量导入图元时出错: {ex.Message}");
                MessageBox.Show($"批量导入图元时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 从Excel行数据创建FileStorage对象
        /// </summary>
        private FileStorage CreateFileStorageFromRow(DataRow row)
        {
            try
            {
                var fileStorage = new FileStorage
                {                    
                    CategoryId = GetIntValue(row, "分类ID"),
                    CategoryType = "sub", // 默认为主分类
                    FileName = GetStringValue(row, "文件名"),
                    DisplayName = GetStringValue(row, "显示名称"),
                    FilePath = GetStringValue(row, "文件路径"),
                    FileType = GetStringValue(row, "文件类型"),
                    FileSize = GetLongValue(row, "文件大小"),
                    BlockName = GetStringValue(row, "元素块名"),
                    LayerName = GetStringValue(row, "图层名称"),
                    ColorIndex = GetIntValue(row, "颜色索引"),
                    Scale = GetDoubleValue(row, "比例", 1.0),
                    PreviewImageName = GetStringValue(row, "预览图片名称"),
                    PreviewImagePath = GetStringValue(row, "预览图片路径"),
                    IsPreview = GetIntValue(row, "是否预览", 0),
                    CreatedBy = GetStringValue(row, "创建者"),
                    Title = GetStringValue(row, "标题"),
                    Version = GetIntValue(row, "版本号", 1),
                    IsActive = GetIntValue(row, "是否激活", 1),
                    IsPublic = GetIntValue(row, "是否公开", 1),
                    Description = GetStringValue(row, "描述"),
                    CreatedAt = DateTime.Now,
                    UpdatedAt = DateTime.Now
                };

                return fileStorage;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"创建FileStorage对象时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 从Excel行数据创建 JSON 属性字典（新方案）
        /// </summary>
        private Dictionary<string, string> CreateFileAttributeFromRow(DataRow row, int storageFileId)
        {
            // 始终返回非空字典，避免上层空引用
            var attrs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // 防御式检查，避免传入空行导致异常
                if (row == null || row.Table == null)
                {
                    return attrs;
                }

                // 写入关联主键（字符串形式，便于JSON统一）
                attrs["FileStorageId"] = storageFileId.ToString();

                // 局部函数，字符串有值才写入
                void AddIfNotEmpty(string key, string value)
                {
                    if (!string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(value))
                    {
                        attrs[key.Trim()] = value.Trim();
                    }
                }

                // 局部函数，数值有效才写入，统一用英文小数点
                void AddIfNumber(string key, double value)
                {
                    if (Math.Abs(value) > 1e-12)
                    {
                        attrs[key] = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }
                }

                // 按你原有字段映射写入
                AddIfNumber("Width", GetDoubleValue(row, "宽度"));
                AddIfNumber("Height", GetDoubleValue(row, "高度"));
                AddIfNumber("Length", GetDoubleValue(row, "长度"));
                AddIfNumber("Angle", GetDoubleValue(row, "角度"));
                AddIfNumber("BasePointX", GetDoubleValue(row, "基点X"));
                AddIfNumber("BasePointY", GetDoubleValue(row, "基点Y"));
                AddIfNumber("BasePointZ", GetDoubleValue(row, "基点Z"));

                AddIfNotEmpty("MediumName", GetStringValue(row, "介质"));
                AddIfNotEmpty("Specifications", GetStringValue(row, "规格"));
                AddIfNotEmpty("Material", GetStringValue(row, "材质"));
                AddIfNotEmpty("StandardNumber", GetStringValue(row, "标准编号"));

                AddIfNumber("Power", GetDoubleValue(row, "功率"));
                AddIfNumber("Volume", GetDoubleValue(row, "容积"));
                AddIfNumber("Pressure", GetDoubleValue(row, "压力"));
                AddIfNumber("Temperature", GetDoubleValue(row, "温度"));
                AddIfNumber("Diameter", GetDoubleValue(row, "直径"));
                AddIfNumber("OuterDiameter", GetDoubleValue(row, "外径"));
                AddIfNumber("InnerDiameter", GetDoubleValue(row, "内径"));
                AddIfNumber("Thickness", GetDoubleValue(row, "厚度"));
                AddIfNumber("Weight", GetDoubleValue(row, "重量"));

                AddIfNotEmpty("Model", GetStringValue(row, "型号"));
                AddIfNotEmpty("Remarks", GetStringValue(row, "备注"));
                AddIfNotEmpty("Customize1", GetStringValue(row, "自定义1"));
                AddIfNotEmpty("Customize2", GetStringValue(row, "自定义2"));
                AddIfNotEmpty("Customize3", GetStringValue(row, "自定义3"));

                // 写入时间戳，便于后续追踪
                attrs["CreatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                attrs["UpdatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                return attrs;
            }
            catch (Exception ex)
            {
                // 记录日志并返回当前已收集到的字典（不抛出，防止批量导入中断）
                LogManager.Instance.LogError($"创建JSON属性字典时出错: {ex.Message}");
                return attrs;
            }
        }

        /// <summary>
        /// 将旧 FileAttribute 转为 JSON 字典（过渡桥接）
        /// </summary>
        private Dictionary<string, string> ConvertFileAttributeToDictionary(FileAttribute fileAttribute)
        {
            // 创建不区分大小写字典
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 空对象直接返回空字典
            if (fileAttribute == null) return dict;

            // 通过反射遍历旧模型属性，自动收集非空值
            foreach (var p in typeof(FileAttribute).GetProperties())
            {
                // 只读取可读属性
                if (!p.CanRead) continue;

                // 过滤技术字段（可按需扩展）
                if (string.Equals(p.Name, "Id", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(p.Name, "CreatedAt", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(p.Name, "UpdatedAt", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // 获取属性值
                var v = p.GetValue(fileAttribute);

                // 空值跳过
                if (v == null) continue;

                // 转换字符串
                var s = Convert.ToString(v);

                // 空字符串跳过
                if (string.IsNullOrWhiteSpace(s)) continue;

                // 写入字典
                dict[p.Name] = s.Trim();
            }

            return dict;
        }

        /// <summary>
        /// 读取Excel文件到DataTable
        /// </summary>
        private DataTable ReadExcelToDataTable(string filePath)
        {
            try
            {
                DataTable dataTable = new DataTable();

                // 使用EPPlus读取Excel
                using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // 读取第一个工作表

                    // 读取标题行
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var cellValue = worksheet.Cells[1, col].Value?.ToString() ?? "";
                        dataTable.Columns.Add(cellValue);
                    }

                    // 读取数据行
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var dataRow = dataTable.NewRow();
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            dataRow[col - 1] = worksheet.Cells[row, col].Value ?? DBNull.Value;
                        }
                        dataTable.Rows.Add(dataRow);
                    }
                }

                return dataTable;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"读取Excel文件时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 辅助方法
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="columnName">列</param>
        /// <returns></returns>
        private string GetStringValue(DataRow row, string columnName)
        {
            try
            {
                if (row.Table.Columns.Contains(columnName) && row[columnName] != DBNull.Value)
                {
                    return row[columnName].ToString();
                }
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 获取整型值
        /// </summary>
        /// <param name="row"></param>
        /// <param name="columnName"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        private int GetIntValue(DataRow row, string columnName, int defaultValue = 0)
        {
            try
            {
                if (row.Table.Columns.Contains(columnName) && row[columnName] != DBNull.Value)
                {
                    if (int.TryParse(row[columnName].ToString(), out int result))
                    {
                        return result;
                    }
                }
                return defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// 获取长整型值
        /// </summary>
        /// <param name="row"></param>
        /// <param name="columnName"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        /// 
        private long GetLongValue(DataRow row, string columnName, long defaultValue = 0)
        {
            try
            {
                if (row.Table.Columns.Contains(columnName) && row[columnName] != DBNull.Value)
                {
                    if (long.TryParse(row[columnName].ToString(), out long result))
                    {
                        return result;
                    }
                }
                return defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// 获取双精度值
        /// </summary>
        /// <param name="row"></param>
        /// <param name="columnName"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        private double GetDoubleValue(DataRow row, string columnName, double defaultValue = 0.0)
        {
            try
            {
                if (row.Table.Columns.Contains(columnName) && row[columnName] != DBNull.Value)
                {
                    if (double.TryParse(row[columnName].ToString(), out double result))
                    {
                        return result;
                    }
                }
                return defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// 获取布尔值
        /// </summary>
        /// <param name="row"></param>
        /// <param name="columnName"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        private bool GetBoolValue(DataRow row, string columnName, bool defaultValue = false)
        {
            try
            {
                if (row.Table.Columns.Contains(columnName) && row[columnName] != DBNull.Value)
                {
                    if (bool.TryParse(row[columnName].ToString(), out bool result))
                    {
                        return result;
                    }
                    // 处理"是"/"否"等中文表示
                    string value = row[columnName].ToString().ToLower();
                    if (value == "是" || value == "true" || value == "1")
                        return true;
                    if (value == "否" || value == "false" || value == "0")
                        return false;
                }
                return defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }

    }
}
