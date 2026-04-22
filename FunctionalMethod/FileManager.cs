using System.Data;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Serialization;
using static GB_NewCadPlus_LM.WpfMainWindow;
using DataTable = System.Data.DataTable;
using MessageBox = System.Windows.MessageBox;
using TextBox = System.Windows.Controls.TextBox;
using FileStorage = GB_NewCadPlus_LM.FunctionalMethod.DatabaseManager.FileStorage;
using FileAttribute = GB_NewCadPlus_LM.FunctionalMethod.DatabaseManager.FileAttribute;
using CadCategory = GB_NewCadPlus_LM.FunctionalMethod.DatabaseManager.CadCategory;
using CadSubcategory = GB_NewCadPlus_LM.FunctionalMethod.DatabaseManager.CadSubcategory;

namespace GB_NewCadPlus_LM.FunctionalMethod
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
        /// 当前选择的文件属性
        /// </summary>
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
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GB_NewCadPlus_LM");
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
            var path = GetPipeAttrSavePath(isOutlet);
            if (!File.Exists(path))
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var xs = new XmlSerializer(typeof(List<PipeAttrEntry>));
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var obj = xs.Deserialize(fs) as List<PipeAttrEntry>;
                    if (obj == null) return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

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
            if (attrs == null) return;
            var path = GetPipeAttrSavePath(isOutlet);

            try
            {
                // 规范化并过滤：Trim 键/值，丢弃空键，按不区分大小写合并同名键
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var kv in attrs)
                {
                    var k = (kv.Key ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(k)) continue;
                    var v = (kv.Value ?? string.Empty).Trim();
                    dict[k] = v;
                }

                // 如果没有有效项，则删除旧文件并返回
                if (dict.Count == 0)
                {
                    try { if (File.Exists(path)) File.Delete(path); }
                    catch { }
                    System.Diagnostics.Debug.WriteLine($"[FileManager] Nothing to save for {path}");
                    return;
                }

                var list = dict.Select(kv => new PipeAttrEntry { Key = kv.Key, Value = kv.Value }).ToList();

                var xs = new XmlSerializer(typeof(List<PipeAttrEntry>));
                var tmpFile = path + ".tmp";

                var settings = new System.Xml.XmlWriterSettings
                {
                    Indent = true,
                    Encoding = new System.Text.UTF8Encoding(true), // 带 BOM，符合编辑器/Windows 习惯
                    NewLineChars = "\r\n"
                };

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
                    // 回退：简单覆盖
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
        /// 上传文件并保存到数据库
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
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
            ItemsControl categoryPropertiesDataGrid, // 新增参数
            WpfMainWindow wpfMainWindow // 新增参数
        )
        {
            List<string> uploadedFiles = new List<string>(); // 记录已上传的文件路径，用于回滚
            FileStorage? savedFileStorage = null; // 记录已保存的文件记录
            FileAttribute? savedFileAttribute = null; // 记录已保存的属性记录
            bool transactionSuccess = false;

            try
            {
                if (string.IsNullOrEmpty(_selectedFilePath) || _selectedCategoryNode == null)
                {
                    throw new Exception("文件路径或分类节点为空");
                }

                // 1. 获取文件信息
                var fileInfo = new FileInfo(_selectedFilePath);
                string fileName = fileInfo.Name;
                string displayName = Path.GetFileNameWithoutExtension(fileName);
                string description = $"上传文件: {fileName}";
                var fileStorage = new FileStorage();
                // 2. 使用FileManager上传主文件到服务器指定路径
                using (var fileStream = File.OpenRead(_selectedFilePath))
                {
                    fileStorage = await UploadFileAsync(_databaseManager,
                        categoryId,
                        _selectedCategoryNode.Level == 0 ? "main" : "sub",
                        fileName,
                        fileStream,
                        description,
                        Environment.UserName
                    );

                    // 保存上传后的文件信息
                    _currentFileStorage = fileStorage;
                    savedFileStorage = fileStorage;
                    uploadedFiles.Add(fileStorage.FilePath); // 记录已上传的文件路径
                }

                // 3. 如果有预览图片，上传预览图片
                string previewStoredPath = null;
                if (!string.IsNullOrEmpty(_selectedPreviewImagePath) && File.Exists(_selectedPreviewImagePath))
                {
                    var previewInfo = new FileInfo(_selectedPreviewImagePath);
                    string previewFileName = $"{Path.GetFileNameWithoutExtension(_selectedPreviewImagePath)}_preview{previewInfo.Extension}";

                    using (var previewStream = File.OpenRead(_selectedPreviewImagePath))
                    {
                        // 生成预览文件存储路径
                        string previewStoredName = $"{Guid.NewGuid()}{previewInfo.Extension}";
                        previewStoredPath = Path.Combine(
                            Path.GetDirectoryName(_currentFileStorage.FilePath),
                            previewStoredName);

                        // 复制预览图片到同一目录
                        File.Copy(_selectedPreviewImagePath, previewStoredPath, true);

                        _currentFileStorage.PreviewImageName = previewStoredName;
                        _currentFileStorage.PreviewImagePath = previewStoredPath;
                        uploadedFiles.Add(previewStoredPath); // 记录预览文件路径
                    }
                }

                // 4. 创建文件属性对象
                _currentFileAttribute = new FileAttribute
                {
                    //FileStorageId = _currentFileStorage.Id,
                    FileName = _currentFileStorage.FileName,
                    CreatedAt = DateTime.Now,
                    UpdatedAt = DateTime.Now
                };

                // 5. 从属性编辑网格中获取属性值
                var gridProperties = categoryPropertiesDataGrid.ItemsSource as List<CategoryPropertyEditModel>;
                if (gridProperties != null)
                {
                    foreach (var property in gridProperties)
                    {
                        // 使用实例调用
                        wpfMainWindow.SetFileAttributeProperty(_currentFileAttribute, property.PropertyName1, property.PropertyValue1);
                        wpfMainWindow.SetFileAttributeProperty(_currentFileAttribute, property.PropertyName2, property.PropertyValue2);
                    }
                }
                if (_currentFileAttribute.FileName == null)
                {
                    MessageBox.Show("请填写文件名称", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 6. 保存文件属性到数据库
                int attributeResult = await _databaseManager.AddFileAttributeAsync(_currentFileAttribute);
                if (attributeResult <= 0)
                {
                    LogManager.Instance.LogInfo("保存文件属性失败");

                }
                else
                {
                    LogManager.Instance.LogInfo("保存文件属性到数据库:成功");
                }
                //获取文件属性ID
                _currentFileAttribute = await _databaseManager.GetFileAttributeAsync(_currentFileStorage.DisplayName);
                if (_currentFileAttribute == null || _currentFileAttribute.Id == null)
                {
                    LogManager.Instance.LogInfo("获取文件属性ID失败");
                    // 发生异常，需要回滚操作
                    await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, _currentFileStorage, _currentFileAttribute);
                    return;
                }
                _currentFileStorage.FileAttributeId = _currentFileAttribute.Id;

                //新加文件到数据库中
                var fileResult = await _databaseManager.AddFileStorageAsync(_currentFileStorage);
                if (fileResult == 0)
                {
                    LogManager.Instance.LogInfo("保存文件记录到数据库:失败");
                    // 发生异常，需要回滚操作
                    await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, _currentFileStorage, _currentFileAttribute);
                    return;
                }
                else
                {
                    LogManager.Instance.LogInfo("保存文件记录到数据库:成功");
                }
                ;
                _currentFileStorage = await _databaseManager.GetFileStorageAsync(_currentFileStorage.FileHash);//获取文件的基本信息
                _currentFileAttribute.FileStorageId = _currentFileStorage.Id;//文件属性ID

                await _databaseManager.UpdateFileAttributeAsync(_currentFileAttribute);//更新文件属性
                // 8. 处理标签信息
                await ProcessFileTags(_currentFileStorage.Id, properties);

                // 9. 更新分类统计
                var updateBool = await _databaseManager.UpdateCategoryStatisticsAsync(
                    _currentFileStorage.CategoryId,
                    _currentFileStorage.CategoryType);

                // 如果所有操作都成功，标记事务成功
                transactionSuccess = true;
                // 11. 刷新分类树和界面显示
                // 替换为：
                //await RefreshCurrentCategoryFilesAsync();
                await wpfMainWindow.RefreshCurrentCategoryDisplayAsync(_selectedCategoryNode);
                MessageBox.Show($"文件已成功上传并保存到服务器指定路径\n文件路径: {_currentFileStorage.FilePath}",
                    "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                // 发生异常，需要回滚操作
                await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, _currentFileStorage, _currentFileAttribute);
                throw new Exception($"文件上传和数据库保存失败: {ex.Message}", ex);
            }
            finally
            {
                // 如果事务失败，执行回滚
                if (!transactionSuccess)
                {
                    await FileManager.RollbackFileUpload(_databaseManager, uploadedFiles, _currentFileStorage, _currentFileAttribute);
                }
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
                            TagName = property.PropertyValue1,
                            CreatedAt = DateTime.Now
                        };
                        // 这里需要在DatabaseManager中添加添加标签的方法
                        var addFileTagBool = await _databaseManager.AddFileTagAsync(tag);
                        if (addFileTagBool)
                        {
                            LogManager.Instance.LogInfo($"添加标签 {tag.TagName} 成功");
                        }

                    }

                    // 处理标签2
                    if (property.PropertyName2?.StartsWith("标签") == true && !string.IsNullOrEmpty(property.PropertyValue2))
                    {
                        var tag = new FileTag
                        {
                            FileId = fileId,
                            TagName = property.PropertyValue2,
                            CreatedAt = DateTime.Now
                        };
                        var addFileTagBool = await _databaseManager.AddFileTagAsync(tag);
                        if (addFileTagBool)
                        {
                            LogManager.Instance.LogInfo($"添加标签 {tag.TagName} 成功");
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
        /// 批量导入图元
        /// </summary>
        public async Task BatchImportGraphicsAsync(string excelFilePath , CategoryTreeNode _selectedCategoryNode,ItemsControl _categoryTreeView, List<CategoryTreeNode> _categoryTreeNodes)
        {
            try
            {
                LogManager.Instance.LogInfo($"开始批量导入图元: {excelFilePath}");

                // 读取Excel文件
                DataTable dataTable = ReadExcelToDataTable(excelFilePath);

                if (dataTable == null || dataTable.Rows.Count == 0)
                {
                    MessageBox.Show("Excel文件中没有数据", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int successCount = 0;
                int failCount = 0;

                // 遍历每一行数据
                foreach (DataRow row in dataTable.Rows)
                {
                    try
                    {
                        // 创建FileStorage对象
                        var fileStorage = CreateFileStorageFromRow(row);

                        if (fileStorage != null)
                        {
                            #region 原有代码
                            // 保存到数据库
                            //int fileId = await _databaseManager.AddFileStorageAsync(fileStorage);

                            //if (fileId > 0)
                            //{
                            //    // 创建FileAttribute对象
                            //    var fileAttribute = CreateFileAttributeFromRow(row, fileId);

                            //    if (fileAttribute != null)
                            //    {
                            //        // 保存文件属性
                            //        await _databaseManager.AddFileAttributeAsync(fileAttribute);
                            //    }

                            //    successCount++;
                            //    LogManager.Instance.LogInfo($"成功导入图元: {fileStorage.DisplayName}");
                            //}
                            //else
                            //{
                            //    failCount++;
                            //    LogManager.Instance.LogWarning($"导入图元失败: {fileStorage?.DisplayName}");
                            //}
                            #endregion
                            // 再插入 attribute，并检查 attribute 返回值；仅两者都成功时记为成功。
                            int fileId = await _databaseManager.AddFileStorageAsync(fileStorage);

                            if (fileId > 0)
                            {
                                var fileAttribute = CreateFileAttributeFromRow(row, fileId);

                                int attrId = 0;
                                if (fileAttribute != null)
                                {
                                    attrId = await _databaseManager.AddFileAttributeAsync(fileAttribute);
                                }

                                if (attrId > 0)
                                {
                                    successCount++;
                                    LogManager.Instance.LogInfo($"成功导入图元: {fileStorage.DisplayName} (Id={fileId}, AttrId={attrId})");
                                }
                                else
                                {
                                    // 属性插入失败：标记失败并记录日志。可选：删除已插入的 storage（避免孤立数据）
                                    failCount++;
                                    LogManager.Instance.LogWarning($"导入图元失败（属性写入失败）: {fileStorage?.DisplayName} (StorageId={fileId})");
                                    try
                                    {
                                        // 尝试清理孤立 storage
                                        if (fileId > 0)
                                        {
                                            await _databaseManager.DeleteFileStorageAsync(fileId);
                                            LogManager.Instance.LogInfo($"已删除孤立的 cad_file_storage Id={fileId}");
                                        }
                                    }
                                    catch (Exception cleanupEx)
                                    {
                                        LogManager.Instance.LogInfo($"清理孤立 cad_file_storage 失败: {cleanupEx.Message}");
                                    }
                                }
                            }
                            else
                            {
                                failCount++;
                                LogManager.Instance.LogWarning($"导入图元失败: {fileStorage?.DisplayName}");
                            }
                        }
                        else
                        {
                            failCount++;
                            LogManager.Instance.LogWarning("创建FileStorage对象失败");
                        }
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        LogManager.Instance.LogError($"导入单个图元时出错: {ex.Message}");
                    }
                }

                // 显示结果
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
                    /*
                        CategoryId("分类ID", typeof(int));
                        CategoryType("分类类型", typeof(string));
                       FileName ("文件名", typeof(string));
                       DisplayName ("显示名称", typeof(string));
                        FilePath("文件路径", typeof(string));
                        FileType("文件类型", typeof(string));
                        FileSize("文件大小", typeof(long));
                        ElementBlockName("元素块名", typeof(string));
                        LayerName("图层名称", typeof(string));
                        ColorIndex("颜色索引", typeof(int));
                        PreviewImageName("预览图片名称", typeof(string));
                        PreviewImagePath("预览图片路径", typeof(string));
                        IsPreview("是否预览", typeof(int));
                        CreatedBy("创建者", typeof(string));
                        Title("标题", typeof(string));
                        Keywords("关键字", typeof(string));
                        UpdatedBy("更新者", typeof(string));
                        Version("版本号", typeof(int));
                        IsActive("是否激活", typeof(int));
                        IsPublic("是否公开", typeof(int));
                        Description("描述", typeof(string));
       
                     */
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
        /// 从Excel行数据创建FileAttribute对象
        /// </summary>
        private FileAttribute CreateFileAttributeFromRow(DataRow row, int storageFileId)
        {
            try
            {
                var fileAttribute = new FileAttribute
                {
                    /*
                      { "FileStorageId", "存储文件ID" },
                      { "Length", "长度" },
                      { "Width", "宽度" },
                      { "Height", "高度" },
                      { "Angle", "角度" },
                      { "BasePointX", "基点X" },
                      { "BasePointY", "基点Y" },
                      { "BasePointZ", "基点Z" },
                      { "MediumName", "介质" },
                      { "Specifications", "规格" },
                      { "Material", "材质" },
                      { "StandardNumber", "标准编号" },
                      { "Power", "功率" },
                      { "Volume", "容积" },
                      { "Pressure", "压力" },
                      { "Temperature", "温度" },
                      { "Diameter", "直径" },
                      { "OuterDiameter", "外径" },
                      { "InnerDiameter", "内径" },
                      { "Thickness", "厚度" },
                      { "Weight", "重量" },
                      { "Model", "型号" },
                      { "Remarks", "备注" },
                      { "Customize1", "自定义1" },
                      { "Customize2", "自定义2" },
                      { "Customize3", "自定义3" }
                     */
                    FileStorageId = storageFileId,
                    Width = (decimal?)GetDoubleValue(row, "宽度"),
                    Height = (decimal?)GetDoubleValue(row, "高度"),
                    Length = (decimal?)GetDoubleValue(row, "长度"),
                    Angle = (decimal?)GetDoubleValue(row, "角度"),
                    BasePointX = (decimal?)GetDoubleValue(row, "基点X"),
                    BasePointZ = (decimal?)GetDoubleValue(row, "基点Y"),
                    BasePointY = (decimal?)GetDoubleValue(row, "基点Z"),
                    MediumName = GetStringValue(row, "介质"),
                    Specifications = GetStringValue(row, "规格"),
                    Material = GetStringValue(row, "材质"),
                    StandardNumber = GetStringValue(row, "标准编号"),
                    Power = GetStringValue(row, "功率"),
                    Volume = GetStringValue(row, "容积"),
                    Pressure = GetStringValue(row, "压力"),
                    Temperature = GetStringValue(row, "温度"),
                    Diameter = GetStringValue(row, "直径"),
                    OuterDiameter = GetStringValue(row, "外径"),
                    InnerDiameter = GetStringValue(row, "内径"),
                    Thickness = GetStringValue(row, "厚度"),
                    Weight = GetStringValue(row, "重量"),
                    Model = GetStringValue(row, "型号"),
                    Remarks = GetStringValue(row, "备注"),
                    Customize1 = GetStringValue(row, "自定义1"),
                    Customize2 = GetStringValue(row, "自定义2"),
                    Customize3 = GetStringValue(row, "自定义3"),
                    CreatedAt = DateTime.Now,
                    UpdatedAt = DateTime.Now
                };

                return fileAttribute;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"创建FileAttribute对象时出错: {ex.Message}");
                return null;
            }
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
