using GB_NewCadPlus_IV.FunctionalMethod;

namespace GB_NewCadPlus_IV.UniFiedStandards
{
    /// <summary>
    /// 获取路径相关
    /// </summary>
    public static class GetPath
    {
        #region 路径相关 这个路径下可以存一些本程序自己的配置或其它相关的文件（后加入的图库文件），比较方便； 
        /// <summary>
        /// 拿到本app的local的路径，并创建GB_CADPLUS文件夹
        /// </summary>
        public static string AppPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GB_CADPLUS");

        /// <summary>
        /// 引用文件referenceFile文件夹  
        /// </summary>
        public static string referenceFile = Path.Combine(AppPath, "ReferenceFile");

        /// <summary>
        /// 文件路径与名称  resourcesFile
        /// </summary>
        public static string? filePathAndName = null;

        /// <summary>
        /// 缓存存储路径（可由用户指定，优先于默认的 AppData 路径）。如果用户没有指定，则为 null，此时会使用 AppDataPath 作为缓存目录的基础路径。
        /// </summary>
        public static string? _cacheStoragePath;

        /// <summary>
        /// 本地缓存根目录（与 WpfMainWindow.AppPath 保持一致）
        /// </summary>
        public static string AppDataPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "GB_CADPLUS");

        /// <summary>
        /// 预览图缓存目录（AppDataPath 下的子目录）用于存储从数据库下载的预览图，避免重复下载，提高性能
        /// </summary>
        public static string PreviewCachePath
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(_cacheStoragePath))
                    return Path.Combine(_cacheStoragePath, "PreviewCache");
                return Path.Combine(AppDataPath, "PreviewCache");
            }
            set
            {
                if (_cacheStoragePath != value)
                {
                    LogManager.Instance.LogWarning($"DwgCachePath 被修改！旧值={_cacheStoragePath}，新值={value}，堆栈：{Environment.StackTrace}");
                    _cacheStoragePath = value;
                }
            }
        }

        /// <summary>
        /// Dwg文件缓存目录（AppDataPath 下的子目录）用于存储从数据库下载的Dwg文件，避免重复下载，提高性能
        /// </summary>
        public static string DwgCachePath
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(_cacheStoragePath))
                    return Path.Combine(_cacheStoragePath, "DwgCache");
                return Path.Combine(AppDataPath, "DwgCache");
            }
            set
            {
                if (_cacheStoragePath != value)
                {
                    LogManager.Instance.LogWarning($"DwgCachePath 被修改！旧值={_cacheStoragePath}，新值={value}，堆栈：{Environment.StackTrace}");
                    _cacheStoragePath = value;
                }
            }
        }
        #endregion

        /// <summary>
        /// 拿到要插入文件的地址（C:\Users\Administrator\AppData\Local），如果没有文件，就把资源文件复制过去
        /// </summary>
        /// <param name="table">块表与块表记录</param>
        /// <param name="bytes">文件字节</param>
        /// <param name="fileName">文件名</param>
        /// <param name="over">是否覆盖</param>
        /// <returns>返回文件的objectId</returns>
        public static ObjectId GetBlockFormA(this SymbolTable<BlockTable, BlockTableRecord> table, byte[] bytes, string fileName, bool over)
        {
            if (!Directory.Exists(referenceFile)) //如果不存在这个文件夹，我们就创建这个文件夹
                Directory.CreateDirectory(referenceFile);
            filePathAndName = Path.Combine(referenceFile, fileName + ".dwg");//获得引用文件全路径与文件名
            if (!File.Exists(filePathAndName))
                File.WriteAllBytes(filePathAndName, bytes);

            return table.GetBlockFrom(filePathAndName, over);
        }
        /// <summary>
        /// 拿到要插入文件的地址（C:\Users\Administrator\AppData\Local），如果没有文件，就把资源文件复制过去
        /// </summary>
        /// <param name="table">块表与块表记录</param>
        /// <param name="bytes">文件字节</param>
        /// <param name="fileName">文件名</param>
        /// <returns>返回文件的objectId</returns>
        public static void GetFileFormA(byte[] bytes, string fileName)
        {
            if (!Directory.Exists(referenceFile)) //如果不存在这个文件夹，我们就创建这个文件夹
                Directory.CreateDirectory(referenceFile);
            filePathAndName = Path.Combine(referenceFile, fileName + ".dwg");//获得引用文件全路径与文件名
            if (!File.Exists(filePathAndName))
                File.WriteAllBytes(filePathAndName, bytes);
        }
        /// <summary>
        /// 拿到要插入文件的地址（C:\Users\Administrator\AppData\Local），如果没有文件，就把资源文件复制过去
        /// </summary>
        /// <param name="table">块表与块表记录</param>
        /// <param name="bytes">文件字节</param>
        /// <param name="fileName">文件名</param>
        /// <param name="blockName">块名</param>
        ///  /// <param name="over">是否覆盖</param>
        /// <returns>返回文件的objectId,块名</returns>
        public static ObjectId GetBlockFormA(this SymbolTable<BlockTable, BlockTableRecord> table, byte[] bytes, string fileName, string blockName, bool over)
        {
            if (!Directory.Exists(referenceFile)) //如果不存在这个文件夹，我们就创建这个文件夹
                Directory.CreateDirectory(referenceFile);
            var filePathAndName = Path.Combine(referenceFile, fileName + ".dwg");//获得引用文件全路径与文件名
            if (!File.Exists(filePathAndName))
                File.WriteAllBytes(filePathAndName, bytes);

            return table.GetBlockFrom(filePathAndName, blockName, over);
        }
        /// <summary>
        /// 拿到要插入文件的地址（C:\Users\Administrator\AppData\Local），如果没有文件，就把资源文件复制过去
        /// </summary>
        /// <param name="table">块表与块表记录</param>
        /// <param name="filePathAndName">文件路径与文件名称</param>
        /// <param name="blockName">块名</param>
        ///  /// <param name="over">是否覆盖</param>
        /// <returns>返回文件的objectId,块名</returns>
        public static ObjectId GetBlockFormA(this SymbolTable<BlockTable, BlockTableRecord> table, string filePathAndName, string blockName, bool over)
        {
            try
            {
                return table.GetBlockFrom(filePathAndName, blockName, over);
            }
            catch (Exception ex)
            {
                Env.Editor.WriteMessage($"获取块时出错: {ex.Message}");
                return table.GetBlockFrom(filePathAndName, blockName, over);
            }
            finally
            {
                table = null;
            }
        }

        /// <summary>
        /// 设置一个存储dwg文件的list类型变量ListDwgFile
        /// </summary>
        public static List<string> ListDwgFile { get; set; }//设置一个存储dwg文件的list类型变量ListDwgFile
    }
}
