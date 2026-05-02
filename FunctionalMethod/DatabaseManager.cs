using System.Data;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Serialization;
using Dapper;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static GB_NewCadPlus_IV.WpfMainWindow;
using DataTable = System.Data.DataTable;
using MessageBox = System.Windows.MessageBox;
using TextBox = System.Windows.Controls.TextBox;
using FileStorage = GB_NewCadPlus_IV.FunctionalMethod.DatabaseManager.FileStorage;
// 过渡期保留旧别名，后续全部改完可删除
using FileAttribute = GB_NewCadPlus_IV.FunctionalMethod.DatabaseManager.FileAttribute;
using CadCategory = GB_NewCadPlus_IV.FunctionalMethod.DatabaseManager.CadCategory;
using CadSubcategory = GB_NewCadPlus_IV.FunctionalMethod.DatabaseManager.CadSubcategory;
using FileTag = GB_NewCadPlus_IV.FunctionalMethod.DatabaseManager.FileTag;
using FileAccessLog = GB_NewCadPlus_IV.FunctionalMethod.DatabaseManager.FileAccessLog;
using DeviceInfo = GB_NewCadPlus_IV.UniFiedStandards.DeviceInfo;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 数据库访问类  
    /// </summary>
    public class DatabaseManager
    {
        /// <summary>
        /// 对外公开数据库连接（注意：调用方负责不要忘记关闭/处置）
        /// </summary>
        public MySqlConnection GetConnection()
        {
            return new MySqlConnection(_connectionString);
        }

        // -----------------------------
        // 简易异步方法占位实现（避免引用处编译错误）
        // 说明：这些方法为占位实现，返回默认值或抛出未实现异常。
        // 在接入真实后端时，请替换为完整实现。
        // -----------------------------

        /// <summary>
        /// 文件访问日志模型（占位）
        /// </summary>
        public class FileAccessLog
        {
            public int FileId { get; set; }
            public string? UserName { get; set; }
            public string? ActionType { get; set; }
            public DateTime AccessTime { get; set; }
            public string? IpAddress { get; set; }
        }

        /// <summary>
        /// 文件标签模型（占位）
        /// </summary>
        public class FileTag
        {
            public int FileId { get; set; }
            public string? TagName { get; set; }
            public DateTime CreatedAt { get; set; }
        }

        /// <summary>
        /// 添加文件访问日志（占位）
        /// </summary>
        public virtual async Task<bool> AddFileAccessLogAsync(FileAccessLog log)
        {
            await Task.Yield();
            // 占位：默认记录成功
            return true;
        }

        /// <summary>
        /// 删除文件（占位）
        /// </summary>
        public virtual async Task<int> DeleteFileAsync(int fileId, string deletedBy)
        {
            await Task.Yield();
            // 占位：返回 1 表示已删除
            return 1;
        }

        public virtual async Task<bool> DeleteFileAttributeAsync(long attributeId)
        {
            await Task.Yield();
            return true;
        }

        public virtual async Task<bool> DeleteFileStorageAsync(long storageId)
        {
            await Task.Yield();
            return true;
        }

        public virtual async Task<int> AddFileAttributeAsync(FileAttribute attribute)
        {
            await Task.Yield();
            // 旧链路已废弃，返回0避免“假成功”
            LogManager.Instance.LogWarning("AddFileAttributeAsync 已废弃，请改用 AddFileStorageAndAttributesJsonAsync。");
            return 0;
        }

        public virtual async Task<int> AddFileStorageAsync(FileStorage storage)
        {
            await Task.Yield();
            // 旧链路已废弃，返回0避免“假成功”
            LogManager.Instance.LogWarning("AddFileStorageAsync 旧写入链路已废弃，请改用 AddFileStorageAndAttributesJsonAsync。");
            return 0;
        }


        //public virtual async Task<FileStorage?> GetFileStorageAsync(string fileHash)
        //{
        //    await Task.Yield();
        //    return null;
        //}

        public virtual async Task<bool> UpdateFileAttributeAsync(FileAttribute attribute)
        {
            await Task.Yield();
            return true;
        }

        public virtual async Task<bool> UpdateFileStorageAsync(FileStorage storage)
        {
            await Task.Yield();
            return true;
        }

        public virtual async Task<bool> AddFileTagAsync(FileTag tag)
        {
            await Task.Yield();
            return true;
        }

        public virtual async Task<int> AddFileAccessLogAsync(object accessLog)
        {
            await Task.Yield();
            return 1;
        }


        /// <summary>
        /// 用户实体（对应 users 表）
        /// </summary>
        public class User
        {
            /// <summary>
            /// 用户 ID
            /// </summary>
            public int Id { get; set; }
            /// <summary>
            /// 用户名
            /// </summary>
            public string? Username { get; set; }
            /// <summary>
            /// 密码哈希
            /// </summary>
            public string? PasswordHash { get; set; }
            // <summary>
            /// 显示名称
            /// </summary>
            public string? DisplayName { get; set; }

            /// <summary>
            /// 性别
            /// </summary>
            public string? Gender { get; set; }
            /// <summary>
            /// 手机号码
            /// </summary>
            public string? Phone { get; set; }
            /// <summary>
            /// 电子邮箱
            /// </summary>
            public string? Email { get; set; }
            /// <summary>
            /// 部门 ID
            /// </summary>
            public int? DepartmentId { get; set; }

            /// <summary>
            /// 角色
            /// </summary>
            public string? Role { get; set; }
            /// <summary>
            /// 状态
            /// </summary>
            public int Status { get; set; }
            /// <summary>
            /// 创建时间
            /// </summary>
            public DateTime CreatedAt { get; set; }

            /// <summary>
            /// 更新时间
            /// </summary>
            public DateTime UpdatedAt { get; set; }
        }

        /// <summary>
        /// 根据用户名查询用户（用于注册后获取 id）
        /// </summary>
        /// <param name="username"></param>
        /// <returns>匹配的 User 或 null</returns>
        public async Task<User> GetUserByUsernameAsync(string username)
        {
            if (string.IsNullOrWhiteSpace(username))
                return null;

            const string sql = @"
                SELECT
                    id AS Id,
                    username AS Username,
                    password_hash AS PasswordHash,
                    display_name AS DisplayName,
                    gender AS Gender,
                    phone AS Phone,
                    email AS Email,
                    department_id AS DepartmentId,
                    role AS Role,
                    status AS Status,
                    created_at AS CreatedAt,
                    updated_at AS UpdatedAt
                FROM users
                WHERE username = @Username
                LIMIT 1";

            try
            {
                using var conn = GetConnection();
                var user = await conn.QuerySingleOrDefaultAsync<User>(sql, new { Username = username }).ConfigureAwait(false);
                return user;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetUserByUsernameAsync 出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 检查指定数据库中是否缺少核心表。
        /// 返回：
        /// - 若数据库不存在：返回包含单项 "__DATABASE_MISSING__"
        /// - 若数据库存在但缺少表：返回缺失表名列表（不为空）
        /// - 若一切正常：返回空列表
        /// </summary>
        public static List<string> CheckMissingCoreTables(string server, int port, string user, string password, string database = "cad_sw_library")
        {
            var missing = new List<string>();
            try
            {
                // 连接到目标数据库以检查表
                var connStr = $"Server={server};Port={port};Database={database};Uid={user};Pwd={password};";
                using var conn = new MySqlConnection(connStr);
                conn.Open();

                // 需要保证的核心表（含 CAD / SW / 设备表）
                var required = new[]
                {
                    "cad_categories",
                    "cad_subcategories",
                    "cad_file_storage",
                    "cad_block_attributes_json",
                    "system_config",
                    "users",
                    "departments",
                    "department_users",
                    "category_department_map",
                    "sw_categories",
                    "sw_subcategories",
                    "sw_graphics",
                    "device_info"
                };

                var sql = @"SELECT table_name FROM information_schema.tables
                        WHERE table_schema = @schema AND table_name IN @names";
                var found = conn.Query<string>(sql, new { schema = database, names = required }).AsList();

                foreach (var t in required)
                {
                    if (!found.Contains(t))
                        missing.Add(t);
                }

                return missing;
            }
            catch (MySqlException mex)
            {
                // 数据库不存在
                if (mex.Number == 1049)
                {
                    return new List<string> { "__DATABASE_MISSING__" };
                }
                return new List<string> { $"__DB_ERROR__:{mex.Message}" };
            }
            catch (Exception)
            {
                return new List<string> { "__DB_CHECK_FAILED__" };
            }
        }

        /// <summary>
        /// 创建数据库（若不存在）并创建/修复核心表结构（新方案：cad_block_attributes_json）。
        /// 返回 true 表示成功（创建成功或已存在且修复成功），false 表示失败。
        /// </summary>
        public static bool CreateDatabaseAndCoreTables(string server, int port, string user, string password, string database = "cad_sw_library")
        {
            try
            {
                // =========================
                // 第1步：确保数据库存在
                // =========================
                var masterConn = $"Server={server};Port={port};Uid={user};Pwd={password};";
                using (var conn = new MySqlConnection(masterConn))
                {
                    conn.Open();
                    // 创建数据库（若不存在），统一字符集与排序规则，避免中文乱码
                    var createDbSql = $"CREATE DATABASE IF NOT EXISTS `{database}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;";
                    conn.Execute(createDbSql);
                }

                // =========================
                // 第2步：连接目标数据库
                // =========================
                var dbConn = $"Server={server};Port={port};Database={database};Uid={user};Pwd={password};Allow User Variables=True;";
                using (var conn = new MySqlConnection(dbConn))
                {
                    conn.Open();

                    // 统一执行DDL的小工具，便于阅读和排错
                    void Exec(string sql)
                    {
                        conn.Execute(sql);
                    }

                    // 检查列是否存在
                    bool ColumnExists(string tableName, string columnName)
                    {
                        const string sql = @"
SELECT COUNT(1)
FROM information_schema.columns
WHERE table_schema = @schema
  AND table_name = @table
  AND column_name = @column;";
                        return conn.QuerySingle<int>(sql, new { schema = database, table = tableName, column = columnName }) > 0;
                    }

                    // 检查索引是否存在
                    bool IndexExists(string tableName, string indexName)
                    {
                        const string sql = @"
SELECT COUNT(1)
FROM information_schema.statistics
WHERE table_schema = @schema
  AND table_name = @table
  AND index_name = @index;";
                        return conn.QuerySingle<int>(sql, new { schema = database, table = tableName, index = indexName }) > 0;
                    }

                    // 检查外键约束是否存在
                    bool ForeignKeyExists(string tableName, string fkName)
                    {
                        const string sql = @"
SELECT COUNT(1)
FROM information_schema.table_constraints
WHERE table_schema = @schema
  AND table_name = @table
  AND constraint_name = @fk
  AND constraint_type = 'FOREIGN KEY';";
                        return conn.QuerySingle<int>(sql, new { schema = database, table = tableName, fk = fkName }) > 0;
                    }

                    // =========================
                    // 第3步：创建核心表（若不存在）
                    // =========================

                    // CAD 主分类表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `cad_categories` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `name` VARCHAR(200) NOT NULL,
    `display_name` VARCHAR(200) NULL,
    `subcategory_ids` TEXT NULL,
    `sort_order` INT DEFAULT 0,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // CAD 子分类表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `cad_subcategories` (
    `id` INT NOT NULL PRIMARY KEY,
    `parent_id` INT NOT NULL,
    `name` VARCHAR(200) NOT NULL,
    `display_name` VARCHAR(200) NULL,
    `sort_order` INT DEFAULT 0,
    `level` INT DEFAULT 1,
    `subcategory_ids` TEXT NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX `idx_cad_sub_parent` (`parent_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // CAD 文件主表（沿用你现有项目字段）
                    Exec(@"
CREATE TABLE IF NOT EXISTS `cad_file_storage` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `category_id` INT NULL,
    `file_attribute_id` VARCHAR(64) NULL COMMENT '兼容字段：可存属性配置名或业务ID',
    `file_name` VARCHAR(512) NULL,
    `file_stored_name` VARCHAR(255) NOT NULL,
    `display_name` VARCHAR(255) NOT NULL,
    `file_type` VARCHAR(255) NULL,
    `file_hash` VARCHAR(255) NOT NULL,
    `block_name` VARCHAR(255) NULL,
    `layer_name` VARCHAR(100) NOT NULL,
    `color_index` INT NOT NULL DEFAULT 256,
    `file_path` VARCHAR(500) NOT NULL,
    `preview_image_name` VARCHAR(255) NULL,
    `preview_image_path` VARCHAR(500) NULL,
    `file_size` BIGINT NULL,
    `created_at` TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
    `updated_at` TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    `is_preview` TINYINT(1) DEFAULT 0,
    `version` INT DEFAULT 1,
    `description` TEXT NULL,
    `is_active` TINYINT(1) DEFAULT 1,
    `created_by` VARCHAR(255) NULL,
    `category_type` VARCHAR(50) DEFAULT 'sub',
    `title` VARCHAR(255) NULL,
    `keywords` TEXT NULL,
    `is_public` TINYINT(1) DEFAULT 1,
    `updated_by` VARCHAR(255) NULL,
    `last_accessed_at` DATETIME NULL,
    `is_tianzheng` TINYINT(1) NULL,
    `scale` DOUBLE NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 新属性表（JSON方案核心）
                    Exec(@"
CREATE TABLE IF NOT EXISTS `cad_block_attributes_json` (
    `attr_id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY COMMENT '属性记录唯一标识',
    `file_id` INT NOT NULL COMMENT '关联 cad_file_storage.id',
    `config_name` VARCHAR(100) NULL COMMENT '配置名称，如 DN50 配置',
    `attributes_json` TEXT NULL COMMENT 'JSON格式属性字典',
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX `idx_fileid` (`file_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 系统配置表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `system_config` (
    `config_key` VARCHAR(200) PRIMARY KEY,
    `config_value` TEXT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 部门表（包含 cad_category_id，兼容 MySqlAuthService 中的同步逻辑）
                    Exec(@"
CREATE TABLE IF NOT EXISTS `departments` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `cad_category_id` INT NULL,
    `name` VARCHAR(200) NOT NULL,
    `display_name` VARCHAR(200) NULL,
    `description` TEXT NULL,
    `manager_user_id` INT NULL,
    `sort_order` INT DEFAULT 0,
    `is_active` TINYINT(1) DEFAULT 1,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 用户表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `users` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `username` VARCHAR(100) NOT NULL UNIQUE,
    `password_hash` VARCHAR(512) NULL,
    `display_name` VARCHAR(200) NULL,
    `gender` VARCHAR(16) NULL,
    `phone` VARCHAR(32) NULL,
    `email` VARCHAR(200) NULL,
    `department_id` INT NULL,
    `role` VARCHAR(64) NULL,
    `status` TINYINT DEFAULT 1,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX `idx_users_department` (`department_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 部门-用户关系表（可选多对多）
                    Exec(@"
CREATE TABLE IF NOT EXISTS `department_users` (
    `department_id` INT NOT NULL,
    `user_id` INT NOT NULL,
    PRIMARY KEY (`department_id`,`user_id`),
    INDEX `idx_department_users_user` (`user_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 分类-部门映射表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `category_department_map` (
    `category_id` INT NOT NULL PRIMARY KEY,
    `department_id` INT NOT NULL UNIQUE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 文件访问日志
                    Exec(@"
CREATE TABLE IF NOT EXISTS `file_access_logs` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `file_id` INT NULL,
    `user_name` VARCHAR(200) NULL,
    `action_type` VARCHAR(50) NULL,
    `ip_address` VARCHAR(64) NULL,
    `user_agent` VARCHAR(512) NULL,
    `access_time` DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 文件标签表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `file_tags` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `file_id` INT NULL,
    `tag_name` VARCHAR(200) NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 文件版本历史
                    Exec(@"
CREATE TABLE IF NOT EXISTS `file_version_history` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `file_id` INT NULL,
    `version` INT NULL,
    `file_name` VARCHAR(512) NULL,
    `stored_file_name` VARCHAR(512) NULL,
    `file_path` VARCHAR(1024) NULL,
    `file_size` BIGINT NULL,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_by` VARCHAR(200) NULL,
    `change_description` TEXT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // SW 分类表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `sw_categories` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `name` VARCHAR(200) NOT NULL,
    `display_name` VARCHAR(200) NULL,
    `sort_order` INT DEFAULT 0,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // SW 子分类表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `sw_subcategories` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `category_id` INT NOT NULL,
    `parent_id` INT NOT NULL DEFAULT 0,
    `name` VARCHAR(200) NOT NULL,
    `display_name` VARCHAR(200) NULL,
    `sort_order` INT DEFAULT 0,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX `idx_sw_sub_cat` (`category_id`),
    INDEX `idx_sw_sub_parent` (`parent_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // SW 图元表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `sw_graphics` (
    `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `subcategory_id` INT NOT NULL,
    `file_name` VARCHAR(512) NOT NULL,
    `display_name` VARCHAR(512) NULL,
    `file_path` VARCHAR(1024) NULL,
    `preview_image_path` VARCHAR(1024) NULL,
    `file_size` BIGINT NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX `idx_sw_graphics_sub` (`subcategory_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // 设备信息表
                    Exec(@"
CREATE TABLE IF NOT EXISTS `device_info` (
    `device_id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    `device_name` VARCHAR(255) NOT NULL,
    `device_type` VARCHAR(100) NULL,
    `medium_name` VARCHAR(100) NULL,
    `specifications` VARCHAR(255) NULL,
    `material` VARCHAR(100) NULL,
    `quantity` INT DEFAULT 0,
    `drawing_number` VARCHAR(100) NULL,
    `power` DECIMAL(18,6) NULL,
    `volume` DECIMAL(18,6) NULL,
    `pressure` DECIMAL(18,6) NULL,
    `temperature` DECIMAL(18,6) NULL,
    `diameter` DECIMAL(18,6) NULL,
    `length` DECIMAL(18,6) NULL,
    `thickness` DECIMAL(18,6) NULL,
    `weight` DECIMAL(18,6) NULL,
    `model` VARCHAR(255) NULL,
    `remarks` TEXT NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // =========================
                    // 第4步：修复已有库的列/索引/外键（增量修复）
                    // =========================

                    // 确保 cad_file_storage 关键兼容列存在
                    if (!ColumnExists("cad_file_storage", "file_attribute_id"))
                    {
                        Exec("ALTER TABLE `cad_file_storage` ADD COLUMN `file_attribute_id` VARCHAR(64) NULL COMMENT '兼容字段：属性配置名或业务ID';");
                    }
                    if (!ColumnExists("cad_file_storage", "category_type"))
                    {
                        Exec("ALTER TABLE `cad_file_storage` ADD COLUMN `category_type` VARCHAR(50) DEFAULT 'sub';");
                    }

                    // 确保常用索引存在
                    if (!IndexExists("cad_file_storage", "idx_cfs_category"))
                    {
                        Exec("ALTER TABLE `cad_file_storage` ADD INDEX `idx_cfs_category` (`category_id`, `category_type`);");
                    }
                    if (!IndexExists("cad_file_storage", "idx_cfs_attr_biz"))
                    {
                        Exec("ALTER TABLE `cad_file_storage` ADD INDEX `idx_cfs_attr_biz` (`file_attribute_id`);");
                    }
                    if (!IndexExists("cad_file_storage", "idx_cfs_file_hash"))
                    {
                        Exec("ALTER TABLE `cad_file_storage` ADD INDEX `idx_cfs_file_hash` (`file_hash`);");
                    }

                    // 确保 departments 的 cad_category_id 索引存在（兼容 MySqlAuthService）
                    if (!IndexExists("departments", "idx_cad_category_id"))
                    {
                        Exec("ALTER TABLE `departments` ADD INDEX `idx_cad_category_id` (`cad_category_id`);");
                    }

                    // 确保 JSON 属性表外键存在（file_id -> cad_file_storage.id）
                    if (!ForeignKeyExists("cad_block_attributes_json", "fk_file_id"))
                    {
                        // 先确保 file_id 有索引
                        if (!IndexExists("cad_block_attributes_json", "idx_fileid"))
                        {
                            Exec("ALTER TABLE `cad_block_attributes_json` ADD INDEX `idx_fileid` (`file_id`);");
                        }

                        // 添加外键，启用级联删除，防止孤儿属性记录
                        Exec(@"
ALTER TABLE `cad_block_attributes_json`
ADD CONSTRAINT `fk_file_id`
FOREIGN KEY (`file_id`) REFERENCES `cad_file_storage`(`id`)
ON DELETE CASCADE
ON UPDATE CASCADE;");
                    }

                    // =========================
                    // 第5步：可选清理提醒（不自动删旧表）
                    // =========================
                    // 这里不自动 DROP 旧 cad_file_attributes，避免误删历史数据。
                    // 若你确认旧表永不再用，可在数据库中手工删除。

                    return true;
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"CreateDatabaseAndCoreTables 失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 主窗口
        /// </summary>
        private readonly WpfMainWindow _wpfMainWindow;
        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        public readonly string _connectionString;
        /// <summary>
        /// 数据库是否可用
        /// </summary>
        public bool IsDatabaseAvailable { get; private set; } = true;
        /// <summary>
        /// 数据库管理类构造函数
        /// </summary>
        /// <param name="connectionString"> 链接字符串
        ///  </param>
        public DatabaseManager(string connectionString)
        {
            _connectionString = connectionString;
            IsDatabaseAvailable = TestDatabaseConnection();
        }
        /// <summary>
        /// 测试数据库连接
        /// </summary>
        /// <returns></returns>
        private bool TestDatabaseConnection()
        {
            try
            {
                using var connection = GetConnection();
                connection.Open();
                LogManager.Instance.LogInfo("数据库连接测试成功");
                return true;
            }
            catch (MySqlException ex)
            {
                LogManager.Instance.LogInfo($"MySQL连接错误: {ex.Number} - {ex.Message}");
                switch (ex.Number)
                {
                    case 0:
                        LogManager.Instance.LogInfo("无法连接到MySQL服务器");
                        break;
                    case 1042:
                        LogManager.Instance.LogInfo("无法解析主机名");
                        break;
                    case 1045:
                        LogManager.Instance.LogInfo("用户名或密码错误");
                        break;
                    case 1049:
                        LogManager.Instance.LogInfo("未知数据库");
                        break;
                    case 2002:
                        LogManager.Instance.LogInfo("连接超时或服务器无响应");
                        break;
                    default:
                        LogManager.Instance.LogInfo($"MySQL错误代码: {ex.Number}");
                        break;
                }
                return false;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"数据库连接测试失败: {ex.Message}");
                return false;
            }
        }

        #region 部门与人员同步方法

        /// <summary>
        /// 新增：部门实体
        /// </summary>
        public class Department
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string DisplayName { get; set; }
            public int SortOrder { get; set; }
            public bool IsActive { get; set; }
            public DateTime CreatedAt { get; set; }
            public DateTime UpdatedAt { get; set; }
        }

        /// <summary>
        /// 获取所有部门（用于注册窗口下拉列表）
        /// </summary>
        public async Task<List<Department>> GetAllDepartmentsAsync()
        {
            const string sql = @"
              SELECT
                  id AS Id,
                  name AS Name,
                  display_name AS DisplayName,
                  sort_order AS SortOrder,
                  is_active AS IsActive,
                  created_at AS CreatedAt,
                  updated_at AS UpdatedAt
              FROM departments
              ORDER BY sort_order, name";
            try
            {
                using var connection = new MySqlConnection(_connectionString);
                var depts = await connection.QueryAsync<Department>(sql).ConfigureAwait(false);
                return depts.AsList();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetAllDepartmentsAsync 出错: {ex.Message}");
                return new List<Department>();
            }
        }


        /// <summary>
        /// 将 cad_categories 中尚未映射到 departments 的分类逐条创建部门并建立映射。
        /// 保持幂等：已存在的映射或同名部门不会重复创建（会尝试复用同名部门）。
        /// </summary>
        public async Task SyncDepartmentsFromCadCategoriesAsync()
        {
            try
            {
                var categories = await GetAllCadCategoriesAsync().ConfigureAwait(false);
                if (categories == null || categories.Count == 0) return;

                using var conn = GetConnection();
                await conn.OpenAsync().ConfigureAwait(false);

                foreach (var cat in categories)
                {
                    // 检查是否已有映射
                    var mapSql = "SELECT department_id FROM category_department_map WHERE category_id = @CategoryId";
                    var mapped = await conn.QueryFirstOrDefaultAsync<int?>(mapSql, new { CategoryId = cat.Id }).ConfigureAwait(false);
                    if (mapped.HasValue) continue;

                    // 尝试按名称查找已有部门
                    var deptSql = "SELECT id FROM departments WHERE name = @Name LIMIT 1";
                    var deptId = await conn.QueryFirstOrDefaultAsync<int?>(deptSql, new { Name = cat.Name }).ConfigureAwait(false);
                    if (!deptId.HasValue)
                    {
                        // 插入新部门
                        var insertDeptSql = @"INSERT INTO departments (name, display_name, sort_order, created_at, updated_at) 
                                              VALUES (@Name, @DisplayName, @SortOrder, NOW(), NOW());
                                              SELECT LAST_INSERT_ID();";
                        var newDeptId = await conn.ExecuteScalarAsync<int>(insertDeptSql, new { Name = cat.Name, DisplayName = cat.DisplayName ?? cat.Name, SortOrder = cat.SortOrder }).ConfigureAwait(false);
                        deptId = newDeptId;
                    }

                    if (deptId.HasValue)
                    {
                        // 建立映射
                        var insertMap = "INSERT INTO category_department_map (category_id, department_id) VALUES (@CategoryId, @DepartmentId)";
                        await conn.ExecuteAsync(insertMap, new { CategoryId = cat.Id, DepartmentId = deptId.Value }).ConfigureAwait(false);
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategoriesAsync 出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 删除与分类相关的部门映射与（可选）部门记录。
        /// 当分类被删除时调用：如果该部门没有其他映射且没有用户（或你选择直接删除），则删除部门。
        /// </summary>
        public async Task RemoveDepartmentMappingForCategoryAsync(int categoryId)
        {
            try
            {
                using var conn = GetConnection();
                await conn.OpenAsync().ConfigureAwait(false);

                var getDeptSql = "SELECT department_id FROM category_department_map WHERE category_id = @CategoryId";
                var deptId = await conn.QueryFirstOrDefaultAsync<int?>(getDeptSql, new { CategoryId = categoryId }).ConfigureAwait(false);
                if (!deptId.HasValue)
                {
                    // 无映射， nothing to do
                    return;
                }

                // 删除映射
                var delMapSql = "DELETE FROM category_department_map WHERE category_id = @CategoryId";
                await conn.ExecuteAsync(delMapSql, new { CategoryId = categoryId }).ConfigureAwait(false);

                // 检查该部门是否仍被其它分类映射或有用户
                var usedByCatSql = "SELECT COUNT(*) FROM category_department_map WHERE department_id = @DepartmentId";
                var usedByCat = await conn.QuerySingleAsync<int>(usedByCatSql, new { DepartmentId = deptId.Value }).ConfigureAwait(false);
                var userCountSql = "SELECT COUNT(*) FROM users WHERE department_id = @DepartmentId";
                var userCount = await conn.QuerySingleAsync<int>(userCountSql, new { DepartmentId = deptId.Value }).ConfigureAwait(false);

                if (usedByCat == 0 && userCount == 0)
                {
                    var delDeptSql = "DELETE FROM departments WHERE id = @DepartmentId";
                    await conn.ExecuteAsync(delDeptSql, new { DepartmentId = deptId.Value }).ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"RemoveDepartmentMappingForCategoryAsync 出错: {ex.Message}");
            }
        }

        #endregion

        #region CAD分类操作

        /// <summary>
        /// 添加CAD分类（并同步创建部门映射）
        /// </summary>
        public async Task<int> AddCadCategoryAsync(CadCategory category)
        {
            using var connection = GetConnection();
            var sql = @"INSERT INTO cad_categories (name, display_name, sort_order) 
                VALUES (@Name, @DisplayName, @SortOrder)";
            var affected = await connection.ExecuteAsync(sql, category).ConfigureAwait(false);

            // 异步触发同步（保证分类与部门一致）
            _ = SyncDepartmentsFromCadCategoriesAsync();
            return affected;
        }

        /// <summary>
        /// 修改CAD分类（并同步部门信息）
        /// </summary>
        public async Task<int> UpdateCadCategoryAsync(CadCategory category)
        {
            using var connection = GetConnection();
            var sql = @"UPDATE cad_categories 
                SET name = @Name, display_name = @DisplayName, sort_order = @SortOrder, updated_at = NOW() 
                WHERE id = @Id";
            var affected = await connection.ExecuteAsync(sql, category).ConfigureAwait(false);

            // 如果分类名或显示名变更，更新对应部门（若已存在映射）
            try
            {
                using var conn = GetConnection();
                await conn.OpenAsync().ConfigureAwait(false);
                var mapSql = "SELECT department_id FROM category_department_map WHERE category_id = @CategoryId";
                var deptId = await conn.QueryFirstOrDefaultAsync<int?>(mapSql, new { CategoryId = category.Id }).ConfigureAwait(false);
                if (deptId.HasValue)
                {
                    var updateDeptSql = @"UPDATE departments SET name = @Name, display_name = @DisplayName, updated_at = NOW() WHERE id = @Id";
                    await conn.ExecuteAsync(updateDeptSql, new { Name = category.Name, DisplayName = category.DisplayName ?? category.Name, Id = deptId.Value }).ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"UpdateCadCategoryAsync 同步部门时出错: {ex.Message}");
            }

            // 确保全量同步以修正遗漏
            _ = SyncDepartmentsFromCadCategoriesAsync();
            return affected;
        }

        /// <summary>
        /// 删除CAD分类（并删除部门映射及可选部门）
        /// </summary>
        public async Task<int> DeleteCadCategoryAsync(int id)
        {
            using var connection = GetConnection();
            var sql = "DELETE FROM cad_categories WHERE id = @Id";
            var affected = await connection.ExecuteAsync(sql, new { Id = id }).ConfigureAwait(false);

            // 删除映射并在必要时删除部门
            _ = RemoveDepartmentMappingForCategoryAsync(id);
            return affected;
        }

        /// <summary>
        /// 获取所有CAD分类
        /// </summary>
        /// <returns> 返回List<CadCategory>分类list</returns>
        public async Task<List<CadCategory>> GetAllCadCategoriesAsync()
        {
            try
            {
                const string sql = @"
                                   SELECT 
                                       id AS Id,
                                       name AS Name,
                                       display_name AS DisplayName,
                                       subcategory_ids AS SubcategoryIds,
                                       sort_order AS SortOrder,
                                       created_at AS CreatedAt,
                                       updated_at AS UpdatedAt
                                   FROM cad_categories 
                                   ORDER BY sort_order";

                using var connection = new MySqlConnection(_connectionString);
                var categories = await connection.QueryAsync<CadCategory>(sql);
                LogManager.Instance.LogInfo($"查询返回 {categories.AsList().Count} 条记录");
                return categories.AsList();

            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"数据库查询出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 根据名称获取CAD分类
        /// </summary>
        public async Task<CadCategory> GetCadCategoryByNameAsync(string categoryName)
        {
            const string sql = @"
            SELECT 
               id AS Id,
               name AS Name,
               display_name AS DisplayName,
               subcategory_ids AS SubcategoryIds,
               sort_order AS SortOrder,
               created_at AS CreatedAt,
               updated_at AS UpdatedAt
            FROM cad_categories 
            WHERE name LIKE @Name";
            try
            {
                using var connection = new MySqlConnection(_connectionString);
                var parameters = new Dictionary<string, object>();
                parameters.Add("Name", $"%{categoryName}%");
                var result = await connection.QuerySingleOrDefaultAsync<CadCategory>(sql, parameters);
                return result;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"数据库查询出错: {ex.Message}");
                throw;
            }


        }

        #endregion

        #region CAD子分类操作


        /// <summary>
        /// 获取所有CAD子分类
        /// </summary>
        /// <returns></returns>
        public async Task<List<CadSubcategory>> GetAllCadSubcategoriesAsync()
        {
            using var connection = GetConnection();
            //var sql = "SELECT * FROM cad_subcategories ORDER BY parent_id, id, parent_id, name, display_name, level , subcategory_ids, sort_order";
            const string sql = @"
                               SELECT 
                                   id AS Id,
                                   parent_id AS ParentId,
                                   name AS Name,
                                   display_name AS DisplayName,
                                   sort_order AS SortOrder,
                                   level AS Level,
                                   subcategory_ids AS SubcategoryIds,
                                   created_at AS CreatedAt,
                                   updated_at AS UpdatedAt
                               FROM cad_subcategories 
                               ORDER BY parent_id, id, parent_id, name, display_name, level , subcategory_ids, sort_order";
            return (await connection.QueryAsync<CadSubcategory>(sql)).AsList();
        }

        /// <summary>
        /// 通过Id获取子分类的方法
        /// </summary>
        /// <returns>  </returns>
        public async Task<CadSubcategory> GetCadSubcategoryByIdAsync(int id)
        {
            const string sql = @"
                               SELECT 
                                   id AS Id,
                                   parent_id AS ParentId,
                                   name AS Name,
                                   display_name AS DisplayName,
                                   sort_order AS SortOrder,
                                   level AS Level,
                                   subcategory_ids AS SubcategoryIds,
                                   created_at AS CreatedAt,
                                   updated_at AS UpdatedAt
                               FROM cad_subcategories 
                               WHERE id = @id";

            using var connection = new MySqlConnection(_connectionString);
            return await connection.QuerySingleOrDefaultAsync<CadSubcategory>(sql, new { id });
        }


        /// <summary>
        /// 根据子分类ID获取这个子分类同级的所有兄弟子分类
        /// </summary>
        public async Task<List<CadSubcategory>> GetCadSubcategoriesByCategoryIdAsync(int categoryId)
        {
            const string sql = @"
                               SELECT 
                                    id AS Id,
                                    parent_id AS ParentId,
                                    name AS Name,
                                    display_name AS DisplayName,
                                    sort_order AS SortOrder,
                                    level AS Level,
                                    subcategory_ids AS SubcategoryIds,
                                    created_at AS CreatedAt,
                                    updated_at AS UpdatedAt
                               FROM cad_subcategories 
                               WHERE parent_id = @ParentId 
                               ORDER BY sort_order";
            try
            {
                using var connection = new MySqlConnection(_connectionString);
                var parameters = new Dictionary<string, object>();
                parameters.Add("ParentId", categoryId);

                var subcategories = await connection.QueryAsync<CadSubcategory>(sql, parameters);
                return subcategories.AsList();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"数据库查询出错: {ex.Message}");
                throw;
            }

        }

        /// <summary>
        /// 根据父ID获取子分类（用于递归加载）
        /// </summary>
        public async Task<List<CadSubcategory>> GetCadSubcategoriesByParentIdAsync(int parentId)
        {
            try
            {
                const string sql = @"
                               SELECT 
                                   *
                               FROM cad_subcategories 
                               WHERE parent_id = @parentId 
                               ORDER BY sort_order";

                using var connection = new MySqlConnection(_connectionString);
                var subcategories = await connection.QueryAsync<CadSubcategory>(sql, new { parentId });
                return subcategories.AsList();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"获取子分类时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 添加CAD子分类
        /// </summary>
        /// <param name="subcategory"></param>
        /// <returns>返回受影响的行数</returns>
        public async Task<int> AddCadSubcategoryAsync(CadSubcategory subcategory)
        {
            try
            {
                // 验证输入参数
                if (subcategory == null)
                    throw new ArgumentNullException(nameof(subcategory));

                if (string.IsNullOrEmpty(subcategory.Name))
                    throw new ArgumentException("子分类名称不能为空", nameof(subcategory.Name));

                using var connection = GetConnection();

                // SQL语句修正：移除id字段（假设是自增），修正参数名
                var sql = @"INSERT INTO cad_subcategories ( id,parent_id, name, display_name, sort_order, level, subcategory_ids, created_at, updated_at) 
            VALUES ( @Id, @ParentId, @Name, @DisplayName, @SortOrder, @Level, @SubcategoryIds, NOW(), NOW())";

                var result = await connection.ExecuteAsync(sql, subcategory);
                LogManager.Instance.LogInfo($"成功添加子分类: {subcategory.Name}");
                return result;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"添加CAD子分类时出错: {ex.Message}");
                // throw new Exception($"添加CAD子分类失败: {ex.Message}", ex);
                return 0;
            }
        }
        /// <summary>
        /// 修改CAD子分类
        /// </summary>
        /// <param name="category"></param>
        /// <returns></returns>
        public async Task<int> UpdateCadSubcategoryAsync(CadSubcategory cadSubcategory)
        {
            using var connection = GetConnection();
            var sql = @"UPDATE cad_subcategories 
           SET  parent_id = @ParentId, name = @Name, display_name = @DisplayName, subcategory_ids = @newSubcategoryIds, sort_order = @SortOrder, updated_at = NOW() 
           WHERE id = @Id";
            return await connection.ExecuteAsync(sql, cadSubcategory);
        }
        /// <summary>
        /// 添加更新父级子分类列表的方法
        /// </summary>
        /// <param name="parentId"></param>
        /// <param name="newSubcategoryId"></param>
        /// <returns></returns>
        public async Task<int> UpdateParentSubcategoryListAsync(int parentId, int newSubcategoryId)
        {
            try
            {
                using var connection = GetConnection();
                string selectSql;// 获取父级记录
                object parameters;
                if (parentId >= 10000)
                {
                    selectSql = "SELECT subcategory_ids FROM cad_subcategories WHERE id = @parentId";   // 父级是子分类
                    parameters = new { parentId };
                }
                else
                {
                    selectSql = "SELECT subcategory_ids FROM cad_categories WHERE id = @parentId";   // 父级是主分类
                    parameters = new { parentId };
                }
                string currentSubcategoryIds = await connection.QuerySingleOrDefaultAsync<string>(selectSql, parameters);
                string newSubcategoryIds;// 更新子分类列表
                if (string.IsNullOrEmpty(currentSubcategoryIds))
                {
                    newSubcategoryIds = newSubcategoryId.ToString();// 创建新的子分类列表
                }
                else
                {
                    var ids = currentSubcategoryIds.Split(',').Select(id => id.Trim()).ToList();// 将字符串转换为列表
                    if (!ids.Contains(newSubcategoryId.ToString()))// 如果不存在
                    {
                        ids.Add(newSubcategoryId.ToString());// 添加
                        newSubcategoryIds = string.Join(",", ids);// 重新组合为字符串
                    }
                    else
                    {
                        newSubcategoryIds = currentSubcategoryIds; // 已存在，不需要更新
                    }
                }
                string updateSql; // 更新数据库
                if (parentId >= 10000)
                {
                    updateSql = "UPDATE cad_subcategories SET subcategory_ids = @newSubcategoryIds, updated_at = NOW() WHERE id = @parentId"; // 更新子分类表
                }
                else
                {
                    updateSql = "UPDATE cad_categories SET subcategory_ids = @newSubcategoryIds, updated_at = NOW() WHERE id = @parentId"; // 更新主分类表
                }

                return await connection.ExecuteAsync(updateSql, new { newSubcategoryIds, parentId });
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"更新父级子分类列表失败: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 添加更新父级子分类列表的方法
        /// </summary>
        /// <param name="parentId"></param>
        /// <param name="newSubcategoryId"></param>
        /// <returns></returns>
        public async Task<int> UpdateParentSubcategoryListAsync(int parentId, string newSubcategoryIds)
        {
            try
            {
                using var connection = GetConnection();// 创建数据库连接
                string updateSql; // 更新数据库
                if (parentId >= 10000)
                {
                    updateSql = "UPDATE cad_subcategories SET subcategory_ids = @newSubcategoryIds, updated_at = NOW() WHERE id = @parentId"; // 更新子分类表
                }
                else
                {
                    updateSql = "UPDATE cad_categories SET subcategory_ids = @newSubcategoryIds, updated_at = NOW() WHERE id = @parentId"; // 更新主分类表
                }

                return await connection.ExecuteAsync(updateSql, new { newSubcategoryIds, parentId });
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"更新父级子分类列表失败: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 删除CAD子分类
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<int> DeleteCadSubcategoryAsync(int id)
        {
            //这个方法还有不完善的地方，比如子分类下还有子分类或图元，如果不删除子分类下的图元，则无法删除子分类，需要前删除子分类下的图元，不然这个分类下的图元与子分类就是在数据库中的垃圾数据；
            using var connection = GetConnection();
            var sql = "DELETE FROM cad_subcategories WHERE id = @Id";
            return await connection.ExecuteAsync(sql, new { Id = id });
        }
        #endregion

        #region SW分类操作
        /// <summary>
        /// 获取所有SW分类
        /// </summary>
        public async Task<List<SwCategory>> GetAllSwCategoriesAsync()
        {
            const string sql = @"
            SELECT 
                id,
                name,
                display_name,
                sort_order,
                created_at,
                updated_at
            FROM sw_categories 
            ORDER BY sort_order";

            using var connection = new MySqlConnection(_connectionString);
            var categories = await connection.QueryAsync<SwCategory>(sql);
            return categories.AsList();
        }

        /// <summary>
        /// 添加SW分类
        /// </summary>
        /// <param name="category"></param>
        /// <returns></returns>
        public async Task<int> AddSwCategoryAsync(SwCategory category)
        {
            using var connection = GetConnection();
            var sql = @"INSERT INTO sw_categories (name, display_name, sort_order) 
                VALUES (@Name, @DisplayName, @SortOrder)";
            return await connection.ExecuteAsync(sql, category);
        }
        /// <summary>
        /// 修改SW分类
        /// </summary>
        /// <param name="category"></param>
        /// <returns></returns>
        public async Task<int> UpdateSwCategoryAsync(SwCategory category)
        {
            using var connection = GetConnection();
            var sql = @"UPDATE sw_categories 
                SET name = @Name, display_name = @DisplayName, sort_order = @SortOrder, updated_at = NOW() 
                WHERE id = @Id";
            return await connection.ExecuteAsync(sql, category);
        }
        /// <summary>
        /// 删除SW分类
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<int> DeleteSwCategoryAsync(int id)
        {
            using var connection = GetConnection();
            var sql = "DELETE FROM sw_categories WHERE id = @Id";
            return await connection.ExecuteAsync(sql, new { Id = id });
        }
        /// <summary>
        /// 根据父ID获取SW子分类
        /// </summary>
        public async Task<List<SwSubcategory>> GetSwSubcategoriesByParentIdAsync(int parentId)
        {
            using var connection = GetConnection();
            var sql = @"SELECT * FROM sw_subcategories 
                WHERE parent_id = @ParentId 
                ORDER BY sort_order, name";
            return (await connection.QueryAsync<SwSubcategory>(sql, new { ParentId = parentId })).AsList();
        }
        #endregion

        #region SW子分类操作
        /// <summary>
        /// 获取指定SW分类下的所有SW子分类
        /// </summary>
        /// <param name="categoryId"></param>
        /// <returns></returns>
        public async Task<List<SwSubcategory>> GetSwSubcategoriesByCategoryIdAsync(int categoryId)
        {
            using var connection = GetConnection();
            var sql = @"SELECT * FROM sw_subcategories 
                WHERE category_id = @CategoryId 
                ORDER BY parent_id, sort_order, name";
            return (await connection.QueryAsync<SwSubcategory>(sql, new { CategoryId = categoryId })).AsList();
        }
        /// <summary>
        /// 获取所有SW子分类
        /// </summary>
        /// <returns></returns>
        public async Task<List<SwSubcategory>> GetAllSwSubcategoriesAsync()
        {
            using var connection = GetConnection();
            var sql = "SELECT * FROM sw_subcategories ORDER BY category_id, parent_id, sort_order, name";
            return (await connection.QueryAsync<SwSubcategory>(sql)).AsList();
        }
        /// <summary>
        /// 添加SW子分类
        /// </summary>
        /// <param name="subcategory"></param>
        /// <returns></returns>
        public async Task<int> AddSwSubcategoryAsync(SwSubcategory subcategory)
        {
            using var connection = GetConnection();
            var sql = @"INSERT INTO sw_subcategories (category_id, name, display_name, parent_id, sort_order) 
                VALUES (@CategoryId, @Name, @DisplayName, @ParentId, @SortOrder)";
            return await connection.ExecuteAsync(sql, subcategory);
        }
        /// <summary>
        /// 修改SW子分类
        /// </summary>
        /// <param name="subcategory"></param>
        /// <returns></returns>
        public async Task<int> UpdateSwSubcategoryAsync(SwSubcategory subcategory)
        {
            using var connection = GetConnection();
            var sql = @"UPDATE sw_subcategories 
                SET name = @Name, display_name = @DisplayName, parent_id = @ParentId, sort_order = @SortOrder, updated_at = NOW() 
                WHERE id = @Id";
            return await connection.ExecuteAsync(sql, subcategory);
        }
        /// <summary>
        /// 删除SW子分类
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<int> DeleteSwSubcategoryAsync(int id)
        {
            using var connection = GetConnection();
            var sql = "DELETE FROM sw_subcategories WHERE id = @Id";
            return await connection.ExecuteAsync(sql, new { Id = id });
        }
        #endregion

        #region SW图元操作
        /// <summary>
        /// 根据子分类ID获取SW图元
        /// </summary>
        public async Task<List<SwGraphic>> GetSwGraphicsBySubcategoryIdAsync(int subcategoryId)
        {
            const string sql = @"
                               SELECT 
                                   id,
                                   subcategory_id,
                                   file_name,
                                   display_name,
                                   file_path,
                                   preview_image_path,
                                   file_size,
                                   created_at,
                                   updated_at
                               FROM sw_graphics 
                               WHERE subcategory_id = @subcategoryId 
                               ORDER BY file_name";

            using var connection = new MySqlConnection(_connectionString);
            var graphics = await connection.QueryAsync<SwGraphic>(sql, new { subcategoryId });
            return graphics.AsList();
        }

        /// <summary>
        /// 添加SW图元
        /// </summary>
        /// <param name="graphic"></param>
        /// <returns></returns>
        public async Task<int> AddSwGraphicAsync(SwGraphic graphic)
        {
            using var connection = GetConnection();
            var sql = @"INSERT INTO sw_graphics (subcategory_id, file_name, display_name, file_path, preview_image_path, file_size) 
                VALUES (@SubcategoryId, @FileName, @DisplayName, @FilePath, @PreviewImagePath, @FileSize)";
            return await connection.ExecuteAsync(sql, graphic);
        }
        /// <summary>
        /// 修改SW图元
        /// </summary>
        /// <param name="graphic"></param>
        /// <returns></returns>
        public async Task<int> UpdateSwGraphicAsync(SwGraphic graphic)
        {
            using var connection = GetConnection();
            var sql = @"UPDATE sw_graphics 
                SET file_name = @FileName, display_name = @DisplayName, file_path = @FilePath, 
                    preview_image_path = @PreviewImagePath, file_size = @FileSize, updated_at = NOW() 
                WHERE id = @Id";
            return await connection.ExecuteAsync(sql, graphic);
        }
        /// <summary>
        /// 删除SW图元
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public async Task<int> DeleteSwGraphicAsync(int id)
        {
            using var connection = GetConnection();
            var sql = "DELETE FROM sw_graphics WHERE id = @Id";
            return await connection.ExecuteAsync(sql, new { Id = id });
        }
        public async Task<SwGraphic> GetSwGraphicByIdAsync(int id)
        {
            using var connection = GetConnection();
            var sql = "SELECT * FROM sw_graphics WHERE id = @Id";
            return await connection.QueryFirstOrDefaultAsync<SwGraphic>(sql, new { Id = id });
        }
        #endregion

        #region 设备表相关操作

        /// <summary>
        /// 获取所有设备信息（用于设备表生成）
        /// </summary>
        public async Task<List<DeviceInfo>> GetAllDeviceInfoAsync()
        {
            const string sql = @"
            SELECT 
                device_id AS Id,
                device_name AS Name,
                device_type AS Type,
                medium_name AS MediumName,
                specifications AS Specifications,
                material AS Material,
                quantity AS Quantity,
                drawing_number AS DrawingNumber,
                power AS Power,
                volume AS Volume,
                pressure AS Pressure,
                temperature AS Temperature,
                diameter AS Diameter,
                length AS Length,
                thickness AS Thickness,
                weight AS Weight,
                model AS Model,
                remarks AS Remarks
            FROM device_info 
            ORDER BY device_name";

            using var connection = new MySqlConnection(_connectionString);
            var deviceList = await connection.QueryAsync<DeviceInfo>(sql);
            return deviceList.AsList();
        }

        /// <summary>
        /// 批量插入设备信息
        /// </summary>
        public async Task<int> InsertDeviceInfoBatchAsync(List<DeviceInfo> deviceList)
        {
            const string sql = @"
            INSERT INTO device_info (
                device_name, device_type, medium_name, specifications,
                material, quantity, drawing_number, power, volume, pressure,
                temperature, diameter, length, thickness, weight, model, remarks
            ) VALUES (
                @Name, @Type, @MediumName, @Specifications,
                @Material, @Quantity, @DrawingNumber, @Power, @Volume, @Pressure,
                @Temperature, @Diameter, @Length, @Thickness, @Weight, @Model, @Remarks
            )";

            using var connection = new MySqlConnection(_connectionString);
            return await connection.ExecuteAsync(sql, deviceList);
        }

        #endregion

        #region 事务操作示例

        /// <summary>
        /// 事务操作示例：批量更新图元信息
        /// </summary>
        public async Task<bool> UpdateFileBatchAsync(List<FileStorage> file)
        {
            const string updateSql = @"
            UPDATE cad_file_storage 
            SET display_name = @DisplayName,
                file_path = @FilePath,
                preview_image_path = @PreviewImagePath,
                updated_at = NOW()
            WHERE id = @Id";

            using var connection = new MySqlConnection(_connectionString);
            using var transaction = await connection.BeginTransactionAsync();

            try
            {
                await connection.ExecuteAsync(updateSql, file, transaction);
                await transaction.CommitAsync();
                return true;
            }
            catch
            {
                await transaction.RollbackAsync();
                return false;
            }
        }

        #endregion

        #region 系统配置操作
        /// <summary>
        /// 获取系统配置
        /// </summary>
        /// <param name="configKey"></param>
        /// <returns></returns>
        public async Task<string> GetConfigValueAsync(string configKey)
        {
            using var connection = GetConnection();
            var sql = "SELECT config_value FROM system_config WHERE config_key = @ConfigKey";
            return await connection.QueryFirstOrDefaultAsync<string>(sql, new { ConfigKey = configKey });
        }
        /// <summary>
        /// 设置系统配置
        /// </summary>
        /// <param name="configKey"></param>
        /// <param name="configValue"></param>
        /// <returns></returns>
        public async Task<int> SetConfigValueAsync(string configKey, string configValue)
        {
            using var connection = GetConnection();
            var sql = @"INSERT INTO system_config (config_key, config_value) 
                VALUES (@ConfigKey, @ConfigValue) 
                ON DUPLICATE KEY UPDATE config_value = @ConfigValue";
            return await connection.ExecuteAsync(sql, new { ConfigKey = configKey, ConfigValue = configValue });
        }
        /// <summary>
        /// 获取所有系统配置
        /// </summary>
        /// <returns></returns>
        public async Task<Dictionary<string, string>> GetAllConfigAsync()
        {
            using var connection = GetConnection();
            var sql = "SELECT config_key, config_value FROM system_config";
            var result = await connection.QueryAsync<(string, string)>(sql);
            return result.ToDictionary(x => x.Item1, x => x.Item2);
        }

        /// <summary>
        /// 获取CAD分类的最大排序序号
        /// </summary>
        public async Task<int> GetMaxCadCategorySortOrderAsync()
        {
            const string sql = "SELECT COALESCE(MAX(sort_order), 0) FROM cad_categories";

            using var connection = new MySqlConnection(_connectionString);
            var result = await connection.QuerySingleOrDefaultAsync<int>(sql);
            return result;
        }

        /// <summary>
        /// 获取指定父分类下子分类的最大排序序号
        /// </summary>
        public async Task<int> GetMaxCadSubcategorySortOrderAsync(int parentId)
        {
            const string sql = "SELECT COALESCE(MAX(sort_order), 0) FROM cad_subcategories WHERE parent_id = @parentId";

            using var connection = new MySqlConnection(_connectionString);
            var result = await connection.QuerySingleOrDefaultAsync<int>(sql, new { parentId });
            return result;
        }

        /// <summary>
        /// 获取所有子分类的最大排序序号（用于主分类下的直接子分类）
        /// </summary>
        public async Task<int> GetMaxCadSubcategorySortOrderForMainCategoryAsync(int parentId)
        {
            const string sql = "SELECT COALESCE(MAX(sort_order), 0) FROM cad_subcategories WHERE parent_id = @parentId";

            using var connection = new MySqlConnection(_connectionString);
            var result = await connection.QuerySingleOrDefaultAsync<int>(sql, new { parentId });
            return result;
        }

        #endregion

        #region 优化的文件管理方法

        /// <summary>
        /// 获取分类下的所有文件（支持分页和排序）
        /// </summary>
        public async Task<List<FileStorage>> GetFilesByCategoryAsync(int categoryId, string categoryType = "sub",
    int page = 1, int pageSize = 50, string orderBy = "created_at DESC")
        {
            string sql = @"
        SELECT 
            id AS Id,
            category_id AS CategoryId,
            file_attribute_id AS FileAttributeId,
            file_name AS FileName,
            file_stored_name AS FileStoredName,
            file_type AS FileType,
            is_tianzheng AS IsTianZheng,
            file_hash AS FileHash,
            display_name AS DisplayName,
            block_name AS BlockName,
            layer_name AS LayerName,
            color_index AS ColorIndex,
            scale AS Scale,
            file_path AS FilePath,
            preview_image_name AS PreviewImageName,
            preview_image_path AS PreviewImagePath,
            file_size AS FileSize,
            is_preview AS IsPreview,
            version AS Version,
            description AS Description,
            is_active AS IsActive,
            created_by AS CreatedBy,
            category_type AS CategoryType,
            title AS Title,
            keywords AS Keywords,
            is_public AS IsPublic,
            updated_by AS UpdatedBy,
            last_accessed_at AS LastAccessedAt,
            created_at AS CreatedAt,
            updated_at AS UpdatedAt
          FROM cad_file_storage
          WHERE category_id = @CategoryId 
          AND category_type = @CategoryType
          AND is_active = 1";
            try
            {
                if (!string.IsNullOrEmpty(orderBy))
                {
                    sql += $" ORDER BY {orderBy}";
                }
                sql += " LIMIT @offset, @pageSize";

                var parameters = new
                {
                    CategoryId = categoryId,
                    CategoryType = categoryType,
                    offset = (page - 1) * pageSize,
                    pageSize
                };

                using var connection = new MySqlConnection(_connectionString);
                return (await connection.QueryAsync<FileStorage>(sql, parameters)).AsList();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"获取分类下的文件时出错: {ex.Message}");
                return new List<FileStorage>();
            }
        }

        /// <summary>
        ///  获取分类下的所有文件
        /// </summary>
        /// <param name="categoryId">分类Id</param>
        /// <param name="categoryType">分类类型</param>
        /// <returns></returns>
        public async Task<List<FileStorage>> GetFilesByCategoryIdAsync(int categoryId, string categoryType)
        {
            const string sql = @"
                SELECT 
                 id AS Id,
                 category_id AS CategoryId,
                 category_type AS CategoryType,
                 file_name AS FileName,
                 file_stored_name AS FileStoredName,
                 file_path AS FilePath,
                 file_type AS FileType,
                 is_tianzheng AS IsTianZheng,
                 file_size AS FileSize,
                 file_hash AS FileHash,
                 display_name AS DisplayName,
                 block_name AS BlockName,
                 layer_name AS LayerName,
                 color_index AS ColorIndex,
                 scale AS Scale,
                 preview_image_name AS PreviewImageName,
                 preview_image_path AS PreviewImagePath,
                 description AS Description,
                 version AS Version,
                 is_preview AS IsPreview,
                 is_active AS IsActive,
                 created_by AS CreatedBy,
                 created_at AS CreatedAt,
                 updated_at AS UpdatedAt
             FROM cad_file_storage 
             WHERE category_id = @CategoryId 
               AND category_type = @CategoryType
             ORDER BY created_at DESC";

            try
            {
                using var connection = new MySqlConnection(_connectionString);
                var parameters = new Dictionary<string, object>();
                parameters.Add("CategoryId", categoryId);
                parameters.Add("CategoryType", categoryType);
                var result = await connection.QueryAsync<FileStorage>(sql, parameters);
                return result.AsList();
            }
            catch (Exception e)
            {
                Env.Editor.WriteMessage(e.Message);
                return new List<FileStorage>();
            }
        }

        /// <summary>
        /// 兼容旧调用：按子分类读取图元文件。
        /// </summary>
        public async Task<List<FileStorage>> GetFileStorageBySubcategoryIdAsync(int subcategoryId)
        {
            return await GetFilesByCategoryIdAsync(subcategoryId, "sub").ConfigureAwait(false);
        }

        /// <summary>
        /// 兼容旧调用：按文件主键读取单条文件记录。
        /// </summary>
        public async Task<FileStorage?> GetFileByIdAsync(int fileId)
        {
            const string sql = @"
                SELECT 
                    id AS Id,
                    category_id AS CategoryId,
                    file_attribute_id AS FileAttributeId,
                    file_name AS FileName,
                    file_stored_name AS FileStoredName,
                    display_name AS DisplayName,
                    file_type AS FileType,
                    is_tianzheng AS IsTianZheng,
                    file_hash AS FileHash,
                    block_name AS BlockName,
                    layer_name AS LayerName,
                    color_index AS ColorIndex,
                    scale AS Scale,
                    file_path AS FilePath,
                    preview_image_name AS PreviewImageName,
                    preview_image_path AS PreviewImagePath,
                    file_size AS FileSize,
                    is_preview AS IsPreview,
                    version AS Version,
                    description AS Description,
                    is_active AS IsActive,
                    created_by AS CreatedBy,
                    category_type AS CategoryType,
                    title AS Title,
                    keywords AS Keywords,
                    is_public AS IsPublic,
                    updated_by AS UpdatedBy,
                    last_accessed_at AS LastAccessedAt,
                    created_at AS CreatedAt,
                    updated_at AS UpdatedAt
                FROM cad_file_storage
                WHERE id = @Id
                LIMIT 1";

            try
            {
                using var connection = new MySqlConnection(_connectionString);
                return await connection.QueryFirstOrDefaultAsync<FileStorage>(sql, new { Id = fileId }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetFileByIdAsync 出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 按图元主键读取属性记录（JSON方案）
        /// 说明：内部复用 GetFileWithAttributeAsync(fileId)，保持旧调用兼容。
        /// </summary>
        /// <param name="fileStorageId">cad_file_storage.id</param>
        /// <returns>兼容旧模型的 FileAttribute；找不到返回 null</returns>
        public async Task<FileAttribute?> GetFileAttributeByGraphicIdAsync(int fileStorageId)
        {
            // 防御式校验，避免无效主键触发无意义查询
            if (fileStorageId <= 0)
            {
                return null;
            }

            try
            {
                // 复用“主表 + JSON属性表”聚合查询方法
                var result = await GetFileWithAttributeAsync(fileStorageId).ConfigureAwait(false);

                // 返回属性对象（内部已完成 JSON -> FileAttribute 的兼容映射）
                return result.Attribute;
            }
            catch (Exception ex)
            {
                // 记录日志，避免上层崩溃
                LogManager.Instance.LogInfo($"GetFileAttributeByGraphicIdAsync 出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 级联删除图元及其属性（新表：cad_block_attributes_json），并可选删除物理文件
        /// </summary>
        public async Task<bool> DeleteCadGraphicCascadeAsync(int fileId, bool physicalDelete = true)
        {
            try
            {
                using var connection = GetConnection();
                await connection.OpenAsync().ConfigureAwait(false);
                using var transaction = connection.BeginTransaction();

                // 先取主记录，便于后续删除物理文件
                var storage = await connection.QueryFirstOrDefaultAsync<FileStorage>(@"
SELECT 
    id AS Id,
    category_id AS CategoryId,
    category_type AS CategoryType,
    file_attribute_id AS FileAttributeId,
    file_name AS FileName,
    file_stored_name AS FileStoredName,
    display_name AS DisplayName,
    file_path AS FilePath,
    preview_image_path AS PreviewImagePath,
    is_active AS IsActive
FROM cad_file_storage
WHERE id = @Id
LIMIT 1", new { Id = fileId }, transaction).ConfigureAwait(false);

                if (storage == null)
                {
                    transaction.Rollback();
                    return false;
                }

                // 先删附属日志/标签等
                await connection.ExecuteAsync("DELETE FROM file_tags WHERE file_id = @Id", new { Id = fileId }, transaction).ConfigureAwait(false);
                await connection.ExecuteAsync("DELETE FROM file_access_logs WHERE file_id = @Id", new { Id = fileId }, transaction).ConfigureAwait(false);
                await connection.ExecuteAsync("DELETE FROM file_version_history WHERE file_id = @Id", new { Id = fileId }, transaction).ConfigureAwait(false);

                // 删除新属性表记录（核心）
                await connection.ExecuteAsync(
                    "DELETE FROM cad_block_attributes_json WHERE file_id = @Id",
                    new { Id = fileId },
                    transaction).ConfigureAwait(false);

                // 最后删主表
                int affected = await connection.ExecuteAsync(
                    "DELETE FROM cad_file_storage WHERE id = @Id",
                    new { Id = fileId },
                    transaction).ConfigureAwait(false);

                // 提交数据库事务
                transaction.Commit();

                // 按需删除物理文件
                if (physicalDelete)
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(storage.FilePath) && File.Exists(storage.FilePath))
                            File.Delete(storage.FilePath);
                    }
                    catch (Exception exFile)
                    {
                        LogManager.Instance.LogInfo($"删除图元文件失败: {exFile.Message}");
                    }

                    try
                    {
                        if (!string.IsNullOrWhiteSpace(storage.PreviewImagePath) && File.Exists(storage.PreviewImagePath))
                            File.Delete(storage.PreviewImagePath);
                    }
                    catch (Exception exPreview)
                    {
                        LogManager.Instance.LogInfo($"删除预览文件失败: {exPreview.Message}");
                    }
                }

                return affected > 0;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"DeleteCadGraphicCascadeAsync 出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 兼容旧调用：当前版本无单独统计表时仅执行一次计数校验与日志记录。
        /// </summary>
        public async Task UpdateCategoryStatisticsAsync(int categoryId, string categoryType)
        {
            try
            {
                using var connection = GetConnection();
                int count = await connection.QuerySingleOrDefaultAsync<int>(@"
                    SELECT COUNT(*)
                    FROM cad_file_storage
                    WHERE category_id = @CategoryId AND category_type = @CategoryType AND is_active = 1",
                    new { CategoryId = categoryId, CategoryType = categoryType }).ConfigureAwait(false);

                LogManager.Instance.LogInfo($"分类统计已刷新: categoryId={categoryId}, categoryType={categoryType}, count={count}");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"UpdateCategoryStatisticsAsync 出错: {ex.Message}");
            }
        }



        /// <summary>
        /// 根据文件扩展名获取分类下的文件
        /// </summary>
        public async Task<List<FileStorage>> GetFilesByCategoryAndExtensionAsync(int categoryId, string fileType)
        {
            const string sql = @"
        SELECT 
            id AS Id,
            category_id AS CategoryId,
            file_attribute_id AS FileAttributeId,
            file_name AS FileName,
            file_stored_name AS FileStoredName,
            file_type AS FileType,
            file_hash AS FileHash,
            display_name AS DisplayName,
            block_name AS BlockName,
            layer_name AS LayerName,
            color_index AS ColorIndex,
            scale AS Scale,
            file_path AS FilePath,
            preview_image_name AS PreviewImageName,
            preview_image_path AS PreviewImagePath,
            file_size AS FileSize,
            is_preview AS IsPreview,
            version AS Version,
            description AS Description,
            is_active AS IsActive,
            created_by AS CreatedBy,
            category_type AS CategoryType,
            title AS Title,
            keywords AS Keywords,
            is_public AS IsPublic,
            updated_by AS UpdatedBy,
            last_accessed_at AS LastAccessedAt,
            created_at AS CreatedAt,
            updated_at AS UpdatedAt
        FROM cad_file_storage 
        WHERE category_id = @categoryId 
          AND file_type = @fileType
          AND is_active = 1
        ORDER BY created_at DESC";
            try
            {
                using var connection = new MySqlConnection(_connectionString);
                return (await connection.QueryAsync<FileStorage>(sql, new { categoryId, fileType })).AsList();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"根据文件扩展名获取分类下的文件时出错: {ex.Message}");
                return new List<FileStorage>();
            }
        }

        /// <summary>
        /// 搜索文件（支持关键词搜索）
        /// </summary>
        public async Task<List<FileStorage>> SearchFilesAsync(string keyword, int? categoryId = null)
        {
            string sql = @"
        SELECT 
            id AS Id,
            category_id AS CategoryId,
            file_attribute_id AS FileAttributeId,
            file_name AS FileName,
            file_stored_name AS FileStoredName,
            file_type AS FileType,
            file_hash AS FileHash,
            display_name AS DisplayName,
            block_name AS BlockName,
            layer_name AS LayerName,
            color_index AS ColorIndex,
            scale AS Scale,
            file_path AS FilePath,
            preview_image_name AS PreviewImageName,
            preview_image_path AS PreviewImagePath,
            file_size AS FileSize,
            is_preview AS IsPreview,
            version AS Version,
            description AS Description,
            is_active AS IsActive,
            created_by AS CreatedBy,
            category_type AS CategoryType,
            title AS Title,
            keywords AS Keywords,
            is_public AS IsPublic,
            updated_by AS UpdatedBy,
            last_accessed_at AS LastAccessedAt,
            created_at AS CreatedAt,
            updated_at AS UpdatedAt
        FROM cad_file_storage 
        WHERE is_active = 1";
            try
            {
                var parameters = new Dictionary<string, object>();

                if (!string.IsNullOrEmpty(keyword))
                {
                    sql += @" AND (title LIKE @keyword 
         OR file_name LIKE @keyword 
         OR display_name LIKE @keyword 
         OR description LIKE @keyword 
         OR keywords LIKE @keyword)";
                    parameters.Add("keyword", $"%{keyword}%");
                }

                if (categoryId.HasValue)
                {
                    sql += " AND category_id = @categoryId";
                    parameters.Add("categoryId", categoryId.Value);
                }

                sql += " ORDER BY created_at DESC LIMIT 100";

                using var connection = new MySqlConnection(_connectionString);
                return (await connection.QueryAsync<FileStorage>(sql, parameters)).AsList();
            }
            catch (Exception ex)
            {
                Env.Editor.WriteMessage($"搜索文件时出错: {ex.Message}");
                return new List<FileStorage>();
            }

        }

        /// <summary>
        /// 获取文件主记录（按文件哈希）
        /// 说明：在不破坏旧签名的前提下，若 file_attribute_id 为空，会自动回填最新 config_name（来自 JSON 属性表）。
        /// </summary>
        public async Task<FileStorage> GetFileStorageAsync(string filehash)
        {
            // 防御式校验，避免空哈希查询
            if (string.IsNullOrWhiteSpace(filehash))
                return null;

            // 查询文件主表
            const string fileSql = @"
SELECT 
    id AS Id,
    category_id AS CategoryId,
    file_attribute_id AS FileAttributeId,
    file_name AS FileName,
    file_stored_name AS FileStoredName,
    display_name AS DisplayName,
    file_type AS FileType,
    file_hash AS FileHash,
    block_name AS BlockName,
    layer_name AS LayerName,
    color_index AS ColorIndex,
    scale AS Scale,
    file_path AS FilePath,
    preview_image_name AS PreviewImageName,
    preview_image_path AS PreviewImagePath,
    file_size AS FileSize,
    is_preview AS IsPreview,
    version AS Version,
    description AS Description,
    is_active AS IsActive,
    created_by AS CreatedBy,
    category_type AS CategoryType,
    title AS Title,
    keywords AS Keywords,
    is_public AS IsPublic,
    updated_by AS UpdatedBy,
    last_accessed_at AS LastAccessedAt,
    created_at AS CreatedAt,
    updated_at AS UpdatedAt
FROM cad_file_storage
WHERE file_hash = @FileHash
LIMIT 1;";

            // 当主表 file_attribute_id 为空时，回填 JSON 属性表中最新 config_name
            const string latestConfigSql = @"
SELECT config_name
FROM cad_block_attributes_json
WHERE file_id = @FileId
ORDER BY attr_id DESC
LIMIT 1;";

            try
            {
                // 创建数据库连接
                using var connection = new MySqlConnection(_connectionString);

                // 先查主表
                var fileStorageInfo = await connection.QueryFirstOrDefaultAsync<FileStorage>(
                    fileSql, new { FileHash = filehash }).ConfigureAwait(false);

                // 未查到主表记录，直接返回
                if (fileStorageInfo == null)
                    return null;

                // 若兼容字段为空，则尝试从 JSON 表回填配置名
                if (string.IsNullOrWhiteSpace(fileStorageInfo.FileAttributeId))
                {
                    var latestConfigName = await connection.QueryFirstOrDefaultAsync<string>(
                        latestConfigSql, new { FileId = fileStorageInfo.Id }).ConfigureAwait(false);

                    if (!string.IsNullOrWhiteSpace(latestConfigName))
                    {
                        fileStorageInfo.FileAttributeId = latestConfigName;
                    }
                }

                // 返回主记录
                return fileStorageInfo;
            }
            catch (Exception ex)
            {
                // 记录错误日志并返回空
                LogManager.Instance.LogInfo($"获取文件详细信息时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 新增：按文件哈希获取“主记录 + JSON 属性字典 + 配置名”
        /// </summary>
        public async Task<(FileStorage File, Dictionary<string, string> Attributes, string ConfigName)> GetFileStorageWithAttributesByHashAsync(
            string filehash,
            string preferredConfigName = null)
        {
            // 先拿主记录（复用现有方法）
            var file = await GetFileStorageAsync(filehash).ConfigureAwait(false);

            // 主记录不存在时，返回空元组内容
            if (file == null)
                return (null, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), string.Empty);

            // 准备查询属性 JSON 的 SQL（优先配置名）
            const string attrByConfigSql = @"
SELECT config_name AS ConfigName, attributes_json AS AttributesJson
FROM cad_block_attributes_json
WHERE file_id = @FileId
  AND config_name = @ConfigName
ORDER BY attr_id DESC
LIMIT 1;";

            // 兜底查询（取最新）
            const string attrLatestSql = @"
SELECT config_name AS ConfigName, attributes_json AS AttributesJson
FROM cad_block_attributes_json
WHERE file_id = @FileId
ORDER BY attr_id DESC
LIMIT 1;";

            try
            {
                // 创建数据库连接
                using var connection = new MySqlConnection(_connectionString);

                // 确定优先配置名（参数优先，其次主表 file_attribute_id）
                var configName = string.IsNullOrWhiteSpace(preferredConfigName)
                    ? (file.FileAttributeId ?? string.Empty)
                    : preferredConfigName.Trim();

                (string ConfigName, string AttributesJson)? row = null;

                // 优先按配置名查
                if (!string.IsNullOrWhiteSpace(configName))
                {
                    row = await connection.QueryFirstOrDefaultAsync<(string ConfigName, string AttributesJson)>(
                        attrByConfigSql, new { FileId = file.Id, ConfigName = configName }).ConfigureAwait(false);
                }

                // 兜底查最新
                if (row == null || string.IsNullOrWhiteSpace(row.Value.AttributesJson))
                {
                    row = await connection.QueryFirstOrDefaultAsync<(string ConfigName, string AttributesJson)>(
                        attrLatestSql, new { FileId = file.Id }).ConfigureAwait(false);
                }

                // 没有属性记录时返回空字典
                if (row == null || string.IsNullOrWhiteSpace(row.Value.AttributesJson))
                {
                    return (file, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), string.Empty);
                }

                // 反序列化 JSON -> 字典
                var dict = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(row.Value.AttributesJson)
                           ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                // 返回聚合结果
                return (file, dict, row.Value.ConfigName ?? string.Empty);
            }
            catch (Exception ex)
            {
                // 异常时记录日志并返回主记录+空字典
                LogManager.Instance.LogInfo($"GetFileStorageWithAttributesByHashAsync 出错: {ex.Message}");
                return (file, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), string.Empty);
            }
        }

        /// <summary>
        /// 获取文件属性的详细信息（JSON新方案）
        /// 说明：保持旧返回类型 FileAttribute，不改调用方；
        /// 实现流程：先按文件名在 cad_file_storage 找 file -> 再到 cad_block_attributes_json 取 JSON -> 反序列化回填 FileAttribute。
        /// </summary>
        public async Task<FileAttribute> GetFileAttributeAsync(string fileName)
        {
            // 防御式处理，防止空输入
            if (string.IsNullOrWhiteSpace(fileName))
                return null;

            // 提取文件名（带扩展名）和不带扩展名版本，用于模糊匹配
            string noExtName = Path.GetFileNameWithoutExtension(fileName) ?? string.Empty;
            string rawName = Path.GetFileName(fileName) ?? string.Empty;

            // 先在主表里查目标文件（优先按 file_name / display_name / title）
            const string fileSql = @"
SELECT
    id AS Id,
    category_id AS CategoryId,
    file_attribute_id AS FileAttributeId,
    file_name AS FileName,
    file_stored_name AS FileStoredName,
    display_name AS DisplayName,
    file_type AS FileType,
    file_hash AS FileHash,
    block_name AS BlockName,
    layer_name AS LayerName,
    color_index AS ColorIndex,
    scale AS Scale,
    file_path AS FilePath,
    preview_image_name AS PreviewImageName,
    preview_image_path AS PreviewImagePath,
    file_size AS FileSize,
    is_preview AS IsPreview,
    version AS Version,
    description AS Description,
    is_active AS IsActive,
    created_by AS CreatedBy,
    category_type AS CategoryType,
    title AS Title,
    keywords AS Keywords,
    is_public AS IsPublic,
    updated_by AS UpdatedBy,
    last_accessed_at AS LastAccessedAt,
    created_at AS CreatedAt,
    updated_at AS UpdatedAt
FROM cad_file_storage
WHERE is_active = 1
  AND (
        file_name LIKE CONCAT('%', @Name1, '%')
     OR file_name LIKE CONCAT('%', @Name2, '%')
     OR display_name LIKE CONCAT('%', @Name1, '%')
     OR display_name LIKE CONCAT('%', @Name2, '%')
     OR title LIKE CONCAT('%', @Name1, '%')
     OR title LIKE CONCAT('%', @Name2, '%')
  )
ORDER BY updated_at DESC, id DESC
LIMIT 1;";

            // 优先按 config_name=file_attribute_id 精确查 JSON 属性
            const string attrSqlByConfig = @"
SELECT
    attr_id AS AttrId,
    file_id AS FileId,
    config_name AS ConfigName,
    attributes_json AS AttributesJson,
    created_at AS CreatedAt,
    updated_at AS UpdatedAt
FROM cad_block_attributes_json
WHERE file_id = @FileId
  AND config_name = @ConfigName
ORDER BY attr_id DESC
LIMIT 1;";

            // 兜底按 file_id 取最新属性记录
            const string attrSqlByLatest = @"
SELECT
    attr_id AS AttrId,
    file_id AS FileId,
    config_name AS ConfigName,
    attributes_json AS AttributesJson,
    created_at AS CreatedAt,
    updated_at AS UpdatedAt
FROM cad_block_attributes_json
WHERE file_id = @FileId
ORDER BY attr_id DESC
LIMIT 1;";

            try
            {
                using var connection = new MySqlConnection(_connectionString);

                // 查询主表文件记录
                var file = await connection.QueryFirstOrDefaultAsync<FileStorage>(
                    fileSql,
                    new { Name1 = noExtName, Name2 = rawName }).ConfigureAwait(false);

                // 没有匹配到文件，直接返回 null
                if (file == null)
                    return null;

                // 查询 JSON 属性记录
                BlockAttributesJson jsonRow = null;
                var configName = Convert.ToString(file.FileAttributeId)?.Trim();

                if (!string.IsNullOrWhiteSpace(configName))
                {
                    jsonRow = await connection.QueryFirstOrDefaultAsync<BlockAttributesJson>(
                        attrSqlByConfig,
                        new { FileId = file.Id, ConfigName = configName }).ConfigureAwait(false);
                }

                if (jsonRow == null)
                {
                    jsonRow = await connection.QueryFirstOrDefaultAsync<BlockAttributesJson>(
                        attrSqlByLatest,
                        new { FileId = file.Id }).ConfigureAwait(false);
                }

                // 没有属性记录时，返回一个基础对象（兼容旧逻辑）
                if (jsonRow == null || string.IsNullOrWhiteSpace(jsonRow.AttributesJson))
                {
                    return new FileAttribute
                    {
                        FileStorageId = file.Id,
                        FileName = file.FileName,
                        FileAttributeId = file.FileAttributeId,
                        CreatedAt = file.CreatedAt,
                        UpdatedAt = file.UpdatedAt
                    };
                }

                // 反序列化 JSON 到字典
                var dict = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonRow.AttributesJson)
                           ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                // 创建返回对象，先回填基础字段
                var attr = new FileAttribute
                {
                    FileStorageId = file.Id,
                    FileName = file.FileName,
                    FileAttributeId = string.IsNullOrWhiteSpace(jsonRow.ConfigName) ? file.FileAttributeId : jsonRow.ConfigName,
                    CreatedAt = jsonRow.CreatedAt,
                    UpdatedAt = jsonRow.UpdatedAt
                };

                // 通过反射把字典按“属性名”回填到 FileAttribute 模型
                foreach (var p in typeof(FileAttribute).GetProperties())
                {
                    if (!p.CanWrite) continue;
                    if (!dict.TryGetValue(p.Name, out var raw)) continue;
                    if (string.IsNullOrWhiteSpace(raw)) continue;

                    try
                    {
                        var targetType = Nullable.GetUnderlyingType(p.PropertyType) ?? p.PropertyType;

                        if (targetType == typeof(string))
                        {
                            p.SetValue(attr, raw);
                            continue;
                        }

                        if (targetType == typeof(int))
                        {
                            if (int.TryParse(raw, out var v)) p.SetValue(attr, v);
                            continue;
                        }

                        if (targetType == typeof(long))
                        {
                            if (long.TryParse(raw, out var v)) p.SetValue(attr, v);
                            continue;
                        }

                        if (targetType == typeof(decimal))
                        {
                            if (decimal.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                                p.SetValue(attr, v);
                            else if (decimal.TryParse(raw, out var v2))
                                p.SetValue(attr, v2);
                            continue;
                        }

                        if (targetType == typeof(double))
                        {
                            if (double.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                                p.SetValue(attr, v);
                            else if (double.TryParse(raw, out var v2))
                                p.SetValue(attr, v2);
                            continue;
                        }

                        if (targetType == typeof(DateTime))
                        {
                            if (DateTime.TryParse(raw, out var v)) p.SetValue(attr, v);
                            continue;
                        }

                        if (targetType == typeof(bool))
                        {
                            if (bool.TryParse(raw, out var b))
                                p.SetValue(attr, b);
                            else if (raw == "1")
                                p.SetValue(attr, true);
                            else if (raw == "0")
                                p.SetValue(attr, false);
                        }
                    }
                    catch
                    {
                        // 单个字段解析失败不影响整体返回
                    }
                }

                return attr;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"获取文件属性失败(JSON): {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 获取文件详细信息（主表 + 新 JSON 属性表）
        /// 说明：保留旧返回签名 (FileStorage, FileAttribute)，内部把 JSON 反序列化后回填到 FileAttribute（兼容旧调用）。
        /// </summary>
        public async Task<(FileStorage File, FileAttribute Attribute)> GetFileWithAttributeAsync(int fileId)
        {
            // 查询文件主表信息
            const string fileSql = @"
SELECT
    id AS Id,
    category_id AS CategoryId,
    file_attribute_id AS FileAttributeId,
    file_name AS FileName,
    file_stored_name AS FileStoredName,
    display_name AS DisplayName,
    file_type AS FileType,
    file_hash AS FileHash,
    block_name AS BlockName,
    layer_name AS LayerName,
    color_index AS ColorIndex,
    scale AS Scale,
    file_path AS FilePath,
    preview_image_name AS PreviewImageName,
    preview_image_path AS PreviewImagePath,
    file_size AS FileSize,
    is_preview AS IsPreview,
    version AS Version,
    description AS Description,
    is_active AS IsActive,
    created_by AS CreatedBy,
    category_type AS CategoryType,
    title AS Title,
    keywords AS Keywords,
    is_public AS IsPublic,
    updated_by AS UpdatedBy,
    last_accessed_at AS LastAccessedAt,
    created_at AS CreatedAt,
    updated_at AS UpdatedAt
FROM cad_file_storage
WHERE id = @Id
LIMIT 1;";

            // 优先按 config_name = storage.file_attribute_id 查 JSON 属性
            const string attrSqlByConfig = @"
SELECT
    attr_id AS AttrId,
    file_id AS FileId,
    config_name AS ConfigName,
    attributes_json AS AttributesJson,
    created_at AS CreatedAt,
    updated_at AS UpdatedAt
FROM cad_block_attributes_json
WHERE file_id = @FileId
  AND config_name = @ConfigName
ORDER BY attr_id DESC
LIMIT 1;";

            // 若按 config_name 查不到，则按 file_id 取最新一条兜底
            const string attrSqlByLatest = @"
SELECT
    attr_id AS AttrId,
    file_id AS FileId,
    config_name AS ConfigName,
    attributes_json AS AttributesJson,
    created_at AS CreatedAt,
    updated_at AS UpdatedAt
FROM cad_block_attributes_json
WHERE file_id = @FileId
ORDER BY attr_id DESC
LIMIT 1;";

            try
            {
                // 创建数据库连接
                using var connection = new MySqlConnection(_connectionString);

                // 先取主表文件记录
                var file = await connection.QuerySingleOrDefaultAsync<FileStorage>(fileSql, new { Id = fileId }).ConfigureAwait(false);

                // 若文件不存在，直接返回空
                if (file == null)
                {
                    return (null, null);
                }

                // 准备读取 JSON 属性记录
                BlockAttributesJson jsonRow = null;

                // 兼容字段 file_attribute_id，此处作为 config_name 优先匹配
                var configName = Convert.ToString(file.FileAttributeId)?.Trim();

                // 优先按 config_name 查
                if (!string.IsNullOrWhiteSpace(configName))
                {
                    jsonRow = await connection.QuerySingleOrDefaultAsync<BlockAttributesJson>(
                        attrSqlByConfig,
                        new { FileId = file.Id, ConfigName = configName }).ConfigureAwait(false);
                }

                // 兜底按最新记录查
                if (jsonRow == null)
                {
                    jsonRow = await connection.QuerySingleOrDefaultAsync<BlockAttributesJson>(
                        attrSqlByLatest,
                        new { FileId = file.Id }).ConfigureAwait(false);
                }

                // 没有属性记录时，返回 file + null（兼容旧逻辑）
                if (jsonRow == null || string.IsNullOrWhiteSpace(jsonRow.AttributesJson))
                {
                    return (file, null);
                }

                // 反序列化 JSON 为键值字典
                var dict = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonRow.AttributesJson)
                           ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                // 构建兼容旧调用的 FileAttribute 对象
                var attr = new FileAttribute
                {
                    FileStorageId = file.Id,
                    FileName = file.FileName,
                    FileAttributeId = string.IsNullOrWhiteSpace(jsonRow.ConfigName) ? file.FileAttributeId : jsonRow.ConfigName,
                    CreatedAt = jsonRow.CreatedAt,
                    UpdatedAt = jsonRow.UpdatedAt
                };

                // 遍历 FileAttribute 可写属性，按属性名从字典回填
                foreach (var p in typeof(FileAttribute).GetProperties())
                {
                    // 只处理可写属性
                    if (!p.CanWrite) continue;

                    // 字典里没有同名键则跳过
                    if (!dict.TryGetValue(p.Name, out var raw)) continue;

                    // 空值跳过
                    if (string.IsNullOrWhiteSpace(raw)) continue;

                    try
                    {
                        // 处理可空类型
                        var targetType = Nullable.GetUnderlyingType(p.PropertyType) ?? p.PropertyType;

                        // 字符串直接赋值
                        if (targetType == typeof(string))
                        {
                            p.SetValue(attr, raw);
                            continue;
                        }

                        // 整型转换
                        if (targetType == typeof(int))
                        {
                            if (int.TryParse(raw, out var v)) p.SetValue(attr, v);
                            continue;
                        }

                        // 长整型转换
                        if (targetType == typeof(long))
                        {
                            if (long.TryParse(raw, out var v)) p.SetValue(attr, v);
                            continue;
                        }

                        // decimal 转换（用不变文化，兼容小数点）
                        if (targetType == typeof(decimal))
                        {
                            if (decimal.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                            {
                                p.SetValue(attr, v);
                            }
                            else if (decimal.TryParse(raw, out var v2))
                            {
                                p.SetValue(attr, v2);
                            }
                            continue;
                        }

                        // double 转换（用不变文化，兼容小数点）
                        if (targetType == typeof(double))
                        {
                            if (double.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                            {
                                p.SetValue(attr, v);
                            }
                            else if (double.TryParse(raw, out var v2))
                            {
                                p.SetValue(attr, v2);
                            }
                            continue;
                        }

                        // DateTime 转换
                        if (targetType == typeof(DateTime))
                        {
                            if (DateTime.TryParse(raw, out var v)) p.SetValue(attr, v);
                            continue;
                        }

                        // bool 转换（兼容 1/0 与 true/false）
                        if (targetType == typeof(bool))
                        {
                            if (bool.TryParse(raw, out var b))
                            {
                                p.SetValue(attr, b);
                            }
                            else if (raw == "1")
                            {
                                p.SetValue(attr, true);
                            }
                            else if (raw == "0")
                            {
                                p.SetValue(attr, false);
                            }
                        }
                    }
                    catch
                    {
                        // 单字段转换失败不影响整体读取
                    }
                }

                // 返回主表文件记录 + 兼容属性对象
                return (file, attr);
            }
            catch (Exception ex)
            {
                // 输出错误并返回空对象元组
                Env.Editor.WriteMessage($"\n获取文件详细信息时出错: {ex.Message}");
                return (null, null);
            }
        }

        /// <summary>
        /// ===== 新增：cad_file_attributes 全字段查询片段（供多个查询方法复用）=====
        /// </summary>
        private const string CadFileAttributeSelectColumns = @"
    id AS Id,
    category_id AS CategoryId,
    file_storage_id AS FileStorageId,
    file_name AS FileName,
    file_attribute_id AS FileAttributeId,
    description AS Description,
    attribute_group AS AttributeGroup,
    remarks AS Remarks,
    customize1 AS Customize1,
    customize2 AS Customize2,
    customize3 AS Customize3,
    length AS Length,
    width AS Width,
    height AS Height,
    angle AS Angle,
    base_point_x AS BasePointX,
    base_point_y AS BasePointY,
    base_point_z AS BasePointZ,
    model AS Model,
    specifications AS Specifications,
    material AS Material,
    medium_name AS MediumName,
    standard_number AS StandardNumber,
    pressure AS Pressure,
    temperature AS Temperature,
    pressure_rating AS PressureRating,
    operating_pressure AS OperatingPressure,
    operating_temperature AS OperatingTemperature,
    diameter AS Diameter,
    outer_diameter AS OuterDiameter,
    inner_diameter AS InnerDiameter,
    nominal_diameter AS NominalDiameter,
    thickness AS Thickness,
    weight AS Weight,
    density AS Density,
    volume AS Volume,
    flow AS Flow,
    velocity AS Velocity,
    lift AS Lift,
    power AS Power,
    voltage AS Voltage,
    current AS Current,
    frequency AS Frequency,
    conductivity AS Conductivity,
    moisture AS Moisture,
    humidity AS Humidity,
    vacuum AS Vacuum,
    radiation AS Radiation,
    pipe_spec AS PipeSpec,
    pipe_nominal_diameter AS PipeNominalDiameter,
    pipe_wall_thickness AS PipeWallThickness,
    pipe_pressure_class AS PipePressureClass,
    connection_type AS ConnectionType,
    pipe_slope AS PipeSlope,
    anticorrosion_treatment AS AnticorrosionTreatment,
    valve_model AS ValveModel,
    valve_body_material AS ValveBodyMaterial,
    valve_disc_material AS ValveDiscMaterial,
    valve_ball_material AS ValveBallMaterial,
    seal_material AS SealMaterial,
    drive_type AS DriveType,
    open_mode AS OpenMode,
    applicable_medium AS ApplicableMedium,
    flange_model AS FlangeModel,
    flange_type AS FlangeType,
    flange_face_type AS FlangeFaceType,
    flange_standard AS FlangeStandard,
    bolt_spec AS BoltSpec,
    reducer_spec AS ReducerSpec,
    reducer_large_dn AS ReducerLargeDn,
    reducer_small_dn AS ReducerSmallDn,
    reducer_wall_thickness_large AS ReducerWallThicknessLarge,
    reducer_wall_thickness_small AS ReducerWallThicknessSmall,
    reducer_connection_type AS ReducerConnectionType,
    reducer_conicity AS ReducerConicity,
    reducer_eccentric_direction AS ReducerEccentricDirection,
    reducer_applicable_medium AS ReducerApplicableMedium,
    reducer_anticorrosion AS ReducerAnticorrosion,
    pump_model AS PumpModel,
    pump_flow AS PumpFlow,
    pump_head AS PumpHead,
    pump_body_material AS PumpBodyMaterial,
    motor_power AS MotorPower,
    inlet_outlet_diameter AS InletOutletDiameter,
    rated_speed AS RatedSpeed,
    pump_applicable_medium AS PumpApplicableMedium,
    working_pressure AS WorkingPressure,
    protection_level AS ProtectionLevel,
    expansion_joint_model AS ExpansionJointModel,
    bellows_material AS BellowsMaterial,
    flange_or_nozzle_material AS FlangeOrNozzleMaterial,
    compensation_amount AS CompensationAmount,
    expansion_joint_connection_type AS ExpansionJointConnectionType,
    expansion_joint_medium AS ExpansionJointMedium,
    expansion_joint_working_temp AS ExpansionJointWorkingTemp,
    flue_gas_capacity AS FlueGasCapacity,
    desulfurization_efficiency AS DesulfurizationEfficiency,
    droplet_size AS DropletSize,
    spray_layer_count AS SprayLayerCount,
    chimney_spec AS ChimneySpec,
    chimney_diameter AS ChimneyDiameter,
    chimney_height AS ChimneyHeight,
    chimney_material AS ChimneyMaterial,
    chimney_thickness AS ChimneyThickness,
    outlet_wind_speed AS OutletWindSpeed,
    insulation_thickness AS InsulationThickness,
    support_type AS SupportType,
    flue_gas_temperature AS FlueGasTemperature,
    pressure_gauge_model AS PressureGaugeModel,
    thermometer_model AS ThermometerModel,
    filter_model AS FilterModel,
    check_valve_model AS CheckValveModel,
    sprinkler_model AS SprinklerModel,
    flow_meter_model AS FlowMeterModel,
    safety_valve_model AS SafetyValveModel,
    flexible_joint_model AS FlexibleJointModel,
    created_at AS CreatedAt,
    updated_at AS UpdatedAt";

        /// <summary>
        /// ===== 新增：统一构造参数（插入/更新复用）=====
        /// </summary>
        /// <param name="a"></param>
        /// <param name="includeId"></param>
        /// <returns></returns>
        private DynamicParameters BuildFileAttributeParameters(FileAttribute a, bool includeId)
        {
            // 防御式处理，避免空引用
            a ??= new FileAttribute();

            // 业务属性ID为空时自动生成，满足新表非空要求
            if (string.IsNullOrWhiteSpace(a.FileAttributeId))
            {
                a.FileAttributeId = Guid.NewGuid().ToString("N");
            }

            // 时间兜底
            if (a.CreatedAt == default) a.CreatedAt = DateTime.Now;
            if (a.UpdatedAt == default) a.UpdatedAt = DateTime.Now;

            var p = new DynamicParameters();
            if (includeId) p.Add("Id", a.Id);

            p.Add("CategoryId", a.CategoryId);
            p.Add("FileStorageId", a.FileStorageId);
            p.Add("FileName", a.FileName ?? string.Empty);
            p.Add("FileAttributeId", a.FileAttributeId);

            p.Add("Description", a.Description);
            p.Add("AttributeGroup", a.AttributeGroup);
            p.Add("Remarks", a.Remarks);
            p.Add("Customize1", a.Customize1);
            p.Add("Customize2", a.Customize2);
            p.Add("Customize3", a.Customize3);

            p.Add("Length", a.Length); p.Add("Width", a.Width); p.Add("Height", a.Height); p.Add("Angle", a.Angle);
            p.Add("BasePointX", a.BasePointX); p.Add("BasePointY", a.BasePointY); p.Add("BasePointZ", a.BasePointZ);

            p.Add("Model", a.Model); p.Add("Specifications", a.Specifications); p.Add("Material", a.Material);
            p.Add("MediumName", a.MediumName); p.Add("StandardNumber", a.StandardNumber);

            p.Add("Pressure", a.Pressure); p.Add("Temperature", a.Temperature); p.Add("PressureRating", a.PressureRating);
            p.Add("OperatingPressure", a.OperatingPressure); p.Add("OperatingTemperature", a.OperatingTemperature);

            p.Add("Diameter", a.Diameter); p.Add("OuterDiameter", a.OuterDiameter); p.Add("InnerDiameter", a.InnerDiameter);
            p.Add("NominalDiameter", a.NominalDiameter); p.Add("Thickness", a.Thickness); p.Add("Weight", a.Weight); p.Add("Density", a.Density);

            p.Add("Volume", a.Volume); p.Add("Flow", a.Flow); p.Add("Velocity", a.Velocity); p.Add("Lift", a.Lift);
            p.Add("Power", a.Power); p.Add("Voltage", a.Voltage); p.Add("Current", a.Current); p.Add("Frequency", a.Frequency);
            p.Add("Conductivity", a.Conductivity); p.Add("Moisture", a.Moisture); p.Add("Humidity", a.Humidity); p.Add("Vacuum", a.Vacuum); p.Add("Radiation", a.Radiation);

            p.Add("PipeSpec", a.PipeSpec); p.Add("PipeNominalDiameter", a.PipeNominalDiameter); p.Add("PipeWallThickness", a.PipeWallThickness);
            p.Add("PipePressureClass", a.PipePressureClass); p.Add("ConnectionType", a.ConnectionType); p.Add("PipeSlope", a.PipeSlope); p.Add("AnticorrosionTreatment", a.AnticorrosionTreatment);

            p.Add("ValveModel", a.ValveModel); p.Add("ValveBodyMaterial", a.ValveBodyMaterial); p.Add("ValveDiscMaterial", a.ValveDiscMaterial);
            p.Add("ValveBallMaterial", a.ValveBallMaterial); p.Add("SealMaterial", a.SealMaterial); p.Add("DriveType", a.DriveType);
            p.Add("OpenMode", a.OpenMode); p.Add("ApplicableMedium", a.ApplicableMedium);

            p.Add("FlangeModel", a.FlangeModel); p.Add("FlangeType", a.FlangeType); p.Add("FlangeFaceType", a.FlangeFaceType);
            p.Add("FlangeStandard", a.FlangeStandard); p.Add("BoltSpec", a.BoltSpec);

            p.Add("ReducerSpec", a.ReducerSpec); p.Add("ReducerLargeDn", a.ReducerLargeDn); p.Add("ReducerSmallDn", a.ReducerSmallDn);
            p.Add("ReducerWallThicknessLarge", a.ReducerWallThicknessLarge); p.Add("ReducerWallThicknessSmall", a.ReducerWallThicknessSmall);
            p.Add("ReducerConnectionType", a.ReducerConnectionType); p.Add("ReducerConicity", a.ReducerConicity);
            p.Add("ReducerEccentricDirection", a.ReducerEccentricDirection); p.Add("ReducerApplicableMedium", a.ReducerApplicableMedium); p.Add("ReducerAnticorrosion", a.ReducerAnticorrosion);

            p.Add("PumpModel", a.PumpModel); p.Add("PumpFlow", a.PumpFlow); p.Add("PumpHead", a.PumpHead); p.Add("PumpBodyMaterial", a.PumpBodyMaterial);
            p.Add("MotorPower", a.MotorPower); p.Add("InletOutletDiameter", a.InletOutletDiameter); p.Add("RatedSpeed", a.RatedSpeed);
            p.Add("PumpApplicableMedium", a.PumpApplicableMedium); p.Add("WorkingPressure", a.WorkingPressure); p.Add("ProtectionLevel", a.ProtectionLevel);

            p.Add("ExpansionJointModel", a.ExpansionJointModel); p.Add("BellowsMaterial", a.BellowsMaterial); p.Add("FlangeOrNozzleMaterial", a.FlangeOrNozzleMaterial);
            p.Add("CompensationAmount", a.CompensationAmount); p.Add("ExpansionJointConnectionType", a.ExpansionJointConnectionType);
            p.Add("ExpansionJointMedium", a.ExpansionJointMedium); p.Add("ExpansionJointWorkingTemp", a.ExpansionJointWorkingTemp);

            p.Add("FlueGasCapacity", a.FlueGasCapacity); p.Add("DesulfurizationEfficiency", a.DesulfurizationEfficiency); p.Add("DropletSize", a.DropletSize); p.Add("SprayLayerCount", a.SprayLayerCount);
            p.Add("ChimneySpec", a.ChimneySpec); p.Add("ChimneyDiameter", a.ChimneyDiameter); p.Add("ChimneyHeight", a.ChimneyHeight); p.Add("ChimneyMaterial", a.ChimneyMaterial);
            p.Add("ChimneyThickness", a.ChimneyThickness); p.Add("OutletWindSpeed", a.OutletWindSpeed); p.Add("InsulationThickness", a.InsulationThickness);
            p.Add("SupportType", a.SupportType); p.Add("FlueGasTemperature", a.FlueGasTemperature);

            p.Add("PressureGaugeModel", a.PressureGaugeModel); p.Add("ThermometerModel", a.ThermometerModel); p.Add("FilterModel", a.FilterModel);
            p.Add("CheckValveModel", a.CheckValveModel); p.Add("SprinklerModel", a.SprinklerModel); p.Add("FlowMeterModel", a.FlowMeterModel);
            p.Add("SafetyValveModel", a.SafetyValveModel); p.Add("FlexibleJointModel", a.FlexibleJointModel);

            p.Add("CreatedAt", a.CreatedAt);
            p.Add("UpdatedAt", a.UpdatedAt);
            return p;
        }


        /// <summary>
        /// 事务插入：先插入文件存储，再插入 JSON 属性（新表 cad_block_attributes_json）
        /// 说明：保留旧方法签名，兼容现有调用方；旧 FileAttribute 会被转换为字典后存入 JSON。
        /// </summary>
        public async Task<(int StorageId, int AttributeId)> AddFileStorageAndAttributeAsync(FileStorage storage, FileAttribute attribute)
        {
            // 参数保护，防止空引用
            if (storage == null) throw new ArgumentNullException(nameof(storage));
            if (attribute == null) throw new ArgumentNullException(nameof(attribute));

            // 业务属性ID兜底（继续保留，写回 cad_file_storage.file_attribute_id 供兼容旧逻辑）
            if (string.IsNullOrWhiteSpace(attribute.FileAttributeId))
            {
                attribute.FileAttributeId = Guid.NewGuid().ToString("N");
            }

            // 本地函数，把旧 FileAttribute 转成字典（只存有值字段，避免无意义空值）
            Dictionary<string, string> BuildAttrDict(FileAttribute src)
            {
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var props = typeof(FileAttribute).GetProperties();

                foreach (var p in props)
                {
                    // 跳过系统字段/主键字段，避免污染业务属性
                    if (string.Equals(p.Name, nameof(FileAttribute.Id), StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(p.Name, nameof(FileAttribute.CategoryId), StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(p.Name, nameof(FileAttribute.FileStorageId), StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(p.Name, nameof(FileAttribute.CreatedAt), StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(p.Name, nameof(FileAttribute.UpdatedAt), StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var v = p.GetValue(src);
                    if (v == null) continue;

                    var s = Convert.ToString(v);
                    if (string.IsNullOrWhiteSpace(s)) continue;

                    dict[p.Name] = s.Trim();
                }

                return dict;
            }

            // 把旧模型转换为 JSON 文本
            var attrDict = BuildAttrDict(attribute);
            var attrJson = Newtonsoft.Json.JsonConvert.SerializeObject(attrDict);

            using var connection = GetConnection();
            await connection.OpenAsync().ConfigureAwait(false);

            MySqlTransaction tx = null;
            int storageId = 0;

            try
            {
                // 开启事务，确保“主表+属性表+回写”原子性
                tx = connection.BeginTransaction();

                // 先插入文件主表
                const string insertStorageSql = @"
INSERT INTO cad_file_storage
(category_id, file_attribute_id, file_name, file_stored_name, display_name, file_type, is_tianzheng, file_hash,
 block_name, layer_name, color_index, scale, file_path, preview_image_name, preview_image_path, file_size,
 is_preview, version, description, is_active, created_by, category_type, title, keywords, is_public, updated_by,
 last_accessed_at, created_at, updated_at)
VALUES
(@CategoryId, NULL, @FileName, @FileStoredName, @DisplayName, @FileType, @IsTianZheng, @FileHash,
 @BlockName, @LayerName, @ColorIndex, @Scale, @FilePath, @PreviewImageName, @PreviewImagePath, @FileSize,
 @IsPreview, @Version, @Description, @IsActive, @CreatedBy, @CategoryType, @Title, @Keywords, @IsPublic, @UpdatedBy,
 @LastAccessedAt, @CreatedAt, @UpdatedAt);";

                await connection.ExecuteAsync(insertStorageSql, new
                {
                    storage.CategoryId,
                    FileName = storage.FileName ?? string.Empty,
                    FileStoredName = storage.FileStoredName ?? string.Empty,
                    DisplayName = storage.DisplayName ?? storage.FileName ?? string.Empty,
                    FileType = storage.FileType ?? string.Empty,
                    storage.IsTianZheng,
                    FileHash = storage.FileHash ?? string.Empty,
                    BlockName = storage.BlockName ?? string.Empty,
                    LayerName = storage.LayerName ?? string.Empty,
                    ColorIndex = storage.ColorIndex ?? 0,
                    Scale = storage.Scale ?? 1.0,
                    FilePath = storage.FilePath ?? string.Empty,
                    PreviewImageName = storage.PreviewImageName ?? string.Empty,
                    PreviewImagePath = storage.PreviewImagePath ?? string.Empty,
                    FileSize = storage.FileSize ?? 0L,
                    storage.IsPreview,
                    storage.Version,
                    Description = storage.Description ?? string.Empty,
                    storage.IsActive,
                    CreatedBy = storage.CreatedBy ?? Environment.UserName,
                    CategoryType = storage.CategoryType ?? "sub",
                    Title = storage.Title ?? string.Empty,
                    Keywords = storage.Keywords ?? string.Empty,
                    storage.IsPublic,
                    UpdatedBy = storage.UpdatedBy ?? string.Empty,
                    storage.LastAccessedAt,
                    CreatedAt = storage.CreatedAt == default ? DateTime.Now : storage.CreatedAt,
                    UpdatedAt = storage.UpdatedAt == default ? DateTime.Now : storage.UpdatedAt
                }, tx).ConfigureAwait(false);

                // 获取主表自增ID
                storageId = await connection.QuerySingleAsync<int>("SELECT LAST_INSERT_ID()", transaction: tx).ConfigureAwait(false);

                // 插入新属性表（JSON）
                const string insertJsonAttrSql = @"
INSERT INTO cad_block_attributes_json
(file_id, config_name, attributes_json, created_at, updated_at)
VALUES
(@FileId, @ConfigName, @AttributesJson, NOW(), NOW());";

                await connection.ExecuteAsync(insertJsonAttrSql, new
                {
                    FileId = storageId,
                    ConfigName = string.IsNullOrWhiteSpace(attribute.FileAttributeId) ? "default" : attribute.FileAttributeId,
                    AttributesJson = attrJson
                }, tx).ConfigureAwait(false);

                // 取 JSON 属性表主键（为了兼容原返回值 AttributeId）
                int attrId = await connection.QuerySingleAsync<int>("SELECT LAST_INSERT_ID()", transaction: tx).ConfigureAwait(false);

                // 回写 storage.file_attribute_id（保持兼容）
                const string updateStorageSql = @"
UPDATE cad_file_storage
SET file_attribute_id = @FileAttributeId, updated_at = @UpdatedAt
WHERE id = @Id;";

                await connection.ExecuteAsync(updateStorageSql, new
                {
                    FileAttributeId = attribute.FileAttributeId,
                    UpdatedAt = DateTime.Now,
                    Id = storageId
                }, tx).ConfigureAwait(false);

                // 提交事务
                tx.Commit();
                return (storageId, attrId);
            }
            catch (Exception ex)
            {
                // 异常时回滚
                try { tx?.Rollback(); } catch { }
                LogManager.Instance.LogInfo($"AddFileStorageAndAttributeAsync 失败: {ex.Message}");
                return (0, 0);
            }
            finally
            {
                // 释放事务对象
                try { tx?.Dispose(); } catch { }
            }
        }

        /// <summary>
        /// 新方案——插入文件主记录 + JSON属性记录（事务）
        /// </summary>
        /// <param name="storage">中文注释：文件主表对象</param>
        /// <param name="attributes">中文注释：动态属性字典</param>
        /// <param name="configName">中文注释：属性配置名</param>
        /// <returns>中文注释：返回主表ID与属性表ID</returns>
        /// <exception cref="ArgumentNullException">中文注释：主表对象为空时抛出</exception>
        public async Task<(int StorageId, long AttrId)> AddFileStorageAndAttributesJsonAsync(
            FileStorage storage,
            Dictionary<string, string> attributes,
            string configName = "default")
        {
            if (storage == null) throw new ArgumentNullException(nameof(storage));

            using var connection = GetConnection();
            await connection.OpenAsync().ConfigureAwait(false);
            using var tx = connection.BeginTransaction();

            try
            {
                // 先插入 cad_file_storage 主表
                const string insertStorageSql = @"
INSERT INTO cad_file_storage
(category_id, file_name, file_stored_name, display_name, file_type, file_hash, block_name, layer_name, color_index, scale, file_path, preview_image_name, preview_image_path, file_size, is_preview, version, description, is_active, created_by, category_type, title, keywords, is_public, updated_by, last_accessed_at, created_at, updated_at)
VALUES
(@CategoryId, @FileName, @FileStoredName, @DisplayName, @FileType, @FileHash, @BlockName, @LayerName, @ColorIndex, @Scale, @FilePath, @PreviewImageName, @PreviewImagePath, @FileSize, @IsPreview, @Version, @Description, @IsActive, @CreatedBy, @CategoryType, @Title, @Keywords, @IsPublic, @UpdatedBy, @LastAccessedAt, @CreatedAt, @UpdatedAt);";

                await connection.ExecuteAsync(insertStorageSql, new
                {
                    storage.CategoryId,
                    FileName = storage.FileName ?? string.Empty,
                    FileStoredName = storage.FileStoredName ?? string.Empty,
                    DisplayName = storage.DisplayName ?? storage.FileName ?? string.Empty,
                    FileType = storage.FileType ?? string.Empty,
                    FileHash = storage.FileHash ?? string.Empty,
                    BlockName = storage.BlockName ?? string.Empty,
                    LayerName = storage.LayerName ?? string.Empty,
                    ColorIndex = storage.ColorIndex ?? 0,
                    Scale = storage.Scale ?? 1.0,
                    FilePath = storage.FilePath ?? string.Empty,
                    PreviewImageName = storage.PreviewImageName ?? string.Empty,
                    PreviewImagePath = storage.PreviewImagePath ?? string.Empty,
                    FileSize = storage.FileSize ?? 0L,
                    storage.IsPreview,
                    storage.Version,
                    Description = storage.Description ?? string.Empty,
                    storage.IsActive,
                    CreatedBy = storage.CreatedBy ?? Environment.UserName,
                    CategoryType = storage.CategoryType ?? "sub",
                    Title = storage.Title ?? string.Empty,
                    Keywords = storage.Keywords ?? string.Empty,
                    storage.IsPublic,
                    UpdatedBy = storage.UpdatedBy ?? string.Empty,
                    storage.LastAccessedAt,
                    CreatedAt = storage.CreatedAt == default ? DateTime.Now : storage.CreatedAt,
                    UpdatedAt = storage.UpdatedAt == default ? DateTime.Now : storage.UpdatedAt
                }, tx).ConfigureAwait(false);

                // 获取主表自增ID
                int storageId = await connection.QuerySingleAsync<int>("SELECT LAST_INSERT_ID()", transaction: tx).ConfigureAwait(false);

                // 插入 JSON 属性表
                const string insertAttrSql = @"
INSERT INTO cad_block_attributes_json (file_id, config_name, attributes_json, created_at, updated_at)
VALUES (@FileId, @ConfigName, @AttributesJson, NOW(), NOW());";

                await connection.ExecuteAsync(insertAttrSql, new
                {
                    FileId = storageId,
                    ConfigName = string.IsNullOrWhiteSpace(configName) ? "default" : configName,
                    AttributesJson = ToJson(attributes)
                }, tx).ConfigureAwait(false);

                // 获取属性表自增ID
                long attrId = await connection.QuerySingleAsync<long>("SELECT LAST_INSERT_ID()", transaction: tx).ConfigureAwait(false);

                // 提交事务
                tx.Commit();
                return (storageId, attrId);
            }
            catch
            {
                // 异常回滚事务
                tx.Rollback();
                throw;
            }
        }

        #endregion

        #region 管道操作方法

        /// <summary>
        /// 管道模板查询结果（主记录 + JSON属性）
        /// </summary>
        public sealed class PipeTemplateQueryResult
        {
            /// <summary>
            /// 模板主表记录
            /// </summary>
            public FileStorage? Storage { get; set; }

            /// <summary>
            /// 模板属性字典（来自 cad_block_attributes_json）
            /// </summary>
            public Dictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 按入口/出口从数据库获取“最新可用管道模板”及其属性字典
        /// 规则：优先匹配 display_name/file_name/block_name/title/keywords，按 updated_at/id 倒序取一条
        /// </summary>
        /// <param name="isOutlet">中文注释：true=出口模板，false=入口模板</param>
        /// <returns>中文注释：命中时返回模板与属性，未命中返回 null</returns>
        public async Task<PipeTemplateQueryResult?> GetLatestPipeTemplateWithAttributesAsync(bool isOutlet)
        {
            try
            {
                // 根据入口/出口准备关键词
                string kwCnMain = isOutlet ? "出口管道" : "入口管道";
                string kwCnSub = isOutlet ? "出口" : "入口";
                string kwEn = isOutlet ? "outlet" : "inlet";

                // 模糊匹配参数
                string k1 = $"%{kwCnMain}%";
                string k2 = $"%{kwCnSub}%";
                string k3 = $"%{kwEn}%";

                // 查询模板主记录（只取激活记录）
                const string sqlTemplate = @"
                              SELECT
                                  id AS Id,
                                  category_id AS CategoryId,
                                  category_type AS CategoryType,
                                  file_attribute_id AS FileAttributeId,
                                  file_name AS FileName,
                                  file_stored_name AS FileStoredName,
                                  display_name AS DisplayName,
                                  file_type AS FileType,
                                  file_hash AS FileHash,
                                  block_name AS BlockName,
                                  layer_name AS LayerName,
                                  color_index AS ColorIndex,
                                  scale AS Scale,
                                  file_path AS FilePath,
                                  preview_image_name AS PreviewImageName,
                                  preview_image_path AS PreviewImagePath,
                                  file_size AS FileSize,
                                  is_preview AS IsPreview,
                                  version AS Version,
                                  description AS Description,
                                  is_active AS IsActive,
                                  created_by AS CreatedBy,
                                  title AS Title,
                                  keywords AS Keywords,
                                  is_public AS IsPublic,
                                  updated_by AS UpdatedBy,
                                  last_accessed_at AS LastAccessedAt,
                                  created_at AS CreatedAt,
                                  updated_at AS UpdatedAt
                              FROM cad_file_storage
                              WHERE is_active = 1
                                AND (
                                    display_name LIKE @K1 OR display_name LIKE @K2 OR display_name LIKE @K3
                                    OR file_name LIKE @K1 OR file_name LIKE @K2 OR file_name LIKE @K3
                                    OR block_name LIKE @K1 OR block_name LIKE @K2 OR block_name LIKE @K3
                                    OR title LIKE @K1 OR title LIKE @K2 OR title LIKE @K3
                                    OR keywords LIKE @K1 OR keywords LIKE @K2 OR keywords LIKE @K3
                                )
                              ORDER BY updated_at DESC, id DESC
                              LIMIT 1;";

                // 建立连接并查询主记录
                using var connection = GetConnection();
                var storage = await connection.QueryFirstOrDefaultAsync<FileStorage>(
                    sqlTemplate,
                    new { K1 = k1, K2 = k2, K3 = k3 }).ConfigureAwait(false);

                // 未命中模板直接返回 null
                if (storage == null)
                {
                    return null;
                }

                // 查询该模板最新属性JSON
                const string sqlAttr = @"
                             SELECT attributes_json
                             FROM cad_block_attributes_json
                             WHERE file_id = @FileId
                             ORDER BY updated_at DESC, attr_id DESC
                             LIMIT 1;";

                string? attrJson = await connection.QueryFirstOrDefaultAsync<string>(
                    sqlAttr,
                    new { FileId = storage.Id }).ConfigureAwait(false);

                // 反序列化属性字典（空则返回空字典）
                var attrs = string.IsNullOrWhiteSpace(attrJson)
                    ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    : (Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(attrJson)
                       ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));

                // 返回统一结果对象
                return new PipeTemplateQueryResult
                {
                    Storage = storage,
                    Attributes = attrs
                };
            }
            catch (Exception ex)
            {
                // 记录异常日志，返回 null 由上层兜底
                LogManager.Instance.LogInfo($"GetLatestPipeTemplateWithAttributesAsync 出错: {ex.Message}");
                return null;
            }
        }

        #endregion



        /// <summary>
        /// CAD分类
        /// </summary>
        public class CadCategory
        {
            public int Id { get; set; } // 分类ID
            public string Name { get; set; } = string.Empty; // 分类编码名
            public string DisplayName { get; set; } = string.Empty; // 分类显示名
            public string SubcategoryIds { get; set; } = string.Empty; // 子分类ID串
            public int SortOrder { get; set; } // 排序
            public DateTime CreatedAt { get; set; } // 创建时间
            public DateTime UpdatedAt { get; set; } // 更新时间
        }

        /// <summary>
        /// CAD子分类
        /// </summary>
        public class CadSubcategory
        {
            public int Id { get; set; } // 子分类ID
            public int ParentId { get; set; } // 父分类ID
            public string Name { get; set; } = string.Empty; // 子分类编码名
            public string DisplayName { get; set; } = string.Empty; // 子分类显示名
            public string SubcategoryIds { get; set; } = string.Empty; // 下级子分类ID串
            public int SortOrder { get; set; } // 排序
            public int Level { get; set; } // 层级
            public DateTime CreatedAt { get; set; } // 创建时间
            public DateTime UpdatedAt { get; set; } // 更新时间
        }

        /// <summary>
        /// SW分类
        /// </summary>
        public class SwCategory
        {
            public int Id { get; set; } // 分类ID
            public string Name { get; set; } = string.Empty; // 分类名
            public string DisplayName { get; set; } = string.Empty; // 显示名
            public int SortOrder { get; set; } // 排序
            public DateTime CreatedAt { get; set; } // 创建时间
            public DateTime UpdatedAt { get; set; } // 更新时间
        }

        /// <summary>
        /// SW子分类
        /// </summary>
        public class SwSubcategory
        {
            public int Id { get; set; } // 子分类ID
            public int ParentId { get; set; } // 父ID
            public int CategoryId { get; set; } // 分类ID
            public string Name { get; set; } = string.Empty; // 名称
            public string DisplayName { get; set; } = string.Empty; // 显示名
            public int SortOrder { get; set; } // 排序
            public DateTime CreatedAt { get; set; } // 创建时间
            public DateTime UpdatedAt { get; set; } // 更新时间
        }

        /// <summary>
        /// SW图元
        /// </summary>
        public class SwGraphic
        {
            public int Id { get; set; } // 图元ID
            public int SubcategoryId { get; set; } // 子分类ID
            public string FileName { get; set; } = string.Empty; // 文件名
            public string DisplayName { get; set; } = string.Empty; // 显示名
            public string FilePath { get; set; } = string.Empty; // 文件路径
            public string PreviewImagePath { get; set; } = string.Empty; // 预览图路径
            public long FileSize { get; set; } // 文件大小
            public DateTime CreatedAt { get; set; } // 创建时间
            public DateTime UpdatedAt { get; set; } // 更新时间
        }           

        /// <summary>
        /// 文件存储（cad_file_storage）
        /// </summary>
        public class FileStorage
        {
            public int Id { get; set; } // 主键ID
            public int CategoryId { get; set; } // 分类ID
            public string? CategoryType { get; set; } // 分类类型
            public string? FileAttributeId { get; set; } // 属性业务ID（修复为string）
            public string? FileName { get; set; } // 文件名
            public string? FileStoredName { get; set; } // 存储文件名
            public string? DisplayName { get; set; } // 显示名
            public string? FileType { get; set; } // 文件类型
            public int? IsTianZheng { get; set; } // 是否天正
            public string? FileHash { get; set; } // 文件哈希
            public string? BlockName { get; set; } // 块名
            public string? LayerName { get; set; } // 图层名
            public int? ColorIndex { get; set; } // 颜色索引
            public double? Scale { get; set; } // 比例
            public string? FilePath { get; set; } // 文件路径
            public string? PreviewImageName { get; set; } // 预览图名
            public string? PreviewImagePath { get; set; } // 预览图路径
            public long? FileSize { get; set; } // 文件大小
            public int IsPreview { get; set; } // 是否预览
            public int Version { get; set; } // 版本
            public string? Description { get; set; } // 描述
            public int IsActive { get; set; } // 是否启用
            public string? CreatedBy { get; set; } // 创建人
            public string? Title { get; set; } // 标题
            public string? Keywords { get; set; } // 关键字
            public int IsPublic { get; set; } // 是否公开
            public string? UpdatedBy { get; set; } // 更新人
            public DateTime? LastAccessedAt { get; set; } // 最后访问时间
            public DateTime CreatedAt { get; set; } // 创建时间
            public DateTime UpdatedAt { get; set; } // 更新时间
        }

        /// <summary>
        /// 文件属性（cad_file_attributes）
        /// 说明：表中 id/category_id/file_storage_id 是 BIGINT UNSIGNED，模型改为 long 兼容大数据量
        /// </summary>
        public class FileAttribute
        {
            // 主键ID（BIGINT）
            public long Id { get; set; }

            // 分类ID（BIGINT，可空）
            public long? CategoryId { get; set; }

            // 文件存储ID（BIGINT）
            public long FileStorageId { get; set; }

            // 文件名
            public string? FileName { get; set; }

            // 业务属性ID（VARCHAR(64)）
            public string? FileAttributeId { get; set; }

            public string? Description { get; set; } // 描述
            public string? AttributeGroup { get; set; } // 属性分组
            public string? Remarks { get; set; } // 备注
            public string? Customize1 { get; set; } // 自定义1
            public string? Customize2 { get; set; } // 自定义2
            public string? Customize3 { get; set; } // 自定义3

            public decimal? Length { get; set; }
            public decimal? Width { get; set; }
            public decimal? Height { get; set; }
            public decimal? Angle { get; set; }
            public decimal? BasePointX { get; set; }
            public decimal? BasePointY { get; set; }
            public decimal? BasePointZ { get; set; }

            public string? Model { get; set; }
            public string? Specifications { get; set; }
            public string? Material { get; set; }
            public string? MediumName { get; set; }
            public string? StandardNumber { get; set; }
            public decimal? Pressure { get; set; }
            public decimal? Temperature { get; set; }
            public string? PressureRating { get; set; }
            public decimal? OperatingPressure { get; set; }
            public decimal? OperatingTemperature { get; set; }
            public decimal? Diameter { get; set; }
            public decimal? OuterDiameter { get; set; }
            public decimal? InnerDiameter { get; set; }
            public string? NominalDiameter { get; set; }
            public decimal? Thickness { get; set; }
            public decimal? Weight { get; set; }
            public decimal? Density { get; set; }
            public decimal? Volume { get; set; }
            public decimal? Flow { get; set; }
            public decimal? Velocity { get; set; }
            public decimal? Lift { get; set; }
            public decimal? Power { get; set; }
            public decimal? Voltage { get; set; }
            public decimal? Current { get; set; }
            public decimal? Frequency { get; set; }
            public decimal? Conductivity { get; set; }
            public decimal? Moisture { get; set; }
            public decimal? Humidity { get; set; }
            public decimal? Vacuum { get; set; }
            public decimal? Radiation { get; set; }

            public string? PipeSpec { get; set; }
            public string? PipeNominalDiameter { get; set; }
            public decimal? PipeWallThickness { get; set; }
            public string? PipePressureClass { get; set; }
            public string? ConnectionType { get; set; }
            public string? PipeSlope { get; set; }
            public string? AnticorrosionTreatment { get; set; }

            public string? ValveModel { get; set; }
            public string? ValveBodyMaterial { get; set; }
            public string? ValveDiscMaterial { get; set; }
            public string? ValveBallMaterial { get; set; }
            public string? SealMaterial { get; set; }
            public string? DriveType { get; set; }
            public string? OpenMode { get; set; }
            public string? ApplicableMedium { get; set; }

            public string? FlangeModel { get; set; }
            public string? FlangeType { get; set; }
            public string? FlangeFaceType { get; set; }
            public string? FlangeStandard { get; set; }
            public string? BoltSpec { get; set; }

            public string? ReducerSpec { get; set; }
            public string? ReducerLargeDn { get; set; }
            public string? ReducerSmallDn { get; set; }
            public decimal? ReducerWallThicknessLarge { get; set; }
            public decimal? ReducerWallThicknessSmall { get; set; }
            public string? ReducerConnectionType { get; set; }
            public string? ReducerConicity { get; set; }
            public string? ReducerEccentricDirection { get; set; }
            public string? ReducerApplicableMedium { get; set; }
            public string? ReducerAnticorrosion { get; set; }

            public string? PumpModel { get; set; }
            public decimal? PumpFlow { get; set; }
            public decimal? PumpHead { get; set; }
            public string? PumpBodyMaterial { get; set; }
            public decimal? MotorPower { get; set; }
            public string? InletOutletDiameter { get; set; }
            public decimal? RatedSpeed { get; set; }
            public string? PumpApplicableMedium { get; set; }
            public decimal? WorkingPressure { get; set; }
            public string? ProtectionLevel { get; set; }

            public string? ExpansionJointModel { get; set; }
            public string? BellowsMaterial { get; set; }
            public string? FlangeOrNozzleMaterial { get; set; }
            public decimal? CompensationAmount { get; set; }
            public string? ExpansionJointConnectionType { get; set; }
            public string? ExpansionJointMedium { get; set; }
            public decimal? ExpansionJointWorkingTemp { get; set; }

            public decimal? FlueGasCapacity { get; set; }
            public decimal? DesulfurizationEfficiency { get; set; }
            public decimal? DropletSize { get; set; }
            public int? SprayLayerCount { get; set; }

            public string? ChimneySpec { get; set; }
            public decimal? ChimneyDiameter { get; set; }
            public decimal? ChimneyHeight { get; set; }
            public string? ChimneyMaterial { get; set; }
            public decimal? ChimneyThickness { get; set; }
            public decimal? OutletWindSpeed { get; set; }
            public decimal? InsulationThickness { get; set; }
            public string? SupportType { get; set; }
            public decimal? FlueGasTemperature { get; set; }

            public string? PressureGaugeModel { get; set; }
            public string? ThermometerModel { get; set; }
            public string? FilterModel { get; set; }
            public string? CheckValveModel { get; set; }
            public string? SprinklerModel { get; set; }
            public string? FlowMeterModel { get; set; }
            public string? SafetyValveModel { get; set; }
            public string? FlexibleJointModel { get; set; }

            public DateTime CreatedAt { get; set; }
            public DateTime UpdatedAt { get; set; }
        }

        /// <summary>
        /// JSON属性表模型（对应 cad_block_attributes_json）
        /// </summary>
        public class BlockAttributesJson
        {
            public long AttrId { get; set; } // 属性记录主键
            public int FileId { get; set; } // 关联 cad_file_storage.id
            public string? ConfigName { get; set; } // 配置名
            public string? AttributesJson { get; set; } // JSON文本
            public DateTime CreatedAt { get; set; } // 创建时间
            public DateTime UpdatedAt { get; set; } // 更新时间
        }

        /// <summary>
        /// 将字典转成JSON字符串
        /// </summary>
        /// <param name="dict"></param>
        /// <returns></returns>
        private static string ToJson(Dictionary<string, string>? dict)
        {
            return JsonConvert.SerializeObject(dict ?? new Dictionary<string, string>());
        }

        /// <summary>
        /// 将JSON字符串转成字典
        /// </summary>
        /// <param name="json"></param>
        /// <returns></returns>
        private static Dictionary<string, string> FromJson(string? json)
        {
            if (string.IsNullOrWhiteSpace(json)) return new Dictionary<string, string>();
            return JsonConvert.DeserializeObject<Dictionary<string, string>>(json)
                   ?? new Dictionary<string, string>();
        }

        /// <summary>
        /// 按 file_id 获取JSON属性（返回字典）
        /// </summary>
        /// <param name="fileId"></param>
        /// <param name="configName"></param>
        /// <returns></returns>
        public async Task<Dictionary<string, string>> GetAttributesJsonByFileIdAsync(int fileId, string configName = "default")
        {
            const string sql = @"
SELECT attributes_json
FROM cad_block_attributes_json
WHERE file_id = @FileId AND config_name = @ConfigName
ORDER BY attr_id DESC
LIMIT 1;";

            using var connection = GetConnection();
            var json = await connection.QueryFirstOrDefaultAsync<string>(sql, new { FileId = fileId, ConfigName = configName }).ConfigureAwait(false);
            return FromJson(json);
        }

    }
}
