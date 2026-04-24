using Dapper;
using GB_NewCadPlus_IV.UniFiedStandards;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            // 占位：返回模拟的插入结果（>0 表示成功）
            return 1;
        }

        //public virtual async Task<FileAttribute?> GetFileAttributeAsync(string displayName)
        //{
        //    await Task.Yield();
        //    return null;
        //}

        public virtual async Task<int> AddFileStorageAsync(FileStorage storage)
        {
            await Task.Yield();
            return 1;
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

        //public virtual async Task<FileStorage?> GetFileByIdAsync(int fileId)
        //{
        //    await Task.Yield();
        //    return null;
        //}

        //public virtual async Task UpdateCategoryStatisticsAsync(int categoryId, string categoryType)
        //{
        //    await Task.Yield();
        //    // 占位：不返回值
        //}


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
                    "cad_file_attributes",
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
        /// 创建数据库（若不存在）并创建缺失的核心表。
        /// 返回 true 表示成功（即创建完毕或已存在），false 表示失败。
        /// 注意：该方法会在服务器上执行 DDL，请确保凭据具有相应权限。
        /// </summary>
        public static bool CreateDatabaseAndCoreTables(string server, int port, string user, string password, string database = "cad_sw_library")
        {
            try
            {
                // 1) 先连接到服务端，确保数据库存在
                var masterConn = $"Server={server};Port={port};Uid={user};Pwd={password};";
                using (var conn = new MySqlConnection(masterConn))
                {
                    conn.Open();
                    var createDbSql = $"CREATE DATABASE IF NOT EXISTS `{database}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;";
                    conn.Execute(createDbSql);
                }

                // 2) 连接目标数据库，创建核心表
                var dbConn = $"Server={server};Port={port};Database={database};Uid={user};Pwd={password};";
                using (var conn = new MySqlConnection(dbConn))
                {
                    conn.Open();
                    var sql = new StringBuilder();

                    // cad_categories：修复为自增主键，匹配 AddCadCategoryAsync 插入方式
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `cad_categories` (
                        `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                        `name` VARCHAR(200) NOT NULL,
                        `display_name` VARCHAR(200),
                        `subcategory_ids` TEXT,
                        `sort_order` INT DEFAULT 0,
                        `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
                        `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // cad_subcategories：保持非自增（你的代码中有手工传入 Id 的逻辑）
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `cad_subcategories` (
                        `id` INT NOT NULL PRIMARY KEY,
                        `parent_id` INT NOT NULL,
                        `name` VARCHAR(200) NOT NULL,
                        `display_name` VARCHAR(200),
                        `sort_order` INT DEFAULT 0,
                        `level` INT DEFAULT 1,
                        `subcategory_ids` TEXT,
                        `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
                        `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                        INDEX(`parent_id`)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // cad_file_storage：补齐字段 + file_attribute_id 改为业务ID字符串
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `cad_file_storage` (
                        `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                        `category_id` INT NULL,
                        `category_type` VARCHAR(16) DEFAULT 'sub',
                        `file_attribute_id` VARCHAR(64) NULL COMMENT '关联 cad_file_attributes.file_attribute_id（业务ID）',

                        `file_name` VARCHAR(512) NULL,
                        `file_stored_name` VARCHAR(512) NULL,
                        `display_name` VARCHAR(512) NULL,
                        `file_type` VARCHAR(32) NULL,
                        `is_tianzheng` TINYINT NULL,
                        `file_hash` VARCHAR(128) NULL,

                        `block_name` VARCHAR(255) NULL,
                        `layer_name` VARCHAR(255) NULL,
                        `color_index` INT NULL,
                        `scale` DOUBLE DEFAULT 1.0,

                        `file_path` VARCHAR(1024) NULL,
                        `preview_image_name` VARCHAR(512) NULL,
                        `preview_image_path` VARCHAR(1024) NULL,

                        `file_size` BIGINT NULL,
                        `is_preview` TINYINT DEFAULT 0,
                        `version` INT DEFAULT 1,
                        `description` TEXT NULL,
                        `is_active` TINYINT DEFAULT 1,

                        `created_by` VARCHAR(128) NULL,
                        `title` VARCHAR(512) NULL,
                        `keywords` VARCHAR(1024) NULL,
                        `is_public` TINYINT DEFAULT 1,
                        `updated_by` VARCHAR(128) NULL,
                        `last_accessed_at` DATETIME NULL,

                        `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
                        `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,

                        KEY `idx_cfs_category` (`category_id`,`category_type`),
                        KEY `idx_cfs_attr_biz` (`file_attribute_id`),
                        KEY `idx_cfs_file_hash` (`file_hash`)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // cad_file_attributes：保留扩展字段
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `cad_file_attributes` (
                        `id` BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '主键ID',
                        `category_id` BIGINT UNSIGNED NULL COMMENT '分类ID',
                        `file_storage_id` BIGINT UNSIGNED NOT NULL COMMENT '文件存储ID',
                        `file_name` VARCHAR(255) NOT NULL COMMENT '文件名',
                        `file_attribute_id` VARCHAR(64) NOT NULL COMMENT '属性业务ID',
                        `description` TEXT NULL COMMENT '描述',
                        `attribute_group` VARCHAR(50) NULL COMMENT '属性分组',
                        `remarks` TEXT NULL COMMENT '备注',
                        `customize1` VARCHAR(255) NULL COMMENT '自定义1',
                        `customize2` VARCHAR(255) NULL COMMENT '自定义2',
                        `customize3` VARCHAR(255) NULL COMMENT '自定义3',
                        `length` DECIMAL(18,6) NULL COMMENT '长度',
                        `width` DECIMAL(18,6) NULL COMMENT '宽度',
                        `height` DECIMAL(18,6) NULL COMMENT '高度',
                        `angle` DECIMAL(18,6) NULL COMMENT '角度',
                        `base_point_x` DECIMAL(18,6) NULL COMMENT '基点X',
                        `base_point_y` DECIMAL(18,6) NULL COMMENT '基点Y',
                        `base_point_z` DECIMAL(18,6) NULL COMMENT '基点Z',
                        `model` VARCHAR(255) NULL COMMENT '型号',
                        `specifications` VARCHAR(255) NULL COMMENT '规格',
                        `material` VARCHAR(100) NULL COMMENT '材质',
                        `medium_name` VARCHAR(100) NULL COMMENT '介质',
                        `standard_number` VARCHAR(100) NULL COMMENT '标准编号',
                        `pressure` DECIMAL(18,6) NULL COMMENT '压力',
                        `temperature` DECIMAL(18,6) NULL COMMENT '温度',
                        `pressure_rating` VARCHAR(50) NULL COMMENT '压力等级',
                        `operating_pressure` DECIMAL(18,6) NULL COMMENT '操作压力',
                        `operating_temperature` DECIMAL(18,6) NULL COMMENT '操作温度',
                        `diameter` DECIMAL(18,6) NULL COMMENT '直径',
                        `outer_diameter` DECIMAL(18,6) NULL COMMENT '外径',
                        `inner_diameter` DECIMAL(18,6) NULL COMMENT '内径',
                        `nominal_diameter` VARCHAR(50) NULL COMMENT '公称直径DN',
                        `thickness` DECIMAL(18,6) NULL COMMENT '壁厚',
                        `weight` DECIMAL(18,6) NULL COMMENT '重量',
                        `density` DECIMAL(18,6) NULL COMMENT '密度',
                        `volume` DECIMAL(18,6) NULL COMMENT '体积',
                        `flow` DECIMAL(18,6) NULL COMMENT '流量',
                        `velocity` DECIMAL(18,6) NULL COMMENT '流速',
                        `lift` DECIMAL(18,6) NULL COMMENT '扬程',
                        `power` DECIMAL(18,6) NULL COMMENT '功率',
                        `voltage` DECIMAL(18,6) NULL COMMENT '电压',
                        `current` DECIMAL(18,6) NULL COMMENT '电流',
                        `frequency` DECIMAL(18,6) NULL COMMENT '频率',
                        `conductivity` DECIMAL(18,6) NULL COMMENT '电导率',
                        `moisture` DECIMAL(18,6) NULL COMMENT '含湿量',
                        `humidity` DECIMAL(18,6) NULL COMMENT '湿度',
                        `vacuum` DECIMAL(18,6) NULL COMMENT '真空度',
                        `radiation` DECIMAL(18,6) NULL COMMENT '辐射量',
                        `pipe_spec` VARCHAR(100) NULL COMMENT '管道规格',
                        `pipe_nominal_diameter` VARCHAR(50) NULL COMMENT '管道公称直径',
                        `pipe_wall_thickness` DECIMAL(18,6) NULL COMMENT '管道壁厚',
                        `pipe_pressure_class` VARCHAR(50) NULL COMMENT '管道压力等级',
                        `connection_type` VARCHAR(50) NULL COMMENT '连接方式',
                        `pipe_slope` VARCHAR(50) NULL COMMENT '管道坡度',
                        `anticorrosion_treatment` VARCHAR(100) NULL COMMENT '防腐处理',
                        `valve_model` VARCHAR(100) NULL COMMENT '阀门型号',
                        `valve_body_material` VARCHAR(100) NULL COMMENT '阀体材质',
                        `valve_disc_material` VARCHAR(100) NULL COMMENT '阀板材质',
                        `valve_ball_material` VARCHAR(100) NULL COMMENT '球体材质',
                        `seal_material` VARCHAR(100) NULL COMMENT '密封材质',
                        `drive_type` VARCHAR(50) NULL COMMENT '传动方式',
                        `open_mode` VARCHAR(50) NULL COMMENT '开启方式',
                        `applicable_medium` VARCHAR(100) NULL COMMENT '适用介质',
                        `flange_model` VARCHAR(100) NULL COMMENT '法兰型号',
                        `flange_type` VARCHAR(50) NULL COMMENT '法兰类型',
                        `flange_face_type` VARCHAR(50) NULL COMMENT '密封面形式',
                        `flange_standard` VARCHAR(100) NULL COMMENT '法兰标准',
                        `bolt_spec` VARCHAR(100) NULL COMMENT '螺栓规格',
                        `reducer_spec` VARCHAR(100) NULL COMMENT '异径管规格',
                        `reducer_large_dn` VARCHAR(50) NULL COMMENT '大端DN',
                        `reducer_small_dn` VARCHAR(50) NULL COMMENT '小端DN',
                        `reducer_wall_thickness_large` DECIMAL(18,6) NULL COMMENT '大端壁厚',
                        `reducer_wall_thickness_small` DECIMAL(18,6) NULL COMMENT '小端壁厚',
                        `reducer_connection_type` VARCHAR(50) NULL COMMENT '异径管连接方式',
                        `reducer_conicity` VARCHAR(50) NULL COMMENT '锥度',
                        `reducer_eccentric_direction` VARCHAR(50) NULL COMMENT '偏心方向',
                        `reducer_applicable_medium` VARCHAR(100) NULL COMMENT '异径管适用介质',
                        `reducer_anticorrosion` VARCHAR(100) NULL COMMENT '异径管防腐',
                        `pump_model` VARCHAR(100) NULL COMMENT '泵型号',
                        `pump_flow` DECIMAL(18,6) NULL COMMENT '泵流量',
                        `pump_head` DECIMAL(18,6) NULL COMMENT '泵扬程',
                        `pump_body_material` VARCHAR(100) NULL COMMENT '泵体材质',
                        `motor_power` DECIMAL(18,6) NULL COMMENT '电机功率',
                        `inlet_outlet_diameter` VARCHAR(100) NULL COMMENT '进出口直径',
                        `rated_speed` DECIMAL(18,6) NULL COMMENT '额定转速',
                        `pump_applicable_medium` VARCHAR(100) NULL COMMENT '泵适用介质',
                        `working_pressure` DECIMAL(18,6) NULL COMMENT '工作压力',
                        `protection_level` VARCHAR(50) NULL COMMENT '防护等级',
                        `expansion_joint_model` VARCHAR(100) NULL COMMENT '膨胀节型号',
                        `bellows_material` VARCHAR(100) NULL COMMENT '波纹管材质',
                        `flange_or_nozzle_material` VARCHAR(100) NULL COMMENT '法兰/接管材质',
                        `compensation_amount` DECIMAL(18,6) NULL COMMENT '补偿量',
                        `expansion_joint_connection_type` VARCHAR(50) NULL COMMENT '膨胀节连接方式',
                        `expansion_joint_medium` VARCHAR(100) NULL COMMENT '膨胀节适用介质',
                        `expansion_joint_working_temp` DECIMAL(18,6) NULL COMMENT '膨胀节工作温度',
                        `flue_gas_capacity` DECIMAL(18,6) NULL COMMENT '处理烟气量',
                        `desulfurization_efficiency` DECIMAL(18,6) NULL COMMENT '脱硫效率',
                        `droplet_size` DECIMAL(18,6) NULL COMMENT '液滴粒径',
                        `spray_layer_count` INT NULL COMMENT '喷淋层数',
                        `chimney_spec` VARCHAR(100) NULL COMMENT '烟囱规格',
                        `chimney_diameter` DECIMAL(18,6) NULL COMMENT '烟囱直径',
                        `chimney_height` DECIMAL(18,6) NULL COMMENT '烟囱高度',
                        `chimney_material` VARCHAR(100) NULL COMMENT '烟囱材质',
                        `chimney_thickness` DECIMAL(18,6) NULL COMMENT '烟囱壁厚',
                        `outlet_wind_speed` DECIMAL(18,6) NULL COMMENT '出口风速',
                        `insulation_thickness` DECIMAL(18,6) NULL COMMENT '保温层厚度',
                        `support_type` VARCHAR(100) NULL COMMENT '支撑方式',
                        `flue_gas_temperature` DECIMAL(18,6) NULL COMMENT '烟气温度',
                        `pressure_gauge_model` VARCHAR(100) NULL COMMENT '压力表型号',
                        `thermometer_model` VARCHAR(100) NULL COMMENT '温度计型号',
                        `filter_model` VARCHAR(100) NULL COMMENT '过滤器型号',
                        `check_valve_model` VARCHAR(100) NULL COMMENT '止回阀型号',
                        `sprinkler_model` VARCHAR(100) NULL COMMENT '喷淋头型号',
                        `flow_meter_model` VARCHAR(100) NULL COMMENT '流量计型号',
                        `safety_valve_model` VARCHAR(100) NULL COMMENT '安全阀型号',
                        `flexible_joint_model` VARCHAR(100) NULL COMMENT '柔性接头型号',
                        `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
                        `updated_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
                        PRIMARY KEY (`id`),
                        UNIQUE KEY `uk_file_attribute_id` (`file_attribute_id`),
                        KEY `idx_category_id` (`category_id`),
                        KEY `idx_file_storage_id` (`file_storage_id`),
                        KEY `idx_file_name` (`file_name`)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // system_config
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `system_config` (
                        `config_key` VARCHAR(200) PRIMARY KEY,
                        `config_value` TEXT
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;");

                    // departments
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `departments` (
                        `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                        `name` VARCHAR(200) NOT NULL UNIQUE,
                        `display_name` VARCHAR(200),
                        `sort_order` INT DEFAULT 0,
                        `is_active` TINYINT DEFAULT 1,
                        `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
                        `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;");

                    // users (详细人员信息表)
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `users` (
                        `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                        `username` VARCHAR(100) NOT NULL UNIQUE,
                        `password_hash` VARCHAR(512),
                        `display_name` VARCHAR(200),
                        `gender` ENUM('男','女','无信息') DEFAULT '无信息',
                        `phone` VARCHAR(32),
                        `email` VARCHAR(200),
                        `department_id` INT,
                        `role` VARCHAR(64),
                        `status` TINYINT DEFAULT 1,
                        `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
                        `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                        INDEX(`department_id`),
                        CONSTRAINT `fk_users_department` FOREIGN KEY (`department_id`) REFERENCES `departments`(`id`) ON DELETE SET NULL ON UPDATE CASCADE
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;");

                    // department_users (可选的多对多)
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `department_users` (
                        `department_id` INT NOT NULL,
                        `user_id` INT NOT NULL,
                        PRIMARY KEY (`department_id`,`user_id`),
                        INDEX(`user_id`),
                        CONSTRAINT `fk_dept_users_dept` FOREIGN KEY (`department_id`) REFERENCES `departments`(`id`) ON DELETE CASCADE ON UPDATE CASCADE,
                        CONSTRAINT `fk_dept_users_user` FOREIGN KEY (`user_id`) REFERENCES `users`(`id`) ON DELETE CASCADE ON UPDATE CASCADE
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;");

                    // category_department_map (一对一映射 cad_categories.id -> departments.id)
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `category_department_map` (
                        `category_id` INT NOT NULL PRIMARY KEY,
                        `department_id` INT NOT NULL UNIQUE,
                        CONSTRAINT `fk_map_category` FOREIGN KEY (`category_id`) REFERENCES `cad_categories`(`id`) ON DELETE CASCADE ON UPDATE CASCADE,
                        CONSTRAINT `fk_map_department` FOREIGN KEY (`department_id`) REFERENCES `departments`(`id`) ON DELETE CASCADE ON UPDATE CASCADE
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;");

                    // file_access_logs & file_tags & file_version_history (ensure basic tables)
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `file_access_logs` (
                        `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                        `file_id` INT,
                        `user_name` VARCHAR(200),
                        `action_type` VARCHAR(50),
                        `ip_address` VARCHAR(64),
                        `user_agent` VARCHAR(512),
                        `access_time` DATETIME DEFAULT CURRENT_TIMESTAMP
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;");

                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `file_tags` (
                        `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                        `file_id` INT,
                        `tag_name` VARCHAR(200),
                        `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;");

                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `file_version_history` (
                        `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                        `file_id` INT,
                        `version` INT,
                        `file_name` VARCHAR(512),
                        `stored_file_name` VARCHAR(512),
                        `file_path` VARCHAR(1024),
                        `file_size` BIGINT,
                        `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
                        `updated_by` VARCHAR(200),
                        `change_description` TEXT
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;");

                    // sw_categories：SW主分类表
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `sw_categories` (
                        `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                        `name` VARCHAR(200) NOT NULL,
                        `display_name` VARCHAR(200),
                        `sort_order` INT DEFAULT 0,
                        `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
                        `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // sw_subcategories：SW子分类表
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `sw_subcategories` (
                        `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                        `category_id` INT NOT NULL,
                        `parent_id` INT NOT NULL DEFAULT 0,
                        `name` VARCHAR(200) NOT NULL,
                        `display_name` VARCHAR(200),
                        `sort_order` INT DEFAULT 0,
                        `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
                        `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                        KEY `idx_sw_sub_cat` (`category_id`),
                        KEY `idx_sw_sub_parent` (`parent_id`)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // sw_graphics：SW图元表
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `sw_graphics` (
                        `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                        `subcategory_id` INT NOT NULL,
                        `file_name` VARCHAR(512) NOT NULL,
                        `display_name` VARCHAR(512) NULL,
                        `file_path` VARCHAR(1024) NULL,
                        `preview_image_path` VARCHAR(1024) NULL,
                        `file_size` BIGINT NULL,
                        `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
                        `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                        KEY `idx_sw_graphics_sub` (`subcategory_id`)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;");

                    // device_info：设备信息表（与 GetAllDeviceInfoAsync / InsertDeviceInfoBatchAsync 对齐）
                    sql.AppendLine(@"CREATE TABLE IF NOT EXISTS `device_info` (
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

                    conn.Execute(sql.ToString());
                }

                return true;
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

        // 新增：部门实体
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
        /// 兼容旧调用：按图元主键读取属性记录。
        /// </summary>
        public async Task<FileAttribute?> GetFileAttributeByGraphicIdAsync(int fileStorageId)
        {
            try
            {
                var result = await GetFileWithAttributeAsync(fileStorageId).ConfigureAwait(false);
                return result.Attribute;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetFileAttributeByGraphicIdAsync 出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 兼容旧调用：级联删除图元及其属性，并可选删除物理文件。
        /// </summary>
        public async Task<bool> DeleteCadGraphicCascadeAsync(int fileId, bool physicalDelete = true)
        {
            try
            {
                using var connection = GetConnection();
                await connection.OpenAsync().ConfigureAwait(false);
                using var transaction = connection.BeginTransaction();

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

                await connection.ExecuteAsync("DELETE FROM file_tags WHERE file_id = @Id", new { Id = fileId }, transaction).ConfigureAwait(false);
                await connection.ExecuteAsync("DELETE FROM file_access_logs WHERE file_id = @Id", new { Id = fileId }, transaction).ConfigureAwait(false);
                await connection.ExecuteAsync("DELETE FROM file_version_history WHERE file_id = @Id", new { Id = fileId }, transaction).ConfigureAwait(false);
                await connection.ExecuteAsync(
                    "DELETE FROM cad_file_attributes WHERE file_storage_id = @Id OR file_attribute_id = @FileAttributeId",
                    new { Id = fileId, FileAttributeId = storage.FileAttributeId },
                    transaction).ConfigureAwait(false);

                int affected = await connection.ExecuteAsync(
                    "DELETE FROM cad_file_storage WHERE id = @Id",
                    new { Id = fileId },
                    transaction).ConfigureAwait(false);

                transaction.Commit();

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
        /// 获取文件的详细信息（包括属性）
        /// </summary>
        public async Task<(FileStorage File, FileAttribute Attribute)> GetFileWithAttributeAsync(int fileId)
        {
            // 查询文件存储信息
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
        WHERE id = @Id;";

            // 按业务属性ID精确查询
            var attrByBizIdSql = $@"
        SELECT {CadFileAttributeSelectColumns}
        FROM cad_file_attributes
        WHERE file_attribute_id = @FileAttributeId
        LIMIT 1;";

            // 老数据兼容：若 storage.file_attribute_id 里存的是数字主键，尝试按属性主键查
            var attrByLegacyIdSql = $@"
        SELECT {CadFileAttributeSelectColumns}
        FROM cad_file_attributes
        WHERE id = @LegacyAttrId
        LIMIT 1;";

            // 最后兜底：按 file_storage_id 查询
            var attrByStorageIdSql = $@"
        SELECT {CadFileAttributeSelectColumns}
        FROM cad_file_attributes
        WHERE file_storage_id = @FileStorageId
        LIMIT 1;";

            try
            {
                using var connection = new MySqlConnection(_connectionString);

                // 先取文件
                var file = await connection.QuerySingleOrDefaultAsync<FileStorage>(fileSql, new { Id = fileId }).ConfigureAwait(false);
                if (file == null)
                {
                    return (null, null);
                }

                FileAttribute attribute = null;

                // 兼容 FileStorage.FileAttributeId 目前可能是 int 或 string，这里统一转字符串处理
                var storageAttrIdText = Convert.ToString(file.FileAttributeId)?.Trim();

                // 第一优先：按业务ID精确匹配
                if (!string.IsNullOrWhiteSpace(storageAttrIdText))
                {
                    attribute = await connection.QuerySingleOrDefaultAsync<FileAttribute>(
                        attrByBizIdSql,
                        new { FileAttributeId = storageAttrIdText }
                    ).ConfigureAwait(false);
                }

                // 第二优先：老数据兼容（数字主键）
                if (attribute == null && int.TryParse(storageAttrIdText, out var legacyAttrId))
                {
                    attribute = await connection.QuerySingleOrDefaultAsync<FileAttribute>(
                        attrByLegacyIdSql,
                        new { LegacyAttrId = legacyAttrId }
                    ).ConfigureAwait(false);
                }

                // 最后兜底：按 file_storage_id 关联
                if (attribute == null)
                {
                    attribute = await connection.QuerySingleOrDefaultAsync<FileAttribute>(
                        attrByStorageIdSql,
                        new { FileStorageId = file.Id }
                    ).ConfigureAwait(false);
                }

                return (file, attribute);
            }
            catch (Exception ex)
            {
                Env.Editor.WriteMessage($"\n获取文件详细信息时出错: {ex.Message}");
                return (null, null);
            }
        }

        /// <summary>
        /// 获取文件的详细信息（包括属性）
        /// </summary>
        public async Task<FileStorage> GetFileStorageAsync(string filehash)
        {
            if (string.IsNullOrWhiteSpace(filehash))
                return null;

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
             LIMIT 1";

            try
            {
                using var connection = new MySqlConnection(_connectionString);
                var fileStorageInfo = await connection.QueryFirstOrDefaultAsync<FileStorage>(
                    fileSql, new { FileHash = filehash });
                return fileStorageInfo;
            }
            catch (Exception ex)
            {

                LogManager.Instance.LogInfo($"获取文件详细信息时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 获取文件属性的详细信息
        /// </summary>
        public async Task<FileAttribute> GetFileAttributeAsync(string fileName)
        {
            // 按文件名模糊查询属性（兼容带扩展名/不带扩展名）
            string noExtName = Path.GetFileNameWithoutExtension(fileName) ?? string.Empty;
            string rawName = Path.GetFileName(fileName) ?? string.Empty;

            var sql = $@"
        SELECT {CadFileAttributeSelectColumns}
        FROM cad_file_attributes
        WHERE file_name LIKE CONCAT('%', @Name1, '%')
           OR file_name LIKE CONCAT('%', @Name2, '%')
        LIMIT 1;";

            try
            {
                using var connection = new MySqlConnection(_connectionString);
                return await connection.QueryFirstOrDefaultAsync<FileAttribute>(sql, new { Name1 = noExtName, Name2 = rawName }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"获取文件属性失败: {ex.Message}");
                return null;
            }
        }

        // ===== 新增：cad_file_attributes 全字段查询片段（供多个查询方法复用）=====
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

        // ===== 新增：统一构造参数（插入/更新复用）=====
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
        /// 事务插入：先插入文件存储，再插入属性，并回写 file_attribute_id（业务ID）
        /// </summary>
        public async Task<(int StorageId, int AttributeId)> AddFileStorageAndAttributeAsync(FileStorage storage, FileAttribute attribute)
        {
            // 参数保护：防止空对象
            if (storage == null) throw new ArgumentNullException(nameof(storage));
            if (attribute == null) throw new ArgumentNullException(nameof(attribute));

            // 业务属性ID兜底：若为空则自动生成
            if (string.IsNullOrWhiteSpace(attribute.FileAttributeId))
            {
                attribute.FileAttributeId = Guid.NewGuid().ToString("N");
            }

            using var connection = GetConnection();
            await connection.OpenAsync().ConfigureAwait(false);

            MySqlTransaction tx = null;
            int storageId = 0;

            try
            {
                // 开启事务，保证“存储+属性+回写关联”原子性
                tx = connection.BeginTransaction();

                // 1) 先插入 cad_file_storage（先把 file_attribute_id 置空，后面再回写业务ID）
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

                // 取 storage 自增主键
                storageId = await connection.QuerySingleAsync<int>("SELECT LAST_INSERT_ID()", transaction: tx).ConfigureAwait(false);

                // 2) 插入属性表：先挂上 storageId/categoryId
                attribute.FileStorageId = storageId;
                attribute.CategoryId = storage.CategoryId;

                const string insertAttributeSql = @"
        INSERT INTO cad_file_attributes
        (
            category_id, file_storage_id, file_name, file_attribute_id, description, attribute_group, remarks,
            customize1, customize2, customize3, length, width, height, angle, base_point_x, base_point_y, base_point_z,
            model, specifications, material, medium_name, standard_number, pressure, temperature, pressure_rating,
            operating_pressure, operating_temperature, diameter, outer_diameter, inner_diameter, nominal_diameter,
            thickness, weight, density, volume, flow, velocity, lift, power, voltage, current, frequency, conductivity,
            moisture, humidity, vacuum, radiation,
            pipe_spec, pipe_nominal_diameter, pipe_wall_thickness, pipe_pressure_class, connection_type, pipe_slope, anticorrosion_treatment,
            valve_model, valve_body_material, valve_disc_material, valve_ball_material, seal_material, drive_type, open_mode, applicable_medium,
            flange_model, flange_type, flange_face_type, flange_standard, bolt_spec,
            reducer_spec, reducer_large_dn, reducer_small_dn, reducer_wall_thickness_large, reducer_wall_thickness_small,
            reducer_connection_type, reducer_conicity, reducer_eccentric_direction, reducer_applicable_medium, reducer_anticorrosion,
            pump_model, pump_flow, pump_head, pump_body_material, motor_power, inlet_outlet_diameter, rated_speed, pump_applicable_medium,
            working_pressure, protection_level,
            expansion_joint_model, bellows_material, flange_or_nozzle_material, compensation_amount, expansion_joint_connection_type,
            expansion_joint_medium, expansion_joint_working_temp,
            flue_gas_capacity, desulfurization_efficiency, droplet_size, spray_layer_count,
            chimney_spec, chimney_diameter, chimney_height, chimney_material, chimney_thickness, outlet_wind_speed, insulation_thickness,
            support_type, flue_gas_temperature,
            pressure_gauge_model, thermometer_model, filter_model, check_valve_model, sprinkler_model, flow_meter_model, safety_valve_model, flexible_joint_model,
            created_at, updated_at
        )
        VALUES
        (
            @CategoryId, @FileStorageId, @FileName, @FileAttributeId, @Description, @AttributeGroup, @Remarks,
            @Customize1, @Customize2, @Customize3, @Length, @Width, @Height, @Angle, @BasePointX, @BasePointY, @BasePointZ,
            @Model, @Specifications, @Material, @MediumName, @StandardNumber, @Pressure, @Temperature, @PressureRating,
            @OperatingPressure, @OperatingTemperature, @Diameter, @OuterDiameter, @InnerDiameter, @NominalDiameter,
            @Thickness, @Weight, @Density, @Volume, @Flow, @Velocity, @Lift, @Power, @Voltage, @Current, @Frequency, @Conductivity,
            @Moisture, @Humidity, @Vacuum, @Radiation,
            @PipeSpec, @PipeNominalDiameter, @PipeWallThickness, @PipePressureClass, @ConnectionType, @PipeSlope, @AnticorrosionTreatment,
            @ValveModel, @ValveBodyMaterial, @ValveDiscMaterial, @ValveBallMaterial, @SealMaterial, @DriveType, @OpenMode, @ApplicableMedium,
            @FlangeModel, @FlangeType, @FlangeFaceType, @FlangeStandard, @BoltSpec,
            @ReducerSpec, @ReducerLargeDn, @ReducerSmallDn, @ReducerWallThicknessLarge, @ReducerWallThicknessSmall,
            @ReducerConnectionType, @ReducerConicity, @ReducerEccentricDirection, @ReducerApplicableMedium, @ReducerAnticorrosion,
            @PumpModel, @PumpFlow, @PumpHead, @PumpBodyMaterial, @MotorPower, @InletOutletDiameter, @RatedSpeed, @PumpApplicableMedium,
            @WorkingPressure, @ProtectionLevel,
            @ExpansionJointModel, @BellowsMaterial, @FlangeOrNozzleMaterial, @CompensationAmount, @ExpansionJointConnectionType,
            @ExpansionJointMedium, @ExpansionJointWorkingTemp,
            @FlueGasCapacity, @DesulfurizationEfficiency, @DropletSize, @SprayLayerCount,
            @ChimneySpec, @ChimneyDiameter, @ChimneyHeight, @ChimneyMaterial, @ChimneyThickness, @OutletWindSpeed, @InsulationThickness,
            @SupportType, @FlueGasTemperature,
            @PressureGaugeModel, @ThermometerModel, @FilterModel, @CheckValveModel, @SprinklerModel, @FlowMeterModel, @SafetyValveModel, @FlexibleJointModel,
            @CreatedAt, @UpdatedAt
        );";

                var attrParams = BuildFileAttributeParameters(attribute, includeId: false);
                await connection.ExecuteAsync(insertAttributeSql, attrParams, tx).ConfigureAwait(false);

                // 取属性主键（仅返回用）
                int attrId = await connection.QuerySingleAsync<int>("SELECT LAST_INSERT_ID()", transaction: tx).ConfigureAwait(false);

                // 3) 回写 storage.file_attribute_id = 业务ID字符串（关键修复点）
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

                // 事务提交
                tx.Commit();
                return (storageId, attrId);
            }
            catch (Exception ex)
            {
                // 异常回滚
                try { tx?.Rollback(); } catch { }
                LogManager.Instance.LogInfo($"AddFileStorageAndAttributeAsync 失败: {ex.Message}");
                return (0, 0);
            }
            finally
            {
                try { tx?.Dispose(); } catch { }
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
        /// 设备信息
        /// </summary>
        public class DeviceInfo
        {
            public int Id { get; set; } // 设备ID
            public string Name { get; set; } = string.Empty; // 设备名
            public string Type { get; set; } = string.Empty; // 设备类型
            public string MediumName { get; set; } = string.Empty; // 介质
            public string Specifications { get; set; } = string.Empty; // 规格
            public string Material { get; set; } = string.Empty; // 材质
            public int Quantity { get; set; } // 数量
            public string DrawingNumber { get; set; } = string.Empty; // 图号
            public decimal Power { get; set; } // 功率
            public decimal Volume { get; set; } // 容积
            public decimal Pressure { get; set; } // 压力
            public decimal Temperature { get; set; } // 温度
            public decimal Diameter { get; set; } // 直径
            public decimal Length { get; set; } // 长度
            public decimal Thickness { get; set; } // 厚度
            public decimal Weight { get; set; } // 重量
            public string Model { get; set; } = string.Empty; // 型号
            public string Remarks { get; set; } = string.Empty; // 备注
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


    }
}
