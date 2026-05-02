using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// MySQL 认证服务（增强：部门管理与用户-部门分配）
    /// </summary>
    public class MySqlAuthService
    {
        private readonly string _server; // 数据库服务器地址
        private readonly string _port; //  数据库服务器端口
        private readonly string _database = "cad_sw_library"; // 数据库名称
        private readonly string _dbUser = "root"; // 默认用户名（请在生产环境改为配置化、低权限账号）
        private readonly string _dbPassword = "root";

        public MySqlAuthService(string server, string port)
        {
            _server = string.IsNullOrEmpty(server) ? "127.0.0.1" : server;
            _port = string.IsNullOrEmpty(port) ? "3306" : port;
        }

        private string ConnString()
        {
            return $"Server={_server};Port={_port};Database={_database};Uid={_dbUser};Pwd={_dbPassword};Allow User Variables=True;";
        }

        #region 表结构保证与同步
        /// <summary>
        /// 确保 cad_categories 表存在
        /// </summary>
        public void EnsureCategoriesTableExists()
        {
            using (var conn = new MySqlConnection(ConnString()))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                cmd.CommandText = @"
                CREATE TABLE IF NOT EXISTS cad_categories (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    parent_id INT DEFAULT 0,
                    name VARCHAR(200) NOT NULL,
                    display_name VARCHAR(200),
                    sort_order INT DEFAULT 0,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";
                cmd.ExecuteNonQuery();
            }
        }
        /// <summary>
        /// 确保 departments 表存在
        /// </summary>
        public void EnsureDepartmentsTableExists()
        {
            using var conn = new MySqlConnection(ConnString());
            conn.Open();

            // 创建 departments 表（若不存在）
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
                    CREATE TABLE IF NOT EXISTS departments (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        cad_category_id INT DEFAULT NULL,
                        name VARCHAR(200) NOT NULL,
                        display_name VARCHAR(200),
                        description TEXT,
                        manager_user_id INT NULL,
                        sort_order INT DEFAULT 0,
                        is_active TINYINT(1) DEFAULT 1,
                        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                        updated_at DATETIME NULL
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";
                cmd.ExecuteNonQuery();
            }

            // 确保 departments 表包含必要列与索引（幂等）
            EnsureDepartmentsTableSchema(conn);

            // 确保 department_users 表存在（多对多映射）
            try
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        CREATE TABLE IF NOT EXISTS department_users (
                            department_id INT NOT NULL,
                            user_id INT NOT NULL,
                            PRIMARY KEY (department_id, user_id),
                            INDEX (user_id)
                        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"EnsureDepartmentsTableExists: 创建 department_users 失败: {ex.Message}");
            }
        }
        /// <summary>
        /// 确保 departments 表包含 cad_category_id 列与可用索引（如果可能则创建唯一索引 ux_cad_category_id）
        /// </summary>
        private void EnsureDepartmentsTableSchema(MySqlConnection conn)
        {
            // 保证常用列存在（包含 description/display_name 等）
            var requiredColumns = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "cad_category_id", "INT DEFAULT NULL" },
                { "name", "VARCHAR(200) NOT NULL" },
                { "display_name", "VARCHAR(200) NULL" },
                { "description", "TEXT NULL" },
                { "manager_user_id", "INT NULL" },
                { "sort_order", "INT DEFAULT 0" },
                { "is_active", "TINYINT(1) DEFAULT 1" },
                { "created_at", "DATETIME DEFAULT CURRENT_TIMESTAMP" },
                { "updated_at", "DATETIME NULL" }
            };

            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var colCmd = conn.CreateCommand())
            {
                colCmd.CommandText = "SELECT column_name FROM information_schema.columns WHERE table_schema=@schema AND table_name='departments'";
                colCmd.Parameters.AddWithValue("@schema", conn.Database);
                using (var reader = colCmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        existing.Add(reader.GetString(0));
                    }
                }
            }

            foreach (var kv in requiredColumns)
            {
                if (!existing.Contains(kv.Key))
                {
                    try
                    {
                        using (var alter = conn.CreateCommand())
                        {
                            alter.CommandText = $"ALTER TABLE `departments` ADD COLUMN `{kv.Key}` {kv.Value};";
                            alter.ExecuteNonQuery();
                        }
                        LogManager.Instance.LogInfo($"已向 departments 表添加列 {kv.Key}");
                        existing.Add(kv.Key);
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogInfo($"向 departments 表添加列 {kv.Key} 失败: {ex.Message}");
                    }
                }
            }

            // 索引处理（尽量创建唯一索引，否则普通索引）
            try
            {
                using (var idxCheck = conn.CreateCommand())
                {
                    idxCheck.CommandText = @"
                        SELECT COUNT(1)
                        FROM information_schema.statistics
                        WHERE table_schema = @schema AND table_name = 'departments' AND index_name = 'ux_cad_category_id';";
                    idxCheck.Parameters.AddWithValue("@schema", conn.Database);
                    var idxExists = Convert.ToInt32(idxCheck.ExecuteScalar() ?? 0) > 0;
                    if (!idxExists)
                    {
                        using (var dupCheck = conn.CreateCommand())
                        {
                            dupCheck.CommandText = @"
                                SELECT COUNT(*) FROM (
                                  SELECT cad_category_id FROM departments
                                  WHERE cad_category_id IS NOT NULL
                                  GROUP BY cad_category_id HAVING COUNT(*) > 1
                                ) t;";
                            var dupCount = Convert.ToInt32(dupCheck.ExecuteScalar() ?? 0);
                            if (dupCount == 0)
                            {
                                using (var createIdx = conn.CreateCommand())
                                {
                                    createIdx.CommandText = "ALTER TABLE `departments` ADD UNIQUE INDEX `ux_cad_category_id` (`cad_category_id`);";
                                    createIdx.ExecuteNonQuery();
                                }
                                LogManager.Instance.LogInfo("已为 departments.cad_category_id 创建唯一索引 ux_cad_category_id");
                            }
                            else
                            {
                                using (var createIdx = conn.CreateCommand())
                                {
                                    createIdx.CommandText = "ALTER TABLE `departments` ADD INDEX `idx_cad_category_id` (`cad_category_id`);";
                                    createIdx.ExecuteNonQuery();
                                }
                                LogManager.Instance.LogInfo("departments 表存在 cad_category_id 重复，创建非唯一索引 idx_cad_category_id");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"EnsureDepartmentsTableSchema 索引处理出错: {ex.Message}");
            }
        }
        /// <summary>
        /// 将 cad_categories 的顶级分类同步到 departments 表（插入/更新）
        /// </summary>
        public void SyncDepartmentsFromCadCategories()
        {
            using var conn = new MySqlConnection(ConnString());
            conn.Open();
            using var tx = conn.BeginTransaction();
            try
            {
                var sel = conn.CreateCommand();
                sel.Transaction = tx;
                sel.CommandText = "SELECT id, name, display_name, sort_order FROM cad_categories ORDER BY sort_order, id;";
                var categories = new List<(int Id, string Name, string DisplayName, int SortOrder)>();
                using (var reader = sel.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        categories.Add((
                            reader.GetInt32("id"),
                            reader.IsDBNull(reader.GetOrdinal("name")) ? "" : reader.GetString("name"),
                            reader.IsDBNull(reader.GetOrdinal("display_name")) ? "" : reader.GetString("display_name"),
                            reader.IsDBNull(reader.GetOrdinal("sort_order")) ? 0 : reader.GetInt32("sort_order")
                        ));
                    }
                }

                if (categories.Count == 0)
                {
                    tx.Commit();
                    return;
                }

                bool hasCadColumn;
                using (var colCheck = conn.CreateCommand())
                {
                    colCheck.Transaction = tx;
                    colCheck.CommandText = "SELECT COUNT(1) FROM information_schema.columns WHERE table_schema = @schema AND table_name = 'departments' AND column_name = 'cad_category_id';";
                    colCheck.Parameters.AddWithValue("@schema", conn.Database);
                    hasCadColumn = Convert.ToInt32(colCheck.ExecuteScalar() ?? 0) > 0;
                }

                if (hasCadColumn)
                {
                    using (var upsert = conn.CreateCommand())
                    {
                        upsert.Transaction = tx;
                        upsert.CommandText = @"
                            INSERT INTO departments (cad_category_id, name, display_name, sort_order, is_active, created_at)
                            VALUES (@cad, @name, @display, @so, 1, NOW())
                            ON DUPLICATE KEY UPDATE
                                name = VALUES(name),
                                display_name = VALUES(display_name),
                                sort_order = VALUES(sort_order),
                                is_active = 1,
                                updated_at = CURRENT_TIMESTAMP;";
                        upsert.Parameters.Add("@cad", MySqlDbType.Int32);
                        upsert.Parameters.Add("@name", MySqlDbType.VarChar);
                        upsert.Parameters.Add("@display", MySqlDbType.VarChar);
                        upsert.Parameters.Add("@so", MySqlDbType.Int32);

                        foreach (var c in categories)
                        {
                            upsert.Parameters["@cad"].Value = c.Id;
                            upsert.Parameters["@name"].Value = c.Name ?? "";
                            upsert.Parameters["@display"].Value = string.IsNullOrEmpty(c.DisplayName) ? (c.Name ?? "") : c.DisplayName;
                            upsert.Parameters["@so"].Value = c.SortOrder;
                            upsert.ExecuteNonQuery();
                        }
                    }
                }
                else
                {
                    foreach (var c in categories)
                    {
                        using (var find = conn.CreateCommand())
                        {
                            find.Transaction = tx;
                            find.CommandText = "SELECT id FROM departments WHERE name = @name LIMIT 1;";
                            find.Parameters.AddWithValue("@name", c.Name ?? "");
                            var existingId = Convert.ToInt32(find.ExecuteScalar() ?? 0);
                            if (existingId > 0)
                            {
                                using (var upd = conn.CreateCommand())
                                {
                                    upd.Transaction = tx;
                                    upd.CommandText = "UPDATE departments SET display_name=@display, sort_order=@so, updated_at=CURRENT_TIMESTAMP WHERE id=@id;";
                                    upd.Parameters.AddWithValue("@display", string.IsNullOrEmpty(c.DisplayName) ? (c.Name ?? "") : c.DisplayName);
                                    upd.Parameters.AddWithValue("@so", c.SortOrder);
                                    upd.Parameters.AddWithValue("@id", existingId);
                                    upd.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                using (var ins = conn.CreateCommand())
                                {
                                    ins.Transaction = tx;
                                    ins.CommandText = "INSERT INTO departments (name, display_name, sort_order, is_active, created_at) VALUES (@name, @display, @so, 1, NOW());";
                                    ins.Parameters.AddWithValue("@name", c.Name ?? "");
                                    ins.Parameters.AddWithValue("@display", string.IsNullOrEmpty(c.DisplayName) ? (c.Name ?? "") : c.DisplayName);
                                    ins.Parameters.AddWithValue("@so", c.SortOrder);
                                    ins.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                }

                tx.Commit();
            }
            catch (Exception ex)
            {
                try { tx.Rollback(); } catch { }
                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories 失败: {ex.Message}");
                throw;
            }
        }


        #endregion

        #region 用户表与认证（增强）
        /// <summary>
        /// 确保 users 表存在且包含注册/认证/分配所需列。
        /// - 如果不存在 users 表，则创建。
        /// - 如果存在旧表名 gb_user，会在可能时把数据迁移到 users（尽量保留密码/盐）。
        /// - 如果缺少列，会按需 ALTER TABLE ADD COLUMN。
        /// 这样可以避免出现 Unknown column 'salt' 的错误。
        /// </summary>
        public void EnsureUserTableExists()
        {
            try
            {
                using var conn = new MySqlConnection(ConnString());
                conn.Open();

                var create = conn.CreateCommand();
                create.CommandText = @"
                 CREATE TABLE IF NOT EXISTS `users` (
                     `id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                     `username` VARCHAR(100) NOT NULL UNIQUE,
                     `password_hash` VARCHAR(512) NOT NULL,
                     `salt` VARCHAR(64) NOT NULL,
                     `display_name` VARCHAR(200),
                     `gender` ENUM('男','女','无信息') DEFAULT '无信息',
                     `phone` VARCHAR(50),
                     `email` VARCHAR(200),
                     `role` VARCHAR(64),
                     `department_id` INT NOT NULL DEFAULT 0,
                     `department_name` VARCHAR(200) DEFAULT '',
                     `is_active` TINYINT(1) DEFAULT 1,
                     `last_login` DATETIME NULL,
                     `avatar_url` VARCHAR(512) DEFAULT NULL,
                     `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
                     `created_by` VARCHAR(100),
                     `updated_at` DATETIME NULL,
                     INDEX idx_department_id (department_id)
                 ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;";
                create.ExecuteNonQuery();

                EnsureUserTableSchema(conn);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"EnsureUserTableExists 失败: {ex.Message}");
                throw;
            }
        }
        /// <summary>
        /// 检查用户名是否存在（只检查 users 表）
        /// </summary>
        public bool UserExists(string username)
        {
            if (string.IsNullOrWhiteSpace(username)) return false;
            try
            {
                using var conn = new MySqlConnection(ConnString());
                conn.Open();
                var q = conn.CreateCommand();
                q.CommandText = "SELECT COUNT(1) FROM users WHERE username = @u";
                q.Parameters.AddWithValue("@u", username);
                var cnt = Convert.ToInt32(q.ExecuteScalar() ?? 0);
                return cnt > 0;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"UserExists 出错: {ex.Message}");
                // 为了避免潜在重复创建风险，保守返回 true；若要用于调试可以改为 false
                return true;
            }
        }
        /// <summary>
        /// 检查表是否存在
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private static bool TableExists(MySqlConnection conn, string tableName)
        {
            var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = @schema AND table_name = @table";
            cmd.Parameters.AddWithValue("@schema", conn.Database);
            cmd.Parameters.AddWithValue("@table", tableName);
            var exists = Convert.ToInt32(cmd.ExecuteScalar() ?? 0) > 0;
            return exists;
        }
        /// <summary>
        /// 确保 users 表包含所需列，若缺失则添加（幂等）。
        /// </summary>
        private void EnsureUserTableSchema(MySqlConnection conn)
        {
            var requiredColumns = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "password_hash", "VARCHAR(512) NOT NULL" },
                { "salt", "VARCHAR(64) NOT NULL" },
                { "display_name", "VARCHAR(200) NULL" },
                { "gender", "ENUM('男','女','无信息') DEFAULT '无信息'" },
                { "phone", "VARCHAR(50) NULL" },
                { "email", "VARCHAR(200) NULL" },
                { "role", "VARCHAR(64) NULL" },
                { "department_id", "INT NOT NULL DEFAULT 0" },
                { "department_name", "VARCHAR(200) DEFAULT ''" },
                { "is_active", "TINYINT(1) DEFAULT 1" },
                { "last_login", "DATETIME NULL" },
                { "avatar_url", "VARCHAR(512) NULL" },
                { "created_at", "DATETIME DEFAULT CURRENT_TIMESTAMP" },
                { "created_by", "VARCHAR(100) NULL" },
                { "updated_at", "DATETIME NULL" }
            };

            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 1) 安全地读取已存在列（确保 reader/command 都被释放）
            using (var colCmd = conn.CreateCommand())
            {
                colCmd.CommandText = "SELECT column_name FROM information_schema.columns WHERE table_schema=@schema AND table_name='users'";
                colCmd.Parameters.AddWithValue("@schema", conn.Database);
                using (var reader = colCmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        existing.Add(reader.GetString(0));
                    }
                }
            }

            // 2) 添加缺失列（幂等）
            foreach (var kv in requiredColumns)
            {
                if (!existing.Contains(kv.Key))
                {
                    try
                    {
                        using (var alter = conn.CreateCommand())
                        {
                            alter.CommandText = $"ALTER TABLE `users` ADD COLUMN `{kv.Key}` {kv.Value};";
                            alter.ExecuteNonQuery();
                        }
                        LogManager.Instance.LogInfo($"已向 users 表添加列 {kv.Key}");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogInfo($"向 users 表添加列 {kv.Key} 失败: {ex.Message}");
                    }
                }
            }

            // 3) 确保索引存在：使用独立的 command 并为 table/index 使用参数
            try
            {
                using (var idxCmd = conn.CreateCommand())
                {
                    idxCmd.CommandText = @"
                SELECT COUNT(1) FROM information_schema.statistics
                WHERE table_schema=@schema AND table_name=@table AND index_name=@index;";
                    idxCmd.Parameters.AddWithValue("@schema", conn.Database);
                    idxCmd.Parameters.AddWithValue("@table", "users");
                    idxCmd.Parameters.AddWithValue("@index", "idx_department_id");

                    var hasIdx = Convert.ToInt32(idxCmd.ExecuteScalar() ?? 0) > 0;
                    if (!hasIdx)
                    {
                        using (var createIdx = conn.CreateCommand())
                        {
                            createIdx.CommandText = "ALTER TABLE `users` ADD INDEX `idx_department_id` (`department_id`);";
                            createIdx.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"确保 users.idx_department_id 索引时出错: {ex.Message}");
            }
        }
        /// <summary>
        /// 注册用户
        /// </summary>
        /// <param name="username"> 用户名 </param>
        /// <param name="password">密码</param>
        /// <param name="departmentId">部门id</param>
        /// <param name="departmentName">部门名称</param>
        /// <param name="fullName">全名</param>
        /// <param name="email">email</param>
        /// <param name="phone">电话</param>
        /// <param name="role">角色</param>
        /// <param name="createdBy">创建人</param>
        /// <returns></returns>
        public bool RegisterUser(string username, string password, int departmentId = 0, string departmentName = "",
             string? displayname = null, string? gender = null, string? email = null, string? phone = null, string? role = null, string? createdBy = null)
        {


            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password)) return false;
            var salt = GenerateSalt();
            var hash = ComputeHash(password, salt);

            using var conn = new MySqlConnection(ConnString());
            conn.Open();
            using var tx = conn.BeginTransaction();
            try
            {
                var cmd = conn.CreateCommand();
                cmd.Transaction = tx;
                cmd.CommandText = @"
                    INSERT INTO users
                        (username, password_hash, salt, display_name, gender, email, phone, role, department_id, department_name, is_active, created_by)
                    VALUES
                        (@u, @h, @s, @dn, @gender, @em, @ph, @rl, @d, @dname, 1, @cb);";
                cmd.Parameters.AddWithValue("@u", username);
                cmd.Parameters.AddWithValue("@h", hash);
                cmd.Parameters.AddWithValue("@s", salt);
                cmd.Parameters.AddWithValue("@dn", (object?)displayname ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@gender", (object?)gender ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@em", (object?)email ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ph", (object?)phone ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@rl", (object?)role ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@d", departmentId);
                cmd.Parameters.AddWithValue("@dname", departmentName ?? "");
                cmd.Parameters.AddWithValue("@cb", createdBy ?? username);
                cmd.ExecuteNonQuery();

                tx.Commit();
                return true;
            }
            catch (MySqlException mex)
            {
                try { tx.Rollback(); } catch { }
                LogManager.Instance.LogInfo($"RegisterUser MySQL 错误: {mex.Number} - {mex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                try { tx.Rollback(); } catch { }
                LogManager.Instance.LogInfo($"RegisterUser 失败: {ex.Message}");
                return false;
            }
        }
        /// <summary>
        /// 验证用户
        /// </summary>
        /// <param name="username">用户名</param>
        /// <param name="password">密码</param>
        /// <returns></returns>
        public bool AuthenticateUser(string username, string password)
        {
            if (string.IsNullOrEmpty(username)) return false;
            try
            {
                using var conn = new MySqlConnection(ConnString());
                conn.Open();

                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT password_hash, salt FROM users WHERE username=@u LIMIT 1;";
                cmd.Parameters.AddWithValue("@u", username);
                using var reader = cmd.ExecuteReader();
                if (!reader.Read())
                {
                    return false;
                }

                string? storedHash = reader.IsDBNull(reader.GetOrdinal("password_hash")) ? null : reader.GetString("password_hash");
                string? salt = reader.IsDBNull(reader.GetOrdinal("salt")) ? null : reader.GetString("salt");

                if (string.IsNullOrEmpty(storedHash))
                    return false;

                if (string.IsNullOrEmpty(salt))
                {
                    var hashedDirect = ComputeHash(password, "");
                    return string.Equals(storedHash, hashedDirect, StringComparison.OrdinalIgnoreCase);
                }
                else
                {
                    var hashed = ComputeHash(password, salt);
                    return string.Equals(storedHash, hashed, StringComparison.OrdinalIgnoreCase);
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"AuthenticateUser 出错: {ex.Message}");
                return false;
            }
        }
        #endregion

        #region 部门 CRUD 与用户分配
        /// <summary>
        /// 新增部门
        /// </summary>
        /// <param name="name">部门名称</param>
        /// <param name="displayName">部门显示名称</param>
        /// <param name="description">部门描述</param>
        /// <param name="cadCategoryId">CAD分类ID</param>
        /// <param name="sortOrder">排序</param>
        /// <returns></returns>
        public int AddDepartment(string name, string displayName = null, string description = null, int? cadCategoryId = null, int sortOrder = 0)
        {
            using var conn = new MySqlConnection(ConnString());
            conn.Open();
            using var tx = conn.BeginTransaction();
            try
            {
                var cmd = conn.CreateCommand();
                cmd.Transaction = tx;
                cmd.CommandText = @"INSERT INTO departments (cad_category_id, name, display_name, description, sort_order, is_active)
                                    VALUES (@cad, @name, @display, @desc, @so, 1); SELECT LAST_INSERT_ID();";
                cmd.Parameters.AddWithValue("@cad", (object?)cadCategoryId ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@name", name);
                cmd.Parameters.AddWithValue("@display", (object)displayName ?? (object)name);
                cmd.Parameters.AddWithValue("@desc", (object)description ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@so", sortOrder);
                var id = Convert.ToInt32(cmd.ExecuteScalar());
                tx.Commit();
                return id;
            }
            catch
            {
                try { tx.Rollback(); } catch { }
                return 0;
            }
        }
        /// <summary>
        /// 添加用户到部门
        /// </summary>
        /// <param name="username">用户名</param>
        /// <param name="departmentId">部门ID</param>
        /// <returns></returns>
        public bool AssignUserToDepartmentByUsername(string username, int departmentId)
        {
            var deptName = GetDepartmentNameByDeptTableId(departmentId);
            using var conn = new MySqlConnection(ConnString());
            conn.Open();
            using var tx = conn.BeginTransaction();
            try
            {
                var cmd = conn.CreateCommand();
                cmd.Transaction = tx;
                cmd.CommandText = "UPDATE users SET department_id=@d, department_name=@dn WHERE username=@u;";
                cmd.Parameters.AddWithValue("@d", departmentId);
                cmd.Parameters.AddWithValue("@dn", deptName);
                cmd.Parameters.AddWithValue("@u", username);
                var rows = cmd.ExecuteNonQuery();

                // 写入 department_users 映射（若表存在且尚未映射）
                try
                {
                    // 获取用户 id
                    var q = conn.CreateCommand();
                    q.Transaction = tx;
                    q.CommandText = "SELECT id FROM users WHERE username=@u LIMIT 1;";
                    q.Parameters.AddWithValue("@u", username);
                    var userIdObj = q.ExecuteScalar();
                    if (userIdObj != null && int.TryParse(userIdObj.ToString(), out int userId))
                    {
                        // 插入映射（若不存在）
                        var exists = conn.CreateCommand();
                        exists.Transaction = tx;
                        exists.CommandText = "SELECT COUNT(1) FROM department_users WHERE department_id=@d AND user_id=@uid;";
                        exists.Parameters.AddWithValue("@d", departmentId);
                        exists.Parameters.AddWithValue("@uid", userId);
                        var cnt = Convert.ToInt32(exists.ExecuteScalar() ?? 0);
                        if (cnt == 0)
                        {
                            var ins = conn.CreateCommand();
                            ins.Transaction = tx;
                            ins.CommandText = "INSERT INTO department_users (department_id, user_id) VALUES (@d, @uid);";
                            ins.Parameters.AddWithValue("@d", departmentId);
                            ins.Parameters.AddWithValue("@uid", userId);
                            ins.ExecuteNonQuery();
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Instance.LogInfo($"AssignUserToDepartmentByUsername 写入 department_users 失败: {ex.Message}");
                }

                tx.Commit();
                return rows > 0;
            }
            catch
            {
                try { tx.Rollback(); } catch { }
                return false;
            }
        }
        /// <summary>
        /// 修改部门
        /// </summary>
        /// <param name="id">部门ID</param>
        /// <param name="name">部门名称</param>
        /// <param name="displayName">部门显示名称</param>
        /// <param name="description">部门描述</param>
        /// <param name="sortOrder">排序</param>
        /// <param name="managerUserId">部门经理用户ID</param>
        /// <param name="isActive"></param>
        /// <returns></returns>
        public bool UpdateDepartment(int id, string name, string? displayName = null, string? description = null, int sortOrder = 0, int? managerUserId = null, bool? isActive = null)
        {
            using var conn = new MySqlConnection(ConnString());
            conn.Open();
            using var tx = conn.BeginTransaction();
            try
            {
                var cmd = conn.CreateCommand();
                cmd.Transaction = tx;
                cmd.CommandText = @"
                        UPDATE departments SET
                            name = @name,
                            display_name = @display,
                            description = @desc,
                            sort_order = @so,
                            manager_user_id = @mgr,
                            is_active = @ia,
                            updated_at = CURRENT_TIMESTAMP
                        WHERE id = @id;";
                cmd.Parameters.AddWithValue("@name", name);
                cmd.Parameters.AddWithValue("@display", (object?)displayName ?? (object)name);
                cmd.Parameters.AddWithValue("@desc", (object?)description ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@so", sortOrder);
                cmd.Parameters.AddWithValue("@mgr", (object?)managerUserId ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ia", isActive.HasValue ? (isActive.Value ? 1 : 0) : 1);
                cmd.Parameters.AddWithValue("@id", id);
                cmd.ExecuteNonQuery();
                tx.Commit();
                return true;
            }
            catch
            {
                try { tx.Rollback(); } catch { }
                return false;
            }
        }
        /// <summary>
        /// 删除部门
        /// </summary>
        /// <param name="id">部门ID</param>
        /// <returns></returns>
        public bool DeleteDepartment(int id)
        {
            using var conn = new MySqlConnection(ConnString());
            conn.Open();
            using var tx = conn.BeginTransaction();
            try
            {
                var cmd1 = conn.CreateCommand();
                cmd1.Transaction = tx;
                cmd1.CommandText = "UPDATE users SET department_id=0, department_name='' WHERE department_id=@id;";
                cmd1.Parameters.AddWithValue("@id", id);
                cmd1.ExecuteNonQuery();

                var cmd2 = conn.CreateCommand();
                cmd2.Transaction = tx;
                cmd2.CommandText = "DELETE FROM departments WHERE id=@id;";
                cmd2.Parameters.AddWithValue("@id", id);
                cmd2.ExecuteNonQuery();

                tx.Commit();
                return true;
            }
            catch
            {
                try { tx.Rollback(); } catch { }
                return false;
            }
        }

        /// <summary>
        /// 根据用户ID分配用户到部门
        /// </summary>
        /// <param name="userId">用户ID</param>
        /// <param name="departmentId">部门ID</param>
        /// <returns></returns>
        public bool AssignUserToDepartmentByUserId(int userId, int departmentId)
        {
            var deptName = GetDepartmentNameByDeptTableId(departmentId);
            using var conn = new MySqlConnection(ConnString());
            conn.Open();
            using var tx = conn.BeginTransaction();
            try
            {
                var cmd = conn.CreateCommand();
                cmd.Transaction = tx;
                cmd.CommandText = "UPDATE users SET department_id=@d, department_name=@dn WHERE id=@uid;";
                cmd.Parameters.AddWithValue("@d", departmentId);
                cmd.Parameters.AddWithValue("@dn", deptName);
                cmd.Parameters.AddWithValue("@uid", userId);
                var rows = cmd.ExecuteNonQuery();
                tx.Commit();
                return rows > 0;
            }
            catch
            {
                try { tx.Rollback(); } catch { }
                return false;
            }
        }
        /// <summary>
        /// 根据部门ID获取部门名称
        /// </summary>
        /// <param name="departmentTableId">部门ID</param>
        /// <returns></returns>
        private string GetDepartmentNameByDeptTableId(int departmentTableId)
        {
            using var conn = new MySqlConnection(ConnString());
            conn.Open();
            var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT name FROM departments WHERE id=@id LIMIT 1;";
            cmd.Parameters.AddWithValue("@id", departmentTableId);
            var result = cmd.ExecuteScalar();
            return result == null ? "" : result.ToString();
        }

        #endregion

        #region 查询
        /// <summary>
        /// 获取所有部门
        /// </summary>
        /// <returns></returns>
        public List<DepartmentModel> GetDepartmentsWithCounts()
        {
            var list = new List<DepartmentModel>();
            using var conn = new MySqlConnection(ConnString());
            conn.Open();
            var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                    SELECT d.id, d.cad_category_id, d.name, d.display_name, d.description, d.manager_user_id, d.sort_order, d.is_active,
                           (SELECT COUNT(1) FROM users u WHERE u.department_id = d.id) AS user_count
                    FROM departments d
                    ORDER BY d.sort_order, d.name;";
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(new DepartmentModel
                {
                    Id = reader.GetInt32("id"),
                    CadCategoryId = reader.IsDBNull(reader.GetOrdinal("cad_category_id")) ? (int?)null : reader.GetInt32("cad_category_id"),
                    Name = reader.GetString("name"),
                    DisplayName = reader.IsDBNull(reader.GetOrdinal("display_name")) ? reader.GetString("name") : reader.GetString("display_name"),
                    Description = reader.IsDBNull(reader.GetOrdinal("description")) ? "" : reader.GetString("description"),
                    ManagerUserId = reader.IsDBNull(reader.GetOrdinal("manager_user_id")) ? (int?)null : reader.GetInt32("manager_user_id"),
                    SortOrder = reader.IsDBNull(reader.GetOrdinal("sort_order")) ? 0 : reader.GetInt32("sort_order"),
                    IsActive = reader.IsDBNull(reader.GetOrdinal("is_active")) ? true : (reader.GetInt32("is_active") == 1),
                    UserCount = reader.IsDBNull(reader.GetOrdinal("user_count")) ? 0 : reader.GetInt32("user_count")
                });
            }
            return list;
        }
        /// <summary>
        /// 获取部门下的所有用户
        /// </summary>
        /// <param name="departmentId">部门ID</param>
        /// <returns></returns>
        public List<UserModel> GetUsersByDepartmentId(int departmentId)
        {
            var list = new List<UserModel>();
            using var conn = new MySqlConnection(ConnString());
            conn.Open();
            var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT id, username, gender, email, phone, role, is_active FROM users WHERE department_id=@d;";
            cmd.Parameters.AddWithValue("@d", departmentId);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(new UserModel
                {
                    Id = reader.GetInt32("id"),
                    Username = reader.GetString("username"),
                    Gender = reader.IsDBNull(reader.GetOrdinal("gender")) ? "" : reader.GetString("gender"),
                    Email = reader.IsDBNull(reader.GetOrdinal("email")) ? "" : reader.GetString("email"),
                    Phone = reader.IsDBNull(reader.GetOrdinal("phone")) ? "" : reader.GetString("phone"),
                    Role = reader.IsDBNull(reader.GetOrdinal("role")) ? "" : reader.GetString("role"),
                    IsActive = reader.IsDBNull(reader.GetOrdinal("is_active")) ? true : (reader.GetInt32("is_active") == 1)
                });
            }
            return list;
        }
        #endregion

        #region 辅助：哈希与盐
        /// <summary>
        /// 计算密码的哈希值
        /// </summary>
        /// <param name="password">密码</param>
        /// <param name="salt">盐值</param>
        /// <returns></returns>
        private static string ComputeHash(string password, string salt)
        {
            using var sha = SHA256.Create();
            var bytes = Encoding.UTF8.GetBytes((salt ?? "") + password);
            var hashed = sha.ComputeHash(bytes);
            return BitConverter.ToString(hashed).Replace("-", "").ToLowerInvariant();
        }
        /// <summary>
        /// 生成盐值,为密码加密16字节随机数的十六进制表示
        /// </summary>
        /// <returns>盐值</returns>
        private static string GenerateSalt()
        {
            var b = new byte[16];
            using var rng = RandomNumberGenerator.Create();
            rng.GetBytes(b);
            return BitConverter.ToString(b).Replace("-", "").ToLowerInvariant();
        }

        #endregion
    }

    /// <summary>
    /// 简单模型用于返回数据
    /// </summary>
    /// <remarks></remarks>
    public class DepartmentModel
    {
        public int Id { get; set; }
        public int? CadCategoryId { get; set; }
        public string? Name { get; set; }
        public string? DisplayName { get; set; }
        public string? Description { get; set; }
        public int? ManagerUserId { get; set; }
        public int SortOrder { get; set; }
        public bool IsActive { get; set; }
        public int UserCount { get; set; }
    }
    /// <summary>
    /// 用户模型
    /// </summary>
    public class UserModel
    {
        public int Id { get; set; }// 用户ID
        public string? Username { get; set; }// 用户名
        public string? Gender { get; set; }// 性别
        public string? Email { get; set; }// 邮箱
        public string? Phone { get; set; }// 手机
        public string? Role { get; set; }// 角色
        public bool IsActive { get; set; }// 是否激活
    }
}
