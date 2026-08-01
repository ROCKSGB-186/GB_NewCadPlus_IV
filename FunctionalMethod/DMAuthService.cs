// FunctionalMethod/DMAuthService.cs (优化版)
using Dm;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.UniFiedStandards;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 达梦数据库认证与组织架构管理服务
    /// 仅负责达梦数据库的操作，MySQL 逻辑请使用 MySqlAuthService
    /// </summary>
    public class DMAuthService
    {
        #region 基础字段与连接

        private readonly string _server;
        private readonly string _port;
        private readonly string _dbUser;
        private readonly string _dbPassword;
        private readonly string _database;
        private static bool _tablesEnsured = false;
        private static readonly object _ensureLock = new object();

        public DMAuthService(string server, string port, string user = "SYSDBA", string pwd = "SYSDBA")
        {
            _server = string.IsNullOrEmpty(server) ? "127.0.0.1" : server;
            _port = string.IsNullOrEmpty(port) ? "5236" : port;
            _dbUser = (user ?? "SYSDBA").ToUpper();
            _dbPassword = pwd;
            _database = VariableDictionary._dataBaseName ?? "CAD_SW_LIBRARY";
        }

        private string ConnString() =>
            $"Server={_server};Port={_port};User Id={_dbUser};Password={_dbPassword};Schema={_database};";

        private DmConnection CreateOpenConnection()
        {
            var conn = new DmConnection(ConnString());
            conn.Open();
            return conn;
        }

        private static void AddParam(IDbCommand cmd, string name, object value)
        {
            var p = cmd.CreateParameter();
            p.ParameterName = name;
            p.Value = value ?? DBNull.Value;
            cmd.Parameters.Add(p);
        }

        #endregion

        #region 表结构初始化（仅首次）

        public void EnsureAllTablesExist()
        {
            if (_tablesEnsured) return;
            lock (_ensureLock)
            {
                if (_tablesEnsured) return;

                EnsureTable("CAD_CATEGORIES", @"
                    CREATE TABLE CAD_CATEGORIES (
                        ID INT IDENTITY(1,1) PRIMARY KEY,
                        NAME VARCHAR(200) NOT NULL,
                        DISPLAY_NAME VARCHAR(200),
                        SORT_ORDER INT DEFAULT 0,
                        CREATED_AT DATETIME DEFAULT CURRENT_TIMESTAMP
                    )");

                EnsureTable("DEPARTMENTS", @"
                    CREATE TABLE DEPARTMENTS (
                        ID INT IDENTITY(1,1) PRIMARY KEY,
                        CAD_CATEGORY_ID INT NULL,
                        NAME VARCHAR(200) NOT NULL,
                        DISPLAY_NAME VARCHAR(200),
                        DESCRIPTION TEXT,
                        MANAGER_USER_ID INT NULL,
                        SORT_ORDER INT DEFAULT 0,
                        IS_ACTIVE TINYINT DEFAULT 1,
                        CREATED_AT DATETIME DEFAULT CURRENT_TIMESTAMP,
                        UPDATED_AT DATETIME
                    )");

                EnsureTable("USERS", @"
                    CREATE TABLE USERS (
                        ID INT IDENTITY(1,1) PRIMARY KEY,
                        USERNAME VARCHAR(100) NOT NULL UNIQUE,
                        PASSWORD_HASH VARCHAR(512) NOT NULL,
                        SALT VARCHAR(64) NOT NULL,
                        REAL_NAME VARCHAR(200),
                        GENDER VARCHAR(20),
                        PHONE VARCHAR(50),
                        EMAIL VARCHAR(200),
                        DEPARTMENT_ID INT DEFAULT 0,
                        DEPARTMENT_NAME VARCHAR(200),
                        ROLE VARCHAR(64),
                        IS_ACTIVE TINYINT DEFAULT 1,
                        CREATED_AT DATETIME DEFAULT CURRENT_TIMESTAMP
                    )");

                EnsureUsersColumnsExist();
                BackfillUsersNullFields();
                _tablesEnsured = true;
            }
        }

        private void EnsureTable(string tableName, string createSql)
        {
            using var conn = CreateOpenConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(1) FROM USER_TABLES WHERE TABLE_NAME = :t";
            AddParam(cmd, "t", (tableName ?? string.Empty).Trim().ToUpperInvariant());
            var exists = Convert.ToInt32(cmd.ExecuteScalar() ?? 0) > 0;
            if (exists) return;

            cmd.CommandText = createSql;
            cmd.ExecuteNonQuery();
        }

        private void EnsureUsersColumnsExist()
        {
            using var conn = CreateOpenConnection();
            EnsureUserColumn(conn, "REAL_NAME", "VARCHAR(200)");
            EnsureUserColumn(conn, "GENDER", "VARCHAR(20)");
            EnsureUserColumn(conn, "PHONE", "VARCHAR(50)");
            EnsureUserColumn(conn, "EMAIL", "VARCHAR(200)");
        }

        private void EnsureUserColumn(DmConnection conn, string columnName, string dataTypeSql)
        {
            try
            {
                using var existsCmd = conn.CreateCommand();
                existsCmd.CommandText = "SELECT COUNT(1) FROM USER_TAB_COLUMNS WHERE TABLE_NAME='USERS' AND COLUMN_NAME=:c";
                AddParam(existsCmd, "c", columnName);
                if (Convert.ToInt32(existsCmd.ExecuteScalar() ?? 0) > 0) return;

                using var alterCmd = conn.CreateCommand();
                alterCmd.CommandText = $"ALTER TABLE USERS ADD {columnName} {dataTypeSql}";
                alterCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"EnsureUserColumn({columnName}): {ex.Message}");
            }
        }

        private void BackfillUsersNullFields()
        {
            try
            {
                using var conn = CreateOpenConnection();
                // USERS 表没有 DISPLAY_NAME 字段，直接使用 USERNAME 回填 REAL_NAME。
                ExecuteNonQuery(conn, "UPDATE USERS SET REAL_NAME = USERNAME WHERE REAL_NAME IS NULL");
                ExecuteNonQuery(conn, "UPDATE USERS SET GENDER = '无信息' WHERE GENDER IS NULL");
                ExecuteNonQuery(conn, "UPDATE USERS SET PHONE = '未填写' WHERE PHONE IS NULL");
                ExecuteNonQuery(conn, "UPDATE USERS SET EMAIL = '未填写' WHERE EMAIL IS NULL");
                ExecuteNonQuery(conn, "UPDATE USERS SET ROLE = 'user' WHERE ROLE IS NULL");
                ExecuteNonQuery(conn, "UPDATE USERS SET DEPARTMENT_NAME = '未分配' WHERE DEPARTMENT_NAME IS NULL");
                ExecuteNonQuery(conn, "UPDATE USERS SET IS_ACTIVE = 1 WHERE IS_ACTIVE IS NULL");
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"BackfillUsersNullFields: {ex.Message}");
            }
        }

        private static void ExecuteNonQuery(DmConnection conn, string sql)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.ExecuteNonQuery();
        }

        #endregion

        #region 认证

        public bool AuthenticateUser(string username, string password)
        {
            EnsureAllTablesExist();
            if (string.IsNullOrEmpty(username)) return false;
            try
            {
                using var conn = CreateOpenConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT PASSWORD_HASH, SALT FROM USERS WHERE USERNAME = :u AND IS_ACTIVE = 1";
                AddParam(cmd, "u", username);
                using var r = cmd.ExecuteReader();
                if (r.Read())
                {
                    string dbHash = r.GetString(0);
                    string dbSalt = r.GetString(1);
                    return VerifyPassword(password, dbSalt, dbHash);
                }
            }
            catch { }
            return false;
        }

        #endregion

        #region 用户 CRUD

        [Obsolete("请使用 AddUser 方法代替")]
        public bool RegisterUser(string username, string password, int departmentId = 0, string departmentName = "")
        {
            return AddUser(username, password, departmentId, departmentName);
        }

        public bool AddUser(string username, string password, int departmentId = 0, string departmentName = "",
            string role = "user", bool isActive = true, string? realName = null, string? gender = null,
            string? phone = null, string? email = null)
        {
            EnsureAllTablesExist();
            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
                return false;

            var userNameTrim = username.Trim();
            var salt = GenerateSalt();
            var hash = ComputeHash(password, salt);

            try
            {
                using var conn = CreateOpenConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = @"INSERT INTO USERS
(USERNAME, PASSWORD_HASH, SALT, REAL_NAME, GENDER, PHONE, EMAIL, DEPARTMENT_ID, DEPARTMENT_NAME, ROLE, IS_ACTIVE, CREATED_AT)
VALUES (:u, :h, :s, :realName, :g, :p, :e, :d, :deptName, :r, :ia, SYSDATE)";

                AddParam(cmd, "u", userNameTrim);
                AddParam(cmd, "h", hash);
                AddParam(cmd, "s", salt);
                AddParam(cmd, "realName", UserFieldNormalizer.NormalizeRealName(realName, userNameTrim));
                AddParam(cmd, "g", UserFieldNormalizer.NormalizeGender(gender));
                AddParam(cmd, "p", UserFieldNormalizer.NormalizePhone(phone));
                AddParam(cmd, "e", UserFieldNormalizer.NormalizeEmail(email));
                AddParam(cmd, "d", departmentId);
                AddParam(cmd, "deptName", UserFieldNormalizer.NormalizeDepartmentName(departmentName));
                AddParam(cmd, "r", UserFieldNormalizer.NormalizeRole(role));
                AddParam(cmd, "ia", isActive ? 1 : 0);

                return cmd.ExecuteNonQuery() > 0;
            }
            catch { return false; }
        }

        public bool UpdateUser(int userId, string username, string role, bool isActive,
            int? departmentId = null, string? departmentName = null, string? newPassword = null,
            string? realName = null, string? gender = null, string? phone = null, string? email = null)
        {
            EnsureAllTablesExist();
            if (userId <= 0 || string.IsNullOrWhiteSpace(username)) return false;

            try
            {
                using var conn = CreateOpenConnection();

                string finalDeptName = UserFieldNormalizer.NormalizeDepartmentName(departmentName);
                if (departmentId.HasValue && departmentId.Value > 0 && string.IsNullOrWhiteSpace((departmentName ?? string.Empty).Trim()))
                {
                    using var findDept = conn.CreateCommand();
                    findDept.CommandText = "SELECT NAME FROM DEPARTMENTS WHERE ID = :id";
                    AddParam(findDept, "id", departmentId.Value);
                    finalDeptName = UserFieldNormalizer.NormalizeDepartmentName(Convert.ToString(findDept.ExecuteScalar()));
                }

                var finalRealName = UserFieldNormalizer.NormalizeRealName(realName, username);
                var finalGender = UserFieldNormalizer.NormalizeGender(gender);
                var finalPhone = UserFieldNormalizer.NormalizePhone(phone);
                var finalEmail = UserFieldNormalizer.NormalizeEmail(email);

                using var cmd = conn.CreateCommand();

                if (string.IsNullOrWhiteSpace(newPassword))
                {
                    cmd.CommandText = @"UPDATE USERS SET USERNAME=:u, REAL_NAME=:realName, GENDER=:gender,
                        PHONE=:phone, EMAIL=:email, ROLE=:r, IS_ACTIVE=:ia, DEPARTMENT_ID=:d,
                        DEPARTMENT_NAME=:dn WHERE ID=:id";
                }
                else
                {
                    var salt = GenerateSalt();
                    var hash = ComputeHash(newPassword, salt);
                    cmd.CommandText = @"UPDATE USERS SET USERNAME=:u, PASSWORD_HASH=:h, SALT=:s,
                        REAL_NAME=:realName, GENDER=:gender, PHONE=:phone, EMAIL=:email,
                        ROLE=:r, IS_ACTIVE=:ia, DEPARTMENT_ID=:d, DEPARTMENT_NAME=:dn WHERE ID=:id";
                    AddParam(cmd, "h", hash);
                    AddParam(cmd, "s", salt);
                }

                AddParam(cmd, "u", username.Trim());
                AddParam(cmd, "realName", finalRealName);
                AddParam(cmd, "gender", finalGender);
                AddParam(cmd, "phone", finalPhone);
                AddParam(cmd, "email", finalEmail);
                AddParam(cmd, "r", UserFieldNormalizer.NormalizeRole(role));
                AddParam(cmd, "ia", isActive ? 1 : 0);
                AddParam(cmd, "d", departmentId ?? 0);
                AddParam(cmd, "dn", finalDeptName);
                AddParam(cmd, "id", userId);

                return cmd.ExecuteNonQuery() > 0;
            }
            catch { return false; }
        }

        public bool DeleteUser(int userId)
        {
            EnsureAllTablesExist();
            if (userId <= 0) return false;
            try
            {
                using var conn = CreateOpenConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "DELETE FROM USERS WHERE ID = :id";
                AddParam(cmd, "id", userId);
                return cmd.ExecuteNonQuery() > 0;
            }
            catch { return false; }
        }

        public string? GetUserRole(string username)
        {
            if (string.IsNullOrWhiteSpace(username)) return null;
            try
            {
                using var conn = CreateOpenConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT ROLE FROM USERS WHERE UPPER(USERNAME) = :u AND IS_ACTIVE = 1";
                AddParam(cmd, "u", username.Trim().ToUpperInvariant());
                var obj = cmd.ExecuteScalar();
                return (obj == null || obj == DBNull.Value) ? null : Convert.ToString(obj);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"GetUserRole: {ex.Message}");
                return null;
            }
        }

        public List<UserModel> GetUsersByDepartmentId(int departmentId)
        {
            EnsureAllTablesExist();
            var list = new List<UserModel>();
            using var conn = CreateOpenConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT ID, USERNAME, REAL_NAME, GENDER, EMAIL, PHONE, ROLE, DEPARTMENT_NAME, IS_ACTIVE FROM USERS WHERE DEPARTMENT_ID = :d ORDER BY ID";
            AddParam(cmd, "d", departmentId);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var realName = r.IsDBNull(2) ? string.Empty : r.GetString(2);
                list.Add(new UserModel
                {
                    Id = r.GetInt32(0),
                    Username = r.GetString(1),
                    RealName = realName,
                    DisplayName = realName,
                    FullName = realName,
                    Gender = r.IsDBNull(3) ? string.Empty : r.GetString(3),
                    Email = r.IsDBNull(4) ? string.Empty : r.GetString(4),
                    Phone = r.IsDBNull(5) ? string.Empty : r.GetString(5),
                    Role = r.IsDBNull(6) ? string.Empty : r.GetString(6),
                    DepartmentName = r.IsDBNull(7) ? string.Empty : r.GetString(7),
                    IsActive = r.GetInt16(8) == 1
                });
            }
            return list;
        }

        public bool AssignUserToDepartmentByUsername(string username, int departmentId)
        {
            EnsureAllTablesExist();
            if (string.IsNullOrEmpty(username)) return false;
            using var conn = CreateOpenConnection();
            string dname = "";
            using (var gdn = conn.CreateCommand())
            {
                gdn.CommandText = "SELECT NAME FROM DEPARTMENTS WHERE ID = :id";
                AddParam(gdn, "id", departmentId);
                dname = gdn.ExecuteScalar()?.ToString() ?? "";
            }
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE USERS SET DEPARTMENT_ID = :d, DEPARTMENT_NAME = :dn WHERE USERNAME = :u";
            AddParam(cmd, "d", departmentId);
            AddParam(cmd, "dn", dname);
            AddParam(cmd, "u", username);
            return cmd.ExecuteNonQuery() > 0;
        }

        #endregion

        #region 部门 CRUD

        public int AddDepartment(string name, string? displayName = null, string? description = null,
            int? managerUserId = null, int sortOrder = 0)
        {
            EnsureAllTablesExist();
            if (string.IsNullOrWhiteSpace(name)) return 0;
            using var conn = CreateOpenConnection();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"INSERT INTO DEPARTMENTS
(NAME, DISPLAY_NAME, DESCRIPTION, MANAGER_USER_ID, SORT_ORDER, IS_ACTIVE, CREATED_AT)
VALUES (:n, :dn, :desc, :mgr, :so, 1, SYSDATE)";
                AddParam(cmd, "n", name.Trim());
                AddParam(cmd, "dn", displayName ?? name);
                AddParam(cmd, "desc", description);
                AddParam(cmd, "mgr", managerUserId);
                AddParam(cmd, "so", sortOrder);
                cmd.ExecuteNonQuery();
            }
            using var tid = conn.CreateCommand();
            tid.CommandText = "SELECT ID FROM DEPARTMENTS WHERE NAME = :n ORDER BY ID DESC";
            AddParam(tid, "n", name.Trim());
            using var r = tid.ExecuteReader();
            if (r.Read()) return r.GetInt32(0);
            return 0;
        }

        public bool UpdateDepartment(int id, string name, string? displayName = null, string? description = null,
            int sortOrder = 0, int? managerUserId = null, bool? isActive = null)
        {
            EnsureAllTablesExist();
            if (id <= 0) return false;
            using var conn = CreateOpenConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"UPDATE DEPARTMENTS SET NAME=:n, DISPLAY_NAME=:dn, DESCRIPTION=:desc,
MANAGER_USER_ID=:mgr, SORT_ORDER=:so, IS_ACTIVE=:ia, UPDATED_AT=SYSDATE WHERE ID=:id";
            AddParam(cmd, "n", name);
            AddParam(cmd, "dn", displayName);
            AddParam(cmd, "desc", description);
            AddParam(cmd, "mgr", managerUserId);
            AddParam(cmd, "so", sortOrder);
            AddParam(cmd, "ia", (isActive ?? true) ? 1 : 0);
            AddParam(cmd, "id", id);
            return cmd.ExecuteNonQuery() > 0;
        }

        public bool DeleteDepartment(int id)
        {
            EnsureAllTablesExist();
            if (id <= 0) return false;
            using var conn = CreateOpenConnection();
            using var tx = conn.BeginTransaction();
            try
            {
                using (var c1 = conn.CreateCommand())
                {
                    c1.Transaction = tx;
                    c1.CommandText = "UPDATE USERS SET DEPARTMENT_ID = 0, DEPARTMENT_NAME = '' WHERE DEPARTMENT_ID = :id";
                    AddParam(c1, "id", id);
                    c1.ExecuteNonQuery();
                }
                using (var c2 = conn.CreateCommand())
                {
                    c2.Transaction = tx;
                    c2.CommandText = "DELETE FROM DEPARTMENTS WHERE ID = :id";
                    AddParam(c2, "id", id);
                    var result = c2.ExecuteNonQuery() > 0;
                    tx.Commit();
                    return result;
                }
            }
            catch
            {
                tx.Rollback();
                return false;
            }
        }

        public List<DepartmentModel> GetDepartmentsWithCounts()
        {
            EnsureAllTablesExist();
            var list = LoadDepartmentsWithCountsCore();
            if (list.Count > 0) return list;

            LogManager.Instance.LogInfo("GetDepartmentsWithCounts: DEPARTMENTS 为空，尝试从 CAD_CATEGORIES 同步");
            try { SyncDepartmentsFromCadCategories(); }
            catch (Exception ex) { LogManager.Instance.LogInfo($"同步部门失败: {ex.Message}"); }

            list = LoadDepartmentsWithCountsCore();
            if (list.Count > 0) return list;

            LogManager.Instance.LogInfo("GetDepartmentsWithCounts: 改为 CAD_CATEGORIES 兜底");
            return LoadDepartmentsFromCategoriesFallback();
        }

        private List<DepartmentModel> LoadDepartmentsWithCountsCore()
        {
            var list = new List<DepartmentModel>();
            using var conn = CreateOpenConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"SELECT D.ID, D.CAD_CATEGORY_ID, D.NAME, D.DISPLAY_NAME,
                        D.DESCRIPTION, D.SORT_ORDER, D.IS_ACTIVE,
                        (SELECT COUNT(1) FROM USERS U WHERE U.DEPARTMENT_ID = D.ID) AS USER_COUNT
                        FROM DEPARTMENTS D ORDER BY D.SORT_ORDER";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var realName = r.IsDBNull(3) ? (r.IsDBNull(2) ? string.Empty : Convert.ToString(r.GetValue(2)))
                    : Convert.ToString(r.GetValue(3));
                var name = r.IsDBNull(2) ? string.Empty : Convert.ToString(r.GetValue(2));

                list.Add(new DepartmentModel
                {
                    Id = Convert.ToInt32(r.GetValue(0)),
                    CadCategoryId = r.IsDBNull(1) ? null : (int?)Convert.ToInt32(r.GetValue(1)),
                    Name = name,
                    RealName = realName,
                    DisplayName = string.IsNullOrWhiteSpace(realName) ? name : realName, // ⭐ 填充 DisplayName
                    Description = r.IsDBNull(4) ? string.Empty : Convert.ToString(r.GetValue(4)),
                    SortOrder = r.IsDBNull(5) ? 0 : Convert.ToInt32(r.GetValue(5)),
                    IsActive = !r.IsDBNull(6) && Convert.ToInt32(r.GetValue(6)) == 1,
                    UserCount = r.IsDBNull(7) ? 0 : Convert.ToInt32(r.GetValue(7))
                });
            }
            return list;
        }

        private List<DepartmentModel> LoadDepartmentsFromCategoriesFallback()
        {
            var list = new List<DepartmentModel>();
            using var conn = CreateOpenConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT ID, NAME, DISPLAY_NAME, SORT_ORDER FROM CAD_CATEGORIES ORDER BY SORT_ORDER";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var id = Convert.ToInt32(r.GetValue(0));
                var name = r.IsDBNull(1) ? string.Empty : Convert.ToString(r.GetValue(1));
                var displayName = r.IsDBNull(2) ? name : Convert.ToString(r.GetValue(2));

                list.Add(new DepartmentModel
                {
                    Id = id,
                    CadCategoryId = id,
                    Name = name,
                    RealName = displayName,
                    DisplayName = string.IsNullOrWhiteSpace(displayName) ? name : displayName, // ⭐ 填充 DisplayName
                    Description = string.Empty,
                    SortOrder = r.IsDBNull(3) ? 0 : Convert.ToInt32(r.GetValue(3)),
                    IsActive = true,
                    UserCount = 0
                });
            }
            return list;
        }

        #endregion

        #region 部门同步（仅达梦）

        public void SyncDepartmentsFromCadCategories()
        {
            EnsureAllTablesExist();
            using var conn = CreateOpenConnection();
            using var tx = conn.BeginTransaction();
            try
            {
                // 读取分类
                var list = new List<(int Id, string Name, string Display, int So)>();
                using (var sel = conn.CreateCommand())
                {
                    sel.Transaction = tx;
                    sel.CommandText = "SELECT ID, NAME, DISPLAY_NAME, SORT_ORDER FROM CAD_CATEGORIES";
                    using var r = sel.ExecuteReader();
                    while (r.Read())
                    {
                        int id = r.GetInt32(0);
                        string name = r.IsDBNull(1) ? "未命名分类" : r.GetString(1);
                        string disp = r.IsDBNull(2) ? name : r.GetString(2);
                        int sort = r.IsDBNull(3) ? 0 : Convert.ToInt32(r.GetValue(3));
                        list.Add((id, name, disp, sort));
                    }
                }

                if (list.Count == 0) { tx.Commit(); return; }
                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories: 读取到 {list.Count} 条分类");

                var affectedRows = 0;
                foreach (var c in list)
                {
                    int? matchedDepartmentId = null;
                    string matchMode = "";

                    using (var chkByCategory = conn.CreateCommand())
                    {
                        chkByCategory.Transaction = tx;
                        chkByCategory.CommandText = "SELECT ID FROM DEPARTMENTS WHERE CAD_CATEGORY_ID = :id";
                        AddParam(chkByCategory, "id", c.Id);
                        var existingId = chkByCategory.ExecuteScalar();
                        if (existingId != null && existingId != DBNull.Value)
                        {
                            matchedDepartmentId = Convert.ToInt32(existingId);
                            matchMode = "CAD_CATEGORY_ID";
                        }
                    }

                    if (!matchedDepartmentId.HasValue)
                    {
                        using var chkByName = conn.CreateCommand();
                        chkByName.Transaction = tx;
                        chkByName.CommandText = "SELECT ID FROM DEPARTMENTS WHERE NAME = :name";
                        AddParam(chkByName, "name", c.Name);
                        var existingId = chkByName.ExecuteScalar();
                        if (existingId != null && existingId != DBNull.Value)
                        {
                            matchedDepartmentId = Convert.ToInt32(existingId);
                            matchMode = "NAME";
                        }
                    }

                    using var cmd = conn.CreateCommand();
                    cmd.Transaction = tx;
                    if (matchedDepartmentId.HasValue)
                    {
                        cmd.CommandText = @"UPDATE DEPARTMENTS SET CAD_CATEGORY_ID=:cid, NAME=:n,
DISPLAY_NAME=:d, SORT_ORDER=:s, IS_ACTIVE=1, UPDATED_AT=SYSDATE WHERE ID=:deptId";
                        AddParam(cmd, "deptId", matchedDepartmentId.Value);
                    }
                    else
                    {
                        cmd.CommandText = @"INSERT INTO DEPARTMENTS
(CAD_CATEGORY_ID, NAME, DISPLAY_NAME, SORT_ORDER, IS_ACTIVE, CREATED_AT)
VALUES (:cid, :n, :d, :s, 1, SYSDATE)";
                    }
                    AddParam(cmd, "cid", c.Id);
                    AddParam(cmd, "n", c.Name);
                    AddParam(cmd, "d", string.IsNullOrEmpty(c.Display) ? c.Name : c.Display);
                    AddParam(cmd, "s", c.So);
                    affectedRows += cmd.ExecuteNonQuery();

                    if (matchedDepartmentId.HasValue)
                        LogManager.Instance.LogInfo($"同步: 分类 {c.Id}-{c.Name} 通过 {matchMode} 命中部门 ID={matchedDepartmentId.Value}");
                }

                tx.Commit();
                LogManager.Instance.LogInfo($"SyncDepartmentsFromCadCategories: 完成，受影响 {affectedRows} 行");
            }
            catch
            {
                tx.Rollback();
                throw;
            }
        }

        #endregion

        #region 工具方法

        public List<TableInfo> GetTables()
        {
            EnsureAllTablesExist();
            var list = new List<TableInfo>();
            using var conn = CreateOpenConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"SELECT T.TABLE_NAME, C.COMMENTS FROM USER_TABLES T
LEFT JOIN USER_TAB_COMMENTS C ON T.TABLE_NAME = C.TABLE_NAME ORDER BY T.TABLE_NAME";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                list.Add(new TableInfo
                {
                    TableName = r.GetString(0),
                    TableComment = r.IsDBNull(1) ? "" : r.GetString(1)
                });
            }
            return list;
        }

        #endregion

        #region 加密（PBKDF2）

        private static string GenerateSalt()
        {
            byte[] bytes = new byte[32];
            using var rng = RandomNumberGenerator.Create();  // 使用新的 API，但输出格式不变
            rng.GetBytes(bytes);
            return Convert.ToBase64String(bytes);
        }

        private static string ComputeHash(string password, string salt)
        {
            using var sha = SHA256.Create();
            byte[] bytes = Encoding.UTF8.GetBytes(password + salt);
            return Convert.ToBase64String(sha.ComputeHash(bytes));
        }

        private static bool VerifyPassword(string password, string salt, string storedHash)
        {
            return ComputeHash(password, salt) == storedHash;
        }

        #endregion
    }
}
