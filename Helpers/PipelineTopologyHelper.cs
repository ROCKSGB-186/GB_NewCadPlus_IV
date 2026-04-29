using Autodesk.AutoCAD.DatabaseServices; // 引入 ObjectId 等数据库类型 // 中文注释
using Autodesk.AutoCAD.Geometry; // 引入 Point3d 类型 // 中文注释
using GB_NewCadPlus_IV.FunctionalMethod;
using System; // 使用系统命名空间，提供基础类型和异常等 // 中文注释
using System.Collections.Generic; // 使用泛型集合类型 // 中文注释
using System.Linq; // 使用 LINQ 扩展方法（如需） // 中文注释
using System.Text; // 文本处理（保留） // 中文注释
using System.Threading.Tasks; // 任务并发（保留） // 中文注释

namespace GB_NewCadPlus_IV.Helpers // 命名空间与原文件保持一致 // 中文注释
{
    /// <summary>
    /// 管道拓扑优化助手：提供管线落图后的自动并线功能 // 中文注释
    /// </summary>
    internal static class PipelineTopologyHelper // 静态类，包含若干静态辅助方法 // 中文注释
    {

        /// <summary>
        /// 主入口：在“进口/出口管线”落图成功后调用（第一批只做端点并线） // 中文注释
        /// </summary>
        /// <param name="tr">数据库事务对象，用于操作管线数据</param>
        /// <param name="newPipeId">新插入的管线ID</param>
        /// <param name="pipeRole">管线角色，如“进口”或“出口”</param>
        public static void PostProcessAfterPipePlaced(DBTrans tr, ObjectId newPipeId, string pipeRole)
        {
            // 事务为空直接返回，避免空引用异常
            if (tr == null) return;
            // 新管线ID无效直接返回，避免无效对象访问
            if (newPipeId == ObjectId.Null) return;

            try
            {
                // 当前正在处理的管线ID，初始为刚插入的新管线
                ObjectId workingPipeId = newPipeId;
                // 是否继续链式并线的开关
                bool continueMerge = true;
                // 安全计数器，防止异常数据导致死循环
                int safeLoop = 0;
                // 最大循环次数，通常够用且能避免卡死
                const int maxLoop = 30;

                // 本地函数-统计字典中非空值数量，用于判断“参数完整度”
                int CountNonEmpty(Dictionary<string, string> map)
                {
                    // 空字典返回0
                    if (map == null) return 0;
                    // 计数器
                    int n = 0;
                    // 遍历所有键值对
                    foreach (var kv in map)
                    {
                        // 值非空白则计数+1
                        if (!string.IsNullOrWhiteSpace(kv.Value)) n++;
                    }
                    // 返回统计值
                    return n;
                }

                // 外层循环，每并线成功一次就再扫描一轮（支持A-B-C连续并线）
                while (continueMerge && safeLoop < maxLoop)
                {
                    // 循环次数+1
                    safeLoop++;
                    // 先置为false，只有本轮真的并成功才置回true
                    continueMerge = false;

                    // 读取当前工作管线的轴线与元数据
                    if (!TryGetPipeAxisAndMeta(tr, workingPipeId, out Point3d aStart, out Point3d aEnd, out Dictionary<string, string> aMeta))
                    {
                        // 读取失败则终止后处理
                        LogManager.Instance.LogInfo("\nPostProcessAfterPipePlaced: 读取当前管线轴线失败，终止并线。");
                        break;
                    }

                    // 角色兜底，确保规则判定时有 PipeRole
                    if (!aMeta.ContainsKey("PipeRole"))
                    {
                        // 若上游传入空，则写空串
                        aMeta["PipeRole"] = pipeRole ?? string.Empty;
                    }

                    // 按当前轴线查找附近候选管线（粗筛）
                    List<ObjectId> candidates = FindNearbyPipes(tr, aStart, aEnd, workingPipeId);
                    // 没有候选则本轮结束
                    if (candidates == null || candidates.Count == 0)
                    {
                        // 无候选时直接结束外层循环
                        break;
                    }

                    // 遍历候选列表，寻找第一个可并对象
                    foreach (ObjectId candidateId in candidates)
                    {
                        // 空ID跳过
                        if (candidateId == ObjectId.Null) continue;
                        // 自身跳过
                        if (candidateId == workingPipeId) continue;

                        // 读取候选对象的轴线和元数据
                        if (!TryGetPipeAxisAndMeta(tr, candidateId, out Point3d bStart, out Point3d bEnd, out Dictionary<string, string> bMeta))
                        {
                            // 候选读取失败则检查下一个
                            continue;
                        }

                        // 候选角色兜底（防止缺失导致规则失效）
                        if (!bMeta.ContainsKey("PipeRole"))
                        {
                            // 优先沿用传入角色，后续可改为按块名推断
                            bMeta["PipeRole"] = pipeRole ?? string.Empty;
                        }

                        // 进行“端点接触+近似共线+业务兼容”判定
                        if (!CanMergeAtEndpoint(aStart, aEnd, bStart, bEnd, aMeta, bMeta))
                        {
                            // 不满足并线条件，继续检查下一个候选
                            continue;
                        }

                        // 组合四个端点，后续取最远两点作为合并后的新轴线
                        Point3d[] pts = new[] { aStart, aEnd, bStart, bEnd };
                        // 最大距离初始值
                        double maxDist = -1.0;
                        // 合并后起点默认值
                        Point3d mergedStart = aStart;
                        // 合并后终点默认值
                        Point3d mergedEnd = aEnd;

                        // 双层循环找四点中最远点对
                        for (int i = 0; i < pts.Length; i++)
                        {
                            // 内层从 i+1 开始，避免重复比较
                            for (int j = i + 1; j < pts.Length; j++)
                            {
                                // 计算点距
                                double d = pts[i].DistanceTo(pts[j]);
                                // 若更大则更新最远点对
                                if (d > maxDist)
                                {
                                    // 更新最大距离
                                    maxDist = d;
                                    // 更新合并起点
                                    mergedStart = pts[i];
                                    // 更新合并终点
                                    mergedEnd = pts[j];
                                }
                            }
                        }

                        // 若最远距离过小，视为异常，不执行并线
                        if (maxDist <= 1e-8)
                        {
                            // 异常情况继续下一候选
                            continue;
                        }

                        // 按“参数完整度”决定保留哪条线（参数多者优先）
                        int aScore = CountNonEmpty(aMeta);
                        // 候选参数完整度
                        int bScore = CountNonEmpty(bMeta);

                        // 定义保留对象ID
                        ObjectId keepId;
                        // 定义删除对象ID
                        ObjectId eraseId;
                        // 定义需要合并回保留对象的来源属性
                        Dictionary<string, string> inheritMap;

                        // 若候选参数更完整，则保留候选并删除当前
                        if (bScore > aScore)
                        {
                            // 保留候选对象
                            keepId = candidateId;
                            // 删除当前对象
                            eraseId = workingPipeId;
                            // 把当前对象属性并入保留对象
                            inheritMap = aMeta;
                        }
                        else
                        {
                            // 默认保留当前对象
                            keepId = workingPipeId;
                            // 删除候选对象
                            eraseId = candidateId;
                            // 把候选对象属性并入保留对象
                            inheritMap = bMeta;
                        }

                        // 执行几何并线（更新保留对象几何并删除另一条）
                        bool mergedOk = TryMergePipes(tr, keepId, eraseId, mergedStart, mergedEnd);
                        // 并线失败则继续下一个候选
                        if (!mergedOk) continue;

                        // 并线成功后执行属性合并（空值不覆盖非空值）
                        MergePipeProperties(tr, keepId, inheritMap);

                        // 更新工作对象为当前保留对象，支持下一轮链式并线
                        workingPipeId = keepId;
                        // 标记本轮有并线，外层继续
                        continueMerge = true;
                        // 本轮已处理一次并线，跳出候选循环，进入下一轮重新扫描
                        break;
                    }
                }

                // 输出完成提示
                Env.Editor?.WriteMessage("\n管线端点并线后处理完成。");
            }
            catch (Exception ex)
            {
                // 统一容错，不阻断主命令，记录日志便于排查
                LogManager.Instance.LogInfo($"\nPostProcessAfterPipePlaced 执行异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 尝试读取管线的轴线（起点/终点）和属性元数据（块属性 + 扩展字典），并转换到世界坐标系
        /// </summary>
        /// <param name="tr">事务对象</param>
        /// <param name="pipeId">管线对象的 ObjectId</param>
        /// <param name="start">输出参数：管线轴线起点</param>
        /// <param name="end">输出参数：管线轴线终点</param>
        /// <param name="meta">输出参数：管线属性元数据字典</param>
        /// <returns>成功返回 true，失败返回 false</returns>
        private static bool TryGetPipeAxisAndMeta(DBTrans tr, ObjectId pipeId, out Point3d start, out Point3d end, out Dictionary<string, string> meta)
        {
            // 先给输出参数默认值，防止任何分支未赋值
            start = Point3d.Origin;
            // 终点默认值
            end = Point3d.Origin;
            // 初始化属性字典，忽略大小写
            meta = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 基础判空，事务为空直接失败
            if (tr == null) return false;
            // ObjectId 为空直接失败
            if (pipeId == ObjectId.Null) return false;

            try
            {
                // 按块参照读取，因为你的管线是“PL线制作的属性块”
                var br = tr.GetObject(pipeId, OpenMode.ForRead) as BlockReference;
                // 不是块参照则失败
                if (br == null) return false;
                // 已删除对象不处理
                if (br.IsErased) return false;

                // 读取块定义记录（优先动态块记录，普通块也兼容）
                ObjectId btrId = br.DynamicBlockTableRecord != ObjectId.Null ? br.DynamicBlockTableRecord : br.BlockTableRecord;
                // 打开块表记录
                var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                // 块记录无效则失败
                if (btr == null) return false;

                // 记录块名到元数据，后续可用于过滤“进口/出口”类型
                meta["BlockName"] = btr.Name ?? string.Empty;
                // 记录图层
                meta["Layer"] = br.Layer ?? string.Empty;
                // 记录句柄
                meta["Handle"] = br.Handle.ToString();                

                // 在块定义里查找“最像管道主体”的 Polyline（优先非闭合，按端点跨度最大）
                Polyline bodyPl = null; // 最终选中的主体线
                double bestScore = double.MinValue; // 评分越大越优先

                foreach (ObjectId entId in btr) // 遍历块定义实体
                {
                    var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity; // 读取实体
                    if (ent is not Polyline pl) continue; // 只看Polyline

                    if (pl.NumberOfVertices < 2) continue; // 顶点不足跳过

                    Point3d p0 = pl.GetPoint3dAt(0); // 首点
                    Point3d pN = pl.GetPoint3dAt(pl.NumberOfVertices - 1); // 末点

                    double span = p0.DistanceTo(pN); // 端点跨度（主体线一般更大）
                    double score = span; // 基础分

                    if (!pl.Closed) score += 1000000.0; // 强优先非闭合线（箭头通常闭合）

                    if (score > bestScore) // 更新最优候选
                    {
                        bestScore = score; // 更新评分
                        bodyPl = pl; // 更新主体
                    }
                }

                // 记录最大长度用于比较
                double maxLen = -1.0;

                // 遍历块定义中的实体
                foreach (ObjectId entId in btr)
                {
                    // 读取实体
                    var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                    // 过滤空实体
                    if (ent == null) continue;

                    // 只处理 Polyline
                    if (ent is Polyline pl)
                    {
                        // 保护性读取长度
                        double len = 0.0;
                        // 捕获极端异常，避免中断
                        try { len = pl.Length; } catch { len = 0.0; }

                        // 长度更大则作为候选主体
                        if (len > maxLen)
                        {
                            // 更新最大长度
                            maxLen = len;
                            // 更新主体 Polyline
                            bodyPl = pl;
                        }
                    }
                }

                // 没找到 Polyline 主体则失败
                if (bodyPl == null) return false;
                // 顶点不足 2 个无法构成轴线
                if (bodyPl.NumberOfVertices < 2) return false;

                // 取块内局部坐标首末点作为轴线端点
                Point3d localStart = bodyPl.GetPoint3dAt(0);
                // 块内终点
                Point3d localEnd = bodyPl.GetPoint3dAt(bodyPl.NumberOfVertices - 1);

                // 通过块变换矩阵转到世界坐标（非常关键）
                start = br.BlockTransform * localStart;
                // 终点转世界坐标
                end = br.BlockTransform * localEnd;

                // 若首末点重合（闭合或异常），尝试用几何包围盒对角兜底
                if (start.DistanceTo(end) <= 1e-8)
                {
                    // 尝试读取几何包围盒
                    try
                    {
                        var ext = br.GeometricExtents;
                        // 用对角点构造兜底轴线
                        start = ext.MinPoint;
                        end = ext.MaxPoint;
                    }
                    catch
                    {
                        // 包围盒失败则保持原值，后面会统一失败返回
                    }
                }

                // 再次校验轴线有效性
                if (start.DistanceTo(end) <= 1e-8) return false;

                // 读取块属性（AttributeReference）
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 打开属性引用
                    var ar = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                    // 过滤空对象
                    if (ar == null) continue;
                    // 读取并清洗 Tag
                    string tag = (ar.Tag ?? string.Empty).Trim();
                    // 空 Tag 跳过
                    if (string.IsNullOrWhiteSpace(tag)) continue;
                    // 读取文本值
                    string val = ar.TextString ?? string.Empty;

                    // 双写键，便于后续兼容读取
                    meta[tag] = val;
                    // 带前缀键
                    meta["ATT:" + tag] = val;
                }

                // 读取扩展字典 XRecord（如果存在）
                if (br.ExtensionDictionary != ObjectId.Null)
                {
                    // 打开扩展字典
                    var dict = tr.GetObject(br.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                    // 字典有效再处理
                    if (dict != null)
                    {
                        // 遍历字典项
                        foreach (DBDictionaryEntry entry in dict)
                        {
                            // 读取键名
                            string key = (entry.Key ?? string.Empty).Trim();
                            // 空键跳过
                            if (string.IsNullOrWhiteSpace(key)) continue;

                            // 读取 XRecord
                            var xrec = tr.GetObject(entry.Value, OpenMode.ForRead) as Xrecord;
                            // 无数据跳过
                            if (xrec?.Data == null) continue;

                            // 取首个值作为字符串（与现有项目逻辑一致）
                            var arr = xrec.Data.AsArray();
                            // 空数组跳过
                            if (arr == null || arr.Length == 0) continue;
                            // 转字符串
                            string xv = arr[0].Value?.ToString() ?? string.Empty;

                            // 写入字典（不覆盖已有同名属性）
                            if (!meta.ContainsKey(key)) meta[key] = xv;
                            // 同时写前缀键
                            meta["XREC:" + key] = xv;
                        }
                    }
                }

                // 角色兜底（你后续可按实际字段名改这里）
                if (!meta.ContainsKey("PipeRole"))
                {
                    // 按块名猜测
                    if (meta["BlockName"].Contains("进口")) meta["PipeRole"] = "Inlet";
                    else if (meta["BlockName"].Contains("出口")) meta["PipeRole"] = "Outlet";
                }

                // 成功读取，返回 true
                return true;
            }
            catch (Exception ex)
            {
                // 记录异常并安全返回 false，避免中断主命令
                LogManager.Instance.LogInfo($"\nTryGetPipeAxisAndMeta 执行异常: {ex.Message}");
                return false;
            }
        }

        // 搜索候选管线（包围盒粗筛） // 中文注释
        private static List<ObjectId> FindNearbyPipes(DBTrans tr, Point3d start, Point3d end, ObjectId selfId)
        {
            // 创建返回列表 // 中文注释
            var result = new List<ObjectId>();
            // 判空保护 // 中文注释
            if (tr == null) return result;

            // 搜索缓冲（图纸单位，首版可调） // 中文注释
            const double searchBuffer = 200.0;

            // 计算当前管线轴线包围盒（加缓冲） // 中文注释
            double minX = Math.Min(start.X, end.X) - searchBuffer;
            double minY = Math.Min(start.Y, end.Y) - searchBuffer;
            double minZ = Math.Min(start.Z, end.Z) - searchBuffer;
            double maxX = Math.Max(start.X, end.X) + searchBuffer;
            double maxY = Math.Max(start.Y, end.Y) + searchBuffer;
            double maxZ = Math.Max(start.Z, end.Z) + searchBuffer;

            // 本地函数：判断两个 AABB 是否相交 // 中文注释
            bool IntersectsAabb(double aMinX, double aMinY, double aMinZ, double aMaxX, double aMaxY, double aMaxZ,
                                double bMinX, double bMinY, double bMinZ, double bMaxX, double bMaxY, double bMaxZ)
            {
                if (aMaxX < bMinX || aMinX > bMaxX) return false;
                if (aMaxY < bMinY || aMinY > bMaxY) return false;
                if (aMaxZ < bMinZ || aMinZ > bMaxZ) return false;
                return true;
            }

            try
            {
                // 遍历当前空间实体 // 中文注释
                foreach (ObjectId id in tr.CurrentSpace)
                {
                    // 跳过空ID // 中文注释
                    if (id == ObjectId.Null) continue;
                    // 跳过自身 // 中文注释
                    if (id == selfId) continue;

                    // 先尝试读取为实体 // 中文注释
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    // 非实体跳过 // 中文注释
                    if (ent == null) continue;
                    // 已删除跳过 // 中文注释
                    if (ent.IsErased) continue;
                    // 只处理块参照（你的管线对象模型） // 中文注释
                    if (!(ent is BlockReference)) continue;

                    // 尝试获取候选轴线 // 中文注释
                    if (!TryGetPipeAxisAndMeta(tr, id, out Point3d bStart, out Point3d bEnd, out Dictionary<string, string> _))
                        continue;

                    // 计算候选轴线包围盒 // 中文注释
                    double bMinX = Math.Min(bStart.X, bEnd.X);
                    double bMinY = Math.Min(bStart.Y, bEnd.Y);
                    double bMinZ = Math.Min(bStart.Z, bEnd.Z);
                    double bMaxX = Math.Max(bStart.X, bEnd.X);
                    double bMaxY = Math.Max(bStart.Y, bEnd.Y);
                    double bMaxZ = Math.Max(bStart.Z, bEnd.Z);

                    // 粗筛：包围盒相交才入候选 // 中文注释
                    if (!IntersectsAabb(minX, minY, minZ, maxX, maxY, maxZ, bMinX, bMinY, bMinZ, bMaxX, bMaxY, bMaxZ))
                        continue;

                    // 加入候选 // 中文注释
                    result.Add(id);
                }
            }
            catch (Exception ex)
            {
                // 出错时记录日志并返回已有结果 // 中文注释
                LogManager.Instance.LogInfo($"\nFindNearbyPipes 执行异常: {ex.Message}");
            }

            // 返回候选列表 // 中文注释
            return result;
        }

        // 判定“端点接触 + 近似共线 + 业务可并” // 中文注释
        private static bool CanMergeAtEndpoint(Point3d a1, Point3d a2, Point3d b1, Point3d b2, Dictionary<string, string> ma, Dictionary<string, string> mb)
        {
            // ========================= 基础容差参数（第一批先写死，后续可迁移到 AutoCadHelper） ========================= // 中文注释
            double endpointTol = 20; // 端点接触容差：两端点距离小于等于该值视为“接触” // 中文注释
            double angleTolDeg = 1.0; // 共线角度容差（度）：允许 1 度以内偏差 // 中文注释
            double angleTolRad = angleTolDeg * Math.PI / 180.0; // 把角度容差从度转换为弧度 // 中文注释
            double cosTol = Math.Cos(angleTolRad); // 角度阈值转为点积阈值，便于快速判断方向夹角 // 中文注释

            // ========================= 防御性处理：空字典兜底，避免空引用 ========================= // 中文注释
            ma ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // 若 ma 为空则补一个空字典 // 中文注释
            mb ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // 若 mb 为空则补一个空字典 // 中文注释

            // ========================= 第一步：轴线有效性判断（避免零长度线段参与计算） ========================= // 中文注释
            if (a1.DistanceTo(a2) <= 1e-8) return false; // A 管线起终点重合，判定无效不可并 // 中文注释
            if (b1.DistanceTo(b2) <= 1e-8) return false; // B 管线起终点重合，判定无效不可并 // 中文注释

            // ========================= 第二步：端点接触判断（至少有一组端点需要贴合） ========================= // 中文注释
            double d11 = a1.DistanceTo(b1); // A起点 到 B起点 的距离 // 中文注释
            double d12 = a1.DistanceTo(b2); // A起点 到 B终点 的距离 // 中文注释
            double d21 = a2.DistanceTo(b1); // A终点 到 B起点 的距离 // 中文注释
            double d22 = a2.DistanceTo(b2); // A终点 到 B终点 的距离 // 中文注释
            double minEndpointDist = Math.Min(Math.Min(d11, d12), Math.Min(d21, d22)); // 取四组端点距离中的最小值 // 中文注释
            if (minEndpointDist > endpointTol) return false; // 最小端点距离超容差，说明不是端点相接，不能并线 // 中文注释

            // ========================= 第三步：近似共线判断（方向要平行或反向平行） ========================= // 中文注释
            Vector3d va = (a2 - a1).GetNormal(); // 计算 A 管线方向单位向量 // 中文注释
            Vector3d vb = (b2 - b1).GetNormal(); // 计算 B 管线方向单位向量 // 中文注释
            double absDot = Math.Abs(va.DotProduct(vb)); // 方向点积绝对值，越接近 1 越平行 // 中文注释
            if (absDot < cosTol) return false; // 小于阈值说明夹角过大，不满足近似共线 // 中文注释

            // ========================= 第四步：业务规则兼容判断（角色/图层/关键属性） ========================= // 中文注释

            // 本地函数：按“候选键名列表”读取第一个非空值（兼容 ATT: 前缀和无前缀） // 中文注释
            string GetFirstMeta(Dictionary<string, string> map, params string[] keys)
            {
                // 遍历所有候选键名 // 中文注释
                foreach (string key in keys)
                {
                    // 字典中存在该键才继续 // 中文注释
                    if (!map.TryGetValue(key, out string v)) continue;
                    // 去掉首尾空白后再判空 // 中文注释
                    string s = (v ?? string.Empty).Trim();
                    // 非空即返回 // 中文注释
                    if (!string.IsNullOrWhiteSpace(s)) return s;
                }
                // 所有候选都没值则返回空串 // 中文注释
                return string.Empty;
            }

            // 本地函数：如果双方都给了值则必须相等；只要有一方没给值就放行（第一批策略） // 中文注释
            bool IsCompatible(string v1, string v2)
            {
                // 一方缺失信息时先放行，避免首版过严导致无法并线 // 中文注释
                if (string.IsNullOrWhiteSpace(v1) || string.IsNullOrWhiteSpace(v2)) return true;
                // 双方都有值时必须完全一致（忽略大小写） // 中文注释
                return string.Equals(v1.Trim(), v2.Trim(), StringComparison.OrdinalIgnoreCase);
            }

            // 读取角色（进口/出口） // 中文注释
            string roleA = GetFirstMeta(ma, "PipeRole", "ATT:PipeRole", "角色", "ATT:角色"); // 从 A 元数据读取角色 // 中文注释
            string roleB = GetFirstMeta(mb, "PipeRole", "ATT:PipeRole", "角色", "ATT:角色"); // 从 B 元数据读取角色 // 中文注释
            if (!IsCompatible(roleA, roleB)) return false; // 双方角色同时存在且不一致，则不允许并线 // 中文注释

            // 读取图层（可选约束：双方都有图层时要求一致） // 中文注释
            string layerA = GetFirstMeta(ma, "Layer"); // 读取 A 图层 // 中文注释
            string layerB = GetFirstMeta(mb, "Layer"); // 读取 B 图层 // 中文注释
            if (!IsCompatible(layerA, layerB)) return false; // 双方图层同时存在且不一致，则不允许并线 // 中文注释

            // 读取关键业务属性：系统 // 中文注释
            string sysA = GetFirstMeta(ma, "PipeSystem", "ATT:PipeSystem", "系统", "ATT:系统"); // 读取 A 系统字段 // 中文注释
            string sysB = GetFirstMeta(mb, "PipeSystem", "ATT:PipeSystem", "系统", "ATT:系统"); // 读取 B 系统字段 // 中文注释
            if (!IsCompatible(sysA, sysB)) return false; // 系统冲突则不允许并线 // 中文注释

            // 读取关键业务属性：规格 // 中文注释
            string specA = GetFirstMeta(ma, "Spec", "ATT:Spec", "规格", "ATT:规格"); // 读取 A 规格字段 // 中文注释
            string specB = GetFirstMeta(mb, "Spec", "ATT:Spec", "规格", "ATT:规格"); // 读取 B 规格字段 // 中文注释
            if (!IsCompatible(specA, specB)) return false; // 规格冲突则不允许并线 // 中文注释

            // 读取关键业务属性：管径 // 中文注释
            string dnA = GetFirstMeta(ma, "DN", "ATT:DN", "管径", "ATT:管径"); // 读取 A 管径字段 // 中文注释
            string dnB = GetFirstMeta(mb, "DN", "ATT:DN", "管径", "ATT:管径"); // 读取 B 管径字段 // 中文注释
            if (!IsCompatible(dnA, dnB)) return false; // 管径冲突则不允许并线 // 中文注释

            // ========================= 所有检查通过，允许端点并线 ========================= // 中文注释
            return true; // 返回 true 表示满足“端点接触 + 近似共线 + 业务可并” // 中文注释
        }

        // 执行并线：重算一条块的参数（基于 keep 块定义内主 Polyline 重写首末点），删除另一条 // 中文注释
        private static bool TryMergePipes(DBTrans tr, ObjectId keepId, ObjectId eraseId, Point3d mergedStart, Point3d mergedEnd)
        {
            // 判空保护 // 中文注释
            if (tr == null) return false;
            // ID 判空保护 // 中文注释
            if (keepId == ObjectId.Null || eraseId == ObjectId.Null) return false;
            // 自己和自己合并无意义 // 中文注释
            if (keepId == eraseId) return false;

            try
            {
                // 打开保留对象（可写） // 中文注释
                var keepBr = tr.GetObject(keepId, OpenMode.ForWrite) as BlockReference;
                // 打开删除对象（可写） // 中文注释
                var eraseEnt = tr.GetObject(eraseId, OpenMode.ForWrite) as Entity;

                // 基础有效性检查 // 中文注释
                if (keepBr == null || eraseEnt == null) return false;
                if (keepBr.IsErased || eraseEnt.IsErased) return false;

                // 打开 keep 的块定义（可写） // 中文注释
                var keepBtr = tr.GetObject(keepBr.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                if (keepBtr == null) return false;

                // 保护：若该块定义被多个参照共享，直接返回 false，避免影响其它管线 // 中文注释
                try
                {
                    var refs = keepBtr.GetBlockReferenceIds(true, false);
                    if (refs != null && refs.Count > 1)
                    {
                        LogManager.Instance.LogInfo("\nTryMergePipes: 块定义被多个参照共享，首版跳过并线以避免连带修改。");
                        return false;
                    }
                }
                catch
                {
                    // 忽略引用计数异常，继续尝试 // 中文注释
                }

                // 在 keep 块定义中找“主体 Polyline”（长度最大） // 中文注释
                Polyline bodyPl = null;
                double maxLen = -1.0;

                foreach (ObjectId entId in keepBtr)
                {
                    var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                    if (ent is Polyline pl)
                    {
                        double len = 0.0;
                        try { len = pl.Length; } catch { len = 0.0; }
                        if (len > maxLen)
                        {
                            maxLen = len;
                            bodyPl = pl;
                        }
                    }
                }

                // 没找到主体线则失败 // 中文注释
                if (bodyPl == null) return false;

                // 把世界坐标的合并端点转换到 keep 块局部坐标 // 中文注释
                Matrix3d inv = keepBr.BlockTransform.Inverse();
                Point3d localStart = inv * mergedStart;
                Point3d localEnd = inv * mergedEnd;

                // 打开主体 Polyline 为可写 // 中文注释
                var bodyPlWrite = tr.GetObject(bodyPl.ObjectId, OpenMode.ForWrite) as Polyline;
                if (bodyPlWrite == null) return false;

                // 清空原有顶点 // 中文注释
                while (bodyPlWrite.NumberOfVertices > 0)
                {
                    bodyPlWrite.RemoveVertexAt(bodyPlWrite.NumberOfVertices - 1);
                }

                // 使用两点重建为单段轴线 // 中文注释
                bodyPlWrite.AddVertexAt(0, new Point2d(localStart.X, localStart.Y), 0.0, 0.0, 0.0);
                bodyPlWrite.AddVertexAt(1, new Point2d(localEnd.X, localEnd.Y), 0.0, 0.0, 0.0);
                bodyPlWrite.Closed = false;

                // 删除被合并对象 // 中文注释
                eraseEnt.Erase();

                // 成功 // 中文注释
                return true;
            }
            catch (Exception ex)
            {
                // 记录异常 // 中文注释
                LogManager.Instance.LogInfo($"\nTryMergePipes 执行异常: {ex.Message}");
                return false;
            }
        }

        // 合并属性：将来源字典写入目标块（同名优先，空值不覆盖非空值） // 中文注释
        private static void MergePipeProperties(DBTrans tr, ObjectId targetId, Dictionary<string, string> fromMap)
        {
            // 判空保护 // 中文注释
            if (tr == null) return;
            if (targetId == ObjectId.Null) return;
            if (fromMap == null || fromMap.Count == 0) return;

            try
            {
                // 打开目标实体 // 中文注释
                var ent = tr.GetObject(targetId, OpenMode.ForWrite) as Entity;
                if (ent == null || ent.IsErased) return;

                // 优先写块属性（AttributeReference） // 中文注释
                if (ent is BlockReference br)
                {
                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        var ar = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                        if (ar == null) continue;

                        string tag = (ar.Tag ?? string.Empty).Trim();
                        if (string.IsNullOrWhiteSpace(tag)) continue;

                        if (!fromMap.TryGetValue(tag, out string srcVal)) continue;
                        srcVal = srcVal ?? string.Empty;

                        // 策略：目标已有值且非空时不覆盖；目标为空时才填充 // 中文注释
                        string cur = ar.TextString ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(cur) && !string.IsNullOrWhiteSpace(srcVal))
                        {
                            ar.TextString = srcVal;
                        }
                    }
                }

                // 再写扩展字典同名键（仅更新已有键，避免首版引入新键结构） // 中文注释
                if (ent.ExtensionDictionary != ObjectId.Null)
                {
                    var dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;
                    if (dict != null)
                    {
                        foreach (DBDictionaryEntry entry in dict)
                        {
                            string key = (entry.Key ?? string.Empty).Trim();
                            if (string.IsNullOrWhiteSpace(key)) continue;

                            if (!fromMap.TryGetValue(key, out string srcVal)) continue;
                            srcVal = srcVal ?? string.Empty;
                            if (string.IsNullOrWhiteSpace(srcVal)) continue;

                            var xrec = tr.GetObject(entry.Value, OpenMode.ForWrite) as Xrecord;
                            if (xrec == null) continue;

                            // 读取当前值（首项） // 中文注释
                            string cur = string.Empty;
                            try
                            {
                                var arr = xrec.Data?.AsArray();
                                if (arr != null && arr.Length > 0) cur = arr[0].Value?.ToString() ?? string.Empty;
                            }
                            catch { }

                            // 仅当前为空时写入 // 中文注释
                            if (string.IsNullOrWhiteSpace(cur))
                            {
                                xrec.Data = new ResultBuffer(new TypedValue((int)DxfCode.Text, srcVal));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录异常，不中断主流程 // 中文注释
                LogManager.Instance.LogInfo($"\nMergePipeProperties 执行异常: {ex.Message}");
            }
        }
    } 
} 
