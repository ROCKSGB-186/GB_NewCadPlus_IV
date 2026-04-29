
using Autodesk.AutoCAD.DatabaseServices;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Helpers;
using Mysqlx.Crud;
using OfficeOpenXml;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using AttributeCollection = Autodesk.AutoCAD.DatabaseServices.AttributeCollection;
using DataTable = System.Data.DataTable;



namespace GB_NewCadPlus_IV.UniFiedStandards
{
    /// <summary>
    /// 设备属性块信息类（用于存储CAD块中的设备信息）
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

        /// <summary>
        /// 属性字典（中文属性名-值）
        /// </summary>
        public Dictionary<string, string> Attributes { get; set; }

        /// <summary>
        /// 英文属性名对照（中文属性名-英文属性名）
        /// </summary>
        public Dictionary<string, string> EnglishNames { get; set; }

        /// <summary>
        /// 相同设备的数量统计
        /// </summary>
        public int Count { get; set; }

        /// <summary>
        /// 构造函数初始化字典和默认值
        /// </summary>
        public DeviceInfo()
        {
            Attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            EnglishNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Count = 1;  // 默认数量为1
        }
    }

    /// <summary>
    /// 统一表生成器
    /// </summary>
    public class UnifiedTableGenerator
    {

        //实例化WPF窗体
        WpfMainWindow wpfMainWindow = new WpfMainWindow();


        /// <summary>
        /// 主命令：生成设备材料表（按类型拆分为多个表）arrowEntities
        /// </summary>
        /// <summary>
        /// 主命令：生成设备材料表（按类型拆分为多个表）
        /// 修复要点：
        /// - 不再在此处重复调用 GetSelection（避免需要第二次选择的问题）
        /// - 直接使用 SelectAndAnalyzeBlocks 内的选择逻辑
        /// - 每生成并插入一个表后立即刷新界面，确保用户能看到刚插入的表
        /// </summary>
        [CommandMethod("GenerateDeviceTable")]
        public void GenerateDeviceTable()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            try
            {
                // 使用 DynamicBlockOperations 的新签名（返回设备列表 + 所选 ObjectId[]）
                var (devices, selIds) = DynamicBlockOperations.SelectAndAnalyzeBlocks(ed, doc.Database);
                if (devices == null || devices.Count == 0)
                {
                    ed.WriteMessage("\n未找到可用的设备信息。");
                    return;
                }

                // 基于所选图元推断比例分母（优先使用所选图元的布局/视口）
                //double scaleDenom = GB_NewCadPlus_IV.Helpers.AutoCadHelper.GetScaleDenominatorForSelection(doc.Database, selIds, roundToCommon: false);
                double scaleDenom = AutoCadHelper.GetScale();
                // 按 Type 分组并为每组生成独立表
                var groups = devices.GroupBy(e => string.IsNullOrWhiteSpace(e.Type) ? "设备" : e.Type);
                foreach (var g in groups)
                {
                    List<DeviceInfo>? list = g.ToList();
                    // 把推断到的比例传入 CreateDeviceTable（新签名）
                    CreateDeviceTable(doc.Database, list, scaleDenom);
                    //CreateDeviceTable(doc.Database, list, 100);
                    try { ed.Regen(); Application.UpdateScreen(); } catch { }
                    ed.WriteMessage($"\n已为类型 '{g.Key}' 生成表，包含 {list.Count} 条汇总项（使用比例分母 {scaleDenom}）。");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n生成设备表失败: {ex.Message}");
            }
        }

        /// <summary>
        /// “生成(工艺)管道表”：GeneratePipeTableFromSelection
        /// - 交互式选择实体
        /// - 提取属性（优先使用属性中包含“长度”的字段作为长度）
        /// - 按属性组合或几何尺寸分组并累加长度
        /// - 最终以表格形式（沿用本类表格样式）生成一张或多张“管道”表
        /// - 新增功能：自动调整每列列宽以适应最长文字
        /// </summary>
        [CommandMethod("GeneratePipeTableFromSelection")]
        public void GeneratePipeTableFromSelection()
        {
            // 获取当前活动文档，若为空直接返回，避免空引用异常
            var doc = Application.DocumentManager.MdiActiveDocument;
            // 无活动文档时不继续执行
            if (doc == null) return;

            // 在统一的文档事务中执行，保持与现有代码风格一致
            AutoCadHelper.ExecuteInDocumentTransaction((d, tr) =>
            {
                // 获取编辑器对象，用于交互提示和日志输出
                var ed = d.Editor;
                try
                {
                    // 提示用户选择要统计的管道图元
                    ed.WriteMessage("\n开始生成管道表：请选择要统计的管道图元（完成选择回车）.");
                    // 设置选择提示文本
                    var pso = new PromptSelectionOptions { MessageForAdding = "\n请选择要统计的管道图元：" };
                    // 执行选择
                    var psr = ed.GetSelection(pso);
                    // 若未成功选择则直接结束
                    if (psr.Status != PromptStatus.OK || psr.Value == null)
                    {
                        ed.WriteMessage("\n未选择实体或已取消。");
                        return;
                    }

                    // 取到所选对象ID数组
                    var selIds = psr.Value.GetObjectIds();
                    // 空选择保护
                    if (selIds == null || selIds.Length == 0)
                    {
                        ed.WriteMessage("\n未选择任何实体。");
                        return;
                    }

                    // 统一比例来源——强制使用 AutoCadHelper.GetScale(...)（禁用缓存，避免刚切视口时读到旧值）
                    double rawScale = AutoCadHelper.GetScale(useCache: false);
                    // 把 GetScale 返回值统一归一化为“比例分母”（例如 0.01->100，100->100，1->1）
                    double scaleDenom = TextFontsStyleHelper.DetermineScaleDenominator(rawScale, null, roundToCommon: false);
                    // 兜底保护，防止非法值导致表格高度异常
                    if (double.IsNaN(scaleDenom) || double.IsInfinity(scaleDenom) || scaleDenom <= 0.0)
                    {
                        // 非法时回退到 1:1
                        scaleDenom = 1.0;
                    }

                    // 输出日志，便于你核对“GetScale原值”和“最终分母值”
                    ed.WriteMessage($"\nGetScale原值: {rawScale}，归一化比例分母: {scaleDenom}（将用于表格文字高度与行高计算）。");

                    // 图形单位到米的换算常量（按你现有逻辑保持 1000）
                    const double unitToMeters = 1000.0;
                    // 管段号候选键
                    string[] pipeNoKeys = new[] { "管段号", "管段编号", "Pipeline No", "Pipeline", "Pipe No", "No" };
                    // 起点候选键
                    string[] startKeys = new[] { "起点", "始点", "From" };
                    // 终点候选键
                    string[] endKeys = new[] { "终点", "止点", "To" };

                    // 局部函数——按候选键优先、再按模糊键名匹配，提取首个有效属性值
                    string FindFirstAttrValueLocal(Dictionary<string, string>? attrs, string[] candidates)
                    {
                        // 空字典直接返回空字符串
                        if (attrs == null) return string.Empty;
                        // 先走精确键匹配
                        foreach (var c in candidates)
                        {
                            // 命中且有值则立即返回
                            if (attrs.TryGetValue(c, out var v) && !string.IsNullOrWhiteSpace(v)) return v;
                        }
                        // 再走包含匹配作为兜底
                        foreach (var kv in attrs)
                        {
                            // 空键跳过
                            if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                            foreach (var c in candidates)
                            {
                                // 键名包含候选词且值不空时返回
                                if (kv.Key.IndexOf(c, StringComparison.OrdinalIgnoreCase) >= 0 && !string.IsNullOrWhiteSpace(kv.Value))
                                    return kv.Value;
                            }
                        }
                        // 未命中返回空
                        return string.Empty;
                    }

                    // 逐条管道记录列表（每个选中实体对应一行）
                    var perPipeList = new List<DeviceInfo>();
                    // 顺序编号，用于默认 Name
                    int seqIndex = 0;

                    // 遍历所选实体，提取长度/起点/终点/管段号等信息
                    foreach (var id in selIds)
                    {
                        // 序号递增
                        seqIndex++;
                        try
                        {
                            // 读取实体对象
                            var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            // 无效实体跳过
                            if (ent == null) continue;

                            // 读取实体属性字典（块属性、扩展字典、XData）
                            var attrMap = GetEntityAttributeMap(tr, ent);

                            // 长度（米）初始为 NaN，表示尚未取到
                            double length_m = double.NaN;
                            // 先从属性里找“长度”字段（优先）
                            if (attrMap != null)
                            {
                                foreach (var k in attrMap.Keys)
                                {
                                    if (!string.IsNullOrWhiteSpace(k) && k.IndexOf("长度", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        var parsed = ParseLengthValueFromAttribute(attrMap[k]);
                                        if (!double.IsNaN(parsed) && parsed > 0.0)
                                        {
                                            length_m = parsed;
                                            break;
                                        }
                                    }
                                }
                            }

                            // 若属性没取到长度，再按实体几何类型回退计算
                            if (double.IsNaN(length_m))
                            {
                                if (ent is Line lineEnt)
                                {
                                    length_m = lineEnt.Length / unitToMeters;
                                }
                                else if (ent is Polyline plEnt)
                                {
                                    length_m = plEnt.Length / unitToMeters;
                                }
                                else if (ent is BlockReference brEnt)
                                {
                                    try
                                    {
                                        double l = DynamicBlockOperations.GetLength(brEnt);
                                        if (!double.IsNaN(l) && l > 0.0) length_m = l;
                                    }
                                    catch { }
                                }

                                if (double.IsNaN(length_m))
                                {
                                    try
                                    {
                                        var ext = ent.GeometricExtents;
                                        double sizeX = Math.Abs(ext.MaxPoint.X - ext.MinPoint.X);
                                        length_m = sizeX / unitToMeters;
                                    }
                                    catch
                                    {
                                        length_m = 0.0;
                                    }
                                }
                            }

                            // 优先从属性取起点/终点
                            string startStr = FindFirstAttrValueLocal(attrMap, startKeys);
                            string endStr = FindFirstAttrValueLocal(attrMap, endKeys);

                            // 属性缺失时，从几何信息回退
                            if (string.IsNullOrWhiteSpace(startStr) || string.IsNullOrWhiteSpace(endStr))
                            {
                                if (ent is Line lineEnt2)
                                {
                                    if (string.IsNullOrWhiteSpace(startStr))
                                        startStr = $"X={lineEnt2.StartPoint.X:F3},Y={lineEnt2.StartPoint.Y:F3}";
                                    if (string.IsNullOrWhiteSpace(endStr))
                                        endStr = $"X={lineEnt2.EndPoint.X:F3},Y={lineEnt2.EndPoint.Y:F3}";
                                }
                                else if (ent is Polyline plEnt2)
                                {
                                    try
                                    {
                                        var p0 = plEnt2.GetPoint3dAt(0);
                                        var pN = plEnt2.GetPoint3dAt(plEnt2.NumberOfVertices - 1);
                                        if (string.IsNullOrWhiteSpace(startStr))
                                            startStr = $"X={p0.X:F3},Y={p0.Y:F3}";
                                        if (string.IsNullOrWhiteSpace(endStr))
                                            endStr = $"X={pN.X:F3},Y={pN.Y:F3}";
                                    }
                                    catch { }
                                }
                                else if (ent is BlockReference brEnt2)
                                {
                                    try
                                    {
                                        var (s, e) = DynamicBlockOperations.GetEndPoints(brEnt2);
                                        if (string.IsNullOrWhiteSpace(startStr))
                                            startStr = $"X={s.X:F3},Y={s.Y:F3}";
                                        if (string.IsNullOrWhiteSpace(endStr))
                                            endStr = $"X={e.X:F3},Y={e.Y:F3}";
                                    }
                                    catch { }
                                }

                                if (string.IsNullOrWhiteSpace(startStr) || string.IsNullOrWhiteSpace(endStr))
                                {
                                    try
                                    {
                                        var ext = ent.GeometricExtents;
                                        if (string.IsNullOrWhiteSpace(startStr))
                                            startStr = $"X={ext.MinPoint.X:F3},Y={ext.MinPoint.Y:F3}";
                                        if (string.IsNullOrWhiteSpace(endStr))
                                            endStr = $"X={ext.MaxPoint.X:F3},Y={ext.MaxPoint.Y:F3}";
                                    }
                                    catch
                                    {
                                        if (string.IsNullOrWhiteSpace(startStr)) startStr = "N/A";
                                        if (string.IsNullOrWhiteSpace(endStr)) endStr = "N/A";
                                    }
                                }
                            }

                            // 读取管段号
                            string pipeNo = FindFirstAttrValueLocal(attrMap, pipeNoKeys);
                            // 若无管段号且是块参照，则尝试从块名中提取数字
                            if (string.IsNullOrWhiteSpace(pipeNo))
                            {
                                if (ent is BlockReference brEnt3)
                                {
                                    try
                                    {
                                        var btr = tr.GetObject(brEnt3.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                        if (btr != null)
                                        {
                                            var nm = btr.Name ?? string.Empty;
                                            var m = Regex.Match(nm, @"\d+");
                                            if (m.Success) pipeNo = m.Value;
                                        }
                                    }
                                    catch { }
                                }
                            }

                            // 构建单条记录对象
                            var info = new DeviceInfo
                            {
                                Name = $"PIPE_{seqIndex}",
                                Type = "管道"
                            };

                            // 拷贝原属性到记录中
                            if (attrMap != null)
                            {
                                foreach (var kv in attrMap) info.Attributes[kv.Key] = kv.Value;
                            }

                            // 写入标准化字段
                            if (!string.IsNullOrWhiteSpace(pipeNo)) info.Attributes["管段号"] = pipeNo;
                            info.Attributes["起点"] = startStr;
                            info.Attributes["终点"] = endStr;
                            info.Attributes["长度(m)"] = length_m.ToString("F3");
                            info.Attributes["累计长度(m)"] = length_m.ToString("F3");
                            if (info.Attributes.TryGetValue("介质", out var medVal) && !info.Attributes.ContainsKey("介质名称"))
                            {
                                info.Attributes["介质名称"] = medVal;
                            }

                            // 加入结果列表
                            perPipeList.Add(info);
                        }
                        catch (System.Exception exEnt)
                        {
                            // 单实体失败不影响整体
                            ed.WriteMessage($"\n处理实体 {id} 时出错: {exEnt.Message}");
                        }
                    }

                    // 无有效记录则结束
                    if (perPipeList.Count == 0)
                    {
                        ed.WriteMessage("\n未生成任何管道记录。");
                        return;
                    }

                    // 局部函数——从管段号中提取数字用于排序
                    int ExtractPipeNoNumber(string s)
                    {
                        if (string.IsNullOrWhiteSpace(s)) return int.MaxValue;
                        var m = Regex.Match(s, @"\d+");
                        if (m.Success && int.TryParse(m.Value, out var v)) return v;
                        return int.MaxValue;
                    }

                    // 按管段号数字排序，保证表格顺序稳定
                    var finalList = perPipeList
                        .Select((e, idx) => new { Item = e, Orig = idx })
                        .OrderBy(x =>
                        {
                            if (x.Item.Attributes.TryGetValue("管段号", out var pn) && !string.IsNullOrWhiteSpace(pn))
                                return (ExtractPipeNoNumber(pn), 0, x.Orig);
                            return (int.MaxValue, 1, x.Orig);
                        })
                        .Select(x => x.Item)
                        .ToList();

                    // ================== 关键修改区域 ==================

                    // 创建管道明细表，并传入统一后的比例分母
                    // 注意：CreateDeviceTableWithType 是 void 方法，内部已处理列宽调整，因此不需要接收返回值
                    CreateDeviceTableWithType(d.Database, finalList, "管道明细", scaleDenom);

                    // 删除原本依赖 tableId 的代码块，因为方法内部已经完成了所有操作
                    // if (!tableId.IsNull)
                    // {
                    //     AutoFitTableColumns(tr, tableId, ed);
                    // }

                    // ================== 结束关键修改区域 ==================

                    // 完成提示
                    ed.WriteMessage($"\n管道表已生成，共 {finalList.Count} 条记录（每选中一条管道一行，使用比例分母 {scaleDenom}）。");
                }
                catch (System.Exception ex)
                {
                    // 整体异常捕获，防止命令中断
                    d.Editor.WriteMessage($"\n生成管道表时发生错误: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 自动调整表格列宽以适应内容
        /// </summary>
        /// <param name="tr">当前事务</param>
        /// <param name="tableId">表格对象的 ObjectId</param>
        /// <param name="ed">编辑器对象，用于获取文字 extents</param>
        private void AutoFitTableColumns(Transaction tr, ObjectId tableId, Editor ed)
        {
            try
            {
                // 以写模式打开表格对象
                var table = tr.GetObject(tableId, OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.Table;
                if (table == null) return;

                // 获取表格的行数和列数
                int numRows = table.Rows.Count;
                int numCols = table.Columns.Count;

                // 遍历每一列
                for (int col = 0; col < numCols; col++)
                {
                    double maxWidth = 0.0;
                    // 获取当前列的文字高度（假设所有单元格文字高度一致，取第一个非空单元格的高度）
                    double textHeight = 2.5; // 默认高度，可根据实际情况调整

                    // 遍历该列的所有行
                    for (int row = 0; row < numRows; row++)
                    {
                        // 获取单元格的文本内容
                        string cellText = table.Cells[row, col].TextString;
                        if (string.IsNullOrWhiteSpace(cellText)) continue;

                        // 获取单元格的文字高度
                        if (table.Cells[row, col].TextHeight > 0)
                        {
                            textHeight = Convert.ToDouble(table.Cells[row, col].TextHeight);
                        }

                        // 计算文本的显示宽度
                        // 简单估算：汉字宽度约为高度的 1.0-1.2 倍，英文约为 0.6-0.8 倍
                        // 更准确的方法是使用 Geometry.TextBounds，但需要知道字体和样式
                        // 这里使用一种简化的估算方法：每个字符平均宽度为 textHeight * 0.8
                        // 对于中文，可能需要更宽，这里乘以 1.0 作为系数
                        double estimatedWidth = cellText.Length * textHeight * 0.8;

                        // 考虑一些常用汉字的宽度，适当增加系数
                        // 如果包含中文，系数可以更大一些，比如 1.0
                        bool hasChinese = Regex.IsMatch(cellText, @"[\u4e00-\u9fff]");
                        if (hasChinese)
                        {
                            estimatedWidth = cellText.Length * textHeight * 1.0;
                        }

                        // 更新最大宽度
                        if (estimatedWidth > maxWidth)
                        {
                            maxWidth = estimatedWidth;
                        }
                    }

                    // 设置列宽，增加一定的边距（Padding），例如左右各加 0.5 倍文字高度
                    if (maxWidth > 0)
                    {
                        double padding = textHeight * 1.0; // 左右边距总和
                        table.Columns[col].Width = maxWidth + padding;
                    }
                    else
                    {
                        // 如果列中没有内容，设置一个最小宽度
                        table.Columns[col].Width = textHeight * 5;
                    }
                }

                // 强制更新表格显示
                table.GenerateLayout();
            }
            catch (Exception ex)
            {
                // 记录异常，但不中断程序
                ed.WriteMessage($"\n自动调整列宽时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 【优化版】高级自动调整表格列宽，考虑合并单元格和实际文本宽度
        /// </summary>
        /// <param name="table">要调整的表格对象</param>
        private void AutoFitTableColumnsAdvanced(Table table)
        {
            if (table == null) return;

            int numRows = table.Rows.Count;
            int numCols = table.Columns.Count;

            if (numRows == 0 || numCols == 0) return;

            // 获取典型文字高度：优先使用数据行（通常从第4行开始，索引3）的高度，如果没有则用第0行
            double textHeight = 2.5;
            try
            {
                // 尝试获取数据行第一格的高度，通常数据行高度更具代表性
                int sampleRow = numRows > 3 ? 3 : 0;
                var h = table.Cells[sampleRow, 0].TextHeight;
                if (h > 0) textHeight = Convert.ToDouble(h);
            }
            catch { }

            // 遍历每一列
            for (int col = 0; col < numCols; col++)
            {
                double maxWidthInCol = 0.0;

                // 遍历该列的每一行
                for (int row = 0; row < numRows; row++)
                {
                    var cell = table.Cells[row, col];

                    // 1. 检查当前单元格是否被合并
                    bool isMerged = Convert.ToBoolean(cell.IsMerged);
                    bool shouldSkip = false;

                    if (isMerged)
                    {
                        // 关键逻辑：如果是合并单元格，我们只计算“左上角”那个单元格的宽度贡献
                        bool isTopLeft = true;

                        // 向左检查：如果左边的单元格也是合并状态，则当前格不是最左
                        if (col > 0 && Convert.ToBoolean(table.Cells[row, col - 1].IsMerged))
                        {
                            isTopLeft = false;
                        }

                        // 向上检查：如果上面的单元格也是合并状态，则当前格不是最上
                        if (row > 0 && Convert.ToBoolean(table.Cells[row - 1, col].IsMerged))
                        {
                            isTopLeft = false;
                        }

                        // 如果不是左上角单元格，跳过计算，避免重复累加宽度
                        if (!isTopLeft)
                        {
                            shouldSkip = true;
                        }
                    }

                    if (shouldSkip) continue;

                    string cellText = cell.TextString;
                    if (string.IsNullOrWhiteSpace(cellText)) continue;

                    // 计算文本估算宽度
                    double estimatedWidth = 0.0;

                    // 更精确的估算：区分中英文
                    int chineseCount = System.Text.RegularExpressions.Regex.Matches(cellText, @"[\u4e00-\u9fff]").Count;
                    int otherCount = cellText.Length - chineseCount;

                    // 中文系数 1.0，英文/数字系数 0.6
                    estimatedWidth = (chineseCount * 1.0 + otherCount * 0.6) * textHeight;

                    // 增加左右边距 (Padding)，例如每侧 0.5 倍字高
                    double padding = textHeight * 2; // 稍微减小边距，让列更紧凑
                    estimatedWidth += padding;

                    if (estimatedWidth > maxWidthInCol)
                    {
                        maxWidthInCol = estimatedWidth;
                    }
                }

                // 设置列宽，设置一个最小宽度以防万一
                // 【修改点】减小最小宽度系数，从 4.0 改为 3.0，让短列更紧凑
                double minWidth = textHeight * 3;
                if (maxWidthInCol < minWidth) maxWidthInCol = minWidth;

                // 设置最大宽度限制，防止某列特别长导致表格过宽
                double maxWidthLimit = textHeight * 60.0; // 稍微增加最大限制，适应长文本
                if (maxWidthInCol > maxWidthLimit) maxWidthInCol = maxWidthLimit;

                table.Columns[col].Width = maxWidthInCol;
            }

            // 强制更新表格布局
            table.GenerateLayout();
        }

        /// <summary>
        /// 提取中文字符
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private string ExtractChineseCharacters(string? input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            var matches = Regex.Matches(input!, @"[\u4e00-\u9fff]+");
            if (matches.Count == 0) return string.Empty;
            return string.Concat(matches.Cast<Match>().Select(m => m.Value)).Trim();
        }

        /// <summary>
        /// 【新增公共方法】仅提取数据，不插入表格
        /// 供 WPF 窗口调用
        /// </summary>
        public List<DeviceInfo> ExtractPipeDataFromSelection(Editor ed, Database db)
        {
            var deviceList = new List<DeviceInfo>();

            // 1. 选择实体
            var psr = ed.GetSelection();
            if (psr.Status != PromptStatus.OK || psr.Value == null) return null;

            var selIds = psr.Value.GetObjectIds();
            if (selIds.Length == 0) return null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 2. 遍历提取 (这里简化了逻辑，你可以复制 GeneratePipeTableFromSelection 中的核心提取代码)
                foreach (var id in selIds)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent is BlockReference br)
                    {
                        var info = new DeviceInfo();
                        info.Name = br.Name;
                        info.Type = "管道"; // 或其他逻辑

                        // 提取属性
                        var attrMap = GetEntityAttributeMap(tr, ent);
                        if (attrMap != null)
                        {
                            foreach (var kv in attrMap)
                            {
                                info.Attributes[kv.Key] = kv.Value;
                            }
                        }

                        // 提取管段号等关键信息 (参考 GeneratePipeTableFromSelection 中的逻辑)
                        if (info.Attributes.TryGetValue("管段号", out var pipeNo))
                        {
                            info.Name = pipeNo; // 用管段号作为 Name 以便匹配
                        }

                        deviceList.Add(info);
                    }
                }
                tr.Commit();
            }
            return deviceList;
        }

        /// <summary>
        /// 从管道标题中提取管道等级，例如 "350-AR-1002-1.0G11" -> "1.0G11"
        /// 优先匹配含小数点的等级格式（如 1.0G11），再做宽松匹配。
        /// </summary>
        private string ExtractPipeClassFromTitle(string? title)
        {
            if (string.IsNullOrWhiteSpace(title)) return string.Empty;
            title = title!.Trim();

            // 先整体搜索常见格式：数字.数字 + 可选字母数字（如 1.0G11）
            var m = Regex.Match(title, @"\d+\.\d+[A-Za-z0-9]*", RegexOptions.IgnoreCase);
            if (m.Success) return m.Value;

            // 如果没有小数点形式，按分隔符拆分并从后向前查找合适片段
            var parts = title.Split(new[] { '-', '_', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = parts.Length - 1; i >= 0; i--)
            {
                var seg = parts[i].Trim();
                if (string.IsNullOrEmpty(seg)) continue;

                // 常见样式：数字+字母+数字 或 字母+数字（如 1G11 / G11）
                if (Regex.IsMatch(seg, @"^\d+[A-Za-z]+\d*$", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(seg, @"^[A-Za-z]+\d+$", RegexOptions.IgnoreCase))
                {
                    return seg;
                }

                // 也接受含字母后跟数字的片段
                if (Regex.IsMatch(seg, @"[A-Za-z]\d", RegexOptions.IgnoreCase))
                    return seg;
            }

            // 退回到更宽松的全局匹配：数字.数字 或 带字母的数字段
            var m2 = Regex.Match(title, @"\d+\.\d+|[A-Za-z]*\d+[A-Za-z]+\d*", RegexOptions.IgnoreCase);
            if (m2.Success) return m2.Value;

            return string.Empty;
        }

        /// <summary>
        /// 迁移：为 UI 按钮 [插入管道表] 提供命令入口（已迁移到此类）
        /// 简单实现：直接调用 GeneratePipeTableFromSelection，使用统一逻辑（包含选择、分组、插入点提示）。
        /// 如果 WPF 需直接调用此方法，请在 WPF 中触发此命令或直接调用 GeneratePipeTableFromSelection。
        /// </summary>
        [CommandMethod("InsertPipeTable")]
        public void InsertPipeTable()
        {
            // 迁移后直接复用 GeneratePipeTableFromSelection 的逻辑
            GeneratePipeTableFromSelection();
        }

        /// <summary>
        /// 从管道标题中提取管道号，例如 "350-AR-1002-1.0G11" -> "AR-1002" 
        /// 优先返回字母-数字形式的片段，若无法匹配返回 empty 
        /// </summary>
        private string ExtractPipeCodeFromTitle(string? title)
        {
            if (string.IsNullOrWhiteSpace(title)) return string.Empty;
            title = title!.Trim();

            // 常见形式：以字母开头 + '-' + 数字，例如 AR-1002
            var m = Regex.Match(title, @"[A-Za-z]+-\d+");
            if (m.Success) return m.Value;

            // 保险：也尝试在 -...- 中间提取类似模式
            var m2 = Regex.Match(title, @"-(?<code>[A-Za-z]+-\d+)-");
            if (m2.Success && m2.Groups["code"].Success) return m2.Groups["code"].Value;

            // 若仍找不到，可以尝试更宽松的匹配（包含数字在后）
            var m3 = Regex.Match(title, @"[A-Za-z0-9]+-[0-9A-Za-z]+");
            if (m3.Success) return m3.Value;

            return string.Empty;
        }

        /// <summary>
        /// ----------- 辅助：从实体中读取属性（AttributeReference / Xrecord / XData） ------------
        /// </summary>
        /// <param name="tr"></param>
        /// <param name="ent"></param>
        /// <returns></returns>
        private Dictionary<string, string> GetEntityAttributeMap(Transaction tr, Entity ent)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (ent == null) return map;

                // 1) AttributeReference（块参照）
                if (ent is BlockReference br)
                {
                    try
                    {
                        var attCol = br.AttributeCollection;
                        foreach (ObjectId attId in attCol)
                        {
                            try
                            {
                                var ar = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                if (ar != null)
                                {
                                    var tag = (ar.Tag ?? string.Empty).Trim();
                                    var val = (ar.TextString ?? string.Empty).Trim();
                                    if (!string.IsNullOrEmpty(tag) && !map.ContainsKey(tag)) map[tag] = val;
                                }
                            }
                            catch { /* 忽略单个属性读取失败 */ }
                        }
                    }
                    catch { /* 忽略 */ }
                }

                // 2) ExtensionDictionary 的 Xrecord
                try
                {
                    if (ent.ExtensionDictionary != ObjectId.Null)
                    {
                        var extDict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                        if (extDict != null)
                        {
                            foreach (var entry in extDict)
                            {
                                try
                                {
                                    var xrec = tr.GetObject(entry.Value, OpenMode.ForRead) as Xrecord;
                                    if (xrec != null && xrec.Data != null)
                                    {
                                        var vals = xrec.Data.Cast<TypedValue>().Select(tv => tv.Value?.ToString() ?? "").ToArray();
                                        var key = entry.Key ?? string.Empty;
                                        var value = string.Join("|", vals);
                                        if (!map.ContainsKey(key)) map[key] = value;
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                }
                catch { /* 忽略 */ }

                // 3) RegApp XData
                try
                {
                    var db = ent.Database;
                    var rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
                    foreach (ObjectId appId in rat)
                    {
                        try
                        {
                            var app = tr.GetObject(appId, OpenMode.ForRead) as RegAppTableRecord;
                            if (app == null) continue;
                            var appName = app.Name;
                            var rb = ent.GetXDataForApplication(appName);
                            if (rb != null)
                            {
                                var vals = rb.Cast<TypedValue>().Select(tv => tv.Value?.ToString() ?? "").ToArray();
                                var key = $"XDATA:{appName}";
                                var value = string.Join("|", vals);
                                if (!map.ContainsKey(key)) map[key] = value;
                            }
                        }
                        catch { }
                    }
                }
                catch { /* 忽略 */ }
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nGetEntityAttributeMap 异常: {ex.Message}");
            }
            return map;
        }

        /// <summary>
        /// ----------- 辅助：从属性字符串中解析长度（返回米，支持 mm/m） ------------
        /// </summary>
        /// <param name="rawValue"></param>
        /// <returns></returns>
        private double ParseLengthValueFromAttribute(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue)) return double.NaN;
            try
            {
                var s = rawValue.Trim().ToLowerInvariant();
                bool containsMm = s.Contains("mm") || s.Contains("毫米");
                bool containsM = (s.Contains("m") && !containsMm) || s.Contains("米");

                var m = Regex.Match(s, @"[-+]?[0-9]*\.?[0-9]+");
                if (!m.Success) return double.NaN;
                if (!double.TryParse(m.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double value))
                    return double.NaN;

                if (containsMm) return value / 1000.0;
                if (containsM) return value;

                // 无单位启发式：若值 >=1000 视为 mm
                if (value >= 1000.0) return value / 1000.0;
                return value;
            }
            catch { return double.NaN; }
        }


        #region

        /// <summary>
        /// 线段信息,定义存储线段信息的类
        /// </summary>
        public class LineSegmentInfo
        {
            /// <summary>
            /// 中点坐标
            /// </summary>
            public Point3d MidPoint { get; set; }             // 中点坐标
            /// <summary>
            /// 线宽（毫米）
            /// </summary>
            public double LineWeight { get; set; }            // 线宽（毫米）
            /// <summary>
            /// 线段ID
            /// </summary>
            public ObjectId Id { get; set; }                  // 线段ID
            /// <summary>
            /// 起点坐标
            /// </summary>
            public Point3d StartPoint { get; set; }           // 起点坐标
            /// <summary>
            /// 终点坐标
            /// </summary>
            public Point3d EndPoint { get; set; }             // 终点坐标
            /// <summary>
            /// 中点坐标（多段线专用）
            /// </summary>
            public List<Point3d>? MidPoints { get; set; }      // 中间点（多段线专用）
            /// <summary>
            /// 线段长度
            /// </summary>
            public double Length { get; set; }                // 线段长度
            /// <summary>
            /// 线段角度
            /// </summary>
            public double Angle { get; set; }                 // 线段角度（弧度）
            /// <summary>
            /// 所在图层
            /// </summary>
            public string? Layer { get; set; }                 // 所在图层
            /// <summary>
            /// 颜色索引
            /// </summary>
            public int ColorIndex { get; set; }               // 颜色索引
            /// <summary>
            /// 线型比例
            /// </summary>
            public double LinetypeScale { get; set; }         // 线型比例
            /// <summary>
            /// 实体类型
            /// </summary>
            public string? EntityType { get; set; }            // 实体类型（Line/Polyline）
        }

        /// <summary>
        /// 获取所有选择的线段信息
        /// </summary>
        [CommandMethod("CollectLineInfo")]
        public static void CollectLineInfo()
        {
            // 获取当前文档和编辑器
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // 创建选择过滤器，只允许选择直线(LINE)和轻量多段线(LWPOLYLINE)
                TypedValue[] filterValues = new TypedValue[] {
                        new TypedValue((int)DxfCode.Operator, "<OR"),
                        new TypedValue((int)DxfCode.Start, "LINE"),  // 使用 DxfCode.Start
                        new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),  // 使用 DxfCode.Start
                        new TypedValue((int)DxfCode.Operator, "OR>")
                    };

                SelectionFilter filter = new SelectionFilter(filterValues);

                // 设置选择选项
                PromptSelectionOptions opts = new PromptSelectionOptions
                {
                    MessageForAdding = "\n选择线段或多段线: ",
                    AllowDuplicates = false
                };

                // 获取用户选择
                PromptSelectionResult selResult = ed.GetSelection(opts, filter);

                if (selResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\n未选择对象或选择已取消。");
                    return;
                }

                SelectionSet selectionSet = selResult.Value;
                ed.WriteMessage($"\n已选择 {selectionSet.Count} 个对象");

                // 存储所有线段信息的列表
                List<LineSegmentInfo> lineInfos = new List<LineSegmentInfo>();

                // 开始事务处理
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 遍历所有选中的对象
                    foreach (SelectedObject selObj in selectionSet)
                    {
                        if (selObj == null) continue;

                        Entity entity = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;

                        if (entity is Line line)
                        {
                            ed.WriteMessage("\n找到直线对象");
                            lineInfos.Add(ProcessLine(line, tr));
                        }
                        else if (entity is Polyline pline) // 处理轻量多段线
                        {
                            ed.WriteMessage("\n找到多段线对象");
                            lineInfos.Add(ProcessPolyline(pline, tr));
                        }
                        else
                        {
                            ed.WriteMessage($"\n跳过不支持的类型: {entity?.GetType().Name}");
                        }
                    }

                    // 输出收集到的信息
                    if (lineInfos.Count > 0)
                    {
                        PrintLineInfos(ed, lineInfos);
                    }
                    else
                    {
                        ed.WriteMessage("\n未找到可处理的线段对象");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n错误: {ex.Message}\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// 处理直线(LINE)对象
        /// </summary>
        /// <param name="line"></param>
        /// <param name="tr"></param>
        /// <returns></returns>
        private static LineSegmentInfo ProcessLine(Line line, Transaction tr)
        {
            // 计算线段角度（与X轴正方向的夹角，弧度）
            Vector3d lineVector = line.EndPoint - line.StartPoint;
            return new LineSegmentInfo
            {
                Id = line.ObjectId,
                StartPoint = line.StartPoint,
                EndPoint = line.EndPoint,
                MidPoints = new List<Point3d>(),
                Length = line.Length,
                Angle = lineVector.GetAngleTo(Vector3d.XAxis),
                Layer = GetLayerName(line.LayerId, tr),
                ColorIndex = line.ColorIndex,
                LinetypeScale = line.LinetypeScale,
                EntityType = "LINE"
            };
        }

        /// <summary>
        /// 处理多段线(POLYLINE)对象
        /// </summary>
        /// <param name="pline"></param>
        /// <param name="tr"></param>
        /// <returns></returns>
        private static LineSegmentInfo ProcessPolyline(Polyline pline, Transaction tr)
        {
            List<Point3d> vertices = new List<Point3d>();
            int numVertices = pline.NumberOfVertices;

            // 获取所有顶点
            for (int i = 0; i < numVertices; i++)
            {
                vertices.Add(pline.GetPoint3dAt(i));
            }

            // 计算总长度（考虑闭合情况）
            double totalLength = pline.Length;

            // 获取第一段线段的角度
            double angle = 0;
            if (numVertices >= 2)
            {
                Vector3d vector = pline.GetPoint3dAt(1) - pline.GetPoint3dAt(0);
                angle = vector.AngleOnPlane(new Plane(Point3d.Origin, Vector3d.ZAxis));
            }

            return new LineSegmentInfo
            {
                Id = pline.ObjectId,
                StartPoint = pline.StartPoint,
                EndPoint = pline.EndPoint,
                MidPoints = vertices.Skip(1).Take(vertices.Count - 2).ToList(),
                Length = totalLength,
                Angle = angle,
                Layer = GetLayerName(pline.LayerId, tr),
                ColorIndex = pline.ColorIndex,
                LinetypeScale = pline.LinetypeScale,
                EntityType = "LWPOLYLINE"
            };
        }

        /// <summary>
        /// 根据图层ID获取图层名称
        /// </summary>
        /// <param name="layerId"></param>
        /// <param name="tr"></param>
        /// <returns></returns>
        private static string GetLayerName(ObjectId layerId, Transaction tr)
        {
            if (layerId.IsNull) return "0";

            try
            {
                LayerTableRecord ltr = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
                return ltr?.Name ?? "0";
            }
            catch
            {
                return "0";
            }
        }

        /// <summary>
        /// 输出收集到的线段信息
        /// </summary>
        /// <param name="ed"></param>
        /// <param name="infos"></param>
        private static void PrintLineInfos(Editor ed, List<LineSegmentInfo> infos)
        {
            ed.WriteMessage("\n\n===== 线段信息报告 =====");
            ed.WriteMessage($"\n共处理 {infos.Count} 个线段对象");

            foreach (var info in infos)
            {
                ed.WriteMessage("\n--------------------------------");
                ed.WriteMessage($"\n对象ID: {info.Id}");
                ed.WriteMessage($"\n类型: {info.EntityType}");
                ed.WriteMessage($"\n起点: X={info.StartPoint.X:F2}, Y={info.StartPoint.Y:F2}, Z={info.StartPoint.Z:F2}");
                ed.WriteMessage($"\n终点: X={info.EndPoint.X:F2}, Y={info.EndPoint.Y:F2}, Z={info.EndPoint.Z:F2}");

                if (info.MidPoints?.Count > 0)
                {
                    ed.WriteMessage($"\n中间点({info.MidPoints.Count}个):");
                    foreach (var pt in info.MidPoints)
                    {
                        ed.WriteMessage($"\n  X={pt.X:F2}, Y={pt.Y:F2}, Z={pt.Z:F2}");
                    }
                }

                ed.WriteMessage($"\n长度: {info.Length:F2}");
                ed.WriteMessage($"\n角度: {RadiansToDegrees(info.Angle):F1}°");
                ed.WriteMessage($"\n图层: {info.Layer}");
                ed.WriteMessage($"\n颜色索引: {info.ColorIndex}");
                ed.WriteMessage($"\n线型比例: {info.LinetypeScale:F2}");
            }

            ed.WriteMessage("\n\n===== 报告结束 =====");
        }

        /// <summary>
        /// 弧度转角度
        /// </summary>
        /// <param name="radians"></param>
        /// <returns></returns>
        private static double RadiansToDegrees(double radians)
        {
            return radians * (180.0 / Math.PI);
        }

        #endregion

        #region 可直接替换的增强实现：在起点/终点自动拾取最近部件（块参照），读取属性，并按白名单把同名字段写回管道属性。

        // 端点拾取属性白名单（按你项目可继续补充）
        private static readonly HashSet<string> _pipeEndpointAttrWhitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
         {
             "位号","设备位号","设备名称","名称","Name","Tag","TAG",
             "介质","介质名称","规格","型号","材质","管径","DN","PN",
             "管道标题","管段号","流向","系统","压力","温度","备注","Remark"
         };

        #endregion

        #region 同步管道的实现方法

        /// <summary>
        /// 存储从示例块中分析出的管道信息
        /// </summary>
        public class SamplePipeInfo
        {
            /// <summary>
            /// 模板管道
            /// </summary>
            public Polyline? PipeBodyTemplate { get; set; }
            /// <summary>
            /// 模板方向箭头
            /// </summary>
            public Polyline? DirectionArrowTemplate { get; set; }
            /// <summary>
            /// 属性定义
            /// </summary>
            public List<AttributeDefinition> AttributeDefinitions { get; set; } = new List<AttributeDefinition>();
            /// <summary>
            /// 基点
            /// </summary>
            public Point3d BasePoint { get; set; }
        }


        /// <summary>
        /// 同步管道\属性
        /// </summary>        
        [CommandMethod("SyncPipeProperties")]
        public void SyncPipeProperties()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 选择线段
            var lineSelResult = ed.GetSelection(
                new PromptSelectionOptions { MessageForAdding = "\n请选择要同步的线段 (LINE 或 LWPOLYLINE):" },
                new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "LINE,LWPOLYLINE") })
            );
            if (lineSelResult.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n操作取消。");
                return;
            }
            var sourceLineIds = lineSelResult.Value.GetObjectIds().ToList();

            // 选择示例管线块（作为样例）
            var blockSelResult = ed.GetEntity("\n请选择示例管线块:");
            if (blockSelResult.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n操作取消。");
                return;
            }

            using (var tr = new DBTrans())
            {
                try
                {
                    // 读取示例块参照
                    var sampleBlockRef = tr.GetObject(blockSelResult.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (sampleBlockRef == null)
                    {
                        ed.WriteMessage("\n错误：选择的不是块参照。");
                        return;
                    }

                    // 解析示例块（提取 polyline / arrow / attribute definitions）
                    var sampleInfo = AnalyzeSampleBlock(tr, sampleBlockRef);
                    // 检查模板
                    if (sampleInfo?.PipeBodyTemplate == null)
                    {
                        ed.WriteMessage("\n错误：示例块中未找到作为管道主体的 Polyline。");
                        return;
                    }

                    // 收集并构建顶点顺序
                    var lineSegments = CollectLineSegments(tr, sourceLineIds);
                    if (lineSegments == null || lineSegments.Count == 0)
                    {
                        ed.WriteMessage("\n未找到可处理的线段。");
                        return;
                    }
                    // 构建顶点顺序
                    var orderedVertices = BuildOrderedVerticesFromSegments(lineSegments, 0.1);
                    if (orderedVertices == null || orderedVertices.Count < 2)
                    {
                        ed.WriteMessage("\n顶点不足，无法生成管线。");
                        return;
                    }

                    // 读取示例块的所有属性
                    var sampleAttrMap = GetEntityAttributeMap(tr, sampleBlockRef) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    // 加载上次保存的属性（此处取决于 sampleBlockRef 名称中是否包含入口/出口关键词）
                    bool sampleIsOutlet = (sampleBlockRef.Name ?? string.Empty).ToLowerInvariant().Contains("出口") ||
                                          (sampleBlockRef.Name ?? string.Empty).ToLowerInvariant().Contains("outlet");

                    // 从磁盘读取历史属性
                    var savedAttrsSync = FileManager.LoadLastPipeAttributes(sampleIsOutlet);

                    // 打开属性编辑窗，传入合并后的初始字典
                    using (var editor = new PipeAttributeEditorForm(savedAttrsSync))
                    {
                        var dr = editor.ShowDialog();
                        if (dr != DialogResult.OK)
                        {
                            ed.WriteMessage("\n已取消属性编辑，停止同步操作。");
                            return;
                        }

                        // 保存历史属性供下次使用
                        var editedAttrs = editor.Attributes ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        FileManager.SaveLastPipeAttributes(sampleIsOutlet, editedAttrs);

                        // 把用户修改后的属性写回示例块（只写存在的 AttributeReference）
                        try
                        {
                            var sampleBrWrite = tr.GetObject(sampleBlockRef.ObjectId, OpenMode.ForWrite) as BlockReference;
                            if (sampleBrWrite != null)
                            {
                                // 遍历属性
                                foreach (ObjectId aid in sampleBrWrite.AttributeCollection)
                                {
                                    try
                                    {
                                        // 获取属性引用
                                        var ar = tr.GetObject(aid, OpenMode.ForWrite) as AttributeReference;
                                        // 跳过无效的属性
                                        if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue;
                                        // 检查属性是否在编辑字典中
                                        if (editedAttrs.TryGetValue(ar.Tag, out var newVal))
                                        {
                                            // 更新属性
                                            ar.TextString = newVal ?? string.Empty;
                                            // 对齐
                                            try { ar.AdjustAlignment(db); } catch { }
                                        }
                                    }
                                    catch { /* 单个属性写回失败不阻塞整体 */ }
                                }
                            }
                        }
                        catch (System.Exception exWriteSample)
                        {
                            ed.WriteMessage($"\n写回示例块属性时出错: {exWriteSample.Message}");
                        }
                    }
                    // 开始构建新管道块
                    double pipelineLength = 0.0;
                    for (int i = 0; i < orderedVertices.Count - 1; i++)
                        pipelineLength += orderedVertices[i].DistanceTo(orderedVertices[i + 1]);
                    // 计算中点
                    var (midPoint, midAngle) = ComputeMidPointAndAngle(orderedVertices, pipelineLength);
                    // 计算目标向量
                    Vector3d targetDir = ComputeDirectionAtPoint(orderedVertices, midPoint, 1e-6);
                    // 计算聚合线段向量
                    Vector3d segmentDir = ComputeAggregateSegmentDirection(lineSegments);
                    // 如果聚合线段向量与目标向量方向相反，则反转目标向量
                    if (!segmentDir.IsZeroLength() && targetDir.DotProduct(segmentDir) < 0)
                        targetDir = -targetDir;
                    // 归一化目标向量  
                    Vector3d targetDirNormalized = targetDir.IsZeroLength() ? Vector3d.XAxis : targetDir.GetNormal();
                    // 构建管道
                    Polyline pipeLocal = BuildPipePolylineLocal(sampleInfo.PipeBodyTemplate, orderedVertices, midPoint);

                    // 复制属性定义（基于示例块的定义）——先克隆示例定义（保持字段顺序与名称）
                    var attDefsLocal = CloneAttributeDefinitionsLocal(sampleInfo.AttributeDefinitions, midPoint, 0.0, pipelineLength, sampleBlockRef.Name)
                                        ?? new List<AttributeDefinition>();

                    // 重新读取示例块属性（刚刚可能已被编辑并写回）
                    var latestSampleAttrs = GetEntityAttributeMap(tr, sampleBlockRef) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    // 生成标题（优先属性中的管道标题）
                    string pipeTitle = latestSampleAttrs.TryGetValue("管道标题", out var sampleTitle) && !string.IsNullOrWhiteSpace(sampleTitle)
                                       ? sampleTitle
                                       : sampleBlockRef.Name ?? "管道";

                    // 为每一段生成局部坐标下的箭头与标题（相对于 midPoint），仅使用分段箭头/文字
                    var arrowEntities = CreateDirectionalArrowsAndTitles(tr, sampleInfo, orderedVertices, midPoint, pipeTitle, sampleBlockRef.Name);

                    // ---------- 关键修正：确保新块的属性定义字段严格以示例块的 AttributeDefinitions 为准 ----------
                    // 1) 若示例存在属性定义，则按照示例顺序保留字段，仅更新 TextString（不新增示例中不存在的字段）
                    // 2) 若示例没有属性定义，则以 latestSampleAttrs 为基础创建属性定义（按键排序以保证稳定性）
                    if (sampleInfo.AttributeDefinitions != null && sampleInfo.AttributeDefinitions.Count > 0)
                    {
                        // 使用示例定义作为基准，更新文本值（保留原有位置/高度/顺序）
                        var latestDict = new Dictionary<string, string>(latestSampleAttrs, StringComparer.OrdinalIgnoreCase);
                        foreach (var def in attDefsLocal)
                        {
                            if (string.IsNullOrWhiteSpace(def.Tag)) continue;
                            if (latestDict.TryGetValue(def.Tag, out var val))
                            {
                                def.TextString = val ?? string.Empty;
                            }
                            // 临时显示设置（随后统一隐藏/显示处理）
                            def.Invisible = false;
                            def.Constant = false;
                        }
                    }
                    else
                    {
                        // 示例无属性定义：根据 latestSampleAttrs 动态创建属性定义（按 Key 排序）
                        attDefsLocal.Clear();
                        double attHeight = 3.5;
                        double yOffsetBase = -attHeight * 2.0;
                        int idx = 0;
                        foreach (var kv in latestSampleAttrs.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
                        {
                            if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                            attDefsLocal.Add(new AttributeDefinition
                            {
                                Tag = kv.Key,
                                Position = new Point3d(0, yOffsetBase - idx * attHeight * 1.2, 0),
                                Rotation = 0.0,
                                TextString = kv.Value ?? string.Empty,
                                Height = attHeight,
                                Invisible = false,
                                Constant = false
                            });
                            idx++;
                        }
                    }

                    // 新增或覆盖 起点/终点 属性定义（保持原逻辑，若示例中已有这些字段则更新其值，否则新增）
                    Point3d worldStart = orderedVertices.First();
                    Point3d worldEnd = orderedVertices.Last();
                    string startCoordStr = $"X={worldStart.X:F3},Y={worldStart.Y:F3}";
                    string endCoordStr = $"X={worldEnd.X:F3},Y={worldEnd.Y:F3}";
                    int nextSegNum = GetNextPipeSegmentNumber(db);

                    // 取管段号，优先从属性或标题提取
                    string extractedPipeNo = string.Empty;
                    if (latestSampleAttrs.TryGetValue("管道标题", out var titleFromSample) && !string.IsNullOrWhiteSpace(titleFromSample))
                    {
                        extractedPipeNo = ExtractPipeCodeFromTitle(titleFromSample);
                    }
                    if (string.IsNullOrWhiteSpace(extractedPipeNo))
                    {
                        extractedPipeNo = ExtractPipeCodeFromTitle(sampleBlockRef.Name);
                    }
                    if (string.IsNullOrWhiteSpace(extractedPipeNo))
                    {
                        if (latestSampleAttrs.TryGetValue("管段号", out var pn) && !string.IsNullOrWhiteSpace(pn))
                            extractedPipeNo = pn;
                        else if (latestSampleAttrs.TryGetValue("管段编号", out var pn2) && !string.IsNullOrWhiteSpace(pn2))
                            extractedPipeNo = pn2;
                    }
                    if (string.IsNullOrWhiteSpace(extractedPipeNo))
                    {
                        extractedPipeNo = nextSegNum.ToString("D4");
                    }

                    // 局部函数：按示例定义优先更新 / 新增（仅当示例中不存在该字段时新增）管道号
                    void SetOrAddAttrLocal(string tag, string text)
                    {
                        // 按示例定义优先更新 查找是否已存在该属性定义
                        var existing = attDefsLocal.FirstOrDefault(a => string.Equals(a.Tag, tag, StringComparison.OrdinalIgnoreCase));
                        if (existing != null)
                        {
                            existing.TextString = text;// 更新属性值 按示例定义优先更新
                            existing.Invisible = false;// 临时显示设置（随后统一隐藏/显示处理）
                            existing.Constant = false;// 临时常量设置（随后统一取消常量处理）
                        }
                        else
                        {
                            // 只有在示例没有任何定义时允许新增；但为兼容性仍允许新增起/终/段号
                            attDefsLocal.Add(new AttributeDefinition
                            {
                                Tag = tag,
                                Position = new Point3d(0, (attDefsLocal.Count > 0 ? attDefsLocal[0].Position.Y - attDefsLocal[0].Height * 1.2 : -3.5), 0),
                                Rotation = 0.0,
                                TextString = text,
                                Height = attDefsLocal.Count > 0 ? attDefsLocal[0].Height : 3.5,
                                Invisible = false,
                                Constant = false
                            });
                        }
                    }

                    SetOrAddAttrLocal("始点", startCoordStr);
                    SetOrAddAttrLocal("终点", endCoordStr);
                    SetOrAddAttrLocal("管段号", extractedPipeNo);

                    //// 移除块定义中的中点“管道标题”属性（避免在块中重复显示中点标题）
                    //attDefsLocal.RemoveAll(ad => string.Equals(ad.Tag, "管道标题", StringComparison.OrdinalIgnoreCase));

                    //// 将其余属性设置为隐藏（块内不显示），以保持行为与现有逻辑一致
                    foreach (var ad in attDefsLocal)
                    {
                        ad.Invisible = true;
                        ad.Constant = false;
                    }

                    // 构建块定义并插入新块
                    string desiredName = sampleBlockRef.Name;
                    // 创建块定义 若名称已存在则添加后缀
                    string newBlockName = BuildPipeBlockDefinition(tr, desiredName, (Polyline)pipeLocal.Clone(), arrowEntities, attDefsLocal);
                    // 准备属性值字典
                    var attValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var a in attDefsLocal)
                    {
                        if (string.IsNullOrWhiteSpace(a?.Tag)) continue;// 跳过无效标签 准备属性值字典
                        attValues[a.Tag] = a.TextString ?? string.Empty;//  收集属性值
                    }
                    // 插入新块并设置属性
                    //var newBrId = InsertPipeBlockWithAttributes(tr, midPoint, newBlockName, 0.0, attValues);
                    // 在插入新管道块后（newBrId 拿到后）调用后处理
                    var newBrId = InsertPipeBlockWithAttributes(tr, midPoint, newBlockName, 0.0, attValues);
                    PipelineTopologyHelper.PostProcessAfterPipePlaced(tr, newBrId, sampleIsOutlet ? "Outlet" : "Inlet");
                    var newBr = tr.GetObject(newBrId, OpenMode.ForWrite) as BlockReference;// 插入新块并设置属性
                    if (newBr != null)
                        newBr.Layer = sampleInfo.PipeBodyTemplate.Layer;// 插入新块并设置属性 继承图层

                    // 删除原始线段
                    foreach (var seg in lineSegments)
                    {
                        var ent = tr.GetObject(seg.Id, OpenMode.ForWrite) as Entity;// 删除原始线段
                        if (ent != null)
                            ent.Erase();
                    }

                    tr.Commit();
                    ed.WriteMessage($"\n管线块已生成：新增/更新属性 [始点][终点][管段号]={extractedPipeNo}。仅显示字段：管道标题。");
                }
                catch (Exception ex)
                {
                    ed.WriteMessage($"\n发生错误: {ex.Message}\n{ex.StackTrace}");
                    tr.Abort();
                }
            }
        }

        /// <summary>
        /// 新增表单窗口：PipeAttributeEditorForm —— 编辑示例图元的属性表（键不可改，值可编辑）
        /// </summary>
        public class PipeAttributeEditorForm : Form
        {
            /// <summary>
            /// 属性表 属性表网格
            /// </summary>
            private DataGridView _dataGridView;
            /// <summary>
            /// 确认按钮
            /// </summary>
            private Button _btnOk;// 确认按钮
            /// <summary>
            /// 取消按钮
            /// </summary>
            private Button _btnCancel;// 取消按钮 确认和取消按钮
            /// <summary>
            /// 属性表
            /// </summary>
            private Dictionary<string, string> _attributes;// 属性表/ 存储属性表

            /// <summary>
            /// 属性表编辑后的属性表
            /// </summary>
            public Dictionary<string, string> Attributes => new Dictionary<string, string>(_attributes, StringComparer.OrdinalIgnoreCase);

            /// <summary>
            /// 属性表编辑窗口
            /// </summary>
            /// <param name="initialAttributes"></param>
            public PipeAttributeEditorForm(Dictionary<string, string> initialAttributes)
            {
                _attributes = new Dictionary<string, string>(initialAttributes ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase);
                InitializeComponent();
                LoadAttributesToGrid();
            }

            /// <summary>
            /// 初始化控件
            /// </summary>
            private void InitializeComponent()
            {
                this.Text = "示例管道属性编辑";
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.StartPosition = FormStartPosition.CenterParent;
                this.ClientSize = new Size(640, 420);
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.MinimizeBox = false;
                this.ShowInTaskbar = false;
                this.AutoScaleMode = AutoScaleMode.Font;

                _dataGridView = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    RowHeadersVisible = false,
                    SelectionMode = DataGridViewSelectionMode.CellSelect,
                    MultiSelect = false
                };

                var colKey = new DataGridViewTextBoxColumn { Name = "Key", HeaderText = "字段", ReadOnly = true };
                var colVal = new DataGridViewTextBoxColumn { Name = "Value", HeaderText = "值", ReadOnly = false };

                _dataGridView.Columns.Add(colKey);
                _dataGridView.Columns.Add(colVal);

                _btnOk = new Button { Text = "完成", DialogResult = DialogResult.OK, Width = 90, Height = 30 };
                _btnCancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Width = 90, Height = 30 };

                _btnOk.Click += BtnOk_Click;
                _btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

                // 底部按钮面板，右对齐
                var panel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    Height = 50,
                    FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                    Padding = new Padding(8),
                    WrapContents = false
                };

                // 添加按钮到面板（右到左）
                panel.Controls.Add(_btnCancel);
                panel.Controls.Add(_btnOk);

                // 设置接受/取消按钮
                this.AcceptButton = _btnOk;
                this.CancelButton = _btnCancel;

                // 按钮与网格先后添加，保证 DockFill 占满剩余空间
                this.Controls.Add(_dataGridView);
                this.Controls.Add(panel);
            }

            /// <summary>
            /// 填充属性表
            /// </summary>
            private void LoadAttributesToGrid()
            {
                _dataGridView.Rows.Clear();
                foreach (var kv in _attributes.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
                {
                    _dataGridView.Rows.Add(kv.Key, kv.Value);
                }
                if (_dataGridView.Rows.Count > 0)
                    _dataGridView.CurrentCell = _dataGridView.Rows[0].Cells[1];
            }

            /// <summary>
            /// 确认按钮点击事件
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            private void BtnOk_Click(object sender, EventArgs e)
            {
                // 保存网格中用户编辑的值回 _attributes
                try
                {
                    var newDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < _dataGridView.Rows.Count; i++)
                    {
                        var row = _dataGridView.Rows[i];
                        if (row.IsNewRow) continue;
                        var keyCell = row.Cells["Key"].Value;
                        var valCell = row.Cells["Value"].Value;
                        if (keyCell == null) continue;
                        string key = keyCell.ToString() ?? string.Empty;
                        string val = valCell?.ToString() ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(key)) continue;
                        newDict[key] = val;
                    }
                    _attributes = newDict;
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("保存属性失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// 扫描模型空间中已有块引用的属性，找出最大已用管段号并返回下一个编号（整数）
        /// 编号规则：解析属性值内的首个连续数字序列作为编号；若无，跳过。
        /// </summary>
        /// <param name="db">当前数据库</param>
        /// <returns>下一个管段号（从 1 开始）</returns>
        private int GetNextPipeSegmentNumber(Database db)
        {
            int max = 0;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                    foreach (ObjectId id in ms)
                    {
                        try
                        {
                            var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (ent is BlockReference br)
                            {
                                foreach (ObjectId aid in br.AttributeCollection)
                                {
                                    try
                                    {
                                        var ar = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                                        if (ar == null) continue;
                                        if (!string.Equals(ar.Tag, "管段号", StringComparison.OrdinalIgnoreCase) &&
                                            !string.Equals(ar.Tag, "管段编号", StringComparison.OrdinalIgnoreCase))
                                            continue;

                                        string txt = ar.TextString ?? string.Empty;
                                        if (string.IsNullOrWhiteSpace(txt)) continue;

                                        // 提取首个连续数字序列
                                        var m = Regex.Match(txt, @"\d+");
                                        if (m.Success && int.TryParse(m.Value, out int val))
                                        {
                                            if (val > max) max = val;
                                        }
                                    }
                                    catch { /* 忽略单个属性读取问题 */ }
                                }
                            }
                        }
                        catch { /* 忽略单个实体读取问题 */ }
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }
            return max + 1;
        }

        /// <summary>
        /// 将箭头几何按照指定方向对齐
        /// </summary>
        private (Polyline outline, Solid? fill) AlignArrowToDirection(Polyline arrowTemplate, Solid? fillTemplate, Vector3d direction)
        {
            // 计算模板主方向
            Vector3d dir = direction.IsZeroLength() ? Vector3d.XAxis : direction.GetNormal();
            // 计算模板侧向
            Vector3d yAxis = Vector3d.ZAxis.CrossProduct(dir);
            if (yAxis.IsZeroLength())// 如果主向和侧向平行，则侧向为 Y 轴
                yAxis = Vector3d.YAxis;// 侧向为 Z 轴
            else
                yAxis = yAxis.GetNormal();// 计算侧向
            // 计算对齐矩阵
            Matrix3d alignMatrix = Matrix3d.AlignCoordinateSystem(
                Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                Point3d.Origin, dir, yAxis, Vector3d.ZAxis
            );
            // 对齐模板
            var outline = (Polyline)arrowTemplate.Clone();
            outline.TransformBy(alignMatrix);// 对齐
            // 对齐填充
            Solid? fill = null;
            if (fillTemplate != null)
            {
                // 对齐填充
                fill = (Solid)fillTemplate.Clone();
                fill.TransformBy(alignMatrix);// 对齐
            }
            return (outline, fill);
        }

        /// <summary>
        /// 获取箭头
        /// </summary>
        /// <param name="segments"></param>
        /// <returns></returns>
        private static Vector3d ComputeAggregateSegmentDirection(List<LineSegmentInfo> segments)
        {
            if (segments == null || segments.Count == 0)
                return new Vector3d(0, 0, 0);

            Vector3d sum = new Vector3d(0, 0, 0);
            foreach (var seg in segments)
            {
                Vector3d dir = seg.EndPoint - seg.StartPoint;
                if (!dir.IsZeroLength())
                    sum += dir.GetNormal();
            }

            return sum.IsZeroLength() ? new Vector3d(0, 0, 0) : sum.GetNormal();
        }

        /// <summary>
        /// 计算某点附近的方向向量（优先使用与 referencePoint 最近的线段）
        /// </summary>
        private static Vector3d ComputeDirectionAtPoint(List<Point3d> orderedVertices, Point3d referencePoint, double tol = 1e-6)
        {
            if (orderedVertices == null || orderedVertices.Count < 2)
                return Vector3d.XAxis;

            Vector3d fallbackDir = ComputePathDirectionVector(orderedVertices, tol);
            double bestDist = double.MaxValue;
            Vector3d bestDir = fallbackDir.IsZeroLength() ? Vector3d.XAxis : fallbackDir;

            for (int i = 0; i < orderedVertices.Count - 1; i++)
            {
                Point3d start = orderedVertices[i];
                Point3d end = orderedVertices[i + 1];
                Vector3d segment = end - start;
                if (segment.IsZeroLength())
                    continue;

                Point3d projected = ProjectPointToSegment(referencePoint, start, end);
                double dist = referencePoint.DistanceTo(projected);
                if (dist + tol < bestDist)
                {
                    bestDist = dist;
                    bestDir = segment.GetNormal();
                }
            }

            if (!bestDir.IsZeroLength() && !fallbackDir.IsZeroLength() && bestDir.DotProduct(fallbackDir) < 0)
            {
                bestDir = -bestDir;
            }

            return bestDir.IsZeroLength() ? fallbackDir : bestDir;
        }

        /// <summary>
        /// 将点投影到指定线段上
        /// </summary>
        private static Point3d ProjectPointToSegment(Point3d point, Point3d segmentStart, Point3d segmentEnd)
        {
            Vector3d segment = segmentEnd - segmentStart;
            if (segment.IsZeroLength())
                return segmentStart;

            Vector3d toPoint = point - segmentStart;
            double t = toPoint.DotProduct(segment) / segment.DotProduct(segment);
            t = Math.Max(0.0, Math.Min(1.0, t));
            return segmentStart + segment * t;
        }

        /// <summary>
        /// 计算整条路径的总体方向向量（UCS，Z=+）
        /// </summary>
        private static Vector3d ComputePathDirectionVector(List<Point3d> orderedVertices, double tol = 1e-6)
        {
            if (orderedVertices == null || orderedVertices.Count < 2)
                return Vector3d.XAxis;

            // 直接用整体起点→终点的向量，保证箭头指向终点（流向）
            Vector3d overall = orderedVertices.Last() - orderedVertices.First();
            if (overall.Length > tol)
                return overall.GetNormal();

            // 回退：选择最长段方向
            double maxLen = 0.0;
            Vector3d longestDir = Vector3d.XAxis;
            for (int i = 0; i < orderedVertices.Count - 1; i++)
            {
                Vector3d v = orderedVertices[i + 1] - orderedVertices[i];
                if (v.Length > maxLen)
                {
                    maxLen = v.Length;
                    longestDir = v.GetNormal();
                }
            }
            return longestDir;
        }

        /// <summary>
        /// 获取选择的线段信息
        /// </summary>
        /// <param name="orderedVertices">有序顶点列表</param>
        /// <param name="totalLength">总长度</param>
        /// <returns></returns>
        private (Point3d midPoint, double midAngle) ComputeMidPointAndAngle(List<Point3d> orderedVertices, double totalLength)
        {
            double halfLen = totalLength / 2.0;
            double acc = 0.0;
            Point3d midPoint = orderedVertices[0];
            double midAngle = 0.0;

            for (int i = 0; i < orderedVertices.Count - 1; i++)
            {
                var p1 = orderedVertices[i];
                var p2 = orderedVertices[i + 1];
                double segLen = p1.DistanceTo(p2);
                if (acc + segLen >= halfLen)
                {
                    double t = (halfLen - acc) / segLen;
                    midPoint = new Point3d(
                        p1.X + (p2.X - p1.X) * t,
                        p1.Y + (p2.Y - p1.Y) * t,
                        p1.Z + (p2.Z - p1.Z) * t
                    );
                    midAngle = ComputeSegmentAngleUcs(p1, p2);
                    break;
                }
                acc += segLen;
            }
            return (midPoint, midAngle);
        }

        /// <summary>
        /// 计算线段角度
        /// </summary>
        /// <param name="p1">起点</param>
        /// <param name="p2">终点</param>
        /// <returns>线段在UCS中的角度</returns>
        private static double ComputeSegmentAngleUcs(Point3d p1, Point3d p2)
        {
            // 当前UCS的XY平面，保证与AutoCAD旋转角同一参考
            var plane = new Plane(Point3d.Origin, Vector3d.ZAxis);
            Vector3d dir = (p2 - p1).GetNormal();
            double angle = dir.AngleOnPlane(plane); // 以正X为0，逆时针为正
                                                    // 归一化到 [0, 2π)
            if (angle < 0) angle += 2.0 * Math.PI;
            return angle;
        }

        /// <summary>
        /// 构建局部坐标的管线 Polyline
        /// </summary>
        /// <param name="template">模板 Polyline</param>
        /// <param name="verticesWorld">全局坐标系下的顶点列表</param>
        /// <param name="midPointWorld">全局坐标系下的中点</param>
        /// <returns>局部坐标系下的管线 Polyline</returns>
        private Polyline BuildPipePolylineLocal(Polyline template, List<Point3d> verticesWorld, Point3d midPointWorld)
        {
            var pl = new Polyline();
            double lineWeightScale = VariableDictionary.textBoxScale;
            for (int i = 0; i < verticesWorld.Count; i++)
            {
                var local = new Point2d(verticesWorld[i].X - midPointWorld.X, verticesWorld[i].Y - midPointWorld.Y);
                //var local = new Point2d(verticesWorld[i].X, verticesWorld[i].Y);
                pl.AddVertexAt(i, local, 0,
                    template.ConstantWidth * lineWeightScale,
                    template.ConstantWidth * lineWeightScale);
            }

            pl.Layer = template.Layer;
            pl.Color = template.Color;
            pl.LineWeight = template.LineWeight;
            pl.Linetype = template.Linetype;
            pl.LinetypeScale = template.LinetypeScale;
            pl.Elevation = 0;
            pl.Normal = Vector3d.ZAxis;
            pl.Closed = false;
            return pl;
        }

        /// <summary>
        /// 新增：创建方向箭头（轮廓 + 填充）
        /// </summary>
        /// <param name="arrowLength">箭头长度</param>
        /// <param name="arrowHeight">箭头高度</param>
        /// <param name="colorIndex">颜色索引</param>
        /// <param name="pipeTemplate">管道模板</param>
        /// <returns>轮廓和填充的元组</returns>
        private (Polyline outline, Solid fill) CreateArrowTriangleFilled(double arrowLength, double arrowHeight, short colorIndex, Polyline pipeTemplate)
        {
            // 三角顶点（局部坐标，尖端朝 +X）
            var tip = new Point2d(arrowLength / 2.0, 0.0);
            var leftBottom = new Point2d(-arrowLength / 2.0, -arrowHeight / 2.0);
            var leftTop = new Point2d(-arrowLength / 2.0, arrowHeight / 2.0);

            // 轮廓
            var arrow = new Polyline();
            arrow.AddVertexAt(0, tip, 0, 0, 0);
            arrow.AddVertexAt(1, leftBottom, 0, 0, 0);
            arrow.AddVertexAt(2, leftTop, 0, 0, 0);
            arrow.Closed = true;
            arrow.Layer = pipeTemplate.Layer;
            //arrow.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex);
            //arrow.Linetype = pipeTemplate.Linetype;
            arrow.LinetypeScale = pipeTemplate.LinetypeScale;
            arrow.LineWeight = pipeTemplate.LineWeight;
            arrow.Elevation = 0;
            arrow.Normal = Vector3d.ZAxis;

            // 填充（二维实心三角形）
            var solid = new Solid(
                new Point3d(tip.X, tip.Y, 0),
                new Point3d(leftBottom.X, leftBottom.Y, 0),
                new Point3d(leftTop.X, leftTop.Y, 0),
                new Point3d(leftTop.X, leftTop.Y, 0) // 三角形第四点与第三点相同
            );
            solid.Layer = pipeTemplate.Layer;
            solid.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex);
            solid.LineWeight = pipeTemplate.LineWeight;
            solid.Normal = Vector3d.ZAxis;

            return (arrow, solid);
        }

        /// <summary>
        /// 修改：块定义构建，加入额外实体（管线 + 箭头）
        /// </summary>
        /// <param name="tr">数据库事务</param>
        /// <param name="desiredName">期望的块名称</param>
        /// <param name="pipeLocal">管道轮廓</param>
        /// <param name="overlayEntities">附加实体列表（一般为箭头轮廓和填充）</param>
        /// <param name="attDefsLocal">属性定义列表</param>
        /// <returns>块定义名称</returns>
        private string BuildPipeBlockDefinition(DBTrans tr, string desiredName, Polyline pipeLocal, List<Entity> overlayEntities, List<AttributeDefinition> attDefsLocal)
        {
            string finalName = desiredName;
            int suf = 1;
            while (tr.BlockTable.Has(finalName))
                finalName = desiredName + "_PIPEGEN_" + suf++;

            tr.BlockTable.Add(
                finalName,
                btr =>
                {
                    btr.Origin = Point3d.Origin;
                },
                () =>
                {
                    var entities = new List<Entity>
                    {
                        (Polyline)pipeLocal.Clone()
                    };
                    if (overlayEntities != null)
                    {
                        foreach (var entity in overlayEntities)
                        {
                            if (entity == null)
                                continue;

                            var clone = entity.Clone() as Entity;
                            if (clone != null)
                                entities.Add(clone);
                        }
                    }
                    return entities;
                },
                () => attDefsLocal
            );

            return finalName;
        }

        /// <summary>
        /// 根据名称确定箭头样式
        /// </summary>
        /// <param name="blockName">块名称</param>
        /// <returns>箭头样式元组</returns>
        private (short colorIndex, double length, double height) DetermineArrowStyleByName(string blockName)
        {
            string nameLower = (blockName ?? string.Empty).ToLowerInvariant();
            bool isOutlet = nameLower.Contains("出口") || nameLower.Contains("outlet");
            bool isInlet = nameLower.Contains("入口") || nameLower.Contains("inlet");

            // 出口=黄色(ACI 2)，入口=绿色(ACI 3)，默认黄色
            short colorIndex = isInlet ? (short)6 : (short)2;
            if (!isInlet && !isOutlet)
            {
                colorIndex = 6;
            }

            return (colorIndex, 8.0, 2.5);
        }

        /// <summary>
        /// 创建属性定义
        /// </summary>
        /// <param name="defs">属性定义列表</param>
        /// <param name="midPointWorld">中点位置（世界坐标系）</param>
        /// <param name="finalRotation">最终旋转角度</param>
        /// <param name="pipelineLength">管道长度</param>
        /// <param name="titleFallback">管道标题后备值</param>
        /// <returns>属性定义列表</returns>
        private List<AttributeDefinition> CloneAttributeDefinitionsLocal(List<AttributeDefinition> defs, Point3d midPointWorld, double finalRotation, double pipelineLength, string titleFallback)
        {
            var result = new List<AttributeDefinition>();
            bool hasTitle = false;

            foreach (var def in defs)
            {
                var cloned = def.Clone() as AttributeDefinition;
                if (cloned == null) continue;

                // 转为局部坐标（相对中点）
                var localPos = new Point3d(def.Position.X - midPointWorld.X, def.Position.Y - midPointWorld.Y, 0);
                cloned.Position = localPos;
                cloned.Rotation = def.Rotation;
                cloned.Invisible = def.Invisible;
                cloned.Constant = def.Constant;
                cloned.Tag = def.Tag;
                cloned.TextString = def.TextString;
                cloned.Height = def.Height;

                if (!string.IsNullOrWhiteSpace(cloned.Tag))
                {
                    var tagLower = cloned.Tag.ToLowerInvariant();
                    if (tagLower.Contains("长度") || tagLower.Contains("length"))
                    {
                        double baseValue = 0.0;
                        if (double.TryParse(cloned.TextString, out double parsed)) baseValue = parsed;
                        cloned.TextString = (baseValue + pipelineLength).ToString("0.###");
                    }
                    if (string.Equals(cloned.Tag, "管道标题", StringComparison.OrdinalIgnoreCase))
                    {
                        hasTitle = true;
                        cloned.Position = Point3d.Origin;
                        cloned.Rotation = finalRotation;
                        cloned.Invisible = false;
                        if (string.IsNullOrWhiteSpace(cloned.TextString))
                            cloned.TextString = titleFallback ?? "管道";
                    }
                }

                result.Add(cloned);
            }

            if (!hasTitle)
            {
                result.Add(new AttributeDefinition
                {
                    Tag = "管道标题",
                    Position = Point3d.Origin,
                    Rotation = finalRotation,
                    TextString = string.IsNullOrWhiteSpace(titleFallback) ? "管道" : titleFallback,
                    Height = defs != null && defs.Count > 0 ? defs[0].Height : 2.5,
                    Invisible = false,
                    Constant = false
                });
            }

            return result;
        }

        /// <summary>
        /// 插入管道块
        /// </summary>
        /// <param name="tr">数据库事务</param>
        /// <param name="insertPointWorld">插入点（世界坐标）</param>
        /// <param name="blockName">块名称</param>
        /// <param name="rotation">旋转角度</param>
        /// <param name="attValues">属性值字典</param>
        /// <returns>新插入块的对象ID</returns>
        private ObjectId InsertPipeBlockWithAttributes(DBTrans tr, Point3d insertPointWorld, string blockName, double rotation, Dictionary<string, string> attValues)
        {
            ObjectId btrId = tr.BlockTable[blockName];
            ObjectId newBrId = tr.CurrentSpace.InsertBlock(insertPointWorld, btrId, rotation: rotation, atts: attValues);
            return newBrId;
        }

        /// <summary>
        /// 新增：根据首尾相连的线段集合，按连通顺序构建连续顶点列表（起点、每个连接点、终点）
        /// </summary>
        /// <param name="segments">线段集合</param>
        /// <param name="tol">容差</param>
        /// <returns></returns>
        private List<Point3d> BuildOrderedVerticesFromSegments(List<LineSegmentInfo> segments, double tol = 1e-6)
        {
            var result = new List<Point3d>();// 结果顶点列表
            if (segments == null || segments.Count == 0) return result;

            // 比较两点是否相等（使用容差）
            static bool PointsEqual(Point3d a, Point3d b, double tol)
            {
                return Math.Abs(a.X - b.X) <= tol && Math.Abs(a.Y - b.Y) <= tol && Math.Abs(a.Z - b.Z) <= tol;
            }
            // 构建唯一点列表并统计度数（出现次数）
            var uniquePoints = new List<Point3d>();
            Func<Point3d, int> getIndex = p =>
            {
                for (int i = 0; i < uniquePoints.Count; i++)
                {
                    if (PointsEqual(uniquePoints[i], p, tol)) return i;
                }
                uniquePoints.Add(p);
                return uniquePoints.Count - 1;
            };
            // 构建索引列表
            var counts = new List<int>();
            var segPairs = new List<(int s, int e)>();
            foreach (var seg in segments)
            {
                var si = getIndex(seg.StartPoint);
                var ei = getIndex(seg.EndPoint);
                segPairs.Add((si, ei));

                // ensure counts capacity
                while (counts.Count < uniquePoints.Count) counts.Add(0);
                counts[si]++;
                counts[ei]++;
            }
            // 找到链的端点：度为1的点（非闭合链）
            int startPointIndex = -1;
            for (int i = 0; i < counts.Count; i++)
            {
                if (counts[i] == 1)
                {
                    startPointIndex = i;
                    break;
                }
            }
            // 若都是度 >=2（闭合回路或多分支），退回到第一个段的起点
            if (startPointIndex == -1)
            {
                startPointIndex = segPairs.Count > 0 ? segPairs[0].s : 0;
            }
            // 从 startPointIndex 开始按链遍历段
            var visited = new bool[segPairs.Count];
            Point3d current = uniquePoints[startPointIndex];
            result.Add(current);
            bool progressed;
            do
            {
                progressed = false;
                for (int i = 0; i < segPairs.Count; i++)
                {
                    if (visited[i]) continue;
                    var (si, ei) = segPairs[i];
                    if (PointsEqual(uniquePoints[si], current, tol))
                    {
                        // forward
                        var next = uniquePoints[ei];
                        if (!PointsEqual(next, result.Last(), tol))
                            result.Add(next);
                        current = next;
                        visited[i] = true;
                        progressed = true;
                        break;
                    }
                    else if (PointsEqual(uniquePoints[ei], current, tol))
                    {
                        // reverse
                        var next = uniquePoints[si];
                        if (!PointsEqual(next, result.Last(), tol))
                            result.Add(next);
                        current = next;
                        visited[i] = true;
                        progressed = true;
                        break;
                    }
                }
            } while (progressed);

            // 新增校验：确保最终的方向与线段聚合方向一致
            try
            {
                if (result.Count >= 2)
                {
                    var overallVec = result.Last() - result.First();
                    if (!overallVec.IsZeroLength())
                    {
                        var agg = ComputeAggregateSegmentDirection(segments);
                        if (!agg.IsZeroLength())
                        {
                            // 如果总体向量与聚合向量点积为负，则反转顶点顺序
                            if (overallVec.DotProduct(agg) < 0)
                            {
                                result.Reverse();
                            }
                        }
                    }
                }
            }
            catch
            {
                // 容错：若聚合计算失败，不影响已有顺序
            }
            return result;
        }

        /// <summary>
        /// 分析示例块，提取管道、箭头和属性信息
        /// </summary>
        private SamplePipeInfo AnalyzeSampleBlock(Transaction tr, BlockReference blockRef)
        {
            //获取块的属性 初始化结果对象
            var info = new SamplePipeInfo();
            // 打开块表记录
            var btr = (BlockTableRecord)tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead);
            info.BasePoint = btr.Origin;//获取块的基点
            //遍历块中的对象 遍历块中的所有实体
            List<Polyline> polylines = new List<Polyline>();

            foreach (ObjectId id in btr)
            {
                //获取对象
                var dbObj = tr.GetObject(id, OpenMode.ForRead);
                if (dbObj is Polyline pl)
                {
                    //创建副本收集所有多段线
                    polylines.Add(pl.Clone() as Polyline);
                }
                else if (dbObj is AttributeDefinition attDef)//属性定义获取属性定义
                {
                    //创建副本添加属性定义到结果对象 添加属性定义到结果对象的属性定义列表
                    info.AttributeDefinitions.Add(attDef.Clone() as AttributeDefinition);
                }
            }
            //创建结果对象 分析多段线以识别管道主体和方向箭头
            if (polylines.Count == 0) return info;

            // 假设最长的Polyline是管道主体
            polylines = polylines.OrderByDescending(p => p.Length).ToList();
            info.PipeBodyTemplate = polylines[0];//设置管道主体模板设置管道主体模板为最长的多段线

            // 假设闭合的、有3个顶点的Polyline是方向箭头
            info.DirectionArrowTemplate = polylines.FirstOrDefault(p => p.Closed && p.NumberOfVertices == 3);

            if (info.DirectionArrowTemplate != null)
            {
                //获取箭头尖端的点 将箭头移动到原点，便于后续变换
                Point3d arrowTip = info.DirectionArrowTemplate.GetPoint3dAt(0);//获取箭头尖端的点假设第一个顶点是箭头尖端
                //创建一个矩阵，将箭头移动到原点创建变换矩阵将箭头移动到原点 创建变换矩阵将箭头移动到原点
                Matrix3d toOrigin = Matrix3d.Displacement(Point3d.Origin - arrowTip);
                //将箭头移动到原点应用变换 将变换应用到箭头模板上
                info.DirectionArrowTemplate.TransformBy(toOrigin);
            }

            return info;
        }

        /// <summary>
        /// 从选择的ObjectId集合中收集所有线段信息
        /// </summary>
        private List<LineSegmentInfo> CollectLineSegments(Transaction tr, List<ObjectId> ids)
        {
            var segments = new List<LineSegmentInfo>();
            foreach (var id in ids)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead);
                if (ent is Line line)
                {
                    segments.Add(ProcessLine(line, tr));
                }
                else if (ent is Polyline pl)
                {
                    for (int i = 0; i < pl.NumberOfVertices - 1; i++)
                    {
                        if (pl.GetSegmentType(i) == SegmentType.Line)
                        {
                            var p1 = pl.GetPoint3dAt(i);
                            var p2 = pl.GetPoint3dAt(i + 1);
                            var vec = p2 - p1;
                            segments.Add(new LineSegmentInfo
                            {
                                StartPoint = p1,
                                EndPoint = p2,
                                Length = vec.Length,
                                Angle = vec.GetAngleTo(Vector3d.XAxis, Vector3d.ZAxis),
                                Layer = pl.Layer,
                                ColorIndex = pl.ColorIndex,
                                LinetypeScale = pl.LinetypeScale,
                                EntityType = "POLYLINE_SEGMENT"
                            });
                        }
                    }
                }
            }
            return segments;
        }

        #endregion

        #region 绘制管道线

        /// <summary>
        /// 新增：通过点击采集点并生成管道块（入口/出口两种命令）DrawOutletPipeByClicks
        /// </summary>
        [CommandMethod("DrawOutletPipeByClicks")]
        public void DrawOutletPipeByClicks()
        {
            DrawPipeByClicks(isOutlet: true);
        }

        /// <summary>
        /// 新增：通过点击采集点并生成管道块（入口/出口两种命令）DrawInletPipeByClicks
        /// </summary>
        [CommandMethod("DrawInletPipeByClicks")]
        public void DrawInletPipeByClicks()
        {
            DrawPipeByClicks(isOutlet: false);
        }

        /// <summary>
        /// 主实现：交互式采点并基于示例块生成管道块
        /// 说明：Polyline pipeLocal = BuildPipePolylineLocal
        ///— 1) 首先要求用户选择一个示例管道块（作为模板/属性来源）；
        ///— 2) 交互式点击多个点（每次点击会记录坐标），按 Esc 或在提示阶段取消结束采点；
        ///— 3) 采集结束后用示例块生成管道块（复用现有的 Clone/Build/Insert 逻辑）。
        /// </summary>
        //private void DrawPipeByClicks(bool isOutlet)
        //{
        //    var doc = Application.DocumentManager.MdiActiveDocument;
        //    if (doc == null) return;
        //    var ed = doc.Editor;
        //    Database db = doc.Database;
        //    try
        //    {
        //        // 先尝试在块表中寻找示例块（优先按名称包含“出口/出口/outlet/入口/inlet”）
        //        ObjectId sampleBtrId = ObjectId.Null;
        //        string sampleBtrName = string.Empty;
        //        using (var tx = db.TransactionManager.StartTransaction())
        //        {
        //            // 遍历块表
        //            var bt = (BlockTable)tx.GetObject(db.BlockTableId, OpenMode.ForRead);
        //            foreach (ObjectId btrId in bt)// 遍历块表
        //            {
        //                try
        //                {
        //                    var btr = tx.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
        //                    if (btr == null) continue;
        //                    var nm = (btr.Name ?? string.Empty).ToLowerInvariant();
        //                    if (isOutlet)
        //                    {
        //                        if (nm.Contains("出口") || nm.Contains("outlet"))
        //                        {
        //                            sampleBtrId = btrId;
        //                            sampleBtrName = btr.Name;
        //                            break;
        //                        }
        //                    }
        //                    else
        //                    {
        //                        if (nm.Contains("入口") || nm.Contains("inlet"))
        //                        {
        //                            sampleBtrId = btrId;
        //                            sampleBtrName = btr.Name;
        //                            break;
        //                        }
        //                    }
        //                }
        //                catch { /* 忽略单条读取错误 */ }
        //            }
        //            tx.Commit();
        //        }
        //        // 没有找到合适的样例块
        //        bool tempInserted = false;
        //        // 样例块引用
        //        BlockReference sampleBr = null;
        //        // 开始事务
        //        using (var tr = new DBTrans())
        //        {
        //            try
        //            {
        //                // 如果在块表里找到了合适的样例块定义，先把它插入到当前空间（临时），位置用图纸原点
        //                if (!sampleBtrId.IsNull)
        //                {
        //                    Point3d insertPoint = Point3d.Origin;
        //                    // 缩放比例
        //                    var scale = VariableDictionary.textBoxScale;

        //                    var scale3d = new Scale3d(scale, scale, scale);

        //                    // 插入到当前空间
        //                    ObjectId tempSampleBrId = tr.CurrentSpace.InsertBlock(insertPoint, sampleBtrId, scale3d, rotation: 0.0, atts: null);
        //                    sampleBr = tr.GetObject(tempSampleBrId, OpenMode.ForWrite) as BlockReference;
        //                    tempInserted = true;
        //                }
        //                else
        //                {
        //                    // 未找到自动样例块，回退到用户选择示例块（原有流程）
        //                    ed.WriteMessage("\n未在块表中找到自动样例块，请手动选择示例块作为模板。");
        //                    var peo = new PromptEntityOptions("\n请选择示例管线块（作为模板，用于属性与样式）:");
        //                    peo.SetRejectMessage("\n请选择块参照对象.");
        //                    peo.AddAllowedClass(typeof(BlockReference), true);
        //                    var per = ed.GetEntity(peo);
        //                    if (per.Status != PromptStatus.OK)
        //                    {
        //                        ed.WriteMessage("\n未选择示例块，取消操作。");
        //                        tr.Abort();
        //                        return;
        //                    }
        //                    sampleBr = tr.GetObject(per.ObjectId, OpenMode.ForRead) as BlockReference;
        //                    // ensure writable for attribute write later
        //                    if (sampleBr != null && !sampleBr.IsWriteEnabled) sampleBr.UpgradeOpen();
        //                    tempInserted = false;
        //                }

        //                if (sampleBr == null)
        //                {
        //                    ed.WriteMessage("\n无法获取示例块参照，取消操作。");
        //                    tr.Abort();
        //                    return;
        //                }

        //                // 读取示例块属性并弹出属性编辑窗体（在插入后立即编辑）
        //                var sampleAttrMap = GetEntityAttributeMap(tr, sampleBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //                // 加载上次保存的属性（入口/出口区分）
        //                var loadedAttrs = FileManager.LoadLastPipeAttributes(isOutlet);

        //                // 弹出属性编辑窗，传入合并结果
        //                using (var editor = new PipeAttributeEditorForm(loadedAttrs))
        //                {
        //                    var dr = editor.ShowDialog();
        //                    if (dr != DialogResult.OK)
        //                    {
        //                        // 用户取消属性编辑，移除临时插入的样例块并退出
        //                        if (tempInserted)
        //                        {
        //                            try
        //                            {
        //                                var brToDel = tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference;
        //                                brToDel?.Erase();
        //                            }
        //                            catch { }
        //                        }
        //                        tr.Abort();
        //                        ed.WriteMessage("\n已取消属性编辑，操作终止。");
        //                        return;
        //                    }

        //                    // 用户确认后获取编辑结果并保存为默认供下次使用
        //                    var editedAttrs = editor.Attributes ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //                    FileManager.SaveLastPipeAttributes(isOutlet, editedAttrs);

        //                    // 把用户修改后的属性写回示例块（只写存在的 AttributeReference）
        //                    try
        //                    {
        //                        var sampleBrWrite = tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference;
        //                        if (sampleBrWrite != null)
        //                        {
        //                            foreach (ObjectId aid in sampleBrWrite.AttributeCollection)
        //                            {
        //                                try
        //                                {
        //                                    var ar = tr.GetObject(aid, OpenMode.ForWrite) as AttributeReference;
        //                                    if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue;
        //                                    if (editedAttrs.TryGetValue(ar.Tag, out var newVal))
        //                                    {
        //                                        ar.TextString = newVal ?? string.Empty;
        //                                        try { ar.AdjustAlignment(db); } catch { }
        //                                    }
        //                                }
        //                                catch { /* 单个属性写回失败不阻塞整体 */ }
        //                            }
        //                        }
        //                    }
        //                    catch { /* 忽略写回失败 */ }
        //                }
        //                // 属性编辑完成后，开始采点（起点支持右键/回车直接取消）
        //                var points = new List<Point3d>();
        //                var firstOpts = new PromptPointOptions("\n指定起点（点击或输入坐标，右键/回车取消）：");
        //                firstOpts.AllowNone = true; // 允许 None，这样右键/回车可作为取消

        //                // 可选关键字：输入“取消”也可终止
        //                firstOpts.Keywords.Add("取消");
        //                firstOpts.AppendKeywordsToMessage = true;

        //                var firstRes = ed.GetPoint(firstOpts);

        //                // 右键/回车（None）或关键字“取消”都走友好取消提示
        //                if (firstRes.Status == PromptStatus.None ||
        //                    (firstRes.Status == PromptStatus.Keyword &&
        //                     string.Equals(firstRes.StringResult, "取消", StringComparison.OrdinalIgnoreCase)))
        //                {
        //                    ed.WriteMessage("\n已取消绘制：未指定起点。");
        //                    if (tempInserted)
        //                    {
        //                        try
        //                        {
        //                            var brToDel = tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference;
        //                            brToDel?.Erase();
        //                        }
        //                        catch { }
        //                    }
        //                    tr.Abort();
        //                    return;
        //                }

        //                // ESC 或其它异常状态
        //                if (firstRes.Status != PromptStatus.OK)
        //                {
        //                    ed.WriteMessage("\n已取消绘制：起点输入已终止。");
        //                    if (tempInserted)
        //                    {
        //                        try
        //                        {
        //                            var brToDel = tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference;
        //                            brToDel?.Erase();
        //                        }
        //                        catch { }
        //                    }
        //                    tr.Abort();
        //                    return;
        //                }

        //                points.Add(firstRes.Value);

        //                while (true)
        //                {
        //                    // 允许“回车/空格/右键”结束采点，同时支持 ESC 取消输入并结束
        //                    var nextOpts = new PromptPointOptions("\n指定下一个点（点击或输入坐标），右键/回车结束：");
        //                    nextOpts.UseBasePoint = true;
        //                    nextOpts.BasePoint = points.Last();
        //                    nextOpts.AllowNone = true; // 关键：允许 None，右键通常会走到 None（等同回车）

        //                    // 可选：增加显式关键字，用户也可输入“完成”
        //                    nextOpts.Keywords.Add("完成");
        //                    nextOpts.AppendKeywordsToMessage = true;

        //                    var nextRes = ed.GetPoint(nextOpts);

        //                    if (nextRes.Status == PromptStatus.OK)
        //                    {
        //                        var pt = nextRes.Value;
        //                        if (pt.IsEqualTo(points.Last()))
        //                            break;
        //                        points.Add(pt);
        //                        continue;
        //                    }

        //                    // None = 回车/空格/右键，作为“结束绘制”
        //                    if (nextRes.Status == PromptStatus.None)
        //                        break;

        //                    // 关键字“完成”也结束绘制
        //                    if (nextRes.Status == PromptStatus.Keyword &&
        //                        string.Equals(nextRes.StringResult, "完成", StringComparison.OrdinalIgnoreCase))
        //                        break;

        //                    // ESC（Cancel）或其他状态，统一结束当前采点
        //                    break;
        //                }

        //                if (points.Count < 2)
        //                {
        //                    ed.WriteMessage("\n采集点不足（至少需要两个点），取消生成。");
        //                    // 清理临时样例块
        //                    if (tempInserted)
        //                    {
        //                        try { var brToDel = tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference; brToDel?.Erase(); }
        //                        catch { }
        //                    }
        //                    tr.Abort();
        //                    return;
        //                }

        //                // 使用示例块（sampleBr）生成管道块（与以前逻辑一致）
        //                var sampleInfo = AnalyzeSampleBlock(tr, sampleBr);
        //                if (sampleInfo?.PipeBodyTemplate == null)
        //                {
        //                    ed.WriteMessage("\n示例块中未找到管道主体（PolyLine），无法生成管道。");
        //                    // 清理临时样例块
        //                    if (tempInserted)
        //                    {
        //                        try { var brToDel = tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference; brToDel?.Erase(); }
        //                        catch { }
        //                    }
        //                    tr.Abort();
        //                    return;
        //                }

        //                // 计算管道总长度与中点
        //                //double pipelineLength = 0.0;
        //                //for (int i = 0; i < points.Count - 1; i++)
        //                //    pipelineLength += points[i].DistanceTo(points[i + 1]);
        //                // 计算本次绘制管线长度（带兜底算法，避免异常点导致长度=0）
        //                double pipelineLength = ComputePipelineLengthByPoints(points);
        //                if (pipelineLength <= 0.0)
        //                {
        //                    ed.WriteMessage("\n警告：未能有效计算管线长度，长度将按 0 处理。");
        //                }
        //                // 计算中点位置和整体方向（用于确定块的旋转和属性布局）
        //                var (midPoint, midAngle) = ComputeMidPointAndAngle(points, pipelineLength);
        //                Vector3d targetDir = ComputeDirectionAtPoint(points, midPoint, 1e-6);

        //                Vector3d segmentDir = ComputeAggregateSegmentDirection(BuildOrderedLineSegmentsFromPoints(points));
        //                if (!segmentDir.IsZeroLength() && targetDir.DotProduct(segmentDir) < 0)
        //                    targetDir = -targetDir;
        //                Vector3d targetDirNormalized = targetDir.IsZeroLength() ? Vector3d.XAxis : targetDir.GetNormal();

        //                // 局部 polyline（以 midPoint 为基准）
        //                var pipeLocal = BuildPipePolylineLocal(sampleInfo.PipeBodyTemplate, points, midPoint);

        //                // 重新读取示例块属性（以保证使用用户已编辑的值）
        //                var latestSampleAttrs = GetEntityAttributeMap(tr, sampleBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //                string pipeTitle = latestSampleAttrs.TryGetValue("管道标题", out var sampleTitle) && !string.IsNullOrWhiteSpace(sampleTitle)
        //                                 ? sampleTitle
        //                                 : (sampleBr.IsDynamicBlock ? sampleBr.Name : sampleBr.Name) ?? "管道";

        //                // 为每段生成局部坐标下的箭头与标题（相对于 midPoint），仅使用分段箭头/文字
        //                var arrowEntities = CreateDirectionalArrowsAndTitles(tr, sampleInfo, points, midPoint, pipeTitle, sampleBr.Name);

        //                // ---------- 关键修正：属性定义严格以示例块 AttributeDefinitions 为准 ----------
        //                var attDefsLocal = CloneAttributeDefinitionsLocal(sampleInfo.AttributeDefinitions, midPoint, 0.0, pipelineLength, sampleBr.Name)
        //                                    ?? new List<AttributeDefinition>();

        //                if (sampleInfo.AttributeDefinitions != null && sampleInfo.AttributeDefinitions.Count > 0)
        //                {
        //                    // 示例有定义：按示例字段更新值（不随意新增无示例的字段）
        //                    var latestDict = new Dictionary<string, string>(latestSampleAttrs, StringComparer.OrdinalIgnoreCase);
        //                    foreach (var def in attDefsLocal)
        //                    {
        //                        if (string.IsNullOrWhiteSpace(def.Tag)) continue;
        //                        if (latestDict.TryGetValue(def.Tag, out var v))
        //                            def.TextString = v ?? string.Empty;
        //                        def.Invisible = false;
        //                        def.Constant = false;
        //                    }
        //                }
        //                else
        //                {
        //                    // 示例无定义：根据 latestSampleAttrs 动态创建属性定义（按 Key 排序）
        //                    attDefsLocal.Clear();
        //                    double attHeight = 2.5;
        //                    double yOffsetBase = -attHeight * 2.0;
        //                    int extraIdx = 0;
        //                    foreach (var kv in latestSampleAttrs.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
        //                    {
        //                        if (string.IsNullOrWhiteSpace(kv.Key)) continue;
        //                        attDefsLocal.Add(new AttributeDefinition
        //                        {
        //                            Tag = kv.Key,
        //                            Position = new Point3d(0, yOffsetBase - extraIdx * attHeight * 1.2, 0),
        //                            Rotation = 0.0,
        //                            TextString = kv.Value ?? string.Empty,
        //                            Height = attHeight,
        //                            Invisible = false,
        //                            Constant = false
        //                        });
        //                        extraIdx++;
        //                    }
        //                }

        //                // 确保始点/终点/管段号属性（若示例已包含则更新其值，否则新增）
        //                string startCoordStr = $"X={points.First().X:F3},Y={points.First().Y:F3}";
        //                string endCoordStr = $"X={points.Last().X:F3},Y={points.Last().Y:F3}";

        //                // 修正错误：在调用 SetOrAddAttr 时必须传入有效的 ref 变量及合适的偏移参数
        //                // 计算用于新增属性时的位置参数（基于 attDefsLocal 第一个定义或默认值）
        //                double finalAttHeight = attDefsLocal.Count > 0 ? attDefsLocal[0].Height : 2.5;
        //                double finalYOffsetBase = attDefsLocal.Count > 0 ? attDefsLocal[0].Position.Y - finalAttHeight * 1.2 : -finalAttHeight * 2.0;
        //                int extraIndex = 0; // 追加属性时的索引（ref 参数）

        //                // 正确调用：传入 ref extraIndex, finalYOffsetBase, finalAttHeight
        //                SetOrAddAttr(attDefsLocal, "始点", startCoordStr, ref extraIndex, finalYOffsetBase, finalAttHeight);
        //                SetOrAddAttr(attDefsLocal, "终点", endCoordStr, ref extraIndex, finalYOffsetBase, finalAttHeight);

        //                // 将本次绘制长度写回属性（优先更新已有长度字段，不存在则新增）
        //                SetOrAddLengthAttrs(attDefsLocal, pipelineLength, ref extraIndex, finalYOffsetBase, finalAttHeight);

        //                int nextSegNum = GetNextPipeSegmentNumber(db);
        //                string extractedPipeNo = string.Empty;
        //                if (latestSampleAttrs.TryGetValue("管道标题", out var titleFromSample) && !string.IsNullOrWhiteSpace(titleFromSample))
        //                    extractedPipeNo = ExtractPipeCodeFromTitle(titleFromSample);
        //                if (string.IsNullOrWhiteSpace(extractedPipeNo))
        //                {
        //                    if (latestSampleAttrs.TryGetValue("管段号", out var pn) && !string.IsNullOrWhiteSpace(pn))
        //                        extractedPipeNo = pn;
        //                }
        //                if (string.IsNullOrWhiteSpace(extractedPipeNo))
        //                    extractedPipeNo = nextSegNum.ToString("D4");

        //                // 如果 attDefsLocal 中没有管段号字段，按之前逻辑补入（此处保证示例定义为准，但仍要保证管段号存在）
        //                if (!attDefsLocal.Any(a => string.Equals(a.Tag, "管段号", StringComparison.OrdinalIgnoreCase)))
        //                {
        //                    double attH = attDefsLocal.Count > 0 ? attDefsLocal[0].Height : 2.5;
        //                    double yBase = attDefsLocal.Count > 0 ? attDefsLocal[0].Position.Y - attH * 1.2 : -attH * 2.0;
        //                    attDefsLocal.Add(new AttributeDefinition
        //                    {
        //                        Tag = "管段号",
        //                        Position = new Point3d(0, yBase - attDefsLocal.Count * attH * 1.2, 0),
        //                        Rotation = 0.0,
        //                        TextString = extractedPipeNo,
        //                        Height = attH,
        //                        Invisible = false,
        //                        Constant = false
        //                    });
        //                }
        //                else
        //                {
        //                    var pnDef = attDefsLocal.First(a => string.Equals(a.Tag, "管段号", StringComparison.OrdinalIgnoreCase));
        //                    pnDef.TextString = extractedPipeNo;
        //                }

        //                // 在构建块定义之前：移除/隐藏中点“管道标题”
        //                //attDefsLocal.RemoveAll(ad => string.Equals(ad.Tag, "管道标题", StringComparison.OrdinalIgnoreCase));
        //                foreach (var ad in attDefsLocal)
        //                {
        //                    ad.Invisible = true;
        //                    ad.Constant = false;
        //                }

        //                // 构建块定义并插入新块
        //                string desiredName = sampleBr.Name;
        //                string newBlockName = BuildPipeBlockDefinition(tr, desiredName, (Polyline)pipeLocal.Clone(), arrowEntities, attDefsLocal);

        //                var attValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //                foreach (var a in attDefsLocal)
        //                {
        //                    if (string.IsNullOrWhiteSpace(a?.Tag)) continue;
        //                    attValues[a.Tag] = a.TextString ?? string.Empty;
        //                }

        //                var newBrId = InsertPipeBlockWithAttributes(tr, midPoint, newBlockName, 0.0, attValues);

        //                PipelineTopologyHelper.PostProcessAfterPipePlaced(tr, newBrId, isOutlet ? "Outlet" : "Inlet");
        //                var newBr = tr.GetObject(newBrId, OpenMode.ForWrite) as BlockReference;
        //                if (newBr != null)
        //                    newBr.Layer = sampleInfo.PipeBodyTemplate.Layer;

        //                // 清理：如果之前我们临时插入了样例块，删除它（不影响已创建的管线块）foreach (var ad in attDefsLocal)
        //                if (tempInserted)
        //                {
        //                    try
        //                    {
        //                        var brToDel = tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference;
        //                        brToDel?.Erase();
        //                    }
        //                    catch { /* 忽略删除失败 */ }
        //                }

        //                tr.Commit();
        //                ed.WriteMessage($"\n管道已生成（{(isOutlet ? "出口" : "入口")}），点数: {points.Count}，管段号: {extractedPipeNo}");
        //            }
        //            catch (Exception ex)
        //            {
        //                tr.Abort();
        //                ed.WriteMessage($"\n生成管道时出错: {ex.Message}");
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ed.WriteMessage($"\n操作失败: {ex.Message}");
        //    }
        //}

        /// <summary>
        /// 主实现：交互式采点并基于示例块生成管道块（增强：起终点自动拾取最近部件并合并同名属性）
        /// </summary>
        //private void DrawPipeByClicks(bool isOutlet)
        //{
        //    var doc = Application.DocumentManager.MdiActiveDocument;
        //    if (doc == null) return;
        //    var ed = doc.Editor;
        //    Database db = doc.Database;

        //    try
        //    {
        //        // 先在块表中尝试自动寻找示例块（入口/出口）
        //        ObjectId sampleBtrId = ObjectId.Null;
        //        using (var tx = db.TransactionManager.StartTransaction())
        //        {
        //            var bt = (BlockTable)tx.GetObject(db.BlockTableId, OpenMode.ForRead);
        //            foreach (ObjectId btrId in bt)
        //            {
        //                try
        //                {
        //                    var btr = tx.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
        //                    if (btr == null) continue;
        //                    var nm = (btr.Name ?? string.Empty).ToLowerInvariant();
        //                    if (isOutlet)
        //                    {
        //                        if (nm.Contains("出口") || nm.Contains("outlet"))
        //                        {
        //                            sampleBtrId = btrId;
        //                            break;
        //                        }
        //                    }
        //                    else
        //                    {
        //                        if (nm.Contains("入口") || nm.Contains("inlet"))
        //                        {
        //                            sampleBtrId = btrId;
        //                            break;
        //                        }
        //                    }
        //                }
        //                catch { }
        //            }
        //            tx.Commit();
        //        }

        //        bool tempInserted = false;
        //        BlockReference sampleBr = null;

        //        using (var tr = new DBTrans())
        //        {
        //            try
        //            {
        //                // 若自动找到样例块，则临时插入一份作为模板
        //                if (!sampleBtrId.IsNull)
        //                {
        //                    var scale = VariableDictionary.textBoxScale;
        //                    if (double.IsNaN(scale) || scale <= 0) scale = 1.0;
        //                    ObjectId tempSampleBrId = tr.CurrentSpace.InsertBlock(Point3d.Origin, sampleBtrId, new Scale3d(scale, scale, scale), 0.0, null);
        //                    sampleBr = tr.GetObject(tempSampleBrId, OpenMode.ForWrite) as BlockReference;
        //                    tempInserted = true;
        //                }
        //                else
        //                {
        //                    // 未自动匹配时，手工选择示例块
        //                    ed.WriteMessage("\n未自动找到示例块，请手动选择一个块参照作为模板。");
        //                    var peo = new PromptEntityOptions("\n请选择示例管线块：");
        //                    peo.SetRejectMessage("\n请选择块参照对象。");
        //                    peo.AddAllowedClass(typeof(BlockReference), true);
        //                    var per = ed.GetEntity(peo);
        //                    if (per.Status != PromptStatus.OK)
        //                    {
        //                        ed.WriteMessage("\n未选择示例块，操作取消。");
        //                        tr.Abort();
        //                        return;
        //                    }
        //                    sampleBr = tr.GetObject(per.ObjectId, OpenMode.ForRead) as BlockReference;
        //                    if (sampleBr != null && !sampleBr.IsWriteEnabled) sampleBr.UpgradeOpen();
        //                    tempInserted = false;
        //                }

        //                if (sampleBr == null)
        //                {
        //                    ed.WriteMessage("\n无法获取示例块，取消。");
        //                    tr.Abort();
        //                    return;
        //                }

        //                // 弹出属性编辑（沿用你原有逻辑）
        //                var loadedAttrs = FileManager.LoadLastPipeAttributes(isOutlet);
        //                using (var editor = new PipeAttributeEditorForm(loadedAttrs))
        //                {
        //                    var dr = editor.ShowDialog();
        //                    if (dr != DialogResult.OK)
        //                    {
        //                        if (tempInserted)
        //                        {
        //                            try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { }
        //                        }
        //                        tr.Abort();
        //                        ed.WriteMessage("\n已取消属性编辑，终止。");
        //                        return;
        //                    }

        //                    var editedAttrs = editor.Attributes ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //                    FileManager.SaveLastPipeAttributes(isOutlet, editedAttrs);

        //                    try
        //                    {
        //                        var sampleBrWrite = tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference;
        //                        if (sampleBrWrite != null)
        //                        {
        //                            foreach (ObjectId aid in sampleBrWrite.AttributeCollection)
        //                            {
        //                                var ar = tr.GetObject(aid, OpenMode.ForWrite) as AttributeReference;
        //                                if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue;
        //                                if (editedAttrs.TryGetValue(ar.Tag, out var newVal))
        //                                {
        //                                    ar.TextString = newVal ?? string.Empty;
        //                                    try { ar.AdjustAlignment(db); } catch { }
        //                                }
        //                            }
        //                        }
        //                    }
        //                    catch { }
        //                }

        //                // 采集起点
        //                var points = new List<Point3d>();
        //                var firstOpts = new PromptPointOptions("\n指定起点（右键/回车取消）：");
        //                firstOpts.AllowNone = true;
        //                firstOpts.Keywords.Add("取消");
        //                firstOpts.AppendKeywordsToMessage = true;
        //                var firstRes = ed.GetPoint(firstOpts);

        //                if (firstRes.Status == PromptStatus.None ||
        //                    (firstRes.Status == PromptStatus.Keyword && string.Equals(firstRes.StringResult, "取消", StringComparison.OrdinalIgnoreCase)))
        //                {
        //                    ed.WriteMessage("\n已取消：未指定起点。");
        //                    if (tempInserted) { try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { } }
        //                    tr.Abort();
        //                    return;
        //                }
        //                if (firstRes.Status != PromptStatus.OK)
        //                {
        //                    ed.WriteMessage("\n已取消：起点输入终止。");
        //                    if (tempInserted) { try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { } }
        //                    tr.Abort();
        //                    return;
        //                }
        //                points.Add(firstRes.Value);

        //                // 循环采集后续点
        //                while (true)
        //                {
        //                    var nextOpts = new PromptPointOptions("\n指定下一个点（右键/回车结束）：");
        //                    nextOpts.UseBasePoint = true;
        //                    nextOpts.BasePoint = points.Last();
        //                    nextOpts.AllowNone = true;
        //                    nextOpts.Keywords.Add("完成");
        //                    nextOpts.AppendKeywordsToMessage = true;

        //                    var nextRes = ed.GetPoint(nextOpts);
        //                    if (nextRes.Status == PromptStatus.OK)
        //                    {
        //                        var pt = nextRes.Value;
        //                        if (pt.IsEqualTo(points.Last())) break;
        //                        points.Add(pt);
        //                        continue;
        //                    }
        //                    if (nextRes.Status == PromptStatus.None) break;
        //                    if (nextRes.Status == PromptStatus.Keyword && string.Equals(nextRes.StringResult, "完成", StringComparison.OrdinalIgnoreCase)) break;
        //                    break;
        //                }

        //                if (points.Count < 2)
        //                {
        //                    ed.WriteMessage("\n采点不足，取消生成。");
        //                    if (tempInserted) { try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { } }
        //                    tr.Abort();
        //                    return;
        //                }

        //                // 分析样例块
        //                var sampleInfo = AnalyzeSampleBlock(tr, sampleBr);
        //                if (sampleInfo?.PipeBodyTemplate == null)
        //                {
        //                    ed.WriteMessage("\n示例块中未找到管道主体。");
        //                    if (tempInserted) { try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { } }
        //                    tr.Abort();
        //                    return;
        //                }

        //                // 计算长度与中点
        //                double pipelineLength = ComputePipelineLengthByPoints(points);
        //                var (midPoint, _) = ComputeMidPointAndAngle(points, pipelineLength);

        //                // 构建局部管线
        //                var pipeLocal = BuildPipePolylineLocal(sampleInfo.PipeBodyTemplate, points, midPoint);

        //                // 读取样例块属性（基线属性）
        //                var latestSampleAttrs = GetEntityAttributeMap(tr, sampleBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //                #region  ======================== 增强开始：起终点自动拾取并合并属性 ========================
        //                //// 根据比例计算搜索半径（你可按项目再调）
        //                //double pickRadius = Math.Max(200.0, VariableDictionary.textBoxScale * 5.0);

        //                //// 查找起点附近最近部件
        //                //var startBr = FindNearestBlockReferenceNearPoint(tr, points.First(), pickRadius, sampleBr.ObjectId);

        //                //// 查找终点附近最近部件
        //                //var endBr = FindNearestBlockReferenceNearPoint(tr, points.Last(), pickRadius, sampleBr.ObjectId);

        //                //// 读取起点部件属性
        //                //var startMap = startBr != null ? (GetEntityAttributeMap(tr, startBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
        //                //                               : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //                //// 读取终点部件属性
        //                //var endMap = endBr != null ? (GetEntityAttributeMap(tr, endBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
        //                //                           : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //                //// 统计命中数量
        //                //int startHit = 0;
        //                //int endHit = 0;

        //                //// 起点属性优先覆盖同名字段（仅白名单）
        //                //foreach (var kv in startMap)
        //                //{
        //                //    string key = (kv.Key ?? string.Empty).Trim();
        //                //    if (string.IsNullOrWhiteSpace(key)) continue;
        //                //    if (!IsPipeEndpointAttrAllowed(key)) continue;

        //                //    if (latestSampleAttrs.ContainsKey(key))
        //                //    {
        //                //        latestSampleAttrs[key] = kv.Value ?? string.Empty;
        //                //        startHit++;
        //                //    }
        //                //}

        //                //// 终点属性补充覆盖同名字段（仅白名单）
        //                //foreach (var kv in endMap)
        //                //{
        //                //    string key = (kv.Key ?? string.Empty).Trim();
        //                //    if (string.IsNullOrWhiteSpace(key)) continue;
        //                //    if (!IsPipeEndpointAttrAllowed(key)) continue;

        //                //    if (latestSampleAttrs.ContainsKey(key))
        //                //    {
        //                //        latestSampleAttrs[key] = kv.Value ?? string.Empty;
        //                //        endHit++;
        //                //    }
        //                //}

        //                //// 把起终点关键提示字段写入（可用于后续检查）
        //                //latestSampleAttrs["始点"] = $"X={points.First().X:F3},Y={points.First().Y:F3}";
        //                //latestSampleAttrs["终点"] = $"X={points.Last().X:F3},Y={points.Last().Y:F3}";

        //                //// 输出命令行日志，明确是否真的拾取到了部件
        //                //ed.WriteMessage($"\n端点拾取半径={pickRadius:F2}；起点部件={(startBr != null ? startBr.Name : "未命中")}；终点部件={(endBr != null ? endBr.Name : "未命中")}。");
        //                //ed.WriteMessage($"\n端点同名属性合并：起点命中={startHit}，终点命中={endHit}。");
        //                #endregion ======================== 增强结束 ========================

        //                // ======================== 增强开始：起终点自动拾取并合并属性（含名称/管道标题/管段号） ========================
        //                // 根据比例计算搜索半径（可按项目调参）
        //                double pickRadius = Math.Max(200.0, VariableDictionary.textBoxScale * 5.0);

        //                // 查找起点/终点附近最近部件（块参照）
        //                var startBr = FindNearestBlockReferenceNearPoint(tr, points.First(), pickRadius, sampleBr.ObjectId);
        //                var endBr = FindNearestBlockReferenceNearPoint(tr, points.Last(), pickRadius, sampleBr.ObjectId);

        //                // 读取端点部件属性映射
        //                var startMap = startBr != null
        //                    ? (GetEntityAttributeMap(tr, startBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
        //                    : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //                var endMap = endBr != null
        //                    ? (GetEntityAttributeMap(tr, endBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
        //                    : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //                // 候选键集合（按你现场字段继续补）
        //                string[] nameKeys = new[] { "名称", "设备名称", "Name", "TAG", "Tag", "位号", "设备位号" };
        //                string[] titleKeys = new[] { "管道标题", "标题", "Title" };
        //                string[] pipeNoKeys = new[] { "管段号", "管段编号", "Pipeline No", "Pipe No", "No" };

        //                // 本地函数：从属性字典里按候选键取第一个非空值（支持包含匹配）
        //                string FindFirstAttrValueLocal(Dictionary<string, string> attrs, string[] keys)
        //                {
        //                    if (attrs == null || attrs.Count == 0) return string.Empty;

        //                    // 先精确键匹配
        //                    foreach (var k in keys)
        //                    {
        //                        if (attrs.TryGetValue(k, out var v) && !string.IsNullOrWhiteSpace(v))
        //                            return v.Trim();
        //                    }

        //                    // 再做包含匹配（兼容“起点名称”“设备名称1”等）
        //                    foreach (var kv in attrs)
        //                    {
        //                        if (string.IsNullOrWhiteSpace(kv.Key) || string.IsNullOrWhiteSpace(kv.Value)) continue;
        //                        foreach (var k in keys)
        //                        {
        //                            if (kv.Key.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
        //                                return kv.Value.Trim();
        //                        }
        //                    }

        //                    return string.Empty;
        //                }

        //                // 从起点/终点部件提取“名称”
        //                string startName = FindFirstAttrValueLocal(startMap, nameKeys);
        //                string endName = FindFirstAttrValueLocal(endMap, nameKeys);

        //                // 从端点或样例提取“管道标题/管段号”
        //                string startTitle = FindFirstAttrValueLocal(startMap, titleKeys);
        //                string endTitle = FindFirstAttrValueLocal(endMap, titleKeys);
        //                string sampleTitle = latestSampleAttrs.TryGetValue("管道标题", out var st) ? (st ?? string.Empty).Trim() : string.Empty;
        //                string finalTitle = !string.IsNullOrWhiteSpace(startTitle) ? startTitle
        //                                : !string.IsNullOrWhiteSpace(endTitle) ? endTitle
        //                                : sampleTitle;

        //                string startPipeNo = FindFirstAttrValueLocal(startMap, pipeNoKeys);
        //                string endPipeNo = FindFirstAttrValueLocal(endMap, pipeNoKeys);
        //                string samplePipeNo = latestSampleAttrs.TryGetValue("管段号", out var spn) ? (spn ?? string.Empty).Trim() : string.Empty;
        //                string finalPipeNo = !string.IsNullOrWhiteSpace(startPipeNo) ? startPipeNo
        //                                 : !string.IsNullOrWhiteSpace(endPipeNo) ? endPipeNo
        //                                 : samplePipeNo;

        //                // 始点/终点字段优先写“名称值”，若未取到则回退坐标
        //                string startCoordStr = $"X={points.First().X:F3},Y={points.First().Y:F3}";
        //                string endCoordStr = $"X={points.Last().X:F3},Y={points.Last().Y:F3}";

        //                latestSampleAttrs["始点"] = !string.IsNullOrWhiteSpace(startName) ? startName : startCoordStr;
        //                latestSampleAttrs["终点"] = !string.IsNullOrWhiteSpace(endName) ? endName : endCoordStr;

        //                // 把端点名称额外保留到独立字段（可选，但建议保留便于排查）
        //                if (!string.IsNullOrWhiteSpace(startName)) latestSampleAttrs["始点名称"] = startName;
        //                if (!string.IsNullOrWhiteSpace(endName)) latestSampleAttrs["终点名称"] = endName;

        //                // 管道标题与管段号按“端点优先、样例兜底”写回
        //                if (!string.IsNullOrWhiteSpace(finalTitle)) latestSampleAttrs["管道标题"] = finalTitle;
        //                if (!string.IsNullOrWhiteSpace(finalPipeNo)) latestSampleAttrs["管段号"] = finalPipeNo;

        //                // 命令行诊断输出
        //                ed.WriteMessage($"\n端点拾取半径={pickRadius:F2}；起点部件={(startBr != null ? startBr.Name : "未命中")}；终点部件={(endBr != null ? endBr.Name : "未命中")}。");
        //                ed.WriteMessage($"\n端点名称：始点='{latestSampleAttrs["始点"]}'，终点='{latestSampleAttrs["终点"]}'。");
        //                ed.WriteMessage($"\n标题/管段号：管道标题='{(latestSampleAttrs.ContainsKey("管道标题") ? latestSampleAttrs["管道标题"] : "")}'，管段号='{(latestSampleAttrs.ContainsKey("管段号") ? latestSampleAttrs["管段号"] : "")}'。");
        //                // ======================== 增强结束 ========================

        //                // 生成标题与箭头
        //                string pipeTitle = latestSampleAttrs.TryGetValue("管道标题", out var sample_Title) && !string.IsNullOrWhiteSpace(sample_Title)
        //                                 ? sample_Title
        //                                 : sampleBr.Name ?? "管道";
        //                var arrowEntities = CreateDirectionalArrowsAndTitles(tr, sampleInfo, points, midPoint, pipeTitle, sampleBr.Name);

        //                // 克隆属性定义并回填值
        //                var attDefsLocal = CloneAttributeDefinitionsLocal(sampleInfo.AttributeDefinitions, midPoint, 0.0, pipelineLength, sampleBr.Name)
        //                                   ?? new List<AttributeDefinition>();

        //                if (sampleInfo.AttributeDefinitions != null && sampleInfo.AttributeDefinitions.Count > 0)
        //                {
        //                    foreach (var def in attDefsLocal)
        //                    {
        //                        if (string.IsNullOrWhiteSpace(def.Tag)) continue;
        //                        if (latestSampleAttrs.TryGetValue(def.Tag, out var v))
        //                            def.TextString = v ?? string.Empty;
        //                        def.Invisible = false;
        //                        def.Constant = false;
        //                    }
        //                }
        //                else
        //                {
        //                    attDefsLocal.Clear();
        //                    double attHeight = 2.5;
        //                    double yOffsetBase = -attHeight * 2.0;
        //                    int idx = 0;
        //                    foreach (var kv in latestSampleAttrs.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
        //                    {
        //                        if (string.IsNullOrWhiteSpace(kv.Key)) continue;
        //                        attDefsLocal.Add(new AttributeDefinition
        //                        {
        //                            Tag = kv.Key,
        //                            Position = new Point3d(0, yOffsetBase - idx * attHeight * 1.2, 0),
        //                            Rotation = 0.0,
        //                            TextString = kv.Value ?? string.Empty,
        //                            Height = attHeight,
        //                            Invisible = false,
        //                            Constant = false
        //                        });
        //                        idx++;
        //                    }
        //                }

        //                // 确保关键属性存在
        //                int nextSegNum = GetNextPipeSegmentNumber(db);
        //                string extractedPipeNo = nextSegNum.ToString("D4");
        //                if (latestSampleAttrs.TryGetValue("管段号", out var pn) && !string.IsNullOrWhiteSpace(pn))
        //                    extractedPipeNo = pn;

        //                double finalAttHeight = attDefsLocal.Count > 0 ? attDefsLocal[0].Height : 2.5;
        //                double finalYOffsetBase = attDefsLocal.Count > 0 ? attDefsLocal[0].Position.Y - finalAttHeight * 1.2 : -finalAttHeight * 2.0;
        //                int extraIndex = 0;

        //                SetOrAddAttr(attDefsLocal, "始点", latestSampleAttrs["始点"], ref extraIndex, finalYOffsetBase, finalAttHeight);
        //                SetOrAddAttr(attDefsLocal, "终点", latestSampleAttrs["终点"], ref extraIndex, finalYOffsetBase, finalAttHeight);
        //                SetOrAddLengthAttrs(attDefsLocal, pipelineLength, ref extraIndex, finalYOffsetBase, finalAttHeight);
        //                SetOrAddAttr(attDefsLocal, "管段号", extractedPipeNo, ref extraIndex, finalYOffsetBase, finalAttHeight);

        //                // 属性设为隐藏（保持你原行为）
        //                foreach (var ad in attDefsLocal)
        //                {
        //                    ad.Invisible = true;
        //                    ad.Constant = false;
        //                }

        //                // 构建并插入新块
        //                string newBlockName = BuildPipeBlockDefinition(tr, sampleBr.Name, (Polyline)pipeLocal.Clone(), arrowEntities, attDefsLocal);
        //                var attValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //                foreach (var a in attDefsLocal)
        //                {
        //                    if (string.IsNullOrWhiteSpace(a?.Tag)) continue;
        //                    attValues[a.Tag] = a.TextString ?? string.Empty;
        //                }

        //                var newBrId = InsertPipeBlockWithAttributes(tr, midPoint, newBlockName, 0.0, attValues);
        //                PipelineTopologyHelper.PostProcessAfterPipePlaced(tr, newBrId, isOutlet ? "Outlet" : "Inlet");

        //                var newBr = tr.GetObject(newBrId, OpenMode.ForWrite) as BlockReference;
        //                if (newBr != null) newBr.Layer = sampleInfo.PipeBodyTemplate.Layer;

        //                // 清理临时样例块
        //                if (tempInserted)
        //                {
        //                    try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { }
        //                }

        //                tr.Commit();
        //                ed.WriteMessage($"\n管道已生成（{(isOutlet ? "出口" : "入口")}），点数={points.Count}，管段号={extractedPipeNo}。");
        //            }
        //            catch (Exception ex)
        //            {
        //                tr.Abort();
        //                ed.WriteMessage($"\n生成管道时出错: {ex.Message}");
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ed.WriteMessage($"\n操作失败: {ex.Message}");
        //    }
        //}

        /// <summary>
        /// 修整版：交互式采点并基于示例块生成管道块（入口/出口共用）
        /// 说明：
        /// 1）优先块表自动匹配入口/出口样例；
        /// 2）未命中时手工选择样例；
        /// 3）历史属性文件为空时，自动用样例块当前属性兜底，避免编辑页空表；
        /// 4）保留端点自动拾取名称/标题/管段号回填逻辑。
        /// </summary>
        //private void DrawPipeByClicks(bool isOutlet)
        //{
        //    // 获取当前活动文档
        //    var doc = Application.DocumentManager.MdiActiveDocument;
        //    // 无文档直接结束
        //    if (doc == null) return;

        //    // 获取编辑器
        //    var ed = doc.Editor;
        //    // 获取数据库
        //    Database db = doc.Database;

        //    try
        //    {
        //        // 先在块表中自动匹配“入口/出口”样例块定义
        //        ObjectId sampleBtrId = ObjectId.Null;
        //        using (var tx = db.TransactionManager.StartTransaction())
        //        {
        //            // 打开块表
        //            var bt = (BlockTable)tx.GetObject(db.BlockTableId, OpenMode.ForRead);
        //            // 遍历块定义
        //            foreach (ObjectId btrId in bt)
        //            {
        //                try
        //                {
        //                    // 读取块定义记录
        //                    var btr = tx.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
        //                    // 空记录跳过
        //                    if (btr == null) continue;

        //                    // 块名转小写便于匹配
        //                    var nm = (btr.Name ?? string.Empty).ToLowerInvariant();

        //                    // 按入口/出口关键字匹配
        //                    if (isOutlet)
        //                    {
        //                        // 出口模式匹配“出口/outlet”
        //                        if (nm.Contains("出口") || nm.Contains("outlet"))
        //                        {
        //                            sampleBtrId = btrId;
        //                            break;
        //                        }
        //                    }
        //                    else
        //                    {
        //                        // 入口模式匹配“入口/inlet”
        //                        if (nm.Contains("入口") || nm.Contains("inlet"))
        //                        {
        //                            sampleBtrId = btrId;
        //                            break;
        //                        }
        //                    }
        //                }
        //                catch
        //                {
        //                    // 单条读取异常忽略，继续扫描
        //                }
        //            }

        //            // 提交只读事务
        //            tx.Commit();
        //        }

        //        // 标记样例是否是临时插入
        //        bool tempInserted = false;
        //        // 样例块引用
        //        BlockReference sampleBr = null;

        //        // 开启绘制事务
        //        using (var tr = new DBTrans())
        //        {
        //            try
        //            {
        //                // 自动匹配命中时，临时插入样例块到原点
        //                if (!sampleBtrId.IsNull)
        //                {
        //                    // 读取比例
        //                    var scale = VariableDictionary.textBoxScale;
        //                    // 比例保护
        //                    if (double.IsNaN(scale) || scale <= 0) scale = 1.0;

        //                    // 插入临时样例块
        //                    ObjectId tempSampleBrId = tr.CurrentSpace.InsertBlock(Point3d.Origin, sampleBtrId, new Scale3d(scale, scale, scale), 0.0, null);
        //                    // 打开可写引用
        //                    sampleBr = tr.GetObject(tempSampleBrId, OpenMode.ForWrite) as BlockReference;
        //                    // 标记为临时插入
        //                    tempInserted = true;
        //                }
        //                else
        //                {
        //                    // 自动未命中时回退手工选择样例块
        //                    ed.WriteMessage("\n未自动找到示例块，请手动选择一个块参照作为模板。");
        //                    // 创建选择提示
        //                    var peo = new PromptEntityOptions("\n请选择示例管线块：");
        //                    // 设置拒绝消息
        //                    peo.SetRejectMessage("\n请选择块参照对象。");
        //                    // 限制只能选块参照
        //                    peo.AddAllowedClass(typeof(BlockReference), true);
        //                    // 执行选择
        //                    var per = ed.GetEntity(peo);

        //                    // 选择取消则终止
        //                    if (per.Status != PromptStatus.OK)
        //                    {
        //                        ed.WriteMessage("\n未选择示例块，操作取消。");
        //                        tr.Abort();
        //                        return;
        //                    }

        //                    // 读取用户选择的样例块
        //                    sampleBr = tr.GetObject(per.ObjectId, OpenMode.ForRead) as BlockReference;
        //                    // 升级为可写
        //                    if (sampleBr != null && !sampleBr.IsWriteEnabled) sampleBr.UpgradeOpen();
        //                    // 手选样例不是临时插入
        //                    tempInserted = false;
        //                }

        //                // 样例获取失败直接退出
        //                if (sampleBr == null)
        //                {
        //                    ed.WriteMessage("\n无法获取示例块，取消。");
        //                    tr.Abort();
        //                    return;
        //                }

        //                // 读取历史属性（入口/出口分开）
        //                var loadedAttrs = FileManager.LoadLastPipeAttributes(isOutlet) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //                // 兜底——如果历史属性为空，则用样例块当前属性填充，避免编辑页空表
        //                if (loadedAttrs.Count == 0)
        //                {
        //                    // 读取样例块属性字典
        //                    var sampleAttrMap = GetEntityAttributeMap(tr, sampleBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //                    // 逐项写入编辑初值
        //                    foreach (var kv in sampleAttrMap)
        //                    {
        //                        if (string.IsNullOrWhiteSpace(kv.Key)) continue;
        //                        loadedAttrs[kv.Key.Trim()] = kv.Value ?? string.Empty;
        //                    }

        //                    // 打印兜底日志
        //                    ed.WriteMessage($"\n已用样例块属性兜底初始化编辑项：{loadedAttrs.Count} 项。");
        //                }

        //                // 弹出属性编辑窗
        //                using (var editor = new PipeAttributeEditorForm(loadedAttrs))
        //                {
        //                    // 显示编辑窗
        //                    var dr = editor.ShowDialog();

        //                    // 用户取消则清理并退出
        //                    if (dr != DialogResult.OK)
        //                    {
        //                        if (tempInserted)
        //                        {
        //                            try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { }
        //                        }

        //                        tr.Abort();
        //                        ed.WriteMessage("\n已取消属性编辑，终止。");
        //                        return;
        //                    }

        //                    // 获取编辑结果
        //                    var editedAttrs = editor.Attributes ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //                    // 保存到历史属性文件
        //                    FileManager.SaveLastPipeAttributes(isOutlet, editedAttrs);

        //                    // 把编辑结果回写到样例块已有属性
        //                    try
        //                    {
        //                        // 打开样例块可写对象
        //                        var sampleBrWrite = tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference;
        //                        if (sampleBrWrite != null)
        //                        {
        //                            // 遍历属性引用
        //                            foreach (ObjectId aid in sampleBrWrite.AttributeCollection)
        //                            {
        //                                // 读取属性引用
        //                                var ar = tr.GetObject(aid, OpenMode.ForWrite) as AttributeReference;
        //                                // 无效属性跳过
        //                                if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue;

        //                                // 命中同名键则回写
        //                                if (editedAttrs.TryGetValue(ar.Tag, out var newVal))
        //                                {
        //                                    ar.TextString = newVal ?? string.Empty;
        //                                    try { ar.AdjustAlignment(db); } catch { }
        //                                }
        //                            }
        //                        }
        //                    }
        //                    catch
        //                    {
        //                        // 回写失败不阻塞主流程
        //                    }
        //                }

        //                // 采点列表
        //                var points = new List<Point3d>();

        //                // 首点提示
        //                var firstOpts = new PromptPointOptions("\n指定起点（右键/回车取消）：");
        //                firstOpts.AllowNone = true;
        //                firstOpts.Keywords.Add("取消");
        //                firstOpts.AppendKeywordsToMessage = true;

        //                // 读取首点
        //                var firstRes = ed.GetPoint(firstOpts);

        //                // 首点取消
        //                if (firstRes.Status == PromptStatus.None ||
        //                    (firstRes.Status == PromptStatus.Keyword && string.Equals(firstRes.StringResult, "取消", StringComparison.OrdinalIgnoreCase)))
        //                {
        //                    ed.WriteMessage("\n已取消：未指定起点。");
        //                    if (tempInserted) { try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { } }
        //                    tr.Abort();
        //                    return;
        //                }

        //                // 首点非正常状态
        //                if (firstRes.Status != PromptStatus.OK)
        //                {
        //                    ed.WriteMessage("\n已取消：起点输入终止。");
        //                    if (tempInserted) { try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { } }
        //                    tr.Abort();
        //                    return;
        //                }

        //                // 加入首点
        //                points.Add(firstRes.Value);

        //                // 循环采集后续点
        //                while (true)
        //                {
        //                    // 后续点提示
        //                    var nextOpts = new PromptPointOptions("\n指定下一个点（右键/回车结束）：");
        //                    // 启用基点橡皮筋
        //                    nextOpts.UseBasePoint = true;
        //                    nextOpts.BasePoint = points.Last();
        //                    nextOpts.AllowNone = true;
        //                    nextOpts.Keywords.Add("完成");
        //                    nextOpts.AppendKeywordsToMessage = true;

        //                    // 读取后续点
        //                    var nextRes = ed.GetPoint(nextOpts);

        //                    // 正常点输入
        //                    if (nextRes.Status == PromptStatus.OK)
        //                    {
        //                        var pt = nextRes.Value;
        //                        if (pt.IsEqualTo(points.Last())) break;
        //                        points.Add(pt);
        //                        continue;
        //                    }

        //                    // 回车/右键结束
        //                    if (nextRes.Status == PromptStatus.None) break;
        //                    // 关键字完成结束
        //                    if (nextRes.Status == PromptStatus.Keyword && string.Equals(nextRes.StringResult, "完成", StringComparison.OrdinalIgnoreCase)) break;
        //                    // 其它状态统一结束
        //                    break;
        //                }

        //                // 至少两点
        //                if (points.Count < 2)
        //                {
        //                    ed.WriteMessage("\n采点不足，取消生成。");
        //                    if (tempInserted) { try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { } }
        //                    tr.Abort();
        //                    return;
        //                }

        //                // 本地函数——分析样例块（支持一级嵌套块回退）
        //                SamplePipeInfo AnalyzeSampleWithFallback()
        //                {
        //                    // 先按当前样例直接分析
        //                    var info = AnalyzeSampleBlock(tr, sampleBr);
        //                    if (info != null && info.PipeBodyTemplate != null) return info;

        //                    try
        //                    {
        //                        // 打开样例块定义
        //                        var hostBtr = tr.GetObject(sampleBr.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
        //                        if (hostBtr == null) return info;

        //                        // 尝试从一级嵌套块中找真正的管道模板
        //                        foreach (ObjectId id in hostBtr)
        //                        {
        //                            var nested = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
        //                            if (nested == null) continue;

        //                            var nestedInfo = AnalyzeSampleBlock(tr, nested);
        //                            if (nestedInfo != null && nestedInfo.PipeBodyTemplate != null)
        //                            {
        //                                // 命中嵌套模板则返回
        //                                return nestedInfo;
        //                            }
        //                        }
        //                    }
        //                    catch
        //                    {
        //                        // 嵌套分析失败则回退原结果
        //                    }

        //                    return info;
        //                }

        //                // 分析样例块（含嵌套回退）
        //                var sampleInfo = AnalyzeSampleWithFallback();

        //                // 未找到管道主体则提示并退出
        //                if (sampleInfo == null || sampleInfo.PipeBodyTemplate == null)
        //                {
        //                    ed.WriteMessage("\n示例块中未找到管道主体。");
        //                    if (tempInserted) { try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { } }
        //                    tr.Abort();
        //                    return;
        //                }

        //                // 计算管线长度
        //                double pipelineLength = ComputePipelineLengthByPoints(points);
        //                // 计算中点
        //                var midTuple = ComputeMidPointAndAngle(points, pipelineLength);
        //                Point3d midPoint = midTuple.midPoint;

        //                // 构建局部管线
        //                var pipeLocal = BuildPipePolylineLocal(sampleInfo.PipeBodyTemplate, points, midPoint);

        //                // 读取样例属性作为基线
        //                var latestSampleAttrs = GetEntityAttributeMap(tr, sampleBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //                // 端点拾取半径
        //                double pickRadius = Math.Max(200.0, VariableDictionary.textBoxScale * 5.0);

        //                // 找起点附近部件
        //                var startBr = FindNearestBlockReferenceNearPoint(tr, points.First(), pickRadius, sampleBr.ObjectId);
        //                // 找终点附近部件
        //                var endBr = FindNearestBlockReferenceNearPoint(tr, points.Last(), pickRadius, sampleBr.ObjectId);

        //                // 读取起点部件属性
        //                var startMap = startBr != null
        //                    ? (GetEntityAttributeMap(tr, startBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
        //                    : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //                // 读取终点部件属性
        //                var endMap = endBr != null
        //                    ? (GetEntityAttributeMap(tr, endBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
        //                    : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //                // 候选键
        //                string[] nameKeys = { "名称", "设备名称", "Name", "TAG", "Tag", "位号", "设备位号" };
        //                string[] titleKeys = { "管道标题", "标题", "Title" };
        //                string[] pipeNoKeys = { "管段号", "管段编号", "Pipeline No", "Pipe No", "No" };

        //                // 本地函数，按候选键找首个非空值
        //                string FindFirstAttrValueLocal(Dictionary<string, string> attrs, string[] keys)
        //                {
        //                    if (attrs == null || attrs.Count == 0) return string.Empty;

        //                    foreach (var k in keys)
        //                    {
        //                        if (attrs.TryGetValue(k, out var v) && !string.IsNullOrWhiteSpace(v))
        //                            return v.Trim();
        //                    }

        //                    foreach (var kv in attrs)
        //                    {
        //                        if (string.IsNullOrWhiteSpace(kv.Key) || string.IsNullOrWhiteSpace(kv.Value)) continue;
        //                        foreach (var k in keys)
        //                        {
        //                            if (kv.Key.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
        //                                return kv.Value.Trim();
        //                        }
        //                    }

        //                    return string.Empty;
        //                }

        //                // 提取端点名称
        //                string startName = FindFirstAttrValueLocal(startMap, nameKeys);
        //                string endName = FindFirstAttrValueLocal(endMap, nameKeys);

        //                // 提取标题与管段号
        //                string startTitle = FindFirstAttrValueLocal(startMap, titleKeys);
        //                string endTitle = FindFirstAttrValueLocal(endMap, titleKeys);
        //                string sampleTitle = latestSampleAttrs.TryGetValue("管道标题", out var st) ? (st ?? string.Empty).Trim() : string.Empty;
        //                string finalTitle = !string.IsNullOrWhiteSpace(startTitle) ? startTitle
        //                                : !string.IsNullOrWhiteSpace(endTitle) ? endTitle
        //                                : sampleTitle;

        //                string startPipeNo = FindFirstAttrValueLocal(startMap, pipeNoKeys);
        //                string endPipeNo = FindFirstAttrValueLocal(endMap, pipeNoKeys);
        //                string samplePipeNo = latestSampleAttrs.TryGetValue("管段号", out var spn) ? (spn ?? string.Empty).Trim() : string.Empty;
        //                string finalPipeNo = !string.IsNullOrWhiteSpace(startPipeNo) ? startPipeNo
        //                                 : !string.IsNullOrWhiteSpace(endPipeNo) ? endPipeNo
        //                                 : samplePipeNo;

        //                // 坐标回退值
        //                string startCoordStr = $"X={points.First().X:F3},Y={points.First().Y:F3}";
        //                string endCoordStr = $"X={points.Last().X:F3},Y={points.Last().Y:F3}";

        //                // 始点终点优先写名称，否则写坐标
        //                latestSampleAttrs["始点"] = !string.IsNullOrWhiteSpace(startName) ? startName : startCoordStr;
        //                latestSampleAttrs["终点"] = !string.IsNullOrWhiteSpace(endName) ? endName : endCoordStr;

        //                // 附加保存始点/终点名称
        //                if (!string.IsNullOrWhiteSpace(startName)) latestSampleAttrs["始点名称"] = startName;
        //                if (!string.IsNullOrWhiteSpace(endName)) latestSampleAttrs["终点名称"] = endName;

        //                // 回填标题/管段号
        //                if (!string.IsNullOrWhiteSpace(finalTitle)) latestSampleAttrs["管道标题"] = finalTitle;
        //                if (!string.IsNullOrWhiteSpace(finalPipeNo)) latestSampleAttrs["管段号"] = finalPipeNo;

        //                // 命令行日志
        //                ed.WriteMessage($"\n端点拾取半径={pickRadius:F2}；起点部件={(startBr != null ? startBr.Name : "未命中")}；终点部件={(endBr != null ? endBr.Name : "未命中")}。");
        //                ed.WriteMessage($"\n端点名称：始点='{latestSampleAttrs["始点"]}'，终点='{latestSampleAttrs["终点"]}'。");
        //                ed.WriteMessage($"\n标题/管段号：管道标题='{(latestSampleAttrs.ContainsKey("管道标题") ? latestSampleAttrs["管道标题"] : "")}'，管段号='{(latestSampleAttrs.ContainsKey("管段号") ? latestSampleAttrs["管段号"] : "")}'。");

        //                // 生成标题文字内容
        //                string pipeTitle = latestSampleAttrs.TryGetValue("管道标题", out var sample_Title) && !string.IsNullOrWhiteSpace(sample_Title)
        //                                 ? sample_Title
        //                                 : sampleBr.Name ?? "管道";

        //                // 创建箭头与标题实体
        //                var arrowEntities = CreateDirectionalArrowsAndTitles(tr, sampleInfo, points, midPoint, pipeTitle, sampleBr.Name);

        //                // 克隆属性定义
        //                var attDefsLocal = CloneAttributeDefinitionsLocal(sampleInfo.AttributeDefinitions, midPoint, 0.0, pipelineLength, sampleBr.Name)
        //                                   ?? new List<AttributeDefinition>();

        //                // 有模板属性定义时按Tag回填
        //                if (sampleInfo.AttributeDefinitions != null && sampleInfo.AttributeDefinitions.Count > 0)
        //                {
        //                    foreach (var def in attDefsLocal)
        //                    {
        //                        if (string.IsNullOrWhiteSpace(def.Tag)) continue;
        //                        if (latestSampleAttrs.TryGetValue(def.Tag, out var v))
        //                            def.TextString = v ?? string.Empty;
        //                        def.Invisible = false;
        //                        def.Constant = false;
        //                    }
        //                }
        //                else
        //                {
        //                    // 无模板属性定义时，按字典动态生成
        //                    attDefsLocal.Clear();
        //                    double attHeight = 2.5;
        //                    double yOffsetBase = -attHeight * 2.0;
        //                    int idx = 0;

        //                    foreach (var kv in latestSampleAttrs.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
        //                    {
        //                        if (string.IsNullOrWhiteSpace(kv.Key)) continue;
        //                        attDefsLocal.Add(new AttributeDefinition
        //                        {
        //                            Tag = kv.Key,
        //                            Position = new Point3d(0, yOffsetBase - idx * attHeight * 1.2, 0),
        //                            Rotation = 0.0,
        //                            TextString = kv.Value ?? string.Empty,
        //                            Height = attHeight,
        //                            Invisible = false,
        //                            Constant = false
        //                        });
        //                        idx++;
        //                    }
        //                }

        //                // 生成管段号（样例优先，序号兜底）
        //                int nextSegNum = GetNextPipeSegmentNumber(db);
        //                string extractedPipeNo = nextSegNum.ToString("D4");
        //                if (latestSampleAttrs.TryGetValue("管段号", out var pn) && !string.IsNullOrWhiteSpace(pn))
        //                    extractedPipeNo = pn;

        //                // 计算附加属性定位参数
        //                double finalAttHeight = attDefsLocal.Count > 0 ? attDefsLocal[0].Height : 2.5;
        //                double finalYOffsetBase = attDefsLocal.Count > 0 ? attDefsLocal[0].Position.Y - finalAttHeight * 1.2 : -finalAttHeight * 2.0;
        //                int extraIndex = 0;

        //                // 确保关键属性存在
        //                SetOrAddAttr(attDefsLocal, "始点", latestSampleAttrs["始点"], ref extraIndex, finalYOffsetBase, finalAttHeight);
        //                SetOrAddAttr(attDefsLocal, "终点", latestSampleAttrs["终点"], ref extraIndex, finalYOffsetBase, finalAttHeight);
        //                SetOrAddLengthAttrs(attDefsLocal, pipelineLength, ref extraIndex, finalYOffsetBase, finalAttHeight);
        //                SetOrAddAttr(attDefsLocal, "管段号", extractedPipeNo, ref extraIndex, finalYOffsetBase, finalAttHeight);

        //                // 保持原行为，属性默认隐藏
        //                foreach (var ad in attDefsLocal)
        //                {
        //                    ad.Invisible = true;
        //                    ad.Constant = false;
        //                }

        //                // 块名兜底，避免空名
        //                string desiredName = string.IsNullOrWhiteSpace(sampleBr.Name) ? "PIPE_TEMPLATE" : sampleBr.Name;

        //                // 构建块定义并插入
        //                string newBlockName = BuildPipeBlockDefinition(tr, desiredName, (Polyline)pipeLocal.Clone(), arrowEntities, attDefsLocal);

        //                // 整理属性值字典
        //                var attValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //                foreach (var a in attDefsLocal)
        //                {
        //                    if (string.IsNullOrWhiteSpace(a?.Tag)) continue;
        //                    attValues[a.Tag] = a.TextString ?? string.Empty;
        //                }

        //                // 插入新块
        //                var newBrId = InsertPipeBlockWithAttributes(tr, midPoint, newBlockName, 0.0, attValues);

        //                // 后处理拓扑
        //                PipelineTopologyHelper.PostProcessAfterPipePlaced(tr, newBrId, isOutlet ? "Outlet" : "Inlet");

        //                // 继承图层
        //                var newBr = tr.GetObject(newBrId, OpenMode.ForWrite) as BlockReference;
        //                if (newBr != null) newBr.Layer = sampleInfo.PipeBodyTemplate.Layer;

        //                // 清理临时样例块
        //                if (tempInserted)
        //                {
        //                    try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { }
        //                }

        //                // 提交事务
        //                tr.Commit();

        //                // 输出完成消息
        //                ed.WriteMessage($"\n管道已生成（{(isOutlet ? "出口" : "入口")}），点数={points.Count}，管段号={extractedPipeNo}。");
        //            }
        //            catch (Exception ex)
        //            {
        //                // 事务内异常回滚
        //                tr.Abort();
        //                ed.WriteMessage($"\n生成管道时出错: {ex.Message}");
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        // 总异常兜底
        //        ed.WriteMessage($"\n操作失败: {ex.Message}");
        //    }
        //}

        /// <summary>
        /// 按候选键提取第一个非空属性值（先精确匹配，再包含匹配）
        /// </summary>
        //private static string FindFirstNonEmptyAttrValueForPipe(Dictionary<string, string> attrs, string[] keys)
        //{
        //    // 空字典直接返回空
        //    if (attrs == null || attrs.Count == 0) return string.Empty;

        //    // 先做精确键匹配
        //    foreach (var k in keys)
        //    {
        //        if (attrs.TryGetValue(k, out var v) && !string.IsNullOrWhiteSpace(v))
        //            return v.Trim();
        //    }

        //    // 再做包含匹配（兼容“设备名称1”“起点名称”等）
        //    foreach (var kv in attrs)
        //    {
        //        if (string.IsNullOrWhiteSpace(kv.Key) || string.IsNullOrWhiteSpace(kv.Value)) continue;
        //        foreach (var k in keys)
        //        {
        //            if (kv.Key.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
        //                return kv.Value.Trim();
        //        }
        //    }

        //    // 未命中返回空
        //    return string.Empty;
        //}

        /// <summary>
        /// 判断点是否在包围盒内（带容差）
        /// </summary>
        private static bool IsPointInsideExtentsForPipeInheritance(Point3d p, Extents3d ext, double tol)
        {
            // X 方向判断
            if (p.X < ext.MinPoint.X - tol || p.X > ext.MaxPoint.X + tol) return false;
            // Y 方向判断
            if (p.Y < ext.MinPoint.Y - tol || p.Y > ext.MaxPoint.Y + tol) return false;
            // Z 方向判断
            if (p.Z < ext.MinPoint.Z - tol || p.Z > ext.MaxPoint.Z + tol) return false;
            // 全部通过即命中
            return true;
        }

        /// <summary>
        /// 查找与指定端点“相交/命中”的候选块（优先包围盒命中，其次插入点接近）
        /// </summary>
        //private List<BlockReference> FindBlocksCrossingPointForPipeInheritance(DBTrans tr, Point3d point, ObjectId excludeId, double tol)
        //{
        //    // 包围盒命中列表
        //    var insideHits = new List<BlockReference>();
        //    // 近距离命中列表（包围盒未命中，但插入点很近）
        //    var nearHits = new List<BlockReference>();

        //    // 遍历当前空间实体
        //    foreach (ObjectId id in tr.CurrentSpace)
        //    {
        //        // 排除指定对象（通常是样例块本身）
        //        if (!excludeId.IsNull && id == excludeId) continue;

        //        // 读取实体
        //        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
        //        // 过滤无效实体
        //        if (ent == null || ent.IsErased) continue;

        //        // 仅处理块参照（便于后续继承属性）
        //        var br = ent as BlockReference;
        //        if (br == null) continue;

        //        // 默认非包围盒命中
        //        bool inside = false;

        //        // 尝试读取包围盒并判断端点是否落入
        //        try
        //        {
        //            var ext = br.GeometricExtents;
        //            inside = IsPointInsideExtentsForPipeInheritance(point, ext, tol);
        //        }
        //        catch
        //        {
        //            // 包围盒失败则保持 inside=false，走距离兜底
        //        }

        //        // 包围盒命中优先加入 insideHits
        //        if (inside)
        //        {
        //            insideHits.Add(br);
        //            continue;
        //        }

        //        // 包围盒未命中时，按插入点距离兜底
        //        var d = br.Position.DistanceTo(point);
        //        if (d <= tol * 2.0)
        //        {
        //            nearHits.Add(br);
        //        }
        //    }

        //    // 按“距离端点近”排序，提高稳定性
        //    insideHits = insideHits.OrderBy(b => b.Position.DistanceTo(point)).ToList();
        //    nearHits = nearHits.OrderBy(b => b.Position.DistanceTo(point)).ToList();

        //    // 先拼 insideHits，再拼 nearHits
        //    var result = new List<BlockReference>();
        //    result.AddRange(insideHits);
        //    result.AddRange(nearHits);

        //    // 返回候选
        //    return result;
        //}

        /// <summary>
        /// 完成版——手选示例块 + 端点相交继承参数 + 同名标签同步（增强：读取到0不覆盖，字段标准化）
        /// </summary>
        private void DrawPipeByClicks(bool isOutlet)
        {
            // 获取当前文档
            var doc = Application.DocumentManager.MdiActiveDocument;
            // 无文档直接返回
            if (doc == null) return;

            // 获取编辑器和数据库
            var ed = doc.Editor;
            Database db = doc.Database;

            // 本地函数——判断值是否应跳过（空或0）
            bool IsZeroLikeForPipe(string raw)
            {
                // 空值兜底
                string text = (raw ?? string.Empty).Trim();
                // 空串直接视为无效继承值
                if (string.IsNullOrWhiteSpace(text)) return true;

                // 兼容全角0和中文空格
                text = text.Replace('０', '0').Replace('\u3000', ' ').Trim();

                // 纯0直接判定
                if (string.Equals(text, "0", StringComparison.OrdinalIgnoreCase)) return true;

                // 先尝试纯数字判定
                double dv;
                if (double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out dv))
                {
                    if (Math.Abs(dv) < 1e-12) return true;
                }
                else if (double.TryParse(text, out dv))
                {
                    if (Math.Abs(dv) < 1e-12) return true;
                }

                // 带单位0判定（例如 0MPa、0℃、0 m/s）
                var m = System.Text.RegularExpressions.Regex.Match(
                    text,
                    @"^\s*([+-]?(?:0+(?:[.,]0+)?|[.,]0+))\s*[^0-9]*\s*$",
                    System.Text.RegularExpressions.RegexOptions.CultureInvariant);

                if (m.Success)
                {
                    string n = (m.Groups[1].Value ?? string.Empty).Replace(',', '.');
                    if (double.TryParse(n, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out dv))
                    {
                        if (Math.Abs(dv) < 1e-12) return true;
                    }
                }

                // 其余值保留
                return false;
            }

            // 本地函数——按候选键提取首个“非空且非0”值
            string FindFirstNonZeroAttrValueLocal(Dictionary<string, string> attrs, string[] keys)
            {
                // 空字典直接返回空
                if (attrs == null || attrs.Count == 0) return string.Empty;

                // 先做精确键匹配
                foreach (var k in keys)
                {
                    if (!attrs.TryGetValue(k, out var v)) continue;
                    string val = (v ?? string.Empty).Trim();
                    if (IsZeroLikeForPipe(val)) continue;
                    return val;
                }

                // 再做包含匹配
                foreach (var kv in attrs)
                {
                    if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                    string val = (kv.Value ?? string.Empty).Trim();
                    if (IsZeroLikeForPipe(val)) continue;

                    foreach (var k in keys)
                    {
                        if (kv.Key.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
                            return val;
                    }
                }

                // 未命中返回空
                return string.Empty;
            }

            // 本地函数——从多个候选中取第一个非空且非0值
            string PickFirstNonZero(params string[] candidates)
            {
                foreach (var c in candidates)
                {
                    string v = (c ?? string.Empty).Trim();
                    if (IsZeroLikeForPipe(v)) continue;
                    return v;
                }
                return string.Empty;
            }

            try
            {
                // 统一手选样例，不走自动匹配
                bool tempInserted = false;
                BlockReference sampleBr = null;

                // 开启事务
                using (var tr = new DBTrans())
                {
                    try
                    {
                        // 提示先选择示例块
                        ed.WriteMessage($"\n请先选择{(isOutlet ? "出口" : "入口")}示例管线块作为模板。");
                        var peo = new PromptEntityOptions("\n请选择示例管线块：");
                        peo.SetRejectMessage("\n请选择块参照对象。");
                        peo.AddAllowedClass(typeof(BlockReference), true);
                        var per = ed.GetEntity(peo);

                        // 取消选择直接结束
                        if (per.Status != PromptStatus.OK)
                        {
                            ed.WriteMessage("\n未选择示例块，操作取消。");
                            tr.Abort();
                            return;
                        }

                        // 读取示例块并升级为可写
                        sampleBr = tr.GetObject(per.ObjectId, OpenMode.ForRead) as BlockReference;
                        if (sampleBr != null && !sampleBr.IsWriteEnabled) sampleBr.UpgradeOpen();

                        // 判空保护
                        if (sampleBr == null)
                        {
                            ed.WriteMessage("\n无法获取示例块，取消。");
                            tr.Abort();
                            return;
                        }

                        // 读取历史属性
                        var loadedAttrs = FileManager.LoadLastPipeAttributes(isOutlet) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                        // 历史属性为空时，用样例块属性兜底
                        if (loadedAttrs.Count == 0)
                        {
                            var sampleAttrMap = GetEntityAttributeMap(tr, sampleBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            foreach (var kv in sampleAttrMap)
                            {
                                if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                                loadedAttrs[kv.Key.Trim()] = kv.Value ?? string.Empty;
                            }
                            ed.WriteMessage($"\n历史属性为空，已用样例块属性兜底初始化 {loadedAttrs.Count} 项。");
                        }

                        // 弹出属性编辑窗口
                        using (var editor = new PipeAttributeEditorForm(loadedAttrs))
                        {
                            var dr = editor.ShowDialog();
                            if (dr != DialogResult.OK)
                            {
                                tr.Abort();
                                ed.WriteMessage("\n已取消属性编辑，终止。");
                                return;
                            }

                            // 保存编辑结果
                            var editedAttrs = editor.Attributes ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            FileManager.SaveLastPipeAttributes(isOutlet, editedAttrs);

                            // 回写示例块已有属性
                            try
                            {
                                var sampleBrWrite = tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference;
                                if (sampleBrWrite != null)
                                {
                                    foreach (ObjectId aid in sampleBrWrite.AttributeCollection)
                                    {
                                        var ar = tr.GetObject(aid, OpenMode.ForWrite) as AttributeReference;
                                        if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue;
                                        if (editedAttrs.TryGetValue(ar.Tag, out var newVal))
                                        {
                                            ar.TextString = newVal ?? string.Empty;
                                            try { ar.AdjustAlignment(db); } catch { }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // 忽略单点失败
                            }
                        }

                        // 采集起点
                        var points = new List<Point3d>();
                        var firstOpts = new PromptPointOptions("\n指定起点（右键/回车取消）：");
                        firstOpts.AllowNone = true;
                        firstOpts.Keywords.Add("取消");
                        firstOpts.AppendKeywordsToMessage = true;
                        var firstRes = ed.GetPoint(firstOpts);

                        // 起点取消处理
                        if (firstRes.Status == PromptStatus.None ||
                            (firstRes.Status == PromptStatus.Keyword && string.Equals(firstRes.StringResult, "取消", StringComparison.OrdinalIgnoreCase)))
                        {
                            ed.WriteMessage("\n已取消：未指定起点。");
                            tr.Abort();
                            return;
                        }

                        // 起点异常状态处理
                        if (firstRes.Status != PromptStatus.OK)
                        {
                            ed.WriteMessage("\n已取消：起点输入终止。");
                            tr.Abort();
                            return;
                        }

                        // 加入起点
                        points.Add(firstRes.Value);

                        // 循环采集后续点
                        while (true)
                        {
                            var nextOpts = new PromptPointOptions("\n指定下一个点（右键/回车结束）：");
                            nextOpts.UseBasePoint = true;
                            nextOpts.BasePoint = points.Last();
                            nextOpts.AllowNone = true;
                            nextOpts.Keywords.Add("完成");
                            nextOpts.AppendKeywordsToMessage = true;

                            var nextRes = ed.GetPoint(nextOpts);

                            if (nextRes.Status == PromptStatus.OK)
                            {
                                var pt = nextRes.Value;
                                if (pt.IsEqualTo(points.Last())) break;
                                points.Add(pt);
                                continue;
                            }

                            if (nextRes.Status == PromptStatus.None) break;
                            if (nextRes.Status == PromptStatus.Keyword && string.Equals(nextRes.StringResult, "完成", StringComparison.OrdinalIgnoreCase)) break;
                            break;
                        }

                        // 至少需要两点
                        if (points.Count < 2)
                        {
                            ed.WriteMessage("\n采点不足，取消生成。");
                            tr.Abort();
                            return;
                        }

                        // 分析样例块
                        var sampleInfo = AnalyzeSampleBlock(tr, sampleBr);

                        // 样例无主体时做一级嵌套回退
                        if (sampleInfo == null || sampleInfo.PipeBodyTemplate == null)
                        {
                            try
                            {
                                var hostBtr = tr.GetObject(sampleBr.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                if (hostBtr != null)
                                {
                                    foreach (ObjectId id in hostBtr)
                                    {
                                        var nested = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                                        if (nested == null) continue;

                                        var nestedInfo = AnalyzeSampleBlock(tr, nested);
                                        if (nestedInfo != null && nestedInfo.PipeBodyTemplate != null)
                                        {
                                            sampleInfo = nestedInfo;
                                            break;
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // 嵌套回退失败不抛出
                            }
                        }

                        // 仍无主体则失败
                        if (sampleInfo == null || sampleInfo.PipeBodyTemplate == null)
                        {
                            ed.WriteMessage("\n示例块中未找到管道主体。");
                            tr.Abort();
                            return;
                        }

                        // 计算长度与中点
                        double pipelineLength = ComputePipelineLengthByPoints(points);
                        var (midPoint, _) = ComputeMidPointAndAngle(points, pipelineLength);

                        // 构建局部管线
                        var pipeLocal = BuildPipePolylineLocal(sampleInfo.PipeBodyTemplate, points, midPoint);

                        // 读取样例属性作为基线
                        var latestSampleAttrs = GetEntityAttributeMap(tr, sampleBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                        // 端点命中容差（按比例自适应）
                        double hitTol = Math.Max(1.0, VariableDictionary.textBoxScale * 0.2);

                        // 查找起终点命中候选
                        var startCandidates = FindBlocksCrossingPoint(tr, points.First(), sampleBr.ObjectId, hitTol);
                        var endCandidates = FindBlocksCrossingPoint(tr, points.Last(), sampleBr.ObjectId, hitTol);

                        // 取最优候选
                        var startBr = startCandidates.FirstOrDefault();
                        var endBr = endCandidates.FirstOrDefault();

                        // 读取命中图元属性
                        var startMap = startBr != null
                            ? (GetEntityAttributeMap(tr, startBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
                            : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                        var endMap = endBr != null
                            ? (GetEntityAttributeMap(tr, endBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
                            : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                        // 关键字段候选键
                        string[] nameKeys = { "名称", "设备名称", "Name", "TAG", "Tag", "位号", "设备位号" };
                        string[] titleKeys = { "管道标题", "标题", "Title" };
                        string[] pipeNoKeys = { "管段号", "管段编号", "管道号", "Pipeline No", "Pipe No", "No" };

                        // 提取端点名称（自动跳过0）
                        string startName = FindFirstNonZeroAttrValueLocal(startMap, nameKeys);
                        string endName = FindFirstNonZeroAttrValueLocal(endMap, nameKeys);

                        // 提取标题（起点优先，终点次之，样例兜底；自动跳过0）
                        string startTitle = FindFirstNonZeroAttrValueLocal(startMap, titleKeys);
                        string endTitle = FindFirstNonZeroAttrValueLocal(endMap, titleKeys);
                        string sampleTitle = latestSampleAttrs.TryGetValue("管道标题", out var st) ? (st ?? string.Empty).Trim() : string.Empty;
                        string finalTitle = PickFirstNonZero(startTitle, endTitle, sampleTitle);

                        // 提取管段号（起点优先，终点次之，样例兜底；自动跳过0）
                        string startPipeNo = FindFirstNonZeroAttrValueLocal(startMap, pipeNoKeys);
                        string endPipeNo = FindFirstNonZeroAttrValueLocal(endMap, pipeNoKeys);
                        string samplePipeNo = latestSampleAttrs.TryGetValue("管段号", out var spn) ? (spn ?? string.Empty).Trim() : string.Empty;
                        string finalPipeNo = PickFirstNonZero(startPipeNo, endPipeNo, samplePipeNo);

                        // 坐标兜底
                        string startCoordStr = $"X={points.First().X:F3},Y={points.First().Y:F3}";
                        string endCoordStr = $"X={points.Last().X:F3},Y={points.Last().Y:F3}";

                        // 写入标准字段
                        latestSampleAttrs["起点"] = !string.IsNullOrWhiteSpace(startName) ? startName : startCoordStr;
                        latestSampleAttrs["终点"] = !string.IsNullOrWhiteSpace(endName) ? endName : endCoordStr;

                        // 附加调试字段
                        if (!string.IsNullOrWhiteSpace(startName)) latestSampleAttrs["起点名称"] = startName;
                        if (!string.IsNullOrWhiteSpace(endName)) latestSampleAttrs["终点名称"] = endName;

                        // 仅当非0时回填标题与标准管号
                        if (!string.IsNullOrWhiteSpace(finalTitle) && !IsZeroLikeForPipe(finalTitle))
                            latestSampleAttrs["管道标题"] = finalTitle;
                        if (!string.IsNullOrWhiteSpace(finalPipeNo) && !IsZeroLikeForPipe(finalPipeNo))
                            latestSampleAttrs["管段号"] = finalPipeNo;

                        // 同名标签同步（排除“名称”；且值为0不覆盖）
                        foreach (var kv in startMap)
                        {
                            if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                            if (string.Equals(kv.Key, "名称", StringComparison.OrdinalIgnoreCase)) continue;
                            if (string.Equals(kv.Key, "Name", StringComparison.OrdinalIgnoreCase)) continue;

                            string incoming = (kv.Value ?? string.Empty).Trim();
                            if (IsZeroLikeForPipe(incoming)) continue;

                            if (latestSampleAttrs.ContainsKey(kv.Key))
                                latestSampleAttrs[kv.Key] = incoming;
                        }

                        foreach (var kv in endMap)
                        {
                            if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                            if (string.Equals(kv.Key, "名称", StringComparison.OrdinalIgnoreCase)) continue;
                            if (string.Equals(kv.Key, "Name", StringComparison.OrdinalIgnoreCase)) continue;

                            string incoming = (kv.Value ?? string.Empty).Trim();
                            if (IsZeroLikeForPipe(incoming)) continue;

                            if (latestSampleAttrs.ContainsKey(kv.Key))
                                latestSampleAttrs[kv.Key] = incoming;
                        }

                        // 字段标准化（根治重复：起点/终点/管段号）
                        latestSampleAttrs = NormalizePipeAttributeKeys(latestSampleAttrs);

                        // 日志
                        ed.WriteMessage($"\n端点命中容差={hitTol:F2}；起点命中={(startBr != null ? startBr.Name : "无")}；终点命中={(endBr != null ? endBr.Name : "无")}。");

                        // 确定标题
                        string pipeTitle = latestSampleAttrs.TryGetValue("管道标题", out var sample_Title) && !string.IsNullOrWhiteSpace(sample_Title) && !IsZeroLikeForPipe(sample_Title)
                                         ? sample_Title
                                         : sampleBr.Name ?? "管道";

                        // 创建箭头与标题实体
                        var arrowEntities = CreateDirectionalArrowsAndTitles(tr, sampleInfo, points, midPoint, pipeTitle, sampleBr.Name);

                        // 克隆模板属性定义
                        var attDefsLocal = CloneAttributeDefinitionsLocal(sampleInfo.AttributeDefinitions, midPoint, 0.0, pipelineLength, sampleBr.Name)
                                           ?? new List<AttributeDefinition>();

                        // 按模板定义回填值（值为0不回填覆盖）
                        if (sampleInfo.AttributeDefinitions != null && sampleInfo.AttributeDefinitions.Count > 0)
                        {
                            foreach (var def in attDefsLocal)
                            {
                                if (string.IsNullOrWhiteSpace(def.Tag)) continue;
                                if (latestSampleAttrs.TryGetValue(def.Tag, out var v))
                                {
                                    string vv = (v ?? string.Empty).Trim();
                                    if (!IsZeroLikeForPipe(vv))
                                        def.TextString = vv;
                                }
                                def.Invisible = false;
                                def.Constant = false;
                            }
                        }
                        else
                        {
                            // 模板无属性定义则动态生成（值为0不生成）
                            attDefsLocal.Clear();
                            double attHeight = 2.5;
                            double yOffsetBase = -attHeight * 2.0;
                            int idx = 0;

                            foreach (var kv in latestSampleAttrs.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
                            {
                                if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                                string vv = (kv.Value ?? string.Empty).Trim();
                                if (IsZeroLikeForPipe(vv)) continue;

                                attDefsLocal.Add(new AttributeDefinition
                                {
                                    Tag = kv.Key,
                                    Position = new Point3d(0, yOffsetBase - idx * attHeight * 1.2, 0),
                                    Rotation = 0.0,
                                    TextString = vv,
                                    Height = attHeight,
                                    Invisible = false,
                                    Constant = false
                                });
                                idx++;
                            }
                        }

                        // 最终管段号（继承值为0时回退自动编号）
                        int nextSegNum = GetNextPipeSegmentNumber(db);
                        string extractedPipeNo = nextSegNum.ToString("D4");
                        if (latestSampleAttrs.TryGetValue("管段号", out var pn))
                        {
                            string pnv = (pn ?? string.Empty).Trim();
                            if (!IsZeroLikeForPipe(pnv) && !string.IsNullOrWhiteSpace(pnv))
                                extractedPipeNo = pnv;
                        }

                        // 计算附加属性布局参数
                        double finalAttHeight = attDefsLocal.Count > 0 ? attDefsLocal[0].Height : 2.5;
                        double finalYOffsetBase = attDefsLocal.Count > 0 ? attDefsLocal[0].Position.Y - finalAttHeight * 1.2 : -finalAttHeight * 2.0;
                        int extraIndex = 0;

                        // 仅写标准字段（清理旧的始点/起始/管道号写法）
                        SetOrAddCanonicalPipeAttrs(
                            attDefsLocal,
                            latestSampleAttrs.TryGetValue("起点", out var pStart) ? pStart : string.Empty,
                            latestSampleAttrs.TryGetValue("终点", out var pEnd) ? pEnd : string.Empty,
                            extractedPipeNo,
                            ref extraIndex,
                            finalYOffsetBase,
                            finalAttHeight);

                        // 长度字段保持原逻辑
                        SetOrAddLengthAttrs(attDefsLocal, pipelineLength, ref extraIndex, finalYOffsetBase, finalAttHeight);

                        // 保持原行为，定义默认隐藏
                        foreach (var ad in attDefsLocal)
                        {
                            ad.Invisible = true;
                            ad.Constant = false;
                        }

                        // 块名兜底
                        string desiredName = string.IsNullOrWhiteSpace(sampleBr.Name) ? "PIPE_TEMPLATE" : sampleBr.Name;

                        // 构建块定义
                        string newBlockName = BuildPipeBlockDefinition(tr, desiredName, (Polyline)pipeLocal.Clone(), arrowEntities, attDefsLocal);

                        // 整理属性值（值为0不写入最终字典）
                        var attValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var a in attDefsLocal)
                        {
                            if (string.IsNullOrWhiteSpace(a?.Tag)) continue;
                            string vv = (a.TextString ?? string.Empty).Trim();
                            if (IsZeroLikeForPipe(vv)) continue;
                            attValues[a.Tag] = vv;
                        }

                        // 插入新块
                        var newBrId = InsertPipeBlockWithAttributes(tr, midPoint, newBlockName, 0.0, attValues);

                        // 后处理拓扑关系
                        PipelineTopologyHelper.PostProcessAfterPipePlaced(tr, newBrId, isOutlet ? "Outlet" : "Inlet");

                        // 继承图层
                        var newBr = tr.GetObject(newBrId, OpenMode.ForWrite) as BlockReference;
                        if (newBr != null) newBr.Layer = sampleInfo.PipeBodyTemplate.Layer;

                        // 兼容分支（当前为手选样例，不会进入）
                        if (tempInserted)
                        {
                            try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { }
                        }

                        // 提交事务
                        tr.Commit();

                        // 输出成功消息
                        ed.WriteMessage($"\n管道已生成（{(isOutlet ? "出口" : "入口")}），点数={points.Count}，管段号={extractedPipeNo}。");
                    }
                    catch (Exception ex)
                    {
                        // 事务内异常回滚
                        tr.Abort();
                        ed.WriteMessage($"\n生成管道时出错: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // 总异常兜底
                ed.WriteMessage($"\n操作失败: {ex.Message}");
            }
        }


        #region 新增 3 个私有辅助方法

        /// <summary>
        /// 判断点是否落在包围盒内（带容差）
        /// </summary>
        private static bool IsPointInsideExtents(Point3d p, Extents3d ext, double tol)
        {
            // X 范围判断
            if (p.X < ext.MinPoint.X - tol || p.X > ext.MaxPoint.X + tol) return false;
            // Y 范围判断
            if (p.Y < ext.MinPoint.Y - tol || p.Y > ext.MaxPoint.Y + tol) return false;
            // Z 范围判断
            if (p.Z < ext.MinPoint.Z - tol || p.Z > ext.MaxPoint.Z + tol) return false;
            // 全部通过则命中
            return true;
        }

        /// <summary>
        /// 查找“与端点相交/命中”的候选块参照（优先包围盒包含，其次插入点接近）
        /// </summary>
        private List<BlockReference> FindBlocksCrossingPoint(DBTrans tr, Point3d point, ObjectId excludeId, double tol)
        {
            // 结果列表
            var hits = new List<(BlockReference br, bool inside, double dist)>();

            // 遍历当前空间
            foreach (ObjectId id in tr.CurrentSpace)
            {
                // 排除自身
                if (!excludeId.IsNull && id == excludeId) continue;

                // 读取实体
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                // 过滤无效实体
                if (ent == null || ent.IsErased) continue;

                // 仅处理块参照（因为要读属性）
                var br = ent as BlockReference;
                if (br == null) continue;

                bool inside = false;
                double dist = br.Position.DistanceTo(point);

                // 优先判断点是否在包围盒内
                try
                {
                    var ext = br.GeometricExtents;
                    inside = IsPointInsideExtents(point, ext, tol);
                }
                catch
                {
                    // 包围盒失败时保持 inside=false，后续走距离兜底
                }

                // 命中条件：包围盒命中，或插入点足够近
                if (inside || dist <= tol * 2.0)
                {
                    hits.Add((br, inside, dist));
                }
            }

            // 排序规则：包围盒命中优先，再按距离升序
            var ordered = hits
                .OrderByDescending(x => x.inside)
                .ThenBy(x => x.dist)
                .Select(x => x.br)
                .ToList();

            // 返回候选列表
            return ordered;
        }

        /// <summary>
        /// 从属性字典按候选键提取第一个非空值（先精确，再包含）
        /// </summary>
        //private static string FindFirstAttrValue(Dictionary<string, string> attrs, string[] keys)
        //{
        //    // 空字典保护
        //    if (attrs == null || attrs.Count == 0) return string.Empty;

        //    // 先做精确键匹配
        //    foreach (var k in keys)
        //    {
        //        if (attrs.TryGetValue(k, out var v) && !string.IsNullOrWhiteSpace(v))
        //            return v.Trim();
        //    }

        //    // 再做包含匹配
        //    foreach (var kv in attrs)
        //    {
        //        if (string.IsNullOrWhiteSpace(kv.Key) || string.IsNullOrWhiteSpace(kv.Value)) continue;
        //        foreach (var k in keys)
        //        {
        //            if (kv.Key.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
        //                return kv.Value.Trim();
        //        }
        //    }

        //    // 未命中
        //    return string.Empty;
        //}

        #endregion

        //SetOrAddAttrByAliases DrawPipeByClicks NormalizePipeAttributeKeys


        /// <summary>
        /// 辅助：把属性定义中存在或不存在的 Tag 设置/新增值
        /// </summary>
        private void SetOrAddAttr(List<AttributeDefinition> attDefs, string tag, string text, ref int extraIndex, double yOffsetBase, double attHeight)
        {
            var existing = attDefs.FirstOrDefault(a => string.Equals(a.Tag, tag, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.TextString = text;
                existing.Invisible = false;
                existing.Constant = false;
            }
            else
            {
                attDefs.Add(new AttributeDefinition
                {
                    Tag = tag,
                    Position = new Point3d(0, yOffsetBase - extraIndex * attHeight * 1.2, 0),
                    Rotation = 0.0,
                    TextString = text,
                    Height = attHeight,
                    Invisible = false,
                    Constant = false
                });
                extraIndex++;
            }
        }

        /// <summary>
        /// 根据采集点计算管线长度（绘图单位），并提供 Polyline 兜底计算。
        /// </summary>
        private double ComputePipelineLengthByPoints(List<Point3d> points)
        {
            if (points == null || points.Count < 2) return 0.0;

            double total = 0.0;
            for (int i = 0; i < points.Count - 1; i++)
            {
                total += points[i].DistanceTo(points[i + 1]);
            }

            // 正常累加结果可用
            if (total > 0.0) return total;

            // 兜底：用临时 Polyline 再算一次
            try
            {
                using (var pl = new Polyline())
                {
                    for (int i = 0; i < points.Count; i++)
                    {
                        pl.AddVertexAt(i, new Point2d(points[i].X, points[i].Y), 0, 0, 0);
                    }
                    return pl.Length;
                }
            }
            catch
            {
                return 0.0;
            }
        }

        /// <summary>
        /// 回写“长度”相关属性：优先更新已有字段；若不存在则新增“长度(mm)”和“Length(mm)”。
        /// </summary>
        private void SetOrAddLengthAttrs(
            List<AttributeDefinition> attDefs,
            double pipelineLength,
            ref int extraIndex,
            double yOffsetBase,
            double attHeight)
        {
            if (attDefs == null) return;

            // 统一格式：默认按绘图单位（你的表头是 mm，所以这里按 mm 写）
            string lenMmText = pipelineLength.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);

            bool updatedAny = false;
            foreach (var def in attDefs)
            {
                if (def == null || string.IsNullOrWhiteSpace(def.Tag)) continue;

                string tag = def.Tag.Trim();
                bool isLengthTag =
                    tag.IndexOf("长度", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    tag.IndexOf("length", StringComparison.OrdinalIgnoreCase) >= 0;

                if (!isLengthTag) continue;

                // 仅写“长度”类字段，不动管段号等字段
                def.TextString = lenMmText;
                def.Invisible = false;
                def.Constant = false;
                updatedAny = true;
            }

            // 如果模板里没有任何“长度”字段，则新增两个常用字段
            if (!updatedAny)
            {
                attDefs.Add(new AttributeDefinition
                {
                    Tag = "长度(mm)",
                    Position = new Point3d(0, yOffsetBase - extraIndex * attHeight * 1.2, 0),
                    Rotation = 0.0,
                    TextString = lenMmText,
                    Height = attHeight,
                    Invisible = false,
                    Constant = false
                });
                extraIndex++;

                attDefs.Add(new AttributeDefinition
                {
                    Tag = "Length(mm)",
                    Position = new Point3d(0, yOffsetBase - extraIndex * attHeight * 1.2, 0),
                    Rotation = 0.0,
                    TextString = lenMmText,
                    Height = attHeight,
                    Invisible = false,
                    Constant = false
                });
                extraIndex++;
            }
        }

        /// <summary>
        /// 仅写入标准字段（起点/终点/管段号），并清理别名字段，避免重复
        /// </summary>
        private void SetOrAddCanonicalPipeAttrs(
            List<AttributeDefinition> attDefs,
            string startText,
            string endText,
            string pipeNoText,
            ref int extraIndex,
            double yOffsetBase,
            double attHeight)
        {
            // 参数保护
            if (attDefs == null) return;

            // 清理旧别名字段，避免重复
            attDefs.RemoveAll(a =>
                a != null &&
                !string.IsNullOrWhiteSpace(a.Tag) &&
                (string.Equals(a.Tag, "始点", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(a.Tag, "起始", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(a.Tag, "管道号", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(a.Tag, "管段编号", StringComparison.OrdinalIgnoreCase)));

            // 统一只写标准字段
            SetOrAddAttr(attDefs, "起点", (startText ?? string.Empty).Trim(), ref extraIndex, yOffsetBase, attHeight);
            SetOrAddAttr(attDefs, "终点", (endText ?? string.Empty).Trim(), ref extraIndex, yOffsetBase, attHeight);
            SetOrAddAttr(attDefs, "管段号", (pipeNoText ?? string.Empty).Trim(), ref extraIndex, yOffsetBase, attHeight);
        }

        /// <summary>
        /// 辅助：根据点列表构建按顺序的线段信息（仅用于方向聚合计算）
        /// </summary>
        //private List<LineSegmentInfo> BuildOrderedLineSegmentsFromPoints(List<Point3d> pts)
        //{
        //    var segs = new List<LineSegmentInfo>();
        //    for (int i = 0; i < pts.Count - 1; i++)
        //    {
        //        segs.Add(new LineSegmentInfo
        //        {
        //            StartPoint = pts[i],
        //            EndPoint = pts[i + 1],
        //            Length = pts[i].DistanceTo(pts[i + 1]),
        //            Angle = ComputeSegmentAngleUcs(pts[i], pts[i + 1])
        //        });
        //    }
        //    return segs;
        //}

        /// <summary>
        /// 新增辅助：为每段生成方向三角和管道标题文字（放在 deviceTableGenerator 类内，放置在 BuildOrderedLineSegmentsFromPoints 之后）
        /// </summary>
        /// <param name="sampleInfo"></param>
        /// <param name="verticesWorld"></param>
        /// <param name="pipeTitle"></param>
        /// <param name="sampleBlockName"></param>
        /// <returns></returns>
        //private List<Entity> CreateDirectionalArrowsAndTitles(DBTrans tr, SamplePipeInfo sampleInfo, List<Point3d> verticesWorld, Point3d midPointWorld, string pipeTitle, string sampleBlockName)
        //{
        //    var overlay = new List<Entity>();
        //    if (sampleInfo == null || verticesWorld == null || verticesWorld.Count < 2) return overlay;

        //    // 优先使用用户在TextBox_绘图比例中设置的比例值
        //    var scaleDenom = VariableDictionary.textBoxScale;
        //    if (scaleDenom <= 0) // 如果获取失败，使用原有逻辑
        //    {
        //        AutoCadHelper.GetAndApplyActiveDrawingScale();//获取当前绘图比例
        //        scaleDenom = VariableDictionary.blockScale;
        //    }
        //    // 箭头模板与填充准备：若无模板则用默认三角di 
        //    Polyline arrowTemplate = sampleInfo.DirectionArrowTemplate;
        //    Solid? fillTemplate = null;
        //    double explicitArrowLength = 10.0;
        //    double explicitArrowHeight = 3.0;
        //    if (arrowTemplate == null)
        //    {
        //        var (colorIdx, length, height) = DetermineArrowStyleByName(sampleBlockName);
        //        explicitArrowLength = length;
        //        explicitArrowHeight = height;
        //        var (outline, fill) = CreateArrowTriangleFilled(length, height, colorIdx, sampleInfo.PipeBodyTemplate);
        //        arrowTemplate = outline;
        //        fillTemplate = fill;
        //    }            

        //    // 标题最终高度：基准 3.5 * 比例分母（与表格一致）
        //    double finalTitleHeight = FontsStyleHelper.ComputeScaledHeight(3.5, scaleDenom);

        //    // 遍历每一段，生成箭头并在箭头“上方”放置居中对齐的标题文字
        //    for (int i = 0; i < verticesWorld.Count - 1; i++)
        //    {
        //        var p1 = verticesWorld[i];
        //        var p2 = verticesWorld[i + 1];
        //        var seg = p2 - p1;
        //        if (seg.IsZeroLength()) continue;

        //        var dir = seg.GetNormal();
        //        var mid = new Point3d((p1.X + p2.X) / 2.0, (p1.Y + p2.Y) / 2.0, (p1.Z + p2.Z) / 2.0);

        //        Polyline? outlineAligned = null;
        //        Solid? fillAligned = null;
        //        try
        //        {
        //            // 箭头模板对齐
        //            (outlineAligned, fillAligned) = AlignArrowToDirection(arrowTemplate, fillTemplate, dir);
        //            // 箭头模板平移
        //            var localDisp = mid - midPointWorld;
        //            if (outlineAligned != null)
        //            {
        //                // 箭头模板平移
        //                outlineAligned.TransformBy(Matrix3d.Displacement(new Vector3d(localDisp.X, localDisp.Y, localDisp.Z)));
        //                // 箭头模板设置图层
        //                outlineAligned.Layer = sampleInfo.PipeBodyTemplate.Layer;
        //                // 箭头模板添加到 overlay
        //                overlay.Add(outlineAligned);
        //            }
        //            if (fillAligned != null)//填充
        //            {
        //                // 填充模板平移
        //                fillAligned.TransformBy(Matrix3d.Displacement(new Vector3d(localDisp.X, localDisp.Y, localDisp.Z)));
        //                // 填充模板设置图层
        //                fillAligned.Layer = sampleInfo.PipeBodyTemplate.Layer;
        //                overlay.Add(fillAligned);
        //            }
        //        }
        //        catch
        //        {
        //            // 忽略箭头生成异常，继续生成标题
        //        }

        //        try
        //        {
        //            // 计算文字放置方向：取段法线的+90度方向作为“上方”
        //            var perp = new Vector3d(-dir.Y, dir.X, 0.0);
        //            if (perp.IsZeroLength())
        //                perp = Vector3d.YAxis;
        //            else
        //                perp = perp.GetNormal();

        //            // 确保 perp 指向图纸上侧（全局 +Y）
        //            if (perp.DotProduct(Vector3d.YAxis) < 0)
        //                perp = -perp;

        //            // 估算箭头半高以确定文字偏移，优先使用已对齐实体的几何包围盒
        //            double arrowHalfHeight = explicitArrowHeight / 2.0;
        //            try
        //            {
        //                // 获取实体尺寸
        //                Entity sizeEntity = (Entity?)outlineAligned ?? (Entity?)fillAligned;
        //                if (sizeEntity != null)
        //                {
        //                    var ext = sizeEntity.GeometricExtents;// 获取实体尺寸
        //                    arrowHalfHeight = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y) / 2.0;// 计算箭头半高
        //                    if (arrowHalfHeight < 1e-6) arrowHalfHeight = explicitArrowHeight / 2.0;// 如果获取尺寸失败，使用默认值
        //                }
        //            }
        //            catch { arrowHalfHeight = explicitArrowHeight / 2.0; }// 如果获取尺寸失败，使用默认值

        //            // 文字偏移：箭头上方 + 与文字高度相关的间距
        //            double offset = arrowHalfHeight + finalTitleHeight * 0.8;
        //            var worldTextPos = mid + perp * offset;
        //            var localTextPos = new Point3d(worldTextPos.X - midPointWorld.X, worldTextPos.Y - midPointWorld.Y, worldTextPos.Z - midPointWorld.Z);

        //            // 文字方向：沿段方向，保证可读（不倒置）
        //            double segAngle = ComputeSegmentAngleUcs(p1, p2);
        //            double textRot = segAngle;
        //            if (Math.Cos(textRot) < 0) textRot += Math.PI;
        //            if (textRot > Math.PI) textRot -= 2.0 * Math.PI;
        //            if (textRot <= -Math.PI) textRot += 2.0 * Math.PI;

        //            // 创建 DBText 并设置为居中对齐
        //            var dbText = new DBText
        //            {
        //                Position = localTextPos,
        //                Height = finalTitleHeight,
        //                TextString = string.IsNullOrWhiteSpace(pipeTitle) ? sampleBlockName ?? "管道" : pipeTitle,
        //                Rotation = textRot,
        //                Layer = sampleInfo.PipeBodyTemplate.Layer,
        //                Normal = Vector3d.ZAxis,
        //                Oblique = 0.0
        //            };

        //            // 设置对齐点并置中（水平 + 垂直）
        //            try
        //            {
        //                dbText.AlignmentPoint = localTextPos;
        //                dbText.HorizontalMode = TextHorizontalMode.TextCenter;
        //                dbText.VerticalMode = TextVerticalMode.TextVerticalMid;
        //            }
        //            catch
        //            {
        //                // 某些 API/版本对这些属性有限制，忽略异常
        //            }

        //            // 应用样式并保证高度按当前比例（FontsStyleHelper 内部也会确保 TextStyle 存在）
        //            try
        //            {
        //                FontsStyleHelper.ApplyTitleToDBText(tr, dbText, scaleDenom);
        //            }
        //            catch
        //            {
        //                // 若样式应用失败，仍使用 dbText 的 Height
        //            }

        //            overlay.Add(dbText);
        //        }
        //        catch
        //        {
        //            // 忽略该段文字生成异常
        //        }
        //    }

        //    return overlay;

        //}

        private List<Entity> CreateDirectionalArrowsAndTitles(DBTrans tr, SamplePipeInfo sampleInfo, List<Point3d> verticesWorld, Point3d midPointWorld, string pipeTitle, string sampleBlockName)
        {
            var overlay = new List<Entity>();
            if (sampleInfo == null || verticesWorld == null || verticesWorld.Count < 2) return overlay;

            // 优先使用用户在TextBox_绘图比例中设置的比例值
            var scaleFactor = VariableDictionary.textBoxScale;
            if (scaleFactor <= 0) // 如果获取失败，使用原有逻辑
            {
                //AutoCadHelper.GetAndApplyActiveDrawingScale();//获取当前绘图比例
                scaleFactor = AutoCadHelper.GetAndApplyActiveDrawingScale();//获取当前绘图比例
            }

            // 箭头模板与填充准备：若无模板则用默认三角
            Polyline arrowTemplate = sampleInfo.DirectionArrowTemplate;
            Solid? fillTemplate = null;
            double explicitArrowLength = 8;  // 基础长度
            double explicitArrowHeight = 2.0;   // 基础高度
            if (arrowTemplate == null)
            {
                // 根据名称确定箭头样式
                var (colorIdx, length, height) = DetermineArrowStyleByName(sampleBlockName);
                explicitArrowLength = length * scaleFactor;  // 应用比例
                explicitArrowHeight = height * scaleFactor;  // 应用比例
                // 创建箭头
                var (outline, fill) = CreateArrowTriangleFilled(explicitArrowLength, explicitArrowHeight, colorIdx, sampleInfo.PipeBodyTemplate);
                arrowTemplate = outline;
                fillTemplate = fill;
            }
            else
            {
                // 如果有模板箭头，也按比例缩放
                try
                {
                    if (scaleFactor != 1.0)
                    {
                        // 克隆模板并按比例缩放
                        arrowTemplate = (Polyline)arrowTemplate.Clone();
                        Matrix3d scaleMatrix = Matrix3d.Scaling(scaleFactor, Point3d.Origin);
                        arrowTemplate.TransformBy(scaleMatrix);

                        if (fillTemplate != null)
                        {
                            fillTemplate = (Solid)fillTemplate.Clone();
                            fillTemplate.TransformBy(scaleMatrix);
                        }
                    }
                }
                catch
                {
                    // 如果缩放失败，使用原始模板
                }
            }

            // 标题最终高度：基准 3.5 * 比例分母（与表格一致）
            double finalTitleHeight = TextFontsStyleHelper.ComputeScaledHeight(4, scaleFactor);

            // 遍历每一段，生成箭头并在箭头"上方"放置居中对齐的标题文字
            for (int i = 0; i < verticesWorld.Count - 1; i++)
            {
                var p1 = verticesWorld[i];
                var p2 = verticesWorld[i + 1];
                var seg = p2 - p1;
                if (seg.IsZeroLength()) continue;

                var dir = seg.GetNormal();
                var mid = new Point3d((p1.X + p2.X) / 2.0, (p1.Y + p2.Y) / 2.0, (p1.Z + p2.Z) / 2.0);

                Polyline? outlineAligned = null;
                Solid? fillAligned = null;
                try
                {
                    // 箭头模板对齐（注意：这里只做旋转，缩放已在上面处理）
                    (outlineAligned, fillAligned) = AlignArrowToDirection(arrowTemplate, fillTemplate, dir);

                    // 箭头模板平移
                    var localDisp = mid - midPointWorld;
                    if (outlineAligned != null)
                    {
                        // 箭头模板平移
                        outlineAligned.TransformBy(Matrix3d.Displacement(new Vector3d(localDisp.X, localDisp.Y, localDisp.Z)));
                        // 箭头模板设置图层
                        outlineAligned.Layer = sampleInfo.PipeBodyTemplate.Layer;
                        // 箭头模板添加到 overlay
                        overlay.Add(outlineAligned);
                    }
                    if (fillAligned != null)//填充
                    {
                        // 填充模板平移
                        fillAligned.TransformBy(Matrix3d.Displacement(new Vector3d(localDisp.X, localDisp.Y, localDisp.Z)));
                        // 填充模板设置图层
                        fillAligned.Layer = sampleInfo.PipeBodyTemplate.Layer;
                        overlay.Add(fillAligned);
                    }
                }
                catch
                {
                    // 忽略箭头生成异常，继续生成标题
                }

                try
                {
                    // 计算文字放置方向：取段法线的+90度方向作为"上方"
                    var perp = new Vector3d(-dir.Y, dir.X, 0.0);
                    if (perp.IsZeroLength())
                        perp = Vector3d.YAxis;
                    else
                        perp = perp.GetNormal();

                    // 确保 perp 指向图纸上侧（全局 +Y）
                    if (perp.DotProduct(Vector3d.YAxis) < 0)
                        perp = -perp;

                    // 估算箭头半高以确定文字偏移，优先使用已对齐实体的几何包围盒
                    double arrowHalfHeight = explicitArrowHeight / 2.0;
                    try
                    {
                        // 获取实体尺寸
                        Entity sizeEntity = (Entity?)outlineAligned ?? (Entity?)fillAligned;
                        if (sizeEntity != null)
                        {
                            var ext = sizeEntity.GeometricExtents;// 获取实体尺寸
                            arrowHalfHeight = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y) / 2.0;// 计算箭头半高
                            if (arrowHalfHeight < 1e-6) arrowHalfHeight = explicitArrowHeight / 2.0;// 如果获取尺寸失败，使用默认值
                        }
                    }
                    catch { arrowHalfHeight = explicitArrowHeight / 2.0; }// 如果获取尺寸失败，使用默认值

                    // 文字偏移：箭头上方 + 与文字高度相关的间距（按比例调整）
                    //double offset = (arrowHalfHeight + finalTitleHeight * 0.8) * scaleFactor; // 应用比例
                    double offset = (400 + finalTitleHeight * 0.75); // 应用比例
                    var worldTextPos = mid + perp * offset;// 文字放在箭头上方一定距离处
                    // 计算文字的局部坐标位置（相对于 midPointWorld）
                    var localTextPos = new Point3d(worldTextPos.X - midPointWorld.X, worldTextPos.Y - midPointWorld.Y, worldTextPos.Z - midPointWorld.Z);

                    // 文字方向：沿段方向，保证可读（不倒置）
                    double segAngle = ComputeSegmentAngleUcs(p1, p2);
                    double textRot = segAngle;
                    if (Math.Cos(textRot) < 0) textRot += Math.PI;
                    if (textRot > Math.PI) textRot -= 2.0 * Math.PI;
                    if (textRot <= -Math.PI) textRot += 2.0 * Math.PI;

                    // 创建标题文字对象
                    var dbText = new DBText
                    {
                        // 文字实际内容，优先使用管道标题，没有则回退到块名，再没有则显示“管道”
                        TextString = string.IsNullOrWhiteSpace(pipeTitle) ? sampleBlockName ?? "管道" : pipeTitle,

                        // 设置文字高度
                        Height = finalTitleHeight,

                        // 这里先给 Position 一个值，作为兼容性兜底
                        Position = localTextPos,

                        // 设置文字旋转角度，使文字沿管段方向显示
                        Rotation = textRot,

                        // 设置图层，仍然跟随管道主体图层
                        Layer = sampleInfo.PipeBodyTemplate.Layer,

                        // 设置法向量，保持文字位于当前 XY 平面
                        Normal = Vector3d.ZAxis,

                        // 设置倾斜角为 0，不做斜体处理
                        Oblique = 0.0,

                        // 根据图层名设置颜色
                        Color = sampleInfo.PipeBodyTemplate.Layer.Contains("进口") ? Color.FromColorIndex(ColorMethod.ByAci, 1) :
                                sampleInfo.PipeBodyTemplate.Layer.Contains("出口") ? Color.FromColorIndex(ColorMethod.ByAci, 2) :
                                sampleInfo.PipeBodyTemplate.Color,

                        // 关键设置——把文字对齐方式改成“中间居中”
                        Justify = AttachmentPoint.MiddleCenter,

                        // 关键设置——让文字的“中心点”对齐到目标点，而不是首字符落点对齐
                        AlignmentPoint = localTextPos
                    };

                    // 再次显式设置水平居中，增强兼容性
                    dbText.HorizontalMode = TextHorizontalMode.TextCenter;

                    // 再次显式设置垂直居中，增强兼容性
                    dbText.VerticalMode = TextVerticalMode.TextVerticalMid;

                    // 先应用您项目里的标题文字样式（文字样式、高度、注释性等）
                    try
                    {
                        TextFontsStyleHelper.ApplyTitleToDBText(tr, dbText, scaleFactor);
                    }
                    catch
                    {
                        // 若样式应用失败，则保留当前 DBText 基本设置继续执行
                    }

                    // 非常关键——让 AutoCAD 根据 Justify 和 AlignmentPoint 重新计算文字位置
                    try
                    {
                        dbText.AdjustAlignment(tr.Database);
                    }
                    catch
                    {
                        // 某些场景下对象尚未加入数据库，可能会失败，这里忽略异常即可
                    }

                    overlay.Add(dbText);
                }
                catch
                {
                    // 忽略该段文字生成异常
                }
            }

            return overlay;
        }

        #endregion


        /// <summary>
        /// 辅助命令：列出选中动态块的所有可用属性
        /// </summary>
        [CommandMethod("LISTDYNPROPS")]
        public void ListDynamicBlockProperties()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 提示用户选择动态块
            PromptEntityOptions peo = new PromptEntityOptions("\n请选择动态块以查看其属性: ");
            peo.SetRejectMessage("\n只能选择块参照对象!");
            peo.AddAllowedClass(typeof(BlockReference), true);

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
                return;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 获取块参照对象
                BlockReference blockRef = trans.GetObject(per.ObjectId, OpenMode.ForRead) as BlockReference;

                if (blockRef != null && blockRef.IsDynamicBlock)
                {
                    ed.WriteMessage("\n========== 动态块属性列表 ==========");

                    // 遍历并显示所有动态属性
                    DynamicBlockReferencePropertyCollection dynProps = blockRef.DynamicBlockReferencePropertyCollection;

                    for (int i = 0; i < dynProps.Count; i++)
                    {
                        DynamicBlockReferenceProperty dynProp = dynProps[i];
                        ed.WriteMessage($"\n  属性 {i + 1}:");
                        ed.WriteMessage($"\n  名称: {dynProp.PropertyName}");
                        ed.WriteMessage($"\n  描述: {dynProp.Description}");
                        ed.WriteMessage($"\n  参数类型: {dynProp.Value.GetType().Name}");
                        ed.WriteMessage($"\n  当前值: {dynProp.Value}");
                        ed.WriteMessage($"\n  单位类型: {dynProp.UnitsType}");
                        ed.WriteMessage($"\n  是否只读: {dynProp.ReadOnly}");
                        ed.WriteMessage($"\n  是否可见: {dynProp.Show}");
                        ed.WriteMessage("\n" + new string('-', 30));
                    }
                }
                else
                {
                    ed.WriteMessage("\n所选对象不是动态块!");
                }

                trans.Commit();
            }
        }


        /// <summary>
        /// 根据设备列表构建规范化的动态列名列表（去重、排除、同义词归一化）
        /// 说明：
        /// - 从每个设备的 Attributes 收集原始键
        /// - 排除包含关键字的属性（如 图层/块名/比例/标题/角度/颜色/法兰/属性 等）
        /// - 使用 NormalizeAttributeKey 做同义词归一化（例如 阀体材料 -> 材料）
        /// - 返回稳定排序的列名列表
        /// </summary>
        private List<string> BuildDynamicColumnList(List<DeviceInfo> deviceList)
        {
            var excludeSubstrings = new[]
            {
                "管道标题","管段号","起点","始点","终点","止点","管道等级",
                "介质","介质名称","Medium","Medium Name","操作温度","操作压力",
                "隔热隔声代号","是否防腐","Length","长度","长度(m)",
                "图层","块名","比例","标题","角度","颜色","法兰","属性"
            };

            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var e in deviceList)
            {
                if (e?.Attributes == null) continue;
                foreach (var raw in e.Attributes.Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrWhiteSpace(raw)) continue;
                    if (excludeSubstrings.Any(s => !string.IsNullOrWhiteSpace(s) && raw.IndexOf(s, StringComparison.OrdinalIgnoreCase) >= 0))
                        continue;

                    // 忽略标准/图号同义词（这些由列名映射或专列处理时可能需要单独合并）
                    if (raw.IndexOf("标准", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        raw.IndexOf("图号", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        raw.IndexOf("DWG", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        raw.IndexOf("STD", StringComparison.OrdinalIgnoreCase) >= 0)
                        continue;

                    var mapped = NormalizeAttributeKey(raw);
                    if (string.IsNullOrWhiteSpace(mapped)) continue;
                    if (!set.Contains(mapped))
                        set.Add(mapped);
                }
            }

            // 保证稳定排序：先按常用列顺序，其余按字母序
            var preferred = new[] { "名称", "规格", "图号或标准号", "材料", "数量", "介质名称" };
            var result = new List<string>();

            foreach (var p in preferred)
            {
                if (set.Contains(p))
                {
                    result.Add(p);
                    set.Remove(p);
                }
            }

            var rest = set.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
            result.AddRange(rest);

            // 至少保留一列（用于避免空表）
            if (result.Count == 0) result.Add("名称");

            return result;
        }

        #region 表与图元的属性映射与同步辅助方法


        /// <summary>
        /// 同步设备表到块（表 -> 图元 / 图元 -> 表 / 双向(表 <-> 图元)）
        /// </summary>
        [CommandMethod(nameof(SyncdeviceTableToBlocks))]
        public void SyncdeviceTableToBlocks()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            try
            {
                var result = MessageBox.Show(
                    "请选择同步方向：\n\n" +
                    "A. 表 -> 图元：将表格数据同步到选中的块\n" +
                    "B. 图元 -> 表：将选中块的属性同步到表格\n" +
                    "C. 双向同步：表格和块互相同步",
                    "同步选项",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Cancel) return;

                if (result == DialogResult.Yes) // 表 -> 图元
                {
                    SyncTableToBlocks();
                }
                else if (result == DialogResult.No) // 图元 -> 表
                {
                    SyncBlocksToTable();
                }
                else // 双向
                {
                    SyncTableToBlocks();
                    SyncBlocksToTable();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n同步设备表到块时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 将表格数据同步到选中的块（表 -> 图元）
        /// </summary>
        private void SyncTableToBlocks()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                ed.WriteMessage("\n请选择要同步的表格对象...");
                var peo = new PromptEntityOptions("\n请选择表格: ");
                peo.SetRejectMessage("\n必须选择表格对象。");
                peo.AddAllowedClass(typeof(Table), true);

                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var table = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table;
                    if (table == null)
                    {
                        ed.WriteMessage("\n选择的对象不是表格。");
                        return;
                    }

                    // 常见表头结构：0=title,1=headerRow1,2=headerRow2, 数据从第3行开始
                    int dataStartRow = 3;
                    if (table.Rows.Count <= dataStartRow)
                    {
                        ed.WriteMessage("\n表格行数过少，无法识别数据行，请确认表格结构。");
                        return;
                    }

                    // 读取列头（优先 row2，再 row1），并构建 header list
                    int cols = table.Columns.Count;
                    var headers = new List<string>(cols);
                    for (int c = 0; c < cols; c++)
                    {
                        string h2 = (table.Cells[2, c].TextString ?? string.Empty).Trim();
                        string h1 = (table.Cells[1, c].TextString ?? string.Empty).Trim();
                        string header = string.IsNullOrWhiteSpace(h2) ? h1 : h2;
                        header = header.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? header;
                        headers.Add(header);
                    }

                    // 寻找标识列
                    int idCol = -1;
                    string[] idKeywords = new[] { "部件ID", "部件 Id", "部件编号", "管段号", "管段编号", "Pipeline", "Pipe No", "ID", "序号" };
                    for (int c = 0; c < headers.Count; c++)
                    {
                        var header = headers[c];
                        if (string.IsNullOrWhiteSpace(header)) continue;
                        foreach (var kw in idKeywords)
                        {
                            if (header.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                idCol = c;
                                break;
                            }
                        }
                        if (idCol != -1) break;
                    }

                    // 将表中每行构建为 header->value 字典
                    var tableRows = new List<Dictionary<string, string>>();
                    for (int r = dataStartRow; r < table.Rows.Count; r++)
                    {
                        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        for (int c = 0; c < cols; c++)
                        {
                            try
                            {
                                string val = table.Cells[r, c].TextString ?? string.Empty;
                                var key = headers[c] ?? $"Col{c}";
                                if (!dict.ContainsKey(key)) dict[key] = val;
                            }
                            catch { }
                        }
                        tableRows.Add(dict);
                    }

                    // 选择要同步到的块参照（用户选择）
                    ed.WriteMessage("\n请选择要同步到的块参照(多选)：");
                    var pso = new PromptSelectionOptions { MessageForAdding = "\n请选择块参照: " };
                    var filter = new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "INSERT") });
                    var psr = ed.GetSelection(pso, filter);
                    if (psr.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\n未选择块参照或已取消。");
                        return;
                    }

                    var selectedBlockIds = psr.Value.GetObjectIds().ToList();
                    if (selectedBlockIds.Count == 0)
                    {
                        ed.WriteMessage("\n未选择任何块参照。");
                        return;
                    }

                    // 不可回写字段（规范化后）
                    var nonWritable = new[]
                    {
                "起点","终点","始点","止点","起止点","起止","位置","坐标","方向","角度",
                "管段号","管段编号","管道标题","长度","长度(m)","累计长度","累计长度(mm)"
            }.Select(k => NormalizeAttributeKey(k)).Where(x => !string.IsNullOrWhiteSpace(x)).ToHashSet(StringComparer.OrdinalIgnoreCase);

                    // 为快速匹配，建立选中块的属性快照（Tag->value），并同时尝试从属性中找到 ID 值（若有）
                    var blockIdToAttrMap = new Dictionary<ObjectId, Dictionary<string, string>>();
                    var valueToBlockIds = new Dictionary<string, List<ObjectId>>(StringComparer.OrdinalIgnoreCase); // idValue -> blocks
                    foreach (var bid in selectedBlockIds)
                    {
                        try
                        {
                            var br = tr.GetObject(bid, OpenMode.ForRead) as BlockReference;
                            if (br == null) continue;
                            var attrs = GetEntityAttributeMap(tr, br) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            blockIdToAttrMap[bid] = attrs;

                            // 尝试从属性中找到一个候选 id 值（基于 idKeywords 或 部件ID 标签）
                            foreach (var kv in attrs)
                            {
                                if (string.IsNullOrWhiteSpace(kv.Key) || string.IsNullOrWhiteSpace(kv.Value)) continue;
                                // 若键名包含 idKeywords，则把该值作为候选
                                if (idKeywords.Any(k => kv.Key.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0) ||
                                    kv.Key.Equals("部件ID", StringComparison.OrdinalIgnoreCase) ||
                                    kv.Key.Equals("部件 Id", StringComparison.OrdinalIgnoreCase))
                                {
                                    var v = kv.Value.Trim();
                                    if (!string.IsNullOrWhiteSpace(v))
                                    {
                                        if (!valueToBlockIds.ContainsKey(v)) valueToBlockIds[v] = new List<ObjectId>();
                                        valueToBlockIds[v].Add(bid);
                                    }
                                }
                            }
                        }
                        catch { }
                    }

                    int updatedCount = 0;
                    var unmatched = new List<string>();

                    // Helper: 写回单个块的单个字段（考虑动态属性 & AttributeReference）
                    void WriteValueToBlock(BlockReference brWrite, string headerRaw, string newVal)
                    {
                        if (brWrite == null) return;
                        string headerLine = headerRaw?.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? headerRaw;
                        var normalizedHeader = NormalizeAttributeKey(headerLine);
                        if (!string.IsNullOrWhiteSpace(normalizedHeader) && nonWritable.Contains(normalizedHeader)) return;

                        string cleaned = CleanAttributeText(newVal ?? string.Empty);

                        bool matched = false;

                        // 1) 动态属性
                        try
                        {
                            if (brWrite.IsDynamicBlock)
                            {
                                var dyn = brWrite.DynamicBlockReferencePropertyCollection;
                                foreach (DynamicBlockReferenceProperty prop in dyn)
                                {
                                    try
                                    {
                                        var propNorm = NormalizeAttributeKey(prop.PropertyName ?? string.Empty);
                                        if (string.Equals(propNorm, normalizedHeader, StringComparison.OrdinalIgnoreCase) ||
                                            (prop.PropertyName ?? string.Empty).IndexOf(headerLine, StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            try
                                            {
                                                var targetType = prop.Value?.GetType() ?? typeof(string);
                                                object conv;
                                                if (targetType == typeof(string))
                                                    conv = cleaned;
                                                else
                                                    conv = Convert.ChangeType(cleaned, targetType, System.Globalization.CultureInfo.InvariantCulture);
                                                prop.Value = conv;
                                            }
                                            catch
                                            {
                                                try { prop.Value = cleaned; } catch { }
                                            }
                                            matched = true;
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                        catch { /* ignore */ }

                        // 2) 普通 AttributeReference
                        try
                        {
                            foreach (ObjectId aid in brWrite.AttributeCollection)
                            {
                                try
                                {
                                    var ar = tr.GetObject(aid, OpenMode.ForWrite) as AttributeReference;
                                    if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue;
                                    var tagNorm = NormalizeAttributeKey(ar.Tag ?? string.Empty);
                                    if (string.Equals(tagNorm, normalizedHeader, StringComparison.OrdinalIgnoreCase) ||
                                        string.Equals(ar.Tag, headerLine, StringComparison.OrdinalIgnoreCase) ||
                                        ar.Tag.IndexOf(headerLine, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                        normalizedHeader.IndexOf(tagNorm ?? string.Empty, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        ar.TextString = cleaned;
                                        try { ar.AdjustAlignment(db); } catch { }
                                        matched = true;
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }

                        if (!matched)
                        {
                            // 3) 宽松包含匹配（尝试）
                            try
                            {
                                foreach (ObjectId aid in brWrite.AttributeCollection)
                                {
                                    try
                                    {
                                        var ar = tr.GetObject(aid, OpenMode.ForWrite) as AttributeReference;
                                        if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue;
                                        if (!string.IsNullOrWhiteSpace(headerLine) &&
                                            (ar.Tag.IndexOf(headerLine, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                             headerLine.IndexOf(ar.Tag, StringComparison.OrdinalIgnoreCase) >= 0))
                                        {
                                            ar.TextString = cleaned;
                                            try { ar.AdjustAlignment(db); } catch { }
                                            matched = true;
                                            break;
                                        }
                                    }
                                    catch { }
                                }
                            }
                            catch { }
                        }
                    }

                    // 主匹配逻辑：按表行遍历，找到要更新的块（优先 ID 列匹配，其次按序或宽松匹配）
                    for (int i = 0; i < tableRows.Count; i++)
                    {
                        var rowDict = tableRows[i];
                        string idValue = string.Empty;
                        if (idCol >= 0)
                        {
                            var idKey = headers[idCol];
                            if (!string.IsNullOrWhiteSpace(idKey) && rowDict.TryGetValue(idKey, out var tval))
                                idValue = (tval ?? string.Empty).Trim();
                        }

                        List<ObjectId> matchedBlocks = new List<ObjectId>();

                        if (!string.IsNullOrWhiteSpace(idValue) && valueToBlockIds.TryGetValue(idValue, out var bls))
                        {
                            matchedBlocks.AddRange(bls);
                        }

                        // 如果没有通过 ID 匹配，且选中块数 == 表行数，则按顺序映射（允许用户先手动选中目标块）
                        if (matchedBlocks.Count == 0 && selectedBlockIds.Count == tableRows.Count)
                        {
                            var bid = selectedBlockIds[i];
                            matchedBlocks.Add(bid);
                        }

                        // 回退：尝试宽松在每个选中块里查找任意属性值等于 idValue 或与表行里关键列匹配
                        if (matchedBlocks.Count == 0 && !string.IsNullOrWhiteSpace(idValue))
                        {
                            foreach (var kvp in blockIdToAttrMap)
                            {
                                if (kvp.Value.Values.Any(v => string.Equals(v?.Trim(), idValue, StringComparison.OrdinalIgnoreCase)))
                                    matchedBlocks.Add(kvp.Key);
                            }
                        }

                        // 最终如果仍然为空，则跳过该行并记录诊断
                        if (matchedBlocks.Count == 0)
                        {
                            unmatched.Add($"未匹配表行 {dataStartRow + i + 1} 的标识 '{idValue}' (可手动按顺序选择块以映射)。");
                            continue;
                        }

                        // 对每个匹配的块，写回所有列（跳过 idCol 与非可写字段）
                        foreach (var bid in matchedBlocks)
                        {
                            try
                            {
                                var brWrite = tr.GetObject(bid, OpenMode.ForWrite) as BlockReference;
                                if (brWrite == null) continue;

                                foreach (var kv in rowDict)
                                {
                                    var header = kv.Key;
                                    // 跳过标识列
                                    if (idCol >= 0 && string.Equals(header, headers[idCol], StringComparison.OrdinalIgnoreCase))
                                        continue;

                                    WriteValueToBlock(brWrite, header, kv.Value);
                                    updatedCount++;
                                }
                            }
                            catch { }
                        }
                    }

                    tr.Commit();

                    ed.WriteMessage($"\n同步完成，尝试更新属性项数（估计）: {updatedCount}。");
                    if (unmatched.Count > 0)
                    {
                        ed.WriteMessage("\n未匹配的表行样例（最多显示20条）：");
                        foreach (var s in unmatched.Take(20)) ed.WriteMessage("\n  " + s);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage($"\n同步失败: {ex.Message}");
            }
        }


        /// <summary>
        /// 将选中块的属性同步到表格（图元 -> 表）
        /// </summary>
        private void SyncBlocksToTable()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                ed.WriteMessage("\n请选择要同步的块参照...");
                var pso = new PromptSelectionOptions { MessageForAdding = "\n请选择块参照: " };
                var filter = new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "INSERT") });
                var psr = ed.GetSelection(pso, filter);
                if (psr.Status != PromptStatus.OK) return;

                var blockIds = psr.Value.GetObjectIds();
                if (blockIds == null || blockIds.Length == 0) return;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // 收集每个块的属性字典
                    var blockDataList = new List<Dictionary<string, string>>();
                    var idKeywords = new[] { "部件ID", "部件 Id", "部件编号", "管段号", "管段编号", "Pipeline", "Pipe No", "ID", "序号" };

                    foreach (var blockId in blockIds)
                    {
                        try
                        {
                            var br = tr.GetObject(blockId, OpenMode.ForRead) as BlockReference;
                            if (br == null) continue;
                            var attrs = ExtractBlockAttributes(br) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                            // 如果块没有明显的 key -> value，则尝试动态块属性
                            if ((attrs == null || attrs.Count == 0) && br.IsDynamicBlock)
                            {
                                try
                                {
                                    var dyn = br.DynamicBlockReferencePropertyCollection;
                                    foreach (DynamicBlockReferenceProperty p in dyn)
                                    {
                                        if (!string.IsNullOrWhiteSpace(p.PropertyName))
                                            attrs[p.PropertyName] = p.Value?.ToString() ?? string.Empty;
                                    }
                                }
                                catch { }
                            }

                            blockDataList.Add(attrs);
                        }
                        catch { }
                    }

                    ed.WriteMessage("\n请选择要同步到的表格...");
                    var peo = new PromptEntityOptions("\n请选择表格: ");
                    peo.SetRejectMessage("\n必须选择表格对象。");
                    peo.AddAllowedClass(typeof(Table), true);

                    var per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK) return;

                    var table = tr.GetObject(per.ObjectId, OpenMode.ForWrite) as Table;
                    if (table == null)
                    {
                        ed.WriteMessage("\n选择的对象不是表格。");
                        return;
                    }

                    // 表头读取：优先第2行，其次第1行
                    int cols = table.Columns.Count;
                    var headers = new List<string>(cols);
                    for (int c = 0; c < cols; c++)
                    {
                        string h2 = (table.Cells[2, c].TextString ?? string.Empty).Trim();
                        string h1 = (table.Cells[1, c].TextString ?? string.Empty).Trim();
                        string header = string.IsNullOrWhiteSpace(h2) ? h1 : h2;
                        header = header.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? header;
                        headers.Add(header);
                    }

                    // 寻找标识列
                    int dataStartRow = 3;
                    int idCol = -1;
                    for (int c = 0; c < headers.Count; c++)
                    {
                        if (string.IsNullOrWhiteSpace(headers[c])) continue;
                        foreach (var kw in idKeywords)
                        {
                            if (headers[c].IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                idCol = c;
                                break;
                            }
                        }
                        if (idCol != -1) break;
                    }

                    // 确保表有足够行以容纳数据（追加模式：在现有数据末尾追加）
                    int existingRows = table.Rows.Count;
                    int needRows = Math.Max(0, blockDataList.Count - Math.Max(0, existingRows - dataStartRow));
                    if (needRows > 0)
                    {
                        table.InsertRows(existingRows - 1, 1, needRows);
                    }

                    // 写入每个块的数据
                    for (int i = 0; i < blockDataList.Count; i++)
                    {
                        var data = blockDataList[i];
                        // 决定目标行：若表有 ID 列并且块提供了 ID 值，则尝试按 ID 找到行，否则按顺序追加/对应
                        int targetRow = dataStartRow + i; // 默认按顺序映射

                        if (idCol >= 0)
                        {
                            // 获取块的 id 值（先查 key 指定字段，再宽松匹配）
                            string idValue = string.Empty;
                            foreach (var k in idKeywords)
                            {
                                if (data.TryGetValue(k, out var v) && !string.IsNullOrWhiteSpace(v))
                                {
                                    idValue = v.Trim();
                                    break;
                                }
                            }
                            if (string.IsNullOrWhiteSpace(idValue))
                            {
                                // 宽松查找：键中含关键字
                                foreach (var kv in data)
                                {
                                    if (string.IsNullOrWhiteSpace(kv.Key) || string.IsNullOrWhiteSpace(kv.Value)) continue;
                                    if (idKeywords.Any(k => kv.Key.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                                    {
                                        idValue = kv.Value.Trim();
                                        break;
                                    }
                                }
                            }

                            if (!string.IsNullOrWhiteSpace(idValue))
                            {
                                // 查找表中已有行
                                bool found = false;
                                for (int r = dataStartRow; r < table.Rows.Count; r++)
                                {
                                    try
                                    {
                                        var cellTxt = (table.Cells[r, idCol].TextString ?? string.Empty).Trim();
                                        if (!string.IsNullOrWhiteSpace(cellTxt) && string.Equals(cellTxt, idValue, StringComparison.OrdinalIgnoreCase))
                                        {
                                            targetRow = r;
                                            found = true;
                                            break;
                                        }
                                    }
                                    catch { }
                                }
                                if (!found)
                                {
                                    // 如果没找到，则选择按顺序行（i），或在表尾追加
                                    int candidate = dataStartRow + i;
                                    if (candidate >= table.Rows.Count)
                                    {
                                        table.InsertRows(table.Rows.Count - 1, 1, 1);
                                    }
                                    targetRow = candidate;
                                }
                            }
                        }

                        // 写入每个表头列
                        for (int c = 0; c < headers.Count; c++)
                        {
                            string header = headers[c];
                            if (string.IsNullOrWhiteSpace(header)) continue;

                            try
                            {
                                // 忽略不可回写字段
                                var normalized = NormalizeAttributeKey(header);
                                var nonWritable = new[]
                                {
                            "起点","终点","始点","止点","起止点","起止","位置","坐标","方向","角度",
                            "管段号","管段编号","管道标题","长度","长度(m)","累计长度","累计长度(mm)"
                        }.Select(k => NormalizeAttributeKey(k)).Where(x => !string.IsNullOrWhiteSpace(x)).ToHashSet(StringComparer.OrdinalIgnoreCase);
                                if (!string.IsNullOrWhiteSpace(normalized) && nonWritable.Contains(normalized))
                                    continue;

                                // 尝试通过 GetAttributeValueByMappedKey 找到值（支持同义词/归一化）
                                string value = GetAttributeValueByMappedKey(data, NormalizeAttributeKey(header));
                                if (string.IsNullOrWhiteSpace(value))
                                {
                                    // 再尝试直接以 header 或包含匹配
                                    if (data.TryGetValue(header, out var v)) value = v;
                                    else
                                    {
                                        foreach (var kv in data)
                                        {
                                            if (string.IsNullOrWhiteSpace(kv.Key) || string.IsNullOrWhiteSpace(kv.Value)) continue;
                                            if (kv.Key.IndexOf(header, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                header.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                                            {
                                                value = kv.Value;
                                                break;
                                            }
                                        }
                                    }
                                }

                                if (!string.IsNullOrWhiteSpace(value))
                                {
                                    table.Cells[targetRow, c].TextString = value;
                                }
                            }
                            catch { }
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage($"\n成功将 {blockIds.Length} 个块参照的属性同步到表格（按表头匹配，已跳过几何/位置类不可回写字段）。");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n将块属性同步到表格时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 从表格中提取数据
        /// </summary>
        /// <param name="table">表格对象</param>
        /// <returns>表格数据字典</returns>
        //private Dictionary<string, string> ExtractTableData(Table table)
        //{
        //    var data = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //    try
        //    {
        //        // 遍历表格的行和列，提取数据
        //        for (int row = 0; row < table.Rows.Count; row++)
        //        {
        //            for (int col = 0; col < table.Columns.Count; col++)
        //            {
        //                var cellText = table.Cells[row, col].TextString ?? string.Empty;
        //                var headerText = table.Cells[0, col].TextString ?? string.Empty; // 使用第一行作为标题

        //                if (!string.IsNullOrWhiteSpace(headerText) && !string.IsNullOrWhiteSpace(cellText))
        //                {
        //                    // 避免重复键，使用组合键
        //                    var key = $"{headerText}_{row}";
        //                    if (!data.ContainsKey(key))
        //                    {
        //                        data[key] = cellText;
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        Application.DocumentManager.MdiActiveDocument?.Editor?.WriteMessage($"\n提取表格数据时出错: {ex.Message}");
        //    }

        //    return data;
        //}

        /// <summary>
        /// 从块参照中提取属性
        /// </summary>
        /// <param name="blockRef">块参照</param>
        /// <returns>属性数据字典</returns>
        private Dictionary<string, string> ExtractBlockAttributes(BlockReference blockRef)
        {
            var data = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // 遍历块参照的所有属性
                foreach (ObjectId attId in blockRef.AttributeCollection)
                {
                    var attRef = blockRef.Database.TransactionManager.TopTransaction.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                    if (attRef == null) continue;

                    var tag = attRef.Tag ?? string.Empty;
                    var value = attRef.TextString ?? string.Empty;

                    if (!string.IsNullOrWhiteSpace(tag))
                    {
                        data[tag] = value;
                    }
                }
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument?.Editor?.WriteMessage($"\n提取块属性时出错: {ex.Message}");
            }

            return data;
        }

        #endregion


        /// <summary>
        /// 属性同义词映射与过滤规则 NEW
        /// </summary>
        /// <param name="rawKey"></param>
        /// <returns></returns>
        private string NormalizeAttributeKey(string rawKey)
        {
            if (string.IsNullOrWhiteSpace(rawKey)) return string.Empty;

            // 1) 规范化：trim + 小写
            string k = rawKey.Trim();

            // 2) 去掉常见单位符号/括号/空白并小写（便于匹配 "T(℃)" / "T(°C)" 等）
            string normalized = k.ToLowerInvariant();
            normalized = Regex.Replace(normalized, @"[\s\[\]\(\){}]", ""); // 去空格与括号
            normalized = normalized.Replace("℃", "c").Replace("°c", "c").Replace("°", ""); // 规范度符号为 c
            normalized = normalized.Replace("（", "").Replace("）", "");

            // 3) 先尝试同义词字典匹配（包含匹配）
            foreach (var kv in DictionaryHelper.AttributeSynonyms)
            {
                string synKey = kv.Key;
                if (string.IsNullOrWhiteSpace(synKey)) continue;
                string synNorm = synKey.ToLowerInvariant();
                synNorm = Regex.Replace(synNorm, @"[\s\[\]\(\){}]", "");
                synNorm = synNorm.Replace("℃", "c").Replace("°c", "c").Replace("°", "");

                if (normalized.IndexOf(synNorm, StringComparison.OrdinalIgnoreCase) >= 0)
                    return kv.Value;
            }

            // 4) 回退：去掉单位后返回原始的修剪结果（便于直接比较）
            //    返回去单位/去括号后的原文（首字母大写或原样均可）
            var final = Regex.Replace(k, @"[\s\[\]\(\){}℃°]", "").Trim();
            return string.IsNullOrWhiteSpace(final) ? k : final;
        }

        /// <summary>
        /// 把管道属性字典按标准字段归一化（起点/终点/管段号）
        /// </summary>
        private Dictionary<string, string> NormalizePipeAttributeKeys(Dictionary<string, string> src)
        {
            // 创建输出字典
            var dst = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            // 空字典保护
            if (src == null || src.Count == 0) return dst;

            // 先整体拷贝
            foreach (var kv in src)
            {
                string k = (kv.Key ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(k)) continue;
                dst[k] = (kv.Value ?? string.Empty).Trim();
            }

            // 起点别名合并到“起点”
            string startCanonical = string.Empty;
            if (dst.TryGetValue("起点", out var s1) && !string.IsNullOrWhiteSpace(s1)) startCanonical = s1.Trim();
            if (string.IsNullOrWhiteSpace(startCanonical) && dst.TryGetValue("始点", out var s2) && !string.IsNullOrWhiteSpace(s2)) startCanonical = s2.Trim();
            if (string.IsNullOrWhiteSpace(startCanonical) && dst.TryGetValue("起始", out var s3) && !string.IsNullOrWhiteSpace(s3)) startCanonical = s3.Trim();
            if (!string.IsNullOrWhiteSpace(startCanonical)) dst["起点"] = startCanonical;
            dst.Remove("始点");
            dst.Remove("起始");

            // 终点仅保留“终点”
            if (dst.TryGetValue("终点", out var e1) && !string.IsNullOrWhiteSpace(e1))
                dst["终点"] = e1.Trim();

            // 管号别名合并到“管段号”
            string pipeNoCanonical = string.Empty;
            if (dst.TryGetValue("管段号", out var p1) && !string.IsNullOrWhiteSpace(p1)) pipeNoCanonical = p1.Trim();
            if (string.IsNullOrWhiteSpace(pipeNoCanonical) && dst.TryGetValue("管段编号", out var p2) && !string.IsNullOrWhiteSpace(p2)) pipeNoCanonical = p2.Trim();
            if (string.IsNullOrWhiteSpace(pipeNoCanonical) && dst.TryGetValue("管道号", out var p3) && !string.IsNullOrWhiteSpace(p3)) pipeNoCanonical = p3.Trim();
            if (!string.IsNullOrWhiteSpace(pipeNoCanonical)) dst["管段号"] = pipeNoCanonical;
            dst.Remove("管段编号");
            dst.Remove("管道号");

            // 返回标准化结果
            return dst;
        }

        /// <summary>
        /// 根据目标（规范）列名，从属性集合中查找匹配的源属性值
        /// - 优先精确键匹配（属性键与目标相同）
        /// - 其次查找归一化后等于目标的属性键，返回第一个非空值
        /// </summary>
        private string GetAttributeValueByMappedKey(Dictionary<string, string>? attrs, string mappedKey)
        {
            if (attrs == null || string.IsNullOrWhiteSpace(mappedKey)) return string.Empty;
            // 1) 精确键
            if (attrs.TryGetValue(mappedKey, out var exact) && !string.IsNullOrWhiteSpace(exact))
                return exact;
            // 2) 归一化后匹配
            foreach (var kv in attrs)
            {
                if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                var norm = NormalizeAttributeKey(kv.Key);
                if (string.Equals(norm, mappedKey, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(kv.Value))
                    return kv.Value;
            }
            // 3) 回退：包含匹配（宽松）
            foreach (var kv in attrs)
            {
                if (string.IsNullOrWhiteSpace(kv.Key) || string.IsNullOrWhiteSpace(kv.Value)) continue;
                if (kv.Key.IndexOf(mappedKey, StringComparison.OrdinalIgnoreCase) >= 0)
                    return kv.Value;
            }
            return string.Empty;
        }

        /// <summary>
        /// 获取或创建表格样式
        /// </summary>
        private ObjectId GetOrCreateTableStyle(Database db, Transaction trans)
        {
            try
            {
                // 获取表格样式字典
                DBDictionary tableStyleDict = trans.GetObject(db.TableStyleDictionaryId, OpenMode.ForRead) as DBDictionary;

                ObjectId tableStyleId = ObjectId.Null;

                // 尝试获取Standard表格样式
                if (tableStyleDict.Contains("DeviceTableStyle"))
                {
                    tableStyleId = tableStyleDict.GetAt("DeviceTableStyle");
                }
                else if (tableStyleDict.Contains("_DeviceTableStyle"))
                {
                    tableStyleId = tableStyleDict.GetAt("_DeviceTableStyle");
                }
                else
                {
                    // 创建自定义(设备)表格样式
                    tableStyleId = CreateDeviceTableStyle(db, trans, tableStyleDict);
                }

                return tableStyleId;
            }
            catch
            {
                return ObjectId.Null;
            }
        }

        /// <summary>
        /// 创建自定义(设备)表格样式
        /// </summary>
        private ObjectId CreateDeviceTableStyle(Database db, Transaction trans, DBDictionary tableStyleDict)
        {
            try
            {
                // 升级字典访问权限
                tableStyleDict.UpgradeOpen();

                // 创建新的表格样式
                TableStyle newTableStyle = new TableStyle();
                newTableStyle.Name = "DeviceTableStyle";// 表格样式名称

                // 设置标题行样式
                newTableStyle.SetAlignment(CellAlignment.MiddleCenter, (int)RowType.TitleRow);
                newTableStyle.SetTextHeight(3.5, (int)RowType.TitleRow);
                newTableStyle.SetColor(Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1), (int)RowType.TitleRow);
                newTableStyle.SetMargin(cellMargin: default, 1.5, RowType.TitleRow.ToString());

                // 设置表头行样式
                newTableStyle.SetAlignment(CellAlignment.MiddleCenter, (int)RowType.HeaderRow);
                newTableStyle.SetTextHeight(2.5, (int)RowType.HeaderRow);
                newTableStyle.SetColor(Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2), (int)RowType.HeaderRow);
                newTableStyle.SetMargin(cellMargin: default, 1.0, RowType.HeaderRow.ToString());

                // 设置数据行样式
                newTableStyle.SetAlignment(CellAlignment.MiddleCenter, (int)RowType.DataRow);
                newTableStyle.SetTextHeight(2.0, (int)RowType.DataRow);
                newTableStyle.SetColor(Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 7), (int)RowType.DataRow);
                newTableStyle.SetMargin(cellMargin: default, 1.0, RowType.DataRow.ToString());

                // 设置网格线
                newTableStyle.SetGridLineWeight(LineWeight.LineWeight025, (int)GridLineType.AllGridLines, (int)RowType.TitleRow);
                newTableStyle.SetGridLineWeight(LineWeight.LineWeight025, (int)GridLineType.AllGridLines, (int)RowType.HeaderRow);
                newTableStyle.SetGridLineWeight(LineWeight.LineWeight025, (int)GridLineType.AllGridLines, (int)RowType.DataRow);

                // 添加到字典
                tableStyleDict.SetAt("DeviceTableStyle", newTableStyle);
                trans.AddNewlyCreatedDBObject(newTableStyle, true);

                return newTableStyle.ObjectId;
            }
            catch
            {
                return ObjectId.Null;
            }
        }

        /// <summary>
        /// 设置表格样式 - 修正版
        /// </summary>
        private void SetTableStyle(Database db, Table table, Transaction trans, double scaleDenominator)
        {
            try
            {
                if (table == null) return;

                ObjectId tableStyleId = GetOrCreateTableStyle(db, trans);
                if (!tableStyleId.IsNull)
                {
                    table.TableStyle = tableStyleId;
                }

                // 优先使用传入的 scaleDenominator（>0），否则从数据库/视口读取
                double scaleDenom = scaleDenominator > 0.0 ? scaleDenominator : 0.0;
                if (scaleDenom <= 0.0)
                {
                    try
                    {
                        scaleDenom = GetScaleDenominatorForDatabase(db, roundToCommon: false);
                    }
                    catch
                    {
                        scaleDenom = 1.0;
                    }
                }

                // 应用按比例计算的文字高度与行高（会在单元格 TextHeight 上设置）
                ApplyScaledHeightsToTable(table, scaleDenom);

                // 基于单元格文字高度和内容自适应列宽
                AutoResizeColumns(table);

                // 设置网格线粗细
                try
                {
                    SetTableBorders(table, scaleDenom);
                }
                catch { /* 忽略 */ }

                // 设置单元格对齐与兜底 TextHeight
                SetCellStyles(table);
            }
            catch (Exception ex)
            {
                SetBasicTableStyle(table);
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage($"\nSetTableStyle 异常，已回退基本样式: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置单元格样式（只设置对齐与必要的默认字体样式，避免再次硬编码高度）
        /// </summary>
        private void SetCellStyles(Table table)
        {
            try
            {
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        var cell = table.Cells[r, c];
                        // 居中对齐为默认
                        cell.Alignment = CellAlignment.MiddleCenter;

                        // 确保如果某些单元格未设置 TextHeight，则使用相应行的默认值（安全兜底）
                        if (cell.TextHeight <= 0.0)
                        {
                            if (r == 0) cell.TextHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, 1.0);
                            else if (r == 1 || r == 2) cell.TextHeight = TextFontsStyleHelper.ComputeScaledHeight(2.5, 1.0);
                            else cell.TextHeight = TextFontsStyleHelper.ComputeScaledHeight(2.0, 1.0);
                        }
                    }
                }
            }
            catch
            {
                // 忽略单元格样式设置错误（保持容错）
            }
        }

        /// <summary>
        /// 自适应调整列宽（使用单元格 TextHeight 作为尺度基准，保证按相同比例放缩）
        /// 说明：避免直接依赖外部的比例值，改为基于已设置的单元格 TextHeight 计算列宽。
        /// </summary>
        //private void AutoResizeColumns(Table table)
        //{
        //    try
        //    {
        //        for (int j = 0; j < table.Columns.Count; j++)
        //        {
        //            double maxWidth = 10.0; // 最小宽度（绘图单位）
        //            double sampleTextHeight = 0.0;

        //            // 1) 估算该列使用的平均文字高度（取第一非零的 TextHeight）
        //            for (int i = 0; i < table.Rows.Count; i++)
        //            {
        //                try
        //                {
        //                    var th = table.Cells[i, j].TextHeight;
        //                    if (th > 0.0)
        //                    {
        //                        sampleTextHeight = Convert.ToDouble(th);
        //                        break;
        //                    }
        //                }
        //                catch { }
        //            }
        //            // 若未找到有效 TextHeight，尝试行内查找（兜底）
        //            if (sampleTextHeight <= 0.0)
        //            {
        //                // 取全表第一个非零 TextHeight
        //                for (int ii = 0; ii < table.Rows.Count && sampleTextHeight <= 0.0; ii++)
        //                {
        //                    for (int jj = 0; jj < table.Columns.Count && sampleTextHeight <= 0.0; jj++)
        //                    {
        //                        try { sampleTextHeight = Convert.ToDouble(table.Cells[ii, jj].TextHeight); }
        //                        catch { sampleTextHeight = 0.0; }
        //                    }
        //                }
        //            }

        //            // 保守兜底：若仍然无有效 TextHeight，则使用一个合理默认
        //            if (sampleTextHeight <= 0.0)
        //                sampleTextHeight = TextFontsStyleHelper.ComputeScaledHeight(2.0, 1.0);

        //            // 2) 计算每行文本的估算字符宽度（以字符单位计），并据此求出最大需求
        //            for (int i = 0; i < table.Rows.Count; i++)
        //            {
        //                string text = string.Empty;
        //                try { text = table.Cells[i, j].TextString ?? string.Empty; } catch { text = string.Empty; }
        //                if (!string.IsNullOrEmpty(text))
        //                {
        //                    double charUnits = EstimateTextWidth(text); // 近似单位
        //                    // 根据文字高度估算所需宽度（经验系数）
        //                    double charWidthFactor = 0.55; // 每字符宽度相对 TextHeight 的经验系数，可微调
        //                    double estWidth = Math.Ceiling(charUnits * sampleTextHeight * charWidthFactor);
        //                    if (estWidth > maxWidth) maxWidth = estWidth;
        //                }
        //            }

        //            // 3) 限制列宽范围，避免过窄或过宽
        //            double minColWidth = Math.Max(8.0, sampleTextHeight * 2.5);
        //            double maxColWidth = Math.Max(60.0, sampleTextHeight * 40.0);

        //            double finalWidth = Math.Max(minColWidth, Math.Min(maxColWidth, maxWidth));

        //            try { table.SetColumnWidth(j, finalWidth); } catch { /* 忽略设置失败 */ }
        //        }
        //    }
        //    catch
        //    {
        //        // 兜底：保持现状，不抛出异常以免中断调用方流程
        //    }
        //}

        /// <summary>
        /// 设置表格边框并按比例选择合适的线宽（尝试性，API 不同版本行为不同，用 try/catch 包裹）
        /// </summary>
        private void SetTableBorders(Table table, double scaleDenominator)
        {
            if (table == null) return;

            try
            {
                // 根据比例分母选择更合适的线宽（经验映射）
                LineWeight lw = LineWeight.LineWeight025;
                if (scaleDenominator >= 1000.0) lw = LineWeight.LineWeight070;
                else if (scaleDenominator >= 500.0) lw = LineWeight.LineWeight050;
                else if (scaleDenominator >= 200.0) lw = LineWeight.LineWeight050;
                else lw = LineWeight.LineWeight025;

                bool applied = false;

                // 优先尝试基于单元格的新 API（反射，兼容不同版本）
                try
                {
                    for (int r = 0; r < table.Rows.Count && !applied; r++)
                    {
                        for (int c = 0; c < table.Columns.Count; c++)
                        {
                            try
                            {
                                var cell = table.Cells[r, c];
                                if (cell == null) continue;

                                var cellType = cell.GetType();

                                // 1) 尝试调用可能存在的 SetGridLineWeight 方法（Cell 级）
                                var mi = cellType.GetMethod("SetGridLineWeight", new Type[] { typeof(LineWeight), typeof(int), typeof(int) });
                                if (mi != null)
                                {
                                    mi.Invoke(cell, new object[] { lw, (int)GridLineType.AllGridLines, 0 });
                                    applied = true;
                                    break;
                                }

                                // 2) 尝试设置名为 GridLineWeight 的可写属性（某些 API 以属性形式暴露）
                                var prop = cellType.GetProperty("GridLineWeight");
                                if (prop != null && prop.CanWrite && prop.PropertyType == typeof(LineWeight))
                                {
                                    prop.SetValue(cell, lw);
                                    applied = true;
                                    break;
                                }

                                // 3) 尝试存在的其它重载（最通用的最后尝试）
                                var candidates = cellType.GetMethods().Where(m => m.Name == "SetGridLineWeight").ToArray();
                                foreach (var cand in candidates)
                                {
                                    var ps = cand.GetParameters();
                                    if (ps.Length == 1 && ps[0].ParameterType == typeof(LineWeight))
                                    {
                                        cand.Invoke(cell, new object[] { lw });
                                        applied = true;
                                        break;
                                    }
                                    else if (ps.Length == 2 && ps[0].ParameterType == typeof(LineWeight) && ps[1].ParameterType == typeof(int))
                                    {
                                        cand.Invoke(cell, new object[] { lw, (int)GridLineType.AllGridLines });
                                        applied = true;
                                        break;
                                    }
                                }
                                if (applied) break;
                            }
                            catch
                            {
                                // 单个单元格失败继续尝试其它单元格
                                continue;
                            }
                        }
                    }
                }
                catch
                {
                    // 忽略单元格级别尝试的任何异常，回退到表级 API
                    applied = false;
                }

                // 如果单元格级应用未生效，回退到表级方法（某些旧版本仍然有效）
                if (!applied)
                {
                    try
                    {
                        // 尝试直接调用 Table.SetGridLineWeight（可能已被标记为 Obsolete，但在部分版本仍可用）
#pragma warning disable 0618
                        table.SetGridLineWeight(lw, (int)GridLineType.AllGridLines, 0);
                        table.SetGridLineWeight(lw, (int)GridLineType.AllGridLines, 1);
                        table.SetGridLineWeight(lw, (int)GridLineType.AllGridLines, 2);
#pragma warning restore 0618
                        applied = true;
                    }
                    catch
                    {
                        // 再尝试通过反射设置可能存在的 GridLineWeight 属性（表级）
                        try
                        {
                            var tType = table.GetType();
                            var prop = tType.GetProperty("GridLineWeight");
                            if (prop != null && prop.CanWrite && prop.PropertyType == typeof(LineWeight))
                            {
                                prop.SetValue(table, lw);
                                applied = true;
                            }
                        }
                        catch
                        {
                            // 最终兜底：不再尝试逐单元格设置，保持程序可用性
                        }
                    }
                }
            }
            catch
            {
                // 忽略任何边框设置失败，保证功能不中断
            }
        }

        /// <summary>
        /// 设置基本表格样式（备选方案）
        /// </summary>
        private void SetBasicTableStyle(Table table)
        {
            try
            {
                // 基本的行高设置
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    table.SetRowHeight(i, i == 0 ? 12.0 : 8.0);
                }

                // 基本的列宽设置
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    table.SetColumnWidth(j, 25.0);
                }

                // 设置基本边框
                //table.SetGridLineWeight(LineWeight.LineWeight025, (int)GridLineType.AllGridLines);
            }
            catch
            {
                // 忽略错误
            }
        }

        /// <summary>
        /// 填充固定列头
        /// pipeGroupCount: 管道组（"管道 Pipe (m)"）的子列数（可以大于初始 8）
        /// 基本固定列索引说明（基列数 baseFixedCols = 10）：
        /// 0: 管道标题
        /// 1: 管段号
        /// 2-3: 起点/终点（合并为组）
        /// 4: 管道等级
        /// 5-7: 设计条件（介质/温度/压力）
        /// 8-9: 隔热及防腐（Code/Antisepsis）
        /// 随后 pipeGroupCount 个列为管道组子列（名称/材料/...等）
        /// </summary>
        private void Fill_PipeLine_FixedHeaders(Table table, int pipeGroupCount)
        {
            try
            {
                // 基础固定列（0..9）
                // 管道标题：第0列（索引0），跨2行
                table.MergeCells(CellRange.Create(table, 1, 0, 2, 0));
                table.Cells[1, 0].TextString = "管道标题\nPipe Title";

                // 管段号：第1列（索引1），跨2行 管道号
                table.MergeCells(CellRange.Create(table, 1, 1, 2, 1));
                table.Cells[1, 1].TextString = "管段号\nPipeline\nNo.";

                // 管段起止点：第2-3列（索引2,3），row1 合并为组，row2 分别为 起点/终点
                table.MergeCells(CellRange.Create(table, 1, 2, 1, 3));
                table.Cells[1, 2].TextString = "管段起止点\nPipeline From Start To End";
                table.Cells[2, 2].TextString = "起点\nFrom";
                table.Cells[2, 3].TextString = "终点\nTo";

                // 管道等级：第4列（索引4），跨2行
                table.MergeCells(CellRange.Create(table, 1, 4, 2, 4));
                table.Cells[1, 4].TextString = "管道\n等级\nPipe Class";

                // 设计条件：第5-7列（索引5..7），row1 合并为组，row2: 介质名称/操作温度/操作压力
                table.MergeCells(CellRange.Create(table, 1, 5, 1, 7));
                table.Cells[1, 5].TextString = "设计条件 \nDesign Condition";
                table.Cells[2, 5].TextString = "介质名称\nMedium Name";
                table.Cells[2, 6].TextString = "操作温度\nT(℃)";
                table.Cells[2, 7].TextString = "操作压力\nP(MPaG)";

                // 隔热及防腐：第8-9列（索引8..9），row1 合并为组，row2: Code / Antisepsis
                table.MergeCells(CellRange.Create(table, 1, 8, 1, 9));
                table.Cells[1, 8].TextString = "隔热及防腐 \nInsul.& Antisepsis";
                table.Cells[2, 8].TextString = "隔热隔声代号\nCode";
                table.Cells[2, 9].TextString = "是否防腐\nAntisepsis";

                // 管道组：从第10列开始，动态宽度由 pipeGroupCount 决定
                int pipeGroupStart = 10;
                int pipeGroupEnd = pipeGroupStart + Math.Max(0, pipeGroupCount - 1);
                if (pipeGroupEnd >= table.Columns.Count) pipeGroupEnd = table.Columns.Count - 1;
                if (pipeGroupStart < table.Columns.Count)
                {
                    table.MergeCells(CellRange.Create(table, 1, pipeGroupStart, 1, pipeGroupEnd));
                    table.Cells[1, pipeGroupStart].TextString = "管道\nPipe (m)";
                    table.Cells[1, pipeGroupStart].Alignment = CellAlignment.MiddleCenter;

                    // 默认子列标题（前8项为常用）；额外列留空，由 CreateDeviceTableWithType 填写具体名字
                    string[] defaultPipeSubHeaders = new[]
                    {
                         "名称\nName",
                        "材料\nMaterial",
                        "图号或标准号\nDWG.No./ STD.No.",
                        "数量\nQuan.",
                        "泵前/后\nPump F/B",
                        "核算流速\n(M/S)",
                        "管道长度(mm)\nLength(mm)",
                        "累计长度(mm)\nAllLength(mm)"

                    };

                    for (int i = 0; i <= pipeGroupEnd - pipeGroupStart; i++)
                    {
                        int col = pipeGroupStart + i;
                        if (i < defaultPipeSubHeaders.Length)
                            table.Cells[2, col].TextString = defaultPipeSubHeaders[i];
                        else
                            table.Cells[2, col].TextString = string.Empty; // 额外列由上层动态填写列名
                    }
                }
            }
            catch
            {
                // 容错：忽略任何设置异常
            }
        }

        /// <summary>
        /// 估算文本宽度
        /// </summary>
        private double EstimateTextWidth(string text)
        {
            double width = 0;
            foreach (char c in text)
            {
                if (c > 127) // 中文字符
                    width += 1.5;
                else
                    width += 0.8;
            }
            return width + 2.0; // 增加一些边距
            //foreach (char c in text)
            //{
            //    if (c > 127) // 中文字符
            //        width += 3.0;
            //    else
            //        width += 1.5;
            //}
            //return width + 4.0; // 增加一些边距
        }


        /// <summary>
        /// 根据视口尺度因子和比例字符串确定最终的比例分母
        /// </summary>
        /// <param name="table"></param>
        /// <param name="scaleDenominator"></param>
        private void ApplyScaledHeightsToTable(Autodesk.AutoCAD.DatabaseServices.Table table, double scaleDenominator)
        {
            // 将约定的基准字高按比例应用到表格的Title/Header/Data单元格（仅设置单元格 TextHeight，避免修改全局 TextStyle）
            if (table == null) return;

            double titleHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, scaleDenominator);
            double headerHeight = TextFontsStyleHelper.ComputeScaledHeight(2.5, scaleDenominator);
            double dataHeight = TextFontsStyleHelper.ComputeScaledHeight(2.0, scaleDenominator);

            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    try
                    {
                        if (r == 0)
                            table.Cells[r, c].TextHeight = titleHeight;
                        else if (r == 1 || r == 2)
                            table.Cells[r, c].TextHeight = headerHeight;
                        else
                            table.Cells[r, c].TextHeight = dataHeight;
                    }
                    catch
                    {
                        // 某些单元在某些 AutoCAD 版本可能不可写，忽略
                    }
                }

                // 可选：把行高也以字高为参考设置（保持视觉一致）
                try
                {
                    if (r == 0)
                        table.SetRowHeight(r, Math.Max(8.0, titleHeight * 3.0));
                    else if (r == 1 || r == 2)
                        table.SetRowHeight(r, Math.Max(6.0, headerHeight * 2.5));
                    else
                        table.SetRowHeight(r, Math.Max(5.0, dataHeight * 2.2));
                }
                catch { }
            }
        }
        /// <summary>
        /// 获取数据库的缩放因子
        /// </summary>
        /// <param name="db"></param>
        /// <param name="roundToCommon"></param>
        /// <returns></returns>
        //private double GetScaleDenominatorForDatabase(Database db, bool roundToCommon = false)
        //{
        //    try
        //    {
        //        // 优先读取视口 scale factor（0.01 表示 1:100）
        //        double vpScaleFactor = GetActiveViewportScaleFactor(db);
        //        // Delegate to FontsStyleHelper.DetermineScaleDenominator 做归一化与容错
        //        return FontsStyleHelper.DetermineScaleDenominator(viewportScaleFactor: vpScaleFactor, scaleString: null, roundToCommon: roundToCommon);
        //    }
        //    catch
        //    {
        //        return 1.0;
        //    }
        //}

        /// <summary>
        /// 导出到Excel命令
        /// </summary>
        [CommandMethod(nameof(ExportTableToExcel))]
        public void ExportTableToExcel()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // 选择表格
                PromptEntityOptions opts = new PromptEntityOptions("\n请选择要导出的表格：");
                opts.SetRejectMessage("\n请选择一个表格对象。");
                opts.AddAllowedClass(typeof(Table), true);

                PromptEntityResult selResult = ed.GetEntity(opts);
                if (selResult.Status != PromptStatus.OK)
                    return;

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    Table table = trans.GetObject(selResult.ObjectId, OpenMode.ForRead) as Table;
                    if (table == null)
                    {
                        ed.WriteMessage("\n选择的对象不是表格。");
                        return;
                    }

                    // 导出到Excel
                    ExportTableToExcelFile(table, ed);
                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\n导出Excel时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 导出表格到Excel文件（保留合并单元格与文本、简单样式）
        /// 改进：合并检测改为严格检查候选矩形内所有单元格均未被标记为已合并且为空，以避免产生部分重叠的合并区域（解决 "Can't delete/overwrite merged cells" 错误）。
        /// </summary>
        private void ExportTableToExcelFile(Table table, Editor ed)
        {
            // 获取保存路径
            System.Windows.Forms.SaveFileDialog saveDialog = new System.Windows.Forms.SaveFileDialog();
            saveDialog.Filter = "Excel文件|*.xlsx";
            saveDialog.Title = "保存设备材料表";
            saveDialog.FileName = $"设备材料表_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            if (saveDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;

            try
            {
                // 创建Excel工作簿 使用EPPlus创建Excel文件
                using (ExcelPackage package = new ExcelPackage())
                {
                    // 创建工作表
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("设备材料表");
                    // 获取行数和列数 设置默认列宽和行高
                    int rows = table.Rows.Count;
                    int cols = table.Columns.Count;

                    // 先把所有文本写入单元格（仅文本，不处理合并）
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            // 处理文本 读取单元格文本
                            string cellText = table.Cells[r, c].TextString ?? string.Empty;
                            worksheet.Cells[r + 1, c + 1].Value = string.IsNullOrEmpty(cellText) ? "" : cellText;
                            worksheet.Cells[r + 1, c + 1].Style.WrapText = true;
                        }
                    }

                    // 标记已被合并/处理的单元格，初始为 false
                    var mergedMark = new bool[rows, cols];

                    // 扫描每个单元格，找到非空且未处理的起始单元格后尝试扩展为矩形合并区域
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            // 跳过已处理的单元格
                            if (mergedMark[r, c])
                                continue;

                            // 读取当前单元格文本并跳过空单元格
                            string txt = (table.Cells[r, c].TextString ?? string.Empty).Trim();
                            if (string.IsNullOrEmpty(txt))
                            {
                                mergedMark[r, c] = true;
                                continue;
                            }

                            // 计算最大水平扩展：要求右侧单元为空且未被标记
                            int maxH = 1;
                            while (c + maxH < cols)
                            {
                                if (mergedMark[r, c + maxH]) break;
                                var rightTxt = (table.Cells[r, c + maxH].TextString ?? string.Empty).Trim();
                                if (!string.IsNullOrEmpty(rightTxt)) break;
                                maxH++;
                            }

                            // 计算最大垂直扩展：对于每一行，要求从 c..c+maxH-1 都为空且未被标记
                            int maxV = 1;
                            while (r + maxV < rows)
                            {
                                bool rowOk = true;
                                for (int cc = c; cc < c + maxH; cc++)
                                {
                                    if (mergedMark[r + maxV, cc])
                                    {
                                        rowOk = false;
                                        break;
                                    }
                                    var downTxt = (table.Cells[r + maxV, cc].TextString ?? string.Empty).Trim();
                                    if (!string.IsNullOrEmpty(downTxt))
                                    {
                                        rowOk = false;
                                        break;
                                    }
                                }
                                if (!rowOk) break;
                                maxV++;
                            }

                            // 规则变更：
                            // - 保持第1-3行（0-based 0..2）原有合并逻辑
                            // - 第4行及以后（r >= 3）禁止任何方向的合并（既禁止横向也禁止纵向）
                            // - 若起始行在第1-3行，但合并会跨过第3行边界，则限制垂直合并使其不会跨入第4行（即 r+maxV-1 <= 2）
                            if (r >= 3)
                            {
                                // 第4行以后的起始单元：禁止横向和纵向合并
                                maxH = 1;
                                maxV = 1;
                            }
                            else
                            {
                                // 起始行在 0..2：允许横向合并，但垂直合并不能跨入第4行（index >=3）
                                int maxAllowedV = 3 - r; // 例如：r=0 -> maxAllowedV=3 (rows 0,1,2)，r=1 ->2, r=2 ->1
                                if (maxV > maxAllowedV) maxV = maxAllowedV;
                            }

                            // 进一步确保矩形内部所有单元均未被标记（防止与先前合并产生部分重叠）
                            bool rectangleClear = true;
                            for (int rr = r; rr < r + maxV && rectangleClear; rr++)
                            {
                                for (int cc = c; cc < c + maxH; cc++)
                                {
                                    if (mergedMark[rr, cc])
                                    {
                                        rectangleClear = false;
                                        break;
                                    }
                                }
                            }

                            if (!rectangleClear)
                            {
                                // 如果候选矩形内部有已标记单元，则退回为单元格不合并（标记当前单元）
                                mergedMark[r, c] = true;
                                continue;
                            }

                            // 只有当矩形尺寸大于1才合并，否则单个单元标记为已处理
                            if (maxH > 1 || maxV > 1)
                            {
                                int excelRow1 = r + 1; // Excel 行号 1-based
                                int excelCol1 = c + 1;
                                int excelRow2 = r + maxV;
                                int excelCol2 = c + maxH;

                                try
                                {
                                    worksheet.Cells[excelRow1, excelCol1, excelRow2, excelCol2].Merge = true;
                                    worksheet.Cells[excelRow1, excelCol1, excelRow2, excelCol2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                    worksheet.Cells[excelRow1, excelCol1, excelRow2, excelCol2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                }
                                catch
                                {
                                    // 合并冲突时回退为不合并，保持单元格内容
                                    worksheet.Cells[excelRow1, excelCol1].Value = worksheet.Cells[excelRow1, excelCol1].Value;
                                }

                                // 标记该矩形已被处理
                                for (int rr = r; rr < r + maxV; rr++)
                                    for (int cc = c; cc < c + maxH; cc++)
                                        mergedMark[rr, cc] = true;
                            }
                            else
                            {
                                mergedMark[r, c] = true;
                            }
                        }
                    }

                    // 特殊处理：若第1行为标题并在 AutoCAD 中被合并（大多数场景是如此），确保 Excel 中也是合并并加粗居中
                    try
                    {
                        string firstCell = (table.Cells[0, 0].TextString ?? string.Empty).Trim();
                        bool otherEmpty = true;
                        for (int cc = 1; cc < cols; cc++)
                        {
                            if (!string.IsNullOrWhiteSpace(table.Cells[0, cc].TextString ?? string.Empty))
                            {
                                otherEmpty = false;
                                break;
                            }
                        }
                        if (!string.IsNullOrEmpty(firstCell) && otherEmpty)
                        {
                            worksheet.Cells[1, 1, 1, cols].Merge = true;
                            worksheet.Cells[1, 1].Style.Font.Bold = true;
                            worksheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheet.Cells[1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        }
                    }
                    catch { /* 忽略 */ }

                    // 自动调整列宽
                    worksheet.Cells.AutoFitColumns();

                    // 基本样式设置
                    worksheet.Cells.Style.Font.Name = "宋体";
                    if (rows > 0 && cols > 0)
                    {
                        worksheet.Cells[1, 1, rows, cols].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[1, 1, rows, cols].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[1, 1, rows, cols].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[1, 1, rows, cols].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    // 保存文件
                    package.SaveAs(new FileInfo(saveDialog.FileName));
                }

                ed.WriteMessage($"\n设备材料表已成功导出到: {saveDialog.FileName}");
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\n导出Excel文件时发生错误: {ex.Message}");
            }
        }

        #region 同步表格

        /// <summary>
        /// 【临时】用于 WPF 窗口传递待同步的数据列表
        /// </summary>
        public static List<DeviceInfo> TempSyncData { get; set; } = new List<DeviceInfo>();

        /// <summary>
        /// 同步表格
        /// </summary>
        [CommandMethod(nameof(SyncTableToEntities))]
        public void SyncTableToEntities()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // 选择表格
                var peo = new PromptEntityOptions("\n请选择要同步的表格（Table）：");
                peo.SetRejectMessage("\n请选择一个表格对象。");
                peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Table), true);
                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var table = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Table;
                    if (table == null)
                    {
                        ed.WriteMessage("\n选中的对象不是表格。");
                        return;
                    }

                    // 自动识别数据起始行（常见：0=title,1=headerRow1,2=headerRow2, 数据自3起）
                    int dataStartRow = 3;
                    if (table.Rows.Count <= dataStartRow)
                    {
                        ed.WriteMessage("\n表格行数过少，无法识别数据行，请确认表格结构。");
                        return;
                    }

                    // 查找标识列（用于匹配图元）
                    int idCol = -1;
                    string[] idKeywords = new[] { "部件ID", "部件 Id", "部件编号", "管段号", "管段编号", "Pipeline", "Pipe No", "ID", "序号" };
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        string h1 = (table.Cells[1, c].TextString ?? string.Empty).Trim();
                        string h2 = (table.Cells[2, c].TextString ?? string.Empty).Trim();
                        string header = string.IsNullOrWhiteSpace(h2) ? h1 : h2;
                        foreach (var kw in idKeywords)
                        {
                            if (!string.IsNullOrWhiteSpace(header) && header.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                idCol = c;
                                break;
                            }
                        }
                        if (idCol != -1) break;
                    }

                    if (idCol == -1)
                    {
                        ed.WriteMessage("\n未能在表头中找到标识列（例如“部件ID”或“管段号”）。请在表格中包含用于匹配块的标识列后重试。");
                        return;
                    }

                    // 准备要扫描的空间（Model & Paper）
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var spaceIds = new List<ObjectId>();
                    try
                    {
                        if (bt.Has(BlockTableRecord.ModelSpace))
                            spaceIds.Add(bt[BlockTableRecord.ModelSpace]);
                        if (bt.Has(BlockTableRecord.PaperSpace))
                            spaceIds.Add(bt[BlockTableRecord.PaperSpace]);
                    }
                    catch { }

                    // 预扫描建立索引（按可能的标识属性值）
                    var blockIndex = new Dictionary<string, List<ObjectId>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var spaceId in spaceIds)
                    {
                        try
                        {
                            var space = tr.GetObject(spaceId, OpenMode.ForRead) as BlockTableRecord;
                            if (space == null) continue;
                            foreach (ObjectId entId in space)
                            {
                                try
                                {
                                    var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                                    if (ent is BlockReference br)
                                    {
                                        foreach (ObjectId aid in br.AttributeCollection)
                                        {
                                            try
                                            {
                                                var ar = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                                                if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue;
                                                string tag = ar.Tag.Trim();
                                                // 如果标签看起来像 ID，则加入索引
                                                if (idKeywords.Any(k => tag.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                                                {
                                                    string val = (ar.TextString ?? string.Empty).Trim();
                                                    if (!string.IsNullOrWhiteSpace(val))
                                                    {
                                                        if (!blockIndex.ContainsKey(val)) blockIndex[val] = new List<ObjectId>();
                                                        blockIndex[val].Add(br.ObjectId);
                                                    }
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }

                    // 不可回写字段（归一化后匹配）
                    var nonWritable = new[]
                    {
                "起点","终点","始点","止点","起止点","起止","位置","坐标","方向","角度",
                "管段号","管段编号","管道标题","长度","长度(m)","累计长度","累计长度(mm)"
            }.Select(k => NormalizeAttributeKey(k)).Where(x => !string.IsNullOrWhiteSpace(x)).ToHashSet(StringComparer.OrdinalIgnoreCase);

                    var unmatchedDiagnostics = new List<string>();
                    int updatedCount = 0;

                    // 遍历每条数据行
                    for (int r = dataStartRow; r < table.Rows.Count; r++)
                    {
                        string idCell = (table.Cells[r, idCol].TextString ?? string.Empty).Trim();
                        if (string.IsNullOrWhiteSpace(idCell)) continue;

                        List<ObjectId> matchedBlocks;
                        if (!blockIndex.TryGetValue(idCell, out var idxList))
                        {
                            // 回退搜索：在空间中查找任一块其任意属性值等于 idCell
                            matchedBlocks = new List<ObjectId>();
                            foreach (var spaceId in spaceIds)
                            {
                                try
                                {
                                    var space = tr.GetObject(spaceId, OpenMode.ForRead) as BlockTableRecord;
                                    if (space == null) continue;
                                    foreach (ObjectId entId in space)
                                    {
                                        try
                                        {
                                            var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                                            if (ent is BlockReference br)
                                            {
                                                bool found = false;
                                                foreach (ObjectId aid in br.AttributeCollection)
                                                {
                                                    try
                                                    {
                                                        var ar = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                                                        if (ar == null) continue;
                                                        if (string.Equals((ar.TextString ?? string.Empty).Trim(), idCell, StringComparison.OrdinalIgnoreCase))
                                                        {
                                                            found = true;
                                                            break;
                                                        }
                                                    }
                                                    catch { }
                                                }
                                                if (found) matchedBlocks.Add(br.ObjectId);
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                catch { }
                            }
                        }
                        else
                        {
                            matchedBlocks = idxList.ToList();
                        }

                        if (matchedBlocks == null || matchedBlocks.Count == 0) continue;

                        // 遍历每列，写回（跳过标识列与不可回写字段）
                        for (int ci = 0; ci < table.Columns.Count; ci++)
                        {
                            if (ci == idCol) continue;

                            string head2 = (table.Cells[2, ci].TextString ?? string.Empty).Trim();
                            string head1 = (table.Cells[1, ci].TextString ?? string.Empty).Trim();
                            string header = string.IsNullOrWhiteSpace(head2) ? head1 : head2;
                            if (string.IsNullOrWhiteSpace(header)) continue;

                            var headerLine = header.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                            var normalizedHeader = NormalizeAttributeKey(headerLine);

                            // 跳过不可回写字段
                            if (!string.IsNullOrWhiteSpace(normalizedHeader) && nonWritable.Contains(normalizedHeader))
                                continue;

                            string newValue = (table.Cells[r, ci].TextString ?? string.Empty).Trim();
                            string cleanedValue = CleanAttributeText(newValue);

                            bool anyMatchedForThisCell = false;

                            // 写回每个匹配到的块
                            foreach (var bid in matchedBlocks)
                            {
                                try
                                {
                                    var br = tr.GetObject(bid, OpenMode.ForWrite) as BlockReference;
                                    if (br == null) continue;

                                    // 1) 尝试写回动态块属性（若存在）
                                    try
                                    {
                                        if (br.IsDynamicBlock)
                                        {
                                            var dynProps = br.DynamicBlockReferencePropertyCollection;
                                            foreach (DynamicBlockReferenceProperty prop in dynProps)
                                            {
                                                try
                                                {
                                                    var propNameNorm = NormalizeAttributeKey(prop.PropertyName ?? string.Empty);
                                                    if (string.Equals(propNameNorm, normalizedHeader, StringComparison.OrdinalIgnoreCase) ||
                                                        prop.PropertyName.IndexOf(headerLine, StringComparison.OrdinalIgnoreCase) >= 0)
                                                    {
                                                        // 尝试转换类型再赋值，若失败则直接赋字符串
                                                        try
                                                        {
                                                            var targetType = prop.Value?.GetType() ?? typeof(string);
                                                            object conv;
                                                            if (targetType == typeof(string))
                                                                conv = cleanedValue;
                                                            else
                                                            {
                                                                conv = Convert.ChangeType(cleanedValue, targetType, System.Globalization.CultureInfo.InvariantCulture);
                                                            }
                                                            prop.Value = conv;
                                                        }
                                                        catch
                                                        {
                                                            try { prop.Value = cleanedValue; } catch { }
                                                        }
                                                        anyMatchedForThisCell = true;
                                                    }
                                                }
                                                catch { }
                                            }
                                        }
                                    }
                                    catch { /* 忽略动态块写回失败 */ }

                                    // 2) 尝试写回普通 AttributeReference（优先精确标签/规范化匹配，再做部分匹配）
                                    try
                                    {
                                        foreach (ObjectId aid in br.AttributeCollection)
                                        {
                                            try
                                            {
                                                var ar = tr.GetObject(aid, OpenMode.ForWrite) as AttributeReference;
                                                if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue;
                                                var tagNorm = NormalizeAttributeKey(ar.Tag ?? string.Empty);

                                                if (string.Equals(tagNorm, normalizedHeader, StringComparison.OrdinalIgnoreCase) ||
                                                    string.Equals(ar.Tag, headerLine, StringComparison.OrdinalIgnoreCase) ||
                                                    ar.Tag.IndexOf(headerLine, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                    normalizedHeader.IndexOf(tagNorm ?? string.Empty, StringComparison.OrdinalIgnoreCase) >= 0)
                                                {
                                                    ar.TextString = cleanedValue ?? string.Empty;
                                                    try { ar.AdjustAlignment(db); } catch { }
                                                    anyMatchedForThisCell = true;
                                                }
                                            }
                                            catch { /* 单个属性写入失败忽略 */ }
                                        }
                                    }
                                    catch { /* 忽略 */ }

                                    // 3) 若没有任何匹配，尝试用部分包含匹配再写一次（宽松匹配）
                                    if (!anyMatchedForThisCell)
                                    {
                                        try
                                        {
                                            foreach (ObjectId aid in br.AttributeCollection)
                                            {
                                                try
                                                {
                                                    var ar = tr.GetObject(aid, OpenMode.ForWrite) as AttributeReference;
                                                    if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue;
                                                    if (!string.IsNullOrWhiteSpace(headerLine) &&
                                                        (ar.Tag.IndexOf(headerLine, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                         headerLine.IndexOf(ar.Tag, StringComparison.OrdinalIgnoreCase) >= 0))
                                                    {
                                                        ar.TextString = cleanedValue ?? string.Empty;
                                                        try { ar.AdjustAlignment(db); } catch { }
                                                        anyMatchedForThisCell = true;
                                                        break;
                                                    }
                                                }
                                                catch { }
                                            }
                                        }
                                        catch { }
                                    }

                                    if (anyMatchedForThisCell) updatedCount++;
                                }
                                catch { /* 单个块写回问题忽略 */ }
                            } // end foreach block

                            if (!anyMatchedForThisCell)
                            {
                                unmatchedDiagnostics.Add($"行{r + 1} 列{ci + 1} 头:'{headerLine}' 值:'{newValue}' ID:'{idCell}'");
                            }
                        } // end for columns
                    } // end for rows

                    tr.Commit();

                    ed.WriteMessage($"\n同步完成，尝试更新属性项数（估计）: {updatedCount}。");
                    if (unmatchedDiagnostics.Count > 0)
                    {
                        ed.WriteMessage($"\n未匹配的单元（部分样例，最多显示20条）：");
                        foreach (var s in unmatchedDiagnostics.Take(20))
                            ed.WriteMessage("\n  " + s);
                        ed.WriteMessage("\n提示：若某些表头未被匹配，请检查表头与块属性标签的命名或在 DictionaryHelper.AttributeSynonyms/ChineseToEnglish 中补充同义词映射。");
                    }
                } // using tr
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage($"\n同步失败: {ex.Message}");
            }
        }

        #region 新增：从 Excel/表格数据同步属性到 CAD 图元

        /// <summary>
        /// 【临时静态变量】用于 WPF 窗口传递待同步的 Excel 数据列表
        /// </summary>
        public static List<DeviceInfo> TempExcelSyncData { get; set; } = new List<DeviceInfo>();

        /// <summary>
        /// 【新命令】从 WPF 传递的 Excel 表格数据同步属性到选中的 CAD 图元
        /// 命令名: SyncExcelDataToCad
        /// 匹配规则：优先匹配块属性中的 "管段号"，其次匹配 "Name"
        /// </summary>
        [CommandMethod("SyncExcelDataToCad")]
        public void SyncExcelDataToCadEntities()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            // 1. 检查是否有待同步数据
            if (TempExcelSyncData == null || TempExcelSyncData.Count == 0)
            {
                ed.WriteMessage("\n错误：没有待同步的表格数据。请先在 WPF 窗口中加载 Excel 并点击同步按钮。");
                return;
            }

            // 2. 提示用户选择要更新的图元
            ed.WriteMessage("\n请选择要同步属性的管道/设备图元（支持多选，回车结束）：");
            var psr = ed.GetSelection();
            if (psr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n已取消选择。");
                TempExcelSyncData.Clear(); // 清空数据
                return;
            }

            var selectedIds = psr.Value.GetObjectIds();
            if (selectedIds.Length == 0)
            {
                ed.WriteMessage("\n未选择任何图元。");
                TempExcelSyncData.Clear();
                return;
            }

            // 3. 构建查找字典：以 "管段号" 为 Key，DeviceInfo 为 Value
            // 这样可以在 O(1) 时间内找到对应行的数据
            var dataDict = new Dictionary<string, DeviceInfo>(StringComparer.OrdinalIgnoreCase);
            int skippedRows = 0;

            foreach (var item in TempExcelSyncData)
            {
                // 优先使用 "管段号"，其次使用 "Name"
                string key = string.Empty;
                if (item.Attributes != null && item.Attributes.TryGetValue("管段号", out var pipeNo))
                {
                    key = pipeNo;
                }
                else if (!string.IsNullOrWhiteSpace(item.Name))
                {
                    key = item.Name;
                }

                if (!string.IsNullOrWhiteSpace(key))
                {
                    // 如果 Key 重复，后出现的覆盖先出现的
                    dataDict[key] = item;
                }
                else
                {
                    skippedRows++;
                }
            }

            if (skippedRows > 0)
            {
                ed.WriteMessage($"\n警告：表格中有 {skippedRows} 行数据缺少'管段号'或'Name'，已被跳过。");
            }

            if (dataDict.Count == 0)
            {
                ed.WriteMessage("\n错误：表格中没有有效的可匹配数据。请检查 Excel 列名是否包含'管段号'。");
                TempExcelSyncData.Clear();
                return;
            }

            int successCount = 0;
            int failCount = 0;
            int noMatchCount = 0;

            // 4. 遍历选中的图元，更新属性
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (var id in selectedIds)
                {
                    try
                    {
                        var ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                        if (ent is BlockReference br)
                        {
                            // 获取图元的标识符（管段号）
                            string entityKey = GetEntityKeyFromBlock(br, tr);

                            if (string.IsNullOrWhiteSpace(entityKey))
                            {
                                failCount++; // 无法识别的图元
                                continue;
                            }

                            // 在表格数据中查找匹配项
                            if (dataDict.TryGetValue(entityKey, out var targetData))
                            {
                                // 匹配成功，开始更新属性
                                UpdateBlockAttributesFromExcel(br, targetData, tr);
                                successCount++;
                            }
                            else
                            {
                                // 未找到匹配数据
                                noMatchCount++;
                            }
                        }
                        else
                        {
                            // 非块参照，跳过
                            failCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        ed.WriteMessage($"\n处理图元 {id} 时出错: {ex.Message}");
                        failCount++;
                    }
                }
                tr.Commit();
            }

            ed.WriteMessage($"\n同步完成：成功更新 {successCount} 个，未匹配 {noMatchCount} 个，其他失败 {failCount} 个。");

            // 5. 清空临时数据，防止重复同步
            TempExcelSyncData.Clear();

            // 6. 刷新屏幕
            ed.Regen();
            Application.UpdateScreen();
        }

        /// <summary>
        /// 辅助：从块中获取唯一标识键（用于匹配 Excel）
        /// </summary>
        private string GetEntityKeyFromBlock(BlockReference br, Transaction tr)
        {
            // 尝试从块属性中读取 "管段号"
            foreach (ObjectId attId in br.AttributeCollection)
            {
                var ar = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (ar != null && !string.IsNullOrWhiteSpace(ar.Tag))
                {
                    // 这里可以根据你的块属性 Tag 名称进行调整
                    // 常见的 Tag 名称可能是 "管段号", "PIPE_NO", "TAG", "NAME" 等
                    if (string.Equals(ar.Tag, "管段号", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(ar.Tag, "PIPE_NO", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(ar.Tag, "PIPELINE NO", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!string.IsNullOrWhiteSpace(ar.TextString))
                            return ar.TextString.Trim();
                    }
                }
            }

            // 如果没找到管段号，可以尝试返回块名或其他标识，但通常管段号是最准确的
            return string.Empty;
        }

        /// <summary>
        /// 辅助：根据 Excel 数据更新块属性
        /// </summary>
        private void UpdateBlockAttributesFromExcel(BlockReference br, DeviceInfo data, Transaction tr)
        {
            if (data.Attributes == null || data.Attributes.Count == 0) return;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                var ar = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (ar != null && !string.IsNullOrWhiteSpace(ar.Tag))
                {
                    // 如果表格数据中包含与属性 Tag 相同的键，则更新
                    // 注意：Dictionary 初始化时用了 OrdinalIgnoreCase，所以不区分大小写
                    if (data.Attributes.TryGetValue(ar.Tag, out var newValue))
                    {
                        // 只有当值不同时才更新，提高效率
                        if (ar.TextString != newValue)
                        {
                            ar.TextString = newValue;
                            // 可选：调整对齐方式，防止文字溢出或位置偏移
                            try { ar.AdjustAlignment(br.Database); } catch { }
                        }
                    }
                }
            }
        }

        #endregion

        /// <summary>
        /// 清理属性文本
        /// </summary>
        /// <param name="raw">原始属性文本</param>
        /// <returns></returns>
        private string CleanAttributeText(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            string s = raw.Trim();

            // 1) 处理 AutoCAD 富文本包装，如 {\fSimSun|b0|i0|c134|p2;酒精}
            var m = Regex.Match(s, @"^\{\\[^;{}]*;(?<text>.*)\}$", RegexOptions.Singleline);
            if (m.Success)
            {
                s = m.Groups["text"].Value;
            }

            // 2) 移除反斜线控制序列（如 \f…, \H…, \p… 等），但保留数字、℃、°、字母和中文
            s = Regex.Replace(s, @"\\[A-Za-z]+\d*;?", string.Empty);

            // 3) 去掉外层花括号残留
            s = s.Replace("{", "").Replace("}", "");

            // 4) 去掉不可见控制字符
            s = Regex.Replace(s, @"[\x00-\x1F\x7F]", string.Empty);

            // 5) 修剪并规范化连续空白
            s = Regex.Replace(s, @"\s+", " ").Trim();

            return s;
        }
        /// <summary>
        /// 导入表格数据
        /// </summary>
        [CommandMethod(nameof(ImportTableFromExcel))]
        public void ImportTableFromExcel()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // 选择 Excel 文件
                using (var ofd = new System.Windows.Forms.OpenFileDialog())
                {
                    ofd.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                    ofd.Title = "选择要导入的 Excel 文件";
                    if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
                    string filePath = ofd.FileName;
                    if (string.IsNullOrWhiteSpace(filePath) || !System.IO.File.Exists(filePath))
                    {
                        ed.WriteMessage("\n未找到选择的文件。");
                        return;
                    }

                    // 读取 Excel 内容到内存（EPPlus）
                    List<string[]> excelData = new List<string[]>();
                    var mergedRanges = new List<(int r1, int c1, int r2, int c2)>();
                    int excelRows = 0, excelCols = 0;
                    using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
                    {
                        if (package.Workbook.Worksheets.Count == 0)
                        {
                            ed.WriteMessage("\nExcel 文件中未找到工作表。");
                            return;
                        }
                        var ws = package.Workbook.Worksheets[0]; // 使用第一个工作表
                        if (ws.Dimension == null)
                        {
                            ed.WriteMessage("\n工作表为空。");
                            return;
                        }

                        excelRows = ws.Dimension.End.Row;
                        excelCols = ws.Dimension.End.Column;

                        // 读取单元格文本
                        for (int r = 1; r <= excelRows; r++)
                        {
                            var rowArr = new string[excelCols];
                            for (int c = 1; c <= excelCols; c++)
                            {
                                var cell = ws.Cells[r, c];
                                string text = cell?.Text ?? string.Empty;
                                rowArr[c - 1] = text;
                            }
                            excelData.Add(rowArr);
                        }

                        // 收集合并单元格（Excel 地址如 "A1:C1"）
                        foreach (var addr in ws.MergedCells)
                        {
                            try
                            {
                                var a = new OfficeOpenXml.ExcelAddress(addr);
                                mergedRanges.Add((a.Start.Row - 1, a.Start.Column - 1, a.End.Row - 1, a.End.Column - 1));
                            }
                            catch
                            {
                                // 忽略无法解析的合并范围
                            }
                        }
                    }

                    // 让用户在 CAD 中选择要替换的 Table
                    PromptEntityOptions peo = new PromptEntityOptions("\n请选择要用 Excel 内容替换的 CAD 表格（Table）:");
                    peo.SetRejectMessage("\n请选择一个表格对象。");
                    peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Table), true);
                    PromptEntityResult per = ed.GetEntity(peo);
                    if (per.Status != PromptStatus.OK) return;

                    // 写回 CAD 表格 —— 仅替换数据，不改变表结构/样式/合并状态
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var table = tr.GetObject(per.ObjectId, OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.Table;
                        if (table == null)
                        {
                            ed.WriteMessage("\n选中的对象不是表格。");
                            return;
                        }

                        try
                        {
                            int cadRows = table.Rows.Count;
                            int cadCols = table.Columns.Count;

                            // 仅在重叠区域写入数据，避免调整表的行列数或样式
                            int maxR = Math.Min(cadRows, excelRows);
                            int maxC = Math.Min(cadCols, excelCols);

                            // 处理 Excel 的合并区域：只写入合并区域的左上角单元格，跳过合并区域内的其余单元（避免覆盖已合并单元）
                            var skipCell = new bool[excelRows, excelCols];
                            foreach (var m in mergedRanges)
                            {
                                int r1 = m.r1, c1 = m.c1, r2 = m.r2, c2 = m.c2;
                                if (r1 < 0 || c1 < 0 || r2 >= excelRows || c2 >= excelCols) continue;
                                for (int rr = r1; rr <= r2; rr++)
                                {
                                    for (int cc = c1; cc <= c2; cc++)
                                    {
                                        if (rr == r1 && cc == c1) continue; // 留下左上角可写
                                        skipCell[rr, cc] = true;
                                    }
                                }
                            }

                            // 将重叠区域的数据写回 CAD 表格（谨慎写入每个单元，单元写入失败时忽略以保证不修改样式）
                            for (int r = 0; r < maxR; r++)
                            {
                                for (int c = 0; c < maxC; c++)
                                {
                                    // 如果 Excel 在此处属于合并范围且不是左上角，则跳过写入（以免破坏 CAD 的合并格）
                                    if (r < excelRows && c < excelCols && skipCell[r, c])
                                        continue;

                                    string val = string.Empty;
                                    if (r < excelData.Count && c < excelData[r].Length)
                                        val = excelData[r][c] ?? string.Empty;

                                    try
                                    {
                                        // 只改 TextString，不改对齐、行高、列宽、合并等
                                        table.Cells[r, c].TextString = val;
                                    }
                                    catch
                                    {
                                        // 某些单元格可能属于 CAD 合并区域的次单元，写入会失败。忽略并继续。
                                        continue;
                                    }
                                }
                            }

                            // 如果 Excel 的第一行是合并标题并 CAD 侧已合并，则尽量保证左上角单元居中加粗，但不改变 CAD 合并结构或样式
                            try
                            {
                                if (excelRows >= 1 && excelCols >= 1)
                                {
                                    string firstCellExcel = excelData.Count > 0 && excelData[0].Length > 0 ? (excelData[0][0] ?? string.Empty).Trim() : string.Empty;
                                    if (!string.IsNullOrEmpty(firstCellExcel))
                                    {
                                        // 如果 CAD 表的第一行在视觉上是标题（例如大部分列为空），只更新左上角文本（已写入），不修改样式
                                        // 不做合并/加粗/列宽/行高调整，完全保留 CAD 端样式
                                    }
                                }
                            }
                            catch { /* 忽略 */ }

                            tr.Commit();
                            ed.WriteMessage($"\n已将 Excel ({System.IO.Path.GetFileName(filePath)}) 的数据写入选中的表格（仅覆盖重叠单元，不改动表格样式/合并/尺寸）。");
                        }
                        catch (System.Exception exInner)
                        {
                            tr.Abort();
                            ed.WriteMessage($"\n将 Excel 写入表格时出错: {exInner.Message}");
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\n导入失败: {ex.Message}");
            }
        }

        #endregion

        #region 比例相关
        /// <summary>
        /// 获取当前绘图比例（优先使用用户在WPF界面输入的值）
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="roundToCommon">是否四舍五入到常见比例</param>
        /// <returns>绘图比例分母</returns>
        private double GetDrawingScaleDenominator(Database db, bool roundToCommon = false)
        {
            // 优先从WPF界面的TextBox获取用户输入的比例值
            double userScale = GetDrawingScaleFromWpf();
            if (userScale > 0)
            {
                return roundToCommon ? RoundToCommonScale(userScale) : userScale;
            }

            // 如果WPF界面不可用或输入无效，则使用原有逻辑
            return GetScaleDenominatorForDatabase(db, roundToCommon);
        }

        /// <summary>
        /// 从WPF界面获取用户输入的绘图比例
        /// </summary>
        /// <returns>用户输入的比例值，如果获取失败返回0</returns>
        private double GetDrawingScaleFromWpf()
        {
            try
            {
                // 获取WPF主窗口实例
                var wpfWindow = AutoCadHelper.GetWpfWindow();
                if (wpfWindow != null)
                {
                    // 使用反射获取TextBox_绘图比例控件
                    var textBoxField = wpfWindow.GetType().GetField("TextBox_绘图比例",
                        System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    if (textBoxField != null)
                    {
                        var textBox = textBoxField.GetValue(wpfWindow) as System.Windows.Controls.TextBox;
                        if (textBox != null &&
                            double.TryParse(textBox.Text, out double scale) &&
                            scale > 0)
                        {
                            return scale;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 如果获取失败，返回0表示使用默认逻辑
                Application.DocumentManager.MdiActiveDocument?.Editor?.WriteMessage($"\n获取绘图比例失败: {ex.Message}");
            }

            return 0;
        }

        /// <summary>
        /// 将比例四舍五入到常见值
        /// </summary>
        /// <param name="scale">原始比例</param>
        /// <returns>四舍五入后的比例</returns>
        private double RoundToCommonScale(double scale)
        {
            // 常见的比例值
            double[] commonScales = { 1, 5, 10, 20, 25, 50, 100, 200, 500, 1000 };

            double closestScale = commonScales[0];
            double minDiff = Math.Abs(scale - commonScales[0]);

            foreach (double commonScale in commonScales)
            {
                double diff = Math.Abs(scale - commonScale);
                if (diff < minDiff)
                {
                    minDiff = diff;
                    closestScale = commonScale;
                }
            }
            return closestScale;
        }

        /// <summary>
        /// 获取数据库中的绘图比例分母
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="roundToCommon">是否四舍五入到常见比例</param>
        /// <returns>比例分母</returns>
        private double GetScaleDenominatorForDatabase(Database db, bool roundToCommon = true)
        {
            // 优先使用用户在WPF界面输入的比例值
            double userScale = GetDrawingScaleFromWpf();
            if (userScale > 0)
            {
                return roundToCommon ? RoundToCommonScale(userScale) : userScale;
            }
            // 如果用户没有输入或输入无效，则使用原有逻辑
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    // 尝试从当前视口获取比例
                    try
                    {
                        var ed = doc.Editor;
                        var cv = ed.CurrentViewportObjectId;
                        if (cv != ObjectId.Null)
                        {
                            using (var tr = db.TransactionManager.StartTransaction())
                            {
                                var vp = tr.GetObject(cv, OpenMode.ForRead) as Viewport;
                                if (vp != null)
                                {
                                    // 检查是否为布局视口（非模型空间视口）
                                    var currentSpace = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                                    if (currentSpace != null && currentSpace.IsLayout)
                                    {
                                        // 返回视口的自定义比例
                                        return vp.CustomScale;
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // 忽略错误，继续使用默认值
                    }
                }
            }
            catch
            {
                // 忽略错误
            }
            // 默认返回100（1:100）
            return roundToCommon ? 100.0 : 100.0;
        }

        /// <summary>
        /// 创建设备材料表
        /// </summary>
        /// <param name="db">当前数据库</param>
        /// <param name="deviceList">设备信息列表</param>
        /// <param name="scaleDenominator">绘图比例分母（用于计算文字高度和行高）</param>
        private void CreateDeviceTable(Database db, List<DeviceInfo> deviceList, double scaleDenominator = 0.0)
        {
            // 如果设备列表为空，直接返回
            if (deviceList == null) return;

            // 开启数据库事务
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 获取块表（只读）
                    BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    // 获取当前空间（模型空间或图纸空间，可写）
                    BlockTableRecord currentSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    // 构建动态列列表（根据所有设备的属性提取出的唯一列名）
                    var dynamicColumns = BuildDynamicColumnList(deviceList);
                    // 如果没有提取到任何列，默认至少包含“名称”列
                    if (dynamicColumns == null) dynamicColumns = new List<string> { "名称" };

                    // 计算总列数
                    int totalColumns = Math.Max(1, dynamicColumns.Count);
                    // 计算数据行数
                    int dataRows = deviceList?.Count ?? 0;
                    // 计算总行数：1行标题 + 1行表头 + N行数据
                    int totalRows = 1 + 1 + dataRows;

                    // 创建新的表格对象
                    Table table = new Table();
                    // 设置表格的行数和列数
                    table.SetSize(totalRows, totalColumns);

                    // 获取当前文档的编辑器对象，用于用户交互
                    var ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                    if (ed == null)
                    {
                        // 如果无法获取编辑器（例如在非交互环境下），直接退出以避免后续错误
                        return;
                    }

                    // 提示用户在 CAD 中指定表格的插入位置
                    PromptPointResult ppr = ed.GetPoint("\n指定插入位置: ");
                    // 如果用户成功指定了点
                    if (ppr.Status == PromptStatus.OK)
                    {
                        // 设置表格的插入点
                        table.Position = ppr.Value;
                    }

                    // 确定有效的比例分母：如果传入值有效则使用传入值，否则从数据库中获取当前绘图比例
                    double effectiveScaleDenom = scaleDenominator > 0.0 ? scaleDenominator : GetDrawingScaleDenominator(db, true);

                    // 设置表格的样式（包括文字高度、行高、边框等，基于比例分母计算）
                    SetTableStyle(db, table, trans, effectiveScaleDenom);

                    // --- 生成表格标题 ---
                    string titleName = string.Empty;
                    // 尝试从第一个设备的属性或名称中提取标题依据
                    if (deviceList != null && deviceList.Count > 0)
                    {
                        var first = deviceList[0];
                        // 优先查找“名称”属性
                        if (first.Attributes != null && first.Attributes.TryGetValue("名称", out var nv) && !string.IsNullOrWhiteSpace(nv))
                            titleName = nv;
                        // 其次使用 DeviceInfo 的 Name 字段
                        else if (!string.IsNullOrWhiteSpace(first.Name))
                            titleName = first.Name;
                    }
                    // 如果仍未获取到名称，使用默认值“设备”
                    if (string.IsNullOrWhiteSpace(titleName))
                        titleName = "设备";

                    // 提取标题中的中文字符，用于生成更规范的中文标题
                    string chinese = ExtractChineseCharacters(titleName);
                    if (string.IsNullOrWhiteSpace(chinese))
                        chinese = titleName;

                    // 组合最终标题，格式如：“泵 - 材料明细表”
                    string fullTitle = $"{chinese} - 材料明细表";

                    // 合并第一行的所有单元格作为标题栏
                    table.MergeCells(CellRange.Create(table, 0, 0, 0, totalColumns - 1));
                    // 设置标题文本
                    table.Cells[0, 0].TextString = fullTitle;
                    // 设置标题居中对齐
                    table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

                    // --- 填充表头（第二行）---
                    for (int c = 0; c < dynamicColumns.Count && c < table.Columns.Count; c++)
                    {
                        // 获取当前列的键名（中文）
                        string key = dynamicColumns[c];
                        // 查找对应的英文翻译，如果没有则使用原键名
                        string english = DictionaryHelper.ChineseToEnglish.ContainsKey(key) ? DictionaryHelper.ChineseToEnglish[key] : key;

                        // 设置表头文本：如果有英文对照，则显示为“中文\n英文”，否则只显示中文
                        table.Cells[1, c].TextString = string.Equals(english, key) ? key : (key + "\n" + english);
                        // 设置表头居中对齐
                        table.Cells[1, c].Alignment = CellAlignment.MiddleCenter;
                    }

                    // --- 填充数据行 ---
                    int dataStart = 2; // 数据从第3行开始（索引2）
                    for (int r = 0; r < dataRows; r++)
                    {
                        // 获取当前行的设备对象
                        var item = deviceList[r];
                        // 计算当前行在表格中的索引
                        int rowIndex = dataStart + r;

                        // 遍历每一列填充数据
                        for (int c = 0; c < dynamicColumns.Count && c < table.Columns.Count; c++)
                        {
                            // 获取当前列的键名
                            string colKey = dynamicColumns[c];
                            // 根据键名从设备属性中获取对应的值
                            string val = GetAttributeValueByMappedKey(item.Attributes, colKey);

                            // 特殊处理：如果列是“名称”且值为空，使用 item.Name
                            if (string.Equals(colKey, "名称", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(val))
                                val = item.Name ?? string.Empty;

                            // 特殊处理：如果列是“数量”且值为空，使用 item.Count
                            if (string.Equals(colKey, "数量", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(val))
                                val = item.Count > 0 ? item.Count.ToString() : string.Empty;

                            // 将值填入单元格
                            table.Cells[rowIndex, c].TextString = val ?? string.Empty;
                        }
                    }

                    // 尝试应用基于比例的行列高度调整
                    try { ApplyScaledHeightsToTable(table, effectiveScaleDenom); } catch { }

                    // 【关键修改】在事务提交前，应用高级自动列宽调整
                    // 这行代码确保了设备表也能像管道表一样，根据内容自动调整列宽，处理合并单元格
                    AutoFitTableColumnsAdvanced(table);

                    // 强制更新表格布局，使上述样式和宽度更改生效
                    table.GenerateLayout();

                    // 将表格添加到当前空间
                    currentSpace.AppendEntity(table);
                    // 将表格对象注册到事务中
                    trans.AddNewlyCreatedDBObject(table, true);

                    // 提交事务，保存更改
                    trans.Commit();
                }
                catch
                {
                    // 如果发生异常，回滚事务
                    trans.Abort();
                    // 重新抛出异常以便上层捕获
                    throw;
                }
            }
        }

        /// <summary>
        /// 创建指定类型的设备/管道表格，并自动调整列宽以适应内容
        /// </summary>
        /// <param name="db">当前数据库</param>
        /// <param name="deviceList">设备或管道信息列表</param>
        /// <param name="typeTitle">表格标题（如“管道明细”）</param>
        /// <param name="scaleDenominator">比例分母，用于计算文字高度和行高</param>
        public void CreateDeviceTableWithType(Database db, List<DeviceInfo> deviceList, string typeTitle, double scaleDenominator = 0.0)
        {
            // 如果列表为空，直接返回，不创建空表
            if (deviceList == null || deviceList.Count == 0) return;

            // 获取当前活动文档
            var doc = Application.DocumentManager.MdiActiveDocument;
            // 如果文档为空，返回
            if (doc == null) return;
            // 获取编辑器对象
            var ed = doc.Editor;

            // 提示用户指定插入位置
            PromptPointOptions ppo = new PromptPointOptions($"\n'{typeTitle}' 表：指定插入位置（点击或输入点）：");
            // 不允许直接回车跳过
            ppo.AllowNone = false;
            // 获取用户输入的点位
            var ppr = ed.GetPoint(ppo);
            // 如果用户取消或未指定点位
            if (ppr.Status != PromptStatus.OK)
            {
                // 输出提示信息
                ed.WriteMessage("\n未指定插入位置，跳过该类型表的插入。");
                // 结束方法
                return;
            }
            // 获取插入点坐标
            Point3d insertPosition = ppr.Value;

            // 定义基础固定列数（索引 0..9）
            const int baseFixedCols = 10;
            // 定义管段号候选键数组
            string[] pipeNoKeys = new[] { "管段号", "管道号", "管段编号", "Pipeline No", "Pipeline", "Pipe No", "No" };
            // 定义起点候选键数组
            string[] startKeys = new[] { "起点", "始点", "From" };
            // 定义终点候选键数组
            string[] endKeys = new[] { "终点", "止点", "To" };

            // 迁移 "介质" -> "介质名称"，确保字段统一
            foreach (var e in deviceList)
            {
                // 如果属性字典为空，跳过
                if (e.Attributes == null) continue;
                // 如果存在"介质"且不为空
                if (e.Attributes.TryGetValue("介质", out var val) && !string.IsNullOrWhiteSpace(val))
                {
                    // 如果"介质名称"不存在或为空，则赋值
                    if (!e.Attributes.ContainsKey("介质名称") || string.IsNullOrWhiteSpace(e.Attributes["介质名称"]))
                        e.Attributes["介质名称"] = val;
                }
                // 如果"介质名称"仍不存在，尝试其他英文键
                if (!e.Attributes.ContainsKey("介质名称"))
                {
                    // 尝试 "Medium"
                    if (e.Attributes.TryGetValue("Medium", out var mv) && !string.IsNullOrWhiteSpace(mv)) e.Attributes["介质名称"] = mv;
                    // 尝试 "Medium Name"
                    else if (e.Attributes.TryGetValue("Medium Name", out var mn) && !string.IsNullOrWhiteSpace(mn)) e.Attributes["介质名称"] = mn;
                }
            }

            // 收集所有属性键，用于动态生成列
            var allAttrKeysSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var e in deviceList)
            {
                // 如果属性字典为空，跳过
                if (e.Attributes == null) continue;
                // 遍历所有键
                foreach (var k in e.Attributes.Keys)
                {
                    // 忽略空键
                    if (string.IsNullOrWhiteSpace(k)) continue;
                    // 添加到集合中
                    allAttrKeysSet.Add(k);
                }
            }

            // 定义需要移除的保留/固定列相关键
            var reservedKeySubstrings = new[]
            {
                "管道标题","管段号","起点","始点","终点","止点","管道等级",
                "介质","介质名称","Medium","Medium Name","操作温度","操作压力",
                "隔热隔声代号","是否防腐","Length","长度"
            };
            // 从动态列集合中移除这些保留键
            foreach (var key in allAttrKeysSet.ToList())
            {
                // 如果键包含任何保留子字符串
                if (reservedKeySubstrings.Any(s => !string.IsNullOrWhiteSpace(s) &&
                    key.IndexOf(s, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    // 从集合中移除
                    allAttrKeysSet.Remove(key);
                }
            }

            // 定义管道组首选字段顺序
            var pipeGroupPreferred = new[] { "名称", "材料", "图号或标准号", "数量", "泵前/后" };
            // 创建管道组列列表
            var pipeGroupColumns = new List<string>();
            // 按首选顺序添加存在的键
            foreach (var pk in pipeGroupPreferred)
            {
                // 如果集合中包含该键
                if (allAttrKeysSet.Contains(pk))
                {
                    // 添加到列表
                    pipeGroupColumns.Add(pk);
                    // 从集合中移除，避免重复
                    allAttrKeysSet.Remove(pk);
                }
            }
            // 确保首选字段即使不在原始数据中也占位（可选，这里保持原逻辑）
            foreach (var pk in pipeGroupPreferred)
            {
                // 如果列表中还没有
                if (!pipeGroupColumns.Contains(pk))
                    // 添加进去（这可能导致空列，视需求而定，原代码如此）
                    pipeGroupColumns.Add(pk);
            }

            // 处理剩余属性，按字母排序
            var remainingAttrs = allAttrKeysSet.OrderBy(k => k, StringComparer.OrdinalIgnoreCase).ToList();
            // 再次过滤，确保不包含首选字段
            remainingAttrs = remainingAttrs.Where(k => !pipeGroupPreferred.Any(pk => string.Equals(pk, k, StringComparison.OrdinalIgnoreCase))).ToList();
            // 将剩余属性追加到管道组列后面
            pipeGroupColumns.AddRange(remainingAttrs);

            // 计算管道组列的数量
            int pipeGroupCount = pipeGroupColumns.Count;

            // 其余动态列（通常为空，原代码保留但未使用，这里保持兼容）
            var restKeys = new List<string>();

            // 局部函数：查找第一个匹配的属性值
            string FindFirstAttrValue(Dictionary<string, string>? attrs, string[] candidates)
            {
                // 如果字典为空，返回空字符串
                if (attrs == null) return string.Empty;
                // 精确匹配
                foreach (var c in candidates)
                {
                    if (attrs.TryGetValue(c, out var v) && !string.IsNullOrWhiteSpace(v)) return v;
                }
                // 模糊匹配
                foreach (var kv in attrs)
                {
                    if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                    foreach (var c in candidates)
                    {
                        if (kv.Key.IndexOf(c, StringComparison.OrdinalIgnoreCase) >= 0 && !string.IsNullOrWhiteSpace(kv.Value))
                            return kv.Value;
                    }
                }
                // 未找到返回空
                return string.Empty;
            }

            // 局部函数：解析管段号中的数字部分
            int ParsePipeNumberNumeric(string? s)
            {
                // 如果为空，返回最大值
                if (string.IsNullOrWhiteSpace(s)) return int.MaxValue;
                // 正则匹配数字
                var m = Regex.Match(s!, @"\d+");
                // 如果匹配成功且能解析为整数
                if (m.Success && int.TryParse(m.Value, out int v)) return v;
                // 否则返回最大值
                return int.MaxValue;
            }

            // 对设备列表进行排序，依据管段号数字部分
            var sortedDeviceList = deviceList
                .Select((e, idx) => new { Item = e, OrigIndex = idx })
                .OrderBy(x =>
                {
                    // 获取管段号文本
                    var txt = FindFirstAttrValue(x.Item.Attributes, pipeNoKeys);
                    // 解析数字
                    int num = ParsePipeNumberNumeric(txt);
                    // 返回排序元组
                    return (num, x.OrigIndex);
                })
                .Select(x => x.Item)
                .ToList();

            // 锁定文档并开始事务
            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 获取块表
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    // 获取当前空间（模型空间或图纸空间）
                    var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    // 计算总列数
                    int fixedCols = baseFixedCols + pipeGroupCount;
                    int dynamicCols = restKeys.Count;
                    int totalCols = fixedCols + dynamicCols;
                    // 计算数据行数
                    int dataRows = sortedDeviceList.Count;
                    // 计算总行数（标题1 + 表头2 + 数据行）
                    int rows = 1 + 2 + dataRows; // 注意：原代码是1+3，这里根据Fill_PipeLine_FixedHeaders通常占用2行表头调整，若原逻辑是3行表头请改回3

                    // 创建表格对象
                    var table = new Autodesk.AutoCAD.DatabaseServices.Table();
                    // 设置表格尺寸
                    table.SetSize(rows, Math.Max(1, totalCols));
                    // 设置表格插入位置
                    table.Position = insertPosition;

                    // 获取初始比例分母
                    double initialScaleDenom = scaleDenominator > 0.0 ? scaleDenominator : GetDrawingScaleDenominator(db, true);
                    // 设置表格样式（包括文字高度等）
                    SetTableStyle(db, table, tr, initialScaleDenom);

                    // 合并标题行单元格
                    table.MergeCells(CellRange.Create(table, 0, 0, 0, table.Columns.Count - 1));
                    // 设置标题文本
                    table.Cells[0, 0].TextString = $"{typeTitle} - 设备材料明细表";
                    // 设置标题对齐方式
                    table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

                    // 填充固定的表头（前10列）
                    Fill_PipeLine_FixedHeaders(table, pipeGroupCount);

                    // 填充管道组动态列的表头
                    int pipeStart = baseFixedCols;
                    for (int i = 0; i < pipeGroupColumns.Count && (pipeStart + i) < table.Columns.Count; i++)
                    {
                        // 获取列名
                        string header = pipeGroupColumns[i];
                        // 获取英文对照
                        string english = DictionaryHelper.ChineseToEnglish.ContainsKey(header) ? DictionaryHelper.ChineseToEnglish[header] : header;
                        // 设置表头文本（中文+英文）
                        table.Cells[2, pipeStart + i].TextString = header + (english == header ? "" : "\n" + english);
                    }

                    // 填充其余动态列的表头（如果有的话）
                    for (int i = 0; i < restKeys.Count; i++)
                    {
                        var key = restKeys[i];
                        int col = fixedCols + i;
                        if (col < table.Columns.Count)
                        {
                            table.Cells[1, col].TextString = key;
                            table.Cells[2, col].TextString = (DictionaryHelper.ChineseToEnglish.ContainsKey(key) ? DictionaryHelper.ChineseToEnglish[key] : key);
                        }
                    }

                    // 开始填充数据行
                    int dataStartRow = 3; // 假设表头占用了0,1,2三行
                    // 遍历排序后的设备/管道列表，逐行填充表格数据
                    for (int r = 0; r < sortedDeviceList.Count; r++)
                    {
                        // 获取当前行的数据对象
                        var item = sortedDeviceList[r];
                        // 计算当前数据在表格中的实际行索引（起始行 + 偏移量）
                        int rowIndex = dataStartRow + r;

                        // ================== 填充第0列：管道标题 ==================
                        string title = null;
                        // 优先从属性字典中查找“管道标题”
                        if (item.Attributes != null && item.Attributes.TryGetValue("管道标题", out var tv) && !string.IsNullOrWhiteSpace(tv))
                            title = tv;
                        else
                        {
                            // 如果属性中没有，则从 Name 字段截取（通常格式为 PIPE_XXX，截取后半部分）
                            var nm = item.Name ?? string.Empty;
                            int pos = nm.LastIndexOf('_');
                            // 如果存在下划线且后面有内容，则截取；否则使用完整名称
                            title = pos >= 0 && pos < nm.Length - 1 ? nm.Substring(pos + 1) : nm;
                        }
                        // 将标题写入表格第0列
                        table.Cells[rowIndex, 0].TextString = title ?? string.Empty;

                        // ================== 填充第1列：管段号 ==================
                        string pipeNoVal = string.Empty;
                        // 尝试从属性中根据预设的关键字数组查找管段号
                        if (item.Attributes != null)
                            pipeNoVal = FindFirstAttrValue(item.Attributes, pipeNoKeys);

                        // 如果仍未找到，尝试从标题中提取管段号（例如从 "350-AR-1002" 中提取 "AR-1002"）
                        if (string.IsNullOrWhiteSpace(pipeNoVal) && !string.IsNullOrWhiteSpace(title))
                        {
                            pipeNoVal = ExtractPipeCodeFromTitle(title);
                            // 如果提取失败，尝试提取标题中的纯数字作为备用
                            if (string.IsNullOrWhiteSpace(pipeNoVal))
                            {
                                var m = Regex.Match(title, @"\d+");
                                if (m.Success) pipeNoVal = m.Value;
                            }
                        }

                        // 最后的兜底策略：尝试其他常见的管段号键名
                        if (string.IsNullOrWhiteSpace(pipeNoVal) && item.Attributes != null)
                        {
                            var fallbackKeys = new[] { "管段编号", "Pipeline No", "Pipeline", "Pipe No", "No" };
                            foreach (var k in fallbackKeys)
                            {
                                if (item.Attributes.TryGetValue(k, out var v) && !string.IsNullOrWhiteSpace(v))
                                {
                                    pipeNoVal = v;
                                    break;
                                }
                            }
                        }

                        // 如果最终找到了管段号，写入表格第1列
                        if (!string.IsNullOrWhiteSpace(pipeNoVal))
                            table.Cells[rowIndex, 1].TextString = pipeNoVal;

                        // ================== 填充第2列：起点 ==================
                        // 从属性中查找起点信息（支持“起点”、“始点”、“From”等键）
                        var startVal = FindFirstAttrValue(item.Attributes, startKeys);
                        if (!string.IsNullOrWhiteSpace(startVal)) table.Cells[rowIndex, 2].TextString = startVal;

                        // ================== 填充第3列：终点 ==================
                        // 从属性中查找终点信息（支持“终点”、“止点”、“To”等键）
                        var endVal = FindFirstAttrValue(item.Attributes, endKeys);
                        if (!string.IsNullOrWhiteSpace(endVal)) table.Cells[rowIndex, 3].TextString = endVal;

                        // ================== 填充第4列：管道等级 ==================
                        string pipeClass = string.Empty;
                        // 优先直接查找“管道等级”属性
                        if (item.Attributes != null && item.Attributes.TryGetValue("管道等级", out var cls) && !string.IsNullOrWhiteSpace(cls))
                        {
                            pipeClass = cls;
                        }
                        else
                        {
                            // 如果没找到，尝试从标题中提取等级（例如从 "...-1.0G11" 中提取 "1.0G11"）
                            pipeClass = ExtractPipeClassFromTitle(title);
                            // 如果标题中也没有，尝试查找其他可能的等级键名
                            if (string.IsNullOrWhiteSpace(pipeClass) && item.Attributes != null)
                            {
                                var fallbackKeys = new[] { "等级", "Class", "管级", "级别" };
                                foreach (var fk in fallbackKeys)
                                {
                                    if (item.Attributes.TryGetValue(fk, out var fv) && !string.IsNullOrWhiteSpace(fv))
                                    {
                                        pipeClass = fv;
                                        break;
                                    }
                                }
                            }
                            // 如果成功提取或找到等级，将其回填到属性字典中，方便后续使用
                            if (!string.IsNullOrWhiteSpace(pipeClass) && item.Attributes != null)
                                item.Attributes["管道等级"] = pipeClass;
                        }
                        // 写入表格第4列
                        if (!string.IsNullOrWhiteSpace(pipeClass))
                            table.Cells[rowIndex, 4].TextString = pipeClass;

                        // ================== 填充第5列：介质名称 ==================
                        // 查找介质相关信息
                        var mediumVal = FindFirstAttrValue(item.Attributes, new[] { "介质名称", "介质", "Medium Name" });
                        if (!string.IsNullOrWhiteSpace(mediumVal)) table.Cells[rowIndex, 5].TextString = mediumVal;

                        // ================== 填充第6列：操作温度 ==================
                        string opTemp = string.Empty;
                        if (item.Attributes != null)
                        {
                            // 智能查找温度键：支持包含“操作温度”、“T(”、“℃”、“°C”的键名
                            var tempKey = item.Attributes.Keys.FirstOrDefault(k => !string.IsNullOrWhiteSpace(k) &&
                                (k.IndexOf("操作温度", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                k.IndexOf("温度", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 k.IndexOf("T(", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 k.IndexOf("℃", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 k.IndexOf("°C", StringComparison.OrdinalIgnoreCase) >= 0));
                            if (!string.IsNullOrWhiteSpace(tempKey))
                                opTemp = item.Attributes[tempKey];
                        }
                        // 如果智能查找未果，尝试精确匹配“操作温度”
                        if (!string.IsNullOrWhiteSpace(opTemp))
                            table.Cells[rowIndex, 6].TextString = opTemp;
                        else if (item.Attributes != null && item.Attributes.TryGetValue("操作温度", out var tval) && !string.IsNullOrWhiteSpace(tval))
                            table.Cells[rowIndex, 6].TextString = tval;

                        // ================== 填充第7列：操作压力 ==================
                        // 智能查找压力键：支持包含“操作压力”、“P(”、“MPa”的键名
                        var pressureKey = item.Attributes.Keys.FirstOrDefault(k => !string.IsNullOrWhiteSpace(k) &&
                            (k.IndexOf("操作压力", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            k.IndexOf("压力", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             k.IndexOf("压力等级", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             k.IndexOf("P(MPaG)", StringComparison.OrdinalIgnoreCase) >= 0));
                        if (!string.IsNullOrWhiteSpace(pressureKey))
                            table.Cells[rowIndex, 7].TextString = item.Attributes[pressureKey];

                        // ================== 填充第8列：隔热隔声代号 ==================
                        if (item.Attributes != null && item.Attributes.TryGetValue("隔热隔声代号", out var code) && !string.IsNullOrWhiteSpace(code))
                            table.Cells[rowIndex, 8].TextString = code;

                        // ================== 填充第9列：是否防腐 ==================
                        if (item.Attributes != null && item.Attributes.TryGetValue("是否防腐", out var anti) && !string.IsNullOrWhiteSpace(anti))
                            table.Cells[rowIndex, 9].TextString = anti;

                        // ================== 填充管道组动态列（特定逻辑处理） ==================
                        for (int i = 0; i < pipeGroupColumns.Count; i++)
                        {
                            // 计算当前动态列在表格中的实际列索引
                            int col = baseFixedCols + i;
                            // 防止列索引越界
                            if (col >= table.Columns.Count) break;

                            // 获取当前列对应的属性键名
                            var headerKey = pipeGroupColumns[i];

                            // --- 特殊处理1：核算流速 ---
                            if (string.Equals(headerKey, "核算流速", StringComparison.OrdinalIgnoreCase))
                            {
                                var flows = new List<string>();
                                if (item.Attributes != null)
                                {
                                    // 遍历所有属性，收集包含“流速”、“流量”或“flow”的值
                                    foreach (var kv in item.Attributes)
                                    {
                                        if (kv.Key.IndexOf("流速", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                            kv.Key.IndexOf("流量", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                            kv.Key.IndexOf("flow", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            // 避免重复添加
                                            if (!string.IsNullOrWhiteSpace(kv.Value) && !flows.Contains(kv.Value)) flows.Add(kv.Value);
                                        }
                                    }
                                }
                                // 如果属性中有明确的“核算流速”，优先放在最前面
                                if (item.Attributes != null && item.Attributes.TryGetValue("核算流速", out var hv) && !string.IsNullOrWhiteSpace(hv) && !flows.Contains(hv))
                                    flows.Insert(0, hv);

                                // 将收集到的流速值用逗号连接后填入单元格
                                if (flows.Count > 0) table.Cells[rowIndex, col].TextString = string.Join(", ", flows);
                                continue;
                            }

                            // --- 特殊处理2：数量 ---
                            if (string.Equals(headerKey, "数量", StringComparison.OrdinalIgnoreCase))
                            {
                                // 优先使用属性中的“数量”，如果没有则使用 DeviceInfo 自带的 Count 属性
                                if (item.Attributes != null && item.Attributes.TryGetValue("数量", out var qv) && !string.IsNullOrWhiteSpace(qv))
                                    table.Cells[rowIndex, col].TextString = qv;
                                else
                                    table.Cells[rowIndex, col].TextString = item.Count.ToString();
                                continue;
                            }

                            // --- 特殊处理3：图号或标准号 ---
                            if (string.Equals(headerKey, "图号或标准号", StringComparison.OrdinalIgnoreCase))
                            {
                                string matched = string.Empty;
                                if (item.Attributes != null)
                                {
                                    // 按优先级查找：标准 -> 标准号 -> 图号 -> DWG.No. -> STD.No.
                                    if (item.Attributes.TryGetValue("标准号", out var stdVal) && !string.IsNullOrWhiteSpace(stdVal)) matched = stdVal;
                                    if (string.IsNullOrWhiteSpace(matched))
                                    {
                                        if (item.Attributes.TryGetValue("图号", out var stdVal2) && !string.IsNullOrWhiteSpace(stdVal2)) matched = stdVal2;
                                    }
                                    if (string.IsNullOrWhiteSpace(matched))
                                    {
                                        if (item.Attributes.TryGetValue("标准號", out var dwgVal) && !string.IsNullOrWhiteSpace(dwgVal)) matched = dwgVal;
                                        else if (item.Attributes.TryGetValue("DWG.No.", out var dwgDot) && !string.IsNullOrWhiteSpace(dwgDot)) matched = dwgDot;
                                        else if (item.Attributes.TryGetValue("STD.No.", out var stdDot) && !string.IsNullOrWhiteSpace(stdDot)) matched = stdDot;
                                    }

                                    // 如果上述精确匹配都失败，进行模糊匹配（键名包含“标准”、“图号”、“DWG”、“STD”）
                                    if (string.IsNullOrWhiteSpace(matched))
                                    {
                                        foreach (var kv in item.Attributes)
                                        {
                                            if (string.IsNullOrWhiteSpace(kv.Key) || string.IsNullOrWhiteSpace(kv.Value)) continue;
                                            if (kv.Key.IndexOf("标准", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                kv.Key.IndexOf("图号", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                kv.Key.IndexOf("DWG", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                kv.Key.IndexOf("STD", StringComparison.OrdinalIgnoreCase) >= 0)
                                            {
                                                matched = kv.Value;
                                                break;
                                            }
                                        }
                                    }
                                }

                                // 写入匹配到的图号或标准号
                                if (!string.IsNullOrWhiteSpace(matched))
                                    table.Cells[rowIndex, col].TextString = matched;

                                continue;
                            }

                            // --- 通用处理：其他动态列 ---
                            string matchedValue = string.Empty;
                            if (item.Attributes != null)
                            {
                                // 首先尝试精确匹配键名
                                if (item.Attributes.TryGetValue(headerKey, out var dv) && !string.IsNullOrWhiteSpace(dv))
                                    matchedValue = dv;
                                else
                                {
                                    // 如果精确匹配失败，尝试模糊匹配（键名互相包含）
                                    foreach (var kv in item.Attributes)
                                    {
                                        if (string.IsNullOrWhiteSpace(kv.Key) || string.IsNullOrWhiteSpace(kv.Value)) continue;
                                        if (kv.Key.IndexOf(headerKey, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                            headerKey.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            matchedValue = kv.Value;
                                            break;
                                        }
                                    }
                                }
                            }
                            // 写入匹配到的值
                            if (!string.IsNullOrWhiteSpace(matchedValue))
                                table.Cells[rowIndex, col].TextString = matchedValue;
                        }

                        // ================== 填充其余动态列（剩余属性） ==================
                        for (int di = 0; di < restKeys.Count; di++)
                        {
                            // 计算列索引
                            int col = fixedCols + di;
                            // 防止列索引越界
                            if (col >= table.Columns.Count) break;

                            // 获取剩余的属性键
                            var key = restKeys[di];
                            // 如果属性中存在该键且值不为空，则填入表格
                            if (item.Attributes != null && item.Attributes.TryGetValue(key, out var v) && !string.IsNullOrWhiteSpace(v))
                            {
                                table.Cells[rowIndex, col].TextString = v;
                            }
                        }
                    }

                    // ================== 关键修改区域 ==================

                    // 计算最终应用的比例分母
                    double appliedScaleDenom = scaleDenominator > 0.0 ? scaleDenominator : initialScaleDenom;
                    if (appliedScaleDenom <= 0.0) appliedScaleDenom = GetDrawingScaleDenominator(db, true);

                    // 输出日志
                    ed.WriteMessage($"\n插入表格时使用的比例分母: {appliedScaleDenom}");

                    // 第一步：先应用缩放后的文字高度和行高
                    // 必须在调整列宽之前执行，因为列宽计算依赖 TextHeight
                    try { ApplyScaledHeightsToTable(table, appliedScaleDenom); } catch { }

                    // 第二步：自动调整列宽以适应内容
                    // 此时 TextHeight 已正确设置，计算结果更准确
                    //AutoResizeColumns(table);

                    // ================== 结束关键修改区域 ==================

                    // 【新增】在事务提交前，自动调整列宽
                    AutoFitTableColumnsAdvanced(table);
                    // 将表格添加到当前空间
                    currentSpace.AppendEntity(table);
                    // 将表格对象加入事务管理
                    tr.AddNewlyCreatedDBObject(table, true);

                    // 提交事务
                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    // 发生异常时中止事务
                    tr.Abort();
                    // 输出错误信息
                    ed.WriteMessage($"\n创建表格失败: {ex.Message}");
                    // 重新抛出异常
                    throw;
                }
            }

            // 刷新屏幕显示
            try
            {
                ed.Regen();
                Application.UpdateScreen();
            }
            catch { }
        }

        /// <summary>
        /// 自动调整表格列宽，使每列宽度适应其内容中最长的文字
        /// </summary>
        /// <param name="table">要调整的表格对象</param>
        private void AutoResizeColumns(Autodesk.AutoCAD.DatabaseServices.Table table)
        {
            // 获取表格的总行数
            int numRows = table.Rows.Count;
            // 获取表格的总列数
            int numCols = table.Columns.Count;

            // 遍历每一列，计算该列所需的最大宽度
            for (int col = 0; col < numCols; col++)
            {
                // 初始化当前列的最大宽度为最小值
                double maxWidth = 0.0;
                // 初始化一个默认的文字高度，防止单元格高度未设置时出错
                double defaultTextHeight = 2.5;

                // 遍历该列的每一行
                for (int row = 0; row < numRows; row++)
                {
                    // 获取当前单元格的文本内容
                    string cellText = table.Cells[row, col].TextString;

                    // 如果单元格为空，跳过
                    if (string.IsNullOrWhiteSpace(cellText)) continue;

                    // 获取当前单元格的文字高度
                    var textHeight = Convert.ToDouble(table.Cells[row, col].TextHeight);

                    // 如果文字高度无效，使用默认值
                    if (textHeight <= 0) textHeight = defaultTextHeight;

                    // 估算文本宽度
                    // 简单算法：字符数 * 文字高度 * 系数
                    double estimatedWidth = 0.0;

                    // 检查是否包含中文字符
                    bool hasChinese = System.Text.RegularExpressions.Regex.IsMatch(cellText, @"[\u4e00-\u9fff]");

                    if (hasChinese)
                    {
                        // 如果有中文，使用较大系数 1.1（汉字较宽）
                        estimatedWidth = cellText.Length * textHeight * 1.1;
                    }
                    else
                    {
                        // 如果全是英文/数字，使用较小系数 0.7
                        estimatedWidth = cellText.Length * textHeight * 0.7;
                    }

                    // 更新当前列的最大宽度
                    if (estimatedWidth > maxWidth)
                    {
                        maxWidth = estimatedWidth;
                    }
                }

                // 设置列宽
                // 增加一定的边距（Padding），例如左右各加 1.0 倍文字高度，避免文字贴边
                double padding = defaultTextHeight * 2.0;

                // 如果计算出的最大宽度大于0，则设置列宽
                if (maxWidth > 0)
                {
                    // 设置列宽为最大内容宽度加上边距
                    table.Columns[col].Width = maxWidth + padding;
                }
                else
                {
                    // 如果列中没有有效内容，设置一个默认最小宽度
                    table.Columns[col].Width = defaultTextHeight * 10.0;
                }
            }

            // 强制重新生成表格布局，使列宽生效
            table.GenerateLayout();
        }
        #endregion
    }
}

/// <summary>
/// 属性编辑表单
/// </summary>
public partial class AttributeForm : Form
{
    /// <summary>
    /// 块表记录
    /// </summary>
    private BlockTableRecord _blockTableRecord; // 关联的块定义
    /// <summary>
    /// 数据表格
    /// </summary>
    private DataGridView dataGridView; // 数据表格控件
    /// <summary>
    /// 添加行按钮
    /// </summary>
    private Button btnAddRow; // 添加行按钮
    /// <summary>
    /// 删除行按钮
    /// </summary>
    private Button btnDeleteRow; // 删除行按钮
    /// <summary>
    /// 保存按钮
    /// </summary>
    private Button btnSave; // 保存按钮
    /// <summary>
    /// 取消按钮
    /// </summary>
    private Button btnCancel; // 取消按钮
    /// <summary>
    /// 数据表格
    /// </summary>
    private DataTable _dataTable; // 存储属性数据的表格
    /// <summary>
    /// 公开数据表格供外部访问
    /// </summary>
    public DataTable DataTable => _dataTable;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="blockTableRecord"></param>
    /// <param name="blockDataTables"></param>
    public AttributeForm(BlockTableRecord blockTableRecord, Dictionary<string, DataTable> blockDataTables)
    {
        InitializeComponent();
        _blockTableRecord = blockTableRecord;

        // 加载已有数据或初始化新表格
        if (blockDataTables.ContainsKey(_blockTableRecord.Name))
        {
            _dataTable = blockDataTables[_blockTableRecord.Name].Copy();
        }
        else
        {
            _dataTable = new DataTable();
            _dataTable.Columns.Add("序号", typeof(int));
            _dataTable.Columns.Add("部件ID", typeof(string));
            _dataTable.Columns.Add("部件名", typeof(string));
            _dataTable.Columns.Add("参数", typeof(string));

            // 添加初始行
            DataRow row = _dataTable.NewRow();
            row["序号"] = _dataTable.Rows.Count + 1;
            row["部件ID"] = $"id{_dataTable.Rows.Count + 1:D4}";
            _dataTable.Rows.Add(row);
        }
        if (dataGridView is null)
        {
            return;
        }
        else
        {
            // 绑定数据到表格控件
            dataGridView.DataSource = _dataTable;

            // 设置列标题
            dataGridView.Columns["序号"].HeaderText = "序号";
            dataGridView.Columns["部件ID"].HeaderText = "部件ID";
            dataGridView.Columns["部件名"].HeaderText = "部件名";
            dataGridView.Columns["参数"].HeaderText = "参数";
            // 设置序号列为只读
            dataGridView.Columns["序号"].ReadOnly = true;
            dataGridView.Columns["部件ID"].ReadOnly = true;
        }

    }

    /// <summary>
    /// 初始化表单控件
    /// </summary>
    private void InitializeComponent()
    {
        this.dataGridView = new DataGridView();
        this.btnAddRow = new Button();
        this.btnDeleteRow = new Button();
        this.btnSave = new Button();
        this.btnCancel = new Button();
        ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
        this.SuspendLayout();

        // 数据表格控件设置
        this.dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        this.dataGridView.Location = new Point(12, 12);
        this.dataGridView.Size = new Size(560, 300);
        this.dataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        this.dataGridView.AllowUserToAddRows = false;
        this.dataGridView.AllowUserToDeleteRows = false;
        this.dataGridView.ReadOnly = false;
        this.dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

        // 添加行按钮
        this.btnAddRow.Text = "添加一行属性";
        this.btnAddRow.Location = new Point(12, 318);
        this.btnAddRow.Size = new Size(120, 30);
        this.btnAddRow.Click += new EventHandler(this.btnAddRow_Click);

        // 删除行按钮
        this.btnDeleteRow.Text = "删除选定属性";
        this.btnDeleteRow.Location = new Point(142, 318);
        this.btnDeleteRow.Size = new Size(120, 30);
        this.btnDeleteRow.Click += new EventHandler(this.btnDeleteRow_Click);

        // 保存按钮
        this.btnSave.Text = "确定";
        this.btnSave.DialogResult = DialogResult.OK;
        this.btnSave.Location = new Point(392, 318);
        this.btnSave.Size = new Size(80, 30);
        this.btnSave.Click += new EventHandler(this.btnSave_Click);

        // 取消按钮
        this.btnCancel.Text = "取消";
        this.btnCancel.DialogResult = DialogResult.Cancel;
        this.btnCancel.Location = new Point(482, 318);
        this.btnCancel.Size = new Size(80, 30);

        // 表单设置
        this.ClientSize = new Size(584, 361);
        this.Controls.Add(this.dataGridView);
        this.Controls.Add(this.btnAddRow);
        this.Controls.Add(this.btnDeleteRow);
        this.Controls.Add(this.btnSave);
        this.Controls.Add(this.btnCancel);
        this.MinimumSize = new Size(600, 400);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.Text = "自定义属性 - ";

        this.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
    }

    /// <summary>
    /// 添加行按钮点击事件
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void btnAddRow_Click(object sender, EventArgs e)
    {
        DataRow row = _dataTable.NewRow();
        row["序号"] = _dataTable.Rows.Count + 1;
        row["部件ID"] = $"id{_dataTable.Rows.Count + 1:D4}";
        _dataTable.Rows.Add(row);

        // 滚动到最后一行
        dataGridView.FirstDisplayedScrollingRowIndex = dataGridView.RowCount - 1;
    }

    /// <summary>
    /// 删除行按钮点击事件
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void btnDeleteRow_Click(object sender, EventArgs e)
    {
        if (dataGridView.SelectedRows.Count == 0)
        {
            MessageBox.Show("请先选择要删除的行！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (MessageBox.Show("确定要删除选中的行吗？", "确认删除",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
        {
            return;
        }

        // 反向遍历选中行，避免集合修改异常
        var selectedIndices = dataGridView.SelectedRows.Cast<DataGridViewRow>()
            .Select(row => row.Index)
            .OrderByDescending(i => i)
            .ToList();

        foreach (int index in selectedIndices)
        {
            if (index >= 0 && index < _dataTable.Rows.Count)
            {
                _dataTable.Rows.RemoveAt(index);
            }
        }

        // 重新计算序号
        for (int i = 0; i < _dataTable.Rows.Count; i++)
        {
            _dataTable.Rows[i]["序号"] = i + 1;
            _dataTable.Rows[i]["部件ID"] = $"id{i + 1:D4}";
        }
    }

    /// <summary>
    /// 保存按钮点击事件
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void btnSave_Click(object sender, EventArgs e)
    {
        dataGridView.EndEdit(); // 结束编辑，保存修改

        // 验证数据
        foreach (DataRow row in _dataTable.Rows)
        {
            if (string.IsNullOrWhiteSpace(row["部件名"]?.ToString()))
            {
                MessageBox.Show("部件名不能为空！", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
    }

}

