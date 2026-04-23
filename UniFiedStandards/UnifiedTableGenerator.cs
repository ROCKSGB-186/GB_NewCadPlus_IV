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
        /// <summary>
        /// 设备名称
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 设备类型（如阀门、法兰等）
        /// </summary>
        public string Type { get; set; } = string.Empty;

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
    public partial class UnifiedTableGenerator
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
        /// “生成(工艺)管道表”：GeneratePipeTableFromSelection GeneratedeviceTable
        /// - 交互式选择实体
        /// - 提取属性（优先使用属性中包含“长度”的字段作为长度）
        /// - 按属性组合或几何尺寸分组并累加长度
        /// - 最终以表格形式（沿用本类表格样式）生成一张或多张“管道”表（如果选择中包含多个不同管道规格，会生成一张表列出所有分组）
        /// 说明：此方法同时可被命令调用或者由 UI 层调用（无需 UI 层再自行实现分组/样式）
        /// </summary>
        [CommandMethod("GeneratePipeTableFromSelection")]
        public void GeneratePipeTableFromSelection()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            AutoCadHelper.ExecuteInDocumentTransaction((d, tr) =>
            {
                var ed = d.Editor;
                try
                {
                    ed.WriteMessage("\n开始生成管道表：请选择要统计的管道图元（完成选择回车）.");
                    var pso = new PromptSelectionOptions { MessageForAdding = "\n请选择要统计的管道图元：" };
                    var psr = ed.GetSelection(pso);
                    if (psr.Status != PromptStatus.OK || psr.Value == null)
                    {
                        ed.WriteMessage("\n未选择实体或已取消。");
                        return;
                    }

                    var selIds = psr.Value.GetObjectIds();
                    if (selIds == null || selIds.Length == 0)
                    {
                        ed.WriteMessage("\n未选择任何实体。");
                        return;
                    }

                    // 基于所选图元推断比例分母（优先）
                    double scaleDenom = AutoCadHelper.GetScaleDenominatorForSelection(d.Database, selIds, roundToCommon: false);
                    ed.WriteMessage($"\n已推断到比例分母: {scaleDenom}（将用于表格文字高度与行高计算）。");

                    const double unitToMeters = 1000.0;
                    string[] pipeNoKeys = new[] { "管段号", "管段编号", "Pipeline No", "Pipeline", "Pipe No", "No" };
                    string[] startKeys = new[] { "起点", "始点", "From" };
                    string[] endKeys = new[] { "终点", "止点", "To" };

                    string FindFirstAttrValueLocal(Dictionary<string, string>? attrs, string[] candidates)
                    {
                        if (attrs == null) return string.Empty;
                        foreach (var c in candidates)
                        {
                            if (attrs.TryGetValue(c, out var v) && !string.IsNullOrWhiteSpace(v)) return v;
                        }
                        foreach (var kv in attrs)
                        {
                            if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                            foreach (var c in candidates)
                            {
                                if (kv.Key.IndexOf(c, StringComparison.OrdinalIgnoreCase) >= 0 && !string.IsNullOrWhiteSpace(kv.Value))
                                    return kv.Value;
                            }
                        }
                        return string.Empty;
                    }

                    var perPipeList = new List<DeviceInfo>();
                    int seqIndex = 0;

                    foreach (var id in selIds)
                    {
                        seqIndex++;
                        try
                        {
                            var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (ent == null) continue;

                            var attrMap = GetEntityAttributeMap(tr, ent);

                            double length_m = double.NaN;
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

                            string startStr = FindFirstAttrValueLocal(attrMap, startKeys);
                            string endStr = FindFirstAttrValueLocal(attrMap, endKeys);

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

                            string pipeNo = FindFirstAttrValueLocal(attrMap, pipeNoKeys);
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

                            var info = new DeviceInfo
                            {
                                Name = $"PIPE_{seqIndex}",
                                Type = "管道"
                            };

                            if (attrMap != null)
                            {
                                foreach (var kv in attrMap) info.Attributes[kv.Key] = kv.Value;
                            }

                            if (!string.IsNullOrWhiteSpace(pipeNo)) info.Attributes["管段号"] = pipeNo;
                            info.Attributes["起点"] = startStr;
                            info.Attributes["终点"] = endStr;
                            info.Attributes["长度(m)"] = length_m.ToString("F3");
                            info.Attributes["累计长度(m)"] = length_m.ToString("F3");
                            if (info.Attributes.TryGetValue("介质", out var medVal) && !info.Attributes.ContainsKey("介质名称"))
                            {
                                info.Attributes["介质名称"] = medVal;
                            }
                            perPipeList.Add(info);
                        }
                        catch (System.Exception exEnt)
                        {
                            ed.WriteMessage($"\n处理实体 {id} 时出错: {exEnt.Message}");
                        }
                    }

                    if (perPipeList.Count == 0)
                    {
                        ed.WriteMessage("\n未生成任何管道记录。");
                        return;
                    }

                    int ExtractPipeNoNumber(string s)
                    {
                        if (string.IsNullOrWhiteSpace(s)) return int.MaxValue;
                        var m = Regex.Match(s, @"\d+");
                        if (m.Success && int.TryParse(m.Value, out var v)) return v;
                        return int.MaxValue;
                    }

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

                    CreateDeviceTableWithType(d.Database, finalList, "管道明细", scaleDenom);
                    ed.WriteMessage($"\n管道表已生成，共 {finalList.Count} 条记录（每选中一条管道一行）。");
                }
                catch (System.Exception ex)
                {
                    d.Editor.WriteMessage($"\n生成管道表时发生错误: {ex.Message}");
                }
            });
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
        //private Dictionary<string, string> GetEntityAttributeMap(Transaction tr, Entity ent)
        //{
        //    var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //    try
        //    {
        //        if (ent == null) return map;

        //        // 1) AttributeReference（块参照）
        //        if (ent is BlockReference br)
        //        {
        //            try
        //            {
        //                var attCol = br.AttributeCollection;
        //                foreach (ObjectId attId in attCol)
        //                {
        //                    try
        //                    {
        //                        var ar = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
        //                        if (ar != null)
        //                        {
        //                            var tag = (ar.Tag ?? string.Empty).Trim();
        //                            var val = (ar.TextString ?? string.Empty).Trim();
        //                            if (!string.IsNullOrEmpty(tag) && !map.ContainsKey(tag)) map[tag] = val;
        //                        }
        //                    }
        //                    catch { /* 忽略单个属性读取失败 */ }
        //                }
        //            }
        //            catch { /* 忽略 */ }
        //        }

        //        // 2) ExtensionDictionary 的 Xrecord
        //        try
        //        {
        //            if (ent.ExtensionDictionary != ObjectId.Null)
        //            {
        //                var extDict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
        //                if (extDict != null)
        //                {
        //                    foreach (var entry in extDict)
        //                    {
        //                        try
        //                        {
        //                            var xrec = tr.GetObject(entry.Value, OpenMode.ForRead) as Xrecord;
        //                            if (xrec != null && xrec.Data != null)
        //                            {
        //                                var vals = xrec.Data.Cast<TypedValue>().Select(tv => tv.Value?.ToString() ?? "").ToArray();
        //                                var key = entry.Key ?? string.Empty;
        //                                var value = string.Join("|", vals);
        //                                if (!map.ContainsKey(key)) map[key] = value;
        //                            }
        //                        }
        //                        catch { }
        //                    }
        //                }
        //            }
        //        }
        //        catch { /* 忽略 */ }

        //        // 3) RegApp XData
        //        try
        //        {
        //            var db = ent.Database;
        //            var rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
        //            foreach (ObjectId appId in rat)
        //            {
        //                try
        //                {
        //                    var app = tr.GetObject(appId, OpenMode.ForRead) as RegAppTableRecord;
        //                    if (app == null) continue;
        //                    var appName = app.Name;
        //                    var rb = ent.GetXDataForApplication(appName);
        //                    if (rb != null)
        //                    {
        //                        var vals = rb.Cast<TypedValue>().Select(tv => tv.Value?.ToString() ?? "").ToArray();
        //                        var key = $"XDATA:{appName}";
        //                        var value = string.Join("|", vals);
        //                        if (!map.ContainsKey(key)) map[key] = value;
        //                    }
        //                }
        //                catch { }
        //            }
        //        }
        //        catch { /* 忽略 */ }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nGetEntityAttributeMap 异常: {ex.Message}");
        //    }
        //    return map;
        //}

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
            /// 模板箭头填充颜色（优先读取模板中的 Solid/Hatch 颜色）
            /// </summary>
            public Autodesk.AutoCAD.Colors.Color? DirectionArrowFillColor { get; set; }
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
                    var newBrId = InsertPipeBlockWithAttributes(tr, midPoint, newBlockName, 0.0, attValues);
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


        //public void SyncPipeProperties()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    if (doc == null) return;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;

        //    // 选择线段
        //    var lineSelResult = ed.GetSelection(
        //        new PromptSelectionOptions { MessageForAdding = "\n请选择要同步的线段 (LINE 或 LWPOLYLINE):" },
        //        new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "LINE,LWPOLYLINE") })
        //    );
        //    if (lineSelResult.Status != PromptStatus.OK)
        //    {
        //        ed.WriteMessage("\n操作取消。");
        //        return;
        //    }
        //    var sourceLineIds = lineSelResult.Value.GetObjectIds().ToList();

        //    // 选择示例管线块（作为样例）
        //    var blockSelResult = ed.GetEntity("\n请选择示例管线块:");
        //    if (blockSelResult.Status != PromptStatus.OK)
        //    {
        //        ed.WriteMessage("\n操作取消。");
        //        return;
        //    }

        //    using (var tr = new DBTrans())
        //    {
        //        try
        //        {
        //            // 读取示例块参照
        //            var sampleBlockRef = tr.GetObject(blockSelResult.ObjectId, OpenMode.ForRead) as BlockReference;
        //            if (sampleBlockRef == null)
        //            {
        //                ed.WriteMessage("\n错误：选择的不是块参照。");
        //                return;
        //            }

        //            // 解析示例块（提取 polyline / arrow / attribute definitions）
        //            var sampleInfo = AnalyzeSampleBlock(tr, sampleBlockRef);
        //            // 检查模板
        //            if (sampleInfo?.PipeBodyTemplate == null)
        //            {
        //                ed.WriteMessage("\n错误：示例块中未找到作为管道主体的 Polyline。");
        //                return;
        //            }

        //            // 收集并构建顶点顺序
        //            var lineSegments = CollectLineSegments(tr, sourceLineIds);
        //            if (lineSegments == null || lineSegments.Count == 0)
        //            {
        //                ed.WriteMessage("\n未找到可处理的线段。");
        //                return;
        //            }
        //            // 构建顶点顺序
        //            var orderedVertices = BuildOrderedVerticesFromSegments(lineSegments, 0.1);
        //            if (orderedVertices == null || orderedVertices.Count < 2)
        //            {
        //                ed.WriteMessage("\n顶点不足，无法生成管线。");
        //                return;
        //            }

        //            // 读取示例块的所有属性
        //            var sampleAttrMap = GetEntityAttributeMap(tr, sampleBlockRef) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //            // 加载上次保存的属性（此处取决于 sampleBlockRef 名称中是否包含入口/出口关键词）
        //            bool sampleIsOutlet = (sampleBlockRef.Name ?? string.Empty).ToLowerInvariant().Contains("出口") ||
        //                                  (sampleBlockRef.Name ?? string.Empty).ToLowerInvariant().Contains("outlet");

        //            // 从磁盘读取历史属性
        //            var savedAttrsSync = FileManager.LoadLastPipeAttributes(sampleIsOutlet);

        //            // 关键修复：编辑窗口初始数据应包含“示例块当前属性 + 历史属性”
        //            // 1) 先放示例块属性，保证首次打开也能看到完整字段
        //            var initialAttrsForEditor = new Dictionary<string, string>(sampleAttrMap, StringComparer.OrdinalIgnoreCase);
        //            // 2) 再覆盖历史属性，保证用户上次编辑过的值优先显示
        //            foreach (var kv in savedAttrsSync)
        //            {
        //                initialAttrsForEditor[kv.Key] = kv.Value;
        //            }

        //            // 打开属性编辑窗，传入合并后的初始字典
        //            using (var editor = new PipeAttributeEditorForm(initialAttrsForEditor))
        //            {
        //                var dr = editor.ShowDialog();
        //                if (dr != DialogResult.OK)
        //                {
        //                    ed.WriteMessage("\n已取消属性编辑，停止同步操作。");
        //                    return;
        //                }

        //                // 保存历史属性供下次使用
        //                var editedAttrs = editor.Attributes ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //                FileManager.SaveLastPipeAttributes(sampleIsOutlet, editedAttrs);

        //                // 把用户修改后的属性写回示例块（只写存在的 AttributeReference）
        //                try
        //                {
        //                    var sampleBrWrite = tr.GetObject(sampleBlockRef.ObjectId, OpenMode.ForWrite) as BlockReference;
        //                    if (sampleBrWrite != null)
        //                    {
        //                        // 遍历属性
        //                        foreach (ObjectId aid in sampleBrWrite.AttributeCollection)
        //                        {
        //                            try
        //                            {
        //                                // 获取属性引用
        //                                var ar = tr.GetObject(aid, OpenMode.ForWrite) as AttributeReference;
        //                                // 跳过无效的属性
        //                                if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue;
        //                                // 检查属性是否在编辑字典中
        //                                if (editedAttrs.TryGetValue(ar.Tag, out var newVal))
        //                                {
        //                                    // 更新属性
        //                                    ar.TextString = newVal ?? string.Empty;
        //                                    // 对齐
        //                                    try { ar.AdjustAlignment(db); } catch { }
        //                                }
        //                            }
        //                            catch { /* 单个属性写回失败不阻塞整体 */ }
        //                        }
        //                    }
        //                }
        //                catch (System.Exception exWriteSample)
        //                {
        //                    ed.WriteMessage($"\n写回示例块属性时出错: {exWriteSample.Message}");
        //                }
        //            }
        //            // 开始构建新管道块
        //            double pipelineLength = 0.0;
        //            for (int i = 0; i < orderedVertices.Count - 1; i++)
        //                pipelineLength += orderedVertices[i].DistanceTo(orderedVertices[i + 1]);
        //            // 计算中点
        //            var (midPoint, midAngle) = ComputeMidPointAndAngle(orderedVertices, pipelineLength);
        //            // 计算目标向量
        //            Vector3d targetDir = ComputeDirectionAtPoint(orderedVertices, midPoint, 1e-6);
        //            // 计算聚合线段向量
        //            Vector3d segmentDir = ComputeAggregateSegmentDirection(lineSegments);
        //            // 如果聚合线段向量与目标向量方向相反，则反转目标向量
        //            if (!segmentDir.IsZeroLength() && targetDir.DotProduct(segmentDir) < 0)
        //                targetDir = -targetDir;
        //            // 归一化目标向量  
        //            Vector3d targetDirNormalized = targetDir.IsZeroLength() ? Vector3d.XAxis : targetDir.GetNormal();
        //            // 构建管道
        //            Polyline pipeLocal = BuildPipePolylineLocal(sampleInfo.PipeBodyTemplate, orderedVertices, midPoint);

        //            // 复制属性定义（基于示例块的定义）——先克隆示例定义（保持字段顺序与名称）
        //            var attDefsLocal = CloneAttributeDefinitionsLocal(sampleInfo.AttributeDefinitions, midPoint, 0.0, pipelineLength, sampleBlockRef.Name)
        //                                ?? new List<AttributeDefinition>();

        //            // 重新读取示例块属性（刚刚可能已被编辑并写回）
        //            var latestSampleAttrs = GetEntityAttributeMap(tr, sampleBlockRef) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        //            // 生成标题（优先属性中的管道标题）
        //            string pipeTitle = latestSampleAttrs.TryGetValue("管道标题", out var sampleTitle) && !string.IsNullOrWhiteSpace(sampleTitle)
        //                               ? sampleTitle
        //                               : sampleBlockRef.Name ?? "管道";

        //            // 为每一段生成局部坐标下的箭头与标题（相对于 midPoint），仅使用分段箭头/文字
        //            var arrowEntities = CreateDirectionalArrowsAndTitles(tr, sampleInfo, orderedVertices, midPoint, pipeTitle, sampleBlockRef.Name);

        //            // ---------- 关键修正：确保新块的属性定义字段严格以示例块的 AttributeDefinitions 为准 ----------
        //            // 1) 若示例存在属性定义，则按照示例顺序保留字段，仅更新 TextString（不新增示例中不存在的字段）
        //            // 2) 若示例没有属性定义，则以 latestSampleAttrs 为基础创建属性定义（按键排序以保证稳定性）
        //            if (sampleInfo.AttributeDefinitions != null && sampleInfo.AttributeDefinitions.Count > 0)
        //            {
        //                // 使用示例定义作为基准，更新文本值（保留原有位置/高度/顺序）
        //                var latestDict = new Dictionary<string, string>(latestSampleAttrs, StringComparer.OrdinalIgnoreCase);
        //                foreach (var def in attDefsLocal)
        //                {
        //                    if (string.IsNullOrWhiteSpace(def.Tag)) continue;
        //                    if (latestDict.TryGetValue(def.Tag, out var val))
        //                    {
        //                        def.TextString = val ?? string.Empty;
        //                    }
        //                    // 临时显示设置（随后统一隐藏/显示处理）
        //                    def.Invisible = false;
        //                    def.Constant = false;
        //                }
        //            }
        //            else
        //            {
        //                // 示例无属性定义：根据 latestSampleAttrs 动态创建属性定义（按 Key 排序）
        //                attDefsLocal.Clear();
        //                double attHeight = 3.5;
        //                double yOffsetBase = -attHeight * 2.0;
        //                int idx = 0;
        //                foreach (var kv in latestSampleAttrs.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
        //                {
        //                    if (string.IsNullOrWhiteSpace(kv.Key)) continue;
        //                    attDefsLocal.Add(new AttributeDefinition
        //                    {
        //                        Tag = kv.Key,
        //                        Position = new Point3d(0, yOffsetBase - idx * attHeight * 1.2, 0),
        //                        Rotation = 0.0,
        //                        TextString = kv.Value ?? string.Empty,
        //                        Height = attHeight,
        //                        Invisible = false,
        //                        Constant = false
        //                    });
        //                    idx++;
        //                }
        //            }

        //            // 新增或覆盖 起点/终点 属性定义（保持原逻辑，若示例中已有这些字段则更新其值，否则新增）
        //            Point3d worldStart = orderedVertices.First();
        //            Point3d worldEnd = orderedVertices.Last();
        //            string startCoordStr = $"X={worldStart.X:F3},Y={worldStart.Y:F3}";
        //            string endCoordStr = $"X={worldEnd.X:F3},Y={worldEnd.Y:F3}";
        //            int nextSegNum = GetNextPipeSegmentNumber(db);

        //            // 取管段号，优先从属性或标题提取
        //            string extractedPipeNo = string.Empty;
        //            if (latestSampleAttrs.TryGetValue("管道标题", out var titleFromSample) && !string.IsNullOrWhiteSpace(titleFromSample))
        //            {
        //                extractedPipeNo = ExtractPipeCodeFromTitle(titleFromSample);
        //            }
        //            if (string.IsNullOrWhiteSpace(extractedPipeNo))
        //            {
        //                extractedPipeNo = ExtractPipeCodeFromTitle(sampleBlockRef.Name);
        //            }
        //            if (string.IsNullOrWhiteSpace(extractedPipeNo))
        //            {
        //                if (latestSampleAttrs.TryGetValue("管段号", out var pn) && !string.IsNullOrWhiteSpace(pn))
        //                    extractedPipeNo = pn;
        //                else if (latestSampleAttrs.TryGetValue("管段编号", out var pn2) && !string.IsNullOrWhiteSpace(pn2))
        //                    extractedPipeNo = pn2;
        //            }
        //            if (string.IsNullOrWhiteSpace(extractedPipeNo))
        //            {
        //                extractedPipeNo = nextSegNum.ToString("D4");
        //            }

        //            // 局部函数：按示例定义优先更新 / 新增（仅当示例中不存在该字段时新增）管道号
        //            void SetOrAddAttrLocal(string tag, string text)
        //            {
        //                // 按示例定义优先更新 查找是否已存在该属性定义
        //                var existing = attDefsLocal.FirstOrDefault(a => string.Equals(a.Tag, tag, StringComparison.OrdinalIgnoreCase));
        //                if (existing != null)
        //                {
        //                    existing.TextString = text;// 更新属性值 按示例定义优先更新
        //                    existing.Invisible = false;// 临时显示设置（随后统一隐藏/显示处理）
        //                    existing.Constant = false;// 临时常量设置（随后统一取消常量处理）
        //                }
        //                else
        //                {
        //                    // 只有在示例没有任何定义时允许新增；但为兼容性仍允许新增起/终/段号
        //                    attDefsLocal.Add(new AttributeDefinition
        //                    {
        //                        Tag = tag,
        //                        Position = new Point3d(0, (attDefsLocal.Count > 0 ? attDefsLocal[0].Position.Y - attDefsLocal[0].Height * 1.2 : -3.5), 0),
        //                        Rotation = 0.0,
        //                        TextString = text,
        //                        Height = attDefsLocal.Count > 0 ? attDefsLocal[0].Height : 3.5,
        //                        Invisible = false,
        //                        Constant = false
        //                    });
        //                }
        //            }

        //            SetOrAddAttrLocal("始点", startCoordStr);
        //            SetOrAddAttrLocal("终点", endCoordStr);
        //            SetOrAddAttrLocal("管段号", extractedPipeNo);

        //            //// 移除块定义中的中点“管道标题”属性（避免在块中重复显示中点标题）
        //            //attDefsLocal.RemoveAll(ad => string.Equals(ad.Tag, "管道标题", StringComparison.OrdinalIgnoreCase));

        //            //// 将其余属性设置为隐藏（块内不显示），以保持行为与现有逻辑一致
        //            foreach (var ad in attDefsLocal)
        //            {
        //                ad.Invisible = true;
        //                ad.Constant = false;
        //            }

        //            // 构建块定义并插入新块
        //            string desiredName = sampleBlockRef.Name;
        //            // 创建块定义 若名称已存在则添加后缀
        //            string newBlockName = BuildPipeBlockDefinition(tr, desiredName, (Polyline)pipeLocal.Clone(), arrowEntities, attDefsLocal);
        //            // 准备属性值字典
        //            var attValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //            foreach (var a in attDefsLocal)
        //            {
        //                if (string.IsNullOrWhiteSpace(a?.Tag)) continue;// 跳过无效标签 准备属性值字典
        //                attValues[a.Tag] = a.TextString ?? string.Empty;//  收集属性值
        //            }
        //            // 插入新块并设置属性
        //            var newBrId = InsertPipeBlockWithAttributes(tr, midPoint, newBlockName, 0.0, attValues);
        //            var newBr = tr.GetObject(newBrId, OpenMode.ForWrite) as BlockReference;// 插入新块并设置属性
        //            if (newBr != null)
        //                newBr.Layer = sampleInfo.PipeBodyTemplate.Layer;// 插入新块并设置属性 继承图层

        //            // 删除原始线段
        //            foreach (var seg in lineSegments)
        //            {
        //                var ent = tr.GetObject(seg.Id, OpenMode.ForWrite) as Entity;
        //                if (ent != null) ent.Erase();
        //            }

        //            tr.Commit();
        //            ed.WriteMessage($"\n管线块已生成：新增/更新属性 [始点][终点][管段号]={extractedPipeNo}。");
        //        }
        //        catch (Exception ex)
        //        {
        //            ed.WriteMessage($"\n发生错误: {ex.Message}\n{ex.StackTrace}");
        //            tr.Abort();
        //        }
        //    }

        //    try
        //    {
        //        ed.Regen();
        //        Application.UpdateScreen();
        //    }
        //    catch { }
        //}

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

                // 设置网格线粗细
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
        /// 修复：设置表格样式 - 修正版
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
        /// 设置单元格样式（只设置对齐与必要的默认字体样式，避免修改全局 TextStyle）
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
        private void AutoResizeColumns(Table table)
        {
            try
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    double maxWidth = 10.0; // 最小宽度（绘图单位）
                    double sampleTextHeight = 0.0;

                    // 1) 估算该列使用的平均文字高度（取第一非零的 TextHeight）
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        try
                        {
                            var th = table.Cells[i, j].TextHeight;
                            if (th > 0.0)
                            {
                                sampleTextHeight = Convert.ToDouble(th);
                                break;
                            }
                        }
                        catch { }
                    }
                    // 若未找到有效 TextHeight，尝试行内查找（兜底）
                    if (sampleTextHeight <= 0.0)
                    {
                        // 取全表第一个非零 TextHeight
                        for (int ii = 0; ii < table.Rows.Count && sampleTextHeight <= 0.0; ii++)
                        {
                            for (int jj = 0; jj < table.Columns.Count && sampleTextHeight <= 0.0; jj++)
                            {
                                try { sampleTextHeight = Convert.ToDouble(table.Cells[ii, jj].TextHeight); }
                                catch { sampleTextHeight = 0.0; }
                            }
                        }
                    }

                    // 保守兜底：若仍然无有效 TextHeight，则使用一个合理默认
                    if (sampleTextHeight <= 0.0)
                        sampleTextHeight = TextFontsStyleHelper.ComputeScaledHeight(2.0, 1.0);

                    // 2) 计算每行文本的估算字符宽度（以字符单位计），并据此求出最大需求
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        string text = string.Empty;
                        try { text = table.Cells[i, j].TextString ?? string.Empty; } catch { text = string.Empty; }
                        if (!string.IsNullOrEmpty(text))
                        {
                            double charUnits = EstimateTextWidth(text); // 近似单位
                            // 根据文字高度估算所需宽度（经验系数）
                            double charWidthFactor = 0.55; // 每字符宽度相对 TextHeight 的经验系数，可微调
                            double estWidth = Math.Ceiling(charUnits * sampleTextHeight * charWidthFactor);
                            if (estWidth > maxWidth) maxWidth = estWidth;
                        }
                    }

                    // 3) 限制列宽范围，避免过窄或过宽
                    double minColWidth = Math.Max(8.0, sampleTextHeight * 2.5);
                    double maxColWidth = Math.Max(60.0, sampleTextHeight * 40.0);

                    double finalWidth = Math.Max(minColWidth, Math.Min(maxColWidth, maxWidth));

                    try { table.SetColumnWidth(j, finalWidth); } catch { /* 忽略设置失败 */ }
                }
            }
            catch
            {
                // 兜底：保持现状，不抛出异常以免中断调用方流程
            }
        }

        /// <summary>
        /// 设置表格边框并按比例选择合适的线宽（尝试性，API 不同版本行为不同，用 try/catch 包裹）
        /// </summary>
        private void SetTableBorders(Table table, double scaleDenominator)
        {
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
                // 基础的行高设置
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    table.SetRowHeight(i, i == 0 ? 12.0 : 8.0);
                }

                // 基础的列宽设置
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

                // 管段号：第1列（索引1），跨2行
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
        /// <param name="db">数据库对象</param>
        /// <param name="roundToCommon">是否四舍五入到常见比例</param>
        /// <returns>绘图比例分母</returns>
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
                                    if (mergedMark[r + maxV, cc]) { rowOk = false; break; }
                                    var downTxt = (table.Cells[r + maxV, cc].TextString ?? string.Empty).Trim();
                                    if (!string.IsNullOrEmpty(downTxt)) { rowOk = false; break; }
                                }
                                if (!rowOk) break;
                                maxV++;
                            }

                            // 规则变更：
                            // - 保持第1-3行（0-based 0..2）原有合并逻辑
                            // - 第4行及以后（r >= 3）禁止任何方向的合并（既禁止横向也禁止纵向）
                            // - 若起始行在第3行及以前，但合并会跨过第3行边界，则限制垂直合并使其不会跨入第4行（即 r+maxV-1 <= 2）
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
                    // 这里移除对未定义变量 excelRows/excelCols/excelData 的引用，避免编译错误。

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
        #endregion
        #region 同步表格
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
                                                if (idKeywords.Any(k => tag.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0) ||
                                                    tag.Equals("部件ID", StringComparison.OrdinalIgnoreCase) ||
                                                    tag.Equals("部件 Id", StringComparison.OrdinalIgnoreCase))
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
                                catch { /* 忽略单个实体读取问题 */ }
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
                                                                conv = Convert.ChangeType(cleanedValue, targetType, System.Globalization.CultureInfo.InvariantCulture);
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
                doc.Editor.WriteMessage($"\n导入失败: {ex.Message}");
            }
        }
        #endregion
    }
}

