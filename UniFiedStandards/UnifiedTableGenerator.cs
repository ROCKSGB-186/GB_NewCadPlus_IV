using Autodesk.AutoCAD.DatabaseServices;
using GB_NewCadPlus_IV.DisplayPages;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Helpers;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;          // 用于 CellRangeAddress 等辅助类
using NPOI.XSSF.UserModel;  // 仅用于创建 .xlsx 工作簿
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;             // FileStream 必需
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using GB_NewCadPlus_IV.DisplayPages;
using static Autodesk.AutoCAD.Features.PointCloud.PointCloudColorMapping.ClassificationRamp;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using AttributeCollection = Autodesk.AutoCAD.DatabaseServices.AttributeCollection;
using BorderStyle = NPOI.SS.UserModel.BorderStyle;
using Database = Autodesk.AutoCAD.DatabaseServices.Database;
using DataTable = System.Data.DataTable;
using DoubleCollection = Autodesk.AutoCAD.Geometry.DoubleCollection;
using DrawOrderTable = Autodesk.AutoCAD.DatabaseServices.DrawOrderTable;
using HorizontalAlignment = NPOI.SS.UserModel.HorizontalAlignment;
using MessageBox = System.Windows.Forms.MessageBox;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Table = Autodesk.AutoCAD.DatabaseServices.Table;



/// 设备属性块信息类和统一表生成器类
namespace GB_NewCadPlus_IV.UniFiedStandards
{
    /// <summary>
    /// 设备属性块信息类（用于存储CAD块中的设备信息）
    /// </summary>
    public class DeviceInfo
    {
        /// <summary>
        /// 设备ID
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// 设备名
        /// </summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>
        /// 设备类型
        /// </summary>
        public string Type { get; set; } = string.Empty;
        /// <summary>
        /// 介质
        /// </summary>
        public string MediumName { get; set; } = string.Empty;
        /// <summary>
        /// 规格
        /// </summary>
        public string Specifications { get; set; } = string.Empty;
        /// <summary>
        /// 材质
        /// </summary>
        public string Material { get; set; } = string.Empty;
        /// <summary>
        /// 数量
        /// </summary>
        public int Quantity { get; set; }
        /// <summary>
        /// 图号
        /// </summary>
        public string DrawingNumber { get; set; } = string.Empty;
        /// <summary>
        /// 功率
        /// </summary>
        public decimal Power { get; set; }
        /// <summary>
        /// 容积
        /// </summary>
        public decimal Volume { get; set; }
        /// <summary>
        /// 压力
        /// </summary>
        public decimal Pressure { get; set; }
        /// <summary>
        /// 温度
        /// </summary>
        public decimal Temperature { get; set; }
        /// <summary>
        /// 直径
        /// </summary>
        public decimal Diameter { get; set; }
        /// <summary>
        /// 长度
        /// </summary>
        public decimal Length { get; set; }
        /// <summary>
        /// 厚度
        /// </summary>
        public decimal Thickness { get; set; }
        /// <summary>
        /// 重量
        /// </summary>
        public decimal Weight { get; set; }
        /// <summary>
        /// 型号
        /// </summary>
        public string Model { get; set; } = string.Empty;
        /// <summary>
        /// 标准号
        /// </summary>
        public string DeviceSTDNo { get; set; } = string.Empty;
        /// <summary>
        /// 备注
        /// </summary>
        public string Remarks { get; set; } = string.Empty;


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


        #region 新设置表生成方法

        /// <summary>
        /// 主命令：生成设备表（增强版：确保只使用选中图元数据，并智能合并相同设备）
        /// </summary>
        [CommandMethod(nameof(GenerateDeviceTable))]
        public void GenerateDeviceTable()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            try
            {
                LogManager.Instance.LogInfo("[设备表][命令开始] GenerateDeviceTable");
                // 1. 使用原有的选择与分析逻辑（不变）
                var (devices, selIds) = DynamicBlockOperations.SelectAndAnalyzeBlocks(ed, doc.Database);
                LogManager.Instance.LogInfo($"[设备表][选择返回] SelectedObjectCount={selIds?.Length ?? 0}, DeviceCount={devices?.Count ?? 0}");
                if (devices == null || devices.Count == 0)
                {
                    LogManager.Instance.LogWarning("[设备表][命令结束] 未获得可用设备图元。");
                    ed.WriteMessage("\n未找到可用的设备信息。");
                    return;
                }

                // 2. ★ 新增：智能合并重复设备（完全相同才合并，否则独立）
                var mergedDevices = AggregateSelectedDeviceQuantities(devices);
                LogManager.Instance.LogInfo($"[设备表][NAME聚合完成] SourceCount={devices.Count}, AggregatedCount={mergedDevices.Count}");
                ed.WriteMessage($"\n[设备数量统计] 原始图元 {devices.Count} 个，合并后 {mergedDevices.Count} 行。" );

                // 3. ★ 新增：按名称排序，使表格看上去整齐，并避免合并逻辑误判
                var sortedDevices = mergedDevices
                    .OrderBy(d => d.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                // 4. 原有比例计算（不变）
                double scaleDenom = AutoCadHelper.GetScale();

                // 5. 按 Type 分组，为每组生成独立表（原有样式不变）
                var groups = sortedDevices.GroupBy(e => string.IsNullOrWhiteSpace(e.Type) ? "设备" : e.Type);
                // 整个设备表批次只选择一次起始位置，后续每个独立表格自动向下排列。
                PromptPointResult deviceTablePoint = ed.GetPoint("\n指定设备表批量插入起始位置: ");
                if (deviceTablePoint.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\n未指定设备表批量插入位置，已取消设备表生成。");
                    LogManager.Instance.LogWarning($"[设备表][批量插入取消] Status={deviceTablePoint.Status}");
                    return;
                }

                Point3d nextDeviceTablePosition = deviceTablePoint.Value; // 记录下一张设备表的插入位置，初始为用户指定的起始点。
                var deviceTableIds = new List<ObjectId>(); // 保存每张独立设备表的 ObjectId，便于后续统一调整列宽和重新排列表格。
                double deviceTableGap = Math.Max(2.0, scaleDenom * 2.0); // 默认垂直间距为 2 个比例单位，避免表格重叠
                LogManager.Instance.LogInfo(
                    $"[设备表][批量插入开始] StartPoint={deviceTablePoint.Value}, GroupCount={groups.Count()}, VerticalGap={deviceTableGap}");

                // 遍历按设备类型分组后的设备数据，为每一种设备类型分别生成一张独立的 AutoCAD 设备表。
                foreach (var g in groups)
                {
                    // 将当前设备类型分组转换为列表，便于统计行数、记录日志并传递给表格创建方法。
                    List<DeviceInfo> list = g.ToList();

                    // 记录当前设备类型、表格行数以及每个设备的汇总数量，便于核对设备表生成数据。
                    LogManager.Instance.LogInfo(
                        $"[设备表][生成类型表] Type={g.Key}, RowCount={list.Count}, Quantities=[{string.Join(",", list.Select(item => $"{item.Name}:{item.Quantity}"))}]");

                    // 根据当前设备列表、比例分母和插入位置创建一张设备表。
                    // 每个设备类型单独创建一张表，避免不同类型的设备混在同一张表中。
                    DeviceTableCreateResult tableResult = CreateDeviceTable(
                        doc.Database,
                        list,
                        scaleDenom,
                        nextDeviceTablePosition);

                    // 只有在表格创建成功并返回有效 ObjectId 时，才记录表格并计算下一张表的位置。
                    if (tableResult.TableId != ObjectId.Null)
                    {
                        // 保存当前表格的 ObjectId，供后续统一调整列宽和重新排列表格使用。
                        deviceTableIds.Add(tableResult.TableId);

                        // AutoCAD 表格从插入位置向下占用空间。
                        // 当前表格创建完成后，将下一张表放置在当前表格下方，并保留设定的垂直间距。
                        nextDeviceTablePosition = new Point3d(
                            // 下一张表沿用批量插入起始点的 X 坐标，保证所有表格左侧对齐。
                            deviceTablePoint.Value.X,
                            // 根据当前表格高度和表格间距计算下一张表的 Y 坐标。
                            nextDeviceTablePosition.Y - tableResult.Height - deviceTableGap,
                            // 保留原插入点的 Z 坐标。
                            nextDeviceTablePosition.Z);

                        // 记录当前表格的位置、实际高度以及下一张表的预计插入位置，便于排查布局问题。
                        LogManager.Instance.LogInfo(
                            $"[设备表][批量表格位置] Type={g.Key}, TableId={tableResult.TableId}, Position={tableResult.Position}, Height={tableResult.Height}, NextPosition={nextDeviceTablePosition}");
                    }
                    // 在 AutoCAD 命令行中提示当前设备类型表格的生成结果。
                    ed.WriteMessage($"\n已为类型 '{g.Key}' 生成表，包含 {list.Count} 条汇总项（使用比例分母 {scaleDenom}）。");
                }
            
                // 使用固定列宽和固定行高替代按内容自动计算列宽。
                // 该方法会重新打开已经插入图纸的表格，因此表格外框也会同步更新。
                SetDeviceTablesFixedSize(
                    doc.Database,
                    deviceTableIds,
                    scaleDenom);

                // 重新生成当前图形显示，确保图纸立即显示更新后的表格外框。
                try
                {
                    // 更新 AutoCAD 图形数据库的显示内容。
                    ed.Regen();

                    // 刷新 AutoCAD 屏幕。
                    Application.UpdateScreen();
                }
                catch
                {
                    // 显示刷新失败不影响已经提交的表格布局更新。
                }
               
                // 统一列宽并重新生成布局后，表格高度可能发生变化。
                // 因此需要按照每张表的最终高度重新计算垂直排列位置。
                Point3d nextBatchPosition = ReflowDeviceTables(
                    doc.Database,
                    deviceTableIds,
                    deviceTablePoint.Value,
                    deviceTableGap);

                // 记录本批次设备表的最终数量和布局处理完成状态。
                LogManager.Instance.LogInfo(
                    $"[设备表][批量插入完成] TableCount={deviceTableIds.Count}, ReflowStart={deviceTablePoint.Value}, Gap={deviceTableGap}");
 
                // 6. 使用原始选中图元统计螺栓，避免设备表合并影响每个图元的螺栓数量。
                List<BoltStatisticRow> boltRows = BuildBoltStatistics(
                    devices,
                    out int boltCandidateCount,
                    out int skippedBoltComponentCount);
                ed.WriteMessage(
                    $"\n[螺栓统计] 已分析 {devices.Count} 个设备图元，命中 {boltCandidateCount} 个法兰/对夹候选图元，得到 {boltRows.Count} 条有效汇总记录。");
                if (boltRows.Count > 0)
                {
                    // 螺丝统计表作为独立表格自动接续在所有设备表下方，不再要求用户二次指定位置。
                    CreateBoltStatisticsTable(doc, boltRows, scaleDenom, nextBatchPosition);
                    try { ed.Regen(); Application.UpdateScreen(); } catch { }
                    ed.WriteMessage($"\n已生成螺栓统计表，包含 {boltRows.Count} 种螺栓规格。");
                }
                else if (skippedBoltComponentCount > 0)
                {
                    ed.WriteMessage($"\n已识别 {skippedBoltComponentCount} 个法兰/对夹部件，但缺少有效的螺栓规格或螺栓数量，未生成螺栓统计表。");
                }
                else
                {
                    ed.WriteMessage("\n未识别到法兰/对夹、法兰或管端盲板图元，未生成螺栓统计表。");
                }
            }
            catch (System.Exception ex)
            {
                LogManager.Instance.LogError($"[设备表][命令失败] Error={ex.Message}");
                ed.WriteMessage($"\n生成设备表失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 批量设置已经插入图纸的设备表的固定列宽和固定行高。
        /// </summary>
        /// <param name="database">当前图形数据库。</param>
        /// <param name="tableIds">需要设置尺寸的设备表对象 ID 集合。</param>
        /// <param name="scaleDenominator">当前图纸比例分母。</param>
        private void SetDeviceTablesFixedSize(
            Database database,
            IEnumerable<ObjectId> tableIds,
            double scaleDenominator)
        {
            // 检查数据库对象是否有效。
            if (database == null)
            {
                return;
            }

            // 检查表格 ID 集合是否为空。
            if (tableIds == null)
            {
                return;
            }

            // 启动事务，批量修改已经插入图纸的设备表。
            using (Transaction transaction =
                database.TransactionManager.StartTransaction())
            {
                // 遍历本次生成的全部设备表。
                foreach (ObjectId tableId in tableIds)
                {
                    // 跳过无效的表格对象 ID。
                    if (tableId == ObjectId.Null)
                    {
                        continue;
                    }

                    // 以写入模式打开当前表格。
                    Table table =
                        transaction.GetObject(
                            tableId,
                            OpenMode.ForWrite) as Table;

                    // 如果对象不是 AutoCAD 表格，则跳过。
                    if (table == null)
                    {
                        continue;
                    }

                    // 设置当前表格的固定列宽和固定行高。
                    // 方法内部同时会重新生成布局和重建表格内部图形块。
                    SetDeviceTableFixedSize(
                        table,
                        scaleDenominator);

                    // 记录当前表格尺寸已经被重新设置。
                    LogManager.Instance.LogInfo(
                        $"[设备表][固定尺寸] TableId={tableId}, " +
                        $"Columns={table.NumColumns}, Rows={table.NumRows}, " +
                        $"ScaleDenominator={scaleDenominator}");
                }

                // 提交所有表格的列宽、行高和外框几何修改。
                transaction.Commit();
            }
        }

        /// <summary>
        /// 按设备表固定版式设置列宽和行高。
        /// 列顺序必须与设备表实际列顺序保持一致：
        /// 名称、规格、材料、数量、图号或标准号。
        /// </summary>
        /// <param name="table">需要设置尺寸的 AutoCAD 表格。</param>
        /// <param name="scaleDenominator">当前图纸比例分母。</param>
        private void SetDeviceTableFixedSize(
            Table table,
            double scaleDenominator)
        {
            // 检查表格对象是否有效。
            if (table == null)
            {
                return;
            }

            // 防止比例分母为 0 或负数导致列宽、行高无效。
            if (scaleDenominator <= 0)
            {
                scaleDenominator = 1.0;
            }

            // 设置图纸基准列宽。
            // 这些数值按照图纸上的显示尺寸定义，再乘以比例分母转换为模型空间尺寸。
            double nameColumnWidth = 35.0 * scaleDenominator;
            double specificationColumnWidth = 38.0 * scaleDenominator;
            double materialColumnWidth = 25.0 * scaleDenominator;
            double quantityColumnWidth = 15.0 * scaleDenominator;
            double drawingNumberColumnWidth = 35.0 * scaleDenominator;

            // 设置标题行高度。
            // 标题行适当加高，便于显示“图号或标准号”等较长标题。
            double titleRowHeight = 10.0 * scaleDenominator;

            // 设置普通数据行高度。
            double dataRowHeight = 8.0 * scaleDenominator;

            // 当前设备表至少需要包含五列。
            if (table.NumColumns >= 5)
            {
                // 第 0 列：名称。
                table.SetColumnWidth(0, nameColumnWidth);

                // 第 1 列：规格。
                table.SetColumnWidth(1, specificationColumnWidth);

                // 第 2 列：材料。
                table.SetColumnWidth(2, materialColumnWidth);

                // 第 3 列：数量。
                table.SetColumnWidth(3, quantityColumnWidth);

                // 第 4 列：图号或标准号。
                table.SetColumnWidth(4, drawingNumberColumnWidth);
            }

            // 设置标题行高度。
            if (table.NumRows > 0)
            {
                table.SetRowHeight(0, titleRowHeight);
            }

            // 设置所有数据行高度。
            for (int rowIndex = 1; rowIndex < table.NumRows; rowIndex++)
            {
                // 统一设置普通数据行高度，避免不同表格的行高不一致。
                table.SetRowHeight(rowIndex, dataRowHeight);
            }

            // 根据新的列宽和行高重新计算表格布局。
            table.GenerateLayout();

            // 强制 AutoCAD 重建表格内部图形块，使外框线同步变化。
            table.RecomputeTableBlock(true);

            // 标记表格图形数据已经发生变化。
            table.RecordGraphicsModified(true);
        }

        /// <summary>
        /// WPF“生成设备表”按钮使用的唯一命令入口。
        /// 使用独立命令名，避免 CAD 中旧插件命令缓存或同名命令冲突。
        /// </summary>
        [CommandMethod("GenerateDeviceTableWithBoltStatistics")]
        public void GenerateDeviceTableWithBoltStatistics()
        {
            Document document = Application.DocumentManager.MdiActiveDocument;
            if (document == null)
            {
                return;
            }

            document.Editor.WriteMessage("\n[设备表新版入口] 已进入带螺栓统计的设备表生成逻辑。\n");
            GenerateDeviceTable();
        }

        /// <summary>
        /// 按属性块 NAME 值聚合本次选择的图元，并将相同图元数量累计到 Quantity。
        /// NAME 是业务上的部件身份；块定义名称只在属性块缺少 NAME 时作为兜底值。
        /// </summary>
        private List<DeviceInfo> AggregateSelectedDeviceQuantities(List<DeviceInfo> devices)
        {
            if (devices == null || devices.Count == 0)
            {
                return new List<DeviceInfo>();
            }

            var aggregated = new Dictionary<string, DeviceInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (DeviceInfo source in devices.Where(device => device != null))
            {
                // 优先使用属性块中的 NAME 判断相同图元，不能使用 AutoCAD 块定义名代替 NAME。
                string componentName = GetFirstAttributeValue(source.Attributes,
                    "NAME", "Name", "部件名称", "组件名称", "名称", "设备名称");
                if (string.IsNullOrWhiteSpace(componentName))
                {
                    componentName = (source.Name ?? string.Empty).Trim();
                }

                string specification = GetFirstAttributeValue(source.Attributes,
                    "规格型号", "规格", "型号", "MODEL", "SPECIFICATION", "SPEC");
                string material = GetFirstAttributeValue(source.Attributes,
                    "材料", "材质", "MATERIAL", "MATERIALS", "MATL");
                string standardNumber = GetFirstAttributeValue(source.Attributes,
                    "图号或标准号", "图号", "标准号", "DRAWINGNO", "STANDARDNO");

                // 相同图元只按 NAME 判断；规格、材质和标准号作为该 NAME 首条记录的显示属性保留。
                string groupingKey = NormalizeDeviceGroupingPart(componentName);

                int sourceQuantity = GetSelectedDeviceQuantity(source);
                LogManager.Instance.LogInfo(
                    $"[设备表][NAME聚合输入] BlockName={source.Name}, NAME={componentName}, GroupKey={groupingKey}, RawQuantity={GetFirstAttributeValue(source.Attributes, "数量", "QTY", "QUANTITY", "Quantity")}, ParsedQuantity={sourceQuantity}");
                if (!aggregated.TryGetValue(groupingKey, out DeviceInfo target))
                {
                    target = CloneDeviceInfo(source);
                    target.Name = componentName;
                    target.Specifications = specification;
                    target.Material = material;
                    target.DrawingNumber = standardNumber;
                    target.Quantity = sourceQuantity;
                    target.Count = sourceQuantity;
                    target.Attributes["NAME"] = componentName;
                    target.Attributes["名称"] = componentName;
                    target.Attributes["规格"] = specification;
                    target.Attributes["材料"] = material;
                    target.Attributes["图号或标准号"] = standardNumber;
                    target.Attributes["数量"] = sourceQuantity.ToString(CultureInfo.InvariantCulture);
                    aggregated[groupingKey] = target;
                    LogManager.Instance.LogInfo($"[设备表][NAME聚合新行] NAME={componentName}, Quantity={target.Quantity}");
                }
                else
                {
                    target.Quantity += sourceQuantity;
                    target.Count = target.Quantity;
                    target.Attributes["数量"] = target.Quantity.ToString(CultureInfo.InvariantCulture);
                    LogManager.Instance.LogInfo($"[设备表][NAME聚合累计] NAME={componentName}, AddedQuantity={sourceQuantity}, TotalQuantity={target.Quantity}");
                }
            }

            List<DeviceInfo> result = aggregated.Values.ToList();
            for (int index = 0; index < result.Count; index++)
            {
                result[index].Id = index + 1;
                LogManager.Instance.LogInfo($"[设备表][NAME聚合结果] Id={result[index].Id}, NAME={result[index].Name}, Quantity={result[index].Quantity}, Count={result[index].Count}");
            }

            return result;
        }

        /// <summary>
        /// 读取单个选中图元的数量；没有数量属性时，一个图元按一件计算。
        /// </summary>
        private static int GetSelectedDeviceQuantity(DeviceInfo device)
        {
            string rawQuantity = string.Empty;
            if (device?.Attributes != null)
            {
                foreach (string key in new[] { "数量", "QTY", "QUANTITY", "Quantity" })
                {
                    if (device.Attributes.TryGetValue(key, out string value) && !string.IsNullOrWhiteSpace(value))
                    {
                        rawQuantity = value.Trim();
                        break;
                    }
                }
            }

            int quantity = ParsePositiveInteger(rawQuantity);
            if (quantity > 0)
            {
                return quantity;
            }

            return device != null && device.Quantity > 0 ? device.Quantity : 1;
        }

        /// <summary>
        /// 统一设备分组键的文本格式，避免空格和大小写差异造成重复行。
        /// </summary>
        private static string NormalizeDeviceGroupingPart(string value)
        {
            return (value ?? string.Empty).Trim().ToUpperInvariant();
        }

        /// <summary>
        /// 螺栓统计表中的单条汇总数据。
        /// </summary>
        private sealed class BoltStatisticRow
        {
            public string Specification { get; set; } = string.Empty;
            public int Quantity { get; set; }
            public string Length { get; set; } = string.Empty;
            public string Material { get; set; } = string.Empty;
        }

        /// <summary>
        /// 记录单张设备表创建后的实体、位置和高度，供批量排列使用。
        /// </summary>
        private sealed class DeviceTableCreateResult
        {
            public ObjectId TableId { get; set; } = ObjectId.Null;
            public Point3d Position { get; set; }
            public double Height { get; set; }
        }

        /// <summary>
        /// 从本次选中的图元中提取螺栓统计数据。
        /// 仅统计法兰/对夹连接图元，以及名称包含“法兰”或“管端盲板”的部件。
        /// </summary>
        private List<BoltStatisticRow> BuildBoltStatistics(
            IEnumerable<DeviceInfo> devices,
            out int candidateCount,
            out int skippedCount)
        {
            candidateCount = 0;
            skippedCount = 0;
            if (devices == null)
            {
                return new List<BoltStatisticRow>();
            }

            var sourceRows = new List<BoltStatisticRow>();
            foreach (DeviceInfo device in devices.Where(d => d != null))
            {
                if (!IsBoltStatisticCandidate(device))
                {
                    // 记录未命中原因，便于确认实际选择的图元是否包含螺栓属性。
                    LogManager.Instance.LogInfo(
                        $"[螺栓统计][候选跳过] Name={device.Name}, AttributeCount={device.Attributes?.Count ?? 0}, " +
                        $"HasBoltAttribute={HasAnyAttributeValue(device.Attributes, "BOLT_QTY", "BOLT_LENGTH", "BOLT_MATL", "BOLT_SPEC", "BOLT_SPECIFICATION")}");
                    continue;
                }

                candidateCount++;

                string specification = GetFirstAttributeValue(device.Attributes,
                    "螺栓规格", "BOLT_SPEC", "BOLT_SPECIFICATION", "BOLT_SIZE", "BoltSpec", "BoltSpecification");
                // 螺栓数量以属性块中的 BOLT_QTY 为准；其余名称仅用于兼容历史图元。
                int quantity = ParsePositiveInteger(GetFirstAttributeValue(device.Attributes,
                    "BOLT_QTY", "螺栓数量", "螺栓数量n", "螺栓数", "BOLT_HOLES", "BOLT_COUNT", "BOLT_NUM", "BoltCount"));
                // 螺栓长度以属性块中的 BOLT_LENGTH 为准；没有该 Tag 时再读取历史别名。
                string length = GetFirstAttributeValue(device.Attributes,
                    "BOLT_LENGTH", "螺栓长度", "BOLT_LEN", "BoltLength");
                // 螺栓材质以属性块中的 BOLT_MATL 为准；没有该 Tag 时再读取历史别名。
                string material = GetFirstAttributeValue(device.Attributes,
                    "BOLT_MATL", "螺栓材质", "螺栓材料", "BOLT_MATERIAL", "BoltMaterial", "MATERIAL", "材质", "材料");

                // 记录螺栓字段的原始读取结果，确认统计表使用的是属性块实际值。
                LogManager.Instance.LogInfo(
                    $"[螺栓统计][候选读取] Name={device.Name}, Specification={specification}, " +
                    $"Quantity={quantity}, Length={length}, Material={material}");

                NormalizeBoltSpecification(ref specification, ref length);
                if (string.IsNullOrWhiteSpace(specification) || quantity <= 0)
                {
                    skippedCount++;
                    continue;
                }

                sourceRows.Add(new BoltStatisticRow
                {
                    Specification = specification,
                    Quantity = quantity,
                    Length = length,
                    Material = material
                });
            }

            return sourceRows
                .GroupBy(row => string.Join("|", row.Specification, row.Length, row.Material), StringComparer.OrdinalIgnoreCase)
                .Select(group => new BoltStatisticRow
                {
                    Specification = group.First().Specification,
                    Quantity = group.Sum(row => row.Quantity),
                    Length = group.First().Length,
                    Material = group.First().Material
                })
                .OrderBy(row => GetBoltSpecificationSortValue(row.Specification))
                .ThenBy(row => row.Specification, StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.Length, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>
        /// 判断图元是否属于需要统计螺栓的法兰连接部件。
        /// </summary>
        private static bool IsBoltStatisticCandidate(DeviceInfo device)
        {
            // 只要存在明确的螺栓属性，就应当纳入统计，不能强制依赖连接方式或块名称。
            if (HasAnyAttributeValue(device.Attributes,
                "BOLT_QTY", "BOLT_LENGTH", "BOLT_MATL", "BOLT_SPEC", "BOLT_SPECIFICATION", "BOLT_SIZE"))
            {
                return true;
            }

            string connectionType = GetFirstAttributeValue(device.Attributes,
                "连接方式", "连接形式", "CONN_TYPE", "CONNTYPE", "DNCONN_TYPE", "CONNECTION_TYPE", "CONNECTION_MODE", "ConnectionType");
            string normalizedConnectionType = Regex.Replace(connectionType ?? string.Empty, @"\s|[-_/]", string.Empty);
            bool isFlangeConnection = string.Equals(normalizedConnectionType, "法兰", StringComparison.OrdinalIgnoreCase) ||
                                      string.Equals(normalizedConnectionType, "法兰连接", StringComparison.OrdinalIgnoreCase) ||
                                      string.Equals(normalizedConnectionType, "对夹", StringComparison.OrdinalIgnoreCase) ||
                                      string.Equals(normalizedConnectionType, "对夹连接", StringComparison.OrdinalIgnoreCase);
            if (isFlangeConnection)
            {
                return true;
            }

            string componentName = GetFirstAttributeValue(device.Attributes, "部件名称", "组件名称", "名称", "COMPONENT_NAME");
            string name = string.Join(" ", new[] { device.Name, componentName }.Where(value => !string.IsNullOrWhiteSpace(value)));
            return name.IndexOf("法兰", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("管端盲板", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// 判断属性字典中是否至少存在一个指定的非空属性值。
        /// </summary>
        private static bool HasAnyAttributeValue(Dictionary<string, string> attributes, params string[] keys)
        {
            return keys != null && keys.Any(key => !string.IsNullOrWhiteSpace(GetFirstAttributeValue(attributes, key)));
        }

        /// <summary>
        /// 按候选 Tag 读取第一个非空属性值，兼容不同块库的中英文属性名。
        /// </summary>
        //private static string GetFirstAttributeValue(Dictionary<string, string> attributes, params string[] keys)
        //{
        //    if (attributes == null || keys == null)
        //    {
        //        return string.Empty;
        //    }

        //    foreach (string key in keys)
        //    {
        //        if (attributes.TryGetValue(key, out string value) && !string.IsNullOrWhiteSpace(value))
        //        {
        //            return value.Trim();
        //        }
        //    }

        //    foreach (KeyValuePair<string, string> item in attributes)
        //    {
        //        if (string.IsNullOrWhiteSpace(item.Key) || string.IsNullOrWhiteSpace(item.Value))
        //        {
        //            continue;
        //        }

        //        string normalizedAttributeKey = NormalizeAttributeLookupKey(item.Key);
        //        if (keys.Any(key => string.Equals(
        //            normalizedAttributeKey,
        //            NormalizeAttributeLookupKey(key),
        //            StringComparison.OrdinalIgnoreCase)))
        //        {
        //            return item.Value.Trim();
        //        }
        //    }

        //    return string.Empty;
        //}

        /// <summary>
        /// 归一化属性 Tag，兼容 BOLT_COUNT、Bolt Count 等仅分隔符不同的名称。
        /// </summary>
        private static string NormalizeAttributeLookupKey(string key)
        {
            return Regex.Replace(key ?? string.Empty, @"[\s_\-\.\(\)]+", string.Empty).Trim();
        }

        /// <summary>
        /// 解析螺栓数量，兼容“8 个”“8套”等带单位的属性值。
        /// </summary>
        private static int ParsePositiveInteger(string value)
        {
            System.Text.RegularExpressions.Match match = Regex.Match(value ?? string.Empty, @"\d+");
            return match.Success && int.TryParse(match.Value, out int quantity) && quantity > 0 ? quantity : 0;
        }

        /// <summary>
        /// 拆分“ M16×80 ”等规格中的长度；独立长度属性优先保留。
        /// </summary>
        private static void NormalizeBoltSpecification(ref string specification, ref string length)
        {
            specification = (specification ?? string.Empty).Trim();
            length = (length ?? string.Empty).Trim();
            System.Text.RegularExpressions.Match match = Regex.Match(specification, @"(?i)M\s*(?<diameter>\d+(?:\.\d+)?)\s*(?:[x×*]\s*(?<length>\d+(?:\.\d+)?))?");
            if (!match.Success)
            {
                return;
            }

            specification = "M" + match.Groups["diameter"].Value;
            if (string.IsNullOrWhiteSpace(length) && match.Groups["length"].Success)
            {
                length = match.Groups["length"].Value;
            }
        }

        /// <summary>
        /// 提取规格中的直径数值，用于按 M12、M16、M20 的自然顺序排列。
        /// </summary>
        private static decimal GetBoltSpecificationSortValue(string specification)
        {
            System.Text.RegularExpressions.Match match = Regex.Match(specification ?? string.Empty, @"(?i)M\s*(\d+(?:\.\d+)?)");
            return match.Success && decimal.TryParse(match.Groups[1].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal value)
                ? value
                : decimal.MaxValue;
        }

        /// <summary>
        /// 按设备表的样式创建四列螺栓统计表。
        /// </summary>
        private DeviceTableCreateResult CreateBoltStatisticsTable(Document document, List<BoltStatisticRow> rows, double scaleDenominator, Point3d insertPosition)
        {
            // 统一使用调用方传入的活动文档，避免编辑器和数据库来自不同图纸。
            if (document == null || rows == null || rows.Count == 0)
            {
                return new DeviceTableCreateResult();
            }

            // 从同一个文档获取编辑器和数据库。
            Editor editor = document.Editor;
            Database database = document.Database;
            if (editor == null)
            {
                return new DeviceTableCreateResult();
            }

            // 使用设备表批次计算出的最终位置，螺丝统计表不再额外要求用户指定点位。
            ObjectId currentSpaceId = database.CurrentSpaceId;
            LogManager.Instance.LogInfo(
                $"[螺栓统计表][插入准备] Point={insertPosition}, TileMode={database.TileMode}, CurrentSpaceId={currentSpaceId}");

            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTableRecord currentSpace = transaction.GetObject(currentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    if (currentSpace == null)
                    {
                        LogManager.Instance.LogWarning($"[螺栓统计表][插入失败] CurrentSpaceId={currentSpaceId} 对应的空间对象为空。");
                        return new DeviceTableCreateResult();
                    }

                    LogManager.Instance.LogInfo(
                        $"[螺栓统计表][当前空间] Name={currentSpace.Name}, IsLayout={currentSpace.IsLayout}, IsAnonymous={currentSpace.IsAnonymous}");

                    const int columnCount = 4;
                    Table table = new Table();
                    table.SetSize(rows.Count + 2, columnCount);
                    table.Position = insertPosition;
                    SetTableStyle(database, table, transaction, scaleDenominator);
                    table.MergeCells(CellRange.Create(table, 0, 0, 0, columnCount - 1));
                    table.Cells[0, 0].TextString = "螺栓统计表";
                    table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

                    string[] headers = { "螺栓规格", "螺栓数量", "长度", "材质" };
                    for (int columnIndex = 0; columnIndex < headers.Length; columnIndex++)
                    {
                        table.Cells[1, columnIndex].TextString = headers[columnIndex];
                        table.Cells[1, columnIndex].Alignment = CellAlignment.MiddleCenter;
                    }

                    for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                    {
                        BoltStatisticRow row = rows[rowIndex];
                        int tableRowIndex = rowIndex + 2;
                        table.Cells[tableRowIndex, 0].TextString = row.Specification;
                        table.Cells[tableRowIndex, 1].TextString = row.Quantity.ToString(CultureInfo.InvariantCulture);
                        table.Cells[tableRowIndex, 2].TextString = row.Length;
                        table.Cells[tableRowIndex, 3].TextString = row.Material;
                        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                        {
                            table.Cells[tableRowIndex, columnIndex].Alignment = CellAlignment.MiddleCenter;
                        }
                    }

                    ApplyScaledHeightsToTable(table, scaleDenominator);
                    AutoFitTableColumnsAdvanced(table, scaleDenominator);
                    table.GenerateLayout();
                    double tableHeight = GetTableHeight(table);
                    currentSpace.AppendEntity(table);
                    transaction.AddNewlyCreatedDBObject(table, true);
                    LogManager.Instance.LogInfo(
                        $"[螺栓统计表][实体追加] Space={currentSpace.Name}, TableObjectId={table.ObjectId}, Position={table.Position}");
                    transaction.Commit();
                    LogManager.Instance.LogInfo(
                        $"[螺栓统计表][插入完成] Space={currentSpace.Name}, TableHandle={table.Handle}");
                    return new DeviceTableCreateResult
                    {
                        TableId = table.ObjectId,
                        Position = insertPosition,
                        Height = tableHeight
                    };
                }
                catch
                {
                    transaction.Abort();
                    throw;
                }
            }
        }

        /// <summary>
        /// 深拷贝一个设备信息对象，避免共享引用导致后续修改污染
        /// </summary>
        private DeviceInfo CloneDeviceInfo(DeviceInfo source)
        {
            if (source == null) return new DeviceInfo(); // 防御：空源返回新实例

            // 创建新的 DeviceInfo 并逐字段复制（基本类型与字符串直接复制）
            var clone = new DeviceInfo
            {
                Id = source.Id,
                Name = source.Name,
                Type = source.Type,
                MediumName = source.MediumName,
                Specifications = source.Specifications,
                Material = source.Material,
                Quantity = source.Quantity,
                DrawingNumber = source.DrawingNumber,
                Power = source.Power,
                Volume = source.Volume,
                Pressure = source.Pressure,
                Temperature = source.Temperature,
                Diameter = source.Diameter,
                Length = source.Length,
                Thickness = source.Thickness,
                Weight = source.Weight,
                Model = source.Model,
                DeviceSTDNo = source.DeviceSTDNo,
                Remarks = source.Remarks,
                Count = source.Count
            };

            // 字典使用新的实例并拷贝键值（忽略空源）
            clone.Attributes = source.Attributes != null
                ? new Dictionary<string, string>(source.Attributes, StringComparer.OrdinalIgnoreCase)
                : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            clone.EnglishNames = source.EnglishNames != null
                ? new Dictionary<string, string>(source.EnglishNames, StringComparer.OrdinalIgnoreCase)
                : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            return clone; // 返回深拷贝对象
        }

        #endregion

        /// <summary>
        /// 自动调整表格列宽、统一所有单元格的文字高度，并根据用户比例缩放所有尺寸
        /// </summary>
        /// <param name="table">CAD表格对象</param>
        /// <summary>
        /// 自动调整CAD表格：统一设置字高（第一行3倍比例，其余行2.5倍比例），
        /// 并根据文本内容自动调整列宽，同时设置合适的行高。
        /// </summary>
        /// <param name="table">要处理的CAD表格对象</param>
        private void AutoFitTableColumnsAdvanced(Table table, double scaleDenominator)
        {
            // 1. 如果传入的表格对象为空，则直接返回，避免后续操作出错。
            if (table == null) return;

            // 2. 必须使用创建表格时传入的比例，避免自动读取的当前视口比例与本表不一致。
            double textInputScale = scaleDenominator > 0.0 ? scaleDenominator : AutoCadHelper.GetScale();
            if (textInputScale <= 0.0) textInputScale = 1.0;

            // 3. 获取表格的总行数和总列数。
            int numRows = table.Rows.Count;
            int numCols = table.Columns.Count;

            // 4. 如果表格没有任何行或列，无需调整，直接返回。
            if (numRows == 0 || numCols == 0) return;

            // ==================== 第一步：按设备表实际行类型设置字高 ====================
            // 5. 与 ApplyScaledHeightsToTable 保持相同的标题、表头和数据行字高。
            double titleHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, textInputScale);
            double headerHeight = TextFontsStyleHelper.ComputeScaledHeight(2.5, textInputScale);
            double contentHeight = TextFontsStyleHelper.ComputeScaledHeight(2.5, textInputScale);

            // 6. 先写入单元格字高，再根据最终字高计算列宽。
            try
            {
                for (int row = 0; row < numRows; row++)
                {
                    // 第0行为标题，第1行为列标题，其余为数据行。
                    double rowTextHeight = row == 0 ? titleHeight : (row == 1 ? headerHeight : contentHeight);

                    for (int col = 0; col < numCols; col++)
                    {
                        var cell = table.Cells[row, col];
                        if (cell != null)
                            cell.TextHeight = rowTextHeight;
                    }
                }
            }
            catch { /* 忽略异常：例如表格被锁定或某些单元格无法设置，不影响后续列宽计算 */ }

            // ==================== 第二步：设置列宽（使用每个单元格实际已经设置好的字高） ====================
            // 13. 逐列处理，每一列独立计算所需的最大宽度
            for (int col = 0; col < numCols; col++)
            {
                // 14. 用于记录当前列中所有单元格所需的最大宽度（初始为0）
                double maxWidthInCol = 0.0;

                // 15. 遍历当前列的所有行
                for (int row = 0; row < numRows; row++)
                {
                    // 16. 获取当前单元格对象
                    var cell = table.Cells[row, col];

                    // ---------- 合并单元格处理：只让合并区域的左上角参与宽度计算，避免重复 ----------
                    // 17. 判断当前单元格是否属于合并区域
                    bool isMerged = Convert.ToBoolean(cell.IsMerged);
                    bool shouldSkip = false;   // 标记是否跳过本次宽度计算

                    if (isMerged)
                    {
                        // 18. 假设当前单元格是合并区域的左上角，然后向左、向上验证
                        bool isTopLeft = true;

                        // 19. 如果左边有列，且左边单元格也是合并状态，说明当前单元格不是最左列
                        if (col > 0 && Convert.ToBoolean(table.Cells[row, col - 1].IsMerged))
                            isTopLeft = false;

                        // 20. 如果上边有行，且上边单元格也是合并状态，说明当前单元格不是最上行
                        if (row > 0 && Convert.ToBoolean(table.Cells[row - 1, col].IsMerged))
                            isTopLeft = false;

                        // 21. 如果不是左上角，则标记跳过，不参与宽度计算
                        if (!isTopLeft) shouldSkip = true;
                    }

                    // 22. 如果需要跳过（非左上角的合并单元格），直接进入下一个单元格
                    if (shouldSkip) continue;

                    // 23. 获取单元格的文本内容。
                    string cellText = cell.TextString;

                    // 24. 如果文本为空或仅包含空白字符，则跳过此单元格（不贡献宽度）
                    if (string.IsNullOrWhiteSpace(cellText)) continue;

                    // 25. 关键：获取该单元格已经被设置好的实际字高（可能是 headerHeight 或 contentHeight）
                    double cellTextHeight = Convert.ToDouble(cell.TextHeight);

                    // 26. 如果获取的字高无效（<=0），则使用内容行的字高作为保底值
                    if (cellTextHeight <= 0) cellTextHeight = contentHeight;

                    // 27. 文本可能包含换行符，需要按行计算宽度，取最长的一行
                    var lines = cellText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    double maxLineWidth = 0.0;  // 记录当前单元格中最长行的宽度

                    // 28. 遍历每一行文本
                    foreach (var line in lines)
                    {
                        // 29. 按显示宽度估算全角字符和半角字符，兼容中文标点及其他全角字符。
                        int chnCount = System.Text.RegularExpressions.Regex.Matches(line, @"[\u4e00-\u9fff]").Count;
                        int fullWidthCount = line.Count(character => character > 255) - chnCount;
                        int otherCount = line.Length - chnCount - fullWidthCount;

                        // 30. 预留 AutoCAD 字体实际字面宽度、标点和格式差异造成的误差。
                        double lineWidth = (chnCount * 1.25 + fullWidthCount * 1.25 + otherCount * 0.75) * cellTextHeight;

                        // 32. 保留所有行中的最大宽度
                        if (lineWidth > maxLineWidth) maxLineWidth = lineWidth;
                    }

                    // 33. 单元格总宽度增加足够的左右边距，避免文字贴线或越过表线。
                    double estimatedWidth = (maxLineWidth + cellTextHeight * 2.0) * 1.15;

                    // 34. 如果当前单元格的估算宽度大于当前列之前记录的最大宽度，则更新列的最大宽度
                    if (estimatedWidth > maxWidthInCol)
                        maxWidthInCol = estimatedWidth;
                }

                // 35. 列宽保护：设置最小宽度，防止空值列过窄。
                double minWidth = contentHeight * 1.0;
                if (maxWidthInCol < minWidth) maxWidthInCol = minWidth;

                // 36. 使用 AutoCAD 当前版本实际写入表格实体的列宽接口，确保边界线同步更新。
                table.SetColumnWidth(col, maxWidthInCol);
            }

            // ==================== 第三步：设置与字高匹配的最终行高 ====================
            for (int row = 0; row < numRows; row++)
            {
                // 让行高与 ApplyScaledHeightsToTable 的规则完全一致，避免后续布局再次压缩文字。
                double rowHeight = row == 0
                    ? Math.Max(8.0, titleHeight * 3.0)
                    : row == 1
                        ? Math.Max(6.0, headerHeight * 3)
                        : Math.Max(5.0, contentHeight * 3);
                table.SetRowHeight(row, rowHeight);
            }

            // 重新生成布局，使列线、行线和文字使用同一组最终尺寸。
            table.GenerateLayout();
        }

        #region 新生成管道表\导出管道表方法

        /// <summary>
        /// 生成(工艺)管道表：交互选择实体 → 提取数据 → 在 CAD 中绘制表格
        /// </summary>
        [CommandMethod(nameof(GeneratePipeTableFromSelection))]
        public void GeneratePipeTableFromSelection()
        {
            // 获取当前活动文档。
            Document doc =
                Application.DocumentManager.MdiActiveDocument;

            // 当前没有活动文档时直接结束。
            if (doc == null)
            {
                return;
            }

            // 使用文档事务执行管道选择、数据读取和表格生成。
            AutoCadHelper.ExecuteInDocumentTransaction((document, transaction) =>
            {
                // 获取当前文档编辑器。
                Editor ed = document.Editor;

                try
                {
                    // 提示用户选择当前生成的管道主体。
                    // 当前正式管道由 PipelineCadObjectService 创建为带 PIPEID 的 Polyline。
                    ed.WriteMessage(
                        "\n开始生成管道表：请选择当前生成的管道主体，完成后按回车。");

                    // 创建选择过滤器，只允许选择当前管道主体使用的轻量多段线。
                    TypedValue[] filterValues =
                    {
                        // 开始 OR 条件，兼容当前轻量多段线和旧版二维多段线。
                        new TypedValue(
                            (int)DxfCode.Operator,
                            "<OR"),

                        // 当前管道主体通常是 LWPOLYLINE。
                        new TypedValue(
                            (int)DxfCode.Start,
                            "LWPOLYLINE"),

                        // 兼容历史管道可能使用的 POLYLINE。
                        new TypedValue(
                            (int)DxfCode.Start,
                            "POLYLINE"),

                        // 结束 OR 条件。
                        new TypedValue(
                            (int)DxfCode.Operator,
                            "OR>")
                    };

                    // 根据实体类型创建管道主体过滤器。
                    SelectionFilter pipeFilter =
                        new SelectionFilter(filterValues);

                    // 设置管道主体选择提示。
                    PromptSelectionOptions selectionOptions =
                        new PromptSelectionOptions
                        {
                            // 提示用户选择 Polyline 管道主体。
                            MessageForAdding =
                                "\n请选择要统计的管道主体：",

                            // 不允许重复选择同一个管道主体。
                            AllowDuplicates = false
                        };

                    // 只选择 Polyline，避免把管道标题和流向符号选入管道记录。
                    PromptSelectionResult selectionResult =
                        ed.GetSelection(
                            selectionOptions,
                            pipeFilter);

                    // 用户取消选择或没有选择有效对象时结束。
                    if (selectionResult.Status != PromptStatus.OK ||
                        selectionResult.Value == null)
                    {
                        ed.WriteMessage(
                            "\n未选择管道块或选择已取消。");

                        LogManager.Instance.LogWarning(
                            $"[管道表][选择取消] Status={selectionResult.Status}");

                        return;
                    }

                    // 获取用户选择的块参照对象 ID。
                    ObjectId[] selectedIds =
                        selectionResult.Value.GetObjectIds();

                    // 记录本次选择的原始对象数量。
                    LogManager.Instance.LogInfo(
                        $"[管道表][选择完成] SelectedObjectCount={selectedIds.Length}");

                    // 创建有效管道主体对象 ID 列表。
                    List<ObjectId> pipeIds =
                        new List<ObjectId>();

                    // 遍历所有用户选择的对象，使用与右键“管道修改”相同的 PIPEID 判定规则。
                    foreach (ObjectId objectId in selectedIds)
                    {
                        // 以只读方式打开当前管道对象。
                        Entity entity =
                            transaction.GetObject(
                                objectId,
                                OpenMode.ForRead) as Entity;

                        // 当前对象不是 Polyline 时跳过。
                        Polyline pipeline = entity as Polyline;
                        if (pipeline == null)
                        {
                            LogManager.Instance.LogWarning(
                                $"[管道表][跳过非Polyline] ObjectId={objectId}, " +
                                $"EntityType={entity?.GetType().Name ?? "null"}");

                            continue;
                        }

                        // 使用管道修改页面相同的统一属性读取方法。
                        // 该方法能够读取编码后的属性 Tag 和主体扩展字典中的成对属性数据。
                        Dictionary<string, string> attributes =
                            PipelineEndpointPropertyHelper.ReadEntityProperties(
                                transaction,
                                pipeline);

                        // PIPEID 是当前正式管道主体的唯一识别标识。
                        string pipeId = GetFirstAttributeValue(
                            attributes,
                            "PIPEID");

                        // 没有 PIPEID 的普通 Polyline 不是当前生成的正式管道。
                        if (string.IsNullOrWhiteSpace(pipeId))
                        {
                            LogManager.Instance.LogWarning(
                                $"[管道表][跳过非正式管道] ObjectId={objectId}, " +
                                "Reason=PIPEID为空");

                            continue;
                        }

                        // 保存有效的当前管道主体对象 ID。
                        // 起点、终点和管段号由统一属性读取方法提供，不再使用旧版属性扫描逻辑判定管道。
                        pipeIds.Add(objectId);

                        // 读取管段号，便于核对当前对象是否携带正确的管道属性。
                        string pipeNumber =
                            GetFirstAttributeValue(
                                attributes,
                                "TAG_NO",
                                "管段号",
                                "管段编号",
                                "Pipeline No",
                                "Pipeline",
                                "Pipe No");

                        // 记录当前正式管道的对象 ID、PIPEID 和管段号。
                        LogManager.Instance.LogInfo(
                            $"[管道表][接受正式管道] ObjectId={objectId}, " +
                            $"PipeId={pipeId}, PipeNumber={pipeNumber}, " +
                            $"AttributeCount={attributes.Count}");
                    }

                    // 没有发现可处理的管道主体时结束。
                    if (pipeIds.Count == 0)
                    {
                        ed.WriteMessage(
                            "\n选择的对象中没有识别到正式管道，请确认选择的是带 PIPEID 的管道主体。");

                        LogManager.Instance.LogWarning(
                            "[管道表][无有效管道主体] 未找到带 PIPEID 的管道 Polyline。");

                        return;
                    }

                    // 使用当前管道块的统一提取方法读取管道数据。
                    List<DeviceInfo> finalList =
                        ExtractPipeDataFromObjectIds(
                            pipeIds.ToArray(),
                            transaction,
                            ed);

                    // 没有提取到数据时结束。
                    if (finalList == null ||
                        finalList.Count == 0)
                    {
                        ed.WriteMessage(
                            "\n已识别管道块，但没有提取到有效管道数据。");

                        LogManager.Instance.LogWarning(
                            $"[管道表][数据为空] PipeBlockCount={pipeIds.Count}");

                        return;
                    }

                    // 获取当前图纸比例分母。
                    double scaleDenominator =
                        AutoCadHelper.GetScale();

                    // 比例分母无效时使用 1，避免表格尺寸计算异常。
                    if (scaleDenominator <= 0.0 ||
                        double.IsNaN(scaleDenominator) ||
                        double.IsInfinity(scaleDenominator))
                    {
                        scaleDenominator = 1.0;
                    }

                    // 记录最终参与生成管道表的数据数量。
                    LogManager.Instance.LogInfo(
                        $"[管道表][数据提取完成] " +
                        $"PipeBlockCount={pipeIds.Count}, " +
                        $"RecordCount={finalList.Count}, " +
                        $"ScaleDenominator={scaleDenominator}");

                    // 保持现有管道表的格式和绘制方法不变。
                    // 这里只替换数据来源，不修改表头、列顺序、边框和表格样式。
                    CreateDeviceTableWithType(
                        document.Database,
                        finalList,
                        "管道明细",
                        scaleDenominator);

                    // 向命令行输出生成结果。
                    ed.WriteMessage(
                        $"\n管道表已生成，共 {finalList.Count} 条记录（使用比例分母 {scaleDenominator}）。");

                    // 记录管道表生成成功日志。
                    LogManager.Instance.LogInfo(
                        $"[管道表][生成完成] RecordCount={finalList.Count}");
                }
                catch (System.Exception ex)
                {
                    // 记录完整异常信息，便于定位管道表生成失败原因。
                    LogManager.Instance.LogError(
                        $"[管道表][生成失败] Error={ex}");

                    // 向 AutoCAD 命令行输出异常信息。
                    ed.WriteMessage(
                        $"\n生成管道表时发生错误：{ex.Message}");
                }
            });
        }

        /// <summary>
        /// 判断属性字典中是否存在指定候选属性，并且属性值不为空。
        /// </summary>
        private bool ContainsAttribute(
            Dictionary<string, string> attributes,
            params string[] keys)
        {
            // 属性字典为空时，直接返回 false。
            if (attributes == null ||
                keys == null)
            {
                return false;
            }

            // 遍历所有候选属性名。
            foreach (string key in keys)
            {
                // 跳过空的候选属性名。
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                // 属性存在且值不为空时，认为匹配成功。
                if (attributes.TryGetValue(
                        key,
                        out string value) &&
                    !string.IsNullOrWhiteSpace(value))
                {
                    return true;
                }
            }

            // 没有找到有效属性时返回 false。
            return false;
        }

        /// <summary>
        /// 按候选属性名顺序获取第一个非空属性值。此处示例实现保留原有查找逻辑，返回第一个匹配的非空属性值
        /// </summary>
        private static string GetFirstAttributeValue(Dictionary<string, string> attrs, params string[] keys)
        {
            // 防御性检查
            if (attrs == null || keys == null || keys.Length == 0) return string.Empty;

            // 先按传入顺序尝试直匹配（不区分大小写）
            foreach (var k in keys)
            {
                if (string.IsNullOrWhiteSpace(k)) continue;
                if (attrs.TryGetValue(k, out var v) && !string.IsNullOrWhiteSpace(v)) return v;
            }

            // 再按字典中的键做不区分大小写的匹配作为兜底
            foreach (var kv in attrs)
            {
                foreach (var k in keys)
                {
                    if (string.IsNullOrWhiteSpace(k)) continue;
                    if (string.Equals(kv.Key, k, StringComparison.OrdinalIgnoreCase)
                        && !string.IsNullOrWhiteSpace(kv.Value))
                    {
                        return kv.Value;
                    }
                }
            }
            return string.Empty;
        }


        /// <summary>
        /// 【核心方法】从指定的 ObjectId[] 中提取管道数据，返回排序后的 DeviceInfo 列表。
        /// 不包含选择交互，不负责输出（CAD表/Excel），供 Generate 和 Export 共用。
        /// </summary>
        /// <param name="selIds">已选择的实体 ObjectId 数组</param>
        /// <param name="tr">活动事务</param>
        /// <param name="ed">编辑器（用于日志输出）</param>
        /// <returns>排序后的管道数据列表，无数据时返回空列表</returns>
        private List<DeviceInfo> ExtractPipeDataFromObjectIds(ObjectId[] selIds, Transaction tr, Editor ed)
        {
            const double unitToMeters = 1000.0;
            string[] pipeNoKeys = new[] { "管段号", "管段编号", "Pipeline No", "Pipeline", "Pipe No", "TAG_NO" };
            string[] startKeys = new[] { "起点", "START_POINT", "From" };
            string[] endKeys = new[] { "终点", "END_POINT", "To" };

            // 局部函数：按候选键优先、再按模糊键名匹配
            string FindFirstAttrValueLocal(Dictionary<string, string>? attrs, string[] candidates)
            {
                if (attrs == null) return string.Empty;// 防御：空字典返回空字符串
                foreach (var c in candidates) // 优先按候选键名顺序查找
                {
                    // 如果字典中存在该键且值非空，则直接返回该值
                    if (attrs.TryGetValue(c, out var v) && !string.IsNullOrWhiteSpace(v)) return v;
                }
                foreach (var kv in attrs) // 模糊匹配：按包含关系查找键名
                {
                    if (string.IsNullOrWhiteSpace(kv.Key)) continue; // 跳过空键
                    foreach (var c in candidates) // 遍历候选键名，忽略大小写匹配
                    {
                        // 如果当前字典键名包含候选键名且值非空，则返回该值
                        if (kv.Key.IndexOf(c, StringComparison.OrdinalIgnoreCase) >= 0 && !string.IsNullOrWhiteSpace(kv.Value))
                            return kv.Value; // 找到首个匹配的值并返回
                    }
                }
                return string.Empty;
            }
            // 局部函数：从属性值中解析长度（支持带单位的字符串，如 "12.5m"、"12500mm"、"41ft" 等）
            var perPipeList = new List<DeviceInfo>();
            int seqIndex = 0; // 用于生成唯一的设备名称（PIPE_1, PIPE_2, ...）
            // 遍历所有选择的实体 ID
            foreach (var id in selIds)
            {
                seqIndex++; // 序号递增，用于生成唯一名称
                try
                {
                    // --- 获取实体对象 ---
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;
                    // --- 获取管道属性字典 ---
                    // 管道修改页面使用 PipelineEndpointPropertyHelper 读取主体属性，
                    // 这里必须复用同一方法，才能正确解析 PIPEID 和扩展字典中的成对属性。
                    var attrMap = ent is Polyline pipelineEntity
                        ? PipelineEndpointPropertyHelper.ReadEntityProperties(tr, pipelineEntity)
                        : GetEntityAttributeMap(tr, ent);

                    // --- 长度提取 ---
                    double length_m = double.NaN;
                    if (attrMap != null)
                    {
                        foreach (var k in attrMap.Keys)
                        {
                            if (!string.IsNullOrWhiteSpace(k) &&
                                (k.IndexOf("长度", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 string.Equals(k, "PIPE_LENGTH", StringComparison.OrdinalIgnoreCase) ||
                                 string.Equals(k, "LENGTH", StringComparison.OrdinalIgnoreCase)))
                            {
                                var parsed = ParseLengthValueFromAttribute(attrMap[k]);
                                if (!double.IsNaN(parsed) && parsed > 0.0) { length_m = parsed; break; }
                            }
                        }
                    }
                    if (double.IsNaN(length_m))
                    {
                        if (ent is Line lineEnt)
                            length_m = lineEnt.Length / unitToMeters;
                        else if (ent is Polyline plEnt)
                            length_m = plEnt.Length / unitToMeters;
                        else if (ent is BlockReference brEnt)
                        {
                            try { double l = DynamicBlockOperations.GetLength(brEnt); if (!double.IsNaN(l) && l > 0.0) length_m = l; } catch { }
                        }
                        if (double.IsNaN(length_m))
                        {
                            try { var ext = ent.GeometricExtents; length_m = Math.Abs(ext.MaxPoint.X - ext.MinPoint.X) / unitToMeters; }
                            catch { length_m = 0.0; }
                        }
                    }

                    // --- 起点/终点提取 ---
                    string startStr = FindFirstAttrValueLocal(attrMap, startKeys);
                    string endStr = FindFirstAttrValueLocal(attrMap, endKeys);
                    if (string.IsNullOrWhiteSpace(startStr) || string.IsNullOrWhiteSpace(endStr))
                    {
                        if (ent is Line lineEnt2)
                        {
                            if (string.IsNullOrWhiteSpace(startStr)) startStr = $"X={lineEnt2.StartPoint.X:F3},Y={lineEnt2.StartPoint.Y:F3}";
                            if (string.IsNullOrWhiteSpace(endStr)) endStr = $"X={lineEnt2.EndPoint.X:F3},Y={lineEnt2.EndPoint.Y:F3}";
                        }
                        else if (ent is Polyline plEnt2)
                        {
                            try
                            {
                                var p0 = plEnt2.GetPoint3dAt(0);
                                var pN = plEnt2.GetPoint3dAt(plEnt2.NumberOfVertices - 1);
                                if (string.IsNullOrWhiteSpace(startStr)) startStr = $"X={p0.X:F3},Y={p0.Y:F3}";
                                if (string.IsNullOrWhiteSpace(endStr)) endStr = $"X={pN.X:F3},Y={pN.Y:F3}";
                            }
                            catch { }
                        }
                        else if (ent is BlockReference brEnt2)
                        {
                            try
                            {
                                var (s, e) = DynamicBlockOperations.GetEndPoints(brEnt2);
                                if (string.IsNullOrWhiteSpace(startStr)) startStr = $"X={s.X:F3},Y={s.Y:F3}";
                                if (string.IsNullOrWhiteSpace(endStr)) endStr = $"X={e.X:F3},Y={e.Y:F3}";
                            }
                            catch { }
                        }
                        if (string.IsNullOrWhiteSpace(startStr) || string.IsNullOrWhiteSpace(endStr))
                        {
                            try
                            {
                                var ext = ent.GeometricExtents;
                                if (string.IsNullOrWhiteSpace(startStr)) startStr = $"X={ext.MinPoint.X:F3},Y={ext.MinPoint.Y:F3}";
                                if (string.IsNullOrWhiteSpace(endStr)) endStr = $"X={ext.MaxPoint.X:F3},Y={ext.MaxPoint.Y:F3}";
                            }
                            catch
                            {
                                if (string.IsNullOrWhiteSpace(startStr)) startStr = "N/A";
                                if (string.IsNullOrWhiteSpace(endStr)) endStr = "N/A";
                            }
                        }
                    }

                    // --- 管段号提取 ---
                    string pipeNo = FindFirstAttrValueLocal(attrMap, pipeNoKeys);
                    if (string.IsNullOrWhiteSpace(pipeNo) && ent is BlockReference brEnt3)
                    {
                        try
                        {
                            var btr = tr.GetObject(brEnt3.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                            if (btr != null)
                            {
                                var m = Regex.Match(btr.Name ?? string.Empty, @"\d+");
                                if (m.Success) pipeNo = m.Value;
                            }
                        }
                        catch { }
                    }

                    // --- 构建记录 ---
                    var info = new DeviceInfo { Name = $"PIPE_{seqIndex}", Type = "管道" };
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

            if (perPipeList.Count == 0) return perPipeList;

            // --- 按管段号数字排序 ---
            int ExtractPipeNoNumber(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return int.MaxValue;
                var m = Regex.Match(s, @"\d+");
                if (m.Success && int.TryParse(m.Value, out var v)) return v;
                return int.MaxValue;
            }

            return perPipeList
                .Select((e, idx) => new { Item = e, Orig = idx })
                .OrderBy(x =>
                {
                    if (x.Item.Attributes.TryGetValue("管段号", out var pn) && !string.IsNullOrWhiteSpace(pn))
                        return (ExtractPipeNoNumber(pn), 0, x.Orig);
                    return (int.MaxValue, 1, x.Orig);
                })
                .Select(x => x.Item)
                .ToList();
        }

        /// <summary>
        /// 导出管道表到 Excel：交互选择实体 → 提取数据 → 弹出保存对话框 → 写入 .xlsx
        /// </summary>
        [CommandMethod(nameof(ExportTableToExcel))]
        public void ExportTableToExcel()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            // 1. 选择 AutoCAD 表格
            var peo = new PromptEntityOptions("\n请选择要导出的 CAD 表格：");
            peo.SetRejectMessage("请选择 AutoCAD Table 对象。");
            peo.AddAllowedClass(typeof(Table), exactMatch: true);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n未选择表格或已取消。");
                return;
            }

            List<List<string>> tableData = new List<List<string>>();// 用于存储表格文本数据的二维列表
            List<ExcelMergeRange> mergedRegions = new List<ExcelMergeRange>();// 用于存储检测到的合并区域列表
            // 开始事务读取表格数据和检测合并区域
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                var table = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table; // 获取表格对象
                if (table == null)
                {
                    ed.WriteMessage("\n所选对象不是表格。");
                    return;
                }
                //表格行数
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;// 获取表格行列数

                // 2. 读取所有文本
                for (int r = 0; r < rows; r++)
                {
                    var rowData = new List<string>(); // 存储当前行的文本数据
                    for (int c = 0; c < cols; c++) // 获取每个单元格的文本内容，空单元格会返回 null，我们转换为 "" 以便后续处理
                    {
                        rowData.Add(table.Cells[r, c].TextString ?? ""); // 为行数据添加当前单元格的文本内容
                    }
                    tableData.Add(rowData);// 将当前行的数据添加到表格数据列表中
                }

                // 3. 通过空单元格检测合并区域（修正版：严格防止重叠）
                bool[,] processed = new bool[rows, cols];

                // 先标记所有空单元格为"未被处理"（它们可能是合并子单元格，也可能是真正的空单元格）
                // 我们只把非空单元格作为合并区域的左上角来检测

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        // 跳过已被处理的单元格（包括空单元格和已归属合并区域的单元格）
                        if (processed[r, c]) continue;

                        string text = tableData[r][c];

                        // 空单元格：可能是某个合并区域的子单元格，暂时标记但不处理
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            processed[r, c] = true; // 标记为已处理，后续不会被当作起点
                            continue;
                        }

                        // 非空单元格：从此处开始向右、向下探测最大矩形
                        int maxRow = r;
                        int maxCol = c;

                        // 向右探测：检查同一行右侧连续的空单元格
                        while (maxCol + 1 < cols)
                        {
                            if (!processed[r, maxCol + 1] && string.IsNullOrWhiteSpace(tableData[r][maxCol + 1]))
                                maxCol++;
                            else
                                break;
                        }

                        // 向下探测：检查每一行，确保从 c 到 maxCol 的所有单元格都是空且未被处理
                        while (maxRow + 1 < rows)
                        {
                            bool rowAllEmpty = true;
                            for (int cc = c; cc <= maxCol; cc++)
                            {
                                if (processed[maxRow + 1, cc] || !string.IsNullOrWhiteSpace(tableData[maxRow + 1][cc]))
                                {
                                    rowAllEmpty = false;
                                    break;
                                }
                            }
                            if (rowAllEmpty)
                                maxRow++;
                            else
                                break;
                        }

                        // 如果探测到的矩形 > 1x1，则记录为合并区域
                        if (maxRow > r || maxCol > c)
                        {
                            // 标记整个矩形区域为已处理
                            for (int i = r; i <= maxRow; i++)
                            {
                                for (int j = c; j <= maxCol; j++)
                                {
                                    processed[i, j] = true;
                                }
                            }
                            mergedRegions.Add(new ExcelMergeRange(r, maxRow, c, maxCol));
                        }
                        else
                        {
                            // 单单元格，标记为已处理
                            processed[r, c] = true;
                        }
                    }
                }

                tr.Commit();
            }

            ed.WriteMessage($"\n检测到 {mergedRegions.Count} 个合并区域。");

            // 4. 导出到 Excel
            using (var sfd = new System.Windows.Forms.SaveFileDialog())
            {
                sfd.Filter = "Excel 工作簿 (*.xlsx)|*.xlsx";
                sfd.FileName = $"管道明细导出_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                if (sfd.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

                WriteTableToExcel(tableData, mergedRegions, sfd.FileName);
            }

            ed.WriteMessage($"\n表格已成功导出到 Excel，共 {tableData.Count} 行。");
        }

        /// <summary>
        /// Excel 合并区域定义，表示从 (TopRow, LeftCol) 到 (BottomRow, RightCol) 的矩形区域需要合并。
        /// </summary>
        private struct ExcelMergeRange
        {
            public int TopRow, BottomRow, LeftCol, RightCol;
            public ExcelMergeRange(int t, int b, int l, int r)
            {
                TopRow = t; BottomRow = b; LeftCol = l; RightCol = r;
            }
        }

        /// <summary>
        /// 将表格数据写入 Excel 文件，并应用合并区域。
        /// </summary>
        /// <param name="data"> 表格数据</param>
        /// <param name="merges"> 合并区域列表</param>
        /// <param name="filePath"> 文件路径</param>
        private void WriteTableToExcel(List<List<string>> data, List<ExcelMergeRange> merges, string filePath)
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Sheet1");

            // ---------- 定义样式：文字居中 + 细实线边框 ----------
            var style = workbook.CreateCellStyle();
            // 水平居中
            style.Alignment = HorizontalAlignment.Center;
            // 垂直居中
            style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            // 上边框
            style.BorderTop = BorderStyle.Thin;
            // 下边框
            style.BorderBottom = BorderStyle.Thin;
            // 左边框
            style.BorderLeft = BorderStyle.Thin;
            // 右边框
            style.BorderRight = BorderStyle.Thin;
            // 自动换行关闭（避免行高异常，如果你的内容需要换行可改为 true）
            style.WrapText = false;

            // ---------- 写入数据并应用样式 ----------
            for (int r = 0; r < data.Count; r++)
            {
                var row = sheet.CreateRow(r);// 创建新行
                for (int c = 0; c < data[r].Count; c++)
                {
                    var cell = row.CreateCell(c); // 创建新单元格
                    cell.SetCellValue(data[r][c] ?? ""); // 设置单元格文本
                    cell.CellStyle = style;   // 应用样式
                }
            }

            // ---------- 添加合并区域（带重叠保护）----------
            int addedCount = 0; // 统计成功添加的合并区域数量
            foreach (var merge in merges) // 遍历每个合并区域
            {
                try
                {
                    bool overlap = false; // 标记是否与现有合并区域重叠
                    for (int i = 0; i < sheet.NumMergedRegions; i++) // 遍历已存在的合并区域
                    {
                        var existing = sheet.GetMergedRegion(i); // 获取现有合并区域
                        if (existing.Intersects(new NPOI.SS.Util.CellRangeAddress(
                            merge.TopRow, merge.BottomRow, merge.LeftCol, merge.RightCol))) // 如果新区域与现有区域有交集，则认为重叠
                        {
                            overlap = true; // 标记为重叠
                            break;
                        }
                    }
                    // 只有在不重叠的情况下才添加合并区域，避免因重叠导致的异常
                    if (!overlap)
                    {
                        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(
                            merge.TopRow, merge.BottomRow,
                            merge.LeftCol, merge.RightCol)); // 添加合并区域
                        addedCount++; // 成功添加的合并区域数量加一
                    }
                }
                catch { /* 跳过无法添加的合并区域 */ }
            }

            // ---------- 手动计算并设置列宽 ----------
            int cols = (data.Count > 0) ? data[0].Count : 0;
            double[] maxCharWidths = new double[cols];
            for (int r = 0; r < data.Count; r++)
            {
                for (int c = 0; c < data[r].Count && c < cols; c++)
                {
                    string text = data[r][c] ?? "";
                    double w = GetTextDisplayWidth(text);
                    if (w > maxCharWidths[c]) maxCharWidths[c] = w;
                }
            }

            // 设置列宽（增加 2 个字符的缓冲，转换为 Excel 内部单位 1/256 字符）
            for (int c = 0; c < cols; c++)
            {
                int charWidth = (int)Math.Ceiling(maxCharWidths[c] + 2.0);
                sheet.SetColumnWidth(c, charWidth * 256);
            }

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                workbook.Write(fs);
            workbook.Close();
        }

        /// <summary>
        /// 估算文本在 Excel 中的显示宽度（中文字符按 2 单位，其他字符按 1 单位）
        /// </summary>
        private double GetTextDisplayWidth(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0; // 空文本宽度为 0
            double width = 0; // 遍历每个字符，累加宽度
            foreach (char ch in text)
            {
                // 中文、全角符号范围按 2 倍宽度计算
                if (ch >= 0x4e00 && ch <= 0x9fff ||
                    ch >= 0x3000 && ch <= 0x303f ||
                    ch >= 0xff00 && ch <= 0xffef)
                    width += 2.0;
                else
                    width += 1.0;
            }
            return width;
        }
        #endregion

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
            return string.Concat(matches.Cast<System.Text.RegularExpressions.Match>().Select(m => m.Value)).Trim();
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
        [CommandMethod(nameof(InsertPipeTable))]
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
        /// 辅助：从实体中读取属性（AttributeReference / Xrecord / XData）
        /// </summary>
        /// <param name="tr">事件</param>
        /// <param name="ent">选中的实体</param>
        /// <returns>返回实体属性字典</returns>
        private Dictionary<string, string> GetEntityAttributeMap(Transaction tr, Entity ent)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);// 使用不区分大小写的字典收集属性字段
            try
            {
                if (ent == null) return map;// 如果实体为空，直接返回空字典
                // 如果实体是块参照，尝试获取其属性集合
                if (ent is BlockReference br)
                {
                    try
                    {
                        var attCol = br.AttributeCollection;// 获取块参照的属性集合
                        foreach (ObjectId attId in attCol) // 遍历属性集合
                        {
                            try
                            {
                                var ar = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;// 获取属性引用对象
                                if (ar != null)
                                {
                                    var tag = (ar.Tag ?? string.Empty).Trim();// 属性标签
                                    var val = (ar.TextString ?? string.Empty).Trim();// 属性值
                                    if (!string.IsNullOrEmpty(tag) && !map.ContainsKey(tag)) map[tag] = val;// 添加属性标签到字典中，避免重复键
                                }
                            }
                            catch { /* 忽略单个属性读取失败 */ }
                        }
                    }
                    catch { /* 忽略 */ }
                }

                try
                {
                    if (ent.ExtensionDictionary != ObjectId.Null)// 检查实体是否有扩展字典
                    {
                        var extDict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary; // 获取扩展字典对象
                        if (extDict != null)
                        {
                            foreach (var entry in extDict)
                            {
                                try
                                {
                                    var xrec = tr.GetObject(entry.Value, OpenMode.ForRead) as Xrecord;// 获取 Xrecord 对象
                                    if (xrec != null && xrec.Data != null) // 检查 Xrecord 是否有数据
                                    {
                                        var vals = xrec.Data.Cast<TypedValue>().Select(tv => tv.Value?.ToString() ?? "").ToArray(); // 将 Xrecord 的数据转换为字符串数组
                                        var key = entry.Key ?? string.Empty; // 获取 Xrecord 的键名
                                        var value = string.Join("|", vals); // 将数组连接为单个字符串，使用 '|' 分隔
                                        if (!map.ContainsKey(key)) map[key] = value; // 添加到属性字典中，避免重复键
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                }
                catch { /* 忽略 */ }

                try
                {
                    var db = ent.Database; // 获取实体所在的数据库
                    var rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead); // 获取注册应用程序表
                    foreach (ObjectId appId in rat)
                    {
                        try
                        {
                            var app = tr.GetObject(appId, OpenMode.ForRead) as RegAppTableRecord; // 获取注册应用程序记录
                            if (app == null) continue; // 如果获取失败则跳过
                            var appName = app.Name; // 获取注册应用程序的名称
                            var rb = ent.GetXDataForApplication(appName); // 获取实体的 XData
                            if (rb != null)
                            {
                                var vals = rb.Cast<TypedValue>().Select(tv => tv.Value?.ToString() ?? "").ToArray(); // 将 XData 的数据转换为字符串数组
                                var key = $"XDATA:{appName}"; // 使用 "XDATA:应用程序名" 作为键名
                                var value = string.Join("|", vals); // 将数组连接为单个字符串，使用 '|' 分隔
                                if (!map.ContainsKey(key)) map[key] = value; // 添加到属性字典中，避免重复键  
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
        [CommandMethod(nameof(CollectLineInfo))]
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
        [CommandMethod(nameof(SyncPipeProperties))]
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
            var sourceLineIds = lineSelResult.Value.GetObjectIds().ToList();// 转为列表，方便后续处理

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
                            var sampleBrWrite = tr.GetObject(sampleBlockRef.ObjectId, OpenMode.ForWrite) as BlockReference;// 获取可写的示例块参照
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
                    string pipeTitle = latestSampleAttrs.TryGetValue("PIPELINETITLE", out var sampleTitle) && !string.IsNullOrWhiteSpace(sampleTitle)
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
                            if (string.IsNullOrWhiteSpace(def.Tag)) continue;// 跳过无效标签
                            if (latestDict.TryGetValue(def.Tag, out var val)) // 若示例中存在该字段，则更新其值
                            {
                                def.TextString = val ?? string.Empty; // 更新属性值
                            }
                            // 临时显示设置（随后统一隐藏/显示处理）
                            def.Invisible = false;
                            def.Constant = false;
                        }
                    }
                    else
                    {
                        // 示例无属性定义：根据 latestSampleAttrs 动态创建属性定义（按 Key 排序）
                        attDefsLocal.Clear(); // 清空原有定义
                        double attHeight = 3.5; // 默认高度
                        double yOffsetBase = -attHeight * 2.0; // 从 midPoint 向下偏移
                        int idx = 0; // 索引用于计算位置
                        foreach (var kv in latestSampleAttrs.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase)) // 按键排序
                        {
                            if (string.IsNullOrWhiteSpace(kv.Key)) continue; // 跳过无效标签
                            attDefsLocal.Add(new AttributeDefinition
                            {
                                Tag = kv.Key, // 使用属性键作为标签
                                Position = new Point3d(0, yOffsetBase - idx * attHeight * 1.2, 0), // 垂直排列
                                Rotation = 0.0, // 保持水平
                                TextString = kv.Value ?? string.Empty, // 使用属性值作为文本
                                Height = attHeight, // 设置默认高度
                                Invisible = false, // 临时显示设置（随后统一隐藏/显示处理）
                                Constant = false // 临时常量设置（随后统一取消常量处理）
                            });
                            idx++;
                        }
                    }

                    // 新增或覆盖 起点/终点 属性定义（保持原逻辑，若示例中已有这些字段则更新其值，否则新增）
                    Point3d worldStart = orderedVertices.First(); // 获取起点坐标
                    Point3d worldEnd = orderedVertices.Last(); // 获取终点坐标
                    string startCoordStr = $"X={worldStart.X:F3},Y={worldStart.Y:F3}"; // 格式化起点坐标
                    string endCoordStr = $"X={worldEnd.X:F3},Y={worldEnd.Y:F3}"; // 格式化终点坐标
                    int nextSegNum = GetNextPipeSegmentNumber(db); // 获取下一个管段号（用于默认值）

                    // 取管段号，优先从属性或标题提取
                    string extractedPipeNo = string.Empty;
                    // 尝试从属性中获取管段号
                    if (latestSampleAttrs.TryGetValue("PIPELINETITLE", out var titleFromSample) && !string.IsNullOrWhiteSpace(titleFromSample))
                    {
                        extractedPipeNo = ExtractPipeCodeFromTitle(titleFromSample); // 尝试从管道标题中提取管段号
                    }
                    // 若仍未能提取，则尝试从属性中获取 TAG_NO 或 管段编号
                    if (string.IsNullOrWhiteSpace(extractedPipeNo))
                    {
                        if (latestSampleAttrs.TryGetValue("TAG_NO", out var pn) && !string.IsNullOrWhiteSpace(pn))
                            extractedPipeNo = pn;
                    }
                    // 若仍未能提取，则使用下一个管段号作为默认值
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
                    // 设置或新增 起点/终点/管段号 属性
                    SetOrAddAttrLocal("START_POINT", startCoordStr);
                    SetOrAddAttrLocal("END_POINT", endCoordStr);
                    SetOrAddAttrLocal("TAG_NO", extractedPipeNo);

                    //// 移除块定义中的中点“管道标题”属性（避免在块中重复显示中点标题）
                    //attDefsLocal.RemoveAll(ad => string.Equals(ad.Tag, "管道标题", StringComparison.OrdinalIgnoreCase));

                    //// 将其余属性设置为隐藏（块内不显示），以保持行为与现有逻辑一致
                    foreach (var ad in attDefsLocal)
                    {
                        ad.Invisible = true;// 块内不显示
                        ad.Constant = false;// 块内不为常量
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

                    // ================== 2. 插入新管道块（原有逻辑）==================
                    var newPipeId = InsertPipeBlockWithAttributes(tr, midPoint, newBlockName, 0.0, attValues);
                    PipelineTopologyHelper.PostProcessAfterPipePlaced(tr, newPipeId, sampleIsOutlet ? "Outlet" : "Inlet");
                    var newBr = tr.GetObject(newPipeId, OpenMode.ForWrite) as BlockReference;// 插入新块并设置属性
                    if (newBr != null)
                        newBr.Layer = sampleInfo.PipeBodyTemplate.Layer;// 插入新块并设置属性 继承图层
                                                                        // ================== 3. 交叉检测 ==================
                                                                        // 提取新管道的世界坐标线段
                    var newPath = GetPipePathVertices(newBr, tr);
                    if (newPath == null || newPath.Count < 2)
                    {
                        ed.WriteMessage("\n新管道未能提取路径几何，跳过交叉检测。");
                        tr.Commit();
                        return;
                    }

                    // 遍历已有管道块
                    var existingPipes = GetExistingPipes(db, tr); // 需实现遍历模型空间块参照
                    bool hasIntersection = false;
                    BlockReference crossingPipe = null;
                    Point3d firstIntersection = Point3d.Origin;// 插入点原始坐标

                    foreach (var oldBr in existingPipes)
                    {
                        var oldPath = GetPipePathVertices(oldBr, tr);
                        if (oldPath == null || oldPath.Count < 2) continue;

                        var interPts = GetPathIntersections(newPath, oldPath);
                        if (interPts.Count > 0)
                        {
                            hasIntersection = true; // 
                            crossingPipe = oldBr;
                            firstIntersection = interPts[0]; // 取第一个交点
                            break;
                        }
                    }

                    // ================== 4. 交互处理 ==================
                    if (hasIntersection && crossingPipe != null)
                    {
                        // 获取被交叉管道的名称
                        string oldPipeName = "未知";
                        var oldAttrs = GetEntityAttributeMap(tr, crossingPipe);
                        if (oldAttrs.TryGetValue("名称", out var nm)) oldPipeName = nm;
                        else oldPipeName = crossingPipe.Name ?? "未知";
                        Double BackgroundMask = 5; //遮挡块的大小基数
                        using (var dlg = new PipeCrossingDialog(oldPipeName))
                        {
                            dlg.ShowDialog();
                            switch (dlg.SelectedAction)
                            {
                                case PipeCrossingDialog.CrossingAction.Connect:
                                    // (Y) 交叉相连 – 保持现状，可补充连接逻辑
                                    ed.WriteMessage($"\n管道 [{oldPipeName}] 与新管道交叉相连。");
                                    break;

                                case PipeCrossingDialog.CrossingAction.Cover:
                                    // (U) 新管道在上
                                    {
                                        var maskId = CreateBackgroundMask(firstIntersection, BackgroundMask, tr, db, sampleInfo.PipeBodyTemplate.Layer);
                                        SetDrawOrderBetween(tr, db,
                                            entityBelow: crossingPipe.ObjectId, // 旧管道在下
                                            mask: maskId,
                                            entityAbove: newPipeId); // 新管道在上
                                        ed.WriteMessage($"\n已创建遮罩，新管道覆盖管道 [{oldPipeName}]。");
                                    }
                                    break;

                                case PipeCrossingDialog.CrossingAction.Under:
                                    // (D) 新管道在下
                                    {
                                        var maskId = CreateBackgroundMask(firstIntersection, BackgroundMask, tr, db, sampleInfo.PipeBodyTemplate.Layer);
                                        SetDrawOrderBetween(tr, db,
                                            entityBelow: newPipeId, // 新管道在下
                                            mask: maskId,
                                            entityAbove: crossingPipe.ObjectId); // 旧管道在上
                                        ed.WriteMessage($"\n已创建遮罩，新管道位于管道 [{oldPipeName}] 下方。");
                                    }
                                    break;

                                default:
                                    // 取消 or 关闭窗口 → 可直接删除已插入的新管道块
                                    newBr.Erase();
                                    tr.Commit();
                                    ed.WriteMessage("\n操作已取消。");
                                    return;
                            }
                        }
                    }

                    // 删除原始线段
                    foreach (var seg in lineSegments)
                    {
                        var ent = tr.GetObject(seg.Id, OpenMode.ForWrite) as Entity;// 删除原始线段
                        if (ent != null)
                            ent.Erase();// 删除原始线段
                    }
                    Application.SetSystemVariable("WIPEOUTFRAME", 0);//设置屏蔽罩的边框为0，关闭；
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
            /// <param name="initialAttributes"> 属性窗口中显示的属性字段 </param>
            public PipeAttributeEditorForm(Dictionary<string, string> initialAttributes)
            {
                // 初始化属性表
                _attributes = new Dictionary<string, string>(initialAttributes ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase);
                InitializeComponent(); // 初始化控件
                FillAttributesToGrid();// 填充属性表到网格
            }

            /// <summary>
            /// 初始化控件
            /// </summary>
            private void InitializeComponent()
            {
                this.Text = "示例管道属性编辑"; // 设置窗体标题
                this.FormBorderStyle = FormBorderStyle.FixedDialog; // 设置窗体为固定对话框
                this.StartPosition = FormStartPosition.CenterParent; // 设置窗体启动位置为父窗体中心
                this.ClientSize = new System.Drawing.Size(640, 800); // 设置窗体大小
                this.MaximizeBox = false;   // 禁用最大化按钮
                this.MinimizeBox = false; // 禁用最小化按钮
                this.MinimizeBox = false;   // 禁用最小化按钮
                this.ShowInTaskbar = false; // 不在任务栏显示
                this.AutoScaleMode = AutoScaleMode.Font; // 设置自动缩放模式为字体
                //初始化属性表 DataGridView
                _dataGridView = new DataGridView
                {
                    Dock = DockStyle.Fill, // 填充整个窗体
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, // 列宽自适应填充
                    AllowUserToAddRows = false, // 禁止用户添加行
                    AllowUserToDeleteRows = false, // 禁止用户删除行
                    RowHeadersVisible = false, // 隐藏行头
                    SelectionMode = DataGridViewSelectionMode.CellSelect, // 设置选择模式为单元格选择
                    MultiSelect = false // 禁止多选
                };
                // 添加列：Key（字段）和 Value（值）
                var colKey = new DataGridViewTextBoxColumn { Name = "Key", HeaderText = "字段", ReadOnly = true };
                var colVal = new DataGridViewTextBoxColumn { Name = "Value", HeaderText = "值", ReadOnly = false };
                // 添加Key列到 DataGridView
                _dataGridView.Columns.Add(colKey);
                _dataGridView.Columns.Add(colVal); // 添加Value列到 DataGridView
                // 初始化"完成"按钮
                _btnOk = new Button { Text = "完成", DialogResult = DialogResult.OK, Width = 90, Height = 30 };
                _btnCancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Width = 90, Height = 30 }; // 初始化"取消"按钮
                // 绑定按钮点击事件
                _btnOk.Click += BtnOk_Click;
                // 绑定取消按钮点击事件，关闭窗体
                _btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };
                // 底部按钮面板，右对齐
                var panel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom, // 设置面板停靠在底部
                    Height = 50, // 设置面板高度
                    FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft, // 设置流向为从右到左
                    Padding = new Padding(8), // 设置内边距
                    WrapContents = false // 禁止换行
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
            private void FillAttributesToGrid()
            {
                _dataGridView.Rows.Clear();// 清除现有行
                //foreach (var kv in _attributes.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase)) // 循环填充属性表到网格 按键排序以保证稳定性
                //{
                //    _dataGridView.Rows.Add(kv.Key, kv.Value);// 添加行
                //}
                foreach (var kv in _attributes) // 循环填充属性表到网格 按键排序以保证稳定性
                {
                    _dataGridView.Rows.Add(kv.Key, kv.Value);// 添加行
                }
                if (_dataGridView.Rows.Count > 0)// 如果有行，选中第一行的值列
                    _dataGridView.CurrentCell = _dataGridView.Rows[0].Cells[1]; // 选中第一行的值列
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
        /// 按统一列宽后的最终表格高度重新排列所有独立设备表，并返回下一张表的起始位置。
        /// </summary>
        private Point3d ReflowDeviceTables(Database db, List<ObjectId> tableIds, Point3d startPosition, double verticalGap)
        {
            if (db == null || tableIds == null || tableIds.Count == 0)
                return startPosition;

            Point3d nextPosition = startPosition;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    foreach (ObjectId tableId in tableIds.Where(id => id != ObjectId.Null))
                    {
                        Table table = trans.GetObject(tableId, OpenMode.ForWrite) as Table;
                        if (table == null) continue;

                        // 统一列宽后重新生成布局，再读取最终行高，避免表线和下一张表位置仍使用旧尺寸。
                        table.Position = nextPosition;
                        table.GenerateLayout();
                        double tableHeight = GetTableHeight(table);
                        LogManager.Instance.LogInfo(
                            $"[设备表][最终排布] TableId={tableId}, Position={nextPosition}, Height={tableHeight}, Width={table.Width}");
                        nextPosition = new Point3d(
                            startPosition.X,
                            nextPosition.Y - tableHeight - verticalGap,
                            startPosition.Z);
                    }

                    trans.Commit();
                }
                catch
                {
                    trans.Abort();
                    throw;
                }
            }

            return nextPosition;
        }

        /// <summary>
        /// 汇总表格所有行的最终高度。
        /// </summary>
        private static double GetTableHeight(Table table)
        {
            if (table == null) return 0.0;

            double height = 0.0;
            for (int row = 0; row < table.Rows.Count; row++)
            {
                height += table.Rows[row].Height;
            }
            return height;
        }

        #endregion

        #region 绘制管道线

        /// <summary>
        /// 新增：通过点击采集点并生成管道块（入口/出口两种命令）DrawOutletPipeByClicks
        /// </summary>
        [CommandMethod(nameof(DrawOutletPipeByClicks))]
        public void DrawOutletPipeByClicks()
        {
            DrawPipeByClicks(isOutlet: true);
        }

        /// <summary>
        /// 新增：通过点击采集点并生成管道块（入口/出口两种命令）DrawInletPipeByClicks
        /// </summary>
        [CommandMethod(nameof(DrawInletPipeByClicks))]
        public void DrawInletPipeByClicks()
        {
            DrawPipeByClicks(isOutlet: false);
        }

        /// <summary>
        /// 绘制管道方法   完成版——手选示例块 + 端点相交继承参数 + 同名标签同步（增强：读取到0不覆盖，字段标准化）
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
                    if (string.IsNullOrWhiteSpace(kv.Key)) continue;// 空键跳过
                    string val = (kv.Value ?? string.Empty).Trim();// 空值跳过
                    if (IsZeroLikeForPipe(val)) continue;
                    // 检查是否包含任一候选键
                    foreach (var k in keys)
                    {
                        if (kv.Key.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0) // 包含匹配
                            return val;
                    }
                }

                // 未命中返回空
                return string.Empty;
            }

            // 本地函数——从多个候选中取第一个非空且非0值
            string PickFirstNonZero(params string[] candidates)
            {
                foreach (var c in candidates) // 循环
                {
                    string v = (c ?? string.Empty).Trim();
                    if (IsZeroLikeForPipe(v)) continue;
                    return v;
                }
                return string.Empty;
            }

            try
            {
                // ================== 1. 选择示例块，采集路径点 ==================
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
                        if (sampleBr != null && !sampleBr.IsWriteEnabled) sampleBr.UpgradeOpen(); // 如果是只读，则升级为可写

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
                            var sampleAttrMap = GetEntityAttributeMap(tr, sampleBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);// 拿到样例块属性用于兜底
                            foreach (var kv in sampleAttrMap)// 循环样例块属性
                            {
                                if (string.IsNullOrWhiteSpace(kv.Key)) continue; // 空键跳过
                                loadedAttrs[kv.Key.Trim()] = kv.Value ?? string.Empty; // 空值也要保留，避免后续编辑窗口中缺失
                            }
                            ed.WriteMessage($"\n历史属性为空，已用样例块属性兜底初始化 {loadedAttrs.Count} 项。");
                        }

                        // 确保属性窗口包含必须的 canonical 字段（PIPELINETITLE, START_POINT, END_POINT, TAG_NO）
                        var requiredCanonicals = new[] { "PIPELINETITLE", "START_POINT", "END_POINT", "TAG_NO" };
                        var sampleAttrMap2 = GetEntityAttributeMap(tr, sampleBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);// 拿到样例块属性用于兜底
                        foreach (var c in requiredCanonicals)//循环必须字段
                        {
                            // 在历史属性中查询是否已必须字段，如存在则跳过
                            bool exists = TryFindFirstValue(loadedAttrs, c, out var _);
                            if (exists) continue;

                            // 尝试从样例块属性中按 canonical 查找一个候选值
                            if (TryFindFirstValue(sampleAttrMap2, c, out var fromSample) && !string.IsNullOrWhiteSpace(fromSample))
                            {
                                loadedAttrs[c] = fromSample.Trim();
                            }
                            else
                            {
                                // 插入空键以便在属性编辑窗口中显示
                                loadedAttrs[c] = string.Empty;
                            }
                        }

                        // 弹出属性编辑窗口
                        Dictionary<string, string> editedAttrsFromEditor = null;
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
                            // 保存为最后使用的属性（缓存）
                            FileManager.SaveLastPipeAttributes(isOutlet, editedAttrs);
                            // 保留一份用于后续合并到最新样例属性
                            editedAttrsFromEditor = new Dictionary<string, string>(editedAttrs, StringComparer.OrdinalIgnoreCase);

                            // 回写示例块已有属性
                            try
                            {
                                var sampleBrWrite = tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference; // 升级为可写
                                if (sampleBrWrite != null) // 如果示例块不为空，则尝试回写属性
                                {
                                    foreach (ObjectId aid in sampleBrWrite.AttributeCollection) // 循环示例块的属性集合
                                    {
                                        var ar = tr.GetObject(aid, OpenMode.ForWrite) as AttributeReference; // 把示例块中的当前属性升级为可写
                                        if (ar == null || string.IsNullOrWhiteSpace(ar.Tag)) continue; // 如果属性为空或标签为空，则跳过
                                        if (editedAttrs.TryGetValue(ar.Tag, out var newVal)) // 如果编辑后的属性中存在当前标签，则获取新值
                                        {
                                            ar.TextString = newVal ?? string.Empty; // 回写新值（空值也要回写，避免后续编辑窗口中缺失）
                                            try { ar.AdjustAlignment(db); } catch { } // 调整对齐，忽略异常
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
                        firstOpts.AllowNone = true; // 允许回车取消
                        firstOpts.Keywords.Add("取消"); // 添加取消关键字
                        firstOpts.AppendKeywordsToMessage = true; // 显示关键字提示
                        var firstRes = ed.GetPoint(firstOpts); // 获取起点输入

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
                            nextOpts.UseBasePoint = true; // 使用基点
                            nextOpts.BasePoint = points.Last(); // 基点为上一个点
                            nextOpts.AllowNone = true; // 允许回车结束
                            nextOpts.Keywords.Add("完成"); // 添加完成关键字
                            nextOpts.AppendKeywordsToMessage = true; // 显示关键字提示
                            // 获取下一个点输入
                            var nextRes = ed.GetPoint(nextOpts);
                            // 处理输入状态
                            if (nextRes.Status == PromptStatus.OK)
                            {
                                var pt = nextRes.Value; // 获取输入点
                                if (pt.IsEqualTo(points.Last())) break; // 如果与上一个点相同，则结束
                                points.Add(pt); // 加入新点
                                continue;
                            }
                            // 处理取消或完成状态
                            if (nextRes.Status == PromptStatus.None) break; // 回车结束
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
                                var hostBtr = tr.GetObject(sampleBr.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord; // 获取样例块的块定义
                                if (hostBtr != null) // 如果块定义不为空，则尝试分析嵌套块
                                {
                                    foreach (ObjectId id in hostBtr) // 循环块定义中的所有对象
                                    {
                                        var nested = tr.GetObject(id, OpenMode.ForRead) as BlockReference; // 尝试获取嵌套块参照
                                        if (nested == null) continue; // 如果嵌套块参照为空，则跳过
                                        // 分析嵌套块
                                        var nestedInfo = AnalyzeSampleBlock(tr, nested);
                                        if (nestedInfo != null && nestedInfo.PipeBodyTemplate != null) // 如果嵌套块分析成功且有主体，则使用嵌套块信息
                                        {
                                            sampleInfo = nestedInfo; // 替换样例信息为嵌套块信息
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
                        var (midPoint, _) = ComputeMidPointAndAngle(points, pipelineLength);// 计算中点和角度

                        // 构建局部管线
                        var pipeLocal = BuildPipePolylineLocal(sampleInfo.PipeBodyTemplate, points, midPoint);

                        // 读取样例属性作为基线
                        var latestSampleAttrs = GetEntityAttributeMap(tr, sampleBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                        var wpfScale = AutoCadHelper.GetScale();

                        // 端点命中容差（按比例自适应）
                        double hitTol = Math.Max(1.0, wpfScale * 0.2);

                        // 查找起点命中图元（块参照），用于继承属性
                        var startCandidates = FindBlocksCrossingPoint(tr, points.First(), sampleBr.ObjectId, hitTol);
                        // 查找终点命中图元（块参照），用于继承属性
                        var endCandidates = FindBlocksCrossingPoint(tr, points.Last(), sampleBr.ObjectId, hitTol);// 允许起点和终点命中同一个块参照

                        // 取最优候选
                        var startBr = startCandidates.FirstOrDefault();
                        var endBr = endCandidates.FirstOrDefault();

                        // 读取起点的命中图元属性
                        var startMap = startBr != null
                            ? (GetEntityAttributeMap(tr, startBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
                            : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        // 读取终点的命中图元属性
                        var endMap = endBr != null
                            ? (GetEntityAttributeMap(tr, endBr) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
                            : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                        // 日志
                        ed.WriteMessage($"\n端点命中容差={hitTol:F2}；起点命中={(startBr != null ? startBr.Name : "无")}；终点命中={(endBr != null ? endBr.Name : "无")}。");

                        // 确定标题
                        string pipeTitle = latestSampleAttrs.TryGetValue("PIPELINETITLE", out var sample_Title) && !string.IsNullOrWhiteSpace(sample_Title) && !IsZeroLikeForPipe(sample_Title)
                                         ? sample_Title
                                         : sampleBr.Name ?? "管道"; // 获取样例块的名称作为标题兜底

                        // 创建箭头与标题实体
                        var arrowEntities = CreateDirectionalArrowsAndTitles(tr, sampleInfo, points, midPoint, pipeTitle, sampleBr.Name);

                        // 克隆模板属性定义
                        var attDefsLocal = CloneAttributeDefinitionsLocal(sampleInfo.AttributeDefinitions, midPoint, 0.0, pipelineLength, sampleBr.Name)
                                           ?? new List<AttributeDefinition>();

                        // 按模板定义回填值（值为0不回填覆盖）
                        if (sampleInfo.AttributeDefinitions != null && sampleInfo.AttributeDefinitions.Count > 0)
                        {
                            foreach (var def in attDefsLocal)// 循环属性字段
                            {
                                if (string.IsNullOrWhiteSpace(def.Tag)) continue; // 空键跳过
                                if (latestSampleAttrs.TryGetValue(def.Tag, out var v)) // 尝试从最新样例属性中获取值
                                {
                                    string vv = (v ?? string.Empty).Trim(); // 空值跳过
                                    if (!IsZeroLikeForPipe(vv)) // 仅当值非0时才回填
                                        def.TextString = vv; // 
                                }
                                def.Invisible = false;
                                def.Constant = false;
                            }
                        }
                        else
                        {
                            // 模板无属性定义则动态生成（值为0不生成）
                            attDefsLocal.Clear();// 清空模板属性定义
                            double attHeight = 2.5; // 默认属性高度
                            double yOffsetBase = -attHeight * 2.0; // 默认起始偏移
                            int idx = 0; // 索引计数器
                            // 按键名排序后生成属性定义
                            foreach (var kv in latestSampleAttrs)
                            {
                                if (string.IsNullOrWhiteSpace(kv.Key)) continue; // 空键跳过
                                string vv = (kv.Value ?? string.Empty).Trim(); // 空值跳过
                                if (IsZeroLikeForPipe(vv)) continue; // 值为0不生成
                                // 生成属性定义
                                attDefsLocal.Add(new AttributeDefinition
                                {
                                    Tag = kv.Key, // 标签
                                    Position = new Point3d(0, yOffsetBase - idx * attHeight * 1.2, 0),// 位置
                                    Rotation = 0.0, // 旋转角度
                                    TextString = vv, // 文本内容
                                    Height = attHeight,// 高度
                                    Invisible = false,// 可见
                                    Constant = false// 非常量
                                });
                                idx++; // 索引递增
                            }
                        }

                        // 最终管段号（继承值为0时回退自动编号）
                        int nextSegNum = GetNextPipeSegmentNumber(db);
                        string extractedPipeNo = nextSegNum.ToString("D4"); // 默认编号为 4 位数
                        if (latestSampleAttrs.TryGetValue("TAG_NO", out var pn)) // 尝试从最新样例属性中获取管段号
                        {
                            string pnv = (pn ?? string.Empty).Trim(); // 空值跳过
                            if (!IsZeroLikeForPipe(pnv) && !string.IsNullOrWhiteSpace(pnv)) // 仅当值非0且非空时才使用
                                extractedPipeNo = pnv; // 从最新样例属性中获取管段号
                        }

                        // 计算附加属性布局参数
                        double finalAttHeight = attDefsLocal.Count > 0 ? attDefsLocal[0].Height : 2.5;// 默认属性文字高度
                        double finalYOffsetBase = attDefsLocal.Count > 0 ? attDefsLocal[0].Position.Y - finalAttHeight * 1.2 : -finalAttHeight * 2.0; // 默认起始偏移
                        int extraIndex = 0; // 附加属性索引计数器

                        // 长度字段保持原逻辑
                        SetOrAddLengthAttrs(attDefsLocal, pipelineLength, ref extraIndex, finalYOffsetBase, finalAttHeight);// 设置长度属性字段

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
                        var newPipeId = InsertPipeBlockWithAttributes(tr, midPoint, newBlockName, 0.0, attValues);
                        var newBr = tr.GetObject(newPipeId, OpenMode.ForWrite) as BlockReference;

                        // ================== 4. 交互处理 (WPF 多交叉点，带调试) ==================
                        if (newPipeId != ObjectId.Null)
                        {
                            var newPipeBr = tr.GetObject(newPipeId, OpenMode.ForWrite) as BlockReference;
                            if (newPipeBr != null)
                            {
                                var newPath = GetPipePathVertices(newPipeBr, tr);
                                if (newPath != null && newPath.Count >= 2)
                                {
                                    var existingPipes = GetExistingPipes(db, tr);
                                    ed.WriteMessage($"\n[调试] 已有管道数量: {existingPipes.Count}");

                                    // 计算遮罩尺寸（避免双重缩放）
                                    wpfScale = AutoCadHelper.GetScale();
                                    double maskFinalSize = 0;
                                    if (sampleInfo.PipeBodyTemplate != null &&
                                        sampleInfo.PipeBodyTemplate.ConstantWidth > 0)
                                    {
                                        maskFinalSize = Math.Max(1.0, sampleInfo.PipeBodyTemplate.ConstantWidth * wpfScale * 10);
                                    }
                                    else
                                    {
                                        maskFinalSize = Math.Max(1.0, 5.0 * wpfScale);   // 5 为基础尺寸，可调
                                    }
                                    ed.WriteMessage($"\n[调试] 遮罩边长 (世界单位): {maskFinalSize:F2}, 比例: {wpfScale:F2}");

                                    var crossOptions = new List<PipeCrossingMultiDialogWpf.CrossingOptionViewModel>();
                                    double endpointTol = Math.Max(1.0, wpfScale * 0.3);
                                    // 循环
                                    foreach (var oldBr in existingPipes)
                                    {
                                        if (oldBr.ObjectId == newPipeBr.ObjectId) continue;
                                        var oldPath = GetPipePathVertices(oldBr, tr);
                                        if (oldPath == null || oldPath.Count < 2) continue;
                                        //插入屏蔽块
                                        var intersections = GetPathIntersections(newPath, oldPath);
                                        if (intersections.Count == 0) continue;

                                        ed.WriteMessage($"\n[调试] 与管道 '{oldBr.Name}' 交叉点总数: {intersections.Count}");

                                        var validIntersections = intersections.Where(p =>
                                            !IsNearEndpoint(p, newPath, endpointTol) &&
                                            !IsNearEndpoint(p, oldPath, endpointTol)
                                        ).ToList();

                                        ed.WriteMessage($"  其中非端点交叉点数量: {validIntersections.Count}");

                                        // 同一对管道在 PL 顶点或相邻线段处可能返回重复交点。
                                        // 这里只按当前新旧管道去重，不合并不同管道之间的同坐标交叉关系。
                                        double intersectionTol = Math.Max(1e-6, wpfScale * 0.001);
                                        var uniqueIntersections = new List<Point3d>();
                                        foreach (var intersection in validIntersections)
                                        {
                                            if (!uniqueIntersections.Any(existing =>
                                                    existing.DistanceTo(intersection) <= intersectionTol))
                                            {
                                                uniqueIntersections.Add(intersection);
                                            }
                                        }

                                        ed.WriteMessage($"  去重后的交叉点数量: {uniqueIntersections.Count}");

                                        foreach (var pt in uniqueIntersections)
                                        {
                                            string oldName = "未知";
                                            var attrs = GetEntityAttributeMap(tr, oldBr);
                                            if (attrs.TryGetValue("名称", out var nm)) oldName = nm;
                                            else oldName = oldBr.Name ?? "未知";

                                            crossOptions.Add(new PipeCrossingMultiDialogWpf.CrossingOptionViewModel
                                            {
                                                Intersection = pt,
                                                PipeName = oldName,
                                                OldPipeId = oldBr.ObjectId,
                                                NewPipeId = newPipeId
                                            });
                                        }
                                    }

                                    ed.WriteMessage($"\n[调试] 最终收集到的交叉点数量: {crossOptions.Count}");

                                    if (crossOptions.Count > 0)
                                    {
                                        var dlg = new PipeCrossingMultiDialogWpf(crossOptions);
                                        bool? result = dlg.ShowDialogWithOwner();
                                        ed.WriteMessage($"\n[调试] 对话框结果: {result}, IsCancelled: {dlg.IsCancelled}");

                                        if (result == false || dlg.IsCancelled)
                                        {
                                            newPipeBr.Erase(true);
                                            tr.Commit();
                                            ed.WriteMessage("\n操作已取消。");
                                            return;
                                        }
                                        else
                                        {
                                            foreach (var opt in dlg.Options)
                                            {
                                                if (opt.SelectedAction == PipeCrossingMultiDialogWpf.CrossingAction.Connect)
                                                {
                                                    ed.WriteMessage($"\n[调试] 交叉点 {opt.IntersectionDisplay} 选择相连，跳过遮罩。");
                                                    continue;
                                                }

                                                ObjectId entityBelow, entityAbove;
                                                bool newPipeIsAbove;
                                                if (opt.SelectedAction == PipeCrossingMultiDialogWpf.CrossingAction.Cover)
                                                {
                                                    entityBelow = opt.OldPipeId;
                                                    entityAbove = opt.NewPipeId;
                                                    newPipeIsAbove = true;
                                                    ed.WriteMessage($"\n[调试] 处理方式: 覆盖, 旧管道在下, 新管道在上");
                                                }
                                                else
                                                {
                                                    entityBelow = opt.NewPipeId;
                                                    entityAbove = opt.OldPipeId;
                                                    newPipeIsAbove = false;
                                                    ed.WriteMessage($"\n[调试] 处理方式: 下方, 新管道在下, 旧管道在上");
                                                }

                                                ed.WriteMessage($"\n[调试] 开始创建遮罩，中心点: {opt.IntersectionDisplay}, 尺寸: {maskFinalSize:F2}");

                                                var maskId = CreateBackgroundMask(
                                                    opt.Intersection,
                                                    maskFinalSize,
                                                    tr,
                                                    db,
                                                    sampleInfo.PipeBodyTemplate.Layer ?? "0"
                                                );

                                                ed.WriteMessage($"\n[调试] 遮罩创建完成，ObjectId: {maskId}");

                                                SetNewPipeDrawOrderWithoutMovingExistingPipe(
                                                    tr,
                                                    db,
                                                    opt.NewPipeId,
                                                    opt.OldPipeId,
                                                    maskId,
                                                    newPipeIsAbove);
                                                ed.WriteMessage($"\n[调试] 绘图次序调整完毕，顺序：{entityBelow} -> {maskId} -> {entityAbove}");
                                            }
                                            ed.WriteMessage("\n[完成] 所有交叉点处理完毕。");
                                        }
                                    }
                                    else
                                    {
                                        ed.WriteMessage("\n[调试] 没有需要处理的交叉点（可能都在端点）。");
                                    }
                                }
                                else
                                {
                                    ed.WriteMessage("\n[调试] 新管道路径提取失败。");
                                }
                            }
                            else
                            {
                                ed.WriteMessage("\n[调试] 新管道块参照无效。");
                            }
                        }
                        else
                        {
                            ed.WriteMessage("\n[调试] 新管道插入失败 (newPipeId 为空)。");
                        }

                        // 下面的“删除原始线段”等代码保持不变

                        // 后处理拓扑关系
                        PipelineTopologyHelper.PostProcessAfterPipePlaced(tr, newPipeId, isOutlet ? "Outlet" : "Inlet");

                        // 继承图层

                        if (newBr != null) newBr.Layer = sampleInfo.PipeBodyTemplate.Layer;

                        // 保存最新的属性到上次属性缓存（包含 PIPELINETITLE, TAG_NO, START_POINT, END_POINT）
                        try
                        {
                            FileManager.SaveLastPipeAttributes(isOutlet, latestSampleAttrs); // 保存最新的属性到上次属性缓存（包含 PIPELINETITLE, TAG_NO, START_POINT, END_POINT）
                        }
                        catch { }

                        // 兼容分支（当前为手选样例，不会进入）
                        if (tempInserted)
                        {
                            try { (tr.GetObject(sampleBr.ObjectId, OpenMode.ForWrite) as BlockReference)?.Erase(); } catch { }
                        }
                        Application.SetSystemVariable("WIPEOUTFRAME", 0);//设置屏蔽罩的边框为0，关闭；
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
         
        /// <summary>
        /// 管道交叉处理选择对话框（三个操作按钮）
        /// </summary>
        public class PipeCrossingDialog : Form
        {
            public enum CrossingAction
            {
                Connect,    // 交叉相连
                Cover,      // 不相连且覆盖
                Under,      // 不相连且在下方
                Cancel
            }

            public CrossingAction SelectedAction { get; private set; } = CrossingAction.Cancel;

            public PipeCrossingDialog(string existingPipeName)
            {
                this.Text = "管道交叉检测";
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.StartPosition = FormStartPosition.CenterParent;
                this.ClientSize = new System.Drawing.Size(450, 180);
                this.MaximizeBox = false;
                this.MinimizeBox = false;

                var lbl = new Label
                {
                    Text = $"检测到与管道 [{existingPipeName}] 有交叉。请选择处理方式：",
                    Location = new System.Drawing.Point(12, 20),
                    AutoSize = true
                };

                var btnConnect = new Button { Text = "交叉相连 (&Y)", Location = new System.Drawing.Point(30, 70), Width = 110 };
                var btnCover = new Button { Text = "不相连且覆盖 (&U)", Location = new System.Drawing.Point(155, 70), Width = 130 };
                var btnUnder = new Button { Text = "不相连且在下方 (&D)", Location = new System.Drawing.Point(300, 70), Width = 130 };
                var btnCancel = new Button { Text = "取消", Location = new System.Drawing.Point(180, 120), Width = 80 };

                btnConnect.Click += (s, e) => { SelectedAction = CrossingAction.Connect; this.Close(); };
                btnCover.Click += (s, e) => { SelectedAction = CrossingAction.Cover; this.Close(); };
                btnUnder.Click += (s, e) => { SelectedAction = CrossingAction.Under; this.Close(); };
                btnCancel.Click += (s, e) => { SelectedAction = CrossingAction.Cancel; this.Close(); };

                this.Controls.Add(lbl);
                this.Controls.Add(btnConnect);
                this.Controls.Add(btnCover);
                this.Controls.Add(btnUnder);
                this.Controls.Add(btnCancel);

                this.AcceptButton = btnConnect;
                this.CancelButton = btnCancel;
            }
        }
        /// <summary>
        /// 获取现有管道
        /// </summary>
        /// <param name="db"></param>
        /// <param name="tr"></param>
        /// <returns></returns>
        private List<BlockReference> GetExistingPipes(Database db, Transaction tr)
        {
            var pipes = new List<BlockReference>();
            var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            var btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
            foreach (ObjectId id in btr)
            {
                var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                if (br != null && IsPipeBlock(br))
                    pipes.Add(br);
            }
            return pipes;
        }

        /// <summary>
        /// 从管道块参照中提取多段线的世界坐标顶点列表
        /// </summary>
        private List<Point3d>? GetPipePathVertices(BlockReference br, Transaction tr)
        {
            var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            if (btr == null) return null;
            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent is Polyline poly)
                {
                    var pts = new List<Point3d>();
                    for (int i = 0; i < poly.NumberOfVertices; i++)
                        pts.Add(poly.GetPoint3dAt(i));
                    // 应用块变换
                    return pts.Select(p => p.TransformBy(br.BlockTransform)).ToList();
                }
            }
            // 降级：使用包围盒近似
            try
            {
                var ext = br.GeometricExtents;
                return new List<Point3d> { ext.MinPoint, ext.MaxPoint };
            }
            catch { return null; }
        }

        /// <summary>
        /// 计算两条路径（由顶点序列定义）的所有近似交点
        /// </summary>
        private List<Point3d> GetPathIntersections(List<Point3d> pathA, List<Point3d> pathB)
        {
            var intersections = new List<Point3d>();
            for (int i = 0; i < pathA.Count - 1; i++)
            {
                var segA = new LineSegment3d(pathA[i], pathA[i + 1]);
                for (int j = 0; j < pathB.Count - 1; j++)
                {
                    var segB = new LineSegment3d(pathB[j], pathB[j + 1]);
                    var pts = segA.IntersectWith(segB); // 返回 Point3d[]
                    if (pts != null)
                    {
                        foreach (Point3d p in pts)
                            intersections.Add(p);
                    }
                }
            }
            return intersections;
        }
        
        /// <summary>
        /// 在指定中心点创建一个正方形 Wipeout 遮罩（自动匹配背景色，且隐藏边框）
        /// </summary>
        /// <param name="center">遮罩中心点（世界坐标）</param>
        /// <param name="size">遮罩边长（世界单位，已含比例）</param>
        /// <param name="tr">当前事务</param>
        /// <param name="db">数据库</param>
        /// <param name="layer">目标图层</param>
        public ObjectId CreateBackgroundMask(Point3d center, double size, Transaction tr, Database db, string layer = "0")
        {
            // 计算矩形四个角点（注意：必须闭合，即首尾点重合）
            double half = size / 2.0;
            Point2dCollection pts = new Point2dCollection(5)
            {
                new Point2d(center.X - half, center.Y - half),
                new Point2d(center.X + half, center.Y - half),
                new Point2d(center.X + half, center.Y + half),
                new Point2d(center.X - half, center.Y + half),
                new Point2d(center.X - half, center.Y - half)  // 闭合
            };

            Wipeout wipe = new Wipeout();
            wipe.SetDatabaseDefaults(db);
            // 设置四边形节点，需要提供法向量（通常为 Z 轴）
            wipe.SetFrom(pts, new Vector3d(0, 0, 1));

            // 设置图层（回退到 "0" 层）
            wipe.Layer = string.IsNullOrWhiteSpace(layer) ? "0" : layer;

            // 临时关闭 Wipeout 边框显示，以保持视觉干净
            object oldWipeoutFrame = Application.GetSystemVariable("WIPEOUTFRAME");
            Application.SetSystemVariable("WIPEOUTFRAME", 0);

            ObjectId wipeId = ObjectId.Null;
            try
            {
                BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                btr.AppendEntity(wipe);
                tr.AddNewlyCreatedDBObject(wipe, true);
                wipeId = wipe.ObjectId;
            }
            finally
            {
                // 恢复系统变量原值
                Application.SetSystemVariable("WIPEOUTFRAME", oldWipeoutFrame);
            }

            return wipeId;
        }


        /// <summary>
        /// 调整三个实体的绘图次序：从下到上依次为 entityBelow -> mask -> entityAbove
        /// </summary>
        /// <param name="tr">事务</param>
        /// <param name="db">数据库</param>
        /// <param name="entityBelow">最下层实体 ID</param>
        /// <param name="mask">中间遮罩 ID</param>
        /// <param name="entityAbove">最上层实体 ID</param>
        public void SetDrawOrderBetween(Transaction tr, Database db,
            ObjectId entityBelow, ObjectId mask, ObjectId entityAbove)
        {
            if (entityBelow.IsNull || mask.IsNull || entityAbove.IsNull) return;

            // 获取模型空间
            var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            var msId = bt[BlockTableRecord.ModelSpace];
            var ms = tr.GetObject(msId, OpenMode.ForRead) as BlockTableRecord;

            // DrawOrderTableId 属性会自动创建 DrawOrderTable（若尚不存在）
            ObjectId dotId = ms.DrawOrderTableId;
            if (dotId.IsNull) return;

            var dot = tr.GetObject(dotId, OpenMode.ForWrite) as DrawOrderTable;
            if (dot == null) return;

            // 只调整当前交叉点涉及的三个对象，不再把整条下方管道移动到模型空间最底部。
            // 这样处理 A、B、C 多条管道时，后一个交叉点不会直接破坏前一个交叉点的局部关系。
            dot.MoveAbove(new ObjectIdCollection { mask }, entityBelow);
            dot.MoveAbove(new ObjectIdCollection { entityAbove }, mask);

            db.TransactionManager.QueueForGraphicsFlush();
        }

        /// <summary>
        /// 设置新管道与既有管道之间的局部绘图次序，但不移动既有管道。
        /// </summary>
        /// <param name="tr">当前事务</param>
        /// <param name="db">当前数据库</param>
        /// <param name="newPipeId">新管道块参照 ID</param>
        /// <param name="existingPipeId">既有管道块参照 ID，仅作为定位参照</param>
        /// <param name="maskId">当前交叉点的 Wipeout ID</param>
        /// <param name="newPipeIsAbove">新管道是否位于既有管道上方</param>
        private void SetNewPipeDrawOrderWithoutMovingExistingPipe(
            Transaction tr,
            Database db,
            ObjectId newPipeId,
            ObjectId existingPipeId,
            ObjectId maskId,
            bool newPipeIsAbove)
        {
            // 任意一个对象无效时，不执行绘图次序调整。
            if (newPipeId.IsNull || existingPipeId.IsNull || maskId.IsNull) return;

            // 获取模型空间的绘图次序表。
            var blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            var modelSpaceId = blockTable[BlockTableRecord.ModelSpace];
            var modelSpace = tr.GetObject(modelSpaceId, OpenMode.ForRead) as BlockTableRecord;
            var drawOrderTableId = modelSpace.DrawOrderTableId;
            if (drawOrderTableId.IsNull) return;

            // 以写入方式打开绘图次序表。
            var drawOrderTable = tr.GetObject(drawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
            if (drawOrderTable == null) return;

            if (newPipeIsAbove)
            {
                // 覆盖关系：既有管道 -> 遮罩 -> 新管道。
                // 这里只移动遮罩和新管道，不移动既有管道。
                drawOrderTable.MoveAbove(new ObjectIdCollection { maskId }, existingPipeId);
                drawOrderTable.MoveBelow(new ObjectIdCollection { maskId }, newPipeId);
                drawOrderTable.MoveAbove(new ObjectIdCollection { newPipeId }, maskId);
            }
            else
            {
                // 下方关系：新管道 -> 遮罩 -> 既有管道。
                // 既有管道仍然只作为遮罩定位参照，不改变其与其他对象的既有关系。
                drawOrderTable.MoveAbove(new ObjectIdCollection { maskId }, newPipeId);
                drawOrderTable.MoveBelow(new ObjectIdCollection { maskId }, existingPipeId);
                drawOrderTable.MoveBelow(new ObjectIdCollection { newPipeId }, maskId);
            }

            // 请求 AutoCAD 立即刷新图形显示。
            db.TransactionManager.QueueForGraphicsFlush();
        }

        private bool IsPipeBlock(BlockReference br)
        {
            string name = br.Name ?? "";
            return name.IndexOf("管道", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   name.IndexOf("PIPE", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// 判断指定点是否在路径的起点或终点附近（容差内）
        /// </summary>
        private bool IsNearEndpoint(Point3d point, List<Point3d> path, double tolerance)
        {
            if (path == null || path.Count < 2) return false;
            return point.DistanceTo(path.First()) <= tolerance ||
                   point.DistanceTo(path.Last()) <= tolerance;
        }

        /// <summary>
        /// 在属性字典中按照 canonicalKey 查找首个非空值（考虑所有 alias）
        /// </summary>
        public static bool TryFindFirstValue(IDictionary<string, string> rawMap, string canonicalKey, out string value)
        {
            value = string.Empty;
            if (rawMap == null || string.IsNullOrWhiteSpace(canonicalKey)) return false;

            // 1) 直接查找 canonicalKey 本身
            if (rawMap.TryGetValue(canonicalKey, out var v1) && !string.IsNullOrWhiteSpace(v1))
            {
                value = v1.Trim();
                return true;
            }

            // 2) 查找已知 alias
            var aliases = GetAliases(canonicalKey);
            foreach (var a in aliases)
            {
                if (rawMap.TryGetValue(a, out var v) && !string.IsNullOrWhiteSpace(v))
                {
                    value = v.Trim();
                    return true;
                }
            }

            // 3) 宽松匹配：键名包含匹配
            foreach (var kv in rawMap)
            {
                if (string.IsNullOrWhiteSpace(kv.Key)) continue;
                foreach (var a in aliases)
                {
                    if (kv.Key.IndexOf(a, StringComparison.OrdinalIgnoreCase) >= 0 && !string.IsNullOrWhiteSpace(kv.Value))
                    {
                        value = kv.Value.Trim();
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// 获取 canonical（首选英文）对应的别名列表（包含 canonical 本身）
        /// </summary>
        public static IReadOnlyList<string> GetAliases(string canonical)
        {
            if (string.IsNullOrWhiteSpace(canonical)) return Array.Empty<string>(); // 空输入返回空列表
            if (DictionaryHelper._canonicalToAliases.TryGetValue(canonical, out var list)) return list; // 找到对应的别名列表
            return new[] { canonical }; // 未找到则返回仅包含 canonical 本身的列表
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

                bool inside = false; // 标记点是否在包围盒内
                double dist = br.Position.DistanceTo(point);// 计算插入点距离

                // 优先判断点是否在包围盒内
                try
                {
                    var ext = br.GeometricExtents;// 获取包围盒
                    inside = IsPointInsideExtents(point, ext, tol); // 通过给出的参数判断选定位置是不是在包围盒内
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


        #endregion

        /// <summary>
        /// 辅助：把属性定义中存在或不存在的 Tag 设置/新增值
        /// </summary>
        private void SetOrAddAttr(List<AttributeDefinition> attDefs, string tag, string text, ref int extraIndex, double yOffsetBase, double attHeight)
        {
            // 参数保护
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
                    "是（Yes）. 表 -> 图元：将表格数据同步到选中的块\n" +
                    "否（No）. 图元 -> 表：将选中块的属性同步到表格\n" +
                    "取消（Cancel）. 双向同步：表格和块互相同步",
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
                // 获取用户选择的表格对象
                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;
                // 开始事务处理
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var table = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table; // 获取到表格对象
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
                    var headers = new List<string>(cols); // 存储列头
                    for (int c = 0; c < cols; c++) // 遍历每一列
                    {
                        string h2 = (table.Cells[2, c].TextString ?? string.Empty).Trim(); // 优先使用第二行作为列头
                        string h1 = (table.Cells[1, c].TextString ?? string.Empty).Trim(); // 备用使用第一行作为列头
                        string header = string.IsNullOrWhiteSpace(h2) ? h1 : h2; // 如果第二行为空，则使用第一行
                        header = header.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? header; // 取第一行文本作为列头
                        headers.Add(header); // 添加到 headers 列表
                    }

                    // 寻找标识列
                    int idCol = -1;
                    string[] idKeywords = new[] { "TAG_NO", "MODEL", "QTY", "DRAWINGNO.STANDARDNO", "NAME", "Pipeline", "Pipe No", "ID", "序号" };
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
                    var idKeywords = new[] { "TAG_NO", "PIPELINETITLE", "NAME", "MODEL", "DRAWINGNO.STANDARDNO", "START_POINT", "END_POINT", "DN", "PN", "WORK_TEMP", "WORK_PRESSURE", "MEDIUM", "HOT\\SOUND_ISOLACODE", "IS_ANTICORRO" };

                    foreach (var blockId in blockIds) // 遍历每个选中的块
                    {
                        try
                        {
                            var br = tr.GetObject(blockId, OpenMode.ForRead) as BlockReference; // 获取块参照对象
                            if (br == null) continue; // 如果不是块参照则跳过
                            var attrs = ExtractBlockAttributes(br) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // 提取块的属性字典

                            // 如果块没有明显的 key -> value，则尝试动态块属性
                            if ((attrs == null || attrs.Count == 0) && br.IsDynamicBlock)
                            {
                                try
                                {
                                    var dyn = br.DynamicBlockReferencePropertyCollection; // 获取动态块属性集合
                                    foreach (DynamicBlockReferenceProperty p in dyn) // 遍历每个动态属性
                                    {
                                        if (!string.IsNullOrWhiteSpace(p.PropertyName)) // 如果属性名不为空
                                            attrs[p.PropertyName] = p.Value?.ToString() ?? string.Empty; // 将属性名和值加入字典
                                    }
                                }
                                catch { }
                            }

                            blockDataList.Add(attrs); // 将该块的属性字典加入列表
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

            // 3) 仅使用完整的同义词匹配，避免不同业务字段因包含相同文字而串列。
            foreach (var kv in DictionaryHelper.AttributeSynonyms)
            {
                string synKey = kv.Key;
                if (string.IsNullOrWhiteSpace(synKey)) continue;
                string synNorm = synKey.ToLowerInvariant();
                synNorm = Regex.Replace(synNorm, @"[\s\[\]\(\){}]", "");
                synNorm = synNorm.Replace("℃", "c").Replace("°c", "c").Replace("°", "");

                if (string.Equals(normalized, synNorm, StringComparison.OrdinalIgnoreCase))
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
                        scaleDenom = AutoCadHelper.GetScale();
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


                // 管段号：第1列（索引1），跨2行 管道号
                table.MergeCells(Autodesk.AutoCAD.DatabaseServices.CellRange.Create(table, 1, 0, 2, 0));
                table.Cells[1, 0].TextString = "管段号\nPipeline\nNo.";

                // 管段起止点：第2-3列（索引2,3），row1 合并为组，row2 分别为 起点/终点
                table.MergeCells(Autodesk.AutoCAD.DatabaseServices.CellRange.Create(table, 1, 1, 1, 2));
                table.Cells[1, 1].TextString = "管段起止点\nPipeline From Start To End";
                table.Cells[2, 1].TextString = "起点\nFrom";
                table.Cells[2, 2].TextString = "终点\nTo";

                // 管道等级：第4列（索引4），跨2行
                table.MergeCells(Autodesk.AutoCAD.DatabaseServices.CellRange.Create(table, 1, 3, 2, 3));
                table.Cells[1, 3].TextString = "管道\n等级\nPipe Class";

                // 设计条件：第5-7列（索引5..7），row1 合并为组，row2: 介质名称/操作温度/操作压力
                table.MergeCells(Autodesk.AutoCAD.DatabaseServices.CellRange.Create(table, 1, 4, 1, 6));
                table.Cells[1, 4].TextString = "设计条件 \nDesign Condition";
                table.Cells[2, 4].TextString = "介质名称\nMedium Name";
                table.Cells[2, 5].TextString = "操作温度\nT(℃)";
                table.Cells[2, 6].TextString = "操作压力\nP(MPaG)";

                // 隔热及防腐：第8-9列（索引8..9），row1 合并为组，row2: Code / Antisepsis
                table.MergeCells(Autodesk.AutoCAD.DatabaseServices.CellRange.Create(table, 1, 7, 1, 8));
                table.Cells[1, 7].TextString = "隔热及防腐 \nInsul. & Antisepsis";
                table.Cells[2, 7].TextString = "隔热隔声代号\nCode";
                table.Cells[2, 8].TextString = "是否防腐\nAntisepsis";

                // 管道组：从第10列开始，动态宽度由 pipeGroupCount 决定
                int pipeGroupStart = 9;
                int pipeGroupEnd = pipeGroupStart + Math.Max(0, pipeGroupCount - 1);
                if (pipeGroupEnd >= table.Columns.Count) pipeGroupEnd = table.Columns.Count - 1;
                if (pipeGroupStart < table.Columns.Count)
                {
                    table.MergeCells(Autodesk.AutoCAD.DatabaseServices.CellRange.Create(table, 1, pipeGroupStart, 1, pipeGroupEnd));
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
        /// 根据视口尺度因子和比例字符串确定最终的比例分母
        /// </summary>
        /// <param name="table"></param>
        /// <param name="scaleDenominator"></param>
        private void ApplyScaledHeightsToTable(Autodesk.AutoCAD.DatabaseServices.Table table, double scaleDenominator)
        {
            // 将约定的基准字高按比例应用到表格的Title/Header/Data单元格（仅设置单元格 TextHeight，避免修改全局 TextStyle）
            if (table == null) return;

            double titleHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, scaleDenominator);
            double headerHeight = TextFontsStyleHelper.ComputeScaledHeight(2.8, scaleDenominator);
            double dataHeight = TextFontsStyleHelper.ComputeScaledHeight(2.5, scaleDenominator);

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
                        table.SetRowHeight(r, Math.Max(6.0, headerHeight * 2.8));
                    else
                        table.SetRowHeight(r, Math.Max(5.0, dataHeight * 2.2));
                }
                catch { }
            }
        }
             

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
                    int dataStartRow = 2;
                    if (table.Rows.Count <= dataStartRow)
                    {
                        ed.WriteMessage("\n表格行数过少，无法识别数据行，请确认表格结构。");
                        return;
                    }

                    // 查找标识列（用于匹配图元）
                    int idCol = -1;
                    string[] idKeywords = new[] { "管段号", "名称", "规格", "材料", "数量", "标准号", "管段等级", "介质名称", "温度", "压力等级", "隔热", "防腐" };
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
                        ed.WriteMessage("\n未能在表头中找到标识列（例如“名称”或“管段号”）。请在表格中包含用于匹配块的标识列后重试。");
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
                    if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                    {
                        ed.WriteMessage("\n未找到选择的文件。");
                        return;
                    }

                    // ====== 使用 NPOI 读取 Excel 内容 ======
                    List<string[]> excelData = new List<string[]>();
                    var mergedRanges = new List<(int r1, int c1, int r2, int c2)>(); // 存储合并区域(0-based)
                    int excelRows = 0, excelCols = 0;

                    IWorkbook workbook;
                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        workbook = new XSSFWorkbook(fs);
                    }

                    if (workbook.NumberOfSheets == 0)
                    {
                        ed.WriteMessage("\nExcel 文件中未找到工作表。");
                        return;
                    }

                    ISheet sheet = workbook.GetSheetAt(0); // 使用第一个工作表
                    if (sheet == null || sheet.LastRowNum < 0)
                    {
                        ed.WriteMessage("\n工作表为空。");
                        return;
                    }

                    // 计算实际数据区域：最大行号和最大列号
                    excelRows = sheet.LastRowNum + 1;   // 0‑based 转 1‑based
                                                        // 统计最大列数
                    int maxCol = 0;
                    for (int r = 0; r < excelRows; r++)
                    {
                        IRow row = sheet.GetRow(r);
                        if (row != null)
                        {
                            int lastCellNum = row.LastCellNum; // 1‑based
                            if (lastCellNum > maxCol)
                                maxCol = lastCellNum;
                        }
                    }
                    excelCols = maxCol;

                    DataFormatter formatter = new DataFormatter();

                    // 读取单元格文本
                    for (int r = 0; r < excelRows; r++)
                    {
                        IRow row = sheet.GetRow(r);
                        string[] rowArr = new string[excelCols];
                        for (int c = 0; c < excelCols; c++)
                        {
                            ICell cell = row?.GetCell(c);
                            string text = cell != null ? formatter.FormatCellValue(cell) : string.Empty;
                            rowArr[c] = text;
                        }
                        excelData.Add(rowArr);
                    }

                    // 收集合并单元格（直接使用 NPOI 的合并区域，所有索引已是 0‑based）
                    for (int i = 0; i < sheet.NumMergedRegions; i++)
                    {
                        CellRangeAddress region = sheet.GetMergedRegion(i);
                        int r1 = region.FirstRow;
                        int c1 = region.FirstColumn;
                        int r2 = region.LastRow;
                        int c2 = region.LastColumn;
                        // 只加入有效的合并区域（至少跨越两行/列）
                        if (r1 <= r2 && c1 <= c2 && r1 >= 0 && c1 >= 0 && r2 < excelRows && c2 < excelCols)
                        {
                            mergedRanges.Add((r1, c1, r2, c2));
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

                            int maxR = Math.Min(cadRows, excelRows);
                            int maxC = Math.Min(cadCols, excelCols);

                            // 构建“跳过”标记数组：用于处理 Excel 侧的合并单元格
                            bool[,] skipCell = new bool[excelRows, excelCols];
                            foreach (var m in mergedRanges)
                            {
                                int r1 = m.r1, c1 = m.c1, r2 = m.r2, c2 = m.c2;
                                // 边界保护
                                if (r1 < 0 || c1 < 0 || r2 >= excelRows || c2 >= excelCols) continue;
                                for (int rr = r1; rr <= r2; rr++)
                                {
                                    for (int cc = c1; cc <= c2; cc++)
                                    {
                                        // 保留左上角可写，其余跳过
                                        if (rr == r1 && cc == c1) continue;
                                        skipCell[rr, cc] = true;
                                    }
                                }
                            }

                            // 将重叠区域的数据写回 CAD 表格
                            for (int r = 0; r < maxR; r++)
                            {
                                for (int c = 0; c < maxC; c++)
                                {
                                    // 如果 Excel 此格属于合并区域且不是左上角，跳过
                                    if (r < excelRows && c < excelCols && skipCell[r, c])
                                        continue;

                                    string val = string.Empty;
                                    if (r < excelData.Count && c < excelData[r].Length)
                                        val = excelData[r][c] ?? string.Empty;

                                    try
                                    {
                                        table.Cells[r, c].TextString = val;
                                    }
                                    catch
                                    {
                                        // 某些单元格可能属于 CAD 侧的合并区域，写入失败则忽略
                                    }
                                }
                            }

                            // 原逻辑：对标题行的轻微处理（不修改样式，仅保留文本已写入）
                            // 此处不做额外操作

                            tr.Commit();
                            ed.WriteMessage($"\n已将 Excel ({Path.GetFileName(filePath)}) 的数据写入选中的表格（仅覆盖重叠单元，不改动表格样式/合并/尺寸）。");
                        }
                        catch (Exception exInner)
                        {
                            tr.Abort();
                            ed.WriteMessage($"\n将 Excel 写入表格时出错: {exInner.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument?.Editor?.WriteMessage($"\n导入失败: {ex.Message}");
            }
        }

        #endregion


        #region 表格相关

        /// <summary>
        /// 创建设备材料表
        /// </summary>
        /// <param name="db">当前数据库</param>
        /// <param name="deviceList">设备信息列表</param>
        /// <param name="scaleDenominator">绘图比例分母（用于计算文字高度和行高）</param>
        //private void CreateDeviceTable(Database db, List<DeviceInfo> deviceList, double scaleDenominator = 1, IEnumerable<string> includedFields = null)
        //{
        //    // 如果设备列表为空，直接返回
        //    if (db == null || deviceList == null || deviceList.Count == 0) return;

        //    // 按“名称+规格”进行合并统计：相同名称且相同规格的设备合并为一行，并累加数量
        //    var sameDeviceInfos = new Dictionary<string, DeviceInfo>(StringComparer.OrdinalIgnoreCase);
        //    foreach (var _oneDeviceInfo in deviceList)
        //    {
        //        // 空对象保护
        //        if (_oneDeviceInfo == null) continue;

        //        // 提取名称
        //        string _oneDeviceName = GetAttributeValueByMappedKey(_oneDeviceInfo.Attributes, "名称");
        //        if (string.IsNullOrWhiteSpace(_oneDeviceName)) _oneDeviceName = _oneDeviceInfo.Name ?? string.Empty;

        //        // 提取规格
        //        string _oneDeviceSpec = GetAttributeValueByMappedKey(_oneDeviceInfo.Attributes, "规格"); // 首先尝试“规格”
        //        if (string.IsNullOrWhiteSpace(_oneDeviceSpec)) _oneDeviceSpec = GetAttributeValueByMappedKey(_oneDeviceInfo.Attributes, "规格型号"); // 其次尝试“规格型号”
        //        if (string.IsNullOrWhiteSpace(_oneDeviceSpec)) _oneDeviceSpec = _oneDeviceInfo.Specifications ?? string.Empty; // 最后回退到实体字段


        //        // 提取材料
        //        string _oneDeviceMaterial = GetAttributeValueByMappedKey(_oneDeviceInfo.Attributes, "材料"); // 尝试从属性字典获取材料
        //        if (string.IsNullOrWhiteSpace(_oneDeviceMaterial)) _oneDeviceMaterial = GetAttributeValueByMappedKey(_oneDeviceInfo.Attributes, "材质"); // 兼容“材质”键名
        //        if (string.IsNullOrWhiteSpace(_oneDeviceMaterial)) _oneDeviceMaterial = _oneDeviceInfo.Material ?? string.Empty; // 回退到 DeviceInfo.Material

        //        // 提取规格
        //        string _oneDeviceSTDNo = GetAttributeValueByMappedKey(_oneDeviceInfo.Attributes, "图号或标准号"); // 尝试常用键
        //        if (string.IsNullOrWhiteSpace(_oneDeviceSTDNo)) _oneDeviceSTDNo = GetAttributeValueByMappedKey(_oneDeviceInfo.Attributes, "图号"); // 备用键
        //        if (string.IsNullOrWhiteSpace(_oneDeviceSTDNo)) _oneDeviceSTDNo = _oneDeviceInfo.DrawingNumber ?? string.Empty; // 回退到字段

        //        // 计算当前记录有效数量（优先 Quantity，再次 Count，最后默认为 1）
        //        int _oneDeviceInfoNo = 1;
        //        if (_oneDeviceInfo.Quantity > 0) _oneDeviceInfoNo = _oneDeviceInfo.Quantity;
        //        else if (_oneDeviceInfo.Count > 0) _oneDeviceInfoNo = _oneDeviceInfo.Count;
        //        else
        //        {
        //            string rawQty = GetAttributeValueByMappedKey(_oneDeviceInfo.Attributes, "数量");
        //            if (int.TryParse(rawQty, out int parsedQty) && parsedQty > 0) _oneDeviceInfoNo = parsedQty;
        //        }

        //        // 以“名称+规格”作为唯一键
        //        string groupKey = $"{_oneDeviceName.Trim()}||{_oneDeviceSpec.Trim()}||{_oneDeviceMaterial.Trim()}||{_oneDeviceSTDNo.Trim()}";// 联合键决定是否合并

        //        // 首次出现则创建分组 getGroup
        //        if (!sameDeviceInfos.TryGetValue(groupKey, out var getGroup))
        //        {
        //            var cloneOneDeviceInfo = CloneDeviceInfo(_oneDeviceInfo); // 深拷贝源对象以避免修改原始引用
        //            cloneOneDeviceInfo.Name = _oneDeviceName; // 设备信息中的名称
        //            cloneOneDeviceInfo.Specifications = _oneDeviceSpec; // 规格
        //            cloneOneDeviceInfo.Material = _oneDeviceMaterial; // 材料
        //            cloneOneDeviceInfo.DrawingNumber = _oneDeviceSTDNo;// 标准号
        //            cloneOneDeviceInfo.Quantity = _oneDeviceInfoNo; // 数量
        //            cloneOneDeviceInfo.Count = _oneDeviceInfoNo; // 计数

        //            if (cloneOneDeviceInfo.Attributes == null) cloneOneDeviceInfo.Attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //            cloneOneDeviceInfo.Attributes["名称"] = _oneDeviceName;
        //            cloneOneDeviceInfo.Attributes["规格"] = _oneDeviceSpec;
        //            cloneOneDeviceInfo.Attributes["材料"] = _oneDeviceMaterial;
        //            cloneOneDeviceInfo.Attributes["图号或标准号"] = _oneDeviceSTDNo;
        //            cloneOneDeviceInfo.Attributes["数量"] = _oneDeviceInfoNo.ToString();

        //            sameDeviceInfos[groupKey] = cloneOneDeviceInfo;
        //        }
        //        else
        //        {
        //            // 已存在同名称同规格行：数量累加
        //            getGroup.Quantity += _oneDeviceInfoNo;// 累加数量字段
        //            getGroup.Count += _oneDeviceInfoNo;// 累加计数字段
        //            if (getGroup.Attributes == null) getGroup.Attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        //            getGroup.Attributes["数量"] = getGroup.Quantity.ToString();// 更新属性字典中的数量显示
        //            // 对于其他属性：仅在目标为空时补充，不覆盖已有值
        //            if (string.IsNullOrWhiteSpace(getGroup.Material) && !string.IsNullOrWhiteSpace(_oneDeviceMaterial))
        //                getGroup.Material = _oneDeviceMaterial.Trim(); // 补充材料字段
        //            if (string.IsNullOrWhiteSpace(getGroup.Specifications) && !string.IsNullOrWhiteSpace(_oneDeviceSpec))
        //                getGroup.Specifications = _oneDeviceSpec.Trim(); // 补充规格字段
        //            if (string.IsNullOrWhiteSpace(getGroup.DrawingNumber) && !string.IsNullOrWhiteSpace(_oneDeviceSTDNo))
        //                getGroup.DrawingNumber = _oneDeviceSTDNo.Trim(); // 补充图号字段
        //        }
        //    }

        //    // 将分组结果转换为列表，继续使用原来表格样式生成逻辑（保持样式不变）
        //    var mergedDeviceList = sameDeviceInfos.Values.ToList(); // 合并后的设备列表
        //    if (mergedDeviceList.Count == 0) return; // 无数据则返回

        //    // 开启数据库事务
        //    using (Transaction trans = db.TransactionManager.StartTransaction())
        //    {
        //        try
        //        {
        //            // 获取块表（只读）
        //            BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        //            // 获取当前空间（模型空间或图纸空间，可写）
        //            BlockTableRecord currentSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

        //            // 构建动态列列表（根据所有设备的属性提取出的唯一列名）
        //            // 定义默认显示字段（按当前需求默认仅保留：名称/规格/数量）
        //            var defaultFields = new List<string> { "名称", "规格", "材料", "数量", "图号或标准号" };

        //            // 构建最终显示列：优先使用外部传入字段，未传入时使用默认字段
        //            var dynamicColumns = (includedFields ?? defaultFields)
        //                .Where(f => !string.IsNullOrWhiteSpace(f))
        //                .Select(f => NormalizeAttributeKey(f.Trim()))
        //                .Where(f => !string.IsNullOrWhiteSpace(f))
        //                .Distinct(StringComparer.OrdinalIgnoreCase)
        //                .ToList();

        //            // 防御处理：若筛选结果为空，则回退到默认字段
        //            if (dynamicColumns.Count == 0) dynamicColumns = defaultFields;

        //            int totalColumns = Math.Max(1, dynamicColumns.Count); // 计算总列数
        //            int dataRows = mergedDeviceList.Count;// 计算数据行数
        //            int totalRows = 1 + 1 + dataRows; // 计算总行数：1行标题 + 1行表头 + N行数据

        //            Table table = new Table();  // 创建新的表格对象
        //            table.SetSize(totalRows, totalColumns);    // 设置表格的行数和列数

        //            // 获取当前文档的编辑器对象，用于用户交互
        //            var ed = Application.DocumentManager.MdiActiveDocument?.Editor;
        //            if (ed == null) return;// 无编辑器则返回

        //            // 提示用户在 CAD 中指定表格的插入位置
        //            PromptPointResult ppr = ed.GetPoint("\n指定插入位置: ");
        //            // 如果用户成功指定了点
        //            if (ppr.Status == PromptStatus.OK) table.Position = ppr.Value; // 设置表格的插入点

        //            // 确定有效的比例分母：如果传入值有效则使用传入值，否则从数据库中获取当前绘图比例
        //            double effectiveScaleDenom = scaleDenominator > 0.0 ? scaleDenominator : VariableDictionary.wpfTextBoxScale;

        //            // 设置表格的样式（包括文字高度、行高、边框等，基于比例分母计算）
        //            SetTableStyle(db, table, trans, effectiveScaleDenom);

        //            // --- 生成表格标题 ---
        //            string titleName = string.Empty;
        //            // 尝试从第一个设备的属性或名称中提取标题依据
        //            if (mergedDeviceList.Count > 0)
        //            {
        //                var first = mergedDeviceList[0];
        //                // 优先查找“名称”属性
        //                if (first.Attributes != null && first.Attributes.TryGetValue("名称", out var nv) && !string.IsNullOrWhiteSpace(nv)) titleName = nv;
        //                else if (!string.IsNullOrWhiteSpace(first.Name)) titleName = first.Name; // 其次使用 DeviceInfo 的 Name 字段
        //            }
        //            // 如果仍未获取到名称，使用默认值“设备”
        //            if (string.IsNullOrWhiteSpace(titleName)) titleName = "设备";

        //            // 提取标题中的中文字符，用于生成更规范的中文标题
        //            string chinese = ExtractChineseCharacters(titleName);
        //            if (string.IsNullOrWhiteSpace(chinese)) chinese = titleName;

        //            // 组合最终标题，格式如：“泵 - 材料明细表”
        //            string fullTitle = $"{chinese}";

        //            // 合并第一行的所有单元格作为标题栏
        //            table.MergeCells(Autodesk.AutoCAD.DatabaseServices.CellRange.Create(table, 0, 0, 0, totalColumns - 1));
        //            // 设置标题文本
        //            table.Cells[0, 0].TextString = fullTitle;
        //            // 设置标题居中对齐
        //            table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

        //            // --- 填充表头（第二行）---
        //            for (int c = 0; c < dynamicColumns.Count && c < table.Columns.Count; c++)
        //            {
        //                // 获取当前列的键名（中文）
        //                string key = dynamicColumns[c];
        //                // 查找对应的英文翻译，如果没有则使用原键名
        //                string english = DictionaryHelper.ChineseToEnglish.ContainsKey(key) ? DictionaryHelper.ChineseToEnglish[key] : key;

        //                // 设置表头文本：如果有英文对照，则显示为“中文\n英文”，否则只显示中文
        //                table.Cells[1, c].TextString = string.Equals(english, key) ? key : (key + "\n" + english);
        //                // 设置表头居中对齐
        //                table.Cells[1, c].Alignment = CellAlignment.MiddleCenter;
        //            }

        //            // --- 填充数据行 ---
        //            int dataStart = 2; // 数据从第3行开始（索引2）
        //            for (int r = 0; r < dataRows; r++)
        //            {
        //                // 获取当前行的设备对象
        //                var item = mergedDeviceList[r];
        //                // 计算当前行在表格中的索引
        //                int rowIndex = dataStart + r;

        //                // 遍历每一列填充数据
        //                for (int c = 0; c < dynamicColumns.Count && c < table.Columns.Count; c++)
        //                {
        //                    // 获取当前列的键名
        //                    string colKey = dynamicColumns[c];
        //                    if (colKey == "名称")
        //                    {
        //                        colKey = "NAME";
        //                    }
        //                    else if(colKey == "规格")
        //                    {
        //                        colKey = "MODEL";
        //                    }
        //                    else if (colKey == "数量")
        //                    {
        //                        colKey = "QTY";
        //                    }
        //                    else if (colKey == "材料")
        //                    {
        //                        colKey = "MEDIUM";
        //                    }
        //                    else if (colKey == "图号或标准号")
        //                    {
        //                        colKey = "DRAWINGNO.STANDARDNO";
        //                    }
        //                    // 根据键名从设备属性中获取对应的值
        //                    string val = GetAttributeValueByMappedKey(item.Attributes, colKey);

        //                    if (string.Equals(colKey, "名称", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(val))
        //                        val = item.Name ?? string.Empty;
        //                    else if (string.Equals(colKey, "规格", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(val)) // 特殊处理：如果列是“规格”且值为空，使用 item.Specifications
        //                        val = item.Specifications ?? string.Empty;
        //                    else if (string.Equals(colKey, "数量", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(val))  // 特殊处理：如果列是“数量”且值为空，优先使用 item.Quantity，其次 item.Count
        //                    {
        //                        if (item.Quantity > 0)
        //                            val = item.Quantity.ToString();
        //                        else
        //                            val = item.Count > 0 ? item.Count.ToString() : string.Empty;
        //                    }
        //                    else if (string.Equals(colKey, "材料", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(val)) // 特殊处理：如果列是“材料”且值为空，使用 item.Material
        //                        val = item.Material ?? string.Empty;
        //                    else if (string.Equals(colKey, "图号与设计标准", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(val))// 特殊处理：如果列是“图号与设计标准”且值为空，使用 item.Material
        //                        val = item.Material ?? string.Empty;

        //                    // 将获取到的值写入表格单元格，如果值为 null 则写入空字符串
        //                    table.Cells[rowIndex, c].TextString = val ?? string.Empty;
        //                }
        //            }

        //            // 尝试应用基于比例的行列高度调整
        //            try { ApplyScaledHeightsToTable(table, effectiveScaleDenom); } catch { }

        //            // 【关键修改】在事务提交前，应用高级自动列宽调整
        //            // 这行代码确保了设备表也能像管道表一样，根据内容自动调整列宽，处理合并单元格
        //            AutoFitTableColumnsAdvanced(table);

        //            // 强制更新表格布局，使上述样式和宽度更改生效
        //            table.GenerateLayout();

        //            // 将表格添加到当前空间
        //            currentSpace.AppendEntity(table);
        //            // 将表格对象注册到事务中
        //            trans.AddNewlyCreatedDBObject(table, true);

        //            // 提交事务，保存更改
        //            trans.Commit();
        //        }
        //        catch
        //        {
        //            // 如果发生异常，回滚事务
        //            trans.Abort();
        //            // 重新抛出异常以便上层捕获
        //            throw;
        //        }
        //    }
        //}

        #region 创建设备材料表新方法

        /// <summary>
        /// 创建设备表格
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="deviceList">设备信息列表</param>
        /// <param name="scaleDenominator">比例分母</param>
        /// <param name="insertPosition">插入位置</param>
        /// <param name="includedFields">包含的字段列表</param>
        /// <returns>设备表创建结果</returns>
        private DeviceTableCreateResult CreateDeviceTable(
             Database db,
             List<DeviceInfo> deviceList,
             double scaleDenominator,
             Point3d insertPosition,
             IEnumerable<string> includedFields = null)
               {
                   // 参数校验：数据库和设备列表不能为空
                   if (db == null || deviceList == null || deviceList.Count == 0)
                       return new DeviceTableCreateResult();
             
                   LogManager.Instance.LogInfo($"[设备表][表格创建开始] InputRowCount={deviceList.Count}, Quantities=[{string.Join(",", deviceList.Select(item => $"{item.   Name}:     {item.Quantity}"))}]");
             
                   // ─── 本地辅助函数：从属性字典中安全读取第一个匹配键的值，并做Trim和空值归一化 ───
                   static string SafeAttr(Dictionary<string, string> attrs, string[] keys, string fallback)
                   {
                       if (attrs != null)
                       {
                           foreach (var k in keys)
                           {
                               if (attrs.TryGetValue(k, out var v) && !string.IsNullOrWhiteSpace(v))
                                   return v.Trim();
                           }
                       }
                       return fallback?.Trim() ?? string.Empty;
                   }
             
                   // ─── 1. 合并相同设备（按 名称/规格/材料/图号 分组） ───
                   // 使用 LINQ 简化聚合逻辑，并保证数量累加正确
                   var mergedDeviceList = deviceList
                       .Where(d => d != null)                                         // 过滤空对象
                       .Select(d =>
                       {
                           // 提前提取并缓存常用字段（只做一次 Trim）
                           var attrs = d.Attributes;
                           var name = SafeAttr(attrs, new[] { "名称" }, d.Name);
                           var spec = SafeAttr(attrs, new[] { "规格", "规格型号" }, d.Specifications);
                           var material = SafeAttr(attrs, new[] { "材料", "材质" }, d.Material);
                           var stdNo = SafeAttr(attrs, new[] { "图号或标准号", "图号" }, d.DrawingNumber);
             
                           // 计算记录数量：优先级 Quantity > Count > 属性“数量” > 默认1
                           int qty = d.Quantity > 0 ? d.Quantity : (d.Count > 0 ? d.Count : 0);
                           if (qty == 0)
                           {
                               var rawQty = SafeAttr(attrs, new[] { "数量" }, string.Empty);
                               if (!int.TryParse(rawQty, out qty) || qty <= 0)
                                   qty = 1;
                           }
             
                           return new
                           {
                               Key = $"{name}||{spec}||{material}||{stdNo}",  // 合并分组键
                               Source = d,
                               Name = name,
                               Spec = spec,
                               Material = material,
                               StdNo = stdNo,
                               Qty = qty
                           };
                       })
                       .GroupBy(x => x.Key, StringComparer.OrdinalIgnoreCase)          // 按合并键分组（不区分大小写）
                       .Select(g =>
                       {
                           // 取组内第一个作为模板，累加数量
                           var first = g.First();
                           var clone = CloneDeviceInfo(first.Source);   // 深拷贝原始对象
                           clone.Name = first.Name;
                           clone.Specifications = first.Spec;
                           clone.Material = first.Material;
                           clone.DrawingNumber = first.StdNo;
                           clone.Quantity = g.Sum(i => i.Qty);          // 累加数量
                           clone.Count = clone.Quantity;                // 同步 Count
                           if (clone.Attributes == null)
                               clone.Attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                           // 更新属性字典，保持一致性
                           clone.Attributes["名称"] = clone.Name;
                           clone.Attributes["规格"] = clone.Specifications;
                           clone.Attributes["材料"] = clone.Material;
                           clone.Attributes["图号或标准号"] = clone.DrawingNumber;
                           clone.Attributes["数量"] = clone.Quantity.ToString();
             
                           LogManager.Instance.LogInfo($"[设备表][表格二次聚合] NAME={clone.Name}, SourceRows={g.Count()}, FinalQuantity={clone.Quantity}");
                           return clone;
                       })
                       .ToList();
             
                   if (mergedDeviceList.Count == 0)
                       return new DeviceTableCreateResult();
             
                   LogManager.Instance.LogInfo($"[设备表][表格二次聚合完成] OutputRowCount={mergedDeviceList.Count}");
             
                   // ─── 2. 创建表格并填充数据 ───
                   using (Transaction trans = db.TransactionManager.StartTransaction())
                   {
                       try
                       {
                           // 获取块表和当前空间（模型空间或图纸空间）
                           BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                           BlockTableRecord currentSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                           if (currentSpace == null)
                               return new DeviceTableCreateResult();
             
                           // ─── 2.1 确定表格的列（字段） ───
                           var defaultFields = new List<string> { "名称", "规格", "材料", "数量", "图号或标准号" };
                           var dynamicColumns = (includedFields ?? defaultFields)
                               .Where(f => !string.IsNullOrWhiteSpace(f))
                               .Select(f => NormalizeAttributeKey(f.Trim()))
                               .Where(f => !string.IsNullOrWhiteSpace(f))
                               .Distinct(StringComparer.OrdinalIgnoreCase)
                               .ToList();
                           if (dynamicColumns.Count == 0)
                               dynamicColumns = defaultFields;
             
                           int totalColumns = Math.Max(1, dynamicColumns.Count);
                           int dataRows = mergedDeviceList.Count;
                           int totalRows = 1 + 1 + dataRows;   // 标题行 + 表头行 + 数据行
             
                           // ─── 2.2 创建表格对象并设置基本属性 ───
                           Table table = new Table();
                           table.SetSize(totalRows, totalColumns);
                           table.Position = insertPosition;   // 使用传入的插入点
             
                           double effectiveScaleDenom = scaleDenominator > 0.0
                               ? scaleDenominator
                               : VariableDictionary.wpfTextBoxScale;
                           SetTableStyle(db, table, trans, effectiveScaleDenom);   // 应用样式（文字高度、边框等）
             
                           // ─── 2.3 生成表格标题 ───
                           var firstDevice = mergedDeviceList.FirstOrDefault();
                           string titleName;
                           if (firstDevice != null && firstDevice.Attributes != null &&
                               firstDevice.Attributes.TryGetValue("名称", out var nameVal) && !string.IsNullOrWhiteSpace(nameVal))
                           {
                               titleName = nameVal;
                           }
                           else if (!string.IsNullOrWhiteSpace(firstDevice?.Name))
                           {
                               titleName = firstDevice.Name;
                           }
                           else
                           {
                               titleName = "设备";
                           }
                           string chinese = ExtractChineseCharacters(titleName);
                           if (string.IsNullOrWhiteSpace(chinese))
                               chinese = titleName;
                           string fullTitle = $"{chinese}";
             
                           // 合并第一行所有单元格作为标题
                           table.MergeCells(Autodesk.AutoCAD.DatabaseServices.CellRange.Create(table, 0, 0, 0, totalColumns - 1));
                           table.Cells[0, 0].TextString = fullTitle;
                           table.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;
             
                           // ─── 2.4 填充表头（第二行） ───
                           for (int c = 0; c < dynamicColumns.Count && c < table.Columns.Count; c++)
                           {
                               string key = dynamicColumns[c];
                               string english = DictionaryHelper.ChineseToEnglish.ContainsKey(key)
                                   ? DictionaryHelper.ChineseToEnglish[key]
                                   : key;
                               // 如果中英文不同，显示为 “中文\n英文”
                               table.Cells[1, c].TextString = string.Equals(english, key) ? key : (key + "\n" + english);
                               table.Cells[1, c].Alignment = CellAlignment.MiddleCenter;
                           }
             
                           // ─── 2.5 定义列名映射（将中文列名映射到内部属性键） ───
                           static string MapColumnKey(string col)
                           {
                               return col switch
                               {
                                   "名称" => "NAME",
                                   "规格" => "MODEL",
                                   "数量" => "QTY",
                                    "材料" => "MATERIAL",
                                   "图号或标准号" => "DRAWINGNO.STANDARDNO",
                                   _ => col
                               };
                           }
             
                           // ─── 2.6 统一函数：从设备项和列键获取最终显示值 ───
                           string GetValueForColumn(DeviceInfo item, string rawColKey)
                           {
                               // 数量列特殊处理：使用聚合后的 Quantity，不回退到属性中的旧值
                               if (string.Equals(rawColKey, "数量", StringComparison.OrdinalIgnoreCase))
                               {
                                   string quantityText = item.Quantity > 0
                                       ? item.Quantity.ToString(CultureInfo.InvariantCulture)
                                       : (item.Count > 0 ? item.Count.ToString(CultureInfo.InvariantCulture) : "1");
                                   LogManager.Instance.LogInfo(
                                       $"[设备表][数量列取值] NAME={item.Name}, Quantity={item.Quantity}, Count={item.Count}, AttributeQuantity={GetFirstAttributeValue        (item.Attributes, "数量")}, DisplayQuantity={quantityText}");
                                   return quantityText;
                               }

                                // 固定标准列优先使用聚合后的强类型字段，避免属性字典中
                                // 不同业务标签的模糊匹配造成名称、材料、数量和介质串列。
                                if (string.Equals(rawColKey, "名称", StringComparison.OrdinalIgnoreCase))
                                    return item.Name ?? string.Empty;
                                if (string.Equals(rawColKey, "规格", StringComparison.OrdinalIgnoreCase))
                                    return item.Specifications ?? string.Empty;
                                if (string.Equals(rawColKey, "材料", StringComparison.OrdinalIgnoreCase))
                                    return item.Material ?? string.Empty;
                                if (string.Equals(rawColKey, "图号或标准号", StringComparison.OrdinalIgnoreCase))
                                    return item.DrawingNumber ?? string.Empty;
             
                               var colKey = MapColumnKey(rawColKey);
                               // 先从属性字典取值
                               string val = GetAttributeValueByMappedKey(item.Attributes, colKey);
                               if (string.IsNullOrWhiteSpace(val))
                               {
                                   // 若属性字典无值，则从实体字段回退
                                   if (string.Equals(colKey, "NAME", StringComparison.OrdinalIgnoreCase))
                                       return item.Name ?? string.Empty;
                                   if (string.Equals(colKey, "MODEL", StringComparison.OrdinalIgnoreCase))
                                       return item.Specifications ?? string.Empty;
                                   if (string.Equals(colKey, "QTY", StringComparison.OrdinalIgnoreCase))
                                       return item.Quantity > 0 ? item.Quantity.ToString() : (item.Count > 0 ? item.Count.ToString() : string.Empty);
                                   if (string.Equals(colKey, "MEDIUM", StringComparison.OrdinalIgnoreCase))
                                       return item.Material ?? string.Empty;
                                   if (string.Equals(colKey, "DRAWINGNO.STANDARDNO", StringComparison.OrdinalIgnoreCase))
                                       return item.DrawingNumber ?? string.Empty;
                               }
                               return val ?? string.Empty;
                           }
             
                           // ─── 2.7 填充数据行 ───
                           int dataStart = 2;   // 数据从第3行开始（索引2）
                           for (int r = 0; r < dataRows; r++)
                           {
                               var item = mergedDeviceList[r];
                               int rowIndex = dataStart + r;
                               for (int c = 0; c < dynamicColumns.Count && c < table.Columns.Count; c++)
                               {
                                   string display = GetValueForColumn(item, dynamicColumns[c]);
                                   table.Cells[rowIndex, c].TextString = display ?? string.Empty;
                                   LogManager.Instance.LogInfo($"[设备表][表格写入] Row={rowIndex}, Column={dynamicColumns[c]}, NAME={item.Name}, Value={display ??    string.Empty}");
                               }
                           }
                            
                           // ─── 2.10 计算表格总高度（供调用者参考） ───
                           double tableHeight = 0.0;
                           for (int row = 0; row < table.Rows.Count; row++)
                           {
                               tableHeight += table.Rows[row].Height;
                           }
             
                           // ─── 2.11 将表格追加到当前空间并提交事务 ───
                           currentSpace.AppendEntity(table);
                           trans.AddNewlyCreatedDBObject(table, true);
                           trans.Commit();
             
                           return new DeviceTableCreateResult
                           {
                               TableId = table.ObjectId,
                               Position = insertPosition,
                               Height = tableHeight
                           };
                       }
                       catch
                       {
                           trans.Abort();
                           throw;
                       }
                   }
               }

        #endregion


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
            const int baseFixedCols = 9;
            // 定义管段号候选键数组
            string[] pipeNoKeys = new[] { "管段号", "管道号", "管段编号", "Pipeline No", "Pipeline", "Pipe No", "TAG_NO" };
            // 定义起点候选键数组
            string[] startKeys = new[] { "起点", "START_POINT", "From" };
            // 定义终点候选键数组
            string[] endKeys = new[] { "终点", "END_POINT", "To" };

            // 迁移 "介质" -> "介质名称"，确保字段统一
            foreach (var e in deviceList)
            {
                // 如果属性字典为空，跳过
                if (e.Attributes == null) continue;
                // 如果存在"介质"且不为空
                //if (e.Attributes.TryGetValue("MEDIUM", out var val) && !string.IsNullOrWhiteSpace(val))
                //{
                //    // 如果"介质名称"不存在或为空，则赋值
                //    if (!e.Attributes.ContainsKey("MEDIUM") || string.IsNullOrWhiteSpace(e.Attributes["MEDIUM"]))
                //        e.Attributes["MEDIUM"] = val;
                //}
                // 如果"介质名称"仍不存在，尝试其他英文键
                if (!e.Attributes.ContainsKey("MEDIUM"))
                {
                    // 尝试 "Medium"
                    if (e.Attributes.TryGetValue("Medium", out var mv) && !string.IsNullOrWhiteSpace(mv)) e.Attributes["MEDIUM"] = mv;
                    // 尝试 "Medium Name"
                    else if (e.Attributes.TryGetValue("Medium Name", out var mn) && !string.IsNullOrWhiteSpace(mn)) e.Attributes["MEDIUM"] = mn;
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
            var pipeGroupPreferred = new[] { "名称", "规格", "材料", "数量", "图号或标准号", "泵前/后" };
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
            // pipeGroupColumns.AddRange(remainingAttrs); // 1: 管道表，在14列(即所有preferred字段)后的就不要了
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
                    double initialScaleDenom = scaleDenominator > 0.0 ? scaleDenominator : VariableDictionary.wpfTextBoxScale;
                    // 设置表格样式（包括文字高度等）
                    SetTableStyle(db, table, tr, initialScaleDenom);

                    // 合并标题行单元格
                    table.MergeCells(Autodesk.AutoCAD.DatabaseServices.CellRange.Create(table, 0, 0, 0, table.Columns.Count - 1));
                    // 设置标题文本
                    table.Cells[0, 0].TextString = $"{typeTitle}";
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

                        #region ================== 填充第0列：管道标题 ==================
                        // 初始化标题变量
                        string title = null;
                        // 优先从属性字典中查找“管道标题”
                        if (item.Attributes != null && item.Attributes.TryGetValue("PIPELINETITLE", out var tv) && !string.IsNullOrWhiteSpace(tv))
                            title = tv;
                        else
                        {
                            // 如果属性中没有，则从 Name 字段截取（通常格式为 PIPE_XXX，截取后半部分）
                            var nm = item.Name ?? string.Empty;
                            int pos = nm.LastIndexOf('_');
                            // 如果存在下划线且后面有内容，则截取；否则使用完整名称
                            title = pos >= 0 && pos < nm.Length - 1 ? nm.Substring(pos + 1) : nm;
                        }
                        #endregion

                        #region ================== 填充第1列：管段号 ==================

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
                            var fallbackKeys = new[] { "管段编号", "Pipeline No", "Pipeline", "Pipe No", "TAG_NO" };
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
                            table.Cells[rowIndex, 0].TextString = pipeNoVal;
                        #endregion

                        #region ================== 填充第2列：起点 ==================
                        // 从属性中查找起点信息（支持“起点”、“STRAT_POINT”、“From”等键）
                        var startVal = FindFirstAttrValue(item.Attributes, startKeys);
                        if (!string.IsNullOrWhiteSpace(startVal)) table.Cells[rowIndex, 1].TextString = startVal;

                        #endregion

                        #region ================== 填充第3列：终点 ==================
                        // 从属性中查找终点信息（支持“终点”、“止点”、“To”等键）
                        var endVal = FindFirstAttrValue(item.Attributes, endKeys);
                        if (!string.IsNullOrWhiteSpace(endVal)) table.Cells[rowIndex, 2].TextString = endVal;

                        #endregion

                        #region ================== 填充第4列：管道等级 ==================

                        string pipeClass = string.Empty;
                        // 优先直接查找“管道等级”属性
                        if (item.Attributes != null && item.Attributes.TryGetValue("PIPE_CLASS", out var cls) && !string.IsNullOrWhiteSpace(cls))
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
                                var fallbackKeys = new[] { "等级", "Class", "管级", "PIPE_CLASS" };
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
                            table.Cells[rowIndex, 3].TextString = pipeClass;

                        #endregion

                        #region ================== 填充第5列：介质名称 ==================
                        // 查找介质相关信息
                        var mediumVal = FindFirstAttrValue(item.Attributes, new[] { "介质名称", "适用介质", "MEDIUM" });
                        if (!string.IsNullOrWhiteSpace(mediumVal)) table.Cells[rowIndex, 4].TextString = mediumVal;
                        #endregion

                        #region ================== 填充第6列：操作温度 ==================

                        string opTemp = string.Empty;
                        if (item.Attributes != null)
                        {
                            // 智能查找温度键：支持包含“操作温度”、“T(”、“℃”、“°C”的键名
                            var tempKey = item.Attributes.Keys.FirstOrDefault(k => !string.IsNullOrWhiteSpace(k) &&
                                (k.IndexOf("操作温度", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 k.IndexOf("WORK_TEMP", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 k.IndexOf("T(", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 k.IndexOf("℃", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 k.IndexOf("°C", StringComparison.OrdinalIgnoreCase) >= 0));
                            if (!string.IsNullOrWhiteSpace(tempKey))
                                opTemp = item.Attributes[tempKey];
                        }
                        // 如果智能查找未果，尝试精确匹配“操作温度”
                        if (!string.IsNullOrWhiteSpace(opTemp))
                            table.Cells[rowIndex, 5].TextString = opTemp;
                        else if (item.Attributes != null && item.Attributes.TryGetValue("操作温度", out var tval) && !string.IsNullOrWhiteSpace(tval))
                            table.Cells[rowIndex, 6].TextString = tval;
                        #endregion

                        #region ================== 填充第7列：操作压力 ==================
                        // 智能查找压力键：支持包含“操作压力”、“P(”、“MPa”的键名
                        var pressureKey = item.Attributes.Keys.FirstOrDefault(k => !string.IsNullOrWhiteSpace(k) &&
                            (k.IndexOf("操作压力", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             k.IndexOf("WORK_PRESSURE", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             k.IndexOf("压力等级", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             k.IndexOf("PN", StringComparison.OrdinalIgnoreCase) >= 0));
                        if (!string.IsNullOrWhiteSpace(pressureKey))
                            table.Cells[rowIndex, 6].TextString = item.Attributes[pressureKey];
                        #endregion

                        #region ================== 填充第8列：隔热隔声代号 ==================
                        if (item.Attributes != null && item.Attributes.TryGetValue("HOT\\SOUND_ISOLACODE", out var code) && !string.IsNullOrWhiteSpace(code))
                            table.Cells[rowIndex, 7].TextString = code;
                        #endregion

                        #region ================== 填充第9列：是否防腐 ==================
                        // 
                        if (item.Attributes != null && item.Attributes.TryGetValue("IS_ANTICORRO", out var anti) && !string.IsNullOrWhiteSpace(anti))
                            table.Cells[rowIndex, 8].TextString = anti;
                        #endregion

                        #region ================== 填充第10列：名称 ==================
                        // 
                        if (item.Attributes != null && item.Attributes.TryGetValue("NAME", out var NAME) && !string.IsNullOrWhiteSpace(NAME))
                            table.Cells[rowIndex, 9].TextString = NAME;
                        #endregion

                        #region ================== 填充第11列：规格 ==================
                        // 
                        string DNPN = string.Empty;
                        if (item.Attributes != null && item.Attributes.TryGetValue("DN", out var DN) && !string.IsNullOrWhiteSpace(DN))
                            DNPN = DN;
                        if (item.Attributes != null && item.Attributes.TryGetValue("PN", out var PN) && !string.IsNullOrWhiteSpace(PN))
                        {
                            DNPN = $"{DNPN}/{PN}";
                        }
                            table.Cells[rowIndex, 10].TextString = DNPN;
                        #endregion

                        #region ================== 填充第12列：介质 ==================
                        // 
                        if (item.Attributes != null && item.Attributes.TryGetValue("MEDIUM", out var MEDIUM) && !string.IsNullOrWhiteSpace(MEDIUM))
                            table.Cells[rowIndex, 11].TextString = MEDIUM;
                        #endregion

                        #region ================== 填充第13列：数量 ==================
                        // 
                        if (item.Attributes != null && item.Attributes.TryGetValue("QTY", out var QTY) && !string.IsNullOrWhiteSpace(QTY))
                            table.Cells[rowIndex, 12].TextString = QTY;
                        #endregion

                        #region ================== 填充第14列：图号与设计标准号 ==================
                        // 
                        if (item.Attributes != null && item.Attributes.TryGetValue("DRAWINGNO.STANDARDNO", out var DRAWINGNO) && !string.IsNullOrWhiteSpace(DRAWINGNO))
                            table.Cells[rowIndex, 13].TextString = DRAWINGNO;
                        #endregion

                        #region ================== 填充第15列：泵前、后 ==================
                        // 
                        if (item.Attributes != null && item.Attributes.TryGetValue("PUMP_BEFORE_AFTER", out var PUMP_BEFORE_AFTER) && !string.IsNullOrWhiteSpace(PUMP_BEFORE_AFTER))
                            table.Cells[rowIndex, 14].TextString = PUMP_BEFORE_AFTER;
                        #endregion


                        #region ================== 填充其余动态列（剩余属性） ==================
                        // 遍历剩余的动态属性键，填充到表格中
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
                        #endregion

                    }

                    #region ================== 关键修改区域 ==================

                    // 计算最终应用的比例分母
                    double appliedScaleDenom = scaleDenominator > 0.0 ? scaleDenominator : initialScaleDenom;
                    if (appliedScaleDenom <= 0.0) appliedScaleDenom = VariableDictionary.wpfTextBoxScale;

                    // 输出日志
                    ed.WriteMessage($"\n插入表格时使用的比例分母: {appliedScaleDenom}");

                    #endregion

                    // 第一步：先应用缩放后的文字高度和行高
                    // 必须在调整列宽之前执行，因为列宽计算依赖 TextHeight
                    try { ApplyScaledHeightsToTable(table, appliedScaleDenom); } catch { }
                                

                    // ================== 结束关键修改区域 ==================

                    // 【新增】在事务提交前，自动调整列宽
                    AutoFitTableColumnsAdvanced(table, appliedScaleDenom);
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


        #region 兼容

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
            // 6) 额外去掉换行符，避免表格显示问题
            s = Regex.Replace(s, "\r", " ").Trim();
            // 7）去掉换行符，避免表格显示问题
            s = Regex.Replace(s, "\n", " ").Trim();
            return s;
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
            // 如果没有找到闭合三角形箭头，则尝试寻找其他闭合多段线作为箭头
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
        /// 构建局部坐标的管线 Polyline
        /// </summary>
        /// <param name="template">模板 Polyline</param>
        /// <param name="verticesWorld">全局坐标系下的顶点列表</param>
        /// <param name="midPointWorld">全局坐标系下的中点</param>
        /// <returns>局部坐标系下的管线 Polyline</returns>
        private Polyline BuildPipePolylineLocal(Polyline template, List<Point3d> verticesWorld, Point3d midPointWorld)
        {
            double lineWeightScale = AutoCadHelper.GetScale();// 获取当前图形的比例分母，用于缩放线宽

            var pl = new Polyline();

            // 如果模板或顶点为空，直接返回空多段线，避免空引用
            if (template == null || verticesWorld == null || verticesWorld.Count == 0)
                return pl;
            // 你要求的管道线宽：0.3
            double pipeLineWidth = 0.3 * lineWeightScale;

            // 依次把世界坐标点转换为相对中点的局部坐标
            for (int i = 0; i < verticesWorld.Count; i++)
            {
                var worldPt = verticesWorld[i];
                var localPt = new Point2d(
                    worldPt.X - midPointWorld.X,
                    worldPt.Y - midPointWorld.Y);

                // 以 0.3 的固定宽度创建顶点，确保生成的管道线宽度稳定
                pl.AddVertexAt(i, localPt, 0.0, pipeLineWidth, pipeLineWidth);
            }

            // 继承模板的图层、颜色、线型等属性
            pl.Layer = template.Layer;
            pl.Color = template.Color;
            pl.LineWeight = template.LineWeight;
            pl.Linetype = template.Linetype;
            pl.LinetypeScale = template.LinetypeScale;

            // 统一设置为 XY 平面上的 2D 线
            pl.Elevation = 0;
            pl.Normal = Vector3d.ZAxis;
            pl.Closed = false;

            // 额外设置常量宽度，避免某些情况下显示不一致
            pl.ConstantWidth = pipeLineWidth;

            return pl;
        }

        /// <summary>
        /// 辅助：在视觉树中递归查找名为 name 的子控件（泛型）
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="parent"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        private static T? FindChildByName<T>(DependencyObject parent, string name) where T : DependencyObject
        {
            if (parent == null) return null;
            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is FrameworkElement fe && fe.Name == name && child is T t) return t;
                var result = FindChildByName<T>(child, name);
                if (result != null) return result;
            }
            return null;
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
            var result = new List<AttributeDefinition>();// 结果属性定义列表
            bool hasTitle = false; // 标记是否已经存在管道标题属性
            // 遍历原始属性定义列表
            foreach (var def in defs)
            {
                var cloned = def.Clone() as AttributeDefinition; // 克隆属性定义
                if (cloned == null) continue; // 如果克隆失败则跳过

                // 转为局部坐标（相对中点）
                var localPos = new Point3d(def.Position.X - midPointWorld.X, def.Position.Y - midPointWorld.Y, 0); // 局部坐标，Z=0
                cloned.Position = localPos; // 设置克隆属性定义的位置
                cloned.Rotation = def.Rotation; // 保留原始旋转角度
                cloned.Invisible = def.Invisible; // 保留原始可见性
                cloned.Constant = def.Constant; // 保留原始常量属性
                cloned.Tag = def.Tag; // 保留原始标签
                cloned.TextString = def.TextString; // 保留原始文本内容
                cloned.Height = def.Height; // 保留原始高度
                // 保留其他属性（如宽度、样式等）
                if (!string.IsNullOrWhiteSpace(cloned.Tag))
                {
                    var tagLower = cloned.Tag.ToLowerInvariant(); // 转为小写以便比较
                    if (tagLower.Contains("长度") || tagLower.Contains("length")) // 如果标签包含“长度”或“length”，则更新文本内容为管道长度
                    {
                        double baseValue = 0.0; // 默认基准值为0
                        if (double.TryParse(cloned.TextString, out double parsed)) baseValue = parsed; // 尝试解析原始文本为数字
                        cloned.TextString = (baseValue + pipelineLength).ToString("0.###"); // 更新文本内容为基准值加上管道长度，保留三位小数
                    }
                    if (string.Equals(cloned.Tag, "PIPELINETITLE", StringComparison.OrdinalIgnoreCase)) // 如果标签为“管道标题”，则标记已存在标题
                    {
                        hasTitle = true; // 标记已存在管道标题
                        cloned.Position = Point3d.Origin; // 将标题位置设置为局部原点
                        cloned.Rotation = finalRotation; // 将标题旋转角度设置为最终旋转角度
                        cloned.Invisible = false; // 确保标题可见
                        if (string.IsNullOrWhiteSpace(cloned.TextString))
                            cloned.TextString = titleFallback ?? "PIPELINE"; // 如果标题文本为空，则使用后备值或默认值
                    }
                }

                result.Add(cloned); // 将克隆的属性定义添加到结果列表
            }

            if (!hasTitle) // 如果没有找到管道标题属性，则创建一个新的管道标题属性
            {
                // 创建新的管道标题属性定义
                result.Add(new AttributeDefinition
                {
                    Tag = "PIPELINETITLE", // 标签为“PIPELINETITLE”
                    Position = Point3d.Origin, // 位置为局部原点
                    Rotation = finalRotation, // 旋转角度为最终旋转角度
                    TextString = string.IsNullOrWhiteSpace(titleFallback) ? "PIPELINE" : titleFallback, // 文本内容为后备值或默认值
                    Height = defs != null && defs.Count > 0 ? defs[0].Height : 2.5, // 高度为原始属性定义的高度或默认值
                    Invisible = false, // 确保标题可见
                    Constant = false // 设置为非常量
                });
            }

            return result; // 返回结果属性定义列表
        }

        /// <summary>
        /// 创建方向箭头和标题文字
        /// </summary>
        /// <param name="tr">事务</param>
        /// <param name="sampleInfo"> 管道信息</param>
        /// <param name="verticesWorld"> 世界坐标 </param>
        /// <param name="midPointWorld">最小点坐标</param>
        /// <param name="pipeTitle">管道标题</param>
        /// <param name="sampleBlockName">示例块名称</param>
        /// <returns>返回管道实体列表</returns>
        private List<Entity> CreateDirectionalArrowsAndTitles(DBTrans tr, SamplePipeInfo sampleInfo, List<Point3d> verticesWorld, Point3d midPointWorld, string pipeTitle, string sampleBlockName)
        {
            var overlay = new List<Entity>();// 结果实体列表
            if (sampleInfo == null || verticesWorld == null || verticesWorld.Count < 2) return overlay;

            // 优先使用用户在TextBox_绘图比例中设置的比例值
            var scaleFactor = AutoCadHelper.GetScale();

            // 箭头模板与填充准备：若无模板则用默认三角
            Polyline arrowTemplate = sampleInfo.DirectionArrowTemplate;
            Solid? fillTemplate = null;
            double explicitArrowLength = 8;  // 基础长度
            double explicitArrowHeight = 2.0;   // 基础高度
            if (arrowTemplate == null)
            {
                // 根据名称确定箭头样式
                var (colorIdx, length, height) = DetermineArrowStyleByName(sampleBlockName);
                explicitArrowLength = length * scaleFactor;
                explicitArrowHeight = height * scaleFactor;
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
                var p1 = verticesWorld[i];     // 一段管线的P1点；
                var p2 = verticesWorld[i + 1]; // 管线的P2点；
                var seg = p2 - p1;             // P2与P1点间的差值；
                if (seg.IsZeroLength() || seg.Length <= 50.0 * AutoCadHelper.GetScale()) continue;  //判断距离是不是大于50

                var dir = seg.GetNormal();
                var mid = new Point3d((p1.X + p2.X) / 2.0, (p1.Y + p2.Y) / 2.0, (p1.Z + p2.Z) / 2.0); //计算P1与P2的中点；

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
                    double offset = (4 * scaleFactor + finalTitleHeight * 0.75); // 应用比例
                    //offset = Math.Max(finalTitleHeight * 0.75, arrowHalfHeight + finalTitleHeight * 0.25);
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
            // 返回生成的箭头和标题实体列表
            return overlay;
        }

        /// <summary>
        /// 根据名称确定箭头样式
        /// </summary>
        /// <param name="blockName">块名称</param>
        /// <returns>箭头样式元组</returns>
        private (short colorIndex, double length, double height) DetermineArrowStyleByName(string blockName)
        {
            string nameLower = (blockName ?? string.Empty).ToLowerInvariant();// 检查传进来的块名是不是空
            bool isOutlet = nameLower.Contains("出口") || nameLower.Contains("outlet"); // 判断块名内有没有"出口"或"outlet"

            bool isInlet = nameLower.Contains("进口") || nameLower.Contains("入口") || nameLower.Contains("inlet"); // 判断块名内有没有"入口"或"inlet"

            // 出口=黄色(ACI 2)，入口=绿色(ACI 3)，默认黄色
            short colorIndex = isInlet ? (short)6 : (short)2;
            if (!isInlet && !isOutlet)
            {
                colorIndex = 2;
            }

            return (colorIndex, 10.0, 2.0);
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
        /// 兼容补丁：获取下一个管段号。
        /// </summary>
        private int GetNextPipeSegmentNumber(Database db)
        {
            int max = 0;
            if (db == null) return 1;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                foreach (ObjectId id in ms)
                {
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;
                    foreach (ObjectId aid in br.AttributeCollection)
                    {
                        var ar = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                        if (ar == null) continue;
                        if (!string.Equals(ar.Tag, "管段号", StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(ar.Tag, "管段编号", StringComparison.OrdinalIgnoreCase)) continue;

                        var m = Regex.Match(ar.TextString ?? string.Empty, @"\d+");
                        if (m.Success && int.TryParse(m.Value, out var n) && n > max) max = n;
                    }
                }
                tr.Commit();
            }
            return max + 1;
        }

        /// <summary>
        /// 兼容补丁：构建块定义并返回块名。
        /// </summary>
        private string BuildPipeBlockDefinition(DBTrans tr, string desiredName, Polyline pipeLocal, List<Entity> overlayEntities, List<AttributeDefinition> attDefsLocal)
        {
            string finalName = string.IsNullOrWhiteSpace(desiredName) ? "PIPE_BLOCK" : desiredName; // 默认块名
            int suf = 1; // 后缀计数器
            while (tr.BlockTable.Has(finalName)) finalName = (desiredName ?? "PIPE_BLOCK") + "_PIPEGEN_" + suf++; // 确保块名唯一
            // 创建块定义
            tr.BlockTable.Add(
                finalName, // 块名
                btr => { btr.Origin = Point3d.Origin; },
                () =>
                {
                    var entities = new List<Entity>(); // 块定义实体列表
                    if (pipeLocal != null) entities.Add((Polyline)pipeLocal.Clone()); // 克隆管道 Polyline
                    if (overlayEntities != null) // 克隆覆盖实体
                    {
                        foreach (var e in overlayEntities) // 遍历覆盖实体
                        {
                            if (e == null) continue; // 如果为空则跳过
                            var c = e.Clone() as Entity; // 克隆实体
                            if (c != null) entities.Add(c); // 如果克隆成功则添加到列表
                        }
                    }
                    return entities; // 返回块定义实体列表
                },
                () => attDefsLocal ?? new List<AttributeDefinition>() // 返回属性定义列表
            );
            // 返回最终块名
            return finalName;
        }

        /// <summary>
        /// 兼容补丁：插入块并写入属性。
        /// </summary>
        private ObjectId InsertPipeBlockWithAttributes(DBTrans tr, Point3d insertPointWorld, string blockName, double rotation, Dictionary<string, string> attValues)
        {
            ObjectId btrId = tr.BlockTable[blockName];
            return tr.CurrentSpace.InsertBlock(insertPointWorld, btrId, rotation: rotation, atts: attValues);
        }

        #endregion
    }


    /// <summary>
    /// 管道属性编辑器（带字段排序与中文别名展示）。
    /// </summary>
    public class PipeAttributeEditorForm : Form
    {
        private readonly DataGridView _dataGridView;
        private readonly Button _btnOk;
        private readonly Button _btnCancel;
        private Dictionary<string, string> _attributes;

        /// <summary>
        /// 编辑后属性结果。
        /// </summary>
        public Dictionary<string, string> Attributes => new Dictionary<string, string>(_attributes, StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 管道属性编辑表单构造函数，接受初始属性字典（可为 null）。表单设计为固定大小的对话框，包含一个 DataGridView 用于显示和编辑属性，以及“完成”和“取消”按钮。DataGridView 的第一列显示属性字段（只读），第二列显示属性值（可编辑）。加载时根据 DictionaryHelper 的规则对属性进行排序和别名展示。点击“完成”时收集 DataGridView 中的属性值并更新 Attributes 属性；点击“取消”则关闭表单不保存更改。
        /// </summary>
        /// <param name="initialAttributes"></param>
        public PipeAttributeEditorForm(Dictionary<string, string> initialAttributes)
        {
            _attributes = new Dictionary<string, string>(initialAttributes ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase);

            Text = "示例管道属性编辑";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(760, 520);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;

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
            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn { Name = "Key", HeaderText = "字段", ReadOnly = true });
            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn { Name = "Value", HeaderText = "值", ReadOnly = false });

            _btnOk = new Button { Text = "完成", DialogResult = DialogResult.OK, Width = 90, Height = 30 };
            _btnCancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Width = 90, Height = 30 };
            _btnOk.Click += BtnOk_Click;
            _btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 52,
                FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                Padding = new Padding(8),
                WrapContents = false
            };
            panel.Controls.Add(_btnCancel);
            panel.Controls.Add(_btnOk);

            Controls.Add(_dataGridView);
            Controls.Add(panel);

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;

            LoadAttributesToGrid();
        }

        /// <summary>
        /// 字段排序与别名展示：规则来自 DictionaryHelper，可配置。
        /// </summary>
        private void LoadAttributesToGrid()
        {
            _dataGridView.Rows.Clear();

            var ordered = _attributes
                .OrderBy(kv => DictionaryHelper.GetPipeAttributeSortPriority(kv.Key))
                .ThenBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var kv in ordered)
            {
                _dataGridView.Rows.Add(DictionaryHelper.GetPipeAttributeDisplayAlias(kv.Key), kv.Value);
                _dataGridView.Rows[_dataGridView.Rows.Count - 1].Tag = kv.Key;
            }

            if (_dataGridView.Rows.Count > 0)
                _dataGridView.CurrentCell = _dataGridView.Rows[0].Cells[1];
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            try
            {
                var newDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < _dataGridView.Rows.Count; i++)
                {
                    var row = _dataGridView.Rows[i];
                    if (row.IsNewRow) continue;

                    var key = (row.Tag as string) ?? row.Cells["Key"].Value?.ToString() ?? string.Empty;
                    var val = row.Cells["Value"].Value?.ToString() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(key)) continue;

                    newDict[key] = val;
                }

                _attributes = newDict;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存属性失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
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
        this.dataGridView.Location = new System.Drawing.Point(12, 12);
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
            System.Windows.Forms.MessageBox.Show("请先选择要删除的行！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

