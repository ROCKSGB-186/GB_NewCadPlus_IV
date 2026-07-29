using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows.Input;
using Google.Protobuf.WellKnownTypes;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_IV.ViewModels
{
    public class PipeCalcViewModel : INotifyPropertyChanged
    {
        /// <summary>
        /// 属性更改事件处理程序
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// 已更改属性
        /// </summary>
        /// <param name="name"></param>
        protected void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        #region ==================== 数据持久化 ====================
        private const string DataFileName = "PipeCalcData.json";

        private string GetDataFilePath()
        {
            // 获取 AppData 目录
            string appDataDir = UniFiedStandards.GetPath.AppPath;
            // 如果 AppPath 返回的是目录，拼接文件名；如果是完整路径，直接返回
            if (Directory.Exists(appDataDir) || !Path.HasExtension(appDataDir))
            {
                return Path.Combine(appDataDir, DataFileName); // 拼接目录部分和数据文件名，得到完整的数据文件路径
            }
            return appDataDir;// 如果 AppPath 返回的是完整路径，则直接返回
        }

        public void LoadData()
        {
            string filePath = GetDataFilePath();
            if (!File.Exists(filePath)) return;

            try
            {
                string json = File.ReadAllText(filePath);
                var data = JsonSerializer.Deserialize<CalcDataModel>(json);
                if (data == null) return;

                // ---- 脱硫系统 ----
                InletSO2 = data.InletSO2;
                OutletSO2 = data.OutletSO2;
                AtmPressure = data.AtmPressure;
                OutletTemp = data.OutletTemp;
                StandardFlow = data.StandardFlow;
                StackVelocity = data.StackVelocity;
                EquipCount = data.EquipCount;
                CaSRatio = data.CaSRatio;

                // ---- 循环泵 ----
                LiquidGasRatio = data.LiquidGasRatio;
                ResidenceTime = data.ResidenceTime;
                SprayLayers = data.SprayLayers;
                PumpHead = data.PumpHead;
                MotorFactor = data.MotorFactor;
                SelectedPumpFlow = data.SelectedPumpFlow;
                MotorPower = data.MotorPower;
                PumpSpareCount = data.PumpSpareCount;

                // ---- 管道 ----
                InletVelocity = data.InletVelocity;
                InletDiameter = data.InletDiameter;
                OutletVelocity = data.OutletVelocity;
                OutletDiameter = data.OutletDiameter;

                // ---- 氧化风机 ----
                FanHead = data.FanHead;
                FanMotorFactor = data.FanMotorFactor;
                SelectedFanFlow = data.SelectedFanFlow;
                FanMotorPower = data.FanMotorPower;
                FanRunningCount = data.FanRunningCount;
                FanSpareCount = data.FanSpareCount;
                FanOutletVelocity = data.FanOutletVelocity;
                FanOutletDiameter = data.FanOutletDiameter;

                // ---- 介质/条件 ----
                InletMedium = data.InletMedium;
                InletCondition = data.InletCondition;
                OutletMedium = data.OutletMedium;
                OutletCondition = data.OutletCondition;
                FanOutletMedium = data.FanOutletMedium;
                FanOutletCondition = data.FanOutletCondition;

                // ---- 石灰石浆液泵（远） ----
                LimePurity = data.LimePurity;
                SlurryDensity = data.SlurryDensity;
                FarSystemCount = data.FarSystemCount;
                FarRunHours = data.FarRunHours;
                FarPumpCount = data.FarPumpCount;
                FarPumpHead = data.FarPumpHead;
                FarMotorFactor = data.FarMotorFactor;
                FarPumpEfficiency = data.FarPumpEfficiency;
                FarSelectedFlow = data.FarSelectedFlow;
                FarMotorPower = data.FarMotorPower;

                // ---- 石灰石浆液泵（近） ----
                NearSystemCount = data.NearSystemCount;
                NearRunHours = data.NearRunHours;
                NearPumpCount = data.NearPumpCount;
                NearPumpHead = data.NearPumpHead;
                NearMotorFactor = data.NearMotorFactor;
                NearPumpEfficiency = data.NearPumpEfficiency;
                NearSelectedFlow = data.NearSelectedFlow;
                NearMotorPower = data.NearMotorPower;

                // ---- 石膏排出泵（远） ----
                GypsumDensity = data.GypsumDensity;
                FarGypsumSystemCount = data.FarGypsumSystemCount;
                FarGypsumRunHours = data.FarGypsumRunHours;
                FarGypsumPumpCount = data.FarGypsumPumpCount;
                FarGypsumHead = data.FarGypsumHead;
                FarGypsumMotorFactor = data.FarGypsumMotorFactor;
                FarGypsumEfficiency = data.FarGypsumEfficiency;
                FarGypsumSelectedFlow = data.FarGypsumSelectedFlow;
                FarGypsumMotorPower = data.FarGypsumMotorPower;

                // ---- 石膏排出泵（近） ----
                NearGypsumSystemCount = data.NearGypsumSystemCount;
                NearGypsumRunHours = data.NearGypsumRunHours;
                NearGypsumPumpCount = data.NearGypsumPumpCount;
                NearGypsumHead = data.NearGypsumHead;
                NearGypsumMotorFactor = data.NearGypsumMotorFactor;
                NearGypsumEfficiency = data.NearGypsumEfficiency;
                NearGypsumSelectedFlow = data.NearGypsumSelectedFlow;
                NearGypsumMotorPower = data.NearGypsumMotorPower;

                // ---- 缓冲泵 ----
                BufferRunHours = data.BufferRunHours;
                BufferSystemCount = data.BufferSystemCount;
                BufferPumpCount = data.BufferPumpCount;
                BufferHead = data.BufferHead;
                BufferMotorFactor = data.BufferMotorFactor;
                BufferEfficiency = data.BufferEfficiency;
                BufferSelectedFlow = data.BufferSelectedFlow;
                BufferMotorPower = data.BufferMotorPower;

                // ---- 滤液水泵（远） ----
                FiltrateDensity = data.FiltrateDensity;
                FiltrateFarSystemCount = data.FiltrateFarSystemCount;
                FiltrateFarPumpCount = data.FiltrateFarPumpCount;
                FiltrateFarHead = data.FiltrateFarHead;
                FiltrateFarMotorFactor = data.FiltrateFarMotorFactor;
                FiltrateFarEfficiency = data.FiltrateFarEfficiency;
                FiltrateFarSelectedFlow = data.FiltrateFarSelectedFlow;
                FiltrateFarMotorPower = data.FiltrateFarMotorPower;

                // ---- 滤液水泵（近） ----
                FiltrateNearSystemCount = data.FiltrateNearSystemCount;
                FiltrateNearPumpCount = data.FiltrateNearPumpCount;
                FiltrateNearHead = data.FiltrateNearHead;
                FiltrateNearMotorFactor = data.FiltrateNearMotorFactor;
                FiltrateNearEfficiency = data.FiltrateNearEfficiency;
                FiltrateNearSelectedFlow = data.FiltrateNearSelectedFlow;
                FiltrateNearMotorPower = data.FiltrateNearMotorPower;

                // ---- 事故泵（远） ----
                AccidentDensity = data.AccidentDensity;
                AccidentFarPumpCount = data.AccidentFarPumpCount;
                AccidentFarHead = data.AccidentFarHead;
                AccidentFarMotorFactor = data.AccidentFarMotorFactor;
                AccidentFarEfficiency = data.AccidentFarEfficiency;
                AccidentFarSelectedFlow = data.AccidentFarSelectedFlow;
                AccidentFarMotorPower = data.AccidentFarMotorPower;

                // ---- 事故泵（近） ----
                AccidentNearPumpCount = data.AccidentNearPumpCount;
                AccidentNearHead = data.AccidentNearHead;
                AccidentNearMotorFactor = data.AccidentNearMotorFactor;
                AccidentNearEfficiency = data.AccidentNearEfficiency;
                AccidentNearSelectedFlow = data.AccidentNearSelectedFlow;
                AccidentNearMotorPower = data.AccidentNearMotorPower;

                // ---- 地坑泵 ----
                SumpPumpCount = data.SumpPumpCount;
                SumpHead = data.SumpHead;
                SumpMotorFactor = data.SumpMotorFactor;
                SumpEfficiency = data.SumpEfficiency;
                SumpSelectedFlow = data.SumpSelectedFlow;
                SumpMotorPower = data.SumpMotorPower;

                CalculateAll();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"加载数据失败: {ex.Message}");
            }
        }

        public void SaveData()
        {
            try
            {
                var data = new CalcDataModel();

                // ---- 脱硫系统 ----
                data.InletSO2 = InletSO2;
                data.OutletSO2 = OutletSO2;
                data.AtmPressure = AtmPressure;
                data.OutletTemp = OutletTemp;
                data.StandardFlow = StandardFlow;
                data.StackVelocity = StackVelocity;
                data.EquipCount = EquipCount;
                data.CaSRatio = CaSRatio;

                // ---- 循环泵 ----
                data.LiquidGasRatio = LiquidGasRatio;
                data.ResidenceTime = ResidenceTime;
                data.SprayLayers = SprayLayers;
                data.PumpHead = PumpHead;
                data.MotorFactor = MotorFactor;
                data.SelectedPumpFlow = SelectedPumpFlow;
                data.MotorPower = MotorPower;
                data.PumpSpareCount = PumpSpareCount;

                // ---- 管道 ----
                data.InletVelocity = InletVelocity;
                data.InletDiameter = InletDiameter;
                data.OutletVelocity = OutletVelocity;
                data.OutletDiameter = OutletDiameter;

                // ---- 氧化风机 ----
                data.FanHead = FanHead;
                data.FanMotorFactor = FanMotorFactor;
                data.SelectedFanFlow = SelectedFanFlow;
                data.FanMotorPower = FanMotorPower;
                data.FanRunningCount = FanRunningCount;
                data.FanSpareCount = FanSpareCount;
                data.FanOutletVelocity = FanOutletVelocity;
                data.FanOutletDiameter = FanOutletDiameter;

                // ---- 介质/条件 ----
                data.InletMedium = InletMedium;
                data.InletCondition = InletCondition;
                data.OutletMedium = OutletMedium;
                data.OutletCondition = OutletCondition;
                data.FanOutletMedium = FanOutletMedium;
                data.FanOutletCondition = FanOutletCondition;

                // ---- 石灰石浆液泵（远） ----
                data.LimePurity = LimePurity;
                data.SlurryDensity = SlurryDensity;
                data.FarSystemCount = FarSystemCount;
                data.FarRunHours = FarRunHours;
                data.FarPumpCount = FarPumpCount;
                data.FarPumpHead = FarPumpHead;
                data.FarMotorFactor = FarMotorFactor;
                data.FarPumpEfficiency = FarPumpEfficiency;
                data.FarSelectedFlow = FarSelectedFlow;
                data.FarMotorPower = FarMotorPower;

                // ---- 石灰石浆液泵（近） ----
                data.NearSystemCount = NearSystemCount;
                data.NearRunHours = NearRunHours;
                data.NearPumpCount = NearPumpCount;
                data.NearPumpHead = NearPumpHead;
                data.NearMotorFactor = NearMotorFactor;
                data.NearPumpEfficiency = NearPumpEfficiency;
                data.NearSelectedFlow = NearSelectedFlow;
                data.NearMotorPower = NearMotorPower;

                // ---- 石膏排出泵（远） ----
                data.GypsumDensity = GypsumDensity;
                data.FarGypsumSystemCount = FarGypsumSystemCount;
                data.FarGypsumRunHours = FarGypsumRunHours;
                data.FarGypsumPumpCount = FarGypsumPumpCount;
                data.FarGypsumHead = FarGypsumHead;
                data.FarGypsumMotorFactor = FarGypsumMotorFactor;
                data.FarGypsumEfficiency = FarGypsumEfficiency;
                data.FarGypsumSelectedFlow = FarGypsumSelectedFlow;
                data.FarGypsumMotorPower = FarGypsumMotorPower;

                // ---- 石膏排出泵（近） ----
                data.NearGypsumSystemCount = NearGypsumSystemCount;
                data.NearGypsumRunHours = NearGypsumRunHours;
                data.NearGypsumPumpCount = NearGypsumPumpCount;
                data.NearGypsumHead = NearGypsumHead;
                data.NearGypsumMotorFactor = NearGypsumMotorFactor;
                data.NearGypsumEfficiency = NearGypsumEfficiency;
                data.NearGypsumSelectedFlow = NearGypsumSelectedFlow;
                data.NearGypsumMotorPower = NearGypsumMotorPower;

                // ---- 缓冲泵 ----
                data.BufferRunHours = BufferRunHours;
                data.BufferSystemCount = BufferSystemCount;
                data.BufferPumpCount = BufferPumpCount;
                data.BufferHead = BufferHead;
                data.BufferMotorFactor = BufferMotorFactor;
                data.BufferEfficiency = BufferEfficiency;
                data.BufferSelectedFlow = BufferSelectedFlow;
                data.BufferMotorPower = BufferMotorPower;

                // ---- 滤液水泵（远） ----
                data.FiltrateDensity = FiltrateDensity;
                data.FiltrateFarSystemCount = FiltrateFarSystemCount;
                data.FiltrateFarPumpCount = FiltrateFarPumpCount;
                data.FiltrateFarHead = FiltrateFarHead;
                data.FiltrateFarMotorFactor = FiltrateFarMotorFactor;
                data.FiltrateFarEfficiency = FiltrateFarEfficiency;
                data.FiltrateFarSelectedFlow = FiltrateFarSelectedFlow;
                data.FiltrateFarMotorPower = FiltrateFarMotorPower;

                // ---- 滤液水泵（近） ----
                data.FiltrateNearSystemCount = FiltrateNearSystemCount;
                data.FiltrateNearPumpCount = FiltrateNearPumpCount;
                data.FiltrateNearHead = FiltrateNearHead;
                data.FiltrateNearMotorFactor = FiltrateNearMotorFactor;
                data.FiltrateNearEfficiency = FiltrateNearEfficiency;
                data.FiltrateNearSelectedFlow = FiltrateNearSelectedFlow;
                data.FiltrateNearMotorPower = FiltrateNearMotorPower;

                // ---- 事故泵（远） ----
                data.AccidentDensity = AccidentDensity;
                data.AccidentFarPumpCount = AccidentFarPumpCount;
                data.AccidentFarHead = AccidentFarHead;
                data.AccidentFarMotorFactor = AccidentFarMotorFactor;
                data.AccidentFarEfficiency = AccidentFarEfficiency;
                data.AccidentFarSelectedFlow = AccidentFarSelectedFlow;
                data.AccidentFarMotorPower = AccidentFarMotorPower;

                // ---- 事故泵（近） ----
                data.AccidentNearPumpCount = AccidentNearPumpCount;
                data.AccidentNearHead = AccidentNearHead;
                data.AccidentNearMotorFactor = AccidentNearMotorFactor;
                data.AccidentNearEfficiency = AccidentNearEfficiency;
                data.AccidentNearSelectedFlow = AccidentNearSelectedFlow;
                data.AccidentNearMotorPower = AccidentNearMotorPower;

                // ---- 地坑泵 ----
                data.SumpPumpCount = SumpPumpCount;
                data.SumpHead = SumpHead;
                data.SumpMotorFactor = SumpMotorFactor;
                data.SumpEfficiency = SumpEfficiency;
                data.SumpSelectedFlow = SumpSelectedFlow;
                data.SumpMotorPower = SumpMotorPower;

                string json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });

                string filePath = GetDataFilePath();
                string dir = Path.GetDirectoryName(filePath);

                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                File.WriteAllText(filePath, json);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"保存数据失败: {ex.Message}");
            }
        }

        #endregion

        #region ==================== 构造函数 ====================
        public PipeCalcViewModel()
        {
            LoadData();
            CalculateAll();
        }
        #endregion

        #region ==================== 插入表格命令 ====================
        /// <summary>
        /// 导出设备表
        /// </summary>
        public ICommand ExportDeviceTablesCommand => new RelayCommand(ExecuteExportDeviceTables);
        /// <summary>
        /// 导入设备表
        /// </summary>
        private void ExecuteExportDeviceTables()
        {
            var scale = AutoCadHelper.GetScale();
            // 获取当前文档
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                System.Windows.MessageBox.Show("未找到打开的 AutoCAD 文档。");
                return;
            }
            using (DocumentLock docLock = doc.LockDocument()) // 打开当前锁定的文档
            {
                var allRows = GetTableRows();
                if (allRows == null || allRows.Count == 0)
                {
                    System.Windows.MessageBox.Show("没有数据可导出。");
                    return;
                }

                // ---- 按 Group 分组 ----
                var groups = new List<KeyValuePair<string, List<ParameterRow>>>(); // 分组
                string currentGroup = null; // 当前分组
                List<ParameterRow> currentList = null; // 当前分组的行
                // 遍历每一行
                foreach (var row in allRows)
                {
                    if (string.IsNullOrEmpty(row.Group)) // 如果当前行没有分组，则跳过
                        continue;
                    // 如果当前行分组与当前分组不一致，则添加当前分组
                    if (row.Group != currentGroup)
                    {
                        // 添加当前分组
                        if (currentList != null && currentList.Count > 0)
                            groups.Add(new KeyValuePair<string, List<ParameterRow>>(currentGroup, currentList)); // 添加当前分组
                        currentGroup = row.Group; // 更新当前分组
                        currentList = new List<ParameterRow>(); // 创建新的行列表
                    }
                    currentList.Add(row); // 添加行
                }
                // 添加最后一组
                if (currentList != null && currentList.Count > 0)
                    groups.Add(new KeyValuePair<string, List<ParameterRow>>(currentGroup, currentList)); // 添加当前分组
                // 如果没有分组，则跳过
                if (groups.Count == 0)
                {
                    System.Windows.MessageBox.Show("没有可导出的设备分组。");
                    return;
                }

                // ---- 让用户选择第一个插入点 ----
                Editor ed = doc.Editor;
                PromptPointResult ppr = ed.GetPoint("\n请指定第一个设备表格的插入点: ");
                if (ppr.Status != PromptStatus.OK) return;
                // 获取第一个插入点
                Point3d basePoint = ppr.Value;
                const double spacing = 0.8; // 表格之间的水平间距 (mm)

                // 遍历每一组
                int successCount = 0;
                foreach (var group in groups) // 遍历每一组
                {
                    List<ParameterRow> rows = group.Value;  // 当前分组的行
                    if (rows == null || rows.Count == 0) continue; // 如果当前组没有行，则跳过

                    // ---- 计算当前表格的总宽度 ----
                    double tableWidth = GB_NewCadPlus_IV.Helpers.CadTableHelper.GetTableWidth(rows, scale);

                    // ---- 插入表格 ----
                    GB_NewCadPlus_IV.Helpers.CadTableHelper.InsertTable(doc, rows, basePoint);
                    successCount++;

                    // ---- 更新下一个插入点（X 轴向右移动） ----
                    basePoint = new Point3d(basePoint.X + tableWidth * spacing, basePoint.Y, basePoint.Z);
                }

                System.Windows.MessageBox.Show($"已成功导出 {successCount} 个设备表格（水平排列）。");
            }

        }
        /// <summary>
        /// 插入计算表格
        /// </summary>
        public ICommand InsertTableCommand => new RelayCommand(ExecuteInsertTable);
        /// <summary>
        /// 插入计算表格
        /// </summary>
        private void ExecuteInsertTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                System.Windows.MessageBox.Show("未找到打开的 AutoCAD 文档。");
                return;
            }

            var rows = GetTableRows(); // 获取计算表的行内数据   
            Editor ed = doc.Editor;
            PromptPointResult ppr = ed.GetPoint("\n请指定表格插入点: ");
            if (ppr.Status != PromptStatus.OK) return;

            CadTableHelper.InsertTable(doc, rows, ppr.Value); // 插入表 
        }
        #endregion


        #region ==================== 脱硫系统输入 ====================
        private double _inletSO2 = 210;
        public double InletSO2
        {
            get => _inletSO2;
            set
            {
                if (_inletSO2 != value)
                {
                    _inletSO2 = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _outletSO2 = 35;
        public double OutletSO2
        {
            get => _outletSO2;
            set
            {
                if (_outletSO2 != value)
                {
                    _outletSO2 = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _atmPressure = 93090;
        public double AtmPressure
        {
            get => _atmPressure;
            set
            {
                if (_atmPressure != value)
                {
                    _atmPressure = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _outletTemp = 30;
        public double OutletTemp
        {
            get => _outletTemp;
            set
            {
                if (_outletTemp != value)
                {
                    _outletTemp = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _standardFlow = 1735000;
        public double StandardFlow
        {
            get => _standardFlow;
            set
            {
                if (_standardFlow != value)
                {
                    _standardFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _stackVelocity = 15.4;
        public double StackVelocity
        {
            get => _stackVelocity;
            set
            {
                if (_stackVelocity != value)
                {
                    _stackVelocity = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _equipCount = 2;
        public int EquipCount
        {
            get => _equipCount;
            set
            {
                if (_equipCount != value)
                {
                    _equipCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        // Ca/S 比例
        private double _caSRatio = 1.03;
        public double CaSRatio { get => _caSRatio; set { if (_caSRatio != value) { _caSRatio = value; OnPropertyChanged(); CalculateAll(); SaveData(); } } }

        // 循环氧化区直径（输出）
        private double _oxidationDiameter;
        public double OxidationDiameter { get => _oxidationDiameter; private set { if (_oxidationDiameter != value) { _oxidationDiameter = value; OnPropertyChanged(); } } }

        
        #endregion

        #region ==================== 脱硫系统输出 ====================
        private double _actualFlow;
        /// <summary>
        /// 实际流量
        /// </summary>
        public double ActualFlow
        {
            get => _actualFlow;
            private set
            {
                if (_actualFlow != value)  // ← 只在实际值变化时才触发通知
                {
                    _actualFlow = value;
                    // 添加这行，输出到 Visual Studio 输出窗口
                    System.Diagnostics.Debug.WriteLine($"ActualFlow setter 被调用，新值 = {value}");
                    OnPropertyChanged();
                }
            }
        }
      
        private double _stackDiameter;
        /// <summary>
        /// 循环diameter（默认取最近500倍数）
        /// </summary>
        public double StackDiameter
        {
            get => _stackDiameter;
            private set
            {
                if (_stackDiameter != value) // ← 添加这行
                {
                    _stackDiameter = value;
                    OnPropertyChanged();
                    // ★ 同步更新取值（默认取最近500倍数）
                    StackDiameterAdopted = Math.Ceiling(_stackDiameter / 500.0) * 500.0;
                }
            }
        }
        private double _stackDiameterAdopted ;
        /// <summary>
        /// 
        /// </summary>
        public double StackDiameterAdopted
        {
            get => _stackDiameterAdopted;
            set
            {
                // 向上取整到最近的 500 倍数
                // 500→500, 501→1000, 1501→2000, 2001→2500
                double rounded = Math.Ceiling(value / 500.0) * 500.0;

                if (_stackDiameterAdopted != rounded)
                {
                    _stackDiameterAdopted = rounded;
                    OnPropertyChanged();

                    // 如果取值变化后需要联动其他计算（如核算流速等），在这里调用
                    // Recalculate();
                }
            }
        }
        private double _singleUnitFlow;
        public double SingleUnitFlow
        {
            get => _singleUnitFlow;
            private set
            {
                if (_singleUnitFlow != value) // ← 添加这行
                {
                    _singleUnitFlow = value;
                    OnPropertyChanged();
                }
            }
        }

        private double _so2Moles;

        public double So2Moles
        {
            get => _so2Moles;
            private set
            {
                if (_so2Moles != value) // ← 添加这行
                {
                    _so2Moles = value;
                    OnPropertyChanged();
                }
            }
        }

        #endregion

        #region ==================== 循环泵输入 ====================
        private double _liquidGasRatio = 2.2;
        public double LiquidGasRatio
        {
            get => _liquidGasRatio;
            set
            {
                if (_liquidGasRatio != value)
                {
                    _liquidGasRatio = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _residenceTime = 4.35;
        public double ResidenceTime
        {
            get => _residenceTime;
            set
            {
                if (_residenceTime != value)
                {
                    _residenceTime = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _sprayLayers = 1;
        public int SprayLayers
        {
            get => _sprayLayers;
            set
            {
                if (_sprayLayers != value)
                {
                    _sprayLayers = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _pumpHead = 20;
        public double PumpHead
        {
            get => _pumpHead;
            set
            {
                if (_pumpHead != value)
                {
                    _pumpHead = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _motorFactor = 1.15;
        public double MotorFactor
        {
            get => _motorFactor;
            set
            {
                if (_motorFactor != value)
                {
                    _motorFactor = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _selectedPumpFlow = 1950;
        public double SelectedPumpFlow
        {
            get => _selectedPumpFlow;
            set
            {
                if (_selectedPumpFlow != value)
                {
                    _selectedPumpFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        // 备用泵台数
        private int _pumpSpareCount = 1;
        public int PumpSpareCount { get => _pumpSpareCount; set { if (_pumpSpareCount != value) { _pumpSpareCount = value; OnPropertyChanged(); SaveData(); } } }

        #endregion

        #region ==================== 循环泵输出 ====================
        private double _totalSlurryCirc;
        public double TotalSlurryCirc
        {
            get => _totalSlurryCirc;
            private set
            {
                if (_totalSlurryCirc != value)
                {
                    _totalSlurryCirc = value;
                    OnPropertyChanged();
                }
            }
        }

        private double _oxidationVolume;
        public double OxidationVolume
        {
            get => _oxidationVolume;
            private set
            {
                if (_oxidationVolume != value)
                {
                    _oxidationVolume = value;
                    OnPropertyChanged();
                }
            }
        }

        private double _oxidationHeight;
        public double OxidationHeight
        {
            get => _oxidationHeight;
            private set
            {
                if (_oxidationHeight != value)
                {
                    _oxidationHeight = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _runningPumps;
        public int RunningPumps
        {
            get => _runningPumps;
            private set
            {
                if (_runningPumps != value)
                {
                    _runningPumps = value;
                    OnPropertyChanged();
                }
            }
        }

        private double _singlePumpFlowCalc;
        public double SinglePumpFlowCalc
        {
            get => _singlePumpFlowCalc;
            private set
            {
                if (_singlePumpFlowCalc != value)
                {
                    _singlePumpFlowCalc = value;
                    OnPropertyChanged();
                }
            }
        }

        private double _calcPower;
        public double CalcPower
        {
            get => _calcPower;
            private set
            {
                if (_calcPower != value)
                {
                    _calcPower = value;
                    OnPropertyChanged();
                }
            }
        }

        private double _recommendedMotorPower;
        public double RecommendedMotorPower
        {
            get => _recommendedMotorPower;
            private set
            {
                if (_recommendedMotorPower != value)
                {
                    _recommendedMotorPower = value;
                    OnPropertyChanged();
                }
            }
        }

        // 注意：MotorPower 是输出属性（允许用户手动选型），但也是用户可编辑的，所以也视为输入
        private double _motorPower = 185;
        public double MotorPower
        {
            get => _motorPower;
            set
            {
                if (_motorPower != value)
                {
                    _motorPower = value;
                    OnPropertyChanged();
                    SaveData();  // 仅保存，不需要重新计算（因为它是选型值，不影响其他计算）
                }
            }
        }
        #endregion

        #region ==================== 管道尺寸输入 ====================
        private double _inletVelocity = 1.5;
        public double InletVelocity
        {
            get => _inletVelocity;
            set
            {
                if (_inletVelocity != value)
                {
                    _inletVelocity = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _inletDiameter = 350;
        public double InletDiameter
        {
            get => _inletDiameter;
            set
            {
                if (_inletDiameter != value)
                {
                    _inletDiameter = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _outletVelocity = 1.8;
        public double OutletVelocity
        {
            get => _outletVelocity;
            set
            {
                if (_outletVelocity != value)
                {
                    _outletVelocity = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _outletDiameter = 300;
        public double OutletDiameter
        {
            get => _outletDiameter;
            set
            {
                if (_outletDiameter != value)
                {
                    _outletDiameter = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }
        #endregion

        #region ==================== 管道尺寸输出 ====================
        private double _calcInletDiameter;
        public double CalcInletDiameter { get => _calcInletDiameter; private set { _calcInletDiameter = value; OnPropertyChanged(); } }

        private double _checkInletVelocity;
        public double CheckInletVelocity { get => _checkInletVelocity; private set { _checkInletVelocity = value; OnPropertyChanged(); } }

        private double _calcOutletDiameter;
        public double CalcOutletDiameter { get => _calcOutletDiameter; private set { _calcOutletDiameter = value; OnPropertyChanged(); } }

        private double _checkOutletVelocity;
        public double CheckOutletVelocity { get => _checkOutletVelocity; private set { _checkOutletVelocity = value; OnPropertyChanged(); } }
        #endregion

        #region ==================== 氧化风机输入 ====================
        private double _fanHead = 80; // kPa
        public double FanHead
        {
            get => _fanHead;
            set
            {
                if (_fanHead != value)
                {
                    _fanHead = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _fanMotorFactor = 1.15;
        public double FanMotorFactor
        {
            get => _fanMotorFactor;
            set
            {
                if (_fanMotorFactor != value)
                {
                    _fanMotorFactor = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _selectedFanFlow = 16; // m3/h
        public double SelectedFanFlow
        {
            get => _selectedFanFlow;
            set
            {
                if (_selectedFanFlow != value)
                {
                    _selectedFanFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        // 配置台数
        private int _fanRunningCount = 1;
        public int FanRunningCount { get => _fanRunningCount; set { if (_fanRunningCount != value) { _fanRunningCount = value; OnPropertyChanged(); SaveData(); } } }

        // 备用台数
        private int _fanSpareCount = 1;
        public int FanSpareCount { get => _fanSpareCount; set { if (_fanSpareCount != value) { _fanSpareCount = value; OnPropertyChanged(); SaveData(); } } }

        // 氧化风出口管道相关
        private double _fanOutletVelocity = 8;
        public double FanOutletVelocity { get => _fanOutletVelocity; set { if (_fanOutletVelocity != value) { _fanOutletVelocity = value; OnPropertyChanged(); CalculateAll(); SaveData(); } } }

        private double _fanOutletDiameter = 200;
        public double FanOutletDiameter { get => _fanOutletDiameter; set { if (_fanOutletDiameter != value) { _fanOutletDiameter = value; OnPropertyChanged(); CalculateAll(); SaveData(); } } }

        private double _calcFanOutletDiameter;
        public double CalcFanOutletDiameter { get => _calcFanOutletDiameter; private set { if (_calcFanOutletDiameter != value) { _calcFanOutletDiameter = value; OnPropertyChanged(); } } }

        private double _checkFanOutletVelocity;
        public double CheckFanOutletVelocity { get => _checkFanOutletVelocity; private set { if (_checkFanOutletVelocity != value) { _checkFanOutletVelocity = value; OnPropertyChanged(); } } }

        // 介质、使用条件（用于ComboBox）
        public List<string> MediumOptions { get; } = new List<string> { "循环浆液", "氧化风", "石灰石浆液", "石膏浆液", "滤液水" };
        public List<string> ConditionOptions { get; } = new List<string> { "泵前", "泵后" };

        private string _inletMedium = "循环浆液";
        public string InletMedium { get => _inletMedium; set { if (_inletMedium != value) { _inletMedium = value; OnPropertyChanged(); SaveData(); } } }

        private string _inletCondition = "泵前";
        public string InletCondition { get => _inletCondition; set { if (_inletCondition != value) { _inletCondition = value; OnPropertyChanged(); SaveData(); } } }

        private string _outletMedium = "循环浆液";
        public string OutletMedium { get => _outletMedium; set { if (_outletMedium != value) { _outletMedium = value; OnPropertyChanged(); SaveData(); } } }

        private string _outletCondition = "泵后";
        public string OutletCondition { get => _outletCondition; set { if (_outletCondition != value) { _outletCondition = value; OnPropertyChanged(); SaveData(); } } }

        private string _fanOutletMedium = "氧化风";
        public string FanOutletMedium { get => _fanOutletMedium; set { if (_fanOutletMedium != value) { _fanOutletMedium = value; OnPropertyChanged(); SaveData(); } } }

        private string _fanOutletCondition = "泵后";
        public string FanOutletCondition { get => _fanOutletCondition; set { if (_fanOutletCondition != value) { _fanOutletCondition = value; OnPropertyChanged(); SaveData(); } } }

        #endregion

        #region ==================== 氧化风机输出 ====================


        private double _requiredAir;
        public double RequiredAir { get => _requiredAir; private set { _requiredAir = value; OnPropertyChanged(); } }

        private double _requiredPressure;
        public double RequiredPressure { get => _requiredPressure; private set { _requiredPressure = value; OnPropertyChanged(); } }

        private double _fanCalcPower;
        public double FanCalcPower { get => _fanCalcPower; private set { _fanCalcPower = value; OnPropertyChanged(); } }

        private double _fanRecommendedPower;
        public double FanRecommendedPower { get => _fanRecommendedPower; private set { _fanRecommendedPower = value; OnPropertyChanged(); } }

        private double _fanMotorPower = 30;
        public double FanMotorPower
        {
            get => _fanMotorPower;
            set
            {
                if (_fanMotorPower != value)
                {
                    _fanMotorPower = value;
                    OnPropertyChanged();
                    SaveData();
                }
            }
        }
        #endregion

        #region ==================== 石灰石浆液泵（远）输入 ====================
        private double _limePurity = 0.9;
        public double LimePurity
        {
            get => _limePurity;
            set
            {
                if (_limePurity != value)
                {
                    _limePurity = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _slurryDensity = 1200;
        public double SlurryDensity
        {
            get => _slurryDensity;
            set
            {
                if (_slurryDensity != value)
                {
                    _slurryDensity = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _farSystemCount = 2;
        public int FarSystemCount
        {
            get => _farSystemCount;
            set
            {
                if (_farSystemCount != value)
                {
                    _farSystemCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _farRunHours = 8;
        public int FarRunHours
        {
            get => _farRunHours;
            set
            {
                if (_farRunHours != value)
                {
                    _farRunHours = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _farPumpCount = 1;
        public int FarPumpCount
        {
            get => _farPumpCount;
            set
            {
                if (_farPumpCount != value)
                {
                    _farPumpCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _farPumpHead = 35;
        public double FarPumpHead
        {
            get => _farPumpHead;
            set
            {
                if (_farPumpHead != value)
                {
                    _farPumpHead = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _farMotorFactor = 1.3;
        public double FarMotorFactor
        {
            get => _farMotorFactor;
            set
            {
                if (_farMotorFactor != value)
                {
                    _farMotorFactor = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _farPumpEfficiency = 0.5;
        public double FarPumpEfficiency
        {
            get => _farPumpEfficiency;
            set
            {
                if (_farPumpEfficiency != value)
                {
                    _farPumpEfficiency = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _farSelectedFlow = 15;
        public double FarSelectedFlow
        {
            get => _farSelectedFlow;
            set
            {
                if (_farSelectedFlow != value)
                {
                    _farSelectedFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }


        #endregion

        #region ==================== 石灰石浆液泵（远）输出 ====================
        private double _limeConsumption;
        public double LimeConsumption { get => _limeConsumption; private set { _limeConsumption = value; OnPropertyChanged(); } }

        private double _farCalcFlow;
        public double FarCalcFlow { get => _farCalcFlow; private set { _farCalcFlow = value; OnPropertyChanged(); } }

        private double _farCalcPower;
        public double FarCalcPower { get => _farCalcPower; private set { _farCalcPower = value; OnPropertyChanged(); } }

        private double _farRecommendedPower;
        public double FarRecommendedPower { get => _farRecommendedPower; private set { _farRecommendedPower = value; OnPropertyChanged(); } }

        private double _farMotorPower = 5.5;
        public double FarMotorPower
        {
            get => _farMotorPower;
            set
            {
                if (_farMotorPower != value)
                {
                    _farMotorPower = value;
                    OnPropertyChanged();
                    SaveData();
                }
            }
        }
        #endregion

        #region ==================== 石灰石浆液泵（近）输入 ====================
        private int _nearSystemCount = 1;
        public int NearSystemCount
        {
            get => _nearSystemCount;
            set
            {
                if (_nearSystemCount != value)
                {
                    _nearSystemCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _nearRunHours = 8;
        public int NearRunHours
        {
            get => _nearRunHours;
            set
            {
                if (_nearRunHours != value)
                {
                    _nearRunHours = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _nearPumpCount = 1;
        public int NearPumpCount
        {
            get => _nearPumpCount;
            set
            {
                if (_nearPumpCount != value)
                {
                    _nearPumpCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _nearPumpHead = 15;
        public double NearPumpHead
        {
            get => _nearPumpHead;
            set
            {
                if (_nearPumpHead != value)
                {
                    _nearPumpHead = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _nearMotorFactor = 1.3;
        public double NearMotorFactor
        {
            get => _nearMotorFactor;
            set
            {
                if (_nearMotorFactor != value)
                {
                    _nearMotorFactor = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _nearPumpEfficiency = 0.5;
        public double NearPumpEfficiency
        {
            get => _nearPumpEfficiency;
            set
            {
                if (_nearPumpEfficiency != value)
                {
                    _nearPumpEfficiency = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _nearSelectedFlow = 10;
        public double NearSelectedFlow
        {
            get => _nearSelectedFlow;
            set
            {
                if (_nearSelectedFlow != value)
                {
                    _nearSelectedFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }


        #endregion

        #region ==================== 石灰石浆液泵（近）输出 ====================
        private double _nearCalcFlow;
        public double NearCalcFlow { get => _nearCalcFlow; private set { _nearCalcFlow = value; OnPropertyChanged(); } }

        private double _nearCalcPower;
        public double NearCalcPower { get => _nearCalcPower; private set { _nearCalcPower = value; OnPropertyChanged(); } }

        private double _nearRecommendedPower;
        public double NearRecommendedPower { get => _nearRecommendedPower; private set { _nearRecommendedPower = value; OnPropertyChanged(); } }

        private double _nearMotorPower = 2.2;
        public double NearMotorPower
        {
            get => _nearMotorPower;
            set
            {
                if (_nearMotorPower != value)
                {
                    _nearMotorPower = value;
                    OnPropertyChanged();
                    SaveData();
                }
            }
        }
        #endregion

        #region ==================== 石膏排出泵（远）输入 ====================
        private double _gypsumDensity = 1130;
        public double GypsumDensity
        {
            get => _gypsumDensity;
            set
            {
                if (_gypsumDensity != value)
                {
                    _gypsumDensity = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _farGypsumSystemCount = 2;
        public int FarGypsumSystemCount
        {
            get => _farGypsumSystemCount;
            set
            {
                if (_farGypsumSystemCount != value)
                {
                    _farGypsumSystemCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _farGypsumRunHours = 8;
        public int FarGypsumRunHours
        {
            get => _farGypsumRunHours;
            set
            {
                if (_farGypsumRunHours != value)
                {
                    _farGypsumRunHours = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _farGypsumPumpCount = 1;
        public int FarGypsumPumpCount
        {
            get => _farGypsumPumpCount;
            set
            {
                if (_farGypsumPumpCount != value)
                {
                    _farGypsumPumpCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _farGypsumHead = 20;
        public double FarGypsumHead
        {
            get => _farGypsumHead;
            set
            {
                if (_farGypsumHead != value)
                {
                    _farGypsumHead = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _farGypsumMotorFactor = 1.3;
        public double FarGypsumMotorFactor
        {
            get => _farGypsumMotorFactor;
            set
            {
                if (_farGypsumMotorFactor != value)
                {
                    _farGypsumMotorFactor = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _farGypsumEfficiency = 0.5;
        public double FarGypsumEfficiency
        {
            get => _farGypsumEfficiency;
            set
            {
                if (_farGypsumEfficiency != value)
                {
                    _farGypsumEfficiency = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _farGypsumSelectedFlow = 15;
        public double FarGypsumSelectedFlow
        {
            get => _farGypsumSelectedFlow;
            set
            {
                if (_farGypsumSelectedFlow != value)
                {
                    _farGypsumSelectedFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        #endregion

        #region ==================== 石膏排出泵（远）输出 ====================
        private double _gypsumProduction;
        public double GypsumProduction { get => _gypsumProduction; private set { _gypsumProduction = value; OnPropertyChanged(); } }

        private double _gypsumWaterCarry;
        public double GypsumWaterCarry { get => _gypsumWaterCarry; private set { _gypsumWaterCarry = value; OnPropertyChanged(); } }

        private double _gypsumSlurryFlow;
        public double GypsumSlurryFlow { get => _gypsumSlurryFlow; private set { _gypsumSlurryFlow = value; OnPropertyChanged(); } }

        private double _farGypsumCalcFlow;
        public double FarGypsumCalcFlow { get => _farGypsumCalcFlow; private set { _farGypsumCalcFlow = value; OnPropertyChanged(); } }

        private double _farGypsumCalcPower;
        public double FarGypsumCalcPower { get => _farGypsumCalcPower; private set { _farGypsumCalcPower = value; OnPropertyChanged(); } }

        private double _farGypsumRecommendedPower;
        public double FarGypsumRecommendedPower { get => _farGypsumRecommendedPower; private set { _farGypsumRecommendedPower = value; OnPropertyChanged(); } }

        private double _farGypsumMotorPower = 5.5;
        public double FarGypsumMotorPower
        {
            get => _farGypsumMotorPower;
            set
            {
                if (_farGypsumMotorPower != value)
                {
                    _farGypsumMotorPower = value;
                    OnPropertyChanged();
                    SaveData();
                }
            }
        }
        #endregion

        #region ==================== 石膏排出泵（近）输入 ====================
        private int _nearGypsumSystemCount = 1;
        public int NearGypsumSystemCount
        {
            get => _nearGypsumSystemCount;
            set
            {
                if (_nearGypsumSystemCount != value)
                {
                    _nearGypsumSystemCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _nearGypsumRunHours = 8;
        public int NearGypsumRunHours
        {
            get => _nearGypsumRunHours;
            set
            {
                if (_nearGypsumRunHours != value)
                {
                    _nearGypsumRunHours = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _nearGypsumPumpCount = 1;
        public int NearGypsumPumpCount
        {
            get => _nearGypsumPumpCount;
            set
            {
                if (_nearGypsumPumpCount != value)
                {
                    _nearGypsumPumpCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _nearGypsumHead = 10;
        public double NearGypsumHead
        {
            get => _nearGypsumHead;
            set
            {
                if (_nearGypsumHead != value)
                {
                    _nearGypsumHead = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _nearGypsumMotorFactor = 1.3;
        public double NearGypsumMotorFactor
        {
            get => _nearGypsumMotorFactor;
            set
            {
                if (_nearGypsumMotorFactor != value)
                {
                    _nearGypsumMotorFactor = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _nearGypsumEfficiency = 0.5;
        public double NearGypsumEfficiency
        {
            get => _nearGypsumEfficiency;
            set
            {
                if (_nearGypsumEfficiency != value)
                {
                    _nearGypsumEfficiency = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _nearGypsumSelectedFlow = 15;
        public double NearGypsumSelectedFlow
        {
            get => _nearGypsumSelectedFlow;
            set
            {
                if (_nearGypsumSelectedFlow != value)
                {
                    _nearGypsumSelectedFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        #endregion

        #region ==================== 石膏排出泵（近）输出 ====================
        private double _nearGypsumCalcFlow;
        public double NearGypsumCalcFlow { get => _nearGypsumCalcFlow; private set { _nearGypsumCalcFlow = value; OnPropertyChanged(); } }

        private double _nearGypsumCalcPower;
        public double NearGypsumCalcPower { get => _nearGypsumCalcPower; private set { _nearGypsumCalcPower = value; OnPropertyChanged(); } }

        private double _nearGypsumRecommendedPower;
        public double NearGypsumRecommendedPower { get => _nearGypsumRecommendedPower; private set { _nearGypsumRecommendedPower = value; OnPropertyChanged(); } }


        private double _nearGypsumMotorPower = 2.2;
        public double NearGypsumMotorPower
        {
            get => _nearGypsumMotorPower;
            set
            {
                if (_nearGypsumMotorPower != value)
                {
                    _nearGypsumMotorPower = value;
                    OnPropertyChanged();
                    SaveData();
                }
            }
        }
        #endregion

        #region ==================== 缓冲泵输入 ====================
        // 将 BufferRunHours 从 int 改为 double，因为 Excel 中是 5.7
        private double _bufferRunHours = 5.7;
        public double BufferRunHours
        {
            get => _bufferRunHours;
            set
            {
                if (_bufferRunHours != value)
                {
                    _bufferRunHours = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _bufferSystemCount = 3;
        public int BufferSystemCount
        {
            get => _bufferSystemCount;
            set
            {
                if (_bufferSystemCount != value)
                {
                    _bufferSystemCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _bufferPumpCount = 1;
        public int BufferPumpCount
        {
            get => _bufferPumpCount;
            set
            {
                if (_bufferPumpCount != value)
                {
                    _bufferPumpCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _bufferHead = 40;
        public double BufferHead
        {
            get => _bufferHead;
            set
            {
                if (_bufferHead != value)
                {
                    _bufferHead = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _bufferMotorFactor = 1.3;
        public double BufferMotorFactor
        {
            get => _bufferMotorFactor;
            set
            {
                if (_bufferMotorFactor != value)
                {
                    _bufferMotorFactor = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _bufferEfficiency = 0.5;
        public double BufferEfficiency
        {
            get => _bufferEfficiency;
            set
            {
                if (_bufferEfficiency != value)
                {
                    _bufferEfficiency = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _bufferSelectedFlow = 55;
        public double BufferSelectedFlow
        {
            get => _bufferSelectedFlow;
            set
            {
                if (_bufferSelectedFlow != value)
                {
                    _bufferSelectedFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }


        #endregion

        #region ==================== 缓冲泵输出 ====================
        private double _bufferCalcFlow;
        public double BufferCalcFlow { get => _bufferCalcFlow; private set { _bufferCalcFlow = value; OnPropertyChanged(); } }

        private double _bufferCalcPower;
        public double BufferCalcPower { get => _bufferCalcPower; private set { _bufferCalcPower = value; OnPropertyChanged(); } }

        private double _bufferRecommendedPower;
        public double BufferRecommendedPower { get => _bufferRecommendedPower; private set { _bufferRecommendedPower = value; OnPropertyChanged(); } }

        private double _bufferMotorPower = 18.5;
        public double BufferMotorPower
        {
            get => _bufferMotorPower;
            set
            {
                if (_bufferMotorPower != value)
                {
                    _bufferMotorPower = value;
                    OnPropertyChanged();
                    SaveData();
                }
            }
        }
        #endregion

        #region ==================== 滤液水泵（远）输入 ====================
        private double _filtrateDensity = 1000;
        public double FiltrateDensity
        {
            get => _filtrateDensity;
            set
            {
                if (_filtrateDensity != value)
                {
                    _filtrateDensity = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _filtrateFarSystemCount = 2;
        public int FiltrateFarSystemCount
        {
            get => _filtrateFarSystemCount;
            set
            {
                if (_filtrateFarSystemCount != value)
                {
                    _filtrateFarSystemCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _filtrateFarPumpCount = 1;
        public int FiltrateFarPumpCount
        {
            get => _filtrateFarPumpCount;
            set
            {
                if (_filtrateFarPumpCount != value)
                {
                    _filtrateFarPumpCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _filtrateFarHead = 35;
        public double FiltrateFarHead
        {
            get => _filtrateFarHead;
            set
            {
                if (_filtrateFarHead != value)
                {
                    _filtrateFarHead = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _filtrateFarMotorFactor = 1.3;
        public double FiltrateFarMotorFactor
        {
            get => _filtrateFarMotorFactor;
            set
            {
                if (_filtrateFarMotorFactor != value)
                {
                    _filtrateFarMotorFactor = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _filtrateFarEfficiency = 0.5;
        public double FiltrateFarEfficiency
        {
            get => _filtrateFarEfficiency;
            set
            {
                if (_filtrateFarEfficiency != value)
                {
                    _filtrateFarEfficiency = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _filtrateFarSelectedFlow = 50;
        public double FiltrateFarSelectedFlow
        {
            get => _filtrateFarSelectedFlow;
            set
            {
                if (_filtrateFarSelectedFlow != value)
                {
                    _filtrateFarSelectedFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }


        #endregion

        #region ==================== 滤液水泵（远）输出 ====================
        private double _filtrateFarCalcFlow;
        public double FiltrateFarCalcFlow { get => _filtrateFarCalcFlow; private set { _filtrateFarCalcFlow = value; OnPropertyChanged(); } }

        private double _filtrateFarCalcPower;
        public double FiltrateFarCalcPower { get => _filtrateFarCalcPower; private set { _filtrateFarCalcPower = value; OnPropertyChanged(); } }

        private double _filtrateFarRecommendedPower;
        public double FiltrateFarRecommendedPower { get => _filtrateFarRecommendedPower; private set { _filtrateFarRecommendedPower = value; OnPropertyChanged(); } }

        private double _filtrateFarMotorPower = 15;
        public double FiltrateFarMotorPower
        {
            get => _filtrateFarMotorPower;
            set
            {
                if (_filtrateFarMotorPower != value)
                {
                    _filtrateFarMotorPower = value;
                    OnPropertyChanged();
                    SaveData();
                }
            }
        }
        #endregion

        #region ==================== 滤液水泵（近）输入 ====================
        private int _filtrateNearSystemCount = 1;
        public int FiltrateNearSystemCount
        {
            get => _filtrateNearSystemCount;
            set
            {
                if (_filtrateNearSystemCount != value)
                {
                    _filtrateNearSystemCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _filtrateNearPumpCount = 1;
        public int FiltrateNearPumpCount
        {
            get => _filtrateNearPumpCount;
            set
            {
                if (_filtrateNearPumpCount != value)
                {
                    _filtrateNearPumpCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _filtrateNearHead = 15;
        public double FiltrateNearHead
        {
            get => _filtrateNearHead;
            set
            {
                if (_filtrateNearHead != value)
                {
                    _filtrateNearHead = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _filtrateNearMotorFactor = 1.3;
        public double FiltrateNearMotorFactor
        {
            get => _filtrateNearMotorFactor;
            set
            {
                if (_filtrateNearMotorFactor != value)
                {
                    _filtrateNearMotorFactor = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _filtrateNearEfficiency = 0.5;
        public double FiltrateNearEfficiency
        {
            get => _filtrateNearEfficiency;
            set
            {
                if (_filtrateNearEfficiency != value)
                {
                    _filtrateNearEfficiency = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _filtrateNearSelectedFlow = 50;
        public double FiltrateNearSelectedFlow
        {
            get => _filtrateNearSelectedFlow;
            set
            {
                if (_filtrateNearSelectedFlow != value)
                {
                    _filtrateNearSelectedFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }


        #endregion

        #region ==================== 滤液水泵（近）输出 ====================
        private double _filtrateNearCalcFlow;
        public double FiltrateNearCalcFlow { get => _filtrateNearCalcFlow; private set { _filtrateNearCalcFlow = value; OnPropertyChanged(); } }

        private double _filtrateNearCalcPower;
        public double FiltrateNearCalcPower { get => _filtrateNearCalcPower; private set { _filtrateNearCalcPower = value; OnPropertyChanged(); } }

        private double _filtrateNearRecommendedPower;
        public double FiltrateNearRecommendedPower { get => _filtrateNearRecommendedPower; private set { _filtrateNearRecommendedPower = value; OnPropertyChanged(); } }

        private double _filtrateNearMotorPower = 7.5;
        public double FiltrateNearMotorPower
        {
            get => _filtrateNearMotorPower;
            set
            {
                if (_filtrateNearMotorPower != value)
                {
                    _filtrateNearMotorPower = value;
                    OnPropertyChanged();
                    SaveData();
                }
            }
        }
        #endregion

        #region ==================== 事故泵（远）输入 ====================
        private double _accidentDensity = 1150;
        public double AccidentDensity
        {
            get => _accidentDensity;
            set
            {
                if (_accidentDensity != value)
                {
                    _accidentDensity = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private int _accidentFarPumpCount = 1;
        public int AccidentFarPumpCount
        {
            get => _accidentFarPumpCount;
            set
            {
                if (_accidentFarPumpCount != value)
                {
                    _accidentFarPumpCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _accidentFarHead = 35;
        public double AccidentFarHead
        {
            get => _accidentFarHead;
            set
            {
                if (_accidentFarHead != value)
                {
                    _accidentFarHead = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _accidentFarMotorFactor = 1.3;
        public double AccidentFarMotorFactor
        {
            get => _accidentFarMotorFactor;
            set
            {
                if (_accidentFarMotorFactor != value)
                {
                    _accidentFarMotorFactor = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _accidentFarEfficiency = 0.5;
        public double AccidentFarEfficiency
        {
            get => _accidentFarEfficiency;
            set
            {
                if (_accidentFarEfficiency != value)
                {
                    _accidentFarEfficiency = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _accidentFarSelectedFlow = 40;
        public double AccidentFarSelectedFlow
        {
            get => _accidentFarSelectedFlow;
            set
            {
                if (_accidentFarSelectedFlow != value)
                {
                    _accidentFarSelectedFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }


        #endregion

        #region ==================== 事故泵（远）输出 ====================
        private double _accidentCalcFlow;
        public double AccidentCalcFlow { get => _accidentCalcFlow; private set { _accidentCalcFlow = value; OnPropertyChanged(); } }

        private double _accidentFarCalcPower;
        public double AccidentFarCalcPower { get => _accidentFarCalcPower; private set { _accidentFarCalcPower = value; OnPropertyChanged(); } }

        private double _accidentFarRecommendedPower;
        public double AccidentFarRecommendedPower { get => _accidentFarRecommendedPower; private set { _accidentFarRecommendedPower = value; OnPropertyChanged(); } }

        private double _accidentFarMotorPower = 15;
        public double AccidentFarMotorPower
        {
            get => _accidentFarMotorPower;
            set
            {
                if (_accidentFarMotorPower != value)
                {
                    _accidentFarMotorPower = value;
                    OnPropertyChanged();
                    SaveData();
                }
            }
        }
        #endregion

        #region ==================== 事故泵（近）输入 ====================
        private int _accidentNearPumpCount = 1;
        public int AccidentNearPumpCount
        {
            get => _accidentNearPumpCount;
            set
            {
                if (_accidentNearPumpCount != value)
                {
                    _accidentNearPumpCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _accidentNearHead = 15;
        public double AccidentNearHead
        {
            get => _accidentNearHead;
            set
            {
                if (_accidentNearHead != value)
                {
                    _accidentNearHead = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _accidentNearMotorFactor = 1.3;
        public double AccidentNearMotorFactor
        {
            get => _accidentNearMotorFactor;
            set
            {
                if (_accidentNearMotorFactor != value)
                {
                    _accidentNearMotorFactor = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _accidentNearEfficiency = 0.5;
        public double AccidentNearEfficiency
        {
            get => _accidentNearEfficiency;
            set
            {
                if (_accidentNearEfficiency != value)
                {
                    _accidentNearEfficiency = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _accidentNearSelectedFlow = 40;
        public double AccidentNearSelectedFlow
        {
            get => _accidentNearSelectedFlow;
            set
            {
                if (_accidentNearSelectedFlow != value)
                {
                    _accidentNearSelectedFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }


        #endregion

        #region ==================== 事故泵（近）输出 ====================
        private double _accidentNearCalcPower;
        public double AccidentNearCalcPower { get => _accidentNearCalcPower; private set { _accidentNearCalcPower = value; OnPropertyChanged(); } }

        private double _accidentNearRecommendedPower;
        public double AccidentNearRecommendedPower { get => _accidentNearRecommendedPower; private set { _accidentNearRecommendedPower = value; OnPropertyChanged(); } }

        private double _accidentNearMotorPower = 5.5;
        public double AccidentNearMotorPower
        {
            get => _accidentNearMotorPower;
            set
            {
                if (_accidentNearMotorPower != value)
                {
                    _accidentNearMotorPower = value;
                    OnPropertyChanged();
                    SaveData();
                }
            }
        }
        #endregion

        #region ==================== 地坑泵输入 ====================
        private int _sumpPumpCount = 1;
        public int SumpPumpCount
        {
            get => _sumpPumpCount;
            set
            {
                if (_sumpPumpCount != value)
                {
                    _sumpPumpCount = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _sumpHead = 15;
        public double SumpHead
        {
            get => _sumpHead;
            set
            {
                if (_sumpHead != value)
                {
                    _sumpHead = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _sumpMotorFactor = 1.3;
        public double SumpMotorFactor
        {
            get => _sumpMotorFactor;
            set
            {
                if (_sumpMotorFactor != value)
                {
                    _sumpMotorFactor = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _sumpEfficiency = 0.4;
        public double SumpEfficiency
        {
            get => _sumpEfficiency;
            set
            {
                if (_sumpEfficiency != value)
                {
                    _sumpEfficiency = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }

        private double _sumpSelectedFlow = 10;
        public double SumpSelectedFlow
        {
            get => _sumpSelectedFlow;
            set
            {
                if (_sumpSelectedFlow != value)
                {
                    _sumpSelectedFlow = value;
                    OnPropertyChanged();
                    CalculateAll();
                    SaveData();
                }
            }
        }


        #endregion

        #region ==================== 地坑泵输出 ====================
        private double _sumpCalcFlow;
        public double SumpCalcFlow { get => _sumpCalcFlow; private set { _sumpCalcFlow = value; OnPropertyChanged(); } }

        private double _sumpCalcPower;
        public double SumpCalcPower { get => _sumpCalcPower; private set { _sumpCalcPower = value; OnPropertyChanged(); } }

        private double _sumpRecommendedPower;
        public double SumpRecommendedPower { get => _sumpRecommendedPower; private set { _sumpRecommendedPower = value; OnPropertyChanged(); } }

        private double _sumpMotorPower = 2.2;
        public double SumpMotorPower
        {
            get => _sumpMotorPower;
            set
            {
                if (_sumpMotorPower != value)
                {
                    _sumpMotorPower = value;
                    OnPropertyChanged();
                    SaveData();
                }
            }
        }
        #endregion

        #region ==================== 综合计算入口 ====================
        private bool _isCalculating = false;
        /// <summary>
        /// 综合计算入口
        /// </summary>
        public void CalculateAll()
        { // 防止递归调用
            if (_isCalculating)
            {
                System.Diagnostics.Debug.WriteLine("⚠️ 检测到递归调用，已跳过");
                return;
            }

            _isCalculating = true;
            try
            {
                System.Diagnostics.Debug.WriteLine("=== 开始计算 ===");
                CalculateDesulfurization();   // 脱硫系统
                CalculateCirculationPump();   // 循环泵 
                CalculatePipes();             // 管道尺寸 
                CalculateOxidationFan();      // 氧化风机 
                CalculateLimeSlurryPumpFar(); // 石灰石浆液泵（远）
                CalculateLimeSlurryPumpNear();// 石灰石浆液泵（近）
                CalculateGypsumPumpFar();     // 石膏排出泵（远） 
                CalculateGypsumPumpNear();    // 石膏排出泵（近） 
                CalculateBufferPump();        // 缓冲泵 
                CalculateFiltratePumpFar();   // 滤液水泵（远） 
                CalculateFiltratePumpNear();  // 滤液水泵（近） 
                CalculateAccidentPumpFar();   // 事故泵（远） 
                CalculateAccidentPumpNear();  // 事故泵（近） 
                CalculateSumpPump();          // 地坑泵 
                System.Diagnostics.Debug.WriteLine("=== 计算完成 ===");
            }
            finally
            {
                _isCalculating = false;
            }

        }
        #endregion

        #region ==================== 各模块计算函数 ====================

        /// <summary>脱硫系统</summary>
        private void CalculateDesulfurization()
        {
            // 工况烟气量 = 标况烟气量 * 101325 / 大气压 * (273 + 温度) / 273
            ActualFlow = StandardFlow * 101325 / AtmPressure * (273 + OutletTemp) / 273;

            // 烟囱直径 = 0.0188 * sqrt(工况流量 / 出口风速) * 1000
            if (StackVelocity > 0)
                StackDiameter = 0.0188 * Math.Sqrt(ActualFlow / StackVelocity) * 1000;
            else
                StackDiameter = 0;

            // 单台脱硫装置标况烟气量 = 总标况烟气量 / 设备数量
            SingleUnitFlow = StandardFlow / EquipCount;

            // SO2摩尔数 = 标况烟气量 * (入口浓度 - 出口浓度) / 64 / 1000000
            So2Moles = StandardFlow * (InletSO2 - OutletSO2) / 64 / 1000000;

            OxidationDiameter = StackDiameter / 1000; // 循环氧化区直径 = 烟囱直径 / 1000 (m)
        }

        /// <summary>循环泵</summary>
        private void CalculateCirculationPump()
        {
            // 浆液循环总量 = 单台标况烟气量 * 液气比 / 1000
            TotalSlurryCirc = SingleUnitFlow * LiquidGasRatio / 1000;

            // 循环氧化区有效容积 = 循环总量 / 60 * 停留时间
            OxidationVolume = TotalSlurryCirc / 60 * ResidenceTime;

            // 循环氧化区直径 = 烟囱直径 / 1000 (m)
            double diamM = StackDiameter / 1000;

            // 循环氧化区有效高度 = 容积 / (π * r²)
            //OxidationHeight = (diamM > 0) ? OxidationVolume * 4 / (Math.PI * diamM * diamM) : 0;
            OxidationHeight = OxidationVolume * 4 / (Math.PI * diamM * diamM);

            // 配置循环泵台数 = 喷淋层数 * 设备数量
            RunningPumps = SprayLayers * EquipCount;

            // 单台循环泵流量 = 循环总量 / 泵台数
            //SinglePumpFlowCalc = (RunningPumps > 0) ? TotalSlurryCirc / RunningPumps : 0;
            SinglePumpFlowCalc = TotalSlurryCirc / RunningPumps;

            // 计算功率 = 流量 * 扬程 * 9.81 * 密度(1150) / 3600 / 效率(0.8) / 传动效率(0.98) / 1000 * 电机系数
            double density = 1150;
            double efficiency = 0.8;
            double transEff = 0.98;
            CalcPower = SelectedPumpFlow * PumpHead * 9.81 * density / 3600 / efficiency / transEff / 1000 * MotorFactor;
            //CalcPower = SelectedPumpFlow * PumpHead * 9.81 * 1150 / 3600 / 0.8 / 0.98 / 1000 * MotorFactor;
            RecommendedMotorPower = PowerStandardHelper.GetStandardPower(CalcPower);
        }

        /// <summary>管道尺寸</summary>
        private void CalculatePipes()
        {
            double flow = SinglePumpFlowCalc;
            if (flow > 0)
            {
                // 进口管道
                CalcInletDiameter = Math.Sqrt(flow / 3600 / InletVelocity / Math.PI * 4) * 1000;
                CheckInletVelocity = flow / 3600 / (Math.PI * Math.Pow(InletDiameter / 1000, 2) / 4);

                // 出口管道（循环浆液）
                CalcOutletDiameter = Math.Sqrt(flow / 3600 / OutletVelocity / Math.PI * 4) * 1000;
                CheckOutletVelocity = flow / 3600 / (Math.PI * Math.Pow(OutletDiameter / 1000, 2) / 4);
            }

            // 氧化风出口管道
            // Excel: C52*60/3600 = C52/60，即把 m3/h 转换为 m3/min 再转换为 m3/s
            // 所以相当于 (SelectedFanFlow / 60) / 3600 = SelectedFanFlow / 216000
            if (SelectedFanFlow > 0)
            {
                // 方案1：保持与Excel一致
                //double fanFlow = SelectedFanFlow * 60; // 先转换为 m3/min（Excel中的做法）
                //CalcFanOutletDiameter = Math.Sqrt(fanFlow / 3600 / FanOutletVelocity / Math.PI * 4) * 1000;

                // 方案2：更清晰的写法（结果相同）
                double fanFlowM3s = SelectedFanFlow / 3600; // m3/h 直接转 m3/s
                CalcFanOutletDiameter = Math.Sqrt(fanFlowM3s / FanOutletVelocity / Math.PI * 4) * 1000;

                CheckFanOutletVelocity = SelectedFanFlow / 3600 / (Math.PI * Math.Pow(FanOutletDiameter / 1000, 2) / 4);
            }
        }

        /// <summary>氧化风机</summary>
        private void CalculateOxidationFan()
        {
            // 所需氧化风量 = 0.5 * 22.4 * SO2摩尔数 * 0.9 / 0.21 / 0.3 / 60 (Nm3/min)
            RequiredAir = 0.5 * 22.4 * So2Moles * 0.9 / 0.21 / 0.3 / 60;

            // 所需氧化风压 = 9.8 * 氧化区高度 (kPa)
            RequiredPressure = 9.8 * OxidationHeight;

            // 计算功率 = (流量m3/h * 60 * 风压kPa * 1000 * 电机系数) / (3600 * 102 * 效率 * 传动效率 * 9.8)
            double efficiency = 0.9;
            double transEff = 0.98;
            FanCalcPower = (SelectedFanFlow * 60 * FanHead * 1000 * FanMotorFactor) / (3600 * 102 * efficiency * transEff * 9.8);
            FanRecommendedPower = PowerStandardHelper.GetStandardPower(FanCalcPower);
        }

        /// <summary>石灰石浆液泵（远）</summary>
        private void CalculateLimeSlurryPumpFar()
        {
            // 单套系统石灰石消耗量 = SO2摩尔数 * Ca/S * 100 / 纯度
            // ✅ 修正：使用 CaSRatio 属性，而非硬编码 1.03
            LimeConsumption = So2Moles * CaSRatio * 100 / LimePurity;

            // 计算流量 = 消耗量 * 系统数量 / 浓度(0.2) / 密度 / 泵台数 * 24 / 运行时长
            double concentration = 0.2;
            FarCalcFlow = LimeConsumption * FarSystemCount / concentration / SlurryDensity / FarPumpCount * 24 / FarRunHours;

            // 功率 = 选型流量 * 扬程 * 9.81 * 密度 / 3600 / 效率 / 传动效率 / 1000 * 电机系数
            FarCalcPower = FarSelectedFlow * FarPumpHead * 9.81 * SlurryDensity / 3600 / FarPumpEfficiency / 0.98 / 1000 * FarMotorFactor;
            FarRecommendedPower = PowerStandardHelper.GetStandardPower(FarCalcPower);
        }

        /// <summary>石灰石浆液泵（近）</summary>
        private void CalculateLimeSlurryPumpNear()
        {
            // ✅ 这里不需要 Ca/S，因为 LimeConsumption 已经在远型中计算好了
            double concentration = 0.2;
            NearCalcFlow = LimeConsumption * NearSystemCount / concentration / SlurryDensity / NearPumpCount * 24 / NearRunHours;
            NearCalcPower = NearSelectedFlow * NearPumpHead * 9.81 * SlurryDensity / 3600 / NearPumpEfficiency / 0.98 / 1000 * NearMotorFactor;
            NearRecommendedPower = PowerStandardHelper.GetStandardPower(NearCalcPower);
        }

        /// <summary>石膏排出泵（远）</summary>
        private void CalculateGypsumPumpFar()
        {
            // ✅ 修正：使用 CaSRatio 属性
            // 石膏产量 = (石灰石消耗量*(1-纯度) + SO2摩尔数*(Ca/S-1)*100 + SO2摩尔数*172) / 0.9
            GypsumProduction = (LimeConsumption * (1 - LimePurity) + So2Moles * (CaSRatio - 1) * 100 + So2Moles * 172) / 0.9;

            // 石膏携带水量 = 石膏产量 * 0.1 + SO2摩尔数 * 36
            GypsumWaterCarry = GypsumProduction * 0.1 + So2Moles * 36;

            // 石膏浆液产量 = 石膏产量 / 0.2 / 石膏密度 (m3/h)
            double gypsumConc = 0.2;
            GypsumSlurryFlow = GypsumProduction / gypsumConc / GypsumDensity;

            // 流量 = 单套浆液产量 * 24 / 系统数量 / 泵台数 (Excel: =C96*24/C98/C99)
            FarGypsumCalcFlow = GypsumSlurryFlow * 24 / FarGypsumSystemCount / FarGypsumPumpCount;

            // 功率 = 选型流量 * 扬程 * 9.81 * 密度 / 3600 / 效率 / 传动效率 / 1000 * 电机系数
            FarGypsumCalcPower = FarGypsumSelectedFlow * FarGypsumHead * 9.81 * GypsumDensity / 3600 / FarGypsumEfficiency / 0.98 / 1000 * FarGypsumMotorFactor;
            FarGypsumRecommendedPower = PowerStandardHelper.GetStandardPower(FarGypsumCalcPower);
        }

        /// <summary>石膏排出泵（近）</summary>
        private void CalculateGypsumPumpNear()
        {
            // 流量 = 单套浆液产量 * 24 / 系统数量 / 泵台数 (Excel: =C96*24/C110/C111)
            NearGypsumCalcFlow = GypsumSlurryFlow * 24 / NearGypsumSystemCount / NearGypsumPumpCount;

            NearGypsumCalcPower = NearGypsumSelectedFlow * NearGypsumHead * 9.81 * GypsumDensity / 3600 / NearGypsumEfficiency / 0.98 / 1000 * NearGypsumMotorFactor;
            NearGypsumRecommendedPower = PowerStandardHelper.GetStandardPower(NearGypsumCalcPower);
        }

        /// <summary>缓冲泵</summary>
        private void CalculateBufferPump()
        {
            // 流量 = 单套浆液产量 * 系统数量 * 24 / 运行时长 / 泵台数
            // Excel: =C96*C122*24/C121/C123
            BufferCalcFlow = GypsumSlurryFlow * BufferSystemCount * 24 / BufferRunHours / BufferPumpCount;

            BufferCalcPower = BufferSelectedFlow * BufferHead * 9.81 * GypsumDensity / 3600 / BufferEfficiency / 0.98 / 1000 * BufferMotorFactor;
            BufferRecommendedPower = PowerStandardHelper.GetStandardPower(BufferCalcPower);
        }

        /// <summary>滤液水泵（远）</summary>
        private void CalculateFiltratePumpFar()
        {
            // 流量 = (缓冲泵选型流量 * 石膏密度 / 1000 * 0.8 * 0.9 + 5)
            FiltrateFarCalcFlow = (BufferSelectedFlow * GypsumDensity / 1000 * 0.8 * 0.9 + 5);

            FiltrateFarCalcPower = FiltrateFarSelectedFlow * FiltrateFarHead * 9.81 * FiltrateDensity / 3600 / FiltrateFarEfficiency / 0.98 / 1000 * FiltrateFarMotorFactor;
            FiltrateFarRecommendedPower = PowerStandardHelper.GetStandardPower(FiltrateFarCalcPower);
        }

        /// <summary>滤液水泵（近）</summary>
        private void CalculateFiltratePumpNear()
        {
            FiltrateNearCalcFlow = (BufferSelectedFlow * GypsumDensity / 1000 * 0.8 * 0.9 + 5);

            FiltrateNearCalcPower = FiltrateNearSelectedFlow * FiltrateNearHead * 9.81 * FiltrateDensity / 3600 / FiltrateNearEfficiency / 0.98 / 1000 * FiltrateNearMotorFactor;
            FiltrateNearRecommendedPower = PowerStandardHelper.GetStandardPower(FiltrateNearCalcPower);
        }

        /// <summary>事故泵（远）</summary>
        private void CalculateAccidentPumpFar()
        {
            // 流量 = 氧化区容积 / 8
            AccidentCalcFlow = OxidationVolume / 8;

            AccidentFarCalcPower = AccidentFarSelectedFlow * AccidentFarHead * 9.81 * AccidentDensity / 3600 / AccidentFarEfficiency / 0.98 / 1000 * AccidentFarMotorFactor;
            AccidentFarRecommendedPower = PowerStandardHelper.GetStandardPower(AccidentFarCalcPower);
        }

        /// <summary>事故泵（近）</summary>
        private void CalculateAccidentPumpNear()
        {
            AccidentNearCalcPower = AccidentNearSelectedFlow * AccidentNearHead * 9.81 * AccidentDensity / 3600 / AccidentNearEfficiency / 0.98 / 1000 * AccidentNearMotorFactor;
            AccidentNearRecommendedPower = PowerStandardHelper.GetStandardPower(AccidentNearCalcPower);
        }

        /// <summary>地坑泵</summary>
        private void CalculateSumpPump()
        {
            // 流量 = 3 * 3 * 3 / 8 = 3.375 (m3/h)
            SumpCalcFlow = 27.0 / 8.0;

            SumpCalcPower = SumpSelectedFlow * SumpHead * 9.81 * AccidentDensity / 3600 / SumpEfficiency / 0.98 / 1000 * SumpMotorFactor;
            SumpRecommendedPower = PowerStandardHelper.GetStandardPower(SumpCalcPower);
        }

        #endregion

        #region ==================== 生成表格数据 ====================
        /// <summary>
        /// 获取计算表的行内数据
        /// </summary>
        /// <returns></returns>
        public List<ParameterRow> GetTableRows()
        {
            var rows = new List<ParameterRow>();

            // ==================== 脱硫系统 ====================
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "入口烟气SO2浓度", Value = InletSO2.ToString("F0"), Unit = "mg/Nm3", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "出口烟气SO2浓度", Value = OutletSO2.ToString("F0"), Unit = "mg/Nm3", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "当地大气压", Value = AtmPressure.ToString("F0"), Unit = "Pa", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "脱硫系统出口温度", Value = OutletTemp.ToString("F0"), Unit = "℃", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "单套净化系统出口标况烟气量", Value = StandardFlow.ToString("F0"), Unit = "Nm3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "单套净化系统出口工况烟气量", Value = ActualFlow.ToString("F2"), Unit = "Nm3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "烟囱出口风速", Value = StackVelocity.ToString("F1"), Unit = "m/s", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "烟囱出口直径", Value = StackDiameter.ToString("F0"), Unit = "mm", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "烟囱出口直径取值", Value = StackDiameter.ToString("F0"), Unit = "mm", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "单套净化系统设备数量", Value = EquipCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "单台脱硫装置标况烟气量", Value = SingleUnitFlow.ToString("F2"), Unit = "Nm3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "单套净化系统SO2摩尔数", Value = So2Moles.ToString("F4"), Unit = "kmol/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "Ca/S", Value = CaSRatio.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "液气比", Value = LiquidGasRatio.ToString("F2"), Unit = "L/Nm3", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "脱硫系统", Name = "喷淋层数", Value = SprayLayers.ToString(), Unit = "层", Remark = "输入" });

            // ==================== 循环泵 ====================

            rows.Add(new ParameterRow { Group = "循环泵", Name = "浆液循环总量", Value = TotalSlurryCirc.ToString("F2"), Unit = "m3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "停留时间", Value = ResidenceTime.ToString("F2"), Unit = "min", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "循环氧化区有效容积", Value = OxidationVolume.ToString("F2"), Unit = "m3", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "循环氧化区直径", Value = OxidationDiameter.ToString("F2"), Unit = "m", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "循环氧化区有效高度", Value = OxidationHeight.ToString("F2"), Unit = "m", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "配置循环泵台数（连续运行）", Value = RunningPumps.ToString(), Unit = "台", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "备用", Value = PumpSpareCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "单台循环泵流量", Value = SinglePumpFlowCalc.ToString("F0"), Unit = "m3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "流量（选型）", Value = SelectedPumpFlow.ToString("F0"), Unit = "m3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "扬程", Value = PumpHead.ToString("F0"), Unit = "m", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "电机系数", Value = MotorFactor.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "计算功率", Value = CalcPower.ToString("F2"), Unit = "kW", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "推荐电机功率", Value = RecommendedMotorPower.ToString("F0"), Unit = "kW", Remark = "推荐" });
            rows.Add(new ParameterRow { Group = "循环泵", Name = "电机功率", Value = MotorPower.ToString("F0"), Unit = "kW", Remark = "输入" });

            // ==================== 进口管道尺寸 ====================
            rows.Add(new ParameterRow { Group = "进口管道", Name = "介质", Value = InletMedium, Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "进口管道", Name = "使用条件", Value = InletCondition, Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "进口管道", Name = "流速", Value = InletVelocity.ToString("F2"), Unit = "m/s", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "进口管道", Name = "管道尺寸", Value = CalcInletDiameter.ToString("F0"), Unit = "mm", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "进口管道", Name = "取值 DN", Value = InletDiameter.ToString("F0"), Unit = "mm", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "进口管道", Name = "核算流速", Value = CheckInletVelocity.ToString("F2"), Unit = "m/s", Remark = "计算" });

            // ==================== 出口管道（循环浆液） ====================
            rows.Add(new ParameterRow { Group = "出口管道（循环浆液）", Name = "介质", Value = OutletMedium, Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "出口管道（循环浆液）", Name = "使用条件", Value = OutletCondition, Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "出口管道（循环浆液）", Name = "流速", Value = OutletVelocity.ToString("F2"), Unit = "m/s", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "出口管道（循环浆液）", Name = "管道尺寸", Value = CalcOutletDiameter.ToString("F0"), Unit = "mm", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "出口管道（循环浆液）", Name = "取值 DN", Value = OutletDiameter.ToString("F0"), Unit = "mm", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "出口管道（循环浆液）", Name = "核算流速", Value = CheckOutletVelocity.ToString("F2"), Unit = "m/s", Remark = "计算" });

            // ==================== 氧化风机 ====================
            rows.Add(new ParameterRow { Group = "氧化风机", Name = "配置氧化风机台数（连续运行）", Value = FanRunningCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "氧化风机", Name = "备用", Value = FanSpareCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "氧化风机", Name = "所需氧化风量", Value = RequiredAir.ToString("F2"), Unit = "Nm3/min", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "氧化风机", Name = "所需氧化风压", Value = RequiredPressure.ToString("F2"), Unit = "kPa", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "氧化风机", Name = "扬程", Value = FanHead.ToString("F0"), Unit = "kPa", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "氧化风机", Name = "电机系数", Value = FanMotorFactor.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "氧化风机", Name = "流量（选型）", Value = SelectedFanFlow.ToString("F0"), Unit = "m3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "氧化风机", Name = "计算功率", Value = FanCalcPower.ToString("F2"), Unit = "kW", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "氧化风机", Name = "推荐电机功率", Value = FanRecommendedPower.ToString("F0"), Unit = "kW", Remark = "推荐" });
            rows.Add(new ParameterRow { Group = "氧化风机", Name = "电机功率", Value = FanMotorPower.ToString("F0"), Unit = "kW", Remark = "输入" });

            // ==================== 出口管道（氧化风） ====================
            rows.Add(new ParameterRow { Group = "出口管道（氧化风）", Name = "介质", Value = FanOutletMedium, Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "出口管道（氧化风）", Name = "使用条件", Value = FanOutletCondition, Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "出口管道（氧化风）", Name = "流速", Value = FanOutletVelocity.ToString("F2"), Unit = "m/s", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "出口管道（氧化风）", Name = "管道尺寸", Value = CalcFanOutletDiameter.ToString("F0"), Unit = "mm", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "出口管道（氧化风）", Name = "取值 DN", Value = FanOutletDiameter.ToString("F0"), Unit = "mm", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "出口管道（氧化风）", Name = "核算流速", Value = CheckFanOutletVelocity.ToString("F2"), Unit = "m/s", Remark = "计算" });

            // ==================== 石灰石浆液泵（远） ====================
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "石灰石纯度", Value = LimePurity.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "单套系统消耗量", Value = LimeConsumption.ToString("F0"), Unit = "kg/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "浆液平均密度", Value = SlurryDensity.ToString("F0"), Unit = "kg/m3", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "供给远处系统数量", Value = FarSystemCount.ToString(), Unit = "套", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "运行时长", Value = FarRunHours.ToString(), Unit = "h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "配置台数", Value = FarPumpCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "计算单台浆液泵流量", Value = FarCalcFlow.ToString("F2"), Unit = "m3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "流量（选型）", Value = FarSelectedFlow.ToString("F0"), Unit = "m3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "扬程", Value = FarPumpHead.ToString("F0"), Unit = "m", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "电机系数", Value = FarMotorFactor.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "效率", Value = FarPumpEfficiency.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "计算功率", Value = FarCalcPower.ToString("F2"), Unit = "kW", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "推荐电机功率", Value = FarRecommendedPower.ToString("F0"), Unit = "kW", Remark = "推荐" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（远）", Name = "电机功率", Value = FarMotorPower.ToString("F0"), Unit = "kW", Remark = "输入" });

            // ==================== 石灰石浆液泵（近） ====================
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（近）", Name = "供给近处系统数量", Value = NearSystemCount.ToString(), Unit = "套", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（近）", Name = "运行时长", Value = NearRunHours.ToString(), Unit = "h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（近）", Name = "配置台数", Value = NearPumpCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（近）", Name = "计算单台浆液泵流量", Value = NearCalcFlow.ToString("F2"), Unit = "m3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（近）", Name = "流量（选型）", Value = NearSelectedFlow.ToString("F0"), Unit = "m3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（近）", Name = "扬程", Value = NearPumpHead.ToString("F0"), Unit = "m", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（近）", Name = "电机系数", Value = NearMotorFactor.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（近）", Name = "效率", Value = NearPumpEfficiency.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（近）", Name = "计算功率", Value = NearCalcPower.ToString("F2"), Unit = "kW", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（近）", Name = "推荐电机功率", Value = NearRecommendedPower.ToString("F0"), Unit = "kW", Remark = "推荐" });
            rows.Add(new ParameterRow { Group = "石灰石浆液泵（近）", Name = "电机功率", Value = NearMotorPower.ToString("F0"), Unit = "kW", Remark = "输入" });

            // ==================== 石膏排出泵（远） ====================
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "单套系统石膏产量", Value = GypsumProduction.ToString("F0"), Unit = "kg/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "单套系统石膏携带水量", Value = GypsumWaterCarry.ToString("F0"), Unit = "kg/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "石膏堆积密度", Value = "1000", Unit = "kg/m3", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "石膏浆液平均密度", Value = GypsumDensity.ToString("F0"), Unit = "kg/m3", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "单套系统石膏浆液产量", Value = GypsumSlurryFlow.ToString("F2"), Unit = "m3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "远处系统数量", Value = FarGypsumSystemCount.ToString(), Unit = "套", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "运行时长", Value = FarGypsumRunHours.ToString(), Unit = "h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "配置台数", Value = FarGypsumPumpCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "计算单台浆液泵流量", Value = FarGypsumCalcFlow.ToString("F2"), Unit = "m3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "流量（选型）", Value = FarGypsumSelectedFlow.ToString("F0"), Unit = "m3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "扬程", Value = FarGypsumHead.ToString("F0"), Unit = "m", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "电机系数", Value = FarGypsumMotorFactor.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "效率", Value = FarGypsumEfficiency.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "计算功率", Value = FarGypsumCalcPower.ToString("F2"), Unit = "kW", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "推荐电机功率", Value = FarGypsumRecommendedPower.ToString("F0"), Unit = "kW", Remark = "推荐" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（远）", Name = "电机功率", Value = FarGypsumMotorPower.ToString("F0"), Unit = "kW", Remark = "输入" });

            // ==================== 石膏排出泵（近） ====================
            rows.Add(new ParameterRow { Group = "石膏排出泵（近）", Name = "近处系统数量", Value = NearGypsumSystemCount.ToString(), Unit = "套", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（近）", Name = "运行时长", Value = NearGypsumRunHours.ToString(), Unit = "h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（近）", Name = "配置台数", Value = NearGypsumPumpCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（近）", Name = "计算单台浆液泵流量", Value = NearGypsumCalcFlow.ToString("F2"), Unit = "m3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（近）", Name = "流量（选型）", Value = NearGypsumSelectedFlow.ToString("F0"), Unit = "m3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（近）", Name = "扬程", Value = NearGypsumHead.ToString("F0"), Unit = "m", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（近）", Name = "电机系数", Value = NearGypsumMotorFactor.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（近）", Name = "效率", Value = NearGypsumEfficiency.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（近）", Name = "计算功率", Value = NearGypsumCalcPower.ToString("F2"), Unit = "kW", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（近）", Name = "推荐电机功率", Value = NearGypsumRecommendedPower.ToString("F0"), Unit = "kW", Remark = "推荐" });
            rows.Add(new ParameterRow { Group = "石膏排出泵（近）", Name = "电机功率", Value = NearGypsumMotorPower.ToString("F0"), Unit = "kW", Remark = "输入" });

            // ==================== 缓冲泵 ====================
            rows.Add(new ParameterRow { Group = "缓冲泵", Name = "运行时长", Value = BufferRunHours.ToString("F1"), Unit = "h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "缓冲泵", Name = "供给系统数量", Value = BufferSystemCount.ToString(), Unit = "套", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "缓冲泵", Name = "配置台数", Value = BufferPumpCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "缓冲泵", Name = "计算单台浆液泵流量", Value = BufferCalcFlow.ToString("F2"), Unit = "m3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "缓冲泵", Name = "流量（选型）", Value = BufferSelectedFlow.ToString("F0"), Unit = "m3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "缓冲泵", Name = "扬程", Value = BufferHead.ToString("F0"), Unit = "m", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "缓冲泵", Name = "电机系数", Value = BufferMotorFactor.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "缓冲泵", Name = "效率", Value = BufferEfficiency.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "缓冲泵", Name = "计算功率", Value = BufferCalcPower.ToString("F2"), Unit = "kW", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "缓冲泵", Name = "推荐电机功率", Value = BufferRecommendedPower.ToString("F0"), Unit = "kW", Remark = "推荐" });
            rows.Add(new ParameterRow { Group = "缓冲泵", Name = "电机功率", Value = BufferMotorPower.ToString("F0"), Unit = "kW", Remark = "输入" });

            // ==================== 滤液水泵（远） ====================
            rows.Add(new ParameterRow { Group = "滤液水泵（远）", Name = "滤液水密度", Value = FiltrateDensity.ToString("F0"), Unit = "kg/m3", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（远）", Name = "供给系统数量", Value = FiltrateFarSystemCount.ToString(), Unit = "套", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（远）", Name = "配置台数", Value = FiltrateFarPumpCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（远）", Name = "计算单台浆液泵流量", Value = FiltrateFarCalcFlow.ToString("F2"), Unit = "m3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "滤液水泵（远）", Name = "流量（选型）", Value = FiltrateFarSelectedFlow.ToString("F0"), Unit = "m3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（远）", Name = "扬程", Value = FiltrateFarHead.ToString("F0"), Unit = "m", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（远）", Name = "电机系数", Value = FiltrateFarMotorFactor.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（远）", Name = "效率", Value = FiltrateFarEfficiency.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（远）", Name = "计算功率", Value = FiltrateFarCalcPower.ToString("F2"), Unit = "kW", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "滤液水泵（远）", Name = "推荐电机功率", Value = FiltrateFarRecommendedPower.ToString("F0"), Unit = "kW", Remark = "推荐" });
            rows.Add(new ParameterRow { Group = "滤液水泵（远）", Name = "电机功率", Value = FiltrateFarMotorPower.ToString("F0"), Unit = "kW", Remark = "输入" });

            // ==================== 滤液水泵（近） ====================
            rows.Add(new ParameterRow { Group = "滤液水泵（近）", Name = "供给系统数量", Value = FiltrateNearSystemCount.ToString(), Unit = "套", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（近）", Name = "配置台数", Value = FiltrateNearPumpCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（近）", Name = "计算单台浆液泵流量", Value = FiltrateNearCalcFlow.ToString("F2"), Unit = "m3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "滤液水泵（近）", Name = "流量（选型）", Value = FiltrateNearSelectedFlow.ToString("F0"), Unit = "m3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（近）", Name = "扬程", Value = FiltrateNearHead.ToString("F0"), Unit = "m", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（近）", Name = "电机系数", Value = FiltrateNearMotorFactor.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（近）", Name = "效率", Value = FiltrateNearEfficiency.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "滤液水泵（近）", Name = "计算功率", Value = FiltrateNearCalcPower.ToString("F2"), Unit = "kW", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "滤液水泵（近）", Name = "推荐电机功率", Value = FiltrateNearRecommendedPower.ToString("F0"), Unit = "kW", Remark = "推荐" });
            rows.Add(new ParameterRow { Group = "滤液水泵（近）", Name = "电机功率", Value = FiltrateNearMotorPower.ToString("F0"), Unit = "kW", Remark = "输入" });

            // ==================== 事故泵（远） ====================
            rows.Add(new ParameterRow { Group = "事故泵（远）", Name = "事故浆液密度", Value = AccidentDensity.ToString("F0"), Unit = "kg/m3", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "事故泵（远）", Name = "计算单台浆液泵流量", Value = AccidentCalcFlow.ToString("F2"), Unit = "m3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "事故泵（远）", Name = "配置台数", Value = AccidentFarPumpCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "事故泵（远）", Name = "流量（选型）", Value = AccidentFarSelectedFlow.ToString("F0"), Unit = "m3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "事故泵（远）", Name = "扬程", Value = AccidentFarHead.ToString("F0"), Unit = "m", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "事故泵（远）", Name = "电机系数", Value = AccidentFarMotorFactor.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "事故泵（远）", Name = "效率", Value = AccidentFarEfficiency.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "事故泵（远）", Name = "计算功率", Value = AccidentFarCalcPower.ToString("F2"), Unit = "kW", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "事故泵（远）", Name = "推荐电机功率", Value = AccidentFarRecommendedPower.ToString("F0"), Unit = "kW", Remark = "推荐" });
            rows.Add(new ParameterRow { Group = "事故泵（远）", Name = "电机功率", Value = AccidentFarMotorPower.ToString("F0"), Unit = "kW", Remark = "输入" });

            // ==================== 事故泵（近） ====================
            rows.Add(new ParameterRow { Group = "事故泵（近）", Name = "配置台数", Value = AccidentNearPumpCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "事故泵（近）", Name = "流量（选型）", Value = AccidentNearSelectedFlow.ToString("F0"), Unit = "m3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "事故泵（近）", Name = "扬程", Value = AccidentNearHead.ToString("F0"), Unit = "m", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "事故泵（近）", Name = "电机系数", Value = AccidentNearMotorFactor.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "事故泵（近）", Name = "效率", Value = AccidentNearEfficiency.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "事故泵（近）", Name = "计算功率", Value = AccidentNearCalcPower.ToString("F2"), Unit = "kW", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "事故泵（近）", Name = "推荐电机功率", Value = AccidentNearRecommendedPower.ToString("F0"), Unit = "kW", Remark = "推荐" });
            rows.Add(new ParameterRow { Group = "事故泵（近）", Name = "电机功率", Value = AccidentNearMotorPower.ToString("F0"), Unit = "kW", Remark = "输入" });

            // ==================== 地坑泵 ====================
            rows.Add(new ParameterRow { Group = "地坑泵", Name = "计算单台浆液泵流量", Value = SumpCalcFlow.ToString("F2"), Unit = "m3/h", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "地坑泵", Name = "配置台数", Value = SumpPumpCount.ToString(), Unit = "台", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "地坑泵", Name = "流量（选型）", Value = SumpSelectedFlow.ToString("F0"), Unit = "m3/h", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "地坑泵", Name = "扬程", Value = SumpHead.ToString("F0"), Unit = "m", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "地坑泵", Name = "电机系数", Value = SumpMotorFactor.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "地坑泵", Name = "效率", Value = SumpEfficiency.ToString("F2"), Unit = "", Remark = "输入" });
            rows.Add(new ParameterRow { Group = "地坑泵", Name = "计算功率", Value = SumpCalcPower.ToString("F2"), Unit = "kW", Remark = "计算" });
            rows.Add(new ParameterRow { Group = "地坑泵", Name = "推荐电机功率", Value = SumpRecommendedPower.ToString("F0"), Unit = "kW", Remark = "推荐" });
            rows.Add(new ParameterRow { Group = "地坑泵", Name = "电机功率", Value = SumpMotorPower.ToString("F0"), Unit = "kW", Remark = "输入" });

            return rows;
        }

        #endregion
    }
}
