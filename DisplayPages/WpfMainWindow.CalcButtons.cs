using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.UniFiedStandards;
using System;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace GB_NewCadPlus_IV
{
    public partial class WpfMainWindow
    {
        /// <summary>
        /// 点击“重载CSV”按钮
        /// </summary>
        private void 重载CSV按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ReloadCalcCsvTables(true); // 强制重新加载CSV
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"重载 CSV 失败: {ex.Message}");
                System.Windows.MessageBox.Show(
                    $"重载 CSV 失败：{ex.Message}",
                    "错误",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void 转换CSV按钮_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var ofd = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "选择总表Excel",
                    Filter = "Excel 文件 (*.xlsx)|*.xlsx",
                    FileName = "_00_脱硫系统计算_总表.xlsx",
                    InitialDirectory = @"D:\03-客户文件\沈阳铝镁院\模板"
                };

                if (ofd.ShowDialog() != true)
                    return;
                // 加载Excel并生成页面
                LoadFromMasterExcelAndBuildUi(ofd.FileName, exportCsv: true);

                MessageBox.Show("已按Excel动态生成页面并导出总CSV。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"按Excel转换失败: {ex.Message}");
                MessageBox.Show($"转换失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 插入计算表：将当前“计算数据表”整体生成临时DWG并插入当前图纸空间
        /// </summary>
        private void 插入计算表_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 先确保当前数据已完成一次联动重算，避免插入旧值
                try { RecalculateAllCalcCsvTables(); } catch { /* 若重算失败，后续按当前值继续 */ }

                // 1) 生成临时DWG
                var (dwgPath, blockName) = BuildCalcTablesTempDwg();

                if (string.IsNullOrWhiteSpace(dwgPath) || !File.Exists(dwgPath))
                {
                    MessageBox.Show("生成临时计算表文件失败。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 2) 获取插入点（在当前活动文档中交互拾取）
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("未找到活动的 AutoCAD 文档。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                Autodesk.AutoCAD.Geometry.Point3d insertPoint;
                using (doc.LockDocument())
                {
                    var ppr = doc.Editor.GetPoint("\n请选择“计算表整体”插入点：");
                    if (ppr.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        return;

                    insertPoint = ppr.Value;
                }

                // 3) 插入到当前空间（复用现有统一插入能力）
                var insertedId = AutoCadHelper.InsertBlockFromExternalDwg(dwgPath, blockName, insertPoint);
                if (insertedId == Autodesk.AutoCAD.DatabaseServices.ObjectId.Null)
                {
                    MessageBox.Show("插入失败：未能将计算表导入当前图纸。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                MessageBox.Show("计算表已整体插入当前图纸空间。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
                try
                {
                    // ... 现有插入逻辑成功后
                    TryDeleteTempFile(dwgPath);
                }
                catch
                {
                    // 忽略清理异常，不影响主流程
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogError($"插入计算表失败: {ex.Message}");
                MessageBox.Show($"插入计算表失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
    }
}
