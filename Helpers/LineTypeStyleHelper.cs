using GB_NewCadPlus_IV.FunctionalMethod;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 线型/线宽/线型比例的工具类（中文注释）
    /// 责任：提供自适应 LinetypeScale 计算与批量应用命令，且不更改全局系统变量或内置 linetype 定义
    /// </summary>
    public static class LineTypeStyleHelper
    {
        /// <summary>
        /// 默认文字样式名称（用于 DASH 定义中的字体依赖，实际创建时会检查是否存在，不存在则使用默认字体）
        /// </summary>
        private const string DefaultTextStyleName = "tJText";
        /// <summary>
        /// 默认线型名称（确保存在的 DASH 线型名称，若已存在则不覆盖）
        /// </summary>
        private const string DefaultDashName = "DASH";
        #region EnsureDashLinetype（DBTrans 版本） —— 推荐在项目内优先使用 DBTrans 调用
        /// <summary>
        /// 确保 DASH 线型存在（DBTrans 版本）
        /// 说明：只有在图面中缺失 DASH 时创建，默认不修改全局系统变量
        /// </summary>
        public static ObjectId EnsureDashLinetype(DBTrans tr, double uiScale, string textStyleName = DefaultTextStyleName, bool syncSystemVariables = false, double? ltscaleOverride = null, double? celtscaleOverride = null)
        {
            if (tr == null) throw new ArgumentNullException(nameof(tr));
            double ltScale = ltscaleOverride ?? uiScale;
            if (double.IsNaN(ltScale) || ltScale <= 0) ltScale = 1.0;
            ltScale = Math.Max(0.1, ltScale);

            // 如果已存在 DASH，直接返回（不覆盖）
            if (tr.LinetypeTable.Has(DefaultDashName))
            {
                return tr.LinetypeTable[DefaultDashName];
            }

            // 创建 DASH 定义（简单模式）
            tr.LinetypeTable.Add(DefaultDashName, ltr =>
            {
                ltr.Name = DefaultDashName;
                ltr.AsciiDescription = " - - - - - ";
                ltr.PatternLength = 1.0 * ltScale;
                ltr.NumDashes = 2;
                ltr.SetDashLengthAt(0, 0.6 * ltScale);
                ltr.SetDashLengthAt(1, -0.4 * ltScale);
            });

            if (syncSystemVariables)
            {
                try { Application.SetSystemVariable("LTSCALE", ltscaleOverride ?? ltScale); } catch { }
                try { Application.SetSystemVariable("CELTSCALE", celtscaleOverride ?? 1.0); } catch { }
            }

            return tr.LinetypeTable[DefaultDashName];
        }
        #endregion

        #region EnsureDashLinetype（Transaction 兼容重载） —— 方便在原生 Transaction 场景使用
        /// <summary>
        /// Transaction 兼容版 EnsureDashLinetype：在原生 Transaction 上确保 DASH 存在并返回其 ObjectId
        /// 注：不改变系统变量（除非 syncSystemVariables == true）
        /// </summary>
        public static ObjectId EnsureDashLinetype(Transaction tr, Database db, double uiScale, bool syncSystemVariables = false, double? ltscaleOverride = null, double? celtscaleOverride = null)
        {
            if (tr == null) throw new ArgumentNullException(nameof(tr));
            if (db == null) throw new ArgumentNullException(nameof(db));

            double ltScale = ltscaleOverride ?? uiScale;
            if (double.IsNaN(ltScale) || ltScale <= 0) ltScale = 1.0;
            ltScale = Math.Max(0.1, ltScale);

            var ltt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);
            if (ltt.Has(DefaultDashName))
            {
                return ltt[DefaultDashName];
            }

            ltt.UpgradeOpen();
            // 创建 dash
            var ltrDef = new LinetypeTableRecord
            {
                Name = DefaultDashName,
                AsciiDescription = " - - - - - "
            };
            ltrDef.PatternLength = 1.0 * ltScale;
            ltrDef.NumDashes = 2;
            ltrDef.SetDashLengthAt(0, 0.6 * ltScale);
            ltrDef.SetDashLengthAt(1, -0.4 * ltScale);
            ObjectId newId = ltt.Add(ltrDef);
            tr.AddNewlyCreatedDBObject(ltrDef, true);

            if (syncSystemVariables)
            {
                try { Application.SetSystemVariable("LTSCALE", ltscaleOverride ?? ltScale); } catch { }
                try { Application.SetSystemVariable("CELTSCALE", celtscaleOverride ?? 1.0); } catch { }
            }

            return newId;
        }
        #endregion

        #region LinetypeScale 计算与应用
        /// <summary>
        /// 读取系统变量并计算全局因子（LTSCALE * CELTSCALE * PSLTSCALE）
        /// </summary>
        public static double ComputeGlobalLinetypeScaleFactor()
        {
            return AutoCadHelper.ComputeGlobalLinetypeScaleFactor();
        }

        /// <summary>
        /// 根据期望的单段可视长度计算推荐 LinetypeScale
        /// recommended = desiredPatternLength / globalFactor
        /// </summary>
        public static double ComputeRecommendedLinetypeScale(double desiredPatternLength)
        {
            double globalFactor = ComputeGlobalLinetypeScaleFactor();
            if (desiredPatternLength <= 0) desiredPatternLength = 100.0;
            double recommended = desiredPatternLength / globalFactor;
            if (double.IsNaN(recommended) || recommended <= 0) recommended = 1.0;
            return recommended;
        }

        /// <summary>
        /// 对单个实体应用推荐 LinetypeScale（Transaction 版本）
        /// </summary>
        public static bool ApplyAdaptiveLinetypeScaleToEntity(Transaction tr, ObjectId entId, double desiredPatternLength, string linetypeNameToMatch = DefaultDashName)
        {
            try
            {
                var ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                if (ent == null) return false;
                if (!string.Equals(ent.Linetype, linetypeNameToMatch, StringComparison.OrdinalIgnoreCase)) return false;
                ent.LinetypeScale = ComputeRecommendedLinetypeScale(desiredPatternLength);
                return true;
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\nApplyAdaptiveLinetypeScaleToEntity 失败: {ex.Message}");
                return false;
            }
        }
        #endregion

        #region 命令：NormalizeDashLinetype（原生 Transaction 版本）
        /// <summary>
        /// 命令：NormalizeDashLinetype
        /// 交互式：输入目标单段长度（绘图单位），对选择集（或整个当前空间）中 linetype 为 DASH 的实体应用推荐 LinetypeScale。
        /// </summary>
        [Autodesk.AutoCAD.Runtime.CommandMethod("NormalizeDashLinetype")]
        public static void NormalizeDashLinetype_Command()
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;
                var ed = doc.Editor;

                // 安全提示用户输入正数（回车使用默认 100）
                double desiredLen = AutoCadHelper.PromptForPositiveDouble("\n请输入目标断线段长度（绘图单位），回车使用默认值 100:", 100.0);

                // 先在当前事务中确保 DASH 存在（Transaction 兼容版）
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    // 获取一个可用的 uiScale（优先 WPF/VariableDictionary）
                    double uiScale = AutoCadHelper.GetScale(true);
                    // 确保 DASH 存在（若不存在创建）
                    var dashId = EnsureDashLinetype(tr, doc.Database, uiScale, syncSystemVariables: false);

                    // 处理用户选择或当前空间全部 DASH
                    var selRes = ed.GetSelection();
                    bool useSelection = selRes.Status == PromptStatus.OK;
                    int applied = 0;
                    if (useSelection)
                    {
                        foreach (var id in selRes.Value.GetObjectIds())
                        {
                            try
                            {
                                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                                if (ent == null) continue;
                                // 使用 LinetypeId 比较更稳健
                                if (ent.LinetypeId == dashId || string.Equals(ent.Linetype, DefaultDashName, StringComparison.OrdinalIgnoreCase))
                                {
                                    // 升级为写入并设置 LinetypeScale
                                    var entW = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                                    if (entW != null)
                                    {
                                        entW.LinetypeScale = ComputeRecommendedLinetypeScale(desiredLen);
                                        applied++;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        var btr = (BlockTableRecord)tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForRead);
                        foreach (ObjectId id in btr)
                        {
                            try
                            {
                                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                                if (ent == null) continue;
                                if (ent.LinetypeId == dashId || string.Equals(ent.Linetype, DefaultDashName, StringComparison.OrdinalIgnoreCase))
                                {
                                    var entW = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                                    if (entW != null)
                                    {
                                        entW.LinetypeScale = ComputeRecommendedLinetypeScale(desiredLen);
                                        applied++;
                                    }
                                }
                            }
                            catch { }
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage($"\n已应用 LinetypeScale 到 {applied} 个 DASH 实体（目标段长: {desiredLen}）。");
                    LogManager.Instance.LogInfo($"\nNormalizeDashLinetype 完成: applied={applied}, desiredLen={desiredLen}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\nNormalizeDashLinetype 执行失败: {ex.Message}");
                try { Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage($"\n错误: {ex.Message}"); } catch { }
            }
        }
        #endregion

        #region 辅助：恢复默认系统变量（可选）
        /// <summary>
        /// 可选：将 LTSCALE/CELTSCALE 恢复为 1.0（谨慎调用）
        /// </summary>
        public static void RestoreDefaultLtscaleCeltscale()
        {
            try { Application.SetSystemVariable("LTSCALE", 1.0); } catch { }
            try { Application.SetSystemVariable("CELTSCALE", 1.0); } catch { }
            AutoCadHelper.LogWithSafety("已尝试将 LTSCALE/CELTSCALE 恢复为 1.0（如权限允许）。");
        }
        #endregion
    }
}
