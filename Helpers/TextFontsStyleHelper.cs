using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_IV.Helpers
{/// <summary>
 /// 辅助类：处理字体样式
 /// 说明：统一使用 DrawingScaleService 获取当前视图比例并据此计算最终字高，保证与项目其它地方比例一致。
 /// </summary>
    public static class TextFontsStyleHelper
    {
        // 建议放在 Command 类“辅助方法”区域
        private const string DefaultTextStyleName = "tJText";

        /// <summary>
        /// 新增：确保标题文本样式
        /// 确保并应用 _TitleStyle 以及辅助函数（放到 Command 类内部、靠近 TextStyleAndLayerInfo 定义处）EnsureTitleTextStyle
        /// </summary>
        /// <param name="tr"></param>
        public static void EnsureTitleTextStyle(DBTrans tr)
        {
            // 查找 _TitleStyle 如果不存在 _TitleStyle，则创建它
            if (!tr.TextStyleTable.Has("_TitleStyle"))
            {
                tr.TextStyleTable.Add("_TitleStyle", ttr =>
                {
                    ttr.FileName = "gbenor.shx";
                    ttr.BigFontFileName = "gbcbig.shx";
                    ttr.XScale = 0.8;
                });
            }
            else
            {
                /// 更新已有的 _TitleStyle
                tr.TextStyleTable.Change("_TitleStyle", ttr =>
                {
                    ttr.FileName = "gbenor.shx";
                    ttr.BigFontFileName = "gbcbig.shx";
                    ttr.XScale = 0.8;
                });
            }
        }

        /// <summary>
        /// 新增：确保标签文本样式
        /// </summary>
        /// <param name="tr"></param>
        public static void EnsureLabelTextStyle(DBTrans tr)
        {
            // 查找 _LabelStyle 如果不存在 _LabelStyle，则创建它
            if (!tr.TextStyleTable.Has("_LabelStyle"))
            {
                tr.TextStyleTable.Add("_LabelStyle", ttr =>
                {
                    ttr.FileName = "gbenor.shx";
                    ttr.BigFontFileName = "gbcbig.shx";
                    ttr.XScale = 0.8;
                });
            }
            else
            {
                /// 更新已有的 _LabelStyle
                tr.TextStyleTable.Change("_LabelStyle", ttr =>
                {
                    ttr.FileName = "gbenor.shx";
                    ttr.BigFontFileName = "gbcbig.shx";
                    ttr.XScale = 0.8;
                });
            }
        }

        /// <summary>
        /// 新增：确保字体样式
        /// </summary>
        /// <param name="tr"></param>
        public static void EnsureFontsStyle(DBTrans tr)
        {
            // 查找 _FontsStyle 如果不存在 _FontsStyle，则创建它
            if (!tr.TextStyleTable.Has("_FontsStyle"))
            {
                tr.TextStyleTable.Add("_FontsStyle", ttr =>
                {
                    ttr.FileName = "gbenor.shx";
                    ttr.BigFontFileName = "gbcbig.shx";
                    ttr.XScale = 0.8;
                });
            }
            else
            {
                /// 更新已有的 _FontsStyle
                tr.TextStyleTable.Change("_FontsStyle", ttr =>
                {
                    ttr.FileName = "gbenor.shx";
                    ttr.BigFontFileName = "gbcbig.shx";
                    ttr.XScale = 0.8;
                });
            }
        }

        /// <summary>
        /// 确保文本样式存在并正确设置，如果不存在则创建，如果已存在则更新其属性以符合预期的样式要求。
        /// 注意：关键改动 —— 将 TextSize 设为 0，避免“样式固定高度”覆盖实体高度。实体高度应由创建实体时显式设置。
        /// </summary>
        /// <param name="tr">数据库事务对象</param>
        /// <param name="textStyleName">文本样式名称</param>
        /// <returns>文本样式的 ObjectId</returns>
        public static ObjectId EnsureTextStyle(DBTrans tr, string textStyleName = DefaultTextStyleName)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return ObjectId.Null;
            using (doc.LockDocument()) // 锁定文档以安全修改
            {
                if (textStyleName == DefaultTextStyleName)
                {
                    if (!tr.TextStyleTable.Has(textStyleName))
                    {
                        tr.TextStyleTable.Add(textStyleName, ttr =>
                        {
                            ttr.FileName = "gbenor.shx";
                            ttr.BigFontFileName = "hzfs.shx";
                            ttr.XScale = 0.8;
                            // 关键：设为 0，避免样式固定高度反向覆盖实体高度
                            // 将 TextSize 设为 0 表示样式不固定高度，实体创建时通过 Height/TextHeight 显式控制。
                            ttr.TextSize = 0.0;
                        });
                    }
                    else
                    {
                        tr.TextStyleTable.Change(textStyleName, ttr =>
                        {
                            ttr.FileName = "gbenor.shx";
                            ttr.BigFontFileName = "hzfs.shx";
                            ttr.XScale = 0.8;
                            // 关键：设为 0，避免样式固定高度反向覆盖实体高度
                            ttr.TextSize = 0.0;
                        });
                    }
                }
                if (textStyleName == "HZFS")
                    if (!tr.TextStyleTable.Has(textStyleName))
                    {
                        tr.TextStyleTable.Add(textStyleName, ttr =>
                        {
                            ttr.FileName = "romanc.shx";
                            ttr.BigFontFileName = "hzfs.shx";
                            ttr.XScale = 0.8;
                            // 关键：设为 0，避免样式固定高度反向覆盖实体高度
                            ttr.TextSize = 0.0;
                        });
                    }
                    else
                    {
                        tr.TextStyleTable.Change(textStyleName, ttr =>
                        {
                            ttr.FileName = "romanc.shx";
                            ttr.BigFontFileName = "hzfs.shx";
                            ttr.XScale = 0.8;
                            ttr.TextSize = 0.0;
                        });
                    }
            }


            return tr.TextStyleTable[textStyleName];
        }

        /// <summary>
        /// 根据基本字高与图纸比例计算最终字高
        /// baseHeight: 比例为 1:1 时的字高（例如 3.5）
        /// scaleDenominator: 比例分母，例如 1 表示 1:1，100 表示 1:100。
        /// 如果传入 <= 0，则自动从 DrawingScaleService.GetScale(true) 读取视图比例并转换为分母。
        /// </summary>
        /// <param name="baseHeight"></param>
        /// <param name="scaleDenominator"></param>
        /// <returns></returns>
        public static double ComputeScaledHeight(double baseHeight, double scaleDenominator = 1.0)
        {
            if (baseHeight <= 0) baseHeight = 3.5;

            double denom = scaleDenominator;

            if (denom <= 0.0)
            {
                try
                {
                    // 从 DrawingScaleService 读取视图比例（可能为规范化值，如 0.01 表示 1:100）
                    double viewportScaleFactor = AutoCadHelper.GetScale(true);
                    // DetermineScaleDenominator 方法能把视图因子（0.01）或直接传入分母（如 100）归一为分母形式
                    denom = DetermineScaleDenominator(viewportScaleFactor, null, false);
                }
                catch
                {
                    denom = 1.0;
                }
            }

            if (double.IsNaN(denom) || double.IsInfinity(denom) || denom <= 0.0)
                denom = 1.0;

            // 采用简单乘法：基准高度 * 分母（例如 base 3.5, denom 100 -> 350）
            return baseHeight * denom;
        }

        /// <summary>
        /// 将 _TitleStyle 应用于 DBText（设置 TextStyleId 与 WidthFactor = 0.75）
        /// scaleDenominator: 传入比例分母（1 表示 1:1，100 表示 1:100），如果不传或传 <=0 则自动读取当前视图比例。
        /// </summary>
        public static void ApplyTitleToDBText(DBTrans tr, DBText dbText, double scaleDenominator = 0.0)
        {
            try
            {
                if (dbText == null) return;
                EnsureTitleTextStyle(tr);
                dbText.TextStyleId = tr.TextStyleTable["_TitleStyle"];
                dbText.WidthFactor = 0.75;
                // 根据比例设置字高：基准 3.5 -> 按分母计算（1:100 => 350）
                dbText.Height = ComputeScaledHeight(3.5, scaleDenominator);
            }
            catch { }
        }

        /// <summary>
        /// 将 _TitleStyle 应用于 MText（设置 TextStyleId 并尽量调整宽度）
        /// scaleDenominator: 传入比例分母（1 表示 1:1，100 表示 1:100），如果不传或传 <=0 则自动读取当前视图比例。
        /// </summary>
        public static void ApplyTitleToMText(DBTrans tr, MText mt, double scaleDenominator = 0.0)
        {
            try
            {
                if (mt == null) return;
                EnsureTitleTextStyle(tr);
                mt.TextStyleId = tr.TextStyleTable["_TitleStyle"];
                // 按图纸比例设置高度
                mt.Height = ComputeScaledHeight(3.5, scaleDenominator);
                // MText 无 WidthFactor：按请求将宽度尝试设置为高度的若干倍再乘以 0.75 以近似“宽度因子”
                if (mt.Width <= 0)
                    mt.Width = Math.Max(1.0, mt.Height * 10.0) * 0.75;
                else
                    mt.Width = mt.Width * 0.75;
            }
            catch { }
        }

        /// <summary>
        /// 尝试从多种输入获取并归一化为“比例分母”形式（scaleDenominator），确保返回 >= 1。
        /// 输入可以是：
        /// - viewportScaleFactor: 视口尺度因子（如 0.01 表示 1:100）或直接的分母（如 100）
        /// - scaleString: "1:100"、"100" 等
        /// 如果两者都提供，优先使用 viewportScaleFactor。
        /// </summary>
        /// <param name="viewportScaleFactor">视口尺度因子或分母，可为 null</param>
        /// <param name="scaleString">字符串形式尺度，可为 null</param>
        /// <param name="roundToCommon">是否将结果舍入到常用比例（用于 UI 显示），默认 false</param>
        /// <returns>比例分母（例如 1、100、200）</returns>
        public static double DetermineScaleDenominator(double? viewportScaleFactor = null, string scaleString = null, bool roundToCommon = false)
        {
            double denom = 1.0;

            if (viewportScaleFactor.HasValue)
            {
                var v = viewportScaleFactor.Value;
                if (v <= 0)
                {
                    denom = 1.0;
                }
                else if (v >= 1.0)
                {
                    // 可能直接传入分母（如 100）
                    denom = v;
                }
                else
                {
                    // 典型视口比例因子，例如 0.01 -> 100
                    denom = 1.0 / v;
                }
            }
            else if (!string.IsNullOrWhiteSpace(scaleString))
            {
                var parsed = ParseScaleString(scaleString);
                if (parsed > 0)
                    denom = parsed;
                else
                    denom = 1.0;
            }
            else
            {
                // 未知环境，回退到 1:1
                denom = 1.0;
            }

            // 防止极小或非法值
            if (double.IsNaN(denom) || double.IsInfinity(denom) || denom <= 0)
                denom = 1.0;

            if (roundToCommon)
                denom = RoundToCommonDenominator(denom);

            return Math.Max(1.0, denom);
        }

        /// <summary>
        /// 解析类似 "1:100"、"1/100"、"100"、"1:50" 的尺度字符串并返回分母（例如 100），解析失败返回 -1
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static double ParseScaleString(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return -1;
            s = s.Trim();
            // 支持 1:100 或 1/100
            if (s.Contains(":") || s.Contains("/"))
            {
                var delim = s.Contains(":") ? ':' : '/';
                var parts = s.Split(delim);
                if (parts.Length == 2)
                {
                    // 常见形式左为 1，右为 100 -> 返回右侧
                    if (double.TryParse(parts[1].Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out var right))
                        return right > 0 ? right : -1;
                }
                return -1;
            }
            // 直接数字，例如 "100"
            if (double.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out var val))
            {
                return val > 0 ? val : -1;
            }
            return -1;
        }

        /// <summary>
        /// 将传入的分母舍入到常用的比例集合（用于 UI 显示或自动选择）。
        /// 常用集合示例：1,2,5,10,20,25,50,100,200,250,500,1000
        /// </summary>
        /// <param name="denom"></param>
        /// <returns></returns>
        public static double RoundToCommonDenominator(double denom)
        {
            if (denom <= 1) return 1.0;
            double[] common = new double[] { 1, 2, 5, 10, 20, 25, 50, 100, 200, 250, 500, 1000, 2000 };
            double best = common[0];
            double bestDiff = Math.Abs(denom - best);
            foreach (var c in common)
            {
                var d = Math.Abs(denom - c);
                if (d < bestDiff)
                {
                    best = c;
                    bestDiff = d;
                }
            }
            return best;
        }

        #region 文字样式

        /// <summary>
        /// 文字样式与图层信息
        /// 关键改动：
        /// - TextStyle 使用 EnsureTextStyle 创建时 TextSize = 0（样式不固定高度）
        /// - 在此方法中按当前视图比例计算实体高度并显式赋值给 MText/MLeader，避免样式覆盖
        /// </summary>
        /// <param name="tr"></param>
        /// <param name="layerName"></param>
        /// <param name="textColor"></param>
        /// <param name="tJText"></param>
        /// <param name="mt"></param>
        /// <param name="mld"></param>
        public static void TextStyleAndLayerInfo(DBTrans tr, string layerName, short textColor, string tJText, ref MText mt, ref MLeader mld)
        {

            // 获取或创建目标图层，优先使用 btnBlockLayer 作为图层名称，如果没有则使用 layerName，确保图层存在并设置颜色。
            string targetLayer = LayerControlHelper.GetOrCreateTargetLayer(tr, layerName, textColor);
            // 获取或创建文字样式，并获取其 ObjectId 以供后续使用。
            ObjectId textStyleId = TextFontsStyleHelper.EnsureTextStyle(tr, tJText);

            // 使用当前视图的比例（分母）来计算实体的最终高度与箭头大小
            // AutoCadHelper.GetScale(true) 返回视图因子或分母信息，ComputeScaledHeight 会把它统一为分母形式
            double uiScaleDenominator = TextFontsStyleHelper.DetermineScaleDenominator(AutoCadHelper.GetScale(true));
            // 计算出在当前视图下实体应有的绝对高度（例如基准 3.5 在 1:100 时会得到 350）
            double defaultTextHeight = TextFontsStyleHelper.ComputeScaledHeight(3.5, uiScaleDenominator);
            double defaultArrowSize = TextFontsStyleHelper.ComputeScaledHeight(2.0, uiScaleDenominator);

            mt = new MText
            {
                Attachment = AttachmentPoint.MiddleCenter,
                Height = defaultTextHeight, // 显式设置实体高度，确保样式 TextSize=0 时生效
                ColorIndex = textColor,
                TextStyleId = textStyleId // 确保 MText 使用该文本样式（但样式不固定高度）
            };

            mld = new MLeader
            {
                Layer = targetLayer,
                ColorIndex = textColor,
                TextAttachmentType = TextAttachmentType.AttachmentBottomOfTopLine,
                ContentType = ContentType.MTextContent,
                LeaderLineColor = Color.FromColorIndex(ColorMethod.ByAci, textColor),
                LeaderLineTypeId = Env.Database.LinetypeTableId,
                MText = mt,
                TextStyleId = textStyleId,
                TextHeight = defaultTextHeight, // 显式设置 MLeader 的文字高度
                ArrowSize = defaultArrowSize
            };
        }
        /// <summary>
        /// 检查文字样式与图层信息
        /// </summary>
        /// <param name="tr"></param>
        /// <param name="layerName"></param>
        /// <param name="textColor"></param>
        /// <param name="tJText"></param>
        public static void TextStyleAndLayerInfo(DBTrans tr, string layerName, short textColor, string tJText)
        {
            _ = LayerControlHelper.GetOrCreateTargetLayer(tr, layerName, textColor);
            _ = TextFontsStyleHelper.EnsureTextStyle(tr, tJText);
        }

        #endregion
    }
}
