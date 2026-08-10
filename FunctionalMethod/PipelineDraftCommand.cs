using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using GB_NewCadPlus_IV.DisplayPages;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.Models;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using AutoCADApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 管道绘制草稿命令。
    /// 当前负责验证多点采集、端点属性识别和服务器基础数据调用，正式参数页面在后续步骤接入。
    /// </summary>
    public static class PipelineDraftCommand
    {
        /// <summary>
        /// 通过命令行测试管道草稿采集流程。
        /// </summary>
        [CommandMethod("PIPELINE_DRAFT")]
        public static async void DrawPipelineDraft()
        {
            Document document = AutoCADApplication.DocumentManager.MdiActiveDocument;
            if (document == null)
            {
                return;
            }

            Editor editor = document.Editor;
            LogManager.Instance.LogInfo(
                $"[PIPELINE_DRAFT][开始] Document={document.Name}, Database={document.Database.Filename}");
            try
            {
                string pipeRole = CollectPipeRole(editor);
                if (string.IsNullOrWhiteSpace(pipeRole))
                {
                    LogManager.Instance.LogInfo("[PIPELINE_DRAFT][取消] 未选择有效的管道角色。");
                    return;
                }
                LogManager.Instance.LogInfo($"[PIPELINE_DRAFT][角色完成] PipeRole={pipeRole}");

                List<Point3d> points = CollectPoints(editor);
                if (points.Count < 2)
                {
                    LogManager.Instance.LogWarning(
                        $"[PIPELINE_DRAFT][取消] 采点数量不足：PointCount={points.Count}");
                    editor.WriteMessage("\n至少需要指定起点和终点，管道草稿已取消。\n");
                    return;
                }
                LogManager.Instance.LogInfo(
                    $"[PIPELINE_DRAFT][采点完成] PointCount={points.Count}, Start={points[0]}, End={points[points.Count - 1]}");

                const double endpointTolerance = 2.0;
                PipelineEndpointSource startSource = PipelineEndpointPropertyHelper.FindNearestEntity(
                    document.Database,
                    points[0],
                    endpointTolerance);
                PipelineEndpointSource endSource = PipelineEndpointPropertyHelper.FindNearestEntity(
                    document.Database,
                    points[points.Count - 1],
                    endpointTolerance);
                LogManager.Instance.LogInfo(
                    $"[PIPELINE_DRAFT][端点读取完成] StartHandle={startSource?.Handle ?? "无"}, EndHandle={endSource?.Handle ?? "无"}, Tolerance={endpointTolerance}");

                StandardApiService apiService = new StandardApiService();
                PipelineDefaultsResponseClient defaultsResponse =
                    await apiService.GetPipelineDefaultsAsync();
                LogManager.Instance.LogInfo(
                    $"[PIPELINE_DRAFT][默认值完成] Success={defaultsResponse.Success}, AttributeCount={defaultsResponse.Attributes.Count}, Message={defaultsResponse.Message}");
                PipelineEndpointMergeResult mergeResult = PipelineEndpointPropertyHelper.MergeEndpointProperties(
                    defaultsResponse.Attributes,
                    startSource,
                    endSource);
                LogManager.Instance.LogInfo(
                    $"[PIPELINE_DRAFT][端点合并完成] AttributeCount={mergeResult.Attributes.Count}, ConflictCount={mergeResult.StartEndConflicts.Count}");

                editor.WriteMessage(
                    $"\n管道草稿点数={points.Count}，起点图元={startSource?.Handle ?? "无"}，终点图元={endSource?.Handle ?? "无"}，属性数={mergeResult.Attributes.Count}。\n");
                foreach (string conflict in mergeResult.StartEndConflicts)
                {
                    editor.WriteMessage($"\n属性冲突：{conflict}");
                }

                string standardNo = GetAttribute(mergeResult.Attributes, "DRAWINGNO.STANDARDNO");
                string dn = GetAttribute(mergeResult.Attributes, "DN");
                string pn = GetAttribute(mergeResult.Attributes, "PN");
                if (!string.IsNullOrWhiteSpace(standardNo) &&
                    !string.IsNullOrWhiteSpace(dn) &&
                    !string.IsNullOrWhiteSpace(pn))
                {
                    PipelineDesignStandardMatchResponseClient standardResponse =
                        await apiService.MatchPipelineDesignStandardAsync(
                            new PipelineDesignStandardMatchRequestClient
                            {
                                DrawingStandardNo = standardNo,
                                DN = dn,
                                PN = pn,
                                Schedule = GetAttribute(mergeResult.Attributes, "SCHEDULE"),
                                PipeMaterial = GetAttribute(mergeResult.Attributes, "PIPE_MATL"),
                                Medium = GetAttribute(mergeResult.Attributes, "MEDIUM")
                            });
                    LogManager.Instance.LogInfo(
                        $"[PIPELINE_DRAFT][GB规范完成] Success={standardResponse.Success}, MatchCount={standardResponse.MatchCount}, Message={standardResponse.Message}");

                    editor.WriteMessage(
                        $"\nGB 设计规范匹配：成功={standardResponse.Success}，匹配数={standardResponse.MatchCount}，消息={standardResponse.Message}\n");

                    if (standardResponse.Success)
                    {
                        MergeStandardAttributes(mergeResult.Attributes, standardResponse.Attributes);
                    }
                }
                else
                {
                    editor.WriteMessage("\n缺少 DRAWINGNO.STANDARDNO、DN 或 PN，暂不请求 GB 设计规范。\n");
                }

                PipelineFieldCatalogResponseClient catalogResponse =
                    await apiService.GetPipelineFieldCatalogAsync();
                LogManager.Instance.LogInfo(
                    $"[PIPELINE_DRAFT][字段目录完成] Success={catalogResponse.Success}, FieldCount={catalogResponse.Fields.Count}, Message={catalogResponse.Message}");
                if (!catalogResponse.Success)
                {
                    editor.WriteMessage($"\n管道字段目录获取失败：{catalogResponse.Message}\n");
                    return;
                }

                mergeResult.Attributes["PIPE_ROLE"] = pipeRole;
                PipelineParameterWindow parameterWindow = new PipelineParameterWindow();
                parameterWindow.Initialize(
                    pipeRole,
                    catalogResponse.Fields,
                    mergeResult.Attributes,
                    mergeResult.StartEndConflicts);

                bool? dialogResult = parameterWindow.ShowDialog();
                if (dialogResult != true)
                {
                    LogManager.Instance.LogInfo("[PIPELINE_DRAFT][取消] 用户取消管道参数确认。");
                    editor.WriteMessage("\n用户取消管道参数确认，未生成 CAD 管道。\n");
                    return;
                }
                LogManager.Instance.LogInfo(
                    $"[PIPELINE_DRAFT][参数确认完成] AttributeCount={parameterWindow.ConfirmedAttributes.Count}, Title={GetAttribute(parameterWindow.ConfirmedAttributes, "PIPELINETITLE")}");

                LogManager.Instance.LogInfo("[PIPELINE_DRAFT][准备落图] 即将获取当前文档写锁。");
                PipelineCadPlacementResult placementResult;
                using (DocumentLock documentLock = document.LockDocument())
                {
                    LogManager.Instance.LogInfo("[PIPELINE_DRAFT][文档锁成功] 已获取当前文档写锁。");
                    placementResult = PipelineCadObjectService.PlacePipeline(
                        document.Database,
                        points,
                        parameterWindow.ConfirmedAttributes,
                        pipeRole);

                    // CAD 对象事务提交后立即重生成当前视图，确保标题和流向符号及时显示。
                    editor.Regen();
                    LogManager.Instance.LogInfo("[PIPELINE_DRAFT][显示刷新完成] 已执行当前视图 Regen。");
                }
                LogManager.Instance.LogInfo(
                    $"[PIPELINE_DRAFT][落图完成] PipeId={placementResult.PipeId}, PipelineObjectId={placementResult.PipelineObjectId}, TitleObjectId={placementResult.TitleObjectId}, FlowObjectId={placementResult.FlowDirectionObjectId}");
                editor.WriteMessage(
                    $"\n管道已完成绘制：PipeId={placementResult.PipeId}，点数={points.Count}，标题={parameterWindow.ConfirmedAttributes["PIPELINETITLE"]}\n");
            }
            catch (Exception exception)
            {
                LogManager.Instance.LogError(
                    $"[PIPELINE_DRAFT][异常] Type={exception.GetType().FullName}, Message={exception.Message}, StackTrace={exception.StackTrace}");
                editor.WriteMessage($"\n管道草稿命令失败：{exception.Message}\n");
            }
        }

        /// <summary>
        /// 选择统一管道参数流程使用的进口或出口角色。
        /// </summary>
        private static string CollectPipeRole(Editor editor)
        {
            PromptKeywordOptions options = new PromptKeywordOptions(
                "\n选择管道角色 [进口/出口] <进口>：")
            {
                AllowNone = true
            };
            options.Keywords.Add("进口", "进口", "进口");
            options.Keywords.Add("出口", "出口", "出口");

            PromptResult result = editor.GetKeywords(options);
            if (result.Status == PromptStatus.None)
            {
                return PipelineRoles.Import;
            }

            if (result.Status != PromptStatus.OK)
            {
                return string.Empty;
            }

            return string.Equals(result.StringResult, "出口", StringComparison.OrdinalIgnoreCase)
                ? PipelineRoles.Export
                : PipelineRoles.Import;
        }

        /// <summary>
        /// 将服务器规范匹配返回的真实属性合并到当前管道属性。
        /// </summary>
        private static void MergeStandardAttributes(
            IDictionary<string, string> target,
            IDictionary<string, string> standardAttributes)
        {
            if (target == null || standardAttributes == null)
            {
                return;
            }

            foreach (KeyValuePair<string, string> pair in standardAttributes)
            {
                if (string.IsNullOrWhiteSpace(pair.Key))
                {
                    continue;
                }

                target[pair.Key.Trim()] = pair.Value ?? string.Empty;
                LogManager.Instance.LogInfo(
                    $"[管道规范属性同步] Tag={pair.Key}, NewValue={pair.Value ?? string.Empty}");
            }
        }

        /// <summary>
        /// 连续采集管道点，Z 撤销上一点，回车或右键完成。
        /// </summary>
        private static List<Point3d> CollectPoints(Editor editor)
        {
            List<Point3d> points = new List<Point3d>();
            while (true)
            {
                PromptPointOptions options = new PromptPointOptions(
                    points.Count == 0
                        ? "\n指定管道起点（Z 撤销，回车完成）："
                        : "\n指定管道下一点（Z 撤销，回车完成）：")
                {
                    AllowNone = true,
                    UseDashedLine = points.Count > 0
                };

                options.Keywords.Add("Z", "Z", "撤销");
                if (points.Count > 0)
                {
                    options.BasePoint = points[points.Count - 1];
                    options.UseBasePoint = true;
                }

                PromptPointResult pointResult = editor.GetPoint(options);
                if (pointResult.Status == PromptStatus.Keyword &&
                    string.Equals(pointResult.StringResult, "Z", StringComparison.OrdinalIgnoreCase))
                {
                    if (points.Count > 0)
                    {
                        points.RemoveAt(points.Count - 1);
                    }
                    continue;
                }

                if (pointResult.Status == PromptStatus.None)
                {
                    break;
                }

                if (pointResult.Status != PromptStatus.OK)
                {
                    points.Clear();
                    break;
                }

                points.Add(pointResult.Value);
            }

            return points;
        }

        /// <summary>
        /// 按不区分大小写的 Tag 读取属性。
        /// </summary>
        private static string GetAttribute(
            IDictionary<string, string> attributes,
            string tag)
        {
            return attributes != null && attributes.TryGetValue(tag, out string value)
                ? value ?? string.Empty
                : string.Empty;
        }
    }
}
