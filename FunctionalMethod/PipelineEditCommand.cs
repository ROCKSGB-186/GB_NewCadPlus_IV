using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using GB_NewCadPlus_IV.DisplayPages;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using AutoCADApplication = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_IV.FunctionalMethod
{
    /// <summary>
    /// 管道选择和参数编辑命令。
    /// </summary>
    public static class PipelineEditCommand
    {
        /// <summary>
        /// 选择带 PIPEID 的 Polyline 管道并打开通用参数页面。
        /// </summary>
        [CommandMethod("PIPELINE_EDIT")]
        public static async void EditPipeline()
        {
            Document document = AutoCADApplication.DocumentManager.MdiActiveDocument;
            if (document == null)
            {
                return;
            }

            Editor editor = document.Editor;
            PromptEntityOptions options = new PromptEntityOptions("\n选择要编辑的管道：");
            options.SetRejectMessage("\n请选择管道 Polyline。\n");
            options.AddAllowedClass(typeof(Polyline), true);

            PromptEntityResult selection = editor.GetEntity(options);
            if (selection.Status != PromptStatus.OK)
            {
                return;
            }

            Dictionary<string, string> attributes;
            using (Transaction transaction = document.Database.TransactionManager.StartTransaction())
            {
                Entity entity = transaction.GetObject(selection.ObjectId, OpenMode.ForRead) as Entity;
                if (!(entity is Polyline))
                {
                    editor.WriteMessage("\n选择的对象不是管道 Polyline。\n");
                    return;
                }

                attributes = PipelineEndpointPropertyHelper.ReadEntityProperties(transaction, entity);
                if (entity is Polyline selectedPipeline)
                {
                    double currentLength = PipelineCadEditService.GetPolylineLength(selectedPipeline);
                    attributes["PIPE_LENGTH"] = currentLength.ToString(
                        "0.###",
                        CultureInfo.InvariantCulture);
                    LogManager.Instance.LogInfo(
                        $"[PIPELINE_EDIT][当前长度读取] ObjectId={selection.ObjectId}, CurrentLength={attributes["PIPE_LENGTH"]}, VertexCount={selectedPipeline.NumberOfVertices}");
                }
                transaction.Commit();
            }

            LogManager.Instance.LogInfo(
                $"[PIPELINE_EDIT][属性读取完成] ObjectId={selection.ObjectId}, AttributeCount={attributes.Count}, HasPipeId={attributes.ContainsKey("PIPEID")}");

            if (!attributes.TryGetValue("PIPEID", out string pipeId) || string.IsNullOrWhiteSpace(pipeId))
            {
                editor.WriteMessage("\n该 Polyline 没有 PIPEID，不是新管道对象。\n");
                return;
            }

            try
            {
                StandardApiService apiService = new StandardApiService();
                PipelineFieldCatalogResponseClient catalog =
                    await apiService.GetPipelineFieldCatalogAsync();
                if (!catalog.Success)
                {
                    editor.WriteMessage($"\n管道字段目录获取失败：{catalog.Message}\n");
                    LogManager.Instance.LogWarning(
                        $"[PIPELINE_EDIT][字段目录失败] Message={catalog.Message}");
                    return;
                }

                LogManager.Instance.LogInfo(
                    $"[PIPELINE_EDIT][字段目录完成] FieldCount={catalog.Fields.Count}, AttributeCount={attributes.Count}");

                PipelineParameterWindow window = new PipelineParameterWindow();
                window.Initialize(
                    attributes.TryGetValue("PIPE_ROLE", out string role) ? role : PipelineRoles.Import,
                    catalog.Fields,
                    attributes,
                    new List<string>());

                bool? dialogResult = window.ShowDialog();
                if (dialogResult != true)
                {
                    return;
                }

                using (DocumentLock documentLock = document.LockDocument())
                {
                    LogManager.Instance.LogInfo("[PIPELINE_EDIT][文档锁成功] 已获取当前文档写锁。");
                    PipelineCadEditService.UpdatePipeline(
                        document.Database,
                        selection.ObjectId,
                        window.ConfirmedAttributes);
                    editor.Regen();
                }
                LogManager.Instance.LogInfo(
                    $"[PIPELINE_EDIT][回写完成] PipeId={pipeId}, AttributeCount={window.ConfirmedAttributes.Count}");
                editor.WriteMessage($"\n管道参数已更新，PipeId={pipeId}。\n");
            }
            catch (Exception exception)
            {
                LogManager.Instance.LogInfo($"管道编辑命令异常：PipeId={pipeId}，错误={exception.Message}");
                editor.WriteMessage($"\n管道编辑失败：{exception.Message}\n");
            }
        }
    }
}
