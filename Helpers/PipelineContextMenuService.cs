using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System;
using System.Collections.Generic;
using GB_NewCadPlus_IV.FunctionalMethod;
using AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadDocument = Autodesk.AutoCAD.ApplicationServices.Document;
using AcadMenuItem = Autodesk.AutoCAD.Windows.MenuItem;

namespace GB_NewCadPlus_IV.Helpers
{
    /// <summary>
    /// 管道 Polyline 的 AutoCAD 右键菜单。
    /// </summary>
    public static class PipelineContextMenuService
    {
        private static ContextMenuExtension _contextMenu;
        private static AcadMenuItem _editItem;
        private static bool _registered;

        public static void Register()
        {
            if (_registered)
            {
                return;
            }

            _contextMenu = new ContextMenuExtension();
            _editItem = new AcadMenuItem("管道修改");
            _editItem.Click += EditItem_Click;
            _contextMenu.MenuItems.Add(_editItem);

            AcadApplication.AddObjectContextMenuExtension(
                RXClass.GetClass(typeof(Polyline)),
                _contextMenu);
            _registered = true;
        }

        public static void Unregister()
        {
            if (!_registered)
            {
                return;
            }

            try
            {
                _editItem.Click -= EditItem_Click;
                AcadApplication.RemoveObjectContextMenuExtension(
                    RXClass.GetClass(typeof(Polyline)),
                    _contextMenu);
            }
            finally
            {
                _editItem = null;
                _contextMenu = null;
                _registered = false;
            }
        }

        private static void EditItem_Click(object sender, EventArgs e)
        {
            AcadDocument document = AcadApplication.DocumentManager.MdiActiveDocument;
            if (document == null)
            {
                return;
            }

            ObjectId pipelineObjectId = GetSelectedPipelineObjectId(document);
            if (pipelineObjectId == ObjectId.Null)
            {
                document.Editor.WriteMessage("\n当前 Polyline 不是管道对象，未执行管道修改。\n");
                return;
            }

            document.SendStringToExecute(
                "PIPELINE_EDIT ",
                true,
                false,
                false);
        }

        private static ObjectId GetSelectedPipelineObjectId(AcadDocument document)
        {
            PromptSelectionResult selection = document.Editor.SelectImplied();
            if (selection.Status != PromptStatus.OK || selection.Value.Count == 0)
            {
                return ObjectId.Null;
            }

            using (Transaction transaction = document.Database.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selectedObject in selection.Value)
                {
                    if (selectedObject == null)
                    {
                        continue;
                    }

                    Polyline pipeline = transaction.GetObject(
                        selectedObject.ObjectId,
                        OpenMode.ForRead) as Polyline;
                    if (pipeline == null)
                    {
                        continue;
                    }

                    Dictionary<string, string> properties =
                        PipelineEndpointPropertyHelper.ReadEntityProperties(transaction, pipeline);
                    if (properties.TryGetValue("PIPEID", out string pipeId) &&
                        !string.IsNullOrWhiteSpace(pipeId))
                    {
                        transaction.Commit();
                        return selectedObject.ObjectId;
                    }
                }

                transaction.Commit();
            }

            return ObjectId.Null;
        }
    }
}
