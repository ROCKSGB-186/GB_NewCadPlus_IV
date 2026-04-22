using Autodesk.AutoCAD.PlottingServices;
using GB_NewCadPlus_IV.FunctionalMethod;
using GB_NewCadPlus_IV.UniFiedStandards;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static GB_NewCadPlus_IV.FunctionalMethod.Command;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace GB_NewCadPlus_IV.Helpers
{
    public static class LayerControlHelper
    {
        /// <summary>
        /// 统一图层入口：解析图层名 + 解析颜色 + 确保图层存在
        /// </summary>
        public static string GetOrCreateTargetLayer(DBTrans tr, string? preferredLayerName = null, short? preferredColor = null)
        {
            string layerName = LayerControlHelper.ResolveTargetLayerName(preferredLayerName);
            short color = preferredColor ?? LayerControlHelper.ResolveLayerColor();
            return LayerDictionaryHelper.EnsureTargetLayer(tr, layerName, color);
        }

        #region 模型空间检查图层

        /// <summary>
        /// 关闭图层
        /// </summary>
        [CommandMethod(nameof(CloseLayer))]
        public static void CloseLayer()
        {
            try
            {
                using var tr = new DBTrans();
                if (VariableDictionary.btnState)
                {
                    SaveAllTjLayerStates();//保存所有条件图层状态
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n已保存 {layerOriStates.Count} 个图层的初始状态");
                    // 打开图层表  
                    //var layerTable = tr.LayerTable;

                    // 遍历图层并进行一些修改（这里以关闭部分图层为例）  
                    foreach (var layerId in layerOriStates)
                    {
                        //块表记录
                        LayerTableRecord layerRecord = tr.GetObject(layerId.ObjectId, OpenMode.ForWrite) as LayerTableRecord;

                        if (VariableDictionary.allTjtLayer.Contains(layerRecord.Name))
                        {
                            layerRecord.IsOff = true;//关闭图层
                        }
                    }
                }
                else
                {
                    // 还原图层状态  
                    RestoreLayerStates(layerOriStates);
                }
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("关闭图层失败！");
                LogManager.Instance.LogInfo($"\n错误信息：{ex.Message}");
            }
        }

        /// <summary>  
        /// 保存当前文档中所有图层的状态  
        /// </summary>  
        /// <returns>图层状态列表</returns>  
        public static void SaveLayerStates()
        {
            // 获取当前活动文档的数据库  
            Database db = Application.DocumentManager.MdiActiveDocument.Database;

            // 使用事务读取图层信息  
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 清空储存图层列表变量
                    layerOriStates.Clear();
                    // 打开图层表  
                    LayerTable layerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    // 遍历所有图层  
                    foreach (ObjectId layerId in layerTable)
                    {
                        // 获取图层表记录  
                        LayerTableRecord layerRecord = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
                        // 创建并保存图层状态  
                        LayerState state = new LayerState
                        {
                            ObjectId = layerRecord.ObjectId,
                            Name = layerRecord.Name,//图层名称
                            IsOff = layerRecord.IsOff,//图层是否关闭
                            #region 图层其它属性
                            IsFrozen = layerRecord.IsFrozen, // 图层是否冻结
                            //IsLocked = layerRecord.IsLocked,// 图层是否锁定
                            //IsPlottable = layerRecord.IsPlottable,// 图层是否可打印  
                            ColorIndex = layerRecord.Color.ColorIndex,// 图层颜色索引  
                            //LineWeight = layerRecord.LineWeight, // 图层线宽  
                            //PlotStyleName = layerRecord.PlotStyleName,// 图层打印样式名称
                            //PlotStyleNameId = layerRecord.PlotStyleNameId,// 图层打印样式名称ID
                            //DrawStream = layerRecord.DrawStream,// 图层绘制数据
                            //IsHidden = layerRecord.IsHidden,//   图层是否隐藏
                            //IsReconciled = layerRecord.IsReconciled,//   图层是否同步
                            //LinetypeObjectId = layerRecord.LinetypeObjectId,// 图层线型对象ID
                            //MaterialId = layerRecord.MaterialId,// 图层材质ID
                            //MergeStyle = layerRecord.MergeStyle,// 图层合并样式
                            //OwnerId = layerRecord.OwnerId,// 图层所有者ID
                            //Transparency = layerRecord.Transparency,// 图层透明度
                            //ViewportVisibilityDefault = layerRecord.ViewportVisibilityDefault,// 图层视图可见性默认值
                            //XData = layerRecord.XData,// 图层X数据
                            //Annotative = layerRecord.Annotative,// 图层注释
                            //Description = layerRecord.Description,// 图层描述
                            //HasSaveVersionOverride = layerRecord.HasSaveVersionOverride,// 图层保存版本覆盖
                            #endregion
                        };
                        layerOriStates.Add(state);
                    }
                    tr.Commit();
                }
                catch (Exception ex)
                {
                    // 处理可能的异常  
                    Application.ShowAlertDialog($"保存图层状态时发生错误：{ex.Message}");
                    tr.Abort();
                }
            }
        }

        /// <summary>  
        /// 保存当前文档中所有图层的状态  
        /// </summary>  
        /// <returns>图层状态列表</returns>  
        public static void SaveAllTjLayerStates()
        {
            // 使用事务读取图层信息  
            using (DBTrans tr = new())
            {
                try
                {
                    layerOriStates.Clear();
                    // 打开图层表  
                    var layerTable = tr.LayerTable;
                    // 遍历所有图层  
                    foreach (ObjectId layerId in layerTable)
                    {
                        // 获取图层表记录  
                        LayerTableRecord layerRecord = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;

                        if (VariableDictionary.allTjtLayer.Contains(layerRecord.Name))
                        {
                            // 创建并保存图层状态  
                            LayerState state = new LayerState
                            {
                                ObjectId = layerRecord.ObjectId,
                                Name = layerRecord.Name,
                                IsOff = layerRecord.IsOff,
                                #region 图层其它属性
                                IsFrozen = layerRecord.IsFrozen, // 图层是否冻结
                                //IsLocked = layerRecord.IsLocked,// 图层是否锁定
                                //IsPlottable = layerRecord.IsPlottable,// 图层是否可打印  
                                //ColorIndex = layerRecord.Color.ColorIndex,// 图层颜色索引  
                                //LineWeight = layerRecord.LineWeight, // 图层线宽  
                                //PlotStyleName = layerRecord.PlotStyleName,// 图层打印样式名称
                                //PlotStyleNameId = layerRecord.PlotStyleNameId,// 图层打印样式名称ID
                                //DrawStream = layerRecord.DrawStream,// 图层绘制数据
                                //IsHidden = layerRecord.IsHidden,//   图层是否隐藏
                                //IsReconciled = layerRecord.IsReconciled,//   图层是否同步
                                //LinetypeObjectId = layerRecord.LinetypeObjectId,// 图层线型对象ID
                                //MaterialId = layerRecord.MaterialId,// 图层材质ID
                                //MergeStyle = layerRecord.MergeStyle,// 图层合并样式
                                //OwnerId = layerRecord.OwnerId,// 图层所有者ID
                                //Transparency = layerRecord.Transparency,// 图层透明度
                                //ViewportVisibilityDefault = layerRecord.ViewportVisibilityDefault,// 图层视图可见性默认值
                                //XData = layerRecord.XData,// 图层X数据
                                //Annotative = layerRecord.Annotative,// 图层注释
                                //Description = layerRecord.Description,// 图层描述
                                //HasSaveVersionOverride = layerRecord.HasSaveVersionOverride,// 图层保存版本覆盖
                                #endregion
                            };
                            layerOriStates.Add(state);
                        }
                    }
                    tr.Commit();
                }
                catch (Exception ex)
                {
                    // 处理可能的异常  
                    Application.ShowAlertDialog($"保存图层状态时发生错误：{ex.Message}");
                    tr.Abort();
                }
            }
        }

        /// <summary>  
        /// 还原之前保存的图层状态  
        /// </summary>  
        /// <param name="savedLayerStates">之前保存的图层状态列表</param>  
        public static void RestoreLayerStates(List<LayerState> savedLayerStates)
        {
            // 获取当前活动文档的数据库  
            Database db = Application.DocumentManager.MdiActiveDocument.Database;

            // 使用事务还原图层状态  
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 打开图层表  
                    LayerTable layerTable = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    // 遍历保存的图层状态  
                    foreach (LayerState state in savedLayerStates)
                    {
                        // 检查图层是否存在  
                        if (layerTable.Has(state.Name))
                        {
                            // 获取图层表记录并打开写入模式  
                            ObjectId layerId = layerTable[state.Name];
                            LayerTableRecord layerRecord = trans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                            // 还原图层状态  
                            layerRecord.IsOff = state.IsOff;
                            #region 图层的其它属性
                            layerRecord.IsFrozen = state.IsFrozen; // 图层是否冻结
                            layerRecord.IsLocked = state.IsLocked;// 图层是否锁定
                            //layerRecord.IsPlottable = state.IsPlottable;//图层是否可打印  
                            //layerRecord.Color = Color.FromColorIndex(ColorMethod.ByAci, state.ColorIndex);// 还原颜色和线宽  
                            //layerRecord.LineWeight = state.LineWeight;// 图层线宽 
                            #endregion
                            /*
                              ObjectId = layerRecord.ObjectId,
                                Name = layerRecord.Name,
                                IsOff = layerRecord.IsOff,
                                #region 图层其它属性
                                IsFrozen = layerRecord.IsFrozen, // 图层是否冻结
                                IsLocked = layerRecord.IsLocked,// 图层是否锁定
                                IsPlottable = layerRecord.IsPlottable,// 图层是否可打印  
                                ColorIndex = layerRecord.Color.ColorIndex,// 图层颜色索引  
                                LineWeight = layerRecord.LineWeight, // 图层线宽  
                                PlotStyleName = layerRecord.PlotStyleName,// 图层打印样式名称
                                PlotStyleNameId = layerRecord.PlotStyleNameId,// 图层打印样式名称ID
                                DrawStream = layerRecord.DrawStream,// 图层绘制数据
                                IsHidden = layerRecord.IsHidden,//   图层是否隐藏
                                IsReconciled = layerRecord.IsReconciled,//   图层是否同步
                                LinetypeObjectId = layerRecord.LinetypeObjectId,// 图层线型对象ID
                                MaterialId = layerRecord.MaterialId,// 图层材质ID
                                MergeStyle = layerRecord.MergeStyle,// 图层合并样式
                                OwnerId = layerRecord.OwnerId,// 图层所有者ID
                                Transparency = layerRecord.Transparency,// 图层透明度
                                ViewportVisibilityDefault = layerRecord.ViewportVisibilityDefault,// 图层视图可见性默认值
                                XData = layerRecord.XData,// 图层X数据
                                Annotative = layerRecord.Annotative,// 图层注释
                                Description = layerRecord.Description,// 图层描述
                                HasSaveVersionOverride = layerRecord.HasSaveVersionOverride,// 图层保存版本覆盖
                                #endregion
                             */
                        }
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    // 处理可能的异常  
                    Application.ShowAlertDialog($"还原图层状态时发生错误：{ex.Message}");
                    trans.Abort();
                }
            }
        }

        /// <summary>
        /// 图层删除和恢复命令
        /// </summary>
        [CommandMethod("ToggleLayerDeletion")]
        public static void ToggleLayerDeletion()
        {
            // 获取当前文档和数据库
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 使用事务处理
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 如果尚未删除图层，执行删除操作
                    if (!_isLayersDeleted)
                    {
                        // 提示用户选择要保留的图层
                        PromptEntityResult selectedLayerResult = ed.GetEntity("请选择要保留的图层");
                        if (selectedLayerResult.Status != PromptStatus.OK)
                        {
                            ed.WriteMessage("\n未选择图层，操作取消。");
                            return;
                        }
                        // 打开图层表
                        LayerTable layerTable = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                        // 获取选中的图层名称
                        Entity selectEntity = trans.GetObject(selectedLayerResult.ObjectId, OpenMode.ForRead) as Entity;
                        //LayerTable selectLayerTable = trans.GetObject(selectEntity.ObjectId, OpenMode.ForRead) as LayerTable;
                        //LayerTableRecord selectedLayer = trans.GetObject(selectLayerTable.ObjectId, OpenMode.ForRead) as LayerTableRecord;

                        //LayerTableRecord selectedLayer1 = trans.GetObject(selectedLayerResult.ObjectId, OpenMode.ForRead) as LayerTableRecord;
                        if (selectEntity == null)
                        {
                            ed.WriteMessage("\n未找到图层，操作取消。");
                            return;
                        }
                        //LayerTableRecord selectedLayer = trans.GetObject(selectLayerTable.ObjectId, OpenMode.ForRead) as LayerTableRecord;
                        string preserveLayerName = selectEntity.Layer;

                        // 初始化原始图层信息字典
                        _originalLayerInfos = new Dictionary<string, LayerState>();

                        // 准备删除图层的事务
                        trans.GetObject(db.LayerTableId, OpenMode.ForWrite);

                        // 遍历所有图层
                        foreach (ObjectId layerId in layerTable)
                        {
                            LayerTableRecord layer = trans.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;

                            // 跳过0图层和选中的图层
                            if (layer.Name == "0" || layer.Name == "Defpoints" || layer.Name == preserveLayerName)
                                continue;

                            // 记录原始图层信息
                            _originalLayerInfos[layer.Name] = new LayerState
                            {
                                Name = layer.Name,
                                ColorIndex = layer.Color.ColorIndex,
                                LinetypeObjectId = layer.LinetypeObjectId,
                                LineWeight = layer.LineWeight,
                                IsFrozen = layer.IsFrozen,
                                IsLocked = layer.IsLocked,
                                IsOff = layer.IsOff
                            };

                            // 删除图层（切换写入模式）
                            layer.UpgradeOpen();
                            layer.Erase();
                        }

                        // 提交事务
                        trans.Commit();
                        Env.Editor.Redraw();

                        // 设置删除状态
                        _isLayersDeleted = true;
                        ed.WriteMessage($"\n已删除除 {preserveLayerName} 和 0 图层外的所有图层。");
                    }
                    else
                    {
                        // 恢复图层
                        LayerTable layerTable = trans.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable;

                        // 重新创建之前删除的图层
                        foreach (var layerInfo in _originalLayerInfos)
                        {
                            // 创建新的图层表记录
                            LayerTableRecord newLayer = new LayerTableRecord
                            {
                                Name = layerInfo.Value.Name,
                                Color = Color.FromColorIndex(ColorMethod.ByAci, layerInfo.Value.ColorIndex),// 还原颜色和线宽
                                LineWeight = layerInfo.Value.LineWeight
                            };

                            // 设置线型
                            newLayer.LinetypeObjectId = layerInfo.Value.LinetypeObjectId;

                            // 设置图层状态
                            newLayer.IsLocked = layerInfo.Value.IsLocked;
                            newLayer.IsOff = layerInfo.Value.IsOff;

                            // 将图层添加到图层表
                            layerTable.Add(newLayer);
                            trans.AddNewlyCreatedDBObject(newLayer, true);
                        }

                        // 提交事务
                        trans.Commit();

                        // 重置状态
                        _isLayersDeleted = false;
                        ed.WriteMessage("\n已恢复所有删除的图层。");
                    }
                }
                catch (Exception ex)
                {
                    // 错误处理
                    ed.WriteMessage($"\n发生错误：{ex.Message}");
                    trans.Abort();
                }
            }
        }

        /// <summary>
        /// 关闭图层
        /// </summary>
        [CommandMethod(nameof(CloseAllLayer))]
        public static void CloseAllLayer()
        {
            // 使用事务修改图层状态  
            using (var tr = new DBTrans())
            {
                try
                {
                    //// 第三步：等待用户确认  
                    //PromptResult pr = Application.DocumentManager.MdiActiveDocument.Editor.GetString("\n已修改图层状态，是否还原? [是/否] <是>: ");

                    //// 判断是否还原  
                    //if (pr.Status == PromptStatus.OK &&
                    //    (string.IsNullOrEmpty(pr.StringResult) ||
                    //     pr.StringResult.Trim().ToLower() == "是" ||
                    //     pr.StringResult.Trim().ToLower() == "y"))
                    //{
                    //    // 还原图层状态  
                    //    RestoreLayerStates(originalLayerStates);
                    //    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n已还原所有图层的初始状态");
                    //}
                    // 第一步：保存当前所有图层的初始状态  

                    if (VariableDictionary.btnState)
                    {
                        SaveLayerStates();
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n已保存 {layerOriStates.Count} 个图层的初始状态");
                        // 打开图层表  
                        LayerTable layerTable = tr.GetObject(tr.Database.LayerTableId, OpenMode.ForRead) as LayerTable;

                        // 遍历图层并进行一些修改（这里以关闭部分图层为例）  
                        foreach (ObjectId layerId in layerTable)
                        {
                            LayerTableRecord layerRecord = tr.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;

                            if (!VariableDictionary.selectTjtLayer.Contains(layerRecord.Name))
                            {
                                //layerRecord.IsOff = true;//关闭图层
                                layerRecord.IsFrozen = true;//冻结图层
                            }
                        }
                    }
                    else
                    {
                        // 还原图层状态  
                        RestoreLayerStates(layerOriStates);
                    }
                    tr.Commit();
                    Env.Editor.Redraw();
                }
                catch (Exception ex)
                {
                    // 处理可能的异常  
                    Application.ShowAlertDialog($"修改图层状态时发生错误：{ex.Message}");
                    tr.Abort();
                }
            }

        }

        /// <summary>
        /// 打开图层
        /// </summary>
        [CommandMethod(nameof(OpenLayer2))]
        public static void OpenLayer2()
        {
            try
            {
                if (!readLayerONOFFState)
                    layerOnOff();
                using var tr = new DBTrans();
                if (!VariableDictionary.btnState)
                {
                    foreach (var layerName in layerOnOffDic)
                    {
                        var layerLtr = tr.LayerTable.GetRecord(layerName.Key, OpenMode.ForWrite);
                        if (layerLtr != null)
                        {
                            foreach (var layer in VariableDictionary.tjtBtn)
                            {
                                if (layer == layerName.Key)
                                { layerLtr.IsOff = false; layerLtr.IsFrozen = false; }
                                else
                                {
                                    //layerLtr.IsOff = true;
                                    layerLtr.IsFrozen = true;
                                    break;
                                }
                            }
                        }
                    }
                    VariableDictionary.btnState = true;
                }

                foreach (var layerName in layerOnOffDic)
                {
                    var layerLtr = tr.LayerTable.GetRecord(layerName.Key, OpenMode.ForWrite);
                    if (layerLtr != null)
                    {
                        layerLtr.IsOff = layerName.Value;

                    }
                    VariableDictionary.btnState = false;
                }
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo($"打开图层失败:{ex}");
            }

        }

        /// <summary>
        /// 图层开或关
        /// </summary>
        public static void layerOnOff()
        {
            try
            {
                using var tr = new DBTrans();
                //获取当前所有图层名称并循环委托
                tr.LayerTable.GetRecordNames().ForEach(action: (layerName) =>
                {
                    var layerLtr = tr.LayerTable.GetRecord(layerName, OpenMode.ForWrite);
                    if (layerLtr != null)
                        layerOnOffDic.Add(layerName, layerLtr.IsOff);
                });
                readLayerONOFFState = true;
                tr.Commit();
                Env.Editor.Redraw();

            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("图层开或关失败！");
                LogManager.Instance.LogInfo(ex.Message);
            }

        }

        /// <summary>
        /// 打开图层
        /// </summary>
        [CommandMethod(nameof(OpenLayer))]
        public static void OpenLayer()
        {
            try
            {
                using var tr = new DBTrans();
                tr.LayerTable.GetRecordNames().ForEach(action: (layname) =>
                {
                    foreach (var layer in VariableDictionary.tjtBtn)
                    {
                        if (layer != null)
                        {
                            if ((layname.Contains(layer)))//判断layer图层里是不是有传进来的关键字
                            {
                                var ltr = tr.LayerTable.GetRecord(layname, OpenMode.ForWrite);
                                if (ltr.IsOff == true)
                                {
                                    //ltr.IsOff = false;
                                    ltr.IsFrozen = false;
                                }
                                else
                                {
                                    //ltr.IsOff = true;
                                    ltr.IsFrozen = true;
                                }
                                ; // 关闭图层
                                layname.Print();
                            }
                        }
                    }
                });
                tr.Commit();
                Env.Editor.Redraw();
            }
            catch (Exception ex)
            {
                // 记录错误日志  
                LogManager.Instance.LogInfo("打开图层失败！");
            }
        }

        /// <summary>
        /// 冻结VariableDictionary.selectTjtLayer图层
        /// </summary>
        [CommandMethod(nameof(IsFrozenLayer))]
        public static void IsFrozenLayer()
        {
            try
            {
                using var tr = new DBTrans();
                // 循环遍历图层表中的所有图层名称
                foreach (var layname in tr.LayerTable.GetRecordNames())
                {
                    // 如果图层名称包含 "|"，则取 "|" 后面的部分作为比较名称，否则直接使用图层名称
                    string compareName = layname.Contains("|")
                        ? layname.Split('|')[1]
                        : layname;
                    // 如果比较名称不在 VariableDictionary.selectTjtLayer 中，则跳过该图层
                    if (!VariableDictionary.selectTjtLayer.Contains(compareName))
                    {
                        continue;
                    }
                    // 获取图层表记录并以写入模式打开
                    var ltr = tr.LayerTable.GetRecord(layname, OpenMode.ForWrite);
                    if (ltr == null)
                    {
                        continue;
                    }
                    // 判断当前图层是否将要被冻结
                    bool willFreeze = !ltr.IsFrozen;

                    // AutoCAD 不允许冻结当前图层
                    if (willFreeze && tr.Database.Clayer == ltr.ObjectId)
                    {
                        // 尝试切换到另一个图层
                        ObjectId switchLayerId = ObjectId.Null;

                        // 优先切换到 0 图层
                        if (tr.LayerTable.Has("0") &&
                            !string.Equals(ltr.Name, "0", StringComparison.OrdinalIgnoreCase))
                        {
                            switchLayerId = tr.LayerTable["0"];//获取0图层的ObjectId
                        }
                        else
                        {
                            // 找一个非当前、未冻结、未关闭的图层
                            foreach (var otherLayerName in tr.LayerTable.GetRecordNames())
                            {
                                // 跳过当前图层
                                if (string.Equals(otherLayerName, layname, StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }
                                // 获取其他图层的记录
                                var otherLayer = tr.LayerTable.GetRecord(otherLayerName, OpenMode.ForRead);
                                if (otherLayer != null && !otherLayer.IsFrozen && !otherLayer.IsOff)
                                {
                                    // 找到一个可切换的图层，记录其 ObjectId 并退出循环
                                    switchLayerId = otherLayer.ObjectId;
                                    break;

                                }
                            }
                        }
                        // 如果没有找到可切换的图层，记录信息并跳过冻结当前图层
                        if (switchLayerId.IsNull)
                        {
                            Env.Editor.WriteMessage($"\n图层“{layname}”是当前图层，且没有可切换图层，已跳过冻结。");
                            continue;
                        }
                        // 切换当前图层到可用图层
                        tr.Database.Clayer = switchLayerId;
                    }
                    //
                    ltr.IsFrozen = !ltr.IsFrozen;
                    layname.Print();
                }

                tr.Commit();
                Env.Editor.Regen();
                Env.Editor.Redraw();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo("打开图层失败！");
                LogManager.Instance.LogInfo(ex.Message);
            }
        }

        #endregion

        #region  布局空间检查图层

        [CommandMethod("ChangePlotSetting")]
        public static void ChangePlotSetting()
        {
            //开启事务
            DBTrans tr = new();
            // 引用布局管理器 LayoutManager
            var acLayoutMgr = LayoutManager.Current;
            // 读取当前布局，在命令行窗口显示布局名
            var acLayout = tr.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead) as Layout;
            // 输出当前布局名和设备名
            LogManager.Instance.LogInfo("\nCurrent layout: " + acLayout.LayoutName);
            LogManager.Instance.LogInfo("\nCurrent device name: " + 516 + acLayout.PlotConfigurationName);
            // 从布局中获取 PlotInfo
            PlotInfo acPlInfo = new PlotInfo();
            acPlInfo.Layout = acLayout.ObjectId;
            // 复制布局中的 PlotSettings
            PlotSettings acPlSet = new PlotSettings(acLayout.ModelType);
            acPlSet.CopyFrom(acLayout);
            // 更新 PlotSettings 对象的 PlotConfigurationName 属性
            PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
            acPlSetVdr.SetPlotConfigurationName(acPlSet, "DWF6 ePlot.pc3", "ANSI_(8.50_x_11.00_Inches)");
            // 更新布局
            acLayout.UpgradeOpen();
            acLayout.CopyFrom(acPlSet);
            // 输出已更新的布局设备名
            LogManager.Instance.LogInfo("\nNew device name: " + acLayout.PlotConfigurationName);
            // 将新对象保存到数据库
            tr.Commit();
        }

        /// <summary>
        /// 找到视口中的外部参照图层并冻结
        /// </summary>
        [CommandMethod("FindXrefLayersInViewport")]
        public static void FindXrefLayersInViewport()
        {
            // 获取当前文档的编辑器对象
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            // 提示用户选择一个视口
            PromptEntityOptions peo = new("\n请选择一个布局视口: ");
            peo.SetRejectMessage("\n请选择一个视口对象。\n");
            peo.AddAllowedClass(typeof(Viewport), true);
            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n未选择视口，操作已取消。");
                return;
            }

            // 开始事务
            using DBTrans tr = new();
            try
            {
                // 获取选定的视口对象
                Viewport selectedViewport = tr.GetObject<Viewport>(per.ObjectId, OpenMode.ForWrite);
                if (selectedViewport == null)
                {
                    ed.WriteMessage("\n无法打开选定的视口。");
                    return;
                }
                // 获取当前文档的编辑器对象  
                //DBTrans tr = new();
                // 获取布局管理器单例  
                LayoutManager lm = LayoutManager.Current;
                // 获取当前布局的名称  
                string curLayoutName = lm.CurrentLayout;
                // 根据布局名获取对应的 ObjectId  
                ObjectId layoutId = lm.GetLayoutId(curLayoutName);
                // 打开布局对象以便读取  
                Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                // 获取 PaperSpace 的 BlockTableRecord  
                BlockTableRecord psBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                // 打开 BlockTable  
                BlockTable bt = tr.GetObject(tr.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                // 获取模型空间的  
                BlockTableRecord msBtr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                // 获取图层表  
                var layerTable = tr.GetObject(tr.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                // 遍历模型空间中的所有实体  
                foreach (ObjectId layerTableId in layerTable)
                {
                    // 打开子实体  
                    var layerTableRecord = tr.GetObject(layerTableId, OpenMode.ForRead) as LayerTableRecord;
                    // 如果子实体为空则跳过  
                    if (layerTableRecord == null) continue;
                    // 获取子实体所在的图层名称  
                    string layerName = layerTableRecord.Name;
                    // 检查图层名是否包含“|”符号  
                    int separatorIndex = layerName.IndexOf('|');
                    if (separatorIndex >= 0)
                    {
                        // 获取“|”符号后面的部分  
                        string layerSuffix = layerName.Substring(separatorIndex + 1);
                        // 循环遍历选择的图层  
                        foreach (var selectLayer in VariableDictionary.selectTjtLayer)
                        {
                            // 如果图层后缀与选择的图层完全匹配  
                            if (string.Equals(layerSuffix, selectLayer, StringComparison.OrdinalIgnoreCase))
                            {
                                // 获取图层的 ObjectId  
                                //ObjectId layerId = layerTable[layerName];

                                tr.GetObject(layerTableId, OpenMode.ForWrite);
                                // 冻结图层到视口中  
                                ObjectIdCollection layerIds = new ObjectIdCollection(new[] { layerTableId });
                                // 冻结图层到视口中  
                                // 修复代码以确保 activeVp 对象已打开为写模式，并且 layerIds 不为空。
                                try
                                {
                                    // 确保 activeVp 已打开为写模式
                                    var viewport = tr.GetObject(selectedViewport.ObjectId, OpenMode.ForWrite) as Viewport;
                                    if (selectedViewport == null)
                                    {
                                        throw new InvalidOperationException("无法将视口对象打开为写模式。");
                                    }
                                    // 确保 layerIds 不为空
                                    if (layerIds == null)
                                    {
                                        throw new ArgumentException("图层 ID 列表为空或无效。");
                                    }
                                    // 冻结图层
                                    selectedViewport.FreezeLayersInViewport(layerIds.GetEnumerator());

                                }
                                catch (Exception ex)
                                {
                                    // 处理异常
                                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"错误: {ex.Message}");
                                }

                                LogManager.Instance.LogInfo($"\n图层 '{layerName}' 已在视口中冻结。");
                            }
                        }
                    }
                }
                tr.Commit();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n发生错误: {ex.Message}");
            }
        }
        /// <summary>
        /// 找到视口中的外部参照图层并冻结
        /// </summary>
        [CommandMethod("FindXrefLayersInViewport1")]
        public static void FindXrefLayersInViewport1()
        {
            // 获取当前文档的编辑器对象  
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            // 开始事务  
            using DBTrans tr = new();
            try
            {
                // 获取布局管理器单例  
                LayoutManager lm = LayoutManager.Current;
                // 获取当前布局的名称  
                string curLayoutName = lm.CurrentLayout;
                // 根据布局名获取对应的 ObjectId  
                ObjectId layoutId = lm.GetLayoutId(curLayoutName);
                // 打开布局对象以便读取  
                Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;

                // 检查是否在布局中  
                if (layout != null && !layout.ModelType)
                {
                    // 如果在布局中，直接获取当前视口  
                    Viewport selectedViewport = null;
                    foreach (ObjectId objId in layout.GetViewports())
                    {
                        selectedViewport = tr.GetObject(objId, OpenMode.ForWrite) as Viewport;
                        if (selectedViewport != null && selectedViewport.Number > 1)
                        {
                            break;
                        }
                    }

                    if (selectedViewport == null)
                    {
                        ed.WriteMessage("\n未找到有效的视口。");
                        return;
                    }

                    ProcessViewportLayers(tr, selectedViewport);
                }
                else
                {
                    // 提示用户选择一个视口  
                    PromptEntityOptions peo = new("\n请选择一个布局视口: ");
                    peo.SetRejectMessage("\n请选择一个视口对象。\n");
                    peo.AddAllowedClass(typeof(Viewport), true);
                    PromptEntityResult per = ed.GetEntity(peo);

                    if (per.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\n未选择视口，操作已取消。");
                        return;
                    }

                    // 获取选定的视口对象  
                    Viewport selectedViewport = tr.GetObject<Viewport>(per.ObjectId, OpenMode.ForWrite);
                    if (selectedViewport == null)
                    {
                        ed.WriteMessage("\n无法打开选定的视口。");
                        return;
                    }

                    ProcessViewportLayers(tr, selectedViewport);
                }

                tr.Commit();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n发生错误: {ex.Message}");
            }
        }
        /// <summary>
        /// 处理视口中的图层进行开关
        /// </summary>
        /// <param name="tr"></param>
        /// <param name="selectedViewport"></param>
        private static void ProcessViewportLayers(DBTrans tr, Viewport selectedViewport)
        {
            // 获取图层表  
            var layerTable = tr.GetObject(tr.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
            // 遍历图层表中的所有图层  
            foreach (ObjectId layerTableId in layerTable)
            {
                var layerTableRecord = tr.GetObject(layerTableId, OpenMode.ForRead) as LayerTableRecord;
                if (layerTableRecord == null) continue;

                string layerName = layerTableRecord.Name;
                int separatorIndex = layerName.IndexOf('|');
                if (separatorIndex >= 0)
                {
                    string layerSuffix = layerName.Substring(separatorIndex + 1);
                    foreach (var selectLayer in VariableDictionary.selectTjtLayer)
                    {
                        if (string.Equals(layerSuffix, selectLayer, StringComparison.OrdinalIgnoreCase))
                        {
                            tr.GetObject(layerTableId, OpenMode.ForWrite);
                            ObjectIdCollection layerIds = new ObjectIdCollection(new[] { layerTableId });
                            try
                            {
                                selectedViewport.FreezeLayersInViewport(layerIds.GetEnumerator());
                            }
                            catch (Exception ex)
                            {
                                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"错误: {ex.Message}");
                            }

                            LogManager.Instance.LogInfo($"\n图层 '{layerName}' 已在视口中冻结。");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 找到视口中的外部参照图层并冻结
        /// </summary>
        [CommandMethod("FindXrefLayersInViewportOpen")]
        public static void FindXrefLayersInViewportOpen()
        {
            // 获取当前文档的编辑器对象
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            // 提示用户选择一个视口
            PromptEntityOptions peo = new("\n请选择一个布局视口: ");
            peo.SetRejectMessage("\n请选择一个视口对象。\n");
            peo.AddAllowedClass(typeof(Viewport), true);
            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n未选择视口，操作已取消。");
                return;
            }

            // 开始事务
            using DBTrans tr = new();
            try
            {
                // 获取选定的视口对象
                Viewport selectedViewport = tr.GetObject<Viewport>(per.ObjectId, OpenMode.ForWrite);
                if (selectedViewport == null)
                {
                    ed.WriteMessage("\n无法打开选定的视口。");
                    return;
                }

                // 获取图层表
                var layerTable = tr.GetObject(tr.Database.LayerTableId, OpenMode.ForRead) as LayerTable;

                // 遍历图层表中的所有图层
                foreach (ObjectId layerTableId in layerTable)
                {
                    // 打开图层记录
                    var layerTableRecord = tr.GetObject(layerTableId, OpenMode.ForRead) as LayerTableRecord;
                    if (layerTableRecord == null) continue;

                    // 获取图层名称
                    string layerName = layerTableRecord.Name;

                    // 检查图层名是否包含“|”符号
                    int separatorIndex = layerName.IndexOf('|');
                    if (separatorIndex >= 0)
                    {
                        // 获取“|”符号后面的部分
                        string layerSuffix = layerName.Substring(separatorIndex + 1);

                        // 循环遍历选择的图层
                        foreach (var selectLayer in VariableDictionary.selectTjtLayer)
                        {
                            // 如果图层后缀与选择的图层完全匹配
                            if (string.Equals(layerSuffix, selectLayer, StringComparison.OrdinalIgnoreCase))
                            {
                                // 解冻图层
                                ObjectIdCollection layerIds = new ObjectIdCollection(new[] { layerTableId });
                                try
                                {
                                    // 确保 selectedViewport 已打开为写模式
                                    var viewport = tr.GetObject(selectedViewport.ObjectId, OpenMode.ForWrite) as Viewport;
                                    if (viewport == null)
                                    {
                                        throw new InvalidOperationException("无法将视口对象打开为写模式。");
                                    }

                                    // 解冻图层
                                    selectedViewport.ThawLayersInViewport(layerIds.GetEnumerator());
                                }
                                catch (Exception ex)
                                {
                                    // 处理异常
                                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"错误: {ex.Message}");
                                }

                                LogManager.Instance.LogInfo($"\n图层 '{layerName}' 已在视口中解冻。");
                            }
                        }
                    }
                }
                tr.Commit();
            }
            catch (Exception ex)
            {
                LogManager.Instance.LogInfo($"\n发生错误: {ex.Message}");
            }
        }

        #endregion



        /// <summary>
        /// 确保图层存在
        /// </summary>
        /// <param name="tr">数据库事务</param>
        /// <param name="layerName">图层名</param>
        //private static void EnsureLayerExists(DBTrans tr, string layerName, short colorIndex)
        //{
        //    if (!tr.LayerTable.Has(layerName))
        //    {
        //        //tr.LayerTable.Add(layerName);
        //        // 创建新图层并设置颜色140
        //        tr.LayerTable.Add(layerName, lt =>
        //        {
        //            lt.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
        //        });
        //    }
        //}



        /// <summary>
        ///  解析层颜色
        /// </summary>
        /// <param name="fallback">默认颜色索引</param>
        /// <returns></returns>
        public static short ResolveLayerColor(short fallback = 7)
        {
            try
            {
                short color = Convert.ToInt16(VariableDictionary.layerColorIndex);
                return color > 0 ? color : fallback;
            }
            catch
            {
                return fallback;
            }
        }
        /// <summary>
        /// 解析目标层名称
        /// </summary>
        /// <param name="preferred">首选图层名称</param>
        /// <returns></returns>

        public static string ResolveTargetLayerName(string? preferred = null)
        {
            var layer = preferred
                        ?? VariableDictionary.layerName
                        ?? VariableDictionary.btnBlockLayer
                        ?? "0";
            return string.IsNullOrWhiteSpace(layer) ? "0" : layer;
        }

    }
}
