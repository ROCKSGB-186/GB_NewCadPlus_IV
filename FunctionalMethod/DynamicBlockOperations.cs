using Autodesk.AutoCAD.DatabaseServices;
using GB_NewCadPlus_IV.Helpers;
using GB_NewCadPlus_IV.UniFiedStandards;
using Mysqlx.Crud;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.IO;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using Dm.util;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using AttributeCollection = Autodesk.AutoCAD.DatabaseServices.AttributeCollection;
using DataTable = System.Data.DataTable;

namespace GB_NewCadPlus_IV.FunctionalMethod
{

    /// <summary>
    /// 动态块信息结构
    /// </summary>
    public class DynamicBlockInfo
    {
        public Point3d StartPoint { get; set; }     // 起点坐标
        public Point3d EndPoint { get; set; }       // 终点坐标
        public Point3d MidPoint { get; set; }       // 中点坐标
        public double Length { get; set; }          // 长度
        public double Rotation { get; set; }        // 旋转角度（弧度）
        public Point3d Position { get; set; }       // 插入点坐标
    }
    /// <summary>
    /// 动态块操作类
    /// </summary>
    internal class DynamicBlockOperations
    {
        #region 命令方法示例

        /// <summary>
        /// 测试命令 - 获取动态块信息
        /// </summary>
        [CommandMethod("TestGetBlockInfo")]
        public static void TestGetBlockInfo()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 选择块参照
            PromptEntityOptions opts = new PromptEntityOptions("\n选择动态块:");
            opts.SetRejectMessage("\n请选择块参照.");
            opts.AddAllowedClass(typeof(BlockReference), true);

            PromptEntityResult result = ed.GetEntity(opts);
            if (result.Status != PromptStatus.OK) return;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 获取块参照
                BlockReference blockRef = trans.GetObject(result.ObjectId, OpenMode.ForRead) as BlockReference;

                // 获取块信息
                DynamicBlockInfo info = GetBlockInfo(blockRef);

                // 显示信息
                ed.WriteMessage($"\n===== 动态块信息 =====");
                ed.WriteMessage($"\n起点: {info.StartPoint}");
                ed.WriteMessage($"\n终点: {info.EndPoint}");
                ed.WriteMessage($"\n中点: {info.MidPoint}");
                ed.WriteMessage($"\n长度: {info.Length:F3}");
                ed.WriteMessage($"\n旋转角度: {info.Rotation * 180 / Math.PI:F2} 度");
                ed.WriteMessage($"\n插入点: {info.Position}");

                trans.Commit();
            }
        }

        /// <summary>
        /// 测试命令 - 移动动态块
        /// </summary>
        [CommandMethod("TestMoveBlock")]
        public static void TestMoveBlock()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 选择块参照
            PromptEntityOptions opts = new PromptEntityOptions("\n选择要移动的动态块:");
            opts.SetRejectMessage("\n请选择块参照.");
            opts.AddAllowedClass(typeof(BlockReference), true);

            PromptEntityResult result = ed.GetEntity(opts);
            if (result.Status != PromptStatus.OK) return;

            // 选择新位置
            PromptPointOptions pointOpts = new PromptPointOptions("\n选择新位置:");
            PromptPointResult pointResult = ed.GetPoint(pointOpts);
            if (pointResult.Status != PromptStatus.OK) return;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 获取块参照
                BlockReference blockRef = trans.GetObject(result.ObjectId, OpenMode.ForRead) as BlockReference;

                // 移动块
                MoveBlock(blockRef, pointResult.Value);

                trans.Commit();
            }
        }

        #endregion



        #region 动态块拉伸

        /// <summary>
        /// 绘制管线动态块
        /// </summary>
        [CommandMethod("Draw_GD_PipeLine_DynamicBlock")]
        public void Draw_GD_PipeLine_DynamicBlock()
        {
            // 获取当前文档和数据库
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 开始事务
            using (var tr = new DBTrans())
            {
                try
                {
                    ObjectId insertBlockObjectId = new ObjectId();
                    //插入块
                    //Command.GB_InsertBlock_(new Point3d(0, 0, 0), 0, ref insertBlockObjectId);

                    if (insertBlockObjectId == new ObjectId())
                    {
                        Env.Editor.WriteMessage("未找到块");
                        return;
                    }
                    ;
                    // 打开块参照
                    BlockReference? blockRef = tr.GetObject(insertBlockObjectId, OpenMode.ForWrite) as BlockReference;

                    if (blockRef == null)
                    {
                        doc.Editor.WriteMessage("\n没有找到有效的块参照.");
                        return;
                    }
                    // 第二步：提示用户指定起点
                    PromptPointOptions startPointOpts = new PromptPointOptions("\n请指定拉伸的起点: ");
                    startPointOpts.AllowNone = false; // 不允许直接回车
                    PromptPointResult startPointResult = ed.GetPoint(startPointOpts);
                    if (startPointResult.Status != PromptStatus.OK)
                    {
                        Env.Editor.WriteMessage("\n未指定起点,取消绘制管道");
                        tr.BlockTable.Remove(insertBlockObjectId);
                        return;
                    }

                    // 第三步：提示用户指定终点（使用橡皮筋效果）
                    PromptPointOptions endPointOpts = new PromptPointOptions("\n请指定拉伸的终点: ");
                    endPointOpts.BasePoint = startPointResult.Value; // 设置基点为起点
                    endPointOpts.UseBasePoint = true; // 启用基点
                    endPointOpts.UseDashedLine = true; // 使用虚线显示橡皮筋效果
                    PromptPointResult endPointResult = ed.GetPoint(endPointOpts);
                    if (endPointResult.Status != PromptStatus.OK)
                    {
                        Env.Editor.WriteMessage("\n未指定终点,取消绘制管道");

                        tr.BlockTable.Remove(insertBlockObjectId);
                        return;
                    }
                    //拿到新的起点和终点
                    Point3d newStartPoint = startPointResult.Value;
                    Point3d newEndPoint = endPointResult.Value;
                    double newAngle = 0;
                    if (newEndPoint.X == newStartPoint.X)
                        newAngle = -Math.Atan2(newEndPoint.Y - newStartPoint.Y, newEndPoint.X - newStartPoint.X);
                    if (newEndPoint.Y == newStartPoint.Y)
                        newAngle = Math.Atan2(newEndPoint.Y - newStartPoint.Y, newEndPoint.X - newStartPoint.X);

                    // 确保事务处于写入模式
                    if (!blockRef.IsWriteEnabled) blockRef.UpgradeOpen();

                    DynamicBlockReferencePropertyCollection dynProps = blockRef.DynamicBlockReferencePropertyCollection;

                    //如果起点和终点相同，则退出
                    if (newStartPoint == newEndPoint)
                    {
                        doc.Editor.WriteMessage("\n起点和终点不能相同.");
                        return;
                    }
                    //else if (newStartPoint.X < newEndPoint.X)//如果起点在终点的右侧，则交换起点和终点
                    //{
                    //    newStartPoint = endPointResult.Value;
                    //    newEndPoint = startPointResult.Value;
                    //}
                    //else if (Convert.ToInt32(newStartPoint.X) == Convert.ToInt32(newEndPoint.X) && newStartPoint.Y > newEndPoint.Y)//如果起点在终点的下方，则交换起点和终点
                    //{
                    //    //newStartPoint = endPointResult.Value;
                    //    //newEndPoint = startPointResult.Value;
                    //    newStartPoint = startPointResult.Value;
                    //    newEndPoint = endPointResult.Value;
                    //}
                    if (Convert.ToInt32(newStartPoint.X) == Convert.ToInt32(newEndPoint.X))//如果起点和终点在同一X方向，则移动块到Y轴方向
                    {
                        //移动块
                        MoveBlock(blockRef, new Point3d(newStartPoint.X, (newStartPoint.Y + newEndPoint.Y) / 2, 0));
                        // 设置角度
                        SetDynamicBlockNewAngle(dynProps, newAngle, ed);

                    }
                    else if (Convert.ToInt32(newStartPoint.Y) == Convert.ToInt32(newEndPoint.Y))
                    {
                        //移动块
                        MoveBlock(blockRef, new Point3d((newStartPoint.X + newEndPoint.X) / 2, newStartPoint.Y, 0));
                        // 设置角度
                        SetDynamicBlockNewAngle(dynProps, newAngle, ed);
                    }
                    // 打开块表记录
                    BlockTableRecord? blockTableRecord = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                    // 查找多段线（Polyline）
                    Polyline polyline = new Polyline();
                    if (blockTableRecord != null)
                        foreach (ObjectId entId in blockTableRecord)
                        {
                            Entity? ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                            if (ent != null && ent is Polyline pl)
                            {
                                polyline = pl;
                                break;
                            }
                        }

                    if (polyline == null)
                    {
                        doc.Editor.WriteMessage("\n在动态块中没有找到多段线.");
                        return;
                    }
                    // 获取原始多段线的起点和终点（在块局部坐标系中）
                    var originalStartPoint = polyline.GetPoint3dAt(0);
                    var originalEndPoint = polyline.GetPoint3dAt(polyline.NumberOfVertices - 1);

                    // 转换为世界坐标系
                    var worldStartPoint = blockRef.BlockTransform * originalStartPoint;
                    var worldEndPoint = blockRef.BlockTransform * originalEndPoint;

                    // 计算多段线的原始方向向量
                    Vector3d originalDirection = (worldEndPoint - worldStartPoint).GetNormal();
                    //计算拉伸距离（投影到原始多段线方向）
                    //如果是向右拉伸，则拉伸距离为负数；如果向左拉伸，则拉伸距离为正数
                    var startStretchDistance = (newStartPoint - worldStartPoint).DotProduct(originalDirection);

                    //如果向左拉伸，则拉伸距离为负数；如果向右拉伸，则拉伸距离为正数
                    var endStretchDistance = (newEndPoint - worldEndPoint).DotProduct(originalDirection);

                    // 设置起始点与终点
                    SetDynamicBlockEndPoints(dynProps, startStretchDistance, endStretchDistance, ed);

                    // 提交事务
                    tr.Commit();
                    Env.Editor.Redraw();
                }
                catch (Exception ex)
                {
                    // 处理可能的异常
                    doc.Editor.WriteMessage($"\n发生错误: {ex.Message}");
                }
            }
        }


        /// <summary>
        /// 动态块拉伸
        /// </summary>
        [CommandMethod("StretchDynamicBlockPolyline")]
        public void StretchDynamicBlockPolyline()
        {
            // 获取当前文档和数据库
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 开始事务
            using (var tr = new DBTrans())
            {
                try
                {
                    // 提示用户选择动态块参照
                    PromptEntityOptions promptOpts = new PromptEntityOptions("\n选择要拉伸的动态块:");
                    promptOpts.SetRejectMessage("\n请选择一个块参照.");
                    promptOpts.AddAllowedClass(typeof(BlockReference), true);
                    PromptEntityResult result = doc.Editor.GetEntity(promptOpts);

                    // 如果用户取消选择，则退出
                    if (result.Status != PromptStatus.OK)
                        return;

                    // 打开块参照
                    BlockReference? blockRef = tr.GetObject(result.ObjectId, OpenMode.ForWrite) as BlockReference;

                    if (blockRef == null)
                    {
                        doc.Editor.WriteMessage("\n没有找到有效的块参照.");
                        return;
                    }
                    // 第二步：提示用户指定起点
                    PromptPointOptions startPointOpts = new PromptPointOptions("\n请指定拉伸的起点: ");
                    startPointOpts.AllowNone = false; // 不允许直接回车
                    PromptPointResult startPointResult = ed.GetPoint(startPointOpts);
                    if (startPointResult.Status != PromptStatus.OK)
                        return;
                    // 第三步：提示用户指定终点（使用橡皮筋效果）
                    PromptPointOptions endPointOpts = new PromptPointOptions("\n请指定拉伸的终点: ");
                    endPointOpts.BasePoint = startPointResult.Value; // 设置基点为起点
                    endPointOpts.UseBasePoint = true; // 启用基点
                    endPointOpts.UseDashedLine = true; // 使用虚线显示橡皮筋效果
                    PromptPointResult endPointResult = ed.GetPoint(endPointOpts);
                    if (endPointResult.Status != PromptStatus.OK)
                        return;

                    //拿到新的起点和终点
                    Point3d newStartPoint = startPointResult.Value;
                    Point3d newEndPoint = endPointResult.Value;
                    double newAngle = Math.Atan2(newEndPoint.Y - newStartPoint.Y, newEndPoint.X - newStartPoint.X);
                    // 创建新块参照的克隆
                    BlockReference newBlockRef = (BlockReference)blockRef.Clone();

                    // 添加到当前空间
                    var newBlockRefObjectId = tr.CurrentSpace.AddEntity(newBlockRef);

                    // 确保事务处于写入模式
                    if (!newBlockRef.IsWriteEnabled) newBlockRef.UpgradeOpen();

                    DynamicBlockReferencePropertyCollection dynProps = newBlockRef.DynamicBlockReferencePropertyCollection;

                    //如果起点和终点相同，则退出
                    if (newStartPoint == newEndPoint)
                    {
                        doc.Editor.WriteMessage("\n起点和终点不能相同.");
                        return;
                    }
                    else if (newStartPoint.X > newEndPoint.X)//如果起点在终点的右侧，则交换起点和终点
                    {
                        newStartPoint = endPointResult.Value;
                        newEndPoint = startPointResult.Value;
                    }
                    else if (Convert.ToInt16(newStartPoint.X) == Convert.ToInt16(newEndPoint.X) && newStartPoint.Y > newEndPoint.Y)//如果起点在终点的下方，则交换起点和终点
                    {
                        newStartPoint = endPointResult.Value;
                        newEndPoint = startPointResult.Value;
                    }
                    if (Convert.ToInt16(newStartPoint.X) == Convert.ToInt16(newEndPoint.X))//如果起点和终点在同一X方向，则移动块到Y轴方向
                    {
                        //移动块
                        MoveBlock(newBlockRef, new Point3d(newStartPoint.X, (newStartPoint.Y + newEndPoint.Y) / 2, 0));
                        // 设置角度
                        SetDynamicBlockNewAngle(dynProps, newAngle, ed);

                    }
                    else if (Convert.ToInt16(newStartPoint.Y) == Convert.ToInt16(newEndPoint.Y))
                    {
                        //移动块
                        MoveBlock(newBlockRef, new Point3d((newStartPoint.X + newEndPoint.X) / 2, newStartPoint.Y, 0));
                    }
                    // 打开块表记录
                    BlockTableRecord? blockTableRecord = tr.GetObject(newBlockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                    // 查找多段线（Polyline）
                    Polyline polyline = new Polyline();
                    if (blockTableRecord != null)
                        foreach (ObjectId entId in blockTableRecord)
                        {
                            Entity? ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                            if (ent != null && ent is Polyline pl)
                            {
                                polyline = pl;
                                break;
                            }
                        }

                    if (polyline == null)
                    {
                        doc.Editor.WriteMessage("\n在动态块中没有找到多段线.");
                        return;
                    }
                    // 获取原始多段线的起点和终点（在块局部坐标系中）
                    var originalStartPoint = polyline.GetPoint3dAt(0);
                    var originalEndPoint = polyline.GetPoint3dAt(polyline.NumberOfVertices - 1);

                    // 转换为世界坐标系
                    var worldStartPoint = newBlockRef.BlockTransform * originalStartPoint;
                    var worldEndPoint = newBlockRef.BlockTransform * originalEndPoint;

                    // 计算多段线的原始方向向量
                    Vector3d originalDirection = (worldEndPoint - worldStartPoint).GetNormal();
                    //计算拉伸距离（投影到原始多段线方向）
                    //如果是向右拉伸，则拉伸距离为负数；如果向左拉伸，则拉伸距离为正数
                    var startStretchDistance = (newStartPoint - worldStartPoint).DotProduct(originalDirection);

                    //如果向左拉伸，则拉伸距离为负数；如果向右拉伸，则拉伸距离为正数
                    var endStretchDistance = (newEndPoint - worldEndPoint).DotProduct(originalDirection);

                    // 设置起始点与终点
                    SetDynamicBlockEndPoints(dynProps, startStretchDistance, endStretchDistance, ed);

                    // 提交事务
                    tr.Commit();
                    Env.Editor.Redraw();
                }
                catch (Exception ex)
                {
                    // 处理可能的异常
                    doc.Editor.WriteMessage($"\n发生错误: {ex.Message}");
                }
            }
        }


        #endregion

        /// <summary>
        /// 设置动态块的角度
        /// </summary>
        /// <param name="dynProps">动态块</param>
        /// <param name="newAngle">新角度</param>
        /// <param name="ed">Editor</param>
        private static void SetDynamicBlockNewAngle(DynamicBlockReferencePropertyCollection dynProps, double newAngle, Editor ed)
        {
            //循环动态块的所有属性
            foreach (DynamicBlockReferenceProperty dynProp in dynProps)
            {
                if (dynProp.PropertyName.Contains("管道角度"))
                {
                    if (!dynProp.ReadOnly)
                    {
                        // 输出拉伸参数信息，帮助调试
                        ed.WriteMessage($"\n角度参数: {dynProp.PropertyName}");
                        // 将度转换为弧度
                        //double rotationRadians = pdr.Value * Math.PI / 180.0;

                        // 设置旋转参数
                        dynProp.Value = newAngle;
                        ed.WriteMessage($"\n设置角度：（{newAngle}）");
                        break;
                    }
                }
            }
        }
        
        /// <summary>
        /// 设置动态块的起点和终点坐标
        /// </summary>
        /// <param name="blockRef">块参照对象</param>
        /// <param name="startPoint">起点坐标</param>
        /// <param name="endPoint">终点坐标</param>
        /// <param name="ed">编辑器对象</param>
        private static void SetDynamicBlockEndPoints(DynamicBlockReferencePropertyCollection dynProps, double startPoint, double endPoint, Editor ed)
        {

            foreach (DynamicBlockReferenceProperty dynProp in dynProps)
            {
                // 查找起点相关属性
                if (dynProp.PropertyName.Contains("管道向左拉伸"))
                {
                    if (!dynProp.ReadOnly)
                    {
                        // 输出拉伸参数信息，帮助调试
                        ed.WriteMessage($"\n起点拉伸参数: {dynProp.PropertyName}");
                        if (startPoint <= 0)
                        {
                            // 设置起点坐标
                            dynProp.Value = Convert.ToDouble(dynProp.Value) + Math.Abs(startPoint);
                        }
                        else
                        {
                            //dynProp.Value = startPoint - Convert.ToDouble(dynProp.Value);
                            if (startPoint > Convert.ToDouble(dynProp.Value))
                                dynProp.Value = startPoint - Convert.ToDouble(dynProp.Value);
                            else
                                dynProp.Value = Convert.ToDouble(dynProp.Value) - startPoint;
                        }
                        ed.WriteMessage($"\n设置起点: ({dynProp.Value})");
                    }
                }
                // 查找终点相关属性
                else if (dynProp.PropertyName.Contains("管道向右拉伸"))
                {
                    if (!dynProp.ReadOnly)
                    {
                        // 输出拉伸参数信息，帮助调试
                        ed.WriteMessage($"\n起点拉伸参数: {dynProp.PropertyName}");
                        if (endPoint >= 0)
                        {
                            dynProp.Value = Convert.ToDouble(dynProp.Value) + endPoint;
                        }
                        else
                        {
                            if (Math.Abs(endPoint) > Convert.ToDouble(dynProp.Value))
                                dynProp.Value = Math.Abs(endPoint) - Convert.ToDouble(dynProp.Value);
                            else
                                dynProp.Value = Convert.ToDouble(dynProp.Value) - Math.Abs(endPoint);
                        }
                        ed.WriteMessage($"\n设置终点: ({dynProp.Value})");
                    }
                }
            }
        }

        #region 获取动态块信息的方法

        /// <summary>
        /// 获取动态块的端点坐标（在模型空间中）
        /// </summary>
        /// <param name="blockRef">块参照对象</param>
        /// <returns>包含起点和终点的元组</returns>
        public static (Point3d startPoint, Point3d endPoint) GetEndPoints(BlockReference blockRef)
        {
            // 获取当前数据库
            Database db = blockRef.Database;

            // 初始化返回值
            Point3d worldStartPoint = Point3d.Origin;
            Point3d worldEndPoint = Point3d.Origin;

            // 开始事务
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 获取块表记录
                    BlockTableRecord blockTableRecord = trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                    // 查找多段线
                    foreach (ObjectId entId in blockTableRecord)
                    {
                        Entity ent = trans.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent is Polyline polyline)
                        {
                            // 获取多段线的局部坐标端点
                            Point3d localStartPoint = polyline.GetPoint3dAt(0);
                            Point3d localEndPoint = polyline.GetPoint3dAt(polyline.NumberOfVertices - 1);

                            // 转换到世界坐标系
                            worldStartPoint = blockRef.BlockTransform * localStartPoint;
                            worldEndPoint = blockRef.BlockTransform * localEndPoint;

                            break; // 找到第一个多段线就退出
                        }
                    }

                    // 提交事务
                    trans.Commit();
                }
                catch (Exception ex)
                {
                    // 错误处理
                    Env.Editor.WriteMessage($"\n获取端点时发生错误: {ex.Message}");
                }
            }
            // 返回世界坐标系下的起始点与终点
            return (worldStartPoint, worldEndPoint);
        }

        /// <summary>
        /// 获取动态块的中点坐标（在模型空间中）
        /// </summary>
        /// <param name="blockRef">块参照对象</param>
        /// <returns>中点坐标</returns>
        public static Point3d GetMidPoint(BlockReference blockRef)
        {
            // 获取端点
            var (startPoint, endPoint) = GetEndPoints(blockRef);

            // 计算中点
            Point3d midPoint = new Point3d(
                (startPoint.X + endPoint.X) / 2,
                (startPoint.Y + endPoint.Y) / 2,
                (startPoint.Z + endPoint.Z) / 2
            );
            // 返回中点
            return midPoint;
        }

        /// <summary>
        /// 获取动态块的长度
        /// </summary>
        /// <param name="blockRef">块参照对象</param>
        /// <returns>多段线长度</returns>
        public static double GetLength(BlockReference blockRef)
        {
            // 直接从动态块参数获取长度
            foreach (DynamicBlockReferenceProperty prop in blockRef.DynamicBlockReferencePropertyCollection)
            {
                // 查找包含长度信息的参数
                if (prop.PropertyName.Contains("Length") ||
                    prop.PropertyName.Contains("Distance") ||
                    prop.PropertyName.Contains("Stretch"))
                {
                    try
                    {
                        // 返回参数值作为长度
                        return Convert.ToDouble(prop.Value);
                    }
                    catch
                    {
                        // 如果转换失败，继续查找
                        continue;
                    }
                }
            }
            // 获取端点
            var (startPoint, endPoint) = GetEndPoints(blockRef);
            // 计算距离
            double length = startPoint.DistanceTo(endPoint);
            // 返回距离
            return length;
        }

        /// <summary>
        /// 获取动态块的所有信息
        /// </summary>
        /// <param name="blockRef">块参照对象</param>
        /// <returns>包含所有信息的结构</returns>
        public static DynamicBlockInfo GetBlockInfo(BlockReference blockRef)
        {
            var info = new DynamicBlockInfo();
            // 获取端点
            var (startPoint, endPoint) = GetEndPoints(blockRef);
            info.StartPoint = startPoint;
            info.EndPoint = endPoint;
            // 计算中点
            info.MidPoint = GetMidPoint(blockRef);
            // 计算长度
            info.Length = GetLength(blockRef);
            // 获取旋转角度
            info.Rotation = blockRef.Rotation;
            // 获取插入点
            info.Position = blockRef.Position;
            // 返回所有信息
            return info;
        }

        #endregion


        #region 操作动态块的方法

        /// <summary>
        /// 移动动态块到新位置
        /// </summary>
        /// <param name="blockRef">块参照对象</param>
        /// <param name="newPosition">新位置坐标</param>
        public static void MoveBlock(BlockReference blockRef, Point3d newPosition)
        {
            // 获取当前数据库
            Database db = blockRef.Database;

            // 开始事务
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 打开块参照进行写操作
                    BlockReference br = trans.GetObject(blockRef.ObjectId, OpenMode.ForWrite) as BlockReference;

                    // 计算移动向量
                    Vector3d moveVector = newPosition - br.Position;

                    // 创建移动矩阵
                    Matrix3d moveMatrix = Matrix3d.Displacement(moveVector);

                    // 应用变换
                    br.TransformBy(moveMatrix);

                    // 提交事务
                    trans.Commit();

                    // 输出信息
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n动态块已移动到: {newPosition}");
                }
                catch (Exception ex)
                {
                    // 错误处理
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n移动块时发生错误: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 旋转动态块
        /// </summary>
        /// <param name="blockRef">块参照对象</param>
        /// <param name="rotationAngle">旋转角度（弧度）</param>
        /// <param name="basePoint">旋转基点（可选，默认为块插入点）</param>
        public static void RotateBlock(BlockReference blockRef, double rotationAngle, Point3d? basePoint = null)
        {
            // 获取当前数据库
            Database db = blockRef.Database;

            // 开始事务
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 打开块参照进行写操作
                    BlockReference? br = trans.GetObject(blockRef.ObjectId, OpenMode.ForWrite) as BlockReference;

                    // 确定旋转基点
                    Point3d rotationBase = basePoint ?? br.Position;

                    // 创建旋转矩阵
                    Matrix3d rotationMatrix = Matrix3d.Rotation(rotationAngle, Vector3d.ZAxis, rotationBase);

                    // 应用变换
                    br.TransformBy(rotationMatrix);

                    // 提交事务
                    trans.Commit();

                    // 输出信息
                    double angleDegrees = rotationAngle * 180 / Math.PI;
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n动态块已旋转 {angleDegrees:F2} 度");
                }
                catch (Exception ex)
                {
                    // 错误处理
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n旋转块时发生错误: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 拉伸动态块到指定的端点
        /// </summary>
        /// <param name="blockRef">块参照对象</param>
        /// <param name="newStartPoint">新的起点坐标</param>
        /// <param name="newEndPoint">新的终点坐标</param>
        public static void StretchBlock(BlockReference blockRef, Point3d newStartPoint, Point3d newEndPoint)
        {
            // 获取当前数据库和编辑器
            Database db = blockRef.Database;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            // 开始事务
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 打开块参照进行写操作
                    BlockReference br = trans.GetObject(blockRef.ObjectId, OpenMode.ForWrite) as BlockReference;

                    // 获取当前端点
                    var (currentStartPoint, currentEndPoint) = GetEndPoints(br);

                    // 计算当前方向向量
                    Vector3d currentDirection = (currentEndPoint - currentStartPoint).GetNormal();

                    // 计算拉伸距离
                    double startStretchDistance = (newStartPoint - currentStartPoint).DotProduct(currentDirection);
                    double endStretchDistance = (newEndPoint - currentEndPoint).DotProduct(currentDirection);

                    // 查找并设置动态块参数
                    DynamicBlockReferencePropertyCollection props = br.DynamicBlockReferencePropertyCollection;

                    foreach (DynamicBlockReferenceProperty prop in props)
                    {
                        // 输出所有参数名称（用于调试）
                        ed.WriteMessage($"\n找到参数: {prop.PropertyName}");

                        // 设置拉伸参数
                        if (prop.PropertyName.Contains("Distance") ||
                            prop.PropertyName.Contains("Length") ||
                            prop.PropertyName.Contains("Stretch"))
                        {
                            // 计算新的总长度
                            double newLength = currentStartPoint.DistanceTo(currentEndPoint) +
                                             Math.Abs(startStretchDistance) +
                                             Math.Abs(endStretchDistance);

                            // 尝试设置参数值
                            try
                            {
                                prop.Value = newLength;
                                ed.WriteMessage($"\n设置参数 {prop.PropertyName} 为: {newLength}");
                            }
                            catch
                            {
                                ed.WriteMessage($"\n无法设置参数 {prop.PropertyName}");
                            }
                        }
                    }

                    // 提交事务
                    trans.Commit();

                    ed.WriteMessage($"\n动态块拉伸完成");
                }
                catch (Exception ex)
                {
                    // 错误处理
                    ed.WriteMessage($"\n拉伸块时发生错误: {ex.Message}");
                }
            }
        }


        /// <summary>
        /// 选择并分析属性块或动态属性块，提取设备信息并整理
        /// </summary>
        /// <param name="ed">AutoCAD编辑器对象，用于用户交互</param>
        /// <param name="db">当前图形数据库对象</param>
        /// <returns>设备信息列表，包含从属性块中提取的所有设备数据</returns>
        public static (List<DeviceInfo> deviceList, ObjectId[] selectedIds) SelectAndAnalyzeBlocks(Editor ed, Database db)
        {
            var deviceList =new List<DeviceInfo>();

            // 与原 SelectAndAnalyzeBlocks 保持交互提示一致
            var opts = new PromptSelectionOptions
            {
                MessageForAdding = "\n请选择要生成设备表的属性块或动态属性块（完成后回车）：",
                AllowDuplicates = true
            };

            var selcetElementS = ed.GetSelection(opts);// 获取选择图元实例
            if (selcetElementS.Status != PromptStatus.OK || selcetElementS.Value == null)
                return (deviceList, Array.Empty<ObjectId>());

            var selIds = selcetElementS.Value.GetObjectIds();// 获取选择图元实例Object的Id集合
            if (selIds == null || selIds.Length == 0)
                return (deviceList, Array.Empty<ObjectId>());
            LogManager.Instance.LogInfo($"[设备表][选择完成] SelectedObjectCount={selIds.Length}");
            //开启事务
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject _selectElement in selcetElementS.Value)//循环每个实例
                {
                    try
                    {
                        if (_selectElement == null || _selectElement.ObjectId == ObjectId.Null) continue; // 如果这个选择的实例为空则跳过这个进行下一个实例
                        var _elementBr = tr.GetObject(_selectElement.ObjectId, OpenMode.ForRead) as BlockReference; // 拿到这个图元实例的引用块
                        if (_elementBr == null)
                        {
                            LogManager.Instance.LogWarning($"[设备表][图元跳过] ObjectId={_selectElement.ObjectId}, Reason=不是BlockReference");
                            continue;
                        }

                        var _elementBtr = tr.GetObject(_elementBr.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord; //拿到这个图元的块表记录
                        if (_elementBtr == null) continue;
                        // 初始化一个设备容器
                        var device = new DeviceInfo
                        {
                            Name = _elementBtr.Name ?? string.Empty, //设备名称
                            Type = DetermineDeviceType(_elementBtr.Name), // 类型
                            Attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), //设备属性
                            EnglishNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase), //设备英文名
                            Count = 1 //设备数量
                        };

                        // 读取静态属性
                        foreach (ObjectId _elementAttribute in _elementBr.AttributeCollection) //循环每个图元属性集合,拿到集合中的一个属性值
                        {
                            try
                            {
                                var _elementAttributeRef = tr.GetObject(_elementAttribute, OpenMode.ForRead) as AttributeReference; // 拿到图属性引用
                                if (_elementAttributeRef == null) continue; // 如果这个图元属性引用为空则跳过进行下一下
                                var eARTag = (_elementAttributeRef.Tag ?? string.Empty).Trim(); //拿到图元属性标题
                                var eARVal = (_elementAttributeRef.TextString ?? string.Empty).Trim(); //拿到图元属性的文字
                                if (!string.IsNullOrEmpty(eARTag))//如果图元属性标题不为空
                                {
                                    device.Attributes[eARTag] = eARVal; // 赋值这个设备的表标题
                                    // 如果有中英文字典则映射
                                    device.EnglishNames[eARTag] = (DictionaryHelper.ChineseToEnglish.ContainsKey(eARTag) ? DictionaryHelper.ChineseToEnglish[eARTag] : eARTag);
                                }
                            }
                            catch { /* 忽略单个属性读取失败 */ }
                        }

                         // 读取扩展字典中的 XRecord 属性。
                         // 规范同步、管道属性继承等流程可能将 CONNECTION_TYPE、BOLT_COUNT、BOLT_SPEC 等业务字段保存到 XRecord，
                         // 这些字段不在 AttributeCollection 中，必须与块的静态属性合并后再生成设备表。
                         Dictionary<string, string> xrecordProperties =
                             PipelineEndpointPropertyHelper.ReadEntityProperties(tr, _elementBr);
                         foreach (KeyValuePair<string, string> property in xrecordProperties)
                         {
                             if (string.IsNullOrWhiteSpace(property.Key)) continue;
                             device.Attributes[property.Key.Trim()] = property.Value?.Trim() ?? string.Empty;
                             device.EnglishNames[property.Key.Trim()] =
                                 DictionaryHelper.ChineseToEnglish.TryGetValue(property.Key.Trim(), out string englishName)
                                     ? englishName
                                     : property.Key.Trim();
                         }

                         // 读取动态块属性，兼容连接方式、规格和数量由动态属性提供的块。
                         ProcessDynamicProperties(_elementBr, device);

                         string allAttributes = string.Join("; ", device.Attributes
                             .OrderBy(attribute => attribute.Key, StringComparer.OrdinalIgnoreCase)
                             .Select(attribute => $"{attribute.Key}={attribute.Value}"));
                         LogManager.Instance.LogInfo(
                             $"[设备表][属性读取完成] ObjectId={_elementBr.ObjectId}, BlockName={_elementBtr.Name}, AttributeCount={device.Attributes.Count}, Attributes=[{allAttributes}]");

                        deviceList.add(device);
                    }
                    catch (Exception ex)
                    {
                        LogManager.Instance.LogError($"[设备表][图元分析失败] ObjectId={_selectElement?.ObjectId}, Error={ex.Message}");
                    }
                }

                tr.Commit();
            }

            LogManager.Instance.LogInfo($"[设备表][选择分析完成] DeviceCount={deviceList.Count}, SelectedObjectCount={selIds.Length}");

            return (deviceList, selIds);
        }

        /// <summary>
        /// 处理动态块中的动态属性（如拉伸参数、可见性参数等）
        /// </summary>
        /// <param name="blockRef">动态块引用对象</param>
        /// <param name="device">设备信息对象，用于存储提取的属性值</param>
        private static void ProcessDynamicProperties(BlockReference blockRef, DeviceInfo device)
        {
            try
            {
                if (blockRef == null || device == null)
                    return;

                // 读取动态块属性并保留原始属性名（不加“动态_”前缀），
                // 以便后续逻辑正确识别 CONNECTION_TYPE、BOLT_SPEC、BOLT_COUNT 等字段。
                DynamicBlockReferencePropertyCollection dynProps = blockRef.DynamicBlockReferencePropertyCollection;
                foreach (DynamicBlockReferenceProperty dynProp in dynProps)
                {
                    if (dynProp == null || string.IsNullOrWhiteSpace(dynProp.PropertyName))
                        continue;

                    string propName = dynProp.PropertyName.Trim();
                    string propValue = dynProp.Value?.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(propValue))
                        continue;

                    device.Attributes[propName] = propValue;
                    device.EnglishNames[propName] =
                        DictionaryHelper.ChineseToEnglish.TryGetValue(propName, out string englishName)
                            ? englishName
                            : propName;
                }
            }
            catch (Exception ex)
            {
                // 动态属性读取失败时不中断流程，保留已有静态属性与 XRecord 属性。
                Env.Editor.WriteMessage($"读取动态块属性失败: {ex.Message}");
            }
        }

        #endregion

        #region 获取设备信息
        /// <summary>
        /// 确定设备类型
        /// </summary>
        private static string DetermineDeviceType(string blockName)
        {
            if (string.IsNullOrWhiteSpace(blockName)) return "其他";
            string lowerName = blockName.ToLowerInvariant();

            // 每项包含多个关键字（中文/英文），匹配其中任意一个即判定为该类型（顺序重要：更精确的放前面）
            var mapping = new (string[] keys, string type)[]
            {
                (new[]{ "阀","valve","阀门" }, "阀门"),
                (new[]{ "法兰","flange" }, "法兰"),
                (new[]{ "泵","pump" }, "泵"),
                (new[]{ "流量计","flowmeter","flow meter" }, "流量计"),
                (new[]{ "定位器","positioner" }, "定位器"),
                (new[]{ "执行机构","actuator" }, "执行机构"),
                (new[]{ "仪表","instrument","meter" }, "仪表"),
                (new[]{ "异径管","reducer" }, "异径管"),
                (new[]{ "管帽","pipecap" }, "管帽"),
                (new[]{ "视镜","sightglass" }, "视镜"),
                (new[]{ "膨胀节","expansionjoint" }, "膨胀节"),
                (new[]{ "滤网","filter" }, "滤网"),
                (new[]{ "篮式过滤器","basketfilter" }, "篮式过滤器"),
                (new[]{ "漏斗","funnel" }, "漏斗"),
                (new[]{ "接头","joint","fitting","connector" }, "接头"),
                (new[]{ "取样","sampling" }, "取样"),
                (new[]{ "放空","vent" }, "放空"),
                (new[]{ "排放","discharge" }, "排放"),
                (new[]{ "垫片","gasket" }, "垫片"),
                (new[]{ "螺栓","bolt" }, "螺栓"),
                (new[]{ "螺母","nut" }, "螺母"),
                (new[]{ "泵前/后","pump f","pump b" }, "泵周边件"),
                (new[]{ "管","pipe" }, "管道")
            };

            foreach (var map in mapping)
            {
                foreach (var key in map.keys)
                {
                    if (string.IsNullOrWhiteSpace(key)) continue;
                    if (lowerName.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0)
                        return map.type;
                }
            }

            return "其他";
        }

        /// <summary>
        /// 生成设备唯一键
        /// </summary>
        private static string GenerateDeviceKey(DeviceInfo device)
        {
            return $"{device.Name}_{string.Join("_", device.Attributes.Values)}";
        }
        #endregion
     
    }
}
