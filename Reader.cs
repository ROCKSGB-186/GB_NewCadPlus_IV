namespace GB_NewCadPlus_IV
{
    internal class Reader
    {
        //[CommandMethod(nameof(CreateOutlineCommand))]
        ///// 定义一个 AutoCAD 命令，命令名为 CreateOutlineCommand  
        //public void CreateOutlineCommand()
        //{
        //    PromptSelectionOptions selOpts = new PromptSelectionOptions
        //    {
        //        MessageForAdding = "\n选择图元生成精确轮廓: " // 设置选择图元的提示信息  
        //    };
        //    PromptSelectionResult selRes = Env.Editor.GetSelection(selOpts); // 提示用户选择图元  
        //    if (selRes.Status != PromptStatus.OK) return; // 如果用户取消选择，则退出命令  
        //    double offsetDist = 10.0;  // 固定偏移距离 // 设置固定的偏移距离  
        //    Env.Editor.WriteMessage($"\n将使用偏移距离: {offsetDist}"); // 向命令行输出将使用的偏移距离  
        //    using (var tr = new DBTrans()) // 启动一个事务，确保数据库操作的原子性  
        //    {
        //        try
        //        {
        //            Polyline outline = CreateExactOutline(tr, selRes.Value.GetObjectIds(), offsetDist); // 调用 CreateExactOutline 方法创建精确的外轮廓 
        //            if (outline != null) // 如果外轮廓创建成功  
        //            {
        //                tr.CurrentSpace.AddEntity(outline); // 将外轮廓添加到模型空间  
        //                tr.Commit(); // 提交事务，保存更改  
        //                Env.Editor.Redraw(); // 刷新编辑器，显示结果  
        //                Env.Editor.WriteMessage("\n外轮廓创建成功."); // 向命令行输出成功消息  
        //            }
        //            else
        //            {
        //                Env.Editor.WriteMessage("\n创建外轮廓失败."); // 向命令行输出失败消息  
        //            }
        //        }
        //        catch (System.Exception ex)
        //        {
        //            Env.Editor.WriteMessage("\nError: " + ex.Message); // 捕获异常，并向命令行输出错误信息  
        //            tr.Abort(); // 回滚事务，放弃更改  
        //        }
        //    }
        //}

        ///// <summary>
        ///// 创建外扩线段
        ///// </summary>
        ///// <param name="tr"></param>
        ///// <param name="ids"></param>
        ///// <param name="offsetDist"></param>
        ///// <returns></returns>
        //private Polyline CreateExactOutline(DBTrans tr, ObjectId[] ids, double offsetDist)
        //{
        //    const double scale = 1000.0; // 定义缩放因子，用于提高精度  
        //    Clipper clipper = new Clipper(); // 创建 Clipper 对象，用于多边形裁剪和偏移  
        //    int validEntityCount = 0;
        //    Env.Editor.WriteMessage($"\n选中的图元数量: {ids.Length}");// 调试：打印选中的图元数量  
        //    foreach (ObjectId id in ids) // 遍历所有选中的图元  
        //    {
        //        Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity; // 获取图元对象  
        //        if (ent == null) // 如果图元为空，则跳过  
        //        {
        //            Env.Editor.WriteMessage($"\n跳过空图元: {id}");
        //            continue;
        //        }
        //        Env.Editor.WriteMessage($"\n处理图元类型: {ent.GetType().Name}");
        //        Env.Editor.WriteMessage($"\n图元详细信息: {DescribeEntity(ent)}");
        //        try
        //        {
        //            // 调试：打印每个图元的类型  
        //            Env.Editor.WriteMessage($"\n处理图元类型: {ent.GetType().Name}");

        //            if (ent is BlockReference blkRef)// 如果图元是块参照  
        //            {
        //                TraverseBlock(tr, blkRef, clipper, scale, Matrix3d.Identity); // 递归处理块参照  
        //            }
        //            else
        //            {
        //                AddEntityToClipper(ent, clipper, scale, Matrix3d.Identity); // 将图元添加到 Clipper 对象中  
        //                validEntityCount++;
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            Env.Editor.WriteMessage($"\n处理图元时发生异常: {ex.Message}");
        //        }
        //    }
        //    if (validEntityCount == 0)
        //    {
        //        Env.Editor.WriteMessage("\n警告：没有有效的可处理图元。");
        //        return null;
        //    }

        //    List<List<IntPoint>> unionResult = new List<List<IntPoint>>(); // 创建一个列表，用于存储多边形联合的结果  
        //    clipper.Execute(ClipType.ctUnion, unionResult, PolyFillType.pftNonZero, PolyFillType.pftNonZero);

        //    // 调试：检查 unionResult 是否为空
        //    if (unionResult.Count == 0)
        //    {
        //        Env.Editor.WriteMessage("\n警告：未能生成有效的联合多边形。");
        //        return null;
        //    }

        //    try
        //    {
        //        clipper.Execute(ClipType.ctUnion, unionResult, PolyFillType.pftNonZero, PolyFillType.pftNonZero);

        //        // 调试：打印联合结果详情  
        //        Env.Editor.WriteMessage($"\nunionResult 数量: {unionResult.Count}");
        //        foreach (var path in unionResult)
        //        {
        //            Env.Editor.WriteMessage($"\n路径点数: {path.Count}");
        //            // 可选：打印具体的点坐标  
        //            foreach (var point in path)
        //            {
        //                Env.Editor.WriteMessage($"\n点坐标: ({point.X}, {point.Y})");
        //            }
        //        }

        //        if (unionResult.Count == 0)
        //        {
        //            Env.Editor.WriteMessage("\n警告：未能生成有效的联合多边形。可能原因：");
        //            Env.Editor.WriteMessage("\n1. 输入图形可能不适合布尔操作");
        //            Env.Editor.WriteMessage("\n2. 坐标转换可能存在问题");
        //            Env.Editor.WriteMessage("\n3. 缩放因子可能不合适");
        //            return null;
        //        }

        //        // 后续代码保持不变  
        //        ClipperOffset co = new ClipperOffset(2.0);// 创建 ClipperOffset 对象，用于多边形偏移  
        //        co.AddPaths(unionResult, JoinType.jtRound, EndType.etClosedPolygon);// 设置偏移的连接类型和结束类型  

        //        List<List<IntPoint>> offsetPaths = new List<List<IntPoint>>(); // 创建一个列表，用于存储偏移后的多边形  
        //        co.Execute(ref offsetPaths, offsetDist * scale);

        //        if (offsetPaths.Count > 0 && offsetPaths[0].Count > 0)
        //        {
        //            return ConvertToPolyline(offsetPaths[0], 1.0 / scale);
        //        }
        //        else
        //        {
        //            Env.Editor.WriteMessage("\n警告：未能生成有效的偏移路径。");
        //            return null;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Env.Editor.WriteMessage($"\n布尔操作发生异常: {ex.Message}");
        //        return null;
        //    }
        //}

        ///// <summary>
        ///// 辅助方法来描述实体
        ///// </summary>
        ///// <param name="ent">实体</param>
        ///// <returns></returns>
        //private string DescribeEntity(Entity ent)
        //{
        //    if (ent is Line line)
        //        return $"Line: Start({line.StartPoint.X},{line.StartPoint.Y}), End({line.EndPoint.X},{line.EndPoint.Y})";
        //    if (ent is Polyline pl)
        //        return $"Polyline: Vertices({pl.NumberOfVertices}), Closed({pl.Closed})";
        //    if (ent is Circle circle)
        //        return $"Circle: Center({circle.Center.X},{circle.Center.Y}), Radius({circle.Radius})";
        //    if (ent is Arc arc)
        //        return $"Arc: Center({arc.Center.X},{arc.Center.Y}), Radius({arc.Radius})";
        //    return "Unknown entity type";
        //}

        ///// <summary>
        ///// 横向块
        ///// </summary>
        ///// <param name="tr">事务</param>
        ///// <param name="blkRef">块表记录</param>
        ///// <param name="clipper">Clipper</param>
        ///// <param name="scale">比例</param>
        ///// <param name="mat">Matrix3d</param>
        //private void TraverseBlock(DBTrans tr, BlockReference blkRef, Clipper clipper, double scale, Matrix3d mat)
        //{
        //    BlockTableRecord btr = tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord; // 获取块表记录  
        //    Matrix3d finalMat = blkRef.BlockTransform.PreMultiplyBy(mat); // 计算最终的变换矩阵  

        //    foreach (ObjectId id in btr) // 遍历块内的所有图元  
        //    {
        //        Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity; // 获取图元对象  
        //        if (ent is BlockReference innerBlk) // 如果图元是嵌套的块参照  
        //        {
        //            TraverseBlock(tr, innerBlk, clipper, scale, finalMat); // 递归处理嵌套的块参照  
        //        }
        //        else
        //        {
        //            AddEntityToClipper(ent, clipper, scale, finalMat); // 将图元添加到 Clipper 对象中  
        //        }
        //    }
        //}

        ///// <summary>
        ///// 添加实体到Clipper
        ///// </summary>
        ///// <param name="ent">实体</param>
        ///// <param name="clipper">Clipper</param>
        ///// <param name="scale">比例</param>
        ///// <param name="mat">Matrix3d</param>
        //private void AddEntityToClipper(Entity ent, Clipper clipper, double scale, Matrix3d mat)
        //{
        //    List<List<IntPoint>> paths = new List<List<IntPoint>>(); // 创建一个列表，用于存储图元的路径  

        //    switch (ent) // 根据图元的类型进行处理  
        //    {
        //        case Line line: // 如果图元是直线  
        //            paths.Add(ProcessLine(line, scale, mat)); // 处理直线  
        //            break;
        //        case Polyline pl: // 如果图元是多段线  
        //            paths.Add(ProcessPolyline(pl, scale, mat)); // 处理多段线  
        //            break;
        //        case Circle circle: // 如果图元是圆  
        //            paths.Add(ProcessCircle(circle, scale, mat)); // 处理圆  
        //            break;
        //        case Arc arc: // 如果图元是圆弧  
        //            paths.Add(ProcessArc(arc, scale, mat)); // 处理圆弧  
        //            break;
        //    }
        //    foreach (var path in paths) // 遍历所有路径  
        //    {
        //        clipper.AddPath(path, PolyType.ptSubject, true); // 将路径添加到 Clipper 对象中  
        //    }
        //}

        ///// <summary>
        ///// 处理线
        ///// </summary>
        ///// <param name="line">线</param>
        ///// <param name="scale">比例</param>
        ///// <param name="mat">Matrix3d</param>
        ///// <returns></returns>
        //private List<IntPoint> ProcessLine(Line line, double scale, Matrix3d mat)
        //{
        //    Point3d start = line.StartPoint.TransformBy(mat); // 获取直线的起点，并进行变换  
        //    Point3d end = line.EndPoint.TransformBy(mat); // 获取直线的终点，并进行变换  
        //    Vector3d dir = (end - start).GetNormal(); // 计算直线的方向向量  
        //    Vector3d perp = new Vector3d(-dir.Y, dir.X, 0) * line.Thickness * 0.5; // 计算垂直于直线的向量，考虑线宽  
        //    return new List<IntPoint> // 返回一个包含四个 IntPoint 的列表，表示矩形的四个角点  
        //    {
        //        SafeToIntPoint(start + perp, scale), // 计算并添加一个角点  
        //        SafeToIntPoint(end + perp, scale), // 计算并添加一个角点  
        //        SafeToIntPoint(end - perp, scale), // 计算并添加一个角点  
        //        SafeToIntPoint(start - perp, scale)  // 计算并添加一个角点  
        //    };
        //}

        ///// <summary>
        ///// 外理Polyline
        ///// </summary>
        ///// <param name="pl">Polyline</param>
        ///// <param name="scale">比例</param>
        ///// <param name="mat">Matrix3d</param>
        ///// <returns></returns>
        //private List<IntPoint> ProcessPolyline(Polyline pl, double scale, Matrix3d mat)
        //{
        //    List<IntPoint> path = new List<IntPoint>(); // 创建一个列表，用于存储多段线的路径  
        //    for (int i = 0; i < pl.NumberOfVertices; i++) // 遍历多段线的每个顶点  
        //    {
        //        if (pl.GetSegmentType(i) == SegmentType.Arc) // 如果当前段是圆弧  
        //        {
        //            CircularArc2d arc = pl.GetArcSegment2dAt(i); // 获取圆弧段  
        //            path.AddRange(DiscretizeArc(arc, scale, mat)); // 将圆弧离散化为一系列点  
        //        }
        //        else
        //        {
        //            Point2d pt = pl.GetPoint2dAt(i); // 获取顶点坐标  
        //            path.Add(SafeToIntPoint(new Point3d(pt.X, pt.Y, 0).TransformBy(mat), scale)); // 将顶点坐标转换为 IntPoint，并进行变换  
        //        }
        //    }
        //    return path; // 返回多段线的路径  
        //}

        ///// <summary>
        ///// 处理圆
        ///// </summary>
        ///// <param name="circle">圆</param>
        ///// <param name="scale">比例</param>
        ///// <param name="mat">Matrix3d</param>
        ///// <returns></returns>
        //private List<IntPoint> ProcessCircle(Circle circle, double scale, Matrix3d mat)
        //{
        //    List<IntPoint> points = new List<IntPoint>(); // 创建一个列表，用于存储圆的离散点  

        //    // 额外的输入有效性检查  
        //    if (circle == null)
        //        throw new ArgumentNullException(nameof(circle), "圆对象不能为空");

        //    // 检查圆心坐标是否有效  
        //    if (double.IsNaN(circle.Center.X) || double.IsNaN(circle.Center.Y) ||
        //        double.IsInfinity(circle.Center.X) || double.IsInfinity(circle.Center.Y))
        //    {
        //        throw new ArgumentException($"圆心坐标无效: ({circle.Center.X}, {circle.Center.Y})");
        //    }

        //    // 检查半径是否合法  
        //    if (circle.Radius <= 0 || double.IsNaN(circle.Radius) || double.IsInfinity(circle.Radius))
        //    {
        //        throw new ArgumentException($"圆半径无效: {circle.Radius}");
        //    }

        //    Point3d transformedCenter = circle.Center.TransformBy(mat); // 获取圆心坐标，并进行变换  
        //    double transformedRadius = circle.Radius; // 获取圆的半径  

        //    // 限制离散化的段数，防止过大的坐标  
        //    int segments = Math.Min(Math.Max(8, (int)(transformedRadius / 10)), 128);

        //    for (int i = 0; i < segments; i++) // 遍历每个段  
        //    {
        //        double angle = 2 * Math.PI * i / segments; // 计算角度  
        //        double x = transformedCenter.X + transformedRadius * Math.Cos(angle); // 计算 x 坐标  
        //        double y = transformedCenter.Y + transformedRadius * Math.Sin(angle); // 计算 y 坐标  

        //        try
        //        {
        //            // 使用更安全的坐标转换方法  
        //            points.Add(SafeToIntPoint(new Point3d(x, y, 0), scale));
        //        }
        //        catch (Exception ex)
        //        {
        //            // 记录详细的错误信息  
        //            throw new InvalidOperationException(
        //                $"圆坐标转换失败。圆心: ({transformedCenter.X}, {transformedCenter.Y}), " +
        //                $"半径: {transformedRadius}, 角度: {angle}, " +
        //                $"坐标: ({x}, {y}), 比例尺: {scale}", ex);
        //        }
        //    }

        //    return points; // 返回圆的离散点列表  
        //}

        ///// <summary>
        ///// 处理弧
        ///// </summary>
        ///// <param name="arc">弧</param>
        ///// <param name="scaleFactor">比例因子</param>
        ///// <param name="mat">Matrix3d</param>
        ///// <returns></returns>
        ///// <exception cref="ArgumentException"></exception>
        //private List<IntPoint> ProcessArc(Arc arc, double scaleFactor, Matrix3d mat)
        //{
        //    List<IntPoint> points = new List<IntPoint>(); // 创建一个列表，用于存储圆弧的离散点  

        //    CoordinateSystem3d coordSys = mat.CoordinateSystem3d; // 获取坐标系  
        //    Vector3d xAxis = coordSys.Xaxis; // 获取 X 轴向量  
        //    Vector3d yAxis = coordSys.Yaxis; // 获取 Y 轴向量  
        //    double scaleX = xAxis.Length; // 获取 X 轴的缩放比例  
        //    double scaleY = yAxis.Length; // 获取 Y 轴的缩放比例  

        //    if (Math.Abs(scaleX - scaleY) > 1e-6) // 如果 X 轴和 Y 轴的缩放比例不一致  
        //    {
        //        throw new ArgumentException("Non-uniform scaling - arcs may become ellipses."); // 抛出异常，提示不支持非均匀缩放  
        //    }

        //    bool isMirrored = xAxis.DotProduct(Vector3d.XAxis) < 0; // 判断是否镜像  
        //    double sign = isMirrored ? -1 : 1; // 如果镜像，则 sign 为 -1，否则为 1  
        //    double effectiveRadius = arc.Radius * scaleX; // 计算有效半径  

        //    double rotationAngle = GetZRotationFromMatrix(mat); // 获取 Z 轴的旋转角度  

        //    double startAngle = arc.StartAngle + rotationAngle; // 计算起始角度  
        //    double endAngle = arc.EndAngle + rotationAngle; // 计算终止角度  

        //    if (isMirrored) // 如果镜像  
        //    {
        //        startAngle = Math.PI - startAngle; // 调整起始角度  
        //        endAngle = Math.PI - endAngle; // 调整终止角度  
        //        (startAngle, endAngle) = (endAngle, startAngle); // 交换起始角度和终止角度  
        //    }

        //    int segments = CalculateArcSegments(startAngle, endAngle, effectiveRadius); // 计算离散化的段数  
        //    for (int i = 0; i <= segments; i++) // 遍历每个段  
        //    {
        //        double angle = startAngle + (endAngle - startAngle) * i / segments; // 计算角度  
        //        double x = arc.Center.X + effectiveRadius * Math.Cos(angle); // 计算 x 坐标  
        //        double y = arc.Center.Y + effectiveRadius * Math.Sin(angle); // 计算 y 坐标  
        //        points.Add(SafeToIntPoint(new Point3d(x, y, 0), scaleFactor)); // 将坐标转换为 IntPoint，并添加到列表中  
        //    }
        //    return points; // 返回圆弧的离散点列表  
        //}

        ///// <summary>
        ///// 从矩阵中获取旋转
        ///// </summary>
        ///// <param name="mat">Matrix3d</param>
        ///// <returns></returns>
        ///// <exception cref="ArgumentException"></exception>
        //private double GetZRotationFromMatrix(Matrix3d mat)
        //{
        //    CoordinateSystem3d coordSys = mat.CoordinateSystem3d; // 获取坐标系  
        //    Vector3d xAxis = coordSys.Xaxis; // 获取 X 轴向量  
        //    if (xAxis.Length < Tolerance.Global.EqualPoint) // 如果 X 轴向量的长度小于容差  
        //    {
        //        throw new ArgumentException("X axis vector has zero length."); // 抛出异常，提示 X 轴向量长度为零  
        //    }
        //    xAxis = xAxis.DivideBy(xAxis.Length); // 对 X 轴向量进行归一化  

        //    double dot = xAxis.DotProduct(Vector3d.XAxis); // 计算 X 轴向量与全局 X 轴向量的点积  
        //    double angle = Math.Acos(Math.Min(1.0, Math.Max(-1.0, dot))); // 计算角度  

        //    if (xAxis.Y < -Tolerance.Global.EqualPoint) // 如果 Y 轴分量小于容差  
        //    {
        //        angle = -angle; // 调整角度  
        //    }
        //    else if (xAxis.Y > Tolerance.Global.EqualPoint) // 如果 Y 轴分量大于容差  
        //    {
        //        angle = 2 * Math.PI - angle; // 调整角度  
        //    }
        //    return angle; // 返回 Z 轴的旋转角度  
        //}

        ///// <summary>
        ///// 计算分段弧
        ///// </summary>
        ///// <param name="start">开始</param>
        ///// <param name="end">结束</param>
        ///// <param name="radius">半径</param>
        ///// <returns></returns>
        //private int CalculateArcSegments(double start, double end, double radius)
        //{
        //    double angleSweep = Math.Abs(end - start); // 计算角度差  
        //    int minSegments = 8; // 设置最小段数  
        //    int radiusBased = (int)Math.Ceiling(radius / 10.0) * 4; // 基于半径计算段数  
        //    return Math.Max(minSegments, (int)(angleSweep / (5.0))); // 返回最大段数  
        //}

        ///// <summary>
        ///// 离散弧
        ///// </summary>
        ///// <param name="arc">圆弧</param>
        ///// <param name="scale">比例</param>
        ///// <param name="mat">Matrix3d</param>
        ///// <returns></returns>
        ///// <exception cref="ArgumentException"></exception>
        //private IEnumerable<IntPoint> DiscretizeArc(CircularArc2d arc, double scale, Matrix3d mat)
        //{
        //    CoordinateSystem3d coordSys = mat.CoordinateSystem3d; // 获取坐标系  
        //    Vector3d xAxis = coordSys.Xaxis; // 获取 X 轴向量  
        //    Vector3d yAxis = coordSys.Yaxis; // 获取 Y 轴向量  

        //    double scaleX = xAxis.Length; // 获取 X 轴的缩放比例  
        //    double scaleY = yAxis.Length; // 获取 Y 轴的缩放比例  

        //    if (Math.Abs(scaleX - scaleY) > 1e-6) // 如果 X 轴和 Y 轴的缩放比例不一致  
        //    {
        //        throw new ArgumentException("Non-uniform scaling not supported for arcs."); // 抛出异常，提示不支持非均匀缩放  
        //    }
        //    double actualScale = scaleX; // 获取实际的缩放比例  

        //    double rotationAngle = GetZRotationAngle(coordSys); // 获取 Z 轴的旋转角度  
        //    bool isMirrored = xAxis.DotProduct(Vector3d.XAxis) < 0; // 判断是否镜像  

        //    Point3d center3d = new Point3d(arc.Center.X, arc.Center.Y, 0).TransformBy(mat); // 获取圆心坐标，并进行变换  

        //    double startAngle = arc.StartAngle + rotationAngle; // 计算起始角度  
        //    double endAngle = arc.EndAngle + rotationAngle; // 计算终止角度  

        //    if (isMirrored) // 如果镜像  
        //    {
        //        startAngle = Math.PI - startAngle; // 调整起始角度  
        //        endAngle = Math.PI - endAngle; // 调整终止角度  
        //        (startAngle, endAngle) = (endAngle, startAngle); // 交换起始角度和终止角度  
        //    }

        //    if (arc.IsClockWise) // 如果是顺时针圆弧  
        //    {
        //        double sweep = startAngle - endAngle; // 计算扫描角度  
        //        if (sweep < 0) sweep += 2 * Math.PI; // 如果扫描角度小于 0，则加上 2π  
        //        endAngle = startAngle - sweep; // 调整终止角度  
        //    }

        //    double angleSweep = endAngle - startAngle; // 计算角度差  
        //    int segments = Math.Max(8, (int)(angleSweep / (5.0 * Math.PI / 180))); // 计算离散化的段数  

        //    for (int i = 0; i <= segments; i++) // 遍历每个段  
        //    {
        //        double angle = startAngle + angleSweep * i / segments; // 计算角度  
        //        double x = center3d.X + arc.Radius * actualScale * Math.Cos(angle); // 计算 x 坐标  
        //        double y = center3d.Y + arc.Radius * actualScale * Math.Sin(angle); // 计算 y 坐标  
        //        yield return SafeToIntPoint(new Point3d(x, y, 0), scale); // 将坐标转换为 IntPoint，并返回  
        //    }
        //}

        ///// <summary>
        ///// 获取Z旋转角度
        ///// </summary>
        ///// <param name="coordSys">坐标系统3d</param>
        ///// <returns></returns>
        //private double GetZRotationAngle(CoordinateSystem3d coordSys)
        //{
        //    Vector3d xAxisProj = new Vector3d(coordSys.Xaxis.X, coordSys.Xaxis.Y, 0); // 获取 X 轴在 XY 平面的投影  
        //    if (xAxisProj.Length < 1e-6) return 0; // 如果投影长度小于容差，则返回 0  
        //    double dot = xAxisProj.DotProduct(Vector3d.XAxis); // 计算投影与全局 X 轴的点积  
        //    double angle = Math.Acos(dot / xAxisProj.Length); // 计算角度  
        //    if (coordSys.Xaxis.Y < 0) angle = -angle; // 如果 Y 轴分量小于 0，则调整角度  
        //    if (coordSys.Xaxis.DotProduct(Vector3d.XAxis) < 0) // 如果 X 轴与全局 X 轴的点积小于 0  
        //    {
        //        angle = Math.PI - angle; // 调整角度  
        //    }
        //    return angle; // 返回 Z 轴的旋转角度  
        //}

        ///// <summary>
        ///// 获取int坐标
        ///// </summary>
        ///// <param name="pt">Point3d</param>
        ///// <param name="scale">比例</param>
        ///// <returns></returns>
        ///// <exception cref="ArgumentException"></exception>
        //private IntPoint SafeToIntPoint(Point3d pt, double scale)
        //{
        //    Env.Editor.WriteMessage($"\n转换坐标: ({pt.X}, {pt.Y}), 缩放: {scale}");
        //    try
        //    {
        //        // 首先检查原始坐标是否合理  
        //        if (double.IsNaN(pt.X) || double.IsNaN(pt.Y) ||
        //            double.IsInfinity(pt.X) || double.IsInfinity(pt.Y))
        //        {
        //            throw new ArgumentException($"Invalid coordinate: ({pt.X}, {pt.Y})");
        //        }

        //        // 计算最大可安全缩放的值  
        //        double safeMaxValue = Math.Sqrt(Int64.MaxValue);

        //        // 检查原始坐标的绝对值是否超出安全范围  
        //        if (Math.Abs(pt.X) > safeMaxValue / scale ||
        //            Math.Abs(pt.Y) > safeMaxValue / scale)
        //        {
        //            throw new OverflowException(
        //                $"Coordinate ({pt.X}, {pt.Y}) is too large for scale {scale}. " +
        //                $"Maximum safe value is {safeMaxValue / scale}.");
        //        }

        //        // 进行安全缩放  
        //        //double x = pt.X * scale;
        //        //double y = pt.Y * scale;

        //        // 使用 checked 块确保转换安全  
        //        //try
        //        //{
        //        //    return new IntPoint(
        //        //        checked((long)Math.Round(x)),
        //        //        checked((long)Math.Round(y))
        //        //    );
        //        //}
        //        //catch (OverflowException ex)
        //        //{
        //        //    throw new ArgumentException(
        //        //        $"Coordinate conversion error: {ex.Message}. " +
        //        //        $"Scaled coordinates: ({x}, {y})", ex);
        //        //}
        //        try
        //        {
        //            long x = checked((long)Math.Round(pt.X * scale));
        //            long y = checked((long)Math.Round(pt.Y * scale));

        //            Env.Editor.WriteMessage($"\n转换结果: ({x}, {y})");

        //            return new IntPoint(x, y);
        //        }
        //        catch (Exception ex)
        //        {
        //            Env.Editor.WriteMessage($"\n坐标转换异常: {ex.Message}");
        //            throw;
        //        }
        //    }
        //    catch (OverflowException ex)
        //    {
        //        throw new ArgumentException("Coordinate conversion error: " + ex.Message); // 捕获 OverflowException 异常，并抛出 ArgumentException 异常  
        //    }
        //}

        ///// <summary>
        ///// 生成PL线
        ///// </summary>
        ///// <param name="path"></param>
        ///// <param name="inverseScale">逆比例</param>
        ///// <returns></returns>
        //private Polyline ConvertToPolyline(List<IntPoint> path, double inverseScale)
        //{
        //    Polyline pl = new Polyline(); // 创建一个新的 Polyline 对象  
        //    if (path == null) return pl; //Handle null path 如果路径为空，则返回一个空的 Polyline 对象  

        //    for (int i = 0; i < path.Count; i++) // 遍历路径中的每个点  
        //    {
        //        pl.AddVertexAt(i, new Point2d( // 将点添加到 Polyline 对象中  
        //            path[i].X * inverseScale, // 计算 X 坐标  
        //            path[i].Y * inverseScale), 0, 0, 0); // 计算 Y 坐标  
        //    }
        //    if (path.Count > 2) // ensure it is valid 确保路径有效，至少包含 3 个点  
        //        pl.Closed = true; // 如果路径有效，则将 Polyline 对象设置为闭合  
        //    return pl; // 返回 Polyline 对象  
        //}



        /*  使用CAD内部命令
          21.2 使用CAD内部命令

          SendStringToExecute()函数会延时执行（非同步），它在.NET命令结束时才会被调用。

          1,var doc=Acap.DocumentManager.MdiActiveDocument;
          2,string fileName = "D:\\test.dwg";
          3,doc.SendStringToExecute("Saveas\n"+"LT2004\n"+fileName+"\n",true,false,false);

          需要同步执行时，可以用Editor.RunLisp()方法，设置RunLispFlag特性为      
          RunLispFlag.AcedEvaluateLisp即可同步执行。

          Env.Print("\n123");
          Env.Editor.RunLisp("(princ \"\n456\")", EditorEx.RunLispFlag.AcedEvaluateLisp);
          Env.Print("\n789");

        执行结果为：
          命令: TEST
          123
          456
          789

        在执行命令前取消在执行的多个嵌套命令，可以参考以下方法。
          //创建Esc命令 By edata 
          string esc = "";
          string cmds = (string)Acap.GetSystemVariable("CMDNAMES");
          if (cmds.Length > 0)
          {
              int cmdNum = cmds.Split(new char[] { '\'' }).Length;
              for (int i = 0; i < cmdNum; i++)
                  esc += '\x03';
          }
          doc.SendStringToExecute(esc + cmdbtn.CmdStr + "\n", true, false, true);
        */

        /* 选择外部参照 、子实体选择、子实体特性选择
          // 第一步：让用户点击选择外部参照中的一个图元
                PromptEntityOptions peo = new PromptEntityOptions("\n请选择外部参照中的图元: ");
                peo.AllowObjectOnLockedLayer = true; // 允许选择锁定图层
                peo.AllowNone = false; // 不允许不选择
                peo.SetRejectMessage("\n请选择外部参照中的图元。");
                peo.AddAllowedClass(typeof(Entity), true); // 只允许选择实体类型
                peo.AddAllowedClass(typeof(BlockReference), true); // 允许选择块参照
                peo.AddAllowedClass(typeof(SubentRef), true);// 允许选择子实体引用
                peo.AddAllowedClass(typeof(SubentityType), true); // 允许选择子实体类型
                peo.AddAllowedClass(typeof(SubEntityTraits), true); // 允许选择子实体特性
                peo.AddAllowedClass(typeof(SubentityOverrule), true); // 允许选择子实体重写
                peo.AddAllowedClass(typeof(SubentityId), true); // 允许选择子实体ID
                peo.AddAllowedClass(typeof(SubentityGeometry), true); // 允许选择子实体几何
                peo.AddAllowedClass(typeof(SubDMesh), true); // 允许选择子D网格
                peo.AddAllowedClass(typeof(SubEntity), true); // 允许选择子实体
                peo.AddAllowedClass(typeof(SubentityType), true); // 允许选择子实体类型

                peo.AddAllowedClass(typeof(Polyline), true); // 允许选择多段线 
                peo.AddAllowedClass(typeof(Line), true); // 允许选择线段
                peo.AddAllowedClass(typeof(Circle), true); // 允许选择圆
                peo.AddAllowedClass(typeof(Arc), true); // 允许选择圆弧
                peo.AddAllowedClass(typeof(Ellipse), true); // 允许选择椭圆
                peo.AddAllowedClass(typeof(Solid), true); // 允许选择实体
                peo.AddAllowedClass(typeof(MText), true); // 允许选择多行文本
                peo.AddAllowedClass(typeof(Leader), true); // 允许选择引线
                peo.AddAllowedClass(typeof(Hatch), true); // 允许选择填充
                peo.AddAllowedClass(typeof(Polyline2d), true); // 允许选择二维多段线
                peo.AddAllowedClass(typeof(Polyline3d), true); // 允许选择三维多段线
                peo.AddAllowedClass(typeof(Spline), true); // 允许选择样条曲线
                peo.AddAllowedClass(typeof(Region), true); // 允许选择区域
                peo.AddAllowedClass(typeof(Solid3d), true); // 允许选择三维实体
                peo.AddAllowedClass(typeof(Mesh), true); // 允许选择网格
                peo.AddAllowedClass(typeof(Trace), true); // 允许选择追踪
                peo.AddAllowedClass(typeof(Leader), true); // 允许选择引线
                peo.AddAllowedClass(typeof(Annotation), true); // 允许选择注释对象
                peo.AddAllowedClass(typeof(Underlay), true); // 允许选择附加图像
                peo.AddAllowedClass(typeof(UnderlayImage), true); // 允许选择图像附加图像
                peo.AddAllowedClass(typeof(UnderlayDwf), true); // 允许选择DWF附加图像
                peo.AddAllowedClass(typeof(UnderlayDwfx), true); // 允许选择DWFx附加图像
                peo.AddAllowedClass(typeof(UnderlaySvg), true); // 允许选择SVG附加图像
                peo.AddAllowedClass(typeof(UnderlayDgn), true); // 允许选择DGN附加图像
                peo.AddAllowedClass(typeof(UnderlayPdf), true); // 允许选择PDF附加图像
         */

        /* 以下是 PromptEntityOptions 的一些常用属性和方法：
        //Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        //Editor ed = doc.Editor;

         PromptEntityOptions 是用于提示用户选择实体的选项类
         PromptEntityOptions 是用于提示用户选择实体的选项类。它提供了一些属性和方法，用于设置和获取选择的实体的条件和选项。

        以下是 PromptEntityOptions 的一些常用属性和方法：

        Message：用于设置提示消息，向用户说明需要选择的实体的类型或其他相关信息。
        AllowNone：用于设置是否允许用户选择空实体。默认值为 false，即不允许选择空实体。
        AllowObjectSnap：用于设置是否允许用户通过对象捕捉来选择实体。默认值为 true，即允许使用对象捕捉。
        AddAllowedClass：用于添加允许选择的实体类型。可以通过 RXClass.GetClass() 方法获取实体类型。例如，AddAllowedClass(RXClass.GetClass(typeof(Line))) 表示允许选择直线实体。
        SetRejectMessage：用于设置当用户选择不符合条件的实体时显示的错误消息。
        // 选择外部参照
        //方法一
        //PromptEntityOptions opt = new PromptEntityOptions("选择一个标注实体 ");
        //PromptEntityResult res = ed.GetEntity(opt);
        //opt.SetRejectMessage("您必须选择一个实体 ");
        //opt.AddAllowedClass(typeof(BlockReference), true);//如果是想要的类型，就加入到opt内

         */

        /* 演示如何将一个文件以实体图元的形式插入到当前图中

         在C#中，可以使用AutoCAD的COM接口来进行二次开发。以下是一个示例代码，演示如何将一个文件以实体图元的形式插入到当前图中：

        [assembly: CommandClass(typeof(InsertEntityCommand))]

        public class InsertEntityCommand
        {
         [CommandMethod("InsertEntity")]
            public void InsertEntity()
             {
        // 获取当前文档和数据库
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;

        // 提示用户选择要插入的文件
        PromptOpenFileOptions options = new PromptOpenFileOptions("选择要插入的文件");
        options.Filter = "Drawing Files (*.dwg)|*.dwg";
        PromptFileNameResult result = doc.Editor.GetFileNameForOpen(options);

        if (result.Status == PromptStatus.OK)
        {
            string filePath = result.StringResult;

            // 开始事务
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 打开要插入的文件
                Database insertDb = new Database(false, true);
                insertDb.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, false, "");

                // 获取要插入的实体
                using (BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable)
                {
                    using (BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord)
                    {
                        // 创建一个块参照
                        BlockReference blockRef = new BlockReference(Point3d.Origin, insertDb.BlockTableId);
                        btr.AppendEntity(blockRef);
                        tr.AddNewlyCreatedDBObject(blockRef, true);

                        // 将块参照的属性设置为要插入的实体
                        using (BlockTableRecord insertBtr = tr.GetObject(insertDb.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord)
                        {
                            foreach (ObjectId objId in insertBtr)
                            {
                                Entity entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                                if (entity != null)
                                {
                                    Entity entityCopy = entity.Clone() as Entity;
                                    blockRef.AttributeCollection.AppendAttribute(entityCopy);
                                    tr.AddNewlyCreatedDBObject(entityCopy, true);
                                }
                            }
                        }
                    }
                }

                // 提交事务
                tr.Commit();
            }

            // 提示插入成功
            doc.Editor.WriteMessage("文件已成功插入到当前图中。");
        }
        else
        {
            // 提示用户未选择文件
            doc.Editor.WriteMessage("未选择要插入的文件。");
        }
    }
    }

         */

        /* 将一个 `Point3d` 结构体的三维坐标转换为当前视口平面上的二维坐标
         * 在CAD二次开发中，有时需要将一个 `Point3d` 结构体的三维坐标转换为当前视口平面上的二维坐标。
         * 可以使用 `Vector3d` 类提供的 `GetVectorTo` 方法和 `Viewport` 对象的 `Convert2d` 方法实现这个目的。
         * 下面是一个示例代码：

        public Point2d Point3DTo2D(Point3d point, Viewport viewport)
        {
            Vector3d vector = viewport.ViewDirection;
            Point3d origin = viewport.ViewTarget;
            vector = vector.TransformBy(viewport.ViewportTwist);
            vector = vector.GetNormal();
            Point3d delta = point - origin;
            double distance = delta.DotProduct(vector);
            Point3d projectedPoint = origin + distance * vector;
            Point2d resultPoint = viewport.Convert2d(projectedPoint);
            return resultPoint;
        }

        在此示例代码中，`Vector3d` 类的 `GetVectorTo` 方法用于计算从观察点到点 `point` 的向量。
        `viewport.ViewDirection` 属性获取当前视口的观察方向向量，`viewport.ViewTarget` 
        属性获取当前视口的观察目标点。`TransformBy` 方法用于应用视口的视图旋转。
        与此相应的，`viewport.ViewportTwist` 属性是当前视口的旋转角度。`GetNormal` 方法返回向量的单位向量。

        在得到从观察点到点 `point` 的向量后，可以计算出该向量在当前视口的投影距离 `distance`。
        然后，可以计算出在当前视口投影上的点 `projectedPoint`。
        最后，使用 `Viewport` 对象的 `Convert2d` 方法将该点从 3D 坐标系转换为 2D 坐标系。
        最后将返回一个 `Point2d` 结构体，其中包含转换后的 x 和 y 坐标。
         * 
         * 
         在C#中，将 Point3d 结构体的三维坐标转换为当前视图平面上的二维坐标，
        可以使用 Vector3d 类提供的 GetVectorTo 方法实现。该方法的功能是计算从一个 3D 点到另一个 3D 点的向量，
        在常见的 CAD 开发环境下，该方法可以用于将 3D 点转换为当前视口的二维坐标。下面是一个示例代码：
        public Point2d Point3DTo2D(Point3d point, Viewport viewport)
        {
            Vector3d vector = viewport.ViewDirection;
            Point3d origin = viewport.ViewTarget;
            vector = vector.TransformBy(viewport.ViewportTwist);
            vector = vector.GetNormal();
            Point3d delta = point - origin;
            double distance = delta.DotProduct(vector);
            Point3d projectedPoint = origin + distance * vector;
            Point3d wcsToDcs = new Point3d(viewport.ViewTransform.Inverse().Translation);
            Point3d dcsPoint = projectedPoint.TransformBy(Matrix3d.WorldToPlane(wcsToDcs, vector));
            Point2d resultPoint = new Point2d(dcsPoint.X, dcsPoint.Y);
            return resultPoint;
        }
        在此示例代码中，Vector3d 类的 GetVectorTo 方法用于计算从观察点到点 point 的向量。
        viewport.ViewDirection 属性获取当前视口的观察方向向量，viewport.ViewTarget 属性获取当前视口的观察目标点。TransformBy 方法用于应用视口的视图旋转。与此相应的，viewport.ViewportTwist 属性是当前视口的旋转角度。GetNormal 方法返回向量的单位向量。

        在得到从观察点到点 point 的向量后，可以计算出该向量在当前视口的投影距离 distance。
        然后，可以计算出在当前视口投影上的点 projectedPoint。
        最后，应用视口的视图转换将该点从世界坐标系转换为当前视口的坐标系，
        并应用视口的平面转换将点从 3D 坐标系转换为 2D 坐标系。最后将返回一个 Point2d 结构体，
        其中包含转换后的 x 和 y 坐标。



        ==============================================================
         Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 提示用户选择第一个点
            PromptPointResult ppr1 = ed.GetPoint("\n选择第一个点：");
            if (ppr1.Status != PromptStatus.OK) return;
            Point3d pt1 = ppr1.Value;

            // 提示用户选择第二个点
            PromptPointOptions ppo = new PromptPointOptions("\n选择第二个点：");
            ppo.BasePoint = pt1;
            ppo.UseBasePoint = true;
            PromptPointResult ppr2 = ed.GetPoint(ppo);
            if (ppr2.Status != PromptStatus.OK) return;
            Point3d pt2 = ppr2.Value;

            // 计算矩形的四个角点
            Point3d pt3 = new Point3d(pt1.X, pt2.Y, 0);
            Point3d pt4 = new Point3d(pt2.X, pt1.Y, 0);

            // 开始绘制矩形
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                Polyline rect = new Polyline();
                rect.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                rect.AddVertexAt(1, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                rect.AddVertexAt(2, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                rect.AddVertexAt(3, new Point2d(pt4.X, pt4.Y), 0, 0, 0);
                rect.Closed = true;
                btr.AppendEntity(rect);
                tr.AddNewlyCreatedDBObject(rect, true);

                tr.Commit();
            }

            要对一个实体画轮廓并删除实体，可以使用以下代码：
            public void DrawEntityOutlineAndDelete(Entity ent)
            {
                // 获取当前文档和数据库
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;

                // 开启事务
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 获取实体的边界框
                    Extents3d ext = ent.GeometricExtents;

                    // 创建一个矩形
                    Polyline rect = new Polyline();
                    rect.AddVertexAt(0, new Point2d(ext.MinPoint.X, ext.MinPoint.Y), 0, 0, 0);
                    rect.AddVertexAt(1, new Point2d(ext.MaxPoint.X, ext.MinPoint.Y), 0, 0, 0);
                    rect.AddVertexAt(2, new Point2d(ext.MaxPoint.X, ext.MaxPoint.Y), 0, 0, 0);
                    rect.AddVertexAt(3, new Point2d(ext.MinPoint.X, ext.MaxPoint.Y), 0, 0, 0);
                    rect.Closed = true;

                    // 将矩形添加到模型空间
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    btr.AppendEntity(rect);
                    tr.AddNewlyCreatedDBObject(rect, true);

                    // 删除实体
                    ent.Erase();

                    // 提交事务
                    tr.Commit();
                }
            }

         */

        /* 标注类的详细的参数解释
        DimStyleTableRecord Properties：

        Properties	Description(描述)
        Dimadec	    角度标注保留的有效位数
        Dimalt	    控制是否显示换算单位标注值中的零
        Dimaltd	    控制换算单位中小数的位数
        Dimaltf	    控制换算单位中的比例因子
        Dimaltrnd	决定换算单位的舍入
        Dimalttd	设置标注换算单位公差值小数位的位数
        Dimalttz	控制是否对公差值作消零处理
        Dimaltu 	为所有标注样式族(角度标注除外)换算单位设置单位格式
        Dimaltz	    控制是否对换算单位标注值作消零处理。
        Dimapost	为所有标注类型（角度标注除外）的换算标注测量值指定文字前缀或后缀（或两者都指定)
        Dimarcsym	控制弧长标注中圆弧符号的显示
        Dimasz	    控制尺寸线、引线箭头的大小
        Dimatfit 	控制尺寸线、引线箭头的大小。并控制钩线的大小 
        Dimaunit 	设置角度标注的单位格式
        Dimazin 	对角度标注作消零处理
        Dimblk 	    设置尺寸线或引线末端显示的箭头块
        Dimblk1     当 DIMSAH 为开时，设置尺寸线第一个端点的箭头
        Dimblk2	    当 DIMSAH 为开时，设置尺寸线第二个端点的箭头
        Dimcen 	    控制由 DIMCENTER、DIMDIAMETER 和 DIMRADIUS 绘制的圆或圆弧的圆心标记和中心线
        Dimclrd 	为尺寸线、箭头和标注引线指定颜色
        Dimclre 	为尺寸界线指定颜色
        Dimclrt 	为标注文字指定颜色
        Dimdec 	    设置标注主单位显示的小数位位数
        Dimdle 	    当使用小斜线代替箭头进行标注时，设置尺寸线超出尺寸界线的距离
        Dimdli 	    控制基线标注中尺寸线的间距
        Dimdsep 	指定一个单独的字符作为创建十进制标注时使用的小数分隔符
        Dimexe 	    指定尺寸界线超出尺寸线的距离
        Dimexo 	    指定尺寸界线偏离原点的距离
        Dimfrac 	设置当 DIMLUNIT 被设为 4（建筑）或 5（分数）时的分数格式
        Dimfxlen 	设置固定延长线的值。
        DimfxlenOn 	设置是否显示固定延长线
        Dimgap 	    在尺寸线分段以放置标注文字时，设置标注文字周围的距离
        Dimjogang 	确定折弯半径标注中尺寸线的横向段角度
        Dimjust 	控制标注文字的水平位置
        Dimldrblk 	指定引线的箭头类型
        Dimlfac 	设置线性标注测量值的比例因子
        Dimlim 	    将极限尺寸生成为缺省文字
        Dimltex1 	设置第一条尺寸界线的线型
        Dimltex2 	设置第二条尺寸界线的线型
        Dimltype 	设置尺寸线的线型
        Dimlunit 	为所有标注类型（角度标注除外）设置单位
        Dimlwd 	    指定尺寸线的线宽
        Dimlwe 	    指定尺寸界线的线宽
        Dimpost 	指定标注测量值的文字前缀或后缀（或两者都指定）
        Dimrnd 	    将所有标注距离舍入到指定值
        Dimsah 	    控制尺寸线箭头块的显示
        Dimscale 	为标注变量（指定尺寸、距离或偏移量）设置全局比例因子
        Dimsd1   	控制是否禁止显示第一条尺寸线
        Dimsd2   	控制是否禁止显示第二条尺寸线
        Dimse1   	控制是否禁止显示第一条尺寸界线
        Dimse2   	控制是否禁止显示第二条尺寸界线
        Dimsoxd 	控制是否重新定义拖动的标注对象
        Dimtad 	    控制是否允许尺寸线绘制到尺寸界线之外
        Dimtdec 	显示当前标注样式
        Dimtfac 	控制文字相对尺寸线的垂直位置
        Dimtfill 	设置标注主单位的公差值显示的小数位数
        Dimtfillclr 	设置用来计算标注分数或公差文字的高度的比例因子
        Dimtih 	    控制所有标注类型（坐标标注除外）的标注文字在尺寸界线内的位置
        Dimtix 	    在尺寸界线之间绘制文字
        Dimtm 	    当 DIMTOL 或 DIMLIM 为开时，为标注文字设置最大下偏差
        Dimtmove     设置标注文字的移动规则
        Dimtofl     控制是否将尺寸线绘制在尺寸界线之间（即使文字放置在尺寸界线之外）
        Dimtoh 	    控制标注文字在尺寸界线外的位置
        Dimtol 	    将公差添加到标注文字中
        Dimtolj     设置公差值相对名词性标注文字的垂直对正方式
        Dimtp 	    当 DIMTOL 或 DIMLIM 为开时，为标注文字设置最大上偏差
        Dimtsz 	    指定线性标注、半径标注以及直径标注中替代箭头的小斜线尺寸
        Dimtvp 	    控制尺寸线上方或下方标注文字的垂直位置
       * Dimtxsty    指定标注的文字样式
       * Dimtxt 	 指定标注文字的高度，除非当前文字样式具有固定的高度
        Dimtzin     控制是否对公差值作消零处理
        Dimupt 	    控制用户定位文字的选项
        Dimzin 	    控制是否对主单位值作消零处理
        IsModifiedForRecompute 	Assesses the modified state of this dimension style. 

         Dimension各参数含义
            发表于2014 年 5 月 21 日由boitboy
            Each dimension has the capability of overriding the settings assigned to it by a dimension style. The following properties are available for most dimension objects:
        每个维度都具有覆盖由维度样式分配给它的设置的能力。以下属性可用于大多数维度对象:

            Dimatfit//指定仅在延长线内显示尺寸线，并在延长线内外强制显示尺寸文本和箭头。控制尺寸线、引线箭头的大小。并控制钩线的大小 
            Specifies the display of dimension lines inside extension lines only, and forces dimension text and arrowheads inside or outside extension lines.


            Dimaltrnd//指定替代单位的舍入
            Specifies the rounding of alternate units.设置主单位–舍入值

            Dimasz//指定尺寸线箭头、引线箭头和钩线的大小。
            Specifies the size of dimension line arrowheads, leader line arrowheads, and hook lines.

            设置标注箭头大小，引线箭头大小，和hook lines

            Dimaunit//指定角度尺寸的单位格式
            Specifies the unit format for angular dimensions.设置角度标注的单位格式

            Dimcen//指定径向和直径尺寸的中心标记的类型和尺寸。
            Specifies the type and size of center mark for radial and diametric dimensions.

            Dimclre
            设置尺寸界线的颜色.

            Dimdsep
            设置线型标注的小数分隔符.

            Dimfrac//指定分数值的格式。
            Specifies the format of fractional values in 标注和公差.

            Dimlfac//设置线型标注的全局比例
            Specifies a global scale factor for linear dimension measurements.


            Dimltex1, Dimltex2
            设置尺寸界线的线型.

            Dimlwd//指定尺寸线的线重。
            Specifies the lineweight for the dimension line.

            Dimlwe
            设置尺寸界线的线宽

            Dimjust//指定维度文本的水平对齐。
            Specifies the horizontal justification for dimension text.

            Dimrnd//指定尺寸测量的距离舍入。
            Specifies the distance rounding for dimension measurements.

            Dimsd1, Dimsd2//指定尺寸线的抑制。
            Specifies the suppression of the dimension lines.

            Dimse1, Dimse2//指定对延长线的抑制。
            Specifies the suppression of extension lines.

            Dimtfac//指定公差值的文本高度相对于尺寸文本高度的比例因子。
            Specifies a scale factor for the text height of tolerance values relative to the dimension text height.

            Dimlunit//指定除角外的所有维度的单位格式。
            Specifies the unit format for all dimensions except angular.

            Dimtm//指定尺寸文本的最小公差限制。
            Specifies the minimum tolerance limit for dimension text.

            Dimtol//指定是否与尺寸文本一起显示公差。
            Specifies if tolerances are displayed with the dimension text.

            Dimtolj//指定相对于标称尺寸文本的公差值的垂直对齐。
            Specifies the vertical justification of tolerance values relative to the nominal dimension text.

            Dimtp//指定尺寸文本的最大公差限制。
            Specifies the maximum tolerance limit for dimension text.

            Dimzin//指定在尺寸值中抑制前导零和尾随零以及零英尺和英寸测量值。
            Specifies the suppression of leading and trailing zeros, and zero foot and inch measurements in dimension values.

            Prefix//指定维度值前缀。
            Specifies the dimension value prefix.

            Suffix//指定维度值后缀。
            Specifies the dimension value suffix.

            TextPrecision
            设置角度标注文字的精度.

            TextPosition
            设置文字位置.

            TextRotation//指定维度文本的旋转角度。
            Specifies the rotation angle of the dimension text.
        Dimatfit

            按对话框内容组织:

            Dimdle    —    尺寸线   —  超出标记    | 100

            Dimclrd  —   尺寸线 —   颜色       | ByBlock

            Dimexe    —   尺寸界线 —  超出尺寸线  | 100

            Dimexo    —   尺寸界线 —  起点偏移量  | 250

            Dimdli     —  基线间距  — 0



            Dimblk — 符号和箭头 | 建筑标记

            Dimblk1 Dimblk2 — 符号和箭头 — 箭头1/箭头2 | 建筑标记

            Dimazs — 符号和箭头  —  箭头大小 | 100

            Dimtxsty — 文字 — 文字样式 | TSSD_Dimension

            Dimclrt  — 文字 — 文字颜色 | 7

            Dimtxt   — 文字 — 文字高度 | 350

            Dimgap   — 文字 — 从尺寸线偏移 | 100

            Dimtad   — 文字 — 文字位置垂直 | 1 — 上方

            Dimjust  — 文字 — 文字位置水平

            Dimtih与Dimtoh共同确定文字对齐样式

            Dimtih   — 文字 — 文字对齐 在尺寸界线内的时候–关

            Dimtoh   — 文字 — 文字对齐 在尺寸界线上的时候–关

            关 将文字与尺寸线对齐 开 水平绘制文字

            Dimatfit/Dimtix/Dimsoxd共同确定调整选项

            DIMATFIT — 调整–

            0 将文字和箭头均放置于尺寸界线之外
            1 先移动箭头，然后移动文字 
            2 先移动文字，然后移动箭头
            3 移动文字和箭头中较合适的一个 –对结构设计，应选择2

            Dimtix   — 调整 — 文字始终保持在尺寸界线之间 | 开

            Dimsoxd   — 调整 — 空间不足时，是否显示箭头 | 开

            Dimtoh   — 调整 — 尺寸线上方，不带引线

            Dimtmove — 调整 — 文字位置 | 2–尺寸线上方，不带引线

            Dimtofl  — 调整 — 优化-在尺寸界线之间绘制尺寸线 | 开

            Dimtdec  — 公差 — 公差精度 | 0

            Dimdec — 精度 — 0


         */

        /*  CAD开发错误消息
        case Acad::eOutOfMemory: lstrcpy(Glb_AcadErrorInfo, _T(“内存不足”)); break;

          case Acad::eInvalidInput: lstrcpy(Glb_AcadErrorInfo, _T(“无效的参数”)); break;

          case Acad::eInvalidInput: lstrcpy(Glb_AcadErrorInfo, _T(“无效的参数”)); break;

          case Acad::eInvalidInput: lstrcpy(Glb_AcadErrorInfo, _T(“无效的参数”)); break;

          case Acad::eInvalidInput: lstrcpy(Glb_AcadErrorInfo, _T(“无效的参数”)); break;

          case Acad::eUserBreak: lstrcpy(Glb_AcadErrorInfo, _T(“操作被取消”)); break;

          case Acad::eCommandNotFound: lstrcpy(Glb_AcadErrorInfo, _T(“无法解析命令”)); break;

          case Acad::eObjectLocked: lstrcpy(Glb_AcadErrorInfo, _T(“对象被锁定”)); break;

          case Acad::eObjectLocked: lstrcpy(Glb_AcadErrorInfo, _T(“对象被锁定”)); break;

          case Acad::eInvalidDwgVersion: lstrcpy(Glb_AcadErrorInfo, _T(“无效的操作”)); break;

          Acad::eInvalidInput: lstrcpy(Glb_AcadErrorInfo, _T(“无效输入”)); break;

          Acad::eAmbiguousInput: lstrcpy(Glb_AcadErrorInfo, _T(“输入模糊”)); break;

          Acad::eNullObjectId: lstrcpy(Glb_AcadErrorInfo, _T(“空对象 ID”)); break;-

          Acad::eBrokenObject: lstrcpy(Glb_AcadErrorInfo, _T(“破损的对象”)); break;

          Acad::eNullPointer: lstrcpy(Glb_AcadErrorInfo, _T(“空指针”)); break;

          Acad::eUndefinedView: lstrcpy(Glb_AcadErrorInfo, _T(“未定义的视图”)); break;

          Acad::eNoActiveTransactions: lstrcpy(Glb_AcadErrorInfo, _T(“没有活动事务”)); break;

          Acad::eNoDatabase: lstrcpy(Glb_AcadErrorInfo, _T(“没有数据库”)); break;

          Acad::eDatabaseNotInitialized: lstrcpy(Glb_AcadErrorInfo, _T(“数据库未初始化”)); break;

          Acad::eInvalidDWGVersion: lstrcpy(Glb_AcadErrorInfo, _T(“无效的 DWG 版本”)); break;

          Acad::eOpenWhileCommandActive: lstrcpy(Glb_AcadErrorInfo, _T(“在命令活动时打开”)); break;

          Acad::eCloseWhileCommandActive: lstrcpy(Glb_AcadErrorInfo, _T(“在命令活动时关闭”)); break;

          Acad::eMustOpenDrawing: lstrcpy(Glb_AcadErrorInfo, _T(“必须打开绘图”)); break;

          Acad::eMustOpenTemplate: lstrcpy(Glb_AcadErrorInfo, _T(“必须打开模板”)); break;

          Acad::eMustOpenSheetSet: lstrcpy(Glb_AcadErrorInfo, _T(“必须打开工作表集”)); break;

          Acad::eMustOpenDrawingOrTemplate: lstrcpy(Glb_AcadErrorInfo, _T(“必须打开绘图或模板”)); break;

          Acad::eInvalidDxfId: lstrcpy(Glb_AcadErrorInfo, _T(“无效的 DXF ID”)); break;

          Acad::eNotHandledYet: lstrcpy(Glb_AcadErrorInfo, _T(“尚未处理”)); break;

          Acad::eNotApplicable: lstrcpy(Glb_AcadErrorInfo, _T(“不适用”)); break;

          Acad::eInvalidViewportObjectId: lstrcpy(Glb_AcadErrorInfo, _T(“无效的视口对象 ID”)); break;

          Acad::eNotInPaperspace: lstrcpy(Glb_AcadErrorInfo, _T(“不在图纸空间”)); break;

          Acad::eInvalidPlotArea: lstrcpy(Glb_AcadErrorInfo, _T(“无效的绘图区域”)); break;

          Acad::ePaperSpaceViewportExists: lstrcpy(Glb_AcadErrorInfo, _T(“图纸空间视口已存在”)); break;

          Acad::eLayoutNameExists: lstrcpy(Glb_AcadErrorInfo, _T(“布局名称已存在”)); break;

          Acad::eNoLayout: lstrcpy(Glb_AcadErrorInfo, _T(“没有布局”)); break;

          case Acad::eObjectToBeDeleted: lstrcpy(Glb_AcadErrorInfo, _T(“对象将被删除”)); break;

          case Acad::eNullObjectId: lstrcpy(Glb_AcadErrorInfo, _T(“空的对象 ID”)); break;

          case Acad::eInvalidDwgVersion: lstrcpy(Glb_AcadErrorInfo, _T(“无效的 DWG 版本”)); break;

          case Acad::eDuplicateRecordName: lstrcpy(Glb_AcadErrorInfo, _T(“记录名称重复”)); break;

          case Acad::eInvalidOwnerObject: lstrcpy(Glb_AcadErrorInfo, _T(“无效的所有者对象”)); break;

          case Acad::eInvalidResBuf: lstrcpy(Glb_AcadErrorInfo, _T(“无效的 ResBuf”)); break;

          case Acad::eDuplicateKey: lstrcpy(Glb_AcadErrorInfo, _T(“键重复”)); break;

          case Acad::eBufferTooSmall: lstrcpy(Glb_AcadErrorInfo, _T(“缓冲区太小”)); break;

          case Acad::eInvalidKey: lstrcpy(Glb_AcadErrorInfo, _T(“无效的键”)); break;

          case Acad::eNoActiveTransactions: lstrcpy(Glb_AcadErrorInfo, _T(“没有活动的事务”)); break;

          case Acad::eDocumentSwitchDisabled: lstrcpy(Glb_AcadErrorInfo, _T(“文档切换已禁用”)); break;

          case Acad::eEndOfObject: lstrcpy(Glb_AcadErrorInfo, _T(“对象结束”)); break;

          case Acad::eEndOfFile: lstrcpy(Glb_AcadErrorInfo, _T(“文件结束”)); break;

          case Acad::eNotImplementedYet: lstrcpy(Glb_AcadErrorInfo, _T(“尚未实现”)); break;

          case Acad::eUnknownDwgVersion: lstrcpy(Glb_AcadErrorInfo, _T(“未知的 DWG 版本”)); break;

          case Acad::eHandleExistsAlready: lstrcpy(Glb_AcadErrorInfo, _T(“句柄已存在”)); break;

          case Acad::eNullPtr: lstrcpy(Glb_AcadErrorInfo, _T(“空指针”)); break;

          case Acad::eNoLongerApplicable: lstrcpy(Glb_AcadErrorInfo, _T(“不再适用”)); break;

          case Acad::eEntityInInactiveLayout: lstrcpy(Glb_AcadErrorInfo, _T(“非活动布局中的实体”)); break;

          case Acad::eCannotRestoreFromAcisFile: lstrcpy(Glb_AcadErrorInfo, _T(“无法从 ACIS 文件中恢复”)); break;

          case Acad::eInvalidInput: lstrcpy(Glb_AcadErrorInfo, _T(“无效的输入”)); break;

          内存不足：case Acad::eOutOfMemory: lstrcpy(Glb_AcadErrorInfo, _T(“内存不足”)); break;

          锁定错误：case Acad::eLockViolation: lstrcpy(Glb_AcadErrorInfo, _T(“锁定错误”)); break;

          无法打开文件：case Acad::eFileNotFound: lstrcpy(Glb_AcadErrorInfo, _T(“无法打开文件”)); break;

          访问被拒绝：case Acad::eAccessDenied: lstrcpy(Glb_AcadErrorInfo, _T(“访问被拒绝”)); break;

          无效的DWG文件格式：case Acad::eInvalidDwgVersion: lstrcpy(Glb_AcadErrorInfo, _T(“无效的DWG文件格式”)); break;

          无效的组码：case Acad::eInvalidGroupCode: lstrcpy(Glb_AcadErrorInfo, _T(“无效的组码”)); break;

          无效的块名称：case Acad::eBadDxfSequence: lstrcpy(Glb_AcadErrorInfo, _T(“无效的块名称”)); break;

          无效的段落类型：case Acad::eBadDxfLineNumber: lstrcpy(Glb_AcadErrorInfo, _T(“无效的段落类型”)); break;

          命令错误：case Acad::eCommandWasInProgress: lstrcpy(Glb_AcadErrorInfo, _T(“命令错误”)); break;

          对象未找到：case Acad::eNullObjectId: lstrcpy(Glb_AcadErrorInfo, _T(“对象未找到”)); break;

          错误的对象类型：case Acad::eWrongObjectType: lstrcpy(Glb_AcadErrorInfo, _T(“错误的对象类型”)); break;

          不支持的版本：case Acad::eNotImplementedYet: lstrcpy(Glb_AcadErrorInfo, _T(“不支持的版本”)); break;

          数据错误：case Acad::eDataError: lstrcpy(Glb_AcadErrorInfo, _T(“数据错误”)); break;

          错误的指针：case Acad::eInvalidIndex: lstrcpy(Glb_AcadErrorInfo, _T(“错误的指针”)); break;

          无效的ID：case Acad::eInvalidIxDxf: lstrcpy(Glb_AcadErrorInfo, _T(“无效的ID”)); break;

          错误的参数：case Acad::eBadDxfFile: lstrcpy(Glb_AcadErrorInfo, _T(“错误的参数”)); break;

          不支持的操作：case Acad::eUnsupportedDwgVersion: lstrcpy(Glb_AcadErrorInfo, _T(“不支持的操作”)); break;

          未知错误：default: lstrcpy(Glb_AcadErrorInfo, _T(“未知错误”)); break;
          ————————————————

          版权声明：本文为博主原创文章，遵循 CC 4.0 BY-SA 版权协议，转载请附上原文出处链接和本声明。

          原文链接：https://blog.csdn.net/ultramand/article/details/130418848

        */

        /*  AutoCAD.net-错误消息大全

         AutoCAD.net-错误消息大全
            case Acad::eOk:lstrcpy(Glb_AcadErrorInfo,_T("正确"));break;
            case Acad::eNotImplementedYet:lstrcpy(Glb_AcadErrorInfo,_T("尚未实现"));break;
            case Acad::eNotApplicable:lstrcpy(Glb_AcadErrorInfo,_T("不合适的"));break;
            case Acad::eInvalidInput:lstrcpy(Glb_AcadErrorInfo,_T("无效的输入"));break;
            case Acad::eAmbiguousInput:lstrcpy(Glb_AcadErrorInfo,_T("模糊不清的输入"));break;
            case Acad::eAmbiguousOutput:lstrcpy(Glb_AcadErrorInfo,_T("模糊不清的输出"));break;
            case Acad::eOutOfMemory:lstrcpy(Glb_AcadErrorInfo,_T("内存不足"));break;
            case Acad::eBufferTooSmall:lstrcpy(Glb_AcadErrorInfo,_T("缓冲区太小"));break;
            case Acad::eInvalidOpenState:lstrcpy(Glb_AcadErrorInfo,_T("无效的打开状态"));break;
            case Acad::eEntityInInactiveLayout:lstrcpy(Glb_AcadErrorInfo,_T("实体不在活动布局上"));break;
            case Acad::eHandleExists:lstrcpy(Glb_AcadErrorInfo,_T("句柄已存在"));break;
            case Acad::eNullHandle:lstrcpy(Glb_AcadErrorInfo,_T("空句柄"));break;
            case Acad::eBrokenHandle:lstrcpy(Glb_AcadErrorInfo,_T("损坏的句柄"));break;
            case Acad::eUnknownHandle:lstrcpy(Glb_AcadErrorInfo,_T("未知句柄"));break;
            case Acad::eHandleInUse:lstrcpy(Glb_AcadErrorInfo,_T("句柄被占用"));break;
            case Acad::eNullObjectPointer:lstrcpy(Glb_AcadErrorInfo,_T("对象指针为空"));break;
            case Acad::eNullObjectId:lstrcpy(Glb_AcadErrorInfo,_T("对象ID为空"));break;
            case Acad::eNullBlockName:lstrcpy(Glb_AcadErrorInfo,_T("块名称为空"));break;
            case Acad::eContainerNotEmpty:lstrcpy(Glb_AcadErrorInfo,_T("容器不为空"));break;
            case Acad::eNullEntityPointer:lstrcpy(Glb_AcadErrorInfo,_T("实体指针为空"));break;
            case Acad::eIllegalEntityType:lstrcpy(Glb_AcadErrorInfo,_T("非法的实体类型"));break;
            case Acad::eKeyNotFound:lstrcpy(Glb_AcadErrorInfo,_T("关键字未找到"));break;
            case Acad::eDuplicateKey:lstrcpy(Glb_AcadErrorInfo,_T("重复的关键字"));break;
            case Acad::eInvalidIndex:lstrcpy(Glb_AcadErrorInfo,_T("无效的索引"));break;
            case Acad::eDuplicateIndex:lstrcpy(Glb_AcadErrorInfo,_T("重复的索引"));break;
            case Acad::eAlreadyInDb:lstrcpy(Glb_AcadErrorInfo,_T("已经在数据库中了"));break;
            case Acad::eOutOfDisk:lstrcpy(Glb_AcadErrorInfo,_T("硬盘容量不足"));break;
            case Acad::eDeletedEntry:lstrcpy(Glb_AcadErrorInfo,_T("已经删除的函数入口"));break;
            case Acad::eNegativeValueNotAllowed:lstrcpy(Glb_AcadErrorInfo,_T("不允许输入负数"));break;
            case Acad::eInvalidExtents:lstrcpy(Glb_AcadErrorInfo,_T("无效的空间范围"));break;
            case Acad::eInvalidAdsName:lstrcpy(Glb_AcadErrorInfo,_T("无效的ADS名称"));break;
            case Acad::eInvalidSymbolTableName:lstrcpy(Glb_AcadErrorInfo,_T("无效的符号名称"));break;
            case Acad::eInvalidKey:lstrcpy(Glb_AcadErrorInfo,_T("无效的关键字"));break;
            case Acad::eWrongObjectType:lstrcpy(Glb_AcadErrorInfo,_T("错误的类型"));break;
            case Acad::eWrongDatabase:lstrcpy(Glb_AcadErrorInfo,_T("错误的数据库"));break;
            case Acad::eObjectToBeDeleted:lstrcpy(Glb_AcadErrorInfo,_T("对象即将被删除"));break;
            case Acad::eInvalidDwgVersion:lstrcpy(Glb_AcadErrorInfo,_T("不合理的DWG版本"));break;
            case Acad::eAnonymousEntry:lstrcpy(Glb_AcadErrorInfo,_T("多重入口"));break;
            case Acad::eIllegalReplacement:lstrcpy(Glb_AcadErrorInfo,_T("非法的替代者"));break;
            case Acad::eEndOfObject:lstrcpy(Glb_AcadErrorInfo,_T("对象结束"));break;
            case Acad::eEndOfFile:lstrcpy(Glb_AcadErrorInfo,_T("文件结束"));break;
            case Acad::eIsReading:lstrcpy(Glb_AcadErrorInfo,_T("正在读取"));break;
            case Acad::eIsWriting:lstrcpy(Glb_AcadErrorInfo,_T("正在写入"));break;
            case Acad::eNotOpenForRead:lstrcpy(Glb_AcadErrorInfo,_T("不是只读打开"));break;
            case Acad::eNotOpenForWrite:lstrcpy(Glb_AcadErrorInfo,_T("不是可写打开"));break;
            case Acad::eNotThatKindOfClass:lstrcpy(Glb_AcadErrorInfo,_T("类型不匹配"));break;
            case Acad::eInvalidBlockName:lstrcpy(Glb_AcadErrorInfo,_T("不合理的块名称"));break;
            case Acad::eMissingDxfField:lstrcpy(Glb_AcadErrorInfo,_T("DXF字段缺失"));break;
            case Acad::eDuplicateDxfField:lstrcpy(Glb_AcadErrorInfo,_T("DXF字段重复"));break;
            case Acad::eInvalidDxfCode:lstrcpy(Glb_AcadErrorInfo,_T("不合理的DXF编码"));break;
     

            case Acad::eInvalidResBuf:lstrcpy(Glb_AcadErrorInfo,_T("不合理的ResBuf"));break;
            case Acad::eBadDxfSequence:lstrcpy(Glb_AcadErrorInfo,_T("不正确的DXF顺序"));break;
            case Acad::eFilerError:lstrcpy(Glb_AcadErrorInfo,_T("文件错误"));break;
            case Acad::eVertexAfterFace:lstrcpy(Glb_AcadErrorInfo,_T("顶点在面后面"));break;
            case Acad::eInvalidFaceVertexIndex:lstrcpy(Glb_AcadErrorInfo,_T("不合理的面顶点顺序"));break;
            case Acad::eInvalidMeshVertexIndex:lstrcpy(Glb_AcadErrorInfo,_T("不合理的mesh顺序"));break;
            case Acad::eOtherObjectsBusy:lstrcpy(Glb_AcadErrorInfo,_T("其它对象忙"));break;
            case Acad::eMustFirstAddBlockToDb:lstrcpy(Glb_AcadErrorInfo,_T("必须先把块加入到数据库"));break;
            case Acad::eCannotNestBlockDefs:lstrcpy(Glb_AcadErrorInfo,_T("不可以嵌套块定义"));break;
            case Acad::eDwgRecoveredOK:lstrcpy(Glb_AcadErrorInfo,_T("修复DWG完成"));break;
            case Acad::eDwgNotRecoverable:lstrcpy(Glb_AcadErrorInfo,_T("无法修复DWG"));break;
            case Acad::eDxfPartiallyRead:lstrcpy(Glb_AcadErrorInfo,_T("DXF部分读取"));break;
            case Acad::eDxfReadAborted:lstrcpy(Glb_AcadErrorInfo,_T("读取DXF终止"));break;
            case Acad::eDxbPartiallyRead:lstrcpy(Glb_AcadErrorInfo,_T("DXB部分读取"));break;
            case Acad::eDwgCRCDoesNotMatch:lstrcpy(Glb_AcadErrorInfo,_T("DWG文件的CRC不匹配"));break;
            case Acad::eDwgSentinelDoesNotMatch:lstrcpy(Glb_AcadErrorInfo,_T("DWG文件的校验不匹配"));break;
            case Acad::eDwgObjectImproperlyRead:lstrcpy(Glb_AcadErrorInfo,_T("DWG文件错误读取"));break;
            case Acad::eNoInputFiler:lstrcpy(Glb_AcadErrorInfo,_T("没有找到输入过滤"));break;
            case Acad::eDwgNeedsAFullSave:lstrcpy(Glb_AcadErrorInfo,_T("DWG需要完全保存"));break;
            case Acad::eDxbReadAborted:lstrcpy(Glb_AcadErrorInfo,_T("DXB读取终止"));break;
            case Acad::eFileLockedByACAD:lstrcpy(Glb_AcadErrorInfo,_T("文件被ACAD锁定"));break;
            case Acad::eFileAccessErr:lstrcpy(Glb_AcadErrorInfo,_T("无法读取文件"));break;
            case Acad::eFileSystemErr:lstrcpy(Glb_AcadErrorInfo,_T("文件系统错误"));break;
            case Acad::eFileInternalErr:lstrcpy(Glb_AcadErrorInfo,_T("文件内部错误"));break;
            case Acad::eFileTooManyOpen:lstrcpy(Glb_AcadErrorInfo,_T("文件被打开太多次"));break;
            case Acad::eFileNotFound:lstrcpy(Glb_AcadErrorInfo,_T("未找到文件"));break;
            case Acad::eDwkLockFileFound:lstrcpy(Glb_AcadErrorInfo,_T("找到DWG锁定文件"));break;
            case Acad::eWasErased:lstrcpy(Glb_AcadErrorInfo,_T("对象被删除"));break;
            case Acad::ePermanentlyErased:lstrcpy(Glb_AcadErrorInfo,_T("对象被永久删除"));break;
            case Acad::eWasOpenForRead:lstrcpy(Glb_AcadErrorInfo,_T("对象只读打开"));break;
            case Acad::eWasOpenForWrite:lstrcpy(Glb_AcadErrorInfo,_T("对象可写打开"));break;
            case Acad::eWasOpenForUndo:lstrcpy(Glb_AcadErrorInfo,_T("对象撤销打开"));break;
            case Acad::eWasNotifying:lstrcpy(Glb_AcadErrorInfo,_T("对象被通知"));break;
            case Acad::eWasOpenForNotify:lstrcpy(Glb_AcadErrorInfo,_T("对象通知打开"));break;
            case Acad::eOnLockedLayer:lstrcpy(Glb_AcadErrorInfo,_T("对象在锁定图层上"));break;
            case Acad::eMustOpenThruOwner:lstrcpy(Glb_AcadErrorInfo,_T("必须经过所有者打开"));break;
            case Acad::eSubentitiesStillOpen:lstrcpy(Glb_AcadErrorInfo,_T("子对象依然打开着"));break;
            case Acad::eAtMaxReaders:lstrcpy(Glb_AcadErrorInfo,_T("超过最大打开次数"));break;
            case Acad::eIsWriteProtected:lstrcpy(Glb_AcadErrorInfo,_T("对象被写保护"));break;
            case Acad::eIsXRefObject:lstrcpy(Glb_AcadErrorInfo,_T("对象是XRef"));break;
            case Acad::eNotAnEntity:lstrcpy(Glb_AcadErrorInfo,_T("对象不是实体"));break;
            case Acad::eHadMultipleReaders:lstrcpy(Glb_AcadErrorInfo,_T("被多重打开"));break;
            case Acad::eDuplicateRecordName:lstrcpy(Glb_AcadErrorInfo,_T("重复的记录名称"));break;
            case Acad::eXRefDependent:lstrcpy(Glb_AcadErrorInfo,_T("依赖于XREF"));break;
            case Acad::eSelfReference:lstrcpy(Glb_AcadErrorInfo,_T("引用自身"));break;
            case Acad::eMissingSymbolTable:lstrcpy(Glb_AcadErrorInfo,_T("丢失符号化表"));break;
            case Acad::eMissingSymbolTableRec:lstrcpy(Glb_AcadErrorInfo,_T("丢失符号化记录"));break;
            case Acad::eWasNotOpenForWrite:lstrcpy(Glb_AcadErrorInfo,_T("不是可写打开"));break;
            case Acad::eCloseWasNotifying:lstrcpy(Glb_AcadErrorInfo,_T("对象关闭,正在执行通知"));break;
            case Acad::eCloseModifyAborted:lstrcpy(Glb_AcadErrorInfo,_T("对象关闭,修改被取消"));break;
     
            case Acad::eClosePartialFailure:lstrcpy(Glb_AcadErrorInfo,_T("对象关闭,部分操作未成功"));break;
            case Acad::eCloseFailObjectDamaged:lstrcpy(Glb_AcadErrorInfo,_T("对象被损坏,关闭失败"));break;
            case Acad::eCannotBeErasedByCaller:lstrcpy(Glb_AcadErrorInfo,_T("对象不可以被当前呼叫者删除"));break;
            case Acad::eCannotBeResurrected:lstrcpy(Glb_AcadErrorInfo,_T("不可以复活"));break;
            case Acad::eWasNotErased:lstrcpy(Glb_AcadErrorInfo,_T("对象未删除"));break;
            case Acad::eInsertAfter:lstrcpy(Glb_AcadErrorInfo,_T("在后面插入"));break;
            case Acad::eFixedAllErrors:lstrcpy(Glb_AcadErrorInfo,_T("修复了所有错误"));break;
            case Acad::eLeftErrorsUnfixed:lstrcpy(Glb_AcadErrorInfo,_T("剩下一些错误未修复"));break;
            case Acad::eUnrecoverableErrors:lstrcpy(Glb_AcadErrorInfo,_T("不可恢复的错误"));break;
            case Acad::eNoDatabase:lstrcpy(Glb_AcadErrorInfo,_T("没有数据库"));break;
            case Acad::eXdataSizeExceeded:lstrcpy(Glb_AcadErrorInfo,_T("扩展数据长度太大"));break;
            case Acad::eRegappIdNotFound:lstrcpy(Glb_AcadErrorInfo,_T("没有找到扩展数据注册ID"));break;
            case Acad::eRepeatEntity:lstrcpy(Glb_AcadErrorInfo,_T("重复实体"));break;
            case Acad::eRecordNotInTable:lstrcpy(Glb_AcadErrorInfo,_T("表中未找到记录"));break;
            case Acad::eIteratorDone:lstrcpy(Glb_AcadErrorInfo,_T("迭代器完成"));break;
            case Acad::eNullIterator:lstrcpy(Glb_AcadErrorInfo,_T("空的迭代器"));break;
            case Acad::eNotInBlock:lstrcpy(Glb_AcadErrorInfo,_T("不在块中"));break;
            case Acad::eOwnerNotInDatabase:lstrcpy(Glb_AcadErrorInfo,_T("所有者不在数据库中"));break;
            case Acad::eOwnerNotOpenForRead:lstrcpy(Glb_AcadErrorInfo,_T("所有者不是只读打开"));break;
            case Acad::eOwnerNotOpenForWrite:lstrcpy(Glb_AcadErrorInfo,_T("所有者不是可写打开"));break;
            case Acad::eExplodeBeforeTransform:lstrcpy(Glb_AcadErrorInfo,_T("在变换之前就被炸开了"));break;
            case Acad::eCannotScaleNonUniformly:lstrcpy(Glb_AcadErrorInfo,_T("不可以不同比例缩放"));break;
            case Acad::eNotInDatabase:lstrcpy(Glb_AcadErrorInfo,_T("不在数据库中"));break;
            case Acad::eNotCurrentDatabase:lstrcpy(Glb_AcadErrorInfo,_T("不是当前数据库"));break;
            case Acad::eIsAnEntity:lstrcpy(Glb_AcadErrorInfo,_T("是一个实体"));break;
            case Acad::eCannotChangeActiveViewport:lstrcpy(Glb_AcadErrorInfo,_T("不可以改变活动视口"));break;
            case Acad::eNotInPaperspace:lstrcpy(Glb_AcadErrorInfo,_T("不在图纸空间中"));break;
            case Acad::eCommandWasInProgress:lstrcpy(Glb_AcadErrorInfo,_T("正在执行命令"));break;
            case Acad::eGeneralModelingFailure:lstrcpy(Glb_AcadErrorInfo,_T("创建模型失败"));break;
            case Acad::eOutOfRange:lstrcpy(Glb_AcadErrorInfo,_T("超出范围"));break;
            case Acad::eNonCoplanarGeometry:lstrcpy(Glb_AcadErrorInfo,_T("没有平面几何对象"));break;
            case Acad::eDegenerateGeometry:lstrcpy(Glb_AcadErrorInfo,_T("退化的几何对象"));break;
            case Acad::eInvalidAxis:lstrcpy(Glb_AcadErrorInfo,_T("无效的轴线"));break;
            case Acad::ePointNotOnEntity:lstrcpy(Glb_AcadErrorInfo,_T("点不在实体上"));break;
            case Acad::eSingularPoint:lstrcpy(Glb_AcadErrorInfo,_T("单一的点"));break;
            case Acad::eInvalidOffset:lstrcpy(Glb_AcadErrorInfo,_T("无效的偏移"));break;
            case Acad::eNonPlanarEntity:lstrcpy(Glb_AcadErrorInfo,_T("没有平面的实体"));break;
            case Acad::eCannotExplodeEntity:lstrcpy(Glb_AcadErrorInfo,_T("不可分解的实体"));break;
            case Acad::eStringTooLong:lstrcpy(Glb_AcadErrorInfo,_T("字符串太短"));break;
            case Acad::eInvalidSymTableFlag:lstrcpy(Glb_AcadErrorInfo,_T("无效的符号化表标志"));break;
            case Acad::eUndefinedLineType:lstrcpy(Glb_AcadErrorInfo,_T("没有定义的线型"));break;
            case Acad::eInvalidTextStyle:lstrcpy(Glb_AcadErrorInfo,_T("无效的字体样式"));break;
            case Acad::eTooFewLineTypeElements:lstrcpy(Glb_AcadErrorInfo,_T("太少的线型要素"));break;
            case Acad::eTooManyLineTypeElements:lstrcpy(Glb_AcadErrorInfo,_T("太多的线型要素"));break;
            case Acad::eExcessiveItemCount:lstrcpy(Glb_AcadErrorInfo,_T("过多的项目"));break;
            case Acad::eIgnoredLinetypeRedef:lstrcpy(Glb_AcadErrorInfo,_T("忽略线型定义描述"));break;
            case Acad::eBadUCS:lstrcpy(Glb_AcadErrorInfo,_T("不好的用户坐标系"));break;
            case Acad::eBadPaperspaceView:lstrcpy(Glb_AcadErrorInfo,_T("不好的图纸空间视图"));break;
            case Acad::eSomeInputDataLeftUnread:lstrcpy(Glb_AcadErrorInfo,_T("一些输入数据未被读取"));break;
            case Acad::eNoInternalSpace:lstrcpy(Glb_AcadErrorInfo,_T("不是内部空间"));break;
     

            case Acad::eInvalidDimStyle:lstrcpy(Glb_AcadErrorInfo,_T("无效的标注样式"));break;
            case Acad::eInvalidLayer:lstrcpy(Glb_AcadErrorInfo,_T("无效的图层"));break;
            case Acad::eUserBreak:lstrcpy(Glb_AcadErrorInfo,_T("用户打断"));break;
            case Acad::eDwgNeedsRecovery:lstrcpy(Glb_AcadErrorInfo,_T("DWG文件需要修复"));break;
            case Acad::eDeleteEntity:lstrcpy(Glb_AcadErrorInfo,_T("删除实体"));break;
            case Acad::eInvalidFix:lstrcpy(Glb_AcadErrorInfo,_T("无效的方位"));break;
            case Acad::eFSMError:lstrcpy(Glb_AcadErrorInfo,_T("FSM错误"));break;
            case Acad::eBadLayerName:lstrcpy(Glb_AcadErrorInfo,_T("不好的图层名称"));break;
            case Acad::eLayerGroupCodeMissing:lstrcpy(Glb_AcadErrorInfo,_T("图层分组编码丢失"));break;
            case Acad::eBadColorIndex:lstrcpy(Glb_AcadErrorInfo,_T("不好的颜色索引号"));break;
            case Acad::eBadLinetypeName:lstrcpy(Glb_AcadErrorInfo,_T("不好的线型名称"));break;
            case Acad::eBadLinetypeScale:lstrcpy(Glb_AcadErrorInfo,_T("不好的线型缩放比例"));break;
            case Acad::eBadVisibilityValue:lstrcpy(Glb_AcadErrorInfo,_T("不好的可见性值"));break;
            case Acad::eProperClassSeparatorExpected:lstrcpy(Glb_AcadErrorInfo,_T("本身类未找到预期的分割符号(?)"));break;
            case Acad::eBadLineWeightValue:lstrcpy(Glb_AcadErrorInfo,_T("不好的线宽值"));break;
            case Acad::eBadColor:lstrcpy(Glb_AcadErrorInfo,_T("不好的颜色"));break;
            case Acad::ePagerError:lstrcpy(Glb_AcadErrorInfo,_T("页面错误"));break;
            case Acad::eOutOfPagerMemory:lstrcpy(Glb_AcadErrorInfo,_T("页面内存不足"));break;
            case Acad::ePagerWriteError:lstrcpy(Glb_AcadErrorInfo,_T("页面不可写"));break;
            case Acad::eWasNotForwarding:lstrcpy(Glb_AcadErrorInfo,_T("不是促进(?)"));break;
            case Acad::eInvalidIdMap:lstrcpy(Glb_AcadErrorInfo,_T("无效的ID字典"));break;
            case Acad::eInvalidOwnerObject:lstrcpy(Glb_AcadErrorInfo,_T("无效的所有者"));break;
            case Acad::eOwnerNotSet:lstrcpy(Glb_AcadErrorInfo,_T("未设置所有者"));break;
            case Acad::eWrongSubentityType:lstrcpy(Glb_AcadErrorInfo,_T("错误的子对象类型"));break;
            case Acad::eTooManyVertices:lstrcpy(Glb_AcadErrorInfo,_T("太多节点"));break;
            case Acad::eTooFewVertices:lstrcpy(Glb_AcadErrorInfo,_T("太少节点"));break;
            case Acad::eNoActiveTransactions:lstrcpy(Glb_AcadErrorInfo,_T("不活动的事务"));break;
            case Acad::eNotTopTransaction:lstrcpy(Glb_AcadErrorInfo,_T("不是最顶层的事务"));break;
            case Acad::eTransactionOpenWhileCommandEnded:lstrcpy(Glb_AcadErrorInfo,_T("在命令结束的时候打开(/开始)事务"));break;
            case Acad::eInProcessOfCommitting:lstrcpy(Glb_AcadErrorInfo,_T("在提交事务的过程中"));break;
            case Acad::eNotNewlyCreated:lstrcpy(Glb_AcadErrorInfo,_T("不是新创建的"));break;
            case Acad::eLongTransReferenceError:lstrcpy(Glb_AcadErrorInfo,_T("长事务引用错误"));break;
            case Acad::eNoWorkSet:lstrcpy(Glb_AcadErrorInfo,_T("没有工作集"));break;
            case Acad::eAlreadyInGroup:lstrcpy(Glb_AcadErrorInfo,_T("已经在组中了"));break;
            case Acad::eNotInGroup:lstrcpy(Glb_AcadErrorInfo,_T("不在组中"));break;
            case Acad::eInvalidREFIID:lstrcpy(Glb_AcadErrorInfo,_T("无效的REFIID"));break;
            case Acad::eInvalidNormal:lstrcpy(Glb_AcadErrorInfo,_T("无效的标准"));break;
            case Acad::eInvalidStyle:lstrcpy(Glb_AcadErrorInfo,_T("无效的样式"));break;
            case Acad::eCannotRestoreFromAcisFile:lstrcpy(Glb_AcadErrorInfo,_T("不可以从Acis(?)文件中恢复"));break;
            case Acad::eMakeMeProxy:lstrcpy(Glb_AcadErrorInfo,_T("自我代理"));break;
            case Acad::eNLSFileNotAvailable:lstrcpy(Glb_AcadErrorInfo,_T("无效的NLS文件"));break;
            case Acad::eNotAllowedForThisProxy:lstrcpy(Glb_AcadErrorInfo,_T("不允许这个代理"));break;
            case Acad::eNotSupportedInDwgApi:lstrcpy(Glb_AcadErrorInfo,_T("在Dwg Api中不支持"));break;
            case Acad::ePolyWidthLost:lstrcpy(Glb_AcadErrorInfo,_T("多段线宽度丢失"));break;
            case Acad::eNullExtents:lstrcpy(Glb_AcadErrorInfo,_T("空的空间范围"));break;
            case Acad::eExplodeAgain:lstrcpy(Glb_AcadErrorInfo,_T("再一次分解"));break;
            case Acad::eBadDwgHeader:lstrcpy(Glb_AcadErrorInfo,_T("坏的DWG文件头"));break;
            case Acad::eLockViolation:lstrcpy(Glb_AcadErrorInfo,_T("锁定妨碍当前操作"));break;
            case Acad::eLockConflict:lstrcpy(Glb_AcadErrorInfo,_T("锁定冲突"));break;
            case Acad::eDatabaseObjectsOpen:lstrcpy(Glb_AcadErrorInfo,_T("数据库对象打开"));break;
     
            case Acad::eLockChangeInProgress:lstrcpy(Glb_AcadErrorInfo,_T("锁定改变中"));break;
            case Acad::eVetoed:lstrcpy(Glb_AcadErrorInfo,_T("禁止"));break;
            case Acad::eNoDocument:lstrcpy(Glb_AcadErrorInfo,_T("没有文档"));break;
            case Acad::eNotFromThisDocument:lstrcpy(Glb_AcadErrorInfo,_T("不是从这个文档"));break;
            case Acad::eLISPActive:lstrcpy(Glb_AcadErrorInfo,_T("LISP活动"));break;
            case Acad::eTargetDocNotQuiescent:lstrcpy(Glb_AcadErrorInfo,_T("目标文档活动中"));break;
            case Acad::eDocumentSwitchDisabled:lstrcpy(Glb_AcadErrorInfo,_T("禁止文档转换"));break;
            case Acad::eInvalidContext:lstrcpy(Glb_AcadErrorInfo,_T("无效的上下文环境"));break;
            case Acad::eCreateFailed:lstrcpy(Glb_AcadErrorInfo,_T("创建失败"));break;
            case Acad::eCreateInvalidName:lstrcpy(Glb_AcadErrorInfo,_T("创建无效名称"));break;
            case Acad::eSetFailed:lstrcpy(Glb_AcadErrorInfo,_T("设置失败"));break;
            case Acad::eDelDoesNotExist:lstrcpy(Glb_AcadErrorInfo,_T("删除对象不存在"));break;
            case Acad::eDelIsModelSpace:lstrcpy(Glb_AcadErrorInfo,_T("删除模型空间"));break;
            case Acad::eDelLastLayout:lstrcpy(Glb_AcadErrorInfo,_T("删除最后一个布局"));break;
            case Acad::eDelUnableToSetCurrent:lstrcpy(Glb_AcadErrorInfo,_T("删除后无法设置当前对象"));break;
            case Acad::eDelUnableToFind:lstrcpy(Glb_AcadErrorInfo,_T("没有找到删除对象"));break;
            case Acad::eRenameDoesNotExist:lstrcpy(Glb_AcadErrorInfo,_T("重命名对象不存在"));break;
            case Acad::eRenameIsModelSpace:lstrcpy(Glb_AcadErrorInfo,_T("不可以重命令模型空间"));break;
            case Acad::eRenameInvalidLayoutName:lstrcpy(Glb_AcadErrorInfo,_T("重命名无效的布局名称"));break;
            case Acad::eRenameLayoutAlreadyExists:lstrcpy(Glb_AcadErrorInfo,_T("重命名布局名称已存在"));break;
            case Acad::eRenameInvalidName:lstrcpy(Glb_AcadErrorInfo,_T("重命名无效名称"));break;
            case Acad::eCopyDoesNotExist:lstrcpy(Glb_AcadErrorInfo,_T("拷贝不存在"));break;
            case Acad::eCopyIsModelSpace:lstrcpy(Glb_AcadErrorInfo,_T("拷贝是模型空间"));break;
            case Acad::eCopyFailed:lstrcpy(Glb_AcadErrorInfo,_T("拷贝失败"));break;
            case Acad::eCopyInvalidName:lstrcpy(Glb_AcadErrorInfo,_T("拷贝无效名称"));break;
            case Acad::eCopyNameExists:lstrcpy(Glb_AcadErrorInfo,_T("拷贝名称存在"));break;
            case Acad::eProfileDoesNotExist:lstrcpy(Glb_AcadErrorInfo,_T("配置名称不存在"));break;
            case Acad::eInvalidFileExtension:lstrcpy(Glb_AcadErrorInfo,_T("无效的文件后缀名成"));break;
            case Acad::eInvalidProfileName:lstrcpy(Glb_AcadErrorInfo,_T("无效的配置文件名称"));break;
            case Acad::eFileExists:lstrcpy(Glb_AcadErrorInfo,_T("文件存在"));break;
            case Acad::eProfileIsInUse:lstrcpy(Glb_AcadErrorInfo,_T("配置文件存在"));break;
            case Acad::eCantOpenFile:lstrcpy(Glb_AcadErrorInfo,_T("打开文件失败"));break;
            case Acad::eNoFileName:lstrcpy(Glb_AcadErrorInfo,_T("没有文件名称"));break;
            case Acad::eRegistryAccessError:lstrcpy(Glb_AcadErrorInfo,_T("读取注册表错误"));break;
            case Acad::eRegistryCreateError:lstrcpy(Glb_AcadErrorInfo,_T("创建注册表项错误"));break;
            case Acad::eBadDxfFile:lstrcpy(Glb_AcadErrorInfo,_T("坏的DXF文件"));break;
            case Acad::eUnknownDxfFileFormat:lstrcpy(Glb_AcadErrorInfo,_T("未知的DXF文件格式"));break;
            case Acad::eMissingDxfSection:lstrcpy(Glb_AcadErrorInfo,_T("丢失DXF分段"));break;
            case Acad::eInvalidDxfSectionName:lstrcpy(Glb_AcadErrorInfo,_T("无效的DXF分段名称"));break;
            case Acad::eNotDxfHeaderGroupCode:lstrcpy(Glb_AcadErrorInfo,_T("无效的DXF组码"));break;
            case Acad::eUndefinedDxfGroupCode:lstrcpy(Glb_AcadErrorInfo,_T("没有定义DXF组码"));break;
            case Acad::eNotInitializedYet:lstrcpy(Glb_AcadErrorInfo,_T("没有初始化"));break;
            case Acad::eInvalidDxf2dPoint:lstrcpy(Glb_AcadErrorInfo,_T("无效的DXF二维点"));break;
            case Acad::eInvalidDxf3dPoint:lstrcpy(Glb_AcadErrorInfo,_T("无效的DXD三维点"));break;
            case Acad::eBadlyNestedAppData:lstrcpy(Glb_AcadErrorInfo,_T("坏的嵌套应用程序数据"));break;
            case Acad::eIncompleteBlockDefinition:lstrcpy(Glb_AcadErrorInfo,_T("不完整的块定义"));break;
            case Acad::eIncompleteComplexObject:lstrcpy(Glb_AcadErrorInfo,_T("不完整的合成(?复杂)对象"));break;
            case Acad::eBlockDefInEntitySection:lstrcpy(Glb_AcadErrorInfo,_T("块定义在实体段中"));break;
            case Acad::eNoBlockBegin:lstrcpy(Glb_AcadErrorInfo,_T("没有块开始"));break;
            case Acad::eDuplicateLayerName:lstrcpy(Glb_AcadErrorInfo,_T("重复的图层名称"));break;
     

            case Acad::eBadPlotStyleName:lstrcpy(Glb_AcadErrorInfo,_T("不好的打印样式名称"));break;
            case Acad::eDuplicateBlockName:lstrcpy(Glb_AcadErrorInfo,_T("重复的块名称"));break;
            case Acad::eBadPlotStyleType:lstrcpy(Glb_AcadErrorInfo,_T("不好的打印样式类型"));break;
            case Acad::eBadPlotStyleNameHandle:lstrcpy(Glb_AcadErrorInfo,_T("不好的打印样式名称句柄"));break;
            case Acad::eUndefineShapeName:lstrcpy(Glb_AcadErrorInfo,_T("没有定义形状名称"));break;
            case Acad::eDuplicateBlockDefinition:lstrcpy(Glb_AcadErrorInfo,_T("重复的块定义"));break;
            case Acad::eMissingBlockName:lstrcpy(Glb_AcadErrorInfo,_T("丢失了块名称"));break;
            case Acad::eBinaryDataSizeExceeded:lstrcpy(Glb_AcadErrorInfo,_T("二进制数据长度太长"));break;
            case Acad::eObjectIsReferenced:lstrcpy(Glb_AcadErrorInfo,_T("对象被引用"));break;
            case Acad::eNoThumbnailBitmap:lstrcpy(Glb_AcadErrorInfo,_T("没有缩略图"));break;
            case Acad::eGuidNoAddress:lstrcpy(Glb_AcadErrorInfo,_T("未找到GUID地址"));break;
            case Acad::eMustBe0to2:lstrcpy(Glb_AcadErrorInfo,_T("必须是0到2"));break;
            case Acad::eMustBe0to3:lstrcpy(Glb_AcadErrorInfo,_T("必须是0到3"));break;
            case Acad::eMustBe0to4:lstrcpy(Glb_AcadErrorInfo,_T("必须是0到4"));break;
            case Acad::eMustBe0to5:lstrcpy(Glb_AcadErrorInfo,_T("必须是0到5"));break;
            case Acad::eMustBe0to8:lstrcpy(Glb_AcadErrorInfo,_T("必须是0到8"));break;
            case Acad::eMustBe1to8:lstrcpy(Glb_AcadErrorInfo,_T("必须是1到8"));break;
            case Acad::eMustBe1to15:lstrcpy(Glb_AcadErrorInfo,_T("必须是1到15"));break;
            case Acad::eMustBePositive:lstrcpy(Glb_AcadErrorInfo,_T("必须为正数"));break;
            case Acad::eMustBeNonNegative:lstrcpy(Glb_AcadErrorInfo,_T("必须为非负数"));break;
            case Acad::eMustBeNonZero:lstrcpy(Glb_AcadErrorInfo,_T("不可以等于0"));break;
            case Acad::eMustBe1to6:lstrcpy(Glb_AcadErrorInfo,_T("必须是1到6"));break;
            case Acad::eNoPlotStyleTranslationTable:lstrcpy(Glb_AcadErrorInfo,_T("没有打印样式事务表(?)"));break;
            case Acad::ePlotStyleInColorDependentMode:lstrcpy(Glb_AcadErrorInfo,_T("打印样式依赖颜色"));break;
            case Acad::eMaxLayouts:lstrcpy(Glb_AcadErrorInfo,_T("最大布局数量"));break;
            case Acad::eNoClassId:lstrcpy(Glb_AcadErrorInfo,_T("没有类ID"));break;
            case Acad::eUndoOperationNotAvailable:lstrcpy(Glb_AcadErrorInfo,_T("撤销操作无效"));break;
            case Acad::eUndoNoGroupBegin:lstrcpy(Glb_AcadErrorInfo,_T("撤销操作没有组开始"));break;
            case Acad::eHatchTooDense:lstrcpy(Glb_AcadErrorInfo,_T("填充太密集"));break;
            case Acad::eOpenFileCancelled:lstrcpy(Glb_AcadErrorInfo,_T("打开文件取消"));break;
            case Acad::eNotHandled:lstrcpy(Glb_AcadErrorInfo,_T("没有处理"));break;
            case Acad::eMakeMeProxyAndResurrect:lstrcpy(Glb_AcadErrorInfo,_T("将自己变成代理然后复活"));break;
            case Acad::eFileMissingSections:lstrcpy(Glb_AcadErrorInfo,_T("文件丢失分段"));break;
            case Acad::eRepeatedDwgRead:lstrcpy(Glb_AcadErrorInfo,_T("重复的读取DWG文件"));break;
            case Acad::eWrongCellType:lstrcpy(Glb_AcadErrorInfo,_T("错误的单元格类型"));break;
            case Acad::eCannotChangeColumnType:lstrcpy(Glb_AcadErrorInfo,_T("不可以改变列类型"));break;
            case Acad::eRowsMustMatchColumns:lstrcpy(Glb_AcadErrorInfo,_T("行必须匹配列"));break;
            case Acad::eFileSharingViolation:lstrcpy(Glb_AcadErrorInfo,_T("文件共享妨碍"));break;
            case Acad::eUnsupportedFileFormat:lstrcpy(Glb_AcadErrorInfo,_T("不支持的文件格式"));break;
            case Acad::eObsoleteFileFormat:lstrcpy(Glb_AcadErrorInfo,_T("废弃的文件格式"));break;
            case Acad::eDwgShareDemandLoad:lstrcpy(Glb_AcadErrorInfo,_T("DWG共享要求加载(?)"));break;
            case Acad::eDwgShareReadAccess:lstrcpy(Glb_AcadErrorInfo,_T("DWG共享读取"));break;
            case Acad::eDwgShareWriteAccess:lstrcpy(Glb_AcadErrorInfo,_T("DWG共享写入"));break;
            case Acad::eLoadFailed:lstrcpy(Glb_AcadErrorInfo,_T("加载失败"));break;
            case Acad::eDeviceNotFound:lstrcpy(Glb_AcadErrorInfo,_T("驱动未找到"));break;
            case Acad::eNoCurrentConfig:lstrcpy(Glb_AcadErrorInfo,_T("没有当前配置"));break;
            case Acad::eNullPtr:lstrcpy(Glb_AcadErrorInfo,_T("空指针"));break;
            case Acad::eNoLayout:lstrcpy(Glb_AcadErrorInfo,_T("没有布局"));break;
            case Acad::eIncompatiblePlotSettings:lstrcpy(Glb_AcadErrorInfo,_T("不兼容的打印设置"));break;
            case Acad::eNonePlotDevice:lstrcpy(Glb_AcadErrorInfo,_T("没有打印驱动"));break;
     

            case Acad::eNoMatchingMedia:lstrcpy(Glb_AcadErrorInfo,_T("没有匹配的打印尺寸"));break;
            case Acad::eInvalidView:lstrcpy(Glb_AcadErrorInfo,_T("无效的视图"));break;
            case Acad::eInvalidWindowArea:lstrcpy(Glb_AcadErrorInfo,_T("无效的窗口范围"));break;
            case Acad::eInvalidPlotArea:lstrcpy(Glb_AcadErrorInfo,_T("无效的打印范围"));break;
            case Acad::eCustomSizeNotPossible:lstrcpy(Glb_AcadErrorInfo,_T("用户输入的打印尺寸不可能存在"));break;
            case Acad::ePageCancelled:lstrcpy(Glb_AcadErrorInfo,_T("纸张取消"));break;
            case Acad::ePlotCancelled:lstrcpy(Glb_AcadErrorInfo,_T("打印取消"));break;
            case Acad::eInvalidEngineState:lstrcpy(Glb_AcadErrorInfo,_T("无效的引擎状态"));break;
            case Acad::ePlotAlreadyStarted:lstrcpy(Glb_AcadErrorInfo,_T("已经开始在打印了"));break;
            case Acad::eNoErrorHandler:lstrcpy(Glb_AcadErrorInfo,_T("没有错误处理"));break;
            case Acad::eInvalidPlotInfo:lstrcpy(Glb_AcadErrorInfo,_T("无效的打印信息"));break;
            case Acad::eNumberOfCopiesNotSupported:lstrcpy(Glb_AcadErrorInfo,_T("不支持打印份数"));break;
            case Acad::eLayoutNotCurrent:lstrcpy(Glb_AcadErrorInfo,_T("不是当前布局"));break;
            case Acad::eGraphicsNotGenerated:lstrcpy(Glb_AcadErrorInfo,_T("绘图对象创建失败(?)"));break;
            case Acad::eCannotPlotToFile:lstrcpy(Glb_AcadErrorInfo,_T("不可以打印到文件"));break;
            case Acad::eMustPlotToFile:lstrcpy(Glb_AcadErrorInfo,_T("必须打印到文件"));break;
            case Acad::eNotMultiPageCapable:lstrcpy(Glb_AcadErrorInfo,_T("不支持多种纸张"));break;
            case Acad::eBackgroundPlotInProgress:lstrcpy(Glb_AcadErrorInfo,_T("正在后台打印"));break;
            case Acad::eSubSelectionSetEmpty:lstrcpy(Glb_AcadErrorInfo,_T("子选择集被设置为空"));break;
            case Acad::eInvalidObjectId:lstrcpy(Glb_AcadErrorInfo,_T("无效的对象ID或者对象ID不在当前数据库"));break;
            case Acad::eInvalidXrefObjectId:lstrcpy(Glb_AcadErrorInfo,_T("无效的XREF对象ID或者XREF对象ID不在当前数据库"));break;
            case Acad::eNoViewAssociation:lstrcpy(Glb_AcadErrorInfo,_T("未找到对应的视图对象"));break;
            case Acad::eNoLabelBlock:lstrcpy(Glb_AcadErrorInfo,_T("视口未找到关联的块"));break;
            case Acad::eUnableToSetViewAssociation:lstrcpy(Glb_AcadErrorInfo,_T("设置视图关联视口失败"));break;
            case Acad::eUnableToGetViewAssociation:lstrcpy(Glb_AcadErrorInfo,_T("无法找到关联的视图"));break;
            case Acad::eUnableToSetLabelBlock:lstrcpy(Glb_AcadErrorInfo,_T("无法设置关联的块"));break;
            case Acad::eUnableToGetLabelBlock:lstrcpy(Glb_AcadErrorInfo,_T("无法获取关联的块"));break;
            case Acad::eUnableToRemoveAssociation:lstrcpy(Glb_AcadErrorInfo,_T("无法移除视口关联对象"));break;
            case Acad::eUnableToSyncModelView:lstrcpy(Glb_AcadErrorInfo,_T("无法同步视口和模型空间视图"));break;
            case Acad::eSecInitializationFailure:lstrcpy(Glb_AcadErrorInfo,_T("SEC(?)初始化错误"));break;
            case Acad::eSecErrorReadingFile:lstrcpy(Glb_AcadErrorInfo,_T("SEC(?)读取文件错误"));break;
            case Acad::eSecErrorWritingFile:lstrcpy(Glb_AcadErrorInfo,_T("SEC(?)写入文件错误"));break;
            case Acad::eSecInvalidDigitalID:lstrcpy(Glb_AcadErrorInfo,_T("SEC(?)无效的数字ID"));break;
            case Acad::eSecErrorGeneratingTimestamp:lstrcpy(Glb_AcadErrorInfo,_T("SEC(?)创建时间戳错误"));break;
            case Acad::eSecErrorComputingSignature:lstrcpy(Glb_AcadErrorInfo,_T("SEC(?)电子签名错误"));break;
            case Acad::eSecErrorWritingSignature:lstrcpy(Glb_AcadErrorInfo,_T("SEC(?)写入签名错误"));break;
            case Acad::eSecErrorEncryptingData:lstrcpy(Glb_AcadErrorInfo,_T("SEC(?)加密数据错误"));break;
            case Acad::eSecErrorCipherNotSupported:lstrcpy(Glb_AcadErrorInfo,_T("SEC(?)不支持的密码"));break;
            case Acad::eSecErrorDecryptingData:lstrcpy(Glb_AcadErrorInfo,_T("SEC(?)解密数据错误"));break;
            //case Acad::eInetBase:lstrcpy(Glb_AcadErrorInfo,_T("网络错误"));break;
            case Acad::eInetOk:lstrcpy(Glb_AcadErrorInfo,_T("网络正常"));break;
            case Acad::eInetInCache:lstrcpy(Glb_AcadErrorInfo,_T("在缓冲区中"));break;
            case Acad::eInetFileNotFound:lstrcpy(Glb_AcadErrorInfo,_T("网络文件不存在"));break;
            case Acad::eInetBadPath:lstrcpy(Glb_AcadErrorInfo,_T("不好的网络路径"));break;
            case Acad::eInetTooManyOpenFiles:lstrcpy(Glb_AcadErrorInfo,_T("打开太多网络文件"));break;
            case Acad::eInetFileAccessDenied:lstrcpy(Glb_AcadErrorInfo,_T("打开网络文件被拒绝"));break;
            case Acad::eInetInvalidFileHandle:lstrcpy(Glb_AcadErrorInfo,_T("无效的网络文件句柄"));break;
            case Acad::eInetDirectoryFull:lstrcpy(Glb_AcadErrorInfo,_T("网络文件夹目录已满"));break;
            case Acad::eInetHardwareError:lstrcpy(Glb_AcadErrorInfo,_T("网络硬件错误"));break;
            case Acad::eInetSharingViolation:lstrcpy(Glb_AcadErrorInfo,_T("违反网络共享"));break;
     

            case Acad::eInetDiskFull:lstrcpy(Glb_AcadErrorInfo,_T("网络硬盘满了"));break;
            case Acad::eInetFileGenericError:lstrcpy(Glb_AcadErrorInfo,_T("网络文件创建错误"));break;
            case Acad::eInetValidURL:lstrcpy(Glb_AcadErrorInfo,_T("无效的URL地址"));break;
            case Acad::eInetNotAnURL:lstrcpy(Glb_AcadErrorInfo,_T("不是URL地址"));break;
            case Acad::eInetNoWinInet:lstrcpy(Glb_AcadErrorInfo,_T("没有WinInet(?)"));break;
            case Acad::eInetOldWinInet:lstrcpy(Glb_AcadErrorInfo,_T("旧的WinInet(?)"));break;
            case Acad::eInetNoAcadInet:lstrcpy(Glb_AcadErrorInfo,_T("无法连接ACAD网站"));break;
            case Acad::eInetNotImplemented:lstrcpy(Glb_AcadErrorInfo,_T("无法应用网络"));break;
            case Acad::eInetProtocolNotSupported:lstrcpy(Glb_AcadErrorInfo,_T("网络协议不支持"));break;
            case Acad::eInetCreateInternetSessionFailed:lstrcpy(Glb_AcadErrorInfo,_T("创建网络会话失败"));break;
            case Acad::eInetInternetSessionConnectFailed:lstrcpy(Glb_AcadErrorInfo,_T("连接网络会话失败"));break;
            case Acad::eInetInternetSessionOpenFailed:lstrcpy(Glb_AcadErrorInfo,_T("打开网络会话失败"));break;
            case Acad::eInetInvalidAccessType:lstrcpy(Glb_AcadErrorInfo,_T("无效的网络接收类型"));break;
            case Acad::eInetFileOpenFailed:lstrcpy(Glb_AcadErrorInfo,_T("打开网络文件失败"));break;
            case Acad::eInetHttpOpenRequestFailed:lstrcpy(Glb_AcadErrorInfo,_T("打开HTTP协议失败"));break;
            case Acad::eInetUserCancelledTransfer:lstrcpy(Glb_AcadErrorInfo,_T("用户取消了网络传输"));break;
            case Acad::eInetHttpBadRequest:lstrcpy(Glb_AcadErrorInfo,_T("不合理的网络请求"));break;
            case Acad::eInetHttpAccessDenied:lstrcpy(Glb_AcadErrorInfo,_T("HTTP协议拒绝"));break;
            case Acad::eInetHttpPaymentRequired:lstrcpy(Glb_AcadErrorInfo,_T("HTTP协议要求付费"));break;
            case Acad::eInetHttpRequestForbidden:lstrcpy(Glb_AcadErrorInfo,_T("禁止HTTP请求"));break;
            case Acad::eInetHttpObjectNotFound:lstrcpy(Glb_AcadErrorInfo,_T("HTTP对象未找到"));break;
            case Acad::eInetHttpBadMethod:lstrcpy(Glb_AcadErrorInfo,_T("不合理的HTTP请求方法"));break;
            case Acad::eInetHttpNoAcceptableResponse:lstrcpy(Glb_AcadErrorInfo,_T("不接受的HTTP回复"));break;
            case Acad::eInetHttpProxyAuthorizationRequired:lstrcpy(Glb_AcadErrorInfo,_T("要求HTTP代理授权"));break;
            case Acad::eInetHttpTimedOut:lstrcpy(Glb_AcadErrorInfo,_T("HTTP超时"));break;
            case Acad::eInetHttpConflict:lstrcpy(Glb_AcadErrorInfo,_T("HTTP冲突"));break;
            case Acad::eInetHttpResourceGone:lstrcpy(Glb_AcadErrorInfo,_T("网络资源被用光"));break;
            case Acad::eInetHttpLengthRequired:lstrcpy(Glb_AcadErrorInfo,_T("HTTP请求长度是必须的"));break;
            case Acad::eInetHttpPreconditionFailure:lstrcpy(Glb_AcadErrorInfo,_T("HTTP预处理失败"));break;
            case Acad::eInetHttpRequestTooLarge:lstrcpy(Glb_AcadErrorInfo,_T("HTTP请求太大"));break;
            case Acad::eInetHttpUriTooLong:lstrcpy(Glb_AcadErrorInfo,_T("URL地址太长"));break;
            case Acad::eInetHttpUnsupportedMedia:lstrcpy(Glb_AcadErrorInfo,_T("HTTP不支持的媒体"));break;
            case Acad::eInetHttpServerError:lstrcpy(Glb_AcadErrorInfo,_T("HTTP服务器错误"));break;
            case Acad::eInetHttpNotSupported:lstrcpy(Glb_AcadErrorInfo,_T("HTTP不支持"));break;
            case Acad::eInetHttpBadGateway:lstrcpy(Glb_AcadErrorInfo,_T("HTTP网关错误"));break;
            case Acad::eInetHttpServiceUnavailable:lstrcpy(Glb_AcadErrorInfo,_T("HTTP服务当前不可用"));break;
            case Acad::eInetHttpGatewayTimeout:lstrcpy(Glb_AcadErrorInfo,_T("HTTP网关超时"));break;
            case Acad::eInetHttpVersionNotSupported:lstrcpy(Glb_AcadErrorInfo,_T("HTTP版本不支持"));break;
            case Acad::eInetInternetError:lstrcpy(Glb_AcadErrorInfo,_T("HTTP网络错误"));break;
            case Acad::eInetGenericException:lstrcpy(Glb_AcadErrorInfo,_T("HTTP常规异常"));break;
            case Acad::eInetUnknownError:lstrcpy(Glb_AcadErrorInfo,_T("HTTP未知错误"));break;
            case Acad::eAlreadyActive:lstrcpy(Glb_AcadErrorInfo,_T("已经是活动的了"));break;
            case Acad::eAlreadyInactive:lstrcpy(Glb_AcadErrorInfo,_T("已经是不活动的了"));break;
     
         */

    }
}
