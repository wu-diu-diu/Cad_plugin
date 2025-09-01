using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static CoDesignStudy.Cad.PlugIn.Command;
using CADApplication = Autodesk.AutoCAD.ApplicationServices.Application;


namespace CoDesignStudy.Cad.PlugIn
{
    public static class CadDrawingHelper
    {
        public static Dictionary<string, (double Count, string Info)> componentStats = new Dictionary<string, (double, string)>();

        public static Point3d GetCenterFromExtents(Extents3d ext)
        {
            double centerX = (ext.MinPoint.X + ext.MaxPoint.X) / 2.0;
            double centerY = (ext.MinPoint.Y + ext.MaxPoint.Y) / 2.0;
            double centerZ = (ext.MinPoint.Z + ext.MaxPoint.Z) / 2.0;
            return new Point3d(centerX, centerY, centerZ);
        }
        public static void MergeLightingFromModelReply(string modelReplyJson)
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            try
            {
                // 反序列化模型回复
                var lightingDesignResponse = JsonConvert.DeserializeObject<LightingDesignResponse>(modelReplyJson);

                // 构造路径，读取房间分析结果 JSON 文件
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "cad_rect_result.json");
                if (File.Exists(filePath))
                {
                    string analysisJson = File.ReadAllText(filePath);
                    // 直接反序列化为 RoomInfo（结构相同）
                    var roomInfo = JsonConvert.DeserializeObject<LightingDesignResponse.RoomInfo>(analysisJson);
                    lightingDesignResponse.room_info = roomInfo;
                    string finalJson = JsonConvert.SerializeObject(lightingDesignResponse, Formatting.Indented);

                    // 自动生成递增的文件名 final_result_1.json, final_result_2.json, ...
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    int index = 1;
                    string savePath;

                    do
                    {
                        savePath = Path.Combine(desktopPath, $"final_result_{index}.json");
                        index++;
                    }
                    while (File.Exists(savePath));

                    File.WriteAllText(savePath, finalJson, System.Text.Encoding.UTF8);
                    ed.WriteMessage($"\n合并结果已导出到: {savePath}");
                }
                else
                {
                    ed.WriteMessage($"\n找不到分析结果文件: {filePath}");
                }

                // TODO: 在此处使用 lightingDesignResponse，比如生成CAD布图等后续逻辑

                ed.WriteMessage("\n模型和分析数据已成功合并。");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n处理过程中发生错误: {ex.Message}");
            }
        }
        /// <summary>
        /// 最重要的函数，从模型输出的Json中解析照明设计结果并插入图元
        /// </summary>
        public static void InsertLightingFromModelReply(string modelReplyJson)
        {
            var obj = JsonConvert.DeserializeObject<LightingDesignResponse>(modelReplyJson);

            // 插入注释信息
            List<double> annotation_coords = obj.lighting_design.annotation_position_mm;
            if (annotation_coords != null && annotation_coords.Count == 2)
            {
                Point3d annotationPoint = new Point3d(annotation_coords[0], annotation_coords[1], 0);
                InsertLightingInfo(annotationPoint, obj.lighting_design.fixture_count, obj.lighting_design.power_w, obj.lighting_design.mounting_height_mm / 1000);
            }
            // 保存灯具坐标点以供线路连接
            List<Point3d> insertPoints = new List<Point3d>();
            string lightType = obj.lighting_design.fixture_type;
            string lightName = lightType.Contains("吸顶") ? "感应式吸顶灯" :
                   lightType.Contains("防爆") ? "防爆灯" :
                   (lightType.Contains("面板") || lightType.Contains("荧光")) ? "双管荧光灯" :
                   "gen_light";
            string LightLayer = "照明";
            int lightCount = obj.lighting_design.fixture_count;
            int lightRotation = obj.lighting_design.fixture_rotations_degrees;
            // 插入灯具
            foreach (var point in obj.lighting_design.fixture_positions_mm)
            {
                if (point.Count >= 2)
                {
                    double x = point[0];
                    double y = point[1];
                    double z = 0;
                    insertPoints.Add(new Point3d(x, y, z));

                    ObjectId id = InsertBlockFromDwg(new Point3d(x, y, z), LightLayer, lightName, lightRotation);
                    // 记录本次插入的图元ID
                    InsertTracker.AddEntity(id);
                    // 统计信息
                }
            }
            // 统计所有的插入灯具信息
            AddOrUpdateComponent(lightName, lightCount, obj.lighting_design.power_w.ToString() + "W");
            // 统计上一次插入的灯具信息
            InsertTracker.AddComponentCount(lightName, lightCount);
            // 绘制灯具之间的连线
            string wiringLayer = "照明连线";
            double lightLength = 0;
            var lightcolor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2);
            foreach (var line in obj.lighting_design.fixture_wiring_lines_mm)
            {
                if (line.Count >= 2)
                {
                    var start = line[0];
                    var end = line[1];
                    if (start.Count >= 2 && end.Count >= 2)
                    {
                        Point3d p1 = new Point3d(start[0], start[1], 0);
                        Point3d p2 = new Point3d(end[0], end[1], 0);
                        (ObjectId polyId, int length) = DrawPolyLineBetweenPoints(p1, p2, wiringLayer, 20f, lightcolor);
                        lightLength += length; // 累加长度
                        InsertTracker.AddEntity(polyId);
                    }
                }
            }
            // 绘制引出线
            foreach (var line in obj.lighting_design.power_outlet_connection_line_mm)
            {
                if (line.Count >= 2)
                {
                    var start = line[0];
                    var end = line[1];
                    if (start.Count >= 2 && end.Count >= 2)
                    {
                        Point3d p1 = new Point3d(start[0], start[1], 0);
                        Point3d p2 = new Point3d(end[0], end[1], 0);
                        (ObjectId polyId, int length) = DrawPolyLineBetweenPoints(p1, p2, wiringLayer, 20f, lightcolor);
                        lightLength += length; // 累加长度
                        InsertTracker.AddEntity(polyId);
                    }
                }
            }
            lightLength = Math.Round(lightLength / 1000.0, 1);
            // 统计插入的照明线路
            AddOrUpdateComponent("照明线路", lightLength, "BV-500 4mm²");
            // 统计上一次插入的照明线路信息
            InsertTracker.AddComponentCount("照明线路", lightLength);
            // 插入插座
            int switch_count = obj.switch_position.fixture_count;
            if (obj.socket_positions != null)
            {
                string socketLayer = "插座"; // 替换为你项目中的图层名
                foreach (var socket in obj.socket_positions)
                {
                    if (socket.position_mm.Count >= 2)
                    {
                        double x = socket.position_mm[0];
                        double y = socket.position_mm[1];
                        double z = 0;
                        double rotationDeg = socket.rotation_degrees;

                        ObjectId id = InsertBlockFromDwg(new Point3d(x, y, z), socketLayer, "三相五孔插座", rotationDeg);
                        InsertTracker.AddEntity(id);
                    }
                }
            }
            // 统计插座信息
            AddOrUpdateComponent("三相五孔插座", switch_count, "220V 10A");
            // 统计上一次插入的插座信息
            InsertTracker.AddComponentCount("三相五孔插座", switch_count);
            // 绘制插座之间的连线
            wiringLayer = "插座连线";
            double socketLength = 0;
            var socketcolor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 6);
            foreach (var line in obj.socket_wiring_lines_mm)
            {
                if (line.Count >= 2)
                {
                    var start = line[0];
                    var end = line[1];
                    if (start.Count >= 2 && end.Count >= 2)
                    {
                        Point3d p1 = new Point3d(start[0], start[1], 0);
                        Point3d p2 = new Point3d(end[0], end[1], 0);
                        (ObjectId polyId, int length) = DrawPolyLineBetweenPoints(p1, p2, wiringLayer, 20f, socketcolor);
                        socketLength += length; // 累加长度
                        InsertTracker.AddEntity(polyId);
                    }
                }
            }
            socketLength = Math.Round(socketLength / 1000.0, 1);
            // 统计插入的插座线路
            AddOrUpdateComponent("插座线路", socketLength, "BV-500 2mm²");
            // 统计上一次插入的插座线路信息
            InsertTracker.AddComponentCount("插座线路", socketLength);
            // 插入开关
            if (obj.switch_position?.position_mm != null && obj.switch_position.position_mm.Count >= 2)
            {
                string switchLayer = "开关"; // 替换为你项目中的图层名
                double x = obj.switch_position.position_mm[0];
                double y = obj.switch_position.position_mm[1];
                double z = 0;
                ObjectId id = InsertBlockFromDwg(new Point3d(x, y, z), switchLayer, "开关");
                InsertTracker.AddEntity(id);
            }
            // 统计开关信息
            AddOrUpdateComponent("开关", 1, "220V 10A");
            // 统计上一次插入的开关信息
            InsertTracker.AddComponentCount("开关", 1);
            // 完成本次记录
            InsertTracker.CommitInsert();
        }
        /// <summary>
        /// 每次生成一组元件后，将这组元件的数量，类型等统计到一个字典中，方便后续生成材料清单
        /// </summary>
        private static void AddOrUpdateComponent(string type, double count, string info)
        {
            if (componentStats.ContainsKey(type))
            {
                var old = componentStats[type];
                componentStats[type] = (old.Count + count, old.Info); // 保持原 info
            }
            else
            {
                componentStats[type] = (count, info);
            }
        }
        /// <summary>
        /// 给定插入坐标，图层名，块名，从指定路径插入块参照，并可设置旋转角度。依靠该函数实现图纸自动绘制
        /// </summary>
        public static ObjectId InsertBlockFromDwg(Point3d insertPoint, string targetLayer, string blockName, double rotationDegrees = 0)
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;
            // 构建 DWG 文件路径
            string blocksDirectory = @"C:\Users\武丢丢\Documents\gen_light\";
            string dwgFilePath = Path.Combine(blocksDirectory, blockName + ".dwg");

            if (!File.Exists(dwgFilePath))
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog($"找不到块文件：{dwgFilePath}");
                return ObjectId.Null;
            }

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 导入块定义到当前数据库
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord { Name = targetLayer };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    // 插入块参照，设置旋转角度（单位为弧度）
                    double rotationRadians = rotationDegrees * Math.PI / 180.0;
                    BlockReference br = new BlockReference(insertPoint, bt[blockName])
                    {
                        Layer = targetLayer,
                        Rotation = rotationRadians
                    };
                    blockId = modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    tr.Commit();
                }
            }
            return blockId;
        }
        /// <summary>
        /// 在图纸中插入灯具信息文本，格式为：
        /// 上：功率（如 "100W"）
        /// 下：安装高度（如 "4.5"）
        /// 左：灯具数量（如 "4"）
        /// </summary>
        /// <param name="basePosition">文本中心点（用于功率位置）</param>
        /// <param name="fixtureCount">灯具数量</param>
        /// <param name="powerW">功率（单位 W）</param>
        /// <param name="mountingHeight">安装高度（单位 m）</param>
        /// <param name="textHeight">文字高度（默认 300）</param>
        /// <param name="layerName">图层名称（可选）</param>
        //[CommandMethod("PP", CommandFlags.Session)]
        private static void InsertLightingInfo(Point3d basePosition, int fixtureCount, int powerW, int mountingHeight, double textHeight = 300, string layerName = "PUB_TEXT")
        {
            //Point3d basePosition = new Point3d(34640, 55259, 0);
            //int fixtureCount = 4;
            //int powerW = 100;
            //double mountingHeight = 4.5;
            //int textHeight = 300;
            //string layerName = "PUB_TEXT";
            // 插入中间的功率文本（如 "100W"）
            ObjectId id1 = InsertTextAt($"{powerW}W", basePosition, layerName, textHeight);
            InsertTracker.AddEntity(id1);

            // 插入下方的安装高度文本（如 "4.5"）
            var below = new Point3d(basePosition.X, basePosition.Y - textHeight * 1.4, basePosition.Z);
            ObjectId id2 = InsertTextAt($"{mountingHeight:0.##}", below, layerName, textHeight);
            InsertTracker.AddEntity(id2);

            // 插入左侧的灯具数量文本（如 "4"）
            var left = new Point3d(basePosition.X - textHeight * 1.4, basePosition.Y - textHeight * 0.8, basePosition.Z);
            ObjectId id3 = InsertTextAt($"{fixtureCount}", left, layerName, textHeight);
            InsertTracker.AddEntity(id3);

            // 计算直线起止点（在瓦数和高度之间，略微留白）
            double halfLineLength = textHeight * 2.0;
            var lineY = basePosition.Y - 85; // 在两者中间
            Point3d lineStart = new Point3d(basePosition.X - halfLineLength + 500, lineY, basePosition.Z);
            Point3d lineEnd = new Point3d(basePosition.X + halfLineLength + 500, lineY, basePosition.Z);

            // 绘制黄色横线
            ObjectId id4 = DrawYellowLine(lineStart, lineEnd, layerName);
            InsertTracker.AddEntity(id4);
        }
        /// <summary>
        /// 在指定坐标插入单行文本（DBText）
        /// </summary>
        /// <param name="text">要插入的文本内容</param>
        /// <param name="position">插入点坐标</param>
        /// <param name="layerName">可选，插入到的图层，默认为当前图层</param>
        private static ObjectId InsertTextAt(string text, Point3d position, string layerName = null, double textHeight = 300)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId textID;

            using (DocumentLock docLock = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                MText mtext = new MText
                {
                    Contents = text,
                    Location = position,
                    TextHeight = textHeight,
                    Width = 0, // 0 表示不自动换行，可设置为具体值以控制列宽
                    Attachment = AttachmentPoint.BottomLeft, // 左下角对齐
                    Layer = layerName,
                    Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2)
                };

                //// 创建图层（如果不存在）
                //if (!string.IsNullOrEmpty(layerName))
                //{
                //    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                //    if (!lt.Has(layerName))
                //    {
                //        lt.UpgradeOpen();
                //        LayerTableRecord newLayer = new LayerTableRecord { Name = layerName };
                //        lt.Add(newLayer);
                //        tr.AddNewlyCreatedDBObject(newLayer, true);
                //    }
                //    dbText.Layer = layerName;
                //}

                textID = modelSpace.AppendEntity(mtext);
                tr.AddNewlyCreatedDBObject(mtext, true);

                tr.Commit();
            }
            return textID;
        }
        private static ObjectId DrawYellowLine(Point3d start, Point3d end, string layerName)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId yellowlineID;

            using (DocumentLock docLock = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 确保图层存在
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(layerName))
                {
                    lt.UpgradeOpen();
                    LayerTableRecord newLayer = new LayerTableRecord { Name = layerName };
                    lt.Add(newLayer);
                    tr.AddNewlyCreatedDBObject(newLayer, true);
                }

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                Line line = new Line(start, end)
                {
                    Layer = layerName,
                    Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2) // 2 = 黄色
                };

                yellowlineID = modelSpace.AppendEntity(line);
                tr.AddNewlyCreatedDBObject(line, true);

                tr.Commit();
            }
            return yellowlineID;
        }
        /// <summary>
        /// 两点之间绘制多段线
        /// </summary>
        private static (ObjectId, int) DrawPolyLineBetweenPoints(Point3d pt1, Point3d pt2, string layerName, double lineWidth, Autodesk.AutoCAD.Colors.Color color)
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            int polylineLength = 0;
            ObjectId polyId = ObjectId.Null;

            using (DocumentLock docLock = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 确保图层存在
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(layerName))
                {
                    lt.UpgradeOpen();
                    LayerTableRecord newLayer = new LayerTableRecord { Name = layerName };
                    lt.Add(newLayer);
                    tr.AddNewlyCreatedDBObject(newLayer, true);
                }

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // 创建多段线
                Polyline poly = new Polyline();
                poly.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, (float)lineWidth, (float)lineWidth); // 线宽35f
                poly.AddVertexAt(1, new Point2d(pt2.X, pt2.Y), 0, (float)lineWidth, (float)lineWidth); // 线宽35f
                poly.Layer = layerName;
                poly.Color = color;
                poly.ConstantWidth = (float)lineWidth;

                polyId = modelSpace.AppendEntity(poly);
                tr.AddNewlyCreatedDBObject(poly, true);
                // 获取多段线长度
                polylineLength = (int)poly.Length;
                tr.Commit();
            }
            return (polyId, polylineLength);
        }
        /// <summary>
        /// 使用多段线绘制识别到的房间矩形，辅助检查识别的区域
        /// </summary>
        private static void DrawRectanglePolyline(List<int[]> rectPoints, string layerName = "Rect")
        {
            if (rectPoints == null || rectPoints.Count != 4)
                throw new ArgumentException("必须提供四个矩形角点。");

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (DocumentLock docLock = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 确保图层存在
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(layerName))
                {
                    lt.UpgradeOpen();
                    LayerTableRecord newLayer = new LayerTableRecord { Name = layerName };
                    lt.Add(newLayer);
                    tr.AddNewlyCreatedDBObject(newLayer, true);
                }

                // 获取模型空间
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // 创建闭合多段线
                Polyline poly = new Polyline();
                for (int i = 0; i < 4; i++)
                {
                    int[] pt = rectPoints[i];
                    poly.AddVertexAt(i, new Point2d(pt[0], pt[1]), 0, 0, 0);
                }
                poly.Closed = true;
                poly.Layer = layerName;
                poly.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 255, 0); // 可自定义颜色
                poly.ConstantWidth = 30f;

                modelSpace.AppendEntity(poly);
                tr.AddNewlyCreatedDBObject(poly, true);
                tr.Commit();
            }
        }
        /// <summary>
        /// 绘制材料清单表格
        /// </summary>
        /// <param name="Type">元件类型</param>
        /// <param name="Count">元件数量</param>
        /// <param name="Info">元件注释信息</param>
        public static void DrawComponentTable6ColsWithBlocks(List<(string Type, double Count, string Info)> statsList)
        {
            double rowHeight = 800;
            short colorIndex = 7;
            LineWeight lineWeight = LineWeight.LineWeight200;
            Point3d insertPoint = new Point3d(3599, 17119, 0);  // 在一个固定的位置绘制表格
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            using (DocumentLock docLock = doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                int rowCount = statsList.Count;
                int totalRowCount = rowCount + 2;
                int colCount = 6;

                // 指定列宽
                double[] colWidths = { 900, 4000, 4000, 900, 1200, 1500 };
                string[] headerTitles = { "编号", "名称", "规范", "单位", "数量", "图例" };

                // 计算每列起始X坐标
                double[] colX = new double[colCount + 1];
                colX[0] = insertPoint.X;
                for (int i = 1; i <= colCount; i++)
                {
                    colX[i] = colX[i - 1] + colWidths[i - 1];
                }

                double tableHeight = totalRowCount * rowHeight;

                // 画水平线
                for (int i = 0; i <= totalRowCount; i++)
                {
                    double y = insertPoint.Y - i * rowHeight;
                    var start = new Point3d(colX[0], y, 0);
                    var end = new Point3d(colX[colCount], y, 0);
                    var line = new Line(start, end)
                    {
                        Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, colorIndex),
                        LineWeight = lineWeight
                    };
                    btr.AppendEntity(line);
                    tr.AddNewlyCreatedDBObject(line, true);
                }

                // 画垂直线
                for (int j = 0; j <= colCount; j++)
                {
                    Point3d start;
                    if (j == 0 || j == colCount)
                    {
                        start = new Point3d(colX[j], insertPoint.Y, 0);
                    }
                    else
                    {
                        start = new Point3d(colX[j], insertPoint.Y - rowHeight, 0);
                    }
                    var end = new Point3d(colX[j], insertPoint.Y - tableHeight, 0);
                    var line = new Line(start, end)
                    {
                        Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, colorIndex),
                        LineWeight = lineWeight
                    };
                    btr.AppendEntity(line);
                    tr.AddNewlyCreatedDBObject(line, true);
                }
                // 插入标题：第 0 行，居中插入
                {
                    double titleCenterX = (colX[colCount] + colX[0]) / 2;
                    double titleY = insertPoint.Y - rowHeight * 0.6;

                    var mtext = new MText
                    {
                        Contents = "设备材料清单",
                        TextHeight = rowHeight * 0.5,
                        Location = new Point3d(titleCenterX, titleY, 0),
                        Attachment = AttachmentPoint.MiddleCenter,
                        Width = colX[colCount] - colX[0]
                    };

                    btr.AppendEntity(mtext);
                    tr.AddNewlyCreatedDBObject(mtext, true);
                }
                // 插入表头行（第1行）
                for (int j = 0; j < colCount; j++)
                {
                    double x = colX[j] + 100;
                    double y = insertPoint.Y - rowHeight * 1.6;

                    var text = new DBText
                    {
                        TextString = headerTitles[j],
                        Position = new Point3d(x, y, 0),
                        Height = rowHeight * 0.35,
                        Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7)
                    };

                    btr.AppendEntity(text);
                    tr.AddNewlyCreatedDBObject(text, true);
                }

                // 插入文本：列顺序 = 行号, type, info, "只", count, 空
                for (int i = 0; i < rowCount; i++)
                {
                    var (type, count, info) = statsList[i];
                    string unit = type.EndsWith("线路") ? "米" : "只";
                    string[] values = {
                                        (i+1).ToString(),    // 第1列：行号
                                        type,            // 第2列：元件类型
                                        info,            // 第3列：元件信息
                                        unit,            // 第4列：单位
                                        count.ToString(),// 第5列：数量
                                        ""               // 图例（插入块）
                                    };

                    for (int j = 0; j < colCount; j++)
                    {
                        // 块插入逻辑（第六列）
                        if (j == 5)
                        {
                            double blockX = colX[j] + colWidths[j] / 2;
                            double blockY = insertPoint.Y - (i + 2) * rowHeight - rowHeight / 2;
                            var blockPoint = new Point3d(blockX, blockY, 0);
                            InsertBlockFromDwg(blockPoint, "PUB_TEXT", type, 0);
                            continue;
                        }
                        double x = colX[j] + 100; // 边距内缩
                        double y = insertPoint.Y - (i + 2) * rowHeight - rowHeight * 0.6;

                        var mtext = new MText
                        {
                            Contents = values[j],
                            Location = new Point3d(x, y, 0),
                            TextHeight = rowHeight * 0.35,
                            Width = colWidths[j] - 200, // 控制列内宽度，避免溢出
                            Attachment = AttachmentPoint.MiddleLeft, // 文字对齐方式
                            Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7)
                        };
                        btr.AppendEntity(mtext);
                        tr.AddNewlyCreatedDBObject(mtext, true);
                    }
                }

                tr.Commit();
            }
        }
        /// <summary>
        /// 识别到的房间轮廓进行缩小处理
        /// </summary>
        /// <param name="rectPoints">识别到的房间轮廓的四个点</param>
        /// <param name="offsetX">x轴方向缩小量</param>
        /// <param name="offsetY">y轴方向缩小量</param>
        public static List<int[]> ShrinkRectangle(List<int[]> rectPoints, int offsetX, int offsetY)
        {
            if (rectPoints == null || rectPoints.Count != 4)
                throw new ArgumentException("矩形必须有4个点");

            // 将每个点包装成带索引的结构
            var sorted = rectPoints
                .Select(pt => new { X = pt[0], Y = pt[1], Point = pt })
                .ToList();

            // 找出四个角
            var leftBottom = sorted.OrderBy(p => p.X).ThenBy(p => p.Y).First().Point;
            var leftTop = sorted.OrderBy(p => p.X).ThenByDescending(p => p.Y).First().Point;
            var rightTop = sorted.OrderByDescending(p => p.X).ThenByDescending(p => p.Y).First().Point;
            var rightBottom = sorted.OrderByDescending(p => p.X).ThenBy(p => p.Y).First().Point;

            // 缩小每个角点
            var newPoints = new List<int[]>
            {
                new int[] { leftBottom[0] + offsetX, leftBottom[1] + offsetY },       // 左下
                new int[] { leftTop[0] + offsetX, leftTop[1] - offsetY * 2 },             // 左上
                new int[] { rightTop[0] - offsetX * 2, rightTop[1] - offsetY * 2 },           // 右上
                new int[] { rightBottom[0] - offsetX * 2, rightBottom[1] + offsetY }      // 右下
            };

            return newPoints;
        }

    }
}