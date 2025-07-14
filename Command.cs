using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Autodesk.Windows;
using BoundingRectangle;
using Markdig;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.Schema;
using static System.Net.Mime.MediaTypeNames;
using CADApplication = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: ExtensionApplication(typeof(CoDesignStudy.Cad.PlugIn.Command))]
[assembly: CommandClass(typeof(CoDesignStudy.Cad.PlugIn.Command))]
namespace CoDesignStudy.Cad.PlugIn
{
    /// <summary>
    /// 入口
    /// </summary>
    
    public class Command : IExtensionApplication
    {
        #region 成员变量
        public string FinalPrompt;
        public string FinalReply;
        public static PaletteSetDlg DlgInstance;
        Dictionary<string, (double Count, string Info)> componentStats = new Dictionary<string, (double, string)>();

        #endregion

        #region 初始化
        public void Initialize()
        {
            PrjExploreHelper.InitMenu();
        }

        public void Terminate()
        {

        }
        #endregion

        #region 测试命令

        [CommandMethod("TT", CommandFlags.Session)]
        public static void SDF_CMF()
        {
            try
            {
                Document curDoc = CADApplication.DocumentManager.MdiActiveDocument;
                if (curDoc == null)
                {
                    throw new System.Exception("当前未开启任何文件！");
                }
                // 以下代码初始化了一个paletteset类为mainpaletteset，用来管理自定义的面板
                PrjExploreHelper.InitPalette();
                // palettesetdlg是继承自usercontrol，用户实现具体的界面和交互逻辑，它是一个winform控件，不能直接在cad界面中显示，必须添加到paletteset中
                DlgInstance = new PaletteSetDlg();
                // 将自定义的winform控件添加到mainpaletteset之后，侧边栏才能显示自定义的控件比如发送按钮，聊天的背景颜色等等。
                PrjExploreHelper.MainPaletteset.Add("测试界面", DlgInstance);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("测试", ex.Message);
            }
        }
        public static void ShowModelReplyInPanel(string reply)
        {
            try
            {
                Document curDoc = CADApplication.DocumentManager.MdiActiveDocument;
                if (curDoc == null)
                {
                    throw new System.Exception("当前未开启任何文件！");
                }

                // 初始化 PaletteSet（只初始化一次）
                if (PrjExploreHelper.MainPaletteset == null)
                {
                    PrjExploreHelper.InitPalette();
                }

                // 初始化面板实例（只初始化一次）
                if (Command.DlgInstance == null)
                {
                    Command.DlgInstance = new PaletteSetDlg();
                    PrjExploreHelper.MainPaletteset.Add("测试界面", Command.DlgInstance);
                }

                // 展示模型输出内容
                Command.DlgInstance.AppendMessageSync("AI", reply).Wait();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("显示面板失败", ex.Message);
            }
        }

        [CommandMethod("URL", CommandFlags.Session)]
        public static void URL()
        {
            try
            {
                Document curDoc = CADApplication.DocumentManager.MdiActiveDocument;
                if (curDoc == null)
                {
                    throw new System.Exception("当前未开启任何文件！");
                }
                PrjExploreHelper.InitWebPalette();
                var panel = new WebServicePanel("http://1.119.159.4:48840/");
                PrjExploreHelper.MainPaletteset.Add("Web服务", panel);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("测试", ex.Message);
            }
        }

        [CommandMethod("AA", CommandFlags.Session)]
        public void HighlightALLTexts()
        {
            // DocumentManager是对多个文档进行管理的，每个dwg文件视为一个文档
            // MdiActiveDocument表示当前激活的文档，即当前界面打开的文档
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;

            // 每一个dwg文件可以看作一个数据库，类似一个excel表格，存储着不同类别的数据
            // blocktable是所有的图元数据，即线，圆，所有能在图纸中选中的图形
            // layertable是所有的图层数据
            // dimstyletable是所有的文字数据

            Database db = doc.Database;  // 获取当前文档的数据库对象
            Editor ed = doc.Editor;  // 获取当前文档的编辑器对象
            // 启动事务，这里的tr可以看作图元，即图纸
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 以只读模式打开图元的数据库的
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;  // 获取模型空间

                foreach (ObjectId id in btr)
                {
                    // ent是当前遍历到的autocad图形对象，包括文字，直线，圆等所有可见图形对象都继承自Entity
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent is DBText text)
                    {
                        if (IsChineseText(text.TextString))
                        {
                            ent.Highlight();
                        }
                    }
                }
                tr.Commit();
            }
        }
        [CommandMethod("BB", CommandFlags.Session)]
        public void HighlightBlocksByName()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                foreach (ObjectId id in space)
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent is BlockReference br)
                    {
                        // 获取块名
                        string blockName = "";
                        if (br.BlockTableRecord.IsValid)
                        {
                            BlockTableRecord brDef = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                            blockName = brDef.Name;
                        }

                        // 这里可以根据块名进行筛选，比如高亮所有名为"ZhaoMing"的块
                        if (blockName == "door1" || blockName == "door2" || blockName == "door3")
                        {
                            ent.Highlight();
                        }
                    }
                }
                tr.Commit();
            }
        }
        private bool IsChineseText(string text)
        {
            // 正则表达式匹配中文字符（包括标点符号）
            return Regex.IsMatch(text, @"[\u4e00-\u9fa5]");
        }

        [CommandMethod("QQ", CommandFlags.Session)]
        public void DrawCircleWithLisp()
        {
            // 获取当前文档
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;  // 打开当前激活的文档
            // AutoLISP 代码：在(100,100)处画半径为50的圆
            string lisp = "(command \"CIRCLE\" '(100 100 0) 1500) ";
            // 执行 AutoLISP 代码
            doc.SendStringToExecute(lisp, true, false, false);
        }
        [CommandMethod("EE", CommandFlags.Session)]
        public void ExportExcel()
        {
            var statsList = componentStats
                        .Select(kv => (Type: kv.Key, Count: kv.Value.Count, Info: kv.Value.Info))
                        .ToList();

            //ExportStatisticsToExcel(statsList, @"D:\最终统计.xlsx");
            Point3d insertPoint = new Point3d(34640, 55259, 0);

            DrawComponentTable6ColsWithBlocks(statsList);
        }

        [CommandMethod("PRINT_ALL_BLOCK_NAMES", CommandFlags.Session)]
        public void PrintAllBlockNames()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                doc.Editor.WriteMessage("\n当前图纸中的所有块名称：");
                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    // 过滤掉匿名块和布局块（如需全部显示可去掉此判断）
                    if (!btr.IsAnonymous && !btr.IsLayout)
                    {
                        doc.Editor.WriteMessage($"\n{btr.Name}");
                    }
                }
                tr.Commit();
            }
        }
        [CommandMethod("EXTRACT_LAYER_ENTS", CommandFlags.Session)]
        public void ExtractEntitiesFromLayer()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 提示用户输入图层名
            PromptStringOptions pso = new PromptStringOptions("\n请输入要提取的图层名：");
            pso.AllowSpaces = true;
            PromptResult pr = ed.GetString(pso);

            string targetLayer = pr.StringResult.Trim();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                List<Entity> foundEntities = new List<Entity>();

                foreach (ObjectId id in modelSpace)
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent != null && ent.Layer == targetLayer)
                    {
                        foundEntities.Add(ent);
                    }
                }

                ed.WriteMessage($"\n图层 \"{targetLayer}\" 内共找到 {foundEntities.Count} 个图元：");
                foreach (var ent in foundEntities)
                {
                    if (ent is BlockReference br)
                    {
                        // 获取块名
                        string blockName = "";
                        if (br.BlockTableRecord.IsValid)
                        {
                            BlockTableRecord brDef = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                            blockName = brDef.Name;
                        }
                        ed.WriteMessage($"\n类型: BlockReference, 块名: {blockName}, 对象ID: {ent.ObjectId}");
                    }
                    else
                    {
                        ed.WriteMessage($"\n类型: {ent.GetType().Name}, 对象ID: {ent.ObjectId}");
                    }
                }

                // 便于debug时查看
                System.Diagnostics.Debugger.Break();

                tr.Commit();
            }
        }
        private static Point3d GetCenterFromExtents(Extents3d ext)
        {
            double centerX = (ext.MinPoint.X + ext.MaxPoint.X) / 2.0;
            double centerY = (ext.MinPoint.Y + ext.MaxPoint.Y) / 2.0;
            double centerZ = (ext.MinPoint.Z + ext.MaxPoint.Z) / 2.0;
            return new Point3d(centerX, centerY, centerZ);
        }
        public class AnalysisResult
        {
            public Dictionary<string, List<double[]>> Blocks { get; set; } = new Dictionary<string, List<double[]>>();
            public Dictionary<string, int> BlockCount { get; set; } = new Dictionary<string, int>();
            public Dictionary<string, double> PolylineLengthByLayer { get; set; } = new Dictionary<string, double>();
            public List<object> Texts { get; set; } = new List<object>();
            public List<double[]> RectPoints { get; set; } = new List<double[]>();
        }

        public class LightingDesignResponse
        {
            public RoomInfo room_info { get; set; }
            public LightingDesign lighting_design { get; set; }
            public List<SocketPosition> socket_positions { get; set; }
            public SwitchPosition switch_position { get; set; }

            public class RoomInfo
            {
                public string room_type { get; set; }
                public List<List<int>> coordinates_mm { get; set; }
                public Dimensions dimensions { get; set; }
                public int illuminance_standard_lx { get; set; }
            }

            public class Dimensions
            {
                public int length_mm { get; set; }
                public int width_mm { get; set; }
                public double area_m2 { get; set; }
            }

            public class LightingDesign
            {
                public string fixture_type { get; set; }
                public int fixture_count { get; set; }
                public int power_w { get; set; }
                public int mounting_height_mm { get; set; }
                public int fixture_rotations_degrees { get; set; }
                public List<double> annotation_position_mm { get; set; }
                public List<List<double>> fixture_positions_mm { get; set; }
                public List<List<List<double>>> fixture_wiring_lines_mm { get; set; }
                public List<List<List<double>>> power_outlet_connection_line_mm { get; set; }

            }

            public class SocketPosition
            {
                public List<double> position_mm { get; set; }
                public int rotation_degrees { get; set; }
                public int fixture_count { get; set; }
            }

            public class SwitchPosition
            {
                public List<double> position_mm { get; set; }
                public int fixture_count { get; set; }
            }
        }


        [CommandMethod("SELECT_RECT_PRINT", CommandFlags.Session)]
        public async void SelectEntitiesByRectangleAndPrintInfo()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            // 保存门的坐标
            List<Point3d> doorPoints = new List<Point3d>();
            string[] doorNames = new[] { "$DorLib2D$00000001", "$DorLib2D$00000002" };

            HashSet<string> keepLayers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "COLUMN",
                "PUB_HATCH",
                "PUB_TEXT",
                "STAIR",
                "WALL",
                "WINDOW",
                "WINDOW_TEXT",
                "0"
            };

            // 1. 让用户用鼠标框选一个矩形
            PromptPointResult ppr1 = ed.GetPoint("\n请指定矩形的第一个角点: ");
            if (ppr1.Status != PromptStatus.OK) return;
            PromptCornerOptions pco = new PromptCornerOptions("\n请指定对角点: ", ppr1.Value);
            PromptPointResult ppr2 = ed.GetCorner(pco);
            if (ppr2.Status != PromptStatus.OK) return;

            // 1. 框选完成后，获取鼠标位置
            System.Drawing.Point mousePos = System.Windows.Forms.Control.MousePosition;

            // 2. 弹出输入框（可用InputBox或自定义窗体）
            string instruction = ShowInputBoxAt(mousePos, "请输入指令：");
            Point3d pt1 = ppr1.Value;
            Point3d pt2 = ppr2.Value;

            // 2. 计算矩形范围
            double minX = Math.Min(pt1.X, pt2.X);
            double minY = Math.Min(pt1.Y, pt2.Y);
            double maxX = Math.Max(pt1.X, pt2.X);
            double maxY = Math.Max(pt1.Y, pt2.Y);
            Extents3d rect = new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));

            var result = new AnalysisResult();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                int count = 0;
                foreach (ObjectId id in modelSpace)
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;
                    if (ent is Line line) continue;

                    // 只处理在 keepLayers 中的图层
                    if (!keepLayers.Contains(ent.Layer))
                        continue;

                    // 判断实体是否有几何范围
                    try
                    {
                        Extents3d ext = ent.GeometricExtents;
                        // 判断实体范围与矩形是否相交
                        if (ext.MinPoint.X < minX || ext.MaxPoint.X > maxX ||
                            ext.MinPoint.Y < minY || ext.MaxPoint.Y > maxY)
                        {
                            continue;
                        }
                    }
                    catch
                    {
                        continue; // 某些实体可能没有几何范围
                    }

                    count++;
                    ent.Highlight();
                    if (ent is BlockReference br)
                    {
                        string blockName = "";
                        if (br.BlockTableRecord.IsValid)
                        {
                            BlockTableRecord brDef = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                            blockName = brDef.Name;
                        }
                        // 统计数量
                        if (!result.BlockCount.ContainsKey(blockName))
                            result.BlockCount[blockName] = 0;
                        result.BlockCount[blockName]++;
                        // 记录中心坐标
                        if (!result.Blocks.ContainsKey(blockName))
                            result.Blocks[blockName] = new List<double[]>();
                        try
                        {
                            var center = GetCenterFromExtents(br.GeometricExtents);
                            result.Blocks[blockName].Add(new double[] { center.X, center.Y, center.Z });
                        }
                        catch { }
                        ed.WriteMessage($"\n类型: BlockReference, 块名: {blockName}, 插入点: {br.Position}, 对象ID: {ent.ObjectId}");
                    }
                    else if (ent is Hatch hatch)
                    {
                        if (hatch.Layer.Equals("PUB_HATCH", StringComparison.OrdinalIgnoreCase))
                        {
                            string blockName = "_FZH";
                            if (!result.BlockCount.ContainsKey(blockName))
                                result.BlockCount[blockName] = 0;
                            result.BlockCount[blockName]++;

                            if (!result.Blocks.ContainsKey(blockName))
                                result.Blocks[blockName] = new List<double[]>();
                            try
                            {
                                var center = GetCenterFromExtents(hatch.GeometricExtents);
                                result.Blocks[blockName].Add(new double[] { center.X, center.Y, center.Z });
                            }
                            catch { }
                        }
                    }
                    else if (ent is DBText dBText)
                    {
                        string content = dBText.TextString?.Trim(); // 去掉首尾空格，防止因空格导致匹配失败

                        if (!string.IsNullOrEmpty(content) &&
                            (content.EndsWith("室") || content.EndsWith("间")))
                        {
                            result.Texts.Add(new
                            {
                                content = content,
                                position = new double[] { dBText.Position.X, dBText.Position.Y, dBText.Position.Z }
                            });
                            ed.WriteMessage($"\n类型: DBText, 内容: \"{content}\", 位置: {dBText.Position}, 对象ID: {ent.ObjectId}");
                        }
                    }
                    else if (ent is Polyline polyline)
                    {
                        string layerName = polyline.Layer;
                        if (!result.PolylineLengthByLayer.ContainsKey(layerName))
                            result.PolylineLengthByLayer[layerName] = 0;
                        result.PolylineLengthByLayer[layerName] += polyline.Length;

                        ed.WriteMessage($"\n类型: Polyline, 图层: {layerName}, 长度: {polyline.Length}, 对象ID: {ent.ObjectId}");
                    }
                    else
                    {
                        ed.WriteMessage($"\n类型: {ent.GetType().Name}, 基点: {ent.GeometricExtents.MinPoint}, 对象ID: {ent.ObjectId}");
                    }
                }
                ed.WriteMessage($"\n共选中 {count} 个图元。");
                tr.Commit();
            }
            // 柱子坐标提取出来，计算最小外接矩形
            if (result.Blocks.ContainsKey("_FZH"))
            {
                var rectPoints = BoundingRectangle.Program
                .CalculateBoundingRectangle(result.Blocks["_FZH"])
                .Select(p => new double[] { p.X, p.Y }).ToList();
                result.RectPoints = rectPoints;
            }

            // 提取门坐标
            foreach (string doorName in doorNames)
            {
                if (result.Blocks.ContainsKey(doorName))
                {
                    foreach (var coords in result.Blocks[doorName])
                    {
                        if (coords.Length >= 2)
                        {
                            double x = coords[0];
                            double y = coords[1];
                            double z = coords.Length >= 3 ? coords[2] : 0;

                            doorPoints.Add(new Point3d(x, y, z));
                        }
                    }
                }
            }
            // 最小外接矩形的坐标变为整数
            var intRectPoints = result.RectPoints
                .Select(pt => pt.Select(x => (int)x).ToArray())
                .ToList();
            // 绘制一个矩形方便查看算法识别的最小外接矩形范围
            DrawRectanglePolyline(intRectPoints);
            // 外接矩形的坐标转为字符串（如 [[32267, 52942], [41142, 52942], ...]）
            string coordinatesStr = "[" + string.Join(", ", intRectPoints.Select(
                                        pt => $"[{string.Join(", ", pt)}]"
                                    )) + "]";
            // 门的坐标转为字符串
            string doorPositionStr = "[" + string.Join(", ", doorPoints.Select(
                                        pt => $"[{(int)pt.X}, {(int)pt.Y}]"
                                    )) + "]";

            var textObj = result.Texts[0];
            // 反射或 dynamic 获取 content 字段
            string roomType = "";
            var prop = textObj.GetType().GetProperty("content");
            if (prop != null)
                roomType = prop.GetValue(textObj)?.ToString();

            string CaculatePromptTemplate = Prompt.CaculatePrompt2;
            // 生成完整Prompt
            string prompt = Prompt.GetLightingPrompt(roomType, coordinatesStr, doorPositionStr, instruction);

            //PrjExploreHelper.InitPalette();
            //PaletteSetDlg dlg = new PaletteSetDlg();
            //PrjExploreHelper.MainPaletteset.Add("test", dlg);
            string reply = await DlgInstance.SendAsync(prompt, instruction);

            //dlg.BeginInvoke((MethodInvoker)(async () =>
            //{
            //    string FinalReply = await dlg.SendAsync(prompt);
            //}));
            //ed.WriteMessage($"reply:{FinalReply}");
            //string finalReply = SendPromptAndWaitReply(dlg, prompt);

            //// 非流式调用模型
            //string reply = Task.Run(() => PaletteSetDlg.CallLLMAsync(prompt)).GetAwaiter().GetResult();
            string Thinking_content = "";

            var match = Regex.Match(reply, @"```json\s*([\s\S]+?)\s*```");

            if (!match.Success)
            {
                ed.WriteMessage("未找到JSON内容");
                return;
            }

            string ModelReplyJson = match.Groups[1].Value;
            Thinking_content = Regex.Replace(reply, @"```json\s*([\s\S]+?)\s*```", "").Trim();

            var obj = JsonConvert.DeserializeObject<LightingDesignResponse>(ModelReplyJson);
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
            string lightName = "";
            if (lightType.Contains("吸顶"))
                lightName = "感应式吸顶灯";
            else if (lightType.Contains("防爆"))
                lightName = "防爆灯";
            else if (lightType.Contains("面板") || lightType.Contains("荧光"))
                lightName = "双管荧光灯";
            else
                lightName = "gen_light";
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

                    InsertBlockFromDwg(new Point3d(x, y, z), LightLayer, lightName, lightRotation);
                    // 统计信息
                }
            }
            AddOrUpdateComponent(lightName, lightCount, obj.lighting_design.power_w.ToString()+"W");
            // 绘制灯具之间的连线
            string wiringLayer = "照明连线";
            double lightLength = 0;
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
                        lightLength += DrawPolyLineBetweenPoints(p1, p2, wiringLayer);
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
                        lightLength += DrawPolyLineBetweenPoints(p1, p2, wiringLayer);
                    }
                }
            }
            lightLength = Math.Round(lightLength / 1000.0, 1);
            AddOrUpdateComponent("照明线路", lightLength, "BV-500 4mm²");
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

                        InsertBlockFromDwg(new Point3d(x, y, z), socketLayer, "三相五孔插座", rotationDeg);
                    }
                }
            }
            AddOrUpdateComponent("三相五孔插座", switch_count, "220V 10A");
            // 插入开关
            if (obj.switch_position?.position_mm != null && obj.switch_position.position_mm.Count >= 2)
            {
                string switchLayer = "开关"; // 替换为你项目中的图层名
                double x = obj.switch_position.position_mm[0];
                double y = obj.switch_position.position_mm[1];
                double z = 0;
                InsertBlockFromDwg(new Point3d(x, y, z), switchLayer, "开关");
            }
            AddOrUpdateComponent("开关", 1, "220V 10A");
            // 最短路径生成线路
            //DrawOpenPolyline(insertPoints, "照明-WIRE");
            // 插入注释
            //InsertLightingInfo()

            // 输出结构化JSON
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(result, Newtonsoft.Json.Formatting.Indented);

            // 导出到文件
            try
            {
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "cad_rect_result.json");
                File.WriteAllText(filePath, json, System.Text.Encoding.UTF8);
                ed.WriteMessage($"\n已导出到: {filePath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n导出JSON文件失败: {ex.Message}");
            }

            // 记录

        }
        private void AddOrUpdateComponent(string type, double count, string info)
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

        public string ShowInputBoxAt(System.Drawing.Point location, string title)
        {
            Form inputForm = new Form();
            inputForm.StartPosition = FormStartPosition.Manual;
            inputForm.Location = location;
            inputForm.Width = 400;
            inputForm.Height = 150;
            inputForm.Text = title;

            TextBox textBox = new TextBox { Dock = DockStyle.Fill, Multiline = true };
            Button okButton = new Button { Text = "确定", Dock = DockStyle.Bottom };
            okButton.Click += (s, e) => inputForm.DialogResult = DialogResult.OK;

            // 支持回车直接提交
            textBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter && !e.Shift)
                {
                    e.SuppressKeyPress = true;
                    inputForm.DialogResult = DialogResult.OK;
                }
            };

            inputForm.Controls.Add(textBox);
            inputForm.Controls.Add(okButton);

            if (inputForm.ShowDialog() == DialogResult.OK)
                return textBox.Text.Trim();
            return "";
        }

        public void InsertBlockFromDwg(Point3d insertPoint, string targetLayer, string blockName, double rotationDegrees = 0)
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            // 构建 DWG 文件路径
            string blocksDirectory = @"C:\Users\武丢丢\Documents\gen_light\";  // ✅ 修改为你实际的路径
            string dwgFilePath = Path.Combine(blocksDirectory, blockName + ".dwg");

            if (!File.Exists(dwgFilePath))
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog($"找不到块文件：{dwgFilePath}");
                return;
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
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    tr.Commit();
                }
            }
        }
        [CommandMethod("test", CommandFlags.Session)]
        public void InsertBlockFromDwgTest()
        {
            Point3d insertPoint = new Point3d(34640, 55259, 0);
            string targetLayer = "AI";
            double rotationDegrees = 90;
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\gen_light\双管荧光灯.dwg"; // 固定块文件路径
            string blockName = "AI_Gen_Light"; // 块名

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
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    tr.Commit();
                }
            }
        }


        public static List<Point3d> SortPointsToClosedLoop(List<Point3d> inputPoints)
        {
            if (inputPoints == null || inputPoints.Count == 0)
                return new List<Point3d>();

            List<Point3d> sorted = new List<Point3d>();
            HashSet<int> visited = new HashSet<int>();

            Point3d current = inputPoints[0];
            sorted.Add(current);
            visited.Add(0);

            while (visited.Count < inputPoints.Count)
            {
                double minDist = double.MaxValue;
                int nearestIndex = -1;

                for (int i = 0; i < inputPoints.Count; i++)
                {
                    if (visited.Contains(i)) continue;

                    double dist = current.DistanceTo(inputPoints[i]);
                    if (dist < minDist)
                    {
                        minDist = dist;
                        nearestIndex = i;
                    }
                }

                if (nearestIndex != -1)
                {
                    current = inputPoints[nearestIndex];
                    sorted.Add(current);
                    visited.Add(nearestIndex);
                }
            }

            // 闭合回原点
            sorted.Add(sorted[0]);

            return sorted;
        }

        public void DrawClosedPolyline(List<Point3d> lampPoints, string layerName)
        {
            var sortedPoints = SortPointsToClosedLoop(lampPoints);

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (DocumentLock docLock = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 创建图层（如果不存在）
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(layerName))
                {
                    lt.UpgradeOpen();
                    LayerTableRecord newLayer = new LayerTableRecord { Name = layerName };
                    lt.Add(newLayer);
                    tr.AddNewlyCreatedDBObject(newLayer, true);
                }

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                Polyline poly = new Polyline();
                for (int i = 0; i < sortedPoints.Count; i++)
                {
                    poly.AddVertexAt(i, new Point2d(sortedPoints[i].X, sortedPoints[i].Y), 0, 0, 0);
                }

                poly.Closed = true;  // ✅ 关闭曲线

                poly.Layer = layerName;

                poly.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 153, 153);
                poly.ConstantWidth = 35f;

                btr.AppendEntity(poly);
                tr.AddNewlyCreatedDBObject(poly, true);

                tr.Commit();
            }
        }
        public static List<Point3d> SortPointsNearestPath(List<Point3d> inputPoints)
        {
            if (inputPoints == null || inputPoints.Count == 0)
                return new List<Point3d>();

            List<Point3d> sorted = new List<Point3d>();
            HashSet<int> visited = new HashSet<int>();

            Point3d current = inputPoints[0];
            sorted.Add(current);
            visited.Add(0);

            while (visited.Count < inputPoints.Count)
            {
                double minDist = double.MaxValue;
                int nearestIndex = -1;

                for (int i = 0; i < inputPoints.Count; i++)
                {
                    if (visited.Contains(i)) continue;

                    double dist = current.DistanceTo(inputPoints[i]);
                    if (dist < minDist)
                    {
                        minDist = dist;
                        nearestIndex = i;
                    }
                }

                if (nearestIndex != -1)
                {
                    current = inputPoints[nearestIndex];
                    sorted.Add(current);
                    visited.Add(nearestIndex);
                }
            }

            return sorted; // ❌ 不再加 sorted[0]
        }

        public void DrawOpenPolyline(List<Point3d> lampPoints, string layerName)
        {
            var sortedPoints = SortPointsNearestPath(lampPoints);

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (DocumentLock docLock = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(layerName))
                {
                    lt.UpgradeOpen();
                    LayerTableRecord newLayer = new LayerTableRecord { Name = layerName };
                    lt.Add(newLayer);
                    tr.AddNewlyCreatedDBObject(newLayer, true);
                }

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                Polyline poly = new Polyline();
                for (int i = 0; i < sortedPoints.Count; i++)
                {
                    poly.AddVertexAt(i, new Point2d(sortedPoints[i].X, sortedPoints[i].Y), 0, 0, 0);
                }

                // ❌ 不闭合 poly.Closed = true;
                poly.Layer = layerName;
                poly.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 153, 153);
                poly.ConstantWidth = 35f;

                btr.AppendEntity(poly);
                tr.AddNewlyCreatedDBObject(poly, true);

                tr.Commit();
            }
        }

        public void DrawRectanglePolyline(List<int[]> rectPoints, string layerName = "Rect")
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
        public void ExportStatisticsToExcel(List<(string Type, int Count, string Info)> stats, string filePath)
        {
            // 创建工作簿
            IWorkbook workbook = new XSSFWorkbook(); // .xlsx 格式
            ISheet sheet = workbook.CreateSheet("元件统计");

            // 创建表头
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("元件类型");
            headerRow.CreateCell(1).SetCellValue("个数");
            headerRow.CreateCell(2).SetCellValue("元件信息");

            // 填充数据
            for (int i = 0; i < stats.Count; i++)
            {
                var item = stats[i];
                IRow row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(item.Type);
                row.CreateCell(1).SetCellValue(item.Count);
                row.CreateCell(2).SetCellValue(item.Info);
            }

            // 自动列宽（可选）
            for (int i = 0; i < 3; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            // 保存到文件
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }

        public void DrawLineBetweenPoints(Point3d pt1, Point3d pt2, string layerName)
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
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

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // 创建直线
                Line line = new Line(pt1, pt2)
                {
                    Layer = layerName,
                    Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 153, 204), // 粉红色
                    LineWeight = LineWeight.LineWeight211 // 0.50mm线宽
                };

                modelSpace.AppendEntity(line);
                tr.AddNewlyCreatedDBObject(line, true);

                tr.Commit();
            }
        }
        public int DrawPolyLineBetweenPoints(Point3d pt1, Point3d pt2, string layerName)
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            int polylineLength = 0;

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
                poly.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 35f, 35f); // 线宽35f
                poly.AddVertexAt(1, new Point2d(pt2.X, pt2.Y), 0, 35f, 35f); // 线宽35f
                poly.Layer = layerName;
                poly.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2); // 黄色
                poly.ConstantWidth = 35f;

                modelSpace.AppendEntity(poly);
                tr.AddNewlyCreatedDBObject(poly, true);
                // 获取多段线长度
                polylineLength = (int)poly.Length;
                tr.Commit();
            }
            return polylineLength;
        }
        [CommandMethod("MM", CommandFlags.Session)]
        public void TestMarkdig()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            var result = Markdown.ToHtml("| 灯具类型 | 个数 | 瓦数 | \r\n|---------|-----|-----|\r\n | 吸顶灯 | 2 | 20W | | 防爆灯 | 3 | 50W | | 荧光灯 | 5 | 60W |", pipeline);
            ed.WriteMessage($"{result}");   // prints: <p>This is a text with some <em>emphasis</em></p>
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
        public void InsertLightingInfo(Point3d basePosition, int fixtureCount, int powerW, int mountingHeight, double textHeight = 300, string layerName = "PUB_TEXT")
        {
            //Point3d basePosition = new Point3d(34640, 55259, 0);
            //int fixtureCount = 4;
            //int powerW = 100;
            //double mountingHeight = 4.5;
            //int textHeight = 300;
            //string layerName = "PUB_TEXT";
            // 插入中间的功率文本（如 "100W"）
            InsertTextAt($"{powerW}W", basePosition, layerName, textHeight);

            // 插入下方的安装高度文本（如 "4.5"）
            var below = new Point3d(basePosition.X, basePosition.Y - textHeight * 1.4, basePosition.Z);
            InsertTextAt($"{mountingHeight:0.##}", below, layerName, textHeight);

            // 插入左侧的灯具数量文本（如 "4"）
            var left = new Point3d(basePosition.X - textHeight * 1.4, basePosition.Y - textHeight * 0.8, basePosition.Z);
            InsertTextAt($"{fixtureCount}", left, layerName, textHeight);

            // 计算直线起止点（在瓦数和高度之间，略微留白）
            double halfLineLength = textHeight * 2.0;
            var lineY = basePosition.Y - 85; // 在两者中间
            Point3d lineStart = new Point3d(basePosition.X - halfLineLength + 500, lineY, basePosition.Z);
            Point3d lineEnd = new Point3d(basePosition.X + halfLineLength +500, lineY, basePosition.Z);

            // 绘制黄色横线
            DrawYellowLine(lineStart, lineEnd, layerName);
        }
        /// <summary>
        /// 在指定坐标插入单行文本（DBText）
        /// </summary>
        /// <param name="text">要插入的文本内容</param>
        /// <param name="position">插入点坐标</param>
        /// <param name="layerName">可选，插入到的图层，默认为当前图层</param>
        public void InsertTextAt(string text, Point3d position, string layerName = null, double textHeight = 300)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

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

                modelSpace.AppendEntity(mtext);
                tr.AddNewlyCreatedDBObject(mtext, true);

                tr.Commit();
            }
        }
        public void DrawYellowLine(Point3d start, Point3d end, string layerName)
        {
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

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                Line line = new Line(start, end)
                {
                    Layer = layerName,
                    Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2) // 2 = 黄色
                };

                modelSpace.AppendEntity(line);
                tr.AddNewlyCreatedDBObject(line, true);

                tr.Commit();
            }
        }
        public void DrawComponentTable6ColsWithBlocks(List<(string Type, double Count, string Info)> statsList)
        {
            double rowHeight = 800;
            short colorIndex = 7;
            LineWeight lineWeight = LineWeight.LineWeight200;
            Point3d insertPoint = new Point3d(3599, 17119, 0);
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
                double[] colWidths = { 900, 4000, 4000, 900, 900, 1500 };
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
                            double blockY = insertPoint.Y - (i+2) * rowHeight - rowHeight / 2;
                            var blockPoint = new Point3d(blockX, blockY, 0);
                            InsertBlockFromDwg(blockPoint, "PUB_TEXT", type, 0);
                            continue;
                        }
                        double x = colX[j] + 100; // 边距内缩
                        double y = insertPoint.Y - (i+2) * rowHeight - rowHeight * 0.6;

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
        #endregion
    }
    }
