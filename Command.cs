using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
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
        public static PaletteSetDlg DlgInstance;
        // 存储上一次插入的图元
        //List<ObjectId> lastInsertedEntities = new List<ObjectId>();

        #endregion

        #region 初始化
        public void Initialize()
        {
            //PrjExploreHelper.InitMenu();
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
        [CommandMethod("ALIYUN_QWEN_CHAT", CommandFlags.Session)]
        public async void AliyunQwenChat()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            string apiKey = Environment.GetEnvironmentVariable("DASHSCOPE_API_KEY");
            if (string.IsNullOrEmpty(apiKey))
            {
                ed.WriteMessage("\nAPI Key 未设置。请确保环境变量 'DASHSCOPE_API_KEY' 已设置。");
                return;
            }

            string url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions";
            string jsonContent = @"{
                ""model"": ""qwen-plus"",
                ""messages"": [
                    {
                        ""role"": ""system"",
                        ""content"": ""You are a helpful assistant.""
                    },
                    {
                        ""role"": ""user"", 
                        ""content"": ""你是谁？""
                    }
                ]
            }";

            string result = await SendPostRequestAsync(url, jsonContent, apiKey);
            ed.WriteMessage($"\n模型输出: {result}");
        }

        private static readonly HttpClient httpClient = new HttpClient();

        private static async Task<string> SendPostRequestAsync(string url, string jsonContent, string apiKey)
        {
            using (var content = new StringContent(jsonContent, Encoding.UTF8, "application/json"))
            {
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
                httpClient.DefaultRequestHeaders.Accept.Clear();
                httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                HttpResponseMessage response = await httpClient.PostAsync(url, content);

                if (response.IsSuccessStatusCode)
                {
                    return await response.Content.ReadAsStringAsync();
                }
                else
                {
                    return $"请求失败: {response.StatusCode}";
                }
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
            public List<List<List<double>>> socket_wiring_lines_mm { get; set; }
            public SwitchPosition switch_position { get; set; }

            public class RoomInfo
            {
                public Dictionary<string, List<double[]>> Blocks { get; set; } = new Dictionary<string, List<double[]>>();
                public Dictionary<string, int> BlockCount { get; set; } = new Dictionary<string, int>();
                public Dictionary<string, double> PolylineLengthByLayer { get; set; } = new Dictionary<string, double>();
                public List<object> Texts { get; set; } = new List<object>();
                public List<double[]> RectPoints { get; set; } = new List<double[]>();
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

        /// <summary>
        /// 最重要的命令，用户运行这个命令选定一个区域，程序会识别区域内指定的元件的坐标信息，嵌入到prompt模板中
        /// 同时
        /// </summary>
        [CommandMethod("SELECT_RECT_PRINT", CommandFlags.Session)]
        public async void SelectEntitiesByRectangleAndPrintInfo()
        {
            // 开始记录新一轮插入的图元
            InsertTracker.BeginNewInsert();

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
                            var center = CadDrawingHelper.GetCenterFromExtents(br.GeometricExtents);
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
                                var center = CadDrawingHelper.GetCenterFromExtents(hatch.GeometricExtents);
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
            // 给矩形加一个偏差，缩小因为识别柱子坐标带来的内轮廓误差
            var shrunkPoints = CadDrawingHelper.ShrinkRectangle(intRectPoints, 125, 125);
            // 绘制一个矩形方便查看算法识别的最小外接矩形范围
            //DrawRectanglePolyline(shrunkPoints);
            // 外接矩形的坐标转为字符串（如 [[32267, 52942], [41142, 52942], ...]）
            string coordinatesStr = "[" + string.Join(", ", shrunkPoints.Select(
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
            //RoomInputCache.SetRoomDrawingInputs(roomType, coordinatesStr, doorPositionStr);
            InsertTracker.SetRoomDrawingInputs(roomType, coordinatesStr, doorPositionStr);
            string CaculatePromptTemplate = Prompt.CaculatePrompt2;
            // 生成完整Prompt
            string prompt = Prompt.GetPreparePrompt(roomType, coordinatesStr, doorPositionStr);
            InsertTracker.WaitingRoomPromptInfo = prompt;
            await DlgInstance.AppendMessageAsync("AI", "**已识别房间信息，请继续输入您的设计指令，如无则回复无**");

            //PrjExploreHelper.InitPalette();
            //PaletteSetDlg dlg = new PaletteSetDlg();
            //PrjExploreHelper.MainPaletteset.Add("test", dlg);

            //dlg.BeginInvoke((MethodInvoker)(async () =>
            //{
            //    string FinalReply = await dlg.SendAsync(prompt);
            //}));
            //ed.WriteMessage($"reply:{FinalReply}");
            //string finalReply = SendPromptAndWaitReply(dlg, prompt);

            //// 非流式调用模型
            //string reply = Task.Run(() => PaletteSetDlg.CallLLMAsync(prompt)).GetAwaiter().GetResult();
            // 输出结构化JSON
            //string reply = await DlgInstance.SendAsync(prompt, "");
            //var match = Regex.Match(reply, @"```json\s*([\s\S]+?)\s*```");

            //if (!match.Success)
            //{
            //    ed.WriteMessage("未找到JSON内容");
            //    return;
            //}

            //string ModelReplyJson = match.Groups[1].Value;
            //InsertLightingFromModelReply(ModelReplyJson);
            //string json = Newtonsoft.Json.JsonConvert.SerializeObject(result, Newtonsoft.Json.Formatting.Indented);

            // 导出到文件
            //try
            //{
            //    string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "cad_rect_result.json");
            //    File.WriteAllText(filePath, json, System.Text.Encoding.UTF8);
            //    ed.WriteMessage($"\n已导出到: {filePath}");
            //}
            //catch (System.Exception ex)
            //{
            //    ed.WriteMessage($"\n导出JSON文件失败: {ex.Message}");
            //}
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
        #endregion
    }
    }
