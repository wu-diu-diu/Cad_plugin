using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using DocumentFormat.OpenXml.Presentation;
using FunctionCallingAI;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Microsoft.IO.RecyclableMemoryStreamManager;
using CADApplication = Autodesk.AutoCAD.ApplicationServices.Application;


[assembly: ExtensionApplication(typeof(CoDesignStudy.Cad.PlugIn.Command))]
[assembly: CommandClass(typeof(CoDesignStudy.Cad.PlugIn.Command))]
namespace CoDesignStudy.Cad.PlugIn
{
    /// <summary>
    /// 入口
    /// </summary>
    // 定义数据模型类
    public class CADResponse
    {
        public bool success { get; set; }
        public string message { get; set; }
        public int total_images { get; set; }
        public int processed_images { get; set; }
        public Dictionary<string, List<RoomData>> results { get; set; }
        public Dictionary<string, string> errors { get; set; }
    }

    public class RoomData
    {
        public string room_name { get; set; }
        public List<List<List<double>>> cad_coordinates { get; set; }
    }

    public class Command : IExtensionApplication
    {
        #region 成员变量
        public static PaletteSetDlg DlgInstance;
        // 存储上一次插入的图元
        //List<ObjectId> lastInsertedEntities = new List<ObjectId>();
        Dictionary<string, double> Cadparam = new Dictionary<string, double>();
        public static string ImagePath = "";
        public static string serverUrl = "http://127.0.0.1:8000";

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
        [CommandMethod("QWENSDK", CommandFlags.Session)]
        public static async void Qwen()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            
            try
            {
                string apiKey = Environment.GetEnvironmentVariable("DASHSCOPE_API_KEY");
                if (string.IsNullOrEmpty(apiKey))
                {
                    ed.WriteMessage("\nAPI Key 未设置。请确保环境变量 'DASHSCOPE_API_KEY' 已设置。");
                    return;
                }

                var client = new QwenClient("qwen-plus", apiKey);
                
                // 初始化对话历史，包含系统消息
                List<ChatMessage> messages = new List<ChatMessage>
                {
                    new SystemChatMessage("You are a helpful assistant for AutoCAD users. You can help with CAD operations, drawing commands, and general questions.")
                };

                ed.WriteMessage("\n=== QwenSDK 对话模式已启动 ===");
                ed.WriteMessage("\n提示：输入消息进行对话，输入 'exit' 或 'quit' 退出");
                ed.WriteMessage("\n" + new string('-', 50));

                while (true)
                {
                    // 获取用户输入
                    PromptStringOptions pso = new PromptStringOptions("\n用户: ")
                    {
                        AllowSpaces = true
                    };
                    PromptResult pr = ed.GetString(pso);
                    
                    if (pr.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\n对话已取消。");
                        break;
                    }

                    string userInput = pr.StringResult.Trim();
                    
                    // 检查退出命令
                    if (string.IsNullOrEmpty(userInput))
                    {
                        continue;
                    }
                    
                    if (userInput.ToLower() == "exit" || userInput.ToLower() == "quit")
                    {
                        ed.WriteMessage("\n=== 对话结束 ===");
                        break;
                    }
                    
                    if (userInput.ToLower() == "clear")
                    {
                        // 清除对话历史，只保留系统消息
                        messages.Clear();
                        messages.Add(new SystemChatMessage("You are a helpful assistant for AutoCAD users. You can help with CAD operations, drawing commands, and general questions."));
                        ed.WriteMessage("\n对话历史已清除。");
                        continue;
                    }

                    // 添加用户消息到历史
                    messages.Add(new UserChatMessage(userInput));
                    
                    try
                    {
                        ed.WriteMessage("\nAI正在思考...");
                        
                        // 发送请求到AI
                        var result = await client.CompleteChatAsync(messages);
                        
                        // 显示AI回复
                        ed.WriteMessage($"\nAI: {result.Content}");
                        
                        // 添加AI回复到历史
                        messages.Add(new AssistantChatMessage(result.Content));
                        
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n请求失败: {ex.Message}");
                        ed.WriteMessage("\n请重试或检查网络连接。");
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n启动对话模式失败: {ex.Message}");
            }
        }
        
        [CommandMethod("QWENSTREAM", CommandFlags.Session)]
        public static async void QwenStream()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            
            try
            {
                string apiKey = Environment.GetEnvironmentVariable("DASHSCOPE_API_KEY");
                if (string.IsNullOrEmpty(apiKey))
                {
                    ed.WriteMessage("\nAPI Key 未设置。请确保环境变量 'DASHSCOPE_API_KEY' 已设置。");
                    return;
                }

                var client = new QwenClient("qwen-plus", apiKey);
                
                List<ChatMessage> messages = new List<ChatMessage>
                {
                    new SystemChatMessage("You are a helpful assistant."),
                    new UserChatMessage("告诉我CAD中模型空间和布局空间的区别")
                };

                ed.WriteMessage("\n=== 流式输出测试开始 ===");
                ed.WriteMessage("\n问题：告诉我CAD中模型空间和布局空间的区别");
                ed.WriteMessage("\n" + new string('-', 40));
                ed.WriteMessage("\nAI回复: ");

                // 用于累积完整回复的变量
                string fullResponse = "";
                
                // 定义流式回调处理函数
                StreamingResponseHandler streamHandler = (StreamingChatCompletionUpdate update) =>
                {
                    if (!string.IsNullOrEmpty(update.Content))
                    {
                        // 实时输出每个片段
                        ed.WriteMessage(update.Content);
                        fullResponse += update.Content;
                    }
                    
                    // 检查是否完成
                    if (update.IsFinished)
                    {
                        ed.WriteMessage("\n" + new string('-', 40));
                        ed.WriteMessage($"\n流式输出完成！完整回复长度: {fullResponse.Length} 字符");
                        ed.WriteMessage($"\n结束原因: {update.FinishReason}");
                    }
                };

                // 发送流式请求
                await client.CompleteChatStreamingAsync(messages, streamHandler);
                
                ed.WriteMessage("\n=== 流式输出测试结束 ===");
                
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n流式请求失败: {ex.Message}");
                ed.WriteMessage($"\n错误详情: {ex.StackTrace}");
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

        [CommandMethod("PLOT_DISPLAY", CommandFlags.Session)]
        public void PlotToPng()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            string outputDir = @"C:\Users\武丢丢\Documents\cadpdf";
            string outputFile = Path.Combine(outputDir, "output.png");

            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            using (doc.LockDocument()) // 添加文档锁定
            using (Transaction tr = db.TransactionManager.StartTransaction())
            using (PlotEngine pe = PlotFactory.CreatePublishEngine())
            using (PlotProgressDialog ppd = new PlotProgressDialog(false, 1, true))
            {
                try
                {
                    // 获取布局和创建 PlotSettings
                    LayoutManager lm = LayoutManager.Current;
                    ObjectId layoutId = lm.GetLayoutId(lm.CurrentLayout);
                    Layout layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);

                    PlotSettings ps = new PlotSettings(layout.ModelType);
                    ps.CopyFrom(layout);

                    // 设置打印参数
                    PlotSettingsValidator psv = PlotSettingsValidator.Current;
                    psv.SetPlotConfigurationName(ps, "PublishToWeb PNG.pc3", null);
                    psv.RefreshLists(ps);

                    string customMediaName = "UserDefinedRaster (3840.00 x 2160.00像素)";
                    var mediaList = psv.GetCanonicalMediaNameList(ps);
                    if (mediaList.Contains(customMediaName))
                        psv.SetCanonicalMediaName(ps, customMediaName);

                    psv.SetUseStandardScale(ps, true);
                    psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
                    psv.SetPlotRotation(ps, PlotRotation.Degrees000);
                    psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Display);

                    // 创建 PlotInfo 并验证
                    PlotInfo pi = new PlotInfo { Layout = layoutId, OverrideSettings = ps };
                    PlotInfoValidator piv = new PlotInfoValidator { MediaMatchingPolicy = MatchingPolicy.MatchEnabled };
                    piv.Validate(pi);

                    // 执行打印
                    System.IO.Directory.SetCurrentDirectory(outputDir);
                    ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "出图进度");
                    ppd.OnBeginPlot();

                    pe.BeginPlot(ppd, null);
                    pe.BeginDocument(pi, doc.Name, null, 1, true, outputFile);

                    PlotPageInfo ppi = new PlotPageInfo();
                    pe.BeginPage(ppi, pi, true, null);
                    pe.BeginGenerateGraphics(null);
                    pe.EndGenerateGraphics(null);
                    pe.EndPage(null);
                    pe.EndDocument(null);
                    pe.EndPlot(null);

                    tr.Commit();
                    ed.WriteMessage($"\nPNG 图片已输出到: {outputFile}");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n出图失败: {ex.Message}");
                    tr.Abort();
                }
            }
        }
        [CommandMethod("PLOT_WINDOW", CommandFlags.Session)]
        public void PlotWindowPng()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            string outputDir = @"C:\Users\武丢丢\Documents\upload_test";
            string outputFile = Path.Combine(outputDir, "output.png");
            ImagePath = outputFile;
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // 定义需要打印的图层名称列表
            List<string> remainedLyaer = new List<string>
            {
                "PUB_WALL",
                "PUB_TEXT",
                "PUB_HATCH",
                "WINDOW",
                "COLUMN"
                // 在这里添加你需要打印的图层名称
            };

            // 1. 切换到布局1
            LayoutManager lm = LayoutManager.Current;
            if (lm.CurrentLayout != "布局1")
            {
                try
                {
                    lm.CurrentLayout = "布局1";
                    ed.WriteMessage("\n已切换到布局1");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n切换到布局1失败: {ex.Message}");
                    return;
                }
            }

            // 2. 让用户用鼠标框选一个窗口区域
            PromptPointResult ppr1 = ed.GetPoint("\n请指定窗口的第一个角点: ");
            if (ppr1.Status != PromptStatus.OK) return;
            PromptCornerOptions pco = new PromptCornerOptions("\n请指定对角点: ", ppr1.Value);
            PromptPointResult ppr2 = ed.GetCorner(pco);
            if (ppr2.Status != PromptStatus.OK) return;
            Point3d pt1 = ppr1.Value;
            Point3d pt2 = ppr2.Value;

            // 3. 计算窗口的四个角点
            double xmin = Math.Min(pt1.X, pt2.X);
            double xmax = Math.Max(pt1.X, pt2.X);
            double ymin = Math.Min(pt1.Y, pt2.Y);
            double ymax = Math.Max(pt1.Y, pt2.Y);
            Cadparam["xmin"] = xmin;
            Cadparam["xmax"] = xmax;
            Cadparam["ymin"] = ymin;
            Cadparam["ymax"] = ymax;
            Cadparam["originx"] = -13.0;
            Cadparam["originy"] = 11.0;
            Point3d leftBottom = new Point3d(xmin, ymin, 0);
            Point3d rightBottom = new Point3d(xmax, ymin, 0);
            Point3d rightTop = new Point3d(xmax, ymax, 0);
            Point3d leftTop = new Point3d(xmin, ymax, 0);
            double originx = -13.0;
            double originy = 11.0;

            // 4. 打印窗口四个点坐标和原点偏移量
            ed.WriteMessage($"\n窗口四个角点坐标：");
            ed.WriteMessage($"\n左下: ({leftBottom.X}, {leftBottom.Y})");
            ed.WriteMessage($"\n右下: ({rightBottom.X}, {rightBottom.Y})");
            ed.WriteMessage($"\n右上: ({rightTop.X}, {rightTop.Y})");
            ed.WriteMessage($"\n左上: ({leftTop.X}, {leftTop.Y})");
            ed.WriteMessage($"\n窗口范围：xmin={xmin}, xmax={xmax}, ymin={ymin}, ymax={ymax}");
            ed.WriteMessage($"\n窗口左下角（原点）偏移量：({originx}, {originy})");

            // 5. 记录所有图层的原始状态

            Dictionary<string, bool> originalLayerStates = new Dictionary<string, bool>();
            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId layerId in lt)
                {
                    LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    originalLayerStates[ltr.Name] = ltr.IsOff;
                    // 只保留指定的图层可见
                    ltr.IsOff = !remainedLyaer.Contains(ltr.Name);
                }
                tr.Commit();
                ed.WriteMessage($"\n已设置图层可见性，只显示 {remainedLyaer.Count} 个指定图层");
            }

            try
            {
                // 6. 执行出图
                using (doc.LockDocument())
                using (Transaction tr = db.TransactionManager.StartTransaction())
                using (PlotEngine pe = PlotFactory.CreatePublishEngine())
                using (PlotProgressDialog ppd = new PlotProgressDialog(false, 1, true))
                {
                    try
                    {
                        ObjectId layoutId = lm.GetLayoutId("布局1");
                        Layout layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);

                        PlotSettings ps = new PlotSettings(layout.ModelType);
                        ps.CopyFrom(layout);
                        PlotSettingsValidator psv = PlotSettingsValidator.Current;
                        psv.SetPlotConfigurationName(ps, "PublishToWeb PNG.pc3", null);
                        psv.RefreshLists(ps);

                        string customMediaName = "UserDefinedRaster (3840.00 x 2160.00像素)";
                        var mediaList = psv.GetCanonicalMediaNameList(ps);
                        if (mediaList.Contains(customMediaName))
                            psv.SetCanonicalMediaName(ps, customMediaName);

                        // 7. 设置窗口区域出图
                        try
                        {
                            // 先设置窗口区域
                            psv.SetPlotWindowArea(ps, new Extents2d(xmin, ymin, xmax, ymax));
                            // 再设置为窗口类型
                            psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception)
                        {
                            // 如果窗口类型设置失败，回退到范围类型
                            ed.WriteMessage("\n窗口类型设置失败，使用范围类型出图");
                            psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
                        }

                        psv.SetUseStandardScale(ps, true);
                        psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
                        psv.SetPlotRotation(ps, PlotRotation.Degrees090);

                        PlotInfo pi = new PlotInfo { Layout = layoutId, OverrideSettings = ps };
                        PlotInfoValidator piv = new PlotInfoValidator { MediaMatchingPolicy = MatchingPolicy.MatchEnabled };
                        piv.Validate(pi);

                        System.IO.Directory.SetCurrentDirectory(outputDir);
                        ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "出图进度");
                        ppd.OnBeginPlot();
                        pe.BeginPlot(ppd, null);
                        pe.BeginDocument(pi, doc.Name, null, 1, true, outputFile);

                        PlotPageInfo ppi = new PlotPageInfo();
                        pe.BeginPage(ppi, pi, true, null);
                        pe.BeginGenerateGraphics(null);
                        pe.EndGenerateGraphics(null);
                        pe.EndPage(null);
                        pe.EndDocument(null);
                        pe.EndPlot(null);

                        tr.Commit();
                        ed.WriteMessage($"\nPNG 图片已输出到: {outputFile}");

                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n出图失败: {ex.Message}");
                        tr.Abort();
                        throw; // 重新抛出异常以便在finally中处理
                    }
                }
            }
            finally
            {
                // 8. 恢复图层状态
                try
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                        foreach (ObjectId layerId in lt)
                        {
                            LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                            if (originalLayerStates.ContainsKey(ltr.Name))
                                ltr.IsOff = originalLayerStates[ltr.Name];
                        }
                        tr.Commit();
                        ed.WriteMessage("\n已恢复图层原始状态");
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n恢复图层状态失败: {ex.Message}");
                }
            }
        }
        [CommandMethod("LIST_PLOT_PAPERS", CommandFlags.Session)]
        public void ListPlotPapers()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayoutManager lm = LayoutManager.Current;
                ObjectId layoutId = lm.GetLayoutId(lm.CurrentLayout);
                Layout layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);

                PlotSettings ps = new PlotSettings(layout.ModelType);
                ps.CopyFrom(layout);

                PlotSettingsValidator psv = PlotSettingsValidator.Current;

                // 设置打印机名称
                string printerName = "PublishToWeb PNG.pc3"; // 可替换为你需要的打印机
                psv.SetPlotConfigurationName(ps, printerName, null);
                psv.RefreshLists(ps);

                // 获取所有纸张尺寸
                var mediaNames = psv.GetCanonicalMediaNameList(ps);

                ed.WriteMessage($"\n打印机: {printerName} 支持的纸张尺寸：");
                foreach (string name in mediaNames)
                {
                    ed.WriteMessage($"\n{name}");
                }
                tr.Commit();
            }
        }
        [CommandMethod("UPLOAD_CAD", CommandFlags.Session)]
        public async void UploadCADCommand()
        {
            var doc = CADApplication.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            if (!File.Exists(ImagePath))
            {
                ed.WriteMessage($"\n图片文件不存在: {ImagePath}");
                return;
            }

            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromMinutes(5); // 设置5分钟超时

                    using (var form = new MultipartFormDataContent())
                    {
                        byte[] fileBytes = File.ReadAllBytes(ImagePath);
                        var fileContent = new ByteArrayContent(fileBytes);
                        fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("image/png");
                        form.Add(fileContent, "files", Path.GetFileName(ImagePath));

                        // 构造参数JSON
                        string paramJson = $@"{{
                            ""Xmin"": {Cadparam["xmin"]},
                            ""Ymin"": {Cadparam["ymin"]},
                            ""Xmax"": {Cadparam["xmax"]},
                            ""Ymax"": {Cadparam["ymax"]},
                            ""originx"": {Cadparam["originx"]},
                            ""originy"": {Cadparam["originy"]}
                        }}";


                        form.Add(new StringContent(paramJson, Encoding.UTF8), "cad_params");


                        // 发送POST请求到正确的端点
                        string uploadUrl = $"{serverUrl}/upload-and-process";
                        var response = await client.PostAsync(uploadUrl, form);
                        string result = await response.Content.ReadAsStringAsync();

                        if (response.IsSuccessStatusCode)
                        {
                            ed.WriteMessage($"\n✅ 图片和参数已成功发送到服务器");

                            // 使用强类型解析
                            try
                            {
                                var cadResponse = JsonConvert.DeserializeObject<CADResponse>(result);

                                ed.WriteMessage($"\n📊 处理结果: {cadResponse.processed_images}/{cadResponse.total_images} 个图像成功");

                                if (cadResponse.success && cadResponse.results != null)
                                {
                                    foreach (var imageResult in cadResponse.results)
                                    {
                                        string imageName = imageResult.Key;
                                        List<RoomData> rooms = imageResult.Value;

                                        ed.WriteMessage($"\n图片: {imageName}");
                                        ed.WriteMessage($"发现 {rooms.Count} 个房间:");

                                        foreach (var room in rooms)
                                        {
                                            ed.WriteMessage($"  - {room.room_name}: {room.cad_coordinates[0].Count} 个坐标点");

                                            // CreateRoomInCAD(room.room_name, room.cad_coordinates);
                                        }
                                    }
                                }
                            }
                            catch (System.Exception parseEx)
                            {
                                ed.WriteMessage($"\n⚠️ 解析返回结果失败: {parseEx.Message}");
                            }
                        }
                        else
                        {
                            ed.WriteMessage($"\n❌ 发送失败，状态码：{response.StatusCode}");
                            ed.WriteMessage($"\n错误详情：{result}");
                        }
                    }
                }
            }
            catch (HttpRequestException httpEx)
            {
                ed.WriteMessage($"\n❌ 网络请求失败: {httpEx.Message}");
                ed.WriteMessage("\n请检查服务器地址和网络连接");
            }
            catch (TaskCanceledException timeoutEx)
            {
                ed.WriteMessage($"\n❌ 请求超时: {timeoutEx.Message}");
                ed.WriteMessage("\n图片处理可能需要更长时间，请稍后重试");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n❌ 发生异常: {ex.Message}");
                ed.WriteMessage($"\n详细错误: {ex.StackTrace}");
            }
        }
        [CommandMethod("TEST_NETWORK", CommandFlags.Session)]
        public static async void TestNetwork()
        {
            var ed = CADApplication.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(10);
                    var response = await client.GetAsync("https://www.baidu.com");
                    ed.WriteMessage($"\n网络测试成功，状态码: {response.StatusCode}");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n网络测试失败: {ex.Message}");
            }
        }
        #endregion
    }
}
