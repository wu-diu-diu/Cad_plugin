using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using BoundingRectangle;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
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
                PrjExploreHelper.InitPalette();
                PaletteSetDlg dlg = new PaletteSetDlg();
                PrjExploreHelper.MainPaletteset.Add("测试界面", dlg);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("测试", ex.Message);
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
        private void DrawGreenRectangleByPoints(List<double[]> rectPoints, Database db)
        {
            if (rectPoints == null || rectPoints.Count != 4) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                for (int i = 0; i < 4; i++)
                {
                    double[] start = rectPoints[i];
                    double[] end = rectPoints[(i + 1) % 4];
                    Line line = new Line(
                        new Point3d(start[0], start[1], 0),
                        new Point3d(end[0], end[1], 0)
                    );
                    line.ColorIndex = 3; // 绿色
                    modelSpace.AppendEntity(line);
                    tr.AddNewlyCreatedDBObject(line, true);
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

        [CommandMethod("SELECT_RECT_PRINT", CommandFlags.Session)]
        public void SelectEntitiesByRectangleAndPrintInfo()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

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
            if (result.Blocks.ContainsKey("_FZH"))
            {
                var rectPoints = BoundingRectangle.Program
                .CalculateBoundingRectangle(result.Blocks["_FZH"])
                .Select(p => new double[] { p.X, p.Y }).ToList();
                result.RectPoints = rectPoints;
            }
            var intRectPoints = result.RectPoints
                .Select(pt => pt.Select(x => (int)x).ToArray())
                .ToList();
            // 转为字符串（如 [[32267, 52942], [41142, 52942], ...]）
            string coordinatesStr = "[" + string.Join(", ", intRectPoints.Select(
                                        pt => $"[{string.Join(", ", pt)}]"
                                    )) + "]";

            var textObj = result.Texts[0];
            // 反射或 dynamic 获取 content 字段
            string roomType = "";
            var prop = textObj.GetType().GetProperty("content");
            if (prop != null)
                roomType = prop.GetValue(textObj)?.ToString();
            string CaculatePromptTemplate = Prompt.CaculatePrompt;
            // 生成完整Prompt
            FinalPrompt = string.Format(CaculatePromptTemplate, roomType, coordinatesStr);

            Task.Run(async () =>
            {
                try
                {
                    await SelectEntitiesByRectangleAndPrintInfoAsync();
                }
                catch (System.Exception ex)
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n错误: {ex.Message}");
            }
            });

            // 输出结构化JSON
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(result, Newtonsoft.Json.Formatting.Indented);
            ed.WriteMessage($"\n{json}");

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
        }
        private async Task SelectEntitiesByRectangleAndPrintInfoAsync()
        {
            PaletteSetDlg dlg = new PaletteSetDlg();
            Action<string> updateFunc = null;
            // 等待 AI 控件初始化并获得更新函数
            var aiContentControl = await dlg.AppendMessageAsync("AI", "", true, setter => updateFunc = setter);
            await dlg.GetAIResponse(FinalPrompt, updateFunc);
        }

        [CommandMethod("DELETE_LAYER_AND_ENTITY", CommandFlags.Session)]
        public void DeleteLayerAndEntity()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 提示用户输入图层名
            PromptStringOptions pso = new PromptStringOptions("\n请输入要删除的图层名: ");
            pso.AllowSpaces = true;
            PromptResult pr = ed.GetString(pso);
            if (pr.Status != PromptStatus.OK) return;
            string layerName = pr.StringResult.Trim();

            using (DocumentLock docLock = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // 检查图层是否存在
                if (!lt.Has(layerName))
                {
                    ed.WriteMessage($"\n图层 \"{layerName}\" 不存在。");
                    return;
                }

                // 不能删除0层、Defpoints层或当前图层
                LayerTableRecord layerRec = (LayerTableRecord)tr.GetObject(lt[layerName], OpenMode.ForWrite);
                if (layerRec.IsDependent || layerRec.IsErased || layerRec.Name == "0" || layerRec.Name.ToUpper() == "DEFPOINTS" || db.Clayer == layerRec.ObjectId)
                {
                    ed.WriteMessage($"\n不能删除0层、Defpoints层或当前图层。");
                    return;
                }

                // 删除所有属于该图层的实体
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForWrite);
                    List<ObjectId> toErase = new List<ObjectId>();
                    foreach (ObjectId entId in btr)
                    {
                        Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent != null && ent.Layer == layerName)
                        {
                            toErase.Add(entId);
                        }
                    }
                    foreach (ObjectId entId in toErase)
                    {
                        Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                        ent.Erase();
                    }
                }

                // 删除图层
                layerRec.Erase();

                tr.Commit();
                ed.WriteMessage($"\n图层 \"{layerName}\" 及其所有实体已删除。");
            }
        }
        [CommandMethod("DELETE_EXCEPT_LAYERS", CommandFlags.Session)]
        public void DeleteExceptLayers()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 需要保留的图层名（全部大写，便于比较）
            HashSet<string> keepLayers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "COLUMN",
                "PUB_HATCH",
                "PUB_TEXT",
                "STAIR",
                "WALL",
                "WINDOW",
                "WINDOW_TEXT",
                "0",           // 通常建议保留0层
                "DEFPOINTS"    // 通常建议保留Defpoints层
            };

            using (DocumentLock docLock = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                List<string> toDelete = new List<string>();

                // 收集要删除的图层名
                foreach (ObjectId layerId in lt)
                {
                    LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                    string lname = ltr.Name;
                    if (!keepLayers.Contains(lname) && !ltr.IsDependent && !ltr.IsErased && db.Clayer != ltr.ObjectId)
                    {
                        toDelete.Add(lname);
                    }
                }

                // 删除每个图层上的实体和图层本身
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (string layerName in toDelete)
                {
                    // 删除实体
                    foreach (ObjectId btrId in bt)
                    {
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForWrite);
                        List<ObjectId> toErase = new List<ObjectId>();
                        foreach (ObjectId entId in btr)
                        {
                            Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                            if (ent != null && ent.Layer.Equals(layerName, StringComparison.OrdinalIgnoreCase))
                            {
                                toErase.Add(entId);
                            }
                        }
                        foreach (ObjectId entId in toErase)
                        {
                            Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                            ent.Erase();
                        }
                    }
                    // 删除图层
                    LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(lt[layerName], OpenMode.ForWrite);
                    ltr.Erase();
                    ed.WriteMessage($"\n已删除图层及其实体: {layerName}");
                }

                tr.Commit();
                ed.WriteMessage($"\n操作完成，已保留指定图层，其余图层及实体全部删除。");
            }
        }
        #endregion
    }
    }
