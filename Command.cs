using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
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

        [CommandMethod("ZZ", CommandFlags.Session)]
        public void InsertBlockFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\test.dwg"; // 
            string blockName = "ZhaoMing";
            string targetLayer = "照明";

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }
                // 基准点坐标
                int baseX = 15548;
                int baseY = 25927;
                // 房间宽度
                int offsetX1 = 6317;
                // 房间长度
                int offsetY1 = 9234;
                // 第一个房间和第二个房间之间的距离
                int offsetX2 = 7000;
                // 第一个房间和第三个房间的距离
                int offsetX3 = 14000;

                var insertData = new List<(Point3d point, double angleDeg)>
    {
                    (new Point3d(baseX, baseY, 0), 0),
                    (new Point3d(baseX + offsetX1, baseY, 0), 90),
                    (new Point3d(baseX, baseY + offsetY1, 0), 270),
                    (new Point3d(baseX + offsetX1, baseY + offsetY1, 0), 180),
                    (new Point3d(baseX + offsetX2, baseY, 0), 0),
                    (new Point3d(baseX + offsetX1 + offsetX2, baseY, 0), 90),
                    (new Point3d(baseX + offsetX2, baseY + offsetY1, 0), 270),
                    (new Point3d(baseX + offsetX1 + offsetX2, baseY + offsetY1, 0), 180),
                    (new Point3d(baseX + offsetX3, baseY, 0), 0),
                    (new Point3d(baseX + offsetX1 + offsetX3, baseY, 0), 90),
                    (new Point3d(baseX + offsetX3, baseY + offsetY1, 0), 270),
                    (new Point3d(baseX + offsetX1 + offsetX3, baseY + offsetY1, 0), 180),
                };

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;



                    // 获取块定义
                    if (!bt.Has(blockName))
                    {
                        doc.Editor.WriteMessage($"\n未找到块定义：{blockName}");
                        return;
                    }
                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    foreach (ObjectId layerId in lt)
                    {
                        LayerTableRecord ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                        string layerName = ltr.Name;
                        // 这里可以输出到命令行或收集到列表
                        doc.Editor.WriteMessage($"\n图层名: {layerName}");
                    }
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    foreach (var (point, angleDeg) in insertData)
                    {
                        double angleInRadians = angleDeg * Math.PI / 180.0;
                        BlockReference br = new BlockReference(point, bt[blockName])
                        {
                            Rotation = angleInRadians,
                            Layer = targetLayer
                        };

                        modelSpace.AppendEntity(br);
                        tr.AddNewlyCreatedDBObject(br, true);
                    }
                    tr.Commit();
                }
            }
        }

        [CommandMethod("ZL", CommandFlags.Session)]
        public void InsertZhaoMingLineFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\zhaoming_line.dwg"; // 
            string blockName = "ZhaoMing-Line";
            string targetLayer = "WIRE-照明";
            Point3d insertPoint = new Point3d(15553, 25923, 0); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }

        [CommandMethod("LL", CommandFlags.Session)]
        public void InsertZhaoMingloutiFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\louti-zhaoming.dwg"; // 
            string blockName = "ZhaoMing-Louti";
            string targetLayer = "EQUIP-照明";
            Point3d insertPoint = new Point3d(18629, 24048, 0); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }
        [CommandMethod("YY", CommandFlags.Session)]
        public void InsertYinJiZhaoMingFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\yinji-new.dwg"; // 
            string blockName = "ZhaoMing-Yinji";
            string targetLayer = "照明";
            Point3d insertPoint = new Point3d(15772, 25615, 2); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }
        [CommandMethod("YL", CommandFlags.Session)]
        public void InsertYinJiZhaoLineFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\yinji-line.dwg"; // 
            string blockName = "Yinji-Line";
            string targetLayer = "WIRE-应急";
            Point3d insertPoint = new Point3d(17127, 25746, 0); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }
        [CommandMethod("CC", CommandFlags.Session)]
        public void InsertChaZuoFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\chazuo.dwg"; // 
            string blockName = "Chazuo";
            string targetLayer = "插座";
            Point3d insertPoint = new Point3d(15615, 27917, 0); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }
        [CommandMethod("CL", CommandFlags.Session)]
        public void InsertChaZuoLineFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\chazuo-line.dwg"; // 
            string blockName = "Chazuo-Line";
            string targetLayer = "WIRE-插座";
            Point3d insertPoint = new Point3d(15615, 32413, 2); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }
        [CommandMethod("ZS", CommandFlags.Session)]
        public void InsertZhuShiFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\zhaoming-text.dwg"; // 
            string blockName = "ZhaoMing-Text";
            string targetLayer = "照明";
            Point3d insertPoint = new Point3d(18289, 30122, 0); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }

        [CommandMethod("EE", CommandFlags.Session)]
        public void InsertZhuShiErrorFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\zhaoming-text-error.dwg"; // 
            string blockName = "ZhaoMing-Text-error";
            string targetLayer = "照明";
            Point3d insertPoint = new Point3d(18635, 23173, 1); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }

        [CommandMethod("TL", CommandFlags.Session)]
        public void InsertTuLiFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\tuli.dwg"; // 
            string blockName = "TuLi";
            string targetLayer = "PUB_TEXT";
            Point3d insertPoint = new Point3d(4182, 4113, 2); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }

        [CommandMethod("SM", CommandFlags.Session)]
        public void InsertShuoMingFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\shuoming_test.dwg"; // 
            string blockName = "dljfaldkf";
            string targetLayer = "PUB_TEXT";
            Point3d insertPoint = new Point3d(23827, 13394, 0); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }

        [CommandMethod("ST", CommandFlags.Session)]
        public void InsertSentenceFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\sentence.dwg"; // 
            string blockName = "Sentence";
            string targetLayer = "PUB_TEXT";
            Point3d insertPoint = new Point3d(37268, 28790, 0); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }

        [CommandMethod("WW", CommandFlags.Session)]
        public void InsertSwitchFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\switch.dwg"; // 
            string blockName = "Switch";
            string targetLayer = "插座";
            Point3d insertPoint = new Point3d(20950, 26139, 0); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }

        [CommandMethod("DL", CommandFlags.Session)]
        public void InsertDLFromDwg()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            string dwgFilePath = @"C:\Users\武丢丢\Documents\donglixiang.dwg"; // 
            string blockName = "Sentence";
            string targetLayer = "PUB_TEXT";
            Point3d insertPoint = new Point3d(21226, 25006, 2); // 插入点

            using (DocumentLock docLock = doc.LockDocument())
            {
                // 块插入操作应在事务外部
                // 这一步导入块定义到数据库中
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(dwgFilePath, System.IO.FileShare.Read, true, "");
                    db.Insert(blockName, sourceDb, true);
                }

                // 这一步将块参照，即块定义的实例，插入到图纸的指定位置中
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // 确保图层存在，不存在则创建
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(targetLayer))
                    {
                        lt.UpgradeOpen();
                        LayerTableRecord newLayer = new LayerTableRecord
                        {
                            Name = targetLayer
                        };
                        lt.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                    }

                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    br.Layer = targetLayer; // 指定插入的图层
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
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
        [CommandMethod("DELETE_LAYER_AND_ENTS", CommandFlags.Session)]
        public void DeleteLayerAndEntities(string layerName)
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

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
        private static Point3d GetCenterFromExtents(Extents3d ext)
        {
            double centerX = (ext.MinPoint.X + ext.MaxPoint.X) / 2.0;
            double centerY = (ext.MinPoint.Y + ext.MaxPoint.Y) / 2.0;
            double centerZ = (ext.MinPoint.Z + ext.MaxPoint.Z) / 2.0;
            return new Point3d(centerX, centerY, centerZ);
        }
        [CommandMethod("SELECT_RECT_PRINT", CommandFlags.Session)]
        public void SelectEntitiesByRectangleAndPrintInfo()
        {
            Document doc = CADApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

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

            var result = new
            {
                blocks = new Dictionary<string, List<double[]>>(), // 块名 -> List<中心坐标[x,y,z]>
                blockCount = new Dictionary<string, int>(),         // 块名 -> 数量
                polylineLengthByLayer = new Dictionary<string, double>(),       // 图层名 -> 总长度
                texts = new List<object>()                          // { 内容, 位置[x,y,z] }
            };

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
                        if (!result.blockCount.ContainsKey(blockName))
                            result.blockCount[blockName] = 0;
                        result.blockCount[blockName]++;
                        // 记录中心坐标
                        if (!result.blocks.ContainsKey(blockName))
                            result.blocks[blockName] = new List<double[]>();
                        try
                        {
                            var center = GetCenterFromExtents(br.GeometricExtents);
                            result.blocks[blockName].Add(new double[] { center.X, center.Y, center.Z });
                        }
                        catch { }
                        ed.WriteMessage($"\n类型: BlockReference, 块名: {blockName}, 插入点: {br.Position}, 对象ID: {ent.ObjectId}");
                    }
                    else if (ent is DBText dBText)
                    {
                        result.texts.Add(new
                        {
                            content = dBText.TextString,
                            position = new double[] { dBText.Position.X, dBText.Position.Y, dBText.Position.Z }
                        });
                        ed.WriteMessage($"\n类型: DBText, 内容: \"{dBText.TextString}\", 位置: {dBText.Position}, 对象ID: {ent.ObjectId}");
                    }
                    else if (ent is Polyline polyline)
                    {
                        string layerName = polyline.Layer;
                        if (!result.polylineLengthByLayer.ContainsKey(layerName))
                            result.polylineLengthByLayer[layerName] = 0;
                        result.polylineLengthByLayer[layerName] += polyline.Length;

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
        #endregion
    }
}
