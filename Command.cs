using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
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
                            //ent.UpgradeOpen();
                            // 
                            if (ent.IsWriteEnabled)
                            {
                                ent.Highlight();
                                ed.WriteMessage($"\n找到中文标注: {text.TextString}");
                            }
                            else
                            {
                                //  如果实体是以只读模式打开，则修改前需要调用upgradeopen，切换为写模式
                                ent.UpgradeOpen();
                                ent.Highlight();
                                ed.WriteMessage($"\n找到中文标注: {text.TextString}");
                            }
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
            string targetLayer = "TEST";

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
                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
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
            string targetLayer = "照明";
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
                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
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
            string targetLayer = "应急";
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
                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
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
                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
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
                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
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
                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
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
                    BlockReference br = new BlockReference(insertPoint, bt[blockName]);
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    tr.Commit();
                }
            }
        }
        #endregion
    }
}
