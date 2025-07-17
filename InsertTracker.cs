using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Org.BouncyCastle.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoDesignStudy.Cad.PlugIn
{
    public static class RoomInputCache
    {
        public static string RoomType { get; set; }
        public static string CoordinatesStr { get; set; }
        public static string DoorPositionStr { get; set; }

        public static bool HasValidData()
        {
            return !string.IsNullOrEmpty(RoomType) &&
                   !string.IsNullOrEmpty(CoordinatesStr) &&
                   !string.IsNullOrEmpty(DoorPositionStr);
        }
        public static void SetRoomDrawingInputs(string roomType, string coordinatesStr, string doorPositionStr)
        {
            RoomType = roomType;
            CoordinatesStr = coordinatesStr;
            DoorPositionStr = doorPositionStr;
        }
        public static (string roomType, string coordinatesStr, string doorPositionStr) GetLastRoomDrawingInputs()
        {
            return (RoomType, CoordinatesStr, DoorPositionStr);
        }

    }
    public static class InsertTracker
    {
        private static List<ObjectId> currentBatch = new List<ObjectId>();
        private static List<ObjectId> lastBatch = new List<ObjectId>();

        private static Dictionary<string, double> currentComponentCounts = new Dictionary<string, double>();
        private static Dictionary<string, double> lastComponentCounts = new Dictionary<string, double>();
        private static string lastRoomType = "";
        private static string lastCoordinatesStr = "";
        private static string lastDoorPositionStr = "";

        public static void BeginNewInsert()
        {
            currentBatch.Clear();
            currentComponentCounts.Clear();
        }
        public static void SetRoomDrawingInputs(string roomType, string coordinatesStr, string doorPositionStr)
        {
            lastRoomType = roomType;
            lastCoordinatesStr = coordinatesStr;
            lastDoorPositionStr = doorPositionStr;
        }
        public static (string roomType, string coordinatesStr, string doorPositionStr) GetLastRoomDrawingInputs()
        {
            return (lastRoomType, lastCoordinatesStr, lastDoorPositionStr);
        }

        public static void AddEntity(ObjectId id)
        {
            if (!id.IsNull)
                currentBatch.Add(id);
        }

        public static void AddComponentCount(string name, double count)
        {
            if (!string.IsNullOrEmpty(name))
            {
                if (currentComponentCounts.ContainsKey(name))
                    currentComponentCounts[name] += count;
                else
                    currentComponentCounts[name] = count;
            }
        }

        public static void CommitInsert()
        {
            lastBatch = new List<ObjectId>(currentBatch);
            lastComponentCounts = new Dictionary<string, double>(currentComponentCounts);
            currentBatch.Clear();
            currentComponentCounts.Clear();
        }

        public static bool HasLastInsert() => lastBatch.Count > 0;

        public static void DeleteLastInserted(Database db, Dictionary<string, (double Count, string Info)> componentStats)
        {
            if (lastBatch == null || lastBatch.Count == 0)
                return;
            using (DocumentLock docLock = Application.DocumentManager.MdiActiveDocument.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in lastBatch)
                {
                    try
                    {
                        Entity ent = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                        ent?.Erase();
                    }
                    catch { }
                }
                tr.Commit();
            }

            foreach (var kv in lastComponentCounts)
            {
                string name = kv.Key;
                double toRemove = kv.Value;

                if (componentStats.ContainsKey(name))
                {
                    var (count, info) = componentStats[name];
                    double newCount = count - toRemove;
                    if (newCount <= 0)
                        componentStats.Remove(name);
                    else
                        componentStats[name] = (newCount, info);
                }
            }

            lastBatch.Clear();
            lastComponentCounts.Clear();
        }
    }
}
