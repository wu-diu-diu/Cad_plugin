using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoDesignStudy.Cad.PlugIn
{
    public class DeviceRecord
    {
        public string DeviceType { get; set; }      // 设备类型，如“吸顶灯”、“插座”
        public int Count { get; set; }              // 数量
        public string Info { get; set; }          // 功率（W）
        // 构造函数，在new一个类的实例的时候自动调用
        public DeviceRecord(string deviceType, int count, string info)
        {
            DeviceType = deviceType;
            Count = count;
            Info = info;
        }
    }
}
