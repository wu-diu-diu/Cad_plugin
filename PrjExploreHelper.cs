using System;
using System.Windows.Forms;

using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using XCAD = Autodesk.AutoCAD;
using XCADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using XcadPopupMenuItem = Autodesk.AutoCAD.Interop.AcadPopupMenuItem;
using XcadPopupMenu = Autodesk.AutoCAD.Interop.AcadPopupMenu;
using XcadToolbar = Autodesk.AutoCAD.Interop.AcadToolbar;
using XcadToolbarItem = Autodesk.AutoCAD.Interop.AcadToolbarItem;
using XcadMenuGroups = Autodesk.AutoCAD.Interop.AcadMenuGroups;
using XcadMenuBar = Autodesk.AutoCAD.Interop.AcadMenuBar;
using Autodesk.AutoCAD.Customization;


namespace CoDesignStudy.Cad.PlugIn
{
    public class PrjExploreHelper
    {
        //主托盘
        public static PaletteSet MainPaletteset { get; set; }

        //菜单
        static XcadPopupMenu MainMenu { get; set; }

        /// <summary>
        /// 初始化协同面板
        /// </summary>
        internal static void InitPalette()
        {
            try
            {
                MainPaletteset = new PaletteSet("ChatCAD", Guid.NewGuid());
                MainPaletteset.Style = PaletteSetStyles.Snappable | PaletteSetStyles.ShowCloseButton | PaletteSetStyles.ShowAutoHideButton;
                MainPaletteset.DockEnabled = DockSides.Left | DockSides.None;
                MainPaletteset.Visible = true;

                double width = (int)XCADApp.MainWindow.DeviceIndependentSize.Width;
                double height = (int)XCADApp.MainWindow.DeviceIndependentSize.Height;
                int palettesetWidth = Convert.ToInt32(width * 0.25);  // 侧边栏的宽度
                int palettesetHeight = Convert.ToInt32(height * 0.5);  // 侧边栏的高度
                MainPaletteset.Size = new System.Drawing.Size(palettesetWidth, palettesetHeight);

                MainPaletteset.Dock = DockSides.Left;
                MainPaletteset.KeepFocus = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("初始化面板:", ex.Message);
            }
        }
        internal static void InitWebPalette()
        {
            try
            {
                MainPaletteset = new PaletteSet("Web服务", Guid.NewGuid());
                MainPaletteset.Style = PaletteSetStyles.Snappable | PaletteSetStyles.ShowCloseButton | PaletteSetStyles.ShowAutoHideButton;
                MainPaletteset.DockEnabled = DockSides.Left | DockSides.None;
                MainPaletteset.Visible = true;

                double width = (int)XCADApp.MainWindow.DeviceIndependentSize.Width;
                double height = (int)XCADApp.MainWindow.DeviceIndependentSize.Height;
                int palettesetWidth = Convert.ToInt32(width * 0.7);  // 侧边栏的宽度
                int palettesetHeight = Convert.ToInt32(height * 0.5);  // 侧边栏的高度
                MainPaletteset.Size = new System.Drawing.Size(palettesetWidth, palettesetHeight);

                MainPaletteset.Dock = DockSides.Left;
                MainPaletteset.KeepFocus = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("初始化面板:", ex.Message);
            }
        }

        internal static void InitMenu()
        {
            XcadMenuGroups menuGroups = (XcadMenuGroups)XCADApp.MenuGroups;
            XcadMenuBar XcadMenuBar = (XcadMenuBar)XCADApp.MenuBar;
            MainMenu = menuGroups.Item(0).Menus.Add("测试菜单");
            XcadPopupMenuItem pmi = MainMenu.AddMenuItem(MainMenu.Count + 1, "显示面板", "GW_ShowPalette ");
            MainMenu.InsertInMenuBar(XcadMenuBar.Count - 1);
        }
    }
}
