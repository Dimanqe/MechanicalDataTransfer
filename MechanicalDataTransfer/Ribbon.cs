using System;
using System.Linq;
using System.Drawing;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Windows;
using System.Windows.Forms;
using System.Windows;
using NPOI.Util.Collections;

namespace MechanicalDataTransfer
{
    [Transaction(TransactionMode.Manual)]
    public class Ribbon : IExternalApplication
    {
        private const string TabName = "1ДСК";

        public Result OnStartup(UIControlledApplication application)
        {
            if (SystemInformation.UserDomainName != "PROGRESS" && SystemInformation.UserDomainName != "dsk1.man" &&
                SystemInformation.UserDomainName != "KAPP@DSK1" &&
                SystemInformation.UserDomainName != "AUTH") return Result.Succeeded;

            if (!ComponentManager.Ribbon.Tabs.Any(i => i.Name == TabName)) application.CreateRibbonTab(TabName);
            var RibbonName = "Комплект Марка";
            var RibbonPanel = application.GetRibbonPanels(TabName).FirstOrDefault(i => i.Name == RibbonName);
            if (RibbonPanel == null) RibbonPanel = application.CreateRibbonPanel(TabName, RibbonName);

            var DllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;

            var PushButtonData1 = new PushButtonData("Перенос расчётных данных", "Перенос расчётных данных", DllPath,
                "MechanicalDataTransfer.MechanicalDataTransferCom");
            PushButtonData1.LargeImage = ImgToSource(Properties.Resources.MechanicalTransferIcon);
            PushButtonData1.ToolTip = "Перенос расчётных данных";
            var pushButton1 = (PushButton)RibbonPanel.AddItem(PushButtonData1);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private BitmapSource ImgToSource(Bitmap source)
        {
            return System.Windows.Interop.Imaging
                .CreateBitmapSourceFromHBitmap(source.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
        }
    }
}