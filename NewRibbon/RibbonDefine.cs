using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;

namespace NewRibbon
{
    class RibbonDefine : IExternalApplication
    {
        // define a method that will create our tab and button

        static void AddRibbonPanel(UIControlledApplication application)
        {
            // Create a custom ribbon tab
            String tabName = "KOPEC";
            application.CreateRibbonTab(tabName);

            // Add a new ribbon panel
            RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "ROBOCZA");
            RibbonPanel ribbonPanel0 = application.CreateRibbonPanel(tabName, "Total Length");
            RibbonPanel ribbonPanel1 = application.CreateRibbonPanel(tabName, "ShP");
            RibbonPanel ribbonPanel2 = application.CreateRibbonPanel(tabName, "SQL");

            // Get dll assembly path
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            // create push buttons for Tabs
            PushButtonData b1DataRobocza = new PushButtonData(
                "cmdElementInfo",
                "xxx" + System.Environment.NewLine + "yyy",
                thisAssemblyPath,
                "ElementInfo.ShowElementInfo");

            PushButtonData b1DataTotalLength = new PushButtonData(
                "cmdCurveTotalLength",
                "Total" + System.Environment.NewLine + "  Length ",
                thisAssemblyPath,
                "TotalLength.CurveTotalLength");

            PushButtonData b1DataShP = new PushButtonData(
                "cmdAddSharedParam",
                "Dodaj" + System.Environment.NewLine + "Parametry",
                thisAssemblyPath,
                "SharedParametersLoad.SharedParam");

            PushButtonData b1SQL1 = new PushButtonData(
               "cmdUtworzTabele",
               "Dodaj" + System.Environment.NewLine + "Tabele SQL",
               thisAssemblyPath,
               "SharedParametersLoad.SharedParam");

            //Add buttons to Ribbon --> tabs
            PushButton pb1ElemetInfo = ribbonPanel.AddItem(b1DataRobocza) as PushButton;
            pb1ElemetInfo.ToolTip = "Wybierz Element";
            BitmapImage pb1ImageElementInfo = new BitmapImage(new Uri("pack://application:,,,/NewRibbon;component/Resources/AdskSample_ico.png"));
            pb1ElemetInfo.LargeImage = pb1ImageElementInfo;

            PushButton pb1TotalLength = ribbonPanel0.AddItem(b1DataTotalLength) as PushButton;
            pb1TotalLength.ToolTip = "Wybierz linie";
            BitmapImage pb1ImageTotalLength = new BitmapImage(new Uri("pack://application:,,,/NewRibbon;component/Resources/totalLength.png"));
            pb1TotalLength.LargeImage = pb1ImageTotalLength;

            PushButton pb1LoadShP = ribbonPanel1.AddItem(b1DataShP) as PushButton;
            pb1LoadShP.ToolTip = "Zaladuj Parametry";
            BitmapImage pb1ImageLoadShP = new BitmapImage(new Uri("pack://application:,,,/NewRibbon;component/Resources/add_shared_parameter.ico"));
            pb1LoadShP.LargeImage = pb1ImageLoadShP;

            PushButton pbSqlRvt = ribbonPanel2.AddItem(b1SQL1) as PushButton;
            pbSqlRvt.ToolTip = "Utworz Tabele SQL";
            BitmapImage pbImageLoadSQL = new BitmapImage(new Uri("pack://application:,,,/NewRibbon;component/Resources/add_shared_parameter.ico"));
            pbSqlRvt.LargeImage = pbImageLoadSQL;



        }
        public Result OnShutdown(UIControlledApplication application)
        {
            // do nothing
            // return result.succeeded
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            // call our method that will load up our toolbar
            // return result.succeeded
            AddRibbonPanel(application);
            return Result.Succeeded;
        }

    }
}