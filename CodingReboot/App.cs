#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.Versioning;
using System.Windows.Markup;

#endregion

namespace CodingReboot
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            // 1. Create ribbon tab
            string tabName = "CodingReboot";
            try
            {
                app.CreateRibbonTab(tabName);
            }
            catch (Exception)
            {
                Debug.Print("Tab already exists.");
            }

            // 2. Create ribbon panel 
            RibbonPanel panelmod1 = Utils.CreateRibbonPanel(app, tabName, "MOD01");
            RibbonPanel panelmod2 = Utils.CreateRibbonPanel(app, tabName, "MOD02");
            RibbonPanel panelRidge = Utils.CreateRibbonPanel(app, tabName, "Ridge");

            // 3. Create button data instances
            //PushButtonData btnData1 = Command1.GetButtonData();
            //PushButtonData btnData2 = Command2.GetButtonData();
            PushButtonData MOD0101 = MOD0101_FizzBuzz.GetButtonData();
            PushButtonData MOD0102 = MOD0102_HiddenMessage.GetButtonData();
            PushButtonData MOD0103 = MOD0103_MovingDay.GetButtonData();
            PushButtonData MOD01addexcel = MOD01AddExcelSheets.GetButtonData();
            PushButtonData MOD0201Schedules = MOD0201_Schedules.GetButtonData();


            // 4. Create buttons
            //PushButton myButton1 = panel.AddItem(btnData1) as PushButton;
            //PushButton myButton2 = panel.AddItem(btnData2) as PushButton;
            PushButton myButton3 = panelmod1.AddItem(MOD0101) as PushButton;
            PushButton myButton4 = panelmod1.AddItem(MOD0102) as PushButton;
            PushButton myButton5 = panelmod1.AddItem(MOD0103) as PushButton;
            PushButton myButton6 = panelRidge.AddItem(MOD01addexcel) as PushButton;
            PushButton myButton7 = panelmod2.AddItem(MOD0201Schedules) as PushButton;
            //PushButton myButton7 = panel.AddItem(MOD01addexcel) as PushButton;


            //5. create split button (swops the button depending on choice)
            //SplitButtonData splitbuttonData = new SplitButtonData("split1","split\rButton");
            //SplitButton splitbutton = panel.AddItem(splitbuttonData) as SplitButton;
            //splitbutton.AddPushButton(btnData1);
            //splitbutton.AddPushButton(btnData2);

            //6. create pulldown button
            //PulldownButtonData pullbuttondata = new PulldownButtonData("pulldown1", "Pull Down");
            //pullbuttondata.LargeImage = ButtonDataClass.BitmapToImageSource(Properties.Resources.Red_32);
            //pullbuttondata.Image = ButtonDataClass.BitmapToImageSource(Properties.Resources.Red_16);
            //PulldownButton pulldownbutton = panelmod1.AddItem(pullbuttondata) as PulldownButton;
            //pulldownbutton.AddPushButton(MOD0103);
            //pulldownbutton.AddPushButton(MOD0102);

            //7. stacked buttons (vertically stacked) 
            //panel.AddStackedItems(MOD01addexcel, MOD01addexcel2);

            //NOTE:
            //    To create a new tool, copy lines 35 and 39 and rename the variables to "btnData3" and "myButton3".
            //     Change the name of the tool in the arguments of line

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }


    }
}
