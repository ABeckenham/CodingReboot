#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD0101_FizzBuzz : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            int numVariable = 250;
            double eleStart = 0;
            double flrHeight = 15;
            int planNumberCount = 0;
            int ceilingNumberCount = 0;
            int sheetNumberCount = 0;

            //collect for the titleblock
            FilteredElementCollector titleblockecollector = new FilteredElementCollector(doc);
            titleblockecollector.OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType();

            //collect the viewfamilytype for each view type: floor plan, ceiling plan
            FilteredElementCollector viewtypecollector = new FilteredElementCollector(doc);
            viewtypecollector.OfClass(typeof(ViewFamilyType));

            ViewFamilyType VFTfloorplan = null;
            ViewFamilyType VFTceilingplan = null;

            foreach (ViewFamilyType curType in viewtypecollector)
            {
                if (curType.ViewFamily == ViewFamily.FloorPlan)
                {
                    VFTfloorplan = curType;
                    
                }
            
                else if (curType.ViewFamily == ViewFamily.CeilingPlan)
                {
                    VFTceilingplan = curType;
                    
                }
            }

            Transaction t = new Transaction(doc);
            t.Start("Creating sheets, plans, and ceilings for the FizzBuzz challenge");

            //loop through each number variable between 1 - 250
            //create a level
            for (int i = 0; i <= numVariable; i++)
            {
                Level newLevel = Level.Create(doc, eleStart);
                eleStart = eleStart + flrHeight;
                newLevel.Name = "Level_" + i.ToString();

                //if(i % 3 == 0) // if dvisible by 3 - create floor plan and name it FIZZ_
                //{
                //    if (i % 5 != 0) // if the number is divisible by both 3 and 5 
                //    {
                //        ViewPlan newfPlan = ViewPlan.Create(doc, VFTfloorplan.Id, newLevel.Id);
                //        newfPlan.Name = "FIZZ_"+ i.ToString();
                //        planNumberCount = planNumberCount + 1;
                //    }
                //    else //divisable by 3 and 5 - create a sheet and name it FIZZBUZZ_
                //    {
                //        ViewSheet newsheet = ViewSheet.Create(doc, titleblockecollector.FirstElement().Id);
                //        newsheet.Name = "FIZZBUZZ_" + i.ToString();
                //        sheetNumberCount = sheetNumberCount + 1;
                //    }
                //}
                //else // if divisible by 5 - create a ceiling plan and Name it BUZZ_
                //{ 
                //   if(i % 5 == 0)
                //   {
                //    ViewPlan newCPlan = ViewPlan.Create(doc, VFTceilingplan.Id, newLevel.Id);
                //    newCPlan.Name = "BUZZ_" + i.ToString();
                //    ceilingNumberCount = ceilingNumberCount + 1;
                //   }
                //   else continue;                         
                //}

                if (i % 3 == 0 && i % 5 == 0) // if dvisible by 3 - create floor plan and name it FIZZ_
                {
                    ViewSheet newsheet = ViewSheet.Create(doc, titleblockecollector.FirstElement().Id);
                    newsheet.Name = "FIZZBUZZ_" + i.ToString();
                    sheetNumberCount = sheetNumberCount + 1;

                    //bonus
                    ViewPlan floorplan = ViewPlan.Create(doc, VFTfloorplan.Id, newLevel.Id);
                    XYZ point = new XYZ(1,0.5,0);
                    Viewport vp = Viewport.Create(doc, newsheet.Id, floorplan.Id, point);

                }
                else if(i % 3 == 0)
                {
                    ViewPlan newfPlan = ViewPlan.Create(doc, VFTfloorplan.Id, newLevel.Id);
                    newfPlan.Name = "FIZZ_" + i.ToString();
                    planNumberCount = planNumberCount + 1;

                    
                }
                
                else if (i % 5 == 0)
                {
                    ViewPlan newCPlan = ViewPlan.Create(doc, VFTceilingplan.Id, newLevel.Id);
                    newCPlan.Name = "BUZZ_" + i.ToString();
                    ceilingNumberCount = ceilingNumberCount + 1;
                }
                
                else continue;

                
            }


            TaskDialog.Show("Total Created Elements", ("Total Number of plans: " + planNumberCount) +
                Environment.NewLine + "Total Number of Ceilingplans: " + ceilingNumberCount +
                 Environment.NewLine + "Total number of Sheets: " + sheetNumberCount);

            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "MOD0101_FizzBuzz";
            string buttonTitle = "FizzBuzz";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for FizzBuzz");

            return myButtonData1.Data;
        }
    }
}
