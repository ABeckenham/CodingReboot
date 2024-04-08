#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.Reflection;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]

    public class MOD0101 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            //create a transaction to lock the model
            Transaction t = new Transaction(doc);
            //this text describes what the tool is doing
            t.Start("Doing something in Revit");
            

            //text variables
            string text1 = "this is my text1";
            string text2 = "this is my text2";
            
            // Revit API works in Feet and inches
            // convert meters to feet 
            double meters = 4;
            double metersToFeet = meters * 3.28084;

            //convert mm to feet
            double mm = 3500;
            double mmtofeet = mm / 304.8;

            //find the reminder when dividing (ie. the modulo or mod)
            double remainder1 = 100 % 10; //equals 0 (100 divided by 10 = 10)
            double remainder2 = 100 % 9; // equals 1 (100 divided by 9 = 11 with remainder of 1

            //increment a number by 1
            double number1 = 100;
            number1++;
            number1--;
            number1 += 10;// takes value + on 10 to it.

            //boolean operators
            // == equals
            // != not equal
            // > greater than
            // < less than 
            // >= more than or equal <= less than or equal

            if (number1 > 10)
            {
                // do something 
            }

            if (number1 > 6)
            {
                // do something
            }
            else if (number1 == 20)                    
            {
                //blar
            }
            else
            {
               //do something
            }

            // compound conditional stastement

            if ( number1 > 10 && number1 == 100)
            {
                //do something
            }

            //look for either thing using || "or" statements
            if (number1>10 || number1 ==100)
            {
                //do something
            }

            //lists define a list and the data type
            // new is used for instantiating, creating a new object
            List<string> list1 = new List<string>();

            list1.Add(text1);
            list1.Add(text2);
            list1.Add("This is more text");

            //create list and add items to it 
            List<int> list2 = new List<int> { 1, 2, 3, 4, 5 };


            //LOOPS
            //loop through a list
            int letterCounter = 0;
            foreach (string currentstring in list1)
            {
                //letterCounter = letterCounter + currentstring.Length;
                letterCounter += currentstring.Length;
            }

            //for loop
            int numbercount = 0;
            int counter = 100;
            for (int i = 0; i <= counter; i++)
            {
                numbercount += i;
            }

            TaskDialog.Show("Number counter", "the number counter is " + numbercount.ToString());


            //create a floor level
            double elevation = 10;
            Level newlevel = Level.Create(doc, elevation);

            newlevel.Name = "my new level";


            //create new views 
            //filtered element collector
            FilteredElementCollector collector1 = new FilteredElementCollector(doc);
            collector1.OfClass(typeof(ViewFamilyType));

            ViewFamilyType floorplanVFT = null;
            foreach (ViewFamilyType curVFT in collector1)
            {
                if (curVFT.ViewFamily == ViewFamily.FloorPlan)
                {
                    floorplanVFT = curVFT;
                    break;
                }
            }    


            ViewPlan newPlan = ViewPlan.Create(doc, floorplanVFT.Id, newlevel.Id);
            newPlan.Name = "my new floor plan"; 

            //get ceiling plan view family type

            ViewFamilyType cielingplanVFT = null;
            foreach (ViewFamilyType curVFT in collector1)
            {
                if (curVFT.ViewFamily == ViewFamily.CeilingPlan)
                {
                    cielingplanVFT = curVFT;
                    break;
                }
            }

            //create a ceiling plan
            ViewPlan newCPlan = ViewPlan.Create(doc, cielingplanVFT.Id, newlevel.Id);
            newPlan.Name = "my new Ceiling plan";


            //create a sheet
            //filtered element collector
            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            collector2.OfCategory(BuiltInCategory.OST_TitleBlocks);

            ViewSheet newsheet = ViewSheet.Create(doc, collector2.FirstElementId());
            newsheet.Name = "my new sheet";
            newsheet.SheetNumber = "101";

            //add a view to a sheet using a viewport
            //first create a point
            XYZ insPoint = new XYZ(1, 0.5, 0);

            Viewport newviewport = Viewport.Create(doc, newsheet.Id, newPlan.Id, insPoint);

            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
