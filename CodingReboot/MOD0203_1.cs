#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD0203_1: IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            //1. GET MODEL LINES (get all elements you want to tag)
            FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
            collector.OfClass(typeof(CurveElement));

            //2. create reference array and point list
            ReferenceArray referenceArray = new ReferenceArray();
            List<XYZ> pointList = new List<XYZ>();

            //3. Loop through lines (get geometry)
            foreach(ModelLine curLine in collector)
            {
                //3a. get midpoint of line 
                Curve curve = curLine.GeometryCurve;
                XYZ midPoint = curve.Evaluate(0.75, true);

                //7. check line is vertical

                if (Utils.IsLineVertical(curve) == false)
                    continue;

                //3b. add lines as a reference to ref array list
                referenceArray.Append(new Reference(curLine));

                //3c. add midpoint to list 
                pointList.Add(midPoint);
            }

            //the list of points will be in order of items placed so we need to
            //sort elements into a useable list - references need to be next to each other
            //4. order list left to right 
            List<XYZ> sortedList = pointList.OrderBy(p => p.X).ThenBy(p => p.Y).ToList();
            XYZ point1 = sortedList.First();
            XYZ point2 = sortedList.Last();

            //5. Create line for dimension (not a visible line, but one commited to memory)
            //Line dimLine = Line.CreateBound(point1, point2);
            Line dimLine = Line.CreateBound(point1, new XYZ(point2.X, point1.Y,0));

            //6. create dimension
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create dimension");
                Dimension newdim = doc.Create.NewDimension(doc.ActiveView, dimLine, referenceArray);
                t.Commit();
            }


            return Result.Succeeded;
        }

        

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Command 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 2");

            return myButtonData1.Data;
        }
    }
}
