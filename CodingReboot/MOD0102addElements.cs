#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Reflection;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD0102addElements : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            //1. pick elements and filter them into list
            UIDocument uidoc = uiapp.ActiveUIDocument;
            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select Elements");

            // 2a. filter selected elements for curves
            List<CurveElement> allCurves = new List<CurveElement>();
            foreach (Element elem in pickList) 
            {
                if (elem is CurveElement)
                {
                    allCurves.Add(elem as CurveElement);
                }
            }

            //2b. filter selected elements for model curve
            List<CurveElement> modelCurves = new List<CurveElement>();
            foreach (Element elem in pickList)
            {
                if (elem is CurveElement)
                {
                    CurveElement curveElem = elem as CurveElement;
                    if (curveElem.CurveElementType == CurveElementType.ModelCurve) 
                    {
                        modelCurves.Add(curveElem);
                    }
                }
            }

            //3. curve data
            foreach(CurveElement currentCurve in modelCurves)
            {
                Curve curve = currentCurve.GeometryCurve;
                XYZ startPoint = curve.GetEndPoint(0);
                XYZ endPoint = curve.GetEndPoint(1);

                GraphicsStyle curStyle = currentCurve.LineStyle as GraphicsStyle;

                Debug.Print(curStyle.Name);

            }

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Revit elements");

                //4. create wall
                Level newlevel = Level.Create(doc, 20);
                Curve curCurve1 = modelCurves[0].GeometryCurve;

                Wall.Create(doc, curCurve1, newlevel.Id, false);

                FilteredElementCollector wallTypes = new FilteredElementCollector(doc);
                wallTypes.OfClass(typeof(WallType)); 

                Curve curCurve2 = modelCurves[1].GeometryCurve;
                Wall.Create(doc, curCurve2, wallTypes.FirstElementId(), newlevel.Id, 20, 0, false, false);

                t.Commit();
            }


            return Result.Succeeded;
        }


        internal string MyFirstMethod()
        {
            return "This is my first method!";
        }

        internal void MySecondMethod()
        {
            Debug.Print("This is my second method!");

        }

        internal string MyThirdMethod(string input)
        {
            return "This is my third method : "+ input; 
        }


        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

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
