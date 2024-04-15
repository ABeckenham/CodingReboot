#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
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
    public class MOD0102 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            //1.PICK ELEMENTS AND FILTER THEM INTO A LIST
            UIDocument uidoc = uiapp.ActiveUIDocument;
            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("select Element");


            TaskDialog.Show("Selection", "You have selected "+ pickList.Count.ToString()+" Elements");

            //2a. filter selected elements for curve elements
            List<CurveElement> allCurves = new List<CurveElement>();
            foreach (Element elem in pickList) 
            {
                if (elem is CurveElement)
                {
                    allCurves.Add(elem as CurveElement);
                }
            }

            //2b. filter selected elements for model curves
            List<CurveElement> modelCurves = new List<CurveElement>();
            foreach (Element elem in pickList)
            {
                //CurveElement curveElem = elem as CurveElement;
                CurveElement curveElem = elem as CurveElement;
                if(curveElem.CurveElementType == CurveElementType.ModelCurve)
                {
                    modelCurves.Add(curveElem);
                }
            }

            // 3. curve data
            foreach (CurveElement currentCurve in modelCurves)
            {
                Curve curve = currentCurve.GeometryCurve;
                XYZ startPoint = curve.GetEndPoint(0); //zero index gets the start point
                XYZ endPoint = curve.GetEndPoint(1); //1 index gets the end 

                GraphicsStyle curStyle = currentCurve.LineStyle as GraphicsStyle;

                Debug.Print(curStyle.Name);

            }

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Revit Elements");
                //4. Create a wall
                Level newLevel = Level.Create(doc, 20);
                Curve curCurve1 = modelCurves[0].GeometryCurve;

                //Wall.Create(doc, curCurve1, newLevel.Id, false); // select a random wall type

                //5. create a wall specific
                FilteredElementCollector wallTypes = new FilteredElementCollector(doc);
                wallTypes.OfClass(typeof(WallType));

                Curve curCurve2 = modelCurves[1].GeometryCurve;
                WallType wallTypeSelect = getWallTypeByName(doc, "Basic Wall");
                Wall.Create(doc, curCurve2, wallTypeSelect.Id, newLevel.Id, 10, 0, false, false);

                //6. get system types
                FilteredElementCollector systemType = new FilteredElementCollector(doc);
                systemType.OfClass(typeof(MEPSystemType));

                //7. get duct system type
                MEPSystemType ductSystemType = getMEPSystemTypeByName(doc, "Supply Air"); 
                
                //8. get duct family types
                FilteredElementCollector DuctType = new FilteredElementCollector(doc);
                DuctType.OfClass(typeof(DuctType));

                //9. create a duct
                Curve curCurve3 = modelCurves[2].GeometryCurve;
                Duct newDuct = Duct.Create(doc,ductSystemType.Id, DuctType.FirstElementId(), 
                    newLevel.Id, curCurve3.GetEndPoint(0), curCurve3.GetEndPoint(1));

                //10. get pipe system type
                MEPSystemType pipeSystemType = getMEPSystemTypeByName(doc, "Domestic Hot Water");

                //11. get pipe family type                
                FilteredElementCollector pipeType = new FilteredElementCollector(doc);
                pipeType.OfClass(typeof(PipeType));

                //12. create a pipe
                Curve curCurve4 = modelCurves[3].GeometryCurve;
                Pipe newPipe = Pipe.Create(doc, pipeSystemType.Id, pipeType.FirstElementId(),
                newLevel.Id, curCurve3.GetEndPoint(0), curCurve3.GetEndPoint(1));

                //13. use out new methods
                string testdtring = MyFirstMethod();
                MySecondMethod();
                string teststring2 = MyThirdMethod("lalala");

                //15. switch statement
                int numberValue = 5;
                string numAsString = "";

                switch (numberValue)
                {
                    case 1:
                        numAsString = "One";
                        break;

                    case 2:
                        numAsString = "Two";
                        break;

                    case 3:
                        numAsString = "Three";
                        break;

                    case 4:
                        numAsString = "Four";
                        break;

                    case 5:
                        numAsString = "Five";
                        break;

                    default:
                        numAsString = "Zero";
                        break;

                }

                //16. advanced switch stagements
                Curve curCurve5 = modelCurves[4].GeometryCurve;
                GraphicsStyle Curve5GS = modelCurves[1].LineStyle as GraphicsStyle;

                WallType wallType1 = getWallTypeByName(doc, "Storefront");
                WallType wallType2 = getWallTypeByName(doc, "Exterior - Brick on CMU");

                switch(Curve5GS.Name)
                {
                    case "<Thin Lines>":
                        Wall.Create(doc, curCurve5, wallType1.Id, newLevel.Id, 20, 0, false, false);
                        break;

                    case "<Wide Lines>":
                        Wall.Create(doc, curCurve5, wallType1.Id, newLevel.Id, 20, 0, false, false);
                        break;

                    default:
                        Wall.Create(doc, curCurve5, newLevel.Id, false);
                        break;
                }

                t.Commit();
            }

            return Result.Succeeded;
        }
        
        //custom methods =================================

        //method with return value string
        internal string MyFirstMethod()
        {
            return "This is my first method!";
        }

        //method with out return value
        internal void MySecondMethod()
        {
            Debug.Print("Thus is my second method!");
        }

        internal string MyThirdMethod(string input)
        {
            return "This is my Third Method: " + input;
        }

        //get wall type method
        internal WallType getWallTypeByName(Document doc, string typename)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));
            foreach(WallType curType in collector)
            {
                if (curType.Name == typename)
                {
                    return curType; 
                }
            }
            return null;
        }


        internal MEPSystemType getMEPSystemTypeByName(Document doc, string typename) 
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(MEPSystemType));
            
            foreach(MEPSystemType curType in collector)
            {
                if(curType.Name == typename) 
                {
                    return curType; 
                }
            }
            return null;
        }

        //add-inn toolbar method
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
