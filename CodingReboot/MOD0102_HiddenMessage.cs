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
using System.Security.Cryptography;
using System.Windows.Controls;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD0102_HiddenMessage : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here



            //create user selection prompt
            UIDocument uidoc = uiapp.ActiveUIDocument;
            TaskDialog.Show("Select Lines", "Select some lines to convert to Revit Elements");
            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("select Element");
            

            //Cast the element list to curve elements

            //List<CurveElement> allCurve = new List<CurveElement>();
            //foreach (Element elem in pickList) 
            //{
            //    if (elem is CurveElement)
            //    {
            //        allCurve.Add(elem as CurveElement);
            //    }
            //}

            //create a list of model lines and detail lines 
                //CurveElement curveElem = elem as CurveElement;
            
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

            TaskDialog.Show("Curves", $"you selected {modelCurves.Count} lines.");
            
            //select model curves by name
            // filtered elements collector for: level wall, pipe, and duct type, SYSTEM TYPE
            FilteredElementCollector levelcollect = new FilteredElementCollector(doc);
            levelcollect.OfCategory(BuiltInCategory.OST_Levels);            
           

            FilteredElementCollector DuctType = new FilteredElementCollector(doc);
            DuctType.OfClass(typeof(DuctType));

            FilteredElementCollector pipeType = new FilteredElementCollector(doc);
            pipeType.OfClass(typeof(PipeType));

            //Get system types for the pipe and ducting : "Domestic Hot Water", "Supply Air"
            FilteredElementCollector systemCollector = new FilteredElementCollector(doc);
            systemCollector.OfClass(typeof(MEPSystemType));

            //make a duct/pipe needs = doc, systemType, ductType, LevelId, start and end points           
            //foreach (MEPSystemType MEPtype in systemCollector) 
            //{
            //    if(MEPtype.Name == "Domestic Hot Water")
            //    {
            //        syshotwater = MEPtype;
            //        break;
            //    }

            //}

            //foreach (MEPSystemType MEPtype in systemCollector)
            //{
            //    if (MEPtype.Name == "Supply Air")
            //    {
            //        sysairsupply = MEPtype;
            //        break;
            //    }                    
            //}
            //you are repeating yourself make a method.
            //get mepsystem type 
            MEPSystemType syshotwater = GetMEPSystemType(doc, "Domestic Hot Water");
            MEPSystemType sysairsupply = GetMEPSystemType(doc, "Supply Air");

            //get current doc level 1 
            //Level selectLevel = null;
            //foreach (Level level in levelcollect)
            //{
            //    if(level.Name == "Level 1")
            //    {
            //        selectLevel = level;
            //    }
            //}

            //make a method to collect the two wall types

            WallType storeWall = GetWallTypeByName(doc, "Storefront");
            WallType genWall = GetWallTypeByName(doc, @"Generic - 8""");        

            List<ElementId> linestohide = new List<ElementId>();

            //start the creation of elements

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Revit Elements");

                Level newLevel = Level.Create(doc, 10);

                // switch 
                foreach (CurveElement curCurve in modelCurves)
                {                    
                    //CurveElementType
                    Curve curvve = curCurve.GeometryCurve;
                    //need to deal with circles that dont have a start or end point
                    if (curvve.IsBound == true)
                    {
                        XYZ startpoint = curvve.GetEndPoint(0);
                        XYZ endpoint = curvve.GetEndPoint(1);

                        GraphicsStyle curStyle = curCurve.LineStyle as GraphicsStyle;

                        switch (curStyle.Name)
                        {
                            case "A-GLAZ":
                                Wall.Create(doc, curvve, storeWall.Id, newLevel.Id, 10, 0, false, false);
                                break;

                            case "A-WALL":
                                Wall.Create(doc, curvve, genWall.Id, newLevel.Id, 10, 0, false, false);
                                break;

                            case "M-DUCT":
                                Duct.Create(doc, sysairsupply.Id, DuctType.FirstElementId(),
                                    newLevel.Id, startpoint, endpoint);
                                break;

                            case "P-PIPE":
                                Pipe.Create(doc, syshotwater.Id, pipeType.FirstElementId(),
                                    newLevel.Id, startpoint, endpoint);
                                break;

                            default:
                                linestohide.Add(curCurve.Id);
                                break;
                        }

                    }
                    else
                    {
                        linestohide.Add(curCurve.Id);
                        continue;
                    }
                        
                }

                doc.ActiveView.HideElements(linestohide);

                t.Commit();
            }
                //create methods for as much as possible
            

            return Result.Succeeded;
        }
        
        // my new methods ===========================

        internal MEPSystemType GetMEPSystemType(Document doc, string sysName)
        {
            FilteredElementCollector Collector = new FilteredElementCollector(doc);
            Collector.OfClass(typeof(MEPSystemType));


            foreach (MEPSystemType curType in Collector)
            {
                if (curType.Name == sysName)
                {
                    return curType;
                }
            }    
            return null;            
        }

        internal WallType GetWallTypeByName(Document doc, string name)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach (WallType curType in collector)
            {
                if (curType.Name == name) 
                {
                    return curType;
                }
            }
            return null;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "MOD0102_HiddenMessage";
            string buttonTitle = "Hidden Message";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.icons8_invisible_pulsar_color_3296,
                Properties.Resources.icons8_invisible_pulsar_color_1696,
                "This is a tooltip for Hidden Message");

            return myButtonData1.Data;
        }
    }
}
