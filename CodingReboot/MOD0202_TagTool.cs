#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Reflection;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD0202_TagTool : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            View curView = doc.ActiveView;
            FilteredElementCollector Collector = new FilteredElementCollector(doc, curView.Id);

            //List<BuiltInCategory> catList = new List<BuiltInCategory>();
            //catList.Add(BuiltInCategory.OST_Areas);
            //catList.Add(BuiltInCategory.OST_Doors);
            //catList.Add(BuiltInCategory.OST_Furniture);
            //catList.Add(BuiltInCategory.OST_LightingFixtures);
            //catList.Add(BuiltInCategory.OST_Rooms);
            //catList.Add(BuiltInCategory.OST_Walls);

            //ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catList);
            //Collector.WherePasses(catfilter).WhereElementIsNotElementType();

            List<BuiltInCategory> catListFP = new List<BuiltInCategory>();
            catListFP.Add(BuiltInCategory.OST_Doors);
            catListFP.Add(BuiltInCategory.OST_Furniture);
            catListFP.Add(BuiltInCategory.OST_LightingFixtures);
            catListFP.Add(BuiltInCategory.OST_Rooms);
            catListFP.Add(BuiltInCategory.OST_Walls);
            catListFP.Add(BuiltInCategory.OST_Windows);

            

            List<BuiltInCategory> catListCP = new List<BuiltInCategory>();
            catListCP.Add(BuiltInCategory.OST_LightingFixtures);
            catListCP.Add(BuiltInCategory.OST_Rooms);

            List<BuiltInCategory> catListAP = new List<BuiltInCategory>();
            catListAP.Add(BuiltInCategory.OST_Areas);

            List<BuiltInCategory> catListS = new List<BuiltInCategory>();
            catListS.Add(BuiltInCategory.OST_Rooms);

                        
           
            //could i make multiple collectors for each view then run the script on which view it is??

            //make multiple catfilters for each view type. 


            //LINQ
            FamilySymbol curDoorTag = Utils.GetTagbyName(doc, "M_Door Tag");
            FamilySymbol curAreaTag = Utils.GetTagbyName(doc, "M_Area");
            FamilySymbol curCwallTag = Utils.GetTagbyName(doc, "M_Curtain Wall Tag");
            FamilySymbol curWallTag = Utils.GetTagbyName(doc, "M_Wall Tag");
            FamilySymbol curFurnTag = Utils.GetTagbyName(doc, "M_Furniture Tag");
            FamilySymbol curLightTag = Utils.GetTagbyName(doc, "Lighting Fixture Tag");
            FamilySymbol curRoomTag = Utils.GetTagbyName(doc, "M_Room Tag");
            FamilySymbol curWindowTag = Utils.GetTagbyName(doc, "M_Window Tag");

            //Dictionary
            //2. create dictionary for tags
            Dictionary<string, FamilySymbol> tags = new Dictionary<string, FamilySymbol>();
            tags.Add("Areas", curAreaTag);
            tags.Add("Curtain Walls", curCwallTag);
            tags.Add("Doors", curDoorTag);
            tags.Add("Furniture", curFurnTag);
            tags.Add("Lighting Fixtures", curLightTag);
            tags.Add("Rooms", curRoomTag);
            tags.Add("Walls", curWallTag);
            tags.Add("Windows", curWindowTag);

            //create counter
            int counter = 0;

            ///////////////////view type
            ViewType curViewType = curView.ViewType;


            if (curViewType == ViewType.FloorPlan)
            {
                //floorplan cat list
                ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catListFP);
                Collector.WherePasses(catfilter).WhereElementIsNotElementType();

                foreach (Element curElem in Collector)
                {
                    //get location point, or midpoint
                    XYZ insPoint = Utils.GetTagLocationPoint(curElem);

                    //get the correct tag based on the dictionary, and therefore category of the element
                    FamilySymbol curTagType = tags[curElem.Category.Name];
                    Reference curRef = new Reference(curElem);

                    //tag; Curtain walls, doors, furniture, rooms,  walls, and windows 
                    //M_Curtain Wall Tag, M_Door Tag, M_Furniture Tag, M_Room Tag, M_Wall Tag, M_Window Tag
                    // window tag offset 3' in y direction


                    if (curElem.Category.Name == "Walls")
                    {
                        Wall curWall = curElem as Wall;
                        WallType curWallType = curWall.WallType;
                        if (curWallType.Kind == WallKind.Curtain)
                        {
                            //create curtain wall tag
                            IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
                            curRef, true, TagOrientation.Horizontal, insPoint);
                        }
                        else
                        {
                            IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
                            curRef, true, TagOrientation.Horizontal, insPoint);
                            //create regular wall tag
                        }
                    }
                    else
                    {
                        IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
                        curRef, false, TagOrientation.Horizontal, insPoint);
                    }

                }
            }
            else if (curViewType == ViewType.CeilingPlan)
            {   //tag; Light Fixtures, and rooms. "M_Light Fixture Tag", "M_Room Tag"
                //ceilingplan cat list
                ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catListCP);
                Collector.WherePasses(catfilter).WhereElementIsNotElementType();

                foreach (Element curElem in Collector)
                {
                    //get location point, or midpoint
                    XYZ insPoint = Utils.GetTagLocationPoint(curElem);
                    //get the correct tag based on the dictionary, and therefore category of the element
                    FamilySymbol curTagType = tags[curElem.Category.Name];
                    Reference curRef = new Reference(curElem);

                }
            }

            else if (curViewType == ViewType.AreaPlan)
            {
                //ceilingplan cat list
                ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catListAP);
                Collector.WherePasses(catfilter).WhereElementIsNotElementType();

                foreach (Element curElem in Collector)
                {
                    //get location point, or midpoint
                    XYZ insPoint = Utils.GetTagLocationPoint(curElem);

                    //get the correct tag based on the dictionary, and therefore category of the element
                    FamilySymbol curTagType = tags[curElem.Category.Name];
                    Reference curRef = new Reference(curElem);

                    if (curElem.Category.Name == "Areas")
                    {
                        ViewPlan curAreaPlan = curView as ViewPlan;
                        Area curArea = curElem as Area;
                        AreaTag newAreaTag = doc.Create.NewAreaTag(curAreaPlan, curArea, new UV(insPoint.X, insPoint.Y));
                        newAreaTag.TagHeadPosition = new XYZ(insPoint.X, insPoint.Y, 0);
                        newAreaTag.HasLeader = false;
                    }
                    else continue;
                }
            }

            else if (curViewType == ViewType.Section)
            {
                ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catListS);
                Collector.WherePasses(catfilter).WhereElementIsNotElementType();
                foreach (Element curElem in Collector)
                {
                    //get location point, or midpoint
                    XYZ insPoint = Utils.GetTagLocationPoint(curElem);

                    //get the correct tag based on the dictionary, and therefore category of the element
                    FamilySymbol curTagType = tags[curElem.Category.Name];
                    Reference curRef = new Reference(curElem);

                }
            }

            else
            {
                TaskDialog.Show("Error", "This tool only works for areaplans ,ceilingplans, floorplans and sections");
            }    
            
            counter++;
            TaskDialog.Show("Number of Tags", "There are " + counter + " number of tags created!");


            return Result.Succeeded;
        }  

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "MOD0202Tag";
            string buttonTitle = "MOD0202Tag";

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
