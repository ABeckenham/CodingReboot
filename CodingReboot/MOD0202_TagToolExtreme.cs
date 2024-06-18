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
using System.Xml;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD0202_TagToolExtreme : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            //View curView = doc.ActiveView;
            

            FilteredElementCollector ViewCollector = new FilteredElementCollector(doc);
            ViewCollector.OfCategory(BuiltInCategory.OST_Views)
            .WhereElementIsNotElementType();
                                  
            
            List<BuiltInCategory> catListFP = new List<BuiltInCategory>();
            catListFP.Add(BuiltInCategory.OST_Doors);
            catListFP.Add(BuiltInCategory.OST_Furniture);
            catListFP.Add(BuiltInCategory.OST_LightingFixtures);
            catListFP.Add(BuiltInCategory.OST_Rooms);
            catListFP.Add(BuiltInCategory.OST_Walls);
            catListFP.Add(BuiltInCategory.OST_Windows);

            //could i make multiple collectors for each view then run the script on which view it is??
              //make multiple catfilters for each view type. 

            List<BuiltInCategory> catListCP = new List<BuiltInCategory>();
            catListCP.Add(BuiltInCategory.OST_LightingFixtures);
            catListCP.Add(BuiltInCategory.OST_Rooms);

            List<BuiltInCategory> catListAP = new List<BuiltInCategory>();
            catListAP.Add(BuiltInCategory.OST_Areas);

            List<BuiltInCategory> catListS = new List<BuiltInCategory>();
            catListS.Add(BuiltInCategory.OST_Rooms);

            

            //LINQ
            FamilySymbol curDoorTag = Utils.GetTagbyName(doc, "M_Door Tag");
            FamilySymbol curAreaTag = Utils.GetTagbyName(doc, "M_Area Tag");                   
            FamilySymbol curCwallTag = Utils.GetTagbyName(doc, "M_Curtain Wall Tag");
            FamilySymbol curWallTag = Utils.GetTagbyName(doc, "M_Wall Tag");
            FamilySymbol curFurnTag = Utils.GetTagbyName(doc, "M_Furniture Tag");
            FamilySymbol curLightTag = Utils.GetTagbyName(doc, "M_Lighting Fixture Tag");
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
            int viewcounter = 0;

            ///////////////////view type
            //ViewType curViewType = curView.ViewType;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("insert Tags");

                foreach (View curView in ViewCollector)
                {
                    
                    ViewType curViewType = curView.ViewType;
                    FilteredElementCollector Collector = new FilteredElementCollector(doc, curView.Id);

                    if (curViewType == ViewType.ThreeD)
                        continue;                    

                    viewcounter++;
                    if (curViewType == ViewType.FloorPlan)
                    {
                        //floorplan cat list
                        ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catListFP);
                        Collector.WherePasses(catfilter).WhereElementIsNotElementType();

                        foreach (Element curElem in Collector)
                        {
                            //tag; Curtain walls, doors, furniture, rooms,  walls, and windows 
                            //M_Curtain Wall Tag, M_Door Tag, M_Furniture Tag, M_Room Tag, M_Wall Tag, M_Window Tag
                            // window tag offset 3' in y direction
                            //get location point, or midpoint
                            XYZ insPoint = Utils.GetTagLocationPoint(curElem);
                            if (insPoint == null)
                                continue;
                            //get the correct tag based on the dictionary, and therefore category of the element
                            FamilySymbol curTagType = tags[curElem.Category.Name];
                            Reference curRef = new Reference(curElem);

                            if (curElem.Category.Name == "Walls")
                            {
                                Wall curWall = curElem as Wall;
                                WallType curWallType = curWall.WallType;
                                if (curWallType.Kind == WallKind.Curtain)
                                {
                                    //create curtain wall tag
                                    IndependentTag newTag = IndependentTag.Create(doc, curCwallTag.Id, curView.Id,
                                    curRef, true, TagOrientation.Horizontal, insPoint);
                                    counter++;
                                }
                                else
                                {
                                    //create regular wall tag
                                    IndependentTag newTag = IndependentTag.Create(doc, curWallTag.Id, curView.Id,
                                    curRef, true, TagOrientation.Horizontal, insPoint);
                                    counter++;
                                }
                            }

                            else if (curElem.Category.Name == "Windows")
                            {
                                IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
                                curRef, false, TagOrientation.Horizontal, insPoint);
                                newTag.TagHeadPosition = new XYZ(insPoint.X, (insPoint.Y + 3), 0);
                                newTag.HasLeader = false;
                                counter++;
                            }

                            else
                            {
                                IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
                                curRef, false, TagOrientation.Horizontal, insPoint);
                                newTag.HasLeader = false;
                                counter++;
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
                            if (insPoint == null)
                                continue;
                            //get the correct tag based on the dictionary, and therefore category of the element
                            FamilySymbol curTagType = tags[curElem.Category.Name];
                            Reference curRef = new Reference(curElem);

                            IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
                                    curRef, true, TagOrientation.Horizontal, insPoint);
                            newTag.HasLeader = false;
                            newTag.TagHeadPosition = new XYZ(insPoint.X, insPoint.Y, 0);
                            counter++;
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
                            if (insPoint == null)
                                continue;
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
                                counter++;
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
                            if (insPoint == null)
                                continue;
                            //get the correct tag based on the dictionary, and therefore category of the element
                            FamilySymbol curTagType = tags[curElem.Category.Name];
                            Reference curRef = new Reference(curElem);

                            IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
                                    curRef, true, TagOrientation.Horizontal, insPoint);
                            newTag.HasLeader = false;
                            newTag.TagHeadPosition = new XYZ(insPoint.X, insPoint.Y, (insPoint.Z + 3));
                            counter++;
                        }
                    }

                    else continue;
                    

                }    TaskDialog.Show("Number of Tagged views", viewcounter + " Views have been tagged" + Environment.NewLine +
                     "There are " + counter + " number of tags created!");

                t.Commit();

            }

            return Result.Succeeded;
        }  

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "MOD0202TagExtreme";
            string buttonTitle = "TagEXTREME";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.icons8_color_palette_pulsar_gradient_3296,
                Properties.Resources.icons8_color_palette_pulsar_gradient_1696,
                "Tag all the views with all the tags");

            return myButtonData1.Data;
        }
    }
}
