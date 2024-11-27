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
using System.Windows.Controls;

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
            ///////////////////view type
            ViewType curViewType = curView.ViewType;


            //List<BuiltInCategory> catList = new List<BuiltInCategory>();
            //catList.Add(BuiltInCategory.OST_Areas);
            //catList.Add(BuiltInCategory.OST_Doors);
            //catList.Add(BuiltInCategory.OST_Furniture);
            //catList.Add(BuiltInCategory.OST_LightingFixtures);
            //catList.Add(BuiltInCategory.OST_Rooms);
            //catList.Add(BuiltInCategory.OST_Walls);

            //ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catList);
            //Collector.WherePasses(catfilter).WhereElementIsNotElementType();

            //List<BuiltInCategory> catListFP = new List<BuiltInCategory>();
            //catListFP.Add(BuiltInCategory.OST_Doors);
            //catListFP.Add(BuiltInCategory.OST_Furniture);
            //catListFP.Add(BuiltInCategory.OST_LightingFixtures);
            //catListFP.Add(BuiltInCategory.OST_Rooms);
            //catListFP.Add(BuiltInCategory.OST_Walls);
            //catListFP.Add(BuiltInCategory.OST_Windows);

            ////could i make multiple collectors for each view then run the script on which view it is??
            //  //make multiple catfilters for each view type. 

            //List<BuiltInCategory> catListCP = new List<BuiltInCategory>();
            //catListCP.Add(BuiltInCategory.OST_LightingFixtures);
            //catListCP.Add(BuiltInCategory.OST_Rooms);

            //List<BuiltInCategory> catListAP = new List<BuiltInCategory>();
            //catListAP.Add(BuiltInCategory.OST_Areas);

            //List<BuiltInCategory> catListS = new List<BuiltInCategory>();
            //catListS.Add(BuiltInCategory.OST_Rooms);

            //** tips from runthrough - do these Lists inside the loop!
            //** He created a dictionary for the viewtypes!!


            //2. get the categories for the view type 
            List<BuiltInCategory> catList = new List<BuiltInCategory>();
            Dictionary<ViewType, List<BuiltInCategory>> viewTypeCatDic = Utils.GetViewTypeCatDictionary();

            //if the view type is of one you dont want, you can do the tryget to test and get at the same time
            if(viewTypeCatDic.TryGetValue(curViewType, out catList ) == false)
            {
                TaskDialog.Show("Error", "sorry. Cannot add tags to this view type.");
                return Result.Failed;
            }

            //3. get elements to tag for the view type
            ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catList);
            FilteredElementCollector elemCollector = new FilteredElementCollector(doc, curView.Id);
            elemCollector.WherePasses(catfilter).WhereElementIsNotElementType();

            //4. loop through elemsnt to tag
            int counter = 0;

            //Dictionary
            //6. create dictionary for tags
            Dictionary<string, FamilySymbol> tagdictionary = Utils.GetTagDictionary(doc);

            using (Transaction t = new Transaction(doc))
            {

                t.Start("tag elements in view");

                foreach (Element curElem in elemCollector)
                {
                    bool addleader = false;

                    if (curElem.Location == null)
                        continue;
                    //5. get insertion point based on element type 
                    XYZ point = Utils.GetTagLocationPoint(curElem.Location);
                    if (point == null)
                        continue;

                    // 7. get element data
                    string catName = curElem.Category.Name;

                    //10. check cat name for walls
                    if(catName == "Walls")
                    {
                        addleader = true;
                        if(Utils.IsCurtainWall(curElem))
                        {
                            catName = "Curtain Walls";
                        }
                    }

                    //8. get tag based on element type
                    FamilySymbol elemTag = tagdictionary[catName];

                    //9.tag element
                    if (catName == "Areas")
                    {
                        ViewPlan curareaplan = curView as ViewPlan;
                        Area curArea = curElem as Area;

                        AreaTag newTag = doc.Create.NewAreaTag(curareaplan,curArea,new UV(point.X, point.Y));
                        newTag.TagHeadPosition = new XYZ(point.X, point.Y, 0);
                    }
                    else
                    {
                        IndependentTag newtag = IndependentTag.Create(doc, elemTag.Id, curView.Id, new Reference(curElem),
                            addleader, TagOrientation.Horizontal, point);
                        //9a. offset tags as needed
                        if(catName == "Windows")                        
                            newtag.TagHeadPosition = point.Add(new XYZ(0, 3, 0));

                        if(curView.ViewType == ViewType.Section)
                            newtag.TagHeadPosition = point.Add(new XYZ(0,0,3));                           
                        
                    }

                    counter++;


                }
                t.Commit();
                
            }

            TaskDialog.Show("Complete", $"Added {counter} tags in view");

            //using (Transaction t = new Transaction(doc))
            //{
            //    t.Start("insert Tags");

            //    if (curViewType == ViewType.FloorPlan)
            //    {
            //        //floorplan cat list
            //        ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catListFP);
            //        Collector.WherePasses(catfilter).WhereElementIsNotElementType();

            //        foreach (Element curElem in Collector)
            //        {
            //            //tag; Curtain walls, doors, furniture, rooms,  walls, and windows 
            //            //M_Curtain Wall Tag, M_Door Tag, M_Furniture Tag, M_Room Tag, M_Wall Tag, M_Window Tag
            //            // window tag offset 3' in y direction
            //            //get location point, or midpoint
            //            XYZ insPoint = Utils.GetTagLocationPoint(curElem.Location);
            //            if (insPoint == null)
            //                continue;
            //            //get the correct tag based on the dictionary, and therefore category of the element
            //            FamilySymbol curTagType = tags[curElem.Category.Name];
            //            Reference curRef = new Reference(curElem);

            //            if (curElem.Category.Name == "Walls")
            //            {
            //                Wall curWall = curElem as Wall;
            //                WallType curWallType = curWall.WallType;
            //                if (curWallType.Kind == WallKind.Curtain)
            //                {
            //                    //create curtain wall tag
            //                    IndependentTag newTag = IndependentTag.Create(doc, curCwallTag.Id, curView.Id,
            //                    curRef, true, TagOrientation.Horizontal, insPoint);
            //                    counter++;
            //                }
            //                else
            //                {
            //                    //create regular wall tag
            //                    IndependentTag newTag = IndependentTag.Create(doc, curWallTag.Id, curView.Id,
            //                    curRef, true, TagOrientation.Horizontal, insPoint);
            //                    counter++;
            //                }
            //            }

            //            else if (curElem.Category.Name == "Windows")
            //            {
            //                IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
            //                curRef, false, TagOrientation.Horizontal, insPoint);
            //                newTag.TagHeadPosition = new XYZ(insPoint.X, (insPoint.Y + 3), 0);
            //                newTag.HasLeader = false;
            //                counter++;
            //            }

            //            else
            //            {
            //                IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
            //                curRef, false, TagOrientation.Horizontal, insPoint);
            //                newTag.HasLeader = false;
            //                counter++;
            //            }

            //        }
            //    }

            //    else if (curViewType == ViewType.CeilingPlan)
            //    {   //tag; Light Fixtures, and rooms. "M_Light Fixture Tag", "M_Room Tag"
            //        //ceilingplan cat list
            //        ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catListCP);
            //        Collector.WherePasses(catfilter).WhereElementIsNotElementType();

            //        foreach (Element curElem in Collector)
            //        {
            //            //get location point, or midpoint
            //            XYZ insPoint = Utils.GetTagLocationPoint(curElem.Location);
            //            if(insPoint == null)
            //             continue; 
            //            //get the correct tag based on the dictionary, and therefore category of the element
            //            FamilySymbol curTagType = tags[curElem.Category.Name];
            //            Reference curRef = new Reference(curElem);

            //            IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
            //                    curRef, true, TagOrientation.Horizontal, insPoint);
            //            newTag.HasLeader = false;
            //            newTag.TagHeadPosition = new XYZ(insPoint.X, insPoint.Y, 0);
            //            counter++;
            //        }
            //    }

            //    else if (curViewType == ViewType.AreaPlan)
            //    {
            //        //ceilingplan cat list
            //        ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catListAP);
            //        Collector.WherePasses(catfilter).WhereElementIsNotElementType();

            //        foreach (Element curElem in Collector)
            //        {
            //            //get location point, or midpoint
            //            XYZ insPoint = Utils.GetTagLocationPoint(curElem.Location);
            //            if (insPoint == null)
            //                continue;
            //            //get the correct tag based on the dictionary, and therefore category of the element
            //            FamilySymbol curTagType = tags[curElem.Category.Name];
            //            Reference curRef = new Reference(curElem);

            //            if (curElem.Category.Name == "Areas")
            //            {
            //                ViewPlan curAreaPlan = curView as ViewPlan;
            //                Area curArea = curElem as Area;
            //                AreaTag newAreaTag = doc.Create.NewAreaTag(curAreaPlan, curArea, new UV(insPoint.X, insPoint.Y));
            //                newAreaTag.TagHeadPosition = new XYZ(insPoint.X, insPoint.Y, 0);
            //                newAreaTag.HasLeader = false;
            //                counter++;
            //            }
            //            else continue;
            //        }
            //    }

            //    else if (curViewType == ViewType.Section)
            //    {
            //        ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catListS);
            //        Collector.WherePasses(catfilter).WhereElementIsNotElementType();
            //        foreach (Element curElem in Collector)
            //        {
            //            //get location point, or midpoint
            //            XYZ insPoint = Utils.GetTagLocationPoint(curElem.Location);
            //            if (insPoint == null)
            //                continue;
            //            //get the correct tag based on the dictionary, and therefore category of the element
            //            FamilySymbol curTagType = tags[curElem.Category.Name];
            //            Reference curRef = new Reference(curElem);

            //            IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
            //                    curRef, true, TagOrientation.Horizontal, insPoint);
            //            newTag.HasLeader = false;
            //            newTag.TagHeadPosition = new XYZ(insPoint.X, insPoint.Y, (insPoint.Z+3));
            //            counter++;
            //        }
            //    }

            //    else
            //    {
            //        TaskDialog.Show("Error", "This tool only works for areaplans ,ceilingplans, floorplans and sections");
            //    }
                                
            //    TaskDialog.Show("Number of Tags", "There are " + counter + " number of tags created!");

            //    t.Commit();

            //}

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
                Properties.Resources.icons8_dog_tag_pulsar_color_3296,
                Properties.Resources.icons8_dog_tag_pulsar_color_1696,
                "Tag the current view with the appropriate tags");

            return myButtonData1.Data;
        }
    }
}
