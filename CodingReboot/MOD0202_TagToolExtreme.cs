#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        //public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        //{
        //    // this is a variable for the Revit application
        //    UIApplication uiapp = commandData.Application;

        //    // this is a variable for the current Revit model
        //    Document doc = uiapp.ActiveUIDocument.Document;

        //    // Your code goes here

        //    //View curView = doc.ActiveView;

        //    FilteredElementCollector ViewCollector = new FilteredElementCollector(doc);
        //    ViewCollector.OfCategory(BuiltInCategory.OST_Views);
        //    ViewCollector.WhereElementIsNotElementType();

        //    //filtered out the viewtemplates

        //    //4. loop through elemsnt to tag
        //    int counter = 0;

        //    foreach (View curView in ViewCollector)
        //    {                             
        //        ViewType curViewType = curView.ViewType;

        //        //2. get the categories for the view type 
        //        List<BuiltInCategory> catList = new List<BuiltInCategory>();
        //        Dictionary<ViewType, List<BuiltInCategory>> viewTypeCatDic = GetViewTypeCatDictionary();

                

        //        //if the view type is of one you dont want, you can do the tryget to test and get at the same time
        //        if (viewTypeCatDic.TryGetValue(curViewType, out catList) == false)
        //        {
        //            continue;
        //        }

        //        //3. get elements to tag for the view type
        //        ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catList);
        //        FilteredElementCollector elemCollector = new FilteredElementCollector(doc, curView.Id);
        //        elemCollector.WherePasses(catfilter).WhereElementIsNotElementType();

                

        //        //Dictionary
        //        //6. create dictionary for tags
        //        Dictionary<string, FamilySymbol> tagdictionary = GetTagDictionary(doc);

        //        using (Transaction t = new Transaction(doc))
        //        {

        //            t.Start("tag elements in view");

        //            foreach (Element curElem in elemCollector)
        //            {
        //                bool addleader = false;

        //                if (curElem.Location == null)
        //                    continue;
        //                //5. get insertion point based on element type 
        //                XYZ point = Utils.GetTagLocationPoint(curElem.Location);
        //                if (point == null)
        //                    continue;

        //                // 7. get element data
        //                string catName = curElem.Category.Name;

        //                //10. check cat name for walls
        //                if (catName == "Walls")
        //                {
        //                    addleader = true;
        //                    if (IsCurtainWall(curElem))
        //                    {
        //                        catName = "Curtain Walls";
        //                    }
        //                }

        //                //8. get tag based on element type
        //                FamilySymbol elemTag = tagdictionary[catName];

        //                //9.tag element
        //                if (catName == "Areas")
        //                {
        //                    ViewPlan curareaplan = curView as ViewPlan;
        //                    Area curArea = curElem as Area;

        //                    AreaTag newTag = doc.Create.NewAreaTag(curareaplan, curArea, new UV(point.X, point.Y));
        //                    newTag.TagHeadPosition = new XYZ(point.X, point.Y, 0);
        //                }
        //                else
        //                {
        //                    IndependentTag newtag = IndependentTag.Create(doc, elemTag.Id, curView.Id, new Reference(curElem),
        //                        addleader, TagOrientation.Horizontal, point);
        //                    //9a. offset tags as needed
        //                    if (catName == "Windows")
        //                        newtag.TagHeadPosition = point.Add(new XYZ(0, 3, 0));

        //                    if (curView.ViewType == ViewType.Section)
        //                        newtag.TagHeadPosition = point.Add(new XYZ(0, 0, 3));

        //                }

        //                counter++;


        //            }
        //            t.Commit();

        //        }

        //        TaskDialog.Show("Complete", $"Added {counter} tags in view");
        //    }
        //}


        //internal static PushButtonData GetButtonData()
        //{
        //    // use this method to define the properties for this command in the Revit ribbon
        //    string buttonInternalName = "MOD0202TagExtreme";
        //    string buttonTitle = "TagEXTREME";

        //    ButtonDataClass myButtonData1 = new ButtonDataClass(
        //        buttonInternalName,
        //        buttonTitle,
        //        MethodBase.GetCurrentMethod().DeclaringType?.FullName,
        //        Properties.Resources.icons8_color_palette_pulsar_gradient_3296,
        //        Properties.Resources.icons8_color_palette_pulsar_gradient_1696,
        //        "Tag all the views with all the tags");

        //    return myButtonData1.Data;
        //}

    }
}  
    
