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
    public class MOD0202 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            //1a. filtered element collect by view
            View curView = doc.ActiveView;
            FilteredElementCollector collector = new FilteredElementCollector(doc, curView.Id);

            //1b. element multicategory filters
            List<BuiltInCategory> catList = new List<BuiltInCategory>();
            catList.Add(BuiltInCategory.OST_Rooms);
            catList.Add(BuiltInCategory.OST_Doors);
            catList.Add(BuiltInCategory.OST_Walls);

            ElementMulticategoryFilter catfilter = new ElementMulticategoryFilter(catList);

            collector.WherePasses(catfilter).WhereElementIsNotElementType();

            //LINQ
            FamilySymbol curDoorTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Door Tag"))
                .First();

            FamilySymbol curRoomTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Room Tag"))
                .First();

            FamilySymbol curWallTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Wall Tag"))
                .First();



            //Dictionaries
            //2. create dictionary for tags
            Dictionary<string, FamilySymbol> tags = new Dictionary<string, FamilySymbol>();
            tags.Add("Doors", curDoorTag);
            tags.Add("Rooms", curRoomTag);
            tags.Add("Walls", curWallTag);


            //why use one? Fast, no loops, very efficient 
            using (Transaction t = new Transaction(doc))
            {
                t.Start("insert Tags");
                foreach (Element curElem in collector)
                {
                    //3. get point from location
                    XYZ insPoint;
                    LocationPoint locPoint;
                    LocationCurve locCurve;

                    Location curLoc = curElem.Location;
                    if (curLoc == null)
                        continue;

                    locPoint = curLoc as LocationPoint;
                    if (locPoint != null)
                    {
                        //is a location point
                        insPoint = locPoint.Point;
                    }
                    else
                    {
                        // is a location curve
                        locCurve = curLoc as LocationCurve;
                        Curve curCurve = locCurve.Curve;
                        //insPoint = curCurve.GetEndPoint(1);
                        insPoint = Utils.GetMidpointBetweenTwoPoints(curCurve.GetEndPoint(0), curCurve.GetEndPoint(1));
                    }


                    // 4. create reference to element 
                    Reference curRef = new Reference(curElem);
                    FamilySymbol curTagType = tags[curElem.Category.Name];

                    //5a. place tag
                    IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id, 
                        curRef, false,TagOrientation.Horizontal, insPoint);
                    
                }
                t.Commit();
            }

            return Result.Succeeded;
        }



        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "MOD0202";
            string buttonTitle = "MOD0202";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.icons8_tag_pulsar_color_3296,
                Properties.Resources.icons8_tag_pulsar_color_1696,
                "This is a tooltip for Button 2");

            return myButtonData1.Data;
        }
    }
}
