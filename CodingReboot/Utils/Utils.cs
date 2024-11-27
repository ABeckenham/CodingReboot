using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;

namespace CodingReboot
{
    internal static class Utils
    {
        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel currentPanel = GetRibbonPanelByName(app, tabName, panelName);

            if (currentPanel == null)
                currentPanel = app.CreateRibbonPanel(tabName, panelName);

            return currentPanel;
        }

        internal static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }

        internal static Level GetLevelByName(Document doc, string levelName)
        {
            FilteredElementCollector Collector = new FilteredElementCollector(doc);
            Collector.OfClass(typeof(Level)); Collector.WhereElementIsElementType();

            foreach (Level curlevel in Collector)
            {
                if (curlevel.Name == levelName)
                {
                    return curlevel;
                }
            }
            return null;
        }

        internal static Element GetWallTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector Collector = new FilteredElementCollector(doc);
            Collector.OfClass(typeof(WallType));

            foreach (Element curWall in Collector)
            {
                if (curWall.Name == typeName)
                {
                    return curWall;
                }
            }
            return null;
        }


        internal static string GetParameterValueAsString(Element element, string paramName)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter myParam = paramList.First();

            return myParam.AsString();
        }

        internal static double GetParameterValueAsDouble(Element element, string paramName)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter Param = paramList.First();

            return Param.AsDouble();
        }

        internal static void SetParameterValue(Element element, string paramName, string value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter Param = paramList.First();

            Param.Set(value);
        }

        internal static void SetParameterValue(Element element, string paramName, double value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter Param = paramList.First();

            Param.Set(value);
        }

        internal static void SetParameterValue(Element element, string paramName, int value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter Param = paramList.First();

            Param.Set(value);
        }


        internal static FamilySymbol GetFamilySymbolByName(Document doc, string famName, string fsName)
        {//equlivant of a family type
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));

            foreach (FamilySymbol fs in collector)
            {
                if (fs.Name == fsName && fs.FamilyName == famName)
                    return fs;
            }
            return null;
        }

        internal static FilteredElementCollector GetAllRooms(Document doc)
        {
            FilteredElementCollector RoomCollector = new FilteredElementCollector(doc);
            RoomCollector.OfCategory(BuiltInCategory.OST_Rooms);

            return RoomCollector;
        }

        internal static XYZ GetMidpointBetweenTwoPoints(XYZ point1, XYZ point2)
        {
            XYZ midPoint = new XYZ(
                (point1.X + point2.X) / 2,
                (point1.Y + point2.Y) / 2,
                (point1.Z + point2.Z) / 2);
            return midPoint;
        }

        internal static FamilySymbol GetTagbyName(Document doc, string tagName)
        {
            FamilySymbol curTag = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilySymbol))
                            .Cast<FamilySymbol>()
                            .Where(x => x.FamilyName.Equals(tagName))
                            .First();

            return curTag;
        }


        internal static XYZ GetTagLocationPoint(Location loc)
        {
            XYZ insPoint;
            //get the location
            
            if (loc == null)
                return null;
            else
            {
                LocationPoint locPoint = loc as LocationPoint;
                if (locPoint != null)
                {
                    insPoint = locPoint.Point;
                }
                else
                {
                    LocationCurve locCurve = loc as LocationCurve;
                    Curve curCurve = locCurve.Curve;
                    insPoint = Utils.GetMidpointBetweenTwoPoints(curCurve.GetEndPoint(0), curCurve.GetEndPoint(1));
                }
            }
            //need some additional logic for curved walls 
            return insPoint;
        }
        internal static bool IsCurtainWall(Element curWallElem)
        {
            Wall curWall = curWallElem as Wall;

            if (curWall.WallType.Kind == WallKind.Curtain)
            {
                return true;
            }
            else return false;
        }

        internal static Dictionary<string, FamilySymbol> GetTagDictionary(Document doc)
        {
            Dictionary<string, FamilySymbol> tagDict = new Dictionary<string, FamilySymbol>();
            tagDict.Add("Areas", Utils.GetTagbyName(doc, "M_Area Tag"));
            tagDict.Add("Curtain Walls", Utils.GetTagbyName(doc, "M_Curtain Wall Tag"));
            tagDict.Add("Doors", Utils.GetTagbyName(doc, "M_Door Tag"));
            tagDict.Add("Furniture", Utils.GetTagbyName(doc, "M_Furniture Tag"));
            tagDict.Add("Lighting Fixtures", Utils.GetTagbyName(doc, "M_Lighting Fixture Tag"));
            tagDict.Add("Rooms", Utils.GetTagbyName(doc, "M_Room Tag"));
            tagDict.Add("Walls", Utils.GetTagbyName(doc, "M_Wall Tag"));
            tagDict.Add("Windows", Utils.GetTagbyName(doc, "M_Window Tag"));

            return tagDict;

        }

        internal static Dictionary<ViewType, List<BuiltInCategory>> GetViewTypeCatDictionary()
        {
            Dictionary<ViewType, List<BuiltInCategory>> dictionary = new Dictionary<ViewType, List<BuiltInCategory>>();

            dictionary.Add(ViewType.FloorPlan, new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Rooms,
                BuiltInCategory.OST_Windows,
                BuiltInCategory.OST_Doors,
                BuiltInCategory.OST_Furniture,
                BuiltInCategory.OST_Walls
            });

            dictionary.Add(ViewType.AreaPlan, new List<BuiltInCategory> { BuiltInCategory.OST_Areas });
            dictionary.Add(ViewType.CeilingPlan, new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Rooms,
                BuiltInCategory.OST_LightingFixtures

            });

            dictionary.Add(ViewType.Section, new List<BuiltInCategory> { BuiltInCategory.OST_Rooms });
            return dictionary;
        }

        internal static bool IsLineVertical(Curve curLine)
        {
            XYZ p1 = curLine.GetEndPoint(0);
            XYZ p2 = curLine.GetEndPoint(1);
            if (Math.Abs(p1.X - p2.X) < Math.Abs(p1.Y - p2.Y))
                return true;
            return false;
        }


        public static class Convert
        {
            public static double ConvertMMtoFT(double mmDim)
            {
                //convert millimeters to feet
                double convert = (mmDim / 25.4) / 12;

                return convert;
            }

            public static double ConvertCMtoFT(double cmDim)
            {
                //convert centimeters to feet
                double convert = (cmDim / 2.54) / 12;

                return convert;
            }

            public static double ConvertMtoFT(double mDim)
            {
                //convert meters to feet
                return mDim * 3.28084;
            }

            public static Document OpenDocumentFile(string filePath, bool audit  = false, bool detachfromCentral = false)
            {
                return null; 
            }
        }

            
    }

}
