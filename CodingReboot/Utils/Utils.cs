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


        internal static XYZ GetTagLocationPoint(Element curElem)
        {
            XYZ insPoint;
            LocationPoint locPoint;
            LocationCurve locCurve;

            //get the location
            Location curLoc = curElem.Location;
            if (curLoc == null)
                return null;
            else
            {
                locPoint = curLoc as LocationPoint;
                if (locPoint != null)
                {
                    insPoint = locPoint.Point;
                }
                else
                {
                    locCurve = curLoc as LocationCurve;
                    Curve curCurve = locCurve.Curve;
                    insPoint = Utils.GetMidpointBetweenTwoPoints(curCurve.GetEndPoint(0), curCurve.GetEndPoint(1));
                }
            }

            return insPoint;
        }

    }
}
