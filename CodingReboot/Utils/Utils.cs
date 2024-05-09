using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            Collector.OfClass(typeof(Level));Collector.WhereElementIsElementType();

            foreach (Level curlevel in Collector)
            {
                if(curlevel.Name == levelName)
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

        

    }
}
