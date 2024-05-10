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
    public class MOD0201 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Door schedule created");
                //01.schedule
                ElementId catId = new ElementId(BuiltInCategory.OST_Doors);
                ViewSchedule newschedule = ViewSchedule.CreateSchedule(doc, catId);
                newschedule.Name = "My Door Schedule";

                //02a Get parameters for fields
                FilteredElementCollector doorCollector = new FilteredElementCollector(doc);
                doorCollector.OfCategory(BuiltInCategory.OST_Doors);
                doorCollector.WhereElementIsNotElementType();

                Element doorInst = doorCollector.FirstElement();

                Parameter doorNumParam = doorInst.LookupParameter("Mark");
                Parameter doorLevelParam = doorInst.LookupParameter("Level");

                //BUILT-IN PARAMETERS - use revitlookup to see what sort of parameter it is
                Parameter doorWidthParam = doorInst.get_Parameter(BuiltInParameter.DOOR_WIDTH);
                Parameter doorHeightParam = doorInst.get_Parameter(BuiltInParameter.DOOR_HEIGHT);
                Parameter doorTypeParam = doorInst.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM);


                //02b. create fields
                ScheduleField doorNumField = newschedule.Definition.AddField(ScheduleFieldType.Instance, doorNumParam.Id);
                ScheduleField doorLevelField = newschedule.Definition.AddField(ScheduleFieldType.Instance, doorLevelParam.Id);
                ScheduleField doorHeightField = newschedule.Definition.AddField(ScheduleFieldType.ElementType, doorHeightParam.Id);
                ScheduleField doorWidthField = newschedule.Definition.AddField(ScheduleFieldType.ElementType, doorWidthParam.Id);
                ScheduleField doorTypeField = newschedule.Definition.AddField(ScheduleFieldType.Instance, doorTypeParam.Id);

                doorLevelField.IsHidden = true;
                doorWidthField.DisplayType = ScheduleFieldDisplayType.Totals;

                //03. filter by level
                Level filterLevel = Utils.GetLevelByName(doc, "Level 1");
                ScheduleFilter LevelFilter = new ScheduleFilter(doorLevelField.FieldId, ScheduleFilterType.Equal, filterLevel.Id);
                newschedule.Definition.AddFilter(LevelFilter);

                //04. Grouping 
                ScheduleSortGroupField typeSort = new ScheduleSortGroupField(doorTypeField.FieldId);
                typeSort.ShowHeader = true;
                typeSort.ShowFooter = true;
                typeSort.ShowBlankLine = true;
                newschedule.Definition.AddSortGroupField(typeSort);

                //05. sorting 
                ScheduleSortGroupField markSort = new ScheduleSortGroupField(doorNumField.FieldId);
                newschedule.Definition.AddSortGroupField(markSort);

                // 05.set totals
                newschedule.Definition.IsItemized = true;
                newschedule.Definition.ShowGrandTotal = true;
                newschedule.Definition.ShowGrandTotalTitle = true;
                newschedule.Definition.ShowGrandTotalCount = true;

                t.Commit();
            }

            //06. c# filter a list for uniquw items
            List<string> rawStrings = new List<string>() { "a", "b", "c", "c", "d", "e", "e", "f", "g", "h" };
            List<string> uniqueStrings = rawStrings.Distinct().ToList();
            uniqueStrings.Sort();

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "MOD0201";
            string buttonTitle = "MOD0201";

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
