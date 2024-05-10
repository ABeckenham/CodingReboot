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
using System.Reflection;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD0201_Schedules : IExternalCommand
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
                t.Start("Creating Door Schedule");

                // create a schedule of Rooms

                //get all the rooms to find all departs of the rooms 
                FilteredElementCollector allRooms = Utils.GetAllRooms(doc);

                //Get all department names as list 
                List<string> roomDeptList = new List<string>();
                foreach (Room curRoom in allRooms)
                {
                    string curDept = Utils.GetParameterValueAsString(curRoom, "Department");
                    roomDeptList.Add(curDept);
                }
                              

                //get all room parameters needed : Name, number, department, comments, area, level
                Element roomInst = allRooms.FirstElement();                
                Parameter roomNumberParam = roomInst.get_Parameter(BuiltInParameter.ROOM_NUMBER);
                Parameter roomNameParam = roomInst.get_Parameter(BuiltInParameter.ROOM_NAME);
                Parameter roomDeptParam = roomInst.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT);
                Parameter roomCommParam = roomInst.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                Parameter roomAreaParam = roomInst.get_Parameter(BuiltInParameter.ROOM_AREA);
                Parameter roomLevelParam = roomInst.get_Parameter(BuiltInParameter.LEVEL_NAME);
                                   
                //create schedule for rooms
                foreach(string curString in roomDeptList)
                {
                    ElementId catId = new ElementId(BuiltInCategory.OST_Rooms);
                    ViewSchedule newschedule = ViewSchedule.CreateSchedule(doc, catId);
                    newschedule.Name = "Dept - " + curString;

                    //create schedule Fields
                    ScheduleField roomNumberField = newschedule.Definition.AddField(ScheduleFieldType.Instance, roomNumberParam.Id);
                    ScheduleField roomNameField = newschedule.Definition.AddField(ScheduleFieldType.Instance, roomNameParam.Id);
                    ScheduleField roomDeptField = newschedule.Definition.AddField(ScheduleFieldType.Instance, roomDeptParam.Id);
                    ScheduleField roomCommField = newschedule.Definition.AddField(ScheduleFieldType.Instance, roomCommParam.Id);
                    ScheduleField roomAreaField = newschedule.Definition.AddField(ScheduleFieldType.Instance, roomAreaParam.Id);
                    ScheduleField roomLevelField = newschedule.Definition.AddField(ScheduleFieldType.Instance, roomLevelParam.Id);

                    //set displays
                    roomLevelField.IsHidden = true;
                    roomLevelField.DisplayType = ScheduleFieldDisplayType.Totals;
                    roomAreaField.DisplayType = ScheduleFieldDisplayType.Totals;

                    // 05.set totals
                    newschedule.Definition.IsItemized = true;
                    newschedule.Definition.ShowGrandTotal = true;
                    newschedule.Definition.ShowGrandTotalCount = true;

                    //grouping
                    ScheduleSortGroupField sortbyRoomName = new ScheduleSortGroupField(roomNameField.FieldId);

                }

                

                t.Commit();

            }

                return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "MOD0201_Schedules";
            string buttonTitle = "Create Schedules";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.icons8_year_view_pulsar_color_3296,
                Properties.Resources.icons8_year_view_pulsar_color_1696,
                "This is a tooltip for Schedules");

            return myButtonData1.Data;
        }
    }
}
