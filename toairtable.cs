#region Namespaces
using AirtableApiClient;
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
using System.Threading.Tasks;

#endregion

namespace RAA_ExportToAirtable
{
	[Transaction(TransactionMode.Manual)]
	public class Command1 : IExternalCommand
	{
		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			// this is a variable for the Revit application
			UIApplication uiapp = commandData.Application;

			// this is a variable for the current Revit model
			Document doc = uiapp.ActiveUIDocument.Document;

			// Your code goes here
			List<SpatialElement> elemList = new FilteredElementCollector(doc)
				.WhereElementIsNotElementType()
				.OfCategory(BuiltInCategory.OST_Rooms)
				.Cast<SpatialElement>()
				.ToList();

			List<List<string>> roomData = new List<List<string>>();
			foreach(SpatialElement curRoom in elemList)
			{
				List<string> room = new List<string>();
				room.Add(curRoom.Id.ToString());
				room.Add(curRoom.Number);
				room.Add(curRoom.Name);
				room.Add(curRoom.Level.Name);
				room.Add(curRoom.Area.ToString());
				roomData.Add(room);
			}

			WriteToAirtable(roomData);
			
			TaskDialog.Show("Export to Airtable", "Exporting room data to Airtable.");

			return Result.Succeeded;
		}
		internal static async void WriteToAirtable(List<List<string>> roomData)
		{
			// use Nuget to add the following packages:
			// Airtable (https://github.com/ngocnicholas/airtable.net)
			// System.Text.Json

			string baseId = "ENTER YOUR BASE ID HERE";
			string accessToken = "ENTER YOUR PERSONAL ACCESS TOKEN";
			string tableName = "Rooms";

			using (AirtableBase curBase = new AirtableBase(accessToken, baseId))
			{
				if (curBase == null) return;

				foreach (List<string> room in roomData)
				{
					var fields = new Fields();
					fields.AddField("Room ID", room[0]);
					fields.AddField("Room Number", room[1]);
					fields.AddField("Room Name", room[2]);
					fields.AddField("Room Area", room[3]);

					Task<AirtableCreateUpdateReplaceRecordResponse> task = curBase.CreateRecord(tableName, fields, true);
					var response = await task;

					if (!response.Success)
					{
						string errorMessage = null;
						if (response.AirtableApiError is AirtableApiException)
						{
							errorMessage = response.AirtableApiError.ErrorMessage;
						}
						else
						{
							errorMessage = "Unknown error";
						}
						// Report errorMessage
					}
					else
					{
						var record = response.Record;
					}
				}
			}
		}
		internal static PushButtonData GetButtonData()
		{
			// use this method to define the properties for this command in the Revit ribbon
			string buttonInternalName = "btnCommand1";
			string buttonTitle = "Button 1";

			ButtonDataClass myButtonData1 = new ButtonDataClass(
				buttonInternalName,
				buttonTitle,
				MethodBase.GetCurrentMethod().DeclaringType?.FullName,
				Properties.Resources.Blue_32,
				Properties.Resources.Blue_16,
				"This is a tooltip for Button 1");

			return myButtonData1.Data;
		}

	}
}
