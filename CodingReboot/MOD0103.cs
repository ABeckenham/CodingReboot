#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Reflection;
using Autodesk.Revit.DB.Structure;
using System.Linq;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD0103 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here
            Building theature = new Building("grand Opera House", "5 main street", 4, 35000);
            Building Office = new Building("BigBusiness", "6 main street", 4, 58000);
            Building hotel = new Building("hotellia", "7 main street", 4, 45000);

            List<Building> buildings = new List<Building>();
            buildings.Add(theature);
            buildings.Add(Office);
            buildings.Add(hotel);

            Neighbourhood downtown = new Neighbourhood("Downtown", "Middletown", "CT", buildings);

            TaskDialog.Show("Test", $"There are {downtown.GetBuildingCount()}" + $" buildings in the {downtown.Name} neighbourhood");

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);

            FamilySymbol curFS = Utils.GetFamilySymbolByName(doc, "Desk", "60\" x 30\"");

            using ( Transaction t = new Transaction(doc))
            {
                foreach (SpatialElement room in collector)
                {
                    t.Start("Insert family into room");
                    curFS.Activate();

                    LocationPoint loc = room.Location as LocationPoint;
                    XYZ roompoint = loc.Point as XYZ;

                    FamilyInstance fi = doc.Create.NewFamilyInstance(roompoint, curFS, StructuralType.NonStructural);

                    string name = Utils.GetParameterValueAsString(room, "Department");
                    Utils.SetParameterValue(room, "Floor Finish", "CT");
                }
                t.Commit();

                string myline = "one, two, three, four, five";
                string[] splitline = myline.Split(',');
                TaskDialog.Show("test", splitline[0].Trim());
                TaskDialog.Show("test", splitline[3].Trim());
            }
                



            return Result.Succeeded;

        }

        

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

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

    public class Building
    {
        //properties
        public string Name { get; set; }
        public string Address { get; set; }
        public int NumFloors { get; set; }
        public double Area { get; set; }

        //constructor
        public Building(string _name, string _address, int _numFloors, double _area) 
        {
            Name = _name;
            Address = _address;
            NumFloors = _numFloors;
            Area = _area;
        }
    }

    public class Neighbourhood
    {
        public string Name { get; set; }
        public string City { get; set; }
        public string State { get; set; }

        public List<Building> BuildingList { get; set; }
        public Neighbourhood(string _name, string _State, string _City, List<Building> buildings)
        {
            Name = _name;
            City = _City;
            State = _State;
            BuildingList = buildings;
           
        }

        //add methods to class
        public int GetBuildingCount()
        {
            return BuildingList.Count;
        }


    }
}
