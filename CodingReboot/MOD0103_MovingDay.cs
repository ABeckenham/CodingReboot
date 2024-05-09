#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD0103_MovingDay : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            //prompt user to select file 
            ////using Windows.Forms
            //Forms.OpenFileDialog selectFile = new Forms.OpenFileDialog();
            //selectFile.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            //selectFile.InitialDirectory = "C:\\";
            //selectFile.Multiselect = false;

            string excelFile = "C:\\AMY\\Coding\\CodingReboot\\RAB_Module_03_Challenge_Files\\RAB_Module 03_Furniture.xlsx";


            //open excel file 
            ///////////////////////////////////
            //creating an instance of excel in the code
                      
            List<List<string>> typesExceldata = GetExceldata(excelFile, "Furniture types");

            //create the furnitureTypes             
            List<FurnitureTypes> FurnitureTypeslist = new List<FurnitureTypes>();

            foreach (List<string> curlist in typesExceldata)
            {
                if (typesExceldata.IndexOf(curlist) ==0)
                {
                    continue;
                }
                else
                {
                    FurnitureTypes curFurnType = new FurnitureTypes(curlist[0], curlist[1], curlist[2]);
                    FurnitureTypeslist.Add(curFurnType);
                }
            }

            //create the furniture sets
            
            List<List<string>> furnExceldata = GetExceldata(excelFile, "Furniture sets");
            List<FurnitureSet> FurnitureSetsList = new List<FurnitureSet>();
            
            foreach (List<string> curlist in furnExceldata)
            {
                if (furnExceldata.IndexOf(curlist) == 0)
                {
                    continue;
                }
                else
                {
                    string curSet = curlist[0];
                    string curRoom = curlist[1];
                    string curInclFurniture = curlist[2];
                    List<string> splitlist = curInclFurniture.Split(',').Select(s => s.Trim()).ToList();
                    List<FurnitureTypes> stringIncludedFurniture = new List<FurnitureTypes>();

                    foreach (string curstring in splitlist)                         
                    {
                        foreach (FurnitureTypes furntype in FurnitureTypeslist)
                        {
                            if (curstring == furntype.FurnitureName)
                            {
                                stringIncludedFurniture.Add(furntype);
                            }
                            else continue;
                        }                        
                    }
                    FurnitureSet curFurnSet = new FurnitureSet(curSet, curRoom, stringIncludedFurniture);
                    FurnitureSetsList.Add(curFurnSet);                                        
                }
            }


            using (Transaction t = new Transaction(doc))
            {
                t.Start("Insert family into room");
                //collect all the rooms, location point and furniture set parameter
                //rooms are spatial element class
                FilteredElementCollector RoomCollector = new FilteredElementCollector(doc);
                RoomCollector.OfCategory(BuiltInCategory.OST_Rooms);

                
                foreach (SpatialElement curRoom in RoomCollector)
                {
                    LocationPoint curPoint = curRoom.Location as LocationPoint;
                    XYZ roomPoint = curPoint.Point as XYZ;

                    string curRoomParamFurnSet = Utils.GetParameterValueAsString(curRoom, "Furniture Set");

                    //Place furniture by furniture set, into each room
                    //Update the room furniture count using method from furniture set

                    List<FurnitureTypes> curRoomFurnTypes = new List<FurnitureTypes>();

                    foreach (FurnitureSet FS in FurnitureSetsList)
                    {
                        if (curRoomParamFurnSet == FS.SetName)
                        {
                            List<FurnitureTypes> RoomFurnitureTypes = FS.IncludedFurniture;
                            foreach (FurnitureTypes curfurntype in RoomFurnitureTypes)
                            {
                                FamilySymbol curFamSymbol = Utils.GetFamilySymbolByName(doc, curfurntype.revitFamilyName, curfurntype.revitFamilyType);
                                curFamSymbol.Activate();
                                FamilyInstance curFamInstance = doc.Create.NewFamilyInstance(roomPoint, curFamSymbol, StructuralType.NonStructural);
                            }
                            Utils.SetParameterValue(curRoom, "Furniture Count", FS.GetFurnitureCount());
                        }
                        else continue;
                    }
                }
                t.Commit();
            }        
            return Result.Succeeded;
        }

        private static List<List<string>> GetExceldata(string excelFile, string worksheetname)
        {
            Excel.Application excel = new Excel.Application();
            //creating an instance of excel workbook
            Excel.Workbook workbook = excel.Workbooks.Open(excelFile);
            // in excel the index starts at 1 ... wierd.
            Excel.Worksheet worksheet = workbook.Worksheets[worksheetname];
            //create range of the data in the worksheet - where data exist
            //Excel.Range range = worksheet.UsedRange;
            //if an error happens cast it to another type 
            Excel.Range setsRange = (Excel.Range)worksheet.UsedRange;
            // get rows and columns count
            int rows = setsRange.Rows.Count;
            int cols = setsRange.Columns.Count;
             // read excel data into a list            
            List<List<string>> excelData = new List<List<string>>();

            //forloop to go through, iterate through everything
            for (int i = 1; i <= rows; i++)
            {
                List<string> rowData = new List<string>();
                for (int j = 1; j <= cols; j++)
                {
                    string ColData = Convert.ToString(worksheet.Cells[i, j].Value);
                    rowData.Add(ColData);                    
                }
                excelData.Add(rowData);                             

            }
            workbook.Save();
            excel.Quit();
            return excelData;
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
}

public class FurnitureSet
{
    public string SetName { get; set; }
    public string RoomType { get; set; }
    public List<FurnitureTypes> IncludedFurniture { get; set; }

    //CONSTRUCTOR
    public FurnitureSet(string _SetName, string _RoomType, List<FurnitureTypes> _IncludedFurniture)
    {
        SetName = _SetName;
        RoomType = _RoomType;
        IncludedFurniture = _IncludedFurniture;
    }


    //methods

    public int GetFurnitureCount()
    {
        return IncludedFurniture.Count();
    }
}

public class FurnitureTypes
{
    public string FurnitureName { get; set; }
    public string revitFamilyName { get; set; }
    public string revitFamilyType { get; set; }

    //CONSTRUCTOR
    public FurnitureTypes(string _FurnitureName, string _revitFamilyName, string _revitFamilyType )
    {
        FurnitureName = _FurnitureName;
        revitFamilyName = _revitFamilyName;
        revitFamilyType = _revitFamilyType;

    }        

}
