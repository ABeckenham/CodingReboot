#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD01AddExcelSheets : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            //prompt user to select file 
            //using Windows.Forms
            Forms.OpenFileDialog selectFile = new Forms.OpenFileDialog();
            selectFile.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            selectFile.InitialDirectory = "C:\\";
            selectFile.Multiselect = false;

            string excelFile = "";

            if (selectFile.ShowDialog() == Forms.DialogResult.OK)
                excelFile = selectFile.FileName;

            if (excelFile == "")
            {
                TaskDialog.Show("Error", "Please select an Excel File");
                return Result.Failed;
            }

            //open excel file 
            ///////////////////////////////////
            //creating an instance of excel in the code
            Excel.Application excel = new Excel.Application();
            //creating an instance of excel workbook
            Excel.Workbook workbook = excel.Workbooks.Open(excelFile);
            // in excel the index starts at 1 ... wierd.
            Excel.Worksheet worksheet = workbook.Worksheets["TIDP"];
            //create range of the data in the worksheet - where data exist
            //Excel.Range range = worksheet.UsedRange;
            //if an error happens cast it to another type 
            Excel.Range range = (Excel.Range)worksheet.UsedRange;

            // get rows and columns count
            int rows = range.Rows.Count;
            int cols = range.Columns.Count;

            // read excel data into a list
            /////////////////////////////////////
            //nested list, could also use an array
            List<List<string>> excelData = new List<List<string>>();

            //forloop to go through, iterate through everything
            for(int i = 1; i <= rows; i++)
            {
                List<string> rowData = new List<string>();
                for (int j = 1; j <= cols; j++)
                {
                    string ColData = Convert.ToString(worksheet.Cells[i, j].Value);
                    //here you can specify column or row hardcoded to get specific lists of things [row, column]
                    //functional breakdown columns 6
                    //spatial column 7
                    //form column 8
                    //Number column 9
                    //sheet title column 10
                    //string funcData = Convert.ToString(worksheet.Cells[i, 6].Value);
                    //string SpatData = Convert.ToString(worksheet.Cells[i, 7].Value);
                    //string formData = Convert.ToString(worksheet.Cells[i, 8].Value);
                    //string NumberData = Convert.ToString(worksheet.Cells[i, 9].Value);
                    //string nameData = Convert.ToString(worksheet.Cells[i, 10].Value);
                    rowData.Add(ColData);
                }
                excelData.Add(rowData);
            }

            //List<string> funcData = excelData[*:6];
            //List<string> SpatData = excelData[*:7];
            //List<string> formData = excelData[*:8];
            //List<string> NumberData = excelData[*:9];
            //List<string> nameData = excelData[*:10];


            //create sheets, inside a look for each element in excel
            // ooooo create each sheet from the column data!!! [i,6], [i,7] etc and set the name to that.
            //collect sheet type, just get first for now 
            FilteredElementCollector sheetType = new FilteredElementCollector(doc);                
            sheetType.OfCategory(BuiltInCategory.OST_TitleBlocks);

            ViewSheet newSheet = ViewSheet.Create(doc, sheetType.FirstElementId());

            newSheet.Name = "my new sheet";
            newSheet.SheetNumber = "101";



            ////write to excel - write a list of the drawings that have been made in Revit, and the file name 
            ////Create new worksheet
            //////////////////////////////////////////

            //Excel.Worksheet newWorksheet = workbook.Worksheets.Add();
            //newWorksheet.Name = "Test Interop.Excel";

            ////write data to excel
            //for (int k = 1; k <= 10; k++) //row loop
            //{
            //    for (int j = 1; j <= 10; j++) //column loop
            //    {
            //        newWorksheet.Cells[k,j].Value = "Row " + k.ToString() + ": Column " + j.ToString();
            //    }
            //}


            //get project info to check the roles and project code matches in the excel



            //save and close excel file
            workbook.Save();
            excel.Quit();

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "addSheetExcel";
            string buttonTitle = "Add Sheets from TIDP";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.icons8_insert_page_pulsar_color_3296,
                Properties.Resources.icons8_insert_page_pulsar_color_1696,
                "This is a tooltip for Add Sheets from TIDP");

            return myButtonData1.Data;
        }
    }
}
