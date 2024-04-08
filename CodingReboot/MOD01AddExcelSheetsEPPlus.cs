#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using Excel = OfficeOpenXml;
using Forms = System.Windows.Forms;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD01AddExcelSheetsEPPlus : IExternalCommand
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
                TaskDialog.Show("Error", "Please select an Excel File, make sure that it is closed");
                return Result.Failed;
            }

            //set epplus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //open excel file
            ExcelPackage excel = new ExcelPackage(excelFile);
            ExcelWorkbook workbook = excel.Workbook;
            ExcelWorksheet worksheet = workbook.Worksheets[2];

            // get rows and columns count
            int rows = worksheet.Dimension.Rows;
            int cols = worksheet.Dimension.Columns;

            // read excel data into a list
            //nested list, could also use an array
            List<List<string>> excelData = new List<List<string>>();

            //forloop to go through, iterate through everything
            for (int i = 1; i <= rows; i++)
            {
                List<string> rowData = new List<string>();
                for (int j = 1; j <= cols; j++)
                {
                    //string cellContent = worksheet.Cells[i, j].Value.ToString();
                    //avoid null values
                    string cellContent = Convert.ToString(worksheet.Cells[i, j].Value);
                    //here you can specify column or row hardcoded to get specific lists of things
                    rowData.Add(cellContent);
                }
                excelData.Add(rowData);
            }

            


            //write to excel - write a list of the drawings that have been made in Revit, and the file name 
            //Create new worksheet
            ////////////////////////////////////////
            ExcelWorksheet newWorkSheet = workbook.Worksheets.Add("Test EEPlus");
            

            //write data to excel
            for (int k = 1; k <= 10; k++) //row loop
            {
                for (int j = 1; j <= 10; j++) //column loop
                {
                    newWorkSheet.Cells[k,j].Value = "Row " + k.ToString() + ": Column " + j.ToString();
                }
            }

            //save and close excel file
            excel.Save();
            excel.Dispose();


            return Result.Succeeded;
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
