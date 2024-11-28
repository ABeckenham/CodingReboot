#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Reflection;
using Forms = System.Windows.Forms;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class RTemplate_ExViewFamilyTypes : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here
            //This tool will be used to quickly export out ViewFamilyTypes, which are the viewtype.types:(01)GAPlans in each ViewType:elevation,plan,section

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewFamilyType));

            //collct current file name 
            string filename = System.IO.Path.GetFileNameWithoutExtension(doc.PathName);
            string filepath = filename + "_ViewFamilyTypes.xlsx";
            ///UI to select the folder path for the export
            //prompt user to select file 
            string folderpath = "";

            using (Forms.FolderBrowserDialog selectFile = new Forms.FolderBrowserDialog())
            {
                selectFile.Description = "Select the folder location for your excel file";
                if (selectFile.ShowDialog() == Forms.DialogResult.OK)
                {
                    folderpath = selectFile.SelectedPath;
                }
                else
                {
                    TaskDialog.Show("err6 7 hbgor", "Please select a folder");
                    return Result.Failed;
                }
            }

            string savepath = System.IO.Path.Combine(folderpath, filepath);

            try
            {
                //create excel file 
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(savepath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookpart = document.AddWorkbookPart();
                    workbookpart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                    WorksheetPart worksheetpart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetpart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new SheetData());

                    Sheets sheets = workbookpart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet()
                    {
                        Id = workbookpart.GetIdOfPart(worksheetpart),
                        SheetId = 1,
                        Name = "FamilyParameters"
                    };

                    sheets.AppendChild(sheet);
                    SheetData sheetData = worksheetpart.Worksheet.GetFirstChild<SheetData>();

                    // Create header row
                    Row headerRow = new Row();
                    headerRow.AppendChild(CreateCell("ViewType"));
                    headerRow.AppendChild(CreateCell("ViewFamilyType"));
                    sheetData.AppendChild(headerRow);


                    //add records 
                    foreach (ViewFamilyType viewType in collector)
                    {
                        string name = viewType.Name;
                        ViewFamily vType = viewType.ViewFamily;

                        Row datarow = new Row();

                        //add datat to row

                        datarow.AppendChild(CreateCell(name));
                        datarow.AppendChild(CreateCell(vType.ToString()));
                        sheetData.AppendChild(datarow);

                    }
                    workbookpart.Workbook.Save();

                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to create Excel file.\nError: {ex.Message}");

                return Result.Failed;
            }


            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "RTemplate_ViewFamilyTypes";
            string buttonTitle = "RTemplate_ViewFamilyTypes";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Red_32,
                Properties.Resources.Red_16,
                "This is a tooltip for Button 2");

            return myButtonData1.Data;
        }
        private Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(text ?? "N/A");
            return cell;
        }
    }
}
