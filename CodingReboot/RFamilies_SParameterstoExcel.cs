#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Forms = System.Windows.Forms;
#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class RFamilies_SParameterstoExcel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get Revit application and document
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            string folderpath = "";

            //using a form, get folder path, with error handling
            using (Forms.FolderBrowserDialog selectFile = new Forms.FolderBrowserDialog())
            {
                selectFile.Description = "Select folder containing Revit family files";
                if (selectFile.ShowDialog() == Forms.DialogResult.OK)
                {
                    folderpath = selectFile.SelectedPath;
                }
                else
                {
                    TaskDialog.Show("Error", "Please select a folder");
                    return Result.Failed;
                }
            }

            try
            {
                string savepath = Path.Combine(folderpath, "FamilyParameters_Complete.xlsx");
                string[] filedata = Directory.GetFiles(folderpath, "*.rfa");
                int totalFiles = filedata.Length;

                if (totalFiles == 0)
                {
                    TaskDialog.Show("Error", "No .rfa files found in selected folder.");
                    return Result.Failed;
                }

                using (var progressBar = new ProgressBar())
                {
                    progressBar.Show();
                    int currentFile = 0;
                    int processedFiles = 0;
                    int skippedFiles = 0;

                    using (SpreadsheetDocument document = SpreadsheetDocument.Create(savepath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = document.AddWorkbookPart();
                        workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new SheetData());

                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                        Sheet sheet = new Sheet()
                        {
                            Id = workbookPart.GetIdOfPart(worksheetPart),
                            SheetId = 1,
                            Name = "FamilyParameters"
                        };
                        sheets.AppendChild(sheet);

                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                        // Create header row
                        Row headerRow = new Row();
                        headerRow.AppendChild(CreateCell("Family Name"));
                        headerRow.AppendChild(CreateCell("Parameter Name"));
                        headerRow.AppendChild(CreateCell("Storage Type"));
                        headerRow.AppendChild(CreateCell("Parameter Group"));
                        headerRow.AppendChild(CreateCell("Is Shared"));
                        headerRow.AppendChild(CreateCell("GUID"));
                        headerRow.AppendChild(CreateCell("Is Instance"));
                        headerRow.AppendChild(CreateCell("Formula"));
                        sheetData.AppendChild(headerRow);

                        foreach (string file in filedata)
                        {
                            currentFile++;
                            Forms.Application.DoEvents();

                            progressBar.UpdateProgress($"Processing family {currentFile} of {totalFiles}", (currentFile * 100) / totalFiles);

                            if (Path.GetExtension(file).Equals(".rfa", StringComparison.OrdinalIgnoreCase))
                            {
                                Document familyDoc = null;
                                try
                                {
                                    familyDoc = uiapp.Application.OpenDocumentFile(file);
                                    string filename = Path.GetFileNameWithoutExtension(file);

                                    if (familyDoc.IsFamilyDocument)
                                    {
                                        FamilyManager famMan = familyDoc.FamilyManager;

                                        foreach (FamilyParameter param in famMan.Parameters)
                                        {
                                            Row dataRow = new Row();

                                            // Add family name and parameter name
                                            dataRow.AppendChild(CreateCell(filename));
                                            dataRow.AppendChild(CreateCell(param.Definition.Name));

                                            // Add storage type safely
                                            string storageType = "Unknown";
                                            try
                                            {
                                                storageType = param.StorageType.ToString();
                                            }
                                            catch { }
                                            dataRow.AppendChild(CreateCell(storageType));

                                            // Add parameter group safely
                                            string paramGroup = "Other";
                                            try
                                            {
                                                paramGroup = param.Definition.ParameterGroup.ToString();
                                            }
                                            catch { }
                                            dataRow.AppendChild(CreateCell(paramGroup));

                                            // Add is shared and GUID
                                            dataRow.AppendChild(CreateCell(param.IsShared.ToString()));
                                            dataRow.AppendChild(CreateCell(param.IsShared ? param.GUID.ToString() : "N/A"));

                                            // Add instance/type info
                                            dataRow.AppendChild(CreateCell(param.IsInstance.ToString()));

                                            // Add formula if any
                                            string value = "None";
                                            try
                                            {
                                                if (famMan.CurrentType != null)
                                                {
                                                    var paramValue = famMan.CurrentType.AsValueString(param);
                                                    value = paramValue ?? "None";
                                                }
                                            }
                                            catch
                                            {
                                                value = "None";
                                            }
                                            dataRow.AppendChild(CreateCell(value));

                                            sheetData.AppendChild(dataRow);
                                        }
                                        processedFiles++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    skippedFiles++;
                                    continue;
                                }
                                finally
                                {
                                    if (familyDoc != null)
                                    {
                                        familyDoc.Close(false);
                                        Forms.Application.DoEvents();
                                    }
                                }
                            }
                        }

                        workbookPart.Workbook.Save();
                    }

                    TaskDialog summaryDialog = new TaskDialog("Operation Complete");
                    summaryDialog.MainInstruction = "Family parameter export complete";
                    summaryDialog.MainContent = $"Successfully processed {processedFiles} families\n" +
                                              $"Skipped {skippedFiles} families\n" +
                                              $"Excel file saved to:\n{savepath}";
                    summaryDialog.Show();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to create Excel file.\nError: {ex.Message}");
                return Result.Failed;
            }
        }

        private Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(text ?? "N/A");
            return cell;
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "Family Parameter Export";
            string buttonTitle = "Export All Parameters";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "Export all parameters from families");

            return myButtonData1.Data;
        }
    }

    public class ProgressBar : Forms.Form
    {
        private Forms.ProgressBar progressBar;
        private Forms.Label label;
        private Forms.Timer messageTimer;

        public ProgressBar()
        {
            this.Width = 400;
            this.Height = 100;
            this.FormBorderStyle = Forms.FormBorderStyle.FixedDialog;
            this.ControlBox = false;
            this.StartPosition = Forms.FormStartPosition.CenterScreen;
            this.Text = "Processing Families...";

            messageTimer = new Forms.Timer();
            messageTimer.Interval = 100;
            messageTimer.Tick += (s, e) => Forms.Application.DoEvents();
            messageTimer.Start();

            label = new Forms.Label();
            label.Width = 360;
            label.Location = new System.Drawing.Point(10, 10);

            progressBar = new Forms.ProgressBar();
            progressBar.Width = 360;
            progressBar.Location = new System.Drawing.Point(10, 30);
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;

            this.Controls.Add(label);
            this.Controls.Add(progressBar);
        }

        protected override void OnFormClosing(Forms.FormClosingEventArgs e)
        {
            messageTimer.Stop();
            messageTimer.Dispose();
            base.OnFormClosing(e);
        }

        public void UpdateProgress(string text, int percentage)
        {
            if (!this.IsDisposed && !this.Disposing)
            {
                if (this.InvokeRequired)
                {
                    try
                    {
                        this.Invoke(new Action(() =>
                        {
                            if (!this.IsDisposed && !this.Disposing)
                            {
                                label.Text = text;
                                progressBar.Value = percentage;
                                this.Update();
                            }
                        }));
                    }
                    catch (ObjectDisposedException) { }
                }
                else
                {
                    label.Text = text;
                    progressBar.Value = percentage;
                    this.Update();
                }
            }
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "RFamilies_SParameterstoExcel";
            string buttonTitle = "Dim Detective - rooms";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.icons8_measure_pulsar_Blue_3296,
                Properties.Resources.icons8_measure_pulsar_Blue_1696,
                "Dimension all rooms");

            return myButtonData1.Data;
        }
    }
}
