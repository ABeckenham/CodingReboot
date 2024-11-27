#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using Forms = System.Windows.Forms;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class RFamilies_Exportplus : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here


            ///UI to select the folder path for the export
            //prompt user to select file 
            string folderpath = "";
            using(Forms.FolderBrowserDialog selectFile = new Forms.FolderBrowserDialog())
            {
                selectFile.Description = "Select the folder location";
                if (selectFile.ShowDialog() == Forms.DialogResult.OK) 
                {
                    folderpath =selectFile.SelectedPath;
                }
                else
                {
                    TaskDialog.Show("error", "Please select a folder");
                    return Result.Failed;
                }
            }

            //create progress bar


            //create folder in path
            //retrieve file name
            string fileName = doc.Title;
            string savepath = Path.Combine(folderpath, fileName);

            //variables for later
            int totalFamilies = 0;
            int totalinPlace = 0;
            int totalNests = 0;
                      


            //error handling for in-place families
            //collect familyName
            Dictionary<string, List<FamilySymbol>> FamilyDictionary = new Dictionary<string, List<FamilySymbol>>();

            //create dictionary for families against category name
            //Collect all loadable families
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            var familySymbols = collector.OfClass(typeof(FamilySymbol));

            foreach (FamilySymbol familySymbol in familySymbols)
            {
                if (familySymbol.Family.IsInPlace)
                {
                    continue;
                }
                else
                {
                    totalinPlace++;
                }

                string categoryName = familySymbol.Category?.Name;
                if (categoryName != null)
                {
                    //if the family category is not in the dictionary, add new key/list
                    if (!FamilyDictionary.ContainsKey(categoryName))
                    {
                        FamilyDictionary[categoryName] = new List<FamilySymbol>();                                     
                        //create folder for category
                        string categorypath = Path.Combine (folderpath, categoryName);
                    }

                    //if it does exist add the symbol to the current key list
                    FamilyDictionary[categoryName].Add(familySymbol);
                    totalFamilies++;

                    //search for nested families
                    //to collect nested families you collect the nestedFamilyIds
                    var nestedFamilyIDs = familySymbol.Family.GetFamilySymbolIds();
                    foreach(ElementId nestId in nestedFamilyIDs)
                    {
                        FamilySymbol nestedSymbol = doc.GetElement(nestId) as FamilySymbol;
                        FamilyDictionary[categoryName].Add(nestedSymbol);
                        //create folder for nested Families
                        string categorypath = Path.Combine(folderpath, categoryName,familySymbol.ToString());
                    }

                }

                




                //save family
                //create folder for each family with nests in the category folder

                //task dialogue - number of families, number of families with nests








                return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "RFamilies_Exportplus";
            string buttonTitle = "Family Export Plus";

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
