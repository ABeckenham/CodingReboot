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

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class WPF01 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            //Step 01: put any code needed for the form here

            //Step 02: Open the form
            MyForms.MyForm currentForm = new MyForms.MyForm()
            {
                Width = 500,
                Height = 450,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost=true,
            };

            currentForm.ShowDialog();

            //Step 04: get form data and do something
            if(currentForm.DialogResult == false)
            {
                return Result.Cancelled;
            }

            //do something
            string textboxresult = currentForm.tbxFile.Text;

            bool checkbox1Value = currentForm.getCheckbox1();

            string radioButtonValue = currentForm.GetGroup1();

            TaskDialog.Show("Test", "textbox result is " + textboxresult);
            if( checkbox1Value == true) 
            {
                TaskDialog.Show("test", "Check box 1 was selected");
            }

            TaskDialog.Show("test", radioButtonValue);

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Command 1";

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
