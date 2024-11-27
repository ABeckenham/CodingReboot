#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Windows.Media.Media3D;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD0203_DimD1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            //1. get all grids in view           
            FilteredElementCollector Collector = new FilteredElementCollector(doc, doc.ActiveView.Id)
            .OfClass(typeof(Grid)).WhereElementIsNotElementType();


            //2. reference array and point list            
            ReferenceArray arrayList = new ReferenceArray();
            ReferenceArray arrayListv = new ReferenceArray();
            List<XYZ> pointList = new List<XYZ>();
            List<XYZ> pointListv = new List<XYZ>();

            foreach (Grid curGrid in Collector)
            {
                Curve curCurve = curGrid.Curve;
                double length = curCurve.Length;
                XYZ locPoint = curCurve.Evaluate((length - 3), false);                
                //XYZ locPoint = curCurve.GetEndPoint(0);                                               

                if (Utils.IsLineVertical(curCurve))
                {
                    arrayListv.Append(new Reference(curGrid));
                    pointListv.Add(locPoint);
                }
                else
                {
                    arrayList.Append(new Reference(curGrid));
                    pointList.Add(locPoint);
                }
                
            }

            //pointlist == horizontals 

            List<XYZ> sortedList = pointList.OrderBy(p => p.X).ThenBy(p => p.Y).ToList();
            XYZ point1 = sortedList.First();
            XYZ point2 = sortedList.Last();
           
            Autodesk.Revit.DB.Line dimline = Autodesk.Revit.DB.Line.CreateBound(point1, new XYZ(point1.X, point2.Y, 0));

            //pointlistv == verticals 

            List<XYZ> sortedListh = pointListv.OrderBy(p => p.X).ThenBy(p => p.Y).ToList();
            XYZ point1v = sortedListh.First();
            XYZ point2v = sortedListh.Last();

            Autodesk.Revit.DB.Line dimlinev = Autodesk.Revit.DB.Line.CreateBound(new XYZ(point1v.X,point2v.Y,0), point2v);

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Dimension Grids");
                Dimension newdim = doc.Create.NewDimension(doc.ActiveView, dimline, arrayList);
                Dimension newdimv = doc.Create.NewDimension(doc.ActiveView, dimlinev, arrayListv);
                t.Commit();
            }

            return Result.Succeeded;
        }



        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "MOD0203_DimD1";
            string buttonTitle = "Dim Detective - grids";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.icons8_measure_pulsar_color_3296,
                Properties.Resources.icons8_measure_pulsar_color_1696,
                "Dimension all grids");

            return myButtonData1.Data;
        }
    }
}
