#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Media3D;

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD0203_DimD2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here
            //1. collect all the rooms in the project
            FilteredElementCollector roomcollector = Utils.GetAllRooms(doc);
                      

            double counter = 0;

            foreach (Room rm in roomcollector)
            {
                //3. set options and get room boundaries
                SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
                options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
                List<BoundarySegment> seglist = rm.GetBoundarySegments(options).First().ToList();

                ReferenceArray referenceArray = new ReferenceArray();
                List<XYZ> pointlist = new List<XYZ>();
                List<XYZ> pointlisth = new List<XYZ>();

                //2. loop through and get room boundaries as segments
                foreach (BoundarySegment curSeg in seglist)
                {
                    
                    Curve boundCurve = curSeg.GetCurve();
                    XYZ midpoint = boundCurve.Evaluate(0.5, true);

                    if (Utils.IsLineVertical(boundCurve))
                    {
                        //3. get the boundary walls
                        Element curWall = doc.GetElement(curSeg.ElementId);

                        referenceArray.Append(new Reference(curWall));
                        pointlist.Add(midpoint);
                    }
                }
                XYZ point1 = pointlist.First();
                XYZ point2 = pointlist.Last();

                using (Transaction t = new Transaction(doc))
                {

                    t.Start("Create dimensions in Rooms");
                                      
                   
                    //4. dimension the horizontal and vertical walls of each room.
                    Line dimline = Line.CreateBound(point1, point2);
                    Dimension newdim = doc.Create.NewDimension(doc.ActiveView, dimline, referenceArray);
                    
                    counter++;
                    t.Commit();
                }
            }
                       
        
            TaskDialog.Show("Counter", counter.ToString());



            // make sure not to overlap the dimension 
            //5. add a counter variable to keep track of the number of dimensions created


            return Result.Succeeded;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "MOD0203_DimD2";
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
