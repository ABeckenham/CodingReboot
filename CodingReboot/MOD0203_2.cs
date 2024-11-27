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

#endregion

namespace CodingReboot
{
    [Transaction(TransactionMode.Manual)]
    public class MOD0203_2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            //1. Select Room
            Reference curRef = uiapp.ActiveUIDocument.Selection.PickObject(ObjectType.Element, "Select a room");
            Room curRoom = doc.GetElement(curRef) as Room; 

            //2. create reference array and point list
            ReferenceArray referenceArray = new ReferenceArray();
            List<XYZ> pointlist = new List<XYZ>();  

            //3. set options and ge troom boundaries
            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
            options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;

            List<BoundarySegment> boundarySegList = curRoom.GetBoundarySegments(options).First().ToList();
            //return ilist by default so this TOLIST converts it to a regular list

            //4. Loop through rom boundaries
            foreach(BoundarySegment curSeg in boundarySegList)
            {
                //4a.get boundary geometry
                Curve boundCurve = curSeg.GetCurve();
                XYZ midpoint = boundCurve.Evaluate(0.25, true);

                //4b. check if line is vertical
                if (Utils.IsLineVertical(boundCurve)==false)
                {
                    //5. get boundary wall
                    //each boundary segment knows what creates it 
                    //we have to dimension the wall, we cannot dimension to the boundary seg
                    Element curWall = doc.GetElement(curSeg.ElementId);

                    //6. add to ref and point array
                    referenceArray.Append(new Reference(curWall));
                    pointlist.Add(midpoint);
                }

                //7. create line for dimension
                XYZ point1 = pointlist.First();
                XYZ point2 = pointlist.Last();
                //Line dimLine = Line.CreateBound(point1,new XYZ(point2.X, point1.Y,0));
                Line dimLine = Line.CreateBound(point1, point2);


                //8. create dimension 
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create dimension");
                    Dimension newDim = doc.Create.NewDimension(doc.ActiveView, dimLine, referenceArray);
                    t.Commit();
                }
                
            }

            //this method will still dim to the wall centreline eventhough we selected finish. 
            //we will cover this in another class

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Command 2";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Red_32,
                Properties.Resources.Red_16,
                "This is a tooltip for Button 2");

            return myButtonData1.Data;
        }
    }
}
