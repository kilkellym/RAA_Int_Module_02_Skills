#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace RAA_Int_Module_02_Skills
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // 1a. Filtered Element Collector by view
            View curView = doc.ActiveView;
            FilteredElementCollector collector = new FilteredElementCollector(doc, curView.Id);

            // 1b. ElementMultiCategoryFilter
            List<BuiltInCategory> catList = new List<BuiltInCategory>();
            catList.Add(BuiltInCategory.OST_Rooms);
            catList.Add(BuiltInCategory.OST_Doors);
            catList.Add(BuiltInCategory.OST_Walls);
            catList.Add(BuiltInCategory.OST_Areas);

            ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(catList);
            collector.WherePasses(catFilter).WhereElementIsNotElementType();

            // 1c. use LINQ to get family symbol by name
            FamilySymbol curDoorTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Door Tag"))
                .First();

            FamilySymbol curRoomTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Room Tag"))
                .First();

            FamilySymbol curWallTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Wall Tag"))
                .First();

            // 2. create dictionary for tags
            Dictionary<string, FamilySymbol> tags = new Dictionary<string, FamilySymbol>();
            tags.Add("Doors", curDoorTag);
            tags.Add("Rooms", curRoomTag);
            tags.Add("Walls", curWallTag);

            int counter = 0;
            using(Transaction t = new Transaction(doc))
            {
                t.Start("Insert tags");
                foreach (Element curElem in collector)
                {
                    // 3. get point from location
                    XYZ insPoint;
                    LocationPoint locPoint;
                    LocationCurve locCurve;

                    Location curLoc = curElem.Location;

                    if (curLoc == null)
                        continue;

                    locPoint = curLoc as LocationPoint;
                    if (locPoint != null)
                    {
                        // is a location point
                        insPoint = locPoint.Point;
                    }
                    else
                    {   // is a location curve
                        locCurve = curLoc as LocationCurve;
                        Curve curCurve = locCurve.Curve;
                        //insPoint = curCurve.GetEndPoint(1);
                        insPoint = GetMidpointBetweenTwoPoints(curCurve.GetEndPoint(0), curCurve.GetEndPoint(1));
                    }

                    // 6. check wall type
                    if (curElem.Category.Name == "Walls")
                    {
                        Wall curWall = curElem as Wall;
                        WallType curWallType = curWall.WallType;
                       
                        if (curWallType.Kind == WallKind.Curtain)
                        {
                            TaskDialog.Show("Test", "Found a curtain wall!");
                        }
                    }

                    ViewType curViewType = curView.ViewType;
                    
                    if(curViewType == ViewType.AreaPlan)
                    {
                        TaskDialog.Show("Test", "This is an area plan!");
                    }

                    FamilySymbol curTagType = tags[curElem.Category.Name];

                    // 4. create reference to element
                    Reference curRef = new Reference(curElem);

                    // 5a. place tag
                    if(IsElementTagged(curView, curElem) == false)
                    {
                        IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id,
                            curRef, false, TagOrientation.Horizontal, insPoint);
                        counter++;
                    }

                    // 5b. place area tag
                    //if (curElem.Category.Name == "Areas")
                    //{
                    //    ViewPlan curAreaPlan = curView as ViewPlan;
                    //    Area curArea = curElem as Area;

                    //    AreaTag curAreaTag = doc.Create.NewAreaTag(curAreaPlan, curArea, new UV(insPoint.X, insPoint.Y));
                    //    curAreaTag.TagHeadPosition = new XYZ(insPoint.X, insPoint.Y, 0);
                    //    curAreaTag.HasLeader = false;
                    //}

                }
                t.Commit();
            }

            TaskDialog.Show("Complete", $"Tagged {counter} elements.");

            return Result.Succeeded;
        }

        private bool IsElementTagged(View curView, Element curElem)
        {
            FilteredElementCollector collector = new FilteredElementCollector(curElem.Document, curView.Id);
            collector.OfClass(typeof(IndependentTag)).WhereElementIsNotElementType();

            foreach(IndependentTag curTag in collector)
            {
                List<ElementId> taggedIds = curTag.GetTaggedLocalElementIds().ToList();

                foreach(ElementId taggedId in taggedIds)
                {
                    if (taggedId.Equals(curElem.Id))
                        return true;
                }
            }
            return false;

        }

        private XYZ GetMidpointBetweenTwoPoints(XYZ point1, XYZ point2)
        {
            XYZ midPoint = new XYZ(
                (point1.X + point2.X) / 2,
                (point1.Y + point2.Y) / 2,
                (point1.Z + point2.Z) / 2);

            return midPoint;
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
