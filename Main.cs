using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Z8_CreateOpenings
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Получаем доступ к Revit, активному документу, базе данных документа

            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(d => d.Title.Contains("ОВ")).FirstOrDefault();
            if (ovDoc==null)
            {
                TaskDialog.Show("Error", "Не загружен файл ОВ");
                return Result.Cancelled;
            }
            // проверяем загружено ли семейство отверстия
            var openingFamily = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(f => f.Name.Equals("Z7_CreateOpenings_Opening"))
                .FirstOrDefault();
            if (openingFamily==null)
            {
                TaskDialog.Show("Error", "Не загружено семейство отверстия");
                return Result.Cancelled;
            }

            //находим все экземпляры воздуховодов
            List<Duct> allDucts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .OfType<Duct>()
                .ToList();

            //находим все экземпляры трубопроводов
            List<Pipe> allPipes = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Pipe))
                .WhereElementIsNotElementType()
                .OfType<Pipe>()
                .ToList();

            // получаем 3D вид
            var active3DView = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(v=>!v.IsTemplate)
                .FirstOrDefault();
            if (active3DView==null)
            {
                TaskDialog.Show("Error", "Не найден 3D вид");
                return Result.Cancelled;
            }


            Transaction t = new Transaction(arDoc);
            t.Start("Размещение отверстий");

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, active3DView);
            // выполняем проверку на пересечения стен с воздуховодами
            foreach (var duct in allDucts)
            {
               Line line = (duct.Location as LocationCurve).Curve as Line;
                XYZ point = line.GetEndPoint(0);
                XYZ direction = line.Direction;
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(i => i.Proximity <= line.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach (var intersect in intersections)
                {
                    double proximity = intersect.Proximity;
                    Reference reference = intersect.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ insertPoint = point + (direction * proximity);
                    // размещаем экземпляр семейства
                    FamilyInstance opening = arDoc.Create.NewFamilyInstance(insertPoint, openingFamily, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    Parameter width = opening.LookupParameter("ADSK_Отверстие_Ширина");
                    Parameter height = opening.LookupParameter("ADSK_Отверстие_Высота");
                    double offset = mmToFeet(150);
                    width.Set(duct.Diameter + offset) ;
                    height.Set(duct.Diameter + offset);
                }
            }
            // выполняем проверку на пересечения стен с трубопроводами
            foreach (var pipe in allPipes)
            {
                Line line = (pipe.Location as LocationCurve).Curve as Line;
                XYZ point = line.GetEndPoint(0);
                XYZ direction = line.Direction;
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(i => i.Proximity <= line.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach (var intersect in intersections)
                {
                    double proximity = intersect.Proximity;
                    Reference reference = intersect.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ insertPoint = point + (direction * proximity);
                    // размещаем экземпляр семейства
                    FamilyInstance opening = arDoc.Create.NewFamilyInstance(insertPoint, openingFamily, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    Parameter width = opening.LookupParameter("ADSK_Отверстие_Ширина");
                    Parameter height = opening.LookupParameter("ADSK_Отверстие_Высота");
                    double offset = mmToFeet(50);
                    width.Set(pipe.Diameter + offset);
                    height.Set(pipe.Diameter + offset);
                }
            }
            t.Commit();
            return Result.Succeeded;
        }

        private static double mmToFeet(double lenght)
        {
            return UnitUtils.ConvertToInternalUnits(lenght, UnitTypeId.Millimeters);
        }

        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
    }
}
