using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ViewScheduleSheetSpreader
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class ViewScheduleSheetSpreaderCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            List<ViewSchedule> viewScheduleList = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Schedules)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .OrderBy(vs => vs.Name, new AlphanumComparatorFastString())
                .ToList();

            List<Family> titleBlockFamilysList = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(f => f.FamilyCategory.Id.IntegerValue.Equals((int)BuiltInCategory.OST_TitleBlocks))
                .OrderBy(f => f.Name, new AlphanumComparatorFastString())
                .ToList();

            //Вызов формы
            ViewScheduleSheetSpreaderWPF viewScheduleSheetSpreaderWPF = new ViewScheduleSheetSpreaderWPF(doc, viewScheduleList, titleBlockFamilysList);
            viewScheduleSheetSpreaderWPF.ShowDialog();
            if (viewScheduleSheetSpreaderWPF.DialogResult != true)
            {
                return Result.Cancelled;
            }
            return Result.Succeeded;
        }
    }
}
