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

            List<Definition> paramDefinitionsList = new List<Definition>();
            var iterador = doc.ParameterBindings.ForwardIterator();
            while (iterador.MoveNext())
            {
                Definition wdwdwd = iterador.Key;
                paramDefinitionsList.Add(iterador.Key);
            }

            //Вызов формы
            ViewScheduleSheetSpreaderWPF viewScheduleSheetSpreaderWPF = new ViewScheduleSheetSpreaderWPF(doc, viewScheduleList, titleBlockFamilysList, paramDefinitionsList);
            viewScheduleSheetSpreaderWPF.ShowDialog();
            if (viewScheduleSheetSpreaderWPF.DialogResult != true)
            {
                return Result.Cancelled;
            }

            List<ViewSchedule> selectedViewScheduleList = viewScheduleSheetSpreaderWPF.SelectedViewScheduleCollection.ToList();
            FamilySymbol firstSheetType = viewScheduleSheetSpreaderWPF.FirstSheetType;
            FamilySymbol followingSheetType = viewScheduleSheetSpreaderWPF.FollowingSheetType;
            Parameter sheetFormatParameter = viewScheduleSheetSpreaderWPF.SheetFormatParameter;
            Definition groupingParameterDefinition = viewScheduleSheetSpreaderWPF.GroupingParameterDefinition;
            double xOffset = viewScheduleSheetSpreaderWPF.XOffset / 304.8;
            double yOffset = viewScheduleSheetSpreaderWPF.YOffset / 304.8;
            string sheetSizeVariantName = viewScheduleSheetSpreaderWPF.SheetSizeVariantName;

            using (TransactionGroup tg = new TransactionGroup(doc))
            {
                tg.Start("Спецификации на листы");
                foreach (ViewSchedule viewSchedule in selectedViewScheduleList)
                {
                    List<ElementId> allElemInViewSchedule = new FilteredElementCollector(doc, viewSchedule.Id).ToElementIds().ToList();
                    TableData tableData = viewSchedule.GetTableData();
                    TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);
                    int nRows = sectionData.NumberOfRows;

                    double rowHightSumm = 0;
                    int sheetNumber = 69;
                    int startSheetNumber = sheetNumber;
                    ViewSheet currentViewSheet = null;
                    //Добавить параметр группировки

                    XYZ viewLocation = new XYZ(xOffset, yOffset, 0);
                    for (int i = 0; i < nRows; i++)
                    {
                        if(currentViewSheet == null && startSheetNumber == sheetNumber)
                        {
                            using (Transaction t = new Transaction(doc))
                            {
                                t.Start("Создание 1-го листа");
                                currentViewSheet = ViewSheet.Create(doc, firstSheetType.Id);
                                ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_TitleBlocks);
                                FamilyInstance firstViewSheetFrameFamilyInstance = doc.GetElement(currentViewSheet.GetDependentElements(filter).First()) as FamilyInstance;
                                firstViewSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set(sheetNumber.ToString());
                                firstViewSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NAME).Set("Спецификация оборудования");
                                if (sheetSizeVariantName == "radioButton_Instance")
                                {
                                    firstViewSheetFrameFamilyInstance.get_Parameter(sheetFormatParameter.Definition).Set(3);
                                }
                                t.Commit();
                            }
                        }
                        using (Transaction t = new Transaction(doc))
                        {
                            t.Start("Проверка размера спецификации");

                            //ScheduleDefinition definition = viewSchedule.Definition;
                            //ScheduleField field = FindField(viewSchedule, sortParam.Id);
                            //if (field == null)
                            //{
                            //    field = definition.AddField(ScheduleFieldType.Instance, sortParam.Id);
                            //    field.IsHidden = true;
                            //}
                            //ScheduleFilter filter = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, sheetNumber.ToString());
                            //definition.AddFilter(filter);

                            ScheduleSheetInstance newViewSchedule = ScheduleSheetInstance.Create(doc, currentViewSheet.Id, viewSchedule.Id, viewLocation);
                            BoundingBoxXYZ bb = newViewSchedule.get_BoundingBox(currentViewSheet);
                            double scheduleHeight = bb.Max.Y - bb.Min.Y - 0.00695538057742782152230971128609;
                            double scheduleHeightMM = scheduleHeight * 304.8;
                            t.Commit();
                        }    
                    }


                    //using (Transaction t = new Transaction(doc))
                    //{
                    //    for (int i = 0; i < nRows; i++)
                    //    {
                    //        double rowHeight = sectionData.GetRowHeight(i);
                    //        rowHightSumm += rowHeight;
                    //        if (startSheetNumber == sheetNumber & rowHightSumm <= 200 / 380.4)
                    //        {
                    //            if (sectionData.CanRemoveRow(i) == true)
                    //            {
                    //                List<ElementId> idsInRow = null;
                    //                t.Start("Test");
                    //                sectionData.RemoveRow(i);
                    //                List<ElementId> elesTwo = new FilteredElementCollector(doc, viewSchedule.Id).ToElementIds().ToList();
                    //                idsInRow = allElemInViewSchedule.Except(elesTwo).ToList();
                    //                t.RollBack();
                    //                t.Start("Test");
                    //                foreach (ElementId elementId in idsInRow)
                    //                {
                    //                    doc.GetElement(elementId).LookupParameter("ADSK_Группирование").Set(sheetNumber.ToString());
                    //                }
                    //                t.Commit();
                    //            }
                    //        }
                    //        else if (startSheetNumber != sheetNumber & rowHightSumm <= 240 / 380.4)
                    //        {
                    //            if (sectionData.CanRemoveRow(i) == true)
                    //            {
                    //                List<ElementId> idsInRow = null;
                    //                t.Start("Test");
                    //                sectionData.RemoveRow(i);
                    //                List<ElementId> elesTwo = new FilteredElementCollector(doc, viewSchedule.Id).ToElementIds().ToList();
                    //                idsInRow = allElemInViewSchedule.Except(elesTwo).ToList();
                    //                t.RollBack();
                    //                t.Start("Test");
                    //                foreach (ElementId elementId in idsInRow)
                    //                {
                    //                    doc.GetElement(elementId).LookupParameter("ADSK_Группирование").Set(sheetNumber.ToString());
                    //                }
                    //                t.Commit();
                    //            }
                    //        }
                    //        else
                    //        {
                    //            rowHightSumm = rowHeight;
                    //            sheetNumber += 1;

                    //            List<ElementId> idsInRow = null;
                    //            t.Start("Test");
                    //            sectionData.RemoveRow(i);
                    //            List<ElementId> elesTwo = new FilteredElementCollector(doc, viewSchedule.Id).ToElementIds().ToList();
                    //            idsInRow = allElemInViewSchedule.Except(elesTwo).ToList();
                    //            t.RollBack();
                    //            t.Start("Test");
                    //            foreach (ElementId elementId in idsInRow)
                    //            {
                    //                doc.GetElement(elementId).LookupParameter("ADSK_Группирование").Set(sheetNumber.ToString());
                    //            }
                    //            t.Commit();

                    //            t.Start("Test");
                    //            foreach (ElementId elementId in idsInRow)
                    //            {
                    //                doc.GetElement(elementId).LookupParameter("ADSK_Группирование").Set(sheetNumber.ToString());
                    //            }
                    //            t.Commit();
                    //        }
                    //    }
                    //}
                }
                tg.Assimilate();
            }

            return Result.Succeeded;
        }
        public static ScheduleField FindField(ViewSchedule schedule, ElementId paramId)
        {
            ScheduleDefinition definition = schedule.Definition;
            ScheduleField foundField = null;

            foreach (ScheduleFieldId fieldId in definition.GetFieldOrder())
            {
                foundField = definition.GetField(fieldId);
                if (foundField.ParameterId == paramId)
                {
                    return foundField;
                }
            }
            return null;
        }
    }
}
