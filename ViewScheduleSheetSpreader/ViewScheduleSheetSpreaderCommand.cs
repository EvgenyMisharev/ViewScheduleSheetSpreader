﻿using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ViewScheduleSheetSpreader
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class ViewScheduleSheetSpreaderCommand : IExternalCommand
    {
        ViewScheduleSheetSpreaderProgressBarWPF viewScheduleSheetSpreaderProgressBarWPF;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

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
            int sheetNumber = viewScheduleSheetSpreaderWPF.FirstSheetNumber;

            int startSheetNumber = sheetNumber;
            ViewSheet currentViewSheet = null;
            XYZ viewLocation = new XYZ(xOffset, yOffset, 0);

            Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
            newWindowThread.SetApartmentState(ApartmentState.STA);
            newWindowThread.IsBackground = true;
            newWindowThread.Start();
            int stepSpecificationProcessing = 0;
            int stepRowProcessing = 0;
            int stepPlacementSpecificationsOnSheets = 0;
            Thread.Sleep(100);
            viewScheduleSheetSpreaderProgressBarWPF.pb_SpecificationProcessingProgressBar.Dispatcher.Invoke(() => viewScheduleSheetSpreaderProgressBarWPF.pb_SpecificationProcessingProgressBar.Minimum = 0);
            viewScheduleSheetSpreaderProgressBarWPF.pb_SpecificationProcessingProgressBar.Dispatcher.Invoke(() => viewScheduleSheetSpreaderProgressBarWPF.pb_SpecificationProcessingProgressBar.Maximum = selectedViewScheduleList.Count);
            viewScheduleSheetSpreaderProgressBarWPF.pb_RowProcessingProgressBar.Dispatcher.Invoke(() => viewScheduleSheetSpreaderProgressBarWPF.pb_RowProcessingProgressBar.Minimum = 0);
            viewScheduleSheetSpreaderProgressBarWPF.pb_PlacementSpecificationsOnSheetsProgressBar.Dispatcher.Invoke(() => viewScheduleSheetSpreaderProgressBarWPF.pb_PlacementSpecificationsOnSheetsProgressBar.Minimum = 0);

            using (TransactionGroup tg = new TransactionGroup(doc))
            {
                tg.Start("Спецификации на листы");
                foreach (ViewSchedule viewSchedule in selectedViewScheduleList)
                {
                    TableData mainTableData = viewSchedule.GetTableData();
                    TableSectionData mainSectionData = mainTableData.GetSectionData(SectionType.Body);
                    int nRows = mainSectionData.NumberOfRows;
                    viewScheduleSheetSpreaderProgressBarWPF.pb_RowProcessingProgressBar.Dispatcher.Invoke(() => viewScheduleSheetSpreaderProgressBarWPF.pb_RowProcessingProgressBar.Maximum = nRows);
                    viewScheduleSheetSpreaderProgressBarWPF.pb_RowProcessingProgressBar.Dispatcher.Invoke(() => viewScheduleSheetSpreaderProgressBarWPF.pb_RowProcessingProgressBar.Value = stepRowProcessing);
                    viewScheduleSheetSpreaderProgressBarWPF.pb_PlacementSpecificationsOnSheetsProgressBar.Dispatcher.Invoke(() => viewScheduleSheetSpreaderProgressBarWPF.pb_PlacementSpecificationsOnSheetsProgressBar.Value = stepPlacementSpecificationsOnSheets);
                    
                    List<ElementId> allElemsInViewSchedule = new FilteredElementCollector(doc, viewSchedule.Id).ToElementIds().ToList();
                    List<List<ElementId>> elementIdListByRow = new List<List<ElementId>>();

                    using (Transaction t = new Transaction(doc))
                    {

                        for (int i = 0; i < nRows; i++)
                        {
                            stepRowProcessing++;
                            viewScheduleSheetSpreaderProgressBarWPF.pb_RowProcessingProgressBar.Dispatcher.Invoke(() => viewScheduleSheetSpreaderProgressBarWPF.pb_RowProcessingProgressBar.Value = stepRowProcessing);
                            if (mainSectionData.CanRemoveRow(i))
                            {
                                t.Start("Выбор элементов в строке");
                                mainSectionData.RemoveRow(i);
                                List<ElementId> elemsInViewScheduleWithoutRow = new FilteredElementCollector(doc, viewSchedule.Id).ToElementIds().ToList();
                                List<ElementId> idsInRow = allElemsInViewSchedule.Except(elemsInViewScheduleWithoutRow).ToList();
                                elementIdListByRow.Add(idsInRow);
                                t.RollBack();
                            }
                            else if (!mainSectionData.CanRemoveRow(i) && i == nRows - 1)
                            {
                                List<ElementId> elementIds = new List<ElementId>();
                                foreach (List<ElementId> idList in elementIdListByRow)
                                {
                                    foreach (ElementId id in idList)
                                    {
                                        elementIds.Add(id);
                                    }
                                }
                                List<ElementId> idsInRow = allElemsInViewSchedule.Except(elementIds).ToList();
                                elementIdListByRow.Add(idsInRow);
                            }
                        }
                    }

                    viewScheduleSheetSpreaderProgressBarWPF.pb_PlacementSpecificationsOnSheetsProgressBar.Dispatcher.Invoke(() => viewScheduleSheetSpreaderProgressBarWPF.pb_PlacementSpecificationsOnSheetsProgressBar.Maximum = elementIdListByRow.Count);
                    foreach (List<ElementId> idList in elementIdListByRow)
                    {
                        stepPlacementSpecificationsOnSheets++;
                        viewScheduleSheetSpreaderProgressBarWPF.pb_PlacementSpecificationsOnSheetsProgressBar.Dispatcher.Invoke(() => viewScheduleSheetSpreaderProgressBarWPF.pb_PlacementSpecificationsOnSheetsProgressBar.Value = stepPlacementSpecificationsOnSheets);
                        if (currentViewSheet == null && startSheetNumber == sheetNumber)
                        {
                            using (Transaction t = new Transaction(doc))
                            {
                                t.Start("Создание 1-го листа");
                                currentViewSheet = ViewSheet.Create(doc, firstSheetType.Id);
                                ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_TitleBlocks);
                                FamilyInstance firstViewSheetFrameFamilyInstance = doc.GetElement(currentViewSheet.GetDependentElements(elementCategoryFilter).First()) as FamilyInstance;
                                firstViewSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set(sheetNumber.ToString());
                                firstViewSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NAME).Set("Спецификация оборудования");
                                if (sheetSizeVariantName == "radioButton_Instance")
                                {
                                    firstViewSheetFrameFamilyInstance.LookupParameter(sheetFormatParameter.Definition.Name).Set(3);
                                }
                                t.Commit();
                            }
                        }
                        using (Transaction t = new Transaction(doc))
                        {
                            t.Start("Заполнение номера листа в группировку");
                            foreach (ElementId elementId in idList)
                            {
                                doc.GetElement(elementId).get_Parameter(groupingParameterDefinition).Set(sheetNumber.ToString());
                            }
                            t.Commit();

                            t.Start("Проверка высоты спецификации");
                            ViewSchedule tmpViewSchedule = doc.GetElement(viewSchedule.Duplicate(ViewDuplicateOption.Duplicate)) as ViewSchedule;
                            tmpViewSchedule.Name = $"{viewSchedule.Name}_Лист{sheetNumber}";
                            ScheduleDefinition definition = tmpViewSchedule.Definition;
                            ScheduleField field = FindField(doc, groupingParameterDefinition, definition);
                            if (field == null)
                            {
                                field = definition.AddField(ScheduleFieldType.Instance, (doc.GetElement(field.ParameterId) as ParameterElement).Id);
                                field.IsHidden = true;
                            }
                            ScheduleFilter filter = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, sheetNumber.ToString());
                            definition.AddFilter(filter);

                            if((yOffset - viewLocation.Y).Equals(0))
                            {
                                definition.ShowTitle = true;
                            }
                            else
                            {
                                definition.ShowTitle = false;
                            }

                            ScheduleSheetInstance tmpScheduleSheetInstance = ScheduleSheetInstance.Create(doc, currentViewSheet.Id, tmpViewSchedule.Id, viewLocation);
                            BoundingBoxXYZ bb = tmpScheduleSheetInstance.get_BoundingBox(currentViewSheet);
                            double scheduleHeight = bb.Max.Y - bb.Min.Y - 0.00695538057742782152230971128609;
                            double summScheduleHeight = scheduleHeight + (yOffset - viewLocation.Y);

                            if (startSheetNumber == sheetNumber && summScheduleHeight < 230 / 304.8)
                            {
                                if (stepPlacementSpecificationsOnSheets != elementIdListByRow.Count) t.RollBack();
                                else
                                {
                                    t.Commit();
                                    viewLocation = viewLocation - scheduleHeight * XYZ.BasisY;
                                } 

                            }
                            else if (startSheetNumber != sheetNumber && summScheduleHeight < 270 / 304.8)
                            {
                                if (stepPlacementSpecificationsOnSheets != elementIdListByRow.Count) t.RollBack();
                                else
                                {
                                    t.Commit();
                                    viewLocation = viewLocation - scheduleHeight * XYZ.BasisY;
                                }
                            }
                            else
                            {
                                t.Commit();
                                sheetNumber++;
                                viewLocation = new XYZ(xOffset, yOffset, 0);

                                t.Start("Заполнение номера листа в группировку");
                                foreach (ElementId elementId in idList)
                                {
                                    doc.GetElement(elementId).get_Parameter(groupingParameterDefinition).Set(sheetNumber.ToString());
                                }
                                t.Commit();
                                t.Start("Создание последующего листа");
                                currentViewSheet = ViewSheet.Create(doc, followingSheetType.Id);
                                ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_TitleBlocks);
                                FamilyInstance firstViewSheetFrameFamilyInstance = doc.GetElement(currentViewSheet.GetDependentElements(elementCategoryFilter).First()) as FamilyInstance;
                                firstViewSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set(sheetNumber.ToString());
                                firstViewSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NAME).Set("Спецификация оборудования");
                                if (sheetSizeVariantName == "radioButton_Instance")
                                {
                                    firstViewSheetFrameFamilyInstance.LookupParameter(sheetFormatParameter.Definition.Name).Set(3);
                                }
                                t.Commit();
                            }
                        }
                    }
                    stepSpecificationProcessing++;
                    viewScheduleSheetSpreaderProgressBarWPF.pb_SpecificationProcessingProgressBar.Dispatcher.Invoke(() => viewScheduleSheetSpreaderProgressBarWPF.pb_SpecificationProcessingProgressBar.Value = stepSpecificationProcessing);
                    stepRowProcessing = 0;
                    stepPlacementSpecificationsOnSheets = 0;
                }
                viewScheduleSheetSpreaderProgressBarWPF.Dispatcher.Invoke(() => viewScheduleSheetSpreaderProgressBarWPF.Close());
                tg.Assimilate();
            }
            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            TaskDialog.Show("Revit",$"Время выполнения: {elapsedTime}");
            return Result.Succeeded;
        }

        private static ScheduleField FindField(Document doc, Definition groupingParameterDefinition, ScheduleDefinition definition)
        {
            ScheduleField foundField = null;
            int fieldCount = definition.GetFieldCount();
            for(int i =0; i< fieldCount; i++)
            {
                foundField = definition.GetField(i);
                ParameterElement param = doc.GetElement(foundField.ParameterId) as ParameterElement;
                Definition def = param.GetDefinition();
                if(def.Name == groupingParameterDefinition.Name
                    && def.ParameterGroup == groupingParameterDefinition.ParameterGroup
                    && def.ParameterType == groupingParameterDefinition.ParameterType)
                {
                    return foundField;
                }
            }
            return null;
        }

        private void ThreadStartingPoint()
        {
            viewScheduleSheetSpreaderProgressBarWPF = new ViewScheduleSheetSpreaderProgressBarWPF();
            viewScheduleSheetSpreaderProgressBarWPF.Show();
            System.Windows.Threading.Dispatcher.Run();
        }
    }
}