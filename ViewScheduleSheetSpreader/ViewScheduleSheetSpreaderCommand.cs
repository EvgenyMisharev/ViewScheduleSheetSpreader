using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ViewScheduleSheetSpreader
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class ViewScheduleSheetSpreaderCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            int.TryParse(commandData.Application.Application.VersionNumber, out int versionNumber);
            if (versionNumber < 2023)
            {
                TaskDialog.Show("Revit", "К сожалению, возможность разделения спецификаций по листам доступна только с версии Revit 2023 и выше!");
                return Result.Cancelled;
            }

#if R2023 || R2024 || R2025
            List<ViewSchedule> viewScheduleList = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Schedules)
                            .OfClass(typeof(ViewSchedule))
                            .Cast<ViewSchedule>()
                            .Where(vs => vs.IsTitleblockRevisionSchedule == false)
                            .OrderBy(vs => vs.Name, new AlphanumComparatorFastString())
                            .ToList();

            List<Family> titleBlockFamilysList = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(f => f.FamilyCategory.Id.IntegerValue.Equals((int)BuiltInCategory.OST_TitleBlocks))
                .OrderBy(f => f.Name, new AlphanumComparatorFastString())
                .ToList();

            // Вызов формы
            ViewScheduleSheetSpreaderWPF viewScheduleSheetSpreaderWPF = new ViewScheduleSheetSpreaderWPF(doc, viewScheduleList, titleBlockFamilysList);
            viewScheduleSheetSpreaderWPF.ShowDialog();
            if (viewScheduleSheetSpreaderWPF.DialogResult != true)
            {
                return Result.Cancelled;
            }

            List<ViewSchedule> selectedViewScheduleList = viewScheduleSheetSpreaderWPF.SelectedViewScheduleCollection.ToList();
            FamilySymbol firstSheetType = viewScheduleSheetSpreaderWPF.FirstSheetType;
            FamilySymbol followingSheetType = viewScheduleSheetSpreaderWPF.FollowingSheetType;
            Parameter sheetFormatParameter = viewScheduleSheetSpreaderWPF.SheetFormatParameter;
            double xOffset = viewScheduleSheetSpreaderWPF.XOffset / 304.8;
            double yOffset = viewScheduleSheetSpreaderWPF.YOffset / 304.8;
            string sheetSizeVariantName = viewScheduleSheetSpreaderWPF.SheetSizeVariantName;
            int sheetNumber = viewScheduleSheetSpreaderWPF.FirstSheetNumber;
            string headerInSpecificationHeaderVariantName = viewScheduleSheetSpreaderWPF.HeaderInSpecificationHeaderVariantName;
            double specificationHeaderHeight = viewScheduleSheetSpreaderWPF.SpecificationHeaderHeight / 304.8;
            double placementHeightOnFirstSheet = 230 / 304.8;
            double placementHeightOnFollowingSheet = 270 / 304.8;
            string sheetNumberSuffix = viewScheduleSheetSpreaderWPF.SheetNumberSuffix;

            //Проверка разделения
            foreach (ViewSchedule viewSchedule in selectedViewScheduleList)
            {
                if (viewSchedule.IsSplit())
                {
                    TaskDialog.Show("Revit", $"Спецификация \"{viewSchedule.Name}\" уже разделена и не может быть размещена на листы!");
                    return Result.Cancelled;
                }
            }

            if (headerInSpecificationHeaderVariantName == "radioButton_No")
            {
                yOffset = yOffset - specificationHeaderHeight;
                placementHeightOnFirstSheet = placementHeightOnFirstSheet - specificationHeaderHeight;
                placementHeightOnFollowingSheet = placementHeightOnFollowingSheet - specificationHeaderHeight;

                int currentSheetNumber = sheetNumber;
                ViewSheet currentSheet = null;
                XYZ viewLocation = new XYZ(xOffset, yOffset, 0);
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Спецификации на листы");
                    ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_TitleBlocks);

                    foreach (ViewSchedule viewSchedule in selectedViewScheduleList)
                    {
                        ScheduleHeightsOnSheet scheduleHeightsOnSheet = viewSchedule.GetScheduleHeightsOnSheet();
                        double viewScheduleTitleHeight = scheduleHeightsOnSheet.TitleHeight;
                        double viewScheduleColumnHeaderHeight = scheduleHeightsOnSheet.ColumnHeaderHeight;
                        double viewScheduleBodyRowHeight = scheduleHeightsOnSheet.GetBodyRowHeights().Sum();
                        if (viewScheduleBodyRowHeight == 0) continue;

                        if (currentSheet == null)
                        {
                            currentSheet = ViewSheet.Create(doc, firstSheetType.Id);
                            FamilyInstance firstViewSheetFrameFamilyInstance = doc.GetElement(currentSheet.GetDependentElements(elementCategoryFilter).First()) as FamilyInstance;
                            firstViewSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set($"{currentSheetNumber}{sheetNumberSuffix}");
                            firstViewSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NAME).Set("Спецификация оборудования, изделий и материалов");
                            if (sheetSizeVariantName == "radioButton_Instance")
                            {
                                firstViewSheetFrameFamilyInstance.LookupParameter(sheetFormatParameter.Definition.Name).Set(3);
                            }

                            if (viewScheduleBodyRowHeight - placementHeightOnFirstSheet < 0)
                            {
                                ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation);
                                placementHeightOnFirstSheet = placementHeightOnFirstSheet - viewScheduleBodyRowHeight;
                                viewLocation = new XYZ(viewLocation.X, viewLocation.Y - viewScheduleBodyRowHeight, 0);
                            }
                            else
                            {
                                int segmentsCount = (int)Math.Ceiling(((viewScheduleBodyRowHeight - placementHeightOnFirstSheet) / placementHeightOnFollowingSheet));
                                List<double> viewScheduleHeightsList = new List<double>();
                                viewScheduleHeightsList.Add(placementHeightOnFirstSheet);
                                for (int i = 1; i < segmentsCount; i++)
                                {
                                    viewScheduleHeightsList.Add(placementHeightOnFollowingSheet);
                                }
                                viewSchedule.Split(viewScheduleHeightsList);
                                int realSegmentsCount = viewSchedule.GetSegmentCount();
                                for (int i = 0; i < realSegmentsCount; i++)
                                {
                                    double segH = viewSchedule.GetSegmentHeight(i);
                                    if (segH == double.MaxValue)
                                    {
                                        ScheduleSheetInstance scheduleSheetInstanceTMP = ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, i);
                                        BoundingBoxXYZ bb = scheduleSheetInstanceTMP.get_BoundingBox(currentSheet);
                                        double segmentHeight = new XYZ(bb.Max.X, bb.Max.Y, 0).DistanceTo(new XYZ(bb.Max.X, bb.Min.Y, 0)) - 0.01391076115;
                                        doc.Delete(scheduleSheetInstanceTMP.Id);
                                        if (segmentHeight <= 0)
                                        {
                                            viewSchedule.DeleteSegment(i);
                                            realSegmentsCount--;
                                        }
                                    }
                                }

                                ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, 0);

                                double lastSegmentSegmentHeight = 0;
                                viewLocation = new XYZ(xOffset, yOffset, 0);
                                currentSheetNumber++;
                                for (int i = 1; i < realSegmentsCount; i++)
                                {
                                    currentSheet = ViewSheet.Create(doc, followingSheetType.Id);
                                    FamilyInstance currentSheetFrameFamilyInstance = doc.GetElement(currentSheet.GetDependentElements(elementCategoryFilter).First()) as FamilyInstance;
                                    currentSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set($"{currentSheetNumber}{sheetNumberSuffix}");
                                    currentSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NAME).Set("Спецификация оборудования, изделий и материалов");
                                    if (sheetSizeVariantName == "radioButton_Instance")
                                    {
                                        currentSheetFrameFamilyInstance.LookupParameter(sheetFormatParameter.Definition.Name).Set(3);
                                    }
                                    ScheduleSheetInstance scheduleSheetInstance = ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, i);
                                    currentSheetNumber++;
                                    if (i == realSegmentsCount - 1)
                                    {
                                        BoundingBoxXYZ bb = scheduleSheetInstance.get_BoundingBox(currentSheet);
                                        lastSegmentSegmentHeight = new XYZ(bb.Max.X, bb.Max.Y, 0).DistanceTo(new XYZ(bb.Max.X, bb.Min.Y, 0)) - 0.01391076115;
                                    }
                                }
                                placementHeightOnFollowingSheet = placementHeightOnFollowingSheet - lastSegmentSegmentHeight;
                                viewLocation = new XYZ(viewLocation.X, viewLocation.Y - lastSegmentSegmentHeight, 0);
                            }
                        }
                        else
                        {
                            if (currentSheetNumber == sheetNumber)
                            {
                                if (viewScheduleBodyRowHeight - placementHeightOnFirstSheet < 0)
                                {
                                    ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation);
                                    placementHeightOnFirstSheet = placementHeightOnFirstSheet - viewScheduleBodyRowHeight;
                                    viewLocation = new XYZ(viewLocation.X, viewLocation.Y - viewScheduleBodyRowHeight, 0);
                                }
                                else
                                {
                                    int segmentsCount = (int)Math.Ceiling(((viewScheduleBodyRowHeight - placementHeightOnFirstSheet) / placementHeightOnFollowingSheet));
                                    List<double> viewScheduleHeightsList = new List<double>();
                                    viewScheduleHeightsList.Add(placementHeightOnFirstSheet);
                                    for (int i = segmentsCount; i > 0; i--)
                                    {
                                        viewScheduleHeightsList.Add(placementHeightOnFollowingSheet);
                                    }
                                    viewSchedule.Split(viewScheduleHeightsList);
                                    int realSegmentsCount = viewSchedule.GetSegmentCount();
                                    for (int i = 0; i < realSegmentsCount; i++)
                                    {
                                        double segH = viewSchedule.GetSegmentHeight(i);
                                        if (segH == double.MaxValue)
                                        {
                                            ScheduleSheetInstance scheduleSheetInstanceTMP = ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, i);
                                            BoundingBoxXYZ bb = scheduleSheetInstanceTMP.get_BoundingBox(currentSheet);
                                            double segmentHeight = new XYZ(bb.Max.X, bb.Max.Y, 0).DistanceTo(new XYZ(bb.Max.X, bb.Min.Y, 0)) - 0.01391076115;
                                            doc.Delete(scheduleSheetInstanceTMP.Id);
                                            if (segmentHeight <= 0)
                                            {
                                                viewSchedule.DeleteSegment(i);
                                                realSegmentsCount--;
                                            }
                                        }
                                    }

                                    ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, 0);
                                    viewLocation = new XYZ(xOffset, yOffset, 0);

                                    currentSheetNumber++;
                                    double lastSegmentSegmentHeight = 0;
                                    for (int i = 1; i < realSegmentsCount; i++)
                                    {
                                        currentSheet = ViewSheet.Create(doc, followingSheetType.Id);
                                        FamilyInstance currentSheetFrameFamilyInstance = doc.GetElement(currentSheet.GetDependentElements(elementCategoryFilter).First()) as FamilyInstance;
                                        currentSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set($"{currentSheetNumber}{sheetNumberSuffix}");
                                        currentSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NAME).Set("Спецификация оборудования, изделий и материалов");
                                        if (sheetSizeVariantName == "radioButton_Instance")
                                        {
                                            currentSheetFrameFamilyInstance.LookupParameter(sheetFormatParameter.Definition.Name).Set(3);
                                        }
                                        ScheduleSheetInstance scheduleSheetInstance = ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, i);
                                        currentSheetNumber++;
                                        if (i == realSegmentsCount - 1)
                                        {
                                            BoundingBoxXYZ bb = scheduleSheetInstance.get_BoundingBox(currentSheet);
                                            lastSegmentSegmentHeight = new XYZ(bb.Max.X, bb.Max.Y, 0).DistanceTo(new XYZ(bb.Max.X, bb.Min.Y, 0)) - 0.01391076115;
                                        }
                                    }
                                    placementHeightOnFollowingSheet = placementHeightOnFollowingSheet - lastSegmentSegmentHeight;
                                    viewLocation = new XYZ(viewLocation.X, viewLocation.Y - lastSegmentSegmentHeight, 0);
                                }
                            }
                            else
                            {
                                if (viewScheduleBodyRowHeight - placementHeightOnFollowingSheet < 0)
                                {
                                    ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation);
                                    placementHeightOnFollowingSheet = placementHeightOnFollowingSheet - viewScheduleBodyRowHeight;
                                    viewLocation = new XYZ(viewLocation.X, viewLocation.Y - viewScheduleBodyRowHeight, 0);
                                }
                                else
                                {
                                    int segmentsCount = (int)Math.Ceiling(((viewScheduleBodyRowHeight - placementHeightOnFollowingSheet) / (270 / 304.8 - specificationHeaderHeight)));
                                    List<double> viewScheduleHeightsList = new List<double>();
                                    viewScheduleHeightsList.Add(placementHeightOnFollowingSheet);

                                    placementHeightOnFollowingSheet = 270 / 304.8 - specificationHeaderHeight;
                                    for (int i = segmentsCount; i > 0; i--)
                                    {
                                        viewScheduleHeightsList.Add(placementHeightOnFollowingSheet);
                                    }
                                    viewSchedule.Split(viewScheduleHeightsList);
                                    int realSegmentsCount = viewSchedule.GetSegmentCount();
                                    for (int i = 0; i < realSegmentsCount; i++)
                                    {
                                        double segH = viewSchedule.GetSegmentHeight(i);
                                        if (segH == double.MaxValue)
                                        {
                                            ScheduleSheetInstance scheduleSheetInstanceTMP = ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, i);
                                            BoundingBoxXYZ bb = scheduleSheetInstanceTMP.get_BoundingBox(currentSheet);
                                            double segmentHeight = new XYZ(bb.Max.X, bb.Max.Y, 0).DistanceTo(new XYZ(bb.Max.X, bb.Min.Y, 0)) - 0.01391076115;
                                            doc.Delete(scheduleSheetInstanceTMP.Id);
                                            if (segmentHeight <= 0)
                                            {
                                                viewSchedule.DeleteSegment(i);
                                                realSegmentsCount--;
                                            }
                                        }
                                    }

                                    ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, 0);
                                    viewLocation = new XYZ(xOffset, yOffset, 0);

                                    double lastSegmentSegmentHeight = 0;
                                    for (int i = 1; i < realSegmentsCount; i++)
                                    {
                                        currentSheet = ViewSheet.Create(doc, followingSheetType.Id);
                                        FamilyInstance currentSheetFrameFamilyInstance = doc.GetElement(currentSheet.GetDependentElements(elementCategoryFilter).First()) as FamilyInstance;
                                        currentSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set($"{currentSheetNumber}{sheetNumberSuffix}");
                                        currentSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NAME).Set("Спецификация оборудования, изделий и материалов");
                                        if (sheetSizeVariantName == "radioButton_Instance")
                                        {
                                            currentSheetFrameFamilyInstance.LookupParameter(sheetFormatParameter.Definition.Name).Set(3);
                                        }
                                        ScheduleSheetInstance scheduleSheetInstance = ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, i);
                                        currentSheetNumber++;
                                        if (i == realSegmentsCount - 1)
                                        {
                                            BoundingBoxXYZ bb = scheduleSheetInstance.get_BoundingBox(currentSheet);
                                            lastSegmentSegmentHeight = new XYZ(bb.Max.X, bb.Max.Y, 0).DistanceTo(new XYZ(bb.Max.X, bb.Min.Y, 0)) - 0.01391076115;
                                        }
                                    }
                                    placementHeightOnFollowingSheet = placementHeightOnFollowingSheet - lastSegmentSegmentHeight;
                                    viewLocation = new XYZ(viewLocation.X, viewLocation.Y - lastSegmentSegmentHeight, 0);
                                }
                            }
                        }
                    }
                    t.Commit();
                }
            }
            else
            {
                int currentSheetNumber = sheetNumber;
                ViewSheet currentSheet = null;
                XYZ viewLocation = new XYZ(xOffset, yOffset, 0);
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Спецификации на листы");
                    ElementCategoryFilter elementCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_TitleBlocks);

                    foreach (ViewSchedule viewSchedule in selectedViewScheduleList)
                    {
                        ScheduleHeightsOnSheet scheduleHeightsOnSheet = viewSchedule.GetScheduleHeightsOnSheet();
                        double viewScheduleTitleHeight = scheduleHeightsOnSheet.TitleHeight;
                        double viewScheduleColumnHeaderHeight = scheduleHeightsOnSheet.ColumnHeaderHeight;
                        double viewScheduleBodyRowHeight = scheduleHeightsOnSheet.GetBodyRowHeights().Sum();
                        if (viewScheduleBodyRowHeight == 0) continue;

                        placementHeightOnFirstSheet = placementHeightOnFirstSheet - viewScheduleTitleHeight - viewScheduleColumnHeaderHeight;
                        placementHeightOnFollowingSheet = placementHeightOnFollowingSheet - viewScheduleTitleHeight - viewScheduleColumnHeaderHeight;

                        if (currentSheet == null)
                        {
                            currentSheet = ViewSheet.Create(doc, firstSheetType.Id);
                            FamilyInstance firstViewSheetFrameFamilyInstance = doc.GetElement(currentSheet.GetDependentElements(elementCategoryFilter).First()) as FamilyInstance;
                            firstViewSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set($"{currentSheetNumber}{sheetNumberSuffix}");
                            firstViewSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NAME).Set("Спецификация оборудования, изделий и материалов");
                            if (sheetSizeVariantName == "radioButton_Instance")
                            {
                                firstViewSheetFrameFamilyInstance.LookupParameter(sheetFormatParameter.Definition.Name).Set(3);
                            }

                            if (viewScheduleBodyRowHeight - placementHeightOnFirstSheet < 0)
                            {
                                ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation);
                                placementHeightOnFirstSheet = placementHeightOnFirstSheet - viewScheduleBodyRowHeight;
                                viewLocation = new XYZ(viewLocation.X, viewLocation.Y - viewScheduleBodyRowHeight, 0);
                            }
                            else
                            {
                                int segmentsCount = (int)Math.Ceiling(((viewScheduleBodyRowHeight - placementHeightOnFirstSheet) / placementHeightOnFollowingSheet));
                                List<double> viewScheduleHeightsList = new List<double>();
                                viewScheduleHeightsList.Add(placementHeightOnFirstSheet);
                                for (int i = 1; i < segmentsCount; i++)
                                {
                                    viewScheduleHeightsList.Add(placementHeightOnFollowingSheet);
                                }
                                viewSchedule.Split(viewScheduleHeightsList);
                                int realSegmentsCount = viewSchedule.GetSegmentCount();

                                ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, 0);

                                double lastSegmentSegmentHeight = 0;
                                viewLocation = new XYZ(xOffset, yOffset, 0);
                                currentSheetNumber++;
                                for (int i = 1; i < realSegmentsCount; i++)
                                {
                                    currentSheet = ViewSheet.Create(doc, followingSheetType.Id);
                                    FamilyInstance currentSheetFrameFamilyInstance = doc.GetElement(currentSheet.GetDependentElements(elementCategoryFilter).First()) as FamilyInstance;
                                    currentSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set($"{currentSheetNumber}{sheetNumberSuffix}");
                                    currentSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NAME).Set("Спецификация оборудования, изделий и материалов");
                                    if (sheetSizeVariantName == "radioButton_Instance")
                                    {
                                        currentSheetFrameFamilyInstance.LookupParameter(sheetFormatParameter.Definition.Name).Set(3);
                                    }
                                    ScheduleSheetInstance scheduleSheetInstance = ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, i);
                                    currentSheetNumber++;
                                    if (i == realSegmentsCount - 1)
                                    {
                                        BoundingBoxXYZ bb = scheduleSheetInstance.get_BoundingBox(currentSheet);
                                        lastSegmentSegmentHeight = new XYZ(bb.Max.X, bb.Max.Y, 0).DistanceTo(new XYZ(bb.Max.X, bb.Min.Y, 0)) - 0.01391076115;
                                    }
                                }
                                placementHeightOnFollowingSheet = placementHeightOnFollowingSheet - lastSegmentSegmentHeight;
                                viewLocation = new XYZ(viewLocation.X, viewLocation.Y - lastSegmentSegmentHeight, 0);
                            }
                        }
                        else
                        {
                            if (currentSheetNumber == sheetNumber)
                            {
                                if (viewScheduleBodyRowHeight - placementHeightOnFirstSheet < 0)
                                {
                                    ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation);
                                    placementHeightOnFirstSheet = placementHeightOnFirstSheet - viewScheduleBodyRowHeight;
                                    viewLocation = new XYZ(viewLocation.X, viewLocation.Y - viewScheduleBodyRowHeight, 0);
                                }
                                else
                                {
                                    int segmentsCount = (int)Math.Ceiling(((viewScheduleBodyRowHeight - placementHeightOnFirstSheet) / placementHeightOnFollowingSheet));
                                    List<double> viewScheduleHeightsList = new List<double>();
                                    viewScheduleHeightsList.Add(placementHeightOnFirstSheet);
                                    for (int i = segmentsCount; i > 0; i--)
                                    {
                                        viewScheduleHeightsList.Add(placementHeightOnFollowingSheet);
                                    }
                                    viewSchedule.Split(viewScheduleHeightsList);
                                    int realSegmentsCount = viewSchedule.GetSegmentCount();

                                    ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, 0);
                                    viewLocation = new XYZ(xOffset, yOffset, 0);

                                    currentSheetNumber++;
                                    double lastSegmentSegmentHeight = 0;
                                    for (int i = 1; i < realSegmentsCount; i++)
                                    {
                                        currentSheet = ViewSheet.Create(doc, followingSheetType.Id);
                                        FamilyInstance currentSheetFrameFamilyInstance = doc.GetElement(currentSheet.GetDependentElements(elementCategoryFilter).First()) as FamilyInstance;
                                        currentSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set($"{currentSheetNumber}{sheetNumberSuffix}");
                                        currentSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NAME).Set("Спецификация оборудования, изделий и материалов");
                                        if (sheetSizeVariantName == "radioButton_Instance")
                                        {
                                            currentSheetFrameFamilyInstance.LookupParameter(sheetFormatParameter.Definition.Name).Set(3);
                                        }
                                        ScheduleSheetInstance scheduleSheetInstance = ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, i);
                                        currentSheetNumber++;
                                        if (i == realSegmentsCount - 1)
                                        {
                                            BoundingBoxXYZ bb = scheduleSheetInstance.get_BoundingBox(currentSheet);
                                            lastSegmentSegmentHeight = new XYZ(bb.Max.X, bb.Max.Y, 0).DistanceTo(new XYZ(bb.Max.X, bb.Min.Y, 0)) - 0.01391076115;
                                        }
                                    }
                                    placementHeightOnFollowingSheet = placementHeightOnFollowingSheet - lastSegmentSegmentHeight;
                                    viewLocation = new XYZ(viewLocation.X, viewLocation.Y - lastSegmentSegmentHeight, 0);
                                }
                            }
                            else
                            {
                                if (viewScheduleBodyRowHeight - placementHeightOnFollowingSheet < 0)
                                {
                                    ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation);
                                    placementHeightOnFollowingSheet = placementHeightOnFollowingSheet - viewScheduleBodyRowHeight;
                                    viewLocation = new XYZ(viewLocation.X, viewLocation.Y - viewScheduleBodyRowHeight, 0);
                                }
                                else
                                {
                                    int segmentsCount = (int)Math.Ceiling(((viewScheduleBodyRowHeight - placementHeightOnFollowingSheet) / (270 / 304.8 - viewScheduleTitleHeight - viewScheduleColumnHeaderHeight)));
                                    List<double> viewScheduleHeightsList = new List<double>();
                                    viewScheduleHeightsList.Add(placementHeightOnFollowingSheet);

                                    placementHeightOnFollowingSheet = 270 / 304.8 - viewScheduleTitleHeight - viewScheduleColumnHeaderHeight;
                                    for (int i = segmentsCount; i > 0; i--)
                                    {
                                        viewScheduleHeightsList.Add(placementHeightOnFollowingSheet);
                                    }
                                    viewSchedule.Split(viewScheduleHeightsList);
                                    int realSegmentsCount = viewSchedule.GetSegmentCount();

                                    ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, 0);
                                    viewLocation = new XYZ(xOffset, yOffset, 0);

                                    double lastSegmentSegmentHeight = 0;
                                    for (int i = 1; i < realSegmentsCount; i++)
                                    {
                                        currentSheet = ViewSheet.Create(doc, followingSheetType.Id);
                                        FamilyInstance currentSheetFrameFamilyInstance = doc.GetElement(currentSheet.GetDependentElements(elementCategoryFilter).First()) as FamilyInstance;
                                        currentSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set($"{currentSheetNumber}{sheetNumberSuffix}");
                                        currentSheetFrameFamilyInstance.get_Parameter(BuiltInParameter.SHEET_NAME).Set("Спецификация оборудования, изделий и материалов");
                                        if (sheetSizeVariantName == "radioButton_Instance")
                                        {
                                            currentSheetFrameFamilyInstance.LookupParameter(sheetFormatParameter.Definition.Name).Set(3);
                                        }
                                        ScheduleSheetInstance scheduleSheetInstance = ScheduleSheetInstance.Create(doc, currentSheet.Id, viewSchedule.Id, viewLocation, i);
                                        currentSheetNumber++;
                                        if (i == realSegmentsCount - 1)
                                        {
                                            BoundingBoxXYZ bb = scheduleSheetInstance.get_BoundingBox(currentSheet);
                                            lastSegmentSegmentHeight = new XYZ(bb.Max.X, bb.Max.Y, 0).DistanceTo(new XYZ(bb.Max.X, bb.Min.Y, 0)) - 0.01391076115;
                                        }
                                    }
                                    placementHeightOnFollowingSheet = placementHeightOnFollowingSheet - lastSegmentSegmentHeight;
                                    viewLocation = new XYZ(viewLocation.X, viewLocation.Y - lastSegmentSegmentHeight, 0);
                                }
                            }
                        }
                    }
                    t.Commit();
                }
            }
#endif
            return Result.Succeeded;
        }
    }
}
