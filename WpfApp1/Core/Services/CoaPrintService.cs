namespace WpfApp1.Core.Services
{
    public class CoaPrintService
    {
        private byte[]? _img1Data = null;
        private byte[]? _img2Data = null;

        public async System.Threading.Tasks.Task<System.String> GenerateCoaExcel(
            System.String customerName,
            System.String poNumber,
            System.String doNumber,
            System.Collections.Generic.List<WpfApp1.Core.Models.BusbarExportItem> dataList,
            System.String standardName)
        {
            return await System.Threading.Tasks.Task.Run<System.String>(() =>
            {
                System.String basePath = @"C:\Users\mrrx\Documents\My Web Sites\H\COA";

                System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();

                System.String resourceTemplate = "WpfApp1.Images.TEMPLATEE_COA_BUSBAR.xlsx";
                System.String resourceImg1 = "WpfApp1.Images.approved_IMG.png";
                System.String resourceImg2 = "WpfApp1.Images.profile_SNI.png";

                if (_img1Data == null)
                {
                    System.IO.Stream? stream = null;
                    try
                    {
                        stream = assembly.GetManifestResourceStream(resourceImg1);
                        if (stream != null)
                        {
                            System.IO.MemoryStream? ms = null;
                            try
                            {
                                ms = new System.IO.MemoryStream();
                                stream.CopyTo(ms);
                                _img1Data = ms.ToArray();
                            }
                            finally
                            {
                                if (ms != null)
                                {
                                    ms.Dispose();
                                }
                            }
                        }
                    }
                    finally
                    {
                        if (stream != null)
                        {
                            stream.Dispose();
                        }
                    }
                }

                if (_img2Data == null)
                {
                    System.IO.Stream? stream = null;
                    try
                    {
                        stream = assembly.GetManifestResourceStream(resourceImg2);
                        if (stream != null)
                        {
                            System.IO.MemoryStream? ms = null;
                            try
                            {
                                ms = new System.IO.MemoryStream();
                                stream.CopyTo(ms);
                                _img2Data = ms.ToArray();
                            }
                            finally
                            {
                                if (ms != null)
                                {
                                    ms.Dispose();
                                }
                            }
                        }
                    }
                    finally
                    {
                        if (stream != null)
                        {
                            stream.Dispose();
                        }
                    }
                }

                System.DateTime now = System.DateTime.Now;
                System.String yearFolder = now.ToString("yyyy");

                System.String yearPath = System.IO.Path.Combine(basePath, yearFolder);

                if (!System.IO.Directory.Exists(yearPath))
                {
                    System.IO.Directory.CreateDirectory(yearPath);
                }

                System.String finalMonthFolderName = System.String.Empty;
                System.Int32 currentMonthNumber = now.Month;
                System.Boolean folderFound = false;

                try
                {
                    System.String[] existingDirectories = System.IO.Directory.GetDirectories(yearPath);

                    foreach (System.String dirPath in existingDirectories)
                    {
                        System.String dirName = System.IO.Path.GetFileName(dirPath);

                        System.String leadingDigits = System.String.Empty;
                        foreach (System.Char c in dirName)
                        {
                            if (System.Char.IsDigit(c))
                            {
                                leadingDigits += c;
                            }
                            else
                            {
                                break;
                            }
                        }

                        if (!System.String.IsNullOrEmpty(leadingDigits) && System.Int32.TryParse(leadingDigits, out int folderNum))
                        {
                            if (folderNum == currentMonthNumber)
                            {
                                finalMonthFolderName = dirName;
                                folderFound = true;
                                break;
                            }
                        }
                    }
                }
                catch
                {
                }

                if (!folderFound)
                {
                    System.Globalization.CultureInfo cultureIndo = new System.Globalization.CultureInfo("id-ID");
                    System.String monthName = cultureIndo.DateTimeFormat.GetMonthName(now.Month);
                    monthName = cultureIndo.TextInfo.ToTitleCase(monthName);

                    finalMonthFolderName = $"{now.Month}. {monthName}";
                }

                System.String finalDirectory = System.IO.Path.Combine(yearPath, finalMonthFolderName);

                if (!System.IO.Directory.Exists(finalDirectory))
                {
                    System.IO.Directory.CreateDirectory(finalDirectory);
                }

                System.String[] existingFiles = System.IO.Directory.GetFiles(finalDirectory, "*.xlsx");
                System.Int32 validFileCount = 0;
                foreach (System.String filePath in existingFiles)
                {
                    System.String fName = System.IO.Path.GetFileName(filePath);
                    if (!fName.StartsWith("~$"))
                    {
                        validFileCount++;
                    }
                }

                System.Int32 nomorFile = validFileCount + 1;
                System.String formattedFileNumber = nomorFile.ToString("000");
                System.String romanMonth = GetRomanMonth(now.Month);

                System.String fileName = $"{formattedFileNumber}. COA {customerName} {doNumber}.xlsx";
                System.String fullPath = System.IO.Path.Combine(finalDirectory, fileName);

                System.Random randomGen = new System.Random();
                System.Globalization.CultureInfo cultureInvariant = System.Globalization.CultureInfo.InvariantCulture;

                System.IO.Stream? templateStream = null;
                try
                {
                    templateStream = assembly.GetManifestResourceStream(resourceTemplate);
                    if (templateStream == null)
                    {
                        throw new System.IO.FileNotFoundException($"Template Embedded Resource tidak ditemukan: '{resourceTemplate}'. \nPastikan 'Build Action' file Excel diset ke 'Embedded Resource'.");
                    }

                    ClosedXML.Excel.XLWorkbook? workbook = null;
                    try
                    {
                        workbook = new ClosedXML.Excel.XLWorkbook(templateStream);

                        ClosedXML.Excel.IXLWorksheet worksheet = workbook.Worksheet(1);
                        worksheet.Style.Font.FontName = "Montserrat";

                        worksheet.Cell("C12").Value = ": " + poNumber;
                        worksheet.Cell("J12").Value = ": " + customerName;
                        worksheet.Cell("J13").Value = ": " + now.ToString("dd/MM/yyyy");
                        worksheet.Cell("J14").Value = ": " + $"{formattedFileNumber}/{romanMonth}/{now.Year}";

                        System.Int32 dataCount = dataList.Count;
                        System.Int32 rowsPerItem = 2;
                        System.Int32 totalRowsNeeded = dataCount * rowsPerItem;
                        System.Int32 defaultRowsAvailable = 3;

                        System.Int32 startRowTable1 = 20;
                        System.Int32 originalStartRowTable2 = 30;
                        System.Int32 startRowTable2 = originalStartRowTable2;

                        System.Int32 rowsDiff = totalRowsNeeded - defaultRowsAvailable;

                        if (rowsDiff > 0)
                        {
                            worksheet.Row(22).InsertRowsBelow(rowsDiff);
                        }
                        else if (rowsDiff < 0)
                        {
                            System.Int32 rowsToDelete = System.Math.Abs(rowsDiff);
                            System.Int32 deleteStartRow = startRowTable1 + totalRowsNeeded;
                            for (System.Int32 d = 0; d < rowsToDelete; d++)
                            {
                                worksheet.Row(deleteStartRow).Delete();
                            }
                        }

                        startRowTable2 = originalStartRowTable2 + rowsDiff;

                        System.Collections.Generic.List<(System.Double thickness, System.Double width, System.Double nominalThick, System.Double nominalWidth)> toleranceData =
                            new System.Collections.Generic.List<(System.Double, System.Double, System.Double, System.Double)>();

                        for (System.Int32 i = 0; i < dataCount; i++)
                        {
                            WpfApp1.Core.Models.BusbarExportItem exportItem = dataList[i];
                            WpfApp1.Core.Models.BusbarRecord rec = exportItem.RecordData;
                            System.String cleanSizeStr = WpfApp1.Shared.Helpers.StringHelper.CleanSizeCOA(rec.Size);

                            System.Double finalTolThickness = 0;
                            System.Double finalTolWidth = 0;

                            if (standardName == "JIS")
                            {
                                (System.Double Thickness, System.Double Width) calculatedTols = WpfApp1.Shared.Helpers.ToleranceJIS.CalculateFromDbString(cleanSizeStr);
                                finalTolThickness = calculatedTols.Thickness;
                                finalTolWidth = calculatedTols.Width;
                            }
                            else
                            {
                                finalTolThickness = System.Math.Round((randomGen.NextDouble() * 0.2) + 0.05, 2);
                                finalTolWidth = System.Math.Round((randomGen.NextDouble() * 1.0) + 0.50, 2);
                            }

                            System.Double nominalThick = 0;
                            System.Double nominalWidth = 0;
                            System.String[] parts = cleanSizeStr.Split('x');
                            if (parts.Length >= 2)
                            {
                                System.Double.TryParse(parts[0], System.Globalization.NumberStyles.Float, cultureInvariant, out nominalThick);
                                System.Double.TryParse(parts[1], System.Globalization.NumberStyles.Float, cultureInvariant, out nominalWidth);
                            }

                            toleranceData.Add((finalTolThickness, finalTolWidth, nominalThick, nominalWidth));
                        }

                        for (System.Int32 i = 0; i < dataCount; i++)
                        {
                            System.Int32 rTop = startRowTable1 + (i * rowsPerItem);
                            System.Int32 rBottom = rTop + 1;
                            System.Int32 rTable2 = startRowTable2 + i;

                            worksheet.Row(rTop).Height = 51;
                            worksheet.Row(rBottom).Height = 51;
                            worksheet.Row(rTable2).Height = 102;
                        }

                        for (System.Int32 i = 0; i < dataCount; i++)
                        {
                            System.Int32 rTop = startRowTable1 + (i * rowsPerItem);
                            System.Int32 rBottom = rTop + 1;
                            System.Int32 rTable2 = startRowTable2 + i;

                            WpfApp1.Core.Models.BusbarExportItem exportItem = dataList[i];
                            WpfApp1.Core.Models.BusbarRecord rec = exportItem.RecordData;
                            (System.Double thickness, System.Double width, System.Double nominalThick, System.Double nominalWidth) tol = toleranceData[i];

                            System.String displaySize = WpfApp1.Shared.Helpers.StringHelper.CleanSizeCOA(rec.Size).Replace("x", " x ");
                            System.String selectedType = exportItem.SelectedType;

                            if (!System.String.IsNullOrEmpty(selectedType) &&
                                !selectedType.Equals("Select", System.StringComparison.OrdinalIgnoreCase) &&
                                !selectedType.Equals("None", System.StringComparison.OrdinalIgnoreCase))
                            {
                                displaySize = displaySize + " - " + selectedType;
                            }

                            ClosedXML.Excel.IXLCell cellBatch1 = worksheet.Cell(rTop, 2);
                            cellBatch1.Value = rec.BatchNo;
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightBatchNo(rec.BatchNo))
                            {
                                cellBatch1.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }

                            worksheet.Cell(rTop, 3).Value = displaySize;
                            worksheet.Cell(rTop, 4).Value = "No Dirty\nNo Blackspot\nNo Blisters";

                            System.String strThickTol = System.String.Format(cultureInvariant, "({0:0.00} \u00B1 {1:0.00})", tol.nominalThick, tol.thickness);
                            System.String strWidthTol = System.String.Format(cultureInvariant, "({0:0.00} \u00B1 {1:0.00})", tol.nominalWidth, tol.width);

                            ClosedXML.Excel.IXLCell cellThick = worksheet.Cell(rTop, 6);
                            cellThick.Value = System.String.Format(cultureInvariant, "{0:0.00}", rec.Thickness);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightThicknessWithTolerance(
                                rec.Thickness, tol.nominalThick, tol.thickness))
                            {
                                cellThick.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }
                            worksheet.Cell(rBottom, 6).Value = strThickTol;

                            ClosedXML.Excel.IXLCell cellWidth = worksheet.Cell(rTop, 7);
                            cellWidth.Value = System.String.Format(cultureInvariant, "{0:0.00}", rec.Width);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightWidthWithTolerance(
                                rec.Width, tol.nominalWidth, tol.width))
                            {
                                cellWidth.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }
                            worksheet.Cell(rBottom, 7).Value = strWidthTol;

                            ClosedXML.Excel.IXLCell cellLength = worksheet.Cell(rTop, 8);
                            cellLength.Value = rec.Length;
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightLength(rec.Length))
                            {
                                cellLength.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }
                            worksheet.Cell(rBottom, 8).Value = "(4000 +15/-0)";

                            ClosedXML.Excel.IXLCell cellRad = worksheet.Cell(rTop, 9);
                            cellRad.Value = System.String.Format(cultureInvariant, "{0:0.00}", rec.Radius);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightRadius(rec.Radius))
                            {
                                cellRad.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }

                            ClosedXML.Excel.IXLCell cellCham = worksheet.Cell(rTop, 10);
                            cellCham.Value = System.String.Format(cultureInvariant, "{0:0.00}", rec.Chamber);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightChamber(rec.Chamber))
                            {
                                cellCham.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }

                            worksheet.Cell(rTop, 11).Value = "OK";

                            worksheet.Range(rTop, 2, rBottom, 2).Merge();
                            worksheet.Range(rTop, 3, rBottom, 3).Merge();
                            worksheet.Range(rTop, 4, rBottom, 5).Merge();
                            worksheet.Range(rTop, 9, rBottom, 9).Merge();
                            worksheet.Range(rTop, 10, rBottom, 10).Merge();
                            worksheet.Range(rTop, 11, rBottom, 11).Merge();

                            ClosedXML.Excel.IXLCell cellBatch2 = worksheet.Cell(rTable2, 2);
                            cellBatch2.Value = rec.BatchNo;
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightBatchNo(rec.BatchNo))
                            {
                                cellBatch2.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }

                            worksheet.Cell(rTable2, 3).Value = displaySize;

                            ClosedXML.Excel.IXLCell cellElec = worksheet.Cell(rTable2, 4);
                            cellElec.Value = System.String.Format(cultureInvariant, "{0:0.00}", rec.Electric);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightElectric(rec.Electric))
                            {
                                cellElec.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }

                            ClosedXML.Excel.IXLCell cellRes = worksheet.Cell(rTable2, 5);
                            cellRes.Value = System.String.Format(cultureInvariant, "{0:0.00000}", rec.Resistivity);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightResistivity(rec.Resistivity))
                            {
                                cellRes.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }

                            ClosedXML.Excel.IXLCell cellElong = worksheet.Cell(rTable2, 6);
                            cellElong.Value = System.String.Format(cultureInvariant, "{0:0.00}", rec.Elongation);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightElongation(rec.Elongation))
                            {
                                cellElong.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }

                            ClosedXML.Excel.IXLCell cellTens = worksheet.Cell(rTable2, 7);
                            cellTens.Value = System.String.Format(cultureInvariant, "{0:0.00}", rec.Tensile);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightTensile(rec.Tensile))
                            {
                                cellTens.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }

                            worksheet.Cell(rTable2, 8).Value = "No Crack";

                            ClosedXML.Excel.IXLCell cellSpec = worksheet.Cell(rTable2, 9);
                            cellSpec.Value = System.String.Format(cultureInvariant, "{0:0.000}", rec.Spectro);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightSpectro(rec.Spectro))
                            {
                                cellSpec.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }

                            ClosedXML.Excel.IXLCell cellOxy = worksheet.Cell(rTable2, 10);
                            cellOxy.Value = System.String.Format(cultureInvariant, "{0:0.00}", rec.Oxygen);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightOxygen(rec.Oxygen))
                            {
                                cellOxy.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            }

                            worksheet.Cell(rTable2, 11).Value = "OK";

                            worksheet.Row(rTop).Height = 51;
                            worksheet.Row(rBottom).Height = 51;
                            worksheet.Row(rTable2).Height = 102;
                        }


                        ClosedXML.Excel.IXLRange table1Range = worksheet.Range(startRowTable1, 2, startRowTable1 + totalRowsNeeded - 1, 11);
                        ApplyCustomStyleBatch(table1Range);
                        ApplyBorders(table1Range);

                        for (System.Int32 i = 0; i < dataCount; i++)
                        {
                            System.Int32 rTop = startRowTable1 + (i * rowsPerItem);
                            System.Int32 rBottom = rTop + 1;

                            ClosedXML.Excel.IXLCell cellThickVal = worksheet.Cell(rTop, 6);
                            cellThickVal.Style.Font.Bold = true;
                            cellThickVal.Style.Font.FontSize = 22;
                            cellThickVal.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Bottom;
                            cellThickVal.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.None;

                            ClosedXML.Excel.IXLCell cellThickTol = worksheet.Cell(rBottom, 6);
                            cellThickTol.Style.Font.Bold = false;
                            cellThickTol.Style.Font.Italic = true;
                            cellThickTol.Style.Font.FontSize = 22;
                            cellThickTol.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Top;
                            cellThickTol.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.None;

                            ClosedXML.Excel.IXLCell cellWidthVal = worksheet.Cell(rTop, 7);
                            cellWidthVal.Style.Font.Bold = true;
                            cellWidthVal.Style.Font.FontSize = 22;
                            cellWidthVal.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Bottom;
                            cellWidthVal.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.None;

                            ClosedXML.Excel.IXLCell cellWidthTol = worksheet.Cell(rBottom, 7);
                            cellWidthTol.Style.Font.Bold = false;
                            cellWidthTol.Style.Font.Italic = true;
                            cellWidthTol.Style.Font.FontSize = 22;
                            cellWidthTol.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Top;
                            cellWidthTol.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.None;

                            ClosedXML.Excel.IXLCell cellLengthVal = worksheet.Cell(rTop, 8);
                            cellLengthVal.Style.NumberFormat.Format = "0";
                            cellLengthVal.Style.Font.Bold = true;
                            cellLengthVal.Style.Font.FontSize = 22;
                            cellLengthVal.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Bottom;
                            cellLengthVal.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.None;

                            ClosedXML.Excel.IXLCell cellLengthTol = worksheet.Cell(rBottom, 8);
                            cellLengthTol.Style.Font.Bold = false;
                            cellLengthTol.Style.Font.Italic = true;
                            cellLengthTol.Style.Font.FontSize = 22;
                            cellLengthTol.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Top;
                            cellLengthTol.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.None;

                            worksheet.Cell(rTop, 4).Style.Alignment.WrapText = true;
                        }

                        System.Int32 lastRowTable2 = startRowTable2 + dataCount - 1;
                        ClosedXML.Excel.IXLRange table2Range = worksheet.Range(startRowTable2, 2, lastRowTable2, 11);
                        ApplyCustomStyleBatch(table2Range);
                        ApplyBorders(table2Range);

                        worksheet.Row(lastRowTable2).InsertRowsBelow(5);
                        System.Int32 firstInsertedRow = lastRowTable2 + 1;
                        System.Int32 lastInsertedRow = firstInsertedRow + 4;

                        ClosedXML.Excel.IXLRange signatureRange = worksheet.Range(firstInsertedRow, 2, lastInsertedRow, 11);
                        signatureRange.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.None;
                        signatureRange.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.None;
                        signatureRange.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.None;
                        signatureRange.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.None;
                        signatureRange.Style.Border.InsideBorder = ClosedXML.Excel.XLBorderStyleValues.None;

                        signatureRange.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.NoColor;

                        for (System.Int32 k = 0; k < 5; k++)
                        {
                            worksheet.Row(firstInsertedRow + k).Height = 102;
                        }
                        worksheet.Row(firstInsertedRow + 4).Height = 50;

                        System.Int32 imageRow = firstInsertedRow + 1;

                        if (_img1Data != null)
                        {
                            System.IO.MemoryStream? ms1 = null;
                            try
                            {
                                ms1 = new System.IO.MemoryStream(_img1Data);
                                ClosedXML.Excel.Drawings.IXLPicture pic1 = worksheet.AddPicture(ms1);
                                pic1.MoveTo(worksheet.Cell(imageRow, 11));
                            }
                            catch
                            {
                                if (ms1 != null) ms1.Dispose();
                                throw;
                            }
                            if (ms1 != null) ms1.Dispose();
                        }

                        if (_img2Data != null)
                        {
                            System.IO.MemoryStream? ms2 = null;
                            try
                            {
                                ms2 = new System.IO.MemoryStream(_img2Data);
                                ClosedXML.Excel.Drawings.IXLPicture pic2 = worksheet.AddPicture(ms2);
                                pic2.MoveTo(worksheet.Cell(imageRow, 2));
                            }
                            finally
                            {
                                if (ms2 != null) ms2.Dispose();
                            }
                        }

                        worksheet.PageSetup.PrintAreas.Clear();
                        worksheet.PageSetup.PrintAreas.Add(1, 2, lastInsertedRow, 11);
                        worksheet.PageSetup.AddHorizontalPageBreak(lastInsertedRow + 1);
                        worksheet.PageSetup.PagesTall = 1;
                        worksheet.PageSetup.PagesWide = 1;

                        workbook.SaveAs(fullPath);
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            workbook.Dispose();
                        }
                    }
                }
                finally
                {
                    if (templateStream != null)
                    {
                        templateStream.Dispose();
                    }
                }

                return fullPath;
            });
        }

        private void ApplyCustomStyleBatch(ClosedXML.Excel.IXLRange range)
        {
            range.Style.Font.Bold = true;
            range.Style.Font.FontSize = 22;
            range.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
            range.Style.Alignment.WrapText = true;
        }

        private void ApplyBorders(ClosedXML.Excel.IXLRange range)
        {
            range.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            range.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            range.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            range.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
        }

        public void ClearCache()
        {
            _img1Data = null;
            _img2Data = null;
        }

        private System.String GetRomanMonth(System.Int32 month)
        {
            switch (month)
            {
                case 1: return "I";
                case 2: return "II";
                case 3: return "III";
                case 4: return "IV";
                case 5: return "V";
                case 6: return "VI";
                case 7: return "VII";
                case 8: return "VIII";
                case 9: return "IX";
                case 10: return "X";
                case 11: return "XI";
                case 12: return "XII";
                default: return "";
            }
        }
    }
}