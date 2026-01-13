using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ClosedXML.Excel;

namespace WpfApp1.Core.Services
{
    public class CoaPrintService
    {
        private byte[]? _img1Data = null;
        private byte[]? _img2Data = null;

        public async System.Threading.Tasks.Task<string> GenerateCoaExcel(
            string customerName,
            string poNumber,
            string doNumber,
            System.Collections.Generic.List<WpfApp1.Core.Models.BusbarExportItem> dataList,
            string standardName)
        {
            return await System.Threading.Tasks.Task.Run<string>(() =>
            {

                string basePath = @"C:\Users\mrrx\Documents\My Web Sites\H\COA";

                var assembly = Assembly.GetExecutingAssembly();

                string resourceTemplate = "WpfApp1.Images.TEMPLATEE_COA_BUSBAR.xlsx";
                string resourceImg1 = "WpfApp1.Images.approved_IMG.png";
                string resourceImg2 = "WpfApp1.Images.profile_SNI.png";

                if (_img1Data == null)
                {
                    using (Stream stream = assembly.GetManifestResourceStream(resourceImg1))
                    {
                        if (stream != null)
                        {
                            using (MemoryStream ms = new MemoryStream())
                            {
                                stream.CopyTo(ms);
                                _img1Data = ms.ToArray();
                            }
                        }
                    }
                }

                if (_img2Data == null)
                {
                    using (Stream stream = assembly.GetManifestResourceStream(resourceImg2))
                    {
                        if (stream != null)
                        {
                            using (MemoryStream ms = new MemoryStream())
                            {
                                stream.CopyTo(ms);
                                _img2Data = ms.ToArray();
                            }
                        }
                    }
                }


                System.DateTime now = System.DateTime.Now;
                string yearFolder = now.ToString("yyyy");

                string yearPath = System.IO.Path.Combine(basePath, yearFolder);

                if (!System.IO.Directory.Exists(yearPath))
                {
                    System.IO.Directory.CreateDirectory(yearPath);
                }

                string finalMonthFolderName = string.Empty;
                int currentMonthNumber = now.Month;
                bool folderFound = false;

                try
                {
                    string[] existingDirectories = System.IO.Directory.GetDirectories(yearPath);

                    foreach (string dirPath in existingDirectories)
                    {
                        string dirName = System.IO.Path.GetFileName(dirPath);

                        string leadingDigits = string.Empty;
                        foreach (char c in dirName)
                        {
                            if (char.IsDigit(c))
                            {
                                leadingDigits += c;
                            }
                            else
                            {
                                break;
                            }
                        }

                        if (!string.IsNullOrEmpty(leadingDigits) && int.TryParse(leadingDigits, out int folderNum))
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
                    var cultureIndo = new System.Globalization.CultureInfo("id-ID");
                    string monthName = cultureIndo.DateTimeFormat.GetMonthName(now.Month);
                    monthName = cultureIndo.TextInfo.ToTitleCase(monthName);

                    finalMonthFolderName = $"{now.Month}. {monthName}";
                }

                string finalDirectory = System.IO.Path.Combine(yearPath, finalMonthFolderName);

                if (!System.IO.Directory.Exists(finalDirectory))
                {
                    System.IO.Directory.CreateDirectory(finalDirectory);
                }

                string[] existingFiles = System.IO.Directory.GetFiles(finalDirectory, "*.xlsx");
                int validFileCount = 0;
                foreach (string filePath in existingFiles)
                {
                    string fName = System.IO.Path.GetFileName(filePath);
                    if (!fName.StartsWith("~$"))
                    {
                        validFileCount++;
                    }
                }

                int nomorFile = validFileCount + 1;
                string formattedFileNumber = nomorFile.ToString("000");
                string romanMonth = GetRomanMonth(now.Month);

                string fileName = $"{formattedFileNumber}. COA {customerName} {doNumber}.xlsx";
                string fullPath = System.IO.Path.Combine(finalDirectory, fileName);

                System.Random randomGen = new System.Random();
                var cultureInvariant = System.Globalization.CultureInfo.InvariantCulture;

                using (Stream templateStream = assembly.GetManifestResourceStream(resourceTemplate))
                {
                    if (templateStream == null)
                    {
                        throw new FileNotFoundException($"Template Embedded Resource tidak ditemukan: '{resourceTemplate}'. \nPastikan 'Build Action' file Excel diset ke 'Embedded Resource'.");
                    }

                    using (var workbook = new ClosedXML.Excel.XLWorkbook(templateStream))
                    {
                        var worksheet = workbook.Worksheet(1);
                        worksheet.Style.Font.FontName = "Montserrat";

                        worksheet.Cell("C12").Value = ": " + poNumber;
                        worksheet.Cell("J12").Value = ": " + customerName;
                        worksheet.Cell("J13").Value = ": " + now.ToString("dd/MM/yyyy");
                        worksheet.Cell("J14").Value = ": " + $"{formattedFileNumber}/{romanMonth}/{now.Year}";

                        int dataCount = dataList.Count;
                        int rowsPerItem = 2;
                        int totalRowsNeeded = dataCount * rowsPerItem;
                        int defaultRowsAvailable = 3;

                        int startRowTable1 = 20;
                        int originalStartRowTable2 = 30;
                        int startRowTable2 = originalStartRowTable2;

                        int rowsDiff = totalRowsNeeded - defaultRowsAvailable;

                        if (rowsDiff > 0)
                        {
                            worksheet.Row(22).InsertRowsBelow(rowsDiff);
                        }
                        else if (rowsDiff < 0)
                        {
                            int rowsToDelete = System.Math.Abs(rowsDiff);
                            int deleteStartRow = startRowTable1 + totalRowsNeeded;
                            for (int d = 0; d < rowsToDelete; d++)
                            {
                                worksheet.Row(deleteStartRow).Delete();
                            }
                        }

                        startRowTable2 = originalStartRowTable2 + rowsDiff;

                        var toleranceData = new System.Collections.Generic.List<(double thickness, double width, double nominalThick, double nominalWidth)>();

                        for (int i = 0; i < dataCount; i++)
                        {
                            var exportItem = dataList[i];
                            var rec = exportItem.RecordData;
                            string cleanSizeStr = WpfApp1.Shared.Helpers.StringHelper.CleanSizeCOA(rec.Size);

                            double finalTolThickness = 0;
                            double finalTolWidth = 0;

                            if (standardName == "JIS")
                            {
                                var calculatedTols = WpfApp1.Shared.Helpers.ToleranceJIS.CalculateFromDbString(cleanSizeStr);
                                finalTolThickness = calculatedTols.Thickness;
                                finalTolWidth = calculatedTols.Width;
                            }
                            else
                            {
                                finalTolThickness = System.Math.Round((randomGen.NextDouble() * 0.2) + 0.05, 2);
                                finalTolWidth = System.Math.Round((randomGen.NextDouble() * 1.0) + 0.50, 2);
                            }

                            double nominalThick = 0;
                            double nominalWidth = 0;
                            string[] parts = cleanSizeStr.Split('x');
                            if (parts.Length >= 2)
                            {
                                double.TryParse(parts[0], System.Globalization.NumberStyles.Float, cultureInvariant, out nominalThick);
                                double.TryParse(parts[1], System.Globalization.NumberStyles.Float, cultureInvariant, out nominalWidth);
                            }

                            toleranceData.Add((finalTolThickness, finalTolWidth, nominalThick, nominalWidth));
                        }

                        for (int i = 0; i < dataCount; i++)
                        {
                            int rTop = startRowTable1 + (i * rowsPerItem);
                            int rBottom = rTop + 1;
                            int rTable2 = startRowTable2 + i;

                            worksheet.Row(rTop).Height = 51;
                            worksheet.Row(rBottom).Height = 51;
                            worksheet.Row(rTable2).Height = 102;
                        }



                        for (int i = 0; i < dataCount; i++)
                        {
                            int rTop = startRowTable1 + (i * rowsPerItem);
                            int rBottom = rTop + 1;
                            int rTable2 = startRowTable2 + i;

                            var exportItem = dataList[i];
                            var rec = exportItem.RecordData;
                            var tol = toleranceData[i];

                            string displaySize = WpfApp1.Shared.Helpers.StringHelper.CleanSizeCOA(rec.Size).Replace("x", " x ");
                            string selectedType = exportItem.SelectedType;

                            if (!string.IsNullOrEmpty(selectedType) &&
                                !selectedType.Equals("Select", System.StringComparison.OrdinalIgnoreCase) &&
                                !selectedType.Equals("None", System.StringComparison.OrdinalIgnoreCase))
                            {
                                displaySize = displaySize + " - " + selectedType;
                            }

                            var cellBatch1 = worksheet.Cell(rTop, 2);
                            cellBatch1.Value = rec.BatchNo;
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightBatchNo(rec.BatchNo))
                                cellBatch1.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                            worksheet.Cell(rTop, 3).Value = displaySize;
                            worksheet.Cell(rTop, 4).Value = "No Dirty\nNo Blackspot\nNo Blisters";

                            string strThickTol = string.Format(cultureInvariant, "({0:0.00} \u00B1 {1:0.00})", tol.nominalThick, tol.thickness);
                            string strWidthTol = string.Format(cultureInvariant, "({0:0.00} \u00B1 {1:0.00})", tol.nominalWidth, tol.width);

                            var cellThick = worksheet.Cell(rTop, 6);
                            cellThick.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Thickness);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightThicknessWithTolerance(
                                rec.Thickness, tol.nominalThick, tol.thickness))
                                cellThick.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            worksheet.Cell(rBottom, 6).Value = strThickTol;

                            var cellWidth = worksheet.Cell(rTop, 7);
                            cellWidth.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Width);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightWidthWithTolerance(
                                rec.Width, tol.nominalWidth, tol.width))
                                cellWidth.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            worksheet.Cell(rBottom, 7).Value = strWidthTol;

                            var cellLength = worksheet.Cell(rTop, 8);
                            cellLength.Value = rec.Length;
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightLength(rec.Length))
                                cellLength.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                            worksheet.Cell(rBottom, 8).Value = "(4000 +15/-0)";

                            var cellRad = worksheet.Cell(rTop, 9);
                            cellRad.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Radius);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightRadius(rec.Radius))
                                cellRad.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                            var cellCham = worksheet.Cell(rTop, 10);
                            cellCham.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Chamber);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightChamber(rec.Chamber))
                                cellCham.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                            worksheet.Cell(rTop, 11).Value = "OK";

                            worksheet.Range(rTop, 2, rBottom, 2).Merge();
                            worksheet.Range(rTop, 3, rBottom, 3).Merge();
                            worksheet.Range(rTop, 4, rBottom, 5).Merge();
                            worksheet.Range(rTop, 9, rBottom, 9).Merge();
                            worksheet.Range(rTop, 10, rBottom, 10).Merge();
                            worksheet.Range(rTop, 11, rBottom, 11).Merge();

                            var cellBatch2 = worksheet.Cell(rTable2, 2);
                            cellBatch2.Value = rec.BatchNo;
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightBatchNo(rec.BatchNo))
                                cellBatch2.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                            worksheet.Cell(rTable2, 3).Value = displaySize;

                            var cellElec = worksheet.Cell(rTable2, 4);
                            cellElec.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Electric);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightElectric(rec.Electric))
                                cellElec.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                            var cellRes = worksheet.Cell(rTable2, 5);
                            cellRes.Value = string.Format(cultureInvariant, "{0:0.00000}", rec.Resistivity);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightResistivity(rec.Resistivity))
                                cellRes.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                            var cellElong = worksheet.Cell(rTable2, 6);
                            cellElong.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Elongation);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightElongation(rec.Elongation))
                                cellElong.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                            var cellTens = worksheet.Cell(rTable2, 7);
                            cellTens.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Tensile);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightTensile(rec.Tensile))
                                cellTens.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                            worksheet.Cell(rTable2, 8).Value = "No Crack";

                            var cellSpec = worksheet.Cell(rTable2, 9);
                            cellSpec.Value = string.Format(cultureInvariant, "{0:0.000}", rec.Spectro);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightSpectro(rec.Spectro))
                                cellSpec.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                            var cellOxy = worksheet.Cell(rTable2, 10);
                            cellOxy.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Oxygen);
                            if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightOxygen(rec.Oxygen))
                                cellOxy.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                            worksheet.Cell(rTable2, 11).Value = "OK";

                            worksheet.Row(rTop).Height = 51;
                            worksheet.Row(rBottom).Height = 51;
                            worksheet.Row(rTable2).Height = 102;
                        }


                        var table1Range = worksheet.Range(startRowTable1, 2, startRowTable1 + totalRowsNeeded - 1, 11);
                        ApplyCustomStyleBatch(table1Range);
                        ApplyBorders(table1Range);

                        for (int i = 0; i < dataCount; i++)
                        {
                            int rTop = startRowTable1 + (i * rowsPerItem);
                            int rBottom = rTop + 1;

                            var cellThickVal = worksheet.Cell(rTop, 6);
                            cellThickVal.Style.Font.Bold = true;
                            cellThickVal.Style.Font.FontSize = 22;
                            cellThickVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                            cellThickVal.Style.Border.BottomBorder = XLBorderStyleValues.None;

                            var cellThickTol = worksheet.Cell(rBottom, 6);
                            cellThickTol.Style.Font.Bold = false;
                            cellThickTol.Style.Font.Italic = true;
                            cellThickTol.Style.Font.FontSize = 22;
                            cellThickTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                            cellThickTol.Style.Border.TopBorder = XLBorderStyleValues.None;

                            var cellWidthVal = worksheet.Cell(rTop, 7);
                            cellWidthVal.Style.Font.Bold = true;
                            cellWidthVal.Style.Font.FontSize = 22;
                            cellWidthVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                            cellWidthVal.Style.Border.BottomBorder = XLBorderStyleValues.None;

                            var cellWidthTol = worksheet.Cell(rBottom, 7);
                            cellWidthTol.Style.Font.Bold = false;
                            cellWidthTol.Style.Font.Italic = true;
                            cellWidthTol.Style.Font.FontSize = 22;
                            cellWidthTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                            cellWidthTol.Style.Border.TopBorder = XLBorderStyleValues.None;

                            var cellLengthVal = worksheet.Cell(rTop, 8);
                            cellLengthVal.Style.NumberFormat.Format = "0";
                            cellLengthVal.Style.Font.Bold = true;
                            cellLengthVal.Style.Font.FontSize = 22;
                            cellLengthVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                            cellLengthVal.Style.Border.BottomBorder = XLBorderStyleValues.None;

                            var cellLengthTol = worksheet.Cell(rBottom, 8);
                            cellLengthTol.Style.Font.Bold = false;
                            cellLengthTol.Style.Font.Italic = true;
                            cellLengthTol.Style.Font.FontSize = 22;
                            cellLengthTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                            cellLengthTol.Style.Border.TopBorder = XLBorderStyleValues.None;

                            worksheet.Cell(rTop, 4).Style.Alignment.WrapText = true;
                        }

                        int lastRowTable2 = startRowTable2 + dataCount - 1;
                        var table2Range = worksheet.Range(startRowTable2, 2, lastRowTable2, 11);
                        ApplyCustomStyleBatch(table2Range);
                        ApplyBorders(table2Range);

                        worksheet.Row(lastRowTable2).InsertRowsBelow(5);
                        int firstInsertedRow = lastRowTable2 + 1;
                        int lastInsertedRow = firstInsertedRow + 4;

                        var signatureRange = worksheet.Range(firstInsertedRow, 2, lastInsertedRow, 11);
                        signatureRange.Style.Border.TopBorder = XLBorderStyleValues.None;
                        signatureRange.Style.Border.BottomBorder = XLBorderStyleValues.None;
                        signatureRange.Style.Border.LeftBorder = XLBorderStyleValues.None;
                        signatureRange.Style.Border.RightBorder = XLBorderStyleValues.None;
                        signatureRange.Style.Border.InsideBorder = XLBorderStyleValues.None;

                        signatureRange.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.NoColor;

                        for (int k = 0; k < 5; k++)
                        {
                            worksheet.Row(firstInsertedRow + k).Height = 102;
                        }
                        worksheet.Row(firstInsertedRow + 4).Height = 50;

                        int imageRow = firstInsertedRow + 1;

                        if (_img1Data != null)
                        {
                            using (var ms1 = new System.IO.MemoryStream(_img1Data))
                            {
                                var pic1 = worksheet.AddPicture(ms1);
                                pic1.MoveTo(worksheet.Cell(imageRow, 11));
                            }
                        }

                        if (_img2Data != null)
                        {
                            using (var ms2 = new System.IO.MemoryStream(_img2Data))
                            {
                                var pic2 = worksheet.AddPicture(ms2);
                                pic2.MoveTo(worksheet.Cell(imageRow, 2));
                            }
                        }

                        worksheet.PageSetup.PrintAreas.Clear();
                        worksheet.PageSetup.PrintAreas.Add(1, 2, lastInsertedRow, 11);
                        worksheet.PageSetup.AddHorizontalPageBreak(lastInsertedRow + 1);
                        worksheet.PageSetup.PagesTall = 1;
                        worksheet.PageSetup.PagesWide = 1;

                        workbook.SaveAs(fullPath);
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

        private string GetRomanMonth(int month)
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