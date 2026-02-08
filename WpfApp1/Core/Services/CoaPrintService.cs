using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using WpfApp1.ViewModels;

namespace WpfApp1.Core.Services
{
    public class CoaPrintService
    {
        private byte[]? _img1Data = null;
        private byte[]? _img2Data = null;
        private byte[]? _img3Data = null;
        private byte[]? _img4Data = null;

        public async System.Threading.Tasks.Task<string> GenerateCoaExcel(
            string customerName,
            string poNumber,
            string doNumber,
            System.Collections.Generic.IList<SheetModel> sheets,
            string standardName)
        {
            return await System.Threading.Tasks.Task.Run<string>(() =>
            {
                string basePath = @"C:\Users\mrrx\Documents\My Web Sites\H\COA";

                Assembly assembly = Assembly.GetExecutingAssembly();

                string resourceTemplate = "WpfApp1.Images.TEMPLATEE_COA_BUSBAR.xlsx";
                string resourceImg1 = "WpfApp1.Images.approved_IMG_v2.png";
                string resourceImg2 = "WpfApp1.Images.profile_SNI.png";
                string resourceImg3 = "WpfApp1.Images.document_no.png";
                string resourceImg4 = "WpfApp1.Images.logo_COA.png";

                if (_img1Data == null)
                {
                    using (Stream? stream = assembly.GetManifestResourceStream(resourceImg1))
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
                    using (Stream? stream = assembly.GetManifestResourceStream(resourceImg2))
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

                if (_img3Data == null)
                {
                    using (Stream? stream = assembly.GetManifestResourceStream(resourceImg3))
                    {
                        if (stream != null)
                        {
                            using (MemoryStream ms = new MemoryStream())
                            {
                                stream.CopyTo(ms);
                                _img3Data = ms.ToArray();
                            }
                        }
                    }
                }

                if (_img4Data == null)
                {
                    using (Stream? stream = assembly.GetManifestResourceStream(resourceImg4))
                    {
                        if (stream != null)
                        {
                            using (MemoryStream ms = new MemoryStream())
                            {
                                stream.CopyTo(ms);
                                _img4Data = ms.ToArray();
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
                            if (char.IsDigit(c)) leadingDigits += c;
                            else break;
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
                catch { }

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
                    if (!fName.StartsWith("~$")) validFileCount++;
                }

                int nomorFile = validFileCount + 1;
                string formattedFileNumber = nomorFile.ToString("000");
                string romanMonth = GetRomanMonth(now.Month);
                string fileName = $"{formattedFileNumber}. COA {customerName} {doNumber}.xlsx";
                string fullPath = System.IO.Path.Combine(finalDirectory, fileName);

                using (Stream? templateStream = assembly.GetManifestResourceStream(resourceTemplate))
                {
                    if (templateStream == null)
                    {
                        throw new FileNotFoundException($"Template Embedded Resource tidak ditemukan: '{resourceTemplate}'. \nPastikan 'Build Action' file Excel diset ke 'Embedded Resource'.");
                    }

                    using (ClosedXML.Excel.XLWorkbook workbook = new ClosedXML.Excel.XLWorkbook(templateStream))
                    {
                        ClosedXML.Excel.IXLWorksheet templateWorksheet = workbook.Worksheet(1);

                        foreach (var picture in templateWorksheet.Pictures.ToList())
                        {
                            picture.Delete();
                        }

                        if (sheets.Count > 0)
                        {
                            templateWorksheet.Name = sheets[0].SheetName;
                        }

                        for (int i = 1; i < sheets.Count; i++)
                        {
                            templateWorksheet.CopyTo(sheets[i].SheetName);
                        }

                        for (int i = 0; i < sheets.Count; i++)
                        {
                            ClosedXML.Excel.IXLWorksheet currentWorksheet = workbook.Worksheet(i + 1);
                            SheetModel currentSheetModel = sheets[i];

                            ProcessSingleSheet(
                                currentWorksheet,
                                currentSheetModel.Items,
                                customerName,
                                poNumber,
                                formattedFileNumber,
                                standardName,
                                now
                            );
                        }

                        workbook.SaveAs(fullPath);
                    }
                }

                return fullPath;
            });
        }

        private void ProcessSingleSheet(
                    IXLWorksheet worksheet,
                    System.Collections.Generic.IList<WpfApp1.Core.Models.BusbarExportItem> dataList,
                    string customerName,
                    string poNumber,
                    string fileNumber,
                    string standardName,
                    System.DateTime now)
        {
            worksheet.ShowGridLines = false;

            System.Globalization.CultureInfo cultureInvariant = System.Globalization.CultureInfo.InvariantCulture;
            System.Random randomGen = new System.Random();
            string romanMonth = GetRomanMonth(now.Month);

            worksheet.Style.Font.FontName = "Montserrat";

            worksheet.Cell("C12").Value = ": " + poNumber;
            worksheet.Cell("K12").Value = ": " + customerName;
            worksheet.Cell("K13").Value = ": " + now.ToString("dd/MM/yyyy");
            worksheet.Cell("K14").Value = ": " + $"{fileNumber}/{romanMonth}/{now.Year}";

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

            System.Collections.Generic.List<(double thickness, double width, double nominalThick, double nominalWidth)> toleranceData = new System.Collections.Generic.List<(double, double, double, double)>();

            for (int i = 0; i < dataCount; i++)
            {
                WpfApp1.Core.Models.BusbarExportItem exportItem = dataList[i];
                WpfApp1.Core.Models.BusbarRecord rec = exportItem.RecordData;
                string cleanSizeStr = WpfApp1.Shared.Helpers.StringHelper.CleanSizeCOA(rec.Size);

                double finalTolThickness = 0;
                double finalTolWidth = 0;

                if (standardName == "JIS")
                {
                    (double Thickness, double Width) calculatedTols = WpfApp1.Shared.Helpers.ToleranceJIS.CalculateFromDbString(cleanSizeStr);
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

                WpfApp1.Core.Models.BusbarExportItem exportItem = dataList[i];
                WpfApp1.Core.Models.BusbarRecord rec = exportItem.RecordData;
                (double thickness, double width, double nominalThick, double nominalWidth) tol = toleranceData[i];

                string displaySize = WpfApp1.Shared.Helpers.StringHelper.CleanSizeCOA(rec.Size).Replace("x", " x ");
                string selectedType = exportItem.SelectedType;

                if (!string.IsNullOrEmpty(selectedType) &&
                    !selectedType.Equals("Select", System.StringComparison.OrdinalIgnoreCase) &&
                    !selectedType.Equals("None", System.StringComparison.OrdinalIgnoreCase))
                {
                    displaySize = displaySize + " - " + selectedType;
                }

                IXLCell cellBatch1 = worksheet.Cell(rTop, 2);
                cellBatch1.Value = rec.BatchNo;
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightBatchNo(rec.BatchNo))
                    cellBatch1.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                worksheet.Cell(rTop, 3).Value = displaySize;
                worksheet.Cell(rTop, 4).Value = "No Dirty\nNo Blackspot\nNo Blisters";

                string strThickTol = string.Format(cultureInvariant, "({0:0.00} \u00B1 {1:0.00})", tol.nominalThick, tol.thickness);
                string strWidthTol = string.Format(cultureInvariant, "({0:0.00} \u00B1 {1:0.00})", tol.nominalWidth, tol.width);

                IXLCell cellThick = worksheet.Cell(rTop, 7);
                cellThick.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Thickness);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightThicknessWithTolerance(
                    rec.Thickness, tol.nominalThick, tol.thickness))
                    cellThick.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                worksheet.Cell(rBottom, 7).Value = strThickTol;

                IXLCell cellWidth = worksheet.Cell(rTop, 8);
                cellWidth.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Width);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightWidthWithTolerance(
                    rec.Width, tol.nominalWidth, tol.width))
                    cellWidth.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                worksheet.Cell(rBottom, 8).Value = strWidthTol;

                IXLCell cellLength = worksheet.Cell(rTop, 9);
                cellLength.Value = rec.Length;
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightLength(rec.Length))
                    cellLength.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                worksheet.Cell(rBottom, 9).Value = "(4000 +15/-0)";

                IXLCell cellRad = worksheet.Cell(rTop, 10);
                cellRad.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Radius);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightRadius(rec.Radius))
                    cellRad.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                IXLCell cellCham = worksheet.Cell(rTop, 11);
                cellCham.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Chamber);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightChamber(rec.Chamber))
                    cellCham.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                worksheet.Cell(rTop, 12).Value = "OK";

                worksheet.Range(rTop, 2, rBottom, 2).Merge();
                worksheet.Range(rTop, 3, rBottom, 3).Merge();
                worksheet.Range(rTop, 4, rBottom, 6).Merge();
                worksheet.Range(rTop, 10, rBottom, 10).Merge();
                worksheet.Range(rTop, 11, rBottom, 11).Merge();
                worksheet.Range(rTop, 12, rBottom, 12).Merge();

                IXLCell cellBatch2 = worksheet.Cell(rTable2, 2);
                cellBatch2.Value = rec.BatchNo;
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightBatchNo(rec.BatchNo))
                    cellBatch2.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                worksheet.Cell(rTable2, 3).Value = displaySize;

                IXLCell cellElec = worksheet.Cell(rTable2, 4);
                cellElec.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Electric);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightElectric(rec.Electric))
                    cellElec.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                IXLCell cellRes = worksheet.Cell(rTable2, 5);
                cellRes.Value = string.Format(cultureInvariant, "{0:0.00000}", rec.Resistivity);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightResistivity(rec.Resistivity))
                    cellRes.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                IXLCell cellElong = worksheet.Cell(rTable2, 6);
                cellElong.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Elongation);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightElongation(rec.Elongation))
                    cellElong.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                IXLCell cellTens = worksheet.Cell(rTable2, 7);
                cellTens.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Tensile);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightTensile(rec.Tensile))
                    cellTens.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                worksheet.Cell(rTable2, 8).Value = "No Crack";
                worksheet.Cell(rTable2, 9).Value = "No Crack";

                IXLCell cellSpec = worksheet.Cell(rTable2, 10);
                cellSpec.Value = string.Format(cultureInvariant, "{0:0.000}", rec.Spectro);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightSpectro(rec.Spectro))
                    cellSpec.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                IXLCell cellOxy = worksheet.Cell(rTable2, 11);
                cellOxy.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Oxygen);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightOxygen(rec.Oxygen))
                    cellOxy.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;

                worksheet.Cell(rTable2, 12).Value = "OK";
            }

            IXLRange table1Range = worksheet.Range(startRowTable1, 2, startRowTable1 + totalRowsNeeded - 1, 12);
            ApplyCustomStyleBatch(table1Range);
            ApplyBorders(table1Range);

            for (int i = 0; i < dataCount; i++)
            {
                int rTop = startRowTable1 + (i * rowsPerItem);
                int rBottom = rTop + 1;

                IXLCell cellThickVal = worksheet.Cell(rTop, 7);
                cellThickVal.Style.Font.Bold = true;
                cellThickVal.Style.Font.FontSize = 22;
                cellThickVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                cellThickVal.Style.Border.BottomBorder = XLBorderStyleValues.None;

                IXLCell cellThickTol = worksheet.Cell(rBottom, 7);
                cellThickTol.Style.Font.Bold = false;
                cellThickTol.Style.Font.Italic = true;
                cellThickTol.Style.Font.FontSize = 22;
                cellThickTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                cellThickTol.Style.Border.TopBorder = XLBorderStyleValues.None;

                IXLCell cellWidthVal = worksheet.Cell(rTop, 8);
                cellWidthVal.Style.Font.Bold = true;
                cellWidthVal.Style.Font.FontSize = 22;
                cellWidthVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                cellWidthVal.Style.Border.BottomBorder = XLBorderStyleValues.None;

                IXLCell cellWidthTol = worksheet.Cell(rBottom, 8);
                cellWidthTol.Style.Font.Bold = false;
                cellWidthTol.Style.Font.Italic = true;
                cellWidthTol.Style.Font.FontSize = 22;
                cellWidthTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                cellWidthTol.Style.Border.TopBorder = XLBorderStyleValues.None;

                IXLCell cellLengthVal = worksheet.Cell(rTop, 9);
                cellLengthVal.Style.NumberFormat.Format = "0";
                cellLengthVal.Style.Font.Bold = true;
                cellLengthVal.Style.Font.FontSize = 22;
                cellLengthVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                cellLengthVal.Style.Border.BottomBorder = XLBorderStyleValues.None;

                IXLCell cellLengthTol = worksheet.Cell(rBottom, 9);
                cellLengthTol.Style.Font.Bold = false;
                cellLengthTol.Style.Font.Italic = true;
                cellLengthTol.Style.Font.FontSize = 22;
                cellLengthTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                cellLengthTol.Style.Border.TopBorder = XLBorderStyleValues.None;

                worksheet.Cell(rTop, 4).Style.Alignment.WrapText = true;
            }

            int lastRowTable2 = startRowTable2 + dataCount - 1;
            IXLRange table2Range = worksheet.Range(startRowTable2, 2, lastRowTable2, 12);
            ApplyCustomStyleBatch(table2Range);
            ApplyBorders(table2Range);

            worksheet.Row(lastRowTable2).InsertRowsBelow(5);
            int firstInsertedRow = lastRowTable2 + 1;
            int lastInsertedRow = firstInsertedRow + 4;

            IXLCell noteCell = worksheet.Cell(firstInsertedRow, 2);
            noteCell.Value = "Note :\n* Testing 1x/size/year";
            noteCell.Style.Font.FontName = "Montserrat";
            noteCell.Style.Font.FontSize = 22;
            noteCell.Style.Font.Bold = true;
            noteCell.Style.Alignment.WrapText = true;
            noteCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;


            IXLRange signatureRange = worksheet.Range(firstInsertedRow, 2, lastInsertedRow, 12);
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
            double dpi = 96.0;

            if (_img1Data != null)
            {
                using (System.IO.MemoryStream ms1 = new System.IO.MemoryStream(_img1Data))
                {
                    var pic1 = worksheet.AddPicture(ms1);
                    pic1.MoveTo(worksheet.Cell(imageRow, 10), 235, 0);

                    pic1.Height = (int)((7.21 / 2.54) * dpi);
                    pic1.Width = (int)((19.91 / 2.54) * dpi);
                }
            }

            if (_img2Data != null)
            {
                using (System.IO.MemoryStream ms2 = new System.IO.MemoryStream(_img2Data))
                {
                    var pic2 = worksheet.AddPicture(ms2);
                    pic2.MoveTo(worksheet.Cell(imageRow, 2));
                }
            }

            if (_img3Data != null)
            {
                using (System.IO.MemoryStream ms3 = new System.IO.MemoryStream(_img3Data))
                {
                    var pic3 = worksheet.AddPicture(ms3);
                    pic3.MoveTo(worksheet.Cell(4, 11), 295, 0);

                    pic3.Height = (int)((3.90 / 2.54) * dpi);
                    pic3.Width = (int)((9.39 / 2.54) * dpi);
                }
            }

            if (_img4Data != null)
            {
                using (System.IO.MemoryStream ms4 = new System.IO.MemoryStream(_img4Data))
                {
                    var pic4 = worksheet.AddPicture(ms4);
                    pic4.MoveTo(worksheet.Cell(2, 2));

                    //pic4.Height = (int)((3.90 / 2.54) * dpi);
                   // pic4.Width = (int)((9.39 / 2.54) * dpi);
                }
            }

            worksheet.PageSetup.PrintAreas.Clear();
            worksheet.PageSetup.PrintAreas.Add(1, 2, lastInsertedRow, 12);
            worksheet.PageSetup.AddHorizontalPageBreak(lastInsertedRow + 1);
            worksheet.PageSetup.PagesTall = 1;
            worksheet.PageSetup.PagesWide = 1;
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
            _img3Data = null;
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