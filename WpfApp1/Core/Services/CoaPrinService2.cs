
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using WpfApp1.ViewModels;
using WpfApp1.Shared.Helpers; 

namespace WpfApp1.Core.Services
{
    public class CoaPrintService2
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
                if (sheets == null || sheets.Count == 0)
                {
                    throw new ArgumentException("Tidak ada sheet yang diberikan untuk generate COA.", nameof(sheets));
                }

                bool hasAnyItems = false;
                foreach (SheetModel sheet in sheets)
                {
                    if (sheet != null && sheet.Items != null && sheet.Items.Count > 0)
                    {
                        hasAnyItems = true;
                        break;
                    }
                }

                if (!hasAnyItems)
                {
                    throw new InvalidOperationException("Semua sheet kosong. Tidak ada data untuk di-export.");
                }

                string basePath = @"C:\Users\mrrx\Documents\My Web Sites\H\COA";

                Assembly assembly = Assembly.GetExecutingAssembly();

                string resourceTemplate = "WpfApp1.Shared.Images.TEMPLATEE_COA_SIMEN.xlsx";
                string resourceImg1 = "WpfApp1.Shared.Images.approved_IMG_v2.png";
                string resourceImg2 = "WpfApp1.Shared.Images.profile_SNI.png";
                string resourceImg3 = "WpfApp1.Shared.Images.document_no_simens_comp.png";
                string resourceImg4 = "WpfApp1.Shared.Images.logo_COA-comp.png";

                if (_img1Data == null)
                {
                    using (Stream? stream = assembly.GetManifestResourceStream(resourceImg1))
                    {
                        if (stream != null)
                        {
                            using (MemoryStream ms = new MemoryStream())
                            {
                                stream.CopyTo(ms);
                                _img1Data = ImageHelper.CompressImage(ms.ToArray(), 19.91, 7.21);
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
                                _img2Data = ImageHelper.CompressImage(ms.ToArray(), 22.17, 11.48);
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
                                _img3Data = ImageHelper.CompressImage(ms.ToArray(), 9.37, 3.89);
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
                                _img4Data = ImageHelper.CompressImage(ms.ToArray(), 9.55, 3.68);
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

                string safeDoNumber = doNumber.Replace(":", "-").Replace("/", "-").Replace("\\", "-");
                string fileName = $"{formattedFileNumber}. COA - PT. SIEMENS INDONESIA {safeDoNumber}.xlsx";
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
                                poNumber,
                                formattedFileNumber,
                                standardName,
                                now
                            );
                        }

                        workbook.SaveAs(fullPath);
                    }
                }

                FileInfo fileInfo = new FileInfo(fullPath);
                if (!fileInfo.Exists || fileInfo.Length == 0)
                {
                    throw new IOException($"File tidak berhasil dibuat atau kosong: {fullPath}");
                }

                return fullPath;
            });
        }

        private void ProcessSingleSheet(
            IXLWorksheet worksheet,
            System.Collections.Generic.IList<WpfApp1.Core.Models.BusbarExportItem> dataList,
            string poNumber,
            string fileNumber,
            string standardName,
            System.DateTime now)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (dataList == null || dataList.Count == 0) return;

            worksheet.ShowGridLines = false;

            System.Globalization.CultureInfo cultureInvariant = System.Globalization.CultureInfo.InvariantCulture;
            System.Random randomGen = new System.Random();
            string romanMonth = GetRomanMonth(now.Month);

            worksheet.Style.Font.FontName = "Montserrat";

            worksheet.Cell("C14").Value = ": " + poNumber;
            worksheet.Cell("M15").Value = ": " + now.ToString("dd/MM/yyyy");
            worksheet.Cell("M16").Value = ": " + $"{fileNumber}/{romanMonth}/{now.Year}";

            int dataCount = dataList.Count;
            int rowsPerItem = 2;
            int totalRowsNeeded = dataCount * rowsPerItem;
            int defaultRowsAvailable = 3;

            int startRowTable1 = 22;

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

                worksheet.Row(rTop).Height = 47;
                worksheet.Row(rBottom).Height = 47;
            }

            for (int i = 0; i < dataCount; i++)
            {
                int rTop = startRowTable1 + (i * rowsPerItem);
                int rBottom = rTop + 1;

                WpfApp1.Core.Models.BusbarExportItem exportItem = dataList[i];
                WpfApp1.Core.Models.BusbarRecord rec = exportItem.RecordData;
                var tol = toleranceData[i];

                string displaySize = WpfApp1.Shared.Helpers.StringHelper.CleanSizeCOA(rec.Size).Replace("x", " x ");
                string selectedType = exportItem.SelectedType;

                if (!string.IsNullOrEmpty(selectedType) &&
                    !selectedType.Equals("Select", System.StringComparison.OrdinalIgnoreCase) &&
                    !selectedType.Equals("None", System.StringComparison.OrdinalIgnoreCase))
                {
                    displaySize = displaySize + " - " + selectedType;
                }

                IXLRange itemRange = worksheet.Range(rTop, 1, rBottom, 15);
                itemRange.Style.Font.FontSize = 16;

                IXLCell cellBatch = worksheet.Cell(rTop, 1);
                cellBatch.Value = rec.BatchNo;
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightBatchNo(rec.BatchNo))
                {
                    cellBatch.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                }
                worksheet.Range(rTop, 1, rBottom, 1).Merge();

                worksheet.Cell(rTop, 2).Value = displaySize;
                worksheet.Range(rTop, 2, rBottom, 2).Merge();

                IXLCell cellVisual = worksheet.Cell(rTop, 3);
                cellVisual.Value = "No Crack\nNo Pores\nNo Blisters\nNo Inclusion";
                worksheet.Range(rTop, 3, rBottom, 3).Merge();

                string strThickTol = string.Format(cultureInvariant, "({0:0.00} \u00B1 {1:0.00})", tol.nominalThick, tol.thickness);
                string strWidthTol = string.Format(cultureInvariant, "({0:0.00} \u00B1 {1:0.00})", tol.nominalWidth, tol.width);

                IXLCell cellThick = worksheet.Cell(rTop, 4);
                cellThick.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Thickness);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightThicknessWithTolerance(rec.Thickness, tol.nominalThick, tol.thickness))
                {
                    cellThick.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                }
                worksheet.Cell(rBottom, 4).Value = strThickTol;

                IXLCell cellWidth = worksheet.Cell(rTop, 5);
                cellWidth.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Width);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightWidthWithTolerance(rec.Width, tol.nominalWidth, tol.width))
                {
                    cellWidth.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                }
                worksheet.Cell(rBottom, 5).Value = strWidthTol;

                IXLCell cellLength = worksheet.Cell(rTop, 6);
                cellLength.Value = rec.Length;
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightLength(rec.Length))
                {
                    cellLength.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                }
                worksheet.Cell(rBottom, 6).Value = "(4000 +15/-0)";

                IXLCell cellElec = worksheet.Cell(rTop, 7);
                cellElec.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Electric);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightElectric(rec.Electric))
                {
                    cellElec.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                }
                worksheet.Range(rTop, 7, rBottom, 7).Merge();

                IXLCell cellRes = worksheet.Cell(rTop, 8);
                cellRes.Value = string.Format(cultureInvariant, "{0:0.00000}", rec.Resistivity);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightResistivity(rec.Resistivity))
                {
                    cellRes.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                }
                worksheet.Range(rTop, 8, rBottom, 8).Merge();

                worksheet.Cell(rTop, 9).Value = "OK";
                worksheet.Range(rTop, 9, rBottom, 9).Merge();

                IXLCell cellHard = worksheet.Cell(rTop, 10);
                cellHard.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Hardness);
                worksheet.Range(rTop, 10, rBottom, 10).Merge();

                IXLCell cellSpectro = worksheet.Cell(rTop, 11);
                if (rec.Spectro > 0)
                {
                    string spectroVal = rec.Spectro.ToString("0.000", cultureInvariant);
                    cellSpectro.Value = "'" + spectroVal;
                }
                worksheet.Range(rTop, 11, rBottom, 11).Merge();

                worksheet.Range(rTop, 12, rBottom, 12).Merge();
                worksheet.Range(rTop, 13, rBottom, 13).Merge();

                IXLCell cellOxy = worksheet.Cell(rTop, 14);
                cellOxy.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Oxygen);
                if (WpfApp1.Shared.Helpers.CellValidationHelper.ShouldHighlightOxygen(rec.Oxygen))
                {
                    cellOxy.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Red;
                }
                worksheet.Range(rTop, 14, rBottom, 14).Merge();

                worksheet.Cell(rTop, 15).Value = "OK";
                worksheet.Range(rTop, 15, rBottom, 15).Merge();
            }

            int lastDataRow = startRowTable1 + totalRowsNeeded - 1;
            IXLRange tableRange = worksheet.Range(startRowTable1, 1, lastDataRow, 15);

            ApplyCustomStyleBatch(tableRange);
            ApplyBorders(tableRange);

            worksheet.Range(startRowTable1, 3, lastDataRow, 3).Style.Font.FontSize = 14;

            for (int i = 0; i < dataCount; i++)
            {
                int rTop = startRowTable1 + (i * rowsPerItem);
                int rBottom = rTop + 1;

                IXLCell cellThickVal = worksheet.Cell(rTop, 4);
                cellThickVal.Style.Font.Bold = true;
                cellThickVal.Style.Font.FontSize = 16;
                cellThickVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                cellThickVal.Style.Border.BottomBorder = XLBorderStyleValues.None;

                IXLCell cellThickTol = worksheet.Cell(rBottom, 4);
                cellThickTol.Style.Font.Bold = false;
                cellThickTol.Style.Font.Italic = true;
                cellThickTol.Style.Font.FontSize = 16;
                cellThickTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                cellThickTol.Style.Border.TopBorder = XLBorderStyleValues.None;

                IXLCell cellWidthVal = worksheet.Cell(rTop, 5);
                cellWidthVal.Style.Font.Bold = true;
                cellWidthVal.Style.Font.FontSize = 16;
                cellWidthVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                cellWidthVal.Style.Border.BottomBorder = XLBorderStyleValues.None;

                IXLCell cellWidthTol = worksheet.Cell(rBottom, 5);
                cellWidthTol.Style.Font.Bold = false;
                cellWidthTol.Style.Font.Italic = true;
                cellWidthTol.Style.Font.FontSize = 16;
                cellWidthTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                cellWidthTol.Style.Border.TopBorder = XLBorderStyleValues.None;

                IXLCell cellLengthVal = worksheet.Cell(rTop, 6);
                cellLengthVal.Style.NumberFormat.Format = "0";
                cellLengthVal.Style.Font.Bold = true;
                cellLengthVal.Style.Font.FontSize = 16;
                cellLengthVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                cellLengthVal.Style.Border.BottomBorder = XLBorderStyleValues.None;

                IXLCell cellLengthTol = worksheet.Cell(rBottom, 6);
                cellLengthTol.Style.Font.Bold = false;
                cellLengthTol.Style.Font.Italic = true;
                cellLengthTol.Style.Font.FontSize = 16;
                cellLengthTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                cellLengthTol.Style.Border.TopBorder = XLBorderStyleValues.None;

                worksheet.Cell(rTop, 3).Style.Alignment.WrapText = true;
            }

            worksheet.Row(lastDataRow).InsertRowsBelow(4);
            int firstInsertedRow = lastDataRow + 1;
            int lastInsertedRow = firstInsertedRow + 3;

            worksheet.Row(firstInsertedRow).Height = 94;
            worksheet.Row(firstInsertedRow + 1).Height = 94;
            worksheet.Row(firstInsertedRow + 2).Height = 94;
            worksheet.Row(firstInsertedRow + 3).Height = 65;

            IXLRange signatureRange = worksheet.Range(firstInsertedRow, 1, lastInsertedRow, 15);
            signatureRange.Style.Border.TopBorder = XLBorderStyleValues.None;
            signatureRange.Style.Border.BottomBorder = XLBorderStyleValues.None;
            signatureRange.Style.Border.LeftBorder = XLBorderStyleValues.None;
            signatureRange.Style.Border.RightBorder = XLBorderStyleValues.None;
            signatureRange.Style.Border.InsideBorder = XLBorderStyleValues.None;
            signatureRange.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.NoColor;

            int imageRow = firstInsertedRow + 1;
            double dpi = 96.0;

            if (_img1Data != null)
            {
                using (MemoryStream ms1 = new MemoryStream(_img1Data))
                {
                    var pic1 = worksheet.AddPicture(ms1);
                    pic1.MoveTo(worksheet.Cell(imageRow, 12), 90, 20);
                    pic1.Height = (int)((5.19 / 2.54) * dpi);
                    pic1.Width = (int)((14.29 / 2.54) * dpi);
                }
            }

            if (_img2Data != null)
            {
                using (MemoryStream ms2 = new MemoryStream(_img2Data))
                {
                    var pic2 = worksheet.AddPicture(ms2);
                    pic2.MoveTo(worksheet.Cell(imageRow, 1));
                    pic2.Height = (int)((8.45 / 2.54) * dpi);
                    pic2.Width = (int)((16.31 / 2.54) * dpi);
                }
            }

            if (_img3Data != null)
            {
                using (MemoryStream ms3 = new MemoryStream(_img3Data))
                {
                    var pic3 = worksheet.AddPicture(ms3);
                    pic3.MoveTo(worksheet.Cell(2, 14), 60, 0);
                    pic3.Height = (int)((2.93 / 2.54) * dpi);
                    pic3.Width = (int)((6.69 / 2.54) * dpi);
                }
            }

            if (_img4Data != null)
            {
                using (MemoryStream ms4 = new MemoryStream(_img4Data))
                {
                    var pic4 = worksheet.AddPicture(ms4);
                    pic4.MoveTo(worksheet.Cell(2, 1));
                    pic4.Height = (int)((3.13 / 2.54) * dpi);
                    pic4.Width = (int)((8.08 / 2.54) * dpi);
                }
            }

            worksheet.PageSetup.PrintAreas.Clear();
            worksheet.PageSetup.PrintAreas.Add(1, 1, lastInsertedRow, 15);
            worksheet.PageSetup.AddHorizontalPageBreak(lastInsertedRow + 1);
            worksheet.PageSetup.PagesTall = 1;
            worksheet.PageSetup.PagesWide = 1;
        }

        private void ApplyCustomStyleBatch(ClosedXML.Excel.IXLRange range)
        {
            range.Style.Font.Bold = true;
            range.Style.Font.FontSize = 16;
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
            _img4Data = null;
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