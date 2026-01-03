using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace WpfApp1.Core.Services
{
    public class CoaPrintService
    {
        private bool? _img1Exists = null;
        private bool? _img2Exists = null;

        // PERUBAHAN 1: Tanda tangan metode diganti menjadi async Task<string>
        public async System.Threading.Tasks.Task<string> GenerateCoaExcel(
            string customerName,
            string poNumber,
            string doNumber,
            System.Collections.Generic.List<WpfApp1.Core.Models.BusbarExportItem> dataList,
            string standardName)
        {
            // PERUBAHAN 2: Seluruh logika berat dibungkus Task.Run agar non-blocking
            return await System.Threading.Tasks.Task.Run<string>(() =>
            {
                string templatePath = @"C:\Users\mrrx\Documents\My Web Sites\H\TEMPLATE_COA_BUSBAR.xlsx";
                string basePath = @"C:\Users\mrrx\Documents\My Web Sites\H\COA";
                string pathImg1 = @"C:\Users\mrrx\Documents\Custom Office Templates\WpfApp1\WpfApp1\Images\approved_IMG.png";
                string pathImg2 = @"C:\Users\mrrx\Documents\Custom Office Templates\WpfApp1\WpfApp1\Images\profile_SNI.png";

                if (!System.IO.File.Exists(templatePath))
                {
                    throw new System.IO.FileNotFoundException($"File template tidak ditemukan: {templatePath}");
                }

                System.DateTime now = System.DateTime.Now;
                string yearFolder = now.ToString("yyyy");

                var cultureIndo = new System.Globalization.CultureInfo("id-ID");
                string monthName = cultureIndo.DateTimeFormat.GetMonthName(now.Month);
                monthName = cultureIndo.TextInfo.ToTitleCase(monthName);

                string monthFolder = $"{now.Month}. {monthName}";
                string finalDirectory = System.IO.Path.Combine(basePath, yearFolder, monthFolder);

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

                using (var workbook = new ClosedXML.Excel.XLWorkbook(templatePath))
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

                    if (totalRowsNeeded > defaultRowsAvailable)
                    {
                        int rowsToInsert = totalRowsNeeded - defaultRowsAvailable;
                        worksheet.Row(22).InsertRowsBelow(rowsToInsert);
                        startRowTable2 = originalStartRowTable2 + rowsToInsert;
                    }

                    var toleranceData = new System.Collections.Generic.List<(double thickness, double width, double nominalThick, double nominalWidth)>();

                    for (int i = 0; i < dataCount; i++)
                    {
                        var rec = dataList[i].RecordData;
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

                        var rec = dataList[i].RecordData;
                        var tol = toleranceData[i];

                        worksheet.Cell(rTop, 2).Value = rec.BatchNo;
                        worksheet.Cell(rTop, 3).Value = rec.Size;
                        worksheet.Cell(rTop, 4).Value = "No Dirty\nNo Blackspot\nNo Blisters";

                        string strThickTol = string.Format(cultureInvariant, "({0:0.00} \u00B1 {1:0.00})", tol.nominalThick, tol.thickness);
                        string strWidthTol = string.Format(cultureInvariant, "({0:0.00} \u00B1 {1:0.00})", tol.nominalWidth, tol.width);

                        worksheet.Cell(rTop, 6).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Thickness);
                        worksheet.Cell(rBottom, 6).Value = strThickTol;

                        worksheet.Cell(rTop, 7).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Width);
                        worksheet.Cell(rBottom, 7).Value = strWidthTol;

                        worksheet.Cell(rTop, 8).Value = rec.Length;
                        worksheet.Cell(rBottom, 8).Value = "(4000 ± 15)";

                        worksheet.Cell(rTop, 9).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Radius);
                        worksheet.Cell(rTop, 10).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Chamber);
                        worksheet.Cell(rTop, 11).Value = "OK";

                        worksheet.Range(rTop, 2, rBottom, 2).Merge();
                        worksheet.Range(rTop, 3, rBottom, 3).Merge();
                        worksheet.Range(rTop, 4, rBottom, 5).Merge();
                        worksheet.Range(rTop, 9, rBottom, 9).Merge();
                        worksheet.Range(rTop, 10, rBottom, 10).Merge();
                        worksheet.Range(rTop, 11, rBottom, 11).Merge();

                        worksheet.Cell(rTable2, 2).Value = rec.BatchNo;
                        worksheet.Cell(rTable2, 3).Value = rec.Size;
                        worksheet.Cell(rTable2, 4).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Electric);
                        worksheet.Cell(rTable2, 5).Value = string.Format(cultureInvariant, "{0:0.00000}", rec.Resistivity);
                        worksheet.Cell(rTable2, 6).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Elongation);
                        worksheet.Cell(rTable2, 7).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Tensile);
                        worksheet.Cell(rTable2, 8).Value = "No Crack";
                        worksheet.Cell(rTable2, 9).Value = string.Format(cultureInvariant, "{0:0.000}", rec.Spectro);
                        worksheet.Cell(rTable2, 10).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Oxygen);
                        worksheet.Cell(rTable2, 11).Value = "OK";
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

                    worksheet.Row(lastRowTable2).InsertRowsBelow(6);

                    int firstInsertedRow = lastRowTable2 + 1;
                    int lastInsertedRow = firstInsertedRow + 5;

                    var signatureRange = worksheet.Range(firstInsertedRow, 2, lastInsertedRow, 11);
                    signatureRange.Style.Border.TopBorder = XLBorderStyleValues.None;
                    signatureRange.Style.Border.BottomBorder = XLBorderStyleValues.None;
                    signatureRange.Style.Border.LeftBorder = XLBorderStyleValues.None;
                    signatureRange.Style.Border.RightBorder = XLBorderStyleValues.None;
                    signatureRange.Style.Border.InsideBorder = XLBorderStyleValues.None;

                    for (int k = 0; k < 5; k++)
                    {
                        worksheet.Row(firstInsertedRow + k).Height = 102;
                    }
                    worksheet.Row(firstInsertedRow + 5).Height = 50;

                    int imageRow = firstInsertedRow + 2;

                    if (!_img1Exists.HasValue)
                        _img1Exists = System.IO.File.Exists(pathImg1);

                    if (!_img2Exists.HasValue)
                        _img2Exists = System.IO.File.Exists(pathImg2);

                    if (_img1Exists.Value)
                    {
                        var pic1 = worksheet.AddPicture(pathImg1);
                        pic1.MoveTo(worksheet.Cell(imageRow, 11));
                    }

                    if (_img2Exists.Value)
                    {
                        var pic2 = worksheet.AddPicture(pathImg2);
                        pic2.MoveTo(worksheet.Cell(imageRow, 2));
                    }

                    worksheet.PageSetup.PrintAreas.Clear();
                    worksheet.PageSetup.PrintAreas.Add(1, 2, lastInsertedRow, 11);
                    worksheet.PageSetup.AddHorizontalPageBreak(lastInsertedRow + 1);
                    worksheet.PageSetup.PagesTall = 1;
                    worksheet.PageSetup.PagesWide = 1;

                    workbook.SaveAs(fullPath);
                }

                string pdfPath = System.IO.Path.ChangeExtension(fullPath, ".pdf");
                ConvertExcelToPdf(fullPath, pdfPath);

                return fullPath;
            });
        }

        private void ConvertExcelToPdf(string excelFile, string pdfFile)
        {
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();

            try
            {
                workbook.LoadFromFile(excelFile);

                foreach (Spire.Xls.Worksheet sheet in workbook.Worksheets)
                {
                    Spire.Xls.PageSetup setup = sheet.PageSetup;
                    setup.FitToPagesWide = 1;
                    setup.FitToPagesTall = 1;
                }

                workbook.SaveToFile(pdfFile, Spire.Xls.FileFormat.PDF);
            }
            catch (System.Exception ex)
            {
                throw new System.Exception("Gagal konversi PDF dengan Spire: " + ex.Message, ex);
            }
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
            _img1Exists = null;
            _img2Exists = null;
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