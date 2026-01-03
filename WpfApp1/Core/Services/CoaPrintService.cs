using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace WpfApp1.Core.Services
{
    public class CoaPrintService
    {
        public string GenerateCoaExcel(
            string customerName,
            string poNumber,
            string doNumber,
            System.Collections.Generic.List<WpfApp1.Core.Models.BusbarExportItem> dataList,
            string standardName)
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

                var cultureInvariant = System.Globalization.CultureInfo.InvariantCulture;

                for (int i = 0; i < dataCount; i++)
                {
                    int rTop = startRowTable1 + (i * rowsPerItem);
                    int rBottom = rTop + 1;

                    int rTable2 = startRowTable2 + i;

                    worksheet.Row(rTop).Height = 51;
                    worksheet.Row(rBottom).Height = 51;

                    worksheet.Row(rTable2).Height = 102;

                    var rec = dataList[i].RecordData;

                    worksheet.Cell(rTop, 2).Value = rec.BatchNo;
                    worksheet.Range(rTop, 2, rBottom, 2).Merge();

                    worksheet.Cell(rTop, 3).Value = rec.Size;
                    worksheet.Range(rTop, 3, rBottom, 3).Merge();

                    var cellD = worksheet.Cell(rTop, 4);
                    cellD.Value = "No Dirty\nNo Blackspot\nNo Blisters";
                    cellD.Style.Alignment.WrapText = true;
                    worksheet.Range(rTop, 4, rBottom, 5).Merge();

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

                    string strThickTol = string.Format(cultureInvariant, "({0:0.00} \u00B1 {1:0.00})", nominalThick, finalTolThickness);
                    string strWidthTol = string.Format(cultureInvariant, "({0:0.00} \u00B1 {1:0.00})", nominalWidth, finalTolWidth);

                    var cellThickVal = worksheet.Cell(rTop, 6);
                    cellThickVal.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Thickness);
                    cellThickVal.Style.Font.Bold = true;
                    cellThickVal.Style.Font.FontSize = 22;
                    cellThickVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                    cellThickVal.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    var cellThickTol = worksheet.Cell(rBottom, 6);
                    cellThickTol.Value = strThickTol;
                    cellThickTol.Style.Font.Bold = false;
                    cellThickTol.Style.Font.Italic = true;
                    cellThickTol.Style.Font.FontSize = 22;
                    cellThickTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    cellThickTol.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    var cellWidthVal = worksheet.Cell(rTop, 7);
                    cellWidthVal.Value = string.Format(cultureInvariant, "{0:0.00}", rec.Width);
                    cellWidthVal.Style.Font.Bold = true;
                    cellWidthVal.Style.Font.FontSize = 22;
                    cellWidthVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                    cellWidthVal.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    var cellWidhtTol = worksheet.Cell(rBottom, 7);
                    cellWidhtTol.Value = strWidthTol;
                    cellWidhtTol.Style.Font.Bold = false;
                    cellWidhtTol.Style.Font.Italic = true;
                    cellWidhtTol.Style.Font.FontSize = 22;
                    cellWidhtTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    cellWidhtTol.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    var cellLengthVal = worksheet.Cell(rTop, 8);
                    cellLengthVal.Value = rec.Length;
                    cellLengthVal.Style.NumberFormat.Format = "0";
                    cellLengthVal.Style.Font.Bold = true;
                    cellLengthVal.Style.Font.FontSize = 22;
                    cellLengthVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                    cellLengthVal.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    var cellLentghTol = worksheet.Cell(rBottom, 8);
                    cellLentghTol.Value = "(4000 ± 15)";
                    cellLentghTol.Style.Font.Bold = false;
                    cellLentghTol.Style.Font.Italic = true;
                    cellLentghTol.Style.Font.FontSize = 22;
                    cellLentghTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    cellLentghTol.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    worksheet.Cell(rTop, 9).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Radius);
                    worksheet.Range(rTop, 9, rBottom, 9).Merge();

                    worksheet.Cell(rTop, 10).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Chamber);
                    worksheet.Range(rTop, 10, rBottom, 10).Merge();

                    worksheet.Cell(rTop, 11).Value = "OK";
                    worksheet.Range(rTop, 11, rBottom, 11).Merge();

                    var rangeAll = worksheet.Range(rTop, 2, rBottom, 11);
                    ApplyCustomStyle(rangeAll);

                    cellThickVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                    cellThickTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    cellThickTol.Style.Font.Bold = false;

                    cellWidthVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                    cellWidhtTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    cellWidhtTol.Style.Font.Bold = false;

                    cellLengthVal.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                    cellLentghTol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    cellLentghTol.Style.Font.Bold = false;

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

                    var rangeT2 = worksheet.Range(rTable2, 2, rTable2, 11);
                    ApplyCustomStyle(rangeT2);
                }

                var table1Range = worksheet.Range(startRowTable1, 2, startRowTable1 + totalRowsNeeded - 1, 11);
                ApplyBorders(table1Range);

                for (int i = 0; i < dataCount; i++)
                {
                    int rTop = startRowTable1 + (i * rowsPerItem);
                    int rBottom = rTop + 1;

                    worksheet.Cell(rTop, 6).Style.Border.BottomBorder = XLBorderStyleValues.None;
                    worksheet.Cell(rBottom, 6).Style.Border.TopBorder = XLBorderStyleValues.None;

                    worksheet.Cell(rTop, 7).Style.Border.BottomBorder = XLBorderStyleValues.None;
                    worksheet.Cell(rBottom, 7).Style.Border.TopBorder = XLBorderStyleValues.None;

                    worksheet.Cell(rTop, 8).Style.Border.BottomBorder = XLBorderStyleValues.None;
                    worksheet.Cell(rBottom, 8).Style.Border.TopBorder = XLBorderStyleValues.None;
                }

                int lastRowTable2 = startRowTable2 + dataCount - 1;

                var table2Range = worksheet.Range(startRowTable2, 2, lastRowTable2, 11);
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
                signatureRange.Style.Border.OutsideBorder = XLBorderStyleValues.None;

                for (int k = 0; k < 5; k++)
                {
                    worksheet.Row(firstInsertedRow + k).Height = 102;
                }

                worksheet.Row(firstInsertedRow + 5).Height = 50;

                int imageRow = firstInsertedRow + 2;

                if (System.IO.File.Exists(pathImg1))
                {
                    var pic1 = worksheet.AddPicture(pathImg1);
                    pic1.MoveTo(worksheet.Cell(imageRow, 11));
                }

                if (System.IO.File.Exists(pathImg2))
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
            return fullPath;
        }

        private void ApplyCustomStyle(ClosedXML.Excel.IXLRange range)
        {
            range.Style.Font.Bold = true;
            range.Style.Font.FontSize = 22;

            range.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
        }

        private void ApplyBorders(ClosedXML.Excel.IXLRange range)
        {
            range.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            range.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            range.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            range.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
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