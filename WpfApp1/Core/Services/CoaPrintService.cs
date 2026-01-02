using System;
using System.Collections.Generic;
using WpfApp1.Core.Models;

namespace WpfApp1.Core.Services
{
    public class CoaPrintService
    {
        public string GenerateCoaExcel(
            string customerName,
            string poNumber,
            string doNumber,
            System.Collections.Generic.List<WpfApp1.Core.Models.BusbarExportItem> dataList)
        {
            string templatePath = @"C:\Users\mrrx\Documents\My Web Sites\H\TEMPLATE_COA_BUSBAR.xlsx";
            string basePath = @"C:\Users\mrrx\Documents\My Web Sites\H\COA";

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
            int nomorFile = existingFiles.Length + 1;

            string formattedFileNumber = nomorFile.ToString("00");
            string romanMonth = GetRomanMonth(now.Month);

            string fileName = $"{formattedFileNumber}. COA {customerName} {doNumber}.xlsx";
            string fullPath = System.IO.Path.Combine(finalDirectory, fileName);

            using (var workbook = new ClosedXML.Excel.XLWorkbook(templatePath))
            {
                var worksheet = workbook.Worksheet(1);

                worksheet.Cell("C12").Value = ": " + poNumber;
                worksheet.Cell("J12").Value = ": " + customerName;
                worksheet.Cell("J13").Value = ": " + now.ToString("dd/MM/yyyy");
                worksheet.Cell("J14").Value = ": " + $"{formattedFileNumber}/{romanMonth}/{now.Year}";

                int dataCount = dataList.Count;
                int startRowTable1 = 20;
                int originalStartRowTable2 = 30;
                int startRowTable2 = originalStartRowTable2;

                if (dataCount > 3)
                {
                    int rowsToInsert = dataCount - 3;
                    worksheet.Row(22).InsertRowsBelow(rowsToInsert);
                    startRowTable2 = originalStartRowTable2 + rowsToInsert;
                }

                var cultureInvariant = System.Globalization.CultureInfo.InvariantCulture;

                for (int i = 0; i < dataCount; i++)
                {
                    int r1 = startRowTable1 + i;
                    int r2 = startRowTable2 + i;

                    worksheet.Row(r1).Height = 102;
                    worksheet.Row(r2).Height = 102;

                    // 2. JANGAN ubah Width (Hapus logika worksheet.Column(col).Width) 
                    // agar mengikuti template yang sudah ada.
                    var rec = dataList[i].RecordData;

                    worksheet.Cell(r1, 2).Value = rec.BatchNo;
                    worksheet.Cell(r1, 3).Value = rec.Size;

                    var cellD = worksheet.Cell(r1, 4);
                    cellD.Value = "No Dirty\nNo Blackspot\nNo Blisters";
                    cellD.Style.Alignment.WrapText = true;
                    worksheet.Range(r1, 4, r1, 5).Merge();

                    worksheet.Cell(r1, 6).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Thickness);
                    worksheet.Cell(r1, 7).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Width);

                    if (double.TryParse(rec.Length.ToString(), out double lengthValue))
                    {
                        worksheet.Cell(r1, 8).Value = System.Convert.ToInt32(lengthValue);
                    }
                    else
                    {
                        worksheet.Cell(r1, 8).Value = rec.Length;
                    }

                    worksheet.Cell(r1, 9).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Radius);
                    worksheet.Cell(r1, 10).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Chamber);

                    worksheet.Cell(r1, 11).Value = "OK";

                    var rangeT1 = worksheet.Range(r1, 2, r1, 11);
                    ApplyCustomStyle(rangeT1);

                    worksheet.Cell(r2, 2).Value = rec.BatchNo;
                    worksheet.Cell(r2, 3).Value = rec.Size;

                    worksheet.Cell(r2, 4).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Electric);
                    //worksheet.Cell(r2, 5).Value = string.Format(cultureInvariant, "{0}", rec.Resistivity);
                    worksheet.Cell(r2, 5).Value = string.Format(cultureInvariant, "{0:0.00000}", rec.Resistivity);
                    worksheet.Cell(r2, 6).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Elongation);
                    worksheet.Cell(r2, 7).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Tensile);

                    worksheet.Cell(r2, 8).Value = "No Crack";

                    worksheet.Cell(r2, 9).Value = string.Format(cultureInvariant, "{0}", rec.Spectro);
                    worksheet.Cell(r2, 10).Value = string.Format(cultureInvariant, "{0:0.00}", rec.Oxygen);

                    worksheet.Cell(r2, 11).Value = "OK";

                    var rangeT2 = worksheet.Range(r2, 2, r2, 11);
                    ApplyCustomStyle(rangeT2);
                }

                ApplyBorders(worksheet.Range(startRowTable1, 2, startRowTable1 + dataCount - 1, 11));
                ApplyBorders(worksheet.Range(startRowTable2, 2, startRowTable2 + dataCount - 1, 11));

                worksheet.PageSetup.PagesTall = 1;
                worksheet.PageSetup.PagesWide = 1;

                workbook.SaveAs(fullPath);
            }

            return fullPath;
        }

        private void ApplyCustomStyle(ClosedXML.Excel.IXLRange range)
        {
            range.Style.Font.Bold = true;
            range.Style.Font.FontName = "Montserrat";
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