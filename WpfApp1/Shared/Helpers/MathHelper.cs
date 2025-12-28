using System;
using System.Collections.Generic;
using System.Text;
using WpfApp1.Shared.Helpers;

namespace WpfApp1.Shared.Helpers
{
    public static class MathHelper
    {
        public static double GetMergedOrAverageValue(ClosedXML.Excel.IXLWorksheet sheet_YLB, int startRow, string columnLetter)
        {
            var cellFirst = sheet_YLB.Cell(startRow, columnLetter);
            if (cellFirst.IsMerged()) return StringHelper.ParseCustomDecimal(cellFirst.GetString());

            var val1 = StringHelper.ParseCustomDecimal(cellFirst.GetString());
            var val2 = StringHelper.ParseCustomDecimal(sheet_YLB.Cell(startRow + 1, columnLetter).GetString());

            if (val1 == 0) return val2;
            if (val2 == 0) return val1;

            return (val1 + val2) / 2.0;
        }
    }
}