using System;
using System.Collections.Generic;
using System.Text;
using WpfApp1.Shared.Helpers;

namespace WpfApp1.Shared.Helpers
{
    public static class MathHelper
    {
        public static double GetMergedOrAverageValue(System.Data.DataTable table, int rowIndex, int colIndex)
        {
            if (rowIndex >= table.Rows.Count) return 0.0;

            object raw1 = table.Rows[rowIndex][colIndex];
            string str1 = raw1 != null ? raw1.ToString() ?? "" : "";
            double val1 = StringHelper.ParseCustomDecimal(str1);

            if (rowIndex + 1 < table.Rows.Count)
            {
                object raw2 = table.Rows[rowIndex + 1][colIndex];
                string str2 = raw2 != null ? raw2.ToString() ?? "" : "";
                double val2 = StringHelper.ParseCustomDecimal(str2);

                if (val1 == 0) return val2;
                if (val2 == 0) return val1;

                return (val1 + val2) / 2.0;
            }

            return val1;
        }
    }
}