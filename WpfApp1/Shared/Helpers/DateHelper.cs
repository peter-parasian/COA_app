namespace WpfApp1.Shared.Helpers
{
    public static class DateHelper
    {
        public static System.String StandardizeDate(System.String rawDate, System.Int32 expectedMonth, System.Int32 expectedYear)
        {
            if (System.String.IsNullOrWhiteSpace(rawDate))
            {
                return System.String.Empty;
            }

            if (System.DateTime.TryParse(rawDate, out System.DateTime parsedDate))
            {
                if (expectedYear > 2000 && parsedDate.Year != expectedYear)
                {
                    parsedDate = new System.DateTime(expectedYear, parsedDate.Month, parsedDate.Day);
                }

                if (expectedMonth > 0 && parsedDate.Month != expectedMonth)
                {
                    if (parsedDate.Day <= 12)
                    {
                        System.Int32 newMonth = parsedDate.Day;
                        System.Int32 newDay = parsedDate.Month;

                        if (newMonth == expectedMonth)
                        {
                            parsedDate = new System.DateTime(parsedDate.Year, newMonth, newDay);
                        }
                    }
                }

                return parsedDate.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            }

            return rawDate;
        }

        private static readonly System.Collections.Generic.Dictionary<System.String, System.Int32> _monthCache =
            new System.Collections.Generic.Dictionary<System.String, System.Int32>(System.StringComparer.OrdinalIgnoreCase);

        public static System.Int32 GetMonthNumber(System.String monthName)
        {
            if (System.String.IsNullOrWhiteSpace(monthName))
            {
                return 0;
            }

            if (_monthCache.TryGetValue(monthName, out System.Int32 cachedMonth))
            {
                return cachedMonth;
            }

            try
            {
                System.Int32 month = System.DateTime.ParseExact(monthName, "MMMM",
                    System.Globalization.CultureInfo.InvariantCulture).Month;
                _monthCache[monthName] = month;
                return month;
            }
            catch
            {
                return 0;
            }
        }

        public static System.String NormalizeMonthFolder(System.String rawMonth)
        {
            if (System.String.IsNullOrWhiteSpace(rawMonth))
            {
                return System.String.Empty;
            }

            for (System.Int32 i = 0; i < rawMonth.Length; i++)
            {
                if (!System.Char.IsDigit(rawMonth[i]))
                {
                    continue;
                }

                System.Int32 start = i;
                while (i < rawMonth.Length && System.Char.IsDigit(rawMonth[i]))
                {
                    i++;
                }

                System.String numberText = rawMonth.Substring(start, i - start);
                if (!System.Int32.TryParse(numberText, out System.Int32 monthNumber))
                {
                    continue;
                }

                if (monthNumber < 1 || monthNumber > 12)
                {
                    continue;
                }

                return new System.Globalization.DateTimeFormatInfo().GetMonthName(monthNumber);
            }
            return System.String.Empty;
        }
    }
}