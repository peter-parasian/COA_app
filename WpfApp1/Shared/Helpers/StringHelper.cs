using System;
using System.Collections.Generic;
using System.Text;

namespace WpfApp1.Shared.Helpers
{
    public static class StringHelper
    {
        public static string CleanSizeText(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;

            string text = raw.ToUpper();

            int start = -1;
            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsDigit(text[i]))
                {
                    start = i;
                    break;
                }
            }

            if (start == -1) return string.Empty;

            string substring = text.Substring(start);

            int xIndex = substring.IndexOf('X');
            if (xIndex == -1) return string.Empty;

            if (xIndex + 1 >= substring.Length || !char.IsDigit(substring[xIndex + 1]))
                return string.Empty;

            int end = xIndex + 1;
            while (end < substring.Length && char.IsDigit(substring[end]))
                end++;

            string result = substring.Substring(0, end).Trim();

            string remaining = substring.Substring(end).Trim();

            string keyword = string.Empty;

            string cleanRemaining = string.Empty;
            for (int i = 0; i < remaining.Length; i++)
            {
                if (char.IsLetterOrDigit(remaining[i]))
                {
                    cleanRemaining += remaining[i];
                }
            }

            if (cleanRemaining.Contains("FR"))
            {
                int frIndex = cleanRemaining.IndexOf("FR");
                if (frIndex >= 0)
                {
                    keyword = "FR";
                }
            }
            else if (!string.IsNullOrEmpty(cleanRemaining))
            {
                for (int i = 0; i < cleanRemaining.Length; i++)
                {
                    if (cleanRemaining[i] == 'B' && i + 1 < cleanRemaining.Length &&
                        char.IsDigit(cleanRemaining[i + 1]))
                    {
                        int bStart = i;
                        int bEnd = i + 1;
                        while (bEnd < cleanRemaining.Length && char.IsDigit(cleanRemaining[bEnd]))
                        {
                            bEnd++;
                        }

                        if (bEnd - bStart >= 2)
                        {
                            keyword = cleanRemaining.Substring(bStart, bEnd - bStart);
                            break;
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(keyword))
            {
                result = result + " " + keyword;
            }

            return result.Trim();
        }

        public static string DetermineTLJTable(string size_mm)
        {
            string cleanSize = size_mm.ToUpper().Replace(" ", "");

            int xIndex = cleanSize.IndexOf('X');
            if (xIndex == -1) return "TLJ500";

            string beforeX = cleanSize.Substring(0, xIndex);
            string afterX = cleanSize.Substring(xIndex + 1);

            string afterXDigits = "";
            for (int i = 0; i < afterX.Length; i++)
            {
                if (char.IsDigit(afterX[i]))
                {
                    afterXDigits += afterX[i];
                }
                else
                {
                    break;
                }
            }

            if (int.TryParse(beforeX, out int firstDimension) &&
                int.TryParse(afterXDigits, out int secondDimension))
            {
                if (firstDimension <= 10 && secondDimension <= 100)
                {
                    return "TLJ350";
                }
            }

            return "TLJ500";
        }

        public static string ProcessRawBatchString(string rawBatch)
        {
            if (string.IsNullOrEmpty(rawBatch)) return string.Empty;

            var batchList = new System.Collections.Generic.List<string>();
            string[] batches = rawBatch.Split(new[] { '\n', '\r' }, System.StringSplitOptions.RemoveEmptyEntries);
            foreach (string b in batches)
            {
                string t = b.Trim();
                if (!string.IsNullOrEmpty(t)) batchList.Add(t);
            }
            return string.Join("\n", batchList);
        }

        public static double ParseCustomDecimal(string rawInput)
        {
            if (string.IsNullOrWhiteSpace(rawInput)) return 0.0;
            string cleanInput = rawInput.Replace(",", ".").Trim();
            if (double.TryParse(cleanInput, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double result)) return result;
            return 0.0;
        }
    }
}