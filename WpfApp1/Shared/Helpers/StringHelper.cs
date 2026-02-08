using System;
using System.Collections.Generic;
using System.Text;

namespace WpfApp1.Shared.Helpers
{
    public static class StringHelper
    {
        public static string CleanSizeCOA(string rawSize)
        {
            if (string.IsNullOrWhiteSpace(rawSize))
            {
                return string.Empty;
            }

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            bool hasX = false;
            bool hasDigitsAfterX = false;

            foreach (char c in rawSize)
            {
                if (char.IsDigit(c))
                {
                    sb.Append(c);
                    if (hasX)
                    {
                        hasDigitsAfterX = true;
                    }
                }
                else if (c == 'x' || c == 'X')
                {
                    if (hasX)
                    {
                        break;
                    }
                    sb.Append('x');
                    hasX = true;
                }
                else if (c == '.' || c == ',')
                {
                    sb.Append('.');
                }
                else
                {
                    if (hasX && hasDigitsAfterX)
                    {
                        break;
                    }
                }
            }

            return sb.ToString();
        }

        public static string CleanSizeText(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;

            System.Text.StringBuilder resultBuilder = new System.Text.StringBuilder();

            string text = raw.ToUpper();

            int startIndex = -1;

            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsDigit(text[i]))
                {
                    startIndex = i;
                    break;
                }
            }

            if (startIndex == -1) return string.Empty;

            int idx = startIndex;

            bool hasX = false;
            int dimensionIterator = startIndex;

            while (dimensionIterator < text.Length)
            {
                char c = text[dimensionIterator];

                if (char.IsDigit(c))
                {
                    resultBuilder.Append(c);
                }
                else if (c == 'X')
                {
                    if (dimensionIterator + 1 < text.Length && char.IsDigit(text[dimensionIterator + 1]))
                    {
                        resultBuilder.Append('x');
                        hasX = true;
                    }
                    else
                    {
                        break;
                    }
                }
                else
                {
                    if (hasX) break;
                }
                dimensionIterator++;
            }

            if (!hasX || resultBuilder.Length == 0) return string.Empty;

            string keyword = string.Empty;

            System.Text.StringBuilder remainingBuilder = new System.Text.StringBuilder();
            for (int k = dimensionIterator; k < text.Length; k++)
            {
                if (char.IsLetterOrDigit(text[k]))
                {
                    remainingBuilder.Append(text[k]);
                }
            }
            string cleanRemaining = remainingBuilder.ToString();

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
                for (int j = 0; j < cleanRemaining.Length; j++)
                {
                    if (cleanRemaining[j] == 'B' && j + 1 < cleanRemaining.Length &&
                        char.IsDigit(cleanRemaining[j + 1]))
                    {
                        int bStart = j;
                        int bEnd = j + 1;
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
                resultBuilder.Append(' ');
                resultBuilder.Append(keyword);
            }

            return resultBuilder.ToString();
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