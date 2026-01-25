namespace WpfApp1.Shared.Helpers
{
    public static class StringHelper
    {
        public static System.String CleanSizeCOA(System.String rawSize)
        {
            if (System.String.IsNullOrWhiteSpace(rawSize))
            {
                return System.String.Empty;
            }

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            System.Boolean hasX = false;
            System.Boolean hasDigitsAfterX = false;

            foreach (System.Char c in rawSize)
            {
                if (System.Char.IsDigit(c))
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

        public static System.String CleanSizeText(System.String raw)
        {
            if (System.String.IsNullOrWhiteSpace(raw))
            {
                return System.String.Empty;
            }

            System.Text.StringBuilder resultBuilder = new System.Text.StringBuilder();

            System.String text = raw.ToUpper();

            System.Int32 startIndex = -1;

            for (System.Int32 i = 0; i < text.Length; i++)
            {
                if (System.Char.IsDigit(text[i]))
                {
                    startIndex = i;
                    break;
                }
            }

            if (startIndex == -1)
            {
                return System.String.Empty;
            }

            System.Int32 idx = startIndex;

            System.Boolean hasX = false;
            System.Int32 dimensionIterator = startIndex;

            while (dimensionIterator < text.Length)
            {
                System.Char c = text[dimensionIterator];

                if (System.Char.IsDigit(c))
                {
                    resultBuilder.Append(c);
                }
                else if (c == 'X')
                {
                    if (dimensionIterator + 1 < text.Length && System.Char.IsDigit(text[dimensionIterator + 1]))
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
                    if (hasX)
                    {
                        break;
                    }
                }
                dimensionIterator++;
            }

            if (!hasX || resultBuilder.Length == 0)
            {
                return System.String.Empty;
            }

            System.String keyword = System.String.Empty;

            System.Text.StringBuilder remainingBuilder = new System.Text.StringBuilder();
            for (System.Int32 k = dimensionIterator; k < text.Length; k++)
            {
                if (System.Char.IsLetterOrDigit(text[k]))
                {
                    remainingBuilder.Append(text[k]);
                }
            }
            System.String cleanRemaining = remainingBuilder.ToString();

            if (cleanRemaining.Contains("FR"))
            {
                System.Int32 frIndex = cleanRemaining.IndexOf("FR");
                if (frIndex >= 0)
                {
                    keyword = "FR";
                }
            }
            else if (!System.String.IsNullOrEmpty(cleanRemaining))
            {
                for (System.Int32 j = 0; j < cleanRemaining.Length; j++)
                {
                    if (cleanRemaining[j] == 'B' && j + 1 < cleanRemaining.Length &&
                        System.Char.IsDigit(cleanRemaining[j + 1]))
                    {
                        System.Int32 bStart = j;
                        System.Int32 bEnd = j + 1;
                        while (bEnd < cleanRemaining.Length && System.Char.IsDigit(cleanRemaining[bEnd]))
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

            if (!System.String.IsNullOrEmpty(keyword))
            {
                resultBuilder.Append(' ');
                resultBuilder.Append(keyword);
            }

            return resultBuilder.ToString();
        }

        public static System.String DetermineTLJTable(System.String size_mm)
        {
            System.String cleanSize = size_mm.ToUpper().Replace(" ", "");

            System.Int32 xIndex = cleanSize.IndexOf('X');
            if (xIndex == -1)
            {
                return "TLJ500";
            }

            System.String beforeX = cleanSize.Substring(0, xIndex);
            System.String afterX = cleanSize.Substring(xIndex + 1);

            System.String afterXDigits = "";
            for (System.Int32 i = 0; i < afterX.Length; i++)
            {
                if (System.Char.IsDigit(afterX[i]))
                {
                    afterXDigits += afterX[i];
                }
                else
                {
                    break;
                }
            }

            if (System.Int32.TryParse(beforeX, out System.Int32 firstDimension) &&
                System.Int32.TryParse(afterXDigits, out System.Int32 secondDimension))
            {
                if (firstDimension <= 10 && secondDimension <= 100)
                {
                    return "TLJ350";
                }
            }

            return "TLJ500";
        }

        public static System.String ProcessRawBatchString(System.String rawBatch)
        {
            if (System.String.IsNullOrEmpty(rawBatch))
            {
                return System.String.Empty;
            }

            System.Collections.Generic.List<System.String> batchList = new System.Collections.Generic.List<System.String>();
            System.String[] batches = rawBatch.Split(new[] { '\n', '\r' }, System.StringSplitOptions.RemoveEmptyEntries);
            foreach (System.String b in batches)
            {
                System.String t = b.Trim();
                if (!System.String.IsNullOrEmpty(t))
                {
                    batchList.Add(t);
                }
            }
            return System.String.Join("\n", batchList);
        }

        public static System.Double ParseCustomDecimal(System.String rawInput)
        {
            if (System.String.IsNullOrWhiteSpace(rawInput))
            {
                return 0.0;
            }
            System.String cleanInput = rawInput.Replace(",", ".").Trim();
            if (System.Double.TryParse(cleanInput, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out System.Double result))
            {
                return result;
            }
            return 0.0;
        }
    }
}