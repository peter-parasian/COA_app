using System.Runtime.CompilerServices;

namespace WpfApp1.Shared.Helpers
{
    public static class ToleranceJIS
    {
        private static readonly System.Collections.Generic.Dictionary<(double, double), double> _widthExceptions = new()
        {
            { (6, 12), 1.25 },
            { (9.65, 152.4), 1.52 },
            { (12, 100), 1.20 },
            { (12, 120), 1.25 },
            { (12, 125), 1.00 },
            { (12.7, 76.2), 1.02 },
            { (12.7, 101.6), 1.27 },
            { (12.7, 127), 1.00 },
            { (15, 60), 1.50 },
            { (15, 80), 1.50 },
            { (15, 100), 1.50 },
            { (15, 150), 1.50 },
            { (5, 125), 1.25 },
            { (6, 125), 1.25 },
            { (7.76, 12.7), 1.02 }
        };

        public static (double Thickness, double Width) CalculateFromDbString(string rawInput)
        {
            if (string.IsNullOrEmpty(rawInput)) return (0, 0);

            System.ReadOnlySpan<char> span = rawInput.AsSpan();
            int xIndex = span.IndexOf('x'); 

            if (xIndex == -1) return (0, 0); 

            double d1 = double.Parse(span.Slice(0, xIndex), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);
            double d2 = double.Parse(span.Slice(xIndex + 1), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture);

            return ExecuteLogic(d1, d2);
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
        private static (double, double) ExecuteLogic(double d1, double d2)
        {
            if (_widthExceptions.TryGetValue((d1, d2), out double width))
            {
                return (GetThicknessTol(d1), width);
            }

            return (GetThicknessTol(d1), GetWidthRule(d1, d2));
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
        private static double GetThicknessTol(double d1)
        {
            if (d1 < 4) return 0.08;
            if (d1 < 6) return 0.10;
            if (d1 < 10) return 0.12;
            if (d1 < 15) return 0.15;
            return 0.20;
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
        private static double GetWidthRule(double d1, double d2)
        {
            if (d1 < 4) return 0.80;
            if (d1 < 6) return 1.00;
            if (d1 < 10) return 1.00;

            if (d1 == 10) 
            {
                if (d2 >= 200) return 2.00;
                if (d2 >= 180) return 1.80;
                if (d2 >= 160) return 1.60;
                if (d2 >= 150) return 1.50;
                if (d2 >= 140) return 1.40;
                if (d2 >= 125) return 1.25;
                if (d2 >= 120) return 1.20;
                return 1.00;
            }

            if (d1 < 15) return 1.00;
            return 1.50;
        }
    }
}