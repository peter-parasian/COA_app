using System.Runtime.CompilerServices;

namespace WpfApp1.Shared.Helpers
{
    public static class ToleranceJIS
    {
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

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static (double, double) ExecuteLogic(double d1, double d2)
        {
            return (GetThicknessTol(d1), GetWidthRule(d1, d2));
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static double GetThicknessTol(double thickness)
        {
            // Rumus: =IF(A1<=3.2,0.08,IF(A1<=5,0.1,IF(A1<=8,0.12,IF(A1<=12,0.15,IF(A1<=20,0.2,IF(A1<=30,1.2%*A1,0))))))
            if (thickness <= 3.2) return 0.08;
            if (thickness <= 5) return 0.1;
            if (thickness <= 8) return 0.12;
            if (thickness <= 12) return 0.15;
            if (thickness <= 20) return 0.2;
            if (thickness <= 30) return 0.012 * thickness; 

            return 0; 
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static double GetWidthRule(double thickness, double width)
        {
            // Rumus: =IF(AND(A1<=3.2,C1<=100),0.8,IF(AND(A1>3.2,C1<=100),1,IF(AND(A1>3.2,C1>100),1%*C1)))
            if (thickness <= 3.2 && width <= 100) return 0.8;

            if (thickness > 3.2)
            {
                if (width <= 100) return 1.0;
                if (width > 100) return 0.01 * width; 
            }

            return 0; 
        }
    }
}