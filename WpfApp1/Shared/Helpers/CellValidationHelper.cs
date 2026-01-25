namespace WpfApp1.Shared.Helpers
{
    public static class CellValidationHelper
    {
        public static System.Boolean ShouldHighlightBatchNo(System.String batchNo)
        {
            return System.String.IsNullOrWhiteSpace(batchNo);
        }

        public static System.Boolean ShouldHighlightLength(System.Int32 length)
        {
            return length <= 0 || length < 4015 || length > 4000;
        }

        public static System.Boolean ShouldHighlightThickness(System.Double thickness)
        {
            return thickness <= 0;
        }

        public static System.Boolean ShouldHighlightWidth(System.Double width)
        {
            return width <= 0;
        }

        public static System.Boolean ShouldHighlightRadius(System.Double radius)
        {
            return radius <= 0;
        }

        public static System.Boolean ShouldHighlightChamber(System.Double chamber)
        {
            return chamber <= 0 || chamber > 3.5;
        }

        public static System.Boolean ShouldHighlightElectric(System.Double electric)
        {
            return electric <= 0 || electric < 100;
        }

        public static System.Boolean ShouldHighlightResistivity(System.Double resistivity)
        {
            return resistivity <= 0 || resistivity > 0.15328;
        }

        public static System.Boolean ShouldHighlightElongation(System.Double elongation)
        {
            return elongation <= 0 || elongation < 15;
        }

        public static System.Boolean ShouldHighlightTensile(System.Double tensile)
        {
            return tensile <= 0 || tensile < 245 || tensile > 315;
        }

        public static System.Boolean ShouldHighlightSpectro(System.Double spectro)
        {
            return spectro <= 0 || spectro < 99.96;
        }

        public static System.Boolean ShouldHighlightOxygen(System.Double oxygen)
        {
            return oxygen <= 0 || oxygen > 10;
        }

        public static System.Boolean ShouldHighlightThicknessWithTolerance(
            System.Double actualThickness,
            System.Double nominalThickness,
            System.Double tolerance)
        {
            if (actualThickness <= 0)
            {
                return true;
            }

            System.Double lowerBound = nominalThickness - tolerance;
            System.Double upperBound = nominalThickness + tolerance;

            return actualThickness < lowerBound || actualThickness > upperBound;
        }

        public static System.Boolean ShouldHighlightWidthWithTolerance(
            System.Double actualWidth,
            System.Double nominalWidth,
            System.Double tolerance)
        {
            if (actualWidth <= 0)
            {
                return true;
            }

            System.Double lowerBound = nominalWidth - tolerance;
            System.Double upperBound = nominalWidth + tolerance;

            return actualWidth < lowerBound || actualWidth > upperBound;
        }
    }
}