namespace WpfApp1.Shared.Helpers
{
    public static class CellValidationHelper
    {
        public static bool ShouldHighlightBatchNo(string batchNo)
        {
            return string.IsNullOrWhiteSpace(batchNo);
        }

        public static bool ShouldHighlightLength(int length)
        {
            return length <= 0 || length < 4015 || length > 4000;
        }

        public static bool ShouldHighlightThickness(double thickness)
        {
            return thickness <= 0;
        }

        public static bool ShouldHighlightWidth(double width)
        {
            return width <= 0;
        }

        public static bool ShouldHighlightRadius(double radius)
        {
            return radius <= 0;
        }

        public static bool ShouldHighlightChamber(double chamber)
        {
            return chamber <= 0 || chamber > 3.5;
        }

        public static bool ShouldHighlightElectric(double electric)
        {
            return electric <= 0 || electric < 100;
        }

        public static bool ShouldHighlightResistivity(double resistivity)
        {
            return resistivity <= 0 || resistivity > 0.15328;
        }

        public static bool ShouldHighlightElongation(double elongation)
        {
            return elongation <= 0 || elongation < 15;
        }

        public static bool ShouldHighlightTensile(double tensile)
        {
            return tensile <= 0 || tensile < 245 || tensile > 315;
        }

        public static bool ShouldHighlightSpectro(double spectro)
        {
            return spectro <= 0 || spectro < 99.96;
        }

        public static bool ShouldHighlightOxygen(double oxygen)
        {
            return oxygen <= 0 || oxygen > 10;
        }

        public static bool ShouldHighlightThicknessWithTolerance(
            double actualThickness,
            double nominalThickness,
            double tolerance)
        {
            if (actualThickness <= 0) return true;

            double lowerBound = nominalThickness - tolerance;
            double upperBound = nominalThickness + tolerance;

            return actualThickness < lowerBound || actualThickness > upperBound;
        }

        public static bool ShouldHighlightWidthWithTolerance(
            double actualWidth,
            double nominalWidth,
            double tolerance)
        {
            if (actualWidth <= 0) return true;

            double lowerBound = nominalWidth - tolerance;
            double upperBound = nominalWidth + tolerance;

            return actualWidth < lowerBound || actualWidth > upperBound;
        }
    }
}