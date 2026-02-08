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
            return length <= 0 || length < 4000 || length > 4015;
        }

        public static bool ShouldHighlightThicknessWithTolerance(
            double actualThickness,
            double nominalThickness,
            double tolerance)
        {
            if (actualThickness <= 0) return true;

            double min = nominalThickness - tolerance;
            double max = nominalThickness + tolerance;

            return actualThickness < min || actualThickness > max;
        }

        public static bool ShouldHighlightWidthWithTolerance(
            double actualWidth,
            double nominalWidth,
            double tolerance)
        {
            if (actualWidth <= 0) return true;

            double min = nominalWidth - tolerance;
            double max = nominalWidth + tolerance;

            return actualWidth < min || actualWidth > max;
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
    }
}