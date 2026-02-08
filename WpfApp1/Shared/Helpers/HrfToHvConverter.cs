using System;
using System.Collections.Generic;

namespace WpfApp1.Shared.Helpers
{
    public static class HrfToHvConverter
    {
        private readonly struct HardnessPoint
        {
            public readonly double Hrf;
            public readonly double Hv;

            public HardnessPoint(double hrf, double hv)
            {
                Hrf = hrf;
                Hv = hv;
            }
        }

        private static readonly System.Collections.Generic.List<HardnessPoint> ConversionTable = new System.Collections.Generic.List<HardnessPoint>
        {
            new HardnessPoint(82.6, 90),
            new HardnessPoint(85.114, 95),
            new HardnessPoint(87.0, 100),
            new HardnessPoint(88.75, 105),
            new HardnessPoint(90.5, 110),
            new HardnessPoint(92.271, 115),
            new HardnessPoint(93.6, 120),
            new HardnessPoint(95.0, 125),
            new HardnessPoint(96.4, 130),
            new HardnessPoint(97.514, 135),
            new HardnessPoint(99.0, 140),
            new HardnessPoint(100.2, 145),
            new HardnessPoint(101.4, 150),
            new HardnessPoint(102.5, 155),
            new HardnessPoint(103.6, 160),
            new HardnessPoint(104.686, 165),
            new HardnessPoint(105.5, 170),
            new HardnessPoint(106.35, 175),
            new HardnessPoint(107.2, 180),
            new HardnessPoint(108.057, 185),
            new HardnessPoint(108.7, 190),
            new HardnessPoint(109.4, 195),
            new HardnessPoint(110.1, 200),
            new HardnessPoint(110.786, 205),
            new HardnessPoint(111.3, 210),
            new HardnessPoint(111.85, 215),
            new HardnessPoint(112.4, 220),
            new HardnessPoint(112.829, 225),
            new HardnessPoint(113.4, 230),
            new HardnessPoint(113.85, 235),
            new HardnessPoint(114.3, 240),
            new HardnessPoint(114.7, 245),
            new HardnessPoint(115.1, 250)
        };

        public static double Convert(double hrfInput)
        {
            if (ConversionTable.Count == 0) return 0;

            double minHrf = ConversionTable[0].Hrf; 
            double maxHrf = ConversionTable[ConversionTable.Count - 1].Hrf; 

            if (hrfInput < minHrf || hrfInput > maxHrf)
            {
                throw new System.ArgumentOutOfRangeException(nameof(hrfInput), $"Nilai HRF {hrfInput} diluar batas ({minHrf} - {maxHrf})");
            }

            for (int i = 0; i < ConversionTable.Count; i++)
            {
                HardnessPoint upperPoint = ConversionTable[i];

                if (upperPoint.Hrf >= hrfInput)
                {
                    if (upperPoint.Hrf == hrfInput || i == 0)
                    {
                        return upperPoint.Hv;
                    }

                    HardnessPoint lowerPoint = ConversionTable[i - 1];

                    // 3. Rumus Interpolasi Linear
                    // Scale = (UpperHRF - Input) / (UpperHRF - LowerHRF)
                    double scale = (upperPoint.Hrf - hrfInput) / (upperPoint.Hrf - lowerPoint.Hrf);

                    // Result = UpperHV - (Scale * (UpperHV - LowerHV))
                    double resultHv = upperPoint.Hv - (scale * (upperPoint.Hv - lowerPoint.Hv));

                    return resultHv;
                }
            }
            return 0;
        }
    }
}