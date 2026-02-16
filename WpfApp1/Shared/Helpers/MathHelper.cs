using System;
using System.Collections.Generic;
using System.Text;
using WpfApp1.Shared.Helpers;

namespace WpfApp1.Shared.Helpers
{
    public static class MathHelper
    {
        public static (double Tensile, double Elongation) CalculateTensileAndElongation(double t1, double t2, double e1, double e2)
        {
            double finalTensile = 0;
            double finalElongation = 0;

            bool hasT1 = t1 > 0;
            bool hasT2 = t2 > 0;
            bool hasE1 = e1 > 0;
            bool hasE2 = e2 > 0;

            if (hasT1 && hasT2 && hasE1 && hasE2)
            {
                if (t1 <= t2)
                {
                    finalTensile = t1;
                    finalElongation = e1;
                }
                else
                {
                    finalTensile = t2;
                    finalElongation = e2;
                }
            }
            else if (hasT1 && hasT2 && (hasE1 || hasE2))
            {
                finalTensile = System.Math.Min(t1, t2);
                finalElongation = System.Math.Max(e1, e2);
            }
            else if ((hasT1 || hasT2) && hasE1 && hasE2)
            {
                finalTensile = System.Math.Max(t1, t2);
                finalElongation = System.Math.Max(e1, e2);
            }
            else
            {
                finalTensile = System.Math.Max(t1, t2);
                finalElongation = System.Math.Max(e1, e2);
            }
            return (System.Math.Round(finalTensile, 2), System.Math.Round(finalElongation, 2));
        }

        public static double CalculateTensileStrengthKgmm2(double tensileNmm2)
        {
            if (tensileNmm2 < 0)
            {
                return 0.0;
            }

            double tensileStrengthKgmm2 = tensileNmm2 * 0.101972;
            return tensileStrengthKgmm2;
        }

        public static double CalculateElectricalResistivity(double iacs)
        {
            if (iacs <= 0)
            {
                return 0.0;
            }

            double electricalResistivity = 1.7241 / iacs;
            return electricalResistivity;
        }

        public static double CalculateElectricalConductivity(double iacs)
        {
            if (iacs <= 0)
            {
                return 0.0;
            }

            double electricalConductivity = iacs / 1.7241;
            return electricalConductivity; 
        }

        public static double CalculateConductorResisten(double iacs, double diameter)
        {
            if (iacs <= 0 || diameter <= 0)
            {
                return 0.0;
            }

            double conductorResisten = (1.7241 * 4.0) / (System.Math.PI * System.Math.Pow(diameter, 2) * iacs);
            return conductorResisten;
        }
    }
}