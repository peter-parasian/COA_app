using System;
using System.Collections.Generic;
using System.Globalization;

namespace WpfApp1.Shared.Helpers
{
    public static class WireValidationHelper
    {
        private enum ValidationUnit { OhmM, OhmKM }
        private class WireTolerance
        {
            public double DiameterNom { get; set; } = 0.0;
            public double DiameterTol { get; set; } = 0.0;
            public double ElongMin { get; set; } = double.MinValue;
            public double ElongMax { get; set; } = double.MaxValue;
            public double TensileKgMin { get; set; } = double.MinValue;
            public double TensileKgMax { get; set; } = double.MaxValue;
            public double CopperContentMin { get; set; } = double.MinValue;
            public double CondResMax { get; set; } = double.MaxValue;
            public ValidationUnit CondResUnit { get; set; } = ValidationUnit.OhmKM; 
            public double IACSMin { get; set; } = double.MinValue;
            public double YieldMin { get; set; } = double.MinValue;
            public double TensileNMin { get; set; } = double.MinValue;
            public double TensileNMax { get; set; } = double.MaxValue;
            public double ElecCondMin { get; set; } = double.MinValue;
            public double ElecResMax { get; set; } = double.MaxValue;

        }

        private static readonly Dictionary<string, WireTolerance> _wireTolerances = new Dictionary<string, WireTolerance>(StringComparer.OrdinalIgnoreCase);

        static WireValidationHelper()
        {

            // PT. CANNING INDONESIAN PRODUCTS (Size 1.20) -> Canning|1.20
            WireTolerance tol1 = new WireTolerance();
            tol1.DiameterNom = 1.20;
            tol1.DiameterTol = 0.01;
            tol1.ElongMin = 25.0;
            tol1.ElongMax = 31.0;
            tol1.TensileKgMin = 22.0;
            tol1.TensileKgMax = 30.0;
            tol1.CopperContentMin = 99.95;
            tol1.CondResMax = 14.51;
            tol1.IACSMin = 100.0;
            tol1.CondResUnit = ValidationUnit.OhmKM;
            _wireTolerances["Canning|1.20"] = tol1;

            // PT. INDOLAKTO (Size 1.20) -> Indolakto|1.20
            WireTolerance tol2 = new WireTolerance();
            tol2.DiameterNom = 1.20;
            tol2.DiameterTol = 0.01;
            tol2.ElongMin = 25.0;
            tol2.ElongMax = 31.0;
            tol2.TensileKgMin = 22.0;
            tol2.TensileKgMax = 30.0;
            tol2.IACSMin = 100.0;
            _wireTolerances["Indolakto|1.20"] = tol2;

            // PT. Indonesia Multi Colour Printing (Size 1.20) -> Multi Colour|1.20
            WireTolerance tol3 = new WireTolerance();
            tol3.DiameterNom = 1.20;
            tol3.DiameterTol = 0.01;
            tol3.ElongMin = 28.0;
            tol3.ElongMax = 34.0;
            tol3.TensileKgMin = 22.0;
            tol3.TensileKgMax = 30.0;
            tol3.IACSMin = 100.0;
            _wireTolerances["Multi Colour|1.20"] = tol3;

            // PT. NESTLE INDONESIA (Size 1.20) -> Nestle|1.20
            WireTolerance tol4 = new WireTolerance();
            tol4.DiameterNom = 1.20;
            tol4.DiameterTol = 0.04;
            tol4.ElongMin = 22.0;
            tol4.ElongMax = 28.0;
            tol4.YieldMin = 180.0;
            tol4.TensileNMin = 245.0;
            tol4.TensileNMax = 285.0;
            tol4.CopperContentMin = 99.95;
            tol4.ElecCondMin = 57.50;
            tol4.ElecResMax = 0.01739;
            _wireTolerances["Nestle|1.20"] = tol4;

            // PT. CANNING INDONESIAN PRODUCTS (Size 1.24) -> Canning|1.24
            WireTolerance tol5 = new WireTolerance();
            tol5.DiameterNom = 1.24;
            tol5.DiameterTol = 0.01;
            tol5.ElongMin = 25.0;
            tol5.ElongMax = 31.0;
            tol5.TensileKgMin = 22.0;
            tol5.TensileKgMax = 30.0;
            tol5.CopperContentMin = 99.95;
            tol5.CondResMax = 14.51;
            tol5.IACSMin = 100.0;
            tol1.CondResUnit = ValidationUnit.OhmKM; 
            _wireTolerances["Canning|1.24"] = tol5;

            // PT. COMETA CAN (Size 1.24) -> Cometa|1.24
            WireTolerance tol6 = new WireTolerance();
            tol6.DiameterNom = 1.24;
            tol6.DiameterTol = 0.01;
            tol6.ElongMin = 25.0;
            tol6.ElongMax = 31.0;
            tol6.TensileKgMin = 22.0;
            tol6.TensileKgMax = 30.0;
            tol6.CopperContentMin = 99.95;
            tol6.ElecResMax = 0.017241;
            tol6.IACSMin = 100.0;
            _wireTolerances["Cometa|1.24"] = tol6;

            // PT. Indonesia Multi Colour Printing (Size 1.24) -> Multi Colour|1.24
            WireTolerance tol7 = new WireTolerance();
            tol7.DiameterNom = 1.24;
            tol7.DiameterTol = 0.01;
            tol7.ElongMin = 28.0;
            tol7.ElongMax = 34.0;
            tol7.TensileKgMin = 22.0;
            tol7.TensileKgMax = 30.0;
            tol7.IACSMin = 100.0;
            _wireTolerances["Multi Colour|1.24"] = tol7;

            // PT. ALMICOS PRATAMA (Size 1.38) -> Almicos|1.38
            WireTolerance tol9 = new WireTolerance();
            tol9.DiameterNom = 1.38;
            tol9.DiameterTol = 0.01;
            tol9.ElongMin = 28.0;
            tol9.ElongMax = 34.0;
            tol9.TensileKgMin = 22.0;
            tol9.TensileKgMax = 30.0;
            tol9.IACSMin = 100.0;
            _wireTolerances["Almicos|1.38"] = tol9;

            // PT. Avia Avian (Size 1.38) -> Avia Avian|1.38
            WireTolerance tol10 = new WireTolerance();
            tol10.DiameterNom = 1.38;
            tol10.DiameterTol = 0.01;
            tol10.ElongMin = 28.0;
            tol10.ElongMax = 34.0;
            tol10.TensileKgMin = 22.0;
            tol10.TensileKgMax = 30.0;
            tol10.IACSMin = 100.0;
            _wireTolerances["Avia Avian|1.38"] = tol10;

            // PT. COMETA CAN (Size 1.38) -> Cometa|1.38
            WireTolerance tol8 = new WireTolerance();
            tol8.DiameterNom = 1.38;
            tol8.DiameterTol = 0.01;
            tol8.ElongMin = 28.0;
            tol8.ElongMax = 34.0;
            tol8.TensileKgMin = 22.0;
            tol8.TensileKgMax = 30.0;
            tol8.CopperContentMin = 99.95;
            tol8.ElecResMax = 0.017241;
            tol8.IACSMin = 100.0;
            _wireTolerances["Cometa|1.38"] = tol8;

            // PT. EKA TIMUR RAYA (Size 1.38) -> Eka Timur|1.38
            WireTolerance tol11 = new WireTolerance();
            tol11.DiameterNom = 1.38;
            tol11.DiameterTol = 0.01;
            tol11.ElongMin = 28.0;
            tol11.ElongMax = 34.0;
            tol11.TensileKgMin = 22.0;
            tol11.TensileKgMax = 30.0;
            tol11.IACSMin = 100.0;
            _wireTolerances["Eka Timur|1.38"] = tol11;

            // PT. Energy Lautan Nusantara (Size 1.38) -> Energy Lautan|1.38
            WireTolerance tol12 = new WireTolerance();
            tol12.DiameterNom = 1.38;
            tol12.DiameterTol = 0.03;
            tol12.ElongMin = 28.0;
            tol12.ElongMax = 34.0;
            tol12.TensileKgMin = 22.0;
            tol12.TensileKgMax = 30.0;
            tol12.IACSMin = 100.0;
            _wireTolerances["Energy Lautan|1.38"] = tol12;

            // MASAMI PASIFIK (Size 1.38) -> Masami Pasifik|1.38
            WireTolerance tol13 = new WireTolerance();
            tol13.DiameterNom = 1.38;
            tol13.DiameterTol = 0.01;
            tol13.ElongMin = 28.0;
            tol13.ElongMax = 34.0;
            tol13.TensileKgMin = 22.0;
            tol13.TensileKgMax = 30.0;
            tol13.IACSMin = 100.0;
            _wireTolerances["Masami Pasifik|1.38"] = tol13;

            // PT. Metal Manufakturing Indonesia (Size 1.38) -> Metal Manufacturing|1.38
            WireTolerance tol14 = new WireTolerance();
            tol14.DiameterNom = 1.38;
            tol14.DiameterTol = 0.01;
            tol14.ElongMin = 28.0;
            tol14.ElongMax = 34.0;
            tol14.TensileKgMin = 22.0;
            tol14.TensileKgMax = 30.0;
            tol14.IACSMin = 100.0;
            _wireTolerances["Metal Manufacturing|1.38"] = tol14;

            // PT. Indonesia Multi Colour Printing (Size 1.38) -> Multi Colour|1.38
            WireTolerance tol15 = new WireTolerance();
            tol15.DiameterNom = 1.38;
            tol15.DiameterTol = 0.01;
            tol15.ElongMin = 28.0;
            tol15.ElongMax = 34.0;
            tol15.TensileKgMin = 22.0;
            tol15.TensileKgMax = 30.0;
            tol15.IACSMin = 100.0;
            _wireTolerances["Multi Colour|1.38"] = tol15;

            // PT. PRISMACABLE MITRATAMA INDUSTRIES (Size 1.38) -> Prisma Cable|1.38
            WireTolerance tol16 = new WireTolerance();
            tol16.DiameterNom = 1.38;
            tol16.DiameterTol = 0.03;
            tol16.ElongMin = 25.0;
            tol16.ElongMax = double.MaxValue;
            tol16.TensileKgMin = 22.0;
            tol16.TensileKgMax = 30.0;
            tol16.CopperContentMin = 99.95;
            tol16.IACSMin = 100.0;
            _wireTolerances["Prisma Cable|1.38"] = tol16;

            // PT. Avia Avian (Size 1.50) -> Avia Avian|1.50
            WireTolerance tol18 = new WireTolerance();
            tol18.DiameterNom = 1.50;
            tol18.DiameterTol = 0.01;
            tol18.ElongMin = 28.0;
            tol18.ElongMax = 34.0;
            tol18.TensileKgMin = 22.0;
            tol18.TensileKgMax = 30.0;
            tol18.IACSMin = 100.0;
            _wireTolerances["Avia Avian|1.50"] = tol18;

            // PT. COMETA CAN (Size 1.50) -> Cometa|1.50
            WireTolerance tol17 = new WireTolerance();
            tol17.DiameterNom = 1.50;
            tol17.DiameterTol = 0.01;
            tol17.ElongMin = 28.0;
            tol17.ElongMax = 34.0;
            tol17.TensileKgMin = 22.0;
            tol17.TensileKgMax = 30.0;
            tol17.CopperContentMin = 99.95;
            tol17.ElecResMax = 0.017241;
            tol17.IACSMin = 100.0;
            _wireTolerances["Cometa|1.50"] = tol17;

            // PT. Indonesia Multi Colour Printing (Size 1.50) -> Multi Colour|1.50
            WireTolerance tol19 = new WireTolerance();
            tol19.DiameterNom = 1.50;
            tol19.DiameterTol = 0.01;
            tol19.ElongMin = 28.0;
            tol19.ElongMax = 34.0;
            tol19.TensileKgMin = 22.0;
            tol19.TensileKgMax = 30.0;
            tol19.IACSMin = 100.0;
            _wireTolerances["Multi Colour|1.50"] = tol19;

            // PT. Indowire Prima Soft (Size 1.60) -> Indowire|1.60
            WireTolerance tol20 = new WireTolerance();
            tol20.DiameterNom = 1.60;
            tol20.DiameterTol = 0.03;
            tol20.ElongMin = 30.0;
            tol20.ElongMax = double.MaxValue;
            tol20.IACSMin = 100.0;
            _wireTolerances["Indowire (Soft)|1.60"] = tol20;

            // PT. Indowire Prima Hard (Size 1.60) -> Indowire|1.60
            WireTolerance tol23 = new WireTolerance();
            //tol23.DiameterNom = 1.60;
           // tol23.DiameterTol = 0.03;
           // tol23.IACSMin = 100.0;
            _wireTolerances["Indowire (Hard)|1.60"] = tol23;

            // PT. Magnakabel Nusantara (Size 1.60) -> Magnakabel|1.60
            WireTolerance tol21 = new WireTolerance();
            tol21.DiameterNom = 1.60;
            tol21.DiameterTol = 0.03;
            tol21.ElongMin = 25.0;
            tol21.ElongMax = double.MaxValue;
            tol21.IACSMin = 100.0;
            _wireTolerances["Magnakabel|1.60"] = tol21;

            // PT. Metal Manufakturing Indonesia (Size 1.60) -> Metal Manufacturing|1.60
            WireTolerance tol22 = new WireTolerance();
            tol22.DiameterNom = 1.60;
            tol22.DiameterTol = 0.03;
            tol22.ElongMin = 30.0;
            tol22.ElongMax = double.MaxValue;
            tol22.TensileKgMin = 22.0;
            tol22.TensileKgMax = 30.0;
            tol22.IACSMin = 100.0;
            _wireTolerances["Metal Manufacturing|1.60"] = tol22;
        }

        private static string GetWireKey(string customer, string size)
        {
            if (string.IsNullOrWhiteSpace(customer) || string.IsNullOrWhiteSpace(size)) return string.Empty;

            double sizeNum;
            if (!double.TryParse(size, NumberStyles.Float, CultureInfo.InvariantCulture, out sizeNum)) return string.Empty;

            string formattedSize = sizeNum.ToString("0.00", CultureInfo.InvariantCulture);

            return customer.Trim() + "|" + formattedSize;
        }

        public static bool ShouldHighlightDiameter(string customer, string size, double value)
        {
            string key = GetWireKey(customer, size);
            if (string.IsNullOrEmpty(key) || !_wireTolerances.TryGetValue(key, out WireTolerance? tol) || tol == null) return false;

            if (tol.DiameterNom == 0.0 || tol.DiameterTol == 0.0) return false;

            double min = tol.DiameterNom - tol.DiameterTol;
            double max = tol.DiameterNom + tol.DiameterTol;

            return value <= 0 || value < min || value > max;
        }

        public static bool ShouldHighlightElongation(string customer, string size, double value)
        {
            string key = GetWireKey(customer, size);
            if (string.IsNullOrEmpty(key) || !_wireTolerances.TryGetValue(key, out WireTolerance? tol) || tol == null) return false;

            return value <= 0 || value < tol.ElongMin || value > tol.ElongMax;
        }

        public static bool ShouldHighlightYield(string customer, string size, double value)
        {
            string key = GetWireKey(customer, size);
            if (string.IsNullOrEmpty(key) || !_wireTolerances.TryGetValue(key, out WireTolerance? tol) || tol == null) return false;

            return value <= 0 || value < tol.YieldMin;
        }

        public static bool ShouldHighlightTensile(string customer, string size, double tensileN)
        {
            string key = GetWireKey(customer, size);
            if (string.IsNullOrEmpty(key) || !_wireTolerances.TryGetValue(key, out WireTolerance? tol) || tol == null) return false;

            if (tensileN <= 0) return true;

            if (tol.TensileNMin != double.MinValue || tol.TensileNMax != double.MaxValue)
            {
                if (tensileN < tol.TensileNMin || tensileN > tol.TensileNMax) return true;
            }
            else if (tol.TensileKgMin != double.MinValue || tol.TensileKgMax != double.MaxValue)
            {
                double tensileKg = MathHelper.CalculateTensileStrengthKgmm2(tensileN);
                if (tensileKg < tol.TensileKgMin || tensileKg > tol.TensileKgMax) return true;
            }

            return false;
        }

        public static bool ShouldHighlightIACS(string customer, string size, double value)
        {
            string key = GetWireKey(customer, size);
            if (string.IsNullOrEmpty(key) || !_wireTolerances.TryGetValue(key, out WireTolerance? tol) || tol == null) return false;

            return value <= 0 || value < tol.IACSMin;
        }

        public static bool ShouldHighlightElectricalResistivity(string customer, string size, double iacs)
        {
            string key = GetWireKey(customer, size);
            if (string.IsNullOrEmpty(key) || !_wireTolerances.TryGetValue(key, out WireTolerance? tol) || tol == null) return false;

            if (iacs <= 0) return true;

            double res = MathHelper.CalculateElectricalResistivity(iacs);

            return res > tol.ElecResMax;
        }

        public static bool ShouldHighlightElectricalConductivity(string customer, string size, double iacs)
        {
            string key = GetWireKey(customer, size);
            if (string.IsNullOrEmpty(key) || !_wireTolerances.TryGetValue(key, out WireTolerance? tol) || tol == null) return false;

            if (iacs <= 0) return true;

            double cond = MathHelper.CalculateElectricalConductivity(iacs);

            return cond < tol.ElecCondMin;
        }

        public static bool ShouldHighlightConductorResistance(string customer, string size, double iacs, double diameter)
        {
            string key = GetWireKey(customer, size);
            if (string.IsNullOrEmpty(key) || !_wireTolerances.TryGetValue(key, out WireTolerance? tol) || tol == null) return false;

            if (iacs <= 0 || diameter <= 0) return true;

            double condRes = MathHelper.CalculateConductorResisten(iacs, diameter);

            if (tol.CondResUnit == ValidationUnit.OhmKM)
            {
                condRes = condRes * 1000.0;
            }

            return condRes > tol.CondResMax;
        }
    }
}