using System;
using System.Collections.Generic;
using System.Text;

namespace WpfApp1.Core.Models
{
    public struct BusbarRecord
    {
        public int Id;

        public string BatchNo;

        public string Size, Year, Month, ProdDate, BendTest;
        public double Thickness, Width, Length, Radius, Chamber, Electric, Resistivity, Elongation, Tensile, Spectro, Oxygen;
    }
}