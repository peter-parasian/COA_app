using System;
using System.Collections.Generic;
using System.Text;

namespace WpfApp1.Core.Models
{
    public struct BusbarRecord
    {
        public int Id, Length;
        public string BatchNo;
        public string Size, Year, Month, ProdDate, BendTest;
        public double Thickness, Width, Radius, Chamber, Electric, Resistivity, Elongation, Tensile, Hardness, Spectro, Oxygen;
    }
}