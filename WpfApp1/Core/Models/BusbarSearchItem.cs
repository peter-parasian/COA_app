namespace WpfApp1.Core.Models
{
    public class BusbarSearchItem
    {
        public System.Int32 No { get; set; }
        public System.String Specification { get; set; } = System.String.Empty;
        public System.String DateProd { get; set; } = System.String.Empty;
        public WpfApp1.Core.Models.BusbarRecord FullRecord { get; set; }
    }
}