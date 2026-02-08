namespace WpfApp1.Core.Models
{
    public class WireSearchItem
    {
        public int No { get; set; }
        public string Specification { get; set; } = string.Empty;
        public string CustomerName { get; set; } = string.Empty;
        public string DateProd { get; set; } = string.Empty;
        public WireRecord FullRecord { get; set; }
    }
}