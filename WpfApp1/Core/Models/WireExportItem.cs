using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace WpfApp1.Core.Models
{
    public class WireExportItem : INotifyPropertyChanged
    {
        private WireRecord _record;

        public WireExportItem(WireRecord record)
        {
            _record = record;
        }

        public WireRecord RecordData => _record;

        public string CustomerName => _record.Customer;

        public string Specification => _record.Size;

        public string DateProd => _record.Date;

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}