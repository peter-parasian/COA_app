using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace WpfApp1.Core.Models
{
    public class BusbarExportItem : INotifyPropertyChanged
    {
        private BusbarRecord _record;

        public BusbarExportItem(BusbarRecord record)
        {
            _record = record;
        }

        public BusbarRecord RecordData => _record;

        public int Id => _record.Id;
        public string Specification => _record.Size;
        public string DateProd => _record.ProdDate;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}