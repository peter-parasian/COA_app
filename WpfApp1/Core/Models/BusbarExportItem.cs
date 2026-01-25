namespace WpfApp1.Core.Models
{
    public class BusbarExportItem : System.ComponentModel.INotifyPropertyChanged
    {
        private readonly WpfApp1.Core.Models.BusbarRecord _record;
        private System.String _selectedType = "Select";

        public BusbarExportItem(WpfApp1.Core.Models.BusbarRecord record)
        {
            _record = record;
        }

        public WpfApp1.Core.Models.BusbarRecord RecordData
        {
            get
            {
                return _record;
            }
        }

        public System.Int32 Id
        {
            get
            {
                return _record.Id;
            }
        }

        public System.String Specification
        {
            get
            {
                return _record.Size;
            }
        }

        public System.String DateProd
        {
            get
            {
                return _record.ProdDate;
            }
        }

        public System.String SelectedType
        {
            get
            {
                return _selectedType;
            }
            set
            {
                if (_selectedType != value)
                {
                    _selectedType = value;
                    OnPropertyChanged();
                }
            }
        }

        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] System.String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
        }
    }
}