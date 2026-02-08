using System.Collections.ObjectModel;
using WpfApp1.Core.Models;

namespace WpfApp1.ViewModels
{
    public class SheetModel : BaseViewModel
    {
        private string _sheetName = string.Empty;

        public string SheetName
        {
            get => _sheetName;
            set
            {
                _sheetName = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<BusbarExportItem> Items { get; set; }
            = new ObservableCollection<BusbarExportItem>();

        public SheetModel(string name)
        {
            _sheetName = name;
        }
    }
}