using System.Collections.ObjectModel;
using WpfApp1.Core.Models;

namespace WpfApp1.ViewModels
{
    public class WireSheetModel : BaseViewModel
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

        public ObservableCollection<WireExportItem> Items { get; set; }
            = new ObservableCollection<WireExportItem>();

        public WireSheetModel(string name)
        {
            _sheetName = name;
        }
    }
}