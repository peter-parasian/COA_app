using MahApps.Metro.Controls;  
using System.Threading.Tasks;
using System.Windows;

namespace WpfApp1.Views
{
    // Ubah inheritance dari Window menjadi MetroWindow
    public partial class MainWindow : MetroWindow
    {
        private WpfApp1.ViewModels.MainViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();

            _viewModel = new WpfApp1.ViewModels.MainViewModel();

            this.DataContext = _viewModel;

            _viewModel.OnShowMessage += (msg) =>
            {
                // Gunakan MahApps MessageDialogStyle (Opsional, disini tetap MessageBox standar dulu)
                System.Windows.MessageBox.Show(
                    msg,
                    "Informasi Sistem",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
            };
        }

        private async void ButtonMode1_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (_viewModel.IsBusy) return;

            try
            {
                _viewModel.IsBusy = true;

                await System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        _viewModel.ImportExcelToSQLite();
                    }
                    catch (System.Exception ex)
                    {
                        this.Dispatcher.Invoke(() =>
                        {
                            System.Windows.MessageBox.Show(
                                $"ERROR FATAL:\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                                "Import Gagal",
                                System.Windows.MessageBoxButton.OK,
                                System.Windows.MessageBoxImage.Error);
                        });
                    }
                });

                _viewModel.ShowBlankPage = true;
            }
            finally
            {
                _viewModel.IsBusy = false;
            }
        }

        private void ButtonMode2_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            _viewModel.ButtonMode2_Click();
        }

        private void ButtonMode3_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            _viewModel.ButtonMode3_Click();
        }

        private void ButtonMode4_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            _viewModel.ButtonMode4_Click();
        }

        private void ButtonBack_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            _viewModel.BackToMenu();
        }
    }
}