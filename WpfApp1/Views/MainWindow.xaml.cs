using System.Windows;
using WpfApp1.ViewModels;

namespace WpfApp1.Views
{
    public partial class MainWindow : System.Windows.Window
    {
        private MainViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new MainViewModel();
        }

        private void ButtonMode1_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            ((System.Windows.Controls.Button)sender).IsEnabled = false;

            System.Threading.Tasks.Task.Run(() =>
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
                finally
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        ((System.Windows.Controls.Button)sender).IsEnabled = true;
                    });
                }
            });
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
    }
}