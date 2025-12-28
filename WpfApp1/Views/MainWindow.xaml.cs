using System;
using System.Threading.Tasks;
using System.Windows;
using WpfApp1.ViewModels;

namespace WpfApp1.Views
{
    public partial class MainWindow : Window
    {
        private MainViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();

            _viewModel = new MainViewModel();

            this.DataContext = _viewModel;

            _viewModel.OnShowMessage += (msg) =>
            {
                MessageBox.Show(
                    msg,
                    "Informasi Sistem",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            };
        }

        private async void ButtonMode1_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel.IsBusy) return;

            try
            {
                _viewModel.IsBusy = true;

                await Task.Run(() =>
                {
                    try
                    {
                        _viewModel.ImportExcelToSQLite();
                    }
                    catch (Exception ex)
                    {
                        this.Dispatcher.Invoke(() =>
                        {
                            MessageBox.Show(
                                $"ERROR FATAL:\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                                "Import Gagal",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
                        });
                    }
                });

                // Logika yang diminta: Setelah import selesai (dan popup ditutup), ganti ke halaman kosong
                _viewModel.ShowBlankPage = true;
            }
            finally
            {
                _viewModel.IsBusy = false;
            }
        }

        private void ButtonMode2_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.ButtonMode2_Click();
        }

        private void ButtonMode3_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.ButtonMode3_Click();
        }

        private void ButtonMode4_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.ButtonMode4_Click();
        }

        private void ButtonBack_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.BackToMenu();
        }
    }
}