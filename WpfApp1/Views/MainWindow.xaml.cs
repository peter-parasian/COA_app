namespace WpfApp1.Views
{
    public partial class MainWindow : MahApps.Metro.Controls.MetroWindow
    {
        private WpfApp1.ViewModels.MainViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new WpfApp1.ViewModels.MainViewModel();
            this.DataContext = _viewModel;
            _viewModel.OnShowMessage += (msg) => { System.Windows.MessageBox.Show(msg, "Informasi Sistem", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information); };
            this.StateChanged += MainWindow_StateChanged;
        }

        private void MainWindow_StateChanged(System.Object sender, System.EventArgs e)
        {
            if (this.WindowState == System.Windows.WindowState.Minimized)
            {
                System.GC.Collect(2, System.GCCollectionMode.Forced, false);
                System.GC.WaitForPendingFinalizers();
            }
        }

        private async void ButtonMode1_Click(System.Object sender, System.Windows.RoutedEventArgs e)
        {
            if (_viewModel.IsBusy)
            {
                return;
            }

            _viewModel.BusyMessage = "Importing Excel to Database...";
            _viewModel.IsBusy = true;

            try
            {
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
                        throw;
                    }
                });

                this.Dispatcher.Invoke(() => { _viewModel.ShowBlankPage = true; });
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"Error: {ex.Message}",
                    "Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
            finally
            {
                _viewModel.IsBusy = false;
            }
        }

        private void ButtonMode2_Click(System.Object sender, System.Windows.RoutedEventArgs e) { _viewModel.ButtonMode2_Click(); }
        private void ButtonMode3_Click(System.Object sender, System.Windows.RoutedEventArgs e) { _viewModel.ButtonMode3_Click(); }
        private void ButtonMode4_Click(System.Object sender, System.Windows.RoutedEventArgs e) { _viewModel.ButtonMode4_Click(); }
        private void ButtonBack_Click(System.Object sender, System.Windows.RoutedEventArgs e) { _viewModel.BackToMenu(); }

        protected override void OnClosed(System.EventArgs e)
        {
            base.OnClosed(e);
            _viewModel = null;
            this.DataContext = null;
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }
    }
}