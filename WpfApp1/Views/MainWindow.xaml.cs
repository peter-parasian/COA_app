using System.Threading.Tasks;
using System.Windows;
using MahApps.Metro.Controls;

namespace WpfApp1.Views
{
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
                System.Windows.MessageBox.Show(
                    msg,
                    "Informasi Sistem",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
            };

            this.StateChanged += MainWindow_StateChanged;
        }

        private void MainWindow_StateChanged(object sender, System.EventArgs e)
        {
            if (this.WindowState == System.Windows.WindowState.Minimized)
            {
                System.GC.Collect(2, System.GCCollectionMode.Forced, false);
                System.GC.WaitForPendingFinalizers();
            }
        }

        private async void ButtonMode1_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (_viewModel.IsBusy) return;

            _viewModel.IsBusy = true;

            bool canCloseProgress = false;
            System.Windows.Window? progressWindow = null;

            try
            {
                progressWindow = new System.Windows.Window
                {
                    Title = "Processing...",
                    Width = 300,
                    Height = 150,
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner,
                    Owner = this,
                    ResizeMode = System.Windows.ResizeMode.NoResize,
                    WindowStyle = System.Windows.WindowStyle.None,
                    BorderBrush = System.Windows.Media.Brushes.Black,
                    BorderThickness = new System.Windows.Thickness(1),
                    Content = new System.Windows.Controls.StackPanel
                    {
                        Margin = new System.Windows.Thickness(20),
                        Children =
                        {
                            new System.Windows.Controls.TextBlock
                            {
                                Text = "Importing Excel to Database...",
                                FontSize = 14,
                                TextAlignment = System.Windows.TextAlignment.Center,
                                Margin = new System.Windows.Thickness(0, 10, 0, 20)
                            },
                            new System.Windows.Controls.ProgressBar
                            {
                                IsIndeterminate = true,
                                Height = 20
                            },
                            new System.Windows.Controls.TextBlock
                            {
                                Text = "Please wait, do not close this window...",
                                FontSize = 10,
                                TextAlignment = System.Windows.TextAlignment.Center,
                                Margin = new System.Windows.Thickness(0, 20, 0, 0),
                                Foreground = System.Windows.Media.Brushes.Gray
                            }
                        }
                    }
                };

                progressWindow.Closing += (s, args) =>
                {
                    if (!canCloseProgress)
                    {
                        args.Cancel = true;
                    }
                };

                progressWindow.Show();

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

                canCloseProgress = true;
                progressWindow.Close();

                this.Dispatcher.Invoke(() =>
                {
                    _viewModel.ShowBlankPage = true;
                });
            }
            catch (System.Exception ex)
            {
                if (progressWindow != null)
                {
                    canCloseProgress = true;
                    progressWindow.Close();
                }

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