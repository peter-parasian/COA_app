using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using WpfApp1.ViewModels;

namespace WpfApp1.Views
{
    public partial class MainMenuView : UserControl
    {
        public MainMenuView()
        {
            InitializeComponent();
        }

        private async void ButtonMode1_Click(object sender, RoutedEventArgs e)
        {
            MainViewModel? viewModel = this.DataContext as MainViewModel;

            if (viewModel == null || viewModel.IsBusy)
            {
                return;
            }

            viewModel.BusyMessage = "Importing Excel to Database...";
            viewModel.IsBusy = true;

            try
            {
                await Task.Run(() =>
                {
                    try
                    {
                        viewModel?.ImportExcelToSQLite();
                    }
                    catch (System.Exception ex)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
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

                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (viewModel != null)
                    {
                        viewModel.ShowBlankPage = true;
                    }
                });
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
                if (viewModel != null)
                {
                    viewModel.IsBusy = false;
                }
            }
        }

        private async void ButtonMode2_Click(object sender, RoutedEventArgs e)
        {
            MainViewModel? viewModel = this.DataContext as MainViewModel;

            if (viewModel == null || viewModel.IsBusy)
            {
                return;
            }

            viewModel.BusyMessage = "Importing Excel to Database...";
            viewModel.IsBusy = true;

            try
            {
                await Task.Run(() =>
                {
                    try
                    {
                        viewModel?.ImportExcelToSQLite();
                    }
                    catch (System.Exception ex)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
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

                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (viewModel != null)
                    {
                        viewModel.ShowMode2Page = true;
                    }
                });
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
                if (viewModel != null)
                {
                    viewModel.IsBusy = false;
                }
            }
        }

        private async void ButtonMode3_Click(object sender, RoutedEventArgs e)
        {
            // Cast DataContext to specific type explicitly (no var)
            MainViewModel? viewModel = this.DataContext as MainViewModel;

            if (viewModel == null || viewModel.IsBusy)
            {
                return;
            }

            viewModel.BusyMessage = "Importing Wire Excel to Database...";
            viewModel.IsBusy = true;

            try
            {
                await Task.Run(() =>
                {
                    try
                    {
                        // CHANGED: Call ImportWireToSQLite for Mode 3 instead of generic Import
                        viewModel?.ImportWireToSQLite();
                    }
                    catch (System.Exception ex)
                    {
                        // Use Application.Current.Dispatcher for UI access from background thread
                        Application.Current.Dispatcher.Invoke(() =>
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

                // Switch to Mode 3 view via property change
                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (viewModel != null)
                    {
                        viewModel.ShowMode3Page = true;
                    }
                });
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
                if (viewModel != null)
                {
                    viewModel.IsBusy = false;
                }
            }
        }

        private void ButtonMode4_Click(object sender, RoutedEventArgs e)
        {
            MainViewModel? viewModel = this.DataContext as MainViewModel;
            if (viewModel != null)
            {
                viewModel.ButtonMode4_Click();
            }
        }
    }
}