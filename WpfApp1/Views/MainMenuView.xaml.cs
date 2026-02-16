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
                        viewModel?.ImportWireToSQLite();
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