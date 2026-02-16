using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using MahApps.Metro.Controls;

namespace WpfApp1.Views
{
    public partial class MainWindow : MetroWindow
    {
        private WpfApp1.ViewModels.MainViewModel? _viewModel;

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

            SetMainContent(new MainMenuView());

            if (_viewModel != null)
            {
                _viewModel.PropertyChanged += ViewModel_PropertyChanged;
            }
        }

        private void ViewModel_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(WpfApp1.ViewModels.MainViewModel.ShowBlankPage))
            {
                if (_viewModel != null)
                {
                    if (_viewModel.ShowBlankPage)
                    {
                        SetMainContent(new Mode1View());
                    }
                    else if (!_viewModel.ShowMode2Page && !_viewModel.ShowMode3Page)
                    {
                        SetMainContent(new MainMenuView());
                    }
                }
            }
            else if (e.PropertyName == nameof(WpfApp1.ViewModels.MainViewModel.ShowMode2Page))
            {
                if (_viewModel != null)
                {
                    if (_viewModel.ShowMode2Page)
                    {
                        SetMainContent(new Mode2View());
                    }
                    else if (!_viewModel.ShowBlankPage && !_viewModel.ShowMode3Page)
                    {
                        SetMainContent(new MainMenuView());
                    }
                }
            }
            else if (e.PropertyName == nameof(WpfApp1.ViewModels.MainViewModel.ShowMode3Page))
            {
                if (_viewModel != null)
                {
                    if (_viewModel.ShowMode3Page)
                    {
                        SetMainContent(new Mode3View());
                    }
                    else if (!_viewModel.ShowBlankPage && !_viewModel.ShowMode2Page)
                    {
                        SetMainContent(new MainMenuView());
                    }
                }
            }
        }

        private void SetMainContent(UserControl viewControl)
        {
            MainContentRegion.Content = viewControl;
        }

        private void MainWindow_StateChanged(object? sender, System.EventArgs e)
        {
            if (this.WindowState == System.Windows.WindowState.Minimized)
            {
                System.GC.Collect(2, System.GCCollectionMode.Optimized, false);
                System.GC.WaitForPendingFinalizers();
            }
        }

        protected override void OnClosed(System.EventArgs e)
        {
            if (_viewModel != null)
            {
                _viewModel.PropertyChanged -= ViewModel_PropertyChanged;
            }

            base.OnClosed(e);

            _viewModel = null;
            this.DataContext = null;

            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }
    }
}