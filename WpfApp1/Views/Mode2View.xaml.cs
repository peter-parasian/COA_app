using System.Windows;
using System.Windows.Controls;
using WpfApp1.ViewModels;

namespace WpfApp1.Views
{
    // Fix: Class name must be Mode2View to match x:Class in XAML
    public partial class Mode2View : UserControl
    {
        public Mode2View()
        {
            InitializeComponent();
        }

        private void ButtonBack_Click(object sender, RoutedEventArgs e)
        {
            // Explicit cast
            MainViewModel? viewModel = this.DataContext as MainViewModel;

            if (viewModel != null)
            {
                viewModel.BackToMenu();
            }
        }
    }
}