using System.Windows;
using System.Windows.Controls;
using WpfApp1.ViewModels;

namespace WpfApp1.Views
{
    public partial class Mode3View : UserControl
    {
        public Mode3View()
        {
            InitializeComponent();
        }

        private void ButtonBack_Click(object sender, RoutedEventArgs e)
        {
            MainViewModel? viewModel = this.DataContext as MainViewModel;

            if (viewModel != null)
            {
                viewModel.BackToMenu();
            }
        }
    }
}