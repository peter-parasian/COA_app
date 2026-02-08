using System.Globalization;
using System.Threading;
using System.Windows;
using System.Windows.Markup;

namespace WpfApp1
{
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var indoCulture = new CultureInfo("id-ID");

            Thread.CurrentThread.CurrentCulture = indoCulture;
            Thread.CurrentThread.CurrentUICulture = indoCulture;

            FrameworkElement.LanguageProperty.OverrideMetadata(
                typeof(FrameworkElement),
                new FrameworkPropertyMetadata(
                    XmlLanguage.GetLanguage(indoCulture.IetfLanguageTag)));
        }
    }
}