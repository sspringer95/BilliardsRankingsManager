using MaterialDesignThemes.Wpf;
using System.Windows;

namespace BilliardsRankingsManager
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {

            base.OnStartup(e);

            // 1. Create the helper
            var paletteHelper = new MaterialDesignThemes.Wpf.PaletteHelper();

            // 2. Get the current theme (which is currently empty/default)
            var theme = paletteHelper.GetTheme();

            // 3. Set your "Default" colors (This fixes the invisible underlines globally)
            // You can use any Hex value here.
            theme.SetPrimaryColor(System.Windows.Media.Color.FromRgb(154, 189, 217)); // Example: light Blue
            theme.SetSecondaryColor(System.Windows.Media.Color.FromRgb(23, 133, 24)); // Example: green

            // 4. Apply it
            paletteHelper.SetTheme(theme);

            var mainWindow = new MainWindow();
            mainWindow.Show();
        }
    }
}