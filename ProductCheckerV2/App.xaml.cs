using System.Windows;
using ProductCheckerV2.Common;

namespace ProductCheckerV2
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Handle unhandled exceptions
            DispatcherUnhandledException += App_DispatcherUnhandledException;
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            ModalDialogService.Show(
                $"An unexpected error occurred:\n\n{e.Exception.Message}",
                "Application Error",
                ModalDialogType.Error,
                Current?.MainWindow);

            e.Handled = true;
        }
    }
}
