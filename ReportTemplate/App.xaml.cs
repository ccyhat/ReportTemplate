using System.Windows;

namespace ReportTemplate;

public partial class App : Application
{
    private readonly AutofacBootstrapper _bootstrapper;

    public App()
    {
        _bootstrapper = new AutofacBootstrapper();
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
    }
}
