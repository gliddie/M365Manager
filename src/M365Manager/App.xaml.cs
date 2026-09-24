using System.Windows;
using M365Manager.Core.DependencyInjection;
using M365Manager.Services;
using M365Manager.ViewModels;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace M365Manager;

public partial class App : Application
{
    private readonly IHost _host;

    public App()
    {
        _host = Host.CreateDefaultBuilder()
            .ConfigureServices((_, services) =>
            {
                services.AddM365ManagerCore();

                // Shell services
                services.AddSingleton<IThemeService, ThemeService>();
                services.AddSingleton<IDialogService, DialogService>();
                services.AddSingleton<INavigationService, NavigationService>();

                // ViewModels
                services.AddSingleton<MainViewModel>();
                services.AddSingleton<DashboardViewModel>();
                services.AddSingleton<GroupsViewModel>();
                services.AddSingleton<TeamsViewModel>();
                services.AddSingleton<SharedMailboxesViewModel>();
                services.AddSingleton<RoomResourcesViewModel>();
                services.AddSingleton<TeamsPoliciesViewModel>();
                services.AddSingleton<NewHireViewModel>();
                services.AddSingleton<TraceViewModel>();
                services.AddSingleton<SettingsViewModel>();
                services.AddSingleton<LogsViewModel>();

                // Shell
                services.AddSingleton<MainWindow>();
            })
            .Build();
    }

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        await _host.StartAsync();

        // Before the window is shown, so it never paints in the wrong theme first.
        _host.Services.GetRequiredService<IThemeService>().Initialize();

        var window = _host.Services.GetRequiredService<MainWindow>();
        var mainViewModel = _host.Services.GetRequiredService<MainViewModel>();
        window.DataContext = mainViewModel;
        window.Show();

        // Sign in to Graph, connect Exchange, and any future M365 connector - right away,
        // so nobody has to open Settings then Groups just to connect (see MainViewModel.ConnectAllAsync).
        _ = mainViewModel.ConnectAllAsync();
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        await _host.StopAsync();
        _host.Dispose();
        base.OnExit(e);
    }
}
