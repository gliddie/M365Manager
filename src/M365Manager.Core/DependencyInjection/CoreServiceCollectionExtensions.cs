using M365Manager.Core.Exchange;
using M365Manager.Core.M365;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;
using M365Manager.Core.Sql;
using M365Manager.Data.Logging;
using Microsoft.Extensions.DependencyInjection;

namespace M365Manager.Core.DependencyInjection;

public static class CoreServiceCollectionExtensions
{
    /// <summary>Registers all Core + Data services used by the app.</summary>
    public static IServiceCollection AddM365ManagerCore(this IServiceCollection services)
    {
        services.AddSingleton<ISettingsService, SettingsService>();
        services.AddSingleton<IConnectionStringProvider, SettingsConnectionStringProvider>();
        services.AddSingleton<ILogService, SqlLogService>();

        // Each service that signs in to an M365 workload is registered once as its own
        // concrete singleton, then exposed both under its specific interface and under
        // IM365Connector so IStartupConnectionService can connect it automatically at app
        // startup. Adding a future connector (Teams, SharePoint, ...) is just: implement
        // IM365Connector on the service and repeat this three-line pattern - no other
        // startup/UI code needs to change.
        services.AddSingleton<M365AuthService>();
        services.AddSingleton<IM365AuthService>(sp => sp.GetRequiredService<M365AuthService>());
        services.AddSingleton<IM365Connector>(sp => sp.GetRequiredService<M365AuthService>());

        // Embedded PowerShell + Exchange Online
        services.AddSingleton<PowerShellHost>();
        services.AddSingleton<ExchangeService>();
        services.AddSingleton<IExchangeService>(sp => sp.GetRequiredService<ExchangeService>());
        services.AddSingleton<IM365Connector>(sp => sp.GetRequiredService<ExchangeService>());

        services.AddSingleton<IStartupConnectionService, StartupConnectionService>();

        return services;
    }
}
