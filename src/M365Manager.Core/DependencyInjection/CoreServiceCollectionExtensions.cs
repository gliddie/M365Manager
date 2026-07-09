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
        services.AddSingleton<IM365AuthService, M365AuthService>();

        // Embedded PowerShell + Exchange Online
        services.AddSingleton<PowerShellHost>();
        services.AddSingleton<IExchangeService, ExchangeService>();

        return services;
    }
}
