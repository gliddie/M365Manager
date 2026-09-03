using M365Manager.Core.ActiveDirectory;
using M365Manager.Core.Exchange;
using M365Manager.Core.Groups;
using M365Manager.Core.M365;
using M365Manager.Core.NewHire;
using M365Manager.Core.Notifications;
using M365Manager.Core.PowerShell;
using M365Manager.Core.RoomResources;
using M365Manager.Core.Settings;
using M365Manager.Core.SharedMailboxes;
using M365Manager.Core.SharePoint;
using M365Manager.Core.Sql;
using M365Manager.Core.Teams;
using M365Manager.Core.TeamsPolicies;
using M365Manager.Core.Trace;
using M365Manager.Data.Logging;
using M365Manager.Data.Naming;
using M365Manager.Data.NewHire;
using M365Manager.Data.Rooms;
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
        services.AddSingleton<GraphRestClient>();

        // Embedded PowerShell + Exchange Online. The transcript is what the console views on the
        // feature pages render; PowerShellHost feeds it centrally, so no service opts in.
        services.AddSingleton<IPowerShellTranscript, PowerShellTranscript>();
        services.AddSingleton<PowerShellHost>();
        services.AddSingleton<ExchangeService>();
        services.AddSingleton<IExchangeService>(sp => sp.GetRequiredService<ExchangeService>());
        services.AddSingleton<IM365Connector>(sp => sp.GetRequiredService<ExchangeService>());

        // SharePoint Online (PnP.PowerShell) - locks down site sharing on internal teams.
        services.AddSingleton<SharePointService>();
        services.AddSingleton<ISharePointService>(sp => sp.GetRequiredService<SharePointService>());
        services.AddSingleton<IM365Connector>(sp => sp.GetRequiredService<SharePointService>());

        // Microsoft Teams admin (MicrosoftTeams module) - Grant-Cs*Policy and phone number
        // assignment have no Graph equivalent. Same three-line connector pattern as the two above.
        services.AddSingleton<TeamsPowerShellService>();
        services.AddSingleton<ITeamsPowerShellService>(sp => sp.GetRequiredService<TeamsPowerShellService>());
        services.AddSingleton<IM365Connector>(sp => sp.GetRequiredService<TeamsPowerShellService>());

        services.AddSingleton<IStartupConnectionService, StartupConnectionService>();

        // Outbound notification mail. Prefers the configured SMTP relay and falls back to
        // Graph /me/sendMail, so features never talk to a transport directly.
        services.AddSingleton<INotificationMailService, NotificationMailService>();

        // Groups: distribution/security/dynamic group administration. Sibling of ExchangeService,
        // which keeps connection, search and membership.
        services.AddSingleton<IGroupNamingService, GroupNamingService>();
        services.AddSingleton<IGroupAdminService, GroupAdminService>();

        // Teams feature: naming convention + orchestration.
        services.AddSingleton<ITeamNamingRepository, SqlTeamNamingRepository>();
        services.AddSingleton<ITeamNamingService, TeamNamingService>();
        services.AddSingleton<ITeamsService, TeamsService>();

        // Shared Mailboxes feature: reuses the Teams naming acronym table (dbo.TeamNameAcronyms)
        // and the ExchangeService connection - no new IM365Connector needed.
        services.AddSingleton<ISharedMailboxNamingService, SharedMailboxNamingService>();
        services.AddSingleton<ISharedMailboxService, SharedMailboxService>();

        // Room & Resource mailboxes: also reuses the Teams naming acronym table and the
        // ExchangeService connection, plus its own editable site/time-zone list (dbo.RoomSites).
        services.AddSingleton<IRoomSiteRepository, SqlRoomSiteRepository>();
        services.AddSingleton<IRoomNamingService, RoomNamingService>();
        services.AddSingleton<IRoomResourceService, RoomResourceService>();

        // Teams policy groups: Entra group membership that drives Teams meeting policies.
        // Graph only, like Trace.
        services.AddSingleton<ITeamsPolicyService, TeamsPolicyService>();

        // Trace application: Entra guest accounts. Graph only - no Exchange connection involved.
        services.AddSingleton<ITraceGuestService, TraceGuestService>();

        // New Hire wizard: enables an employee for Teams telephony. The only feature that reads a
        // second database (the legacy telephony catalog on the same server, hence its own connection
        // string provider) and the only one that writes to on-premises AD.
        services.AddSingleton<INewHireConnectionStringProvider, NewHireConnectionStringProvider>();
        services.AddSingleton<INewHireRepository, SqlNewHireRepository>();
        services.AddSingleton<IActiveDirectoryService, ActiveDirectoryService>();
        services.AddSingleton<INewHireService, NewHireService>();

        return services;
    }
}
