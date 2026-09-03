using System.Collections.ObjectModel;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Core.M365;
using M365Manager.Core.Notifications;
using M365Manager.Core.Settings;
using M365Manager.Data.Logging;
using M365Manager.Data.Naming;
using M365Manager.Data.NewHire;
using M365Manager.Data.Rooms;

namespace M365Manager.ViewModels;

public sealed partial class SettingsViewModel : ObservableObject
{
    private readonly ISettingsService _settings;
    private readonly ILogService _log;
    private readonly IM365AuthService _auth;
    private readonly ITeamNamingRepository _naming;
    private readonly IRoomSiteRepository _roomSites;
    private readonly INotificationMailService _mail;
    private readonly INewHireRepository _newHire;

    // --- SQL ---
    [ObservableProperty] private string _sqlServer = "";
    [ObservableProperty] private string _sqlDatabase = "M365Manager";
    [ObservableProperty] private bool _sqlUseIntegratedSecurity;
    [ObservableProperty] private string _sqlUserId = "";
    [ObservableProperty] private string _sqlPassword = "";

    // --- M365 ---
    [ObservableProperty] private string _tenantId = "";
    [ObservableProperty] private string _clientId = "";
    [ObservableProperty] private string _domain = "";
    [ObservableProperty] private string _defaultMailDomain = "";
    [ObservableProperty] private string _sharePointAdminUrl = "";

    // --- SMTP (notification e-mail) ---
    [ObservableProperty] private string _smtpHost = "";
    [ObservableProperty] private string _smtpPort = "25";
    [ObservableProperty] private bool _smtpUseSsl;
    [ObservableProperty] private string _smtpFromAddress = "";
    [ObservableProperty] private string _smtpFromDisplayName = "";
    [ObservableProperty] private string _smtpReplyToAddress = "";
    [ObservableProperty] private string _smtpBccAddress = "";
    [ObservableProperty] private string _smtpUserName = "";
    [ObservableProperty] private string _smtpPassword = "";
    [ObservableProperty] private string _smtpTestRecipient = "";

    // --- Shared Mailboxes ---
    [ObservableProperty] private string _sharedMailboxRetentionPolicyName = "";
    [ObservableProperty] private string _sharedMailboxRoleAssignmentPolicyName = "";
    [ObservableProperty] private string _sharedMailboxGroupMembershipHelpUrl = "";
    [ObservableProperty] private string _sharedMailboxHelpLinks = "";

    // --- Room & resource mailboxes ---
    [ObservableProperty] private string _roomMailboxDomain = "";
    [ObservableProperty] private string _roomGroupDomain = "";
    [ObservableProperty] private string _roomListOwnerGroup = "";
    [ObservableProperty] private string _roomGlobalAdminGroup = "";
    [ObservableProperty] private string _roomBookingWindowDays = "90";
    [ObservableProperty] private string _roomMaximumDurationMinutes = "1440";

    // --- Teams policies ---
    [ObservableProperty] private string _teamsPolicyGroupPrefixes = "";

    // --- New Hire (telephony) ---
    [ObservableProperty] private string _newHireDatabase = NewHireSettings.DefaultDatabase;
    [ObservableProperty] private bool _newHireWriteAdAttributes = true;
    [ObservableProperty] private string _newHireLdapServer = "";
    [ObservableProperty] private string _newHireAdUserName = "";
    [ObservableProperty] private string _newHireAdPassword = "";
    [ObservableProperty] private string _newHireDeploymentLocator = "";
    [ObservableProperty] private string _newHirePhoneNumberType = "";
    [ObservableProperty] private string _newHireTeamsUpgradePolicyName = "";
    [ObservableProperty] private bool _newHireNotifyEmployee = true;
    [ObservableProperty] private string _newHireAdminNotificationAddress = "";
    [ObservableProperty] private string _newHireStatusMessage = "";

    // --- Trace application ---
    [ObservableProperty] private string _traceCompanyName = "";
    [ObservableProperty] private string _traceRedirectUrl = "";
    [ObservableProperty] private string _traceInvitationMessage = "";

    // --- Room sites (site code -> time zone + regional admin group) ---
    [ObservableProperty] private RoomSite? _selectedRoomSite;
    [ObservableProperty] private string _newSiteCode = "";
    [ObservableProperty] private string _newSiteTimeZone = "";
    [ObservableProperty] private string _newSiteRegion = "";
    [ObservableProperty] private string _newSiteAdminGroup = "";
    [ObservableProperty] private string _roomSiteStatusMessage = "";

    public ObservableCollection<RoomSite> RoomSites { get; } = new();

    // --- Team naming acronyms ---
    [ObservableProperty] private TeamNameAcronym? _selectedAcronym;
    [ObservableProperty] private string _newAcronym = "";
    [ObservableProperty] private string _newTranslation = "";
    [ObservableProperty] private string _acronymStatusMessage = "";

    public ObservableCollection<TeamNameAcronym> Acronyms { get; } = new();

    // --- UI state ---
    [ObservableProperty] private string _statusMessage = "";
    [ObservableProperty] private bool _isBusy;

    public string SettingsFilePath => _settings.SettingsFilePath;

    public SettingsViewModel(ISettingsService settings, ILogService log, IM365AuthService auth, ITeamNamingRepository naming, IRoomSiteRepository roomSites, INotificationMailService mail, INewHireRepository newHire)
    {
        _settings = settings;
        _log = log;
        _auth = auth;
        _naming = naming;
        _roomSites = roomSites;
        _mail = mail;
        _newHire = newHire;

        var s = _settings.Current;
        SqlServer = s.Sql.Server;
        SqlDatabase = s.Sql.Database;
        SqlUseIntegratedSecurity = s.Sql.UseIntegratedSecurity;
        SqlUserId = s.Sql.UserId;
        SqlPassword = s.Sql.Password;
        TenantId = s.M365.TenantId;
        ClientId = s.M365.ClientId;
        Domain = s.M365.Domain;
        DefaultMailDomain = s.M365.DefaultMailDomain;
        SharePointAdminUrl = s.M365.SharePointAdminUrl;
        SmtpHost = s.Smtp.Host;
        SmtpPort = s.Smtp.Port.ToString();
        SmtpUseSsl = s.Smtp.UseSsl;
        SmtpFromAddress = s.Smtp.FromAddress;
        SmtpFromDisplayName = s.Smtp.FromDisplayName;
        SmtpReplyToAddress = s.Smtp.ReplyToAddress;
        SmtpBccAddress = s.Smtp.BccAddress;
        SmtpUserName = s.Smtp.UserName;
        SmtpPassword = s.Smtp.Password;
        SmtpTestRecipient = auth.CurrentUser?.Upn ?? "";
        SharedMailboxRetentionPolicyName = s.SharedMailboxes.RetentionPolicyName;
        SharedMailboxRoleAssignmentPolicyName = s.SharedMailboxes.RoleAssignmentPolicyName;
        SharedMailboxGroupMembershipHelpUrl = s.SharedMailboxes.GroupMembershipHelpUrl;
        SharedMailboxHelpLinks = s.SharedMailboxes.MailboxHelpLinks;
        RoomMailboxDomain = s.RoomResources.RoomMailboxDomain;
        RoomGroupDomain = s.RoomResources.RoomGroupDomain;
        RoomListOwnerGroup = s.RoomResources.RoomListOwnerGroup;
        RoomGlobalAdminGroup = s.RoomResources.GlobalAdminGroup;
        RoomBookingWindowDays = s.RoomResources.DefaultBookingWindowDays.ToString();
        RoomMaximumDurationMinutes = s.RoomResources.DefaultMaximumDurationMinutes.ToString();
        TeamsPolicyGroupPrefixes = s.TeamsPolicies.GroupPrefixes;
        NewHireDatabase = s.NewHire.Database;
        NewHireWriteAdAttributes = s.NewHire.WriteAdAttributes;
        NewHireLdapServer = s.NewHire.LdapServer;
        NewHireAdUserName = s.NewHire.AdUserName;
        NewHireAdPassword = s.NewHire.AdPassword;
        NewHireDeploymentLocator = s.NewHire.DeploymentLocator;
        NewHirePhoneNumberType = s.NewHire.PhoneNumberType;
        NewHireTeamsUpgradePolicyName = s.NewHire.TeamsUpgradePolicyName;
        NewHireNotifyEmployee = s.NewHire.NotifyEmployee;
        NewHireAdminNotificationAddress = s.NewHire.AdminNotificationAddress;
        TraceCompanyName = s.Trace.CompanyName;
        TraceRedirectUrl = s.Trace.RedirectUrl;
        TraceInvitationMessage = s.Trace.InvitationMessage;

        _ = LoadAcronymsAsync();
        _ = LoadRoomSitesAsync();
    }

    private async Task LoadRoomSitesAsync()
    {
        try
        {
            var all = await _roomSites.GetAllAsync();
            RoomSites.Clear();
            foreach (var site in all)
                RoomSites.Add(site);
            RoomSiteStatusMessage = $"{RoomSites.Count} site(s) loaded.";
        }
        catch (Exception ex)
        {
            RoomSiteStatusMessage = $"Could not load sites: {ErrorText.Describe(ex)}";
        }
    }

    /// <summary>
    /// Imports the legacy RoomTimeZones.csv (columns Code, Zone, Admins) that used to live on the
    /// automation server. Re-importing a corrected file updates existing rows instead of
    /// duplicating them.
    /// </summary>
    [RelayCommand]
    private async Task ImportRoomSitesAsync()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "CSV file (*.csv;*.txt)|*.csv;*.txt|All files (*.*)|*.*",
            Title = "Select RoomTimeZones.csv",
        };
        if (dialog.ShowDialog() != true)
            return;

        try
        {
            var sites = RoomSiteCsv.Parse(await File.ReadAllLinesAsync(dialog.FileName));
            if (sites.Count == 0)
            {
                RoomSiteStatusMessage = "No usable rows found - expected columns 'Code', 'Zone' and 'Admins'.";
                return;
            }

            var (added, updated) = await _roomSites.UpsertManyAsync(sites);
            await LoadRoomSitesAsync();
            RoomSiteStatusMessage = $"Imported {sites.Count} row(s) from {Path.GetFileName(dialog.FileName)}: {added} added, {updated} updated.";
        }
        catch (Exception ex)
        {
            RoomSiteStatusMessage = $"Import failed: {ErrorText.Describe(ex)}";
        }
    }

    [RelayCommand]
    private async Task AddRoomSiteAsync()
    {
        if (string.IsNullOrWhiteSpace(NewSiteCode) || string.IsNullOrWhiteSpace(NewSiteTimeZone))
        {
            RoomSiteStatusMessage = "Enter at least a site code and a time zone.";
            return;
        }

        try
        {
            await _roomSites.AddAsync(new RoomSite
            {
                SiteCode = NewSiteCode.Trim().ToUpperInvariant(),
                TimeZone = NewSiteTimeZone.Trim(),
                Region = NewSiteRegion.Trim(),
                RegionalAdminGroup = NewSiteAdminGroup.Trim(),
            });
            NewSiteCode = "";
            NewSiteTimeZone = "";
            NewSiteRegion = "";
            NewSiteAdminGroup = "";
            await LoadRoomSitesAsync();
        }
        catch (Exception ex)
        {
            RoomSiteStatusMessage = $"Add failed: {ErrorText.Describe(ex)}";
        }
    }

    [RelayCommand]
    private async Task DeleteRoomSiteAsync()
    {
        if (SelectedRoomSite is null)
        {
            RoomSiteStatusMessage = "Select a site to delete.";
            return;
        }

        try
        {
            await _roomSites.DeleteAsync(SelectedRoomSite.Id);
            await LoadRoomSitesAsync();
        }
        catch (Exception ex)
        {
            RoomSiteStatusMessage = $"Delete failed: {ErrorText.Describe(ex)}";
        }
    }

    [RelayCommand]
    private async Task SaveRoomSiteChangesAsync()
    {
        try
        {
            foreach (var site in RoomSites)
                await _roomSites.UpdateAsync(site);
            RoomSiteStatusMessage = "Changes saved.";
        }
        catch (Exception ex)
        {
            RoomSiteStatusMessage = $"Save failed: {ErrorText.Describe(ex)}";
        }
    }

    private async Task LoadAcronymsAsync()
    {
        try
        {
            var all = await _naming.GetAllAsync();
            Acronyms.Clear();
            foreach (var a in all)
                Acronyms.Add(a);
            AcronymStatusMessage = $"{Acronyms.Count} entries loaded.";
        }
        catch (Exception ex)
        {
            AcronymStatusMessage = $"Could not load acronyms: {ErrorText.Describe(ex)}";
        }
    }

    [RelayCommand]
    private async Task AddAcronymAsync()
    {
        if (string.IsNullOrWhiteSpace(NewAcronym) || string.IsNullOrWhiteSpace(NewTranslation))
        {
            AcronymStatusMessage = "Enter both an acronym and its translation.";
            return;
        }

        try
        {
            await _naming.AddAsync(new TeamNameAcronym { Acronym = NewAcronym, Translation = NewTranslation });
            NewAcronym = "";
            NewTranslation = "";
            await LoadAcronymsAsync();
        }
        catch (Exception ex)
        {
            AcronymStatusMessage = $"Add failed: {ErrorText.Describe(ex)}";
        }
    }

    [RelayCommand]
    private async Task DeleteAcronymAsync()
    {
        if (SelectedAcronym is null)
        {
            AcronymStatusMessage = "Select an entry to delete.";
            return;
        }

        try
        {
            await _naming.DeleteAsync(SelectedAcronym.Id);
            await LoadAcronymsAsync();
        }
        catch (Exception ex)
        {
            AcronymStatusMessage = $"Delete failed: {ErrorText.Describe(ex)}";
        }
    }

    [RelayCommand]
    private async Task SaveAcronymChangesAsync()
    {
        try
        {
            foreach (var a in Acronyms)
                await _naming.UpdateAsync(a);
            AcronymStatusMessage = "Changes saved.";
        }
        catch (Exception ex)
        {
            AcronymStatusMessage = $"Save failed: {ErrorText.Describe(ex)}";
        }
    }

    private void ApplyToSettings()
    {
        var s = _settings.Current;
        s.Sql.Server = SqlServer.Trim();
        s.Sql.Database = SqlDatabase.Trim();
        s.Sql.UseIntegratedSecurity = SqlUseIntegratedSecurity;
        s.Sql.UserId = SqlUserId.Trim();
        s.Sql.Password = SqlPassword;
        s.M365.TenantId = TenantId.Trim();
        s.M365.ClientId = ClientId.Trim();
        s.M365.Domain = Domain.Trim();
        s.M365.DefaultMailDomain = DefaultMailDomain.Trim();
        s.M365.SharePointAdminUrl = SharePointAdminUrl.Trim();
        s.Smtp.Host = SmtpHost.Trim();
        s.Smtp.Port = int.TryParse(SmtpPort.Trim(), out var port) && port > 0 ? port : 25;
        s.Smtp.UseSsl = SmtpUseSsl;
        s.Smtp.FromAddress = SmtpFromAddress.Trim();
        s.Smtp.FromDisplayName = SmtpFromDisplayName.Trim();
        s.Smtp.ReplyToAddress = SmtpReplyToAddress.Trim();
        s.Smtp.BccAddress = SmtpBccAddress.Trim();
        s.Smtp.UserName = SmtpUserName.Trim();
        s.Smtp.Password = SmtpPassword;
        s.SharedMailboxes.RetentionPolicyName = SharedMailboxRetentionPolicyName.Trim();
        s.SharedMailboxes.RoleAssignmentPolicyName = SharedMailboxRoleAssignmentPolicyName.Trim();
        s.SharedMailboxes.GroupMembershipHelpUrl = SharedMailboxGroupMembershipHelpUrl.Trim();
        s.SharedMailboxes.MailboxHelpLinks = SharedMailboxHelpLinks;
        s.RoomResources.RoomMailboxDomain = RoomMailboxDomain.Trim();
        s.RoomResources.RoomGroupDomain = RoomGroupDomain.Trim();
        s.RoomResources.RoomListOwnerGroup = RoomListOwnerGroup.Trim();
        s.RoomResources.GlobalAdminGroup = RoomGlobalAdminGroup.Trim();
        s.RoomResources.DefaultBookingWindowDays = int.TryParse(RoomBookingWindowDays.Trim(), out var window) && window > 0 ? window : 90;
        s.RoomResources.DefaultMaximumDurationMinutes = int.TryParse(RoomMaximumDurationMinutes.Trim(), out var duration) && duration > 0 ? duration : 1440;
        s.TeamsPolicies.GroupPrefixes = TeamsPolicyGroupPrefixes;
        s.NewHire.Database = NewHireDatabase.Trim();
        s.NewHire.WriteAdAttributes = NewHireWriteAdAttributes;
        s.NewHire.LdapServer = NewHireLdapServer.Trim();
        s.NewHire.AdUserName = NewHireAdUserName.Trim();
        s.NewHire.AdPassword = NewHireAdPassword;
        s.NewHire.DeploymentLocator = NewHireDeploymentLocator.Trim();
        s.NewHire.PhoneNumberType = NewHirePhoneNumberType.Trim();
        s.NewHire.TeamsUpgradePolicyName = NewHireTeamsUpgradePolicyName.Trim();
        s.NewHire.NotifyEmployee = NewHireNotifyEmployee;
        s.NewHire.AdminNotificationAddress = NewHireAdminNotificationAddress.Trim();
        s.Trace.CompanyName = TraceCompanyName.Trim();
        s.Trace.RedirectUrl = TraceRedirectUrl.Trim();
        s.Trace.InvitationMessage = TraceInvitationMessage;
        _settings.Save(s);
    }

    [RelayCommand]
    private void Save()
    {
        ApplyToSettings();
        StatusMessage = $"Settings saved to {_settings.SettingsFilePath}";
    }

    /// <summary>Puts the invitation text back to the wording the legacy invite script used.</summary>
    [RelayCommand]
    private void ResetTraceInvitationMessage()
    {
        TraceInvitationMessage = TraceSettings.DefaultInvitationMessage;
        StatusMessage = "Invitation text reset to the default - click Save to keep it.";
    }

    [RelayCommand]
    private async Task TestSqlAsync()
    {
        IsBusy = true;
        StatusMessage = "Testing SQL connection...";
        try
        {
            ApplyToSettings();
            var ok = await _log.TestConnectionAsync();
            if (ok)
            {
                await _log.WriteAsync(new LogEntry
                {
                    Area = "System",
                    Action = "ConnectionTest",
                    EventCode = "STAR",
                    Severity = Severity.Success,
                    TargetObject = SqlServer,
                    UserUpn = _auth.CurrentUser?.Upn,
                    Message = "SQL connection test succeeded from M365Manager.",
                });
                StatusMessage = "SQL connection OK - a test entry was written to the log.";
            }
            else
            {
                StatusMessage = "Could not connect. Check server, database and credentials.";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"SQL error: {ErrorText.Describe(ex)}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    /// The telephony database sits on the same server as the app's own, so only the catalog name can
    /// be wrong here - and a wrong one is worth catching in Settings rather than three steps into a
    /// new hire.
    /// </summary>
    [RelayCommand]
    private async Task TestNewHireSqlAsync()
    {
        IsBusy = true;
        NewHireStatusMessage = "Testing the telephony database connection...";
        try
        {
            ApplyToSettings();
            NewHireStatusMessage = await _newHire.TestConnectionAsync()
                ? $"Connected to '{NewHireDatabase.Trim()}' and read the site table."
                : $"Could not connect to '{NewHireDatabase.Trim()}' on {SqlServer.Trim()}. Check the database name.";
        }
        catch (Exception ex)
        {
            NewHireStatusMessage = $"SQL error: {ErrorText.Describe(ex)}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    /// Sends a real notification through whichever transport is configured, so a wrong host, port,
    /// TLS setting or blocked relay surfaces here instead of silently in a WARN log line three
    /// mailbox creations later.
    /// </summary>
    [RelayCommand]
    private async Task SendTestMailAsync()
    {
        var recipient = SmtpTestRecipient.Trim();
        if (recipient.Length == 0)
        {
            StatusMessage = "Enter an address to send the test message to.";
            return;
        }

        IsBusy = true;
        StatusMessage = "Sending test message...";
        try
        {
            ApplyToSettings();
            var html = $"""
                <p>Test message from M365Manager.</p>
                <p>Sent {DateTime.Now:yyyy-MM-dd HH:mm} via {(_mail.IsSmtpConfigured ? "the configured SMTP relay" : "Microsoft Graph as the signed-in account")}.
                If you can read this, notification e-mails for new shared mailboxes and teams will go out the same way.</p>
                """;

            await _mail.SendAsync(new[] { recipient }, "M365Manager - SMTP test", html);
            StatusMessage = $"Test message sent to {recipient} from {_mail.SenderDescription}. Check the inbox (and the junk folder).";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Sending failed: {ErrorText.Describe(ex)}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task SignInM365Async()
    {
        IsBusy = true;
        StatusMessage = "Opening sign-in...";
        try
        {
            ApplyToSettings();
            var user = await _auth.SignInAsync();
            StatusMessage = $"Signed in as {user.DisplayName} ({user.Upn}).";

            try
            {
                await _log.WriteAsync(new LogEntry
                {
                    Area = "System",
                    Action = "SignIn",
                    EventCode = "INFO",
                    Severity = Severity.Info,
                    UserUpn = user.Upn,
                    Message = "Interactive M365 sign-in succeeded.",
                });
            }
            catch
            {
                // Logging is best-effort here; SQL may not be configured yet.
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Sign-in failed: {ErrorText.Describe(ex)}";
        }
        finally
        {
            IsBusy = false;
        }
    }
}
