using System.IO;
using System.Text.Json;
using M365Manager.Core.Security;

namespace M365Manager.Core.Settings;

/// <summary>
/// Loads/saves <see cref="AppSettings"/> as JSON in %APPDATA%\M365Manager\settings.json.
/// Secret fields are encrypted with DPAPI before being written and decrypted on load.
/// </summary>
public sealed class SettingsService : ISettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public string SettingsFilePath { get; }

    public AppSettings Current { get; private set; } = new();

    public SettingsService()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "M365Manager");
        Directory.CreateDirectory(dir);
        SettingsFilePath = Path.Combine(dir, "settings.json");
        Load();
    }

    public AppSettings Load()
    {
        if (!File.Exists(SettingsFilePath))
        {
            Current = new AppSettings();
            return Current;
        }

        try
        {
            var json = File.ReadAllText(SettingsFilePath);
            var settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();

            // Decrypt secrets that were encrypted at rest.
            settings.Sql.Password = SecretProtector.Unprotect(settings.Sql.Password) ?? "";
            settings.Smtp.Password = SecretProtector.Unprotect(settings.Smtp.Password) ?? "";

            Current = settings;
        }
        catch
        {
            Current = new AppSettings();
        }

        return Current;
    }

    public void Save(AppSettings settings)
    {
        // Build a copy with encrypted secrets so we never write plaintext to disk.
        var persisted = new AppSettings
        {
            Sql = new SqlSettings
            {
                Server = settings.Sql.Server,
                Database = settings.Sql.Database,
                UseIntegratedSecurity = settings.Sql.UseIntegratedSecurity,
                UserId = settings.Sql.UserId,
                Password = SecretProtector.Protect(settings.Sql.Password) ?? "",
            },
            M365 = new M365Settings
            {
                TenantId = settings.M365.TenantId,
                ClientId = settings.M365.ClientId,
                Domain = settings.M365.Domain,
                DefaultMailDomain = settings.M365.DefaultMailDomain,
                SharePointAdminUrl = settings.M365.SharePointAdminUrl,
            },
            Smtp = new SmtpSettings
            {
                Host = settings.Smtp.Host,
                Port = settings.Smtp.Port,
                UseSsl = settings.Smtp.UseSsl,
                FromAddress = settings.Smtp.FromAddress,
                FromDisplayName = settings.Smtp.FromDisplayName,
                ReplyToAddress = settings.Smtp.ReplyToAddress,
                BccAddress = settings.Smtp.BccAddress,
                UserName = settings.Smtp.UserName,
                Password = SecretProtector.Protect(settings.Smtp.Password) ?? "",
            },
            SharedMailboxes = new SharedMailboxSettings
            {
                RetentionPolicyName = settings.SharedMailboxes.RetentionPolicyName,
                RoleAssignmentPolicyName = settings.SharedMailboxes.RoleAssignmentPolicyName,
                GroupMembershipHelpUrl = settings.SharedMailboxes.GroupMembershipHelpUrl,
                MailboxHelpLinks = settings.SharedMailboxes.MailboxHelpLinks,
            },
            RoomResources = new RoomResourceSettings
            {
                RoomMailboxDomain = settings.RoomResources.RoomMailboxDomain,
                RoomGroupDomain = settings.RoomResources.RoomGroupDomain,
                RoomListOwnerGroup = settings.RoomResources.RoomListOwnerGroup,
                GlobalAdminGroup = settings.RoomResources.GlobalAdminGroup,
                DefaultBookingWindowDays = settings.RoomResources.DefaultBookingWindowDays,
                DefaultMaximumDurationMinutes = settings.RoomResources.DefaultMaximumDurationMinutes,
            },
            TeamsPolicies = new TeamsPolicySettings
            {
                GroupPrefixes = settings.TeamsPolicies.GroupPrefixes,
            },
            Trace = new TraceSettings
            {
                CompanyName = settings.Trace.CompanyName,
                RedirectUrl = settings.Trace.RedirectUrl,
                InvitationMessage = settings.Trace.InvitationMessage,
            },
            NewHire = new NewHireSettings
            {
                Database = settings.NewHire.Database,
                WriteAdAttributes = settings.NewHire.WriteAdAttributes,
                LdapServer = settings.NewHire.LdapServer,
                DeploymentLocator = settings.NewHire.DeploymentLocator,
                PhoneNumberType = settings.NewHire.PhoneNumberType,
                TeamsUpgradePolicyName = settings.NewHire.TeamsUpgradePolicyName,
                NotifyEmployee = settings.NewHire.NotifyEmployee,
                AdminNotificationAddress = settings.NewHire.AdminNotificationAddress,
            },
        };

        var json = JsonSerializer.Serialize(persisted, JsonOptions);
        File.WriteAllText(SettingsFilePath, json);

        // Keep the in-memory copy with plaintext secrets for immediate use.
        Current = settings;
    }
}
