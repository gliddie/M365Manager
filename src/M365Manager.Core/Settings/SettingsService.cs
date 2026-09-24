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

    /// <summary>
    /// Reading is deliberately more forgiving than writing: a settings file gets hand-edited during
    /// troubleshooting, and a stray trailing comma or a commented-out line should not cost the user
    /// their configuration.
    /// </summary>
    private static readonly JsonSerializerOptions ReadOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        AllowTrailingCommas = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
    };

    public string SettingsFilePath { get; }

    public AppSettings Current { get; private set; } = new();

    public string? LastLoadError { get; private set; }

    public string? QuarantinedFilePath { get; private set; }

    public SettingsService()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "M365Manager");
        Directory.CreateDirectory(dir);
        SettingsFilePath = Path.Combine(dir, "settings.json");
        Load();
    }

    /// <summary>
    /// Reads the settings file section by section.
    ///
    /// This used to be one try/catch around the whole deserialization, falling back to a blank
    /// <see cref="AppSettings"/>. Two things made that dangerous. A single unreadable value - one
    /// mistyped port, one hand-edited enum - discarded *every* section, not just its own. And
    /// because the blank result then sat in <see cref="Current"/>, the next <see cref="Save"/> from
    /// anywhere in the app wrote those blanks over a settings file that was still almost entirely
    /// intact, turning a read error into permanent loss.
    ///
    /// Now each section is bound on its own, so a bad one costs only itself; the failure is
    /// reported through <see cref="LastLoadError"/> instead of being swallowed; and whenever
    /// anything failed to load, the original file is copied aside first (see
    /// <see cref="QuarantinedFilePath"/>) so a later Save cannot destroy it.
    /// </summary>
    public AppSettings Load()
    {
        LastLoadError = null;
        QuarantinedFilePath = null;

        if (!File.Exists(SettingsFilePath))
        {
            Current = new AppSettings();
            return Current;
        }

        string json;
        try
        {
            json = File.ReadAllText(SettingsFilePath);
        }
        catch (Exception ex)
        {
            // Locked or unreadable. Nothing is quarantined here on purpose - the file is fine, we
            // just cannot see it, and copying it would likely fail for the same reason.
            LastLoadError = $"The settings file could not be read ({ex.Message}). "
                            + "The app is running with default settings - do not save until this is resolved, "
                            + "or the defaults will replace your configuration.";
            Current = new AppSettings();
            return Current;
        }

        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(json, new JsonDocumentOptions
            {
                AllowTrailingCommas = true,
                CommentHandling = JsonCommentHandling.Skip,
            });
        }
        catch (JsonException ex)
        {
            LastLoadError = $"The settings file is not valid JSON ({ex.Message}). "
                            + "The app is running with default settings.";
            QuarantinedFilePath = Quarantine();
            Current = new AppSettings();
            return Current;
        }

        using var _ = document;
        var settings = new AppSettings();
        var failedSections = new List<string>();
        var root = document.RootElement;

        settings.Sql = Section(root, nameof(AppSettings.Sql), settings.Sql, failedSections);
        settings.M365 = Section(root, nameof(AppSettings.M365), settings.M365, failedSections);
        settings.Smtp = Section(root, nameof(AppSettings.Smtp), settings.Smtp, failedSections);
        settings.SharedMailboxes = Section(root, nameof(AppSettings.SharedMailboxes), settings.SharedMailboxes, failedSections);
        settings.RoomResources = Section(root, nameof(AppSettings.RoomResources), settings.RoomResources, failedSections);
        settings.TeamsPolicies = Section(root, nameof(AppSettings.TeamsPolicies), settings.TeamsPolicies, failedSections);
        settings.Trace = Section(root, nameof(AppSettings.Trace), settings.Trace, failedSections);
        settings.NewHire = Section(root, nameof(AppSettings.NewHire), settings.NewHire, failedSections);
        settings.Ui = Section(root, nameof(AppSettings.Ui), settings.Ui, failedSections);

        // Unprotect never throws - it returns the input unchanged when a value is not DPAPI
        // ciphertext - so these cannot take the whole load down with them.
        settings.Sql.Password = SecretProtector.Unprotect(settings.Sql.Password) ?? "";
        settings.Smtp.Password = SecretProtector.Unprotect(settings.Smtp.Password) ?? "";
        settings.NewHire.AdPassword = SecretProtector.Unprotect(settings.NewHire.AdPassword) ?? "";

        if (failedSections.Count > 0)
        {
            LastLoadError = $"These settings sections could not be read and were reset to defaults: "
                            + $"{string.Join(", ", failedSections)}. Check them before saving.";
            QuarantinedFilePath = Quarantine();
        }

        Current = settings;
        return Current;
    }

    /// <summary>
    /// Binds one top-level section, falling back to the already-constructed default when it is
    /// missing or unreadable. A missing section is normal (the file predates it) and is not
    /// reported; only a present-but-broken one counts as a failure.
    /// </summary>
    private static T Section<T>(JsonElement root, string name, T fallback, ICollection<string> failed)
        where T : class
    {
        if (!root.TryGetProperty(name, out var element) || element.ValueKind == JsonValueKind.Null)
            return fallback;

        try
        {
            return element.Deserialize<T>(ReadOptions) ?? fallback;
        }
        catch (Exception)
        {
            failed.Add(name);
            return fallback;
        }
    }

    /// <summary>
    /// Copies the current settings file aside so a later <see cref="Save"/> cannot overwrite the
    /// only copy of a configuration we failed to parse. Best-effort: if even this fails there is
    /// nothing useful left to do about it.
    /// </summary>
    /// <returns>Path of the copy, or null when it could not be made.</returns>
    private string? Quarantine()
    {
        try
        {
            var path = $"{SettingsFilePath}.unreadable-{DateTime.Now:yyyyMMdd-HHmmss}.bak";
            File.Copy(SettingsFilePath, path, overwrite: true);
            return path;
        }
        catch
        {
            return null;
        }
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
                AdUserName = settings.NewHire.AdUserName,
                AdPassword = SecretProtector.Protect(settings.NewHire.AdPassword) ?? "",
                DeploymentLocator = settings.NewHire.DeploymentLocator,
                PhoneNumberType = settings.NewHire.PhoneNumberType,
                TeamsUpgradePolicyName = settings.NewHire.TeamsUpgradePolicyName,
                NotifyEmployee = settings.NewHire.NotifyEmployee,
                AdminNotificationAddress = settings.NewHire.AdminNotificationAddress,
            },
            Ui = new UiSettings
            {
                Theme = settings.Ui.Theme,
            },
        };

        var json = JsonSerializer.Serialize(persisted, JsonOptions);
        File.WriteAllText(SettingsFilePath, json);

        // Keep the in-memory copy with plaintext secrets for immediate use.
        Current = settings;
    }
}
