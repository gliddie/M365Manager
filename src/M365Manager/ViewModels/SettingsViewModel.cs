using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Core.M365;
using M365Manager.Core.Settings;
using M365Manager.Data.Logging;

namespace M365Manager.ViewModels;

public sealed partial class SettingsViewModel : ObservableObject
{
    private readonly ISettingsService _settings;
    private readonly ILogService _log;
    private readonly IM365AuthService _auth;

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

    // --- UI state ---
    [ObservableProperty] private string _statusMessage = "";
    [ObservableProperty] private bool _isBusy;

    public string SettingsFilePath => _settings.SettingsFilePath;

    public SettingsViewModel(ISettingsService settings, ILogService log, IM365AuthService auth)
    {
        _settings = settings;
        _log = log;
        _auth = auth;

        var s = _settings.Current;
        SqlServer = s.Sql.Server;
        SqlDatabase = s.Sql.Database;
        SqlUseIntegratedSecurity = s.Sql.UseIntegratedSecurity;
        SqlUserId = s.Sql.UserId;
        SqlPassword = s.Sql.Password;
        TenantId = s.M365.TenantId;
        ClientId = s.M365.ClientId;
        Domain = s.M365.Domain;
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
        _settings.Save(s);
    }

    [RelayCommand]
    private void Save()
    {
        ApplyToSettings();
        StatusMessage = $"Settings saved to {_settings.SettingsFilePath}";
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
            StatusMessage = $"SQL error: {ex.Message}";
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
            StatusMessage = $"Sign-in failed: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }
}
