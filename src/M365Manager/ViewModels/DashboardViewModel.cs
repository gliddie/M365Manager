using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Core.M365;
using M365Manager.Core.Settings;

namespace M365Manager.ViewModels;

public sealed partial class DashboardViewModel : ObservableObject
{
    private readonly ISettingsService _settings;
    private readonly IM365AuthService _auth;

    [ObservableProperty]
    private string _sqlStatus = "";

    [ObservableProperty]
    private string _m365Status = "";

    public DashboardViewModel(ISettingsService settings, IM365AuthService auth)
    {
        _settings = settings;
        _auth = auth;
        Refresh();
    }

    [RelayCommand]
    private void Refresh()
    {
        var sql = _settings.Current.Sql;
        SqlStatus = string.IsNullOrWhiteSpace(sql.Server)
            ? "Not configured yet - open Settings."
            : $"{sql.Server}  •  {sql.Database}";

        M365Status = _auth.IsSignedIn
            ? $"Signed in as {_auth.CurrentUser!.Upn}"
            : "Not signed in - open Settings to connect.";
    }
}
