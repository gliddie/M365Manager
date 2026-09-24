using System.Windows;
using Microsoft.Win32;
using M365Manager.Core.Settings;

namespace M365Manager.Services;

public interface IThemeService
{
    /// <summary>What the user picked. <see cref="AppTheme.System"/> follows Windows.</summary>
    AppTheme Preference { get; }

    /// <summary>The theme actually on screen right now - never <see cref="AppTheme.System"/>.</summary>
    AppTheme Effective { get; }

    /// <summary>Applies the stored preference. Called once at startup, before the window is shown.</summary>
    void Initialize();

    /// <summary>Switches theme and persists the choice.</summary>
    void SetPreference(AppTheme theme);
}

/// <summary>
/// Swaps the colour dictionary in <see cref="Application.Resources"/> between Colors.Light.xaml and
/// Colors.Dark.xaml. Both define the same keys, so anything that references a brush through
/// <c>DynamicResource</c> re-reads it on the swap; a <c>StaticResource</c> reference keeps whatever
/// it resolved at load time and will NOT follow the theme.
///
/// The dictionary is found by a marker key rather than by index, so adding another dictionary to
/// App.xaml later cannot silently swap the wrong one.
/// </summary>
public sealed class ThemeService : IThemeService
{
    /// <summary>Key present in both colour dictionaries and in no other, used to locate the one to replace.</summary>
    private const string MarkerKey = "PageBrush";

    private const string LightSource = "Themes/Colors.Light.xaml";
    private const string DarkSource = "Themes/Colors.Dark.xaml";

    private readonly ISettingsService _settings;
    private bool _watchingSystem;

    public ThemeService(ISettingsService settings)
    {
        _settings = settings;
    }

    public AppTheme Preference => _settings.Current.Ui.Theme;

    public AppTheme Effective { get; private set; } = AppTheme.Light;

    public void Initialize() => Apply(Preference);

    public void SetPreference(AppTheme theme)
    {
        var settings = _settings.Current;
        settings.Ui.Theme = theme;
        try
        {
            _settings.Save(settings);
        }
        catch
        {
            // A settings file that cannot be written must not stop the theme from switching -
            // the user would otherwise click the toggle and see nothing happen.
        }

        Apply(theme);
    }

    private void Apply(AppTheme theme)
    {
        var effective = theme == AppTheme.System ? ReadWindowsTheme() : theme;
        Effective = effective;

        var app = Application.Current;
        if (app is null)
            return;

        var source = effective == AppTheme.Dark ? DarkSource : LightSource;
        var replacement = new ResourceDictionary { Source = new Uri(source, UriKind.Relative) };

        var dictionaries = app.Resources.MergedDictionaries;
        for (var i = 0; i < dictionaries.Count; i++)
        {
            if (dictionaries[i].Contains(MarkerKey))
            {
                dictionaries[i] = replacement;
                TrackSystemTheme(theme);
                return;
            }
        }

        // No colour dictionary merged yet (first call during startup) - add one.
        dictionaries.Add(replacement);
        TrackSystemTheme(theme);
    }

    /// <summary>
    /// While the preference is "System", follow Windows if the user flips app colour mode with the
    /// app already running. Subscribed only once, and only in that mode.
    /// </summary>
    private void TrackSystemTheme(AppTheme preference)
    {
        if (preference != AppTheme.System || _watchingSystem)
            return;

        SystemEvents.UserPreferenceChanged += (_, e) =>
        {
            if (e.Category != UserPreferenceCategory.General)
                return;
            if (Preference != AppTheme.System)
                return;

            Application.Current?.Dispatcher.Invoke(() => Apply(AppTheme.System));
        };
        _watchingSystem = true;
    }

    /// <summary>
    /// Windows' "Choose your default app mode" setting. 0 = dark, 1 = light; a missing value means
    /// a Windows build that predates the setting, so light.
    /// </summary>
    private static AppTheme ReadWindowsTheme()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            var value = key?.GetValue("AppsUseLightTheme");
            if (value is int useLight)
                return useLight == 0 ? AppTheme.Dark : AppTheme.Light;
        }
        catch
        {
            // Registry unreadable (policy, sandbox) - fall through to light.
        }

        return AppTheme.Light;
    }
}
