using System.Diagnostics;
using System.Text.RegularExpressions;

namespace M365Manager.Services;

/// <summary>
/// Shared helpers for surfacing an interactive sign-in prompt (device code / browser URL) that
/// a connector reports via its onPrompt callback. Used by both the app-wide startup connect
/// (MainViewModel) and the per-page manual "Connect" retry (GroupsViewModel).
/// </summary>
public static class DeviceCodePrompt
{
    /// <summary>Extracts the short sign-in code from a prompt like "...enter the code ABCD1234 to authenticate.".</summary>
    public static string? ExtractCode(string prompt)
    {
        var match = Regex.Match(prompt, @"code\s+([A-Za-z0-9][A-Za-z0-9-]{4,})");
        return match.Success ? match.Groups[1].Value : null;
    }

    /// <summary>Opens the first URL found in the prompt in the default browser. Returns false if none was found or it couldn't be opened.</summary>
    public static bool TryOpenBrowser(string prompt)
    {
        var match = Regex.Match(prompt, @"https?://\S+");
        if (!match.Success)
            return false;

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = match.Value,
                UseShellExecute = true,
            });
            return true;
        }
        catch
        {
            // If the browser cannot be opened, the URL is still shown in the banner.
            return false;
        }
    }
}
