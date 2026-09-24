using System.Linq;
using System.Windows;
using M365Manager.Controls;

namespace M365Manager.Services;

public interface IDialogService
{
    /// <summary>
    /// Asks the operator to confirm something that cannot be undone. Returns true only on an
    /// explicit confirm - closing the dialog any other way counts as "no".
    /// </summary>
    /// <param name="title">Question form, naming the object: "Delete LST.EMEA.Sales?"</param>
    /// <param name="message">What will happen, in plain words.</param>
    /// <param name="confirmLabel">Verb on the confirm button: "Delete group", not "OK".</param>
    /// <param name="detail">Optional extra context - what else goes with it, what is logged first.</param>
    bool ConfirmDestructive(string title, string message, string confirmLabel, string? detail = null);
}

/// <summary>
/// Shows the app's own <see cref="ConfirmDialog"/> instead of <see cref="MessageBox"/>.
///
/// MessageBox was used in exactly one place, while four other destructive actions relied on an
/// inline "yes I am sure" checkbox and two had no confirmation at all. Routing every one of them
/// through here makes the confirmation consistent, themed, and impossible to skip by scrolling past.
/// </summary>
public sealed class DialogService : IDialogService
{
    public bool ConfirmDestructive(string title, string message, string confirmLabel, string? detail = null)
    {
        var dialog = new ConfirmDialog(title, message, confirmLabel, detail)
        {
            Owner = ActiveWindow(),
        };

        return dialog.ShowDialog() == true;
    }

    /// <summary>
    /// The window to centre on. Falls back to the main window, then to none - a dialog without an
    /// owner still works, it just centres on screen rather than on the app.
    /// </summary>
    private static Window? ActiveWindow()
    {
        var app = Application.Current;
        if (app is null)
            return null;

        return app.Windows.OfType<Window>().FirstOrDefault(w => w.IsActive)
               ?? app.MainWindow;
    }
}
