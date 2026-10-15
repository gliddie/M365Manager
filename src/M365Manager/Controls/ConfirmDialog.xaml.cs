using System.Windows;

namespace M365Manager.Controls;

/// <summary>
/// Modal confirmation for a destructive action. Created through
/// <see cref="M365Manager.Services.IDialogService"/> rather than directly, so view models never
/// reference a window type.
///
/// Cancel is the default focus and Esc cancels: the safe answer should be the one you get by
/// reflex. Confirm is deliberately not the Enter key.
/// </summary>
public partial class ConfirmDialog : Window
{
    /// <param name="destructive">
    /// False for a confirmation that is not a loss (e.g. assigning a paid licence): info icon and the
    /// screen's primary button instead of the warning triangle and the red confirm button.
    /// </param>
    public ConfirmDialog(string title, string message, string confirmLabel, string? detail, bool destructive = true)
    {
        InitializeComponent();

        TitleText.Text = title;
        MessageText.Text = message;
        ConfirmButton.Content = confirmLabel;

        if (!destructive)
        {
            // SetResourceReference = DynamicResource, so a theme switch still reaches the dialog.
            IconDisc.SetResourceReference(System.Windows.Controls.Border.BackgroundProperty, "InfoSoftBrush");
            IconGlyph.SetResourceReference(System.Windows.Controls.TextBlock.ForegroundProperty, "InfoBrush");
            IconGlyph.Text = ""; // Info
            ConfirmButton.SetResourceReference(StyleProperty, "PrimaryButton");
        }

        if (!string.IsNullOrWhiteSpace(detail))
        {
            DetailText.Text = detail;
            DetailPanel.Visibility = Visibility.Visible;
        }

        Loaded += (_, _) => CancelButton.Focus();
    }

    private void Confirm_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
