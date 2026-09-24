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
    public ConfirmDialog(string title, string message, string confirmLabel, string? detail)
    {
        InitializeComponent();

        TitleText.Text = title;
        MessageText.Text = message;
        ConfirmButton.Content = confirmLabel;

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
