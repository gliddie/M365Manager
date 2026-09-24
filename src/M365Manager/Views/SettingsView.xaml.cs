using System.Windows;
using System.Windows.Controls;
using M365Manager.ViewModels;

namespace M365Manager.Views;

public partial class SettingsView : UserControl
{
    public SettingsView()
    {
        InitializeComponent();
    }

    /// <summary>
    /// Seeds one PasswordBox from the view model. A PasswordBox cannot be data-bound, so it has to
    /// be filled in code.
    ///
    /// This hangs off each box's own Loaded event rather than the view's: the settings sections are
    /// tabs now, and WPF does not build a tab's content tree until that tab is first selected. The
    /// view-level handler this replaces reached for all three boxes at once and would have found
    /// two of them null.
    /// </summary>
    private void PasswordBox_Loaded(object sender, RoutedEventArgs e)
    {
        if (DataContext is not SettingsViewModel vm || sender is not PasswordBox box)
            return;

        box.Password = box.Name switch
        {
            nameof(SqlPasswordBox) => vm.SqlPassword,
            nameof(SmtpPasswordBox) => vm.SmtpPassword,
            nameof(AdPasswordBox) => vm.NewHireAdPassword,
            _ => box.Password,
        };
    }

    private void SqlPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (DataContext is SettingsViewModel vm)
            vm.SqlPassword = SqlPasswordBox.Password;
    }

    private void SmtpPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (DataContext is SettingsViewModel vm)
            vm.SmtpPassword = SmtpPasswordBox.Password;
    }

    private void AdPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (DataContext is SettingsViewModel vm)
            vm.NewHireAdPassword = AdPasswordBox.Password;
    }
}
