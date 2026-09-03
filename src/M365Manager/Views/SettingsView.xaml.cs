using System.Windows;
using System.Windows.Controls;
using M365Manager.ViewModels;

namespace M365Manager.Views;

public partial class SettingsView : UserControl
{
    public SettingsView()
    {
        InitializeComponent();
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        // Seed the PasswordBoxes from the view model (they cannot be data-bound directly).
        if (DataContext is not SettingsViewModel vm)
            return;

        SqlPasswordBox.Password = vm.SqlPassword;
        SmtpPasswordBox.Password = vm.SmtpPassword;
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
}
