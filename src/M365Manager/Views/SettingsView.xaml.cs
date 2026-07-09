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
        // Seed the PasswordBox from the view model (it cannot be data-bound directly).
        if (DataContext is SettingsViewModel vm)
            SqlPasswordBox.Password = vm.SqlPassword;
    }

    private void SqlPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (DataContext is SettingsViewModel vm)
            vm.SqlPassword = SqlPasswordBox.Password;
    }
}
