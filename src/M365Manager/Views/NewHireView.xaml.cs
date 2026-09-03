using System.Windows.Controls;
using M365Manager.ViewModels;

namespace M365Manager.Views;

public partial class NewHireView : UserControl
{
    public NewHireView()
    {
        InitializeComponent();

        // The site list comes from the telephony database, so it is loaded when the page is first
        // shown rather than in the constructor - a cold start shouldn't wait on SQL.
        Loaded += async (_, _) =>
        {
            if (DataContext is NewHireViewModel vm)
                await vm.EnsureLocationsLoadedAsync();
        };
    }
}
