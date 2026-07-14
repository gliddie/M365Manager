using System.Windows.Controls;
using M365Manager.Core.Exchange;
using M365Manager.ViewModels;

namespace M365Manager.Views;

public partial class GroupsView : UserControl
{
    public GroupsView()
    {
        InitializeComponent();
    }

    // DataGrid.SelectedItems isn't bindable, so multi-selection is forwarded here for the
    // bulk "Remove selected" command (see GroupsViewModel.UpdateMemberSelection).
    private void MembersDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (DataContext is GroupsViewModel vm && sender is DataGrid grid)
            vm.UpdateMemberSelection(grid.SelectedItems.Cast<GroupMemberInfo>());
    }
}
