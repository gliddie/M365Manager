namespace M365Manager.ViewModels;

/// <summary>
/// A page whose view is a tab strip, so the shell can land on a particular tab when something
/// navigates here - a Home tile that says "Create shared mailbox" should open the create form, not
/// the overview the page normally starts on.
///
/// The view binds its TabControl's SelectedIndex to <see cref="SelectedTabIndex"/>; nothing else
/// reads it.
/// </summary>
public interface ITabbedPage
{
    int SelectedTabIndex { get; set; }
}
