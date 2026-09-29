namespace M365Manager.ViewModels;

/// <summary>
/// A page that wants to know when the shell shows it - for work that must not run at startup, such
/// as a live read that needs a connection the startup sign-in may not have finished yet.
/// </summary>
public interface IActivatablePage
{
    void OnActivated();
}