using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Data.Logging;

namespace M365Manager.ViewModels;

public sealed partial class LogsViewModel : ObservableObject
{
    private readonly ILogService _log;

    [ObservableProperty] private string _statusMessage = "";
    [ObservableProperty] private bool _isBusy;

    public ObservableCollection<LogEntry> Entries { get; } = new();

    public LogsViewModel(ILogService log)
    {
        _log = log;
    }

    [RelayCommand]
    private async Task RefreshAsync()
    {
        IsBusy = true;
        StatusMessage = "Loading...";
        try
        {
            var items = await _log.GetRecentAsync(200);
            Entries.Clear();
            foreach (var item in items)
                Entries.Add(item);
            StatusMessage = $"{Entries.Count} entries (newest first).";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Could not load logs: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }
}
