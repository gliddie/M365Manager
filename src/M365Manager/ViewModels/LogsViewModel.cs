using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Core.Logging;
using M365Manager.Data.Logging;

namespace M365Manager.ViewModels;

public sealed partial class LogsViewModel : ObservableObject
{
    private const string AnyValue = "(all)";

    private readonly ILogService _log;

    [ObservableProperty] private string _statusMessage = "";
    [ObservableProperty] private bool _isBusy;

    // --- Filter ---
    [ObservableProperty] private string _searchText = "";
    [ObservableProperty] private string _taskNumber = "";
    [ObservableProperty] private string _selectedUser = AnyValue;
    [ObservableProperty] private string _selectedArea = AnyValue;
    [ObservableProperty] private string _selectedAction = AnyValue;
    [ObservableProperty] private string _selectedSeverity = AnyValue;
    [ObservableProperty] private DateTime? _fromDate;
    [ObservableProperty] private DateTime? _toDate;
    [ObservableProperty] private int _maxResults = 1000;

    public ObservableCollection<string> Users { get; } = new() { AnyValue };
    public ObservableCollection<string> Areas { get; } = new() { AnyValue };
    public ObservableCollection<string> Actions { get; } = new() { AnyValue };

    public IReadOnlyList<string> Severities { get; } =
        new[] { AnyValue }.Concat(Enum.GetNames<Severity>()).ToList();

    public IReadOnlyList<int> ResultLimits { get; } = new[] { 200, 500, 1000, 5000, 20000 };

    public ObservableCollection<LogEntry> Entries { get; } = new();

    public LogsViewModel(ILogService log)
    {
        _log = log;
    }

    /// <summary>
    /// Loads the drop-down values from the whole table, so filtering by a user or area that hasn't
    /// appeared in the currently displayed page is still possible.
    /// </summary>
    private async Task LoadFilterOptionsAsync()
    {
        try
        {
            var options = await _log.GetFilterOptionsAsync();

            // Clearing a bound ItemsSource makes the ComboBox push null back into the bound
            // SelectedItem, so each selection has to be captured first and restored afterwards -
            // otherwise refreshing the list silently drops the filter the operator just set.
            var user = SelectedUser;
            var area = SelectedArea;
            var action = SelectedAction;

            Replace(Users, options.Users, user);
            Replace(Areas, options.Areas, area);
            Replace(Actions, options.Actions, action);

            SelectedUser = user;
            SelectedArea = area;
            SelectedAction = action;
        }
        catch
        {
            // Best-effort: without SQL the drop-downs just stay at "(all)".
        }
    }

    private static void Replace(ObservableCollection<string> target, IReadOnlyList<string> values, string keepSelected)
    {
        target.Clear();
        target.Add(AnyValue);
        foreach (var value in values)
            target.Add(value);

        // A value that is still selected but no longer present in the log (e.g. its last entry
        // aged out) stays choosable rather than silently resetting the filter.
        if (!string.IsNullOrEmpty(keepSelected) && !target.Contains(keepSelected))
            target.Add(keepSelected);
    }

    [RelayCommand]
    private async Task RefreshAsync()
    {
        IsBusy = true;
        StatusMessage = "Loading...";
        try
        {
            // Built before the drop-downs are reloaded, so a repopulated ItemsSource cannot
            // influence what is queried.
            var query = new LogQuery
            {
                // The grid shows UTC, so the pickers are read as UTC too - no surprise shift
                // between what was typed and what comes back.
                FromUtc = FromDate?.Date,
                ToUtc = ToDate?.Date.AddDays(1).AddTicks(-1),
                UserUpn = Normalize(SelectedUser),
                Area = Normalize(SelectedArea),
                Action = Normalize(SelectedAction),
                TaskNumber = string.IsNullOrWhiteSpace(TaskNumber) ? null : TaskNumber.Trim(),
                Severity = Enum.TryParse<Severity>(Normalize(SelectedSeverity), out var severity) ? severity : null,
                SearchText = string.IsNullOrWhiteSpace(SearchText) ? null : SearchText.Trim(),
                MaxResults = MaxResults,
            };

            await LoadFilterOptionsAsync();

            var items = await _log.QueryAsync(query);
            Entries.Clear();
            foreach (var item in items)
                Entries.Add(item);

            StatusMessage = Entries.Count >= MaxResults
                ? $"{Entries.Count} entries (newest first) - limit reached, narrow the filter or raise the limit."
                : $"{Entries.Count} entries (newest first).";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Could not load logs: {ErrorText.Describe(ex)}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    private static string? Normalize(string value)
        => string.IsNullOrWhiteSpace(value) || value == AnyValue ? null : value;

    [RelayCommand]
    private async Task ClearFilterAsync()
    {
        SearchText = "";
        TaskNumber = "";
        SelectedUser = AnyValue;
        SelectedArea = AnyValue;
        SelectedAction = AnyValue;
        SelectedSeverity = AnyValue;
        FromDate = null;
        ToDate = null;
        await RefreshAsync();
    }

    /// <summary>Exports exactly what the grid currently shows - i.e. the filtered result.</summary>
    [RelayCommand]
    private void ExportToExcel()
    {
        if (Entries.Count == 0)
        {
            StatusMessage = "Nothing to export - the filtered list is empty.";
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = $"Logs-{DateTime.Now:yyyyMMdd-HHmm}.xlsx",
        };
        if (dialog.ShowDialog() != true)
            return;

        try
        {
            LogExportService.ExportToExcel(dialog.FileName, Entries.ToList());
            StatusMessage = $"Exported {Entries.Count} entr{(Entries.Count == 1 ? "y" : "ies")} to {dialog.FileName}.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Export failed: {ErrorText.Describe(ex)}";
        }
    }
}
