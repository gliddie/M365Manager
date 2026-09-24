using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Core.ActiveDirectory;
using M365Manager.Core.NewHire;
using M365Manager.Controls;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;
using M365Manager.Data.NewHire;

namespace M365Manager.ViewModels;

/// <summary>
/// One step outcome as the result card shows it. Mirrors <see cref="NewHireStep"/> with the bits the
/// XAML binds to.
/// </summary>
public sealed class NewHireStepViewModel
{
    public required string Name { get; init; }
    public required string Detail { get; init; }
    public required bool Succeeded { get; init; }
    public string? Error { get; init; }

    public string Icon => Succeeded ? "OK" : "FAILED";

    /// <summary>The error on a failed step; the value on a successful one.</summary>
    public string Summary => Succeeded ? Detail : $"{Detail} - {Error}";
}

/// <summary>
/// The New Hire wizard: four steps on one page rather than the four modal wizard windows of the
/// standalone tool, so it reads like the rest of the app. Steps run forward only - going back would
/// mean re-validating the employee anyway, and Start over does that cleanly.
/// </summary>
public sealed partial class NewHireViewModel : ObservableObject
{
    private readonly INewHireService _newHire;
    private readonly ISettingsService _settings;

    /// <summary>Backs the console on this page - AD writes and Teams cmdlets both land in it.</summary>
    public IPowerShellTranscript Transcript { get; }

    // --- step state ---

    /// <summary>1 = employee, 2 = number, 3 = review, 4 = result.</summary>
    [ObservableProperty] private int _step = 1;

    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusMessage = "Enter an employee and click Look up.";

    // --- step 1: employee ---

    [ObservableProperty] private string _identity = "";
    [ObservableProperty] private string _taskNumber = "";
    [ObservableProperty] private UcLocation? _selectedLocation;
    [ObservableProperty] private string _employeeSummary = "";
    [ObservableProperty] private string _licenceSummary = "";
    [ObservableProperty] private string _blocker = "";
    [ObservableProperty] private string _warnings = "";

    public ObservableCollection<UcLocation> Locations { get; } = new();

    // --- step 2: number ---

    [ObservableProperty] private UcDidRange? _selectedRange;
    [ObservableProperty] private FreeDid? _selectedDid;
    [ObservableProperty] private string _rangeSummary = "";

    public ObservableCollection<UcDidRange> DidRanges { get; } = new();
    public ObservableCollection<FreeDid> FreeDids { get; } = new();

    // --- step 3: review ---

    [ObservableProperty] private string _reviewName = "";
    [ObservableProperty] private string _reviewAccount = "";
    [ObservableProperty] private string _reviewOffice = "";
    [ObservableProperty] private string _reviewNumber = "";
    [ObservableProperty] private string _reviewSipAddress = "";
    [ObservableProperty] private string _reviewPolicies = "";
    [ObservableProperty] private string _reviewAdStep = "";

    // --- step 4: result ---

    [ObservableProperty] private string _resultHeadline = "";
    [ObservableProperty] private bool _resultSucceeded;
    [ObservableProperty] private string _teamsState = "";

    // Status severity, so a failed lookup stops looking like a successful one.
    [ObservableProperty] private AlertSeverity _statusSeverity = AlertSeverity.Info;

    private void Status(string text, AlertSeverity severity = AlertSeverity.Info)
    {
        StatusSeverity = severity;
        StatusMessage = text;
    }

    public ObservableCollection<NewHireStepViewModel> ResultSteps { get; } = new();

    private AdUser? _user;

    public NewHireViewModel(INewHireService newHire, ISettingsService settings, IPowerShellTranscript transcript)
    {
        _newHire = newHire;
        _settings = settings;
        Transcript = transcript;
    }

    // The step rail and the cards both bind to these rather than comparing Step in XAML.
    public bool IsStep1 => Step == 1;
    public bool IsStep2 => Step == 2;
    public bool IsStep3 => Step == 3;
    public bool IsStep4 => Step == 4;

    // "Already behind us", so the stepper can show a step as done rather than only marking the
    // current one. Three states - done / current / upcoming - are what make a stepper readable as
    // progress instead of as four buttons.
    public bool IsStep1Done => Step > 1;
    public bool IsStep2Done => Step > 2;
    public bool IsStep3Done => Step > 3;

    /// <summary>True once step 1 found an employee who may actually be enabled.</summary>
    public bool CanContinueFromEmployee => _user is not null && SelectedLocation is not null && Blocker.Length == 0;

    public bool HasBlocker => Blocker.Length > 0;
    public bool HasWarnings => Warnings.Length > 0;

    partial void OnStepChanged(int value)
    {
        OnPropertyChanged(nameof(IsStep1));
        OnPropertyChanged(nameof(IsStep2));
        OnPropertyChanged(nameof(IsStep3));
        OnPropertyChanged(nameof(IsStep4));

        OnPropertyChanged(nameof(IsStep1Done));
        OnPropertyChanged(nameof(IsStep2Done));
        OnPropertyChanged(nameof(IsStep3Done));
    }

    partial void OnBlockerChanged(string value)
    {
        OnPropertyChanged(nameof(HasBlocker));
        OnPropertyChanged(nameof(CanContinueFromEmployee));
    }

    partial void OnWarningsChanged(string value) => OnPropertyChanged(nameof(HasWarnings));

    /// <summary>
    /// Picking a different site re-runs the lookup with that office, because the site decides both
    /// the blocker checks and which numbers are on offer.
    /// </summary>
    partial void OnSelectedLocationChanged(UcLocation? value)
    {
        OnPropertyChanged(nameof(CanContinueFromEmployee));

        if (value is not null && _user is not null && !string.Equals(value.Name, _user.Office, StringComparison.OrdinalIgnoreCase))
            _ = LookUpAsync();
    }

    /// <summary>Loads the office list once the page is first shown, so a cold start isn't blocked on SQL.</summary>
    public async Task EnsureLocationsLoadedAsync()
    {
        if (Locations.Count > 0)
            return;

        try
        {
            var locations = await _newHire.GetLocationsAsync();
            foreach (var location in locations)
                Locations.Add(location);
        }
        catch (Exception ex)
        {
            Status($"Could not load the site list: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
    }

    // ----- step 1 -----

    [RelayCommand]
    private async Task LookUpAsync()
    {
        if (string.IsNullOrWhiteSpace(Identity))
        {
            Status("Enter a SamAccountName, UPN or e-mail address.", AlertSeverity.Error);
            return;
        }

        IsBusy = true;
        Transcript.BeginOperation($"New hire lookup: {Identity.Trim()}");
        EmployeeSummary = "";
        LicenceSummary = "";
        Blocker = "";
        Warnings = "";

        try
        {
            await EnsureLocationsLoadedAsync();

            // SelectedLocation is the operator's override; on the first lookup it is null and the
            // office from the AD account is used instead.
            var lookup = await _newHire.LookUpAsync(Identity.Trim(), SelectedLocation?.Name);

            if (!lookup.Succeeded || lookup.User is null)
            {
                _user = null;
                Status(lookup.ErrorMessage ?? "Lookup failed.", AlertSeverity.Error);
                OnPropertyChanged(nameof(CanContinueFromEmployee));
                return;
            }

            _user = lookup.User;
            EmployeeSummary = $"{lookup.User.DisplayName} ({lookup.User.UserPrincipalName})";
            LicenceSummary = lookup.License.Summary;

            // Only adopt the matched site when the operator has not picked one themselves.
            if (SelectedLocation is null && lookup.Location is not null)
            {
                SelectedLocation = Locations.FirstOrDefault(l => l.Id == lookup.Location.Id) ?? lookup.Location;
            }

            Blocker = lookup.Blocker ?? "";
            Warnings = string.Join("\n", lookup.Warnings);

            StatusMessage = Blocker.Length > 0
                ? "This employee cannot be enabled - see below."
                : "Employee found. Continue to pick a number.";

            OnPropertyChanged(nameof(CanContinueFromEmployee));
        }
        catch (Exception ex)
        {
            Status($"Lookup failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task GoToNumberAsync()
    {
        if (!CanContinueFromEmployee || SelectedLocation is null)
            return;

        if (string.IsNullOrWhiteSpace(TaskNumber))
        {
            Status("Enter the ticket/task number authorizing this change.", AlertSeverity.Error);
            return;
        }

        Step = 2;
        await LoadRangesAsync();
    }

    // ----- step 2 -----

    private async Task LoadRangesAsync()
    {
        if (SelectedLocation is null)
            return;

        IsBusy = true;
        DidRanges.Clear();
        FreeDids.Clear();
        SelectedDid = null;
        RangeSummary = "";

        try
        {
            var ranges = await _newHire.GetDidRangesAsync(SelectedLocation.LocationCode);
            foreach (var range in ranges)
                DidRanges.Add(range);

            StatusMessage = DidRanges.Count == 0
                ? $"No usable number blocks are configured for {SelectedLocation.LocationCode}."
                : $"{DidRanges.Count} number block(s) for {SelectedLocation.LocationCode}. Pick one, then search for free numbers.";

            // With a single block there is nothing to choose - select it so one click less is needed.
            if (DidRanges.Count == 1)
                SelectedRange = DidRanges[0];
        }
        catch (Exception ex)
        {
            Status($"Could not load number blocks: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task FindFreeNumbersAsync()
    {
        if (SelectedRange is null)
        {
            Status("Select a number block first.", AlertSeverity.Error);
            return;
        }

        IsBusy = true;
        FreeDids.Clear();
        SelectedDid = null;

        try
        {
            var free = await _newHire.FindFreeDidsAsync(SelectedRange);
            foreach (var did in free)
                FreeDids.Add(did);

            RangeSummary = $"{SelectedRange.DidStart} - {SelectedRange.DidEnd}: {FreeDids.Count} free of "
                           + $"{CountInRange(SelectedRange)} numbers";

            if (FreeDids.Count == 0)
            {
                Status("That block is fully allocated. Pick a different one.", AlertSeverity.Error);
                return;
            }

            // Picked at random rather than lowest-first, as the legacy tool did: consecutive numbers
            // handed out in sequence make a whole team trivially guessable from one of them.
            SelectedDid = FreeDids[Random.Shared.Next(FreeDids.Count)];
            Status($"{FreeDids.Count} number(s) free. {SelectedDid.E164} is proposed - change it if you like.", AlertSeverity.Success);
        }
        catch (Exception ex)
        {
            Status($"Number search failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private static long CountInRange(UcDidRange range) =>
        long.TryParse(range.DidStart, out var start) && long.TryParse(range.DidEnd, out var end) && end >= start
            ? end - start + 1
            : 0;

    [RelayCommand]
    private void GoToReview()
    {
        if (_user is null || SelectedLocation is null || SelectedDid is null)
        {
            Status("Pick a number before continuing.", AlertSeverity.Error);
            return;
        }

        var request = BuildRequest();
        var newHire = _settings.Current.NewHire;

        ReviewName = _user.DisplayName;
        ReviewAccount = $"{_user.SamAccountName} ({_user.UserPrincipalName})";
        ReviewOffice = $"{SelectedLocation.Name} - {SelectedLocation.LocationCode}";
        ReviewNumber = request.PhoneNumber;
        ReviewSipAddress = request.SipAddress;
        ReviewPolicies = $"Voice routing, dial plan, emergency calling and emergency call routing: "
                         + $"{request.PolicyName}. Teams upgrade: {newHire.TeamsUpgradePolicyName}.";
        ReviewAdStep = newHire.WriteAdAttributes
            ? $"msRTCSIP-Line = {request.LineUri} and msRTCSIP-PrimaryUserAddress = {request.SipAddress} are written "
              + "to the on-premises account first, so the Entra Connect sync agrees with Teams instead of overwriting it."
            : "On-premises AD is not touched (turned off under Settings > New Hire). Only do this once the "
              + "msRTCSIP-* attributes are no longer synchronised - otherwise the next sync can clear the number again.";

        Step = 3;
        StatusMessage = "Review the summary, then apply.";
    }

    private NewHireRequest BuildRequest() => new()
    {
        TaskNumber = TaskNumber.Trim(),
        User = _user!,
        Location = SelectedLocation!,
        Did = SelectedDid!.Did,
    };

    // ----- step 3 -----

    [RelayCommand]
    private async Task ConfigureAsync()
    {
        if (_user is null || SelectedLocation is null || SelectedDid is null)
            return;

        IsBusy = true;
        ResultSteps.Clear();
        Transcript.BeginOperation($"Enable for telephony: {_user.UserPrincipalName} on +{SelectedDid.Did}");

        try
        {
            var result = await _newHire.ConfigureAsync(BuildRequest(), msg => StatusMessage = msg);

            foreach (var step in result.Steps)
            {
                ResultSteps.Add(new NewHireStepViewModel
                {
                    Name = step.Name,
                    Detail = step.Detail,
                    Succeeded = step.Succeeded,
                    Error = step.Error,
                });
            }

            ResultSucceeded = result.Succeeded;

            if (!result.Succeeded)
            {
                ResultHeadline = $"{_user.DisplayName} could not be enabled: {result.ErrorMessage}";
                Status("The run failed. The steps below show how far it got.", AlertSeverity.Error);
            }
            else if (result.NumberPendingSync)
            {
                ResultHeadline = $"{_user.DisplayName} has been given +{SelectedDid.Did} in Active Directory. "
                                 + "Teams would not take the number directly because it is managed on-premises, "
                                 + "so it arrives with the next Entra Connect sync.";
                Status("Done - the number reaches Teams via directory sync.", AlertSeverity.Success);
            }
            else if (result.HasFailedSteps)
            {
                ResultHeadline = $"{_user.DisplayName} is enabled on +{SelectedDid.Did}, "
                                 + "but some steps failed - check them below.";
                Status("Finished with warnings.", AlertSeverity.Warning);
            }
            else
            {
                ResultHeadline = $"{_user.DisplayName} is enabled on +{SelectedDid.Did}"
                                 + (result.EmployeeNotified ? " and has been notified." : ".");
                Status("Done.", AlertSeverity.Success);
            }

            TeamsState = BuildTeamsState(result, $"tel:+{SelectedDid.Did}");

            Step = 4;
        }
        catch (Exception ex)
        {
            Status($"The run failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    /// What Teams holds right after the run. The read-back happens seconds after the change, so on a
    /// tenant where the number arrives by directory sync it will still show the old line - saying so
    /// is the difference between "it worked, wait" and an operator re-running the wizard.
    /// </summary>
    private static string BuildTeamsState(NewHireResult result, string expectedLineUri)
    {
        if (result.VoiceState is null)
            return "Teams could not be read back - the changes may still have landed.";

        var state = result.VoiceState;
        var summary = $"Teams reports: line {Or(state.LineUri, "(none)")}, enterprise voice "
                      + $"{(state.EnterpriseVoiceEnabled ? "on" : "off")}, voice routing "
                      + $"{Or(state.OnlineVoiceRoutingPolicy, "(global)")}, dial plan "
                      + $"{Or(state.TenantDialPlan, "(global)")}, mode "
                      + $"{Or(state.TeamsUpgradeEffectiveMode, "(unknown)")}.";

        if (!string.Equals(state.LineUri, expectedLineUri, StringComparison.OrdinalIgnoreCase))
            summary += " The line still shows the previous value - Entra Connect can take up to 30 minutes "
                       + "to carry the new number across. The policies above apply immediately.";

        return summary;
    }

    private static string Or(string value, string fallback) => value.Length == 0 ? fallback : value;

    [RelayCommand]
    private void Back()
    {
        if (Step > 1 && Step < 4)
            Step--;
    }

    /// <summary>
    /// Clears everything for the next hire. The office list is kept - it is the same for every run
    /// and re-reading it just makes the page wait on SQL again.
    /// </summary>
    [RelayCommand]
    private void StartOver()
    {
        _user = null;
        Identity = "";
        TaskNumber = "";
        SelectedLocation = null;
        EmployeeSummary = "";
        LicenceSummary = "";
        Blocker = "";
        Warnings = "";
        DidRanges.Clear();
        FreeDids.Clear();
        SelectedRange = null;
        SelectedDid = null;
        RangeSummary = "";
        ResultSteps.Clear();
        ResultHeadline = "";
        TeamsState = "";
        Step = 1;
        Status("Enter an employee and click Look up.", AlertSeverity.Error);
    }
}
