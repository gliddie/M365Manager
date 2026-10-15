using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Controls;
using M365Manager.Core.RoomResources;
using M365Manager.Core.Telephony;
using M365Manager.Data.NewHire;

namespace M365Manager.ViewModels;

/// <summary>
/// The Telephony tab of the manage panel: gives the selected room a number and the site's voice
/// policies. Three steps on one tab - look up, pick a number, assign - because the room is already
/// chosen in the grid and the ticket is already in the panel above.
///
/// Not part of the Create form on purpose: a new room account has no licence, and Teams refuses a
/// number until it has one with Teams Phone (Teams Rooms Pro) and has picked that licence up. The
/// Create tab therefore ends with a jump to this tab, which offers a free Teams Rooms licence - and
/// assigns it only once the operator confirms.
/// </summary>
public sealed partial class RoomResourcesViewModel
{
    /// <summary>Index of the Telephony tab in the manage panel - see RoomResourcesView.xaml.</summary>
    private const int TelephonyTabIndex = 2;

    /// <summary>Which manage sub-tab is shown; set when the Create tab jumps here.</summary>
    [ObservableProperty] private int _manageTabIndex;

    [ObservableProperty] private string _telephonyAccountSummary = "";
    [ObservableProperty] private string _telephonyLicenceSummary = "";
    [ObservableProperty] private string _telephonyBlocker = "";
    [ObservableProperty] private string _telephonyWarnings = "";
    [ObservableProperty] private string _telephonyPlan = "";
    [ObservableProperty] private string _telephonyResult = "";
    [ObservableProperty] private AlertSeverity _telephonyResultSeverity = AlertSeverity.Info;

    [ObservableProperty] private UcLocation? _selectedTelephonyLocation;
    [ObservableProperty] private UcDidRange? _selectedTelephonyRange;
    [ObservableProperty] private FreeDid? _selectedTelephonyDid;

    public ObservableCollection<UcLocation> TelephonyLocations { get; } = new();
    public ObservableCollection<UcDidRange> TelephonyRanges { get; } = new();
    public ObservableCollection<FreeDid> TelephonyFreeDids { get; } = new();

    private RoomTelephonyLookup? _telephonyLookup;

    /// <summary>Set by Create, so its "Set up telephony" button knows which room to open.</summary>
    [ObservableProperty] private string _lastCreatedRoomAddress = "";

    public bool HasTelephonyLookup => _telephonyLookup is not null;
    public bool HasTelephonyBlocker => TelephonyBlocker.Length > 0;

    /// <summary>A free Teams Rooms licence the room could get - only ever assigned after a confirm.</summary>
    public RoomLicenceOffer? TelephonyLicenceOffer => _telephonyLookup?.LicenceOffer;
    public bool HasTelephonyLicenceOffer => TelephonyLicenceOffer is not null;

    /// <summary>The licence needs a usage location, and that comes from a usable site.</summary>
    public bool CanAssignTelephonyLicence =>
        TelephonyLicenceOffer is not null && SelectedTelephonyLocation is { } site && _telephony.CheckUsageLocation(site) is null;
    public bool CanPickTelephonyNumber => HasTelephonyLookup && !HasTelephonyBlocker && SelectedTelephonyLocation is not null;
    public bool CanAssignTelephonyNumber => CanPickTelephonyNumber && SelectedTelephonyDid is not null && TelephonyPlan.Length > 0;
    public bool HasLastCreatedRoom => LastCreatedRoomAddress.Length > 0;

    partial void OnLastCreatedRoomAddressChanged(string value) => OnPropertyChanged(nameof(HasLastCreatedRoom));

    partial void OnTelephonyBlockerChanged(string value)
    {
        OnPropertyChanged(nameof(HasTelephonyBlocker));
        NotifyTelephonyGates();
    }

    partial void OnTelephonyPlanChanged(string value) => NotifyTelephonyGates();

    private void NotifyTelephonyGates()
    {
        OnPropertyChanged(nameof(HasTelephonyLookup));
        OnPropertyChanged(nameof(CanPickTelephonyNumber));
        OnPropertyChanged(nameof(CanAssignTelephonyNumber));
        OnPropertyChanged(nameof(TelephonyLicenceOffer));
        OnPropertyChanged(nameof(HasTelephonyLicenceOffer));
        OnPropertyChanged(nameof(CanAssignTelephonyLicence));
    }

    /// <summary>Called when the grid selection changes - nothing found for the previous room applies any more.</summary>
    private void ResetTelephony()
    {
        _telephonyLookup = null;
        TelephonyAccountSummary = "";
        TelephonyLicenceSummary = "";
        TelephonyBlocker = "";
        TelephonyWarnings = "";
        TelephonyResult = "";
        SelectedTelephonyLocation = null;
        TelephonyRanges.Clear();
        SelectedTelephonyRange = null;
        NotifyTelephonyGates();
    }

    // ----- step 1: look the room up -----

    [RelayCommand]
    private async Task LookUpTelephonyAsync()
    {
        if (SelectedRoom is null)
        {
            ManageStatus("Select a room in the grid first.", AlertSeverity.Error);
            return;
        }

        if (string.Equals(SelectedRoom.ResourceKind, "Equipment", StringComparison.OrdinalIgnoreCase))
        {
            ManageStatus("Only rooms can be set up for telephony, not equipment mailboxes.", AlertSeverity.Error);
            return;
        }

        var room = SelectedRoom;
        IsManageBusy = true;
        ResetTelephony();
        ManageStatus($"Reading {room.PrimarySmtpAddress} from Entra and Teams...");
        Transcript.BeginOperation($"Telephony lookup: {room.PrimarySmtpAddress}");

        try
        {
            if (TelephonyLocations.Count == 0 && _telephony.IsConfigured)
            {
                foreach (var location in await _telephony.GetLocationsAsync())
                    TelephonyLocations.Add(location);
            }

            var lookup = await _telephony.LookUpAsync(room.PrimarySmtpAddress, room.SiteCode);
            if (!ReferenceEquals(room, SelectedRoom))
                return; // the operator moved on while this ran

            if (!lookup.Succeeded)
            {
                ManageStatus(lookup.ErrorMessage ?? "Lookup failed.", AlertSeverity.Error);
                return;
            }

            _telephonyLookup = lookup;
            TelephonyAccountSummary = $"{lookup.DisplayName} ({lookup.UserPrincipalName})"
                                      + $" · usage location {(lookup.CurrentUsageLocation.Length == 0 ? "not set" : lookup.CurrentUsageLocation)}"
                                      + (lookup.AccountEnabled ? "" : " · account disabled");
            TelephonyLicenceSummary = lookup.License.Summary;
            TelephonyWarnings = string.Join("\n", lookup.Warnings);
            TelephonyBlocker = lookup.Blocker ?? "";

            // Pre-select the site the room's code points at; the operator can still change it.
            SelectedTelephonyLocation = lookup.SuggestedLocation is { } suggested
                ? TelephonyLocations.FirstOrDefault(l => l.Id == suggested.Id) ?? suggested
                : null;

            ManageStatus(TelephonyBlocker.Length > 0
                    ? "This room cannot be given a number right now - see the Telephony tab."
                    : SelectedTelephonyLocation is null
                        ? "Room found. Pick its site, then a number."
                        : "Room found. Pick a number block and search for a free number.",
                TelephonyBlocker.Length > 0 ? AlertSeverity.Warning : AlertSeverity.Info);
        }
        catch (Exception ex)
        {
            ManageStatus($"Lookup failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
            NotifyTelephonyGates();
        }
    }

    // ----- optional: licence, only on request -----

    /// <summary>
    /// Offered when the room has no Teams Phone and a Teams Rooms licence is free. Never automatic:
    /// licences cost money, so the operator is asked and has to confirm (user decision 2026-10-15).
    /// </summary>
    [RelayCommand]
    private async Task AssignTelephonyLicenceAsync()
    {
        if (string.IsNullOrWhiteSpace(ManageTaskNumber))
        {
            ManageStatus("Enter the ticket/task number authorizing this change.", AlertSeverity.Error);
            return;
        }

        if (_telephonyLookup is not { LicenceOffer: { } offer } lookup || SelectedTelephonyLocation is not { } site || !CanAssignTelephonyLicence)
        {
            ManageStatus("Look the room up and pick its site first - the licence needs the site's country as usage location.", AlertSeverity.Error);
            return;
        }

        var usage = UsageLocationResolver.Resolve(site.Country, site.Name);
        if (!_dialogs.Confirm(
                $"Assign {offer.DisplayName} to {lookup.DisplayName}?",
                $"One of the tenant's licences is used for this room ({offer.StockText}). It includes Teams Phone, so the room can then get a number.",
                "Assign licence",
                $"{lookup.UserPrincipalName}\n\nUsage location: "
                + (string.Equals(usage.Code, lookup.CurrentUsageLocation, StringComparison.OrdinalIgnoreCase)
                    ? $"{usage.Code} (unchanged)"
                    : $"{(lookup.CurrentUsageLocation.Length == 0 ? "(not set)" : lookup.CurrentUsageLocation)} -> {usage.Code}, the country of {site.Name}")
                + "\n\nTeams usually takes a few minutes to pick the licence up before a number can be assigned."))
            return;

        IsManageBusy = true;
        Transcript.BeginOperation($"Room licence: {lookup.UserPrincipalName}");
        var licensed = false;
        try
        {
            var result = await _telephony.AssignLicenceAsync(new RoomLicenceRequest
            {
                TaskNumber = ManageTaskNumber.Trim(),
                UserPrincipalName = lookup.UserPrincipalName,
                CurrentUsageLocation = lookup.CurrentUsageLocation,
                Location = site,
                Offer = offer,
            });

            licensed = result.Succeeded;
            ManageStatus(result.Succeeded
                    ? $"{offer.DisplayName} assigned to {lookup.DisplayName}. Reading the room again..."
                    : $"The licence could not be assigned: {result.ErrorMessage}",
                result.Succeeded ? AlertSeverity.Success : AlertSeverity.Error);
        }
        catch (Exception ex)
        {
            ManageStatus($"The licence could not be assigned: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }

        if (!licensed)
            return;

        // Read the room again, so the licence shows and the number steps unlock. Keep the site the
        // operator picked if it was not the suggested one.
        await LookUpTelephonyAsync();
        if (_telephonyLookup is not null && SelectedTelephonyLocation?.Id != site.Id)
            SelectedTelephonyLocation = TelephonyLocations.FirstOrDefault(l => l.Id == site.Id) ?? site;
    }

    // ----- step 2: site, block, number -----

    /// <summary>The site decides the blocks on offer, the policy names and the usage location - reload all of it.</summary>
    partial void OnSelectedTelephonyLocationChanged(UcLocation? value)
    {
        TelephonyRanges.Clear();
        SelectedTelephonyRange = null;
        NotifyTelephonyGates();

        if (_telephonyLookup is null)
            return;

        // A site problem blocks the run, but only as long as that site is the chosen one.
        TelephonyBlocker = _telephonyLookup.Blocker ?? (value is null ? null : _telephony.CheckLocation(value)) ?? "";
        if (value is null)
            return;
        if (TelephonyBlocker.Length == 0)
            _ = LoadTelephonyRangesAsync(value);
    }

    /// <summary>The ticket of the last creation, carried over when Create jumps to the Telephony tab.</summary>
    private string _lastCreatedTaskNumber = "";

    /// <summary>
    /// The creation's requester, carried over too: when the telephony failed during Create, their mail
    /// was held back and goes out from the Telephony tab instead.
    /// </summary>
    private string _lastCreatedRequester = "";

    /// <summary>
    /// Rooms created in this session whose account got a password here - their later telephony mail
    /// points the requester at the admin for it, as the creation mail would have.
    /// </summary>
    private readonly HashSet<string> _roomsWithNewPassword = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The new room account's password, shown once to copy; never stored.</summary>
    [ObservableProperty] private string _createdPassword = "";
    [ObservableProperty] private string _createdAccountUpn = "";
    [ObservableProperty] private string _createdDisplayName = "";

    /// <summary>The number the new room got in the same run; empty when telephony was not asked for or failed.</summary>
    [ObservableProperty] private string _createdPhoneNumber = "";

    public bool HasCreatedPassword => CreatedPassword.Length > 0;

    /// <summary>The result card shows as soon as there is a password or a number to copy.</summary>
    public bool HasCreatedAccountCard => CreatedPassword.Length > 0 || CreatedPhoneNumber.Length > 0;

    partial void OnCreatedPasswordChanged(string value)
    {
        OnPropertyChanged(nameof(HasCreatedPassword));
        OnPropertyChanged(nameof(HasCreatedAccountCard));
    }

    partial void OnCreatedPhoneNumberChanged(string value) => OnPropertyChanged(nameof(HasCreatedAccountCard));

    [RelayCommand]
    private void CopyCreatedPhoneNumber() => CopyToClipboard(CreatedPhoneNumber, "Phone number");

    [RelayCommand]
    private void CopyCreatedPassword() => CopyToClipboard(CreatedPassword, "Password");

    [RelayCommand]
    private void CopyCreatedUpn() => CopyToClipboard(CreatedAccountUpn, "Sign-in name");

    [RelayCommand]
    private void CopyCreatedDisplayName() => CopyToClipboard(CreatedDisplayName, "Display name");

    private void CopyToClipboard(string value, string what)
    {
        if (value.Length == 0)
            return;

        try
        {
            System.Windows.Clipboard.SetText(value);
            Status($"{what} copied to the clipboard.", AlertSeverity.Success);
        }
        catch (Exception ex)
        {
            // The clipboard can be held open by another process; the text box still allows Ctrl+C.
            Status($"Could not copy - select the text and press Ctrl+C instead ({ex.Message}).", AlertSeverity.Warning);
        }
    }

    private void RememberCreatedRoom(RoomResourceCreationResult result, ResourceKind kind, string taskNumber)
    {
        // The jump to the Telephony tab is only offered when there is something left to do there:
        // telephony was not asked for, or it failed.
        var telephonyDone = result.Telephony is { Succeeded: true };
        LastCreatedRoomAddress = result.Succeeded && kind == ResourceKind.Room && !telephonyDone ? result.PrimarySmtpAddress ?? "" : "";
        _lastCreatedTaskNumber = taskNumber;
        _lastCreatedRequester = telephonyDone ? "" : RequesterIdentity.Trim();
        if (result.Succeeded && result.Password is not null && result.PrimarySmtpAddress is { } address)
            _roomsWithNewPassword.Add(address);
    }

    private async Task LoadTelephonyRangesAsync(UcLocation location)
    {
        try
        {
            var ranges = await _telephony.GetDidRangesAsync(location.LocationCode);
            if (!ReferenceEquals(location, SelectedTelephonyLocation))
                return;

            foreach (var range in ranges)
                TelephonyRanges.Add(range);

            if (TelephonyRanges.Count == 0)
                ManageStatus($"No usable number blocks are configured for {location.LocationCode}.", AlertSeverity.Warning);
            else if (TelephonyRanges.Count == 1)
                SelectedTelephonyRange = TelephonyRanges[0];
        }
        catch (Exception ex)
        {
            ManageStatus($"Could not load number blocks: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
    }

    /// <summary>Numbers found in another block would be assigned with that block's type - drop them.</summary>
    partial void OnSelectedTelephonyRangeChanged(UcDidRange? value)
    {
        TelephonyFreeDids.Clear();
        SelectedTelephonyDid = null;
    }

    partial void OnSelectedTelephonyDidChanged(FreeDid? value)
    {
        TelephonyPlan = value is null || BuildTelephonyRequest() is not { } request ? "" : DescribeTelephonyPlan(request);
        NotifyTelephonyGates();
    }

    [RelayCommand]
    private async Task FindTelephonyNumbersAsync()
    {
        if (SelectedTelephonyRange is not { } range)
        {
            ManageStatus("Pick a number block first.", AlertSeverity.Error);
            return;
        }

        if (range.Type == TeamsNumberType.Unknown)
        {
            ManageStatus("This block has no number type (Direct Routing or Operator Connect) in the telephony database, "
                         + "so its numbers are not handed out. Have the type set on the range first.", AlertSeverity.Error);
            return;
        }

        IsManageBusy = true;
        TelephonyFreeDids.Clear();
        SelectedTelephonyDid = null;
        try
        {
            foreach (var did in await _telephony.FindFreeDidsAsync(range))
                TelephonyFreeDids.Add(did);

            if (TelephonyFreeDids.Count == 0)
            {
                ManageStatus("That block is fully allocated. Pick a different one.", AlertSeverity.Error);
                return;
            }

            // Random rather than lowest-first, as in New Hire: sequential numbers are guessable.
            SelectedTelephonyDid = TelephonyFreeDids[Random.Shared.Next(TelephonyFreeDids.Count)];
            ManageStatus($"{TelephonyFreeDids.Count} number(s) free. {SelectedTelephonyDid.E164} is proposed - change it if you like.",
                AlertSeverity.Success);
        }
        catch (Exception ex)
        {
            ManageStatus($"Number search failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    private RoomTelephonyRequest? BuildTelephonyRequest() =>
        _telephonyLookup is { } lookup && SelectedTelephonyLocation is { } location
        && SelectedTelephonyRange is { } range && SelectedTelephonyDid is { } did
            ? new RoomTelephonyRequest
            {
                TaskNumber = ManageTaskNumber.Trim(),
                UserPrincipalName = lookup.UserPrincipalName,
                DisplayName = lookup.DisplayName,
                CurrentUsageLocation = lookup.CurrentUsageLocation,
                Location = location,
                Range = range,
                Did = did.Did,
            }
            : null;

    /// <summary>What Assign will do, shown before it does it.</summary>
    private static string DescribeTelephonyPlan(RoomTelephonyRequest request)
    {
        var usage = request.TargetUsageLocation;
        var lines = new List<string>
        {
            $"Number:  {request.PhoneNumber} ({request.Range.TypeLabel})",
            $"Policies:  {PhonePolicyPlan.Describe(request.Range.Type, request.PolicyName)}",
            request.ChangesUsageLocation
                ? $"Usage location:  {(request.CurrentUsageLocation.Length == 0 ? "(not set)" : request.CurrentUsageLocation)} -> {usage.Code} ({usage.Note})"
                : $"Usage location:  {usage.Code} (unchanged)",
            "Inventory:  recorded in the telephony database as a Common Area Phone, as the import classifies room accounts.",
        };
        if (usage.NeedsCheck)
            lines.Add($"Check the country: {usage.Note}");
        return string.Join("\n", lines);
    }

    // ----- step 3: assign -----

    [RelayCommand]
    private async Task AssignTelephonyNumberAsync()
    {
        if (string.IsNullOrWhiteSpace(ManageTaskNumber))
        {
            ManageStatus("Enter the ticket/task number authorizing this change.", AlertSeverity.Error);
            return;
        }

        if (!CanAssignTelephonyNumber || BuildTelephonyRequest() is not { } request)
        {
            ManageStatus("Look the room up and pick a number first.", AlertSeverity.Error);
            return;
        }

        var room = SelectedRoom; // the grid selection may change while this runs
        IsManageBusy = true;
        TelephonyResult = "";
        Transcript.BeginOperation($"Room telephony: {request.UserPrincipalName} on {request.PhoneNumber}");
        try
        {
            var result = await _telephony.ConfigureAsync(request, msg => ManageStatusMessage = msg);

            var lines = result.Steps.Select(s => s.Succeeded
                ? $"OK      {s.Name}: {s.Detail}"
                : $"FAILED  {s.Name}: {s.Detail} - {s.Error}").ToList();
            if (result.VoiceState is { } state)
                lines.Add($"Teams reports: line {(state.LineUri.Length == 0 ? "(none)" : state.LineUri)}, enterprise voice "
                          + $"{(state.EnterpriseVoiceEnabled ? "on" : "off")}, voice routing "
                          + $"{(state.OnlineVoiceRoutingPolicy.Length == 0 ? "(global)" : state.OnlineVoiceRoutingPolicy)}, dial plan "
                          + $"{(state.TenantDialPlan.Length == 0 ? "(global)" : state.TenantDialPlan)}.");

            TelephonyResult = string.Join("\n", lines);
            TelephonyResultSeverity = !result.Succeeded ? AlertSeverity.Error
                : result.HasFailedSteps ? AlertSeverity.Warning
                : AlertSeverity.Success;

            if (!result.Succeeded)
            {
                ManageStatus($"{request.DisplayName} could not be given a number: {result.ErrorMessage}", AlertSeverity.Error);
                return;
            }

            // One mail to the requester, with the room's details and the number - also the one that
            // was held back when the telephony failed during Create.
            string? mailWarning = null;
            if (room is not null && ManageRequesterIdentity.Trim().Length > 0)
            {
                ManageStatusMessage = "Sending the requester the room's details...";
                mailWarning = await _rooms.SendTelephonyReadyMailAsync(room, ManageRequesterIdentity.Trim(), request.TaskNumber,
                    request.PhoneNumber, _roomsWithNewPassword.Contains(room.PrimarySmtpAddress));
            }

            var warn = result.HasFailedSteps || mailWarning is not null;
            ManageStatus((result.HasFailedSteps
                    ? $"{request.DisplayName} holds {request.PhoneNumber}, but some steps failed - see the Telephony tab."
                    : $"{request.DisplayName} is set up for telephony on {request.PhoneNumber} ({request.Range.TypeLabel}).")
                + (mailWarning is null ? "" : $" WARNING: {mailWarning}"),
                warn ? AlertSeverity.Warning : AlertSeverity.Success);

            // The number is gone from the pool; a second click must not offer it again.
            _telephonyLookup = null;
            TelephonyFreeDids.Clear();
            SelectedTelephonyDid = null;
        }
        catch (Exception ex)
        {
            ManageStatus($"Telephony setup failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
            NotifyTelephonyGates();
        }
    }

    // ----- jump from Create -----

    /// <summary>Opens the room just created on the Telephony tab, so the ticket can be finished there once it is licensed.</summary>
    [RelayCommand]
    private async Task SetUpTelephonyForCreatedRoomAsync()
    {
        var address = LastCreatedRoomAddress;
        if (address.Length == 0)
            return;

        await RefreshOverviewAsync();
        var row = Rooms.FirstOrDefault(r => string.Equals(r.PrimarySmtpAddress, address, StringComparison.OrdinalIgnoreCase));
        if (row is null)
        {
            Status($"{address} is not in the overview yet. Refresh it on the first tab in a moment.", AlertSeverity.Warning);
            return;
        }

        OverviewFilterText = "";
        SelectedRoom = row;
        ManageTaskNumber = _lastCreatedTaskNumber;
        ManageRequesterIdentity = _lastCreatedRequester;
        SelectedTabIndex = 0;
        ManageTabIndex = TelephonyTabIndex;
        LastCreatedRoomAddress = "";
    }
}
