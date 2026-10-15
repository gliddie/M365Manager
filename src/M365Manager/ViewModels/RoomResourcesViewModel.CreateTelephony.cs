using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Controls;
using M365Manager.Core.RoomResources;
using M365Manager.Core.Telephony;
using M365Manager.Data.NewHire;

namespace M365Manager.ViewModels;

/// <summary>
/// The licence and telephony part of the Create form. Two independent choices made up front (user
/// requests 2026-10-15): a Teams Rooms licence - with or without a number - and a phone number. Both
/// run in the same Create, and the requester gets a single mail once everything is done.
///
/// They share the site picker: a licence needs the site's country as usage location; a number also
/// needs the site's blocks and policy names.
/// </summary>
public sealed partial class RoomResourcesViewModel
{
    [ObservableProperty] private bool _createWithTelephony;
    [ObservableProperty] private bool _createAssignLicence;
    [ObservableProperty] private UcLocation? _createTelephonyLocation;
    [ObservableProperty] private UcDidRange? _createTelephonyRange;
    [ObservableProperty] private FreeDid? _createTelephonyDid;
    [ObservableProperty] private string _createLicenceText = "";
    [ObservableProperty] private string _createTelephonySiteProblem = "";
    [ObservableProperty] private string _createTelephonyPlan = "";
    [ObservableProperty] private string _createTelephonyResult = "";
    [ObservableProperty] private AlertSeverity _createTelephonyResultSeverity = AlertSeverity.Info;

    public ObservableCollection<UcDidRange> CreateTelephonyRanges { get; } = new();
    public ObservableCollection<FreeDid> CreateTelephonyFreeDids { get; } = new();

    private RoomLicenceOffer? _createLicenceOffer;
    private bool _createLicenceOfferLoaded;

    /// <summary>The site code of the name as typed, from the last preview - used to suggest the site.</summary>
    private string _previewSiteCode = "";

    /// <summary>True once the operator picked a site by hand; then typing the name no longer changes it.</summary>
    private bool _createLocationPickedByHand;
    private bool _applyingSuggestedLocation;

    /// <summary>The site picker is needed by either option.</summary>
    public bool ShowCreateSite => CreateWithTelephony || CreateAssignLicence;

    public bool CanPickCreateNumber =>
        CreateWithTelephony && CreateTelephonyLocation is not null && CreateTelephonySiteProblem.Length == 0;

    partial void OnCreateTelephonySiteProblemChanged(string value) => OnPropertyChanged(nameof(CanPickCreateNumber));

    partial void OnCreateWithTelephonyChanged(bool value) => OnCreateOptionChanged();

    partial void OnCreateAssignLicenceChanged(bool value)
    {
        OnCreateOptionChanged();
        if (value && !_createLicenceOfferLoaded)
            _ = LoadCreateLicenceOfferAsync();
    }

    private void OnCreateOptionChanged()
    {
        OnPropertyChanged(nameof(ShowCreateSite));
        if (ShowCreateSite)
            _ = EnsureCreateLocationsAsync();
        RecheckCreateSite();
    }

    private async Task EnsureCreateLocationsAsync()
    {
        try
        {
            if (!_telephony.IsConfigured)
            {
                Status("The telephony database is not configured (it holds the sites). Open Settings > New Hire and enter its name.",
                    AlertSeverity.Error);
                CreateWithTelephony = false;
                CreateAssignLicence = false;
                return;
            }

            if (TelephonyLocations.Count == 0)
                foreach (var location in await _telephony.GetLocationsAsync())
                    TelephonyLocations.Add(location);

            SuggestCreateLocation();
        }
        catch (Exception ex)
        {
            Status($"Could not load the site list: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
    }

    private async Task LoadCreateLicenceOfferAsync()
    {
        CreateLicenceText = "Checking the tenant's Teams Rooms licences...";
        try
        {
            _createLicenceOffer = await _telephony.GetLicenceOfferAsync();
            _createLicenceOfferLoaded = true;

            if (_createLicenceOffer is { Available: > 0 } offer)
            {
                CreateLicenceText = $"{offer.StockText}. Asked again for confirmation when you click Create.";
            }
            else
            {
                CreateLicenceText = _createLicenceOffer is null
                    ? "The tenant has no Teams Rooms licence that includes Teams Phone."
                    : $"No Teams Rooms licence is free ({_createLicenceOffer.Consumed} of {_createLicenceOffer.Enabled} {_createLicenceOffer.DisplayName} in use).";
                CreateAssignLicence = false;
            }
        }
        catch (Exception ex)
        {
            CreateLicenceText = $"The licences could not be read: {ErrorText.Describe(ex)}";
            CreateAssignLicence = false;
        }
        finally
        {
            UpdateCreateTelephonyPlan();
        }
    }

    /// <summary>Called after every preview - the name's site code picks the site until the operator picks one.</summary>
    private void OnPreviewSiteCodeChanged(string siteCode)
    {
        _previewSiteCode = siteCode;
        if (ShowCreateSite)
            SuggestCreateLocation();
    }

    private void SuggestCreateLocation()
    {
        if (_createLocationPickedByHand || _previewSiteCode.Length == 0)
            return;

        var match = TelephonyLocations.FirstOrDefault(l =>
            string.Equals(l.LocationCode.Trim(), _previewSiteCode, StringComparison.OrdinalIgnoreCase));
        if (match is null || ReferenceEquals(match, CreateTelephonyLocation))
            return;

        _applyingSuggestedLocation = true;
        try
        {
            CreateTelephonyLocation = match;
        }
        finally
        {
            _applyingSuggestedLocation = false;
        }
    }

    partial void OnCreateTelephonyLocationChanged(UcLocation? value)
    {
        if (!_applyingSuggestedLocation && value is not null)
            _createLocationPickedByHand = true;

        RecheckCreateSite();
    }

    /// <summary>
    /// What the chosen site has to offer depends on the options: a licence needs only a country, a
    /// number also needs a site that is live on Teams voice. Reloads the blocks when a number is wanted.
    /// </summary>
    private void RecheckCreateSite()
    {
        CreateTelephonyRanges.Clear();
        CreateTelephonyRange = null;

        var site = CreateTelephonyLocation;
        CreateTelephonySiteProblem = site is null ? ""
            : CreateWithTelephony ? _telephony.CheckLocation(site) ?? ""
            : CreateAssignLicence ? _telephony.CheckUsageLocation(site) ?? ""
            : "";
        // Depends on the options and the site too, not only on the problem text - which stays "" for
        // a good site and so would not raise anything on its own.
        OnPropertyChanged(nameof(CanPickCreateNumber));
        UpdateCreateTelephonyPlan();

        if (site is not null && CreateWithTelephony && CreateTelephonySiteProblem.Length == 0)
            _ = LoadCreateRangesAsync(site);
    }

    private async Task LoadCreateRangesAsync(UcLocation location)
    {
        try
        {
            var ranges = await _telephony.GetDidRangesAsync(location.LocationCode);
            if (!ReferenceEquals(location, CreateTelephonyLocation) || !CreateWithTelephony)
                return;

            CreateTelephonyRanges.Clear();
            foreach (var range in ranges)
                CreateTelephonyRanges.Add(range);

            if (CreateTelephonyRanges.Count == 0)
                CreateTelephonySiteProblem = $"No usable number blocks are configured for {location.LocationCode}.";
            else if (CreateTelephonyRanges.Count == 1)
                CreateTelephonyRange = CreateTelephonyRanges[0];
        }
        catch (Exception ex)
        {
            CreateTelephonySiteProblem = $"Could not load number blocks: {ErrorText.Describe(ex)}";
        }
    }

    /// <summary>Numbers found in another block would be assigned with that block's type - drop them.</summary>
    partial void OnCreateTelephonyRangeChanged(UcDidRange? value)
    {
        CreateTelephonyFreeDids.Clear();
        CreateTelephonyDid = null;
    }

    partial void OnCreateTelephonyDidChanged(FreeDid? value) => UpdateCreateTelephonyPlan();

    [RelayCommand]
    private async Task FindCreateTelephonyNumbersAsync()
    {
        if (CreateTelephonyRange is not { } range)
        {
            Status("Pick a number block first.", AlertSeverity.Error);
            return;
        }

        if (range.Type == TeamsNumberType.Unknown)
        {
            Status("This block has no number type (Direct Routing or Operator Connect) in the telephony database, "
                   + "so its numbers are not handed out. Have the type set on the range first.", AlertSeverity.Error);
            return;
        }

        IsBusy = true;
        CreateTelephonyFreeDids.Clear();
        CreateTelephonyDid = null;
        try
        {
            foreach (var did in await _telephony.FindFreeDidsAsync(range))
                CreateTelephonyFreeDids.Add(did);

            if (CreateTelephonyFreeDids.Count == 0)
            {
                Status("That block is fully allocated. Pick a different one.", AlertSeverity.Error);
                return;
            }

            // Random rather than lowest-first, as in New Hire: sequential numbers are guessable.
            CreateTelephonyDid = CreateTelephonyFreeDids[Random.Shared.Next(CreateTelephonyFreeDids.Count)];
            Status($"{CreateTelephonyFreeDids.Count} number(s) free. {CreateTelephonyDid.E164} is proposed - change it if you like.",
                AlertSeverity.Success);
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

    /// <summary>The number the form asks for, or null when telephony is off or not complete yet.</summary>
    private RoomCreateTelephony? BuildCreateTelephony() =>
        CreateWithTelephony && Kind == ResourceKind.Room
        && CreateTelephonyLocation is { } location && CreateTelephonyRange is { } range && CreateTelephonyDid is { } did
            ? new RoomCreateTelephony { Location = location, Range = range, Did = did.Did }
            : null;

    /// <summary>The licence the form asks for, or null when unticked, none is free, or no site is picked.</summary>
    private RoomCreateLicence? BuildCreateLicence() =>
        CreateAssignLicence && Kind == ResourceKind.Room
        && _createLicenceOffer is { Available: > 0 } offer && CreateTelephonyLocation is { } location
            ? new RoomCreateLicence { Location = location, Offer = offer }
            : null;

    private void UpdateCreateTelephonyPlan()
    {
        var licence = BuildCreateLicence();
        var phone = BuildCreateTelephony();
        if (licence is null && phone is null)
        {
            CreateTelephonyPlan = "";
            return;
        }

        var site = (licence?.Location ?? phone!.Location)!;
        var usage = UsageLocationResolver.Resolve(site.Country, site.Name);
        var lines = new List<string>();
        if (licence is not null)
            lines.Add($"Licence:  {licence.Offer.DisplayName} - assigned after you confirm");
        if (phone is not null)
        {
            lines.Add($"Number:  {phone.PhoneNumber} ({phone.Range.TypeLabel})");
            lines.Add($"Policies:  {PhonePolicyPlan.Describe(phone.Range.Type, phone.Location.LocationCode)}");
            if (licence is null)
                lines.Add("Licence:  none assigned by the app - Teams only takes the number once the room has Teams Phone");
        }
        lines.Add($"Usage location:  {usage.Code} ({usage.Note})");
        lines.Add("The requester's e-mail goes out once all of this is done.");
        CreateTelephonyPlan = string.Join("\n", lines);
    }

    /// <summary>The Create click's extra checks: complete choices, and the licence confirm.</summary>
    /// <returns>False when the creation must not start.</returns>
    private bool ConfirmCreateTelephony()
    {
        if (Kind != ResourceKind.Room)
            return true;

        if (CreateWithTelephony && BuildCreateTelephony() is null)
        {
            Status("Telephony is ticked: pick the site, a number block and a number - or untick it.", AlertSeverity.Error);
            return false;
        }

        if (!CreateAssignLicence)
            return true;

        if (BuildCreateLicence() is not { } licence)
        {
            Status(_createLicenceOffer is { Available: > 0 }
                    ? "The licence is ticked: pick the site - its country becomes the usage location - or untick it."
                    : "The licence is ticked, but no Teams Rooms licence is free. Untick it.",
                AlertSeverity.Error);
            return false;
        }

        var usage = UsageLocationResolver.Resolve(licence.Location.Country, licence.Location.Name);
        return _dialogs.Confirm(
            $"Assign {licence.Offer.DisplayName} to the new room?",
            $"One of the tenant's licences is used for {PreviewDisplayName} ({licence.Offer.StockText}).",
            "Assign licence and create",
            $"{PreviewAddress}\n\nUsage location: {usage.Code}, the country of {licence.Location.Name}.\n\n"
            + "Cancel stops the creation - untick the licence box to create the room without it.");
    }

    /// <summary>The licence and telephony steps of the run, in order.</summary>
    private void ShowCreateTelephonyResult(RoomTelephonyResult? licence, RoomTelephonyResult? telephony)
    {
        var parts = new[] { licence, telephony }.Where(r => r is not null).Select(r => r!).ToList();
        if (parts.Count == 0)
        {
            CreateTelephonyResult = "";
            return;
        }

        var lines = new List<string>();
        foreach (var result in parts)
        {
            lines.AddRange(result.Steps.Select(s => s.Succeeded
                ? $"OK      {s.Name}: {s.Detail}"
                : $"FAILED  {s.Name}: {s.Detail} - {s.Error}"));
            if (result.Steps.Count == 0 && result.ErrorMessage is { } error)
                lines.Add($"FAILED  {error}");
        }

        CreateTelephonyResult = string.Join("\n", lines);
        CreateTelephonyResultSeverity = parts.Any(r => !r.Succeeded) ? AlertSeverity.Error
            : parts.Any(r => r.HasFailedSteps) ? AlertSeverity.Warning
            : AlertSeverity.Success;
    }

    private void ClearCreateTelephony()
    {
        CreateWithTelephony = false;
        CreateAssignLicence = false;
        CreateTelephonyLocation = null;
        _createLocationPickedByHand = false;
        _createLicenceOfferLoaded = false; // the stock changed with this creation
        CreateLicenceText = "";
        CreateTelephonySiteProblem = "";
        CreateTelephonyPlan = "";
    }
}
