using System.Globalization;
using System.Text.Json;
using M365Manager.Core.M365;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Teams;
using M365Manager.Core.Telephony;
using M365Manager.Data.Logging;
using M365Manager.Data.NewHire;

namespace M365Manager.Core.RoomResources;

/// <summary>
/// Sets a room account up for Teams telephony: usage location from the site's country, a free
/// number from the site's blocks (Direct Routing or Operator Connect, as the block says), the
/// site's voice policies, and a row in the telephony inventory.
///
/// Built after Switchboard's PhoneSetupService for common area phones, which is the same kind of
/// account - a cloud-only resource with Teams Phone, no on-premises AD object. So, unlike New Hire,
/// nothing is written to AD, and the policies follow the common area phone rule: Direct Routing all
/// four, Operator Connect the dial plan only (<see cref="PhonePolicyPlan"/>).
///
/// The licence is only ever assigned on request: when the room has no Teams Phone and the tenant has
/// a free Teams Rooms licence that carries it, the lookup offers it, and <see cref="AssignLicenceAsync"/>
/// runs only after the operator confirmed (decided by the user 2026-10-15 - licences cost money).
/// </summary>
public interface IRoomTelephonyService
{
    /// <summary>True once the telephony database (uchelper) is configured under Settings > New Hire.</summary>
    bool IsConfigured { get; }

    /// <summary>The sites available in the picker, ordered by name.</summary>
    Task<IReadOnlyList<UcLocation>> GetLocationsAsync(CancellationToken ct = default);

    /// <summary>
    /// Resolves the room account and works out whether it can be given a number. Never throws for a
    /// room-shaped problem (no licence, already has a number) - those come back as the blocker.
    /// </summary>
    /// <param name="roomAddress">Primary SMTP address of the room mailbox.</param>
    /// <param name="siteCode">The room's site code ("NBK"), used to suggest the matching site.</param>
    Task<RoomTelephonyLookup> LookUpAsync(string roomAddress, string siteCode, CancellationToken ct = default);

    /// <summary>Why a site cannot be used for a room number, or null when it can.</summary>
    string? CheckLocation(UcLocation location);

    /// <summary>Why a site's country cannot serve as usage location (all a licence needs), or null.</summary>
    string? CheckUsageLocation(UcLocation location);

    Task<IReadOnlyList<UcDidRange>> GetDidRangesAsync(string locationCode, CancellationToken ct = default);
    Task<IReadOnlyList<FreeDid>> FindFreeDidsAsync(UcDidRange range, CancellationToken ct = default);

    /// <summary>
    /// The tenant's Teams Rooms licence with Teams Phone and the most free units, or null when the
    /// tenant has none. For the Create form, which asks about the licence up front.
    /// </summary>
    Task<RoomLicenceOffer?> GetLicenceOfferAsync(CancellationToken ct = default);

    /// <summary>
    /// Assigns the offered Teams Rooms licence. Sets the usage location to the site's country first -
    /// Entra refuses a licence for an account without one. Only call after the operator confirmed.
    /// </summary>
    Task<RoomTelephonyResult> AssignLicenceAsync(RoomLicenceRequest request, CancellationToken ct = default);

    /// <summary>Applies everything. Failed steps are reported in the result, not thrown.</summary>
    Task<RoomTelephonyResult> ConfigureAsync(RoomTelephonyRequest request, Action<string>? onProgress = null, CancellationToken ct = default);
}

public sealed class RoomTelephonyService : IRoomTelephonyService
{
    /// <summary>
    /// How long to keep retrying while Teams does not accept the account yet - for all Teams steps
    /// together. A licence assigned minutes before reaches the number service with a delay.
    /// </summary>
    public static readonly TimeSpan ProvisioningWait = TimeSpan.FromMinutes(10);
    private static readonly TimeSpan RetryInterval = TimeSpan.FromSeconds(30);

    /// <summary>
    /// What Teams answers while an account has not reached the policy and number services yet.
    /// These are retried; anything else fails at once. Taken from Switchboard, where each was seen live.
    /// </summary>
    private static readonly string[] NotReadyMarkers =
    {
        "does not exist",
        "doesn't exist",
        "not found",
        "could not be found",
        "couldn't be found",
        "cannot find",
        "couldn't find",
        "not provisioned",
        "ManagementObjectNotFound",
        // A licence assigned shortly before: "User '...' lacks appropriate licenses to assign a
        // 'DirectRouting' number". After the deadline the message is shown as is.
        "licens",
        // Same lag, worded differently - seen live on a new room 2026-10-15, seconds after its
        // Teams Rooms Pro licence: "User with AccountType 'IneligibleUser' does not support
        // assigning numbers online." Get-CsOnlineUser showed AssignedPlan {} at that moment.
        "IneligibleUser",
        "does not support assigning numbers",
    };

    /// <summary>What Import-UCData.ps1 writes for a cloud account with a number and no app owner - keeps the row stable across imports.</summary>
    private const string EndpointType = "Common Area Phone";

    /// <summary>What the import writes as Office for accounts without an on-premises AD user.</summary>
    private const string CloudOnlyOffice = "Cloud Only Identity";

    private readonly INewHireRepository _repository;
    private readonly ITeamsPowerShellService _teams;
    private readonly GraphRestClient _graph;
    private readonly IPowerShellTranscript _transcript;
    private readonly IM365AuthService _auth;
    private readonly ILogService _log;
    private readonly IRoomAccountService _accounts;

    public RoomTelephonyService(
        INewHireRepository repository,
        ITeamsPowerShellService teams,
        GraphRestClient graph,
        IPowerShellTranscript transcript,
        IM365AuthService auth,
        ILogService log,
        IRoomAccountService accounts)
    {
        _repository = repository;
        _teams = teams;
        _graph = graph;
        _transcript = transcript;
        _auth = auth;
        _log = log;
        _accounts = accounts;
    }

    public bool IsConfigured => _repository.IsConfigured;

    public Task<IReadOnlyList<UcLocation>> GetLocationsAsync(CancellationToken ct = default) =>
        _repository.GetLocationsAsync(ct);

    public Task<IReadOnlyList<UcDidRange>> GetDidRangesAsync(string locationCode, CancellationToken ct = default) =>
        _repository.GetDidRangesAsync(locationCode, ct);

    public Task<IReadOnlyList<FreeDid>> FindFreeDidsAsync(UcDidRange range, CancellationToken ct = default) =>
        _repository.FindFreeDidsAsync(range, ct);

    // ----- lookup -----

    public async Task<RoomTelephonyLookup> LookUpAsync(string roomAddress, string siteCode, CancellationToken ct = default)
    {
        var address = (roomAddress ?? "").Trim();
        if (address.Length == 0)
            return RoomTelephonyLookup.Failed("Select a room first.");

        if (!_repository.IsConfigured)
            return RoomTelephonyLookup.Failed(
                "The telephony database is not configured. Open Settings > New Hire and enter its name.");

        if (!_teams.IsConnected)
            return RoomTelephonyLookup.Failed(
                "Not connected to Microsoft Teams. Restart the app to sign in, or check the connection status on the dashboard.");

        JsonElement account;
        try
        {
            account = await FindAccountAsync(address, ct)
                      ?? throw new InvalidOperationException($"No Entra account found for {address}.");
        }
        catch (Exception ex)
        {
            return RoomTelephonyLookup.Failed($"The room account could not be read from Entra: {ex.Message}");
        }

        var upn = TeamsLicenseReader.Text(account, "userPrincipalName");
        var displayName = TeamsLicenseReader.Text(account, "displayName");
        var usageLocation = TeamsLicenseReader.Text(account, "usageLocation");
        var enabled = account.TryGetProperty("accountEnabled", out var e) && e.ValueKind == JsonValueKind.True;

        var license = await TeamsLicenseReader.ReadAsync(_graph, upn, ct);

        TeamsVoiceState? voice = null;
        string? teamsError = null;
        try
        {
            voice = await _teams.TryGetVoiceStateAsync(upn, ct);
        }
        catch (Exception ex)
        {
            teamsError = ex.Message;
        }

        // The import records cloud accounts under their UPN prefix.
        string? inventoryNumber = null;
        try
        {
            inventoryNumber = (await _repository.GetEmployeeStateAsync(upn.Split('@')[0], ct)).AssignedNumber;
        }
        catch (Exception ex)
        {
            // EF hides the real cause in the InnerException.
            return RoomTelephonyLookup.Failed($"The telephony database could not be read: {(ex.InnerException ?? ex).Message}");
        }

        var teamsNumber = string.IsNullOrWhiteSpace(voice?.LineUri) ? null : voice!.LineUri;
        var assigned = teamsNumber ?? inventoryNumber;

        UcLocation? suggested = null;
        var code = (siteCode ?? "").Trim();
        if (code.Length > 0)
        {
            var locations = await _repository.GetLocationsAsync(ct);
            suggested = locations.FirstOrDefault(l => string.Equals(l.LocationCode.Trim(), code, StringComparison.OrdinalIgnoreCase));
        }

        var warnings = new List<string>();
        if (!license.Checked)
            warnings.Add($"The licence could not be read from Entra ({license.Error}). If Teams Phone is missing, assigning the number will fail.");
        else if (license.TeamsPhonePending)
            warnings.Add("The Teams Phone licence is assigned but Microsoft has not finished provisioning it. Assigning the number may fail for a few more minutes.");
        if (teamsError is not null)
            warnings.Add($"Teams could not be asked about the account ({teamsError}).");
        else if (voice is null && !license.DefinitelyUnlicensed)
            warnings.Add("Teams does not know the account yet - usual for a few minutes after the licence was assigned. "
                         + $"The run keeps retrying for up to {ProvisioningWait.TotalMinutes:0} minutes.");
        else if (voice is { IsIneligibleForNumbers: true } && !license.DefinitelyUnlicensed)
            warnings.Add("Teams has not picked up the licence yet (AccountType 'IneligibleUser') and refuses numbers until it has. "
                         + $"This can take a while after the licence was assigned; the run keeps retrying for up to {ProvisioningWait.TotalMinutes:0} minutes. "
                         + "If it still fails, look up again later.");
        if (!enabled)
            warnings.Add("The account is disabled in Entra (New-Mailbox -Room creates it that way). The number works regardless, "
                         + "but a Teams Rooms device can only sign in once the account is enabled with a password.");
        if (code.Length > 0 && suggested is null)
            warnings.Add($"No site in the telephony database has the location code '{code}'. Pick the room's site below.");

        // Only looked for when it matters: an account that has Teams Phone needs no offer.
        RoomLicenceOffer? offer = null;
        string? stockProblem = null;
        if (license.DefinitelyUnlicensed)
        {
            try
            {
                offer = await FindLicenceOfferAsync(ct);
                stockProblem = offer is null
                    ? "The tenant has no Teams Rooms licence that includes Teams Phone."
                    : offer.Available == 0 ? $"No Teams Rooms licence is free ({offer.Consumed} of {offer.Enabled} {offer.DisplayName} in use)." : null;
                if (offer is { Available: 0 })
                    offer = null;
            }
            catch (Exception ex)
            {
                stockProblem = $"The tenant's licences could not be read ({ex.Message}).";
            }
        }

        return new RoomTelephonyLookup
        {
            Succeeded = true,
            UserPrincipalName = upn,
            DisplayName = displayName,
            CurrentUsageLocation = usageLocation,
            AccountEnabled = enabled,
            License = license,
            Voice = voice,
            AssignedNumber = assigned,
            SuggestedLocation = suggested,
            LicenceOffer = assigned is null ? offer : null,
            Warnings = warnings,
            Blocker = FindBlocker(displayName, assigned, teamsNumber is not null, license, offer, stockProblem),
        };
    }

    /// <summary>
    /// The tenant's Teams Rooms licence that carries Teams Phone (MCOEV), the one with the most free
    /// units. Picked by content, not by a fixed part number: that excludes Teams Rooms Basic (no Teams
    /// Phone) and survives the next licensing rename.
    /// </summary>
    public Task<RoomLicenceOffer?> GetLicenceOfferAsync(CancellationToken ct = default) => FindLicenceOfferAsync(ct);

    private async Task<RoomLicenceOffer?> FindLicenceOfferAsync(CancellationToken ct)
    {
        using var json = await _graph.GetAsync("/subscribedSkus?$select=skuId,skuPartNumber,prepaidUnits,consumedUnits,servicePlans", ct: ct);
        if (!json.RootElement.TryGetProperty("value", out var skus) || skus.ValueKind != JsonValueKind.Array)
            return null;

        var offers = new List<RoomLicenceOffer>();
        foreach (var sku in skus.EnumerateArray())
        {
            var part = TeamsLicenseReader.Text(sku, "skuPartNumber");
            if (!part.Contains("Teams_Rooms", StringComparison.OrdinalIgnoreCase))
                continue;

            var hasPhone = sku.TryGetProperty("servicePlans", out var plans) && plans.ValueKind == JsonValueKind.Array
                && plans.EnumerateArray().Any(p => string.Equals(TeamsLicenseReader.Text(p, "servicePlanName"),
                    TeamsLicenseReader.TeamsPhoneServicePlan, StringComparison.OrdinalIgnoreCase));
            if (!hasPhone || !Guid.TryParse(TeamsLicenseReader.Text(sku, "skuId"), out var skuId))
                continue;

            var enabled = sku.TryGetProperty("prepaidUnits", out var prepaid) && prepaid.TryGetProperty("enabled", out var e)
                && e.TryGetInt32(out var n) ? n : 0;
            var consumed = sku.TryGetProperty("consumedUnits", out var c) && c.TryGetInt32(out var m) ? m : 0;
            offers.Add(new RoomLicenceOffer(skuId, part, enabled, consumed));
        }

        return offers.OrderByDescending(o => o.Available).FirstOrDefault();
    }

    public async Task<RoomTelephonyResult> AssignLicenceAsync(RoomLicenceRequest request, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();
        var steps = new List<RoomTelephonyStep>();
        var upn = request.UserPrincipalName;

        if (string.IsNullOrWhiteSpace(request.TaskNumber))
            return await FailAsync(correlationId, request.TaskNumber, upn, "Enter the ticket/task number authorizing this change.");
        // Only the country matters for a licence - the site need not be live on Teams voice, since
        // the licence may be assigned without any telephony.
        if (CheckUsageLocation(request.Location) is { } locationProblem)
            return await FailAsync(correlationId, request.TaskNumber, upn, locationProblem);

        // Re-read right before assigning: the count on screen may be minutes old, and a licence
        // must never be taken that is not actually free.
        RoomLicenceOffer? offer;
        try
        {
            offer = await FindLicenceOfferAsync(ct);
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, request.TaskNumber, upn, $"The tenant's licences could not be read: {ex.Message}");
        }
        if (offer is null || offer.SkuId != request.Offer.SkuId)
            return await FailAsync(correlationId, request.TaskNumber, upn, $"{request.Offer.DisplayName} is no longer available in the tenant.");
        if (offer.Available == 0)
            return await FailAsync(correlationId, request.TaskNumber, upn,
                $"No {offer.DisplayName} licence is free any more ({offer.Consumed} of {offer.Enabled} in use).");

        var target = UsageLocationResolver.Resolve(request.Location.Country, request.Location.Name).Code!;
        await LogAsync(correlationId, "STAR", $"Room licence assignment has started: {offer.DisplayName} ({offer.StockText})",
            request.TaskNumber, upn, Severity.Info);

        if (!string.Equals(target, request.CurrentUsageLocation, StringComparison.OrdinalIgnoreCase))
        {
            var from = request.CurrentUsageLocation.Length == 0 ? "(not set)" : request.CurrentUsageLocation;
            steps.Add(await TryStepAsync("Usage location", $"{from} -> {target}", async () =>
            {
                using var _ = await _graph.PatchAsync($"/users/{Uri.EscapeDataString(upn)}", new { usageLocation = target }, ct);
            }));
        }
        else
        {
            steps.Add(RoomTelephonyStep.Ok("Usage location", $"{target} (unchanged)"));
        }

        if (steps[^1].Succeeded)
        {
            steps.Add(await TryStepAsync("Licence", offer.DisplayName, async () =>
            {
                using var _ = await _graph.PostAsync($"/users/{Uri.EscapeDataString(upn)}/assignLicense", new
                {
                    addLicenses = new[] { new { skuId = offer.SkuId, disabledPlans = Array.Empty<Guid>() } },
                    removeLicenses = Array.Empty<Guid>(),
                }, ct);
            }));
        }

        foreach (var step in steps)
            await LogAsync(correlationId, step.Succeeded ? "STEP" : "FAIL",
                step.Succeeded ? $"{step.Name}: {step.Detail}" : $"{step.Name} ({step.Detail}): {step.Error}",
                request.TaskNumber, upn, step.Succeeded ? Severity.Success : Severity.Error);

        var licensed = steps.Count == 2 && steps[1].Succeeded;
        await LogAsync(correlationId, licensed ? "DONE" : "WARN",
            licensed ? $"Assigned {offer.DisplayName} to {upn}" : "Room licence assignment failed",
            request.TaskNumber, upn, licensed ? Severity.Success : Severity.Warning);

        return new RoomTelephonyResult
        {
            Succeeded = licensed,
            ErrorMessage = steps.FirstOrDefault(s => !s.Succeeded)?.Error,
            Steps = steps,
            CorrelationId = correlationId,
        };
    }

    private Task<JsonElement?> FindAccountAsync(string address, CancellationToken ct) => _accounts.FindAsync(address, ct);

    private static string? FindBlocker(string displayName, string? assigned, bool inTeams, TeamsLicenseState license,
        RoomLicenceOffer? offer, string? stockProblem)
    {
        if (assigned is not null)
            return $"{displayName} already holds {assigned}" + (inTeams ? " in Teams" : " in the telephony inventory")
                   + ". Changing a room's number is not part of this setup - remove the old assignment first.";

        if (license.DefinitelyUnlicensed)
            return $"{displayName} has no licence that includes Teams Phone, so no number can be assigned. Entra reports: {license.Summary}. "
                   + (offer is not null
                       ? $"{offer.StockText} - pick the room's site and assign one below."
                       : $"{stockProblem} Get a Teams Rooms Pro licence assigned, wait a few minutes, then look up again.");

        return null;
    }

    public string? CheckUsageLocation(UcLocation location)
    {
        var usage = UsageLocationResolver.Resolve(location.Country, location.Name);
        return usage.IsResolved
            ? null
            : $"The usage location for '{location.Name}' cannot be worked out: {usage.Note} Fix the site's Country in the telephony database.";
    }

    public string? CheckLocation(UcLocation location)
    {
        if (location.TeamsVoiceEnabled == 2)
            return $"The site '{location.Name}' is not enabled for Teams enterprise voice.";
        if (location.TeamsVoiceEnabled != 1)
            return $"The site '{location.Name}' has no Teams voice status recorded. Have the site configured first.";
        if (string.IsNullOrWhiteSpace(location.LocationCode))
            return $"The site '{location.Name}' has no location code, so no policies can be granted.";

        var usage = UsageLocationResolver.Resolve(location.Country, location.Name);
        if (!usage.IsResolved)
            return $"The usage location for '{location.Name}' cannot be worked out: {usage.Note} Fix the site's Country in the telephony database.";

        return null;
    }

    // ----- apply -----

    public async Task<RoomTelephonyResult> ConfigureAsync(RoomTelephonyRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();
        var steps = new List<RoomTelephonyStep>();
        var upn = request.UserPrincipalName;

        if (string.IsNullOrWhiteSpace(request.TaskNumber))
            return await FailAsync(correlationId, request, "Enter the ticket/task number authorizing this change.");
        if (request.Range.Type == TeamsNumberType.Unknown)
            return await FailAsync(correlationId, request,
                $"The range of {request.PhoneNumber} has no number type (Direct Routing or Operator Connect) in the telephony database.");
        if (CheckLocation(request.Location) is { } locationProblem)
            return await FailAsync(correlationId, request, locationProblem);

        // Re-checked right before the change: the list on screen may be minutes old.
        try
        {
            if (!await _repository.IsDidFreeAsync(request.Did, ct))
                return await FailAsync(correlationId, request, $"{request.PhoneNumber} has been taken in the meantime - choose another number.");
            if (await _teams.GetNumberHolderAsync(request.PhoneNumber, ct) is { } holder)
                return await FailAsync(correlationId, request,
                    $"{request.PhoneNumber} is already assigned in Teams ({holder}) although the telephony database shows it as free.");
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, request, $"Could not check whether {request.PhoneNumber} is still free: {ex.Message}");
        }

        var typeName = TeamsNumberTypes.CmdletValue(request.Range.Type);
        var target = request.TargetUsageLocation.Code!;

        await LogAsync(correlationId, "STAR", "Room telephony setup has started", request, Severity.Info);
        await LogAsync(correlationId, "INPUT",
            $"Room={upn}; Site={request.Location.Name} ({request.Location.LocationCode}); Number={request.PhoneNumber}; "
            + $"Type={typeName}; UsageLocation={target}", request, Severity.Info);

        // --- usage location: Teams refuses a number for an account without one ---
        if (request.ChangesUsageLocation)
        {
            Report(onProgress, $"Setting the usage location to {target}...");
            var from = request.CurrentUsageLocation.Length == 0 ? "(not set)" : request.CurrentUsageLocation;
            steps.Add(await TryStepAsync("Usage location", $"{from} -> {target}", async () =>
            {
                using var _ = await _graph.PatchAsync($"/users/{Uri.EscapeDataString(upn)}", new { usageLocation = target }, ct);
            }));
        }
        else
        {
            steps.Add(RoomTelephonyStep.Ok("Usage location", $"{target} (unchanged)"));
        }

        var deadline = DateTimeOffset.UtcNow + ProvisioningWait;

        // --- number + Enterprise Voice; without it nothing below makes sense ---
        Report(onProgress, $"Assigning {request.PhoneNumber} ({request.Range.TypeLabel}) in Teams...");
        var numberStep = await TryStepAsync("Phone number", $"{request.PhoneNumber} ({request.Range.TypeLabel})",
            () => WhenReadyAsync(upn, deadline, () => _teams.AssignPhoneNumberAsync(upn, request.PhoneNumber, typeName, ct), ct));
        steps.Add(numberStep);

        if (numberStep.Succeeded)
        {
            // --- voice policies: Direct Routing all four, Operator Connect the dial plan only ---
            foreach (var (name, cmdlet) in request.VoicePolicies)
            {
                ct.ThrowIfCancellationRequested();
                Report(onProgress, $"Granting {name} '{request.PolicyName}'...");
                steps.Add(await TryStepAsync(name, request.PolicyName,
                    () => WhenReadyAsync(upn, deadline, () => _teams.GrantPolicyAsync(cmdlet, upn, request.PolicyName, ct), ct)));
            }

            // --- inventory: recorded whenever the number landed, even if a policy failed - otherwise
            // the number keeps showing as free and is handed out twice. Shaped like Import-UCData.ps1's row.
            string Granted(string cmdlet) => PhonePolicyPlan.Grants(request.Range.Type, cmdlet) ? request.PolicyName : "";
            Report(onProgress, "Recording the endpoint in the telephony database...");
            steps.Add(await TryStepAsync("Telephony inventory", request.PhoneNumber, () =>
                _repository.AddEndpointAsync(new UcEndpoint
                {
                    DisplayName = request.DisplayName,
                    SamAccountName = upn.Split('@')[0],
                    SipAddress = upn,
                    LineUri = request.LineUri,
                    Did = request.Did,
                    Type = EndpointType,
                    VoicePolicy = Granted(PhonePolicyPlan.VoiceRoutingCmdlet),
                    DialPlan = Granted(PhonePolicyPlan.DialPlanCmdlet),
                    EnterpriseVoiceEnabled = "True",
                    HostedVoiceMail = "False",
                    EmergencyCallingPolicy = Granted(PhonePolicyPlan.EmergencyCallingCmdlet),
                    EmergencyCallRoutingPolicy = Granted(PhonePolicyPlan.EmergencyCallRoutingCmdlet),
                    Office = CloudOnlyOffice,
                }, typeName, target, ct)));
        }

        // --- read back what Teams holds ---
        TeamsVoiceState? voiceState = null;
        try
        {
            Report(onProgress, "Reading the configuration back from Teams...");
            voiceState = await _teams.TryGetVoiceStateAsync(upn, ct);
        }
        catch (Exception ex)
        {
            await LogAsync(correlationId, "WARN", $"Could not read the Teams configuration back: {ex.Message}", request, Severity.Warning);
        }

        // Every step, not only the failed ones: the audit log is the only record of the run.
        foreach (var step in steps)
            await LogAsync(correlationId,
                step.Succeeded ? "STEP" : "FAIL",
                step.Succeeded ? $"{step.Name}: {step.Detail}" : $"{step.Name} ({step.Detail}): {step.Error}",
                request, step.Succeeded ? Severity.Success : Severity.Error);

        var failed = steps.Count(s => !s.Succeeded);
        await LogAsync(correlationId,
            failed == 0 ? "DONE" : "WARN",
            failed == 0
                ? $"Room telephony setup has completed: {upn} on {request.PhoneNumber} ({request.Range.TypeLabel})"
                : $"Room telephony setup finished with {failed} failed step(s)",
            request, failed == 0 ? Severity.Success : Severity.Warning);

        return new RoomTelephonyResult
        {
            Succeeded = numberStep.Succeeded,
            ErrorMessage = numberStep.Error,
            Steps = steps,
            VoiceState = voiceState,
            CorrelationId = correlationId,
        };
    }

    /// <summary>
    /// Runs a Teams command, retrying while Teams says the account is not there yet - until the
    /// deadline of the whole run.
    /// </summary>
    private async Task WhenReadyAsync(string upn, DateTimeOffset deadline, Func<Task> action, CancellationToken ct)
    {
        var attempts = 0;
        while (true)
        {
            attempts++;
            try
            {
                await action();
                return;
            }
            catch (Exception ex) when (IsNotReady(ex))
            {
                if (DateTimeOffset.UtcNow + RetryInterval >= deadline)
                    throw new InvalidOperationException(
                        $"Teams does not accept {upn} yet after {ProvisioningWait.TotalMinutes:0} minutes (last answer: {ex.Message}). "
                        + "Check the licence, then run the setup again later.", ex);

                _transcript.Write(TranscriptLineKind.Information, string.Create(CultureInfo.InvariantCulture,
                    $"Teams does not accept {upn} yet - trying again in {RetryInterval.TotalSeconds:0} s (attempt {attempts})."));
                await Task.Delay(RetryInterval, ct);
            }
        }
    }

    private static bool IsNotReady(Exception ex) =>
        NotReadyMarkers.Any(m => ex.Message.Contains(m, StringComparison.OrdinalIgnoreCase));

    private static async Task<RoomTelephonyStep> TryStepAsync(string name, string detail, Func<Task> action)
    {
        try
        {
            await action();
            return RoomTelephonyStep.Ok(name, detail);
        }
        catch (Exception ex)
        {
            return RoomTelephonyStep.Failed(name, detail, ex.Message);
        }
    }

    private static void Report(Action<string>? onProgress, string message) => onProgress?.Invoke(message);

    private Task<RoomTelephonyResult> FailAsync(Guid correlationId, RoomTelephonyRequest request, string error) =>
        FailAsync(correlationId, request.TaskNumber, request.UserPrincipalName, error);

    private async Task<RoomTelephonyResult> FailAsync(Guid correlationId, string taskNumber, string upn, string error)
    {
        await LogAsync(correlationId, "FAIL", error, taskNumber, upn, Severity.Error);
        return RoomTelephonyResult.Failed(correlationId, error);
    }

    private Task LogAsync(Guid correlationId, string eventCode, string message, RoomTelephonyRequest request, Severity severity) =>
        LogAsync(correlationId, eventCode, message, request.TaskNumber, request.UserPrincipalName, severity);

    private async Task LogAsync(Guid correlationId, string eventCode, string message, string taskNumber, string upn, Severity severity)
    {
        try
        {
            await _log.WriteAsync(new LogEntry
            {
                CorrelationId = correlationId,
                Area = "RoomResources",
                Action = "Telephony",
                EventCode = eventCode,
                Severity = severity,
                TargetObject = upn,
                TaskNumber = taskNumber,
                UserUpn = _auth.CurrentUser?.Upn,
                Message = message,
            });
        }
        catch
        {
            // Logging is best-effort - SQL may not be reachable.
        }
    }
}
