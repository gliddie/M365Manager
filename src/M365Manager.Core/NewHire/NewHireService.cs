using System.Text.Json;
using M365Manager.Core.ActiveDirectory;
using M365Manager.Core.M365;
using M365Manager.Core.Notifications;
using M365Manager.Core.Settings;
using M365Manager.Core.Teams;
using M365Manager.Data.Logging;
using M365Manager.Data.NewHire;

namespace M365Manager.Core.NewHire;

/// <summary>See <see cref="INewHireService"/>.</summary>
public sealed class NewHireService : INewHireService
{
    private readonly INewHireRepository _repository;
    private readonly IActiveDirectoryService _directory;
    private readonly GraphRestClient _graph;
    private readonly ITeamsPowerShellService _teams;
    private readonly INotificationMailService _mail;
    private readonly ISettingsService _settings;
    private readonly IM365AuthService _auth;
    private readonly ILogService _log;

    public NewHireService(
        INewHireRepository repository,
        IActiveDirectoryService directory,
        GraphRestClient graph,
        ITeamsPowerShellService teams,
        INotificationMailService mail,
        ISettingsService settings,
        IM365AuthService auth,
        ILogService log)
    {
        _repository = repository;
        _directory = directory;
        _graph = graph;
        _teams = teams;
        _mail = mail;
        _settings = settings;
        _auth = auth;
        _log = log;
    }

    // ----- step 1: who is this, and can they be enabled -----

    public Task<IReadOnlyList<UcLocation>> GetLocationsAsync(CancellationToken ct = default) =>
        _repository.GetLocationsAsync(ct);

    public async Task<NewHireLookup> LookUpAsync(string identity, string? officeOverride = null, CancellationToken ct = default)
    {
        var value = (identity ?? "").Trim();
        if (value.Length == 0)
            return NewHireLookup.Failed("Enter an employee to look up.");

        if (!_repository.IsConfigured)
            return NewHireLookup.Failed(
                "The telephony database is not configured. Open Settings > New Hire and enter its name.");

        // A UPN or mail address is what people paste; AD is searched by SamAccountName.
        var sam = value.Contains('@') ? value[..value.IndexOf('@')] : value;

        AdUser? user;
        try
        {
            user = await _directory.FindUserAsync(sam, ct);
        }
        catch (Exception ex)
        {
            return NewHireLookup.Failed($"Active Directory could not be reached: {ex.Message}");
        }

        if (user is null)
            return NewHireLookup.Failed($"No Active Directory account found for '{sam}'.");

        var state = await _repository.GetEmployeeStateAsync(user.SamAccountName, ct);
        var license = await GetLicenseStateAsync(user.UserPrincipalName, ct);

        // The office attribute is regularly empty or wrong on a brand-new account, so the operator
        // can pick a site instead - that override wins.
        var officeName = string.IsNullOrWhiteSpace(officeOverride) ? user.Office : officeOverride.Trim();
        var location = officeName.Length == 0
            ? null
            : await _repository.FindLocationByNameAsync(officeName, ct);

        var warnings = new List<string>();

        if (!license.Checked)
            warnings.Add($"The licence could not be read from Entra ({license.Error}). "
                         + "Continuing without it - if the Teams Phone licence is missing, assigning the number will fail.");
        else if (license.TeamsPhonePending)
            warnings.Add("The Teams Phone licence is assigned but Microsoft has not finished provisioning it. "
                         + "Assigning the number may fail for another few minutes.");

        if (user.LineUri is not null)
            warnings.Add($"The AD account already carries a line URI ({user.LineUri}). Continuing will replace it.");

        if (string.IsNullOrWhiteSpace(user.Mail))
            warnings.Add("The AD account has no mail address, so no SIP address can be derived and no notification can be sent.");

        return new NewHireLookup
        {
            Succeeded = true,
            User = user,
            Location = location,
            State = state,
            License = license,
            Warnings = warnings,
            Blocker = FindBlocker(user, state, license, location, officeName),
        };
    }

    /// <summary>
    /// The service plan that actually decides whether someone can hold a phone number. Granted by
    /// E5, by Teams Phone Standard and by Teams Phone with Calling Plan - never by E3. Checking the
    /// service plan rather than the SKU keeps this correct across licensing renames.
    /// </summary>
    private const string TeamsPhoneServicePlan = "MCOEV";

    /// <summary>
    /// Reads the employee's licences from Entra. The telephony database has an <c>e3licensed</c>
    /// table filled by a nightly import, and reading it would be marginally faster - but it only
    /// models E3, which this tenant no longer uses, so it reports every E5 user as unlicensed.
    /// Entra is the source of truth; the import would have to be reworked before the table could be
    /// trusted again.
    ///
    /// A failure here is deliberately not fatal: an unreachable Graph should not stop a run that
    /// would otherwise succeed, and a genuinely missing licence still surfaces when Teams refuses
    /// the number.
    /// </summary>
    private async Task<TeamsLicenseState> GetLicenseStateAsync(string upn, CancellationToken ct)
    {
        try
        {
            var licenses = await _graph.GetAllPagesAsync(
                $"/users/{Uri.EscapeDataString(upn)}/licenseDetails", ct: ct);

            var skus = new List<string>();
            var provisioned = false;
            var pending = false;

            foreach (var license in licenses)
            {
                var sku = Text(license, "skuPartNumber");
                if (sku.Length > 0)
                    skus.Add(sku);

                if (!license.TryGetProperty("servicePlans", out var plans) || plans.ValueKind != JsonValueKind.Array)
                    continue;

                foreach (var plan in plans.EnumerateArray())
                {
                    if (!string.Equals(Text(plan, "servicePlanName"), TeamsPhoneServicePlan, StringComparison.OrdinalIgnoreCase))
                        continue;

                    // The same plan can appear under several SKUs; one provisioned copy is enough.
                    if (string.Equals(Text(plan, "provisioningStatus"), "Success", StringComparison.OrdinalIgnoreCase))
                        provisioned = true;
                    else
                        pending = true;
                }
            }

            return new TeamsLicenseState
            {
                Checked = true,
                HasTeamsPhone = provisioned,
                TeamsPhonePending = pending && !provisioned,
                Skus = skus,
            };
        }
        catch (Exception ex)
        {
            return new TeamsLicenseState { Checked = false, Error = ex.Message };
        }
    }

    private static string Text(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : "";

    /// <summary>
    /// The one reason this employee cannot be enabled, in the order that matters. The legacy tool
    /// raised each of these as a message box and then let the operator carry on anyway - which is how
    /// unlicensed accounts ended up half configured.
    /// </summary>
    private static string? FindBlocker(AdUser user, UcEmployeeState state, TeamsLicenseState license, UcLocation? location, string officeName)
    {
        if (state.IsAlreadyEnabled)
            return $"{user.DisplayName} is already enabled for enterprise voice "
                   + $"(number {state.Endpoint?.LineUri ?? "unknown"}). Use a change request instead of the wizard.";

        // Only block when Entra actually answered and said no. A licence that is merely still
        // provisioning is a warning, and a failed lookup must not invent a blocker.
        if (license.Checked && !license.HasTeamsPhone && !license.TeamsPhonePending)
            return $"{user.DisplayName} has no Teams Phone licence, so no number can be assigned. "
                   + $"Entra reports: {license.Summary}. Assign a licence that includes Teams Phone "
                   + "(E5, Teams Phone Standard or Teams Phone with Calling Plan), then run the wizard again.";

        if (officeName.Length == 0)
            return "The AD account has no office set. Pick the employee's site below to continue.";

        if (location is null)
            return $"The office '{officeName}' is not in the telephony database. Pick a known site below, "
                   + "or have the site added first.";

        if (location.TeamsVoiceEnabled == 2)
            return $"The site '{location.Name}' is not enabled for Teams enterprise voice, so this employee "
                   + "cannot be enabled. If there is a ticket, close it accordingly.";

        if (location.TeamsVoiceEnabled != 1)
            return $"The site '{location.Name}' has no Teams voice status recorded. Have the site configured first.";

        if (string.IsNullOrWhiteSpace(location.LocationCode))
            return $"The site '{location.Name}' has no location code, so no policies can be granted.";

        return null;
    }

    // ----- step 2: pick a number -----

    public Task<IReadOnlyList<UcDidRange>> GetDidRangesAsync(string locationCode, CancellationToken ct = default) =>
        _repository.GetDidRangesAsync(locationCode, ct);

    public Task<IReadOnlyList<FreeDid>> FindFreeDidsAsync(UcDidRange range, CancellationToken ct = default) =>
        _repository.FindFreeDidsAsync(range, ct);

    // ----- step 3: apply -----

    public async Task<NewHireResult> ConfigureAsync(NewHireRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();
        var newHire = _settings.Current.NewHire;
        var steps = new List<NewHireStep>();

        if (string.IsNullOrWhiteSpace(request.TaskNumber))
            return await FailAsync(correlationId, request, "Enter the ticket/task number authorizing this change.");

        if (string.IsNullOrWhiteSpace(request.User.Mail))
            return await FailAsync(correlationId, request,
                "The AD account has no mail address, so no SIP address can be derived. Fix the account first.");

        await LogAsync(correlationId, "STAR", "New hire telephony setup has started", request, Severity.Info);
        await LogAsync(correlationId, "INPUT",
            $"User={request.User.UserPrincipalName}; Office={request.Location.Name} ({request.Location.LocationCode}); "
            + $"Number={request.PhoneNumber}; WriteAd={newHire.WriteAdAttributes}",
            request, Severity.Info);

        // The run is recorded before anything changes, so a run that dies halfway still leaves a
        // trace against the employee rather than vanishing.
        int logId;
        try
        {
            logId = await _repository.StartRunAsync(
                request.User.SamAccountName,
                _auth.CurrentUser?.Upn ?? Environment.UserName,
                request.TaskNumber,
                ct);
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, request, $"Could not open a run in the telephony database: {ex.Message}");
        }

        // --- on-premises AD first ---
        //
        // Where Entra Connect still synchronises the msRTCSIP-* attributes, on-premises is the
        // authoritative source: a number set only in Teams is overwritten - or cleared - on the next
        // sync. Writing the same values here first makes that sync a no-op.
        var adResults = new List<AdAttributeResult>();
        if (newHire.WriteAdAttributes)
        {
            Report(onProgress, "Writing the SIP address and line URI to Active Directory...");
            try
            {
                adResults.AddRange(await _directory.WriteTelephonyAttributesAsync(
                    request.User, request.LineUri, request.SipAddress, newHire.DeploymentLocator, ct));

                foreach (var result in adResults)
                    steps.Add(new NewHireStep($"AD {result.Attribute}", result.Value, result.Error));
            }
            catch (Exception ex)
            {
                steps.Add(NewHireStep.Failed("Active Directory", request.LineUri, ex.Message));
                await LogAsync(correlationId, "FAIL", $"AD write failed: {ex.Message}", request, Severity.Error);
            }
        }
        else
        {
            steps.Add(NewHireStep.Ok("Active Directory", "skipped - AD attribute sync is turned off in Settings"));
        }

        // --- Teams ---
        Report(onProgress, $"Assigning {request.PhoneNumber} in Teams...");
        steps.Add(await TryStepAsync("Phone number", request.PhoneNumber,
            () => _teams.AssignPhoneNumberAsync(request.User.UserPrincipalName, request.PhoneNumber, newHire.PhoneNumberType, ct)));

        // All four site policies carry the location code as their name - that is the tenant's naming
        // convention and how the legacy script granted them.
        var policies = new (string Name, string Cmdlet, string PolicyName)[]
        {
            ("Voice routing policy", "Grant-CsOnlineVoiceRoutingPolicy", request.PolicyName),
            ("Dial plan", "Grant-CsTenantDialPlan", request.PolicyName),
            ("Emergency calling policy", "Grant-CsTeamsEmergencyCallingPolicy", request.PolicyName),
            ("Emergency call routing policy", "Grant-CsTeamsEmergencyCallRoutingPolicy", request.PolicyName),
            ("Teams upgrade policy", "Grant-CsTeamsUpgradePolicy", newHire.TeamsUpgradePolicyName),
        };

        foreach (var (name, cmdlet, policyName) in policies)
        {
            ct.ThrowIfCancellationRequested();
            Report(onProgress, $"Granting {name} '{policyName}'...");
            steps.Add(await TryStepAsync(name, policyName,
                () => _teams.GrantPolicyAsync(cmdlet, request.User.UserPrincipalName, policyName, ct)));
        }

        // --- record the endpoint ---
        //
        // The nightly import would eventually pick this up, but not before the next operator has
        // already handed the same number to somebody else.
        Report(onProgress, "Recording the endpoint in the telephony database...");
        steps.Add(await TryStepAsync("Telephony inventory", request.PhoneNumber, () =>
            _repository.AddEndpointAsync(new UcEndpoint
            {
                DisplayName = request.User.DisplayName,
                SamAccountName = request.User.SamAccountName,
                SipAddress = request.SipAddress,
                LineUri = request.LineUri,
                Did = request.Did,
                Type = "CsUser",
                VoicePolicy = request.PolicyName,
                DialPlan = request.PolicyName,
                EnterpriseVoiceEnabled = "true",
                HostedVoiceMail = "true",
                EmergencyCallingPolicy = request.PolicyName,
                EmergencyCallRoutingPolicy = request.PolicyName,
                Office = request.Location.Name,
            }, ct)));

        // --- read back what Teams actually holds ---
        Report(onProgress, "Reading the configuration back from Teams...");
        TeamsVoiceState? voiceState = null;
        try
        {
            voiceState = await _teams.GetVoiceStateAsync(request.User.UserPrincipalName, ct);
        }
        catch (Exception ex)
        {
            // A failed read-back says nothing about whether the changes landed, so it is a note
            // rather than a failure of the run.
            await LogAsync(correlationId, "WARN", $"Could not read the Teams configuration back: {ex.Message}", request, Severity.Warning);
        }

        // --- notify ---
        //
        // Where Entra Connect owns msRTCSIP-Line, Set-CsPhoneNumberAssignment can refuse the
        // assignment because the number is "managed on-premises". That is not a failed run: the AD
        // write above already put the number where the sync picks it up, so the employee does get
        // their number - it just arrives on the next sync cycle rather than immediately. So the
        // number counts as delivered if either side accepted it.
        var employeeNotified = false;
        var numberStep = steps.First(s => s.Name == "Phone number");
        var adLineWritten = steps.Any(s => s.Name == "AD msRTCSIP-Line" && s.Succeeded);
        var numberDelivered = numberStep.Succeeded || adLineWritten;

        if (newHire.NotifyEmployee && numberDelivered)
        {
            Report(onProgress, $"Notifying {request.User.Mail}...");
            var (subject, html) = NewHireNotificationMail.Enabled(request.User.GivenName, request.PhoneNumber);
            var notified = await TryStepAsync("Employee notification", request.User.Mail,
                () => _mail.SendAsync(new[] { request.User.Mail }, subject, html, ct: ct));
            steps.Add(notified);
            employeeNotified = notified.Succeeded;
        }

        var failedSteps = steps.Where(s => !s.Succeeded).ToList();

        if (failedSteps.Count > 0 && !string.IsNullOrWhiteSpace(newHire.AdminNotificationAddress))
        {
            var (subject, html) = NewHireNotificationMail.Failure(request, steps, logId);
            try
            {
                await _mail.SendAsync(SplitAddresses(newHire.AdminNotificationAddress), subject, html, ct: ct);
            }
            catch (Exception ex)
            {
                await LogAsync(correlationId, "WARN", $"Could not notify the telephony team: {ex.Message}", request, Severity.Warning);
            }
        }

        // --- close the run ---
        try
        {
            await _repository.CompleteRunAsync(logId, row => ApplyStepsToLog(row, request, steps, newHire), ct);
        }
        catch (Exception ex)
        {
            await LogAsync(correlationId, "WARN", $"Could not close the run in the telephony database: {ex.Message}", request, Severity.Warning);
        }

        foreach (var step in failedSteps)
            await LogAsync(correlationId, "FAIL", $"{step.Name} ({step.Detail}): {step.Error}", request, Severity.Error);

        await LogAsync(correlationId,
            failedSteps.Count == 0 ? "DONE" : "WARN",
            failedSteps.Count == 0
                ? $"New hire telephony setup has completed: {request.User.UserPrincipalName} on {request.PhoneNumber}"
                : $"New hire telephony setup finished with {failedSteps.Count} failed step(s)",
            request,
            failedSteps.Count == 0 ? Severity.Success : Severity.Warning);

        return new NewHireResult
        {
            // The run counts as succeeded when the employee ends up with their number; a failed
            // notification or inventory row is worth showing but does not undo that.
            Succeeded = numberDelivered,
            ErrorMessage = numberDelivered ? null : numberStep.Error,
            NumberPendingSync = !numberStep.Succeeded && adLineWritten,
            Steps = steps,
            VoiceState = voiceState,
            NewHireLogId = logId,
            CorrelationId = correlationId,
            EmployeeNotified = employeeNotified,
        };
    }

    /// <summary>
    /// Runs one step and turns any exception into a recorded failure. Steps are deliberately not
    /// allowed to abort the run: a half-configured account needs the rest attempted, and the operator
    /// needs to see exactly which one broke.
    /// </summary>
    private static async Task<NewHireStep> TryStepAsync(string name, string detail, Func<Task> action)
    {
        try
        {
            await action();
            return NewHireStep.Ok(name, detail);
        }
        catch (Exception ex)
        {
            return NewHireStep.Failed(name, detail, ex.Message);
        }
    }

    /// <summary>
    /// Maps the steps onto the legacy newhirelog columns (1 = succeeded, 0 = failed, with a matching
    /// *Error column). The shape is odd but it is what the existing reports over that table read.
    /// </summary>
    private static void ApplyStepsToLog(
        UcNewHireLog row,
        NewHireRequest request,
        IReadOnlyList<NewHireStep> steps,
        NewHireSettings settings)
    {
        row.AllVariablesSet = 1;
        row.SnTicket = request.TaskNumber;

        // The legacy script's credential-file and MSOnline checks have no equivalent here - the app
        // signs in once at startup - so they are recorded as satisfied rather than left at 0, which
        // in the old reports means "the script never ran".
        row.CrdFileExisting = 1;
        row.CrdFileUpToDate = 1;
        row.TeamsConnected = 1;
        row.TeamsConnectedError = "";
        row.MsolConnected = 1;
        row.MsolConnectedError = "";
        row.SqlServerModule = 1;

        void Ad(string attribute, Action<int, string> assign)
        {
            var step = steps.FirstOrDefault(s => s.Name == $"AD {attribute}");
            if (step is null)
            {
                // Either the AD step is turned off, or it threw before reaching this attribute.
                assign(settings.WriteAdAttributes ? 0 : 1,
                    settings.WriteAdAttributes ? "not attempted" : "skipped - AD attribute sync is turned off");
                return;
            }

            assign(step.Succeeded ? 1 : 0, step.Error ?? "");
        }

        Ad("msRTCSIP-Line", (ok, err) => { row.MsRtcSipLine = ok; row.MsRtcSipLineError = err; });
        Ad("msRTCSIP-PrimaryUserAddress", (ok, err) => { row.MsRtcSipPrimaryUserAddress = ok; row.MsRtcSipPrimaryUserAddressError = err; });
        Ad("msRTCSIP-DeploymentLocator", (ok, err) => { row.MsRtcSipDeploymentLocator = ok; row.MsRtcSipDeploymentLocatorError = err; });
        Ad("msRTCSIP-FederationEnabled", (ok, err) => { row.MsRtcSipFederationEnabled = ok; row.MsRtcSipFederationEnabledError = err; });
        Ad("msRTCSIP-InternetAccessEnabled", (ok, err) => { row.MsRtcSipInternetAccessEnabled = ok; row.MsRtcSipInternetAccessEnabledError = err; });
        Ad("msRTCSIP-UserEnabled", (ok, err) => { row.MsRtcSipUserEnabled = ok; row.MsRtcSipUserEnabledError = err; });

        void Teams(string name, Action<int, string> assign)
        {
            var step = steps.FirstOrDefault(s => s.Name == name);
            assign(step?.Succeeded == true ? 1 : 0, step?.Error ?? "not attempted");
        }

        Teams("Voice routing policy", (ok, err) => { row.VoiceRoutingPolicy = ok; row.VoiceRoutingPolicyError = err; });
        Teams("Dial plan", (ok, err) => { row.DialPlan = ok; row.DialPlanError = err; });
        Teams("Emergency calling policy", (ok, err) => { row.EmergencyCallingPolicy = ok; row.EmergencyCallingPolicyError = err; });
        Teams("Emergency call routing policy", (ok, err) => { row.EmergencyCallRoutingPolicy = ok; row.EmergencyCallRoutingPolicyError = err; });
        Teams("Teams upgrade policy", (ok, err) => { row.TeamsUpgradePolicy = ok; row.TeamsUpgradePolicyError = err; });

        // "EVEnabled" in the legacy schema means "the user can make calls", which is now the phone
        // number assignment - Set-CsPhoneNumberAssignment enables enterprise voice as part of it.
        Teams("Phone number", (ok, err) => { row.EvEnabled = ok; row.EvEnabledError = err; });
    }

    private static string[] SplitAddresses(string value) =>
        value.Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    private static void Report(Action<string>? onProgress, string message) => onProgress?.Invoke(message);

    private async Task<NewHireResult> FailAsync(Guid correlationId, NewHireRequest request, string error)
    {
        await LogAsync(correlationId, "FAIL", error, request, Severity.Error);
        return NewHireResult.Failed(correlationId, error);
    }

    private async Task LogAsync(Guid correlationId, string eventCode, string message, NewHireRequest request, Severity severity)
    {
        try
        {
            await _log.WriteAsync(new LogEntry
            {
                CorrelationId = correlationId,
                Area = "NewHire",
                Action = "EnableTelephony",
                EventCode = eventCode,
                Severity = severity,
                TargetObject = request.User.UserPrincipalName,
                TaskNumber = request.TaskNumber,
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
