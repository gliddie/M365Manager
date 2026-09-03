using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using PowerShellSdk = System.Management.Automation.PowerShell;

namespace M365Manager.Core.PowerShell;

/// <summary>
/// Hosts a single long-lived PowerShell runspace in-process (Microsoft.PowerShell.SDK).
/// Used to run the bundled ExchangeOnlineManagement module. The runspace keeps the
/// Exchange Online session alive between calls. Access is serialized (runspaces are
/// not thread-safe).
/// </summary>
public sealed class PowerShellHost : IDisposable
{
    /// <summary>How many pipeline objects get formatted into the transcript per command.</summary>
    private const int MaxTranscriptObjects = 20;

    private readonly Runspace _runspace;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly IPowerShellTranscript _transcript;
    private bool _exchangeModuleImported;
    private bool _pnpModuleImported;
    private bool _teamsModuleImported;

    public PowerShellHost(IPowerShellTranscript transcript)
    {
        _transcript = transcript;

        var iss = InitialSessionState.CreateDefault();
        iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;
        _runspace = RunspaceFactory.CreateRunspace(iss);
        _runspace.Open();
    }

    /// <summary>
    /// Locates a bundled module manifest next to the app. Bundling rather than relying on whatever
    /// is installed on the machine keeps every box on the same module version - a mismatched
    /// Graph/Teams module is exactly the kind of thing that works on one PC and not the next.
    /// </summary>
    private static string FindBundledModuleManifest(string moduleName)
    {
        var root = Path.Combine(AppContext.BaseDirectory, "Modules", moduleName);
        if (!Directory.Exists(root))
            throw new FileNotFoundException(
                $"{moduleName} module not found at '{root}'. Restore it once per machine with "
                + "'pwsh scripts/Restore-Modules.ps1'.");

        var versionDir = Directory.GetDirectories(root)
            .OrderByDescending(d => d)
            .FirstOrDefault()
            ?? throw new FileNotFoundException($"No version folder found under '{root}'.");

        var psd1 = Path.Combine(versionDir, $"{moduleName}.psd1");
        if (!File.Exists(psd1))
            throw new FileNotFoundException($"Module manifest not found: '{psd1}'.");

        return psd1;
    }

    private async Task ImportBundledModuleAsync(string moduleName, CancellationToken ct)
    {
        var psd1 = FindBundledModuleManifest(moduleName);
        await InvokeAsync(ps => ps
            .AddCommand("Import-Module")
            .AddParameter("Name", psd1)
            .AddParameter("ErrorAction", "Stop"), ct: ct, suppressTranscript: true);
    }

    /// <summary>Imports the Exchange module once per runspace.</summary>
    public async Task EnsureExchangeModuleAsync(CancellationToken ct = default)
    {
        if (_exchangeModuleImported)
            return;

        await ImportBundledModuleAsync("ExchangeOnlineManagement", ct);
        _exchangeModuleImported = true;
    }

    /// <summary>Imports the PnP.PowerShell module once per runspace.</summary>
    public async Task EnsurePnPModuleAsync(CancellationToken ct = default)
    {
        if (_pnpModuleImported)
            return;

        await ImportBundledModuleAsync("PnP.PowerShell", ct);
        _pnpModuleImported = true;
    }

    /// <summary>Imports the MicrosoftTeams module once per runspace (Grant-Cs*, Set-CsPhoneNumberAssignment).</summary>
    public async Task EnsureTeamsModuleAsync(CancellationToken ct = default)
    {
        if (_teamsModuleImported)
            return;

        await ImportBundledModuleAsync("MicrosoftTeams", ct);
        _teamsModuleImported = true;
    }

    /// <summary>
    /// Builds and runs a command against the shared runspace, returning the output objects.
    /// Throws <see cref="PowerShellException"/> if the command reports errors.
    ///
    /// The command and its output are mirrored into <see cref="IPowerShellTranscript"/> for the
    /// console view, unless <paramref name="suppressTranscript"/> is set - which sign-in and module
    /// imports do, since those carry device codes and tokens.
    /// </summary>
    public async Task<IReadOnlyList<PSObject>> InvokeAsync(
        Action<PowerShellSdk> build,
        Action<string>? onInformation = null,
        CancellationToken ct = default,
        bool suppressTranscript = false)
    {
        await _gate.WaitAsync(ct);
        try
        {
            return await Task.Run(() =>
            {
                // Capture console output (the MSAL device-code prompt is written via
                // Console.WriteLine and would otherwise be invisible in a GUI app).
                TextWriter? originalOut = null;
                TextWriter? originalError = null;
                ForwardingTextWriter? forwarder = null;

                if (onInformation is not null)
                {
                    forwarder = new ForwardingTextWriter(onInformation);
                    originalOut = Console.Out;
                    originalError = Console.Error;
                    Console.SetOut(forwarder);
                    Console.SetError(forwarder);
                }

                try
                {
                    using var ps = PowerShellSdk.Create();
                    ps.Runspace = _runspace;
                    build(ps);

                    // Written before Invoke, so a long-running step is visible while it runs
                    // rather than only once it finishes.
                    if (!suppressTranscript)
                        _transcript.Write(TranscriptLineKind.Command, CommandRenderer.Render(ps.Commands.Commands));

                    // Also surface PowerShell's own streams (Write-Host / warnings).
                    if (onInformation is not null)
                    {
                        ps.Streams.Information.DataAdded += (_, e) =>
                            onInformation(ps.Streams.Information[e.Index].ToString());
                        ps.Streams.Warning.DataAdded += (_, e) =>
                            onInformation(ps.Streams.Warning[e.Index].Message);
                    }

                    if (!suppressTranscript)
                    {
                        ps.Streams.Warning.DataAdded += (_, e) =>
                            _transcript.Write(TranscriptLineKind.Warning, ps.Streams.Warning[e.Index].Message);
                        ps.Streams.Information.DataAdded += (_, e) =>
                            _transcript.Write(TranscriptLineKind.Information, ps.Streams.Information[e.Index].ToString());
                    }

                    var results = ps.Invoke();

                    if (!suppressTranscript)
                    {
                        WriteResultsToTranscript(results);

                        foreach (var error in ps.Streams.Error)
                            _transcript.Write(TranscriptLineKind.Error, error.ToString());
                    }

                    if (ps.HadErrors && ps.Streams.Error.Count > 0)
                        throw new PowerShellException(ps.Streams.Error[0].ToString());

                    return (IReadOnlyList<PSObject>)results.ToList();
                }
                finally
                {
                    if (originalOut is not null)
                        Console.SetOut(originalOut);
                    if (originalError is not null)
                        Console.SetError(originalError);
                    forwarder?.Dispose();
                }
            }, ct);
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <summary>
    /// Formats pipeline output the way PowerShell itself would print it, by piping it back through
    /// Out-String on the same runspace. That picks up the module's own format definitions, so a
    /// mailbox object comes out as the familiar table rather than a dump of every property.
    ///
    /// Called while the gate is held and on the same thread as the invoke, so reusing the runspace
    /// here is safe.
    /// </summary>
    private void WriteResultsToTranscript(IReadOnlyCollection<PSObject> results)
    {
        if (results.Count == 0)
            return;

        var shown = results.Take(MaxTranscriptObjects).ToArray();

        try
        {
            using var formatter = PowerShellSdk.Create();
            formatter.Runspace = _runspace;
            formatter.AddCommand("Out-String")
                     .AddParameter("InputObject", shown)
                     .AddParameter("Width", 200);

            var formatted = formatter.Invoke<string>();
            var text = string.Concat(formatted).TrimEnd();

            if (text.Length > 0)
                _transcript.Write(TranscriptLineKind.Output, text);
        }
        catch
        {
            // Formatting must never break the actual operation - fall back to a plain count.
            _transcript.Write(TranscriptLineKind.Output, $"<{results.Count} object(s) returned>");
        }

        if (results.Count > shown.Length)
            _transcript.Write(TranscriptLineKind.Output, $"... {results.Count - shown.Length} more object(s) not shown");
    }

    public void Dispose()
    {
        _runspace.Dispose();
        _gate.Dispose();
    }
}
