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
    private readonly Runspace _runspace;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private bool _exchangeModuleImported;

    public PowerShellHost()
    {
        var iss = InitialSessionState.CreateDefault();
        iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;
        _runspace = RunspaceFactory.CreateRunspace(iss);
        _runspace.Open();
    }

    /// <summary>Locates the bundled ExchangeOnlineManagement manifest next to the app.</summary>
    private static string FindExchangeModuleManifest()
    {
        var root = Path.Combine(AppContext.BaseDirectory, "Modules", "ExchangeOnlineManagement");
        if (!Directory.Exists(root))
            throw new FileNotFoundException(
                $"ExchangeOnlineManagement module not found at '{root}'. It must ship in the app's Modules folder.");

        var versionDir = Directory.GetDirectories(root)
            .OrderByDescending(d => d)
            .FirstOrDefault()
            ?? throw new FileNotFoundException($"No version folder found under '{root}'.");

        var psd1 = Path.Combine(versionDir, "ExchangeOnlineManagement.psd1");
        if (!File.Exists(psd1))
            throw new FileNotFoundException($"Module manifest not found: '{psd1}'.");

        return psd1;
    }

    /// <summary>Imports the Exchange module once per runspace.</summary>
    public async Task EnsureExchangeModuleAsync(CancellationToken ct = default)
    {
        if (_exchangeModuleImported)
            return;

        var psd1 = FindExchangeModuleManifest();
        await InvokeAsync(ps => ps
            .AddCommand("Import-Module")
            .AddParameter("Name", psd1)
            .AddParameter("ErrorAction", "Stop"), ct: ct);

        _exchangeModuleImported = true;
    }

    /// <summary>
    /// Builds and runs a command against the shared runspace, returning the output objects.
    /// Throws <see cref="PowerShellException"/> if the command reports errors.
    /// </summary>
    public async Task<IReadOnlyList<PSObject>> InvokeAsync(
        Action<PowerShellSdk> build,
        Action<string>? onInformation = null,
        CancellationToken ct = default)
    {
        await _gate.WaitAsync(ct);
        try
        {
            return await Task.Run(() =>
            {
                using var ps = PowerShellSdk.Create();
                ps.Runspace = _runspace;
                build(ps);

                // Surface host messages (e.g. the device-code sign-in prompt) live.
                if (onInformation is not null)
                {
                    ps.Streams.Information.DataAdded += (_, e) =>
                        onInformation(ps.Streams.Information[e.Index].ToString());
                    ps.Streams.Warning.DataAdded += (_, e) =>
                        onInformation(ps.Streams.Warning[e.Index].Message);
                }

                var results = ps.Invoke();

                if (ps.HadErrors && ps.Streams.Error.Count > 0)
                    throw new PowerShellException(ps.Streams.Error[0].ToString());

                return (IReadOnlyList<PSObject>)results.ToList();
            }, ct);
        }
        finally
        {
            _gate.Release();
        }
    }

    public void Dispose()
    {
        _runspace.Dispose();
        _gate.Dispose();
    }
}
