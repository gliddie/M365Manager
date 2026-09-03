# Portable build for testing on another machine

`scripts/Publish-Portable.ps1` produces a self-contained win-x64 build and zips it up so it
can be copied to a machine without .NET installed - useful for quick cross-machine testing
before a real release/signing process exists.

## Usage

```powershell
pwsh scripts/Publish-Portable.ps1
```

Produces `<Desktop>\M365Manager-portable.zip`. Pass `-OutputZip <path>` to write it somewhere
else instead (e.g. a shared drive).

## What it does

- `dotnet publish` the WPF app (`src/M365Manager/M365Manager.csproj`) as `Release`,
  `-r win-x64 --self-contained true`, into `publish/M365Manager` (gitignored).
- Zips that folder's contents.

The result bundles the .NET/WPF runtime and the PowerShell modules the app hosts in-process
(`src/M365Manager/Modules`: ExchangeOnlineManagement, PnP.PowerShell, MicrosoftTeams), so the
target machine only needs to unzip and run `M365Manager.exe` - no .NET runtime install required
there.

`Modules/` is gitignored, so it is *not* restored by a fresh clone. On a build machine that has
never run it, `pwsh scripts/Restore-Modules.ps1` has to run once before publishing - otherwise the
zip ships without the modules and every feature that needs one fails at first use with a
`FileNotFoundException`. Worth checking the publish output after adding a module:

```powershell
Get-ChildItem publish\M365Manager\Modules -Depth 2 -Filter *.psd1 |
    Where-Object { $_.BaseName -eq $_.Directory.Parent.Name }
```

That should list one manifest per bundled module - the layout
`Modules\<Name>\<Version>\<Name>.psd1` is exactly what `PowerShellHost.FindBundledModuleManifest`
looks for.

## On the target machine

- Windows SmartScreen will likely flag the unsigned exe on first run - "More info" ->
  "Run anyway".
- Settings (SQL server, Tenant/Client ID, domain) are stored per-machine in
  `%AppData%\M365Manager\settings.json` and start empty - they need to be entered once in
  Settings before the app-startup auto-connect (see `IM365Connector`) has anything to connect to.

### Extra requirements for the New Hire page

This is the only feature that reads a second database and the only one that writes to
on-premises Active Directory, so it needs two things the other pages don't:

- **Settings > New Hire > Telephony database.** The number inventory (sites, DID ranges, blocked
  numbers, endpoints) lives in its own catalog - `uchelper` by default - on the *same* server as
  the one in the SQL Server section, whose server and credentials it reuses. Until the catalog name
  is filled in, the page has nothing to read and every lookup fails. "Test connection" there
  verifies the name before a real run does.
- **A domain-joined machine, and an admin with rights on the msRTCSIP-\* attributes.** The AD step
  binds over Kerberos as the signed-in Windows user - deliberately, so no service account
  credentials are stored anywhere. That differs from the legacy NewHireWizard, which carried a
  service account and therefore always had the same rights regardless of who ran it. If the
  operator lacks write access, the AD step fails visibly in the result list rather than silently,
  but the number then never reaches on-premises AD - and on a tenant where Entra Connect
  synchronises those attributes, on-premises is the authoritative source (see
  `INewHireService`). The fix is to delegate write access on the msRTCSIP-\* attributes to the
  admin group, not to turn the AD step off.

  On a machine that is not domain-joined, set a domain controller under Settings > New Hire -
  otherwise the search base cannot be discovered from RootDSE.
