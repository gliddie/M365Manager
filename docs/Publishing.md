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
- **A machine that can reach a domain controller, and an account allowed to write the
  msRTCSIP-\* attributes.**

  By default the AD step binds as the signed-in Windows user. That only works where the desktop
  session itself holds those rights. Where admin work is done from a separate admin account - the
  usual split, and the case here - the desktop session is the unprivileged account, and every write
  comes back `Access is denied` even though reads succeed. Microsoft 365 hides this difference
  because Graph, Exchange, SharePoint and Teams all sign in interactively as the admin; LDAP has no
  such sign-in to reuse and takes the process token.

  The fix is **Settings > New Hire > Active Directory account**: enter the admin account and its
  password (`GLOBAL\admin.jdoe` or `admin.jdoe@corp.contoso.com`; DPAPI-encrypted at rest, like the
  SQL and SMTP passwords). Those credentials are used for both the search and the write. Leaving
  the fields empty keeps the signed-in-user behaviour.

  Delegating the msRTCSIP-\* attributes to the operator's everyday account is the other way to
  solve it and needs no credentials stored - but it grants directory write access to an account
  that is not otherwise privileged, so prefer naming the admin account here.

  A failed AD write shows up in the wizard's result list rather than silently, and the message names
  the identity that was refused. Do not "fix" it by turning the AD step off: on a tenant where Entra
  Connect synchronises those attributes, on-premises is the authoritative source for the number and
  the next sync would undo the Teams assignment (see `INewHireService`).

  On a machine that is not domain-joined, set a domain controller under Settings > New Hire -
  otherwise the search base cannot be discovered from RootDSE.
