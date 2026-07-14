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

The result bundles the .NET/WPF runtime and the ExchangeOnlineManagement PowerShell module
(`src/M365Manager/Modules`), so the target machine only needs to unzip and run
`M365Manager.exe` - no .NET runtime install required there.

## On the target machine

- Windows SmartScreen will likely flag the unsigned exe on first run - "More info" ->
  "Run anyway".
- Settings (SQL server, Tenant/Client ID, domain) are stored per-machine in
  `%AppData%\M365Manager\settings.json` and start empty - they need to be entered once in
  Settings before the app-startup auto-connect (see `IM365Connector`) has anything to connect to.
