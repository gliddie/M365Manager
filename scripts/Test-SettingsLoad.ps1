<#
.SYNOPSIS
    Checks that SettingsService.Load survives a damaged settings file.
.DESCRIPTION
    The app has no test project, so this stands in for one on the path where a regression is
    expensive: Load used to wrap the whole deserialization in one try/catch and fall back to a
    blank AppSettings. A single unreadable value discarded every section, and because the blanks
    then sat in Current, the next Save wrote them over a file that was still almost intact.

    These checks pin that down: one bad section must not cost the others, the failure must be
    reported rather than swallowed, and the original file must be copied aside so a later Save
    cannot destroy it.

    Run this after touching SettingsService, AppSettings, or anything about how settings are
    serialized - adding a section, changing a type, introducing an enum.

    SettingsService resolves its own path through Environment.GetFolderPath, which asks Windows
    directly and ignores $env:APPDATA, so the checks cannot be redirected into a sandbox. They run
    against the real %APPDATA%\M365Manager\settings.json, which is backed up first and restored
    afterwards - including on failure.
.PARAMETER Configuration
    Build configuration to load the assembly from. Defaults to Debug.
.EXAMPLE
    pwsh scripts/Test-SettingsLoad.ps1
.EXAMPLE
    pwsh scripts/Test-SettingsLoad.ps1 -Configuration Release
#>
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot

$dll = Get-ChildItem -Path (Join-Path $repoRoot "src\M365Manager.Core\bin\$Configuration") `
                     -Filter "M365Manager.Core.dll" -Recurse -ErrorAction SilentlyContinue |
       Select-Object -First 1
if (-not $dll) {
    throw "M365Manager.Core.dll not found under src\M365Manager.Core\bin\$Configuration - run 'dotnet build' first."
}

Add-Type -Path $dll.FullName | Out-Null

$dir = Join-Path ([Environment]::GetFolderPath("ApplicationData")) "M365Manager"
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$file = Join-Path $dir "settings.json"
$backup = "$file.pretest-backup"
$hadSettings = Test-Path $file
if ($hadSettings) { Copy-Item $file $backup -Force }

$script:passed = 0
$script:failed = 0

function Check {
    param([string]$Name, [string]$Json, [scriptblock]$Assert)

    Get-ChildItem $dir -Filter "*.unreadable-*.bak" | Remove-Item -Force -ErrorAction SilentlyContinue
    Set-Content -Path $file -Value $Json -Encoding UTF8

    $service = [M365Manager.Core.Settings.SettingsService]::new()
    $result = & $Assert $service

    if ($result -eq $true) {
        Write-Host "PASS  $Name" -ForegroundColor Green
        $script:passed++
    }
    else {
        Write-Host "FAIL  $Name -> $result" -ForegroundColor Red
        $script:failed++
    }
}

try {
    Write-Host "Checking SettingsService.Load against $file"
    Write-Host ""

    Check "clean file loads" '{ "Sql": { "Server": "sql01", "Database": "m365" }, "Ui": { "Theme": "Dark" } }' {
        param($s)
        if ($s.Current.Sql.Server -ne "sql01") { return "Sql.Server = $($s.Current.Sql.Server)" }
        if ("$($s.Current.Ui.Theme)" -ne "Dark") { return "Theme = $($s.Current.Ui.Theme)" }
        if ($null -ne $s.LastLoadError) { return "unexpected error: $($s.LastLoadError)" }
        $true
    }

    # The regression this script exists for.
    Check "one bad section keeps the rest" '{ "Sql": { "Server": "sql01" }, "Smtp": { "Port": "not-a-number" }, "M365": { "TenantId": "t-1" } }' {
        param($s)
        if ($s.Current.Sql.Server -ne "sql01") { return "Sql lost: $($s.Current.Sql.Server)" }
        if ($s.Current.M365.TenantId -ne "t-1") { return "M365 lost: $($s.Current.M365.TenantId)" }
        if ($null -eq $s.LastLoadError) { return "failure was not reported" }
        if ($s.LastLoadError -notmatch "Smtp") { return "error does not name Smtp: $($s.LastLoadError)" }
        $true
    }

    Check "bad section is quarantined" '{ "Sql": { "Server": "keepme" }, "Smtp": { "Port": "nope" } }' {
        param($s)
        if ($null -eq $s.QuarantinedFilePath) { return "nothing quarantined" }
        if (-not (Test-Path $s.QuarantinedFilePath)) { return "quarantine file missing" }
        if ((Get-Content $s.QuarantinedFilePath -Raw) -notmatch "keepme") { return "quarantine lost the original content" }
        $true
    }

    Check "malformed json is reported" '{ "Sql": { "Server": "x" ' {
        param($s)
        if ($null -eq $s.LastLoadError) { return "failure was not reported" }
        if ($null -eq $s.QuarantinedFilePath) { return "nothing quarantined" }
        $true
    }

    # Settings files get hand-edited during troubleshooting; a stray comma must not cost a config.
    Check "tolerates comments and trailing commas" @'
{
  // set during troubleshooting
  "Sql": { "Server": "sql01", },
}
'@ {
        param($s)
        if ($s.Current.Sql.Server -ne "sql01") { return "Sql.Server = $($s.Current.Sql.Server)" }
        if ($null -ne $s.LastLoadError) { return "unexpected error: $($s.LastLoadError)" }
        $true
    }

    # A section missing entirely is normal - the file predates it - and is not a failure.
    Check "missing section is not an error" '{ "Sql": { "Server": "sql01" } }' {
        param($s)
        if ($null -ne $s.LastLoadError) { return "reported an error for a missing section: $($s.LastLoadError)" }
        if ("$($s.Current.Ui.Theme)" -ne "System") { return "Ui default wrong: $($s.Current.Ui.Theme)" }
        $true
    }
}
finally {
    Get-ChildItem $dir -Filter "*.unreadable-*.bak" | Remove-Item -Force -ErrorAction SilentlyContinue
    if ($hadSettings) {
        Move-Item $backup $file -Force
    }
    else {
        Remove-Item $file -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Write-Host "$script:passed passed, $script:failed failed"
if ($script:failed -gt 0) { exit 1 }
