<#
.SYNOPSIS
    Restores the PowerShell modules M365Manager bundles and hosts in-process
    (ExchangeOnlineManagement, PnP.PowerShell, MicrosoftTeams) into src/M365Manager/Modules.
.DESCRIPTION
    These modules are not committed to git (see .gitignore) because they're large binary
    module trees. Run this once per dev/build machine before building or publishing.
    The app's .csproj already copies everything under Modules\ to the output directory.
.EXAMPLE
    pwsh scripts/Restore-Modules.ps1
#>
$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$modulesDir = Join-Path $repoRoot "src\M365Manager\Modules"
New-Item -ItemType Directory -Path $modulesDir -Force | Out-Null

$modules = @("ExchangeOnlineManagement", "PnP.PowerShell", "MicrosoftTeams")
foreach ($name in $modules) {
    Write-Host "Restoring $name to $modulesDir ..."
    Save-Module -Name $name -Path $modulesDir -Force
}

Write-Host "Done. Restored: $($modules -join ', ')"
