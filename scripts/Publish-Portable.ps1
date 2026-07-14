<#
.SYNOPSIS
    Builds a self-contained, portable win-x64 release of M365Manager and zips it.
.DESCRIPTION
    Runs `dotnet publish` (Release, win-x64, --self-contained) into publish/M365Manager,
    then packages that folder into a zip. The target machine needs no .NET install - just
    unzip and run M365Manager.exe.
.PARAMETER OutputZip
    Path of the zip file to create. Defaults to <Desktop>\M365Manager-portable.zip.
.EXAMPLE
    pwsh scripts/Publish-Portable.ps1
.EXAMPLE
    pwsh scripts/Publish-Portable.ps1 -OutputZip D:\Transfer\M365Manager.zip
#>
param(
    [string]$OutputZip = (Join-Path ([Environment]::GetFolderPath("Desktop")) "M365Manager-portable.zip")
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$publishDir = Join-Path $repoRoot "publish\M365Manager"
$csproj = Join-Path $repoRoot "src\M365Manager\M365Manager.csproj"

Write-Host "Publishing self-contained win-x64 release to $publishDir..."
dotnet publish $csproj -c Release -r win-x64 --self-contained true -o $publishDir
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed with exit code $LASTEXITCODE" }

if (Test-Path $OutputZip) { Remove-Item $OutputZip -Force }

Write-Host "Zipping to $OutputZip..."
Compress-Archive -Path (Join-Path $publishDir '*') -DestinationPath $OutputZip -CompressionLevel Optimal

$sizeMb = [Math]::Round((Get-Item $OutputZip).Length / 1MB, 1)
Write-Host "Done: $OutputZip ($sizeMb MB)"
