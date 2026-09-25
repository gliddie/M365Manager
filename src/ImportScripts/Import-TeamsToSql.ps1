<#
.SYNOPSIS
Imports Microsoft Teams-specific details into SQL Server for the M365Manager Teams overview grid.

.DESCRIPTION
Companion script to Import-M365GroupsToSql.ps1. That script already imports every M365 group
(including the group behind each Team) into dbo.Groups/dbo.GroupMembers. This script only adds
the Team-specific attributes into dbo.Teams (see src/M365Manager.Data/Scripts/04_Teams.sql),
keyed by the same group id:
  - From Microsoft Graph (app-only, certificate auth): Visibility, CreatedDateTime, IsArchived,
    AllowToAddGuests - the effective value: tenant-wide Group.Unified setting AND the team's own
    Group.Unified.Guest setting, read via Invoke-MgGraphRequest (no extra Graph submodule).
  - From Exchange Online (app-only, certificate auth - separate app registration/certificate
    from the Graph one): HiddenFromAddressListsEnabled, WelcomeMessageEnabled, SharePointSiteUrl.
    These three have no Graph equivalent - they're Exchange-only UnifiedGroup properties.

It is designed to run on a separate server as a scheduled task, a few minutes AFTER
Import-M365GroupsToSql.ps1 (dbo.Teams has a foreign key on dbo.Groups - a team whose group row
doesn't exist yet is skipped and logged, not inserted).

REQUIRED PERMISSIONS
- Microsoft Graph (app-only): Group.Read.All, Directory.Read.All, Team.ReadBasic.All
- Exchange Online (app-only): Exchange Administrator or Compliance Administrator role assigned
  to the app registration (Connect-ExchangeOnline -AppId/-CertificateThumbprint/-Organization).
  See docs/AppRegistration-Setup.md for the app-only certificate setup steps.

REQUIRED MODULES
- Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Teams
- ExchangeOnlineManagement

IMPORTANT
- Replace the values marked with TODO before running the script.
- If a group is not found in dbo.Teams, the GUI simply shows fewer columns for it (no live
  M365 fallback for the overview grid - that's this script's job).
#>

[CmdletBinding()]
param(
    # --- Microsoft Graph (app-only) ---
    [Parameter()]
    [string]$TenantId = "TODO_TENANT_ID",

    [Parameter()]
    [string]$ClientId = "TODO_GRAPH_CLIENT_ID",

    [Parameter()]
    [string]$CertificateThumbprint = "TODO_GRAPH_CERTIFICATE_THUMBPRINT",

    # --- Exchange Online (app-only) - separate app registration/certificate from Graph above ---
    [Parameter()]
    [string]$ExoAppId = "TODO_EXO_APP_ID",

    [Parameter()]
    [string]$ExoCertificateThumbprint = "TODO_EXO_CERTIFICATE_THUMBPRINT",

    [Parameter()]
    [string]$Organization = "TODO_TENANT.onmicrosoft.com",

    # --- SQL ---
    [Parameter()]
    [string]$SqlServer = "TODO_SQL_SERVER",

    [Parameter()]
    [string]$SqlDatabase = "M365Manager",

    [Parameter()]
    [string]$SqlUser = "",

    [Parameter()]
    [string]$SqlPassword = "",

    [Parameter()]
    [switch]$UseIntegratedSecurity = $false,

    [Parameter()]
    [switch]$UseSqlEncryption = $true,

    [Parameter()]
    [string]$LogFile = "",

    [Parameter()]
    [int]$MaxRetries = 3,

    [Parameter()]
    [int]$MaxLogFiles = 14,

    [Parameter()]
    [int]$LogRetentionDays = 30,

    [Parameter()]
    [switch]$VerboseDebug = $false
)

$ErrorActionPreference = 'Stop'

function Write-ImportLog {
    param([string]$Message)

    $line = "[{0}] {1}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Message
    Write-Host $line
    if (-not [string]::IsNullOrWhiteSpace($LogFile)) {
        $logDirectory = Split-Path -Path $LogFile -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDirectory)) {
            New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        }
        Add-Content -Path $LogFile -Value $line
    }
}

function Write-DebugLog {
    param([string]$Message)
    if ($VerboseDebug) { Write-ImportLog $Message }
}

function Rotate-ImportLogs {
    if ([string]::IsNullOrWhiteSpace($LogFile)) { return }
    $logDirectory = Split-Path -Path $LogFile -Parent
    if ([string]::IsNullOrWhiteSpace($logDirectory) -or -not (Test-Path -Path $logDirectory)) { return }

    $files = Get-ChildItem -Path $logDirectory -Filter "*.log" -File | Sort-Object LastWriteTimeUtc -Descending
    $filesToDelete = $files | Where-Object { $_.LastWriteTimeUtc -lt (Get-Date).AddDays(-$LogRetentionDays) }
    foreach ($file in $filesToDelete) {
        try { Remove-Item -Path $file.FullName -Force -ErrorAction Stop } catch { }
    }

    $files = Get-ChildItem -Path $logDirectory -Filter "*.log" -File | Sort-Object LastWriteTimeUtc -Descending
    if ($files.Count -gt $MaxLogFiles) {
        $files | Select-Object -Skip $MaxLogFiles | ForEach-Object {
            try { Remove-Item -Path $_.FullName -Force -ErrorAction Stop } catch { }
        }
    }
}

function Connect-GraphAppOnly {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft Graph PowerShell module is not installed. Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $existingContext = Get-MgContext -ErrorAction SilentlyContinue
    if ($existingContext) { return }

    Write-DebugLog 'Connecting to Microsoft Graph with certificate thumbprint...'
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome | Out-Null

    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        throw 'Microsoft Graph connection could not be established.'
    }
}

function Connect-ExchangeAppOnly {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module is not installed. Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    Write-DebugLog 'Connecting to Exchange Online with certificate thumbprint...'
    Connect-ExchangeOnline -AppId $ExoAppId -CertificateThumbprint $ExoCertificateThumbprint -Organization $Organization -ShowBanner:$false
}

function Get-ConnectionString {
    if ($UseIntegratedSecurity) {
        $encryptionPart = if ($UseSqlEncryption) { "Encrypt=True;TrustServerCertificate=True;" } else { "Encrypt=False;TrustServerCertificate=True;" }
        return "Server=$SqlServer;Database=$SqlDatabase;Integrated Security=True;$encryptionPart"
    }
    if ([string]::IsNullOrWhiteSpace($SqlUser) -or [string]::IsNullOrWhiteSpace($SqlPassword)) {
        throw "Provide SqlUser and SqlPassword or use -UseIntegratedSecurity."
    }
    $encryptionPart = if ($UseSqlEncryption) { "Encrypt=True;TrustServerCertificate=True;" } else { "Encrypt=False;TrustServerCertificate=True;" }
    return "Server=$SqlServer;Database=$SqlDatabase;User Id=$SqlUser;Password=$SqlPassword;$encryptionPart"
}

function Invoke-SqlCommand {
    param(
        [string]$Sql,
        [hashtable]$Parameters = @{}
    )

    $connection = New-Object System.Data.SqlClient.SqlConnection(Get-ConnectionString)
    $command = $connection.CreateCommand()
    $command.CommandText = $Sql
    foreach ($key in $Parameters.Keys) {
        $value = $Parameters[$key]
        $param = $command.Parameters.Add("@$key", [System.Data.SqlDbType]::NVarChar)

        if ($null -eq $value) { $param.Value = [System.DBNull]::Value; continue }
        if ($value -is [guid]) { $param.SqlDbType = [System.Data.SqlDbType]::UniqueIdentifier; $param.Value = [guid]$value }
        elseif ($value -is [bool]) { $param.SqlDbType = [System.Data.SqlDbType]::Bit; $param.Value = [bool]$value }
        elseif ($value -is [datetime]) { $param.SqlDbType = [System.Data.SqlDbType]::DateTime2; $param.Value = [datetime]$value }
        else { $param.Value = [string]$value }
    }

    $connection.Open()
    try { $command.ExecuteNonQuery() | Out-Null }
    finally { $connection.Dispose() }
}

function Ensure-Schema {
    $sql = @"
IF OBJECT_ID(N'dbo.Teams', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Teams (
        GroupId UNIQUEIDENTIFIER NOT NULL PRIMARY KEY,
        Visibility NVARCHAR(32) NULL,
        IsArchived BIT NOT NULL DEFAULT 0,
        AllowToAddGuests BIT NULL,
        HiddenFromAddressListsEnabled BIT NULL,
        WelcomeMessageEnabled BIT NULL,
        SharePointSiteUrl NVARCHAR(512) NULL,
        CreatedDateTime DATETIME2(3) NULL,
        LastImportedAtUtc DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        CONSTRAINT FK_Teams_Groups FOREIGN KEY (GroupId) REFERENCES dbo.Groups(Id)
    );
END
"@
    Invoke-SqlCommand -Sql $sql
}

function Test-GroupExistsInSql {
    param([guid]$GroupId)
    $sql = "SELECT 1 FROM dbo.Groups WHERE Id = @Id"
    $connection = New-Object System.Data.SqlClient.SqlConnection(Get-ConnectionString)
    $command = $connection.CreateCommand()
    $command.CommandText = $sql
    $command.Parameters.Add("@Id", [System.Data.SqlDbType]::UniqueIdentifier).Value = $GroupId
    $connection.Open()
    try { return $null -ne $command.ExecuteScalar() }
    finally { $connection.Dispose() }
}

function Write-TeamRecord {
    param([pscustomobject]$Team)

    $sql = @"
MERGE dbo.Teams AS target
USING (SELECT @GroupId AS GroupId) AS source ON target.GroupId = source.GroupId
WHEN MATCHED THEN
    UPDATE SET
        Visibility = @Visibility,
        IsArchived = @IsArchived,
        AllowToAddGuests = @AllowToAddGuests,
        HiddenFromAddressListsEnabled = @HiddenFromAddressListsEnabled,
        WelcomeMessageEnabled = @WelcomeMessageEnabled,
        SharePointSiteUrl = @SharePointSiteUrl,
        CreatedDateTime = @CreatedDateTime,
        LastImportedAtUtc = SYSUTCDATETIME()
WHEN NOT MATCHED THEN
    INSERT (GroupId, Visibility, IsArchived, AllowToAddGuests, HiddenFromAddressListsEnabled, WelcomeMessageEnabled, SharePointSiteUrl, CreatedDateTime, LastImportedAtUtc)
    VALUES (@GroupId, @Visibility, @IsArchived, @AllowToAddGuests, @HiddenFromAddressListsEnabled, @WelcomeMessageEnabled, @SharePointSiteUrl, @CreatedDateTime, SYSUTCDATETIME());
"@

    $parameters = @{
        GroupId = [guid]$Team.GroupId
        Visibility = $Team.Visibility
        IsArchived = [bool]$Team.IsArchived
        AllowToAddGuests = $Team.AllowToAddGuests
        HiddenFromAddressListsEnabled = $Team.HiddenFromAddressListsEnabled
        WelcomeMessageEnabled = $Team.WelcomeMessageEnabled
        SharePointSiteUrl = $Team.SharePointSiteUrl
        CreatedDateTime = $Team.CreatedDateTime
    }

    Invoke-SqlCommand -Sql $sql -Parameters $parameters
}

# ----- Guest access (AllowToAddGuests) -----
#
# Read through plain Graph REST (Invoke-MgGraphRequest, part of Microsoft.Graph.Authentication,
# which is loaded anyway) instead of Get-MgGroupSettingTemplate / Get-MgGroupSetting. The template
# cmdlet lives in Microsoft.Graph.Identity.DirectoryManagement, and on the scheduler box every
# version of that submodule that could be installed next to the pinned Groups/Teams modules
# collided with them ("Assembly with same name is already loaded") and broke Get-MgGroup itself.
# REST needs no extra module at all - the same approach the app's TeamsService takes.
#
# The value stored is the EFFECTIVE one: guests can be added to a team only if the tenant-wide
# Group.Unified setting allows it AND the team has no Group.Unified.Guest setting forbidding it.
# A team without its own setting inherits the tenant value - previously that case came out NULL,
# which is what nearly every "external" team looked like.

# Well-known id of the Group.Unified.Guest template; used when the template list can't be read.
$script:GuestTemplateFallbackId = '08d542b9-071f-4e16-94b0-74abb372e3d9'
$script:GuestSettingsUnavailable = $false

function Invoke-GraphGet {
    param([string]$Uri)
    Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
}

function Get-SettingValue {
    param($Setting, [string]$Name)
    $entry = $Setting.values | Where-Object { $_.name -eq $Name } | Select-Object -First 1
    if ($null -eq $entry -or [string]::IsNullOrWhiteSpace($entry.value)) { return $null }
    return [bool]::Parse($entry.value)
}

function Get-GuestAccessContext {
    $templateId = $script:GuestTemplateFallbackId
    try {
        $templates = Invoke-GraphGet -Uri 'https://graph.microsoft.com/v1.0/groupSettingTemplates?$select=id,displayName'
        $match = $templates.value | Where-Object { $_.displayName -eq 'Group.Unified.Guest' } | Select-Object -First 1
        if ($match) { $templateId = $match.id }
    }
    catch {
        Write-DebugLog "Could not list group setting templates ($($_.Exception.Message)) - using the well-known Group.Unified.Guest id."
    }

    # Tenant default: Group.Unified/AllowToAddGuests. No such setting means the Microsoft 365
    # default, which is "allowed".
    $tenantAllows = $true
    try {
        $tenantSettings = Invoke-GraphGet -Uri 'https://graph.microsoft.com/v1.0/groupSettings'
        $unified = $tenantSettings.value | Where-Object { $_.displayName -eq 'Group.Unified' } | Select-Object -First 1
        if ($unified) {
            $value = Get-SettingValue -Setting $unified -Name 'AllowToAddGuests'
            if ($null -ne $value) { $tenantAllows = $value }
        }
    }
    catch {
        Write-ImportLog "WARNING: Could not read the tenant-wide group settings ($($_.Exception.Message)) - AllowToAddGuests will be left NULL. The Graph app needs Directory.Read.All."
        $script:GuestSettingsUnavailable = $true
    }

    [pscustomobject]@{ TemplateId = $templateId; TenantAllows = $tenantAllows }
}

function Get-AllowToAddGuests {
    param([string]$GroupId, $Context)

    if ($script:GuestSettingsUnavailable) { return $null }
    if (-not $Context.TenantAllows) { return $false }   # the tenant switch overrides every team

    try {
        $settings = Invoke-GraphGet -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/settings"
    }
    catch {
        # A permission problem would repeat for every team - report it once and stop asking.
        if ($_.Exception.Message -match '403|Forbidden|Authorization_RequestDenied') {
            Write-ImportLog "WARNING: Not allowed to read group settings ($($_.Exception.Message)) - AllowToAddGuests will be left NULL for the remaining teams. The Graph app needs Directory.Read.All."
            $script:GuestSettingsUnavailable = $true
        }
        else {
            Write-DebugLog "Could not read settings of group $GroupId`: $($_.Exception.Message)"
        }
        return $null
    }

    $guestSetting = $settings.value | Where-Object { $_.templateId -eq $Context.TemplateId } | Select-Object -First 1
    if (-not $guestSetting) { return $true }   # no own setting - inherits the tenant's "allowed"

    $value = Get-SettingValue -Setting $guestSetting -Name 'AllowToAddGuests'
    if ($null -eq $value) { return $true }
    return $value
}

try {
    if (-not [string]::IsNullOrWhiteSpace($LogFile)) { Rotate-ImportLogs }

    Write-ImportLog 'Ensuring SQL schema...'
    Ensure-Schema

    Write-ImportLog 'Connecting to Microsoft Graph...'
    $attempt = 0
    $connected = $false
    while (-not $connected -and $attempt -lt $MaxRetries) {
        $attempt++
        try { Connect-GraphAppOnly; $connected = $true }
        catch {
            Write-ImportLog "Graph connection attempt $attempt failed: $($_.Exception.Message)"
            if ($attempt -ge $MaxRetries) { throw }
            Start-Sleep -Seconds 10
        }
    }

    Write-ImportLog 'Connecting to Exchange Online...'
    Connect-ExchangeAppOnly

    Write-ImportLog 'Querying Teams-enabled groups from Microsoft Graph...'
    $groups = Get-MgGroup -All -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -ConsistencyLevel eventual -CountVariable teamCount -ErrorAction Stop
    Write-ImportLog "Found $($groups.Count) Teams-enabled groups."

    $guestContext = Get-GuestAccessContext
    Write-ImportLog "Guest access: tenant-wide AllowToAddGuests = $($guestContext.TenantAllows)."

    $imported = 0
    $skipped = 0
    foreach ($group in $groups) {
        $groupId = [guid]$group.Id
        Write-DebugLog "Processing team: $($group.DisplayName) [$groupId]"

        if (-not (Test-GroupExistsInSql -GroupId $groupId)) {
            Write-ImportLog "Skipping '$($group.DisplayName)' [$groupId] - not yet present in dbo.Groups (run Import-M365GroupsToSql.ps1 first)."
            $skipped++
            continue
        }

        $isArchived = $false
        try {
            $team = Get-MgTeam -TeamId $group.Id -ErrorAction Stop
            $isArchived = [bool]$team.IsArchived
        }
        catch {
            Write-DebugLog "Could not read Get-MgTeam for '$($group.DisplayName)': $($_.Exception.Message)"
        }

        $allowGuests = Get-AllowToAddGuests -GroupId $group.Id -Context $guestContext

        $hidden = $null
        $welcome = $null
        $siteUrl = $null
        try {
            $identity = if ($group.MailNickname) { $group.MailNickname } else { $group.Mail }
            $ug = Get-UnifiedGroup -Identity $identity -ErrorAction Stop
            $hidden = [bool]$ug.HiddenFromAddressListsEnabled
            $welcome = [bool]$ug.WelcomeMessageEnabled
            $siteUrl = $ug.SharePointSiteUrl
        }
        catch {
            Write-DebugLog "Could not read Get-UnifiedGroup for '$($group.DisplayName)': $($_.Exception.Message)"
        }

        Write-TeamRecord -Team ([pscustomobject]@{
            GroupId = $group.Id
            Visibility = $group.Visibility
            IsArchived = $isArchived
            AllowToAddGuests = $allowGuests
            HiddenFromAddressListsEnabled = $hidden
            WelcomeMessageEnabled = $welcome
            SharePointSiteUrl = $siteUrl
            CreatedDateTime = $group.CreatedDateTime
        })
        $imported++
    }

    Write-ImportLog "Import completed: $imported team(s) imported, $skipped skipped (group not yet in dbo.Groups)."
    exit 0
}
catch {
    Write-ImportLog "Import failed: $($_.Exception.Message)"
    exit 1
}
finally {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
}
