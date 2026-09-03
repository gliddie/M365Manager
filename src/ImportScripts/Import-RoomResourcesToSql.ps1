<#
.SYNOPSIS
Imports room and equipment mailboxes into SQL Server for the M365Manager Rooms & Resources grid.

.DESCRIPTION
Populates dbo.RoomResources (see src/M365Manager.Data/Scripts/06_RoomResources.sql) from Exchange
Online: every mailbox with RecipientTypeDetails "RoomMailbox" or "EquipmentMailbox", plus its
calendar processing (delegates, booking window, in/out-of-policy lists), regional configuration
(time zone) and room-list membership.

The access model and booking policy are DERIVED from what Exchange actually reports, so rooms
created before this app existed - or changed directly in the admin center - still show up with the
right classification:
  - BookInPolicy populated            -> Restricted        (only that group may book)
  - delegate group ends in .RRS.OutOfPolicy.DE -> GeneralUseSiteDelegates
  - otherwise                         -> GeneralUseCustomDelegates
  - AllowRecurringMeetings $false     -> Hoteling, else Standard

Rooms and equipment aren't groups, so unlike Import-TeamsToSql.ps1 this needs only an Exchange
Online connection - no Microsoft Graph connection is required.

REQUIRED PERMISSIONS
- Exchange Online (app-only): Exchange Administrator or Recipient Administrator role assigned to
  the app registration (Connect-ExchangeOnline -AppId/-CertificateThumbprint/-Organization).
  See docs/AppRegistration-Setup.md for the app-only certificate setup steps.

REQUIRED MODULES
- ExchangeOnlineManagement

IMPORTANT
- Replace the values marked with TODO before running the script.
#>

[CmdletBinding()]
param(
    # --- Exchange Online (app-only) ---
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
        elseif ($value -is [int]) { $param.SqlDbType = [System.Data.SqlDbType]::Int; $param.Value = [int]$value }
        elseif ($value -is [datetime]) { $param.SqlDbType = [System.Data.SqlDbType]::DateTime2; $param.Value = [datetime]$value }
        else { $param.Value = [string]$value }
    }

    $connection.Open()
    try { $command.ExecuteNonQuery() | Out-Null }
    finally { $connection.Dispose() }
}

function Ensure-Schema {
    $sql = @"
IF OBJECT_ID(N'dbo.RoomSites', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.RoomSites (
        Id INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
        SiteCode NVARCHAR(16) NOT NULL,
        TimeZone NVARCHAR(128) NOT NULL,
        Region NVARCHAR(16) NULL,
        RegionalAdminGroup NVARCHAR(256) NULL
    );
    CREATE UNIQUE INDEX IX_RoomSites_SiteCode ON dbo.RoomSites(SiteCode);
END

IF OBJECT_ID(N'dbo.RoomResources', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.RoomResources (
        ExchangeGuid UNIQUEIDENTIFIER NOT NULL PRIMARY KEY,
        DisplayName NVARCHAR(256) NOT NULL,
        Alias NVARCHAR(128) NOT NULL,
        PrimarySmtpAddress NVARCHAR(256) NOT NULL,
        ResourceKind NVARCHAR(16) NOT NULL,
        Capacity INT NULL,
        Building NVARCHAR(128) NULL,
        Floor NVARCHAR(32) NULL,
        Office NVARCHAR(256) NULL,
        SiteCode NVARCHAR(16) NULL,
        TimeZone NVARCHAR(128) NULL,
        RoomListName NVARCHAR(256) NULL,
        DelegateGroup NVARCHAR(256) NULL,
        UsersGroup NVARCHAR(256) NULL,
        AccessModel NVARCHAR(32) NULL,
        BookingPolicy NVARCHAR(32) NULL,
        BookingWindowInDays INT NULL,
        MaximumDurationInMinutes INT NULL,
        AllowRecurringMeetings BIT NULL,
        CreatedDateTime DATETIME2(3) NULL,
        IsDeletedInM365 BIT NOT NULL DEFAULT 0,
        LastImportedAtUtc DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        LastSeenAtUtc DATETIME2(3) NULL
    );
    CREATE UNIQUE INDEX IX_RoomResources_PrimarySmtpAddress ON dbo.RoomResources(PrimarySmtpAddress);
    CREATE INDEX IX_RoomResources_SiteCode ON dbo.RoomResources(SiteCode);
END
"@
    Invoke-SqlCommand -Sql $sql
}

function Write-RoomRecord {
    param([pscustomobject]$Room)

    $sql = @"
MERGE dbo.RoomResources AS target
USING (SELECT @ExchangeGuid AS ExchangeGuid) AS source ON target.ExchangeGuid = source.ExchangeGuid
WHEN MATCHED THEN
    UPDATE SET
        DisplayName = @DisplayName, Alias = @Alias, PrimarySmtpAddress = @PrimarySmtpAddress,
        ResourceKind = @ResourceKind, Capacity = @Capacity, Building = @Building, Floor = @Floor,
        Office = @Office, SiteCode = @SiteCode, TimeZone = @TimeZone, RoomListName = @RoomListName,
        DelegateGroup = @DelegateGroup, UsersGroup = @UsersGroup, AccessModel = @AccessModel,
        BookingPolicy = @BookingPolicy, BookingWindowInDays = @BookingWindowInDays,
        MaximumDurationInMinutes = @MaximumDurationInMinutes, AllowRecurringMeetings = @AllowRecurringMeetings,
        CreatedDateTime = @CreatedDateTime, IsDeletedInM365 = 0,
        LastImportedAtUtc = SYSUTCDATETIME(), LastSeenAtUtc = SYSUTCDATETIME()
WHEN NOT MATCHED THEN
    INSERT (ExchangeGuid, DisplayName, Alias, PrimarySmtpAddress, ResourceKind, Capacity, Building, Floor,
            Office, SiteCode, TimeZone, RoomListName, DelegateGroup, UsersGroup, AccessModel, BookingPolicy,
            BookingWindowInDays, MaximumDurationInMinutes, AllowRecurringMeetings, CreatedDateTime,
            IsDeletedInM365, LastImportedAtUtc, LastSeenAtUtc)
    VALUES (@ExchangeGuid, @DisplayName, @Alias, @PrimarySmtpAddress, @ResourceKind, @Capacity, @Building, @Floor,
            @Office, @SiteCode, @TimeZone, @RoomListName, @DelegateGroup, @UsersGroup, @AccessModel, @BookingPolicy,
            @BookingWindowInDays, @MaximumDurationInMinutes, @AllowRecurringMeetings, @CreatedDateTime,
            0, SYSUTCDATETIME(), SYSUTCDATETIME());
"@

    $parameters = @{
        ExchangeGuid = [guid]$Room.ExchangeGuid
        DisplayName = $Room.DisplayName
        Alias = $Room.Alias
        PrimarySmtpAddress = $Room.PrimarySmtpAddress
        ResourceKind = $Room.ResourceKind
        Capacity = $Room.Capacity
        Building = $Room.Building
        Floor = $Room.Floor
        Office = $Room.Office
        SiteCode = $Room.SiteCode
        TimeZone = $Room.TimeZone
        RoomListName = $Room.RoomListName
        DelegateGroup = $Room.DelegateGroup
        UsersGroup = $Room.UsersGroup
        AccessModel = $Room.AccessModel
        BookingPolicy = $Room.BookingPolicy
        BookingWindowInDays = $Room.BookingWindowInDays
        MaximumDurationInMinutes = $Room.MaximumDurationInMinutes
        AllowRecurringMeetings = $Room.AllowRecurringMeetings
        CreatedDateTime = $Room.CreatedDateTime
    }

    Invoke-SqlCommand -Sql $sql -Parameters $parameters
}

function Update-MissingRoomFlags {
    param([string[]]$SeenGuids)

    # Rooms that were in the cache but no longer come back from Exchange are flagged rather than
    # deleted, so the grid can still show what happened (same convention as dbo.Groups).
    if ($SeenGuids.Count -eq 0) { return }

    $list = ($SeenGuids | ForEach-Object { "'" + $_ + "'" }) -join ","
    $sql = "UPDATE dbo.RoomResources SET IsDeletedInM365 = 1 WHERE ExchangeGuid NOT IN ($list)"
    Invoke-SqlCommand -Sql $sql
}

function Get-SiteCodeFromName {
    param([string]$DisplayName)

    if ([string]::IsNullOrWhiteSpace($DisplayName)) { return $null }
    $space = $DisplayName.IndexOf(' ')
    if ($space -lt 1) { return $null }
    return $DisplayName.Substring(0, $space).ToUpper()
}

function Get-RoomListForMember {
    param([string]$Address, [hashtable]$RoomListMembers)

    foreach ($listName in $RoomListMembers.Keys) {
        if ($RoomListMembers[$listName] -contains $Address.ToLower()) { return $listName }
    }
    return $null
}

try {
    if (-not [string]::IsNullOrWhiteSpace($LogFile)) { Rotate-ImportLogs }

    Write-ImportLog 'Ensuring SQL schema...'
    Ensure-Schema

    Write-ImportLog 'Connecting to Exchange Online...'
    $attempt = 0
    $connected = $false
    while (-not $connected -and $attempt -lt $MaxRetries) {
        $attempt++
        try { Connect-ExchangeAppOnly; $connected = $true }
        catch {
            Write-ImportLog "Exchange Online connection attempt $attempt failed: $($_.Exception.Message)"
            if ($attempt -ge $MaxRetries) { throw }
            Start-Sleep -Seconds 10
        }
    }

    # Room-list membership is read once up front rather than per room - one call per list instead
    # of one per mailbox.
    Write-ImportLog 'Reading room lists...'
    $roomListMembers = @{}
    try {
        $roomLists = Get-DistributionGroup -RecipientTypeDetails RoomList -ResultSize Unlimited -ErrorAction Stop
        foreach ($list in $roomLists) {
            $members = @()
            try {
                $members = (Get-DistributionGroupMember -Identity $list.Identity -ResultSize Unlimited -ErrorAction Stop |
                            ForEach-Object { $_.PrimarySmtpAddress.ToString().ToLower() })
            }
            catch {
                Write-DebugLog "Could not read members of room list '$($list.DisplayName)': $($_.Exception.Message)"
            }
            $roomListMembers[$list.DisplayName] = $members
        }
        Write-ImportLog "Found $($roomLists.Count) room list(s)."
    }
    catch {
        Write-ImportLog "WARNING: Could not enumerate room lists ($($_.Exception.Message)) - RoomListName will be left NULL."
    }

    Write-ImportLog 'Querying room and equipment mailboxes...'
    $mailboxes = Get-EXOMailbox -RecipientTypeDetails RoomMailbox,EquipmentMailbox -ResultSize Unlimited `
        -Properties ExchangeGuid, DisplayName, Alias, PrimarySmtpAddress, RecipientTypeDetails, ResourceCapacity, Office, WhenMailboxCreated -ErrorAction Stop
    Write-ImportLog "Found $($mailboxes.Count) room/equipment mailbox(es)."

    $imported = 0
    $seenGuids = @()

    foreach ($mailbox in $mailboxes) {
        Write-DebugLog "Processing: $($mailbox.DisplayName) [$($mailbox.ExchangeGuid)]"
        $address = $mailbox.PrimarySmtpAddress.ToString()

        $kind = if ($mailbox.RecipientTypeDetails -eq 'EquipmentMailbox') { 'Equipment' } else { 'Room' }

        $timeZone = $null
        try { $timeZone = (Get-MailboxRegionalConfiguration -Identity $address -ErrorAction Stop).TimeZone }
        catch { Write-DebugLog "No regional configuration for '$($mailbox.DisplayName)': $($_.Exception.Message)" }

        $delegateGroup = $null
        $usersGroup = $null
        $accessModel = $null
        $bookingPolicy = $null
        $window = $null
        $duration = $null
        $recurring = $null

        try {
            $cp = Get-CalendarProcessing -Identity $address -ErrorAction Stop

            $window = $cp.BookingWindowInDays
            $duration = $cp.MaximumDurationInMinutes
            $recurring = [bool]$cp.AllowRecurringMeetings

            if ($cp.ResourceDelegates -and $cp.ResourceDelegates.Count -gt 0) {
                $delegateGroup = $cp.ResourceDelegates[0].ToString()
            }
            if ($cp.BookInPolicy -and $cp.BookInPolicy.Count -gt 0) {
                $usersGroup = $cp.BookInPolicy[0].ToString()
            }

            # Derive the access model the same way the app models it.
            if ($usersGroup) {
                $accessModel = 'Restricted'
            }
            elseif ($delegateGroup -and $delegateGroup -match '\.RRS\.OutOfPolicy\.DE$') {
                $accessModel = 'GeneralUseSiteDelegates'
            }
            else {
                $accessModel = 'GeneralUseCustomDelegates'
            }

            $bookingPolicy = if ($recurring) { 'Standard' } else { 'Hoteling' }
        }
        catch {
            Write-DebugLog "No calendar processing for '$($mailbox.DisplayName)': $($_.Exception.Message)"
        }

        # Office is stored as free text ("Building 3, Floor 2"); split it back out best-effort.
        $building = $null
        $floor = $null
        if ($mailbox.Office) {
            if ($mailbox.Office -match 'Building\s+([^,]+)') { $building = $Matches[1].Trim() }
            if ($mailbox.Office -match 'Floor\s+(\S+)') { $floor = $Matches[1].Trim() }
            elseif ($mailbox.Office -match 'Ground Floor') { $floor = '0' }
        }

        Write-RoomRecord -Room ([pscustomobject]@{
            ExchangeGuid = $mailbox.ExchangeGuid
            DisplayName = $mailbox.DisplayName
            Alias = $mailbox.Alias
            PrimarySmtpAddress = $address
            ResourceKind = $kind
            Capacity = if ($null -ne $mailbox.ResourceCapacity) { [int]$mailbox.ResourceCapacity } else { $null }
            Building = $building
            Floor = $floor
            Office = $mailbox.Office
            SiteCode = Get-SiteCodeFromName -DisplayName $mailbox.DisplayName
            TimeZone = $timeZone
            RoomListName = Get-RoomListForMember -Address $address -RoomListMembers $roomListMembers
            DelegateGroup = $delegateGroup
            UsersGroup = $usersGroup
            AccessModel = $accessModel
            BookingPolicy = $bookingPolicy
            BookingWindowInDays = if ($null -ne $window) { [int]$window } else { $null }
            MaximumDurationInMinutes = if ($null -ne $duration) { [int]$duration } else { $null }
            AllowRecurringMeetings = $recurring
            CreatedDateTime = $mailbox.WhenMailboxCreated
        })

        $seenGuids += ([guid]$mailbox.ExchangeGuid).ToString()
        $imported++
    }

    Update-MissingRoomFlags -SeenGuids $seenGuids

    Write-ImportLog "Import completed: $imported room(s)/resource(s) imported."
    exit 0
}
catch {
    Write-ImportLog "Import failed: $($_.Exception.Message)"
    exit 1
}
finally {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
}
