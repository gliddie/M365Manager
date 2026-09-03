<#
.SYNOPSIS
Imports Shared Mailboxes into SQL Server for the M365Manager Shared Mailboxes overview grid.

.DESCRIPTION
Populates dbo.SharedMailboxes (see src/M365Manager.Data/Scripts/05_SharedMailboxes.sql) from
Exchange Online: every mailbox with RecipientTypeDetails "SharedMailbox", plus its .ED/.AU/.RE
access-group addresses and the .ED group's ManagedBy (recorded as the mailbox's "Owners").

All MBX.* access groups are read once at the start and held in an in-memory index keyed by both
group name and SMTP address. Each mailbox is then matched against that index by its derived
address and, failing that, by its derived name - the second form catches groups that were renamed
to the current convention but kept their original address. Only when the Editor group is still
unresolved does the script ask Exchange which groups actually hold permission on that mailbox,
which costs a round-trip and is therefore reserved for the few odd cases.

A mailbox with no access groups at all is normal and is imported with those columns empty.

Shared mailboxes aren't Microsoft 365 groups, so unlike Import-TeamsToSql.ps1 this script only
needs an Exchange Online connection - no Microsoft Graph connection is required.

REQUIRED PERMISSIONS
- Exchange Online (app-only): Exchange Administrator or Recipient Administrator role assigned to
  the app registration (Connect-ExchangeOnline -AppId/-CertificateThumbprint/-Organization).
  See docs/AppRegistration-Setup.md for the app-only certificate setup steps.

REQUIRED MODULES
- ExchangeOnlineManagement

IMPORTANT
- Replace the values marked with TODO before running the script.
- If a mailbox is not found in dbo.SharedMailboxes, the GUI simply shows fewer rows for it (no
  live M365 fallback for the overview grid - that's this script's job).
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
        elseif ($value -is [datetime]) { $param.SqlDbType = [System.Data.SqlDbType]::DateTime2; $param.Value = [datetime]$value }
        else { $param.Value = [string]$value }
    }

    $connection.Open()
    try { $command.ExecuteNonQuery() | Out-Null }
    finally { $connection.Dispose() }
}

function Ensure-Schema {
    $sql = @"
IF OBJECT_ID(N'dbo.SharedMailboxes', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SharedMailboxes (
        ExchangeGuid UNIQUEIDENTIFIER NOT NULL PRIMARY KEY,
        DisplayName NVARCHAR(256) NOT NULL,
        Alias NVARCHAR(128) NOT NULL,
        PrimarySmtpAddress NVARCHAR(256) NOT NULL,
        EDGroupAddress NVARCHAR(256) NULL,
        AUGroupAddress NVARCHAR(256) NULL,
        REGroupAddress NVARCHAR(256) NULL,
        Owners NVARCHAR(MAX) NULL,
        RequireSenderAuthenticationEnabled BIT NULL,
        CreatedDateTime DATETIME2(3) NULL,
        IsDeletedInM365 BIT NOT NULL DEFAULT 0,
        LastImportedAtUtc DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        LastSeenAtUtc DATETIME2(3) NULL
    );
    CREATE UNIQUE INDEX IX_SharedMailboxes_PrimarySmtpAddress ON dbo.SharedMailboxes(PrimarySmtpAddress);
END
"@
    Invoke-SqlCommand -Sql $sql
}

$Script:AccessGroupIndex = $null

function Initialize-AccessGroupIndex {
    <#
      Loads every MBX.* group once into an in-memory index, keyed by both its name and its primary
      SMTP address (lower-cased).

      This replaces the previous three Get-DistributionGroup calls per mailbox. With a couple of
      thousand shared mailboxes that was ~6,700 Exchange round-trips per run - slow enough to take
      the better part of an hour and a reliable way to get throttled. It also removes the need to
      tell "does not exist" apart from "the call failed" by parsing error text: a group is simply
      either in the index or it isn't.
    #>
    $index = @{}
    $groups = Get-DistributionGroup -ResultSize Unlimited -Filter "Name -like 'MBX.*'" -ErrorAction Stop

    foreach ($group in $groups) {
        $address = $group.PrimarySmtpAddress.ToString()
        $index[$group.Name.ToLower()] = $group
        $index[$address.ToLower()] = $group
    }

    $Script:AccessGroupIndex = $index
    Write-ImportLog "Indexed $($groups.Count) MBX.* access group(s)."
}

function Resolve-AccessGroup {
    <#
      Looks an access group up in the index by name or address.

      .Found   - $true when the group exists
      .Owners  - ManagedBy entries when found
      .Address - the group's REAL primary SMTP address, which is not necessarily the identity it
                 was looked up by (a renamed group keeps its original address)
    #>
    param([string]$GroupIdentity)

    $miss = [pscustomobject]@{ Found = $false; Owners = @(); Address = $null }
    if ([string]::IsNullOrWhiteSpace($GroupIdentity)) { return $miss }

    $group = $Script:AccessGroupIndex[$GroupIdentity.ToLower()]
    if (-not $group) { return $miss }

    return [pscustomobject]@{
        Found   = $true
        Owners  = @($group.ManagedBy)
        Address = $group.PrimarySmtpAddress.ToString()
    }
}

function Get-GroupsFromMailboxPermissions {
    <#
      Returns the access groups actually granted FullAccess on the mailbox, as a hashtable of
      tier -> group name.

      Deriving the group address from the mailbox display name is not reliable on its own:
      renaming a group's display name in the admin center does NOT change its SMTP address, so a
      group renamed to "MBX.UL Arkansas Communication.ED" still answers to its original
      "MBX.UL.ArkansasCommunication.ED@ul.com" and a lookup by the derived address misses it. The
      mailbox's own permissions are the ground truth - same approach ShrMbxChgOwner.ps1 used.
    #>
    param([string]$MailboxAddress)

    $result = @{}
    try {
        $perms = Get-MailboxPermission -Identity $MailboxAddress -ErrorAction Stop |
                 Where-Object { $_.User -like 'MBX*' }

        foreach ($perm in $perms) {
            $leaf = ($perm.User.ToString() -split '/')[-1].Trim()
            foreach ($tier in @('ED', 'AU', 'RE')) {
                if ($leaf -like "*.$tier" -and -not $result.ContainsKey($tier)) {
                    $result[$tier] = $leaf
                }
            }
        }
    }
    catch {
        Write-DebugLog "Could not read permissions of '$MailboxAddress': $($_.Exception.Message)"
    }
    return $result
}

function Update-MissingMailboxFlags {
    param([string[]]$SeenGuids)

    # Shared mailboxes that were in the cache but no longer come back from Exchange are flagged
    # rather than deleted, so the grid can still show what happened (same convention as dbo.Groups
    # and the rooms importer).
    if ($SeenGuids.Count -eq 0) { return }

    $list = ($SeenGuids | ForEach-Object { "'" + $_ + "'" }) -join ","
    $sql = "UPDATE dbo.SharedMailboxes SET IsDeletedInM365 = 1 WHERE ExchangeGuid NOT IN ($list)"
    Invoke-SqlCommand -Sql $sql
}

function Write-SharedMailboxRecord {
    param([pscustomobject]$Mailbox)

    $sql = @"
MERGE dbo.SharedMailboxes AS target
USING (SELECT @ExchangeGuid AS ExchangeGuid) AS source ON target.ExchangeGuid = source.ExchangeGuid
WHEN MATCHED THEN
    UPDATE SET
        DisplayName = @DisplayName,
        Alias = @Alias,
        PrimarySmtpAddress = @PrimarySmtpAddress,
        EDGroupAddress = @EDGroupAddress,
        AUGroupAddress = @AUGroupAddress,
        REGroupAddress = @REGroupAddress,
        Owners = @Owners,
        RequireSenderAuthenticationEnabled = @RequireSenderAuthenticationEnabled,
        CreatedDateTime = @CreatedDateTime,
        IsDeletedInM365 = 0,
        LastImportedAtUtc = SYSUTCDATETIME(),
        LastSeenAtUtc = SYSUTCDATETIME()
WHEN NOT MATCHED THEN
    INSERT (ExchangeGuid, DisplayName, Alias, PrimarySmtpAddress, EDGroupAddress, AUGroupAddress, REGroupAddress, Owners, RequireSenderAuthenticationEnabled, CreatedDateTime, IsDeletedInM365, LastImportedAtUtc, LastSeenAtUtc)
    VALUES (@ExchangeGuid, @DisplayName, @Alias, @PrimarySmtpAddress, @EDGroupAddress, @AUGroupAddress, @REGroupAddress, @Owners, @RequireSenderAuthenticationEnabled, @CreatedDateTime, 0, SYSUTCDATETIME(), SYSUTCDATETIME());
"@

    $parameters = @{
        ExchangeGuid = [guid]$Mailbox.ExchangeGuid
        DisplayName = $Mailbox.DisplayName
        Alias = $Mailbox.Alias
        PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
        EDGroupAddress = $Mailbox.EDGroupAddress
        AUGroupAddress = $Mailbox.AUGroupAddress
        REGroupAddress = $Mailbox.REGroupAddress
        Owners = $Mailbox.Owners
        RequireSenderAuthenticationEnabled = $Mailbox.RequireSenderAuthenticationEnabled
        CreatedDateTime = $Mailbox.CreatedDateTime
    }

    Invoke-SqlCommand -Sql $sql -Parameters $parameters
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

    Write-ImportLog 'Indexing access groups...'
    Initialize-AccessGroupIndex

    Write-ImportLog 'Querying shared mailboxes...'
    $mailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited `
        -Properties ExchangeGuid, DisplayName, Alias, PrimarySmtpAddress, RequireSenderAuthenticationEnabled, WhenMailboxCreated -ErrorAction Stop
    Write-ImportLog "Found $($mailboxes.Count) shared mailboxes."

    $imported = 0
    $skipped = 0
    $seenGuids = @()
    foreach ($mailbox in $mailboxes) {
        Write-DebugLog "Processing shared mailbox: $($mailbox.DisplayName) [$($mailbox.ExchangeGuid)]"

        $atIndex = $mailbox.PrimarySmtpAddress.IndexOf('@')
        if ($atIndex -lt 1) {
            Write-DebugLog "Skipping '$($mailbox.DisplayName)' - PrimarySmtpAddress has no domain."
            $skipped++
            continue
        }
        $domain = $mailbox.PrimarySmtpAddress.Substring($atIndex + 1)

        # Access groups are named after the mailbox's DISPLAY name - "MBX.<display name>.ED", with
        # the display name's spaces kept in the group NAME and removed in the alias/ADDRESS. The
        # index holds both keys, so trying each in turn also catches a group that was renamed to the
        # current convention but kept its original SMTP address.
        $aliasStem = ($mailbox.DisplayName -replace ' ', '')

        $ed = Resolve-AccessGroup -GroupIdentity "MBX.$aliasStem.ED@$domain"
        if (-not $ed.Found) { $ed = Resolve-AccessGroup -GroupIdentity "MBX.$($mailbox.DisplayName).ED" }

        $au = Resolve-AccessGroup -GroupIdentity "MBX.$aliasStem.AU@$domain"
        if (-not $au.Found) { $au = Resolve-AccessGroup -GroupIdentity "MBX.$($mailbox.DisplayName).AU" }

        $re = Resolve-AccessGroup -GroupIdentity "MBX.$aliasStem.RE@$domain"
        if (-not $re.Found) { $re = Resolve-AccessGroup -GroupIdentity "MBX.$($mailbox.DisplayName).RE" }

        # Only when the Editor group - the one carrying the owners - still hasn't been found is it
        # worth asking Exchange what actually holds permission on this mailbox. That call costs a
        # round-trip per mailbox, so it is reserved for the groups whose names follow neither form.
        if (-not $ed.Found) {
            $granted = Get-GroupsFromMailboxPermissions -MailboxAddress $mailbox.PrimarySmtpAddress
            if ($granted.ContainsKey('ED')) { $ed = Resolve-AccessGroup -GroupIdentity $granted['ED'] }
            if (-not $au.Found -and $granted.ContainsKey('AU')) { $au = Resolve-AccessGroup -GroupIdentity $granted['AU'] }
            if (-not $re.Found -and $granted.ContainsKey('RE')) { $re = Resolve-AccessGroup -GroupIdentity $granted['RE'] }
        }

        # A mailbox with no access groups at all is perfectly normal - plenty of shared mailboxes
        # were never given the .ED/.AU/.RE structure. It is written with empty group columns, not
        # skipped.
        $owners = if ($ed.Owners.Count -gt 0) { ($ed.Owners -join ';') } else { $null }

        Write-SharedMailboxRecord -Mailbox ([pscustomobject]@{
            ExchangeGuid = $mailbox.ExchangeGuid
            DisplayName = $mailbox.DisplayName
            Alias = $mailbox.Alias
            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
            EDGroupAddress = if ($ed.Found) { $ed.Address } else { $null }
            AUGroupAddress = if ($au.Found) { $au.Address } else { $null }
            REGroupAddress = if ($re.Found) { $re.Address } else { $null }
            Owners = $owners
            RequireSenderAuthenticationEnabled = $mailbox.RequireSenderAuthenticationEnabled
            CreatedDateTime = $mailbox.WhenMailboxCreated
        })
        $seenGuids += ([guid]$mailbox.ExchangeGuid).ToString()
        $imported++
    }

    Update-MissingMailboxFlags -SeenGuids $seenGuids

    Write-ImportLog "Import completed: $imported shared mailbox(es) imported, $skipped skipped."
    exit 0
}
catch {
    Write-ImportLog "Import failed: $($_.Exception.Message)"
    exit 1
}
finally {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
}
