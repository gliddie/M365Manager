<#
.SYNOPSIS
Imports Microsoft 365 groups and their members into SQL Server for fast GUI lookups.

.DESCRIPTION
This script connects to Microsoft 365 using a registered app with certificate-based authentication,
queries all groups (distribution, mail-enabled security, security, Microsoft 365 / Unified, dynamic) with
their DIRECT members, and stores the data in SQL Server.

It is designed to run on a separate server as a scheduled task (once a day at the moment).

SPEED (rebuilt 2026-10): one SQL connection for the run, each group's members replaced with one
DELETE + one bulk insert, and members fetched for 20 groups per Graph $batch request. The first version
opened a connection per statement (two per member row), MERGEd row by row and re-expanded every
nested group under each parent - about nine hours for ~37,000 groups.

REQUIRED PERMISSIONS FOR THE REGISTERED APP
- Microsoft Graph:
  - GroupMember.Read.All
  - Group.Read.All
  - Directory.Read.All
  - User.Read.All
- If the app should also change groups later, add:
  - Group.ReadWrite.All
- Exchange Online PowerShell (Role-based access control):
  - Exchange Administrator or Compliance Administrator (only needed for Exchange-specific discovery paths)

IMPORTANT
- Replace the values marked with TODO before running the script.
- The script uses a certificate-based app-only authentication flow.
- For managed environments, you may also need to consent to these permissions in Entra ID.

NOTES
- Direct members only. Nested groups are not expanded: every group is imported in its own right, so a
  nested group's members are already under that group. A member that is a group gets IsNested = 1 and
  Path = '<group> > <member group>' (the app reads IsNested = 0 as the user members).
- Members removed in M365 disappear from the cache on the next run (rows are replaced, not merged).
- A group deleted while the run is going is skipped and flagged IsDeletedInM365; other per-group errors
  are logged, the run continues and ends with exit code 1 (stops after 50 such errors).
- If a group is not found in SQL, the GUI can fall back to live Microsoft 365 lookups.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$TenantId = "TODO_TENANT_ID",

    [Parameter()]
    [string]$ClientId = "TODO_CLIENT_ID",

    [Parameter()]
    [string]$CertificateThumbprint = "TODO_CERTIFICATE_THUMBPRINT",

    [Parameter()]
    [string]$SqlServer = "TODO_SQL_SERVER",

    [Parameter()]
    [string]$SqlDatabase = "M365Manager",

    [Parameter()]
    [string]$SqlCredentialPath = "D:\SCRIPTS\Secure\m365manager-sql.xml",

    [Parameter()]
    [switch]$UseIntegratedSecurity = $false,

    [Parameter()]
    [switch]$UseSqlEncryption = $true,

    [Parameter()]
    [string]$GraphScope = "https://graph.microsoft.com/.default",

    [Parameter()]
    [int]$BatchSize = 100,

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
$script:SqlConnection = $null

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

    if ($VerboseDebug) {
        Write-ImportLog $Message
    }
}

function Rotate-ImportLogs {
    if ([string]::IsNullOrWhiteSpace($LogFile)) {
        return
    }

    $logDirectory = Split-Path -Path $LogFile -Parent
    if ([string]::IsNullOrWhiteSpace($logDirectory)) {
        return
    }

    if (-not (Test-Path -Path $logDirectory)) {
        return
    }

    $files = Get-ChildItem -Path $logDirectory -Filter "*.log" -File | Sort-Object LastWriteTimeUtc -Descending
    $filesToDelete = $files | Where-Object { $_.LastWriteTimeUtc -lt (Get-Date).AddDays(-$LogRetentionDays) }
    foreach ($file in $filesToDelete) {
        try { Remove-Item -Path $file.FullName -Force -ErrorAction Stop } catch { }
    }

    $files = Get-ChildItem -Path $logDirectory -Filter "*.log" -File | Sort-Object LastWriteTimeUtc -Descending
    if ($files.Count -gt $MaxLogFiles) {
        $filesToRemove = $files | Select-Object -Skip $MaxLogFiles
        foreach ($file in $filesToRemove) {
            try { Remove-Item -Path $file.FullName -Force -ErrorAction Stop } catch { }
        }
    }
}

function Get-GraphAccessToken {
    param([string]$TenantId, [string]$ClientId, [string]$CertificateThumbprint)

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft Graph PowerShell module is not installed. Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    Write-DebugLog 'Checking for existing Microsoft Graph context...'
    $existingContext = Get-MgContext -ErrorAction SilentlyContinue
    if ($existingContext) {
        return $existingContext
    }

    Write-DebugLog 'Connecting to Microsoft Graph with certificate thumbprint...'
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome | Out-Null

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        throw 'Microsoft Graph connection could not be established.'
    }

    return $context
}

function Read-SqlCredential {
    # Export-Clixml protects the password with DPAPI: only the Windows account that wrote the file,
    # on the machine it was written on, can read it back. Create it AS the task's account:
    #   Get-Credential | Export-Clixml D:\SCRIPTS\Secure\m365manager-sql.xml
    $account = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    if (-not (Test-Path -LiteralPath $SqlCredentialPath)) {
        throw "SQL credential file not found: $SqlCredentialPath (running as $account)."
    }

    try {
        $credential = Import-Clixml -LiteralPath $SqlCredentialPath
    }
    catch {
        throw "SQL credential file $SqlCredentialPath could not be decrypted as $account. Only the Windows account that created it can read it, and only on the same machine. $($_.Exception.Message)"
    }

    if ($credential -isnot [pscredential]) {
        throw "SQL credential file $SqlCredentialPath does not contain a credential (create it with Get-Credential | Export-Clixml)."
    }

    return $credential
}

function Get-ConnectionString {
    # Built once per run: the credential file is decrypted on the first SQL command, every later
    # command reuses the result.
    if ($script:ConnectionString) { return $script:ConnectionString }

    # The builder escapes the values - a ';' or '=' in the password broke the old string concatenation.
    $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $builder['Data Source'] = $SqlServer
    $builder['Initial Catalog'] = $SqlDatabase
    $builder['Encrypt'] = [bool]$UseSqlEncryption
    $builder['TrustServerCertificate'] = $true

    if ($UseIntegratedSecurity) {
        $builder['Integrated Security'] = $true
    }
    else {
        $credential = Read-SqlCredential
        $builder['User ID'] = $credential.UserName
        $builder['Password'] = $credential.GetNetworkCredential().Password
    }

    $script:ConnectionString = $builder.ConnectionString
    return $script:ConnectionString
}

function Invoke-SqlCommand {
    param(
        [string]$Sql,
        [hashtable]$Parameters = @{},
        # SqlClient's default is 30 s. Fine per row; not for set-based statements over the whole table.
        [int]$TimeoutSeconds = 30,
        # Return the affected row count instead of discarding it.
        [switch]$ReturnRowCount
    )

    $connection = Get-SqlConnection
    $command = $connection.CreateCommand()
    $command.CommandText = $Sql
    $command.CommandTimeout = $TimeoutSeconds
    foreach ($key in $Parameters.Keys) {
        $value = $Parameters[$key]
        $param = $command.Parameters.Add("@$key", [System.Data.SqlDbType]::NVarChar)

        if ($null -eq $value) {
            $param.Value = [System.DBNull]::Value
            continue
        }

        if ($value -is [guid]) {
            $param.SqlDbType = [System.Data.SqlDbType]::UniqueIdentifier
            $param.Value = [guid]$value
        }
        elseif ($value -is [bool]) {
            $param.SqlDbType = [System.Data.SqlDbType]::Bit
            $param.Value = [bool]$value
        }
        elseif ($value -is [int] -or $value -is [long] -or $value -is [short] -or $value -is [byte]) {
            $param.SqlDbType = [System.Data.SqlDbType]::Int
            $param.Value = [int]$value
        }
        elseif ($value -is [datetime]) {
            $param.SqlDbType = [System.Data.SqlDbType]::DateTime2
            $param.Value = [datetime]$value
        }
        else {
            $param.Value = [string]$value
        }
    }

    try {
        $affected = $command.ExecuteNonQuery()
        if ($ReturnRowCount) { return $affected }
    }
    finally {
        $command.Dispose()
    }
}

function Get-SqlConnection {
    # One connection for the whole run. Opening one per statement - two statements per member row -
    # was the bulk of the old nine-hour runtime. Reopened if the server dropped it.
    if ($null -eq $script:SqlConnection -or $script:SqlConnection.State -ne [System.Data.ConnectionState]::Open) {
        if ($null -ne $script:SqlConnection) { $script:SqlConnection.Dispose() }
        $script:SqlConnection = New-Object System.Data.SqlClient.SqlConnection(Get-ConnectionString)
        $script:SqlConnection.Open()
    }
    return $script:SqlConnection
}

function Get-SqlUtcNow {
    # The SQL server's clock, not this machine's: LastSeenAtUtc is stamped with SYSUTCDATETIME()
    # on the server, so the run start has to come from the same clock.
    $connection = New-Object System.Data.SqlClient.SqlConnection(Get-ConnectionString)
    $command = $connection.CreateCommand()
    $command.CommandText = 'SELECT SYSUTCDATETIME()'
    $connection.Open()
    try { return [datetime]$command.ExecuteScalar() }
    finally { $connection.Dispose() }
}

function Ensure-Schema {
    $sql = @"
IF OBJECT_ID(N'dbo.Groups', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Groups (
        Id UNIQUEIDENTIFIER NOT NULL PRIMARY KEY,
        ExternalId NVARCHAR(256) NULL,
        DisplayName NVARCHAR(256) NULL,
        MailNickname NVARCHAR(128) NULL,
        Alias NVARCHAR(128) NULL,
        PrimarySmtpAddress NVARCHAR(256) NULL,
        GroupTypes NVARCHAR(512) NULL,
        RecipientTypeDetails NVARCHAR(128) NULL,
        MailEnabled BIT NOT NULL DEFAULT 0,
        SecurityEnabled BIT NOT NULL DEFAULT 0,
        IsAssignableToRole BIT NOT NULL DEFAULT 0,
        IsHiddenInOutlookClients BIT NOT NULL DEFAULT 0,
        IsMembershipDynamic BIT NOT NULL DEFAULT 0,
        MembershipRule NVARCHAR(4000) NULL,
        Notes NVARCHAR(4000) NULL,
        ManagedBy NVARCHAR(MAX) NULL,
        CreatedDateTime DATETIME2(3) NULL,
        UpdatedDateTime DATETIME2(3) NULL,
        LastImportedAtUtc DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        IsDeletedInM365 BIT NOT NULL DEFAULT 0,
        LastSeenAtUtc DATETIME2(3) NULL
    );
END

IF OBJECT_ID(N'dbo.GroupMembers', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.GroupMembers (
        Id BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        GroupId UNIQUEIDENTIFIER NOT NULL,
        MemberType NVARCHAR(64) NOT NULL,
        MemberId NVARCHAR(256) NULL,
        DisplayName NVARCHAR(256) NULL,
        PrimarySmtpAddress NVARCHAR(256) NULL,
        UserPrincipalName NVARCHAR(256) NULL,
        UserType NVARCHAR(64) NULL,
        IsNested BIT NOT NULL DEFAULT 0,
        Path NVARCHAR(2048) NULL,
        LastImportedAtUtc DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        CONSTRAINT FK_GroupMembers_Groups FOREIGN KEY (GroupId) REFERENCES dbo.Groups(Id)
    );
END

IF COL_LENGTH(N'dbo.Groups', N'IsDeletedInM365') IS NULL
BEGIN
    ALTER TABLE dbo.Groups ADD IsDeletedInM365 BIT NOT NULL CONSTRAINT DF_Groups_IsDeletedInM365 DEFAULT (0);
END

IF COL_LENGTH(N'dbo.Groups', N'LastSeenAtUtc') IS NULL
BEGIN
    ALTER TABLE dbo.Groups ADD LastSeenAtUtc DATETIME2(3) NULL;
END

IF COL_LENGTH(N'dbo.GroupMembers', N'UserPrincipalName') IS NULL
BEGIN
    ALTER TABLE dbo.GroupMembers ADD UserPrincipalName NVARCHAR(256) NULL;
END

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_Groups_PrimarySmtpAddress' AND object_id = OBJECT_ID(N'dbo.Groups'))
    CREATE INDEX IX_Groups_PrimarySmtpAddress ON dbo.Groups (PrimarySmtpAddress);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_GroupMembers_GroupId' AND object_id = OBJECT_ID(N'dbo.GroupMembers'))
    CREATE INDEX IX_GroupMembers_GroupId ON dbo.GroupMembers (GroupId);
"@

    Invoke-SqlCommand -Sql $sql
}

function Write-GroupRecord {
    param(
        [pscustomobject]$Group,
        [string]$GroupKey,
        [string]$GroupTypeSummary
    )

    $sql = @"
MERGE dbo.Groups AS target
USING (SELECT @Id AS Id) AS source ON target.Id = source.Id
WHEN MATCHED THEN
    UPDATE SET
        ExternalId = @ExternalId,
        DisplayName = @DisplayName,
        MailNickname = @MailNickname,
        Alias = @Alias,
        PrimarySmtpAddress = @PrimarySmtpAddress,
        GroupTypes = @GroupTypes,
        RecipientTypeDetails = @RecipientTypeDetails,
        MailEnabled = @MailEnabled,
        SecurityEnabled = @SecurityEnabled,
        IsAssignableToRole = @IsAssignableToRole,
        IsHiddenInOutlookClients = @IsHiddenInOutlookClients,
        IsMembershipDynamic = @IsMembershipDynamic,
        MembershipRule = @MembershipRule,
        Notes = @Notes,
        ManagedBy = @ManagedBy,
        UpdatedDateTime = SYSUTCDATETIME(),
        LastImportedAtUtc = SYSUTCDATETIME(),
        IsDeletedInM365 = 0,
        LastSeenAtUtc = SYSUTCDATETIME()
WHEN NOT MATCHED THEN
    INSERT (Id, ExternalId, DisplayName, MailNickname, Alias, PrimarySmtpAddress, GroupTypes, RecipientTypeDetails, MailEnabled, SecurityEnabled, IsAssignableToRole, IsHiddenInOutlookClients, IsMembershipDynamic, MembershipRule, Notes, ManagedBy, CreatedDateTime, UpdatedDateTime, LastImportedAtUtc, IsDeletedInM365, LastSeenAtUtc)
    VALUES (@Id, @ExternalId, @DisplayName, @MailNickname, @Alias, @PrimarySmtpAddress, @GroupTypes, @RecipientTypeDetails, @MailEnabled, @SecurityEnabled, @IsAssignableToRole, @IsHiddenInOutlookClients, @IsMembershipDynamic, @MembershipRule, @Notes, @ManagedBy, SYSUTCDATETIME(), SYSUTCDATETIME(), SYSUTCDATETIME(), 0, SYSUTCDATETIME());
"@

    $parameters = @{
        Id = [guid]$Group.id
        ExternalId = $Group.id
        DisplayName = $Group.displayName
        MailNickname = $Group.mailNickname
        Alias = $Group.mailNickname
        PrimarySmtpAddress = $Group.mail
        GroupTypes = ($Group.groupTypes -join ',')
        RecipientTypeDetails = $Group.recipientTypeDetails
        MailEnabled = [int]($Group.mailEnabled -eq $true)
        SecurityEnabled = [int]($Group.securityEnabled -eq $true)
        IsAssignableToRole = [int]($Group.isAssignToRole -eq $true)
        IsHiddenInOutlookClients = [int]($Group.hideFromOutlookClients -eq $true)
        IsMembershipDynamic = [int]($Group.membershipRule -ne $null)
        MembershipRule = $Group.membershipRule
        Notes = $Group.notes
        ManagedBy = ($Group.managedBy -join ';')
    }

    Invoke-SqlCommand -Sql $sql -Parameters $parameters
}

function Get-GroupsFromGraph {
    param([object]$GraphContext)

    Write-DebugLog 'Starting Graph group query...'

    # Microsoft Graph's group resource has no "RecipientTypeDetails", "Notes" or "ManagedBy"
    # property (those are Exchange-only concepts) - RecipientTypeDetails simply doesn't exist,
    # the notes-equivalent field is "Description", and owners must come from the Owners
    # navigation property. -ExpandProperty Owners resolves them inline in the same call
    # instead of one Get-MgGroupOwner round-trip per group.
    Write-DebugLog 'Using Microsoft Graph PowerShell cmdlets for groups (expanding owners)...'
    $groups = Get-MgGroup -All -ExpandProperty Owners -ErrorAction Stop
    if ($null -ne $groups) {
        return @($groups | ForEach-Object {
            $managedByList = @()
            foreach ($owner in @($_.Owners)) {
                if ($null -eq $owner) { continue }

                $ownerAdditional = $owner.AdditionalProperties
                $ownerDisplay = $ownerAdditional['displayName']
                $ownerUpn = $ownerAdditional['userPrincipalName']
                $ownerMail = $ownerAdditional['mail']

                if (-not [string]::IsNullOrWhiteSpace($ownerDisplay)) {
                    $managedByList += $ownerDisplay
                }
                elseif (-not [string]::IsNullOrWhiteSpace($ownerUpn)) {
                    $managedByList += $ownerUpn
                }
                elseif (-not [string]::IsNullOrWhiteSpace($ownerMail)) {
                    $managedByList += $ownerMail
                }
                else {
                    $managedByList += $owner.Id
                }
            }

            [pscustomobject]@{
                id = $_.Id
                displayName = $_.DisplayName
                mailNickname = $_.MailNickname
                mail = $_.Mail
                groupTypes = @($_.GroupTypes)
                recipientTypeDetails = $null
                mailEnabled = $_.MailEnabled
                securityEnabled = $_.SecurityEnabled
                hideFromOutlookClients = $_.HideFromOutlookClients
                isAssignToRole = $_.IsAssignableToRole
                isMembershipDynamic = $_.IsMembershipDynamic
                membershipRule = $_.MembershipRule
                notes = $_.Description
                createdDateTime = $_.CreatedDateTime
                managedBy = $managedByList
            }
        })
    }

    throw 'No groups were returned by Microsoft Graph.'
}

function Limit-Text {
    # SqlBulkCopy refuses a value longer than its column instead of truncating it like an INSERT
    # with ANSI_WARNINGS OFF would - one oversized display name would fail the whole group.
    param([string]$Value, [int]$Max)
    if ([string]::IsNullOrEmpty($Value)) { return $Value }
    if ($Value.Length -le $Max) { return $Value }
    return $Value.Substring(0, $Max)
}

function Write-GroupMembers {
    # Replaces the group's cached member rows in one transaction: delete, then one bulk insert.
    # The old per-row MERGE never deleted anything, so members removed in M365 stayed in the cache
    # for good; and it cost two round trips per member.
    param(
        [guid]$GroupId,
        [string]$GroupDisplayName,
        [object[]]$Members
    )

    $table = New-Object System.Data.DataTable
    foreach ($column in 'GroupId', 'MemberType', 'MemberId', 'DisplayName', 'PrimarySmtpAddress', 'UserPrincipalName', 'UserType', 'IsNested', 'Path') {
        $type = switch ($column) { 'GroupId' { [guid] } 'IsNested' { [bool] } default { [string] } }
        [void]$table.Columns.Add($column, $type)
    }

    foreach ($member in @($Members)) {
        if ($null -eq $member) { continue }

        $odataType = [string]$member['@odata.type']
        $isGroup = $odataType -eq '#microsoft.graph.group'
        $mail = [string]$member['mail']
        $upn = [string]$member['userPrincipalName']
        $displayName = [string]$member['displayName']
        $primary = if (-not [string]::IsNullOrWhiteSpace($mail)) { $mail } elseif (-not [string]::IsNullOrWhiteSpace($upn)) { $upn } else { $null }

        $row = $table.NewRow()
        $row['GroupId'] = $GroupId
        $row['MemberType'] = if ($isGroup) { 'Group' } else { 'User' }
        $row['MemberId'] = Limit-Text ([string]$member['id']) 256
        $row['DisplayName'] = Limit-Text $displayName 256
        $row['PrimarySmtpAddress'] = if ($null -eq $primary) { [DBNull]::Value } else { Limit-Text $primary 256 }
        $row['UserPrincipalName'] = if ($isGroup -or [string]::IsNullOrWhiteSpace($upn)) { [DBNull]::Value } else { Limit-Text $upn 256 }
        $row['UserType'] = if ($isGroup -or [string]::IsNullOrWhiteSpace([string]$member['userType'])) { [DBNull]::Value } else { Limit-Text ([string]$member['userType']) 64 }
        # Same meaning as before: "this member is itself a group". The app reads IsNested = 0 as the
        # direct (user) members, e.g. for the member count on the Teams page.
        $row['IsNested'] = $isGroup
        $row['Path'] = if ($isGroup) { Limit-Text "$GroupDisplayName > $displayName" 2048 } else { '' }
        $table.Rows.Add($row)
    }

    $connection = Get-SqlConnection
    $transaction = $connection.BeginTransaction()
    try {
        $delete = $connection.CreateCommand()
        $delete.Transaction = $transaction
        $delete.CommandText = 'DELETE FROM dbo.GroupMembers WHERE GroupId = @GroupId'
        [void]$delete.Parameters.Add('@GroupId', [System.Data.SqlDbType]::UniqueIdentifier)
        $delete.Parameters['@GroupId'].Value = $GroupId
        [void]$delete.ExecuteNonQuery()
        $delete.Dispose()

        if ($table.Rows.Count -gt 0) {
            $bulk = New-Object System.Data.SqlClient.SqlBulkCopy($connection, [System.Data.SqlClient.SqlBulkCopyOptions]::Default, $transaction)
            try {
                $bulk.DestinationTableName = 'dbo.GroupMembers'
                $bulk.BulkCopyTimeout = 300
                foreach ($column in $table.Columns) {
                    [void]$bulk.ColumnMappings.Add($column.ColumnName, $column.ColumnName)
                }
                # LastImportedAtUtc is not mapped - its column default (SYSUTCDATETIME) fills it.
                $bulk.WriteToServer($table)
            }
            finally {
                $bulk.Close()
            }
        }

        $transaction.Commit()
    }
    catch {
        try { $transaction.Rollback() } catch { }
        throw
    }
    finally {
        $transaction.Dispose()
    }
}

function Set-GroupDeleted {
    # A group deleted after the group list was read. Write-GroupRecord already stamped it as seen,
    # so the end-of-run marking would miss it - flag it here.
    param([guid]$GroupId)
    Invoke-SqlCommand -Sql 'UPDATE dbo.Groups SET IsDeletedInM365 = 1, UpdatedDateTime = SYSUTCDATETIME() WHERE Id = @Id' -Parameters @{ Id = $GroupId } | Out-Null
}

function Test-IsNotFound {
    # Graph's answer for an object deleted after the group list was read - normal in a run over
    # ~37,000 groups, and no reason to throw the run away.
    param($ErrorRecord)
    $text = "$($ErrorRecord.Exception.Message) $($ErrorRecord.ErrorDetails.Message)"
    return $text -match 'Request_ResourceNotFound|ResourceNotFound|does not exist|\(404\)|NotFound'
}

function Get-MembersBatch {
    # Direct members of up to 20 groups per Graph $batch request - ~1,900 requests for ~37,000
    # groups instead of one each. Each group gets back one of:
    #   @{ Status = 'ok';    Members = <member hashtables> }
    #   @{ Status = 'gone';  Message = ... }   (deleted after the group list was read)
    #   @{ Status = 'error'; Message = ... }
    # Throttled (429) and transient (5xx) answers are retried for that group only, honouring
    # Retry-After. Groups with more than 999 members continue through @odata.nextLink.
    param([object[]]$Groups)

    $results = @{}
    $pending = @($Groups)
    $attempt = 0
    $maxAttempts = [Math]::Max($MaxRetries, 1) + 2

    while ($pending.Count -gt 0 -and $attempt -lt $maxAttempts) {
        $attempt++
        $requests = @()
        $index = 0
        foreach ($group in $pending) {
            $index++
            $requests += @{
                id     = [string]$index
                method = 'GET'
                url    = "/groups/$($group.id)/members?" + '$select=id,displayName,mail,userPrincipalName,userType&$top=999'
            }
        }

        try {
            $body = @{ requests = $requests } | ConvertTo-Json -Depth 5
            $response = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' -Body $body -ContentType 'application/json' -ErrorAction Stop
        }
        catch {
            # The batch call itself failed (network, throttled as a whole): retry all of it.
            Write-ImportLog "Graph batch request failed (attempt $attempt of $maxAttempts): $($_.Exception.Message)"
            Start-Sleep -Seconds ([Math]::Min(60, 10 * $attempt))
            continue
        }

        $retry = @()
        $waitSeconds = 0
        foreach ($answer in @($response['responses'])) {
            $group = $pending[[int]$answer['id'] - 1]
            $status = [int]$answer['status']
            $answerBody = $answer['body']

            if ($status -eq 200) {
                $members = New-Object System.Collections.Generic.List[object]
                foreach ($item in @($answerBody['value'])) { if ($null -ne $item) { $members.Add($item) } }
                $next = $answerBody['@odata.nextLink']
                try {
                    while ($next) {
                        $page = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
                        foreach ($item in @($page['value'])) { if ($null -ne $item) { $members.Add($item) } }
                        $next = $page['@odata.nextLink']
                    }
                    $results[[string]$group.id] = @{ Status = 'ok'; Members = $members.ToArray() }
                }
                catch {
                    $results[[string]$group.id] = if (Test-IsNotFound $_) {
                        @{ Status = 'gone'; Message = $_.Exception.Message }
                    } else {
                        @{ Status = 'error'; Message = "reading further member pages failed: $($_.Exception.Message)" }
                    }
                }
            }
            elseif ($status -eq 404) {
                $results[[string]$group.id] = @{ Status = 'gone'; Message = [string]$answerBody['error']['message'] }
            }
            elseif ($status -eq 429 -or $status -ge 500) {
                $retry += $group
                $after = 0
                if ($answer['headers'] -and [int]::TryParse([string]$answer['headers']['Retry-After'], [ref]$after)) {
                    $waitSeconds = [Math]::Max($waitSeconds, $after)
                }
            }
            else {
                $code = [string]$answerBody['error']['code']
                $message = [string]$answerBody['error']['message']
                $results[[string]$group.id] = @{ Status = 'error'; Message = "($status) $code $message".Trim() }
            }
        }

        $pending = $retry
        if ($pending.Count -gt 0) {
            $waitSeconds = [Math]::Max($waitSeconds, 5 * $attempt)
            Write-DebugLog "Graph throttled/busy for $($pending.Count) group(s) - retrying in $waitSeconds s."
            Start-Sleep -Seconds ([Math]::Min(120, $waitSeconds))
        }
    }

    foreach ($group in $pending) {
        $results[[string]$group.id] = @{ Status = 'error'; Message = "Graph kept refusing the member query (throttled or unavailable) after $maxAttempts attempts." }
    }

    return $results
}

try {
    if (-not [string]::IsNullOrWhiteSpace($LogFile)) {
        Rotate-ImportLogs
    }

    Write-ImportLog 'Ensuring SQL schema...'
    Ensure-Schema

    Write-ImportLog 'Connecting to Microsoft Graph...'
    $graphContext = $null
    $attempt = 0
    while ($null -eq $graphContext -and $attempt -lt $MaxRetries) {
        $attempt++
        try {
            $graphContext = Get-GraphAccessToken -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        catch {
            Write-ImportLog "Graph connection attempt $attempt failed: $($_.Exception.Message)"
            if ($attempt -ge $MaxRetries) { throw }
            Start-Sleep -Seconds 10
        }
    }

    # Every group written below gets LastSeenAtUtc = now; whatever is older than this afterwards
    # was not in Graph this time round.
    $runStartUtc = Get-SqlUtcNow

    Write-ImportLog 'Querying groups from Microsoft Graph...'
    $groups = Get-GroupsFromGraph -GraphContext $graphContext

    Write-ImportLog "Importing $($groups.Count) groups..."

    # One group failing used to end the whole run - after hours, with nothing marked deleted. Now a
    # group deleted meanwhile is skipped and flagged, other per-group errors are logged and counted.
    # Many errors in one run point at something systemic (expired certificate, Graph outage), so the
    # run still stops at this limit instead of logging the same failure 37,000 times.
    $maxGroupErrors = 50
    $groupErrors = 0
    $groupsGone = 0
    $memberRows = 0
    $batchSize = 20          # Graph's $batch limit
    $progressEvery = 2000
    $watch = [System.Diagnostics.Stopwatch]::StartNew()

    for ($offset = 0; $offset -lt $groups.Count; $offset += $batchSize) {
        $chunk = @($groups[$offset..([Math]::Min($offset + $batchSize, $groups.Count) - 1)])

        # The group row first - the member rows reference it. A group whose row cannot be written
        # counts as a failed group; its members are not fetched.
        $recorded = @()
        foreach ($group in $chunk) {
            try {
                Write-GroupRecord -Group $group -GroupKey ([guid]$group.id).ToString() -GroupTypeSummary ($group.groupTypes -join ',')
                $recorded += $group
            }
            catch {
                $groupErrors++
                Write-ImportLog "ERROR in group '$($group.displayName)' [$($group.id)]: $($_.Exception.Message)"
                if ($groupErrors -ge $maxGroupErrors) {
                    throw "Stopped after $groupErrors failed groups - see the errors above."
                }
            }
        }

        $answers = if ($recorded.Count -gt 0) { Get-MembersBatch -Groups $recorded } else { @{} }

        foreach ($group in $recorded) {
            $groupId = [guid]$group.id
            $answer = $answers[[string]$group.id]
            try {
                switch ($answer.Status) {
                    'ok' {
                        Write-GroupMembers -GroupId $groupId -GroupDisplayName $group.displayName -Members $answer.Members
                        $memberRows += @($answer.Members).Count
                    }
                    'gone' {
                        $groupsGone++
                        Write-ImportLog "Group '$($group.displayName)' [$groupId] was deleted during the run - marked as deleted."
                        Set-GroupDeleted -GroupId $groupId
                    }
                    default {
                        throw $(if ($answer) { $answer.Message } else { 'no answer from Graph for this group' })
                    }
                }
            }
            catch {
                $groupErrors++
                Write-ImportLog "ERROR in group '$($group.displayName)' [$groupId]: $($_.Exception.Message)"
                if ($groupErrors -ge $maxGroupErrors) {
                    throw "Stopped after $groupErrors failed groups - see the errors above."
                }
            }
        }

        $done = [Math]::Min($offset + $batchSize, $groups.Count)
        if ($done % $progressEvery -lt $batchSize -or $done -eq $groups.Count) {
            Write-ImportLog ("{0} of {1} groups done, {2} member rows, {3:hh\:mm\:ss} elapsed." -f $done, $groups.Count, $memberRows, $watch.Elapsed)
        }
    }

    if ($groupsGone -gt 0) {
        Write-ImportLog "$groupsGone group(s) were deleted in M365 while the run was going."
    }

    Write-ImportLog 'Marking groups not seen in this run as deleted in M365...'
    # Used to be "WHERE CONVERT(NVARCHAR(36), Id) NOT IN (<every id of this run>)": a literal list of
    # ~37,000 GUIDs, compared against a converted column no index could serve. It outgrew the 30 s
    # command timeout and failed every run, after nine hours of importing. The LastSeenAtUtc stamp
    # every group got above says the same thing and is one indexed-range condition. Groups already
    # flagged are left alone instead of being re-flagged on every run.
    # Guarded like before: an empty Graph result must never mark the whole table as deleted.
    if ($groups.Count -gt 0) {
        $markDeletedSql = @"
UPDATE dbo.Groups
SET IsDeletedInM365 = 1,
    UpdatedDateTime = SYSUTCDATETIME()
WHERE IsDeletedInM365 = 0
  AND (LastSeenAtUtc IS NULL OR LastSeenAtUtc < @RunStartUtc)
"@
        $marked = Invoke-SqlCommand -Sql $markDeletedSql -Parameters @{ RunStartUtc = $runStartUtc } -TimeoutSeconds 300 -ReturnRowCount
        Write-ImportLog "$marked group(s) newly marked as deleted in M365."
    }

    if ($groupErrors -gt 0) {
        # Finished and saved, but not clean - exit 1 so the task history shows it.
        Write-ImportLog "Import completed with $groupErrors failed group(s) - see the errors above."
        exit 1
    }

    Write-ImportLog 'Import completed successfully.'
    exit 0
}
catch {
    Write-ImportLog "Import failed: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($null -ne $script:SqlConnection) { $script:SqlConnection.Dispose() }
}
