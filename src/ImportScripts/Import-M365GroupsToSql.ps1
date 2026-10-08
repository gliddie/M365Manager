<#
.SYNOPSIS
Imports Microsoft 365 groups and their members into SQL Server for fast GUI lookups.

.DESCRIPTION
This script connects to Microsoft 365 using a registered app with certificate-based authentication,
queries all supported group types (distribution, mail-enabled security, security, Microsoft 365 / Unified,
dynamic distribution, and nested memberships), and stores the data in SQL Server.

It is designed to run on a separate server as a scheduled task every 4 hours.

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
- This script intentionally imports both direct and nested group membership via recursive traversal.
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
$script:ImportedGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:ImportedMemberKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

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

    $connection = New-Object System.Data.SqlClient.SqlConnection(Get-ConnectionString)
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

    $connection.Open()
    try {
        $affected = $command.ExecuteNonQuery()
        if ($ReturnRowCount) { return $affected }
    }
    finally {
        $connection.Dispose()
    }
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

function Ensure-GroupExists {
    param(
        [guid]$GroupId,
        [string]$DisplayName = "",
        [string]$PrimarySmtpAddress = ""
    )

    $sql = @"
IF NOT EXISTS (SELECT 1 FROM dbo.Groups WHERE Id = @Id)
BEGIN
    INSERT INTO dbo.Groups (Id, DisplayName, PrimarySmtpAddress, CreatedDateTime, UpdatedDateTime, LastImportedAtUtc)
    VALUES (@Id, @DisplayName, @PrimarySmtpAddress, SYSUTCDATETIME(), SYSUTCDATETIME(), SYSUTCDATETIME());
END
"@

    $parameters = @{
        Id = $GroupId
        DisplayName = $DisplayName
        PrimarySmtpAddress = $PrimarySmtpAddress
    }

    Invoke-SqlCommand -Sql $sql -Parameters $parameters
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

function Write-MemberRecord {
    param(
        [guid]$GroupId,
        [string]$MemberType,
        [string]$MemberId,
        [string]$DisplayName,
        [string]$PrimarySmtpAddress,
        [string]$UserPrincipalName,
        [string]$UserType,
        [bool]$IsNested = $false,
        [string]$Path = ""
    )

    $sql = @"
MERGE dbo.GroupMembers AS target
USING (SELECT @GroupId AS GroupId, @MemberId AS MemberId) AS source ON target.GroupId = source.GroupId AND target.MemberId = source.MemberId
WHEN MATCHED THEN
    UPDATE SET
        MemberType = @MemberType,
        DisplayName = @DisplayName,
        PrimarySmtpAddress = @PrimarySmtpAddress,
        UserPrincipalName = @UserPrincipalName,
        UserType = @UserType,
        IsNested = @IsNested,
        Path = @Path,
        LastImportedAtUtc = SYSUTCDATETIME()
WHEN NOT MATCHED THEN
    INSERT (GroupId, MemberType, MemberId, DisplayName, PrimarySmtpAddress, UserPrincipalName, UserType, IsNested, Path, LastImportedAtUtc)
    VALUES (@GroupId, @MemberType, @MemberId, @DisplayName, @PrimarySmtpAddress, @UserPrincipalName, @UserType, @IsNested, @Path, SYSUTCDATETIME());
"@

    Ensure-GroupExists -GroupId $GroupId

    $parameters = @{
        GroupId = $GroupId
        MemberType = $MemberType
        MemberId = $MemberId
        DisplayName = $DisplayName
        PrimarySmtpAddress = $PrimarySmtpAddress
        UserPrincipalName = $UserPrincipalName
        UserType = $UserType
        IsNested = [int]$IsNested
        Path = $Path
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

function Get-GroupMembersFromGraph {
    param([string]$GroupId)

    $members = @()

    Write-DebugLog "Querying members for group $GroupId"
    Write-DebugLog 'Using Microsoft Graph PowerShell cmdlets for members...'

    $memberObjects = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop
    foreach ($member in @($memberObjects)) {
        # Get-MgGroupMember returns generic DirectoryObject instances: DisplayName/Mail/
        # UserPrincipalName/UserType are not promoted to top-level properties, they live in
        # AdditionalProperties (same reason @odata.type is read from there below).
        $additional = $member.AdditionalProperties

        $displayName = $additional['displayName']
        if ([string]::IsNullOrWhiteSpace($displayName) -and $member.PSObject.Properties.Name -contains 'DisplayName') {
            $displayName = $member.DisplayName
        }

        $mail = $additional['mail']
        if ([string]::IsNullOrWhiteSpace($mail) -and $member.PSObject.Properties.Name -contains 'Mail') {
            $mail = $member.Mail
        }

        $userPrincipalName = $additional['userPrincipalName']
        if ([string]::IsNullOrWhiteSpace($userPrincipalName) -and $member.PSObject.Properties.Name -contains 'UserPrincipalName') {
            $userPrincipalName = $member.UserPrincipalName
        }

        $userType = $additional['userType']
        if ([string]::IsNullOrWhiteSpace($userType) -and $member.PSObject.Properties.Name -contains 'UserType') {
            $userType = $member.UserType
        }

        $members += [pscustomobject]@{
            Id = $member.Id
            Type = $additional['@odata.type']
            DisplayName = $displayName
            Mail = $mail
            UserPrincipalName = $userPrincipalName
            UserType = $userType
        }
    }

    return $members
}

function Test-IsNotFound {
    # Graph's answer for an object deleted after the group list was read - normal in a run that takes
    # hours over ~37,000 groups, and no reason to throw the whole run away.
    param($ErrorRecord)
    $text = "$($ErrorRecord.Exception.Message) $($ErrorRecord.ErrorDetails.Message)"
    return $text -match 'Request_ResourceNotFound|ResourceNotFound|does not exist|\(404\)|NotFound'
}

function Resolve-NestedMembers {
    param(
        [string]$GroupId,
        [string]$GroupDisplayName,
        [string]$Path = "",
        # Tracks groups already expanded in this top-level group's nested chain, so that a cycle
        # (A contains B contains A) cannot recurse forever, and a diamond (C nested under both A
        # and B) is only walked once. Left $null on the initial call; recursive calls pass the
        # same set down so it accumulates across the whole tree.
        [System.Collections.Generic.HashSet[string]]$VisitedGroupIds = $null
    )

    if ($null -eq $VisitedGroupIds) {
        $VisitedGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }
    if (-not $VisitedGroupIds.Add($GroupId)) {
        Write-DebugLog "Skipping already-visited nested group '$GroupDisplayName' [$GroupId] (cycle or diamond membership)"
        return
    }

    try {
        $members = Get-GroupMembersFromGraph -GroupId $GroupId
    }
    catch {
        # A nested group deleted while the run was going: skip that branch only. The top-level group
        # (no $Path yet) is handled by the main loop, which marks it deleted.
        if ((Test-IsNotFound $_) -and -not [string]::IsNullOrWhiteSpace($Path)) {
            Write-ImportLog "Skipping nested group '$GroupDisplayName' [$GroupId] under '$Path' - it no longer exists."
            return
        }
        throw
    }

    foreach ($member in $members) {
        if ($member.Type -eq '#microsoft.graph.group') {
            $nestedPath = if ([string]::IsNullOrWhiteSpace($Path)) { $GroupDisplayName + ' > ' + $member.DisplayName } else { $Path + ' > ' + $member.DisplayName }
            Write-MemberRecord -GroupId ([guid]$GroupId) -MemberType 'Group' -MemberId $member.Id -DisplayName $member.DisplayName -PrimarySmtpAddress $member.Mail -UserPrincipalName '' -UserType '' -IsNested $true -Path $nestedPath
            Resolve-NestedMembers -GroupId $member.Id -GroupDisplayName $member.DisplayName -Path $nestedPath -VisitedGroupIds $VisitedGroupIds
        }
        else {
            $primaryAddress = $null
            if (-not [string]::IsNullOrWhiteSpace($member.Mail)) {
                $primaryAddress = $member.Mail
            }
            elseif (-not [string]::IsNullOrWhiteSpace($member.UserPrincipalName)) {
                $primaryAddress = $member.UserPrincipalName
            }

            Write-MemberRecord -GroupId ([guid]$GroupId) -MemberType 'User' -MemberId $member.Id -DisplayName $member.DisplayName -PrimarySmtpAddress $primaryAddress -UserPrincipalName $member.UserPrincipalName -UserType $member.UserType -IsNested $false -Path $Path
        }
    }
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

    foreach ($group in $groups) {
        $groupId = [guid]$group.id
        $groupKey = $groupId.ToString()

        try {
            Write-DebugLog "Processing group: $($group.displayName) [$groupKey]"
            if ($script:ImportedGroupIds.Add($groupKey)) {
                Write-GroupRecord -Group $group -GroupKey $groupKey -GroupTypeSummary ($group.groupTypes -join ',')
            }

            Resolve-NestedMembers -GroupId $group.id -GroupDisplayName $group.displayName
        }
        catch {
            if (Test-IsNotFound $_) {
                # Deleted after the list was read. Write-GroupRecord already stamped it as seen, so
                # the end-of-run marking would miss it - flag it here instead.
                $groupsGone++
                Write-ImportLog "Group '$($group.displayName)' [$groupKey] was deleted during the run - marked as deleted."
                Invoke-SqlCommand -Sql "UPDATE dbo.Groups SET IsDeletedInM365 = 1, UpdatedDateTime = SYSUTCDATETIME() WHERE Id = @Id" -Parameters @{ Id = $groupKey } | Out-Null
                continue
            }

            $groupErrors++
            Write-ImportLog "ERROR in group '$($group.displayName)' [$groupKey]: $($_.Exception.Message)"
            if ($groupErrors -ge $maxGroupErrors) {
                throw "Stopped after $groupErrors failed groups - see the errors above."
            }
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
