#  This section gets Bookings that are "shared booking calendars"
# Connect to Exchange Online
Connect-ExchangeOnline

# All Bookings calendars (scheduling mailboxes)
$SchedulingMailboxes = Get-EXOMailbox -RecipientTypeDetails SchedulingMailbox -ResultSize Unlimited | select DisplayName, PrimarySmtpAddress |Sort-Object User

# For each calendar, list users with full access, excluding system principals
$BookingsReport = foreach ($Mailbox in $SchedulingMailboxes) {

    $FullAccessPermissions = Get-MailboxPermission -Identity $Mailbox.PrimarySmtpAddress |
                             ? { -not $_.IsInherited -and
                                 $_.AccessRights -contains 'FullAccess' -and
                                 $_.User -notmatch 'NT AUTHORITY|S-1-5-'
                             } |  select @{n='User';e={$_.User}}, @{n='Right';e={'FullAccess'}}

    foreach ($Permission in $FullAccessPermissions) {
        [pscustomobject]@{
            BookingCalendar = $Mailbox.DisplayName
            BookingAddress  = $Mailbox.PrimarySmtpAddress
            User            = $Permission.User
            Right           = $Permission.Right
        }
    }
}
$BookingsReport | export-csv c:\temp\BookingsReport.csv -NoTypeInformation

#This section gets Bookings that are associated for a licensed users personal mailbox

# Pre-req: Connect-ExchangeOnline
[Net.ServicePointManager]::SecurityProtocol  = [Net.SecurityProtocolType]::Tls12
[Net.ServicePointManager]::Expect100Continue = $false

# Get tenant ID from your EXO session
$exo = Get-ConnectionInformation
$tenantId = [string]$exo.TenantID
$tenantid = "70115954-0ccd-45f0-87bd-03b2a3587569"

# Pull user mailboxes
$users = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties ExchangeGuid,CustomAttribute4,CustomAttribute8
function Test-PersonalBookingsProvisioned {
    param(
        [Parameter(Mandatory)][string]$ExchangeGuid32,
        [Parameter(Mandatory)][string]$TenantId
    )
    try {
        # Light-weight web session with expected cookies (cookies do not appear to be required, commenting out for now)
        $s = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        #$s.Cookies.Add((New-Object System.Net.Cookie("ClientId","ABCDEF0123456789ABCDEF0123456789","/","outlook.office.com")))
        #$s.Cookies.Add((New-Object System.Net.Cookie("OIDC","1","/","outlook.office.com")))

        # Build the expected URI & probe the user's Personal Bookings page from the Bookings Service API
        $uri = "https://outlook.office.com/BookingsService/api/V1/bookingBusinessesc2/mbx:$ExchangeGuid32@$TenantId/services"
        $r = Invoke-WebRequest -Uri $uri -UseBasicParsing -Method GET -WebSession $s -ErrorAction Stop -Headers @{
            "x-anchormailbox" = "mbx:$ExchangeGuid32@$TenantId"
            "x-req-source"    = "BookWithMe"
            "accept"          = "*/*"
        }

        # If we got here without throwing an error, the page is provisioned
        return $true
    }
    catch { return $false }
}

$results = foreach ($u in $users) {
    $id32   = $u.ExchangeGuid.ToString('N').ToLower()
    $domain = ($u.PrimarySmtpAddress -split '@')[-1]
    if (Test-PersonalBookingsProvisioned -ExchangeGuid32 $id32 -TenantId $tenantId) {
        [pscustomobject]@{
            DisplayName         = $u.DisplayName
            PrimarySmtpAddress  = $u.PrimarySmtpAddress
            AlphaOrg2 = $u.CustomAttribute4
            AlphaOrg4 = $u.CustomAttribute8
            PersonalBookingsUrl = "https://outlook.office.com/bookwithme/user/$id32@$domain?anonymous&ep=plink"
        }
    }
}

$results | Sort-Object DisplayName
# Optional export:
 $results | Export-Csv .\PersonalBookings_Enabled.csv -NoTypeInformation