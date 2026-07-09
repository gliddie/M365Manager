# Configure Surface Hub Devices
# 08/Changed Domain Controller

Add-PSSnapin *Exchange* -erroraction SilentlyContinue

$uDate = get-date -uformat %D
$Date  = $uDate.Replace("/", "-")
$uTime = get-date -uformat %T
$Time  = $uTime.Replace(":", "")

$DC		   = "usnbkadds001p.global.ul.com"

write-host "`nStarting Configuration for Surface Hub Device" -ForegroundColor Yellow
write-host "Name of the Service Account to Configure: " -ForegroundColor Cyan -NoNewline

$DevName = Read-Host
$DevName = $DevName.Trim()
If ($DevName -notlike "*@*")
{
    $DevName = $DevName + "@global.ul.com"
}

write-host "Enter the Email Address For this Account: " -ForegroundColor Cyan -NoNewline
$EmAddr = Read-Host
$GblEmAddr = $EmAddr.Replace("ul.com","global.ul.com")

write-host "Enabling MailUser and Assigning Licenses for " $DevName -ForegroundColor Green
$atUPN  = $DevName.indexOf("@")
Enable-MailUser $DevName -Alias $DevName.substring(0,$atUPN) -ExternalEmailaddress $EmAddr -PrimarySmtpAddress $EmAddr -DomainController:$DC
Start-Sleep -Seconds 15
Set-MailUser $DevName -EmailAddresses (((Get-MailUser $DevName -DomainController:$DC).EmailAddresses)+=($GblEmAddr)) -CustomAttribute15 "EnableMailUser PS Date: $Date PS Time: $Time" -DomainController:$DC
Set-MsolUser -UserPrincipalName $DevName -UsageLocation US -ErrorAction Silentlycontinue
Set-MsolUserLicense -UserPrincipalName $DevName -AddLicenses "ul:EXCHANGEENTERPRISE" -erroraction SilentlyContinue

$AvailLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MEETING_ROOM"}
If ($AvailLic.ActiveUnits -lt $AvailLic.ConsumedUnits)
{
    Set-MsolUserLicense -UserPrincipalName $DevName -AddLicenses "ul:MEETING_ROOM" -erroraction SilentlyContinue
}
else
{
    write-host "No available Microsoft Teams Meeting Rooms Licenses.  License must be procured and then assigned to this account" -ForegroundColor Red
}

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
Import-PSSession $session -AllowClobber
$easpolicy = Get-MobileDeviceMailboxPolicy "SurfaceHub Policy"

$LDAPFilter = "(userPrincipalName=" + $DevName + ")"
$ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties DisplayName
write-host "Is the current DisplayName of '" $ADuser.DisplayName "' correct (Y/N)? " -ForegroundColor Cyan -NoNewline
$UpdDispName = Read-Host
If ($updDispName -ne "Y")
{
    write-host "Enter the DisplayName for the account" $User.UPN -ForegroundColor Cyan -NoNewline
    $ActDisName = read-host
    Set-ADUser $ADuser -DisplayName $ActDisName
}

write-host "Enter the Employee Number of Where This Device will be installed or who will be the Room Delegates (separate multiple individuals with a comma)." -foregroundColor Cyan -NoNewline
$DevDelegate = Read-Host
$DevDelegate = $DevDelegate + "@global.ul.com"

#write-host "What region will this device be depolyed in (NBK/FRK/NWT)? " -ForegroundColor Cyan -NoNewline
#$RegDeploy = Read-Host

#Make sure Mailbox has been created

$MbxExists = [bool](get-mailbox $DevName -ErrorAction SilentlyContinue)
write-host "Waiting for Mailbox Creation to Complete" -ForegroundColor Cyan -NoNewline

do {
    write-host ".." -ForegroundColor Cyan -NoNewline
    Start-Sleep -Seconds 30
    $mbxExists = [bool](get-mailbox $DevName -ErrorAction SilentlyContinue)
} while ($MbxExists -eq $False)

write-host "`nChanging the mailbox type to a Room Mailbox"
set-mailbox $DevName -Type “Room”
write-host "Assigning the ActiveSync Policy for the SurfaceHubs"
Set-CASMailbox $DevName -ActiveSyncMailboxPolicy $easPolicy.id

write-host "Configuring the Booking Policy to the Standard UL Booking Policies"
do{
    Start-Sleep -Seconds 10
    $mbxType = get-mailbox $Devname
} while ($mbxType.ResourceType -ne "Room")

Set-CalendarProcessing $DevName -BookingWindowInDays 90 -MaximumDurationInMinutes 1440 -ConflictPercentageAllowed 20 -MaximumConflictInstances 3 -AutomateProcessing:AutoAccept
Set-CalendarProcessing $DevName -RemovePrivateProperty $true -DeleteComments $false -DeleteSubject $false -AddOrganizerToSubject $true -AllRequestOutOfPolicy $true -AllBookInPolicy $true -AllRequestInPolicy $false -ResourceDelegates $DevDelegate -AddAdditionalResponse $true -AdditionalResponse “This is a Surface Hub room!”
write-host "Granting Global RRS Admins permission to the mailbox"
Add-MailboxPermission $DevName -AccessRights Fullaccess -User "dbs.crp.rrs.admins"

write-host "Enabling the CSMeetingRoom in the region this device will be deployed"
$SIPAddr = "sip:" + $EmAddr
Enable-CsUser $DevName -SipAddress $SIPAddr -HostingProviderProxyFqdn sipfed.online.lync.com 

#switch ($RegDeploy)
#{
#    "NBK"
#    {
#        Enable-CsMeetingRoom -Identity $DevName -RegistrarPool NBKFE03.ul.com -SipAddressType EmailAddress
#    }
#    "FRK"
#    {
#        Enable-CsMeetingRoom -Identity $DevName -RegistrarPool FRKFE02.ul.com -SipAddressType EmailAddress
#    }
#    "NWT"
#    {
#        Enable-CsMeetingRoom -Identity $DevName -RegistrarPool NWTFE02.ul.com -SipAddressType EmailAddress
#    }
#}

Set-MsolUserLicense -UserPrincipalName $DevName -RemoveLicenses "ul:EXCHANGEENTERPRISE" -erroraction SilentlyContinue
write-host "`nMoving AD Object to the O365 Licensed PC Services OU"
Move-ADObject -Identity $ADUser.ObjectGUID -TargetPath "OU=PCServices,OU=O365 Licensed,OU=ServiceAccounts,OU=Enterprise,DC=global,DC=ul,DC=com"
write-host "`nSurface Hub Device set-up complete" -ForegroundColor Red