write-host "Enter Name of Shared Mailbox: " -ForegroundColor Cyan -NoNewline
$ShrMbx = Read-Host

write-host ""
write-host "Mailbox" $ShrMbx "Details: " -ForegroundColor Yellow

get-MailboxStatistics $ShrMbx |fl TotalItemSize
$fldr = Get-MailboxFolderStatistics $ShrMbx
write-host "Mailbox folder count: " $Fldr.count

write-host ""
write-host "Mailbox Permissions: " -ForegroundColor Yellow
get-MailboxPermission $ShrMbx | Where-Object {$_.IsInherited -eq $False}

write-host ""
write-host "Mailbox Folder Permissions: " -ForegroundColor Yellow
get-MailboxFolderPermission $ShrMbx | ft User,AccessRights

$DLGrpED = "MBX." + $ShrMbx + ".ED"
$DLGrpAU = "MBX." + $ShrMbx + ".AU"
$DLGrpRE = "MBX." + $ShrMbx + ".RE"

$EDExists = [bool](Get-Group $DLGrpED -ErrorAction SilentlyContinue)
If ($EDExists -eq $True)
{
    write-host "Editors Group Details" -ForegroundColor Yellow
    Get-Group $DLGrpED | fl Notes
    Get-DistributionGroupMember $DLGrpED |Out-Host
}

$AUExists = [bool](Get-Group $DLGrpAU -ErrorAction SilentlyContinue)
If ($AUExists -eq $True)
{
    write-host "Authors Group Details" -ForegroundColor Yellow
    If ($EDExists -eq $False)
    {
        Get-Group $DLGrpAU | fl Notes
    }
    Get-DistributionGroupMember $DLGrpAU |Out-Host
}

$REExists = [bool](Get-Group $DLGrpRE -ErrorAction SilentlyContinue)
If ($REExists -eq $True)
{
    write-host "Readers Group Details" -ForegroundColor Yellow
    If (($EDExists -eq $False) -and ($AUExists -eq $False))
    {
        Get-Group $DLGrpRE | fl Notes
    }
    Get-DistributionGroupMember $DLGrpRE |Out-Host
}

write-host ""
write-host "Shared Mailbox Rules: " -ForegroundColor Yellow
Get-InboxRule -Mailbox 96151 |ft Name,Enabled,Priority

write-host "Do you wish to send a email to the mailbox requesting a new owner (Y/N)?" -ForegroundColor Red -NoNewline
$SendMsg = Read-Host