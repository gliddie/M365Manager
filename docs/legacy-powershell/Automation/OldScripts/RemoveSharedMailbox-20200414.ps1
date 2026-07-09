# ==========================================================================================================================
#
# Connect to Office 365
#
#
# Called by:  SharedMailboxAdminMenu.ps1
#
#  04/14/2020 - SAG - Added line to show mailbox tips/owners before proceeding with the deletion

# ==========================================================================================================================
invoke-expression -Command .\ConnectO365.ps1
# ==========================================================================================================================

write-host "Name of Shared Mailbox: " -ForegroundColor Cyan -NoNewline
$InputSharedMBX = Read-Host

$ShrExists = [bool](Get-Mailbox $InputSharedMBX -ErrorAction SilentlyContinue)
If ($ShrExists -eq "True")
{
    write-host "Mailbox Owner Details: " (get-mailbox $InputSharedMBX).MailTip
    Write-Host "Confirm Removal of the " $InputSharedMBX "Shared Mailbox (Y/N): " -ForegroundColor Red -NoNewline
    $Cont = Read-Host
    If ($Cont = "Y")
    {
        write-host "Shared Mailbox Owner: " get-mailbox $InputSharedMBX 
        Write-host "Reason for Removal: " -ForegroundColor Cyan -NoNewline
        $Reason = Read-Host
        $Reason = "Reason for Removal: " + $Reason
        $OutFileName = "c:\temp\" + $inputSharedMBX + ".txt"
        $MorGrps = "Y"

        start-Transcript

        Write-Output "Removing Shared Mailbox" > $OutFileName
        Write-Output $Reason >> $OutFileName

        Get-Mailbox $InputSharedMBX >> $OutFileName

        Get-Mailbox $InputSharedMBX |fl >> $OutFileName

        Write-Output "Mailbox Permissions" >> $OutFileName

        Get-MailboxPermission $InputSharedMBX | where {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous") -and ($_.IsInherited -notlike "True")} |ft User,AccessRights
        Get-MailboxFolderPermission $InputSharedMBX | where {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous")} |ft User,AccessRights
        Get-MailboxPermission $InputSharedMBX | where {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous") -and ($_.IsInherited -notlike "True")} | ft User,AccessRights >> $OutFileName

        Write-Output "MailboxFolder Permissions" >> $OutFileName

        Get-MailboxFolderPermission $InputSharedMBX >> $OutFileName

        Write-Output "Recipient Permissions" >> $OutFileName

        Get-RecipientPermission $InputSharedMBX >> $OutFileName

        Write-Output "Mailbox Message Statistics - Number of Messages in Mailbox" >> $OutFileName

        Get-MailboxStatistics $InputSharedMBX |ft >> $OutFileName

        Write-Output "Mailbox Message Folder Statistics - Number of Messages in Mailbox" >> $OutFileName

        Get-MailboxFolderStatistics $InputSharedMBX |ft Name,ItemsInFolder >> $OutFileName

        :RemGrps Do
        {
            write-host "Name of Mailbox Security Group: " -ForegroundColor Cyan -NoNewline
            $InputDL = Read-Host

            If (($InputDL -ne "None") -and ($InputDL -ne "No") -and ($InputDL -ne ""))
            {

	            Write-Output "Security Access Group Details" >> $OutFileName
	
	            Get-DistributionGroup $InputDL
                Get-DistributionGroup $InputDL >> $OutFileName

	            Write-Output "Security Access Group Full Details" >> $OutFileName

                Get-DistributionGroup $InputDL |fl >> $OutFileName

	            Write-Output "Security Access Group Manager and Notes" >> $OutFileName

                Get-Group $InputDL |fl ManagedBy,Notes >> $OutFileName

	            Write-Output "Security Access Group Membership" >> $OutFileName

                Get-DistributionGroupMember $InputDL |ft Alias,Name,RecipientType >> $OutFileName

                Remove-DistributionGroup $InputDL -BypassSecurityGroupManagerCheck >> $OutFileName

	            Write-Host ""
                write-host "Are there additional groups created for access to this mailbox? (Y/N) " -foregroundColor Cyan -NoNewline
                $MorGrps = Read-Host
            }
            else
            {
                $MorGrps = ""
            }
        } While ($MorGrps -eq "Y" -or $MorGrps -eq "y")

        Remove-Mailbox $InputSharedMBX >> $OutFileName

        Stop-transcript
    }
}
else
{
    Write-Host ""
    Write-Host "Shared Mailbox Identified does not exist" -ForegroundColor Red
}