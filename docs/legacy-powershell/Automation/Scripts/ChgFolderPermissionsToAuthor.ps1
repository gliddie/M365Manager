write-host "Enter the employee of the mailbox to grant elevated permissions to: " -ForegroundColor Cyan -NoNewline
$MbxUser = Read-Host

If ($mbxUser -notlike "*@global.ul.com")
{
    $MbxUser = $MbxUser + "@global.ul.com"
}

write-host "`nPermissions at the Mailbox Level:" -ForegroundColor Cyan
Get-MailboxPermission $MbxUser |Where-Object {$_.User -like "*ul.com*"}

write-host "`nPermissions at the Mailbox Folder Level:" -ForegroundColor Cyan
Get-MailboxFolderPermission $MbxUser |Where-Object {$_.AccessRights -notlike "None"}

write-host "Enter the employee number of the individual to modify permissions to PublishingAuthor for (0 to quit): " -ForegroundColor Cyan -NoNewline
$AuAccess = Read-Host

if ($AuAccess -ne "0")
{
    $MbxPerm = Get-MailboxPermission $MbxUser |Where-Object {$_.User -like "*$AUAccess*"}
    If ($MbxPerm.Count -ne 0)
    {
        write-host "Removing " $MbxPerm.AccessRights " from the mailbox" -ForegroundColor Red
        Remove-MailboxPermission $MbxUser -User $AUAccess -AccessRights $MbxPerm.AccessRights -confirm:$False
 #       write-host "Adding PublishingAuthor rights to the mailbox" -ForegroundColor Green
 #       Add-MailboxPermission $MbxUser -User $AUAccess -AccessRights "PublishingAuthor"
    }

    write-host "Setting permissions on: " $MbxUser " <-- " $AUAccess " (" PublishingAuthor ")" -ForegroundColor Cyan
    write-Host "GrantSendOnBehalf to Delegate:  " $AUAccess " ( " (get-mailbox $AUAccess).DisplayName " ) " -ForegroundColor Cyan
    Set-Mailbox -Identity $MbxUser -GrantSendOnBehalfTo ((Get-Mailbox -Identity $MbxUser).GrantSendOnBehalfTo += $AUAccess) -ErrorAction Silentlycontinue

	$MBXFolders = Get-MailboxFolderStatistics -Identity $MbxUser | % {$_.Identity.ToString().Split("\")[1..100] -join "\"}
    $MbxAccess = Get-MailboxPermission $MbxUser |Where-Object {$_.User -like "*$AUAccess*"}
    $FldrAccess = Get-MailboxFolderPermission $MbxUser |Where-Object {$_.AccessRights -notlike "None"}

	ForEach ($Folder in $MBXFolders)
    {
	    If ($Folder.Equals("Recoverable Items") -or $Folder.Equals("Calendar Logging") -or $Folder.Equals("Deletions") -or $Folder.Equals("Purges") -or $Folder.Equals("Versions"))
        {
	        # Ignore folder
        }
    	else
        {
            if ($Folder.Equals("Top of Information Store"))
            {
                If ($fldrAccess.AccessRights -ne "PublishingAuthor")
                {
    	    		write-host "Folder: " $MbxUser ":\ " -NoNewline
                    write-host "<-- Removing " $FldrAccess.AccessRights -ForegroundColor Red -NoNewline
                    Remove-MailboxFolderPermission -Identity ($MbxUser + ":\") -User $AUAccess -confirm:$False
                    write-host "  <-- Adding PublishingAuthor Access" -ForegroundColor Green
                    Add-MailboxFolderPermission -Identity ($MbxUser + ":\") -User $AUAccess -AccessRights "PublishingAuthor"
                }
                else
                {
                    write-host "Folder: " $MbxUser ":\ " -NoNewline
                    write-host "<-- No Changes" -ForegroundColor Cyan
                }
		    }
    		else
            {
                If ($fldrAccess.AccessRights -ne "PublishingAuthor")
                {
                    write-host "Folder: " $MbxUser ":\" $Folder -NoNewline
                    write-host " <-- Removing " $FldrAccess.AccessRights -ForegroundColor Red -NoNewline
                    Remove-MailboxFolderPermission -Identity ($MbxUser + ":\" + $Folder) -User $AUAccess -Confirm:$False
                    write-host "  <-- Adding PublishingAuthor Access" -ForegroundColor Green
                    Add-MailboxFolderPermission -Identity ($MbxUser + ":\" + $Folder) -User $AUAccess -AccessRights "PublishingAuthor"
                }
                else
                {
                    write-host "Folder: " $MbxUser ":\" $Folder -NoNewline
                    write-host " <-- No Changes" -ForegroundColor Cyan
                }
		    }
	    }
    }
}
# =============================================================================================================================================