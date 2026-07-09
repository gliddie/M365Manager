####
#### M&A Activities Menu
#  Called from O365 Admin Menu to perform Shared Mailbox Administration
#
####
Function GetUserDN($strUID)
{
#####   This function connects to Active Directory and gets the record for the user 
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = "LDAP://DC=global,DC=ul,DC=com"
    $objSearcher.Filter = [string]::format("(sAMAccountName={0})",$strUID)
    $ux = $null
    $ux = $objSearcher.FindOne()
    
    if ($ux -eq $null)
    {
        return $null
    } 
    else
    {
        return $ux.Properties.distinguishedname
    }
}

Function GetAcctInfo($ENo)
{
    $ADExists = [bool]($strDN = GetUserDN $ENo -ErrorAction SilentlyContinue)
    If ($ADExists -eq "True")
    {
        $strDN = GetUserDN $ENo
        $inf = get-mailbox $ENo -ErrorAction SilentlyContinue
        write-host "Reviewing settings for: " $inf.Name
        $strUserPath = [string]::format("LDAP://{0}", $strDN)
        $u = new-object System.DirectoryServices.DirectoryEntry($strUserPath)
		write-host "Proxy Addresses For This Account : " -ForegroundColor Green
		$u.ProxyAddresses.value.split(" ")
    }
    else
    {
        Write-Host ""
        Write-Host "Active Directory Account Does Not Exist for Emp#" $ENo -ForegroundColor Red
    }

    If ($strDN -eq $null)
    {
        $ADCmt = "*****No Active Directory Account For This User*****"
        $strUserPath = "No Active Directory Account for this User"
    }
    else
    {
        $ADCmt = $u.ExtensionAttribute14.value
    }

	If ($inf.ForwardingSMTPAddress -ne $null)
	{
	    write-host "Email Proxy Address To Be Added  : " $inf.ForwardingSMTPAddress -ForegroundColor Green
			
		$RemForw = "N"
		write-host "Continue to remove forwarding for this user (Y/N)? " -ForegroundColor Yellow -NoNewline
		$RemForw = Read-Host
			
		if (($u.ProxyAddresses.value -like $inf.ForwardingSMTPAddress) -and ($RemForw -eq "Y"))
		{
			write-host "The address " $inf.ForwardingSMtPAddress " is already present on this account"
		}
		else
		{
			$u.ProxyAddresses.value = $u.ProxyAddresses.value + $ENoMbx.ForwardingSmtpAddress
			$u.CommitChanges()
			set-mailbox $Eno -ForwardingSMTPAddress $null
		}
	}
	else
	{
		if ($strDN -ne $null)
        {
            write-host "No forwarding address configured on this account do you wish to add a legacy address to AD (Y/N)? " -ForegroundColor Red -NoNewline
		    $AddLegacy = read-host
		    If ($AddLegacy -eq "Y")
		    {
			    write-host "Enter the Legacy Address to Add to the AD Account: " -ForegroundColor Cyan -NoNewline
			    $LegacyAddr = read-host
			    $u.ProxyAddresses.value = $u.ProxyAddresses.value + ("smtp:" + $LegacyAddr)
			    $u.CommitChanges()
		    }
		    write-host ""
        }
	}
}

$MbxChg = "1"

Do
{
    write-host ""
    write-host "M&A Activites Admin Menu" -ForegroundColor Magenta
    write-host
    Write-Host "     Enter ( 1) Run Report of M&A Users and O365 Licenses"
    write-host "           ( 2) Grant SVC.CRP.BTMigration Account Access"
    write-host "           ( 3) Add Legacy Address/Remove Forwarder/Remove Migration Account Access"
    write-host "           ( 4) Remove SVC.CRP.BTMigration Account Access"
    Write-Host "           ( 5) Get Mailbox Statistics for M&A Users"
    write-host "           ( 6) Delete all Contents and Personal Folders for in a Users Mailbox (delete folders not working yet)"
    write-host ""
    write-Host "           ( 0) to Return to the O365 Admin Menu"
    write-Host "     Enter Option? " -ForegroundColor Red -NoNewline
    $MbxChg = Read-Host
    write-host ""

#    invoke-Expression -Command .\MAActivitesProcessInstructions.ps1
        		
		switch ($MbxChg)
		{
			1
                {
                    write-host "Hit Return when the Input file in the e:\Automation\MAActivites\Input directory ready" -ForegroundColor Yellow -NoNewline
                    $Cnt = Read-Host
                    invoke-expression -Command .\MALicenseInformation.ps1
                    pause
                }
            2
                {
                    write-host "Grant SVC.NBK.BTMigration Account Access to M&A Users mailboxes"
                    write-host "Hit Return when the Input file in the e:\Automation\MAActivities\Input directory ready" -ForegroundColor Yellow
                    $Cnt = Read-Host
                    $MAUsers = Import-Csv e:\Automation\MAActivities\Input\MAActivities.csv
                    
                    foreach ($MAUsers in $MAUsers)
                    {
                        $UsrExists = [bool](get-mailbox $MAUsers.UPN -ErrorAction SilentlyContinue)
                        If ($UsrExists -eq "True")
                        {
                            write-host "Granting SVC.CRP.BTMigration FullAccess to" $MAUsers.UPN  -ForegroundColor Yellow
                            Add-MailboxPermission $MAUsers.UPN -AccessRights Fullaccess -User svc.crp.btmigration@global.ul.com
                        }
                        else
                        {
                            write-host "User" $MAUsers.UPN "does not exits"
                        }
                    }
                }
            3
                {
					write-host "Enter Employee Number of individual to Make Changes for: " -ForegroundColor Yellow -NoNewline
					$ENo = Read-Host
					$ENoMbx = get-mailbox $ENo -ErrorAction SilentlyContinue
					GetAcctInfo($ENo)
					# remove SVC account access to mailbox
					If ($ENoMbx -ne $null)
                    {
                        $RemAcc = read-host "Do you want to remove the service account access to this mailbox (Y/N)? "
					    If ($RemAcc -eq "Y")
					    {
						    Remove-MailboxPermission $ENo -AccessRights FullAccess -User "SVC.CRP.BTMigration@global.ul.com" -Confirm:$false
                        }
                    }
                    else
                    {
                        write-host "No mailbox exists for Emp#" $ENo -ForegroundColor Red
                    }
                }
            4
				{
                    write-host "Remove SVC.NBK.BTMigration Account Access from migrated M&A Users mailboxes"
                    write-host "Hit Return when the Input file in the e:\Automation\MAActivities\Input directory ready" -ForegroundColor Yellow
                    $Cnt = Read-Host
                    pause

                    $MAUsers = Import-Csv e:\Automation\MAActivities\Input\MAActivities.csv
                    foreach ($MAUsers in $MAUsers)
                    {
                        $UsrExists = [bool](get-mailbox $MAUsers.UPN -ErrorAction SilentlyContinue)

                        If ($UsrExists -eq "True")
                        {
                            write-host "Removing SVC.CRP.BTMigration FullAccess from" $MAUsers.UPN  -ForegroundColor Yellow
                            remove-MailboxPermission $MAUsers.UPN -AccessRights Fullaccess -User svc.crp.btmigration@global.ul.com -Confirm:$False
                        }
                        else
                        {
                            write-host "User" $MAUsers.UPN "does not exits"
                        }
                    }
                }
            5
                {
                    write-host "Hit Return when the Input file in the e:\Automation\MAActivites\Input directory ready" -ForegroundColor Yellow -NoNewline
                    $Cnt = Read-Host
                    
                    write-host "Executing Report for M&A Licensing....." -ForegroundColor Yellow

                    $MAUsers = Import-Csv e:\Automation\MAActivities\Input\MAActivities.csv

                    $text = "UserPrincipalName,DisplayName,ItemCount"
                    Out-File -FilePath c:\temp\MAMailboxItemCount.csv -InputObject $text

                    Foreach ($MAUsers in $MAUsers)
                    {
                        $Usr = $MAUsers.UPN + "@global.ul.com"
                        $UsrExists = [bool](get-mailbox $Usr -ErrorAction SilentlyContinue)

                        If ($usrExists -eq "True")
                        {
                            $stsmbx = get-mailboxstatistics $Usr
                            $text = "{0},""{1}"",{2}" -f $Usr, $stsmbx.DisplayName, $stsmbx.ItemCount
                        }
                        else
                        {
                            Write-Host "User " $MAUsers.UPN " does not exist"
                            $text = "{0},{1},{2}" -f $Usr, "False", "User Does Not Exist"
                        }
                        Out-File -FilePath c:\temp\MAMailboxItemCount.csv -InputObject $text -Append
                    }
                    Write-Host "Report file generated at c:\temp\MALicenseInformation.csv.  If you would like to maintain the input CSV file please rename it manually." -ForegroundColor Red
                }
            6
                {
                    Write-Host "Enter Employee Number for mailbox to clear" -ForegroundColor Yellow -NoNewline
                    $Usr = Read-Host
                    
                    Get-MailboxStatistics $usr |ft DisplayName,ItemCount
                    
# http://kb.cloudiway.com/how-to-delete-a-mailbox-content-in-office-365-or-exchange/
                    Search-Mailbox -Identity $Usr -DeleteContent -Force

# Deleted folders

                    get-mailboxfolderstatistics $MailboxName | Where-Object{$_.FolderType -eq "User Created" -band $_.ItemsInFolderAndSubFolders -eq 0} | ForEach-Object{  
# Bind to the Inbox Folder  
                    "Deleting Folder " + $_.FolderPath    
                    try
                    {  
                        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId((Convertid $_.FolderId))     
                        $ewsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
                        write-host $ewsFolder       
                        if($ewsFolder.TotalCount -eq 0)
                        {  
#                            $ewsFolder.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
                            write-host $ewsFolder " - Folder Deleted"
                        }
                    }
                    catch 
					{
					#	No Actions to Perform
					}
					}
				}
  }} While ($MbxChg -ne 0)