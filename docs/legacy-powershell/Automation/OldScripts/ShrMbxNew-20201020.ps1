<#################################################################################
# 
# PowerShell source code
# Revision v1.01
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Create New Shared Mailbox in Office 365
#    'Called By    : SharedMailboxAdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 05/27/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '               06/03/2011 MDS Added Retention Policy and User Role Assignment 
#    '               11/30/2011 SAG Remove SMTP Forwarding and Added Quota Settings
#    '               06/11/2012 SAG Added -RetentionHoldEnabled $true
#    '               03/28/2013 SAG Commented out adding full permission for 
#    '                  ACL.UL.FullMailboxRight
#    '               04/22/2013 SAG Modified quota setting to 25GB per change in 
#    '                  O365 Service Description
#    '               10/17/2013 SAG Extracted Retention Policy Setting and created
#    '                  separate Script as Wave15 upgrade errors when setting within
#    '                  this script
#    '               05/04/2018 SAG Added code to set the MessageCopyForSendOnBehalf and 
#    '                  MessageCopyForSendOnBehalf to enbled.  This ensures messages sent
#    '                  from the mailbox are stored the the Sent folder of the shared mailbox
#    '               04/29/2019 SAG Modified the code so that mailboxes can be created with the @ul.org address
#    '               03/23/2020 SAG Changed the retention policy to an active 3 yr policy.
#    '               05/25/2020 SAG Modified code to use an input form
#    '               06/02/2020 SAG Changed Get-Mailbox to Get-EXOMailbox
#    '               06/05/2020 SAG Changed back from Get-EXOMailbox to Get-Mailbox
#    '               06/09/2020 SAG Added loop to make sure that the mailbox/folders have been created
# ==========================================================================
#
#################################################################################>
 
$WhoAmI			= WhoAmI
$RoutingDomain	= "@global.ul.com"
$LogDirectory	= "E:\Automation\NewSharedMailbox\Log"
$LogFile		= $LogDirectory + "\" + "Log-NewSharedMailbox.log"
$ReportDirectory= "E:\Automation\NewSharedMailbox\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-NewSharedMailbox-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
# =============================================================================================================================================
# Begin Shared mailbox Creation
# Add start record to log file
    $LineToWrite = "STAR" + "`t" + "NewSharedMailbox script has started"
	WriteLogEvent
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	WriteLogEvent
                                
invoke-expression -Command .\ConnectO365.ps1

#Input Form Details

function Write-InputDetails
{
    $LineToWrite = "INPUT" + "`t" + "Shared Mailbox Form Input Details"
    WriteReportEvent 
    $LineToWrite = "NAME" + "`t" + $Global:txtDispName.Text
    WriteReportEvent       
    $LineToWrite = "EADDR" + "`t" + $Global:txtMbxAddr.Text
    WriteReportEvent
    $LineToWrite = "OADDR" + "`t" + $Global:txtLegAddr.Text
    WriteReportEvent
    $LineToWrite = "OWNER" + "`t" + $Global:txtGrpOwnr.Text.Trim()
    WriteReportEvent
    $LineToWrite = "TASK" + "`t" + $Global:txtInpTaskNo.Text
    WriteReportEvent

    If ($Global:txtEDGrpName.Text -ne $null)
    {
        $LineToWrite = "EDGR" + "`t" + $Global:txtEDGrpName.Text
        WriteReportEvent
        $LineToWrite = "EDGRAD" + "`t" + $Global:txtEDGrpAddr.Text
        WriteReportEvent
        $LineToWrite = "EDGRME" + "`t" + $Global:txtEDGrpMbr.Text -replace(" ","")
        WriteReportEvent
    }

    If ($Global:txtAUGrpName.Text -ne $null)
    {
        $LineToWrite = "AUGR" + "`t" + $Global:txtAUGrpName.Text
        WriteReportEvent
        $LineToWrite = "AUGRAD" + "`t" + $Global:txtAUGrpAddr.Text
        WriteReportEvent
        $LineToWrite = "AUGRME" + "`t" + $Global:txtAUGrpMbr.Text -replace(" ","")
        WriteReportEvent
    }

    If ($Global:txtREGrpName.Text -ne $null)
    {
        $LineToWrite = "REGR" + "`t" + $Global:txtREGrpName.Text
        WriteReportEvent
        $LineToWrite = "REGRAD" + "`t" + $Global:txtREGrpAddr.Text
        WriteReportEvent
        $LineToWrite = "REGRME" + "`t" + $Global:txtREGrpMbr.Text -replace(" ","")
        WriteReportEvent
    }
}

# Apply Mailbox Folder Permissions
function Add-FolderPermissions
{
    if ($GrpAddr.contains("@"))
	{
		if ((Get-Mailbox $Global:txtMbxAddr.Text).MailTip -eq $null)
        {
            $NewTip = ("Owners: " + (((Get-DistributionGroup $GrpAddr).Managedby) -join ", "))
            if ($NewTip.Length -gt 175)
            {
                write-host "Maximum mail tip length exceeded truncating to 175 characters." -foregroundcolor Red
                $NewTip = $NewTip.Substring(0,175)
                $OldTip = (Get-Mailbox $Global:txtMbxAddr.Text).MailTip
                write-host "Replacing old MailTip: " $OldTip.Substring(16,$OldTip.Length-36) -foregroundcolor Cyan
                write-host "                 With: " $NewTip
		        $LineToWrite = "REPL  " + "`t" + "Replacing old MailTip with " + $NewTip
		        WriteReportEvent
            }
            Set-Mailbox $Global:txtMbxAddr.Text -MailTip $NewTip
        }

		# 3/23/2020 is the date when all new mailboxes started getting the 3 yr retention policy
        if (((Get-Mailbox $Global:txtMbxAddr.Text).WhenCreated -gt (Get-Date -Year 2020 -Month 3 -Day 23)) -and ((Get-Mailbox $Global:txtMbxAddr.Text).RetentionPolicy -notlike "*3 yr*"))
        {
            Set-MailBox $Global:txtMbxAddr.Text -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL MRM Policy - 3 yr Delete" -RetentionHoldEnabled $false
    	    Set-Mailbox $Global:txtMbxAddr.Text -UseDatabaseQuotaDefaults $False -IssueWarningQuota 45GB -ProhibitSendQuota 49.75GB -ProhibitSendReceiveQuota 50GB
            write-host "3 yr Retention Policy Applied to" $Global:txtMbxAddr.Text
		    $LineToWrite = "UPD  " + "`t" + "Retention Policy Set " + $Global:txtMbxAddr.Text
	        WriteReportEvent
		# Check that IMAP and POP3 are disabled
    	    $ProtocolCheck = (get-CASMailbox $Global:txtMbxAddr.Text) 
    		if ($ProtocolCheck.ImapEnabled -eq $true)
	    	{
		        Set-CASMailBox $Global:txtMbxAddr.Text -ImapEnabled $false
    		    $LineToWrite = "INFO" + "`t" + $Global:txtDispName.Text + "`t" + "IMAP Protocol Disabled"
	    	    WriteReportEvent
		    }
    		if ($ProtocolCheck.PopEnabled -eq $true)
	    	{
		        Set-CASMailBox $Global:txtMbxAddr.Text -PopEnabled $false
    		    $LineToWrite = "INFO" + "`t" + $Global:txtDispName.Text + "`t" + "POP3 Protocol Disabled"
	    	    WriteReportEvent
		    }
        }

        $Permission = ""
        if ($GrpAddr.contains(".ED@"))
        {
            $Permission = "Editor"
       		write-host "Setting permissions on: " $Global:txtMbxAddr.Text " <-- " $GrpAddr " (Editor)" -ForegroundColor Cyan
	    	$LineToWrite = "INFO" + "`t" + "Setting Editor Permissions for " + $Global:txtMbxAddr.Text +  " for " + $GrpAddr
		   	WriteReportEvent
            Write-Host "GrantSendOnBehalf to Delegate Group:  " $GrpAddr -ForegroundColor Cyan
            Set-Mailbox -Identity $Global:txtMbxAddr.Text -GrantSendOnBehalfTo ((Get-Mailbox -Identity $Global:txtMbxAddr.Text).GrantSendOnBehalfTo += $GrpAddr)
            Add-MailboxPermission $Global:txtMbxAddr.Text –User $GrpAddr –AccessRights FullAccess
            Add-RecipientPermission $Global:txtMbxAddr.Text -Trustee $GrpAddr –AccessRights SendAs -Confirm:$false
        }
        elseif ($GrpAddr.contains(".AU@"))
        {
            $Permission = "PublishingAuthor"
            write-host "Setting permissions on: " $Global:txtMbxAddr.Textr " <-- " $GrpAddr " (PublishingAuthor)" -ForegroundColor Cyan
	    	$LineToWrite = "INFO" + "`t" + "Setting Author Permissions for " + $Global:txtMbxAddr.Text +  " for " + $GrpAddr
		   	WriteReportEvent
            Write-Host "GrantSendOnBehalf to Delegate Group:  " $GrpAddr -ForegroundColor Cyan
            Set-Mailbox -Identity $Global:txtMbxAddr.Text -GrantSendOnBehalfTo ((Get-Mailbox -Identity $Global:txtMbxAddr.Text).GrantSendOnBehalfTo += $GrpAddr)
        }
        elseif ($GrpAddr.contains(".RE@"))
        {
            $Permission = "Reviewer"
            write-host "Setting permissions on: " $Global:txtMbxAddr.Text " <-- " $GrpAddr " (Reviewer)" -ForegroundColor Cyan
	    	$LineToWrite = "INFO" + "`t" + "Setting Read Permissions on: " + $Global:txtMbxAddr.Text + " for " + $GrpAddr
		   	WriteReportEvent
        }
        else
        {
            write-host "Group name does not follow naming stand there is no .ED, .AU or .RE found." -ForegroundColor Red
            write-host "Group Name: " $GrpAddr
            $LineToWrite = "INFO" + "`t" + "Group name does not follow naming stand there is no .ED, .AU or .RE found " + $Global:txtMbxAddr.Text
		   	WriteReportEvent
        }
        
        If ($Permission -ne "")
        {
            
            Do
            {
                Start-Sleep -Seconds 2
                $MBXFoldersExist = [bool](Get-MailboxFolderStatistics -Identity $Global:txtMbxAddr.Text -ErrorAction SilentlyContinue)
            } while ($MBXFoldersExist -ne "True")

            $MBXFolders = Get-MailboxFolderStatistics -Identity $Global:txtMbxAddr.Text | % {$_.Identity.ToString().Split("\")[1..100] -join "\"}

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
                        Add-MailboxFolderPermission -Identity ($Global:txtMbxAddr.Text + ":\") -User $GrpAddr -AccessRights $Permission
                        $LineToWrite = "ADD  " + "`t" + $Permission + " permission to: " + $Global:txtMbxAddr.Text + ":\"
                    }
					else
                    {
					    Add-MailboxFolderPermission -Identity ($Global:txtMbxAddr.Text + ":\" + $Folder) -User $GrpAddr -AccessRights $Permission
                        $LineToWrite = "ADD  " + "`t" + $Permission + " permission to: " + $Global:txtMbxAddr.Text + ":\" + $Folder
					}
                    WriteReportEvent
				}
			}
        }
        else
        {
            write-host "Skipping this group since it does not follow our naming standard"
        }
    }	
}

#
# =============================================================================================================================================
# Starting mailbox creation

invoke-expression -command .\ShrMbxNewForm.ps1

If ($Global:Result -eq "OK")
{
    # Declare Drive | Folders | and Files
	    $LogDrive		= "e:\Automation\NewSharedMailbox\Report\Report-"
	    $ReportFile		= $LogDrive + $FileName + $Global:txtDispName.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
    # =============================================================================================================================================
    # Begin DistributionGroup creation

        $LineToWrite = "STAR" + "`t" + $FileName + " script has started"
	    WriteReportEvent
	    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	    WriteReportEvent

        Write-InputDetails

	    If ([bool](Get-Mailbox $Global:txtMbxAddr.Text -ErrorAction SilentlyContinue))
        {
            
            Output = $wshell.Popup(("There is already a mailbox created with this email address " + (Get-Mailbox $Global:txtMbxAddr.Text).DisplayName + " created on " + (Get-Mailbox $Global:txtMbxAddr.Text).WhenMailboxCreated),0,"Mailbox Already Exists",0+32)
            $LineToWrite = "STOP" + "`t" + "Shared Maibox Already Exists"
            WriteReportEvent
            Invoke-Expression -Command e:\O365AdminShared\EmailTemplates\SharedMailboxAlreadyExists.oft
        }
        else
        {
		    $atMAIL   = $Global:txtMbxAddr.text.indexOf("@")
            $LeftName = $Global:txtMbxAddr.Text.substring(0,$atMAIL)
		    $addMail  = $Global:txtLegAddr.Text.split(",")

		    New-Mailbox -Name $Global:txtDispName.Text -shared -Alias $LeftName -PrimarySmtpAddress $Global:txtMbxAddr.Text
		    Set-CASMailbox $LeftName -ImapEnabled $false -PopEnabled $false
		    set-mailbox $Global:txtDispName.Text -MessageCopyForSendOnBehalfEnabled $true -MessageCopyForSentAsEnabled $true
		
            Set-Mailbox $LeftName -EmailAddresses (((Get-Mailbox $Global:txtDispName.Text).EmailAddresses)+=($LeftName + $RoutingDomain)) -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL MRM Policy - 3 yr Delete" -RetentionHoldEnabled $false -CustomAttribute15 "NewSharedMailbox PS Date: $Date PS Time: $Time"

		    ForEach ($mail in $addMail)
		    {
		        $add = $mail

		        if($add.contains("@"))
		        {
		            Set-Mailbox $LeftName -EmailAddresses (((Get-Mailbox $LeftName).EmailAddresses)+=$add)
		            $LineToWrite = "INFO" + "`t" + $Global:txtDispName.Text + "`t" + $add + "`t" + " proxyAddress added"
    			    WriteReportEvent
		        }
                else
		        {
		             #write-host "Nothing to add"
		    	    $LineToWrite = "INFO" + "`t" + $Global:txtDispName.Text + "`t" + " no legacyAddress to add"
			        WriteReportEvent
		        }
		    }

            # Creating Security Groups and Applying Permissions
            If (($Global:txtEDGrpName.Text -ne $null) -and ($Global:txtEDGrpName.Text -ne ""))
            {
                $GrpName = $Global:txtEDGrpName.Text
                $GrpAddr = $Global:txtEDGrpAddr.Text
                $GrpMem = $Global:txtEDGrpMbr.Text
                CreateUSG
                Add-FolderPermissions
            }
            If (($Global:txtAUGrpName.Text -ne $null) -and ($Global:txtAUGrpName.Text -ne ""))
            {
                $GrpName = $Global:txtAUGrpName.Text
                $GrpAddr = $Global:txtAUGrpAddr.Text
                $GrpMem = $Global:txtAUGrpMbr.Text
                CreateUSG
                Add-FolderPermissions
            }
            If (($Global:txtREGrpName.Text -ne $null) -and ($Global:txtREGrpName.Text -ne ""))
            {
                $GrpName = $Global:txtREGrpName.Text
                $GrpAddr = $Global:txtREGrpAddr.Text
                $GrpMem = $Global:txtREGrpMbr.Text
                CreateUSG
                Add-FolderPermissions
            }

        # =============================================================================================================================================

        Invoke-Expression -Command e:\O365AdminShared\EmailTemplates\SharedMailboxNew.oft

        $LineToWrite = "STOP" + "`t" + "Shared Maibox Creation Complete"
        WriteReportEvent
    }
}
else
{
    write-host "Shared Mailbox Creation Cancelled" -ForegroundColor Red
}

# End of script #