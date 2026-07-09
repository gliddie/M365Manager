#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Adds full access permission for Shared Mailboxes
#    'Called By    : NewSharedmailbox.ps1
#                  : EUMMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : 
#    'Date Created : 05/22/2012 02:00:00 PM
#    'Comments     :  
#    '
#    'History      :         
#    '             : 09/04/2103 - SAG - Added 'SendAs' Recipient Permission for editors of mailbox
#    '             : 09/28/2019 - SAG - Added code to populate the MailTip field on the mailbox
#    '
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "AddMailboxPermissionSharedMailbox"
	$LogDrive		= "E:"
	$LogPath		= "\Automation"
	$LogFolder		= "\" + $FileNAme
	$LogDirectory		= $LogDrive + $LogPath + $LogFolder + "\"
	$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	$InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
	$ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
# Retrieve the user name
	$WhoAmI			= WhoAmI
		
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name
	
# =============================================================================================================================================

function CheckLogFiles
{
	# Create Files folder $LogDirectory if it's not present
	if (Test-Path $LogDirectory)
		{
		# the directory is present
		}
	else
		{ 
		mkdir $LogDirectory
		}
}
#end CheckLogFiles

function WriteLogEvent
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LocalMachineName + "`t" + $LineToWrite
	Out-File -filepath $LogFile -append -noClobber -inputObject $RecordEvent
}
#end WriteLogEvent

function WriteReportEvent
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LineToWrite
	Out-File -filepath $ReportFile -append -noClobber -inputObject $RecordEvent
}
#end WriteReportEvent

# Setup Folders and Files	
	CheckLogFiles
	$LineToWrite = "STAR" + "`t" + $FileName + " script has started"
	WriteLogEvent
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	WriteLogEvent
		
# =============================================================================================================================================

#	
# =============================================================================================================================================
#
# Connect to Office 365
#
invoke-expression -Command .\ConnectO365.ps1

# =============================================================================================================================================
#
# Add Mailbox Permission
#
if (Test-Path $InputFile) 
{					
	$colUsers = Import-CSV $InputFile
  	Foreach($objUser in $colUsers) 
	{
        Write-Host "Applying Mailbox Permissions to "$objUser.MbxName -ForegroundColor Cyan
        Add-MailboxPermission $objUser.MbxName –User $objUser.AccessGroup –AccessRights FullAccess
        Add-RecipientPermission $objUser.MbxName -Trustee $objUser.AccessGroup –AccessRights SendAs -Confirm:$false

        $NewTip = ("Owners: "+ ($DLInfo.ManagedBy -join ", "))
        if ($NewTip.Length -gt 175)
        {
            write-host "Maximum mail tip length exceeded truncating to 175 characters." -foregroundcolor Red
            $NewTip = $NewTip.Substring(0,175)
        }
        
        if ((get-mailbox $objUser.MbxName).MailTip -ne $null)
        {
            write-host "Replacing old MailTip: " (get-mailbox $objUser.MbxName).MailTip -foregroundcolor Cyan
            write-host "                 With: " $NewTip
        }

        Set-Mailbox $objUser.MbxName -MailTip $NewTip

	}
	# Rename the input file for future reference
	Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")
}

$LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
WriteLogEvent

# ============================================================================================================================================