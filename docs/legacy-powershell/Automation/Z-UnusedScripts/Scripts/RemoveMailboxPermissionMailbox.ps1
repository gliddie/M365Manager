#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Adds full access permission for Shared Mailboxes
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : 
#    'Date Created : 11/20/2012 03:45:00 PM
#    '               
#    '
#    'History      :         
#    '             : 
#    '
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "RemoveMailboxPermissionMailbox"
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

function CheckLogFiles {
	# Create Files folder $LogDirectory if it's not present
	if (Test-Path $LogDirectory)
		{
		# the directory is present
		}
	else
		{ 
		mkdir $LogDirectory
		}
} #end CheckLogFiles

function WriteLogEvent {
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LocalMachineName + "`t" + $LineToWrite
	Out-File -filepath $LogFile -append -noClobber -inputObject $RecordEvent
} #end WriteLogEvent

function WriteReportEvent {
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LineToWrite
	Out-File -filepath $ReportFile -append -noClobber -inputObject $RecordEvent
} #end WriteReportEvent

# Setup Folders and Files	
	CheckLogFiles
	$LineToWrite = "STAR" + "`t" + $FileName + " script has started"
	WriteLogEvent
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	WriteLogEvent
		
# =============================================================================================================================================

# Connect to Office 365
if ($Session -eq $null) {
	write-host "Connect to Office 365."
	$LiveCred	= Get-Credential
	$Session	= New-PSSession -ConfigurationName Microsoft.Exchange `
		-ConnectionUri https://ps.outlook.com/powershell/ `
		-Credential $LiveCred `
		-Authentication Basic `
		-AllowRedirection
	if ($? -eq $true){
		Import-PSSession $Session
	}
	else {
		write-host "Could not establish session with Office 365."
		$LineToWrite = "STOP" + "`t" + "Could not establish session with Office 365." + "`n"
		WriteLogEvent
	}
}

else {
	write-host "Session with Office 365 already exists."
}
# Remove Mailbox Permission

if (Test-Path $InputFile)
{					
	$colUsers = Import-CSV $InputFile
  	Foreach($objUser in $colUsers) 
	{
                write-host "Removing Account " $objUser.AccessGroup " from mailbox " $objuser.MbxName
		$LineToWrite = "Removing Account " + $objUser.AccessGroup + " from mailbox " + $objuser.MbxName
		WriteLogEvent
		Remove-MailboxPermission $objUser.MbxName –User $objUser.AccessGroup –AccessRights FullAccess -Confirm:$FALSE
	}
	# Rename the input file for future reference
	Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")
}
# End of script & Remove PS Session
Remove-PSSession $Session
$Session = $null

$LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
WriteLogEvent

# ============================================================================================================================================