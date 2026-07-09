#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Sets folder permission levels
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 01/21/2012 09:15:00 PM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '             : 
#    '
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "GrantFullAccess"
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

# Grant Full Access to Mailbox
if (Test-Path $InputFile) {					
	$MBXs = Import-CSV $InputFile
  	ForEach ($MBX in $MBXs) {
		if ($MBX.Mail.contains("@")) {
   			write-host "Granting FullAccess permissions on: " $MBX.Mail " <-- " $MBX.DelegateToAdd " 
			$LineToWrite = "INFO" + "`t" + $MBX.Mail + "`t" + $MBX.DelegateToAdd
			WriteReportEvent
			Set-MailboxPermission $MBX.Mail -User $MBX.DelegateToAdd -AccessRights FullAccess
			}
		}
}	
	# Rename the input file for future reference
	Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")

# End of script & Remove PS Session
Remove-PSSession $Session
$Session = $null

$LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
WriteLogEvent

# =============================================================================================================================================