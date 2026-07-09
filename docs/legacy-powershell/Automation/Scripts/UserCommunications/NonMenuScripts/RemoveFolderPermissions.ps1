#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Removes folder permission levels
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Kelly Salvatori (HP)
#    'Date Created : 11/08/2011 06:30:00 PM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '             : 
#    '             : 
#    '
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "RemoveFolderPermissions"
	$LogDrive		= "E:"
	$LogPath		= "\Automation"
	$LogFolder		= "\" + $FileNAme
	$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
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

function RemoveValues($Value,$List) {
	$Array = @()
	$List | foreach {
		if ($_ -notlike "*$Value*") {
			$Array += $_
		}
	}
	return $Array
} # end RemoveValues
		
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

# Set mailbox folder permissions
if (Test-Path $InputFile) {					
	$MBXs = Import-CSV $InputFile
  	ForEach ($MBX in $MBXs) {
		if ($MBX.Mail.contains("@")) {
   			write-host "Removing permissions on: " $MBX.Mail " <-- " $MBX.DelegateToAdd " (" $MBX.Permission ")" 
			$LineToWrite = "INFO" + "`t" + $MBX.Mail + "`t" + $MBX.DelegateToAdd + "`t" + $MBX.Permission + "`n"
			WriteReportEvent
			
			# Remove the delegate from a potential list of delegates
			$MBXCurrent = Get-Mailbox -Identity $MBX.Mail 
			$NewDelegates = RemoveValues $MBX.DelegateToAdd.Substring(0,$MBX.DelegateToAdd.IndexOf("@")) $MBXCurrent.GrantSendOnBehalfTo
			Set-Mailbox -Identity $MBX.Mail -GrantSendOnBehalfTo $NewDelegates
			
			# Remove the folder permissions
			$MBXFolders = Get-MailboxFolderStatistics -Identity $MBX.Mail | % {$_.Identity.ToString().Split("\")[1..100] -join "\"}
			ForEach ($Folder in $MBXFolders) {
				if ($Folder.Equals("Recoverable Items") -or $Folder.Equals("Deletions") -or $Folder.Equals("Purges") -or $Folder.Equals("Versions")) {
				# Ignore folder
				}
				else {
					if ($Folder.Equals("Top of Information Store")) {
						Remove-MailboxFolderPermission -Identity ($MBX.Mail + ":\") -User $MBX.DelegateToAdd -Confirm:$false
					}
					else {
						Remove-MailboxFolderPermission -Identity ($MBX.Mail + ":\" + $Folder) -User $MBX.DelegateToAdd -Confirm:$false
					}
				}
			}	
		}
		else {
			Write-Host $DG.Mail " - Invalid mailbox."
			$LineToWrite = "ERROR" + "`t" + $MBX.Mail + " - Invalid mailbox."
			WriteReportEvent
		}
	}
# Rename the input file for future reference & Remove PS Session

# Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")
 }
else {
		write-host "Input file not found."
}
# End of script #
#	Remove-PSSession $Session
#	$Session = $null

#	$LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
#	WriteLogEvent

# =============================================================================================================================================