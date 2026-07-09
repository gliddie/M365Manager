#################################################################################
#
# PowerShell source code
# Revision Draft v0.7
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Add Members to Distribution Groups
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 09/07/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    '             : Modified 3/1/2012 - Sandi Glazebrook
#    '             : Added -ResultSize Unlimited so that DL's with large memberships will populate without issue.
#    '             : Created new script to remove DG group members
#
#
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "RemoveDgMembers"
	$LogDrive		= "E:"
	$LogPath		= "\Automation"
	$LogFolder		= "\" + $FileName
	$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
	
	$LogFile		= $LogDirectory + "\Log\Log-" + $FileName + ".log"
	$InputFile		= $LogDirectory + "\Input\Input-" + $FileName + ".csv"
	$ReportFile		= $LogDirectory + "\Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
# Retrieve the user name and set DG Owner
	$WhoAmI			= WhoAmI
	
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name
	
# =============================================================================================================================================

function CheckLogFiles {
	# Create Files folder $LogDirectory if it's not present
	if (Test-Path $LogDirectory) {
		# the directory is present
	}
	else { 
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
if ($Session.State -ne "Opened") {
	if ($Session -ne $null) {
	Remove-PSSession $Session
	}
	
	write-host "Connect to Office 365."

       $LiveCred	= Get-Credential
       $global:Session	= New-PSSession `
	   		-ConfigurationName Microsoft.Exchange `
			-ConnectionUri https://ps.outlook.com/powershell/ `
			-Credential $LiveCred `
			-Authentication Basic `
			-AllowRedirection
		Import-PSSession $Session
}
else {
	write-host "Office 365 Connection Ready"
}

if ($Session.State -eq "Opened") {
	write-host "Import CSV File."

	if (Test-Path $InputFile) {					
		$DGroups = Import-CSV $InputFile


		if ($DGroups -ne $null) {
			write-host "Start removing members."

			ForEach ($DG in $DGroups) {
				If ($TestGroup = Get-DistributionGroup $Dg.DgMail) {
					$CurrentMembers = Get-DistributionGroupMember $Dg.DgMail -ResultSize Unlimited
					write-host "Removing Members from Distribution Group: " $DG.DgDisplayName
					$LineToWrite = $DG.DgDisplayName + " - Removing Members from Distribution Group."
					WriteReportEvent
					
					$addMember  = $Dg.MemberMail.split(",")
					if ($? -eq $true){
						ForEach ($mail in $addMember) {
							if ($mail.contains("@")) {
								if ($NewMember = Get-Recipient $mail -erroraction SilentlyContinue) {
									if ($CurrentMembers -match $NewMember.Name) 
									{
										Remove-DistributionGroupMember $Dg.DgMail -Member $mail -Confirm:$False
										write-host $mail " was removed."
										$LineToWrite = "`t" + "Success" + "`t" + $mail + " - removed from Distribution Group. " + $Dg.DgMail
										WriteReportEvent
									}
									else {
										write-host $mail " is not a member."
										$LineToWrite = "`t" + "Warn" + "`t" + $mail + " - is not a member of Distribution Group. " + $Dg.DgMail
										WriteReportEvent
									}
								}
								else {
									write-host $mail " - ERROR finding recipient."
									$LineToWrite = "`t" + "ERROR" + "`t" + $mail + " - ERROR finding recipient. " + $Dg.DgMail
									WriteReportEvent
								}
							}
						
							else {
								write-host "Nothing to add"
							}
						}
					}
				}
				else {
				Write-Host $Dg.DgMail " - ERROR finding Distribution Group."
				$LineToWrite = "ERROR" + "`t" + $Dg.DgMail + " - ERROR finding Distribution Group."
				WriteReportEvent
				}

			}
		}
	}

	else {
		write-host "Input file not found."
	}
}
else {
	write-host "Could not estahblish session with Office 365."
	$LineToWrite = "STOP" + "`t" + "Could not establish session with Office 365." + "`n"
	WriteLogEvent
}

# Rename the input file for future reference
	Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")

# End of script #
#	Remove-PSSession $Session
#	$Session = $null

	$LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
	WriteLogEvent

# =============================================================================================================================================