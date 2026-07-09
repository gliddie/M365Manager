#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Creates new Universal Security Groups
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Kelly Salvatori (HP)
#    'Date Created : 10/27/2011 12:00:00 PM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 11/08/2011 04:30:00 PM        
#    '             :    Added additional error checking
#    '             : 
#    '
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "NewRoomAdmin"
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

# Begin Room Admin List creation
if (Test-Path $InputFile) {					
	$RoomAdmins = Import-CSV $InputFile
  	ForEach ($DG in $RoomAdmins) {
		if ($DG.Mail.contains("@")) {
   			$atMail = $DG.Mail.indexOf("@")
			$DGAlias = $DG.Mail.substring(0,$atMail)
			$DGManagedByMembers = $DG.ManagedBy.Split(",")
			
			# Create the Room Admin List
			New-DistributionGroup -Name $DG.Name `
				-PrimarySmtpAddress $DG.Mail `
				-Alias $DGAlias `
				-ManagedBy $DGManagedByMembers `
				-Type Security `
				| Out-Null

			Set-DistributionGroup $DG.Name -HiddenFromAddressListsEnabled $true


			if (Get-DistributionGroup $DG.Mail) {
				write-host "Room Admin List created: " $DG.Name " (" $DG.Mail ")"
				$LineToWrite = "INFO" + "`t" + $DG.Name + "`t" + $DG.Mail + "Room Admin List Created" + "`n"
				WriteReportEvent
				
				$RoomAdmin	= Get-DistributionGroup $DG.Mail
			
				# Configure the Room Admin List
				# The 'MemberDepartRestriction' and 'MemberJoinRestriction' parameters must be "Closed" for Room Admin List so these are not set.		 		
			
				# Lines below will take an additional input to hide from address book.
				# [Boolean]$DG.Hide = $false
				# $DG.Hide = [System.Convert]::ToBoolean("True")
				# 
				# Add this as a switch to the Set-DistibutionGroup cmdlet below
				# -HiddenFromAddressListsEnabled $DG.Hide `
			
				Set-DistributionGroup $RoomAdmin.PrimarySMTPAddress `
		 			-RequireSenderAuthenticationEnabled $false `
		 			-BypassSecurityGroupManagerCheck `
					-CustomAttribute15 ($FileName + " PS Date: " + (get-date -uformat %D) + " PS Time: " + (get-date -uformat %T))
				
				# Add Room Admin List members
				write-host ("Adding Members to Room Admin List: " + $RoomAdmin.Name + " (" + $RoomAdmin.PrimarySMTPAddress + ")")
				$LineToWrite = "`t" + $RoomAdmin.Name + "`t" + $RoomAdmin.PrimarySMTPAddress + "`t" + "Adding Members to Distribution Group." + "`n"
				WriteReportEvent
				
				$addMember = $Dg.Members.split(",")
				if ($? -eq $true) {
					ForEach ($member in $addMember) {
						if ($member.contains("@")) {
							if (Get-Mailbox $member) {
								Add-DistributionGroupMember $RoomAdmin.PrimarySMTPAddress -Member $member -BypassSecurityGroupManagerCheck
								write-host "Added member:" $member
								$LineToWrite = "`t" + "PASS" + "`t" + $member + "`t" + "Added to Distribution Group" + "`t" + $RoomAdmin.PrimarySMTPAddress + "`n"
								WriteReportEvent
							}
							else {
								write-host "ERROR finding member: " $member
								$LineToWrite = "`t" + "FAIL" + "`t" + $member + "`t" + "ERROR finding recipient" + "`t" + $RoomAdmin.PrimarySMTPAddress + "`n"
								WriteReportEvent
							}
						}
						else {
							write-host "ERROR invalid member: " $member
							$LineToWrite = "`t" + "FAIL" + "`t" + $member + "`t" + "ERROR invalid recipient" + "`t" + $RoomAdmin.PrimarySMTPAddress + "`n"
							WriteReportEvent
						}
					}
				}	
				else {
					write-host "No members to add"
				}
			}	
			else {
				Write-Host "Room Admin List not created: " $DG.Name " (" $DG.Mail ")"
				$LineToWrite = "FAIL" + "`t" + $DG.Name + "`t" + $DG.Mail + "`t" + "Room Admin List not created" + "`n"
				WriteReportEvent
			}
		}
		else {
			write-host "ERROR invalid Room Admin List: " $DG.Mail
			$LineToWrite = "`t" + "FAIL" + "`t" + $DG.Mail + "`t" + "ERROR invalid Room Admin List" + "`n"
			WriteReportEvent
		}
	}
# Rename the input file for future reference & Remove PS Session
Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")
 }
else {
	write-host "Input file not found."
}

# End of script #
	Remove-PSSession $Session
	$Session = $null

	$LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
	WriteLogEvent

# =============================================================================================================================================