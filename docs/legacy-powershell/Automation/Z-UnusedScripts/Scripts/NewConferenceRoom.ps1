#################################################################################
# 
# PowerShell source code
# Revision v1.1
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Create new Conference Rooms in Exchange Online
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 08/01/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 
#    '             : 11/30/2011 SAG Added Quota Setting        
#    '             : 03/07/2012 SAG Remove the Set-CalendarProces -AuomateProcessing
#    '             :     this is already in the ConfigConferenceRoom script
#    '             : 07/25/2013 SAG Remove granting the ACL.UL.FullMailboxRight to
#    '             :     the room
#    '                
#    '
# ==========================================================================
#
#################################################################################

$WhoAmI			= WhoAmI

$LogDirectory		= "E:\Automation\NewConferenceRoom\Log"
$LogFile		= $LogDirectory + "\" + "Log-NewConferenceRoom.log"

$InputDirectory		= "E:\Automation\NewConferenceRoom\Input"
$InputFile		= $InputDirectory + "\" + "Input-NewConferenceRoom.csv"

$ReportDirectory	= "E:\Automation\NewConferenceRoom\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-NewConferenceRoom"


	# Retrieve the local server name
		$Machine = get-wmiobject "Win32_ComputerSystem"
		$LocalMachineName = $Machine.Name


	# Retrieve the current Date and Time for use in log files
		$uDate = get-date -uformat %D
		$uTime = get-date -uformat %T
				
		$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"
		

	# Convert the current Date and Time to format MM-DD-YY and HHMMSS for use in file names	
		$Date  = $uDate.Replace("/", "-")
		$Time  = $uTime.Replace(":", "")
		
		$ReportFile = $ReportFile + "-" + "Date" + $Date + "Time" + $Time + ".Log"	

	
	# Create logging folder $LogDirectory if it's not present
		if (Test-Path $LogDirectory)
		{
			# the directory is present
		}
		
		else
		
		{
			mkdir $LogDirectory
		}
		
		
	# Add start record to log file
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "NewConferenceRoom script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		

# =============================================================================================================================================

# Connect to Office 365
#	$LiveCred = Get-Credential
#	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
#	Import-PSSession $Session

if (Test-Path $InputFile)
  {

	$RoomMailboxes = Import-CSV $InputFile

	ForEach ($Room in $RoomMailboxes)
	  {
		$atMAIL   = $Room.MAIL.indexOf("@")
		$LeftName = $Room.MAIL.substring(0,$atMAIL)

		New-Mailbox -Name $Room.Name -Room -DisplayName $Room.Name -Alias $LeftName -PrimarySmtpAddress $Room.Mail -Office $Room.Location -ResourceCapacity $Room.Capacity
		Set-Mailbox $LeftName -IssueWarningQuota 0.5GB -ProhibitSendQuota 0.75GB -ProhibitSendReceiveQuota 1.0GB
		Set-Mailbox $LeftName -CustomAttribute15 "NewConferenceRoom PS Date: $Date PS Time: $Time"
		
#		Add-MailboxPermission $LeftName -User "ACL.UL.FullMailboxRight" -AccessRights FullAccess

		$LineToWrite = $RecordEvent + "INFO" + "`t" + $Room.Name + "`t" + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		
	  }

	# Rename the input file for future reference & Remove PS Session
		Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
#		Remove-PSSession $Session

  }

else

  {
	write-host "Could not find input file :^) " $InputFile
	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Could not find input file " + "`n"
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
#	Remove-PSSession $Session
  }

# =============================================================================================================================================


# Retrieve the current Date and Time for use in log files
	$uDate = get-date -uformat %D
	$uTime = get-date -uformat %T

	$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"

# Write end record to the log file
	$LineToWrite = $RecordEvent + "INFO" + "`t" + "This instance is stopping." + "`n"
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

# End of script #
