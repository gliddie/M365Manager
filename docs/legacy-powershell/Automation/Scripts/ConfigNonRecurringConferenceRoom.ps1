<#################################################################################
# 
# PowerShell source code
# Revision v1.5
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Configure NonRecurring Meeting Conference Rooms in Exchange Online
#    'Called By    : RoomResourceAdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 08/01/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 06/19/2018 - S.Glazebrook -  Created from the ConfigConferenceRoom script to handling NonRecurring rooms spaces
# ==========================================================================
#
#################################################################################>

$text = get-content e:\automation\scripts\RoomResourceNonRecurringPolicy.htm

$WhoAmI			= WhoAmI
$Folder			= ":\Calendar"

$LogDirectory		= "E:\Automation\ConfigRestrictedConferenceRoom\Log"
$LogFile		= $LogDirectory + "\" + "Log-ConfigRestrictedConferenceRoom.log"

$InputDirectory		= "E:\Automation\ConfigRestrictedConferenceRoom\Input"
$InputFile		= $InputDirectory + "\" + "Input-ConfigNonRecurringConferenceRoom.csv"

$ReportDirectory	= "E:\Automation\ConfigRestrictedConferenceRoom\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-ConfigNonRecurringConferenceRoom"

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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "ConfigConferenceRoom script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		

# =============================================================================================================================================

if (Test-Path $InputFile)
  {

	$RoomMailboxes = Import-CSV $InputFile

	ForEach ($Room in $RoomMailboxes)
	  {
		$atMAIL   = $Room.MAIL.indexOf("@")
		$LeftName = $Room.MAIL.substring(0,$atMAIL)
		[Boolean]$Room.AllRequestOutofPolicy = [System.Convert]::ToBoolean($Room.AllRequestOutofPolicy)

		Set-Mailbox $Room.Mail -CustomAttribute15 "ConfigConferenceRoom PS Date: $Date PS Time: $Time"

        Set-CalendarProcessing $LeftName -ResourceDelegates $Room.BookOutPolicy -ProcessExternalMeetingMessages $false -BookingWindowInDays $Room.BookingWindowInDays -AutomateProcessing $Room.AutomateProcessing -AddAdditionalResponse $true -AdditionalResponse $text -DeleteComments $false -DeleteSubject $False
        Set-CalendarProcessing $LeftName -AllowRecurringMeetings $False -AllRequestOutOfPolicy $False -AllRequestInPolicy $True -AllowConflicts $false -MaximumConflictInstances 0 -ConflictPercentageAllowed 0
    
#		Add-MailboxPermission $LeftName -User "ACL.UL.FullMailboxRight" -AccessRights FullAccess
		Add-MailboxPermission $LeftName -User "DBS.CRP.RRS.Admins" -AccessRights FullAccess
		Add-MailboxPermission $LeftName -User $Room.ResourceDelegates -AccessRights FullAccess
		Add-MailboxPermission $LeftName -User $Room.RRSRegionalOwner -AccessRights FullAccess

		Set-MailboxFolderPermission ($LeftName + $Folder) -User Default -AccessRights Reviewer
		Set-MailboxRegionalConfiguration $LeftName -TimeZone $Room.TimeZone
		
		#may have to add a Resource Delgate -ResourceDelegates 
		#add to RoomList here

		write-host "Adding " $Room.Name "to " $Room.RoomListName -ForegroundColor Yellow

		Add-DistributionGroupMember $Room.RoomListName -Member $Room.Name

		$LineToWrite = $RecordEvent + "INFO" + "`t" + $Room.Name + "`t" + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

		Set-MailboxCalendarConfiguration $LeftName -WorkingHoursTimeZone $Room.TimeZone
		#Write-Host "Time zone    = " $Room.TimeZone

		If ($Room.WorkHours -eq '$true')
		  {
		      Set-CalendarProcessing $LeftName -ScheduleOnlyDuringWorkHours $true
		      Set-MailboxCalendarConfiguration $LeftName -WorkDays $Room.WorkDays -WorkingHoursStartTime $Room.StartTime -WorkingHoursEndTime $Room.EndTime
		  }
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
