<#################################################################################
# 
# PowerShell source code
# Revision v1.5
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Configure Conference Rooms in Exchange Online
#    'Called By    : RoomResourceAdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 06/26/2012 03:35:00 PM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 06/26/2012 SAG Built from the Config Conference Room and
#    '                      added elements for restricted rooms
#    '               03/28/2013 SAG Commented out granting full permission to
#    '                      ACL.UL.FullMailboxRight 
#    '             : 07/11/2017 SAG Added Setting to retain Subject and comments
#    '                      when configuring the CalendarProcessing  
# ==========================================================================
#
#################################################################################>

#$text = "The purpose of this notification is to inform you of the status of your room/resource request.  If the notification indicates your meeting request has been Accepted, you do not need to do anything further.  If the notification indicates your meeting request is Tentative, please check to see if the room/resource is available at the requested time.  If it is not available at the requested time, please reschedule by removing the original room/resource and selecting a new room/resource.  Also, if your request is more than 90 days in the future, the designated Room/Resource Delegates will review and Accept or Decline your request."
$text = get-content e:\automation\scripts\RoomResourcePolicy.htm

$WhoAmI			= WhoAmI
$Folder			= ":\Calendar"

$LogDirectory		= "E:\Automation\ConfigRestrictedConferenceRoom\Log"
$LogFile		= $LogDirectory + "\" + "Log-ConfigRestrictedConferenceRoom.log"

$InputDirectory		= "E:\Automation\ConfigRestrictedConferenceRoom\Input"
$InputFile		= $InputDirectory + "\" + "Input-ConfigRestrictedConferenceRoom.csv"

$ReportDirectory	= "E:\Automation\ConfigRestrictedConferenceRoom\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-ConfigRestrictedConferenceRoom"


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
		[Boolean]$Room.AllRequestOutofPolicy = [System.Convert]::ToBoolean($Room.AllRequestOutofPolicy)

		Set-Mailbox $LeftName -CustomAttribute15 "ConfigConferenceRoom PS Date: $Date PS Time: $Time"

		Set-CalendarProcessing $LeftName -AutomateProcessing AutoAccept
		Set-CalendarProcessing $LeftName -AllRequestOutOfPolicy $Room.AllRequestOutofPolicy -AllBookInPolicy $Room.AllRequestOutOfPolicy -RequestOutOfPolicy $Room.ResourceUsers -BookInPolicy $Room.ResourceUsers -ResourceDelegates $Room.ResourceDelegates -ProcessExternalMeetingMessages $false -BookingWindowInDays $Room.BookingWindowInDays -AutomateProcessing $Room.AutomateProcessing -AllowConflicts $false -MaximumConflictInstances 3 -ConflictPercentageAllowed 20 -AddAdditionalResponse $true -AdditionalResponse $text -DeleteComments $false -DeleteSubject $False
		

#		Add-MailboxPermission $LeftName -User "ACL.UL.FullMailboxRight" -AccessRights FullAccess
		Add-MailboxPermission $LeftName -User "DBS.CRP.RRS.Admins" -AccessRights FullAccess
		Add-MailboxPermission $LeftName -User $Room.ResourceDelegates -AccessRights FullAccess
		Add-MailboxPermission $LeftName -User $Room.RRSRegionalOwner -AccessRights FullAccess

		Set-MailboxFolderPermission ($LeftName + $Folder) -User Default -AccessRights Reviewer
		Set-MailboxRegionalConfiguration $LeftName -TimeZone $Room.TimeZone
		
		#may have to add a Resource Delgate -ResourceDelegates 
		#add to RoomList here
		$Room.RoomListName
		$Room.Name
		Add-DistributionGroupMember $Room.RoomListName -Member $Room.Name

		$LineToWrite = $RecordEvent + "INFO" + "`t" + $Room.Name + "`t" + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

		If ($Room.WorkHours -eq '$true')
		  {
		  
		  Set-CalendarProcessing $LeftName -ScheduleOnlyDuringWorkHours $true
		  Set-MailboxCalendarConfiguration $LeftName -WorkingHoursTimeZone $Room.TimeZone -WorkDays $Room.WorkDays -WorkingHoursStartTime $Room.StartTime -WorkingHoursEndTime $Room.EndTime

		  #Write-Host "Left name is = " $Leftname
		  #Write-Host "Time zone    = " $Room.TimeZone
		  #Write-Host "Start time   = " $Room.StartTime
		  #Write-Host "End time     = " $Room.EndTime
		  
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
