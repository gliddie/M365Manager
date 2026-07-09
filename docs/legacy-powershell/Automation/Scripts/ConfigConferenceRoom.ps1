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
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 08/01/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 08/22/2011 - M.Skrove (HP) - Added section to set 
#    '                     working hours only restriction 
#    '               08/23/2011 - M.Skrove (HP) - Hard coded the 
#    '                     -AllowConflicts to $false 
#    '                     Added working days only restriction 
#    '               08/26/2011 - M.Skrove (HP) - Convert BookInPolicy to 
#    '                     Boolean 
#    '               07/25/2012 - SAG - Add Set-CalendarProcess -DeleteComments
#    '                     and -DeleteSubject to $FALSE'
#    '               10/2/2012 - SAG = Moved set-MailboxCalendarConfiguration 
#    '                     so that the -Time Zone is set regardless of the 
#    '                     working hours being set
#    '               3/28/2013 - SAG - Commented out adding full permissions for the
#    '                     ACL.UL.FullMailbox Right
#    '               7/11/2017 - SAG - Modified the text that gets sent to import an
#    '                     html file so the details are formatted nicer.
# ==========================================================================
#
#################################################################################>

#$text = "The purpose of this notification is to inform you of the status of your room/resource request.  If the notification indicates your meeting request has been Accepted, you do not need to do anything further.  If the notification indicates your meeting request is Tentative, please check to see if the room/resource is available at the requested time.  If it is not available at the requested time, please reschedule by removing the original room/resource and selecting a new room/resource.  Also, if your request is more than 90 days in the future, the designated Room/Resource Delegates will review and Accept or Decline your request."
$text = get-content e:\automation\scripts\RoomResourcePolicy.htm

$WhoAmI			= WhoAmI
$Folder			= ":\Calendar"

$LogDirectory		= "E:\Automation\ConfigConferenceRoom\Log"
$LogFile		= $LogDirectory + "\" + "Log-ConfigConferenceRoom.log"

$InputDirectory		= "E:\Automation\ConfigConferenceRoom\Input"
$InputFile		= $InputDirectory + "\" + "Input-ConfigConferenceRoom.csv"

$ReportDirectory	= "E:\Automation\ConfigConferenceRoom\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-ConfigConferenceRoom"


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
		[Boolean]$Room.BookInPolicy = [System.Convert]::ToBoolean($Room.BookInPolicy)

		Set-Mailbox $Room.Mail -CustomAttribute15 "ConfigConferenceRoom PS Date: $Date PS Time: $Time"

		Set-CalendarProcessing $Room.Mail -AutomateProcessing AutoAccept
		Set-CalendarProcessing $Room.Mail -AllRequestOutOfPolicy $Room.BookInPolicy -ResourceDelegates $Room.BookOutPolicy -ProcessExternalMeetingMessages $false -BookingWindowInDays $Room.BookingWindowInDays -AutomateProcessing $Room.AutomateProcessing -AllowConflicts $false -MaximumConflictInstances 3 -ConflictPercentageAllowed 20 -AddAdditionalResponse $true -AdditionalResponse $text -DeleteComments $false -DeleteSubject $False

#		Add-MailboxPermission $LeftName -User "ACL.UL.FullMailboxRight" -AccessRights FullAccess
		Add-MailboxPermission $Room.Mail -User "DBS.CRP.RRS.Admins" -AccessRights FullAccess
		Add-MailboxPermission $Room.Mail -User $Room.BookOutPolicy -AccessRights FullAccess
		Add-MailboxPermission $Room.Mail -User $Room.RRSRegionalOwner -AccessRights FullAccess

		Set-MailboxFolderPermission ($Room.Mail + $Folder) -User Default -AccessRights Reviewer
        If ($Room.TimeZone -notlike "")
        {
		    Set-MailboxRegionalConfiguration $Room.Mail -TimeZone $Room.TimeZone
		}
		#may have to add a Resource Delgate -ResourceDelegates 
		#add to RoomList here

		write-host "Adding " $Room.Name "to " $Room.RoomListName -ForegroundColor Yellow

		Add-DistributionGroupMember $Room.RoomListName -Member $Room.Name

		$LineToWrite = $RecordEvent + "INFO" + "`t" + $Room.Name + "`t" + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

		Set-MailboxCalendarConfiguration $Room.Mail -WorkingHoursTimeZone $Room.TimeZone
		#Write-Host "Time zone    = " $Room.TimeZone

		If ($Room.WorkHours -eq '$true')
		  {
		  
		  Set-CalendarProcessing $Room.Mail -ScheduleOnlyDuringWorkHours $true
		  Set-MailboxCalendarConfiguration $Room.Mail -WorkDays $Room.WorkDays -WorkingHoursStartTime $Room.StartTime -WorkingHoursEndTime $Room.EndTime

		  #Write-Host "Left name is = " $Leftname
		  #Write-Host "Start time   = " $Room.StartTime
		  #Write-Host "End time     = " $Room.EndTime
		  
		  }
		
	  }

	# Rename the input file for future reference & Remove PS Session
    Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log") -ErrorAction SilentlyContinue
    if (Test-Path $InputFile)
    {
        $loop = 0
        Do
        {
            write-host "The file" $InputFile "is open by another process please close the file and hit returnt to continue" -ForegroundColor Red -NoNewline
            $Cont = read-host
            Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log") -ErrorAction SilentlyContinue
            $loop++
        } while ((Test-Path $InputFile) -and ($loop -lt 10))
        
        If ($loop -ge 10)
        {
            Write-Host "Unable to rename this file" $InputFule "please manually rename it to" ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
        }
    }
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
