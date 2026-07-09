#################################################################################
# 
# PowerShell source code
# Revision v1.5
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Configure new Equipment in Exchange Online
#    'Called By    : RoomResourceAdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 08/01/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : Mitch Skrove (HP)        
#    '               08/22/2011 10:00:00 AM 
#    '               Added section to set working hours only restriction 
#    '
#    '               08/23/2011 Mitch Skrove - Hard coded the -AllowConflicts to $false 
#    '                  Added working days only restriction 
#    '               08/26/2011 Mitch Skrove - Convert BookInPolicy to Boolean 
#    '               03/07/2012 - SAG - Add Set-CalendarProcess -AutoAccept Automate to
#    '                    this script
#    '               07/25/2012 - SAG - Add Set-CalendarProcess -DeleteComments and 
#    '                    -DeleteSubject to $FALSE'
#    '               10/2/2012 - SAG - Moved Set-MailboxCalendarConfiguration so that the
#    '                    -TimeZone is configured even if the work hours are not set
#    '               03/28/2013 - SAG - Commented out adding full permissions for the
#    '                     ACL.UL.FullMailbox Right
#    '               05/058/2016 - SAG - Added code for Restricted Equipment
# ==========================================================================
#
#################################################################################

$text = "The purpose of this notification is to inform you of the status of your room/resource request.  If the notification indicates your meeting request has been Accepted, you do not need to do anything further.  If the notification indicates your meeting request is Tentative, please check to see if the room/resource is available at the requested time.  If it is not available at the requested time, please reschedule by removing the original room/resource and selecting a new room/resource.  Also, if your request is more than 90 days in the future, the designated Room/Resource Delegates will review and Accept or Decline your request."

$WhoAmI			= WhoAmI
$Folder			= ":\Calendar"

$LogDirectory		= "E:\Automation\ConfigEquipment\Log"
$LogFile		= $LogDirectory + "\" + "Log-ConfigEquipment.log"

$InputDirectory		= "E:\Automation\ConfigEquipment\Input"
$InputFile		= $InputDirectory + "\" + "Input-ConfigEquipment.csv"

$ReportDirectory	= "E:\Automation\ConfigEquipment\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-ConfigEquipment"

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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "ConfigEquipment script has started"
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
    $EquipmentMailboxes = Import-CSV $InputFile

	ForEach ($EquipMbx in $EquipmentMailboxes)
	{
        write-host "Processing" $EquipMbx.Name -ForegroundColor Red
        $atMAIL   = $EquipMbx.MAIL.indexOf("@")
		$LeftName = $EquipMbx.MAIL.substring(0,$atMAIL)
#		[Boolean]$EquipMbx.BookInPolicy = [System.Convert]::ToBoolean($EquipMbx.BookInPolicy)

		Set-CalendarProcessing $LeftName -AutomateProcessing AutoAccept

		If (($EquipMbx.Restricted -ne "Y") -or ($EquipMbx.Restricted -ne "Yes"))
		{
			Set-Mailbox $LeftName -CustomAttribute15 "ConfigEquipment PS Date: $Date PS Time: $Time"
			Set-CalendarProcessing $LeftName -AllRequestOutOfPolicy $true -ResourceDelegates $EquipMbx.ResourceDelegates -ProcessExternalMeetingMessages $false -BookingWindowInDays $EquipMbx.BookingWindowInDays -AutomateProcessing $EquipMbx.AutomateProcessing -AllowConflicts $false -MaximumDurationInMinutes 1440 -MaximumConflictInstances 3 -ConflictPercentageAllowed 20 -AddAdditionalResponse $true -AdditionalResponse $text -DeleteComments $false -DeleteSubject $False
		}
		else
		{
			Set-Mailbox $LeftName -CustomAttribute15 "ConfigRestrictedEquipment PS Date: $Date PS Time: $Time"	
			Set-CalendarProcessing $LeftName -AllRequestOutOfPolicy $false -AllBookInPolicy $false -RequestOutOfPolicy $EquipMbx.ResourceUsers -RequestInPolicy $EquipMbx.ResourceUsers -AllRequestInPolicy $false -BookInPolicy $EquipMbx.ResourceUsers -ResourceDelegates $EquipMbx.ResourceDelegates -ProcessExternalMeetingMessages $false -BookingWindowInDays $EquipMbx.BookingWindowInDays -AutomateProcessing $EquipMbx.AutomateProcessing -AllowConflicts $false -MaximumDurationInMintues 1440 -MaximumConflictInstances 3 -ConflictPercentageAllowed 20 -AddAdditionalResponse $true -AdditionalResponse $text -DeleteComments $false -DeleteSubject $False
		}

		Write-Host "BookInPolicy is = " $EquipMbx.BookInPolicy
		if ($EquipMbx.Restricted -like "*Y*")
		{
		    Write-host "ResourceUsers is = " $EquipMbx.ResourceUsers
		}
		Write-Host "ResourceDelegates is = " $EquipMbx.ResourceDelegates
		Write-Host "BookingWindowInDays is = " $EquipMbx.BookingWindowInDays
		Write-Host "AutomateProcessing is = " $EquipMbx.AutomateProcessing

#		Add-MailboxPermission $LeftName -User "ACL.UL.FullMailboxRight" -AccessRights FullAccess
		Add-MailboxPermission $LeftName -User $EquipMbx.ResourceDelegates -AccessRights FullAccess
		Add-MailboxPermission $LeftName -User $EquipMbx.RRSRegionalOwner -AccessRights FullAccess
		Set-MailboxFolderPermission ($LeftName + $Folder) -User Default -AccessRights Reviewer
        If ($EquipMbx.TimeZone -notlike "")
        {
		    Set-MailboxRegionalConfiguration $LeftName -TimeZone $EquipMbx.TimeZone
        }
		
		$LineToWrite = $RecordEvent + "INFO" + "`t" + $EquipMbx.Name + "`t" + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

		Set-MailboxCalendarConfiguration $LeftName -WorkingHoursTimeZone $EquipMbx.TimeZone -WorkDays $EquipMbx.WorkDays -WorkingHoursStartTime $EquipMbx.StartTime -WorkingHoursEndTime $EquipMbx.EndTime

		If ($EquipMbx.WorkHours -eq "True")
        {
            Set-CalendarProcessing $LeftName -ScheduleOnlyDuringWorkHours $true
		    Set-MailboxCalendarConfiguration $LeftName -WorkDays $EquipMbx.WorkDays -WorkingHoursStartTime $EquipMbx.StartTime -WorkingHoursEndTime $EquipMbx.EndTime
 
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
