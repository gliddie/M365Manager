#################################################################################
# 
# PowerShell source code
# Revision v1.1
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Create new Equipment in Exchange Online
#    'Called By    : RoomResourceMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 08/01/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 03/07/2012  SAG Added Quota Settings and removed 
#    '             :    set-calendarprocessing  
#    '             : 07/25/2013  SAG Commented out adding the ACL.UL.FullMailboxRight
#    '             :    to the equipment mailbox.     
#    '             : 05/05/2016  SAG Added Code for Restricted Equipment
#    '
# ==========================================================================
#
#################################################################################

$WhoAmI			= WhoAmI

$LogDirectory		= "E:\Automation\NewEquipment\Log"
$LogFile		= $LogDirectory + "\" + "Log-NewEquipment.log"

$InputDirectory		= "E:\Automation\NewEquipment\Input"
$InputFile		= $InputDirectory + "\" + "Input-NewEquipment.csv"

$ReportDirectory	= "E:\Automation\NewEquipment\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-NewEquipment"


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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "NewEquipment script has started"
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
		$atMAIL   = $EquipMbx.MAIL.indexOf("@")
		$LeftName = $EquipMbx.MAIL.substring(0,$atMAIL)

		New-Mailbox -Name $EquipMbx.Name -Equipment -DisplayName $EquipMbx.Name -Alias $LeftName -PrimarySmtpAddress $EquipMbx.Mail
		Set-user $EquipMbx.Name -Office $Room.Location
		Set-Mailbox $LeftName -IssueWarningQuota 0.5GB -ProhibitSendQuota 0.75GB -ProhibitSendReceiveQuota 1.0GB
		
		If ($EquipMbx.Restricted -ne "Y")
		{
			Set-Mailbox $LeftName -CustomAttribute15 "NewEquipment PS Date: $Date PS Time: $Time"
		}
		else
		{		
			Set-Mailbox $LeftName -CustomAttribute15 "NewRestrictedEquipment PS Date: $Date PS Time: $Time"
		}
		
#		Add-MailboxPermission $LeftName -User "ACL.UL.FullMailboxRight" -AccessRights FullAccess
		
		$LineToWrite = $RecordEvent + "INFO" + "`t" + $EquipMbx.Name + "`t" + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
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