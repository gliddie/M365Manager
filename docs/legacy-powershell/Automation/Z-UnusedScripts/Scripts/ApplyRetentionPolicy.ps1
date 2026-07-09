#################################################################################
# 
# PowerShell source code
# Revision v1.01
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Create Apply Retention Policy in Office 365
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook (extracted out of New Shared Mailbox)
#    'Date Created : 10/17/2013 03:30:00 PM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '               
#    '
# ==========================================================================
#
#################################################################################
 
$WhoAmI			= WhoAmI
$PrimaryMailDomain	= "@ul.com"
$RoutingDomain		= "@global.ul.com"
 
$LogDirectory		= "E:\Automation\ApplyRetentionPolicy\Log"
$LogFile		= $LogDirectory + "\" + "Log-ApplyRetentionPolicy.log"
 
$InputDirectory		= "E:\Automation\ApplyRetentionPolicy\Input"
$InputFile		= $InputDirectory + "\" + "Input-ApplyRetentionPolicy.csv"
 
$ReportDirectory	= "E:\Automation\ApplyRetentionPolicy\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-ApplyRetentionPolicy"
 
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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "ApplyRetentionPolicy script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
#	
# =============================================================================================================================================
#
# Connect to Office 365
#
$Session = get-PSSession
#
if ($Session -eq $null)
{
	write-host "Connect to Office 365."
	$LiveCred = Get-Credential
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
#
	if ($? -eq $true)
	{
		Import-PSSession $Session
	}
	else
	{
		write-host "Could not establish session with Office 365."
		$LineToWrite = "STOP" + "`t" + "Could not establish session with Office 365." + "`n"
		WriteLogEvent
	}
}
else
{
	write-host "Session with Office 365 already exists."
	write-host ""
}
#
# =============================================================================================================================================
#
if (Test-Path $InputFile)
  {

	$ApplyRetentionPolicy = Import-CSV $InputFile

	ForEach ($Mailbox in $ApplyRetentionPolicy)
	  {
		Set-Mailbox $Mailbox.Name –RetentionHoldEnabled $true –StartDateForRetentionHold 04/01/2011
        write-host "Retention Policy Applied to" $Mailbox.Name
		$LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Name + "`t" + $add + "`t" + " Retention Policy Set" + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

	  }

	# Rename the input file for future reference & Remove PS Session
		Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
		Remove-PSSession $Session

  }

else

  {
	write-host "Could not find input file :^) " $InputFile
	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Could not find input file " + "`n"
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	Remove-PSSession $Session
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