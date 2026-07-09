
$WhoAmI			= WhoAmI

$LogDirectory		= "E:\Automation\RetentionHoldEnabled\Log"
$LogFile		= $LogDirectory + "\" + "Log-RetentionHoldEnabled.log"

$InputDirectory		= "E:\Automation\RetentionHoldEnabled\Input"
$InputFile		= $InputDirectory + "\" + "Input-RetentionHoldEnabled.csv"

$ReportDirectory	= "E:\Automation\RetentionHoldEnabled\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-RetentionHoldEnabled"

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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "ConfigMailBox script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		

# =============================================================================================================================================

$LiveCred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
Import-PSSession $Session

if (Test-Path $InputFile)
{					
	$UserAccounts = Import-CSV $InputFile

	ForEach ($User in $UserAccounts)
	{


		Set-Mailbox $User.UPN –RetentionHoldEnabled $true –StartDateForRetentionHold 04/01/2011
		$LineToWrite =  "Enabled Retention Hold for:  " + $User.UPN
		write-host $LineToWrite
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite =  "Starting ManagedFolderAssistantEnabled Retention Hold for:  " + $User.UPN
		write-host $LineToWrite
		start-ManagedFolderAssistant $User.UPN
		$LineToWrite =  "Completed ManagedFolderAssistantEnabled Retention Hold for:  " + $User.UPN
		write-host $LineToWrite
		$LineToWrite =  " "
		write-host $LineToWrite
				
	}

	# Rename the input file for future reference
		Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
		
}
else
{
	write-host "Could not find input file :^) " $InputFile
	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Could not find input file " + "`n"
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
}
Remove-PSSession $Session
# =============================================================================================================================================

	
	# Retrieve the current Date and Time for use in log files
		$uDate = get-date -uformat %D
		$uTime = get-date -uformat %T
		
		$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"

	# Write end record to the log file
		$LineToWrite = $RecordEvent + "STOP" + "`t" + "This instance is stopping." + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

# End of script #
