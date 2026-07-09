#################################################################################
# 
# PowerShell source code
# Revision v1.01
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Prepare mailbox for Notes migration
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 04/21/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '              
#    '		     07/05/2011 MDS Modified to accommodate migration input file.
#    '
# ==========================================================================
#
#################################################################################

$WhoAmI			= WhoAmI

$LogDirectory		= "E:\Automation\PrepMbxMig\Log"
$LogFile		= $LogDirectory + "\" + "Log-PrepMbxMig.log"

$InputDirectory		= "E:\Automation\PrepMbxMig\Input"
$InputFile		= $InputDirectory + "\" + "Input-PrepMbxMig.csv"

$ReportDirectory	= "E:\Automation\PrepMbxMig\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-PrepMbxMig"

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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "PrepMbxMig script has started"
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
		$LineToWrite = $RecordEvent + "INFO" + "`t" + $User.EmployeeID + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite


		$MbxForwarder = Get-MailBox $User.EmployeeID
		
		  if ($MbxForwarder.ForwardingSmtpAddress -eq $null )
		  {
			$LineToWrite = $RecordEvent + "INFO" + "`t" + $User.EmployeeID + "`t" + "Forwarder was not set on this account." + "`n"
			Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		  }
		  
		 else
		  {
			$LineToWrite = $RecordEvent + "INFO" + "`t" + $User.EmployeeID + "`t" + "Removing" + "`t" + $MbxForwarder.ForwardingSmtpAddress + "`n"
			Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
			
			Set-MailBox $User.EmployeeID -ForwardingSmtpAddress $null
		  }

	}

	# Rename the input file for future reference
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
		$LineToWrite = $RecordEvent + "STOP" + "`t" + "This instance is stopping." + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

# End of script #
