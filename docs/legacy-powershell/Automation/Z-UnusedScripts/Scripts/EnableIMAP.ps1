#################################################################################
# 
# PowerShell source code
# Revision v1.00
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Configure mailbox to allow IMAP connections
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 04/27/2011 10:00:00 AM
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

$LogDirectory		= "E:\Automation\EnableIMAP\Log"
$LogFile		= $LogDirectory + "\" + "Log-EnableIMAP.log"

$InputDirectory		= "E:\Automation\EnableIMAP\Input"
$InputFile		= $InputDirectory + "\" + "Input-EnableIMAP.csv"

$ReportDirectory	= "E:\Automation\EnableIMAP\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-EnableIMAP"

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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "EnableIMAP script has started"
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
	   if ($User.UPN.contains("@"))
	    {
		$LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

		Set-CASMailBox $User.UPN -ImapEnabled $true

	     }
	    
	   else
	    {
	      # UPN is not valid
	      	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Invalid UPN or empty line in input file " + "`n"
	      	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
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