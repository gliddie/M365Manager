#################################################################################
# 
# PowerShell source code
# Revision v1.00
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Prepare shared mailbox for Notes migration
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 05/31/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 
#                  : 11/30/2011 SAG Added Shared Mailbox Quota Settings        
#    '              
#    '
# ==========================================================================
#
#################################################################################

$WhoAmI			= WhoAmI

$LogDirectory		= "E:\Automation\PrepSharedMbxMig\Log"
$LogFile		= $LogDirectory + "\" + "Log-PrepSharedMbxMig.log"

$InputDirectory		= "E:\Automation\PrepSharedMbxMig\Input"
$InputFile		= $InputDirectory + "\" + "Input-PrepSharedMbxMig.csv"

$ReportDirectory	= "E:\Automation\PrepSharedMbxMig\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-PrepSharedMbxMig"

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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "PrepSharedMbxMig script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		

# =============================================================================================================================================

$LiveCred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
Import-PSSession $Session

if (Test-Path $InputFile)
{					
	$SharedMailboxes = Import-CSV $InputFile

	ForEach ($Mailbox in $SharedMailboxes)
	{
	   if ($Mailbox.Mail.contains("@"))
	    {
		$LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Mail + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite


		#Get Local Part of Mail Address
		$atMAIL   = $Mailbox.MAIL.indexOf("@")
		$LeftName = $Mailbox.MAIL.substring(0,$atMAIL)

		$MbxForwarder = Get-MailBox $LeftName

		Set-MailBox $LeftName -IssueWarningQuota 4.5GB -ProhibitSendQuota 4.75GB -ProhibitSendReceiveQuota 5GB
		
		  if ($MbxForwarder.ForwardingSmtpAddress -eq $null )
		  {
			$LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Mail + "`t" + "Forwarder was not set on this account." + "`n"
			Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		  }
		  
		 else
		  {
			$LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Mail + "`t" + "Removing Forwarder" + "`t" + $MbxForwarder.ForwardingSmtpAddress + "`n"
			Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
			
			Set-MailBox $LeftName -ForwardingSmtpAddress $null
		  }

	     }
	    
	   else
	    {
	      # Mail is not valid
	      	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Invalid mail address in input file " + "`n"
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
	#Remove-PSSession $Session
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
