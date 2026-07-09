#################################################################################
# 
# PowerShell source code
# Revision v1.01
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Create Apply Retention Policy in Office 365
#    'Called By    : NewSharedMailbox.ps1
#                  : EUMMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook (extracted out of New Shared Mailbox)
#    'Date Created : 10/17/2013 03:30:00 PM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 10/13/2016 -SAG - Added setting the RoleAssignmentPolicy and the RetentionPolicy also check to make 
#    '					         sure IMAP and POP3 are disabled if they are not they will be disabled.
#    '             : 02/28/2018 -SAG - Added code so that all IT staff are assigned an active 3 yr retention policy
#    '             : 03/23/2020 -SAG - Modified so that all new mailboxes have the 3 yr retention policy assigned
#    '             : 03/25/2020 -SAG - Added code so if the input file is in use to attempt the rename 10 times and the fail
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
invoke-expression -Command .\ConnectO365.ps1
#
# =============================================================================================================================================
#
if (Test-Path $InputFile)
  {

	$ApplyRetentionPolicy = Import-CSV $InputFile

	ForEach ($Mailbox in $ApplyRetentionPolicy)
	  {
		Set-MailBox $Mailbox.Name -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL MRM Policy - 3 yr Delete" -RetentionHoldEnabled $false
		Set-Mailbox $Mailbox.Name -UseDatabaseQuotaDefaults $False -IssueWarningQuota 45GB -ProhibitSendQuota 49.75GB -ProhibitSendReceiveQuota 50GB
        write-host "3 yr Retention Policy Applied to" $Mailbox.Name
		$LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Name + "`t" + " Retention Policy Set" + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
	# Check that IMAP and POP3 are disabled
		$ProtocolCheck = (get-CASMailbox $Mailbox.Name) 
		if ($ProtocolCheck.ImapEnabled -eq $true)
		{
			Set-CASMailBox $Mailbox.Name -ImapEnabled $false
			$LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Name + "`t" + "IMAP Protocol Disabled" + "`n"
			Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		}
		if ($ProtocolCheck.PopEnabled -eq $true)
		{
			Set-CASMailBox $Mailbox.Name -PopEnabled $false
			$LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Name + "`t" + "POP3 Protocol Disabled" + "`n"
			Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		}
	  }

	# Rename the input file for future reference & Remove PS Session
		Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log") -ErrorAction SilentlyContinue
        if (Test-Path $InputFile)
        {
            $loop = 0
            Do
            {
                write-host "The file " $InputFile "is open by another process please close the file and hit returnt to continue" -ForegroundColor Red -NoNewline
                $Cont = read-host
                Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log") -ErrorAction SilentlyContinue
                $loop++
            } while ((Test-Path $InputFile) -and ($loop -lt 10))
        }

        If ($loop -ge 10)
        {
            Write-Host "Unable to rename this file" $InputFule "please manually rename the file to" ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
        }

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