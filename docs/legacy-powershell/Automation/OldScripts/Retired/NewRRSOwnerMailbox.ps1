#################################################################################
# 
# PowerShell source code
# Revision v1.00
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Create New RRS Owner Mailbox in Office 365
#    'Called By    : RoomResourceMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 08/09/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :  03/08/2012  SAG  Changed name to from NewRrsSharedMailbox to
#    '             :     NewRRSOwnerMailbox and added settings to to reduce quota
#    '             :     to 1GB
#    '             : 
#    '               
#    '
# ==========================================================================
#
#################################################################################
 
$WhoAmI			= WhoAmI
$PrimaryMailDomain	= "@ul.com"
 
$LogDirectory		= "E:\Automation\NewRRSOwnerMailbox\Log"
$LogFile		= $LogDirectory + "\" + "Log-NewRRSOwnerMailbox.log"
 
$InputDirectory		= "E:\Automation\NewRRSOwnerMailbox\Input"
$InputFile		= $InputDirectory + "\" + "Input-NewRRSOwnerMailbox.csv"
 
$ReportDirectory	= "E:\Automation\NewRRSOwnerMailbox\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-NewRRSOwnerMailbox"
 
 
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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "NewRRSOwnerMailbox script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                                
 
# =============================================================================================================================================

# Connect to Office 365
#	  $LiveCred = Get-Credential
#	  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
#	  Import-PSSession $Session

if (Test-Path $InputFile)
  {

	$RRSOwnerMailboxes = Import-CSV $InputFile

	ForEach ($Mailbox in $RRSOwnerMailboxes)
	  {

New-Mailbox -Name $Mailbox.Name -shared -DisplayName $Mailbox.Name -Alias $Mailbox.Name -PrimarySmtpAddress ($Mailbox.Name + $PrimaryMailDomain)
Set-CASMailbox $Mailbox.Name -ImapEnabled $false -PopEnabled $false
Set-Mailbox $Mailbox.Name -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 365 Day Delete" -CustomAttribute15 "RRS Shared Mailbox" -HiddenFromAddressListsEnabled $true
Set-Mailbox $Mailbox.Name -ProhibitSendQuota .75GB -ProhibitSendReceiveQuota 1GB -IssueWarningQuota .5GB


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
