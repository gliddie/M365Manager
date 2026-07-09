#################################################################################
# 
# PowerShell source code
# Revision v1.02
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Configure mailbox for existing mail enabled user account
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 04/18/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '               04/19/2011 MDS Added Remove-PSSession cmdlet
#    '               06/03/2011 MDS Added Retention Policy and User Role 
#    '                      Assignment
#    '               03/28/2013 SAG Commented out granting full permission to
#    '                      ACL.UL.FullMailboxRight 
#    '	             
# ==========================================================================
#
#################################################################################

$WhoAmI			= WhoAmI

$LogDirectory		= "E:\Automation\ConfigMailBox\Log"
$LogFile		= $LogDirectory + "\" + "Log-ConfigMailBox.log"

$InputDirectory		= "E:\Automation\ConfigMailBox\Input"
$InputFile		= $InputDirectory + "\" + "Input-ConfigMailBox.csv"

$ReportDirectory	= "E:\Automation\ConfigMailBox\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-ConfigMailBox"

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
	   if ($User.UPN.contains("@"))
	    {
		$LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

		$atUPN = $User.UPN.indexOf("@")


		$empID		= $User.UPN.substring(0,$atUPN)
		$RoutingAddress	= $User.RoutingAddress


		$LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + $RoutingAddress + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite


		Set-MailBox $User.UPN -ForwardingSmtpAddress $RoutingAddress -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 365 Day Delete"
		Set-CASMailBox $User.UPN -ImapEnabled $false -PopEnabled $false
#		Add-MailboxPermission $User.UPN -User "ACL.UL.FullMailboxRight" -accessRights FullAccess

	     }
	    
	   else
	    {
	      # UPN is not valid
	      	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Invalid UPN in input file " + "`n"
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
