#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Create new DG Contact for Coexistence
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 05/20/2011 10:00:00 AM
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
$PrimaryMail		= "@ul.com"

$LogDirectory		= "E:\Automation\NewDGContact\Log"
$LogFile		= $LogDirectory + "\" + "Log-NewDGContact.log"

$InputDirectory		= "E:\Automation\NewDGContact\Input"
$InputFile		= $InputDirectory + "\" + "Input-NewDGContact.csv"

$ReportDirectory	= "E:\Automation\NewDGContact\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-NewDGContact"


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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "Create DG Contact script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		

# =============================================================================================================================================

$LiveCred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
Import-PSSession $Session

if (Test-Path $InputFile)
{					
   $NewContacts = Import-CSV $InputFile

	ForEach ($Contact in $NewContacts)
	{


	$atMAIL = $Contact.MAIL.indexOf("@")
	$Name   = $Contact.MAIL.substring(0,$atMAIL)
	$RoutingAddress = $Contact.RoutingAddress

	New-MailContact -Name $Contact.DisplayName -DisplayName $Contact.DisplayName -ExternalEmailAddress $RoutingAddress
	$MyContact = Get-Contact $Contact.DisplayName
	
	#Sleep 60
	$LineToWrite = $RecordEvent + "INFO" + "`t" + $Contact.DisplayName + "`t" + $Contact.Mail + "`t" + $RoutingAddress + "`n"
	Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
	
	$LineToWrite = $RecordEvent + "INFO" + "`t" + $MyContact.DistinguishedName + "`n"
	Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
	
	#write-host $MyContact.DistinguishedName
	
		if ($Contact.MAIL.contains($PrimaryMail))
		{
			Set-MailContact $MyContact.DistinguishedName -EmailAddresses ("SMTP:" + $Contact.Mail),("smtp:" + $RoutingAddress) -CustomAttribute15 "NewDGContact PS Date: $Date PS Time: $Time"
		}
		else
		{
			Set-MailContact $MyContact.DistinguishedName -EmailAddresses ("SMTP:" + $Name + $PrimaryMail),("smtp:" + $Contact.Mail),("smtp:" + $RoutingAddress) -CustomAttribute15 "NewDGContact PS Date: $Date PS Time: $Time"
		}

	}
	

  #Rename the input file for future reference
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
