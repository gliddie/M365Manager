#################################################################################
# 
# PowerShell source code
# Revision v1.1
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Create new external contact
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 04/07/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : MDS Change Name (CN) to First_Last, avoid conflict with DirSynced        
#    '               User Account Name (CN)
#    '
# ==========================================================================
#
#################################################################################

$WhoAmI			= WhoAmI

$LogDirectory		= "E:\Automation\NewExternalContact\Log"
$LogFile		= $LogDirectory + "\" + "Log-NewExternalContact.log"

$InputDirectory		= "E:\Automation\NewExternalContact\Input"
$InputFile		= $InputDirectory + "\" + "Input-NewExternalContact2.csv"

$ReportDirectory	= "E:\Automation\NewExternalContact\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-NewExternalContact"


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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "Create External Contact script has started"
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

	$Name = $Contact.Alias
	
	New-MailContact -Name $Name -DisplayName $Contact.DisplayName -ExternalEmailAddress $Contact.RoutingAddress
	$MyContact = Get-Contact $Name
	
	$LineToWrite = $RecordEvent + "INFO" + "`t" + $Name + "`t" + $Contact.DisplayName + "`t" + $Contact.Alias + "`t" + $Contact.Mail + "`t" + $Contact.RoutingAddress + "`n"
	Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
	
	$LineToWrite = $RecordEvent + "INFO" + "`t" + $MyContact.DistinguishedName + "`n"
	Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
	
	  If ($Contact.Mail -eq $Contact.RoutingAddress)
	  {
	   #write-host "email addresses match"
	   #write-host $MyContact.DistinguishedName
	   Set-MailContact $MyContact.DistinguishedName -CustomAttribute15 "NewExternalContact PS Date: $Date PS Time: $Time" 
	  }
	
	  else
	  {
	   #write-host "email addresses Nomatch"
	   #write-host $MyContact.DistinguishedName
	    Set-MailContact $MyContact.DistinguishedName -EmailAddresses ("smtp:" + $Contact.Mail),("SMTP:" + $Contact.RoutingAddress) -CustomAttribute15 "NewExternalContact PS Date: $Date PS Time: $Time" 
	  }
	
	}
	

  #Rename the input file for future reference
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