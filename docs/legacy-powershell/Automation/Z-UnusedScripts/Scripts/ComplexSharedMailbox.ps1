#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Create Complex Shared Mailbox in Office 365
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 06/08/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '                
#    '               
#    '
# ==========================================================================
#
#################################################################################
 
$WhoAmI			= WhoAmI
$PrimaryMailDomain	= "@ul.com"
$RoutingDomain		= "@global.ul.com"

 
$LogDirectory		= "E:\Automation\ComplexSharedMailbox\Log"
$LogFile		= $LogDirectory + "\" + "Log-ComplexSharedMailbox.log"
 
$InputDirectory		= "E:\Automation\ComplexSharedMailbox\Input"
$InputFile		= $InputDirectory + "\" + "Input-ComplexSharedMailbox.csv"
 
$ReportDirectory	= "E:\Automation\ComplexSharedMailbox\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-ComplexSharedMailbox"
 
 
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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "ComplexSharedMailbox script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                                
 
# =============================================================================================================================================

# Connect to Office 365
	  $LiveCred = Get-Credential
	  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
	  Import-PSSession $Session

if (Test-Path $InputFile)
  {

	$SharedMailboxes = Import-CSV $InputFile

	ForEach ($Mailbox in $SharedMailboxes)
	  {
		$atMAIL   = $Mailbox.MAIL.indexOf("@")
		$LeftName = $Mailbox.MAIL.substring(0,$atMAIL)
		$addMail  = $Mailbox.LegacyMail.split(",")
		$cn       = $Mailbox.Name -replace (" ","_")

		New-Mailbox -Name $cn -DisplayName $Mailbox.Name -shared -Alias $cn -PrimarySmtpAddress ($LeftName + $PrimaryMailDomain)
		Set-CASMailbox $cn -ImapEnabled $false -PopEnabled $false
		Add-MailboxPermission $cn -User "ACL.UL.FullMailboxRight" -AccessRights FullAccess
		
		if ($Mailbox.Mail.contains($PrimaryMailDomain))
		  {
		    
		    if ($Mailbox.RoutingAddress.contains($PrimaryMailDomain))
		    {
		      Set-Mailbox $cn -ForwardingAddress $Mailbox.RoutingAddress -EmailAddresses (((Get-Mailbox $cn).EmailAddresses)+=($LeftName + $RoutingDomain)) -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 365 Day Delete" -CustomAttribute15 "ComplexSharedMailbox PS Date: $Date PS Time: $Time"
		    
		    }
		    
		    else
		    {
		      Set-Mailbox $cn -ForwardingSmtpAddress $Mailbox.RoutingAddress -EmailAddresses (((Get-Mailbox $cn).EmailAddresses)+=($LeftName + $RoutingDomain)) -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 365 Day Delete" -CustomAttribute15 "ComplexSharedMailbox PS Date: $Date PS Time: $Time"
		    
		    }
		  
		  }
		
		else
		  {
		    
		    if ($Mailbox.RoutingAddress.contains($PrimaryMailDomain))
		    {
		      Set-Mailbox $cn -ForwardingAddress $Mailbox.RoutingAddress -EmailAddresses (((Get-Mailbox $cn).EmailAddresses)+=$Mailbox.MAIL,($LeftName + $RoutingDomain)) -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 365 Day Delete" -CustomAttribute15 "ComplexSharedMailbox PS Date: $Date PS Time: $Time"
		    
		    }
		    
		    else
		    {
		      Set-Mailbox $cn -ForwardingSmtpAddress $Mailbox.RoutingAddress -EmailAddresses (((Get-Mailbox $cn).EmailAddresses)+=$Mailbox.MAIL,($LeftName + $RoutingDomain)) -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 365 Day Delete" -CustomAttribute15 "ComplexSharedMailbox PS Date: $Date PS Time: $Time"
		    
		    }
		    
		  }

		ForEach ($mail in $addMail)
		  {
		    $add = $mail

		    if($add.contains("@"))
		      {
		        Set-Mailbox $cn -EmailAddresses (((Get-Mailbox $cn).EmailAddresses)+=$add)
		        #write-host $add " was added"
			
			$LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Name + "`t" + $add + "`t" + " proxyAddress added" + "`n"
			Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		      }

		    else
		      {
		         #write-host "Nothing to add"
			$LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Name + "`t" + " no legacyAddress to add" + "`n"
			Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		      }
		  }


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