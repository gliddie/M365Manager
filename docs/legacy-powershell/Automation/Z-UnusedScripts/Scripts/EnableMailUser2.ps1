#################################################################################
# 
# PowerShell source code
# Revision v1.05
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Enable MailUSer for existing AD User account
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
#    '               04/19/2011 MDS Added fingerprint CA15 
#    '		     04/25/2011 MDS Added cmdlet to remove Exchange PSSnapin
#    '		     04/27/2011 MDS Changed routing catch address
#    '		     05/05/2011 MDS Changed folder name to EnableMailUser
#    '		     05/13/2011 MDS Added condition for Internet Address @ul.com
#    '		     07/13/2011 KDH Added DisplayName Check
#    '		     
#    '
# ==========================================================================
#
#################################################################################

$WhoAmI			= WhoAmI
$DC			= "usnbkd300p.global.ul.com"

$LogDirectory		= "E:\Automation\EnableMailUser\Log"
$LogFile		= $LogDirectory + "\" + "Log-EnableMailUser.log"

$InputDirectory		= "E:\Automation\EnableMailUser\Input"
$InputFile		= $InputDirectory + "\" + "Input-EnableMailUser.csv"

$ReportDirectory	= "E:\Automation\EnableMailUser\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-EnableMailUser"

$PrimaryMail		= "@ul.com"
$RoutingDomain		= "@global.ul.com"

	# Add the Exchange PowerShell Snap-in
		Add-PSSnapin *Exchange* -erroraction SilentlyContinue

	# Add the ActiveDirectory modules.
		Import-Module ActiveDirectory

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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "Enable MailUser script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		

# =============================================================================================================================================

if (Test-Path $InputFile)
{					
	$UserAccounts = Import-CSV $InputFile

	ForEach ($User in $UserAccounts)
	{
	   if ($User.UPN.contains("@"))
	    {
		$LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + $User.Mail + "`t" + $User.LegacyMail + "`t" + $DC + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

		$atUPN  = $User.UPN.indexOf("@")

		$atMAIL = $User.MAIL.indexOf("@")


		$empID   = $User.UPN.substring(0,$atUPN)
		$Name    = $User.MAIL.substring(0,$atMAIL)
		$addMail = $User.LegacyMail.split(",")

		$LineToWrite = $RecordEvent + "INFO" + "`t" + $empID + "`t" + $Name + "`n"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite


		Enable-MailUser $User.UPN -Alias $empID -ExternalEmailaddress ($Name + $PrimaryMail) -PrimarySmtpAddress ($Name + $PrimaryMail) -DomainController:$DC
		
		# Start Check for Internet Mail = @ul.org
		   if ($User.Mail.contains($PrimaryMail))
		    {
			  Set-MailUser $User.UPN -EmailAddresses (((Get-MailUser $User.UPN -DomainController:$DC).EmailAddresses)+=($Name + $RoutingDomain)) -CustomAttribute15 "EnableMailUser PS Date: $Date PS Time: $Time" -DomainController:$DC
write-host "Skip adding Notes Internet Address @ul.org"
		    }
		   else
		    {
			  Set-MailUser $User.UPN -EmailAddresses (((Get-MailUser $User.UPN -DomainController:$DC).EmailAddresses)+=$User.MAIL,($Name + $RoutingDomain)) -CustomAttribute15 "EnableMailUser PS Date: $Date PS Time: $Time" -DomainController:$DC
write-host "Add Notes Internet Address @xx.ul.org"
		    }
		
		# End Check for Internet Mail = @ul.org
		
		ForEach ($mail in $addMail)
		{
		  $add = $mail

		if($add.contains("@"))
		  {
		  Set-MailUser $User.UPN -EmailAddresses (((Get-MailUser $User.UPN -DomainController:$DC).EmailAddresses)+=$add) -DomainController:$DC
		  #write-host $add " was added"
			$LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + $Name + "`t" + $add + "`t" + " proxyAddress added" + "`n"
			Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		  }

		else
		  {
		  #write-host "Nothing to add"
			$LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + $Name + "`t" + " no legacyAddress to add" + "`n"
			Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		  }
	        }

		# kdh - UPN is valid.  All other processing has been done.  Lets check the displayname and 
		#       set it to Last, First M. if needed.

		$LDAPFilter = "(userPrincipalName=" + $User.UPN + ")"
		$ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,proxyaddresses
		If ($ADUser.displayName.contains(",") -eq $false)
		{
			If ($ADUser.initials.length -gt 0)
			{
				$ADName = $ADUser.sn + ", " + $ADUser.givenName + " " + $ADUser.initials.substring(0,1) + "."
			} else {
				$ADName = $ADUser.sn + ", " + $ADUser.givenName
			}
		        Set-ADUser $ADuser -DisplayName $ADName
			$LineToWrite = $RecordEvent + "INFO" + "`t" + "Changed displayName to: " + $ADName + "`n"
			Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
			
		}
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
		
}
else
{
	write-host "Could not find input file :^) " $InputFile
	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Could not find input file " + "`n"
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
}

# =============================================================================================================================================

	# Remove the Exchange PowerShell Snap-in
		Remove-PSSnapin *Exchange* -ErrorAction SilentlyContinue
		
	# Retrieve the current Date and Time for use in log files
		$uDate = get-date -uformat %D
		$uTime = get-date -uformat %T
		
		$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"

	# Write end record to the log file
		$LineToWrite = $RecordEvent + "STOP" + "`t" + "This instance is stopping." + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

# End of script #
