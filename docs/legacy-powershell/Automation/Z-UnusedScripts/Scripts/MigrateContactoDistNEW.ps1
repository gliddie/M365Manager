#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Migrate Universal Distribution Group to Universal Security Group
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Kelly Salvatori (HP)
#    'Date Created : 10/26/2011 11:15:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '				 Set 'CustomAttribute15' to "MigrateUDGToUSGFlag" for UDGs that you want to convert
#    '				    For example: Set-MailContact <Group_Name> -CustomAttribute15 "MigrateContactToUSDFlag"
#    '
#    'History      :         
#    '             : 
#    '             : 
#    '
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "MigrateContactToDist"
	$LogDrive		= "E:"
	$LogPath		= "\Automation"
	$LogFolder		= "\" + $FileNAme
	$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
	
	$LogFile		= $LogDirectory + "Log-" + $FileName + ".log"
	$InputFile		= $LogDirectory + "Input-" + $FileName + ".csv"
	$ReportFile		= $LogDirectory + "Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
# Retrieve the user name and set DG Owner
	$WhoAmI			= WhoAmI
		
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name
	
# =============================================================================================================================================

function CheckLogFiles {
	# Create Files folder $LogDirectory if it's not present
	if (Test-Path $LogDirectory)
		{
		# the directory is present
		}
	else
		{ 
		mkdir $LogDirectory
		}
} #end CheckLogFiles

function WriteLogEvent {
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LocalMachineName + "`t" + $LineToWrite
	Out-File -filepath $LogFile -append -noClobber -inputObject $RecordEvent
} #end WriteLogEvent

function WriteReportEvent {
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LineToWrite
	Out-File -filepath $ReportFile -append -noClobber -inputObject $RecordEvent
} #end WriteReportEvent

# Setup Folders and Files	
	CheckLogFiles
		
	$LineToWrite = "STAR" + "`t" + $FileName + " script has started"
	WriteLogEvent
	
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	WriteLogEvent
		
# =============================================================================================================================================

# Connect to Office 365
if ($Session -eq $null) {
	write-host "Connect to Office 365."
	$LiveCred	= Get-Credential
	$Session	= New-PSSession -ConfigurationName Microsoft.Exchange `
		-ConnectionUri https://ps.outlook.com/powershell/ `
		-Credential $LiveCred `
		-Authentication Basic `
		-AllowRedirection
	
	if ($? -eq $true){
		Import-PSSession $Session
	}
	else {
		write-host "Could not establish session with Office 365."
		$LineToWrite = "STOP" + "`t" + "Could not establish session with Office 365." + "`n"
		WriteLogEvent
	}
}

else {
	write-host "Session with Office 365 already exists."
}

write-host "Find flagged MailContacts."
$UDG = Get-MailContact -Filter {CustomAttribute15 -like "*MigrateContactToUSDFlag*"} -Result Unlimited 

if ($UDG -ne $null) {
  	 ForEach ($DG in $UDG) {
		if ($DG.RecipientTypeDetails -eq "MailContact") {
   			write-host "Convert MailContact: " $DG.DisplayName
			$LineToWrite = $DG.DisplayName + "," + $DG.PrimarySmtpAddress + "," + $DG.legacyExchangeDN
			WriteReportEvent

			Remove-MailContact $DG.PrimarySmtpAddress -confirm:$false

			New-DistributionGroup -Name ("LST." + $DG.DisplayName) `
				-PrimarySmtpAddress ("LST." + $DG.PrimarySmtpAddress) `
				-Alias ("LST. " + $DG.Alias) `
				-ManagedBy "MBX.DistGroup.Owner"
			Set-Group -identity $DG.DisplayName `
				-Notes "Owner:  Needed"

			$USG	= Get-DistributionGroup $DG.PrimarySmtpAddress
			$USGProxys	= $Dg.EMailAddresses += ("x500:" + $DG.legacyExchangeDN)

# The 'MemberDepartRestriction' and 'MemberJoinRestriction' parameters must be "Closed" for USG so these are not set.		 		
			Set-DistributionGroup $USG.PrimarySmtpAddress `
		 		-EmailAddresses $USGProxys `
		 		-RequireSenderAuthenticationEnabled $DG.RequireSenderAuthenticationEnabled `
		 		-BypassSecurityGroupManagerCheck `
				-UMDtmfMap $DG.UMDtmfMap `
				-HiddenFromAddressListsEnabled $DG.HiddenFromAddressListsEnabled `
		 		-CustomAttribute15 ($FileName + " PS Date: " + (get-date -uformat %D) + " PS Time: " + (get-date -uformat %T))

			write-host ("Adding Members to USG: " + $USG.PrimarySMTPAddress)
			$LineToWrite = $USG.DisplayName + " - Adding Members to Distribution Group."
			WriteReportEvent
					
		}
		else {
			write-host "Flagged object is not a MailContact: " $DG.PrimarySMTPAddress
		}
	}
}
else {
	write-host "Could not find any MailContacts: " $DG.PrimarySMTPAddress
}


# End of script #
	Remove-PSSession $Session
	$Session = $null

	$LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
	WriteLogEvent

# =============================================================================================================================================