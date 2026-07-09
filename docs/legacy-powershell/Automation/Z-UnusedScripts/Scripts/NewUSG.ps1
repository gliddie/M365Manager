#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Creates new Universal Security Groups
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Kelly Salvatori (HP)
#    'Date Created : 10/27/2011 12:00:00 PM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 11/08/2011 04:30:00 PM        
#    '             :    Added additional error checking
#    'Modified	   : 06/26/2012 01:10:00 PM
#                  :    Added Set-Group to identify the owner and SD Ticket#
#    '             : 
#    '
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "NewUSG"
	$LogDrive		= "E:"
	$LogPath		= "\Automation"
	$LogFolder		= "\" + $FileNAme
	$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
	$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	$InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
	$ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
# Retrieve the user name
	$WhoAmI			= WhoAmI
		
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name
	
# =============================================================================================================================================

function CheckLogFiles
{
	# Create Files folder $LogDirectory if it's not present
	if (Test-Path $LogDirectory)
		{
		# the directory is present
		}
	else
		{ 
		mkdir $LogDirectory
		}
}
#end CheckLogFiles

function WriteLogEvent
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LocalMachineName + "`t" + $LineToWrite
	Out-File -filepath $LogFile -append -noClobber -inputObject $RecordEvent
} #end WriteLogEvent

function WriteReportEvent {
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LineToWrite
	Out-File -filepath $ReportFile -append -noClobber -inputObject $RecordEvent
}
#end WriteReportEvent

# Setup Folders and Files	
	CheckLogFiles
	$LineToWrite = "STAR" + "`t" + $FileName + " script has started"
	WriteLogEvent
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	WriteLogEvent
		
# =============================================================================================================================================

# Connect to Office 365
$Session = get-PSSession
if ($Session -eq $null)
{	
	write-host "Connect to Office 365."
	$LiveCred = Get-Credential
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
	if ($? -eq $true)
	{
		Import-PSSession $Session
	}
	else
	{
		write-host "Could not establish session with Office 365."
		$LineToWrite = "STOP" + "`t" + "Could not establish session with Office 365." + "`n"
		WriteLogEvent
	}
}
else
{
	write-host "Session with Office 365 already exists."
}

# Begin USG creation
if (Test-Path $InputFile) {					
	$USGs = Import-CSV $InputFile
  	ForEach ($DG in $USGs) {
		if ($DG.Mail.contains("@"))
        {
   			$atMail = $DG.Mail.indexOf("@")
			$DGAlias = $DG.Mail.substring(0,$atMail)
			$DGManagedByMembers = $DG.ManagedBy.Split(",")
			
			# Create the USG
			New-DistributionGroup -Name $DG.DisplayName `
				-PrimarySmtpAddress $DG.Mail `
				-Alias $DGAlias `
				-ManagedBy $DGManagedByMembers `
				-Type Security `
				| Out-Null

			Set-Group -identity $DG.DisplayName `
				-Notes ("Owner:  " + $DG.OwnerName + " - Per " + $DG.SDTicketNo)

			if (Get-DistributionGroup $DG.Mail)
            {
				write-host "USG created: " $DG.DisplayName " (" $DG.Mail ")"
				$LineToWrite = "INFO" + "`t" + $DG.DisplayName + "`t" + $DG.Mail + "USG Created" + "`n"
				WriteReportEvent
				
				$USG	= Get-DistributionGroup $DG.Mail
			
				# Configure the USG
				# The 'MemberDepartRestriction' and 'MemberJoinRestriction' parameters must be "Closed" for USG so these are not set.		 		
			
				# Lines below will take an additional input to hide from address book.
				# [Boolean]$DG.Hide = $false
				# $DG.Hide = [System.Convert]::ToBoolean("True")
				# 
				# Add this as a switch to the Set-DistibutionGroup cmdlet below
				# -HiddenFromAddressListsEnabled $DG.Hide `
			
				Set-DistributionGroup $USG.PrimarySMTPAddress `
		 			-RequireSenderAuthenticationEnabled $false `
		 			-BypassSecurityGroupManagerCheck `
					-CustomAttribute15 ($FileName + " PS Date: " + (get-date -uformat %D) + " PS Time: " + (get-date -uformat %T))
				
				# Add USG members
				write-host ("Adding Members to USG: " + $USG.DisplayName + " (" + $USG.PrimarySMTPAddress + ")")
				$LineToWrite = "`t" + $USG.DisplayName + "`t" + $USG.PrimarySMTPAddress + "`t" + "Adding Members to Distribution Group." + "`n"
				WriteReportEvent
				
				$addMember = $Dg.Members.split(",")
				if ($? -eq $true)
                {
					ForEach ($member in $addMember)
                    {
						if ($member.contains("@"))
                        {
							if (Get-Mailbox $member)
                            {
								Add-DistributionGroupMember $USG.PrimarySMTPAddress -Member $member -BypassSecurityGroupManagerCheck
								write-host "Added member:" $member
								$LineToWrite = "`t" + "PASS" + "`t" + $member + "`t" + "Added to Distribution Group" + "`t" + $USG.PrimarySMTPAddress + "`n"
								WriteReportEvent
							}
							else
                            {
								write-host "ERROR finding member: " $member
								$LineToWrite = "`t" + "FAIL" + "`t" + $member + "`t" + "ERROR finding recipient" + "`t" + $USG.PrimarySMTPAddress + "`n"
								WriteReportEvent
							}
						}
						else
                        {
							write-host "ERROR invalid member: " $member
							$LineToWrite = "`t" + "FAIL" + "`t" + $member + "`t" + "ERROR invalid recipient" + "`t" + $USG.PrimarySMTPAddress + "`n"
							WriteReportEvent
						}
					}
				}	
				else
                {
					write-host "No members to add"
				}
			}	
			else
            {
				Write-Host "USG not created: " $DG.DisplayName " (" $DG.Mail ")"
				$LineToWrite = "FAIL" + "`t" + $DG.DisplayName + "`t" + $DG.Mail + "`t" + "USG not created" + "`n"
				WriteReportEvent
			}
		}
		else
        {
			write-host "ERROR invalid USG: " $DG.Mail
			$LineToWrite = "`t" + "FAIL" + "`t" + $DG.Mail + "`t" + "ERROR invalid USG" + "`n"
			WriteReportEvent
		}
	}
# Rename the input file for future reference & Remove PS Session
Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")
}
else
{
	write-host "Input file not found."
}

# End of script #
	Remove-PSSession $Session
	$Session = $null

	$LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
	WriteLogEvent

# =============================================================================================================================================