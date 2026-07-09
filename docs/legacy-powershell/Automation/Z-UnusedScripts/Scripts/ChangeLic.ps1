#################################################################################
#
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Change Office 365 User License
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 10/07/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :
#
#
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "ChangeLic"
	$LogDrive		= "E:"
	$LogPath		= "\Automation"
	$LogFolder		= "\" + $FileNAme
	$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
	
	$LogFile		= $LogDirectory + "Log-" + $FileName + ".log"
	$InputFile		= $LogDirectory + "Input-" + $FileName + ".csv"
	$ReportFile		= $LogDirectory + "Report-" + $FileName + `
						"-Date" + ((get-date -uformat %D).Replace("/", "-")) + `
						"-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"

	write-host $InputFile
	
# Retrieve the user name and set DG Owner
	$WhoAmI			= WhoAmI
	
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name
	
# =============================================================================================================================================

function CheckLogFiles {
	# Create Files folder $LogDirectory if it's not present
	if (Test-Path $LogDirectory) {
		# the directory is present
	}
	else { 
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

# Setup licensing
# 7200 License Pack
$AccountSkuIdAdd = "ul:STANDARDPACK"
[String[]]$DisabledPlans = "MCOSTANDARD","SHAREPOINTSTANDARD"

# 800 License Pack
#$AccountSkuIdAdd = "ul:EXCHANGESTANDARD"
#$DisabledPlans = $null

$AccountSkuIdRemove = "ul:ENTERPRISEPACK"

$UsageLocation = "US"
$LicenseOptions = New-MsolLicenseOptions -AccountSkuId $AccountSkuIdAdd -DisabledPlans $DisabledPlans

# Connect to Office 365
$Creds = Get-Credential
Connect-MsolService -Credential $Creds


if (Test-Path $InputFile) {					
	write-host "Import CSV File."
	$oUsers = Import-CSV $InputFile


	if ($oUsers -ne $null) {
		write-host "Start modifing licenses."

		ForEach ($User in $oUsers) {
			if ($CurrentUser = Get-MsolUser -UserPrincipalName $User.UPN -ErrorAction SilentlyContinue) {
				Write-Host $User.UPN " - " $CurrentUser.DisplayName " - " $AccountSkuIdAdd
				
				Set-MsolUser -UserPrincipalName $User.UPN -UsageLocation $UsageLocation
				Set-MsolUserLicense -UserPrincipalName $User.UPN `
					-RemoveLicenses $AccountSkuIdRemove `
					-AddLicenses $AccountSkuIdAdd `
					-LicenseOptions $LicenseOptions `
					-ErrorAction SilentlyContinue
				
				if ($? -eq $false) {
					Write-Host $User.UPN " - ERROR setting license."
					$LineToWrite = "ERROR" + "`t" + $User.UPN + " - ERROR setting license."
					WriteReportEvent
				}
				else {
					Write-Host $User.UPN " - " $CurrentUser.DisplayName " - new license set."
					$LineToWrite = "SUCCESS" + "`t" + $User.UPN + "`t" + $AccountSkuIdAdd
					WriteReportEvent
				}
			}
			else {
				Write-Host $User.UPN " - ERROR finding user."
				$LineToWrite = "ERROR" + "`t" + $User.UPN + " - ERROR finding user."
				WriteReportEvent
			}
		}
	}
}

else {
	write-host "Input file not found."
}

# End of script #

	$LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
	WriteLogEvent

# =============================================================================================================================================