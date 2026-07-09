#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Sets folder permission levels
#    'Called By    : NewSharedMailbox.ps1
#                  : EUMMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Kelly Salvatori (HP)
#    'Date Created : 10/31/2011 12:00:00 PM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '             : 11/08/2011 06:30:00 PM
#    '             : Added 'GrantSendOnBehalfTo', filtered out folders you can't stamp,
#    '             : "Top of Information Store" is referenced properly, improved logging.
#    '             : 4/23/2013 - SAG
#    '             : Added check so that .RE groups are not granted 'GrantSendOnBehalfTo'
#    '             : 10/ - SAG
#    '             : Added "Calendar Logging" folder to the "Ignore Folders" group
#    '
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "AddFolderPermissions"
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
		
#	
# =============================================================================================================================================
#
# Connect to Office 365
#
invoke-expression -Command .\ConnectO365.ps1

# =============================================================================================================================================

#
# Set mailbox folder permissions
if (Test-Path $InputFile) {					
	$MBXs = Import-CSV $InputFile
  	ForEach ($MBX in $MBXs) {
		if ($MBX.Mail.contains("@")) {
   			write-host "Setting permissions on: " $MBX.Mail " <-- " $MBX.DelegateToAdd " (" $MBX.Permission ")" -ForegroundColor Cyan
			$LineToWrite = "INFO" + "`t" + $MBX.Mail + "`t" + $MBX.DelegateToAdd + "`t" + $MBX.Permission + "`n" 
			WriteReportEvent
            if (($MBX.DelegateToAdd.contains(".ED@")) -OR ($MBX.DelegateToAdd.contains(".AU@")))
            {
                Write-Host "GrantSendOnBehalf to Delegate Group:  " $MBX.DelegateToAdd -ForegroundColor Cyan
                Set-Mailbox -Identity $MBX.Mail -GrantSendOnBehalfTo ((Get-Mailbox -Identity $MBX.Mail).GrantSendOnBehalfTo += $MBX.DelegateToAdd)
            }
			$MBXFolders = Get-MailboxFolderStatistics -Identity $MBX.Mail | % {$_.Identity.ToString().Split("\")[1..100] -join "\"}
			ForEach ($Folder in $MBXFolders) {
				if ($Folder.Equals("Recoverable Items") -or $Folder.Equals("Calendar Logging") -or $Folder.Equals("Deletions") -or $Folder.Equals("Purges") -or $Folder.Equals("Versions")) {
				# Ignore folder
				}
				else {
					if ($Folder.Equals("Top of Information Store")) {
						Add-MailboxFolderPermission -Identity ($MBX.Mail + ":\") -User $MBX.DelegateToAdd -AccessRights $MBX.Permission
					}
					else {
						Add-MailboxFolderPermission -Identity ($MBX.Mail + ":\" + $Folder) -User $MBX.DelegateToAdd -AccessRights $MBX.Permission
					}
				}
			}
		}	
		else {
        		Write-Host $DG.Mail " - Invalid mailbox." -ForegroundColor Red
				$LineToWrite = "ERROR" + "`t" + $MBX.Mail + " - Invalid mailbox." 
				WriteReportEvent
		}
	}
	# Rename the input file for future reference
	Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log") -ErrorAction SilentlyContinue
    if (Test-Path $InputFile)
    {
        $loop = 0
        Do
        {
            write-host "The file" $InputFile "is open by another process please close the file and hit returnt to continue" -ForegroundColor Red -NoNewline
            $Cont = read-host
            Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log") -ErrorAction SilentlyContinue
            $loop++
        } while ((Test-Path $InputFile) -and ($loop -lt 10))
        
        If ($loop -ge 10)
        {
            Write-Host "Unable to rename this file" $InputFule "please manually rename it to" ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
        }
    }
}
else {
		write-host "Input file not found." -ForegroundColor Red
}

# End of script & Remove PS Session
#Remove-PSSession $Session
#$Session = $null

$LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
WriteLogEvent

# =============================================================================================================================================