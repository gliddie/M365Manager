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
#    '             : 11/08/2011 - SAG - Added 'GrantSendOnBehalfTo', filtered out folders you can't stamp, "Top of Information Store" is referenced properly, improved logging.
#    '             : 04/23/2013 - SAG - Added check so that .RE groups are not granted 'GrantSendOnBehalfTo'
#    '             : 10/xx/xxxx - SAG - Added "Calendar Logging" folder to the "Ignore Folders" group
#    '             : 04/15/2020 - SAG - Modified code to identify the permission so you don't have to enter it into the input file.
#    '             : 08/22/2024 - SAG - Modified the values for the InputFile and Reportfile.  Added the Year into the report file and the autocreation of new directories for the Year.
#    '                       Also fixed code that is confirming that the new groups are created the Do{ statements were missing in the .AU and .RE groups.
#    '
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
#	$FileName		= "AddFolderPermissions"
#	$LogDrive		= "E:"
#	$LogPath		= "\Automation"
#	$LogFolder		= "\" + $FileNAme
#	$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
#	$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	$InputFile		= "e:\Automation\AddFolderPermissions\Input\Input-AddFolderPermissions.csv"
    $Year = (get-date).ToString("yyyy")
    $Path = "e:\Automation\AddFolderPermissions\Report\" + $Year
    If (Test-Path $path) {} else {New-Item -Path $Path -ItemType Directory}
	$ReportFile		= $Path + "\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
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
#invoke-expression -Command .\ConnectO365.ps1

# =============================================================================================================================================

#
# Set mailbox folder permissions
if (Test-Path $InputFile)
{					
	$MBXs = Import-CSV $InputFile
  	ForEach ($MBX in $MBXs) {
		if ($MBX.Mail.contains("@"))
		{
            $Loop = 0
			$MBXExists = [bool](get-mailbox $MBX.Mail -ErrorAction SilentlyContinue)
            Do{
				If ($MBXExists -eq $False)
				{
					if ($Loop -eq 0)
					{
						write-host "Waiting for Mailbox to be created..." -Foregroundcolor Cyan -NoNewline
						$Loop++
					}
					else	
					{
						write "..." -Foregroundcolor Cyan -NoNewline
					}
					Start-Sleep -Seconds 15
                    $MBXExists = [bool](get-mailbox $MBX.Mail -ErrorAction SilentlyContinue)
				}
			}while($MBXExists -ne "True")
            write-host ""
			
			$MBXFolders = Get-MailboxFolderStatistics -Identity $MBX.Mail | % {$_.Identity.ToString().Split("\")[1..100] -join "\"}

            if ((get-mailbox $MBX.mail).MailTip -eq $null)
            {
                $NewTip = ("Owners: " + (((Get-DistributionGroup $MBX.DelegateToAdd).Managedby) -join ", "))
                if ($NewTip.Length -gt 175)
                {
                    write-host "Maximum mail tip length exceeded truncating to 175 characters." -foregroundcolor Red
                    $NewTip = $NewTip.Substring(0,175)
                    $OldTip = (get-mailbox $MBX.Mail).MailTip
                    write-host "Replacing old MailTip: " $OldTip.Substring(16,$OldTip.Length-36) -foregroundcolor Cyan
                    write-host "                 With: " $NewTip
                }
                Set-Mailbox $MBX.Mail -MailTip $NewTip
            }

			# 3/23/2020 is the date when all new mailboxes started getting the 3 yr retention policy
            $MPolicy = (get-mailbox $mbx.Mail).RetentionPolicy
#            if (((get-mailbox $mbx.Mail).WhenCreated -gt (Get-Date -Year 2020 -Month 3 -Day 23)) -and ($(get-mailbox $mbx.Mail).RetentionPolicy -notlike "*3 yr*"))
            If (((get-mailbox $mbx.Mail).WhenCreated -gt (Get-Date -Year 2020 -Month 3 -Day 23)) -and (($MPolicy -notlike "*3 yr*") -or ($MPolicy -notlike "*Delete*")))
            {
                Set-MailBox $MBX.Mail -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL MRM Policy - 3 yr Delete" -RetentionHoldEnabled $false
    		    Set-Mailbox $MBX.Mail -UseDatabaseQuotaDefaults $False -IssueWarningQuota 45GB -ProhibitSendQuota 49.75GB -ProhibitSendReceiveQuota 50GB
                write-host "3 yr Retention Policy Applied to" $Mailbox.Name
	    	    $LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Name + "`t" + " Retention Policy Set" + "`n"
		        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
			# Check that IMAP and POP3 are disabled
    		    $ProtocolCheck = (get-CASMailbox $MBX.Mail) 
    		    if ($ProtocolCheck.ImapEnabled -eq $true)
	    	    {
		    	    Set-CASMailBox $MBX.Mail -ImapEnabled $false
    			    $LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Name + "`t" + "IMAP Protocol Disabled" + "`n"
	    		    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		        }
    		    if ($ProtocolCheck.PopEnabled -eq $true)
	    	    {
		    	    Set-CASMailBox $MBX.Mail -PopEnabled $false
    			    $LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Name + "`t" + "POP3 Protocol Disabled" + "`n"
	    		    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		        }
            }

            $Permission = ""
            if ($MBX.DelegateToAdd.contains(".ED@"))
            {
                $Permission = "Editor"
       			write-host "Setting permissions on: " $MBX.Mail " <-- " $MBX.DelegateToAdd " (Editor)" -ForegroundColor Cyan
	    		$LineToWrite = "INFO" + "`t" + $MBX.Mail + "`t" + $MBX.DelegateToAdd + "`t" + $Permission + "`n" 
		    	WriteReportEvent
                Write-Host "GrantSendOnBehalf to Delegate Group:  " $MBX.DelegateToAdd -ForegroundColor Cyan
                Do {
                    write-host "checking for DistributionGroup"
                    $Exists = [bool](Get-DistributonGroup $MBX.DelegateToAdd -ErrorAction SientlyContinue)
                    Start-Sleep -Seconds 3
                } while ($Exists -eq $False)
                Set-Mailbox -Identity $MBX.Mail -GrantSendOnBehalfTo ((Get-Mailbox -Identity $MBX.Mail).GrantSendOnBehalfTo += $MBX.DelegateToAdd)
                Add-MailboxPermission $MBX.Mail –User $MBX.DelegateToAdd –AccessRights FullAccess
                Add-RecipientPermission $MBX.Mail -Trustee $MBX.DelegateToAdd –AccessRights SendAs -Confirm:$false
        	}
            elseif ($MBX.DelegateToAdd.contains(".AU@"))
            {
                Do {
                    write-host "checking for DistributionGroup"
                    $Exists = [bool](Get-DistributonGroup $MBX.DelegateToAdd -ErrorAction SientlyContinue)
                    Start-Sleep -Seconds 3
                } while ($Exists -eq $False)
                $Permission = "PublishingAuthor"
                write-host "Setting permissions on: " $MBX.Mail " <-- " $MBX.DelegateToAdd " (PublishingAuthor)" -ForegroundColor Cyan
                Write-Host "GrantSendOnBehalf to Delegate Group:  " $MBX.DelegateToAdd -ForegroundColor Cyan
                Set-Mailbox -Identity $MBX.Mail -GrantSendOnBehalfTo ((Get-Mailbox -Identity $MBX.Mail).GrantSendOnBehalfTo += $MBX.DelegateToAdd)
            }
            elseif ($MBX.DelegateToAdd.contains(".RE@"))
            {
                Do {
                    write-host "checking for DistributionGroup"
                    $Exists = [bool](Get-DistributonGroup $MBX.DelegateToAdd -ErrorAction SientlyContinue)
                    Start-Sleep -Seconds 3
                } while ($Exists -eq $False)
                $Permission = "Reviewer"
                write-host "Setting permissions on: " $MBX.Mail " <-- " $MBX.DelegateToAdd " (Reviewer)" -ForegroundColor Cyan
            }
            else
            {
                write-host "Group name does not follow naming stand there is no .ED, .AU or .RE found." -ForegroundColor Red
                write-host "Group Name: " $mbx.DelegateToAdd
            }

            If ($Permission -ne "")
            {
                ForEach ($Folder in $MBXFolders)
                {
			        If ($Folder.Equals("Recoverable Items") -or $Folder.Equals("Calendar Logging") -or $Folder.Equals("Deletions") -or $Folder.Equals("Purges") -or $Folder.Equals("Versions"))
                    {
				        # Ignore folder
				    }
				    else
                    {
Start-Sleep -Seconds 3
					    if ($Folder.Equals("Top of Information Store"))
                        {
#                            Add-MailboxFolderPermission -Identity ($MBX.Mail) -User $MBX.DelegateToAdd -AccessRights $Permission
                            Add-MailboxFolderPermission -Identity ($MBX.Mail + ":\") -User $MBX.DelegateToAdd -AccessRights $Permission
					    }
					    else
                        {
						    Add-MailboxFolderPermission -Identity ($MBX.Mail + ":\" + $Folder) -User $MBX.DelegateToAdd -AccessRights $Permission
					    }
				    }
			    }
            }
            else
            {
                write-host "Skipping this group since it does not follow our naming standard"
            }
		}	
		else
        {
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
        } while ((Test-Path $InputFile) -and ($loop -lt 5))
        
        If ($loop -ge 5)
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