#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange
#    'Description  : Checks Access to Mailboxes and Allows Removal
#    'Called By    : 
#    'Calls        : Connect0365 scripts
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 10/21/2015
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 10/21/2015 Created for SDAP Team Use# 
#    '
#    '             : 11/27/2017 - Removes access from Mailbox and OneDrive and rehides
#    '             :    the user mailbox from the AddressBook
#    '             : 08/03/2018 - Added check to see if other individuals have been given access if so the disablement process is not executed
# ==========================================================================
#
#################################################################################

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

function WriteLogEvent
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LocalMachineName + "`t" + $LineToWrite
	Out-File -filepath $LogFile -append -noClobber -inputObject $RecordEvent
}

function WriteReportEvent 
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LineToWrite
	Out-File -filepath $ReportFile -append -noClobber -inputObject $RecordEvent
}

#################################################################################
# Declare Drive | Folders | and Files
	$FileName		= "RemoveMailboxPermission"
	$LogDrive		= "E:"
	$LogPath		= "\SDAP"
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

# Setup Folders and Files	
	CheckLogFiles
	$LineToWrite = "STAR" + "`t" + $FileName + " script has started"
	WriteLogEvent
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	WriteLogEvent
    WriteReportEvent
	
write-host
write-Host "Enter Employee Number : " -ForegroundColor Yellow -NoNewline
$ENo = Read-Host
$AccessCnt = 0			

$Inf = get-mailbox $ENo

write-host ""
write-host "DisplayName            : " $inf.DisplayName -ForegroundColor Green
write-host "Employee Number        : " $inf.Alias -ForegroundColor Green
write-host "Employee Type          : " $inf.CustomAttribute1 -ForegroundColor Green
write-host "LitigationHoldEnabled  : " $inf.LitigationHoldEnabled -ForegroundColor Green
write-host "LitigationHoldDate     : " $inf.LitigationHoldDate -ForegroundColor Green
write-host "LigitationHoldOwner    : " $inf.LitigationHoldOwner -ForegroundColor Green
write-host "O365 Retention Comment : " $inf.RetentionComment-ForegroundColor Green

$MbxAccess = get-MailboxPermission $ENo | where {$_.User -like "*ul.com*"}

$RemAccess = "N"

foreach ($MbxAccess in $MbxAccess)
{
    if ($MbxAccess.User -like "*@global*")
    {
        write-host "Accounts Granted Access: " $MbxAccess.User -ForegroundColor Red
        $AccessCnt++
        $RemAccess="Y"
    }
}           

If ($RemAccess -eq "Y")
{
    write-host "Enter Employee Number to Remove From Access or 0 to Exit: " -ForegroundColor Yellow -NoNewline
    $RemENo = Read-Host
    Do
    {
        Remove-MailboxPermission $ENo -AccessRights FullAccess -User $RemENo -Confirm:$False
        $LineToWrite = $WhoAmI + "`t" + "Removed " + $RemENo + " Access to " + $ENo + " Mailbox" + "`n"
        WriteReportEvent
        WriteLogEvent
        Write-Host "Removed" $RemENo "Access from" $ENo "Mailbox" -ForegroundColor Yellow

        If ($inf.CustomAttribute1 -like "*Ex-*")
        {
            If ($AccessCnt -lt 2)
			{
				Set-Mailbox $ENo -AccountDisabled:$True
				write-host "Disabling Mailbox Protocols for OWA and AllowMAC's..." -ForegroundColor Green
				Set-CASMailbox -Identity $ENo -OwaEnabled $false -EwsAllowMacOutlook $false
				$LineToWrite = $RecordEvent + "INFO" + "`t" + $ENo + "`t" + "Disabling O365 Mailbox, OWA Access and MacAccess" + "`n"
				Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
				invoke-expression -Command .\GetADUserInfo.ps1
				if ($Global:u.msExchHideFromAddressLists.value -eq $False)
				{
					write-host "Hiding User from the Address Book..." -ForegroundColor Green
					$LineToWrite = $RecordEvent + "INFO" + "`t" + $ENo + "`t" + "Hiding Account from Address Book" + "`n"
					Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
					$Global:u.msExchHideFromAddressLists.value = $True
					$Global:u.CommitChanges()	
				}
				else
				{
					write-host "User Already Hidden from the Address Book..." -ForegroundColor Red
				}
			}

            If ($inf.CustomAttribute1 -like "*Employee*")
            {
                $RemMgr = $RemENo + "@global.ul.com"
                $Usr = $ENo + "_global_ul_com"
                $Site = Get-SPOSite ("https://ul-my.sharepoint.com/personal/"+ $ENo +"_global_ul_com")
                $OneDrAccess = [bool](Set-SPOUser -site $Site -LoginName $RemMgr -IsSiteCollectionAdmin $False)
                If ($OneDrAccess -eq $True)
                {
                    write-host "Removing " $User.SUPERVISOR "access to OneDrive" -ForegroundColor Yellow
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Removed " + $ENo + " Supervisor Access to OneDrive" + "`n"
	    	        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                    Set-SPOUser -site $Site -LoginName $RemMgr -IsSiteCollectionAdmin $False 
                }
                else
                {
                    write-host "This individual " $User.SUPERVISOR "does not have access to users OneDrive"
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "This individual " + $ENo + " does not have access to users OneDrive" + "`n"
	    	        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                }
           }
        }

        Write-Host ""
        write-host "Enter Employee Number to Remove From Access or 0 to Exit? " -ForegroundColor Yellow -NoNewline
        $RemENo = Read-Host
        $AccessCnt--
    } while ($RemENo -ne 0)
}
else
{
    Write-host
    write-host "There are no users or mailboxes that have been given access to this account" -BackgroundColor Red
    Write-Host
}