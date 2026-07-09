#https://learn.microsoft.com/en-us/exchange/troubleshoot/user-and-shared-mailboxes/mailbox-recovery-in-exchange-online#online-account-hard-deleted-online-mailbox-soft-deleted
$UsrTimeZone = Get-TimeZone
$Year = (get-date).ToString("yyyy")
$ReportPath = "E:\Automation\RemoveSoftDelMailbox\Report\"
If (Test-Path $ReportPath) {} else {New-Item -Path $ReportPath -ItemType Directory}

$Filename = "RemoveSoftDelMailbox"
$ReportFile = $ReportPath + "Report-RemoveSoftDelMailbox-" + $Year + ".log"

$ENo = read-host "Enter Employee Number (Separate multiple with a comma CR to End)"
If ($ENo.Length -ge 5)
{
    
    $ENo = $ENo.Trim() -split ","
    $LineToWrite = $RecordEvent + "NEW " + "`t" + "Launched by: " + $WhoAmI
    WriteReportEvent

Do {
        Foreach ($E in $ENo)
        {
            If ($E -ne 0)
            {
                $UPN = $E + "@global.ul.com"
            #    $UPN = $ENo + "@global.ul.com"
                $Success = ""
                $Exists = [bool]($SDel = Get-mailbox –softdeletedmailbox –identity $UPN |select-object IsInactiveMailbox)
    
                write-host "`nProcessing Mailbox " $UPN
                $LineToWrite = $RecordEvent + "STRT" + "`t" + "Processing Mailbox " + $UPN
                WriteReportEvent

                If ($Exists -eq $True)
                {
                    $LineToWrite = $RecordEvent + "STAT" + "`t`t" + "SoftDeleted Mailbox IsInactiveMailbox = " + $SDel.IsInactiveMailbox
                    WriteReportEvent        
                    If ($SDel.IsInactiveMailbox -eq $True)
                    {
                        write-host "`tIsInactiveMailbox equals $true"
                        $ErrorActionPreference = "SilentlyContinue"
                        $RMbx = Get-Mailbox $UPN -softdeletedmailbox| Select Name, DisplayName, MicrosoftOnlineServicesID, ExchangeGuid
                        $Success = [bool](New-Mailbox -Name $UPN -inactivemailbox $RMbx.ExchangeGuid -MicrosoftOnlineServicesID $RMbx.MicrosoftOnlineServicesID -Password (ConvertTo-SecureString -String 'Pa##w0rd goes here' -AsPlainText -Force))
                        $ErrorActionPreference = "Continue"
	                    If ($Success -eq $True)
	                    {
                            $LineToWrite = $RecordEvent + "STAT" + "`t`t" + "New Mailbox Created"
                            WriteReportEvent 
		                    do
		                    {
			                    write-host "`tChecking if Mailbox exists"
			                    Start-Sleep -Seconds 15
			                    $Exists = [bool](get-mailbox $UPN)
		                    } While ($Exists -eq $False)
	                    }
	                    else
	                    {
		                    write-host "`tUnable to create a new mailbox for: " $UPN -foregroundcolor Red
                            $LineToWrite = $RecordEvent + "FAIL" + "`t`t" + "Unable to Create a New Mailbox"
                            WriteReportEvent 
	                    }
                    }
                    else
                    {
                        write-host "`tIsInactiveMailbox equals $false"
            #            Undo-SoftDeletedMailbox $ENo -WindowsLiveID $UPN -Password (ConvertTo-SecureString -String 'Pa$$word1' -AsPlainText -Force)
                        Undo-SoftDeletedMailbox $E -WindowsLiveID $UPN -Password (ConvertTo-SecureString -String 'Pa$$word1' -AsPlainText -Force)
                        $LineToWrite = $RecordEvent + "UNDO" + "`t`t" + "Restored SoftDeletedMailbox"
                        WriteReportEvent 
                        Start-Sleep -Seconds 15
                    }

                    If (($SDel.IsInactiveMailbox -eq $False) -or ($Success -eq $True))
                    {
	                    do
	                    {
		                    $Exists = [bool]($UsrId = get-MgUser -UserId $UPN -ErrorAction SilentlyContinue)
		                    If ($Exists -eq $False)
		                    {
			                    Start-Sleep -Seconds 15
		                    }
	                    } While ($Exists -eq $False)

	                    $MbxRet = get-mailbox $UPN

	                    If ($MbxRet.DelayReleaseHoldApplied -eq $true)
	                    {
                            $LineToWrite = $RecordEvent + "REMV" + "`t`t" + "Removing the DelayedReleaseHold"
                            WriteReportEvent 
		                    write-host "`tRemoving the DelayedReleaseHold"
		                    set-mailbox $UPN -DelayReleaseHoldApplied $false -DelayHoldApplied $False
	                    }

	                    If ($MbxRet.RetentionHoldEnabled -eq $True)
	                    {
                            $LineToWrite = $RecordEvent + "REMV" + "`t`t" + "Removing the RetentionHoldEnabled"
                            WriteReportEvent
		                    write-host "`tRemoving the RetentionHoldEnabled"
		                    set-mailbox $UPN -RetentionHoldEnabled $False
	                    }

                    #	$UsrId = get-MgUser -UserId $UPN
                        #Remove-MsolUser -UserPrincipalName $UPN -Force
                        Remove-MgUser -UserId $UPN
	                    #Remove-MsolUser -UserPrincipalName $UPN -RemoveFromRecycleBin -force
                        Remove-MgDirectoryDeletedItem -DirectoryObjectId $UsrId.Id
                    }
                }
                else
                {
                    write-host "`tUnable to find SoftDeletedMailbox for: " $UPN -foregroundcolor Red
                    $LineToWrite = $RecordEvent + "FAIL" + "`t`t" + "Unable to find SoftDeletedMailboxCreate a New Mailbox"
                    WriteReportEvent
                }
            }
        }
        write-host "`nEnter Employee Number (Separate multiple with a comma CR to End): " -ForegroundColor Cyan -NoNewline
        $ENo = read-host
        $ENo = $ENo.Trim() -split ","
    }While ($ENo -ne 0)

$LineToWrite = $RecordEvent + "END " + "`t" + "Processing Complete"
WriteReportEvent
}