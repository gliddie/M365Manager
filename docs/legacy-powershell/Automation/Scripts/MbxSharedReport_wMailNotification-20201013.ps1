<#
#
#  Called by:  ReportingMenu.ps1
#
#	08/11/2016 - SAG - Added Details for the RoleAssignmentPolicy to the report
#   02/12/2018 - SAG - Added code to move the files to usnbku134p when complete and send an eamail
#   05/07/2018 - SAG - Changec \Scripts to \ReportHistory on usnbku134p
#   10/07/2019 - SAG - Added a Message Sent Count to know who many email were sent during the threshold period
#   06/02/2020 - SAG - Modified to use Connect-ExchangeOnline and replaced get-mailbox with Get-EXOMailbox
#   10/11/2020 - SAG - Changed USNBKU134P to USNBKUTIL100P
#
#Get-Credential | Export-Clixml c:\temp\myO365File.xml #this is the command used to create a new credential file
#
#>

Function Reconnect
{
    # Close all sessions
#    get-pssession | Remove-PSSession -Confirm:$false

    $MsgSentCnt = 0
#    $0365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
    Connect-ExchangeOnline -Credential $LiveCred
    Import-Module ActiveDirectory
    Import-Module MSOnline
    Connect-MsolService -Credential $LiveCred
#    Import-PSSession $0365Session -AllowClobber
}

Function SendLargeMbxMessage
{
    $SendTo = $MbxDetails.UserPrincipalName
    $message.To.Clear()
    $message.CC.Clear()

    $Mbxtype = "your"
    if ($MbxDetails.RecipientTypeDetails -like "*Shared*")
    {
	    $MbxType = "this shared"
    }

    if ($SizeGB.'TotalItemSize(GB)' -lt 45)
    {
    	write-host " - Mailbox 30GB or Larger" -NoNewline
        $message.Subject = "Action Required: Mailbox Exceeds 30 GB"
        $html = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\LargerThan30GBFinal.html 

	    $frstline = "You are receiving this message because " + $MbxType + " mailbox has reached <b>" + $SizeGB.'TotalItemSize(GB)' + " GB</b>. While you have not exceeded the maximum file size, having a large mailbox can cause:<br>"
    }
    else
    {
    	write-host " - Mailbox 45GB or Larger" -NoNewline
        $message.Subject = "Action Required - Mailbox Quota Exceeds 45 GB"
        $html = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\LargerThan45GBFinalShared.html

#        $message.CC.Add($SendCC)
#        $message.CC.Add("Sandi.Glazebrook@ul.com")    
        $frstline = "You are receiving this message because " + $MbxType + " mailbox has reached <b>" + $SizeGB.'TotalItemSize(GB)' + 	" GB. <font color=red> Once your mailbox reaches 50 GB in size, you will not be able to send or receive email messages until your mailbox size is reduced.</font></b></p>"
    }
    $message.Body = $msgfont +$frstline + $html
    $message.To.Add($SendTo)
	$client.Send($message)
    Write-Host " - Message Sent" -NoNewline
    $MsgSentCnt++
}

Function SendTooManyFoldersMbxMessage
{

    $message.Subject = "Action Required: Mailbox Folder Limit Exceeded" 
    $html = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\TooManyFoldersFinal.html

    $SendTo = $MbxDetails.UserPrincipalName
	$message.To.Clear()
    $message.CC.Clear()
    	
	$Mbxtype = "your"

	if ($MbxDetails.RecipientTypeDetails -like "Shared")
	{
		$MbxType = "this shared"
	}
	
	$frstline = "You are receiving this message because " + $MbxType + " mailbox contains <b>" + $MbxFldr.Count + "</b> user-created folders. Having such as large number of folders in your mailbox can cause:<br>"
	$message.Body = $msgfont + $frstline + $html
	$message.To.Add($SendTo)
	$client.Send($message)
    Write-Host " - Message Sent" -NoNewline
    $MsgSentCnt++
}

# Load credential
$me = whoami
$dir = "c:\users\" + $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)) + "\documents\"
$File = "my" + ($me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))).replace(".","") + "File.xml"
$CredFile = $dir + $File
$LiveCred = Import-Clixml $CredFile
$InputFile = "c:\temp\Input_AllSharedMbxDetails.csv"
#$reconnectThreshold = 500 #reconnect to O365 every 500 mailboxes
$reconnectThreshold = New-TimeSpan -Minutes 30

$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server 
$from = New-Object System.Net.Mail.MailAddress "UL.Technology.Services@ul.com" , "UL Technology Services"           
$to = $from 
$message = new-object  System.Net.Mail.MailMessage $from, $to 
$message.IsBodyHtml = $true
$msgfont = "<basefont face=verdana size=2.5 color=black>"
$SendCC = ""

Write-Host "Starting All Shared Mailbox Report This Will take some time to complete......"
Write-Host ""

Reconnect
$processedCount = 0
$strPath = "c:\temp\AllSharedMbxDetails.csv"
$stopwatch = [diagnostics.stopwatch]::StartNew()

If ((Test-path $InputFile) -eq "True")
{
    $mbxprocessed = import-csv $strPath
    if ($mbxprocessed.count -gt 0)
    {
        $LastProc = $mbxprocessed.EmpNo[$mbxprocessed.count-1]
        $Rest = "Y"
    }
}
else
{
	$Mbx = Get-Recipient -Resultsize Unlimited |where {$_.RecipientTypeDetails -eq "SharedMailbox"}
	"Alias" > $InputFile
	Add-Content $InputFile $mbx.alias
	
	Write-Host "Creating Report on c:\temp\AllSharedMbxDetails.csv" -ForegroundColor Yellow
	write-host ""
	$strPath = "c:\temp\AllSharedMbxDetails.csv"

	$text = "EmpNo,DisplayName,MailboxType,WhenCreated,WhenChanged,RetentionPolicy,RoleAssignmentPolicy,UserCreatedFolders,TotalSizeInBytes,TotalSize(GB)"
	Out-File -FilePath c:\temp\AllSharedMbxDetails.csv -InputObject $text
	Out-File -FilePath c:\temp\AllSharedMbxTooManyFldrsDetails.csv -InputObject $text
	Out-File -FilePath c:\temp\AllSharedMbxOvr30GBDetails.csv -InputObject $text
}

$Cnt = 1

$Mbx = Import-csv $InputFile |Sort-Object "Alias"
Write-Host " Current Error Count: " $Error.Count
Write-Host "Total Number of User Mailboxes: " $Mbx.Count
write-host "MbxCnt,EmpNo,ProcessedCnt,ElapsedTime" -NoNewline

try
{		
	foreach ($Mbx in $Mbx)
	{
#        if($processedCount -ge $reconnectThreshold)
        if($stopwatch.elapsed -ge $reconnectThreshold)
        {
            Reconnect
            $processedCount = 0
            $stopwatch = [diagnostics.stopwatch]::StartNew()
            write-host "MbxCnt,EmpNo,ProcessedCnt,ElapsedTime" -NoNewline
        }

		If (($Rest -eq "Y") -and ($mbxprocessed.count -gt 0))
		{
			If ($LastProc -ne $Mbx.alias)
			{
				$cnt = $cnt + 1
			}
			else
			{
				$Rest = "N"
				$cnt = $cnt + 1
			}
    	}
		else
		{
            $Rest = "N"	
			if (([bool](Get-EXOMailbox $mbx.alias -ErrorAction SilentlyContinue) -eq "True"))
			{
				write-host `n$cnt,$mbx.alias,$processedCount,$stopwatch.elapsed -NoNewline
				$MbxDetails = Get-EXOMailbox $mbx.alias -Properties WhenCreated,WhenChanged,RetentionPolicy,RoleAssignmentPolicy
				$MbxFldr = Get-EXOMailboxFolderStatistics $mbx.alias
				$MbxStats = Get-EXOMailboxStatistics -Identity $mbx.alias
				$MbxStats | Add-Member -MemberType ScriptProperty -Name TotalItemSizeInBytes -Value {$this.TotalItemSize -replace "(.*\()|,| [a-z]*\)", ""}
				$SizeGB = $MbxStats | Select-Object @{Name="TotalItemSize(GB)"; Expression={[math]::Round($_.TotalItemSizeInBytes/1GB,2)}}

				$text = """{0}"",{1},{2},{3},{4},{5},{6},{7},{8},{9}" -f $Mbx.Alias,$MbxDetails.Identity,$MbxDetails.RecipientTypeDetails,$MbxDetails.WhenCreated,$MbxDetails.WhenChanged,$MbxDetails.RetentionPolicy,$MbxDetails.RoleAssignmentPolicy,$MbxFldr.Count,$MbxStats.TotalItemSizeInBytes,$SizeGB.'TotalItemSize(GB)'
				Out-File -FilePath c:\temp\AllSharedMbxDetails.csv -InputObject $text -Append

				If ($MbxFldr.Count -gt 949)
				{
					write-host " - Mailbox with 950 Folders or More" -NoNewline
					Out-File -FilePath c:\temp\AllSharedMbxTooManyFldrsDetails.csv -InputObject $text -Append
                    SendTooManyFoldersMbxMessage
				}

				If ($SizeGB.'TotalItemSize(GB)' -ge "30")
				{
#					write-host "Mailbox 30GB or Larger" $Mbx.Alias
					Out-File -FilePath c:\temp\AllSharedMbxOvr30GBDetails.csv -InputObject $text -Append
                    SendLargeMbxMessage
				}

                # Increment processed counter
                $processedCount++
			}
			else
			{
				write-host $cnt,$mbx.alias," - No Mailbox" -NoNewline
				$text = """{0}"",{1}" -f $Mbx.Alias,"No Mailbox"
				Out-File -FilePath c:\temp\AllSharedMbxDetails.csv -InputObject $text -Append
			}
		$Cnt = $Cnt + 1
        }
	}
}

catch
{
    $from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"
    $to = $from 
    $message = new-object  System.Net.Mail.MailMessage $from, $to
    $message.IsBodyHtml = $true  
    $SendTo = "Sandi.Glazebrook@ul.com"
    $SendCC = ""

    $message.Subject = "Monthly Shared Mailbox Reporting Incomplete"
    #$message.Attachments.Add($attach)
    $message.Body = $msgfont + "<p>The UserMailbox Report has encountered an error and the Catch process is executing. Below is the current completion status.<br>"
    $message.Body = $message.Body + "                        Total Number of Shared Mailboxes:  " + $Mbx.Count
    $message.Body = $message.Body + "                                                  MbxCnt:  " + $cnt + "<br>"
    $message.Body = $message.Body + "                                    Shared Mailbox Alias:  " + $mbx.alias + "<br>"
    $message.Body = $message.Body + "                                         Processed Count:  " + $processedCount + "<br>"
    $message.Body = $message.Body + "                   Messages Sent During Threshold Period:  " + $MsgSentCnt + "<br>"
    $message.Body = $message.Body + "                               Current Threshold Setting:  " + $reconnectThreshold + "<br>"
    $message.Body = $message.Body + "                                Stopwatch Threshold Time:  " + $stopwatch.elapsed + "<br>"
    $message.Body = $message.Body + "Error occurred running the Catch Code, failed to execute:  " + $error[0].Exception.ItemName + "<br>"
    $message.Body = $message.Body + "Error occurred running the Catch Code with error message:  " + $error[0].Exception.Message + "<br>"
    $message.Body = $error[0].Exception |fl -Force
    $message.To.Clear()
	
    $message.Body = $msgfont + $message.body
    $message.To.Add($SendTo)
    $message.CC.Add($SendCC)
    $client.Send($message)

    #Reset the From Address
    
    $from = New-Object System.Net.Mail.MailAddress "UL.Technology.Services@ul.com" , "UL Technology Services"
    $to = $from 
    $message = new-object  System.Net.Mail.MailMessage $from, $to
    $message.IsBodyHtml = $true  

    Reconnect
    write-host ""
    Reconnect
}

$MbxInp = Import-csv "c:\temp\Input_AllSharedMbxDetails.csv" |sort-object "Alias"
$MbxOut = Import-csv "c:\temp\AllSharedMbxDetails.csv"

If ($MbxInp.count -ge $MbxOut.count)
{
    $Date = get-date -Format "yyyyMM"
    
    $InpReportfile = "\\usnbkutil100p\d$\ReportHistory\MailboxAuditReports\AllInputShrMbxUPN_" + $Date + ".csv"
    $OutReportFile1 = "\\usnbkutil100p\d$\ReportHistory\MailboxAuditReports\AllShrMbxDetails_" + $Date + ".csv"
    $OutReportFile2 = "\\usnbkutil100p\d$\ReportHistory\MailboxAuditReports\AllShrMbxOvr30GBDetails_" + $Date + ".csv"
    $OutReportFile3 = "\\usnbkutil100p\d$\ReportHistory\MailboxAuditReports\AllShrMbxTooManyFldrsDetails_" + $Date + ".csv"

    Move-Item "c:\temp\Input_AllSharedMbxDetails.csv" $InpReportFile
    Move-Item "c:\temp\AllSharedMbxDetails.csv" $OutReportFile1
    Move-Item "c:\temp\AllSharedMbxOvr30GBDetails.csv" $OutReportFile2
    Move-Item "c:\temp\AllSharedMbxTooManyFldrsDetails.csv" $OutReportFile3

    $from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"
    $to = $from 
    $message = new-object  System.Net.Mail.MailMessage $from, $to
    $message.IsBodyHtml = $true  
    $SendTo = "Sandi.Glazebrook@ul.com"
    $SendCC = ""
    #$SendCC = "Sandi.Glazebrook@ul.com,RJ.Borja@ul.com"

    $message.Subject = "Monthly Shared Mailbox Reporting Process Complete"
    #$message.Attachments.Add($attach)
    $message.Body = "<p>The SharedMailbox Report has completed.  Below are the details on the number of mailboxes processed.<br>"
    $message.Body = $message.Body + "                        Total Number of Shared Mailboxes:  " + $Mbx.Count
    $message.Body = $message.Body + "                                                  MbxCnt:  " + $cnt + "<br>"
    $message.Body = $message.Body + "                                    Shared Mailbox Alias:  " + $mbx.alias + "<br>"
    $message.Body = $message.Body + "                                         Processed Count:  " + $processedCount + "<br>"
    $message.Body = $message.Body + "                                            Elapsed Time:  " + $stopwatch.elapsed + "<br>"
    $message.To.Clear()
}
$message.Body = $msgfont + $message.body
$message.To.Add($SendTo)
$message.CC.Clear()
#$message.CC.Add($SendCC)
$client.Send($message)

write-host "Done"
stop-transcript
write-host "Done"