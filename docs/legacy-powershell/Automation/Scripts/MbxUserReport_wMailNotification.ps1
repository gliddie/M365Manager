<#
#  Called by:  ReportingMenu.ps1
#
#
#	08/11/2016 - SAG - Added Details for the RoleAssignmentPolicy to the report
#   08/09/2017 - SAG - Added Sort-Object command to the import process of the User Mailboxes
#   01/05/2018 - SAG - Added check to see if this is an active user before sending email
#   01/11/2018 - SAG - Added code to send email on failure of successful completion
#   02/12/2018 - SAG - Added code so when the reports are finished it moves the files to usnbku134p
#   03/06/2018 - SAG - Modifed the write-host statements so that # of Folders or size is written on the same line
#   05/07/2018 - SAG - Modified \scripts directory to \ReportHistory on usnbku134p
#   12/06/2019 - SAG - Moved the location of the statements that write the user detail to be after the email is sent
#   12/09/2019 - SAG - Added code to resolve users with where the $mbx.alias is longer than 20 characters and the account does not have a manager configured
#   03/02/2020 - SAG - Added Try/Catch code when messages are sent to try and account for the connection to the Relay Server timing out
#   06/02/2020 - SAG - Modified to use the Connect-ExchangeOnline and the Get-EXOMailbox
#   Get-Credential | Export-Clixml c:\temp\myO365File.xml #this is the command used to create a new credential file
#   10/11/2020 - SAG - Changed USNBKU134P to USNBKUTIL100P
#
#>
Function Reconnect
{
    write-host ""
	# Close all sessions
    get-pssession | Remove-PSSession -Confirm:$false

    $me = whoami
    $dir = "c:\users\" + $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)) + "\documents\"
    $File = "my" + ($me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))).replace(".","") + "File.xml"
    $CredFile = $dir + $File
    $LiveCred = Import-Clixml $CredFile

#    $0365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
    Connect-ExchangeOnline -Credential $LiveCred
    Import-Module ActiveDirectory
    Import-Module MSOnline
    Connect-MsolService -Credential $LiveCred
#    Import-PSSession $0365Session -AllowClobber
	get-date |fl DateTime
}

Function GetUserDN($strUID)
{
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = "LDAP://DC=global,DC=ul,DC=com"
    $objSearcher.Filter = [string]::format("(sAMAccountName={0})",$strUID)
    $ux = $null
    $ux = $objSearcher.FindOne()
    
    if ($ux -eq $null)
    {
        return $null
    } else {
        return $ux.Properties.distinguishedname
    }
}

Function SendLargeMbxMessage
{
    $message.To.Clear()
    $message.CC.Clear()
    $SendTo = $u.mail.value

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

	    $frstline = $u.givenName.value + "-<p><p>You are receiving this message because " + $MbxType + " mailbox has reached <b>" + $SizeGB.'TotalItemSize(GB)' + " GB</b>. While you have not exceeded the maximum file size, having a large mailbox can cause:<br>"
    }
    else
    {
    	write-host " - Mailbox 45GB or Larger" -NoNewline
        $message.Subject = "Action Required - Mailbox Quota Exceeds 45 GB"
        $html = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\LargerThan45GBFinal.html

        $message.CC.Add($SendCC)
#        $message.CC.Add("Sandi.Glazebrook@ul.com")    
        $frstline = $u.givenName.value + "-<p><p>You are receiving this message because " + $MbxType + " mailbox has reached <b>" + $SizeGB.'TotalItemSize(GB)' + 	" GB. <font color=red> Once your mailbox reaches 50 GB in size, you will not be able to send or receive email messages until your mailbox size is reduced.</font></b></p>"
    }
    $message.Body = $msgfont + $frstline + $html
    $message.To.Add($SendTo)
    try
    {
#    	$client.Send($message)
    }
    catch
    {
        write-host " - Pausing 60 Seconds to try to establish a new connecto to the Relay server"  -NoNewLine
        Start-Sleep -Seconds 60
#        $client.Send($message)
    }
    $Global:MsgSentCnt++
    write-host " - Message Sent" -NoNewLine
}

Function SendTooManyFoldersMbxMessage
{

    $message.Subject = "Action Required: Mailbox Folder Limit Exceeded" 
    $html = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\TooManyFoldersFinal.html

#    $SendTo = $MbxDetails.UserPrincipalName
    $SendTo = $u.mail.value
	
	$Mbxtype = "your"

	if ($MbxDetails.RecipientTypeDetails -eq "Shared")
	{
		$MbxType = "this shared"
	}
	
	$message.To.Clear()
    $message.CC.Clear()

	$frstline = $u.givenName.value + "-<p><p>You are receiving this message because " + $MbxType + " mailbox contains <b>" + $MbxFldr.Count + "</b> user-created folders. Having such as large number of folders in your mailbox can cause:<br>"
	$message.Body = $msgfont + $frstline + $html
	$message.To.Add($SendTo)
    try
    {
#    	$client.Send($message)
    }
    catch
    {
        write-host " - Pausing 60 Seconds to try to establish a new connecto to the Relay server" -NoNewLine
        Start-Sleep -Seconds 60
#        $client.Send($message)
    }
    $Global:MsgSentCnt++
    write-host " - Message Sent" -NoNewLine
}

start-transcript

$InputFile = "c:\temp\Input-UserMbxUPN.csv"
$Rest = "N"
#$reconnectThreshold = 500 #reconnect to O365 every 500 mailboxes
$reconnectThreshold = New-TimeSpan -Minutes 45
$Global:MsgSentCnt = 0

$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server 
$from = New-Object System.Net.Mail.MailAddress "UL.Technology.Services@ul.com" , "UL Technology Services"
$to = $from 
$message = new-object  System.Net.Mail.MailMessage $from, $to 
$message.IsBodyHtml = $true
$msgfont = "<basefont face=verdana size=2.5 color=black>"
$SendCC = ""

Write-Host "Error Count at start of script: " $Error.count

Write-Host "Starting All User Mailbox Report This Will take some time to complete......"
Write-Host ""

Reconnect
$processedCount = 0
$stopwatch = [diagnostics.stopwatch]::StartNew()

If ((Test-path $InputFile) -eq "True")
{
    if ((Test-path "c:\temp\AllUserMbxDetails.csv") -eq "True")
    {
        $mbxprocessed = import-csv c:\temp\AllUserMbxDetails.csv
        If ($mbxprocessed.count -gt 0)
        {
            $LastProc = $mbxprocessed.EmpNo[$mbxprocessed.count-1]
            $Rest = "Y"
        }
    }
}
else
{
	$Mbx = Get-Recipient -Resultsize Unlimited |where {$_.RecipientTypeDetails -eq "UserMailbox"}
	"Alias" > $InputFile
	Add-Content $InputFile $mbx.alias
	
	Write-Host "`nCreating Report on c:\temp\AllUserMbxDetails.csv" -ForegroundColor Yellow
	write-host ""
	$strPath = "C:\temp\AllUserMbxDetails.csv"
	Out-File -FilePath $strPath -InputObject $text

	$text = "EmpNo,DisplayName,MailboxType,BusinessUnit,WhenCreated,WhenChanged,RetentionPolicy,RoleAssignmentPolicy,UserCreatedFolders,TotalSizeInBytes,TotalSize(GB)"
	Out-File -FilePath c:\temp\AllUserMbxDetails.csv -InputObject $text
	Out-File -FilePath c:\temp\AllUserMbxTooManyFldrsDetails.csv -InputObject $text
	Out-File -FilePath c:\temp\AllUserMbxOvr30GBDetails.csv -InputObject $text
}

$Cnt = 1

$Mbx = Import-csv "c:\temp\Input-UserMbxUPN.csv" |sort-object "Alias"
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
            write-host "MbxCnt,EmpNo,ProcessedCnt,ElapsedTime"
        }

        If (($stopwatch.elapsed -gt 10) -and ($stopwatch.elapsed -lt 11))
        {
            write-host " - sleeping for 60 second" -NoNewline
            Start-Sleep -s 60
        }

		If ($Rest -eq "Y")
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
			if (([bool](Get-EXOMailbox $mbx.alias -ErrorAction SilentlyContinue) -eq "True"))
			{
				write-host `n$cnt,$mbx.alias,$processedCount,$stopwatch.elapsed -NoNewline
				$MbxDetails = Get-EXOMailbox $mbx.alias -Properties CustomAttribute1,CustomAttribute2,CustomAttribute8,WhenCreated,WhenChanged,RetentionPolicy,RoleAssignmentPolicy
				$MbxFldr = Get-EXOMailboxFolderStatistics $MbxDetails.UserPrincipalName
				$MbxStats = Get-EXOMailboxStatistics -Identity $MbxDetails.UserPrincipalName
				$MbxStats | Add-Member -MemberType ScriptProperty -Name TotalItemSizeInBytes -Value {$this.TotalItemSize -replace "(.*\()|,| [a-z]*\)", ""}
				$SizeGB = $MbxStats | Select-Object @{Name="TotalItemSize(GB)"; Expression={[math]::Round($_.TotalItemSizeInBytes/1GB,2)}}
				$text = """{0}"",{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}" -f $Mbx.Alias,$MbxDetails.Identity,$MbxDetails.RecipientTypeDetails,$MbxDetails.CustomAttribute8,$MbxDetails.WhenCreated,$MbxDetails.WhenChanged,$MbxDetails.RetentionPolicy,$MbxDetails.RoleAssignmentPolicy,$MbxFldr.Count,$MbxStats.TotalItemSizeInBytes,$SizeGB.'TotalItemSize(GB)'

				If (($MbxDetails.CustomAttribute2 -eq "A") -or ($MbxDetails.Alias -like "*svc*"))
				{
					If ($MbxFldr.Count -gt 949)
					{
						write-host " - Mailbox has 950 Folders or More" -NoNewline

						$strDN = GetUserDN $mbx.alias
						$strUserPath = [string]::format("LDAP://{0}", $strDN)
						$u = new-object System.DirectoryServices.DirectoryEntry($strUserPath)

						$strMgrPath = [string]::format("LDAP://{0}", $u.Manager.value)
						$u1 = new-object System.DirectoryServices.DirectoryEntry($strMgrPath)
						$SendCC = $u1.mail.value
						SendTooManyFoldersMbxMessage
                        Out-File -FilePath c:\temp\AllUserMbxTooManyFldrsDetails.csv -InputObject $text -Append
					}

					If ($SizeGB.'TotalItemSize(GB)' -ge "30")
					{
	#					write-host " - Mailbox 30GB or Larger"

						If ((($mbx.Alias).length) -gt 20)
                        {
                            $strDN = GetUserDN ($mbx.alias).Substring(0,20)
                        }
                        else
                        {
                            $strDN = GetUserDN $mbx.alias
                        }
						$strUserPath = [string]::format("LDAP://{0}", $strDN)
						$u = new-object System.DirectoryServices.DirectoryEntry($strUserPath)

                        if ($u.Manager.value -eq $null)
                        {
                            $SendCC = "EnterpriseMessagingServices@ul.com"
                        }
                        else
                        {
    						$strMgrPath = [string]::format("LDAP://{0}", $u.Manager.value)
	    					$u1 = new-object System.DirectoryServices.DirectoryEntry($strMgrPath)
		    				$SendCC = $u1.mail.value
                        }
			    		SendLargeMbxMessage
						Out-File -FilePath c:\temp\AllUserMbxOvr30GBDetails.csv -InputObject $text -Append

					}
				}
				else
				{
					write-host " - Terminated User" -NoNewline
				}
				
				Out-File -FilePath c:\temp\AllUserMbxDetails.csv -InputObject $text -Append

                # Increment processed counter
                $processedCount++
			}
			else
			{
				write-host " - No Mailbox" -NoNewline
				$text = """{0}"",{1}" -f $Mbx.Alias,"No Mailbox"
				Out-File -FilePath c:\temp\AllUserMbxDetails.csv -InputObject $text -Append
			}
		$Cnt = $Cnt + 1

		}
	}
}

catch
{
    $Catchfrom = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"
    $CatchTo = $CatchFrom 
    $CatchMessage = new-object  System.Net.Mail.MailMessage $from, $to
    $CatchMessage.IsBodyHtml = $true  
    $CatchSendTo = "Sandi.Glazebrook@ul.com"
    $CatchSendCC = ""

    $CatchMessage.Subject = "Monthly User Mailbox Reporting Process"
    $CatchMessage.Body = $msgfont + "<p>The UserMailbox Report has encountered an error and the Catch process is executing.  Below is the current completion status:<br>"
    $CatchMessage.Body = $CatchMessage.Body + "`nMbxCnt:  " + $cnt + "<br>"
    $CatchMessage.Body = $CatchMessage.Body + "Number of messages sent:  " + $Global:MsgSentCnt + "<br>"
    $CatchMessage.Body = $CatchMessage.Body + "Employee Number:  " + $mbx.alias + "<br>"
    $CatchMessage.Body = $CatchMessage.Body + "Processed Count:  " + $processedCount + "<br>"
    $CatchMessage.Body = $Catchmessage.Body + "Elapsed Time:  " + $stopwatch.elapsed + "<br>"
    $CatchMessage.Body = $CatchMessage.Body + "Error Count:  " + $Error.Count + "<br>"
    write-host $error
    $CatchMessage.Body = $CatchMessage.Body + "Error occurred running the Catch Code, failed to execute:  " + $error[1].Exception.ItemName + "<br>"
    $CatchMessage.Body = $CatchMessage.Body + "Error occurred running the Catch Code with error message:  " + $error[1].Exception.Message + "<br>"
    $CatchMessage.To.Clear()
	
    $CatchMessage.Body = $msgfont + $CatchMessage.body
    $CatchMessage.To.Add($CatchSendTo)
    #$CatchMessage.CC.Add($CatchSendCC)
    $client.Send($CatchMessage)

}

$MbxInp = Import-csv "c:\temp\Input-UserMbxUPN.csv" |sort-object "Alias"
$MbxOut = Import-csv "c:\temp\AllUserMbxDetails.csv"

If ($MbxOut.count -ge $MbxInp.count)
{
    $Date = get-date -Format "yyyyMM"
    
    $InpReportfile  = "\\usnbku134p\d$\ReportHistory\MailboxAuditReports\AllUserInput-UserMbxUPN_" + $Date + ".csv"
    $OutReportFile1 = "\\usnbku134p\d$\ReportHistory\MailboxAuditReports\AllUserMbxDetails_" + $Date + ".csv"
    $OutReportFile2 = "\\usnbku134p\d$\ReportHistory\MailboxAuditReports\AllUserMbxOvr30GBDetails_" + $Date + ".csv"
    $OutReportFile3 = "\\usnbku134p\d$\ReportHistory\MailboxAuditReports\AllUserMbxTooManyFldrsDetails_" + $Date + ".csv"

    Move-Item "c:\temp\Input-UserMbxUPN.csv" $InpReportFile
    Move-Item "c:\temp\AllUserMbxDetails.csv" $OutReportFile1
    Move-Item "c:\temp\AllUserMbxOvr30GBDetails.csv" $OutReportFile2
    Move-Item "c:\temp\AllUserMbxTooManyFldrsDetails.csv" $OutReportFile3

    $from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"
    $to = $from 
    $message = new-object  System.Net.Mail.MailMessage $from, $to
    $message.IsBodyHtml = $true  
    $SendTo = "Sandi.Glazebrook@ul.com"
    $SendCC = ""
    #$SendCC = "Sandi.Glazebrook@ul.com,RJ.Borja@ul.com"

    $message.Subject = "Monthly User Mailbox Reporting Process Complete"
    #$message.Attachments.Add($attach)
    $message.Body = "<p>The UserMailbox Report has completed.  Below are the details on the number of mailboxes processed.<br>"
    $message.Body = $message.Body + "                          Total Number of User Mailboxes:  " + $Mbx.Count
    $message.Body = $message.Body + "                                                  MbxCnt:  " + $cnt + "<br>"
    $message.Body = $message.Body + "                                         Employee Number:  " + $mbx.alias + "<br>"
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