<#
#  Called by:  ReportingMenu.ps1
#
#	#   10/12/2022 - SAG - Adapted from the mailbox report w/notificationt hat included the review of number of folders.
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

{
    $htmlSig = Get-Content -LiteralPath c:\temp\EMSSignature.htm
    $message.IsBodyHtml = $true
    $msgfont = "<basefont face=Arial size=2.5 color=black>"
        
    write-host "Sending Mailbox Critical to: " $mbx.alias, $mbx.DisplayName
    If ($oed."Employee Number" -contains $mbx.alias)
    {
        write-host "This is a Gold Employee: " $mbx.alias, $mbx.DisplayName
        # Import Gold user message
        #$html = Get-Content -LiteralPath c:\temp\MailboxQuotaMessage.htm
    }
    else
    {
        write-host "This is NOT a Gold Employee: " $mbx.alias, $mbx.DisplayName
        $mbx = $
        $Size = $stat.totalitemsize.value.ToString()
        $val = ($Size.Substring(0,$Size.IndexOf("(")-3))

        If (($Val -ge 50) -and ($Size -like "*GB*"))
        {
            $message.Subject = "Action Required: Direct Report Mailbox Size is Exceeded"
            $frstline = "You are receiving this message because the individual " + $mbx.Name + " who reports to you is no longer able to send or receive any new messages.  The UL mailbox maximum size is 50 GB and their mailbox is " + ($Size.Substring(0,$Size.IndexOf("(")-1)) + " and has exceeded the maximum size.<br>"
            $html =Get-Content -LiteralPath c:\temp\MailboxoverQuotaMessage.htm
            $lastline = "The 3 yr retention policy will be applied to the  mailbox of " + $mbx.Name + " on " + (get-date).AddDays(31).ToString("yyyy-MMM-dd") + " if we do not receive a response to this message."
        }
        else
        {
            #Between 45GB and 50 GB
            If (($Val -ge 45) -and ($Size -like "*GB*"))
            {
                $message.Subject = "Action Required: Mailbox Size is Critical"
                $frstline = "<p>" + ($mbx.Name).SubString(0,$mbx.Name.IndexOf(" ")) + " - </p><p>You are receiving this message because your UL mailbox maximum size is 50 GB.  Your mailbox is currently <b>" + ($Size.Substring(0,$Size.IndexOf("(")-1)) + " </b> and space is considered critical. In <b>" + (49.75 - $val) + " GB</b> you will no longer be able to send new messages and in <b>" + (50 - $val) + " GB</b> you will no longer be able to send or receive new messages."
                $html = Get-Content -LiteralPath c:\temp\MailboxQuotaMessage.htm
            }
            else
            {
                If (($Val -ge 35) -and ($Size -like "*GB*"))
                {
                    $message.Subject = "Action Required: Large Mailbox Size"
                    $frstline = "<p>" + ($mbx.Name).SubString(0,$mbx.Name.IndexOf(" ")) + " - </p>"
                    $html = Get-Content -LiteralPath c:\temp\MailboxPerformanceMessage.htm
                }
                else
                #Smaller than 35GB
                {
                    $message.Subject = "Action Required: No Mailbox Retention Policy"
                    $frstline = "<p>" + ($mbx.Name).SubString(0,$mbx.Name.IndexOf(" ")) + " - </p>"
                    $html = Get-Content -LiteralPath c:\temp\MailboxSetRetentionMessage.htm
                }
            }
            $lastline = "The 3 yr retention policy will be applied to your mailbox on " + (get-date).AddDays(31).ToString("yyyy-MMM-dd") + " if we do not receive a response to this message."
        }
        $Message.Body = $msgfont + $frstline + $html + $lastline + $htmlSig
    }
    else
    {
        $message.Subject = "Action Required - Mailbox Quota Exceeds 45 GB"
        $html = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\LargerThan45GBFinal.html
        
        $message.To.Clear()
        $frstline = "<p>You are receiving this message because " + $MbxType + " mailbox has reached <b>" + $colAddr.Size + 	" GB. <font color=red> Once your mailbox reaches 50 GB in size, you will not be able to send or receive email messages until your mailbox size is reduced.</font></b></p>"
    }

#	$message.Body = $msgfont + $frstline + $html
    $message.To.Add($SendTo)
	$client.Send($message)
}
#stop-transcript
}

<#
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
#>

start-transcript

#$InputFile = "c:\temp\Input-NoActiveRetentionUPN.csv"
#$Rest = "N"
#$reconnectThreshold = 500 #reconnect to O365 every 500 mailboxes
#$reconnectThreshold = New-TimeSpan -Minutes 60

$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server 
$from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "UL Enterprise Messaging Services"
$to = $from 
$message = new-object  System.Net.Mail.MailMessage $from, $to 
$message.IsBodyHtml = $true
$msgfont = "<basefont face=verdana size=2.5 color=black>"
$SendCC = ""

Write-Host "Starting All User Mailbox Report This Will take some time to complete......"
Write-Host ""

<#
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

    $date = (Get-date).ToString('yyyy-MMdd')
    $RptFile = "c:\temp\MailboxSizeReport-" + $date + ".csv"
	$text = "UPN;Status;DisplayName;PrimaryAddress;MbxType;LegalHold;RetentionPolicy;TotalSize;TotalBytes;WhenCreated;Warning;ProhibSend;ProhibSendReceive;MaxRecSize;RetentionHold"
	Out-File -FilePath $RptFile -InputObject $text
}
#>

$Script:MsgSentCnt = 0
$Cnt = 1
$AllMbx = get-mailbox -Resultsize Unlimited -SortBy Alias
$OED = Import-Csv "\\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv" |Where-Object {((($_."VIP STATUS" -eq "Gold") -or ($_."VIP STATUS" -eq "VP Gold")) -and (($_."ASSIGNMENT STATUS" -eq "A") -or ($_."ASSIGNMENT STATUS" -eq "F")))} |Sort-Object "EMPLOYEE NUMBER"
$AllMbx = Import-csv "c:\temp\Input-UserMbxUPN.csv" |sort-object "Alias"
Write-Host "Total Number of User Mailboxes: " $Mbx.Count
write-host "MbxCnt,EmpNo,ProcessedCnt,ElapsedTime" -NoNewline

# This is the script to get the mailbox size report for sending of messages for mailboxes with no retention policy.

foreach ($Mbx in $AllMbx)
{
    $stat = get-mailboxstatistics $mbx.PrimarySMTPAddress
    $Created = $mbx.WhenMailboxCreated -replace (" PM","")
    $Created = $mbx.WhenMailboxCreated -replace (" AM","")
    $Size = $stat.TotalItemSize -replace (" bytes","")
    $Warning = $mbx.IssueWarningQuota.Substring(0,$mbx.IssueWarningQuota.IndexOf("("))
    $ProSend = $mbx.ProhibitSendQuota.Substring(0,$mbx.ProhibitSendQuota.IndexOf("("))
    $ProSendRec = $mbx.ProhibitSendReceiveQuota.Substring(0,$mbx.ProhibitSendReceiveQuota.IndexOf("("))
    #    $RetHoldDate = $mbx.StartDateForRetentionHold.Substring(0,$mbx.StartDateForRetentionHold.IndexOf("("))
    $ReceiveSize = $mbx.MaxReceiveSize.Substring(0,$mbx.MaxReceiveSize.IndexOf("("))
    $Text = "{0};{1};""{2}"";{3};{4};{5};{6};{7};{8};{9};{10};{11};{12}" -f $mbx.UserprincipalName, $mbx.CustomAttribute2, $mbx.DisplayName, $mbx.PrimarySMTPAddress, $mbx.RecipientTypeDetails, $mbx.LitigationHoldEnabled, $mbx.RetentionPolicy, $Size, $Created, $Warning, $ProSend, $ProSendRec, $ReceiveSize, $mbx.RetentionHoldEnabled
    Out-File -FilePath $RptFile -InputObject $text -Append

    If (($mbx.RetentionHoldEnabled -eq $True) -and ($mbx.CustomAttribute2 -eq "A"))
    {
        If ($SizeGB.'TotalItemSize(GB)' -ge "45")
        {
            If ($oed."Employee Number" -contains $mbx.alias)
            {
                write-host "This is a Gold Employee: " $mbx.alias, $mbx.DisplayName
#                Critical-GoldMsg
            }
            else
            {
                write-host "This is NOT a Gold Employee: " $mbx.alias, $mbx.DisplayName
#                Critical-Msg
            }
        }
        else
        {
            If ($oed."Employee Number" -contains $mbx.alias)
            {
                write-host "This is a Gold Employee: " $mbx.alias, $mbx.DisplayName
#                Large-GoldMsg
            }
            else
            {
                write-host "This is NOT a Gold Employee: " $mbx.alias, $mbx.DisplayName
#                Large-Msg
            }
        }
    }

    #Out-File -FilePath c:\temp\AllUserMbxDetails.csv -InputObject $text -Append

    $processedCount++
}

##############

$from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"
$to = $from 
$message = new-object  System.Net.Mail.MailMessage $from, $to
$message.IsBodyHtml = $true  
$SendTo = "Sandi.Glazebrook@ul.com"
$SendCC = ""
#$SendCC = "Sandi.Glazebrook@ul.com,RJ.Borja@ul.com"

$message.Subject = "Monthly User Mailbox Reporting Process Complete"
#$message.Attachments.Add($attach)
$message.Body = "<p>The Retention Policy Nofication Message has been sent to " + $SentMsg.Count + " mailboxes.<br>"
$message.To.Clear()

$message.Body = $msgfont + $message.body
$message.To.Add($SendTo)
$message.CC.Clear()
#$message.CC.Add($SendCC)
$client.Send($message)

write-host "Done"
stop-transcript