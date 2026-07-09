#Start-Transcript
$colAddr = Import-csv -Path E:\Automation\Scripts\UserCommunications\LargeMailbox\Input\Input-LargeMailboxNotification.csv
write-host "Sending Large Mailbox Message to" $colAddr.count
$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server 

$from = New-Object System.Net.Mail.MailAddress "UL.Technology.Services@ul.com" , "UL Technology Services"           
$to = $from 

$message = new-object  System.Net.Mail.MailMessage $from, $to 
$message.IsBodyHtml = $true

$msgfont = "<basefont face=verdana size=2.5 color=black>"

#$message.Body = $html

foreach ($colAddr in $colAddr)
{
	if ($colAddr.UPN -notlike "*global*")
	{
		$Mbx = get-mailbox $colAddr.UPN
		$SendTo = $Mbx.PrimarySMTPAddress
	}
	else
	{
		$SendTo = $colAddr.UPN
	}
	
	$Mbxtype = "your"
	if ($colAddr.Type -like "*Shared*")
	{
		$MbxType = "this shared"
	}
	
    if ($colAddr.Size -lt 45)
    {
        $message.Subject = "Action Required: Mailbox Exceeds 30 GB"
        $html = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\LargerThan30GBFinal.html 

	    $message.To.Clear()
	    $frstline = "You are receiving this message because " + $MbxType + " mailbox has reached <b>" + $colAddr.Size + " GB</b>. While you have not exceeded the maximum file size, having a large mailbox can cause:<br>"
    }
    else
    {
        $message.Subject = "Action Required - Mailbox Quota Exceeds 45 GB"
        $html = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\LargerThan45GBFinal.html
        
        $message.To.Clear()
        $frstline = "<p>You are receiving this message because " + $MbxType + " mailbox has reached <b>" + $colAddr.Size + 	" GB. <font color=red> Once your mailbox reaches 50 GB in size, you will not be able to send or receive email messages until your mailbox size is reduced.</font></b></p>"
    }

	$message.Body = $msgfont + $frstline + $html
    $message.To.Add($SendTo)
	$client.Send($message)
}
#stop-transcript