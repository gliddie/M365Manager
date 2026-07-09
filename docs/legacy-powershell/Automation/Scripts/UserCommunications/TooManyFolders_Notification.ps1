#Start-Transcript
$colAddr = Import-csv -Path E:\Automation\Scripts\UserCommunications\TooManyFolders\Input\Input-TooManyFoldersNotification.csv
$coladdr.count
$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server 

$from = New-Object System.Net.Mail.MailAddress "UL.Technology.Services@ul.com" , "UL Technology Services"           
$to = $from 

$message = new-object  System.Net.Mail.MailMessage $from, $to 
$message.IsBodyHtml = $true

$message.Subject = "Action Required: Mailbox Folder Limit Exceeded" 
$html = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\TooManyFoldersFinal.html
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

	if ($colAddr.Type -eq "Shared")
	{
		$MbxType = "this shared"
	}
	
	$message.To.Clear()
	$frstline = "You are receiving this message because " + $MbxType + " mailbox contains <b>" + $colAddr.NoFolders + "</b> user-created folders. Having such a large number of folders in your mailbox can cause:<br>"
	$message.Body = $msgfont + $frstline + $html
	$message.To.Add($SendTo)
	$client.Send($message)
}
#stop-transcript