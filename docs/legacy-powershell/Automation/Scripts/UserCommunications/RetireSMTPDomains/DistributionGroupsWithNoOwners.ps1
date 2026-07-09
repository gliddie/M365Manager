#Start-Transcript
write-host "Reviewing all Distribtuion Groups and Reporting on Groups with No Owners" -ForegroundColor Yellow
$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server 

$from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"           
$to = $from 

$message = new-object  System.Net.Mail.MailMessage $from, $to 
$message.IsBodyHtml = $true

$msgfont = "<basefont face=verdana size=2.5 color=black>"

#$message.Body = $html

$DLs = Get-DistributionGroup -ResultSize Unlimited | Where-Object {$_.ManagedBy[0] -eq $null}

write-host "Number of Distribution Group Discovered " $DLs.count

$html2 =""
$SendTo = "Sandi.Glazebrook@ul.com"
$message.Subject = "Distribution Lists with No Managers"
$html1 = "The following Distribution Lists have No Active Owners.<br><br>"
write-host "generating first report"

foreach ($DLs in $DLs)
{
    $html2 = $html2 + "The distribution list " + $DLs.Name + " has " + (Get-DistributionGroupMember $DLs.Alias -Resultsize Unlimited).count + " members and was last updated " + (Get-DistributionGroup $DLs.Alias).WhenChanged + ".<br><br>"
}

$message.To.Clear()
$message.Body = $msgfont + $html1 + $html2
$message.To.Add($SendTo)
$client.Send($message)


$html2 = ""
$html1 = "The following Distribution Lists are owned by MBX.DistGroup.Owner.<br><br>"


$DLs = Get-DistributionGroup -ResultSize Unlimited |Where-Object {$_.ManagedBy[0] -eq "MBX.DistGroup.Owner"}
write-host "Number of Distribution Group Discovered " $DLs.count

foreach ($DLs in $DLs)
{
     $html2 = $html2 + "The distribution list " + $DLs.Name + " has " + (Get-DistributionGroupMember $DLs.Alias).count + "members and was last updated " + (Get-DistributionGroup $DLs.Alias).WhenChanged + ".<br><br>"
} 
$message.To.Clear()
$message.Body = $msgfont + $html1 + $html2
$message.To.Add($SendTo)
$client.Send($message)
#stop-transcript