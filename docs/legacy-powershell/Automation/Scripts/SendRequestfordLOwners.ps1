$Plus30 = (Get-Date).AddDays(30)
$RespDate = (Get-Culture).DateTimeFormat.GetMonthName($Plus30.Month) + " " + $Plus30.Day + ", " + $Plus30.Year
write-host "Group Name: " $Script:txtDLName.text
$SendTo = (Get-DistributionGroupMember $Script:txtDLName.Text) -join (";")

$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server 
$from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"
$to = $from 
$message = new-object  System.Net.Mail.MailMessage $from, $to 
$message.IsBodyHtml = $true
$msgfont = "<basefont face=verdana size=2.5 color=black>"

$message.Subject = "Action Required:  Distribution List Owner Needed for: " + $Script:txtDLName.Text

$message.body = "<p>You are receiving this message because the distribution group " + $Script:txtDLName.Text + " which you are a member of no longer any owners. It is required that all distribution lists must have at least 1 (but not more than 4) individuals business identified as an owner of the group.  Group ownership gives those individuals the permissions to manage the membership of the the group.</p>"
$message.body = $message.body + "<p>If no requests are received to configured a new group owner by " + $RespDate + " this group will be removed from the system. Purged groups cannot be restored therefore if the group is needed you will need to submit a New Distrubiton List request via the Service Desk portal to have a new group created.cv</p>"
$message.body = $message.body + "<p>Thanks,<br>Enterprise Messaging Services Team"

$message.To.Clear()
$message.Body = $msgfont + $message.body
$message.To.Add("Sandi.Glazebrook@ul.com")
#$message.To.Add($SendTo)
write-host $message
$client.Send($message)