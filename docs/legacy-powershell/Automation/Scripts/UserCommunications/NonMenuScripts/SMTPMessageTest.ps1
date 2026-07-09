    $server = "smtp-relay.ul.com"
	$client = new-object system.net.mail.smtpclient $server
    $from = New-Object System.Net.Mail.MailAddress "EnterpriseMesagingServices@ul.com" , "Enterprise Messaging Services"
	$to = $from 
	#$SendTo ="ChuckG@renerofe.com"
	#
	$SendTo = "Sandi.Glazebrook@ul.com"
	$message = new-object  System.Net.Mail.MailMessage $from, $to 
	$message.IsBodyHtml = $true
	$msgfont = "<basefont face=verdana size=2.5 color=black>"
	$message.Subject = "Test Message"
	$message.To.Clear()
	$message.Body = $msgfont + "<p>Hello this is a test message please confirm if you receive this message.<M/p>"
	$message.To.Add($SendTo)
	$client.Send($message)