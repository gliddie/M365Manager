Start-Transcript
#$colAddr = Import-csv -Path E:\Automation\Scripts\UserCommunications\RetireSMTPDomains\Input\Input-RetireXXULCOMDomains.csv
$colAddr = Import-csv -Path E:\Automation\Scripts\UserCommunications\RetireSMTPDomains\Input\Input-RetireLegacyDomains.csv
write-host "Sending Retire Address Message to" $colAddr.count "accounts" -ForegroundColor Yellow
$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server 

$from = New-Object System.Net.Mail.MailAddress "UL.Technology.Services@ul.com" , "UL Technology Services"           
$to = $from 
$replyto = "EnterpriseMessagingServices@ul.com"
$message = new-object  System.Net.Mail.MailMessage $from, $to
$message.IsBodyHtml = $true

$msgfont = "<basefont face=verdana size=2.5 color=black>"

#$message.Body = $html

foreach ($colAddr in $colAddr)
{
    $NoMsg = "N"
    $html2 = ""

    if ($colAddr.Org -eq "UL")
    {
	    if ($colAddr.Type -like "*box")
	    {
		    $MbxExists = [bool](get-mailbox $colAddr.Alias -ErrorAction SilentlyContinue)
            If ($MbxExists -eq "True")
            {
                $Mbx = get-mailbox $colAddr.Alias
		        $SendTo = $Mbx.PrimarySMTPAddress
                if ($colAddr.Type -eq "Mailbox")
                {
                    $message.Subject = "Final Notice:  Action required! Your legacy email address to be retired"
#                    $message.Subject = "Reminder:  Action required! Your legacy email address to be retired"
                    $html1 = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\RetireSMTPDomains\LegacyXXULCOMAddress1.html
                }
                else
                {
                    $message.Subject = "Final Notice:  Action required! Legacy email address to be retired"
#                    $message.Subject = "Reminder:  Action required! Legacy email address to be retired"
                    $html1 = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\RetireSMTPDomains\LegacyXXULCOMSharedAddress1.html
                }
            }
            else
            {
                write-host "Mailbox Not found for " $ColAddr.Alias
                $NoMsg = "Y"
            }
	    }
	    else
	    {
            if ($ColAddr.Type -eq "DistributionGroup")
            {
                $DLExists = [bool](get-distributiongroup $colAddr.Alias -ErrorAction SilentlyContinue)
                If ($DLExists -eq "True")
                {
                    $DL = get-distributiongroup $colAddr.Alias
                    if (($DL.ManagedBy -ne "") -and ($DL.ManagedBy -ne "MBX.DistGroup.Owner"))
                    {
                        $SendTo = ""
                        $Cnt=0
                        foreach ($Owner in $DL.ManagedBy)
                        {                
                            $DLOwner = get-mailbox $DL.ManagedBy[$cnt]
                            If (($cnt -eq 0) -and ($DLOwner.PrimarySmtpAddress -ne $null))
                            {
                                $SendTo = $DLOwner.PrimarySMTPAddress
                            }
                            else
                            {
                                if (($cnt -gt 0) -and ($DLOwner.PrimarySmtpAddress -ne $null))
                                {
                                    $SendTo = $SendTo + "," + $DLOwner.PrimarySMTPAddress
                                }
                            }
                            $cnt = $cnt+1
                        }
                        $message.Subject = "Final Notice:  Action required! Legacy email address to be retired"
#                        $message.Subject = "Reminder:  Action required! Legacy email address to be retired"
                        $html1 = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\RetireSMTPDomains\LegacyXXULCOMDLAddress1.html
                    }
                    else
                    {
                        $SendTo = "Sandi.Glazebrook@ul.com"
                        $message.Subject = "Final Notice:  Distribution List with No Managers"
#                        $message.Subject = "Reminder:  Distribution List with No Managers"
                        $html1 = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\RetireSMTPDomains\NoOwnerDLAddress1.html
                    }
                }
                else
                {
                    write-host "Distribution List Not found for " $ColAddr.Alias
                    $NoMsg = "Y"
                }
            }
            else
            {
                if ($colAddr.Type -eq "MailContact")
                {
                    $MCExists = [bool](get-mailcontact $colAddr.Alias -ErrorAction SilentlyContinue)
                    If ($MCExists -eq "True")
                    {
                        $Mbx = get-mailcontact $colAddr.Alias
                        $SendTo = $Mbx.PrimarySMTPAddress
#                        $SendTo = "Sandi.Glazebrook@ul.com"
                        $message.Subject = "Final Notice:  Action required! Your Legacy email address to be retired"
#                        $message.Subject = "Reminder:  Action required! Your Legacy email address to be retired"
                        $html1 = "You are receiving this message because your UL email account which forwards email to your " + ((get-mailcontact $colAddr.Alias).ExternalEmailAddress).Split(":")[1] + " email address currently contains a legacy email address of " + $colAddr.LegacyAddr + " that was used before UL migrated its email system to Outlook in 2011. This legacy email address includes a two-letter country code as a part of the email address.<br><br>"
                    }
                    else
                    {
                        write-host "Mail Contact Not found for " $ColAddr.Alias
                        $NoMsg = "Y"
                    }
                }
            }
	    }
    }
       
    else

    {

        if ($colAddr.Org -eq "MA")
        {
            if ($colAddr.Type -like "*box")
	        {
                $MbxExists = [bool](get-mailbox $colAddr.Alias -ErrorAction SilentlyContinue)
                If ($MbxExists -eq "True")
                {
		            $Mbx = get-mailbox $colAddr.Alias
		            $SendTo = $Mbx.PrimarySMTPAddress
                    if ($colAddr.Type -eq "Mailbox")
                    {
                        $message.Subject = "Final Notice:  Action required! Your legacy email address to be retired"
#                        $message.Subject = "Reminder:  Action required! Your legacy email address to be retired"
                        $html1 = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\RetireSMTPDomains\LegacyMAAddress1.html
                    }
                    else
                    {
                        $message.Subject = "Final Notice:  Action required! Legacy email address to be retired"
#                        $message.Subject = "Reminder:  Action required! Legacy email address to be retired"
                        $html1 = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\RetireSMTPDomains\LegacyMASharedAddress1.html
                    }
                }
                else
                {
                    write-host "No Mailbox found for " $ColAddr.Alias
                    $NoMsg = "Y"
                }
	        }
	        else
	        {
                if ($ColAddr.Type -eq "DistributionGroup")
                {
                    $DLExists = [bool](get-distributiongroup $colAddr.Alias -ErrorAction SilentlyContinue)
                    If ($DLExists -eq "True")
                    {
                        $DL = get-distributiongroup $colAddr.Alias
                        if (($DL.ManagedBy -ne "") -and ($DL.ManagedBy -ne "MBX.DistGroup.Owner"))
                        {
                            $SendTo = ""
                            $Cnt=0
                            foreach ($Owner in $DL.ManagedBy)
                            {                
                                $DLOwner = get-mailbox $DL.ManagedBy[$cnt]
                                If (($cnt -eq 0) -and ($DLOwner.PrimarySmtpAddress -ne $null))
                                {
                                    $SendTo = $DLOwner.PrimarySMTPAddress
                                }
                                else
                                {
                                    if (($cnt -gt 0) -and ($DLOwner.PrimarySmtpAddress -ne $null))
                                    {
                                        $SendTo = $SendTo + "," + $DLOwner.PrimarySMTPAddress
                                    }
                                }
                                $cnt = $cnt+1
                            }
                            $message.Subject = "Final Notice:  Action required! Legacy email address to be retired"
#                            $message.Subject = "Reminder:  Action required! Legacy email address to be retired"
                            $html1 = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\RetireSMTPDomains\LegacyMADLAddress1.html
                        }
                        else
                        {
                            $SendTo = "Sandi.Glazebrook@ul.com"
                            $message.Subject = "Final Notice:  Distribution List with No Managers"
#                            $message.Subject = "Reminder:  Distribution List with No Managers"
                            $html1 = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\RetireSMTPDomains\NoOwnerDLAddress1.html
                        }
                    }
                    else
                    {
                        write-host "No Distribution List found for " $colAddr.Alias
                        $NoMsg = "Y"
                    }
                }
	        }
        }
    }

    If ($NoMsg -eq "N")
    {

        if ($ColAddr.Type -like "*box*")
        {
            $inputline = "Effective Tuesday, November 1, the email address " + $colAddr.LegacyAddr + " will be removed from the UL system. After this date, customers or other external parties who have your legacy email address saved in their address books will begin to receive undeliverable notifications. Any messages sent to this legacy email address will not be delivered starting on November 1.<br>"
            $html2 = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\RetireSMTPDomains\LegacyAddress2.html
        }
        else
        {
            if ($ColAddr.Type -eq "MailContact")
            {
                $inputline = "Effective Tuesday, November 1, the email address " + $colAddr.LegacyAddr + " will be removed from the UL system. After this date, customers or other external parties who have your legacy email address saved in their address books will begin to receive undeliverable notifications. Any messages sent to this legacy email address will not be delivered starting on November 1.  It is your responsibility to notify individuals of your correct @ul.com email address.<br>"
                $html2 = Get-Content -LiteralPath e:\Automation\Scripts\UserCommunications\RetireSMTPDomains\LegacyAddress2Contacts.html
            }
            else
            {
#                If ($message.Subject -eq "Reminder:  Distribution List with No Managers")
                If ($message.Subject -eq "Final Notice:  Distribution List with No Managers")
                {
                    $inputline = ""
                }
                else
                {
                    $inputline = "Effective Tuesday, November 1, the email address " + $colAddr.LegacyAddr + " will be removed from the UL system. After this date, customers or other external parties who send to this legacy email address will begin to receive undeliverable notifications. As an owner of the distribution list it is your responsibility to notify individuals or to coordinate changes to applications that may be using this address.  The address that should be used to send to this list is " + $DL.PrimarySmtpAddress + ".<br><br>"
                }
                $html2 = "This distribution list can be found in the <b>Global Address List</b> under <b>" + $DL.DisplayName + "</b>.  The number of members in this list is " + (Get-DistributionGroupMember $colAddr.Alias).count + ".  The last time this group was updated is " + (Get-DistributionGroup $colAddr.Alias).WhenChanged + ".<br>"
            }
        }
	
	    $message.To.Clear()
        $message.Body = $msgfont + $html1 + $inputline + $html2
        $message.To.Add($SendTo)
        $message.ReplyTo=$ReplyTo
	    $client.Send($message)
        write-host "Message sent to " $colAddr.Type,$colAddr.Alias,$colAddr.LegacyAddr

    }
}
stop-transcript