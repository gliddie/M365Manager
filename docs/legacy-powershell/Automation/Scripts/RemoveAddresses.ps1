$InpFile = Import-Csv e:\automation\scripts\usercommunications\retiresmtpdomains\input\Input-MAAddresses-Final.csv

$from = New-Object System.Net.Mail.MailAddress "UL.Technology.Services@ul.com" , "UL Technology Services"           
$to = $from 
$replyto = "EnterpriseMessagingServices@ul.com"
$message = new-object  System.Net.Mail.MailMessage $from, $to
$message.IsBodyHtml = $true
write-host "Removing Old email addresses from" $InpFile.count "accounts" -ForegroundColor Yellow
$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server 

$msgfont = "<basefont face=verdana size=2.5 color=black>"

foreach ($InpFile in $InpFile)
{
    switch ($InpFile.Type)
    {
        "DistributionGroup" #need to add code to send this to the DL Owners
        {
            #This code removes the legacy addresses from distribution groups
            $DLExists = [bool](get-distributiongroup $InpFile.Alias -ErrorAction SilentlyContinue)

            If ($DLExists -eq "True")
            {
                $DistGroup = Get-DistributionGroup $InpFile.Alias
                If (($DistGroup.ManagedBy -ne "") -and ($DistGroup.ManagedBy -ne "MBX.DistGroup.Owner"))
                {
                    $SendTo = ""
                    $Cnt=0
                    foreach ($Owner in $DistGroup.ManagedBy)
                    {                
                        $DLOwner = get-mailbox $DistGroup.ManagedBy[$cnt]
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
                    $SendTo = $DistGroup.PrimarySMTPAddress
                    Set-DistributionGroup $InpFile.Alias -EmailAddresses @{remove=$InpFile.LegacyAddress}
                    $message.Subject = "Legacy email address has been removed!"
                    $html1 = "You are receiving this message because the legacy email address of " + $InpFile.LegacyAddress + " assigned to the distribution list " + $DistGroup.DisplayName + " has been removed and email sent using the " + $InpFile.LegacyAddress + " address will no longer be delivered to this group.<br><br>"
                    $SendMsg = "Y"
                }
                else
                {
                    $SendTo = "Sandi.Glazebrook@ul.com"
                    $message.Subject = "Legacy Address removed from Distribution List with No Managers"
                    $html1 = "The legacy email address of " + $InpFile.LegacyAddress + " assigned to the distribution list " + $DistGroup.DisplayName + " has been removed from the group.<br><br>"
                    $SendMsg = "Y"
                }
            }
            else
            {
                $SendTo = "Sandi.Glazebrook@ul.com"
                $message.Subject = "Distribution List Not Found"
                $html1 = "The legacy email address of " + $InpFile.LegacyAddress + " could not be removed as the distribution list " + $InpFile.Alias + " does not exist.<br><br>"
                $SendMsg = "Y"
            }
        }
        
        "SharedMailbox"
        {        
            #This code removes the legacy addresses from shared mailboxes

            $MbxExists = [bool](get-mailbox $InpFile.Alias -ErrorAction SilentlyContinue)
            If ($MbxExists -eq "True")
            {
                $ShrMbx = Get-Mailbox $InpFile.Alias
                $SendTo = $ShrMbx.PrimarySMTPAddress
                Set-Mailbox $InpFile.Alias -EmailAddresses @{remove=$InpFile.LegacyAddress}
                $message.Subject = "Legacy email address has been removed!"
                $html1 = "You are receiving this message because the legacy email address of " + $InpFile.LegacyAddress + " assigned to the shared mailbox " + $ShrMbx.DisplayName + " has been removed and email sent using the " + $InpFile.LegacyAddress + "address will no longer be delivered to this mailbox.<br><br>"
                $SendMsg = "Y"
            }
            else
            {
                $SendTo = "Sandi.Glazebrook@ul.com"
                $message.Subject = "Shared Mailbox Not Found"
                $html1 = "The legacy email address of " + $InpFile.LegacyAddress + " could not be removed as this shared mailbox does not exist.<br><br>"
                $SendMsg = "Y"
            }
        }

        "MailContact"
        {
            #This code removes the legacy addresses from the mail contact records

            $MCExists = [bool](get-mailcontact $InpFile.Alias -ErrorAction SilentlyContinue)
            If ($MCExists -eq "True")
            {                
                $CntRec = Get-MailContact $InpFile.Alias
                $SendTo = $CntRec.PrimarySMTPAddress
                Set-MailContact $InpFile.Alias -EmailAddresses @{remove=$InpFile.LegacyAddress}
                $message.Subject = "Legacy email address has been removed!"
                $html1 = "You are receiving this message because the legacy email address of " + $InpFile.LegacyAddress + " assigned to $CntRec.DisplayName has been removed and email sent using the " + $InpFile.LegacyAddress + "address will no longer be delivered to you.<br><br>"
                $SendMsg = "Y"
            }
            else
            {
                $SendTo = "Sandi.Glazebrook@ul.com"
                $message.Subject = "Mail Contact Not Found"
                $html1 = "The legacy email address of " + $InpFile.LegacyAddress + " could not be removed as this mail contact record does not exist.<br><br>"
                $SendMsg = "Y"
            }
        }

        "Mailbox"
        {
            #This code removes the legacy addresses from the active directory account

            $MbxExists = [bool](get-mailbox $InpFile.Alias -ErrorAction SilentlyContinue)
            $ADAcct = (dsquery.exe USER -SAMID $InpFile.Alias)
            If ($MbxExists -eq "True")
            {
                $UsrMbx = Get-Mailbox $InpFile.Alias
                If ($UsrMbx.RecipientTypeDetails -ne "SharedMailbox")
                {
                    $SendTo = $UsrMbx.PrimarySMTPAddress
                    $RemAddr = "smtp:" + $InpFile.LegacyAddress
                    Set-AdUser $InpFile.Alias -Remove @{ProxyAddresses=$RemAddr}
                    $message.Subject = "Legacy email address has been removed!"
                    $html1 = "You are receiving this message because the legacy email address of " + $InpFile.LegacyAddress + " assigned to your mailbox has been removed. Email sent using the " + $InpFile.LegacyAddress + " address will no longer be delivered to you.<br><br>"
                    $SendMsg = "Y"
                }
                else
                {
                    $SendTo = $UsrMbx.PrimarySMTPAddress
                    Set-Mailbox $InpFile.Alias -EmailAddresses @{remove=$InpFile.LegacyAddress}
                    $message.Subject = "Legacy email address has been removed!"
                    $html1 = "You are receiving this message because the legacy email address of " + $InpFile.LegacyAddress + " assigned to your mailbox has been removed. Email sent using the " + $InpFile.LegacyAddress + "address will no longer be delivered to you.<br><br>"
                    $SendMsg = "Y"
                }
                If ($ADAcct -eq $null)
                {
                    $SendTo = "Sandi.Glazebrook@ul.com"
                    $message.Subject = "AD Account Does Not Exist for User but Shared Mailbox Exists"
                    $html1 = "The legacy email address of " + $InpFile.LegacyAddress + " was removed from the Shared Mailbox however this user no longer has an AD account and the Shared Mailbox should be converted to a UserMailbox so it will be deleted.<br><br>"
                    $SendMsg = "Y"
                }
            }
            else
            {
                $SendTo = "Sandi.Glazebrook@ul.com"
                $message.Subject = "User Mailbox Not Found"
                $html1 = "The legacy email address of " + $InpFile.LegacyAddress + " could not be removed as this user mailbox does not be exist.<br><br>"
                $SendMsg = "Y"
            }
        }
    }

    If ($SendMsg -ne "N")
    {
        $message.To.Clear()
        $message.Body = $msgfont + $html1
        $message.To.Add($SendTo)
        $message.ReplyTo=$ReplyTo
	    $client.Send($message)
        write-host "Message sent for " $InpFile.Type,$InpFile.Alias,$InpFile.LegacyAddress

        $SendMsg = "N"
    }
}