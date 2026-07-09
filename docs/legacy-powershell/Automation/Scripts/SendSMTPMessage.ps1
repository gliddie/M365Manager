####
#### Send SMTP Message
####
#################################################################################
# 
# PowerShell source code
# Revision v1.05
# ==========================================================================
#    'Project      : O365 Admin Team Menu Creation
#    'Description  : O365 Administration Main Menu
#    'Called By    : Check_LegalHold.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 04/28/2018
#    'Comments     :
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '             : SAG - 04/28/2018 Created to Use when permissions to User mailbox are already set for account placed on legal hold
#
#################################################################################>

    $server = "smtp-relay.ul.com"
    $client = new-object system.net.mail.smtpclient $server
    $from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Message Services Team"
    $to = New-Object System.Net.Mail.MailAddress $SendTo
    $message = new-object  System.Net.Mail.MailMessage $from, $to 
    $message.IsBodyHtml = $true
    $msgfont = "<basefont face=verdana size=2.5 color=black>"
    #$message.To.Add("EnterpriseMessagingServices@ul.com")
    $message.To.Clear()
    $message.CC.Clear()
    #$message.To.Add("EnterpriseMessagingServices@ul.com")
    $Message.TO.Add($SendTo)
    If ($SendCC -ne "")
    {
        $Message.CC.Add($SendTo)
    }
    $message.Subject = $MessageSubject
    $message.Body = $msgfont + $MessageBody
    $client.Send($message)