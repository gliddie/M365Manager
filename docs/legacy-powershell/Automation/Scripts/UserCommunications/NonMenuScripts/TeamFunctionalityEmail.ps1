#################################################################################
# 
# PowerShell source code
# Revision v1.2
# ==========================================================================
#    'Project      : Deploy Microsoft Teams
#    'Description  : Sends Email for Addition of Teams Functionality
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Created from the NewUnifiedGroup Script by Sandi Glazebrook
#    'Date Created : 7/7/2017 08:00:00 AM
#
#    'Comments     :  GRP.ENK.GoodGuide
#    '               7/7/2017 Image Files are written to: 
#    '                   http://sharepoint.ul.com/Mail/Forms/Thumbnails.aspx
#    '
# ==========================================================================
#
#################################################################################

# =============================================================================================================================================
$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server
$from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"
$to = $from 
$message = new-object  System.Net.Mail.MailMessage $from, $to 
$message.IsBodyHtml = $true
$msgfont = "<basefont face=verdana size=2.5 color=black>"
$tab = "<p style=""margin-left: 40px"">"
$image8 = "http://sharepoint.ul.com/Mail/image025.jpg"
write-Host "Enter the Name of the Group that the Teams Functionality hs been Added For: " -ForegroundColor Green -NoNewline
$GrpName = Read-Host
Remove-UnifiedGroupLinks -Identity $GrpName -LinkType Owners -links "96151@global.ul.com" -Confirm:$False
write-Host "Enter the Request Number: " -ForegroundColor Green -NoNewline
$RITMnO = Read-Host
Write-Host "Enter the Employee Number of the Group Owner: " -ForegroundColor Green -NoNewline
$SendTo = Read-Host
$SendTo = $SendTo + "@global.ul.com"
$MsgSubject = $GrpName  + " Teams Functionality Added - Per:  " + $RITMNo
$frstline = "<p>Per your request the Microsoft Teams functionality has been added to the O365 group named <font color=green>" + $GrpName + "</font> per service desk request number " + $DG.u_requested_item + "</p>"
$svnthline = "<p>You will now see the group in the left hand navigation pane when you enter Microsft Teams.  To add individuals into the Microsoft Teams use the <font color=blue>'Add more people'</font> feature.</p>"
$ninthline = "<p>If you have additional questions or need additional support please contact the Service Desk.</p><p>Thanks,<br>The Enterprise Messaging Team<Align=left></p>"
$MsgBody = $msgfont + $frstline + $svnthline + $tab + "<img src=$image8></p>" + $ninthline
$message.To.Clear()
$message.CC.Clear()
$message.To.Add($SendTo)
$message.Subject = $MsgSubject
$message.Body = $MsgBody
$client.Send($message)

 # =============================================================================================================================================