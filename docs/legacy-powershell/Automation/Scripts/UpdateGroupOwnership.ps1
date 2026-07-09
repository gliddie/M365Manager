<#
================================================================================ 
 Update Distribution/Security/O365 Group ownership
 ================================================================================ 
 #
 #  10/25/2022 - SAG:Create menu process for the original inline question process
 #  11/02/2022 - Moved the Build-DefaultForm, Get-DLGrpDetails, Get-UniGrpDetails, Hide-Details, UnHide-Details and Write-GrpDetails functions to the DistGroupMenu script
#>

Build-DefaultForm
$Global:form.Text = "Update Group Ownership" 
Add-ActionBoxes
Add-FormStandardButtons
Hide-Details
Publish-Form

############################################################
#Start Of Script

write-host "Updating Distribution/Security/O365 Group Ownership " -ForegroundColor Magenta
write-host
$ResetTip = "N"
$WipeOwners = "N"

If ($Script:chkReqOwner.Checked -eq $True)
{
    $ReportPath = "E:\Automation\RequestOwner\Report\" + $Year + "\"
    $Script:txtRptFile.Text = $ReportPath + "Report-RequestOwner-" + ($DLGrp.Alias -replace("[.]",""))  + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
}
else
{
    $ReportPath = "E:\Automation\UpdateGroupOwnership\Report\" + $Year + "\"
    $Script:txtRptFile.Text = $ReportPath + "Report-UpdateGroupOwnership-"+ ($Script:txtDLName.Text -replace("[.]",""))  + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
}
#If (Test-Path $ReportPath) {} else {New-Item -Path $ReportPath -ItemType Directory}
If (Test-Path $ReportPath) {} else {mkdir $ReportPath | out-Null}

If ($Global:Result -eq "OK")
{
    $ReportFile = $Script:txtRptFile.Text

    If ($Script:chkReqOwner.Checked -eq $True)
    {
        $LineToWrite = $WhoAmI + "`t" + "Distribtuion Request for Owner Has Started"
    }
    else
    {
        $LineToWrite = $WhoAmI + "`t" + "Distribtuion Group Update Has Started"
    }
    WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "Group Details Before Changes"
    WriteReportEvent

    Write-GroupDetails

    If ($Script:chkReqOwner.Checked -eq $True)
    {
    #Send message requesting owners
        If ($Script:txtCurrMbrCnt.Text -le 50)
        {
            $Plus30 = (Get-Date).AddDays(30)
            $RespDate = (Get-Culture).DateTimeFormat.GetMonthName($Plus30.Month) + " " + $Plus30.Day + ", " + $Plus30.Year
            $SendTo = (Get-DistributionGroupMember $Script:txtDLName.Text) -join (";")

            $server = "smtp-relay.ul.com"
            $client = new-object system.net.mail.smtpclient $server 
            $from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"
            $to = $from 
            $message = new-object  System.Net.Mail.MailMessage $from, $to 
            $message.IsBodyHtml = $true
            $msgfont = "<basefont face=verdana size=2.5 color=black>"

            $message.Subject = "Action Required:  Distribution List Owner Needed for: " + $Script:txtDLName.Text

            $reqDLOwnerLink = "https://ul.service-now.com/com.glideapp.servicecatalog_cat_item_view.do?v=1&sysparm_id=938c413908538d001f89f039672b8a39&sysparm_link_parent=bb72be7e130a9b00358550782244b03f&sysparm_catalog=e0d08b13c3330100c8b837659bba8fb4&sysparm_catalog_view=catalog_default&sysparm_view=catalog_default"
            $reqNewDLGrpLink = "https://ul.service-now.com/com.glideapp.servicecatalog_cat_item_view.do?v=1&sysparm_id=707fd4c0a996e40074107e119f5c73ee&sysparm_link_parent=bb72be7e130a9b00358550782244b03f&sysparm_catalog=e0d08b13c3330100c8b837659bba8fb4&sysparm_catalog_view=catalog_default&sysparm_view=catalog_default"
            $reqRemoveDLLink = "https://ul.service-now.com/com.glideapp.servicecatalog_cat_item_view.do?v=1&sysparm_id=f31fdc9df4a961001f89b7ccb50f8a9f&sysparm_link_parent=bb72be7e130a9b00358550782244b03f&sysparm_catalog=e0d08b13c3330100c8b837659bba8fb4&sysparm_catalog_view=catalog_default&sysparm_view=catalog_default"

            $message.body = "<p>You are receiving this message because the distribution group <b>" + $Script:txtDLName.Text + "</b> which you are a member of, no longer has any owners. All distribution lists must have at least 1 (but not more than 4) individuals from the business identified as an owner of the group.  Group ownership gives those individuals the permissions to manage the membership of the the group.</p>"
            $message.body = $message.body + "<p>If this group is no longer need please use the <a href=$reqRemoveDLLink>Remove Shared Mailboxes, Distribution Lists, Resources or Rooms</a> form to request removal of the group.</p>"
            $message.body = $message.body + "<p>Please use the <a href=$reqDLOwnerLink>Change Distribution List Owner(s)</a> form to request changes to the ownership.  In the field requesting the <i>Name of the Distribution List</i> enter <b>" + $Script:txtDLName.Text + "</b> select <b>No</b> that you <i>are not a current owner</i>, and in the <i>Reason for Change</i> select <b>Owner is no longer with organization</b>.</p>"
            $message.body = $message.body + "<p>No additional requests will be sent to obtain a new owner for this group.  If no service desk ticket requests are received to remove the group or to configure a new group owner by <b>$RespDate</b> this group will be purged from the system. Purged groups cannot be restored, therefore if the group is needed a <a href=$reqNewDLGrpLink>New Distrubiton List</a> request must be submitted via the Service Desk portal.</p>"
            $message.body = $message.body + "<p>Thanks,<br>Enterprise Messaging Services Team"

            $message.To.Clear()
            $message.Body = $msgfont + $message.body
            $message.To.Add("Sandi.Glazebrook@ul.com")
            #$message.To.Add($SendTo)
            write-host $message
            $client.Send($message)

            #Update group tip to indicate a message has been sent requesting new groups owners.
            set-distributiongroup $Script:txtDLName.Text -MailTip "Request for Owners sent group will be removed on: " $Plus30.ToString("yyyy/MM/dd")
            set-group $Script:txtDLName.Text -Notes "Request for Owners sent group will be removed on: " $Plus30.ToString("yyyy/MM/dd")

        }
    }
    else
    {

    If ($Script:chkRemOwner.Checked -eq $True)
    {
#       Remove Designated Owners
        $Owner = $Script:txtRemOwner.text -split (",")
        Foreach ($Owner in $Owner)
        {
            $Usr = $Owner.Trim()
            $Exists = [bool]($mbx = get-mailbox $Usr -ErrorAction SilentlyContinue)
            If ($Exists -eq $True)
            {
                If ($Script:DLExists -eq $True)
                {
                    $IsOwner = [bool]((Get-DistributionGroup $Script:txtDLAddr.Text).ManagedBy -like $mbx.Name)
                    If ($IsOwner -eq $True)
                    {
                        write-host "Removing distribution/security group ownership from: " $Usr -ForegroundColor Red
                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing distribution/security group ownership from: " + $Usr
                        WriteReportEvent
                        Set-DistributionGroup $Script:txtDLAddr.Text -ManagedBy @{remove=$mbx.Alias} -BypassSecurityGroupManagerCheck -confirm:$False
                        $ResetTip = "Y"
                    }
                    else
                    {
                        write-host "Individual is not a current distribution/security owner: " $Usr
                        $LineToWrite = $RecordEvent + "ERRO" + "`t" + "Individual is not a current distribution/security  owner: " + $Usr
                        WriteReportEvent
                    }
                }

                If ($Script:UniExists -eq $True)
                {
                    $IsOwner = [bool](((Get-UnifiedGroup $Script:txtDLAddr.Text).ManagedBy) -like $mbx.Name)
                    If ($IsOwner -eq $True)
                    {
                        write-host "Removing unified group ownership from: " $usr -ForegroundColor Red
                        Remove-UnifiedGroupLinks $Script:txtDLAddr.Text -LinkType Owners -Links $mbx.Alias -Confirm:$False
                        Remove-UnifiedGroupLinks $Script:txtDLAddr.Text -LinkType Member -Links $mbx.Alias -Confirm:$False
                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing unified group ownership from: " + $Usr
                        WriteReportEvent
                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing unified group membership from: " + $Usr
                        WriteReportEvent
                        $ResetTip = "Y"
                    }
                    else
                    {
                        write-host "Individual is not a current unified group owner: " $Usr
                        $LineToWrite = $RecordEvent + "ERRO" + "`t" + "Individual is not a current unified group owner: " + $Usr
                        WriteReportEvent
                    }
                }
            }
        }
    }

    If (($Script:chkAddOwner.Checked -eq $True) -or ($Script:chkReplOwner.Checked -eq $True))
    {
#       Add Owners to the Group
        $Owner = $Script:txtNewOwner.text -split (",")
        Foreach ($Owner in $Owner)
        {
            $Usr = $Owner.Trim()
            $Exists = [bool]($mbx = get-mailbox $Usr -ErrorAction SilentlyContinue)
            If ($Exists -eq $True)
            {
                If ($Script:DLExists -eq $True)
                {
                    $IsOwner = [bool]((Get-DistributionGroup $Script:txtDLAddr.Text).MmanagedBy -like $mbx.Name)
                    If ($IsOwner -eq $False)
                    {
                        If (($Script:chkReplOwner.Checked -eq $True) -and ($WipeOwners -eq "N"))
                        {
#                           Replaces Current Owners with first new owner
                            Write-Host "Replacing current owners" -ForegroundColor Red
                            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Replacing all distribution/security group owners"
                            WriteReportEvent
                            write-host "Adding distribution/security group ownership for: " $Usr -ForegroundColor Green
                            $LineToWrite = $RecordEvent + "ADD " + "`t" + "Adding distribution/security group ownership for: " + $Usr
                            WriteReportEvent
                            Set-DistributionGroup $Script:txtDLAddr.Text -ManagedBy $mbx.Alias
                            $WipeOwners = "Y"
                        }
                        else
                        {
#                           Adding Additional Owners
                            write-host "Adding distribution/security group ownership to: " $Usr -ForegroundColor Green
                            Set-DistributionGroup $Script:txtDLAddr.Text -ManagedBy @{add=$mbx.Alias} -BypassSecurityGroupManagerCheck -confirm:$False
                        }
                        $ResetTip = "Y"
                    }
                }

                If ($Script:UniExists -eq $True)
                {

#                   Add New Owners
                    Add-UnifiedGroupLinks $Script:txtDLAddr.Text -LinkType Member -Links $mbx.Alias -Confirm:$False -ErrorAction SilentlyContinue
                    Add-UnifiedGroupLinks $Script:txtDLAddr.Text -LinkType Owners -Links $mbx.Alias -Confirm:$False -ErrorAction SilentlyContinue
                    $LineToWrite = $RecordEvent + "ADD " + "`t" + "Adding unified group membership for: " + $Usr
                    WriteReportEvent
                    $LineToWrite = $RecordEvent + "ADD " + "`t" + "Adding unified group ownership for: " + $Usr
                    WriteReportEvent
                    $ResetTip = "Y"

                    If (($Script:chkReplOwner.Checked -eq $True) -and ($WipeOwners -eq "N") -and ($Script:txtCurrOwner.Text.Length -gt 10))
                    {
#                       Remove All Existing Owners
                        $OldOwner = $Script:txtCurrOwner.Text -split ", "
                        Foreach ($OOwner in $OldOwner)
                        {
                            
                            $OOwner = $OOwner -replace ("(Terminated)","")
                            $mbx = get-mailbox $OOwner
                            Remove-UnifiedGroupLinks $Script:txtDLAddr.Text -LinkType Owners -Links $mbx.Alias -Confirm:$False
                            Remove-UnifiedGroupLinks $Script:txtDLAddr.Text -LinkType Member -Links $mbx.alias -Confirm:$False
                            $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing unified group ownership from: " + $OldOwner
                            WriteReportEvent
                            $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing unified group membership from: " + $OldOwner
                            WriteReportEvent
                            $WipeOwners = "Y"
                        }
                    }
                }
            }
            else
            {
                write-host "This user does not have a licensed mailbox: " $Usr -ForegroundColor Red
                $LineToWrite = $RecordEvent + "ERRO" + "`t" + "This user does not have a licensed mailbox for: " + $Usr
                WriteReportEvent
            }
        }
    }

    If ($ResetTip -eq "Y")
    {
        $Global:Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
        $Global:Form.BackColor = "LightBlue"
        $Global:form.Text = "Update Group Ownership - After Changes" 
        Write-Host "Resetting MailTip" -ForegroundColor Green
        If ($Script:DLExists -eq $True)
        {
#            $Own = (Get-DistributionGroup $Script:txtDLAddr.Text).ManagedBy
            $Global:GrpName = $Script:txtDLAddr.Text
            Build-OwnerMailTip
#            CreateMailTip ($Script:txtDLAddr.Text)
            $MailTip = $Script:Owners
            $GrpTip = "Owners: " + $Script:Owners + " - Per: " + $Script:txtTicketNo.Text
            Set-DistributionGroup $Script:txtDLAddr.Text -MailTip $Script:Owners
            Set-Group $Script:txtDLAddr.Text -Notes $GrpTip
            $Script:txtCurrOwner.Text = $Script:Owners.TrimStart("Owners: ")
#            $Script:txtCurrOwner.Text = ((Get-DistributionGroup $Script:txtDLAddr.Text).ManagedBy -Join ", ")
            $Script:txtCurrMbr.Text = ((Get-DistributionGroupMember $Script:txtDLAddr.Text) -Join ", ")
            write-host "Resetting the distribution/security group MailTip for: " $Script:txtDLAddr.Text -ForegroundColor Green
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Resetting the distribution/security group MailTip for: " + $Script:txtDLAddr.Text
            WriteReportEvent
            write-host "Resetting the distribution/security group Notes for: " $Script:txtDLAddr.Text -ForegroundColor Green
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Resetting the distribution/security group Notes for: " + $Script:txtDLAddr.Text + "`n"
            WriteReportEvent
        }

        If ($Script:UniExists -eq $True)
        {
#            $Own = (Get-UnifiedGroup $Script:txtDLAddr.Text).ManagedBy
            $Global:GrpName = $Script:txtDLAddr.Text
            Build-OwnerMailTip
            $GrpTip = $Script:Owners + " - Per: " + $Script:txtTicketNo.Text
            $MailTip = $Script:Owners + " - Per: " + $Script:txtTicketNo.Text
            Set-UnifiedGroup $Script:txtDLAddr.Text -MailTip $MailTip
            $Script:txtCurrOwner.Text = $Script:Owners.TrimStart("Owners: ")
            $Script:txtCurrMbr.Text = (Get-UnifiedGroupLinks $Script:txtDLAddr.Text -LinkType Member) -join ", "
            write-host "Resetting the unified group MailTip for: " $Script:txtDLAddr.Text -ForegroundColor Green
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Resetting the unified group MailTip for: " + $Script:txtDLAddr.Text + "`n"
            WriteReportEvent
        }

        #Display the group details after requested changes have been made
        $Script:txtMailTip.Text = $MailTip
        $Global:Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
        $Global:Form.BackColor = "LightBlue"
        $Script:txtTicketNo.BackColor = "LightGray"
        $Script:lblNewOwner.Visible = $False
        $Script:txtNewOwner.Visible = $False
        $Script:lblRemOwner.Visible = $False
        $Script:txtRemOwner.Visible = $False
        $Script:chkAddOwner.Visible = $False
        $Script:chkReplOwner.Visible = $False
        $Script:chkRemOwner.Visible = $False
        $Global:cancelButton.Text = "Exit"
        $Global:okButton.Visible = $False
        Publish-Form
        $LineToWrite = $RecordEvent + "INFO" + "Group Details After Changes"
        WriteReportEvent
        Write-GroupDetails
        Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\DistributionListOwnershipChange.oft
        $Output = $wshell.Popup("Group Ownership Update Request Complete.",10,"Complete",0+32)
        write-host "Group Ownership changes complete" -ForegroundColor Cyan
        $LineToWrite = $RecordEvent + "INFO" + "Group Onwership changes Complete"
        WriteReportEvent
    }
    else
    {
        write-host "MailTip not updated for: " $Script:txtDLAddr.Text -ForegroundColor Red
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + "MailTip not updated for: " + $Script:txtDLAddr.Text + "`n"
        WriteReportEvent
        $Output = $wshell.Popup("MailTip not updated for no owner changes made.",10,"Not Updated",0+32)        
    }

    }
}
else
{
    write-host "Group Ownership Update Request Cancelled" -ForegroundColor Red
    $Output = $wshell.Popup("Group Ownership Update Request Cancelled.",10,"Cancelled",0+32)
}