#####################################################################################
#
#   Execute Purge Account
#
#   11/29/2020 - Created New Script from originals SDAPAdminMenu details$Global:RemoveMbxPerm
#   08/06/2021 - Modified the code for the No Show it account had to be in the Disabled OU added DisabledOU or Account Disabled to the If statement.  Also commented out the Get-CSUser code
#   09/22/2021 - Modified the process for getting the deatils of the Extension assigned to the account due to changes to using Teams vs UM and clearing those values
#   02/22/2022 - Modified code to clear the phone extension
#   07/01/2022 - Added code to document the original account purge date
#   09/28/2022 - Added code to block receipt of new messages for individuals on legal hold
#   10/06/2022 - Made removal of calendar events a function and added this to the clean-up process for accounts on legal hold
#   05/09/2023 - V1.8 Added code to set RetentionHoldEnabled to False if it is set to True.  When it is True then softdeletedmailboxes cannot be permanently purged after 30 days
#
#####################################################################################

#Put in SDAP Admin menu

Function Add-ActionPurgeCheckBoxes
{
    If ($Global:txtMblDevice.Text -gt 0)
    {
        $form.Height = $form.Height + 30
        $Global:SelWipeRequired = "Yes"
        ## Selective Wipe Performed
        $Global:Top = $Global:Top + 20
        $Global:chkSelWipe = New-Object Windows.Forms.checkbox
        $Global:chkSelWipe.Left = 200; $Global:chkSelWipe.Width = 200; $Global:chkSelWipe.Top = $Global:Top
        $Global:chkSelWipe.ForeColor = "Red"
        $Global:chkSelWipe.Text = "Selective Wipe Performed"
        $Global:chkSelWipe.Checked = $false   # set a default value
        $Global:chkSelWipe.TabIndex = $Tab++
        $Global:form.Controls.Add($Global:chkSelWipe)
    }
    
    If (($ADUsr.Enabled -eq $True) -and ($ADUsr.ExtensionAttribute1 -like "Ex*"))
    {
        $form.Height = $form.Height + 30
        $Global:ADEnabledRequired = "Yes"
        ## AD Account Enabled for Terminated User
        $Global:Top = $Global:Top + 20
        $Global:chkADEnabCheck = New-Object Windows.Forms.checkbox
        $Global:chkADEnabCheck.Left = 200; $Global:chkADEnabCheck.Width = 200; $Global:chkADEnabCheck.Top = $Global:Top
        $Global:chkADEnabCheck.ForeColor = "Red"
        $Global:chkADEnabCheck.Text = "AD User is Terminated"
        $Global:chkADEnabCheck.Checked = $false   # set a default value
        $Global:chkADEnabCheck.TabIndex = $Tab++
        $Global:form.Controls.Add($Global:chkADEnabCheck)
    }

    If (($ADUsr.ExtensionAttribute1 -notlike "Ex-*") -and (($ADUsr.DistinguishedName -like "*Disabled*") -or ($ADUsr.Enabled -eq $False)))
    {
        $form.Height = $form.Height + 30
        $Global:ADEnabledRequired = "Yes"
        ## Add No Show Button for Account where the user is no an "Ex" something
        $Global:Top = $Global:Top + 20
        $Global:chkADNoShowCheck = New-Object Windows.Forms.checkbox
        $Global:chkADNoShowCheck.Left = 200; $Global:chkADNoShowCheck.Width = 200; $Global:chkADNoShowCheck.Top = $Global:Top
        $Global:chkADNoShowCheck.ForeColor = "Red"
        $Global:chkADNoShowCheck.Text = "AD User is a No Show"
        $Global:chkADNoShowCheck.Checked = $false   # set a default value
        $Global:chkADNoShowCheck.TabIndex = $Tab++
        $Global:form.Controls.Add($Global:chkADNoShowCheck)
    }
}

Function Rem-CalendarEvents
{
    if ($CalEven.count -gt 0)
    {
        write-host "Issuing command to remove Calendar Events this user is the Organizer of" -ForegroundColor Cyan 
        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing " + $CalEven.count + " Calendar Events this individual Organized"
        WriteReportEvent
        Out-File -filepath $ReportFile -append -noClobber -inputObject $CalEven
        Remove-CalendarEvents -Identity $Global:txtInpEmpNo.Text -CancelOrganizedMeetings -QueryWindowInDays 1825 -Confirm:$False
    }
    else
    {
        Write-Host "This individual did not Organize Any Calendar Events" -ForegroundColor Red
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "This individual did not Organize Any Calendar Events"
        WriteReportEvent
     }
}

Function Remove-PurgeMbxFldrAccess
{
    write-host "Check and remove permissions to mailbox and key folders" -ForegroundColor Yellow
    $FldrTypes = import-csv E:\O365AdminShared\Data\FolderTypes.csv
    #Remove access at top of mailbox
    $Global:MbxPerm = Get-MailboxPermission $Global:txtInpEmpNo.Text | Where {($_.User -notlike "*SELF*")}
    Foreach ($Perm in $Global:MbxPerm)
    {
        Remove-MailboxPermission $Global:txtInpEmpNo.Text -User $Perm.User -AccessRights $Perm.AccessRights -confirm:$False |out-null
        write-host "Removing" $Perm.User "-->" $Perm.AccessRights "mailbox permissions"
        $LineToWrite = "REMV" + "`tRemoving " + $Perm.User + " --> " + $Perm.AccessRights + " from mailbox"
        WriteReportEvent
    }
        
    #Remove access to identified Key folder types
    $MBXFldr = Get-MailboxFolderStatistics $Global:txtInpEmpNo.Text
    write-host "There are " $MBXFldr.count " folders to review access to permissions, only key folders will have the access removed." -ForegroundColor Yellow
    foreach ($Fldr in $MBXFldr)
    {
        foreach ($Type in $FldrTypes)
        {
            If ($Fldr.FolderType -eq $Type.FolderType)
            {
                $Folder = $Global:txtInpEmpNo.Text + ":\"
                If ($Fldr.Foldertype -ne "Root")
                {
                    $Folder = $Global:txtInpEmpNo.Text + ":\" + $Fldr.Name
                }

                $Global:FLDRPerm = Get-MailboxFolderPermission $Folder | Where {($_.User -notlike "*Default*") -and $_.User -notlike "*Anonym*"}
                If ($Global:FLDRPerm.count -ne 0)
                {
                    Foreach ($Fldr1 in $Global:FLDRPerm)
                    {
                        #Remove Perissions
                        Remove-MailboxFolderPermission $Folder -User $Fldr1.User.DisplayName -Confirm:$False |out-null
                        write-host "`tRemoving from" $Folder "folder " $Fldr1.User.DisplayName "-->" $Fldr1.AccessRights "permissions"
                        $LineToWrite = "REMV" + "`t" + "`tRemoving from " + $Folder + " folder " + $Fldr1.User.DisplayName + " --> " + $Fldr1.AccessRights + " permissions"
                        WriteReportEvent
                    }
                }
            }
        }
    }
}

$UsrTimeZone = Get-TimeZone
$Year = (get-date).ToString("yyyy")
$ReportPath = "E:\SDAP\PurgeAccount\Report\" + $Year + "\"
If (Test-Path $ReportPath) {} else {New-Item -Path $ReportPath -ItemType Directory}

$Filename = "PurgeAccount"
$LogFile = "E:\SDAP\PurgeAccount\Log\Log-" + $FileName + ".log"
$Global:DoNotPurge = "NoVal"
$Global:SelWipeRequired = ""
$Global:ADEnabledRequired = ""
$Global:txtMblDevice = ""
$Global:chkSelWipe = ""
$Global:chkADEnabCheck = ""
$DLMember = "None"
$DLOwner = "None"

Enter-EmpNoInputForm
Publish-Form

If ($Global:Result -eq "OK")
{
    $Usr = $Global:txtInpEmpNo.Text
    $NoOwnerTip = "Message sent to obtain group owners will be deleted if no response by " + (Get-date).AddDays(30).ToString('yyyy-MMM-dd') + " - Per " + $Global:txtInpTicketNo.Text
    If ($usr.length -le 7)
    {
        $ADExists = [bool](get-ADUser -Filter {SamAccountName -eq $Usr} -ErrorAction SilentlyContinue)
    }
    If ($ADExists -eq $True)
    {
        $ReportFile	= $ReportPath + "Report-PurgeAccount-EmpNo" + $Global:txtInpEmpNo.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
        $Global:GeneralTitle = "Purge Account" #set title for the SDAPGeneralAccountInfo form

        $EmpNo = $Global:txtInpEmpNo.Text + "@global.ul.com"
        $ADUsr = Get-ADUser $Global:txtInpEmpNo.Text -Properties Enabled,DisplayName,ExtensionAttribute1,"msRTCSIP-Line"
        write-host "`n`nStarting Account Purge Activities for" $Global:txtInpEmpNo.Text "-" $ADUsr.DisplayName "...." -ForegroundColor Green
        write-host "Retrieving AD Account Information...." -ForegroundColor Yellow
        $AdminAcct = "A" + $Global:txtInpEmpNo.Text
        $ADUAdm = [bool](get-ADUser -Filter {SamAccountName -eq $AdminAcct})
        $strDN = GetUserDN $Global:txtInpEmpNo.Text
        GetAcctInfo($Global:txtInpEmpNo.Text)
        $UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion
        write-host "Retrieving O365 Licensing Information...." -ForegroundColor Yellow
        $IsLic = Get-MsolUser -UserPrincipalName $EmpNo
        $UserLicense = Get-MsolUser -UserPrincipalName $EmpNo
        write-host "Retrieving Group Ownership/Membership Information...." -ForegroundColor Yellow
        $ADGroupMem = Get-ADPrincipalGroupMembership $Global:txtInpEmpNo.Text -ResourceContextServer global.ul.com
        $AZGroupMem = Get-AzureADUsermembership -ObjectId (Get-MsolUser -userPrincipalName $EmpNo).ObjectID
        $AZGroupOwner = Get-AzureADUserOwnedObject -ObjectId (Get-MsolUser -userPrincipalName $EmpNo).ObjectID
        $CnctExists = [bool] (Get-MailContact $EmpNo -ErrorAction SilentlyContinue)
        $HasExP2 = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:EXCHANGEENTERPRISE"})
        $HasE5 = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:ENTERPRISEPREMIUM"})

        write-host "Checking if account ever had a mailbox license assigned.... " -ForegroundColor Yellow

        if (($HasExP2 -eq $False) -and ($HasE5 -eq $False))
        {
            If ($UserLicense.UsageLocation -eq "US")
            {
                Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicenses "ul:EXCHANGEENTERPRISE"
                Write-Host "Assigned Email License and waiting for mailbox connection in order to purge O365 configs" -ForegroundColor Green
  		        $LineToWrite = $RecordEvent + "LISC" + "`t" + "Assigned License to Account " + $Global:txtInpEmpNo.Text
        	    WriteReportEvent

                Do
                {
                    start-sleep -Seconds 10
                    $Mbx = get-mailbox $Global:txtInpEmpNo.Text -ErrorAction SilentlyContinue
                }   While($mbx.Alias -ne $Global:txtInpEmpNo.Text)
                $HasExP2 = "True"
            }
            else
            {
			    $NoLic = "Y"
            }
        }

        If ($NoLic -ne "Y")
        {
            $MbxExist = [bool](Get-Mailbox $Global:txtInpEmpNo.Text -ErrorAction SilentlyContinue)
            If ($MbxExist -eq $True)
            {
                write-host "Retrieving O365 Group Ownership/Membership Information...." -ForegroundColor Yellow
                $OED = Import-Csv "\\Usnbks140p-spa\itd1\SharedFiles\OED_Extract6.csv" |Where-Object {($_."EMPLOYEE NUMBER" -eq $Global:txtInpEmpNo.Text)}
                $DLMember = Get-AzureADUser -SearchString $Global:txtInpEmpNo.Text | Get-AzureADUserMembership | Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application"))} |Sort-Object DisplayName
                $AZDLOwner = Get-AzureADUser -SearchString $Global:txtInpEmpNo.Text | Get-AzureADUserOwnedObject | Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application"))}
                $DistName = (get-User $Global:txtInpEmpNo.Text).DistinguishedName
                 If ($DistName -like "*'*")
                 {
                    $DistName = $Distname.Replace("'","")
                 }
#                $DLMember = (Get-Recipient -Filter "Members -eq '$DistName'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup | Select Name,RecipientTypeDetails)
                $DLOwner = (Get-Recipient -Filter "ManagedBy -eq '$DistName'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup -SortBy Name | Select Name,RecipientTypeDetails)
                write-host "Retrieving O365 Mailbox/Calendar Information...." -ForegroundColor Yellow
                $Mbx = get-mailbox $Global:txtInpEmpNo.Text
                $MbxOOO = get-mailboxautoreplyconfiguration $Global:txtInpEmpNo.Text
                $MbxPerm = Get-MailboxPermission $Global:txtInpEmpNo.Text | Where {($_.User -notlike "*SELF*")}
                $InboxName = Get-MailboxFolderStatistics $Global:txtInpEmpNo.Text | where {$_.FolderType -eq "Inbox"} |select Name #This accounts for mailboxes not in English
                $MbxFolderPerm = Get-MailboxFolderPermission ($Global:txtInpEmpNo.Text + ":\"+ $InboxName.Name) | Where {($_.User -notlike "*Default*") -and $_.User -notlike "*Anonym*"}
                $CalEven = Remove-CalendarEvents -Identity $Global:txtInpEmpNo.Text -CancelOrganizedMeetings -QueryWindowInDays 1825 -PreviewOnly -Confirm:$False
                write-host "Retrieving Mobile Device/Unified Messaging/Skpe Information...." -ForegroundColor Yellow
                $MblDevice = Get-MobileDevice -Mailbox $Global:txtInpEmpNo.Text
#                $MblDevice = Get-MobileDevice -Mailbox $Global:txtInpEmpNo.Text | ft FriendlyName,DeviceType,DeviceModel,WhenChanged
                If ($ADUsr."msRTCSIP-Line" -like "*=*")
                {
                    $ExtLoc = ($ADUsr."msRTCSIP-Line".IndexOf("=")) +1
                    $EUMMbx = ($ADUsr."msRTCSIP-Line".Substring($ExtLoc,(($ADUsr."msRTCSIP-Line".length)-$ExtLoc)))
                }
            }
            else
            {
			    Write-Host "Checking to see if there is a Mail Contact Record for this individual"
  		        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Checking to see if there is a Mail Contact Record for this individual " + $Global:txtInpEmpNo.Text
    	        WriteReportEvent
				
			    If ($CnctExists -eq "True")
                {
		            write-Host "Removing Contact Record for" $Global:txtInpEmpNo.Text " external address is" (Get-MailContact $Global:txtInpEmpNo.Text).ExternalEmailAddress
				    $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing Contact Record for " + $Global:txtInpEmpNo.Text + " external address is: " + (Get-MailContact $Global:txtInpEmpNo.Text).ExternalEmailAddress
				    WriteReportEvent					
    				Remove-MailContact $Global:txtInpEmpNo.Text -Confirm:$False
                }
                else
                {
	                write-Host "No Mail Contact Record exists for" $Global:txtInpEmpNo.Text
	                $LineToWrite = $RecordEvent + "INFO" + "`t" + "No Mail Contact Record exists for " + $Global:txtInpEmpNo.Text
	                WriteReportEvent					
                }
                write-host "Unable to find mailbox for this user" -ForegroundColor Red
                $Output = $wshell.Popup("Unable to find maibox for " + $Global:txtInpEmpNo.Text + ".",10,"Mailbox Not Found",0+32)
            }
        }

        $Global:OKDetails = "Purge"
        If ($Mbx.LitigationHoldEnabled -eq "True")
        {
            If (($MbxPerm.count -gt 0) -or ($MbxFolderPerm.count -gt 0))
            {
                $Global:OKDetails = "Remove Access"
            }
            else
            {
                $Global:OKDetails = "CleanUp"
            }
        }

        Build-SDAPGeneralAccountInfo
        If (($Global:txtType.Text -like "Ex*") -or (($ADUsr.DistinguishedName -like "*Disabled*") -or ($ADUsr.Enabled -eq $False)))
        {
            Add-ActionPurgeCheckBoxes
            Add-FormStandardButtons
        }
        else
        {
            Add-FormInfoOnlyButtons
        }
        Publish-form

    #  Return from information form and start purge process

        If ($Global:Result -eq "OK")
        {
            # Add start record to log file
            $LineToWrite = $RecordEvent + "STAR" + "`t" + "Purge Account script has started"
            WriteLogEvent
            WriteReportEvent
    	    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI
            WriteLogEvent
            WriteReportEvent
    	    $LineToWrite = $RecordEvent + "INFO" + "`t" + "User TimeZone: " + $UsrTimeZone + "`n"
            WriteLogEvent            
            Write-AccountDetails

            write-host 
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Purging Account for " + $Global:txtInpEmpNo.Text + " - " + $Global:txtName.Text + ", " + $Global:txtType.Text
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Version 1.8"
            WriteReportEvent			

            If (($ADUsr.ExtensionAttribute1 -notlike "*Ex-*") -and ($ADUsr.DistinguishedName -like "*Disabled*") -and ($Global:chkADNoShowCheck.Checked -ne "Checked"))
            {
                $Global:DoNotPurge = "Yes"
            }
            else
            {
                If ($Global:chkADNoShowCheck.Checked -eq "Checked")
                {
                    $Global:DoNotPurge = "No"
                    $LineToWrite = $RecordEvent + "INFO " + "`t" + "The No Show checkbox was selected"
                    WriteReportEvent
                }
            }
           
            If ($Global:DoNotPurge -eq "Yes")
            {
                If ($Mbx.LitigationHoldEnabled -eq "True")
                {
                    set-mailbox $Global:txtInpEmpNo.Text -MaxReceiveSize 0
                    $LineToWrite = $RecordEvent + "BLCK" + "`t" + "Setting the mailbox to prevent receipt of any new messages."
                    WriteReportEvent

                    write-host "Legal Hold enabled running to remove permissions to mailbox and remove from groups and hide from address book function."
                    If ($Global:Hidden -ne "True")
                    {
                        Execute-HideMbx	
                    }
                    If (($Global:RemoveMbxPerm -eq "Yes") -or ($Global:RemoveFldrPerm -eq "Yes"))
                    {
                        Remove-PurgeMbxFldrAccess
                    }
                    If (($DLMember.count -gt 0) -or ($DLOwner.Name.Count -gt 0) -or ($ADGroupMem.Name.count -gt 0))
                    {
                        write-host "Remove group membership/ownership"
                        $ENo = $Global:txtInpEmpNo.Text
                        Remove-DLMemberOwner
                    }
                    Rem-CalendarEvents
                    write-host "This account cannot be purged because the mailbox is on legal hold.  Access to the mailbox has been removed and the individual has been removed from the membership or ownership of any groups." -ForegroundColor Red
                    $Output = $wshell.Popup("This account cannot be purged because the mailbox is on legal hold.  Access to the mailbox has been removed and the individual has been removed from the membership or ownership of any groups.",10,"Unable to Purge",0+32)
                    $LineToWrite = $RecordEvent + "ERR " + "`t" + "This account cannot be purged because the mailbox is on legal.  Access to the mailbox has been removed and the individual has been removed from the membership or ownership of any groups."
                    WriteReportEvent
                }
                else
                {
                    If ($UsrDetails.ProtectedFromAccidentalDeletion -eq "True")
                    {
                        write-host "Legal Hold not enabled sending email to O365 team to review."
                        write-host "Sending email to O365 Admin team to remove AD Object Protection as account is not on Legal Hold"
                        $MsgTo = "LST.O365AdminTeam@ul.com"
                        $MsgCC = ""
                        $MsgSubject = "Action Required:  Review AD Object Protection for Account " + $Global:txtInpEmpNo.Text
                        $MsgBody = "This account came up for purging, it is not on Legal hold but the AD Object is protected.  Please reivew why this is set and remove it the account does not need to be retained.  Once complete inform the SDAP team so that the account can be purged by the SDAP Team."
                        Send-Message
                    }
                    write-host "This account cannot be purged because the user may be an active staff member, the AD Account is not disabled and/or the protected from accidental deletion." -ForegroundColor Red
                    $Output = $wshell.Popup("This account cannot be purged because the user may be an active staff member, the AD Account is not disabled and/or the protected from accidental deletion.",10,"Unable to Purge",0+32)
                    $LineToWrite = $RecordEvent + "ERR " + "`t" + "This account cannot be purged because the user may be an active staff member, the AD Account is not disabled and/or the protected from accidental deletion."
                    WriteReportEvent
                }
            }
            else
            {
                If ((($MblDevice.count-4) -gt 0) -and ($Global:chkSelWipe.Checked -ne "Checked"))
                {
                    write-host "This account cannot be purged because a selective wipe has not been performed on this account." -ForegroundColor Red
                    $Output = $wshell.Popup("This account cannot be purged because a selective wipe has not been performed on this account.",10,"Selective Wipe Required",0+32)
                    $LineToWrite = $RecordEvent + "ERR " + "`t" + "This account cannot be purged because a selective wipe has not been performed on this account."
                    WriteReportEvent
                }
                else
                {
                    If ($mbx.LitigationHoldEnabled -ne $True) #just double checking that the mailbox is not on Legal Hold!
                    {
                        If (($Global:ADEnabledRequired -ne "Yes") -or ($Global:chkADEnabCheck.Checked -eq "Checked") -or ($Global:chkADNoShowCheck.Checked -eq "Checked"))
                        {
                            If ($ADUAdm -eq $True)
                            {
                                write-host "This user has an elevated Admin Account.  Sending email to AD Support Team to purge this account." -ForegroundColor Cyan
                                $LineToWrite = $RecordEvent + "ADMN" + "`t" + "This user has an admin account " + $AdminAcct + " that needs to be purged by the AD Team an email has been sent."
                                WriteReportEvent
                                $MsgTo = "LST.O365AdminTeam@ul.com"
                                $MsgCC = ""
                                $MsgSubject = "Action Required: Purge Admin Account for Terminated Staff " + $AdminAcct
                                $MsgBody = "Please purge the admin account for " + $AdminAcct + " as this individuals 'standard' account has been purged by the SDAP Team."
                                Send-Message
                            }

                            If ($MbxExist -eq $True)
                            {
								If ($Global:RetentionHoldEnabled -eq $True)
                                {
                                    #remove retentionhold setting so mailbox can be permanently purged after 30 days!
									write-host "Set RetentionHoldEnabled to False for this account" -ForegroundColor Cyan
                                    Set-mailbox $Global:txtInpEmpNo.Text -RetentionHoldEnabled $False
                                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Set RetentionHoldEnabled value to False for this account"
                                    WriteReportEvent									
                                }
								
                                If (($Global:RemoveMbxPerm -eq "Yes") -or ($Global:RemoveFldrPerm -eq "Yes"))
                                {
                                    #remove mailbox/folder permissions not sure if this is needed when purging!
                                    Remove-PurgeMbxFldrAccess
                                }

                                Rem-CalendarEvents
                                If ($Mbx.LitigationHoldEnabled -eq $True)
                                {
                                    If ($ADUsr.DistinguishedName -notlike "*Extended*")
                                    {
                                        # Send email to O365 Team to move this account to the correct OU"
                                        write-host "Sending an email to the O365 Team to Move the this Account to the _ExtendedHold OU"
                                        $MsgTo = "Sandi.Glazebrook@ul.com"
                                        $MsgCC = "RJ.Borja@ul.com"
                                        $MsgSubject = "Review Required: Move Account to _ExtendedHold OU"
                                        $MsgBody = "The account " + $Global:txtInpEmpNo.Text + "-" + $mbx.Name + " reached the purge activity. This account is on an active Legal Hold please review and move the account from " + $mbx.DistinguishedName + "AD Container to the _ExtendedHold Container if it has not already been moved."
                                        Send-Message
                                    }
                                    If ($UsrDetails.ProtectedFromAccidentalDeletion -eq $True)
                                    {
                                      write-host "send email to O365 admin team to remove AD Object Projection"
                                    }
                                }

                                If ($Mbx.CustomAttribute1 -like "*Employee*")
                                {
                                    write-host "Checking to see if anyone has been granted access to OneDrive for this account" -ForegroundColor Cyan
                                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Checking to see if anyone has been granted access to OneDrive for this account"
                                    WriteReportEvent
                                    Remove-OneDriveAccess
                                }

                               write-host "Checking to see if this user has a UM Mailbox configured" -ForegroundColor Cyan
                                if ($Global:txtUMExt.Text -ne "(None)")
                                {
                                    write-host "Releasing Phone Number" -ForegroundColor Cyan
                                    $LineToWrite = $RecordEvent + "DISB" + "`t" + "Removing Extension Associated with Account: " + $ADUsr.'msRTCSIP-Line'
                                    WriteReportEvent
#                                    Out-File -filepath $ReportFile -append -noClobber -inputObject $ADUsr.'msRTCSIP-Line'
                                    $PhCfg = Get-ADUser $Global:txtInpEmpNo.Text -Properties msRTCSIP-Line
                                    If ($PhCfg.'msRTCSIP-Line'.Length -gt 0)
                                    {
                                        Set-ADUser $Global:txtInpEmpNo.Text -Remove @{'msRTCSIP-Line'=$ADUsr.'msRTCSIP-Line'} 
                                    }
#                                    Disable-UMMailbox -Identity $Global:txtInpEmpNo.Text -Confirm:$False
                                }
                                else
                                {
                                    write-host "No Extension Associated to this Account" -ForegroundColor Red
                                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "No Extension Associated for this Account"
                                    WriteReportEvent
                                }


                                write-host "Checking to see if this user has Mobile Devices associated with their Account" -ForegroundColor Cyan
                                $MblExist = [bool](Get-MobileDevice -Mailbox $Global:txtInpEmpNo.Text -ErrorAction SilentlyContinue)
                                If (($MblDevice.count -gt 0) -and ($Global:chkSelWipe.Checked -eq "Checked"))
#                                If ((($MblDevice.count-4) -gt 0) -and ($Global:chkSelWipe.Checked -eq "Checked"))
                                {
                                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Selective Wipe Checkbox Value:" + $Global:chkSelWipe.Checked
                                    WriteReportEvent
                                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Mobile Devices Associated with Account"
                                    WriteReportEvent
                                    Out-File -filepath $ReportFile -append -noClobber -inputObject ($MblDevice | ft FriendlyName,DeviceType,DeviceModel,WhenChanged)
                
                                    Write-Host "Removing Mobile Device(s) from account"
                                    $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing Mobile Device(s) from account"
                                    WriteReportEvent
                                    $MblDevice = Get-MobileDevice -Mailbox $Global:txtInpEmpNo.Text
                                    foreach ($MblDevice in $MblDevice)
                                    {
                                        write-host "Removing Mobile Device: (Friendly Name: " $MblDevice.FriendlyName ") - (DeviceModel: " $MblDevice.DeviceModel ")"
                                        Remove-MobileDevice -Identity $MblDevice.DistinguishedName -Confirm:$False
                                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing Mobile Device: (Friendly Name:" + $MblDevice.FriendlyName + ") - (Device Model:" + $MblDevice.DeviceModel + ") from account"
                                        WriteReportEvent
                                    }
                                }
                                else
                                {
                                    write-host "This user does not have any Mobile Devices configured" -ForegroundColor Red
                                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "No Mobile Devices Associated with Account"
                                    WriteReportEvent
                                }
                                write-host "Removing individual from AD Groups and O365 Distribution Lists they are a member or owner of" -ForegroundColor Cyan
                                Remove-DLMemberOwner

                            }
                            else
                            {
        			            If ($CnctExists -eq "True")
                                {
		                            write-Host "Removing Contact Record for" $Global:txtInpEmpNo.Text " external address is" (Get-MailContact $Global:txtInpEmpNo.Text).ExternalEmailAddress
				                    $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing Contact Record for " + $Global:txtInpEmpNo.Text + " external address is: " + (Get-MailContact $Global:txtInpEmpNo.Text).ExternalEmailAddress
			    	                WriteReportEvent					
    	    			            Remove-MailContact $Global:txtInpEmpNo.Text -Confirm:$False
                                }
                            }

    	                    Remove-ADAcct

                            $LineToWrite = $RecordEvent + "INFO" + "`t" + "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
                            WriteReportEvent
                            write-host "`n++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
                            write-host "Purge Request complete" -ForegroundColor Red
                            $Output = $wshell.Popup("Purge request complete.",10,"Complete",0+32)
                        }
                        else
                        {
                            write-host "Purge not performed this users account is not disabled and the checkbox to confirm OK to delete was not selected." -ForegroundColor Red
                            $Output = $wshell.Popup("Purge not performed this users account is not disabled and the checkbox to confirm OK to delete was not selected.",10,"Purge Cancelled",0+32)
                            $LineToWrite = $RecordEvent + "ERR " + "`t" + "Purge not performed this users account is not disabled and the checkbox to confirm OK to delete was not selected."
                            WriteReportEvent
                        }
                    }
                    else
                    {
                        write-host "Purge not performed account is on Legal Hold or the AD Account is not disabled." -ForegroundColor Red
                        $Output = $wshell.Popup("Purge not performed account is on Legal Hold or the AD Account is not disabled.",10,"Purge Cancelled",0+32)
                        $LineToWrite = $RecordEvent + "ERR " + "`t" + "Purge not performed account is on Legal Hold or the AD Account is not disabled."
                        WriteReportEvent
                    }
                }
            }
        }
        else
        {
            write-host "Purge Request cancelled" -ForegroundColor Red
            $Output = $wshell.Popup("Purge request cancelled.",10,"Cancelled",0+32)
        }
        $Global:OKDetails = ""
    }
    else
    {
        write-host "This individual does not have an AD Account.  Purge Process Aborted" -ForegroundColor Red
        $Output = $wshell.Popup("This individual " + $Global:txtInpEmpNo.Text + " does not have an Active Directory account.",10,"No AD Account",0+32)
    }

	# Write end record to the log file
    $LineToWrite = $RecordEvent + "STOP" + "`t" + "This instance is stopping." + "`n"
	WriteLogEvent
}
else
{
    write-host "Request cancelled" -ForegroundColor Red
    $Output = $wshell.Popup("Request cancelled.",10,"Cancelled",0+32)
}