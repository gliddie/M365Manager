#####################################################################################
#
#   Execute Account Standard or Emergency Termination
#
#   12/11/2020 - Created New Script from originals SDAPAdminMenu details$Global:RemoveMbxPerm
#   02/23/2021 - Modifed the Enable/Disable-CSUser to use the "Get-ADUser" and "Set-ADUser -clear"
#   06/02/2021 - Added Function to generate a complex password
#   07/16/2021 - Added *Hold* to the checks when writing CustomAttribute14
#   07/01/2022 - Added values to store a date for the account purge
#   10/14/2022 - Modifed group clean-up due to changes to the Group/Owner clean-up process
#   11/01/2022 - Fixed the removal of mobile devices from accounts
#   03/03/2023 - Added code to prevent the removal of steritech forwarders and putting an OOO on mailbox with that forwarder
#   04/13/2023 - Changed the $Global:u lookups and $strPath that used commitchange to use the set-aduser and move-adobject
#   05/18/2023 - Added code to create files in a directory for the year they were created
#
#####################################################################################

#Put in SDAP Admin menu

Function Add-ActionTerminationCheckBoxes
{
    $Tab = 0
    $Global:chkDelegateAccess = ""
    $Global:txtDelegateAccess = ""

    If ($MbxExist -eq $True)
    {
        $form.Height = $form.Height + 70
        ## Add option to grant delegate Access
        $Global:Top = $Global:Top + 20
        $Global:chkDelegateAccess = New-Object Windows.Forms.checkbox
        $Global:chkDelegateAccess.Left = 200; $Global:chkDelegateAccess.Width = 200; $Global:chkDelegateAccess.Top = $Global:Top
        $Global:chkDelegateAccess.Text = "Grant Delegate Access to emp #"
        $Global:chkDelegateAccess.Checked = $false   # set a default value
        $Global:chkDelegateAccess.TabIndex = $Tab++
        $Global:form.Controls.Add($Global:chkDelegateAccess)
        $Global:txtDelegateAccess = New-Object Windows.Forms.TextBox
        $Global:txtDelegateAccess.TabIndex = $Tab++
        $Global:txtDelegateAccess.Top = $Global:Top; $Global:txtDelegateAccess.Left = 400; $Global:txtDelegateAccess.Width = 70;
        $Global:txtDelegateAccess.Text = ""
        $Global:form.Controls.Add($Global:txtDelegateAccess)    # Add to Form

        ## Purge Extension Requested
        $Global:Top = $Global:Top + 20
        $Global:chkPurgeExt = New-Object Windows.Forms.checkbox
        $Global:chkPurgeExt.Left = 200; $Global:chkPurgeExt.Width = 200; $Global:chkPurgeExt.Top = $Global:Top
        $Global:chkPurgeExt.Text = "Purge Extension Requested until"
        $Global:chkPurgeExt.Checked = $false   # set a default value
        $Global:chkPurgeExt.TabIndex = $Tab++
        $Global:form.Controls.Add($Global:chkPurgeExt)
        $Global:txtPurgeExt = New-Object Windows.Forms.TextBox
        $Global:txtPurgeExt.TabIndex = $Tab++
        $Global:txtPurgeExt.Top = $Global:Top; $Global:txtPurgeExt.Left = 400; $Global:txtPurgeExt.Width = 70;
        $Global:txtPurgeExt.Text = ""
        $Global:form.Controls.Add($Global:txtPurgeExt)    # Add to Form

        ## OOO Referral
        $Global:Top = $Global:Top + 20
        $Global:chkOOORef = New-Object Windows.Forms.checkbox
        $Global:chkOOORef.Left = 200; $Global:chkOOORef.Width = 200; $Global:chkOOORef.Top = $Global:Top
        $Global:chkOOORef.Text = "Custom OOO Referral Address"
        $Global:chkOOORef.Checked = $false   # set a default value
        $Global:chkOOORef.TabIndex = $Tab++
        $Global:form.Controls.Add($Global:chkOOORef)
        $Global:txtOOORef = New-Object Windows.Forms.TextBox
        $Global:txtOOORef.TabIndex = $Tab++
        $Global:txtOOORef.Top = $Global:Top; $Global:txtOOORef.Left = 400; $Global:txtOOORef.Width = 200;
        $Global:txtOOORef.Text = ""
        $Global:form.Controls.Add($Global:txtOOORef)    # Add to Form

        If ((($MblDevice.count-4) -gt 0))
        {
            $Global:Top = $Global:Top + 20
            $Global:chkSelWipe = New-Object Windows.Forms.checkbox
            $Global:chkSelWipe.Left = 200; $Global:chkSelWipe.Width = 200; $Global:chkSelWipe.Top = $Global:Top
            $Global:chkSelWipe.Text = "Mobile Device Wipe Complete"
            $Global:chkSelWipe.Checked = $false   # set a default value
            $Global:chkSelWipe.TabIndex = $Tab++
            $Global:form.Controls.Add($Global:chkSelWipe)
        }
    }
}

Function Generate-UsrPassword
{
    $PassComplexCheck = $False
    $chars = "abcdefghijkmnopqrstuvwxyzABCEFGHJKLMNPQRSTUVWXYZ0123456789$!#&".ToCharArray()
    do{
        $Global:Pwd=""
        1..10 | ForEach { $Global:Pwd += $chars | Get-Random }
        If (($Global:Pwd -cmatch "[A-Z]") -and ($Global:Pwd -cmatch "[a-z]") -and ($Global:Pwd -cmatch "[$!#&]"))
        {
            $PassComplexCheck = $True
        }
    } while ($PassComplexCheck -eq $False)
}

Function Remove-FullAccessMbxPerm
{
    # Reset any permissions that were granted to ReadyOnly access
    $FldrTypes = import-csv E:\O365AdminShared\Data\FolderTypes.csv
    #Remove access at top of mailbox
    $Global:MbxPerm = Get-MailboxPermission $Global:txtInpEmpNo.Text | Where {($_.User -notlike "*SELF*")}
    Foreach ($Perm in $Global:MbxPerm)
    {
        If ($Global:MbxPerm.AccessRights -notlike "*Read*")
        {
            Remove-MailboxPermission $Global:txtInpEmpNo.Text -User $Perm.User -AccessRights $Perm.AccessRights -confirm:$False |out-null
            Add-MailboxPermission $Global:txtInpEmpNo.Text -User $Perm.User -AccessRights ReadPermission -confirm:$False |out-null
            write-host "Resetting" $Perm.User "-->" $Perm.AccessRights "mailbox permissions to ReadyOnly"
            $LineToWrite = "REMV" + "`tResetting " + $Perm.User + " --> " + $Perm.AccessRights + " to ReadOnly"
            WriteReportEvent
        }
        else
        {
            write-host "This individual" $Perm.User "-->" $Perm.AccessRights "already has Read Permission to the mailbox"
            $LineToWrite = "REMV" + "`tThis individual " + $Perm.User + " --> " + $Perm.AccessRights + " already has Read permissions to the mailbox"
            WriteReportEvent            
        }
    }
        
    #Remove access to identified Key folder types
    $MBXFldr = Get-MailboxFolderStatistics $Global:txtInpEmpNo.Text
    $Folder = $Global:txtInpEmpNo.Text + ":\"
    foreach ($Fldr in $MBXFldr)
    {
        foreach ($Type in $FldrTypes)
        {
            If ($Fldr.FolderType -eq $Type.FolderType)
            {
                If ($Fldr.Foldertype -ne "Root")
                {
                    $Folder = $Global:txtInpEmpNo.Text + ":\" + $Fldr.Name
                }
                $Global:FLDRPerm = Get-MailboxFolderPermission $Folder | Where {($_.User -notlike "*Default*") -and $_.User -notlike "*Anonym*"}
                If ($Global:FLDRPerm.count -ne 0)
                {
                    Foreach ($Fldr1 in $Global:FLDRPerm)
                    {
                        #Remove Permissions
                        If ($Fldr1.AccessRights -notlike "*Reviewer*")
                        {
                            Remove-MailboxFolderPermission $Folder -User $Fldr1.User.DisplayName -Confirm:$False |out-null
                            Add-MailboxFolderPermission $Folder -User $Fldr1.User.DisplayName -AccessRights Reviewer
                            write-host "Resetting access to" $Folder "folder for " $Fldr1.User.DisplayName " from: " $Fldr1.AccessRights "to: Reviewer" -ForegroundColor Red
                            $LineToWrite = "REMV" + "`t" + "Resetting access to " + $Folder + " folder " + $Fldr1.User.DisplayName + " from: " + $Fldr1.AccessRights + " to Reviewer"
                            WriteReportEvent
                        }
                        else
                        {
                            write-host "This individual" $Fldr1.User.DisplayName "-->" $Fldr1.AccessRights "already has Read Permission to the " $Folder " folder"
                            $LineToWrite = "REMV" + "`tThis individual " + $Fldr1.User.DisplayName + " --> " + $Fldr1.AccessRights + " already has Read permissions to the " + $Folder + "folder"
                            WriteReportEvent 
                        }
                    }
                }
            }
        }
    }
}

Function Set-000Msg
{
    If ($Global:chkOOORef.Checked -eq "Checked")
    {
        $StdOOOMsg ="The individual you are attempting to reach is unavailable.  In their absence please reach out to " + $Global:txtOOORef.Text + "."
    }
    else
    {
        $Super = Get-AdUser $Mbx.Alias -Properties Manager
        If ($Super.Manager.Length -gt 0)
        {
            $strMgrPath = [string]::format("LDAP://{0}", $Super.Manager)
            $u1 = new-object System.DirectoryServices.DirectoryEntry($strMgrPath)
            $StdOOOMsg ="The individual you are attempting to reach is unavailable.  In their absence please reach out to " + $u1.mail.value + "."
        }
        else
        {
            write-host "User does not have a supervisor configured please enter the address for people to contact. If no information is provided the out of office messag will not be enabled. " -ForegroundColor Cyan -NoNewline
            $Super = Read-Host
            $Super = $Super2.Trim()
            $StdOOOMsg ="The individual you are attempting to reach is unavailable.  In their absence please reach out to " + $Super + "."
        }
    }

    if (($MbxOOO.AutoReplyState -eq "Scheduled") -and (($mbx.ForwardingAddress -notlike "steritech.com") -and ($mbx.ForwardingSMTPAddress -notlike "steritech.com")))
    {
        if ($MbxOOO.InternalMessage -notlike "*In their absence*")
        {
            write-host "Out Of Office is currently enabled with a non-standard out of office messsage. Resetting to the standard out of office message." -ForegroundColor Yellow
            $LineToWrite = $RecordEvent + "OOOR" + "`t" + "Resetting Out Of Office currently enabled  out of office with our standard message"
            WriteReportEvent
        }
    }
    #else
    #{
        If (($Super.Manager.Length -gt 0) -or ($Super.Length -gt 0))
        {
            write-host "Configuring standard out of office messsage." -ForegroundColor Yellow
            $LineToWrite = $RecordEvent + "OOOE" + "`t" + "Configuring standard out of office message"
            WriteReportEvent
            Set-MailboxAutoReplyConfiguration $Global:txtInpEmpNo.Text –AutoReplyState Scheduled –StartTime (get-date) -EndTime (get-date).Adddays(30) –ExternalMessage $StdOOOMsg -InternalMessage $StdOOOMsg
            write-host "The Internal/External responses have been set to:" -ForegroundColor Cyan
            write-host "Internal Message: " (Get-MailboxAutoReplyConfiguration $Global:txtInpEmpNo.Text).InternalMessage -ForegroundColor Cyan
            write-host "External Message: " (Get-MailboxAutoReplyConfiguration $Global:txtInpEmpNo.Text).ExternalMessage -ForegroundColor Cyan
        }
        else
        {
            write-host "Not Configuring an out of office message because no supervisor exists or an address was not provided." -ForegroundColor Red
            $LineToWrite = $RecordEvent + "OOOE" + "`t" + "No out of office message configured no supervisor found or no address was provided"
            WriteReportEvent
        }
    #}
    
}

#########
#   Start of Script
#########

$Script:NoOwnerPath = "\\usnbkemes100p\E:\O365AdminShared\Data\NoOwnerNotifications.csv"
$Global:strUserPath = ""
$Global:u = ""
$Global:inf = ""
$Global:MBXAccess = ""
$strDN = 0
$Cont = "N"

$MgrAccess = "N"
Enter-EmpNoInputForm
Publish-Form

If ($Global:Result -eq "OK")
{
$UsrTimeZone = Get-TimeZone
$Year = (get-date).ToString("yyyy")

If (Test-Path $ReportPath) {} else {New-Item -Path $ReportPath -ItemType Directory}
    $NoOwnerTip = "Message sent to obtain group owners will be deleted if no response by " + (Get-date).AddDays(30).ToString('yyyy-MMM-dd') + " - Per " + $Global:txtInpTicketNo.Text

    If ($Global:chkEmerTerm.Checked -eq "Checked")
    {
        $Filename = "EmergencyDisable"
        $LogFile  = "E:\SDAP\EmergencyDisable\Log\Log-" + $FileName + ".log" 
    }
    else
    {
        $Filename = "StandardTermination"
        $LogFile  = "E:\SDAP\StandardTermination\Log\Log-" + $FileName + ".log"      
    }

    $Usr = $Global:txtInpEmpNo.Text
    If ($usr.length -le 7)
    {
        $ADExists = [bool](get-ADUser -Filter {SamAccountName -eq $Usr} -ErrorAction SilentlyContinue)
    }

    If ($ADExists -eq $True)
    {
        $ReportPath = "E:\SDAP\StandardTermination\Report\" + $Year + "\"
        $ReportFile	= $ReportPath + "Report-StandardTermination-EmpNo" + $Global:txtInpEmpNo.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
        $Global:GeneralTitle = "Standard Termination" #set title for the SDAPGeneralAccountInfo form    
        
        If ($Global:chkEmerTerm.Checked -eq "Checked")
        {
            $ReportPath = "E:\SDAP\EmergencyDisable\Report\" + $Year + "\"
            $ReportFile	= $ReportPath + "Report-EmergencyDisable-EmpNo" + $Global:txtInpEmpNo.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            $Global:GeneralTitle = "Emergency Termination" #set title for the SDAPGeneralAccountInfo form   
        }
        If (Test-Path $ReportPath) {} else {New-Item -Path $ReportPath -ItemType Directory}

    # Add start record to log file
        $LineToWrite = $RecordEvent + "STAR" + "`t" + "Account Termination script has started"
        WriteLogEvent
        WriteReportEvent
    	$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI
        WriteLogEvent
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "User TimeZone: " + $UsrTimeZone
        WriteReportEvent
    	$LineToWrite = $RecordEvent + "INFO" + "`t" + "Script Last Change Date: 05/18/2023" + "`n"
        WriteReportEvent

        $EmpNo = $Global:txtInpEmpNo.Text + "@global.ul.com"
        $ADUsr = Get-ADUser $Global:txtInpEmpNo.Text -Properties Enabled,userCertificate,DisplayName,ExtensionAttribute1
        write-host "`n`nStarting Account Termination for" $Global:txtInpEmpNo.Text "-" $ADUsr.DisplayName "...." -ForegroundColor Red
        write-host "Retrieving AD Account Information...." -ForegroundColor Yellow
        $AdminAcct = "A" + $Global:txtInpEmpNo.Text
        $ADUAdm = [bool](get-ADUser -Filter {SamAccountName -eq $AdminAcct})
        $strDN = GetUserDN $Global:txtInpEmpNo.Text
        GetAcctInfo($Global:txtInpEmpNo.Text)
        $UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion
        write-host "Retrieving O365 Licensing Information...." -ForegroundColor Yellow
        $UserLicense = Get-MsolUser -UserPrincipalName $EmpNo
        write-host "Retrieving Group Ownership/Membership Information...." -ForegroundColor Yellow
        $ADGroupMem = Get-ADPrincipalGroupMembership $Global:txtInpEmpNo.Text -ResourceContextServer global.ul.com
        $AZGroupOwner = Get-AzureADUserOwnedObject -ObjectId (Get-MsolUser -userPrincipalName $EmpNo).ObjectID
        $CnctExists = [bool] (Get-MailContact $EmpNo -ErrorAction SilentlyContinue)
        $HasExP2 = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:EXCHANGEENTERPRISE"})
        $HasE3 = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:ENTERPRISEPACK"})

        write-host "Checking if account ever had a mailbox license assigned.... " -ForegroundColor Yellow

        if (($HasExP2 -eq $False) -and ($HasE3 -eq $False))
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
                $DistName = (get-User $Global:txtInpEmpNo.Text).DistinguishedName
                If ($DistName -like "*'*")
                {
                   $DistName = $Distname.Replace("'","")
                }
                $DLOwner = (Get-Recipient -Filter "ManagedBy -eq '$DistName'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup | Select Name,RecipientTypeDetails)
                write-host "Retrieving O365 Mailbox/Calendar Information...." -ForegroundColor Yellow
                $Mbx = get-mailbox $Global:txtInpEmpNo.Text
                $MbxOOO = get-mailboxautoreplyconfiguration $Global:txtInpEmpNo.Text
                $MbxPerm = Get-MailboxPermission $Global:txtInpEmpNo.Text | Where {($_.User -notlike "*SELF*")}
                $MBXFldr = Get-MailboxFolderStatistics $Global:txtInpEmpNo.Text
                $InboxName = Get-MailboxFolderStatistics $Global:txtInpEmpNo.Text | where {$_.FolderType -eq "Inbox"} |select Name #This accounts for mailboxes not in English
                $MbxFolderPerm = Get-MailboxFolderPermission ($Global:txtInpEmpNo.Text + ":\"+ $InboxName.Name) | Where {($_.User -notlike "*Default*") -and $_.User -notlike "*Anonym*"}
                $CalEven = Remove-CalendarEvents -Identity $Global:txtInpEmpNo.Text -CancelOrganizedMeetings -QueryWindowInDays 1825 -PreviewOnly -Confirm:$False
                write-host "Retrieving Mobile Device/Unified Messaging/Skpe Information...." -ForegroundColor Yellow
                $MblDevice = Get-MobileDevice -Mailbox $Global:txtInpEmpNo.Text
                $TeamsUsr = Get-ADUser $Global:txtInpEmpNo.Text -Properties msRTCSIP-UserEnabled
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
                $Output = $wshell.Popup("Unable to find maibox for " + $Global:txtInpEmpNo.Text + ".",0,"Mailbox Not Found",0+32)
            }
        }

        $Global:OKDetails = "Terminate"

        Build-SDAPGeneralAccountInfo
        Add-ActionTerminationCheckBoxes
        Add-FormStandardButtons
        $Global:InputFocus = $Global:cancelButton
        Publish-form
        Write-AccountDetails

        # Addding details for delegate"
        If ($Global:chkDelegateAccess.Checked -eq "Checked")
        {
            $LineToWrite = $RecordEvent + "INFO" + "`t" + " Delegate Access Checkbox: " + $Global:chkDelegateAccess.Checked
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t" + " Grant Delegate Access To: " + $Global:txtDelegateAccess.Text
            WriteReportEvent
        }

        If ($Global:chkPurgeExt.Checked -eq "Checked")
        {
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Purge Extension Requested: " + $Global:chkPurgeExt.Checked
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "     Requested Purge Date: " + $Global:txtPurgeExt.Text
            WriteReportEvent
        }

    #  Return from information form and start purge process
        If ($Global:Result -eq "OK")
        {
            write-host 
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Starting Standard Termination for " + $Global:txtInpEmpNo.Text + " - " + $Global:txtName.Text + ", " + $Global:txtType.Text
            WriteReportEvent
            $DescriptionComment = "Standard Termination Request - " + $Global:txtInpTicketNo.Text
            If ($Global:chkEmerTerm.Checked -eq "Checked")
            {
                $LineToWrite = $RecordEvent + "INFO" + "`t" + "Starting Emergency Termination for " + $Global:txtInpEmpNo.Text + " - " + $Global:txtName.Text + ", " + $Global:txtType.Text
                $DescriptionComment = "Emergency Disablement Request - " + $Global:txtInpTicketNo.Text
            }

            If ($Global:chkPurgeExt.Checked -eq "Checked")
            {
                $DescriptionComment = $Descriptioncomment + " - Purge Date: " + $Global:txtPurgeExt.Text
            }

            If ($ADUAdm -eq $True)
            {
                write-host "This user has an elevated Admin Account.  Sending email to AD Support Team to disable this account." -ForegroundColor Cyan
                $LineToWrite = $RecordEvent + "ADMN" + "`t" + "This user has an admin account " + $AdminAcct + " that needs to be purged by the AD Team an email has been sent."
                WriteReportEvent
                $MsgTo = "LST.O365AdminTeam@ul.com"
                $MsgCC = "EnterpriseMessagingServices@ul.com"
                $MsgSubject = "Action Required: Disable Admin Account for Terminated Staff " + $AdminAcct
                $MsgBody = "Please disable the admin account for " + $AdminAcct + " as this individuals 'standard' account has has been disabled by the SDAP Team and the termination process has begun."
                Send-Message
            }

            if ($ADUsr.Enabled -eq $true)
            {
                Write-host "Disabling AD Account and AzureAD Account..." -ForegroundColor Green
                $LineToWrite = $RecordEvent + "DISA" + "`t" + "Disabling Active Directory and AzureAD Account for: " + $Global:txtInpEmpNo.Text
                Disable-ADAccount -Identity $Global:txtInpEmpNo.Text
                Set-AzureADUser -ObjectID $EmpNo -AccountEnabled $false
            }
            else
            {
                write-host "AD Account is already disabled..." -ForegroundColor Green
                $LineToWrite = $RecordEvent + "INFO" + "`t" + "AD Account is already disabled: " + $Global:txtInpEmpNo.Text
            }
            WriteReportEvent

            write-host "Resetting AD Description to" $DescriptionComment "..." -ForegroundColor Green
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Resetting Description to " + $DescriptionComment
            WriteReportEvent
            Set-ADUser $Global:txtInpEmpNo.Text -Replace @{'Description'=$DescriptionComment}
#            $Global:u.Description.value = $DescriptionComment
#            $Global:u.CommitChanges()
            
            $v = get-aduser $Global:txtInpEmpNo.Text -Properties * |select msDS-cloudExtensionAttribute1
            $PurgeDate = (Get-Date).AddDays(14).ToString("MM/dd/yyyy")
            If ($v."msDS-cloudExtensionAttribute1".length -gt 0)
            {
                write-host "Resetting Purge Date" -ForegroundColor Green
                $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Resetting Account Purge Date to 14 Days from today"
                Set-ADUser $Global:txtInpEmpNo.Text -Replace  @{"msDS-cloudExtensionAttribute1" = $PurgeDate}
            }
            else
            {
                write-host "Adding Purge Date" -ForegroundColor Green
                $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Adding Account Purge Date to 14 Days from today"
                Set-ADUser $Global:txtInpEmpNo.Text -Add  @{"msDS-cloudExtensionAttribute1" = $PurgeDate}
            }
            WriteReportEvent

            write-host "Resetting On-Prem and AzureCloud Account Password..." -ForegroundColor Green
            $Retry = 0
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Resetting AD Password with a Random Generated Password" +"`n"
            Do
            {
                Generate-UsrPassword
#                $chars = "abcdefghijkmnopqrstuvwxyzABCEFGHJKLMNPQRSTUVWXYZ23456789!#%&?".ToCharArray()
#                $Pwd=""
#                1..10 | ForEach {  $Pwd += $chars | Get-Random }

                Set-ADAccountPassword -Identity $Global:txtInpEmpNo.Text -NewPassword (ConvertTo-SecureString -AsPlainText [string]$Global:Pwd -Force)
                Set-MsolUserPassword –UserPrincipalName $EmpNo –NewPassword $Pwd -ForceChangePassword $False
                Set-AzureADUser -ObjectID $EmpNo -AccountEnabled $false
                sleep 5
                If ($Retry -gt 0)
                {
                    write-host "AD Password was not reset -- Attempting to reset the password again" -ForegroundColor Cyan
                }
                $Retry++
                $ChkPwd = Get-ADUser $Global:txtInpEmpNo.Text -Properties PwdLastSet

                If ($Retry -ge 5)
                {
                    write-host "It appears the AD password was not reset -- please report this to the Enterprise Messaging Services Team" -ForegroundColor Red
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unable to Reset AD Password with a Random Generated Password" +"`n"
                    $CurrDate = (get-date).AddDays(-365)
                }
            } While ($CurrDate -gt ([datetime]::FromFileTime($ChkPwd.PwdLastSet)) -and ($Retry -lt 5))
            WriteReportEvent

            write-host "Revoking O365 Tokens..." -ForegroundColor Green
            Revoke-AzureADUserAllRefreshToken -ObjectID $UserLicense.ObjectID
            $LineToWrite = $RecordEvent + "REVO" + "`t" + "Revoking O365 AzureADUser Tokens"
            WriteReportEvent

            write-host "Revoking O365 Active Sessions..." -ForegroundColor Green
            write-host "If you get the warning that it could not find an active SPO User session to revoke"
            Revoke-SPOUserSession -User $EmpNo -Confirm:$false
            $LineToWrite = $RecordEvent + "REVO" + "`t" + "Revoking O365 Active Sessions"
            WriteReportEvent

            if ($MbxExist -eq $True)
            {
                Set-000Msg
                If (($Global:RemoveMbxPerm -eq "Yes") -or ($Global:RemoveFldrPerm -eq "Yes") -or ($Global:txtDelegateAccess.Text -ne ""))
                {
                    $MbxFldrCnt = $MbxFldr.count
                    $MBXFldr = Get-MailboxFolderStatistics $Global:txtInpEmpNo.Text | % {$_.Identity.ToString().Split("\")[1..$MBXFldrCnt] -join "\"}
                    Remove-FullAccessMbxPerm
                }

                If (($Global:chkDelegateAccess.Checked -eq "Checked") -and ($Global:txtDelegateAccess.Text -ne ""))
                {
                    If ([bool](get-mailbox $Global:txtDelegateAccess.Text -ErrorAction SilentlyContinue))
                    {
                        If ($MbxFldrCnt -eq 0)
                        {
                            $MbxFldrCnt = $MbxFldr.count
                            $MBXFldr = Get-MailboxFolderStatistics $Global:txtInpEmpNo.Text | % {$_.Identity.ToString().Split("\")[1..$MBXFldrCnt] -join "\"}
                        }
                        Set-DelegateAccess
                    }
                    else
                    {
                        write-host "Delegate mailbox not found, unable to grant access to" $Global:txtDelegateAccess.Text "the specified delegate"
                        $LineToWrite = $RecordEvent + "ERRO" + "`t" + "Delegate mailbox not found, unable to grant access to " + $Global:txtDelegateAccess.Text + "the specified delegate" 
                    }
                }

                $MbxAccess = Get-MailboxPermission $EmpNo |where {$_.IsInherited -eq $False} -ErrorAction SilentlyContinue
                if ($MbxAccess.count -lt 1)
                {
                    write-host "Disabling O365 Mailbox..." -ForegroundColor Green
                    Set-Mailbox $Global:txtInpEmpNo.Text -AccountDisabled:$True
                    $LineToWrite = $RecordEvent + "DISA" + "`t" + "Disabling O365 Mailbox"
                    WriteReportEvent        
                    write-host "Disabling Mailbox Protocols for OWA, OWAForDevices, ActiveSync and AllowMAC's..." -ForegroundColor Green
                    Set-CASMailbox -Identity $Global:txtInpEmpNo.Text -OwaEnabled $false -OwaForDevicesEnabled $false -ActiveSyncEnabled $false -EwsAllowMacOutlook $false
                    $LineToWrite = $RecordEvent + "DISA" + "`t" + "Disabling Mailbox Protocols for OWA, OWAForDevices, ActiveSync and AllowMAC's"
                    WriteReportEvent
                    write-host "Setting user to be Hidden from the Address Book..." -ForegroundColor Green
                    Set-ADUser $Global:txtInpEmpNo.Text -Replace @{'msExchHideFromAddressLists'=$True}
#                    $Global:u.msExchHideFromAddressLists.value = $True
#                    $Global:u.CommitChanges()
                    $LineToWrite = $RecordEvent + "HIDE" + "`t" + "Setting user to be Hidden from the Address Book"
                    WriteReportEvent
                }
                else
                {
                    write-host "Individuals were granted access to the maibox not disabling protocols"
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Individuals were granted access to the maibox not disabling protocols"
                    WriteReportEvent
                    
                    if ($Global:u.msExchHideFromAddressLists.value -ne $False)
                    {
                        write-host "Setting user to be Unhidden from the Address Book..." -ForegroundColor Green
                        Set-ADUser $Global:txtInpEmpNo.Text -Replace @{'msExchHideFromAddressLists'=$False}
#                        $Global:u.msExchHideFromAddressLists.value = $False
#                        $Global:u.CommitChanges()
                        $LineToWrite = $RecordEvent + "UNHI" + "`t" + "Setting user to be Unhidden from the Address Book"
                        WriteReportEvent
                    }
                    else
                    {
                        write-host "The AD Account/User Mailbox is already hidden from the Address Book..." -ForegroundColor Red
                        $LineToWrite = $RecordEvent + "INFO" + "`t" + "The AD Account/User Mailbox is already hidden from the Address Book" +"`n"
                        WriteReportEvent
                    }
                }

                write-host "Removing individual from AD Groups and O365 Distribution Lists they are a member or owner of" -ForegroundColor Cyan
                Remove-DLMemberOwner

                if (($Mbx.LitigationHoldEnabled -eq "True") -and ($Global:chkDelegateAccess.Checked -eq "Checked"))
                {
                    $SendTo = $Global:LegalTeamMsgs
                    if ($Mbx.CustomAttribute1 -like "*Employee*")
                    {
#                        $MsgSubject = "Granted Access to Terminated Employee Mailbox and OneDrive"
                        $MsgSubject = "Granted Access to Terminated Employee Mailbox"
                        $MsgBody = "A request was received to grant access to the below terminated individuals Mailbox and OneDrive files.<ul type=""disc""><li>Terminated " + $Mbx.CustomAttribute1 + " details " + $Mbx.Alias + " - " + $Mbx.DisplayName + "</li><li>Individual given access to this information " + $u.Manager + " - " + (get-mailbox $U.manager).DisplayName + "</li></ul>UL Account Provisioning Team</li></ul>"
                    }
                    else
                    {
                        $MsgSubject = "Granted Access to Terminated Non-Employee Mailbox"
                        $MsBody = "A request was received to grant access to the below terminated individuals Mailbox.<ul type=""disc""><li>Terminated " + $Mbx.CustomAttribute1 + " details " + $Mbx.Alias + " - " + $Mbx.DisplayName + "</li><li>Individual given access to this information " + $u.Manager + " - " + (get-mailbox $U.manager).DisplayName + "</li></ul>UL Account Provisioning Team</li></ul>"
                    }
                    write-host "Sending Message to Legal Team"
                    Send-Message
                }

                If (($mbx.ForwardingAddress -ne $null) -or ($mbx.ForwardingSMTPAddress -ne $null)) 
                {
                    write-host "This Account has a forwarding on the mailbox which will be removed..." -ForegroundColor Red
                    $LineToWrite = $RecordEvent + "FORW" + "`t" + "This Account has a forwarding on the mailbox which will be removed"
                    WriteReportEvent
                    if (($mbx.ForwardingAddress -ne "") -and ($mbx.ForwardingAdress -notlike "steritech.com"))
                    {
                        write-host "Removing ForwardingAddress of " $mbx.Forwardingddress "from this account"
                        $LineToWrite = $RecordEvent + "FORW" + "`t" + "Removing ForwardingAddress of" + $mbx.ForwardingAddress + "from this account"
                        WriteReportEvent
                        set-mailbox $EmpNo -ForwardingAddress $null
                    }
                    if (($mbx.ForwardingSMTPAddress -ne "") -and ($mbx.ForwardingSMTPAddress -notlike "steritech.com"))
                    {
                        write-host "Removing ForwardingSMTPAddress of " $mbx.ForwardingSMTPAddress "from this account"
                        $LineToWrite = $RecordEvent + "FORW" + "`t" + "Removing ForwardingSMTPAddress of" + $mbx.ForwardingSMTPAddress + "from this account"
                        WriteReportEvent
                        set-mailbox $EmpNo -ForwardingSmtpAddress $null
                    }
                }
            }
            else
            {
                write-host "No Mailbox exists for this Account...." -ForegroundColor Red
                $LineToWrite = $RecordEvent + "NOMX" + "`t" + "No Mailbox Exists for this Account"
                WriteReportEvent
            }

            If ($Global:txtADObjProt.Text -ne $True)
            {
     		    If ($Global:chkEmerTerm.Checked -eq "Checked")
                {
                    If ($ADUsr.DistinguishedName -notlike "*Emergency*")
#                    If ($Global:struserPath -notlike "*Emergency*")
                    {
                        write-host "Moving AD Object to the EmergencyDisablement OU..." -ForegroundColor Green
                        Move-ADObject ($ADUsr.DistinguishedName -replace("LDAP://","")) 'OU=EmergencyDisablement,OU=_ExtendedHold,DC=global,DC=ul,DC=com'
                        $LineToWrite = $RecordEvent + "MOVE" + "`t" + "Moving AD Object to the _ExtendedHold\EmergencyDisablement OU"
                        WriteReportEvent
                    }
                }
                else
                {
                    If (($Global:chkDelegateAccess.Checked -eq "Checked") -and ($ADUsr.DistinguishedName -notlike "*MGRAcc*"))
#                    If (($Global:chkDelegateAccess.Checked -eq "Checked") -and ($Global:strUserPath -notlike "*MGRAcc*"))
                    {
	            	    write-host "Moving AD Object to the Disabled\MGRAccessGranted OU..." -ForegroundColor Green
    					Move-ADObject ($ADUsr.DistinguishedName -replace("LDAP://","")) 'OU=MGRAccessGranted,OU=Disabled,DC=global,DC=ul,DC=com'
					    $LineToWrite = $RecordEvent + "MOVE" + "`t" + "Moving AD Object to the Disabled\MGRAccessGranted OU" + "`n"
					    WriteReportEvent
				    }
                    else
   	    			{
                        If ($ADUsr.DistinguishedName -notlike "*Disabled*")
#                        If ($Global:strUserPath -notlike "*Disabled*")
                        {
   	    	    			write-host "Moving AD Object to the Disabled OU..." -ForegroundColor Green
    	    	            Move-ADObject ($ADUsr.DistinguishedName -replace("LDAP://","")) 'OU=Disabled,DC=global,DC=ul,DC=com'
        	        		$LineToWrite = $RecordEvent + "MOVE" + "`t" + "Moving AD Object to the Disabled OU"
		    	    	    WriteReportEvent
       		    		}
                        else
                        {
       					    write-host "AD Object is already in the correct disabled OU..." -ForegroundColor Yellow
	   	    			    $LineToWrite = $RecordEvent + "INFO" + "`t" + "AD Object is already in the correct disabled OU"
	            		    WriteReportEvent
                        }
                    }
                }
            }
            else
            {
            	if ($Global:strUserPath -notlike "*_Ext*")
		   		{
	    			write-host "The AD Account is protected and cannot be moved an email is being sent to the Enterprise Email team..." -ForegroundColor Red
    				$LineToWrite = $RecordEvent + "ERRO" + "`t" + "The AD Account is protected and cannot be moved an sending email to the O365 Admin Team"
		    		WriteReportEvent
				    $MsgAltFrom = New-Object System.Net.Mail.MailAddress "DoNotReply@ul.com" , "Account Provisioning Team"
				    $MsgTo = "EnterpriseMessagingServices@ul.com"
                    $MsgCC = ""
				    $MsgSubject = "Standard Disablement - Move Protected AD Account to _ExtendedHold OU"
				    $MsgBody = "The Active Directory Account for " + $Global:txtInpEmpNo.Text + " is currently protected and in the " + $Global:strUserPath + " container.  The account needs to be moved to the _ExtendedHold OU as a result of the user being terminated"
                    Send-Message
			    }
            }
 
            If (($ADUsr.ExtensionAttribute14 -notlike "*Hold*") -and ($Mbx.LitigationHoldEnabled -eq $False))
#            If (($Global:u.ExtensionAttribute14.value -notlike "*Hold*") -and ($Mbx.LitigationHoldEnabled -eq $False))
#            if ((($Global:u.ExtensionAttribute14.value -notlike "*Hold*") -or ($Global:u.ExtensionAttribute14.value -notlike "*Preservation*") -or ($Global:u.ExtensionAttribute14.value -notlike "*Legal*")) -and ($mbx.LitigationHoldEnabled -eq $False))
            {
                write-host "Updating ExtensionAttribute14 to Standard Termination..." -ForegroundColor Green
                $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Updating ExtensionAttribute14 to " + $DescriptionComment
                WriteReportEvent
                Set-ADUser $Global:txtInpEmpNo.Text -Replace @{'ExtensionAttribute14'=$DescriptionComment}
#                $Global:u.ExtensionAttribute14.value = $DescriptionComment
#                $Global:u.CommitChanges()					
            }
            else
            {
                write-host "Account appears to be on Legal Hold not overwriting value...." -ForegroundColor Red
                $LineToWrite = $RecordEvent + "INFO" + "`t" + "Account appears to be on Legal Hold not overwriting ExtensionAttribute14 value"
                WriteReportEvent
            }

            if ($ADUsr.userCertificate.value -ne $null)
#            if ($Global:u.userCertificate.value -ne $null)
            {
                write-host "Clearing User Certificates..." -ForegroundColor Green
                Set-ADUser $Global:txtInpEmpNo.Text -Replace @{'userCertificate'=$null}
#                $Global:u.userCertificate.value = $null
#                $Global:u.CommitChanges()	
                $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Clearing User Certificates in Active Directory"
                WriteReportEvent            				
            }
            else
            {
                write-host "There are no User Certificates saved for this account no actions necessary...." -ForegroundColor Red
                $LineToWrite = $RecordEvent + "INFO" + "`t" + "No Certificates to clear in Active Directory"
                WriteReportEvent 
            }

            if ($MblDevice -ne $null)
            {
                write-host "This user has the following Mobile Devices Configured" -BackgroundColor Red
                $MblDevice  | ft FriendlyName,DeviceType,DeviceModel,WhenChanged
                Out-File -filepath $ReportFile -append -noClobber -inputObject ($MblDevice | ft FriendlyName,DeviceType,DeviceModel,WhenChanged)

                If ($Global:chkSelWipe.Checked -ne "Checked")
                {
                    write-host "Pleae manually perform selective wipes on these devices" -BackgroundColor Red
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "This account has mobile devices that require selective wipes to be performed"
                    WriteReportEvent
                }
                else
                {
                    write-host "The mobile device checkbox was selected - Removing Mobile Devices from this Account" -BackgroundColor Cyan
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "The mobile device checkbox was selected - Removing Mobile Devices from this Account"
                    foreach ($MblDevice in $MblDevice)
                    {
                        Remove-MobileDevice -Identity $MblDevice.Identity -Confirm:$False
                        write-host "Removing Mobile Device: (Friendly Name: " $MblDevice.FriendlyName ") - (DeviceModel: " $MblDevice.DeviceModel ")"
                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing Mobile Device: (Friendly Name:" + $MblDevice.FriendlyName + ") - (Device Model:" + $MblDevice.DeviceModel + ") from account"
                        WriteReportEvent
                    }
                }
            }
            else
            {
                write-host "No mobile devices are configured for this account"
                $LineToWrite = $RecordEvent + "INFO" + "`t" + "No mobile devices are configured for this account"
                WriteReportEvent
            }

            if ($TeamsUsr."msRTCSIP-UserEnabled" -eq $True)
            {
                Set-ADUser $Global:txtInpEmpNo.Text -Clear msrtcsip-userenabled,msRTCSIP-DeploymentLocator,msRTCSIP-FederationEnabled,msRTCSIP-Line,msRTCSIP-OptionFlags,msRTCSIP-PrimaryHomeServer,msRTCSIP-PrimaryUserAddress,msRTCSIP-UserPolicies
                write-host "Disabled Teams User"
                $LineToWrite = $RecordEvent + "INFO" + "`t" + "Disabled Teams User for this account"
                WriteReportEvent
            }
            else
            {
                write-host "This account was not configured as a Teams User"
                $LineToWrite = $RecordEvent + "INFO" + "`t" + "This account was not configured as a Teams User"
                WriteReportEvent
            }


            write-host "Disablement Process Complete..." -ForegroundColor Green
            $LineToWrite = $RecordEvent + "STOP" + "`t" + "Disablement Process Completed!!"
            WriteReportEvent
        }
        else
        {
            write-host "Request cancelled" -ForegroundColor Red
            $Output = $wshell.Popup("Request cancelled.",0,"Cancelled",0+32)
        }
    	# Write end record to the log file
        $LineToWrite = $RecordEvent + "STOP" + "`t" + "This instance is stopping." + "`n"
    	WriteLogEvent
    }
}
else
{
    write-host "Termination Request cancelled" -ForegroundColor Red
    write-host "Request cancelled" -ForegroundColor Red
    $Output = $wshell.Popup("Request cancelled.",0,"Cancelled",0+32)
#    WriteReportEvent
}
