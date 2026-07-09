#New SDAP Team Member
#   01/11/2024 - SAG - Modified to have a form in front to enter the EmpNo and TicketNo.
#   06/19/2025 - SAG - Modified the directory role group changes to reaplce the MSOL commands with Graph
#   06/19/2025 - SAG - Remove adding SDAP members to ACL.UL.FileShareAdmins this is no longer permitted
#   03/20/2026 - SAG working to resolve the errors when adding individuals into a new role group

function Enter-EmpNoInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow"
    $Global:form.Text = "New SDAP Team Member"
    $Global:form.Size = New-Object System.Drawing.Size(450,175) #(W,H)
    $Global:form.StartPosition = "CenterScreen"
    Add-FormStandardButtons
    ## Employee Number
    $Script:lblEmpNo = New-Object System.Windows.Forms.Label   
        $Script:lblEmpNo.Text = "Employee No.:"  
        $Script:lblEmpNo.Top = 20 ; $Script:lblEmpNo.Left = 10; $Script:lblEmpNo.Width=120 ; $Script:lblEmpNo.AutoSize = $true
        $Global:form.Controls.Add($Script:lblEmpNo)    # Add to Form 
        # 
        $Script:txtEmpNo = New-Object Windows.Forms.TextBox  
        $Script:txtEmpNo.Top = 20; $Script:txtEmpNo.Left = 140; $Script:txtEmpNo.Width = 200;  
        $Script:txtEmpNo.Text = ""
        $Global:form.Controls.Add($Script:txtEmpNo)    # Add to Form
        $Script:InputFocus = $Script:txtEmpNo

    ## Ticket Number
    $Script:lblTicketNo = New-Object System.Windows.Forms.Label   
        $Script:lblTicketNo.Text = "Ticket No.:"  
        $Script:lblTicketNo.Top = 50 ; $Script:lblTicketNo.Left = 10; $Script:lblTicketNo.Width=120 ; $Script:lblTicketNo.AutoSize = $true
        $Global:form.Controls.Add($Script:lblTicketNo)    # Add to Form 
        # 
        $Script:txtTicketNo = New-Object Windows.Forms.TextBox  
        $Script:txtTicketNo.Top = 50; $Script:txtTicketNo.Left = 140; $Script:txtTicketNo.Width = 200;  
        $Script:txtTicketNo.Text = ""
        $Global:form.Controls.Add($Script:txtTicketNo)    # Add to Form
}

Function Review-MgGroup
{
#these commands are successful when running in an interactive PS window. 3/20/2026 SAG
    If ($Exists -eq $False)
    {
        Write-Host "Adding to the $GrpName group" -ForegroundColor Green
#        Add-MsolRoleMember -RoleObjectId $AdmGrp -RoleMemberEmailAddress $UPN
        $newRoleMember = @{
          "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UsrObjId"
        }
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $AdmGrp -BodyParameter $newRoleMember
        $LineToWrite = $RecordEvent + "ADDED" + "`t" + "Added to the " + $GrpName + " group"
    }
    else
    {
        Write-host "Already a member of the $GrpName group" -ForegroundColor Red
        $LineToWrite = $RecordEvent + "NOCHG" + "`t" + "Already a member of the " + $GrpName + " group"
    }
    WriteReportEvent
}

Function Review-ExchGroup
{
    If ($Exists -eq $False)
    {
        
        Write-Host "Adding to the $GrpName group" -ForegroundColor Green
        Add-RoleGroupMember -Identity $AdmGrp -Member $UPN
        $LineToWrite = $RecordEvent + "ADDED" + "`t" + "Added to the " + $GrpName + " group"
    }
    else
    {
        Write-host "Already a member of the $GrpName group" -ForegroundColor Red
        $LineToWrite = $RecordEvent + "NOCHG" + "`t" + "Already a member of the " + $GrpName + " group"
    }
    WriteReportEvent
}

Function Review-ADGroup
{
    $Exists = [bool](Get-ADGroupMember -Identity $GrpName |Where-Object {$_.SamAccountName -eq $Script:txtEmpNo.Text})
    If ($Exists -eq $False)
    {
        Write-host "Adding to the $GrpName group" -ForegroundColor Green
        Add-ADGroupMember -Identity $GrpName -Members $Script:txtEmpNo.Text
        $LineToWrite = $RecordEvent + "CHNG" + "`t" + "Added to the " + $GrpName + " group"
    }
    else
    {
       Write-host "Already a member of the $GrpName group" -ForegroundColor Red
       $LineToWrite = $RecordEvent + "NOCHG" + "`t" + "Already a member of the " + $GrpName + " group"
       Write-host "Already a member of the $GrpName" -ForegroundColor Red
    }
    WriteReportEvent
}

$ReportPath = "E:\Automation\NewSDAPTeamMember\Report\" + $Year
If (Test-Path $ReportPath) {} else {mkdir $ReportPath}

#$ENo = Read-Host "Enter Employee Number of New SDAP Team Member"
#$TicketNo = Read-Host "Enter Ticket Number"
Enter-EmpNoInputForm
$Global:okButton.Text = "Proceed"
Publish-Form

If ($Global:Result -eq "OK")
{

    $UPN = $Script:txtEmpNo.Text + "@global.ul.com"
    $ReportFile = $ReportPath + "\NewSDAPTeamMember-EmpNo" + $Script:txtEmpNo.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"

    Write-host "Starting changes for New SDAP Team Member: " $UPN -ForegroundColor Cyan
    $LineToWrite = $WhoAmI + "`t" + "Adding New SDAP Team Member Started"
    WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
    WriteReportEvent

    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Enabling SDAP Roles for: " + $UPN + " - " + (get-ADUser $Script:txtEmpNo.Text).Name
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Ticket Number: " + $Script:txtTicketNo.Text + "`n"
    WriteReportEvent

    #Enable RemotePowerShell
    $Usr = get-user -Identity $UPN
    If ($Usr.RemotePowerShellEnabled -eq $False)
    {
        write-host "Enabling Powershell for: $UPN" -ForegroundColor Green
        Set-User -Identity $UPN -RemotePowerShellEnabled $True -Confirm:$False
        $LineToWrite = $RecordEvent + "CHNG" + "`t" + "Enabled Remote Powershell for: " + $UPN
    }
    else
    {
        $LineToWrite = $RecordEvent + "NOCHG" + "`t" + "Remote Powershell already enabled"
        write-host "Powershell already enabled for: $UPN" -ForegroundColor Red
    }
    WriteReportEvent

#Add to O365 Role Groups
#    $TeamCommAdmin = "baf37b3a-610e-45da-9e62-d9d1e5e8914b" #MSOL
    $TeamCommAdmin = "cd9676a1-0a61-43c3-b830-d69cea9852c9"
#    $ShrPntAdmin = "f28a1f50-f6e7-4571-818b-6a12f2af6b6c" #MSOL
    $ShrPntAdmin = "42d29543-439e-4084-938c-eec9155690f3"
#    $UsrAdmin = "fe930be7-5e62-47db-91af-98c3a49a38b1" #MSOL
    $UsrAdmin = "7c8c4c85-c351-438a-8574-f389eef15a21"
    $UsrObjID = (Get-MgUser -UserID $UPN).ID

#    $Exists = [bool](get-msolrolemember -RoleObjectId $TeamCommAdmin |Where-Object {$_.EmailAddress -eq $upn})
    $Exists = [bool](get-MgDirectoryRole -DirectoryRoleId $TeamCommAdmin -ExpandProperty Members |Where-Object {$_.Members.Id -eq $UsrObjID})
    $AdmGrp = $TeamCommAdmin
    $GrpName = "Teams Communications Administrator"
    Review-MgGroup

#    $Exists = [bool](get-msolrolemember -RoleObjectId $ShrPntAdmin |Where-Object {$_.EmailAddress -eq $upn})
    $Exists = [bool](get-MgDirectoryRole -DirectoryRoleId $ShrPntAdmin -ExpandProperty Members |Where-Object {$_.Members.Id -eq $usrObjId})
    $AdmGrp = $ShrPntAdmin
    $GrpName = "SharePoint Administrator"
    Review-MgGroup

#    $Exists = [bool](get-msolrolemember -RoleObjectId $UsrAdmin |Where-Object {$_.EmailAddress -eq $upn})
    $Exists = [bool](get-MgDirectoryRole -DirectoryRoleId $UsrAdmin -ExpandProperty Members |Where-Object {$_.Members.Id -eq $usrObjId})
    $AdmGrp = $UsrAdmin
    $GrpName = "User Administrator"
    Review-MgGroup

#Add to Exchange Role Groups
    $Exists = [bool](Get-RoleGroupMember -Identity "UL Account Provisioning" |Where-Object {$_.PrimarySmtpAddress -eq $Usr.WindowsEmailAddress})
    $AdmGrp = "UL Account Provisioning"
    $GrpName = "UL Account Provisioning"
    Review-ExchGroup

#Add to ADGroups
    $GrpName = "ACL.NBK.Unisys.TermServ"
    Review-ADGroup
    
#    $GrpName = "ACL.UL.FileShareAdmins"
#    Review-ADGroup

    $GrpName = "ACL.UL.ITUsers"
    Review-ADGroup

    $GrpName = "ADMIN.UL.SDGroupMgr"
    Review-ADGroup

    $GrpName = "ADMIN.UL.ServiceDeskAccountProvisioning"
    Review-ADGroup

    $GrpName = "ADMIN.UL.UNISYS.LYNCPROVISIONER"
    Review-ADGroup

#Add to Priviledged AD Group
    $Exists = [bool](Get-ADGroupMember -Identity "ADMIN.UL.SDAdmin" |Where-Object {$_.SamAccountName -eq $Script:txtEmpNo.Text})
    If ($Exists -eq $False)
    {
#        $Admin = whoami
#        $Admin = $Admin.Replace("global\","global\a")
#        $AdminCred = Get-Credential -UserName $Admin -Message 'Enter Admin Account Password'
        write-host "Adding to the ADMIN.UL.SDAdmin group" -ForegroundColor Green
        Add-ADGroupMember -Identity "ADMIN.UL.SDAdmin" -Members $Script:txtEmpNo.Text -Credential $Global:AdmLiveCred
#        Add-ADGroupMember -Identity "ADMIN.UL.SDAdmin" -Members $Script:txtEmpNo.Text -Credential $AdminCred
        $LineToWrite = $RecordEvent + "CHNG" + "`t" + "Added to the ADMIN.UL.SDAdmin group"
    }
    else
    {
        write-host "Already a member of the ADMIN.UL.SDAdmin group" -ForegroundColor Red
        $LineToWrite = $RecordEvent + "NOCHG" + "`t" + "Already a member of the ADMIN.UL.SDAdmin group"
    }
    WriteReportEvent

#Add to the Access group for the UL Rejected Messages mailbox
    $Licen = [bool](Get-Mailbox $upn -ErrorAction SilentlyContinue)
    If ($Licen -eq $True)
    {
        $Exists = [bool](Get-DistributionGroupMember "MBX.UL Rejected Messages.RE" | Where-Object {$_.PrimarySmtpAddress -eq $Usr.WindowsEmailAddress})
        If ($Exists -eq $False)
        {
            write-host "Adding to the UL Rejected Messages shared mailbox readers group" -ForegroundColor Green
            Add-DistributionGroupMember "MBX.UL Rejected Messages.RE" -Member $Script:txtEmpNo.Text -BypassSecurityGroupManagerCheck
            $LineToWrite = $RecordEvent + "CHNG" + "`t" + "Added to the readers group of the UL Rejected Messages mailbox"
        }
        else
        {
            write-host "Already a member of the UL Rejected Messages shared mailbox readers group" -ForegroundColor Red
            $LineToWrite = $RecordEvent + "NOCHG" + "`t" + "Already a member of the readers group of the UL Rejected Messages mailbox"
        }
    }
    else
    {
        write-host "This individual does not have a licensed mailbox not adding to the UL Rejected Messages shared mailbox readers group" -ForegroundColor Red
        $LineToWrite = $RecordEvent + "NOCHG" + "`t" + "This individual does not have a licensed mailbox not adding to the readers group of the UL Rejected Messages mailbox"
    }
    WriteReportEvent
}
else
{
    write-host "Process Aborted" -ForegroundColor Red
}