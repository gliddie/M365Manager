#New SDAP Team Member

Function Review-MSOLGroup
{
    If ($Exists -eq $False)
    {
        
        Write-Host "Adding to the $GrpName group" -ForegroundColor Green
        Add-MsolRoleMember -RoleObjectId $AdmGrp -RoleMemberEmailAddress $UPN
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
    $Exists = [bool](Get-ADGroupMember -Identity $GrpName |Where-Object {$_.SamAccountName -eq $ENo})
    If ($Exists -eq $False)
    {
        Write-host "Adding to the $GrpName group" -ForegroundColor Green
        Add-ADGroupMember -Identity $GrpName -Members $ENo
        $LineToWrite = $RecordEvent + "CHNG" + "`t" + "Added to the " + $GrpName + " group"
    }
    else
    {
       $LineToWrite = $RecordEvent + "NOCHG" + "`t" + "Already a memver of the " + $GrpName + " group"
       Write-host "Already a member of the $GrpName" -ForegroundColor Red
    }
    WriteReportEvent
}

$ReportPath = "E:\Automation\NewSDAPTeamMember\Report\" + $Year
If (Test-Path $ReportPath) {} else {mkdir $ReportPath}

$ENo = Read-Host "Enter Employee Number of New SDAP Team Member"
$TicketNo = Read-Host "Enter Ticket Number"
$UPN = $ENo + "@global.ul.com"
$ReportFile = $ReportPath + "\NewSDAPTeamMember-EmpNo" + $ENo + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"

$LineToWrite = $WhoAmI + "`t" + "Adding New SDAP Team Member Started"
WriteReportEvent
$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
WriteReportEvent

$LineToWrite = $RecordEvent + "INFO" + "`t" + "Enabling SDAP Roles for: " + $UPN + " - " + (get-ADUser $ENo).Name
WriteReportEvent
$LineToWrite = $RecordEvent + "INFO" + "`t" + "Ticket Number: " + $TicketNo + "`n"
WriteReportEvent

#Enable RemotePowerShell
$Usr = get-user -Identity $UPN
$LineToWrite = $RecordEvent + "NOCHG" + "`t" + "Remote Powershell already enabled"
If ($Usr.RemotePowerShellEnabled -eq $False)
{
    Set-User -Identity $UPN -RemotePowerShellEnabled $True -Confirm:$False
    $LineToWrite = $RecordEvent + "CHNG" + "`t" + "Enabled Remote Powershell for: " + $UPN
}
WriteReportEvent

#Add to O365 Role Groups
    $TeamCommAdmin = "baf37b3a-610e-45da-9e62-d9d1e5e8914b"
    $ShrPntAdmin = "f28a1f50-f6e7-4571-818b-6a12f2af6b6c"
    $UsrAdmin = "fe930be7-5e62-47db-91af-98c3a49a38b1"

    $Exists = [bool](get-msolrolemember -RoleObjectId $TeamCommAdmin |Where-Object {$_.EmailAddress -eq $upn})
    $AdmGrp = $TeamCommAdmin
    $GrpName = "Teams Communications Administrator"
    Review-MSOLGroup

    $Exists = [bool](get-msolrolemember -RoleObjectId $ShrPntAdmin |Where-Object {$_.EmailAddress -eq $upn})
    $AdmGrp = $ShrPntAdmin
    $GrpName = "SharePoint Administrator"
    Review-MSOLGroup

    $Exists = [bool](get-msolrolemember -RoleObjectId $UsrAdmin |Where-Object {$_.EmailAddress -eq $upn})
    $AdmGrp = $UsrAdmin
    $GrpName = "User Administrator"
    Review-MSOLGroup

#Add to Exchange Role Groups
    $Exists = [bool](Get-RoleGroupMember -Identity "UL Account Provisioning" |Where-Object {$_.PrimarySmtpAddress -eq $Usr.WindowsEmailAddress})
    $AdmGrp = "UL Account Provisioning"
    $GrpName = "UL Account Provisioning"
    Review-ExchGroup

#Add to ADGroups
    $GrpName = "ACL.NBK.Unisys.TermServ"
    Review-ADGroup
    
    $GrpName = "ACL.UL.FileShareAdmins"
    Review-ADGroup

    $GrpName = "ACL.UL.ITUsers"
    Review-ADGroup

    $GrpName = "ADMIN.UL.SDGroupMgr"
    Review-ADGroup

    $GrpName = "ADMIN.UL.ServiceDeskAccountProvisioning"
    Review-ADGroup

    $GrpName = "ADMIN.UL.UNISYS.LYNCPROVISIONER"
    Review-ADGroup

#Add to Priviledged AD Group
    $Exists = [bool](Get-ADGroupMember -Identity "ADMIN.UL.SDAdmin" |Where-Object {$_.SamAccountName -eq $ENo})
    If ($Exists -eq $False)
    {
        $Admin = whoami
        $Admin = $Admin.Replace("global\","global\a")
        $AdminCred = Get-Credential -UserName $Admin -Message 'Enter Admin Account Password'
        write-host "Adding to the ADMIN.UL.SDAdmin group" -ForegroundColor Green
        Add-ADGroupMember -Identity "ADMIN.UL.SDAdmin" -Members $ENo -Credential $AdminCred
        $LineToWrite = $RecordEvent + "CHNG" + "`t" + "Added to the ADMIN.UL.SDAdmin group"
    }
    else
    {
        write-host "Already a member of the ADMIN.UL.SDAdmin group" -ForegroundColor Red
        $LineToWrite = $RecordEvent + "NOCHG" + "`t" + "Already a memver of the ADMIN.UL.SDAdmin group"
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
            Add-DistributionGroupMember "MBX.UL Rejected Messages.RE" -Member $ENo -BypassSecurityGroupManagerCheck
            $LineToWrite = $RecordEvent + "CHNG" + "`t" + "Added to the readers group of the UL Rejected Messages mailbox"
        }
        else
        {
            write-host "Already a member of the UL Rejected Messages shared mailbox readers group" -ForegroundColor Red
            $LineToWrite = $RecordEvent + "NOCHG" + "`t" + "Already a memver of the readers group of the UL Rejected Messages mailbox"
        }
    }
    else
    {
        write-host "This individual does not have a licensed mailbox not adding to the UL Rejected Messages shared mailbox readers group" -ForegroundColor Red
        $LineToWrite = $RecordEvent + "NOCHG" + "`t" + "This individual does not have a licensed mailbox not adding to the readers group of the UL Rejected Messages mailbox"
    }
    WriteReportEvent