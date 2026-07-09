function Enter-EmpNoInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow"
    $Global:form.Text = "Employee No."
    $Global:form.Size = New-Object System.Drawing.Size(450,175) #(W,H)
    $Global:form.StartPosition = "CenterScreen"
    Add-FormStandardButtons
    ## Employee Number
    $Global:lblEmpNo = New-Object System.Windows.Forms.Label   
        $Global:lblEmpNo.Text = "Employee No.:"  
        $Global:lblEmpNo.Top = 30 ; $Global:lblEmpNo.Left = 10; $Global:lblEmpNo.Width=120 ; $Global:lblEmpNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblEmpNo)    # Add to Form 
        # 
        $Global:txtInpEmpNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpEmpNo.Top = 30; $Global:txtInpEmpNo.Left = 140; $Global:txtInpEmpNo.Width = 200;  
        $Global:txtInpEmpNo.Text = ""
        $Global:form.Controls.Add($Global:txtInpEmpNo)    # Add to Form
        $Global:InputFocus = $Global:txtInpEmpNo

    ## Ticket Number
    $Global:lblTicketNo = New-Object System.Windows.Forms.Label   
        $Global:lblTicketNo.Text = "Ticket No.:"  
        $Global:lblTicketNo.Top = 60 ; $Global:lblTicketNo.Left = 10; $Global:lblTicketNo.Width=120 ; $Global:lblTicketNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTicketNo)    # Add to Form 
        # 
        $Global:txtInpTicketNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTicketNo.Top = 60; $Global:txtInpTicketNo.Left = 140; $Global:txtInpTicketNo.Width = 200;  
        $Global:txtInpTicketNo.Text = ""
        $Global:form.Controls.Add($Global:txtInpTicketNo)    # Add to Form
}

#############################


$me = whoami
$AAct = "A" + ($me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)))
$Pwd = read-host "Enter password for" $AAct "account" -AsSecureString
$Server = "usnbkadds006p.global.ul.com"
$cred= New-Object System.Management.Automation.PSCredential($AAct,$Pwd)

Enter-EmpNoInputForm
Publish-Form

If ($Global:Result -eq "OK")
{
    $uDate          = (get-date -uformat %D).Replace("/", "-")
    $uTime          = (get-date -uformat %T).Replace(":", "")
    $ReportFile	    = "E:\Automation\SDAPTeamMembers\Report-SDAPTeamMembers-EmpNo" + $Global:txtInpEmpNo.Text + "-Date" + $uDate + "-Time" + $uTime + ".log"
    $ADUser = get-ADUser $Global:txtInpEmpNo.Text
    $upn = $Global:txtInpEmpNo.Text + "@global.ul.com"

    $LineToWrite = $RecordEvent + "STAR" + "`t" + "EnableNewSDAPMember script has started"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
    WriteReportEvent
    write-host "`nStarting SDAP Enablement Process for: " $Global:txtInpEmpNo.Text " " $ADUser.Name -ForegroundColor Cyan
    $LineToWrite = $RecordEvent + "INFO" + "`tStarting Enablement Process for: " + $Global:txtInpEmpNo.Text + " - " + $ADUser.Name
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`tTicket Number: " + $Global:txtInpTicketNo.Text
    WriteReportEvent

    Set-User $upn -RemotePowerShellEnabled $True -Confirm:$False
    $LineToWrite = $RecordEvent + "ENAB" + "`tEnabled Remote Powershell for this user"
    WriteReportEvent

    #Adding to O365 Role groups
    Add-MsolRoleMember -RoleMemberEmailAddress $upn -RoleName "User Administrator"
    $LineToWrite = $RecordEvent + "ENAB" + "`tEnabled User Administrator Role in O365"
    WriteReportEvent
    Add-MsolRoleMember -RoleMemberEmailAddress $upn -RoleName "Teams Communications Administrator"
    $LineToWrite = $RecordEvent + "ENAB" + "`tEnabled Teams Communications Administrator Role in O365"
    WriteReportEvent
    Add-MsolRoleMember -RoleMemberEmailAddress $upn -RoleName "SharePoint Administrator"
    $LineToWrite = $RecordEvent + "ENAB" + "`tEnabled SharePoint Administrator Role in O365"
    WriteReportEvent
    Add-RoleGroupMember -Identity "UL Account Provisioning" -Member $upn
    $LineToWrite = $RecordEvent + "ENAB" + "`tEnabled UL Account Provisioning Role in O365"
    WriteReportEvent

    #Adding to non-priviledged AD group
    Add-ADGroupMember ACL.UL.IntuneServiceDesk -Members $Global:txtInpEmpNo.Text
    $LineToWrite = $RecordEvent + "ENAB" + "`tAdded to the ACL.UL.IntuneServiceDesk Group in AD"
    WriteReportEvent

    #These three commands need to be run with Admin Creds as the groups are priviledged grups
    Add-ADGroupMember ADMIN.UL.SDAdmin -Members $Global:txtInpEmpNo.Text -Credential $Cred -Server $Server
    $LineToWrite = $RecordEvent + "ENAB" + "`tAdded to the ADMIN.UL.SDAdmin Group in AD"
    WriteReportEvent
    Add-ADGroupMember ADMIN.UL.ServiceDeskAccountProvisioning -Members $Global:txtInpEmpNo.Text -Credential $Cred -Server $Server
    $LineToWrite = $RecordEvent + "ENAB" + "`tAdded to the ADMIN.UL.ServiceDeskAccountProvisioning Group in AD"
    WriteReportEvent
    Add-ADGroupMember ADMIN.UL.UNISYS.LYNCPROVISIONER -Members $Global:txtInpEmpNo.Text -Credential $Cred -Server $Server
    $LineToWrite = $RecordEvent + "ENAB" + "`tAdded to the ADMIN.UL.UNISYS.LYNCPROVISIONER Group in AD" + "`n"
    WriteReportEvent

    $LineToWrite = $RecordEvent + "INFO" + "`tEnablement Process Complete for: " + $Global:txtInpEmpNo.Text + " - " + $ADUser.Name
    WriteReportEvent
}