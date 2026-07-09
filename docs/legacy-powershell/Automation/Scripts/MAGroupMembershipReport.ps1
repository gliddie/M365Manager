#Set forwards on User or Shared Mailboxes
#
#    2025/16/18 - SAG - Report of Groups an individual is a member or owner of
#
#####################################################

Function Build-MAGrpMbrOwn
{
    $Script:form = New-Object Windows.Forms.Form 
    $Script:form.FormBorderStyle = "FixedToolWindow" 
    $Script:form.Text = "Report Group Membership/Ownership"
    $Script:form.StartPosition = "CenterScreen" 
    $Script:form.Width = 500 ; $Script:form.Height = 350  # Make the form wider 

    $Script:Top = 30
    ## Project Name
    $Script:lblProject = New-Object System.Windows.Forms.Label   
        $Script:lblProject.Text = "Project Name:"  
        $Script:lblProject.Top = $Script:Top ; $Script:lblProject.Left = 10; $Script:lblProject.Width=150; $Script:lblProject.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblProject)    # Add to Form 
        #
        $Script:txtProject = New-Object Windows.Forms.TextBox
        $Script:txtProject.TabIndex = 0 # set Tab Order 
        $Script:txtProject.Top = $Script:Top ; $Script:txtProject.Left = 140; $Script:txtProject.Width = 150; $Script:txtProject.AutoSize = $true
        $Script:txtProject.Location = New-Object System.Drawing.Size(140,$Script:Top)
        $Script:txtProject.Size = New-Object system.Drawing.Size(300,40)
        $Global:InputFocus = $Global:txtProject
        $Script:form.Controls.Add($Script:txtProject)    # Add to Form 

    $Script:Top = $Script:Top + 30
    ## Enter Employee #s
    $Script:lblEmpID = New-Object System.Windows.Forms.Label   
        $Script:lblEmpID.Text = "Employee Number(s):"  
        $Script:lblEmpID.Top = $Top ; $Script:lblEmpID.Left = 10; $Script:lblEmpID.Width=150; $Script:lblEmpID.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblEmpID)    # Add to Form 
        # 
        $Script:txtEmpID = New-Object Windows.Forms.TextBox
        $Script:txtEmpID.MaxLength = 2000000
        $Script:txtEmpID.TabIndex = 0 # set Tab Order 
        $Script:txtEmpID.Location = New-Object System.Drawing.Size(140,$Script:Top)
        $Script:txtEmpID.Size = New-Object system.Drawing.Size(300,150)
        $Script:txtEmpID.MultiLine = $true
        $Script:txtEmpID.ScrollBars = 'Both'  
        $Script:form.Controls.Add($Script:txtEmpID)    # Add to Form 
}

$me = whoami
$CredENo = $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))
$UsrName = $CredENo + "@global.ul.com"
$dir = "c:\users\" + $CredENo + "\documents\"
$File = "my" + $CredENo + "File.xml"
$Global:CredFile = $dir + $File
#AdminCredFile
$AFile = "myA" + ($me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))).replace(".","") + "File.xml"
$Global:ACredFile = $dir + $AFile
If (Test-Path $Global:ACredFile)
{
    $Global:AdmLiveCred = Import-Clixml $Global:ACredFile
}

$wshell = New-Object -ComObject Wscript.Shell
Build-MAGrpMbrOwn
Add-FormStandardButtons
$Script:OKButton.Text = "Continue"
Publish-Form

$User = (($Script:txtEmpID.text -replace ("`n",",")).split(",") -replace(" ","")).Trim()

#$EmpNo = read-host "Enter Employee Number"
$FName = "c:\temp\" + ($Script:txtProject.Text).Trim() + ".csv"
If (Test-Path "c:\temp\") {} else {New-Item -Path $FName -ItemType Directory}
#write Column Titles
"EmpNo,StaffName,GroupName,Purpose" |out-file $FName

Foreach ($u in $User)
{
    $GrpMember = ""
    $GrpOwner = ""
    $UPN = $u + "@global.ul.com"
    $ADGroupMem = Get-ADPrincipalGroupMembership $u -ResourceContextServer global.ul.com
 #   $AZGroupOwner = Get-AzureADUserOwnedObject -ObjectId (Get-MgUser -UserId $UPN).Id
    $AZGroupOwner = Get-MgUserOwnedObject -UserId (Get-MgUser -UserId $UPN).Id
    $GrpMember = Get-MgUserMemberOf -UserId $UPN -All
#    $GrpMember = Get-MgUserMemberOf -UserId $UPN -All | Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application"))} |Sort-Object DisplayName
#    $GrpMember = Get-AzureADUser -SearchString $UPN | Get-AzureADUserMembership | Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application"))} |Sort-Object DisplayName
    $MUsr = [bool]($DistName = (get-User $u -ErrorAction SilentlyContinue).DistinguishedName)

    If ($MUsr -eq $True)
    {
        If ($DistName -like "*'*")
        {
	        $DistName = $Distname.Replace("'","")
        }
        $DLOwner = (Get-Recipient -Filter "ManagedBy -eq '$DistName'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup | Select Name,RecipientTypeDetails)
    }

#    Groups a Member Of
#    $GrpMember = $ADGroupMem.Name + $DLMember.DisplayName | Sort-Object | Get-Unique
#    Group Owner Of
    $GrpOwner = $DLOwner.Name + $AZGroupOwner.DisplayName | Sort-Object | Get-Unique

    write-host "`n    ################################################################################"
    write-host $u "- ("(get-Aduser $u).Name "-" ((get-ADUser $u -Properties ExtensionAttribute1).ExtensionAttribute1)")"
    write-host "`tMember of: " $GrpMember.count "groups"
    write-host "`t Owner of: " $GrpOwner.Count "groups"

    foreach ($m in $GrpMember)
    {
        $GrpDet = get-mgGroup -GroupId $m.Id
        $GrpType = "Unknown"
        If ($GrpDet.DisplayName -like "ACL.*")
        {
            $GrpType = "AccessControlGroup"
        }
        elseIf ($GrpDet.DisplayName -like "iProsearch*")
        {
            #Do not report membership in this group
        }
        elseIf ($GrpDet.DisplayName -like "ADMIN.*")
        {
            $GrpType = "AdminGroup"
        }
        elseIf (($GrpDet.DisplayName -like "DSG.*") -or ($GrpDet.DisplayName -like "MFA*") -or ($GrpDet.DisplayName -like "All*"))
        {
            $GrpType = "AzureDynamicGroup"
        }
        If ($GrpDet.DisplayName -like "DST.*")
        {
            $GrpType = "ExchangeDynamicGroup"
        }
        elseIf (($GrpDet.DisplayName -like "GRP.*") -or ($GrpDet.DisplayName -like "GRP*"))
        {
            $GrpType = "MicrosoftTeam"
        }
        elseIf ($GrpDet.DisplayName -like "LIC.*")
        {
            $GrpType = "Licenseing Group"
        }
        elseIf ($GrpDet.DisplayName -like "LIC.*")
        {
            $GrpType = "Licenseing Group"
        }
        elseIf ($GrpDet.DisplayName -like "LST.*")
        {
            $GrpType = "ExchangeDistributionGroup"
        }
        elseIf ($GrpDet.DisplayName -like "*OutOfPolicy*")
        {
            $GrpType = "RoomDelegate"
        }
        elseIf ($GrpDet.DisplayName -like "MBX.*")
        {
            $GrpType = "ExchangeSharedMailboxAccessGroup"
        }
        elseIf ($GrpDet.DisplayName -like "MBX.*")
        {
            $GrpType = "ExchangeSharedMailboxAccessGroup"
        }

        If (($GrpDet.DisplayName -notlike "iProsearch*") -and ($GrpDet.DisplayName -notlike "Domain Users"))
        {
            $text = "{0},""{1}"",{2},""{3}"",{4}" -f $u,(get-Aduser $u).Name,"GrpMember",$GrpDet.DisplayName,$GrpType
#            $text = "{0},""{1}"",{2},""{3}"",{4}" -f $u,(get-Aduser $u).Name,"GrpMember",$m,$GrpType
            $Text
            $Text | out-file $FName -append
        }
    }

    foreach ($o in $GrpOwner)
    {
        $GrpType = "Unknown"
        If ($o -like "ACL.*")
        {
            $GrpType = "AccessControlGroup"
        }
        elseIf ($o -like "iProsearch*")
        {
            #Do not report membership in this group
        }
        elseIf ($o -like "ADMIN.*")
        {
            $GrpType = "AdminGroup"
        }
        elseIf (($o -like "DSG.*") -or ($o -like "MFA*") -or ($o -like "All*"))
        {
            $GrpType = "AzureDynamicGroup"
        }
        If ($o -like "DST.*")
        {
            $GrpType = "ExchangeDynamicGroup"
        }
        elseIf (($o -like "GRP.*") -or ($o -like "GRP*"))
        {
            $GrpType = "MicrosoftTeam"
        }
        elseIf ($o -like "LIC.*")
        {
            $GrpType = "Licenseing Group"
        }
        elseIf ($o -like "LIC.*")
        {
            $GrpType = "Licenseing Group"
        }
        elseIf ($o -like "LST.*")
        {
            $GrpType = "ExchangeDistributionGroup"
        }
        elseIf ($o -like "*OutOfPolicy*")
        {
            $GrpType = "RoomDelegate"
        }
        elseIf ($o -like "MBX.*")
        {
            $GrpType = "ExchangeSharedMailboxAccessGroup"
        }
        elseIf ($o -like "MBX.*")
        {
            $GrpType = "ExchangeSharedMailboxAccessGroup"
        }

        $text = "{0},""{1}"",{2},""{3}"",{4}" -f $u,(get-Aduser $u).Name,"GrpOwner",$o,$GrpType
        $Text
        $Text |out-file $FName -append
    }
}
$Output = $wshell.Popup("Report File can be found on usnbkemes100p\" + $FName,0,"Report File",0+32)