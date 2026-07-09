#Default form buttons
function Add-FormStandardButtons
{
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    $Script:buttonPanel = New-Object Windows.Forms.Panel  
    $buttonPanel.Size = New-Object Drawing.Size @(400,40) 
    $buttonPanel.Dock = "Bottom"    
    $Script:cancelButton = New-Object Windows.Forms.Button  
        $Script:cancelButton.Top = $buttonPanel.Height - $Script:cancelButton.Height - 10; $Script:cancelButton.Left = $buttonPanel.Width - $Script:cancelButton.Width - 10 
        $Script:cancelButton.TabIndex = 198
        $Script:cancelButton.Text = "Cancel" 
        $Script:cancelButton.DialogResult = "Cancel" 
        $Script:cancelButton.Anchor = "Right"
    ## Create the OK button, which will anchor to the left of Cancel 
    $Script:okButton = New-Object Windows.Forms.Button   
        $Script:okButton.Top = $cancelButton.Top ; $Script:okButton.Left = $cancelButton.Left - $Script:okButton.Width - 10
        $Script:okButton.TabIndex = 97
        $Script:okButton.Text = $Action
        If ($Script:OKDetails -ne "")
        {
            $Script:okButton.Text = $Script:OKDetails
            $Script:OKDetails = ""
        }
        else
        {
            $Script:okButton.Text = "Continue"
        }
        $Script:okButton.DialogResult = "OK" 
        $Script:okButton.Anchor = "Right"
    ## Add the buttons to the button panel 
    $Script:buttonPanel.Controls.Add($Script:okButton) 
    $Script:buttonPanel.Controls.Add($Script:cancelButton) 
    ## Add the button panel to the form 
    $Script:form.Controls.Add($buttonPanel)
    ## Set Default actions for the buttons 
    $Script:form.AcceptButton = $Script:okButton          # ENTER = ok 
    $Script:form.CancelButton = $Script:cancelButton      # ESCAPE = Cancel
}

Function Publish-Form
{
    If ($Script:InputFocus -eq $null)
    {
        $Script:InputFocus = $Script:okButton
    }
    ## Finalize Form and Show Dialog
    $Script:form.Add_Shown( { $Script:form.Activate(); $Script:InputFocus.Focus() } )  #Activate and Set Focus 
#    $Global:form.Add_Shown( { $form.Activate(); $Global:okButton.Focus() } )  #Activate and Set Focus 
    $Script:Result = $Script:form.ShowDialog()          ## Show the form, and wait for the response
}

Function Build-MAOwnerMemberCleanUp
{
    $Script:form = New-Object Windows.Forms.Form 
    $Script:form.FormBorderStyle = "FixedToolWindow" 
    $Script:form.Text = "Remove Owner/Member Roles"
    $Script:form.StartPosition = "CenterScreen" 
    $Script:form.Width = 500 ; $Script:form.Height = 380  # Make the form wider 

    $Script:Top = 20
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
        $Script:InputFocus = $Script:txtProject
        $Script:form.Controls.Add($Script:txtProject)    # Add to Form       
    
    $Script:Top = $Script:Top + 30
    $Col1 = 30
    $Col2 = 190
    $Col3 = 320
    ## Project Name
    $Script:lblSelect = New-Object System.Windows.Forms.Label   
        $Script:lblSelect.Text = "Select Groups Types to Remove From:"  
        $Script:lblSelect.Top = $Script:Top ; $Script:lblSelect.Left = 10; $Script:lblSelect.Width=150; $Script:lblSelect.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblSelect)    # Add to Form

    $Script:Top = $Script:Top + 25
    ## Remove from LST Groups
    $Script:chkLSTGrp = New-Object Windows.Forms.CheckBox
        $Script:chkLSTGrp.Left = $Col1; $Script:chkLSTGrp.Width = 100; $Script:chkLSTGrp.Top = $Script:Top
        $Script:chkLSTGrp.Text = "LST Groups" 
        $Script:chkLSTGrp.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkLSTGrp)

    ## Remove from MBX Groups
    $Script:chkMBXGrp = New-Object Windows.Forms.CheckBox
        $Script:chkMBXGrp.Left = $Col2; $Script:chkMBXGrp.Width = 100; $Script:chkMBXGrp.Top = $Script:Top
        $Script:chkMBXGrp.Text = "MBX Groups" 
        $Script:chkMBXGrp.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkMBXGrp)

    ## Remove from ACL Groups
    $Script:chkACLGrp = New-Object Windows.Forms.CheckBox
        $Script:chkACLGrp.Left = $Col3; $Script:chkACLGrp.Width = 200; $Script:chkACLGrp.Top = $Script:Top
        $Script:chkACLGrp.Text = "ACL Groups" 
        $Script:chkACLGrp.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkACLGrp)

    $Script:Top = $Script:Top + 25
    ## Remove from Teams Groups
    $Script:chkGRPGrp = New-Object Windows.Forms.CheckBox
        $Script:chkGRPGrp.Left = $Col1; $Script:chkGRPGrp.Width = 160; $Script:chkGRPGrp.Top = $Script:Top
        $Script:chkGRPGrp.Text = "GRP (Teams) Groups" 
        $Script:chkGRPGrp.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkGRPGrp)

    ## Remove from Admin Groups
    $Script:chkADMGrp = New-Object Windows.Forms.CheckBox
        $Script:chkADMGrp.Left = $Col2; $Script:chkADMGrp.Width = 100; $Script:chkADMGrp.Top = $Script:Top
        $Script:chkADMGrp.Text = "ADMIN Groups" 
        $Script:chkADMGrp.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkADMGrp)

    ## Remove from All Groups
    $Script:chkALLGrp = New-Object Windows.Forms.CheckBox
        $Script:chkALLGrp.Left = $Col3; $Script:chkALLGrp.Width = 200; $Script:chkALLGrp.Top = $Script:Top
        $Script:chkALLGrp.Text = "All Groups" 
        $Script:chkALLGrp.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkALLGrp)

    $Script:Top = $Script:Top + $Col1
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
        $Script:txtEmpID.add_Click({
            $Script:OKButton.Visible = $True
        })         
}

Function Cleanup_OwnMemDL
{
    foreach ($o in $DLOwner)
    {
        $Mgrs = (Get-DistributionGroup $O.Name).ManagedBy
        If ($Mgrs.count -gt 1)
        {
            Set-DistributionGroup $O.Name -ManagedBy @{remove="$u"} -BypassSecurityGroupManagerCheck -confirm:$False -ErrorAction SilentlyContinue
            write-host "`tRemoved as owner of group: " $O.Name -ForegroundColor Green
            $LineToWrite = $RecordEvent + "INFO" + "`t`tRemoved as an owner of group: " + $O.Name
        }
        else
        {
            write-host "`tThis person is the last owner of the group: " $o.Name -ForegroundColor Red
            $LineToWrite = $RecordEvent + "INFO" + "`t`tUnable to remove from ownership this person is the last owner of group: " + $O.Name
        }
        WriteReportEvent
    }
                        
    Foreach ($m in $DLMember)
    {
        Remove-DistributionGroupMember $m.DisplayName -Member $u -BypassSecurityGroupManagerCheck -confirm:$False -ErrorAction SilentlyContinue
        write-host "`tRemoved from the Membership of: " $m.DisplayName
        $LineToWrite = $RecordEvent + "INFO" + "`t`tRemoved as member of group: " + $m.DisplayName
        WriteReportEvent
    }
}

Function CleanUp_OwnMemAD
{
    foreach ($o in $DLOwner)
    {
        $ADZExists = ""
        $ADGExists = ""
        $LineToWrite = ""
        $ErrorActionPreference = 'SilentlyContinue'
        $ADZExists = [bool]($ADZ = Get-AzureADGroup -ObjectId $O.ExternalDirectoryObjectId)
        $ADGExists = [bool]($ADG = Get-ADGroup $O.Name)
        $ErrorActionPreference = 'Continue'

        If ($ADExists -eq $True)
        {
            Remove-AzureADGroupOwner -ObjectId $O.ExternalDirectoryObjectId -OwnerId (get-user $u).ExternalDirectoryObjectId
            write-host "`tRemoved as a ower of : " $o.Name -ForegroundColor Green
            $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removed as an owner of group: " + $O.Name
        }
        else
        {
            If ($ADGExists -eq $True)
            {
               write-host "Need code to remove someone as an owner of the AD group" -ForegroundColor Red
               $LineToWrite = $RecordEvent + "ERR " + "`t" + "Need code to remove someone as an owner of an AD group: " + $O.Name
            }
        }
        WriteReportEvent
    }

    Foreach ($m in $DLMember)
    {
        $ADZExists = ""
        $ADGExists = ""
        $LineToWrite = ""
        $ErrorActionPreference = 'SilentlyContinue'
        $ADZExists = [bool]($ADZ = Get-AzureADGroup -ObjectId $m.ObjectId)
        $ADGExists = [bool]($ADG = Get-ADGroup $m.DisplayName)
        $ErrorActionPreference = 'Continue'
        
        If ($ADZExists -eq $True)
        {
            $MemCnt = (Get-AzureADGroupMember -objectid $ADZ.Objectid).count
            $ErrorActionPreference = 'SilentlyContinue'
            $Success = [bool](Remove-AzureADGroupMember -objectid $ADZ.Objectid -MemberId (get-user $u).ExternalDirectoryObjectId)
            $ErrorActionPreference = 'Continue'
            write-host "`tRemoved as a member of: " $m.DisplayName -ForegroundColor Green
            $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removed as a member of AzureADGroup: " + $m.DisplayName
        }
        else
        {
            If ($ADGExists -eq $True)
            {
                Remove-ADGroupMember -Identity $m.DisplayName -Members $u -Confirm:$False
                write-host "`tRemoved as a member of: " $m.DisplayName -ForegroundColor Green
                $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removed as a member of ADGroup: " + $m.DisplayName
            }
        }
        WriteReportEvent
    }
}

Function Cleanup_OwnMemUNI
{
    foreach ($o in $DLOwner)
    {
        $Mgrs = (Get-DistributionGroup $O.Name).ManagedBy
        If ($Mgrs.count -gt 1)
        {
            Set-DistributionGroup $O.Name -ManagedBy @{remove="$u"} -BypassSecurityGroupManagerCheck -confirm:$False -ErrorAction SilentlyContinue
            write-host "`tRemoved as owner of group: " $O.Name -ForegroundColor Green
            $LineToWrite = $RecordEvent + "INFO" + "`t`tRemoved as an owner of group: " + $O.Name
        }
        else
        {
            write-host "`tThis person is the last owner of the group: " $o.Name -ForegroundColor Red
            $LineToWrite = $RecordEvent + "INFO" + "`t`tUnable to remove from ownership this person is the last owner of group: " + $O.Name
        }
        WriteReportEvent
    }
                        
    Foreach ($m in $DLMember)
    {
        $MemCnt = (Get-UnifiedGroup $m.DisplayName).GroupMemberCount
        If ($MemCnt -gt 1)
        {
            Remove-UnifiedGroupLinks $m.ObjectID -LinkType Member -Links $u -Confirm:$False
            write-host "`tRemoved from the Membership of: " $m.DisplayName
            $LineToWrite = $RecordEvent + "INFO" + "`t`tRemoved as member of group: " + $m.DisplayName
        }
        else
        {
            write-host "`tUnable to remove this individual is the last member of: " $m.DisplayName
            $LineToWrite = $RecordEvent + "INFO" + "`t`tUnable to remove this individual is the last member of: " + $m.DisplayName
        }
        WriteReportEvent
    }
}

Function Start-Report
{
    $LineToWrite = $RecordEvent + "STAR" + "`t" + "Generating Report of Group Ownership/Membership has started"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + (whoami)
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "User TimeZone: " + (Get-TimeZone)
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Number of Users to Process: " + $addMember.count
    WriteReportEvent
    $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
    WriteReportEvent
}

$me = whoami

$wshell = New-Object -ComObject Wscript.Shell
Build-MAOwnerMemberCleanUp
Add-FormStandardButtons
$Script:OKButton.Text = "Continue"
Publish-Form

Do
{
    If ($Script:Result -eq "OK")
    {
        If ($Script:txtEmpID.Text.Length -gt 0)
        {
            $ReportFile = "E:\Automation\MAActivities\Reports\RemoveOwnerMember\RemoveOwnerMember-" + $Script:txtProject.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            write-host $ReportFile
            $addMember = (($Script:txtEmpID.text -replace ("`n",",")).split(",") -replace(" ","")).Trim()
            Start-Report
            write-Host "Number of Accounts to Process: " $addMember.Count
            foreach ($u in $addMember)
            {
                If ($u.length -gt 0)
                {
                    $ErrorActionPreference = "SilentlyContinue"
                    $Exists = [bool](get-ADUser $u)
                    $ErrorActionPreference = "Continue"
                    If ($Exists -eq $True)
                    {
                        write-host "`nProcessing UserID: " $u "(" (get-ADUser $u).Name ")"
                        $LineToWrite = $RecordEvent + "CURR" + "`t" + "Employee ID: " + $u
                        WriteReportEvent

                        $DistName = (get-User $u).DistinguishedName
                        If ($DistName -like "*'*")
                        {
                            $DistName = $Distname.Replace("'","")
                        }

                        If (($Script:chkLSTGrp.Checked -eq $True) -or ($Script:chkALLGrp.Checked -eq $True))
                        {
                            $DLOwner = (Get-Recipient -Filter "ManagedBy -eq '$DistName'" -RecipientTypeDetails MailUniversalDistributionGroup | Where-Object {$_.Name -like "LST*"} | Select Name,RecipientTypeDetails) |Sort-Object Name
                            $DLMember = Get-AzureADUser -SearchString $u | Get-AzureADUserMembership -All $True| Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application") -and ($_.DisplayName -like "LST*"))} |Sort-Object DisplayName
                            Cleanup_OwnMemDL
                        }

                        If (($Script:chkMBXGrp.Checked -eq $True) -or ($Script:chkALLGrp.Checked -eq $True))
                        {
                            $DLOwner = (Get-Recipient -Filter "ManagedBy -eq '$DistName'" -RecipientTypeDetails MailUniversalDistributionGroup | Where-Object {$_.Name -like "MBX*"} | Select Name,RecipientTypeDetails) |Sort-Object Name
                            $DLMember = Get-AzureADUser -SearchString $u | Get-AzureADUserMembership -All $True| Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application") -and ($_.DisplayName -like "MBX*"))} |Sort-Object DisplayName
                            Cleanup_OwnMemDL
                        }

                        If (($Script:chkACLGrp.Checked -eq $True) -or ($Script:chkALLGrp.Checked -eq $True))
                        {
                            $DLOwner = (Get-Recipient -Filter "ManagedBy -eq '$DistName'" -RecipientTypeDetails MailUniversalDistributionGroup | Where-Object {$_.Name -like "ACL*"} | Select Name,RecipientTypeDetails,ExternalDirectoryObjectId) |Sort-Object Name
                            $DLMember = Get-AzureADUser -SearchString $u | Get-AzureADUserMembership -All $True| Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application") -and ($_.DisplayName -like "ACL*"))} |Sort-Object DisplayName
                            Cleanup_OwnMemAD
                        }

                        If (($Script:chkUNIGrp.Checked -eq $True) -or ($Script:chkALLGrp.Checked -eq $True))
                        {
                            $DLOwner = Get-AzureADUserOwnedObject -ObjectId ((get-AzureADUser -SearchString ($u + "@global.ul.com")).ObjectId) | Where-object {(($_.ObjectType -eq "Group") -and ($_.DisplayName -like "GRP*"))}
                            $DLMember = Get-AzureADUser -SearchString $u | Get-AzureADUserMembership -All $True| Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application") -and ($_.DisplayName -like "GRP*"))} |Sort-Object DisplayName
                            Cleanup_OwnMemUNI
                        }

                        ($Script:chkALLGrp.Checked -eq $True)
                        {
                            $DLOwner = (Get-Recipient -Filter "ManagedBy -eq '$DistName'" -RecipientTypeDetails MailUniversalDistributionGroup | Select Name,RecipientTypeDetails,ExternalDirectoryObjectId) |Sort-Object Name
                            $DLMember = Get-AzureADUser -SearchString $u | Get-AzureADUserMembership -All $True| Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application") )} |Sort-Object DisplayName
                            write-host "Need extra code to handle the remaining groups.  Some groups cannot be changed because the membership of the group is managed based on a rule.  For others they do not follow our naming standard and may require additional code to be developed." -ForegroundColor DarkRed
                        }

                        $DLOwner = Get-AzureADUserOwnedObject -ObjectId ((get-AzureADUser -SearchString ($u + "@global.ul.com")).ObjectId) | Where-object {(($_.ObjectType -eq "Group") -and ($_.DisplayName -like "GRP*"))} |Sort-Object Name
                        write-host "This individual remains an owner of $DLOwner.count groups" -ForegroundColor Cyan
                        
                        $DLMember = Get-AzureADUser -SearchString $u | Get-AzureADUserMembership -All $True| Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application"))} |Sort-Object DisplayName
                        write-host "This individual remains a member of $DLMember.count groups" -ForegroundColor Cyan
                    }
                    else
                    {
                        write-host "`nAccount Does Not Exist: "$u -ForegroundColor Red
                        $LineToWrite = $RecordEvent + "CURR" + "`t" + "      No AD Account Found for: " + $u
                        WriteReportEvent
                    }
                }
            }
        }
        else
        {
            $output = $wshell.Popup("No employee number or email addresses entered try again.",0,"No Employee Nos.",0+32)
        }
        $Script:txtEmpID.Text = ""
        Publish-Form
    }
}While ($Script:Result -eq "OK")

<#
Function Cleanup-MemberOwner
{
    $ENo = $Global:txtInpEmpNo.Text
    $EmpID = $ENo + "@global.ul.com"
    $ObjectId = (Get-Msoluser -UserPrincipalName $EmpID).ObjectID
    $Usr = (get-mailbox $ENo).Name
    $ADUsrName = (Get-ADUser $ENo).Name
    $NeedOwner = "E:\O365AdminShared\Data\SDAPDLGroupOwnerMember.csv"

    $DistName = (get-User $ENo).DistinguishedName
    If ($DistName -like "*'*")
    {
        $DistName = $Distname.Replace("'","")
    }
    $DLOwner = (Get-Recipient -Filter "ManagedBy -eq '$DistName'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup | Select Name,DisplayName,RecipientTypeDetails) | Sort-Object RecipientTypeDetails,DisplayName
    If ($DLOwner.count -gt 0)
    {
        write-host "Number of groups this person is an owner of: " $DLOwner.count -ForegroundColor Cyan
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Number of groups this person is an owner of: " + $DLOwner.count
    }
    else
    {
        write-host "This person is not an owner of any Azure/AD or O365 groups" -ForegroundColor Red
        $LineToWrite = $RecordEvent + "OWNR" + "`t" + "This person is not an owner of any Azure/AD or O365 groups"
    }
    WriteReportEvent

    $DLMember = Get-AzureADUser -SearchString $ENo | Get-AzureADUserMembership | Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application"))} |Sort-Object DisplayName
    If ($DLMember.count -gt 0)
    {
        write-host "Number of groups this person is a member of: " $DLMember.count -ForegroundColor Cyan
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Number of groups this person is a member of: " + $DLMember.count
    }
    else
    {
        write-host "This person is not a member of any Azure/AD or O365 groups" -ForegroundColor Red
        $LineToWrite = $RecordEvent + "MEMB" + "`t" + "This person is not a member of any Azure/AD or O365 groups"
    }
    WriteReportEvent

    #Review what groups this person an an Owner Of
    If ($DLOwner.count -gt 0)
    {
        foreach ($O in $DLOwner)
        {
            $ADZExists = ""
            $ADGExists = ""
            $DLExists = ""
            $UniExists = ""
            $LineToWrite = ""

            $ErrorActionPreference = 'SilentlyContinue'
            $ADZExists = [bool]($ADZ = Get-AzureADGroup -SearchString $O.DisplayName)
            $ADGExists = [bool]($ADG = Get-ADGroup $O.DisplayName)
            $ErrorActionPreference = 'Continue'
            $DLExists = [bool]($DL = Get-DistributionGroup $O.DisplayName -ErrorAction SilentlyContinue)
            $UniExists = [bool]($Uni = Get-UnifiedGroup $O.DisplayName -ErrorAction SilentlyContinue)
            $Reason = "NoOwners"

            If ($DLExists -eq $True)
            {
                $Type = "DL-O365"
                $Mgrs = (Get-DistributionGroup $O.Name).ManagedBy
            }
            else
            {
                If ($ADGExists -eq $True)
                {
                    $Type = "ADG"
                }
                else
                {
                    If ($UniExists -eq $True)
                    {
                        $Type = "UNI"
                        $Mgrs = (Get-UnifiedGroup $O.Name).ManagedBy
                    }
                    else
                    {
                        If ($ADZExists -eq $True)
                        {
                            $Type = "AZAD"
                        }
                    }
                }
            }

            If ($UniExists -eq $False)
            {
                If ($DLExists -eq $True)
                {
                    If (($O.DisplayName -notlike "DST.ULMAIL*") -and ($O.DisplayName -notlike "DSG.UL*") -and ($O.DisplayName -ne "RCC-Users") -and ($O.DisplayName -notlike "LIC.*"))
                    {
                        If ($Mgrs.count -gt 1)
                        {
                            Set-DistributionGroup $O.Name -ManagedBy @{remove="$ENo"} -BypassSecurityGroupManagerCheck -confirm:$False -ErrorAction SilentlyContinue
                            write-host "`tRemoved as owner of $Type group: " $O.DisplayName -ForegroundColor Green
                        }
                        else
                        {
                            write-host "This person is the last owner of the $Type group: " $O.DisplayName -ForegroundColor Red
                            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unable to remove from ownership this person is the last owner of " + $Type + " group: " + $O.DisplayName
                            $Reason = "LastOwner"
                            $MemCnt = (get-DistributionGroupMember $O.Name -ResultSize Unlimited).Count
                        }
                    }
                }
                else
                {
                    write-host "No code to remove owner from this group type: $Type group: " $O.DisplayName -ForegroundColor Red
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "No Code to Remove Owner from this group type: " + $Type + " group: " + $O.DisplayName           
                }
            }
            else
            {
                If (($UniExists -eq $True) -and ($0.DisplayName -notlike "DST.*") -and ($O.DisplayName -notlike "DSG*") -and ($0.DisplayName -ne "iProSearch") -and ($0.DisplayName -ne "RCC-Users" )-and ($O.DisplayName -notlike "LIC.*"))
                {
                    $MemCnt = (Get-UnifiedGroup $O.Name).GroupMemberCount
                    If ($Mgrs.count -gt 1)
                    {
                        Remove-UnifiedGroupLinks $O.Name -LinkType Owners -Links $ENo -Confirm:$False
                        $MemCnt = (Get-UnifiedGroup $O.Name).GroupMemberCount
                        write-host "`tRemoved as owner of $Type group: " $O.DisplayName -ForegroundColor Green
                        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Removed from ownership this " + $Type + " group: " + $O.DisplayName
                    }
                    else
                    {
                         write-host "This person is the last owner of the $Type group: " $O.DisplayName -ForegroundColor Red
                        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unable to remove from ownership this person is the last owner of " + $Type + " group: " + $O.DisplayName
                        $Reason = "LastOwner"
                    }
                }
                else
                {
                    If (($Type -eq "DL-O365") -or ($Type -eq "UNI"))
                    {
                        write-host "This is a Dynamic group and ownership is not managed" $O.DisplayName -ForegroundColor Red
                    }
                }
            }

        #Print details for removal from group ownership
            If ($LineToWrite -ne "")
            {
                WriteReportEvent
                $Reason = "LastOwner"
            }

        #Checking to see if this person is also a member of the group
            If (($O.Displayname -notlike "DST.ULMAIL*") -and ($Reason -ne "LastOwner"))
            {
                If ($UniExists -eq $False)
                {
                    If ($DLExists -eq $True)
                    {
                        if (($O.DisplayName -notlike "DSG*") -and ($O.DisplayName -notlike "DST*") -and ($O.DisplayName -notlike "MFA_*") -and ($O.DisplayName -ne "RCC-Users") -and ($O.DisplayName -notlike "LIC.*"))
                        {
                            $IsMember = [bool]((Get-DistributionGroupMember $O.DisplayName -ResultSize Unlimited) | ?{$_.name -like $usr})
                            If ($IsMember -eq $True)
                            {
                                Remove-DistributionGroupmember $O.DisplayName -Member $ENo -BypassSecurityGroupManagerCheck -Confirm:$False -ErrorAction SilentlyContinue
                                $MemCnt = (get-DistributionGroupMember $O.DisplayName -ResultSize Unlimited).Count
                                $Text = "{0},{1}" -f $ENo,$O.DisplayName
                            }
                        }
                    }
                    else
                    {
                        If ($ADGExists -eq $True)
                        {
                            write-host "This is an AD group need code" -ForegroundColor Red
                        }
                        else
                        {
                            If ($ADZExists -eq $True)
                            {
                                $MemCnt = (Get-AzureADGroupMember -objectid $ADZ.Objectid).count
                                if (($O.DisplayName -notlike "DSG*") -and ($O.DisplayName -notlike "DST*") -and ($O.DisplayName -notlike "MFA_*") -and ($O.DisplayName -ne "RCC-Users") -and ($O.DisplayName -notlike "LIC.*"))
                                {
                                    $ErrorActionPreference = 'SilentlyContinue'
                                    Remove-AzureADGroupMember -objectid $ADZ.Objectid -MemberId (get-user $ENo).ExternalDirectoryObjectId
                                    $ErrorActionPreference = 'Continue'
                                }
                            }
                        }
                    }
                }
                else
                {
                    If ($UniExists -eq $True)
                    {
                        $MbrCnt = (Get-UnifiedGroup $O.ObjectID).GroupMemberCount
                        If ($MbrCnt -gt 1)
                        {            
                            Remove-UnifiedGroupLinks $O.ObjectID -LinkType Member -Links $ENo -Confirm:$False
                        } 
                    }
                }

                If (($MemCnt -eq "") -or ($MemCnt -eq $null))
                {
                    $MemCnt = 0
                    If ($IsMember -eq $True)
                    {
                        $Reason = $Reason + "LastMember"
                    }
                }

            #Print details for removal from group ownership
                If ($LineToWrite -eq "")
                {
                    If ($IsMember -eq $True)
                    {
                        write-host "`tRemoved as an owner and a member of $Type group: " $O.DisplayName -ForegroundColor Green
                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removed as an owner and a member of " + $Type + " group: " + $O.DisplayName
                    }
                    else
                    {
                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removed as an owner of " + $Type + " group: " + $O.DisplayName
                    }
                    WriteReportEvent
                }

                If (($O.DisplayName -notlike "DSG*") -and ($O.DisplayName -notlike "DST*") -and ($O.DisplayName -notlike "MFA_*") -and ($O.DisplayName -ne "RCC-Users") -and ($O.DisplayName -notlike "LIC.*"))
                {
                    If (($Mgrs.count -le 1) -or ($MemCnt -eq 0)) #Print details to file to be used for sending notificaitons and updating mailtip details.
                    {
                        $text = """{0}"",""{1}"",{2},{3},{4},{5},{6},{7},{8}" -f $O.DisplayName,$O.Name,$Type,$Reason,$Mgrs.Count,$MemCnt,$ENo,$ADUsrName,$Global:txtInpTicketNo.Text
                        Out-File -FilePath $NeedOwner -InputObject $text -Append
                    }
                }
            }
        }
    }

    #Remove from groups where the individual is identified as a group member
    $DLMember = Get-AzureADUser -SearchString $ENo | Get-AzureADUserMembership | Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application"))} |Sort-Object DisplayName
    If ($DLMember.count -gt 0)
    {
        foreach ($O in $DLMember)
        {
            $ADZExists = ""
            $ADGExists = ""
            $DLExists = ""
            $UniExists = ""
            $Mgrs = ""
            $LineToWrite = ""

            $ErrorActionPreference = 'SilentlyContinue'
            $ADZExists = [bool](Get-AzureADGroup -SearchString $O.DisplayName)
            $ADGExists = [bool]($ADG = Get-ADGroup $O.DisplayName) 
            $ErrorActionPreference = 'Continue'
            $DLExists = [bool]($DL = Get-DistributionGroup $O.DisplayName -ErrorAction SilentlyContinue)
            $UniExists = [bool]($Uni = Get-UnifiedGroup $O.DisplayName -ErrorAction SilentlyContinue)
            $Reason = "NoMembers"

            If ($DLExists -eq $True)
            {
                $Type = "DL-O365"
                $Mgrs = (Get-DistributionGroup $O.DisplayName).ManagedBy
            }
            else
            {
                If ($ADGExists -eq $True)
                {
                    $Type = "ADG"
                }
                else
                {
                    If ($UniExists -eq $True)
                    {
                        $Type = "UNI"
                        $Mgrs = (Get-UnifiedGroup $O.DisplayName).ManagedBy
                    }
                    else
                    {
                        If ($ADZExists -eq $True)
                        {
                            $Type = "AZAD"
                        }
                    }
                }
            }

            If (($O.DisplayName -like "DSG*") -or ($O.DisplayName -like "DST*") -or ($O.DisplayName -like "MFA_*") -or ($O.DisplayName -eq "RCC-Users") -or ($O.DisplayName -like "LIC.*"))
            {
                write-host "`tMembership of this group is based on a rule: " $O.DisplayName -ForegroundColor Cyan
                $LineToWrite = $RecordEvent + "REMV" + "`t" + "Membership of this group is based on a rule: " + $O.DisplayName
            }
            else
            {
                If ($Mgrs -eq ((get-mailbox $ENo).Name))
                {
                    If  ($Mgrs.count -le 1)
                    {
                        write-host "`tCannot remove this person is the last owner of this group: " $O.DisplayName -ForegroundColor Red
                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Cannot remove ownership this person is the last owner of this group: " + $O.DisplayName
                    }
                    else
                    {
                        Set-DistributionGroup $O.DisplayName -ManagedBy @{remove="$ENo"} -BypassSecurityGroupManagerCheck -confirm:$False -ErrorAction SilentlyContinue
                        write-host "`tRemoved as owner of $Type group: " $O.DisplayName -ForegroundColor Green
                    }
                    WriteReportEvent
                }

                If ($DLExists -eq $True)
                {
                    $MemCnt = (get-DistributionGroupMember $O.DisplayName -ResultSize Unlimited).Count
                    If (($0.DisplayName -notlike "DST.UL*") -and ($0.DisplayName -ne "iProSearch") -and ($DL.RecipientTypeDetails -like "MailUniversal*"))
                    {
                        If ($O.DisplayName -eq "iProSearch")
                        {
                            write-host "`tGroup is used for Legal Hold Searches - Membership not updated: " $O.DisplayName -ForegroundColor Cyan
                            $LineToWrite = $RecordEvent + "REMV" + "`t" + "Group is used for Legal Hold Searches - Membership not updated: " + $O.DisplayName
                        }
                        else
                        {
                            $ErrorActionPreference = 'SilentlyContinue'
                            Remove-DistributionGroupMember $O.DisplayName -Member $ENo -BypassSecurityGroupManagerCheck -Confirm:$False -ErrorAction SilentlyContinue
                        }
                    }
                    else
                    {
                        If ($O.DisplayName -eq "iProSearch")
                        {
                            write-host "`tGroup is used for Legal Hold Searches - Membership not updated: " $O.DisplayName -ForegroundColor Cyan
                            $LineToWrite = $RecordEvent + "REMV" + "`t" + "Group is used for Legal Hold Searches - Membership not updated: " + $O.DisplayName
                        }
                        else
                        {
                            write-host "`tMembership of this group is not managed in O365: " $O.DisplayName -ForegroundColor Cyan
                            $LineToWrite = $RecordEvent + "REMV" + "`t" + "Membership of this group is not managed in O365: " + $O.DisplayName
                        }
                    }
                }
                else
                {
                    If ($ADGExists -eq $True)
                    {
                        $group = Get-ADGroup -Identity $O.Displayname -Properties member
                        $members = @()
                        $members = $group.member
                        $MemCnt = $members.count
                        If (($O.DisplayName -notlike "DSG*") -and ($O.DisplayName -notlike "DST*") -and ($O.DisplayName -notlike "MFA_*") -and ($O.DisplayName -ne "RCC-Users"))
                        {
   		                    $ErrorActionPreference = 'SilentlyContinue'
                            Remove-ADGroupMember $O.DisplayName -Members $ENo -Confirm:$false
                            $ErrorActionPreference = 'Continue'
                        }
                    }
                    else
                    {
                        If ($UniExists -eq $True)
                        {
                            $MemCnt = (Get-UnifiedGroup $O.DisplayName).GroupMemberCount
                            If ($MemCnt -gt 1)
                            {
                                If (($O.DisplayName -like "DSG*") -or ($O.DisplayName -like "DST*") -or ($O.DisplayName -like "MFA_*") -or ($O.DisplayName -eq "RCC-Users"))
#                                if (($O.DisplayName -notlike "DSG*") -and ($O.DisplayName -notlike "DST*") -and ($O.DisplayName -notlike "MFA_*") -and ($O.DisplayName -notlike "LIC*"))
                                {
                                    write-host "`tMembership of this group is based on a rule: " $O.DisplayName -ForegroundColor Cyan
                                    $LineToWrite = $RecordEvent + "REMV" + "`t" + "Membership of this group is based on a rule: " + $O.DisplayName
                                }
                                else
                                {
#                                    Remove-UnifiedGroupLinks $O.DisplayName -LinkType Member -Links $ENo -Confirm:$False
                                    Remove-UnifiedGroupLinks $O.ObjectID -LinkType Member -Links $ENo -Confirm:$False
                                }
                            }  
                            else
                            {
                                write-host "Membership of this group is based on a rule: " $O.DisplayName -ForegroundColor Cyan
                                $LineToWrite = $RecordEvent + "REMV" + "`t" + "Membership of this group is based on a rule: " + $O.DisplayName
                            }
                        }
                        else
                        {
                            If ($ADZExists -eq $True)
                            {
                                $MemCnt = (Get-AzureADGroupMember -objectid $O.Objectid).count
                                If (($O.DisplayName -like "DSG*") -or ($O.DisplayName -like "DST*") -or ($O.DisplayName -like "MFA_*") -or ($O.DisplayName -eq "RCC-Users") -or ($O.DisplayName -eq "LIC.*"))
#                                if (($O.DisplayName -notlike "DSG*") -and ($O.DisplayName -notlike "DST*") -and ($O.DisplayName -notlike "MFA_*") -and ($O.DisplayName -notlike "LIC*"))
                               {
                                    If ($O.DisplayName -like "LIC*")
                                    {
                                        write-host "`tMembership of is managed this group is managed in other scripts: " $O.DisplayName -ForegroundColor Cyan
                                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Membership of this group is managed in other scripts: " + $O.DisplayName
                                    }
                                    else
                                    {
                                        write-host "`tMembership of this group is based on a rule: " $O.DisplayName -ForegroundColor Cyan
                                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Membership of this group is based on a rule: " + $O.DisplayName
                                    }
                                }
                                else
                                {
                                    $ErrorActionPreference = 'SilentlyContinue'
                                    Remove-AzureADGroupMember -objectid $O.Objectid -MemberId (get-user $ENo).ExternalDirectoryObjectId
                                    $ErrorActionPreference = 'Continue'
                                }
                            }
                            else
                            { 
                                write-host "`tUnable to remove from membership group cannot be managed or was not found: " $O.DisplayName -ForegroundColor Red
                                $LineToWrite = $RecordEvent + "REMV" + "`t" + "Unable to remove from membership group cannot be managed or was not found: " + $O.DisplayName
                            }
                        }
                    }
                }

                If ($LineToWrite -eq "")
                {
                    write-host "`tRemoved as member of $Type group: " $O.DisplayName -ForegroundColor Green
                    $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removed as member of " + $Type + " group: " + $O.DisplayName
                }
            }
            WriteReportEvent

            If (($O.DisplayName -notlike "DSG*") -and ($O.DisplayName -notlike "DST*") -and ($O.DisplayName -notlike "MFA_*") -and ($O.DisplayName -ne "RCC-Users"))
            {
                If (($MemCnt -eq "") -or ($MemCnt -eq $null) -or ($MemCnt -le 1))
                {
                    $text = """{0}"",""{1}"",{2},{3},{4},{5},{6},{7},{8}" -f $O.DisplayName,$O.Name,$Type,"NoMembers",$Mgrs.Count,$MemCnt,$ENo,$ADUsrName,$Global:txtInpTicketNo.Text
                    Out-File -FilePath $NeedOwner -InputObject $text -Append
                }
            }
        }
    }

}
#>