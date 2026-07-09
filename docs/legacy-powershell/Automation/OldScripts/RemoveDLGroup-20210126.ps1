<#
#
#  Called by:  DistributionSecurityGroupMenu.ps1
#              RoomResourceAdminMenu.ps1
#
#  06/14/2018 - Added to to delete Unified Groups
#  08/04/2018 - Added code if there are no members in a DST group to remove without confirming
#  01/28/2019 - Added -ResultSize Unlimited to the get-DistributionListMember statement
#  04/08/2019 - Added -ResultSize Unlimited to the line that initially reports how many individuals in the group
#  12/10/2019 - Added code so if the name does not follow the naming standard the requestor identifies the group type
#  08/08/2020 - Changed over to a GUI Interface
#>

function Build-DLInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Remove Distribution Group" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(450,250) #(W,H)
    $BldDetails = "N"

    ## Get Details used to create
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Distribution List Name:"
        $Global:lblDispName.Top = 10 ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtInpName = New-Object Windows.Forms.TextBox  
        $Global:txtInpName.TabIndex = 0 # set Tab Order 
        $Global:txtInpName.Top = 10; $Global:txtInpName.Left = 130; $Global:txtInpName.Width = 280;  
        $Global:txtInpName.Text = ""   # DisplayName
        $Global:txtInpName.TabIndex = 0
        $Global:form.Controls.Add($Global:txtInpName)    # Add to Form 
        $Global:InputFocus = $Global:txtInpName
 
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 40 ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtInpTaskNo.Top = 40; $Global:txtInpTaskNo.Left = 130; $Global:txtInpTaskNo.Width = 120;  
        $Global:txtInpTaskNo.Text = "TASK"   # Enter ticket number
        $Global:txtInpTaskNo.TabIndex = 1
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form

    ## DL Group         
    $Global:chkDLGroup = New-Object Windows.Forms.checkbox 
        $Global:chkDLGroup.Left = 130; $Global:chkDLGroup.Width = 200; $Global:chkDLGroup.Top = 70  
        $Global:chkDLGroup.Text = "Distribution List" 
        $Global:chkDLGroup.Checked = $true   # set a default value 
        $Global:chkDLGroup.TabIndex = 2
        $Global:form.Controls.Add($Global:chkDLGroup) 
        # Obtain Value with: $Global:chkDLGroup.Checked

    ## Dynamic Group
    $Global:chkDynGroup = New-Object Windows.Forms.checkbox 
        $Global:chkDynGroup.Left = 130; $Global:chkDynGroup.Width = 200; $Global:chkDynGroup.Top = 90  
        $Global:chkDynGroup.Text = "Dynamic Distribution List" 
        $Global:chkDynGroup.Checked = $Global:chkDynGroup.Checked   # set a default value 
        $Global:chkDynGroup.TabIndex = 4
        $Global:form.Controls.Add($Global:chkDynGroup) 
        # Obtain Value with: $Global:chkDynGroup.Checked

    ## O365 Group
    $Global:chkO365Group = New-Object Windows.Forms.checkbox 
        $Global:chkO365Group.Left = 130; $Global:chkO365Group.Width = 200; $Global:chkO365Group.Top = 110  
        $Global:chkO365Group.Text = "O365 Group" 
        $Global:chkO365Group.Checked = $Global:chkO365Group.Checked   # set a default value 
        $Global:chkO365Group.TabIndex = 5
        $Global:form.Controls.Add($Global:chkO365Group) 
        # Obtain Value with: $Global:chkO365Group.Checked

        Add-FormStandardButtons
}

function Build-DLDetailsForm
{
    $BldDetails = "N"
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Remove Distribution Group Details" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(600,400) #(W,H)

    ## Get Details used to create
    ## Display Name
    $LocTop = 10
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Distribution List Name:"
        $Global:lblDispName.Top = $LocTop ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtDLDet = New-Object Windows.Forms.TextBox  
        $Global:txtDLDet.TabIndex = 0 # set Tab Order
        $Global:txtDLDet.Top = $LocTop; $Global:txtDLDet.Left = 130; $Global:txtDLDet.Width = 400;  
        $Global:txtDLDet.Text = $Global:txtInpName.Text
        $Global:txtDLDet.TabIndex = 1
        $Global:form.Controls.Add($Global:txtDLDet)    # Add to Form 
         
    $LocTop = $LocTop + 30
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = $LocTop ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtDetTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtDetTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtDetTaskNo.Top = $LocTop; $Global:txtDetTaskNo.Left = 130; $Global:txtDetTaskNo.Width = 400;
        $Global:txtDetTaskNo.Text = $Global:txtInpTaskNo.Text 
        If ($Global:txtDetTaskNo.Text.Length -lt 7)
        {
            $Global:txtDetTaskNo.Text = "NoOwners/NoMembers"
        }
        $Global:txtDetTaskNo.TabIndex = 2 
        $Global:form.Controls.Add($Global:txtDetTaskNo)    # Add to Form

    $LocTop = $LocTop + 30
    ## DL Notes
    $Global:lblDLNotes = New-Object System.Windows.Forms.Label   
        $Global:lblDLNotes.Text = "Notes:"  
        $Global:lblDLNotes.Top = $LocTop ; $Global:lblDLNotes.Left = 10; $Global:lblDLNotes.Width=120 ; $Global:lblDLNotes.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLNotes)    # Add to Form 
        # 
        $Global:txtInpDLNotes = New-Object Windows.Forms.TextBox  
        $Global:txtInpDLNotes.TabIndex = 0 # set Tab Order 
        $Global:txtInpDLNotes.Top = $LocTop; $Global:txtInpDLNotes.Left = 130; $Global:txtInpDLNotes.Width = 400;  
        $Global:txtInpDLNotes.TabIndex = 3
        $Global:txtInpDLNotes.Text = (get-group $Global:txtInpName.Text).Notes
        $Global:form.Controls.Add($Global:txtInpDLNotes)    # Add to Form

    $LocTop = $LocTop + 30
    ## DL Owners
    $Global:lblDLOwner = New-Object System.Windows.Forms.Label   
        $Global:lblDLOwner.Text = "Group Owner(s):"  
        $Global:lblDLOwner.Top = $LocTop ; $Global:lblDLOwner.Left = 10; $Global:lblDLOwner.Width=120 ; $Global:lblDLOwner.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLOwner)    # Add to Form 
        # 
        $Global:txtInpDLOwner = New-Object Windows.Forms.TextBox  
        $Global:txtInpDLOwner.TabIndex = 0 # set Tab Order 
        $Global:txtInpDLOwner.Top = $LocTop; $Global:txtInpDLOwner.Left = 130; $Global:txtInpDLOwner.Width = 400;  
        $Global:txtInpDLOwner.TabIndex = 4
        $Global:txtInpDLOwner.Text = $RemDL.ManagedBy
        $Global:form.Controls.Add($Global:txtInpDLOwner)    # Add to Form

    $LocTop = $LocTop + 30
    ## DL Members Count
    $Global:lblDLMemCnt = New-Object System.Windows.Forms.Label   
        $Global:lblDLMemCnt.Text = "Number of Member(s):"  
        $Global:lblDLMemCnt.Top = $LocTop ; $Global:lblDLMemCnt.Left = 10; $Global:lblDLMemCnt.Width=120 ; $Global:lblDLMemCnt.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLMemCnt)    # Add to Form 
        # 
        $Global:txtInpDLMemCnt = New-Object Windows.Forms.TextBox  
        $Global:txtInpDLMemCnt.TabIndex = 0 # set Tab Order 
        $Global:txtInpDLMemCnt.Top = $LocTop; $Global:txtInpDLMemCnt.Left = 130; $Global:txtInpDLMemCnt.Width = 400;  
        $Global:txtInpDLMemCnt.TabIndex = 4
        $Global:txtInpDLMemCnt.Text = (Get-DistributionGroupMember $Global:txtInpName.Text).count
        $Global:form.Controls.Add($Global:txtInpDLMemCnt)    # Add to Form

    $LocTop = $LocTop + 30
    ## DL Members
    $Global:lblDLMem = New-Object System.Windows.Forms.Label   
        $Global:lblDLMem.Text = "Group Member(s):"  
        $Global:lblDLMem.Top = $LocTop ; $Global:lblDLMem.Left = 10; $Global:lblDLMem.Width=120 ; $Global:lblDLMem.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLMem)    # Add to Form 
        # 
        $Global:txtInpDLMem = New-Object Windows.Forms.TextBox  
        $Global:txtInpDLMem.TabIndex = 0 # set Tab Order 
        $Global:txtInpDLMem.Location = New-Object System.Drawing.Size(130,$LocTop)
        $Global:txtInpDLMem.Size = New-Object system.Drawing.Size(400,120)
        $Global:txtInpDLMem.MultiLine = $true
        $Global:txtInpDLMem.ScrollBars = 'Both'
        $Global:txtInpDLMem.TabIndex = 6
        $Global:txtInpDLMem.Text = (Get-DistributionGroupMember $Global:txtInpName.Text) -join ", "
        $Global:form.Controls.Add($Global:txtInpDLMem)    # Add to Form

        Add-FormStandardButtons
}

function Build-DynDetailsForm
{
    $BldDetails = "N"
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Remove Dynamic Distribution Group Details" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(600,400) #(W,H)

    ## Get Details used to create
    ## Display Name
    $LocTop = 10
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Dynamic List Name:"
        $Global:lblDispName.Top = $LocTop ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtDLDet = New-Object Windows.Forms.TextBox  
        $Global:txtDLDet.TabIndex = 0 # set Tab Order
        $Global:txtDLDet.Top = $LocTop; $Global:txtDLDet.Left = 130; $Global:txtDLDet.Width = 400;  
        $Global:txtDLDet.Text = $Global:txtInpName.Text
        $Global:txtDLDet.TabIndex = 1
        $Global:form.Controls.Add($Global:txtDLDet)    # Add to Form
         
    $LocTop = $LocTop + 30
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = $LocTop ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtDetTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtDetTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtDetTaskNo.Top = $LocTop; $Global:txtDetTaskNo.Left = 130; $Global:txtDetTaskNo.Width = 400;
        $Global:txtDetTaskNo.Text = $Global:txtInpTaskNo.Text 
        If ($Global:txtDetTaskNo.Text.Length -lt 7)
        {
            $Global:txtDetTaskNo.Text = "NoOwners/NoMembers"
        }
        $Global:txtDetTaskNo.TabIndex = 2 
        $Global:form.Controls.Add($Global:txtDetTaskNo)    # Add to Form

    $LocTop = $LocTop + 30
    ## DL Owners
    $Global:lblDLOwner = New-Object System.Windows.Forms.Label   
        $Global:lblDLOwner.Text = "Group Owner(s):"  
        $Global:lblDLOwner.Top = $LocTop ; $Global:lblDLOwner.Left = 10; $Global:lblDLOwner.Width=120 ; $Global:lblDLOwner.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLOwner)    # Add to Form 
        # 
        $Global:txtInpDLOwner = New-Object Windows.Forms.TextBox  
        $Global:txtInpDLOwner.TabIndex = 0 # set Tab Order 
        $Global:txtInpDLOwner.Top = $LocTop; $Global:txtInpDLOwner.Left = 130; $Global:txtInpDLOwner.Width = 400;  
        $Global:txtInpDLOwner.TabIndex = 4
        $Global:txtInpDLOwner.Text = $RemDL.ManagedBy
        $Global:form.Controls.Add($Global:txtInpDLOwner)    # Add to Form

    $LocTop = $LocTop + 30
    ## DLMembers
    $Global:lblDLMemCnt = New-Object System.Windows.Forms.Label   
        $Global:lblDLMemCnt.Text = "Group Member(s):"  
        $Global:lblDLMemCnt.Top = $LocTop ; $Global:lblDLOwner.Left = 10; $Global:lblDLOwner.Width=120 ; $Global:lblDLOwner.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLMemCnt)    # Add to Form 
        # 
        $Global:txtInpDLMemCnt = New-Object Windows.Forms.TextBo00;  
        $Global:txtInpDLMemCnt.TabIndex = 4
        $Global:txtInpDLMemCnt.Text = $colu.count
        $Global:form.Controls.Add($Global:txtInpDLMemCnt)    # Add to Form

        Add-FormStandardButtons
}

Build-DLInputForm
Publish-Form

If ($Global:Result -eq "OK")
{
    $Reason = "Reason for Removal: " + $Global:txtInpTaskNo.Text

    If (($Global:chkDLGroup.Checked -eq "Checked") -or ($Global:txtInpName.Text -like "LST*"))
    {
        $DstExists = [bool](Get-DistributionGroup $Global:txtInpName.Text -ErrorAction SilentlyContinue)
        If ($DstExists -eq "True")
        {
            $RemDL = Get-DistributionGroup $Global:txtInpName.Text
            Build-DLDetailsForm
            $Global:InputFocus = $Global:txtDLDet
            Publish-Form

            If ($Global:Result -eq "OK")
            {
                $OutFileName = "c:\temp\" + ($Global:txtInpName.Text -replace " ","") + ".txt"
                $Output = $wshell.Popup("Removing group " + $Global:txtInpName.Text + " copy file " + $OutfileName + " to the Team Sharepoint Site",0,"Move Log File",0+32)
                Write-Output "Removing Distribution List" > $OutFileName
                Write-Output $Reason >> $OutFileName
                Write-Output "Distribution Group Information" >> $OutFileName
                get-DistributionGroup $Global:txtInpName.Text >> $OutFileName
                Write-Output "Distribution Group Detailed Information" >> $OutFileName
                get-DistributionGroup $Global:txtInpName.Text |fl >> $OutFileName
                Write-Output "Distribution Group Manager and Notes" >> $OutFileName
                get-Group $Global:txtInpName.Text |fl ManagedBy,Notes >> $OutFileName
                Write-Output "Distribution Group Membership" >> $OutFileName
                get-DistributionGroupMember $Global:txtInpName.Text -ResultSize Unlimited|ft Alias,Name,RecipientType >> $OutFileName
                remove-DistributionGroup $Global:txtInpName.Text -BypassSecurityGroupManagerCheck -Confirm:$False
            }
        }
        else
        {
            $Output = $wshell.Popup("Distribution List " + $Global:txtInpName.Text + " does not exist",0,"No Changes Needed",0+32)
        }
    }

    If (($Global:chkDynGroup.Checked -eq "Checked") -or ($Global:txtInpName.Text -like "DST*"))
    {
        $DstExists = [bool](Get-DynamicDistributionGroup $Global:txtInpName.Text -ErrorAction SilentlyContinue)
        If ($DstExists -eq "True")
        {
            Remove-PSSession (Get-PSSession)
            $Global:LiveCred = Import-Clixml $Global:CredFile
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
            Import-PSSession $Session -AllowClobber

            $g = Get-DynamicDistributionGroup $Global:txtInpName.Text
            write-host "Getting the number of dynamic group members...."
            $colu = Get-Recipient -RecipientPreviewFilter $g.LdapRecipientFilter -ResultSize Unlimited
            Build-DynDetailsForm
            Publish-Form

            If ($Global:Result -eq "OK")
            {
                $Confirm = "Y"
                If ($colu.count -ne 0)
                {
                    Write-host "There are still members in this group should this be deleted (Y/N)? " -ForegroundColor cyan -NoNewline
                    $Confirm = Read-Host
                }

                If ($Confirm -eq "Y")
                {
        `           $OutFileName = "c:\temp\" + ($Global:txtInpName.Text -replace " ","") + ".txt"
                    $Output = $wshell.Popup("Removing group " + $Global:txtInpName.Text + " copy file " + $OutfileName + " to the Team Sharepoint Site",0,"Move Log File",0+32)
                    Write-Output "Removing Dynamic Distribution List" > $OutFileName
                    Write-Output $Reason >> $OutFileName
                    Write-Output "Dynamic Distribution Group Information" >> $OutFileName
                    Get-DynamicDistributionGroup $Global:txtInpName.Text >> $OutFileName
                    Write-Output "Dynamic Distribution Group Detailed Information" >> $OutFileName
                    Get-DynamicDistributionGroup $Global:txtInpName.Text |fl >> $OutFileName
                    Write-Output "Dynamic Distribtuion Group Membership Rule" >> $OutFileName
                    Get-DynamicDistributionGroup $Global:txtInpName.Text |fl RecipientFilter,IncludedRecipients >>$OutFileName
                    Remove-DynamicDistributionGroup $Global:txtInpName.TexL -confirm:$False
                }
                else
                {
                    $Output = $wshell.Popup("Group Not Removed because there are still valid group members",0,"Not Deleted",0+32)
                }
            }
            Remove-PSSession $Session
            invoke-Expression -Command .\ConnectO365.ps1
        }
    }

    If (($Global:chkO365Group.Checked -eq "Checked") -or ($Global:txtInpName.Text -like "GRP*"))
    {
        $ModInputDL = Global:txtInpName.Text.Replace(" ","")
        $ModInputDL = $ModInputDL.Replace(".","")
        $DstExists = [bool](Get-UnifiedGroup $ModInputDL -ErrorAction SilentlyContinue)
        If ($DstExists -eq "True")
        {
            $RemDL = Get-UnifiedGroup $ModInputDL
        `   $OutFileName = "c:\temp\" + ($Global:txtInpName.Text -replace " ","") + ".txt"
            $Output = $wshell.Popup("Removing group " + $Global:txtInpName.Text + " copy file " + $OutfileName + " to the Team Sharepoint Site",0,"Move Log File",0+32)
            Write-Output "Unified Group Information" >> $OutFileName
            Get-UnifiedGroup $ModInputDL >> $OutFileName
            Write-Output "Unified Group Detailed Information" >> $OutFileName
            Get-UnifiedGroup $ModInputDL |fl >> $OutFileName
            Write-Output "Unfied Group Owners and Members " >> $OutFileName
            Get-UnifiedGroupLinks -LinkType Owners $ModInputDL |fl Name >>$OutFileName
            Get-UnifiedGroupLinks -LinkType Members $ModInputDL |fl Name >>$OutFileName
            Remove-UnifiedGroup $ModInputDL -Confirm:$False
        }
    }
    Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\DLRemoval.oft

}
else
{
    $Output = $wshell.Popup("As requested group removal has been cancelled",0,"Do Not Continue",0+32)
}
