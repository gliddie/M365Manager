<#
#
#  Called by:  DistributionSecurityGroupMenu.ps1
#
#  10/27/2020 - Created for new GUI Interface
#>

function Add-ActionBoxes
{
    ## Add Access
    $Global:chkAdd = New-Object Windows.Forms.checkbox 
        $Global:chkAdd.Left = 570; $Global:chkAdd.Width = 200; $Global:chkAdd.Top = 60  
        $Global:chkAdd.Text = "Add Access" 
        $Global:chkAdd.Checked = $Global:chkAdd.Checked   # set a default value 
        $Global:chkAdd.TabIndex = 0
        $Global:form.Controls.Add($Global:chkAdd) 
        # Obtain Value with: $Global:chkThis.Checked

    ## Remove Access
    $Global:chkRemove = New-Object Windows.Forms.checkbox 
        $Global:chkRemove.Left = 570; $Global:chkRemove.Width = 200; $Global:chkRemove.Top = 80  
        $Global:chkRemove.Text = "Remove Access" 
        $Global:chkRemove.Checked = $Global:chkRemove.Checked   # set a default value 
        $Global:chkRemove.TabIndex = 1
        $Global:form.Controls.Add($Global:chkRemove) 
        # Obtain Value with: $Global:chkRemove.Checked

    If ($Global:DL.AcceptMessagesOnlyFromSendersOrmembers -like "ACL.*")
    {
        ## Remove Access
        $Global:chkGroup = New-Object Windows.Forms.checkbox 
            $Global:chkGroup.Left = 570; $Global:chkGroup.Width = 200; $Global:chkGroup.Top = 100  
            $Global:chkGroup.Text = "To/From Group" 
            $Global:chkGroup.Checked = $Global:chkGroup.Checked   # set a default value 
            $Global:chkGroup.TabIndex = 1
            $Global:form.Controls.Add($Global:chkGroup) 
            # Obtain Value with: $Global:chkGroup.Checked
    }
}

function Build-DLInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Review Restricted Distribution Group Access" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(450,250) #(W,H)
    $BldDetails = "N"

    ## Get Details used to create
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "DL Name:"
        $Global:lblDispName.Top = 10 ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtInpName = New-Object Windows.Forms.TextBox  
        $Global:txtInpName.Top = 10; $Global:txtInpName.Left = 130; $Global:txtInpName.Width = 280;  
        $Global:txtInpName.Text = ""   # DisplayName
        $Global:txtInpName.TabIndex = 1
        $Global:form.Controls.Add($Global:txtInpName)    # Add to Form

    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 40 ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.Top = 40; $Global:txtInpTaskNo.Left = 130; $Global:txtInpTaskNo.Width = 120;  
        $Global:txtInpTaskNo.Text = "TASK"   # Enter ticket number
        $Global:txtInpTaskNo.TabIndex = 2
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form

        Add-FormStandardButtons
}

function Build-DLDetailsForm
{
    $BldDetails = "N"
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Distribution Group Restricted Access Details" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(750,300) #(W,H)

    If ($Global:Refresh -eq "N")
    {
        Add-ActionBoxes
    }

    ## Get Details used to create
    ## Display Name
    $LocTop = 10
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "DL Name:"
        $Global:lblDispName.Top = $LocTop ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtDLDet = New-Object Windows.Forms.TextBox  
        $Global:txtDLDet.Top = $LocTop; $Global:txtDLDet.Left = 140; $Global:txtDLDet.Width = 400;  
        $Global:txtDLDet.Text = $Global:DL.Name
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
        $Global:txtDetTaskNo.Top = $LocTop; $Global:txtDetTaskNo.Left = 140; $Global:txtDetTaskNo.Width = 400;
        $Global:txtDetTaskNo.Text = $Global:txtInpTaskNo.Text 
        $Global:txtDetTaskNo.TabIndex = 2
        $Global:form.Controls.Add($Global:txtDetTaskNo)    # Add to Form

    $LocTop = $LocTop + 30
    ## Current Access
    $Global:lblDLAccess = New-Object System.Windows.Forms.Label   
        $Global:lblDLAccess.Text = "Current Access:"  
        $Global:lblDLAccess.Top = $LocTop ; $Global:lblDLAccess.Left = 10; $Global:lblDLAccess.Width=120 ; $Global:lblDLAccess.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLAccess)    # Add to Form 
        # 
        $Global:txtInpDLAccess = New-Object Windows.Forms.TextBox  
        $Global:txtInpDLAccess.Top = $LocTop; $Global:txtInpDLAccess.Left = 140; $Global:txtInpDLAccess.Width = 400;  
        $Global:txtInpDLAccess.TabIndex = 3
        $Global:txtInpDLAccess.Text = $Global:DL.AcceptMessagesOnlyFromSendersOrmembers -join ", "
        $Global:form.Controls.Add($Global:txtInpDLAccess)    # Add to Form

    ## Access Change
    $LocTop = $LocTop + 30
    $Global:lblAccessTo = New-Object System.Windows.Forms.Label   
        $Global:lblAccessTo.Text = "Account to Add/Remove:"
        $Global:lblAccessTo.Top = $LocTop ; $Global:lblAccessTo.Left = 10; $Global:lblAccessTo.Width=120 ;$Global:lblAccessTo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAccessTo)    # Add to Form 
        # 
        $Global:txtAccessTo = New-Object Windows.Forms.TextBox  
        $Global:txtAccessTo.Top = $LocTop; $Global:txtAccessTo.Left = 140; $Global:txtAccessTo.Width = 400;
        $Global:txtAccessTo.TabIndex = 4
        $Global:txtAccessTo.Text = ""
        $Global:form.Controls.Add($Global:txtAccessTo)    # Add to Form 

    If ($Global:Refresh -eq "N")
    {
        If ($Global:DL.AcceptMessagesOnlyFromSendersOrmembers -like "ACL.*")
        {
            ## Access Change
            $LocTop = $LocTop + 30
            $Global:lblGroup = New-Object System.Windows.Forms.Label   
                $Global:lblGroup.Text = "Group to Change:"
                $Global:lblGroup.Top = $LocTop ; $Global:lblGroup.Left = 10; $Global:lblGroup.Width=120 ;$Global:lblGroup.AutoSize = $true
                $Global:form.Controls.Add($Global:lblGroup)    # Add to Form 
                # 
                $Global:txtGroup = New-Object Windows.Forms.TextBox  
                $Global:txtGroup.Top = $LocTop; $Global:txtGroup.Left = 140; $Global:txtGroup.Width = 400;
                $Global:txtGroup.TabIndex = 4
                $Global:txtGroup.Text = ""
                $Global:form.Controls.Add($Global:txtGroup)    # Add to Form 
        }
    }
         
    Add-FormStandardButtons
}

########################

$Global:Refresh = "N"
Build-DLInputForm
Publish-Form

$DstExists = [bool](Get-DistributionGroup $Global:txtInpName.Text -ErrorAction SilentlyContinue)
if ($DstExists -eq "True")
{
    $Global:DL = Get-DistributionGroup $Global:txtInpName.Text

    Build-DLDetailsForm
    Publish-Form

    If ($Global:Result -eq "OK")
    {
        $Global:Refresh = "Y"
        If ($Global:txtAccessTo.Text.Length -gt 0)
        {
            If ($Global:chkAdd.Checked -eq "Checked")
            {
                If ($Global:chkGroup.Checked -eq "Checked")
                {
                    Add-DistributionGroupMember $Global:txtGroup.Text -Member $Global:txtAccessTo.Text -ByPassSecurityGroupManagerCheck -confirm:$False
                }
                else
                {
               
                    If ($Global:txtAccessTo.Text -notlike "*,*")
                    {
                        Set-DistributionGroup $Global:txtInpName.Text -AcceptMessagesOnlyFromSendersOrMembers @{add=$Global:txtAccessTo.Text}
                    }
                    else
                    {
                        $Acc = $Global:txtAccessTo.Text -split ","
                        ForEach ($Acc in $Acc)
                        {
                            Set-DistributionGroup $Global:txtInpName.Text -AcceptMessagesOnlyFromSendersOrMembers @{add=$Acc}
                        }
                    }
                }
            }

            If ($Global:chkRemove.Checked -eq "Checked")
            {
                If ($Global:chkGroup.Checked -eq "Checked")
                {
                    Remove-DistributionGroupMember $Global:txtGroup.Text -Member $Global:txtAccessTo.Text -ByPassSecurityGroupManagerCheck -confirm:$False
                }
                else
                {
                    If ($Global:txtAccessTo.Text -notlike "*,*")
                    {
                        Set-DistributionGroup $Global:txtInpName.Text -AcceptMessagesOnlyFromSendersOrMembers @{remove=$Global:txtAccessTo.Text}
                    }
                    else
                    {
                        $Acc = $Global:txtAccessTo.Text -split ","
                        ForEach ($Acc in $Acc)
                        {
                            Set-DistributionGroup $Global:txtInpName.Text -AcceptMessagesOnlyFromSendersOrMembers @{remove=$Acc}
                        }
                    }
                }
            }
        }
        else
        {
            $Output = $wshell.Popup("No Accounts Listed to Add/Remove from Access.",0,"Request Cancelled",0+32)
        }

        $Global:DL = Get-DistributionGroup $Global:txtInpName.Text
        Build-DLDetailsForm
        Publish-Form
    }
    else
    {
        $Output = $wshell.Popup("Restricted Access to Distribution Group cancelled.",0,"Request Cancelled",0+32)
    }
}
else
{
    $Output = $wshell.Popup("Distribution Group List Does Not Exist - No Changes Made",0,"Not Found",0+32)
}
