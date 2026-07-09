<#
#
#  Called by:  O365MainMenu.ps1
#
#  08/09/2020 - Changed over to a GUI Interface
#  04/17/2021 - Commented out the connect to O365
#  11/02/2022 - Moved the Build-DefaultForm, Get-DLGrpDetails, Get-UniGrpDetails, Hide-Details, UnHide-Details and Write-GrpDetails function from the UpdateGroupOwnership script
#>

function Build-DLMenuForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Distribution List / Security and O365 Group Admin Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 640 ; $form.Height = 650  # Make the form wider 
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 

    $TopLoc = 20 

    ## Label and TextBox  
    ## Title Line
    $Global:lblDLSecTitleLine = New-Object System.Windows.Forms.Label   
        $lblDLSecTitleLine.Text = "Distribution List/Security Group Action"
        $lblDLSecTitleLine.Top = 15 ; $lblDLSecTitleLine.Left = 60; $lblDLSecTitleLine.Width=220 ;$lblDLSecTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($lblDLSecTitleLine)    # Add to Form 

    ## Add/Remove Aliases
    $TopLoc = $TopLoc + 20
    $Global:chkAddRemAlias = New-Object Windows.Forms.RadioButton 
        $Global:chkAddRemAlias.Left = 100; $Global:chkAddRemAlias.Width = 450; $Global:chkAddRemAlias.Top = $TopLoc  
        $Global:chkAddRemAlias.Text = "Add/Remove Additional eMail Alias to Distribution List" 
        $Global:chkAddRemAlias.Checked = $false   # set a default value 
        $Global:chkAddRemAlias.TabIndex = 1
        $Global:form.Controls.Add($Global:chkAddRemAlias) 
        # Obtain Value with: $Global:chkAddRemAlias.Checked
        
    ## New DL
    $TopLoc = $TopLoc + 20
    $Global:chkNewDL = New-Object Windows.Forms.RadioButton 
        $Global:chkNewDL.Left = 100; $Global:chkNewDL.Width = 450; $Global:chkNewDL.Top = $TopLoc
        $Global:chkNewDL.Text = "Create a New Distribtuion Group" 
        $Global:chkNewDL.Checked = $false   # set a default value 
        $Global:chkNewDL.TabIndex = 2
        $Global:form.Controls.Add($Global:chkNewDL) 
        # Obtain Value with: $Global:chkNewDL.Checked
        $Global:InputFocus = $Global:chkNewDL

    ## Change DL/Sec/O365 Group
    $TopLoc = $TopLoc + 20
    $Global:chkChgDLSec = New-Object Windows.Forms.RadioButton 
        $Global:chkChgDLSec.Left = 100; $Global:chkChgDLSec.Width = 450; $Global:chkChgDLSec.Top = $TopLoc  
        $Global:chkChgDLSec.Text = "Change Distribution List/Security Group/Unified Group Ownership" 
        $Global:chkChgDLSec.Checked = $false   # set a default value 
        $Global:chkChgDLSec.TabIndex = 3
        $Global:form.Controls.Add($Global:chkChgDLSec) 
        # Obtain Value with: $Global:chkChgDLSec.Checked

    ## Change DL Restriced Access
    $TopLoc = $TopLoc + 20
    $Global:chkDLRestrict = New-Object Windows.Forms.RadioButton 
        $Global:chkDLRestrict.Left = 100; $Global:chkDLRestrict.Width = 450; $Global:chkDLRestrict.Top = $TopLoc  
        $Global:chkDLRestrict.Text = "Change Distribution List Restricted Access"
        $Global:chkDLRestrict.Checked = $false   # set a default value 
        $Global:chkDLRestrict.TabIndex = 3
        $Global:form.Controls.Add($Global:chkDLRestrict) 
        # Obtain Value with: $Global:chkDLRestrict.Checked

    ## DL/Sec Group Membership
    $TopLoc = $TopLoc + 20
    $Global:chkDLMem = New-Object Windows.Forms.RadioButton 
        $Global:chkDLMem.Left = 100; $Global:chkDLMem.Width = 450; $Global:chkDLMem.Top = $TopLoc  
        $Global:chkDLMem.Text = "Count of Distribution/Security Group Membership" 
        $Global:chkDLMem.Checked = $Global:chkDLMem.Checked   # set a default value 
        $Global:chkDLMem.TabIndex = 4
        $Global:form.Controls.Add($Global:chkDLMem) 
        # Obtain Value with: $Global:chkDLMem.Checked

    ## External Access
    $TopLoc = $TopLoc + 20
    $Global:chkExtAccess = New-Object Windows.Forms.RadioButton 
        $Global:chkExtAccess.Left = 100; $Global:chkExtAccess.Width = 450; $Global:chkExtAccess.Top = $TopLoc  
        $Global:chkExtAccess.Text = "Grant/Deny External Addresses Access to Distribution List or Security Group" 
        $Global:chkExtAccess.Checked = $false   # set a default value 
        $Global:chkExtAccess.TabIndex = 5
        $Global:form.Controls.Add($Global:chkExtAccess) 
        # Obtain Value with: $Global:chkExtAccess.Checked

    ## Add Members
    $TopLoc = $TopLoc + 20
    $Global:chkAddMem = New-Object Windows.Forms.RadioButton 
        $Global:chkAddMem.Left = 100; $Global:chkAddMem.Width = 450; $Global:chkAddMem.Top = $TopLoc  
        $Global:chkAddMem.Text = "M&A Activities Menu" 
        $Global:chkAddMem.Checked = $false   # set a default value 
        $Global:chkAddMem.TabIndex = 6
        $Global:form.Controls.Add($Global:chkAddMem) 
        # Obtain Value with: $Global:chkAddMem.Checked

    ## Remove DL/Security Group 
    $TopLoc = $TopLoc + 20
    $Global:chkRemDLSecGrp = New-Object Windows.Forms.RadioButton 
        $Global:chkRemDLSecGrp.Left = 100; $Global:chkRemDLSecGrp.Width = 450; $Global:chkRemDLSecGrp.Top = $TopLoc  
        $Global:chkRemDLSecGrp.Text = "Remove Distribution List or Security Group" 
        $Global:chkRemDLSecGrp.Checked = $false   # set a default value 
        $Global:chkRemDLSecGrp.TabIndex = 7 
        $Global:form.Controls.Add($Global:chkRemDLSecGrp) 
        # Obtain Value with: $Global:chkRemDLSecGrp.Checked

    ## Remove Members
    $TopLoc = $TopLoc + 20
    $Global:chkRemMem = New-Object Windows.Forms.RadioButton 
        $Global:chkRemMem.Left = 100; $Global:chkRemMem.Width = 450; $Global:chkRemMem.Top = $TopLoc  
        $Global:chkRemMem.Text = "Remove Member(s) from a Distribution List/Security or Unified Group" 
        $Global:chkRemMem.Checked = $false   # set a default value 
        $Global:chkRemMem.TabIndex = 8
        $Global:form.Controls.Add($Global:chkRemMem) 
        # Obtain Value with: $Global:chkRemMem.Checked

    ## Rename DL/Security Group
    $TopLoc = $TopLoc + 20
    $Global:chkRenDL = New-Object Windows.Forms.RadioButton 
        $Global:chkRenDL.Left = 100; $Global:chkRenDL.Width = 450; $Global:chkRenDL.Top = $TopLoc  
        $Global:chkRenDL.Text = "Rename Distribution List or Security Group" 
        $Global:chkRenDL.Checked = $false   # set a default value 
        $Global:chkRenDL.TabIndex = 9
        $Global:form.Controls.Add($Global:chkRenDL) 
        # Obtain Value with: $Global:chkRenDL.Checked

    ## Replace DL/Security Group Membership
    $TopLoc = $TopLoc + 20
    $Global:chkReplMem = New-Object Windows.Forms.RadioButton 
        $Global:chkReplMem.Left = 100; $Global:chkReplMem.Width = 450; $Global:chkReplMem.Top = $TopLoc  
        $Global:chkReplMem.Text = "Replace Distribution List or Security Group Membership" 
        $Global:chkReplMem.Checked = $false   # set a default value 
        $Global:chkReplMem.TabIndex = 10
        $Global:form.Controls.Add($Global:chkReplMem) 
        # Obtain Value with: $Global:chkReplMem.Checked

    ## Update DL/Security Group
    $TopLoc = $TopLoc + 20
    $Global:chkUpdDLSecGrp = New-Object Windows.Forms.RadioButton 
        $Global:chkUpdDLSecGrp.Left = 100; $Global:chkUpdDLSecGrp.Width = 450; $Global:chkUpdDLSecGrp.Top = $TopLoc  
        $Global:chkUpdDLSecGrp.Text = "Update (Add/Remove) Distribution List or Security Group Membership" 
        $Global:chkUpdDLSecGrp.Checked = $false   # set a default value 
        $Global:chkUpdDLSecGrp.TabIndex = 11
        $Global:form.Controls.Add($Global:chkUpdDLSecGrp) 
        # Obtain Value with: $Global:chkUpdDLSecGrp.Checked

    ## Update DL/Security Group Authorized Users
    $TopLoc = $TopLoc + 20
    $Global:chkUpdAuthUser = New-Object Windows.Forms.RadioButton 
        $Global:chkUpdAuthUser.Left = 100; $Global:chkUpdAuthUser.Width = 450; $Global:chkUpdAuthUser.Top = $TopLoc  
        $Global:chkUpdAuthUser.Text = "Update Distribution List or Security Group Authorized Users" 
        $Global:chkUpdAuthUser.Checked = $false   # set a default value 
        $Global:chkUpdAuthUser.TabIndex = 11
        $Global:form.Controls.Add($Global:chkUpdAuthUser) 
        # Obtain Value with: $Global:chkUpdDLSecGrp.Checked
         
    ## Title Line
    $TopLoc = $TopLoc + 40
    $Global:lblDynTitleLine = New-Object System.Windows.Forms.Label   
        $Global:lblDynTitleLine.Text = "Dynamic Distribution Group Actions"
        $Global:lblDynTitleLine.Top = $TopLoc ; $Global:lblDynTitleLine.Left = 60; $Global:lblDynTitleLine.Width=120 ;$Global:lblDynTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDynTitleLine)    # Add to Form 
    
    
    ## New DST Group
    $TopLoc = $TopLoc + 20
    $Global:chkNewDST = New-Object Windows.Forms.RadioButton 
        $Global:chkNewDST.Left = 100; $Global:chkNewDST.Width = 450; $Global:chkNewDST.Top = $TopLoc  
        $Global:chkNewDST.Text = "Create New UL Employee Only and UL Staff Dynamic Distribution Groups" 
        $Global:chkNewDST.Checked = $false   # set a default value 
        $Global:chkNewDST.TabIndex = 12
        $Global:form.Controls.Add($Global:chkNewDST) 
        # Obtain Value with: $Global:chkNewDST.Checked

    ## DST Members
    $TopLoc = $TopLoc + 20
    $Global:chkDSTMem = New-Object Windows.Forms.RadioButton 
        $Global:chkDSTMem.Left = 100; $Global:chkDSTMem.Width = 450; $Global:chkDSTMem.Top = $TopLoc  
        $Global:chkDSTMem.Text = "List of DST Membership" 
        $Global:chkDSTMem.Checked = $false   # set a default value 
        $Global:chkDSTMem.TabIndex = 13
        $Global:form.Controls.Add($Global:chkDSTMem) 
        # Obtain Value with: $Global:chkDSTMem.Checked

    ## Remove DST Group
    $TopLoc = $TopLoc + 20
    $Global:chkRemDST = New-Object Windows.Forms.RadioButton 
        $Global:chkRemDST.Left = 100; $Global:chkRemDST.Width = 450; $Global:chkRemDST.Top = $TopLoc  
        $Global:chkRemDST.Text = "Remove Dynamic Distribution List" 
        $Global:chkRemDST.Checked = $false   # set a default value 
        $Global:chkRemDST.TabIndex = 14
        $Global:form.Controls.Add($Global:chkRemDST) 
        # Obtain Value with: $Global:chkRemDST.Checked

    ## Title Line
    $TopLoc = $TopLoc + 40
    $Global:lblUniGrpTitleLine = New-Object System.Windows.Forms.Label   
        $Global:lblUniGrpTitleLine.Text = "O365 Groups/Teams Actions"
        $Global:lblUniGrpTitleLine.Top = $TopLoc ; $Global:lblUniGrpTitleLine.Left = 60; $Global:lblUniGrpTitleLine.Width=120 ;$Global:lblUniGrpTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblUniGrpTitleLine)    # Add to Form 

    ## Check Team/Sharepoint Site
    $TopLoc = $TopLoc + 20
    $Global:chkTeamShrPtSite = New-Object Windows.Forms.RadioButton 
        $Global:chkTeamShrPtSite.Left = 100; $Global:chkTeamShrPtSite.Width = 450; $Global:chkTeamShrPtSite.Top = $TopLoc  
        $Global:chkTeamShrPtSite.Text = "Check for Team SharePoint Site" 
        $Global:chkTeamShrPtSite.Checked = $false   # set a default value 
        $Global:chkTeamShrPtSite.TabIndex = 15
        $Global:form.Controls.Add($Global:chkTeamShrPtSite) 
        # Obtain Value with: $Global:chkTeamShrPtSite.Checked

    ## SoftDeleted Site
    $TopLoc = $TopLoc + 20
    $Global:chkSoftDelObj = New-Object Windows.Forms.RadioButton 
        $Global:chkSoftDelObj.Left = 100; $Global:chkSoftDelObj.Width = 450; $Global:chkSoftDelObj.Top = $TopLoc  
        $Global:chkSoftDelObj.Text = "Check for SoftDeleted Team SharePoint Sites" 
        $Global:chkSoftDelObj.Checked = $false   # set a default value 
        $Global:chkSoftDelObj.TabIndex = 16
        $Global:form.Controls.Add($Global:chkSoftDelObj) 
        # Obtain Value with: $Global:chkSoftDelObj.Checked

    ## Create O365/Team
    $TopLoc = $TopLoc + 20
    $Global:chkCreUniGrp = New-Object Windows.Forms.RadioButton 
        $Global:chkCreUniGrp.Left = 100; $Global:chkCreUniGrp.Width = 450; $Global:chkCreUniGrp.Top = $TopLoc  
        $Global:chkCreUniGrp.Text = "Create O365 Groups (Teams/Unified Groups)" 
        $Global:chkCreUniGrp.Checked = $Global:chkCreUniGrp.Checked   # set a default value 
        $Global:chkCreUniGrp.TabIndex = 17
        $Global:form.Controls.Add($Global:chkCreUniGrp) 
        # Obtain Value with: $Global:chkCreUniGrp.Checked

    ## Remove Uni Mem
    $TopLoc = $TopLoc + 20
    $Global:chkUniGrpMem = New-Object Windows.Forms.RadioButton 
        $Global:chkUniGrpMem.Left = 100; $Global:chkUniGrpMem.Width = 500; $Global:chkUniGrpMem.Top = $TopLoc  
        $Global:chkUniGrpMem.Text = "Remove Member(s) from a Distribution List/Security or Unified Group" 
        $Global:chkUniGrpMem.Checked = $false   # set a default value 
        $Global:chkUniGrpMem.TabIndex = 18
        $Global:form.Controls.Add($Global:chkUniGrpMem) 
        # Obtain Value with: $Global:chkUniGrpMem.Checked

    ## Misc Powershell Commands
    $TopLoc = $TopLoc + 20
    $Global:chkO356GrpTeams = New-Object Windows.Forms.RadioButton 
        $Global:chkO356GrpTeams.Left = 100; $Global:chkO356GrpTeams.Width = 450; $Global:chkO356GrpTeams.Top = $TopLoc  
        $Global:chkO356GrpTeams.Text = "View All O365 Groups Enabled for Teams" 
        $Global:chkO356GrpTeams.Checked = $false   # set a default value 
        $Global:chkO356GrpTeams.TabIndex = 19
        $Global:form.Controls.Add($Global:chkO356GrpTeams) 
        # Obtain Value with: $Global:chkO356GrpTeams.Checked

    Add-FormStandardButtons
}

Function Build-DefaultForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow"
    $Global:form.Text = "Update Group Ownership"
    If ($Global:chkRemDLSecGrp.Checked -eq $True)
    {
        $Global:form.Text = "Remove Distribution/Security/O365 Group"
    }
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 600 ; $form.Height = 130   # Make the form wider 
  
    $Top = 20
    ## Group Name
    $Script:lblDLName = New-Object System.Windows.Forms.Label   
        $Script:lblDLName.Text = "Distribution Group Name:" 
        $Script:lblDLName.Top = $Top ; $Script:lblDLName.Left = 5; $Script:lblDLName.Width=150 ;$Script:lblDLName.AutoSize = $true 
        $form.Controls.Add($Script:lblDLName)    # Add to Form 
        # 
        $Script:txtDLName = New-Object Windows.Forms.TextBox
        $Script:txtDLName.Top = $Top; $Script:txtDLName.Left = 160; $Script:txtDLName.Width = 250;  
        $Script:txtDLName.Text = ""
        $Global:form.Controls.Add($Script:txtDLName)    # Add to Form
        $Global:InputFocus = $Script:txtDLName
        $Script:txtDLName.Add_Click({
            Hide-Details
            })

    $Script:ButGetDL = New-Object Windows.Forms.Button
        $Script:ButGetDL.Location = New-object System.Drawing.Size(420,$Top)
        $Script:ButGetDL.Size = new-Object System.Drawing.Size(150,20)
        $Script:ButGetDL.Text = "Get Group Details"
        $Script:ButGetDL.TabIndex = 0
        $Global:form.Controls.Add($Script:ButGetDL)
        $Script:ButGetDL.Add_Click({
            $UniGrp = $Script:txtDLName.Text
            write-host "Retreiving Account Details for Group: $Script:DLName" -ForegroundColor Cyan
            If ($Script:txtDLName.Text -notlike "*@*")
            {
                $UniGrp = ($Script:txtDLName.Text -replace("[.]","")).Trim()
                $UniGrp = ($UniGrp -replace(" ","")) + "@ul.onmicrosoft.com"
            }
            $Script:UniExists = [bool]($O365Grp = get-UnifiedGroup $UniGrp -ErrorAction SilentlyContinue)
            $Script:DLExists = [bool]($DLGrp = Get-DistributionGroup $Script:txtDLName.Text -ErrorAction SilentlyContinue)
            If (($Script:DLExists -eq $True) -or ($Script:UniExists -eq $True))
            {
                If ($Script:DLExists -eq $True)
                {
                    Get-DLGrpDetails
                }
                else
                {
                    write-host "Getting Unified Group Details" -ForegroundColor Cyan
                    Get-UniGrpDetails
                }
                UnHide-Details
                $Global:OKButton.Text = "Update"
                $Global:OKButton.visible = $True
                If ($Global:chkRemDLSecGrp.Checked -eq $True)
                {
                    $Global:OKButton.Text = "Remove"
                }
            }
            else
            {
                $Global:InputFocus = $Script:txtTicketNo
                $Script:txtDLName.Text = "Invalid"
                $Script:ButGetDL.Visible = "True"
            }
            })

    $Top = $Top + 30
    ## Ticket Number
    $Script:lblTicketNo = New-Object System.Windows.Forms.Label   
        $Script:lblTicketNo.Text = "Ticket Number:" 
        $Script:lblTicketNo.Top = $Top ; $Script:lblTicketNo.Left = 5; $Script:lblTicketNo.Width=150 ;$Script:lblTicketNo.AutoSize = $true 
        $form.Controls.Add($Script:lblTicketNo)    # Add to Form 
        # 
        $Script:txtTicketNo = New-Object Windows.Forms.ComboBox
        $Script:txtTicketNo.Top = $Top; $Script:txtTicketNo.Left = 160; $Script:txtTicketNo.Width = 200;
        $Script:txtTicketNo.TabIndex = 1
        $Script:txtTicketNo.Text = "TASK"
        [void] $Script:txtTicketNo.Items.Add("NoOwners/NoMembers Request")  # Add element to listbox 
        $Global:form.Controls.Add($Script:txtTicketNo)    # Add to Form

    $Top = $Top + 30
    ## Group Address
    $Script:lblDLAddr = New-Object System.Windows.Forms.Label   
        $Script:lblDLAddr.Text = "Distribution Primary Address:" 
        $Script:lblDLAddr.Top = $Top ; $Script:lblDLAddr.Left = 5; $Script:lblDLAddr.Width=150 ;$Script:lblDLAddr.AutoSize = $true 
        $form.Controls.Add($Script:lblDLAddr)    # Add to Form 
        # 
        $Script:txtDLAddr = New-Object Windows.Forms.TextBox
        $Script:txtDLAddr.Top = $Top; $Script:txtDLAddr.Left = 160; $Script:txtDLAddr.Width = 250;
        $Script:txtDLAddr.ReadOnly = $True; $Script:txtDLAddr.TabStop = $False
        $Script:txtDLAddr.Text = ""
        $Global:form.Controls.Add($Script:txtDLAddr)    # Add to Form

    $Top = $Top + 30
    ## MailTip (Extract the mail sent part)
    $Script:lblMailTip = New-Object System.Windows.Forms.Label   
        $Script:lblMailTip.Text = "MailTip:" 
        $Script:lblMailTip.Top = $Top ; $Script:lblMailTip.Left = 5; $Script:lblMailTip.Width=150 ;$Script:lblMailTip.AutoSize = $true 
        $form.Controls.Add($Script:lblMailTip)    # Add to Form 
        # 
        $Script:txtMailTip = New-Object Windows.Forms.TextBox
        $Script:txtMailTip.Top = $Top; $Script:txtMailTip.Left = 160; $Script:txtMailTip.Width = 450;
        $Script:txtMailTip.ReadOnly = $True; $Script:txtMailTip.TabStop = $False
        $Script:txtMailTip.Text = "No Tip Found"
        $Global:form.Controls.Add($Script:txtMailTip)    # Add to Form

    $Top = $Top + 30
    ## Current Owner
    $Script:lblCurrOwner = New-Object System.Windows.Forms.Label   
        $Script:lblCurrOwner.Text = "Current Owner(s):" 
        $Script:lblCurrOwner.Top = $Top ; $Script:lblCurrOwner.Left = 5; $Script:lblCurrOwner.Width=150 ;$Script:lblCurrOwner.AutoSize = $true 
        $form.Controls.Add($Script:lblCurrOwner)    # Add to Form 
        # 
        $Script:txtCurrOwner = New-Object Windows.Forms.TextBox
        $Script:txtCurrOwner.Top = $Top; $Script:txtCurrOwner.Left = 160; $Script:txtCurrOwner.Width = 450;
        $Script:txtCurrOwner.ReadOnly = $True; $Script:txtCurrOwner.TabStop = $False 
        $Script:txtCurrOwner.Text = ""
        $Global:form.Controls.Add($Script:txtCurrOwner)    # Add to Form

    $Top = $Top + 30
    ## Current Members
    $Script:lblCurrMbr = New-Object System.Windows.Forms.Label   
        $Script:lblCurrMbr.Text = "Current Member(s):" 
        $Script:lblCurrMbr.Top = $Top ; $Script:lblCurrMbr.Left = 5; $Script:lblCurrMbr.Width=150 ;$Script:lblCurrMbr.AutoSize = $true 
        $form.Controls.Add($Script:lblCurrMbr)    # Add to Form 
        # 
        $Script:txtCurrMbr = New-Object Windows.Forms.TextBox
        $Script:txtCurrMbr.MaxLength = 200000
        $Script:txtCurrMbr.Location = New-Object System.Drawing.Size(160,$Top)
        $Script:txtCurrMbr.Size = New-Object system.Drawing.Size(450,70)
        $Script:txtCurrMbr.MultiLine = $true
        $Script:txtcurrMbr.ScrollBars = 'Both'
        $Script:txtCurrMbr.ReadOnly = $true; $Script:txtCurrMbr.TabStop = $False
        $Script:txtCurrMbr.Text = ""
        $Global:form.Controls.Add($Script:txtCurrMbr)    # Add to Form

    $Top = $Top + 80
    ## New Owner
    $Script:lblNewOwner = New-Object System.Windows.Forms.Label   
        $Script:lblNewOwner.Text = "New Owner(s):" 
        $Script:lblNewOwner.Top = $Top ; $Script:lblNewOwner.Left = 5; $Script:lblNewOwner.Width=150 ;$Script:lblNewOwner.AutoSize = $true 
        $form.Controls.Add($Script:lblNewOwner)    # Add to Form 
        # 
        $Script:txtNewOwner = New-Object Windows.Forms.TextBox
        $Script:txtNewOwner.Top = $Top; $Script:txtNewOwner.Left = 160; $Script:txtNewOwner.Width = 450;$Script:txtNewOwner.TabIndex = 5
        $Global:form.Controls.Add($Script:txtNewOwner)    # Add to Form
        $Script:txtNewOwner.Add_Click({
            $Script:txtNewOwner.Text = ""
        })

    $Top = $Top + 30
    ## Remove Owners
    $Script:lblRemOwner = New-Object System.Windows.Forms.Label   
        $Script:lblRemOwner.Text = "Remove Owner(s):" 
        $Script:lblRemOwner.Top = $Top ; $Script:lblRemOwner.Left = 5; $Script:lblRemOwner.Width=150 ;$Script:lblRemOwner.AutoSize = $true 
        $form.Controls.Add($Script:lblRemOwner)    # Add to Form 
        # 
        $Script:txtRemOwner = New-Object Windows.Forms.TextBox
        $Script:txtRemOwner.Top = $Top; $Script:txtRemOwner.Left = 160; $Script:txtRemOwner.Width = 450; $Script:txtRemOwner.Height = 60
        $Script:txtRemOwner.TabIndex = 6
        $Global:form.Controls.Add($Script:txtRemOwner)    # Add to Form
        $Script:txtRemOwner.Add_Click({
            $Script:txtRemOwner.Text = ""
        })

    $Top = $Top + 30
    ## Report File
    $Script:lblRptFile = New-Object System.Windows.Forms.Label   
        $Script:lblRptFile.Text = "Report File:" 
        $Script:lblRptFile.Top = $Top ; $Script:lblRptFile.Left = 5; $Script:lblRptFile.Width=150 ;$Script:lblRptFile.AutoSize = $true 
        $form.Controls.Add($Script:lblRptFile)    # Add to Form 
        # 
        $Script:txtRptFile = New-Object Windows.Forms.TextBox
        $Script:txtRptFile.Top = $Top; $Script:txtRptFile.Left = 160; $Script:txtRptFile.Width = 450; $Script:txtRptFile.Height = 60
        $Script:txtRptFile.ReadOnly = $True; $Script:txtRptFile.TabStop = $False
        $Script:txtRptFile.Text = ""
        $Global:form.Controls.Add($Script:txtRptFile)    # Add to Form
}

Function Add-ActionBoxes
{
    $Left = 630
    ## Add Owner(s) RadioButton 
    $Script:chkAddOwner = New-Object Windows.Forms.Checkbox
        $Script:chkAddOwner.Left = $Left; $Script:chkAddOwner.Width = 200; $Script:chkAddOwner.Top = 60  
        $Script:chkAddOwner.Text = "Add Owner(s)" 
        $Script:chkAddOwner.Checked = $False   # set a default value
        $Script:chkAddOwner.TabIndex = 2
        $Global:form.Controls.Add($Script:chkAddOwner)
        $Script:chkAddOwner.Add_Click({
            $Script:chkReplOwner.Checked = $False
            UnHide-OwnerChanges
         }) 

    ## Replace Owner(s) RadioButton
    $Script:chkReplOwner = New-Object Windows.Forms.Checkbox
        $Script:chkReplOwner.Left = $Left; $Script:chkReplOwner.Width = 200; $Script:chkReplOwner.Top = 90  
        $Script:chkReplOwner.Text = "Replace Owner(s)" 
        $Script:chkReplOwner.Checked = $False   # set a default value 
        $Script:chkReplOwner.TabIndex = 3
        $Global:form.Controls.Add($Script:chkReplOwner)
        $Script:chkReplOwner.Add_Click({
            $Script:chkAddOwner.Checked = $False
            $Script:chkRemOwner.Checked = $False
            UnHide-OwnerChanges
        })

    ## Remove Owner(s) RadioButton
    $Script:chkRemOwner = New-Object Windows.Forms.Checkbox
        $Script:chkRemOwner.Left = $Left; $Script:chkRemOwner.Width = 200; $Script:chkRemOwner.Top = 120  
        $Script:chkRemOwner.Text = "Remove Owner(s)" 
        $Script:chkRemOwner.Checked = $False   # set a default value 
        $Script:chkRemOwner.TabIndex = 4
        $Global:form.Controls.Add($Script:chkRemOwner)
        $Script:chkRemOwner.Add_Click({
            $Script:chkReplOwner.Checked = $False
            UnHide-OwnerChanges
        }) 
}

Function Check-Owner
{
    If (($Script:txtCurrOwner.Text -notlike "*,*") -and ($Script:txtCurrOwner.Text.Length -gt 4))
    {
        $Active = (Get-Mailbox $Script:txtCurrOwner.Text).CustomAttribute2
        If ($Active -eq "T")
        {
           $Script:txtCurrOwner.Text = $Script:txtCurrOwner.Text + " (Terminated)"
           $Script:txtCurrOwner.BackColor = "#EEEDF0"
           $Script:txtCurrOwner.ForeColor = "Red"
           $Script:txtCurrOwner.ReadOnly = $True
        }
    }
}

Function Get-DLGrpDetails
{
    $Script:txtDLName.Text = $DLGrp.Name
    $Script:txtDLAddr.Text = $DLGrp.PrimarySMTPAddress
    If ($DLGrp.MailTip.Length -gt 20)
    {
        $EndofTip = $DLGrp.MailTip.Indexof("</body")-14
        $Script:txtMailTip.Text = $DLGrp.MailTip.SubString(14,$EndofTip)
    }
    $Script:txtCurrOwner.Text = ($DLGrp.ManagedBy) -join ", "
    Check-Owner
    $Script:txtCurrMbr.Text = ((Get-DistributionGroupMember $DLGrp.Name) -join ", ")
    $Script:txtNewOwner.Text = "Enter Email Address or Emp# of individuals to add, separate multiple entries with commas"
    $Script:txtRemOwner.Text = "Enter Email Address or Emp# of individauls to remove, separate multiple entries with commas"
    $Script:txtRptFile.Text = "E:\Automation\UpdateGroupOwnership\Report\Report-UpdateGroupOwnership-" + ($DLGrp.Alias -replace("[.]",""))  + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
    If ($Global:chkRemDLSecGrp.Checked -eq $True)
    {
            $Script:txtRptFile.Text = "E:\Automation\RemoveDLGroup\Report\Report-RemoveDLGroup-" + ($DLGrp.Alias -replace("[.]",""))  + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
    }
}

Function Get-UniGrpDetails
{
    $Script:txtDLName.Text = $O365Grp.DisplayName
    $Script:txtDLAddr.Text = $O365Grp.PrimarySMTPAddress
    If ($O365Grp.MailTip.Length -gt 20)
    {
        $EndofTip = $O365Grp.MailTip.Indexof("</body")-14
        $Script:txtMailTip.Text = $O365Grp.MailTip.SubString(14,$EndofTip)
    }
    $Script:txtCurrOwner.Text = $O365Grp.ManagedBy -join ", "
    Check-Owner
    $Script:txtCurrMbr.Text = (Get-UnifiedGroupLinks $O365Grp -LinkType Member) -join ", "
    $Script:txtNewOwner.Text = "Enter Email Address or Emp# of individuals to add, separate multiple entries with commas"
    $Script:txtRemOwner.Text = "Enter Email Address or Emp# of individauls to remove, separate multiple entries with commas"
    $Script:txtRptFile.Text = "e:\Automation\UpdateGroupOwnership\Report\Report-UpdateGroupOwnership-" + $O365Grp.Alias  + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
}

Function Hide-Details
{
    $Global:form.Width = 600 ; $form.Height = 130
    $Script:ButGetDL.visible = $True
    $Script:txtDLName.ReadOnly = $False
    $Script:lblTicketNo.Visible = $False
    $Script:txtTicketNo.Visible = $False
    $Script:lblDLAddr.Visible = $False
    $Script:txtDLAddr.Visible = $False
    $Script:lblMailTip.Visible = $False
    $Script:txtMailTip.Visible = $False
    $Script:txtMailTip.Text = "No Tip Found"
    $Script:lblCurrOwner.Visible = $False
    $Script:txtCurrOwner.Visible = $False
    $Script:txtCurrOwner.ForeColor = "Black"

    $Script:lblCurrMbr.Visible = $False
    $Script:txtCurrMbr.Visible = $False
    $Script:lblNewOwner.Visible = $False
    $Script:txtNewOwner.Visible = $False
    $Script:lblRemOwner.Visible = $False
    $Script:txtRemOwner.Visible = $False
    $Script:chkAddOwner.Visible = $False
    $Script:chkReplOwner.Visible = $False
    $Script:chkRemOwner.Visible = $False
    $Global:OKButton.visible = $False
    $Global:cancelButton.Visible = $False
}

Function UnHide-Details
{
    $Global:form.Width = 780 ; $form.Height = 450
    $Script:ButGetDL.visible = $False
    $Script:txtDLName.ReadOnly = $True
    $Script:lblDLAddr.Visible = $True
    $Script:txtDLAddr.Visible = $True
    $Script:lblTicketNo.Visible = $True
    $Script:txtTicketNo.Visible = $True
    $Script:lblMailTip.Visible = $True
    $Script:txtMailTip.Visible = $True
    $Script:lblCurrOwner.Visible = $True
    $Script:txtCurrOwner.Visible = $True
    $Script:lblCurrMbr.Visible = $True
    $Script:txtCurrMbr.Visible = $True
    If ($Global:form.Text -notlike "Remove*")
    {
        If ($Script:txtCurrOwner.Text -notlike "*Term*")
        {
            $Script:chkAddOwner.Visible = $True
            $Script:chkRemOwner.Visible = $True
        }
        $Script:chkReplOwner.Visible = $True
    }
    $Global:OKButton.visible = $True
    $Global:cancelButton.Visible = $True
}

Function UnHide-OwnerChanges
{
    If (($Script:chkAddOwner.Checked -eq $True) -or ($Script:chkReplOwner.Checked -eq $True))
    {
        $Script:lblNewOwner.Visible = $True
        $Script:txtNewOwner.Visible = $True
        $Script:txtNewOwner.Text = "Enter Email Address or Emp# of individauls to remove, separate multiple entries with commas"
    }
    else
    {
        $Script:lblNewOwner.Visible = $False
        $Script:txtNewOwner.Visible = $False
    }

    If ($Script:chkReplOwner.Checked -eq $True)
    {
        $Script:lblRemOwner.Visible = $True
        $Script:txtRemOwner.Visible = $True
        $Script:txtRemOwner.BackColor = "LightGray"
        $Script:txtRemOwner.ForeColor = "Red"
        $Script:txtRemOwner.Text = "***ALL EXISTING OWNERS WILL BE REMOVED ***"
        $Script:txtRemOwner.ReadOnly = $True
    }

    If ($Script:chkRemOwner.Checked -eq $True)
    {
        If ($Script:chkAddOwner.Checked -eq $False)
        {
            $Script:lblNewOwner.Visible = $False
            $Script:txtNewOwner.Visible = $False            
        }
        $Script:chkReplOwner.Checked = $False
        $Script:txtRemOwner.BackColor = "White"
        $Script:txtRemOwner.ForeColor = "Black"
        $Script:txtRemOwner.Text = "Enter Email Address or Emp# of individauls to remove, separate multiple entries with commas"
        $Script:txtRemOwner.ReadOnly = $False
        $Script:lblRemOwner.Visible = $True
        $Script:txtRemOwner.Visible = $True
    }

    If (($Script:chkReplOwner.Checked -eq $False) -and ($Script:chkRemOwner.Checked -eq $False))
    {
        $Script:lblRemOwner.Visible = $False
        $Script:txtRemOwner.Visible = $False
    }
} 

Function Write-GroupDetails
{
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "                Group Name: " + $Script:txtDLName.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "             Ticket Number: " + $Script:txtTicketNo.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "     Primary Email Address: " + $Script:txtDLAddr.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "                   MailTip: " + $Script:txtMailTip.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "            Current Owners: " + $Script:txtCurrOwner.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "           Current Members: " + $Script:txtCurrMbr.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "             Owners to Add: " + $Script:txtNewOwner.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "          Owners to Remove: " + $Script:RemNewOwner.Text
    WriteReportEvent
}

#### Start of Script

Build-DLMenuForm
Publish-Form

Do
{
    If ($Global:chkNewDL.Checked -eq "Checked")
    {
        Invoke-Expression -Command .\NewDLGroup.ps1
    }

    If ($Global:chkReplMem.Checked -eq "Checked")
    {
        Invoke-Expression -Command .\ReplGrpMembers.ps1
    }

    If ($Global:chkRenDL.Checked -eq "Checked")
    {
        Invoke-Expression -Command .\RenameDLGroup.ps1
    }

    If ($Global:chkChgDLSec.Checked -eq "Checked")
    {
        Invoke-Expression -Command .\UpdateGroupOwnership.ps1
        <#
        $AnoGrp = "Y"
        DO
        {
            write-host ""
            write-host "     Enter ( 1) Change Ownership for a Single Group"
            write-host "           ( 2) Change Ownership for Multiple Groups"
            write-host ""
            write-host "           ( 0) to Return to the Distribution List and Security Group Admin Menu"
            write-Host "     Enter Option? " -ForegroundColor Red -NoNewline
            $NoGrps = Read-Host
            switch ($NoGrps)
            {
                1
                {
                    invoke-expression -Command .\GroupOwnershipChanges.ps1
                }
                2
                {
                    invoke-expression -Command .\GroupOwnershipChangeMultiple.ps1
                }
            }
            write-host "Change Membership on more Distribution Lists (Y/N)? " -ForegroundColor Yellow -NoNewline
            $AnoGrp = Read-Host
        } while ($anoGrp -eq "Y")
        #>
    }

    If ($Global:chkDLRestrict.Checked -eq "Checked")
    {
        Invoke-Expression -Command .\RestrictedDLAccess.ps1
        Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\DLRestrictedAccessGranted.oft        
    }

    If ($Global:chkRemDLSecGrp.Checked -eq "Checked")
    {
        invoke-expression -Command .\RemoveDLGroup.ps1
    }

    If ($Global:chkUpdDLSecGrp.Checked -eq "Checked")
    {
        invoke-expression -command .\UpdateGroupMembership.ps1
        write-host""
        write-host "List Updating Complete" -ForegroundColor Magenta
    }

    If ($Global:chkUpdAuthUser.Checked -eq "Checked")
    {
        invoke-expression -command .\ModifyDLAuthUsers.ps1
        write-host""
        write-host "Authorized Users Update Commplete" -ForegroundColor Magenta
    }

    If ($Global:chkDLMem.Checked -eq "Checked")
    {
        Write-Host "Enter the Name of the group that you would like the get the membership for " -ForegroundColor Yellow -NoNewline
        $Grp = Read-Host
    	$DynDstExits = [bool](Get-DynamicDistributionGroup $Grp -ErrorAction SilentlyContinue)
        if ($DynDstExits = "True")
        {
            $colu = Get-DistributionGroupMember $Grp -ResultSize Unlimited
            Write-Host "Number of members in the group " -ForegroundColor Yellow -NoNewline
            Write-Host $Grp -ForegroundColor Red -NoNewline
            Write-Host " has " -ForegroundColor Yellow -NoNewline
            write-host $colu.count "members" -ForegroundColor Red
            pause
        }
        else
        {
             Write-Host "Dynamic Distribution List Does Not Exist - No Changes Made" -ForegroundColor Red
        }
    }

    If ($Global:chkExtAccess.Checked -eq "Checked")
    {
        Write-Host "Enter the Name of the group that you would like Grant or Deny External Addresses to use: " -ForegroundColor Yellow -NoNewline
        $Grp = Read-Host
        
        $DstExists = [bool](Get-DistributionGroup $Grp -ErrorAction SilentlyContinue)
        if ($DstExists -eq "True")
        {
            $GrpExt = Get-DistributionGroup $Grp
            write-host "Group Owners:  " $GrpExt.ManagedBy

            if ($GrpExt.RequireSenderAuthenticationEnabled -eq "True")
            {
                write-host "Is the requestor an owner of this distribution list? (Y/N) " -ForegroundColor Yellow -NoNewline
                $Owner = Read-host

                If ($Owner -eq"Y")
                {
                    Write-Host "External addresses are not allowed to send to this group would you like to allow this (Y/N) " -ForegroundColor Yellow -NoNewline
                    $ExtChg = Read-Host
                    if ($ExtChg -eq "Y")
                    {
                        Set-DistributionGroup $Grp -RequireSenderAuthenticationEnabled $false
                        Write-Host "Setting changed to allow external addresses to send to the group it may take 15-20 minutes for this setting to synchronize in the O365 environment" -ForegroundColor Yellow
                    }
                    else
                    {
                        Write-Host "No changes made to the settings for allowing external addresses to send to the group"
                    }
                    pause
                }
                else
                {
                    write-host "No changes made without a curent owner approving the changes" -ForegroundColor Red
                }
            }
            else
            {
                write-host "Is the requestor an owner of this distribution list? (Y/N) " -ForegroundColor Yellow -NoNewline
                $Owner = Read-host

                If ($Owner -eq "Y")
                {
                    Write-Host "External addresses are allowed to send to this group would you like to disallow this (Y/N) " -ForegroundColor Yellow -NoNewline
                    $ExtChg = Read-Host
                    if ($ExtChg -eq "Y")
                    {
                        Set-DistributionGroup $Grp -RequireSenderAuthenticationEnabled $true
                        Write-Host "Setting changed to not allow external addresses to send to the group it may take 15-20 minutes for this setting to synchronize in the O365 environment" -ForegroundColor Yellow
                    }
                    else
                    {
                        Write-Host "No changes made to the settings for allowing external addresses to send to the group"
                    }
                    pause
                }
                else
                {
                    write-host "No changes made without a curent owner approving the changes" -ForegroundColor Red
                }
            }
        }
        else
        {
            Write-Host "Distribution List Does Not Exist - No Changes Made" -ForegroundColor Red
        }                 
    }

    If ($Global:chkAddRemAlias.Checked -eq "Checked")
    {
        Write-Host "Add/Remove Additional eMail Alias to Distribution Group" -ForegroundColor Magenta
        Write-Host "Enter the name of the Distribution List " -ForegroundColor Yellow -NoNewline
        $Grp = Read-Host
        $DstExists = [bool](Get-DistributionGroup $Grp -ErrorAction SilentlyContinue)
        if ($DstExists -eq "True")
        {
            $GrpAlias = Get-DistributionGroup $Grp                    
            write-host "Distribution List Owners on" $Grp ": " $GrpAlias.ManagedBy -ForegroundColor Yellow
            Write-Host "Current Email aliases on" $Grp "   : " -ForegroundColor Yellow -NoNewline
            write-host $GrpAlias.EmailAddresses 
            write-host ""
            write-host "Is the requestor an owner of this distribution list or have you obtained owner approval (Y/N)? " -ForegroundColor Red -NoNewline
            $OwnrAns = Read-Host
            If ($OwnrAns -eq "Y")
            {
                write-host "Do you wish to Add or Remove an Alias enter (A = Add/R = Remove)? " -ForegroundColor Yellow -NoNewline
                $AddRem = Read-Host
                If ($AddRem -eq "A")
                {
                    write-host ""
                    write-host "Enter the eMail Alias you would like to add: " -ForegroundColor Yellow -NoNewline
                    $NewAlias = Read-Host
                    $AddAlias = "smtp:" + $NewAlias
                    Set-DistributionGroup $Grp -EmailAddresses @{Add=$AddAlias}
                    Get-DistributionGroup $Grp |ft *Addresses*
                    Write-Host "Should this be made the new Primary SMTP Address for this Mailbox (Y/N) ?" -ForegroundColor Yellow -NoNewline
                    $NewPrim = Read-Host
                    If ($NewPrim -eq "Y")
                    {
                        Set-DistributionGroup $Grp -PrimarySmtpAddress $NewAlias
                    }
                }
                elseif ($AddRem -eq "R")
                {
                    write-host ""
                    write-host "Enter the eMail Alias you would like to remove: " -ForegroundColor Yellow -NoNewline
                    $RemAlias = Read-Host
                    $GrpDet = Get-DistributionGroup $Grp
                    If ($RemAlias -eq $GrpDet.PrimarySMTPAddress)
                    {
                        write-host "This is the Primary SMTP Alias for this group you must assign a new Primary Alias.  Enter New Primary Alias: " -ForegroundColor Yellow -NoNewline
                        $NewAlias = Read-Host
                        Set-DistributionGroup $Grp -PrimarySmtpAddress $NewAlias                                                       
                    }
                    $RemAlias = "smtp:" + $RemAlias
                    Set-DistributionGroup $Grp -EmailAddresses @{Remove=$RemAlias}
                }
                else
                {
                    Write-Host "No changes made to the configured aliases for this Distribution list"
                }
                Get-DistributionGroup $Grp |ft EmailAddresses
                pause
            }
            else
            {
                Write-Host "Please obtain approval to make this change" -BackgroundColor Red
                pause
            }
        }
        else
        {
             Write-Host "Distribution List Does Not Exist - No Changes Made" -ForegroundColor Red
        }
    }

    If ($Global:chkAddMem.Checked -eq "Checked")
    {
        #  Declare Drive | Folders | and Files
	    $FileName		= "AddDgMembers"
	    $LogDrive		= "e:\Automation"
	    $LogFolder		= "\" + $FileNAme
	    $LogDirectory	= $LogDrive + $LogFolder + "\"
	    $LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	    $InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
	    $ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"

        if (Test-Path $InputFile)
        {
            write-host "Importing the e:\Automation\AddDGMembers\Input-AddDgMembers.csv Input file" -ForegroundColor Yellow
            $InpFile = Import-Csv $InputFile
            $DstExists = [bool](Get-DistributionGroup $InpFile.DgMail -ErrorAction SilentlyContinue)
            $UfgExists = [bool](Get-UnifiedGroup $InpFile.DgMail -ErrorAction SilentlyContinue)
            If (($DstExists -eq "True") -or ($UfgExists -eq "True"))
            {
                $Members = $InpFile.MemberMail.Replace(" ","")
    	        $addMember = $Members.split(",")
                write-host "Number of Members Listed in Input File: " $addmember.count
                if ($? -eq $true)
                {
                    $CurrentMembers = Get-DistributionGroupMember $InpFile.DgMail -ResultSize Unlimited
		            ForEach ($member in $addMember)
                    {
                        if ($member.contains("@"))
                        {
	                        if (Get-Mailbox $member -ErrorAction SilentlyContinue)
                            {
                                if ($CurrentMembers -match (get-mailbox $Member).Name)
                                {
                                    write-host $member " already a member."
           				            $LineToWrite = "`t" + "Warn" + "`t" + $member + " - is already a member of Distribution Group. " + $Grp.GrpName
                                }
                                else
                                {
                                    If ($DstExists -eq "True")
                                    {
                                        Add-DistributionGroupMember $InpFile.DgMail -Member $Member -BypassSecurityGroupManagerCheck
    	  				                Write-Host "Added member:" $member
		    			                $LineToWrite = "`t" + "PASS" + "`t" + $member + "`t" + "Added to Distribution/Security Group" + "`t" + $InpFile.DgDisplayName + "`n"
                                    }
                                    else
                                    {
                                        Add-UnifiedGroupLinks -Identity $InpFile.DgMail -LinkType Member -links $Member -Confirm:$false								    
    					                Write-Host "Added member:" $member
						                $LineToWrite = "`t" + "PASS" + "`t" + $member + "`t" + "Added to Unified Group" + "`t" + $InpFile.DgDisplayName + "`n"
                                    }
                                }
					        }
		    		        else
                            {
						        write-host "ERROR finding member: " $member
						        $LineToWrite = "`t" + "FAIL" + "`t" + $member + "`t" + "ERROR finding recipient" + "`t" + $InpFile.DgDisplayName + "`n"
					        }
				        }
				        else
                        {
                            write-host "ERROR invalid member: " $member
					        $LineToWrite = "`t" + "FAIL" + "`t" + $member + "`t" + "ERROR invalid recipient" + "`t" + $InpFile.DgDisplayName + "`n"
		                }
                        WriteReportEvent
                    }
                    Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")
                }
            }	
		    else
            {
                write-host "Distribution/Security or Unified Group Does Not Exist"
		    }
        }
        else
        {
            Write-Host "Enter the name of the Security Group, Distribution List or Unified Group " -ForegroundColor Yellow -NoNewline
            $Grp = Read-Host
            $DstExists = [bool](Get-DistributionGroup $Grp -ErrorAction SilentlyContinue)
            $UfgExists = [bool](Get-UnifiedGroup $Grp -ErrorAction SilentlyContinue)
            If (($DstExists -eq "True") -or ($UfgExists -eq "True"))
            {
                If ($DstExists -eq "True")
                {
                    $GrpDet = Get-DistributionGroup $Grp
                }
                else
                {
                    $GrpDet = Get-UnifiedGroup $Grp
                }
                write-host ""
                write-host "Distribution/Security/Unified Group Owner(s): " $GrpDet.ManagedBy 
                write-host ""
                write-host "Is the requestor an owner of this security or distribution list/unified group or have you obtained owner approval (Y/N)? " -ForegroundColor Red -NoNewline
                $OwnrAns = Read-Host
                
                If ($OwnrAns -eq "Y")
                {
                    Do
                    {
                        Write-Host "Enter the Employee Number of the individual to add to the membership " -ForegroundColor Yellow -NoNewline
                        $Member = Read-Host
                        $Member = $Member + "@global.ul.com"
                        
                        If ($DstExists -eq "True")
                        {
                            Add-DistributionGroupMember $grp -Member $Member -BypassSecurityGroupManagerCheck
                        }
                        else
                        {
                            Add-UnifiedGroupLinks -Identity $Grp -LinkType Members -links $Member -Confirm:$false
                        }

                        Write-Host $Member "added to " $Grp
                        Write-Host ""
                        Write-Host "Do you have more members to add to this group (Y/N)? " -ForegroundColor Yellow -NoNewline
                        $OwnrAns = Read-Host
                    } while  ($OwnrAns -eq "Y")
                }
           }
           else
           {
                Write-Host "Distribution/Security or Unified Group Does Not Exist" -ForegroundColor Red
           }
        }
        else
        {
            write-host "No changes made without owner approval"
        }
    }
    
    If (($Global:chkRemMem.Checked -eq "Checked") -or ($Global:chkUniGrpMem.Checked -eq "Checked"))
    {
        #  Declare Drive | Folders | and Files
        $FileName		= "RemoveDgMembers"
	    $LogDrive		= "e:\Automation"
	    $LogFolder		= "\" + $FileNAme
	    $LogDirectory	= $LogDrive + $LogFolder + "\"
	    $LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	    $InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
	    $ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"

        if (Test-Path $InputFile)
        {
            write-host "Importing the e:\Automation\RemoveDGMembers\Input-RemoveDgMembers.csv Input file" -ForegroundColor Yellow
            $InpFile = Import-Csv $InputFile
            $DstExists = [bool](Get-DistributionGroup $InpFile.DgMail -ErrorAction SilentlyContinue)
            $UfgExists = [bool](Get-UnifiedGroup $InpFile.DgMail -ErrorAction SilentlyContinue)
            If (($DstExists -eq "True") -or ($UfgExists -eq "True"))
            {
                $Members = $InpFile.MemberMail.Replace(" ","")
                $removeMember = $Members.split(",")
                write-host "Number of Members Listed in Input File: " $removemember.count
	            if ($? -eq $true)
                {
                    $CurrentMembers = Get-DistributionGroupMember $InpFile.DgMail -ResultSize Unlimited
	                ForEach ($member in $removeMember)
                    {
		                if ($member.contains("@"))
                        {
	                        if (Get-Mailbox $member -ErrorAction SilentlyContinue)
                            {
                                if ($CurrentMembers -match (get-mailbox $Member).Name)
                                {
                                    write-host $member " already a member."
        		                    $LineToWrite = "`t" + "Warn" + "`t" + $member + " - is already a member of Distribution Group. " + $Grp.GrpName
                                }
                                else
                                {
                                    If ($DstExists -eq "True")
                                    {
                                        Remove-DistributionGroupMember $InpFile.DgMail -Member $Member -BypassSecurityGroupManagerCheck -Confirm:$false
    	    		                    Write-Host "Added member:" $member
		        	                    $LineToWrite = "`t" + "PASS" + "`t" + $member + "`t" + "Removed from Distribution/Security Group" + "`t" + $InpFile.DgDisplayName + "`n"
                                    }
                                    else
                                    {
                                        Remove-UnifiedGroupLinks -Identity $InpFile.DgMail -LinkType Member -links $Member -Confirm:$false								    
    			                        Write-Host "Added member:" $member
					                    $LineToWrite = "`t" + "PASS" + "`t" + $member + "`t" + "Removed from Unified Group" + "`t" + $InpFile.DgDisplayName + "`n"
                                    }
                                }
					        }
		    		        else
                            {
                                write-host "ERROR finding member: " $member
						        $LineToWrite = "`t" + "FAIL" + "`t" + $member + "`t" + "ERROR finding recipient" + "`t" + $InpFile.DgDisplayName + "`n"
					        }
				        }
	                    else
                        {
		                    write-host "ERROR invalid member: " $member
			                $LineToWrite = "`t" + "FAIL" + "`t" + $member + "`t" + "ERROR invalid recipient" + "`t" + $InpFile.DgDisplayName + "`n"
			            }
                        WriteReportEvent
			        }
                    Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")
                }
            }	
            else
            {
                write-host "Distrubtion/Security or Unified Group Does Not Exist"
	        }
        }
        else
        {
            Write-Host "Enter the name of the Security Group, Distribution List or Unified Group " -ForegroundColor Yellow -NoNewline
            $Grp = Read-Host
            $DstExists = [bool](Get-DistributionGroup $Grp -ErrorAction SilentlyContinue)
            $UfgExists = [bool](Get-UnifiedGroup $Grp -ErrorAction SilentlyContinue)
            If (($DstExists -eq "True") -or ($UfgExists -eq "True"))
            {
                If ($DstExists -eq "True")
                {
                    $GrpDet = Get-DistributionGroup $Grp
                }
                else
                {
                    $GrpDet = Get-UnifiedGroup $Grp
                }
                write-host ""
                write-host "Distribution/Security/Unified Group Owner(s): " $GrpDet.ManagedBy 
                write-host ""
                write-host "Is the requestor an owner of this security or distribution list/unified group or have you obtained owner approval (Y/N)? " -ForegroundColor Red -NoNewline
                $OwnrAns = Read-Host
                
                If ($OwnrAns -eq "Y")
                {
                    Do
                    {
                        Write-Host "Enter the Employee Number of the individual to remove to the membership " -ForegroundColor Yellow -NoNewline
                        $Member = Read-Host
                        $Member = $Member + "@global.ul.com"
                        
                        If ($DstExists -eq "True")
                        {
                            Remove-DistributionGroupMember $grp -Member $Member -BypassSecurityGroupManagerCheck -Confirm:$false
                        }
                        else
                        {
                            Remove-UnifiedGroupLinks -Identity $Grp -LinkType Members -links $Member -Confirm:$false
                        }

                        Write-Host $Member "removed to " $Grp
                        Write-Host ""
                        Write-Host "Do you have more members to remove to this group (Y/N)? " -ForegroundColor Yellow -NoNewline
                        $OwnrAns = Read-Host
                    } while  ($OwnrAns -eq "Y")
                }
            }
            else
            {
                
                Write-Host "Distribution/Security or Unified Group Does Not Exist" -ForegroundColor Red
            }
        }
        else
        {
            write-host "No changes made without owner approval"
        }
    }

    If ($Global:chkDSTMem.Checked -eq "Checked")
    {
        Import-Module ExchangeOnlineManagement
        connect-ExchangeOnline -Credential $Global:LiveCred
        Write-Host "Enter the Name of the DST group that you would like the membership for " -ForegroundColor Yellow -NoNewline
        $DSTGrp = Read-Host
        $DstExists = [bool](Get-DynamicDistributionGroup $DSTGrp -ErrorAction SilentlyContinue)
        if ($DstExists -eq "True")
        {
            $g = Get-DynamicDistributionGroup $DSTGrp
            $colu = Get-DynamicDistributionGroupMember $g.name -ResultSize unlimited
#            $colu = Get-Recipient -RecipientPreviewFilter $g.RecipientFilter -OrganizationalUnit $g.RecipientContainer -ResultSize Unlimited
            Write-Host
            Write-Host "     Number of members in this group " -ForegroundColor Yellow -NoNewline
            Write-Host $colu.count -ForegroundColor Yellow
            Write-Host
        	Write-Host "Writing list of group members to c:\temp\DSTMembership.csv"
			$colu.PrimarySMTPAddress | Out-File c:\temp\DSTMembership.csv
            Write-Host "Request complete hit return to continue" -NoNewline
            $Cont = Read-Host 
        }
        else
        {
            Write-Host "Distribution/Security List Does Not Exist - No Changes Made" -ForegroundColor Red
        }
    }

    If ($Global:chkNewDST.Checked -eq "Checked") 
    {
        Invoke-Expression .\NewDynamicDistributionGroup.ps1
    }

    If ($Global:chkRemDST.Checked -eq "Checked") 
    {
        Invoke-Expression .\RemoveDynamicDistributionGrp.ps1
#        Invoke-Expression .\RemoveDynamicDistributionGroup-New.ps1
    }

    If ($Global:chkCreUniGrp.Checked -eq "Checked") 
    {
        Invoke-Expression .\NewUnifiedGrp.ps1
    }

    If ($Global:chkTeamShrPtSite.Checked -eq "Checked")
    {
        Write-Host "Enter the Name of the O365 Group to check for an active Sharepoint Site: " -ForegroundColor Yellow -NoNewline
        $Grp = Read-Host
        $O365Exists = [bool](Get-UnifiedGroup $Grp -ErrorAction SilentlyContinue)
        if ($O365Exists -eq "True")
        {
            $O365Grp = Get-UnifiedGroup $Grp
            $SPSiteExists = [bool](get-sposite $O365Grp.SharePointSiteUrl)
            If ($SPSiteExists -eq "True")
            {
                write-host "SharePointSite" $0365Grp.SharePointSiteUrl "Exists" -ForegroundColor Green
            }
            else
            {
                write-host "No SharepointSite " $0365Grp.SharePointSiteUrl "Exists" -ForegroundColor Red
            }
        }
        else
        {
            write-host "The O365 Group" $O365Grp "does not exist" -ForegroundColor Red
        }
    }

    If ($Global:chkSoftDelObj.Checked -eq "Checked")
    {
        write-host "Connecting to AzureAD...... " -ForegroundColor Cyan
     	Connect-AzureAD -Credential $Global:LiveCred | Out-Null
        Write-Host "Enter the Name of the O365 Group to check for a soft deleted Sharepoint Site: " -ForegroundColor Cyan -NoNewline
        $Grp = Read-Host
        $O365Exists = [bool](Get-AzureADMSDeletedGroup -SearchString $Grp -ErrorAction SilentlyContinue)
        if ($O365Exists -eq "True")
        {
            $O365Grp = Get-AzureADMSDeletedGroup -SearchString $Grp
            $SPSiteExists = [bool](get-SPODeletedSite $O365Grp.SharePointSiteUrl)
            If ($SPSiteExists -eq "True")
            {
                write-host "SharePointSite" $0365Grp.SharePointSiteUrl "Exists. Do you want to permanently Delete it (Y/N)?: " -ForegroundColor Green -NoNewline
                $SPODel = Read-Host
                If ($SPODel -eq "Y")
                {
                    Write-Host "Permanently removing the soft deleted sharepoint site" -ForegroundColor Yellow
                    Remove-SPODeletedSite $O365Grp.SharePointSiteUrl -Confirm:$False
                }
                else
                {
                    write-host "Not removing the soft deleted sharepoint site" -ForegroundColor Red
                }
            }
            else
            {
                write-host "No SharepointSite " $0365Grp.SharePointSiteUrl "Exists" -ForegroundColor Red
            }
        }
        else
        {
            write-host "A deleted O365 Group named" $Grp "does not exist" -ForegroundColor Red
        }
#      	Remove-AzureADMSDeletedDirectoryObject -Id $O365Grp.Id
        pause
    }

    If ($Global:chkMiscCmd.Checked -eq "Checked")
    {
        write-host "Enter (1) to Show All O365 Groups or (2) to Show a Specific Group: " -ForegroundColor Cyan -NoNewline
        $ShwGrps = Read-Host
               
        switch ($ShwGrps)
        {

            1
            {
                $o365groups = Get-UnifiedGroup
                foreach ($o365group in $o365groups)
                {
                    $TeamEna = "Enabled"
                    try
                    {
                        $teamschannels = Get-TeamChannel -GroupId $o365group.ExternalDirectoryObjectId
                        [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; TeamsEnabled = $true}
                    }
                    catch
                    {
                        $ErrorCode = $_.Exception.ErrorCode
                        switch ($ErrorCode)
                        {
                            "404"
                            {
                               [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; TeamsEnabled = $false}
                               $TeamEna = "NotEnabled"
                               break;
                            }
                            "403"
                            {
                                [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; TeamsEnabled = $true}
                                break;
                            }
                            default
                            {
                                Write-Error ("Unknown ErrorCode trying to 'Get-TeamChannel -GroupId {0}' :: {1}" -f $o365group, $ErrorCode)
                                $TeamEna = "Unknown"
                            }
                        }
                    }
                }
            }
            2
            {
                Write-host "Enter the Name of the Group: " -ForegroundColor Cyan -NoNewline
                $Grp = Read-Host
                $o365groups = Get-UnifiedGroup $Grp
                foreach ($o365group in $o365groups)
                {
                    $TeamEna = "Enabled"
                    try
                    {
                        $teamschannels = Get-TeamChannel -GroupId $o365group.ExternalDirectoryObjectId
                        [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; TeamsEnabled = $true}
                    }
                    catch
                    {
                        $ErrorCode = $_.Exception.ErrorCode
                        switch ($ErrorCode)
                        {
                            "404"
                            {
                                [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; TeamsEnabled = $false}
                                $TeamEna = "NotEnabled"
                                break;
                            }
                            "403"
                            {
#                                [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; TeamsEnabled = $true}
                                break;
                            }
                            default
                            {
                                Write-Error ("Unknown ErrorCode trying to 'Get-TeamChannel -GroupId {0}' :: {1}" -f $o365group, $ErrorCode)
                                $TeamEna = "Unknown"
                            }
                        }
                    }
                    write-host "Status:    " $TeamEna
                    write-host "GroupID:   " $O365group.ExternalDirectoryObjectId
                    write-host "GroupName: " $o365group.DisplayName
                }
                pause
            }
        }
    }
    Build-DLMenuForm
    Publish-Form
}While ($Global:Result -eq "OK")	