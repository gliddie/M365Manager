<#   
#================================================================================ 
# Name: Remove Shared Mailbox
#
# 10/12/2020 - SAG - Redesigned using forms and the old RemoveSharedMailbox.ps1 script
# 05/17/2022 - SAG - Modified the report file to be written out to the E: drive
# 11/10/2022 - SAG - Modified to use a GUI interface
# 04/07/2023 - SAG - Added group member empno details into log information
#================================================================================ 
#>  

function Hide-Details
{
    $Global:form.Width = 600 ; $Global:form.Height = 120
    $Script:txtDispName.ReadOnly = $False
    $Script:ButGetDL.Visible = $True
    $Script:lblTicketNo.Visible = $False
    $Script:txtTicketNo.Visible = $False
    $Script:lblGrpOwnr.Visible = $False
    $Script:txtGrpOwnr.Visible = $False
    $Script:lblCurrTip.Visible = $False
    $Script:txtCurrTip.Visible = $False
    $Script:lblExtUse.Visible = $False
    $Script:chkAllow.Visible = $False
    $Script:chkRestrict.Visible = $False
    $Script:lblEDGrpName.Visible = $False
    $Script:txtEDGrpName.Visible = $False
    $Script:lblEDGrpMbr.Visible = $False
    $Script:txtEDGrpMbr.Visible = $False
    $Script:lblAUGrpName.Visible = $False
    $Script:txtAUGrpName.Visible = $False
    $Script:lblAUGrpMbr.Visible = $False 
    $Script:txtAUGrpMbr.Visible = $False
    $Script:lblREGrpName.Visible = $False
    $Script:txtREGrpName.Visible = $False
    $Script:lblREGrpMbr.Visible = $False
    $Script:txtREGrpMbr.Visible = $False
    $Script:lblRptFile.Visible = $False
    $Script:txtRptFile.Visible = $False
    $Script:chkTemplate.Visible = $False
    $Global:OKButton.Visible = $False
    $Script:MsgTo = ""
    $Script:txtGrpOwnr.Text = "None Found"
    $Script:txtEDGrpName.Text = ""
    $Script:txtEDGrpMbr.Text = ""
    $Script:txtAUGrpName.Text = ""
    $Script:txtAUGrpMbr.Text = ""
    $Script:txtREGrpName.Text = ""
    $Script:txtREGrpMbr.Text = ""
}

function UnHide-Details
{
    $Global:form.Width = 700 ; $Global:form.Height = 580
    $Script:txtDispName.ReadOnly = $True
    $Script:ButGetDL.Visible = $False
    $Script:lblTicketNo.Visible = $True
    $Script:txtTicketNo.Visible = $True
    $Script:lblGrpOwnr.Visible = $True
    $Script:txtGrpOwnr.Visible = $True
    $Script:lblCurrTip.Visible = $True
    $Script:txtCurrTip.Visible = $True
    $Script:lblExtUse.Visible = $True
    $Script:chkAllow.Visible = $True
    $Script:chkRestrict.Visible = $True
    $Script:lblEDGrpName.Visible = $True
    $Script:txtEDGrpName.Visible = $True
    $Script:lblEDGrpMbr.Visible = $True
    $Script:txtEDGrpMbr.Visible = $True
    $Script:lblAUGrpName.Visible = $True
    $Script:txtAUGrpName.Visible = $True
    $Script:lblAUGrpMbr.Visible = $True 
    $Script:txtAUGrpMbr.Visible = $True
    $Script:lblREGrpName.Visible = $True
    $Script:txtREGrpName.Visible = $True
    $Script:lblREGrpMbr.Visible = $True
    $Script:txtREGrpMbr.Visible = $True
    $Script:lblRptFile.Visible = $True
    $Script:txtRptFile.Visible = $True
    $Script:chkTemplate.Visible = $True
}

function Get-EmailDate
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Email Date" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Width = 250; $Global:form.Height = 130

    ## Date of Original Communication
    $Script:lblEmailDate = New-Object System.Windows.Forms.Label   
        $Script:lblEmailDate.Text = "Email Date:"
        $Script:lblEmailDate.Top = 10; $Script:lblEmailDate.Left = 30; $Script:lblEmailDate.Width=70
        $Script:lblEmailDate.Visible = $True
        $Global:form.Controls.Add($Script:lblEmailDate)    # Add to Form 
    $Script:txEmailDate = New-Object Windows.Forms.DateTimePicker
        $Script:txEmailDate.Top = 10; $Script:txEmailDate.Left = 110; $Script:txEmailDate.Width = 100;
        $Script:txEmailDate.Format = [windows.forms.datetimepickerFormat]::custom
        $Script:txEmailDate.CustomFormat = "MM/dd/yyyy"
        $Script:txEmailDate.Visible = $True
        $Script:txEmailDate.Text = (get-date) 
        $Global:form.Controls.Add($Script:txEmailDate)    # Add to Form

    Add-FormStandardButtons
    $Global:OkButton.Visible = $False
    $Global:CancelButton.Text = "Continue"
    Publish-Form
}

function Build-ShrMbxUseDetailsForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Change Shared Mailbox Use" 
    $Global:form.StartPosition = "CenterScreen"

    Add-FormStandardButtons
    $Global:OKButton.AutoSize = $True
    $Global:OKButton.Add_Click({
        If ($Script:txtTicketNo.Text -eq "Email Request")
        {
            Get-EmailDate
        }
    })

    ## Get Details to complete Use change
    $TopLoc = 10
    ## Display Name
    $Script:lblDispName = New-Object System.Windows.Forms.Label   
        $Script:lblDispName.Text = "Shared Mailbox Name:"
        $Script:lblDispName.Top = $TopLoc; $Script:lblDispName.Left = 10; $Script:lblDispName.Width=120
        $Global:form.Controls.Add($Script:lblDispName)    # Add to Form 
    $Script:txtDispName = New-Object Windows.Forms.TextBox  
        $Script:txtDispName.TabIndex = 0 # set Tab Order 
        $Script:txtDispName.Top = $TopLoc; $Script:txtDispName.Left = 130; $Script:txtDispName.Width = 220;
        $Script:txtDispName.Text = ""
        $Global:form.Controls.Add($Script:txtDispName)    # Add to Form
        $Global:InputFocus = $Script:txtDispName
        $Script:txtDispName.Add_Click({
            Hide-Details
         })

    $Script:ButGetDL = New-Object Windows.Forms.Button
        $Script:ButGetDL.Location = New-object System.Drawing.Size(370,$TopLoc)
        $Script:ButGetDL.Size = new-Object System.Drawing.Size(200,20)
        $Script:ButGetDL.Text = "Get Shared Mailbox Details"
        $Script:ButGetDL.TabIndex = 0
        $Global:form.Controls.Add($Script:ButGetDL)
        $Script:ButGetDL.Add_Click({
            write-host "`nRetreiving Account Details for Mailbox:" $Script:txtDispName.Text -ForegroundColor Cyan
            $MbxExists = [bool]($Script:Mbx = Get-Mailbox $Script:txtDispName.Text -ErrorAction SilentlyContinue)
            If ($MbxExists -eq $True)
            {
                $Script:txtDispName.Text = $Script:Mbx.DisplayName
                Check-Reconnect
                Get-MailboxDetails
                UnHide-Details
                $Script:Date = get-date -Format "yyyy-MMdd"
                $Script:txtRptFile.Text = "e:\Automation\ShrMbxUse\Report\ShrMbxUse-" + ($Script:txtDispName.Text -Replace("['.()/\- ]","")) + "-Date" + $Script:Date + ".log"
                $Global:OKButton.Text = "Change"
                $Global:OKButton.visible = $True
            }
            else
            {
                $Script:txtDispName.Text = "Invalid"
                $Script:ButGetDL.Visible = "True"
            }
            })

    $TopLoc = $TopLoc + 30
    ## Ticket Number
    $Script:lblTicketNo = New-Object System.Windows.Forms.Label   
        $Script:lblTicketNo.Text = "Ticket Number:"  
        $Script:lblTicketNo.Top = $TopLoc ; $Script:lblTicketNo.Left = 10; $Script:lblTicketNo.Width=120
        $Script:lblTicketNo.Visible = $False
        $Global:form.Controls.Add($Script:lblTicketNo)    # Add to Form              
    $Script:txtTicketNo = New-Object Windows.Forms.TextBox
        $Script:txtTicketNo.TabIndex = 0 # set Tab Order 
        $Script:txtTicketNo.Top = $TopLoc; $Script:txtTicketNo.Left = 130; $Script:txtTicketNo.Width = 180;
        $Global:form.Controls.Add($Script:txtTicketNo)    # Add to Form

    $TopLoc = $TopLoc + 30
    ## Mailbox Owner
    $Script:lblGrpOwnr = New-Object System.Windows.Forms.Label   
        $Script:lblGrpOwnr.Text = "Mailbox Owner(s):"  
        $Script:lblGrpOwnr.Top = $TopLoc; $Script:lblGrpOwnr.Left = 10; $Script:lblGrpOwnr.Width=120;
        $Global:form.Controls.Add($lblGrpOwnr)    # Add to Form 
    $Script:txtGrpOwnr = New-Object Windows.Forms.TextBox  
        $Script:txtGrpOwnr.TabIndex = 0 # set Tab Order 
        $Script:txtGrpOwnr.Top = $TopLoc; $Script:txtGrpOwnr.Left = 130; $Script:txtGrpOwnr.Width = 320;
        $Script:txtGrpOwnr.Text = "None Found"
        $Script:txtGrpOwnr.ReadOnly = $True
        $Script:txtGrpOwnr.Text = ""
        $Global:form.Controls.Add($Script:txtGrpOwnr)    # Add to Form 

    $TopLoc = $TopLoc + 30
    ## Current Mailbox Tip
    $Script:lblCurrTip = New-Object System.Windows.Forms.Label   
        $Script:lblCurrTip.Text = "Current Tip:"  
        $Script:lblCurrTip.Top = $TopLoc; $Script:lblCurrTip.Left = 10; $Script:lblCurrTip.Width=120;
        $Global:form.Controls.Add($lblCurrTip)    # Add to Form 
    $Script:txtCurrTip = New-Object Windows.Forms.TextBox  
        $Script:txtCurrTip.TabIndex = 0 # set Tab Order 
        $Script:txtCurrTip.Top = $TopLoc; $Script:txtCurrTip.Left = 130; $Script:txtCurrTip.Width = 500;
        $Script:txtCurrTip.ReadOnly = $True
        $Script:txtCurrTip.Text = ""
        $Global:form.Controls.Add($Script:txtCurrTip)    # Add to Form
        
    $TopLoc = $TopLoc + 30
    ## Current Mailbox Use Setting
    $Script:lblExtUse = New-Object System.Windows.Forms.Label   
        $Script:lblExtUse.Text = "External Access:"  
        $Script:lblExtUse.Top = $TopLoc; $Script:lblExtUse.Left = 10; $Script:lblExtUse.Width=120;
        $Global:form.Controls.Add($lblExtUse)    # Add to Form 
    $Script:chkAllow = New-Object Windows.Forms.RadioButton
        $Script:chkAllow.Left = 130; $Script:chkAllow.Width = 120; $Script:chkAllow.Top = ($TopLoc-5)
        $Script:chkAllow.Text = "Allow External Use" 
        $Global:form.Controls.Add($Script:chkAllow)
    $Script:chkRestrict = New-Object Windows.Forms.RadioButton
        $Script:chkRestrict.Left = 260; $Script:chkRestrict.Width = 150; $Script:chkRestrict.Top = ($TopLoc-5)
        $Script:chkRestrict.Text = "Restrict External Use" 
        $Global:form.Controls.Add($Script:chkRestrict)

    $TopLoc = $TopLoc + 30
        ## Editor Old Group Name
        $Script:lblEDGrpName = New-Object System.Windows.Forms.Label
        $Script:lblEDGrpName.Text = ".ED Group:"  
        $Script:lblEDGrpName.Top = $TopLoc ; $Script:lblEDGrpName.Left = 10; $Script:lblEDGrpName.Width=120;
        $Global:form.Controls.Add($Script:lblEDGrpName)    # Add to Form 
    $Script:txtEDGrpName = New-Object Windows.Forms.TextBox  
        $Script:txtEDGrpName.TabIndex = 0 # set Tab Order 
        $Script:txtEDGrpName.Top = $TopLoc; $Script:txtEDGrpName.Left = 130; $Script:txtEDGrpName.Width = 220;  
        $Script:txtEDGrpName.Text = "Not Found"
        $Script:txtEDGrpName.ReadOnly = $True
        $Global:form.Controls.Add($Script:txtEDGrpName)    # Add to Form 

    $TopLoc = $TopLoc + 30
    ## Editor Group Members
    $Script:lblEDGrpMbr = New-Object System.Windows.Forms.Label   
        $Script:lblEDGrpMbr.Text = ".ED Member(s):"  
        $Script:lblEDGrpMbr.Top = $TopLoc; $Script:lblEDGrpMbr.Left = 10; $Script:lblEDGrpMbr.Width=120;
        $Global:form.Controls.Add($Script:lblEDGrpMbr)    # Add to Form 
    $Script:txtEDGrpMbr = New-Object Windows.Forms.TextBox
        $Script:txtEDGrpMbr.TabIndex = 0 # set Tab Order
        $Script:txtEDGrpMbr.Location = New-Object System.Drawing.Size(130,$TopLoc)
        $Script:txtEDGrpMbr.Size = New-Object system.Drawing.Size(520,40)
        $Script:txtEDGrpMbr.BackColor = "#EEEDF0"
        $Script:txtEDGrpMbr.MultiLine = $true
        $Script:txtEDGrpMbr.ScrollBars = 'Both'
        $Script:txtEDGrpMbr.Text = $Script:EDMembers
        $Global:form.Controls.Add($Script:txtEDGrpMbr)    # Add to Form
        $TopLoc = $TopLoc + 20

    $TopLoc = $TopLoc + 30
    # Author Group Name
    $Script:lblAUGrpName = New-Object System.Windows.Forms.Label
        $Script:lblAUGrpName.Text = ".AU Group:" 
        $Script:lblAUGrpName.Top = $TopLoc; $Script:lblAUGrpName.Left = 10; $Script:lblAUGrpName.Width=120;
        $Global:form.Controls.Add($Script:lblAUGrpName)    # Add to Form 
    $Script:txtAUGrpName = New-Object Windows.Forms.TextBox  
        $Script:txtAUGrpName.TabIndex = 0 # set Tab Order 
        $Script:txtAUGrpName.Top = $TopLoc; $Script:txtAUGrpName.Left = 130; $Script:txtAUGrpName.Width = 220;  
        $Script:txtAUGrpName.Text = "Not Found"   # Use Corrent computer name as default
        $Script:txtAUGrpName.ReadOnly = $True
        $Global:form.Controls.Add($Script:txtAUGrpName)    # Add to Form

    $TopLoc = $TopLoc + 30
    ## Author Group Members
    $Script:lblAUGrpMbr = New-Object System.Windows.Forms.Label   
        $Script:lblAUGrpMbr.Text = ".AU Member(s):"  
        $Script:lblAUGrpMbr.Top = $TopLoc; $Script:lblAUGrpMbr.Left = 10; $Script:lblAUGrpMbr.Width=120;
        $Global:form.Controls.Add($Script:lblAUGrpMbr)    # Add to Form 
    $Script:txtAUGrpMbr = New-Object Windows.Forms.TextBox
        $Script:txtAUGrpMbr.TabIndex = 0 # set Tab Order 
        $Script:txtAUGrpMbr.BackColor = "#EEEDF0"
        $Script:txtAUGrpMbr.Location = New-Object System.Drawing.Size(130,$TopLoc)
        $Script:txtAUGrpMbr.Size = New-Object system.Drawing.Size(520,40)
        $Script:txtAUGrpMbr.MultiLine = $true
        $Script:txtAUGrpMbr.ScrollBars = 'Both'
        $Script:txtAUGrpMbr.Text = $Script:AUMembers
        $Global:form.Controls.Add($Script:txtAUGrpMbr)    # Add to Form 
        $TopLoc = $TopLoc + 20

    $TopLoc = $TopLoc + 30
    # Reader Group Name
    $Script:lblREGrpName = New-Object System.Windows.Forms.Label
        $Script:lblREGrpName.Text = ".RE Group:" 
        $Script:lblREGrpName.Top = $TopLoc ; $Script:lblREGrpName.Left = 10; $Script:lblREGrpName.Width=120;
        $Global:form.Controls.Add($Script:lblREGrpName)    # Add to Form 
    $Script:txtREGrpName = New-Object Windows.Forms.TextBox  
        $Script:txtREGrpName.TabIndex = 0 # set Tab Order 
        $Script:txtREGrpName.Top = $TopLoc; $Script:txtREGrpName.Left = 130; $Script:txtREGrpName.Width = 220; 
        $Script:txtREGrpName.Text = "Not Found"
        $Script:txtREGrpName.ReadOnly = $True 
        $Global:form.Controls.Add($Script:txtREGrpName)    # Add to Form 

    $TopLoc = $TopLoc + 30
    ## Reader Group Members
    $Script:lblREGrpMbr = New-Object System.Windows.Forms.Label   
        $Script:lblREGrpMbr.Text = ".RE Member(s):"  
        $Script:lblREGrpMbr.Top = $TopLoc; $Script:lblREGrpMbr.Left = 10; $Script:lblREGrpMbr.Width=120;
        $Global:form.Controls.Add($Script:lblREGrpMbr)    # Add to Form 
    $Script:txtREGrpMbr = New-Object Windows.Forms.TextBox
        $Script:txtREGrpMbr.TabIndex = 0 # set Tab Order 
        $Script:txtREGrpMbr.Location = New-Object System.Drawing.Size(130,$TopLoc)
        $Script:txtREGrpMbr.Size = New-Object system.Drawing.Size(520,40)
        $Script:txtREGrpMbr.BackColor = "#EEEDF0"
        $Script:txtREGrpMbr.MultiLine = $true
        $Script:txtREGrpMbr.ScrollBars = 'Both'
        $Script:txtREGrpMbr.Text = ""
        $Script:txtREGrpMbr.ReadOnly = $True
        $Global:form.Controls.Add($Script:txtREGrpMbr)    # Add to Form
        $TopLoc = $TopLoc + 20

    $TopLoc = $TopLoc + 30
    ## Report File
    $Script:lblRptFile = New-Object System.Windows.Forms.Label   
        $Script:lblRptFile.Text = "Report File:" 
        $Script:lblRptFile.Top = $TopLoc ; $Script:lblRptFile.Left = 10; $Script:lblRptFile.Width=120;
        $form.Controls.Add($Script:lblRptFile)    # Add to Form 
        # 
        $Script:txtRptFile = New-Object Windows.Forms.TextBox
        $Script:txtRptFile.Top = $TopLoc; $Script:txtRptFile.Left = 130; $Script:txtRptFile.Width = 500; $Script:txtRptFile.Height = 60
        $Script:txtRptFile.ReadOnly = $True; $Script:txtRptFile.TabStop = $False
        $Script:txtRptFile.Text = ""
        $Global:form.Controls.Add($Script:txtRptFile)    # Add to Form

    $TopLoc = $TopLoc + 30
    ## To Display the email template
    $Script:chkTemplate = New-Object Windows.Forms.checkbox 
        $Script:chkTemplate.Left = 130; $Script:chkTemplate.Width = 300; $Script:chkTemplate.Top = $TopLoc
        $Script:chkTemplate.Text = "Display Email Template for Change to Mailbox Use" 
        $Script:chkTemplate.Checked = $true   # set a default value
        $Global:form.Controls.Add($Script:chkTemplate)
        $Script:chkTemplate.Add_Click({
            If ($Script:chkTemplate.Checked -eq $True)
            {
                $Global:OKButton.Text = "Change"
                $Script:txtRptFile.Text = "e:\Automation\ShrMbxUse\Report\ShrMbxUse-" + ($Script:txtDispName.Text -Replace("['.()/\- ]","")) + "-Date" + $Script:Date + ".log"
            }
        })
}

function Get-MailboxDetails
{
    If ($Script:Mbx.RequireSenderAuthenticationEnabled -eq $True)
    {
        $Script:chkAllow.Checked = $False
        $Script:chkRestrict.Checked = $True
    }
    else
    {
        $Script:chkAllow.Checked = $True
        $Script:chkRestrict.Checked = $False
    }
    $Script:txtEDGrpName.Text = "Not Found"
    $Script:txtAUGrpName.Text = "Not Found"
    $Script:txtREGrpName.Text = "Not Found"
    $Script:txtGrpOwnr.Text = "None Found"
    $Script:MsgTo = $Script:Mbx.PrimarySMTPAddress

    $MbxPerm = get-mailboxPermission $Script:txtDispName.Text |Where-Object {$_.User -like "*MBX*"}
    foreach ($MbxPerm in $MbxPerm)
    {
        If ([bool](Get-DistributionGroup $MbxPerm.User -ErrorAction SilentlyContinue))
        {
            If ((Get-DistributionGroupMember $MbxPerm.User).count -gt 0)
            {
                If ($MbxPerm.User -like "*.ED")
                {
                    If ($Script:txtEDGrpName.Text -like "Not Found*")
                    {
                        $Script:txtEDGrpName.Text = $MbxPerm.User.DisplayName
                        $Script:txtEDGrpMbr.Text = (Get-DistributionGroupMember $MbxPerm.User).DisplayName -join "; "
                        $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User).ManagedBy -join "; "
                    }
                    else
                    {
                        $Script:txtEDGrpName.Text = $Script:txtEDGrpName.Text + "," + $MbxPerm.User.DisplayName
                        $Script:txtEDGrpMbr.Text = $Script:txtEDGrpMbr.Text + "," + (Get-DistributionGroupMember $MbxPerm.User).DisplayName -join "; "
                        If ($Script:txtGrpOwnr.Text -ne "")
                        {
                            $Script:txtGrpOwnr.Text = $Script:txtGrpOwnr.Text + "," + (Get-DistributionGroup $mbxPerm.User).ManagedBy -join "; "
                        }
                    }
                    $Script:MsgTo = $Script:MsgTo + "; " + ((Get-DistributionGroupMember $MbxPerm.User).DisplayName -join "; ")
                }

                If ($MbxPerm.User -like "*.AU")
                {
                    If ($Script:txtAUGrpName.Text -like "Not Found*")
                    {
                        $Script:txtAUGrpName.Text = $MbxPerm.User.DisplayName
                        $Script:txtAUGrpMbr.Text = (Get-DistributionGroupMember $MbxPerm.User).DisplayName -join "; "
                        $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User).ManagedBy -join "; "
                    }
                    else
                    {
                        $Script:txtAUGrpName.Text = $Script:txtAUGrpName.Text + "," + $MbxPerm.User.DisplayName
                        $Script:txtAUGrpMbr.Text = $Script:txtAUGrpMbr.Text + "," + (Get-DistributionGroupMember $MbxPerm.User).DisplayName -join "; "
                        If ($Script:txtGrpOwnr.Text -ne "")
                        {
                            $Script:txtGrpOwnr.Text = $Script:txtGrpOwnr.Text + "," + (Get-DistributionGroup $mbxPerm.User).ManagedBy -join "; "
                        }
                    }
                    $Script:MsgTo = $Script:MsgTo + "; " + ((Get-DistributionGroupMember $MbxPerm.User).DisplayName -join "; ")
                }

                If ($MbxPerm.User -like "*.RE")
                {
                    If ($Script:txtREGrpName.Text -like "Not Found*")
                    {
                        $Script:txtREGrpName.Text = $MbxPerm.User
                        $Script:txtREGrpMbr.Text = (Get-DistributionGroupMember $MbxPerm.User).DisplayName -join "; "
                        $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User).ManagedBy -join "; "
                    }
                    else
                    {
                        $Script:txtREDGrpName.Text = $Script:txtREGrpName.Text + "," + $MbxPerm.User.DisplayName
                        $Script:txtREGrpMbr.Text = $Script:txtREGrpMbr.Text + "," + (Get-DistributionGroupMember $MbxPerm.User).DisplayName -join "; "
                        If ($Script:txtGrpOwnr.Text -ne "")
                        {
                            $Script:txtGrpOwnr.Text =  $Script:txtGrpOwnr.Text + "," + (Get-DistributionGroup $mbxPerm.User).ManagedBy -join "; "
                        }
                    }
                    $Script:MsgTo = $Script:MsgTo + "; " + ((Get-DistributionGroupMember $MbxPerm.User).DisplayName -join "; ")
                }
            }
        }
    }

    $MbxPerm = get-mailboxfolderPermission $Script:txtDispName.Text |Where-Object {$_.AccessRights -notlike "*None*"}
    foreach ($MbxPerm in $MbxPerm)
    {
        If ([bool](get-distributiongroup $MbxPerm.User.DisplayName -ErrorAction SilentlyContinue))
        {
            If ((Get-DistributionGroupMember $MbxPerm.User.Displayname).Name.count -gt 0)
            {
                If (($MbxPerm.User.Displayname -like "*.ED") -or (($MbxPerm.User.DisplayName -notlike "*,*") -and ($MbxPerm.AccessRights -eq "Editor")))
                {
                    If ($Script:txtEDGrpName.Text -eq "")
                    {
                        $Script:txtEDGrpName.Text = $MbxPerm.User.DisplayName
                        $Script:txtEDGrpMbr.Text = ((Get-DistributionGroupMember $MbxPerm.User).DisplayName) -join "; "
                        $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join "; "
                    }
                    else
                    {
                        If ($Script:txtEDGrpName.Text -notlike ("*"+((get-distributiongroup $mbxperm.user.DisplayName).Name)+"*"))
                        {
                            $Script:txtEDGrpName.Text = $Script:txtEDGrpName.Text + "," + $MbxPerm.User.DisplayName
                            $Script:txtEDGrpMbr.Text = $Script:txtEDGrpMbr.Text + "," + (Get-DistributionGroupMember $MbxPerm.User.Displayname).DisplayName -join "; "
                            $Script:txtGrpOwnr.Text = $Script:txtGrpOwnr.Text + "," + (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join "; "
                        }
                    }
                    If ($Script:txtEDGrpMbr.Text -ne "")
                    {
                        $Script:MsgTo = $MsgTo + "; " + (Get-DistributionGroupMember $MbxPerm.User.Displayname).DisplayName -join "; "
                    }
                }
                If (($MbxPerm.User.Displayname -like "*.AU") -or (($MbxPerm.User.DisplayName -notlike "*,*") -and ($MbxPerm.AccessRights -eq "PublishingAuthor")))
                {
                    $Script:txtAUGrpName.Text = $MbxPerm.User.DisplayName
                    $Script:txtAUGrpMbr.Text = (Get-DistributionGroupMember $MbxPerm.User.Displayname).DisplayName -join "; "
                    If ($Script:txtGrpOwnr.Text -eq "")
                    {
                        $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join "; "
                    }
                    else
                    {
                        If ($Script:txtAUGrpName.Text -notlike ("*"+((get-distributiongroup $mbxperm.user.DisplayName).Name)+"*"))
                        {
                            $Script:txtAUGrpName.Text = $Script:txtAUGrpName.Text + "," + $MbxPerm.User.DisplayName
                            $Script:txtAUGrpMbr.Text = $Script:txtAUGrpMbr.Text + "," + (Get-DistributionGroupMember $MbxPerm.User).Displayname -join "; "
                            If ($Script:txtGrpOwnr.Text -eq "")
                            {
                                $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join "; "
                            }
                        }
                    }
                    If ($Script:txtAUGrpMbr.Text -ne "")
                    {
                        $Script:MsgTo = $MsgTo + "; " + (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join "; "
                    }
                }
                If (($MbxPerm.User.Displayname -like "*.RE") -or (($MbxPerm.User.DisplayName -notlike "*,*") -and ($MbxPerm.AccessRights -eq "Reviewer")))
                {
                    $Script:txtREGrpName.Text = $MbxPerm.User.DisplayName
                    $Script:txtREGrpMbr.Text = (Get-DistributionGroupMember $MbxPerm.User.Displayname).DisplayName -join "; "
                    If ($Script:txtGrpOwnr.Text -eq "")
                    {
                        $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join "; "
                    }
                    else
                    {
                        If ($Script:txtREGrpName.Text -notlike ("*"+((get-distributiongroup $mbxperm.user.DisplayName).Name)+"*"))
                        {
                            $Script:txtREGrpName.Text = $Script:txtREGrpName.Text + "," + $MbxPerm.User.DisplayName
                            $Script:txtREGrpMbr.Text = $Script:txtREGrpMbr.Text + "," + (Get-DistributionGroupMember $MbxPerm.User.Displayname).DisplayName -join "; "
                            If ($Script:txtGrpOwnr.Text -eq "")
                            {
                                $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User).ManagedBy -join "; "
                            }
                        }
                    }
                    If ($Script:txtREGrpMbr.Text -ne "")
                    {
                        $Script:MsgTo = $Script:MsgTo + "; " + (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join "; "
                    }
                }
            }
        }
    }

#    If no group is found see if one exists following our naming standard"
     If ($Script:txtEDGrpName.Text -eq "Not Found")
     {
        $Exists = [bool](Get-DistributionGroup ("MBX." + $Script:txtDispName.Text + ".ED") -ErrorAction SilentlyContinue)
        If ($Exists -eq $True)
        {
            $Script:txtEDGrpName.Text = ("MBX." + $Script:txtDispName.Text + ".ED") + " (Not Configured)"
            $Script:txtEDGrpMbr.Text = (Get-DistributionGroupMember ("MBX." + $Script:txtDispName.Text + ".ED")).DisplayName -join "; "
        }
     }

     If ($Script:txtAUGrpName.Text -eq "Not Found")
     {
        $Exists = [bool](Get-DistributionGroup ("MBX." + $Script:txtDispName.Text + ".AU") -ErrorAction SilentlyContinue)
        If ($Exists -eq $True)
        {
            $Script:txtAUGrpName.Text = ("MBX." + $Script:txtDispName.Text + ".AU") + " (Not Configured)"
            $Script:txtAUGrpMbr.Text = (Get-DistributionGroupMember ("MBX." + $Script:txtDispName.Text + ".AU")).DisplayName -join "; "
        }
     }

     If ($Script:txtREGrpName.Text -eq "Not Found")
     {
        $Exists = [bool](Get-DistributionGroup ("MBX." + $Script:txtDispName.Text + ".RE") -ErrorAction SilentlyContinue)
        If ($Exists -eq $True)
        {
            $Script:txtREGrpName.Text = ("MBX." + $Script:txtDispName.Text + ".RE") + " (Not Configured)"
            $Script:txtREGrpMbr.Text = (Get-DistributionGroupMember ("MBX." + $Script:txtDispName.Text + ".RE")).DisplayName -join "; "
        }
     }

    If ($Script:txtGrpOwnr.Text -eq "")
    {
        $Script:txtGrpOwnr.Text = "None Found"
    }

    $CurrTip = $Script:Mbx.Mailtip -Replace("html","")
    $CurrTip = $CurrTip -Replace("body","")
    $Script:txtCurrTip.Text = $CurrTip -Replace('[<>/]','')

    #Remove Duplicate addresses
    $UniqueAddr = $Script:MsgTo -split ("; ")
    $Script:MsgTo = ($UniqueAddr |select -Unique) -join "; "
}

Function Write-ShrMbxDetails
{
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "               Mailbox Name: " + $Script:txtDispName.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "              Ticket Number: " + $Script:txtTicketNo.Text
    WriteReportEvent
    If ($Script:txtTicketNo.Text -eq "Email Request")
    {
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "               Message Date: " + $Script:txEmailDate.Text
        WriteReportEvent
    }
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "             Mailbox Owners: " + $Script:txtGrpOwnr.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "      Current Mailbox Usage: " + ($Script:Mbx.RequireSenderAuthenticationEnabled) + " = O365 Internal Messages Only"
    If ($Script:Mbx.RequireSenderAuthenticationEnabled -eq $False)
    {
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "      Current Mailbox Usage: " + ($Script:Mbx.RequireSenderAuthenticationEnabled) + " = External Addresses Allowed"
    }
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "               Editor Group: " + $Script:txtEDGrpName.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "       Editor Group Members: " + $Script:txtEDGrpMbr.Text
    WriteReportEvent
    If ($Script:txtEDGrpMbr.Text.Length -gt 0)
    {
        write-host "Getting ED Group members"
        $LineToWrite = $RecordEvent + "INFO" + "`t" + " Editor Group Members EmpNo: " + ((get-DistributionGroupMember $Script:txtEDGrpName.Text).Alias -join "; ")
        WriteReportEvent
    }
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "               Author Group: " + $Script:txtAUGrpName.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "       Author Group Members: " + $Script:txtAUGrpMbr.Text
    WriteReportEvent
    If ($Script:txtAUGrpMbr.Text.Length -gt 0)
    {
        $LineToWrite = $RecordEvent + "INFO" + "`t" + " Author Group Members EmpNo: " + ((get-DistributionGroupMember $Script:txtAUGrpName.Text).Alias -join "; ")
        WriteReportEvent
    }
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "               Reader Group: " + $Script:txtREGrpName.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "       Reader Group Members: " + $Script:txtREGrpMbr.Text
    WriteReportEvent
    If ($Script:txtREGrpMbr.Text.Length -gt 0)
    {
        $LineToWrite = $RecordEvent + "INFO" + "`t" + " Reader Group Members EmpNo: " + ((get-DistributionGroupMember $Script:txtREGrpName.Text).Alias -join "; ")
        WriteReportEvent
    }
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`n"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "              Mailbox Usage: External Email Restricted"
    If ($Script:chkAllow.Checked -eq $True)
    {
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "              Mailbox Usage: External Email Allowed"
    }
    WriteReportEvent
}

Build-ShrMbxUseDetailsForm
Hide-Details
Publish-Form

If ($Global:Result -eq "OK") 
{

#  Recheck to see of the Cancel button was not hit when getting 
    $ReportFile = $Script:txtRptFile.Text
    $LineToWrite = $RecordEvent + "STAR" + "`t" + "Change Shared Maibox Use Script"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI
    WriteReportEvent
 	$LineToWrite = $RecordEvent + "INFO" + "`t" + "Script Last Change Date: 04/27/2023" + "`n"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
    WriteReportEvent
    Write-ShrMbxDetails

    If ($Script:chkAllow.Checked -eq $True)
    {
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "       No change made to mailbox usage restiction"
        If ($Script:Mbx.RequireSenderAuthenticationEnabled -eq $True)
        {
            write-host "Change Mailbox to allow external use"
            pause
            set-mailbox $Script:txtDispName.Text -RequireSenderAuthenticationEnabled $False
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "       Changed Restriction to Allow Recipt of External Messages"
        }
    }

    If ($Script.chkRestrict.Checked -eq $True)
    {
        If ($Script:Mbx.RequireSenderAuthenticationEnabled -eq $False)
        {
            write-host "Change Mailbox to prevent external use"
            pause
            set-mailbox $Script:txtDispName.Text -RequireSenderAuthenticationEnabled $True
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "       Changed Restriction to Deny Recipt of External Messages"
        }
    }
    WriteReportEvent

    If ($Script:chkTemplate.Checked -eq $True)
    {
        invoke-Expression -Command e:\O365AdminShared\EMailTemplates\SharedMailboxUse.oft
    }

    write-host "Shared Mailbox Use Request Complete" -ForegroundColor Cyan
    $LineToWrite = $RecordEvent + "DONE" + "`t" + "Shared Mailbox Use Request Complete"
    WriteReportEvent
}
else
{
    write-host "Shared Mailbox Removal Cancelled" -ForegroundColor Red
    $Output = $wshell.Popup("Shared mailbox removal request cancelled.",0,"Cancelled",0+32)
}
