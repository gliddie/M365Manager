<#   
#================================================================================ 
# Name: Remove Shared Mailbox
#
# 10/12/2020 - SAG - Redesigned using forms and the old RemoveSharedMailbox.ps1 script
# 05/17/2022 - SAG - Modified the report file to be written out to the E: drive
# 11/10/2022 - SAG - Modified to use a GUI interface
# 04/07/2023 - SAG - Added group member empno details into log information
# 04/04/2025 - SAG - Added Forwarding Address to the details gathered and reported
#================================================================================ 
#>  

function Hide-Details
{
    $Global:form.Width = 600 ; $Global:form.Height = 150
    $Script:txtDispName.ReadOnly = $False
    $Script:ButGetDL.Visible = $True
    $Script:lblTicketNo.Visible = $False
    $Script:txtTicketNo.Visible = $False
    $Script:lblFwdAddr.Visible = $False
    $Script:txtFwdAddr.Visible = $False
    $Script:lblGrpOwnr.Visible = $False
    $Script:txtGrpOwnr.Visible = $False
    $Script:lblCurrTip.Visible = $False
    $Script:txtCurrTip.Visible = $False
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
    $Script:lblNonStdAcc.Visible = $False
    $Script:txtNonStdAcc.Visible = $False
    $Script:lblRptFile.Visible = $False
    $Script:txtRptFile.Visible = $False
    $Script:chkTemplate.Visible = $False
    $Script:chkRemEDGrp.Visible = $False
    $Script:chkRemAUGrp.Visible = $False
    $Script:chkRemREGrp.Visible = $False
    $Script:chkReqOwner.Visible = $False
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
    $Global:form.Width = 700 ; $Global:form.Height = 610
    If ($Script:chkReqOwner.Visible -eq $True)
    {
        $Global:form.Width = 700 ; $Global:form.Height = 640
    }
    $Script:txtDispName.ReadOnly = $True
    $Script:ButGetDL.Visible = $False
    $Script:lblTicketNo.Visible = $True
    $Script:txtTicketNo.Visible = $True
    $Script:lblFwdAddr.Visible = $True
    $Script:txtFwdAddr.Visible = $True
    $Script:lblGrpOwnr.Visible = $True
    $Script:txtGrpOwnr.Visible = $True
    $Script:lblCurrTip.Visible = $True
    $Script:txtCurrTip.Visible = $True
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
    $Script:lblNonStdAcc.Visible = $True
    $Script:txtNonStdAcc.Visible = $True
    $Script:lblRptFile.Visible = $True
    $Script:txtRptFile.Visible = $True
    $Script:chkTemplate.Visible = $True
    If ($Script:txtEDGrpName.Text -ne "Not Found")
    {
        $Script:chkRemEDGrp.Visible = $True
    }
    If ($Script:txtAUGrpName.Text -ne "Not Found")
    {
        $Script:chkRemAUGrp.Visible = $True
    }
    If ($Script:txtREGrpName.Text -ne "Not Found")
    {
        $Script:chkRemREGrp.Visible = $True
    }
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

function Build-ShrMbxRemDetailsForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Remove Shared Mailbox" 
    $Global:form.StartPosition = "CenterScreen"

    Add-FormStandardButtons
    $Global:OKButton.AutoSize = $True
    $Global:OKButton.Add_Click({
        If ($Script:txtTicketNo.Text -eq "Email Request")
        {
            Get-EmailDate
        }
    })
    ## Get Details to complete creation
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
#                Check-Reconnect
                Get-MailboxDetails
                UnHide-Details
                $Script:Date = get-date -Format "yyyy-MMdd"
                $Script:txtRptFile.Text = $Path + "\RemoveShrMbx-" + ($Script:txtDispName.Text -Replace("['.()/\- ]","")) + "-Date" + $Script:Date + ".log"
                $Global:OKButton.Text = "Remove"
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
    $Script:txtTicketNo = New-Object Windows.Forms.ComboBox
        $Script:txtTicketNo.TabIndex = 0 # set Tab Order 
        $Script:txtTicketNo.Top = $TopLoc; $Script:txtTicketNo.Left = 130; $Script:txtTicketNo.Width = 180;
        If ($Script:chkReqOwner.Checked -eq $True)
        {
            $Script:txtTicketNo.Text = "No Owners/No Members Request"
        }
        [void] $Script:txtTicketNo.Items.Add("CHG0103470 (3/29 email from Circle Fong)") #CHG ticket to remove 358 mailboxes based on email from Circle Fong (started on 3/29)
        [void] $Script:txtTicketNo.Items.Add("Email Request")  # Add element to listbox  
        [void] $Script:txtTicketNo.Items.Add("No Owners/No Members Request")  # Add element to listbox
        [void] $Script:txtTicketNo.Items.Add("No Remaining Owners Or Members")  # Add element to listbox
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
    ## Forwarding Addresses
    $Script:lblFwdAddr = New-Object System.Windows.Forms.Label   
        $Script:lblFwdAddr.Text = "Forwarding Address:"  
        $Script:lblFwdAddr.Top = $TopLoc ; $Script:lblFwdAddr.Left = 10; $Script:lblFwdAddr.Width=120
        $Script:lblFwdAddr.Visible = $False
        $Global:form.Controls.Add($Script:lblFwdAddr)    # Add to Form              
    $Script:txtFwdAddr = New-Object Windows.Forms.TextBox
        $Script:txtFwdAddr.TabIndex = 0 # set Tab Order 
        $Script:txtFwdAddr.Top = $TopLoc; $Script:txtFwdAddr.Left = 130; $Script:txtFwdAddr.Width = 500;
        $Script:txtFwdAddr.ReadOnly = $True
        $Global:form.Controls.Add($Script:txtFwdAddr)    # Add to Form

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

    ## DeleteEDGroupCheck Box
    $Script:chkRemEDGrp = New-Object Windows.Forms.checkbox 
        $Script:chkRemEDGrp.Left = 360; $Script:chkRemEDGrp.Width = 150; $Script:chkRemEDGrp.Top = $TopLoc
        $Script:chkRemEDGrp.Text = "Delete ED Group" 
        $Script:chkRemEDGrp.Checked = $false   # set a default value 
        $Script:chkRemEDGrp.TabIndex = 5
        $Script:chkRemEDGrp.Checked = $True
        $Global:form.Controls.Add($Script:chkRemEDGrp) 
        # Obtain Value with: $Script:chkRemEDGrp.Checked

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

    ## DeleteAUGroupCheck Box
    $Script:chkRemAUGrp = New-Object Windows.Forms.checkbox 
        $Script:chkRemAUGrp.Left = 360; $Script:chkRemAUGrp.Width = 150; $Script:chkRemAUGrp.Top = $TopLoc
        $Script:chkRemAUGrp.Text = "Delete AU Group" 
        $Script:chkRemAUGrp.Checked = $false   # set a default value 
        $Script:chkRemAUGrp.TabIndex = 5
        $Script:chkRemAUGrp.Checked = $True
        $Global:form.Controls.Add($Script:chkRemAUGrp) 

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

    ## DeleteREGroupCheck Box
    $Script:chkRemREGrp = New-Object Windows.Forms.checkbox 
        $Script:chkRemREGrp.Left = 360; $Script:chkRemREGrp.Width = 150; $Script:chkRemREGrp.Top = $TopLoc
        $Script:chkRemREGrp.Text = "Delete RE Group" 
        $Script:chkRemREGrp.Checked = $false   # set a default value 
        $Script:chkRemREGrp.TabIndex = 5
        $Script:chkRemREGrp.Checked = $True
        $Global:form.Controls.Add($Script:chkRemREGrp) 

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
    ##Non-Standard Access
    $Script:lblNonStdAcc = New-Object System.Windows.Forms.Label   
        $Script:lblNonStdAcc.Text = "NonStandard Access:"  
        $Script:lblNonStdAcc.Top = $TopLoc; $Script:lblNonStdAcc.Left = 10; $Script:lblNonStdAcc.Width=120;
        $Global:form.Controls.Add($Script:lblNonStdAcc)    # Add to Form 
    $Script:txtNonStdAcc = New-Object System.Windows.Forms.ListBox
        $Script:txtNonStdAcc.Top = $TopLoc; $Script:txtNonStdAcc.Left = 130; $Script:txtNonStdAcc.Width = 520; $Script:txtNonStdAcc.Height = 60
        $Script:txtNonStdAcc.BackColor = "#EEEDF0"
        $Global:form.Controls.Add($Script:txtNonStdAcc)    # Add to Form

    $TopLoc = $TopLoc + 60
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
        $Script:chkTemplate.Text = "Display Email Template for Shared Mailbox Removal" 
        $Script:chkTemplate.Checked = $true   # set a default value
        $Global:form.Controls.Add($Script:chkTemplate)
        $Script:chkTemplate.Add_Click({
            If ($Script:chkTemplate.Checked -eq $True)
            {
                $Script:chkReqOwner.Checked = $False
                $Global:OKButton.Text = "Remove"
                $Script:txtRptFile.Text = $Path + "\RemoveShrMbx-" + ($Script:txtDispName.Text -Replace("['.()/\- ]","")) + "-Date" + $Script:Date + ".log"
            }
        })

    $TopLoc = $TopLoc + 30
    $Global:form.Height = $Global:form.Height + 30
    ## Send Message requesting Owner
    $Script:chkReqOwner = New-Object Windows.Forms.checkbox 
        $Script:chkReqOwner.Left = 130; $Script:chkReqOwner.Width = 400; $Script:chkReqOwner.Top = $TopLoc
        $Script:chkReqOwner.Text = "Display Email Template to Request Shared Mailbox Owners"
        $Script:chkReqOwner.Checked = $False   # set a default value
        $Global:form.Controls.Add($Script:chkReqOwner)
        $Script:chkReqOwner.Add_Click({
            If ($Script:chkReqOwner.Checked -eq $True)
            {
                $Script:chkTemplate.Checked = $False
                $Global:OKButton.Text = "Send Message"
                $Script:txtRptFile.Text = "e:\Automation\RequestForOwner\Report\RequestForOwner-" + ($Script:txtDispName.Text -Replace("['.()/\- ]","")) + "-Date" + $Script:Date + ".log"
            }
            else
            {
                $Global:OKButton.Text = "Remove"
                $Script:txtRptFile.Text = $Path + "\RemoveShrMbx-" + ($Script:txtDispName.Text -Replace("['.()/\- ]","")) + "-Date" + $Script:Date + ".log"
            }
        })
}

function Get-MailboxDetails
{
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
           #Owner Friendly Names
            If ($Script:txtGrpOwnr.Text -eq "None Found")
            {
                $GrpMgrs = (Get-DistributionGroup $mbxPerm.User).ManagedBy -split(",")
                MemberFriendlyName
                $Script:txtGrpOwnr.Text = $Script:FriendlyName
            }

            If ((Get-DistributionGroupMember $MbxPerm.User).count -gt 0)
            {
                If ($MbxPerm.User -like "*.ED")
                {
                    If ($Script:txtEDGrpName.Text -like "Not Found*")
                    {
                        $GrpMgrs = (Get-DistributionGroupMember $mbxPerm.User) -split(",")
                        MemberFriendlyName
                        $Script:txtEDGrpMbr.Text = $Script:FriendlyName
                        $Script:txtEDGrpName.Text = $MbxPerm.User
                    }
                    $Script:MsgTo = $Script:MsgTo + "; " + ((Get-DistributionGroupMember $MbxPerm.User) -join "; ")
                }

                If ($MbxPerm.User -like "*.AU")
                {
                    If ($Script:txtAUGrpName.Text -like "Not Found*")
                    {
                        $GrpMgrs = (Get-DistributionGroupMember $mbxPerm.User) -split(",")
                        MemberFriendlyName
                        $Script:txtAUGrpMbr.Text = $Script:FriendlyName
                        $Script:txtAUGrpName.Text = $MbxPerm.User
                    }
                    $Script:MsgTo = $Script:MsgTo + "; " + ((Get-DistributionGroupMember $MbxPerm.User) -join "; ")
                }

                If ($MbxPerm.User -like "*.RE")
                {
                    If ($Script:txtREGrpName.Text -like "Not Found*")
                    {
                        $GrpMgrs = (Get-DistributionGroupMember $mbxPerm.User) -split(",")
                        MemberFriendlyName
                        $Script:txtREGrpMbr.Text = $Script:FriendlyName
                        $Script:txtREGrpName.Text = $MbxPerm.User
                    }
                    $Script:MsgTo = $Script:MsgTo + "; " + ((Get-DistributionGroupMember $MbxPerm.User) -join "; ")
                }
            }
        }
    }

    $MbxPerm = get-mailboxfolderPermission $Script:txtDispName.Text |Where-Object {$_.AccessRights -notlike "*None*"}
    foreach ($MbxPerm in $MbxPerm)
    {
        If ([bool](get-distributiongroup $MbxPerm.User.DisplayName -ErrorAction SilentlyContinue))
        {
            If ((Get-DistributionGroupMember $MbxPerm.User.Displayname).count -gt 0)
            {
                If (($MbxPerm.User.Displayname -like "*.ED") -or (($MbxPerm.User.DisplayName -notlike "*,*") -and ($MbxPerm.AccessRights -eq "Editor")))
                {
                    If ($Script:txtEDGrpName.Text -eq "")
                    {
                        $Script:txtEDGrpName.Text = $MbxPerm.User.DisplayName
                        $Script:txtEDGrpMbr.Text = (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ", "
                        $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
                    }
                    else
                    {
                        If ($Script:txtEDGrpName.Text -notlike ("*"+((get-distributiongroup $mbxperm.user.DisplayName).Name)+"*"))
                        {
                            $Script:txtEDGrpName.Text = $Script:txtEDGrpName.Text + "," + $MbxPerm.User.DisplayName
                            $Script:txtEDGrpMbr.Text = $Script:txtEDGrpMbr.Text + "," + (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ", "
                            $Script:txtGrpOwnr.Text = $Script:txtGrpOwnr.Text + "," + (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
                        }
                    }
                    If ($Script:txtEDGrpMbr.Text -ne "")
                    {
                        $Script:chkReqOwner.Visible = $True
                        $Script:MsgTo = $MsgTo + "; " + ((Get-DistributionGroupMember $MbxPerm.User.Displayname) -join "; ")
                    }
                }
                If (($MbxPerm.User.Displayname -like "*.AU") -or (($MbxPerm.User.DisplayName -notlike "*,*") -and ($MbxPerm.AccessRights -eq "PublishingAuthor")))
                {
                    $Script:txtAUGrpName.Text = $MbxPerm.User.DisplayName
                    $Script:txtAUGrpMbr.Text = (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ", "
                    If ($Script:txtGrpOwnr.Text -eq "")
                    {
                        $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
                    }
                    else
                    {
                        If ($Script:txtAUGrpName.Text -notlike ("*"+((get-distributiongroup $mbxperm.user.DisplayName).Name)+"*"))
                        {
                            $Script:txtAUGrpName.Text = $Script:txtAUGrpName.Text + "," + $MbxPerm.User.DisplayName
                            $Script:txtAUGrpMbr.Text = $Script:txtAUGrpMbr.Text + "," + (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ", "
                            If ($Script:txtGrpOwnr.Text -eq "")
                            {
                                $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
                            }
                        }
                    }
                    If ($Script:txtAUGrpMbr.Text -ne "")
                    {
                        $Script:chkReqOwner.Visible = $True
                        $Script:MsgTo = $MsgTo + "; " + ((Get-DistributionGroupMember $MbxPerm.User.Displayname) -join "; ")
                    }
                }
                If (($MbxPerm.User.Displayname -like "*.RE") -or (($MbxPerm.User.DisplayName -notlike "*,*") -and ($MbxPerm.AccessRights -eq "Reviewer")))
                {
                    $Script:txtREGrpName.Text = $MbxPerm.User.DisplayName
                    $Script:txtREGrpMbr.Text = (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ", "
                    If ($Script:txtGrpOwnr.Text -eq "")
                    {
                        $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
                    }
                    else
                    {
                        If ($Script:txtREGrpName.Text -notlike ("*"+((get-distributiongroup $mbxperm.user.DisplayName).Name)+"*"))
                        {
                            $Script:txtREGrpName.Text = $Script:txtREGrpName.Text + "," + $MbxPerm.User.DisplayName
                            $Script:txtREGrpMbr.Text = $Script:txtREGrpMbr.Text + "," + (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ", "
                            If ($Script:txtGrpOwnr.Text -eq "")
                            {
                                $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
                            }
                        }
                    }
                    If ($Script:txtREGrpMbr.Text -ne "")
                    {
                        $Script:chkReqOwner.Visible = $True
                        $Script:MsgTo = $Script:MsgTo + "; " + ((Get-DistributionGroupMember $MbxPerm.User.Displayname) -join "; ")
                    }
                }
            }
        }
        If ($MbxPerm.User.Displayname -notlike "MBX*")
        {
            $UsrMbxExists = [bool]($UsrMbx = get-mailbox $MbxPerm.User.DisplayName -ErrorAction SilentlyContinue)
            If ($UsrMbxExists -eq $True)
            {
                $Details = $MbxPerm.User.Displayname + " - (" + $MbxPerm.AccessRights +") - Status: " + $UsrMbx.CustomAttribute1
                If (($UsrMbx.CustomAttribute2 -ne "T") -and ($Script:txtGrpOwnr.Text.Length -eq 0))
                {
                    $Script:chkReqOwner.Visible = $True
                    $Script:MsgTo = $Script:MsgTo + "; " + $MbxPerm.User.Displayname
                }
            }
            else
            {
                $Details = $MbxPerm.User.Displayname + " - (" + $MbxPerm.AccessRights +") - Status: Account Deleted"
            }
            [void] $Script:txtNonStdAcc.Items.Add($Details)  # Add element to listbox 
        }
    }

#    If no group is found see if one exists following our naming standard"
    If ($Script:txtEDGrpName.Text -eq "Not Found")
    {
        $Exists = [bool](Get-DistributionGroup ("MBX." + $Script:txtDispName.Text + ".ED") -ErrorAction SilentlyContinue)
        If ($Exists -eq $True)
        {
            $Script:txtEDGrpName.Text = ("MBX." + $Script:txtDispName.Text + ".ED") + " (Not Configured)"
            $Script:txtEDGrpMbr.Text = (Get-DistributionGroupMember ("MBX." + $Script:txtDispName.Text + ".ED")) -join ", "
        }
    }

    If ($Script:txtAUGrpName.Text -eq "Not Found")
    {
       $Exists = [bool](Get-DistributionGroup ("MBX." + $Script:txtDispName.Text + ".AU") -ErrorAction SilentlyContinue)
       If ($Exists -eq $True)
       {
           $Script:txtAUGrpName.Text = ("MBX." + $Script:txtDispName.Text + ".AU") + " (Not Configured)"
           $Script:txtAUGrpMbr.Text = (Get-DistributionGroupMember ("MBX." + $Script:txtDispName.Text + ".AU")) -join ", "
       }
    }

    If ($Script:txtREGrpName.Text -eq "Not Found")
    {
       $Exists = [bool](Get-DistributionGroup ("MBX." + $Script:txtDispName.Text + ".RE") -ErrorAction SilentlyContinue)
       If ($Exists -eq $True)
       {
           $Script:txtREGrpName.Text = ("MBX." + $Script:txtDispName.Text + ".RE") + " (Not Configured)"
           $Script:txtREGrpMbr.Text = (Get-DistributionGroupMember ("MBX." + $Script:txtDispName.Text + ".RE")) -join ", "
       }
    }

    If ($Script:txtGrpOwnr.Text -eq "")
    {
        $Script:txtGrpOwnr.Text = "None Found"
    }
    else
    {
        $Script:chkReqOwner.Visible = $True
    }

    If (($Script:Mbx.ForwardingAddress.length -gt 0) -or ($Script:Mbx.ForwardingSMTPAddress.Length -gt 0))
    {
        $Script:txtFwdAddr.Text = $Script:Mbx.ForwardingSMTPAddress
        If ($Script:Mbx.ForwardingAddress.length -gt 0)
        {
            $Script:txtFwdAddr.Text = $Script:Mbx.ForwardingAddress
        }
    }
    
    $CurrTip = $Script:Mbx.Mailtip -Replace("html","")
    $CurrTip = $CurrTip -Replace("body","")
    $Script:txtCurrTip.Text = $CurrTip -Replace('[<>/]','')

    #Remove Duplicate addresses
    $UniqueAddr = $Script:MsgTo -split ("; ")
    $Script:MsgTo = ($UniqueAddr |select -Unique) -join "; "
}

function Remove-Group
{
    If ($InputDL -notlike "*,*")
    {
        If ($InputDL -like "*(Not*")
        {
            $InputDL = $InputDL.Substring(0,$InputDL.IndexOf("(")-1)
        }
        else
        {
           $InputDL = $InputDL.Trim()
        }
        $Exists = [bool](Get-DistributionGroup $InputDL -ErrorAction SilentlyContinue)
        If ($Exists -eq $True)
        {
            $InputDL = $InputDL.Trim()
            write-host "Removing Access Group: " $InputDL
            Write-Output "Security Access Group Details - " $InputDL >> $OutFileName
            Get-DistributionGroup $InputDL |ft Name,DisplayName,GroupType,PrimarySmtpAddress >> $OutFileName
	        Write-Output "Security Access Group Full Details" >> $OutFileName
            Get-DistributionGroup $InputDL |fl >> $OutFileName
	        Write-Output "Security Access Group Manager and Notes" >> $OutFileName
            Get-Group $InputDL |fl ManagedBy,Notes >> $OutFileName
	        write-Output "Security Access Group Membership" >> $OutFileName
            Get-DistributionGroupMember $InputDL |ft Alias,Name,RecipientType >> $OutFileName
            Remove-DistributionGroup $InputDL -BypassSecurityGroupManagerCheck -confirm:$False >> $OutFileName
        }
    }
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
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "          Mailbox Forwarder: " + $Script:txtFwdAddr.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "        Current Mailbox Tip: " + $Script:txtCurrTip.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "             Mailbox Owners: " + $Script:txtGrpOwnr.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "               Editor Group: " + $Script:txtEDGrpName.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "       Editor Group Members: " + $Script:txtEDGrpMbr.Text
    WriteReportEvent
    If ($Script:txtEDGrpMbr.Text.Length -gt 0)
    {
        $Grp = $Script:txtEDGrpName.Text
        If ($Grp -like "*(Not*")
        {
            $Grp = $Grp.Substring(0,$Grp.IndexOf("(")-1)
        }
        $LineToWrite = $RecordEvent + "INFO" + "`t" + " Editor Group Members EmpNo: " + ((get-DistributionGroupMember $Grp).Alias -join ",")
        WriteReportEvent
    }
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "               Author Group: " + $Script:txtAUGrpName.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "       Author Group Members: " + $Script:txtAUGrpMbr.Text
    WriteReportEvent
    If ($Script:txtAUGrpMbr.Text.Length -gt 0)
    {
        $Grp = $Script:txtAUGrpName.Text
        If ($Grp -like "*(Not*")
        {
            $Grp = $Grp.Substring(0,$Grp.IndexOf("(")-1)
        }
        $LineToWrite = $RecordEvent + "INFO" + "`t" + " Author Group Members EmpNo: " + ((get-DistributionGroupMember $Grp).Alias -join ",")
        WriteReportEvent
    }
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "               Reader Group: " + $Script:txtREGrpName.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "       Reader Group Members: " + $Script:txtREGrpMbr.Text
    WriteReportEvent
    If ($Script:txtREGrpMbr.Text.Length -gt 0)
    {
        $Grp = $Script:txtREGrpName.Text
        If ($Grp -like "*(Not*")
        {
            $Grp = $Grp.Substring(0,$Grp.IndexOf("(")-1)
        }
        $LineToWrite = $RecordEvent + "INFO" + "`t" + " Reader Group Members EmpNo: " + ((get-DistributionGroupMember $Grp).Alias -join ",")
        WriteReportEvent
    }
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "        Non-Standard Access: " + ((get-mailboxfolderPermission $Script:txtDispName.Text |Where-Object {($_.AccessRights -notlike "*None*") -and ($_.User -notlike "MBX*")}).User.DisplayName -join "`n`t`t`t`t`t`t`t`t`t") + "`n"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
    WriteReportEvent
}

$Year = (get-date).ToString("yyyy")
$Path = "e:\Automation\RemoveShrMbx\Report\" + $Year
If (Test-Path $path) {} else {New-Item -Path $Path -ItemType Directory}

Build-ShrMbxRemDetailsForm
Hide-Details
Publish-Form

If ($Global:Result -eq "OK") 
{
#  Recheck to see of the Cancel button was not hit when getting 
    $ReportFile = $Script:txtRptFile.Text
    $LineToWrite = $RecordEvent + "STAR" + "`t" + "Remove Shard Maibox Script"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI
    WriteReportEvent
 	$LineToWrite = $RecordEvent + "INFO" + "`t" + "Script Last Change Date: 11/11/2022" + "`n"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
    WriteReportEvent
    Write-ShrMbxDetails

    If ($Script:chkReqOwner.Checked -eq $False)
    {
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Starting Shared Mailbox Removal: " + $Script:txtDispName.Text
        WriteReportEvent
        $Reason = $Script:txtTicketNo.Text
        $OutFileName = $Script:txtRptFile.Text
        write-host $OutFileName
        
        write-host "Removing Shared Mailbox: " $Script:txtDispName.Text -ForegroundColor Cyan
        Write-Output "Removing Shared Mailbox" >> $OutFileName
        write-Output "Launched by:  $WhoAmI" >> $OutFileName
        Write-Output "Ticket Number:  $Reason" >> $OutFileName
        If ($Reason -eq "EmailResponse")
        {
            Write-Output "Email Request Date:  $Script:txtEmailComm.Text" >> $OutFileName
        }
        Get-Mailbox $Script:txtDispName.Text |ft Name,Alias,Database,ProhibitSendQuota,ExternalDirectoryObjectID >> $OutFileName
        Get-Mailbox $Script:txtDispName.Text |fl >> $OutFileName
        Write-Output "Mailbox Rules" >> $OutFileName
        Get-InboxRule -Mailbox $Script:txtDispName.Text |Where-Object{($_.Description -like "*forward*") -or ($_.Description -like "*redirect*")} |fl Name,Priority,*desc* >> $OutFileName
        Write-Output "Mailbox Permissions" >> $OutFileName
        get-mailboxPermission $Script:txtDispName.Text | where {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous") -and ($_.IsInherited -notlike "True")} | ft User,AccessRights >> $OutFileName
        Write-Output "MailboxFolder Permissions" >> $OutFileName
        get-mailboxfolderPermission $Script:txtDispName.Text |ft FolderName,User,AccessRights,SharingPermissionFlags >> $OutFileName
        Write-Output "Recipient Permissions" >> $OutFileName
        Get-RecipientPermission $Script:txtDispName.Text |ft Identity,Trustee,AccessControlType,AccessRights,Inherited >> $OutFileName
        Write-Output "Mailbox Message Statistics - Number of Messages in Mailbox" >> $OutFileName
        Get-MailboxStatistics $Script:txtDispName.Text |ft DisplayName,ItemCount,StorageLimitSatatus,LastLogonTime >> $OutFileName
        Write-Output "Mailbox Message Folder Statistics - Number of Messages in Mailbox" >> $OutFileName
        Get-MailboxFolderStatistics $Script:txtDispName.Text |ft Name,ItemsInFolder >> $OutFileName

        If ($Script:chkRemEDGrp.Checked -eq "Checked")
        {
            $InputDL = $Script:txtEDGrpName.Text -split ","
            Foreach ($InputDL in $InputDL)
            {
                Remove-Group
            }
        }

        If ($Script:chkRemAUGrp.Checked -eq "Checked")
        {

            $InputDL = $Script:txtAUGrpName.Text -split ","
            Foreach ($InputDL in $InputDL)
            {
                Remove-Group
            }
        }

        If ($Script:chkRemREGrp.Checked -eq "Checked")
        {
            $InputDL = $Script:txtREGrpName.Text -split ","
            Foreach ($InputDL in $InputDL)
            {
                Remove-Group
            }
        }

        Remove-Mailbox $Script:txtDispName.Text -confirm:$False >> $OutFileName
        
        If ($Script:chkTemplate.Checked -eq $True)
        {
            invoke-Expression -Command e:\O365AdminShared\EMailTemplates\SharedMailboxRemoval.oft
        }
        Write-Output "Shared Mailbox Removal Complete" >> $OutFileName
        write-host "Shared Mailbox Removal Complete" -ForegroundColor Cyan
        $LineToWrite = $RecordEvent + "DONE" + "`t" + "Shared Mailbox Removal Complete"
    }
    else
    {
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Sending Request for Shared Mailbox Owners: " + $Script:txtDispName.Text
        WriteReportEvent
        If ($Script:txtCurrTip.Text -like "*Message sent to obtain*")
        {
            write-host "Message already sent to obtain New owners"
            $LineToWrite = $RecordEvent + "NOTS" + "`t" + "Message Already Sent to Obtain Nw Owners"
            WriteReportEvent
        }
        else
        {
#       Send Message to get new owners and set the Mailtips on the mailbox and groups that this action has been taken
            invoke-Expression -Command e:\O365AdminShared\EMailTemplates\ShrMbxNeedOwners.oft
            $MsgSubject = "Action Required:  Need Shared Mailbox Owner for " + $Script:txtDispName.Text
            $Date30Days = (Get-date).AddDays(30).ToString('yyyy-MMM-dd')
            $NoOwnerTip = "Message sent to obtain group owners will be deleted if no response by " + (Get-date).AddDays(30).ToString('yyyy-MMM-dd') + " - Per No Owners Remaining"
            write-host "           SendTo: " $Script:MsgTo
            $LineToWrite = $RecordEvent + "MSGTO" + "`t" + "                  Message To: " + $Script:MsgTo
            WriteReportEvent
            write-host "  Message Subject: " $MsgSubject
            $LineToWrite = $RecordEvent + "SUBJ" + "`t" + "             Message Subject: " + $MsgSubject
            WriteReportEvent
            write-host "Date Plus 30 Days: " $Date30Days
            $LineToWrite = $RecordEvent + "30Da" + "`t" + "                     30 Days: " + $Script:Date
            WriteReportEvent
            $Output = $wshell.Popup("Message details: $Script:MsgTo `nMessage Subject: $MsgSubject `nDate (+30 Days): $Date30Days" + "",30,"Details For Message",0+32)

            If ($Script:txtNonStdAcc.Text -ne "")
            {
                write-host "Check to move active accounts to correct groups if the groups exist"
                $MbxPerm = get-mailboxfolderPermission $Script:txtDispName.Text |Where-Object {$_.AccessRights -notlike "*None*"}
                foreach ($MbxPerm in $MbxPerm)
                {
                    If ($MbxPerm.User.Displayname -notlike "MBX*")
                    {
                        $UsrMbxExists = [bool]($UsrMbx = get-mailbox $MbxPerm.User.DisplayName -ErrorAction SilentlyContinue)
                        If ($UsrMbxExists -eq $True)
                        {
                            If ($UsrMbx.CustomAttribute2 -eq "A")
                            {
                                If (($MbxPerm.AccessRights -eq "Owner") -and ($Script:txtEDGrpName.Text -notlike "Not Found*"))
                                {
                                    $ConfOwner = [bool](Get-DistributionGroup $cript:txtEDGrpName.Text | ?{$_.managedby -like $UsrMbx.Alias})
                                    If ($ConfOwner -eq $False)
                                    {
                                        Add-DistributionGroupMember $Script:txtREGrpName.Text -Member $MbxPerm.User.Displayname -ByPassSecurityGroupManagerCheck -ErrorAction SilentlyContinue
    #                                    Set-DistributionGroup $Script:txtEDGrpName.Text -ManagedBy @{add="$UsrMbx.Alias"} -BypassSecurityGroupManagerCheck -confirm:$False
                                        $LineToWrite = $RecordEvent + "ADDM" + "`t" + "       Added to Editor Group: " + $UsrMbx.Alias + " to " + $Script:txtEDGrpName.Text
                                        WriteReportEvent
    #                                    Remove-MailboxFolderPermission $Script:txtDispName.Text -User $MbxPerm.User.Displayname
                                    }
                                }
                                If ((($MbxPerm.AccessRights -eq "Editor") -or ($MbxPerm.AccessRights -eq "Owner")) -and ($Script:txtEDGrpName.Text -notlike "Not Found*"))
                                {
                                    $ConfMember = [bool]((Get-DistributionGroupMember $Script:txtEDGrpName.Text -ResultSize Unlimited) | ?{$_.name -like $UsrMbx.Alias})
                                    If ($ConfMember -eq $False)
                                    {
                                        Add-DistributionGroupMember $Script:txtEDGrpName.Text -Member $MbxPerm.User.Displayname -ByPassSecurityGroupManagerCheck -ErrorAction SilentlyContinue
                                        $LineToWrite = $RecordEvent + "ADDM" + "`t" + "       Added to Editor Group: " + $UsrMbx.Alias + " to " + $Script:txtEDGrpName.Text
                                        WriteReportEvent
                                    }
                                    else
                                    {
                                        Remove-MailboxFolderPermission $Script:txtDispName.Text -User $MbxPerm.User.Displayname -Confirm:$False
                                        $LineToWrite = $RecordEvent + "REMO" + "`t" + "  Already member of ED Group: " + $UsrMbx.Alias + " to " + $Script:txtEDGrpName.Text
                                        WriteReportEvent
                                    }
                                }
                                If (($MbxPerm.AccessRights -eq "PublishingAuthor") -and ($Script:txtAUGrpName.Text -notlike "Not Found*"))
                                {
                                    $ConfMember = [bool]((Get-DistributionGroupMember Script:txtAUGrpName.Text -ResultSize Unlimited) | ?{$_.name -like $UsrMbx.Alias})
                                    If ($ConfMember -eq $False)
                                    {
                                        Add-DistributionGroupMember $Script:txtAUGrpName.Text -Member $MbxPerm.User.Displayname -ByPassSecurityGroupManagerCheck -ErrorAction SilentlyContinue
                                        $LineToWrite = $RecordEvent + "ADDM" + "`t" + "       Added to Author Group: " + $UsrMbx.Alias + " to " + $Script:txtAUGrpName.Text
                                        WriteReportEvent
                                    }
                                    else
                                    {
                                        Remove-MailboxFolderPermission $Script:txtDispName.Text -User $MbxPerm.User.Displayname -Confirm:$False
                                        $LineToWrite = $RecordEvent + "REMO" + "`t" + "  Already member of AU Group: " + $UsrMbx.Alias + " to " + $Script:txtAUGrpName.Text
                                        WriteReportEvent
                                    }
                                }
                                If (($MbxPerm.AccessRights -eq "Reviewer") -and ($Script:txtREGrpName.Text -notlike "Not Found*"))
                                {
                                    $ConfMember = [bool]((Get-DistributionGroupMember Script:txtREGrpName.Text -ResultSize Unlimited) | ?{$_.name -like $UsrMbx.Alias})
                                    If ($ConfMember -eq $False)
                                    {
                                        Add-DistributionGroupMember $Script:txtREGrpName.Text -Member $MbxPerm.User.Displayname -ByPassSecurityGroupManagerCheck -ErrorAction SilentlyContinue
                                        $LineToWrite = $RecordEvent + "ADDM" + "`t" + "       Added to Reader Group: " + $UsrMbx.Alias + " to " + $Script:txtREGrpName.Text
                                        WriteReportEvent
                                    }
                                    else
                                    {
                                        Remove-MailboxFolderPermission $Script:txtDispName.Text -User $MbxPerm.User.Displayname -Confirm:$False
                                        $LineToWrite = $RecordEvent + "REMO" + "`t" + "  Already member of RE Group: " + $UsrMbx.Alias + " to " + $Script:txtREGrpName.Text
                                        WriteReportEvent
                                    }
                                }
                            }
                            else
                            {
                                Remove-MailboxFolderPermission $Script:txtDispName.Text -User $MbxPerm.User.Displayname 
                                $LineToWrite = $RecordEvent + "EDTIP" + "`t" + "Removing " + $MbxPerm.AccessRights + " folder permissions from: " + $MbxPerm.User.Displayname + "who is a terminated user"
                                WriteReportEvent
                            }
                        }
                        else
                        {
                            Remove-MailboxFolderPermission $Script:txtDispName.Text -User $MbxPerm.User.Displayname
                            $LineToWrite = $RecordEvent + "EDTIP" + "`t" + "Removing " + $MbxPerm.AccessRights + " folder permissions from: " + $MbxPerm.User.Displayname + "who has been purged from the system"
                            WriteReportEvent
                        }
                    }
                }
            }

            $LineToWrite = $RecordEvent + "NTIP" + "`t" + "               New Mail Tips: " + $NoOwnerTip
            WriteReportEvent
            Set-Mailbox $Script:txtDispName.Text -MailTip $NoOwnerTip
            $LineToWrite = $RecordEvent + "EDTIP" + "`t" + "   Resetting Mailbox Tip On: " + $Script:txtDispName.Text
            WriteReportEvent

            If ($Script:txtEDGrpName.Text -ne "Not Found")
            {
                $Grp = $Script:txtEDGrpName.Text -split ", "
                Foreach ($Grp in $Grp)
                {
                    Set-DistributionGroup $Grp -MailTip $NoOwnerTip
                    $LineToWrite = $RecordEvent + "EDTIP" + "`t" + "      Setting NoOwner Tip On: " + $Grp
                    WriteReportEvent
                }
            }
            If ($Script:txtAUGrpName.Text -ne "Not Found")
            {
                $Grp = $Script:txtAUGrpName.Text -split ", "
                Foreach ($Grp in $Grp)
                {
                    Set-DistributionGroup $Grp -MailTip $NoOwnerTip
                    $LineToWrite = $RecordEvent + "EDTIP" + "`t" + "      Setting NoOwner Tip On: " + $Grp
                    WriteReportEvent
                }
            }
            If ($Script:txtREGrpName.Text -ne "Not Found")
            {
                $Grp = $Script:txtREGrpName.Text -split ", "
                Foreach ($Grp in $Grp)
                {
                    Set-DistributionGroup $Grp -MailTip $NoOwnerTip
                    $LineToWrite = $RecordEvent + "EDTIP" + "`t" + "      Setting NoOwner Tip On: " + $Grp
                    WriteReportEvent
                }
            }
        }
        write-host "Shared Mailbox Request for Owners Complete" -ForegroundColor Cyan
        $LineToWrite = $RecordEvent + "DONE" + "`t" + "Shared Mailbox Request for Owners Complete"
        WriteReportEvent
WRITE-HOST "Message to: " $Script:MsgTo
    }
}
else
{
    write-host "Shared Mailbox Removal Cancelled" -ForegroundColor Red
    $Output = $wshell.Popup("Shared mailbox removal request cancelled.",0,"Cancelled",0+32)
}