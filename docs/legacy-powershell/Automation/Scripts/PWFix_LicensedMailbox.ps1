<#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange
#    'Description  : Enables Service Account for use as a Licensed Mailbox
#    'Called By    : SDAPAdminMenu.ps1 and O365AdminMenu.ps1
#    'Calls        : Connect0365 scripts
#    'Parameters   :11
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 04/25/2021
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 04/25/2021 New Script
#                  : 09/14/2021 Fixed code so if the mailbox name starts with UL it does not change replace the "UL" in "ul.com" with uppercase characters and added code to stop execution if the Exchange Modules are not installed
#                  : 10/04/2021 Added code to set the MSOLUser PasswordNeverExpires value to $True for the surfacehub accounts
#                  : 01/28/2023 Added code to check if OutOfPolicy or Delegate is a licensed mailbox or a valid distribution group.
# ==========================================================================
#
#################################################################################>

Function Build-SvcAccountDetails
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Enable/License Service Account" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 600 ; $Global:form.Height = 300  # Make the form wider 
    
    $Script:Left = 20
    $Script:LeftInput = 150
    $Script:Top = 20
    
    #Name of Account (20 Characters)
    $Script:lblActName = New-Object System.Windows.Forms.Label   
        $Script:lblActName.Text = "Account Name:"
        $Script:lblActName.Top = $Script:Top ; $Script:lblActName.Left = $Script:Left; $Script:lblActName.Width=150 ;$Script:lblActName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblActName)    # Add to Form 
    $Script:txtActName = New-Object Windows.Forms.TextBox
        $Script:txtActName.Top = $Script:Top; $Script:txtActName.Left = $Script:LeftInput; $Script:txtActName.Width = 250;
        $Script:txtActName.TabIndex = 1
        $Script:txtActName.Text = ""
        $Script:InputFocus = $Script:txtActName
        $Global:form.Controls.Add($Script:txtActName)    # Add to Form
        $Global:InputFocus = $Script:txtActName
                
    $Script:Top = $Script:Top + 30
    ## TASK #
    $Script:lblTask = New-Object System.Windows.Forms.Label   
        $Script:lblTask.Text = "Task Number:"
        $Script:lblTask.Top = $Script:Top ; $Script:lblTask.Left = $Script:Left; $Script:lblTask.Width=150 ;$Script:lblTask.AutoSize = $true
        $Global:form.Controls.Add($Script:lblTask)    # Add to Form 
    $Script:txtTask = New-Object Windows.Forms.TextBox
        $Script:txtTask.Top = $Script:Top; $Script:txtTask.Left = $Script:LeftInput; $Script:txtTask.Width = 250;  
        $Script:txtTask.Text = "TASK"
        $Script:txtTask.TabIndex = 2
        $Global:form.Controls.Add($Script:txtTask)    # Add to Form

    $Script:Top = $Script:Top + 30
    ## DisplayName
    $Script:lblDispName = New-Object System.Windows.Forms.Label   
        $Script:lblDispName.Text = "Mailbox Display Name:" 
        $Script:lblDispName.Top = $Script:Top ; $Script:lblDispName.Left = $Script:Left; $Script:lblDispName.Width=150 ;$Script:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblDispName)    # Add to Form 
    $Script:txtDispName = New-Object Windows.Forms.TextBox
        $Script:txtDispName.Top = $Script:Top; $Script:txtDispName.Left = $Script:LeftInput; $Script:txtDispName.Width = 250;
        $Script:txtDispName.ReadOnly = $false
        $Script:txtDispName.TabIndex = 3
        $Global:form.Controls.Add($Script:txtDispName)    # Add to Form
        $Script:txtDispName.Add_Click({$Script:txtMailAddr.Text = "Click Here to Get Address"})

    $Script:Top = $Script:Top + 30
    ## EMail Address
    $Script:lblMailAddr = New-Object System.Windows.Forms.Label   
        $Script:lblMailAddr.Text = "Email Address:" 
        $Script:lblMailAddr.Top = $Script:Top ; $Script:lblMailAddr.Left = $Script:Left; $Script:lblMailAddr.Width=150 ;$Script:lblMailAddr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblMailAddr)    # Add to Form 
    $Script:txtMailAddr = New-Object Windows.Forms.TextBox
        $Script:txtMailAddr.Top = $Script:Top; $Script:txtMailAddr.Left = $Script:LeftInput; $Script:txtMailAddr.Width = 250;
        $Script:txtMailAddr.ReadOnly = $false
        $Script:txtMailAddr.TabIndex = 4
        $SCript:txtMailAddr.Text = "Click Here to Get Address"
        $Global:form.Controls.Add($Script:txtMailAddr)    # Add to Form
        $Script:txtMailAddr.Add_Click({
            #Update the Email Address
            $Loc = $Script:txtDispName.Text.Substring(0,$Script:txtDispName.Text.indexof(" "))
            If ($Script:txtDispName.Text -like "*-*") # SurfaceHubs have a "-" in the name however all details after the "-" for all mailboxes is excluded
            {
                $DName = $Script:txtDispName.Text
                If ($DName -like "*-*")
                {
                    $NewDispName = ($DName.Substring(0,$DName.indexof("-"))).TrimEnd()
                }
                else
                {
                    # this checks for a long "-" in the name and it replaces it with a regular dash
                    $NewDispName = ($DName.Substring(0,$DName.indexof("–"))).TrimEnd()
                    $Script:txtDispName.Text = $DName -replace ("–","-")
                }
                $Script:txtMailAddr.Text = ($NewDispName -Replace($Loc,($Loc +".")))
                $Script:txtMailAddr.Text = (($Script:txtMailAddr.Text -Replace(" ","")) + "@ul.com")
            }
            else
            {
                $Script:txtMailAddr.Text = ($Script:txtDispName.Text -Replace($Loc,($Loc +".")))
                $Script:txtMailAddr.Text = (($Script:txtMailAddr.Text -Replace(" ","")) + "@ul.com")
            }
            If ($script:chkEditors.Checked -eq $true)
            {
                $Script:txtEDGrpName.Text = ("MBX." + $Script:txtDispName.Text + ".ED")
                $Script:txtEDGrpAddr.Text = ("MBX." + ($Script:txtMailAddr.Text.Substring(0,$Script:txtMailAddr.Text.IndexOf("@"))) + ".ED@ul.com")
            }
            If ($script:chkAuthors.Checked -eq $true)
            {
                $Script:txtAUGrpName.Text = ("MBX." + $Script:txtDispName.Text + ".AU")
                $Script:txtAUGrpAddr.Text = ("MBX." + ($Script:txtMailAddr.Text.Substring(0,$Script:txtMailAddr.Text.IndexOf("@"))) + ".AU@ul.com")
            }
            If ($script:chkEditors.Checked -eq $true)
            {
                $Script:txtREGrpName.Text = ("MBX." + $Script:txtDispName.Text + ".RE")
                $Script:txtREGrpAddr.Text = ("MBX." + ($Script:txtMailAddr.Text.Substring(0,$Script:txtMailAddr.Text.IndexOf("@"))) + ".RE@ul.com")
            }

            Check-MbxExists
            $Script:chkLicMbx.Visible = $True
            $Script:chkKFIOFR.Visible = $True
            $Script:chkSurHub.Visible = $True
        })
        $Script:lblMailAddr.Visible = $true
        $Script:txtMailAddr.Visible = $true

    $Script:Top = $Script:Top + 30
    ## Account Owner
    $Script:lblAcctOwner = New-Object System.Windows.Forms.Label   
        $Script:lblAcctOwner.Text = "Account Owner:" 
        $Script:lblAcctOwner.Top = $Script:Top ; $Script:lblAcctOwner.Left = $Script:Left; $Script:lblAcctOwner.Width=150 ;$Script:lblAcctOwner.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblAcctOwner)    # Add to Form 
    $Script:txtAcctOwner = New-Object Windows.Forms.TextBox
        $Script:txtAcctOwner.Top = $Script:Top; $Script:txtAcctOwner.Left = $Script:LeftInput; $Script:txtAcctOwner.Width = 250;
        $Script:txtAcctOwner.Text = "(Emp# or Address separate with commas)"
        $Script:txtAcctOwner.ReadOnly = $false
        $Script:txtAcctOwner.TabIndex = 4
        $Global:form.Controls.Add($Script:txtAcctOwner)    # Add to Form
        $Script:lblAcctOwner.visible = $false
        $script:txtAcctOwner.visible = $false
        $Script:txtAcctOwner.Add_Click({
            If ($Script:txtAcctOwner.Text -eq "(Emp# or Address separate with commas)")
            {
                $Script:txtAcctOwner.Text = ""
                $Script:txtAcctOwner.Refresh()
            }
        })

    $Script:Top = $Script:Top + 30
    #Report Details
    $Script:lblRptFile = New-Object System.Windows.Forms.Label   
        $Script:lblRptFile.Text = "Report Details:"  
        $Script:lblRptFile.Top = $Script:Top ; $Script:lblRptFile.Left = $Script:Left; $Script:lblRptFile.Width=100 ;$Script:lblRptFile.AutoSize = $true 
    $Global:form.Controls.Add($Script:lblRptFile)    # Add to Form 
        #
    $Script:txtRptFile = New-Object Windows.Forms.TextBox
        $Script:txtRptFile.Top = $Script:Top; $Script:txtRptFile.Left = $Script:LeftInput; $Script:txtRptFile.Width = 400;
        $Script:txtRptFile.ReadOnly = $false
        $Script:txtRptFile.TabIndex = 5
        $Global:form.Controls.Add($Script:txtRptFile)    # Add to Form        
        $Script:lblRptFile.visible = $false
        $script:txtRptFile.visible = $false

    #Add Checkboxes
    ## Licensed Mailbox 
    $Script:chkLicMbx = New-Object Windows.Forms.RadioButton
        $Script:chkLicMbx.Left = 420; $Script:chkLicMbx.Width = 160; $Script:chkLicMbx.Top = 20
        $Script:chkLicMbx.Text = "Licensed Mailbox" 
        $Script:chkLicMbx.Checked = $Script:chkLicMbx.Checked   # set a default value 
        $Script:chkLicMbx.TabIndex = 0
        $Global:form.Controls.Add($Script:chkLicMbx) 
        # Obtain Value with: $Script:chkThis.Checked
        $Script:chkLicMbx.Visible = $False
        $Script:chkLicMbx.Add_Click({
            If ($Script:txtDispName.Text.Length -ne 0)
            {
                If ($Script:txtActName.Text -notlike "*@*")
                {
                    $Script:txtActName.Text = $Script:txtActName.Text + "@global.ul.com"
                } 
                Check-ADExists
                Check-MbxExists
                If (($Script:ADExists -eq $true) -and ($Script:MbxExists -eq $false))
                {
                    Add-LicensedMbxDetails
                }
                else
                {
                    $script:chkLicMbx.Checked = $false
                }
            }
            else
            {
                $script:chkLicMbx.Checked = $false
            }
        })

    ## Oracle KFI/OFR Mailbox
    $Script:chkKFIOFR = New-Object Windows.Forms.RadioButton
        $Script:chkKFIOFR.Left = 420; $Script:chkKFIOFR.Width = 160; $Script:chkKFIOFR.Top = 40
        $Script:chkKFIOFR.Text = "Oracle KFI/OFR Mailbox" 
        $Script:chkKFIOFR.Checked = $Script:chkKFIOFR.Checked   # set a default value 
        $Script:chkKFIOFR.TabIndex = 1
        $Script:chkKFIOFR.Visible = $False
        $Global:form.Controls.Add($Script:chkKFIOFR) 
        # Obtain Value with: $Script:chkKFIOFR.Checked
        $Script:chkKFIOFR.Add_Click({Add-KFIOFRMbxDetails})

    ## Surface Hub Mailbox
    $Script:chkSurHub = New-Object Windows.Forms.RadioButton
        $Script:chkSurHub.Left = 420; $Script:chkSurHub.Width = 160; $Script:chkSurHub.Top = 60  
        $Script:chkSurHub.Text = "Surface Hub Device" 
        $Script:chkSurHub.Checked = $Script:chkSurHub.Checked   # set a default value 
        $Script:chkSurHub.TabIndex = 1
        $Script:chkSurHub.Visible = $False
        $Global:form.Controls.Add($Script:chkSurHub) 
        # Obtain Value with: $Script:chkSurHub.Checked
        $Script:chkSurHub.Add_Click({Add-SurfaceHubDetails})

    $Script:Top = $Script:Top + 30 #Keep this so data is spaced properly
}

Function Check-ADExists
{
    $Script:ActName = $Script:txtActName.Text
    If (($Script:ActName.Contains("@")) -or ($Script:ActName.Length -gt 20))
    {
        If ($Script:ActName.Text.Length -gt 20)
        {
            $Script:ActName = $Script:ActName.Substring(0,20)
        }

        If ($Script:ActName.Contains("@"))
        {
            $Script:ActName = $Script:ActName.Substring(0,$Script:ActName.IndexOf("@"))
        }
    }

    $accountname = $Script:ActName

    $Script:ADExists = [bool](get-ADUser -Filter {SamAccountName -eq $accountname} -ErrorAction SilentlyContinue)
    If ($Script:ADExists -ne $true)
    {
        write-host "No Service Account Found with the name: " $accountname -ForegroundColor Red
        $Output = $wshell.Popup("No Service Account found for " + $accountname,0,"No Account Found",0+32)
    }
    else
    {
        $Script:txtAcctOwner.Text = (get-aduser $accountname -Properties ExtensionAttribute2).ExtensionAttribute2
    }
}

Function Check-MbxExists
{
    $Script:MbxExists = [bool](get-mailbox $Script:txtMailAddr.Text -ErrorAction Silentlycontinue)
    If ($Script:MbxExists -eq $true)
    {
        write-host "There is already a mailbox with the address: " $script:txtMailAddr.Text -ForegroundColor Red
        $Output = $wshell.Popup("There is already a mailbox with the address " + $Script:txtMailAddr.Text,0,"Address In Use",0+32)
    }
}

Function Add-LicensedMbxDetails
{
#Code to add details for MBX groups

    $Global:form.Text = "Enable Service Account as Licensed Mailbox"
    $Script:lblAcctOwner.visible = $true
    $Script:txtAcctOwner.visible = $true
    $Script:lblRptFile.visible = $true
    $script:txtRptFile.visible = $true
    $Script:chkLicMbx.visible = $false
    $Script:chkKFIOFR.visible = $false
    $Script:chkSurHub.visible = $false
    $script:txtRptFile.Text = "e:\Automation\EnableSvcActAssignLicense\Report\Report-EnableSvcActAssignLicense-" + ($Script:txtDispName.Text -replace(" ","")) + "-Date" + $Date + "-Time" + $Time + ".log"

    ## Editors Group CheckBox
    $Script:chkEditors = New-Object Windows.Forms.checkbox 
    $script:chkEditors.Left = $Script:Left; $script:chkEditors.Width = 160; $script:chkEditors.Top = $Script:Top
    $script:chkEditors.Text = "Create (.ED) Editors Group" 
    $script:chkEditors.Checked = $false   # set a default value 
    $script:chkEditors.TabIndex = 2
    $Global:form.Controls.Add($script:chkEditors)
    $Script:chkEditors.Add_Click({
        If ($script:chkEditors.Checked -eq $true)
        {
            $Global:form.Height = 480
            ## Editor Group Name
            $Script:txtEDGrpName = New-Object Windows.Forms.TextBox  
            $Script:txtEDGrpName.TabIndex = 0 # set Tab Order 
            $Script:txtEDGrpName.Top = ($Script:Top+30); $Script:txtEDGrpName.Left = $Script:Left; $Script:txtEDGrpName.Width = 160;  
            $Script:txtEDGrpName.Text = ("MBX." + $Script:txtDispName.Text + ".ED")   # Use Corrent computer name as default 
            $Script:txtEDGrpName.ReadOnly = $true
            $Global:form.Controls.Add($Script:txtEDGrpName)    # Add to Form 

            ## Editor Group Address
            $Script:txtEDGrpAddr = New-Object Windows.Forms.TextBox  
            $Script:txtEDGrpAddr.TabIndex = 0 # set Tab Order
            $Script:txtEDGrpAddr.Top = ($Script:Top+60); $Script:txtEDGrpAddr.Left = $Script:Left; $Script:txtEDGrpAddr.Width = 160;
            $Script:txtEDGrpAddr.Text = ("MBX." + ($Script:txtMailAddr.Text.Substring(0,$Script:txtMailAddr.Text.IndexOf("@"))) + ".ED@ul.com")
            $Script:txtEDGrpAddr.ReadOnly = $true
            $Global:form.Controls.Add($Script:txtEDGrpAddr)    # Add to Form

            ## Editor Group Members
            $Script:lblEDGrpMbr = New-Object System.Windows.Forms.Label   
            $Script:lblEDGrpMbr.Text = ".ED Member(s):"  
            $Script:lblEDGrpMbr.Top = ($Script:Top+90); $Script:lblEDGrpMbr.Left = $Script:Left; $Script:lblEDGrpMbr.Width=160; $Script:lblEDGrpMbr.AutoSize = $true 
            $Global:form.Controls.Add($Script:lblEDGrpMbr)    # Add to Form 
            # 
            $Script:txtEDGrpMbr = New-Object Windows.Forms.TextBox
            $Script:txtEDGrpMbr.TabIndex = 0 # set Tab Order 
            $Script:txtEDGrpMbr.Location = New-Object System.Drawing.Size($Script:Left,($Script:Top+110))
            $Script:txtEDGrpMbr.Size = New-Object system.Drawing.Size(160,80)
            $Script:txtEDGrpMbr.MultiLine = $true
            $Script:txtEDGrpMbr.ScrollBars = 'Both'  
            $Global:form.Controls.Add($Script:txtEDGrpMbr)    # Add to Form
        }
        else
        {
            If (($Script:chkAuthors.Checked -eq $false) -and ($Script:chkReaders.Checked -eq $false))
            {
                $Global:form.Height = 300
            }
            $Script:txtEDGrpName.Visible = $false
            $Script:txtEDGrpAddr.visible = $false
            $Script:lblEDGrpMbr.visible = $false
            $Script:txtEDGrpMbr.visible = $false
        }         
    })

    ## Authors Group CheckBox
    $Script:chkAuthors = New-Object Windows.Forms.checkbox 
    $Script:chkAuthors.Left = ($Script:Left+180); $Script:chkAuthors.Width = 180; $Script:chkAuthors.Top = $Script:Top
    $Script:chkAuthors.Text = "Create (.AU) Authors Group" 
    $Script:chkAuthors.Checked = $false   # set a default value 
    $Script:chkAuthors.TabIndex = 2
    $Global:form.Controls.Add($Script:chkAuthors)
    $Script:chkAuthors.Add_Click({
        If ($script:chkAuthors.Checked -eq $true)
        {
            $Global:form.Height = 480
            ## Author Group Name
            $Script:txtAUGrpName = New-Object Windows.Forms.TextBox  
            $Script:txtAUGrpName.TabIndex = 0 # set Tab Order 
            $Script:txtAUGrpName.Top = ($Script:Top+30); $Script:txtAUGrpName.Left = $Script:Left+180; $Script:txtAUGrpName.Width = 160;  
            $Script:txtAUGrpName.Text = ("MBX." + $Script:txtDispName.Text + ".AU")   # Use Corrent computer name as default
            $Script:txtAUGrpName.ReadOnly = $true
            $Global:form.Controls.Add($Script:txtAUGrpName)    # Add to Form 

            ## Author Group Address
            $Script:txtAUGrpAddr = New-Object Windows.Forms.TextBox  
            $Script:txtAUGrpAddr.TabIndex = 0 # set Tab Order
            $Script:txtAUGrpAddr.Top = ($Script:Top+60); $Script:txtAUGrpAddr.Left = $Script:Left+180; $Script:txtAUGrpAddr.Width = 160;
            $Script:txtAUGrpAddr.Text = ("MBX." + ($Script:txtMailAddr.Text.Substring(0,$Script:txtMailAddr.Text.IndexOf("@"))) + ".AU@ul.com")
            $Script:txtAUGrpAddr.ReadOnly = $true
            $Global:form.Controls.Add($Script:txtAUGrpAddr)    # Add to Form

            ## Author Group Members
            $Script:lblAUGrpMbr = New-Object System.Windows.Forms.Label   
            $Script:lblAUGrpMbr.Text = ".AU Member(s):"  
            $Script:lblAUGrpMbr.Top = ($Script:Top+90); $Script:lblAUGrpMbr.Left = $Script:Left+180; $Script:lblAUGrpMbr.Width=160; $Script:lblAUGrpMbr.AutoSize = $true 
            $Global:form.Controls.Add($Script:lblAUGrpMbr)    # Add to Form 
            # 
            $Script:txtAUGrpMbr = New-Object Windows.Forms.TextBox
            $Script:txtAUGrpMbr.TabIndex = 0 # set Tab Order 
            $Script:txtAUGrpMbr.Location = New-Object System.Drawing.Size(($Script:Left+180),($Script:Top+110))
            $Script:txtAUGrpMbr.Size = New-Object system.Drawing.Size(160,80)
            $Script:txtAUGrpMbr.MultiLine = $true
            $Script:txtAUGrpMbr.ScrollBars = 'Both'  
            $Global:form.Controls.Add($Script:txtAUGrpMbr)    # Add to Form
        }
        else
        {
            If (($Script:chkEditors.Checked -eq $false) -and ($Script:chkReaders.Checked -eq $false))
            {
                $Global:form.Height = 300
            }
            $Script:txtAUGrpName.Visible = $false
            $Script:txtAUGrpAddr.visible = $false
            $Script:lblAUGrpMbr.visible = $false
            $Script:txtAUGrpMbr.visible = $false
        }
    })         
 
    ## Readers Group CheckBox
    $Script:chkReaders = New-Object Windows.Forms.checkbox 
    $Script:chkReaders.Left = ($Script:Left+360); $Script:chkReaders.Width = 180; $Script:chkReaders.Top = $Script:Top
    $Script:chkReaders.Text = "Create (.RE) Readers Group" 
    $Script:chkReaders.Checked = $false   # set a default value 
    $Script:chkReaders.TabIndex = 2
    $Global:form.Controls.Add($Script:chkReaders)
    $Script:chkReaders.Add_Click({
        If ($script:chkReaders.Checked -eq $true)
        {
            $Global:form.Height = 480
            ## Reader Group Name
            $Script:txtREGrpName = New-Object Windows.Forms.TextBox  
            $Script:txtREGrpName.TabIndex = 0 # set Tab Order 
            $Script:txtREGrpName.Top = ($Script:Top+30); $Script:txtREGrpName.Left = $Script:Left+360; $Script:txtREGrpName.Width = 160;  
            $Script:txtREGrpName.Text = ("MBX." + $Script:txtDispName.Text + ".RE")   # Use Corrent computer name as default 
            $Script:txtREGrpName.ReadOnly = $true
            $Global:form.Controls.Add($Script:txtREGrpName)    # Add to Form 

            ## Reader Group Address
            $Script:txtREGrpAddr = New-Object Windows.Forms.TextBox  
            $Script:txtREGrpAddr.TabIndex = 0 # set Tab Order
            $Script:txtREGrpAddr.Top = ($Script:Top+60); $Script:txtREGrpAddr.Left = $Script:Left+360; $Script:txtREGrpAddr.Width = 160;
            $Script:txtREGrpAddr.Text = ("MBX." + ($Script:txtMailAddr.Text.Substring(0,$Script:txtMailAddr.Text.IndexOf("@"))) + ".RE@ul.com")
            $Script:txtREGrpAddr.ReadOnly = $true
            $Global:form.Controls.Add($Script:txtREGrpAddr)    # Add to Form

            ## Reader Group Members
            $Script:lblREGrpMbr = New-Object System.Windows.Forms.Label   
            $Script:lblREGrpMbr.Text = ".RE Member(s):"  
            $Script:lblREGrpMbr.Top = ($Script:Top+90); $Script:lblREGrpMbr.Left = $Script:Left+360; $Script:lblREGrpMbr.Width=160; $Script:lblREGrpMbr.AutoSize = $true 
            $Global:form.Controls.Add($Script:lblREGrpMbr)    # Add to Form 
            # 
            $Script:txtREGrpMbr = New-Object Windows.Forms.TextBox
            $Script:txtREGrpMbr.TabIndex = 0 # set Tab Order 
            $Script:txtREGrpMbr.Location = New-Object System.Drawing.Size(($Script:Left+360),($Script:Top+110))
            $Script:txtREGrpMbr.Size = New-Object system.Drawing.Size(160,80)
            $Script:txtREGrpMbr.MultiLine = $true
            $Script:txtREGrpMbr.ScrollBars = 'Both'  
            $Global:form.Controls.Add($Script:txtREGrpMbr)    # Add to Form
        }
        else
        {
            If (($Script:chkEditors.Checked -eq $false) -and ($Script:chkAuthors.Checked -eq $false))
            {
                $Global:form.Height = 300
            }
            $Script:txtREGrpName.Visible = $false
            $Script:txtREGrpAddr.visible = $false
            $Script:lblREGrpMbr.visible = $false
            $Script:txtREGrpMbr.visible = $false
        }
    })           

    Add-FormStandardButtons
}

Function Add-KFIOFRMbxDetails
{
    write-host "Add code to add details for KFI/OFR configs (under developent)" -ForegroundColor Red
    Add-FormStandardButtons
}

Function Add-SurfaceHubDetails
{
    $Script:lblAcctOwner.Text = "Room Delegate Group:"
    $Script:txtAcctOwner.Text = "MBX." + $Script:txtActName.Text.substring(4,3) + ".RRS.OutOfPolicy.Surface.DE"
    $Script:lblAcctOwner.visible = $True
    $script:txtAcctOwner.visible = $True
    $Script:lblMailAddr.Visible = $True
    $Script:txtMailAddr.Visible = $True
    $Script:lblRptFile.visible = $True
    $script:txtRptFile.visible = $True

    If ($Script:txtActName.Text -like "*@")
    {
        $Script:txtRptFile.Text = "E:\Automation\EnableSvcActAssignLicense\Report\Report-EnableSvcActAssignLicense-" + $Script:txtActName.Text.Substring(0,$Script:txtActName.Text.IndexOf("@")) + "-Date" + $Date + "-Time" + $Time + ".log"
    }
    else
    {
        $Script:txtRptFile.Text = "E:\Automation\EnableSvcActAssignLicense\Report\Report-EnableSvcActAssignLicense-" + $Script:txtActName.Text + "-Date" + $Date + "-Time" + $Time + ".log"
    }
    Add-FormStandardButtons
}

function Add-SvcFolderPermissions
{
    if ($GrpAddr.contains("@"))
	{
		if ((Get-Mailbox $ADUser.UserPrincipalName).MailTip -eq $null)
        {
            $NewTip = ("Owners: " + (((Get-DistributionGroup $GrpAddr).Managedby) -join ", "))
            if ($NewTip.Length -gt 175)
            {
                write-host "Maximum mail tip length exceeded truncating to 175 characters." -foregroundcolor Red
                $NewTip = $NewTip.Substring(0,175)
                $OldTip = (Get-Mailbox $ADUser.UserPrincipalName).MailTip
                write-host "Replacing old MailTip: " $OldTip.Substring(16,$OldTip.Length-36) -foregroundcolor Cyan
                write-host "                 With: " $NewTip
		        $LineToWrite = "REPL  " + "`t" + "Replacing old MailTip with " + $NewTip
		        WriteReportEvent
            }
            Set-Mailbox $ADUser.UserPrincipalName -MailTip $NewTip
        }

        $Permission = ""
        if ($GrpAddr.contains(".ED@"))
        {
            $Permission = "Editor"
       		write-host "Setting permissions on: " $ADUser.UserPrincipalName " <-- " $GrpAddr " (Editor)" -ForegroundColor Cyan
	    	$LineToWrite = "INFO" + "`t" + "Setting Editor Permissions for " + $ADUser.UserPrincipalName +  " for " + $GrpAddr
		   	WriteReportEvent
            Write-Host "GrantSendOnBehalf to Delegate Group:  " $GrpAddr -ForegroundColor Cyan
            Set-Mailbox -Identity $ADUser.UserPrincipalName -GrantSendOnBehalfTo ((Get-Mailbox -Identity $ADUser.UserPrincipalName).GrantSendOnBehalfTo += $GrpAddr)
            Add-MailboxPermission $ADUser.UserPrincipalName –User $GrpAddr –AccessRights FullAccess
            Add-RecipientPermission $ADUser.UserPrincipalName -Trustee $GrpAddr –AccessRights SendAs -Confirm:$false
        }
        elseif ($GrpAddr.contains(".AU@"))
        {
            $Permission = "PublishingAuthor"
            write-host "Setting permissions on: " $ADUser.UserPrincipalName " <-- " $GrpAddr " (PublishingAuthor)" -ForegroundColor Cyan
	    	$LineToWrite = "INFO" + "`t" + "Setting Author Permissions for " + $ADUser.UserPrincipalName +  " for " + $GrpAddr
		   	WriteReportEvent
            Write-Host "GrantSendOnBehalf to Delegate Group:  " $GrpAddr -ForegroundColor Cyan
            Set-Mailbox -Identity $ADUser.UserPrincipalName -GrantSendOnBehalfTo ((Get-Mailbox -Identity $ADUser.UserPrincipalName).GrantSendOnBehalfTo += $GrpAddr)
        }
        elseif ($GrpAddr.contains(".RE@"))
        {
            $Permission = "Reviewer"
            write-host "Setting permissions on: " $ADUser.UserPrincipalName " <-- " $GrpAddr " (Reviewer)" -ForegroundColor Cyan
	    	$LineToWrite = "INFO" + "`t" + "Setting Read Permissions on: " + $ADUser.UserPrincipalName + " for " + $GrpAddr
		   	WriteReportEvent
        }
        else
        {
            write-host "Group name does not follow naming stand there is no .ED, .AU or .RE found." -ForegroundColor Red
            write-host "Group Name: " $GrpAddr
            $LineToWrite = "INFO" + "`t" + "Group name does not follow naming stand there is no .ED, .AU or .RE found " + $ADUser.UserPrincipalName
		   	WriteReportEvent
        }
        
        If ($Permission -ne "")
        {
            Do
            {
                Start-Sleep -Seconds 2
                $MBXFoldersExist = [bool](Get-MailboxFolderStatistics -Identity $ADUser.UserPrincipalName -ErrorAction SilentlyContinue)
            } while ($MBXFoldersExist -ne "True")

            $MBXFolders = Get-MailboxFolderStatistics -Identity $ADUser.UserPrincipalName | % {$_.Identity.ToString().Split("\")[1..100] -join "\"}

            ForEach ($Folder in $MBXFolders)
            {
		        If ($Folder.Equals("Recoverable Items") -or $Folder.Equals("Calendar Logging") -or $Folder.Equals("Deletions") -or $Folder.Equals("Purges") -or $Folder.Equals("Versions"))
                {
				    # Ignore folder
				}
				else
                {
				    if ($Folder.Equals("Top of Information Store"))
                    {
                        Add-MailboxFolderPermission -Identity ($ADUser.UserPrincipalName + ":\") -User $GrpAddr -AccessRights $Permission
                        $LineToWrite = "ADD  " + "`t" + $Permission + " permission to: " + $ADUser.UserPrincipalName + ":\"
                    }
					else
                    {
					    Add-MailboxFolderPermission -Identity ($ADUser.UserPrincipalName + ":\" + $Folder) -User $GrpAddr -AccessRights $Permission
                        $LineToWrite = "ADD  " + "`t" + $Permission + " permission to: " + $ADUser.UserPrincipalName + ":\" + $Folder
					}
                    WriteReportEvent
				}
			}
        }
        else
        {
            write-host "Skipping this group since it does not follow our naming standard"
        }
    }	
}

function CreateSvcUSG
{				
    $DstExists = [bool](Get-DistributionGroup $GrpAddr -ResultSize Unlimited -ErrorAction SilentlyContinue)
    if ($DstExists -eq $false)
    {
	    if ($GrpAddr.contains("@"))
        {
            $atMail = $GrpAddr.indexOf("@")
		    $DGAlias = $GrpAddr.substring(0,$atMail)
            $DGManagedByMembers = $Script:txtAcctOwner.Text
            $Task = $Script:txtTask.Text

            $LineToWrite = ""
            WriteReportEvent
        			
		    # Create the USG
		    New-DistributionGroup -Name $GrpName `
		        -PrimarySmtpAddress $GrpAddr `
			    -Alias $DGAlias `
			    -ManagedBy $DGManagedByMembers `
			    -RequireSenderAuthenticationEnabled $TRUE `
			    -Type Security | Out-Null
					
            $Owners = "Owners: " + ((Get-DistributionGroup $GrpName).ManagedBy -join (", "))
        
            Set-Group -identity $GrpName `
		        -Notes ($Owners + " - Per: " + $Script:txtTask.Text)

            If ($Owners.Length -gt 175)
            {
                Set-DistributionGroup $GrpName -MailTip $Owners.Substring(0,175)
            }
            else
            {
                Set-DistributionGroup $GrpName -MailTip $Owners
            }
        
            if (Get-DistributionGroup $GrpAddr)
            {
		        write-host ""
                write-host "USG created: " $GrpName " (" $GrpAddr ")" -ForegroundColor Cyan
		        $LineToWrite = "CREATE" + "`t" + $GrpName + "`t" + $GrpAddr + " - USG created"
		        WriteReportEvent
			
			    write-host "USG Ownner: " $DGManagedByMembers -ForegroundColor Cyan
		        $LineToWrite = "OWNER" + "`t" + $DGManagedByMembers + " - USG Owner"
		        WriteReportEvent
				
		        $USG = Get-DistributionGroup $GrpAddr
		        Set-DistributionGroup $GrpAddr `
			        -RequireSenderAuthenticationEnabled $True `
		 		    -BypassSecurityGroupManagerCheck `
				    -CustomAttribute15 ($FileName + " PS Date: " + (get-date -uformat %D) + " PS Time: " + (get-date -uformat %T))
				
		       # Add USG members
                write-host ("Adding Members to USG: " + $GrpName + " (" + $GrpAddr + ")")
		        $LineToWrite = "MEMBER" + "`t" + "Adding Members to Distribution Group " + $GrpName
		        WriteReportEvent
		
                $addMember = $GrpMem.split(",")
		        if ($? -eq $true)
                {
			        ForEach ($member in $addMember)
                    {
                        $member = $member.Trim()
                        if ($member.length -gt 0)
                        {
    				        if (Get-Mailbox $member -ErrorAction SilentlyContinue)
                            {
		    				    Add-DistributionGroupMember $GrpAddr -Member $member -BypassSecurityGroupManagerCheck
			    			    write-host "Added member:" $member
				    		    $LineToWrite = "PASS" + "`t" + "Added " + $member
					        }
					        else
                            {
	    					    write-host "ERROR finding member: " $member
		    				    $LineToWrite = "FAIL" + "`t" + "Failed adding " + $member
			    		    }
                            WriteReportEvent
	    			    }
                    }
		        }	
		        else
                {
			        write-host "No members to add"
			        $LineToWrite = "FAIL" + "`t" + "No Members to Add"
                    WriteReportEvent
		        }
            }	
		    else
            {
		        Write-Host "USG not created: " $GrpName " (" $GrpAddr ")" -ForegroundColor Red
			    $LineToWrite = "FAIL" + "`t" + $GrpName + "`t" + $GrpAddr + "`t" + "USG not created"
			    WriteReportEvent
	        }
        }
    }
    else
    {
        Write-Host "User Security Group Already Exists - No Changes Made" -ForegroundColor Red
    }
}

Function Enable-LicensedMailbox
{
    $LineToWrite = "STAR" + "`t" + "NewSharedMailbox script has started"
	WriteReportEvent
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	WriteReportEvent

    $LineToWrite = "INFO" + "`t" + "    Account Name: " + $Script:txtActName.Text
    WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "     Task Number: " + $Script:txtTask.Text
    WriteReportEvent
    $LineToWrite = "INFO" + "`t" +  "    Display Name: " + $Script:txtDispName.Text
    WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "   Email Address: " + $Script:txtMailAddr.Text
    WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "   Account Owner: " + $Script:txtAcCtOwner.Text
    WriteReportEvent
    If ($script:chkEditors.Checked -eq $true)
    {
        $LineToWrite = "INFO" + "`t" + "   Editors Group: " + $Script:txtEDGrpName.Text + " (" + $Script:txtEDGrpAddr.Text + ")"
        WriteReportEvent
        $LineToWrite = "INFO" + "`t" + " Editors Members: " + $Script:txtEDGrpMbr.Text
        WriteReportEvent
    }
    If ($script:chkAuthors.Checked -eq $true)
    {
        $LineToWrite = "INFO" + "`t" + "   Authors Group: " + $Script:txtAUGrpName.Text + " (" + $Script:txtAUGrpAddr.Text + ")"
        WriteReportEvent
        $LineToWrite = "INFO" + "`t" + " Authors Members: " + $Script:txtAUGrpMbr.Text
        WriteReportEvent
    }
    If ($script:chkReaders.Checked -eq $true)
    {
        $LineToWrite = "INFO" + "`t" + "   Readers Group: " + $Script:txtREGrpName.Text + " (" + $Script:txtREGrpAddr.Text + ")"
        WriteReportEvent
        $LineToWrite = "INFO" + "`t" + " Readers Members: " + $Script:txtREGrpMbr.Text
        WriteReportEvent
    }
    $LineToWrite = "`n"
    WriteReportEvent
}

$Global:OKDetails = "Enable"
$Filename = "LicensedMailbox"
$Global:AddButtons = 0
$uDate = get-date -uformat %D
$Date  = $uDate.Replace("/", "-")
$uTime = get-date -uformat %T
$Time  = $uTime.Replace(":", "")
$Script:ExSnap = [bool](Get-PSSnapin *Exchange* -ErrorAction Silentlycontinue)

If ($Script:ExSnap -eq $True)
{
    Build-SvcAccountDetails
    Publish-Form
    $DC = "usnbkadds001p.global.ul.com"
    $ReportFile = $script:txtRptFile.Text

    If ($Global:Result -eq "OK")
    {
        If ($Script:chkLicMbx.Checked -eq "Checked")
        {
            Write-host "Mail Enabling and Licensing Service Account [$($Script:ActName)]" -ForegroundColor Green
            Enable-LicensedMailbox

            $ADUser = get-ADUser -Filter {SamAccountName -eq $script:ActName} -Properties sn,givenName,initials,displayName,mail,proxyaddresses,info,userWorkstations,UserPrincipalName,PasswordNeverExpires
            write-host "ADuser.UserPrincipalName: " $ADUser.UserPrincipalName
        #Enable Mailbox
            If (([bool](get-Mailbox $Script:txtMailAddr.Text -ErrorAction SilentlyContinue)) -ne "True")
            {
                $LineToWrite = "INFO" + "`t" + "Account Name: " + $Script:txtActName.Text + "`t" + $DC
                WriteReportEvent

                write-host "Enabling for Mail Use" $Script:txtActName.Text -ForegroundColor Green
        # Disconnect from the PS Session in order for the Get-MailUser command to work
#                Remove-PSSession (Get-PSSession)
                $MEnab = ([bool](Get-MailUser $ADUser.UserPrincipalName -ErrorAction SilentlyContinue))
  	            If ($MEnab -eq $true)
                {
                    write-host "This account is already mail enabled" -ForegroundColor Red
    	            $LineToWrite = "INFO" + "`t" + "Account is Already Mail Enabled for: " + $Script:txtActName.Text
   		            WriteReportEvent
#                    Set-MailUser $ADUser.UserPrincipalName -EmailAddresses (((Get-MailUser $ADUser.UserPrincipalName -DomainController:$DC).$Script:txtMailAddr.Text)) -CustomAttribute15 "EnableServiceAcct PS Date: $Date PS Time: $Time"
                    Set-ADUser $script:ActName -Add @{ProxyAddresses=("SMTP:" + $Script:txtMailAddr.Text)}
                    Set-ADUser $script:ActName -Add @{'ExtensionAttribute15'="EnableMailUser PS Date: $Date PS Time: $Time"}
       	            $LineToWrite = "INFO" + "`t" + "Setting Email Address to: " + $Script:txtMailAddr.Text
    	            WriteReportEvent
                }
                else
                {
                    Enable-MailUser $ADUser.UserPrincipalName -Alias $script:ActName -ExternalEmailaddress ($Script:txtMailAddr.Text) -PrimarySmtpAddress ($Script:txtMailAddr.Text) -DomainController:$DC
#                    Set-MailUser $ADUser.UserPrincipalName -EmailAddresses (((Get-MailUser $ADUser.UserPrincipalName -DomainController:$DC).$Script:txtMailAddr.Text)) -CustomAttribute15 "EnableServiceAcct PS Date: $Date PS Time: $Time" -DomainController:$DC
                    Set-ADUser $script:ActName -Add @{ProxyAddresses=("SMTP:" + $Script:txtMailAddr.Text)}
                    Set-ADUser $script:ActName -Add @{'ExtensionAttribute15'="EnableMailUser PS Date: $Date PS Time: $Time"}
                    write-host "Adding Email Address" -ForegroundColor Green
    	            $LineToWrite = "INFO" + "`t" + "Setting Email Address to: " + $Script:txtMailAddr.Text
    	            WriteReportEvent
 		            $LineToWrite = "INFO" + "`t" + "Service Account has been Mail Enabled"
   		            WriteReportEvent
                }
            }

        #Set MailboxDisplayName
            write-host "Setting Mailbox Display Name" -ForegroundColor Green
            Set-ADUser $script:ActName -DisplayName $Script:txtDispName.Text
            $LineToWrite = "INFO" + "`t" + "Old Display Name: " + $ADUser.DisplayName
            WriteReportEvent
            $LineToWrite = "INFO" + "`t" + "New Display Name: " + $Script:txtDispName.Text
            WriteReportEvent

        #AssignLicense 
        # Set UsageLocation, Assign Exchagne P2 Email License and Update ExtensionAttribute5 for this account
        
            write-host "Setting License Usage Location" -ForegroundColor Green
            Set-MsolUser -UserPrincipalName $ADuser.UserPrincipalName -UsageLocation US -PasswordNeverExpires $true
            $SSKID = "ul:EXCHANGEENTERPRISE"
            $LicType = "ExchangeOnline Plan2"
            Write-Host "Assigning " $LicType "Licenses to" $Name -ForegroundColor Green
            Set-MsolUserLicense -UserPrincipalName $ADuser.UserPrincipalName -AddLicenses $SSKID
            $LineToWrite = "INFO" + "`t" + "Assigned " + $SSKID + " License to account"
            WriteReportEvent
            Set-ADUser $script:ActName -Replace @{ExtensionAttribute5 = "Y"}
            Set-ADUser $script:ActName -Replace @{ExtensionAttribute6 = $Script:txtTask.Text}
            $LineToWrite = "UPDA" + "`t" + "Set ExtensionAttribute5=Y and ExtensionAttribute6=" + $Script:txtTask.Text
            WriteReportEvent

            #If USR make sure PwdNeverExpires is set to False
            if (($ADuser.UserPrincipalName -like "USR*") -and ($ADUser.PasswordNeverExpires -eq $True))
            {
                write-host "Account to PasswordNeverExpires to False" -ForegroundColor Green
                Set-ADUser $script:ActName -PasswordNeverExpires $False
            }

        #Wait for Mailbox Creation to Complete
            write-host "Reconnecting to O365 in order to complete the Process" -ForegroundColor Yellow
            write-host "Cred File: " $Global:CredFile
           
            #$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
            #Import-PSSession $session -AllowClobber
            Connect-ExchangeOnline -Credential $Global:LiveCred

            write-host "Starting Mailbox Configuration Process for: " $ADUser.UserPrincipalName -ForegroundColor Cyan

            write-host "Waiting for mailbox creation to complete, this may take several minutes please be patient.." -ForegroundColor Yellow -NoNewline
            Do
            {
                write-host ".." -ForegroundColor Yellow -NoNewline
                start-sleep -Seconds 15
                $MBXExists = [bool](Get-Mailbox $ADUser.UserPrincipalName -ErrorAction SilentlyContinue)
            } while ($MBXExists -eq $False)

            write-host  "`nSetting Mailbox Quota and active 3 yr retention policy" -ForegroundColor Green
            $LineToWrite = "INFO" + "`t" + "Setting Mailbox Quota and active 3 yr Retention Policy"
            WriteReportEvent
            write-host "Setting mailbox retention policy" -ForegroundColor Green
            $LineToWrite = "INFO" + "`t" + "Setting Mailbox Retention Policy"
            WriteReportEvent
            Set-MailBox $ADUser.UserPrincipalName -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL MRM Policy - 3 yr Delete" -RetentionHoldEnabled $false
            write-host "Setting mailbox quotas" -ForegroundColor Green
            $LineToWrite = "INFO" + "`t" + "Setting Mailbox Quota"
            WriteReportEvent
            Set-Mailbox $ADUser.UserPrincipalName -UseDatabaseQuotaDefaults $False -IssueWarningQuota 45GB -ProhibitSendQuota 49.75GB -ProhibitSendReceiveQuota 50GB
            write-host "Setting mailbox protocols" -ForegroundColor Green
            $LineToWrite = "INFO" + "`t" + "Setting Mailbox Protocols"
            WriteReportEvent
            Set-CASMailBox $ADUser.UserPrincipalName -ImapEnabled $false -PopEnabled $false -ActiveSyncEnabled $false

        # Creating Security Groups and Applying Permissions
            write-host "Creating Mailbox Access Groups" -ForegroundColor Green
            If ($script:chkEditors.Checked -eq $true)
            {
                $GrpName = $Script:txtEDGrpName.Text
                $GrpAddr = $Script:txtEDGrpAddr.Text
                $GrpMem = $Script:txtEDGrpMbr.Text
                CreateSvcUSG
                Add-SvcFolderPermissions
            }
            If ($script:chkAuthors.Checked -eq $true)
            {
                $GrpName = $Script:txtAUGrpName.Text
                $GrpAddr = $Script:txtAUGrpAddr.Text
                $GrpMem = $Script:txtAUGrpMbr.Text
                CreateSvcUSG
                Add-SvcFolderPermissions
            }
            If ($script:chkReaders.Checked -eq $true)
            {
                $GrpName = $Sript:txtREGrpName.Text
                $GrpAddr = $Script:txtREGrpAddr.Text
                $GrpMem = $Script:txtREGrpMbr.Text
                CreateSvcUSG
                Add-SvcFolderPermissions
            }
            Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\LicensedServiceAccountMailbox.oft
        }

        If ($Script:chkKFIOFR.Checked -eq "Checked")
        {
            Add-KFIOFRMbxDetails
<#
                2 #Enable/Configure Oracle OFR/KFI Mailbox"
                {
                    Remove-PSSession (Get-PSSession)
                    Invoke-Expression -Command .\EnableSvcActAssignLicense.ps1
                    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
                    Import-PSSession $session -AllowClobber
#                   Create USG group to control access
                    Write-Host "Creating groups to control access to the mailbox.  Hit return when the NewUSG input file has ready" -ForegroundColor Red -NoNewline
                    $Ret = Read-Host
                    Invoke-Expression -Command .\NewUSG.ps1
#                   Add Mailbox/Folder Permissions and Apply Retention Policy
                    Write-Host "Applying folder permissions and retention policy to the mailbox.  Hit return when the AddFolderPermissionsinput file has ready and the mailbox has been provisioned." -ForegroundColor Red -NoNewline
                    $Ret = Read-Host
                    Invoke-Expression -Command .\AddFolderPermissions.ps1
                }
#>
        }

        If ($Script:chkSurHub.Checked -eq "Checked")
        {
            If ($Global:Result -eq "OK")
            {
                # Setup Folders and Files	
	            $LineToWrite = "STAR" + "`tNewSurfaceHub script has started"
	            WriteReportEvent
	            $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	            WriteReportEvent

                $DevName = $Script:txtActName.Text
                write-host "`nStarting Configuration for Surface Hub Device [$($DevName)]" -ForegroundColor Green
                If ($DevName -notlike "*@*")
                {
                    $DevName = $DevName + "@global.ul.com"
                }
                $GblEmAddr = $Script:txtMailAddr.Text.Replace("ul.com","global.ul.com")
                $LineToWrite = "INFO" + "`t" + "        Account Name: " + $DevName
                WriteReportEvent
                $LineToWrite = "INFO" + "`t" + "         Task Number: " + $Script:txtTask.Text
                WriteReportEvent
                $LineToWrite = "INFO" + "`t" + "        Display Name: " + $Script:txtDispName.Text
                WriteReportEvent
                $LineToWrite = "INFO" + "`t" + "       Email Address: " + $Script:txtMailAddr.Text
                WriteReportEvent
                $LineToWrite = "INFO" + "`t" + "Global Email Address: " + $GblEmAddr
                WriteReportEvent
                $LineToWrite = "INFO" + "`t" + "       Room Delegate: " + $Script:txtAcctOwner.Text + "`n"
                WriteReportEvent

#                Remove-PSSession (Get-PSSession)
                Add-PSSnapin *Exchange* -erroraction SilentlyContinue

                write-host "Enabling MailUser and Assigning Licenses for " $DevName -ForegroundColor Green
                $atUPN  = $DevName.indexOf("@")
                $MEnab = ([bool](Get-MailUser $DevName -ErrorAction SilentlyContinue))
  	            If ($MEnab -eq $false)
                {
                    Enable-MailUser $DevName -Alias $DevName.substring(0,$atUPN) -ExternalEmailaddress $Script:txtMailAddr.Text -PrimarySmtpAddress $Script:txtMailAddr.Text -DomainController:$DC
                    $LineToWrite = "INFO" + "`tMail Enabled Account $($DevName)"
                    WriteReportEvent
                }
                Start-Sleep -Seconds 15
#                Set-MailUser $DevName -EmailAddresses (((Get-MailUser $DevName -DomainController:$DC).EmailAddresses)+=($GblEmAddr)) -CustomAttribute15 "EnableMailUser PS Date: $Date PS Time: $Time" -DomainController:$DC
                Set-ADUser $Script:txtActName.Text -Add @{ProxyAddresses=("smtp:" + $GblEmAddr)}
  	            $LineToWrite = "INFO" + "`t" + "Setting Email Address to: " + $GblEmAddr
   	            WriteReportEvent
                Set-ADUser $Script:txtActName.Text -Add @{'ExtensionAttribute15'="EnableMailUser PS Date: $Date PS Time: $Time"}

                Set-MsolUser -UserPrincipalName $DevName -UsageLocation US -PasswordNeverExpires $true -ErrorAction Silentlycontinue
                Set-MsolUserLicense -UserPrincipalName $DevName -AddLicenses "ul:EXCHANGEENTERPRISE" -erroraction SilentlyContinue
                $LineToWrite = "INFO" + "`tTemporarily Assigned Exchange P2 license to Account"
                WriteReportEvent

                $AvailLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MEETING_ROOM"}
                If ($AvailLic.ActiveUnits -gt $AvailLic.ConsumedUnits)
                {
                    Set-MsolUserLicense -UserPrincipalName $DevName -AddLicenses "ul:MEETING_ROOM" -erroraction SilentlyContinue
                    $LineToWrite = "INFO" + "`tAssigned Meeting Room License to Account"
                    WriteReportEvent
                }
                else
                {
                    write-host "No available Microsoft Teams Meeting Rooms Licenses.  License must be procured and then assigned to this account" -ForegroundColor Red
                    $LineToWrite = "ERRO" + "`tNo Available Microsoft Teams Meeting Rooms Licenses.  License must be procured and the assigned this account."
                    WriteReportEvent
                }
                Write-host "Reconnecting to ExchangeOnline"
                Connect-ExchangeOnline -Credential $Global:LiveCred
#                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
#                Import-PSSession $session -AllowClobber
                Write-host "Connection Process complete"
                $easpolicy = Get-MobileDeviceMailboxPolicy "SurfaceHub Policy"
                $LDAPFilter = "(userPrincipalName=" + $DevName + ")"
                $ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties DisplayName
                Set-ADUser $ADuser -DisplayName $Script:txtDispName.Text

                $DevDelegate = ($Script:txtAcctOwner.Text -replace(" ","")) -split(",")

                #confirm that delegates are licensed mailboxes
                write-host "Checking that identified delegates are a security group or a licensed users" -ForegroundColor Cyan
                Foreach ($Dele in $DevDelegate)
                {            
                    $Exists = ""
                    If ($Dele -notlike "MBX.*")
                    {
                        $Exists = [bool](get-mailbox $Dele -ErrorAction SilentlyContinue)
                        If ($Exists -ne $True)
                        {
                            write-host "No mailbox found for delegate: $($Dele)" -ForegroundColor Red
                            $LineToWrite = "ERRO" + "`tNo mailbox found for identified delegate: " + $Dele
                            WriteReportEvent
                        }
                    }
                    else
                    {
                        $Exists = [bool](Get-DistributionGroup $Dele -ErrorAction SilentlyContinue)
                        If ($Exists -ne $True)
                        {
                            write-host "No security group found for delegate: $($Dele)" -ForegroundColor Red
                            $LineToWrite = "ERRO" + "`tNo seurity group found for identified delegate: " + $Dele
                            WriteReportEvent
                            $GrpName = $Script:txtAcctOwner.Text
                            $DGAlias = $Script:txtAcctOwner.Text
                            $GrpAddr = ($Script:txtAcctOwner.Text) + "@ul.com"
                            $DGManagedByMembers = "MBX.RRS.Owner@ul.com"
                            $GrpMem = read-host "Enter employee number of room delegates spearated by commas"

		                    # Create the USG
		                    New-DistributionGroup -Name $GrpName `
		                        -PrimarySmtpAddress $GrpAddr `
			                    -Alias $DGAlias `
			                    -ManagedBy $DGManagedByMembers `
			                    -RequireSenderAuthenticationEnabled $TRUE `
			                    -Type Security | Out-Null
					
                            $Owners = "Owners: " + ((Get-DistributionGroup $GrpName).ManagedBy -join (", "))
        
                            Set-Group -identity $GrpName `
		                        -Notes ($Owners + " - Per: " + $Script:txtTask.Text)

                            $AddMem = $GrpMem -split (",")
                            Foreach ($M IN $AddMem)
                            {
                                Add-DistributionGroupmember $GrpName -Member $M -BypassSecurityGroupManagerCheck
                                write-host "Added Member: " $M
                            }
                        }
                    }
                }

                #Make sure Mailbox has been created
                $MbxExists = [bool](get-mailbox $DevName -ErrorAction SilentlyContinue)
                write-host "Waiting for Mailbox Creation to Complete" -ForegroundColor Cyan -NoNewline
                Do {
                    write-host ".." -ForegroundColor Cyan -NoNewline
                    Start-Sleep -Seconds 30
                    $mbxExists = [bool](get-mailbox $DevName -ErrorAction SilentlyContinue)
                } while ($MbxExists -eq $False)
                $LineToWrite = "INFO" + "`tMailbox creation complete"
                WriteReportEvent

                write-host "`nChanging the mailbox type to a Room Mailbox..." -ForegroundColor Green -NoNewline
                set-mailbox $DevName -Type “Room”
                $mbxType = get-mailbox $Devname
                Do {
                    If ($mbxType.ResourceType -ne "Room")
                    {
                        write-host "..." -ForegroundColor Green -NoNewline
                        Start-Sleep -Seconds 10
                        $mbxType = get-mailbox $Devname
                    }
                } while ($mbxType.ResourceType -ne "Room")
                write-host "`nMailbox conversion complete" -ForegroundColor Green
                $LineToWrite = "INFO" + "`tChanged Mailbox Type to a Room Mailbox"
                WriteReportEvent

                write-host "Assigning the ActiveSync Policy for the SurfaceHubs" -ForegroundColor Green
                Set-CASMailbox $DevName -ActiveSyncMailboxPolicy $easPolicy.id
                $LineToWrite = "INFO" + "`tAssigned ActiveSync policy $($easPolicy.id) to account"
                WriteReportEvent

                write-host "Configuring the Booking Policy to the Standard UL Booking Policies" -ForegroundColor Green
                Set-CalendarProcessing $DevName -BookingWindowInDays 90 -MaximumDurationInMinutes 1440 -ConflictPercentageAllowed 20 -MaximumConflictInstances 3 -AutomateProcessing:AutoAccept
                Set-CalendarProcessing $DevName -RemovePrivateProperty $true -DeleteComments $false -DeleteSubject $false -AddOrganizerToSubject $true -AllRequestOutOfPolicy $true -AllBookInPolicy $true -AllRequestInPolicy $false -ResourceDelegates $DevDelegate -AddAdditionalResponse $true -AdditionalResponse “This is a Surface Hub room!”
                write-host "Granting Global RRS Admins permission to the mailbox" -ForegroundColor Green
                Add-MailboxPermission $DevName -AccessRights Fullaccess -User "dbs.crp.rrs.admins"
                $LineToWrite = "INFO" + "`tConfigured standard Room Booking policy to account"
                WriteReportEvent

                write-host "Enabling the CSMeetingRoom in the region this device will be deployed" -ForegroundColor Green
                $SIPAddr = "sip:" + $Script:txtMailAddr.Text
                $ADName = $DevName.Substring(0,$DevName.Indexof("@"))
                Set-ADUser $ADName -Replace @{'msRTCSIP-PrimaryUserAddress'=$sipAddr}
                Set-ADUser $ADName -Replace @{'msRTCSIP-DeploymentLocator'="sipfed.online.lync.com"}
                Set-ADUser $ADName -Replace @{'msRTCSIP-FederationEnabled'="TRUE"}
                Set-ADUser $ADName -Replace @{'msRTCSIP-InternetAccessEnabled'="TRUE"}
                Set-ADUser $ADName -Replace @{'msRTCSIP-UserEnabled'="TRUE"}
                $LineToWrite = "CONF" + "`tConfigured msRTCSIP-PrimaryUserAddress to $($sipAddr)"
                WriteReportEvent
                $LineToWrite = "CONF" + "`tConfigured msRTCSIP-DeploymentLocator to sipfed.online.lync.com"
                WriteReportEvent
                $LineToWrite = "ENAB" + "`tConfigured msRTCSIP-FederationEnabled to Enabled"
                WriteReportEvent
                $LineToWrite = "ENAB" + "`tConfigured msRTCSIP-InternetAccessEnabled to Enabled"
                WriteReportEvent  
                $LineToWrite = "ENAB" + "`tConfigured msRTCSIP-UserEnabled to Enabled"
                WriteReportEvent  

                Set-MsolUserLicense -UserPrincipalName $DevName -RemoveLicenses "ul:EXCHANGEENTERPRISE" -erroraction SilentlyContinue
                $LineToWrite = "REMO" + "`tRemoved Exchangee P2 license from account"
                WriteReportEvent  

                write-host "`nMoving AD Object to the O365 Licensed OU" -ForegroundColor Green
                $ADObj = Get-ADUser $ADName
                Move-ADObject -Identity $ADObj.ObjectGUID -TargetPath "OU=O365Licensed,OU=SVC,OU=ServiceAccounts,OU=Enterprise,DC=global,DC=ul,DC=com"
                $LineToWrite = "MOVE" + "`tMoved Account to Enterprise/SVC/O365Licensed Active Directory OU`n"
                WriteReportEvent 
                write-host "`nSurface Hub Device set-up complete" -ForegroundColor Red
                $LineToWrite = "MOVE" + "`tSurface Hub Device set-up Complete"
                WriteReportEvent 
            }
            Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\SurfaceHubDeviceConfigured.oft

            $Output = $wshell.Popup("Service account enablement procesess complete.",0,"Complete",0+32)
            write-host "Service Account Enablement Process complete.....`n" -ForegroundColor Green
        }
    }
    else
    {
        $Output = $wshell.Popup("Service account enablement procesess cancelled.",0,"Cancelled",0+32)
        write-host "Service Account enablement process cancelled.....`n" -ForegroundColor Red
    }
}
else
{
    $Output = $wshell.Popup("The Exchange Management Modules are not installed on this PC run this from a device with the Exchange Management Tools installed.",0,"No Exchange Modules",0+32)
}