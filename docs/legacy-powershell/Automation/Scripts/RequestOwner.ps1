<#
#
#  Called by:  DistributionSecurityGroupMenu.ps1
#
#  06/08/2023 - Started coding of process to Request Owners for a group or shared mailbox
#>

function Enter-ObjectName
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Enter Item Name" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(575,200) #(W,H)

    $TopLoc = 30

    $Script:lblDispName = New-Object System.Windows.Forms.Label   
        $Script:lblDispName.Text = "Enter Item Name:"
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
        $Script:ButGetDL.Size = new-Object System.Drawing.Size(150,20)
        $Script:ButGetDL.Text = "Get Item Details"
        $Script:ButGetDL.TabIndex = 0
        $Global:form.Controls.Add($Script:ButGetDL)
        $Script:ButGetDL.Add_Click({
            write-host "`nRetreiving Account Details:" $Script:txtDispName.Text -ForegroundColor Cyan
            $MbxExists = [bool]($Script:Mbx = Get-Mailbox $Script:txtDispName.Text -ErrorAction SilentlyContinue)
            If ($Script:txtDispName.Text -like "MBX*")
            {
                $FindMbx = $Script:txtDispName.Text.Split( ".")[1]
                $MbxExists = [bool]($Script:Mbx = Get-Mailbox $FindMbx -ErrorAction SilentlyContinue)
                If ($MbxExists -eq $True)
                {
                    $Script:txtDispName.Text = $FindMbx
                    write-host "This group is associated to a shared mailbox:  " $Script:txtDispName.Text
                }
            }
            $DLExists = [bool]($Script:Mbx = Get-DistributionGroup $Script:txtDispName.Text -ErrorAction SilentlyContinue)
            $UniExists = [bool]($Script:Mbx = Get-UnifiedGroup $Script:txtDispName.Text -ErrorAction SilentlyContinue)

            If ($MbxExists -eq $True)
            {
                write-host "Mailbox Found: " $Script:txtDispName.Text
                Add-FormStandardButtons
            }
            else
            {
                If ($DLExists -eq $True)
                {
                    write-host "Distribution List Found: " $Script:txtDispName.Text
                    Add-FormStandardButtons
                }
                else
                {
                    If ($UniExists -eq $True)
                    {
                        write-host "Unified Group Found: " $Script:txtDispName.Text
                        Add-FormStandardButtons
                    }
                    else
                    {
                        write-host "Unable to find this item: " $Script:txtDispName.Text
                    }
                }
            }
        })
}

#####
#functions needed to test without the full admin menu
Function Publish-Form
{
    If ($Global:InputFocus -eq $null)
    {
        $Global:InputFocus = $Global:okButton
    }
    ## Finalize Form and Show Dialog
    $Global:form.Add_Shown( { $form.Activate(); $Global:InputFocus.Focus() } )  #Activate and Set Focus 
#    $Global:form.Add_Shown( { $form.Activate(); $Global:okButton.Focus() } )  #Activate and Set Focus 
    $Global:Result = $Global:form.ShowDialog()          ## Show the form, and wait for the response
}

function Add-FormStandardButtons
{
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    $Global:buttonPanel = New-Object Windows.Forms.Panel  
    $buttonPanel.Size = New-Object Drawing.Size @(400,40) 
    $buttonPanel.Dock = "Bottom"    
    $Global:cancelButton = New-Object Windows.Forms.Button  
        $Global:cancelButton.Top = $buttonPanel.Height - $Global:cancelButton.Height - 10; $Global:cancelButton.Left = $buttonPanel.Width - $Global:cancelButton.Width - 10 
        $Global:cancelButton.TabIndex = 98
        $Global:cancelButton.Text = "Cancel" 
        $Global:cancelButton.DialogResult = "Cancel" 
        $Global:cancelButton.Anchor = "Right"
    ## Create the OK button, which will anchor to the left of Cancel 
    $Global:okButton = New-Object Windows.Forms.Button   
        $Global:okButton.Top = $cancelButton.Top ; $Global:okButton.Left = $cancelButton.Left - $Global:okButton.Width - 10
        $Global:okButton.TabIndex = 97
        $Global:okButton.Text = $Action
        If ($Global:OKDetails -ne "")
        {
            $Global:okButton.Text = $Global:OKDetails
            $Global:OKDetails = ""
        }
        else
        {
            $Global:okButton.Text = "Continue"
        }
        $Global:okButton.DialogResult = "OK" 
        $Global:okButton.Anchor = "Right"
    ## Add the buttons to the button panel 
    $Global:buttonPanel.Controls.Add($Global:okButton) 
    $Global:buttonPanel.Controls.Add($Global:cancelButton) 
    ## Add the button panel to the form 
    $Global:form.Controls.Add($buttonPanel)
    ## Set Default actions for the buttons 
    $Global:form.AcceptButton = $Global:okButton          # ENTER = ok 
    $Global:form.CancelButton = $Global:cancelButton      # ESCAPE = Cancel
}

#####

function Build-DLInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Rename Distribution Group" 
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
        $Global:txtInpName.TabIndex = 0 # set Tab Order 
        $Global:txtInpName.Top = 10; $Global:txtInpName.Left = 130; $Global:txtInpName.Width = 280;  
        $Global:txtInpName.Text = ""   # DisplayName
        $Global:txtInpName.TabIndex = 0
        $Global:form.Controls.Add($Global:txtInpName)    # Add to Form

    ## New Display Name
    $Global:lblNewDispName = New-Object System.Windows.Forms.Label   
        $Global:lblNewDispName.Text = "New DL Name:"
        $Global:lblNewDispName.Top = 40 ; $Global:lblNewDispName.Left = 10; $Global:lblNewDispName.Width=120 ;$Global:lblNewDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblNewDispName)    # Add to Form 
        # 
        $Global:txtInpNewName = New-Object Windows.Forms.TextBox  
        $Global:txtInpNewName.TabIndex = 0 # set Tab Order 
        $Global:txtInpNewName.Top = 40; $Global:txtInpNewName.Left = 130; $Global:txtInpNewName.Width = 280;  
        $Global:txtInpNewName.Text = ""   # DisplayName
        $Global:txtInpNewName.TabIndex = 0
        $Global:form.Controls.Add($Global:txtInpNewName)    # Add to Form 
 
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 70 ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtInpTaskNo.Top = 70; $Global:txtInpTaskNo.Left = 130; $Global:txtInpTaskNo.Width = 120;  
        $Global:txtInpTaskNo.Text = "TASK"   # Enter ticket number
        $Global:txtInpTaskNo.TabIndex = 1
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form

        Add-FormStandardButtons
}

function Build-DLDetailsForm
{
    $BldDetails = "N"
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Rename Distribution Group Details" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(750,450) #(W,H)

    ## Get Details used to create
    ## Display Name
    $LocTop = 10
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "DL Name:"
        $Global:lblDispName.Top = $LocTop ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtDLDet = New-Object Windows.Forms.TextBox  
        $Global:txtDLDet.TabIndex = 0 # set Tab Order
        $Global:txtDLDet.Top = $LocTop; $Global:txtDLDet.Left = 130; $Global:txtDLDet.Width = 400;  
        $Global:txtDLDet.Text = $Global:txtInpName.Text
        $Global:txtDLDet.TabIndex = 0
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
        $Global:txtDetTaskNo.TabIndex = 1
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
        $Global:txtInpDLNotes.Text = "(Not Configurable on This Group Type)"
        If ($DSTGrpDetails.RecipientTypeDetails -ne "RoomList")
        {
            $Global:txtInpDLNotes.Text = (get-group $Global:txtInpName.Text).Notes
        }
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
        $Global:txtInpDLOwner.TabIndex = 2
        $Global:txtInpDLOwner.Text = ((Get-DistributionGroup $Global:txtDLDet.Text).ManagedBy -join ", ")
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
        $Global:txtInpDLMemCnt.TabIndex = 3
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
        $Global:txtInpDLMem.TabIndex = 4
        $Global:txtInpDLMem.Text = (Get-DistributionGroupMember $Global:txtInpName.Text).DisplayName -join "; "
        $Global:form.Controls.Add($Global:txtInpDLMem)    # Add to Form

    ## New Display Name
    $LocTop = $LocTop + 140
    $Global:lblNewDispName = New-Object System.Windows.Forms.Label   
        $Global:lblNewDispName.Text = "New DL Name:"
        $Global:lblNewDispName.Top = $LocTop ; $Global:lblNewDispName.Left = 10; $Global:lblNewDispName.Width=120 ;$Global:lblNewDispName.AutoSize = $true
        $Global:form.Controls.Add($Global:lblNewDispName)    # Add to Form 
        # 
        $Global:txtNewDLDet = New-Object Windows.Forms.TextBox  
        $Global:txtNewDLDet.TabIndex = 0 # set Tab Order
        $Global:txtNewDLDet.Top = $LocTop; $Global:txtNewDLDet.Left = 130; $Global:txtNewDLDet.Width = 400;  
        $Global:txtNewDLDet.Text = $Global:txtInpNewName.Text
        $Global:txtNewDLDet.TabIndex = 0
        $Global:form.Controls.Add($Global:txtNewDLDet)    # Add to Form 

    ## New Address
    $LocTop = $LocTop + 30
    $Global:lblDLAddr = New-Object System.Windows.Forms.Label   
        $Global:lblDLAddr.Text = "Email Address:"
        $Global:lblDLAddr.Top = $LocTop ; $Global:lblDLAddr.Left = 10; $Global:lblDLAddr.Width=120 ;$Global:lblDLAddr.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLAddr)    # Add to Form 
        # 
        $Global:txtDLAddr = New-Object Windows.Forms.TextBox  
        $Global:txtDLAddr.TabIndex = 0 # set Tab Order
        $Global:txtDLAddr.Top = $LocTop; $Global:txtDLAddr.Left = 130; $Global:txtDLAddr.Width = 400;
#        $Global:DLAddr = ((Get-Culture).textinfo.totitlecase($Global:txtInpNewName.Text))
        $Global:txtDLAddr.Text = ($Global:txtInpNewName.Text -Replace '[ &/,$#_-]','') + "@ul.com"
        $Global:txtDLAddr.TabIndex = 0
        $Global:form.Controls.Add($Global:txtDLAddr)    # Add to Form 

    ## Authorized Requestor
    $LocTop = $LocTop + 30
    $Global:chkAuthorized = New-Object Windows.Forms.checkbox 
        $Global:chkAuthorized.Left = 130; $Global:chkAuthorized.Width = 200; $Global:chkAuthorized.Top = $LocTop  
        $Global:chkAuthorized.Text = "Authorized Requestor" 
        $Global:chkAuthorized.Checked = $false   # set a default value 
        $Global:chkAuthorized.TabIndex = 5
        $Global:form.Controls.Add($Global:chkAuthorized) 
        # Obtain Value with: $Global:chkAuthorized.Checked

        Add-FormStandardButtons
}

$Global:GrpName = $Global:txtDispName.Text
$Global:EMailAddr = $Global:txtGrpAddr.Text

Enter-ObjectName
Publish-Form

pause


$DstExists = [bool](Get-DistributionGroup $Global:txtInpName.Text -ErrorAction SilentlyContinue)
if ($DstExists -eq "True")
{

    $DSTGrpDetails = Get-DistributionGroup $Global:txtInpName.Text
#    $DSTGrpDetails = Get-DistributionGroup $Global:txtDLDet.Text
    Build-DLDetailsForm
    Publish-Form

    If ($Global:Result -eq "OK")
    {
        If ($Global:chkAuthorized.Checked -eq "Checked")
        {
            If ($DSTGrpDetails.RecipientTypeDetails -ne "RoomList")
            {
                $GrpInf = get-Group $Global:txtDLDet.text
            }

            If ($GrpInf.Notes -like "*Per*")
            {
                $GrpInf.Notes = $GrpInf.Notes.Remove($GrpInf.Notes.IndexOf("Per")) + "Per " + $Global:txtDetTaskNo.Text
            }
            else
            {
                If ($Global:txtDLDet.Text -notlike "*outOfPolicy*")
                {
                    $LSTTag = "Owners: "
                }
                else
                {
                    $LSTTag = "Room Delegates: "
                }
                 
                $GrpInf.Notes = $LSTTag + ((Get-DistributionGroup $Global:txtDLDet.Text).ManagedBy -join (", ")) + " - Per " + $Global:txtDetTaskNo.Text
                $MailTip = $LSTTag + ((Get-DistributionGroup $Global:txtDLDet.Text).ManagedBy -join (", "))
            }

            If ($DSTGrpDetails.RecipientTypeDetails -ne "RoomList")
            {
                Set-Group $Global:txtDLDet.Text -Notes $GrpInf.Notes
            }
            write-host "Resetting distribution group values"
            Set-DistributionGroup $Global:txtDLDet.Text -DisplayName $Global:txtNewDLDet.Text -Name $Global:txtNewDLDet.Text -Alias ($Global:txtNewDLDet.Text -Replace '[ &/,$#_-]','') -EmailAddresses @{add=$Global:txtDLAddr.Text}
            write-host "Resetting distribution group primary email address and mailtip value"
            Set-DistributionGroup $Global:txtNewDLDet.Text -PrimarySmtpAddress $Global:txtDLAddr.Text -MailTip $MailTip
        }
        else
        {
            $Output = $wshell.Popup("Obtain approval from authorized requestor - No Changes Made",0,"Not Authorized",0+32)
        }
        Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\DLRenamed.oft
    }
    else
    {
        $Output = $wshell.Popup("Distribution Group rename cancelled.",0,"Request Cancelled",0+32)
    }
}
else
{
    $Output = $wshell.Popup("Distribution List Does Not Exist - No Changes Made",0,"Not Found",0+32)
}
