# This prints details regarding mailbox access to the console and report file

Function Add-ViewFormStandardButtons
{
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    $Global:ViewButtonPanel = New-Object Windows.Forms.Panel  
    $ViewButtonPanel.Size = New-Object Drawing.Size @(400,40) 
    $ViewButtonPanel.Dock = "Bottom"    
    $Global:ViewCancelButton = New-Object Windows.Forms.Button  
        $Global:ViewCancelButton.Top = $ViewButtonPanel.Height - $Global:ViewCancelButton.Height - 10; $Global:ViewCancelButton.Left = $ViewButtonPanel.Width - $Global:ViewCancelButton.Width - 10 
        $Global:ViewCancelButton.TabIndex = 99
        $Global:ViewCancelButton.Text = "Cancel" 
        $Global:ViewCancelButton.DialogResult = "Cancel" 
        $Global:ViewCancelButton.Anchor = "Right"
    ## Create the OK button, which will anchor to the left of Cancel 
    $Global:ViewokButton = New-Object Windows.Forms.Button
        $Global:ViewokButton.Top = $ViewCancelButton.Top ;$Global:ViewokButton.Left = $ViewCancelButton.Left - $Global:ViewokButton.Width - 10
        $Global:ViewokButton.Text = "Next"
        $Global:ViewokButton.DialogResult = "OK" 
        $Global:ViewokButton.Anchor = "Right"
    ## Add the buttons to the button panel 
    $Global:ViewButtonPanel.Controls.Add($Global:ViewokButton) 
    $Global:ViewButtonPanel.Controls.Add($Global:ViewCancelButton) 
    ## Add the button panel to the form 
    $Global:ViewForm.Controls.Add($ViewButtonPanel)
    ## Set Default actions for the buttons 
    $Global:Viewform.AcceptButton = $Global:ViewokButton          # ENTER = ok 
    $Global:Viewform.CancelButton = $Global:ViewCancelButton      # ESCAPE = Cancel
}

Function Add-ViewUsrDeatilsForm
{
    $Global:ViewForm = New-Object Windows.Forms.Form 
    $Global:ViewForm.FormBorderStyle = "FixedToolWindow" 
    $Global:ViewForm.Text = "User Hold Details" 
    $Global:ViewForm.StartPosition = "CenterScreen" 
    $Global:ViewForm.Width = 670 ; $Viewform.Height = 400   # Make the form wider

    $VTop = 20
    $VLblLeft =15
    $VTxtLeft = 160
    $VLblLeft2 = 300
    $VTxtLeft2 = 410

    ## Employee Number
    $Script:lblVEno = New-Object System.Windows.Forms.Label   
        $Script:lblVEno.Text = "Employee Number:"  
        $Script:lblVEno.Top = $VTop ; $Script:lblVEno.Left = $VLblLeft; $Script:lblVEno.Width=150 ;$Script:lblVEno.AutoSize = $true
        $Script:lblVEno.Visible = $True 
        $ViewForm.Controls.Add($Script:lblVEno)    # Add to Form 
        # 
        $Script:txtVENo = New-Object Windows.Forms.TextBox
        $Script:txtVENo.ReadOnly = $True;
        $Script:txtVENo.Top = $VTop; $Script:txtVENo.Left = $VTxtLeft; $Script:txtVENo.Width = 120;
        $Script:txtVENo.Visible = $True
        $Global:Viewform.Controls.Add($Script:txtVENo)    # Add to Form 

    $VTop = $VTop + 30
    ## Employee Name
    $Script:lblName = New-Object System.Windows.Forms.Label   
        $Script:lblName.Text = "Employee Name:"  
        $Script:lblName.Top = $VTop ; $Script:lblName.Left = $VLblLeft; $Script:lblName.Width=150 ;$Script:lblName.AutoSize = $true
        $Script:lblName.Visible = $True 
        $ViewForm.Controls.Add($Script:lblName)    # Add to Form 
        # 
        $Script:txtName = New-Object Windows.Forms.TextBox
        $Script:txtName.ReadOnly = $True;
        $Script:txtName.Top = $VTop; $Script:txtName.Left = $VTxtLeft; $Script:txtName.Width = 120;
        $Script:txtName.Visible = $True
        $Global:Viewform.Controls.Add($Script:txtName)    # Add to Form 

    ## Employee Type
    $Script:lblType = New-Object System.Windows.Forms.Label   
        $Script:lblType.Text = "Employee Type:"  
        $Script:lblType.Top = $VTop ; $Script:lblType.Left = $VLblLeft2; $Script:lblType.Width=150 ;$Script:lblType.AutoSize = $true 
        $Script:lblType.Visible = $True 
        $ViewForm.Controls.Add($Script:lblType)    # Add to Form 
        # 
        $Script:txtType = New-Object Windows.Forms.TextBox
        $Script:txtType.ReadOnly = $True;
        $Script:txtType.Top = $VTop; $Script:txtType.Left = $VTxtLeft2; $Script:txtType.Width = 120;
        $Script:txtType.Visible = $True
        $Global:ViewForm.Controls.Add($Script:txtType)    # Add to Form 

    $VTop = $VTop + 30
    ## ADContainer
    $Script:lblContainer = New-Object System.Windows.Forms.Label   
        $Script:lblContainer.Text = "AD Container:"  
        $Script:lblContainer.Top = $VTop ; $Script:lblContainer.Left = $VLblLeft; $Script:lblContainer.Width=150 ;$Script:lblContainer.AutoSize = $true 
        $ViewForm.Controls.Add($Script:lblContainer)    # Add to Form 
        # 
        $Script:txtContainer = New-Object Windows.Forms.TextBox  
        $Script:txtContainer.ReadOnly = $True; 
        $Script:txtContainer.Top = $VTop; $Script:txtContainer.Left = $VTXTLeft; $Script:txtContainer.Width = 370;
        $Global:ViewForm.Controls.Add($Script:txtContainer)    # Add to Form 

    $VTop = $VTop + 30
    ## Hold Enabled Date
    $Script:lblDate = New-Object System.Windows.Forms.Label   
        $Script:lblDate.Text = "Hold Enabled Date:"  
        $Script:lblDate.Top = $VTop ; $Script:lblDate.Left = $VLblLeft; $Script:lblDate.Width=150 ;$Script:lblDate.AutoSize = $true 
        $ViewForm.Controls.Add($Script:lblDate)    # Add to Form 
        # 
        $Script:txtDate = New-Object Windows.Forms.TextBox  
        $Script:txtDate.ReadOnly = $True;
        $Script:txtDate.Top = $VTop; $Script:txtDate.Left = $VTxtLeft; $Script:txtDate.Width = 120;  
        $Script:txtDate.Text = ""
        $Global:ViewForm.Controls.Add($Script:txtDate)    # Add to Form 

    ## Hold Enabled By
    $Script:lblBy = New-Object System.Windows.Forms.Label   
        $Script:lblBy.Text = "Hold Enabled By:"  
        $Script:lblBy.Top = $VTop; $Script:lblBy.Left = $VLblLeft2; $Script:lblBy.Width=100 ;$Script:lblBy.AutoSize = $true  
        $ViewForm.Controls.Add($Script:lblBy)    # Add to Form 
        # 
        $Script:txtBy = New-Object Windows.Forms.TextBox  
        $Script:txtBy.ReadOnly = $True; 
        $Script:txtBy.Top = $VTop; $Script:txtBy.Left = $VTxtLeft2; $Script:txtBy.Width = 120;  
        $Global:ViewForm.Controls.Add($Script:txtBy)    # Add to Form 

    $VTop = $VTop + 30
    ## Hold ADOjbect Protected
    $Script:lblProtect = New-Object System.Windows.Forms.Label   
        $Script:lblProtect.Text = "AD Object Protected:"  
        $Script:lblProtect.Top = $VTop ; $Script:lblProtect.Left = $VLblLeft; $Script:lblProtect.Width=150 ;$Script:lblProtect.AutoSize = $true 
        $ViewForm.Controls.Add($Script:lblProtect)    # Add to Form 
        # 
        $Script:txtProtect = New-Object Windows.Forms.TextBox  
        $Script:txtProtect.ReadOnly = $True; 
        $Script:txtProtect.Top = $VTop; $Script:txtProtect.Left = $VTxtLeft; $Script:txtProtect.Width = 120;
        $Global:ViewForm.Controls.Add($Script:txtProtect)    # Add to Form 

    ## Hold ADOjbect Hidden
    $Script:lblHide = New-Object System.Windows.Forms.Label   
        $Script:lblHide.Text = "AD Object Hidden:"  
        $Script:lblHide.Top = $VTop ; $Script:lblHide.Left = $VLblLeft2; $Script:lblHide.Width=160 ;$Script:lblHide.AutoSize = $true 
        $ViewForm.Controls.Add($Script:lblHide)    # Add to Form 
        # 
        $Script:txtHide = New-Object Windows.Forms.TextBox  
        $Script:txtHide.ReadOnly = $True;
        $Script:txtHide.Top = $VTop; $Script:txtHide.Left = $VTxtLeft2; $Script:txtHide.Width = 120;
        $Global:ViewForm.Controls.Add($Script:txtHide)    # Add to Form 

    $VTop = $VTop + 30
    ## Accounts Granted Access
    $Script:lblAccess = New-Object System.Windows.Forms.Label   
        $Script:lblAccess.Text = "Accounts with Access:"  
        $Script:lblAccess.Top = $VTop ; $Script:lblAccess.Left = $VLblLeft; $Script:lblAccess.Width=160 ;$Script:lblAccess.AutoSize = $true 
        $ViewForm.Controls.Add($Script:lblAccess)    # Add to Form 
        # 
        $Script:txtAccess = New-Object Windows.Forms.TextBox  
        $Script:txtAccess.ReadOnly = $True; 
        $Script:txtAccess.Top = $VTop; $Script:txtAccess.Left = $VTxtLeft; $Script:txtAccess.Width = 370; 
        $Script:txtAccess.Text = ""
        $Global:ViewForm.Controls.Add($Script:txtAccess)    # Add to Form 
 
     $VTop = $VTop + 30
     ##InPlace Holds
     $Script:lblInPlace = New-Object System.Windows.Forms.Label   
        $Script:lblInPlace.Text = "InPlace Holds:"  
        $Script:lblInPlace.Top = $VTop ; $Script:lblInPlace.Left = $VLblLeft; $Script:lblInPlace.Width=160 ;$Script:lblInPlace.AutoSize = $true 
        $ViewForm.Controls.Add($Script:lblInPlace)    # Add to Form 
        # 
        $Script:txtInPlace = New-Object Windows.Forms.TextBox  
        $Script:txtInPlace.ReadOnly = $True;
        $Script:txtInPlace.Text = "(None)"
        $Script:txtInPlace.Top = $VTop; $Script:txtInPlace.Left = $VTxtLeft; $Script:txtInPlace.Width = 370; 
        $Global:Viewform.Controls.Add($Script:txtInPlace)    # Add to Form           
 
    $VTop = $VTop + 30
    ## ListBox - Fill with Data From Azure Location Name 
    $Script:lblLoc = New-Object System.Windows.Forms.Label   
        $Script:lblLoc.Text = "Preservation Comment(s):"; $Script:lblLoc.Top = $VTop; $Script:lblLoc.Left = $VLblLeft; $Script:lblLoc.Autosize = $true  
        $Global:ViewForm.Controls.Add($Script:lblLoc)  
        # Listbox for Location Name 
        $Script:locListBox = New-Object System.Windows.Forms.ListBox  
        $Script:locListBox.Top = $VTop; $locListBox.Left = $VTxtLeft; $locListBox.Height = 100; $LocListBox.Width = 370;
        $script:locListBox.TabIndex = 1
        $Global:Viewform.Controls.Add($Script:locListBox) #Add listbox to form 

    Add-ViewFormStandardButtons
}

Function Build-HoldMenu
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Legal Hold Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 670 ; $form.Height = 400   # Make the form wider

    $Top = 30
    $Col2Top = 40
    $Left = 100
    $Col2Left = 340
    $Col1Left = 130
  
    ## New Hold
    $Script:chkNewHold = New-Object System.Windows.Forms.RadioButton
        $Script:chkNewHold.Text = "Enable for New Legal Hold" 
        $Script:chkNewHold.Top = $Top ; $Script:chkNewHold.Left = $Left; $Script:chkNewHold.Width=150 ;$Script:chkNewHold.AutoSize = $true 
        $form.Controls.Add($Script:chkNewHold)    # Add to Form
        $Script:chkNewHold.Add_Click({
            Hide-All
            Uncheck-Options
            $Script:chk40Day.Visible = $True
            $Script:chk180Day.Visible = $True
            $Script:chkStd.Visible = $True
            $Script:lblAttorney.Text = "Attorney Assigned:"
            $Script:txtHRRep.Text = ""
            $Script:lblENos.Text = "Staff to Hold:"
            $Script:txtRelENos.Left = 130; $Script:txtRelENos.Width = 340;
            $Script:txtRelDBOwner.Left = 130; $Script:txtRelDBOwner.Width = 340;
            $Global:okButton.Text = "Enable"
        })

    ## 40 Day Hold
    $Script:chk40Day = New-Object System.Windows.Forms.Checkbox
        $Script:chk40Day.Text = "40 Day Hold"  
        $Script:chk40Day.Top = $Top ; $Script:chk40Day.Left = $Col2Left; $Script:chk40Day.Width=150 ;$Script:chk40Day.AutoSize = $true 
        $Script:chk40Day.Visible = $False
        $form.Controls.Add($Script:chk40Day)    # Add to Form
        $Script:chk40Day.Add_Click({
            Hide-Details
            $Script:chk180Day.Checked = $False
            $Script:chkStd.Checked = $False
            $Script:ButGetENo.visible = $False
            If ($Script:chkNewHold.checked -eq $True)
            {
                $Script:txtHoldName.Text = "40 Day Hold (PutNameHere) - "
                If ($Script:chkSilent.checked -eq $True)
                {
                    $Script:txtHoldName.Text = "NonDisclosed - 40 Day Hold (PutNameHere) - "
                }
                $Script:lblHoldName.Visible = $True
                $Script:txtHoldName.Visible = $True
                $Script:lblENos.Visible = $True
                $Script:txtENos.Visible = $False
                $Script:txtRelENos.Visible = $True
                $Script:lblDBOwner.Visible = $True
                $Script:txtDBOwner.Visible = $False
                $Script:txtRelDBOwner.Visible = $True
                $Script:txtAttorney.Text = "B.Bogaerts"
            }
            If ($Script:chkExistHold.checked -eq $True)
            {
                ViewOpt-ExistingHold
            }
            Show-WhoSilent
            $Script:txtActHolds.Items.Clear()
            Foreach ($Hold in $Script:40DayHolds)
            {
                [void] $Script:txtActHolds.Items.Add($Hold.HoldName)
            }
        })

    $Top = $Top + 20
    ## Existng Hold
    $Script:chkExistHold = New-Object System.Windows.Forms.RadioButton
        $Script:chkExistHold.Text = "Enable Additional Staff for Existing Hold"  
        $Script:chkExistHold.Top = $Top; $Script:chkExistHold.Left = $Left; $Script:chkExistHold.Width=150 ;$Script:chkExistHold.AutoSize = $true 
        $form.Controls.Add($Script:chkExistHold)    # Add to Form
        $Script:chkExistHold.Add_Click({
            Hide-All
            Uncheck-Options
            $Script:chk40Day.Visible = $True
            $Script:chk180Day.Visible = $True
            $Script:chkStd.Visible = $True
            $Script:lblENos.Text = "Current Staff on Hold:"
            $Script:lblAttorney.Text = "Attorney Assigned:"
            $Script:txtRelENos.Left = 330; $Script:txtRelENos.Width = 165
            $Script:txtRelDBOwner.Left = 330; $Script:txtRelDBOwner.Width = 165
            $Script:lblRelENos.Text = ":Staff to Add"
            $Script:lblRelDBOwner.Text = ":DBOwner to Add"
            $Global:okButton.Text = "Enable"
        })

    ## 180 Day
    $Script:chk180Day = New-Object System.Windows.Forms.Checkbox
        $Script:chk180Day.Text = "180 Day Hold"  
        $Script:chk180Day.Top = $Top ; $Script:chk180Day.Left = $Col2Left; $Script:chk180Day.Width=150 ;$Script:chk180Day.AutoSize = $true 
        $Script:chk180Day.Visible = $False
        $form.Controls.Add($Script:chk180Day)    # Add to Form
        $Script:chk180Day.Add_Click({
            Hide-Details
            $Script:chk40Day.Checked = $False
            $Script:chkStd.Checked = $False
            $Script:ButGetENo.visible = $False
            If ($Script:chkNewHold.checked -eq $True)
            {
                $Script:txtHoldName.Text = "180 Day Hold (PutNameHere) - "
                If ($Script:chkSilent.checked -eq $True)
                {
                    $Script:txtHoldName.Text = "NonDisclosed - 180 Day Hold (PutNameHere) - "
                }
                $Script:lblHoldName.Visible = $True
                $Script:txtHoldName.Visible = $True
                $Script:lblENos.Visible = $True
                $Script:txtENos.Visible = $False
                $Script:txtRelENos.Visible = $True
                $Script:lblDBOwner.Visible = $True
                $Script:txtDBOwner.Visible = $False
                $Script:txtRelDBOwner.Visible = $true
                $Script:txtAttorney.Text = "B.Bogearts"
            }
            If ($Script:chkExistHold.checked -eq $True)
            {
                ViewOpt-ExistingHold
            }
            Show-WhoSilent
            $Script:txtActHolds.Items.Clear()
            Foreach ($Hold in $Script:180DayHolds)
            {
                [void] $Script:txtActHolds.Items.Add($Hold.HoldName)
            }
        })
        
    $Top = $Top + 20
    ## Release Standard Hold
    $Script:chkStdRel = New-Object System.Windows.Forms.RadioButton
        $Script:chkStdRel.Text = "Release from Legal Hold" 
        $Script:chkStdRel.Top = $Top ; $Script:chkStdRel.Left = $Left; $Script:chkStdRel.Width=150 ;$Script:chkStdRel.AutoSize = $true 
        $form.Controls.Add($Script:chkStdRel)    # Add to Form
        $Script:chkStdRel.Add_Click({
            Hide-All
            Uncheck-Options
            $Script:chkStdHold.Visible = $True
            $Script:chkfor40180.Visible = $True
            $Script:lblAttorney.Text = "Released by:"
            $Script:lblENos.Text = "Staff on Hold:"
            $Script:txtRelENos.Left = 330; $Script:txtRelENos.Width = 165
            $Script:txtRelDBOwner.Left = 330; $Script:txtRelDBOwner.Width = 165
            $Script:lblRelENos.Text = ":Staff to Release"
            $Script:lblRelDBOwner.Text = ":DBOwner to Release"
            $Global:okButton.Text = "Release"
        })

    ## Standard Hold
    $Script:chkStd = New-Object System.Windows.Forms.Checkbox
        $Script:chkStd.Text = "Standard Hold"  
        $Script:chkStd.Top = $Top ; $Script:chkStd.Left = $Col2Left; $Script:chkStd.Width=150 ;$Script:chkStd.AutoSize = $true 
        $Script:chkStd.Visible = $False
        $form.Controls.Add($Script:chkStd)    # Add to Form
        $Script:chkStd.Add_Click({
            Hide-Details
            $Script:chk40Day.Checked = $False
            $Script:chk180Day.Checked = $False
            $Script:ButGetENo.visible = $False
            If ($Script:chkNewHold.Checked -eq $True)
            {
                $Script:txtHoldName.Text = "(PutNameHere) Legal Hold - "
                If ($Script:chkSilent.checked -eq $True)
                {
                    $Script:txtHoldName.Text = "NonDisclosed - (PutNameHere) Legal Hold - "
                }                
                $Script:txtAttorney.Text = "B.Bogaerts"
                $Script:lblHoldName.Visible = $True
                $Script:txtHoldName.Visible = $True
                $Script:lblENos.Visible = $True
                $Script:txtENos.Visible = $False
                $Script:txtRelENos.Visible = $True
                $Script:lblDBOwner.Visible = $True
                $Script:txtDBOwner.Visible = $False
                $Script:txtRelDBOwner.Visible = $True
            }
            If ($Script:chkExistHold.checked -eq $True)
            {
                ViewOpt-ExistingHold
            }
            Show-WhoSilent
            $Script:txtActHolds.Items.Clear()
            Foreach ($Hold in $Script:StdHolds)
            {
                [void] $Script:txtActHolds.Items.Add($Hold.HoldName)
            }
        })

    $Top = $Top + 20
    ## Rename an existing hold
    $Script:chkRenameHold = New-Object System.Windows.Forms.RadioButton
        $Script:chkRenameHold.Text = "Rename Existing Hold (developing)" 
        $Script:chkRenameHold.Top = $Top ; $Script:chkRenameHold.Left = $Left; $Script:chkRenameHold.Width=150 ;$Script:chkRenameHold.AutoSize = $true 
        $form.Controls.Add($Script:chkRenameHold)    # Add to Form
        $Script:chkRenameHold.Add_Click({
            Hide-All
            Uncheck-Options
            $Script:chkStdHold.Visible = $True
            $Script:chkfor40180.Visible = $True
            $Script:lblENos.Text = "Staff on Hold:"
            $Global:okButton.Text = "Rename"
        })

    $Top = $Top + 20
    ## View Holds for a User
    $Script:chkViewUsr = New-Object System.Windows.Forms.RadioButton
        $Script:chkViewUsr.Text = "View Details for a Specific User(s)" 
        $Script:chkViewUsr.Top = $Top ; $Script:chkViewUsr.Left = $Left; $Script:chkViewUsr.Width=150 ;$Script:chkViewUsr.AutoSize = $true 
        $form.Controls.Add($Script:chkViewUsr)    # Add to Form
        $Script:chkViewUsr.Add_Click({
            Hide-All
            Uncheck-Options
            $Script:txtENo.Visible = $True
            $Script:ButViewENo.Visible = $True
        })

    ## Employee Number
    $Script:txtENo = New-Object Windows.Forms.TextBox
        $Script:txtENo.Top = $Top; $Script:txtENo.Left = ($Col2Left-40); $Script:txtENo.Width = 225;
        $Script:txtENo.Text = "Enter comma separated ENo's"
        $Script:txtENo.Visible = $False
        $form.Controls.Add($Script:txtENo)    # Add to Form
        $Script:txtENo.Add_Click({
            $Script:txtENo.Text = ""
        })

    $Script:GettingData = New-Object Windows.Forms.Label
        $Script:GettingData.Top = ($Top+40); $Script:GettingData.Left = ($Col2Left-100); $Script:GettingData.Width = 425;
        $Script:GettingData.Text = ""
        $Script:GettingData.ForeColor = "Red"
        $Script:GettingData.Visible = $False
        $form.Controls.Add($Script:GettingData)    # Add to Form

    $Script:ButViewENo = New-Object Windows.Forms.Button
        $Script:ButViewENo.Location = New-object System.Drawing.Size(530,($Top-5))
        $Script:ButViewENo.Size = new-Object System.Drawing.Size(100,30)
        $Script:ButViewENo.Text = "Get Details"
        $Script:ButViewENo.TabIndex = 0
        $Script:ButViewENo.Visible = $False
        $form.Controls.Add($Script:ButViewENo)
        $Script:ButViewENo.Add_Click({
            $mbxs = $Script:txtENo.Text -split(",")
            Hide-All
            $Global:cancelButton.Visible = $False
            Foreach ($m in $mbxs)
            {
                $Script:GettingData.Visible = $True
                If ($m -notlike "Enter*")
                {
                    $Script:GettingData.Text = "Getting User Details for: " + $m
                    $Script:GettingData.Refresh()
                    $Exists = [bool]($vmbx = get-mailbox $m -ErrorAction SilentlyContinue)
                    $VADUsr = get-aduser $m -Properties ExtensionAttribute14
                    If ($Exists -eq $True)
                    {
                        $ENo = $m
                        PrtMBxAccess($Script:MbxAccess)
                        $Script:txtVENo.Text = $vmbx.Alias
                        $Script:txtName.Text = $vmbx.DisplayName
                        $Script:txtType.Text = $vmbx.CustomAttribute1
                        $ADU = get-ADUser $m -Properties ProtectedFromAccidentalDeletion
                        $Container = $ADU.DistinguishedName -Replace("CN=","")
                        $Container = $Container -Replace(",OU=","/")
                        $Container = $Container -Replace(",DC=","/")
                        $Script:txtContainer.Text = $Container
                        $Script:txtDate.Text = $vmbx.LitigationHoldDate
                        If ($vmbx.LitigationHoldOwner.Length -gt 0)
                        {
                            $Script:txtBy.Text = (get-mailbox $vmbx.LitigationHoldOwner).DisplayName
                        }
                        $Script:txtProtect.Text = $ADU.ProtectedFromAccidentalDeletion
                        $Script:txtHide.Text = $vmbx.HiddenFromAddressListsEnabled
                        $Script:txtAccess.Text = $Script:MbxAccessForm 
                        $Script:txtInPlace.Text = $vmbx.InPlaceHolds
                        #Build Retention
                        $Script:locListBox.Items.Clear()
                        If ($VADUsr.ExtensionAttribute14.Length -gt 0)
                        {
                            $LocArray = $VADUsr.ExtensionAttribute14.split(",")
                            foreach ($element in $LocArray)
                            {
                                [void] $Script:locListBox.Items.Add($element.TrimStart())  # Add element to listbox 
                            }
                            $Script:locListBox.Refresh()
                        }
                        $Script:StoredHolds = import-csv "E:\O365AdminShared\Data\LegalDBOwners.csv" |Sort-Object EmpNo
                        $Script:DBOwnerHolds = ""
                        Foreach ($SH in $Script:StoredHolds)
                        {
                            If ($SH.EmpNo -eq $ENo)
                            {
                                $Val = "Listed as DBOwner for: " + $SH.CaseName + "/" + $SH.AppType
                                $Script:DBOwnerHolds = $Script:DBOwnerHolds + $SH.CaseName + " , "
                                [void] $Script:locListBox.Items.Add($val)
                            }
                        } 
                        Publish-ViewForm
                    }
                    else
                    {
                        write-host "Mailbox not found for: " $m
                    }
                }
            }
            $Script:chkViewUsr.Checked = $False
            $Script:txtENo.Visible = $False
            $Script:ButViewENo.Visible = $False
            $Script:GettingData.Visible = $False
            $Global:cancelButton.Visible = $True
        })

    ## New Hold Name
    $Script:lblHoldName = New-Object System.Windows.Forms.Label   
        $Script:lblHoldName.Text = "New Hold Name: "
        $Script:lblHoldName.Top = $Top+20; $Script:lblHoldName.Left = 10; $Script:lblHoldName.Width=150 ;$Script:lblHoldName.AutoSize = $true
        $Script:lblHoldName.Visible = $False 
        $form.Controls.Add($Script:lblHoldName)    # Add to Form 
        #
        $Script:txtHoldName = New-Object Windows.Forms.TextBox
        $Script:txtHoldName.Top = $Top; $Script:txtHoldName.Left = $Col1Left; $Script:txtHoldName.Width = 340;
        $Script:txtHoldName.Text = ""
        $Script:txtHoldName.Visible = $False
        $form.Controls.Add($Script:txtHoldName)    # Add to Form
            
#    $Col2Top = 40
    ## For Entire Standard Hold
    $Script:chkStdHold = New-Object System.Windows.Forms.Checkbox
        $Script:chkStdHold.Text = "for Standard Hold"  
        $Script:chkStdHold.Top = $Col2Top ; $Script:chkStdHold.Left = $Col2Left; $Script:chkStdHold.Width=150 ;$Script:chkStdHold.AutoSize = $true 
        $Script:chkStdHold.Visible = $False
        $form.Controls.Add($Script:chkStdHold)    # Add to Form 
        $Script:chkStdHold.Add_Click({
            Hide-All
            $Script:chkStdHold.Checked = $True
            $Script:chkStdHold.Visible = $True
            $Script:chkfor40180.Checked = $False
            $Script:chkfor40180.Visible = $True
            $Script:lblActHolds.Visible = $True
            $Script:txtActHolds.Visible = $True
            $Script:txtActHolds.Visible = $True
            $Script:txtActHolds.Items.Clear()
            Foreach ($Hold in $Script:StdHolds)
            {
                [void] $Script:txtActHolds.Items.Add($Hold.HoldName)
            }
        })

    $Col2Top = $Col2Top + 20
    ## For Entire 40/180 Hold
    $Script:chkfor40180 = New-Object System.Windows.Forms.Checkbox
        $Script:chkfor40180.Text = "for 40 or 180 Day Hold"  
        $Script:chkfor40180.Top = $Col2Top ; $Script:chkfor40180.Left = $Col2Left; $Script:chkfor40180.Width=150 ;$Script:chkfor40180.AutoSize = $true 
        $Script:chkfor40180.Visible = $False
        $form.Controls.Add($Script:chkfor40180)    # Add to Form 
        $Script:chkfor40180.Add_Click({
            Hide-All
            $Script:chkStdHold.Checked = $False
            $Script:chkStdHold.Visible = $True
            $Script:chkfor40180.Checked = $True
            $Script:chkfor40180.Visible = $True
            $Script:lblActHolds.Visible = $True
            $Script:txtActHolds.Visible = $True
            $Script:txtActHolds.Visible = $True
            $Script:txtActHolds.Items.Clear()
            Foreach ($Hold in $Script:40DayHolds)
            {
                [void] $Script:txtActHolds.Items.Add($Hold.HoldName)
            }
            Foreach ($Hold in $Script:180DayHolds)
            {
                [void] $Script:txtActHolds.Items.Add($Hold.HoldName)
            }
        })
        
    $Top = $Top + 30
    ## List of Holds
    $Script:lblActHolds = New-Object System.Windows.Forms.Label   
        $Script:lblActHolds.Text = "Active Legal Holds:"
        $Script:lblActHolds.Top = $Top; $Script:lblActHolds.Left = 10; $Script:lblActHolds.Width=150 ;$Script:lblActHolds.AutoSize = $true
        $Script:lblActHolds.Visible = $False 
        $form.Controls.Add($Script:lblActHolds)    # Add to Form 
        # 
        $Script:txtActHolds = New-Object Windows.Forms.ComboBox
        $Script:40DayHolds = $HoldDetails |Where-Object {$_.HoldName -like "*40 Day*"} |Sort-Object Holdname
        $Script:180DayHolds = $HoldDetails |Where-Object {$_.HoldName -like "*180 Day*"} |Sort-Object Holdname
        $Script:StdHolds = $HoldDetails |Where-Object {($_.HoldName -notlike "*40 Day*") -and ($_.HoldName -notlike "*180 Day*")} |Sort-Object Holdname
        $Script:txtActHolds.Top = $Top; $Script:txtActHolds.Left = $Col1Left; $Script:txtActHolds.Width = 360;
        $Script:txtActHolds.Visible = $False
        $Script:txtActHolds.Text = ""
        $form.Controls.Add($Script:txtActHolds)
        $Script:txtActHolds.Add_Click({
            If (($Script:chkExistHold.Checked -eq $True) -or ($Script:chkStdRel.Checked -eq $True) -or ($Script:chkRenameHold.Checked -eq $True))
            {
                $Script:txtActHolds.Text = ""
                $Script:txtENos.Text = ""
                $Script:txtDBOwner.Text = ""
                If ($Script:chkRenameHold.Checked -eq $False)
                {
                    $Script:txtRelENos.Text = "Enter comma separated ENo's"
                    $Script:txtRelDBOwner.Text = "Enter comma separated ENo's"
                    $Script:txtNoENos.Text = ""
                    $Script:txtNoDBOwner.Text = ""
                    $Script:txtRenameHolds.Text = ""
                }
                $Script:ButGetENo.visible = $true
            }
        })

    $Script:ButGetENo = New-Object Windows.Forms.Button
        $Script:ButGetENo.Location = New-object System.Drawing.Size(505,($Top-10))
        $Script:ButGetENo.Size = new-Object System.Drawing.Size(($Top-10),35)
        $Script:ButGetENo.Text = "Get Enabled Users"
        $Script:ButGetENo.visible = $false
        $form.Controls.Add($Script:ButGetENo)
        $Script:ButGetENo.Add_Click({                
            $Script:ButGetENo.visible = $false
            $Script:FindHold = $Script:txtActHolds.SelectedItem + "*"
            If ($Script:FindHold -ne "*")
            {
                $OwnFound = ""
                $Script:ThisHold = $AllHold |where-object {$_.HoldName -like $Script:FindHold} |Sort-Object EmpNo
                If ($Script:ThisHold.Count -ne "")
                {
                    $Script:txtDBOwner.Text = ""
                    $Script:txtENos.Items.Clear()
                    $Script:txtDBOwner.Items.Clear()
                    foreach ($t in $Script:ThisHold)
                    {
                        If ($t.Enabled -ne "DBOwner")
                        {
                            [void] $Script:txtENos.Items.Add($T.EmpNo + " (" + $T.Name +")")
                        }
                        else
                        {
                            [void] $Script:txtDBOwner.Items.Add($T.EmpNo + " (" + $T.Name +")")
                            $OwnFound = "Yes"
                        }
                    }
                    $Script:txtENos.SelectedIndex = 0
                    If ($OwnFound -eq "Yes")
                    {
                        $Script:txtDBOwner.SelectedIndex = 0
                    }
                    $Script:lblENos.Visible = $True
                    $Script:txtENos.Visible = $True
                    $Script:lblDBOwner.Visible = $True
                    $Script:txtDBOwner.Visible = $True
                    $HoldDetails |where-object {$_.HoldName -like $Script:FindHold}
                    $Script:txtAttorney.Text = ($HoldDetails |where-object {$_.HoldName -like $Script:FindHold}).Attorney
                    If (($Script:chkStd.checked -eq $True) -or ($Script:chkRenameHold.Checked -eq $True))
                    {
                        $Script:lblHRRep.Visible = $False
                        $Script:txtHRRep.Visible = $False
                    }
                    else
                    {
                        $Script:txtHRRep.Text = ($HoldDetails |where-object {$_.HoldName -like $Script:FindHold}).HRRep
                        $Script:lblHRRep.Visible = $True
                        $Script:txtHRRep.Visible = $True
                    }
                    If ($Script:chkStdRel.Checked -eq $False)
                    {
                        $Script:chkSilent.Visible = $True
                        $Script:chkHHold.Visible = $True
                    }
                    else
                    {
                        $Script:lblEMailDate.Visible = $True
                        $Script:txtEMailDate.Visible = $True
                        $Script:lblHRRep.Visible = $False
                        $Script:txtHRRep.Visible = $False
                        $Script:txtAttorney.Text = "A.Uteg"               
                    }
                    If ((($HoldDetails |where-object {$_.HoldName -like $Script:FindHold}).Silent) -eq "True")
                    {
                        $Script:chkSilent.Checked = $True
                    }
                    If ($Script:chkExistHold.Checked -eq $True)
                    {
                        If ($Script:chkStd.Checked -eq $True)
                        {
                            $Script:lblAttorney.Visible = $False
                            $Script:txtAttorney.Visible = $False
                            $Script:chkHHold.Visible = $False
                            $Script:chkSilent.Visible = $False
                        }
                    }
                    If ($Script:chkRenameHold.Checked -eq $False)
                    {
                        $Script:lblRelENos.Visible = $True
                        $Script:txtRelENos.Visible = $True
                        $Script:txtRelDBOwner.Visible = $True
                        $Script:lblRelDBOwner.Visible = $True
                        $Script:lblAttorney.Visible = $True
                        $Script:txtAttorney.Visible = $True
                    }
                    else
                    {
                        $Script:txtRenameHolds.Text = $Script:txtActHolds.SelectedItem
                        $Script:lblRenameHolds.Visible = $True
                        $Script:txtRenameHolds.Visible = $True
                        $Script:chkHHold.Visible = $False
                        $Script:chkSilent.Visible = $False
                    }
                    $Script:txtNoENos.Text = $Script:txtENos.Items.count
                    $Script:txtNoDBOwner.Text = $Script:txtDBOwner.Items.count
                    $Script:lblCountTitle.Visible = $True
                    $Script:txtNoENos.Visible = $True
                    $Script:txtNoDBOwner.Visible = $True
                }
                else
                {
                    write-host "No records found for this hold" -ForegroundColor Red
                }
            }
            else
            {
                write-host "Invalid name of Legal Hold" -ForegroundColor Red
            }
        })

    ## New Hold Name
    $Script:lblHoldName = New-Object System.Windows.Forms.Label   
        $Script:lblHoldName.Text = "New Hold Name: "
        $Script:lblHoldName.Top = $Top; $Script:lblHoldName.Left = 10; $Script:lblHoldName.Width=150 ;$Script:lblHoldName.AutoSize = $true
        $Script:lblHoldName.Visible = $False 
        $form.Controls.Add($Script:lblHoldName)    # Add to Form 
        #
        $Script:txtHoldName = New-Object Windows.Forms.TextBox
        $Script:txtHoldName.Top = $Top; $Script:txtHoldName.Left = $Col1Left; $Script:txtHoldName.Width = 340;
        $Script:txtHoldName.Text = ""
        $Script:txtHoldName.Visible = $False
        $form.Controls.Add($Script:txtHoldName)    # Add to Form

    $Top = $Top + 30
    #Number of Staff Column Title
    $Script:lblCountTitle = New-Object Windows.Forms.Label
        $Script:lblCountTitle.Top = $Top; $Script:lblCountTitle.Left = 285; $Script:lblCountTitle.Width = 80
        $Script:lblCountTitle.Text = "#Enabled"
        $Script:lblCountTitle.Visible = $False
        $form.Controls.Add($Script:lblCountTitle)    # Add to Form

    ## Hardware Hold
    $Script:chkHHold = New-Object System.Windows.Forms.Checkbox
        $Script:chkHHold.Text = "Hardware Hold" 
        $Script:chkHHold.Top = $Top ; $Script:chkHHold.Left = $Col1Left+30; $Script:chkHHold.Width=150 ;$Script:chkHHold.AutoSize = $true 
        $Script:chkHHold.Checked = $True
        $Script:chkHHold.Visible = $False
        $form.Controls.Add($Script:chkHHold)    # Add to Form

    ## Silent Hold
    $Script:chkSilent = New-Object System.Windows.Forms.Checkbox
        $Script:chkSilent.Text = "Silent Hold" 
        $Script:chkSilent.Top = $Top ; $Script:chkSilent.Left = $Col2Left+30; $Script:chkSilent.Width=150 ;$Script:chkSilent.AutoSize = $true 
        $Script:chkSilent.Checked = $False
        $Script:chkSilent.Visible = $False
        $form.Controls.Add($Script:chkSilent)    # Add to Form
        $Script:chkSilent.Add_Click({
            If ($chkSilent.Checked -eq $True)
            {
                 $Script:txtHoldName.Text = "NonDisclosed - " + $Script:txtHoldName.Text
            }
            else
            {
                $Script:txtHoldName.Text = $Script:txtHoldName.Text -replace ("NonDisclosed - ","")
            }
         })

    $Top = $Top + 23
    ## Employee Numbers
    $Script:lblENos  = New-Object System.Windows.Forms.Label   
        $Script:lblENos.Text = "Staff to Hold:"  
        $Script:lblENos.Top = $Top ; $Script:lblENos.Left = 10; $Script:lblENos.Width=150 ;$Script:lblENos.AutoSize = $true
        $Script:lblENos.Visible = $False
        $form.Controls.Add($Script:lblENos)    # Add to Form 
        # 
        $Script:txtENos = New-Object Windows.Forms.ComboBox
        $Script:txtENos.Top = $Top; $Script:txtENos.Left = $Col1Left; $Script:txtENos.Width = 165;
        $Script:txtENos.Visible = $False
        $form.Controls.Add($Script:txtENos)    # Add to Form

    #No of Enabled Staff
    $Script:txtNoENos = New-Object Windows.Forms.TextBox
        $Script:txtNoENos.Top = $Top; $Script:txtNoENos.Left = 300; $Script:txtNoENos.Width = 25;
        $Script:txtNoENos.Visible = $False
        $form.Controls.Add($Script:txtNoENos)    # Add to Form

    ## Release Employee #'s
    $Script:lblRelENos  = New-Object System.Windows.Forms.Label   
        $Script:lblRelENos.Text = ":Staff to Add/Release"  
        $Script:lblRelENos.Top = $Top ; $Script:lblRelENos.Left = 505; $Script:lblRelENos.Width=150 ;$Script:lblRelENos.AutoSize = $true
        $Script:lblRelENos.Visible = $False
        $form.Controls.Add($Script:lblRelENos)    # Add to Form 
        # 
        $Script:txtRelENos = New-Object Windows.Forms.TextBox
        $Script:txtRelENos.Top = $Top; $Script:txtRelENos.Left = 330; $Script:txtRelENos.Width = 165;
        $Script:txtRelENos.Text = "Enter comma separated ENo's"
        $Script:txtRelENos.Visible = $False
        $form.Controls.Add($Script:txtRelENos)    # Add to Form
        $Script:txtRelENos.Add_Click({
            If ($Script:txtRelENos.Text -eq "Enter comma separated ENo's")
            {
                $Script:txtRelENos.Text = ""
                If ($Script:chkStdRel.Checked -eq $True)
                {
                    $Output = $wshell.Popup("Enter 'ALL' to release all enabled or enter the employee numbers separated by commas to release select staff.",10,"Help Info",0+32)
                }
                else
                {
                    $Output = $wshell.Popup("Enter employee numbers separated by commas of individual to add to hold.",10,"Help Info",0+32)
                }
            }
            $Global:okButton.Visible = $True
        })

    $Top = $Top + 30
    ## DBOwners
    $Script:lblDBOwner  = New-Object System.Windows.Forms.Label   
        $Script:lblDBOwner.Text = "DBOwner(s):"  
        $Script:lblDBOwner.Top = $Top ; $Script:lblDBOwner.Left = 10; $Script:lblDBOwner.Width=150 ;$Script:lblDBOwner.AutoSize = $true
        $Script:lblDBOwner.Visible = $False
        $form.Controls.Add($Script:lblDBOwner)    # Add to Form 
        # 
        $Script:txtDBOwner = New-Object Windows.Forms.ComboBox
        $Script:txtDBOwner.Top = $Top; $Script:txtDBOwner.Left = $Col1Left; $Script:txtDBOwner.Width = 165;
        $Script:txtDBOwner.Visible = $False
        $form.Controls.Add($Script:txtDBOwner)    # Add to Form

    #No of DBOwners
        $Script:txtNoDBOwner = New-Object Windows.Forms.TextBox
        $Script:txtNoDBOwner.Top = $Top; $Script:txtNoDBOwner.Left = 300; $Script:txtNoDBOwner.Width = 25;
        $Script:txtNoDBOwner.Visible = $False
        $form.Controls.Add($Script:txtNoDBOwner)    # Add to Form

    ## Release DBOwners
    $Script:lblRelDBOwner  = New-Object System.Windows.Forms.Label   
        $Script:lblRelDBOwner.Text = ":DBOwners to Add/Release"  
        $Script:lblRelDBOwner.Top = $Top ; $Script:lblRelDBOwner.Left = 505; $Script:lblRelDBOwner.Width=150
        $Script:lblRelDBOwner.Visible = $False
        $form.Controls.Add($Script:lblRelDBOwner)    # Add to Form 
        # 
        $Script:txtRelDBOwner = New-Object Windows.Forms.TextBox
        $Script:txtRelDBOwner.Top = $Top; $Script:txtRelDBOwner.Left = 330; $Script:txtRelDBOwner.Width = 165;
        $Script:txtRelDBOwner.Text = "Enter comma separated ENo's"
        $Script:txtRelDBOwner.Visible = $False
        $form.Controls.Add($Script:txtRelDBOwner)    # Add to Form
        $Script:txtRelDBOwner.Add_Click({
            If ($Script:txtRelDBOwner.Text -eq "Enter comma separated ENo's")
            {
                $Script:txtRelDBOwner.Text = ""
                If ($Script:chkStdRel.Checked -eq $True)
                {
                    $Output = $wshell.Popup("Enter 'ALL' to release all enabled or enter the employee numbers separated by commas to release select staff.",10,"Help Info",0+32)
                }
                else
                {
                    $Output = $wshell.Popup("Enter employee numbers separated by commas of individual to add to as a DBOwner.",10,"Help Info",0+32)
                }
            }
            $Global:okButton.Visible = $True
        })

    $Top = $Top + 30
    #Rename Hold Name
    $Script:lblRenameHolds = New-Object System.Windows.Forms.Label   
        $Script:lblRenameHolds.Text = "New Hold Name: "
        $Script:lblRenameHolds.Top = $Top; $Script:lblRenameHolds.Left = 10; $Script:lblRenameHolds.Width=150 ;$Script:lblRenameHolds.AutoSize = $true
        $Script:lblRenameHolds.Visible = $False 
        $form.Controls.Add($Script:lblRenameHolds)    # Add to Form 
        # 
        $Script:txtRenameHolds = New-Object Windows.Forms.TextBox
        $Script:txtRenameHolds.Top = $Top; $Script:txtRenameHolds.Left = $Col1Left; $Script:txtRenameHolds.Width = 360;
        $Script:txtRenameHolds.Visible = $False
        $Script:txtRenameHolds.Text = ""
        $form.Controls.Add($Script:txtRenameHolds)
        $Script:txtRenameHolds.Add_Click({
            $Global:OKButton.Visible = $True
        })


    #Add HR Rep
    $Script:lblHRRep = New-Object System.Windows.Forms.Label
        $Script:lblHRRep.Text = ":HR Rep"
        $Script:lblHRRep.Top = $Top ; $Script:lblHRRep.Left = 505; $Script:lblHRRep.Width=10 ;$Script:lblHRRep.AutoSize = $true
        $Script:lblHRRep.Visible = $False
        $form.Controls.Add($Script:lblHRRep)    # Add to Form 
        # 
        $Script:txtHRRep = New-Object Windows.Forms.ComboBox  
        $Script:txtHRRep.Top = $Top; $Script:txtHRRep.Left = 330; $Script:txtHRRep.Width = 165;
        $Script:HRReps = $HoldDetails.HRRep | Sort-Object |Get-Unique
        Foreach ($HRReps in $Script:HRReps)
        {
            If ($HRReps.Length -gt 0)
            {
                [void] $Script:txtHRRep.Items.Add($HRReps)
            }
        }
        $Script:txtHRRep.Visible = $False
        $form.Controls.Add($Script:txtHRRep)    # Add to Form 
    
    #Add Released By
    $Script:lblAttorney = New-Object System.Windows.Forms.Label
        $Script:lblAttorney.Text = "Released by:"  
        $Script:lblAttorney.Top = $Top ; $Script:lblAttorney.Left = 10; $Script:lblAttorney.Width=150 ;$Script:lblAttorney.AutoSize = $true
        $Script:lblAttorney.Visible = $False
        $form.Controls.Add($Script:lblAttorney)    # Add to Form 
        # 
        $Script:txtAttorney = New-Object Windows.Forms.ComboBox
        $Script:txtAttorney.Top = $Top; $Script:txtAttorney.Left = $Col1Left; $Script:txtAttorney.Width = 165;
        $Script:Attorneys = $HoldDetails.Attorney | Sort-Object |Get-Unique
        Foreach ($Attorneys in $Script:Attorneys)
        {
            If ($Attorneys.length -gt 0)
            {
                [void] $Script:txtAttorney.Items.Add($Attorneys)
            }
        }
        $Script:txtAttorney.Visible = $False
        $form.Controls.Add($Script:txtAttorney)    # Add to Form

    $Top = $Top + 30
    #Add Release Date
    $Script:lblEMailDate = New-Object System.Windows.Forms.Label
        $Script:lblEMailDate.Text = "EMail Date Releasing:"
        $Script:lblEMailDate.Top = $Top; $Script:lblEMailDate.Left = 10; $Script:lblEMailDate.Width=150 ;$Script:lblEMailDate.AutoSize = $true
        $Script:lblEMailDate.Visible = $False
        $form.Controls.Add($Script:lblEMailDate)    # Add to Form
        #
        $Script:txtEMailDate = New-Object Windows.Forms.DateTimePicker
        $Script:txtEMailDate.Top = $Top; $Script:txtEMailDate.Left = $Col1Left; $Script:txtEMailDate.Width = 120;
        $Script:txtEMailDate.Format = [windows.forms.datetimepickerFormat]::custom
        $Script:txtEMailDate.Visible = $False
        $Script:txtEMailDate.CustomFormat = "MM/dd/yyyy"
        $Script:txtEMailDate.Text = (get-date)
        $form.Controls.Add($Script:txtEMailDate)    # Add to Form

        Add-FormStandardButtons
        $Global:okButton.Text = "Proceed"

    Add-ViewUsrDeatilsForm
}

Function Build-RetentComment
{
    #Enable Hold or Add Additional Comment
    If ($Script:chkExistHold.Checked -eq $False)
    {
        $HoldDetails = import-csv "e:\O365AdminShared\Data\Legal-AllHoldDetails.csv"
        If ($HoldDetails -like ("*"+ $HName +"*"))
        {
            #DoNothing this retention comment is arlready listed in the file
        }
        else
        {
            #Add this new hold to the list of all holds
            $Text = "{0},{1},{2},{3},{4},{5}" -f $HName, $Script:txtHRRep.Text, $Script:txtAttorney.Text, $Script:chkHHold.Checked, $Script:chkSilent.Checked, (get-Date).ToString("yyyy-MM-dd")
            Out-File -FilePath "e:\O365AdminShared\Data\Legal-AllHoldDetails.csv" -InputObject $Text -Append
        }
    }

    $Script:FullRetCmt = ""
        
    If ($Script:chkStd.Checked -eq $True)
    {
        $Script:RetCmt = $HName + "\" + $Script:txtAttorney.Text + "\"
    }
    else
    {
        $Script:RetCmt = $HName + " - " + $Script:txtHRRep.Text + "\" + $Script:txtAttorney.Text + "\"
    }
        
    If ($Script:chkHHold.Checked -eq $True)
    {
        $Script:RetCmt = $Script:RetCmt + "HHold"
    }
    else
    {
        $Script:RetCmt = $Script:RetCmt + "NoHHold"
    }
    #set the retention comment if there are other holds

    $Script:NewHoldCmt = $Script:RetCmt
    $Script:Enabled = $True
    If (($Script:u.ExtensionAttribute14.Length -gt 0) -and (($Script:u.ExtensionAttribute14 -notlike "Emergency*") -and ($Script:u.ExtensionAttribute14 -notlike "Standard*") -and ($Script:u.ExtensionAttribute14 -notlike "*Rehire*")))
    {
        If ($Script:u.ExtensionAttribute14 -notlike ("*" + $Script:RetCmt + "*"))
        {
            write-host "Adding ExtensionAttribute14 to comment"
            $Script:RetCmt = $Script:u.ExtensionAttribute14 + ", " + $Script:RetCmt
        }
        else
        {
            write-host "This user is already enabled for this hold" -ForegroundColor Red
            $Script:Enabled = $False
            $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - No change made - already enabled for this hold`n"
        }
    }
}

Function Check-ReportFile
{
    If (Test-Path $ReportFile)
    {
    #  Do Nothing
    }
    else
    {
        write-host "`tCreating new report file for this individual" -ForegroundColor Yellow
        $LineToWrite = "Legal Hold Activities for Emp No. " + $ENo + " [" +$Script:u.cn + "]`n"
        WriteReportEvent       
    }
}

Function Commit-CommentChange
{
    If ($Script:NewCmt -ne $Script:StoredCmt)
    {
#   Updating Comment
        If ($Script:ADExists -ne $True)
        {
           $LineToWrite = $WhoAmI + "`t" + "Updating ExtensionAttribute14 on the O365 mailbox"
           set-mailbox $ENo -RetentionComment (($Script:StoredCmt).Replace($Script:locListBox.SelectedItem,$Script:NewCmt))
        }
        else
        {
            $LineToWrite = $WhoAmI + "`t" + "Updating ExtensionAttribute14 on the AD account"
            write-host "`tNew Comment: " $Script:NewCmt
            Set-ADUser $ENo -Replace @{'ExtensionAttribute14' = $Script:NewCmt}
        }
    }
    else 
    {
        $Output = $wshell.Popup("No changes needed.",0,"No Changes Needed",0+32)    
        $LineToWrite = $WhoAmI + "`t" + "No changes Needed              :"
        WriteReportEvent
    }
}

Function Execute-Hold
{
    If ($Script:Mbx.LitigationHoldEnabled -eq $False)
    {
    #This person is not on legal hold
        Set-Mailbox $ENo -LitigationHoldEnabled $true
        write-host "`tEnabling legal hold for: " $ENo
        $LineToWrite = $WhoAmI + "`t" + "Enabling Legal Hold for " + $ENo + " with Retention Comment of: " + $Script:RetCmt + "`n"
        WriteReportEvent
        $Mem = [bool](Get-DistributionGroupMember iprosearch@ul.onmicrosoft.com | Where-Object {$_.PrimarySMTPAddress -eq $Mbx.PrimarySMTPAddress})
        If ($Mem -eq $False)
        {
            Add-DistributionGroupMember iprosearch@ul.onmicrosoft.com -Member $ENo -BypassSecurityGroupManagerCheck
            write-host "`tAdded User to iProSearch group" -ForegroundColor Green
            $LineToWrite = $WhoAmI + "`t" + "Added user to the iProSearch Group for " + $ENo
        }
        else
        {
            write-host "`tUser is already a member of the iProSearch group" -ForegroundColor Yellow
            $LineToWrite = $WhoAmI + "`t" + "User is Already a member of the iProSearch Group"
        }
        WriteReportEvent
    }
    else
    {
        write-host "`tLegal Hold is already enabled on this account" -ForegroundColor Yellow
        $LineToWrite = $WhoAmI + "`tLegal Hold is already enabled on this account"
        WriteReportEvent
    }

    If ($Script:usrDetails.ProtectedFromAccidentalDeletion -eq $False)
    {
        $Script:Protection = "True"
        Set-ADObjProtect
    }

    If (($Script:Mbx.CustomAttribute1 -like "*Ex-*") -and ($Script:strUserPath -notlike "*_Extend*"))
    {
		   $OULoc = "ExtendedHold OU"
           write-host "`tMoving ADOject to the ExtendedHoldOU for: " $ENo -ForegroundColor Green
           Move-ADObjectOU
	}

    If ($ADExists -eq $True)
    {
        #Build new retention comment based on the AD Account Ext14 attribute
        If (($Script:u.ExtensionAttribute14.length -eq 0) -or ($Script:u.ExtensionAttribute14 -like "Emergency*") -or ($Script:u.ExtensionAttribute14 -like "Standard*") -or ($Script:u.ExtensionAttribute14 -like "Rehire*"))
        {
            $Script:NewCmt = $Script:RetCmt
            $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - Enabled legal hold for this individual`n"
        }
        else
        {
            $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - Added additional legal hold for this individual`n"
            $Script:StoredCmt = $Script:u.ExtensionAttribute14
            $Script:NewCmt = $Script:RetCmt
        }
        $LineToWrite = $WhoAmI + "`t" + "New ADRetention Activity       :`t" + $Script:NewCmt + "`n"
    }
    else
    {
        #Build new retention comment based on the O365Mailbox Ext14 attribute
        $Script:StoredCmt = $Script:Mbx.CustomAttribute14 + $Script:RetCmt
        $LineToWrite = $WhoAmI + "`t" + "New O365MbxRetention Activity         :`t" + $Script:NewCmt + "`n"
        $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - No AD Acount Updated O365 Mailbox CustomAttribute14 for this individual`n"
    }
    WriteReportEvent
    Commit-CommentChange
}

# Gets AD Account Properties
Function GetAcctInfo
{
    $Script:ADCmt = ""
    $Script:u = ""
    
    If ($ADExists -eq "True")
    {
        $strDN = (Get-ADUser $ENo).DistinguishedName
        $Script:strUserPath = [string]::format("LDAP://{0}", $strDN)
        $Script:u = Get-ADUser $ENo -Properties *
        write-host "`tCurrent AD Account Retention Comment: " $Script:u.ExtensionAttribute14
        $Script:ADCmt = $Script:u.ExtensionAttribute14
        $ObjExists = [bool]($Script:UsrDetails = Get-ADObject -Identity $strDN -Properties ProtectedFromAccidentalDeletion -ErrorAction silentlyContinue)
        If ($ObjExists -eq $False)
        {
            $cnt = 0
            Do
            {
                start-sleep -Seconds 5
                $ObjExists = [bool]($Script:UsrDetails = Get-ADObject -Identity $strDN -Properties ProtectedFromAccidentalDeletion -ErrorAction silentlyContinue)
                If ($ObjExists -eq $False)
                {
                    #Attempting to find the AD object
                    $Cnt++
                }
            } while (($Cnt -le 5) -and ($ObjExists -eq $False))

            If ($Cnt -ge 5)
            {
                write-host "`tAD Record not found"
                $LineToWrite = $WhoAmI + "`t" + "AD Record Not Found!"
                WriteReportEvent
            }
        }
     }
    else
    {
       write-host "`tNo ADAccount Found" -ForegroundColor Red
    }

    $MbxExists = [bool]($Script:inf = get-mailbox -identity $ENo -ErrorAction SilentlyContinue)
    If ($Script:chkStdRel.Checked -eq $False)
    {
        If ($MbxExists -eq "True")
        {
            $LineToWrite = $WhoAmI + "`t" + "Account Details Before Adding New Hold: "
            WriteReportEvent
            Log-AcctDetails
            PrtMBxAccess($Script:MbxAccess)
        }
        else
        {
            write-host "`n`tMailbox does not exist for Emp#" $ENo -ForegroundColor Red
        }
    }
}

function Hide-All
{
    $Script:ButGetENo.visible = $False
    $Script:chkStdHold.Visible = $False
    $Script:chkfor40180.Visible = $False
    $Script:chk40Day.Visible = $False
    $Script:chk180Day.Visible = $False
    $Script:chkStd.Visible = $False
    $Script:chkSilent.Visible = $False
    $Script:chkHHold.Visible = $False
    $Script:ButViewENo.Visible = $False
    $Script:txtENo.Visible = $False
    Hide-Details
}

Function Hide-Details
{
    $Script:lblENos.Visible = $False
    $Script:txtENos.Visible = $False
    $Script:lblRelENos.Visible = $False
    $Script:txtRelENos.Visible = $False
    $Script:lblDBOwner.Visible = $False
    $Script:txtDBOwner.Visible = $False
    $Script:lblRelDBOwner.Visible = $False
    $Script:txtRelDBOwner.Visible = $False
    $Script:lblActHolds.Visible = $False
    $Script:txtActHolds.Visible = $False
    $Script:lblRenameHolds.Visible = $False
    $Script:txtRenameHolds.Visible = $False
    $Script:lblHoldName.Visible = $False
    $Script:txtHoldName.Visible = $False
    $Script:lblEMailDate.Visible = $False
    $Script:txtEMailDate.Visible = $False
    $Script:txtAttorney.Text = ""
    $Script:lblAttorney.Visible = $False
    $Script:txtAttorney.Visible = $False
    $Script:lblHRRep.Visible = $False
    $Script:txtHRRep.Visible = $False
    $Script:lblCountTitle.Visible = $False
    $Script:txtNoENos.Visible = $False
    $Script:txtNoDBOwner.Visible = $False
    $Script:ButGetENo.visible = $false
    If ($Script:chkStdRel.checked -eq $True)
    {
        $Script:txtAttorney.Text = "A.Uteg"
    }
    $Script:txtActHolds.Items.Clear()
    $Script:txtDBOwner.Items.Clear()
    $Script:txtENos.Items.Clear()
    $Script:txtNoENos.Text = ""
    $Script:txtNoDBOwner.Text = ""
    $Script:txtENos.Text = ""
    $Script:txtDBOwner.Text = ""
    
    $Script:txtENo.Text = "Enter comma separated ENo's"
    $Script:txtRelENos.Text = "Enter comma separated ENo's"
    $Script:txtRelDBOwner.Text = "Enter comma separated ENo's"
    $Script:txtActHolds.Text = "Select Hold from List"
    $Global:okButton.Visible = $False
}

Function Hide-ViewUsrDetails
{
    $Script:chkNewHold.Visible = $True
    $Script:chk40Day.Visible = $True
    $Script:chk180Day.Visible = $True
    $Script:chkExistHold.Visible = $True
    $Script:chkStdRel.Visible = $True
    $Script:chkViewUsr.Visible = $True
    $Script:lblName.Visible = $False
    $Script:txtName.Visible = $False
    $Script:lblType.Visible = $False
    $Script:txtType.Visible = $False
}

Function Log-AcctDetails
{
    $LineToWrite = $WhoAmI + "`t" + "DisplayName                    :" + "`t" + $Script:Mbx.DisplayName
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "Employee Number                :" + "`t" + $Script:Mbx.Alias
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "Employee Type                  :" + "`t" + $Script:Mbx.CustomAttribute1
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "AD Container                   :" + "`t" + $Script:strUserPath
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "LitigationHoldEnabled          :" + "`t" + $Script:Mbx.LitigationHoldEnabled
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "LitigationHoldDate             :" + "`t" + $Script:Mbx.LitigationHoldDate
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "LitigationHoldOwner            :" + "`t" + $Script:Mbx.LitigationHoldOwner
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "Active InPlace Holds           :" + "`t" + ($Script:Mbx.InPlaceHolds -join ",")
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "Hidden from Address Book O365  :" + "`t" + $Script:Mbx.HiddenFromAddressListsEnabled
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "Hidden from Address Book AD    :" + "`t" + $Script:u.msExchHideFromAddressLists
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "O365 Retention Comment         :" + "`t" + $Script:Mbx.CustomAttribute14
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "AD Retention Comment           :" + "`t" + $Script:ADCmt
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "AD Object Protected            :" + "`t" + $Script:usrDetails.ProtectedFromAccidentalDeletion
    WriteReportEvent
}

Function Move-ADObjectOU
{
    $Script:MoveObj = "False"
    $ErrorActionPreference = "SilentlyContinue"
#    write-host "      bute1: " $Script:u.ExtensionAttribute1
#    write-host "     enable: " $Script:chkNewHold.Checked
#    write-host "   disabled: " $Script:chkStdRel.Checked
#    write-host "strUserPath: " $Script:strUserPath
    If (($Script:u.ExtensionAttribute1 -like "*Ex-*") -and ($Script:chkNewHold.Checked -eq "Checked"))
    {
        $OULoc = "ExtendedHold OU"
        write-host "Moving AD Object to the ExtendedHold OU"
        $LineToWrite = $WhoAmI + "`t" + "Moved to ExtendedHold from     :" + "`t" + $Script:strUserPath
        $Script:MoveObj = [bool](Move-ADObject ($Script:strUserPath -replace("LDAP://","")) 'OU=_ExtendedHold,DC=global,DC=ul,DC=com')
        WriteReportEvent
    }
    elseif (($Script:u.ExtensionAttribute1 -like "*Ex-*") -and ($Script:chkStdRel.Checked -eq "Checked"))
    {
        $OULoc = "Disabled OU"
        write-host "Moving AD Object to the Disabled OU"
        $LineToWrite = $WhoAmI + "`t" + "Moved to ExtendedHold from     :" + "`t" + $Script:strUserPath
        $Script:MoveObj = [bool](Move-ADObject ($Script:strUserPath -replace("LDAP://","")) 'OU=Disabled,DC=global,DC=ul,DC=com')
        WriteReportEvent
    }
    elseif ($Script:u.ExtensionAttribute1 -notlike "*Ex-*")
    {
        $OULoc = "Active"
    }
    $ErrorActionPreference = "Continue"
    write-host "MovedObj: " $script:MoveObj

    If ($Script:MoveObj -eq $True)
    {
        If ($OULoc -like "*Disabled*")
        {
            write-host "Moving AD Object to the Disabled OU"
            $Script:u.DistinguishedName = $Script:u.DistinguishedName.Remove($Script:u.DistinguishedName.IndexOf("OU=")) + "OU=Disabled,DC=global,DC=ul,DC=com"
            $LineToWrite = $WhoAmI + "`t" + "Moved to Disabled OU from     :" + "`t" + $Script:strUserPath
            WriteReportEvent
        }
        elseif ($OULoc -like "*Extend*")
        {
            write-host "Moving AD Object to the ExtendedHold OU"
            $Script:u.DistinguishedName = $Script:u.DistinguishedName.Remove($Script:u.DistinguishedName.IndexOf("OU=")) + "OU=_ExtendedHold,DC=global,DC=ul,DC=com"
            $LineToWrite = $WhoAmI + "`t" + "Moved to ExtendedHold from     :" + "`t" + $Script:strUserPath
            WriteReportEvent
        }
        elseif ($OULoc -like "*Active*")
        {
            $output = $wshell.Popup("This individual needs to be moved back to an Active OU.",0,"Move to Active OU",0+32)
        }
        else
        {
            $Output = $wshell.Popup("Insufficient privilege to move account to the " + $OULoc + ": " + ($Script:strUserPath -replace("LDAP://","")),0,"Insufficient Privilege",0+32)
        }
	}
}

Function Publish-ViewForm
{
    If ($Global:InputFocus -eq $null)
    {
        $Global:InputFocus = $Global:okButton
    }
    ## Finalize Form and Show Dialog
    $Global:ViewForm.Add_Shown( { $form.Activate(); $Global:InputFocus.Focus() } )  #Activate and Set Focus 
    $Global:Result = $Global:ViewForm.ShowDialog()          ## Show the form, and wait for the response
}

Function Remove-LegalHold
{
    #write-all records except for those related to this hold
    $NewHoldsFile = get-content "E:\O365AdminShared\Data\Legal-AllHoldDetails.csv"
    $Date = get-date -Format "yyyy-MMdd"
    $uTime = get-date -uformat %T
    $Time = $uTime.Replace(":", "")
    $BkOutFile = "E:\O365AdminShared\Data\Legal-AllHoldDetails-" + $Date + "-Time" + $Time + ".csv"
    $OutFile = "E:\O365AdminShared\Data\Legal-AllHoldDetails.csv"
    Rename-Item $OutFile $BkOutFile

    ForEach ($h in $NewHoldsFile)
    {
        If ($h -notlike ($Script:txtActHolds.SelectedItem + "*"))
        {
            #Add hold record into the new file
            Out-File -FilePath $OutFile -InputObject $h -Append
        }
        else
        {
            write-host "Hold has been removed from the AllHolds file: $hold"
        }    
    }
}

Function Release-FromHold #replaced Disable-Hold
{
    #This is to release individuals from Legal Holds - checks to so see if this individual is on other holds
    $ADExists = $False
    $StrDN = ""
    $Script:StoredCmt = ""
    $ErrorActionPreference = "SilentlyContinue"
    $ADExists = [bool]($Script:u = Get-ADUser $ENo -Properties *)
    $strdn = (Get-ADUser $ENo).DistinguishedName
    GetAcctInfo
    $ErrorActionPreference = "Continue"
    $MbxExists = [bool]($Script:Mbx = Get-Mailbox $ENo -ErrorAction SilentlyContinue)

    If ($Script:u.ExtensionAttribute14 -like "*,*")
    {
        write-host "`tThis individual is on multiple holds removing only details for this hold" -ForegroundColor Yellow
        $holds = $Script:u.ExtensionAttribute14 -replace (", ","`n`t ")
        write-host "`tHolds this inidividual is enabled for: `n`t"$Holds
        $Holds = $Script:u.ExtensionAttribute14 -split ","
        $NewHoldCmt = $Script:u.ExtensionAttribute14 -split ", "
        $RelHold = "*" + $Script:txtActHolds.SelectedItem + "*"
        $BuildCmt = ""
        Foreach ($H in $Holds)
        {
            If ($H -notlike $RelHold)
            {
                $BuildCmt = $BuildCmt + $H + ", "
            }
        }
        $BuildCmt = $BuildCmt.TrimStart(" ")
        $BuildCmt = $BuildCmt.TrimEnd(", ")
        Set-ADUser $ENo -Replace @{'ExtensionAttribute14' = $BuildCmt}
        $LineToWrite = $WhoAmI + "`t" + "Release (HoldName       )      :" + "`t" + $BuildCmt
        WriteReportEvent 
        $LineToWrite = $WhoAmI + "`t" + "Released (By/Date)             :" + "`t" + $Script:lblAttorney.Text + " - " + $Script:txtEMailDate.Text
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "Account remainson other active holds  `n"
        WriteReportEvent
        $PurgeStat = "`t" + $ENo + " (" + $Script:u.cn + " - " + $Script:u.ExtensionAttribute1 + ") - Account removed from this hold but remains on other active holds`n"
    }
    else
    {
        $Protected = Get-ADObject -Identity $strDN -Properties ProtectedFromAccidentalDeletion
        $PurgeStat = "`t" + $ENo + " (" + $Script:u.cn + ") - Account released from Legal Hold`n"
        Write-host "`tThis user is not on multiple Holds" -ForegroundColor Yellow
        $RelDetails = $Script:txtReleaseby.SelectedItem + "-" + $Script:txtEMailDate.Text + " email"
        $LineToWrite = $WhoAmI + "`t" + "Release Details (Who/Date)     :" + "`t" + $RelDetails
        WriteReportEvent

        set-mailbox $ENo -LitigationHoldEnabled $false
        write-host "`tDisabling Litigation Hold on Mailbox"
        $LineToWrite = $WhoAmI + "`t" + "Disabled Litigation Hold on Mailbox"
        WriteReportEvent

        Remove-DistributionGroupMember iprosearch -Member $ENo -BypassSecurityGroupManagerCheck -confirm:$False
        write-host "`tRemoving Account from iPro Search Group"
        $LineToWrite = $WhoAmI + "`t" + "Removed from iProSearch Group"
        WriteReportEvent

        If ($Script:UsrDetails.ProtectedFromAccidentalDeletion -eq $True)
        {
            Set-ADObject -Identity $strDN -ProtectedFromAccidentalDeletion $False -Credential $Global:AdmLiveCred
            write-host "`tDisabling ADObject Protection"
            $LineToWrite = $WhoAmI + "`t" + "Disabled ADObject Protection"
            WriteReportEvent
        }

        If ($Script:u.ExtensionAttribute1 -like "*Ex-*")
        {
            write-host "`tMoving ADObject from ExtendedHold to Disabled OU"
            Move-ADObjectOU
            $LineToWrite = $WhoAmI + "`t" + "Moved account from the ExtendedHold OU to the Disabled OU"
            WriteReportEvent
            write-host "`tCreate Change Request to Purge this Account" -ForegroundColor Yellow
            $PurgeStat = "`t" + $ENo + " (" + $Script:u.cn + ") - Account Released - Create Account purge request`n"
#            $PurgeStat = "`t" + $ENo + " (" + $Script:u.cn + " - " + $Script:u.ExtensionAttribute1 + ") - Account Released - Create Account purge request`n"
	    }

        If ($Container -like "*Active Directory Account*")
        {
            set-mailbox $ENo -RetentionComment $null -CustomAttribute14 $null
            write-host "`tNo ADAccount - Clearing Retention Comment on O365 Mailbox (ExtensionAttribute14)"
            write-host "`tManually purge this maibox" -ForegroundColor Red
            $LineToWrite = $WhoAmI + "`t" + "No AD Account found clearing the hold details in CustomAttribute14"
            WriteReportEvent
        }
        else
        {
            set-aduser $ENo -clear ExtensionAttribute14
            $Script:u.ExtensionAttribute14 = $null
            $LineToWrite = $WhoAmI + "`t" + "Clearing ADAccount ExtensionAttribute14`n"
            WriteReportEvent
        }

        If ($Script:inf.InPlaceHolds.length -ne 0)
        {
            $Hold = $Script:inf.InPlaceHolds
            foreach ($Hold in $Hold)
            {
                $CaseHold = Get-CaseHoldPolicy ($Hold -replace ("UniH",""))
                $Case = Get-ComplianceCase $CaseHold.CaseID
                $Output = $wshell.Popup("An active InPlace Hold was found on this account contact Legal to close Case Name: " + $Case.Name,0,"Active InplaceHold Found",0+32)
                $LineToWrite = $WhoAmI + "`t" + "Active InPlace Holds found Contact Legal to close the name " + $Case.Name
                WriteReportEvent
            }
        }
    }
    $Script:Recap = $Script:Recap + $PurgeStat
}

Function Rename-ExistingHold
{
    write-host "This section will rename the name of hold for all enabled staff/dbowners once complete"

#    $ThisHold = $AllHold |where-object {$_.HoldName -like $Script:FindHold} |Sort-Object EmpNo
    foreach ($t in $Script:ThisHold)
    {
        If ($t.Enabled -ne "DBOwner")
        {
            write-host "Renamaing retention comment for: " $t.EmpNo
            $ADU = get-ADUser $t.EmpNo -Properties ExtensionAttribute14
            write-host $ADU.ExtensionAttribute14
            $NewCmt = $ADU.ExtensionAttribute14 -replace ($Script:FindHold,$Script:txtRenameHolds.Text)
            write-host "to be replaced: " $NewCmt
            Set-ADUser $t.EmpNo -Replace @{ExtensionAttribute14 = $NewCmt}
            $Script:Recap = $Script:Recap + $t.EmpNo + " (" + $ADU.Name + ") " + " - Renamed retention hold name"
        }
        else
        {
            write-host "Renameing retention comment for DBOwner for: " $t.EmpNo
        }
    }

    #Rename the name of the hold in the all holds file
    #Rename the hold name
}

Function Set-ADObjProtect
{
    $ErrorActionPreference = "SilentlyContinue"
    switch ($Script:Protection)
    {
        "True"
        {
            Set-ADObject -Identity $strDN -ProtectedFromAccidentalDeletion $True -Credential $Global:AdmLiveCred
        }
        "False"
        {
            Set-ADObject -Identity $strDN -ProtectedFromAccidentalDeletion $False -Credential $Global:AdmLiveCred
        }
    }
    $ErrorActionPreference = "Continue"
    Start-Sleep -Seconds 10
    $Script:UsrDetails = Get-ADObject -Identity $strDN -Properties ProtectedfromAccidentalDeletion

    If (($Script:UsrDetails.ProtectedfromAccidentalDeletion -eq "True") -and ($Script:Protection -eq "True"))
    {
        $LineToWrite = $WhoAmI + "`t" + "AD Object Protection Enabled"
    }
    else
    {
        If (($Script:UsrDetails.ProtectedfromAccidentalDeletion -ne "True") -and ($Script:Protection -eq "False"))
        {
            $LineToWrite = $WhoAmI + "`t" + "AD Object Protection is Disabled"
        }
        else
        {
            $LineToWrite = $WhoAmI + "`t" + "Insufficient privilege to modify AD Object Protection"
            $output = $wshell.Popup("Insufficient privilege to modify AD Object Protection on account.",0,"Insufficient Privilege",0+32)
        }
    }
    WriteReportEvent
}

Function Show-WhoSilent
{
    If ($Script:chkExistHold.Checked -eq $False)
    {
        $Script:lblAttorney.Visible = $True
        $Script:txtAttorney.Visible = $True
        If ($Script:chkStd.checked -eq $True)
        {
            $Script:lblHRRep.Visible = $False
            $Script:txtHRRep.Visible = $False
        }
        else
        {
            $Script:lblHRRep.Visible = $True
            $Script:txtHRRep.Visible = $True        
        }
        $Script:chkSilent.Visible = $True
        $Script:chkHHold.Visible = $True
    }
}

Function Uncheck-Options
{
    $Script:chkStdHold.Checked = $False
    $Script:chkfor40180.Checked = $False
    $Script:chk40Day.Checked = $False
    $Script:chk180Day.Checked = $False
    $Script:chkStd.Checked = $False
    $Script:chkSilent.Checked = $False
}

Function Unhide-ViewUsrDetails
{
    $Script:chkNewHold.Visible = $False
    $Script:chk40Day.Visible = $False
    $Script:chk180Day.Visible = $False
    $Script:chkExistHold.Visible = $False
    $Script:chkStdRel.Visible = $False
    $Script:chkViewUsr.Visible = $False
    $Script:lblName.Visible = $True
    $Script:txtName.Visible = $True
    $Script:lblType.Visible = $True
    $Script:txtType.Visible = $True
}

Function ViewOpt-ExistingHold
{
    $Script:lblActHolds.Visible = $True
    $Script:txtActHolds.Visible = $True
    $Script:chkHHold.Visible = $False
    $Script:chkSilent.Visible = $False
}

$whoami = whoami
$wshell = New-Object -ComObject Wscript.Shell
$Year = (get-date).ToString("yyyy")
$HoldDetails = import-csv "e:\O365AdminShared\Data\Legal-AllHoldDetails.csv"
#Copy all hold details to the O365AdminShared\Data directory to add new activites for the day
$Dailyfile = ("\\usnbkutil100p\d$\ReportHistory\LegalHoldHistory\" + $Year + "\LegalHoldReport-" + (get-date).Tostring("yyyy-MMdd") + ".csv")
$ActFile = "e:\O365AdminShared\Data\LegalHoldReport-" + (get-date).Tostring("yyyy-MMdd") + ".csv"
$today = (get-date).toString("yyyy-MM-dd")
If (Test-Path $ActFile) {} else {Copy-Item $DailyFile -Destination $ActFile -force}

Do
{
$AllHold = import-csv $ActFile
Build-HoldMenu
$Global:okButton.Visible = $False
Publish-Form
$ReportFile = ""

If ($Global:Result -eq "OK")
{
    #check to see if employee numbers have been entered into the NewHoldIndividuals or DB Owners fields
    If  ($Script:chkNewHold.Checked -eq $True)
    {
        $HName = ($Script:txtHoldName.Text).TrimEnd(" - ")
        If (($HName -ne "40 Day Hold") -and ($HName -ne "180 Day Hold"))
        {
            $Inuse = [bool]($HoldDetails -like ("*"+(($Script:txtHoldName.Text).TrimEnd(" - "))+"*"))
            If ($Inuse -eq $True)
            {
                Do
                {
                    $Output = $wshell.Popup("This retention comment already exists, please enter a different name or change that you are adding to an existing hold.",10,"Help Info",0+32)
                    Publish-Form
                    $Inuse = [bool]($HoldDetails -like ("*"+(($Script:txtHoldName.Text).TrimEnd(" - "))+"*"))
                }While (($InUse -eq $True) -and ($Global:Result -eq "OK"))
            }
        }

        If ($Global:Result -eq "OK")
        {
            $Script:Recap = ""
            If ($Script:chkNewHold.Checked -eq $True)
            {
                Build-RetentComment
                write-host "`tEnable staff for new hold"
                write-host "`t          New Hold Name: " $Script:txtHoldName.Text
                write-host "`t This Retention Comment: " $Script:RetCmt
                write-host "`t          Staff to Hold: " $Script:txtRelENos.Text
                write-host "`t        DBOwners to Add: " $Script:txtRelDBOwner.Text
                write-host "`t      Attorney Assigned: " $Script:txtAttorney.Text
                If ($Script:chkStdHold.Checked -eq $False)
                {
                    write-host "`t         HRRep Assigned: " $Script:txtHRRep.Text
                }
                write-host "`t            Silent Hold: " $Script:chkSilent.Checked

                If (($Script:txtRelENos.Text -notlike "Enter*") -and ($Script:txtRelENos.Text.Length -gt 0))
                {
                    $EnabSel = ($Script:txtRelENos.Text -replace(" ","")) -split ","
                    write-host "`Number of staff to be Enabled: " $EnabSel.Count
                    $NewHold = "*" + $Script:txtHoldName.Text.TrimEnd(" - ") + "*"
                    Foreach ($ENo in $EnabSel)
                    {
                        If (($ENo.length -ge 5) -and ($ENo -notlike "Enter*"))
                        {
                            $ADExists = $False
                            $StrDN = ""
                            $Script:StoredCmt = ""
                            $ErrorActionPreference = "SilentlyContinue"
                            $ADExists = [bool]($Script:u = Get-ADUser $ENo -Properties *)
                            Build-RetentComment
                            $strdn = (Get-ADUser $ENo).DistinguishedName
                            $ErrorActionPreference = "Continue"
                            $ReportFile = "e:\Automation\LegalHold\Report\LegalHoldDetails-" + $ENo + ".csv"
                            $MbxExists = [bool]($Script:Mbx = get-mailbox $ENo -ErrorAction SilentlyContinue)
                            If (($ADExists -eq $False) -and ($MbxExists -eq $False))
                            {
                                write-host "`nProcessing Hold for: " $ENo -ForegroundColor Cyan
                                write-host "`tNo AD Account or O365 Mailbox found for this user.  Try to restore the account before proceeding." -ForegroundColor Red
                                $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - No AD Account or Mailbox found for this user`n"
                            }
                            else
                            {
                                GetAcctInfo
                                write-host "`nProcessing Hold for: " $ENo - $Script:u.cn -ForegroundColor Cyan
                                If ($Script:u.ExtensionAttribute14 -like $NewHold)
                                {
                                    write-host "`tThis individual is already enabled for this hold (LINE 1382)" -ForegroundColor Red
                                    $LineToWrite = $WhoAmI + "`t" + "This individual is already enabled for this hold"
                                    WriteReportEvent
                                    $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - No change made - already enabled for this hold`n"
                                }
                                else
                                {
                                    Check-ReportFile
                                    Build-RetentComment
                                    Execute-Hold
                                    If ($ADExists -eq $True)
                                    {
                                        $cnt = 0
                                        write-host "`tVerifying that the hold details have been added...." -NoNewline
                                        Do
                                        {
                                            write-host "...." -NoNewline
                                            start-sleep -Seconds 7
                                            $Script:u = Get-ADUser $ENo -Properties *
                                            $Cnt++
                                        } while (($Script:u.ExtensionAttribute14 -ne $NewCmt) -and ($cnt -le 10))
                                        $Script:ADCmt = $Script:u.ExtensionAttribute14
                                    }
                                    If ($cnt -ge 10)
                                    {
                                        write-host "Check the retention comment on this account" -ForegroundColor Red
                                    }
                                    $LineToWrite = $WhoAmI + "`t" + "Account Details After Added New Hold: "
                                    WriteReportEvent
                                    $Mbx = Get-Mailbox $ENo -ErrorAction SilentlyContinue
                                    Log-AcctDetails
                                    write-host "`n`tEnabled staff for new hold" $ENo -ForegroundColor Green

                                    # Add details to the LegalHoldReport
                                    #Enabled,EmpNo,Name,EmpType,#ofHolds,HoldName,HoldEnabled,TermDate
                                    $Text = "{0},{1},{2},{3},{4},{5},{6},{7}" -f "New",$ENo,$Script:u.cn,$Script:Mbx.CustomAttribute1,"",$Script:NewHoldCmt,"",""
                                    Out-File -FilePath $ActFile -InputObject $Text -Append
                                }
                            }
                        }
                    }
                }
                else
                {
                    write-host "`tNo users to add as an active legal hold" -ForegroundColor Red
                }
                
                If (($Script:txtRelDBOwner.Text -notlike "Enter*") -and ($Script:txtRelDBOwner.Text.Length -gt 0))
                {
                    $EnabSel = ($Script:txtRelDBOwner.Text -replace(" ","")) -split ","
                    write-host "`nNumber of DBOwners to be Enabled: " $EnabSel.Count -ForegroundColor Cyan
                    $DBOwner = "Yes"
                    Foreach ($ENo in $EnabSel)
                    {
                        $DBOFile = import-csv "E:\O365AdminShared\Data\LegalDBOwners.csv" | Where-Object {$_.EmpNo -eq $ENo}
                        If (($DBOfile.EmpNo -eq $ENo) -and ($DBOFile.CaseName -eq (($Script:txtHoldName.Text).TrimEnd(" - "))))
                        {
                            write-host "`tDetails already stored for this user and this case" -ForegroundColor Red
                            $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - Individual already listed as a dBOwner for this legal hold`n"
                        }
                        else
                        {
                            $ReportFile = "e:\Automation\LegalHold\Report\LegalHoldDetails-" + $ENo + ".csv"
                            Check-ReportFile
                            $ErrorActionPreference = "SilentlyContinue"
                            $ADExists = [bool]($Script:u = Get-ADUser $ENo -Properties *)
                            $strDN = (Get-ADUser $ENo).DistinguishedName
                            $MbxExists = [bool]($Script:Mbx = get-mailbox $ENo -ErrorAction SilentlyContinue)
                            $ErrorActionPreference = "Continue"
                            If ($ADExists -eq $True)
                            {
                                $LineToWrite = $WhoAmI + "`t" + "Adding User as a dbOwner to the LegalDBOwners.csv file"
                                WriteReportEvent
                                $LineToWrite = $WhoAmI + "`t" + "Preservation Details           :`t" + $Script:txtHoldName.Text
                                WriteReportEvent
                                #Add to dbOwner File
                                If ($Name -like "*_*")
                                {
                                    $Loc = $Name.Indexof("_")
                                    $Name = $Name.Substring(0,$Loc)
                                }
                                $LineToWrite = '{0},{1},{2},{3}' -f $ENo,$Script:u.cn,(($Script:txtHoldName.Text).TrimEnd(" - ")),""
                                Out-File -FilePath "E:\O365AdminShared\Data\LegalDBOwners.csv" -InputObject $LineToWrite -Append
                            }
                            write-host "`tAdd DBOwner for new hold" $ENo
                            $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - Added individual as a dBOwner for this legal hold`n"
                            $Text = "{0},{1},{2},{3},{4},{5},{6},{7}" -f "DBOwner",$ENo,$Script:u.cn,"","",(($Script:txtHoldName.Text).TrimEnd(" - ")),"",""
                            Out-File -FilePath $ActFile -InputObject $Text -Append
                        }
                    }
                }
                else
                {
                    write-host "`tNo DBOwners identified" -ForegroundColor Red
                }
            }
            write-host "`n`nChanges made for this hold activity:" -ForegroundColor Green
            $Script:Recap
        }
        else
        {
            write-host "Process cancelled" -ForegroundColor Red
        }
    }

    If ($Script:chkExistHold.Checked -eq $True)
    {
    #    Enable additional staff for an existing legal hold
        write-host "Enable additional staff for existing hold"
        write-host "        Hold Name: " $Script:txtActHolds.SelectedItem
        write-host "    Staff to Hold: " $Script:txtRelENos.Text
        write-host "  DBOwners to Add: " $Script:txtRelDBOwner.Text
        write-host "Attorney Assigned: " $Script:txtAttorney.Text
        If ($Script:chkStdHold.Checked -eq $False)
        {
            write-host "   HRRep Assigned: " $Script:txtHRRep.Text
        }
        write-host "      Silent Hold: " $Script:chkSilent.Checked

        $Script:Recap = ""
        $HName = (($Script:txtActHolds.SelectedItem).TrimEnd(" - "))
        $EnabSel = ($Script:txtRelENos.Text -replace(" ","")) -split ","
        write-host "Number of staff to be Enabled: " $EnabSel.Count
        $NewHold = "*" + $Script:txtActHolds.SelectedItem.TrimEnd(" - ") + "*"
        Foreach ($ENo in $EnabSel)
        {
            If (($ENo.length -ge 5) -and ($ENo -notlike "Enter*"))
            {
                $ADExists = $False
                $StrDN = ""
                $Script:StoredCmt = ""
                $Script:u = ""
                $ErrorActionPreference = "SilentlyContinue"
                $ADExists = [bool]($Script:u = Get-ADUser $ENo -Properties *)
                $strDN = (Get-ADUser $ENo).DistinguishedName
                $ErrorActionPreference = "Continue"
                $MbxExists = [bool](get-mailbox $ENo -ErrorAction SilentlyContinue)
                $Script:Mbx = get-mailbox $ENo
                write-host "`nProcessing Hold for: " $ENo - $Script:u.cn -ForegroundColor Cyan
                $ReportFile = "e:\Automation\LegalHold\Report\LegalHoldDetails-" + $ENo + ".csv"
                Check-ReportFile
                GetAcctInfo
                If ($Script:ADCmt -like $NewHold)
                {
                    write-host "`tThis individual is already enabled for this hold (LINE 1528)" -ForegroundColor Red
                    $LineToWrite = $WhoAmI + "`t" + "This individual is already enabled for this hold"
                    WriteReportEvent
                    $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - No change made - already enabled for this hold`n"
                }
                else
                {
                    Build-RetentComment
                    Execute-Hold
                    If ($ADExists -eq $True)
                    {
                        $cnt = 0
                        write-host "`tVerifying that the hold details have been added...." -NoNewline
                        Do
                        {
                            write-host "...." -NoNewline
                            start-sleep -Seconds 7
                            $Script:u = Get-ADUser $ENo -Properties *
                            $Cnt++
                        } while (($Script:u.ExtensionAttribute14 -ne $NewCmt) -and ($cnt -le 10))
                        $Script:ADCmt = $Script:u.ExtensionAttribute14
                    }
                    If ($cnt -ge 10)
                    {
                        write-host "`n`tRetention Comment not updated - manual check needed" -ForegroundColor Red
                    }
                    else
                    {
                        write-host "`n`tNew Retention Comment: " $Script:u.ExtensionAttribute14 -ForegroundColor Green
                    }
                    $LineToWrite = $WhoAmI + "`tAccount Details After Added New Hold: "
                    WriteReportEvent
                    $Mbx = get-mailbox $ENo -ErrorAction SilentlyContinue
                    Log-AcctDetails
                    write-host "`tEnabled staff for existing hold" $ENo -ForegroundColor Green
                    
                    # Add details to the LegalHoldReport
#                    $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - Added additional staff for existing hold`n"
                    #Enabled,EmpNo,Name,EmpType,#ofHolds,HoldName,HoldEnabled,TermDate
                    $Text = "{0},{1},{2},{3},{4},{5},{6},{7}" -f "New",$ENo,$Script:u.cn,$Script:Mbx.CustomAttribute1,"",$Script:txtActHolds.SelectedItem,"",""
                    Out-File -FilePath $ActFile -InputObject $Text -Append

                    write-host "`tAdded additional staff for existing hold" $ENo
                }
#                write-host "`tAdded additional staff for existing hold" $ENo
            }
        }
#        write-host "`n`nChanges made for this hold activity:" -ForegroundColor Green
#        $Script:Recap

        If (($Script:txtRelDBOwner.Text -notlike "Enter*") -and ($Script:txtRelDBOwner.Text.Length -gt 0))
        {
            $EnabSel = ($Script:txtRelDBOwner.Text -replace(" ","")) -split ","
            write-host "Number of DBOwners to be Enabled: " $EnabSel.Count
            Foreach ($ENo in $EnabSel)
            {
                If (($Script:txtRelDBOwner.Text -notlike "Enter*") -and ($Script:txtRelDBOwner.Text.Length -gt 0))
                {
                    $EnabSel = ($Script:txtRelDBOwner.Text -replace(" ","")) -split ","
                    $DBOwner = "Yes"
                    Foreach ($ENo in $EnabSel)
                    {
                        $DBOFile = import-csv "E:\O365AdminShared\Data\LegalDBOwners.csv" | Where-Object {$_.EmpNo -eq $ENo}
                        If (($DBOfile.EmpNo -eq $ENo) -and ($DBOFile.CaseName -eq (($Script:txtHoldName.Text).TrimEnd(" - "))))
                        {
                            write-host "`tDetails already stored for this user and this case" -ForegroundColor Red
                            $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - Individual already listed as a dBOwner for this legal hold`n"
                        }
                        else
                        {
                            $ReportFile = "e:\Automation\LegalHold\Report\LegalHoldDetails-" + $ENo + ".csv"
                            Check-ReportFile
                            $ErrorActionPreference = "SilentlyContinue"
                            $ADExists = [bool]($Script:u = Get-ADUser $ENo -Properties *)
                            $strDN = (Get-ADUser $ENo).DistinguishedName
                            $MbxExists = [bool]($Script:Mbx = get-mailbox $ENo -ErrorAction SilentlyContinue)
                            $ErrorActionPreference = "Continue"
                            If ($ADExists -eq $True)
                            {
                                $LineToWrite = $WhoAmI + "`t" + "Adding User as a dbOwner to the LegalDBOwners.csv file"
                                WriteReportEvent
                                $LineToWrite = $WhoAmI + "`t" + "Preservation Details           :`t" + $Script:RetCmt
                                WriteReportEvent
                                #Add to dbOwner File
                                If ($Name -like "*_*")
                                {
                                    $Loc = $Name.Indexof("_")
                                    $Name = $Name.Substring(0,$Loc)
                                }
                                $LineToWrite = '{0},{1},{2},{3}' -f $ENo,$Script:u.cn,$HName,""
                                Out-File -FilePath "E:\O365AdminShared\Data\LegalDBOwners.csv" -InputObject $LineToWrite -Append
                            }
                            write-host "`tAdd additional DBOwner for existing hold" $ENo
                            $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - Added individual as a dBOwner for this legal hold`n"
                            $Text = "{0},{1},{2},{3},{4},{5},{6},{7}" -f "DBOwner",$ENo,$Script:u.cn,"","",$Script:RetCmt,"",""
                            Out-File -FilePath $ActFile -InputObject $Text -Append
                        }
                    }
                }
#                write-host "`n`nChanges made for this hold activity:" -ForegroundColor Green
#                $Script:Recap
            }
        }
        else
        {
            write-host "`tNo changes to existing DBOwner's made" -ForegroundColor Red
        }

        If ($Script:Recap.Length -gt 0)
        {
            write-host "`n`nChanges made for this hold activity:" -ForegroundColor Green
            $Script:Recap
        }
        else
        {
            write-host "`n`nNo changes made for this hold activity:" -ForegroundColor Yellow
        }
    }

    If ($Script:chkStdRel.Checked -eq $True)
    {
    #    Release from a legal hold
        $Script:Recap = ""
        If (($Script:txtRelENos.Text -notlike "Enter*" ) -and ($Script:txtRelENos.Text.Length -gt 0))
        {
            write-host "`nReleasing from Legal Hold - " $Script:txtActHolds.SelectedItem -ForegroundColor Cyan
      
            If ($Script:txtRelENos.Text -eq "All")
            {
                #Ths section does the clean-up for individuals on this hold
                $RelSel = $Script:txtENos.Items
                write-host "                 Staff to Release: " $RelSel
                write-host "       Number of staff to Release: " $RelSel.Count
                Foreach ($u in $RelSel)
                {
                    $ENo = $u.Substring(0,$u.Indexof(" ")) #Extact emp# and removes the name
                    $ReportFile = "e:\Automation\LegalHold\Report\LegalHoldDetails-" + $ENo + ".csv"
                    Check-ReportFile
                    write-host "`n`tReleasing employee number: " $ENo -ForegroundColor Cyan
                    Release-FromHold
                }
            }
            else
            {
                #This section does the clean-up for each individual entered
                $RelSel = ($Script:txtRelENos.Text -replace(" ","")) -split (",")
                write-host "                 Staff to Release: " $RelSel
                write-host "       Number of staff to Release: " $RelSel.Count
                Foreach ($u in $RelSel)
                {
                    $ENo = $u.Trim() #trim of leading or trailing spaces
                    $ReportFile = "e:\Automation\LegalHold\Report\LegalHoldDetails-" + $ENo + ".csv"
                    Check-ReportFile
                    write-host "`n`tReleasing employee number: " $ENo -ForegroundColor Cyan
                    Release-FromHold
                }
            }
        }

        If (($Script:txtRelDBOwner.Text -notlike "Enter*") -and ($Script:txtRelDBOwner.Text.Length -gt 0))
        {
            $DBOFile = "E:\O365AdminShared\Data\LegalDBOwners.csv"
            $DBODetails = get-content "E:\O365AdminShared\Data\LegalDBOwners.csv"
            $NewRecs = get-content "e:\O365adminshared\data\LegalDBowners.csv" | Where-Object {$_ -notmatch $Script:txtActHolds.SelectedItem}
            $RemoveRecs = get-content "e:\O365adminshared\data\LegalDBowners.csv" | Where-Object {$_ -match $Script:txtActHolds.SelectedItem}
            Rename-Item -Path $DBOFile -NewName ("LegalDBOwners-Date" + (((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", ""))) + ".old")
            If ($Script:txtRelDBOwner.Text -eq "All")
            {
                write-host "`tRelease all enabled DBOwners from this hold"
                $RelSel = $Script:txtDBOwner.Items
                write-host "            Release all dBOwner's: " $RelSel
                write-host "   Number of dBOwner's to Release: " $RelSel.Count
                Foreach ($r in $RemoveRecs)
                {
                    $ENo = $r.Substring(0,$r.Indexof(","))
                    $ReportFile = "e:\Automation\LegalHold\Report\LegalHoldDetails-" + $ENo + ".csv"
                    write-host "`tReleasing dBOwner employee number: " $ENo
                    $LineToWrite = $WhoAmI + "`t" + "Removing record in DBOwners.csv file"
                    WriteReportEvent
                    $Script:Recap = $Script:Recap + "`t" + $r + ") - Removed entry from the dBOwners.csv file`n"
                }
                #Add owners not a part of this hold into the file
                Foreach ($n in $NewRecs)
                {
                    add-content $DBOFile $n
                }
            }
            else
            {
                write-host "`tRelease selected DBOwners from this hold"
                $RelSel = ($Script:txtRelDBOwner.Text -replace(" ","")) -split(",")
                write-host "       Release Select dBOwner's: " $RelSel
                write-host " Number of dBOwner's to Release: " $RelSel.Count
                $ThisHoldOwners = get-content "e:\O365adminshared\data\legaldbowners.csv" | Where-Object {$_ -match $Script:txtActHolds.SelectedItem}
                Foreach ($ENo in $RelSel)
                {
                    write-host "`tReleasing dbOwner employee number: " $ENo
                    $ReportFile = "e:\Automation\LegalHold\Report\LegalHoldDetails-" + $ENo + ".csv"
                    $LineToWrite = $WhoAmI + "`t" + "Removing record in DBOwners.csv file"
                    WriteReportEvent
                    #Remove Entry from DBOwners file
                    $RemoveRecs = $RemoveRecs |Where-Object {$_ -notmatch $ENo}
                    $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + ((get-ADUser $ENo).Name) + ") - Removed entry from the dBOwners.csv file`n"
                }
                #Add all recs that were not a part of this hold
                Foreach ($n in $NewRecs)
                {
                    add-content $DBOFile $n
                }
                #If dbowners remain add these back into the file
                If ($RemoveRecs.count -gt 0)
                {
                    Foreach ($r in $RemoveRecs)
                    {
                        add-content $DBOFile $r
                    }
                }
            }
#            $Script:Recap = $Script:Recap + "`t" + $ENo + " (" + $Script:u.cn + ") - Removed entry from the dBOwners.csv file`n"
        }

        If ((($Script:txtRelENos.Text -like "ALL*" ) -or ($Script:txtNoENos.Text -eq 0)) -and (($Script:txtRelDBOwner.Text -like "ALL*") -or ($Script:txtNoDBOwner.Text -eq 0)))
        {
            If (($Script:txtActHolds.SelectedItem -notlike "40 Day Hold - ") -and ($Script:txtActHolds.SelectedItem -notlike "180 Day Hold - "))
            {
                Remove-LegalHold
                $Script:Recap = $Script:Recap + "`tRemoved " + $Script:txtActHolds.SelectedItem + " from the file with all active Legal Holds`n"
            }
            else
            {
                write-host "This hold was not removed from the file as it is the generic entry for 40 Day or 180 Day Holds" -ForegroundColor Yellow
            }
        }

        write-host $Script:Recap
    }

    If ($Script:chkRenameHold.Checked -eq $True)
    {
        Rename-ExistingHold
    }
}
else
{
    write-host "Process cancelled" -ForegroundColor Red
}
} while ($Global:Result -eq "OK")