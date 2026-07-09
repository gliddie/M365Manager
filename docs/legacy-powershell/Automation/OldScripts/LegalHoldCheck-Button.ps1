<#   
================================================================================ 
 Name: Used SampleForm.ps1 from Dan Stolts "ITProGuru" at http://ITProGuru.com/Scripts as a template
 
 New Legal Hold process replaces LegalHoldCheckps1 
 ================================================================================ 
 #
 #  05/17/2020 - SAG:Added new process using forms to perform changes$ConfDisable
 #  05/29/2020 - SAG:Fixed issues with enabling NewHolds
 #  06/02/2020 - SAG:Changed Get-Mailbox to Get-EXOMailbox
 #  07/01/2020 - SAG:Added additional buttons for and details (InPlaceHolds,Hide/Unhide,RemovingMBXAccess)
 #  07/02/2020 - SAG:Added Reporting Features
 #  10/26/2020 - SAG:Repaired comments for accounts with no AD Account
 #  10/27/2020 - SAG:Added code so that access is not removed unless all holds have been released
 #  11/16/2020 - SAG:Moved Get-ENo function to the O365AdminMenu.ps1 script
#>  

#Forms Functions
Function Add-ActionBoxes
{
    ## New Legal Hold CheckBox         
    $Global:chkEnable = New-Object Windows.Forms.checkbox 
        $Global:chkEnable.Left = 570; $Global:chkEnable.Width = 200; $Global:chkEnable.Top = 60  
        $Global:chkEnable.Text = "New Legal Hold" 
        $Global:chkEnable.Checked = $Global:chkEnable.Checked   # set a default value 
        $Global:chkEnable.TabIndex = 0
        $Global:form.Controls.Add($Global:chkEnable) 
        # Obtain Value with: $Global:chkThis.Checked

    ## New 40 Day Hold CheckBox         
    $Global:chk40Day = New-Object Windows.Forms.checkbox 
        $Global:chk40Day.Left = 570; $Global:chk40Day.Width = 200; $Global:chk40Day.Top = 90  
        $Global:chk40Day.Text = "New 40 Day Hold" 
        $Global:chk40Day.Checked = $Global:chk40Day.Checked   # set a default value 
        $Global:chk40Day.TabIndex = 1
         
        $Global:form.Controls.Add($Global:chk40Day) 
        # Obtain Value with: $Global:chk40Day.Checked

    $TopAlignment = 120

    If ($null -ne $Global:inf.LitigationHoldDate)
    {
        ## Disable Comment Hold CheckBox
        $Global:chkDisable = New-Object Windows.Forms.checkbox 
        $Global:chkDisable.Left = 570; $Global:chkDisable.Width = 200; $Global:chkDisable.Top = $TopAlignment 
        $Global:chkDisable.Text = "Disable Legal Hold" 
        $Global:chkDisable.Checked = $Global:chkDisable.Checked
        $Global:chkDisable.TabIndex = 2
        $Global:form.Controls.Add($Global:chkDisable)
        $TopAlignment = $TopAlignment + 30
    }

    If ($Global:u.ExtensionAttribute14.value.length -ne 0)
    {
        ## Modify Hold CheckBox
        $Global:chkModify = New-Object Windows.Forms.checkbox 
        $Global:chkModify.Left = 570; $Global:chkModify.Width = 200; $Global:chkModify.Top = $TopAlignment 
        $Global:chkModify.Text = "Modify Comment" 
        $Global:chkModify.Checked = $Global:chkModify.Checked
        $Global:chkModify.TabIndex = 3
        $Global:form.Controls.Add($Global:chkModify)
        $TopAlignment = $TopAlignment + 30
    }

    If (($Global:u.ExtensionAttribute14.value.length -ne 0) -and ($Global:u.ExtensionAttribute14.value -like "*,*"))
    {
        ## Remove Comment Hold CheckBox
        $Global:chkRemove = New-Object Windows.Forms.checkbox 
        $Global:chkRemove.Left = 570; $Global:chkRemove.Width = 200; $Global:chkRemove.Top = $TopAlignment  
        $Global:chkRemove.Text = "Remove Comment" 
        $Global:chkRemove.Checked = $Global:chkRemove.Checked
        $Global:chkRemove.TabIndex = 4
        $Global:form.Controls.Add($Global:chkRemove)
        $TopAlignment = $TopAlignment + 30
    }

    If ($Global:MbxAccessForm -ne "(None)")
    {
        ## Remove Access to Mailbox
        $Global:chkAccess = New-Object Windows.Forms.checkbox 
        $Global:chkAccess.Left = 570; $Global:chkAccess.Width = 200; $Global:chkAccess.Top = $TopAlignment  
        $Global:chkAccess.Text = "Remove Access" 
        $Global:chkAccess.Checked = $Global:chkAccess.Checked
        $Global:chkAccess.TabIndex = 5
        $Global:form.Controls.Add($Global:chkAccess)
        $TopAlignment = $TopAlignment + 30
    }

    If (($Global:u.ExtensionAttribute1 -like "Ex-*"))
    {
        If ($Global:inf.HiddenFromAddressListsEnabled -eq $True)
        {
            ## Unhide from AddressBook Hold CheckBox
            $Global:chkUnhide = New-Object Windows.Forms.checkbox 
            $Global:chkUnhide.Left = 570; $Global:chkUnhide.Width = 200; $Global:chkUnhide.Top = $TopAlignmenT
            $Global:chkUnhide.Text = "Unhide Account" 
            $Global:chkUnhide.Checked = $Global:chkUnhide.Checked
            $Global:chkUnhide.TabIndex = 6
            $Global:form.Controls.Add($Global:chkUnhide)
            $TopAlignment = $TopAlignment + 30
        }
        else
        {
            ## Hide from AddressBook Hold CheckBox
            $Global:chkHide = New-Object Windows.Forms.checkbox 
            $Global:chkHide.Left = 570; $Global:chkHide.Width = 200; $Global:chkHide.Top = $TopAlignmenT
            $Global:chkHide.Text = "Hide Account" 
            $Global:chkHide.Checked = $Global:chkUnhide.Checked
            $Global:chkHide.TabIndex = 7
            $Global:form.Controls.Add($Global:chkHide)
            $TopAlignment = $TopAlignment + 30
        }

        If ($Global:strUserPath -notlike "*_Extended*")
        {
            ## Move to ExtendedHold CheckBox
            $Global:chkMove = New-Object Windows.Forms.checkbox 
            $Global:chkMove.Left = 570; $Global:chkMove.Width = 200; $Global:chkMove.Top = $TopAlignment
            $Global:chkMove.Text = "Move to ExtendedHold" 
            $Global:chkMove.Checked = $Global:chkMove.Checked
            $Global:chkMove.TabIndex = 8
            $Global:form.Controls.Add($Global:chkMove)
            $TopAlignment = $TopAlignment + 30
        }
    }

    ## Create Report
    $Global:chkReport = New-Object Windows.Forms.checkbox 
        $Global:chkReport.Left = 570; $Global:chkReport.Width = 200; $Global:chkReport.Top = $TopAlignment
        $Global:chkReport.Text = "Generate Report" 
        $Global:chkReport.Checked = $Global:chkReport.Checked   # set a default value 
        $Global:chkReport.TabIndex = 9
        $Global:form.Controls.Add($Global:chkReport) 
        # Obtain Value with: $Global:chk40Day.Checked
}

Function Add-SilentHHHoldBoxes
{
    #Add SilentHold Checkbox
    $Global:chkSilent = New-Object Windows.Forms.checkbox 
        $Global:chkSilent.Left = 160; $Global:chkSilent.Width = 200; $Global:chkSilent.Top = 410  
        $Global:chkSilent.Text = "Enable Silent Hold"
        $Global:chkSilent.Checked = $False
        If ($Global:NewCmt -like "NonDisclosed*")
        {
            $Global:chkSilent.Checked = $True
        }
        $Global:chkSilent.Checked = $Global:chkSilent.Checked   # set a default value 
        $Global:chkSilent.TabIndex = $Global:TabIndex
        $Global:form.Controls.Add($Global:chkSilent) 
        # Obtain Value with: $Global:chkSilent.Checked
        $Global:TabIndex++

    #Add HardwareHold
    $Global:chkHHold = New-Object Windows.Forms.checkbox 
        $Global:chkHHold.Left = 160; $Global:chkHHold.Width = 200; $Global:chkHHold.Top = 440  
        $Global:chkHHold.Text = "Enable Hardware Hold"
        $Global:chkHHold.Checked = $True
        If ($Global:NewCmt -like "*NoHHold")
        {
            $Global:chkHHold.Checked = $False
        } 
        $Global:chkHHold.Checked = $Global:chkHHold.Checked   # set a default value 
        $Global:chkHHold.TabIndex = $Global:TabIndex
        $Global:form.Controls.Add($Global:chkHHold) 
        # Obtain Value with: $Global:chkHHold.Checked
        $Global:TabIndex++
}

Function Build-DefaultForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Legal Hold Details" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 740 ; $form.Height = 640  # Make the form wider 
    
    ## Employee ID
    $Global:lblHost = New-Object System.Windows.Forms.Label   
        $Global:lblHost.Text = "Employee No:" 
        $Global:lblHost.Top = 10 ; $Global:lblHost.Left = 5; $Global:lblHost.Width=150 ;$Global:lblHost.AutoSize = $true 
        $form.Controls.Add($Global:lblHost)    # Add to Form 
        # 
        $Global:txtHost = New-Object Windows.Forms.TextBox
        $Global:txtHost.ReadOnly = $True;
        $Global:txtHost.Top = 10; $Global:txtHost.Left = 160; $Global:txtHost.Width = 120;  
        $Global:txtHost.Text = $ENo
        $Global:form.Controls.Add($Global:txtHost)    # Add to Form
        $Global:InputFocus = $Global:txtHost

    $Top = $Top + 30
    $Script:ButActName = New-Object Windows.Forms.Button
        $Script:ButActName.Location = New-object System.Drawing.Size($LeftInput,$Top)
        $Script:ButActName.Size = new-Object System.Drawing.Size(200,20)
        $Script:ButActName.Text = "Get Account Details"
        $Script:form.Controls.Add($Script:ButActName)
        $Script:ButActName.Add_Click({
            Get-AccountDetails
        })        
 
    ## Employee Name
    $Global:lblName = New-Object System.Windows.Forms.Label   
        $Global:lblName.Text = "Employee Name:"  
        $Global:lblName.Top = 40 ; $Global:lblName.Left = 5; $Global:lblName.Width=150 ;$Global:lblName.AutoSize = $true 
        $form.Controls.Add($Global:lblName)    # Add to Form 
        # 
        $Global:txtName = New-Object Windows.Forms.TextBox
        $Global:txtName.ReadOnly = $True;
        $Global:txtName.Top = 40; $Global:txtName.Left = 160; $Global:txtName.Width = 120;  
        $Global:txtName.Text = $Global:inf.DisplayName
        $Global:form.Controls.Add($Global:txtName)    # Add to Form 

    ## Employee $Type
    $Global:lblType = New-Object System.Windows.Forms.Label   
        $Global:lblType.Text = "Employee Type:"  
        $Global:lblType.Top = 70 ; $Global:lblType.Left = 5; $Global:lblType.Width=150 ;$Global:lblType.AutoSize = $true 
        $form.Controls.Add($Global:lblType)    # Add to Form 
        # 
        $Global:txtType = New-Object Windows.Forms.TextBox
        $Global:txtType.ReadOnly = $True;
        $Global:txtType.Top = 70; $Global:txtType.Left = 160; $Global:txtType.Width = 120;  
        $Global:txtType.Text = $Global:u.ExtensionAttribute1.value
        $Global:form.Controls.Add($Global:txtType)    # Add to Form 

    ## ADContainer
    $Global:lblContainer = New-Object System.Windows.Forms.Label   
        $Global:lblContainer.Text = "AD Container:"  
        $Global:lblContainer.Top = 100 ; $Global:lblContainer.Left = 5; $Global:lblContainer.Width=150 ;$Global:lblContainer.AutoSize = $true 
        $form.Controls.Add($Global:lblContainer)    # Add to Form 
        # 
        $Global:txtContainer = New-Object Windows.Forms.TextBox  
        $Global:txtContainer.ReadOnly = $True; 
        $Global:txtContainer.Top = 100; $Global:txtContainer.Left = 160; $Global:txtContainer.Width = 370;
        $Container = $Global:strUserPath -Replace("LDAP://CN=","")
        $Container = $Container -Replace(",OU=","/")
        $Container = $Container -Replace(",DC=","/")
        $Global:txtContainer.Text = ($Container)
        $Global:form.Controls.Add($Global:txtContainer)    # Add to Form 

    ## Hold Enabled Date
    $Global:lblDate = New-Object System.Windows.Forms.Label   
        $Global:lblDate.Text = "Hold Enabled Date:"  
        $Global:lblDate.Top = 130 ; $Global:lblDate.Left = 5; $Global:lblDate.Width=150 ;$Global:lblDate.AutoSize = $true 
        $form.Controls.Add($Global:lblDate)    # Add to Form 
        # 
        $Global:txtDate = New-Object Windows.Forms.TextBox  
        $Global:txtDate.ReadOnly = $True;
        $Global:txtDate.Top = 130; $Global:txtDate.Left = 160; $Global:txtDate.Width = 120;  
        $Global:txtDate.Text = $Global:inf.LitigationHoldDate
        $Global:form.Controls.Add($Global:txtDate)    # Add to Form 
       # Obtain Value with: $Global:txtDate.Text

    ## Hold Enabled By
    $Global:lblBy = New-Object System.Windows.Forms.Label   
        $Global:lblBy.Text = "Hold Enabled By:"  
        $Global:lblBy.Top = 130; $Global:lblBy.Left = 310; $Global:lblBy.Width=100 ;$Global:lblBy.AutoSize = $true  
        $form.Controls.Add($Global:lblBy)    # Add to Form 
        # 
        $Global:txtBy = New-Object Windows.Forms.TextBox  
        $Global:txtBy.ReadOnly = $True; 
        $Global:txtBy.Top = 130; $Global:txtBy.Left = 410; $Global:txtBy.Width = 120;  
        $Global:txtBy.Text = $Global:inf.LitigationHoldOwner
        $Global:form.Controls.Add($Global:txtBy)    # Add to Form 
       # Obtain Value with: $Global:txtBy.Text

    ## Hold ADOjbect Protected
    $Global:lblProtect = New-Object System.Windows.Forms.Label   
        $Global:lblProtect.Text = "AD Object Protected:"  
        $Global:lblProtect.Top = 160 ; $Global:lblProtect.Left = 5; $Global:lblProtect.Width=150 ;$Global:lblProtect.AutoSize = $true 
        $form.Controls.Add($Global:lblProtect)    # Add to Form 
        # 
        $Global:txtProtect = New-Object Windows.Forms.TextBox  
        $Global:txtProtect.ReadOnly = $True; 
        $Global:txtProtect.Top = 160; $Global:txtProtect.Left = 160; $Global:txtProtect.Width = 120;
        $Global:txtProtect.Text = $Global:UsrDetails.ProtectedFromAccidentalDeletion
        $Global:form.Controls.Add($Global:txtProtect)    # Add to Form 
       # Obtain Value with: $Global:txtProtect.Text

    ## Hold ADOjbect Hidden
    $Global:lblHide = New-Object System.Windows.Forms.Label   
        $Global:lblHide.Text = "AD Object Hidden:"  
        $Global:lblHide.Top = 190 ; $Global:lblHide.Left = 5; $Global:lblHide.Width=160 ;$Global:lblHide.AutoSize = $true 
        $form.Controls.Add($Global:lblHide)    # Add to Form 
        # 
        $Global:txtHide = New-Object Windows.Forms.TextBox  
        $Global:txtHide.ReadOnly = $True;
        $Global:txtHide.Top = 190; $Global:txtHide.Left = 160; $Global:txtHide.Width = 120;
        $Global:txtHide.Text = $Global:inf.HiddenFromAddressListsEnabled
        $Global:form.Controls.Add($Global:txtHide)    # Add to Form 
       # Obtain Value with: $Global:txtHide.Text

    ## Accounts Granted Access
    $Global:lblAccess = New-Object System.Windows.Forms.Label   
        $Global:lblAccess.Text = "Accounts with Access:"  
        $Global:lblAccess.Top = 220 ; $Global:lblAccess.Left = 5; $Global:lblAccess.Width=160 ;$Global:lblAccess.AutoSize = $true 
        $form.Controls.Add($Global:lblAccess)    # Add to Form 
        # 
        $Global:txtAccess = New-Object Windows.Forms.TextBox  
        $Global:txtAccess.ReadOnly = $True; 
        $Global:txtAccess.Top = 220; $Global:txtAccess.Left = 160; $Global:txtAccess.Width = 370; 
        $Global:txtAccess.Text = $Global:MbxAccessForm
        $Global:form.Controls.Add($Global:txtAccess)    # Add to Form 
 
     ##InPlace Holds
     $Global:lblInPlace = New-Object System.Windows.Forms.Label   
        $Global:lblInPlace.Text = "InPlace Holds:"  
        $Global:lblInPlace.Top = 250 ; $Global:lblInPlace.Left = 5; $Global:lblInPlace.Width=160 ;$Global:lblInPlace.AutoSize = $true 
        $form.Controls.Add($Global:lblInPlace)    # Add to Form 
        # 
        $Global:txtInPlace = New-Object Windows.Forms.TextBox  
        $Global:txtInPlace.ReadOnly = $True;
        $Global:txtInPlace.Top = 250; $Global:txtInPlace.Left = 160; $Global:txtInPlace.Width = 370; 
        $Global:txtInPlace.Text = $Global:inf.InPlaceHolds
        $Global:form.Controls.Add($Global:txtInPlace)    # Add to Form           
    
    If ($Container -like "*Active Directory Account*")
    {
        $Global:StoredCmt = $Global:inf.RetentionComment
    }
    If ($Global:StoredCmt.Length -ne 0)
    {
        ## ListBox - Fill with Data From Azure Location Name 
        $Global:lblLoc = New-Object System.Windows.Forms.Label   
            $Global:lblLoc.Text = "Preservation Comment(s):"; $Global:lblLoc.Top = 280; $Global:lblLoc.Left = 5; $Global:lblLoc.Autosize = $true  
            $Global:form.Controls.Add($Global:lblLoc)  
            # Listbox for Location Name 
            $Global:locListBox = New-Object System.Windows.Forms.ListBox  
                $Global:locListBox.Top = 280; $locListBox.Left = 160; $locListBox.Height = 100; $LocListBox.Width = 370;
                $Global:locListBox.TabIndex = 1
                # we need to populate the listbox... Example: $objListBox.Items.Add("Item 1 Test Do NOT USE") 
                # in our case, we will use a call to Azure for our "list"
                If ($Global:StoredCmt.Length -ne 0)
                {
                    $LocArray = $Global:StoredCmt.split(",")
                    foreach ($element in $LocArray)
                    {
                        [void] $Global:locListBox.Items.Add($element.TrimStart())  # Add element to listbox 
                    } 
                }
                $Global:form.Controls.Add($Global:locListBox) #Add listbox to form 
                # Obtain Value with: $Global:locListBox.SelectedItem
    }

    Add-FormStandardButtons
}

Function Build-RelaseForm
{
    #Add Released By
    $Global:lblReleaseby = New-Object System.Windows.Forms.Label
        $Global:lblReleaseby.Text = "HR/Legal Rep Releasing:"  
        $Global:lblReleaseby.Top = 380 ; $Global:lblReleaseby.Left = 5; $Global:lblReleaseby.Width=150 ;$Global:lblReleaseby.AutoSize = $true
        $form.Controls.Add($Global:lblReleaseby)    # Add to Form 
        # 
        $Global:txtReleaseby = New-Object Windows.Forms.ComboBox
        $Global:txtReleaseby.TabIndex = 20 # set Tab Order 
        $Global:txtReleaseby.Top = 380; $Global:txtReleaseby.Left = 160; $Global:txtReleaseby.Width = 120;
        $Global:Attorneys = import-csv "e:\O365AdminShared\Data\Attorneys.csv" | Sort-Object Name
        Foreach ($Attorneys in $Global:Attorneys)
        {
            [void] $Global:txtReleaseby.Items.Add($Attorneys.Name)
        }
        $Global:txtReleaseby.Text = ""   # Use Corrent computer name as default 
        $Global:form.Controls.Add($Global:txtReleaseby)    # Add to Form 
       # Obtain Value with: $Global:txtReleaseby.Text

    #Add Release Date
    $Global:lblEMailDate = New-Object System.Windows.Forms.Label
        $Global:lblEMailDate.Text = "EMail Date Releasing:"
        $Global:lblEMailDate.Top = 410; $Global:lblEMailDate.Left = 5; $Global:lblEMailDate.Width=150 ;$Global:lblEMailDate.AutoSize = $true
        $form.Controls.Add($Global:lblEMailDate)    # Add to Form
        #
        $Global:txtEMailDate = New-Object Windows.Forms.TextBox
        $Global:txtEMailDate.TabIndex = 21 # set Tab Order
        $Global:txtEMailDate.Top = 410; $Global:txtEMailDate.Left = 160; $Global:txtEMailDate.Width = 120;
        $Global:txtEMailDate.Text = (get-date -F MM/dd/yyyy)   # Use Corrent computer name as default 
        $Global:form.Controls.Add($Global:txtEMailDate)    # Add to Form 
        # Obtain Value with: $Global:txtEMailDate.Text
}

Function Enable-NewHoldForm
{
    #Add Preservation Name
    $Global:lblPrevName = New-Object System.Windows.Forms.Label
        $Global:lblPrevName.Text = "Preservation Name:"  
        $Global:lblPrevName.Top = 380 ; $Global:lblPrevName.Left = 5; $Global:lblPrevName.Width=150 ;$Global:lblPrevName.AutoSize = $true
        $form.Controls.Add($Global:lblPrevName)    # Add to Form 
        # 
        If ($Global:chk40Day.Checked -eq "Checked")
        {
            $Global:txtPrevName = New-Object Windows.Forms.TextBox  
            $Global:txtPrevName.TabIndex = 19 # set Tab Order 
            $Global:txtPrevName.Top = 380; $Global:txtPrevName.Left = 160; $Global:txtPrevName.Width = 250;
            $Global:txtPrevName.Text = "40 Day Hold - "
        }
        else
        {
            $Global:txtPrevName = New-Object System.Windows.Forms.ComboBox
            $Global:txtPrevName.TabIndex = 19 # set Tab Order 
            $Global:txtPrevName.Top = 380; $Global:txtPrevName.Left = 160; $Global:txtPrevName.Width = 250;
            $Global:StoredHolds = import-csv "e:\O365AdminShared\Data\LegalHolds.csv" | Sort-Object HoldName
            $Global:AddPrevName = "Yes"
            Foreach ($StoredHolds in $Global:StoredHolds)
            {
                [void] $Global:txtPrevName.Items.Add($StoredHolds.HoldName)
            }
        }
        $Global:form.Controls.Add($Global:txtPrevName)    # Add to Form 
       # Obtain Value with: $Global:txtPrevName.Text

       $Top = 410

        If ($Global:chk40Day.Checked -eq "Checked")
        {
            #Add HRRep
            $Global:lblHRRep = New-Object System.Windows.Forms.Label
            $Global:lblHRRep.Text = "HR Representative:"  
            $Global:lblHRRep.Top = $Top ; $Global:lblHRRep.Left = 5; $Global:lblHRRep.Width=150 ;$Global:lblHRRep.AutoSize = $true
            $form.Controls.Add($Global:lblHRRep)    # Add to Form 
            # 
            $Global:txtHRRep = New-Object Windows.Forms.ComboBox  
            $Global:txtHRRep.TabIndex = 20 # set Tab Order 
            $Global:txtHRRep.Top = $Top; $Global:txtHRRep.Left = 160; $Global:txtHRRep.Width = 120;
            $Global:HRReps = import-csv "e:\O365AdminShared\Data\HRReps.csv" | Sort-Object Name
            Foreach ($HRReps in $Global:HRReps)
            {
                [void] $Global:txtHRRep.Items.Add($HRReps.Name)
            }
            $Global:form.Controls.Add($Global:txtHRRep)    # Add to Form 
            $Top = $Top + 30
        }

    #Add Attorney
    $Global:lblAttorney = New-Object System.Windows.Forms.Label
        $Global:lblAttorney.Text = "Handling Attorney:"
        $Global:lblAttorney.Top = $Top ; $Global:lblAttorney.Left = 5; $Global:lblAttorney.Width=150 ;$Global:lblAttorney.AutoSize = $true
        $form.Controls.Add($Global:lblAttorney)    # Add to Form
        #
        $Global:txtAttorney = New-Object Windows.Forms.ComboBox
        $Global:txtAttorney.TabIndex = 21 # set Tab Order
        $Global:txtAttorney.Top = $Top; $Global:txtAttorney.Left = 160; $Global:txtAttorney.Width = 120;
        $Global:Attorneys = import-csv "e:\O365AdminShared\Data\Attorneys.csv" | Sort-Object Name
        Foreach ($Attorneys in $Global:Attorneys)
        {
            [void] $Global:txtAttorney.Items.Add($Attorneys.Name)
        }
        $Global:form.Controls.Add($Global:txtAttorney)    # Add to Form 
        $Top = $Top + 30

    #Add SilentHold Checkbox
    $Global:chkSilent = New-Object Windows.Forms.checkbox 
        $Global:chkSilent.Left = 160; $Global:chkSilent.Width = 200; $Global:chkSilent.Top = $Top  
        $Global:chkSilent.Text = "Enable Silent Hold" 
        $Global:chkSilent.Checked = $Global:chkSilent.Checked   # set a default value 
        $Global:chkSilent.TabIndex = 22 
        $Global:form.Controls.Add($Global:chkSilent) 
        $Top = $Top + 30

    #Add HardwareHold
    $Global:chkHHold = New-Object Windows.Forms.checkbox 
        $Global:chkHHold.Left = 160; $Global:chkHHold.Width = 200; $Global:chkHHold.Top = $Top  
        $Global:chkHHold.Text = "Enable Hardware Hold"
        $Global:chkHHold.Checked = "True" 
        $Global:chkHHold.Checked = $Global:chkHHold.Checked   # set a default value 
        $Global:chkHHold.TabIndex = 23 
        $Global:form.Controls.Add($Global:chkHHold) 
}

Function Modify-CommentForm
{
    $Global:NewCmt = $Global:locListBox.SelectedItem

    $Global:lblCmt = New-Object System.Windows.Forms.Label   
    $Global:lblCmt.Text = "Upated Comment:"
    $Global:lblCmt.Top = 380 ; $Global:lblCmt.Left = 5; $Global:lblCmt.Width=200 ;$Global:lblCmt.AutoSize = $true 
    $form.Controls.Add($Global:lblCmt)    # Add to Form 
    # 
    $Global:txtCmt = New-Object Windows.Forms.TextBox  
    $Global:txtCmt.TabIndex = 19 # set Tab Order 
    $Global:txtCmt.Top = 380; $Global:txtCmt.Left = 160; $Global:txtCmt.Width = 400;  
    $Global:txtCmt.Text = $Global:locListBox.SelectedItem   # New Prservation Comment
    $Global:form.Controls.Add($Global:TxtCmt)    # Add to Form 

    Add-SilentHHHoldBoxes
}

Function Republish-Form
{
#show updates:
    $Global:FormRefresh = "Y"
    start-sleep -s 15
    GetAcctInfo($Global:ENo)

    If ($Global:ADExists -eq $True)
    {
        $Global:StoredCmt = $Global:u.ExtensionAttribute14.value
    }
    else
    {
        $Global:StoredCmt = (Get-Mailbox $ENo).RetentionComment
    }
    $Global:OKDetails = "Done"
    Build-DefaultForm
    Publish-Form
    $Global:FormRefresh = "N"
}

Function Update-InputFiles
{
    #check if this is a new Attorney
    $AddAttName = "Yes"
    $Global:txtAttorney.Text= ($Global:txtAttorney.Text).TrimEnd()
    $Global:txtAttorney.Text = $Global:txtAttorney.Text -replace (",",".")
    If ($Global:txtAttorney.Text -notlike "*TASK*")
    {
        Foreach ($Attorneys in $Global:Attorneys)
        {
            If ($Global:txtAttorney.Text -eq $Attorneys.Name)
            {
                $AddAttName = "No"
            }
        }
    }
    else
    {
        $AddAttName = "No"
    }

    If ($AddAttName -eq "Yes")
    {
        #Add new Attorney to file
        $LineToWrite = '"{0}"' -f $Global:txtAttorney.Text
        Out-File -FilePath "E:\O365AdminShared\Data\Attorneys.csv" -InputObject $LineToWrite -Append
    }
       
    #check if this is a new HRRep
    $AddHRRep = "Yes"
    If ($Global:txtHRRep.Text -notlike "*TASK*")
    {
        Foreach ($HRReps in $Global:HRReps)
        {
            $Global:txtHRRep.Text = ($Global:txtHRRep.Text).TrimEnd()
            If ($Global:txtHRRep.Text -eq $HRReps.Name)
            {
                $AddHRRep = "No"
            }
        }
    }
    else
    {
        $AddHRRep = "No"
    }

    If ($AddHRRep -eq "Yes")
    {
        #Add new HRRep to file
        $LineToWrite = '"{0}"' -f $Global:txtHRRep.Text
        Out-File -FilePath "E:\O365AdminShared\Data\HRReps.csv" -InputObject $LineToWrite -Append
    }       
        
    #Check if this is a new preservation activity
    $Global:AddPrevName = "Yes"
    If ($Global:txtPrevName.Text -ne "40 Day Hold - ")
    {
        Foreach ($StoredHolds in $Global:StoredHolds)
        {
            If ($Global:txtPrevName.Text -eq $StoredHolds.HoldName)
            {
                $Global:AddPrevName = "No"
            }
        }

        If ($Global:AddPrevName -eq "Yes")
        {
            #Add new Retention comment to file
            $LineToWrite = '"{0}"' -f ($Global:txtPrevName.Text).TrimEnd()
            Out-File -FilePath "E:\O365AdminShared\Data\LegalHolds.csv" -InputObject $LineToWrite -Append
        }
    }
}

# Change Functions
Function Build-RetentComment
{
    Update-InputFiles
    
    $Global:txtPrevName.Text = ($Global:txtPrevName.Text).TrimEnd()
    #Enable Hold or Add Additional Comment
    If ($Global:chkSilent.Checked -eq $True)
    {
        $Global:NewBld = "NonDisclosed - " + $Global:txtPrevName.Text + $Global:txtHRRep.Text + "\" + $Global:txtAttorney.Text + "\"
    }
    else
    {
        $Global:NewBld = $Global:txtPrevName.Text + $Global:txtHRRep.Text + "\" + $Global:txtAttorney.Text + "\"
    }
    If ($chkHHold.Checked -eq $True)
    {
        $Global:NewBld = $Global:NewBld + "HHold"
    }
    else
    {
        $Global:NewBld = $Global:NewBld + "NoHHold"
    }

    If ($Global:StoredCmt.Contains($Global:NewBld))
    {
        $LineToWrite = $WhoAmI + "`t" + "Retention Already Configured   :`t" + $Global:BldCmt
        WriteReportEvent
    }
    else
    {
        If (($Global:u.ExtensionAttribute14.value.length -eq 0) -or ($Global:u.ExtensionAttribute14.value -like "Emergency*") -or ($Global:u.ExtensionAttribute14.value -like "Standard*"))
        {
            $Global:NewCmt = $Global:NewBld
        }
        else
        {
            $Global:NewCmt =  $Global:StoredCmt + ", " + $Global:NewBld
        }
        $LineToWrite = $WhoAmI + "`t" + "New Retention Activity         :`t" + $Global:NewBld
        WriteReportEvent
        Commit-CommentChange
        Execute-Hold
    }
}

Function Commit-CommentChange
{
    If ($Global:NewCmt -ne $Global:StoredCmt)
    {
#   Updating Comment
        If ($Global:ADExists -ne $True)
        {
           set-mailbox $Global:ENo -RetentionComment (($Global:StoredCmt).Replace($Global:locListBox.SelectedItem,$Global:NewCmt))
        }
        else
        {
            $Global:u.ExtensionAttribute14 = $Global:NewCmt
            $Global:u.CommitChanges()
        }
    }
    else 
    {
        $Output = $wshell.Popup("No changes needed.",0,"No Changes Needed",0+32)    
        $LineToWrite = $WhoAmI + "`t" + "No changes Needed              :"
        WriteReportEvent
    }
}

Function Disable-Hold
{
    $ConfDisable = 6
    If ($Global:u.ExtensionAttribute14.value -like "*,*")
    {
        $ConfDisable = $wshell.Popup("There are multiple retention comments on this account are all these hold activities being released?",0,"Multiple Holds",4+32)
#       Return value 6 = Yes, 7 = No     
    }
    If ($ConfDisable -eq 6)
    {
        Build-RelaseForm
        Publish-Form

        set-mailbox $Global:ENo -LitigationHoldEnabled $false
        $Global:AuthDetails = $Global:txtReleaseby.Text + " per " + $Global:txtEMailDate.Text + " email"
        $LineToWrite = $WhoAmI + "`t" + "Release (Details/By/Date)      :" + "`t" + $Global:AuthDetails
        WriteReportEvent

        If ($Global:UsrDetails.ProtectedFromAccidentalDeletion -eq $True)
        {
            $Global:Protection = "False"
            Set-ADObjProtect
        }

        If (($Global:inf.CustomAttribute1 -like "*Ex-*") -and ($Global:strUserPath -like "*_Extend*"))
        {
	        write-host "Moving AD Object from ExtendedHold to Disabled OU"
            Move-ADObjectOU
            $OULoc = "ExtendedHold OU to the Disabled OU"
	    }

        set-aduser $Global:ENo -clear ExtensionAttribute14
        $Global:u.ExtensionAttribute14.value = $null
#        $Global:u.CommitChanges()

        If ($Global:inf.InPlaceHolds.length -ne 0)
        {
            $Hold = $inf.InPlaceHolds
            foreach ($Hold in $Hold)
            {
                $CaseHold = Get-CaseHoldPolicy ($Hold -replace ("UniH",""))
                $Case = Get-ComplianceCase $CaseHold.CaseID
                $Output = $wshell.Popup("An active InPlace Hold was found on this account contact Legal to close Case Name: " + $Case.Name,0,"Active InplaceHold Found",0+32)
            }
        }
        $Global:Republish = "Y"
    }
    else
    {
        $Output = $wshell.Popup("Legal Hold Disablement has been cancelled.",0,"Cancelled",0+32)
        $Global:Republish = "N"
    }
}  

Function Execute-Hold
{
    If ($Global:inf.LitigationHoldEnabled -ne $True)
    {
        #Enable Legal Hold on Account
        set-mailbox $Global:ENo -LitigationHoldEnabled $true
        $LineToWrite = $WhoAmI + "`t" + "Enabling Legal Hold for " + $Global:ENo + " with Retention Comment of: " + $Global:NewCmt + "`n"
        WriteReportEvent
    }

    If (($Global:inf.CustomAttribute1 -like "*Ex-*") -and ($Global:strUserPath -notlike "*_Extend*"))
    {
		$OULoc = "ExtendedHold OU"
        Move-ADObjectOU
	}

    If ($Global:UsrDetails.ProtectedFromAccidentalDeletion -eq $false)
    {
        Set-ADObjProtect
    }
    $Global:Republish = "Y"
}

Function Modify-Comment
{
    $Global:NewCmt = $Global:TxtCmt.Text
    If (($Global:chkHHold.Checked -eq $true) -and ($Global:NewCmt -notlike "*\HHold"))
    {
        $Global:NewCmt = $Global:NewCmt.replace("NoHHold","HHold")
    }
    elseif (($Global:chkHHold.Checked -eq $false) -and ($Global:NewCmt -like "*\HHold"))
    {
        $Global:NewCmt = $Global:NewCmt.replace("HHold","NoHHold")
    }

    If (($Global:chkSilent.Checked -eq $true) -and ($Global:NewCmt -notlike "NonDisclosed*"))
    {
        $Global:NewCmt = "NonDisclosed - " + $Global:NewCmt
    }
    elseif (($Global:chkSilent.Checked -eq $false) -and ($Global:NewCmt -like "NonDisclosed*"))
    {
        $Global:NewCmt = $Global:NewCmt.replace("NonDisclosed - ","")
    }

    If ($Global:NewCmt -ne $Global:StoredCmt)
    {
        $LineToWrite = $WhoAmI + "`t" + "Changing Retention Comment     :`t" + $Global:locListBox.SelectedItem
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "Updated Retention Comment      :`t" + $Global:NewCmt
        WriteReportEvent
        $Global:NewCmt = ($Global:StoredCmt).replace($Global:locListBox.Text,$Global:NewCmt)
        Commit-CommentChange
        $Global:Republish = "Y"
    }
    else 
    {
        $Output = $wshell.Popup("No changes needed.",0,"No Changes",0+32)
    }
}

Function Set-ADObjProtect
{
    $Global:SetObj = ""
    $ErrorActionPreference = "SilentlyContinue"
    switch ($Global:Protection)
    {
        "True"
        {
            (Set-ADObject -Identity $Global:u.DistinguishedName.value -ProtectedFromAccidentalDeletion $True -Credential $AdmCred)
        }
        "False"
        {
            (Set-ADObject -Identity $Global:u.DistinguishedName.value -ProtectedFromAccidentalDeletion $False -Credential $AdmCred)
        }
    }
    $ErrorActionPreference = "Continue"
    Start-Sleep -Seconds 10
    $Protected = Get-ADObject -Identity $Global:u.DistinguishedName.value -Properties ProtectedfromAccidentalDeletion

    If (($Protected.ProtectedfromAccidentalDeletion -eq "True") -and ($Global:Protection -eq "True"))
    {
        $LineToWrite = $WhoAmI + "`t" + "AD Object Protection Enabled"
    }
    else
    {
        If (($Protected.ProtectedfromAccidentalDeletion -ne "True") -and ($Global:Protection -eq "False"))
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

Function Move-ADObjectOU
{
    $Global:MoveObj = "False"
    $ErrorActionPreference = "SilentlyContinue"
    If (($Global:u.ExtensionAttribute1.value -like "*Ex-*") -and (($Global:chkEnable.Checked -eq "Checked") -or ($Global:chkMove.Checked -eq "Checked")))
    {
        $OULoc = "ExtendedHold OU"
        write-host "Moving AD Object to the ExtendedHold OU"
        $LineToWrite = $WhoAmI + "`t" + "Moved to ExtendedHold from     :" + "`t" + $Global:strUserPath
        $Global:MoveObj = [bool](Move-ADObject ($Global:strUserPath -replace("LDAP://","")) 'OU=_ExtendedHold,DC=global,DC=ul,DC=com')
        WriteLogEvent
    }
    elseif (($Global:u.ExtensionAttribute1.value -like "*Ex-*") -and ($Global:chkDisable.Checked -eq "Checked"))
    {
        $OULoc = "Disabled OU"
        write-host "Moving AD Object to the Disabled OU"
        $LineToWrite = $WhoAmI + "`t" + "Moved to ExtendedHold from     :" + "`t" + $Global:strUserPath
        $Global:MoveObj = [bool](Move-ADObject ($Global:strUserPath -replace("LDAP://","")) 'OU=Disabled,DC=global,DC=ul,DC=com')
        WriteLogEvent
    }
    elseif ($Global:u.ExtensionAttribute1.value -notlike "*Ex-*")
    {
        $OULoc = "Active"
    }
    $ErrorActionPreference = "Continue"

    If ($Global:MoveObj -eq $True)
    {
        If ($OULoc -like "*Disabled*")
        {
            write-host "Moving AD Object to the Disabled OU"
            $Global:u.DistinguishedName = $Global:u.DistinguishedName.value.Remove($Global:u.DistinguishedName.value.IndexOf("OU=")) + "OU=Disabled,DC=global,DC=ul,DC=com"
            $LineToWrite = $WhoAmI + "`t" + "Moved to Disabled OU from     :" + "`t" + $Global:strUserPath
            WriteReportEvent
        }
        elseif ($OULoc -like "*Extend*")
        {
            write-host "Moving AD Object to the ExtendedHold OU"
            $Global:u.DistinguishedName = $Global:u.DistinguishedName.value.Remove($Global:u.DistinguishedName.value.IndexOf("OU=")) + "OU=_ExtendedHold,DC=global,DC=ul,DC=com"
            $LineToWrite = $WhoAmI + "`t" + "Moved to ExtendedHold from     :" + "`t" + $Global:strUserPath
            WriteReportEvent
        }
        elseif ($OULoc -like "*Active*")
        {
            $output = $wshell.Popup("This individual needs to be moved back to an Active OU.",0,"Move to Active OU",0+32)
        }
        else
        {
            $Output = $wshell.Popup("Insufficient privilege to move account to the " + $OULoc + ": " + ($Global:strUserPath -replace("LDAP://","")),0,"Insufficient Privilege",0+32)
        }
	}

}

############################################################
#
#Start Of Script

$AdmCred        = Import-Clixml "c:\Users\96151\Documents\myA96151.xml"
$FileName		= "LegalHold"
$LogDrive		= "E:"
$LogPath		= "\Automation"
$LogFolder		= "\" + $FileName
$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"

$Global:chkEnable = ""
$Global:chkModify = ""
$Global:chkRemove = ""
$Global:strUserPath = ""
$Global:u = ""
$Global:inf = ""
$Global:MBXAccess = ""
$strDN = 0
$Global:FormRefresh = "N"
$Global:Protection = "$True"

CheckLogFiles
write-host "Starting Script Legal Hold Check " -ForegroundColor Magenta
write-host
write-Host "`nEnter Employee Number : " -ForegroundColor Red -NoNewline
$Global:ENo = Read-Host
$Global:ENo = $Global:ENo.Trim(" ")
$Global:ReportFile = $LogDirectory + "Report\Report-" + $FileName + "-EmpNo" + $Global:ENo  + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
WriteReportEvent

Do
{
    $Global:ReportFile = $LogDirectory + "Report\Report-" + $FileName + "-EmpNo" + $Global:ENo  + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
    WriteReportEvent
    Get-ENo
    If ($ENo -ne "0")
    {
        $Global:txtReleaseby = ""
        $Global:Republish = "N"
        Build-DefaultForm
        Add-ActionBoxes
        Publish-Form
    }

    # Finished with Dialog Box, Now let's see what the user did... 

    If ($Global:Result -eq "OK")
    {
        #Add New Legal or 40 Day Hold
        If (($Global:chk40Day.Checked -eq "Checked") -or ($Global:chkEnable.Checked -eq "Checked"))
        {
            $Global:txtAttorney = ""
            $Global:txtHRRep = ""
            Enable-NewHoldForm
            Publish-Form
            If ($Global:Result -eq "OK")
            {
                $LineToWrite = $WhoAmI + "`t" + "Adding New 40-Day or Legal Hold Activity"
                WriteReportEvent
                If ($Global:u.ExtensionAttribute14.value.Length -eq 0)
                {
                    $Global:StoredCmt = ""
                }
                Build-RetentComment
            }
            $Global:chk40Day.Checked = ""
            $Global:chkEnable.Checked = ""
        }

        If ($Global:chkDisable.Checked -eq "Checked")
        {
            $LineToWrite = $WhoAmI + "`t" + "Disabling Legal Hold from Account"
            WriteReportEvent
            Disable-Hold
            $Global:chkDisable.Checked = ""
#            $Global:Republish = "Y"
            $Global:chkDisable.Checked = ""
        } 
       
        If ($Global:chkModify.Checked -eq "Checked")
        {
            $LineToWrite = $WhoAmI + "`t" + "Modifying Retention Comment"
            WriteReportEvent
            Modify-CommentForm
            Publish-Form
            Modify-Comment
            $Global:locListBox = ""
            $Global:chkModify.Checked = ""
        }

        If ($Global:chkRemove.Checked -eq "Checked")
        {
            Build-RelaseForm
            Publish-Form
            If ($Global:Result -eq "OK")
            {
                $Global:AuthDetails = "Release By/Details: " + $Global:txtReleaseby.Text + " Per: " + $Global:txtEMailDate.Text + " email"
                $Global:StoredCmt = $Global:u.ExtensionAttribute14.Value
                $Cmt1 = "*, " + $Global:locListBox.SelectedItem + "*"
                $Cmt2 = "*" + $Global:locListBox.SelectedItem + ", *"
                If ($Global:u.ExtensionAttribute14.value -like $Cmt1)
                {
                    $Cmt1 = ", " + $Global:locListBox.SelectedItem
                    $Global:NewCmt = $Global:u.ExtensionAttribute14.replace($Cmt1,"")
                }
                elseif
                ($Global:u.ExtensionAttribute14.value -like $Cmt2)
                {
                    $Cmt2 = $Global:locListBox.SelectedItem + ", "
                    $Global:NewCmt = $Global:u.ExtensionAttribute14.replace($Cmt2,"")
                }
                $LineToWrite = $WhoAmI + "`t" + "Removing Hold Activity         :`t" + $Global:locListBox.SelectedItem
                WriteReportEvent
                $LineToWrite = $WhoAmI + "`t" + "Released Details               :`t" + $Global:AuthDetails
                WriteReportEvent
                Commit-CommentChange
                $Global:Republish = "Y"
                $Global:chkRemove.Checked = ""
            }
            else
            {
                $Output = $wshell.Popup("Hold Comment Removal Cancelled.",0,"Cancelled",0+32)
            }
        }
        
        If ($Global:chkUnhide.Checked -eq "Checked")
        {
            $LineToWrite = $WhoAmI + "`t" + "Unhiding Account from Address Book"
            WriteReportEvent
            write-host "UnHiding User from Outlook Address Book....." -ForegroundColor Magenta
            write-host ""
            $Global:u.msExchHideFromAddressLists.value = $FALSE
            $Global:u.CommitChanges()
            $Global:Republish = "Y"
            $Global:chkUnHide.Checked = ""
        }

        If ($Global:chkHide.Checked -eq "Checked")
        {
            $LineToWrite = $WhoAmI + "`t" + "Hiding Account from Address Book"
            WriteReportEvent
            write-host "Hiding User from Outlook Address Book....." -ForegroundColor Magenta
            write-host ""
            $Global:u.msExchHideFromAddressLists.value = $TRUE
            $Global:u.CommitChanges()
            $Global:Republish = "Y"
            $Global:chkHide.Checked = ""
        }

        If ($Global:chkMove.Checked -eq "Checked")
        {
            $ErrorActionPreference = "SilentlyContinue"
            write-host "Removing Protect from Accidental Deletion in order to move Object to the _ExtendedHold OU"
            $Global:SetObj = [bool](Set-ADObject -Identity $Global:u.DistinguishedName.value -ProtectedFromAccidentalDeletion $false)
            $ErrorActionPreference = "Continue"
            If ($Global:SetObj -eq $True)
            {
                $OULoc = "ExtendedHold OU"
                Move-ADObject
                write-host "Re-enabling Protect from Accidental Deletion"
                $Global:Protection = "$true"
                Set-ADObjProtect
                $Global:Republish = "Y"
                $Global:chkMove.Checked = ""
            }
            else
            {
                $Output = $wshell.Popup("Unable to move account due to insufficient privilege to remove AD Object Protection.  Please set manually move the account.",0,"AD Object Protection",0+32)
            }
            $Global:Republish = "Y"
            $Global:chkMove.Checked = ""
        }

        If ($Global:chkAccess.Checked -eq "Checked")
        {
            $LineToWrite = $WhoAmI + "`t" + "Removing Access From Mailbox"
            WriteReportEvent
            
            # Remove Mailbox Permissions
            If ($Global:u.ExtensionAttribute14.value.length -eq 0)
            {
                $RemoveMbxAccess = Get-MailboxPermission $ENo |Where-Object {$_.User -like "*global.ul.com"}
                If ($RemoveMbxAccess.Count -ne 0)
                {
                    Foreach ($RemoveMbxAccess in $RemoveMbxAccess)
                    {
                        If ($Global:u.ExtensionAttribute1.value -notlike "Ex-*")
                        {
                            Write-Host "Would you like to remove mailbox access from" (get-mailbox $RemoveMbxAccess.User).PrimarySMTPAddress" (Y/N)? "-ForegroundColor Cyan -NoNewline
                            $RmUsr = Read-Host
                        }
                        else
                        {
                            $RmUsr = "Y"
                        }

                        If ($RmUsr -eq "Y")
                        {
                            Remove-MailboxPermission $ENo -User $RemoveMbxAccess.User -AccessRights $RemoveMbxAccess.AccessRights -Confirm:$False
                            $LineToWrite = $WhoAmI + "`t" + "Removing Permissions from      :" + "`t" + $RemoveMbxAccess.User + " - " + $RemoveMbxAccess.AccessRights
                            WriteReportEvent
                        }
                    }
                    $Global:chkAccess.Checked = ""
                }
            }
            else
            {
                $Output = $wshell.Popup("There are other active legal holds on this mailbox access is not being removed.",0,"Active Holds Exist",0+32)
            }

            # Remove Folder Permissons
            $RemoveMbxFldrAccess = Get-MailboxFolderPermission $ENo |Where-Object {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous")}
            If ($RemoveMbxFldrAccess.Count -ne 0)
            {
                Foreach ($RemoveMbxFldrAccess in $RemoveMbxFldrAccess)
                {
                    If ($Global:u.ExtensionAttribute1.value -notlike "Ex-*")
                    {
                        Write-Host "Would you like to remove mailbox folder access from" $RemoveMbxFldrAccess.User.DisplayName "(Y/N)? "-ForegroundColor Cyan -NoNewline
                        $RmUsr = Read-Host
                    }
                    else
                    {
                        $RmUsr = "Y"
                    }

                    If ($RmUsr -eq "Y")
                    {
                        $First = "Y"
                        $DelegateTo = (get-mailbox $RemoveMbxFldrAccess.User.DisplayName).PrimarySMTPAddress
                        $MBXFldrCnt = (Get-MailboxFolderStatistics $ENo).count
                        $MBXFldr = Get-MailboxFolderStatistics $ENo | % {$_.Identity.ToString().Split("\")[1..$MBXFldrCnt] -join "\"}
	
                        ForEach ($Folder in $MBXFldr)
	                    {
                            If ($First -eq "Y")
                            {
                                $First = "N"
                                write-host "Removing " $DelegateTo " - " $RemoveMbxFldrAccess.AccessRights "access rights" -ForegroundColor Green
                                $LineToWrite = $WhoAmI + "`t" + $DelegateTo + " " +  $RemoveMbxFldrAccess.AccessRights + " access rights"
                                WriteReportEvent
                                Remove-MailboxFolderPermission -Identity $ENo -User $DelegateTo -confirm:$False -ErrorAction SilentlyContinue|out-null
                                write-host " <-- from the Top of the Inforamtion Store" -ForegroundColor Green
                                $LineToWrite = $WhoAmI + "`t" + "  <-- from the Top of the Inforamtion Store"
                	            WriteReportEvent
                            }
		                    else
                            {
                                Remove-MailboxFolderPermission -Identity ($ENo + ":\" + $Folder) -User $DelegateTo -AccessRights $RemoveMbxFldrAccess.AccessRights -confirm:$False -ErrorAction SilentlyContinue |out-null
                                write-host " <-- from the folder" $Folder -ForegroundColor Green
                                $LineToWrite = $WhoAmI + "`t" + "  <-- from the folder " + $Folder
                                WriteReportEvent
                            }
                        }
                    }
                }
            }
            $Global:chkAccess.Checked = ""
        }

        If ($Global:chkReport.Checked -eq "Checked")
        {
            $Global:ReportFile = $LogDirectory + "Report\Report-" + $FileName + "LegalReport-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            Do
            {
                GetAcctInfo($Global:ENo)
                ShowDetails
                Write-Host ""
                write-Host "Enter Employee Number or 0 to Exit : " -ForegroundColor Red -NoNewline
                $Global:ENo = Read-Host
          } while ($Global:ENo -ne 0)
          $Global:chkReport.Checked = ""
        }

        If ($Global:Republish -eq "Y")
        {
            Republish-Form
        }
    }

    $Global:StoredCmt = ""
    $Global:locListBox = ""
    $Global:NewCmt = ""
    $Cmt1 = ""
    $Cmt2 = ""

    If ($Global:ENo -ne 0)
    {
        write-Host "`nEnter Employee Number : " -ForegroundColor Red -NoNewline
        $Global:ENo = Read-Host
    }
}while ($Global:ENo -ne "0")
