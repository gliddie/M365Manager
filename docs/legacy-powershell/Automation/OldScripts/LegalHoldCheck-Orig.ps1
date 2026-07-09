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
 #  05/28/2021 - SAG:Modified to use button to get employee details and the processes to enable/disable holds need to review others
 #  06/15/2021 - SAG:Modified the chk40day and that new holds and new hr reps are stored in the selection files
 #  08/10/2021 - SAG:Modified to display DBOwner details for individuals, to be able to add dbOwner details for new holds and documenting the removal of DBOwner deatils
 #  12/01/2021 - SAG:Added Focus statement to the Release function
 #  05/03/2022 - SAG:Modified script to eliminate redisplaying form when no actions are selected, resized form and recoded when details are entered where no mailbox exists
 #  09/30/2022 - SAG:Added code to add/remove the individual from the iProSearch group
 #  02/22/2023 - SAG:Modified code to allow for the script to remove hold comments if a person is a DBOwner
 #  04/27/2023 - SAG:Removed the Employee Type when writing to the DBOwner as this only captures this information at the time the record is written
#>  

#Forms Functions
Function Add-ActionBoxes
{
    ## New Legal Hold RadioButton 
    $Global:chkEnable = New-Object Windows.Forms.RadioButton
        $Global:chkEnable.Left = 570; $Global:chkEnable.Width = 200; $Global:chkEnable.Top = 60  
        $Global:chkEnable.Text = "New Legal Hold" 
        $Global:chkEnable.Checked = $False   # set a default value 
        $Global:chkEnable.TabIndex = 0
        $Global:form.Controls.Add($Global:chkEnable) 

    ## New 40 Day Hold RadioButton      
    $Global:chk40Day = New-Object Windows.Forms.RadioButton
        $Global:chk40Day.Left = 570; $Global:chk40Day.Width = 200; $Global:chk40Day.Top = 90  
        $Global:chk40Day.Text = "New 40 / 180 Day Hold" 
        $Global:chk40Day.Checked = $False   # set a default value 
        $Global:chk40Day.TabIndex = 1
        $Global:form.Controls.Add($Global:chk40Day) 

    $TopAlignment = 120

    If ($Global:inf.LitigationHoldDate -ne $null)
    {
        ## Disable Comment Hold RadioButton
        $Global:chkDisable = New-Object Windows.Forms.RadioButton
        $Global:chkDisable.Left = 570; $Global:chkDisable.Width = 200; $Global:chkDisable.Top = $TopAlignment 
        $Global:chkDisable.Text = "Disable Legal Hold" 
        $Global:chkDisable.Checked = $False
        $Global:chkDisable.TabIndex = 2
        $Global:form.Controls.Add($Global:chkDisable)
        $TopAlignment = $TopAlignment + 30
    }

#    If ($Global:locListBox.Items -ne "")
    If ($Global:u.ExtensionAttribute14.value.length -ne 0)
    {
        ## Modify Hold RadioButton
        $Global:chkModify = New-Object Windows.Forms.RadioButton
        $Global:chkModify.Left = 570; $Global:chkModify.Width = 200; $Global:chkModify.Top = $TopAlignment 
        $Global:chkModify.Text = "Modify Comment" 
        $Global:chkModify.Checked = $False
        $Global:chkModify.TabIndex = 3
        $Global:form.Controls.Add($Global:chkModify)
        $TopAlignment = $TopAlignment + 30
    }

    If ($Global:locListBox.Items -ne "")
#    If (($Global:u.ExtensionAttribute14.value.length -ne 0) -and ($Global:u.ExtensionAttribute14.value -like "*,*"))
    {
        ## Remove Comment Hold RadioButton
        $Global:chkRemove = New-Object Windows.Forms.RadioButton
        $Global:chkRemove.Left = 570; $Global:chkRemove.Width = 200; $Global:chkRemove.Top = $TopAlignment  
        $Global:chkRemove.Text = "Remove Comment" 
        $Global:chkRemove.Checked = $False
        $Global:chkRemove.TabIndex = 4
        $Global:form.Controls.Add($Global:chkRemove)
        $TopAlignment = $TopAlignment + 30
    }

    If ($Global:MbxAccessForm -ne "(None)")
    {
        ## Remove Access to Mailbox RadioButton
        $Global:chkAccess = New-Object Windows.Forms.RadioButton
        $Global:chkAccess.Left = 570; $Global:chkAccess.Width = 200; $Global:chkAccess.Top = $TopAlignment  
        $Global:chkAccess.Text = "Remove Access" 
        $Global:chkAccess.Checked = $False
        $Global:chkAccess.TabIndex = 5
        $Global:form.Controls.Add($Global:chkAccess)
        $TopAlignment = $TopAlignment + 30
    }

    If (($Global:u.ExtensionAttribute1 -like "Ex-*"))
    {
        If ($Global:inf.HiddenFromAddressListsEnabled -eq $True)
        {
            ## Unhide from AddressBook Hold RadioButton
            $Global:chkUnhide = New-Object Windows.Forms.RadioButton
            $Global:chkUnhide.Left = 570; $Global:chkUnhide.Width = 200; $Global:chkUnhide.Top = $TopAlignmenT
            $Global:chkUnhide.Text = "Unhide Account" 
            $Global:chkUnhide.Checked = $False
            $Global:chkUnhide.TabIndex = 6
            $Global:form.Controls.Add($Global:chkUnhide)
            $TopAlignment = $TopAlignment + 30
        }
        else
        {
            ## Hide from AddressBook Hold RadioButton
            $Global:chkHide = New-Object Windows.Forms.RadioButton
            $Global:chkHide.Left = 570; $Global:chkHide.Width = 200; $Global:chkHide.Top = $TopAlignmenT
            $Global:chkHide.Text = "Hide Account" 
            $Global:chkHide.Checked = $False
            $Global:chkHide.TabIndex = 7
            $Global:form.Controls.Add($Global:chkHide)
            $TopAlignment = $TopAlignment + 30
        }

        If ($Global:strUserPath -notlike "*_Extended*")
        {
            ## Move to ExtendedHold RadioButton
            $Global:chkMove = New-Object Windows.Forms.RadioButton
            $Global:chkMove.Left = 570; $Global:chkMove.Width = 200; $Global:chkMove.Top = $TopAlignment
            $Global:chkMove.Text = "Move to ExtendedHold" 
            $Global:chkMove.Checked = $False
            $Global:chkMove.TabIndex = 8
            $Global:form.Controls.Add($Global:chkMove)
            $TopAlignment = $TopAlignment + 30
        }
    }

    ## Database Owner RadioButton
    $Global:chkDBOwner = New-Object Windows.Forms.RadioButton
        $Global:chkDBOwner.Left = 570; $Global:chkDBOwner.Width = 200; $Global:chkDBOwner.Top = $TopAlignment
        $Global:chkDBOwner.Text = "Database Owner" 
        $Global:chkDBOwner.Checked = $False
        $Global:form.Controls.Add($Global:chkDBOwner)
        $TopAlignment = $TopAlignment + 30

    ## Create Report
    $Global:chkReport = New-Object Windows.Forms.RadioButton
        $Global:chkReport.Left = 570; $Global:chkReport.Width = 200; $Global:chkReport.Top = $TopAlignment
        $Global:chkReport.Text = "Generate Report" 
        $Global:chkReport.Checked = $Global:chkReport.Checked   # set a default value 
        $Global:chkReport.TabIndex = 9
        $Global:form.Controls.Add($Global:chkReport) 
}

Function Hide-ActionBoxes
{
    $Global:chkEnable.visible = $false  
    $Global:chk40Day.Visible = $false
    If ($Global:inf.LitigationHoldDate -ne $null)
    {
        $Global:chkDisable.visible = $false
    }

#    If ($Global:locListBox.Items -eq "")
    If ($Global:u.ExtensionAttribute14.value.length -ne 0)
    {
        $Global:chkModify.visible = $false
    }

#    If (($Global:u.ExtensionAttribute14.value.length -ne 0) -and ($Global:u.ExtensionAttribute14.value -like "*,*"))
    If ($Global:u.ExtensionAttribute14.value.length -ne 0)
    {
        $Global:chkRemove.visible = $false
    }

    If ($Global:MbxAccessForm -ne "(None)")
    {
        $Global:chkAccess.visible = $false
    }

    If (($Global:u.ExtensionAttribute1 -like "Ex-*"))
    {
        If ($Global:inf.HiddenFromAddressListsEnabled -eq $True)
        {
            $Global:chkUnhide.visible = $false
        }
        else
        {
            $Global:chkHide.visible = $false
        }

        If ($Global:strUserPath -notlike "*_Extended*")
        {
            $Global:chkMove.visible = $false
        }
    }
    $Global:chkDBOwner.visible = $false
    $Global:chkReport.visible = $false
}

Function Add-SilentHHHoldBoxes
{
    $form.Height = 640
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
        $Global:TabIndex++
}

Function Build-DefaultForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Legal Hold Details" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 740 ; $form.Height = 450   # Make the form wider 
  
    ## Employee ID
    $Global:lblHost = New-Object System.Windows.Forms.Label   
        $Global:lblHost.Text = "Employee No:" 
        $Global:lblHost.Top = 10 ; $Global:lblHost.Left = 5; $Global:lblHost.Width=150 ;$Global:lblHost.AutoSize = $true 
        $form.Controls.Add($Global:lblHost)    # Add to Form 
        # 
        $Global:txtHost = New-Object Windows.Forms.TextBox
        $Global:txtHost.Top = 10; $Global:txtHost.Left = 160; $Global:txtHost.Width = 120;  
        $Global:txtHost.Text = ""
        $Global:form.Controls.Add($Global:txtHost)    # Add to Form
		$Global:txtHost.Add_Click({
			$Global:txtHost.Text = ""
			$Global:txtHost.Top = 10; $Global:txtHost.Left = 160; $Global:txtHost.Width = 120; 
			$Global:ButGetENo.Location = New-object System.Drawing.Size(290,10)
		})
        $Global:InputFocus = $Global:txtHost

        If ($Global:FormRefresh -ne "Y")
        {
            $Global:ButGetENo = New-Object Windows.Forms.Button
            $Global:ButGetENo.Location = New-object System.Drawing.Size(290,10)
            $Global:ButGetENo.Size = new-Object System.Drawing.Size(150,20)
            $Global:ButGetENo.Text = "Get Employee Details"
            $Global:ButGetENo.TabIndex = 1
            $Global:form.Controls.Add($Global:ButGetENo)
            $Global:ButGetENo.Add_Click({                
                $ButGetENo.visible = $false
                $Global:ENo = $Global:txtHost.Text.Trim(" ")
                write-host "Retreiving Account Details for EmpNo: $Global:ENO" -ForegroundColor Cyan
                $Global:txtReleaseby = ""
                $Global:Republish = "Y"
                $Exists = [bool](get-mailbox $global:txthost.Text -ErrorAction SilentlyContinue)
                If ($Exists -eq $True)
                {
                    $Global:ReportFile = $LogDirectory + "Report\Report-" + $FileName + "-EmpNo" + $Global:ENo  + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
                    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
                    WriteReportEvent
                    $LineToWrite = $WhoAmI + "`t" + "Employee Number                :" + "`t" + $Global:txtHost.Text
                    WriteReportEvent
                    Check-Reconnect
                    Get-ENo
                    Refresh-DefaultForm
                    Add-ActionBoxes
                    $Global:OKButton.visible = $True 
                    $Global:okButton.Text = "Continue"
                    $Global:cancelButton.Text = "Next Emp"
                }
                else
                {
					$UPN = $Global:ENo + "@global.ul.com"
					$UseLoc = (Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue).UsageLocation
					If ($UseLoc.Length -gt 0)
					{
						$Global:txtHost.Text = "No Mailbox Found (" + $Global:ENo + ")"
						$Global:txtHost.Width = 150;
						$Global:ButGetENo.Location = New-object System.Drawing.Size(320,10)							
					}
					else
					{
						$Global:txtHost.Text = "No AD Account or Mailbox Never Configured (" + $Global:ENo + ")"
						$Global:txtHost.Width = 280;
						$Global:ButGetENo.Location = New-object System.Drawing.Size(450,10)
					}
					$Global:ButGetENo.Visible = "True"
                }
            })
        }

    ## Employee Name
    $Global:lblName = New-Object System.Windows.Forms.Label   
        $Global:lblName.Text = "Employee Name:"  
        $Global:lblName.Top = 40 ; $Global:lblName.Left = 5; $Global:lblName.Width=150 ;$Global:lblName.AutoSize = $true 
        $form.Controls.Add($Global:lblName)    # Add to Form 
        # 
        $Global:txtName = New-Object Windows.Forms.TextBox
        $Global:txtName.ReadOnly = $True;
        $Global:txtName.Top = 40; $Global:txtName.Left = 160; $Global:txtName.Width = 120;
        $Global:txtName.Text = ""
        $Global:form.Controls.Add($Global:txtName)    # Add to Form 

    ## Employee Type
    $Global:lblType = New-Object System.Windows.Forms.Label   
        $Global:lblType.Text = "Employee Type:"  
        $Global:lblType.Top = 70 ; $Global:lblType.Left = 5; $Global:lblType.Width=150 ;$Global:lblType.AutoSize = $true 
        $form.Controls.Add($Global:lblType)    # Add to Form 
        # 
        $Global:txtType = New-Object Windows.Forms.TextBox
        $Global:txtType.ReadOnly = $True;
        $Global:txtType.Top = 70; $Global:txtType.Left = 160; $Global:txtType.Width = 120;
        $Global:txtType.Text = ""
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
        $Global:txtContainer.Text = ""
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
        $Global:txtDate.Text = ""
        $Global:form.Controls.Add($Global:txtDate)    # Add to Form 
 
    ## Hold Enabled By
    $Global:lblBy = New-Object System.Windows.Forms.Label   
        $Global:lblBy.Text = "Hold Enabled By:"  
        $Global:lblBy.Top = 130; $Global:lblBy.Left = 310; $Global:lblBy.Width=100 ;$Global:lblBy.AutoSize = $true  
        $form.Controls.Add($Global:lblBy)    # Add to Form 
        # 
        $Global:txtBy = New-Object Windows.Forms.TextBox  
        $Global:txtBy.ReadOnly = $True; 
        $Global:txtBy.Top = 130; $Global:txtBy.Left = 410; $Global:txtBy.Width = 120;  
        $Global:txtBy.Text = ""
        $Global:form.Controls.Add($Global:txtBy)    # Add to Form 

    ## Hold ADOjbect Protected
    $Global:lblProtect = New-Object System.Windows.Forms.Label   
        $Global:lblProtect.Text = "AD Object Protected:"  
        $Global:lblProtect.Top = 160 ; $Global:lblProtect.Left = 5; $Global:lblProtect.Width=150 ;$Global:lblProtect.AutoSize = $true 
        $form.Controls.Add($Global:lblProtect)    # Add to Form 
        # 
        $Global:txtProtect = New-Object Windows.Forms.TextBox  
        $Global:txtProtect.ReadOnly = $True; 
        $Global:txtProtect.Top = 160; $Global:txtProtect.Left = 160; $Global:txtProtect.Width = 120;
        $Global:txtProtect.Text = ""
        $Global:form.Controls.Add($Global:txtProtect)    # Add to Form 

    ## Hold ADOjbect Hidden
    $Global:lblHide = New-Object System.Windows.Forms.Label   
        $Global:lblHide.Text = "AD Object Hidden:"  
        $Global:lblHide.Top = 190 ; $Global:lblHide.Left = 5; $Global:lblHide.Width=160 ;$Global:lblHide.AutoSize = $true 
        $form.Controls.Add($Global:lblHide)    # Add to Form 
        # 
        $Global:txtHide = New-Object Windows.Forms.TextBox  
        $Global:txtHide.ReadOnly = $True;
        $Global:txtHide.Top = 190; $Global:txtHide.Left = 160; $Global:txtHide.Width = 120;
        $Global:txtHide.Text = ""
        $Global:form.Controls.Add($Global:txtHide)    # Add to Form 

    ## Accounts Granted Access
    $Global:lblAccess = New-Object System.Windows.Forms.Label   
        $Global:lblAccess.Text = "Accounts with Access:"  
        $Global:lblAccess.Top = 220 ; $Global:lblAccess.Left = 5; $Global:lblAccess.Width=160 ;$Global:lblAccess.AutoSize = $true 
        $form.Controls.Add($Global:lblAccess)    # Add to Form 
        # 
        $Global:txtAccess = New-Object Windows.Forms.TextBox  
        $Global:txtAccess.ReadOnly = $True; 
        $Global:txtAccess.Top = 220; $Global:txtAccess.Left = 160; $Global:txtAccess.Width = 370; 
        $Global:txtAccess.Text = ""
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
        $Global:txtInPlace.Text = ""
        $Global:form.Controls.Add($Global:txtInPlace)    # Add to Form           
    
    ## ListBox - Fill with Data From Azure Location Name 
    $Global:lblLoc = New-Object System.Windows.Forms.Label   
        $Global:lblLoc.Text = "Preservation Comment(s):"; $Global:lblLoc.Top = 280; $Global:lblLoc.Left = 5; $Global:lblLoc.Autosize = $true  
        $Global:form.Controls.Add($Global:lblLoc)  
        # Listbox for Location Name 
        $Global:locListBox = New-Object System.Windows.Forms.ListBox  
        $Global:locListBox.Top = 280; $locListBox.Left = 160; $locListBox.Height = 100; $LocListBox.Width = 370;
        $Global:locListBox.TabIndex = 1
        $Global:form.Controls.Add($Global:locListBox) #Add listbox to form 
        # Obtain Value with: $Global:locListBox.SelectedItem
}

Function Build-RelaseForm
{
    $form.Height = 640
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
        $Global:txtReleaseby.Text = "A.Uteg"   # Use Corrent computer name as default 
        $Global:form.Controls.Add($Global:txtReleaseby)    # Add to Form
        $Global:InputFocus = $Global:txtReleasedby

    #Add Release Date
    $Global:lblEMailDate = New-Object System.Windows.Forms.Label
        $Global:lblEMailDate.Text = "EMail Date Releasing:"
        $Global:lblEMailDate.Top = 410; $Global:lblEMailDate.Left = 5; $Global:lblEMailDate.Width=150 ;$Global:lblEMailDate.AutoSize = $true
        $form.Controls.Add($Global:lblEMailDate)    # Add to Form
        #
        $Global:txtEMailDate = New-Object Windows.Forms.DateTimePicker
        $Global:txtEMailDate.TabIndex = 21 # set Tab Order
        $Global:txtEMailDate.Top = 410; $Global:txtEMailDate.Left = 160; $Global:txtEMailDate.Width = 120;
        $Global:txtEMailDate.Format = [windows.forms.datetimepickerFormat]::custom
        $Global:txtEMailDate.CustomFormat = "MM/dd/yyyy"
        $Global:txtEMailDate.Text = (get-date)
        $Global:form.Controls.Add($Global:txtEMailDate)    # Add to Form 

    $Global:InputFocus = $Global:okButton
}

Function Enable-NewHoldForm
{
    $form.Height = 640
    #Add Preservation Name
    $Global:lblPrevName = New-Object System.Windows.Forms.Label
        $Global:lblPrevName.Text = "Preservation Name:"
        $Global:lblPrevName.Top = 380 ; $Global:lblPrevName.Left = 5; $Global:lblPrevName.Width=150 ;$Global:lblPrevName.AutoSize = $true
        $form.Controls.Add($Global:lblPrevName)    # Add to Form 
    $Global:txtPrevName = New-Object Windows.Forms.ComboBox
        $Global:txtPrevName.TabIndex = 19 # set Tab Order 
        $Global:txtPrevName.Top = 380; $Global:txtPrevName.Left = 160; $Global:txtPrevName.Width = 250;
        $Global:StoredHolds = import-csv "e:\O365AdminShared\Data\LegalHolds.csv" | Sort-Object HoldName
        $Global:AddPrevName = "Yes"
        Foreach ($StoredHolds in $Global:StoredHolds)
        {
            If (($Global:chk40Day.Checked -eq "Checked") -and (($StoredHolds.HoldName -like "40*") -or ($StoredHolds.HoldName -like "180*")))
            {
                [void] $Global:txtPrevName.Items.Add($StoredHolds.HoldName)
            }
            else
            {
                If ((($Global:chkEnable.Checked -eq "Checked") -or ($Global:chkDBOwner.Checked -eq "Checked")) -and ($StoredHolds.HoldName -notlike "40*") -and ($StoredHolds.HoldName -notlike "180*"))
                {
                    [void] $Global:txtPrevName.Items.Add($StoredHolds.HoldName)
                }
            }
        }
        $Global:form.Controls.Add($Global:txtPrevName)    # Add to Form 
        $Global:InputFocus = $Global:txtPrevName

       $Top = 410

        If (($Global:chk40Day.Checked -eq "Checked") -or ($Global:chkDBOwner.Checked -eq "Checked"))
        {
            $Global:txtPrevName.Text = "40 Day Hold - "
            #Add HRRep
            $Global:lblHRRep = New-Object System.Windows.Forms.Label
            $Global:lblHRRep.Text = "HR Representative:"
            If ($Global:chkDBOwner.Checked -eq "Checked")
            {
               $Global:lblHRRep.Text = "DB Type:"
               $Global:txtPrevName.Text = "" 
            }
            $Global:lblHRRep.Top = $Top ; $Global:lblHRRep.Left = 5; $Global:lblHRRep.Width=150 ;$Global:lblHRRep.AutoSize = $true
            $form.Controls.Add($Global:lblHRRep)    # Add to Form 
            # 
            $Global:txtHRRep = New-Object Windows.Forms.ComboBox  
            $Global:txtHRRep.TabIndex = 20 # set Tab Order 
            $Global:txtHRRep.Top = $Top; $Global:txtHRRep.Left = 160; $Global:txtHRRep.Width = 120;
            $Global:HRReps = import-csv "e:\O365AdminShared\Data\HRReps.csv" | Sort-Object Name
            If ($Global:chkDBOwner.Checked -eq "Checked")
            {
                $Global:HRReps = import-csv "e:\O365AdminShared\Data\DBOwnerTypes.csv" | Sort-Object Name
            }
            Foreach ($HRReps in $Global:HRReps)
            {
                [void] $Global:txtHRRep.Items.Add($HRReps.Name)
            }
            $Global:form.Controls.Add($Global:txtHRRep)    # Add to Form 
            $Top = $Top + 30
        }

    If ($Global:chkDBOwner.Checked -eq $False)
    {
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
            $Global:txtAttorney.Text = "B.Bogaerts"
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

Function Refresh-DefaultForm
{
    $form.Height = 450
    $Global:txtHost.ReadOnly = $true
    
    $Global:txtName.Text = $Global:inf.DisplayName
    $Global:txtName.Refresh()

    $Global:txtType.Text = $Global:u.ExtensionAttribute1.value
    $Global:txtType.Refresh()

    $Container = $Global:strUserPath -Replace("LDAP://CN=","")
    $Container = $Container -Replace(",OU=","/")
    $Container = $Container -Replace(",DC=","/")
    $Global:txtContainer.Text = ($Container)
    $Global:lblContainer.Refresh()

    $Global:txtDate.Text = $Global:inf.LitigationHoldDate
    $Global:txtDate.Refresh()

    $Global:txtBy.Text = $Global:inf.LitigationHoldOwner
    $Global:txtBy.Refresh()

    $Global:txtProtect.Text = $Global:UsrDetails.ProtectedFromAccidentalDeletion
    $Global:txtProtect.Refresh()

    $Global:txtHide.Text = $Global:inf.HiddenFromAddressListsEnabled
    $Global:txtHide.Refresh()

    $Global:txtAccess.Text = $Global:MbxAccessForm
    $Global:txtAccess.Refresh()

    $Global:txtInPlace.Text = $Global:inf.InPlaceHolds
    $Global:txtInPlace.Refresh()
            
    If ($Global:ADExists -eq $True)
    {
        $Global:StoredCmt = $Global:u.ExtensionAttribute14.value
    }
    else
    {
        $mbx = Get-Mailbox $ENo
        If ($mbx.RetentionComment.length -ne 0)
        {
            $Global:StoredCmt = $mbx.RetentionComment
        }
        else
        {
            $Global:StoredCmt = $mbx.CustomAttribute14
        }
    }
    If ($Global:StoredCmt.Length -ne 0)
    {
        $LocArray = $Global:StoredCmt.split(",")
        foreach ($element in $LocArray)
        {
            [void] $Global:locListBox.Items.Add($element.TrimStart())  # Add element to listbox 
        }
        $Global:locListBox.Refresh() 
    }

    $Global:StoredHolds = import-csv "E:\O365AdminShared\Data\LegalDBOwners.csv" |Sort-Object EmpNo
    $Global:DBOwnerHolds = ""
    Foreach ($StoredHolds in $Global:StoredHolds)
    {
        If ($StoredHolds.EmpNo -eq $Global:ENo)
        {
            $Val = "Listed as DBOwner for: " + $StoredHolds.CaseName + "/" + $StoredHolds.AppType
            $Global:DBOwnerHolds = $Global:DBOwnerHolds + $StoredHolds.CaseName + " , "
            [void] $Global:locListBox.Items.Add($val)
        }
    }
#write-host $Global:locListBox.Items
#pause
    If ($Global:DBOwnerHolds.Length -gt 0)
    {
        $LineToWrite = $WhoAmI + "`t" + "LegalHold DB Owner             :" + "`t" + ($Global:DBOwnerHolds.TrimEnd(" , "))
        WriteReportEvent 
    }
}

Function Remove-DBOwnerDetails
{
    $DBOFile = "E:\O365AdminShared\Data\LegalDBOwners.csv"
    $DBODetails = get-content "E:\O365AdminShared\Data\LegalDBOwners.csv"
    Rename-Item -Path $DBOFile -NewName ("LegalDBOwners-Date" + (((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", ""))) + ".old")
#    $DBOFileNew = "E:\O365AdminShared\Data\LegalDBOwners-new.csv"
    $HoldName = $Global:locListBox.SelectedItem -Replace("Listed as DBOwner for: ","")
    $HoldName = $HoldName -Replace("Listed as DBOwner for: ","")
    $HoldName = $HoldName.Substring(0,$HoldName.IndexOf("/"))

    foreach ($o in $DBODetails)
    {
        If ($o -notmatch $Global:ENo)
        {
            add-content $DBOFile $o
        }
        else
        {
            If ($o -match $HoldName)
            {
#                Do Nothing released as DBOwner from this hold
            }
            else
            {
                add-content $DBOFile $o
            }
        }
    }
}

Function Republish-Form
{
#show updates:
    start-sleep -s 15
    GetAcctInfo($Global:ENo)

    If ($Global:ADExists -eq $True)
    {
        $Global:StoredCmt = $Global:u.ExtensionAttribute14.value
    }
    else
    {
        $mbx = Get-Mailbox $ENo
        If ($mbx.RetentionComment.length -ne 0)
        {
            $Global:StoredCmt = $mbx.RetentionComment
        }
        else
        {
            $Global:StoredCmt = $mbx.CustomAttribute14
        }
    }

    $Global:okButton.Text = "Next Emp"
    $Global:cancelButton.Text = "Exit"
    $Global:locListBox.Items.Clear()
    Refresh-DefaultForm
    Publish-Form
    $Global:FormRefresh = "N"
    $ButGetENo.visible = $true
    $Global:txtHost.ReadOnly = $false
}

Function Update-InputFiles
{
    If ($Global:chkDBOwner.Checked -eq $False)
    {
        #Check if this is a new preservation activity
        $Global:AddPrevName = "Yes"
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
            $LineToWrite = '"{0}"' -f ($Global:txtPrevName.Text)
            Out-File -FilePath "E:\O365AdminShared\Data\LegalHolds.csv" -InputObject $LineToWrite -Append
        }

        #check if this is a new Attorney
        $AddAttName = "Yes"
        $Global:txtAttorney.Text= ($Global:txtAttorney.Text).TrimEnd()
        $Global:txtAttorney.Text = $Global:txtAttorney.Text -replace (",",".")
        If (($Global:txtAttorney.Text -notlike "") -and ($Global:txtAttorney.Text -notlike "*TASK*"))
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
    }
      
    #check if this is a new HRRep or a new DBOwner
    $AddHRRep = "Yes"
    If (($Global:txtHRRep.Text.length -gt 0) -and ($Global:txtHRRep.Text -notlike "*TASK*"))
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
        If ($Global:chkdbOwner.Checked -eq $false)
        {
            Out-File -FilePath "E:\O365AdminShared\Data\HRReps.csv" -InputObject $LineToWrite -Append
        }
        else
        {
            Out-File -FilePath "E:\O365AdminShared\Data\DBOwnerTypes.csv" -InputObject $LineToWrite -Append
        }
    }
}

# Change Functions
Function Build-RetentComment
{
    Update-InputFiles
    
    $Global:txtPrevName.Text = ($Global:txtPrevName.Text)
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

        If ($Global:Result -eq "OK")
        {
            set-mailbox $Global:ENo -LitigationHoldEnabled $false
            $Global:AuthDetails = $Global:txtReleaseby.Text + " per " + $Global:txtEMailDate.Text + " email"
            $LineToWrite = $WhoAmI + "`t" + "Release (Details/By/Date)      :" + "`t" + $Global:AuthDetails + "`n"
            WriteReportEvent
            Remove-DistributionGroupMember iprosearch -Member $Global:ENo -BypassSecurityGroupManagerCheck -confirm:$False
            $LineToWrite = $WhoAmI + "`t" + "Removed from iProSearch Group"
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

            If ($Container -like "*Active Directory Account*")
            {
                set-mailbox $Global:ENo -RetentionComment $null -CustomAttribute14 $null
            }
            else
            {
                set-aduser $Global:ENo -clear ExtensionAttribute14
                $Global:u.ExtensionAttribute14.value = $null
            }

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
            $Global:lblReleaseby.Visible = $false
            $Global:txtReleaseby.visible = $false
            $Global:lblEMailDate.visible = $false
            $Global:txtEMailDate.visible = $false
        }
        else
        {
            $Output = $wshell.Popup("Legal Hold Disablement has been cancelled.",0,"Cancelled",0+32)
            $Global:Republish = "N"
        }
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
        $success = [bool](Add-DistributionGroupMember iprosearch@ul.onmicrosoft.com -Member $Global:ENo -BypassSecurityGroupManagerCheck -ErrorAction SilentlyContinue)
        If ($Success -eq $True)
        {
            $LineToWrite = $WhoAmI + "`t" + "Added user to the iProSearch Group for " + $Global:ENo
        }
        else
        {
            $LineToWrite = $WhoAmI + "`t" + "User is Already a member of the iProSearch Group"
        }
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
    $Global:lblPrevName.visible = $false
    $Global:txtPrevName.visible = $false
    If ($Global:chk40Day.Checked -eq "Checked")
    {
        $Global:lblHRRep.visible = $false
        $Global:txtHRRep.visible = $false
    }
    $Global:lblAttorney.visible = $false
    $Global:txtAttorney.visible = $false
    $Global:chkSilent.visible = $false
    $Global:chkHHold.visible = $false
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

    $Global:lblCmt.Visible = $false
    $Global:txtCmt.visible = $false
    $Global:lblCmt.Visible = $false
    $Global:txtCmt.visible = $false
    $Global:chkSilent.visible = $false
    $Global:chkHHold.Visible = $false
}

Function Set-ADObjProtect
{
    $Global:SetObj = ""
    $ErrorActionPreference = "SilentlyContinue"
    switch ($Global:Protection)
    {
        "True"
        {
            Set-ADObject -Identity $Global:u.DistinguishedName.value -ProtectedFromAccidentalDeletion $True
#            This is not working
#            Set-ADObject -Identity $Global:u.DistinguishedName.value -ProtectedFromAccidentalDeletion $True -Credential $Global:AdmLiveCred
        }
        "False"
        {
            Set-ADObject -Identity $Global:u.DistinguishedName.value -ProtectedFromAccidentalDeletion $False
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
        WriteReportEvent
    }
    elseif (($Global:u.ExtensionAttribute1.value -like "*Ex-*") -and ($Global:chkDisable.Checked -eq "Checked"))
    {
        $OULoc = "Disabled OU"
        write-host "Moving AD Object to the Disabled OU"
        $LineToWrite = $WhoAmI + "`t" + "Moved to ExtendedHold from     :" + "`t" + $Global:strUserPath
        $Global:MoveObj = [bool](Move-ADObject ($Global:strUserPath -replace("LDAP://","")) 'OU=Disabled,DC=global,DC=ul,DC=com')
        WriteReportEvent
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

Function Restart-Process
{
    $Script:ExectueLegalHold="Y" #Used in Get-ENo to prevent asking for the Employee Number
    $Global:chkEnable = ""
    $Global:chkModify = ""
    $Global:chkRemove = ""
    $Global:strUserPath = ""
    $Global:u = ""
    $Global:inf = ""
    $Global:MBXAccess = ""
    $strDN = 0
    $Global:Protection = "$True"
    $Global:FormRefresh = "N"

    Check-Reconnect
    If (($FirstPass -eq "Y") -or ($Global:Result -eq "OK") -or ($Global:cancelButton.Text -eq "Next Emp"))
    {
        Build-DefaultForm
        Add-FormStandardButtons
        $Global:OKButton.visible = $False
        Publish-Form
    }
}

############################################################
#
#Start Of Script

$FileName		= "LegalHold"
$LogDrive		= "E:"
$LogPath		= "\Automation"
$LogFolder		= "\" + $FileName
$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
$FirstPass      = "Y"
$cred           = ""

$Year = (get-date).ToString("yyyy")
$ReportPath = "E:\Automation\LegalHold\Report\" + $Year + "\"
If (Test-Path $ReportPath) {} else {New-Item -Path $ReportPath -ItemType Directory}

#CheckLogFiles
write-host "Starting Script Legal Hold Check " -ForegroundColor Magenta
write-host
Restart-Process
$FirstPass      = "N"

Do
{
    If (($Global:Result -eq "OK") -and ($Global:txtHost.Text.Length -ge 5))
    {
        #Add New Legal or 40 Day Hold
        Hide-ActionBoxes

        If (($Global:chk40Day.Checked -eq "Checked") -or ($Global:chkEnable.Checked -eq "Checked"))
        {
            $Global:txtAttorney = ""
            $Global:txtHRRep = ""
            $Global:form.Text = "Enable Legal Hold"
            $Global:OKButton.Text = "Enable"
            $Global:OKButton.visible = $True 
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
            else
            {
                $Global:Republish = "N"
            }
            $Global:chk40Day.Checked = ""
            $Global:chkEnable.Checked = ""
        }

        If ($Global:chkDisable.Checked -eq "Checked")
        {
            $LineToWrite = $WhoAmI + "`t" + "Disabling Legal Hold from Account"
            WriteReportEvent
            $Global:form.Text = "Disable Legal Hold"
            $Global:OKButton.Text = "Disable" 
            Disable-Hold
            $Global:chkDisable.Checked = $false
        } 
       
        If ($Global:chkModify.Checked -eq "Checked")
        {
            $LineToWrite = $WhoAmI + "`t" + "Modifying Retention Comment"
            WriteReportEvent
            $Global:form.Text = "Modify Retention Comment"
            $Global:OKButton.Text = "Modify" 
            Modify-CommentForm
            Publish-Form
            Modify-Comment
            $Global:chkModify.Checked = $false
        }

        If ($Global:chkRemove.Checked -eq "Checked")
        {
            Build-RelaseForm
            Publish-Form
            $Global:form.Text = "Remove Retention Comment"
            $Global:OKButton.Text = "Remove"
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

                If ($Global:locListBox.SelectedItem -like "*DBOwner*")
                {
                    Remove-DBOwnerDetails
                    $Global:Republish = "N"
                }
                else
                {
                    Commit-CommentChange
                    $Global:Republish = "Y"
                }
                $Global:chkRemove.Checked = $false
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
            $Global:form.Text = "UnHide from Address Book"
            $Global:OKButton.Text = "UnHide" 
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
            $Global:form.Text = "Hide from Address Book"
            $Global:OKButton.Text = "Hide" 
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
            $Global:form.Text = "Move OU"
            $Global:OKButton.Text = "Move" 
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
                If ($cred -eq "")
                {
                    $me = whoami
                    $CredENo = $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))
                    $AAct = "A" + $CredENo
                    $Pwd = read-host "Enter password for" $AAct "account" -AsSecureString
                    $Server = "usnbkadds006p.global.ul.com"
                    $cred= New-Object System.Management.Automation.PSCredential($AAct,$Pwd)
                }
                Set-ADObject -Identity $Global:u.DistinguishedName.value -ProtectedFromAccidentalDeletion $false -Credential $Cred
#                $Output = $wshell.Popup("Unable to move account due to insufficient privilege to remove AD Object Protection.  Please set manually move the account.",0,"AD Object Protection",0+32)
            }
            $Global:Republish = "Y"
            $Global:chkMove.Checked = ""
        }

        If ($Global:chkAccess.Checked -eq "Checked")
        {
            $LineToWrite = $WhoAmI + "`t" + "Removing Access From Mailbox"
            WriteReportEvent
            $Global:form.Text = "Remove Permissions"
            $Global:OKButton.Text = "Remove" 
            
            # Remove Mailbox Permissions
            If ($Global:u.ExtensionAttribute14.value.length -eq 0)
            {
                $Output = $wshell.Popup("There are other active legal holds on this mailbox access to this mailbox will be removed.",0,"Active Holds Exist",0+32)
            }

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
                                Remove-MailboxFolderPermission -Identity ($ENo + ":\" + $Folder) -User $DelegateTo -confirm:$False -ErrorAction SilentlyContinue |out-null
#                                Remove-MailboxFolderPermission -Identity ($ENo + ":\" + $Folder) -User $DelegateTo -AccessRights $RemoveMbxFldrAccess.AccessRights -confirm:$False -ErrorAction SilentlyContinue |out-null
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

        If ($Global:chkDBOwner.Checked -eq "Checked")
        {
            $Global:form.Text = "Update DBOwner"
            $Global:OKButton.Text = "Add Owner"
            Enable-NewHoldForm
            Publish-Form
            If ($Global:Result -eq "OK")
            {
                $LineToWrite = $WhoAmI + "`t" + "Adding User as a dbOwner to the LegalDBOwners.csv file"
                WriteReportEvent
                $LineToWrite = $WhoAmI + "`t" + "Preservation Details           :`t" + $Global:txtPrevName.Text
                WriteReportEvent
                $LineToWrite = $WhoAmI + "`t" + "DBOwer Type Details            :`t" + $Global:txtHRRep.Text + "`n"
                WriteReportEvent
                #Add to dbOwner File
                $Name = $Global:inf.Name
                If ($Name -like "*_*")
                {
                    $Loc = $Name.Indexof("_")
                    $Name = $Name.Substring(0,$Loc)
                }
#                $LineToWrite = '{0},{1},{2},{3},{4}' -f $Global:ENo,$Global:txtType.Text,$Name,$Global:txtPrevName.Text,$Global:txtHRRep.Text
                $LineToWrite = '{0},{1},{2},{3}' -f $Global:ENo,$Name,$Global:txtPrevName.Text,$Global:txtHRRep.Text
                Out-File -FilePath "E:\O365AdminShared\Data\LegalDBOwners.csv" -InputObject $LineToWrite -Append
                Update-InputFiles
            }
            $Global:lblPrevName.Visible=$False
            $Global:txtPrevName.Visible =$False
            $Global:lblHRRep.visible = $false
            $Global:txtHRRep.visible = $false
            $Global:chkdbOwner.Checked = ""
            $Global:Republish = "N"
        }

        If ($Global:chkReport.Checked -eq "Checked")
        {
            $Global:ReportFile = $ReportPath + "Report-" + $FileName + "LegalReport-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            $ReportFile = $Global:ReportFile
            $Global:form.Text = "Generate Report"
            $Global:OKButton.Text = "Report" 
            Do
            {
                GetAcctInfo($Global:ENo)
                ShowDetails
                $ButGetENo.visible = $true
                $Global:txtHost.ReadOnly = $false
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
    $Global:OKButton.Text = "Next Employee"
    $Global:ButGetENo.visible = $true
    $Global:form.Text = "Legal Hold Details" 
    $Global:txtHost.ReadOnly = $false
    Restart-Process
}while (($Global:Result -eq "OK") -or ($Global:cancelButton.Text -eq "Next Emp"))