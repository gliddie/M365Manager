<#
New Legal Hold script
    Enable
        for a new hold
        for an existing hold

        for multiple  (provide comma delimited list of employee #'s)
        for an individual

    Disable 
        for individual
        for multiple (if on multiple hold only remove comment)


    Get list of all holds (function)
    Get all individuals part of a hold (function include showing the list of users and the # of people on that hold)
    Get all DBOwners part of a hold (function include showing the list of DBOwners and the # of people on that hold)

    NewHold (function)
    Release (function)
    Enable (function)
#>

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
            $Script:chk180Day.Checked = $False
            $Script:chkStd.Checked = $False
            If ($Script:chkNewHold.checked -eq $True)
            {
                $Script:txtHoldName.Text = "40 Day Hold (PutNameHere) - "
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
            $Script:lblActHolds.Visible = $True
            $Script:txtActHolds.Visible = $True
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
            $Script:chk40Day.Checked = $False
            $Script:chkStd.Checked = $False
            If ($Script:chkNewHold.checked -eq $True)
            {
                $Script:txtHoldName.Text = "180 Day Hold (PutNameHere) - "
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
            $Script:chk40Day.Checked = $False
            $Script:chk180Day.Checked = $False
            If ($Script:chkNewHold.Checked -eq $True)
            {
                $Script:txtHoldName.Text = "(PutNameHere) Legal Hold - "
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
            Show-WhoSilent
            $Script:txtActHolds.Items.Clear()
            Foreach ($Hold in $Script:StdHolds)
            {
                [void] $Script:txtActHolds.Items.Add($Hold.HoldName)
            }
        })
        
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
            If (($Script:chkExistHold.Checked -eq $True) -or ($Script:chkStdRel.Checked -eq $True))
            {
                $Script:txtActHolds.Text = ""
                $Script:txtENos.Text = ""
                $Script:txtDBOwner.Text = ""
                $Script:txtRelENos.Text = "Enter comma separated ENo's"
                $Script:txtRelDBOwner.Text = "Enter comma separated ENo's"
                $Script:ButGetENo.visible = $true
            }
        })

    $Script:ButGetENo = New-Object Windows.Forms.Button
        $Script:ButGetENo.Location = New-object System.Drawing.Size(505,90)
        $Script:ButGetENo.Size = new-Object System.Drawing.Size(($Top-10),35)
        $Script:ButGetENo.Text = "Get Enabled Users"
        $Script:ButGetENo.visible = $false
        $form.Controls.Add($Script:ButGetENo)
        $Script:ButGetENo.Add_Click({                
            $Script:ButGetENo.visible = $false
            $FindHold = $Script:txtActHolds.SelectedItem + "*"
            
            If ($FindHold -ne "*")
            {
                $OwnFound = ""
                $ThisHold = $AllHold |where-object {$_.HoldName -like $FindHold} |Sort-Object EmpNo
                If ($ThisHold.Count -ne "")
                {
                    $Script:txtDBOwner.Text = ""
                    $Script:txtENos.Items.Clear()
                    $Script:txtDBOwner.Items.Clear()
                    foreach ($t in $ThisHold)
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
                    $Script:lblRelENos.Visible = $True
                    $Script:txtRelENos.Visible = $True
                    $Script:lblDBOwner.Visible = $True
                    $Script:txtDBOwner.Visible = $True
                    $Script:txtRelDBOwner.Visible = $True
                    $Script:lblRelDBOwner.Visible = $True
                    $HoldDetails |where-object {$_.HoldName -like $FindHold}
                    $Script:txtAttorney.Text = ($HoldDetails |where-object {$_.HoldName -like $FindHold}).Attorney
                    $Script:lblAttorney.Visible = $True
                    $Script:txtAttorney.Visible = $True
                    If ($Script:chkStd.checked -eq $True)
                    {
                        $Script:lblHRRep.Visible = $False
                        $Script:txtHRRep.Visible = $False
                    }
                    else
                    {
                        $Script:txtHRRep.Text = ($HoldDetails |where-object {$_.HoldName -like $FindHold}).HRRep
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
                    If ((($HoldDetails |where-object {$_.HoldName -like $FindHold}).Silent) -eq "True")
                    {
                        $Script:chkSilent.Checked = $True
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
                write-host "Invliad name of Legal Hold" -ForegroundColor Red
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
                    $Output = $wshell.Popup("Enter employee numbers separated by commas of individuals to add to the hold.",10,"Help Info",0+32)
                }
                else
                {
                    $Output = $wshell.Popup("Enter employee numbers separated by commas of individual to add to as a DBOwner.",10,"Help Info",0+32)
                }
            }
            $Global:okButton.Visible = $True
        })

    $Top = $Top + 30
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
}

Function Check-ReportFile
{
    If (Test-Path $ReportFile)
    {
    #  Do Nothing
    }
    else
    {
        $LineToWrite = "Legal Hold Activities for Emp No. " + $ENo + "[" + $Person[1] + " " + $Person[0] + "]`n"
        WriteReportEvent       
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
    $Script:txtENos.Text = ""
    $Script:txtDBOwner.Text = ""

    $Script:txtRelENos.Text = "Enter comma separated ENo's"
    $Script:txtRelDBOwner.Text = "Enter comma separated ENo's"
    $Script:txtActHolds.Text = "Select Hold from List"
    $Global:okButton.Visible = $False
}

Function Move-ADObjectOU
{
    $Script:MoveObj = "False"
    $ErrorActionPreference = "SilentlyContinue"
    If (($Script:u.ExtensionAttribute1.value -like "*Ex-*") -and (($Script:chkEnable.Checked -eq "Checked") -or ($Script:chkMove.Checked -eq "Checked")))
    {
        $OULoc = "ExtendedHold OU"
        write-host "Moving AD Object to the ExtendedHold OU"
        $LineToWrite = $WhoAmI + "`t" + "Moved to ExtendedHold from     :" + "`t" + $Script:strUserPath
        $Script:MoveObj = [bool](Move-ADObject ($Script:strUserPath -replace("LDAP://","")) 'OU=_ExtendedHold,DC=global,DC=ul,DC=com')
        WriteReportEvent
    }
    elseif (($Script:u.ExtensionAttribute1.value -like "*Ex-*") -and ($Script:chkDisable.Checked -eq "Checked"))
    {
        $OULoc = "Disabled OU"
        write-host "Moving AD Object to the Disabled OU"
        $LineToWrite = $WhoAmI + "`t" + "Moved to ExtendedHold from     :" + "`t" + $Script:strUserPath
        $Script:MoveObj = [bool](Move-ADObject ($Script:strUserPath -replace("LDAP://","")) 'OU=Disabled,DC=global,DC=ul,DC=com')
        WriteReportEvent
    }
    elseif ($Script:u.ExtensionAttribute1.value -notlike "*Ex-*")
    {
        $OULoc = "Active"
    }
    $ErrorActionPreference = "Continue"

    If ($Script:MoveObj -eq $True)
    {
        If ($OULoc -like "*Disabled*")
        {
            write-host "Moving AD Object to the Disabled OU"
            $Script:u.DistinguishedName = $Script:u.DistinguishedName.value.Remove($Script:u.DistinguishedName.value.IndexOf("OU=")) + "OU=Disabled,DC=global,DC=ul,DC=com"
            $LineToWrite = $WhoAmI + "`t" + "Moved to Disabled OU from     :" + "`t" + $Script:strUserPath
            WriteReportEvent
        }
        elseif ($OULoc -like "*Extend*")
        {
            write-host "Moving AD Object to the ExtendedHold OU"
            $Script:u.DistinguishedName = $Script:u.DistinguishedName.value.Remove($Script:u.DistinguishedName.value.IndexOf("OU=")) + "OU=_ExtendedHold,DC=global,DC=ul,DC=com"
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

Function Execute-Hold
{
    If ($Script:Mbx.LitigationHoldEnabled -ne $True)
    {
    #This person is not on legal hold
        Set-Mailbox $ENo -LitigationHoldEnabled $true
        write-host "Enabling legal hold for: " $ENo
        $LineToWrite = $WhoAmI + "`t" + "Enabling Legal Hold for " + $ENo + " with Retention Comment of: " + $Script:RetCmt + "`n"
        WriteReportEvent
        $success = [bool](Add-DistributionGroupMember iprosearch@ul.onmicrosoft.com -Member $ENo -BypassSecurityGroupManagerCheck -ErrorAction SilentlyContinue)
        If ($Success -eq $True)
        {
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
        If (($Script:u.ExtensionAttribute14.value.length -eq 0) -or ($Script:u.ExtensionAttribute14.value -like "Emergency*") -or ($Script:u.ExtensionAttribute14.value -like "Standard*"))
        {
            $Script:NewCmt = $Script:RetCmt
            $Recap = $Recap + $ENo + " - Enabled legal hold for this individual`n"
        }
        else
        {
            $Recap = $Recap + $ENo + " - Added additional legal hold for this individual`n"
            $Script:StoredCmt = $Script:u.ExtensionAttribute14.value
            $Script:NewCmt =  $Script:u.ExtensionAttribute14.value + ", " + $Script:RetCmt
        }
        $LineToWrite = $WhoAmI + "`t" + "New ADRetention Activity       :`t" + $Script:NewCmt + "`n"
    }
    else
    {
        #Build new retention comment based on the O365Mailbox Ext14 attribute
        $Script:StoredCmt = $Script:Mbx.CustomAttribute14 + $Script:RetCmt
        $LineToWrite = $WhoAmI + "`t" + "New O365MbxRetention Activity         :`t" + $Script:NewCmt + "`n"
        $Recap = $Recap + $ENo + " - No AD Acount Updated O365 Mailbox CustomAttribute14 for this individual`n"
    }
    WriteReportEvent
    Commit-CommentChange
}

Function Release-FromHold
{
    #This is to release individuals from Legal Holds - checks to so see if this individual is on other holds

}

Function Enable-NewHold
{
}

Function Enable-AddlStaff
{
}

Function Release-Hold
{
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

Function Build-RetentComment
{
    If ($HoldDetails -like ("*"+ $HName +"*"))
    {
        #DoNothing this retention comment is arlready listedi n the file
#        write-host "This retention comment already exists" -ForegroundColor Red
    }
    else
    {
        #Add this new hold to the list of all holds
        If ($Script:chkSilent.Checked -eq $True)
        {
            $Text = "{0},{1},{2},{3},{4},{5}" -f "NonDisclosed - " + $HName, $Script:txtHRRep.Text, $Script:txtAttorney.Text, $Script:chkHHold.Checked, $Script:chkSilent.Checked, (get-Date).ToString("yyyy-MM-dd")
        }
        else
        {
            $Text = "{0},{1},{2},{3},{4},{5}" -f $HName, $Script:txtHRRep.Text, $Script:txtAttorney.Text, $Script:chkHHold.Checked, $Script:chkSilent.Checked, (get-Date).ToString("yyyy-MM-dd")
        }
        Out-File -FilePath "\\usnbkemes100p\e$\O365AdminShared\Data\Legal-AllHoldDetails.csv" -InputObject $Text -Append
    }
    
    $Script:FullRetCmt = ""

    #Enable Hold or Add Additional Comment
    If ($Script:chkSilent.Checked -eq $True)
    {
        $Script:RetCmt = "NonDisclosed - " + $HName + " - " + $Script:txtHRRep.Text + "\" + $Script:txtAttorney.Text + "\"
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

    write-host "Retention Comment: " $Script:RetCmt
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
            $Script:u.ExtensionAttribute14 = $Script:NewCmt
            $Script:u.CommitChanges()
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
#    $ConfDisable = 6
    $ADU = Get-ADUser $ENo -Properties *
    If ($ADU.ExtensionAttribute14 -like "*,*")
    {
        write-host "This individual is on multiple holds removing only details for this hold" -ForegroundColor Yellow
        $holds = $ADU.ExtensionAttribute14 -replace (", ","`n`t ")
        write-host "Holds this inidividual is enabled for: `n`t"$Holds
#       Return value 6 = Yes, 7 = No     
    }
    else
#    If ($ConfDisable -eq 6)
    {

        Write-host "This user is not on multple Holds" -ForegroundColor Green
<#        
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
#>
    }
} 

Function Log-AcctDetails
{
    $LineToWrite = $WhoAmI + "`t" + "DisplayName                    :" + "`t" + $Script:Mbx.DisplayName
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
    $LineToWrite = $WhoAmI + "`t" + "Hidden from Address Book AD    :" + "`t" + $Script:u.msExchHideFromAddressLists.value
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "O365 Retention Comment         :" + "`t" + $Script:Mbx.CustomAttribute14
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "AD Retention Comment           :" + "`t" + $Script:ADCmt
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "AD Object Protected            :" + "`t" + $Script:usrDetails.ProtectedFromAccidentalDeletion
    WriteReportEvent
}

# Gets AD ACcount Properties
Function GetAcctInfo
{
    $Script:ADCmt = ""
    $Script:u = ""
    If ($ADExists -eq "True")
    {
        $ObjExists = [bool]($Script:UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion -ErrorAction silentlyContinue)
        If ($ObjExists -eq $False)
        {
            $cnt = 0
            Do
            {
                start-sleep -Seconds 5
                $strDN = (Get-ADUser $ENo).DistinguishedName
#                $strDN = GetUserDN $ENo -ErrorAction SilentlyContinue
                $ObjExists = [bool]($Script:UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion -ErrorAction silentlyContinue)
                If ($ObjExists -eq $False)
                {
                    #Attempting to find the AD object
                    $Cnt++
                }
            } while (($Cnt -le 5) -and ($ObjExists -eq $False))

            If ($Cnt -ge 5)
            {
                write-host "AD Record not found"
                $LineToWrite = $WhoAmI + "`t" + "AD Record Not Found!"
                WriteReportEvent
            }
        }
     }

#    $MbxExists = [bool](get-mailbox -identity $ENo -ErrorAction SilentlyContinue)
    If ($MbxExists -eq "True")
    {
        $Script:inf = get-mailbox $ENo
        If ([bool](get-ADUser -Filter {SamAccountName -eq $ENo} -ErrorAction SilentlyContinue))
        {
            $Script:strUserPath = [string]::format("LDAP://{0}", $strDN)
            $Script:u = new-object System.DirectoryServices.DirectoryEntry($Script:strUserPath)
            $Script:ADCmt = $Script:u.ExtensionAttribute14.value
        }
        else
        {
            $Script:ADCmt = "*****No Active Directory Account For This User*****"
            $Script:strUserPath = "No Active Directory Account for this User"            
        }

        $LineToWrite = $WhoAmI + "`t" + "Account Details Before Adding New Hold: "
        WriteReportEvent
        Log-AcctDetails
        PrtMBxAccess($Script:MbxAccess)
    }
    else
    {
        write-host""
        write-host "Mailbox does not exist for Emp#" $ENo -ForegroundColor Red
    }
}

# This prints details regarding mailbox access to the console and report file
Function PrtMBxAccess()
{
    $Script:MbxAccess = Get-MailboxPermission $ENo |Where-Object {$_.User -like "*global.ul.com"}
    $Script:MbxFldrAccess = Get-MailboxFolderPermission $ENo |Where-Object {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous")}
    $Script:MbxAccessForm = "(None)"
    foreach ($Script:MbxAccess in $Script:MbxAccess)
    {
#       write-host "Accounts Granted Access       : " $Script:MbxAccess.User -ForegroundColor Yellow
        $UsrExists = [bool](Get-mailbox $Script:MbxAccess.User -ErrorAction SilentlyContinue)
        If ($Script:MbxAccessForm -eq "(None)")
        {
            If ($UsrExists -eq $True)
            {
                $Script:MbxAccessForm = (get-mailbox $Script:MbxAccess.User).PrimarySMTPAddress + " (" + $Script:MbxAccess.AccessRights + ")"
                $SupMbx = (get-mailbox $Script:MbxAccess.User).PrimarySMTPAddress
                $LineToWrite = $WhoAmI + "`tUser Granted Mailbox Permission:" + "`t" + $(get-mailbox $Script:MbxAccess.User).PrimarySMTPAddress + "- PermissionLevel: " + $Script:MbxAccess.AccessRights
                WriteReportEvent
            }
            else
            {
                $Script:MbxAccessForm = $Script:MbxAccess.User + " (" + $Script:MbxAccess.AccessRights + ")"
            }
        }
        else
        {
            $Script:MbxAccessForm = $Script:MbxAccessForm + ", " + (get-mailbox $Script:MbxAccess.User).PrimarySMTPAddress + " (" + $Script:MbxAccess.AccessRights + ")"
            $SupMbx = $SupMbx + ", " + (get-mailbox $Script:MbxAccess.User).PrimarySMTPAddress
            $LineToWrite = $WhoAmI + "`tUser Granted Mailbox Permission:" + "`t" + $(get-mailbox $Script:MbxAccess.User).PrimarySMTPAddress + "- PermissionLevel: " + $Script:MbxAccess.AccessRights
            WriteReportEvent
        }
    }

    foreach ($Script:MbxFldrAccess in $Script:MbxFldrAccess)
    {
        If ($Script:MbxAccessForm -eq "(None)")
        {
            If ($Script:MbxFldrAccess.User.DisplayName -like "MBX*")
            {
                $Script:MbxAccessForm = (get-distributiongroup $Script:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + " (" + $Script:MbxFldrAccess.AccessRights + ")"
                $SupMbx = (get-distributiongroup $Script:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress
                $LineToWrite = $WhoAmI + "`tUser Granted Folder Permission :" + "`t" + $(get-distributiongroup $Script:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + "- PermissionLevel: " + $Script:MbxFldrAccess.AccessRights
            }
            else
            {
                $UsrExists = [bool](Get-mailbox $Script:MbxFldrAccess.User.DisplayName -ErrorAction SilentlyContinue)
                If ($UsrExists -eq $True)
                {
                    $Script:MbxAccessForm = $Script:MbxAccessForm + (get-mailbox $Script:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + " (" + $Script:MbxFldrAccess.AccessRights + ")"
                    $SupMbx = (get-mailbox $Script:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress
                }
                else
                {
                    $Script:MbxAccessForm = $Script:MbxFldrAccess.User.DisplayName + " (" + $Script:MbxFldrAccess.AccessRights + ")"
                }
                $LineToWrite = $WhoAmI + "`tUser Granted Folder Permission :" + "`t" + $(get-mailbox $Script:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + "- PermissionLevel: " + $Script:MbxFldrAccess.AccessRights
            }
            WriteReportEvent
        }
        else
        {
            If ($Script:MbxFldrAccess.User.DisplayName -like "MBX*")
            {
                $Script:MbxAccessForm = $Script:MbxAccessForm + (get-distributiongroup $Script:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + " (" + $Script:MbxFldrAccess.AccessRights + ")"
                $SupMbx = (get-distributiongroup $Script:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress
                $LineToWrite = $WhoAmI + "`tUser Granted Folder Permission :" + "`t" + $(get-distributiongroup $Script:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + "- PermissionLevel: " + $Script:MbxFldrAccess.AccessRights
            }
            else
            {
                $Script:MbxAccessForm = $Script:MbxAccessForm + ", " + (get-mailbox $Script:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + " (" + $Script:MbxFldrAccess.AccessRights + ")"
                $SupMbx = $SupMbx + ", " + (get-mailbox $Script:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress
                $LineToWrite = $WhoAmI + "`tUser Granted Folder Permission :" + "`t" + $(get-mailbox $Script:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + "- PermissionLevel: " + $Script:MbxFldrAccess.AccessRights
            }
            WriteReportEvent
        }
    }

    If ($Script:MbxAccessForm -ne "(None)")
    {
        if (((get-date).AddHours(-1) -le $Global:inf.litigationholddate) -and ($Global:inf.CustomAttribute1 -eq "T"))
        {
            $SendTo = $Global:LegalTeamMsgs
            $SendCC = "Sandi.Glazebrook@ul.com"
            if ($Mbx.CustomAttribute1 -like "*Employee*")
            {
                $MessageSubject = "Access Previously Granted to Terminated Employee Mailbox and OneDrive"
                $MessageBody = "While placing this terminated account on Legal Hold access was previously granted to the Mailbox and OneDrive information.  The details are provided below.<ul type=""disc""><li>Terminated " + $Global:inf.CustomAttribute1 + " details " + $Mbx.Alias + " - " + $Global:inf.DisplayName + "</li><li>Individual given access to this information " + $SupMbx.Alias + " - " + $SupMbx.DisplayName + "</li></ul>UL Account Provisioning Team</li></ul>"
            }
            else
            {
                $MessageSubject = "Access Previously Granted to Terminated Non-Employee Mailbox"
                $MessageBody = "While placing this terminated account on Legal Hold access was previously granted to thes Mailbox information.<ul type=""disc""><li>Terminated " + $Global:inf.CustomAttribute1 + " details " + $Global:inf.Alias + " - " + $Global:inf.DisplayName + "</li><li>Individual given access to this information " + $SupMbx.Alias + " - " + $SupMbx.DisplayName + "</li></ul>UL Account Provisioning Team</li></ul>"
            }
            invoke-expression -Command .\SendSMTPMessage.ps1
        }
    }
}

Function Set-ADObjProtect
{
#    $Script:SetObj = ""
    $ErrorActionPreference = "SilentlyContinue"
    switch ($Script:Protection)
    {
        "True"
        {
#            Set-ADObject -Identity $strDN -ProtectedFromAccidentalDeletion $True
#            This is not working
            Set-ADObject -Identity $Global:u.DistinguishedName.value -ProtectedFromAccidentalDeletion $True -Credential $Global:AdmLiveCred
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


#  Writes events to the Report File
function WriteReportEvent
{
    If (($null -eq $ReportFile) -or ($ReportFile -eq ""))
    {
        $ReportFile = $Script:ReportFile
    }
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LineToWrite
	Out-File -filepath $ReportFile -append -noClobber -inputObject $RecordEvent
} #end WriteReportEvent


$whoami = whoami
$HoldDetails = import-csv "e:\O365AdminShared\Data\Legal-AllHoldDetails.csv"
#Copy all hold details to the O365AdminShared\Data directory to add new activites for the day
$Dailyfile = ("\\usnbkutil100p\d$\ReportHistory\LegalHoldHistory\2023\LegalHoldReport-" + (get-date).Tostring("yyyy-MMdd") + ".csv")
$ActFile = "\\usnbkemes100p\e$\O365AdminShared\Data\LegalHoldReport.csv"
$today = (get-date).toString("yyyy-MM-dd")
If (Test-Path $ActFile)
{
    $whencreated = ((get-ChildItem -File $ActFile).CreationTime).ToString("yyyy-MM-dd")
    If ($whencreated -ne $today)
    {
        Copy-Item $DailyFile -Destination $ActFile -force
    }
}
else
{
    Copy-Item $DailyFile -Destination $ActFile -force
}

#$AllHold = import-csv ("\\usnbkutil100p\d$\ReportHistory\LegalHoldHistory\2023\LegalHoldReport-" + (get-date).Tostring("yyyy-MMdd") + ".csv")
$AllHold = import-csv $ActFile
#$ReportFile = "c:\temp\NewLegalHoldDetails.csv"
$wshell = New-Object -ComObject Wscript.Shell
Build-HoldMenu
$Global:okButton.Visible = $False
Publish-Form

#$Year = (get-date).ToString("yyyy")
#$ReportPath = "E:\Automation\LegalHold\Report\" + $Year + "\"
#If (Test-Path $ReportPath) {} else {New-Item -Path $ReportPath -ItemType Directory}

If ($Global:Result -eq "OK")
{
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
            $Recap = ""
            If ($Script:chkNewHold.Checked -eq $True)
            {
                Build-RetentComment
                write-host "`tEnable staff for new hold"
                write-host "`t          New Hold Name: " $Script:txtHoldName.Text
                write-host "`t Full Retention Comment: " $Script:RetCmt
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
                    write-host "Number of staff to be Enabled: " $EnabSel.Count
                    Foreach ($ENo in $EnabSel)
                    {
                        If (($ENo.length -ge 5) -and ($ENo -notlike "Enter*"))
                        {
                            $ADExists = $False
                            $StrDN = ""
                            $Script:StoredCmt = ""
                            $ErrorActionPreference = "SilentlyContinue"
                            $ADExists = [bool]($strdn = (Get-ADUser $ENo).DistinguishedName)
                            $ErrorActionPreference = "Continue"
                            $MbxExists = [bool]($Script:Mbx = get-mailbox $ENo)
                            $Person = $Script:Mbx.DisplayName -split ", "
                            write-host "Processing Hold for: " $ENo - $Person[1] $Person[0] -ForegroundColor Cyan
                            $NewHold = "*" + $Script:txtHoldName.Text.TrimEnd(" - ") + "*"
                            If ($Script:Mbx.CustomAttribute14 -like $NewHold)
                            {
                                write-host "`tThis individual is already enabled for this hold" -ForegroundColor Red
                                $LineToWrite = $WhoAmI + "`t" + "This individual is already enabled for this hold"
                                WriteReportEvent
                                $Recap = $Recap + $ENo + " - No change made - already enabled for this hold`n"
                            }
                            else
                            {
                                $ReportFile = "e:\Automation\LegalHold\Report\LegalHoldDetails-" + $ENo + ".csv"
                                Check-ReportFile
                                GetAcctInfo
                                Execute-Hold
                                If ($ADExists -eq $True)
                                {
                                    $cnt = 0
                                    write-host "Verifying that the hold details have been added...." -NoNewline
                                    Do
                                    {
                                        write-host "...." -NoNewline
                                        start-sleep -Seconds 7
                                        $Script:u = new-object System.DirectoryServices.DirectoryEntry($Script:strUserPath)
                                        $Cnt++
                                    } while (($Script:u.ExtensionAttribute14 -ne $NewHold) -and ($cnt -le 10))
                                    $Script:ADCmt = $Script:u.ExtensionAttribute14.value
                                }
                                If ($cnt -ge 10)
                                {
                                    write-host "check the retention comment on this account"
                                }
                                $LineToWrite = $WhoAmI + "`t" + "Account Details After Added New Hold: "
                                WriteReportEvent
                                Log-AcctDetails
                                write-host "`tEnabled staff for new hold" $ENo -ForegroundColor Green

                                # Add details to the LegalHoldReport
                                #Enabled,EmpNo,Name,EmpType,#ofHolds,HoldName,HoldEnabled,TermDate
                                $Text = {0},{1},{2},{3},{4},{5},{6},{7} -f ("".$ENo,($Person[1] + " " + $Person[0]),"","",$NewHold,"","")
                                Out-File -FilePath $ActFile -InputObject $Text -Append
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
                    write-host "Number of DBOwners to be Enabled: " $EnabSel.Count
                    $DBOwner = "Yes"
                    Foreach ($ENo in $EnabSel)
                    {
                        $DBOFile = import-csv "E:\O365AdminShared\Data\LegalDBOwners.csv" | Where-Object {$_.EmpNo -eq $ENo}
                        If (($DBOfile.EmpNo -eq $ENo) -and ($DBOFile.CaseName -eq (($Script:txtHoldName.Text).TrimEnd(" - "))))
                        {
                            write-host "`tDetails already stored for this user and this case" -ForegroundColor Red
                            $Recap = $Recap + $ENo + " - Individual already listed as a dBOwner for this legal hold`n"
                        }
                        else
                        {
                            $ReportFile = "e:\Automation\LegalHold\Report\LegalHoldDetails-" + $ENo + ".csv"
                            Check-ReportFile
                            $ErrorActionPreference = "SilentlyContinue"
                            $ADExists = [bool]($strDN = (Get-ADUser $ENo).DistinguishedName)
                            $MbxExists = [bool]($Script:Mbx = get-mailbox $ENo)
                            $ErrorActionPreference = "Continue"
                            If ($ADExists -eq $True)
                            {
                                $LineToWrite = $WhoAmI + "`t" + "Adding User as a dbOwner to the LegalDBOwners.csv file"
                                WriteReportEvent
                                $LineToWrite = $WhoAmI + "`t" + "Preservation Details           :`t" + $Script:txtHoldName.Text
                                WriteReportEvent
                                #Add to dbOwner File
                                $Person = $Script:Mbx.DisplayName -split ", "
                                If ($Name -like "*_*")
                                {
                                    $Loc = $Name.Indexof("_")
                                    $Name = $Name.Substring(0,$Loc)
                                }
                                $LineToWrite = '{0},{1},{2},{3}' -f $ENo,($Person[1] + " " + $Person[0]),(($Script:txtHoldName.Text).TrimEnd(" - ")),""
                                Out-File -FilePath "E:\O365AdminShared\Data\LegalDBOwners.csv" -InputObject $LineToWrite -Append
                            }
                            write-host "Add DBOwner for new hold" $ENo
                            $Recap = $Recap + $ENo + " - Added individual as a dBOwner for this legal hold`n"
                            $Text = {0},{1},{2},{3},{4},{5},{6},{7} -f ("DBOwner".$ENo,($Person[1] + " " + $Person[0]),"","",(($Script:txtHoldName.Text).TrimEnd(" - ")),"","")
                            Out-File -FilePath $ActFile -InputObject $Text -Append
                        }
                    }
                }
                else
                {
                    write-host "No DBOwners identified"
                }

                write-host "Changes made for this hold activity:`n`n" -ForegroundColor Green
                $Recap
            }
        }
        else
        {
            write-host "Process cancelled" -ForegroundColor Red
        }
    }

##############
    If ($Script:chkExistHold.Checked -eq $True)
    {
    #    Enable additional staff for an existing legal hold
        write-host "Enable additional staff for existing hold"
            
        write-host "Enable additional staff for existing hold"
        write-host "        Hold Name: " $Script:txtActHolds.SelectedItem
        write-host "    Staff to Hold: " $Script:txtRelENos.Text
        write-host "  DBOwners to Add: " $Script:txtRelDBOwner.Text
        write-host "Attorney Assigned: " $Script:txtAttorney.Text
        write-host "   HRRep Assigned: " $Script:txtHRRep.Text
        write-host "      Silent Hold: " $Script:chkSilent.Checked
        write-host "Number of staff to be enabled: " $RelSel.Count
            
        If ($Script:chkStdHold.Checked -eq $False)
        {
            write-host "HRRep Assigned: " $Script:txtHRRep.Text
        }

        $Recap = ""
        $HName = ($Script:txtActHolds.SelectedItem).TrimEnd(" - ")
        Build-RetentComment
        $EnabSel = ($Script:txtRelENos.Text -replace(" ","")) -split ","
        write-host "Number of staff to be Enabled: " $EnabSel.Count
        Foreach ($ENo in $EnabSel)
        {
            If (($ENo.length -ge 5) -and ($ENo -notlike "Enter*"))
            {
                $ADExists = $False
                $StrDN = ""
                $Script:StoredCmt = ""
                $ErrorActionPreference = "SilentlyContinue"
                $ADExists = [bool]($strDN = (Get-ADUser $ENo).DistinguishedName)
                $ErrorActionPreference = "Continue"
                $MbxExists = [bool]($Script:Mbx = get-mailbox $ENo)
                $Person = $Script:Mbx.DisplayName -split ", "
                write-host "Processing Hold for: " $ENo - $Person[1] $Person[0] -ForegroundColor Cyan
                $NewHold = "*" + $Script:txtActHolds.SelectedItem.TrimEnd(" - ") + "*"
                If ($Script:Mbx.CustomAttribute14 -like $NewHold)
                {
                    write-host "`tThis individual is already enabled for this hold" -ForegroundColor Red
                    $LineToWrite = $WhoAmI + "`t" + "This individual is already enabled for this hold"
                    WriteReportEvent
                    $Recap = $Recap + $ENo + " - No change made - already enabled for this hold`n"
                }
                else
                {
                    $ReportFile = "e:\Automation\LegalHold\Report\LegalHoldDetails-" + $ENo + ".csv"
                    Check-ReportFile
                    GetAcctInfo
                    Execute-Hold
                    If ($ADExists -eq $True)
                    {
                        $cnt = 0
                        write-host "Verifying that the hold details have been added...." -NoNewline
                        Do
                        {
                            write-host "...." -NoNewline
                            start-sleep -Seconds 7
                            $Script:u = new-object System.DirectoryServices.DirectoryEntry($Script:strUserPath)
                            $Cnt++
                        } while (($Script:u.ExtensionAttribute14 -ne $NewHold) -and ($cnt -le 10))
                        $Script:ADCmt = $Script:u.ExtensionAttribute14.value
                    }
                    If ($cnt -ge 10)
                    {
                        write-host "check the retention comment on this account"
                    }
                    $LineToWrite = $WhoAmI + "`t" + "Account Details After Added New Hold: "
                    WriteReportEvent
                    Log-AcctDetails
                    write-host "`tEnabled staff for new hold" $ENo -ForegroundColor Green
                    
                    # Add details to the LegalHoldReport
                    #Enabled,EmpNo,Name,EmpType,#ofHolds,HoldName,HoldEnabled,TermDate
                    $Text = {0},{1},{2},{3},{4},{5},{6},{7} -f ("".$ENo,($Person[1] + " " + $Person[0]),"","",$NewHold,"","")
                    Out-File -FilePath $ActFile -InputObject $Text -Append
                }

                write-host "Added additional staff for existing hold" $ENo
            }
        }
        $EnabSel = ($Script:txtRelDBOwner.Text -replace(" ","")) -split ","
        write-host "Number of DBOwners to be Enabled: " $EnabSel.Count
        Foreach ($ENo in $EnabSel)
        {

<#

#>


            write-host "Add additional DBOwner for existing hold" $ENo
        }
    }

    If ($Script:chkStdRel.Checked -eq $True)
    {
    #    Release from a legal hold
    #if entire hold is released remove hold details from the HodDetails fle
    #display user account details after changes are made
        If (($Script:txtRelENos.Text -notlike "Enter*" ) -and ($Script:txtRelENos.Text.Length -gt 0))
        {
            write-host "Release staff from existing hold"
            write-host "Hold Name: " $Script:txtActHolds.SelectedItem
            If ($Script:txtRelENos.Text -eq "All")
            {
                $RelSel = $Script:txtENos.Items
                write-host "Release all enabled staff from this hold"
                write-host "Number of staff to be released: " $RelSel.Count
                Foreach ($u in $RelSel)
                {
                    $ENo = $u.Substring(0,$u.Indexof(" "))
                    write-host "Release all staff from this hold" $ENo
                    Disable-Hold
                }
            }
            else
            {
                $RelSel = ($Script:txtRelENos.Text -replace(" ","")) -split (",")
                write-host "Number of staff to be released: " $RelSel.Count
                Foreach ($ENo in $RelSel)
                {
                    write-host "Release selected staff from this hold" $ENo
                    Disable-Hold 
                }
            }
        }

        If (($Script:txtRelDBOwner.Text -notlike "Enter*") -and ($Script:txtRelDBOwner.Text.Length -gt 0))
        {
            If ($Script:txtRelDBOwner.Text -eq "All")
            {
                write-host "Release all enabled DBOwners from this hold"
                $RelSel = $Script:txtDBOwner.Items
                write-host "Number of staff to be released: " $RelSel.Count
                Foreach ($u in $RelSel)
                {
                    $ENo = $u.Substring(0,$u.Indexof(" "))
                    write-host "Release DBOwners from this hold" $ENo
                    Disable-Hold
                } 
            }
            else
            {
                write-host "Release selected DBOwners from this hold"
                $RelSel = ($Script:txtRelDBOwner.Text -replace(" ","")) -split(",")
                write-host "Number of staff to be released: " $RelSel.Count
                Foreach ($ENo in $RelSel)
                {
                    write-host "Release selected DBOwner from this hold" $ENo
                    Disable-Hold
                }  
            }
        }
    }
}
else
{
    write-host "Process cancelled" -ForegroundColor Red
}