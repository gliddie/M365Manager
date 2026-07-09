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

    $Col1 = 340
    $Left = 100
    $TextLeft = 130
  
    ## New Hold
    $Script:chkNewHold = New-Object System.Windows.Forms.RadioButton
        $Script:chkNewHold.Text = "Enable for New Legal Hold" 
        $Script:chkNewHold.Top = 30 ; $Script:chkNewHold.Left = $Left; $Script:chkNewHold.Width=150 ;$Script:chkNewHold.AutoSize = $true 
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
        })

    ## Existng Hold
    $Script:chkExistHold = New-Object System.Windows.Forms.RadioButton
        $Script:chkExistHold.Text = "Enable Additional Staff for Existing Hold"  
        $Script:chkExistHold.Top = 50 ; $Script:chkExistHold.Left = $Left; $Script:chkExistHold.Width=150 ;$Script:chkExistHold.AutoSize = $true 
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
        })         

    ## Release Standard Hold
    $Script:chkStdRel = New-Object System.Windows.Forms.RadioButton
        $Script:chkStdRel.Text = "Release from Legal Hold" 
        $Script:chkStdRel.Top = 70 ; $Script:chkStdRel.Left = $Left; $Script:chkStdRel.Width=150 ;$Script:chkStdRel.AutoSize = $true 
        $form.Controls.Add($Script:chkStdRel)    # Add to Form
        $Script:chkStdRel.Add_Click({
            Hide-All
            Uncheck-Options
#            $Script:chkforUser.Visible = $True
            $Script:chkforEntireStd.Visible = $True
            $Script:chkforEntire40180.Visible = $True
            $Script:lblAttorney.Text = "Released by:"
            $Script:lblENos.Text = "Staff on Hold:"
            $Script:txtRelENos.Left = 330; $Script:txtRelENos.Width = 165
            $Script:txtRelDBOwner.Left = 330; $Script:txtRelDBOwner.Width = 165
        })         

    ## For Entire Standard Hold
    $Script:chkforEntireStd = New-Object System.Windows.Forms.Checkbox
        $Script:chkforEntireStd.Text = "for Standard Hold"  
        $Script:chkforEntireStd.Top = 40 ; $Script:chkforEntireStd.Left = $Col1; $Script:chkforEntireStd.Width=150 ;$Script:chkforEntireStd.AutoSize = $true 
        $Script:chkforEntireStd.Visible = $False
        $form.Controls.Add($Script:chkforEntireStd)    # Add to Form 
        $Script:chkforEntireStd.Add_Click({
            Hide-All
            $Script:chkforEntireStd.Checked = $True
            $Script:chkforEntireStd.Visible = $True
            $Script:chkforEntire40180.Checked = $False
            $Script:chkforEntire40180.Visible = $True
            $Script:lblActHolds.Visible = $True
            $Script:txtActHolds.Visible = $True
            $Script:txtActHolds.Visible = $True
            $Script:txtActHolds.Items.Clear()
            Foreach ($Hold in $Script:StdHolds)
            {
                [void] $Script:txtActHolds.Items.Add($Hold.HoldName)
            }
        })

    ## For Entire 40/180 Hold
    $Script:chkforEntire40180 = New-Object System.Windows.Forms.Checkbox
        $Script:chkforEntire40180.Text = "for 40 or 180 Day Hold"  
        $Script:chkforEntire40180.Top = 60 ; $Script:chkforEntire40180.Left = $Col1; $Script:chkforEntire40180.Width=150 ;$Script:chkforEntire40180.AutoSize = $true 
        $Script:chkforEntire40180.Visible = $False
        $form.Controls.Add($Script:chkforEntire40180)    # Add to Form 
        $Script:chkforEntire40180.Add_Click({
            Hide-All
            $Script:chkforEntireStd.Checked = $False
            $Script:chkforEntireStd.Visible = $True
            $Script:chkforEntire40180.Checked = $True
            $Script:chkforEntire40180.Visible = $True
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

    ## 40 Day Hold
    $Script:chk40Day = New-Object System.Windows.Forms.Checkbox
        $Script:chk40Day.Text = "40 Day Hold"  
        $Script:chk40Day.Top = 30 ; $Script:chk40Day.Left = $Col1; $Script:chk40Day.Width=150 ;$Script:chk40Day.AutoSize = $true 
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
                $Script:txtAttorney.Text = "B.Bogearts"
            }
            Show-WhoSilent
            $Script:txtActHolds.Items.Clear()
            Foreach ($Hold in $Script:40DayHolds)
            {
                [void] $Script:txtActHolds.Items.Add($Hold.HoldName)
            }
        })

    ## 180 Day
    $Script:chk180Day = New-Object System.Windows.Forms.Checkbox
        $Script:chk180Day.Text = "180 Day Hold"  
        $Script:chk180Day.Top = 50 ; $Script:chk180Day.Left = $Col1; $Script:chk180Day.Width=150 ;$Script:chk180Day.AutoSize = $true 
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

    ## Standard Hold
    $Script:chkStd = New-Object System.Windows.Forms.Checkbox
        $Script:chkStd.Text = "Standard Hold"  
        $Script:chkStd.Top = 70 ; $Script:chkStd.Left = $Col1; $Script:chkStd.Width=150 ;$Script:chkStd.AutoSize = $true 
        $Script:chkStd.Visible = $False
        $form.Controls.Add($Script:chkStd)    # Add to Form
        $Script:chkStd.Add_Click({
            $Script:chk40Day.Checked = $False
            $Script:chk180Day.Checked = $False
            If ($Script:chkNewHold.Checked -eq $True)
            {
                $Script:txtHoldName.Text = "(PutNameHere) Legal Hold - "
                $Script:lblHoldName.Visible = $True
                $Script:txtHoldName.Visible = $True
            }
            Show-WhoSilent
            $Script:txtActHolds.Items.Clear()
            Foreach ($Hold in $Script:StdHolds)
            {
                [void] $Script:txtActHolds.Items.Add($Hold.HoldName)
            }
        })
        
    ## List of Holds
    $Script:lblActHolds = New-Object System.Windows.Forms.Label   
        $Script:lblActHolds.Text = "Active Legal Holds:"
        $Script:lblActHolds.Top = 100; $Script:lblActHolds.Left = 10; $Script:lblActHolds.Width=150 ;$Script:lblActHolds.AutoSize = $true
        $Script:lblActHolds.Visible = $False 
        $form.Controls.Add($Script:lblActHolds)    # Add to Form 
        # 
        $Script:txtActHolds = New-Object Windows.Forms.ComboBox
        $Script:40DayHolds = $HoldDetails |Where-Object {$_.HoldName -like "*40 Day*"} |Sort-Object Holdname
        $Script:180DayHolds = $HoldDetails |Where-Object {$_.HoldName -like "*180 Day*"} |Sort-Object Holdname
        $Script:StdHolds = $HoldDetails |Where-Object {($_.HoldName -notlike "*40 Day*") -and ($_.HoldName -notlike "*180 Day*")} |Sort-Object Holdname
        $Script:txtActHolds.Top = 100; $Script:txtActHolds.Left = $TextLeft; $Script:txtActHolds.Width = 340;
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
        $Script:ButGetENo.Location = New-object System.Drawing.Size(485,90)
        $Script:ButGetENo.Size = new-Object System.Drawing.Size(80,35)
        $Script:ButGetENo.Text = "Get Enabled Users"
        $Script:ButGetENo.visible = $false
        $form.Controls.Add($Script:ButGetENo)
        $Script:ButGetENo.Add_Click({                
            $Script:ButGetENo.visible = $false
            $FindHold = $Script:txtActHolds.SelectedItem + "*"
            write-host $FindHold
            
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

                    write-host "No of users on Hold: " $Script:txtENos.Items.count
                    write-host "No of DBOwners on Hold: "$Script:txtDBOwner.Items.count
                    $Script:txtNoENos.Text = $Script:txtENos.Items.count
                    $Script:txtNoDBOwner.Text = $Script:txtDBOwner.Items.count
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
        $Script:lblHoldName.Top = 100; $Script:lblHoldName.Left = 10; $Script:lblHoldName.Width=150 ;$Script:lblHoldName.AutoSize = $true
        $Script:lblHoldName.Visible = $False 
        $form.Controls.Add($Script:lblHoldName)    # Add to Form 
        #
        $Script:txtHoldName = New-Object Windows.Forms.TextBox
        $Script:txtHoldName.Top = 100; $Script:txtHoldName.Left = $TextLeft; $Script:txtHoldName.Width = 340;
        $Script:txtHoldName.Text = ""
        $Script:txtHoldName.Visible = $False
        $form.Controls.Add($Script:txtHoldName)    # Add to Form

    ## Employee Numbers
    $Script:lblENos  = New-Object System.Windows.Forms.Label   
        $Script:lblENos.Text = "Staff to Hold:"  
        $Script:lblENos.Top = 130 ; $Script:lblENos.Left = 10; $Script:lblENos.Width=150 ;$Script:lblENos.AutoSize = $true
        $Script:lblENos.Visible = $False
        $form.Controls.Add($Script:lblENos)    # Add to Form 
        # 
        $Script:txtENos = New-Object Windows.Forms.ComboBox
        $Script:txtENos.Top = 130; $Script:txtENos.Left = $TextLeft; $Script:txtENos.Width = 165;
        $Script:txtENos.Visible = $False
        $form.Controls.Add($Script:txtENos)    # Add to Form

    #No of Enabled Staff
        $Script:txtNoENos = New-Object Windows.Forms.TextBox
        $Script:txtNoENos.Top = 130; $Script:txtNoENos.Left = 300; $Script:txtNoENos.Width = 25;
        $Script:txtNoENos.Visible = $True
        $form.Controls.Add($Script:txtNoENos)    # Add to Form


    ## Release Employee #'s
    $Script:lblRelENos  = New-Object System.Windows.Forms.Label   
        $Script:lblRelENos.Text = ":Staff to Add/Release"  
        $Script:lblRelENos.Top = 130 ; $Script:lblRelENos.Left = 505; $Script:lblRelENos.Width=150 ;$Script:lblRelENos.AutoSize = $true
        $Script:lblRelENos.Visible = $False
        $form.Controls.Add($Script:lblRelENos)    # Add to Form 
        # 
        $Script:txtRelENos = New-Object Windows.Forms.TextBox
        $Script:txtRelENos.Top = 130; $Script:txtRelENos.Left = 330; $Script:txtRelENos.Width = 165;
        $Script:txtRelENos.Text = "Enter comma separated ENo's"
        $Script:txtRelENos.Visible = $False
        $form.Controls.Add($Script:txtRelENos)    # Add to Form
        $Script:txtRelENos.Add_Click({
            If ($Script:txtRelENos.Text -eq "Enter comma separated ENo's")
            {
                $Script:txtRelENos.Text = ""
            }
        })

    #No of DBOwners
        $Script:txtNoDBOwner = New-Object Windows.Forms.TextBox
        $Script:txtNoDBOwner.Top = 160; $Script:txtNoDBOwner.Left = 300; $Script:txtNoDBOwner.Width = 25;
        $Script:txtNoDBOwner.Visible = $True
        $form.Controls.Add($Script:txtNoDBOwner)    # Add to Form

    ## DBOwners
    $Script:lblDBOwner  = New-Object System.Windows.Forms.Label   
        $Script:lblDBOwner.Text = "DBOwner(s):"  
        $Script:lblDBOwner.Top = 160 ; $Script:lblDBOwner.Left = 10; $Script:lblDBOwner.Width=150 ;$Script:lblDBOwner.AutoSize = $true
        $Script:lblDBOwner.Visible = $False
        $form.Controls.Add($Script:lblDBOwner)    # Add to Form 
        # 
        $Script:txtDBOwner = New-Object Windows.Forms.ComboBox
        $Script:txtDBOwner.Top = 160; $Script:txtDBOwner.Left = $TextLeft; $Script:txtDBOwner.Width = 165;
        $Script:txtDBOwner.Visible = $False
        $form.Controls.Add($Script:txtDBOwner)    # Add to Form

    ## Release DBOwners
    $Script:lblRelDBOwner  = New-Object System.Windows.Forms.Label   
        $Script:lblRelDBOwner.Text = ":DBOwners to Add/Release"  
        $Script:lblRelDBOwner.Top = 160 ; $Script:lblRelDBOwner.Left = 505; $Script:lblRelDBOwner.Width=150
        $Script:lblRelDBOwner.Visible = $False
        $form.Controls.Add($Script:lblRelDBOwner)    # Add to Form 
        # 
        $Script:txtRelDBOwner = New-Object Windows.Forms.TextBox
        $Script:txtRelDBOwner.Top = 160; $Script:txtRelDBOwner.Left = 330; $Script:txtRelDBOwner.Width = 165;
        $Script:txtRelDBOwner.Text = "Enter comma separated ENo's"
        $Script:txtRelDBOwner.Visible = $False
        $form.Controls.Add($Script:txtRelDBOwner)    # Add to Form
        $Script:txtRelDBOwner.Add_Click({
            If ($Script:txtRelDBOwner.Text -eq "Enter comma separated ENo's")
            {
                $Script:txtRelDBOwner.Text = ""
            }
        })

    #Add HR Rep
    $Script:lblHRRep = New-Object System.Windows.Forms.Label
        $Script:lblHRRep.Text = ":HR Rep"
        $Script:lblHRRep.Top = 190 ; $Script:lblHRRep.Left = 480; $Script:lblHRRep.Width=10 ;$Script:lblHRRep.AutoSize = $true
        $Script:lblHRRep.Visible = $False
        $form.Controls.Add($Script:lblHRRep)    # Add to Form 
        # 
        $Script:txtHRRep = New-Object Windows.Forms.ComboBox  
        $Script:txtHRRep.Top = 190; $Script:txtHRRep.Left = 305; $Script:txtHRRep.Width = 165;
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
        $Script:lblAttorney.Top = 190 ; $Script:lblAttorney.Left = 10; $Script:lblAttorney.Width=150 ;$Script:lblAttorney.AutoSize = $true
        $Script:lblAttorney.Visible = $False
        $form.Controls.Add($Script:lblAttorney)    # Add to Form 
        # 
        $Script:txtAttorney = New-Object Windows.Forms.ComboBox
        $Script:txtAttorney.Top = 190; $Script:txtAttorney.Left = $TextLeft; $Script:txtAttorney.Width = 165;
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

    ## Silent Hold
    $Script:chkSilent = New-Object System.Windows.Forms.Checkbox
        $Script:chkSilent.Text = "Silent Hold" 
        $Script:chkSilent.Top = 220 ; $Script:chkSilent.Left = $TextLeft; $Script:chkSilent.Width=150 ;$Script:chkSilent.AutoSize = $true 
        $Script:chkSilent.Checked = $False
        $Script:chkSilent.Visible = $False
        $form.Controls.Add($Script:chkSilent)    # Add to Form

    #Add Release Date
    $Script:lblEMailDate = New-Object System.Windows.Forms.Label
        $Script:lblEMailDate.Text = "EMail Date Releasing:"
        $Script:lblEMailDate.Top = 220; $Script:lblEMailDate.Left = 10; $Script:lblEMailDate.Width=150 ;$Script:lblEMailDate.AutoSize = $true
        $Script:lblEMailDate.Visible = $False
        $form.Controls.Add($Script:lblEMailDate)    # Add to Form
        #
        $Script:txtEMailDate = New-Object Windows.Forms.DateTimePicker
        $Script:txtEMailDate.Top = 220; $Script:txtEMailDate.Left = $TextLeft; $Script:txtEMailDate.Width = 120;
        $Script:txtEMailDate.Format = [windows.forms.datetimepickerFormat]::custom
        $Script:txtEMailDate.Visible = $False
        $Script:txtEMailDate.CustomFormat = "MM/dd/yyyy"
        $Script:txtEMailDate.Text = (get-date)
        $form.Controls.Add($Script:txtEMailDate)    # Add to Form

#Add ListBox of DBOwners in a particular hold
#Add a Count of users in a particular hold
#Add a count of DBOwners in a particular hold
#When release for multiple uses can it list just the holds for those individuals?
#Release from Standard hold Or Release from 40/180 Day Hold


        Add-FormStandardButtons
        $Global:okButton.Text = "Proceed"
}

Function Uncheck-Options
{
#    $Script:chkforUser.Checked = $False
    $Script:chkforEntireStd.Checked = $False
    $Script:chkforEntire40180.Checked = $False
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
    }
}

function Hide-All
{
    $Script:ButGetENo.visible = $False
#    $Script:chkforUser.Visible = $False
    $Script:chkforEntireStd.Visible = $False
    $Script:chkforEntire40180.Visible = $False
    $Script:chk40Day.Visible = $False
    $Script:chk180Day.Visible = $False
    $Script:chkStd.Visible = $False
    $Script:chkSilent.Visible = $False

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


$HoldDetails = import-csv "e:\O365AdminShared\Data\Legal-AllHoldDetails.csv"
$AllHold = import-csv ("\\usnbkutil100p\d$\ReportHistory\LegalHoldHistory\2023\LegalHoldReport-" + (get-date).Tostring("yyyy-MMdd") + ".csv")
Build-HoldMenu
$Global:okButton.Visible = $False
publish-form
