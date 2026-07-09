<####  This script create new UL Employee Only and UL Staff Dynamic Distribution Lists
#
# Called by:  DistributionSecurityGroupMenu.ps1
#
# 06/17/2016 - SAG - Added code to check for Displaynames longer than 64 chara and truncate them
#                  - Removed the MailUsers configuraiton as these individuals are maiFl enabled but not licensed
#				   - Added code to allow for creation of DST.SUP groups
# 09/28/2016 - SAG - Removed extra quotes around the executive name in the DST.SUP group code
# 08/31/2017 - SAG - Added code to allow the configuring the Supervisor2/Supervisor3 attributes for the DST.SUP groups 
# 02/09/2018 - SAG - Added code to include MailTips
# 04/17/2021 - SAG - Modified creation for the new Alpha Org Supervisor3 now Supervisor2 and Supervisor2 now Supervisor1 also removed configuring the MailContacts on these groups
# 05/27/2021 - SAG - Modified to inlcude creation of location and supervisor people leader groups
# 11/19/2021 - SAG - Modified to add that senderauthentication is required when creating new groups
# 01/04/2022 - SAG - Modified to include creation of AO2 groups
# 08/11/2022 - SAG - Added configuring restriction to the use of the SUP group
# 08/24/2022 - SAG - Added code to allow for Alpha_Org2 or Alpha_Org4
# 01/05/2023 - SAG - Added code o allow for Supervisor3 groups
# 08/26/2024 - SAG - Added new menu item to create AO2 or AO4 People Leader Groups
#>

Function Build-DynGroupsMenu
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "New DST/DSG Groups Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 480 ; $form.Height = 400  # Make the form wider 
    
    Add-FormStandardButtons
    $Global:form.AcceptButton.visible = $false
    
    $TopLoc = 10 
    $Script:lblTitleLine = New-Object System.Windows.Forms.Label   
        $lblTitleLine.Text = "Select Option:"
        $lblTitleLine.Top = 15 ; $lblTitleLine.Left = 60; $lblTitleLine.Width=120 ;$lblTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine)    # Add to Form 

    ## DST.All Location Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSTAllGrp = New-Object Windows.Forms.RadioButton 
        $Script:chkDSTAllGrp.Left = 100; $Script:chkDSTAllGrp.Width = 450; $Script:chkDSTAllGrp.Top = $TopLoc  
        $Script:chkDSTAllGrp.Text = "Create DST.All Groups" 
        $Script:chkDSTAllGrp.Checked = $false   # set a default value 
        $Script:chkDSTAllGrp.TabIndex = 1
        $Global:form.Controls.Add($Script:chkDSTAllGrp)
        $Script:chkDSTAllGrp.Add_Click({
            $Script:lblGrpLocName.Text = "Enter DST.All Location Name:"
            $Global:form.Height = 380
            DSTAllDetails
        })

    ## DST.All Alpha Org Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSTAllOrgGrp = New-Object Windows.Forms.RadioButton 
        $Script:chkDSTAllOrgGrp.Left = 100; $Script:chkDSTAllOrgGrp.Width = 450; $Script:chkDSTAllOrgGrp.Top = $TopLoc  
        $Script:chkDSTAllOrgGrp.Text = "Create DST.AO2 or AO4 Alpha_Org Groups" 
        $Script:chkDSTAllOrgGrp.Checked = $false   # set a default value 
        $Script:chkDSTAllOrgGrp.TabIndex = 1
        $Global:form.Controls.Add($Script:chkDSTAllOrgGrp)
        $Script:chkDSTAllOrgGrp.Add_Click({
            $Script:lblGrpLocName.Text = "Enter DST.AO2 or AO4 Alpha_Org Name:"
            DSTAllDetails
        })
        
    ## DST.All People Leader Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSTAllPPL = New-Object Windows.Forms.RadioButton 
        $Script:chkDSTAllPPL.Left = 100; $Script:chkDSTAllPPL.Width = 450; $Script:chkDSTAllPPL.Top = $TopLoc  
        $Script:chkDSTAllPPL.Text = "Create DST.All People Leader Groups" 
        $Script:chkDSTAllPPL.Checked = $false   # set a default value 
        $Script:chkDSTAllPPL.TabIndex = 2
        $Global:form.Controls.Add($Script:chkDSTAllPPL)
        $Script:chkDSTAllPPL.Add_Click({
            $Script:lblGrpLocName.Text = "Enter DST.All People Leader Location Name:"
            DSTAllDetails
        })

    ## DST.SUP Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSTSupGrp = New-Object Windows.Forms.RadioButton 
        $Script:chkDSTSupGrp.Left = 100; $Script:chkDSTSupGrp.Width = 450; $Script:chkDSTSupGrp.Top = $TopLoc  
        $Script:chkDSTSupGrp.Text = "Create DST.SUP Groups" 
        $Script:chkDSTSupGrp.Checked = $false   # set a default value 
        $Script:chkDSTSupGrp.TabIndex = 4
        $Global:form.Controls.Add($Script:chkDSTSupGrp)
        $Script:chkDSTSupGrp.Add_Click({
            DSTSUPDetails
            UpdateLocation
            $Script:ButBldGroups.visible = $true
        })

    ## DST.SUP People Leader Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSTSupPPL = New-Object Windows.Forms.RadioButton 
        $Script:chkDSTSupPPL.Left = 100; $Script:chkDSTSupPPL.Width = 450; $Script:chkDSTSupPPL.Top = $TopLoc  
        $Script:chkDSTSupPPL.Text = "Create DST.SUP People Leader Groups" 
        $Script:chkDSTSupPPL.Checked = $false   # set a default value 
        $Script:chkDSTSupPPL.TabIndex = 5
        $Global:form.Controls.Add($Script:chkDSTSupPPL)
        $Script:chkDSTSupPPL.Add_Click({
            DSTSUPDetails
            UpdateLocation
        })
 
    ## DSG.AO2 Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSGAO2 = New-Object Windows.Forms.RadioButton 
        $Script:chkDSGAO2.Left = 100; $Script:chkDSGAO2.Width = 450; $Script:chkDSGAO2.Top = $TopLoc  
        $Script:chkDSGAO2.Text = "Create DSG.AO2 Groups" 
        $Script:chkDSGAO2.Checked = $Script:chkDSGAO2.Checked   # set a default value 
        $Script:chkDSGAO2.TabIndex = 6
        $Global:form.Controls.Add($Script:chkDSGAO2)
        $Script:chkDSGAO2.Add_Click({
            $Script:lblGrpLocName.Text = "Enter DSG.AO2 Alpha Org Level 2 Name:"
            DSTAllDetails
        })         
               
    ## DSG.AO4 Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSGAO4 = New-Object Windows.Forms.RadioButton 
        $Script:chkDSGAO4.Left = 100; $Script:chkDSGAO4.Width = 450; $Script:chkDSGAO4.Top = $TopLoc  
        $Script:chkDSGAO4.Text = "Create DSG.AO4 Groups" 
        $Script:chkDSGAO4.Checked = $false   # set a default value 
        $Script:chkDSGAO4.TabIndex = 7
        $Global:form.Controls.Add($Script:chkDSGAO4)
        $Script:chkDSGAO4.Add_Click({
            $Script:lblGrpLocName.Text = "Enter DSG.AO4 Alpha Org Level 4 Name:"
            DSTAllDetails
        })

    ## DSG.All People Leader Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSGAOPPL = New-Object Windows.Forms.RadioButton 
        $Script:chkDSGAOPPL.Left = 100; $Script:chkDSGAOPPL.Width = 450; $Script:chkDSGAOPPL.Top = $TopLoc  
        $Script:chkDSGAOPPL.Text = "Create DSG.AO2 or AO4 People Leader Groups" 
        $Script:chkDSGAOPPL.Checked = $false   # set a default value 
        $Script:chkDSGAOPPL.TabIndex = 8
        $Global:form.Controls.Add($Script:chkDSGAOPPL)
        $Script:chkDSGAOPPL.Add_Click({
            $Script:lblGrpLocName.Text = "Enter DSG.AO2 or AO4 Alpha_Org Name:"
            DSTAllDetails
        })        
 
    $TopLoc = $TopLoc + 15
        $Script:lblSeperatorLine = New-Object System.Windows.Forms.Label   
        $lblSeperatorLine.Text = "__________________________________________________________________________"
        $lblSeperatorLine.Top = $TopLoc ; $lblSeperatorLine.Left = 1; $lblSeperatorLine.Width=200 ;$lblSeperatorLine.AutoSize = $true
        $lblSeperatorLine.Visible = $false 
        $Global:form.Controls.Add($lblSeperatorLine)    # Add to Form 

    $TopLoc = $TopLoc + 20    ## Details to Add Location Information
    $Script:lblGrpLocName = New-Object System.Windows.Forms.Label   
        $Script:lblGrpLocName.Top = $TopLoc ; $Script:lblGrpLocName.Left = 60; $Script:lblGrpLocName.Width=150 ;$Script:lblGrpLocName.AutoSize = $true
        $Script:lblGrpLocName.visible = $false
        $Global:form.Controls.Add($Script:lblGrpLocName)    # Add to Form 
        # 
        $Script:txtGrpLocName = New-Object Windows.Forms.TextBox  
        $Script:txtGrpLocName.Top = $TopLoc + 20; $Script:txtGrpLocName.Left = 60; $Script:txtGrpLocName.Width = 300;
        $Script:txtGrpLocName.Text = ""
        $Script:txtGrpLocName.visible = $false
        $Global:form.Controls.Add($Script:txtGrpLocName)    # Add to Form 
        $Script:txtGrpLocName.Add_Click({
            $Script:ButBldGroups.visible = $true
            UpdateLocation
        })

    ## Enter Employee Number
    $Script:lblGrpEmpNo = New-Object System.Windows.Forms.Label   
        $Script:lblGrpEmpNo.Top = $TopLoc ; $Script:lblGrpEmpNo.Left = 100; $Script:lblGrpEmpNo.Width=80
        $Script:lblGrpEmpNo.Text = "Exec Emp No:"
        $Script:lblGrpEmpNo.visible = $false
        $Global:form.Controls.Add($Script:lblGrpEmpNo)    # Add to Form 
        # 
        $Script:txtGrpEmpNo = New-Object Windows.Forms.TextBox  
        $Script:txtGrpEmpNo.Top = $TopLoc; $Script:txtGrpEmpNo.Left = 190; $Script:txtGrpEmpNo.Width = 80;
        $Script:txtGrpEmpNo.Text = ""
        $Script:txtGrpEmpNo.visible = $false
        $Global:form.Controls.Add($Script:txtGrpEmpNo)    # Add to Form
        $Script:txtGrpEmpNo.Add_Click({
            UpdateLocation
        })

    $TopLoc = $TopLoc + 30
    ## Details to Add First Name
    $Script:lblGrpFrstName = New-Object System.Windows.Forms.Label   
        $Script:lblGrpFrstName.Top = $TopLoc ; $Script:lblGrpFrstName.Left = 30; $Script:lblGrpFrstName.Width=70
        $Script:lblGrpFrstName.Text = "First Name:" 
        $Script:lblGrpFrstName.visible = $false
        $Global:form.Controls.Add($Script:lblGrpFrstName)    # Add to Form 
        # 
        $Script:txtGrpFrstName = New-Object Windows.Forms.TextBox  
        $Script:txtGrpFrstName.Top = $TopLoc; $Script:txtGrpFrstName.Left = 100; $Script:txtGrpFrstName.Width = 80;
        $Script:txtGrpFrstName.ReadOnly = $true
        $Script:txtGrpFrstName.Text = ""
        $Script:txtGrpFrstName.visible = $false
        $Global:form.Controls.Add($Script:txtGrpFrstName)    # Add to Form

    ## Details to Add Last Name
    $Script:lblGrpLastName = New-Object System.Windows.Forms.Label   
        $Script:lblGrpLastName.Top = $TopLoc ; $Script:lblGrpLastName.Left = 190; $Script:lblGrpLastName.Width=70
        $Script:lblGrpLastName.Text = "Last Name:" 
        $Script:lblGrpLastName.visible = $false
        $Global:form.Controls.Add($Script:lblGrpLastName)    # Add to Form 
        # 
        $Script:txtGrpLastName = New-Object Windows.Forms.TextBox  
        $Script:txtGrpLastName.Top = $TopLoc; $Script:txtGrpLastName.Left = 260; $Script:txtGrpLastName.Width = 80;
        $Script:txtGrpLastName.ReadOnly = $true
        $Script:txtGrpLastName.Text = ""
        $Script:txtGrpLastName.visible = $false
        $Global:form.Controls.Add($Script:txtGrpLastName)    # Add to Form

    $TopLoc = $TopLoc + 30
    ## List of Authorized Users
    $Script:lblAuthUsers = New-Object System.Windows.Forms.Label   
        $Script:lblAuthUsers.Top = $TopLoc ; $Script:lblAuthUsers.Left = 10; $Script:lblAuthUsers.Width=110
        $Script:lblAuthUsers.Text = "Auth User/Mailbox:"
        $Script:lblAuthUsers.visible = $false
        $Global:form.Controls.Add($Script:lblAuthUsers)    # Add to Form 
        # 
        $Script:txtAuthUsers = New-Object Windows.Forms.TextBox  
        $Script:txtAuthUsers.Top = $TopLoc; $Script:txtAuthUsers.Left = 120; $Script:txtAuthUsers.Width = 300;
        $Script:txtAuthUsers.Text = "Enter Employee Number or email addresses separated by commas"
        $Script:txtAuthUsers.BackColor = "White"
        $Script:txtAuthUsers.ForeColor = "Red"
        $Script:txtAuthUsers.visible = $false
        $Global:form.Controls.Add($Script:txtAuthUsers)    # Add to Form
        $Script:txtAuthUsers.Add_Click({
            If (($Script:txtAuthUsers.Text.substring(0,5)) -eq "Enter")
            {
                $Script:txtAuthUsers.Text = $Script:txtGrpEmpNo.Text + ","
                $Script:txtAuthUsers.ForeColor = "Black"
            }
        })            

    $GrpTop = 300
    ## Details to UL.COM Employees Only Group
    $Script:lblULCOMEmpOnlyGrp = New-Object System.Windows.Forms.Label   
        $Script:lblULCOMEmpOnlyGrp.Top = $GrpTop; $Script:lblULCOMEmpOnlyGrp.Left = 10; $Script:lblULCOMEmpOnlyGrp.Width= 110
        $Script:lblULCOMEmpOnlyGrp.Text = "ULCOM Emp Only:"
        $Script:lblULCOMEmpOnlyGrp.visible = $false
        $Global:form.Controls.Add($Script:lblULCOMEmpOnlyGrp)    # Add to Form 
        # 
        $Script:txtULCOMEmpOnlyGrp = New-Object Windows.Forms.TextBox  
        $Script:txtULCOMEmpOnlyGrp.Top = $GrpTop; $Script:txtULCOMEmpOnlyGrp.Left = 120; $Script:txtULCOMEmpOnlyGrp.Width = 300;
        $Script:txtULCOMEmpOnlyGrp.visible = $false
        $Global:form.Controls.Add($Script:txtULCOMEmpOnlyGrp)    # Add to Form

    $GrpTop = $GrpTop + 30
    ## Details to UL.COM Staff Group
    $Script:lblULCOMStaffGrp = New-Object System.Windows.Forms.Label   
        $Script:lblULCOMStaffGrp.Top = $GrpTop ; $Script:lblULCOMStaffGrp.Left = 10; $Script:lblULCOMStaffGrp.Width=110
        $Script:lblULCOMStaffGrp.Text = "ULCOM All Staff:"
        $Script:lblULCOMStaffGrp.visible = $false
        $Global:form.Controls.Add($Script:lblULCOMStaffGrp)    # Add to Form 
        # 
        $Script:txtULCOMStaffGrp = New-Object Windows.Forms.TextBox  
        $Script:txtULCOMStaffGrp.Top = $GrpTop; $Script:txtULCOMStaffGrp.Left = 120; $Script:txtULCOMStaffGrp.Width = 300;
        $Script:txtULCOMStaffGrp.visible = $false
        $Global:form.Controls.Add($Script:txtULCOMStaffGrp)    # Add to Form 

    $Script:ButBldGroups = New-Object Windows.Forms.Button
        $Script:ButBldGroups.Location = New-object System.Drawing.Size(230,250)
        $Script:ButBldGroups.Size = new-Object System.Drawing.Size(130,20)
        $Script:ButBldGroups.Text = "Verify Group Names"
        $Script:ButBldGroups.visible = $false
        $Global:form.Controls.Add($Script:ButBldGroups)
        $Script:ButBldGroups.Add_Click({
            $Script:ULCOMEmpGrpName = "N/A"
            $Script:ULCOMStaffGrpName = "N/A"
            ValidateLocation
            If ($Script:SupFound -ne "N")
            {
                $Global:form.Height = 450
                $Global:form.AcceptButton.visible = $false
                $Script:lblAuthUsers.visible = $true
                $Script:txtAuthUsers.visible = $true
                $Global:form.AcceptButton.visible = $true
                $Script:ButBldGroups.visible = $false
                $Script:lblULCOMEmpOnlyGrp.visible = $true
                $Script:txtULCOMEmpOnlyGrp.visible = $true
                If (($Script:txtULCOMStaffGrp.Text -ne "N/A") -or ($Script:AO2Group -eq "Yes") -or ($Script:AO4Group -eq "Yes"))
                {
                    $Script:lblULCOMStaffGrp.visible = $true
                    $Script:txtULCOMStaffGrp.visible = $true
                }
            }
            else
            {
                    $Global:form.Height = 450
                    $Script:ButBldGroups.visible = $false
                    $Script:lblULCOMEmpOnlyGrp.visible = $true
                    $Script:txtULCOMEmpOnlyGrp.visible = $true
                If ($Script:txtULCOMStaffGrp.Text -ne "N/A")
                {
                    $Script:lblULCOMStaffGrp.visible = $true
                    $Script:txtULCOMStaffGrp.visible = $true
                }
            }
        })
}

Function UpdateLocation
{
    $Script:txtGrpFrstName.Text = ""
    $Script:txtGrpLastName.Text = ""
    $Script:lblULCOMEmpOnlyGrp.visible = $false
    $Script:txtULCOMEmpOnlyGrp.visible = $false
    $Script:lblULCOMStaffGrp.visible = $false
    $Script:txtULCOMStaffGrp.visible = $false
    $Script:lblAuthUsers.visible = $false
    $Script:txtAuthUsers.visible = $false
    $Script:txtAuthUsers.Text = "Enter Employee Number or email addresses separated by commas"
    $Script:txtAuthUsers.BackColor = "White"
    $Script:txtAuthUsers.ForeColor = "Red"
    If (($Script:chkDSTSupGrp.Checked -eq $True) -or ($Script:chkDSTSupPPL.Checked -eq $True))
    {
        $Script:ButBldGroups.Visible = $true
    }
}

Function DSTSUPDetails
{
    $Script:SupFound = "N"
    $Script:ButBldGroups.Location = New-object System.Drawing.Size(130,270)
#    $Global:form.Height = 380
    $lblSeperatorLine.Visible = $true 
    $Script:lblGrpEmpNo.visible = $true
    $Script:txtGrpEmpNo.visible = $true
    $Script:txtGrpEmpNo.Focus()
    $Script:lblGrpFrstName.visible = $true
    $Script:txtGrpFrstName.Visible = $true
    $Script:lblGrpLastName.visible = $true
    $Script:txtGrpLastName.visible = $true
    $Script:lblGrpLocName.visible = $false
    $Script:txtGrpLocName.visible = $false
    $Script:lblULCOMEmpOnlyGrp.visible = $false
    $Script:txtULCOMEmpOnlyGrp.visible = $false
    $Script:lblULCOMStaffGrp.visible = $false
    $Script:txtULCOMStaffGrp.visible = $false
    $Script:lblAuthUsers.visible = $false
    $Script:txtAuthUsers.visible = $false
    $Script:ButBldGroups.visible = $false
    $Script:SupVal = "Y"
}

Function DSTAllDetails
{
    $Global:form.Height = 380
    $Script:ButBldGroups.Location = New-object System.Drawing.Size(130,250)
    $lblSeperatorLine.Visible = $true
    $Script:lblGrpLocName.visible = $true
    $Script:txtGrpLocName.visible = $true
    $Script:txtGrpLocName.Focus()
    $Script:lblGrpEmpNo.visible = $false
    $Script:txtGrpEmpNo.visible = $false
    $Script:lblGrpFrstName.visible = $false
    $Script:txtGrpFrstName.visible = $false
    $Script:lblGrpLastName.visible = $false
    $Script:txtGrpLastName.visible = $false
    $Script:lblAuthUsers.visible = $false
    $Script:txtAuthUsers.visible = $false
    $Script:lblULCOMEmpOnlyGrp.visible = $false
    $Script:txtULCOMEmpOnlyGrp.visible = $false
    $Script:lblULCOMStaffGrp.visible = $false
    $Script:txtULCOMStaffGrp.visible = $false
    $Script:ButBldGroups.visible = $true
    $Script:SupVal = "N"
    $Global:InputFocus = $Global:cancelButton
}

Function SUPDetails
{
    $Exists = [bool](get-mailbox $Script:txtGrpEmpNo.Text -ErrorAction SilentlyContinue)

    If ($Exists -eq $True)
    {
        $SUP = get-ADUser $Script:txtGrpEmpNo.Text
        $Script:txtGrpFrstName.Text = $SUP.GivenName
        $Script:txtGrpLastName.Text = $SUP.SurName
        $Script:txtGrpFrstName.Refresh()
        $Script:txtGrpLastName.Refresh()
        $Script:SupFound = "1"

        $Output = $wshell.Popup("Checking if this individual is a Supervisor1, Supervisor2 or a Supervisor3.  Be patient while data is extracted from the OED_Extract file.",5,"Validate Supervisor",0+32)
        write-host "Checking to see if this individual is a Supervisor1 in Oracle" -ForegroundColor Cyan
        $SupExists = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv | Where-Object{(($_."Supervisor_1_NUMBER" -eq $Script:txtGrpEmpNo.Text) -and ($_."ASSIGNMENT STATUS" -ne "T"))}
        If ($SupExists.count -eq 0)
        {
            $Script:SupFound = "2"
            write-host "Checking to see if this individual is a Supervisor2 in Oracle" -ForegroundColor Cyan
            $SupExists = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv | Where-Object{(($_."Supervisor_2_NUMBER" -eq $Script:txtGrpEmpNo.Text) -and ($_."ASSIGNMENT STATUS" -ne "T"))}
#            If ($SupExists[0].count -eq 0)
            If ($SupExists.Length -eq 0)
            {
                $Script:SupFound = "3"
                write-host "Checking to see if this individual is a Supervisor3 in Oracle" -ForegroundColor Cyan
                $SupExists = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv | Where-Object{(($_."Supervisor_3_NUMBER" -eq $Script:txtGrpEmpNo.Text) -and ($_."ASSIGNMENT STATUS" -ne "T"))}
                If ($SupExists.count -eq 0)
                {
                    $Script:SupFound = "N"
                    $Output = $wshell.Popup("Employee Number " + $Script:txtGrpEmpNo.Text + " is not listed as a Supervisor1, Supervisor2 or Supervisor3.",0,"Not Sup1/Sup2/Sup3",0+32)
                    $Global:form.AcceptButton.visible = $false  
                }
                else
                {
                    $Global:form.AcceptButton.visible = $true
                }
            }
        }
    }
    else
    {
        $Output = $wshell.Popup("This individual " + $Script:txtGrpEmpNo.Text + " does not have an active mailbox.",5,"Validate Supervisor",0+32)
        $Script:SupFound = "N"
        $Script:txtGrpEmpNo.Text = ""
        $Global:form.AcceptButton.visible = $false
    }
}

Function ValidateLocation
{
    $Script:SupFound = "N"
    $Script:ActiveULCOMDSTRecs = ""
    $Script:AO2Group = "No"
    $Script:AO4Group = "No"

    If (($Script:chkDSTAllGrp.Checked -eq $true) -or ($Script:chkDSTAllPPL.Checked -eq $True))
    {
        $Output = $wshell.Popup("Gathering all sites with Active UL.COM staff for this location.  Be patient while data is extracted from the OED_Extract file.",10,"Locations with Active Staff",0+32)
        $Script:ActiveULCOMDSTRecs = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv | Where-Object{((($_.ALPHA_HR_ORG_LEVEL_1 -eq "UL Inc.") -and ($_.Location -eq $Script:txtGrpLocName.Text)) -and (($_."Assignment Status" -eq "A") -or ($_."Assignment Status" -eq "F")))}
    }

    If (($Script:chkDSTSupGrp.Checked -eq $True) -or ($Script:chkDSTSupPPL.Checked -eq $True))
    {
            SUPDetails
    }

    If (($Script:chkDSTAllOrgGrp.Checked -eq $True) -or ($Script:chkDSGAOPPL.Checked -eq $True))
    {
        $Output = $wshell.Popup("Verifying there are Active Staff for the " + $Script:txtGrpLocName.Text+ " AO2 or AO4 organization.  Be patient while data is extracted from the OED_Extract file.",10,"Locations with Active Staff",0+32)
        $Script:ActiveULCOMDSTRecs = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv | Where-Object{(($_.ALPHA_HR_ORG_LEVEL_2 -eq $Script:txtGrpLocName.Text) -and (($_."Assignment Status" -eq "A") -or ($_."Assignment Status" -eq "F")))}
        If ($Script:ActiveULCOMDSTRecs.count -ne 0)
        {
            $Script:AO2Group = "Yes"  
        }
        else
        {
            $Script:ActiveULCOMDSTRecs = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv | Where-Object{(($_.ALPHA_HR_ORG_LEVEL_4 -eq $Script:txtGrpLocName.Text) -and (($_."Assignment Status" -eq "A") -or ($_."Assignment Status" -eq "F")))} 
            If ($Script:ActiveULCOMDSTRecs.count -ne 0)
            {
                $Script:AO4Group = "Yes"  
            }
        }
    }

    If ($Script:chkDSGAO2.Checked -eq $true)
    {
        $Output = $wshell.Popup("Gathering all AO2 Locations be patient while data is extracted from the OED_Extract file.",5,"Locations with Active Staff",0+32)
        $Script:ActiveULCOMDSTRecs = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv | Where-Object{(($_."Assignment Status" -eq "A") -or ($_."Assignment Status" -eq "F"))} | select -unique ALPHA_HR_ORG_LEVEL_2 | Sort-Object ALPHA_HR_ORG_LEVEL_4
    }

    If ($Script:chkDSGAO4.Checked -eq $true)
    {
        $Output = $wshell.Popup("Gathering all AO4 Locations be patient while data is extracted from the OED_Extract file.",5,"Locations with Active Staff",0+32)
        $Script:ActiveULCOMDSTRecs = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv | Where-Object{(($_."Assignment Status" -eq "A") -or ($_."Assignment Status" -eq "F"))} | select -unique ALPHA_HR_ORG_LEVEL_4 | Sort-Object ALPHA_HR_ORG_LEVEL_4
    }

    If ($Script:ActiveULCOMDSTRecs.count -ne 0)
#    If (($Script:ActiveULCOMDSTRecs.count.length -ge 1)
    {
        $Script:lblULCOMEmpOnlyGrp.Text = "ULCOM Emp Only:"
        $Script:lblULCOMStaffGrp.Text = "ULCOM All Staff:"
        If ($Script:chkDSTSupGrp.Checked -eq $True)
        {
                $Script:lblULCOMEmpOnlyGrp.Text = "Emp Only Grp:"
                $Script:lblULCOMStaffGrp.Text = "All Staff Grp:"
        }
        If (($Script:chkDSGAO2.Checked -eq $True) -or ($Script:AO2Group -eq "Yes"))
        {
            $Script:lblULCOMEmpOnlyGrp.Text = "  AO2 Emp Only:"
            $Script:lblULCOMStaffGrp.Text = "  AO2 All Staff:"
        }
        If (($Script:chkDSGAO4.Checked -eq $True) -or ($Script:AO4Group -eq "Yes"))
        {
            $Script:lblULCOMEmpOnlyGrp.Text = "  AO4 Emp Only:"
            $Script:lblULCOMStaffGrp.Text = "  AO4 All Staff:"
        }
        If (($Script:chkDSTSupPPL.Checked -eq $True) -or ($Script:chkDSGAOPPL.Checked -eq $True))
        {
            $Script:lblULCOMEmpOnlyGrp.Text = "        Group Name:"
        }

        BuildGroupNames
    }
    else
    {
        $Script:SupFound = "N"
        $Script:txtGrpLocName.Text = ""
        $Output = $wshell.Popup("There are no active Staff in the identified location " + $Script:txtGrpLocName.Text,10,"No Active Staff",0+32)
    }
}

Function BuildGroupNames
{
    $Global:form.AcceptButton.visible = $true
    $Script:ULCOMEmpGrpName = "N/A"
    $Script:ULCOMStaffGrpName = "N/A"
    If ($Script:chkDSTAllGrp.Checked -eq "Checked")
    {
        If ($Script:ActiveULCOMDSTRecs.Location -eq $Script:txtGrpLocName.Text)
        {
            $Script:ULCOMEmpGrpName = "DST.All UL COM " + $Script:txtGrpLocName.Text + " Employees Only"
	        $Script:ULCOMStaffGrpName = "DST.All UL COM " + $Script:txtGrpLocName.Text + " Staff"
        }
        else
        {
            $Script:ULCOMEmpGrpName = "N/A"
            $Script:ULCOMStaffGrpName = "N/A"
            $Output = $wshell.Popup("No Active UL.COM Staff at the site:" + $Script:txtGrpLocName.Text ,5,"Invalid Location",0+32)
            $Global:form.AcceptButton.visible = $false
        }
        $Script:lblAuthUsers.visible = $true
        $Script:txtAuthUsers.visible = $true
    }

    If ($Script:chkDSTAllOrgGrp.Checked -eq "Checked")
    {
        If ($Script:AO2Group -eq "Yes")
        {
            $Script:ULCOMEmpGrpName = "DST.AO2 " + $Script:txtGrpLocName.Text + " Employees Only"
	        $Script:ULCOMStaffGrpName = "DST.AO2 " + $Script:txtGrpLocName.Text + " Staff"
        }
        else
        {
            If ($Script:AO4Group -eq "Yes")
            {
                $Script:ULCOMEmpGrpName = "DST.AO4 " + $Script:txtGrpLocName.Text + " Employees Only"
	            $Script:ULCOMStaffGrpName = "DST.AO4 " + $Script:txtGrpLocName.Text + " Staff"                
            }
        }
    }

    If ($Script:chkDSTSUPGrp.Checked -eq "Checked")
    {
        $Script:ULCOMEmpGrpName = "DST.SUP " + $Script:txtGrpLastName.Text + " " + $Script:txtGrpFrstName.Text + " Employees Only"
	    $Script:ULCOMStaffGrpName = "DST.SUP " + $Script:txtGrpLastName.Text + " " + $Script:txtGrpFrstName.Text + " Staff"
    }

    If ($Script:chkDSTAllPPL.Checked -eq "Checked")
    {
        If ($Script:ActiveULCOMDSTRecs.Location -eq $Script:txtGrpLocName.Text)
        {
            $Script:lblULCOMEmpOnlyGrp.Text = "  ULCOM Group:"
            $Script:ULCOMEmpGrpName = "DST.All UL COM " + $Script:txtGrpLocName.Text + " People Leaders"
        }
        else
        {
            $Script:ULCOMORGEmpGrpName = "N/A"
            $Output = $wshell.Popup("No Active UL.COM Staff at the site:" + $Script:txtGrpLocName.Text ,5,"Invalid Location",0+32)
            $Global:form.AcceptButton.visible = $false
        }
    }

    If ($Script:chkDSTSUPPPL.Checked -eq "Checked")
    {
        $Script:ULCOMEmpGrpName = "DST.SUP " + $Script:txtGrpLastName.Text + " " + $Script:txtGrpFrstName.Text + " People Leaders"
    }

    If ($Script:chkDSGAO2.Checked -eq "Checked")
    {
        If ($Script:ActiveULCOMDSTRecs.ALPHA_HR_ORG_LEVEL_2 -eq $Script:txtGrpLocName.Text)
        {
            $Script:ULCOMEmpGrpName = "DSG.AO2." + $Script:txtGrpLocName.Text + " Employees Only"
	        $Script:ULCOMStaffGrpName = "DSG.AO2." + $Script:txtGrpLocName.Text + " Staff"
        }
        else
        {
            $Script:ULCOMEmpGrpName = "N/A"
            $Script:ULCOMStaffGrpName = "N/A"
            $Output = $wshell.Popup("Identified AO2 Organization not found: " + $Script:txtGrpLocName.Text ,5,"Invalid Organization",0+32)
            $Global:form.AcceptButton.visible = $false
        }
    }

    If ($Script:chkDSGAO4.Checked -eq "Checked")
    {
        If ($Script:ActiveULCOMDSTRecs.ALPHA_HR_ORG_LEVEL_4 -eq $Script:txtGrpLocName.Text)
        {
            $Script:ULCOMEmpGrpName = "DSG.AO4." + $Script:txtGrpLocName.Text + " Employees Only"
	        $Script:ULCOMStaffGrpName = "DSG.AO4." + $Script:txtGrpLocName.Text + " Staff"
            $Global:form.AcceptButton.visible = $true
        }
        else
        {
            $Script:ULCOMEmpGrpName = "N/A"
            $Script:ULCOMStaffGrpName = "N/A"
            $Output = $wshell.Popup("Identified AO4 Organization not found: " + $Script:txtGrpLocName.Text ,5,"Invalid Organization",0+32)
            $Global:form.AcceptButton.visible = $false
        }
    }

    If ($Script:chkDSGAOPPL.Checked -eq "Checked")
    {
        If ($Script:AO2Group -eq "Yes")
        {
            $Script:ULCOMEmpGrpName = "DSG.AO2. " + $Script:txtGrpLocName.Text + " People Leaders"
        }
        else
        {
            If ($Script:AO4Group -eq "Yes")
            {
                $Script:ULCOMEmpGrpName = "DSG.AO4. " + $Script:txtGrpLocName.Text + " People Leaders"
            }
        }
    }

    $Script:txtULCOMEmpOnlyGrp.Text = $Script:ULCOMEmpGrpName
    $Script:txtULCOMStaffGrp.Text = $Script:ULCOMStaffGrpName
    $Script:ULCOMEmpGrpNameTrunc = $Script:ULCOMEmpGrpName
    $Script:ULCOMStaffGrpNameTrunc = $Script:ULCOMStaffGrpName

    If ($Script:ULCOMEmpGrpName.Length -gt 64)
    {
        $Script:ULCOMEmpGrpNameTrunc = ($Script:ULCOMEmpGrpName.Substring(0,64)).Trim()
    }
    
    If ($Script:ULCOMStaffGrpName.Length -gt 64)
    {
        $Script:ULCOMStaffGrpNameTrunc = ($Script:ULCOMStaffGrpName.Substring(0,64)).Trim()
    }
}

Function Create-DSTAllGroups
{
    If ($Global:Result -eq "OK")
    {
        Write-Host "Creating Groups...." -ForegroundColor Cyan
        If (($Script:chkDSTAllPPL.Checked -eq $True) -or ($Script:chkDSTSUPPPL.Checked -eq $True))
        {
            write-host "Name of the group for UL.COM People Leaders: " $Script:ULCOMEmpGrpName
        }
        else
        {
            write-host "Name of the group for UL.COM Employees Only: " $Script:ULCOMEmpGrpName
            write-host "     Name of the group for all UL.COM Staff: " $Script:ULCOMStaffGrpName
        }

        $ULCOMEmpGrpAlias = $Script:ULCOMEmpGrpName -replace("[(),&) ]","")
        $ULCOMEmpGrpAlias = $ULCOMEmpGrpAlias -replace ("-","")
        $ULCOMEmpGrpAlias = $ULCOMEmpGrpAlias.Replace("UL","")
        $ULCOMEmpGrpAlias = $ULCOMEmpGrpAlias.Replace("loyees","")
        $ULCOMEmpINetAlias = $ULCOMEmpGrpAlias + "@ul.com"

        $ULCOMStaffGrpAlias = $Script:ULCOMStaffGrpName -replace("[(),&) ]","")
        $ULCOMStaffGrpAlias = $ULCOMStaffGrpAlias -replace ("-","")
        $ULCOMStaffGrpAlias = $ULCOMStaffGrpAlias.Replace("UL","")
        $ULCOMStaffINetAlias = $ULCOMStaffGrpAlias + "@ul.com"

        $LocName = $Script:txtGrpLocName.Text
        $FrstName = $Script:txtGrpFrstName.Text
        $LastName = $Script:txtGrpLastName.Text

        If ($Script:chkDSTAllGrp.Checked -eq $True)
        {
            #Create UL.COM Groups
            If ($Script:ActiveULCOMDSTRecs.Location -eq $Script:txtGrpLocName.Text)
            {
                $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMEmpGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
                    Write-Host "Creating New Group for: " $Script:ULCOMEmpGrpName -ForegroundColor Green
                    New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCompany "UL.COM" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -PrimarySmtpAddress $ULCOMEmpINetAlias
    	            $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff in UL.COM that have an employee type of ""Employee"" and are in the ""$LocName"" location.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                    $ULCOMMailTip = "Sends all UL.COM Employees Only assigned to the '" + $LocName + "' site in Oracle."
                    Set-Mailtip -Name $Script:ULCOMEmpGrpName -Note $ULCOMEmpGrpNote -Tip $ULCOMMailTip
                    Apply-UseRestriction ($Script:ULCOMEmpGrpName)
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMEmpGrpName,10,"Group Exists",0+32)
                }

                $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMStaffGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
	                Write-Host "Creating New Group for: " $Script:ULCOMStaffGrpName -ForegroundColor Green
                    New-DynamicDistributionGroup -DisplayName $Script:ULCOMStaffGrpName -Alias $ULCOMStaffGrpAlias -Name $Script:ULCOMStaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCompany "UL.COM" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -PrimarySmtpAddress $ULCOMStaffINetAlias
                    $ULCOMStaffGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff in UL.COM that have an employee type of ""Employee, Consultant, Contractor, Temporary Worker or Freelance Worker"" and are in the ""$LocName"" location.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                    $ULCOMStaffTip = "Sends to all UL.COM Employee and Non-Employee types assigned to the '" + $LocName + "' site in Oracle."
                    Set-Mailtip -Name $Script:ULCOMStaffGrpName -Note $ULCOMStaffGrpNote -Tip $ULCOMStaffTip
                    Apply-UseRestriction ($Script:ULCOMStaffGrpName)
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMStaffGrpName,10,"Group Exists",0+32)
                }
            }
        }

        If ($Script:chkDSTAllOrgGrp.Checked -eq $True)
        {
            If ($Script:AO2Group -eq "Yes")
            {
                $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMEmpGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
                    Write-Host "Creating New Group for: " $Script:ULCOMEmpGrpName -ForegroundColor Green
                    New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute8 $LocName -PrimarySmtpAddress $ULCOMEmpINetAlias
    	            $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff in AlphaOrg2 ""$LocName"" that have an employee type of ""Employee"".  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                    $ULCOMMailTip = "Sends all UL.COM Employees Only assigned to the '" + $LocName + "' site in Oracle."
                    Set-Mailtip -Name $Script:ULCOMEmpGrpName -Note $ULCOMEmpGrpNote -Tip $ULCOMMailTip
                    Apply-UseRestriction ($Script:ULCOMEmpGrpName)
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMEmpGrpName,10,"Group Exists",0+32)
                }

                $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMStaffGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
	                Write-Host "Creating New Group for: " $Script:ULCOMStaffGrpName -ForegroundColor Green
                    New-DynamicDistributionGroup -DisplayName $Script:ULCOMStaffGrpName -Alias $ULCOMStaffGrpAlias -Name $Script:ULCOMStaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute8 $LocName -PrimarySmtpAddress $ULCOMStaffINetAlias
                    $ULCOMStaffGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff in AlphaOrg2 ""$LocName"" that have an employee type of ""Employee, Consultant, Contractor, Temporary Worker or Freelance Worker"".  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                    $ULCOMStaffTip = "Sends to all UL.COM Employee and Non-Employee types assigned to the '" + $LocName + "' site in Oracle."
                    Set-Mailtip -Name $Script:ULCOMEmpGrpName -Note $ULCOMEmpGrpNote -Tip $ULCOMMailTip
                    Apply-UseRestriction ($Script:ULCOMStaffGrpName)
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMStaffGrpName,10,"Group Exists",0+32)
                }
            }

            If ($Script:AO4Group -eq "Yes")
            {
                $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMEmpGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
                    Write-Host "Creating New Group for: " $Script:ULCOMEmpGrpName -ForegroundColor Green
                    New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute4 $LocName -PrimarySmtpAddress $ULCOMEmpINetAlias
    	            $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff in AlphaOrg4 ""$LocName"" that have an employee type of ""Employee"".  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                    $ULCOMMailTip = "Sends all UL.COM Employees Only assigned to the '" + $LocName + "' site in Oracle."
                    Set-Mailtip -Name $Script:ULCOMEmpGrpName -Note $ULCOMEmpGrpNote -Tip $ULCOMMailTip
                    Apply-UseRestriction ($Script:ULCOMEmpGrpName)
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMEmpGrpName,10,"Group Exists",0+32)
                }

                $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMStaffGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
	                Write-Host "Creating New Group for: " $Script:ULCOMStaffGrpName -ForegroundColor Green
                    New-DynamicDistributionGroup -DisplayName $Script:ULCOMStaffGrpName -Alias $ULCOMStaffGrpAlias -Name $Script:ULCOMStaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute4 $LocName -PrimarySmtpAddress $ULCOMStaffINetAlias
                    $ULCOMStaffGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff in AlphaOrg4 ""$LocName"" that have an employee type of ""Employee, Consultant, Contractor, Temporary Worker or Freelance Worker"".  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                    $ULCOMStaffTip = "Sends to all UL.COM Employee and Non-Employee types assigned to the '" + $LocName + "' site in Oracle."
                    Set-Mailtip -Name $Script:ULCOMStaffGrpName -Note $ULCOMStaffGrpNote -Tip $ULCOMStaffTip
                    Apply-UseRestriction ($Script:ULCOMStaffGrpName)
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMStaffGrpName,10,"Group Exists",0+32)
                }
            }
        }

        If (($Script:chkDSTSupGrp.Checked -eq $True) -and ($Script:SupFound -ne "N"))
        {
            switch ($Script:SupFound)
            {
                "1"
                {
                    $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMEmpGrpName -ErrorAction SilentlyContinue)
                    If ($Exists -eq $False)
                    {
                        Write-Host "Creating New Group for: " $Script:ULCOMEmpGrpName -ForegroundColor Green
                        New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute6 $Script:txtGrpEmpNo.Text -PrimarySmtpAddress $ULCOMEmpINetAlias
                        $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisor1 in Oracle.  This group contains staff in UL.COM that have an employee type of ""Employee"".  Manual modification of this group is not possible."
                        $ULCOMMailTip = "Sends all Employees that report to '" + $FrstName + " " + $LastName + "' as Supervisor1 in Oracle."
                        Set-Mailtip -Name $Script:ULCOMEmpGrpName -Note $ULCOMEmpGrpNote -Tip $ULCOMMailTip
                        Apply-UseRestriction ($Script:ULCOMEmpGrpName)
                    }
                    else
                    {
                        $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMEmpGrpName,10,"Group Exists",0+32)
                    }

                    $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMStaffGrpName -ErrorAction SilentlyContinue)
                    If ($Exists -eq $False)
                    {
	                    Write-Host "Creating New Group for: " $Script:ULCOMStaffGrpName -ForegroundColor Green
    	                New-DynamicDistributionGroup -DisplayName $Script:ULCOMStaffGrpName -Alias $ULCOMStaffGrpAlias -Name $Script:ULCOMStaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $Script:txtGrpEmpNo.Text -ConditionalCustomAttribute6 $Script:txtGrpEmpNo.Text -PrimarySmtpAddress $ULCOMStaffINetAlias
                        $ULCOMStaffGrpNote = "The membership of this group is determined at the time the group and contains individuals that report up to " + $FrstName + " " + $LastName + " as Supervisor1 in Oracle.  This group contains staff in UL.COM that have an employee type of ""Employee, Consultant, Contractor or Temporary Worker"".  Manual modification of this group is not possible."
                        $ULCOMStaffTip = "Sends to all Employee and Non-Employee types that report to '"  + $FrstName + " " + $LastName + "' as Supervisor1 in Oracle."
                        Set-Mailtip -Name $Script:ULCOMStaffGrpName -Note $ULCOMStaffGrpNote -Tip $ULCOMStaffTip
                        Apply-UseRestriction ($Script:ULCOMStaffGrpName)
                    }
                    else
                    {
                        Write-Host "Group already exists for: " $Script:ULCOMStaffGrpName -ForegroundColor Red
                    }
                }

                "2"
                {
                    $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMEmpGrpName -ErrorAction SilentlyContinue)
                    If ($Exists -eq $False)
                    {
                        Write-Host "Creating New Group for: " $Script:ULCOMEmpGrpName -ForegroundColor Green
                        New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute7 $Script:txtGrpEmpNo.Text -PrimarySmtpAddress $ULCOMEmpINetAlias
        	            $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisor2 in Oracle.  This group contains staff in UL.COM that have an employee type of ""Employee"".  Manual modification of this group is not possible."
                        $ULCOMMailTip = "Sends to individuals listed as Employee that report to '" + $FrstName + " " + $LastName + "' as Supervisor2 in Oracle."
                        Write-host "Restricting Use of Group to: " $Script:txtGrpEmpNo.Text
                        Set-Mailtip -Name $Script:ULCOMEmpGrpName -Note $ULCOMEmpGrpNote -Tip $ULCOMMailTip
                        Apply-UseRestriction ($Script:ULCOMEmpGrpName)
                    }
                    else
                    {
                        $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMEmpGrpName,10,"Group Exists",0+32)
                    }

                    $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMStaffGrpName -ErrorAction SilentlyContinue)
                    If ($Exists -eq $False)
                    {
	                    Write-Host "Creating New Group for: " $Script:ULCOMStaffGrpName -ForegroundColor Green
                        New-DynamicDistributionGroup -DisplayName $Script:ULCOMStaffGrpName -Alias $ULCOMStaffGrpAlias -Name $Script:ULCOMStaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute7 $Script:txtGrpEmpNo.Text -PrimarySmtpAddress $ULCOMStaffINetAlias
    	                $ULCOMStaffGrpNote = "The membership of this group is determined at the time the group and contains individuals that report up to " + $FrstName + " " + $LastName + " as Supervisor2 in Oracle.  This group contains staff in UL.COM that have an employee type of ""Employee, Consultant, Contractor or Temporary Worker"".  Manual modification of this group is not possible."
                        $ULCOMStaffTip = "Sends to all Employee and Non-Employee types that report to '"  + $FrstName + " " + $LastName + "' as Supervisor2 in Oracle."
                        Write-host "Restricting Use of Group to: " $Script:txtGrpEmpNo.Text
                        Set-Mailtip -Name $Script:ULCOMStaffGrpName -Note $ULCOMStaffGrpNote -Tip $ULCOMStaffTip
                        Apply-UseRestriction ($Script:ULCOMStaffGrpName)
                    }
                }

                "3"
                {
                    $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMEmpGrpName -ErrorAction SilentlyContinue)
                    If ($Exists -eq $False)
                    {
                        Write-Host "Creating New Group for: " $Script:ULCOMEmpGrpName -ForegroundColor Green
                        New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute10 $Script:txtGrpEmpNo.Text -PrimarySmtpAddress $ULCOMEmpINetAlias
    	                $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisor3 in Oracle.  This group contains staff in UL.COM that have an employee type of ""Employee"".  Manual modification of this group is not possible."
                        $ULCOMMailTip = "Sends to individuals listed as Employee that report to '" + $FrstName + " " + $LastName + "' as Supervisor3 in Oracle."
                        Write-host "Restricting Use of Group to: " $Script:txtGrpEmpNo.Text
                        Set-Mailtip -Name $Script:ULCOMEmpGrpName -Note $ULCOMEmpGrpNote -Tip $ULCOMMailTip
                        Apply-UseRestriction ($Script:ULCOMEmpGrpName)
                    }
                    else
                    {
                        $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMEmpGrpName,10,"Group Exists",0+32)
                    }

                    $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMStaffGrpName -ErrorAction SilentlyContinue)
                    If ($Exists -eq $False)
                    {
                        Write-Host "Creating New Group for: " $Script:ULCOMStaffGrpName -ForegroundColor Green
                        New-DynamicDistributionGroup -DisplayName $Script:ULCOMStaffGrpName -Alias $ULCOMStaffGrpAlias -Name $Script:ULCOMStaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute10 $Script:txtGrpEmpNo.Text -PrimarySmtpAddress $ULCOMStaffINetAlias
                        $ULCOMStaffGrpNote = "The membership of this group is determined at the time the group and contains individuals that report up to " + $FrstName + " " + $LastName + " as Supervisor3 in Oracle.  This group contains staff in UL.COM of any employee type.  Manual modification of this group is not possible."
                        $ULCOMStaffTip = "Sends to all Employee and Non-Employee types that report to '"  + $FrstName + " " + $LastName + "' as Supervisor3 in Oracle."
                        Write-host "Restricting Use of Group to: " $Script:txtGrpEmpNo.Text
                        Set-Mailtip -Name $Script:ULCOMStaffGrpName -Note $ULCOMStaffGrpNote -Tip $ULCOMStaffTip
                        Apply-UseRestriction ($Script:ULCOMStaffGrpName)
                    }
                    else
                    {
                        Write-Host "Group already exists for: " $Script:ULCOMStaffGrpName -ForegroundColor Red
                    }
                }
            }
        }

        If ($Script:chkDSTAllPPL.Checked -eq $True)
        {
            If ($Script:ActiveULCOMDSTRecs.Location -eq $Script:txtGrpLocName.Text)
            {
                $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMEmpGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
                    Write-Host "Creating New Group for: " $Script:ULCOMEmpGrpName -ForegroundColor Green
                    New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCompany "UL.COM" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $ULCOMEmpINetAlias
    	            $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and contains UL.COM individuals location in the ""$LocName"" location and is a People Leader in Oracle.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                    $ULCOMMailTip = "Sends to individuals in UL.COM listed as Employee assigned to the '" + $Script:txtGrpLocName.Text + "' site in Oracle."
                    Set-Mailtip -Name $Script:ULCOMEmpGrpName -Note $ULCOMEmpGrpNote -Tip $ULCOMMailTip
                    Apply-UseRestriction ($Script:ULCOMEmpGrpName)
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMEmpGrpName,10,"Group Exists",0+32)
                }
            }
        }

        If ($Script:chkDSGAOPPL.Checked -eq $True)
        {
            If ($Script:AO2Group -eq "Yes")
            {
                write-host "Calling AO2 PPL group" -ForegroundColor cyan
                Create-DSGAOPPLAO2Group
            }
            else
            {
                If ($Script:AO4Group -eq "Yes")
                {
                    write-host "Calling AO4 PPL group" -ForegroundColor cyan
                    Create-DSGAOPPLAO4Group
                }
            }
<#
#            If ($Script:ActiveULCOMDSTRecs.Location -eq $Script:txtGrpLocName.Text)
#            {
                $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMEmpGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
                    Write-Host "Creating New Group for: " $Script:ULCOMEmpGrpName -ForegroundColor Green
                    New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCompany "UL.COM" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $ULCOMEmpINetAlias
    	            $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and contains UL.COM individuals location in the ""$LocName"" location and is a People Leader in Oracle.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                    $ULCOMMailTip = "Sends to individuals in UL.COM listed as Employee assigned to the '" + $Script:txtGrpLocName.Text + "' site in Oracle."
                    Set-Mailtip -Name $Script:ULCOMEmpGrpName -Note $ULCOMEmpGrpNote -Tip $ULCOMMailTip
                    Apply-UseRestriction ($Script:ULCOMEmpGrpName)
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMEmpGrpName,10,"Group Exists",0+32)
                }
#            }
#>
        }

        If ($Script:chkDSTSupPPL.Checked -eq $True)
        {
            $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMEmpGrpName -ErrorAction SilentlyContinue)
            If ($Exists -eq $False)
            {
                Write-Host "Creating New Group for: " $Script:ULCOMEmpGrpName -ForegroundColor Green
                switch ($Script:SupFound)
                {
                    "1"
                    {
                        New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute6 $Script:txtGrpEmpNo.Text -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $ULCOMEmpINetAlias
                        $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisor1 in Oracle.  Manual modification of this group is not possible."
                        $ULCOMMailTip = "Sends to individuals listed as Employee that report to '" + $Script:txtGrpLastName.Text + " " + $Script:txtGrpFrstName.Text + "' as Supervisor1 in Oracle."
                        Set-Mailtip -Name $Script:ULCOMEmpGrpName -Note $ULCOMEmpGrpNote -Tip $ULCOMMailTip
                        Apply-UseRestriction ($Script:ULCOMEmpGrpName)
                    }
                
                    "2"
                    {
                        New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute7 $Script:txtGrpEmpNo.Text -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $ULCOMEmpINetAlias
        	            $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisor2 in Oracle.   Manual modification of this group is not possible."
                        $ULCOMMailTip = "Sends to individuals listed as Employee that report to '" + $Script:txtGrpLastName.Text + " " + $Script:txtGrpLastName.Text + "' as Supervisor2 in Oracle."
                        Set-Mailtip -Name $Script:ULCOMEmpGrpName -Note $ULCOMEmpGrpNote -Tip $ULCOMMailTip
                        Apply-UseRestriction ($Script:ULCOMEmpGrpName)
                    }

                    "3"
                    {
                        New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute10 $Script:txtGrpEmpNo.Text -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $ULCOMEmpINetAlias
        	            $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisor3 in Oracle.  Manual modification of this group is not possible."
                        $ULCOMMailTip = "Sends to individuals listed as Employee that report to '" + $Script:txtGrpLastName.Text + " " + $Script:txtGrpLastName.Text + "' as Supervisor3 in Oracle."
                        Set-Mailtip -Name $Script:ULCOMEmpGrpName -Note $ULCOMEmpGrpNote -Tip $ULCOMMailTip
                        Apply-UseRestriction ($Script:ULCOMEmpGrpName)
                    }
                }
                Set-DynamicDistributionGroup $Script:ULCOMEmpGrpName -Notes $ULCOMEmpGrpNote -MailTip $ULCOMMailTip
            }
            else
            {
                $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMEmpGrpName,10,"Group Exists",0+32)
            }
        }
    }
    else
    {
        $Output = $wshell.Popup("No Groups Created",10,"Not Created",0+32)
    }
}

Function Create-DSGAO4Groups
{
    $AO4Groups = Get-AzureADMSGroup -SearchString "DSG.AO4"
    $cnt = 0

    $bute4 = $Script:txtGrpLocName.Text
    $AO4 = $bute4 -replace (",","")

    $ULCOMEmpGrpAlias = $Script:ULCOMEmpGrpName -replace ("-","")
    $ULCOMEmpGrpAlias = $ULCOMEmpGrpAlias -replace ("-","")
    
    $EmpGrpNick = "DSG.AO4." + $Script:txtGrpLocName.Text + "EmpOnly" -replace("[(),&) ]","")
    $EmpGrpNick = $EmpGrpNick -replace ("-","")
    $EmpGrpDesc = "All Employees in Alpha Org 4 " +  $Script:txtGrpLocName.Text
    $Found = ""
    Foreach ($chk in $AO4Groups)
    {
        If ($chk.DisplayName -eq $Script:ULCOMEmpGrpName)
        {
            $Found = "Yes"
        }
    }
    $Script:Subject = "New DSG.AO4 Dynamic Group Created for " + $Script:txtGrpLocName.Text
    $Script:Body = "As a result of our weekly Alpha Org2 and Alpha Org4 review the following DST.AO4 group was added to the system:<br>"
    If ($Found -eq "")
    {
        Write-host "`nCreating" $Script:ULCOMEmpGrpName -ForegroundColor Green
        New-AzureADMSGroup -Description $EmpGrpDesc -DisplayName $Script:ULCOMEmpGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $EmpGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute1 -contains ""Employee"") and (user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute4 -eq ""$bute4"")" -MembershipRuleProcessingState "On"
        $Script:Body = $Script:Body + "<li>$Script:ULCOMEmpGrpName</li>"
    }
    else
    {
        Write-host "`nGroup" $Script:ULCOMEmpGrpName "already exists" -ForegroundColor Red
    }

    $StaffGrpNick = "DSG.AO4." + $Script:txtGrpLocName.Text + "Staff" -replace("[(),&) ]","")
    $StaffGrpNick = $StaffGrpNick -replace ("-","")
    $StaffGrpDesc = "All Staff in Alpha Org 4 " +  $Script:txtGrpLocName.Text
    $Found = ""
    Foreach ($chk in $AO4Groups)
    {
        If ($chk.DisplayName -eq $Script:ULCOMStaffGrpName)
        {
            $Found = "Yes"
        }
    }
    If ($Found -eq "")
    {
        Write-host "Creating" $Script:ULCOMStaffGrpName -ForegroundColor Green
        New-AzureADMSGroup -Description $StaffGrpDesc -DisplayName $Script:ULCOMStaffGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $StaffGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute4 -eq ""$bute4"")" -MembershipRuleProcessingState "On"
        $Script:Body = $Script:Body + "<li>$Script:ULCOMStaffGrpName</li>"
    }
    else
    {
        Write-host "Group" $Script:ULCOMStaffGrpName "already exists" -ForegroundColor Red
    }

    If ($Script:Body.Length -gt 115)
    {
        Send-Email
    }
}

Function Create-DSGAO2Groups
{
    $AO2Groups = Get-AzureADMSGroup -SearchString "DSG.AO2"
    $cnt = 0

    $bute8 = $Script:txtGrpLocName.Text
    $AO2 = $bute8 -replace (",","")

    $EmpGrpNick = "DSG.AO2." + $Script:txtGrpLocName.Text + "EmpOnly" -replace("[(),&) ]","")
    $EmpGrpNick = $EmpGrpNick -replace ("-","")
    $EmpGrpDesc = "All Employees in Alpha Org 2 " +  $Script:txtGrpLocName.Text
    $Found = ""
    Foreach ($chk in $AO2Groups)
    {
        If ($chk.DisplayName -eq $Script:ULCOMEmpGrpName)
        {
            $Found = "Yes"
        }
    }
    If ($Found -eq "")
    {
        $Script:ULCOMEmpGrpName = $Script:ULCOMEmpGrpName  -replace("[(),) ]","")
        $Script:ULCOMEmpGrpName = $Script:ULCOMEmpGrpName  -replace ("-","")
        Write-host "`nCreating" $Script:ULCOMEmpGrpName -ForegroundColor Green
        New-AzureADMSGroup -Description $EmpGrpDesc -DisplayName $Script:ULCOMEmpGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $EmpGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute1 -contains ""Employee"") and (user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute8 -eq ""$bute8"")" -MembershipRuleProcessingState "On"
        $Script:Subject = "New DSG.AO2 Dynamic Group Created - " + $Script:ULCOMEmpGrpName
        Send-Email
    }
    else
    {
        Write-host "`nGroup" $Script:ULCOMEmpGrpName "already exists" -ForegroundColor Red
    }

    $StaffGrpNick = "DSG.AO2." + $Script:txtGrpLocName.Text + "Staff" -replace("[(),&) ]","")
    $StaffGrpNick = $StaffGrpNick -replace ("-","")
    $StaffGrpDesc = "All Staff in Alpha Org 2 " +  $Script:txtGrpLocName.Text
    $Found = ""
    $Script:ULCOMStaffGrpName = $Script:ULCOMStaffGrpName  -replace("[(),) ]","")
    $Script:ULCOMStaffGrpName = $Script:ULCOMStaffGrpName  -replace ("-","")
    Foreach ($chk in $AO2Groups)
    {
        If ($chk.DisplayName -eq $Script:ULCOMStaffGrpName)
        {
            $Found = "Yes"
        }
    }
    If ($Found -eq "")
    {
        Write-host "Creating" $Script:ULCOMStaffGrpName -ForegroundColor Green
        New-AzureADMSGroup -Description $StaffGrpDesc -DisplayName $Script:ULCOMStaffGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $StaffGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute8 -eq ""$bute8"")" -MembershipRuleProcessingState "On"
        $Script:Subject = "New DSG.AO2 Dynamic Group Created - " + $Script:ULCOMStaffGrpName
        Send-Email
    }
    else
    {
        Write-host "Group" $Script:ULCOMStaffGrpName "already exists" -ForegroundColor Red
    }
}

Function Create-DSGAOPPLAO2Group
{
    $AO2Groups = Get-AzureADMSGroup -SearchString "DSG.AO2"
    $cnt = 0

    $bute8 = $Script:txtGrpLocName.Text
    $AO2 = $bute8 -replace (",","")

    $EmpGrpNick = "DSG.AO2." + $Script:txtGrpLocName.Text + " People Leaders" -replace("[(),&) ]","")
    $EmpGrpNick = $EmpGrpNick -replace ("-","")
    $EmpGrpDesc = "People Leaders in Alpha Org 2 " +  $Script:txtGrpLocName.Text
    $Found = ""
    $Script:ULCOMEmpGrpName = $Script:ULCOMEmpGrpName  -replace("[(),) ]","")
    $Script:ULCOMEmpGrpName = $Script:ULCOMEmpGrpName  -replace ("-","")
    Foreach ($chk in $AO2Groups)
    {

        If ($chk.DisplayName -eq $Script:ULCOMEmpGrpName)
        {
            $Found = "Yes"
        }
    }
    If ($Found -eq "")
    {
        Write-host "`nCreating" $Script:ULCOMEmpGrpName -ForegroundColor Green
        New-AzureADMSGroup -Description $EmpGrpDesc -DisplayName $Script:ULCOMEmpGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $EmpGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute8 -eq ""$bute8"") and (user.extensionAttribute9 -eq ""Y"")" -MembershipRuleProcessingState "On"
        $Script:Subject = "New DSG.AO2 People Leader Dynamic Group Created - " + $Script:ULCOMEmpGrpName
        Send-Email
    }
    else
    {
        Write-host "`nGroup" $Script:ULCOMEmpGrpName "already exists" -ForegroundColor Red
    }
}

Function Create-DSGAOPPLAO4Group
{
    $AO4Groups = Get-AzureADMSGroup -SearchString "DSG.AO4"
    $cnt = 0

    $bute8 = $Script:txtGrpLocName.Text
    $AO4 = $bute8 -replace (",","")

    $EmpGrpNick = "DSG.AO4." + $Script:txtGrpLocName.Text + " People Leaders" -replace("[(),&) ]","")
    $EmpGrpNick = $EmpGrpNick -replace ("-","")
    $EmpGrpDesc = "People Leaders in Alpha Org 4 " +  $Script:txtGrpLocName.Text
    $Found = ""
    $Script:ULCOMEmpGrpName = $Script:ULCOMEmpGrpName  -replace("[(),) ]","")
    $Script:ULCOMEmpGrpName = $Script:ULCOMEmpGrpName  -replace ("-","")
    Foreach ($chk in $AO4Groups)
    {
        If ($chk.DisplayName -eq $Script:ULCOMEmpGrpName)
        {
            $Found = "Yes"
        }
    }
    If ($Found -eq "")
    {
        Write-host "`nCreating" $Script:ULCOMEmpGrpName -ForegroundColor Green
        New-AzureADMSGroup -Description $EmpGrpDesc -DisplayName $Script:ULCOMEmpGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $EmpGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute4 -eq ""$bute8"") and (user.extensionAttribute9 -eq ""Y"")" -MembershipRuleProcessingState "On"
        $Script:Subject = "New DSG.AO4 People Leader Dynamic Group Created - " + $Script:ULCOMEmpGrpName
        Send-Email
    }
    else
    {
        Write-host "`nGroup" $Script:ULCOMEmpGrpName "already exists" -ForegroundColor Red
    }
}

Function Apply-UseRestriction ($Script:UseGroup)
{
    If (($Script:txtAuthUsers.Text -like "Enter*") -or ($Script:txtAuthUsers.Text.Length -eq 0))
    {
        #Do nothing no group use restrictions being applied
    }
    else
    {
        If ($Script:txtAuthUsers.Text.length -ge 0)
        {
            $Auth = (($Script:txtAuthUsers.Text) -split(","))
            foreach ($member in $auth)
            {
                $OldUsers = (Get-DynamicDistributionGroup -identity $Script:UseGroup).AcceptMessagesOnlyFromSendersOrMembers
                Set-DynamicDistributionGroup $Script:UseGroup -AcceptMessagesOnlyFromSendersOrMembers ($OldUsers+=$member)
                write-host "Added Authorized user: " $member
            }
        }
    }
}

Function Send-Email
{
    $server = "smtp-relay.ul.com"
    $client = new-object system.net.mail.smtpclient $server
    $from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"
    $to = $from 
    $message = new-object  System.Net.Mail.MailMessage $from, $to 
    $message.IsBodyHtml = $true
    $msgfont = "<basefont face=verdana size=2.5 color=black>"

    $message.Subject = $Script:Subject
    $message.To.Clear()
    $message.Body = $msgfont + $message.body
    $message.To.Add("Lucja.Glodny@ul.com,Brandon.Richter@ul.com,Vinay.Puligundla@ul.com,Abhishek.Jain@ul.com")
    $message.cc.Add("BJ.Stone@ul.com,Gary.Herbold@ul.com")
    $client.Send($message)
}

Function Set-Mailtip
{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Name,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $Note,
         [Parameter(Mandatory=$true, Position=2)]
         [string] $Tip
    )

    If ($Tip.Length -gt 175)
    {
        $Tip = $Tip.SubString(0,175)
    }
    Set-DynamicDistributionGroup $Name -Notes $Note -MailTip $Tip
}

$Year = (get-date).ToString("yyyy")
$Path = "e:\Automation\NewDynamicGroup\Report\" + $Year
If (Test-Path $Path) {} else {New-Item -Path $Path -ItemType Directory}

Do {
    Build-DynGroupsMenu
    Publish-Form

    If ($Global:Result -eq "OK")
    {
        If (($Script:chkDSTSupGrp.Checked -eq "Checked") -or ($Script:chkDSTSupPPL.Checked -eq "Checked"))
        {
            $OutFileName = $Path + "\Report-" + $Script:txtGrpLastName.Text + $Script:txtGrpFrstName.Text  + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
        }
        else
        {
            $OutFileName = $Path + "\Report-" + ($Script:txtGrpLocName.Text -replace "[ /]","") + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
        }

        $WhoAmI	= WhoAmI
        $LineToWrite = "STAR" + "`t" + "NewDynDistGroup script has started"
        Write-Output $LineToWrite >> $OutFileName
	    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI
        Write-Output $LineToWrite >> $OutFileName
        Write-Output "" >> $OutFileName
        Start-Transcript -Path $OutFileName

        If (($Script:chkDSTAllGrp.Checked -eq $True) -or ($Script:chkDSTSupGrp.Checked -eq $True) -or ($Script:chkDSGAOPPL.Checked -eq $True) -or ($Script:chkDSTAllOrgGrp.Checked -eq $True) -or ($Script:chkDSTSupPPL.Checked -eq $True))
        {
            Create-DSTAllGroups
  	    }

	    If ($Script:chkDSGAO2.Checked -eq "Checked")
	    {
		    Create-DSGAO2Groups
  	    }

	    If ($Script:chkDSGAO4.Checked -eq "Checked")
	    {
		    Create-DSGAO4Groups
  	    }

        If ($Script:chkDSTAllPPL.Checked -eq $True)
        {
            If ($Script:AO2Group -eq "Yes")
            {
                Create-DSGAOPPLAO2Group
            }

            If ($Script:AO4Group -eq "Yes")
            {
                Create-DSGAOPPLAO4Group
            }
        }
    
        Stop-Transcript
    }
}While ($Global:Result -eq "OK")	