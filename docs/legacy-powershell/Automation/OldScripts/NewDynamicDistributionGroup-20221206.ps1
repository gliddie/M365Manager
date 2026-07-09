<####  This script create new UL Employee Only and UL Staff Dynamic Distribution Lists
#
# Called by:  DistributionSecurityGroupMenu.ps1
#
# 06/17/2016 - SAG - Added code to check for Displaynames longer than 64 chara and truncate them
#                  - Removed the MailUsers configuraiton as these individuals are mail enabled but not licensed
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
#>

Function Build-DynGroupsMenu
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "New DST/DSG Groups Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 480 ; $form.Height = 300  # Make the form wider 
    
    Add-FormStandardButtons
    $Global:form.AcceptButton.visible = $false
    
    $TopLoc = 20 
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
        $Script:chkDSTAllPPL.TabIndex = 3
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
        $Script:chkDSTSupGrp.TabIndex = 2
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
        $Script:chkDSTSupPPL.TabIndex = 4
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
        $Script:chkDSGAO4.TabIndex = 5
        $Global:form.Controls.Add($Script:chkDSGAO4)
        $Script:chkDSGAO4.Add_Click({
            $Script:lblGrpLocName.Text = "Enter DSG.AO4 Alpha Org Level 4 Name:"
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

    ## Details to Add First Name
    $Script:lblGrpFrstName = New-Object System.Windows.Forms.Label   
        $Script:lblGrpFrstName.Top = $TopLoc + 30 ; $Script:lblGrpFrstName.Left = 30; $Script:lblGrpFrstName.Width=70
        $Script:lblGrpFrstName.Text = "First Name:" 
        $Script:lblGrpFrstName.visible = $false
        $Global:form.Controls.Add($Script:lblGrpFrstName)    # Add to Form 
        # 
        $Script:txtGrpFrstName = New-Object Windows.Forms.TextBox  
        $Script:txtGrpFrstName.Top = $TopLoc + 30; $Script:txtGrpFrstName.Left = 100; $Script:txtGrpFrstName.Width = 80;
        $Script:txtGrpFrstName.ReadOnly = $true
        $Script:txtGrpFrstName.Text = ""
        $Script:txtGrpFrstName.visible = $false
        $Global:form.Controls.Add($Script:txtGrpFrstName)    # Add to Form

    ## Details to Add Last Name
    $Script:lblGrpLastName = New-Object System.Windows.Forms.Label   
        $Script:lblGrpLastName.Top = $TopLoc + 30 ; $Script:lblGrpLastName.Left = 190; $Script:lblGrpLastName.Width=70
        $Script:lblGrpLastName.Text = "Last Name:" 
        $Script:lblGrpLastName.visible = $false
        $Global:form.Controls.Add($Script:lblGrpLastName)    # Add to Form 
        # 
        $Script:txtGrpLastName = New-Object Windows.Forms.TextBox  
        $Script:txtGrpLastName.Top = $TopLoc + 30; $Script:txtGrpLastName.Left = 260; $Script:txtGrpLastName.Width = 80;
        $Script:txtGrpLastName.ReadOnly = $true
        $Script:txtGrpLastName.Text = ""
        $Script:txtGrpLastName.visible = $false
        $Global:form.Controls.Add($Script:txtGrpLastName)    # Add to Form

    ## List of Authorized Users
    $Script:lblAuthUsers = New-Object System.Windows.Forms.Label   
        $Script:lblAuthUsers.Top = $TopLoc + 60 ; $Script:lblAuthUsers.Left = 10; $Script:lblAuthUsers.Width=110
        $Script:lblAuthUsers.Text = "Auth User/Mailbox:"
        $Script:lblAuthUsers.visible = $false
        $Global:form.Controls.Add($Script:lblAuthUsers)    # Add to Form 
        # 
        $Script:txtAuthUsers = New-Object Windows.Forms.TextBox  
        $Script:txtAuthUsers.Top = $TopLoc + 60; $Script:txtAuthUsers.Left = 120; $Script:txtAuthUsers.Width = 300;
        $Script:txtAuthUsers.Text = ""
        $Script:txtAuthUsers.visible = $false
        $Global:form.Controls.Add($Script:txtAuthUsers)    # Add to Form

    $GrpTop = 280
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

    $GrpTop = $GrpTop + 30
    ## Details to UL.ORG Employees Only Group
    $Script:lblULORGEmpOnlyGrp = New-Object System.Windows.Forms.Label   
        $Script:lblULORGEmpOnlyGrp.Top = $GrpTop ; $Script:lblULORGEmpOnlyGrp.Left = 10; $Script:lblULORGEmpOnlyGrp.Width=110
        $Script:lblULORGEmpOnlyGrp.Text = "ULORG Emp Only:"
        $Script:lblULORGEmpOnlyGrp.visible = $false
        $Global:form.Controls.Add($Script:lblULORGEmpOnlyGrp)    # Add to Form 
        # 
        $Script:txtULORGEmpOnlyGrp = New-Object Windows.Forms.TextBox  
        $Script:txtULORGEmpOnlyGrp.Top = $GrpTop; $Script:txtULORGEmpOnlyGrp.Left = 120; $Script:txtULORGEmpOnlyGrp.Width = 300;
        $Script:txtULORGEmpOnlyGrp.visible = $false
        $Global:form.Controls.Add($Script:txtULORGEmpOnlyGrp)    # Add to Form 

    $GrpTop = $GrpTop + 30
    ## Details to UL.ORG Staff Group
    $Script:lblULORGStaffGrp = New-Object System.Windows.Forms.Label   
        $Script:lblULORGStaffGrp.Top = $GrpTop ; $Script:lblULORGStaffGrp.Left = 10; $Script:lblULORGStaffGrp.Width=110
        $Script:lblULORGStaffGrp.Text = "ULORG All Staff:"
        $Script:lblULORGStaffGrp.visible = $false
        $Global:form.Controls.Add($Script:lblULORGStaffGrp)    # Add to Form 
        # 
        $Script:txtULORGStaffGrp = New-Object Windows.Forms.TextBox  
        $Script:txtULORGStaffGrp.Top = $GrpTop; $Script:txtULORGStaffGrp.Left = 120; $Script:txtULORGStaffGrp.Width = 300;
        $Script:txtULORGStaffGrp.visible = $false
        $Global:form.Controls.Add($Script:txtULORGStaffGrp)    # Add to Form 

    $Script:ButBldGroups = New-Object Windows.Forms.Button
        $Script:ButBldGroups.Location = New-object System.Drawing.Size(130,240)
        $Script:ButBldGroups.Size = new-Object System.Drawing.Size(130,20)
        $Script:ButBldGroups.Text = "Verify Group Names"
        $Script:ButBldGroups.visible = $false
        $Global:form.Controls.Add($Script:ButBldGroups)
        $Script:ButBldGroups.Add_Click({
            $Script:ULCOMEmpGrpName = "N/A"
            $Script:ULCOMStaffGrpName = "N/A"
            $Script:ULORGEmpGrpName = "N/A"
            $Script:ULORGStaffGrpName = "N/A"
            ValidateLocation
            If ($Script:NotFound -eq "N")
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
                If ($Script:txtULORGEmpOnlyGrp.Text -ne "N/A")
                {
                    $Global:form.Height = 480
                    $Script:lblULORGEmpOnlyGrp.visible = $true
                    $Script:txtULORGEmpOnlyGrp.visible = $true
                }
                If ($Script:txtULORGStaffGrp.Text -ne "N/A")
                {
                    $Script:lblULORGStaffGrp.visible = $true
                    $Script:txtULORGStaffGrp.visible = $true
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
    $Script:lblULORGEmpOnlyGrp.visible = $false
    $Script:txtULORGEmpOnlyGrp.visible = $false
    $Script:lblULCOMStaffGrp.visible = $false
    $Script:txtULCOMStaffGrp.visible = $false
    $Script:lblULORGStaffGrp.visible = $false
    $Script:txtULORGStaffGrp.visible = $false
    $Script:lblAuthUsers.visible = $false
    $Script:txtAuthUsers.visible = $false
    If (($Script:chkDSTSupGrp.Checked -eq $True) -or ($Script:chkDSTSupPPL.Checked -eq $True))
    {
        $Script:ButBldGroups.Visible = $true
    }
}

Function DSTSUPDetails
{
    $Script:NotFound = ""
    $Script:ButBldGroups.Location = New-object System.Drawing.Size(130,250)
    $Global:form.Height = 380
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
    $Script:ButBldGroups.Location = New-object System.Drawing.Size(130,240)
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
    $Script:lblULORGEmpOnlyGrp.Text = "ULORG Emp Only:"
    $Script:lblULORGEmpOnlyGrp.visible = $false
    $Script:txtULORGEmpOnlyGrp.visible = $false
    $Script:lblULORGStaffGrp.visible = $false
    $Script:txtULORGStaffGrp.visible = $false
    $Script:ButBldGroups.visible = $true
    $Script:SupVal = "N"
    $Global:InputFocus = $Global:cancelButton
}

Function SUPDetails
{
    $Script:NotFound = "N"
    $Exists = [bool](get-mailbox $Script:txtGrpEmpNo.Text -ErrorAction SilentlyContinue)

    If ($Exists -eq $True)
    {
        $SUP = get-ADUser $Script:txtGrpEmpNo.Text
        $Script:txtGrpFrstName.Text = $SUP.GivenName
        $Script:txtGrpLastName.Text = $SUP.SurName
        $Script:txtGrpFrstName.Refresh()
        $Script:txtGrpLastName.Refresh()
        $Supervisor1 = "Yes"

        $Output = $wshell.Popup("Checking if this individual is a Supervisor1 or Supervisor2.  Be patient while data is extracted from the OED_Extract file.",5,"Validate Supervisor",0+32)
        $SupExists = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv | Where-Object{($_."Supervisor_1_NUMBER" -eq $Script:txtGrpEmpNo.Text)}
        If ($SupExists.count -lt 1)
        {
            $Supervisor1 = "No"
            $SupExists = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv | Where-Object{($_."Supervisor_2_NUMBER" -eq $Script:txtGrpEmpNo.Text)}
            If ($SupExists.count -lt 1)
            {
                $Supervisor1 = "NotFound"
                $Output = $wshell.Popup("Employee Number " + $Script:txtGrpEmpNo.Text + " is not listed as a Supervisor1 or Supervisor2.",0,"Not Sup1/Sup2",0+32)
                $Global:form.AcceptButton.visible = $false  
            }
            else
            {
                $Global:form.AcceptButton.visible = $true            }
        }
        else
        {
                $Global:form.AcceptButton.visible = $false        
        }
    }
    else
    {
        $Output = $wshell.Popup("This individual " + $Script:txtGrpEmpNo.Text + " does not have an active mailbox.",5,"Validate Supervisor",0+32)
        $Script:NotFound = "Y"
        $Script:txtGrpEmpNo.Text = ""
        $Global:form.AcceptButton.visible = $false
    }
}

Function ValidateLocation
{
    $Script:NotFound = "N"
    $Script:ActiveULCOMDSTRecs = ""
    $Script:ActiveULORGDSTRecs = ""
    $Script:AO2Group = "No"
    $Script:AO4Group = "No"

    If (($Script:chkDSTAllGrp.Checked -eq $true) -or ($Script:chkDSTAllPPL.Checked -eq $True))
    {
        $Output = $wshell.Popup("Gathering all sites with Active UL.COM and UL.ORG staff for this location.  Be patient while data is extracted from the OED_Extract file.",10,"Locations with Active Staff",0+32)
        $Script:ActiveULCOMDSTRecs = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv | Where-Object{((($_.ALPHA_HR_ORG_LEVEL_1 -eq "UL Inc.") -and ($_.Location -eq $Script:txtGrpLocName.Text)) -and (($_."Assignment Status" -eq "A") -or ($_."Assignment Status" -eq "F")))}
        $Script:ActiveULORGDSTRecs = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract6.csv | Where-Object{((($_.ALPHA_HR_ORG_LEVEL_1 -ne "UL Inc.") -and ($_.Location -eq $Script:txtGrpLocName.Text)) -and (($_."Assignment Status" -eq "A") -or ($_."Assignment Status" -eq "F")))} 
    }

    If (($Script:chkDSTSupGrp.Checked -eq $True) -or ($Script:chkDSTSupPPL.Checked -eq $True))
    {
            SUPDetails
    }

    If ($Script:chkDSTAllOrgGrp.Checked -eq $True)
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

    If (($Script:ActiveULCOMDSTRecs.count.length -ge 1) -or ($Script:ActiveULORGDSTRecs.count.length -ge 1))
#    If (($Script:ActiveULCOMDSTRecs.count -gt 0) -or ($Script:ActiveULORGDSTRecs.count -gt 0))
    {
        $Script:lblULCOMEmpOnlyGrp.Text = "ULCOM Emp Only:"
        $Script:lblULCOMStaffGrp.Text = "ULCOM All Staff:"
        If ($Script:chkDSTSupGrp.Checked -eq $True)
        {
                $Script:lblULCOMEmpOnlyGrp.Text = "Emp Only Grp:"
                $Script:lblULCOMStaffGrp.Text = "All Staff Grp:"
        }
        If ($Script:chkDSTSupPPL.Checked -eq $True)
        {
            $Script:lblULCOMEmpOnlyGrp.Text = "        Group Name:"
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
        BuildGroupNames
    }
    else
    {
        $Script:NotFound = "Y"
        $Script:txtGrpLocName.Text = ""
        $Output = $wshell.Popup("There are no active Staff in the identified location " + $Script:txtGrpLocName.Text,10,"No Active Staff",0+32)
    }
}

Function BuildGroupNames
{
    $Global:form.AcceptButton.visible = $true
    $Script:ULCOMEmpGrpName = "N/A"
    $Script:ULCOMStaffGrpName = "N/A"
    $Script:ULORGEmpGrpName = "N/A"
    $Script:ULORGStaffGrpName = "N/A"
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
        If ($Script:ActiveULORGDSTRecs.Location -eq $Script:txtGrpLocName.Text)
        {
            $Script:ULORGEmpGrpName = "DST.All UL ORG " + $Script:txtGrpLocName.Text + " Employees Only"
    	    $Script:ULORGStaffGrpName = "DST.All UL ORG " + $Script:txtGrpLocName.Text + " Staff"
        }
        else
        {
            $Script:ULORGEmpGrpName = "N/A"
            $Script:ULORGStaffGrpName = "N/A"
            $Output = $wshell.Popup("No Active UL.ORG Staff at the site:" + $Script:txtGrpLocName.Text ,5,"Invalid Location",0+32)
            If ($Script:ULCOMEmpGrpName -eq "N/A")
            {
                $Global:form.AcceptButton.visible = $false
            }
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
        If ($Script:ActiveULORGDSTRecs.Location -eq $Script:txtGrpLocName.Text)
        {
            $Script:lblULORGEmpOnlyGrp.Text = "  ULORG Group:"
            $Script:ULORGEmpGrpName = "DST.All UL ORG " + $Script:txtGrpLocName.Text + " People Leaders"
            $Global:form.AcceptButton.visible = $true
        }
        else
        {
            $Script:ULORGEmpGrpName = "N/A"
            $Output = $wshell.Popup("No Active UL.ORG Staff at the site:" + $Script:txtGrpLocName.Text ,5,"Invalid Location",0+32)
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

    $Script:txtULCOMEmpOnlyGrp.Text = $Script:ULCOMEmpGrpName
    $Script:txtULCOMStaffGrp.Text = $Script:ULCOMStaffGrpName
    $Script:ULCOMEmpGrpNameTrunc = $Script:ULCOMEmpGrpName
    $Script:ULCOMStaffGrpNameTrunc = $Script:ULCOMStaffGrpName

    $Script:txtULORGEmpOnlyGrp.Text = $Script:ULORGEmpGrpName
    $Script:txtULORGStaffGrp.Text = $Script:ULORGStaffGrpName
    $Script:ULORGEmpGrpNameTrunc = $Script:ULORGEmpGrpName
    $Script:ULORGStaffGrpNameTrunc = $Script:ULORGStaffGrpName

    If ($Script:ULCOMEmpGrpName.Length -gt 64)
    {
        $Script:ULCOMEmpGrpNameTrunc = $Script:ULCOMEmpGrpName.Substring(0,64)
    }
    
    If ($Script:ULCOMStaffGrpName.Length -gt 64)
    {
        $Script:ULCOMStaffGrpNameTrunc = $Script:ULCOMStaffGrpName.Substring(0,64)
    }
    
    If ($Script:ULORGEmpGrpName.Length -gt 64)
    {
        $Script:ULORGEmpGrpNameTrunc = $Script:ULORGEmpGrpName.Substring(0,64)
    }
    
    If ($Script:ULORGStaffGrpName.Length -gt 64)
    {
        $Script:ULORGStaffGrpNameTrunc = $Script:ULORGStaffGrpName.Substring(0,64)
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
            write-host "Name of the group for UL.ORG People Leaders: " $Script:ULORGEmpGrpName
        }
        else
        {
            write-host "Name of the group for UL.COM Employees Only: " $Script:ULCOMEmpGrpName
            write-host "     Name of the group for all UL.COM Staff: " $Script:ULCOMStaffGrpName
            write-host "Name of the group for UL.ORG Employees Only: " $Script:ULORGEmpGrpName
            write-host "     Name of the group for all UL.ORG Staff: " $Script:ULORGStaffGrpName
        }

        $ULCOMEmpGrpAlias = $Script:ULCOMEmpGrpName -replace("[(),&) ]","")
        $ULCOMEmpGrpAlias = $ULCOMEmpGrpAlias -replace ("-","")
        $ULCOMEmpGrpAlias = $ULCOMEmpGrpAlias.Replace("UL","")
        $ULCOMEmpGrpAlias = $ULCOMEmpGrpAlias.Replace("loyees","")
        $ULCOMEmpINetAlias = $ULCOMEmpGrpAlias + "@ul.com"

        $ULORGEmpGrpAlias = $Script:ULORGEmpGrpName -replace("[(),&) ]","")
        $ULORGEmpGrpAlias = $ULORGEmpGrpAlias -replace ("-","")
        $ULORGEmpGrpAlias = $ULORGEmpGrpAlias.Replace("UL","")
        $ULORGEmpGrpAlias = $ULORGEmpGrpAlias.Replace("loyees","")
        $ULORGEmpINetAlias = $ULORGEmpGrpAlias + "@ul.org"

        $ULCOMStaffGrpAlias = $Script:ULCOMStaffGrpName -replace("[(),&) ]","")
        $ULCOMStaffGrpAlias = $ULCOMStaffGrpAlias -replace ("-","")
        $ULCOMStaffGrpAlias = $ULCOMStaffGrpAlias.Replace("UL","")
        $ULCOMStaffINetAlias = $ULCOMStaffGrpAlias + "@ul.com"

        $ULORGStaffGrpAlias = $Script:ULORGStaffGrpName -replace("[(),&) ]","")
        $ULORGStaffGrpAlias = $ULORGStaffGrpAlias -replace ("-","")
        $ULORBStaffGrpAlias = $ULORGStaffGrpAlias.Replace("UL","")
        $ULORGStaffINetAlias = $ULORGStaffGrpAlias + "@ul.org"

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
                    Set-DynamicDistributionGroup $Script:ULCOMEmpGrpName -Notes $ULCOMEmpGrpNote -MailTip $ULCOMMailTip
                    $Script:UseGroup = $Script:ULCOMEmpGrpName
                    Apply-UseRestriction
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
                    Set-DynamicDistributionGroup $Script:ULCOMStaffGrpName -Notes $ULCOMStaffGrpNote -MailTip $ULCOMStaffTip
                    $Script:UseGroup = $Script:ULCOMStaffGrpName
                    Apply-UseRestriction
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMStaffGrpName,10,"Group Exists",0+32)
                }
            }

            #Create UL.ORG Groups
            If ($Script:ActiveULORGDSTRecs.Location -eq $Script:txtGrpLocName.Text)
            {
                $Exists = [bool](get-DynamicDistributionGroup $Script:ULORGEmpGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
                    Write-Host "Creating New Group for: " $Script:ULORGEmpGrpName -ForegroundColor Green
                    New-DynamicDistributionGroup -DisplayName $Script:ULORGEmpGrpName -Alias $ULORGEmpGrpAlias -Name $Script:ULORGEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCompany "UL.ORG" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -PrimarySmtpAddress $ULORGEmpINetAlias
    	            $ULORGEmpGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff in UL.ORG that have an employee type of ""Employee"" and are in the ""$LocName"" location.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                    $ULORGMailTip = "Sends all UL.ORG Employees Only assigned to the '" + $LocName + "' site in Oracle."
                    Set-DynamicDistributionGroup $Script:ULORGEmpGrpName -Notes $ULORGEmpGrpNote -MailTip $ULORGMailTip
                    Apply-UseRestriction
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULORGEmpGrpName,10,"Group Exists",0+32)
                }

                $Exists = [bool](get-DynamicDistributionGroup $Script:ULORGStaffGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
	                Write-Host "Creating New Group for: " $Script:ULORGStaffGrpName -ForegroundColor Green
	                New-DynamicDistributionGroup -DisplayName $Script:ULORGStaffGrpName -Alias $ULORGStaffGrpAlias -Name $Script:ULORGStaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCompany "UL.ORG" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -PrimarySmtpAddress $ULORGStaffINetAlias
                    $ULORGStaffGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff in UL.ORG that have an employee type of ""Employee, Consultant, Contractor, Temporary Worker or Freelance Worker"" and are in the ""$LocName"" location.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                    $ULORGStaffTip = "Sends to all UL.ORG Employee and Non-Employee types assigned to the '" + $LocName + "' site in Oracle."
                    Set-DynamicDistributionGroup $Script:ULORGStaffGrpName -Notes $ULORGStaffGrpNote -MailTip $ULORGStaffTip
                    Script:UseGroup = $Script:ULORGStaffGrpName
                    Apply-UseRestriction
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULORGStaffGrpName,10,"Group Exists",0+32)
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
                    Set-DynamicDistributionGroup $Script:ULCOMEmpGrpName -Notes $ULCOMEmpGrpNote -MailTip $ULCOMMailTip
                    Script:UseGroup = $Script:ULCOMEmpGrpName
                    Apply-UseRestriction
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
                    Set-DynamicDistributionGroup $Script:ULCOMStaffGrpName -Notes $ULCOMStaffGrpNote -MailTip $ULCOMStaffTip
                    Script:UseGroup = $Script:ULCOMStaffGrpName
                    Apply-UseRestriction
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
                    Set-DynamicDistributionGroup $Script:ULCOMEmpGrpName -Notes $ULCOMEmpGrpNote -MailTip $ULCOMMailTip
                    Script:UseGroup = $Script:ULCOMEmpGrpName
                    Apply-UseRestriction
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
                    Set-DynamicDistributionGroup $Script:ULCOMStaffGrpName -Notes $ULCOMStaffGrpNote -MailTip $ULCOMStaffTip
                    Script:UseGroup = $Script:ULCOMStaffGrpName
                    Apply-UseRestriction
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMStaffGrpName,10,"Group Exists",0+32)
                }
            }
        }

        If ($Script:chkDSTSupGrp.Checked -eq $True)
        {
            If ($Supervisor1 -eq "Yes")
            {
                $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMEmpGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
                    Write-Host "Creating New Group for: " $Script:ULCOMEmpGrpName -ForegroundColor Green
                    New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute6 $Script:txtGrpEmpNo.Text -PrimarySmtpAddress $ULCOMEmpINetAlias
                    $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisor1 in Oracle.  This group contains staff in UL.COM that have an employee type of ""Employee"".  Manual modification of this group is not possible."
                    $ULCOMMailTip = "Sends all Employees that report to '" + $FrstName + " " + $LastName+ "' as Supervisor1 in Oracle."
                    Script:UseGroup = $Script:ULCOMEmpGrpName
                    Apply-UseRestriction
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
                    $ULCOMStaffGrpNote = "The membership of this group is determined at the time the group and contains individuals that report up to " + $FrstName + " " + $Last + " as Supervisor1 in Oracle.  This group contains staff in UL.COM that have an employee type of ""Employee, Consultant, Contractor or Temporary Worker"".  Manual modification of this group is not possible."
                    $ULCOMStaffTip = "Sends to all Employee and Non-Employee types that report to '"  + $FrstName + " " + $LastName + "' as Supervisor1 in Oracle."
                    Script:UseGroup = $Script:ULCOMStaffGrpName
                    Apply-UseRestriction
                }
                else
                {
                    Write-Host "Group already exists for: " $Script:ULCOMStaffGrpName -ForegroundColor Red
                }
            }
            else
            {
                $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMEmpGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
                    Write-Host "Creating New Group for: " $Script:ULCOMEmpGrpName -ForegroundColor Green
                    New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute7 $Script:txtGrpEmpNo.Text -PrimarySmtpAddress $ULCOMEmpINetAlias
        	        $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisor2 in Oracle.  This group contains staff in UL.COM that have an employee type of ""Employee"".  Manual modification of this group is not possible."
                    $ULCOMMailTip = "Sends to individuals listed as Employee that report to '" + $FrstName + " " + $LastName + "' as Supervisor2 in Oracle."
                    Write-host "Restricting Use of Group to: " $Script:txtGrpEmpNo.Text
                    Set-DynamicDistributionGroup $Script:ULCOMEmpGrpName -Notes $ULCOMEmpGrpNote -MailTip $ULCOMMailTip -RequireSenderAuthenticationEnabled $True -AcceptMessagesOnlyFromSendersOrMembers $Script:txtGrpEmpNo.Text
                    Script:UseGroup = $Script:ULCOMEmpGrpName
                    Apply-UseRestriction
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
                    Set-DynamicDistributionGroup $Script:ULCOMStaffGrpName -Notes $ULCOMStaffGrpNote -MailTip $ULCOMStaffTip -RequireSenderAuthenticationEnabled $True -AcceptMessagesOnlyFromSendersOrMembers $Script:txtGrpEmpNo.Text
                    $Script:UseGroup = $Script:ULCOMStaffGrpName
                    Apply-UseRestriction
                }
                else
                {
                    Write-Host "Group already exists for: " $Script:ULCOMStaffGrpName -ForegroundColor Red
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
                    Set-DynamicDistributionGroup $Script:ULCOMEmpGrpName -Notes $ULCOMEmpGrpNote -MailTip $ULCOMMailTip
                    $Script:UseGroup = $Script:ULCOMEmpGrpName
                    Apply-UseRestriction
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULCOMEmpGrpName,10,"Group Exists",0+32)
                }
            }

            #Create UL.ORG Groups
            If ($Script:ActiveULORGDSTRecs.Location -eq $Script:txtGrpLocName.Text)
            {
                $Exists = [bool](get-DynamicDistributionGroup $Script:ULORGEmpGrpName -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
	                Write-Host "Creating New Group for: " $Script:ULORGEmpGrpName -ForegroundColor Green
                    New-DynamicDistributionGroup -DisplayName $Script:ULORGEmpGrpName -Alias $ULORGEmpGrpAlias -Name $Script:ULORGEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCompany "UL.ORG" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $ULCOMEmpINetAlias
        	        $ULORGEmpGrpNote = "The membership of this group is determined at the time the group is used and contains UL.ORG individuals location in the ""$LocName"" location and is a People Leader in Oracle.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                    $ULORGMailTip = "Sends to individuals in UL.ORG listed as Employee assigned to the '" + $Script:txtGrpLocName.Text + "' site in Oracle."
                    Set-DynamicDistributionGroup $Script:ULORGEmpGrpName -Notes $ULORGEmpGrpNote -MailTip $ULORGMailTip
                    Script:UseGroup = $Script:ULORGEmpGrpName
                    Apply-UseRestriction
                }
                else
                {
                    $Output = $wshell.Popup("Group already exists for: " + $Script:ULORGEmpGrpName,10,"Group Exists",0+32)
                }
            }
        }

        If ($Script:chkDSTSupPPL.Checked -eq $True)
        {
            $Exists = [bool](get-DynamicDistributionGroup $Script:ULCOMEmpGrpName -ErrorAction SilentlyContinue)
            If ($Exists -eq $False)
            {
                Write-Host "Creating New Group for: " $Script:ULCOMEmpGrpName -ForegroundColor Green
                If ($Script:chkSUPLevel1.Checked -eq "Checked")
                {
                    New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute6 $Script:txtGrpEmpNo.Text -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $ULCOMEmpINetAlias
                    $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisor1 in Oracle.  This group contains staff that have an employee type of ""Employee"".  Manual modification of this group is not possible."
                    $ULCOMMailTip = "Sends to individuals listed as Employee that report to '" + $Script:txtGrpLastName.Text + " " + $Script:txtGrpFrstName.Text + "' as Supervisor1 in Oracle."
                    Set-DynamicDistributionGroup $Script:ULCOMEmpGrpName -Notes $ULCOMEmpGrpNote -MailTip $ULCOMMailTip
                    Script:UseGroup = $Script:ULCOMEmpGrpName
                    Apply-UseRestriction
                }
                else
                {
                    New-DynamicDistributionGroup -DisplayName $Script:ULCOMEmpGrpName -Alias $ULCOMEmpGrpAlias -Name $Script:ULCOMEmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute7 $Script:txtGrpEmpNo.Text -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $ULCOMEmpINetAlias
        	        $ULCOMEmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisro2 in Oracle.  This group contains staff that have an employee type of ""Employee"".  Manual modification of this group is not possible."
                    $ULCOMMailTip = "Sends to individuals listed as Employee that report to '" + $Script:txtGrpLastName.Text + " " + $Script:txtGrpLastName.Text + "' as Supervisor2 in Oracle."
                    Script:UseGroup = $Script:ULCOMEmpGrpName
                    Apply-UseRestriction
                }
                Set-DynamicDistributionGroup $Script:ULCOMEmpGrpName -Notes $ULCOMEmpGrpNote -MailTip $ULCOMMailTip
            else
            {
                $Output = $wshell.Popup("Group already exists for: " + $Script:ULORGEmpGrpName,10,"Group Exists",0+32)
            }
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
    If ($Found -eq "")
    {
        Write-host "`nCreating" $Script:ULCOMEmpGrpName -ForegroundColor Green
        New-AzureADMSGroup -Description $EmpGrpDesc -DisplayName $Script:ULCOMEmpGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $EmpGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute1 -contains ""Employee"") and (user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute4 -eq ""$bute4"")" -MembershipRuleProcessingState "On"
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
    }
    else
    {
        Write-host "Group" $Script:ULCOMStaffGrpName "already exists" -ForegroundColor Red
    }
}

Function Create-DSGAO2Groups
{
    $AO2Groups = Get-AzureADMSGroup -SearchString "DSG.AO2"
    $cnt = 0

    $bute8 = $Script:txtGrpLocName.Text
    $AO2 = $bute8 -replace (",","")

    $EmpGrpNick = "DSG.AO2." + $Script:txtGrpLocName.Text + "EmpOnly" -replace("[(),&) ]","")
    $EmpGrpNick - $EmpGrpNick -replace ("-","")
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
        Write-host "`nCreating" $Script:ULCOMEmpGrpName -ForegroundColor Green
        New-AzureADMSGroup -Description $EmpGrpDesc -DisplayName $Script:ULCOMEmpGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $EmpGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute1 -contains ""Employee"") and (user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute8 -eq ""$bute8"")" -MembershipRuleProcessingState "On"
    }
    else
    {
        Write-host "`nGroup" $Script:ULCOMEmpGrpName "already exists" -ForegroundColor Red
    }

    $StaffGrpNick = "DSG.AO2." + $Script:txtGrpLocName.Text + "Staff" -replace("[(),&) ]","")
    $StaffGrpNick = $StaffGrpNick -replace ("-","")
    $StaffGrpDesc = "All Staff in Alpha Org 2 " +  $Script:txtGrpLocName.Text
    $Found = ""
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
    }
    else
    {
        Write-host "Group" $Script:ULCOMStaffGrpName "already exists" -ForegroundColor Red
    }
}

Function Apply-UseRestriction
{
    If ($Script:txtAuthUsers.Text.length -gt 0)
    {
        $Auth = (($Script:txtAuthUsers.Text) -replace (",", "")) -split(",")
        foreach ($member in $auth)
        {
            $OldUsers = (Get-DynamicDistributionGroup -identity $Script:UseGroup).AcceptMessagesOnlyFromSendersOrMembers
            Set-DynamicDistributionGroup $Script:UseGroup -AcceptMessagesOnlyFromSendersOrMembers ($OldUsers+=$member)
            write-host "Added Authorized user: " $member
        }
    }
}

Do {
    Build-DynGroupsMenu
    Publish-Form

    If ($Global:Result -eq "OK")
    {
        If (($Script:chkDSTSupGrp.Checked -eq "Checked") -or ($Script:chkDSTSupPPL.Checked -eq "Checked"))
        {
            $OutFileName = "e:\Automation\NewDynamicGroup\Report\Report-" + $Script:txtGrpLastName.Text + $Script:txtGrpFrstName.Text  + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
        }
        else
        {
            $OutFileName = "e:\Automation\NewDynamicGroup\Report\Report-" + ($Script:txtGrpLocName.Text -replace "[ /]","") + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
        }

        $WhoAmI	= WhoAmI
        $LineToWrite = "STAR" + "`t" + "NewDynDistGroup script has started"
        Write-Output $LineToWrite >> $OutFileName
	    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI
        Write-Output $LineToWrite >> $OutFileName
        Write-Output "" >> $OutFileName
        Start-Transcript -Path $OutFileName

        If (($Script:chkDSTAllGrp.Checked -eq $True) -or ($Script:chkDSTSupGrp.Checked -eq $True) -or ($Script:chkDSTAllPPL.Checked -eq $True) -or ($Script:chkDSTAllOrgGrp.Checked -eq $True) -or ($Script:chkDSTSupPPL.Checked -eq $True))
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
    
        Stop-Transcript
    }
}While ($Global:Result -eq "OK")	