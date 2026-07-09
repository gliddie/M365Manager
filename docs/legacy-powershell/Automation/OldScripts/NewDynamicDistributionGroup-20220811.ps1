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
#>

Function Build-DynGroupsMenu
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "New DST/DSG Groups Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 450 ; $form.Height = 370  # Make the form wider 
    
    Add-FormStandardButtons

    $TopLoc = 20 
    $Script:lblTitleLine = New-Object System.Windows.Forms.Label   
        $lblTitleLine.Text = "Select Option:"
        $lblTitleLine.Top = 15 ; $lblTitleLine.Left = 60; $lblTitleLine.Width=120 ;$lblTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine)    # Add to Form 

    ## DST.All Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSTAllGrp = New-Object Windows.Forms.RadioButton 
        $Script:chkDSTAllGrp.Left = 100; $Script:chkDSTAllGrp.Width = 450; $Script:chkDSTAllGrp.Top = $TopLoc  
        $Script:chkDSTAllGrp.Text = "Create DST.All Groups" 
        $Script:chkDSTAllGrp.Checked = $false   # set a default value 
        $Script:chkDSTAllGrp.TabIndex = 1
        $Global:form.Controls.Add($Script:chkDSTAllGrp)
        $Script:chkDSTAllGrp.Add_Click({
            $Script:lblGrpLocName.Text = "Enter DST.All Location Name:"
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
            $Script:ButBldGroups.visible = $false
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
            $Script:ButBldGroups.visible = $false
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

    $TopLoc = $TopLoc + 20
        $Script:lblSeperatorLine = New-Object System.Windows.Forms.Label   
        $lblSeperatorLine.Text = "____________________________________________________________________"
        $lblSeperatorLine.Top = $TopLoc ; $lblSeperatorLine.Left = 1; $lblSeperatorLine.Width=200 ;$lblSeperatorLine.AutoSize = $true
        $lblSeperatorLine.Visible = $false 
        $Global:form.Controls.Add($lblSeperatorLine)    # Add to Form 

    $TopLoc = $TopLoc + 20
    ## Details to Add Location Information
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
            $Script:ButBldGroups.visible = $false
        })

    ## Confirm Supervisor1
    $TopLoc = $TopLoc + 20
    $Script:chkSUPLevel1 = New-Object Windows.Forms.CheckBox
        $Script:chkSUPLevel1.Left = 100; $Script:chkSUPLevel1.Width = 85; $Script:chkSUPLevel1.Top = $TopLoc  
        $Script:chkSUPLevel1.Text = "Supervisor1" 
        $Script:chkSUPLevel1.Checked = $false   # set a default value
        $Script:chkSUPLevel1.visible = $false 
        $Global:form.Controls.Add($Script:chkSUPLevel1)
        $Script:chkSUPLevel1.Add_Click({
            $Script:chkSUPLevel2.Checked = $false
            $Script:ButBldGroups.visible = $true
            SUPDetails
        })

    ## Confirm Supervisor2
    $Script:chkSUPLevel2 = New-Object Windows.Forms.CheckBox
        $Script:chkSUPLevel2.Left = 190; $Script:chkSUPLevel2.Width = 90; $Script:chkSUPLevel2.Top = $TopLoc  
        $Script:chkSUPLevel2.Text = "Supervisor2" 
        $Script:chkSUPLevel2.Checked = $false   # set a default value
        $Script:chkSUPLevel2.visible = $false 
        $Global:form.Controls.Add($Script:chkSUPLevel2)
        $Script:chkSUPLevel2.Add_Click({
            $Script:chkSUPLevel1.Checked = $false
            $Script:ButBldGroups.visible = $true
            SUPDetails
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
        $Script:txtGrpFrstName.Text = ""
        $Script:txtGrpFrstName.visible = $false
        $Global:form.Controls.Add($Script:txtGrpFrstName)    # Add to Form
        $Script:txtGrpFrstName.Add_Click({ SUPDetails })

    ## Details to Add Last Name
    $Script:lblGrpLastName = New-Object System.Windows.Forms.Label   
        $Script:lblGrpLastName.Top = $TopLoc + 30 ; $Script:lblGrpLastName.Left = 190; $Script:lblGrpLastName.Width=70
        $Script:lblGrpLastName.Text = "Last Name:" 
        $Script:lblGrpLastName.visible = $false
        $Global:form.Controls.Add($Script:lblGrpLastName)    # Add to Form 
        # 
        $Script:txtGrpLastName = New-Object Windows.Forms.TextBox  
        $Script:txtGrpLastName.Top = $TopLoc + 30; $Script:txtGrpLastName.Left = 260; $Script:txtGrpLastName.Width = 80;
        $Script:txtGrpLastName.Text = ""
        $Script:txtGrpLastName.visible = $false
        $Global:form.Controls.Add($Script:txtGrpLastName)    # Add to Form
        $Script:txtGrpFrstName.Add_Click({ SUPDetails })

    ## Details to UL Employees Only Group
    $Script:lblEmpOnlyGrp = New-Object System.Windows.Forms.Label   
        $Script:lblEmpOnlyGrp.Top = 230; $Script:lblEmpOnlyGrp.Left = 10; $Script:lblEmpOnlyGrp.Width= 100
        $Script:lblEmpOnlyGrp.Text = "Emp Only Group:"
        $Script:lblEmpOnlyGrp.visible = $false
        $Global:form.Controls.Add($Script:lblEmpOnlyGrp)    # Add to Form 
        # 
        $Script:txtEmpOnlyGrp = New-Object Windows.Forms.TextBox  
        $Script:txtEmpOnlyGrp.Top = 230; $Script:txtEmpOnlyGrp.Left = 120; $Script:txtEmpOnlyGrp.Width = 300;
        $Script:txtEmpOnlyGrp.visible = $false
        $Global:form.Controls.Add($Script:txtEmpOnlyGrp)    # Add to Form

    ## Details to UL Staff Group
    $Script:lblStaffGrp = New-Object System.Windows.Forms.Label   
        $Script:lblStaffGrp.Top = 260 ; $Script:lblStaffGrp.Left = 10; $Script:lblStaffGrp.Width=100
        $Script:lblStaffGrp.Text = "All Staff Group:"
        $Script:lblStaffGrp.visible = $false
        $Global:form.Controls.Add($Script:lblStaffGrp)    # Add to Form 
        # 
        $Script:txtStaffGrp = New-Object Windows.Forms.TextBox  
        $Script:txtStaffGrp.Top = 260; $Script:txtStaffGrp.Left = 120; $Script:txtStaffGrp.Width = 300;
        $Script:txtStaffGrp.visible = $false
        $Global:form.Controls.Add($Script:txtStaffGrp)    # Add to Form 

        $Script:ButBldGroups = New-Object Windows.Forms.Button
        $Script:ButBldGroups.Location = New-object System.Drawing.Size(120,230)
        $Script:ButBldGroups.Size = new-Object System.Drawing.Size(150,20)
        $Script:ButBldGroups.Text = "Verify Group Names"
        $Script:ButBldGroups.visible = $false
        $Global:form.Controls.Add($Script:ButBldGroups)
        $Script:ButBldGroups.Add_Click({
            BuildGroupNames
            $Script:lblEmpOnlyGrp.visible = $true
            $Script:txtEmpOnlyGrp.visible = $true
            $Script:lblStaffGrp.visible = $true
            $Script:txtStaffGrp.visible = $true
            $Script:ButBldGroups.visible = $false
        })
}

Function UpdateLocation
{
    $Script:txtGrpFrstName.Text = ""
    $Script:txtGrpLastName.Text = ""
    $Script:chkSUPLevel1.Checked = $false
    $Script:chkSUPLevel2.Checked = $false
    $Script:lblEmpOnlyGrp.visible = $false
    $Script:txtEmpOnlyGrp.visible = $false
    $Script:lblStaffGrp.visible = $false
    $Script:txtStaffGrp.visible = $false
    $Script:ButBldGroups.visible = $true
}

Function DSTSUPDetails
{
    $Global:form.Height = 390
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
    $Script:chkSUPLevel1.visible = $true 
    $Script:chkSUPLevel2.visible = $true
    $Script:lblEmpOnlyGrp.Top = 260
    $Script:lblEmpOnlyGrp.visible = $false
    $Script:txtEmpOnlyGrp.Top = 260
    $Script:txtEmpOnlyGrp.visible = $false
    $Script:lblStaffGrp.Top = 290
    $Script:lblStaffGrp.visible = $false
    $Script:txtStaffGrp.Top = 290
    $Script:txtStaffGrp.visible = $false
    $Script:ButBldGroups.Location = New-object System.Drawing.Size(120,260)
    $Script:ButBldGroups.visible = $false
    $Script:SupVal = "Y"
}

Function DSTAllDetails
{
    $Global:form.Height = 370
    $lblSeperatorLine.Visible = $false
    $Script:lblGrpLocName.visible = $true
    $Script:txtGrpLocName.visible = $true
    $Script:ButBldGroups.visible = $true
    $Script:txtGrpLocName.Focus()
    $Script:lblGrpEmpNo.visible = $false
    $Script:txtGrpEmpNo.visible = $false
    $Script:lblGrpFrstName.visible = $false
    $Script:txtGrpFrstName.visible = $false
    $Script:lblGrpLastName.visible = $false
    $Script:txtGrpLastName.visible = $false
    $Script:chkSUPLevel1.visible = $false
    $Script:chkSUPLevel2.visible = $false
    $Script:lblEmpOnlyGrp.Top = 230
    $Script:lblEmpOnlyGrp.visible = $false
    $Script:txtEmpOnlyGrp.Top = 230
    $Script:txtEmpOnlyGrp.visible = $false
    $Script:lblStaffGrp.Top = 260
    $Script:lblStaffGrp.visible = $false
    $Script:txtStaffGrp.Top = 260
    $Script:txtStaffGrp.visible = $false
    $Script:ButBldGroups.Location = New-object System.Drawing.Size(120,230)
    $Script:ButBldGroups.visible = $true
    $Script:SupVal = "N"
}

Function SUPDetails
{
    $SUP = get-ADUser $Script:txtGrpEmpNo.Text
    $Script:txtGrpFrstName.Text = $SUP.GivenName
    $Script:txtGrpFrstName.ReadOnly = $true
    $Script:txtGrpLastName.Text = $SUP.SurName
    $Script:txtGrpLastName.ReadOnly = $true
    $Script:txtGrpFrstName.Refresh()
    $Script:txtGrpLastName.Refresh()
}

Function BuildGroupNames
{
    If ($Script:chkDSTAllGrp.Checked -eq "Checked")
    {
        $Script:EmpGrpName = "DST.All " + $Script:txtGrpLocName.Text + " UL Employees Only"
	    $Script:StaffGrpName = "DST.All " + $Script:txtGrpLocName.Text + " UL Staff"
    }
    If ($Script:chkDSTSUPGrp.Checked -eq "Checked")
    {
        $Script:EmpGrpName = "DST.SUP " + $Script:txtGrpLastName.Text + " " + $Script:txtGrpFrstName.Text + " UL Employees Only"
	    $Script:StaffGrpName = "DST.SUP " + $Script:txtGrpLastName.Text + " " + $Script:txtGrpFrstName.Text + " UL Staff"
    }
    If ($Script:chkDSTAllPPL.Checked -eq "Checked")
    {
        $Script:EmpGrpName = "DST.All " + $Script:txtGrpLocName.Text + " People Leaders"
	    $Script:StaffGrpName = "N/A"
    }
    If ($Script:chkDSTSUPPPL.Checked -eq "Checked")
    {
        $Script:EmpGrpName = "DST.SUP " + $Script:txtGrpLastName.Text + " " + $Script:txtGrpFrstName.Text + " People Leaders"
	    $Script:StaffGrpName = "N/A"
    }
    If ($Script:chkDSGAO2.Checked -eq "Checked")
    {
        $Script:EmpGrpName = "DSG.AO2." + $Script:txtGrpLocName.Text + " UL Employees Only"
	    $Script:StaffGrpName = "DSG.AO2." + $Script:txtGrpLocName.Text + " UL Staff"
    }
    If ($Script:chkDSGAO4.Checked -eq "Checked")
    {
        $Script:EmpGrpName = "DSG.AO4." + $Script:txtGrpLocName.Text + " UL Employees Only"
	    $Script:StaffGrpName = "DSG.AO4." + $Script:txtGrpLocName.Text + " UL Staff"
    }

    $Script:txtEmpOnlyGrp.Text = $Script:EmpGrpName
    $Script:txtStaffGrp.Text = $Script:StaffGrpName
    $Script:EmpGrpNameTrunc = $Script:EmpGrpName
    $Script:StaffGrpNameTrunc = $Script:StaffGrpName

    If ($Script:EmpGrpName.Length -gt 64)
    {
        $Script:EmpGrpNameTrunc = $Script:EmpGrpName.Substring(0,64)
    }
    
    If ($Script:StaffGrpName.Length -gt 64)
    {
        $Script:StaffGrpNameTrunc = $Script:StaffGrpName.Substring(0,64)
    } 
}

Function Create-DSTAllGroups
{
    If ($Global:Result -eq "OK")
    {
        Write-Host "Creating Groups...." -ForegroundColor Cyan
        If (($Script:chkDSTAllPPL.Checked -eq "Checked") -or ($Script:chkDSTSUPPPL.Checked -eq "Checked"))
        {
            write-host "Name of the group for People Leaders: " $Script:EmpGrpName
        }
        else
        {
            write-host "Name of the group for UL Employees Only: " $Script:EmpGrpName
            write-host "     Name of the group for all UL Staff: " $Script:StaffGrpName
        }

        $EmpGrpAlias = $Script:EmpGrpName -Replace '[ (),-]',''
        $EmpGrpAlias = $EmpGrpAlias.Replace("UL","")
        $EmpGrpAlias = $EmpGrpAlias.Replace("loyees","")
        $EmpINetAlias = $EmpGrpAlias + "@ul.com"

        $StaffGrpAlias = $Script:StaffGrpName -Replace '[ (),-]',''
        $StaffGrpAlias = $StaffGrpAlias.Replace("UL","")
        $StaffINetAlias = $StaffGrpAlias + "@ul.com"

        $LocName = $Script:txtGrpLocName.Text
        $FrstName = $Script:txtGrpFrstName.Text
        $LastName = $Script:txtGrpLastName.Text

        If ($Script:chkDSTAllGrp.Checked -eq "Checked")
        {
            New-DynamicDistributionGroup -DisplayName $Script:EmpGrpName -Alias $EmpGrpAlias -Name $Script:EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -PrimarySmtpAddress $EmpINetAlias
    	    $EmpGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff that have an employee type of ""Employee"" and are in the ""$LocName"" location.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
            $MailTip = "Sends to individuals listed as Employee assigned to the '" + $LocName + "' site in Oracle."

	        New-DynamicDistributionGroup -DisplayName $StaffGrpName -Alias $StaffGrpAlias -Name $StaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -PrimarySmtpAddress $StaffINetAlias
            $StaffGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff that have an employee type of ""Employee, Consultant, Contractor, Temporary Worker or Freelance Worker"" and are in the ""$LocName"" location.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
            $StaffTip = "Sends to all Employee, Contractor, Consultant, Freelance Contractor or Contingent Worker assigned to the '" + $LocName + "' site in Oracle."
        }

        If ($Script:chkDSTSupGrp.Checked -eq "Checked")
        {
            If ($Script:chkSUPLevel1.Checked -eq "Checked")
            {
                New-DynamicDistributionGroup -DisplayName $Script:EmpGrpName -Alias $EmpGrpAlias -Name $Script:EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute6 $Script:txtGrpEmpNo.Text -PrimarySmtpAddress $EmpINetAlias
                $EmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisor1 in Oracle.  This group contains staff that have an employee type of ""Employee"".  Manual modification of this group is not possible."
                $MailTip = "Sends to individuals listed as Employee that report to '" + $FrstName + " " + $LastName+ "' as Supervisor1 in Oracle."
                
    	        New-DynamicDistributionGroup -DisplayName $Script:StaffGrpName -Alias $StaffGrpAlias -Name $Script:StaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $Script:txtGrpEmpNo.Text -ConditionalCustomAttribute6 $Script:txtGrpEmpNo.Text -PrimarySmtpAddress $StaffINetAlias
                $StaffGrpNote = "The membership of this group is determined at the time the group and contains individuals that report up to " + $FrstName + " " + $Last + " as Supervisor1 in Oracle.  This group contains staff that have an employee type of ""Employee, Consultant, Contractor or Temporary Worker"".  Manual modification of this group is not possible."
                $StaffTip = "Sends to all Employee, Contractor, Consultant, Freelance Contractor or Contingent Worker that report to '"  + $FrstName + " " + $LastName + "' as Supervisor1 in Oracle."
            }
            else
            {
                New-DynamicDistributionGroup -DisplayName $Script:EmpGrpName -Alias $EmpGrpAlias -Name $Script:EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute7 $Script:txtGrpEmpNo.Text -PrimarySmtpAddress $EmpINetAlias
        	    $EmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisro2 in Oracle.  This group contains staff that have an employee type of ""Employee"".  Manual modification of this group is not possible."
                $MailTip = "Sends to individuals listed as Employee that report to '" + $FrstName + " " + $LastName + "' as Supervisor2 in Oracle."

                New-DynamicDistributionGroup -DisplayName $StaffGrpName -Alias $StaffGrpAlias -Name $StaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute7 $Script:txtGrpEmpNo.Text -PrimarySmtpAddress $StaffINetAlias
    	        $StaffGrpNote = "The membership of this group is determined at the time the group and contains individuals that report up to " + $FrstName + " " + $LastName + " as Supervisor2 in Oracle.  This group contains staff that have an employee type of ""Employee, Consultant, Contractor or Temporary Worker"".  Manual modification of this group is not possible."
                $StaffTip = "Sends to all Employee, Contractor, Consultant, Freelance Contractor or Contingent Worker that report to '"  + $FrstName + " " + $LastName + "' as Supervisor2 in Oracle."
            }

            Write-host "Restricting Use of Group to: " $Script:txtGrpEmpNo.Text
            Set-DynamicDistributionGroup $Script:EmpGrpName -AcceptMessagesOnlyFromSendersOrMembers $Script:txtGrpEmpNo.Text
            Set-DynamicDistributionGroup $Script:StaffGrpName -AcceptMessagesOnlyFromSendersOrMembers $Script:txtGrpEmpNo.Text
        }

        If ($Script:chkDSTAllPPL.Checked -eq "Checked")
        {
            New-DynamicDistributionGroup -DisplayName $Script:EmpGrpName -Alias $EmpGrpAlias -Name $Script:EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $EmpINetAlias
    	    $EmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals location in the ""$LocName"" location and is a People Leader in Oracle.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
            $MailTip = "Sends to individuals listed as Employee assigned to the '" + $Script:txtGrpLocName.Text + "' site in Oracle."
        }

        If ($Script:chkDSTSupPPL.Checked -eq "Checked")
        {
            If ($Script:chkSUPLevel1.Checked -eq "Checked")
            {
                New-DynamicDistributionGroup -DisplayName $Script:EmpGrpName -Alias $EmpGrpAlias -Name $Script:EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute6 $Script:txtGrpEmpNo.Text -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $EmpINetAlias
                $EmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisor1 in Oracle.  This group contains staff that have an employee type of ""Employee"".  Manual modification of this group is not possible."
                $MailTip = "Sends to individuals listed as Employee that report to '" + $Script:txtGrpLastName.Text + " " + $Script:txtGrpFrstName.Text + "' as Supervisor1 in Oracle."
            }
            else
            {
                New-DynamicDistributionGroup -DisplayName $Script:EmpGrpName -Alias $EmpGrpAlias -Name $Script:EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute7 $Script:txtGrpEmpNo.Text -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $EmpINetAlias
        	    $EmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " +  $FrstName + " " +  $LastName + " as Supervisro2 in Oracle.  This group contains staff that have an employee type of ""Employee"".  Manual modification of this group is not possible."
                $MailTip = "Sends to individuals listed as Employee that report to '" + $Script:txtGrpLastName.Text + " " + $Script:txtGrpLastName.Text + "' as Supervisor2 in Oracle."
            }
        }
        
        Set-DynamicDistributionGroup $Script:EmpGrpName -Notes $EmpGrpNote -MailTip $MailTip -RequireSenderAuthenticationEnabled $True
        If ($Script:chkDSTAllPPL.Checked -ne "Checked")
        {
            Set-DynamicDistributionGroup $StaffGrpName -Notes $StaffGrpNote -MailTip $StaffTip -RequireSenderAuthenticationEnabled $True
        }
    }
    else
    {
        Write-Host "No Groups Created"
        pause
    }
}

Function Create-DSGAO4Groups
{
    $AO4Groups = Get-AzureADMSGroup -SearchString "DSG.AO4"
    $cnt = 0

    $bute4 = $Script:txtGrpLocName.Text
    $AO4 = $bute4 -replace (",","")
    $EmpGrpNick =  "DSG.AO4." + $Script:txtGrpLocName.Text + "EmpOnly" -replace("[ ]","")
    $EmpGrpDesc = "All UL Employees in Alpha Org 4 " +  $Script:txtGrpLocName.Text
    $Found = ""
    Foreach ($chk in $AO4Groups)
    {
        If ($chk.DisplayName -eq $Script:EmpGrpName)
        {
            $Found = "Yes"
        }
    }
    If ($Found -eq "")
    {
        Write-host "`nCreating" $Script:EmpGrpName -ForegroundColor Green
        New-AzureADMSGroup -Description $EmpGrpDesc -DisplayName $Script:EmpGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $EmpGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute1 -contains ""Employee"") and (user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute4 -eq ""$bute4"")" -MembershipRuleProcessingState "On"
    }
    else
    {
        Write-host "`nGroup" $Script:EmpGrpName "already exists" -ForegroundColor Red
    }

    $StaffGrpNick = "DSG.AO4." + $Script:txtGrpLocName.Text + "Staff" -replace("[ ]","")
    $StaffGrpDesc = "All UL Staff in Alpha Org 4 " +  $Script:txtGrpLocName.Text
    $Found = ""
    Foreach ($chk in $AO4Groups)
    {
        If ($chk.DisplayName -eq $Script:StaffGrpName)
        {
            $Found = "Yes"
        }
    }
    If ($Found -eq "")
    {
        Write-host "Creating" $Script:StaffGrpName -ForegroundColor Green
        New-AzureADMSGroup -Description $StaffGrpDesc -DisplayName $Script:StaffGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $StaffGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute4 -eq ""$bute4"")" -MembershipRuleProcessingState "On"
    }
    else
    {
        Write-host "Group" $Script:StaffGrpName "already exists" -ForegroundColor Red
    }
}

Function Create-DSGAO2Groups
{
    $AO2Groups = Get-AzureADMSGroup -SearchString "DSG.AO2"
    $cnt = 0

    $bute8 = $Script:txtGrpLocName.Text
    $AO2 = $bute8 -replace (",","")
    $EmpGrpNick =  "DSG.AO2." + $Script:txtGrpLocName.Text + "EmpOnly" -replace("[ ]","")
    $EmpGrpDesc = "All UL Employees in Alpha Org 2 " +  $Script:txtGrpLocName.Text
    $Found = ""
    Foreach ($chk in $AO2Groups)
    {
        If ($chk.DisplayName -eq $Script:EmpGrpName)
        {
            $Found = "Yes"
        }
    }
    If ($Found -eq "")
    {
        Write-host "`nCreating" $Script:EmpGrpName -ForegroundColor Green
        New-AzureADMSGroup -Description $EmpGrpDesc -DisplayName $Script:EmpGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $EmpGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute1 -contains ""Employee"") and (user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute8 -eq ""$bute8"")" -MembershipRuleProcessingState "On"
    }
    else
    {
        Write-host "`nGroup" $Script:EmpGrpName "already exists" -ForegroundColor Red
    }

    $StaffGrpNick = "DSG.AO2." + $Script:txtGrpLocName.Text + "Staff" -replace("[ ]","")
    $StaffGrpDesc = "All UL Staff in Alpha Org 2 " +  $Script:txtGrpLocName.Text
    $Found = ""
    Foreach ($chk in $AO2Groups)
    {
        If ($chk.DisplayName -eq $Script:StaffGrpName)
        {
            $Found = "Yes"
        }
    }
    If ($Found -eq "")
    {
        Write-host "Creating" $Script:StaffGrpName -ForegroundColor Green
        New-AzureADMSGroup -Description $StaffGrpDesc -DisplayName $Script:StaffGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $StaffGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute8 -eq ""$bute8"")" -MembershipRuleProcessingState "On"
    }
    else
    {
        Write-host "Group" $Script:StaffGrpName "already exists" -ForegroundColor Red
    }
}

## Start Dynamic Group Creation

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

    If (($Script:chkDSTAllGrp.Checked -eq "Checked") -or ($Script:chkDSTSupGrp.Checked -eq "Checked") -or ($Script:chkDSTAllPPL.Checked -eq "Checked") -or ($Script:chkDSTSupPPL.Checked -eq "Checked"))
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