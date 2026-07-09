<#   
================================================================================ 
 Name: Licenses Form
 ================================================================================ 
 #
 # 11/18/2020 - SAG - Added missing code for unassigning licenses
 # 12/16/2020 - SAG - Added Project Plan 5 licenses
 # 05/08/2021 - SAG - Modified how the employee number details is obtained.
 # 07/06/2021 - SAG - Added Visio Plan2 license
 # 08/03/2021 - SAG - Added Power Automate per User Plan license
 # 07/07/2022 - SAG - Cleaned up script to remove licenses removed as part of the Renewal and to change to the new E5 licensing model
 # 05/15/2024 - SAG - Modified to to use Add_Click features and to optimize code, display license changes and to allow updating of licenses for more than one indivdiual without exiting.
 # 08/22/2025 - SAG - Added CoPilot Licenses top the list of available licenses

 PROJECT_P1
#>  

Function Add-AssignedLicenses
{
#    write-host "Updating Assigned License Information" -ForegroundColor Yellow
    $Global:LicDetails.Items.Clear()
    $Global:Action = "Continue"
    $Global:UserDet = ""
#    $Global:UserLicense = Get-MsolUser -UserPrincipalName $Global:UPN -ErrorAction SilentlyContinue
    $Global:UserLicense = Get-MgUserLicenseDetail -UserId $Global:UPN
    $O365Lic = $Global:UserLicense
    $LDAPFilter = "(userPrincipalName=" + $Global:UPN + ")"
    $Global:ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,proxyaddresses,extensionattribute1,extensionattribute4
    O365HasLicenses
    O365Licenses
    $Script:ButGetENo.visible = $false
    $Script:lblUserInf.Text = "Employee Info: "
    $Global:txtUserInf.Text = $Global:UserDet
    $Global:txtUserInf.ReadOnly = $true

    If ($Global:LicAssigned.Length -ne 0)
    {
        $LocArray = $Global:LicAssigned.split(",")
        $i=0   # Counter 
        foreach ($element in $LocArray) { # Loop through Azure list and add to listbox 
            [void] $Global:LicDetails.Items.Add($element.TrimStart())  # Add element to listbox 
            $i ++ 
        } 
    }
    $Global:form.Controls.Add($Global:LicDetails) #Add listbox to form 
    # Obtain Value with: $Global:LicDetails.SelectedItem
}

Function Get-LicenseUsage
{
    If (($Script:chkE5Lic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## E5 License
        $E5Lic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPREMIUM"}
        $Script:lblE5Act.Text = $E5Lic.PrepaidUnits.Enabled
        $Script:lblE5InU.Text = $E5Lic.ConsumedUnits
        $Script:lblE5Avl.Text = ($E5Lic.PrepaidUnits.Enabled-$E5Lic.ConsumedUnits)
        $Script:chkE5Lic.Checked = $False
        ## EMS Licenses
        $EMSLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "EMSPREMIUM"}
        $Script:chkEMSAct.Text = $EMSLic.PrepaidUnits.Enabled
        $Script:chkEMSInU.Text = $EMSLic.ConsumedUnits
        $Script:chkEMSAvl.Text = ($EMSLic.PrepaidUnits.Enabled-$EMSLic.ConsumedUnits)
    }
    If (($Script:chkP2Lic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## P2 Licenses
        $P2Lic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "EXCHANGEENTERPRISE"}
        $Script:lblP2Act.Text = $P2Lic.PrepaidUnits.Enabled
        $Script:lblP2InU.Text = $P2Lic.ConsumedUnits
        $Script:lblP2Avl.Text = ($P2Lic.PrepaidUnits.Enabled-$P2Lic.ConsumedUnits)
        $Script:chkP2Lic.Checked = $False
    }
    If (($Script:chkPBIFLic.Checked -eq $True) -or ($Script:chkPBIPLic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## PowerBI Free Licenses
        $PBIFLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "POWER_BI_STANDARD"}
        $Script:lblPBIFAct.Text = $PBIFLic.PrepaidUnits.Enabled
        $Script:lblPBIFInU.Text = $PBIFLic.ConsumedUnits
        $Script:lblPBIFAvl.Text = ($PBIFLic.PrepaidUnits.Enabled-$PBIFLic.ConsumedUnits)
        $Script:chkPBIFLic.Checked = $False
        
        ## PowerBI Pro Licenses
        $PBIPLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "POWER_BI_PRO"}
        $Script:lblPBIPAct.Text = $E5Lic.PrepaidUnits.Enabled
#        $PBIPMem = Get-AzureADGroupMember -ObjectId 4246671a-2ac5-4b8a-96fd-4e4b7d817791 -all $true
        $PBIPMem = (Get-MgGroupMember -GroupId 4246671a-2ac5-4b8a-96fd-4e4b7d817791 -all).Id
        $Script:lblPBIPInU.Text = $PBIPMem.count
        $Script:lblPBIPAvl.Text = ($E5Lic.PrepaidUnits.Enabled-$PBIPMem.count)
        $Script:chkPBIPLic.Checked = $False
    }
    If (($Script:chkATPDefLic.Checked -eq $True) -or ($First -eq "Yes"))
    {
       ## ATPDef Licenses
        $ATPDefLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "WIN_DEF_ATP"}
        $Script:lblATPDefAct.Text = $ATPDefLic.PrepaidUnits.Enabled
        $Script:lblATPDefInU.Text = $ATPDefLic.ConsumedUnits
        $Script:lblATPDefAvl.Text = ($ATPDefLic.PrepaidUnits.Enabled-$ATPDefLic.ConsumedUnits)
        $Script:chkATPDefLic.Checked = $False
    }
    If (($Script:chkAudioConfLic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## AudioConf Licenses
        $AudioConfLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "MCOMEETADV"}
        $Script:lblAudioConfAct.Text = $E5Lic.PrepaidUnits.Enabled     # This is a sublicense of the E5 license
#        $AudioMem = Get-AzureADGroupMember -ObjectId a9b06203-849c-40c7-93f9-6c404fdcf6d3 -all $true
        $AudioMem = (Get-MgGroupMember -GroupId a9b06203-849c-40c7-93f9-6c404fdcf6d3 -all).Id
        $Script:lblAudioConfInU.Text = $AudioMem.count
        $Script:lblAudioConfAvl.Text = ($E5Lic.PrepaidUnits.Enabled-$AudioMem.count)  # This is a sublicense of the E5 license
        $Script:chkAudioConfLic.Checked = $False
    }
    If (($Script:chkAutoPUsrLic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## Power Automate p/User Licenses
        $AutoPUsrLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "FLOW_PER_USER"}
        $Script:lblAutoPUsrAct.Text = $AutoPUsrLic.PrepaidUnits.Enabled
        $Script:lblAutoPUsrInU.Text = $AutoPUsrLic.ConsumedUnits
        $Script:lblAutoPUsrAvl.Text = ($AutoPUsrLic.PrepaidUnits.Enabled-$AutoPUsrLic.ConsumedUnits)
        $Script:chkAutoPUsrLic.Checked = $False
    }
    If (($Script:chkAutoPAutoPrem.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## Power Automate Premium Licenses
        $AutoPAutoPrem = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "POWERAUTOMATE_ATTENDED_RPA"}
         $Script:lblAutoPAutoPremAct.Text = $AutoPAutoPrem.PrepaidUnits.Enabled
         $Script:lblAutoPAutoPremInU.Text = $AutoPAutoPrem.ConsumedUnits
         $Script:lblAutoPAutoPremAvl.Text = ($AutoPAutoPrem.PrepaidUnits.Enabled-$AutoPAutoPrem.ConsumedUnits)
         $Script:chkAutoPAutoPrem.Checked = $False
    }
    If (($Script:chkCAPLic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## MCOCAP Licenses
        $CAPLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "MCOCAP"}
        $Script:lblCAPAct.Text = $CAPLic.PrepaidUnits.Enabled
        $Script:lblCAPInU.Text = $CAPLic.ConsumedUnits
        $Script:lblCAPAvl.Text = ($CAPLic.PrepaidUnits.Enabled-$CAPLic.ConsumedUnits)
        $Script:chkCAPLic.Checked = $False
    }
    If (($Script:chkCoPilotLic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## CoPilot Licenses
        $CoPilotLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "Microsoft_365_Copilot"}
        $Script:lblCoPilotAct.Text = $CoPilotLic.PrepaidUnits.Enabled
        $Script:lblCoPilotInU.Text = $CoPilotLic.ConsumedUnits
        $Script:lblCoPilotAvl.Text = ($CoPilotLic.PrepaidUnits.Enabled-$CoPilotLic.ConsumedUnits)
        $Script:chkCoPilotLic.Checked = $False
    }
    If (($Script:chkMeetingLic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## Meeting Room Licenses
        $MeetingLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "MEETING_ROOM"}
        $Script:lblMeetingAct.Text = $MeetingLic.PrepaidUnits.Enabled
        $Script:lblMeetingInU.Text = $MeetingLic.ConsumedUnits
        $Script:lblMeetingAvl.Text = ($MeetingLic.PrepaidUnits.Enabled-$MeetingLic.ConsumedUnits)
        $Script:chkMeetingLic.Checked = $False
    }
    If (($Script:chkAppsPUsrLic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## PowerApps Premium (per User Plan) Licenses
        $AppsPUsrLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "POWERAPPS_PER_USER"}
        $Script:lblAppsPUsrAct.Text = $AppsPUsrLic.PrepaidUnits.Enabled
        $Script:lblAppsPUsrInU.Text = $AppsPUsrLic.ConsumedUnits
        $Script:lblAppsPUsrAvl.Text = ($AppsPUsrLic.PrepaidUnits.Enabled-$AppsPUsrLic.ConsumedUnits)
        $Script:chkAppsPUsrLic.Checked = $False
    }
    If (($Script:chkPhoneSysLic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## PhoneSystem Licenses
        $PhoneSysLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "MCOEV"}
        $Script:lblPhoneSysAct.Text = $PhoneSysLic.PrepaidUnits.Enabled
        $Script:lblPhoneSysInU.Text = $PhoneSysLic.ConsumedUnits
        $Script:lblPhoneSysAvl.Text = ($PhoneSysLic.PrepaidUnits.Enabled-$PhoneSysLic.ConsumedUnits)
        $Script:chkPhoneSysLic.Checked = $False
    }
    If (($Script:chkProjP1Lic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## Project Plan 1 Licenses
        $ProjP1Lic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "PROJECT_P1"}
        $Script:lblProjP1Act.Text = $ProjP1Lic.PrepaidUnits.Enabled
        $Script:lblProjP1InU.Text = $ProjP1Lic.ConsumedUnits
        $Script:lblProjP1Avl.Text = ($ProjP1Lic.PrepaidUnits.Enabled-$ProjP1Lic.ConsumedUnits)
        $Script:chkProjP1Lic.Checked = $False
    }
    If (($Script:chkProjP3Lic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## Project Plan 3 Licenses
        $ProjP3Lic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "PROJECTPROFESSIONAL"}
        $Script:lblProjP3Act.Text = $ProjP3Lic.PrepaidUnits.Enabled
        $Script:lblProjP3InU.Text = $ProjP3Lic.ConsumedUnits
        $Script:lblProjP3Avl.Text = ($ProjP3Lic.PrepaidUnits.Enabled-$ProjP3Lic.ConsumedUnits)
        $Script:chkProjP3Lic.Checked = $False
    }
    If (($Script:chkProjP5Lic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## Project Plan 5 Licenses
        $ProjP5Lic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "PROJECTPREMIUM"}
        $Script:lblProjP5Act.Text = $ProjP5Lic.PrepaidUnits.Enabled
        $Script:lblProjP5InU.Text = $ProjP5Lic.ConsumedUnits
        $Script:lblProjP5Avl.Text = ($ProjP5Lic.PrepaidUnits.Enabled-$ProjP5Lic.ConsumedUnits)
        $Script:chkProjP5Lic.Checked = $False   
    }
    If (($Script:chkTeamsProLic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## Teams Premium for Departments
        $TeamsProLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "Microsoft_Teams_Premium"}
        $Script:lblTeamsProAct.Text = $TeamsProLic.PrepaidUnits.Enabled
        $Script:lblTeamsProInU.Text = $TeamsProLic.ConsumedUnits
        $Script:lblTeamsProAvl.Text = ($TeamsProLic.PrepaidUnits.Enabled-$TeamsProLic.ConsumedUnits)
        $Script:chkTeamsProLic.Checked = $False   
    }
    If (($Script:chkVisioP1Lic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## Visio Online Plan 1 Licenses
        $VisioP1Lic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "VISIOONLINE_PLAN1"}
        $Script:lblVisioP1Act.Text = $VisioP1Lic.PrepaidUnits.Enabled
        $Script:lblVisioP1InU.Text = $VisioP1Lic.ConsumedUnits
        $Script:lblVisioP1Avl.Text = ($VisioP1Lic.PrepaidUnits.Enabled-$VisioP1Lic.ConsumedUnits)
        $Script:chkVisioP1Lic.Checked = $False
    }
    If (($Script:chkVisioP2Lic.Checked -eq $True) -or ($First -eq "Yes"))
    {
        ## Visio Online Plan 2 Licenses
        $VisioP2Lic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "VISIOCLIENT"}
        $Script:lblVisioP2Act.Text = $VisioP2Lic.PrepaidUnits.Enabled
        $Script:lblVisioP2InU.Text = $VisioP2Lic.ConsumedUnits
        $Script:lblVisioP2Avl.Text = ($VisioP2Lic.PrepaidUnits.Enabled-$VisioP2Lic.ConsumedUnits)
        $Script:chkVisioP2Lic.Checked = $False
    }
}

Function Build-LicenseStatsForm
{
    $TopLoc = 20 
    $ActLoc = 30
    $InULoc = 92
    $AvlLoc = 150
    $LicLoc = 210

    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "O365 License Statistics" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(700,670) #(W,H)

    ## Label and TextBox  
    ## Title Line
    $Script:lblTitleLine = New-Object System.Windows.Forms.Label   
        $Script:lblTitleLine.Text = "Active           InUse          Avail           License Name"
        $Script:lblTitleLine.Top = 15 ; $Script:lblTitleLine.Left = 30; $Script:lblTitleLine.Width=120 ;$Script:lblTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblTitleLine)    # Add to Form 

    ## E5 Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkE5Lic = New-Object Windows.Forms.checkbox 
        $Script:chkE5Lic.Left = $LicLoc; $Script:chkE5Lic.Width = 280; $Script:chkE5Lic.Top = ($TopLoc-5)  
        $Script:chkE5Lic.Text = "Enterprise E5 License" 
        $Script:chkE5Lic.Checked = $false   # set a default value 
        $Script:chkE5Lic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkE5Lic) 
    $Script:lblE5Act = New-Object System.Windows.Forms.Label
        $Script:lblE5Act.Top = $TopLoc ; $Script:lblE5Act.Left = $ActLoc; $Script:lblE5Act.Width=10 ;$Script:lblE5Act.AutoSize = $true
        $Global:form.Controls.Add($Script:lblE5Act)    # Add to Form
    $Script:lblE5InU = New-Object System.Windows.Forms.Label
        $Script:lblE5InU.Top = $TopLoc ; $Script:lblE5InU.Left = $InULoc; $Script:lblE5InU.Width=10 ;$Script:lblE5InU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblE5InU)    # Add to Form
    $Script:lblE5Avl = New-Object System.Windows.Forms.Label
        $Script:lblE5Avl.Top = $TopLoc ; $Script:lblE5Avl.Left = $AvlLoc; $Script:lblE5Avl.Width=10 ;$Script:lblE5Avl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblE5Avl)    # Add to Form

    ## AudioConf Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkAudioConfLic = New-Object Windows.Forms.checkbox 
        $Script:chkAudioConfLic.Left = $LicLoc+10; $Script:chkAudioConfLic.Width = 280; $Script:chkAudioConfLic.Top = ($TopLoc-5)  
        $Script:chkAudioConfLic.Text = "E5 Audio Conferencing Feature" 
        $Script:chkAudioConfLic.Checked = $false   # set a default value 
        $Script:chkAudioConfLic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkAudioConfLic)
    $Script:lblAudioConfAct = New-Object System.Windows.Forms.Label
        $Script:lblAudioConfAct.Top = $TopLoc ; $Script:lblAudioConfAct.Left = $ActLoc; $Script:lblAudioConfAct.Width=10 ;$Script:lblAudioConfAct.AutoSize = $true
        $Global:form.Controls.Add($Script:lblAudioConfAct)    # Add to Form
    $Script:lblAudioConfInU = New-Object System.Windows.Forms.Label
        $Script:lblAudioConfInU.Top = $TopLoc ; $Script:lblAudioConfInU.Left = $InULoc; $Script:lblAudioConfInU.Width=10 ;$Script:lblAudioConfInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblAudioConfInU)    # Add to Form
    $Script:lblAudioConfAvl = New-Object System.Windows.Forms.Label
        $Script:lblAudioConfAvl.Top = $TopLoc ; $Script:lblAudioConfAvl.Left = $AvlLoc; $Script:lblAudioConfAvl.Width=10 ;$Script:lblAudioConfAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblAudioConfAvl)    # Add to Form

    ## PowerBI Pro Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkPBIPLic = New-Object Windows.Forms.checkbox 
        $Script:chkPBIPLic.Left = $LicLoc+10; $Script:chkPBIPLic.Width = 280; $Script:chkPBIPLic.Top = ($TopLoc-5)  
        $Script:chkPBIPLic.Text = "E5 PowerBI Pro Featuure" 
        $Script:chkPBIPLic.Checked = $false   # set a default value 
        $Script:chkPBIPLic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkPBIPLic)         
    $Script:lblPBIPAct = New-Object System.Windows.Forms.Label
        $Script:lblPBIPAct.Top = $TopLoc ; $Script:lblPBIPAct.Left = $ActLoc ; $Script:lblPBIPAct.Width=10 ;$Script:lblPBIPAct.AutoSize = $true
        $Global:form.Controls.Add($Script:lblPBIPAct)    # Add to Form
    $Script:lblPBIPInU = New-Object System.Windows.Forms.Label
        $Script:lblPBIPInU.Top = $TopLoc ; $Script:lblPBIPInU.Left = $InULoc; $Script:lblPBIPInU.Width=10 ;$Script:lblPBIPInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblPBIPInU)    # Add to Form
    $Script:lblPBIPAvl = New-Object System.Windows.Forms.Label
        $Script:lblPBIPAvl.Top = $TopLoc ; $Script:lblPBIPAvl.Left = $AvlLoc; $Script:lblPBIPAvl.Width=10 ;$Script:lblPBIPAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblPBIPAvl)    # Add to Form

    ## EMS Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkEMSLic = New-Object Windows.Forms.checkbox 
        $Script:chkEMSLic.Left = $LicLoc; $Script:chkEMSLic.Width = 280; $Script:chkEMSLic.Top = ($TopLoc-5)  
        $Script:chkEMSLic.Text = "Enterprise Mobility + Security E5 License" 
        $Script:chkEMSLic.Checked = $false   # set a default value 
        $Script:chkEMSLic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkEMSLic)
    $Script:chkEMSAct = New-Object System.Windows.Forms.Label
        $Script:chkEMSAct.Top = $TopLoc ; $Script:chkEMSAct.Left = $ActLoc; $Script:chkEMSAct.Width=10 ;$Script:chkEMSAct.AutoSize = $true
        $Global:form.Controls.Add($Script:chkEMSAct)    # Add to Form
    $Script:chkEMSInU = New-Object System.Windows.Forms.Label
        $Script:chkEMSInU.Top = $TopLoc ; $Script:chkEMSInU.Left = $InULoc; $Script:chkEMSInU.Width=10 ;$Script:chkEMSInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:chkEMSInU)    # Add to Form
    $Script:chkEMSAvl = New-Object System.Windows.Forms.Label
        $Script:chkEMSAvl.Top = $TopLoc ; $Script:chkEMSAvl.Left = $AvlLoc; $Script:chkEMSAvl.Width=10 ;$Script:chkEMSAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:chkEMSAvl)    # Add to Form
 
    ## P2 Licenses
    $TopLoc = $TopLoc + 20   
    $Script:chkP2Lic = New-Object Windows.Forms.checkbox 
        $Script:chkP2Lic.Left = $LicLoc; $Script:chkP2Lic.Width = 280; $Script:chkP2Lic.Top = ($TopLoc-5)  
        $Script:chkP2Lic.Text = "Exchange Online P2 License" 
        $Script:chkP2Lic.Checked = $false   # set a default value 
        $Script:chkP2Lic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkP2Lic) 
    $Script:lblP2Act = New-Object System.Windows.Forms.Label
        $Script:lblP2Act.Top = $TopLoc ; $Script:lblP2Act.Left = $ActLoc; $Script:lblP2Act.Width=10 ;$Script:lblP2Act.AutoSize = $true
        $Global:form.Controls.Add($Script:lblP2Act)    # Add to Form
    $Script:lblP2InU = New-Object System.Windows.Forms.Label
        $Script:lblP2InU.Top = $TopLoc ; $Script:lblP2InU.Left = $InULoc; $Script:lblP2InU.Width=10 ;$Script:lblP2InU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblP2InU)    # Add to Form
    $Script:lblP2Avl = New-Object System.Windows.Forms.Label
        $Script:lblP2Avl.Top = $TopLoc ; $Script:lblP2Avl.Left = $AvlLoc; $Script:lblP2Avl.Width=10 ;$Script:lblP2Avl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblP2Avl)    # Add to Form

   ## ATPDef Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkATPDefLic = New-Object Windows.Forms.checkbox 
        $Script:chkATPDefLic.Left = $LicLoc; $Script:chkATPDefLic.Width = 280; $Script:chkATPDefLic.Top = ($TopLoc-5)  
        $Script:chkATPDefLic.Text = "Defender Advanced Theat Protection License" 
        $Script:chkATPDefLic.Checked = $false   # set a default value 
        $Script:chkATPDefLic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkATPDefLic)
    $Script:lblATPDefAct = New-Object System.Windows.Forms.Label
        $Script:lblATPDefAct.Top = $TopLoc ; $Script:lblATPDefAct.Left = $ActLoc; $Script:lblATPDefAct.Width=10 ;$Script:lblATPDefAct.AutoSize = $true
        $Global:form.Controls.Add($Script:lblATPDefAct)    # Add to Form
    $Script:lblATPDefInU = New-Object System.Windows.Forms.Label
        $Script:lblATPDefInU.Top = $TopLoc ; $Script:lblATPDefInU.Left = $InULoc; $Script:lblATPDefInU.Width=10 ;$Script:lblATPDefInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblATPDefInU)    # Add to Form
    $Script:lblATPDefAvl = New-Object System.Windows.Forms.Label
        $Script:lblATPDefAvl.Top = $TopLoc ; $Script:lblATPDefAvl.Left = $AvlLoc; $Script:lblATPDefAvl.Width=10 ;$Script:lblATPDefAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblATPDefAvl)    # Add to Form

    ## CoPilot Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkCoPilotLic = New-Object Windows.Forms.checkbox 
        $Script:chkCoPilotLic.Left = $LicLoc; $Script:chkCoPilotLic.Width = 280; $Script:chkCoPilotLic.Top = ($TopLoc-5)  
        $Script:chkCoPilotLic.Text = "Microsoft CoPilot License" 
        $Script:chkCoPilotLic.Checked = $false   # set a default value 
        $Script:chkCoPilotLic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkCoPilotLic)          
    $Script:lblCoPilotAct = New-Object System.Windows.Forms.Label
        $Script:lblCoPilotAct.Top = $TopLoc ; $Script:lblCoPilotAct.Left = $ActLoc; $Script:lblCoPilotAct.Width=10 ;$Script:lblCoPilotAct.AutoSize = $true
        $Global:form.Controls.Add($Script:lblCoPilotAct)    # Add to Form
    $Script:lblCoPilotInU = New-Object System.Windows.Forms.Label
        $Script:lblCoPilotInU.Top = $TopLoc ; $Script:lblCoPilotInU.Left = $InULoc; $Script:lblCoPilotInU.Width=10 ;$Script:lblCoPilotInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblCoPilotInU)    # Add to Form
    $Script:lblCoPilotAvl = New-Object System.Windows.Forms.Label
        $Script:lblCoPilotAvl.Top = $TopLoc ; $Script:lblCoPilotAvl.Left = $AvlLoc; $Script:lblCoPilotAvl.Width=10 ;$Script:lblCoPilotAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblCoPilotAvl)    # Add to Form

    ## Meeting Room Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkMeetingLic = New-Object Windows.Forms.checkbox 
        $Script:chkMeetingLic.Left = $LicLoc; $Script:chkMeetingLic.Width = 280; $Script:chkMeetingLic.Top = ($TopLoc-5)  
        $Script:chkMeetingLic.Text = "Microsoft Meeting Room License" 
        $Script:chkMeetingLic.Checked = $false   # set a default value 
        $Script:chkMeetingLic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkMeetingLic)          
    $Script:lblMeetingAct = New-Object System.Windows.Forms.Label
        $Script:lblMeetingAct.Top = $TopLoc ; $Script:lblMeetingAct.Left = $ActLoc; $Script:lblMeetingAct.Width=10 ;$Script:lblMeetingAct.AutoSize = $true
        $Global:form.Controls.Add($Script:lblMeetingAct)    # Add to Form
    $Script:lblMeetingInU = New-Object System.Windows.Forms.Label
        $Script:lblMeetingInU.Top = $TopLoc ; $Script:lblMeetingInU.Left = $InULoc; $Script:lblMeetingInU.Width=10 ;$Script:lblMeetingInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblMeetingInU)    # Add to Form
    $Script:lblMeetingAvl = New-Object System.Windows.Forms.Label
        $Script:lblMeetingAvl.Top = $TopLoc ; $Script:lblMeetingAvl.Left = $AvlLoc; $Script:lblMeetingAvl.Width=10 ;$Script:lblMeetingAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblMeetingAvl)    # Add to Form

    ## PhoneSystem Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkPhoneSysLic = New-Object Windows.Forms.checkbox 
        $Script:chkPhoneSysLic.Left = $LicLoc; $Script:chkPhoneSysLic.Width = 280; $Script:chkPhoneSysLic.Top = ($TopLoc-5)  
        $Script:chkPhoneSysLic.Text = "Phone System License" 
        $Script:chkPhoneSysLic.Checked = $false   # set a default value 
        $Script:chkPhoneSysLic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkPhoneSysLic)
    $Script:lblPhoneSysAct = New-Object System.Windows.Forms.Label
        $Script:lblPhoneSysAct.Top = $TopLoc ; $Script:lblPhoneSysAct.Left = $ActLoc; $Script:lblPhoneSysAct.Width=10 ;$Script:lblPhoneSysAct.AutoSize = $true
        $Global:form.Controls.Add($Script:lblPhoneSysAct)    # Add to Form
    $Script:lblPhoneSysInU = New-Object System.Windows.Forms.Label
        $Script:lblPhoneSysInU.Top = $TopLoc ; $Script:lblPhoneSysInU.Left = $InULoc; $Script:lblPhoneSysInU.Width=10 ;$Script:lblPhoneSysInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblPhoneSysInU)    # Add to Form
    $Script:lblPhoneSysAvl = New-Object System.Windows.Forms.Label
        $Script:lblPhoneSysAvl.Top = $TopLoc ; $Script:lblPhoneSysAvl.Left = $AvlLoc; $Script:lblPhoneSysAvl.Width=10 ;$Script:lblPhoneSysAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblPhoneSysAvl)    # Add to Form

    ## PowerApps Premium (per User Plan) Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkAppsPUsrLic = New-Object Windows.Forms.checkbox 
        $Script:chkAppsPUsrLic.Left = $LicLoc; $Script:chkAppsPUsrLic.Width = 280; $Script:chkAppsPUsrLic.Top = ($TopLoc-5)  
        $Script:chkAppsPUsrLic.Text = "PowerApps Premium License" 
        $Script:chkAppsPUsrLic.Checked = $false   # set a default value 
        $Script:chkAppsPUsrLic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkAppsPUsrLic) 
    $Script:lblAppsPUsrAct = New-Object System.Windows.Forms.Label
        $Script:lblAppsPUsrAct.Top = $TopLoc ; $Script:lblAppsPUsrAct.Left = $ActLoc; $Script:lblAppsPUsrAct.Width=10 ;$Script:lblAppsPUsrAct.AutoSize = $true
        $Global:form.Controls.Add($Script:lblAppsPUsrAct)    # Add to Form
    $Script:lblAppsPUsrInU = New-Object System.Windows.Forms.Label
        $Script:lblAppsPUsrInU.Top = $TopLoc ; $Script:lblAppsPUsrInU.Left = $InULoc; $Script:lblAppsPUsrInU.Width=10 ;$Script:lblAppsPUsrInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblAppsPUsrInU)    # Add to Form
    $Script:lblAppsPUsrAvl = New-Object System.Windows.Forms.Label
        $Script:lblAppsPUsrAvl.Top = $TopLoc ; $Script:lblAppsPUsrAvl.Left = $AvlLoc; $Script:lblAppsPUsrAvl.Width=10 ;$Script:lblAppsPUsrAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblAppsPUsrAvl)    # Add to Form

    ## Power Automate per User Plan Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkAutoPUsrLic = New-Object Windows.Forms.checkbox 
        $Script:chkAutoPUsrLic.Left = $LicLoc; $Script:chkAutoPUsrLic.Width = 280; $Script:chkAutoPUsrLic.Top = ($TopLoc-5)  
        $Script:chkAutoPUsrLic.Text = "Power Automate p/User Plan License" 
        $Script:chkAutoPUsrLic.Checked = $false   # set a default value 
        $Script:chkAutoPUsrLic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkAutoPUsrLic) 
    $Script:lblAutoPUsrAct = New-Object System.Windows.Forms.Label
        $Script:lblAutoPUsrAct.Top = $TopLoc ; $Script:lblAutoPUsrAct.Left = $ActLoc; $Script:lblAutoPUsrAct.Width=10 ;$Script:lblAutoPUsrAct.AutoSize = $true
        $Global:form.Controls.Add($Script:lblAutoPUsrAct)    # Add to Form
    $Script:lblAutoPUsrInU = New-Object System.Windows.Forms.Label
        $Script:lblAutoPUsrInU.Top = $TopLoc ; $Script:lblAutoPUsrInU.Left = $InULoc; $Script:lblAutoPUsrInU.Width=10 ;$Script:lblAutoPUsrInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblAutoPUsrInU)    # Add to Form
    $Script:lblAutoPUsrAvl = New-Object System.Windows.Forms.Label
        $Script:lblAutoPUsrAvl.Top = $TopLoc ; $Script:lblAutoPUsrAvl.Left = $AvlLoc; $Script:lblAutoPUsrAvl.Width=10 ;$Script:lblAutoPUsrAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblAutoPUsrAvl)    # Add to Form

    ## Power Automate Premium Licenses (was w/Attended RPA)
    $TopLoc = $TopLoc + 20
    $Script:chkAutoPAutoPrem = New-Object Windows.Forms.checkbox 
        $Script:chkAutoPAutoPrem.Left = $LicLoc; $Script:chkAutoPAutoPrem.Width = 280; $Script:chkAutoPAutoPrem.Top = ($TopLoc-5)
#        $Script:chkAutoPAutoPrem.Text = "Power Automate Premium (was w/AttendedRPA) License"
        $Script:chkAutoPAutoPrem.Text = "Power Automate Premium License" 
        $Script:chkAutoPAutoPrem.Checked = $false   # set a default value 
        $Script:chkAutoPAutoPrem.TabIndex = 2
        $Global:form.Controls.Add($Script:chkAutoPAutoPrem) 
    $Script:lblAutoPAutoPremAct = New-Object System.Windows.Forms.Label
        $Script:lblAutoPAutoPremAct.Top = $TopLoc ; $Script:lblAutoPAutoPremAct.Left = $ActLoc; $Script:lblAutoPAutoPremAct.Width=10 ;$Script:lblAutoPAutoPremAct.AutoSize = $true
        $Global:form.Controls.Add($Script:lblAutoPAutoPremAct)    # Add to Form
    $Script:lblAutoPAutoPremInU = New-Object System.Windows.Forms.Label
        $Script:lblAutoPAutoPremInU.Top = $TopLoc ; $Script:lblAutoPAutoPremInU.Left = $InULoc; $Script:lblAutoPAutoPremInU.Width=10 ;$Script:lblAutoPAutoPremInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblAutoPAutoPremInU)    # Add to Form
    $Script:lblAutoPAutoPremAvl = New-Object System.Windows.Forms.Label
        $Script:lblAutoPAutoPremAvl.Top = $TopLoc ; $Script:lblAutoPAutoPremAvl.Left = $AvlLoc; $Script:lblAutoPAutoPremAvl.Width=10 ;$Script:lblAutoPAutoPremAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblAutoPAutoPremAvl)    # Add to Form

    ## PowerBI Free Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkPBIFLic = New-Object Windows.Forms.checkbox 
        $Script:chkPBIFLic.Left = $LicLoc; $Script:chkPBIFLic.Width = 280; $Script:chkPBIFLic.Top = ($TopLoc-5)  
        $Script:chkPBIFLic.Text = "PowerBI (Free) License" 
        $Script:chkPBIFLic.Checked = $false   # set a default value 
        $Script:chkPBIFLic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkPBIFLic)
    $Script:lblPBIFAct = New-Object System.Windows.Forms.Label
        $Script:lblPBIFAct.Top = $TopLoc ; $Script:lblPBIFAct.Left = $ActLoc; $Script:lblPBIFAct.Width=10 ;$Script:lblPBIFAct.AutoSize = $true
        $Global:form.Controls.Add($Script:lblPBIFAct)    # Add to Form
    $Script:lblPBIFInU = New-Object System.Windows.Forms.Label
        $Script:lblPBIFInU.Top = $TopLoc ; $Script:lblPBIFInU.Left = $InULoc; $Script:lblPBIFInU.Width=10 ;$Script:lblPBIFInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblPBIFInU)    # Add to Form
    $Script:lblPBIFAvl = New-Object System.Windows.Forms.Label
        $Script:lblPBIFAvl.Top = $TopLoc ; $Script:lblPBIFAvl.Left = $AvlLoc; $Script:lblPBIFAvl.Width=10 ;$Script:lblPBIFAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblPBIFAvl)    # Add to Form

$ProjP1Lic
    ## Project Plan 1 Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkProjP1Lic = New-Object Windows.Forms.checkbox 
        $Script:chkProjP1Lic.Left = $LicLoc; $Script:chkProjP1Lic.Width = 280; $Script:chkProjP1Lic.Top = ($TopLoc-5)  
        $Script:chkProjP1Lic.Text = "Planner Plan 1 License" 
        $Script:chkProjP1Lic.Checked = $false   # set a default value 
        $Script:chkProjP1Lic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkProjP1Lic)
        $Script:chkProjP1Lic.Add_Click({
            Show-ActionButtons
            })
        $Script:lblProjP1Act = New-Object System.Windows.Forms.Label
        $Script:lblProjP1Act.Top = $TopLoc ; $Script:lblProjP1Act.Left = $ActLoc; $Script:lblProjP1Act.Width=10 ;$Script:lblProjP1Act.AutoSize = $true
        $Global:form.Controls.Add($Script:lblProjP1Act)    # Add to Form
    $Script:lblProjP1InU = New-Object System.Windows.Forms.Label
        $Script:lblProjP1InU.Top = $TopLoc ; $Script:lblProjP1InU.Left = $InULoc; $Script:lblProjP1InU.Width=10 ;$Script:lblProjP1InU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblProjP1InU)    # Add to Form
    $Script:lblProjP1Avl = New-Object System.Windows.Forms.Label
        $Script:lblProjP1Avl.Top = $TopLoc ; $Script:lblProjP1Avl.Left = $AvlLoc; $Script:lblProjP1Avl.Width=10 ;$Script:lblProjP1Avl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblProjP1Avl)    # Add to Form

    ## Project Plan 3 Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkProjP3Lic = New-Object Windows.Forms.checkbox 
        $Script:chkProjP3Lic.Left = $LicLoc; $Script:chkProjP3Lic.Width = 280; $Script:chkProjP3Lic.Top = ($TopLoc-5)  
        $Script:chkProjP3Lic.Text = "Project Plan 3 License" 
        $Script:chkProjP3Lic.Checked = $false   # set a default value 
        $Script:chkProjP3Lic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkProjP3Lic)
        $Script:chkProjP3Lic.Add_Click({
            Show-ActionButtons
            })
        $Script:lblProjP3Act = New-Object System.Windows.Forms.Label
        $Script:lblProjP3Act.Top = $TopLoc ; $Script:lblProjP3Act.Left = $ActLoc; $Script:lblProjP3Act.Width=10 ;$Script:lblProjP3Act.AutoSize = $true
        $Global:form.Controls.Add($Script:lblProjP3Act)    # Add to Form
    $Script:lblProjP3InU = New-Object System.Windows.Forms.Label
        $Script:lblProjP3InU.Top = $TopLoc ; $Script:lblProjP3InU.Left = $InULoc; $Script:lblProjP3InU.Width=10 ;$Script:lblProjP3InU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblProjP3InU)    # Add to Form
    $Script:lblProjP3Avl = New-Object System.Windows.Forms.Label
        $Script:lblProjP3Avl.Top = $TopLoc ; $Script:lblProjP3Avl.Left = $AvlLoc; $Script:lblProjP3Avl.Width=10 ;$Script:lblProjP3Avl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblProjP3Avl)    # Add to Form

    ## Project Plan 5 Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkProjP5Lic = New-Object Windows.Forms.checkbox 
        $Script:chkProjP5Lic.Left = $LicLoc; $Script:chkProjP5Lic.Width = 280; $Script:chkProjP5Lic.Top = ($TopLoc-5)  
        $Script:chkProjP5Lic.Text = "Project Plan 5 License" 
        $Script:chkProjP5Lic.Checked = $false   # set a default value 
        $Script:chkProjP5Lic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkProjP5Lic)
    $Script:lblProjP5Act = New-Object System.Windows.Forms.Label
        $Script:lblProjP5Act.Top = $TopLoc ; $Script:lblProjP5Act.Left = $ActLoc; $Script:lblProjP5Act.Width=10 ;$Script:lblProjP5Act.AutoSize = $true
        $Global:form.Controls.Add($Script:lblProjP5Act)    # Add to Form
    $Script:lblProjP5InU = New-Object System.Windows.Forms.Label
        $Script:lblProjP5InU.Top = $TopLoc ; $Script:lblProjP5InU.Left = $InULoc; $Script:lblProjP5InU.Width=10 ;$Script:lblProjP5InU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblProjP5InU)    # Add to Form
    $Script:lblProjP5Avl = New-Object System.Windows.Forms.Label
        $Script:lblProjP5Avl.Top = $TopLoc ; $Script:lblProjP5Avl.Left = $AvlLoc; $Script:lblProjP5Avl.Width=10 ;$Script:lblProjP5Avl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblProjP5Avl)    # Add to Form
 
    ## Teams Premium (for Departments) Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkTeamsProLic = New-Object Windows.Forms.checkbox 
        $Script:chkTeamsProLic.Left = $LicLoc; $Script:chkTeamsProLic.Width = 280; $Script:chkTeamsProLic.Top = ($TopLoc-5)  
        $Script:chkTeamsProLic.Text = "Teams Premium" 
        $Script:chkTeamsProLic.Checked = $false   # set a default value 
        $Script:chkTeamsProLic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkTeamsProLic)
    $Script:lblTeamsProAct = New-Object System.Windows.Forms.Label
        $Script:lblTeamsProAct.Top = $TopLoc ; $Script:lblTeamsProAct.Left = $ActLoc; $Script:lblTeamsProAct.Width=10 ;$Script:lblTeamsProAct.AutoSize = $true
        $Global:form.Controls.Add($Script:lblTeamsProAct)    # Add to Form
    $Script:lblTeamsProInU = New-Object System.Windows.Forms.Label
        $Script:lblTeamsProInU.Top = $TopLoc ; $Script:lblTeamsProInU.Left = $InULoc; $Script:lblTeamsProInU.Width=10 ;$Script:lblTeamsProInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblTeamsProInU)    # Add to Form
    $Script:lblTeamsProAvl = New-Object System.Windows.Forms.Label
        $Script:lblTeamsProAvl.Top = $TopLoc ; $Script:lblTeamsProAvl.Left = $AvlLoc; $Script:lblTeamsProAvl.Width=10 ;$Script:lblTeamsProAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblTeamsProAvl)    # Add to Form

    ## MCOCAP Licenses (previously Common Areah Phone)
    $TopLoc = $TopLoc + 20
    $Script:chkCAPLic = New-Object Windows.Forms.checkbox 
        $Script:chkCAPLic.Left = $LicLoc; $Script:chkCAPLic.Width = 280; $Script:chkCAPLic.Top = ($TopLoc-5)  
        $Script:chkCAPLic.Text = "Teams Shared Devices" 
        $Script:chkCAPLic.Checked = $false   # set a default value 
        $Script:chkCAPLic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkCAPLic)
    $Script:lblCAPAct = New-Object System.Windows.Forms.Label
        $Script:lblCAPAct.Top = $TopLoc ; $Script:lblCAPAct.Left = $ActLoc; $Script:lblCAPAct.Width=10 ;$Script:lblCAPAct.AutoSize = $true
        $Global:form.Controls.Add($Script:lblCAPAct)    # Add to Form
    $Script:lblCAPInU = New-Object System.Windows.Forms.Label
        $Script:lblCAPInU.Top = $TopLoc ; $Script:lblCAPInU.Left = $InULoc; $Script:lblCAPInU.Width=10 ;$Script:lblCAPInU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblCAPInU)    # Add to Form
    $Script:lblCAPAvl = New-Object System.Windows.Forms.Label
        $Script:lblCAPAvl.Top = $TopLoc ; $Script:lblCAPAvl.Left = $AvlLoc; $Script:lblCAPAvl.Width=10 ;$Script:lblCAPAvl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblCAPAvl)    # Add to Form

    ## Visio Online Plan 1 Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkVisioP1Lic = New-Object Windows.Forms.checkbox 
        $Script:chkVisioP1Lic.Left = $LicLoc; $Script:chkVisioP1Lic.Width = 280; $Script:chkVisioP1Lic.Top = ($TopLoc-5)  
        $Script:chkVisioP1Lic.Text = "Visio Plan 1 License" 
        $Script:chkVisioP1Lic.Checked = $false   # set a default value 
        $Script:chkVisioP1Lic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkVisioP1Lic)
    $Script:lblVisioP1Act = New-Object System.Windows.Forms.Label
        $Script:lblVisioP1Act.Top = $TopLoc ; $Script:lblVisioP1Act.Left = $ActLoc; $Script:lblVisioP1Act.Width=10 ;$Script:lblVisioP1Act.AutoSize = $true
        $Global:form.Controls.Add($Script:lblVisioP1Act)    # Add to Form
    $Script:lblVisioP1InU = New-Object System.Windows.Forms.Label
        $Script:lblVisioP1InU.Top = $TopLoc ; $Script:lblVisioP1InU.Left = $InULoc; $Script:lblVisioP1InU.Width=10 ;$Script:lblVisioP1InU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblVisioP1InU)    # Add to Form
    $Script:lblVisioP1Avl = New-Object System.Windows.Forms.Label
        $Script:lblVisioP1Avl.Top = $TopLoc ; $Script:lblVisioP1Avl.Left = $AvlLoc; $Script:lblVisioP1Avl.Width=10 ;$Script:lblVisioP1Avl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblVisioP1Avl)    # Add to Form

    ## Visio Online Plan 2 Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkVisioP2Lic = New-Object Windows.Forms.checkbox 
        $Script:chkVisioP2Lic.Left = $LicLoc; $Script:chkVisioP2Lic.Width = 280; $Script:chkVisioP2Lic.Top = ($TopLoc-5)  
        $Script:chkVisioP2Lic.Text = "Visio Plan 2 License" 
        $Script:chkVisioP2Lic.Checked = $false   # set a default value 
        $Script:chkVisioP2Lic.TabIndex = 2
        $Global:form.Controls.Add($Script:chkVisioP2Lic)
    $Script:lblVisioP2Act = New-Object System.Windows.Forms.Label
        $Script:lblVisioP2Act.Top = $TopLoc ; $Script:lblVisioP2Act.Left = $ActLoc; $Script:lblVisioP2Act.Width=10 ;$Script:lblVisioP2Act.AutoSize = $true
        $Global:form.Controls.Add($Script:lblVisioP2Act)    # Add to Form
    $Script:lblVisioP2InU = New-Object System.Windows.Forms.Label
        $Script:lblVisioP2InU.Top = $TopLoc ; $Script:lblVisioP2InU.Left = $InULoc; $Script:lblVisioP2InU.Width=10 ;$Script:lblVisioP2InU.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblVisioP2InU)    # Add to Form
    $Script:lblVisioP2Avl = New-Object System.Windows.Forms.Label
        $Script:lblVisioP2Avl.Top = $TopLoc ; $Script:lblVisioP2Avl.Left = $AvlLoc; $Script:lblVisioP2Avl.Width=10 ;$Script:lblVisioP2Avl.AutoSize = $true
        $Global:form.Controls.Add($Script:lblVisioP2Avl)    # Add to Form

    $TopLoc = $TopLoc + 40
    ## User Details
    $Script:lblUserInf = New-Object System.Windows.Forms.Label
        $Script:lblUserInf.Text = "Employee Number:"
        $Script:lblUserInf.Top = $TopLoc; $Script:lblUserInf.Left=10; $Script:lblUserInf.Width=10 ;$Script:lblUserInf.AutoSize = $true
        $Global:form.Controls.Add($Script:lblUserInf)    # Add to Form
    $Global:txtUserInf = New-Object System.Windows.Forms.TextBox
        $Global:txtUserInf.Top = $TopLoc; $Global:txtUserInf.Left=120; $Global:txtUserInf.Width=100 ;$Global:txtUserInf.AutoSize = $true
        $Global:form.Controls.Add($Global:txtUserInf)    # Add to Form
        $Global:InputFocus = $Global:txtUserInf
        $Global:txtUserInf.Add_Click({
            $Script:ButGetENo.Visible = $True
            $Script:lblUserInf.Text = "Employee Number:"
            $Global:txtUserInf.Text = ""
            $Script:lblUserInf.Visible = $True
            $Global:txtUserInf.ReadOnly = $False
            $Script:lblLicDet.Visible = $False
            $Global:LicDetails.Visible = $False
            $Global:LicDetails.Items.Clear()
            $Global:txtUserInf.Left=120; $Global:txtUserInf.Width=100
            $Script:chkAssign.Visible = $False
            $Script:chkStd.Visible = $False
            $Script:chkRemove.Visible = $False
            $Script:chkReview.Visible = $False
            $Script:chkReset.Visible = $False
            $Script:ButNextEmp.Visible = $False
            $Global:OKButton.Visible = $False
         })

    $Script:ButGetENo = New-Object Windows.Forms.Button
        $Script:ButGetENo.Location = New-object System.Drawing.Size(230,$TopLoc)
        $Script:ButGetENo.Size = new-Object System.Drawing.Size(150,20)
        $Script:ButGetENo.Text = "Get Employee Details"
        $Global:form.Controls.Add($Script:ButGetENo)
        $Script:ButGetENo.Add_Click({
            $Global:UPN = $Global:txtUserInf.Text + "@global.ul.com"
#            $Script:Exists = [bool](get-MSOLUser -UserPrincipalName $Global:UPN -ErrorAction SilentlyContinue)
            $Script:Exists = [bool](get-MgUser -UserId $Global:UPN -ErrorAction SilentlyContinue)
            If ($Script:Exists -eq $True)
            {
                $Script:ENo = $Global:txtUserInf.Text
                $EmpNo = $Script:ENo
                Add-AssignedLicenses
                $Script:lblUserInf.Text = "Employee Info:"
                $Global:txtUserInf.Left=120; $Global:txtUserInf.Width=370
                $Script:lblLicDet.Visible = $True
                $Global:LicDetails.Visible = $True
                $Script:chkAssign.Visible = $True
                Show-ActionButtons
            }
            else
            {
                $Output = $wshell.Popup("Invalid employee number.",0,"Invalid Employee Number",0+32)
                $Global:txtUserInf.Text = ""
            }
        })

    $Script:ButNextEmp = New-Object Windows.Forms.Button
        $Script:ButNextEmp.Location = New-object System.Drawing.Size(500,$TopLoc)
        $Script:ButNextEmp.Size = new-Object System.Drawing.Size(110,20)
        $Script:ButNextEmp.Text = "Next Employee"
        $Script:ButNextEmp.Visible = $False
        $Global:form.Controls.Add($Script:ButNextEmp)
        $Script:ButNextEmp.Add_Click({
            $Script:ButNextEmp.Visible = $False
            $Script:ButGetENo.Visible = $True
            $Script:lblUserInf.Text = "Employee Number:"
            $Global:txtUserInf.Text = ""
            $Script:lblUserInf.Visible = $True
            $Global:txtUserInf.ReadOnly = $False
            $Script:lblLicDet.Visible = $False
            $Global:LicDetails.Visible = $False
            $Global:LicDetails.Items.Clear()
            $Global:txtUserInf.Left=120; $Global:txtUserInf.Width=100
            $Script:chkAssign.Visible = $False
            $Script:chkStd.Visible = $False
            $Script:chkRemove.Visible = $False
            $Script:chkReview.Visible = $False
            $Script:chkReset.Visible = $False
            $Global:OKButton.Visible = $False
        })

    $TopLoc = $TopLoc + 30
 ## ListBox - Fill with License Details
    $Script:lblLicDet = New-Object System.Windows.Forms.Label   
        $Script:lblLicDet.Text = "License Enabled:"; $Script:lblLicDet.Top = $TopLoc; $Script:lblLicDet.Left = 10; $Script:lblLicDet.Autosize = $true
        $Script:lblLicDet.Visible = $False  
        $Global:form.Controls.Add($Script:lblLicDet)  
    $Global:LicDetails = New-Object System.Windows.Forms.ListBox  
        $Global:LicDetails.Top = $TopLoc; $Global:LicDetails.Left = 120; $Global:LicDetails.Height = 130; $Global:LicDetails.Width = 370;
        $Global:LicDetails.Visible = $False
        $Global:LicDetails.TabIndex = 1
        $Global:form.Controls.Add($Global:LicDetails) #Add listbox to form 
}

Function Add-ActionButtons
{
   #Action Side CheckBoxes

    ## Assign a License         
    $Script:chkAssign = New-Object Windows.Forms.RadioButton
        $Script:chkAssign.Left = 530; $Script:chkAssign.Width = 200; $Script:chkAssign.Top = 100  
        $Script:chkAssign.Text = "Assign" 
        $Script:chkAssign.Checked = $Script:chkAssign.Checked   # set a default value 
        $Script:chkAssign.Visible = $False
        $Global:form.Controls.Add($Script:chkAssign)
        $Script:chkAssign.Add_Click({
            $Global:OKButton.Visible = $True
        })

    ## Assign Standard License Set
    $Script:chkStd = New-Object Windows.Forms.RadioButton 
        $Script:chkStd.Left = 530; $Script:chkStd.Width = 200; $Script:chkStd.Top = 125
        $Script:chkStd.Text = "Assign Standard Set"
        $Script:chkStd.Checked = $Script:chkStd.Checked   # set a default value 
        $Script:chkStd.Visible = $False
        $Global:form.Controls.Add($Script:chkStd) 
        $Script:chkStd.Add_Click({
            $Global:OKButton.Visible = $True
        })

    ## Remove a License         
    $Script:chkRemove = New-Object Windows.Forms.RadioButton
        $Script:chkRemove.Left = 530; $Script:chkRemove.Width = 200; $Script:chkRemove.Top = 150
        $Script:chkRemove.Text = "Remove" 
        $Script:chkRemove.Checked = $Script:chkRemove.Checked   # set a default value 
        $Script:chkRemove.Visible = $False
        $Global:form.Controls.Add($Script:chkRemove)
        $Script:chkRemove.Add_Click({
            $Global:OKButton.Visible = $True
        }) 

    ## Review SubLicense Assignment
    $Script:chkReview = New-Object Windows.Forms.RadioButton
        $Script:chkReview.Left = 530; $Script:chkReview.Width = 200; $Script:chkReview.Top = 175
        $Script:chkReview.Text = "Review Options" 
        $Script:chkReview.Checked = $Script:chkReview.Checked   # set a default value 
        $Script:chkReview.Visible = $False
        $Global:form.Controls.Add($Script:chkReview)
        $Script:chkReview.Add_Click({
            $Global:OKButton.Visible = $True
        }) 

    ## Reset License Assignment
    $Script:chkReset = New-Object Windows.Forms.RadioButton
        $Script:chkReset.Left =530; $Script:chkReset.Width = 200; $Script:chkReset.Top = 200  
        $Script:chkReset.Text = "Reset Options" 
        $Script:chkReset.Checked = $Script:chkReset.Checked   # set a default value 
        $Script:chkReset.Visible = $False
        $Global:form.Controls.Add($Script:chkReset)
        $script:chkReset.Add_Click({
            $Global:OKButton.Visible = $True
        }) 
}

Function Remove-E5Grp
{
#    $usr = get-MgUser -UserId $Global:UPN  #This line may not be needed... commenting it out 04/03/2025
    $ErrorActionPreference = "SilentlyContinue"
    Remove-AzureADGroupMember -ObjectId 40dff561-964a-4f7b-9794-dba3542b63c4 -MemberID $Script:MgUsr.Id
    If ($? -eq $true)
    {
     	write-host "Removing " $Script:ENo "from Non-Employee Enabled Features Group" -ForegroundColor Yellow
    }
    Remove-AzureADGroupMember -ObjectId 2aab81de-e54a-4c58-909c-b15273c54dff -MemberID $Script:MgUsr.Id
    If ($? -eq $true)
    {
     	write-host "Removing " $Script:ENo "from Employee Enabled Features Group" -ForegroundColor Yellow
    }
    Remove-AzureADGroupMember -ObjectId 192175db-72ee-4f85-a24f-c64d4681bfd8 -MemberID $Script:MgUsr.Id
    If ($? -eq $true)
    {
     	write-host "Removing " $Script:ENo "from WNS User Enabled Features Group" -ForegroundColor Yellow
    }
    $ErrorActionPreference = "Continue"
    ## E5 License
    $E5Lic = Get-MgSubscribedSku|Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPREMIUM"}
    $Script:lblE5Act.Text = $E5Lic.PrepaidUnits.Enabled
    $Script:lblE5InU.Text = $E5Lic.ConsumedUnits
    $Script:lblE5Avl.Text = ($E5Lic.PrepaidUnits.Enabled-$E5Lic.ConsumedUnits)
    $Script:chkE5Lic.Checked = $False
    ## EMS Licenses
    $EMSLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "EMSPREMIUM"}
    $Script:chkEMSAct.Text = $EMSLic.PrepaidUnits.Enabled
    $Script:chkEMSInU.Text = $EMSLic.ConsumedUnits
    $Script:chkEMSAvl.Text = ($EMSLic.PrepaidUnits.Enabled-$EMSLic.ConsumedUnits)
}

Function Add-CoPilotGrp
{
    $Transcription = Read-Host "Enter (1) for CoPilot w/Transcription; (2) CoPilot w/o transcription"
    switch ($Transcription)
    {
        1
        {
            New-MgGroupMember -GroupId 11c2475a-7daa-4962-b489-9d4bfccc7597 -DirectoryObjectId $Script:MgUsr.Id
            write-host "Adding " $Script:ENo "to Microsoft CoPilot w/Transcription License Group" -ForegroundColor Yellow
        }
        2
        {
            New-MgGroupMember -GroupId 6047997a-6204-48e1-b5c3-96f5f39637ec -DirectoryObjectId $Script:MgUsr.Id
            write-host "Adding " $Script:ENo "to Microsoft CoPilot w/o Transcription License Group" -ForegroundColor Yellow
        }
    }
    ## CoPilot License
    $CoPilotLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "Microsoft_365_Copilot"}
    $Script:lblCoPilotAct.Text = $CoPilotLic.PrepaidUnits.Enabled
    $Script:lblCoPilotInU.Text = $CoPilotLic.ConsumedUnits
    $Script:lblCoPilotAvl.Text = ($CoPilotLic.PrepaidUnits.Enabled-$CoPilotLic.ConsumedUnits)
    $Script:chkCoPilotLic.Checked = $False
}

Function Remove-CoPilotGrp
{
    $ErrorActionPreference = "SilentlyContinue"
    Remove-AzureADGroupMember -ObjectId 6047997a-6204-48e1-b5c3-96f5f39637ec -MemberID $Script:MgUsr.Id #w/o Transcription
    If ($? -eq $true)
    {
     	write-host "Removing " $Script:ENo "from Microsoft CoPilot (w/o Transcription) License Group" -ForegroundColor Yellow
    }
    Remove-AzureADGroupMember -ObjectId 11c2475a-7daa-4962-b489-9d4bfccc7597 -MemberID $Script:MgUsr.Id #w/Transcription
    If ($? -eq $true)
    {
     	write-host "Removing " $Script:ENo "from Microsoft CoPilot w/Transcription License Group" -ForegroundColor Yellow
    }
    $ErrorActionPreference = "Continue"
    ## CoPilot License
    $CoPilotLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "Microsoft_365_Copilot"}
    $Script:lblCoPilotAct.Text = $CoPilotLic.PrepaidUnits.Enabled
    $Script:lblCoPilotInU.Text = $CoPilotLic.ConsumedUnits
    $Script:lblCoPilotAvl.Text = ($CoPilotLic.PrepaidUnits.Enabled-$CoPilotLic.ConsumedUnits)
    $Script:chkCoPilotLic.Checked = $False
}

Function Show-ActionButtons
{
    $Script:chkAssign.Visible = $True
    $Script:chkStd.Visible = $True
    $Script:chkRemove.Visible = $True
    $Script:chkReview.Visible = $True
    $Script:chkReset.Visible = $True
    $Script:chkAssign.Checked = $False
    $Script:chkStd.Checked = $False
    $Script:chkRemove.Checked = $False
    $Script:chkReview.Checked = $False
    $Script:chkReset.Checked = $False
}

Function Reset-Form
{
    #Reset and Hide Action Boxes
    $Script:chkAssign.Visible = $False
    $Script:chkStd.Visible = $False
    $Script:chkRemove.Visible = $False
    $Script:chkReview.Visible = $False
    $Script:chkReset.Visible = $False
    $Script:chkAssign.Checked = $False
    $Script:chkStd.Checked = $False
    $Script:chkRemove.Checked = $False
    $Script:chkReview.Checked = $False
    $Script:chkReset.Checked = $False

    $Script:ButNextEmp.Visible = $True
    $Global:OKButton.Visible = $False
}


Add-Type -Assembly System.Windows.Forms     ## Load the Windows Forms assembly 
## Create the main form

$Year           = (get-Date).Year
$LogDirectory   = "E:\Automation\Licensing\Log"
$LogFile		= $LogDirectory + "\" + "Log-Licensing-" + $Year + ".log"
$wshell = New-Object -ComObject Wscript.Shell
invoke-expression -Command E:\O365AdminShared\Scripts\O365DisabledLicenseFeatures.ps1
$Global:Result = ""

write-host "`nBuilding Menu and Gathering Licensing Details...." -ForegroundColor Cyan

Build-LicenseStatsForm
Add-ActionButtons
$First = "Yes"
Get-LicenseUsage
$First = "No"
Add-FormStandardButtons
$Global:OKButton.Visible = $False
$Global:form.Add_Shown( { $form.Activate(); $Global:txtUserInf.Focus()} )
$Global:result = $Global:form.ShowDialog()          ## Show the form, and wait for the response

Do
{
    If ($Global:Result -eq "OK")
    {
        $MgEmpNo = $Script:ENo + "@"
        $Script:MgUsr = Get-MgUser -Filter "startsWith(UserPrincipalName, '$Global:UPN')"
            
        Write-host "`nProcessing for: " $Script:ENo
        $E5Lic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPREMIUM"}
    
        If ($Script:chkAssign.Checked -eq $True)
        {
       # This section handles licenses where one license needs to be removed before assigning another license
            If ($Script:chkE5Lic.Checked -eq $True)
            {
                Assign-E5Lic
            }
            If ($Script:chkP2Lic.Checked -eq $True)
            {
                Assign-P2Lic
            }
            If ($Script:chkPBIFLic.Checked -eq $True)
            {
                $HasLicense = $Global:HasBIFree
                $LicName = "PowerBI Free"
                $LicSKU = "POWER_BI_STANDARD"
#                If ($WNSUsrGrp -eq $False)
#                {
                    Assign-License
#                }
#                else
#                {
#                    $Output = $wshell.Popup("The Microsoft Fabric Free (PowerBI) license is not approved for WNS Staff",0+32)
#                    $LineToWrite = "REVI" + "`t" + "Microsoft Fabric Free (PowerBI) license is not approved for WNS Staff" + $Global:UPN
#                    WriteLogEvent
#                }
                ## PowerBI Free Licenses
            }
            If ($Script:chkPBIPLic.Checked -eq $True)
            {
                $LicName = "PowerBI Pro"
                $LicSKU = "POWER_BI_PRO"
                Assign-PowerBIProLic
            }
            #This section handles licenses where a license can just be assigned without impacting other licenses that might be assigned
            If ($Script:chkATPDefLic.Checked -eq $True)
            {
                $HasLicense = $Global:HasATPDef
                $LicName = "Windows Defender Advanced Threat Protection"
                $LicSKU = "WIN_DEF_ATP"
                Assign-License
            }
            If ($Script:chkAudioConfLic.Checked -eq $True)
            {
                $HasLicense = $Global:HasAudioConf
                $LicName = "Audio Conferencing"
                $LicSKU = "MCOMEETADV"
                Assign-License
                $AudioConfLic = Get-MgSubscribedSku |Where-Object {$_.SkuPartNumber -eq "MCOMEETADV"}
#                $AudioMem = Get-AzureADGroupMember -ObjectId a9b06203-849c-40c7-93f9-6c404fdcf6d3 -all $true
                $AudioMem = (Get-MgGroupMember -GroupId a9b06203-849c-40c7-93f9-6c404fdcf6d3 -all).Id

#                $Script:lblAudioConfAct.Text = $E5Lic.ActiveUnits
#                $Script:lblAudioConfInU.Text = $AudioMem.count
#                $Script:lblAudioConfAvl.Text = ($E5Lic.ActiveUnits-$AudioMem.count)
#                $Script:chkAudioConfLic.Checked = $False
            }
            If ($Script:chkAutoPUsrLic.Checked -eq $True)
            {
                $HasLicense = $Global:HasAutoPUsr
                $LicName = "Power Automate Premium"
                $LicSKU = "FLOW_PER_USER"
                Assign-License
            }
            If ($Script:chkAutoPAutoPrem.Checked -eq $True)
            {
                $HasLicense = $Global:HasAutoPUsrRPA
                $LicName = "Power Automate Premium (Previously w/Attended RPA)"
                $LicSKU = "POWERAUTOMATE_ATTENDED_RPA"
                Assign-License
            }
            If ($Script:chkCAPLic.Checked -eq $True)
            {
                $HasLicense = $Global:HasCAP
                $LicName = "Common Area Phone"
                $LicSKU = "MCOCAP"
                Assign-License
            }
            If ($Script:chkCoPilotLic.Checked -eq $True)
            {
                $HasLicense = $Global:HasCoPilot
                $LicName = "Microsoft CoPilot"
                $LicSKU = "Microsoft_365_Copilot"
                Add-CoPilotGrp
            }
            If ($Script:chkMeetingLic.Checked -eq $True)
            {
                $HasLicense = $Global:HasMeeting
                $LicName = "Microsoft Meeting Room"
                $LicSKU = "MEETING_ROOM"
                Assign-License
            }
            If ($Script:chkAppsPUsrLic.Checked -eq $True)
            {
                $HasLicense = $Global:HasAppsPUsr
                $LicName = "PowerApps p/User"
                $LicSKU = "POWERAPPS_PER_USER"
                Assign-License
            }
            If ($Script:chkPhoneSysLic.Checked -eq $True)
            {
                $HasLicense = $Global:HasPhone
                $LicName = "Phone System"
                $LicSKU = "MCOEV"
                Assign-License
            }
            If ($Script:chkProjP1Lic.Checked -eq $True)
            {
                $HasLicense = $Global:HasProjP1
                $LicName = "Planner Plan 1"
                $LicSKU = "PROJECT_P1"
                Assign-License
            }
            If ($Script:chkProjP3Lic.Checked -eq $True)
            {
                $HasLicense = $Global:HasProjP3
                $LicName = "Project Online Plan 3"
                $LicSKU = "PROJECTPROFESSIONAL"
                Assign-License
            }
            If ($Script:chkProjP5Lic.Checked -eq $True)
            {
                $HasLicense = $Global:HasProjP5
                $LicName = "Project Online Plan 5"
                $LicSKU = "PROJECTPREMIUM"
                Assign-License
            }
            If ($Script:chkTeamsProLic.Checked -eq $True)
            {
                $HasLicense = $Global:HasTeamsPro
                $LicName = "Teams Premium"
                $LicSKU = "Microsoft_Teams_Premium"
#                $LicSKU = "Teams_Premium_(for_Departments)"
                Assign-License
            }
            If ($Script:chkVisioP1Lic.Checked -eq $True)
            {
                $HasLicense = $Global:HasVisioP1
                $LicName = "Visio Online Plan 1"
                $LicSKU = "VISIOONLINE_PLAN1"
                Assign-License
            }
            If ($Script:chkVisioP2Lic.Checked -eq $True)
            {
                $HasLicense = $Global:HasVisioP2
                $LicName = "Visio Online Plan 2"
                $LicSKU = "VISIOCLIENT"
                Assign-License
            }
            $Script:chkAssign.Checked = $false
        }

        If ($Script:chkRemove.Checked -eq $True)
        {
            If (($Script:chkATPDefLic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Windows Defender Advanced Threat Protection"))
            {
                $HasLicense = $Global:HasATPDef
                $LicName = "Windows Defender Advanced Threat Protection"
                $LicSKU = "WIN_DEF_ATP"
                UnAssign-License
            }
            If (($Script:chkAudioConfLic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Audio Conferencing"))
            {
                $HasLicense = $Global:HasAudioConf
                $LicName = "Audio Conferencing"
                $LicSKU = "MCOMEETADV"
                UnAssign-License
            }
            If (($Script:chkAutoPUsrLic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Power Automate p/User Plan"))
            {
                $HasLicense = $Global:HasAutoPUsr
                $LicName = "Power Automate p/User Plan"
                $LicSKU = "FLOW_PER_USER"
                UnAssign-License
            }
            If (($Script:chkAutoPAutoPrem.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Power Automate Premium"))
            {
                $HasLicense = $Global:HasAutoPUsrRPA
                $LicName = "Power Automate Premium (previously w/Attended RPA)"
                $LicSKU = "POWERAUTOMATE_ATTENDED_RPA"
                UnAssign-License
            }
            If (($Script:chkCAPLic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Common Area Phone"))
            {
                $HasLicense = $Global:HasCAP
                $LicName = "Common Area Phone"
                $LicSKU = "MCOCAP"
                UnAssign-License
            }
            If (($Script:chkE5Lic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Enterprise E5"))
            {
#               $HasLicense = $Global:HasE5
#               $LicName = "Enterpris E5"
#               $LicSKU = "ENTERPRISEPREMIUM"
#               UnAssign-License
                Remove-E5Grp
            }
            If (($Script:chkEMSLic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Enterprise Mobility + Security E5"))
            {
                $HasLicense = $Global:HasEMS
                $LicName = "Enterprise Mobility + Security E5"
                $LicSKU = "EMSPREMIUM"
                UnAssign-License
            }
            If (($Script:chkCoPilotLic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Microsoft CoPilot"))
            {
                $HasLicense = $Global:HasCoPilot
                $LicName = "Microsoft CoPilot"
                $LicSKU = "Microsoft_365_Copilot"
                Remove-CoPilotGrp
                UnAssign-License
            }
            If (($Script:chkMeetingLic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Microsoft Meeting Room"))
            {
                $HasLicense = $Global:HasMeeting
                $LicName = "Microsoft Meeting Room"
                $LicSKU = "MEETING_ROOM"
                UnAssign-License
            }
            If (($Script:chkP2Lic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Exchange Online Plan2"))
            {
                $HasLicense = $Global:HasExP2
                $LicName = "Exchange Online Plan2"
                $LicSKU = "EXCHANGEENTERPRISE"
                UnAssign-License
            }
            If ($Script:chkAppsPUsrLic.Checked -eq $True)
            {
                $HasLicense = $Global:HasAppsPUsr
                $LicName = "PowerApps p/User"
                $LicSKU = "POWERAPPS_PER_USER"
                UnAssign-License
            }
            If (($Script:chkPBIFLic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "PowerBI (Free)"))
            {
                If ($Global:HasBIPro -eq $True)
                {
                    $LineToWrite = "REVI" + "`t" + "PowerBI Pro license assigned to this account removing PowerBI (Free) License" + $Global:UPN
    #                Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "POWER_BI_STANDARD"
                    $HasLicense = $Global:HasBIFree
                    $LicName = "PowerBI Free"
                    $LicSKU = "POWER_BI_STANDARD"
                    UnAssign-License
                }
                else
                {
  	                $Output = $wshell.Popup("The PowerBI (Free) license is a stanard license for all staff and cannot be removed.",0,"License Not Assigned",0+32)
                    $LineToWrite = "REVI" + "`t" + "PowerBI (Free) License is a standard license and cannot be removed " + $Global:UPN
                }
                WriteLogEvent
            }
            If (($Script:chkPBIPLic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "PowerBI Pro"))
            {
                $HasLicense = $Global:HasBIPro
                $LicName = "PowerBI Pro"
                $LicSKU = "POWER_BI_PRO"
#                $Script:MgUsr = Get-MgUser -Filter "startsWith(UserPrincipalName, '$EmpNo')"
                UnAssign-PowerBIProLic
<#                if ($Global:HasBIFree -ne "True")
                {
      	            $LineToWrite = "REVI" + "`t" + "Assigning PowerBI Free License to " + $Global:UPN
                    WriteLogEvent
                    $EmpNo = $Script:ENo
                    $Script:MgUsr = Get-MgUser -Filter "startsWith(UserPrincipalName, '$MgEmpNo')"
                    $Script:MgUsr = Get-MgUser -Filter "startsWith(UserPrincipalName, '$EmpNo')"
                    Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "POWER_BI_STANDARD"
                }
#>
            }
            If (($Script:chkPhoneSysLic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Phone System"))
            {
                $HasLicense = $Global:HasPhone
                $LicName = "Phone System"
                $LicSKU = "MCOEV"
                UnAssign-License
            }
            If (($Script:chkProjP1Lic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Planner Plan 1"))
            {
                $HasLicense = $Global:HasProjP1
                $LicName = "Planner Plan 1"
                $LicSKU = "PROJECT_P1"
                UnAssign-License
            }
            If (($Script:chkProjP3Lic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Project Plan 3"))
            {
                $HasLicense = $Global:HasProjP3
                $LicName = "Project Online Plan 3"
                $LicSKU = "PROJECTPROFESSIONAL"
                UnAssign-License
            }
            If (($Script:chkProjP5Lic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Project Online Plan 5"))
            {
                $HasLicense = $Global:HasProjP5
                $LicName = "Project Online Plan 5"
                $LicSKU = "PROJECTPREMIUM"
                UnAssign-License
            }
            If (($Script:chkTeamsProLic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Project Online Plan 5"))
            {
                $HasLicense = $Global:HasTeamsPro
                $LicName = "Teams Pro"
                $LicSKU = "PROJECTPREMIUM"
                UnAssign-License
            }
            If (($Script:chkVisioP1Lic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Visio Online Plan 1"))
            {
                $HasLicense = $Global:HasVisioP1
                $LicName = "Visio Online Plan 1"
                $LicSKU = "VISIOONLINE_PLAN1"
                UnAssign-License
            }
            If (($Script:chkVisioP2Lic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Visio Online Plan 2"))
            {
                $HasLicense = $Global:HasVisioP2
                $HasLicense = $True
#                $Script:MgUsr = Get-MgUser -Filter "startsWith(UserPrincipalName, '$Global:UPN')"
                $LicName = "Visio Online Plan 2"
                $LicSKU = "VISIOCLIENT"
                UnAssign-License
            }
            $Script:chkRemove.Checked = $false
        }
    
        If ($Script:chkStd.Checked -eq $True)
        {
            $usr = get-MgUser -UserId $Global:UPN
#            $usr = get-msoluser -UserPrincipalName $Global:UPN
#            $EmpGrp = [bool](Get-AzureADGroupMember -ObjectId 2aab81de-e54a-4c58-909c-b15273c54dff -all $true |Where-Object {$_.ObjectID -eq $usr.ObjectID})
            $EmpGrp = [bool](Get-MgGroupMember -GroupId 2aab81de-e54a-4c58-909c-b15273c54dff -all $true |Where-Object {$_.Id -eq $usr.ObjectID})
#            $NonEmpGrp = [bool](Get-AzureADGroupMember -ObjectId 40dff561-964a-4f7b-9794-dba3542b63c4 -all $true |Where-Object {$_.ObjectID -eq $usr.ObjectID})
            $NonEmpGrp = [bool](Get-MgGroupMember -GroupId 40dff561-964a-4f7b-9794-dba3542b63c4 -all |Where-Object {$_.Id -eq $usr.ObjectID})
#            $WNSUsrGrp = [bool](Get-AzureADGroupMember -ObjectId 192175db-72ee-4f85-a24f-c64d4681bfd8 -all $true |Where-Object {$_.ObjectID -eq $usr.ObjectID})
            $WNSUsrGrp = [bool](Get-MgGroupMember -GroupId 192175db-72ee-4f85-a24f-c64d4681bfd8 -all |Where-Object {$_.Id -eq $usr.ObjectID})
#            $AudioGrpMem = [bool](Get-AzureADGroupMember -ObjectId a9b06203-849c-40c7-93f9-6c404fdcf6d3 -all $true |Where-Object {$_.ObjectID -eq $usr.ObjectID})
            $AudioGrpMem = [bool](Get-MgGroupMember -GroupId a9b06203-849c-40c7-93f9-6c404fdcf6d3 -all |Where-Object {$_.Id -eq $usr.ObjectID})
#            $PowerBIProGrpMem = [book](Get-AzureADGroupMember -ObjectId 4246671a-2ac5-4b8a-96fd-4e4b7d817791 -all $true |Where-Object {$_.ObjectID -eq $usr.ObjectID})
            $PowerBIProGrpMem = [book](Get-MgGroupMember -GroupId 4246671a-2ac5-4b8a-96fd-4e4b7d817791 -all |Where-Object {$_.Id -eq $usr.ObjectID})

            If ($EmpGrp -eq $True)
            {
                Write-host "     Assigned to the Enterprise E5 Employee Group" -ForegroundColor Cyan
                RetentPolicy
            }
            else
            {
                If ($NonEmpGrp -eq $True)
                {
                    Write-host "     Assigned to the Enterprise E5 Non-Employee Group" -ForegroundColor Cyan
                    RetentPolicy
                }
                else
                {
                    If ($WNSUsrGrp -eq $True)
                    {
                        Write-host "     Assigned to the WNS Users Group" -ForegroundColor Cyan
                        RetentPolicy
                    }
                    else
                    {
                        Write-host "     No E5 Licenses Assigned to this individual" -ForegroundColor Red
                    }
                }
            }

            If ($AudioGrpMem -eq $True)
            {
                Write-host "     Assigned to the E5 Microsoft Audio Conferencing Sub-License is enabled" -ForegroundColor Cyan
            }

            If ($PowerBIProGrpMem -eq $True)
            {
                Write-host "     Assigned to the E5 Microsoft PowerBI Pro Sub-License is enabled" -ForegroundColor Cyan
            }

 	        write-host "Checking the Retention Policy Configuration and that the PowerBI and EMS licenses are assigned"
   	        StandardLicenses
            $Script:chkStd.Checked = $false
        }

        If ($Script:chkReview.Checked -eq $True)
        {
            If (($Script:chkE5Lic.Checked -eq $True) -or ($Global:LicDetails.SelectedItem -eq "Enterprise E5"))
            {
                EnabledE5Feature
                pause
            }
            $Script:chkReview.Checked = $false
        }

        If ($Script:chkReset.Checked -eq $True)
        {
            $SSKID = "ENTERPRISEPACK"
            $LicType = "Enterprise E3"
            $DisPlan = $Global:DisPlanEmpE3
            if ($ADUser.ExtensionAttribute1 -notlike "Employee*")
            {
               $DisPlan = $Global:DisPlanNonEmpE3
            }

            If ($ADUser.ExtensionAttribute1 -like "Employee*")
            {
 		        write-host "Resetting " $LicType "licenses to standard employee licenses" -ForegroundColor Yellow				
	        }
	        else
	        {
		        write-host "Resetting " $LicType "licenses to standard non-employee licenses" -ForegroundColor Yellow						
	        }

            EnabledE3Feature

  	        $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Resetting " + $LicType + " SubLicense Options to Standard Employee/Non-Employee for " + $Global:UPN
            Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
            $DLO = ($DisPlan.Split(","))
            $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
            Set-MsolUserLicense -UserPrincipalName $Global:UPN -LicenseOptions $MyO365Sku
            StandardLicenses
						
   	        RetentPolicy
            $Script:chkReset.Checked = $false
        }

        Start-Sleep -Seconds 12
        Add-AssignedLicenses
        Get-LicenseUsage
  
        Reset-Form
        If ($Global:txtUserInf.Text -like "*@*")
        {
            $Script:ENo = $Global:txtUserInf.Text.Substring(0,$Global:txtUserInf.Text.IndexOf("@"))
        }
        Publish-Form
    }
}while ($Global:Result -eq "OK")