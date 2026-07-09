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
#>  

<#
Function AvailLicense
{
    $E5Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPREMIUM"}
	$EMSLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EMSPREMIUM"}
    $P2Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EXCHANGEENTERPRISE"}
    $ATPP1Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ATP_ENTERPRISE"}
    $CAPLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOCAP"}
    $PBIFLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWER_BI_STANDARD"}
    $PBIPLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWER_BI_PRO"}
    $AutoPUsrLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:FLOW_PER_USER"}
    $ATPDefLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:WIN_DEF_ATP"}
    $MeetingLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MEETING_ROOM"}
    $PAppsPUsrLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWERAPPS_PER_USER"}
    $PAutoPUsrRPALic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWERAUTOMATE_ATTENDED_RPA"}
    $PhoneSysLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOEV"}
    $AudioConfLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOMEETADV"}
    $ProjP3Lic = = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:PROJECTPROFESSIONAL"}
    $ProjP5Lic = = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:PROJECTPREMIUM"}
    $VisioP1Lic = = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:VISIOONLINE_PLAN1"}
    $VisioP2Lic = = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:VISIOCLIENT"}
         
#    Send O365 Team email that the Staff Type Changes is complete

    write-host "          Current license allocations:"
    write-host "`nActive`t InUse`t Avail`t-  License Name"
    write-host "-------------------------------------------------------------"

    write-host $E5Lic.ActiveUnits "`t" $E5Lic.ConsumedUnits "`t" ($E5Lic.ActiveUnits-$E5Lic.ConsumedUnits) "`t-  Enterprise E5 License"
	write-host $E3Lic.ActiveUnits "`t" $E3Lic.ConsumedUnits "`t" ($E3Lic.ActiveUnits-$E3Lic.ConsumedUnits) "`t-  Enterprise E3 License"
    write-host $P2Lic.ActiveUnits "`t" $P2Lic.ConsumedUnits "`t" ($P2Lic.ActiveUnits-$P2Lic.ConsumedUnits) "`t-  ExchangeOnline P2 License"
    write-host $EMSLic.ActiveUnits "`t" $EMSLic.ConsumedUnits "`t" ($EMSLic.ActiveUnits-$EMSLic.ConsumedUnits) "`t-  Enterprise Mobility + Security E5 License"
    write-host $ATPP1Lic.ActiveUnits "`t" $ATPP1Lic.ConsumedUnits "`t" ($ATPP1Lic.ActiveUnits-$ATPP1Lic.ConsumedUnits) "`t-  Advanced Threat Protection Plan1 License"
    write-host $ATPDefLic.ActiveUnits "`t" $ATPDefLic.ConsumedUnits "`t" ($ATPDefLic.ActiveUnits-$ATPDefLic.ConsumedUnits) "`t-  Defender Advanced Threat Protection License"
    write-host $CAPLic.ActiveUnits "`t" $CAPLic.ConsumedUnits "`t" ($CAPLic.ActiveUnits-$CAPLic.ConsumedUnits) "`t-  Common Area Phone License"
    write-host "N/A`t" $PBIFLic.ConsumedUnits "`t N/A`t-  PowerBI (Free) License"
    write-host $PBIPLic.ActiveUnits "`t" $PBIPLic.ConsumedUnits "`t" ($PBIPLic.ActiveUnits-$PBIPLic.ConsumedUnits) "`t-  PowerBI Pro License"
    write-host $MeetingLic.ActiveUnits "`t" $MeetingLic.ConsumedUnits "`t" ($MeetingLic.ActiveUnits-$MeetingLic.ConsumedUnits) "`t-  Meeting Room License"
    write-host $AutoPUsrLic.ActiveUnits "`t" $AutoPUsrLic.ConsumedUnits "`t" ($AutoPUsrLic.ActiveUnits-$AutoPUsrLic.ConsumedUnits) "`t-  Power Automate per User Plan License"
    write-host $AutoPUsrRPALic.ActiveUnits "`t" $AutoPUsrRPALic.ConsumedUnits "`t" ($AutoPUsrRPALic.ActiveUnits-$AutoPUsrRPALic.ConsumedUnits) "`t-  Power Automate per User w/RPA Plan License"
    write-host $PAppsPUsrLic.ActiveUnits "`t" $PAppsPUsrLic.ConsumedUnits "`t" ($PAppsPUsrLic.ActiveUnits-$PAppsPUsrLic.ConsumedUnits) "`t-  Power Apps per User Plan License"
    write-host $PhoneSysLic.ActiveUnits "`t" $PhoneSysLic.ConsumedUnits "`t" ($PhoneSysLic.ActiveUnits-$PhoneSysLic.ConsumedUnits) "`t-  Phone System License"
    write-host $ProjP3Lic.ActiveUnits "`t" $ProjP3Lic.ConsumedUnits "`t" ($ProjP3Lic.ActiveUnits-$ProjP3Lic.ConsumedUnits) "`t-  Project Plan 3 License"
    write-host $ProjP5Lic.ActiveUnits "`t" $ProjP5Lic.ConsumedUnits "`t" ($ProjP5Lic.ActiveUnits-$ProjP5Lic.ConsumedUnits) "`t-  Project Plan 5 License"
    write-host $VisioP1Lic.ActiveUnits "`t" $VisioP1Lic.ConsumedUnits "`t" ($VisioP1Lic.ActiveUnits-$VisioP1Lic.ConsumedUnits) "`t-  Visio Plan 1 License"
    write-host $VisioP2Lic.ActiveUnits "`t" $VisioP2Lic.ConsumedUnits "`t" ($VisioP2Lic.ActiveUnits-$VisioP2Lic.ConsumedUnits) "`t-  Visio Plan 2 License"
}
#>

Function Add-AssignedLicenses
{
    $Global:UPN = $Global:txtUserInf.Text + "@global.ul.com"
    $Global:Exists = [bool](get-MSOLUser -UserPrincipalName $Global:UPN -ErrorAction SilentlyContinue)
    $Global:Action = "Continue"
    $Global:UserDet = ""
    $Global:UserLicense = Get-MsolUser -UserPrincipalName $Global:UPN -ErrorAction SilentlyContinue
    $LDAPFilter = "(userPrincipalName=" + $Global:UPN + ")"
    $Global:ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,proxyaddresses,extensionattribute1,extensionattribute4
                
    If ($Global:Exists -eq $true)
    {
        O365Licenses
        $Global:ButGetENo.visible = $false

$Global:lblUserInf.Text = "Employee Info: "
$Global:txtUserInf.Text = $Global:UserDet
$Global:txtUserInf.ReadOnly = $true
<#
        $TopLoc = 410
        ## User Details
        $Global:lblUserDetails = New-Object System.Windows.Forms.Label   
            $Global:lblUserDetails.Text = "Employee Info: "
            $Global:lblUserDetails.Top = $TopLoc; $Global:lblUserDetails.Left = 10; $Global:lblUserDetails.Width=120 ;$Global:lblUserDetails.AutoSize = $true 
            $Global:form.Controls.Add($Global:lblUserDetails)    # Add to Form

        $Global:lblUserDetails = New-Object System.Windows.Forms.Label
    #        $Global:lblUserDetails.Text = $Global:UPN
            $Global:lblUserDetails.Text = $Global:UserDet
            $Global:lblUserDetails.Top = $TopLoc; $Global:lblUserDetails.Left=120; $Global:lblUserDetails.Width=10 ;$Global:lblUserDetails.AutoSize = $true
            $Global:form.Controls.Add($Global:lblUserDetails)    # Add to Form
            # 
#
        $TopLoc = 430
#        $TopLoc = $TopLoc + 20
	    ## ListBox - Fill with License Details
        $Global:lblLicDet = New-Object System.Windows.Forms.Label   
            $Global:lblLicDet.Text = "License Enabled:"; $Global:lblLicDet.Top = $TopLoc; $Global:lblLicDet.Left = 10; $Global:lblLicDet.Autosize = $true  
            $Global:form.Controls.Add($Global:lblLicDet)  
            # Listbox for Location Name
            $Global:locListBox = New-Object System.Windows.Forms.ListBox  
                $Global:locListBox.Top = $TopLoc; $locListBox.Left = 120; $locListBox.Height = 130; $LocListBox.Width = 370;
                $Global:locListBox.TabIndex = 1
                # we need to populate the listbox... Example: $objListBox.Items.Add("Item 1 Test Do NOT USE") 
                # in our case, we will use a call to Azure for our "list"
#> 
                If ($Global:LicAssigned.Length -ne 0)
                {
                    $LocArray = $Global:LicAssigned.split(",")
                    $i=0   # Counter 
                    foreach ($element in $LocArray) { # Loop through Azure list and add to listbox 
                        [void] $Global:locListBox.Items.Add($element.TrimStart())  # Add element to listbox 
                        $i ++ 
                    } 
                }
                $Global:form.Controls.Add($Global:locListBox) #Add listbox to form 
                # Obtain Value with: $Global:locListBox.SelectedItem
                Build-AddActionButtons
    }
    else
    {
        $Output = $wshell.Popup("Invalid employee number.",0,"Invalid Employee Number",0+32)
        $Global:txtUserInf.Text = ""
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
    $Global:form.Size = New-Object System.Drawing.Size(700,640) #(W,H)

    ## Label and TextBox  
    ## Title Line
    $Global:lblTitleLine = New-Object System.Windows.Forms.Label   
        $lblTitleLine.Text = "Active           InUse          Avail           License Name"
        $lblTitleLine.Top = 15 ; $lblTitleLine.Left = 30; $lblTitleLine.Width=120 ;$lblTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine)    # Add to Form 

    ## E5 Licenses
    $E5Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPREMIUM"}
    $TopLoc = $TopLoc + 20
    $Global:chkE5Lic = New-Object Windows.Forms.checkbox 
        $Global:chkE5Lic.Left = $LicLoc; $Global:chkE5Lic.Width = 280; $Global:chkE5Lic.Top = ($TopLoc-5)  
        $Global:chkE5Lic.Text = "Enterprise E5 License" 
        $Global:chkE5Lic.Checked = $false   # set a default value 
        $Global:chkE5Lic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkE5Lic) 
    $Global:lblE5Act = New-Object System.Windows.Forms.Label
        $lblE5Act.Text = $E5Lic.ActiveUnits
        $lblE5Act.Top = $TopLoc ; $lblE5Act.Left = $ActLoc; $lblE5Act.Width=10 ;$lblE5Act.AutoSize = $true
        $Global:form.Controls.Add($Global:lblE5Act)    # Add to Form
    $Global:lblE5InU = New-Object System.Windows.Forms.Label
        $lblE5InU.Text = $E5Lic.ConsumedUnits
        $lblE5InU.Top = $TopLoc ; $lblE5InU.Left = $InULoc; $lblE5InU.Width=10 ;$lblE5InU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblE5InU)    # Add to Form
    $Global:lblE5Avl = New-Object System.Windows.Forms.Label
        $lblE5Avl.Text = ($E5Lic.ActiveUnits-$E5Lic.ConsumedUnits)
        $lblE5Avl.Top = $TopLoc ; $lblE5Avl.Left = $AvlLoc; $lblE5Avl.Width=10 ;$lblE5Avl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblE5Avl)    # Add to Form

    ## AudioConf Licenses
    $AudioConfLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOMEETADV"}
    $TopLoc = $TopLoc + 20
    $Global:chkAudioConfLic = New-Object Windows.Forms.checkbox 
        $Global:chkAudioConfLic.Left = $LicLoc+10; $Global:chkAudioConfLic.Width = 280; $Global:chkAudioConfLic.Top = ($TopLoc-5)  
        $Global:chkAudioConfLic.Text = "E5 Audio Conferencing Feature" 
        $Global:chkAudioConfLic.Checked = $false   # set a default value 
        $Global:chkAudioConfLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkAudioConfLic)
    $Global:lblAudioConfAct = New-Object System.Windows.Forms.Label
        $lblAudioConfAct.Text = $E5Lic.ActiveUnits
        $lblAudioConfAct.Top = $TopLoc ; $lblAudioConfAct.Left = $ActLoc; $lblAudioConfAct.Width=10 ;$lblAudioConfAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAudioConfAct)    # Add to Form
    $Global:lblAudioConfInU = New-Object System.Windows.Forms.Label
        $AudioMem = Get-AzureADGroupMember -ObjectId a9b06203-849c-40c7-93f9-6c404fdcf6d3 -all $true
        $lblAudioConfInU.Text = $AudioMem.count
        $lblAudioConfInU.Top = $TopLoc ; $lblAudioConfInU.Left = $InULoc; $lblAudioConfInU.Width=10 ;$lblAudioConfInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAudioConfInU)    # Add to Form
    $Global:lblAudioConfAvl = New-Object System.Windows.Forms.Label
        $lblAudioConfAvl.Text = ($E5Lic.ActiveUnits-$AudioMem.count)
        $lblAudioConfAvl.Top = $TopLoc ; $lblAudioConfAvl.Left = $AvlLoc; $lblAudioConfAvl.Width=10 ;$lblAudioConfAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAudioConfAvl)    # Add to Form

    ## PowerBI Pro Licenses
    $PBIPLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWER_BI_PRO"}
    $TopLoc = $TopLoc + 20
    $Global:chkPBIPLic = New-Object Windows.Forms.checkbox 
        $Global:chkPBIPLic.Left = $LicLoc+10; $Global:chkPBIPLic.Width = 280; $Global:chkPBIPLic.Top = ($TopLoc-5)  
        $Global:chkPBIPLic.Text = "E5 PowerBI Pro Featuure" 
        $Global:chkPBIPLic.Checked = $false   # set a default value 
        $Global:chkPBIPLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkPBIPLic)         
    $Global:lblPBIPAct = New-Object System.Windows.Forms.Label
        $lblPBIPAct.Text = $E5Lic.ActiveUnits
        $lblPBIPAct.Top = $TopLoc ; $lblPBIPAct.Left = $ActLoc ; $lblPBIPAct.Width=10 ;$lblPBIPAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblPBIPAct)    # Add to Form
    $Global:lblPBIPInU = New-Object System.Windows.Forms.Label
        $PBIPMem = Get-AzureADGroupMember -ObjectId 4246671a-2ac5-4b8a-96fd-4e4b7d817791 -all $true
        $lblPBIPInU.Text = $PBIPMem.count
        $lblPBIPInU.Top = $TopLoc ; $lblPBIPInU.Left = $InULoc; $lblPBIPInU.Width=10 ;$lblPBIPInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblPBIPInU)    # Add to Form
    $Global:lblPBIPAvl = New-Object System.Windows.Forms.Label
        $lblPBIPAvl.Text = ($E5Lic.ActiveUnits-$PBIPMem.count)
        $lblPBIPAvl.Top = $TopLoc ; $lblPBIPAvl.Left = $AvlLoc; $lblPBIPAvl.Width=10 ;$lblPBIPAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblPBIPAvl)    # Add to Form

    ## EMS Licenses
    $EMSLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EMSPREMIUM"}
    $TopLoc = $TopLoc + 20
    $Global:chkEMSLic = New-Object Windows.Forms.checkbox 
        $Global:chkEMSLic.Left = $LicLoc; $Global:chkEMSLic.Width = 280; $Global:chkEMSLic.Top = ($TopLoc-5)  
        $Global:chkEMSLic.Text = "Enterprise Mobility + Security E5 License" 
        $Global:chkEMSLic.Checked = $false   # set a default value 
        $Global:chkEMSLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkEMSLic)
    $Global:lblEMSAct = New-Object System.Windows.Forms.Label
        $lblEMSAct.Text = $EMSLic.ActiveUnits
        $lblEMSAct.Top = $TopLoc ; $lblEMSAct.Left = $ActLoc; $lblEMSAct.Width=10 ;$lblEMSAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblEMSAct)    # Add to Form
    $Global:lblEMSInU = New-Object System.Windows.Forms.Label
        $lblEMSInU.Text = $EMSLic.ConsumedUnits
        $lblEMSInU.Top = $TopLoc ; $lblEMSInU.Left = $InULoc; $lblEMSInU.Width=10 ;$lblEMSInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblEMSInU)    # Add to Form
    $Global:lblEMSAvl = New-Object System.Windows.Forms.Label
        $lblEMSAvl.Text = ($EMSLic.ActiveUnits-$EMSLic.ConsumedUnits)
        $lblEMSAvl.Top = $TopLoc ; $lblEMSAvl.Left = $AvlLoc; $lblEMSAvl.Width=10 ;$lblEMSAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblEMSAvl)    # Add to Form
        # 
 
    ## P2 Licenses
    $P2Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EXCHANGEENTERPRISE"}
    $TopLoc = $TopLoc + 20   
    $Global:chkP2Lic = New-Object Windows.Forms.checkbox 
        $Global:chkP2Lic.Left = $LicLoc; $Global:chkP2Lic.Width = 280; $Global:chkP2Lic.Top = ($TopLoc-5)  
        $Global:chkP2Lic.Text = "Exchange Online P2 License" 
        $Global:chkP2Lic.Checked = $false   # set a default value 
        $Global:chkP2Lic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkP2Lic) 
    $Global:lblP2Act = New-Object System.Windows.Forms.Label
        $lblP2Act.Text = $P2Lic.ActiveUnits
        $lblP2Act.Top = $TopLoc ; $lblP2Act.Left = $ActLoc; $lblP2Act.Width=10 ;$lblP2Act.AutoSize = $true
        $Global:form.Controls.Add($Global:lblP2Act)    # Add to Form
    $Global:lblP2InU = New-Object System.Windows.Forms.Label
        $lblP2InU.Text = $P2Lic.ConsumedUnits
        $lblP2InU.Top = $TopLoc ; $lblP2InU.Left = $InULoc; $lblP2InU.Width=10 ;$lblP2InU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblP2InU)    # Add to Form
    $Global:lblP2Avl = New-Object System.Windows.Forms.Label
        $lblP2Avl.Text = ($P2Lic.ActiveUnits-$P2Lic.ConsumedUnits)
        $lblP2Avl.Top = $TopLoc ; $lblP2Avl.Left = $AvlLoc; $lblP2Avl.Width=10 ;$lblP2Avl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblP2Avl)    # Add to Form

    ## MCOCAP Licenses
    $CAPLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOCAP"}
    $TopLoc = $TopLoc + 20
    $Global:chkCAPLic = New-Object Windows.Forms.checkbox 
        $Global:chkCAPLic.Left = $LicLoc; $Global:chkCAPLic.Width = 280; $Global:chkCAPLic.Top = ($TopLoc-5)  
        $Global:chkCAPLic.Text = "Common Area Phone License" 
        $Global:chkCAPLic.Checked = $false   # set a default value 
        $Global:chkCAPLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkCAPLic)
    $Global:lblCAPAct = New-Object System.Windows.Forms.Label
        $lblCAPAct.Text = $CAPLic.ActiveUnits
        $lblCAPAct.Top = $TopLoc ; $lblCAPAct.Left = $ActLoc; $lblCAPAct.Width=10 ;$lblCAPAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblCAPAct)    # Add to Form
    $Global:lblCAPInU = New-Object System.Windows.Forms.Label
        $lblCAPInU.Text = $CAPLic.ConsumedUnits
        $lblCAPInU.Top = $TopLoc ; $lblCAPInU.Left = $InULoc; $lblCAPInU.Width=10 ;$lblCAPInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblCAPInU)    # Add to Form
    $Global:lblCAPAvl = New-Object System.Windows.Forms.Label
        $lblCAPAvl.Text = ($CAPLic.ActiveUnits-$CAPLic.ConsumedUnits)
        $lblCAPAvl.Top = $TopLoc ; $lblCAPAvl.Left = $AvlLoc; $lblCAPAvl.Width=10 ;$lblCAPAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblCAPAvl)    # Add to Form

   ## ATPDef Licenses
    $ATPDefLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:WIN_DEF_ATP"}
    $TopLoc = $TopLoc + 20
    $Global:chkATPDefLic = New-Object Windows.Forms.checkbox 
        $Global:chkATPDefLic.Left = $LicLoc; $Global:chkATPDefLic.Width = 280; $Global:chkATPDefLic.Top = ($TopLoc-5)  
        $Global:chkATPDefLic.Text = "Defender Advanced Theat Protection License" 
        $Global:chkATPDefLic.Checked = $false   # set a default value 
        $Global:chkATPDefLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkATPDefLic)
    $Global:lblATPDefAct = New-Object System.Windows.Forms.Label
        $lblATPDefAct.Text = $ATPDefLic.ActiveUnits
        $lblATPDefAct.Top = $TopLoc ; $lblATPDefAct.Left = $ActLoc; $lblATPDefAct.Width=10 ;$lblATPDefAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblATPDefAct)    # Add to Form
    $Global:lblATPDefInU = New-Object System.Windows.Forms.Label
        $lblATPDefInU.Text = $ATPDefLic.ConsumedUnits
        $lblATPDefInU.Top = $TopLoc ; $lblATPDefInU.Left = $InULoc; $lblATPDefInU.Width=10 ;$lblATPDefInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblATPDefInU)    # Add to Form
    $Global:lblATPDefAvl = New-Object System.Windows.Forms.Label
        $lblATPDefAvl.Text = ($ATPDefLic.ActiveUnits-$ATPDefLic.ConsumedUnits)
        $lblATPDefAvl.Top = $TopLoc ; $lblATPDefAvl.Left = $AvlLoc; $lblATPDefAvl.Width=10 ;$lblATPDefAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblATPDefAvl)    # Add to Form

    ## Meeting Room Licenses
    $MeetingLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MEETING_ROOM"}
    $TopLoc = $TopLoc + 20
    $Global:chkMeetingLic = New-Object Windows.Forms.checkbox 
        $Global:chkMeetingLic.Left = $LicLoc; $Global:chkMeetingLic.Width = 280; $Global:chkMeetingLic.Top = ($TopLoc-5)  
        $Global:chkMeetingLic.Text = "Microsoft Meeting Room License" 
        $Global:chkMeetingLic.Checked = $false   # set a default value 
        $Global:chkMeetingLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkMeetingLic)          
    $Global:lblMeetingAct = New-Object System.Windows.Forms.Label
        $lblMeetingAct.Text = $MeetingLic.ActiveUnits
        $lblMeetingAct.Top = $TopLoc ; $lblMeetingAct.Left = $ActLoc; $lblMeetingAct.Width=10 ;$lblMeetingAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblMeetingAct)    # Add to Form
    $Global:lblMeetingInU = New-Object System.Windows.Forms.Label
        $lblMeetingInU.Text = $MeetingLic.ConsumedUnits
        $lblMeetingInU.Top = $TopLoc ; $lblMeetingInU.Left = $InULoc; $lblMeetingInU.Width=10 ;$lblMeetingInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblMeetingInU)    # Add to Form
    $Global:lblMeetingAvl = New-Object System.Windows.Forms.Label
        $lblMeetingAvl.Text = ($MeetingLic.ActiveUnits-$MeetingLic.ConsumedUnits)
        $lblMeetingAvl.Top = $TopLoc ; $lblMeetingAvl.Left = $AvlLoc; $lblMeetingAvl.Width=10 ;$lblMeetingAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblMeetingAvl)    # Add to Form

    ## PhoneSystem Licenses
    $PhoneSysLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOEV"}
    $TopLoc = $TopLoc + 20
    $Global:chkPhoneSysLic = New-Object Windows.Forms.checkbox 
        $Global:chkPhoneSysLic.Left = $LicLoc; $Global:chkPhoneSysLic.Width = 280; $Global:chkPhoneSysLic.Top = ($TopLoc-5)  
        $Global:chkPhoneSysLic.Text = "Phone System License" 
        $Global:chkPhoneSysLic.Checked = $false   # set a default value 
        $Global:chkPhoneSysLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkPhoneSysLic)
    $Global:lblPhoneSysAct = New-Object System.Windows.Forms.Label
        $lblPhoneSysAct.Text = $PhoneSysLic.ActiveUnits
        $lblPhoneSysAct.Top = $TopLoc ; $lblPhoneSysAct.Left = $ActLoc; $lblPhoneSysAct.Width=10 ;$lblPhoneSysAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblPhoneSysAct)    # Add to Form
    $Global:lblPhoneSysInU = New-Object System.Windows.Forms.Label
        $lblPhoneSysInU.Text = $PhoneSysLic.ConsumedUnits
        $lblPhoneSysInU.Top = $TopLoc ; $lblPhoneSysInU.Left = $InULoc; $lblPhoneSysInU.Width=10 ;$lblPhoneSysInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblPhoneSysInU)    # Add to Form
    $Global:lblPhoneSysAvl = New-Object System.Windows.Forms.Label
        $lblPhoneSysAvl.Text = ($PhoneSysLic.ActiveUnits-$PhoneSysLic.ConsumedUnits)
        $lblPhoneSysAvl.Top = $TopLoc ; $lblPhoneSysAvl.Left = $AvlLoc; $lblPhoneSysAvl.Width=10 ;$lblPhoneSysAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblPhoneSysAvl)    # Add to Form

    ## PowerApps per User Plan Licenses
    $AppsPUsrLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWERAPPS_PER_USER"}
    $TopLoc = $TopLoc + 20
    $Global:chkAppsPUsrLic = New-Object Windows.Forms.checkbox 
        $Global:chkAppsPUsrLic.Left = $LicLoc; $Global:chkAppsPUsrLic.Width = 280; $Global:chkAppsPUsrLic.Top = ($TopLoc-5)  
        $Global:chkAppsPUsrLic.Text = "PowerApps p/User Plan License" 
        $Global:chkAppsPUsrLic.Checked = $false   # set a default value 
        $Global:chkAppsPUsrLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkAppsPUsrLic) 
    $Global:lblAppsPUsrAct = New-Object System.Windows.Forms.Label
        $lblAppsPUsrAct.Text = $AppsPUsrLic.ActiveUnits
        $lblAppsPUsrAct.Top = $TopLoc ; $lblAppsPUsrAct.Left = $ActLoc; $lblAppsPUsrAct.Width=10 ;$lblAppsPUsrAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAppsPUsrAct)    # Add to Form
    $Global:lblAppsPUsrInU = New-Object System.Windows.Forms.Label
        $lblAppsPUsrInU.Text = $AppsPUsrLic.ConsumedUnits
        $lblAppsPUsrInU.Top = $TopLoc ; $lblAppsPUsrInU.Left = $InULoc; $lblAppsPUsrInU.Width=10 ;$lblAppsPUsrInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAppsPUsrInU)    # Add to Form
    $Global:lblAppsPUsrAvl = New-Object System.Windows.Forms.Label
        $lblAppsPUsrAvl.Text = ($AppsPUsrLic.ActiveUnits-$AppsPUsrLic.ConsumedUnits)
        $lblAppsPUsrAvl.Top = $TopLoc ; $lblAppsPUsrAvl.Left = $AvlLoc; $lblAppsPUsrAvl.Width=10 ;$lblAppsPUsrAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAppsPUsrAvl)    # Add to Form

    ## Power Automate per User Plan Licenses
    $AutoPUsrLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:FLOW_PER_USER"}
    $TopLoc = $TopLoc + 20
    $Global:chkAutoPUsrLic = New-Object Windows.Forms.checkbox 
        $Global:chkAutoPUsrLic.Left = $LicLoc; $Global:chkAutoPUsrLic.Width = 280; $Global:chkAutoPUsrLic.Top = ($TopLoc-5)  
        $Global:chkAutoPUsrLic.Text = "Power Automate p/User Plan License" 
        $Global:chkAutoPUsrLic.Checked = $false   # set a default value 
        $Global:chkAutoPUsrLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkAutoPUsrLic) 
    $Global:lblAutoPUsrAct = New-Object System.Windows.Forms.Label
        $lblAutoPUsrAct.Text = $AutoPUsrLic.ActiveUnits
        $lblAutoPUsrAct.Top = $TopLoc ; $lblAutoPUsrAct.Left = $ActLoc; $lblAutoPUsrAct.Width=10 ;$lblAutoPUsrAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAutoPUsrAct)    # Add to Form
    $Global:lblAutoPUsrInU = New-Object System.Windows.Forms.Label
        $lblAutoPUsrInU.Text = $AutoPUsrLic.ConsumedUnits
        $lblAutoPUsrInU.Top = $TopLoc ; $lblAutoPUsrInU.Left = $InULoc; $lblAutoPUsrInU.Width=10 ;$lblAutoPUsrInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAutoPUsrInU)    # Add to Form
    $Global:lblAutoPUsrAvl = New-Object System.Windows.Forms.Label
        $lblAutoPUsrAvl.Text = ($AutoPUsrLic.ActiveUnits-$AutoPUsrLic.ConsumedUnits)
        $lblAutoPUsrAvl.Top = $TopLoc ; $lblAutoPUsrAvl.Left = $AvlLoc; $lblAutoPUsrAvl.Width=10 ;$lblAutoPUsrAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAutoPUsrAvl)    # Add to Form
##
    ## Power Automate per User w/RPA Licenses
    $AutoPUsrRPALic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWERAUTOMATE_ATTENDED_RPA"}
    $TopLoc = $TopLoc + 20
    $Global:chkAutoPUsrRPALic = New-Object Windows.Forms.checkbox 
        $Global:chkAutoPUsrRPALic.Left = $LicLoc; $Global:chkAutoPUsrRPALic.Width = 280; $Global:chkAutoPUsrRPALic.Top = ($TopLoc-5)  
        $Global:chkAutoPUsrRPALic.Text = "Power Automate p/User w/RPA Plan License" 
        $Global:chkAutoPUsrRPALic.Checked = $false   # set a default value 
        $Global:chkAutoPUsrRPALic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkAutoPUsrRPALic) 
    $Global:lblAutoPUsrAct = New-Object System.Windows.Forms.Label
        $lblAutoPUsrAct.Text = $AutoPUsrRPALic.ActiveUnits
        $lblAutoPUsrAct.Top = $TopLoc ; $lblAutoPUsrAct.Left = $ActLoc; $lblAutoPUsrAct.Width=10 ;$lblAutoPUsrAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAutoPUsrAct)    # Add to Form
    $Global:lblAutoPUsrInU = New-Object System.Windows.Forms.Label
        $lblAutoPUsrInU.Text = $AutoPUsrRPALic.ConsumedUnits
        $lblAutoPUsrInU.Top = $TopLoc ; $lblAutoPUsrInU.Left = $InULoc; $lblAutoPUsrInU.Width=10 ;$lblAutoPUsrInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAutoPUsrInU)    # Add to Form
    $Global:lblAutoPUsrAvl = New-Object System.Windows.Forms.Label
        $lblAutoPUsrAvl.Text = ($AutoPUsrRPALic.ActiveUnits-$AutoPUsrRPALic.ConsumedUnits)
        $lblAutoPUsrAvl.Top = $TopLoc ; $lblAutoPUsrAvl.Left = $AvlLoc; $lblAutoPUsrAvl.Width=10 ;$lblAutoPUsrAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAutoPUsrAvl)    # Add to Form
##
    ## PowerBI Free Licenses
    $PBIFLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWER_BI_STANDARD"}
    $TopLoc = $TopLoc + 20
    $Global:chkPBIFLic = New-Object Windows.Forms.checkbox 
        $Global:chkPBIFLic.Left = $LicLoc; $Global:chkPBIFLic.Width = 280; $Global:chkPBIFLic.Top = ($TopLoc-5)  
        $Global:chkPBIFLic.Text = "PowerBI (Free) License" 
        $Global:chkPBIFLic.Checked = $false   # set a default value 
        $Global:chkPBIFLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkPBIFLic)
    $Global:lblPBIFAct = New-Object System.Windows.Forms.Label
        $lblPBIFAct.Text = $PBIFLic.ActiveUnits
        $lblPBIFAct.Top = $TopLoc ; $lblPBIFAct.Left = $ActLoc; $lblPBIFAct.Width=10 ;$lblPBIFAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblPBIFAct)    # Add to Form
    $Global:lblPBIFInU = New-Object System.Windows.Forms.Label
        $lblPBIFInU.Text = $PBIFLic.ConsumedUnits
        $lblPBIFInU.Top = $TopLoc ; $lblPBIFInU.Left = $InULoc; $lblPBIFInU.Width=10 ;$lblPBIFInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblPBIFInU)    # Add to Form
    $Global:lblPBIFAvl = New-Object System.Windows.Forms.Label
        $lblPBIFAvl.Text = ($PBIFLic.ActiveUnits-$PBIFLic.ConsumedUnits)
        $lblPBIFAvl.Top = $TopLoc ; $lblPBIFAvl.Left = $AvlLoc; $lblPBIFAvl.Width=10 ;$lblPBIFAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblPBIFAvl)    # Add to Form

    ## Project Plan 3 Licenses
    $ProjP3Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:PROJECTPROFESSIONAL"}
    $TopLoc = $TopLoc + 20
    $Global:chkProjP3Lic = New-Object Windows.Forms.checkbox 
        $Global:chkProjP3Lic.Left = $LicLoc; $Global:chkProjP3Lic.Width = 280; $Global:chkProjP3Lic.Top = ($TopLoc-5)  
        $Global:chkProjP3Lic.Text = "Project Plan 3 License" 
        $Global:chkProjP3Lic.Checked = $false   # set a default value 
        $Global:chkProjP3Lic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkProjP3Lic)
    $Global:lblProjP3Act = New-Object System.Windows.Forms.Label
        $lblProjP3Act.Text = $ProjP3Lic.ActiveUnits
        $lblProjP3Act.Top = $TopLoc ; $lblProjP3Act.Left = $ActLoc; $lblProjP3Act.Width=10 ;$lblProjP3Act.AutoSize = $true
        $Global:form.Controls.Add($Global:lblProjP3Act)    # Add to Form
    $Global:lblProjP3InU = New-Object System.Windows.Forms.Label
        $lblProjP3InU.Text = $ProjP3Lic.ConsumedUnits
        $lblProjP3InU.Top = $TopLoc ; $lblProjP3InU.Left = $InULoc; $lblProjP3InU.Width=10 ;$lblProjP3InU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblProjP3InU)    # Add to Form
    $Global:lblProjP3Avl = New-Object System.Windows.Forms.Label
        $lblProjP3Avl.Text = ($ProjP3Lic.ActiveUnits-$ProjP3Lic.ConsumedUnits)
        $lblProjP3Avl.Top = $TopLoc ; $lblProjP3Avl.Left = $AvlLoc; $lblProjP3Avl.Width=10 ;$lblProjP3Avl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblProjP3Avl)    # Add to Form

    ## Project Plan 5 Licenses
    $ProjP5Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:PROJECTPREMIUM"}
    $TopLoc = $TopLoc + 20
    $Global:chkProjP5Lic = New-Object Windows.Forms.checkbox 
        $Global:chkProjP5Lic.Left = $LicLoc; $Global:chkProjP5Lic.Width = 280; $Global:chkProjP5Lic.Top = ($TopLoc-5)  
        $Global:chkProjP5Lic.Text = "Project Plan 5 License" 
        $Global:chkProjP5Lic.Checked = $false   # set a default value 
        $Global:chkProjP5Lic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkProjP5Lic)
    $Global:lblProjP5Act = New-Object System.Windows.Forms.Label
        $lblProjP5Act.Text = $ProjP5Lic.ActiveUnits
        $lblProjP5Act.Top = $TopLoc ; $lblProjP5Act.Left = $ActLoc; $lblProjP5Act.Width=10 ;$lblProjP5Act.AutoSize = $true
        $Global:form.Controls.Add($Global:lblProjP5Act)    # Add to Form
    $Global:lblProjP5InU = New-Object System.Windows.Forms.Label
        $lblProjP5InU.Text = $ProjP5Lic.ConsumedUnits
        $lblProjP5InU.Top = $TopLoc ; $lblProjP5InU.Left = $InULoc; $lblProjP5InU.Width=10 ;$lblProjP5InU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblProjP5InU)    # Add to Form
    $Global:lblProjP5Avl = New-Object System.Windows.Forms.Label
        $lblProjP5Avl.Text = ($ProjP5Lic.ActiveUnits-$ProjP5Lic.ConsumedUnits)
        $lblProjP5Avl.Top = $TopLoc ; $lblProjP5Avl.Left = $AvlLoc; $lblProjP5Avl.Width=10 ;$lblProjP5Avl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblProjP5Avl)    # Add to Form
        
    ## Visio Online Plan 1 Licenses
    $VisioP1Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:VISIOONLINE_PLAN1"}
    $TopLoc = $TopLoc + 20
    $Global:chkVisioP1Lic = New-Object Windows.Forms.checkbox 
        $Global:chkVisioP1Lic.Left = $LicLoc; $Global:chkVisioP1Lic.Width = 280; $Global:chkVisioP1Lic.Top = ($TopLoc-5)  
        $Global:chkVisioP1Lic.Text = "Visio Plan 1 License" 
        $Global:chkVisioP1Lic.Checked = $false   # set a default value 
        $Global:chkVisioP1Lic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkVisioP1Lic)
    $Global:lblVisioP1Act = New-Object System.Windows.Forms.Label
        $lblVisioP1Act.Text = $VisioP1Lic.ActiveUnits
        $lblVisioP1Act.Top = $TopLoc ; $lblVisioP1Act.Left = $ActLoc; $lblVisioP1Act.Width=10 ;$lblVisioP1Act.AutoSize = $true
        $Global:form.Controls.Add($Global:lblVisioP1Act)    # Add to Form
    $Global:lblVisioP1InU = New-Object System.Windows.Forms.Label
        $lblVisioP1InU.Text = $VisioP1Lic.ConsumedUnits
        $lblVisioP1InU.Top = $TopLoc ; $lblVisioP1InU.Left = $InULoc; $lblVisioP1InU.Width=10 ;$lblVisioP1InU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblVisioP1InU)    # Add to Form
    $Global:lblVisioP1Avl = New-Object System.Windows.Forms.Label
        $lblVisioP1Avl.Text = ($VisioP1Lic.ActiveUnits-$VisioP1Lic.ConsumedUnits)
        $lblVisioP1Avl.Top = $TopLoc ; $lblVisioP1Avl.Left = $AvlLoc; $lblVisioP1Avl.Width=10 ;$lblVisioP1Avl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblVisioP1Avl)    # Add to Form

    ## Visio Online Plan 2 Licenses
    $VisioP2Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:VISIOCLIENT"}
    $TopLoc = $TopLoc + 20
    $Global:chkVisioP2Lic = New-Object Windows.Forms.checkbox 
        $Global:chkVisioP2Lic.Left = $LicLoc; $Global:chkVisioP2Lic.Width = 280; $Global:chkVisioP2Lic.Top = ($TopLoc-5)  
        $Global:chkVisioP2Lic.Text = "Visio Plan 2 License" 
        $Global:chkVisioP2Lic.Checked = $false   # set a default value 
        $Global:chkVisioP2Lic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkVisioP2Lic)
    $Global:lblVisioP2Act = New-Object System.Windows.Forms.Label
        $lblVisioP2Act.Text = $VisioP2Lic.ActiveUnits
        $lblVisioP2Act.Top = $TopLoc ; $lblVisioP2Act.Left = $ActLoc; $lblVisioP2Act.Width=10 ;$lblVisioP2Act.AutoSize = $true
        $Global:form.Controls.Add($Global:lblVisioP2Act)    # Add to Form
    $Global:lblVisioP2InU = New-Object System.Windows.Forms.Label
        $lblVisioP2InU.Text = $VisioP2Lic.ConsumedUnits
        $lblVisioP2InU.Top = $TopLoc ; $lblVisioP2InU.Left = $InULoc; $lblVisioP2InU.Width=10 ;$lblVisioP2InU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblVisioP2InU)    # Add to Form
    $Global:lblVisioP2Avl = New-Object System.Windows.Forms.Label
        $lblVisioP2Avl.Text = ($VisioP2Lic.ActiveUnits-$VisioP2Lic.ConsumedUnits)
        $lblVisioP2Avl.Top = $TopLoc ; $lblVisioP2Avl.Left = $AvlLoc; $lblVisioP2Avl.Width=10 ;$lblVisioP2Avl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblVisioP2Avl)    # Add to Form

##############Added
#    $TopLoc = 410
    $TopLoc = $TopLoc + 40
    ## User Details
    $Global:lblUserInf = New-Object System.Windows.Forms.Label
        $Global:lblUserInf.Text = "Employee Number:"
        $Global:lblUserInf.Top = $TopLoc; $Global:lblUserInf.Left=10; $Global:lblUserInf.Width=10 ;$Global:lblUserInf.AutoSize = $true
        $Global:form.Controls.Add($Global:lblUserInf)    # Add to Form
    $Global:txtUserInf = New-Object System.Windows.Forms.TextBox
        $Global:txtUserInf.Top = $TopLoc; $Global:txtUserInf.Left=120; $Global:txtUserInf.Width=100 ;$Global:txtUserInf.AutoSize = $true
        $Global:form.Controls.Add($Global:txtUserInf)    # Add to Form
        $Global:InputFocus = $Global:txtUserInf
        $Global:txtUserInf.Add_Click({
            $Global:ButGetENo.Visible = $True
            $Global:lblUserInf.Text = "Employee Number:"
            $Global:txtUserInf.Text = ""
            $Global:lblUserInf.Visible = $True
            $Global:txtUserInf.ReadOnly = $False
            $Global:lblLicDet.Visible = $False
            $Global:locListBox.Visible = $False
            $Global:locListBox.Items.Clear()
            $Global:txtUserInf.Left=120; $Global:txtUserInf.Width=100
             })

        $Global:ButGetENo = New-Object Windows.Forms.Button
            $Global:ButGetENo.Location = New-object System.Drawing.Size(230,$TopLoc)
            $Global:ButGetENo.Size = new-Object System.Drawing.Size(150,20)
            $Global:ButGetENo.Text = "Get Employee Details"
            $Global:form.Controls.Add($Global:ButGetENo)
            $Global:ButGetENo.Add_Click({
                $Script:ENo = $Global:txtUserInf.Text
                Add-AssignedLicenses
                $Global:lblUserInf.Text = "Employee Info:"
                $Global:txtUserInf.Left=120; $Global:txtUserInf.Width=370
                $Global:lblLicDet.Visible = $True
                $Global:locListBox.Visible = $True
#                $Global:lblLicDet.Visible = $False
            })

        $TopLoc = $TopLoc + 30
	    ## ListBox - Fill with License Details
        $Global:lblLicDet = New-Object System.Windows.Forms.Label   
            $Global:lblLicDet.Text = "License Enabled:"; $Global:lblLicDet.Top = $TopLoc; $Global:lblLicDet.Left = 10; $Global:lblLicDet.Autosize = $true
            $Global:lblLicDet.Visible = $False  
            $Global:form.Controls.Add($Global:lblLicDet)  
        $Global:locListBox = New-Object System.Windows.Forms.ListBox  
            $Global:locListBox.Top = $TopLoc; $locListBox.Left = 120; $locListBox.Height = 130; $LocListBox.Width = 370;
            $Global:locListBox.Visible = $False
            $Global:locListBox.TabIndex = 1
<#
                If ($Global:LicAssigned.Length -ne 0)
                {
                    $LocArray = $Global:LicAssigned.split(",")
                    $i=0   # Counter 
                    foreach ($element in $LocArray) { # Loop through Azure list and add to listbox 
                        [void] $Global:locListBox.Items.Add($element.TrimStart())  # Add element to listbox 
                        $i ++ 
                    } 
                }
#>
                $Global:form.Controls.Add($Global:locListBox) #Add listbox to form 
                # Obtain Value with: $Global:locListBox.SelectedItem


}

function Build-AddUserDetailsForm
{
<#
    $TopLoc = 410
    ## User Details
    $Global:lblUserInf = New-Object System.Windows.Forms.Label
        $Global:lblUserInf.Text = "Employee Number:"
        $Global:lblUserInf.Top = $TopLoc; $Global:lblUserInf.Left=10; $Global:lblUserInf.Width=10 ;$Global:lblUserInf.AutoSize = $true
        $Global:form.Controls.Add($Global:lblUserInf)    # Add to Form
    $Global:txtUserInf = New-Object System.Windows.Forms.TextBox
        $Global:txtUserInf.Top = $TopLoc; $Global:txtUserInf.Left=120; $Global:txtUserInf.Width=100 ;$Global:txtUserInf.AutoSize = $true
        $Global:form.Controls.Add($Global:txtUserInf)    # Add to Form
        $Global:InputFocus = $Global:txtUserInf
        $Global:txtUserInf.Add_Click({
            $Global:ButGetENo.Visible = $True
            $Global:lblUserInf.Text = "Employee Number:"
            $Global:txtUserInf.Text = ""
            $Global:txtUserInf.ReadOnly = $False
            $Global:locListBox.Items.Clear()
            $Global:txtUserInf.Left=120; $Global:txtUserInf.Width=100
             })

        $Global:ButGetENo = New-Object Windows.Forms.Button
            $Global:ButGetENo.Location = New-object System.Drawing.Size(230,$TopLoc)
            $Global:ButGetENo.Size = new-Object System.Drawing.Size(150,20)
            $Global:ButGetENo.Text = "Get Employee Details"
            $Global:form.Controls.Add($Global:ButGetENo)
            $Global:ButGetENo.Add_Click({
                Add-AssignedLicenses
                $Global:lblUserInf.Text = "Employee Info:"
                $Global:txtUserInf.Left=120; $Global:txtUserInf.Width=370
            })
#>
}

Function Build-AddActionButtons
{
   #Action Side CheckBoxes

    ## Assign a License         
    $Global:chkAssign = New-Object Windows.Forms.RadioButton
        $Global:chkAssign.Left = 530; $Global:chkAssign.Width = 200; $Global:chkAssign.Top = 100  
        $Global:chkAssign.Text = "Assign" 
        $Global:chkAssign.Checked = $Global:chkAssign.Checked   # set a default value 
        $Global:chkAssign.TabIndex = 2 
        $Global:form.Controls.Add($Global:chkAssign) 
        # Obtain Value with: $Global:chkAssign.Checked

    ## Assign Standard License Set
    $Global:chkStd = New-Object Windows.Forms.RadioButton 
        $Global:chkStd.Left = 530; $Global:chkStd.Width = 200; $Global:chkStd.Top = 125
        $Global:chkStd.Text = "Assign Standard Set"
        $Global:chkStd.Checked = $Global:chkStd.Checked   # set a default value 
        $Global:chkStd.TabIndex = 2 
        $Global:form.Controls.Add($Global:chkStd) 

    ## Remove a License         
    $Global:chkRemove = New-Object Windows.Forms.RadioButton
        $Global:chkRemove.Left = 530; $Global:chkRemove.Width = 200; $Global:chkRemove.Top = 150
        $Global:chkRemove.Text = "Remove" 
        $Global:chkRemove.Checked = $Global:chkRemove.Checked   # set a default value 
        $Global:chkRemove.TabIndex = 2 
        $Global:form.Controls.Add($Global:chkRemove) 
        # Obtain Value with: $Global:chkRemove.Checked

    ## Review SubLicense Assignment
    $Global:chkReview = New-Object Windows.Forms.RadioButton
        $Global:chkReview.Left = 530; $Global:chkReview.Width = 200; $Global:chkReview.Top = 175
        $Global:chkReview.Text = "Review Options" 
        $Global:chkReview.Checked = $Global:chkReview.Checked   # set a default value 
        $Global:chkReview.TabIndex = 2 
        $Global:form.Controls.Add($Global:chkReview) 
        # Obtain Value with: $Global:chkReplace.Checked

    ## Reset License Assignment
    $Global:chkReset = New-Object Windows.Forms.RadioButton
        $Global:chkReset.Left =530; $Global:chkReset.Width = 200; $Global:chkReset.Top = 200  
        $Global:chkReset.Text = "Reset Options" 
        $Global:chkReset.Checked = $Global:chkReset.Checked   # set a default value 
        $Global:chkReset.TabIndex = 2 
        $Global:form.Controls.Add($Global:chkReset) 
        # Obtain Value with: $Global:chkReplace.Checked
}

Function Add-FormCloseButtonOnly
{
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    $Global:buttonPanel = New-Object Windows.Forms.Panel  
    $buttonPanel.Size = New-Object Drawing.Size @(400,40) 
    $buttonPanel.Dock = "Bottom"    
    $Global:cancelButton = New-Object Windows.Forms.Button  
        $Global:cancelButton.Top = $buttonPanel.Height - $Global:cancelButton.Height - 10; $Global:cancelButton.Left = $buttonPanel.Width - $Global:cancelButton.Width - 10 
        $Global:cancelButton.Text = "Close" 
        $Global:cancelButton.Anchor = "Right"
    $Global:buttonPanel.Controls.Add($Global:cancelButton) 
    ## Add the button panel to the form 
    $Global:form.Controls.Add($buttonPanel) 
    ## Set Default actions for the buttons 
    $Global:form.CancelButton = $Global:cancelButton      # ESCAPE = Cancel 
}

Function Update-LicenseDetails
{
    O365Licenses
    O365HasLicenses
    $Global:locListBox.Items.Clear()
    If ($Global:LicAssigned.Length -ne 0)
    {
        $LocArray = $Global:LicAssigned.split(",")
        $i=0   # Counter 
        foreach ($element in $LocArray)
        { # Loop through Azure list and add to listbox 
            [void] $Global:locListBox.Items.Add($element.TrimStart())  # Add element to listbox 
            $i ++ 
        } 
    }
    $Global:locListBox.Refresh()
    $Global:okButton.Text = "Next Emp"
    $Global:chkAssign.Visible = $false
    $Global:chkStd.Visible = $false
    $Global:chkReview.Visible = $false
    $Global:chkRemove.Visible = $false
    $Global:chkReset.Visible = $false
}

Function Update-AssignedLicenses
{
    write-host "In Update-AssignedLicenses"
    $Global:UPN = $Script:ENo + "@global.ul.com"
    $Global:Exists = [bool](get-MSOLUser -UserPrincipalName $Global:UPN -ErrorAction SilentlyContinue)
    $Global:Action = "Continue"
    $Global:UserDet = ""
    $Global:UserLicense = Get-MsolUser -UserPrincipalName $Global:UPN -ErrorAction SilentlyContinue
    $LDAPFilter = "(userPrincipalName=" + $Global:UPN + ")"
    $Global:ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,proxyaddresses,extensionattribute1,extensionattribute4
                
    If ($Global:Exists -eq $true)
    {
        O365Licenses
        $Global:lblUserInf.Visible = $false
#        $Global:txtUserInf.Visible = $false
        $Global:ButGetENo.visible = $false

        $Global:locListBox.Items.Clear()
        $LocArray = $Global:LicAssigned.split(",")
        $i=0   # Counter 
            foreach ($element in $LocArray) { # Loop through Azure list and add to listbox 
                [void] $Global:locListBox.Items.Add($element.TrimStart())  # Add element to listbox 
                $i ++ 
            } 
            $Global:locListBox.Visible = $true
#            $Global:lblLicDet.Visible = $true
    }
    else
    {
        $Output = $wshell.Popup("Invalid employee number.",0,"Invalid Employee Number",0+32)
        $Global:txtUserInf.Text = ""
    }
     Build-AddActionButtons
}


Add-Type -Assembly System.Windows.Forms     ## Load the Windows Forms assembly 
## Create the main form

$LogDirectory   = "E:\Automation\Licensing\Log"
$LogFile		= $LogDirectory + "\" + "Log-Licensing.log"
$wshell = New-Object -ComObject Wscript.Shell
invoke-expression -Command E:\O365AdminShared\Scripts\O365DisabledLicenseFeatures.ps1
$Global:Result = ""

write-host "`nGathering Licensing Details...." -ForegroundColor Cyan

Build-LicenseStatsForm
#Build-AddUserDetailsForm
Add-FormStandardButtons
$Global:form.Add_Shown( { $form.Activate(); $Global:txtUserInf.Focus()} )
$Global:result = $Global:form.ShowDialog()          ## Show the form, and wait for the response
O365HasLicenses

Do
{
If ($Global:Result -eq "OK")
{
    write-host "UserInf: " $Global:txtUserInf.Text
    $MgEmpNo = $Script:ENo + "@"
    $Script:MgUsr = Get-MgUser -Filter "startsWith(UserPrincipalName, '$MgEmpNo')"
    write-host "UserID: " $Script:MgUsr
    pause
    Write-host "Processing for: " $Script:ENo
    
    If ($Global:chkAssign.Checked -eq "Checked")
    {
   # This section handles licenses where one license needs to be removed before assigning another license
        If ($Global:chkE5Lic.Checked -eq "Checked")
        {
            Assign-E5Lic
        }
        If ($Global:chkP2Lic.Checked -eq "Checked")
        {
            Assign-P2Lic
        }
        If ($Global:chkPBIFLic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasBIFree
            $LicName = "PowerBI Free"
            $LicSKU = "ul:POWER_BI_STANDARD"
            Assign-License
        }
        If ($Global:chkPBIPLic.Checked -eq "Checked")
        {
            $LicName = "PowerBI Free"
            $LicSKU = "ul:POWER_BI_STANDARD"
            Assign-PowerBIProLic
        }
        #This section handles licenses where a license can just be assigned without impacting other licenses that might be assigned
        If ($Global:chkATPDefLic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasATPDef
            $LicName = "Windows Defender Advanced Threat Protection"
            $LicSKU = "ul:WIN_DEF_ATP"
            Assign-License
        }
        If ($Global:chkAudioConfLic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasAudioConf
            $LicName = "Audio Conferencing"
            $LicSKU = "ul:MCOMEETADV"
            Assign-License
        }
        If ($Global:chkAutoPUsrLic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasAutoPUsr
            $LicName = "Power Automate p/User Plan"
            $LicSKU = "ul:FLOW_PER_USER"
            Assign-License
        }
        If ($Global:chkAutoPUsrRPALic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasAutoPUsrRPA
            $LicName = "Power Automate p/User w/RPA Plan"
            $LicSKU = "ul:POWERAUTOMATE_ATTENDED_RPA"
            Assign-License
        }
        If ($Global:chkCAPLic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasCAP
            $LicName = "Common Area Phone"
            $LicSKU = "ul:MCOCAP"
            Assign-License
        }
        If ($Global:chkMeetingLic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasMeeting
            $LicName = "Microsoft Meeting Room"
            $LicSKU = "ul:MEETING_ROOM"
            Assign-License
        }
        If ($Global:chkAppsPUsrLic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasAppsPUsr
            $LicName = "PowerApps p/User"
            $LicSKU = "ul:POWERAPPS_PER_USER"
            Assign-License
        }
        If ($Global:chkPhoneSysLic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasPhone
            $LicName = "Phone System"
            $LicSKU = "ul:MCOEV"
            Assign-License
        }
        If ($Global:chkProjP3Lic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasProjP3
            $LicName = "Project Online Plan 3"
            $LicSKU = "ul:PROJECTPROFESSIONAL"
            Assign-License
        }
        If ($Global:chkProjP5Lic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasProjP5
            $LicName = "Project Online Plan 5"
            $LicSKU = "ul:PROJECTPREMIUM"
            Assign-License
        }
        If ($Global:chkVisioP1Lic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasVisioP1
            $LicName = "Visio Online Plan 1"
            $LicSKU = "ul:VISIOONLINE_PLAN1"
            Assign-License
        }
        If ($Global:chkVisioP2Lic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasVisioP2
            $LicName = "Visio Online Plan 2"
            $LicSKU = "ul:VISIOCLIENT"
            Assign-License
        }
        $Global:chkAssign.Checked = $false
        Update-AssignedLicenses
    }

    If ($Global:chkRemove.Checked -eq "Checked")
    {
        If (($Global:chkATPDefLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Windows Defender Advanced Threat Protection"))
        {
            $HasLicense = $Global:HasATPDef
            $LicName = "Windows Defender Advanced Threat Protection"
            $LicSKU = "ul:WIN_DEF_ATP"
            UnAssign-License
        }
        If (($Global:chkAudioConfLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Audio Conferencing"))
        {
            $HasLicense = $Global:HasAudioConf
            $LicName = "Audio Conferencing"
            $LicSKU = "ul:MCOMEETADV"
            UnAssign-License
        }
        If (($Global:chkAutoPUsrLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Power Automate p/User Plan"))
        {
            $HasLicense = $Global:HasAutoPUsr
            $LicName = "Power Automate p/User Plan"
            $LicSKU = "ul:FLOW_PER_USER"
            UnAssign-License
        }
        If (($Global:chkAutoPUsrRPALic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Power Automate p/User w/RPA Plan"))
        {
            $HasLicense = $Global:HasAutoPUsrRPA
            $LicName = "Power Automate p/User w/RPA Plan"
            $LicSKU = "ul:POWERAUTOMATE_ATTENDED_RPA"
            UnAssign-License
        }
        If (($Global:chkCAPLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Common Area Phone"))
        {
            $HasLicense = $Global:HasCAP
            $LicName = "Common Area Phone"
            $LicSKU = "ul:MCOCAP"
            UnAssign-License
        }
        If (($Global:chkE5Lic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Enterprise E5"))
        {
            $HasLicense = $Global:HasE5
            $LicName = "Enterpris E5"
            $LicSKU = "ul:ENTERPRISEPREMIUM"
            UnAssign-License
        }
        If (($Global:chkEMSLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Enterprise Mobility + Security E5"))
        {
            $HasLicense = $Global:HasEMS
            $LicName = "Enterprise Mobility + Security E5"
            $LicSKU = "ul:EMSPREMIUM"
            UnAssign-License
        }
        If (($Global:chkMeetingLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Microsoft Meeting Room"))
        {
            $HasLicense = $Global:HasMeeting
            $LicName = "Microsoft Meeting Room"
            $LicSKU = "ul:MEETING_ROOM"
            UnAssign-License             
        }
        If (($Global:chkP2Lic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Exchange Online Plan2"))
        {
            $HasLicense = $Global:HasE2
            $LicName = "Exchange Online Plan2"
            $LicSKU = "ul:EXCHANGEENTERPRISE"
            UnAssign-License
        }
        If ($Global:chkAppsPUsrLic.Checked -eq "Checked")
        {
            $HasLicense = $Global:HasAppsPUsr
            $LicName = "PowerApps p/User"
            $LicSKU = "ul:POWERAPPS_PER_USER"
            UnAssign-License
        }
        If (($Global:chkPBIFLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "PowerBI (Free)"))
        {
            If ($Global:HasBIPro -eq $True)
            {
                $LineToWrite = "REVI" + "`t" + "PowerBI Pro license assigned to this account removing PowerBI (Free) License" + $Global:UPN
#                Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:POWER_BI_STANDARD"
                $HasLicense = $Global:HasBIFree
                $LicName = "PowerBI Free"
                $LicSKU = "ul:POWER_BI_STANDARD"
                UnAssign-License
            }
            else
            {
  	            $Output = $wshell.Popup("The PowerBI (Free) license is a stanard license for all staff and cannot be removed.",0,"License Not Assigned",0+32)
                $LineToWrite = "REVI" + "`t" + "PowerBI (Free) License is a standard license and cannot be removed " + $Global:UPN
            }
            WriteLogEvent
        }
        If (($Global:chkPBIPLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "PowerBI Pro"))
        {
            $HasLicense = $Global:HasBIPro
            $LicName = "PowerBI Pro"
            $LicSKU = "ul:POWER_BI_PRO"
            UnAssign-License
            if ($Global:HasBIFree -ne "True")
            {
      	        $LineToWrite = "REVI" + "`t" + "Assigning PowerBI Free License to " + $Global:UPN
                WriteLogEvent
                $EmpNo = $Script:ENo
                $Script:MgUsr = Get-MgUser -Filter "startsWith(UserPrincipalName, '$MgEmpNo')"
#                $Script:MgUsr = Get-MgUser -Filter "startsWith(UserPrincipalName, '$EmpNo')"
                Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:POWER_BI_STANDARD"
            }
        }
        If (($Global:chkPhoneSysLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Phone System"))
        {
            $HasLicense = $Global:HasPhone
            $LicName = "Phone System"
            $LicSKU = "ul:MCOEV"
            UnAssign-License
        }
        If (($Global:chkProjP3Lic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Project Online Plan 3"))
        {
            $HasLicense = $Global:HasProjP3
            $LicName = "Project Online Plan 3"
            $LicSKU = "ul:PROJECTPROFESSIONAL"
            UnAssign-License
        }
        If (($Global:chkProjP5Lic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Project Online Plan 5"))
        {
            $HasLicense = $Global:HasProjP5
            $LicName = "Project Online Plan 5"
            $LicSKU = "ul:PROJECTPREMIUM"
            UnAssign-License
        }
        If (($Global:chkVisioP1Lic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Visio Online Plan 1"))
        {
            $HasLicense = $Global:HasVisioP1
            $LicName = "Visio Online Plan 1"
            $LicSKU = "ul:VISIOONLINE_PLAN1"
            UnAssign-License
        }
        If (($Global:chkVisioP2Lic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Visio Online Plan 2"))
        {
            $HasLicense = $Global:HasVisioP2
            $LicName = "Visio Online Plan 2"
            $LicSKU = "ul:VISIOCLIENT"
            UnAssign-License
        }
         $Global:chkRemove.Checked = $false
    }
    
    If ($Global:chkStd.Checked -eq "Checked")
    {
        $usr = get-msoluser -UserPrincipalName $Global:UPN
        $EmpGrp = [bool](Get-AzureADGroupMember -ObjectId 2aab81de-e54a-4c58-909c-b15273c54dff -all $true |Where-Object {$_.ObjectID -eq $usr.ObjectID})
        $NonEmpGrp = [bool](Get-AzureADGroupMember -ObjectId 40dff561-964a-4f7b-9794-dba3542b63c4 -all $true |Where-Object {$_.ObjectID -eq $usr.ObjectID})
        $AudioGrpMem = [bool](Get-AzureADGroupMember -ObjectId a9b06203-849c-40c7-93f9-6c404fdcf6d3 -all $true |Where-Object {$_.ObjectID -eq $usr.ObjectID})
        $PowerBIProGrpMem = [book](Get-AzureADGroupMember -ObjectId 4246671a-2ac5-4b8a-96fd-4e4b7d817791 -all $true |Where-Object {$_.ObjectID -eq $usr.ObjectID})

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
                Write-host "     No E5 Licenses Assigned to this individual" -ForegroundColor Red
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
        $Global:chkStd.Checked = $false
    }

    If ($Global:chkReview.Checked -eq "Checked")
    {
        If (($Global:chkE5Lic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Enterprise E5"))
        {
            EnabledE5Feature
            pause
        }
        $Global:chkReview.Checked = $false
    }

    If ($Global:chkReset.Checked -eq "Checked")
    {
        $SSKID = "ul:ENTERPRISEPACK"
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
        $Global:chkReset.Checked = $false
    }

    If ($Global:Exists -eq $true)
    {
        write-host "Getting Updated license assignment" -ForegroundColor Yellow
        start-sleep -Seconds 7
        $Global:UserLicense = Get-MsolUser -UserPrincipalName $Global:UPN
        Update-AssignedLicenses
#        Update-LicenseDetails
    }
#    ## Finalize Form and Show Dialog
    $Global:form.Add_Shown( { $form.Activate(); $okButton.Focus() } )  #Activate and Set Focus
    $Global:result = $Global:form.ShowDialog()          ## Show the form, and wait for the response
    
    If ($Global:Result -eq "OK")
    {
        Build-LicenseStatsForm
#        Build-AddUserDetailsForm
        Add-FormStandardButtons
        $Global:cancelButton.Text = "Cancel"
        $Global:form.Add_Shown( { $form.Activate(); $Global:txtUserInf.Focus()} )
        $Global:result = $Global:form.ShowDialog()          ## Show the form, and wait for the response
        O365HasLicenses
   }
}
}while ($Global:Result -eq "OK")