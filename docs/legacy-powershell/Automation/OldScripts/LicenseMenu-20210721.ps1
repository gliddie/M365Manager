<#   
================================================================================ 
 Name: Licenses Form
 ================================================================================ 
 #
 # 11/18/2020 - SAG - Added missing code for unassigning licenses
 # 12/16/2020 - SAG - Added Project Plan 5 licenses
 # 05/08/2021 - SAG - Modified how the employee number details is obtained.
 # 07/06/2021 - SAG - Added Visio Plan2 license
#>  

Function AvailLicense
{
    $E3Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}
    $P2Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EXCHANGEENTERPRISE"}
    $EMSLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EMS"}
    $ATPP1Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ATP_ENTERPRISE"}
    $CAPLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOCAP"}
    $CRMLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:DYN365_ENTERPRISE_PLAN1"}
    $PBIFLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWER_BI_STANDARD"}
    $PBIPLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWER_BI_PRO"}
    $MRSSLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"}
    $FP2Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:FLOW_P2"}
    $ATPDefLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:WIN_DEF_ATP"}
    $ATPDefMAC = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MDATP_XPLAT"}
    $MeetingLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MEETING_ROOM"}
    $PAppsP2Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWERFLOW_P2"}
    $PhoneSysLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOEV"}
    $AudioConfLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOMEETADV"}
    $RmtAssistLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MICROSOFT_REMOTE_ASSIST"}
    $InfoBarrLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:M365_INSIDER_RISK_MANAGEMENT"}
    $ProjP5Lic = = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:PROJECTPREMIUM"}
    $VisioP1Lic = = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:VISIOONLINE_PLAN1"}
    $VisioP2Lic = = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:VISIOCLIENT"}
         
#    Send O365 Team email that the Staff Type Changes is complete

    write-host "          Current license allocations:"
    write-host "`nActive`t InUse`t Avail`t-  License Name"
    write-host "-------------------------------------------------------------"

    write-host $E3Lic.ActiveUnits "`t" $E3Lic.ConsumedUnits "`t" ($E3Lic.ActiveUnits-$E3Lic.ConsumedUnits) "`t-  Enterprise E3 License"
    write-host $P2Lic.ActiveUnits "`t" $P2Lic.ConsumedUnits "`t" ($P2Lic.ActiveUnits-$P2Lic.ConsumedUnits) "`t-  ExchangeOnline P2 License"
    write-host $EMSLic.ActiveUnits "`t" $EMSLic.ConsumedUnits "`t" ($EMSLic.ActiveUnits-$EMSLic.ConsumedUnits) "`t-  Enterprise Mobility + Security E3 License"
    write-host $ATPP1Lic.ActiveUnits "`t" $ATPP1Lic.ConsumedUnits "`t" ($ATPP1Lic.ActiveUnits-$ATPP1Lic.ConsumedUnits) "`t-  Advanced Threat Protection Plan1 License"
    write-host $ATPDefLic.ActiveUnits "`t" $ATPDefLic.ConsumedUnits "`t" ($ATPDefLic.ActiveUnits-$ATPDefLic.ConsumedUnits) "`t-  Defender Advanced Threat Protection License"
    write-host $ATPDefMAC.ActiveUnits "`t" $ATPDefMAC.ConsumedUnits "`t" ($ATPDefMAC.ActiveUnits-$ATPDefMAC.ConsumedUnits) "`t-  Advanced Threat Protection for MAC License"
    write-host $CAPLic.ActiveUnits "`t" $CAPLic.ConsumedUnits "`t" ($CAPLic.ActiveUnits-$CAPLic.ConsumedUnits) "`t-  Common Area Phone License"
    write-host $CRMLic.ActiveUnits "`t" $CRMLic.ConsumedUnits "`t" ($CRMLic.ActiveUnits-$CRMLic.ConsumedUnits) "`t-  Dynamics 365 Customer Engagement Plan License"
    write-host $MRSSLic.ActiveUnits "`t" $MRSSLic.ConsumedUnits "`t" ($MRSSLic.ActiveUnits-$MRSSLic.ConsumedUnits) "`t-  Microsoft Relationship Sales solution License"
    write-host "N/A`t" $PBIFLic.ConsumedUnits "`t N/A`t-  PowerBI (Free) License"
    write-host $PBIPLic.ActiveUnits "`t" $PBIPLic.ConsumedUnits "`t" ($PBIPLic.ActiveUnits-$PBIPLic.ConsumedUnits) "`t-  PowerBI Pro License"
    write-host $FP2Lic.ActiveUnits "`t" $FP2Lic.ConsumedUnits "`t" ($FP2Lic.ActiveUnits-$FP2Lic.ConsumedUnits) "`t-  Microsoft Flow (Plan2) License"
    write-host $MeetingLic.ActiveUnits "`t" $MeetingLic.ConsumedUnits "`t" ($MeetingLic.ActiveUnits-$MeetingLic.ConsumedUnits) "`t-  Meeting Room License"
    write-host $PAppsP2Lic.ActiveUnits "`t" $PAppsP2Lic.ConsumedUnits "`t" ($PAppsP2Lic.ActiveUnits-$PAppsP2Lic.ConsumedUnits) "`t-  PowerApps Plan 2 License"
    write-host $PhoneSysLic.ActiveUnits "`t" $PhoneSysLic.ConsumedUnits "`t" ($PhoneSysLic.ActiveUnits-$PhoneSysLic.ConsumedUnits) "`t-  Phone System License"
    write-host $AudioConfLic.ActiveUnits "`t" $AudioConfLic.ConsumedUnits "`t" ($AudioConfLic.ActiveUnits-$AudioConfLic.ConsumedUnits) "`t-  Audio Conferencing License"
    write-host $RmtAssistLic.ActiveUnits "`t" $RmtAssistLic.ConsumedUnits "`t" ($RmtAssistLic.ActiveUnits-$RmtAssistLic.ConsumedUnits) "`t-  Microsoft Remote Assist License"
    write-host $InfoBarrLic.ActiveUnits "`t" $InfoBarrLic.ConsumedUnits "`t" ($InfoBarrLic.ActiveUnits-$InfoBarrLic.ConsumedUnits) "`t-  E5 Insider Risk Management License"
    write-host $ProjP5Lic.ActiveUnits "`t" $ProjP5Lic.ConsumedUnits "`t" ($ProjP5Lic.ActiveUnits-$ProjP5Lic.ConsumedUnits) "`t-  Project Plan 5 License"
    write-host $VisioP1Lic.ActiveUnits "`t" $VisioP1Lic.ConsumedUnits "`t" ($VisioP1Lic.ActiveUnits-$VisioP1Lic.ConsumedUnits) "`t-  Visio Plan 1 License"
    write-host $VisioP2Lic.ActiveUnits "`t" $VisioP2Lic.ConsumedUnits "`t" ($VisioP2Lic.ActiveUnits-$VisioP2Lic.ConsumedUnits) "`t-  Visio Plan 2 License"
}

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
        $Global:lblUserInf.Visible = $false
        $Global:txtUserInf.Visible = $false
        $Global:ButGetENo.visible = $false
        $TopLoc = 470
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
        $TopLoc = $TopLoc + 20
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
    $Global:form.Size = New-Object System.Drawing.Size(700,700) #(W,H)

    ## Label and TextBox  
    ## Title Line
    $Global:lblTitleLine = New-Object System.Windows.Forms.Label   
        $lblTitleLine.Text = "Active           InUse          Avail           License Name"
        $lblTitleLine.Top = 15 ; $lblTitleLine.Left = 30; $lblTitleLine.Width=120 ;$lblTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine)    # Add to Form 

    ## E3 Licenses
    $E3Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}
    $TopLoc = $TopLoc + 20
    $Global:chkE3Lic = New-Object Windows.Forms.checkbox 
        $Global:chkE3Lic.Left = $LicLoc; $Global:chkE3Lic.Width = 280; $Global:chkE3Lic.Top = ($TopLoc-5)  
        $Global:chkE3Lic.Text = "Enterprise E3 License" 
        $Global:chkE3Lic.Checked = $false   # set a default value 
        $Global:chkE3Lic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkE3Lic) 
    $Global:lblE3Act = New-Object System.Windows.Forms.Label
        $lblE3Act.Text = $E3Lic.ActiveUnits
        $lblE3Act.Top = $TopLoc ; $lblE3Act.Left = $ActLoc; $lblE3Act.Width=10 ;$lblE3Act.AutoSize = $true
        $Global:form.Controls.Add($Global:lblE3Act)    # Add to Form
    $Global:lblE3InU = New-Object System.Windows.Forms.Label
        $lblE3InU.Text = $E3Lic.ConsumedUnits
        $lblE3InU.Top = $TopLoc ; $lblE3InU.Left = $InULoc; $lblE3InU.Width=10 ;$lblE3InU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblE3InU)    # Add to Form
    $Global:lblE3Avl = New-Object System.Windows.Forms.Label
        $lblE3Avl.Text = ($E3Lic.ActiveUnits-$E3Lic.ConsumedUnits)
        $lblE3Avl.Top = $TopLoc ; $lblE3Avl.Left = $AvlLoc; $lblE3Avl.Width=10 ;$lblE3Avl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblE3Avl)    # Add to Form
 
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

    ## EMS Licenses
    $EMSLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EMS"}
    $TopLoc = $TopLoc + 20
    $Global:chkEMSLic = New-Object Windows.Forms.checkbox 
        $Global:chkEMSLic.Left = $LicLoc; $Global:chkEMSLic.Width = 280; $Global:chkEMSLic.Top = ($TopLoc-5)  
        $Global:chkEMSLic.Text = "Enterprise Mobility + Security E3 License" 
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

    ## ATPP1 Licenses
    $ATPP1Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ATP_ENTERPRISE"}
    $TopLoc = $TopLoc + 20
    $Global:chkATPP1Lic = New-Object Windows.Forms.checkbox 
        $Global:chkATPP1Lic.Left = $LicLoc; $Global:chkATPP1Lic.Width = 280; $Global:chkATPP1Lic.Top = ($TopLoc-5)  
        $Global:chkATPP1Lic.Text = "Advanced Threat Protection Plan1 License" 
        $Global:chkATPP1Lic.Checked = $false   # set a default value 
        $Global:chkATPP1Lic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkATPP1Lic)
    $Global:lblATPP1Act = New-Object System.Windows.Forms.Label
        $lblATPP1Act.Text = $ATPP1Lic.ActiveUnits
        $lblATPP1Act.Top = $TopLoc ; $lblATPP1Act.Left = $ActLoc; $lblATPP1Act.Width=10 ;$lblATPP1Act.AutoSize = $true
        $Global:form.Controls.Add($Global:lblATPP1Act)    # Add to Form
    $Global:lblATPP1InU = New-Object System.Windows.Forms.Label
        $lblATPP1InU.Text = $ATPP1Lic.ConsumedUnits
        $lblATPP1InU.Top = $TopLoc ; $lblATPP1InU.Left = $InULoc; $lblATPP1InU.Width=10 ;$lblATPP1InU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblATPP1InU)    # Add to Form
    $Global:lblATPP1Avl = New-Object System.Windows.Forms.Label
        $lblATPP1Avl.Text = ($ATPP1Lic.ActiveUnits-$ATPP1Lic.ConsumedUnits)
        $lblATPP1Avl.Top = $TopLoc ; $lblATPP1Avl.Left = $AvlLoc; $lblATPP1Avl.Width=10 ;$lblATPP1Avl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblATPP1Avl)    # Add to Form

     ## ATPDefMAC Licenses
    $ATPDefMACLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MDATP_XPLAT"}
    $TopLoc = $TopLoc + 20
    $Global:chkATPDefMACLic = New-Object Windows.Forms.checkbox 
        $Global:chkATPDefMACLic.Left = $LicLoc; $Global:chkATPDefMACLic.Width = 280; $Global:chkATPDefMACLic.Top = ($TopLoc-5)  
        $Global:chkATPDefMACLic.Text = "Advanced Theat Protection for MAC License" 
        $Global:chkATPDefMACLic.Checked = $false   # set a default value 
        $Global:chkATPDefMACLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkATPDefMACLic)
    $Global:lblATPDefMACAct = New-Object System.Windows.Forms.Label
        $lblATPDefMACAct.Text = $ATPDefMACLic.ActiveUnits
        $lblATPDefMACAct.Top = $TopLoc ; $lblATPDefMACAct.Left = $ActLoc; $lblATPDefMACAct.Width=10 ;$lblATPDefMACAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblATPDefMACAct)    # Add to Form
    $Global:lblATPDefMACInU = New-Object System.Windows.Forms.Label
        $lblATPDefMACInU.Text = $ATPDefMACLic.ConsumedUnits
        $lblATPDefMACInU.Top = $TopLoc ; $lblATPDefMACInU.Left = $InULoc; $lblATPDefMACInU.Width=10 ;$lblATPDefMACInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblATPDefMACInU)    # Add to Form
    $Global:lblATPDefMACAvl = New-Object System.Windows.Forms.Label
        $lblATPDefMACAvl.Text = ($ATPDefMACLic.ActiveUnits-$ATPDefMACLic.ConsumedUnits)
        $lblATPDefMACAvl.Top = $TopLoc ; $lblATPDefMACAvl.Left = $AvlLoc; $lblATPDefMACAvl.Width=10 ;$lblATPDefMACAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblATPDefMACAvl)    # Add to Form

    ## AudioConf Licenses
    $AudioConfLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOMEETADV"}
    $TopLoc = $TopLoc + 20
    $Global:chkAudioConfLic = New-Object Windows.Forms.checkbox 
        $Global:chkAudioConfLic.Left = $LicLoc; $Global:chkAudioConfLic.Width = 280; $Global:chkAudioConfLic.Top = ($TopLoc-5)  
        $Global:chkAudioConfLic.Text = "Audio Conferencing License" 
        $Global:chkAudioConfLic.Checked = $false   # set a default value 
        $Global:chkAudioConfLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkAudioConfLic)
    $Global:lblAudioConfAct = New-Object System.Windows.Forms.Label
        $lblAudioConfAct.Text = $AudioConfLic.ActiveUnits
        $lblAudioConfAct.Top = $TopLoc ; $lblAudioConfAct.Left = $ActLoc; $lblAudioConfAct.Width=10 ;$lblAudioConfAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAudioConfAct)    # Add to Form
    $Global:lblAudioConfInU = New-Object System.Windows.Forms.Label
        $lblAudioConfInU.Text = $AudioConfLic.ConsumedUnits
        $lblAudioConfInU.Top = $TopLoc ; $lblAudioConfInU.Left = $InULoc; $lblAudioConfInU.Width=10 ;$lblAudioConfInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAudioConfInU)    # Add to Form
    $Global:lblAudioConfAvl = New-Object System.Windows.Forms.Label
        $lblAudioConfAvl.Text = ($AudioConfLic.ActiveUnits-$AudioConfLic.ConsumedUnits)
        $lblAudioConfAvl.Top = $TopLoc ; $lblAudioConfAvl.Left = $AvlLoc; $lblAudioConfAvl.Width=10 ;$lblAudioConfAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAudioConfAvl)    # Add to Form

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

    ## CRM Licenses
    $CRMLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:DYN365_ENTERPRISE_PLAN1"}
    $TopLoc = $TopLoc + 20
    $Global:chkCRMLic = New-Object Windows.Forms.checkbox 
        $Global:chkCRMLic.Left = $LicLoc; $Global:chkCRMLic.Width = 290; $Global:chkCRMLic.Top = ($TopLoc-5)  
        $Global:chkCRMLic.Text = "Dynamics 365 Customer Engagement Plan License" 
        $Global:chkCRMLic.Checked = $false   # set a default value 
        $Global:chkCRMLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkCRMLic)
    $Global:lblCRMAct = New-Object System.Windows.Forms.Label
        $lblCRMAct.Text = $CRMLic.ActiveUnits
        $lblCRMAct.Top = $TopLoc ; $lblCRMAct.Left = $ActLoc; $lblCRMAct.Width=10 ;$lblCRMAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblCRMAct)    # Add to Form
    $Global:lblCRMInU = New-Object System.Windows.Forms.Label
        $lblCRMInU.Text = $CRMLic.ConsumedUnits
        $lblCRMInU.Top = $TopLoc ; $lblCRMInU.Left = $InULoc; $lblCRMInU.Width=10 ;$lblCRMInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblCRMInU)    # Add to Form
    $Global:lblCRMAvl = New-Object System.Windows.Forms.Label
        $lblCRMAvl.Text = ($CRMLic.ActiveUnits-$CRMLic.ConsumedUnits)
        $lblCRMAvl.Top = $TopLoc ; $lblCRMAvl.Left = $AvlLoc; $lblCRMAvl.Width=10 ;$lblCRMAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblCRMAvl)    # Add to Form

    ## FlowP2 Licenses    
    $FP2Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:FLOW_P2"}
    $TopLoc = $TopLoc + 20
    $Global:chkFP2Lic = New-Object Windows.Forms.checkbox 
        $Global:chkFP2Lic.Left = $LicLoc; $Global:chkFP2Lic.Width = 280; $Global:chkFP2Lic.Top = ($TopLoc-5)  
        $Global:chkFP2Lic.Text = "Microsoft Flow (Plan2) License" 
        $Global:chkFP2Lic.Checked = $false   # set a default value 
        $Global:chkFP2Lic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkFP2Lic)          
    $Global:lblFP2Act = New-Object System.Windows.Forms.Label
        $lblFP2Act.Text = $FP2Lic.ActiveUnits
        $lblFP2Act.Top = $TopLoc ; $lblFP2Act.Left = $ActLoc; $lblFP2Act.Width=10 ;$lblFP2Act.AutoSize = $true
        $Global:form.Controls.Add($Global:lblFP2Act)    # Add to Form
    $Global:lblFP2InU = New-Object System.Windows.Forms.Label
        $lblFP2InU.Text = $FP2Lic.ConsumedUnits
        $lblFP2InU.Top = $TopLoc ; $lblFP2InU.Left = $InULoc; $lblFP2InU.Width=10 ;$lblFP2InU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblFP2InU)    # Add to Form
    $Global:lblFP2Avl = New-Object System.Windows.Forms.Label
        $lblFP2Avl.Text = ($FP2Lic.ActiveUnits-$FP2Lic.ConsumedUnits)
        $lblFP2Avl.Top = $TopLoc ; $lblFP2Avl.Left = $AvlLoc; $lblFP2Avl.Width=10 ;$lblFP2Avl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblFP2Avl)    # Add to Form

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

    ## MRSS Licenses
    $MRSSLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"}
    $TopLoc = $TopLoc + 20
    $Global:chkMRSSLic = New-Object Windows.Forms.checkbox 
        $Global:chkMRSSLic.Left = $LicLoc; $Global:chkMRSSLic.Width = 280; $Global:chkMRSSLic.Top = ($TopLoc-5)  
        $Global:chkMRSSLic.Text = "Microsoft Relationship Sales Solution License" 
        $Global:chkMRSSLic.Checked = $false   # set a default value 
        $Global:chkMRSSLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkMRSSLic)
    $Global:lblMRSSAct = New-Object System.Windows.Forms.Label
        $lblMRSSAct.Text = $MRSSLic.ActiveUnits
        $lblMRSSAct.Top = $TopLoc ; $lblMRSSAct.Left = $ActLoc; $lblMRSSAct.Width=10 ;$lblMRSSAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblMRSSAct)    # Add to Form
    $Global:lblMRSSInU = New-Object System.Windows.Forms.Label
        $lblMRSSInU.Text = $MRSSLic.ConsumedUnits
        $lblMRSSInU.Top = $TopLoc ; $lblMRSSInU.Left = $InULoc; $lblMRSSInU.Width=10 ;$lblMRSSInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblMRSSInU)    # Add to Form
    $Global:lblMRSSAvl = New-Object System.Windows.Forms.Label
        $lblMRSSAvl.Text = ($MRSSLic.ActiveUnits-$MRSSLic.ConsumedUnits)
        $lblMRSSAvl.Top = $TopLoc ; $lblMRSSAvl.Left = $AvlLoc; $lblMRSSAvl.Width=10 ;$lblMRSSAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblMRSSAvl)    # Add to Form

    ## RemoteAssist Licenses
    $RmtAssistLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MICROSOFT_REMOTE_ASSIST"}
    $TopLoc = $TopLoc + 20
    $Global:chkRmtAssistLic = New-Object Windows.Forms.checkbox 
        $Global:chkRmtAssistLic.Left = $LicLoc; $Global:chkRmtAssistLic.Width = 280; $Global:chkRmtAssistLic.Top = ($TopLoc-5)  
        $Global:chkRmtAssistLic.Text = "Microsoft Remote Assist License" 
        $Global:chkRmtAssistLic.Checked = $false   # set a default value 
        $Global:chkRmtAssistLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkRmtAssistLic)
    $Global:lblRmtAssistAct = New-Object System.Windows.Forms.Label
        $lblRmtAssistAct.Text = $RmtAssistLic.ActiveUnits
        $lblRmtAssistAct.Top = $TopLoc ; $lblRmtAssistAct.Left = $ActLoc; $lblRmtAssistAct.Width=10 ;$lblRmtAssistAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblRmtAssistAct)    # Add to Form
    $Global:lblRmtAssistInU = New-Object System.Windows.Forms.Label
        $lblRmtAssistInU.Text = $RmtAssistLic.ConsumedUnits
        $lblRmtAssistInU.Top = $TopLoc ; $lblRmtAssistInU.Left = $InULoc; $lblRmtAssistInU.Width=10 ;$lblRmtAssistInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblRmtAssistInU)    # Add to Form
    $Global:lblRmtAssistAvl = New-Object System.Windows.Forms.Label
        $lblRmtAssistAvl.Text = ($RmtAssistLic.ActiveUnits-$RmtAssistLic.ConsumedUnits)
        $lblRmtAssistAvl.Top = $TopLoc ; $lblRmtAssistAvl.Left = $AvlLoc; $lblRmtAssistAvl.Width=10 ;$lblRmtAssistAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblRmtAssistAvl)    # Add to Form

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

    ## PowerAppsP2 Licenses
    $PAppsP2Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWERFLOW_P2"}
    $TopLoc = $TopLoc + 20
    $Global:chkPAppsP2Lic = New-Object Windows.Forms.checkbox 
        $Global:chkPAppsP2Lic.Left = $LicLoc; $Global:chkPAppsP2Lic.Width = 280; $Global:chkPAppsP2Lic.Top = ($TopLoc-5)  
        $Global:chkPAppsP2Lic.Text = "PowerApps Plan 2 License" 
        $Global:chkPAppsP2Lic.Checked = $false   # set a default value 
        $Global:chkPAppsP2Lic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkPAppsP2Lic) 
    $Global:lblPAppsP2Act = New-Object System.Windows.Forms.Label
        $lblPAppsP2Act.Text = $PAppsP2Lic.ActiveUnits
        $lblPAppsP2Act.Top = $TopLoc ; $lblPAppsP2Act.Left = $ActLoc; $lblPAppsP2Act.Width=10 ;$lblPAppsP2Act.AutoSize = $true
        $Global:form.Controls.Add($Global:lblPAppsP2Act)    # Add to Form
    $Global:lblPAppsP2InU = New-Object System.Windows.Forms.Label
        $lblPAppsP2InU.Text = $PAppsP2Lic.ConsumedUnits
        $lblPAppsP2InU.Top = $TopLoc ; $lblPAppsP2InU.Left = $InULoc; $lblPAppsP2InU.Width=10 ;$lblPAppsP2InU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblPAppsP2InU)    # Add to Form
    $Global:lblPAppsP2Avl = New-Object System.Windows.Forms.Label
        $lblPAppsP2Avl.Text = ($PAppsP2Lic.ActiveUnits-$PAppsP2Lic.ConsumedUnits)
        $lblPAppsP2Avl.Top = $TopLoc ; $lblPAppsP2Avl.Left = $AvlLoc; $lblPAppsP2Avl.Width=10 ;$lblPAppsP2Avl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblPAppsP2Avl)    # Add to Form

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

    ## PowerBI Pro Licenses
    $PBIPLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWER_BI_PRO"}
    $TopLoc = $TopLoc + 20
    $Global:chkPBIPLic = New-Object Windows.Forms.checkbox 
        $Global:chkPBIPLic.Left = $LicLoc; $Global:chkPBIPLic.Width = 280; $Global:chkPBIPLic.Top = ($TopLoc-5)  
        $Global:chkPBIPLic.Text = "PowerBI Pro License" 
        $Global:chkPBIPLic.Checked = $false   # set a default value 
        $Global:chkPBIPLic.TabIndex = 2
        $Global:form.Controls.Add($Global:chkPBIPLic)         
    $Global:lblPBIPAct = New-Object System.Windows.Forms.Label
        $lblPBIPAct.Text = $PBIPLic.ActiveUnits
        $lblPBIPAct.Top = $TopLoc ; $lblPBIPAct.Left = $ActLoc; $lblPBIPAct.Width=10 ;$lblPBIPAct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblPBIPAct)    # Add to Form
    $Global:lblPBIPInU = New-Object System.Windows.Forms.Label
        $lblPBIPInU.Text = $PBIPLic.ConsumedUnits
        $lblPBIPInU.Top = $TopLoc ; $lblPBIPInU.Left = $InULoc; $lblPBIPInU.Width=10 ;$lblPBIPInU.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblPBIPInU)    # Add to Form
    $Global:lblPBIPAvl = New-Object System.Windows.Forms.Label
        $lblPBIPAvl.Text = ($PBIPLic.ActiveUnits-$PBIPLic.ConsumedUnits)
        $lblPBIPAvl.Top = $TopLoc ; $lblPBIPAvl.Left = $AvlLoc; $lblPBIPAvl.Width=10 ;$lblPBIPAvl.AutoSize = $true
        $Global:form.Controls.Add($Global:lblPBIPAvl)    # Add to Form

    ## Project Plan 3 Licenses
    $ProjP3Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:PROJECT_PLAN3_DEPT"}
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
}

function Build-AddUserDetailsForm
{
    $TopLoc = 470
    ## User Details
    $Global:lblUserInf = New-Object System.Windows.Forms.Label
        $Global:lblUserInf.Text = "Employee Number:"
        $Global:lblUserInf.Top = $TopLoc; $Global:lblUserInf.Left=10; $Global:lblUserInf.Width=10 ;$Global:lblUserInf.AutoSize = $true
        $Global:form.Controls.Add($Global:lblUserInf)    # Add to Form
    $Global:txtUserInf = New-Object System.Windows.Forms.TextBox
        $Global:txtUserInf.Top = $TopLoc; $Global:txtUserInf.Left=120; $Global:txtUserInf.Width=100 ;$Global:txtUserInf.AutoSize = $true
        $Global:form.Controls.Add($Global:txtUserInf)    # Add to Form
        $Global:InputFocus = $Global:txtUserInf

        $Global:ButGetENo = New-Object Windows.Forms.Button
            $Global:ButGetENo.Location = New-object System.Drawing.Size(230,$TopLoc)
            $Global:ButGetENo.Size = new-Object System.Drawing.Size(150,20)
            $Global:ButGetENo.Text = "Get Employee Details"
            $Global:form.Controls.Add($Global:ButGetENo)
            $Global:ButGetENo.Add_Click({Add-AssignedLicenses})
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
    $Global:cancelButton.Text = "Next Emp"
}

Function Update-AssignedLicenses
{

write-host "running Update-AssignedLicenses"
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
        $Global:lblUserInf.Visible = $false
        $Global:txtUserInf.Visible = $false
        $Global:ButGetENo.visible = $false

        $Global:locListBox.Items.Clear()
        $LocArray = $Global:LicAssigned.split(",")
        $i=0   # Counter 
            foreach ($element in $LocArray) { # Loop through Azure list and add to listbox 
                [void] $Global:locListBox.Items.Add($element.TrimStart())  # Add element to listbox 
                $i ++ 
            } 
            $Global:locListBox.Visible = $true
            $Global:lblLicDet.Visible = $true

#                $Global:form.Controls.Add($Global:locListBox) #Add listbox to form 
                # Obtain Value with: $Global:locListBox.SelectedItem
#                Build-AddActionButtons
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

<#Do
{
    If ($Global:Result -ne "OK")
    {
#>
        Build-LicenseStatsForm
        Build-AddUserDetailsForm
        Add-FormStandardButtons
        $Global:form.Add_Shown( { $form.Activate(); $Global:txtUserInf.Focus()} )
        $Global:result = $Global:form.ShowDialog()          ## Show the form, and wait for the response
        O365HasLicenses
#   }
#$Global:ButGetENo.Add_Click({Update-AssignedLicenses})

Do
{
#    If ($Global:Result -eq "OK")
#    {
#        O365HasLicenses
        If ($Global:chkAssign.Checked -eq "Checked")
        {
            If ($Global:chkAudioConfLic.Checked -eq "Checked")
            {
                Assign-AudioConferencingLic
            }

            If ($Global:chkE3Lic.Checked -eq "Checked")
            {
                Assign-E3Lic
            }

            If ($Global:chkP2Lic.Checked -eq "Checked")
            {
                Assign-P2Lic
            }

            If ($Global:chkATPDefLic.Checked -eq "Checked")
            {
                Assign-ATPDefLic
            }

            If ($Global:chkATPP1Lic.Checked -eq "Checked")
            {
                Assign-ATPP1Lic
            }

            If ($Global:chkATPDefMACLic.Checked -eq "Checked")
            {
                Assign-ATPPDefMACLic
            }

            If ($Global:chkCAPLic.Checked -eq "Checked")
            {
                Assign-CAPLic
            }

            If ($Global:chkCRMLic.Checked -eq "Checked")
            {
                Assign-CRMLic
            }

            If ($Global:chkEMSLic.Checked -eq "Checked")
            {            
                Assign-EMSLic
            }

            If ($Global:chkMeetingLic.Checked -eq "Checked")
            {
                Assign-MeetingLic
            }

            If ($Global:chkMRSSLic.Checked -eq "Checked")
            { 
                Assign-MRSSLic
            }

            If ($Global:chkP2Lic.Checked -eq "Checked")
            { 
                Assign-P2Lic
            }

            If ($Global:chkPAppsP2Lic.Checked -eq "Checked")
            {
                Assign-PAppsP2Lic
            }

            If ($Global:chkPBIFLic.Checked -eq "Checked")
            {
                Assign-PowerBIFLic
            }

            If ($Global:chkPBIPLic.Checked -eq "Checked")
            {
                Assign-PowerBIProLic
            }

            If ($Global:chkPhoneSysLic.Checked -eq "Checked")
            {
                Assign-PhoneSystemLic
            }
            
            If ($Global:chkProjP3Lic.Checked -eq "Checked")
            {
                Assign-ProjP3Lic
            }

            If ($Global:chkProjP5Lic.Checked -eq "Checked")
            {
                Assign-ProjP5Lic
            }

            If ($Global:chkVisioP1Lic.Checked -eq "Checked")
            {
                Assign-VisioP1Lic
            }

            If ($Global:chkVisioP2Lic.Checked -eq "Checked")
            {
                Assign-VisioP2Lic
            }

            If ($Global:chkRmtAssistLic.Checked -eq "Checked")
            {
                Assign-RemoteAsstLic
            }
            $Global:chkAssign.Checked = $false
        }

        If ($Global:chkRemove.Checked -eq "Checked")
        {
            $Global:chkRemove.Checked = $false
            If (($Global:chkE3Lic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Enterprise E3"))
            {
                UnAssign-E3Lic
            }
			
            If (($Global:chkP2Lic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Exchange Online Plan2"))
            {
                UnAssign-P2Lic
            }			
			
            If (($Global:chkEMSLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Enterprise Mobility Suite/Intune (EMS)"))
            {
                UnAssign-EMSLic
            }			

            If (($Global:chkATPP1Lic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Advanced Threat Protection Plan1"))
            {
                UnAssign-ATPP1Lic
            }

            If (($Global:chkATPDefMACLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Advanced Threat Protection MAC"))
            {
                UnAssign-ATPPDefMACLic
            }
            
            If (($Global:chkAudioConfLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Audio Conferencing"))
            {
                UnAssign-AudioConferencingLic
            }
			
	        If (($Global:chkCAPLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Common Area Phone"))
            {
                UnAssign-CommonAreaPhoneLic
            }

            If (($Global:chkATPDefLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Defender Advanced Threat Protection"))
            {
                UnAssign-ATPDefLic
            }

            If (($Global:chkCRMLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Dynamics 365 Customer Engagement Plan"))
            {
                UnAssign-CRMLic
            }

            If (($Global:chkMeetingLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Teams Rooms Standard"))
            {            
                UnAssign-MeetingLic
            }

            If (($Global:chkMRSSLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Micrsoft Relationship Sales Solution"))
            { 
                UnAssign-MRSSLic
            }

            If (($Global:chkPhoneSysLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Phone System"))
            {
                UnAssign-PhoneSystemLic
            }

            If ($Global:chkProjP3Lic.Checked -eq "Checked")
            {
                UnAssign-ProjP3Lic
            }

            If ($Global:chkProjP5Lic.Checked -eq "Checked")
            {
                UnAssign-ProjP5Lic
            }            			

            If (($Global:chkPAppsP2Lic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "PowerApps Plan2"))
            {
                UnAssign-PAppsP2Lic
            }

            If (($Global:chkPBIPLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "PowerBI Pro"))
            {
                UnAssign-PowerBIProLic
            }

            If ($Global:chkProjP5Lic.Checked -eq "Checked")
            {
                UnAssign-VisioP2Lic
            }

            If ($Global:chkVisioP1Lic.Checked -eq "Checked")
            {
                UnAssign-VisioP1Lic
            }

            If ($Global:chkVisioP2Lic.Checked -eq "Checked")
            {
                UnAssign-VisioP2Lic
            }

            If (($Global:chkRmtAssistLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Microsoft Remote Assist"))
            {
                UnAssign-RemoteAsstLic
            }
        }
    
        If ($Global:chkStd.Checked -eq "Checked")
        {
            If ($HasE3 -eq "True")
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
                RetentPolicy
            }
            else
            {
                Assign-E3Lic
            }
    	    write-host "Checking the Retention Policy Configuration and that the PowerBI and EMS licenses are assigned"
       	    StandardLicenses
            $Global:chkStd.Checked = $false
        }

        If ($Global:chkReview.Checked -eq "Checked")
        {
            If (($Global:chkE3Lic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Enterprise E3"))
            {
                EnabledE3Feature
            }

            If (($Global:chkMRSSLic.Checked -eq "Checked") -or ($Global:locListBox.SelectedItem -eq "Dynamics 365 Customer Engagement Plan"))
            {
                $sskid = "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
                EnabledMRSSFeature
            }
            pause
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
            $Global:UserLicense = Get-MsolUser -UserPrincipalName $Global:UPN
            Update-LicenseDetails
        }
#        ## Finalize Form and Show Dialog
        $Global:from 
        $Global:form.Add_Shown( { $form.Activate(); $okButton.Focus() } )  #Activate and Set Focus
        $Global:result = $Global:form.ShowDialog()          ## Show the form, and wait for the response

        If ($Global:Result -ne "OK")
        {
            Build-LicenseStatsForm
            Build-AddUserDetailsForm
            Add-FormStandardButtons
            $Global:cancelButton.Text = "Cancel"
            $Global:form.Add_Shown( { $form.Activate(); $Global:txtUserInf.Focus()} )
            $Global:result = $Global:form.ShowDialog()          ## Show the form, and wait for the response
            O365HasLicenses
       }


 <#       If ($Global:Result -eq "OK")
        {
            $Global:locListBox.Visible = $false
            $Global:lblLicDet.Visible = $false
            $Global:txtUserInf.Text = ""
            $Global:lblUserInf.Visible = $true
            $Global:txtUserInf.Visible = $true
            $Global:ButGetENo.visible = $true
            $Global:lblUserDetails.Text = ""
            $Global:form.Add_Shown( { $form.Activate(); $okButton.Focus() } )  #Activate and Set Focus
            $Global:result = $Global:form.ShowDialog()          ## Show the form, and wait for the response
            write-host "Showing checked options"
            write-host $Global:chkAssign.Checked
            write-host $Global:chkRemove.Checked
            write-host $Global:chkStd.Checked
            write-host $Global:chkReview.Checked
            write-host $Global:chkReset.Checked
#            O365HasLicenses
        }
#>
#    }
}while ($Global:Result -eq "OK")