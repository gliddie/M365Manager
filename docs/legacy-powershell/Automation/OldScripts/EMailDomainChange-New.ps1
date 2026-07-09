#################################################################################
# 
# PowerShell source code
# Revision v1.0f
# ==========================================================================
#    'Project      : SDAP Menu Development
#    'Description  : Move Accounts to/from UL.COM and UL.ORG email addresses
#    'Called By    : SDAPAdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 08/28/2019
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 08/28/2019 Created from ConvertULORG.ps1 Script
#    '             : 07/09/2021 Modified how the SIP Address is updated and check to see if new primary address is already listed as a proxy address
#                  : 08/19/2021 Made additional changes to removing the SIP from the list of proxy addresses and asking the question to update the proxy address
#                  : 08/25/2021 Additional changes for the proper setting of the SIP address.  It only had a value if SIP was listed in the list of proxy addresses
#                  : 09/09/2021 Fixed error with $usr.msRTCSIP-PrimaryUserAddress to $usr.'msRTCSIP-PrimaryUserAddress'
#                  : 09/20/2021 Adjusted code so that any proxyaddresses that match the new Primary Address is removed.  This is to resolve issues where an individual has been moved back and forth from ul.com and ul.org multiple times.
#                  : 03/02/2022 Updated to use Windows Forms

Function Build-DomainChangeForm
{
    $Global:form = New-Object Windows.Forms.Form
    $Global:form.FormBorderStyle = "FixedToolWindow"
    $Global:form.Text = "Change eMail Domain"
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Width = 740 ; $form.Height = 480  # Make the form wider
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons
    Add-FormStandardButtons
    $Global:okButton.Visible = $False
    
    $Tab = 1
    $Script:Top = 10
    $ActLeft = 270
    ## Employee ID
    $Script:lblEmpID = New-Object System.Windows.Forms.Label
    $Script:lblEmpID.Text = "Employee No:"
    $Script:lblEmpID.Top = $Script:Top ; $Script:lblEmpID.Left = 5; $Script:lblEmpID.Width=120 ;$Script:lblEmpID.AutoSize = $true
    $Global:form.Controls.Add($Script:lblEmpID)    # Add to Form
    $Script:txtEmpID = New-Object Windows.Forms.TextBox
    $Script:txtEmpID.TabIndex = $Tab
    $Script:txtEmpID.Top = $Script:Top; $Script:txtEmpID.Left = 130; $Script:txtEmpID.Width = 120;
    $Script:txtEmpID.Text = $Script:txtInpEmpNo.Text
    $Global:form.Controls.Add($Script:txtEmpID)    # Add to Form
    $Global:InputFocus = $Script:txtEmpID
    $Script:txtEmpID.Add_Click({
        Reset-Form
     })

    $Script:Top = $Script:Top + 30
    ## Ticket
    $Script:lblTicket = New-Object System.Windows.Forms.Label
    $Script:lblTicket.Text = "Ticket Number:"
    $Script:lblTicket.Top = $Script:Top ; $Script:lblTicket.Left = 5; $Script:lblTicket.Width=120 ;$Script:lblTicket.AutoSize = $true 
    $Global:form.Controls.Add($Script:lblTicket)    # Add to Form
    $Script:txtTicket = New-Object Windows.Forms.TextBox
    $Script:txtTicket.TabIndex = $Tab++
    $Script:txtTicket.Top = $Script:Top; $Script:txtTicket.Left = 130; $Script:txtTicket.Width = 120;
    $Script:txtTicket.Text = "TASK"
    $Global:form.Controls.Add($Script:txtTicket)    # Add to Form
    $Script:txtTicket.Add_Click({
        Reset-Form
    })

    $Script:ButGetENo = New-Object Windows.Forms.Button
    $Script:ButGetENo.Location = New-object System.Drawing.Size(260,25)
    $Script:ButGetENo.Size = new-Object System.Drawing.Size(150,20)
    $Script:ButGetENo.Text = "Get Employee Details"
    $Script:ButGetENo.TabIndex = $Tab++
    $Global:form.Controls.Add($Script:ButGetENo)
    $Script:ButGetENo.Add_Click({
        If ($Script:txtEmpID.Text.Length -ge 5)
        {
            $Script:ADExists = ""
            $Script:Matchfound = "Y"
            $ErrorActionPreference = "SilentlyContinue"
            $Script:ADExists = [bool](get-ADUser $Script:txtEmpID.Text)
            $ErrorActionPreference = "Continue"

            If ($Script:ADExists -eq $True)
            {
                $Script:InfoDetails = "Retreiving Account Details for EmpNo: " + $Script:txtEmpID.Text
                Add-ErrorItem
                $Script:ButGetENo.visible = $false
                $Script:ReportFile	= $LogDirectory + "Report\Report-" + $FileName + "-EmpNo" + $Script:txtEmpID.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
                $Script:txtRptFile.Text = $Script:ReportFile

                $Script:ADUsr = Get-ADUser $Script:txtEmpID.Text -Properties GivenName,SurName,Initials,DisplayName,EmailAddress,ProxyAddresses,mail,Company,ExtensionAttribute8,msRTCSIP-PrimaryUserAddress
                $Script:RoutingDomain = "@ul.com"
                If (($Script:ADUsr.ExtensionAttribute8 -like "NFP *") -or ($Script:ADUsr.ExtensionAttribute8 -eq "Standards") -or ($Script:ADUsr.ExtensionAttribute8 -eq "Research") -or ($Script:ADUsr.ExtensionAttribute8 -eq "Support"))
                {
                    $Script:RoutingDomain = "@ul.org"
                }
                $Script:txtDispName.Text = $Script:ADUsr.DisplayName
                $Script:txtOrgName.Text = $Script:ADUsr.Company
                $Script:txtOrgLvl2.Text = $Script:ADUsr.ExtensionAttribute8
                $Script:txtPAddr.Text = $Script:ADUsr.EmailAddress
                $Script:txtSIPAddr.Text = ($Script:ADUsr.'msRTCSIP-PrimaryUserAddress') -replace ("SIP:","")
                If ($Script:txtPAddr.Text -like "*@UL.COM")
                {
                    $Script:butBldAddr.Text = "Configure New UL.ORG Address"
                    $Script:NewPAddr = ($Script:txtPAddr.Text) -replace ("ul.com","ul.org")
                }
                If ($Script:txtPAddr.Text -like "*@UL.ORG")
                {
                    $Script:butBldAddr.Text = "Configure New UL.COM Address"
                    $Script:NewPAddr =  "*" + ($Script:txtPAddr.Text) -replace ("ul.org","ul.com") + "*"
                }
                $Script:NewPAddr = "*" + $NewPAddr + "*"

                Show-BldAddrButton

                If ($Script:ADUsr.ProxyAddresses.Length -ne 0)
                {
                    foreach ($element in $Script:ADUsr.ProxyAddresses)
                    {
                        If ($element -cmatch "smtp:")
                        {
                            [void] $Script:txtProxyAddr.Items.Add($element.Substring(5,$element.Length-5))  # Add element to listbox
                        }
                    }
                }

                If ($Script:txtTicket.Text.Length -le 4)
                {
                    $Script:ErrorDetails = "You must enter a Ticket Number to proceed"
                    Add-ErrorItem                    
                }
            }
            else
            {
                $Script:ErrorDetails = "No account found for " + $Script:txtEmpID.Text
                Add-ErrorItem
            }
        }
    })

    $Script:Top = $Script:Top + 30
    $Script:lbDispName = New-Object System.Windows.Forms.Label
    $Script:lbDispName.Text = "Display Name:"
    $Script:lbDispName.Top = $Script:Top ; $Script:lbDispName.Left = 5; $Script:lbDispName.Width=120 ;$Script:lbDispName.AutoSize = $true 
    $Global:form.Controls.Add($Script:lbDispName)    # Add to Form
    $Script:txtDispName = New-Object Windows.Forms.TextBox
    $Script:txtDispName.ReadOnly = $true;
    $Script:txtDispName.TabIndex = 1000
    $Script:txtDispName.Top = $Script:Top; $Script:txtDispName.Left = 130; $Script:txtDispName.Width = 200;
    $Global:form.Controls.Add($Script:txtDispName)    # Add to Form

    $Script:Top = $Script:Top + 30
    $Script:lbOrgName = New-Object System.Windows.Forms.Label
    $Script:lbOrgName.Text = "Organization:"
    $Script:lbOrgName.Top = $Script:Top ; $Script:lbOrgName.Left = 5; $Script:lbOrgName.Width=120 ;$Script:lbOrgName.AutoSize = $true 
    $Global:form.Controls.Add($Script:lbOrgName)    # Add to Form
    $Script:txtOrgName = New-Object Windows.Forms.TextBox
    $Script:txtOrgName.ReadOnly = $true;
    $Script:txtOrgName.TabIndex = 1000
    $Script:txtOrgName.Top = $Script:Top; $Script:txtOrgName.Left = 130; $Script:txtOrgName.Width = 200;
    $Global:form.Controls.Add($Script:txtOrgName)    # Add to Form

    $Script:Top = $Script:Top + 30
    $Script:lbOrgLvl2 = New-Object System.Windows.Forms.Label
    $Script:lbOrgLvl2.Text = "Alpha Org Level2:"
    $Script:lbOrgLvl2.Top = $Script:Top ; $Script:lbOrgLvl2.Left = 5; $Script:lbOrgLvl2.Width=120 ;$Script:lbOrgLvl2.AutoSize = $true 
    $Global:form.Controls.Add($Script:lbOrgLvl2)    # Add to Form
    $Script:txtOrgLvl2 = New-Object Windows.Forms.TextBox
    $Script:txtOrgLvl2.ReadOnly = $true;
    $Script:txtOrgLvl2.TabIndex = 1000
    $Script:txtOrgLvl2.Top = $Script:Top; $Script:txtOrgLvl2.Left = 130; $Script:txtOrgLvl2.Width = 200;
    $Global:form.Controls.Add($Script:txtOrgLvl2)    # Add to Form

    $Script:butBldAddr = New-Object Windows.Forms.Button
    $Script:butBldAddr.Location = New-object System.Drawing.Size(380,$Script:Top)
    $Script:butBldAddr.Size = new-Object System.Drawing.Size(200,20)
    $Script:butBldAddr.visible = $false
    $Script:butBldAddr.TabIndex = $Tab++
    $Global:form.Controls.Add($Script:butBldAddr)
    $Script:butBldAddr.Add_Click({
        Build-AddressDetails})

    $Script:Top = $Script:Top + 30
    $Script:lblPAddr = New-Object System.Windows.Forms.Label
    $Script:lblPAddr.Text = "Primary Address:"
    $Script:lblPAddr.Top = $Script:Top ; $Script:lblPAddr.Left = 5; $Script:lblPAddr.Width=120 ;$Script:lblPAddr.AutoSize = $true 
    $Global:form.Controls.Add($Script:lblPAddr)    # Add to Form
    $Script:txtPAddr = New-Object Windows.Forms.TextBox
    $Script:txtPAddr.ReadOnly = $true;
    $Script:txtPAddr.TabIndex = 1000
    $Script:txtPAddr.Top = $Script:Top; $Script:txtPAddr.Left = 130; $Script:txtPAddr.Width = 200;
    $Global:form.Controls.Add($Script:txtPAddr)    # Add to Form

    $Script:lblBldAddr = New-Object System.Windows.Forms.Label 
    $Script:lblBldAddr.Text = "Proposed Address"
    $Script:lblBldAddr.Top = $Script:Top ; $Script:lblBldAddr.Left = 590; $Script:lblBldAddr.Width=120 ;$Script:lblBldAddr.AutoSize = $true
    $Script:lblBldAddr.Visible = $false
    $Global:form.Controls.Add($Script:lblBldAddr)    # Add to Form
    $Script:txtBldAddr = New-Object Windows.Forms.TextBox
    $Script:txtBldAddr.ReadOnly = $true; $Script:txtBldAddr.TabIndex = 1000
    $Script:txtBldAddr.visible = $false
    $Script:txtBldAddr.Text = ""
    $Script:txtBldAddr.Top = $Script:Top; $Script:txtBldAddr.Left = 380; $Script:txtBldAddr.Width = 200;
    $Global:form.Controls.Add($Script:txtBldAddr)

    $Script:Top = $Script:Top + 30
    $Script:lblSIPAddr = New-Object System.Windows.Forms.Label
    $Script:lblSIPAddr.Text = "SIP Address:"
    $Script:lblSIPAddr.Top = $Script:Top ; $Script:lblSIPAddr.Left = 5; $Script:lblSIPAddr.Width=120 ;$Script:lblSIPAddr.AutoSize = $true 
    $Global:form.Controls.Add($Script:lblSIPAddr)    # Add to Form
    $Script:txtSIPAddr = New-Object Windows.Forms.TextBox
    $Script:txtSIPAddr.ReadOnly = $true; $Script:txtSIPAddr.TabIndex = 1000
    $Script:txtSIPAddr.Top = $Script:Top; $Script:txtSIPAddr.Left = 130; $Script:txtSIPAddr.Width = 200;
    $Global:form.Controls.Add($Script:txtSIPAddr)    # Add to Form
<#
    $Script:lblAvalAddr = New-Object System.Windows.Forms.Label
    $Script:lblAvalAddr.Text = "New Primary Address"
    $Script:lblAvalAddr.Top = $Script:Top ; $Script:lblAvalAddr.Left = 590; $Script:lblAvalAddr.Width=120 ;$Script:lblAvalAddr.AutoSize = $true 
    $Script:lblAvalAddr.visible = $false
    $Global:form.Controls.Add($Script:lblAvalAddr)    # Add to Form
    $Script:txtAvalAddr = New-Object Windows.Forms.TextBox
    $Script:txtAvalAddr.ReadOnly = $true; $Script:txtAvalAddr.TabIndex = 1000
    $Script:txtAvalAddr.Top = $Script:Top; $Script:txtAvalAddr.Left = 380; $Script:txtAvalAddr.Width = 200;
    $Script:txtAvalAddr.visible = $false
    $Global:form.Controls.Add($Script:txtAvalAddr)    # Add to Form
#>
    $Script:Top = $Script:Top + 30
    $Script:lblProxyAddr = New-Object System.Windows.Forms.Label 
    $Script:lblProxyAddr.Text = "Other ProxyAddr:"
    $Script:lblProxyAddr.Top = $Script:Top ; $Script:lblProxyAddr.Left = 5; $Script:lblProxyAddr.Width=120 ;$Script:lblProxyAddr.AutoSize = $true
    $Global:form.Controls.Add($Script:lblProxyAddr)    # Add to Form
    $Script:txtProxyAddr = New-Object Windows.Forms.ListBox
    $Script:txtProxyAddr.Top = $Script:Top; $Script:txtProxyAddr.Left = 130; $Script:txtProxyAddr.Height = 80; $Script:txtProxyAddr.Width = 200;
    $Script:txtProxyAddr.BackColor = "LightGray"
    $Script:txtProxyAddr.TabIndex = 1000
    $Global:form.Controls.Add($Script:txtProxyAddr)    # Add to Form

    #Show Errors:
    $Script:txtError = New-Object Windows.Forms.ListBox
    $Script:txtError.Top = $Script:Top; $Script:txtError.Left = 350; $Script:txtError.Height = 90; $Script:txtError.Width = 350;
    $Script:txtError.BackColor = "LightGray"
    $Script:txtError.HorizontalScrollBar = $true
    $Global:form.Controls.Add($Script:txtError)    # Add to Form
    $x = $Script:txtError.Items.Add("Starting Rename AD Account......")
    $Script:txtError.Add_Click({$Script:txtError.BackColor = "LightGray"})

    $Script:Top = $Script:Top + 90
    ##ReportDetails
    $Script:lblRptFile = New-Object System.Windows.Forms.Label
    $Script:lblRptFile.Text = "Report Details:"
    $Script:lblRptFile.Top = $Script:Top ; $Script:lblRptFile.Left = 5; $Script:lblRptFile.Width=100 ;$Script:lblRptFile.AutoSize = $true
    $Global:form.Controls.Add($Script:lblRptFile)    # Add to Form
    $Script:txtRptFile = New-Object Windows.Forms.TextBox
    $Script:txtRptFile.ReadOnly = $true; $Script:txtRptFile.TabIndex = 1000
    $Script:txtRptFile.Top = $Script:Top; $Script:txtRptFile.Left = 130; $Script:txtRptFile.Width = 570;
    $Script:txtRptFile.Text = ""
    $Global:form.Controls.Add($Script:txtRptFile)    # Add to Form
    $Script:Top = $Script:Top + 20
}

Function Build-AddressDetails
{
#1.  Change primary email address from @ul.com to @ul.org or @ul.org to @ul.com and check to see if account already as that address as a proxy address.
#1a.  if not change the current address to use the other domain and check to see if that address is in use

    If ($usr.ProxyAddresses -like $Script:NewPAddr)
    {
        $Script:ErrorDetails = "New Address $($script:NewPAddr) Found "
        $Script:PromoteExistingAddr = "Y"
    }
    else
    {
        $Exists = get-mailbox [bool]($mbx = get-mailbox $Script:NewPAddr)
        If ($Exists -eq $True)
        {
            $Script:InfoDetails= "New Address $($script:NewPAddr) In use by $($mbx.DisplayName)"            
        }
        else
        {
            $Script:InfoDetails= "New Address $($script:NewPAddr) Not Found on Current User Account"
        }
    }
    $Script:txtBldAddr.Text = $script:NewPAddr
    Add-ErrorItem
    $Script:lblBldAddr.Visible = $True
    $Script:txtBldAddr.Visible = $True
}

Function Build-AddressDetails-RenameCode
{
    If ($Script:txtTicket.Text.Length -gt 4)
    {
        $SpecDispChara = "[<>?_/+=*&^%$#@!{}.()?0123456789:~ ]"

        # Remove Spaces and other special characters
        $FName = $Script:txtNewFName.Text -replace $SpecDispChara
        $Script:txtAdjFName.Text = $FName
        If ($FName.Substring(1,$FName.length-1) -cnotmatch '[a-z]')
        {
            $FName = $FName.substring(0,1).toupper() + $FName.Substring(1,$FName.length-1).ToLower()
        }
        else
        {
            If ($FNAME.Substring(0,1) -cnotmatch '[A-Z]')
            {
                $FName = $FName.substring(0,1).toupper() + $FName.Substring(1,$FName.length-1)
            }
        }
        If (($FName -like "*-*") -or ($FName -like "*'*"))
        {
            $LocDash = $FName.IndexOf("-")
            $LocHyphen = $FName.IndexOf("'")
            $FName = $FName.substring(0,1).toupper() + $FName.Substring(1,$FName.length-1).ToLower()
            If ($LocDash -gt 0)
            {
                $FName = $FName.substring(0,$LocDash+1) + $FName.substring($LocDash+1,1).ToUpper() + $FName.substring($LocDash+2,$FName.Length-($LocDash+2))
            }

            If ($LocHyphen -gt 0)
            {
                $FName = $FName.substring(0,$LocHyphen+1) + $FName.substring($LocHyphen+1,1).ToUpper() + $FName.substring($LocHyphen+2,$FName.Length-($LocHyphen+2))
            }
        }
        $Script:txtAdjFName.Text = $FName

        If ($Script:txtAdjFName.Text -cnotlike $Script:txtNewFName.Text)
        {
            $Script:InfoDetails = "Fixing formatting of the entered First Name details"
            $Script:lblAdj.Visible = $true
            $Script:txtAdjFName.Visible = $True
            Add-ErrorItem
        }

        $Script:txtAdjInitial.Text = $Script:txtNewInitial.Text -replace $SpecDispChara
        $Script:txtAdjInitial.Text = $Script:txtAdjInitial.Text.ToUpper()
        If ($Script:txtAdjInitial.Text -cnotlike $Script:txtNewInitial.Text)
        {
            $Script:InfoDetails = "Fixing formatting of the entered Middle Initial details"
            $Script:lblAdj.Visible = $true
            $Script:txtAdjInitial.Visible = $true
            Add-ErrorItem
        }

		If ($Script:chkChgLName.Checked -eq $False)
		{
			$Script:txtNewLName.Text = $Script:txtLName.Text
		}
		$LName = $Script:txtNewLName.Text -replace $SpecDispChara
        $Script:txtAdjLName.Text = $LName
        If ($LName.Substring(1,$LName.length-1) -cnotmatch '[a-z]')
        {
            $LName = $LName.substring(0,1).toupper() + $LName.Substring(1,$LName.length-1).ToLower()
        }
        else
        {
            If ($LName.Substring(0,1) -cnotmatch '[A-Z]')
            {
                $LName = $LName.substring(0,1).toupper() + $LName.Substring(1,$LName.length-1)
            }            
        }
        If (($LName -like "*-*") -or ($LName -like "*'*"))
        {
            $LocDash = $LName.IndexOf("-")
            $LocHyphen = $LName.IndexOf("'")
            $LName = $LName.substring(0,1).toupper() + $LName.Substring(1,$LName.length-1).ToLower()
            If ($LocDash -gt 0)
            {
                $LName = $LName.substring(0,$LocDash+1) + $LName.substring($LocDash+1,1).ToUpper() + $LName.substring($LocDash+2,$LName.Length-($LocDash+2))
            }

            If ($LocHyphen -gt 0)
            {
                $LName = $LName.substring(0,$LocHyphen+1) + $LName.substring($LocHyphen+1,1).ToUpper() + $LName.substring($LocHyphen+2,$LName.Length-($LocHyphen+2))
            }
        }
        $Script:txtAdjLName.Text = $LName

        If ($Script:txtAdjLName.Text -cnotlike $Script:txtNewLName.Text)
        {
            $Script:InfoDetails = "Fixing formatting of the entered Last Name details"
            $Script:lblAdj.Visible = $true
            $Script:txtAdjLName.Visible = $true
            Add-ErrorItem
        }

        $Script:butBldAddr.visible = $False
        $Script:lblNewDispName.Visible = $true
        $Script:txtNewDispName.Visible = $true
        If ($Script:chkChgDispOnly.Checked -ne "Checked")
        {
            $Script:lblAvalAddr.visible = $true
            $Script:txtAvalAddr.visible = $true
        $Script:lblBldAddr.visible = $true
        $Script:txtBldAddr.visible = $true
        }

        #Build Email Address
        $SpecAddrChara = "[-']"
        If ($Script:txtNewInitial.Text.Length -eq 0)
        {
            $Script:txtNewDispName.Text = ($Script:txtAdjLName.Text + ", " + $Script:txtAdjFName.Text) -replace $SpecAddrChara
            $Script:txtBldAddr.Text = $Script:txtNewDispName.Text
        }
        else
        {
            $Script:txtNewDispName.Text = ($Script:txtAdjLName.Text + ", " + $Script:txtAdjFName.Text + " " + $Script:txtAdjInitial.Text + ".") -replace $SpecAddrChara
            $Script:txtBldAddr.Text = $Script:txtNewDispName.Text
        }

        If ($Script:txtAdjInitial.Text.Length -eq 0)
        {
            $Script:txtBldAddr.Text = ($Script:txtAdjFName.Text + "." + $Script:txtAdjLName.Text + $Script:RoutingDomain) -replace $SpecAddrChara
            $Script:NewCN = $Script:txtAdjFName.Text + " " + $Script:txtAdjLName.Text
        }
        else
        {
            $Script:txtBldAddr.Text = ($Script:txtAdjFName.Text + "." + $Script:txtAdjInitial.Text + "." + $Script:txtAdjLName.Text + $Script:RoutingDomain) -replace $SpecAddrChara
            $Script:NewCN = $Script:txtAdjFName.Text + " " + $Script:txtAdjInitial.Text + ". " + $Script:txtAdjLName.Text
        }

		If ($Script:chkChgDispOnly.Checked -eq $False)
		{
			$MbxExists = [bool]($Mbx = get-mailbox $Script:txtBldAddr.Text -ErrorAction SilentlyContinue)
			If ($MbxExists -eq $True)
			{
				Addr-InUse
			}
			else
			{
				$Script:InfoDetails = "The email " + $Script:txtBldAddr.Text + " is available"
				Add-ErrorItem
				$Script:txtAvalAddr.Text = $Script:txtBldAddr.Text
				$Script:Matchfound = "N"
			}
		}
		$Global:okButton.ForeColor = "Green"
		$Global:okButton.Text = "Continue"
        $Global:OKButton.visible = $True
    }
    else
    {
        $Global:InputFocus = $Script:txtTicket
        $Script:ErrorDetails = "You must enter a Ticket Number to proceed"
        Add-ErrorItem
    }
    $Global:cancelButton.Text = "Exit"
}

Function Add-ErrorItem
{
    If ($Script:ErrorDetails -ne "")
    {
#        $TextColour = [System.Drawing.Color]::Red
#        $TextColourBrush = New-Object System.Drawing.SolidBrush($TextColour)
        $Script:txtError.ForeColor = "Red"
        $Script:txtError.Items.Add($Script:ErrorDetails)  # Add element to listbox
        $Script:ErrorDetails = ""
    }
    else
    {
#        $TextColour = [System.Drawing.Color]::Cyan
#        $TextColourBrush = New-Object System.Drawing.SolidBrush($TextColour)
        $Script:txtError.Items.Add($Script:InfoDetails)  # Add element to listbox
        $Script:txtError.ForeColor = "Yellow"
        $Script:InfoDetails = ""
    }

#    $Script:txtError.Graphics.TextRenderingHint = 'SingleBitPerPixelGridFit'
#    $Script:txtError.Graphics.DrawString($Script:InfoDetails, $Script:txtError.Font, $TextColourBrush, (new-object System.Drawing.PointF($Script:txtError.Bounds.X, $Script:txtError.Bounds.Y)))
    $Script:txtError.SelectedIndex = $Script:txtError.Items.Count - 1;
}

Function Reset-Form
{
    $Script:ButGetENo.Visible = $True
    $Script:butBldAddr.visible = $false
    $Script:txtBldAddr.Visible = $false
    $Script:lblBldAddr.Visible = $false
    $Script:txtDispName.Text = ""
    $Script:txtOrgName.Text = ""
    $Script:txtOrgLvl2.Text = ""
    $Script:txtPAddr.Text = ""
    $Script:txtSIPAddr.Text = ""
    $Script:txtProxyAddr.Items.Clear()
    $Script:NewPAddr = ""
    $Script:PromoteExistingAddr = "N"
    $Script:butBldAddr.visible = $False
#    $Script:lblAvalAddr.visible = $false
#    $Script:txtAvalAddr.visible = $false
}

Function Show-BldAddrButton
{
    $Script:Matchfound = "Y"
    $Script:MCnt = 0
    $Script:butBldAddr.visible = $True
    $Global:OKButton.visible = $false
}

Function ShowProxy
{
    $cnt = 0
    foreach ($PrxyAddr in $usr.ProxyAddresses)
    {
        if ($cnt -eq 0)
        {
            write-host "        All Proxy Addresses: " $PrxyAddr
            $LineToWrite = "INFO" + "`t" + "   All Proxy Addresses:  " + $PrxyAddr
            $cnt++
        }
        else
        {
            write-host "                             " $PrxyAddr
            $LineToWrite = "INFO" + "`t" + "                        " + $PrxyAddr
        }
        WriteReportEvent
        If ($PrxyAddr -clike "SMTP:*")
        {
            $Script:OldPrimAddr = $PrxyAddr
        }
    }
}

function Add-FormStandardButtons
{
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    $Global:buttonPanel = New-Object Windows.Forms.Panel  
    $buttonPanel.Size = New-Object Drawing.Size @(400,40) 
    $buttonPanel.Dock = "Bottom"    
    $Global:cancelButton = New-Object Windows.Forms.Button  
        $Global:cancelButton.Top = $buttonPanel.Height - $Global:cancelButton.Height - 10; $Global:cancelButton.Left = $buttonPanel.Width - $Global:cancelButton.Width - 10 
        $Global:cancelButton.TabIndex = 99
        $Global:cancelButton.Text = "Cancel" 
        $Global:cancelButton.DialogResult = "Cancel" 
        $Global:cancelButton.Anchor = "Right"
    ## Create the OK button, which will anchor to the left of Cancel 
    $Global:okButton = New-Object Windows.Forms.Button
        $Global:okButton.Top = $cancelButton.Top ;$Global:okButton.Left = $cancelButton.Left - $Global:okButton.Width - 10
        If ($Global:OKDetails -eq "Remove Access")
        {
            $Global:OKButton.Width = 100
            $Global:okButton.Left = $cancelButton.Left - $Global:okButton.Width - 10
        }
        $Global:okButton.TabIndex = 98
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
    ## Finalize Form and Show Dialog
    $Global:form.Add_Shown( { $form.Activate(); $Global:InputFocus.Focus() } )  #Activate and Set Focus 
    $Global:Result = $Global:form.ShowDialog()          ## Show the form, and wait for the response
}

# Declare Drive | Folders | and Files
	$FileName		= "EmailDomainChange"
	$LogDrive		= "E:"
	$LogPath		= "\SDAP"
	$LogFolder		= "\" + $FileNAme
	$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
	$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	$ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
# Retrieve the user name
	$WhoAmI			= WhoAmI
		
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name

$Script:ErrorDetails = ""
Build-DomainChangeForm
Publish-Form
[void] $Script:txtError.Items.Add("SRunning Convert Email Domain.....")

pause

write-host "Running Convert Email Domain....."

write-host "Enter Employee Number for email domain change: " -ForegroundColor cyan -NoNewline
$uid = read-host

Do
{
    $ReportFile	= $LogDirectory + "Report\Report-" + $FileName + "-EmpNo" + $uid + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
    $LineToWrite = "INFO" + "`t" + "Converting Email Addresses for: " + $UID + "`n"
    WriteReportEvent
# Setup Folders and Files	
	invoke-expression -Command E:\O365AdminShared\Scripts\CheckLogFiles.ps1
	$LineToWrite = "STAR" + "`t" + $FileName + " script has started"
	WriteLogEvent
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	WriteLogEvent

    $EmpID = $uid + "@global.ul.com"
    $usr = Get-ADUser $UID -Properties ProxyAddresses,mail,TargetAddress,msRTCSIP-PrimaryUserAddress

    If ($usr.mail -like "*ul.com*")
    {
        write-host "This user will be moved from the UL.COM email domain to the UL.ORG mail domain." -ForegroundColor Green
        $LineToWrite = "INFO" + "`t" + "This user will be moved from the UL.COM email domain to the UL.ORG mail domain "
        WriteReportEvent
    }
    else
    {
        write-host "This user will be moved from the UL.ORG email domain to the UL.COM mail domain." -ForegroundColor Magenta
        $LineToWrite = "INFO" + "`t" + "This user will be moved from the UL.ORG email domain to the UL.COM mail domain "
        WriteReportEvent
    }

    write-host "              Employee Name: " $usr.Name
    $LineToWrite = "INFO" + "`t" + "         Employee Name: " + $usr.Name
    WriteReportEvent
    write-host "       Current Mail Address: " $Usr.mail
    $LineToWrite = "INFO" + "`t" + "  Current Mail Address: " + $Usr.mail
    WriteReportEvent

	switch -Wildcard ($usr.mail)
    {
        "*ul.com"
        {
            $NewAddr = $usr.mail.Replace("ul.com","ul.org")
            $TarAddr = $usr.TargetAddress.Replace("ul.com","ul.org")
            write-host "         New UL.ORG Address: " $NewAddr
            $LineToWrite = "INFO" + "`t" + "    New UL.ORG Address: " + $NewAddr
        }
        "*ul.org"
        {
            $NewAddr = $usr.mail.Replace("ul.org","ul.com")
            $TarAddr = $usr.TargetAddress.Replace("ul.org","ul.com")
            write-host "         New UL.COM Address: " $NewAddr
            $LineToWrite = "INFO" + "`t" + "    New UL.COM Address: " + $NewAddr
        }
    }
    WriteReportEvent
    $NewSIPAddr = "sip:" + $NewAddr

	if ($PrxyAddr -like "sip:*")
	{
		$SIPAddr = $TarAddr.Replace("SMTP:","sip:")
	}
    write-host "     Current Target Address: " $usr.TargetAddress
    $LineToWrite = "INFO" + "`t" + "Current Target Address: " + $usr.TargetAddress
    WriteReportEvent
    write-host "         New Target Address: " $TarAddr
    $LineToWrite = "INFO" + "`t" + "    New Target Address: " + $TarAddr
    WriteReportEvent
    write-host "         New SIP Address: " $NewsipAddr
    $LineToWrite = "INFO" + "`t" + "       New SIP Address: " + $NewSIPAddr
    WriteReportEvent

    ShowProxy

    write-host "Is the the correct user (Y/N)? " -ForegroundColor yellow -NoNewline
    $cont = read-host

    If ($Cont -eq "Y")
    {
        write-host "`nAddress Details After Changes" -ForegroundColor Green
        $LineToWrite = "`nINFO" + "`t" + "Address Details After Changes "
        WriteReportEvent

        Set-ADUser -Identity $usr -EMailAddress $NewAddr
        $AddAddr = "SMTP:"+$NewAddr
        ForEach ($PrxyAddr in $usr.ProxyAddresses)
        {
            If ($PrxyAddr -like $AddAddr)
            {
                #Remove proxy address that is like the new primary address
                Set-AdUser -Identity $usr -Remove @{ProxyAddresses=$PrxyAddr}
			    write-host "`n***  UL.COM or UL.ORG Address removed from existing list of proxy addresses. ****" -ForegroundColor Red
                $LineToWrite = "REMO" + "`t" + "   UL.COM or UL.ORG Address removed from existing list of proxy address"
                WriteReportEvent
            }
            If ($PrxyAddr -like "sip:*")
            {
                Set-ADUser $Usr -Remove @{ProxyAddresses=$PrxyAddr}
			    write-host "`n***  Removing SIP address from list of ProxyAddresses because this is no longer used to configure this address. ****" -ForegroundColor Red
	    	    $LineToWrite = "REMO" + "`t" + "   SIP Address removed from existing list of proxy address"
                WriteReportEvent
            }
        }
        Set-ADUser -Identity $usr -Add @{ProxyAddresses=$AddAddr}
	    Set-AdUser -Identity $usr -Remove @{ProxyAddresses=$usr.TargetAddress}
	    Set-ADUser -Identity $usr -Replace @{targetaddress=$TarAddr}
        $OldAddr = $usr.TargetAddress.Replace("SMTP:","smtp:")
        Set-ADUser -Identity $usr -Add @{ProxyAddresses=$OldAddr} -ErrorAction SilentlyContinue
        Set-ADUser -Identity $usr -Replace @{'msRTCSIP-PrimaryUserAddress'=$NewSIPAddr}

        ForEach ($PrxyAddr in $usr.ProxyAddresses)
        {
            If ($PrxyAddr -like "*sip:*")
            {
                Set-AdUser -Identity $usr -Remove @{ProxyAddresses=$ProxyAddr}
       			write-host "`n***  Old SIP Address $($ProxyAddr) removed from list of proxy addresses. ****" -ForegroundColor Red
    			$LineToWrite = "DELE" + "`t" + "   Old SIP Address $($ProxyAddr) removed from list of proxy addresses" 
                WriteReportEvent
            }
        }
		sleep -Seconds 5
        $usr = Get-ADUser $UID -Properties ProxyAddresses,mail,TargetAddress,msRTCSIP-PrimaryUserAddress
        write-host "`n        New Primary Address: " $usr.mail
        $LineToWrite = "INFO" + "`t" + "   New Primary Address: " + $usr.mail
        WriteReportEvent

        write-host "`n        New Target Address: " $usr.mail
        $LineToWrite = "INFO" + "`t" + "   New Target Address: " + $usr.TargetAddress
        WriteReportEvent

        write-host "`n           New SIP Address: " $usr.mail
        $LineToWrite = "INFO" + "`t" + "      New SIP Address: " + $usr.'msRTCSIP-PrimaryUserAddress'
        WriteReportEvent
        ShowProxy		
    }
    else
    {
            write-host "No changes being made for this individual"
            $LineToWrite = "INFO" + "`t" + "Agent cancelled changes for this individual"
            WriteReportEvent
    }
    write-host "Enter Employee Number of individual for email domain change (0) to exit: " -ForegroundColor cyan -NoNewline
    $uid = read-host
}while ($uid -ne "0")