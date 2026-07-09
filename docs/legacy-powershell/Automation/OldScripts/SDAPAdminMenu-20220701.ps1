####
#### SDAP Admin Menu
####
#################################################################################
# 
# PowerShell source code
# Revision v1.05
# ==========================================================================
#    'Project      : SDAP Menu Creation
#    'Description  : SDAP Administration Main Menu
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 11/17/2016
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#                  : SAG - 11/03/2020 Changed to a GUI format
#                  : SAG - 12/02/2020 Added modules and updated code for Purge process to use GUI format
#                  : SAG - 09/21/2021 Disabled code that referenced UMMailbox added functions for to AssignE3License, RetentionPolicy, StandardLicenses and O365Licenses
#                  : SAG - 09/22/2021 Added code to allow for configuring of a forwarding address for M&A accounts
#                  : SAG - 10/04/2021 Added checks to not remove ownership/membership of DST or DSG groups
#                  : SAG - 10/13/2021 Removed the "under development" details for the Manage Phone Number option
#                  : SAG - 10/14/2021 Added the setting of the mailbox quotas to the Retent function
#                  : SAG - 07/01/2022 Added to add the Account Purge date to the log file value writtten to the msDS-cloudExtensionAttribute1 value and to record the date if legal hold was enabled
#
#################################################################################

#Check if the LogFile directory exists for a given menu item (this may be able to be retired)
function CheckLogFiles
{
	# Create Files folder $LogDirectory if it's not present
	if (Test-Path $LogDirectory)
		{
		# the directory is present
		}
	else
		{ 
		mkdir $LogDirectory
		}
} #end CheckLogFiles

#Default form buttons
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

function Add-FormInfoOnlyButtons
{
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    $Global:buttonPanel = New-Object Windows.Forms.Panel  
    $buttonPanel.Size = New-Object Drawing.Size @(400,40) 
    $buttonPanel.Dock = "Bottom"
    $Global:okButton = New-Object Windows.Forms.Button  
        $Global:okButton.Top = $buttonPanel.Height - $Global:okButton.Height - 10; $Global:okButton.Left = $buttonPanel.Width - $Global:okButton.Width - 10 
        $Global:okButton.TabIndex = 98
        $Global:okButton.Text = "Done" 
        $Global:okButton.DialogResult = "OK" 
        $Global:okButton.Anchor = "Right"      
    ## Add the buttons to the button panel 
    $Global:buttonPanel.Controls.Add($Global:okButton) 
    ## Add the button panel to the form 
    $Global:form.Controls.Add($buttonPanel)
    ## Set Default actions for the buttons 
    $Global:form.AcceptButton = $Global:okButton          # ENTER = ok
    $Global:InputFocus = $Global:okButton
 }

Function AssignE3License
{
	Write-host ""
    $sskid = "ul:ENTERPRISEPACK"
    $LicType = "Enterprise E3"

    If ($ADUser.ExtensionAttribute1 -like "Employee*")
    {
        $DisPlan = $Global:DisPlanEmpE3
    }
    else
    {
        $DisPlan = $Global:DisPlanNonEmpE3
    }
	
	$LicDet = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}
			
    If ($HasE3 -eq $True)
    {
        write-host "`nEnterprise E3 License Already Assigned to " $LicType "License"
        $LineToWrite = $WhoAmI + "`t" + "DELE" + "`t" + "Enterprise E3 License Already Assigned to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
    else
    {
        If ($ADUser.ExtensionAttribute1 -like "Employee*")
	    {
		    write-host "`nAssigning " $EmpNo "an " $LicType "License with Employee Features Enabled" -ForegroundColor Yellow
            $LineToWrite = $WhoAmI + "`t" + "UPDA" + "`t" + "Assigning " + $LicType + "with Employee Features to " + $EmpNo
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	    }
	    else
	    {
		    write-host "`nAssigning " $EmpNo "an " $LicType "License with Non-Employee Features Enabled" -ForegroundColor Yellow
            $LineToWrite = $WhoAmI + "`t" + "UPDA" + "`t" + "Assigning " + $LicType + "with Non-Employee Features to " + $EmpNo
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	    }

        $DLO = ($DisPlan.Split(“,”))
        $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense $SSKID –LicenseOptions $MyO365Sku
        write-host "Waiting for mailbox to be provisioned.." -ForegroundColor Red -NoNewline
        Do{
            Start-Sleep -Seconds 5
            write-host ".." -ForegroundColor Red -NoNewline
        }while([bool](get-mailbox $EmpNo -ErrorAction Silentlycontinue) -eq $False)
        write-host ""
        RetentPolicy
    }

    If ($HasExP2 -eq "True")
    {
        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:EXCHANGEENTERPRISE"
        write-host "`nRemoving Exchange Online P2 License and reassigning to a " $LicType "License"
        $LineToWrite = $WhoAmI + "`t" + + "DELE" + "`t" + "Removing P2 License from " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
}

 #Builds General Account Info Form
Function Build-SDAPGeneralAccountInfo
{
    $Global:form = New-Object Windows.Forms.Form
    $Global:form.FormBorderStyle = "FixedToolWindow"
    $Global:form.Text = $Global:GeneralTitle
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Width = 740 ; $form.Height = 670  # Make the form wider
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons
    
    $Tab = 10
    $Global:Top = 10
    ## Employee ID
    $Global:lblHost = New-Object System.Windows.Forms.Label
    $Global:lblHost.Text = "Employee No:"
    $Global:lblHost.Top = $Global:Top ; $Global:lblHost.Left = 5; $Global:lblHost.Width=120 ;$Global:lblHost.AutoSize = $true
    $Global:form.Controls.Add($Global:lblHost)    # Add to Form
    $Global:txtHost = New-Object Windows.Forms.TextBox
    $Global:txtHost.ReadOnly = $true;
    $Global:txtHost.TabIndex = $Tab++
    $Global:txtHost.Top = $Global:Top; $Global:txtHost.Left = 130; $Global:txtHost.Width = 120;
    $Global:txtHost.Text = $ADUsr.SamAccountName
    $Global:form.Controls.Add($Global:txtHost)    # Add to Form
 
    $Global:Top = $Global:Top + 30
    ## Employee Name
    $Global:lblName = New-Object System.Windows.Forms.Label
    $Global:lblName.Text = "Name:"
    $Global:lblName.Top = $Global:Top ; $Global:lblName.Left = 5; $Global:lblName.Width=120 ;$Global:lblName.AutoSize = $true 
    $Global:form.Controls.Add($Global:lblName)    # Add to Form
    $Global:txtName = New-Object Windows.Forms.TextBox
    $Global:txtName.ReadOnly = $true;
    $Global:txtName.TabIndex = $Tab++ # set Tab Order
    $Global:txtName.Top = $Global:Top; $Global:txtName.Left = 130; $Global:txtName.Width = 120;
    $Global:txtName.Text = $ADUsr.Name
    $Global:form.Controls.Add($Global:txtName)    # Add to Form

    $Global:Top = $Global:Top + 30
    ## Employee Type
    $Global:lblType = New-Object System.Windows.Forms.Label
    $Global:lblType.Text = "Type:"
    $Global:lblType.Top = $Global:Top ; $Global:lblType.Left = 5; $Global:lblType.Width=120 ;$Global:lblType.AutoSize = $true
    $Global:txtType = New-Object Windows.Forms.TextBox
    $Global:txtType.ReadOnly = $true
    If ($ADUsr.ExtensionAttribute1 -notlike "Ex*")
    {
        $Global:txtType.BackColor = "LightGray"
        $Global:txttype.ForeColor = "Red"
        $Global:lbltype.ForeColor = "Red"
        If (($AdUsr.DistinguishedName -notlike "*Disabled*") -and ($ADUsr.Enabled -eq $True))
        {
            $Global:DoNotPurge = "Yes"
        }
    }
    $Global:txtType.TabIndex = $Tab++ # set Tab Order
    $Global:txtType.Top = $Global:Top; $Global:txtType.Left = 130; $Global:txtType.Width = 120;
    $Global:txtType.Text = $ADUsr.ExtensionAttribute1
    $Global:form.Controls.Add($Global:lblType)    # Add to Form
    $Global:form.Controls.Add($Global:txtType)    # Add to Form
 
    $Global:Top = $Global:Top + 30
    ## ADAccountEnabled
    $Global:lblADEnabled = New-Object System.Windows.Forms.Label
    $Global:lblADEnabled.Text = "AD Account Enabled:"
    $Global:lblADEnabled.Top = $Global:Top ; $Global:lblADEnabled.Left = 5; $Global:lblADEnabled.Width=120 ;$Global:lblADEnabled.AutoSize = $true
    $Global:txtADEnabled = New-Object Windows.Forms.TextBox
    $Global:txtADEnabled.ReadOnly = $true
    If ($ADUsr.Enabled -eq "True")
    {
        $Global:txtADEnabled.BackColor = "LightGray"
        $Global:txtADEnabled.ForeColor = "Red"
        $Global:lblADEnabled.ForeColor = "Red"
        $Global:DoNotPurge = "Yes"
    }
    $Global:txtADEnabled.TabIndex = $Tab++ # set Tab Order
    $Global:txtADEnabled.Top = $Global:Top; $Global:txtADEnabled.Left = 130; $Global:txtADEnabled.Width = 120;
    $Global:txtADEnabled.Text = $ADUsr.Enabled
    $Global:form.Controls.Add($Global:lblADEnabled)    # Add to Form
    $Global:form.Controls.Add($Global:txtADEnabled)    # Add to Form
 
    $Global:txtAdmAcct = ""
    If ($ADUAdm -eq $True)
    {
        $form.Height = 700
        $Global:Top = $Global:Top + 30
        ## Admin Account
        $Global:lblAdmAcct = New-Object System.Windows.Forms.Label
        $Global:lblAdmAcct.Text = "Admin Account:"
        $Global:lblAdmAcct.Top = $Global:Top ; $Global:lblAdmAcct.Left = 5; $Global:lblAdmAcct.Width=120 ;$Global:lblAdmAcct.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAdmAcct)    # Add to Form
        $Global:txtAdmAcct = New-Object Windows.Forms.TextBox
        $Global:txtAdmAcct.ReadOnly = $true
        $Global:txtAdmAcct.TabIndex = $Tab++ # set Tab Order
        $Global:txtAdmAcct.Top = $Global:Top; $Global:txtAdmAcct.Left = 130; $Global:txtAdmAcct.Width = 120;
        $Global:txtAdmAcct.Text = $AdminAcct
        $Global:form.Controls.Add($Global:txtAdmAcct)    # Add to Form
    }
   
    $Global:Top = $Global:Top + 30
    #Active Directory OU
    $Global:lblADObj = New-Object System.Windows.Forms.Label
    $Global:lblADObj.Text = "AD Object Location:"
    $Global:lblADObj.Top = $Global:Top ; $Global:lblADObj.Left = 5; $Global:lblADObj.Width=120 ;$Global:lblADObj.AutoSize = $true
    $Global:form.Controls.Add($Global:lblADObj)    # Add to Form
    $Global:txtADObj = New-Object Windows.Forms.TextBox
    $Global:txtADObj.ReadOnly = $true;
    $Global:txtADObj.TabIndex = $Tab++ # set Tab Order
    $Global:txtADObj.Top = $Global:Top; $Global:txtADObj.Left = 130; $Global:txtADObj.Width = 540;
    $Global:txtADObj.Text = $AdUsr.DistinguishedName
    $Global:form.Controls.Add($Global:txtADObj)    # Add to Form

    #Do not display to form if there is no mailbox
    If ($NoLic -ne "Y")
    {
        $Global:Top = $Global:Top + 30
        ##PrimaryEmailAddress
        $Global:lblPrimayAddr = New-Object System.Windows.Forms.Label
        $Global:lblPrimayAddr.Text = "Primary Email Address:"
        $Global:lblPrimayAddr.Top = $Global:Top ; $Global:lblPrimayAddr.Left = 5; $Global:lblPrimayAddr.Width=120 ;$Global:lblPrimayAddr.AutoSize = $true
        $Global:form.Controls.Add($Global:lblPrimayAddr)    # Add to Form
        $Global:txtPrimayAddr = New-Object Windows.Forms.TextBox
        $Global:txtPrimayAddr.ReadOnly = $true;
        $Global:txtPrimayAddr.TabIndex = $Tab++ # set Tab Order
        $Global:txtPrimayAddr.Top = $Global:Top; $Global:txtPrimayAddr.Left = 130; $Global:txtPrimayAddr.Width = 200;
        $Global:txtPrimayAddr.Text = $Mbx.PrimarySmtpAddress
        $Global:form.Controls.Add($Global:txtPrimayAddr)    # Add to Form

        ##SIPAddress
        $Global:lblSIPAddr = New-Object System.Windows.Forms.Label
        $Global:lblSIPAddr.Text = "SIP Address:"
        $Global:lblSIPAddr.Top = $Global:Top ; $Global:lblSIPAddr.Left = 380; $Global:lblSIPAddr.Width=120 ;$Global:lblSIPAddr.AutoSize = $true
        $Global:form.Controls.Add($Global:lblSIPAddr)    # Add to Form
        $Global:txtSIPAddr = New-Object Windows.Forms.TextBox
        $Global:txtSIPAddr.ReadOnly = $true;
        $Global:txtSIPAddr.TabIndex = $Tab++ # set Tab Order
        $Global:txtSIPAddr.Top = $Global:Top; $Global:txtSIPAddr.Left = 470; $Global:txtSIPAddr.Width = 200;
        $LocArray = $mbx.EmailAddresses.split(",")
        foreach ($element in $LocArray)
        {
            If ($element -like "SIP:*")
            {
                $Global:txtSIPAddr.Text = $element
            }
        }
        $Global:form.Controls.Add($Global:txtSIPAddr)    # Add to Form

        $Global:Top = $Global:Top + 30
        ##Addresses
        $Global:lblAddresses = New-Object System.Windows.Forms.Label 
        $Global:lblAddresses.Text = "Other Addresses:"
        $Global:lblAddresses.Top = $Global:Top ; $Global:lblAddresses.Left = 5; $Global:lblAddresses.Width=120 ;$Global:lblAddresses.AutoSize = $true
        $Global:form.Controls.Add($Global:lblAddresses)    # Add to Form
        $Global:txtAddresses = New-Object Windows.Forms.ListBox
        $Global:txtAddresses.TabIndex = $Tab++ # set Tab Order
        $Global:txtAddresses.Top = $Global:Top; $Global:txtAddresses.Left = 130; $Global:txtAddresses.Height = 70; $Global:txtAddresses.Width = 540;
        $Global:txtAddresses.BackColor = "LightGray"
        If ($mbx.EmailAddresses.Length -ne 0)
        {
            $LocArray = $mbx.EmailAddresses.split(",")
            foreach ($element in $LocArray)
            {
                # Loop through Azure list and add to listbox
                If (($element -ne ("SMTP:" + $Mbx.PrimarySmtpAddress)) -and ($element -notlike "SIP:*"))
                {
                    [void] $Global:txtAddresses.Items.Add($element.TrimStart())  # Add element to listbox
                }
            }
        }
        $Global:form.Controls.Add($Global:txtAddresses)    # Add to Form

        $Global:Top = $Global:Top + 80
        #OOOEnabled
        $Global:lblOOOEnabled = New-Object System.Windows.Forms.Label
        $Global:lblOOOEnabled.Text = "Out Of Office Enabled:"
        $Global:lblOOOEnabled.Top = $Global:Top ; $Global:lblOOOEnabled.Left = 5; $Global:lblOOOEnabled.Width=120 ;$Global:lblOOOEnabled.AutoSize = $true
        $Global:form.Controls.Add($Global:lblOOOEnabled)    # Add to Form
        $Global:txtOOOEnabled = New-Object Windows.Forms.TextBox
        $Global:txtOOOEnabled.TabIndex = $Tab++ # set Tab Order
        $Global:txtOOOEnabled.Top = $Global:Top; $Global:txtOOOEnabled.Left = 130; $Global:txtOOOEnabled.Width = 70;
        $Global:txtOOOEnabled.ReadOnly = $true;
        If ($MbxOOO.AutoReplyState -ne "Disabled")
        {
            $Global:txtOOOEnabled.ReadOnly = $false;
            $Global:txtOOOEnabled.BackColor = "LightGray"
            $Global:lblOOOEnabled.ForeColor = "Red"
            $Global:txtOOOEnabled.ForeColor = "Red"
        }
        $Global:txtOOOEnabled.Text = $MbxOOO.AutoReplyState
        $Global:form.Controls.Add($Global:txtOOOEnabled)    # Add to Form

        #ADObjectProtected
        $Global:lblADObjProt = New-Object System.Windows.Forms.Label
        $Global:lblADObjProt.Text = "AD Object Protected:"
        $Global:lblADObjProt.Top = $Global:Top ; $Global:lblADObjProt.Left = 230; $Global:lblADObjProt.Width=140 ;$Global:lblADObjProt.AutoSize = $true
        $Global:txtADObjProt = New-Object Windows.Forms.TextBox
        $Global:txtADObjProt.ReadOnly = $true;
        $Global:txtADObjProt.TabIndex = $Tab++ # set Tab Order
        $Global:txtADObjProt.Top = $Global:Top; $Global:txtADObjProt.Left = 390; $Global:txtADObjProt.Width = 70;
        If ($UsrDetails.ProtectedFromAccidentalDeletion -eq "True")
        {
            $Global:txtADObjProt.BackColor = "LightGray"
            $Global:txtADObjProt.ForeColor = "Red"
            $Global:lblADObjProt.ForeColor = "Red"
            $Global:DoNotPurge = "Yes"
        }
        $Global:txtADObjProt.Text = $UsrDetails.ProtectedFromAccidentalDeletion
        $Global:form.Controls.Add($Global:lblADObjProt)    # Add to Form
        $Global:form.Controls.Add($Global:txtADObjProt)    # Add to Form

        #MemberOf
        $Global:lblGrpMemOf = New-Object System.Windows.Forms.Label
        $Global:lblGrpMemOf.Text = "Groups Member Of:"
        $Global:lblGrpMemOf.Top = $Global:Top ; $Global:lblGrpMemOf.Left = 480; $Global:lblGrpMemOf.Width=110 ;$Global:lblGrpMemOf.AutoSize = $true
        $Global:form.Controls.Add($Global:lblGrpMemOf)    # Add to Form
        $Global:txtGrpMemOf = New-Object Windows.Forms.TextBox
        $Global:txtGrpMemOf.ReadOnly = $true;
        $Global:txtGrpMemOf.TabIndex = $Tab++ # set Tab Order
        $Global:txtGrpMemOf.Top = $Global:Top; $Global:txtGrpMemOf.Left = 600; $Global:txtGrpMemOf.Width = 70;
        $Global:txtGrpMemOf.Text = ($DLMember.count + ($ADGroupMem.Name.count-1))
        $Global:form.Controls.Add($Global:txtGrpMemOf)    # Add to Form

        $Global:Top = $Global:Top + 30
        ##Legal Hold Enabled
        $Global:lblLegalHold = New-Object System.Windows.Forms.Label
        $Global:lblLegalHold.Text = "Legal Hold Enabled:"
        $Global:lblLegalHold.Top = $Global:Top ; $Global:lblLegalHold.Left = 5; $Global:lblLegalHold.Width=120 ;$Global:lblLegalHold.AutoSize = $true
        $Global:txtLegalHold = New-Object Windows.Forms.TextBox
        $Global:txtLegalHold.ReadOnly = $true;
        If ($Mbx.LitigationHoldEnabled -eq "True")
        {
            $Global:txtLegalHold.BackColor = "LightGray"
            $Global:txtLegalHold.ForeColor = "Red"
            $Global:lblLegalHold.ForeColor = "Red"
            $Global:DoNotPurge = "Yes"
        }
        $Global:txtLegalHold.TabIndex = $Tab++ # set Tab Order
        $Global:txtLegalHold.Top = $Global:Top; $Global:txtLegalHold.Left = 130; $Global:txtLegalHold.Width = 70;
        $Global:txtLegalHold.Text = $Mbx.LitigationHoldEnabled
        $Global:form.Controls.Add($Global:lblLegalHold)    # Add to Form
        $Global:form.Controls.Add($Global:txtLegalHold)    # Add to Form

        #CalEvents
        $Global:lblCalEvent = New-Object System.Windows.Forms.Label
        $Global:lblCalEvent.Text = "Calendar Events Organized:"
        $Global:lblCalEvent.Top = $Global:Top ; $Global:lblCalEvent.Left = 230; $Global:lblCalEvent.Width=140 ;$Global:lblCalEvent.AutoSize = $true
        $Global:form.Controls.Add($Global:lblCalEvent)    # Add to Form
        $Global:txtCalEvent = New-Object Windows.Forms.TextBox
        $Global:txtCalEvent.ReadOnly = $true;
        $Global:txtCalEvent.TabIndex = $Tab++ # set Tab Order
        $Global:txtCalEvent.Top = $Global:Top; $Global:txtCalEvent.Left = 390; $Global:txtCalEvent.Width = 70;
        $Global:txtCalEvent.Text = $CalEven.count
        $Global:form.Controls.Add($Global:txtCalEvent)    # Add to Form

        #OwmerOf
        $Global:lblGrpOwnOf = New-Object System.Windows.Forms.Label
        $Global:lblGrpOwnOf.Text = "Groups Owner Of:"
        $Global:lblGrpOwnOf.Top = $Global:Top ; $Global:lblGrpOwnOf.Left = 480; $Global:lblGrpOwnOf.Width=110 ;$Global:lblGrpOwnOf.AutoSize = $true
        $Global:form.Controls.Add($Global:lblGrpOwnOf)    # Add to Form
        $Global:txtGrpOwnOf = New-Object Windows.Forms.TextBox
        $Global:txtGrpOwnOf.ReadOnly = $true;
        $Global:txtGrpOwnOf.TabIndex = $Tab++ # set Tab Order
        $Global:txtGrpOwnOf.Top = $Global:Top; $Global:txtGrpOwnOf.Left = 600; $Global:txtGrpOwnOf.Width = 70;
        $Global:txtGrpOwnOf.Text = $DLOwner.count
        $Global:form.Controls.Add($Global:txtGrpOwnOf)    # Add to Form

        $Global:Top = $Global:Top + 30
        #MobileDevicesConfigured
        $Global:lblMblDevice = New-Object System.Windows.Forms.Label
        $Global:lblMblDevice.Text = "No of Mobile Devices:"
        $Global:lblMblDevice.Top = $Global:Top ; $Global:lblMblDevice.Left = 5; $Global:lblMblDevice.Width=120 ;$Global:lblMblDevice.AutoSize = $true
        $Global:txtMblDevice = New-Object Windows.Forms.TextBox
        $Global:txtMblDevice.ReadOnly = $true;
        $Global:txtMblDevice.TabIndex = $Tab++ # set Tab Order
        $Global:txtMblDevice.Top = $Global:Top; $Global:txtMblDevice.Left = 130; $Global:txtMblDevice.Width = 70;
        $Global:txtMblDevice.Text = 0
        If (($MblDevice.count-4) -gt 0)
        {
            $Global:txtMblDevice.Text = $MblDevice.count-4
            $Global:txtMblDevice.BackColor = "LightGray"
            $Global:txtMblDevice.ForeColor = "Red"
            $Global:lblMblDevice.ForeColor = "Red"
        }
        $Global:form.Controls.Add($Global:lblMblDevice)    # Add to Form
        $Global:form.Controls.Add($Global:txtMblDevice)    # Add to Form

        #Hidden
        $Global:lblHidden = New-Object System.Windows.Forms.Label
        $Global:lblHidden.Text = "Hidden from Address Book:"
        $Global:lblHidden.Top = $Global:Top ; $Global:lblHidden.Left = 230; $Global:lblHidden.Width=140 ;$Global:lblHidden.AutoSize = $true
        $Global:txtHidden = New-Object Windows.Forms.TextBox
        $Global:txtHidden.ReadOnly = $true;
        $Global:txtHidden.TabIndex = $Tab++ # set Tab Order
        $Global:txtHidden.Top = $Global:Top; $Global:txtHidden.Left = 390; $Global:txtHidden.Width = 70;
        If (($Mbx.LitigationHoldEnabled -eq "True") -and ($Global:Hidden -ne "True"))
        {
            $Global:txtHidden.BackColor = "LightGray"
            $Global:txtHidden.ForeColor = "Red"
            $Global:lblHidden.ForeColor = "Red"
        }
        $Global:txtHidden.Text = $Global:Hidden
        $Global:form.Controls.Add($Global:lblHidden)    # Add to Form
        $Global:form.Controls.Add($Global:txtHidden)    # Add to Form

        #UMExtension
        $Global:lblUMExt = New-Object System.Windows.Forms.Label
        $Global:lblUMExt.Text = "Extension:"
        $Global:lblUMExt.Top = $Global:Top ; $Global:lblUMExt.Left = 480; $Global:lblUMExt.Width=110 ;$Global:lblUMExt.AutoSize = $true
        $Global:form.Controls.Add($Global:lblUMExt)    # Add to Form
        $Global:txtUMExt = New-Object Windows.Forms.TextBox
        $Global:txtUMExt.ReadOnly = $true;
        $Global:txtUMExt.TabIndex = $Tab++ # set Tab Order
        $Global:txtUMExt.Top = $Global:Top; $Global:txtUMExt.Left = 540; $Global:txtUMExt.Width = 130;
        $Global:txtUMExt.Text = "(None)"
        If ($ADUsr.'msRTCSIP-Line' -ne "")
        {
            $Global:txtUMExt.Text = $ADUsr.'msRTCSIP-Line'
        }
        $Global:form.Controls.Add($Global:txtUMExt)    # Add to Form

        $Global:Top = $Global:Top + 30
        ##MailboxPermissions
        $Global:lblMbxPermission = New-Object System.Windows.Forms.Label
        $Global:lblMbxPermission.Text = "Mailbox Permissions:"
        $Global:lblMbxPermission.Top = $Global:Top ; $Global:lblMbxPermission.Left = 5; $Global:lblMbxPermission.Width=120 ;$Global:lblMbxPermission.AutoSize = $true
        $Global:form.Controls.Add($Global:lblMbxPermission)    # Add to Form
        $Global:txtMbxPermission = New-Object Windows.Forms.ListBox
        $Global:txtMbxPermission.TabIndex = $Tab++ # set Tab Order
        $Global:txtMbxPermission.Top = $Global:Top; $Global:txtMbxPermission.Left = 130; $Global:txtMbxPermission.Height = 50; $Global:txtMbxPermission.Width = 540;
        $Global:txtMbxPermission.BackColor = "LightGray"
        $Global:RemoveMbxPerm = ""
        [void] $Global:txtMbxPermission.Items.Add("User -->  AccessRights")
        foreach ($element in $MbxPerm)
        {
            # Loop through Azure list and add to listbox
            [void] $Global:txtMbxPermission.Items.Add($element.user + " --> " + $element.AccessRights)  # Add element to listbox
            If ($element.User -notlike "*SELF*")
            {
                $Global:RemoveMbxPerm = "Yes"
            }
        }
        $Global:form.Controls.Add($Global:txtMbxPermission)    # Add to Form

        $Global:Top = $Global:Top + 50
        ##InboxPermissions
        $Global:lblInboxPermission = New-Object System.Windows.Forms.Label 
        $Global:lblInboxPermission.Text = "Inbox Permissions:"
        $Global:lblInboxPermission.Top = $Global:Top ; $Global:lblInboxPermission.Left = 5; $Global:lblInboxPermission.Width=120 ;$Global:lblInboxPermission.AutoSize = $true
        $Global:form.Controls.Add($Global:lblInboxPermission)    # Add to Form
        $Global:txtInboxPermission = New-Object Windows.Forms.ListBox
        $Global:txtInboxPermission.TabIndex = $Tab++ # set Tab Order
        $Global:txtInboxPermission.Top = $Global:Top; $Global:txtInboxPermission.Left = 130; $Global:txtInboxPermission.Height = 50; $Global:txtInboxPermission.Width = 540;
        $Global:txtInboxPermission.BackColor = "LightGray"
        $Global:RemoveFldrPerm = ""
        [void] $Global:txtInboxPermission.Items.Add("User -->  AccessRights")
        foreach ($element in $MbxFolderPerm)
        {
            # Loop through Azure list and add to listbox
            [void] $Global:txtInboxPermission.Items.Add($element.User.DisplayName + " --> " + $element.AccessRights) # Add element to listbox
            If ($element.AccessRights -ne "None")
            {
                $Global:RemoveFldrPerm = "Yes"
            }
        }
        $Global:form.Controls.Add($Global:txtInboxPermission)    # Add to Form

        $Global:Top = $Global:Top + 60
        #Exchange GUID
        $Global:lblExchgGUID = New-Object System.Windows.Forms.Label
        $Global:lblExchgGUID.Text = "Exchange GUID:"
        $Global:lblExchgGUID.Top = $Global:Top ; $Global:lblExchgGUID.Left = 5; $Global:lblExchgGUID.Width=120 ;$Global:lblExchgGUID.AutoSize = $true
        $Global:form.Controls.Add($Global:lblExchgGUID)    # Add to Form
        $Global:txtExchgGUID = New-Object Windows.Forms.TextBox
        $Global:txtExchgGUID.ReadOnly = $true;
        $Global:txtExchgGUID.TabIndex = $Tab++ # set Tab Order
        $Global:txtExchgGUID.Top = $Global:Top; $Global:txtExchgGUID.Left = 130; $Global:txtExchgGUID.Width = 540;
        $Global:txtExchgGUID.Text = $Mbx.ExchangeGuid
        $Global:form.Controls.Add($Global:txtExchgGUID)    # Add to Form
    }
    else
    {
        $form.Height = $form.Height - 200
        $Global:Top = $Global:Top + 30
        ##Show No Mailbox
        $Global:lblNoMailbox = New-Object System.Windows.Forms.Label
        $Global:lblNoMailbox.Text = "O365 Mailbox:"
        $Global:lblNoMailbox.Top = $Global:Top ; $Global:lblNoMailbox.Left = 5; $Global:lblNoMailbox.Width=120 ;$Global:lblNoMailbox.AutoSize = $true
        $Global:form.Controls.Add($Global:lblNoMailbox)    # Add to Form
        $Global:txtNoMailbox = New-Object Windows.Forms.TextBox
        $Global:txtNoMailbox.ReadOnly = $true;
        $Global:txtNoMailbox.TabIndex = $Tab++ # set Tab Order
        $Global:txtNoMailbox.Top = $Global:Top; $Global:txtNoMailbox.Left = 130; $Global:txtNoMailbox.Width = 200;
        $Global:txtNoMailbox.Text = "No Mailbox Ever Configured"
        $Global:form.Controls.Add($Global:txtNoMailbox)    # Add to Form
        
        $Global:Top = $Global:Top + 30
        #show group memberhip for AD groups only
        #MemberOf
        $Global:lblGrpMemOf = New-Object System.Windows.Forms.Label
        $Global:lblGrpMemOf.Text = "AD Groups Member Of:"
        $Global:lblGrpMemOf.Top = $Global:Top ; $Global:lblGrpMemOf.Left = 5; $Global:lblGrpMemOf.Width=120 ;$Global:lblGrpMemOf.AutoSize = $true
        $Global:form.Controls.Add($Global:lblGrpMemOf)    # Add to Form
        $Global:txtGrpMemOf = New-Object Windows.Forms.TextBox
        $Global:txtGrpMemOf.ReadOnly = $true;
        $Global:txtGrpMemOf.TabIndex = $Tab++ # set Tab Order
        $Global:txtGrpMemOf.Top = $Global:Top; $Global:txtGrpMemOf.Left = 130; $Global:txtGrpMemOf.Width = 70;
        $Global:txtGrpMemOf.Text = ($ADGroupMem.Name.count - 1)
        $Global:form.Controls.Add($Global:txtGrpMemOf)    # Add to Form
    }

    $Global:Top = $Global:Top + 30
    #Active Object ID
    $Global:lblADObjID = New-Object System.Windows.Forms.Label
    $Global:lblADObjID.Text = "MSOL Object ID:"
    $Global:lblADObjID.Top = $Global:Top ; $Global:lblADObjID.Left = 5; $Global:lblADObjID.Width=120 ;$Global:lblADObjID.AutoSize = $true
    $Global:form.Controls.Add($Global:lblADObjID)    # Add to Form
    $Global:txtADObjID = New-Object Windows.Forms.TextBox
    $Global:txtADObjID.ReadOnly = $true;
    $Global:txtADObjID.TabIndex = $Tab++ # set Tab Order
    $Global:txtADObjID.Top = $Global:Top; $Global:txtADObjID.Left = 130; $Global:txtADObjID.Width = 540;
    $Global:txtADObjID.Text = $UserLicense.ObjectID
    $Global:form.Controls.Add($Global:txtADObjID)    # Add to Form

    $Global:Top = $Global:Top + 50
    ##ReportDetails
    $Global:lblRptFile = New-Object System.Windows.Forms.Label
    $Global:lblRptFile.Text = "Report Details:"
    $Global:lblRptFile.Top = $Global:Top ; $Global:lblRptFile.Left = 5; $Global:lblRptFile.Width=100 ;$Global:lblRptFile.AutoSize = $true
    $Global:form.Controls.Add($Global:lblRptFile)    # Add to Form
    $Global:txtRptFile = New-Object Windows.Forms.TextBox
    $Global:txtRptFile.ReadOnly = $true;
    $Global:txtRptFile.TabIndex = $Tab++ # set Tab Order
    $Global:txtRptFile.Top = $Global:Top; $Global:txtRptFile.Left = 130; $Global:txtRptFile.Width = 540;
    $Global:txtRptFile.Text = $ReportFile
    $Global:form.Controls.Add($Global:txtRptFile)    # Add to Form
    $Global:Top = $Global:Top + 20
}

function Enter-EmpNoInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow"
    $Global:form.Text = "Employee No."
    $Global:form.Size = New-Object System.Drawing.Size(450,175) #(W,H)
    $Global:form.StartPosition = "CenterScreen"
    Add-FormStandardButtons
    ## Employee Number
    $Global:lblEmpNo = New-Object System.Windows.Forms.Label   
        $Global:lblEmpNo.Text = "Employee No.:"  
        $Global:lblEmpNo.Top = 20 ; $Global:lblEmpNo.Left = 10; $Global:lblEmpNo.Width=120 ; $Global:lblEmpNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblEmpNo)    # Add to Form 
        # 
        $Global:txtInpEmpNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpEmpNo.Top = 20; $Global:txtInpEmpNo.Left = 140; $Global:txtInpEmpNo.Width = 200;  
        $Global:txtInpEmpNo.Text = ""
        $Global:form.Controls.Add($Global:txtInpEmpNo)    # Add to Form
        $Global:InputFocus = $Global:txtInpEmpNo

    If (($Global:chkPurgeAct.Checked -eq "Checked") -or ($Global:chkTermination.Checked -eq "Checked") -or ($Global:chkEnabUsr.Checked -eq "Checked"))
    {
        ## Ticket Number
        $Global:lblTicketNo = New-Object System.Windows.Forms.Label   
            $Global:lblTicketNo.Text = "Ticket No.:"  
            $Global:lblTicketNo.Top = 50 ; $Global:lblTicketNo.Left = 10; $Global:lblTicketNo.Width=120 ; $Global:lblTicketNo.AutoSize = $true
            $Global:form.Controls.Add($Global:lblTicketNo)    # Add to Form 
            # 
            $Global:txtInpTicketNo = New-Object Windows.Forms.TextBox  
            $Global:txtInpTicketNo.Top = 50; $Global:txtInpTicketNo.Left = 140; $Global:txtInpTicketNo.Width = 200;  
            $Global:txtInpTicketNo.Text = ""
            $Global:form.Controls.Add($Global:txtInpTicketNo)    # Add to Form
    }

    If ($Global:chkEnabUsr.Checked -eq "Checked")
    {
        $Global:form.Height = $Global:form.Height + 60
        # StartDate
        $Global:lblDatePicker = New-Object System.Windows.Forms.Label
        $Global:lblDatePicker.Text = "Start Date:"
        $Global:lblDatePicker.Top = 80 ; $Global:lblDatePicker.Left = 10; $Global:lblDatePicker.Width=120 ; $Global:lblDatePicker.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDatePicker)

        # DatePicker
        $Global:txtDatePicker = New-Object System.Windows.Forms.DateTimePicker
        $Global:txtDatePicker.Top = 80; $Global:txtDatePicker.Left = 140; $Global:txtDatePicker.Width = 200; 
        $Global:txtDatePicker.Format = [windows.forms.datetimepickerFormat]::custom
        $Global:txtDatePicker.CustomFormat = "dd/MMM/yyyy"
        $Global:txtDatePicker.Text = (get-date)
        $Global:form.Controls.Add($Global:txtDatePicker)

        ## Forwarder for M&A Only
        $Global:lblForwarder = New-Object System.Windows.Forms.Label   
            $Global:lblForwarder.Text = "Forwarder (for M&A Only):"  
            $Global:lblForwarder.Top = 110 ; $Global:lblForwarder.Left = 10; $Global:lblForwarder.Width=120 ; $Global:lblForwarder.AutoSize = $true
            $Global:form.Controls.Add($Global:lblForwarder)    # Add to Form 
            # 
            $Global:txtForwarder = New-Object Windows.Forms.TextBox  
            $Global:txtForwarder.Top = 110; $Global:txtForwarder.Left = 140; $Global:txtForwarder.Width = 200;  
            $Global:txtForwarder.Text = ""
            $Global:form.Controls.Add($Global:txtForwarder)    # Add to Form
    }

    $Global:chkEmerTerm = ""
    If ($Global:chkTermination.Checked -eq "Checked")
    {
        $Global:form.Height = $Global:form.Height + 30
        #Add Emergency Termination checkbox
        $Global:chkEmerTerm = New-Object Windows.Forms.checkbox 
        $Global:chkEmerTerm.Left = 140; $Global:chkEmerTerm.Width = 250; $Global:chkEmerTerm.Top = 80
        $Global:chkEmerTerm.Text = "Emergency Termination" 
        $Global:chkEmerTerm.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkEmerTerm) 
        # Obtain Value with: $Global:chkEmerTerm.Checked
    }

    If ($Global:chkOOOMsg.Checked -eq "Checked")
    {
        $Global:form.Height = $Global:form.Height + 30
        #Add Radio Box for Standard OOO or custom OOO
        $Global:chkOOOStd = New-Object Windows.Forms.RadioButton
        $Global:chkOOOStd.Left = 140; $Global:chkOOOStd.Width = 250; $Global:chkOOOStd.Top = 50
        $Global:chkOOOStd.Text = "Standard Out Of Office"
        $Global:chkOOOStd.Checked = $true   # set a default value 
        $Global:form.Controls.Add($Global:chkOOOStd) 
        # Obtain Value with: $Global:chkOOOStd.Checked
        $Global:chkoooStd.Add_MouseClick(
        {
#            $Global:form.Controls.Remove($Global:lblOOOCustAddr)
            $Global:form.Controls.Remove($Global:txtOOOCustAddr)
        })

        $Global:form.Height = $Global:form.Height + 30
        #Add Radio Box for Standard OOO or custom OOO
        $Global:chkOOOCust = New-Object Windows.Forms.RadioButton
        $Global:chkOOOCust.Left = 140; $Global:chkOOOCust.Width = 250; $Global:chkOOOCust.Top = 70
        $Global:chkOOOCust.Text = "Custom Out Of Office"
        $Global:chkOOOCust.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkOOOCust)
        $Global:chkooocust.Add_MouseClick(
        {
            $Global:txtOOOCustAddr = New-Object Windows.Forms.TextBox  
            $Global:txtOOOCustAddr.Top = 100; $Global:txtOOOCustAddr.Left = 140; $Global:txtOOOCustAddr.Width = 200;
            $Global:txtoooCustAddr.MaxLength = 1000  
            $Global:txtOOOCustAddr.Text = "(Enter Address)"
            $Global:form.Controls.Add($Global:txtOOOCustAddr)    # Add to Form
            $Global:InputFocus = $Global:txtOOOCustAddr
        })
        # Obtain Value with: $Global:chkchkOOOCust.Checked
    }
}

Function Execute-HideMbx
{
    write-host "Setting user to be Hidden from the Address Book..." -ForegroundColor Yellow
    $Global:u.msExchHideFromAddressLists.value = $True
    $Global:u.CommitChanges()
    $LineToWrite = $RecordEvent + "CHG" + "`t" + "Setting user to be Hidden from the Address Book"
    WriteReportEvent
}

function Enter-GroupInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow"
    $Global:form.Text = "Group Name"
    $Global:form.Size = New-Object System.Drawing.Size(450,175) #(W,H)
    $Global:form.StartPosition = "CenterScreen"
    Add-FormStandardButtons
    ## Mailbox
    $Global:lblGroup = New-Object System.Windows.Forms.Label   
        $Global:lblGroup.Text = "Group Name:"  
        $Global:lblGroup.Top = 30 ; $Global:lblGroup.Left = 10; $Global:lblGroup.Width=120 ; $Global:lblGroup.AutoSize = $true
        $Global:form.Controls.Add($Global:lblGroup)    # Add to Form 
        # 
        $Global:txtInpGroup = New-Object Windows.Forms.TextBox
        $Global:txtInpGroup.Top = 30; $Global:txtInpGroup.Left = 140; $Global:txtInpGroup.Width = 200;  
        $Global:txtInpGroup.Text = ""
        $Global:form.Controls.Add($Global:txtInpGroup)    # Add to Form
        $Global:InputFocus = $Global:txtInpGroup
}

function Enter-MbxInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow"
    $Global:form.Text = "Mailbox Name/Address"
    $Global:form.Size = New-Object System.Drawing.Size(450,175) #(W,H)
    $Global:form.StartPosition = "CenterScreen"
    Add-FormStandardButtons
    ## Mailbox
    $Global:lblMailbox = New-Object System.Windows.Forms.Label   
        $Global:lblMailbox.Text = "Mailbox Name/Address:"  
        $Global:lblMailbox.Top = 30 ; $Global:lblMailbox.Left = 10; $Global:lblMailbox.Width=120 ; $Global:lblMailbox.AutoSize = $true
        $Global:form.Controls.Add($Global:lblMailbox)    # Add to Form 
        # 
        $Global:txtInpMailbox = New-Object Windows.Forms.TextBox
        $Global:txtInpMailbox.Top = 30; $Global:txtInpMailbox.Left = 140; $Global:txtInpMailbox.Width = 200;  
        $Global:txtInpMailbox.Text = ""
        $Global:form.Controls.Add($Global:txtInpMailbox)    # Add to Form
        $Global:InputFocus = $Global:txtInpMailbox
}

Function GetAcctInfo($ENo)
{
    $Global:ADCmt = ""
    $Global:u = ""
    $ADExists = [bool]($strDN = GetUserDN $ENo -ErrorAction SilentlyContinue)
    If ($ADExists -eq "True")
    {
       $Global:UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion
     }

    $MbxExists = [bool](get-mailbox -identity $ENo -ErrorAction SilentlyContinue)
    If ($MbxExists -eq "True")
    {
        $Global:inf = get-mailbox $ENo
        If ([bool](get-ADUser -Filter {SamAccountName -eq $ENo} -ErrorAction SilentlyContinue))
        {
            $Global:strUserPath = [string]::format("LDAP://{0}", $strDN)
            $Global:u = new-object System.DirectoryServices.DirectoryEntry($Global:strUserPath)
            $Global:ADCmt = $Global:u.ExtensionAttribute14.value
            $Global:Hidden = $Global:u.msExchHideFromAddressLists.value
        }
        else
        {
            $Global:ADCmt = "*****No Active Directory Account For This User*****"
            $Global:strUserPath = "No Active Directory Account for this User"            
        }

        If ($Global:FormRefresh -ne "Y")
        {
            If (($Global:GeneralTitle -notlike "Purge*") -and ($Global:GeneralTitle -notlike "*Termination*"))
            {
                ShowDetails($Global:ENo)
                $LineToWrite = "`n"
                WriteReportEvent
            }
        }
    }
    else
    {
        write-host "Mailbox does not exist for Emp#" $ENo -ForegroundColor Red
    }
}

# This function connects to Active Directory and gets the record for the user 
Function GetUserDN($strUID)
{
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = "LDAP://DC=global,DC=ul,DC=com"
    $objSearcher.Filter = [string]::format("(sAMAccountName={0})",$strUID)
    $ux = $null
    $ux = $objSearcher.FindOne()
    
    if ($null -eq $ux)
    {
        return $null
    } 
    else
    {
        return $ux.Properties.distinguishedname
    }
}

Function Publish-Form
{
    ## Finalize Form and Show Dialog
    $Global:form.Add_Shown( { $form.Activate(); $Global:InputFocus.Focus() } )  #Activate and Set Focus 
    $Global:Result = $Global:form.ShowDialog()          ## Show the form, and wait for the response
}

function Remove-ADAcct
{
    Remove-ADUser -Identity $Global:txtInpEmpNo.Text -Confirm:$false
    Write-Host "Removed Active Directory Account for" $Global:txtInpEmpNo.Text -ForegroundColor Cyan
	$LineToWrite = $RecordEvent + "INFO" + "`t" + "Removed Active Directory Account for: " + $Global:txtInpEmpNo.Text
	WriteReportEvent

    If (Get-MsolUser -UserPrincipalName $EmpNo -ErrorAction SilentlyContinue)
    {
        Remove-MsolUser -UserPrincipalName $EmpNo -Force
        write-host "Removing MSOLUser Account " $Global:txtInpEmpNo.Text -ForegroundColor Cyan
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Removed MSOLUserAccount for: " + $Global:txtInpEmpNo.Text
        WriteReportEvent
    }
}

Function Remove-DLMemberOwner
{
    $NeedOwners = ""
    If ($NoLic -ne "Y")
    {
        if ($DLOwner.count -gt 0)
        {
            write-host "`nThese are the O365 groups this person is a manager of:" -ForegroundColor Red
	        write-host "`tNumber of Groups this individual is the Owner of: " $DLOwner.Count
	        $LineToWrite = $RecordEvent + "OWNS" + "`t" + "Number of Groups this individual is the Owner of: " + $DLOwner.Count
            foreach ($DLOwner in $DLOwner)
            {
                If ($DLOwner.RecipientTypeDetails -like "*Universal*")
                {
                    If ($DLOwner.Name -notlike "DST*")
                    {
                        $FindOwner = (Get-DistributionGroup $DLOwner.Name).Managedby.count
                    }
                }
                else
                {
                    If ($DLOwner.RecipientTypeDetails -eq "GroupMailbox")
                    {
                        $FindOwner = (Get-UnifiedGroupLinks $DLOwner.Name -LinkType Owner).count
                    }
                    else
                    {
                        $FindOwner = (Get-DynamicDistributionGroup $DLOwner.Name).Managedby.count
                    }
                }

                If ($FindOwner -eq 1)
                {
                    If ($DLOwner.Name -notlike "DST*")
                    {
                        write-host "`tNeed O365 Group Owner Identified for: " $DLOwner.Name
                        $LineToWrite = $RecordEvent + "O365" + "`t" + "Need O365 Group Owner Identified for: " + $DLOwner.Name
                        $NeedOwners = $NeedOwners + "<br>" + $DLOwner.Name
                    }
                }
                else
                {
                    If ($DLOwner.RecipientTypeDetails -eq "GroupMailbox")
                    {
                        write-host "`tRemoving from the O365 Group ownership of: " $DLOwner.Name
                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing from the O365 ownership of: " + $DLOwner.Name
                        Remove-UnifiedGroupLinks $DLOwner.Name -LinkType Owners -Links $Global:txtInpEmpNo.Text -Confirm:$False
                    }
                    else
                    {
                        If  (($DLOwner.Name -notlike "DST*") -and ($DLOwner.Name -notlike "DSG*"))
                        {
                            If ($DLOwner.RecipientTypeDetails -notlike "*Dynamic*")
                            {
                                write-host "`tRemoving from the O365 DL ownership of: " $DLOwner.Name
                                $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing from the O365 DL ownership of: " + $DLOwner.Name
                                Set-DistributionGroup $DLOwner.Name -ManagedBy @{remove="$Global:txtInpEmpNo.Text"} -BypassSecurityGroupManagerCheck -Confirm:$False
                            }
                            else
                            {
                                write-host "`tRemoving from the O365 Dynamic ownership of: " $DLOwner.Name
                                $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing from the O365 Dynamic ownership of: " + $DLOwner.Name
                                Set-DynamicDistributionGroup $DLOwner.Name -ManagedBy $null -Confirm:$False
                            }
                        }
                }
                WriteReportEvent
                }
            }

            if ($NeedOwners -ne "")
            {
                $MsgTo = "EnterpriseMessagingServices@ul.com"
                $MsgCC = "Sandi.Glazebrook@ul.com"
                $Msg = "No managers of O365 DL's for Terminated User " + $Global:txtInpEmpNo.Text + "-" + (get-mailbox $Global:txtInpEmpNo.Text).Name
                $MsgBody = "<p>The subject user account has been terminated/purged and the below O365 groups now have no owners.<br>" + $NeedOwners
                Send-Message
            }
        }
        else
        {
            write-host "This person is not an owner any O365 groups" -ForegroundColor Red
            $LineToWrite = $RecordEvent + "OWNS" + "`t" + "This person is not an owner of any O365 groups"
	        WriteReportEvent
        }

        if ($DLMember.count -gt 0)
        {
            write-host "These are the O365 groups this person is a member of:" -ForegroundColor Red
	        write-host "`tNumber of Groups this individual is a Member of: " $DLMember.Count
	        $LineToWrite = $RecordEvent + "MEMB" + "`t" + "Number of Groups this individual is the Member of: " + $DLMember.Count
            foreach ($DLMember in $DLMember)
            {
                If (($DLMember.RecipientTypeDetails -like "*Universal*") -and (($DLMember.DisplayName -like "DST*") -or ($DLMember.DisplayName -like "DSG*")))
                {
                    write-host "`tRemoving from the O365 DL membership of: " $DLMember.DisplayName
		            $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing from the O365 DL membership of: " + $DLMember.DisplayName
		            WriteReportEvent
                    Remove-DistributionGroupMember $DLMember.DisplayName -Member $Global:txtInpEmpNo.Text -BypassSecurityGroupManagerCheck -Confirm:$False
                }
                else
                {
                    write-host "`tRemoving from the O365 Group membership of: " $DLMember.DisplayName
                    $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing from the O365 Group membership of: " + $DLMember.DisplayName
                    WriteReportEvent
                    Remove-UnifiedGroupLinks $DLMember.DisplayName -LinkType Members -Links $Global:txtInpEmpNo.Text -Confirm:$False
                }
            }
        }
        else
        {
            write-host "This person is not a member of any O365 groups" -ForegroundColor Red
            $LineToWrite = $RecordEvent + "MEMB" + "`t" + "This person is not a member of any O365 groups"
	        WriteReportEvent
        }
    }

    if ($ADGroupMem.count -gt 0)
    {
        write-host "These are the AD Groups this person is a member of:" -ForegroundColor Red
        foreach ($ADGroupMem in $ADGroupMem)
        {
            If (($ADGroupMem.Name -notlike "DST*") -and ($ADGroupMem.Name -notlike "DSG*") -and ($ADGroupMem.Name -notlike "*IntuneEnforced*"))
            {
                write-host "`tRemoving from the AD Group membership of: " $ADGroupMem.Name
		        $LineToWrite = $RecordEvent + "REMV" + "`t" + "Removing from the AD Group membership of: " + $ADGroupMem.Name
		        WriteReportEvent
		        $ErrorActionPreference = 'SilentlyContinue'
                Remove-ADGroupMember $ADGroupMem.Name -Members $Global:txtInpEmpNo.Text -Confirm:$False
		        $ErrorActionPreference = 'Continue'
            }
        }
    }
    else
    {
        write-host "This person is not a member of any AD Groups" -ForegroundColor Red
        $LineToWrite = $RecordEvent + "MEMB" + "`t" + "This person is not a member of any AD Groups"
	    WriteReportEvent
    }
}

Function Remove-OneDriveAccess
{
    $Site = Get-SPOSite ("https://ul-my.sharepoint.com/personal/"+ $Global:txtInpEmpNo.Text + "_global_ul_com")

    $Erroractionpreference = "SilentlyContinue"
    Get-SPOUser -Site $Site |Where-Object {$_.IsSiteAdmin -eq "True"} |ft Displayname,LoginName,IsSiteAdmin
    $SiteAdmins = Get-SPOUser -Site $Site | Where-Object {$_.IsSiteAdmin -eq "True"}
    $Erroractionpreference = "Continue"
    Out-File -filepath $ReportFile -append -noClobber -inputObject $SiteAdmins

    $RemAccess = "N"
    $Owner = "*" + $Global:txtInpEmpNo.Text + "*"

    foreach ($SiteAdmins in $SiteAdmins)
    {
        if (($SiteAdmins.LoginName -notlike $Owner) -and ($SiteAdmins.LoginName -notlike "*sharepoint*"))
        {
            $RemUser = $SiteAdmins.LoginName
            $RemAccess = "Y"
            write-host "Removing " + $SiteAdmins.DisplayName + " (" + $SiteAdmins.LoginName + ") OneDrive Access to " + $Usr
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Removing " + $SiteAdmins.DisplayName + " (" + $SiteAdmins.LoginName + ") OneDrive Access to " + $Usr
            WriteReportEvent
            Set-SPOUser -site $Site -LoginName $RemUser -IsSiteCollectionAdmin $False
        }
    }

    If ($RemAccess -eq "N")
    {
        write-host "No Additional Staff was granted access to OneDrive" -ForegroundColor Cyan
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "No Additional Staff was granted access to OneDrive"
        WriteReportEvent
    }
    else
    {
        $SiteAdmins = Get-SPOUser -Site $Site |ft Displayname,LoginName,IsSiteAdmin
        Out-File -filepath $ReportFile -append -noClobber -inputObject $SiteAdmins
    }
}

Function O365Licenses
{
    $O365Lic = (Get-MsolUser -UserPrincipalName $EmpNo).Licenses
	Write-Host "`nCurrent License Assignment for " $EmpNo " " $Global:UserLicense.DisplayName "-" $ADUser.ExtensionAttribute1-ForegroundColor Green
    
    If ($O365Lic.count -gt 0)
    {
        $LineToWrite = $RecordEvent + "DISP" + "`t" + "Reviewing O365 Licenses Assigned to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

        foreach ($O365Lic in $O365Lic)
        {
            switch ($O365Lic.AccountSkuID)
            {
                "ul:ENTERPRISEPACK"
                {
                    $text = "`tEnterprise E3"
                }
                "ul:EXCHANGEENTERPRISE"
                {
                    $text = "`tExchange Online Plan 2"
                }
                "ul:POWER_BI_STANDARD"
                {
                    $text = "`tPowerBI (Free)"
                }
                "ul:POWER_BI_PRO"
                {
                    $text = "`tPowerBI Pro" 
                }
                "ul:DYN365_ENTERPRISE_PLAN1"
                {
                    $text = "`tDynamics 365 Customer Engagement Plan"
                }
                "ul:EMS"
                {
                    $text = "`tEnterprise Mobility Suite/Intune (EMS)"
                }
                "ul:POWER_BI_INDIVIDUAL_USER"
                {
                    $text = "`tPower BI for O365"
                }
                "ul:MCOIMP"
                {
                   $text = "`tSkype for Business Online (Plan 1)"
                }
                "ul:ECAL_SERVICES"
                {
                    $text = "`tECAL Services (EOA, EOP, DLP)"
                }
                "ul:SHAREPOINTSTANDARD"
                {
                    $text = "`tSharePoint Online (Plan 1)"
                }
                "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
                {
                    $text = "`tMicrosoft Relationship Sales Solution"
                }
				"ul:ATP_ENTERPRISE"
				{
					$text = "`tAdvanced Threat Protection Plan1"
				}
                "ul:FLOW_P2"
                {
                    $text = "`tMicrosoft Flow (Plan 2)"
                }
                "ul:TEAMS_COMMERCIAL_TRIAL"
                {
                    $text = "Microsoft Teams Commercial Cloud (User Initiated)"
                }
                "ul:WIN_DEF_ATP"
                {
                    $text ="Defender Advanced Threat Protection"
                }
                "ul:MEETING_ROOM"
                {
                    $text = "`tMeeting Room"
                }
                "ul:POWERFLOW_P2"
                {
                    $text = "`tPowerApps Plan 2"
                }
                "ul:MCOEV"
                {
                    $text = "`tPhone System"
                }
                "ul:MCOMEETADV"
                {
                    $text = "`tAudio Conferencing"
                }
                "ul:PROJECTPREMIUM"
                {
                    $text = "`tProject Plan 5"
                }
            }
            Write-Host $text
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + $text
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        }
    }
    else
    {
        Write-Host "`n`tNo licenses assigned to this account" -ForegroundColor Red
        $LineToWrite = $RecordEvent + "DISP" + "`t" + "No O365 Licenses Assigned to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

    }
}

Function RetentPolicy
{
    $MbxCreated = [bool](get-mailbox $EmpNo -ErrorAction SilentlyContinue)

    If ($MbxCreated -eq "True")
    {
        If ($ADUser.extensionattribute4 -eq "IT")
        {
            Set-MailBox $EmpNo -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 3 yr Delete"
            $LineToWrite = $WhoAmI + "`t" + "UPDA" + "`t" + "Updating Retention Policy to IT Policy for " + $EmpNo
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        }
        else
        {
            Set-MailBox $EmpNo -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 365 Day Delete"
            Set-Mailbox $EmpNo –RetentionHoldEnabled $true –StartDateForRetentionHold 04/01/2011
            $LineToWrite = $WhoAmI + "`t" + "UPDA" + "`t" + "Updating Retention Policy to UL Default Policy for " + $EmpNo
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        }
        write-host "Setting mailbox quotas" -ForegroundColor Green
        Set-Mailbox $EmpNo -UseDatabaseQuotaDefaults $False -IssueWarningQuota 45GB -ProhibitSendQuota 49.75GB -ProhibitSendReceiveQuota 50GB
        write-host "Setting mailbox protocols" -ForegroundColor Green
        Set-CASMailBox $EmpNo -ImapEnabled $false -PopEnabled $false -ActiveSyncEnabled $false
    }
    else
    {
        write-host "Mailbox does not exist or is in the process of being created.  If the license was just assigned"
        write-host "please allow 3-5 minutes for the mailbox to be created and the select Option 10 from the menu"
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Mailbox Does Not Exist " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
}

Function Send-Message
{
    $server = "smtp-relay.ul.com"
    $client = new-object system.net.mail.smtpclient $server
    $from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "UL Account Provisioning Team"
    If ($MsgAltFrom -ne "")
    {
	    $from = $MsgAltFrom
    }
    $to = New-Object System.Net.Mail.MailAddress $MsgTo
    $message = new-object  System.Net.Mail.MailMessage $from, $to 
    $message.IsBodyHtml = $true
    $msgfont = "<basefont face=verdana size=2.5 color=black>"
    $message.To.Clear()
    $message.CC.Clear()
    $message.To.Add($MsgTo)
    If ($MsgCC -ne "")
    {
        $message.CC.Add($MsgCC)
    }
    $message.Subject = $MsgSubject
    $message.Body = $msgfont + $MsgBody
    $client.Send($message)
}

Function Set-DelegateAccess
{
    $DelegateTo = $Global:txtDelegateAccess.Text + "@global.ul.com"

    write-host "Setting permissions on: " $Global:txtInpEmpNo.Text " <-- " $DelegateTo "granting Read Permission " -ForegroundColor Cyan

    $First = "Y"
    $MBXFldrCnt = $MBXFldr.count
    write-host "The mailbox" $Global:txtInpEmpNo.TextNo "has" $MBXFldrCnt" folders" -ForegroundColor Cyan
	$LineToWrite = $RecordEvent + "PERM" + "`t" + "The mailbox " + $Global:txtInpEmpNo.Text + " has " + $MBXFldrCnt + " folder(s)" 
	WriteReportEvent

    Add-MailboxPermission $Global:txtInpEmpNo.Text -AccessRights ReadPermission -User $DelegateTo

    If (((Get-CASMailbox -Identity $Global:txtInpEmpNo.Text).OwaEnabled) -eq $False)
    {
        write-host "Enabling OWA Access..." -ForegroundColor Red
        Set-CASMailbox -Identity $Global:txtInpEmpNo.Text -OwaEnabled $True
        $LineToWrite = $RecordEvent + "ENAB" + "`t" + "Enabling OWA Access"
        WriteReportEvent
    }

    # Unhide from Address Book
    if ($Global:u.msExchHideFromAddressLists.value -eq $True)
    {
        write-host "Unhiding user from the Address Book..." -ForegroundColor Red
        $Global:u.msExchHideFromAddressLists.value = $False
        $Global:u.CommitChanges()	
        $LineToWrite = $RecordEvent + "CHNG" + "`t" + "Unhiding user from the Address Book "
        WriteReportEvent
    }
	
    ForEach ($Folder in $MBXFldr)
    {
        write-host "     Granting Read Access to: " $Folder -ForegroundColor Green -NoNewline
        If ($First -eq "Y")
        {
           $First = "N"
            If (Get-MailboxFolderPermission $Global:txtInpEmpNo.Text | where {$_.User -like ((get-mailbox $DelegateTo).DisplayName)})
            {
                $ac = Get-MailboxFolderPermission $Global:txtInpEmpNo.Text | where {$_.User -like ((get-mailbox $DelegateTo).DisplayName)}
                If ($ac.AccessRights -eq "Reviewer")
                {
                    write-host " <--" $DelegateTo "already has (" $ac.AccessRights ") access to this folder" -ForegroundColor Red
                    $LineToWrite = $RecordEvent + "PERM" + "`t" + $DelegateTo + " already has (" + $ac.AccessRights + ") access to the mailbox" 
                    WriteReportEvent
                }
                else
                {
                    If ($Global:u.extensionAttribute2.value -eq "T")
                    {
                        write-host " <--" $DelegateTo " has (" $ac.AccessRights ") access to this terminated users folder resetting to ReadOnly" -ForegroundColor Red
                        Remove-MailboxFolderPermission $Global:txtInpEmpNo.Text -User $DelegateTo -Confirm:$False |out-null
                        $LineToWrite = $RecordEvent + "PERM" + "`t" + $DelegateTo + " removed (" + $ac.AccessRights + ") access and granted ReadOnly access to the mailbox" 
                        WriteReportEvent
                        Add-MailboxFolderPermission -Identity ($Global:txtInpEmpNo.Text + ":\") -User $DelegateTo -AccessRights Reviewer -ErrorAction SilentlyContinue|out-null
                    }
                    else
                    {
                        If ($Global:u.extensionAttribute2.value -eq "A")
                        {
                            write-host " <--" $DelegateTo "already has (" $ac.AccessRights ") access to this Active Users mailbox" -ForegroundColor Red
                            $LineToWrite = $RecordEvent + "PERM" + "`t" + $DelegateTo + " already has (" + $ac.AccessRights + ") access to the Active Users folder" 
	                        WriteReportEvent
                        }
                        else
                        {
                            write-host " <-- Granting ReadOnly Access the this mailbox" -ForegroundColor Green
                            $LineToWrite = $RecordEvent + "PERM" + "`t" + $DelegateTo + " granted ReadOnly access to the mailbox"
                            Add-MailboxFolderPermission -Identity ($Global:txtInpEmpNo.Text + ":\") -User $DelegateTo -AccessRights Reviewer -ErrorAction SilentlyContinue|out-null
                        }
                    }
                    WriteReportEvent
                }
            }
            else
            {
                write-host " <-- Granting ReadOnly Access the this folder" -ForegroundColor Green
                Add-MailboxFolderPermission -Identity ($Global:txtInpEmpNo.Text + ":\") -User $DelegateTo -AccessRights Reviewer |out-null
                $LineToWrite = $RecordEvent + "PERM" + "`t" + $DelegateTo + " granted ReadOnly access to the mailbox" 
                WriteReportEvent
            }
	    }
	    else
        {
            If (Get-MailboxFolderPermission -Identity ($Global:txtInpEmpNo.Text + ":\" + $Folder) -ErrorAction SilentlyContinue | where {$_.User -like ((get-mailbox $DelegateTo).DisplayName)})
            {
                $ac = Get-MailboxFolderPermission -Identity ($Global:txtInpEmpNo.Text + ":\" + $Folder) -ErrorAction SilentlyContinue | where {$_.User -like ((get-mailbox $DelegateTo).DisplayName)}
                If ($ac.AccessRights -eq "Reviewer")
                {
                    write-host " <--" $DelegateTo "already has (" $ac.AccessRights ") access to this folder" -ForegroundColor Red
	                $LineToWrite = $RecordEvent + "PERM" + "`t" + $DelegateTo + " already has (" + $ac.AccessRights + ") access to the folder " + $Folder
	                WriteReportEvent
                }
                else
                {
                    If ($Global:u.extensionAttribute2.value -eq "T")
                    {
                        write-host " <--" $DelegateTo " has (" $ac.AccessRights ") access to this terminated users folder resetting to ReadOnly" -ForegroundColor Red
                        Remove-MailboxFolderPermission ($Global:txtInpEmpNo.Text + ":\" + $Folder) -User $DelegateTo -Confirm:$False |out-null
                        $LineToWrite = $RecordEvent + "PERM" + "`t" + $DelegateTo + " removed (" + $ac.AccessRights + ") access and granted ReadOnly access to the folder" 
                        WriteReportEvent
                        Add-MailboxFolderPermission -Identity ($Global:txtInpEmpNo.Text + ":\" + $Folder) -User $DelegateTo -AccessRights Reviewer -confirm:$False -ErrorAction SilentlyContinue |out-null
                    }
                    else
                    {
                        If ($Global:u.extensionAttribute2.value -eq "A")
                        {
                            write-host " <--" $DelegateTo " has (" $ac.AccessRights ") access to this Active users folder unable to perform reset" -ForegroundColor Cyan
                            $LineToWrite = $RecordEvent + "PERM" + "`t" + $DelegateTo + " has (" + $ac.AccessRights + ") access to this Active Users folder cannot reset" 
                        }
                        else
                        {
                            write-host " <-- Granting ReadOnly Access the this folder" -ForegroundColor Green
                            $LineToWrite = $RecordEvent + "PERM" + "`t" + $DelegateTo + " granted ReadOnly access to the folder"
                            Add-MailboxFolderPermission -Identity ($Global:txtInpEmpNo.Text + ":\" + $Folder) -User $DelegateTo -AccessRights Reviewer -confirm:$False -ErrorAction SilentlyContinue |out-null
                        }
                    }
                    WriteReportEvent
                }
            }
            else
            {
                write-host " <-- Granting ReadOnly Access the this folder" -ForegroundColor Green
                Add-MailboxFolderPermission -Identity ($Global:txtInpEmpNo.Text + ":\" + $Folder) -User $DelegateTo -AccessRights Reviewer -confirm:$False -ErrorAction SilentlyContinue |out-null
	            $LineToWrite = $RecordEvent + "PERM" + "`t" + $DelegateTo + " granted ReadOnly access to the folder " + $Folder 
                WriteReportEvent
            }
		}
	}
}

Function StandardLicenses
{
    If ($HasEMS -eq $False)
	{
	    Write-Host "`nAssigning Enterprise Mobility Suite/Intune (EMS) License to" $EmpNo -ForegroundColor Yellow
		Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:EMS"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Enterprise Mobility Suite/Intune (EMS) License to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}  

    If (($HasBIFree -eq $False) -and ($HasBIPro -eq $False))
    {
	    Write-Host "`nAssigning PowerBI Free License to" $EmpNo -ForegroundColor Yellow
	    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:POWER_BI_STANDARD"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning PowerBI Free License to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}

    If ($HasATPP1 -eq $False)
    {
	    Write-Host "`nAssigning Advanced Threat Protection Plan1 License to" $EmpNo -ForegroundColor Yellow
	    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:ATP_ENTERPRISE"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Advanced Threat Protection Plan1 License to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}

    if ($HasPhone -eq $False)
    {
        Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicenses "ul:MCODEV" -ErrorAction SilentlyContinue
	    $LineToWrite = $RecordEvent + "INFO" + "`t" + "     Assigned Phone System License to " + $User.UPN
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "     Assigned Phone System License"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite		
    }
}

Function Write-AccountDetails
{
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "            Ticket Number: " + $Global:txtInpTicketNo.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "             Employee No.: " + $Global:txtHost.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "            Employee Name: " + $Global:txtName.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "            Employee Type: " + $Global:txtType.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "       AD Account Enabled: " + $Global:txtADEnabled.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "            Admin Account: " + $Global:txtAdmAcct.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "       AD Object Location: " + $Global:txtADObj.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "      Configure Forwarder: " + $Global:txtForwarder.Text
    WriteReportEvent
    If ($NoLic -ne "Y")
    {
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "    Primary Email Address: " + $Global:txtPrimayAddr.Text
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "              SIP Address: " + $Global:txtSIPAddr.Text
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "          Other Addresses: " + ($Global:txtAddresses.Items -join "`n`t`t`t`t`t`t`t`t`t")
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "              OOO Enabled: " + $Global:txtOOOEnabled.Text
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "      AD Object Protected: " + $Global:txtADObjProt.Text
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "       Legal Hold Enabled: " + $Global:txtLegalHold.Text
        WriteReportEvent
        If ($Global:txtLegalHold.Text -eq $True)
        {
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "  Legal Hold Enabled Date: " + (get-mailbox $Global:txtHost.Text).LitigationHoldDate
            WriteReportEvent
        }
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "      # of Mobile Devices: " + $Global:txtMblDevice.Text
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Calendar Events Organized: " + $Global:txtCalEvent.Text
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + " Hidden from Address Book: " + $Global:txtHidden.Text
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "     # of Mailbox Folders: " + $MbxFldr.Count
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "       # Groups Member of: " + $Global:txtGrpMemOf.Text
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "        # Groups Owner of: " + $Global:txtGrpOwnOf.Text
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "             UM Extension: " + $Global:txtUMExt.Text
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "       Account Purge Date: " + (Get-ADUser $Global:txtHost.Text -properties *)."msDS-cloudExtensionAttribute1"
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "      Mailbox Permissions: " + ($Global:txtMbxPermission.Items -join "`n`t`t`t`t`t`t`t`t`t")
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "        Inbox Permissions: " + ($Global:txtInboxPermission.Items -join "`n`t`t`t`t`t`t`t`t`t")
        WriteReportEvent
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "            Exchange GUID: " + $Global:txtExchgGUID.Text
        WriteReportEvent
    }
    else
    {
        $LineToWrite = $RecordEvent + "INFO" + "`t" + " No mailbox Ever licensed: UsageLocation was never set"
        WriteReportEvent
    }
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "           MSOL Object ID: " + $Global:txtADObjID.Text
    WriteReportEvent
}

function WriteLogEvent
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LocalMachineName + "`t" + $LineToWrite
	Out-File -filepath $LogFile -append -noClobber -inputObject $RecordEvent
}

function WriteReportEvent 
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LineToWrite
	Out-File -filepath $ReportFile -append -noClobber -inputObject $RecordEvent
}

####Start Script

$Global:LiveCred = ""
$Global:CredFile = ""
$File = "e:\SDAP\scripts\SDAPAdminMenu.ps1"

if (test-path $file)
{
#   Do nothing E" drive exists
}
else
{
    New-PSDrive -Name E -PSProvider FileSystem -Root \\USNBKEMES100P\e$
}

If ((get-location) -notlike "*SDAPNew\*")
{
    Set-Location E:\SDAP\Scripts
    write-host ""
}

# Retrieve the user name
	$WhoAmI			= WhoAmI
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name

$wshell = New-Object -ComObject Wscript.Shell
$Global:OKDetails = ""

If ($WhoAmI -notlike "*96151")
{
    Invoke-Expression -Command .\SDAPMainMenu.ps1
}
else
{
   Invoke-Expression -Command .\SDAPMainMenu-New.ps1
}