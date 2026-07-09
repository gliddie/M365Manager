#
#  SDAP Team Main Menu
#

function Build-AdminMenuForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "SDAP Admin Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 750 ; $form.Height = 550  # Make the form wider 
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    Add-FormStandardButtons

    #Build Column1
    $LeftCol1 = 60
    $TitleCol1 = 40
    $TopCol1 = 20

## Title Line
    $Global:lblTitleLine1 = New-Object System.Windows.Forms.Label   
        $lblTitleLine1.Text = "Account Activities:"
        $lblTitleLine1.Top = $TopCol1 ; $lblTitleLine1.Left = $TitleCol1; $lblTitleLine1.Width=120 ;$lblTitleLine1.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine1)    # Add to Form 

    ## Enable User Account
    $TopCol1 = $TopCol1 + 20
    $Global:chkEnabUsr = New-Object Windows.Forms.RadioButton 
        $Global:chkEnabUsr.Left = $LeftCol1; $Global:chkEnabUsr.Width = 240; $Global:chkEnabUsr.Top = $TopCol1  
        $Global:chkEnabUsr.Text = "Enable User Account/Configure Mailbox" 
        $Global:chkEnabUsr.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkEnabUsr) 
        # Obtain Value with: $Global:chkEnabUsr.Checked

    ## New Service Account
    $TopCol1 = $TopCol1 + 20
    $Global:chkGenAcct = New-Object Windows.Forms.RadioButton 
        $Global:chkGenAcct.Left = $LeftCol1; $Global:chkGenAcct.Width = 240; $Global:chkGenAcct.Top = $TopCol1  
        $Global:chkGenAcct.Text = "Create New Generic User/Service Account" 
        $Global:chkGenAcct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkGenAcct) 
        # Obtain Value with: $Global:chkGenAcct.Checked

    ## Update Service Account
    $TopCol1 = $TopCol1 + 20
    $Global:chkUpdSvcAcct = New-Object Windows.Forms.RadioButton
        $Global:chkUpdSvcAcct.Left = $LeftCol1; $Global:chkUpdSvcAcct.Width = 240; $Global:chkUpdSvcAcct.Top = $TopCol1
        $Global:chkUpdSvcAcct.Text = "Update Generic Account Details" 
        $Global:chkUPdSvcAcct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkUpdSvcAcct) 

    $TopCol1 = $TopCol1 + 40
    $Global:lblTitleLine3 = New-Object System.Windows.Forms.Label   
        $lblTitleLine3.Text = "Termination Activities:"
        $lblTitleLine3.Top = $TopCol1 ; $lblTitleLine3.Left = $TitleCol1; $lblTitleLine3.Width=120 ;$lblTitleLine3.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine3)    # Add to Form 

    ## Standard Termination
    $TopCol1 = $TopCol1 + 20
    $Global:chkTermination = New-Object Windows.Forms.RadioButton 
        $Global:chkTermination.Left = $LeftCol1; $Global:chkTermination.Width = 250; $Global:chkTermination.Top = $TopCol1  
        $Global:chkTermination.Text = "User Termination (Standard or Emergency)" 
        $Global:chkTermination.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkTermination) 
        # Obtain Value with: $Global:chkTermination.Checked

    ##  Purge Account
    $TopCol1 = $TopCol1 + 20
    $Global:chkPurgeAct = New-Object Windows.Forms.RadioButton 
        $Global:chkPurgeAct.Left = $LeftCol1; $Global:chkPurgeAct.Width = 200; $Global:chkPurgeAct.Top = $TopCol1  
        $Global:chkPurgeAct.Text = "Purge User Account" 
        $Global:chkPurgeAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkPurgeAct) 
        # Obtain Value with: $Global:chkPurgeAct.Checked

    $TopCol1 = $TopCol1 + 40
    $Global:lblTitleLine2 = New-Object System.Windows.Forms.Label   
        $lblTitleLine2.Text = "Rehire Activities:"
        $lblTitleLine2.Top = $TopCol1 ; $lblTitleLine2.Left = $TitleCol1; $lblTitleLine2.Width=120 ;$lblTitleLine2.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine2)    # Add to Form

    ## Rehire
    $TopCol1 = $TopCol1 + 20
    $Global:chkRehire = New-Object Windows.Forms.RadioButton 
        $Global:chkRehire.Left = $LeftCol1; $Global:chkRehire.Width = 200; $Global:chkRehire.Top = $TopCol1  
        $Global:chkRehire.Text = "Reconfigure Account for a Rehire" 
        $Global:chkRehire.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkRehire) 
        # Obtain Value with: $Global:chkRehire.Checked

    ## Reset Protocols
    $TopCol1 = $TopCol1 + 20
    $Global:chkProtocols = New-Object Windows.Forms.RadioButton 
        $Global:chkProtocols.Left = $LeftCol1; $Global:chkProtocols.Width = 200; $Global:chkProtocols.Top = $TopCol1
        $Global:chkProtocols.Text = "Reset Disabled Mailbox Protocols" 
        $Global:chkProtocols.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkProtocols) 
        # Obtain Value with: $Global:chkProtocols.Checked

    $TopCol1 = $TopCol1 + 40
    $Global:lblTitleLine2 = New-Object System.Windows.Forms.Label   
        $lblTitleLine2.Text = "M&A Activities:"
        $lblTitleLine2.Top = $TopCol1 ; $lblTitleLine2.Left = $TitleCol1; $lblTitleLine2.Width=120 ;$lblTitleLine2.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine2)    # Add to Form

    ## Add/Remove Forwarding
    $TopCol1 = $TopCol1 + 20
    $Global:chkFwdAddr = New-Object Windows.Forms.RadioButton 
        $Global:chkFwdAddr.Left = $LeftCol1; $Global:chkFwdAddr.Width = 300; $Global:chkFwdAddr.Top = $TopCol1
        $Global:chkFwdAddr.Text = "Add/Remove Mailbox Forwarding (M&A Accounts Only)" 
        $Global:chkFwdAddr.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkFwdAddr) 
        # Obtain Value with: $Global:chkFwdAddr.Checked

    #Build Column2
    $LeftCol2 = 420
    $TitleCol2 = 400
    $TopCol2 = 20

    $Global:lblTitleLine4 = New-Object System.Windows.Forms.Label   
        $lblTitleLine4.Text = "Change Activities:"
        $lblTitleLine4.Top = $TopCol2 ; $lblTitleLine4.Left = $TitleCol2; $lblTitleLine4.Width=120 ;$lblTitleLine4.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine4)    # Add to Form

    ## Add/Remove Access
    $TopCol2 = $TopCol2 + 20
    $Global:chkAccess = New-Object Windows.Forms.RadioButton 
        $Global:chkAccess.Left = $LeftCol2; $Global:chkAccess.Width = 450; $Global:chkAccess.Top = $TopCol2
        $Global:chkAccess.Text = "Add/Remove Access to a Mailbox/OneDrive" 
        $Global:chkAccess.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkAccess) 
        # Obtain Value with: $Global:chkAccess.Checked

    ## SMTP Address Change
    $TopCol2 = $TopCol2 + 20
    $Global:chkARPAddr = New-Object Windows.Forms.RadioButton 
        $Global:chkARPAddr.Left = $LeftCol2; $Global:chkARPAddr.Width = 450; $Global:chkARPAddr.Top = $TopCol2
        $Global:chkARPAddr.Text = "Add/Remove/Promote SMTP Address for User Account" 
        $Global:chkARPAddr.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkARPAddr) 
        # Obtain Value with: $Global:chkARPAddr.Checked

    ## Hide/UnHide User from Address Book
    $TopCol2 = $TopCol2 + 20
    $Global:chkHideUnHideAddr = New-Object Windows.Forms.RadioButton 
        $Global:chkHideUnHideAddr.Left = $LeftCol2; $Global:chkHideUnHideAddr.Width = 450; $Global:chkHideUnHideAddr.Top = $TopCol2
        $Global:chkHideUnHideAddr.Text = "Hide/UnHide User from Address Book" 
        $Global:chkHideUnHideAddr.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkHideUnHideAddr) 
        # Obtain Value with: $Global:chkHideUnHideAddr.Checked

    ## Manage Licenses
    $TopCol2 = $TopCol2 + 20
    $Global:chkManageLisc = New-Object Windows.Forms.RadioButton 
        $Global:chkManageLisc.Left = $LeftCol2; $Global:chkManageLisc.Width = 450; $Global:chkManageLisc.Top = $TopCol2
        $Global:chkManageLisc.Text = "Manage Licenses for a User" 
        $Global:chkManageLisc.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkManageLisc) 
        # Obtain Value with: $Global:chkManageLisc.Checked

    ## Move To/From ul.com and ul.org
    $TopCol2 = $TopCol2 + 20
    $Global:chkCOMORG = New-Object Windows.Forms.RadioButton 
        $Global:chkCOMORG.Left = $LeftCol2; $Global:chkCOMORG.Width = 300; $Global:chkCOMORG.Top = $TopCol2
        $Global:chkCOMORG.Text = "Move Staff from/to UL.COM/UL.ORG Domains" 
        $Global:chkCOMORG.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkCOMORG) 
        # Obtain Value with: $Global:chkCOMORG.Checked

    ## Rename AD Account
    $TopCol2 = $TopCol2 + 20
    $Global:chkRenADAct = New-Object Windows.Forms.RadioButton 
        $Global:chkRenADAct.Left = $LeftCol2; $Global:chkRenADAct.Width = 250; $Global:chkRenADAct.Top = $TopCol2  
        $Global:chkRenADAct.Text = "Rename a Active Directory Account" 
        $Global:chkRenADAct.Checked = $Global:chkRenADAct.Checked   # set a default value 
        $Global:form.Controls.Add($Global:chkRenADAct) 
        # Obtain Value with: $Global:chkRenADAct.Checked

    ## Review OOO Message
    $TopCol2 = $TopCol2 + 20
    $Global:chkOOOMsg = New-Object Windows.Forms.RadioButton 
        $Global:chkOOOMsg.Left = $LeftCol2; $Global:chkOOOMsg.Width = 450; $Global:chkOOOMsg.Top = $TopCol2
        $Global:chkOOOMsg.Text = "Review/Set Out of Office (OOO) Message" 
        $Global:chkOOOMsg.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkOOOMsg) 
        # Obtain Value with: $Global:chkOOOMsg.Checked

    $TopCol2 = $TopCol2 + 40
    $Global:lblTitleLine4 = New-Object System.Windows.Forms.Label   
        $lblTitleLine4.Text = "Miscellaneous Activities:"
        $lblTitleLine4.Top = $TopCol2 ; $lblTitleLine4.Left = $TitleCol2; $lblTitleLine4.Width=120 ;$lblTitleLine4.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine4)    # Add to Form

    ## Check User or Shared Mailbox Permission
    $TopCol2 = $TopCol2 + 20
    $Global:chkMbxPerm = New-Object Windows.Forms.RadioButton 
        $Global:chkMbxPerm.Left = $LeftCol2; $Global:chkMbxPerm.Width = 450; $Global:chkMbxPerm.Top = $TopCol2
        $Global:chkMbxPerm.Text = "Check User or Shared Mailbox Permissions" 
        $Global:chkMbxPerm.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkMbxPerm) 
        # Obtain Value with: $Global:chkMbxPerm.Checked

    ## Enable UM
    $TopCol2 = $TopCol2 + 20
    $Global:chkEnabUM = New-Object Windows.Forms.RadioButton 
        $Global:chkEnabUM.Left = $LeftCol2; $Global:chkEnabUM.Width = 450; $Global:chkEnabUM.Top = $TopCol2
        $Global:chkEnabUM.Text = "Enable Unified Messaging (UM) for a User" 
        $Global:chkEnabUM.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkEnabUM) 
        # Obtain Value with: $Global:chkEnabUM.Checked

    ## List Addresses
    $TopCol2 = $TopCol2 + 20
    $Global:chkListAddr = New-Object Windows.Forms.RadioButton 
        $Global:chkListAddr.Left = $LeftCol2; $Global:chkListAddr.Width = 450; $Global:chkListAddr.Top = $TopCol2
        $Global:chkListAddr.Text = "List All SMTP Addresses for a Mailbox" 
        $Global:chkListAddr.Checked = $Global:chkListAddr.Checked   # set a default value 
        $Global:form.Controls.Add($Global:chkListAddr) 
        # Obtain Value with: $Global:chkListAddr.Checked

    ## Distribution List/Security Group Information
    $TopCol2 = $TopCol2 + 20
    $Global:chkDLSecGrpInf = New-Object Windows.Forms.RadioButton 
        $Global:chkDLSecGrpInf.Left = $LeftCol2; $Global:chkDLSecGrpInf.Width = 450; $Global:chkDLSecGrpInf.Top = $TopCol2
        $Global:chkDLSecGrpInf.Text = "Distribution List/Security Group Information" 
        $Global:chkDLSecGrpInf.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkDLSecGrpInf) 
        # Obtain Value with: $Global:chkDLSecGrpInf.Checked

    ## Room or Resource Inforamtion
    $TopCol2 = $TopCol2 + 20
    $Global:chkRRInf = New-Object Windows.Forms.RadioButton 
        $Global:chkRRInf.Left = $LeftCol2; $Global:chkRRInf.Width = 450; $Global:chkRRInf.Top = $TopCol2
        $Global:chkRRInf.Text = "Room or Resource information"
        $Global:chkRRInf.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkRRInf) 
        # Obtain Value with: $Global:chkRRInf.Checked

    ## Shared Mailbox Information
    $TopCol2 = $TopCol2 + 20
    $Global:chkShrMbxInf = New-Object Windows.Forms.RadioButton 
        $Global:chkShrMbxInf.Left = $LeftCol2; $Global:chkShrMbxInf.Width = 450; $Global:chkShrMbxInf.Top = $TopCol2
        $Global:chkShrMbxInf.Text = "Shared Mailbox Information" 
        $Global:chkShrMbxInf.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkShrMbxInf) 
        # Obtain Value with: $Gl

    ## User Mailbox Information
    $TopCol2 = $TopCol2 + 20
    $Global:chkUsrMbxInf = New-Object Windows.Forms.RadioButton 
        $Global:chkUsrMbxInf.Left = $LeftCol2; $Global:chkUsrMbxInf.Width = 450; $Global:chkUsrMbxInf.Top = $TopCol2
        $Global:chkUsrMbxInf.Text = "User Mailbox Information" 
        $Global:chkUsrMbxInf.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkUsrMbxInf) 
        # Obtain Value with: $Global:chkUsrMbxInf.Checked

   ## Create New Credential File

    If ($TopCol1 -lt $TopCol2)
    {
        $TopLoc = $TopCol2 + 60
    }
    else
    {
        $TopLoc = $TopCol1 + 60
    }
    $Global:chkNewCred = New-Object Windows.Forms.RadioButton 
        $Global:chkNewCred.Left = 230; $Global:chkNewCred.Width = 450; $Global:chkNewCred.Top = $TopLoc
        $Global:chkNewCred.Text = "Create New Credential File" 
        $Global:chkNewCred.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkNewCred) 
        # Obtain Value with: $Global:chkNewCred.Checked
}

Function Publish-MMForm
{
    ## Finalize Form and Show Dialog
    $Global:form.Add_Shown( { $form.Activate(); $Global:okButton.Focus() } )  #Activate and Set Focus s 
    $Global:MainResult = $Global:form.ShowDialog()          ## Show the form, and wait for the response
}

#### Start of Script

#$Global:LiveCred = ""
#$Global:CredFile = ""
$File = "e:\SDAP\scripts\SDAPAdminMenu.ps1"

if (test-path $file)
{
#   Do nothing E" drive exists
}
else
{
    New-PSDrive -Name E -PSProvider FileSystem -Root \\USNBKM100P\e$
}

write-host ""
$reconnectThreshold = New-TimeSpan -Minutes 60
invoke-expression -Command E:\O365AdminShared\Scripts\SDAPConnectO365.ps1
$stopwatch = [diagnostics.stopwatch]::StartNew()
invoke-expression -Command E:\O365AdminShared\Scripts\O365DisabledLicenseFeatures.ps1

<#$Global:sysvars = get-variable | select -ExpandProperty name
$Global:sysvars += 'sysvar'
write-host $Global:sysvars.count
#>

Do
{
<#
    #Store just system variables
    $sysvars = get-variable | select -ExpandProperty name
    $sysvars += 'sysvar'
    write-host $sysvars.count
#>

    #start Script
    $Global:FormRefresh = "N"
    Build-AdminMenuForm
    Publish-MMForm

    If ($Global:MainResult -eq "OK")
    {
        if ($stopwatch.elapsed -ge $reconnectThreshold)
        {
            # Close all sessions
            write-host "Connection Threshold Exceeded -- Reconnecting to Office365" -ForegroundColor Red
            get-pssession | Remove-PSSession -Confirm:$false
            invoke-expression -Command E:\O365AdminShared\Scripts\SDAPConnectO365.ps1
            $stopwatch = [diagnostics.stopwatch]::StartNew()
        }

        If ($Global:chkEnabUsr.Checked -eq "Checked") 
        {

            invoke-expression -Command .\EnableUserConfigMbx.ps1
        }

        If ($Global:chkGenAcct.Checked -eq "Checked") 
        {

            invoke-expression -Command E:\O365AdminShared\Scripts\NewGenericAccount.ps1
        }

        If ($Global:chkUpdSvcAcct.Checked -eq "Checked")
        {
            invoke-expression -Command E:\O365AdminShared\Scripts\UpdateGenericAccount.ps1
        }

        If ($Global:chkTermination.Checked -eq "Checked")
	    {
            invoke-expression -command .\SDAPTermination.ps1
	    }

        If ($Global:chkPurgeAct.Checked -eq "Checked")
        {
            Invoke-Expression -Command .\SDAPPurgeAccount.ps1
            #Invoke-Expression -Command .\PurgeAccount.ps1
        }

        If ($Global:chkRehire.Checked -eq "Checked")
        {
            Invoke-Expression -Command .\EmergencyRehire.ps1
        }

        If ($Global:chkProtocols.Checked -eq "Checked")
        {
            write-host "Enter the Employee Number of the Mailbox to Reveiw and Reset Disabled Protocols for: " -ForegroundColor Cyan -NoNewLine
            $UsrMbx = Read-Host

            $Filename       = "MbxProtocolReset"
            $LogFile		= "E:\SDAP\MbxProtocolReset\Log\Log-" + $FileName + ".log"
            $ReportFile		= "E:\SDAP\MbxProtocolReset\Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            $LineToWrite = "STAR" + "`t" + $FileName + " script has started"
            WriteLogEvent
            $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
            WriteLogEvent
            WriteReportEvent

            $UsrExists = [bool](get-mailbox $UsrMbx -ErrorAction SilentlyContinue)
            If ($UsrExists -eq "True")
            {
                $Mbx = get-mailbox $UsrMbx
                $CASMbx = Get-CASMailbox -Identity $UsrMbx
			    write-host "            Checking Protocols for:  " $UsrMbx "(" $CASMbx.DisplayName ")"
                write-host "                  Mailbox Disabled:  " $Mbx.AccountDisabled
                Write-Host "  OWA (Outlook Web Access) Enabled:  " $CASMbx.OwaEnabled
                Write-Host "             OWAforDevices Enabled:  " $CASMbx.OWAforDevicesEnabled
                write-host "           Outlook for MAC Enabled:  " $CASMbx.EwsAllowMacOutlook
                $LineToWrite = $WhoAmI + "`t" + "            Checking Protocols for:  " + $UsrMbx + "(" + $CASMbx.DisplayName + ")"
                WriteReportEvent
                $LineToWrite = $WhoAmI + "`t" + "                  Mailbox Disabled:  " + $Mbx.AccountDisabled
                WriteReportEvent
                $LineToWrite = $WhoAmI + "`t" + "  OWA (Outlook Web Access) Enabled:  " + $CASMbx.OwaEnabled
                WriteReportEvent
                $LineToWrite = $WhoAmI + "`t" + "             OWAforDevices Enabled:  " + $CASMbx.OWAforDevicesEnabled
                WriteReportEvent
                $LineToWrite = $WhoAmI + "`t" + "           Outlook for MAC Enabled:  " + $CASMbx.EwsAllowMacOutlook
                WriteReportEvent
            
                write-host "Do you need to reset the Protocols (Y/N)?" -ForegroundColor Cyan -NoNewline
                $RestProt = Read-Host
            
                If ($RestProt -eq "Y")
                {
                    Set-Mailbox $UsrMbx -AccountDisabled:$False
                    Set-CASMailbox -Identity $UsrMbx -OwaEnabled $True -OWAforDevicesEnabled $True -ActiveSyncEnabled $True -EwsAllowMacOutlook $True
                    $Mbx = get-mailbox $UsrMbx
                    $CASMbx = Get-CASMailbox -Identity $UsrMbx
                    write-host "`nNew Protocol Settings"
                    write-host "                  Mailbox Disabled:  " $Mbx.AccountDisabled
                    Write-Host "  OWA (Outlook Web Access) Enabled:  " $CASMbx.OwaEnabled
                    Write-Host "             OWAforDevices Enabled:  " $CASMbx.OWAforDevicesEnabled
                    write-host "Outlook for MAC Enabled (Default: blank:  " $CASMbx.EwsAllowMacOutlook
                    $LineToWrite = $WhoAmI + "`t" + ""
                    WriteReportEvent
                    $LineToWrite = $WhoAmI + "`t" + "New Protocol Values"
                    WriteReportEvent
                    $LineToWrite = $WhoAmI + "`t" + "                  Mailbox Disabled:  " + $Mbx.AccountDisabled
                    WriteReportEvent
                    $LineToWrite = $WhoAmI + "`t" + "  OWA (Outlook Web Access) Enabled:  " + $CASMbx.OwaEnabled
                    WriteReportEvent
                    $LineToWrite = $WhoAmI + "`t" + "             OWAforDevices Enabled:  " + $CASMbx.OWAforDevicesEnabled
                    WriteReportEvent
                    $LineToWrite = $WhoAmI + "`t" + "           Outlook for MAC Enabled:  " + $CASMbx.EwsAllowMacOutlook
                    WriteReportEvent
                }
		        else
			    {
                    $LineToWrite = $WhoAmI + "`t" + "Protocols are not being reset"
                    WriteReportEvent
			    }
            }
            else
            {
                write-host "User Mailbox" $EmpNo "Does Not Exist"
                $LineToWrite = $WhoAmI + "`t" + "       User Mailbox Does Not Exist" + $UsrMbx
                WriteReportEvent
            }
        }

        If ($Global:chkFwdAddr.Checked -eq "Checked")
        {
    #		Add/Remove Forwarding
            write-host "Enter 'A' to Add a Forwarding Address and 'R' to Remove a Forwarding Address (A/R): " -ForegroundColor Green -NoNewline
            $FAddr = Read-Host
            switch ($FAddr)
            {
                "A"
                {
                    write-host "`nThis requires that an input file be created in the \SetForwarding\Input directory." -ForegroundColor Red
    	            write-host "Hit return when the input file is ready" -foregroundcolor Red -nonewline
		            $ready = read-host
		            invoke-expression -Command .\SetForwarding.ps1
                }
                "R"
                {
                    Write-Host "Enter Employee Number of individuals to remove the forwarding address from:" -ForegroundColor Green -NoNewline
                    $EmpNo = Read-Host
                    $mbx = Get-Mailbox $EmpNo -ErrorAction SilentlyContinue
                    If ($mbx.ForwardingAddress -ne "")
                    {
                        write-host "Forwarding Address: " $mbx.ForwardingAddress
                        write-host "SMTP Forwarding Address: " $mbx.ForwardingSMTPAddress
                        write-host "Removing Forwarding Addresses..." -ForegroundColor Green
                        set-mailbox $EmpNo -ForwardingAddress $null -ForwardingSMTPAddress $null
                        $mbx = get-mailbox $EmpNo -ErrorAction SilentlyContinue
                        write-host "Forwarding Address: " $mbx.ForwardingAddress
                        write-host "SMTP Forwarding Address: " $mbx.ForwardingSMTPAddress
                    }
                }
            }
        }

        If ($Global:chkAccess.Checked -eq "Checked") 
        {
            write-host "Enter 'A' to Grant Access, 'R' to Remove a Access or 'V' to View Access to a Mailbox (A/R/V): " -ForegroundColor Green -NoNewline
            $FAddr = Read-Host

            write-host "Enter Employee Number of Mailbox to Give\Remove Access To: " -ForegroundColor Green -NoNewline
            $ENo = Read-Host

            If (get-mailbox $ENo -ErrorAction SilentlyContinue)
            {
                switch ($FAddr)
                {
                    "A"
                    {
                        # Declare Drive | Folders | and Files
                        $FileName		= "FolderPermissions"
	                    $LogDrive		= "E:"
	                    $LogPath		= "\SDAP"
	                    $LogFolder		= "\" + $FileNAme
	                    $LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
	                    $LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	                    $InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
	                    $ReportFile		= $LogDirectory + "Report\ReportAdd-" + $FileName + "-EmpNo" + $ENo + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
                        write-host "Log Detailes can be found in: " $ReportFile -ForegroundColor Cyan
                        invoke-expression -Command .\GetADUserInfo.ps1
                        invoke-expression -Command .\SDAPFldrAccess.ps1                    
                    }
                    "R"
                    {
                        # Declare Drive | Folders | and Files
	                    $FileName		= "RemoveMailboxPermission"
	                    $LogDrive		= "E:"
	                    $LogPath		= "\SDAP"
	                    $LogFolder		= "\" + $FileNAme
	                    $LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
	                    $LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	                    $InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
	                    $ReportFile		= $LogDirectory + "Report\ReportAdd-" + $FileName + "-EmpNo" + $ENo + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
                        write-host "Log Detailes can be found in: " $ReportFile -ForegroundColor Cyan
                        invoke-expression -Command .\RemoveMBXAccess.ps1
                    }
                    "V"
                    {
                        invoke-expression -Command .\ViewFolderPermission.ps1
                    }
                }
            }
            else
            {
                 write-host "`n     No mailbox exists for: " $ENo -ForegroundColor Red
            }
        }

        If ($Global:chkARPAddr.Checked -eq "Checked")
        {
            invoke-Expression -Command .\SMTPAddressChange.ps1
        }

        If ($Global:chkHideUnHideAddr.Checked -eq "Checked")
        {
            # Declare Drive | Folders | and Files
		    $FileName		= "HideUnhide"
		    $LogDrive		= "E:"
		    $LogPath		= "\SDAP"
		    $LogFolder		= "\" + $FileNAme
		    $LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
		    $LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
		    $InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
					
		    # Retrieve the user name
		    $WhoAmI			= WhoAmI
						
		    # Retrieve the local server name
		    $Machine = get-wmiobject "Win32_ComputerSystem"
		    $LocalMachineName = $Machine.Name			
		
		    write-host "Enter Employee Number: " -foregroundcolor Cyan -nonewline
		    $ENo = read-host
            $ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-EmpNo" + $ENo + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
		    $Hide = ""
		    invoke-Expression -command .\GetADUserInfo.ps1
    	    if ($Global:u.msExchHideFromAddressLists.value -eq $true)
		    {
			    write-Host "Account is Already Hidden from the Address Book do you want to Unhide it(Y/N)? " -foregroundcolor Cyan -nonewline
			    $Hide = read-host
			    If ($Hide -eq "Y")
			    {
				    $LineToWrite = $WhoAmI + "`t" + "`nUnhiding Account " + $ENo + " from Address Book"
				    WriteReportEvent			
				    $Global:u.msExchHideFromAddressLists.value = $FALSE
                    $Global:u.CommitChanges()		
			    }
		    }
		    else
		    {
			    write-Host "Account is not Hidden from the Address Book do you want to Hide it(Y/N)? " -foregroundcolor Cyan -nonewline
			    $UnHide = read-host
			    If ($UnHide -eq "Y")
			    {
				    $LineToWrite = $WhoAmI + "`t" + "`nHiding Account " + $ENo + " from Address Book"
				    WriteReportEvent						
				    $Global:u.msExchHideFromAddressLists.value = $TRUE
                    $Global:u.CommitChanges()		
			    }		
		    }
        }

        If ($Global:chkManageLisc.Checked -eq "Checked")
        {
            invoke-expression -Command .\SDAPLicensingMenu.ps1
        }

        If ($Global:chkCOMORG.Checked -eq "Checked")
        {
            invoke-Expression -command .\EMailDomainChange.ps1
        }

        If ($Global:chkRenADAct.Checked -eq "Checked") 
        {
            invoke-expression -command .\RenameADAccount.ps1
        }

        If ($Global:chkOOOMsg.Checked -eq "Checked")
        {
            invoke-expression -command .\UserOOOReview.ps1
        }

        If ($Global:chkMbxPerm.Checked -eq "Checked") 
        {
            invoke-expression .\SDAPMailboxAccessInfo.ps1
        }

        If ($Global:chkEnabUM.Checked -eq "Checked")
        {
    #       This section will be used for assigning UM
            $SfBEnable = "N"
            $SfBEnable = read-host "Do you need to enable SkypeforBusiness or Unified Messaging for this individual (Y/N)?"
            If ($SfBEnable -eq "Y")
            {
                $url = "https://sfbhelper.ul.com/newhireSD.php"
                $ie = New-Object -com internetexplorer.application; 
                $ie.visible = $true; 
                $ie.navigate($url);

                write-host "Using the IE session that was opened configure complete the configuration" -ForegroundColor Red
                pause
            }
        }

        If ($Global:chkListAddr.Checked -eq "Checked")
        {
            invoke-Expression .\SDAPMailboxAddressInfo.ps1
        }

        If ($Global:chkDLSecGrpInf.Checked -eq "Checked")
        {
            invoke-expression .\SDAPDistributionGroupInfo.ps1
        }

        If ($Global:chkRRInf.Checked -eq "Checked")
        {
            invoke-expression .\SDAPRoomResourceMailboxInfo.ps1
        }

        If ($Global:chkShrMbxInf.Checked -eq "Checked")
        {
            invoke-expression .\SDAPSharedMailboxInfo.ps1
        }

        If ($Global:chkUsrMbxInf.Checked -eq "Checked")
        {
		    invoke-expression .\SDAPUserMailboxInfo.ps1
        }

        If ($Global:chkNewCred.Checked -eq "Checked")
        {
            Invoke-Expression -Command .\CreateNewCredFile.ps1
        }
    }
    
<#    #clear all non-system variables stored at the start of this script

    write-host "Current variables= " (get-variable *).count
    get-variable * | where { $_.name -like "*txt*" } | Remove-Variable
    (get-variable *).count
    get-variable * | where { $_.name -like "*lbl*" } | Remove-Variable
    (get-variable *).count
    $Global:sysvars = get-variable | select -ExpandProperty name
    $Global:sysvars += 'sysvar'
    write-host $Global:sysvars.count
#>
}While ($Global:MainResult -eq "OK")	