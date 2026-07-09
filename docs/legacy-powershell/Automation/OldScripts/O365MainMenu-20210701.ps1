#
#   O365 Team Main Menu
#
#
#  06-16-2021 - SAG - Added item to search the Mailbox Audit Log
#

Function Build-AdminMenuForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "O365 Admin Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 540 ; $form.Height = 600  # Make the form wider 
    Add-FormStandardButtons

    $TopLoc = 20 

    ## Label and TextBox  
    ## Title Line
    $Global:lblTitleLine = New-Object System.Windows.Forms.Label   
        $Global:lblTitleLine.Text = "Select Option:"
        $Global:lblTitleLine.Top = 15 ; $Global:lblTitleLine.Left = 60; $Global:lblTitleLine.Width=120 ;$Global:lblTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine)    # Add to Form 

    ## Add New Domain
    $TopLoc = $TopLoc + 20
    $Script:chkAddDom = New-Object Windows.Forms.RadioButton 
        $Script:chkAddDom.Left = 100; $Script:chkAddDom.Width = 450; $Script:chkAddDom.Top = $TopLoc  
        $Script:chkAddDom.Text = "Add New Domain to the Tenant" 
        $Script:chkAddDom.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chkAddDom) 
        # Obtain Value with: $Script:chkAddDom.Checked

    ## Add Enable DKIM keys for New Domain
    $TopLoc = $TopLoc + 20
    $Global:chkDKIMKey = New-Object Windows.Forms.RadioButton 
        $Global:chkDKIMKey.Left = 100; $Global:chkDKIMKey.Width = 450; $Global:chkDKIMKey.Top = $TopLoc  
        $Global:chkDKIMKey.Text = "Add/Enable DKIM Keys for New Domain" 
        $Global:chkDKIMKey.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkDKIMKey) 
        # Obtain Value with: $Global:chkDKIMKey.Checked
        
    ## Configure Service Accounts/Surface Hub Devices
    $TopLoc = $TopLoc + 20
    $Global:chkNewSvcAct = New-Object Windows.Forms.RadioButton 
        $Global:chkNewSvcAct.Left = 100; $Global:chkNewSvcAct.Width = 450; $Global:chkNewSvcAct.Top = $TopLoc  
        $Global:chkNewSvcAct.Text = "Configure Service Accounts/Surface Hub Devices" 
        $Global:chkNewSvcAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkNewSvcAct) 
        # Obtain Value with: $Global:chkNewSvcAct.Checked

<#    ## Discovery Search Mailbox
    $TopLoc = $TopLoc + 20
    $Global:chkDiscvMbx = New-Object Windows.Forms.RadioButton 
        $Global:chkDiscvMbx.Left = 100; $Global:chkDiscvMbx.Width = 450; $Global:chkDiscvMbx.Top = $TopLoc  
        $Global:chkDiscvMbx.Text = "Create Discovery Search Mailbox" 
        $Global:chkDiscvMbx.Checked = $false   # set a default value 
        $Global:chkDiscvMbx.TabIndex = 4
        $Global:form.Controls.Add($Global:chkDiscvMbx) 
        # Obtain Value with: $Global:chkDiscvMbx.Checked
#>

    ## Create New Service Account
    $TopLoc = $TopLoc + 20
    $Global:chkGenAcct = New-Object Windows.Forms.RadioButton 
        $Global:chkGenAcct.Left = 100; $Global:chkGenAcct.Width = 450; $Global:chkGenAcct.Top = $TopLoc  
        $Global:chkGenAcct.Text = "Create New Generic Account" 
        $Global:chkGenAcct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkGenAcct) 
        # Obtain Value with: $Global:chkGenAcct.Checked
        
    ## DL/O365 Group/Security Group 
    $TopLoc = $TopLoc + 20
    $Global:chkDLUniSecGrp = New-Object Windows.Forms.RadioButton 
        $Global:chkDLUniSecGrp.Left = 100; $Global:chkDLUniSecGrp.Width = 450; $Global:chkDLUniSecGrp.Top = $TopLoc  
        $Global:chkDLUniSecGrp.Text = "Distribution List / O365 Group / Security Group Menu" 
        $Global:chkDLUniSecGrp.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkDLUniSecGrp) 
        # Obtain Value with: $Global:chkDLUniSecGrp.Checked

    ## Enable/Disable Compromised Account
    $TopLoc = $TopLoc + 20
    $Global:chkCompAct = New-Object Windows.Forms.RadioButton 
        $Global:chkCompAct.Left = 100; $Global:chkCompAct.Width = 450; $Global:chkCompAct.Top = $TopLoc  
        $Global:chkCompAct.Text = "Enable/Disable Compromised Account" 
        $Global:chkCompAct.Checked = $Global:chkCompAct.Checked   # set a default value 
        $Global:form.Controls.Add($Global:chkCompAct) 
        # Obtain Value with: $Global:chkCompAct.Checked

    ## OneDrive Space Usage
    $TopLoc = $TopLoc + 20
    $Global:chkOneDrv = New-Object Windows.Forms.RadioButton 
        $Global:chkOneDrv.Left = 100; $Global:chkOneDrv.Width = 450; $Global:chkOneDrv.Top = $TopLoc  
        $Global:chkOneDrv.Text = "Get Users OneDrive Space Usage" 
        $Global:chkOneDrv.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkOneDrv) 
        # Obtain Value with: $Global:chkOneDrv.Checked

    ## Locked PC Information
    $TopLoc = $TopLoc + 20
    $Global:chkLockPCInf = New-Object Windows.Forms.RadioButton 
        $Global:chkLockPCInf.Left = 100; $Global:chkLockPCInf.Width = 450; $Global:chkLockPCInf.Top = $TopLoc  
        $Global:chkLockPCInf.Text = "Get PC Locked Information" 
        $Global:chkLockPCInf.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkLockPCInf) 
        # Obtain Value with: $Global:chkLockPCInf.Checked

    ## Legal Hold 
    $TopLoc = $TopLoc + 20
    $Global:chkLglHold = New-Object Windows.Forms.RadioButton 
        $Global:chkLglHold.Left = 100; $Global:chkLglHold.Width = 450; $Global:chkLglHold.Top = $TopLoc
        $Global:chkLglHold.Text = "Legal Hold Activities" 
        $Global:chkLglHold.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkLglHold) 
        # Obtain Value with: $Global:chkLglHold.Checked

    ## Licensing Menu
    $TopLoc = $TopLoc + 20
    $Global:chkLicMenu = New-Object Windows.Forms.RadioButton 
        $Global:chkLicMenu.Left = 100; $Global:chkLicMenu.Width = 450; $Global:chkLicMenu.Top = $TopLoc  
        $Global:chkLicMenu.Text = "Licensing Menu" 
        $Global:chkLicMenu.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkLicMenu) 
        # Obtain Value with: $Global:chkLicMenu.Checked
 
    ## M&A Activities
    $TopLoc = $TopLoc + 20
    $Global:chkMAActv = New-Object Windows.Forms.RadioButton 
        $Global:chkMAActv.Left = 100; $Global:chkMAActv.Width = 450; $Global:chkMAActv.Top = $TopLoc  
        $Global:chkMAActv.Text = "M&A Activities Menu" 
        $Global:chkMAActv.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkMAActv) 
        # Obtain Value with: $Global:chkMAActv.Checked

    ## Reporting
    $TopLoc = $TopLoc + 20
    $Global:chkRptMenu = New-Object Windows.Forms.RadioButton 
        $Global:chkRptMenu.Left = 100; $Global:chkRptMenu.Width = 450; $Global:chkRptMenu.Top = $TopLoc  
        $Global:chkRptMenu.Text = "Reporting Menu" 
        $Global:chkRptMenu.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkRptMenu) 
        # Obtain Value with: $Global:chkRptMenu.Checked

    ## Shared Room/Resource
    $TopLoc = $TopLoc + 20
    $Global:chkRoomRes = New-Object Windows.Forms.RadioButton 
        $Global:chkRoomRes.Left = 100; $Global:chkRoomRes.Width = 450; $Global:chkRoomRes.Top = $TopLoc  
        $Global:chkRoomRes.Text = "Room or Resource Menu" 
        $Global:chkRoomRes.Checked = $Global:chkRoomRes.Checked   # set a default value 
        $Global:form.Controls.Add($Global:chkRoomRes) 
        # Obtain Value with: $Global:chkRoomRes.Checked

    ## Search Mailbox Audit Log
    $TopLoc = $TopLoc + 20
    $Global:chkSrchMbxALog = New-Object Windows.Forms.RadioButton 
        $Global:chkSrchMbxALog.Left = 100; $Global:chkSrchMbxALog.Width = 450; $Global:chkSrchMbxALog.Top = $TopLoc  
        $Global:chkSrchMbxALog.Text = "Search Mailbox Audit Log" 
        $Global:chkSrchMbxALog.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkSrchMbxALog) 
        # Obtain Value with: $Global:chkSrchMbxALog.Checked

    ## Shared Mailbox
    $TopLoc = $TopLoc + 20
    $Global:chkShrMbx = New-Object Windows.Forms.RadioButton 
        $Global:chkShrMbx.Left = 100; $Global:chkShrMbx.Width = 450; $Global:chkShrMbx.Top = $TopLoc  
        $Global:chkShrMbx.Text = "Shared Mailbox Menu" 
        $Global:chkShrMbx.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkShrMbx) 
        # Obtain Value with: $Global:chkShrMbx.Checked

    ## Unified Messaging
    $TopLoc = $TopLoc + 20
    $Global:chkUniMsg = New-Object Windows.Forms.RadioButton 
        $Global:chkUniMsg.Left = 100; $Global:chkUniMsg.Width = 450; $Global:chkUniMsg.Top = $TopLoc  
        $Global:chkUniMsg.Text = "Unified Messaging Menu" 
        $Global:chkUniMsg.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkUniMsg) 
        # Obtain Value with: $Global:chkUniMsg.Checked

    ## Update Service Account
    $TopLoc = $TopLoc + 20
    $Global:chkUpdSvcAcct = New-Object Windows.Forms.RadioButton
        $Global:chkUpdSvcAcct.Left = 100; $Global:chkUpdSvcAcct.Width = 450; $Global:chkUpdSvcAcct.Top = $TopLoc  
        $Global:chkUpdSvcAcct.Text = "Update Generic Account Details" 
        $Global:chkUPdSvcAcct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkUpdSvcAcct) 

    ## User Mailbox
    $TopLoc = $TopLoc + 20
    $Global:chkUsrMbx = New-Object Windows.Forms.RadioButton 
        $Global:chkUsrMbx.Left = 100; $Global:chkUsrMbx.Width = 450; $Global:chkUsrMbx.Top = $TopLoc  
        $Global:chkUsrMbx.Text = "User Mailbox Menu" 
        $Global:chkUsrMbx.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkUsrMbx) 
        # Obtain Value with: $Global:chkUsrMbx.Checked

    ## Create New Credential File
    $TopLoc = $TopLoc + 40
    $Global:chkNewCred = New-Object Windows.Forms.RadioButton 
        $Global:chkNewCred.Left = 100; $Global:chkNewCred.Width = 450; $Global:chkNewCred.Top = $TopLoc  
        $Global:chkNewCred.Text = "Create New Credential File" 
        $Global:chkNewCred.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkNewCred) 
        # Obtain Value with: $Global:chkNewCred.Checked

    ## Misc Powershell Commands
    $TopLoc = $TopLoc + 20
    $Global:chkMiscCmd = New-Object Windows.Forms.RadioButton 
        $Global:chkMiscCmd.Left = 100; $Global:chkMiscCmd.Width = 450; $Global:chkMiscCmd.Top = $TopLoc  
        $Global:chkMiscCmd.Text = "Open Miscellaneous Powershell Commands Document" 
        $Global:chkMiscCmd.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkMiscCmd) 
        # Obtain Value with: $Global:chkMiscCmd.Checked
}

Function Publish-MMForm
{
    ## Finalize Form and Show Dialog
    $Global:form.Add_Shown( { $form.Activate(); $Global:okButton.Focus() } )  #Activate and Set Focus s 
    $Global:MainResult = $Global:form.ShowDialog()          ## Show the form, and wait for the response
}

Function Check-Reconnect
{
#    write-host "Elapsed Time: " $Global:stopwatch.Elapsed
    if ($Global:stopwatch.elapsed -ge $Global:reconnectThreshold)
    {
        # Close all sessions
        write-host "Connection Threshold Exceeded -- Reconnecting to Office365" -ForegroundColor Red
        get-pssession | Remove-PSSession -Confirm:$false
        invoke-expression -Command E:\O365AdminShared\Scripts\ConnectO365.ps1
        $Global:stopwatch = [diagnostics.stopwatch]::StartNew()
    }
}


#### Start of Script

$Global:LiveCred = ""
$Global:CredFile = ""
$File = "e:\Automation\scripts\O365AdminMenu.ps1"

if (test-path $file)
{
#   Do nothing E" drive exists
}
else
{
    New-PSDrive -Name E -PSProvider FileSystem -Root \\USNBKM100P\e$
}

Set-Location E:\Automation\Scripts
write-host ""

$Global:reconnectThreshold = New-TimeSpan -Minutes 60
invoke-expression -Command E:\O365AdminShared\Scripts\ConnectO365.ps1

Do
{
    $Global:OKDetails = "Continue"
    $Global:stopwatch = [diagnostics.stopwatch]::StartNew()
    Build-AdminMenuForm
    Publish-MMForm

    Check-Reconnect
        		
    If ($Global:chkLglHold.Checked -eq "Checked")
    {
#        write-host "Elapsed Time: " $Global:stopwatch.Elapsed
        Invoke-Expression -Command .\LegalHoldCheck.ps1
    }

    If ($Global:chkUniMsg.Checked -eq "Checked")
    {
        invoke-expression -Command .\EUMMenu.ps1
    }

    If ($Global:chkUsrMbx.Checked -eq "Checked")
    {
        invoke-expression -Command .\UserMailboxMenu.ps1
    }

    If ($Global:chkDLUniSecGrp.Checked -eq "Checked")
    {
        invoke-expression -Command .\DistributionSecurityGroupMenu.ps1
    }

    If ($Global:chkSrchMbxALog.Checked -eq "Checked")
    {
        invoke-expression -Command .\AuditLogSearch.ps1
    }

    If ($Global:chkShrMbx.Checked -eq "Checked")
    {
        invoke-expression -Command .\ShrMbxAdminMenu.ps1
    }
    
    If ($Global:chkRoomRes.Checked -eq "Checked")
    {
        invoke-expression -Command .\RoomResourceAdminMenu.ps1
    }

<#    If ($Global:chkDiscvMbx.Checked -eq "Checked")
    {
        Write-Host "Enter Name of Discovery Mailbox:  " -ForegroundColor Yellow -NoNewline
        $DiscMbx = Read-Host
        $DiscMbxAddr = ($DiscMbx -replace '\s','') + "@ul.onmicrosoft.com"
                    
        New-mailbox $DiscMbx –Discovery –PrimarySMTPAddress $DiscMbxAddr 
        Add-MailboxPermission $DiscMbx –AccessRights FullAccess –User “Discovery Management{8f1d3a49ccc740dcb9d8c3cd06b95011}”

        Write-Host "Discovery Mailbox Created:  " $DiscMbx
        pause
    }
#>

    If ($Global:chkGenAcct.Checked -eq "Checked")
    {
        invoke-expression -Command E:\O365AdminShared\Scripts\NewGenericAccount.ps1
    }

    If ($Global:chkUpdSvcAcct.Checked -eq "Checked")
    {
        invoke-expression -Command E:\O365AdminShared\Scripts\UpdateGenericAccount.ps1
    }

    If ($Global:chkLockPCInf.Checked -eq "Checked")
    {
        Invoke-Expression -Command .\Get-UserLockedPCName.ps1
    }

    If ($Global:chkMAActv.Checked -eq "Checked")
    {
        Invoke-Expression -Command .\MAActivitesMenu.ps1
    }
    
    If ($Global:chkLicMenu.Checked -eq "Checked")                  
    {
        Invoke-Expression -Command .\LicenseMenu.ps1
    }

    If ($Global:chkRptMenu.Checked -eq "Checked") 
    {
        Invoke-Expression -Command .\ReportingMenu.ps1
    }

    If ($Script:chkAddDom.Checked -eq "Checked") 
    {
        $Script:form = New-Object Windows.Forms.Form 
        $Script:form.FormBorderStyle = "FixedToolWindow" 
        $Script:form.Text = "Add New Domain Tenant" 
        $Script:form.StartPosition = "CenterScreen"
        $Script:form.Size = New-Object System.Drawing.Size(400,200) #(W,H)
        ## Domain to Add
        $Script:lblAddDomain = New-Object System.Windows.Forms.Label   
        $Script:lblAddDomain.Text = "Domain to Add:"
        $Script:lblAddDomain.Top = 40 ; $Script:lblAddDomain.Left = 10; $Script:lblAddDomain.Width=120 ;$Script:lblAddDomain.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblAddDomain)    # Add to Form 
        $Script:txtAddDomain = New-Object Windows.Forms.TextBox  
        $Script:txtAddDomain.TabIndex = 0 # set Tab Order 
        $Script:txtAddDomain.Top = 40; $Script:txtAddDomain.Left = 130; $Script:txtAddDomain.Width = 200;  
        $Script:txtAddDomain.Text = ""
        $Script:form.Controls.Add($Script:txtAddDomain)    # Add to Form
        $Global:InputFocus = $Script:txtAddDomain 
        Add-FormStandardButtons
        Publish-Form

        If ($Global:Result -eq "OK")
        {
            $NewDom = $Script:txtAddDomain.Text.TrimStart("@")
            New-MsolDomain -Name $NewDom
            If (($NewDom -notlike "*ul.com*") -and ($NewDom -notlike "*ul.org*"))
            {
                $Output = $wshell.Popup("Setting the Domain to be an InternalRelay.",0,"Internal Relay",0+32)
                Set-AcceptedDomain $NewDom -DomainType InternalRelay
            }
        }
        else
        {
            $Output = $wshell.Popup("Adding of new domain to the tenant cancelled.",0,"Cancelled",0+32)
        }
    }

    If ($Global:chkDKIMKey.Checked -eq "Checked") 
    {
        $Global:form = New-Object Windows.Forms.Form 
        $Global:form.FormBorderStyle = "FixedToolWindow" 
        $Global:form.Text = "Add DKIM Key to Domain" 
        $Global:form.StartPosition = "CenterScreen"
        $Global:form.Size = New-Object System.Drawing.Size(400,200) #(W,H)
        ## Domain to Add
        $Script:lblAddDomain = New-Object System.Windows.Forms.Label   
        $Script:lblAddDomain.Text = "Domain Enable for DKIM:"
        $Script:lblAddDomain.Top = 40 ; $Script:lblAddDomain.Left = 10; $Script:lblAddDomain.Width=120 ;$Script:lblAddDomain.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblAddDomain)    # Add to Form 
        $Script:txtAddDomain = New-Object Windows.Forms.TextBox  
        $Script:txtAddDomain.TabIndex = 0 # set Tab Order 
        $Script:txtAddDomain.Top = 40; $Script:txtAddDomain.Left = 140; $Script:txtAddDomain.Width = 200;  
        $Script:txtAddDomain.Text = ""
        $Global:form.Controls.Add($Script:txtAddDomain)    # Add to Form 
        Add-FormStandardButtons
        Publish-Form

        If ($Global:Result -eq "OK")
        {
            $NewDom = $Script:txtAddDomain.Text.TrimStart("@")
            If (($NewDom -like "*ul.com*") -or ($NewDom -like "*ul.org*"))
            {
                New-DkimSigningConfig -DomainName $NewDom -Enabled $true
                $Output = $wshell.Popup("DKIM Configured for: " + $NewDom,0,"Internal Relay",0+32) 
            }
        }
        else
        {
            $Output = $wshell.Popup("Adding of DKIM key for domain cancelled.",0,"Cancelled",0+32)
        }
    }

    If ($Global:chkCompAct.Checked -eq "Checked") 
    {
        Invoke-Expression -Command .\CompromisedAccount.ps1
    }

    If ($Global:chkNewSvcAct.Checked -eq "Checked") 
    {
    <#
        write-host ""
        Write-Host "`n     Enter ( 1) Enable/Configure Licensed Mailbox"
        write-host "           ( 2) Enable/Configure Oracle OFR/KFI Mailbox"
        write-host "           ( 3) Enable/Configure Surface Hub Device"
        write-host "           ( 4) Enable/Configure Remote Assist Account"
        write-host ""
        write-Host "           ( 0) No Changes and Exit"
        write-Host "     Enter Option Above? " -ForegroundColor Red -NoNewline
        $EnabMbx = Read-Host
        $Repeat = "N"
    #>

        Invoke-Expression -Command .\LicensedMailbox.ps1
        pause

        Do
        {
            switch ($EnabMbx)
	        {
                1
                {
#                    Remove-PSSession (Get-PSSession)
                    Invoke-Expression -Command .\LicensedMailbox.ps1

                    pause
                    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
                    Import-PSSession $session -AllowClobber
#                   Create USG group to control access
                    Write-Host "Creating groups to control access to the mailbox.  Hit return when the NewUSG input file has ready" -ForegroundColor Red -NoNewline
                    $Ret = Read-Host
                    Invoke-Expression -Command .\NewUSG.ps1
#                   Add Mailbox/Folder Permissions and Apply Retention Policy
                    Write-Host "Applying folder permissions and retention policy to the mailbox.  Hit return when the AddFolderPermissionsinput file has ready and the mailbox has been provisioned." -ForegroundColor Red -NoNewline
                    $Ret = Read-Host
                    Invoke-Expression -Command .\AddFolderPermissions.ps1
                }

                2
                {
                    Remove-PSSession (Get-PSSession)
                    Invoke-Expression -Command .\EnableSvcActAssignLicense.ps1
                    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
                    Import-PSSession $session -AllowClobber
#                   Create USG group to control access
                    Write-Host "Creating groups to control access to the mailbox.  Hit return when the NewUSG input file has ready" -ForegroundColor Red -NoNewline
                    $Ret = Read-Host
                    Invoke-Expression -Command .\NewUSG.ps1
#                   Add Mailbox/Folder Permissions and Apply Retention Policy
                    Write-Host "Applying folder permissions and retention policy to the mailbox.  Hit return when the AddFolderPermissionsinput file has ready and the mailbox has been provisioned." -ForegroundColor Red -NoNewline
                    $Ret = Read-Host
                    Invoke-Expression -Command .\AddFolderPermissions.ps1
                }
                        
                3
                {
                    Remove-PSSession (Get-PSSession)
                    Invoke-Expression -Command .\SetupSurfaceHubDevice.ps1
                    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
                    Import-PSSession $session -AllowClobber
                }

                4
                {
                    Remove-PSSession (Get-PSSession)
                    Invoke-Expression -Command .\EnableSvcActAssignLicense.ps1
                    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
                    Import-PSSession $session -AllowClobber
#                   Create USG group to control access
                    Write-Host "Creating groups to control access to the mailbox.  Hit return when the NewUSG input file has ready" -ForegroundColor Red -NoNewline
                    $Ret = Read-Host
                    Invoke-Expression -Command .\NewUSG.ps1
#                   Add Mailbox/Folder Permissions and Apply Retention Policy
                    Write-Host "Applying folder permissions and retention policy to the mailbox.  Hit return when the AddFolderPermissions input file has ready" -ForegroundColor Red -NoNewline
                    $Ret = Read-Host
                    Invoke-Expression -Command .\AddFolderPermissions.ps1
                }
            }
            write-host "`nDo you need to enable another device of this type (Y/N)? " -ForegroundColor Cyan -NoNewline
            $Repeat = Read-Host
        }while($Repeat -eq "Y")
    }

    If ($Global:chkOneDrv.Checked -eq "Checked")
	{
		write-host "Enter Employee Number to get usage details for: "
		$EmpID = read-host
		$EmpID = $EmpID + "_global_ul_com"
		write-host "Connecting to SharePoint"
		invoke-expression -Command .\ConnectO365SPO.ps1
		$Site = Get-SPOSite ("https://ul-my.sharepoint.com/personal/"+ $EmpID)
		write-host "Current Site Usage: " $site.StorageUsageCurrent
	}

    If ($Global:chkNewCred.Checked -eq "Checked")
    {
        Invoke-Expression -Command .\CreateNewCredFile.ps1
    }

    If ($Global:chkMiscCmd.Checked -eq "Checked")
    {
        invoke-item '.\Instructions\O365 powershell commands.docx'
    }

}While ($Global:MainResult -eq "OK")	