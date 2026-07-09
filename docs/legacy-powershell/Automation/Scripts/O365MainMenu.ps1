#
#   O365 Team Main Menu
#
#
#  06-16-2021 - SAG - Added item to search the Mailbox Audit Log
#  07-30-2021 - SAG - Working on changes to Service Account Configs
#  02-16-2023 - SAG - Added menu option to configured new SDAP team members
#  05-13-2024 - SAG - Changed the layout of the menu to be in 2 columns and added connecting to the SDAP menu
#  05-22-2024 - SAG - Added option to Permanently Delete an AzureAD Account
#  08-27-2025 - SAG - Added section to delay with SoftDeletedMailboxes that are not being purged

Function Build-AdminMenuForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "O365 Admin Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 700 ; $form.Height = 480  # Make the form wider 
    Add-FormStandardButtons

    #Build Column1
    $LeftCol1 = 60
    $TitleCol1 = 40
    $TopCol1 = 20
    $TopLoc = 20 

    ## Label and TextBox  
    ## Title Line
    $Global:lblTitle1 = New-Object System.Windows.Forms.Label   
        $Global:lblTitle1.Text = "AD Activites"
        $Global:lblTitle1.Top = $TopCol1 ; $Global:lblTitle1.Left = $TitleCol1+60; $Global:lblTitle1.Width=120 ;$Global:lblTitle1.AutoSize = $true 
        $Global:form.Controls.Add($lblTitle1)    # Add to Form 

    ## Add AD Site
    $TopCol1 = $TopCol1 + 20
    $Script:chkADSite = New-Object Windows.Forms.RadioButton 
        $Script:chkADSite.Left = $LeftCol1; $Script:chkADSite.Width = 275; $Script:chkADSite.Top = $TopCol1
        $Script:chkADSite.Text = "Add New AD 3 Letter Site" 
        $Script:chkADSite.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chkADSite) 
 
    ## Create New Service Account
    $TopCol1 = $TopCol1 + 20
    $Global:chkGenAcct = New-Object Windows.Forms.RadioButton 
        $Global:chkGenAcct.Left = $LeftCol1; $Global:chkGenAcct.Width = 275; $Global:chkGenAcct.Top = $TopCol1
        $Global:chkGenAcct.Text = "Create New Generic Account" 
        $Global:chkGenAcct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkGenAcct) 
        # Obtain Value with: $Global:chkGenAcct.Checked

    ## Locked PC Information
    $TopCol1 = $TopCol1 + 20
    $Global:chkLockPCInf = New-Object Windows.Forms.RadioButton 
        $Global:chkLockPCInf.Left = $LeftCol1; $Global:chkLockPCInf.Width = 275; $Global:chkLockPCInf.Top = $TopCol1
        $Global:chkLockPCInf.Text = "Get PC Locked Information" 
        $Global:chkLockPCInf.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkLockPCInf) 
        # Obtain Value with: $Global:chkLockPCInf.Checked

    ## Update Service Account
    $TopCol1 = $TopCol1 + 20
    $Global:chkPermDelAZADUsr = New-Object Windows.Forms.RadioButton
        $Global:chkPermDelAZADUsr.Left = $LeftCol1; $Global:chkPermDelAZADUsr.Width = 275; $Global:chkPermDelAZADUsr.Top = $TopCol1
        $Global:chkPermDelAZADUsr.Text = "Permanently Delete AzureAD User" 
        $Global:chkPermDelAZADUsr.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkPermDelAZADUsr) 

    ## Update Service Account
    $TopCol1 = $TopCol1 + 20
    $Global:chkUpdSvcAcct = New-Object Windows.Forms.RadioButton
        $Global:chkUpdSvcAcct.Left = $LeftCol1; $Global:chkUpdSvcAcct.Width = 275; $Global:chkUpdSvcAcct.Top = $TopCol1
        $Global:chkUpdSvcAcct.Text = "Update Generic Account Details" 
        $Global:chkUPdSvcAcct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkUpdSvcAcct) 

    ## Title Line
    $TopCol1 = $TopCol1 + 35
    $Global:lblTitle2 = New-Object System.Windows.Forms.Label   
        $Global:lblTitle2.Text = "O365 Activites"
        $Global:lblTitle2.Top = $TopCol1 ; $Global:lblTitle2.Left = $TitleCol1+60; $Global:lblTitle2.Width=120 ;$Global:lblTitle2.AutoSize = $true 
        $Global:form.Controls.Add($lblTitle2)    # Add to Form 

    ## Add New Domain
    $TopCol1 = $TopCol1 + 20
    $Script:chkAddDom = New-Object Windows.Forms.RadioButton 
        $Script:chkAddDom.Left = $LeftCol1; $Script:chkAddDom.Width = 275; $Script:chkAddDom.Top = $TopCol1
        $Script:chkAddDom.Text = "Add New Domain to the Tenant" 
        $Script:chkAddDom.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chkAddDom) 
        # Obtain Value with: $Script:chkAddDom.Checked

    ## Add New SDAP Team Member
    $TopCol1 = $TopCol1 + 20
    $Script:chkSDAPMem = New-Object Windows.Forms.RadioButton 
        $Script:chkSDAPMem.Left = $LeftCol1; $Script:chkSDAPMem.Width = 275; $Script:chkSDAPMem.Top = $TopCol1
        $Script:chkSDAPMem.Text = "Add New SDAP Team Member" 
        $Script:chkSDAPMem.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chkSDAPMem) 
        # Obtain Value with: $Script:chkSDAPMem.Checked

    ## Add Enable DKIM keys for New Domain
    $TopCol1 = $TopCol1 + 20
    $Global:chkDKIMKey = New-Object Windows.Forms.RadioButton 
        $Global:chkDKIMKey.Left = $LeftCol1; $Global:chkDKIMKey.Width = 275; $Global:chkDKIMKey.Top = $TopCol1
        $Global:chkDKIMKey.Text = "Add/Enable DKIM Keys for New Domain" 
        $Global:chkDKIMKey.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkDKIMKey) 
        # Obtain Value with: $Global:chkDKIMKey.Checked

    ## Add Remove SoftDeletedMailboxes
    $TopCol1 = $TopCol1 + 20
    $Global:chkSftDelHolds = New-Object Windows.Forms.RadioButton 
        $Global:chkSftDelHolds.Left = $LeftCol1; $Global:chkSftDelHolds.Width = 275; $Global:chkSftDelHolds.Top = $TopCol1
        $Global:chkSftDelHolds.Text = "Remove Holds Applied to SoftDeleted Mailboxes" 
        $Global:chkSftDelHolds.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkSftDelHolds) 
        # Obtain Value with: $Global:chkSftDelHolds.Checked

    ## Search Mailbox Audit Log
    $TopCol1 = $TopCol1 + 20
    $Global:chkSrchMbxALog = New-Object Windows.Forms.RadioButton 
        $Global:chkSrchMbxALog.Left = $LeftCol1; $Global:chkSrchMbxALog.Width = 275; $Global:chkSrchMbxALog.Top = $TopCol1
        $Global:chkSrchMbxALog.Text = "Search Mailbox Audit Log" 
        $Global:chkSrchMbxALog.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkSrchMbxALog) 
        # Obtain Value with: $Global:chkSrchMbxALog.Checked


    ## Title Line
    $TopCol1 = $TopCol1 + 35
    $Global:lblTitle3 = New-Object System.Windows.Forms.Label   
        $Global:lblTitle3.Text = "O365 Miscellaneous Activites"
        $Global:lblTitle3.Top = $TopCol1 ; $Global:lblTitle3.Left = $TitleCol1+40; $Global:lblTitle3.Width=120 ;$Global:lblTitle3.AutoSize = $true 
        $Global:form.Controls.Add($lblTitle3)    # Add to Form 

    ## Configure Service Accounts/Surface Hub Devices
    $TopCol1 = $TopCol1 + 20
    $Global:chkNewSvcAct = New-Object Windows.Forms.RadioButton 
        $Global:chkNewSvcAct.Left = $LeftCol1; $Global:chkNewSvcAct.Width = 275; $Global:chkNewSvcAct.Top = $TopCol1
        $Global:chkNewSvcAct.Text = "Configure Service Accounts/Surface Hub Devices" 
        $Global:chkNewSvcAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkNewSvcAct) 
        # Obtain Value with: $Global:chkNewSvcAct.Checked

    ## Enable/Disable Compromised Account
    $TopCol1 = $TopCol1 + 20
    $Global:chkCompAct = New-Object Windows.Forms.RadioButton 
        $Global:chkCompAct.Left = $LeftCol1; $Global:chkCompAct.Width = 275; $Global:chkCompAct.Top = $TopCol1
        $Global:chkCompAct.Text = "Enable/Disable Compromised Account" 
        $Global:chkCompAct.Checked = $Global:chkCompAct.Checked   # set a default value 
        $Global:form.Controls.Add($Global:chkCompAct) 
        # Obtain Value with: $Global:chkCompAct.Checked

    ## OneDrive Space Usage
    $TopCol1 = $TopCol1 + 20
    $Global:chkOneDrv = New-Object Windows.Forms.RadioButton 
        $Global:chkOneDrv.Left = $LeftCol1; $Global:chkOneDrv.Width = 275; $Global:chkOneDrv.Top = $TopCol1
        $Global:chkOneDrv.Text = "Get Users OneDrive Space Usage" 
        $Global:chkOneDrv.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkOneDrv) 
        # Obtain Value with: $Global:chkOneDrv.Checked


    #Build Column2
    $LeftCol2 = 350
    $TitleCol2 = 375
    $TopCol2 = 20
   
    ## Title Line
    $Global:lblTitle3 = New-Object System.Windows.Forms.Label   
        $Global:lblTitle3.Text = "O365 Menus"
        $Global:lblTitle3.Top = $TopCol2 ; $Global:lblTitle3.Left = $LeftCol2+60; $Global:lblTitle3.Width=120 ;$Global:lblTitle3.AutoSize = $true 
        $Global:form.Controls.Add($lblTitle3)    # Add to Form    
 
    ## DL/O365 Group/Security Group 
    $TopCol2 = $TopCol2 + 20
    $Global:chkDLUniSecGrp = New-Object Windows.Forms.RadioButton 
        $Global:chkDLUniSecGrp.Left = $LeftCol2; $Global:chkDLUniSecGrp.Width = 300; $Global:chkDLUniSecGrp.Top = $TopCol2
        $Global:chkDLUniSecGrp.Text = "Distribution List / O365 Group / Security Group Menu" 
        $Global:chkDLUniSecGrp.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkDLUniSecGrp) 
        # Obtain Value with: $Global:chkDLUniSecGrp.Checked

    ## Legal Hold 
    $TopCol2 = $TopCol2 + 20
    $Global:chkLglHold = New-Object Windows.Forms.RadioButton 
        $Global:chkLglHold.Left = $LeftCol2; $Global:chkLglHold.Width = 275; $Global:chkLglHold.Top = $TopCol2
        $Global:chkLglHold.Text = "Legal Hold Activities Menu" 
        $Global:chkLglHold.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkLglHold) 
        # Obtain Value with: $Global:chkLglHold.Checked

    ## Licensing Menu
    $TopCol2 = $TopCol2 + 20
    $Global:chkLicMenu = New-Object Windows.Forms.RadioButton 
        $Global:chkLicMenu.Left = $LeftCol2; $Global:chkLicMenu.Width = 275; $Global:chkLicMenu.Top = $TopCol2 
        $Global:chkLicMenu.Text = "Licensing Menu" 
        $Global:chkLicMenu.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkLicMenu) 
        # Obtain Value with: $Global:chkLicMenu.Checked
 
    ## M&A Activities
    $TopCol2 = $TopCol2 + 20
    $Global:chkMAActv = New-Object Windows.Forms.RadioButton 
        $Global:chkMAActv.Left = $LeftCol2; $Global:chkMAActv.Width = 275; $Global:chkMAActv.Top = $TopCol2
        $Global:chkMAActv.Text = "M&A Activities Menu" 
        $Global:chkMAActv.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkMAActv) 
        # Obtain Value with: $Global:chkMAActv.Checked

    ## Reporting
    $TopCol2 = $TopCol2 + 20
    $Global:chkRptMenu = New-Object Windows.Forms.RadioButton 
        $Global:chkRptMenu.Left = $LeftCol2; $Global:chkRptMenu.Width = 275; $Global:chkRptMenu.Top = $TopCol2
        $Global:chkRptMenu.Text = "Reporting Menu" 
        $Global:chkRptMenu.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkRptMenu) 
        # Obtain Value with: $Global:chkRptMenu.Checked

    ## Shared Room/Resource
    $TopCol2 = $TopCol2 + 20
    $Global:chkRoomRes = New-Object Windows.Forms.RadioButton 
        $Global:chkRoomRes.Left = $LeftCol2; $Global:chkRoomRes.Width = 275; $Global:chkRoomRes.Top = $TopCol2
        $Global:chkRoomRes.Text = "Room or Resource Menu" 
        $Global:chkRoomRes.Checked = $Global:chkRoomRes.Checked   # set a default value 
        $Global:form.Controls.Add($Global:chkRoomRes) 
        # Obtain Value with: $Global:chkRoomRes.Checked

    ## Account Provisioning Menu
    $TopCol2 = $TopCol2 + 20
    $Global:chkSDAPMenu = New-Object Windows.Forms.RadioButton 
        $Global:chkSDAPMenu.Left = $LeftCol2; $Global:chkSDAPMenu.Width = 275; $Global:chkSDAPMenu.Top = $TopCol2
        $Global:chkSDAPMenu.Text = "SDAP Team Menu" 
        $Global:chkSDAPMenu.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkSDAPMenu) 
        # Obtain Value with: $Global:chkSDAPMenu.Checked

    ## Shared Mailbox
    $TopCol2 = $TopCol2 + 20
    $Global:chkShrMbx = New-Object Windows.Forms.RadioButton 
        $Global:chkShrMbx.Left = $LeftCol2; $Global:chkShrMbx.Width = 275; $Global:chkShrMbx.Top = $TopCol2
        $Global:chkShrMbx.Text = "Shared Mailbox Menu" 
        $Global:chkShrMbx.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkShrMbx) 
        # Obtain Value with: $Global:chkShrMbx.Checked

    ## Unified Messaging
    $TopCol2 = $TopCol2 + 20
    $Global:chkUniMsg = New-Object Windows.Forms.RadioButton 
        $Global:chkUniMsg.Left = $LeftCol2; $Global:chkUniMsg.Width = 275; $Global:chkUniMsg.Top = $TopCol2
        $Global:chkUniMsg.Text = "Unified Messaging Menu" 
        $Global:chkUniMsg.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkUniMsg) 
        # Obtain Value with: $Global:chkUniMsg.Checked

    ## User Mailbox
    $TopCol2 = $TopCol2 + 20
    $Global:chkUsrMbx = New-Object Windows.Forms.RadioButton 
        $Global:chkUsrMbx.Left = $LeftCol2; $Global:chkUsrMbx.Width = 275; $Global:chkUsrMbx.Top = $TopCol2
        $Global:chkUsrMbx.Text = "User Mailbox Menu" 
        $Global:chkUsrMbx.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkUsrMbx) 
        # Obtain Value with: $Global:chkUsrMbx.Checked

    ## Title Line
    $TopCol2 = $TopCol2 + 35
    $Global:lblTitle4 = New-Object System.Windows.Forms.Label   
        $Global:lblTitle4.Text = "Miscellaneous Items"
        $Global:lblTitle4.Top = $TopCol2 ; $Global:lblTitle4.Left = $LeftCol2+40; $Global:lblTitle4.Width=120 ;$Global:lblTitle4.AutoSize = $true 
        $Global:form.Controls.Add($lblTitle4)    # Add to Form    

    ## Create New Credential File
    $TopCol2 = $TopCol2 + 20
    $Global:chkNewCred = New-Object Windows.Forms.RadioButton 
        $Global:chkNewCred.Left = $LeftCol2; $Global:chkNewCred.Width = 275; $Global:chkNewCred.Top = $TopCol2
        $Global:chkNewCred.Text = "Create New Credential File" 
        $Global:chkNewCred.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkNewCred) 
        # Obtain Value with: $Global:chkNewCred.Checked

    ## Find SDAP Activities
    $TopCol2 = $TopCol2 + 20
    $Global:chkSDAPAct = New-Object Windows.Forms.RadioButton 
        $Global:chkSDAPAct.Left = $LeftCol2; $Global:chkSDAPAct.Width = 275; $Global:chkSDAPAct.Top = $TopCol2
        $Global:chkSDAPAct.Text = "Find SDAP Report Files for a User" 
        $Global:chkSDAPAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkSDAPAct) 
        # Obtain Value with: $Global:chkSDAPAct.Checked

    ## Misc Powershell Commands
    $TopCol2 = $TopCol2 + 20
    $Global:chkMiscCmd = New-Object Windows.Forms.RadioButton 
        $Global:chkMiscCmd.Left = $LeftCol2; $Global:chkMiscCmd.Width = 275; $Global:chkMiscCmd.Top = $TopCol2
        $Global:chkMiscCmd.Text = "Open Miscellaneous PS Commands Document" 
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
    if ($Global:stopwatch.elapsed -ge $Global:reconnectThreshold)
    {
        # Close all sessions
        write-host "Connection Threshold Exceeded -- Reconnecting to Office365" -ForegroundColor Red
        invoke-expression -Command E:\O365AdminShared\Scripts\ConnectO365_Graph.ps1
        $Global:stopwatch = [diagnostics.stopwatch]::StartNew()
    }
}


#### Start of Script

$Global:LiveCred = ""
$Global:CredFile = ""
$File = "e:\Automation\scripts\O365AdminMenu.ps1"
$Global:MainResult = "Continue"

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
Add-Type -AssemblyName System.Windows.Forms
invoke-expression -Command E:\O365AdminShared\Scripts\ConnectO365_Graph.ps1

If ($Global:MainResult -eq "Continue")
{
    Do
    {
        $Global:OKDetails = "Continue"
        $Global:stopwatch = [diagnostics.stopwatch]::StartNew()
        Build-AdminMenuForm
        Publish-MMForm

 #       Check-Reconnect
        		
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

        If ($Script:chkSDAPMem.Checked -eq "Checked")
        {
            invoke-expression -Command E:\Automation\Scripts\NewSDAPTeamMember.ps1
        }

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
    
        If ($Script:chkADSite.Checked -eq "Checked")
        {
            Invoke-Expression -Command .\NewADSite.ps1
        }

        If ($Global:chkPermDelAZADUsr.Checked -eq "Checked")
        {
            Invoke-Expression -Command .\PermanentlyDeleteAZADUsr.ps1
        }

        If ($Script:chkAddDom.Checked -eq "Checked") 
        {
            $Global:form = New-Object Windows.Forms.Form 
            $Global:form.FormBorderStyle = "FixedToolWindow" 
            $Global:form.Text = "Add New Domain Tenant" 
            $Global:form.StartPosition = "CenterScreen"
            $Global:form.Size = New-Object System.Drawing.Size(400,200) #(W,H)
            ## Domain to Add
            $Script:lblAddDomain = New-Object System.Windows.Forms.Label   
            $Script:lblAddDomain.Text = "Domain to Add:"
            $Script:lblAddDomain.Top = 40 ; $Script:lblAddDomain.Left = 10; $Script:lblAddDomain.Width=120 ;$Script:lblAddDomain.AutoSize = $true 
            $Global:form.Controls.Add($Script:lblAddDomain)    # Add to Form 
            $Script:txtAddDomain = New-Object Windows.Forms.TextBox  
            $Script:txtAddDomain.TabIndex = 0 # set Tab Order 
            $Script:txtAddDomain.Top = 40; $Script:txtAddDomain.Left = 130; $Script:txtAddDomain.Width = 200;  
            $Script:txtAddDomain.Text = ""
            $Global:form.Controls.Add($Script:txtAddDomain)    # Add to Form
            $Global:InputFocus = $Script:txtAddDomain 
            Add-FormStandardButtons
            Publish-Form

            If ($Global:Result -eq "OK")
            {
                Write-host "Added New Domains: " $Script:txtAddDomain.Text
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

        If ($Global:chkSftDelHolds.Checked -eq "Checked")
        {
            Invoke-Expression -Command .\RemoveSoftDeletedMailboxes.ps1
        }

        If ($Global:chkCompAct.Checked -eq "Checked") 
        {
            Invoke-Expression -Command .\CompromisedAccount.ps1
        }

        If ($Global:chkNewSvcAct.Checked -eq "Checked") 
        {
            Invoke-Expression -Command .\LicensedMailbox.ps1
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

        If ($Global:chkSDAPMenu.checked -eq "Checked")
        {
            $Global:FromO365Team = "Yes"
            cd e:\SDAP\Scripts\
            Invoke-Expression -Command .\SDAPAdminMenu.ps1
            cd e:\Automation\Scripts
            $Global:MainResult = "OK"
        }

        If ($Global:chkNewCred.Checked -eq "Checked")
        {
            Invoke-Expression -Command .\CreateNewCredFile.ps1
        }

        If ($Global:chkSDAPAct.Checked -eq "Checked")
        {
    #        Invoke-Expression -Command .\LocateReportFiles.ps1
            Invoke-Expression -Command "E:\O365AdminShared\Scripts\FindSDAPLogFiles.ps1"
        }

        If ($Global:chkMiscCmd.Checked -eq "Checked")
        {
            invoke-item '.\Instructions\O365 powershell commands.docx'
        }

    }While ($Global:MainResult -eq "OK")
}