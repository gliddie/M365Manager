#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange
#    'Description  : Removes SMTP Address from User Account
#    'Called By    : SDAPAdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 11/22/2017
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 11/22/2017 Created from the EmergencyDisablement script
#    '             : 03/23/2018 Added code into the Promote Option to check if the SIP address also should be changed
#    '             : 01/26/2020 Fixed some formatting issues with the out file
#    '             : 05/03/2021 Added code to set the target address and mail address when changing the primary address
#    '             : 06/01/2021 Fixed the updating of the target and mail address and the displaying of the details.
#    '             : 08/25/2021 Removed duplicate functions WriteReportEvent, WriteLogEvent and getting the whoami information
#    '             : 01/25/2022 Redesigned script to use Forms to provide deatils and mail request
#    '
# ==========================================================================
#
#################################################################################

Function Add-ActionBoxes
{
    $ActLeft = 380
    ## Change Primary Address
    $Script:chkNewPrimary = New-Object Windows.Forms.RadioButton
        $Script:chkNewPrimary.Left = $ActLeft; $Script:chkNewPrimary.Width = 200; $Script:chkNewPrimary.Top = 10  
        $Script:chkNewPrimary.Text = "Add New Primary/SIP Address" 
        $Script:chkNewPrimary.Checked = $False   # set a default value 
        $Script:chkNewPrimary.TabIndex = 1
        $Global:form.Controls.Add($Script:chkNewPrimary) 
        # Obtain Value with: $Script:chkNewPrimary.Checked
        $Script:chkNewPrimary.Add_Click({
            $Script:lblNewAddr.Text = "New Primary Address:"
            $Script:txtNewAddr.Text = ""
            $Script:txtNewAddr.ReadOnly = $false
            Show-InputBoxes
        })

    ## Add Proxy Address
    $Script:chkAddProxy = New-Object Windows.Forms.RadioButton
        $Script:chkAddProxy.Left = $ActLeft; $Script:chkAddProxy.Width = 200; $Script:chkAddProxy.Top = 30
        $Script:chkAddProxy.Text = "Add New Proxy Address"
        $Script:chkAddProxy.Checked = $False   # set a default value 
        $Script:chkAddProxy.TabIndex = 0
        $Global:form.Controls.Add($Script:chkAddProxy) 
        # Obtain Value with: $Global:chkThis.Checked
        $Script:chkAddProxy.Add_Click({
            $Script:lblNewAddr.Text = "New Proxy Address:"
            $Script:txtNewAddr.Text = ""
            $Script:txtNewAddr.ReadOnly = $false
            Show-InputBoxes
        })

    ## Promote Proxy Address
    $Script:chkPromoProxy = New-Object Windows.Forms.RadioButton
        $Script:chkPromoProxy.Left = $ActLeft; $Script:chkPromoProxy.Width = 230; $Script:chkPromoProxy.Top = 50  
        $Script:chkPromoProxy.Text = "Promote Proxy to Primary/SIP Address" 
        $Script:chkPromoProxy.Checked = $False   # set a default value 
        $Script:chkPromoProxy.TabIndex = 1
        $Global:form.Controls.Add($Script:chkPromoProxy) 
        # Obtain Value with: $Script:chkPromoProxy.Checked
        $Script:chkPromoProxy.Add_Click({
            $Script:lblNewAddr.Text = "Selected Address:"
            $Script:txtNewAddr.Text = $Script:txtAddresses.SelectedItem
            $Script:txtNewAddr.ReadOnly = $true
            Show-InputBoxes
            If ($Script:txtAddresses.SelectedItem -eq $null)
            {
                $Output = $wshell.Popup("Select one of the addresses listed in the 'Other Addresses'.",0,"Select An Address",0+32)  
            }
        })

    ## Remove Proxy Address
    $Script:chkRemoveAddr = New-Object Windows.Forms.RadioButton
        $Script:chkRemoveAddr.Left = $ActLeft; $Script:chkRemoveAddr.Width = 200; $Script:chkRemoveAddr.Top = 70
        $Script:chkRemoveAddr.Text = "Remove Existing Proxy Address"
        $Script:chkRemoveAddr.Checked = $False   # set a default value 
        $Script:chkRemoveAddr.TabIndex = 1
        $Global:form.Controls.Add($Script:chkRemoveAddr) 
        # Obtain Value with: $Script:chkRemoveAddr.Checked
        $Script:chkRemoveAddr.Add_Click({
            $Script:lblNewAddr.Text = "Remove Address:"
            $Script:txtNewAddr.Text = $Script:txtAddresses.SelectedItem
            $Script:txtNewAddr.ReadOnly = $true
            Show-InputBoxes
            If ($Script:txtAddresses.SelectedItem -eq $null)
            {
                $Output = $wshell.Popup("Select one of the addresses listed in the 'Other Addresses'.",0,"Select An Address",0+32)  
            }
        })

    If ($ADUser.EmailAddress -ne $Script:SIP)
    {
        ## Update SIP Address
        $Script:chkChgSIP = New-Object Windows.Forms.RadioButton
            $Script:chkChgSIP.Left = $ActLeft; $Script:chkChgSIP.Width = 200; $Script:chkChgSIP.Top = 90  
            $Script:chkChgSIP.Text = "Update SIP Address"
            $Script:chkChgSIP.Checked = $False   # set a default value 
            $Script:chkChgSIP.TabIndex = 0
            $Global:form.Controls.Add($Script:chkChgSIP) 
            # Obtain Value with: $Global:chkThis.Checked
            $Script:chkChgSIP.Add_Click({
                $Script:lblNewAddr.Text = "New SIP Address:"
                $Script:txtNewAddr.Text = $ADUser.EmailAddress
                $Script:txtNewAddr.ReadOnly = $true
                Show-InputBoxes
                $Output = $wshell.Popup("The SIP Address will be set to match the primary email address.",0,"Set SIP Address",0+32) 
            })
        }
}

Function Build-SDAPUserInfoForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Add/Promote/Remove SMTP/SIP Address" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 630 ; $form.Height = 400  # Make the form wider 
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    
    $Tab = 0
    $Top = 10
    ## Employee ID
    $Script:lblHost = New-Object System.Windows.Forms.Label   
        $Script:lblHost.Text = "Employee No:"  
        $Script:lblHost.Top = $Top ; $Script:lblHost.Left = 5; $Script:lblHost.Width=120 ;$Script:lblHost.AutoSize = $true 
        $form.Controls.Add($Script:lblHost)    # Add to Form 
        # 
        $Script:txtHost = New-Object Windows.Forms.TextBox
        $Script:txtHost.ReadOnly = $true; 
        $Script:txtHost.TabIndex = $Tab++
        $Script:txtHost.Top = $Top; $Script:txtHost.Left = 130; $Script:txtHost.Width = 70;  
        $Script:txtHost.Text = $Global:txtInpEmpNo.Text
        $Global:form.Controls.Add($Script:txtHost)    # Add to Form 
 
    $Top = $Top + 30
    ## Employee Name
    $Script:lblName = New-Object System.Windows.Forms.Label   
        $Script:lblName.Text = "Name:"  
        $Script:lblName.Top = $Top ; $Script:lblName.Left = 5; $Script:lblName.Width=120 ;$Script:lblName.AutoSize = $true 
        $form.Controls.Add($Script:lblName)    # Add to Form 
        # 
        $Script:txtName = New-Object Windows.Forms.TextBox
        $Script:txtName.ReadOnly = $true;
        $Script:txtName.TabIndex = $Tab++ # set Tab Order  
        $Script:txtName.Top = $Top; $Script:txtName.Left = 130; $Script:txtName.Width = 120;  
        $Script:txtName.Text = $ADUser.DisplayName
        $Global:form.Controls.Add($Script:txtName)    # Add to Form 

    $Top = $Top + 30
    ## Employee $Type
    $Script:lblType = New-Object System.Windows.Forms.Label   
        $Script:lblType.Text = "Type:"  
        $Script:lblType.Top = $Top ; $Script:lblType.Left = 5; $Script:lblType.Width=120 ;$Script:lblType.AutoSize = $true 
        $form.Controls.Add($Script:lblType)    # Add to Form 
        # 
        $Script:txtType = New-Object Windows.Forms.TextBox
        $Script:txtType.ReadOnly = $true;  
        $Script:txtHost.TabIndex = $Tab++ # set Tab Order 
        $Script:txtType.Top = $Top; $Script:txtType.Left = 130; $Script:txtType.Width = 100;  
        $Script:txtType.Text = $ADUser.ExtensionAttribute1
        $Global:form.Controls.Add($Script:txtType)    # Add to Form
    
    $Top = $Top + 30
    ##PrimaryEmailAddress
    $Script:lblPrimayAddr = New-Object System.Windows.Forms.Label   
        $Script:lblPrimayAddr.Text = "Primary Email Address:"  
        $Script:lblPrimayAddr.Top = $Top ; $Script:lblPrimayAddr.Left = 5; $Script:lblPrimayAddr.Width=120 ;$Script:lblPrimayAddr.AutoSize = $true 
        $form.Controls.Add($Script:lblPrimayAddr)    # Add to Form 
        # 
        $Script:txtPrimayAddr = New-Object Windows.Forms.TextBox
        $Script:txtPrimayAddr.ReadOnly = $true;  
        $Script:txtPrimayAddr.TabIndex = $Tab++ # set Tab Order 
        $Script:txtPrimayAddr.Top = $Top; $Script:txtPrimayAddr.Left = 130; $Script:txtPrimayAddr.Width = 200;
        $Script:txtPrimayAddr.Text = $ADUser.EmailAddress
        $Global:form.Controls.Add($Script:txtPrimayAddr)    # Add to Form

    $Top = $Top + 30
    ##SIPAddress
    $Script:lblSIPAddr = New-Object System.Windows.Forms.Label   
        $Script:lblSIPAddr.Text = "SIP Address:"  
        $Script:lblSIPAddr.Top = $Top ; $Script:lblSIPAddr.Left = 5; $Script:lblSIPAddr.Width=120 ;$Script:lblSIPAddr.AutoSize = $true 
        $form.Controls.Add($Script:lblSIPAddr)    # Add to Form 
        # 
        $Script:txtSIPAddr = New-Object Windows.Forms.TextBox
        $Script:txtSIPAddr.ReadOnly = $true;  
        $Script:txtSIPAddr.TabIndex = $Tab++ # set Tab Order 
        $Script:txtSIPAddr.Top = $Top; $Script:txtSIPAddr.Left = 130; $Script:txtSIPAddr.Width = 200;
        $Script:txtSIPAddr.Text = ($ADUser.'msRTCSIP-PrimaryUserAddress'.Substring(4,$ADUser.'msRTCSIP-PrimaryUserAddress'.Length-4) )
        $Global:form.Controls.Add($Script:txtSIPAddr)    # Add to Form

    $Top = $Top + 30
    ##Addresses
    $Script:lblAddresses = New-Object System.Windows.Forms.Label   
        $Script:lblAddresses.Text = "Other Addresses:"
        $Script:lblAddresses.Top = $Top ; $Script:lblAddresses.Left = 5; $Script:lblAddresses.Width=120 ;$Script:lblAddresses.AutoSize = $true 
        $form.Controls.Add($Script:lblAddresses)    # Add to Form 
        # 
        $Script:txtAddresses = New-Object Windows.Forms.ListBox
#        $Script:txtAddresses.ScrollBars = $true
        $Script:txtAddresses.Top = $Top; $Script:txtAddresses.Left = 130; $Script:txtAddresses.Height = 70; $Script:txtAddresses.Width = 300;
        $Global:form.Controls.Add($Script:txtAddresses)    # Add to Form
        If ($ADUser.EmailAddresses.Length -ne 0)
        {
            $LocArray = $ADUser.proxyAddresses.split(",")
            $i=0   # Counter 
            foreach ($element in $LocArray)
            {
                # Loop through Azure list and add to listbox
                If (($element -ne ("SMTP:" + $ADUser.EmailAddress)) -and ($element -notlike "SIP:*"))
                {
                    [void] $Script:txtAddresses.Items.Add($element.TrimStart("smtp:"))  # Add element to listbox 
                }
            } 
        }
        $Script:txtAddresses.Add_Click({
            If (($Script:chkPromoProxy.Checked -eq "Checked") -or ($Script:chkRemoveAddr.Checked -eq "Checked"))
            {
                $Script:txtNewAddr.Text = $Script:txtAddresses.SelectedItem
            }
        })

    $Top = $Top + 80
    ##New/Selected Address
    $Script:lblNewAddr = New-Object System.Windows.Forms.Label
        $Script:lblNewAddr.Top = $Top ; $Script:lblNewAddr.Left = 5; $Script:lblNewAddr.Width=120 ;$Script:lblNewAddr.AutoSize = $true 
        $form.Controls.Add($Script:lblNewAddr)    # Add to Form 
        # 
        $Script:txtNewAddr = New-Object Windows.Forms.TextBox
        $Script:txtNewAddr.TabIndex = $Tab++ # set Tab Order 
        $Script:txtNewAddr.Top = $Top; $Script:txtNewAddr.Left = 130; $Script:txtNewAddr.Width = 300;
        $Global:form.Controls.Add($Script:txtNewAddr)    # Add to Form
        $Script:lblNewAddr.Visible = $False
        $Script:txtNewAddr.Visible = $False

    $Top = $Top + 30
    ##Ticket Number
    $Script:lblTicketNo = New-Object System.Windows.Forms.Label
        $Script:lblTicketNo.Text = "Ticket Number:"
        $Script:lblTicketNo.Top = $Top ; $Script:lblTicketNo.Left = 5; $Script:lblTicketNo.Width=120 ;$Script:lblTicketNo.AutoSize = $true 
        $form.Controls.Add($Script:lblTicketNo)    # Add to Form 
        # 
        $Script:txtTicketNo = New-Object Windows.Forms.TextBox
        $Script:txtTicketNo.TabIndex = $Tab++ # set Tab Order 
        $Script:txtTicketNo.Top = $Top; $Script:txtTicketNo.Left = 130; $Script:txtTicketNo.Width = 300;
        $Global:form.Controls.Add($Script:txtTicketNo)    # Add to Form
        $Script:lblTicketNo.Visible = $False
        $Script:txtTicketNo.Visible = $False

    $Top = $Top + 30
    ##ReportDetails
    $Script:lblRptFile = New-Object System.Windows.Forms.Label   
        $Script:lblRptFile.Text = "Report Details:"  
        $Script:lblRptFile.Top = $Top ; $Script:lblRptFile.Left = 5; $Script:lblRptFile.Width=100 ;$Script:lblRptFile.AutoSize = $true 
        $form.Controls.Add($Script:lblRptFile)    # Add to Form 
        # 
        $Script:txtRptFile = New-Object Windows.Forms.TextBox
        $Script:txtRptFile.ReadOnly = $true; 
        $Script:txtRptFile.TabIndex = $Tab++ # set Tab Order
        $Script:txtRptFile.Top = $Top; $Script:txtRptFile.Left = 130; $Script:txtRptFile.Width = 450; 
        $Script:txtRptFile.Text = $Script:ReportFile
        $Global:form.Controls.Add($Script:txtRptFile)    # Add to Form
        $Script:lblRptFile.Visible = $False
        $Script:txtRptFile.Visible = $False

    Add-FormStandardButtons
}

Function Show-InputBoxes
{
    $Script:lblNewAddr.Visible = $true
    $Script:txtNewAddr.Visible = $true
    $Script:lblTicketNo.Visible = $true
    $Script:txtTicketNo.Visible = $true
    $Script:lblRptFile.Visible = $true
    $Script:txtRptFile.Visible = $true
}

Function Address-Details
{
    $LineToWrite = "INFO" + "`t" + "         Employee Number: " + $Global:txtInpEmpNo.Text
    WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "           Employee Name: " + $ADUser.DisplayName
    WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "           Employee Type: " + $ADUser.ExtensionAttribute1
    WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "   Primary Email Address: " + $ADUser.EmailAddress
    WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "             SIP Address: " + $ADUser.'msRTCSIP-PrimaryUserAddress'
    WriteReportEvent

    $cnt = 0
    $EAddr = $ADUser.ProxyAddresses |Sort-Object ProxAddresses
    $collection = $EAddr.GetEnumerator()
    foreach ($EAddr in $collection)
    {
        If ($EAddr -like "sip:*")
        {
            If ($ADUser.'msRTCSIP-PrimaryUserAddress' -eq $EAddr)
            {
                Set-AdUser -Identity $Global:txtInpEmpNo.Text -Remove @{ProxyAddresses=$EAddr}
                $LineToWrite = "REMV" + "`t" + "    SIP Removed as Proxy: " + $EAddr
                WriteReportEvent
            }
        }
        else
        {
            If ($EAddr -clike "smtp:*")
            {
                If ($cnt -eq 0)
                {
                    $LineToWrite = "INFO" + "`t" + "         Other Addresses: " + $EAddr
                    $cnt++
                }
                else
                {
                    $LineToWrite = "INFO" + "`t" + "                          " + $EAddr
                }
                WriteReportEvent
            }
        }
    }
    $LineToWrite = "INFO" + "`t" + "          Target Address: " + $ADUser.TargetAddress
    WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "            Mail Address: " + $ADUser.Mail + "`n"
    WriteReportEvent
}

Function Address-Exists
{
    $InputFile = "E:\O365AdminShared\Data\UsrAddresses.csv"
    $ExistAddr = Import-Csv $InputFile
    $NoRec = $ExistAddr.count
    $FindAddr = $Script:txtNewAddr.Text
    $FindAddr = $FindAddr.Trim(" ")

    $AddrDet = ""

    Write-host "Checking to see if: " $FindAddr "is already in use...." -ForegroundColor Red
    $SrchAddr = "SMTP:" + $FindAddr
    $Script:MatchFound = "N"
    foreach ($ExistAddr in $ExistAddr)
    {
        if ($ExistAddr.EmailAddresses -like $SrchAddr)
        {
            $AddrDet = get-mailbox $FindAddr
            $LineToWrite = "FAIL" + "`t" + "No Changes Made Address " + $FindAddr + " in use by " + $AddrDet.Name
            WriteReportEvent     
            write-host "The address" $FindAddr "is use by" $AddrDet.Name
            $Script:MatchFound = "Y"
        }
    }
}

Function Form-PostChange
{
    $Script:txtAddresses.Items.Clear()
    If ($ADUser.EmailAddresses.Length -ne 0)
    {
        $LocArray = $ADUser.proxyAddresses.split(",")
        $i=0   # Counter 
        foreach ($element in $LocArray)
        {
            # Loop through Azure list and add to listbox
            If (($element -ne ("SMTP:" + $ADUser.EmailAddress)) -and ($element -notlike "SIP:*"))
            {
                [void] $Script:txtAddresses.Items.Add($element.TrimStart("smtp:"))  # Add element to listbox 
            }
        } 
    }
    $Script:txtNewAddr.Text = $ADUser.EmailAddress
    $Global:OKButton.Visible = $false
    $Global:CancelButton.Text = "Done"
    $Script:chkNewPrimary.Visible = $false
    $Script:chkAddProxy.Visible = $false
    $Script:chkPromoProxy.Visible = $false
    $Script:chkRemoveAddr.Visible = $false
    $Script:lblNewAddr.Visible = $false
    $Script:txtNewAddr.Visible = $false   
    $Script:lblTicketNo.Visible = $false
    $Script:txtTicketNo.Visible = $false
    $Script:lblRptFile.Visible = $true
    $Script:txtRptFile.Visible = $true
    Publish-Form
}

#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "SMTPAddrChange"
	$LogDrive		= "E:"
	$LogPath		= "\SDAP"
	$LogFolder		= "\" + $FileNAme
	$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
	$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"

Enter-EmpNoInputForm
Publish-Form

If ($Global:Result -eq "OK")
{
    If ($Global:txtInpEmpNo.Text -ne "")
    {
        $ENo = $Global:txtInpEmpNo.Text
        If ($ENo -notlike "*@*")
        {
            $ENo = $ENo + "@global.ul.com"
        }

        $Exists = [bool](Get-MsolUser -UserPrincipalName $ENo -ErrorAction SilentlyContinue)
        
        If ($Exists -eq $True)
        {
            $Script:ReportFile = $LogDirectory + "Report\Report-" + $FileName + "-EmpNo" + $Global:txtInpEmpNo.Text + "-Date"+ ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            CheckLogFiles

            $ADUser = Get-ADUser $Global:txtInpEmpNo.Text -Properties *
            If ($ADUser.'msRTCSIP-PrimaryUserAddress' -ne $null)
            {
                $Script:SIP = $ADUser.'msRTCSIP-PrimaryUserAddress'.substring(4,$ADUser.'msRTCSIP-PrimaryUserAddress'.Length-4)
            }
            Build-SDAPUserInfoForm
            Add-ActionBoxes
            Publish-Form

            If ($Global:Result -eq "OK")
            {
                write-host "Modify SMTP Address has started " -ForegroundColor Magenta
                $LineToWrite = "STAR" + "`t" + $FileName + " script has started"
                WriteLogEvent
                $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
                WriteReportEvent

                If ($Script:chkNewPrimary.Checked -eq "Checked")
                {
                    write-host "New Primary Address"
                    $LineToWrite = "INFO" + "`t" + "           Ticket Number: " + $Script:txtTicketNo.Text
                    WriteReportEvent
                    $LineToWrite = "INFO" + "`t" + "     New Primary Address: " + $Script:txtNewAddr.Text + "`n"
                    WriteReportEvent
                    $LineToWrite = "INFO" + "`t" + "Account Details Prior to Change"
                    WriteReportEvent

                    Address-Details
                    Address-Exists

                    If ($Script:MatchFound -eq "N")
                    {
                        $NewPrimAddr = "SMTP:" + $Script:txtNewAddr.Text
                        $LineToWrite = "ADD " + "`t" + "          Adding Address: " + $NewPrimAddr + "`n"
                        WriteReportEvent

                        #Demote Current Primary Address
                        $PrimAddr = "SMTP:" + $ADUser.EmailAddress
                        $DemoteAddr = "smtp:" + $ADUser.EmailAddress
                        Set-AdUser -Identity $Global:txtInpEmpNo.Text -Remove @{ProxyAddresses=$PrimAddr}
                        Set-ADUser -Identity $Global:txtInpEmpNo.Text -Add @{ProxyAddresses=$DemoteAddr}
                        #Add New Primary Address and Change SIP Address
                        $NewPrimAddr = "SMTP:" + $Script:txtNewAddr.Text
                        $SIPAddr = "sip:" + $Script:txtNewAddr.Text
                        Set-ADUser -Identity $Global:txtInpEmpNo.Text -Add @{ProxyAddresses=$NewPrimAddr}
                        Set-ADUser -Identity $Global:txtInpEmpNo.Text -Replace @{targetAddress=$NewPrimAddr}
                        Set-ADUser -Identity $Global:txtInpEmpNo.Text -Replace @{mail=$Script:txtNewAddr.Text}
                        Set-ADUser -Identity $Global:txtInpEmpNo.Text -Replace @{'msRTCSIP-PrimaryUserAddress'=$SIPAddr}
                        Start-Sleep -Seconds 20
                        $ADUser = Get-ADUser $Global:txtInpEmpNo.Text -Properties *
                        $LineToWrite = "INFO" + "`t" + "Account Details After Change"
                        WriteReportEvent
                        Address-Details
                    }
                    else
                    {
                        $LineToWrite = "FAIL" + "`t" + "Requested Address " + $FindAddr + " is available."
                        WriteReportEvent     
                    }
                    $LineToWrite = "INFO" + "`t" + "Addition of New Primary Address Complete"
                    WriteReportEvent

                }

                If ($Script:chkAddProxy.Checked -eq "Checked")
                {
                    write-host "Add Additional Proxy Address"
                    $LineToWrite = "INFO" + "`t" + "           Ticket Number: " + $Script:txtTicketNo.Text
                    WriteReportEvent
                    $LineToWrite = "INFO" + "`t" + "    Proxy Address to Add: " + $Script:txtNewAddr.Text + "`n"
                    WriteReportEvent
                    $LineToWrite = "INFO" + "`t" + "Account Details Prior to Change"
                    WriteReportEvent

                    Address-Details
                    Address-Exists

                    If ($Script:MatchFound -eq "N")
                    {
                        $AddlAddr = "smtp:" + $Script:txtNewAddr.Text
                        $LineToWrite = "ADD " + "`t" + "          Adding Address: " + $AddlAddr + "`n"
                        WriteReportEvent
                        Set-AdUser -Identity $Global:txtInpEmpNo.Text -Add @{ProxyAddresses=$AddlAddr}
                        Start-Sleep -Seconds 20
                        $ADUser = Get-ADUser $Global:txtInpEmpNo.Text -Properties *
                        $LineToWrite = "INFO" + "`t" + "Account Details After Change"
                        WriteReportEvent
                        Address-Details
                    }
                    $LineToWrite = "INFO" + "`t" + "Addition of Proxy Address Complete"
                    WriteReportEvent
                }

                If ($Script:chkPromoProxy.Checked -eq "Checked")
                {
                    write-host "Promote Proxy Address"
                    $LineToWrite = "INFO" + "`t" + "           Ticket Number: " + $Script:txtTicketNo.Text
                    WriteReportEvent
                    $LineToWrite = "INFO" + "`t" + "      Address to Promote: " + $Script:txtNewAddr.Text + "`n"
                    WriteReportEvent
                    $LineToWrite = "INFO" + "`t" + "Account Details Prior to Change"
                    WriteReportEvent
                    Address-Details

                    $PromoProxyAddr = "SMTP:" + $Script:txtNewAddr.Text
                    $RemProxyAddr = "smtp:" + $Script:txtNewAddr.Text
                    $PrimAddr = "SMTP:" + $ADUser.EmailAddress
                    $DemoteAddr = "smtp:" + $ADUser.EmailAddress
                    $SIPAddr = "sip:" + $Script:txtNewAddr.Text

                    #Demote Current Primary Email Address
                    Set-AdUser -Identity $Global:txtInpEmpNo.Text -Remove @{ProxyAddresses=$PrimAddr}
                    Set-ADUser -Identity $Global:txtInpEmpNo.Text -Add @{ProxyAddresses=$DemoteAddr}
                    #Add New Primary Address and Change SIP Address
                    Set-AdUser -Identity $Global:txtInpEmpNo.Text -Remove @{ProxyAddresses=$RemProxyAddr}
                    Set-ADUser -Identity $Global:txtInpEmpNo.Text -Add @{ProxyAddresses=$PromoProxyAddr}
                    Set-ADUser -Identity $Global:txtInpEmpNo.Text -Replace @{targetAddress=$PromoProxyAddr}
                    Set-ADUser -Identity $Global:txtInpEmpNo.Text -Replace @{mail=$Script:txtNewAddr.Text}
                    Set-ADUser -Identity $Global:txtInpEmpNo.Text -Replace @{'msRTCSIP-PrimaryUserAddress'=$SIPAddr}
                    Start-Sleep -Seconds 20
                    $ADUser = Get-ADUser $Global:txtInpEmpNo.Text -Properties *
                    $LineToWrite = "INFO" + "`t" + "Account Details After Change"
                    WriteReportEvent
                    Address-Details
                    $LineToWrite = "INFO" + "`t" + "Address Promotion Process Complete"
                    WriteReportEvent
                }

                If ($Script:chkRemoveAddr.Checked -eq "Checked")
                {
                    write-host "Remove Address Address"
                    $LineToWrite = "INFO" + "`t" + "           Ticket Number: " + $Script:txtTicketNo.Text
                    WriteReportEvent
                    $LineToWrite = "INFO" + "`t" + "       Address to Remove: " + $Script:txtNewAddr.Text + "`n"
                    WriteReportEvent
                    $LineToWrite = "INFO" + "`t" + "Account Details Prior to Change"
                    WriteReportEvent

                    Address-Details
                    $RemAddr = "smtp:" + $Script:txtNewAddr.Text
                    Set-AdUser -Identity $Global:txtInpEmpNo.Text -Remove @{ProxyAddresses=$RemAddr}
                    Start-Sleep -Seconds 20
                    $ADUser = Get-ADUser $Global:txtInpEmpNo.Text -Properties *
                    $LineToWrite = "INFO" + "`t" + "Account Details After Change"
                    WriteReportEvent
                    Address-Details
                    $LineToWrite = "INFO" + "`t" + "Removal of Proxy Address Complete"
                    WriteReportEvent
                }

                If ($Script:chkChgSIP.Checked -eq "Checked")
                {
                    write-host "Update SIP Address"
                    $LineToWrite = "INFO" + "`t" + "           Ticket Number: " + $Script:txtTicketNo.Text
                    WriteReportEvent
                    $LineToWrite = "INFO" + "`t" + "     Current SIP Address: " + $ADUser.'msRTCSIP-PrimaryUserAddress' + "`n"
                    WriteReportEvent
                    $LineToWrite = "INFO" + "`t" + "Account Details Prior to Change"
                    WriteReportEvent

                    Address-Details
                    $SIPAddr = "sip:" + $Script:txtNewAddr.Text
                    Set-ADUser -Identity $Global:txtInpEmpNo.Text -Replace @{'msRTCSIP-PrimaryUserAddress'=$SIPAddr}
                    Start-Sleep -Seconds 20
                    $ADUser = Get-ADUser $Global:txtInpEmpNo.Text -Properties *
                    $LineToWrite = "INFO" + "`t" + "Account Details After Change"
                    WriteReportEvent
                    Address-Details
                    $LineToWrite = "INFO" + "`t" + "Update of SIP Address Complete"
                    WriteReportEvent

                    $Script:chkChgSIP.Visible = $false
                }
                Form-PostChange
            }
            else
            {
                $Output = $wshell.Popup("Request Cancelled.",0,"Cancelled",0+32)  
            }
        }
        else
        {
            $Output = $wshell.Popup("User Account Not Found.",0,"Not Found",0+32)  
        }
    }
}