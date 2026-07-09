#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange
#    'Description  : Manage Phone Extensions for A User Account
#    'Called By    : SDAPAdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 06/25/2021
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 06/25/2021 Created from the scripts provided by Cristian Rauth (TeamsNewHire.ps1,myClasses.ps1,mySQL.ps1)
#    '             : 10/13/2021 Released script to SDAP Team
#    '             : 10/19/2021 Changed so if account is disabled it clears the emp# and does not populate the details required to assign a phone no.
#    '	           : 12/09/2021 Changes to implemente new SQL server by Cristian Rauth
#    '             : 04/20/2022 Added check that they user has an E3 license since the Phone license is being enabled for all staff who have the E5 licenses
#    '             : 05/25/2022 Commented out lines that assign the old Phone System License as this is now part of the enabled E5 features
# ==========================================================================
#
#################################################################################

Function Build-PhoneDetails
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Manage Phone Number"
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 750 ; $form.Height = 375  # Make the form wider 
    
    $Left = 30
    $LeftInput = 160
    $LeftButton = 520
    $Top = 20
    $Tab = 0

    ## Assign Extension
    $Script:chkAssignExt = New-Object Windows.Forms.RadioButton
    $Script:chkAssignExt.Left = $LeftButton; $Script:chkAssignExt.Width = 200; $Script:chkAssignExt.Top = $Top 
    $Script:chkAssignExt.Text = "Assign Phone Number" 
    $Script:chkAssignExt.Checked = $True
    $Global:form.Controls.Add($Script:chkAssignExt)
    $Script:chkAssignExt.visible = $False

    $Top = $Top + 20
    ## Change Pool
    $Script:chkChgExt = New-Object Windows.Forms.RadioButton
    $Script:chkChgExt.Left = $LeftButton; $Script:chkChgExt.Width = 200; $Script:chkChgExt.Top = $Top 
    $Script:chkChgExt.Text = "Change Phone Number" 
    $Script:chkChgExt.Checked = $False
    $Global:form.Controls.Add($Script:chkChgExt)
    $Script:chkChgExt.visible = $False
    $Script:chkChgExt.Add_MouseClick(
        {
            $Script:lblDIDRange.Visible = $true
            $Script:locDIDRange.Visible = $true
            $Script:GetDIDButDetails.Visible = $true
            $Global:OKButton.Text = "Change"
            $Global:OKButton.visible = $True
        })

    $Top = $Top + 20
    ## Remove Extension
    $Script:chkRemoveExt = New-Object Windows.Forms.RadioButton
    $Script:chkRemoveExt.Left = $LeftButton; $Script:chkRemoveExt.Width = 150; $Script:chkRemoveExt.Top = $Top  
    $Script:chkRemoveExt.Text = "Remove Extension" 
    $Script:chkRemoveExt.Checked = $False
    $Global:form.Controls.Add($Script:chkRemoveExt)
    $Script:chkRemoveExt.Visible = $False         
    $Script:chkRemoveExt.Add_MouseClick(
        {
            $Global:OKButton.Text = "Remove"
            $Global:OKButton.visible = $True
            $Script:lblDIDRange.Visible = $False
            $Script:locDIDRange.Visible = $False
            $Script:txtSelDID.Visible = $False
            $Script:lblNewDID.Visible = $False
            $Script:txtNewDID.Visible = $False
            $Script:GetDIDButDetails.Visible = $False
        })

    $Top = 20
    ## Employee ID
    $Script:lblEmpID = New-Object System.Windows.Forms.Label   
    $Script:lblEmpID.Text = "Employee No.*:" 
    $Script:lblEmpID.Top = $Top ; $Script:lblEmpID.Left = $Left; $Script:lblEmpID.Width = 150 ; $Script:lblEmpID.AutoSize = $true 
    $Global:form.Controls.Add($Script:lblEmpID)    # Add to Form 
    $Script:txtEmpID = New-Object Windows.Forms.TextBox
    $Script:txtEmpID.Top = $Top; $Script:txtEmpID.Left = $LeftInput; $Script:txtEmpID.Width = 150;  
    $Script:txtEmpID.Text = ""
    $Global:form.Controls.Add($Script:txtEmpID)    # Add to Form
    $Global:InputFocus = $Script:txtEmpID
    $Script:ButGetDetails = New-Object Windows.Forms.Button
    $Script:ButGetDetails.Location = New-object System.Drawing.Size(($LeftInput + 160), $Top)
    $Script:ButGetDetails.Size = new-Object System.Drawing.Size(140, 20)
    $Script:ButGetDetails.Text = "Get Account Details"
    $Global:form.Controls.Add($Script:ButGetDetails)
    $Script:ButGetDetails.Add_MouseClick(
        {
            $employeeid = $Script:txtEmpID.Text
            # $evenabled = e:\O365AdminShared\Scripts\MySQL.ps1 -Query "SELECT TeamsEnterpriseVoiceEnabled,TeamsLineUri,DisplayName FROM endpoints WHERE EmployeeID LIKE '$employeeid'"

            ################################### Added by Cristian ############################################
            $evenabled = Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query "SELECT EnterpriseVoiceEnabled,LineUri,DisplayName FROM endpoints WHERE SamAccountName LIKE '$employeeid'"
            ##################################################################################################

            ## if ([string]::IsNullOrEmpty($evenabled.TeamsLineUri)) <-- Cristian: Column Name in DB Changed from TeamsLineUri into LineURI -->
            if ([string]::IsNullOrEmpty($evenabled.LineUri))
            {
                #Do Nothing
            }
            else
            {
                ## $Script:lineuri = $evenabled.TeamsLineUri  <-- Cristian: Column Name in DB Changed from TeamsLineUri into LineURI -->
                $Script:lineuri = $evenabled.LineUri
                $Script:DisplayName = $evenabled.DisplayName
                $Output = $wshell.Popup("User is already enabled for Enterprise Voice and or has a Phone Number.", 0, "Already Enabled", 0 + 32)
            }

            If ($Script:txtEmpID.Text -like "*@*")
            {
                $Script:txtEmpID.Text = (($Script:txtEmpID.Text.Substring(0, $Script:txtEmpID.Text.Indexof("@")))).Trim()
            }
            $Script:EmpNo = ($Script:txtEmpID.Text) -Replace (" ", "")
            $Script:UPN = $Script:EmpNo + "@global.ul.com"

            $IsLicensed = Get-MsolUser -UserPrincipalName $Script:UPN
            
            If (($IsLicensed.Licenses.AccountSkuId -contains "ul:ENTERPRISEPACK") -or ($IsLicensed.Licenses.AccountSkuId -contains "ul:ENTERPRISEPREMIUM"))
            {
                Get-PhoneDetails
                #                $Script:GetDIDButDetails.Visible = $True
            }
            else
            {
                $Output = $wshell.Popup("User " + $Script:UPN + " does not have an O365 E3 license assigned, enable account for email before assigning a phone number.", 0, "No Licenses", 0 + 32)
            }
        })

    $Top = $Top + 30
    ## Ticket No.
    $Script:lblTicketNo = New-Object System.Windows.Forms.Label   
    $Script:lblTicketNo.Text = "Ticket Number*:" 
    $Script:lblTicketNo.Top = $Top ; $Script:lblTicketNo.Left = $Left; $Script:lblTicketNo.Width = 150 ; $Script:lblTicketNo.AutoSize = $true 
    $Global:form.Controls.Add($Script:lblTicketNo)    # Add to Form 
    $Script:txtTicketNo = New-Object Windows.Forms.TextBox
    $Script:txtTicketNo.Top = $Top; $Script:txtTicketNo.Left = $LeftInput; $Script:txtTicketNo.Width = 150; $Script:txtTicketNo.ReadOnly = $true;
    $Script:txtTicketNo.Text = ""
    $Global:form.Controls.Add($Script:txtTicketNo)    # Add to Form

    $Top = $Top + 30
    ## Employee Name
    $Script:lblEmpName = New-Object System.Windows.Forms.Label   
    $Script:lblEmpName.Text = "Employee Name:" 
    $Script:lblEmpName.Top = $Top ; $Script:lblEmpName.Left = $Left; $Script:lblEmpName.Width = 150 ; $Script:lblEmpName.AutoSize = $true
    $Global:form.Controls.Add($Script:lblEmpName)    # Add to Form 
    $Script:txtEmpName = New-Object Windows.Forms.TextBox
    $Script:txtEmpName.Top = $Top; $Script:txtEmpName.Left = $LeftInput; $Script:txtEmpName.Width = 300; $Script:txtEmpName.ReadOnly = $true;
    $Script:txtEmpName.Text = ""
    $Global:form.Controls.Add($Script:txtEmpName)    # Add to Form
        
    $Top = $Top + 30
    ## Employee Typ
    $Script:lblEmpType = New-Object System.Windows.Forms.Label   
    $Script:lblEmpType.Text = "Employee Type:" 
    $Script:lblEmpType.Top = $Top ; $Script:lblEmpType.Left = $Left; $Script:lblEmpType.Width = 150 ; $Script:lblEmpType.AutoSize = $true
    $Global:form.Controls.Add($Script:lblEmpType)    # Add to Form 
    $Script:txtEmpType = New-Object Windows.Forms.TextBox
    $Script:txtEmpType.Top = $Top; $Script:txtEmpType.Left = $LeftInput; $Script:txtEmpType.Width = 150; $Script:txtEmpType.ReadOnly = $true;
    $Script:txtEmpType.Text = ""
    $Global:form.Controls.Add($Script:txtEmpType)    # Add to Form

    $Top = $Top + 30
    ## Office Location
    $Script:lblOffLoc = New-Object System.Windows.Forms.Label   
    $Script:lblOffLoc.Text = "Office Location:" 
    $Script:lblOffLoc.Top = $Top ; $Script:lblOffLoc.Left = $Left; $Script:lblOffLoc.Width = 150 ; $Script:lblOffLoc.AutoSize = $true
    $Global:form.Controls.Add($Script:lblOffLoc)    # Add to Form 
    $Script:txtOffLoc = New-Object Windows.Forms.TextBox
    $Script:txtOffLoc.Top = $Top; $Script:txtOffLoc.Left = $LeftInput; $Script:txtOffLoc.Width = 300; $Script:txtOffLoc.ReadOnly = $true;
    $Script:txtOffLoc.Text = ""
    $Global:form.Controls.Add($Script:txtOffLoc)    # Add to Form

    $Top = $Top + 30
    ## SIP Address
    $Script:lblSIPAddr = New-Object System.Windows.Forms.Label   
    $Script:lblSIPAddr.Text = "SIP Address:" 
    $Script:lblSIPAddr.Top = $Top ; $Script:lblSIPAddr.Left = $Left; $Script:lblSIPAddr.Width = 150 ; $Script:lblSIPAddr.AutoSize = $true
    $Global:form.Controls.Add($Script:lblSIPAddr)    # Add to Form 
    $Script:txtSIPAddr = New-Object Windows.Forms.TextBox
    $Script:txtSIPAddr.Top = $Top; $Script:txtSIPAddr.Left = $LeftInput; $Script:txtSIPAddr.Width = 300; $Script:txtSIPAddr.ReadOnly = $true;
    $Script:txtSIPAddr.Text = ""
    $Global:form.Controls.Add($Script:txtSIPAddr)    # Add to Form

    $Top = $Top + 30
    ## Current Extension
    $Script:lblCurrExt = New-Object System.Windows.Forms.Label   
    $Script:lblCurrExt.Text = "Phone Number:" 
    $Script:lblCurrExt.Top = $Top ; $Script:lblCurrExt.Left = $Left; $Script:lblCurrExt.Width = 150 ; $Script:lblCurrExt.AutoSize = $true 
    $Global:form.Controls.Add($Script:lblCurrExt)    # Add to Form 
    $Script:txtCurrExt = New-Object Windows.Forms.TextBox
    $Script:txtCurrExt.Top = $Top; $Script:txtCurrExt.Left = $LeftInput; $Script:txtCurrExt.Width = 300; $Script:txtCurrExt.ReadOnly = $true; 
    $Script:txtCurrExt.Text = ""
    $Global:form.Controls.Add($Script:txtCurrExt)    # Add to Form

    $Top = $Top + 30
    ## DID Ranges
    $Script:lblDIDRange = New-Object System.Windows.Forms.Label   
    $Script:lblDIDRange.Text = "DID Ranges*:" 
    $Script:lblDIDRange.Top = $Top ; $Script:lblDIDRange.Left = $Left; $Script:lblDIDRange.Width = 150 ; $Script:lblDIDRange.AutoSize = $true 
    $Global:form.Controls.Add($Script:lblDIDRange)    # Add to Form 
    $Script:locDIDRange = New-Object System.Windows.Forms.ListBox
    $Script:locDIDRange.Top = $Top; $Script:locDIDRange.Left = $LeftInput; $Script:locDIDRange.Height = 70; $Script:LocDIDRange.Width = 540;
    $Global:form.Controls.Add($Script:locDIDRange) #Add listbox to form
    $Script:txtSelDID = New-Object Windows.Forms.TextBox
    $Script:txtSelDID.Top = $Top; $Script:txtSelDID.Left = $LeftInput; $Script:txtSelDID.Width = 300; $Script:txtSelDID.ReadOnly = $true ; $Script:txtSelDID.Visible = $False
    $Script:txtSelDID.Text = ""
    $Global:form.Controls.Add($Script:txtSelDID)    # Add to Form
    $Script:GetDIDButDetails = New-Object Windows.Forms.Button
    $Script:GetDIDButDetails.Location = New-object System.Drawing.Size($Left, ($Top + 20))
    $Script:GetDIDButDetails.Size = new-Object System.Drawing.Size(100, 20)
    $Script:GetDIDButDetails.Text = "Get New DID"
    $Global:form.Controls.Add($Script:GetDIDButDetails)
    $Script:lblDIDRange.Visible = $False
    $Script:locDIDRange.Visible = $False
    $Script:GetDIDButDetails.Visible = $False
    $Script:GetDIDButDetails.Add_MouseClick(
        {
            If ($Script:locDIDRange.SelectedItem.Length -gt 0)
            {
                $Selection = ($Script:locDIDRange.SelectedItem).substring(0, 1)
                Get-DIDFinder
                #                $Global:OKButton.Text = "Assign"
            }
            else
            {
                $Output = $wshell.Popup("Must select a DID Range to continue.", 0, "Select DID Range", 0 + 32)
            }
        })

    $Top = $Top + 30
    ## New Suggested DID
    $Script:lblNewDID = New-Object System.Windows.Forms.Label   
    $Script:lblNewDID.Text = "New Suggested DID:"  
    $Script:lblNewDID.Top = $Top ; $Script:lblNewDID.Left = $Left; $Script:lblNewDID.Width = 100 ; $Script:lblNewDID.AutoSize = $true 
    $Global:form.Controls.Add($Script:lblNewDID)    # Add to Form 
    $Script:txtNewDID = New-Object Windows.Forms.TextBox
    $Script:txtNewDID.Top = $Top; $Script:txtNewDID.Left = $LeftInput; $Script:txtNewDID.Width = 300; $Script:txtNewDID.ReadOnly = $true
    $Script:txtNewDID.Text = ""
    $Global:form.Controls.Add($Script:txtNewDID)    # Add to Form
    $Script:lblNewDID.Visible = $False
    $Script:txtNewDID.Visible = $False

    $Top = $Top + 50
    ##ReportDetails
    $Script:lblRptFile = New-Object System.Windows.Forms.Label   
    $Script:lblRptFile.Text = "Report Details:"  
    $Script:lblRptFile.Top = $Top ; $Script:lblRptFile.Left = $Left; $Script:lblRptFile.Width = 100 ; $Script:lblRptFile.AutoSize = $true 
    $Global:form.Controls.Add($Script:lblRptFile)    # Add to Form 
    # 
    $Script:txtRptFile = New-Object Windows.Forms.TextBox
    $Script:txtRptFile.ReadOnly = $true; 
    $Script:txtRptFile.TabIndex = $Tab++ # set Tab Order
    $Script:txtRptFile.Top = $Top; $Script:txtRptFile.Left = $LeftInput; $Script:txtRptFile.Width = 540; 
    $Script:txtRptFile.Text = ""
    $Global:form.Controls.Add($Script:txtRptFile)    # Add to Form
    $Script:lblRptFile.Visible = $false
    $Script:txtRptFile.Visible = $false
}

Function Get-PhoneDetails
{
    $Usr = ($Script:txtEmpID.Text) -Replace (" ", "")
    $ADExists = [bool](get-ADUser -Filter { SamAccountName -eq $Usr } -ErrorAction SilentlyContinue)

    If ($ADExists -eq $False)
    {
        $Output = $wshell.Popup("Account not found for Emp ID: " + $Usr, 0, "No Account Found", 0 + 32)
    }
    else
    {
        $Global:form.Height = 475
        $Script:ADUsr = Get-ADUser $Usr -Properties *
        If ($Script:ADUsr.Enabled -eq $true)
        {
            $Script:txtEmpName.Text = $Script:ADUsr.cn
            $Script:txtEmpType.Text = $Script:ADUsr.extensionAttribute1
            $Script:txtOffLoc.Text = $Script:ADUsr.PhysicalDeliveryOfficeName
            $Script:txtSIPAddr.Text = $Script:ADUsr.'msRTCSIP-PrimaryUserAddress'
            $Script:txtTicketNo.ReadOnly = $False
            $Script:txtCurrExt.Text = $Script:ADUsr.'msRTCSIP-Line'
            $Script:txtRptFile.Text = "e:\SDAP\ManagePhone\Report\Report-ManagePhone" + "-EmpNo" + $Script:txtEmpID.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            $Script:txtEmpID.ReadOnly = $true
            $Script:ButGetDetails.Visible = $False
            If ($Script:ADUsr.'msRTCSIP-Line' -ne $null)
            {
                $Script:chkRemoveExt.Visible = $True
                $Script:chkChgExt.visible = $True
            }
            else
            {
                $Script:chkAssignExt.visible = $True
                $Script:lblDIDRange.Visible = $true
                $Script:locDIDRange.Visible = $true
                $Script:GetDIDButDetails.Visible = $true
                $Global:OKButton.Text = "Assign"
            }
            $Script:lblRptFile.Visible = $true
            $Script:txtRptFile.Visible = $true

            ##Set these with the form variables
            $office = $Script:txtOffLoc.Text
            $sipaddress = $Script:ADUsr.Mail

            if ($office)
            {
                Get-SfBHelperOfficeDetails
            }

            if ($didranges)
            {
                Get-NewDIDMenu
                If ($Script:ADUsr.'msRTCSIP-Line' -ne $null)
                {
                    $Script:GetDIDButDetails.Visible = $false
                }
            }
        }
        else
        {
            $Output = $wshell.Popup("Account disabled for Emp ID: " + $Usr, 0, "Account Disabled", 0 + 32)
            $Script:txtEmpID.Text = ""
        }
    }
}

Function Get-SfBHelperOfficeDetails
{
    #### Look for Office and Settings ####
    $Office = $Script:txtOffLoc.Text
    # $policy = e:\O365AdminShared\Scripts\MySQL.ps1 -Query "SELECT LocationCode FROM locationconfiguration WHERE Name LIKE '$office'"

    ############################# Added By Cristian ################################################
    $policy = Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query "SELECT LocationCode FROM locationconfiguration WHERE Name LIKE '$office'"
    ################################################################################################

    $policy = $policy.LocationCode
    if ($policy)
    {
        ## $didranges = e:\O365AdminShared\Scripts\MySQL.ps1 -Query "SELECT DIDSTART,DIDEND,Utilization,Notes FROM did WHERE (LocationCode LIKE '$policy') AND (SDAP LIKE '1')"
        ################## Added By Cristian ############################################
        $didranges = Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query "SELECT DIDSTART,DIDEND,Utilization,Notes FROM did WHERE (LocationCode LIKE '$policy') AND (SDAP LIKE '1')"
        #################################################################################

        if ($didranges)
        {
        }
        else
        {
            $Output = $wshell.Popup("No DID Ranges found for $[$office]", 0, "No DID Ranges", 0 + 32)
            Write-Host "No DID Ranges found" -ForegroundColor Red
        }
    }
    else
    {
        Write-Host $Output = $wshell.Popup("Office $[$office] not found", 0, "Office Not Found", 0 + 32)
        Remove-Variable * -ErrorAction SilentlyContinue
    }
}

Function Get-NewDIDMenu
{
    $index = 1
    foreach ($did in $didranges)
    {
        $rangesize = $did.DIDEND - $did.DIDSTART + 1
        $Script:locDIDRange.Items.Add([string]$index + ": " + [string]$did.DIDSTART + "  Size: " + $RangeSize + "  Usage: " + [string]$did.utilization + "%  (Notes: " + [string]$did.Notes + ")")
        $index++
    }
}

Function Get-DIDFinder
{
    if ($selection)
    {
        if ($didranges.DIDSTART.Count -gt '1')
        {
            $didstart = $didranges[$selection - 1].DIDSTART
            $didend = $didranges[$selection - 1].DIDEND
            $didwork = $didstart - 1
        }
        else
        {
            $didstart = $didranges.DIDSTART
            $didend = $didranges.DIDEND
        }
       
        $counter = 1
        $didwork = $didstart - 1
        $number1 = $didstart
        $number2 = $didend
        [int64]$worknumber = $number1            ## [int64] Added by Cristian. Datatype returned by SQL is string. ++ does only work with type int
        $tabledid = $number1 -replace ".{5}$"
        $tabledid = $tabledid + '%'

        $numberrange = @()
        do
        {
            $numberrange += $worknumber
            $worknumber++
        } while ($worknumber -le $number2)

        $numberrange = $numberrange | Sort-Object { Get-Random }
        $didcount = $numberrange.Count
        $didcounter = 1
        ## $endpointstable = e:\O365AdminShared\Scripts\MySQL.ps1 -Query "SELECT DID FROM endpoints WHERE DID LIKE '$tabledid'"
        ## $endpointstable += e:\O365AdminShared\Scripts\MySQL.ps1 -Query "SELECT DID FROM blockeddids WHERE did LIKE '$tabledid'"

        ################################## Added By Cristian #############################################
        $query1 = "SELECT DID FROM endpoints WHERE DID LIKE '$tabledid'"
        $query2 = "SELECT DID FROM blockeddids WHERE did LIKE '$tabledid'"

        $endpointstable = Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query $query1
        $endpointstable += Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query $query2


        ##################################################################################################

        foreach ($number in $numberrange)
        {
            $etstatus = ($endpointstable.DID -contains $number)
            if ($etstatus -eq $false) { break }
            $didcounter++
        }

        $didwork = $number
        $didcounter
        if (($didcounter) -gt $didcount)
        {
            $Output = $wshell.Popup("DID Range Exhausted, please choose a different range or contact the Telecom team to order more numbers", 0, "DID Range Exhausted", 0 + 32)
            $didwork = $null
        }
        else
        {
            $Script:locDIDRange.Visible = $False
            $Script:lblDIDRange.Text = "Selected DID Range"
            $Script:txtSelDID.Text = $Script:locDIDRange.SelectedItem; $Script:txtSelDID.Visible = $True
            $Script:GetDIDButDetails.Visible = $False
            $Script:lblNewDID.Visible = $true
            $Script:txtNewDID.Text = "+$didwork"
            $Script:txtNewDID.Visible = $true
            $Global:OKButton.visible = $true
        }
    }
}

function Set-LineURI
{
    $Script:lineuri = "tel:" + $Script:txtNewDID.Text
    $newdid = $Script:txtNewDID.Text.Substring(1, $Script:txtNewDID.Text.Length - 1)
               
    if (([string]::IsNullOrEmpty($evenabled)) -or ($evenenabled.length -eq 0))
    {
        Write-Host "`nEnabling Phone Features for: "$Script:txtEmpID.Text
        $LineToWrite = "INFO" + "`tTicket Number                            : " + $Script:txtTicketNo.Text
        WriteReportEvent
        $LineToWrite = "INFO" + "`tEnabling Phone Features for              : " + $Script:txtEmpID.Text
        WriteReportEvent

        $office2 = $Script:txtOffLoc.Text
        ## $policy2 = (e:\O365AdminShared\Scripts\MySQL.ps1 -Query "SELECT OfficeCode FROM servicenowofficecodes WHERE ADOfficeCode = '$office2'").OfficeCode

        ############################################## Added By Cristian #######################################
        
        $policy2 = (Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query "SELECT LocationCode FROM locationconfiguration WHERE Name = '$office'").LocationCode

        ########################################################################################################
 
        $Script:ADUsr = Get-ADuser -Properties * -Filter { (SAmAccountName -eq $EmpNo) } | Select DisplayName, UserPrincipalName, Mail, msRTCSIP-PrimaryUserAddress, msRTCSIP-DeploymentLocator, msRTCSIP-FederationEnabled, msRTCSIP-InternetAccessEnabled, msRTCSIP-Line, msRTCSIP-PrimaryHomeServer, msRTCSIP-UserEnabled, msRTCSIP-UserPolicies, msRTCSIP-UserPolicy, UserAccountControl, PhysicalDeliveryOfficeName, co
        $sipaddress = "sip:" + $Script:ADUsr.Mail
        Set-ADUser $Script:EmpNo -Replace @{'msRTCSIP-DeploymentLocator' = "sipfed.online.lync.com" }
        Set-ADUser $Script:EmpNo -Replace @{'msRTCSIP-FederationEnabled' = "TRUE" }
        Set-ADUser $Script:EmpNo -Replace @{'msRTCSIP-InternetAccessEnabled' = "TRUE" }
        Set-ADUser $Script:EmpNo -Replace @{'msRTCSIP-UserEnabled' = "TRUE" }
        Set-ADUser $Script:EmpNo -Replace @{'msRTCSIP-Line' = $Script:lineuri }
        Set-ADUser $Script:EmpNo -Replace @{'msRTCSIP-PrimaryUserAddress' = $sipaddress }
#        Set-MsolUserLicense -UserPrincipalName $Script:UPN -AddLicense "ul:MCOEV" -ErrorAction SilentlyContinue

        $LineToWrite = "SET " + "`tSetting DeploymentLocator                : " + "sipfed.online.lync.com"
        WriteReportEvent
        $LineToWrite = "ENAB" + "`tEnabling Federation"
        WriteReportEvent
        $LineToWrite = "ENAB" + "`tEnabling Internet Access"
        WriteReportEvent
        $LineToWrite = "SET " + "`tSetting Phone Number                     : " + $Script:lineuri
        WriteReportEvent       
        $LineToWrite = "SET " + "`tSetting SIP Address                      : " + $sipaddress
        WriteReportEvent       
        
        Write-Host "Granting Online Voice Routing Policy ... " -ForegroundColor Cyan -NoNewline
        Grant-CsOnlineVoiceRoutingPolicy $Script:UPN -PolicyName $policy2 -ErrorAction SilentlyContinue
        $LineToWrite = "SET " + "`tSetting Voice Routing Policy             : " + $policy2
        WriteReportEvent 
        if ($?)
        {
            Write-Host "successful" -ForegroundColor Green
            $LineToWrite = "ASGN" + "`tOnline Voice Routing Policy Assigned"
            $status | Add-Member -MemberType NoteProperty -Name VRP -Value "yes"
        }
        else
        {
            $LineToWrite = "ERR " + "`tError Assigning Online Voice Routing Policy"
            Write-Host "failed" -ForegroundColor Red
        }
        WriteReportEvent 
        
        Write-Host "Granting Tenant Dial Plan ... " -ForegroundColor Cyan -NoNewline
        Grant-CsTenantDialPlan $Script:ADUsr.Mail -PolicyName $policy2 -ErrorAction SilentlyContinue
        #        Grant-CsTenantDialPlan -PolicyName $policy2 -Identity (Get-CSOnlineUser $Script:UPN).sipaddress -ErrorAction SilentlyContinue
        if ($?)
        {
            Write-Host "successful" -ForegroundColor Green
            $LineToWrite = "ASGN" + "`tTenant Dial Plan Assigned"
            $status | Add-Member -MemberType NoteProperty -Name TDP -Value "yes"
        }
        else
        {
            Write-Host "previously enabled" -ForegroundColor Red
            $LineToWrite = "ERR " + "`tError Assigning Tenant Dial Plan may have been previously enabled"
        }

        WriteReportEvent 
        
        Write-Host "Granting Emergency Call Routing Policy ... " -ForegroundColor Cyan -NoNewline
        Grant-CsTeamsEmergencyCallRoutingPolicy $Script:ADUsr.Mail -PolicyName $policy2 -ErrorAction SilentlyContinue
        #        Grant-CsTeamsEmergencyCallRoutingPolicy -Identity (Get-CSOnlineUser $Script:UPN).sipaddress -PolicyName $policy2 -ErrorAction SilentlyContinue
        if ($?)
        {
            Write-Host "successful" -ForegroundColor Green
            $LineToWrite = "ASGN" + "`tEmergency Call Routing Policy Assigned"
            $status | Add-Member -MemberType NoteProperty -Name ECR -Value "yes"
        }

        else
        {
            Write-Host "previously enabled" -ForegroundColor Red
            $LineToWrite = "ERR " + "`tError Assigning Emergency Call Routing Policy may have been previously enabled"
        }

        WriteReportEvent

        Write-Host "Granting Emergency Calling Policy ... " -ForegroundColor Cyan -NoNewline
        Grant-CsTeamsEmergencyCallingPolicy $Script:ADUsr.Mail -PolicyName $policy2 -ErrorAction SilentlyContinue
        #        Grant-CsTeamsEmergencyCallingPolicy -Identity (Get-CSOnlineUser $Script:UPN).sipaddress -PolicyName $policy2 -ErrorAction SilentlyContinue
        if ($?)
        {
            Write-Host "successful" -ForegroundColor Green
            $LineToWrite = "ASGN" + "`tEmergency Calling Policy Assigned"
            $status | Add-Member -MemberType NoteProperty -Name ECP -Value "yes"
        }
        else
        {
            Write-Host "previuosly enabled" -ForegroundColor Red
            $LineToWrite = "ERR " + "`tError Assigning Emergency Calling Policy may have been previously enabled"
        }
        WriteReportEvent 
       
        Write-Host "Granting Teams Only Mode ... " -ForegroundColor Cyan -NoNewline
        Grant-CsTeamsUpgradePolicy $Script:ADUsr.Mail -PolicyName "UpgradeToTeams" -ErrorAction SilentlyContinue
        #        Grant-CsTeamsUpgradePolicy -Identity (Get-CSOnlineUser $Script:UPN).sipaddress -PolicyName "UpgradeToTeams" -ErrorAction SilentlyContinue
        if ($?)
        {
            Write-Host "successful" -ForegroundColor Green
            $LineToWrite = "ASGN" + "`tTeams Only Mode Assigned"
            $status | Add-Member -MemberType NoteProperty -Name TUP -Value "yes"
            $tomstatus = $policy2
        }
        else
        {
            Write-Host "previously enabled" -ForegroundColor Red
            $LineToWrite = "ERR " + "`tError Assigning Teams Only Mode may have been previously enabled"
        }
        WriteReportEvent
        
        Set-CsUser $Script:UPN -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -ErrorAction SilentlyContinue
        $LineToWrite = "ASGN" + "`tEnabled Enterprise Voice and Hosted Voicemail"
        WriteReportEvent 

        Write-Host "Setting DID Number for "$Script:ADUsr.DisplayName" to: "$Script:txtNewDID.Text "..." -ForegroundColor Cyan
        Set-CsUser $Script:UPN -LineURI $Script:lineuri -ErrorAction SilentlyContinue
        if ($?)
        {
            Write-Host "Reserving DID number ..." -ForegroundColor Cyan
            $status | Add-Member -MemberType NoteProperty -Name lineuri -Value "yes"

            ### Reserve DID in SfB Helper ###
            ## e:\O365AdminShared\Scripts\MySQL.ps1 -Query "INSERT INTO endpoints (DID) VALUE ('$newdid')"

            ######################### Added By Cristian #######################e
            Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query "INSERT INTO endpoints (DID) VALUES ('$newdid')"
        }
    }
    else
    {
        Write-Host "User not enabled. Do manually first ..." -ForegroundColor Red
    }
}

function Change-LineURI
{
    $Script:lineuri = "tel:" + $Script:txtNewDID.Text
    $newdid = $Script:txtNewDID.Text.Substring(1, $Script:txtNewDID.Text.Length - 1)
               
    if (([string]::IsNullOrEmpty($evenabled)) -or ($evenenabled.length -eq 0))
    {
        Write-Host "Change Phone Number for: "$Script:txtEmpID.Text
        $LineToWrite = "INFO" + "`tTicket Number                            : " + $Script:txtTicketNo.Text
        WriteReportEvent
        $LineToWrite = "INFO" + "`tChanging Phone Number for                : " + $Script:txtEmpID.Text
        WriteReportEvent

        ## $Rec = E:\O365AdminShared\Scripts\MYSQL.PS1 -Query "SELECT * FROM endpoints WHERE EmployeeID LIKE '$Script:EmpNo'"

        #################################### Added By Cristian ###################################
        $query = "SELECT * FROM endpoints WHERE SamAccountName = '" + $Script:EmpNo + "'"
        Write-Host $query
        Start-Sleep -s 10
        $Rec = Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query $query
        $UsrSQLRec = $Rec.ID
        ##########################################################################################
        ## $UsrSQLRec = $Rec.EndpointID
        $office2 = $Script:txtOffLoc.Text
        ## $policy2 = (e:\O365AdminShared\Scripts\MySQL.ps1 -Query "SELECT OfficeCode FROM servicenowofficecodes WHERE ADOfficeCode = '$office2'").OfficeCode
        
        #################################### Added By Cristian ###################################
        $policy2 = (Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query "SELECT LocationCode FROM locationconfiguration WHERE Name = '$office'").LocationCode
        ##########################################################################################

        Write-Host "Change Phone Number from: "$Rec.DID" to: "$newdid
        $Script:ADUsr = Get-ADuser -Properties * -Filter { (SAmAccountName -eq $EmpNo) } | Select DisplayName, UserPrincipalName, Mail, msRTCSIP-PrimaryUserAddress, msRTCSIP-DeploymentLocator, msRTCSIP-FederationEnabled, msRTCSIP-InternetAccessEnabled, msRTCSIP-Line, msRTCSIP-PrimaryHomeServer, msRTCSIP-UserEnabled, msRTCSIP-UserPolicies, msRTCSIP-UserPolicy, UserAccountControl, PhysicalDeliveryOfficeName, co
        $sipaddress = "sip:" + $Script:ADUsr.Mail
        Set-ADUser $Script:EmpNo -Replace @{'msRTCSIP-Line' = $Script:lineuri }

        $LineToWrite = "INFO" + "`tOld DID Number                           : " + $Rec.DID
        WriteReportEvent
        $LineToWrite = "INFO" + "`tOld TeamsLineURI Number                  : " + $Rec.TeamsLineURI
        WriteReportEvent
        $LineToWrite = "INGO" + "`tOld Voice Policies                       : " + $Rec.TenantDialPlan
        WriteReportEvent
        $LineToWrite = "INFO" + "`tNew DID Number                           : " + $newdid
        WriteReportEvent 
        $LineToWrite = "INFO" + "`tNew TeamsLineURI                         : " + $Script:lineuri
        WriteReportEvent  
        $LineToWrite = "INFO" + "`tNew Voice Policies                       : " + $policy2
        WriteReportEvent

        write-Host "Resetting DID Number in SQL Database" -ForegroundColor Cyan
        Write-Host "Resetting DID Number for "$Script:ADUsr.DisplayName" from: "$Rec.DID" to: "$Script:txtNewDID.Text "..." -ForegroundColor Cyan
        Set-CsUser $Script:UPN -LineURI $Script:lineuri -ErrorAction SilentlyContinue
        ## e:\O365AdminShared\Scripts\MySQL.ps1 -Query "UPDATE endpoints SET DID = '$newdid',TeamsLineURI = '$LineURI' WHERE EndpointID  = '$UsrSQLRec'"

        #################################### Added By Cristian ###################################
        Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query "UPDATE endpoints SET DID = '$newdid',LineURI = '$LineURI' WHERE ID  = '$UsrSQLRec'"
        ##########################################################################################
        
        If ($Rec.CSOnlineVoiceRoutingPolicy -ne $policy2)
        {
            Write-Host "Reassigning Online Voice Routing Policy ... " -ForegroundColor Cyan -NoNewline
            Write-Host "Changing Voice Policies from: "$Rec.TenantDialPlan" to: "$policy2
            Grant-CsOnlineVoiceRoutingPolicy $Script:UPN -PolicyName $policy2 -ErrorAction SilentlyContinue
            if ($?)
            {
                Write-Host "successful" -ForegroundColor Green
                $LineToWrite = "CHG " + "`tChanged Online Voice Routing Policy"
            }
            else
            {
                $LineToWrite = "ERR " + "`tError Changing Online Voice Routing Policy"
                Write-Host "failed" -ForegroundColor Red
            }
            WriteReportEvent 
        
            Write-Host "Changing Tenant Dial Plan ... " -ForegroundColor Cyan -NoNewline
            Grant-CsTenantDialPlan $Script:UPN -PolicyName $policy2 -ErrorAction SilentlyContinue
            if ($?)
            {
                Write-Host "successful" -ForegroundColor Green
                $LineToWrite = "CHG " + "`tChanged Tenant Dial Plan"
            }
            else
            {
                Write-Host "failed" -ForegroundColor Red
                $LineToWrite = "ERR " + "`tError Changing Tenant Dial Plan"
            }
            WriteReportEvent 
        
            Write-Host "Changing Emergency Call Routing Policy ... " -ForegroundColor Cyan -NoNewline
            Grant-CsTeamsEmergencyCallRoutingPolicy $Script:UPN -PolicyName $policy2 -ErrorAction SilentlyContinue
            if ($?)
            {
                Write-Host "successful" -ForegroundColor Green
                $LineToWrite = "CHG " + "`tChanged Emergency Call Routing Policy"
            }
            else
            {
                Write-Host "failed" -ForegroundColor Red
                $LineToWrite = "ERR " + "`tError Changing Emergency Call Routing Policy"
            }
            WriteReportEvent

            Write-Host "Changing Emergency Calling Policy ... " -ForegroundColor Cyan -NoNewline
            Grant-CsTeamsEmergencyCallingPolicy $Script:UPN -PolicyName $policy2 -ErrorAction SilentlyContinue
            if ($?)
            {
                Write-Host "successful" -ForegroundColor Green
                $LineToWrite = "CHG " + "`tChanged Emergency Calling Policy"  
            }
            else
            {
                Write-Host "failed" -ForegroundColor Red
                $LineToWrite = "ERR " + "`tError Changing Emergency Calling Policy"
            }
            WriteReportEvent

            write-Host "Changing Voice Policies in SQL Database" -ForegroundColor Cyan
            ## e:\O365AdminShared\Scripts\MySQL.ps1 -Query "UPDATE endpoints SET CsOnlineVoiceRoutingPolicy = '$policy2', TenantDialPlan = '$policy2',TeamsEmergencyCallingPolicy = '$policy2', TeamsEmergencyCallRoutingPolicy = '$policy2' WHERE EndpointID  = '$UsrSQLRec'"
            
            ############################ Added By Cristian #################################
            Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query "UPDATE endpoints SET VoicePolicy = '$policy2', DialPlan = '$policy2',EmergencyCallingPolicy = '$policy2', EmergencyCallRoutingPolicy = '$policy2' WHERE ID  = '$UsrSQLRec'"
            ################################################################################
        }
    }
    else
    {
        Write-Host "User not enabled. Do manually first ..." -ForegroundColor Red
    }
}

function Remove-PhoneDetails
{
    Write-Host "Disabling Phone Features: "$Script:txtEmpID.Text

    ## $Rec = E:\O365AdminShared\Scripts\MYSQL.PS1 -Query "SELECT * FROM endpoints WHERE EmployeeID LIKE '$Script:EmpNo'"

    #################################### Added By Cristian ###################################
    $Rec = Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query "SELECT * FROM endpoints WHERE SamAccountName LIKE '"+$Script:EmpNo+"'"
    ##########################################################################################

    $LineToWrite = "INFO" + "`tTicket Number                          : " + $Script:txtTicketNo.Text
    WriteReportEvent
    $LineToWrite = "INFO" + "`tDisableing Phone Features for          : " + $Script:txtEmpID.Text
    WriteReportEvent

    write-host "Clearing AD Attributes ..." -ForegroundColor Cyan
    $Script:ADUsr = Get-ADuser -Properties * -Filter { (SAmAccountName -eq $EmpNo) } | Select msRTCSIP-PrimaryUserAddress, msRTCSIP-DeploymentLocator, msRTCSIP-FederationEnabled, msRTCSIP-InternetAccessEnabled, msRTCSIP-Line, msRTCSIP-PrimaryHomeServer, msRTCSIP-UserEnabled, msRTCSIP-UserPolicies, msRTCSIP-UserPolicy, UserAccountControl
    Set-ADUser $Script:EmpNo -Clear 'msRTCSIP-DeploymentLocator'
    Set-ADUser $Script:EmpNo -Clear 'msRTCSIP-FederationEnabled'
    Set-ADUser $Script:EmpNo -Clear 'msRTCSIP-InternetAccessEnabled'
    Set-ADUser $Script:EmpNo -Clear 'msRTCSIP-UserEnabled'
    Set-ADUser $Script:EmpNo -Clear 'msRTCSIP-Line'
    $LineToWrite = "CLEAR" + "`tClearing DeploymentLocator             : " + $ADUsr.'msRTCSIP-DeploymentLocator'
    WriteReportEvent
    $LineToWrite = "DISAB " + "`tClearing/Disabling Federation"
    WriteReportEvent
    $LineToWrite = "DISAB " + "`tClearing/Disabling Internet Access"
    WriteReportEvent
    $LineToWrite = "CLEAR " + "`tClearing Extension                     : " + $ADUsr.'msRTCSIP-Line'
    WriteReportEvent
    $LineToWrite = "NOCHG " + "`tRetaining SIP Address                  : " + $ADUsr.'msRTCSIP-PrimaryUserAddress'
    WriteReportEvent
        
#    write-host "Removing Phone Licnese ..." -ForegroundColor Cyan
#    Set-MsolUserLicense -UserPrincipalName $Script:UPN -RemoveLicense "ul:MCOEV" -ErrorAction SilentlyContinue
#    $LineToWrite = "REMOV" + "`tRemoving Phone License (ul:MCOEV"
#    WriteReportEvent
    
    write-host "Removing Phone Details from SQL Database ..." -ForegroundColor Cyan
    If ($Rec.EndpointID.count -gt 0)
    {
        write-host "Removing Phone Details from SQL Database ..." -ForegroundColor Cyan
        $LineToWrite = "CLEAR" + "`tClearing DisplayName                  : " + $Rec.DisplayName
        WriteReportEvent 
        $LineToWrite = "CLEAR" + "`tClearing SipAddress                   : " + $Rec.SipAddress
        WriteReportEvent 
        $LineToWrite = "CLEAR" + "`tClearing EmployeeID                   : " + $Rec.EmployeeID
        WriteReportEvent 
        $LineToWrite = "CLEAR" + "`tClearing DID                          : " + $Rec.DID
        WriteReportEvent 
        $LineToWrite = "CLEAR" + "`tClearing Type                         : " + $Rec.Type
        WriteReportEvent 
        $LineToWrite = "CLEAR" + "`tClearing HostingProvider              : " + $Rec.HostingProvider
        WriteReportEvent 
        $LineToWrite = "CLEAR" + "`tClearing CSOnLineVoiceRoutingPolicy   : " + $Rec.CsOnLineVoiceRoutingPolicy
        WriteReportEvent 
        $LineToWrite = "CLEAR" + "`tClearing TenantDialPlan               : " + $Rec.TenantDialPlan
        WriteReportEvent 
        $LineToWrite = "CLEAR" + "`tClearing TeamsUpgradeEffectiveMode    : " + $Rec.TeamsUpgradeEffectiveMode
        WriteReportEvent 
        $LineToWrite = "CLEAR" + "`tClearing TeamsEnterpriseVoiceEnabled  : " + $Rec.TeamsEnterpriseVoiceEnabled
        WriteReportEvent 
        $LineToWrite = "CLEAR" + "`tClearing TeamsLineURI                 : " + $Rec.TeamsLineURI
        WriteReportEvent 
        $LineToWrite = "CLEAR" + "`tClearing TeamsHostedVoiceMail         : " + $Rec.TeamsHostedVoiceMail
        WriteReportEvent
        $UsrSQLRec = $Rec.EndpointID
        ## e:\O365AdminShared\Scripts\MYSQL.PS1 -Query "DELETE FROM endpoints WHERE EndpointID = '$UsrSQLRec'"

        ################################## Added by Cristian ####################################
        Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query "DELETE FROM endpoints WHERE ID = '$UsrSQLRec'"
        #########################################################################################


        $LineToWrite = "REMOV " + "`tRemoving Entry in SQL DB             : " + $Rec.TeamsLineURI
        WriteReportEvent 
    }
    else
    {
        $LineToWrite = "NoRec" + "`tNo Record found in the SQL Database"
        WriteReportEvent       
    }
}

####################################################################

#Set-Variable -Name upn -Option AllScope
Set-Variable -Name office -Option AllScope
Set-Variable -Name policy -Option AllScope
Set-Variable -Name didranges -Option AllScope
Set-Variable -Name selection -Option AllScope
Set-Variable -Name didranges -Option AllScope
Set-Variable -Name didwork -Option AllScope
Set-Variable -Name didend -Option AllScope
Set-Variable -Name lineuri -Option AllScope
Set-Variable -Name newdid -Option AllScope
Set-Variable -Name status -Option AllScope
Set-Variable -Name sipaddress -Option AllScope
Set-Variable -Name evenabled -Option AllScope
$status = New-Object -TypeName psobject

##################### Added by Cristian ########################
$Script:dbserver = "usnbkmsfb005p.global.ul.com"
$Script:db = "uchelper"
$Script:dbuser = "sfbhelper"
$Script:dbpw = "Underwr1terS"
################################################################

Build-PhoneDetails
Add-FormStandardButtons
$Global:OKButton.Text = "Assign"
$Global:OKButton.visible = $False
Publish-Form

If (($Script:txtTicketNo.Text.Length -eq 0) -and ($Global:Result -eq "OK") -and ($Script:ADUsr.Enabled -eq $true))
{
    Do
    {
        $Global:InputFocus = $Script:txtTicketNo
        $Output = $wshell.Popup("You must enter the ticket number.", 0, "TicketNo Required", 0 + 32)
        Publish-Form
    } while (($Script:txtTicketNo.Text.Length -eq 0) -and ($Global:Result -eq "OK"))
}

If ($Global:Result -eq "OK")
{
    $Script:EmpNo = ($Script:txtEmpID.Text) -Replace (" ", "")
    $Script:UPN = $Script:EmpNo + "@global.ul.com"
    $ReportFile = $Script:txtRptFile.Text
    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
    WriteReportEvent

    If ($Script:chkAssignExt.Checked -eq "Checked")
    {
        #Perform steps to assign New Extension
        If ($Global:Result -eq "OK")
        {
            #Assign button selected
            $LineToWrite = "INFO" + "`t" + "Assigning Phone Number to Account"
            WriteReportEvent
#            $HasE53 = [bool](((Get-MsolUser -UserPrincipalName $Script:UPN).Licenses) | Where-Object { $_.AccountSkuId -eq "ul:ENTERPRISEPACK" })
#            $HasPhone = [bool](((Get-MsolUser -UserPrincipalName $Script:UPN).Licenses) | Where-Object { $_.AccountSkuId -eq "ul:MCOEV" })
#            If (($HasPhone -ne $True) -and ($HasE3 -eq $True))
#            {
#                Set-MsolUserLicense -UserPrincipalName $Script:UPN -AddLicense "ul:MCOEV"
#                write-host "Phone License Assigned to Account"
                $status | Add-Member -MemberType NoteProperty -Name license -Value "yes"
#            }
            Set-LineURI

            $Name = $Script:ADUsr.DisplayName
            ## e:\O365AdminShared\scripts\MySQL.ps1 -Query "INSERT INTO endpoints (DisplayName,EmployeeID,SipAddress,TeamsLineURI,DID,Type,HostingProvider,CsOnlineVoiceRoutingPolicy,TenantDialPlan,TeamsEnterpriseVoiceEnabled,TeamsHostedVoiceMail,TeamsEmergencyCallingPolicy,TeamsEmergencyCallRoutingPolicy) VALUE ('$Name','$EmpNo','$sipaddress','$lineuri','$newdid','CsUser','sipfed.online.lync.com','$policy','$policy','true','true','$policy','$policy')" 
            
            ######################################## Added by Cristian ###############################################
            
            ## Define SQL Queries:
            $query0 = "SELECT SamAccountName FROM endpoints WHERE SamAccountName = '" + $Script:EmpNo + "'"
            $query1 = "UPDATE endpoints SET DisplayName = '$Name', SamAccountName = '$EmpNo', SipAddress = '$sipaddress', LineURI = '$lineuri', DID = '$newdid', Type = 'User', VoicePolicy = '$policy', DialPlan = '$policy', EnterpriseVoiceEnabled = 'True', HostedVoiceMail = 'True', EmergencyCallingPolicy = '$policy', EmergencyCallRoutingPolicy = '$policy', Office = '" + $Script:ADUsr.PhysicalDeliveryOfficeName + "' WHERE samaccountname = '$EmpNo'"
            $query2 = "INSERT INTO endpoints (DisplayName,SamAccountName,SipAddress,LineURI,DID,Type,VoicePolicy,DialPlan,EnterpriseVoiceEnabled,HostedVoiceMail,EmergencyCallingPolicy,EmergencyCallRoutingPolicy,Office) VALUES ('$Name','$EmpNo','$sipaddress','$lineuri','$newdid','User','$policy','$policy','True','True','$policy','$policy','" + $Script:ADUsr.PhysicalDeliveryOfficeName + "')"
            $query3 = "INSERT INTO newhirelog (samaccountname,adminaccount,snticket) VALUES ('$EmpNo','SDAP','" + $Script:txtTicketNo.Text + "')"

            ## Check if DataSet exists in DB:
            $result = Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query $query0
            if ($result.SamAccountName.Count -ge 1)
            {
                ## If Yes, update existing DataSet
                Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query $query1
            }
            else
            {
                ## If No, create a new DataSet
                Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query $query2
            }
            
            ## Add an entry into the log db.
            Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query $query3
            
            ##########################################################################################################

            Write-Host "Enterprise Voice has been enabled" -ForegroundColor Cyan
            $status | Add-Member -MemberType NoteProperty -Name enterprisevoiceenabled -Value "yes"
            
            if ($status.enterprisevoiceenabled -ne "yes")
            {
                ## e:\O365AdminShared\scripts\MySQL.ps1 -Query "INSERT INTO endpoints (DisplayName,EmployeeID,SipAddress,TeamsLineURI,DID,Type,HostingProvider,CsOnlineVoiceRoutingPolicy,TenantDialPlan,TeamsEnterpriseVoiceEnabled,TeamsHostedVoiceMail,TeamsEmergencyCallingPolicy,TeamsEmergencyCallRoutingPolicy) VALUE ('$Name','$EmpNo','$sipaddress','$lineuri','$newdid','CsUser','sipfed.online.lync.com','$policy','$policy','false','false','$policy','$policy')"
                
                ######################################## Added by Cristian ###############################################
                ## Define SQL Queries:
                $query0 = "SELECT SamAccountName FROM endpoints WHERE SamAccountName = '" + $Script:EmpNo + "'"
                $query1 = "UPDATE endpoints SET DisplayName = '$Name', SamAccountName = '$EmpNo', SipAddress = '$sipaddress', LineURI = '$lineuri', DID = '$newdid', Type = 'User', VoicePolicy = '$policy', DialPlan = '$policy', EnterpriseVoiceEnabled = 'False', HostedVoiceMail = 'False', EmergencyCallingPolicy = '$policy', EmergencyCallRoutingPolicy = '$policy', Office = '" + $Script:ADUsr.PhysicalDeliveryOfficeName + "' WHERE samaccountname = '$EmpNo'"
                $query2 = "INSERT INTO endpoints (DisplayName,SamAccountName,SipAddress,LineURI,DID,Type,VoicePolicy,DialPlan,EnterpriseVoiceEnabled,HostedVoiceMail,EmergencyCallingPolicy,EmergencyCallRoutingPolicy,Office) VALUES ('$Name','$EmpNo','$sipaddress','$lineuri','$newdid','User','$policy','$policy','False','False','$policy','$policy','" + $Script:ADUsr.PhysicalDeliveryOfficeName + "')"

                ## Check if DataSet exists in DB:
                $result = Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query $query0
                if ($result.SamAccountName.Count -ge 1)
                {
                    ## If Yes, update existing DataSet
                    Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query $query1
                }
                else
                {
                    ## If No, create a new DataSet
                    Invoke-Sqlcmd -ServerInstance $dbserver -Database $db -Username $dbuser -Password $dbpw -Query $query2
                }
                
                ##########################################################################################################

                $LineToWrite = "ASGN" + "`tEnterprise voice could not be set Teams Only Mode Assigned"
                Write-Host "Enterprise voice could not be set to true. Wait one hour and then do it manually" -ForegroundColor Red
                $status | Add-Member -MemberType NoteProperty -Name enterprisevoiceenabled -Value "no"
            }
        }
    }

    If ($Script:chkRemoveExt.Checked -eq "Checked")
    {
        #Perform steps to remove extension
        If ($Global:Result -eq "OK")
        {
            $LineToWrite = "INFO" + "`t" + "Removing Phone Nubmer from Account"
            WriteReportEvent
            Remove-PhoneDetails
        }
    }

    If ($Script:chkChgExt.Checked -eq "Checked")
    {
        #Perform steps to change extension
        If ($Global:Result -eq "OK")
        {
            $LineToWrite = "INFO" + "`t" + "Changing Phone Nubmer assigned to Account"
            WriteReportEvent
            write-host "Change button selection is under development"
#            $HasPhone = [bool](((Get-MsolUser -UserPrincipalName $Script:UPN).Licenses) | Where-Object { $_.AccountSkuId -eq "ul:MCOEV" })
#            If (($HasPhone -ne $True) -and ($HasE3 -eq $True))
#            {
#                Set-MsolUserLicense -UserPrincipalName $Script:UPN -AddLicense "ul:MCOEV"
#                write-host "Phone License Assigned to Account"
#                $LineToWrite = "ASGN" + "`tPhone License Assigned to Account"
#                $WriteReportEvent
#            }
            Change-LineURI
        }
    }
}