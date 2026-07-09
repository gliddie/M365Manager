#################################################################################
# 
# PowerShell source code
# Revision v1.05
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Enable MailUSer and Assign O365 License for existing AD User account
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :GET
#    'Author       : Sandi Glazebrook
#    'Date Created : 02/05/2021
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '    01/06/2021 SAG Adapted from the EnableUserAssignLisense.ps1 and the ConfigMailbox.ps1 scripts
#    '    03/03/2021 SAG Added disabling RemotePowerShell for new staff
#    '    03/04/2021 SAG Fixed code when address is in use, also fixed the get-ADPrincipalGroupMembership to add -ResourceContextServer global.ul.com which "broke" after systems were rebuilt.
#    '    04/13/2021 SAG Fixed the Reason not being sent to the email to the O365 Team also added code so if the Primary or Target addresses are blank it will pop-up that there was an issue and also in the post configuration the After details will show in red text.
#    '    04/16/2021 SAG Modified code for the new NFP org names
#    '    04/20/2021 SAG Added code to see if account maybe partially renamed - when ADUsername does not match the details in the firstname
#    '    05/07/2021 SAG Added code to check for non US-Based charactes in the AD username fields
#    '    05/14/2021 SAG Added check to confirm that there are email licenses available
#    '    07/01/2021 SAG Added code to add users to the MFA_Enabled group per request from security and removing spaced from the Generated Email Address
#    '    09/22/2021 SAG Added code to allow for the configuration of a forwarding address this is for use with M&A accounts only
#    '    10/01/2021 SAG Modified the reporting of the forwarding address after changes to write to the Report file not the Log file.  Also remove the extra "`t" + " code in the logging statements
#    '    10/19/2021 SAG Modified to check if the AD Account is disabled per request from G.Singh
#    '    01/12/2022 SAG Modifed to Disable Recordings in Teams
#    '    01/27/2022 SAG Modified code to update individuals that are port of the UL.ORG business unit
#    '    04/05/2022 SAG Added line to override the review of Available vs Consumed licenses $E3Free = [bool]$True
#    '    04/20/2022 SAG Modified to assign the E5 license
#    '    05/11/2022 SAG Modified code to allow the "-" to remain a part of the email address
#    '    07/19/2022 SAG Removed duplicate reporting of the No Busienableness Unit configured logging details.
#    '    08/26/2022 SAG Added reporting the Employee Type into the report details
#    '    09/29/2022 SAG Changed the lookup for the NFP org to look at the AD Company Value
#    '    10/27/2022 SAG Added initial logging if forwarding address is provided
#    '    10/28/2022 SAG Removed the session reconnection as it was using Basic Authentication
#
#################################################################################

Function Build-NoBusinessUnitForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "No Business Unit Configured" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 500 ; $form.Height = 240  # Make the form wider 
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    
    $Tab = 0
    $Global:Top = 30
    ## Message
    $Global:lblHost = New-Object System.Windows.Forms.Label   
        $Global:lblHost.Text = "The Business Unit value is not yet set and no manager is identified for this user. Before continuing with the enablement of this account you must provide the Business Unit for this individual."
        $Global:lblHost.Top = $Global:Top ; $Global:lblHost.Left = 20; $Global:lblHost.Width=440 ; $Global:lblHost.Height=40
        $form.Controls.Add($Global:lblHost)    # Add to Form 
 
    $Global:Top = $Global:Top + 40
    ## UL.COM
    $Global:chkulcom = New-Object Windows.Forms.checkbox 
        $Global:chkulcom.Left = 80; $Global:chkulcom.Width = 200; $Global:chkulcom.Top = $Global:Top
        $Global:chkulcom.Text = "For Profit Staff Member" 
        $Global:chkulcom.Checked = $false   # set a default value 
        $Global:chkulcom.TabIndex = 1
        $Global:form.Controls.Add($Global:chkulcom) 
        # Obtain Value with: $Global:chkulcom.Checked
 
    $Global:Top = $Global:Top + 20
    ## UL.ORG
    $Global:chkulorg = New-Object Windows.Forms.checkbox 
        $Global:chkulorg.Left = 80; $Global:chkulorg.Width = 200; $Global:chkulorg.Top = $Global:Top
        $Global:chkulorg.Text = "Not for Profit Staff Member" 
        $Global:chkulorg.Checked = $false   # set a default value 
        $Global:chkulorg.TabIndex = 1
        $Global:form.Controls.Add($Global:chkulorg) 
        # Obtain Value with: $Global:chkulorg.Checked
 
    $Global:Top = $Global:Top + 20
    ## UL.COM
    $Global:chkunknown = New-Object Windows.Forms.checkbox 
        $Global:chkunknown.Left = 80; $Global:chkunknown.Width = 200; $Global:chkunknown.Top = $Global:Top
        $Global:chkunknown.Text = "Unknown Business Unit" 
        $Global:chkunknown.Checked = $false   # set a default value 
        $Global:chkunknown.TabIndex = 1
        $Global:form.Controls.Add($Global:chkunknown) 
        # Obtain Value with: $Global:chkunknown.Checked

    Add-FormStandardButtons
}

Function Build-AcctEnablePriorDetails
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Enable User - Assign License - Configure Mailbox Details"
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 800 ; $form.Height = 760  # Make the form wider 
    
    $Global:Top = 10
    ## Employee ID
    $Global:lblEmpNo = New-Object System.Windows.Forms.Label   
        $Global:lblEmpNo.Text = "Employee No:"  
        $Global:lblEmpNo.Top = $Global:Top ; $Global:lblEmpNo.Left = 5; $Global:lblEmpNo.Width=150 ;$Global:lblEmpNo.AutoSize = $true 
        $form.Controls.Add($Global:lblEmpNo)    # Add to Form 
        # 
        $Global:txtEmpNo = New-Object Windows.Forms.TextBox  
        $Global:txtEmpNo.ReadOnly = $True;
        $Global:txtEmpNo.Top = $Global:Top; $Global:txtEmpNo.Left = 160; $Global:txtEmpNo.Width = 120;  
        $Global:txtEmpNo.Text = $Global:txtInpEmpNo.Text
        $Global:form.Controls.Add($Global:txtEmpNo)    # Add to Form
        $Global:InputFocus = $Global:txtEmpNo
    $Global:Top = $Global:Top + 30

    ## Primary Address
    $Global:lblPrimaryAddr = New-Object System.Windows.Forms.Label   
        $Global:lblPrimaryAddr.Text = "Primary Address:"  
        $Global:lblPrimaryAddr.Top = $Global:Top ; $Global:lblPrimaryAddr.Left = 5; $Global:lblPrimaryAddr.Width=150 ;$Global:lblPrimaryAddr.AutoSize = $true 
        $form.Controls.Add($Global:lblPrimaryAddr)    # Add to Form 
        # 
        $Global:txtPrimaryAddr = New-Object Windows.Forms.TextBox
        $Global:txtPrimaryAddr.ReadOnly = $True;
        $Global:txtPrimaryAddr.TabIndex = 50 # set Tab Order 
        $Global:txtPrimaryAddr.Top = $Global:Top; $Global:txtPrimaryAddr.Left = 160; $Global:txtPrimaryAddr.Width = 200;  
        $Global:txtPrimaryAddr.Text = $EmailAddress
        $Global:form.Controls.Add($Global:txtPrimaryAddr)    # Add to Form
    $Global:Top = $Global:Top + 40

    ## Settings Prior to Execution
    $Global:lblPrior = New-Object System.Windows.Forms.Label   
        $Global:lblPrior.Text = "AD Account values prior to mail-enabling the account:"
        $Global:lblPrior.ForeColor = "Blue"; 
        $Global:lblPrior.Top = $Global:Top ; $Global:lblPrior.Left = 5; $Global:lblPrior.Width=400 ;$Global:lblPrior.AutoSize = $true 
        $form.Controls.Add($Global:lblPrior)    # Add to Form 
    $Global:Top = $Global:Top + 30

    ## Current Primary Address
    $LabelLeft = 20
    $VariableLeft = 175
    $Global:lblCurrPrimaryAddr = New-Object System.Windows.Forms.Label   
        $Global:lblCurrPrimaryAddr.Text = "Current Primary Address:"  
        $Global:lblCurrPrimaryAddr.Top = $Global:Top ; $Global:lblCurrPrimaryAddr.Left = $LabelLeft; $Global:lblCurrPrimaryAddr.Width=150 ;$Global:lblCurrPrimaryAddr.AutoSize = $true 
        $form.Controls.Add($Global:lblCurrPrimaryAddr)    # Add to Form 
        # 
        $Global:txtCurrPrimaryAddr = New-Object Windows.Forms.TextBox
        $Global:txtCurrPrimaryAddr.ReadOnly = $True; 
        $Global:txtCurrPrimaryAddr.Top = $Global:Top; $Global:txtCurrPrimaryAddr.Left = $VariableLeft; $Global:txtCurrPrimaryAddr.Width = 350;  
        $Global:txtCurrPrimaryAddr.Text = $ADUser.mail
        $Global:form.Controls.Add($Global:txtCurrPrimaryAddr)    # Add to Form
    $Global:Top = $Global:Top + 30

    ## Current Target Address
    $Global:lblTargetAddr = New-Object System.Windows.Forms.Label   
        $Global:lblTargetAddr.Text = "AD Target Address:"  
        $Global:lblTargetAddr.Top = $Global:Top ; $Global:lblTargetAddr.Left = $LabelLeft; $Global:lblTargetAddr.Width=150 ;$Global:lblTargetAddr.AutoSize = $true 
        $form.Controls.Add($Global:lblTargetAddr)    # Add to Form 
        # 
        $Global:txtTargetAddr = New-Object Windows.Forms.TextBox
        $Global:txtTargetAddr.ReadOnly = $True;  
        $Global:txtTargetAddr.Top = $Global:Top; $Global:txtTargetAddr.Left = $VariableLeft; $Global:txtTargetAddr.Width = 350;
        $Global:txtTargetAddr.Text = $ADUser.targetaddress
        If ($ADUser.targetaddress -like "SMTP:*")
        {  
            $Global:txtTargetAddr.Text = (($ADUser.targetaddress) -replace("SMTP:",""))
        }
        $Global:form.Controls.Add($Global:txtTargetAddr)    # Add to Form
    $Global:Top = $Global:Top + 30

    ## Current Proxy Addresses
    $Global:lblProxyAddr = New-Object System.Windows.Forms.Label   
        $Global:lblProxyAddr.Text = "Proxy Addresses:"  
        $Global:lblProxyAddr.Top = $Global:Top ; $Global:lblProxyAddr.Left = $LabelLeft; $Global:lblProxyAddr.Width=150 ;$Global:lblProxyAddr.AutoSize = $true 
        $form.Controls.Add($Global:lblProxyAddr)    # Add to Form 
        # 
        $Global:txtProxyAddr = New-Object Windows.Forms.ListBox
        $Global:txtProxyAddr.BackColor ="LightGray"
        $Global:txtProxyAddr.Top = $Global:Top; $Global:txtProxyAddr.Left = $Variableleft; $Global:txtProxyAddr.Width = 350; $Global:txtProxyAddr.Height = 60;
        Foreach ($Proxy in $ADUser.proxyaddresses)
        {
            If ($Proxy -cnotlike "SMTP:*")
            {
                [void] $Global:txtProxyAddr.Items.Add($Proxy)
            }
        }  
        $Global:form.Controls.Add($Global:txtProxyAddr)    # Add to Form
    $Global:Top = $Global:Top + 70

    ## Current ExtensionAttribute15 Value
    $Global:lblExtAttrib = New-Object System.Windows.Forms.Label   
        $Global:lblExtAttrib.Text = "ExtensionAttribute15:"  
        $Global:lblExtAttrib.Top = $Global:Top ; $Global:lblExtAttrib.Left = $LabelLeft; $Global:lblExtAttrib.Width=150 ;$Global:lblExtAttrib.AutoSize = $true 
        $form.Controls.Add($Global:lblExtAttrib)    # Add to Form 
        # 
        $Global:txtExtAttrib = New-Object Windows.Forms.TextBox  
        $Global:txtExtAttrib.ReadOnly = $True;
        $Global:txtExtAttrib.Top = $Global:Top; $Global:txtExtAttrib.Left = $VariableLeft; $Global:txtExtAttrib.Width = 350;
        $Global:txtExtAttrib.Text = $ADUser.ExtensionAttribute15
        $Global:form.Controls.Add($Global:txtExtAttrib)    # Add to Form
}

Function Build-AcctEnabledAfterDetails
{
    $Global:Top = $Global:Top + 40
    ## Settings After Enablement
    $Global:lblPost = New-Object System.Windows.Forms.Label
    $Global:lblPost.Text = "AD Account values after mail-enabling and configuring the mailbox:"
        If (($ADUserPost.mail.length -le 6) -or ($ADUserPost.targetaddress.length -le 6))
        {   
            $Global:lblPost.ForeColor = "Blue";
        }
        else
        {
            $Global:lblPost.Text = "AD Account values contain errors after mail-enabling and configuring the mailbox:"
            $Global:lblPost.ForeColor = "Red";
        }
        $Global:lblPost.Top = $Global:Top ; $Global:lblPost.Left = 5; $Global:lblPost.Width=400 ;$Global:lblPost.AutoSize = $true 
        $form.Controls.Add($Global:lblPost)    # Add to Form 
    $Global:Top = $Global:Top + 30

    ## Current Primary Address
    $LabelLeft = 20
    $VariableLeft = 175
    $Global:lblPostPrimaryAddr = New-Object System.Windows.Forms.Label   
        $Global:lblPostPrimaryAddr.Text = "User Primary Address:"  
        $Global:lblPostPrimaryAddr.Top = $Global:Top ; $Global:lblPostPrimaryAddr.Left = $LabelLeft; $Global:lblPostPrimaryAddr.Width=150 ;$Global:lblPostPrimaryAddr.AutoSize = $true 
        $form.Controls.Add($Global:lblPostPrimaryAddr)    # Add to Form 
        # 
        $Global:txtPostPrimaryAddr = New-Object Windows.Forms.TextBox
        $Global:txtPostPrimaryAddr.ReadOnly = $True; 
        $Global:txtPostPrimaryAddr.Top = $Global:Top; $Global:txtPostPrimaryAddr.Left = $VariableLeft; $Global:txtPostPrimaryAddr.Width = 350;  
        $Global:txtPostPrimaryAddr.Text = $ADUserPost.mail
        $Global:form.Controls.Add($Global:txtPostPrimaryAddr)    # Add to Form
    $Global:Top = $Global:Top + 30

    ## Current Target Address
    $Global:lblPostTargetAddr = New-Object System.Windows.Forms.Label   
        $Global:lblPostTargetAddr.Text = "AD Target Address:"  
        $Global:lblPostTargetAddr.Top = $Global:Top ; $Global:lblPostTargetAddr.Left = $LabelLeft; $Global:lblPostTargetAddr.Width=150 ;$Global:lblPostTargetAddr.AutoSize = $true 
        $form.Controls.Add($Global:lblPostTargetAddr)    # Add to Form 
        # 
        $Global:txtPostTargetAddr = New-Object Windows.Forms.TextBox
        $Global:txtPostTargetAddr.ReadOnly = $True;  
        $Global:txtPostTargetAddr.Top = $Global:Top; $Global:txtPostTargetAddr.Left = $VariableLeft; $Global:txtPostTargetAddr.Width = 350;
        $Global:txtPostTargetAddr.Text = (($ADUserPost.targetaddress) -replace ("SMTP:",""))
        If ($MgrAddress -eq "Yes")
        {
            $Global:txtPostTargetAddr.Text = $Global:txtPostTargetAddr.Text
        }
        $Global:txtPosTargetAddr.Text
        $Global:form.Controls.Add($Global:txtPostTargetAddr)    # Add to Form
    $Global:Top = $Global:Top + 30

    ## Current Proxy Addresses
    $Global:lblPostProxyAddr = New-Object System.Windows.Forms.Label   
        $Global:lblPostProxyAddr.Text = "Proxy Addresses:"  
        $Global:lblPostProxyAddr.Top = $Global:Top ; $Global:lblPostProxyAddr.Left = $LabelLeft; $Global:lblPostProxyAddr.Width=150 ;$Global:lblPostProxyAddr.AutoSize = $true 
        $form.Controls.Add($Global:lblPostProxyAddr)    # Add to Form 
        # 
        $Global:txtPostProxyAddr = New-Object Windows.Forms.ListBox
        $Global:txtPostProxyAddr.BackColor ="LightGray"
        $Global:txtPostProxyAddr.Top = $Global:Top; $Global:txtPostProxyAddr.Left = $Variableleft; $Global:txtPostProxyAddr.Width = 350; $Global:txtPostProxyAddr.Height = 60;
        Foreach ($Proxy in $ADUserPost.proxyaddresses)
        {
            If ($Proxy -cnotlike "SMTP:*")
            {
                [void] $Global:txtPostProxyAddr.Items.Add($Proxy)
            }
        }  
        $Global:form.Controls.Add($Global:txtPostProxyAddr)    # Add to Form
    $Global:Top = $Global:Top + 70

    ## Business Unit
    $Global:lblBusUnit = New-Object System.Windows.Forms.Label   
        $Global:lblBusUnit.Text = "Business Unit:"
        $Global:lblBusUnit.Top = $Global:Top ; $Global:lblBusUnit.Left = $LabelLeft; $Global:lblBusUnit.Width=150 ;$Global:lblBusUnit.AutoSize = $true 
        $form.Controls.Add($Global:lblBusUnit)    # Add to Form 
        # 
        $Global:txtBusUnit = New-Object Windows.Forms.TextBox
        $Global:txtBusUnit.ReadOnly = $True; 
        $Global:txtBusUnit.Top = $Global:Top; $Global:txtBusUnit.Left = $Variableleft; $Global:txtBusUnit.Width = 200; $Global:txtBusUnit.Height = 60;
        $Global:txtBusUnit.Text = $ADUser.Company
        If ($ADUser.Company.length -lt 1) 
        {
            $Global:txtBusUnit.Width = 350;
            $Global:txtBusUnit.Text = "Used manager BU to determine email domain"
        }
        $Global:form.Controls.Add($Global:txtBusUnit)    # Add to Form
    $Global:Top = $Global:Top + 30

    ## Region
    $Global:lblRegion = New-Object System.Windows.Forms.Label   
        $Global:lblRegion.Text = "Region:"  
        $Global:lblRegion.Top = $Global:Top ; $Global:lblRegion.Left = $LabelLeft; $Global:lblRegion.Width=150 ;$Global:lblRegion.AutoSize = $true 
        $form.Controls.Add($Global:lblRegion)    # Add to Form 
        # 
        $Global:txtRegion = New-Object Windows.Forms.TextBox
        $Global:txtRegion.ReadOnly = $True;
        $Global:txtRegion.Top = $Global:Top; $Global:txtRegion.Left = $Variableleft; $Global:txtRegion.Width = 200; $Global:txtRegion.Height = 60;
        $Global:txtRegion.Text = $MBXDetails.CustomAttribute5  
        $Global:form.Controls.Add($Global:txtRegion)    # Add to Form
    $Global:Top = $Global:Top + 30

    ## Retention Policy Set
    $Global:lblRetPolicy = New-Object System.Windows.Forms.Label   
        $Global:lblRetPolicy.Text = "Retention Policy:"  
        $Global:lblRetPolicy.Top = $Global:Top ; $Global:lblRetPolicy.Left = $LabelLeft; $Global:lblRetPolicy.Width=150 ;$Global:lblRetPolicy.AutoSize = $true 
        $form.Controls.Add($Global:lblRetPolicy)    # Add to Form 
        # 
        $Global:txtRetPolicy = New-Object Windows.Forms.TextBox
        $Global:txtRetPolicy.ReadOnly = $True;
        $Global:txtRetPolicy.Top = $Global:Top; $Global:txtRetPolicy.Left = $Variableleft; $Global:txtRetPolicy.Width = 200; $Global:txtRetPolicy.Height = 60;
        $Global:txtRetPolicy.Text = $MBXDetails.RetentionPolicy
        $Global:form.Controls.Add($Global:txtRetPolicy)    # Add to Form
    $Global:Top = $Global:Top + 30
    
    ##InTune Group
    $Global:lblIntuneGrp = New-Object System.Windows.Forms.Label   
        $Global:lblIntuneGrp.Text = "Intune Policy Group:"  
        $Global:lblIntuneGrp.Top = $Global:Top ; $Global:lblIntuneGrp.Left = $LabelLeft; $Global:lblIntuneGrp.Width=150 ;$Global:lblIntuneGrp.AutoSize = $true 
        $form.Controls.Add($Global:lblIntuneGrp)    # Add to Form 
        # 
        $Global:txtIntuneGrp = New-Object Windows.Forms.TextBox
        $Global:txtIntuneGrp.ReadOnly = $True;
        $Global:txtIntuneGrp.Top = $Global:Top; $Global:txtIntuneGrp.Left = $Variableleft; $Global:txtIntuneGrp.Width = 200; $Global:txtIntuneGrp.Height = 60;
        $Global:txtIntuneGrp.Text = $IntuneGroup
        $Global:form.Controls.Add($Global:txtIntuneGrp)    # Add to Form
    $Global:Top = $Global:Top + 30

    ##Welcome Message
        $Global:lblWelcomeMsg = New-Object System.Windows.Forms.Label   
        $Global:lblWelcomeMsg.Text = "Welcome Message Sent:"  
        $Global:lblWelcomeMsg.Top = $Global:Top ; $Global:lblWelcomeMsg.Left = $LabelLeft; $Global:lblWelcomeMsg.Width=150 ;$Global:lblWelcomeMsg.AutoSize = $true 
        $form.Controls.Add($Global:lblWelcomeMsg)    # Add to Form 
        # 
        $Global:txtWelcomeMsg = New-Object Windows.Forms.TextBox
        $Global:txtWelcomeMsg.ReadOnly = $True;
        $Global:txtWelcomeMsg.Top = $Global:Top; $Global:txtWelcomeMsg.Left = $Variableleft; $Global:txtWelcomeMsg.Width = 200;
        $Global:txtWelcomeMsg.Text = "Yes"
        If ($NotSent -ne "Y")
        {
            $Global:txtWelcomeMsg.Width = 350;
            $Global:txtWelcomeMsg.Text = "No, Multiple Configs - " + $Global:txtReason.Text
        }
        $Global:form.Controls.Add($Global:txtWelcomeMsg)    # Add to Form
    $Global:Top = $Global:Top + 30

    ## Current ExtensionAttribute15 Value
    $Global:lblPostExtAttrib = New-Object System.Windows.Forms.Label   
        $Global:lblPostExtAttrib.Text = "ExtensionAttribute15:"  
        $Global:lblPostExtAttrib.Top = $Global:Top ; $Global:lblPostExtAttrib.Left = $LabelLeft; $Global:lblPostExtAttrib.Width=150 ;$Global:lblPostExtAttrib.AutoSize = $true 
        $form.Controls.Add($Global:lblPostExtAttrib)    # Add to Form 
        # 
        $Global:txtPostExtAttrib = New-Object Windows.Forms.TextBox  
        $Global:txtPostExtAttrib.ReadOnly = $True;
        $Global:txtPostExtAttrib.Top = $Global:Top; $Global:txtPostExtAttrib.Left = $VariableLeft; $Global:txtPostExtAttrib.Width = 350;  
        $Global:txtPostExtAttrib.Text = $ADUserPost.extensionattribute15
        $Global:form.Controls.Add($Global:txtPostExtAttrib)    # Add to Form
    $Global:Top = $Global:Top + 30
  
    #Report Details
    $Global:lblRptFile = New-Object System.Windows.Forms.Label   
        $Global:lblRptFile.Text = "Report Details:"  
        $Global:lblRptFile.Top = $Top ; $Global:lblRptFile.Left = $LabelLeft; $Global:lblRptFile.Width=100 ;$Global:lblRptFile.AutoSize = $true 
        $form.Controls.Add($Global:lblRptFile)    # Add to Form 
        # 
        $Global:txtRptFile = New-Object Windows.Forms.TextBox
        $Global:txtRptFile.ReadOnly = $true; 
        $Global:txtRptFile.TabIndex = $Tab++ # set Tab Order
        $Global:txtRptFile.Top = $Top; $Global:txtRptFile.Left = $VariableLeft; $Global:txtRptFile.Width = 580; 
        $Global:txtRptFile.Text = $ReportFile
        $Global:form.Controls.Add($Global:txtRptFile)    # Add to Form

    Add-FormInfoOnlyButtons
}

Function Build-ConfirmAddress
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Confirm Email Address" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 550 ; $form.Height = 300  # Make the form wider 
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    
    $Tab = 0
    $Global:InUseTop = 30
    ## Message
    $Global:lblHost = New-Object System.Windows.Forms.Label   
        $Global:lblHost.Text = "Please confirm the email address.  If it all capitals please select and/or correct the the address that uses proper case."
        $Global:lblHost.Top = $Global:InUseTop ; $Global:lblHost.Left = 20; $Global:lblHost.Width=440 ; $Global:lblHost.Height=40
        $form.Controls.Add($Global:lblHost)    # Add to Form

    $Global:InUseTop = $Global:InUseTop + 40
    ## AD Username
        $Global:lblADUsrName = New-Object System.Windows.Forms.Label   
            $Global:lblADUsrName.Text = "AD User Name:"
            $Global:lblADUsrName.Top = $Global:InUseTop ; $Global:lblADUsrName.Left = 70; $Global:lblADUsrName.Width=90;
            $form.Controls.Add($Global:lblADUsrName)    # Add to Form 
            # 
            $Global:txtADUsrName = New-Object Windows.Forms.TextBox
            $Global:txtADUsrName.Top = $Global:InUseTop; $Global:txtADUsrName.Left = 170; $Global:txtADUsrName.Width = 250;  
            $Global:txtADUsrName.Text = $ADUser.Name
            $Global:form.Controls.Add($Global:txtADUsrName)    # Add to Form

    If (($ADuser.Name -match $ADUser.givenName) -and ($ADuser.Name -match $ADUser.sn))
    {
#        Do Nothing as the names match       
    }
    else
    {
            $Global:txtADUsrName.ReadOnly = $False;
            $Global:txtADUsrName.BackColor = "LightGray"
            $Global:txtADUsrName.ForeColor = "Red"
            $Global:txtADUsrName.ForeColor = "Red"
            $Global:form.Controls.Add($Global:txtADUsrName)    # Add to Form
        $Global:InUseTop = $Global:InUseTop + 30
        ## Message
        $Global:lblHost1 = New-Object System.Windows.Forms.Label
            $Global:lblHost1.Text = "The AD Username on this account does not match the specific details as shown below.  It appears this account is/was partially renamed."
            $Global:lblHost1.Top = $Global:InUseTop ; $Global:lblHost1.Left = 40; $Global:lblHost1.Width=420 ; $Global:lblHost1.Height=40
            $form.Controls.Add($Global:lblHost1)    # Add to Form
    }

    $Global:InUseTop = $Global:InUseTop + 40
    ## AD First Name
    $Global:lblFirstName = New-Object System.Windows.Forms.Label   
        $Global:lblFirstName.Text = "FirstName:"
        $Global:lblFirstName.Top = $Global:InUseTop ; $Global:lblFirstName.Left = 20; $Global:lblFirstName.Width=60 ;
        $form.Controls.Add($Global:lblFirstName)    # Add to Form 
        # 
        $Global:txtFirstName = New-Object Windows.Forms.TextBox
        $Global:txtFirstName.ReadOnly = $True;
        $Global:txtFirstName.Top = $Global:InUseTop; $Global:txtFirstName.Left = 90; $Global:txtFirstName.Width = 80;
        $Global:txtFirstName.Text = $ADUser.givenName
        $Global:form.Controls.Add($Global:txtFirstName)    # Add to Form
    ## AD MiddleInitial
    $Global:lblMiddleInt = New-Object System.Windows.Forms.Label   
        $Global:lblMiddleInt.Text = "MiddleInitial:"
        $Global:lblMiddleInt.Top = $Global:InUseTop ; $Global:lblMiddleInt.Left = 180; $Global:lblMiddleInt.Width=80;
        $form.Controls.Add($Global:lblMiddleInt)    # Add to Form 
        # 
        $Global:txtMiddleInt = New-Object Windows.Forms.TextBox
        $Global:txtMiddleInt.ReadOnly = $True;
        $Global:txtMiddleInt.TabIndex = 50 # set Tab Order 
        $Global:txtMiddleInt.Top = $Global:InUseTop; $Global:txtMiddleInt.Left = 270; $Global:txtMiddleInt.Width = 20;  
        $Global:txtMiddleInt.Text = $ADUser.Initials
        $Global:form.Controls.Add($Global:txtMiddleInt)    # Add to Form
    ## AD Last Name
    $Global:lblLastName = New-Object System.Windows.Forms.Label   
        $Global:lblLastName.Text = "LastName:"
        $Global:lblLastName.Top = $Global:InUseTop ; $Global:lblLastName.Left = 300; $Global:lblLastName.Width=70;
        $form.Controls.Add($Global:lblLastName)    # Add to Form 
        # 
        $Global:txtLastName = New-Object Windows.Forms.TextBox
        $Global:txtLastName.ReadOnly = $True;
        $Global:txtLastName.TabIndex = 50 # set Tab Order 
        $Global:txtLastName.Top = $Global:InUseTop; $Global:txtLastName.Left = 380; $Global:txtLastName.Width = 100;  
        $Global:txtLastName.Text = $ADUser.sn
        $Global:form.Controls.Add($Global:txtLastName)    # Add to Form

    $Global:InUseTop = $Global:InUseTop + 40
    ## System Generated Address
    $Global:chksysGen = New-Object Windows.Forms.checkbox 
        $Global:chksysGen.Left = 30; $Global:chksysGen.Width = 10; $Global:chksysGen.Top = ($Global:InUseTop-4)
        $Global:chksysGen.Checked = $true   # set a default value 
        $Global:form.Controls.Add($Global:chksysGen) 
        # Obtain Value with: $Global:sysGen.Checked
    $Global:lblSysGenAddr = New-Object System.Windows.Forms.Label   
        $Global:lblSysGenAddr.Text = "Default System Generated Address:"
        $Global:lblSysGenAddr.Top = $Global:InUseTop ; $Global:lblSysGenAddr.Left = 45; $Global:lblSysGenAddr.Width=150 ;$Global:lblSysGenAddr.AutoSize = $true 
        $form.Controls.Add($Global:lblSysGenAddr)    # Add to Form 
        # 
        $Global:txtSysGenAddr = New-Object Windows.Forms.TextBox
        $Global:txtSysGenAddr.TabIndex = 50 # set Tab Order 
        $Global:txtSysGenAddr.Top = $Global:InUseTop; $Global:txtSysGenAddr.Left = 240; $Global:txtSysGenAddr.Width = 200;  
        $Global:txtSysGenAddr.Text = $SysGenAddr
        $Global:form.Controls.Add($Global:txtSysGenAddr)    # Add to Form
        $Global:InputFocus = $Global:txtSysGenAddr

    If ($sysGenAddr -cne $ProperCaseAddress)
    {
        $Global:InUseTop = $Global:InUseTop + 40
        ## System Generated Address
        $Global:chkproperCase = New-Object Windows.Forms.checkbox 
            $Global:chkproperCase.Left = 30; $Global:chkproperCase.Width = 10; ($Global:chkproperCase.Top = $Global:InUseTop-4)
            $Global:chkproperCase.Checked = $false   # set a default value 
            $Global:form.Controls.Add($Global:chkproperCase) 
        # Obtain Value with: $Global:properCase.Checked
        $Global:lblProperCase = New-Object System.Windows.Forms.Label   
            $Global:lblProperCase.Text = "Alternate System Generated Address:"
            $Global:lblProperCase.Top = $Global:InUseTop ; $Global:lblProperCase.Left = 45; $Global:lblProperCase.Width=150 ;$Global:lblProperCase.AutoSize = $true 
            $form.Controls.Add($Global:lblProperCase)    # Add to Form 
            # 
            $Global:txtProperCase = New-Object Windows.Forms.TextBox
            $Global:txtProperCase.TabIndex = 50 # set Tab Order 
            $Global:txtProperCase.Top = $Global:InUseTop; $Global:txtProperCase.Left = 240; $Global:txtProperCase.Width = 200;  
            $Global:txtProperCase.Text = $ProperCaseAddress
            $Global:form.Controls.Add($Global:txtProperCase)    # Add to Form
            $Global:InputFocus = $Global:txtProperCase
    }
 
    Add-FormStandardButtons
}

Function Create-WelcomeMsg
{
#  Mail Variables
    $MsgBody = ""
    $EmailFrom = New-Object System.Net.Mail.MailAddress "DoNotReply@ul.com" , "UL Technology Services"
    $smtpServer = "SMTP-relay.ul.com"
    $Image = "e:\o365AdminShared\Data\WelcomeBanner.jpg"
    $html = Get-Content -LiteralPath e:\o365AdminShared\Data\WelcomeEMail.htm
#  Embed Image
    $att1 = new-object Net.Mail.Attachment($Image)
    $att1.ContentType.MediaType = “image/png”
    $att1.ContentId = “Attachment”

#  More information on Mailmessage on http://technet.microsoft.com/en-us/library/dd347693.aspx
    $message = New-Object system.net.mail.mailmessage

#  Add attachment to the mail
    $message.Attachments.Add($att1)

#  Mail body
    $MsgBody += $msgfont + "<img src='cid:$($att1.ContentId)' /><br>" + $html

#  Mail info
    $message.from = $EmailFrom
    $message.To.Add($MBxDetails.UserPrincipalName)
    $message.Bcc.Add("EnterpriseMessagingServices@ul.com")
    $message.Subject = "Welcome to UL"
    $message.Body = $MsgBody
    $message.IsBodyHTML = $true
    $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 25)
    $SMTPClient.Send($message)

#  Dispose attachments
    $att1.dispose()
}

Function Enter-Reason
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Welcome Previously Sent" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 400 ; $form.Height = 240  # Make the form wider 

    $Global:Top = 30
    ## Message
    $Global:lblDesc = New-Object System.Windows.Forms.Label   
        $Global:lblDesc.Text = "The configuration script for this mailbox has already been run. Please provide reason you are running the configuration again."
        $Global:lblDesc.Top = $Global:Top ; $Global:lblDesc.Left = 20; $Global:lblDesc.Width=340 ; $Global:lblDesc.Height=40
        $form.Controls.Add($Global:lblDesc)    # Add to Form 
    $Global:Top = $Global:Top + 40

    ## Reason
    $Global:lblReason = New-Object System.Windows.Forms.Label   
        $Global:lblReason.Text = "Reason:"  
        $Global:lblReason.Top = $Global:Top ; $Global:lblReason.Left = 60; $Global:lblReason.Width=60 ;$Global:lblReason.AutoSize = $true 
        $form.Controls.Add($Global:lblReason)    # Add to Form 
        # 
        $Global:txtReason = New-Object Windows.Forms.TextBox
        $Global:txtReason.TabIndex = 50 # set Tab Order 
        $Global:txtReason.Top = $Global:Top; $Global:txtReason.Left = 130; $Global:txtReason.Width = 140;  
        $Global:txtReason.Text = ""
        $Global:form.Controls.Add($Global:txtReason)
        $Global:InputFocus = $Global:txtReason

    Add-FormStandardButtons
}

Function Review-AddrInUse
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Default Address Already In Use" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 600 ; $form.Height = 300  # Make the form wider 
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    
    $Tab = 0
    $Global:InUseTop = 20
    ## Message
    $Global:lblHost = New-Object System.Windows.Forms.Label   
        $Global:lblHost.Text = "The default email address of FirstName.LastName@ul.com, FirstName.MI.LastName@ul.com, FirstName.LastName@ul.org or FirstName.MI.LastName@ul.org is already being used.  Based upon our rules the primary email address will be:"
        $Global:lblHost.Top = $Global:InUseTop ; $Global:lblHost.Left = 20; $Global:lblHost.Width=540 ; $Global:lblHost.Height=40
        $form.Controls.Add($Global:lblHost)    # Add to Form 

    $Global:InUseTop = $Global:InUseTop + 60
    ## AD First Name
    $Global:lblFirstName = New-Object System.Windows.Forms.Label   
        $Global:lblFirstName.Text = "FirstName:"
        $Global:lblFirstName.Top = $Global:InUseTop ; $Global:lblFirstName.Left = 20; $Global:lblFirstName.Width=60 ;
        $form.Controls.Add($Global:lblFirstName)    # Add to Form 
        # 
        $Global:txtFirstName = New-Object Windows.Forms.TextBox
        $Global:txtFirstName.ReadOnly = $True;
        $Global:txtFirstName.Top = $Global:InUseTop; $Global:txtFirstName.Left = 90; $Global:txtFirstName.Width = 80;  
        $Global:txtFirstName.Text = $ADUser.givenName
        $Global:form.Controls.Add($Global:txtFirstName)    # Add to Form
    ## AD MiddleInitial
    $Global:lblMiddleInt = New-Object System.Windows.Forms.Label   
        $Global:lblMiddleInt.Text = "MiddleInitial:"
        $Global:lblMiddleInt.Top = $Global:InUseTop ; $Global:lblMiddleInt.Left = 180; $Global:lblMiddleInt.Width=80;
        $form.Controls.Add($Global:lblMiddleInt)    # Add to Form 
        # 
        $Global:txtMiddleInt = New-Object Windows.Forms.TextBox
        $Global:txtMiddleInt.ReadOnly = $True;
        $Global:txtMiddleInt.TabIndex = 50 # set Tab Order 
        $Global:txtMiddleInt.Top = $Global:InUseTop; $Global:txtMiddleInt.Left = 270; $Global:txtMiddleInt.Width = 20;  
        $Global:txtMiddleInt.Text = $ADUser.Initials
        $Global:form.Controls.Add($Global:txtMiddleInt)    # Add to Form
    ## AD Last Name
    $Global:lblLastName = New-Object System.Windows.Forms.Label   
        $Global:lblLastName.Text = "LastName:"
        $Global:lblLastName.Top = $Global:InUseTop ; $Global:lblLastName.Left = 300; $Global:lblLastName.Width=70;
        $form.Controls.Add($Global:lblLastName)    # Add to Form 
        # 
        $Global:txtLastName = New-Object Windows.Forms.TextBox
        $Global:txtLastName.ReadOnly = $True;
        $Global:txtLastName.TabIndex = 50 # set Tab Order 
        $Global:txtLastName.Top = $Global:InUseTop; $Global:txtLastName.Left = 380; $Global:txtLastName.Width = 100;  
        $Global:txtLastName.Text = $ADUser.sn
        $Global:form.Controls.Add($Global:txtLastName)    # Add to Form

    $Global:InUseTop = $Global:InUseTop + 40
    ## System Generated Address
    $Global:lblSysGenAddr = New-Object System.Windows.Forms.Label   
        $Global:lblSysGenAddr.Text = "Requested Address:"
        $Global:lblSysGenAddr.ForeColor = "Red";
        $Global:lblSysGenAddr.Top = $Global:InUseTop ; $Global:lblSysGenAddr.Left = 40; $Global:lblSysGenAddr.Width=150 ;$Global:lblSysGenAddr.AutoSize = $true 
        $form.Controls.Add($Global:lblSysGenAddr)    # Add to Form 
        # 
        $Global:txtSysGenAddr = New-Object Windows.Forms.TextBox
        $Global:txtSysGenAddr.ReadOnly = $True;
        $Global:txtSysGenAddr.TabIndex = 50 # set Tab Order 
        $Global:txtSysGenAddr.Top = $Global:InUseTop; $Global:txtSysGenAddr.Left = 240; $Global:txtSysGenAddr.Width = 200;  
        $Global:txtSysGenAddr.Text = $SysGenAddr
        $Global:form.Controls.Add($Global:txtSysGenAddr)    # Add to Form

    $Global:InUseTop = $Global:InUseTop + 30
    ## InUse Address
    $Global:lblInUseAddr = New-Object System.Windows.Forms.Label   
        $Global:lblInUseAddr.Text = "Email Address In Use:"
        $Global:lblInUseAddr.ForeColor = "Red";
        $Global:lblInUseAddr.Top = $Global:InUseTop ; $Global:lblInUseAddr.Left = 40; $Global:lblInUseAddr.Width=150 ;$Global:lblInUseAddr.AutoSize = $true 
        $form.Controls.Add($Global:lblInUseAddr)    # Add to Form 
        # 
        $Global:txtInUseAddr = New-Object Windows.Forms.TextBox
        $Global:txtInUseAddr.ReadOnly = $True;
        $Global:txtInUseAddr.TabIndex = 50 # set Tab Order 
        $Global:txtInUseAddr.Top = $Global:InUseTop; $Global:txtInUseAddr.Left = 240; $Global:txtInUseAddr.Width = 200;  
        $Global:txtInUseAddr.Text = $AddrInUse
        $Global:form.Controls.Add($Global:txtInUseAddr)    # Add to Form

    $Global:InUseTop = $Global:InUseTop + 30
    ## Primary Address
    $Global:lblPrimaryAddr = New-Object System.Windows.Forms.Label   
        $Global:lblPrimaryAddr.Text = "Available Primary Address:"  
        $Global:lblPrimaryAddr.Top = $Global:InUseTop ; $Global:lblPrimaryAddr.Left = 40; $Global:lblPrimaryAddr.Width=250 ;$Global:lblPrimaryAddr.AutoSize = $true 
        $form.Controls.Add($Global:lblPrimaryAddr)    # Add to Form 
        # 
        $Global:txtPrimaryAddr = New-Object Windows.Forms.TextBox
        $Global:txtPrimaryAddr.TabIndex = 50 # set Tab Order 
        $Global:txtPrimaryAddr.Top = $Global:InUseTop; $Global:txtPrimaryAddr.Left = 240; $Global:txtPrimaryAddr.Width = 200;  
        $Global:txtPrimaryAddr.Text = $EmailAddress
        $Global:form.Controls.Add($Global:txtPrimaryAddr)    # Add to Form
        $Global:InputFocus = $Global:txtPrimaryAddr
    $Global:Top = $Global:Top + 40

    Add-FormStandardButtons
}


$Filename       = "EnableUserConfigMbx"
$LogFile		= "e:\SDAP\EnableUserConfigMbx\Log\Log-" + $FileName + ".log"
$RoutingDomain	= "@global.ul.com"
$PrimaryMail    = "@ul.com"
$DC             = "usnbkadds001p.global.ul.com"
$AddrFile       = "e:\o365AdminShared\Data\WelcomeSent.csv"
$uDate          = get-date -uformat %D
$uTime          = get-date -uformat %T
$Date           = $uDate.Replace("/", "-")
$Time           = $uTime.Replace(":", "")
$NoLicRpt       = "E:\SDAP\EnableUserConfigMbx\Log\NoLic" + "-" + "Date" + $Date + ".Log"
$BadAcct        = "Y"
$nonASCII = "[^\x00-\x7F]"

Enter-EmpNoInputForm
Publish-form

If (($Global:Result -eq "OK") -and ([DateTime]$Global:txtDatePicker.Text -le [DateTime]((get-date).AddDays(14)).ToString("dd/MMM/yyyy")))
{
    Invoke-Expression -command E:\O365AdminShared\Scripts\O365DisabledLicenseFeatures.ps1
    $ADExists = [bool]($Global:strDN = GetUserDN $Global:txtInpEmpNo.Text -ErrorAction SilentlyContinue)
    If ($ADExists -eq $True)
    {
        $ADUsr = Get-ADUser $Global:txtInpEmpNo.Text
        If ($ADUsr.Enabled -eq $true)
        {
            $E5Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPREMIUM"}
            $E5Free = [bool] ($E5Lic.ConsumedUnits -lt $E5Lic.ActiveUnits)
            If ($E5Free -eq $True)
            {
                $ReportFile	= "E:\SDAP\EnableUserConfigMbx\Report\Report-" + $FileName +"-EmpNo" + $Global:txtInpEmpNo.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
                $LDAPFilter = "(userPrincipalName=" + $Global:txtInpEmpNo.Text + "@global.ul.com" + ")"
                $EmpNo = $Global:txtInpEmpNo.Text + "@global.ul.com"
                $ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,company,manager,targetaddress,proxyaddresses,extensionattribute1,extensionattribute5,extensionattribute8,extensionattribute15,msExchHideFromAddressLists
                If (($ADUser.sn -cmatch $nonASCII) -or ($ADUser.givenName -cmatch $nonASCII) -or ($ADUser.initials -cmatch $nonASCII))
                {
                    write-host "There are non-US based characters in the user name on the AD Account these need to be corrected before we can proceed" -ForegroundColor Red
                    $LineToWrite = $RecordEvent + "INFO" + "`tNon-US Based characters found in the AD Username for: " + $Global:txtInpEmpNo.Text + " - " + $ADUser.Name
                    writeLogEvent
                    WriteReportEvent
                    $Output = $wshell.Popup("Non-US Based characters found in the AD Username for " + $Global:txtInpEmpNo.Text + " unable to continue.",0,"Non-US Based Characters",0+32)
                }
                else
                {
                    $strUsrPath = [string]::format("LDAP://{0}", $ADUser.DistinguishedName)
                    $u = (new-object System.DirectoryServices.DirectoryEntry($strUsrPath))
                    $LineToWrite = $RecordEvent + "STAR" + "`tEnable User/Assign License/ConfigMailbox script has started"
		            WriteLogEvent
		            WriteReportEvent
		            $LineToWrite = $RecordEvent + "INFO" + "`tLaunched by: " + $WhoAmI
                    WriteLogEvent
		            WriteReportEvent
                    write-host "`nStarting Enablement Process for: " $Global:txtInpEmpNo.Text " " $ADUser.Name -ForegroundColor Cyan
                    $LineToWrite = $RecordEvent + "INFO" + "`tStarting Enablement Process for: " + $Global:txtInpEmpNo.Text + " - " + $ADUser.Name
                    writeLogEvent
                    WriteReportEvent

                    If ($ADUser.Company -eq "UL.ORG")
                    {
                        $PrimaryMail = "@ul.org"
                    }
                    else
                    {
                        #Look to see if the managers email is UL.com or ul.org
                        If (($ADUser.Manager.Length -gt 0) -and ($ADUser.Company.Length -eq 0))
                        {
                            $strMgrPath = [string]::format("LDAP://{0}", $ADUser.Manager)
                            $MgrMail = (new-object System.DirectoryServices.DirectoryEntry($strMgrPath))

                            $Output = $wshell.Popup("No business unit found for this user, setting address based on the managers BU.",0,"Email Domain",0+32)
                            $MgrAddress = "Yes"
                            If ($MgrMail.Company -eq "UL.ORG")
                            {
                                $PrimaryMail = "@ul.org"
                            }
                        }
                        else
                        {
                            $MgrAddress = "No"
                            If (($ADUser.Company -eq $Null) -or ($ADUser.Company -eq ""))
                            {
                                $PrimaryEmail = "@CannotRegister.com"
                                Build-NoBusinessUnitForm
                                Publish-Form
            
                                If (($Global:Result -eq "OK") -and ($Global:chkunknown.Checked -ne "Checked"))
                                {
                                    If ($Global:chkulcom.Checked -eq "Checked")
                                    {
                                        $PrimaryMail = "@ul.com"
                                        $LineToWrite = $RecordEvent + "INFO" + "`tAgent selected to continue registration using UL.ORG for " + $Global:txtInpEmpNo.Text
                                        WriteReportevent
                                    }

                                    If ($Global:chkulorg.Checked -eq "Checked")
                                    {
                                        $PrimaryMail = "@ul.org"
                                        $LineToWrite = $RecordEvent + "INFO" + "`tAgent selected to continue registration using UL.COM for " + $Global:txtInpEmpNo.Text
                                        WriteReportEvent
                                    }
                                }
                                else
                                {
                                    $Global:Result = "Cancelled"
                                }
                            }
                        }
                    }
        
                    If ($Global:Result -eq "OK")
                    {
                        If ($ADUser.Initials -eq $null)
                        {
                            $EmailAddress = ($ADUser.GivenName + "." + $ADUser.Surname) -replace ("'"," ")
#                            $ProperCaseAddress = ((Get-Culture).textinfo.totitlecase($ADUser.GivenName.ToLower()) + "." + (Get-Culture).textinfo.totitlecase($ADUser.Surname.ToLower()) + $PrimaryMail) -replace (" ","")
                            $ProperCaseAddress = ($ADUser.GivenName + "." + $ADUser.Surname) -replace ("'"," ")
                        }
                        else
                        {
                            $EmailAddress = ($ADUser.GivenName + "." + $ADUser.Initials + "." + $ADUser.Surname) -replace ("'"," ")
#                            $ProperCaseAddress = ((Get-Culture).textinfo.totitlecase($ADUser.GivenName.ToLower()) + "." + $ADUser.Initials + "." + (Get-Culture).textinfo.totitlecase($ADUser.Surname.ToLower()) + $PrimaryMail) -replace (" ","")
                            $ProperCaseAddress = ($ADUser.GivenName + "." + $ADUser.Initials + "." + $ADUser.Surname) -replace ("'"," ")
                        }
                        $ProperCaseAddress = ((Get-Culture).textinfo.totitlecase($ProperCaseAddress))

                        If ($EmailAddress -cne $ProperCaseAddress)
                        {
#                           $EmailAddress = $ProperCaseAddress -Replace '[ &/,$#_-]',''
                            $EmailAddress = ($ProperCaseAddress -Replace '[ &/,$#_]','')
                        }
                        else
                        {
#                           $EmailAddress = $EmailAddress -Replace '[ &/,$#_-]',''
                            $EmailAddress = $EmailAddress -Replace '[ &/,$#_]',''
                        }
                        $SysGenAddr = ($EmailAddress -Replace(" ","")) + $PrimaryMail
                        $EmailAddress = ($EmailAddress -Replace(" ","")) + $PrimaryMail
                        $ProperCaseAddress = ($ProperCaseAddress -Replace(" ","")) + $PrimaryMail

                        ##check to see if email address is in use

                        $MBXExists = [bool](get-mailbox $EmailAddress -ErrorAction SilentlyContinue)
                        $Cnt = 0
                        $AddrInuse = "N"
                        $MbxAliasMatch = "N"
                        $MatchENo = "*" + $Global:txtInpEmpNo.Text + "*"
                        If ((get-mailbox $EmailAddress -ErrorAction SilentlyContinue).WindowsLiveID -like $MatchENo)
                        {
                            $MbxAliasMatch = "Y"
                        }
                        If (($MBXExists -eq $True) -and ((get-mailbox $EmailAddress).Alias -ne $Global:txtInpEmpNo.Text) -and ($MbxAliasMatch -eq "N"))
                        {
                            write-host "The Address" $EmailAddress "is in use by employee number" (get-mailbox $EmailAddress -ErrorAction SilentlyContinue).Alias -ForegroundColor Red
                            $AddrInuse = $EmailAddress
                            If ($ADUser.Initials -eq $null)
                            {
                                $Middle = "X"
                                $EmailAddress = (Get-Culture).textinfo.totitlecase($ADUser.GivenName.ToLower()) + "." + $Middle + "." + (Get-Culture).textinfo.totitlecase($ADUser.Surname.ToLower()) + $PrimaryMail
                                Do
                                {
                                    $MBXExists = [bool](get-mailbox $EmailAddress -ErrorAction SilentlyContinue)
                                    If ($MBXExists -eq $True)
                                    {
                                        $Cnt++
                                        $EmailAddress = (Get-Culture).textinfo.totitlecase($ADUser.GivenName.ToLower()) + "." + $Middle + $Cnt + "." + (Get-Culture).textinfo.totitlecase($ADUser.Surname.ToLower()) + $PrimaryMail
                                        write-host "Checking if " $EmailAddress "is in use" -ForegroundColor Red
                                    }
                                } while ($MbxExists -eq $True)
                            }
                            else
                            {
                                $Middle = $ADUser.Initials
                                Do
                                {
                                    $Cnt++
                                    $EmailAddress = (Get-Culture).textinfo.totitlecase($ADUser.GivenName.ToLower()) + "." + $Middle + $Cnt + "." + (Get-Culture).textinfo.totitlecase($ADUser.Surname.ToLower()) + $PrimaryMail
                                    $MBXExists = [bool](get-mailbox $EmailAddress -ErrorAction SilentlyContinue)
                                    write-host "Checking if " $EmailAddress "is in use" -ForegroundColor Red
                                } while ($MbxExists -eq $True)
                            }

                            $AvailableAddress = $EmailAddress
                            Do
                            {
                                $EmailAddress = $AvailableAddress
                                Review-AddrInUse
                                Publish-Form
                                $EmailAddress = $Global:txtPrimaryAddr.Text
                                $AddrInuse = $Global:txtPrimaryAddr.Text
                                $MBXExists = [bool](get-mailbox $EmailAddress -ErrorAction SilentlyContinue)
                            }while (($MBXExists -eq $True) -and ($Global:Result -eq "OK"))
                        }
                        else
                        {
                            If ((get-mailbox $EmailAddress -ErrorAction SilentlyContinue).Alias -ne (($Global:txtInpEmpNo.Text)) -and ($EmailAddress.Substring(0,$EmailAddress.indexof("@"))))
                            {
                                Build-ConfirmAddress
                                Publish-Form

                                If ($Global:Result -eq "OK")
                                {
                                    $BadAcct = "N"
                                    If ($Global:chksysGen.Checked -eq "Checked")
                                    {
                                        $SysGenAddr = $Global:txtSysGenAddr.Text
                                    }
                                    else
                                    {
                                        $SysGenAddr = $Global:txtProperCase.Text
                                    }

                                    If (([bool]((get-mailbox $sysGenAddr -ErrorAction SilentlyContinue)) -ne $False) -and ($MbxAliasMatch -ne "Y"))
                                    {
                                        Do
                                        {
                                            $EmailAddress = $SysGenAddr
                                            $AddrInuse = "Y"
                                            Review-AddrInUse
                                            Publish-Form
                                            $EmailAddress = $Global:txtPrimaryAddr.Text
                                            $AddrInuse = $Global:txtPrimaryAddr.Text
                                            $MBXExists = [bool](get-mailbox $EmailAddress -ErrorAction SilentlyContinue)
                                        }while (($MBXExists -eq $True) -and ($Global:Result -eq "OK"))
                                    }
                                    else
                                    {
                                        #Added this to try and resolve not picking up a new address being entered
                                        $EMailAddress = $SysGenAddr
                                    }
                                }
                            }
                        }

                        If ($Global:Result -eq "OK")
                        {
		                    $LineToWrite = $RecordEvent + "INFO" + "`t" + $ADUser.Name + "(" + $EmailAddress + ")`t" + $DC
                            WriteReportEvent
                            $LineToWrite = $RecordEvent + "INFO" + "`t           Ticket Number: " + $Global:txtInpTicketNo.Text
                            WriteReportEvent
                            $LineToWrite = $RecordEvent + "INFO" + "`t              Start Date: " + $Global:txtDatePicker.Text + "`n"
                            WriteReportEvent
                            If ($ADUser.Initials.Length -eq 0)
                            {
                                $LineToWrite = $RecordEvent + "INFO" + "`t      FirstName LastName:     " + $ADUser.givenName + " " + $ADUser.sn
                            }
                            else
                            {
                                $LineToWrite = $RecordEvent + "INFO" + "`t   FirstName MI LastName:     " + $ADUser.givenName + " " + $ADUser.Initials + ". " + $ADUser.sn
                            }
		                    WriteReportEvent
		                    $LineToWrite = $RecordEvent + "INFO" + "`t    Current Mail Address:     " + $ADUser.Mail
		                    WriteReportEvent
		                    $LineToWrite = $RecordEvent + "INFO" + "`t           Employee Type:     " + $ADUser.extensionAttribute1 + "`n"
		                    WriteReportEvent
    	                    $LineToWrite = $RecordEvent + "INFO" + "`tAD Account values prior to mail enabling account:"
		                    WriteReportEvent
		                    $LineToWrite = $RecordEvent + "INFO" + "`t       User Mail Address:     " + $ADUser.Mail
		                    WriteReportEvent
		                    $LineToWrite = $RecordEvent + "INFO" + "`t     User Target Address:   " + $ADUser.targetaddress
		                    WriteReportEvent
		                    $LineToWrite = $RecordEvent + "INFO" + "`t    User Proxy Addresses:  " + ($ADUser.proxyaddresses -join "`n`t`t`t`t`t`t`t`t`t")
		                    WriteReportEvent
		                    $LineToWrite = $RecordEvent + "INFO" + "`t     ExtensionAttibute15:   " + $ADUser.extensionattribute15
		                    WriteReportEvent
		                    $LineToWrite = $RecordEvent + "INFO" + "`t  Add Forwarding Address:   " + $Global:txtForwarder.Text
		                    WriteReportEvent

                            #check to see if account is already mail enabled
#                            Remove-PSSession (Get-PSSession)
                            Disconnect-ExchangeOnline -Confirm:$False
                            $ErrorActionPreference = "SilentlyContinue"  #Temporarily surpress error events
                            write-host "This user will be assigned" $EmailAddress "as the primary email address" $ADUser.Name -ForegroundColor Green
                            $MEnab = ([bool](Get-MailUser $EmpNo -ErrorAction SilentlyContinue))
                            If ($ADUser.displayName.contains(",") -eq $false)
                            {
			                    If ($ADUser.initials.length -gt 0)
			                    {
			                        $ADName = $ADUser.sn + ", " + $ADUser.givenName + " " + $ADUser.initials.substring(0,1) + "."
			                    }
                                else
                                {
				                    $ADName = $ADUser.sn + ", " + $ADUser.givenName
			                    }
    	                    }

                            write-host "Enabling MailUser and Assigning Licenses for" $ADUser.Name -ForegroundColor Green
                            If ($MEnab -ne $True)
                            {
                                Enable-MailUser $EmpNo -Alias $Global:txtInpEmpNo.Text -ExternalEmailaddress $EmailAddress -PrimarySmtpAddress $EmailAddress -DomainController:$DC
                                $LineToWrite = $RecordEvent + "INFO" + "`t     Mail Enabled the Account"
      		                    WriteReportEvent
                            }
                            else
                            {
                                write-host "This account is already mail enabled" -ForegroundColor Red
			                    $LineToWrite = $RecordEvent + "INFO" + "`t     Account is Already Mail Enabled"
      		                    WriteReportEvent
                            }
			
                            #Start Check for Internet Mail = @ul.com or @ul.org
                            $Name = ($EmailAddress.Substring(0,($EmailAddress.IndexOf("@"))))
#                            Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:TEAMS_COMMERCIAL_TRIAL" -erroraction SilentlyContinue
                            $ErrorActionPreference = "Continue" #Re-enable error events
                            If ($Emailaddress -contains $Name)
                            {
            #                    Do nothing, coded this way as the -notcontains does not work in as expected
                                $LineToWrite = $RecordEvent + "INFO" + "`t     Set-Mailuser Not executed as the $Emailaddress contains $name"
      		                    WriteReportEvent
                            }
                            else
                            {
                                $LineToWrite = $RecordEvent + "INFO" + "`t     Name: " + $Name
   		                        WriteReportEvent
                                $LineToWrite = $RecordEvent + "INFO" + "`t     Routing Domain: " + $RoutingDomain
   		                        WriteReportEvent
                                $LineToWrite = $RecordEvent + "INFO" + "`t     Domain Controller: " + $DC
   		                        WriteReportEvent
                                $RoutingAddress = ($Name + $RoutingDomain)

                                if ($EmailAddress -like "*@ul.*")
                                {
                                    $LineToWrite = $RecordEvent + "INFO" + "`t     Set-Mailuser has @ul. in the name"
      		                        WriteReportEvent
#                                    Set-MailUser $EmpNo -EmailAddresses (((Get-MailUser $EmpNo -DomainController:$DC).EmailAddresses)+=($Name + $RoutingDomain)) -CustomAttribute15 "EnableMailUser PS Date: $Date PS Time: $Time" -DomainController:$DC
                                    Set-MailUser $EmpNo -EmailAddresses (((Get-MailUser $EmpNo -DomainController:$DC).EmailAddresses)+=($Name + $RoutingDomain)) -DomainController:$DC
                                }
                                else
                                {
                                    $LineToWrite = $RecordEvent + "INFO" + "`t     Set-Mailuser does not have @ul. in the name"
      		                        WriteReportEvent
#                                    Set-MailUser $EmpNo -EmailAddresses (((Get-MailUser $EmpNo -DomainController:$DC).EmailAddresses)+=$EmailAddress,($Name + $RoutingDomain)) -CustomAttribute15 "EnableMailUser PS Date: $Date PS Time: $Time" -DomainController:$DC
                                    Set-MailUser $EmpNo -EmailAddresses (((Get-MailUser $EmpNo -DomainController:$DC).EmailAddresses)+=$EmailAddress,($Name + $RoutingDomain)) -DomainController:$DC
                                }
                                Set-MailUser $EmpNo -CustomAttribute15 "EnableMailUser PS Date: $Date PS Time: $Time" -DomainController:$DC
                            }
#                            $ErrorActionPreference = "Continue" #Re-enable error events

                            If ($ADUser.displayName.contains(",") -eq $false)
                            {
	                            Set-ADUser $ADuser -DisplayName $ADName
		                        $LineToWrite = $RecordEvent + "INFO" + "`tChanged displayName to: " + $ADName
		                        WriteReportEvent
	                        }

                            #Set UsageLocation, Assign Free PowerBI License and enable Standard Email License for this Employee Type
                            Set-MsolUser -UserPrincipalName $EmpNo -UsageLocation US -ErrorAction Silentlycontinue

                            $HasE5 = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPREMIUM"})
                            $HasEMS = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:EMSPREMIUM"})
	                        $HasPBIF = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:Power_BI_Standard"})

                            $E5Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPREMIUM"}
                            $EMSLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EMSPREMIUM"}

	                        if ($E5Lic.ConsumedUnits -ge $E5Lic.ActiveUnits)
	                        {
		                        If (Test-Path $NoLicRpt)
		                        {
                                    #Do Nothing File Exists and an email was already sent to the team
			                        $LineToWrite = $RecordEvent + "INFO" + "`t     Assigned " + $LicType + " license to " + $Global:txtInpEmpNo.Text + " - " + $Name
			                        Out-File -FilePath $NoLicRpt -InputObject $LineToWrite -Append -NoClobber
			                        $LineToWrite = "E5 Licenses Assigned " + (Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}).ConsumedUnits + " E5 License Allocation " +(Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPREMIUM"}).ActiveUnits
			                        Out-File -FilePath $NoLicRpt -InputObject $LineToWrite -Append -NoClobber
		                        }
		                        else
		                        {
                                    #Location to write file for tracking the frequency that no E4 licenses are available
                                    #Also prevents this message from being sent more than 1 time per day
			                        Out-File -FilePath $NoLicRpt -InputObject "No E5 Licenses Available"
			                        $LineToWrite = ""
			                        Out-File -FilePath $NoLicRpt -InputObject $LineToWrite -Append -NoClobber
			                        $LineToWrite = "Assigned " + $LicType + " license to " + $Global:txtInpEmpNo.Text + " - " + $EmailAddress
    		                        Out-File -FilePath $NoLicRpt -InputObject $LineToWrite -Append -NoClobber
			                        $LineToWrite = "E5 Licenses Assigned " + (Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}).ConsumedUnits + " E5 License Allocation " + (Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPREMIUM"}).ActiveUnits
			                        Out-File -FilePath $NoLicRpt -InputObject $LineToWrite -Append -NoClobber
                                    #Send email to notify appropriate individuals that no E5 licenses are available
                                    $E5Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPREMIUM"}				
			                        $server = "smtp-relay.ul.com"
			                        $client = new-object system.net.mail.smtpclient $server
			                        $from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "UL Account Provisioning Process"           
			                        $to = $from 
			                        $SendTo = "LST.O365AdminTeam@ul.com,BJ.Stone@ul.com"#,Ginger.M.Tucibat@ul.com"
			                        $message = new-object  System.Net.Mail.MailMessage $from, $to 
			                        $message.IsBodyHtml = $true
			                        $msgfont = "<basefont face=verdana size=2.5 color=black>"
			                        $message.Subject = "No Enterprise E5 Licenses Available"
			                        $message.To.Clear()
			                        $message.Body = $msgfont + "<p>Person executing script: " + $whoami + "<p>Against user account:  " + $EmpNo + "<p>You are receiving this message because there are no Enterprise E5 licenses available.  <ul type=""disc""><li>E5 Licenses Active/Consumed/Available:  " + $E5Lic.ActiveUnits + "/" + $E5Lic.ConsumedUnits + "/" + ($E5Lic.ActiveUnits-$E5Lic.ConsumedUnits)
			                        $message.To.Add($SendTo)
			                        $client.Send($message)				
		                        }
	                        }
 

                            If ($HasEMS -eq $False)
                            {

                                If ($ADUser.ExtensionAttribute1 -like "Employee*")
                                {
                                    write-host "Assigning standard License Set for individual who are Employees" -ForegroundColor Green
                                    Add-AzureADGroupMember -ObjectId 2aab81de-e54a-4c58-909c-b15273c54dff -RefObjectId (get-msoluser -UserPrincipalName $EmpNo).ObjectID
                                    $LicType = "E5 Employee"
                                }
                                else
                                {
                                    If ($ADUser.ExtensionAttribute1 -notlike "Ex-*")
                                    {
                                        write-host "Assigning standard license set for individuals who are Non-Employees" -ForegroundColor Green

                                    }
                                    else
                                    {
                                        write-host "Assigning standard Non-Employee license set however the type listed shows they are an Ex- staff member" -ForegroundColor Red
                                    }
                                    Add-AzureADGroupMember -ObjectId 40dff561-964a-4f7b-9794-dba3542b63c4 -RefObjectId (get-msoluser -UserPrincipalName $EmpNo).ObjectID
                                    $LicType = "E5 Non-Employee"
                                }
                                $LineToWrite = $RecordEvent + "ASSG" + "`t     Assigned " + $LicType + " License"
                                WriteReportEvent
                            }
                            else
                            {
                            #might need to add check that the person is in the right group.
                                write-host "E5 License Already assigned to this account" -ForegroundColor Red
                                $LineToWrite = $RecordEvent + "INFO" + "`t     E5 License Already Assigned"
		                        WriteReportEvent
                            }
                                       
                            if ($HasPBIF -eq "True")
                            {
		                        $LineToWrite = $RecordEvent + "INFO" + "`t     PowerBI Free License Already Assigned"
                                WriteReportEvent	
	                        }
	                        else
	                        {
		                        Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicenses "ul:POWER_BI_STANDARD" -erroraction SilentlyContinue
                                $HasPBIF = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:Power_BI_Standard"})
                                If ($HasPBIF -eq "True")
                                {
                                    $LineToWrite = $RecordEvent + "ASSG" + "`t     Assigned PowerBI Free License"
                                }
                                else
                                {
		                            $LineToWrite = $RecordEvent + "ASSG" + "`t     Unable to Assign PowerBI Free License"
                                    Write-Host "Unable to PowerBI Free License" -ForegroundColor Red
                                }
		                        WriteReportEvent	            
	                        }

                            If ($ADUser.msExchHideFromAddressLists -eq "True")
                            {
                                write-host "Setting user to be Hidden from the Address Book" -ForegroundColor Red
                                write-host
                                $strUserPath = [string]::format("LDAP://{0}", $ADUser.DistinguishedName)
                                $u = new-object System.DirectoryServices.DirectoryEntry($strUserPath) -ErrorAction SilentlyContinue
                                $u.msExchHideFromAddressLists.value = $False
                                $U.CommitChanges()
                            }

                            $IntuneMem = "off"
                            $ADGroupMem = Get-ADPrincipalGroupMembership $Global:txtInpEmpNo.Text -ResourceContextServer global.ul.com
                            foreach ($ADGroupMem in $ADGroupMem)
                            {
                                if ($ADGroupMem.SamAccountName -like "*IntuneEnfor*")
                                {
                                    $IntuneMem = "1"
                                    $IntuneGroup = "Previously Assigned"
                                }
                            }
                            If ($IntuneMem -eq "off")
                            {
                                If ($ADUser.ExtensionAttribute5 -eq "Europe")
                                {
                                    Add-ADGroupMember -Identity "ACL.EU.IntuneEnforced" -Members $Global:txtInpEmpNo.Text -Confirm:$False
                                    write-host  "Adding to the ACL.EU.IntuneEnforced group" -ForegroundColor Green
                                    $LineToWrite = $RecordEvent + "ADD " + "`t     Adding to the ACL.EU.IntuneEnforced group"
		                            WriteReportEvent
                                    $IntuneGroup = "ACL.EU.IntuneEnforced"
                                }
                                elseif (($ADUser.ExtensionAttribute5 -eq "Asia Pacific") -or ($ADUser.ExtensionAttribute5 -eq "APAC & MEA"))
                                        {
                                            Add-ADGroupMember -Identity "ACL.AP.IntuneEnforced" -Members $Global:txtInpEmpNo.Text -Confirm:$False
                                            write-host  "Adding to the ACL.AP.IntuneEnforced group" -ForegroundColor Green
                                            $LineToWrite = $RecordEvent + "ADD " + "`t     Adding to the ACL.AP.IntuneEnforced group"
		                                    WriteReportEvent
                                            $IntuneGroup = "ACL.AP.IntuneEnforced"
                                        }
                                        else
                                        {
                                            Add-ADGroupMember -Identity "ACL.UL.IntuneEnforced" -Members $Global:txtInpEmpNo.Text -Confirm:$False
                                            write-host  "Adding to the ACL.UL.IntuneEnforced group" -ForegroundColor Green
                                            $LineToWrite = $RecordEvent + "ADD " + "`t     Adding to the ACL.UL.IntuneEnforced group"
                                            WriteReportEvent
                                            $IntuneGroup = "ACL.UL.IntuneEnforced"
                                        }
                            }
			                else
	                        {
			                    $LineToWrite = $RecordEvent + "INFO" + "`t     Already a member an IntuneEnforced group"
				                WriteReportEvent				
			                }

                            write-host "Checking to see if this individual is already a member of the MFA_Enabled Group (this may take a few minutes)" -ForegroundColor Cyan
                            $MFAMem = [bool](get-AzureADGroupMember -ObjectId abf2ba7d-50e2-4ca6-99c8-2e9c245cfed1 -all $true |Where-Object {$_.UserPrincipalName -eq $EmpNo})
                            If ($MFAMem -eq $False)
                            {
                                write-host "Adding user to the MFA_Enabled Group" -ForegroundColor Green
                                Add-AzureADGroupMember -ObjectId abf2ba7d-50e2-4ca6-99c8-2e9c245cfed1 -RefObjectId ((get-msoluser -UserPrincipalName $EmpNo).objectid)
                                $LineToWrite = $RecordEvent + "ADDMEM" + "`t     Added to MFA_Enabled Group"
        	                    WriteReportEvent
                            }
                            else
                            {
                                write-host "Already a member of the MFA_Enabled Group" -ForegroundColor Red
                                $LineToWrite = $RecordEvent + "NOCHG" + "`t     Already MFA_Enabled Group Member"
        	                    WriteReportEvent
                            }

                            write-host "Disabling RemotePowerShell" -ForegroundColor Green
                            $DisFile = "E:\SDAP\DisableRemotePS\Input\DisableRemotePS.csv"
                            $Global:txtInpEmpNo.Text | out-file -filepath $DisFile -Append
#                            Set-User $Global:txtInpEmpNo.Text -RemotePowerShellEnabled $False -Confirm:$False
                            $LineToWrite = $RecordEvent + "DISA" + "`t     Added user to the DisableRemotePS.csv file"
    	                    WriteReportEvent

                            write-host "Enabling/Configuring Teams" -ForegroundColor Green
                            $sipAddress = "sip:" + $EmailAddress
                            $LineToWrite = $RecordEvent + "SET " + "`t     Enabling Teams"
    	                    WriteReportEvent                
                            Set-ADUser $ADuser -Replace @{'msRTCSIP-PrimaryUserAddress'=$sipAddress}
                            Set-ADUser $ADuser -Replace @{'msRTCSIP-DeploymentLocator'="sipfed.online.lync.com"}
                            Set-ADUser $ADuser -Replace @{'msRTCSIP-FederationEnabled'="TRUE"}
                            Set-ADUser $ADuser -Replace @{'msRTCSIP-InternetAccessEnabled'="TRUE"}
                            Set-ADUser $ADuser -Replace @{'msRTCSIP-UserEnabled'="TRUE"}

                            write-host "Reconnecting to O365 in order to complete the Process" -ForegroundColor Yellow
#                            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential (Import-Clixml $Global:CredFile) -Authentication Basic -AllowRedirection
#                            Import-PSSession $session -AllowClobber
                            Connect-ExchangeOnline -Credential $Global:LiveCred
                            write-host "Starting Mailbox Configuration Process for: " $Global:txtInpEmpNo.Text -ForegroundColor Cyan

                            $cnt = 0
                            write-host "Waiting for mailbox creation to complete, this may take several minutes please be patient.." -ForegroundColor Yellow -NoNewline
                            Do
                            {
                                $cnt++
                                write-host ".." -ForegroundColor Yellow -NoNewline
                                start-sleep -Seconds 15
                                $MBXExists = [bool](Get-Mailbox $ADUser.UserPrincipalName -ErrorAction SilentlyContinue)
                            } while (($MBXExists -eq $False) -and ($cnt -lt 30))

                            If ($MBXExists -eq $True)
                            {
                                write-host  "`nSetting Mailbox Quota and active 3 yr retention policy" -ForegroundColor Green
                                write-host "Setting mailbox retention policy" -ForegroundColor Green
                                Set-MailBox $ADUser.UserPrincipalName -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL MRM Policy - 3 yr Delete" -RetentionHoldEnabled $false
                                write-host "Setting mailbox quotas" -ForegroundColor Green
                                Set-Mailbox $ADUser.UserPrincipalName -UseDatabaseQuotaDefaults $False -IssueWarningQuota 45GB -ProhibitSendQuota 49.75GB -ProhibitSendReceiveQuota 50GB
                                write-host "Setting mailbox protocols" -ForegroundColor Green
                                Set-CASMailBox $ADUser.UserPrincipalName -ImapEnabled $false -PopEnabled $false -ActiveSyncEnabled $false

                                If ($Global:txtForwarder.Text -like "*@*")
                                {
                                    write-host "Setting mailbox forwarding address on mailbox" -ForegroundColor Green
                                    Set-Mailbox $ADUser.UserPrincipalName -ForwardingSMTPAddress $Global:txtForwarder.Text
                                    $LineToWrite = $RecordEvent + "ADD   " + "`t     Added Forwarding Address of " + $Global:txtForwarder.Text + "`n"
        	                        WriteReportEvent
                                }
  	        
                                $ADUserPost = Get-ADUser $Global:txtInpEmpNo.Text -Properties *

                                If (($ADUserPost.mail.length -le 6) -or ($ADUserPost.targetaddress.length -le 6))
                                {
                                    $Output = $wshell.Popup("Email Addresses were not configured properly contact the O365 Admin Team " + $Global:txtInpEmpNo.Text + ".",0,"Configuration Failed",0+32)
                                    write-host "Email Addresses were not configured properly contact the O365 Admin Team: " $Global:txtInpEmpNo.Text -ForegroundColor Red
                                    $LineToWrite = $RecordEvent + "INFO" + "`tEmail Addresses were not configured properly contact the O365 Admin Team:" + $Global:txtInpEmpNo.Text
                                    WriteReportEvent
                                    $LineToWrite = $RecordEvent + "INFO" + "`tAD Account values contain errors after mail enabling account:"
                                }
                                else
                                {
           	                        $LineToWrite = $RecordEvent + "INFO" + "`tAD Account values after mail enabling account:"
                                }
		                        WriteReportEvent

                                If ($MgrAddress -eq "No")
                                {
	                                $LineToWrite = $RecordEvent + "INFO" + "`t     Business Unit: " + $ADUserPost.Company
                                }
                                else
                                {
                                    $LineToWrite = $RecordEvent + "NOBU" + "`t     Business Unit: Not Configured/Determined based on Managers Address"
                                }
                                WriteReportEvent
		                        $LineToWrite = $RecordEvent + "INFO" + "`t     User Mail Address:     " + $ADUserPost.mail
		                        WriteReportEvent
		                        $LineToWrite = $RecordEvent + "INFO" + "`t     User Target Address:   " + $ADUserPost.targetaddress
		                        WriteReportEvent
		                        $LineToWrite = $RecordEvent + "INFO" + "`t     User Proxy Addresses:  " + ($ADUserPost.proxyaddresses -join "`n`t`t`t`t`t`t`t`t`t")
		                        WriteReportEvent
		                        $LineToWrite = $RecordEvent + "INFO" + "`t     ExtensionAttibute15:   " + $ADUserPost.extensionattribute15
		                        WriteReportEvent
  
                                $MbxDetails = Get-Mailbox $ADUser.UserPrincipalName
                                $LineToWrite = $RecordEvent + "INFO" + "`t     User Principal Name: " + $MBXDetails.UserPrincipalName
	                            WriteReportEvent
	                            $LineToWrite = $RecordEvent + "INFO" + "`t     Primary Email Address: " + $MBXDetails.PrimarySmtpAddress
	                            WriteReportEvent
	                            $LineToWrite = $RecordEvent + "INFO" + "`t     SIP Address: " + $ADUserPost."msRTCSIP-PrimaryUserAddress"
	                            WriteReportEvent                
                                $LineToWrite = $RecordEvent + "INFO" + "`t     SMTP Forwarding Address set to " + $MBXDetails.ForwardingSMTPAddress
                                WriteReportEvent
	                            $LineToWrite = $RecordEvent + "INFO" + "`t     Region: " + $MBXDetails.CustomAttribute5
	                            WriteReportEvent
                                $LineToWrite = $RecordEvent + "SET " + "`t     Retention Policy set to " + $MBXDetails.RetentionPolicy
                                WriteReportEvent

                                #Send Welcome Message to Mailbox
                                $MsgSent = import-csv "e:\o365AdminShared\Data\WelcomeSent.csv"
                                $NotSent = "Y"
                                foreach ($MsgSent in $MsgSent)
                                {
                                    If ($MsgSent.UPN -eq $MBXDetails.UserPrincipalName)
                                    {
                                        $NotSent="N"
                                    }
                                }
  
                                $client = new-object system.net.mail.smtpclient $server
                                $from = New-Object System.Net.Mail.MailAddress "DoNotReply@ul.com" , "UL Technology Services"
                                $to = $from 
                                $message = new-object  System.Net.Mail.MailMessage $from, $to 
                                $message.IsBodyHtml = $true
                                $message.To.Clear()
                                $message.CC.Clear()
                                $message.Bcc.Clear()
        
                                $HasEMS = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:EMSPREMIUM"})
                                If ($NotSent -eq "Y")
                                {
                                    Create-WelcomeMsg
                                    $Details = $MBXDetails.UserPrincipalName + "," + $MBXDetails.PrimarySmtpAddress
                                    Add-Content $AddrFile $Details
                                    Write-host "Welcome Message Sent to this Address" -ForegroundColor Cyan
                                    $LineToWrite = $RecordEvent + "MAIL" + "`t     Welcome Message Sent to this Mailbox"
                                    WriteReportEvent
                                    WriteLogEvent
                                }
                                else
                                {
                                    $EMailFrom = New-Object System.Net.Mail.MailAddress "DoNotReply@ul.com" , "UL Technology Services"
                                    $smtpServer = "SMTP-relay.ul.com"
                                    $message.from = $EmailFrom
                                    $message.To.Add("EnterpriseMessagingServices@ul.com")
                                    $message.Body = $MsgBody
                                    $message.IsBodyHTML = $true
                                    $message.Subject = "SDAP Team Running ConfigMailbox Multiple Times for: " + $Global:txtInpEmpNo.Text

                                    Enter-Reason
                                    Publish-Form
                                    $message.Body = "<p>SDAP Team has run the ConfigMailbox script multiple times for: " + $Global:txtInpEmpNo.Text + ", " + $MBXDetails.PrimarySmtpAddress + "<br><p>WhoAmI: " + $whoAmi + "<br><p>Reason: " + $Global:txtReason.Text + "<br>"
                                    $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 25)
                                    $SMTPClient.Send($message)
                                    Write-Host "Welcome Message has already been sent to this Address" -ForegroundColor Red
                                    $LineToWrite = $RecordEvent + "MAIL" + "`t     Welcome message was previously sent to this Mailbox"
	                                WriteReportEvent
                                    WriteLogEvent
                                    $LineToWrite = $RecordEvent + "MAIL" + "`t     Multiple Config Reason: " + $Global:txtReason.Text
	                                WriteReportEvent
                                    WriteLogEvent
                                }

                                Build-AcctEnablePriorDetails
                                Build-AcctEnabledAfterDetails
                                Publish-Form

                                #Write end record to the log file
                                write-host "Enablement and mailbox configuration Process complete`n" -ForegroundColor Cyan
                            }
                            else
                            {
                                write-host "Enablement terminated no mailbox exists`n" -ForegroundColor Red
                            } 
                            $LineToWrite = $RecordEvent + "STOP" + "`tThis instance is stopping for: " + $Global:txtInpEmpNo.Text + " - " + $ADUser.Name + "`n"
                            WriteReportEvent
	                        writeLogEvent
                        }
                        else
                        {
                            If ($Global:Result -ne "Cancel")
                            {
                                $Output = $wshell.Popup("Business Unit Unknown enable user and configure mailbox process cancelled.",0,"Cancelled",0+32)
                                $LineToWrite = $RecordEvent + "INFO" + "`tAgent selected unknown business unit and registration is being cancelled for " + $Global:txtInpEmpNo.Text
                                writeLogEvent
                                WriteReportEvent
                            }
                            else
                            {
		                        $LineToWrite = $RecordEvent + "INFO" + "`t" + $ADUser.Name + "(" + $EmailAddress + ")`t" + $DC
                                WriteReportEvent
                                $LineToWrite = $RecordEvent + "INFO" + "`tTicket Number: " + $Global:txtInpTicketNo.Text
                                WriteReportEvent
                                $LineToWrite = $RecordEvent + "INFO" + "`t  AD Username: " + $ADUser.Name
                                WriteReportEvent
                                If ($ADUser.Initials.Length -eq 0)
                                {
                                    $LineToWrite = $RecordEvent + "INFO" + "`t    FirstName LastName:     " + $ADUser.givenName + " " + $ADUser.sn
                                }
                                else
                                {
                                    $LineToWrite = $RecordEvent + "INFO" + "`t FirstName MI LastName:     " + $ADUser.givenName + " " + $ADUser.Initials + ". " + $ADUser.sn
                                }
		                        WriteReportEvent

                                $Output = $wshell.Popup("Account Enablement and mailbox configuration process cancelled.",0,"Cancelled",0+32)
                                $LineToWrite = $RecordEvent + "INFO" + "`tAccount Enablement and Mailbox Configuration Process Cancelled for:" + $Global:txtInpEmpNo.Text
                                WriteReportEvent
                                WriteLogEvent
                            }
                        }
                    }
                    else
                    {
                        If ($Global:Result -eq "Cancelled")
                        {
                            $Output = $wshell.Popup("Business Unit Unknown enable user and configure mailbox process cancelled.",0,"Cancelled",0+32)
                            $LineToWrite = $RecordEvent + "INFO" + "`tAgent selected unknown business unit and registration is being cancelled for " + $Global:txtInpEmpNo.Text
                            writeLogEvent
                            WriteReportEvent
                        }
                        else
                        {
                            $O365Lic = (Get-MsolUser -UserPrincipalName $Global:txtInpEmpNo.Text).Licenses
                            If ($O365Lic.Count -lt 3)
                            {
                                $Output = $wshell.Popup("This account does not appear to have the standard license set configured.  To check license assignments use the Manage Licenses for a User under the Change Activities of the SDAP Menu.",0,"Check Licenses",0+32)
                                $LineToWrite = $RecordEvent + "ERR " + "`tLicenses assigned mailbox not yet created"
                                WriteReportEvent
                                WriteLogEvent
                            }
                            else
                            {
                                $Output = $wshell.Popup("O365 Licenses assigned but mailbox creation has not yet completed for this account.",0,"No Mailbox",0+32)
                                $LineToWrite = $RecordEvent + "ERR " + "`tCheck License Assignment"
		                        WriteReportEvent
                                WriteLogEvent
                            }
                        }
                    }
                }
            }
            else
            {
                $Output = $wshell.Popup("No Email Licenses Available - Please notify the O365 Admin Team.",0,"No Licenses",0+32)
                write-host "No Email Licenses Available - Please notify the O365 Admin Team." -ForegroundColor Red
            }
	    }
        else
        {
            $Output = $wshell.Popup("Account disabled for Emp ID: " + $Global:txtInpEmpNo.Text,0,"Account Disabled",0+32)
        }
    }
    else
    {
        $Output = $wshell.Popup("No Active Directory Account Exists " + $Global:txtInpEmpNo.Text + ".",0,"No Account Found",0+32)
        write-host "No Active Directory Account Exists for: " $Global:txtInpEmpNo.Text -ForegroundColor Red
        $LineToWrite = $RecordEvent + "INFO" + "`tNo Active Directory Account Exists for:" + $Global:txtInpEmpNo.Text
        WriteReportEvent
    }
}
else
{
    If ($Global:txtDatePicker.Text -gt (get-date).AddDays(14))
    {
        $Output = $wshell.Popup("The start date for this user is more than 2 weeks out - Account enablement process is being cancelled.",0,"Cancelled",0+32)
        write-host "The start date for this user is more than 2 weeks out; Enablement process is being cancelled" -ForegroundColor Red
    }
    else
    {
        $Output = $wshell.Popup("Account Enablement and mailbox configuration process cancelled.",0,"Cancelled",0+32)
        $LineToWrite = $RecordEvent + "INFO" + "`tAccount Enablement and Mailbox Configuration Process Cancelled for:" + $Global:txtInpEmpNo.Text
        WriteLogEvent
    }
}