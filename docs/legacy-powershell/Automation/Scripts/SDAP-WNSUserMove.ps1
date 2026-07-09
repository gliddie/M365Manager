#
#  SDAP Reset Retention Policy
#
#  11/09/2021 - SAG - Added new script
#  11/10/2021 - SAG - Added checks for when the mailbox was created and if before 3/23/2020 allows selection of the 1 yr or 3 yr retention policy
#  05/04/2023 - SAG - Added details to log the size of the TotalDeletedItems size.  If 100GB then the Archive must be enabled to resolve issue that we see with not being able to reschedule calendar events
#  04/22/2025 - SAG - Added Modified the 6yr policy to be 3yrArchive/6yrDelete and added a 6 yr policy that is 2yrArchive and 6yrDelete
###########################################################################################

Function Display-CurrentSettings
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Current User Settings"
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 500 ; $Global:form.Height = 470  # Make the form wider 
    
    $Script:Left = 20
    $Script:LeftInput = 150
    $Script:Top = 20

    #User Mailbox
    $Script:lblDispName = New-Object System.Windows.Forms.Label   
        $Script:lblDispName.Text = "User Name:"
        $Script:lblDispName.Top = $Script:Top ; $Script:lblDispName.Left = $Script:Left; $Script:lblDispName.Width=150 ;$Script:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblDispName)    # Add to Form 
    $Script:txDispName = New-Object Windows.Forms.TextBox
        $Script:txDispName.Top = $Script:Top; $Script:txDispName.Left = $Script:LeftInput; $Script:txDispName.Width = 250;
        $Script:txDispName.ReadOnly = $true
        $Script:txDispName.Text = $CurrMbxSettings.DisplayName
        $Global:form.Controls.Add($Script:txDispName)    # Add to Form
    
    $Script:Top = $Script:Top + 30
    #Current Retention Policy (20 Characters)
    $Script:lblCurrLoc = New-Object System.Windows.Forms.Label   
        $Script:lblCurrLoc.Text = "Current Location:"
        $Script:lblCurrLoc.Top = $Script:Top ; $Script:lblCurrLoc.Left = $Script:Left; $Script:lblCurrLoc.Width=150 ;$Script:lblCurrLoc.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblCurrLoc)    # Add to Form 
    $Script:txtCurrLoc = New-Object Windows.Forms.TextBox
        $Script:txtCurrLoc.Top = $Script:Top; $Script:txtCurrLoc.Left = $Script:LeftInput; $Script:txtCurrLoc.Width = 250;
        $Script:txtCurrLoc.ReadOnly = $true
        $Script:txtCurrLoc.Text = $CurrMbxSettings.RetentionPolicy
        $Global:form.Controls.Add($Script:txtCurrLoc)    # Add to Form
                
    $Script:Top = $Script:Top + 30
    ## TASK #
    $Script:lblCreated = New-Object System.Windows.Forms.Label   
        $Script:lblCreated.Text = "Mailbox Created:"
        $Script:lblCreated.Top = $Script:Top ; $Script:lblCreated.Left = $Script:Left; $Script:lblCreated.Width=150 ;$Script:lblCreated.AutoSize = $true
        $Global:form.Controls.Add($Script:lblCreated)    # Add to Form 
    $Script:txtCreated = New-Object Windows.Forms.TextBox
        $Script:txtCreated.Top = $Script:Top; $Script:txtCreated.Left = $Script:LeftInput; $Script:txtCreated.Width = 250;
        $Script:txtCreated.ReadOnly = $true  
        $Script:txtCreated.Text = ($Currmbxsettings.WhenCreated).ToString("MM/dd/yyyy")
        $Global:form.Controls.Add($Script:txtCreated)    # Add to Form
                
    $Script:Top = $Script:Top + 30
    ## Mailbox Size #
    $Script:lblSize = New-Object System.Windows.Forms.Label   
        $Script:lblSize.Text = "Mailbox Size:"
        $Script:lblSize.Top = $Script:Top ; $Script:lblSize.Left = $Script:Left; $Script:lblSize.Width=150 ;$Script:lblSize.AutoSize = $true
        $Global:form.Controls.Add($Script:lblSize)    # Add to Form 
    $Script:txtSize = New-Object Windows.Forms.TextBox
        $Script:txtSize.Top = $Script:Top; $Script:txtSize.Left = $Script:LeftInput; $Script:txtSize.Width = 250;
        $Script:txtSize.ReadOnly = $true  
        $Script:txtSize.Text = $CurrMbxSize
        $Global:form.Controls.Add($Script:txtSize)    # Add to Form

    $Script:Top = $Script:Top + 30
    ## Mailbox Size #
    $Script:lblDeleSize = New-Object System.Windows.Forms.Label   
        $Script:lblDeleSize.Text = "Deleted Item Size:"
        $Script:lblDeleSize.Top = $Script:Top ; $Script:lblDeleSize.Left = $Script:Left; $Script:lblDeleSize.Width=150 ;$Script:lblDeleSize.AutoSize = $true
        $Global:form.Controls.Add($Script:lblDeleSize)    # Add to Form 
    $Script:txtDeleSize = New-Object Windows.Forms.TextBox
        $Script:txtDeleSize.Top = $Script:Top; $Script:txtDeleSize.Left = $Script:LeftInput; $Script:txtDeleSize.Width = 250;
        $Script:txtDeleSize.ReadOnly = $true  
        $Script:txtDeleSize.Text = $CurrMbxDeleSize
        $Global:form.Controls.Add($Script:txtDeleSize)    # Add to Form

    $Script:Top = $Script:Top + 30
    ## Archive Mailbox Size #
    $Script:lblArchSize = New-Object System.Windows.Forms.Label   
        $Script:lblArchSize.Text = "Archive Mailbox Size:"
        $Script:lblArchSize.Top = $Script:Top ; $Script:lblArchSize.Left = $Script:Left; $Script:lblArchSize.Width=150 ;$Script:lblArchSize.AutoSize = $true
        $Global:form.Controls.Add($Script:lblArchSize)    # Add to Form 
    $Script:txtArchSize = New-Object Windows.Forms.TextBox
        $Script:txtArchSize.Top = $Script:Top; $Script:txtArchSize.Left = $Script:LeftInput; $Script:txtArchSize.Width = 250;
        $Script:txtArchSize.ReadOnly = $true  
        $Script:txtArchSize.Text = $ArchMbxSize
        $Global:form.Controls.Add($Script:txtArchSize)    # Add to Form

    $Script:Top = $Script:Top + 30
    ## Archive Mailbox Size #
    $Script:lblArchDeleSize = New-Object System.Windows.Forms.Label   
        $Script:lblArchDeleSize.Text = "Deleted Item Size:"
        $Script:lblArchDeleSize.Top = $Script:Top ; $Script:lblArchDeleSize.Left = $Script:Left; $Script:lblArchDeleSize.Width=150 ;$Script:lblArchDeleSize.AutoSize = $true
        $Global:form.Controls.Add($Script:lblArchDeleSize)    # Add to Form 
    $Script:txtArchDeleSize = New-Object Windows.Forms.TextBox
        $Script:txtArchDeleSize.Top = $Script:Top; $Script:txtArchDeleSize.Left = $Script:LeftInput; $Script:txtArchDeleSize.Width = 250;
        $Script:txtArchDeleSize.ReadOnly = $true  
        $Script:txtArchDeleSize.Text = $ArchMbxDeleSize
        $Global:form.Controls.Add($Script:txtArchDeleSize)    # Add to Form
                
    $Script:Top = $Script:Top + 30
    ## TASK #
    $Script:lblTicket = New-Object System.Windows.Forms.Label   
        $Script:lblTicket.Text = "Ticket Number:"
        $Script:lblTicket.Top = $Script:Top ; $Script:lblTicket.Left = $Script:Left; $Script:lblTicket.Width=150 ;$Script:lblTicket.AutoSize = $true
        $Global:form.Controls.Add($Script:lblTicket)    # Add to Form
    $Script:txtTicket = New-Object Windows.Forms.TextBox
        $Script:txtTicket.Top = $Script:Top; $Script:txtTicket.Left = $Script:LeftInput; $Script:txtTicket.Width = 250;
#        $Script:txtTicket.Text = ($Currmbxsettings.WhenTicket).ToString("MM/dd/yyyy")
        $Global:form.Controls.Add($Script:txtTicket)    # Add to Form
        $Global:InputFocus = $Script:txtTicket

    $LeftCol = 20
    $RightCol = 250
    $PolicyWidth = 220
    $Script:Top = $Script:Top + 30
    ## Set 1 yr Suspended Policy
    $Script:chk1yrSUS = New-Object Windows.Forms.RadioButton
        $Script:chk1yrSUS.Left = $LeftCol; $Script:chk1yrSUS.Width = $PolicyWidth; $Script:chk1yrSUS.Top = $Script:Top
        $Script:chk1yrSUS.Text = "Apply 1yr-Suspended policy" 
        $Script:chk1yrSUS.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chk1yrSUS)

    ## Set 4 yr Delete Active Policy
    $Script:chk4yrAct = New-Object Windows.Forms.RadioButton
        $Script:chk4yrAct.Left = $RightCol; $Script:chk4yrAct.Width = $PolicyWidth; $Script:chk4yrAct.Top = $Script:Top
        $Script:chk4yrAct.Text = "Apply 4yr-Delete policy" 
        $Script:chk4yrAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chk4yrAct)	

    $Script:Top = $Script:Top + 20
    ## Set 3 yr Delete Active Policy
    $Script:chk3yrAct = New-Object Windows.Forms.RadioButton
        $Script:chk3yrAct.Left = $LeftCol; $Script:chk3yrAct.Width = $PolicyWidth; $Script:chk3yrAct.Top = $Script:Top
        $Script:chk3yrAct.Text = "Apply 3yr-Delete policy" 
        $Script:chk3yrAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chk3yrAct)

    ## Set 5 yr Delete Active Policy
    $Script:chk5yrAct = New-Object Windows.Forms.RadioButton
        $Script:chk5yrAct.Left = $RightCol; $Script:chk5yrAct.Width = $PolicyWidth; $Script:chk5yrAct.Top = $Script:Top
        $Script:chk5yrAct.Text = "Apply 5yr-Delete policy" 
        $Script:chk5yrAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chk5yrAct)			

    $Script:Top = $Script:Top + 20
    ## Set 3 yr Archive Active Policy
    $Script:chk3yrArchAct = New-Object Windows.Forms.RadioButton
        $Script:chk3yrArchAct.Left = $LeftCol ; $Script:chk3yrArchAct.Width = $PolicyWidth; $Script:chk3yrArchAct.Top = $Script:Top
        $Script:chk3yrArchAct.Text = "Apply 3yr-Archive-Delete policy" 
        $Script:chk3yrArchAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chk3yrArchAct)

    ## Set 5 yr Delete Active Policy
    $Script:chk5yrArchAct = New-Object Windows.Forms.RadioButton
        $Script:chk5yrArchAct.Left = $RightCol; $Script:chk5yrArchAct.Width = $PolicyWidth; $Script:chk5yrArchAct.Top = $Script:Top
        $Script:chk5yrArchAct.Text = "Apply 5 yr-Archive-Delete policy" 
        $Script:chk5yrArchAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chk5yrArchAct)			

    $Script:Top = $Script:Top + 20
    ## Set 3 yr Delete-1 yr Archive Active Policy
    $Script:chk3yr1yrArchAct = New-Object Windows.Forms.RadioButton
        $Script:chk3yr1yrArchAct.Left = $LeftCol; $Script:chk3yr1yrArchAct.Width = $PolicyWidth; $Script:chk3yr1yrArchAct.Top = $Script:Top
        $Script:chk3yr1yrArchAct.Text = "Apply 3yr Delete-1yr Archive policy" 
        $Script:chk3yr1yrArchAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chk3yr1yrArchAct)

    ## Set 6 yr Delete Active Policy
    $Script:chk6yr2yrArchAct = New-Object Windows.Forms.RadioButton
        $Script:chk6yr2yrArchAct.Left = $RightCol; $Script:chk6yr2yrArchAct.Width = $PolicyWidth; $Script:chk6yr2yrArchAct.Top = $Script:Top
        $Script:chk6yr2yrArchAct.Text = "Apply 6yr Delete-2yr Archive policy" 
        $Script:chk6yr2yrArchAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chk6yr2yrArchAct)

    $Script:Top = $Script:Top + 20
    ## Set 3 yr Delete-2 yr Archive Active Policy
    $Script:chk3yr2yrArchAct = New-Object Windows.Forms.RadioButton
        $Script:chk3yr2yrArchAct.Left = $LeftCol; $Script:chk3yr2yrArchAct.Width = $PolicyWidth; $Script:chk3yr2yrArchAct.Top = $Script:Top
        $Script:chk3yr2yrArchAct.Text = "Apply 3yr Delete-2yr Archive policy" 
        $Script:chk3yr2yrArchAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chk3yr2yrArchAct)

    ## Set 6 yr Delete Active Policy
    $Script:chk6yrArchAct = New-Object Windows.Forms.RadioButton
        $Script:chk6yrArchAct.Left = $RightCol; $Script:chk6yrArchAct.Width = $PolicyWidth; $Script:chk6yrArchAct.Top = $Script:Top
        $Script:chk6yrArchAct.Text = "Apply 6yr Delete-3yr Archive policy" 
        $Script:chk6yrArchAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chk6yrArchAct)
<#
    $Script:Top = $Script:Top + 20
    ## Set 10 yr Delete Active Policy
    $TopCol1 = $TopCol1 + 20
    $Script:chk10yrAct = New-Object Windows.Forms.RadioButton
        $Script:chk10yrAct.Left = 100; $Script:chk10yrAct.Width = $PolicyWidth; $Script:chk10yrAct.Top = $Script:Top
        $Script:chk10yrAct.Text = "Apply 10 yr-Archive-Delete policy" 
        $Script:chk10yrAct.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Script:chk10yrAct)
#>
    Add-FormStandardButtons
}

Enter-EmpNoInputForm
Publish-form

If ($Global:Result -eq "OK")
{
    $ReportPath = "E:\SDAP\WNSUserTransfer\Report\" + $Year + "\"
    If (Test-Path $ReportPath) {} else {New-Item -Path $ReportPath -ItemType Directory}
        
    $Exists = [bool]($CurrMbxSettings = get-mailbox $Global:txtInpEmpNo.Text -ErrorAction SilentlyContinue)
    If ($Exists -eq $True)
    {
        $CurrCASMbxSettings = Get-CASMailbox $Global:txtInpEmpNo.Text
        $CurrMbxSize = (Get-EXOMailboxStatistics $Global:txtInpEmpNo.Text).TotalItemSize
        $CurrMbxDeleSize = (Get-EXOMailboxStatistics $Global:txtInpEmpNo.Text).TotalDeletedItemSize
		$ArchExists = [bool]($ArchiveMailbox = get-mailbox $Global:txtInpEmpNo.Text -Archive -ErrorAction SilentlyContinue)
        If ($ArchExists -eq $True)
        {
            $ArchMbxSize = (Get-EXOMailboxStatistics $Global:txtInpEmpNo.Text -archive).TotalItemSize
            $ArchMbxDeleSize = (Get-EXOMailboxStatistics $Global:txtInpEmpNo.Text -archive).TotalDeletedItemSize
        }

        # Display the current settings for this user
        Select-Policy
        Publish-Form

        If ($Global:Result -eq "OK")
        {
            $ReportFile	= $ReportPath + "\Report-ResetRetention-EmpNo" + $Global:txtInpEmpNo.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"

            $LineToWrite = $RecordEvent + "STAR" + "`tReset Retention Policy"
            WriteReportEvent
	        $LineToWrite = $RecordEvent + "STAR" + "`tLaunched by: " + $WhoAmI + "`n"
            WriteReportEvent        
            $LineToWrite = $RecordEvent + "INFO" + "`tResetting Retention Settings for: " + $Global:txtInpEmpNo.Text + " - (" + $CurrMbxSettings.DisplayName + ")"
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`tTicket Number: " + $Script:txtTicket.Text + "`n"
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t               Current Mailbox Size: " + $Script:txtSize.Text
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t          Current Deleted Item Size: " + $Script:txtDeleSize.Text
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t               Archive Mailbox Size: " + $Script:txtArchSize.Text
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t          Archive Deleted Item Size: " + $Script:txtArchDeleSize.Text
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t              Mailbox Creation Date: " + $CurrCASMbxSettings.whenCreated + "`n"
            WriteReportEvent

            write-host "`nResetting Retention Settings for: " $Global:txtInpEmpNo.Text "- [" $CurrMbxSettings.DisplayName "]" -ForegroundColor Red
            write-host "Current Settings:"
            write-host "   Current Mailbox Size: " $Script:txtSize.Text
            write-host "   Current Deleted Items Size: " $Script:txtDeleSize.Text
            write-host "   Archive Mailbox Size: " $Script:txtArchSize.Text
            write-host "   RoleAssignmentPolicy: " $CurrMbxSettings.RoleAssignmentPolicy
            write-host "   RetentionPolicy: " $CurrMbxSettings.RetentionPolicy
            write-host "   RetentionHoldEnabled: " $CurrMbxSettings.RetentionHoldEnabled
            write-host "   UseDatabaseQuotaDevaults: " $CurrMbxSettings.UseDatabaseQuotaDefaults
            write-host "   IssueWarningQuota: " $CurrMbxSettings.IssueWarningQuota
            write-host "   ProhibitSendQuota: " $CurrMbxSettings.ProhibitSendQuota
            write-host "   ProhibitSendReceiveQuota: " $CurrMbxSettings.ProhibitSendReceiveQuota
            write-host "   ImapEnabled: " $CurrCASMbxSettings.ImapEnabled
            write-host "   PopEnabled: " $CurrCASMbxSettings.PopEnabled
            write-host "   ActiveSyncEnabled: " $CurrCASMbxSettings.ActiveSyncEnabled
			write-host "   Archive Enabled: " $ArchExists
            $LineToWrite = $RecordEvent + "INFO" + "`t     Current Role Assignment Policy: " + $CurrMbxSettings.RoleAssignmentPolicy
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t           Current Retention Policy: " + $CurrMbxSettings.RetentionPolicy
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t     Current Retention Hold Enabled: " + $CurrMbxSettings.RetentionHoldEnabled
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t  Current Retention Hold Start Date: " + $CurrMbxSettings.StartDateForRetentionHold
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`tCurrent Use Database Quota Defaults: " + $CurrMbxSettings.UseDatabaseQuotaDefaults
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t        Current Issue Warning Quota: " + $CurrMbxSettings.IssueWarningQuota
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t        Current Prohibit Send Quota: " + $CurrMbxSettings.ProhibitSendQuota
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`tCurrent Prohibit Send-Receive Quota: " + $CurrMbxSettings.ProhibitSendReceiveQuota
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t                Current ImapEnabled: " + $CurrCASMbxSettings.ImapEnabled
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t                 Current PopEnabled: " + $CurrCASMbxSettings.PopEnabled
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t        Current Active Sync Enabled: " + $CurrCASMbxSettings.ActiveSyncEnabled
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t            Mailbox Archive Enabled: " + $ArchExists + "`n"
            WriteReportEvent

            If ($Apply1yrPolicy -eq "Yes")
            {
                #Apply 1 yr suspended policy
	            write-host  "Re-Setting Mailbox Quota and active 1 yr Suspended retention policy" -ForegroundColor Green
                Set-MailBox $Global:txtInpEmpNo.Text -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 365 Day Delete" –RetentionHoldEnabled $true –StartDateForRetentionHold 04/01/2011
            }
            else
            {
				$CurrLoc = ""
				$EnableArchive = "No"
				If ($Script:chk3yrAct.Checked -eq "Checked")
				{
					write-host  "Re-Setting Mailbox Quota and active 3 yr retention retention policy" -ForegroundColor Green
					$CurrLoc = "UL MRM Policy - 3 yr Delete"
				}
				If ($Script:chk3yrArchAct.Checked -eq "Checked")
				{
					write-host  "Re-Setting Mailbox Quota and active 3 yr Archive-Delete retention policy" -ForegroundColor Green
					$CurrLoc = "UL MRM Policy - 3 yr Delete-Archive"
					$EnableArchive = "Yes"
				}
				If ($Script:chk3yr1yrArchAct.Checked -eq "Checked")
				{
					write-host  "Re-Setting Mailbox Quota and active 3 yr Delete - 1yr Archive retention policy" -ForegroundColor Green
					$CurrLoc = "UL MRM Policy - 3 yr Delete-1 yr Archive"
					$EnableArchive = "Yes"
				}
				If ($Script:chk3yr2yrArchAct.Checked -eq "Checked")
				{
					write-host  "Re-Setting Mailbox Quota and active 3 yr Delete - 2yr Archive retention policy" -ForegroundColor Green
					$CurrLoc = "UL MRM Policy - 3 yr Delete-2 yr Archive"
					$EnableArchive = "Yes"
				}
				If ($Script:chk4yrAct.Checked -eq "Checked")
				{
					write-host  "Re-Setting Mailbox Quota and active 4 yr Delete retention policy" -ForegroundColor Green
					$CurrLoc = "UL MRM Policy - 4 yr Delete"
				}
				If ($Script:chk5yrAct.Checked -eq "Checked")
				{
					write-host  "Re-Setting Mailbox Quota and active 5 yr Delete retention policy" -ForegroundColor Green
					CurrLoc = "UL MRM Policy - 5 yr Delete"
				}				
				If ($Script:chk5yrArchAct.Checked -eq "Checked")
				{
					write-host  "Re-Setting Mailbox Quota and active 5 yr Archive-Delete retention policy" -ForegroundColor Green
					$CurrLoc = "UL MRM Policy - 5 yr Archive"
					$EnableArchive = "Yes"
				}
				If ($Script:chk6yr2yrArchAct.Checked -eq "Checked")
				{
					write-host  "Re-Setting Mailbox Quota and active 6 yr Delete-2yr Archive retention policy" -ForegroundColor Green
					$CurrLoc = "UL MRM Policy - 6 yr Delete-2 yr Archive"
					$EnableArchive = "Yes"
				}
				If ($Script:chk6yrArchAct.Checked -eq "Checked")
				{
					write-host  "Re-Setting Mailbox Quota and active 6 yr Delete-3yr Archive retention policy" -ForegroundColor Green
					$CurrLoc = "UL MRM Policy - 6 yr Delete-3 yr Archive"
					$EnableArchive = "Yes"
				}
<#
				If ($Script:chk10yrAct.Checked -eq "Checked")
				{
					write-host  "Re-Setting Mailbox Quota and active 10 yr Archive-Delete retention policy" -ForegroundColor Green
					$CurrLoc = "UL MRM Policy - 10 yr Archive"
					$EnableArchive = "Yes"
				}
#>
                #Apply Active policy
                Set-MailBox $Global:txtInpEmpNo.Text -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy $CurrLoc  -RetentionHoldEnabled $false -StartDateForRetentionHold $null
				
				If (($EnableArchive -eq "Yes") -and ($ArchExists -eq $False))
				{
						write-host  "Enabling Archive Mailbox" -ForegroundColor Green
						Enable-Mailbox $Global:txtInpEmpNo.Text -Archive
				}
            }
            write-host "Re-Setting mailbox quotas" -ForegroundColor Green
            Set-Mailbox $Global:txtInpEmpNo.Text -UseDatabaseQuotaDefaults $False -IssueWarningQuota 45GB -ProhibitSendQuota 49.75GB -ProhibitSendReceiveQuota 50GB
            write-host "Re-Setting mailbox protocols" -ForegroundColor Green
            Set-CASMailBox $Global:txtInpEmpNo.Text -ImapEnabled $false -PopEnabled $false -ActiveSyncEnabled $false

            $PostMbxSettings = get-mailbox $Global:txtInpEmpNo.Text
            $PostCASMbxSettings = Get-CASMailbox $Global:txtInpEmpNo.Text
			$PostArchiveMailbox = [bool](get-mailbox $Global:txtInpEmpNo.Text -Archive -ErrorAction SilentlyContinue)
            write-host "Current Settings:"
            write-host "   RoleAssignmentPolicy: " $PostMbxSettings.RoleAssignmentPolicy
            write-host "   RetentionPolicy: " $PostMbxSettings.RetentionPolicy
            write-host "   RetentionHoldEnabled: " $PostMbxSettings.RetentionHoldEnabled
            write-host "   UseDatabaseQuotaDevaults: " $PostMbxSettings.UseDatabaseQuotaDefaults
            write-host "   IssueWarningQuota: " $PostMbxSettings.IssueWarningQuota
            write-host "   ProhibitSendQuota: " $PostMbxSettings.ProhibitSendQuota
            write-host "   ProhibitSendReceiveQuota: " $PostMbxSettings.ProhibitSendReceiveQuota
            write-host "   ImapEnabled: " $PostCASMbxSettings.ImapEnabled
            write-host "   PopEnabled: " $PostCASMbxSettings.PopEnabled
            write-host "   ActiveSyncEnabled: " $PostCASMbxSettings.ActiveSyncEnabled
			write-host "   ArchiveEnabled: " $PostArchiveMailbox

            $LineToWrite = $RecordEvent + "INFO" + "`t        Post Role Assignment Policy: " + $PostMbxSettings.RoleAssignmentPolicy
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t              Post Retention Policy: " + $PostMbxSettings.RetentionPolicy
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t        Post Retention Hold Enabled: " + $PostMbxSettings.RetentionHoldEnabled
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t     Post Retention Hold Start Date: " + $PostMbxSettings.StartDateForRetentionHold
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t   Post Use Database Quota Defaults: " + $PostMbxSettings.UseDatabaseQuotaDefaults
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t           Post Issue Warning Quota: " + $PostMbxSettings.IssueWarningQuota
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t           Post Prohibit Send Quota: " + $PostMbxSettings.ProhibitSendQuota
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t   Post Prohibit Send-Receive Quota: " + $PostMbxSettings.ProhibitSendReceiveQuota
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t                   Post ImapEnabled: " + $PostCASMbxSettings.ImapEnabled
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t                    Post PopEnabled: " + $PostCASMbxSettings.PopEnabled
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t           Post Active Sync Enabled: " + $PostCASMbxSettings.ActiveSyncEnabled
            WriteReportEvent
            $LineToWrite = $RecordEvent + "INFO" + "`t               Post Archive Enabled: " + $PostArchiveMailbox + "`n"
            WriteReportEvent

            $LineToWrite = $RecordEvent + "STOP" + "`tThis instance is stopping for: " + $Global:txtInpEmpNo.Text + " - (" + $PostMbxSettings.DisplayName + ")"
            WriteReportEvent
        }
        else
        {
            $Output = $wshell.Popup("Resetting of retention policy cancelled for " + $Global:txtInpEmpNo.Text,5,"Cancelled",0+32)
        }
    }
    else
    {
        $Output = $wshell.Popup("No mailbox found for " + $Global:txtInpEmpNo.Text,5,"No mailbox found",0+32)
    }
}