<####
#### MA Activities Menu
#
#  Called by:  O365AdminMenu.ps1
#
#  04/15/2020 - SAG - Combined the AddMailboxPermissionSharedMailbox and ApplyRetentionPolicy scripts into the AddFolderPermission script for the Option 1 and 3
#  05/20/2022 - SAG - Added Shared Mailbox Removal with Input File
#  11/04/2022 - SAG - Added New Script to add access groups for an existing mailbox
#  08/07/2024 - SAG - Added Menu Item for setting OneDrive or Sharepoints to ReadOnly or Unlocked
#
####>

function Add-FormStandardButtons
{
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    $Script:buttonPanel = New-Object Windows.Forms.Panel  
    $buttonPanel.Size = New-Object Drawing.Size @(400,40) 
    $buttonPanel.Dock = "Bottom"    
    $Script:cancelButton = New-Object Windows.Forms.Button  
        $Script:cancelButton.Top = $buttonPanel.Height - $Script:cancelButton.Height - 10; $Script:cancelButton.Left = $buttonPanel.Width - $Script:cancelButton.Width - 10 
        $Script:cancelButton.TabIndex = 198
        $Script:cancelButton.Text = "Cancel" 
        $Script:cancelButton.DialogResult = "Cancel" 
        $Script:cancelButton.Anchor = "Right"
    ## Create the OK button, which will anchor to the left of Cancel 
    $Script:okButton = New-Object Windows.Forms.Button   
        $Script:okButton.Top = $cancelButton.Top ; $Script:okButton.Left = $cancelButton.Left - $Script:okButton.Width - 10
        $Script:okButton.TabIndex = 97
        $Script:okButton.Text = $Action
        If ($Script:OKDetails -ne "")
        {
            $Script:okButton.Text = $Script:OKDetails
            $Script:OKDetails = ""
        }
        else
        {
            $Script:okButton.Text = "Continue"
        }
        $Script:okButton.DialogResult = "OK" 
        $Script:okButton.Anchor = "Right"
    ## Add the buttons to the button panel 
    $Script:buttonPanel.Controls.Add($Script:okButton) 
    $Script:buttonPanel.Controls.Add($Script:cancelButton) 
    ## Add the button panel to the form 
    $Script:form.Controls.Add($buttonPanel)
    ## Set Default actions for the buttons 
    $Script:form.AcceptButton = $Script:okButton          # ENTER = ok 
    $Script:form.CancelButton = $Script:cancelButton      # ESCAPE = Cancel
}

Function Publish-Form
{
    If ($Script:InputFocus -eq $null)
    {
        $Script:InputFocus = $Script:okButton
    }
    ## Finalize Form and Show Dialog
    $Script:form.Add_Shown( { $Script:form.Activate(); $Script:InputFocus.Focus() } )  #Activate and Set Focus 
#    $Global:form.Add_Shown( { $form.Activate(); $Global:okButton.Focus() } )  #Activate and Set Focus 
    $Script:Result = $Script:form.ShowDialog()          ## Show the form, and wait for the response
}

#  Writes events to the Report File
function WriteReportEvent
{
    If (($null -eq $ReportFile) -or ($ReportFile -eq ""))
    {
        $ReportFile = $Script:ReportFile
    }
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LineToWrite
	Out-File -filepath $ReportFile -append -noClobber -inputObject $RecordEvent
} #end WriteReportEvent

function Build-MAMenuForm
{
    $Script:form = New-Object Windows.Forms.Form 
    $Script:form.FormBorderStyle = "FixedToolWindow" 
    $Script:form.Text = "M&A Activites Admin Menu" 
    $Script:form.StartPosition = "CenterScreen" 
    $Script:form.Width = 550 ; $form.Height = 400  # Make the form wider

    $TopLoc = 20 
    ## Label and TextBox  
    ## Title Line
    $Script:lblDLSecTitleLine = New-Object System.Windows.Forms.Label   
        $Script:lblDLSecTitleLine.Text = "Select Option:"
        $Script:lblDLSecTitleLine.Top = $TopLoc ; $Script:lblDLSecTitleLine.Left = 60; $Script:lblDLSecTitleLine.Width=220 ;$Script:lblDLSecTitleLine.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblDLSecTitleLine)    # Add to Form 

    ## Add to an AD Group
    $TopLoc = $TopLoc + 20
    $Script:chkADGroup = New-Object Windows.Forms.RadioButton
        $Script:chkADGroup.Left = 100; $Script:chkADGroup.Width = 450; $Script:chkADGroup.Top = $TopLoc  
        $Script:chkADGroup.Text = "Add To/Remove From an ADGroup" 
        $Script:chkADGroup.Checked = $false   # set a default value 
        $Script:chkADGroup.TabIndex = 2
        $Script:form.Controls.Add($Script:chkADGroup) 

    ## Add Add/Remove Forwarders
    $TopLoc = $TopLoc + 20
    $Script:chkForwarders = New-Object Windows.Forms.RadioButton
        $Script:chkForwarders.Left = 100; $Script:chkForwarders.Width = 450; $Script:chkForwarders.Top = $TopLoc  
        $Script:chkForwarders.Text = "Add/Remove Forwarders" 
        $Script:chkForwarders.Checked = $false   # set a default value 
        $Script:chkForwarders.TabIndex = 2
        $Script:form.Controls.Add($Script:chkForwarders) 

    ## Grant Migration Account Access
    $TopLoc = $TopLoc + 20
    $Script:chkAddMigAcct = New-Object Windows.Forms.RadioButton
        $Script:chkAddMigAcct.Left = 100; $Script:chkAddMigAcct.Width = 450; $Script:chkAddMigAcct.Top = $TopLoc  
        $Script:chkAddMigAcct.Text = "Add/Remove SVC.CRP.Migration Account Access" 
        $Script:chkAddMigAcct.Checked = $false   # set a default value 
        $Script:chkAddMigAcct.TabIndex = 3
        $Script:form.Controls.Add($Script:chkAddMigAcct)

    ## Cleanup Owner/Member in Groups
    $TopLoc = $TopLoc + 20
    $Script:chkCUOwnMem = New-Object Windows.Forms.RadioButton
        $Script:chkCUOwnMem.Left = 100; $Script:chkCUOwnMem.Width = 450; $Script:chkCUOwnMem.Top = $TopLoc  
        $Script:chkCUOwnMem.Text = "CleanUp Owner/Membership of Groups" 
        $Script:chkCUOwnMem.Checked = $false   # set a default value 
        $Script:chkCUOwnMem.TabIndex = 3
        $Script:form.Controls.Add($Script:chkCUOwnMem)

    ## Disable Mailbox Protocols
    $TopLoc = $TopLoc + 20
    $Script:chkMbxProt = New-Object Windows.Forms.RadioButton
        $Script:chkMbxProt.Left = 100; $Script:chkMbxProt.Width = 450; $Script:chkMbxProt.Top = $TopLoc  
        $Script:chkMbxProt.Text = "Enable/Disable Mailbox Protocols" 
        $Script:chkMbxProt.Checked = $false   # set a default value 
        $Script:chkMbxProt.TabIndex = 4
        $Script:form.Controls.Add($Script:chkMbxProt)   

    ## Enable OOO
    $TopLoc = $TopLoc + 20
    $Script:chkEnablOOO = New-Object Windows.Forms.RadioButton
        $Script:chkEnablOOO.Left = 100; $Script:chkEnablOOO.Width = 450; $Script:chkEnablOOO.Top = $TopLoc  
        $Script:chkEnablOOO.Text = "Enable/Disable Out-Of-Office Message for External Only" 
        $Script:chkEnablOOO.Checked = $false   # set a default value 
        $Script:chkEnablOOO.TabIndex = 4
        $Script:form.Controls.Add($Script:chkEnablOOO)      

    ## Hide/Unhide from Address Book
    $TopLoc = $TopLoc + 20
    $Script:chkHideUnhide = New-Object Windows.Forms.RadioButton
        $Script:chkHideUnhide.Left = 100; $Script:chkHideUnhide.Width = 450; $Script:chkHideUnhide.Top = $TopLoc  
        $Script:chkHideUnhide.Text = "Hide/Unhide from Address Book" 
        $Script:chkHideUnhide.Checked = $false   # set a default value 
        $Script:chkHideUnhide.TabIndex = 3
        $Script:form.Controls.Add($Script:chkHideUnhide)

    ## Lock/Unlock OneDrive/SharePoint sites
    $TopLoc = $TopLoc + 20
    $Script:chkLockUnlock = New-Object Windows.Forms.RadioButton
        $Script:chkLockUnlock.Left = 100; $Script:chkLockUnlock.Width = 450; $Script:chkLockUnlock.Top = $TopLoc  
        $Script:chkLockUnlock.Text = "Lock/Unlock OneDrive/Sharepoint site" 
        $Script:chkLockUnlock.Checked = $false   # set a default value 
        $Script:chkLockUnlock.TabIndex = 3
        $Script:form.Controls.Add($Script:chkLockUnlock)

    ## Change external access to mailbox
    $TopLoc = $TopLoc + 20
    $Script:chkMoveLicense = New-Object Windows.Forms.RadioButton
        $Script:chkMoveLicense.Left = 100; $Script:chkMoveLicense.Width = 450; $Script:chkMoveLicense.Top = $TopLoc  
        $Script:chkMoveLicense.Text = "Move to New Licensing Group" 
        $Script:chkMoveLicense.Checked = $false   # set a default value 
        $Script:chkMoveLicense.TabIndex = 5
        $Script:form.Controls.Add($Script:chkMoveLicense) 

    ## Change external access to mailbox
    $TopLoc = $TopLoc + 20
    $Script:chkMoveOU = New-Object Windows.Forms.RadioButton
        $Script:chkMoveOU.Left = 100; $Script:chkMoveOU.Width = 450; $Script:chkMoveOU.Top = $TopLoc  
        $Script:chkMoveOU.Text = "Move User Account to IAM Restricted OU" 
        $Script:chkMoveOU.Checked = $false   # set a default value 
        $Script:chkMoveOU.TabIndex = 5
        $Script:form.Controls.Add($Script:chkMoveOU) 

#Legal Hold
#Partial Purge Process
#EmpID,ULAddr,Forwarder
    ## Report of Users and Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkGrpMemOwn = New-Object Windows.Forms.RadioButton
        $Script:chkGrpMemOwn.Left = 100; $Script:chkGrpMemOwn.Width = 450; $Script:chkGrpMemOwn.Top = $TopLoc
        $Script:chkGrpMemOwn.Text = "Report Group Ownership/Membership" 
        $Script:chkGrpMemOwn.Checked = $false   # set a default value 
        $Script:chkGrpMemOwn.TabIndex = 6
        $Script:form.Controls.Add($Script:chkGrpMemOwn) 

    ## Report of Users and Licenses
    $TopLoc = $TopLoc + 20
    $Script:chkUsrLic = New-Object Windows.Forms.RadioButton
        $Script:chkUsrLic.Left = 100; $Script:chkUsrLic.Width = 450; $Script:chkUsrLic.Top = $TopLoc
        $Script:chkUsrLic.Text = "Report of M&A Users and O365 Licenses" 
        $Script:chkUsrLic.Checked = $false   # set a default value 
        $Script:chkUsrLic.TabIndex = 6
        $Script:form.Controls.Add($Script:chkUsrLic) 

    ## Get Mailbox Statistics
    $TopLoc = $TopLoc + 20
    $Script:chkStats = New-Object Windows.Forms.RadioButton
        $Script:chkStats.Left = 100; $Script:chkStats.Width = 450; $Script:chkStats.Top = $TopLoc  
        $Script:chkStats.Text = "Get Mailbox Statistics for M&A Users" 
        $Script:chkStats.Checked = $false   # set a default value 
        $Script:chkStats.TabIndex = 7
        $Script:form.Controls.Add($Script:chkStats) 
}

    $me = whoami
    $CredENo = $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))
    $UsrName = $CredENo + "@global.ul.com"
    $dir = "c:\users\" + $CredENo + "\documents\"
    $File = "my" + $CredENo + "File.xml"
    $Global:CredFile = $dir + $File
    #AdminCredFile
    $AFile = "myA" + ($me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))).replace(".","") + "File.xml"
    $Global:ACredFile = $dir + $AFile
    If (Test-Path $Global:ACredFile)
    {
        $Global:AdmLiveCred = Import-Clixml $Global:ACredFile
    }
    Build-MAMenuForm
    Add-FormStandardButtons
    $Script:OKButton.Text = "Continue"
    $Script:InputFocus = $Script:okButton
    Publish-Form

Do
{
    If ($Script:chkUsrLic.Checked -eq "Checked")
    {
        write-host "Hit Return when the Input file in the e:\Automation\MAActivites\Input directory ready" -ForegroundColor Yellow -NoNewline
        $Cnt = Read-Host
        invoke-expression -Command .\MALicenseInformation.ps1
        pause
    }

    If ($Script:chkGrpMemOwn.Checked -eq "Checked")
    {
        invoke-expression -Command .\MAGroupMembershipReport.ps1
    }

    If ($Script:chkAddMigAcct.Checked -eq "Checked")
    {
        invoke-expression -Command .\MAMigAccount.ps1
        $Script:chkAddMigAcct.Checked = $False
    }

    If ($Script:chkCUOwnMem.Checked -eq "Checked")
    {
        invoke-expression -Command .\MARemoveOwnerMember.ps1
        $Script:chkcuOwnMem.Checked = $False
    }

    If ($Script:chkHideUnhide.Checked -eq "Checked")
    {
        invoke-expression -Command .\MAHideUnhideAccount.ps1
        $Script:chkHideUnhide.Checked = $False
    }

    If ($Script:chkLockUnlock.Checked -eq "Checked")
    {
        write-host "Developing this process"
#        invoke-expression -Command .\MAHideUnhideAccount.ps1
        $Script:chkLockUnlock.Checked = $False
    }

    If ($Script:chkADGroup.Checked -eq "Checked")                 
    {
        invoke-expression -Command .\MAADGroup.ps1
        $Script:chkADGroup.Checked = $False
    }
   
    If ($Script:chkForwarders.Checked -eq "Checked")                 
    {
        invoke-expression -Command .\MAForwarders.ps1
        $Script:chkForwarders.Checked = $False
    }

    If ($Script:chkMbxProt.Checked -eq "Checked")
    {
        invoke-expression -Command .\MAProtocols.ps1
        $Script:chkMbxProt.Checked = $False 
    }

    If ($Script:chkEnablOOO.Checked -eq "Checked")
    {
        invoke-expression -Command .\MAEnableDisableOOO.ps1
        $Script:chkEnablOOO.Checked = $False        
    }

    If ($Script:chkStats.Checked -eq "Checked")
    {
        write-host "Hit Return when the Input file in the e:\Automation\MAActivites\Input directory ready" -ForegroundColor Yellow -NoNewline
        $Cnt = Read-Host
                    
        write-host "Executing Report for M&A Mailbox Statistics....." -ForegroundColor Yellow

        $MAUsers = Import-Csv e:\Automation\MAActivities\Input\MAActivities.csv

        $text = "UserPrincipalName,DisplayName,ItemCount"
        Out-File -FilePath c:\temp\MAMailboxItemCount.csv -InputObject $text

        Foreach ($MAUsers in $MAUsers)
        {
            $Usr = $MAUsers.UPN + "@global.ul.com"
            $UsrExists = [bool](get-mailbox $Usr -ErrorAction SilentlyContinue)

            If ($usrExists -eq "True")
            {
                $stsmbx = get-mailboxstatistics $Usr
                $text = "{0},""{1}"",{2}" -f $Usr, $stsmbx.DisplayName, $stsmbx.ItemCount
            }
            else
            {
                Write-Host "User " $MAUsers.UPN " does not exist"
                $text = "{0},{1},{2}" -f $Usr, "False", "User Does Not Exist"
            }
            Out-File -FilePath c:\temp\MAMailboxItemCount.csv -InputObject $text -Append
        }
        Write-Host "Report file generated at c:\temp\MALicenseInformation.csv.  If you would like to maintain the input CSV file please rename it manually." -ForegroundColor Red
        $Script:chkStats.Checked = $False
    }

    If ($Script:chkMoveLicense.Checked -eq "Checked")
    {
        invoke-expression -Command .\MAMoveLicense.ps1
        $Script:chkMoveLicense.Checked = $False
    }

    If ($Script:chkMoveOU.Checked -eq "Checked")
    {
        invoke-expression -Command .\MAMoveADOU.ps1
        $Script:chkMoveOU.Checked = $False
    }

    If ($Script:Result -eq "OK")
    {
        $Script:InputFocus = $Script:okButton
        Publish-Form
    }

}While ($Script:Result -eq "OK")	