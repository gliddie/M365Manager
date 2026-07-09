#Set forwards on User or Shared Mailboxes
#
#    2024/04/24 - SAG - Created Script to add or remove the migration account access
#
#####################################################

Function Build-MAMigAccount
{
    $Script:form = New-Object Windows.Forms.Form 
    $Script:form.FormBorderStyle = "FixedToolWindow" 
    $Script:form.Text = "Add/Remove Migration Account Access"
    $Script:form.StartPosition = "CenterScreen" 
    $Script:form.Width = 500 ; $Script:form.Height = 350  # Make the form wider 

    $Script:Top = 30
    ## Project Name
    $Script:lblProject = New-Object System.Windows.Forms.Label   
        $Script:lblProject.Text = "Project Name:"  
        $Script:lblProject.Top = $Script:Top ; $Script:lblProject.Left = 10; $Script:lblProject.Width=150; $Script:lblProject.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblProject)    # Add to Form 
        #
        $Script:txtProject = New-Object Windows.Forms.TextBox
        $Script:txtProject.TabIndex = 0 # set Tab Order 
        $Script:txtProject.Top = $Script:Top ; $Script:txtProject.Left = 140; $Script:txtProject.Width = 150; $Script:txtProject.AutoSize = $true
        $Script:txtProject.Location = New-Object System.Drawing.Size(140,$Script:Top)
        $Script:txtProject.Size = New-Object system.Drawing.Size(300,40)
        $Global:InputFocus = $Global:txtProject
        $Script:form.Controls.Add($Script:txtProject)    # Add to Form 

    $Script:Top = $Script:Top + 30
    ## Add Add/Remove Forwarders
    $Script:chkGrant = New-Object Windows.Forms.RadioButton
        $Script:chkGrant.Left = 175; $Script:chkGrant.Width = 100; $Script:chkGrant.Top = $Script:Top
        $Script:chkGrant.Text = "Grant Access" 
        $Script:chkGrant.Checked = $false   # set a default value 
#        $Script:chkGrant.TabIndex = 3
        $Script:form.Controls.Add($Script:chkGrant)
    $Script:chkRemove = New-Object Windows.Forms.RadioButton
        $Script:chkRemove.Left = 300; $Script:chkRemove.Width = 150; $Script:chkRemove.Top = $Script:Top
        $Script:chkRemove.Text = "Remove Access" 
        $Script:chkRemove.Checked = $false   # set a default value 
#        $Script:chkRemove.TabIndex = 3
        $Script:form.Controls.Add($Script:chkRemove) 

    $Script:Top = $Script:Top + 30
    ## Enter Employee #s
    $Script:lblEmpID = New-Object System.Windows.Forms.Label   
        $Script:lblEmpID.Text = "Employee Number(s):"  
        $Script:lblEmpID.Top = $Top ; $Script:lblEmpID.Left = 10; $Script:lblEmpID.Width=150; $Script:lblEmpID.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblEmpID)    # Add to Form 
        # 
        $Script:txtEmpID = New-Object Windows.Forms.TextBox
        $Script:txtEmpID.MaxLength = 2000000
        $Script:txtEmpID.TabIndex = 0 # set Tab Order 
        $Script:txtEmpID.Location = New-Object System.Drawing.Size(140,$Script:Top)
        $Script:txtEmpID.Size = New-Object system.Drawing.Size(300,150)
        $Script:txtEmpID.MultiLine = $true
        $Script:txtEmpID.ScrollBars = 'Both'  
        $Script:form.Controls.Add($Script:txtEmpID)    # Add to Form 
}

Function Start-Report
{
    write-host "Logging activites in Report File: " $ReportFile -ForegroundColor Yellow
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + (whoami)
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "User TimeZone: " + (Get-TimeZone)
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Project Name: " + $Script:txtProject.Text
    WriteReportEvent

    If ($addMember.count -lt 1)
    {
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Number of Users to Process: 1"
    }
    else
    {
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Number of Users to Process: " + $addMember.count
    }
    WriteReportEvent
    $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
    WriteReportEvent
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

$wshell = New-Object -ComObject Wscript.Shell
Build-MAMigAccount
Add-FormStandardButtons
$Script:OKButton.Text = "Continue"
Publish-Form

Do
{
    If ($Script:Result -eq "OK")
    {
        If ($Script:chkGrant.Checked -eq  $True)
        {
            $ReportFile = "E:\Automation\MAActivities\Reports\MigrationAccountAccess\AddAccess-" + ($Script:txtProject.Text).Trim() + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            write-host "Granting SVC.CRP.Migration FullAccess" -ForegroundColor Yellow
            $LineToWrite = $RecordEvent + "STAR" + "`t" + "Granting SVC.CRP.Migration Accounts Access script has started"

            Start-Report
            $addMember = (($Script:txtEmpID.text -replace ("`n",",")).split(",") -replace(" ","")).Trim()
            write-Host "Number of Accounts to Process: " $addMember.Count
            foreach ($u in $addMember)
            {
                write-host "Processing UserID: " $u
                $LineToWrite = $RecordEvent + "CURR" + "`t" + "Employee ID: " + $u
                WriteReportEvent
                $MbxExists = [bool](get-mailbox $u -ErrorAction SilentlyContinue)
                If ($MbxExists -eq "True")
                {
                    If ([bool](Get-MailboxPermission $u|Where-Object {$_.User -eq "svc.crp.migration@global.ul.com"}))
                    {
                        Write-Host "Migration Account already has access to: " $u
                        $LineToWrite = $RecordEvent + "CURR" + "`t" + "The Migration Account already has access to: " + $u
                    }
                    else
                    {
                        Write-Host "Granting Migration Account access to: " $u
                        Add-MailboxPermission $u -AccessRights Fullaccess -User svc.crp.migration@global.ul.com -AutoMapping:$False
                        $LineToWrite = $RecordEvent + "CURR" + "`t" + "Adding Migration Account access to: " + $u
                    }
                    write-host "LineToWrite:" $LineToWrite
                    WriteReportEvent
                }
                else
                {
                    $LineToWrite = $RecordEvent + "CURR" + "`t" + "No mailbox found for: " + $u
                }
                $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
                WriteReportEvent
            }
        }

        If ($Script:chkRemove.Checked -eq  $True)
        {
            $ReportFile = "E:\Automation\MAActivities\Reports\MigrationAccountAccess\RemoveAccess-" + ($Script:txtProject.Text).Trim() + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            write-host "Removing SVC.CRP.Migration FullAccess"  -ForegroundColor Yellow
            $LineToWrite = $RecordEvent + "STAR" + "`t" + "Removing SVC.CRP.Migration Account Access script has started"
            Start-Report
            $addMember = (($Script:txtEmpID.text -replace ("`n",",")).split(",") -replace(" ","")).Trim()
            write-Host "Number of Accounts to Process: " $addMember.Count
            foreach ($u in $addMember)
            {
                write-host "Processing UserID: " $u
                $LineToWrite = $RecordEvent + "CURR" + "`t" + "Employee ID: " + $u
                WriteReportEvent
                $MbxExists = [bool](get-mailbox $u -ErrorAction SilentlyContinue)
                If ($MbxExists -eq "True")
                {
                    If ([bool](Get-MailboxPermission $u|Where-Object {$_.User -eq "svc.crp.migration@global.ul.com"}))
                    {
                        Write-Host "Removing Migration Account access from: " $u
                        Remove-MailboxPermission $u -AccessRights Fullaccess -User svc.crp.migration@global.ul.com -Confirm:$False
                        $LineToWrite = $RecordEvent + "CURR" + "`t" + "Removing Migration Account access from: " + $u
                    }
                    else
                    {
                        Write-Host "Migration Account does not have access to: " $u
                        $LineToWrite = $RecordEvent + "CURR" + "`t" + "The Migration Account does not have access to: " + $u
                    }
                    WriteReportEvent
                }
                else
                {
                    $LineToWrite = $RecordEvent + "CURR" + "`t" + "No mailbox found for: " + $u
                }
                $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
                WriteReportEvent
            }
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Migration Account Access Proecessing Complete"
            WriteReportEvent
        }
        $Script:chkGrant.Checked = $False
        $Script:chkRemove.Checked = $False
        $Script:txtProject.Text = ""
        $Script:txtEmpID.Text = ""
        Publish-Form
    }
}While ($Script:Result -eq "OK")
