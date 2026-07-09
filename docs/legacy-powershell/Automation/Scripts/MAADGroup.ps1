#Set forwards on User or Shared Mailboxes
#
#    2026/04/01 - SAG - Created Script to move accounts to another License Group
#
#####################################################

Function Build-MAMgADGroup
{
    $Script:form = New-Object Windows.Forms.Form 
    $Script:form.FormBorderStyle = "FixedToolWindow" 
    $Script:form.Text = "Add To/Remove From an AD Group"
    $Script:form.StartPosition = "CenterScreen" 
    $Script:form.Width = 500 ; $Script:form.Height = 380  # Make the form wider 

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
    ## Select AD Group
    $Script:lblADGroup = New-Object System.Windows.Forms.Label   
        $Script:lblADGroup.Text = "AD Group Name:"  
        $Script:lblADGroup.Top = $Script:Top ; $Script:lblADGroup.Left = 10; $Script:lblADGroup.Width=150; $Script:lblADGroup.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblADGroup)    # Add to Form 
        # 
        $Script:txtADGroup = New-Object Windows.Forms.TextBox
        $Script:txtADGroup.TabIndex = 0 # set Tab Order 
        $Script:txtADGroup.Location = New-Object System.Drawing.Size(140,$Script:Top)
        $Script:txtADGroup.Size = New-Object system.Drawing.Size(300,40)
        $Script:form.Controls.Add($Script:txtADGroup)    # Add to Form 

    $Script:Top = $Script:Top + 30
    ## Enable/Disable OOO
    $Script:chkAddTo = New-Object Windows.Forms.RadioButton
        $Script:chkAddTo.Left = 140; $Script:chkAddTo.Width = 150; $Script:chkAddTo.Top = $Script:Top
        $Script:chkAddTo.Text = "Add to Membership" 
        $Script:chkAddTo.Checked = $false   # set a default value 

        $Script:form.Controls.Add($Script:chkAddTo)
    $Script:chkRemoveFrom = New-Object Windows.Forms.RadioButton
        $Script:chkRemoveFrom.Left = 290; $Script:chkRemoveFrom.Width = 170; $Script:chkRemoveFrom.Top = $Script:Top
        $Script:chkRemoveFrom.Text = "Remove from Membership"
        $Script:chkRemoveFrom.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkRemoveFrom)

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
    $LineToWrite = $RecordEvent + "STAR" + "`t" + "Add/Remove User(s) from AD Group script has started"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + (whoami)
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "User TimeZone: " + (Get-TimeZone)
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Number of Users to Process: " + $addMember.count
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


Build-MAMgADGroup

Add-FormStandardButtons
$Script:OKButton.Text = "Continue"
Publish-Form

Do
{
    If ($Script:Result -eq "OK")
    {
        If ($Script:txtEmpID.Text.Length -gt 0)
        {
            If ($Script:chkAddTo.Checked -eq $True)
            {
                $ReportFile = "E:\Automation\MAActivities\Reports\ModifyADGroups\AddToGroup-" + $Script:txtProject.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
                write-host $ReportFile
                Start-Report
                $addMember = (($Script:txtEmpID.text -replace ("`n",",")).split(",") -replace(" ","")).Trim()
                write-Host "Number of Accounts to Process: " $addMember.Count

                foreach ($u in $addMember)
                {
                    write-host "`nProcessing UserID: " $u -ForegroundColor Cyan
                    $LineToWrite = $RecordEvent + "CURR" + "`t" + "                  Employee ID: " + $u
                    WriteReportEvent
                    $ErrorActionPreference = "SilentlyContinue"
                    $ADExists = $False
                    $ADExists = [bool](get-aduser $u)
                    $ErrorActionPreference = "Continue"

                    If ($ADExists -eq $True)
                    {
                        Add-ADGroupMember $Script:txtADGroup -Member $u
                        $LineToWrite = $RecordEvent + "ADD " + "`t" + "                Added User to: " + $Script:txtADGroup
                    }
                    else
                    {
                        write-host "User Not Found: " $u
                        $LineToWrite = $RecordEvent + "ERR " + "`t" + "  User not found or not added: " + $Script:txtADGroup
                    }
                    WriteReportEvent
                    $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
                    WriteReportEvent
                }
            }

            If ($Script:chkRemoveFrom.Checked -eq $True)
            {
                $ReportFile = "E:\Automation\MAActivities\Reports\ModifyADGroups\RemoveFromGroup-" + $Script:txtProject.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
                write-host $ReportFile
                Start-Report
                $addMember = (($Script:txtEmpID.text -replace ("`n",",")).split(",") -replace(" ","")).Trim()
                write-Host "Number of Accounts to Process: " $addMember.Count

                foreach ($u in $addMember)
                {
                    write-host "`nProcessing UserID: " $u -ForegroundColor Cyan
                    $LineToWrite = $RecordEvent + "CURR" + "`t" + "                  Employee ID: " + $u
                    WriteReportEvent
                    $ErrorActionPreference = "SilentlyContinue"
                    $ADExists = $False
                    $ADExists = [bool](get-aduser $u)
                    $ErrorActionPreference = "Continue"

                    If ($ADExists -eq $True)
                    {
                        Remove-ADGroupMember $Script:txtADGroup -Member $u -confirm:$False
                        $LineToWrite = $RecordEvent + "ADD " + "`t" + "            Removed User from: " + $Script:txtADGroup
                    }
                    else
                    {
                        write-host "User Not Found: " $u
                        $LineToWrite = $RecordEvent + "ERR " + "`t" + "User not found or not removed: " + $Script:txtADGroup
                    }
                    WriteReportEvent
                    $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
                    WriteReportEvent
                }
            }
        }
        else
        {
            $output = $wshell.Popup("No employee number or email addresses entered try again.",0,"No Employee Nos.",0+32)
        }
        $Script:txtEmpID.Text = ""
        Publish-Form
    }
}While ($Script:Result -eq "OK")
