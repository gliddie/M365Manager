#Set forwards on User or Shared Mailboxes
#
#    2024/04/24 - SAG - Created Script to add or remove the migration account access
#
#####################################################

Function Build-HideUnhideAccount
{
    $Script:form = New-Object Windows.Forms.Form 
    $Script:form.FormBorderStyle = "FixedToolWindow" 
    $Script:form.Text = "Hide/Unhide Account from Address Book"
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
    $Script:chkHide = New-Object Windows.Forms.RadioButton
        $Script:chkHide.Left = 140; $Script:chkHide.Width = 150; $Script:chkHide.Top = $Script:Top
        $Script:chkHide.Text = "Hide From Address Book" 
        $Script:chkHide.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkHide)
        $Script:chkHide.add_Click({
            $Script:OKButton.Visible = $false
        }) 
    $Script:chkUnhide = New-Object Windows.Forms.RadioButton
        $Script:chkUnhide.Left = 290; $Script:chkUnhide.Width = 160; $Script:chkUnhide.Top = $Script:Top
        $Script:chkUnhide.Text = "Unhide from Address Book" 
        $Script:chkUnhide.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkUnhide)
        $Script:chkUnhide.add_Click({
            $Script:OKButton.Visible = $false
        })  

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
        $Script:txtEmpID.add_Click({
            $Script:OKButton.Visible = $True
        }) 
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
Build-HideUnhideAccount
Add-FormStandardButtons
$Script:OKButton.Text = "Continue"
$Script:OKButton.Visible = $False
Publish-Form

Do
{
    If ($Script:Result -eq "OK")
    {
        $addMember = (($Script:txtEmpID.text -replace ("`n",",")).split(",") -replace(" ","")).Trim()
        write-Host "Number of Accounts to Process: " $addMember.Count
        If ($Script:chkHide.Checked -eq  $True)
        {
            $ReportFile = "E:\Automation\MAActivities\Reports\HideUnhide\Hide-" + ($Script:txtProject.Text).Trim() + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            write-host "Hiding Mailboxes from the Address Book" -ForegroundColor Yellow
            $LineToWrite = $RecordEvent + "STAR" + "`t" + "Hide Unhide Mailboxes from the Address Book script has started"
            Start-Report
            Foreach ($u in $addmember)
            {
                If ($u.length -gt 0)
                {
                    write-host "Processing UserID: " $u
                    $LineToWrite = $RecordEvent + "CURR" + "`t" + "Employee ID: " + $u
                    WriteReportEvent
                    $MbxExists = [bool]($mbx = get-mailbox $u -ErrorAction SilentlyContinue)
                    If ($MbxExists -eq "True")
                    {
                        If ($mbx.HiddenFromAddressListsEnabled -eq $False)
                        {
                            Write-Host "Hide mailbox from the address book: " $u
                            Set-ADUser $u -Replace @{'msExchHideFromAddressLists'=$True}
                            $LineToWrite = $RecordEvent + "CURR" + "`t" + "Hiding mailbox from the address book: " + $u
                        }
                        else
                        {
                            Write-Host "Mailbox is already hidden from the address book: " $u
                            $LineToWrite = $RecordEvent + "CURR" + "`t" + "Mailbox is already hidden from the address book: " + $u
                        }
                        WriteReportEvent
                    }
                    else
                    {
                        $LineToWrite = $RecordEvent + "CURR" + "`t" + "No mailbox found for: " + $u
                        WriteReportEvent
                    }
                    $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
                    WriteReportEvent
                }
            }
            $LineToWrite = $RecordEvent + "DONE" + "`t" + "Hiding Mailboxes Complete"
            WriteReportEvent
        }

        If ($Script:chkUnhide.Checked -eq  $True)
        {
            $ReportFile = "E:\Automation\MAActivities\Reports\MigrationAccountAccess\Unhide-" + ($Script:txtProject.Text).Trim() + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            write-host "Unhiding Mailboxes from the Address Book"  -ForegroundColor Yellow
            $LineToWrite = $RecordEvent + "STAR" + "`t" + "Unhiding Mailboxes from the Address Book script has started"
            Start-Report
            Foreach ($u in $addmember)
            {
                If ($u.length -gt 0)
                {
                    write-host "Processing UserID: " $u
                    $LineToWrite = $RecordEvent + "CURR" + "`t" + "Employee ID: " + $u
                    WriteReportEvent
                    $MbxExists = [bool]($mbx = get-mailbox $u -ErrorAction SilentlyContinue)
                    If ($MbxExists -eq "True")
                    {
                        If ($mbx.HiddenFromAddressListsEnabled -eq $True)
                        {
                            Write-Host "Unhiding mailbox from the address book: " $u
                            Set-ADUser $u -Replace @{'msExchHideFromAddressLists'=$False}
                            $LineToWrite = $RecordEvent + "CURR" + "`t" + "Unhiding mailbox from the address book: " + $u
                        }
                        else
                        {
                            Write-Host "Mailbox is already unhidden from the address book: : " $u
                            $LineToWrite = $RecordEvent + "CURR" + "`t" + "Mailbox is already unhidden from the address book: " + $u
                        }
                        WriteReportEvent
                    }
                    else
                    {
                        $LineToWrite = $RecordEvent + "CURR" + "`t" + "No mailbox found for: " + $u
                        WriteReportEvent
                    }
                    $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
                    WriteReportEvent
                }
            }
            $LineToWrite = $RecordEvent + "DONE" + "`t" + "UnHiding Mailboxes Complete"
            WriteReportEvent
       }

        $Script:OKButton.Visible = $False
        $Script:chkHide.Checked = $False
        $Script:chkUnhide.Checked = $False
        $Script:txtProject.Text = ""
        $Script:txtEmpID.Text = ""
        Publish-Form
    }
}While ($Script:Result -eq "OK")
