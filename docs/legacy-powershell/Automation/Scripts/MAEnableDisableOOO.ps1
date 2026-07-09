#Set OOO on Mailboxes
#
#    2024/04/18 - SAG - Created Script to set OOO for FIME divestiture
#    2026/03/25 - SAG - Updated for Project Elvis (pureEHS) divestiture
#    2026/03/31 - SAG - Updated OOO for Project Elvis as the messages was adjusged after the original 3/25 change.
#
#####################################################

Function Build-MAOutOfOffice
{
    $Script:form = New-Object Windows.Forms.Form 
    $Script:form.FormBorderStyle = "FixedToolWindow" 
    $Script:form.Text = "Enable/Disable Out Of Office Message"
    $Script:form.StartPosition = "CenterScreen" 
    $Script:form.Width = 500 ; $Script:form.Height = 440  # Make the form wider 

    $Script:Top = 20
    ## Current OOO Message
    $Script:lblSampleMsg = New-Object System.Windows.Forms.Label   
        $Script:lblSampleMsg.Text = "New OOO Message:"  
        $Script:lblSampleMsg.Top = $Script:Top ; $Script:lblSampleMsg.Left = 10; $Script:lblSampleMsg.Width=150; $Script:lblSampleMsg.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblSampleMsg)    # Add to Form 
        #
        $Script:txtSampleMsg = New-Object Windows.Forms.TextBox
        $Script:txtSampleMsg.TabIndex = 0 # set Tab Order 
        $Script:txtSampleMsg.Top = $Script:Top ; $Script:txtSampleMsg.Left = 140; $Script:txtSampleMsg.Width = 150; $Script:txtSampleMsg.AutoSize = $true
        $Script:txtSampleMsg.Location = New-Object System.Drawing.Size(140,$Script:Top)
        $Script:txtSampleMsg.Size = New-Object system.Drawing.Size(300,90)
        $Script:txtSampleMsg.MultiLine = $true
#        $SampleMsg = "Please note that the UL Solutions email system automatically forwarded your email to ConfiguredForwarder@domain.com.`nFor more information about the sale of the UL Solutions payments functional testing business please read the press release here: <a href='http://www.ul.com'>UL Solutions</a>.</>"
        $SampleMsg = "Please note that the UL Solutions email system automatically forwarded your email to NAME (AD Name) at ConfiguredForwarder@domain.com.`nFor inquiries or more information, please reach out directly to NAME (AD Name).</>"
        $Script:txtSampleMsg.Text = $SampleMsg
        $Script:txtSampleMsg.ReadOnly = $True
        $Script:form.Controls.Add($Script:txtSampleMsg)    # Add to Form

    $Script:Top = $Script:Top + 90
    ## Enable/Disable OOO
    $Script:chkEnable = New-Object Windows.Forms.RadioButton
        $Script:chkEnable.Left = 140; $Script:chkEnable.Width = 150; $Script:chkEnable.Top = $Script:Top
        $Script:chkEnable.Text = "Enable Out Of Office" 
        $Script:chkEnable.Checked = $false   # set a default value 
        $Script:chkEnable.add_Click({
            $Script:lblNoDays.Visible = $True
             $Script:txtNoDays.visible = $True
             $Script:OKButton.Visible = $False
        })
        $Script:form.Controls.Add($Script:chkEnable)
    $Script:chkDisable = New-Object Windows.Forms.RadioButton
        $Script:chkDisable.Left = 290; $Script:chkDisable.Width = 150; $Script:chkDisable.Top = $Script:Top
        $Script:chkDisable.Text = "Disable Out Of Office" 
        $Script:chkDisable.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkDisable)
        $Script:chkDisable.add_Click({
            $Script:lblNoDays.Visible = $False
            $Script:txtNoDays.visible = $False
            $Script:OKButton.Visible = $False
        })

    $Script:Top = $Script:Top + 25
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
        $Script:InputFocus = $Script:txtProject
        $Script:form.Controls.Add($Script:txtProject)    # Add to Form 

    $Script:Top = $Script:Top + 30
    ## Number of Days
    $Script:lblNoDays = New-Object System.Windows.Forms.Label   
        $Script:lblNoDays.Text = "No Days Active:"  
        $Script:lblNoDays.Top = $Script:Top ; $Script:lblNoDays.Left = 10; $Script:lblNoDays.Width=150; $Script:lblNoDays.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblNoDays)    # Add to Form 
        #
        $Script:txtNoDays = New-Object Windows.Forms.TextBox
        $Script:txtNoDays.TabIndex = 0 # set Tab Order 
        $Script:txtNoDays.Top = $Script:Top ; $Script:txtNoDays.Left = 140; $Script:txtNoDays.Width = 150; $Script:txtNoDays.AutoSize = $true
        $Script:txtNoDays.Location = New-Object System.Drawing.Size(140,$Script:Top)
        $Script:txtNoDays.Size = New-Object system.Drawing.Size(50,40)
        $Script:form.Controls.Add($Script:txtNoDays)    # Add to Form 

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
    $LineToWrite = $RecordEvent + "STAR" + "`t" + "Enable/Disable Out Of Office script has started"
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

$wshell = New-Object -ComObject Wscript.Shell
Build-MAOutOfOffice
Add-FormStandardButtons
$Script:OKButton.Text = "Continue"
$Script:OKButton.Visible = $False
Publish-Form

Do
{
    If ($Script:Result -eq "OK")
    {
        If ($Script:txtEmpID.Text.Length -gt 0)
        {
            $addMember = (($Script:txtEmpID.text -replace ("`n",",")).split(",") -replace(" ","")).Trim()
            write-Host "Number of Accounts to Process: " $addMember.Count
            If ($Script:chkEnable.Checked -eq $True)
            {
                $ReportFile = "E:\Automation\MAActivities\Reports\OutOfOffice\OutOfOfficeEnable-" + $Script:txtProject.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
                Start-Report
                foreach ($u in $addMember)
                {
                    $Addr = ""
                    $Name = ""
                    If ($u.length -gt 0)
                    {
                        write-host "`nProcessing Mailbox: " $u -ForegroundColor Cyan
                        $LineToWrite = $RecordEvent + "CURR" + "`t" + "                  Employee ID: " + $u
                        WriteReportEvent
                        $MbxExists = [bool]($mbx = get-mailbox $u)
                        If ($MbxExists -eq $True)
                        {
                            $Name = (get-aduser $u).Name
                            If ($mbx.ForwardingSMTPAddress.length -gt 0)
                            {
                                $Addr = $mbx.ForwardingSMTPAddress.substring(5,($mbx.ForwardingSMTPAddress.length-5))
                            }
                            else
                            {
                                If ($mbx.ForwardingAddress.length -gt 0)
                                {
                                    $Addr = $mbx.ForwardingAddress.substring(5,($mbx.ForwardingAddress.length-5))
                                }
                            }

                            If ($Addr.Length -gt 0)
                            {
                                $CurMsg = get-MailboxAutoReplyConfiguration $u
                                $LineToWrite = $RecordEvent + "CURR" + "`t" + "      Current AutoReply State: " + $CurMsg.AutoReplyState
                                WriteReportEvent
                                If ($CurMsg.AutoReplyState -eq "Enabled")
                                {
                                    $LineToWrite = $RecordEvent + "CURR" + "`t" + "     Current Internal Message: " + $CurMsg.InternalMessage
                                    WriteReportEvent
                                    $LineToWrite = $RecordEvent + "CURR" + "`t" + "     Current External Message: " + $CurMsg.ExternalMessage
                                    WriteReportEvent
                                }
                                $ExternalMessage = @"
<p>Please note that the UL Solutions email system automatically forwarded your email to $Name at $Addr.</p><p> For more information, please reach out direct to $Name.</p>
"@
                                Set-MailboxAutoReplyConfiguration -Identity $u -AutoReplyState Scheduled -InternalMessage $null -ExternalMessage $ExternalMessage -EndTime (get-date).AddDays(+$Script:txtNoDays.Text)
                                $LineToWrite = $RecordEvent + "CURR" + "`t" + "     Enabled External Message: " + $Addr
                                WriteReportEvent
                                $CurMsg = get-MailboxAutoReplyConfiguration $u
                                $LineToWrite = $RecordEvent + "NEW " + "`t" + "          New AutoReply State: " + $CurMsg.AutoReplyState
                                WriteReportEvent
                                $LineToWrite = $RecordEvent + "NEW " + "`t" + "           AutoReply End Date: " + $CurMsg.EndTime
                                WriteReportEvent
                            }
                            else
                            {
                                Write-host "This mailbox does not have a forwarder set." -ForegroundColor Red
                                $LineToWrite = $RecordEvent + "ERR " + "`t" + "     No Forwarding Configured: " + $u
                                WriteReportEvent
                            }
                        }
                    }
                    else
                    {
                        write-host "`tNo Mailbox found for: " $u -ForegroundColor Red
                        $LineToWrite = $RecordEvent + "ERR " + "`t" + "         No mailbox found for: " + $u
                        WriteReportEvent
                    }
                    $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
                    WriteReportEvent
                }
            }

            If ($Script:chkDisable.Checked -eq $True)
            {
                $ReportFile = "E:\Automation\MAActivities\Reports\OutOfOffice\OutOfOfficeDisable-" + $Script:txtProject.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
                Start-Report
                foreach ($u in $addMember)
                {
                    If ($u.length -gt 0)
                    {
                        write-host "`nProcessing Mailbox: " $u -ForegroundColor Cyan
                        $LineToWrite = $RecordEvent + "CURR" + "`t" + "                  Employee ID: " + $u
                        WriteReportEvent
                        $MbxExists = [bool]($mbx = get-mailbox $u)
                        If ($MbxExists -eq $True)
                        {
                            $CurMsg = get-MailboxAutoReplyConfiguration $u
                            $LineToWrite = $RecordEvent + "CURR" + "`t" + "      Current AutoReply State: " + $CurMsg.AutoReplyState
                            WriteReportEvent
                            Set-MailboxAutoReplyConfiguration -Identity $u -AutoReplyState Disabled
                            $CurMsg = get-MailboxAutoReplyConfiguration $u
                            $LineToWrite = $RecordEvent + "NEW " + "`t" + "          New AutoReply State: " + $CurMsg.AutoReplyState
                            WriteReportEvent
                        }
                        else
                        {
                            write-host "`tNo Mailbox found for: " $u -ForegroundColor Red
                            $LineToWrite = $RecordEvent + "ERR " + "`t" + "         No mailbox found for: " + $u
                            WriteReportEvent
                        }
                        $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
                        WriteReportEvent
                    }
                }
            }
        }
        else
        {
            $output = $wshell.Popup("No employee number or email addresses entered try again.",0,"No Employee Nos.",0+32)
        }
        $Script:txtEmpID.Text = ""
        $Script:txtProject.Text = ""
        $Script:txtNoDays.Text = ""
        $Script:OKButton.Visible = $False
        Publish-Form
    }
}While ($Script:Result -eq "OK")
