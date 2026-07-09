#Set OOO on Mailboxes
#
#    2024/04/18 - SAG - Created Script to move accounts to another AD Container
#
#####################################################

Function Build-MAProtocols
{
    $Script:form = New-Object Windows.Forms.Form 
    $Script:form.FormBorderStyle = "FixedToolWindow" 
    $Script:form.Text = "Enable/Disable Mailbox Protocols"
    $Script:form.StartPosition = "CenterScreen" 
    $Script:form.Width = 500 ; $Script:form.Height = 410  # Make the form wider 

    $Script:Top = 20
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
    ## Enable/Disable OOO
    $Script:chkEnable = New-Object Windows.Forms.RadioButton
        $Script:chkEnable.Left = 140; $Script:chkEnable.Width = 150; $Script:chkEnable.Top = $Script:Top
        $Script:chkEnable.Text = "Enable Protocols" 
        $Script:chkEnable.Checked = $false   # set a default value 
        $Script:chkEnable.add_Click({
            $Script:lblSelect.Text = "Select Protocols to Enable:"
            $Script:OKButton.Visible = $False
        })
        $Script:form.Controls.Add($Script:chkEnable)
    $Script:chkDisable = New-Object Windows.Forms.RadioButton
        $Script:chkDisable.Left = 290; $Script:chkDisable.Width = 150; $Script:chkDisable.Top = $Script:Top
        $Script:chkDisable.Text = "Disable Potocols" 
        $Script:chkDisable.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkDisable)
        $Script:chkDisable.add_Click({
            $Script:lblSelect.Text = "Select Protocols to Disable:"
            $Script:OKButton.Visible = $False
        })

    $Script:Top = $Script:Top + 25
    ## Project Name
    $Script:lblSelect = New-Object System.Windows.Forms.Label   
        $Script:lblSelect.Text = "Select Protocols to Change:"  
        $Script:lblSelect.Top = $Script:Top ; $Script:lblSelect.Left = 10; $Script:lblSelect.Width=150; $Script:lblSelect.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblSelect)    # Add to Form

    $Script:Top = $Script:Top + 25
    ## Enable/Disable MAPI
    $Script:chkMAPI = New-Object Windows.Forms.CheckBox
        $Script:chkMAPI.Left = 30; $Script:chkMAPI.Width = 100; $Script:chkMAPI.Top = $Script:Top
        $Script:chkMAPI.Text = "MAPI" 
        $Script:chkMAPI.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkMAPI)

    ## Enable/Disable POP
    $Script:chkPOP = New-Object Windows.Forms.CheckBox
        $Script:chkPOP.Left = 140; $Script:chkPOP.Width = 100; $Script:chkPOP.Top = $Script:Top
        $Script:chkPOP.Text = "POP" 
        $Script:chkPOP.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkPOP)

    ## Enable/Disable OWA
    $Script:chkOWA = New-Object Windows.Forms.CheckBox
        $Script:chkOWA.Left = 240; $Script:chkOWA.Width = 200; $Script:chkOWA.Top = $Script:Top
        $Script:chkOWA.Text = "OWA (Outlook Web Access)" 
        $Script:chkOWA.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkOWA)

    $Script:Top = $Script:Top + 25
    ## Enable/Disable IMAP
    $Script:chkIMAP = New-Object Windows.Forms.CheckBox
        $Script:chkIMAP.Left = 30; $Script:chkIMAP.Width = 100; $Script:chkIMAP.Top = $Script:Top
        $Script:chkIMAP.Text = "IMAP" 
        $Script:chkIMAP.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkIMAP)

    ## Enable/Disable Mobile
    $Script:chkActSyn = New-Object Windows.Forms.CheckBox
        $Script:chkActSyn.Left = 140; $Script:chkActSyn.Width = 100; $Script:chkActSyn.Top = $Script:Top
        $Script:chkActSyn.Text = "ActiveSync" 
        $Script:chkActSyn.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkActSyn)

    ## Enable/Disable MAC
    $Script:chkMAC = New-Object Windows.Forms.CheckBox
        $Script:chkMAC.Left = 240; $Script:chkMAC.Width = 200; $Script:chkMAC.Top = $Script:Top
        $Script:chkMAC.Text = "MAC (Outlook for MAC)" 
        $Script:chkMAC.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkMAC)

    $Script:Top = $Script:Top + 25
    ## Enable/Disable Mobile
    $Script:chkMobile = New-Object Windows.Forms.CheckBox
        $Script:chkMobile.Left = 30; $Script:chkMobile.Width = 100; $Script:chkMobile.Top = $Script:Top
        $Script:chkMobile.Text = "Outlook Mobile" 
        $Script:chkMobile.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkMobile)

    ## Restrict Send
    $Script:chkSend = New-Object Windows.Forms.CheckBox
        $Script:chkSend.Left = 140; $Script:chkSend.Width = 100; $Script:chkSend.Top = $Script:Top
        $Script:chkSend.Text = "Restrict Send" 
        $Script:chkSend.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkSend)

    ## Restrict Receive
    $Script:chkReceive = New-Object Windows.Forms.CheckBox
        $Script:chkReceive.Left = 240; $Script:chkReceive.Width = 200; $Script:chkReceive.Top = $Script:Top
        $Script:chkReceive.Text = "Restrict Receive" 
        $Script:chkReceive.Checked = $false   # set a default value 
        $Script:form.Controls.Add($Script:chkReceive)

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
    $LineToWrite = $RecordEvent + "STAR" + "`t" + "Enable/Disable Mailbox Protocols has started"
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

Function LogSettings
{
    $LineToWrite = $RecordEvent + $RunSt + "`t" + "                    MAPI Enabled: " + $casmbx.MapiEnabled
    WriteReportEvent                        
    $LineToWrite = $RecordEvent + $RunSt + "`t" + "                     POP Enabled: " + $casmbx.PopEnabled
    WriteReportEvent
    $LineToWrite = $RecordEvent + $RunSt + "`t" + "                    IMAP Enabled: " + $casmbx.ImapEnabled
    WriteReportEvent
    $LineToWrite = $RecordEvent + $RunSt + "`t" + "OWA (Outlook Web Access) Enabled: " + $casmbx.OWAEnabled
    WriteReportEvent
    $LineToWrite = $RecordEvent + $RunSt + "`t" + "              ActiveSync Enabled: " + $casmbx.ActiveSyncEnabled
    WriteReportEvent
    $LineToWrite = $RecordEvent + $RunSt + "`t" + "              MACOutlook Enabled: " + $casmbx.MacOutlookEnabled
    WriteReportEvent
    $LineToWrite = $RecordEvent + $RunSt + "`t" + "          Outlook Mobile Enabled: " + $casmbx.OutlookMobileEnabled
    WriteReportEvent
    $LineToWrite = $RecordEvent + $RunSt + "`t" + "          Outlook Mobile Enabled: " + $casmbx.OWAforDevicesEnabled
    WriteReportEvent
    $LineToWrite = $RecordEvent + $RunSt + "`t" + "                   Max Send Size: " + $mbx.MaxSendSize
    WriteReportEvent
    $LineToWrite = $RecordEvent + $RunSt + "`t" + "                Max Receive Size: " + $mbx.MaxReceiveSize
    WriteReportEvent
}

$wshell = New-Object -ComObject Wscript.Shell
Build-MAProtocols
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
                $ReportFile = "E:\Automation\MAActivities\Reports\Protocols\ProtocolsEnable-" + $Script:txtProject.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            }
            else
            {
                $ReportFile = "E:\Automation\MAActivities\Reports\Protocols\ProtocolsDisable-" + $Script:txtProject.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            }

            Start-Report
            foreach ($u in $addMember)
            {
                write-host "`nProcessing Mailbox: " $u -ForegroundColor Cyan
                $LineToWrite = $RecordEvent + "CURR" + "`t" + "                     Employee ID: " + $u
                WriteReportEvent
                $MbxExists = [bool]($CASmbx = get-casmailbox $u -ErrorAction SilentlyContinue)
                If ($MbxExists -eq $True)
                {
                    $mbx = get-mailbox $u
                    $LineToWrite = ""
                    WriteReportEvent
                    $RunSt = "BEFR"
                    $LineToWrite = $RecordEvent + $RunSt + "`t" + "         Current Settings:"
                    WriteReportEvent
                    LogSettings

                    If ($Script:chkDisable.Checked -eq $True)
                    {
                        If (($casmbx.MapiEnabled -eq $True) -and ($Script:chkMAPI.Checked -eq $True))
                        {
                            Set-CASMailbox $u -MAPIEnabled $False
                            $LineToWrite = $RecordEvent + "CHG " + "`t" + "Disabled MAPI"
                            WriteReportEvent
                        }

                        If (($casmbx.OWAEnabled -eq $True)-and ($Script:chkOWA.Checked -eq $True))
                        {
                            Set-CASMailbox $u -OWAEnabled $False
                            $LineToWrite = $RecordEvent + "CHG " + "`t" + "Disabled OWA (Outlook Web Access)"
                            WriteReportEvent
                        }

                        If (($casmbx.MacOutlookEnabled -eq $True)-and ($Script:chkMAC.Checked -eq $True))
                        {
                            Set-CASMailbox $u -MacOutlookEnabled $False
                            $LineToWrite = $RecordEvent + "CHG " + "`t" + "Disabled MAC Outlook"
                            WriteReportEvent
                        }

                        If (($casmbx.OutlookMobileEnabled -eq $True)-and ($Script:chkMobile.Checked -eq $True))
                        {
                            Set-CASMailbox $u -OutlookMobileEnabled $False
                            $LineToWrite = $RecordEvent + "CHG " + "`t" + "Disabled Outlook Mobile"
                            WriteReportEvent
                        }

                        If (($mbx.MaxSendSize -gt 0) -and ($Script:chkSend.Checked -eq $true))
                        {
                            Set-Mailbox $u -MaxSendSize 0
                            $LineToWrite = $RecordEvent + "CHG " + "`t" + "Set MaxSendSize to 0"
                            WriteReportEvent
                        }

                        If (($mbx.MaxReceiveSize -gt 0) -and ($Script:chkReceive.Checked -eq $true))
                        {
                            Set-Mailbox $u -MaxReceiveSize 0
                            $LineToWrite = $RecordEvent + "CHG " + "`t" + "Set MaxReceiveSize 0"
                            WriteReportEvent
                        }
                    }

                    If ($Script:chkEnable.Checked -eq $True)
                    {
                        If (($casmbx.MapiEnabled -eq $False) -and ($Script:chkMAPI.Checked -eq $True))
                        {
                            Set-CASMailbox $u -MAPIEnabled $True
                            $LineToWrite = $RecordEvent + "CHG " + "`t" + "Enabled MAPI"
                            WriteReportEvent
                        }
                        
                        If (($casmbx.OWAEnabled -eq $False)-and ($Script:chkOWA.Checked -eq $True))
                        {
                            Set-CASMailbox $u -OWAEnabled $True
                            $LineToWrite = $RecordEvent + "CHG " + "`t" + "Enabled OWA (Outlook Web Access)"
                            WriteReportEvent
                        }

                        If (($casmbx.MacOutlookEnabled -eq $False)-and ($Script:chkMAC.Checked -eq $True))
                        {
                            Set-CASMailbox $u -MacOutlookEnabled $True
                            $LineToWrite = $RecordEvent + "CHG " + "`t" + "Enabled MAC Outlook"
                            WriteReportEvent
                        }

                        If (($casmbx.OutlookMobileEnabled -eq $False)-and ($Script:chkMobile.Checked -eq $True))
                        {
                            Set-CASMailbox $u -OutlookMobileEnabled $True
                            $LineToWrite = $RecordEvent + "CHG " + "`t" + "Enabled Outlook Mobile"
                            WriteReportEvent
                        }

                        If (($mbx.MaxSendSize -notlike "*MB*") -and ($Script:chkSend.Checked -eq $true))
                        {
                            Set-Mailbox $u -MaxSendSize 35MB
                            $LineToWrite = $RecordEvent + "CHG " + "`t" + "Set MaxSendSize to 35MB"
                            WriteReportEvent
                        }

                        If (($mbx.MaxReceiveSize -notlike "*MB*") -and ($Script:chkReceive.Checked -eq $true))
                        {
                            Set-Mailbox $u -MaxReceiveSize 36MB
                            $LineToWrite = $RecordEvent + "CHG " + "`t" + "Set MaxReceiveSize 36MB"
                            WriteReportEvent
                        }
                    }
                        
                    If ($casmbx.PopEnabled -eq $True)
                    {
                        Set-CASMailbox $u -PopEnabled $False
                        $LineToWrite = $RecordEvent + "CHG " + "`t" + "Disabled POP this is not allowed by policy"
                        WriteReportEvent
                    }

                    If ($casmbx.ImapEnabled -eq $True)
                    {
                        Set-CASMailbox $u -ImapEnabled $False
                        $LineToWrite = $RecordEvent + "CHG " + "`t" + "Disabled IMap this is not allowed by policy"
                        WriteReportEvent
                    }

                    If ($casmbx.OWAforDevicesEnabled -eq $True)
                    {
                        Set-CASMailbox $u -OWAforDevicesEnabled $False
                        $LineToWrite = $RecordEvent + "CHG " + "`t" + "Disabled OWAforDevices this is not allowed by policy"
                        WriteReportEvent
                    }

                    If ($casmbx.ActiveSyncEnabled -eq $True)
                    {
                        Set-CASMailbox $u -OWAforDevicesEnabled $False
                        $LineToWrite = $RecordEvent + "CHG " + "`t" + "Disabled ActiveSync this is not allowed by policy"
                        WriteReportEvent
                    }

                    $CASmbx = get-CASmailbox $u
                    $mbx = get-mailbox $u
                    $LineToWrite = ""
                    WriteReportEvent
                    $RunSt = "AFTR"
                    $LineToWrite = $RecordEvent + $RunSt + "`t" + "   Settings After Changes:"
                    WriteReportEvent
                    LogSettings
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
        else
        {
            $output = $wshell.Popup("No employee number or email addresses entered try again.",0,"No Employee Nos.",0+32)
        }
        $Script:txtEmpID.Text = ""
        $Script:txtProject.Text = ""
        $Script:chkEnable.Checked = $False
        $Script:chkDisable.Checked = $False
        $Script:chkMAPI.Checked = $False
        $Script:chkPOP.Checked = $False
        $Script:chkOWA.Checked = $False
        $Script:chkIMAP.Checked = $False
        $Script:chkActSyn.Checked = $False
        $Script:chkMAC.Checked = $False
        $Script:chkMobile.Checked = $False
        $Script:chkSend.Checked = $False
        $Script:chkReceive.Checked = $False
        $Script:OKButton.Visible = $False
        Publish-Form
    }
}While ($Script:Result -eq "OK")
