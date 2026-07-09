#Set forwards on User or Shared Mailboxes
#
#    2024/04/18 - SAG - Created Script to Add/Remove Forwarders for a Divestiture.  Allows these changes by providing individual employee numbers or through an input file
#
#####################################################

Function Build-MAForwardersMenu
{
    $Script:form = New-Object Windows.Forms.Form 
    $Script:form.FormBorderStyle = "FixedToolWindow" 
    $Script:form.Text = "Add or Remove Forwarding Addresses"
    $Script:form.StartPosition = "CenterScreen" 
    $Script:form.Width = 400 ; $Script:form.Height = 300  # Make the form wider 
    
    $Script:Left = 60
    $Script:Col2Left = 90
    $Script:Top = 20
    
    ## Multiple Mailboxes
    $Script:chkAddMultiMbx = New-Object System.Windows.Forms.RadioButton
        $Script:chkAddMultiMbx.Text = "Add Forwarders for Muiltiple Mailboxes"
        $Script:chkAddMultiMbx.Top = $Top ; $Script:chkAddMultiMbx.Left = $Left; $Script:chkAddMultiMbx.Width=150 ;$Script:chkAddMultiMbx.AutoSize = $true 
        $form.Controls.Add($Script:chkAddMultiMbx)    # Add to Form
        $Script:chkAddMultiMbx.Add_Click({
            Uncheck-Items
            $Script:chkAddMultiMbx.Checked = $True
            Hide-Details
            $Script:ButViewENo.Visible = $False
            $Script:chkFileExists.Visible = $True
            $Script:chkUseTemplate.Visible = $True
            $Script:lblFileName.Visible = $False
            $Script:txtFileName.Visible = $False
            $Script:lblENo.Visible = $False
            $Script:txtENo.Visible = $False
            $Script:lblForwAddr.visible = $False
            $Script:txtForwAddr.Visible = $false
            $Script:okButton.Text = "Configure"
        })

    $Script:Top = $Script:Top + 20
    ## Individual Mailboxes
    $Script:chkIndvMbx = New-Object System.Windows.Forms.RadioButton
        $Script:chkIndvMbx.Text = "Add Forwarders for an Individual Mailbox"
        $Script:chkIndvMbx.Top = $Top ; $Script:chkIndvMbx.Left = $Left; $Script:chkIndvMbx.Width=150 ;$Script:chkIndvMbx.AutoSize = $true 
        $form.Controls.Add($Script:chkIndvMbx)    # Add to Form
        $Script:chkIndvMbx.Add_Click({
            Uncheck-Items
            $Script:chkIndvMbx.Checked = $True
            Hide-Details
            $Script:ButViewENo.Visible = $False
            $Script:chkFileExists.Visible = $False
            $Script:chkUseTemplate.Visible = $False
            $Script:lblFileName.Visible = $False
            $Script:txtFileName.Visible = $False
            $Script:lblENo.Visible = $True
            $Script:txtENo.Visible = $True
            $Script:lblENo.Text = "Add To: "
            $Script:txtENo.Text = "Enter EmpNo or Email Address"
            $Script:lblForwAddr.visible = $True
            $Script:txtForwAddr.Visible = $True
            $Script:lblKeepCopy.Visible = $True
            $Script:chkKeepCopyYes.Visible = $True
            $Script:chkKeepCopyNo.Visible = $True
            $Script:okButton.Text = "Configure"
        })

    $Script:Top = $Script:Top + 30
    ## Multiple Mailboxes
    $Script:chkRmvMultiMbx = New-Object System.Windows.Forms.RadioButton
        $Script:chkRmvMultiMbx.Text = "Remove Forwarders for Muiltiple Mailboxes"
        $Script:chkRmvMultiMbx.Top = $Top ; $Script:chkRmvMultiMbx.Left = $Left; $Script:chkRmvMultiMbx.Width=150 ;$Script:chkRmvMultiMbx.AutoSize = $true 
        $form.Controls.Add($Script:chkRmvMultiMbx)    # Add to Form
        $Script:chkRmvMultiMbx.Add_Click({
            Uncheck-Items
            $Script:chkRmvMultiMbx.Checked = $True
            Hide-Details
            $Script:ButViewENo.Visible = $False
            $Script:chkFileExists.Visible = $True
            $Script:chkUseTemplate.Visible = $True
            $Script:lblFileName.Visible = $False
            $Script:txtFileName.Visible = $False
            $Script:lblENo.Visible = $False
            $Script:txtENo.Visible = $False
            $Script:lblForwAddr.visible = $False
            $Script:txtForwAddr.Visible = $False
            $Script:okButton.Text = "Remove"
        })

    $Script:Top = $Script:Top + 20
    ## Individual Mailboxes
    $Script:chkRmvIndvMbx = New-Object System.Windows.Forms.RadioButton
        $Script:chkRmvIndvMbx.Text = "Remove Forwarders for an Individual Mailbox"
        $Script:chkRmvIndvMbx.Top = $Top ; $Script:chkRmvIndvMbx.Left = $Left; $Script:chkRmvIndvMbx.Width=150 ;$Script:chkRmvIndvMbx.AutoSize = $true 
        $form.Controls.Add($Script:chkRmvIndvMbx)    # Add to Form
        $Script:chkRmvIndvMbx.Add_Click({
            Uncheck-Items
            $Script:chkRmvIndvMbx.Checked = $True
            Hide-Details
            $Script:ButViewENo.Visible = $False
            $Script:chkFileExists.Visible = $False
            $Script:chkUseTemplate.Visible = $False
            $Script:lblFileName.Visible = $False
            $Script:txtFileName.Visible = $False
            $Script:lblENo.Visible = $True
            $Script:txtENo.Visible = $True
            $Script:lblENo.Text = "Remove From: "
            $Script:txtENo.Text = "Enter EmpNo or Email Address"
            $Script:lblForwAddr.visible = $False
            $Script:txtForwAddr.Visible = $False
            $Script:okButton.Text = "Remove"
        })

    $Top = $Top + 30
    ## Get Details for Individual Mailbox Update
    $Script:lblENo = New-Object System.Windows.Forms.Label
        $Script:lblENo.Top = $Top ; $Script:lblENo.Left = 30; $Script:lblENo.Width=120 ; $Script:lblENo.AutoSize = $true
        $Script:lblENo.Visible = $False
        $Script:form.Controls.Add($Script:lblENo)    # Add to Form 
    $Script:txtENo = New-Object System.Windows.Forms.Textbox
        $Script:txtENo.Text = "Enter EmpNo or Email Address" 
        $Script:txtENo.Top = $Top ; $Script:txtENo.Left = 160; $Script:txtENo.Width=200 ;$Script:txtENo.AutoSize = $true 
        $Script:txtENo.Visible = $False
        $Script:form.Controls.Add($Script:txtENo)    # Add to Form
        $Script:txtENo.Add_Click({
            $Script:txtENo.Text = ""
            $Script:ButViewENo.Visible = $True
           })

    $Script:lblForwAddr = New-Object System.Windows.Forms.Label
        $Script:lblForwAddr.Text = "Forwarding Address: "
        $Script:lblForwAddr.Top = ($Top + 20) ; $Script:lblForwAddr.Left = 30; $Script:lblForwAddr.Width=120 ; $Script:lblForwAddr.AutoSize = $true
        $Script:lblForwAddr.Visible = $False
        $Script:form.Controls.Add($Script:lblForwAddr)    # Add to Form 
    $Script:txtForwAddr = New-Object System.Windows.Forms.Textbox
        $Script:txtForwAddr.Text = "" 
        $Script:txtForwAddr.Top = ($Top + 20) ; $Script:txtForwAddr.Left = 160; $Script:txtForwAddr.Width=200 ;$Script:txtForwAddr.AutoSize = $true 
        $Script:txtForwAddr.Visible = $False
        $Script:form.Controls.Add($Script:txtForwAddr)    # Add to Form

    $Script:lblKeepCopy = New-Object System.Windows.Forms.Label
        $Script:lblKeepCopy.Text = "Keep A Copy:"
        $Script:lblKeepCopy.Top = ($Top + 40) ; $Script:lblKeepCopy.Left = 30; $Script:lblKeepCopy.Width=150 ;$Script:lblKeepCopy.AutoSize = $true 
        $Script:lblKeepCopy.Visible = $False
        $Script:form.Controls.Add($Script:lblKeepCopy)    # Add to Form 
    $Script:chkKeepCopyYes = New-Object System.Windows.Forms.Checkbox
        $Script:chkKeepCopyYes.Text = "Yes"
        $Script:chkKeepCopyYes.Top = ($Top + 40) ; $Script:chkKeepCopyYes.Left = 160; $Script:chkKeepCopyYes.Width=50 ;$Script:chkKeepCopyYes.AutoSize = $true 
        $Script:chkKeepCopyYes.Visible = $False
        $Script:chkKeepCopyYes.Checked = $False
        $Script:form.Controls.Add($Script:chkKeepCopyYes)    # Add to Form
        $Script:chkKeepCopyYes.Add_Click({
            $Script:chkKeepCopyNo.Checked = $False
           })
    $Script:chkKeepCopyNo = New-Object System.Windows.Forms.Checkbox
        $Script:chkKeepCopyNo.Text = "No"
        $Script:chkKeepCopyNo.Top = ($Top + 40) ; $Script:chkKeepCopyNo.Left = 210; $Script:chkKeepCopyNo.Width=50 ;$Script:chkKeepCopyNo.AutoSize = $true 
        $Script:chkKeepCopyNo.Visible = $False
        $Script:chkKeepCopyNo.Checked = $True
        $Script:form.Controls.Add($Script:chkKeepCopyNo)    # Add to Form
        $Script:chkKeepCopyNo.Add_Click({
            $Script:chkKeepCopyYes.Checked = $False
           })

    $Script:ButViewENo = New-Object Windows.Forms.Button
        $Script:ButViewENo.Location = New-object System.Drawing.Size(145,($Top+75))
        $Script:ButViewENo.Size = new-Object System.Drawing.Size(140,20)
        $Script:ButViewENo.Text = "Get Current Details"
        $Script:ButViewENo.TabIndex = 10
        $Script:ButViewENo.Visible = $False
        $form.Controls.Add($Script:ButViewENo)
        $Script:ButViewENo.Add_Click({
            If ([bool]($mbx = get-mailbox $Script:txtENo.Text -ErrorAction SilentlyContinue))
            {
                #Pouplate details and unhide the OK button
                $Script:txtMbxName.Text = $mbx.DisplayName
                $Script:txtCurrForwAddr.Text = $mbx.ForwardingSMTPAddress
                If ($mbx.ForwardingSMTPAddress.Length -gt -0)
                {
                    $Script:txtCurrForwAddr.Text = $mbx.ForwardingSMTPAddress
                }
                else
                {
                    $Script:txtCurrForwAddr.Text = $mbx.ForwardingAddress
                }
                If ($Script:txtCurrForwAddr.Text -ne $null)
                {
                    If ($Script:chkIndvMbx.Checked -eq $True)
                    {
                        $Script:OKButton.Text = "Update"
                    }
                    else
                    {
                        $Script:OKButton.Text = "Remove"
                    }
                }
                else
                {
                    $Script:OKButton.Text = "Continue"
                }
                $Script:txtDelivSet.Text = $mbx.DeliverToMailboxAndForward

                $Script:form.Width = 400 ; $Script:form.Height = 400  # Make the form wider
                $Script:lblMbxName.Visible = $True
                $Script:txtMbxName.Visible = $True
                $Script:lblCurrForwAddr.Visible = $True
                $Script:txtCurrForwAddr.Visible = $True
                $Script:lblDelivSet.Visible = $True
                $Script:txtDelivSet.Visible = $True
                $Script:OKButton.Visible = $True
            }
            else
            {
                $Output = $wshell.Popup("No mailbox found for: " + $Script:txtENo.Text,0,"Check Column Titles",0+32)
            }
            $Script:ButViewENo.Visible = $False
            $Global:OKButton.Visible = $True
        })

    $MbxTop = $Top + 95
    #Current Mailbox User
    $Script:lblMbxName = New-Object System.Windows.Forms.Label
            $Script:lblMbxName.Text = "Mailbox DisplayName: "
            $Script:lblMbxName.Top = $MbxTop ; $Script:lblMbxName.Left = 30; $Script:lblMbxName.Width=120 ; $Script:lblMbxName.AutoSize = $true
            $Script:lblMbxName.Visible = $False
            $Script:form.Controls.Add($Script:lblMbxName)    # Add to Form 
        $Script:txtMbxName = New-Object System.Windows.Forms.Textbox
            $Script:txtMbxName.Text = "" 
            $Script:txtMbxName.Top = $MbxTop ; $Script:txtMbxName.Left = 160; $Script:txtMbxName.Width=200 ;$Script:txtMbxName.AutoSize = $true 
            $Script:txtMbxName.Visible = $False
            $Script:form.Controls.Add($Script:txtMbxName)    # Add to Form

    $MbxTop = $MbxTop + 20
    #Current Forwarding Address
    $Script:lblCurrForwAddr = New-Object System.Windows.Forms.Label
            $Script:lblCurrForwAddr.Text = "Current Forwarder: "
            $Script:lblCurrForwAddr.Top = $MbxTop ; $Script:lblCurrForwAddr.Left = 30; $Script:lblCurrForwAddr.Width=120 ; $Script:lblCurrForwAddr.AutoSize = $true
            $Script:lblCurrForwAddr.Visible = $False
            $Script:form.Controls.Add($Script:lblCurrForwAddr)    # Add to Form 
        $Script:txtCurrForwAddr = New-Object System.Windows.Forms.Textbox
            $Script:txtCurrForwAddr.Text = "" 
            $Script:txtCurrForwAddr.Top = $MbxTop ; $Script:txtCurrForwAddr.Left = 160; $Script:txtCurrForwAddr.Width=200 ;$Script:txtCurrForwAddr.AutoSize = $true 
            $Script:txtCurrForwAddr.Visible = $False
            $Script:form.Controls.Add($Script:txtCurrForwAddr)    # Add to Form

    $MbxTop = $MbxTop + 20
    #Current Deliver Status
    $Script:lblDelivSet = New-Object System.Windows.Forms.Label
            $Script:lblDelivSet.Text = "Keep A Copy: "
            $Script:lblDelivSet.Top = $MbxTop ; $Script:lblDelivSet.Left = 30; $Script:lblDelivSet.Width=120 ; $Script:lblDelivSet.AutoSize = $true
            $Script:lblDelivSet.Visible = $False
            $Script:form.Controls.Add($Script:lblDelivSet)    # Add to Form 
        $Script:txtDelivSet = New-Object System.Windows.Forms.Textbox
            $Script:txtDelivSet.Text = "" 
            $Script:txtDelivSet.Top = $MbxTop ; $Script:txtDelivSet.Left = 160; $Script:txtDelivSet.Width=200 ;$Script:txtDelivSet.AutoSize = $true 
            $Script:txtDelivSet.Visible = $False
            $Script:form.Controls.Add($Script:txtDelivSet)    # Add to Form

    $FwDetails = $Top + 70
    ## Input File Details
    $Script:lblUsrName = New-Object System.Windows.Forms.Label
        $Script:lblUsrName.Text = "User Name: "
        $Script:lblUsrName.Top = $FwDetails ; $Script:lblUsrName.Left = 30; $Script:lblUsrName.Width=120 ; $Script:lblUsrName.AutoSize = $true
        $Script:lblUsrName.Visible = $False
        $Script:form.Controls.Add($Script:lblUsrName)    # Add to Form 
    $Script:txtUsrName = New-Object System.Windows.Forms.Textbox
        $Script:txtUsrName.Text = "(Sample: path:\Name)" 
        $Script:txtUsrName.Top = $FwDetails ; $Script:txtUsrName.Left = 160; $Script:txtUsrName.Width=200 ;$Script:txtUsrName.AutoSize = $true 
        $Script:txtUsrName.Visible = $False
        $Script:form.Controls.Add($Script:txtUsrName)    # Add to Form

    ## Input File Details
    $Script:lblCurrFw = New-Object System.Windows.Forms.Label
        $Script:lblCurrFw.Text = "Current Forwarder: "
        $Script:lblCurrFw.Top = $FwDetails+20 ; $Script:lblCurrFw.Left = 30; $Script:lblCurrFw.Width=120 ; $Script:lblCurrFw.AutoSize = $true
        $Script:lblCurrFw.Visible = $False
        $Script:form.Controls.Add($Script:lblCurrFw)    # Add to Form 
    $Script:txtCurrFw = New-Object System.Windows.Forms.Textbox
        $Script:txtCurrFw.Text = "(Sample: path:\Name)" 
        $Script:txtCurrFw.Top = $FwDetails+20 ; $Script:txtCurrFw.Left = 160; $Script:txtCurrFw.Width=200 ;$Script:txtCurrFw.AutoSize = $true 
        $Script:txtCurrFw.Visible = $False
        $Script:form.Controls.Add($Script:txtCurrFw)    # Add to Form

    ## Use Template or give name of file
    $Script:chkFileExists = New-Object System.Windows.Forms.Checkbox
        $Script:chkFileExists.Text = "Input File Exists"
        $Script:chkFileExists.Top = $Top ; $Script:chkFileExists.Left = 50; $Script:chkFileExists.Width=150 ;$Script:chkFileExists.AutoSize = $true
        $Script:chkFileExists.Visible = $False
        $form.Controls.Add($Script:chkFileExists)    # Add to Form
        $Script:chkFileExists.Add_Click({
            $Output = $wshell.Popup("Column titles in your input file must be 'Mbx,Forwarder'",0,"Check Column Titles",0+32)
            $Script:lblFileName.Visible = $True
            $Script:txtFileName.Visible = $True
            $Script:lblKeepCopy.Visible = $True
            $Script:chkKeepCopyYes.Visible = $True
            $Script:chkKeepCopyNo.Visible = $True
            $Script:OKButton.Visible = $True
        })

    $Script:FileDetails

    ## Use Template or give name of file
    $Script:chkUseTemplate = New-Object System.Windows.Forms.Checkbox
        $Script:chkUseTemplate.Text = "Create Using Template"
        $Script:chkUseTemplate.Top = $Top ; $Script:chkUseTemplate.Left = 170; $Script:chkUseTemplate.Width=150 ;$Script:chkUseTemplate.AutoSize = $true
        $Script:chkUseTemplate.Visible = $False
        $form.Controls.Add($Script:chkUseTemplate)    # Add to Form
        $Script:chkUseTemplate.Add_Click({
            $Script:chkFileExists.Checked = $False
            $Output = $wshell.Popup("Opening the template file in excel ths will take a few moments.  The file is set to read-only use SaveAs to save your changes",0,"Create Input File",0+32)
            Invoke-Expression -Command e:\O365AdminShared\Data\MAForwdersTemplate.csv
            $Script:lblFileName.Visible = $True
            $Script:txtFileName.Visible = $True
            $Script:lblKeepCopy.Visible = $True
            $Script:chkKeepCopyYes.Visible = $True
            $Script:chkKeepCopyNo.Visible = $True
            $Script:OKButton.Visible = $True
        })

    $Top = $Top + 20
    ## Input File Details
    $Script:lblFileName = New-Object System.Windows.Forms.Label
        $Script:lblFileName.Text = "Enter Path:\Filename: "
        $Script:lblFileName.Top = $Top ; $Script:lblFileName.Left = 30; $Script:lblFileName.Width=120 ; $Script:lblFileName.AutoSize = $true
        $Script:lblFileName.Visible = $False
        $Script:form.Controls.Add($Script:lblFileName)    # Add to Form 
    $Script:txtFileName = New-Object System.Windows.Forms.Textbox
        $Script:txtFileName.Text = "(Sample: path:\Name)" 
        $Script:txtFileName.Top = $Top ; $Script:txtFileName.Left = 160; $Script:txtFileName.Width=200 ;$Script:txtFileName.AutoSize = $true 
        $Script:txtFileName.Visible = $False
        $Script:form.Controls.Add($Script:txtFileName)    # Add to Form
        $Script:txtFileName.Add_Click({
            If ($Script:txtFileName.Text -eq "(Sample: path:\Name)")
            {
                $Script:txtFileName.Text = ""
            }
        })
}

Function Uncheck-Items
{
    $Script:chkAddMultiMbx.Checked = $False
    $Script:chkIndvMbx.Checked = $False
    $Script:chkRmvMultiMbx.Checked = $False
    $Script:chkRmvIndvMbx.Checked = $False
    $Script:chkFileExists.Checked = $False
    $Script:chkUseTemplate.Checked = $False
    $Script:chkFileExists.Checked = $False
    $Script:chkUseTemplate.Checked = $Flase
    $Script:txtFileName.Text = "(Sample: path:\Name)"
    $Script:chkKeepCopyYes.Checked = $False
    $Script:chkKeepCopyNo.Checked = $True 
}

Function Hide-Details
{
    $Script:form.Width = 400 ; $Script:form.Height = 300  # Make the form wider
    $Script:lblMbxName.Visible = $false
    $Script:txtMbxName.Visible = $False
    $Script:lblCurrForwAddr.Visible = $False
    $Script:txtCurrForwAddr.Visible = $False
    $Script:lblDelivSet.Visible = $False
    $Script:txtDelivSet.Visible = $False
    $Script:lblENo.Visible = $False
    $Script:txtENo.Visible = $False
    $Script:lblForwAddr.Visible = $False
    $Script:txtForwAddr.Visible = $False
    $Script:chkFileExists.Visible = $False
    $Script:chkUseTemplate.Visible = $False
    $Script:lblFileName.Visible = $False
    $Script:txtFileName.Visible = $False
    $Script:lblKeepCopy.Visible = $False
    $Script:chkKeepCopyYes.Visible = $False
    $Script:chkKeepCopyNo.Visible = $False
    $Script:OKButton.Visible = $False
}

Function Start-Report
{
    $LineToWrite = $RecordEvent + "STAR" + "`t" + "MAForwarders script has started"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + (whoami)
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "User TimeZone: " + (Get-TimeZone)
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Number of Mailboxes to Process: " + ($InpFile.Count)
    WriteReportEvent
    $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
    WriteReportEvent
}

Function Existing-Config
{
    $LineToWrite = $RecordEvent + "CURR" + "`t" + "                 Mailbox: " + $mbx.Alias
    WriteReportEvent
    $LineToWrite = $RecordEvent + "CURR" + "`t" + "     Mailbox DisplayName: " + $mbx.DisplayName

    WriteReportEvent
    If (($mbx.ForwardingSMTPAddress -eq $Null) -and ($mbx.ForwardingAddress -eq $Null))
    {
        $LineToWrite = $RecordEvent + "CURR" + "`t" + " No Forwarding Addresses configured"
    }
    else
    {
        If ($mbx.ForwardingSMTPAddress -ne $Null)
        {
            $LineToWrite = $RecordEvent + "CURR" + "`t" + " Current SMTP Forwarding: " + $mbx.ForwardingSMTPAddress
        }
        else
        {
            If ($mbx.ForwardingAddress -ne $Null)
            {  
                $LineToWrite = $RecordEvent + "CURR" + "`t" + "      Current Forwarding: " + $mbx.ForwardingAddress
            }
        }
    }
    WriteReportEvent
    $LineToWrite = $RecordEvent + "CURR" + "`t" + "Current Delivery Setting: " + $mbx.DeliverToMailboxAndForward
    WriteReportEvent
}

Function New-Config
{
    $LineToWrite = $RecordEvent + "NEW " + "`t" + "           New Forwarder: " + $mbx.ForwardingSMTPAddress
    WriteReportEvent
    $LineToWrite = $RecordEvent + "NEW " + "`t" + "  Keep A Copy of Message: " + $mbx.DeliverToMailboxAndForward
    WriteReportEvent
    $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
    WriteReportEvent
}

$wshell = New-Object -ComObject Wscript.Shell
Build-MAForwardersMenu
Add-FormStandardButtons
$Script:OKButton.Visible = $False
Publish-Form

Do
{
    If ($Script:Result -eq "OK")
    {
        $KeepCopy = $False
        If ($Script:chkKeepCopyYes.Checked -eq $True)
        {
            $KeepCopy = $True
        }

        If ($Script:chkAddMultiMbx.Checked -eq $True)
        {
            #Add Using Input file
            write-host "Add Using Input File"
            If (Test-Path $Script:txtFileName.Text)
            {
                $InpFile = Import-csv $Script:txtFileName.Text
                If (($InpFile.Mbx.Count -gt 0) -and ($InpFile.Forwarder.Count -gt 0))
                {
                    $ReportFile	= "E:\Automation\MAActivities\Reports\AddForwarders\Report-AddForwarders-" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
                    Start-Report

                    write-host "Number of Mailboxes to Process: " $InpFile.Count
                    ForEach ($Inp in $InpFile)
                    {
                        write-host "Processing: " $Inp.Mbx -NoNewline
                        If ([bool]($mbx = get-mailbox $inp.Mbx -ErrorAction SilentlyContinue))
                        {
                            write-host " - Current Forwarder: $mbx.ForwardingSMTPAddress $mbx.ForwardingAddress -DeliveryStatus: " $mbx.DeliverToMailboxAndForward
                            Existing-Config
                            set-mailbox $inp.Mbx -ForwardingSmtpAddress $Inp.Forwarder -DeliverToMailboxAndForward $KeepCopy
                            $mbx = get-mailbox $inp.Mbx
                            New-Config
                        }
                    }
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Additional of Mailbox Forwarders Proecessing Complete"
                    WriteReportEvent
                }
                else
                {
                    $Output = $wshell.Popup("There are " + $InpFile.count + " records in this file but the column titles do not match the 'Mbx,Forwarder' titles.  Please review the column titles on the input file.",0,"Invalid Column Titles",0+32)
                }
            }
            else
            {
                write-host "File Not Found"
            }
        }

        If ($Script:chkIndvMbx.Checked -eq $True)
        {
            #Add Individual
            write-host "Add for Individual Mailbox"
            #write Current settings to a log file
            write-host "DisplayName: " $mbx.DisplayName
            Write-host "Current Forwarding Address: " $Script:txtCurrForwAddr
            Write-host "Current Delivery Status: " $Mbx.DeliverToMailboxAndForward

            $ReportFile	= "E:\Automation\MAActivities\Reports\AddForwarders\Report-AddForwarders-" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            Start-Report
            Existing-Config
            set-mailbox $Script:txtENo.Text -ForwardingSmtpAddress $Script:txtForwAddr.Text -DeliverToMailboxAndForward $KeepCopy
            $mbx = get-mailbox $Script:txtENo.Text
            New-Config
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Addition of Mailbox Forwarder Proecessing Complete"
            WriteReportEvent
        }

        If ($Script:chkRmvMultiMbx.Checked -eq $True)
        {
            #Remove using input file
            write-host "Remove Using Input File"
            If (Test-Path $Script:txtFileName.Text)
            {
                $InpFile = Import-csv $Script:txtFileName.Text
                If (($InpFile.Mbx.Count -gt 0) -and ($InpFile.Forwarder.Count -gt 0))
                {
                    $ReportFile	= "E:\Automation\MAActivities\Reports\RemoveForwarders\Report-RemoveForwarders-" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
                    write-host "Number of Mailboxes to Process: " $InpFile.Count
                    ForEach ($Inp in $InpFile)
                    {
                        write-host "Processing: " $Inp.Mbx -NoNewline
                        If ([bool]($mbx = get-mailbox $inp.Mbx -ErrorAction SilentlyContinue))
                        {
                            Existing-Config
                            write-host " - Current Forwarder: $mbx.ForwardingSMTPAddress $mbx.ForwardingAddress -DeliveryStatus: " $mbx.DeliverToMailboxAndForward
                            set-mailbox $Inp.Mbx -ForwardingSmtpAddress $null -ForwardingAddress $null
                            $mbx = get-mailbox $inp.Mbx
                            New-Config
                        }
                    }
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Removal of Mailbox Forwarders Proecessing Complete"
                    WriteReportEvent
                }
                else
                {
                    $Output = $wshell.Popup("There are " + $InpFile.count + " records in this file but the column titles do not match the 'Mbx,Forwarder' titles.  Please review the column titles on the input file.",0,"Invalid Column Titles",0+32)
                }
            }
            else
            {
                write-host "File Not Found"
            }
        }

        If ($Script:chkRmvIndvMbx.Checked -eq $True)
        {
            #Remove Individual
            write-host "Remove for Individual Mailbox"
            #write Current settings to a log file
            write-host "DisplayName: " $mbx.DisplayName
            Write-host "Current Forwarding Address: " $Script:txtCurrForwAddr
            Write-host "Current Delivery Status: " $Mbx.DeliverToMailboxAndForward

            $ReportFile	= "E:\Automation\MAActivities\Reports\RemoveForwarders\Report-RemoveForwarders-" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            Start-Report
            Existing-Config
            set-mailbox $Script:txtENo.Text -ForwardingSmtpAddress $null -ForwardingAddress $null
            $mbx = get-mailbox $Script:txtENo.Text
            New-Config
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Removal of Mailbox Forwarder Proecessing Complete"
            WriteReportEvent
        }
        Uncheck-Items
        Hide-Details
        Publish-Form
    }
}While ($Script:Result -eq "OK")