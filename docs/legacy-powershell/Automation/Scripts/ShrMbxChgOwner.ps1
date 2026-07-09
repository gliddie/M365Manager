<#
#  This script changes the owner on a shared mailbox
#
#  08/27/2021 - Changed the Add/Remove/Replace owner to be a radio button rather than a checkbox
#  08/16/2022 - Modified to check for owners at the top of the information store
#  06/29/2023 - Adding code to add logging of changes
#  09/26/2024 - Modified the Mailbox Ownership details to display the DisplayName rather than the Name of the mailbox which is now a GUID
#>

function Build-ChgOwnerInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Change Shared Mailbox Owner" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(450,250) #(W,H)
    $BldDetails = "N"

    ## Get Details used to create
    ## Display Name
    $Script:lblDispName = New-Object System.Windows.Forms.Label   
        $Script:lblDispName.Text = "Shared Mailbox Name:"
        $Script:lblDispName.Top = 20 ; $Script:lblDispName.Left = 10; $Script:lblDispName.Width=120 ;$Script:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblDispName)    # Add to Form 
        # 
        $Script:txtDispName = New-Object Windows.Forms.TextBox  
        $Script:txtDispName.TabIndex = 0 # set Tab Order 
        $Script:txtDispName.Top = 20; $Script:txtDispName.Left = 130; $Script:txtDispName.Width = 280;  
        $Script:txtDispName.Text = ""   # DisplayName
        $Script:txtDispName.TabIndex = 0
        $Global:form.Controls.Add($Script:txtDispName)    # Add to Form

    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 50 ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtInpTaskNo.Top = 50; $Global:txtInpTaskNo.Left = 130; $Global:txtInpTaskNo.Width = 120;  
        $Global:txtInpTaskNo.Text = "TASK"   # Enter ticket number
        $Global:txtInpTaskNo.TabIndex = 1
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form

        Add-FormStandardButtons
}

function Build-ShrMbxChgDetailsForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Change Shared Mailbox Owner" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(440,390) #(W,H)

    ## Get Details to complete creation
    $TopLoc = 10
    ## Display Name
    $Script:lblDispName = New-Object System.Windows.Forms.Label   
        $Script:lblDispName.Text = "Shared Mailbox Name:"
        $Script:lblDispName.Top = $TopLoc; $Script:lblDispName.Left = 10; $Script:lblDispName.Width=150 ;$Script:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblDispName)    # Add to Form 
        # 
        $Script:txtDispName = New-Object Windows.Forms.TextBox  
        $Script:txtDispName.TabIndex = 0 # set Tab Order 
        $Script:txtDispName.Top = $TopLoc; $Script:txtDispName.Left = 130; $Script:txtDispName.Width = 220;  
        $Script:txtDispName.Text = $ShrMbx.DisplayName   # DisplayName
        $Global:form.Controls.Add($Script:txtDispName)    # Add to Form

    $TopLoc = $TopLoc + 30
    ## Ticket Number
    $Script:lblTicketNo = New-Object System.Windows.Forms.Label   
        $Script:lblTicketNo.Text = "Ticket Number:"
        $Script:lblTicketNo.Top = $TopLoc; $Script:lblTicketNo.Left = 10; $Script:lblTicketNo.Width=150 ;$Script:lblTicketNo.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblTicketNo)    # Add to Form 
        # 
        $Script:txtTicketNo = New-Object Windows.Forms.TextBox  
        $Script:txtTicketNo.TabIndex = 0 # set Tab Order 
        $Script:txtTicketNo.Top = $TopLoc; $Script:txtTicketNo.Left = 130; $Script:txtTicketNo.Width = 220;  
        $Script:txtTicketNo.Text = $Global:txtInpTaskNo.Text   # DisplayName
        $Global:form.Controls.Add($Script:txtTicketNo)    # Add to Form

    $TopLoc = $TopLoc + 30
    ## Mailbox Owner
    $Script:lblGrpOwnr = New-Object System.Windows.Forms.Label   
        $Script:lblGrpOwnr.Text = "Mailbox Owner(s):"  
        $Script:lblGrpOwnr.Top = $TopLoc; $Script:lblGrpOwnr.Left = 10; $Script:lblGrpOwnr.Width=150; $Script:lblGrpOwnr.AutoSize = $true 
        $Global:form.Controls.Add($lblGrpOwnr)    # Add to Form 
        # 
        $Script:txtGrpOwnr = New-Object Windows.Forms.TextBox  
        $Script:txtGrpOwnr.TabIndex = 0 # set Tab Order 
        $Script:txtGrpOwnr.Top = $TopLoc; $Script:txtGrpOwnr.Left = 130; $Script:txtGrpOwnr.Width = 220;
        If ($Script:MbxOwner -eq "")
        {
            $Script:MbxOwner = "(None)"
        } 
        $Script:txtGrpOwnr.Text = $Script:MbxOwner
        $Global:form.Controls.Add($Script:txtGrpOwnr)    # Add to Form
        
    $TopLoc = $TopLoc + 30
    ## Mailbox Tip
    $Script:lblMbxTip = New-Object System.Windows.Forms.Label   
        $Script:lblMbxTip.Text = "Mailbox Tip:"  
        $Script:lblMbxTip.Top = $TopLoc; $Script:lblMbxTip.Left = 10; $Script:lblMbxTip.Width=150; $Script:lblMbxTip.AutoSize = $true 
        $Global:form.Controls.Add($lblMbxTip)    # Add to Form 
        # 
        $Script:txtMbxTip = New-Object Windows.Forms.TextBox  
        $Script:txtMbxTip.TabIndex = 0 # set Tab Order 
        $Script:txtMbxTip.Top = $TopLoc; $Script:txtMbxTip.Left = 130; $Script:txtMbxTip.Width = 220;
        $start = $ShrMbx.MailTip.IndexOf("y>")
        $end = $ShrMbx.MailTip.IndexOf("</")
        $Tip = $ShrMbx.MailTip.Substring($start+3,($end-$start-3))
        $Script:txtMbxTip.Text = $Tip
        $Global:form.Controls.Add($Script:txtMbxTip)    # Add to Form          

    If ($Global:EDAccess -ne "")
    {
        $TopLoc = $TopLoc + 30
        ## Editor Old Group Name
        $Script:lblEDGrpName = New-Object System.Windows.Forms.Label
        $Script:lblEDGrpName.Text = ".ED Group:"  
        $Script:lblEDGrpName.Top = $TopLoc ; $Script:lblEDGrpName.Left = 10; $Script:lblEDGrpName.Width=150; $Script:lblEDGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblEDGrpName)    # Add to Form 
        #
        $Script:txtEDGrpName = New-Object Windows.Forms.TextBox  
        $Script:txtEDGrpName.TabIndex = 0 # set Tab Order 
        $Script:txtEDGrpName.Top = $TopLoc; $Script:txtEDGrpName.Left = 130; $Script:txtEDGrpName.Width = 220;  
        $Script:txtEDGrpName.Text = $Global:EDAccess
        $Global:form.Controls.Add($Script:txtEDGrpName)    # Add to Form 
    }

    If ($AUAccess -ne "")
    {
        $TopLoc = $TopLoc + 30
        ## Author Group Name
        $Script:lblAUGrpName = New-Object System.Windows.Forms.Label
        $Script:lblAUGrpName.Text = ".AU Group:" 
        $Script:lblAUGrpName.Top = $TopLoc; $Script:lblAUGrpName.Left = 10; $Script:lblAUGrpName.Width=150; $Script:lblAUGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblAUGrpName)    # Add to Form 
        #
        $Script:txtAUGrpName = New-Object Windows.Forms.TextBox  
        $Script:txtAUGrpName.TabIndex = 0 # set Tab Order 
        $Script:txtAUGrpName.Top = $TopLoc; $Script:txtAUGrpName.Left = 130; $Script:txtAUGrpName.Width = 220;  
        $Script:txtAUGrpName.Text = $Global:AUAccess   # Use Corrent computer name as default 
        $Global:form.Controls.Add($Script:txtAUGrpName)    # Add to Form
    }

    If ($REAccess -ne "")
    {
        $TopLoc = $TopLoc + 30
        # Reader Group Name
        $Script:lblREGrpName = New-Object System.Windows.Forms.Label
        $Script:lblREGrpName.Text = ".RE Group:" 
        $Script:lblREGrpName.Top = $TopLoc ; $Script:lblREGrpName.Left = 10; $Script:lblREGrpName.Width=150; $Script:lblREGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblREGrpName)    # Add to Form 
        #
        $Global:txtREGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtREGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtREGrpName.Top = $TopLoc; $Global:txtREGrpName.Left = 130; $Global:txtREGrpName.Width = 220;  
        $Global:txtREGrpName.Text = $Global:REAccess
        $Global:form.Controls.Add($Global:txtREGrpName)    # Add to Form 
    }

    $TopLoc = $TopLoc + 30
    ## Employee # to Add/Remove
        $Script:lblEmpNo = New-Object System.Windows.Forms.Label
        $Script:lblEmpNo.Text = "Emp# or Address:" 
        $Script:lblEmpNo.Top = $TopLoc ; $Script:lblEmpNo.Left = 10; $Script:lblEmpNo.Width=150; $Script:lblEmpNo.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblEmpNo)    # Add to Form 
        #
        $Script:txtEmpNo = New-Object Windows.Forms.TextBox  
        $Script:txtEmpNo.TabIndex = 0 # set Tab Order 
        $Script:txtEmpNo.Top = $TopLoc; $Script:txtEmpNo.Left = 130; $Script:txtEmpNo.Width = 220;  
        $Script:txtEmpNo.Text = ""
        $Global:form.Controls.Add($Script:txtEmpNo)

    $TopLoc = $TopLoc + 30
    ## Add Owner
    $Script:chkAddOwner = New-Object Windows.Forms.RadioButton
        $Script:chkAddOwner.Left = 50; $Script:chkAddOwner.Width = 80; $Script:chkAddOwner.Top = $TopLoc
        $Script:chkAddOwner.Text = "Add Owner" 
        $Script:chkAddOwner.Checked = $false   # set a default value 
        $Script:chkAddOwner.TabIndex = 5
        $Global:form.Controls.Add($Script:chkAddOwner) 
        # Obtain Value with: $Script:chkRemoveOwner.Checked
               
    ## Remove Owner
    $Script:chkRemoveOwner = New-Object Windows.Forms.RadioButton
        $Script:chkRemoveOwner.Left = 150; $Script:chkRemoveOwner.Width = 110; $Script:chkRemoveOwner.Top = $TopLoc
        $Script:chkRemoveOwner.Text = "Remove Owner" 
        $Script:chkRemoveOwner.Checked = $false   # set a default value 
        $Script:chkRemoveOwner.TabIndex = 5
        $Global:form.Controls.Add($Script:chkRemoveOwner) 
        # Obtain Value with: $Script:chkRemoveOwner.Checked

    ## Replace Owner
    $Script:chkReplaceOwner = New-Object Windows.Forms.RadioButton
        $Script:chkReplaceOwner.Left = 270; $Script:chkReplaceOwner.Width = 120; $Script:chkReplaceOwner.Top = $TopLoc
        $Script:chkReplaceOwner.Text = "Replace Owners" 
        $Script:chkReplaceOwner.Checked = $false   # set a default value 
        $Script:chkReplaceOwner.TabIndex = 5
        $Global:form.Controls.Add($Script:chkReplaceOwner) 
        # Obtain Value with: $Script:chkReplaceOwner.Checked

    $TopLoc = $TopLoc + 30
    ## Authorized Requestor
    $Script:chkAuthorized = New-Object Windows.Forms.checkbox 
        $Script:chkAuthorized.Left = 130; $Script:chkAuthorized.Width = 200; $Script:chkAuthorized.Top = $TopLoc
        $Script:chkAuthorized.Text = "Authorized Requestor" 
        $Script:chkAuthorized.Checked = $false   # set a default value 
        $Script:chkAuthorized.TabIndex = 5
        $Global:form.Controls.Add($Script:chkAuthorized) 
        # Obtain Value with: $Script:chkAuthorized.Checked

        Add-FormStandardButtons
}

Function Chg-Owner
{
    If ($Script:chkReplaceOwner.checked -eq "Checked")
    {
        write-host "Replacing Owners in " $GrpName
        $LineToWrite = "CHGOWN" + "`t" + "Replacing Owners in: " + $GrpName
        WriteReportEvent

        If ($Script:txtEmpNo.Text -notlike "*,*")
        {
            Set-DistributionGroup $GrpName -ManagedBy $Script:txtEmpNo.Text -BypassSecurityGroupManagerCheck
        }
        else
        {
            $Own = ($Script:txtEmpNo.Text -split ",").Trim()
            $first = "Y"
            foreach ($own in $own)
            {
                If ($first -eq "Y")
                {
                    Set-DistributionGroup $GrpName -ManagedBy $Own -BypassSecurityGroupManagerCheck
                    $first = "N"
                }
                else
                {
                    Set-DistributionGroup $GrpName -ManagedBy @{add=$Own} -BypassSecurityGroupManagerCheck
                }
                $LineToWrite = "CHGOWN" + "`t" + "Adding " + $Own + " in: " + $GrpName
                WriteReportEvent
            }
        }
    }
    If ($Script:chkAddOwner.checked -eq "Checked")
    {
        write-host "Adding Owner to " $GrpName
        If ($Script:txtEmpNo.Text -notlike "*,*")
        {
            Set-DistributionGroup $GrpName -ManagedBy @{add=($Script:txtEmpNo.Text).Trim()} -BypassSecurityGroupManagerCheck
        }
        else
        {
            $Own = ($Script:txtEmpNo.Text -split ",").Trim()
            foreach ($Own in $Own)
            {
                $LineToWrite = "CHGOWN" + "`t" + "Adding " + $Own + " in: " + $GrpName
                WriteReportEvent
                Set-DistributionGroup $GrpName -ManagedBy @{add=$Own} -BypassSecurityGroupManagerCheck
            }
        }
        Start-Sleep -Seconds 10
    }
    If ($Script:chkRemoveOwner.checked -eq "Checked")
    {
        write-host "Removing Owner from " $GrpName
        $Own = ($Script:txtEmpNo.Text -split ",").Trim()
        foreach ($own in $own)
        {
            $LineToWrite = "CHGOWN" + "`t" + "Removing " + $Own + " in: " + $GrpName
            WriteReportEvent
            Set-DistributionGroup $GrpName -ManagedBy @{remove=$Own} -BypassSecurityGroupManagerCheck
        }
        Start-Sleep -Seconds 10
    }

    $group = (Get-DistributionGroup $GrpName).ManagedBy
    $ow = "Owner: "
    foreach ($g in $group)
    {
        $usr = (get-mailbox $g).displayName -split (", ")
        $ow = $ow + $usr[1] + " " + $usr[0] + ", "
    }
    $ow = $ow.TrimEnd(", ")

    Set-DistributionGroup $GrpName -MailTip ($ow + " - Per: " + $Global:txtInpTaskNo.Text)
    $LineToWrite = "UPGTIP" + "`t" + "Updating distribution group tip to: " + ($ow + " - Per: " + $Script:txtTicketNo.Text)
    WriteReportEvent

    Set-Group $GrpName -Notes ($ow + " - Per: " + $Script:txtTicketNo.Text)
    $LineToWrite = "UPNOTE" + "`t" + "Updating group note to: " + ($ow + " - Per: " + $Script:txtTicketNo.Text)
    WriteReportEvent

    If ($Script:TipUpdate -eq "N")
    {
        $mbxTip = (get-mailbox $Script:txtDispName.Text).MailTip
        If (($MbxTip -like "*Owner*") -or ($mbxTip.Length -eq 0))
        {
            write-host "Setting Shared Mailbox Tip"
            set-mailbox $Script:txtDispName.Text -MailTip $ow
            $LineToWrite = "UPMTIP" + "`t" + "Updating mailbox tip to:" + $ow
            WriteReportEvent
            $Script:TipUpdate = "Y"
        }
    }
}

Function Get-MbxOwner
{
    Foreach ($O in $Own)
    {
        $usr = (get-mailbox $o).displayName -split (", ")
        $ow = $ow + $usr[1] + " " + $usr[0] + ", "
    }
    $Script:MbxOwner = ""
    If ($ow.length -gt 0)
    {
        $Script:MbxOwner = $ow.TrimEnd(", ")
    }
}

$Year = (get-date).ToString("yyyy")
$Path = "e:\Automation\ShrMbxOwnerChange\Report\" + $Year
If (Test-Path $path) {} else {New-Item -Path $Path -ItemType Directory}
$Global:EDAccess = ""
$Global:AUAccess = ""
$Global:REAccess = ""
$Script:MbxOwner = ""
$Script:TipUpdate = "N"
$Valid = ""
$Global:OKDetails = "Continue"
$Global:InputFocus = $Script:txtDispName
Build-ChgOwnerInputForm
Publish-Form

If ($Global:Result -eq "OK")
{
    $ShrExists = [bool](Get-Mailbox $Script:txtDispName.Text -ErrorAction SilentlyContinue)
    If ($ShrExists -eq "True")
    {
        $ShrMbx = Get-Mailbox $Script:txtDispName.Text
        $MbxPerm = get-MailboxPermission $Script:txtDispName.Text |Where-Object {$_.User -like "MBX*"}
        $MbxFldrPerm = get-MailboxFolderPermission $Script:txtDispName.Text |Where-Object {$_.User -like "MBX*"}

        #Check for groups that are granted access at the top of the information store
        foreach ($MbxPerm in $MbxPerm)
        {
            If (($MbxPerm.User -like "*ED") -or ($MbxPerm.AccessRights -eq "FullAccess"))
            {
                $Global:EDAccess = $MbxPerm.User
                $Own = (Get-DistributionGroup $mbxPerm.User).ManagedBy
                Get-MbxOwner
            }
        }

        #Check for groups that are granted access at the folder level
        foreach ($MbxPerm in $MbxFldrPerm)
        {
            If (($MbxPerm.User.Displayname -like "*ED") -or ($MbxPerm.AccessRights -eq "Editor"))
            {
                $Global:EDAccess = $MbxPerm.User.DisplayName
                If ($Script:MbxOwner -eq "")
                {
                    $Own = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy
                    Get-MbxOwner
                }
            }
            If (($MbxPerm.User.Displayname -like "*AU") -or ($MbxPerm.AccessRights -eq "PublishingAuthor"))
            {
                $Global:AUAccess = $MbxPerm.User.DisplayName
                If ($Script:MbxOwner -eq "")
                {
                    $Own = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy
                    Get-MbxOwner
                }
            }
            If (($MbxPerm.User.Displayname -like "*RE") -or ($MbxPerm.AccessRights -eq "Reviewer"))
            {
                $Global:REAccess = $MbxPerm.User.DisplayName
                If ($Script:MbxOwner -eq "")
                {
                    $Own = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedB
                    Get-MbxOwner
                }
            }
        }

        $Global:OKDetails = "Change"
        Build-ShrMbxChgDetailsForm
        Publish-Form

        If ($Global:Result -eq "OK")
        {
            $ReportFile = $Path + "\Report-" + $Script:txtDispName.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            $LineToWrite = "STAR" + "`t" + "Change Shared Mailbox Owner script has started"
            WriteReportEvent
            $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	        WriteReportEvent
            $LineToWrite = "INPUT" + "`t" + "Input Form Details"
            WriteReportEvent
            $LineToWrite = "NAME" + "`t" + $Script:txtDispName.Text
            WriteReportEvent
            $LineToWrite = "TICK" + "`t" + $Script:txtTicketNo.Text
            WriteReportEvent
            $LineToWrite = "OWNER" + "`t" + $Script:MbxOwner
            WriteReportEvent
            $LineToWrite = "MBXTIP" + "`t" + $Script:txtMbxTip.Text
            WriteReportEvent
            $LineToWrite = "EDGRP" + "`t" + $Global:EDAccess
            WriteReportEvent
            $LineToWrite = "AUGRP" + "`t" + $Global:AUAccess
            WriteReportEvent
            $LineToWrite = "REGRP" + "`t" + $Global:REAccess
            WriteReportEvent
            $LineToWrite = "OWNCHG" + "`t" + $Script:txtEmpNo.Text
            WriteReportEvent
            $LineToWrite = "CHGADD" + "`t" + $Script:chkAddOwner.Checked
            WriteReportEvent
            $LineToWrite = "REMOWN" + "`t" + $Script:chkRemoveOwner.Checked
            WriteReportEvent
            $LineToWrite = "REPLOWN" + "`t" + $Script:chkReplaceOwner.Checked + "`n"
            WriteReportEvent

            If ($Script:chkAuthorized.checked -eq "Checked")
            {
                If (($Script:txtEmpNo.Text.Length -eq 5) -or ($Script:txtEmpNo.Text.Length -eq 6))
                {
                    $Script:txtEmpNo.Text = $Script:txtEmpNo.Text + "@global.ul.com"
                    $MbxExists = [bool](get-mailbox $Script:txtEmpNo.Text -ErrorAction SilentlyContinue)
                }
                else
                {
                    If ($Script:txtEmpNo.Text -like "*,*")
                    {
                        $Own = $Script:txtEmpNo.Text -split ","
                        foreach ($own in $own)
                        {
                            $MbxExists = [bool](get-mailbox $own -ErrorAction SilentlyContinue)
                            If ($MbxExists -eq "True")
                            {
                                If ($Valid -eq "")
                                {
                                    $Valid = $Own
                                }
                                else
                                {
                                    $Valid = $Valid + "," + $Own
                                }
                            }
                            else
                            {
                                $Output = $wshell.Popup("No mailbox for employee found for " + $Own + " individual will not be added.",0,"Employee Not Found",0+32)
                            }
                        }
                        If ($Valid -like "*,*")
                        {
                            $Own = $Valid -split ","
                        }
                        else
                        {
                            $Own = $Valid
                        }
                    }
                }
                    
                If ($Global:EDAccess -ne "")
                {
                    $GrpName = $Global:EDAccess
                    Chg-Owner
                }

                If ($Global:AUAccess -ne "")
                {
                    $GrpName = $Global:AUAccess
                    Chg-Owner
                }

                If ($Global:REAccess -ne "")
                {
                    $GrpName = $Global:REAccess
                    Chg-Owner
                }

                $LineToWrite = "EXIT" + "`t" + "Change Shared Mailbox Owner script has ended"
                WriteReportEvent
                Invoke-Expression -Command e:\O365AdminShared\EmailTemplates\SharedMailboxOwnershipChange.oft
            }
            else
            {
                $Output = $wshell.Popup("No changes made obtain authorization from an existing mailbox owner of " + $Script:txtDispName.Text + " mailbox.",0,"Not Found",0+32)
            }
        }
        else
        {
            $Output = $wshell.Popup("Shared mailbox ownership change request cancelled.",0,"Cancelled",0+32)
        }
    }
}
else
{
    $Output = $wshell.Popup("Shared mailbox ownership change request cancelled.",0,"Cancelled",0+32)
}
