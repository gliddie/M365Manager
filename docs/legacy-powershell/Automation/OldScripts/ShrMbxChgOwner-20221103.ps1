<#
#  This script changes the owner on a shared mailbox
#
#  08/27/2021 - Changed the Add/Remove/Replace owner to be a radio button rather than a checkbox
#  08/16/2022 - Modified to check for owners at the top of the information store
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
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Shared Mailbox Name:"
        $Global:lblDispName.Top = 20 ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtDispName = New-Object Windows.Forms.TextBox  
        $Global:txtDispName.TabIndex = 0 # set Tab Order 
        $Global:txtDispName.Top = 20; $Global:txtDispName.Left = 130; $Global:txtDispName.Width = 280;  
        $Global:txtDispName.Text = ""   # DisplayName
        $Global:txtDispName.TabIndex = 0
        $Global:form.Controls.Add($Global:txtDispName)    # Add to Form

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
    $Global:form.Size = New-Object System.Drawing.Size(440,360) #(W,H)

    ## Get Details to complete creation
    $TopLoc = 10
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Shared Mailbox Name:"
        $Global:lblDispName.Top = $TopLoc; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=150 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtDispName = New-Object Windows.Forms.TextBox  
        $Global:txtDispName.TabIndex = 0 # set Tab Order 
        $Global:txtDispName.Top = $TopLoc; $Global:txtDispName.Left = 130; $Global:txtDispName.Width = 220;  
        $Global:txtDispName.Text = (get-mailbox $Global:txtDispName.Text).DisplayName   # DisplayName
        $Global:form.Controls.Add($Global:txtDispName)    # Add to Form

    $TopLoc = $TopLoc + 30
    ## Ticket Number
    $Global:lblTaskNoDet = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNoDet.Text = "Ticket Number:"
        $Global:lblTaskNoDet.Top = $TopLoc; $Global:lblTaskNoDet.Left = 10; $Global:lblTaskNoDet.Width=150 ;$Global:lblTaskNoDet.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblTaskNoDet)    # Add to Form 
        # 
        $Global:txtTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtTaskNo.Top = $TopLoc; $Global:txtTaskNo.Left = 130; $Global:txtTaskNo.Width = 220;  
        $Global:txtTaskNo.Text = $Global:txtInpTaskNo.Text   # DisplayName
        $Global:form.Controls.Add($Global:txtTaskNo)    # Add to Form

    $TopLoc = $TopLoc + 30
    ## Mailbox Owner
    $Global:lblGrpOwnr = New-Object System.Windows.Forms.Label   
        $Global:lblGrpOwnr.Text = "Mailbox Owner(s):"  
        $Global:lblGrpOwnr.Top = $TopLoc; $Global:lblGrpOwnr.Left = 10; $Global:lblGrpOwnr.Width=150; $Global:lblGrpOwnr.AutoSize = $true 
        $Global:form.Controls.Add($lblGrpOwnr)    # Add to Form 
        # 
        $Global:txtGrpOwnr = New-Object Windows.Forms.TextBox  
        $Global:txtGrpOwnr.TabIndex = 0 # set Tab Order 
        $Global:txtGrpOwnr.Top = $TopLoc; $Global:txtGrpOwnr.Left = 130; $Global:txtGrpOwnr.Width = 220;
        If ($Script:MbxOwner -eq "")
        {
            $Script:MbxOwner = "(None)"
        } 
        $Global:txtGrpOwnr.Text = $Script:MbxOwner
        $Global:form.Controls.Add($Global:txtGrpOwnr)    # Add to Form 

    If ($Global:EDAccess -ne "")
    {
        $TopLoc = $TopLoc + 30
        ## Editor Old Group Name
        $Global:lblEDGrpName = New-Object System.Windows.Forms.Label
        $Global:lblEDGrpName.Text = ".ED Group:"  
        $Global:lblEDGrpName.Top = $TopLoc ; $Global:lblEDGrpName.Left = 10; $Global:lblEDGrpName.Width=150; $Global:lblEDGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblEDGrpName)    # Add to Form 
        #
        $Global:txtEDGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtEDGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtEDGrpName.Top = $TopLoc; $Global:txtEDGrpName.Left = 130; $Global:txtEDGrpName.Width = 220;  
        $Global:txtEDGrpName.Text = $Global:EDAccess
        $Global:form.Controls.Add($Global:txtEDGrpName)    # Add to Form 
    }

    If ($AUAccess -ne "")
    {
        $TopLoc = $TopLoc + 30
        ## Author Group Name
        $Global:lblAUGrpName = New-Object System.Windows.Forms.Label
        $Global:lblAUGrpName.Text = ".AU Group:" 
        $Global:lblAUGrpName.Top = $TopLoc; $Global:lblAUGrpName.Left = 10; $Global:lblAUGrpName.Width=150; $Global:lblAUGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAUGrpName)    # Add to Form 
        #
        $Global:txtAUGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtAUGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtAUGrpName.Top = $TopLoc; $Global:txtAUGrpName.Left = 130; $Global:txtAUGrpName.Width = 220;  
        $Global:txtAUGrpName.Text = $Global:AUAccess   # Use Corrent computer name as default 
        $Global:form.Controls.Add($Global:txtAUGrpName)    # Add to Form

    }

    If ($REAccess -ne "")
    {
        $TopLoc = $TopLoc + 30
        # Reader Group Name
        $Global:lblREGrpName = New-Object System.Windows.Forms.Label
        $Global:lblREGrpName.Text = ".RE Group:" 
        $Global:lblREGrpName.Top = $TopLoc ; $Global:lblREGrpName.Left = 10; $Global:lblREGrpName.Width=150; $Global:lblREGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblREGrpName)    # Add to Form 
        #
        $Global:txtREGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtREGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtREGrpName.Top = $TopLoc; $Global:txtREGrpName.Left = 130; $Global:txtREGrpName.Width = 220;  
        $Global:txtREGrpName.Text = $Global:REAccess
        $Global:form.Controls.Add($Global:txtREGrpName)    # Add to Form 
       # Obtain Value with: $txtGrpOwnr.Text
    }

    $TopLoc = $TopLoc + 30
    ## Employee # to Add/Remove
        $Global:lblEmpNo = New-Object System.Windows.Forms.Label
        $Global:lblEmpNo.Text = "Emp# or Address:" 
        $Global:lblEmpNo.Top = $TopLoc ; $Global:lblEmpNo.Left = 10; $Global:lblEmpNo.Width=150; $Global:lblEmpNo.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblEmpNo)    # Add to Form 
        #
        $Global:txtEmpNo = New-Object Windows.Forms.TextBox  
        $Global:txtEmpNo.TabIndex = 0 # set Tab Order 
        $Global:txtEmpNo.Top = $TopLoc; $Global:txtEmpNo.Left = 130; $Global:txtEmpNo.Width = 220;  
        $Global:txtEmpNo.Text = ""
        $Global:form.Controls.Add($Global:txtEmpNo)

    $TopLoc = $TopLoc + 30
    ## Add Owner
    $Global:chkAddOwner = New-Object Windows.Forms.RadioButton
        $Global:chkAddOwner.Left = 50; $Global:chkAddOwner.Width = 80; $Global:chkAddOwner.Top = $TopLoc
        $Global:chkAddOwner.Text = "Add Owner" 
        $Global:chkAddOwner.Checked = $false   # set a default value 
        $Global:chkAddOwner.TabIndex = 5
        $Global:form.Controls.Add($Global:chkAddOwner) 
        # Obtain Value with: $Global:chkRemoveOwner.Checked
               
    ## Remove Owner
    $Global:chkRemoveOwner = New-Object Windows.Forms.RadioButton
        $Global:chkRemoveOwner.Left = 150; $Global:chkRemoveOwner.Width = 110; $Global:chkRemoveOwner.Top = $TopLoc
        $Global:chkRemoveOwner.Text = "Remove Owner" 
        $Global:chkRemoveOwner.Checked = $false   # set a default value 
        $Global:chkRemoveOwner.TabIndex = 5
        $Global:form.Controls.Add($Global:chkRemoveOwner) 
        # Obtain Value with: $Global:chkRemoveOwner.Checked

    ## Replace Owner
    $Global:chkReplaceOwner = New-Object Windows.Forms.RadioButton
        $Global:chkReplaceOwner.Left = 270; $Global:chkReplaceOwner.Width = 120; $Global:chkReplaceOwner.Top = $TopLoc
        $Global:chkReplaceOwner.Text = "Replace Owners" 
        $Global:chkReplaceOwner.Checked = $false   # set a default value 
        $Global:chkReplaceOwner.TabIndex = 5
        $Global:form.Controls.Add($Global:chkReplaceOwner) 
        # Obtain Value with: $Global:chkReplaceOwner.Checked

    $TopLoc = $TopLoc + 30
    ## Authorized Requestor
    $Global:chkAuthorized = New-Object Windows.Forms.checkbox 
        $Global:chkAuthorized.Left = 130; $Global:chkAuthorized.Width = 200; $Global:chkAuthorized.Top = $TopLoc
        $Global:chkAuthorized.Text = "Authorized Requestor" 
        $Global:chkAuthorized.Checked = $false   # set a default value 
        $Global:chkAuthorized.TabIndex = 5
        $Global:form.Controls.Add($Global:chkAuthorized) 
        # Obtain Value with: $Global:chkAuthorized.Checked

        Add-FormStandardButtons
}

Function Chg-Owner
{
    If ($Global:chkReplaceOwner.checked -eq "Checked")
    {
        write-host "Replacing Owner  in " $GrpName
        If ($Global:txtEmpNo.Text -notlike "*,*")
        {
            Set-DistributionGroup $GrpName -ManagedBy $Global:txtEmpNo.Text -BypassSecurityGroupManagerCheck
        }
        else
        {
            $Own = $Global:txtEmpNo.Text -split ","
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
            }
        }
    }
    If ($Global:chkAddOwner.checked -eq "Checked")
    {
        write-host "Adding Owner to " $GrpName
        If ($Global:txtEmpNo.Text -notlike "*,*")
        {
            Set-DistributionGroup $GrpName -ManagedBy @{add=$Global:txtEmpNo.Text} -BypassSecurityGroupManagerCheck
        }
        else
        {
            $Own = $Global:txtEmpNo.Text -split ","
            foreach ($Own in $Own)
            {
                Set-DistributionGroup $GrpName -ManagedBy @{add=$Own} -BypassSecurityGroupManagerCheck
            }
        }
        Start-Sleep -Seconds 10
    }
    If ($Global:chkRemoveOwner.checked -eq "Checked")
    {
        write-host "Removing Owner from " $GrpName

        $Own = $Global:txtEmpNo.Text -split ","
        foreach ($own in $own)
        {
            Set-DistributionGroup $GrpName -ManagedBy @{remove=$Own} -BypassSecurityGroupManagerCheck
        }
        Start-Sleep -Seconds 10
    }

    Set-DistributionGroup $GrpName -MailTip ("Owner: " + ((get-DistributionGroup $GrpName).ManagedBy -join ", ") + " - Per " + $Global:txtTaskNo.Text)
    Set-Group $GrpName -Notes ("Owner: " + ((get-DistributionGroup $GrpName).ManagedBy -join ", ") + " - Per " + $Global:txtTaskNo.Text)
    If (((get-mailbox $Global:txtDispName.Text).MailTip -like "*Owner*") -or ((get-mailbox $Global:txtDispName.Text).MailTip.Length -eq 0))
    {
        write-host "Setting Shared Mailbox Tip"
        set-mailbox $Global:txtDispName.Text -MailTip ("Owner: " + ((get-DistributionGroup $GrpName).ManagedBy -join ", "))
    }
}

$Global:EDAccess = ""
$Global:AUAccess = ""
$Global:REAccess = ""
$Script:MbxOwner = ""
$Valid = ""
$Global:OKDetails = "Continue"
$Global:InputFocus = $Global:txtDispName
Build-ChgOwnerInputForm
Publish-Form

If ($Global:Result -eq "OK")
{
    $ShrExists = [bool](Get-Mailbox $Global:txtDispName.Text -ErrorAction SilentlyContinue)
    If ($ShrExists -eq "True")
    {
        $MbxPerm = get-MailboxPermission $Global:txtDispName.Text |Where-Object {$_.User -like "MBX*"}
        $MbxFldrPerm = get-MailboxFolderPermission $Global:txtDispName.Text |Where-Object {$_.User -like "MBX*"}

        #Check for groups that are granted access at the top of the information store
        foreach ($MbxPerm in $MbxPerm)
        {
            If (($MbxPerm.User -like "*ED") -or ($MbxPerm.AccessRights -eq "FullAccess"))
            {
                $Global:EDAccess = $MbxPerm.User
                $Script:MbxOwner = (Get-DistributionGroup $mbxPerm.User).ManagedBy -join ", "
            }
        }

        #Check for groups that are granted access at the folder level
        foreach ($MbxPerm in $MbxFldrPerm)
        {
            If (($MbxPerm.User.Displayname -like "*ED") -or ($MbxPerm.AccessRights -eq "Editor"))
            {
                $Global:EDAccess = $MbxPerm.User.DisplayName
                $Script:MbxOwner = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
            }
            If (($MbxPerm.User.Displayname -like "*AU") -or ($MbxPerm.AccessRights -eq "PublishingAuthor"))
            {
                $Global:AUAccess = $MbxPerm.User.DisplayName
                If ($Script:MbxOwner -eq "")
                {
                    $Script:MbxOwner = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
                }
            }
            If (($MbxPerm.User.Displayname -like "*RE") -or ($MbxPerm.AccessRights -eq "Reviewer"))
            {
                $Global:REAccess = $MbxPerm.User.DisplayName
                If ($Script:MbxOwner -eq "")
                {
                    $Script:MbxOwner = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
                }
            }
        }

        $Global:OKDetails = "Change"
        Build-ShrMbxChgDetailsForm
        Publish-Form

        If ($Global:Result -eq "OK")
        {
            If ($Global:chkAuthorized.checked -eq "Checked")
            {
                If ($Global:txtEmpNo.Text.Length -eq 5)
                {
                    $Global:txtEmpNo.Text = $Global:txtEmpNo.Text + "@global.ul.com"
                    $MbxExists = [bool](get-mailbox $Global:txtEmpNo.Text -ErrorAction SilentlyContinue)
                }
                else
                {
                    If ($Global:txtEmpNo.Text -like "*,*")
                    {
                        $Own = $Global:txtEmpNo.Text -split ","
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
                Invoke-Expression -Command e:\O365AdminShared\EmailTemplates\SharedMailboxOwnershipChange.oft
            }
            else
            {
                $Output = $wshell.Popup("No changes made obtain authorization from an existing mailbox owner of " + $Global:txtDispName.Text + " mailbox.",0,"Not Found",0+32)
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
