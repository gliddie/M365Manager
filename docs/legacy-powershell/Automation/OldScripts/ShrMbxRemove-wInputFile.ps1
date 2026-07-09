<#   
#================================================================================ 
# Name: Remove Shared Mailbox
#
# 10/12/2020 - SAG - Redesigned using forms and the old RemoveSharedMailbox.ps1 script
# 05/17/2022 - SAG - Modified the report file to be written out to the E: drive
#================================================================================ 
#>  

function Build-RemoveMbxInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Remove Shared Mailbox" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(450,250) #(W,H)
    $BldDetails = "N"

    Add-FormStandardButtons

    ## Get Details used to create
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Shared Mailbox Name:"
        $Global:lblDispName.Top = 20 ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtInpName = New-Object Windows.Forms.TextBox  
        $Global:txtInpName.TabIndex = 0 # set Tab Order 
        $Global:txtInpName.Top = 20; $Global:txtInpName.Left = 130; $Global:txtInpName.Width = 280;  
        $Global:txtInpName.Text = ""   # DisplayName
        $Global:txtInpName.TabIndex = 0
        $Global:form.Controls.Add($Global:txtInpName)    # Add to Form

    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 50 ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtInpTaskNo.Top = 50; $Global:txtInpTaskNo.Left = 130; $Global:txtInpTaskNo.Width = 120;  
        $Global:txtInpTaskNo.Text = "CHG0098108"   # Enter ticket number
        $Global:txtInpTaskNo.TabIndex = 1
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form

    ## NoOwners
    $Global:chkNoOwner = New-Object Windows.Forms.checkbox 
        $Global:chkNoOwner.Left = 100; $Global:chkNoOwner.Width = 250; $Global:chkNoOwner.Top = 90
        $Global:chkNoOwner.Text = "No Response to Request for Owners" 
        $Global:chkNoOwner.Checked = $false   # set a default value 
        $Global:chkNoOwner.TabIndex = 5
        $Global:form.Controls.Add($Global:chkNoOwner) 
        # Obtain Value with: $Global:chkNoOwner.Checked

    ## NoOwnersResponse
    $Global:chkNoOwnerResp = New-Object Windows.Forms.checkbox 
        $Global:chkNoOwnerResp.Left = 100; $Global:chkNoOwnerResp.Width = 250; $Global:chkNoOwnerResp.Top = 120
        $Global:chkNoOwnerResp.Text = "Response to Request for Owners" 
        $Global:chkNoOwnerResp.Checked = $false   # set a default value 
        $Global:chkNoOwnerResp.TabIndex = 5
        $Global:form.Controls.Add($Global:chkNoOwnerResp) 
        # Obtain Value with: $Global:chkNoOwnerResp.Checked
}

function Build-ShrMbxRemDetailsForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Remove Shared Mailbox" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Width = 450 ; $form.Height = 180
#    $Global:form.Size = New-Object System.Drawing.Size(660,550) #(W,H)
    $Global:form.AutoSize = $True

    Add-FormStandardButtons

    ## Get Details to complete creation
    $TopLoc = 10
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Shared Mailbox Name:"
        $Global:lblDispName.Top = $TopLoc; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=180 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtDispName = New-Object Windows.Forms.TextBox  
        $Global:txtDispName.TabIndex = 0 # set Tab Order 
        $Global:txtDispName.Top = $TopLoc; $Global:txtDispName.Left = 130; $Global:txtDispName.Width = 220;  
        $Global:txtDispName.Text = (get-mailbox $Global:txtInpName.Text).DisplayName   # DisplayName
        $Global:form.Controls.Add($Global:txtDispName)    # Add to Form

    $TopLoc = $TopLoc + 30
    If (($Global:chkNoOwner.Checked -eq "Checked") -or ($Global:chkNoOwnerResp.Checked -eq "Checked"))
    {
        ## Date of Original Communication
        $Global:lblCommDate = New-Object System.Windows.Forms.Label   
            $Global:lblCommDate.Text = "Email Date:"
            $Global:lblCommDate.Top = $TopLoc; $Global:lblCommDate.Left = 10; $Global:lblCommDate.Width=180 ;$Global:lblCommDate.AutoSize = $true 
            $Global:form.Controls.Add($Global:lblCommDate)    # Add to Form 
            # 
            $Global:txtCommDate = New-Object Windows.Forms.TextBox  
            $Global:txtCommDate.TabIndex = 0 # set Tab Order 
            $Global:txtCommDate.Top = $TopLoc; $Global:txtCommDate.Left = 130; $Global:txtCommDate.Width = 220;  
            $Global:txtCommDate.Text = ""
            $Global:form.Controls.Add($Global:txtCommDate)    # Add to Form
    }
    else
    {
        ## Communication Date
        $Global:lblCommDate = New-Object System.Windows.Forms.Label   
            $Global:lblCommDate.Text = "Ticket Number:"
            $Global:lblCommDate.Top = $TopLoc; $Global:lblCommDate.Left = 10; $Global:lblCommDate.Width=180 ;$Global:lblCommDate.AutoSize = $true 
            $Global:form.Controls.Add($Global:lblCommDate)    # Add to Form 
            # 
            $Global:txtTaskNo = New-Object Windows.Forms.TextBox  
            $Global:txtTaskNo.TabIndex = 0 # set Tab Order 
            $Global:txtTaskNo.Top = $TopLoc; $Global:txtTaskNo.Left = 130; $Global:txtTaskNo.Width = 220;  
            $Global:txtTaskNo.Text = $Global:txtInpTaskNo.Text   # DisplayName
            $Global:form.Controls.Add($Global:txtTaskNo)    # Add to Form
    }

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
        $Global:txtGrpOwnr.Text = $MbxOwner
        $Global:form.Controls.Add($Global:txtGrpOwnr)    # Add to Form 

    If ($Script:EDAccess -ne "")
    {
        $form.Height = $form.Height + 110
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
        $Global:txtEDGrpName.Text = $Script:EDAccess
        $Global:form.Controls.Add($Global:txtEDGrpName)    # Add to Form 

    ## DeleteEDGroupCheck Box
    $Global:chkRemEDGrp = New-Object Windows.Forms.checkbox 
        $Global:chkRemEDGrp.Left = 360; $Global:chkRemEDGrp.Width = 150; $Global:chkRemEDGrp.Top = $TopLoc
        $Global:chkRemEDGrp.Text = "Delete ED Group" 
        $Global:chkRemEDGrp.Checked = $false   # set a default value 
        $Global:chkRemEDGrp.TabIndex = 5
        $Global:chkRemEDGrp.Checked = $True
        $Global:form.Controls.Add($Global:chkRemEDGrp) 
        # Obtain Value with: $Global:chkRemEDGrp.Checked

        $TopLoc = $TopLoc + 30
        ## Editor Group Members
        $Global:lblEDGrpMbr = New-Object System.Windows.Forms.Label   
        $Global:lblEDGrpMbr.Text = ".ED Member(s):"  
        $Global:lblEDGrpMbr.Top = $TopLoc; $Global:lblEDGrpMbr.Left = 10; $Global:lblEDGrpMbr.Width=150; $Global:lblEDGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblEDGrpMbr)    # Add to Form 
        # 
        $Global:txtEDGrpMbr = New-Object Windows.Forms.TextBox
        $Global:txtEDGrpMbr.TabIndex = 0 # set Tab Order
        $Global:txtEDGrpMbr.Location = New-Object System.Drawing.Size(130,$TopLoc)
        $Global:txtEDGrpMbr.Size = New-Object system.Drawing.Size(520,60)
        $Global:txtEDGrpMbr.MultiLine = $true
        $Global:txtEDGrpMbr.ScrollBars = 'Both'
        $Global:txtEDGrpMbr.Text = $Script:EDMembers
        $Global:form.Controls.Add($Global:txtEDGrpMbr)    # Add to Form
        $TopLoc = $TopLoc + 40
    }

    If ($AUAccess -ne "")
    {
        $TopLoc = $TopLoc + 30
        $form.Height = $form.Height + 110
        # Author Group Name
        $Global:lblAUGrpName = New-Object System.Windows.Forms.Label
        $Global:lblAUGrpName.Text = ".AU Group:" 
        $Global:lblAUGrpName.Top = $TopLoc; $Global:lblAUGrpName.Left = 10; $Global:lblAUGrpName.Width=150; $Global:lblAUGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAUGrpName)    # Add to Form 
        #
        $Global:txtAUGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtAUGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtAUGrpName.Top = $TopLoc; $Global:txtAUGrpName.Left = 130; $Global:txtAUGrpName.Width = 220;  
        $Global:txtAUGrpName.Text = $Script:AUAccess   # Use Corrent computer name as default 
        $Global:form.Controls.Add($Global:txtAUGrpName)    # Add to Form

        ## DeleteAUGroupCheck Box
        $Global:chkRemAUGrp = New-Object Windows.Forms.checkbox 
        $Global:chkRemAUGrp.Left = 360; $Global:chkRemAUGrp.Width = 150; $Global:chkRemAUGrp.Top = $TopLoc
        $Global:chkRemAUGrp.Text = "Delete AU Group" 
        $Global:chkRemAUGrp.Checked = $false   # set a default value 
        $Global:chkRemAUGrp.TabIndex = 5
        $Global:chkRemAUGrp.Checked = $True
        $Global:form.Controls.Add($Global:chkRemAUGrp) 

        $TopLoc = $TopLoc + 30
        ## Author Group Members
        $Global:lblAUGrpMbr = New-Object System.Windows.Forms.Label   
        $Global:lblAUGrpMbr.Text = ".AU Member(s):"  
        $Global:lblAUGrpMbr.Top = $TopLoc; $Global:lblAUGrpMbr.Left = 10; $Global:lblAUGrpMbr.Width=150; $Global:lblAUGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAUGrpMbr)    # Add to Form 
        # 
        $Global:txtAUGrpMbr = New-Object Windows.Forms.TextBox
        $Global:txtAUGrpMbr.TabIndex = 0 # set Tab Order 
        $Global:txtAUGrpMbr.Text = $Script:AUMembers
        $Global:txtAUGrpMbr.Location = New-Object System.Drawing.Size(130,$TopLoc)
        $Global:txtAUGrpMbr.Size = New-Object system.Drawing.Size(520,60)
        $Global:txtAUGrpMbr.MultiLine = $true
        $Global:txtAUGrpMbr.ScrollBars = 'Both'
        $Global:txtAUGrpMbr.Text = $Script:AUMembers
        $Global:form.Controls.Add($Global:txtAUGrpMbr)    # Add to Form 
        $TopLoc = $TopLoc + 40
    }

    If ($REAccess -ne "")
    {
        $TopLoc = $TopLoc + 30
        $form.Height = $form.Height + 110
        # Reader Group Name
        $Global:lblREGrpName = New-Object System.Windows.Forms.Label
        $Global:lblREGrpName.Text = ".RE Group:" 
        $Global:lblREGrpName.Top = $TopLoc ; $Global:lblREGrpName.Left = 10; $Global:lblREGrpName.Width=150; $Global:lblREGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblREGrpName)    # Add to Form 
        #
        $Global:txtREGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtREGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtREGrpName.Top = $TopLoc; $Global:txtREGrpName.Left = 130; $Global:txtREGrpName.Width = 220;  
        $Global:txtREGrpName.Text = $Script:REAccess
        $Global:form.Controls.Add($Global:txtREGrpName)    # Add to Form 

        ## DeleteREGroupCheck Box
        $Global:chkRemREGrp = New-Object Windows.Forms.checkbox 
        $Global:chkRemREGrp.Left = 360; $Global:chkRemREGrp.Width = 150; $Global:chkRemREGrp.Top = $TopLoc
        $Global:chkRemREGrp.Text = "Delete RE Group" 
        $Global:chkRemREGrp.Checked = $false   # set a default value 
        $Global:chkRemREGrp.TabIndex = 5
        $Global:chkRemREGrp.Checked = $True
        $Global:form.Controls.Add($Global:chkRemREGrp) 

        $TopLoc = $TopLoc + 30
        ## Reader Group Members
        $Global:lblREGrpMbr = New-Object System.Windows.Forms.Label   
        $Global:lblREGrpMbr.Text = ".RE Member(s):"  
        $Global:lblREGrpMbr.Top = $TopLoc; $Global:lblREGrpMbr.Left = 10; $Global:lblREGrpMbr.Width=150; $Global:lblREGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblREGrpMbr)    # Add to Form 
        # 
        $Global:txtREGrpMbr = New-Object Windows.Forms.TextBox
        $Global:txtREGrpMbr.TabIndex = 0 # set Tab Order 
        $Global:txtREGrpMbr.Text = $Script:REMembers
        $Global:txtREGrpMbr.Location = New-Object System.Drawing.Size(130,$TopLoc)
        $Global:txtREGrpMbr.Size = New-Object system.Drawing.Size(520,60)
        $Global:txtREGrpMbr.MultiLine = $true
        $Global:txtREGrpMbr.ScrollBars = 'Both'
        $Global:txtREGrpMbr.Text = $Script:REMembers
        $Global:form.Controls.Add($Global:txtREGrpMbr)    # Add to Form
        $TopLoc = $TopLoc + 40
    }

    $TopLoc = $TopLoc + 30
    ## Not for Display Box to Not Display the email template
    $form.Height = $form.Height + 20
    $Global:chkTemplate = New-Object Windows.Forms.checkbox 
        $Global:chkTemplate.Left = 130; $Global:chkTemplate.Width = 300; $Global:chkTemplate.Top = $TopLoc
        $Global:chkTemplate.Text = "Display Email Template for Shared Mailbox Removal" 
        $Global:chkTemplate.Checked = $false   # set a default value
        $Global:form.Controls.Add($Global:chkTemplate)
}

function Remove-Group
{
    If ($InputDL -notlike "*,*")
    {
        $InputDL = $InputDL.Trim()
        $Exists = [bool](Get-DistributionGroup $InputDL -ErrorAction SilentlyContinue)
        If ($Exists -eq $True)
        {
            $InputDL = $InputDL.Trim()
            write-host "Removing Access Group: " $InputDL
            Write-Output "Security Access Group Details - " $InputDL >> $OutFileName
            Get-DistributionGroup $InputDL >> $OutFileName
	        Write-Output "Security Access Group Full Details" >> $OutFileName
            Get-DistributionGroup $InputDL |fl >> $OutFileName
	        Write-Output "Security Access Group Manager and Notes" >> $OutFileName
            Get-Group $InputDL |fl ManagedBy,Notes >> $OutFileName
	        write-Output "Security Access Group Membership" >> $OutFileName
            Get-DistributionGroupMember $InputDL |ft Alias,Name,RecipientType >> $OutFileName
            Remove-DistributionGroup $InputDL -BypassSecurityGroupManagerCheck -confirm:$False >> $OutFileName
        }
    }
}

$mbx = Import-Csv "c:\temp\RemoveShrMbx.csv"
write-host "Number of Mailboxes to Process: " $mbx.count
Foreach ($mbx in $mbx)
{
    $Script:MbxOwner = ""
    $Script:EDAccess = ""
    $Script:AUAccess = ""
    $Script:REAccess = ""
    $Global:chkRemEDGrp = ""
    $Global:chkRemAUGrp = ""
    $Global:chkRemREGrp = ""
    $Global:chkRemOtherGrp = ""
    $GrpName = ""
    $Valid = ""
    Build-RemoveMbxInputForm
    $Global:txtInpName.Text = $mbx.Name
    $Global:Result = "OK"

    If ($Global:Result -eq "OK")
    {
        $ShrExists = [bool](Get-Mailbox $Global:txtInpName.Text -ErrorAction SilentlyContinue)
        If ($ShrExists -eq "True")
        {
            $MbxPerm = get-mailboxPermission $Global:txtInpName.Text |Where-Object {$_.User -like "*MBX*"}
            foreach ($MbxPerm in $MbxPerm)
            {
                If ([bool](Get-DistributionGroup $MbxPerm.User -ErrorAction SilentlyContinue))
                {
                    If ($MbxPerm.User -like "*.ED")
                    {
                        If ($Script:EDAccess -eq "")
                        {
                            $Script:EDAccess = $MbxPerm.User
                            $Script:EDMembers = (Get-DistributionGroupMember $MbxPerm.User) -join ","
                            $Script:MbxOwner = (Get-DistributionGroup $mbxPerm.User).ManagedBy -join ", "
                        }
                        else
                        {
                            $Script:EDAccess = $Script:EDAccess + "," + $MbxPerm.User
                            $Script:EDMembers = $Script:EDMembers + "," + (Get-DistributionGroupMember $MbxPerm.User) -join ","
                            If ($Script:MbxOwner -ne "")
                            {
                                $Script:MbxOwner = $Script:MbxOwner + "," + (Get-DistributionGroup $mbxPerm.User).ManagedBy -join ", "
                            }
                        }
                    }

                    If ($MbxPerm.User -like "*.AU")
                    {
                        If ($Script:AUAccess -eq "")
                        {
                            $Script:AUAccess = $MbxPerm.User
                            $Script:AUMembers = (Get-DistributionGroupMember $MbxPerm.User) -join ","
                            $Script:MbxOwner = (Get-DistributionGroup $mbxPerm.User).ManagedBy -join ", "
                        }
                        else
                        {
                            $Script:AUAccess = $Script:AUAccess + "," + $MbxPerm.User
                            $Script:AUMembers = $Script:AUMembers + "," + (Get-DistributionGroupMember $MbxPerm.User) -join ","
                            If ($Script:MbxOwner -ne "")
                            {
                                $Script:MbxOwner = $Script:MbxOwner + "," + (Get-DistributionGroup $mbxPerm.User).ManagedBy -join ", "
                            }
                        }
                    }

                    If ($MbxPerm.User -like "*.RE")
                    {
                        If ($Script:REAccess -eq "")
                        {
                            $Script:REAccess = $MbxPerm.User
                            $Script:REMembers = (Get-DistributionGroupMember $MbxPerm.User) -join ","
                            $Script:MbxOwner = (Get-DistributionGroup $mbxPerm.User).ManagedBy -join ", "
                        }
                        else
                        {
                            $Script:REAccess = $Script:REAccess + "," + $MbxPerm.User
                            $Script:REMembers = $Script:REMembers + "," + (Get-DistributionGroupMember $MbxPerm.User) -join ","
                            If ($Script:MbxOwner -ne "")
                            {
                                $Script:MbxOwner =  $Script:MbxOwner+ "," + (Get-DistributionGroup $mbxPerm.User).ManagedBy -join ", "
                            }
                        }
                    }
                }
            }

            $MbxPerm = get-mailboxFolderPermission $Global:txtInpName.Text |Where-Object {$_.AccessRights -notlike "*None*"}
            foreach ($MbxPerm in $MbxPerm)
            {
                If ([bool](get-distributiongroup $mbxperm.user.DisplayName -ErrorAction SilentlyContinue))
                {
                    $GrpName = $MbxPerm.User.DisplayName
                    If ($MbxPerm.User.DisplayName -notlike "*,*")
                    {
                        $GrpName = (get-distributiongroup $MbxPerm.User.DisplayName).Name
                    }
                
                    If ((($GrpName -like "*.ED") -and ($GrpName -notlike "")) -or (($MbxPerm.User.DisplayName -notlike "*,*") -and ($MbxPerm.AccessRights -eq "Editor")))
                    {
                        If ($Script:EDAccess -eq "")
                        {
                            $Script:EDAccess = $GrpName
                            $Script:EDMembers = (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ","
                            $Script:MbxOwner = (Get-DistributionGroup $MbxPerm.User.DisplayName).ManagedBy -join ", "
                        }
                        else
                        {

                            If ($Script:EDAccess -notlike ("*"+$GrpName+"*"))
                            {
                                $Script:EDAccess = $Script:EDAccess + "," + $GrpName
                                $Script:EDMembers = $Script:EDMembers + "," + (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ","
                                $Script:MbxOwner = $Script:MbxOwner + "," + (Get-DistributionGroup $MbxPerm.User.DisplayName).ManagedBy -join ", "
                            }
                        }
                    }
                    If ((($GrpName -like "*.AU") -and ($GrpName -notlike "")) -or (($MbxPerm.User.DisplayName -notlike "*,*") -and ($MbxPerm.AccessRights -eq "PublishingAuthor")))
                    {
                        $Script:AUAccess = $GrpName
                        $Script:AUMembers = (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ","
                        If ($Script:MbxOwner -eq "")
                        {
                            $Script:MbxOwner = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
                        }
                        else
                        {
                            If ($Script:AUAccess -notlike ("*"+$GrpName+"*"))
                            {
                                $Script:AUAccess = $Script:AUAccess + "," + $GrpName
                                $Script:AUMembers = $Script:AUMembers + "," + (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ","
                                If ($Script:MbxOwner -eq "")
                                {
                                    $Script:MbxOwner = (Get-DistributionGroup $MbxPerm.User.DisplayName).ManagedBy -join ", "
                                }
                            }
                        }
                    }
                    If ((($GrpName -like "*.RE") -and ($GrpName -notlike "")) -or (($MbxPerm.User.DisplayName -notlike "*,*") -and ($MbxPerm.AccessRights -eq "Reviewer")))
                    {
                        $Script:REAccess = $GrpName
                        $Script:REMembers = (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ","
                        If ($Script:MbxOwner -eq "")
                        {
                            $Script:MbxOwner = (Get-DistributionGroup $MbxPerm.User.DisplayName).ManagedBy -join ", "
                        }
                        else
                        {
                            If ($Script:REAccess -notlike ("*"+$GrpName+"*"))
                            {
                                $Script:REAccess = $Script:REAccess + "," + $GrpName
                                $Script:REMembers = $Script:REMembers + "," + (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ","
                                If ($Script:MbxOwner -eq "")
                                {
                                    $Script:MbxOwner = (Get-DistributionGroup $MbxPerm.User.DisplayName).ManagedBy -join ", "
                                }
                            }
                        }
                    }
                }
            }

            $Global:OKDetails = "Remove"
            Build-ShrMbxRemDetailsForm
            Publish-Form

            If ($Global:Result -eq "OK")
            {
                $Global:Reason = $Global:txtTaskNo.Text
        
                If ($Global:chkNoOwner.Checked -eq "Checked")
                {
                    $Global:Reason = "No response to No Owners email request sent on: " + $Global:txtCommDate.Text
                }
            
                If ($Global:chkNoOwnerResp.Checked -eq "Checked")
                {
                    $Global:Reason = "Mailbox user response to No Owners email request received on: " + $Global:txtCommDate.Text
                }

                write-host "`nRemoving Shared Mailbox: " $Global:txtDispName.Text -ForegroundColor Green
                $Date = get-date -Format "yyyy-MMdd"
                $OutFileName = "e:\Automation\RemoveShrMbx\Report\RemoveShrMbx-" + ($Global:txtInpName.Text -Replace("['.()/\- ]","")) + "-Date" + $Date + ".log"

                Write-Output "Removing Shared Mailbox" > $OutFileName
                Write-Output "Ticket Number:  $Global:Reason" >> $OutFileName
                write-Output "Launched by:  $WhoAmI" >> $OutFileName
                Get-Mailbox $Global:txtInpName.Text >> $OutFileName
                Get-Mailbox $Global:txtInpName.Text |fl >> $OutFileName
                Write-Output "Mailbox Rules" >> $OutFileName
                Get-InboxRule -Mailbox $Global:txtInpName.Text |Where-Object{($_.Description -like "*forward*") -or ($_.Description -like "*redirect*")} |fl Name,Priority,*desc* >> $OutFileName
                Write-Output "Mailbox Permissions" >> $OutFileName
                Get-MailboxPermission $Global:txtInpName.Text | where {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous") -and ($_.IsInherited -notlike "True")} | ft User,AccessRights >> $OutFileName
                Write-Output "MailboxFolder Permissions" >> $OutFileName
                Get-MailboxFolderPermission $Global:txtInpName.Text >> $OutFileName
                Write-Output "Recipient Permissions" >> $OutFileName
                Get-RecipientPermission $Global:txtInpName.Text >> $OutFileName
                Write-Output "Mailbox Message Statistics - Number of Messages in Mailbox" >> $OutFileName
                Get-MailboxStatistics $Global:txtInpName.Text |ft >> $OutFileName
                Write-Output "Mailbox Message Folder Statistics - Number of Messages in Mailbox" >> $OutFileName
                Get-MailboxFolderStatistics $Global:txtInpName.Text |ft Name,ItemsInFolder >> $OutFileName

                If ($Global:chkRemEDGrp.Checked -eq "Checked")
                {
                    $InputDL = $Global:txtEDGrpName.Text -split ","
                    Foreach ($InputDL in $InputDL)
                    {
                        Remove-Group
                    }
                    $InputDL = $Script:EDAccess
                    Remove-Group
                }

                If ($Global:chkRemAUGrp.Checked -eq "Checked")
                {
                    $InputDL = $Global:txtAUGrpName.Text -split ","
                    Foreach ($InputDL in $InputDL)
                    {
                        Remove-Group
                    }            }

                If ($Global:chkRemREGrp.Checked -eq "Checked")
                {
                    $InputDL = $Global:txtREGrpName.Text -split ","
                    Foreach ($InputDL in $InputDL)
                    {
                        Remove-Group
                    }
                }

                Remove-Mailbox $Global:txtInpName.Text -confirm:$False >> $OutFileName

                If (($Global:chkNoOwner.Checked -ne "Checked") -and ($Global:chkNoOwnerResp.Checked -ne "Checked"))
                {
                    If ($Global:chkTemplate.Checked -eq "Checked")
                    {
                        invoke-Expression -Command e:\O365AdminShared\EMailTemplates\SharedMailboxRemoval.oft
                    }
                }
                write-host "Mailbox Removal Report can be found at: " $OutFileName
            }
            else
            {
                $Output = $wshell.Popup("Shared mailbox removal request cancelled.",0,"Cancelled",0+32)
            }
        }
        else
        {
            Write-Host "Shared mailbox" $mbx.Name "does not exist." -ForegroundColor Red
        }
    }
    else
    {
        $Output = $wshell.Popup("Shared mailbox removal request cancelled.",0,"Cancelled",0+32)
    }
}
