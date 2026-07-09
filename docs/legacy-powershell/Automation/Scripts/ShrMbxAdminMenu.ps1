<####
#### Shared Mailbox Admin Menu
#
#  Called by:  O365AdminMenu.ps1
#
#  04/15/2020 - SAG - Combined the AddMailboxPermissionSharedMailbox and ApplyRetentionPolicy scripts into the AddFolderPermission script for the Option 1 and 3
#  05/20/2022 - SAG - Added Shared Mailbox Removal with Input File
#  11/04/2022 - SAG - Added New Script to add access groups for an existing mailbox
#
####>

function Add-FolderPermissions
{
write-host "Getting Folder Permissions" -ForegroundColor Red
    if ($GrpAddr.contains("@"))
	{
		if ((Get-Mailbox $MbxAddr).MailTip -eq $null)
        {
            $NewTip = ("Owners: " + (((Get-DistributionGroup $GrpAddr).Managedby) -join ", "))
            if ($NewTip.Length -gt 175)
            {
                write-host "Maximum mail tip length exceeded truncating to 175 characters." -foregroundcolor Red
                $NewTip = $NewTip.Substring(0,175)
                $OldTip = (Get-Mailbox $MbxAddr).MailTip
                write-host "Replacing old MailTip: " $OldTip.Substring(16,$OldTip.Length-36) -foregroundcolor Cyan
                write-host "                 With: " $NewTip
		        $LineToWrite = "REPL  " + "`t" + "Replacing old MailTip with " + $NewTip
		        WriteReportEvent
            }
            Set-Mailbox $MbxAddr -MailTip $NewTip
        }

		# 3/23/2020 is the date when all new mailboxes started getting the 3 yr retention policy
        if (((Get-Mailbox $MbxAddr).WhenCreated -gt (Get-Date -Year 2020 -Month 3 -Day 23)) -and ((Get-Mailbox $MbxAddr).RetentionPolicy -notlike "*3 yr*"))
        {
            Set-MailBox $MbxAddr -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL MRM Policy - 3 yr Delete" -RetentionHoldEnabled $false
    	    Set-Mailbox $MbxAddr -UseDatabaseQuotaDefaults $False -IssueWarningQuota 45GB -ProhibitSendQuota 49.75GB -ProhibitSendReceiveQuota 50GB
            write-host "3 yr Retention Policy Applied to" $MbxAddr
		    $LineToWrite = "UPD  " + "`t" + "Retention Policy Set " + $MbxAddr
	        WriteReportEvent
		# Check that IMAP and POP3 are disabled
    	    $ProtocolCheck = (get-CASMailbox $MbxAddr) 
    		if ($ProtocolCheck.ImapEnabled -eq $true)
	    	{
		        Set-CASMailBox $MbxAddr -ImapEnabled $false
    		    $LineToWrite = "INFO" + "`t" + $Global:txtDispName.Text + "`t" + "IMAP Protocol Disabled"
	    	    WriteReportEvent
		    }
    		if ($ProtocolCheck.PopEnabled -eq $true)
	    	{
		        Set-CASMailBox $MbxAddr -PopEnabled $false
    		    $LineToWrite = "INFO" + "`t" + $Global:txtDispName.Text + "`t" + "POP3 Protocol Disabled"
	    	    WriteReportEvent
		    }
        }

        $Permission = ""
        if ($GrpAddr.contains(".ED@"))
        {
            $Permission = "Editor"
       		write-host "Setting permissions on: " $MbxAddr " <-- " $GrpAddr " (Editor)" -ForegroundColor Cyan
	    	$LineToWrite = "INFO" + "`t" + "Setting Editor Permissions for " + $MbxAddr +  " for " + $GrpAddr
		   	WriteReportEvent
            Write-Host "GrantSendOnBehalf to Delegate Group:  " $GrpAddr -ForegroundColor Cyan
            Set-Mailbox -Identity $MbxAddr -GrantSendOnBehalfTo ((Get-Mailbox -Identity $MbxAddr).GrantSendOnBehalfTo += $GrpAddr)
            Add-MailboxPermission $MbxAddr –User $GrpAddr –AccessRights FullAccess
            Add-RecipientPermission $MbxAddr -Trustee $GrpAddr –AccessRights SendAs -Confirm:$false
        }
        elseif ($GrpAddr.contains(".AU@"))
        {
            $Permission = "PublishingAuthor"
            write-host "Setting permissions on: " $MbxAddrr " <-- " $GrpAddr " (PublishingAuthor)" -ForegroundColor Cyan
	    	$LineToWrite = "INFO" + "`t" + "Setting Author Permissions for " + $MbxAddr +  " for " + $GrpAddr
		   	WriteReportEvent
            Write-Host "GrantSendOnBehalf to Delegate Group:  " $GrpAddr -ForegroundColor Cyan
            Set-Mailbox -Identity $MbxAddr -GrantSendOnBehalfTo ((Get-Mailbox -Identity $MbxAddr).GrantSendOnBehalfTo += $GrpAddr)
        }
        elseif ($GrpAddr.contains(".RE@"))
        {
            $Permission = "Reviewer"
            write-host "Setting permissions on: " $MbxAddr " <-- " $GrpAddr " (Reviewer)" -ForegroundColor Cyan
	    	$LineToWrite = "INFO" + "`t" + "Setting Read Permissions on: " + $MbxAddr + " for " + $GrpAddr
		   	WriteReportEvent
        }
        else
        {
            write-host "Group name does not follow naming stand there is no .ED, .AU or .RE found." -ForegroundColor Red
            write-host "Group Name: " $GrpAddr
            $LineToWrite = "INFO" + "`t" + "Group name does not follow naming stand there is no .ED, .AU or .RE found " + $MbxAddr
		   	WriteReportEvent
        }
        
        If ($Permission -ne "")
        {
            Do
            {
                Start-Sleep -Seconds 2
                $MBXFoldersExist = [bool](Get-MailboxFolderStatistics -Identity $MbxAddr -ErrorAction SilentlyContinue)
            } while ($MBXFoldersExist -ne "True")

            $MBXFolders = Get-MailboxFolderStatistics -Identity $MbxAddr | % {$_.Identity.ToString().Split("\")[1..100] -join "\"}

            ForEach ($Folder in $MBXFolders)
            {
		        If ($Folder.Equals("Recoverable Items") -or $Folder.Equals("Calendar Logging") -or $Folder.Equals("Deletions") -or $Folder.Equals("Purges") -or $Folder.Equals("Versions") -or $Folder.Equals("Audits") -or $Folder.Equals("SubstrateHolds") -or $Folder.Equals("DiscoveryHolds"))
                {
				    # Ignore folder
				}
				else
                {
				    if ($Folder.Equals("Top of Information Store"))
                    {
                        Add-MailboxFolderPermission -Identity ($MbxAddr + ":\") -User $GrpAddr -AccessRights $Permission
                        $LineToWrite = "ADD  " + "`t" + $Permission + " permission to: " + $MbxAddr + ":\"
                    }
					else
                    {
   					    Add-MailboxFolderPermission -Identity ($MbxAddr + ":\" + $Folder) -User $GrpAddr -AccessRights $Permission
                        $LineToWrite = "ADD  " + "`t" + $Permission + " permission to: " + $MbxAddr + ":\" + $Folder
                    }
                    WriteReportEvent
				}
			}
        }
        else
        {
            write-host "Skipping this group since it does not follow our naming standard"
        }
    }	
}

function Build-ShrMbxMenuForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Shared Mailbox Admin Admin Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 550 ; $form.Height = 350  # Make the form wider

    $TopLoc = 20 
    ## Label and TextBox  
    ## Title Line
    $Global:lblDLSecTitleLine = New-Object System.Windows.Forms.Label   
        $lblDLSecTitleLine.Text = "Select Option:"
        $lblDLSecTitleLine.Top = $TopLoc ; $lblDLSecTitleLine.Left = 60; $lblDLSecTitleLine.Width=220 ;$lblDLSecTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($lblDLSecTitleLine)    # Add to Form 

    ## New ShrMbx
    $TopLoc = $TopLoc + 20
    $Global:chkNewShrMbx = New-Object Windows.Forms.RadioButton
        $Global:chkNewShrMbx.Left = 100; $Global:chkNewShrMbx.Width = 450; $Global:chkNewShrMbx.Top = $TopLoc
        $Global:chkNewShrMbx.Text = "Create a New Shared Mailbox" 
        $Global:chkNewShrMbx.Checked = $false   # set a default value 
        $Global:chkNewShrMbx.TabIndex = 1
        $Global:form.Controls.Add($Global:chkNewShrMbx) 
        # Obtain Value with: $Global:chkNewShrMbx.Checked
#        $Global:InputFocus = $Global:chkNewShrMbx

    ## Change Ownership
    $TopLoc = $TopLoc + 20
    $Global:chkChgOwner = New-Object Windows.Forms.RadioButton
        $Global:chkChgOwner.Left = 100; $Global:chkChgOwner.Width = 450; $Global:chkChgOwner.Top = $TopLoc  
        $Global:chkChgOwner.Text = "Change Shared Mailbox Ownership" 
        $Global:chkChgOwner.Checked = $false   # set a default value 
        $Global:chkChgOwner.TabIndex = 2
        $Global:form.Controls.Add($Global:chkChgOwner) 
        # Obtain Value with: $Global:chkChgOwner.Checked

    ## Add Access Group
    $TopLoc = $TopLoc + 20
    $Global:chkAccessGrp = New-Object Windows.Forms.RadioButton
        $Global:chkAccessGrp.Left = 100; $Global:chkAccessGrp.Width = 450; $Global:chkAccessGrp.Top = $TopLoc  
        $Global:chkAccessGrp.Text = "Add Access Group (.ED, .AU or .RE)" 
        $Global:chkAccessGrp.Checked = $false   # set a default value 
        $Global:chkAccessGrp.TabIndex = 3
        $Global:form.Controls.Add($Global:chkAccessGrp) 
        # Obtain Value with: $Global:chkAccessGrp.Checked

    ## Add/Remove Alias
    $TopLoc = $TopLoc + 20
    $Global:chkAlias = New-Object Windows.Forms.RadioButton
        $Global:chkAlias.Left = 100; $Global:chkAlias.Width = 450; $Global:chkAlias.Top = $TopLoc  
        $Global:chkAlias.Text = "Add/Remove Email Alias from Shared Mailbox" 
        $Global:chkAlias.Checked = $Global:chkAlias.Checked   # set a default value 
        $Global:chkAlias.TabIndex = 6
        $Global:form.Controls.Add($Global:chkAlias) 
        # Obtain Value with: $Global:chkAlias.Checked

    ## Change external access to mailbox
    $TopLoc = $TopLoc + 20
    $Global:chkUse = New-Object Windows.Forms.RadioButton
        $Global:chkUse.Left = 100; $Global:chkUse.Width = 450; $Global:chkUse.Top = $TopLoc  
        $Global:chkUse.Text = "Grant/Deny External Addresses Access to Shared Mailbox" 
        $Global:chkUse.Checked = $false   # set a default value 
        $Global:chkUse.TabIndex = 7
        $Global:form.Controls.Add($Global:chkUse) 
        # Obtain Value with: $Global:chkUse.Checked

    ## Change external access to mailbox
    $TopLoc = $TopLoc + 20
    $Global:chkStats = New-Object Windows.Forms.RadioButton
        $Global:chkStats.Left = 100; $Global:chkStats.Width = 450; $Global:chkStats.Top = $TopLoc  
        $Global:chkStats.Text = "Review Shared Mailbox Statistics" 
        $Global:chkStats.Checked = $false   # set a default value 
        $Global:chkStats.TabIndex = 7
        $Global:form.Controls.Add($Global:chkStats) 
        # Obtain Value with: $Global:chkStats.Checked

    ## Rename Shared Mailbox
    $TopLoc = $TopLoc + 20
    $Global:chkRenShrMbx = New-Object Windows.Forms.RadioButton
        $Global:chkRenShrMbx.Left = 100; $Global:chkRenShrMbx.Width = 450; $Global:chkRenShrMbx.Top = $TopLoc  
        $Global:chkRenShrMbx.Text = "Rename Shared Mailbox" 
        $Global:chkRenShrMbx.Checked = $false   # set a default value 
        $Global:chkRenShrMbx.TabIndex = 4 
        $Global:form.Controls.Add($Global:chkRenShrMbx) 
        # Obtain Value with: $Global:chkRenShrMbx.Checked

    ## Remove Shared Mailbox
    $TopLoc = $TopLoc + 20
    $Global:chkRemShrMbx = New-Object Windows.Forms.RadioButton
        $Global:chkRemShrMbx.Left = 100; $Global:chkRemShrMbx.Width = 450; $Global:chkRemShrMbx.Top = $TopLoc  
        $Global:chkRemShrMbx.Text = "Remove Shared Mailbox"
        $Global:chkRemShrMbx.TabIndex = 5
        $Global:form.Controls.Add($Global:chkRemShrMbx) 
        # Obtain Value with: $Global:chkRemShrMbx.Checked

    ## Restore Deleted mailbox
    $TopLoc = $TopLoc + 20
    $Global:UnDelete = New-Object Windows.Forms.RadioButton
        $Global:UnDelete.Left = 100; $Global:UnDelete.Width = 450; $Global:UnDelete.Top = $TopLoc  
        $Global:UnDelete.Text = "Restore a Deleted Shared Mailbox (under construction)"
        $Global:UnDelete.Checked = $false   # set a default value 
        $Global:UnDelete.TabIndex = 7
        $Global:form.Controls.Add($Global:UnDelete) 
        # Obtain Value with: $Global:UnDelete.Checked
}

Function Build-ShrMbxChgForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Shared Mailbox Additional Access Group"
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(550,250) #(W,H)

    ## Get Details to complete creation
    $Top = 10
    ## Display Name
    $Script:lblDispName = New-Object System.Windows.Forms.Label   
        $Script:lblDispName.Text = "Shared Mailbox Name:"
        $Script:lblDispName.Top = $Top; $Script:lblDispName.Left = 10; $Script:lblDispName.Width=150 ;$Script:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblDispName)    # Add to Form 
        # 
        $Script:txtDispName = New-Object Windows.Forms.TextBox  
        $Script:txtDispName.TabIndex = 0 # set Tab Order 
        $Script:txtDispName.Top = $Top; $Script:txtDispName.Left = 130; $Script:txtDispName.Width = 220;  
        $Script:txtDispName.Text = ""  # DisplayName
        $Global:form.Controls.Add($Script:txtDispName)    # Add to Form
        $Global:InputFocus = $Script:txtDispName
        $Script:txtDispName.Add_Click({
            $Script:txtDispName.Text = ""
            $Script:txtDispName.ReadOnly = $False
            Hide-ShrAddGrp
        })

    $Script:ButGetMbx = New-Object Windows.Forms.Button
        $Script:ButGetMbx.Location = New-object System.Drawing.Size(360,$Top)
        $Script:ButGetMbx.Size = new-Object System.Drawing.Size(130,20)
        $Script:ButGetMbx.Text = "Get Mailbox Details"
        $Script:ButGetMbx.TabIndex = 0
        $Global:form.Controls.Add($Script:ButGetMbx)
        $Script:ButGetMbx.Add_Click({
        write-host "`nRetreiving Mailbox Details for: " $Script:txtDispName.Text -ForegroundColor Cyan
            $MbxExists = [bool]($Mbx = Get-Mailbox $Script:txtDispName.Text -ErrorAction SilentlyContinue)
            If ($MbxExists -eq $True)
            {
                $Script:txtDispName.Text = $Mbx.DisplayName
                $Script:txtDispName.ReadOnly = $True
                UnHide-ShrAddGrp
                $Script:chkNewEDGrp.Visible = $True
                $MbxPerm = get-MailboxPermission $Script:txtDispName.Text |Where-Object {$_.User -like "MBX*"}
                foreach ($MbxPerm in $MbxPerm)
                {
                    $MbxPerm = get-MailboxPermission $Script:txtDispName.Text |Where-Object {$_.User -like "MBX*"}
                    If ($MbxPerm.AccessRights -eq "FullAccess")
                    {
                        $Script:txtEDGrpName.Text = $MbxPerm.User
                        If ($MbxPerm.User -like "*ED")
                        {
                            # Build Current Group Owner Details
                            $Own = (Get-DistributionGroup $MbxPerm.User).ManagedBy
                            Foreach ($o in $own)
                            {
                                $Script:Owners = $Script:Owners + ((get-mailbox $o).DisplayName -split ", ")[1] + " " + ((get-mailbox $o).DisplayName -split ", ")[0] + ", "
                            }
                            $Script:Owners = $Script:Owners.TrimEnd(", ")
                            $Script:txtGrpOwnr.Text = $Script:Owners
#                            $Script:txtGrpOwnr.Text = (Get-DistributionGroup $mbxPerm.User -ErrorAction SilentlyContinue).ManagedBy -join ", "
                            $Script:chkNewEDGrp.Visible = $False
                            $Script:lblEDGrpName.Visible = $True
                            $Script:txtEDGrpName.Visible = $True
                            $Script:txtEDGrpName.ReadOnly = $True
                        }
                    }
                }

                $Script:chkNewAUGrp.Visible = $True
                $Script:chkNewREGrp.Visible = $True
                $MbxFldrPerm = get-MailboxFolderPermission $Script:txtDispName.Text |Where-Object {$_.User -like "MBX*"}
                foreach ($MbxFldrPerm in $MbxFldrPerm)
                {
                    If ($Script:txtGrpOwnr.Text -eq "")
                    {
                        $Script:txtGrpOwnr.Text = (Get-DistributionGroup $MbxFldrPerm.User.DisplayName -ErrorAction SilentlyContinue).ManagedBy -join ", "
                    }
                    If ($MbxFldrPerm.User.DisplayName -like "*AU")
                    {
                        $Script:txtAUGrpName.Text = $MbxFldrPerm.User.DisplayName
                        $Script:chkNewAUGrp.Visible = $False
                        $Script:lblAUGrpName.Visible = $True
                        $Script:txtAUGrpName.Visible = $True
                    }
                    If ($MbxFldrPerm.User.DisplayName -like "*RE")
                    {
                        $Script:txtREGrpName.Text = $MbxFldrPerm.User.DisplayName
                        $Script:chkNewREGrp.Visible = $False
                        $Script:lblREGrpName.Visible = $True
                        $Script:txtREGrpName.Visible = $True
                    }
                }

                If (($Script:chkNewEDGrp.Visible -eq $True) -or ($Script:chkNewAUGrp.Visible -eq $True) -or ($Script:chkNewREGrp.Visible -eq $True))
                {
					If ($Script:txtGrpOwnr.Text -eq "")
					{
						$Script:txtGrpOwnr.Text = "No Owners Found - Enter Owners Employee # or Email Address"
						$Script:txtGrpOwnr.ReadOnly = $False
					}
                    $Global:OKButton.Text = "Create"
                    $Script:txtRptFile.Text	= "e:\Automation\AddFolderPermissions\Report\Report-AddFolderPermissions-" + (($Script:txtDispName.Text).Replace(" ","")) + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
                }
                else
                {
                    $Script:lblTicketNo.Visible = $False
                    $Script:txtTicketNo.Visible = $False
                    $Script:lblRptFile.Visible = $False
                    $Script:txtRptFile.Text = "All Access Groups Exist - No additional changes possible"
                }
            }
            else
            {
                Hide-ShrAddGrp
                $Script:txtDispName.Text = "Invalid"
                $Script:ButGetMbx.Visible = "True"
                Hide-ShrAddGrp
            }})

    $Top = $Top + 30
    ## Ticket Number
    $Script:lblTicketNo = New-Object System.Windows.Forms.Label
        $Script:lblTicketNo.Text = "Ticket Number:"
        $Script:lblTicketNo.Top = $Top; $Script:lblTicketNo.Left = 10; $Script:lblTicketNo.Width=150 ;$Script:lblTicketNo.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblTicketNo)    # Add to Form 
        # 
        $Script:txtTicketNo = New-Object Windows.Forms.ComboBox
        $Script:txtTicketNo.TabIndex = 0 # set Tab Order 
        $Script:txtTicketNo.Top = $Top; $Script:txtTicketNo.Left = 130; $Script:txtTicketNo.Width = 220;  
        $Script:txtTicketNo.Text = "TASK"
        [void] $Script:txtTicketNo.Items.Add("NoOwners/NoMembers Request")  # Add element to listbox 
        $Global:form.Controls.Add($Script:txtTicketNo)    # Add to Form

    $Top = $Top + 30
    ## Mailbox Owner
    $Script:lblGrpOwnr = New-Object System.Windows.Forms.Label   
        $Script:lblGrpOwnr.Text = "Mailbox Owner(s):"  
        $Script:lblGrpOwnr.Top = $Top; $Script:lblGrpOwnr.Left = 10; $Script:lblGrpOwnr.Width=150; $Script:lblGrpOwnr.AutoSize = $true 
        $Global:form.Controls.Add($lblGrpOwnr)    # Add to Form 
        # 
        $Script:txtGrpOwnr = New-Object Windows.Forms.TextBox  
        $Script:txtGrpOwnr.TabIndex = 0 # set Tab Order 
        $Script:txtGrpOwnr.Top = $Top; $Script:txtGrpOwnr.Left = 130; $Script:txtGrpOwnr.Width = 220;
        $Script:txtGrpOwnr.Text = ""
        $Script:txtGrpOwnr.ReadOnly = $True
        $Global:form.Controls.Add($Script:txtGrpOwnr)    # Add to Form 

    $Top = $Top + 30
    ## Editor Group Name
    $Script:lblEDGrpName = New-Object System.Windows.Forms.Label
        $Script:lblEDGrpName.Text = ".ED Group:"  
        $Script:lblEDGrpName.Top = $Top ; $Script:lblEDGrpName.Left = 10; $Script:lblEDGrpName.Width=150; $Script:lblEDGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblEDGrpName)    # Add to Form 
        #
        $Script:txtEDGrpName = New-Object Windows.Forms.TextBox  
        $Script:txtEDGrpName.TabIndex = 0 # set Tab Order 
        $Script:txtEDGrpName.Top = $Top; $Script:txtEDGrpName.Left = 130; $Script:txtEDGrpName.Width = 220;  
        $Script:txtEDGrpName.Text = ""
        $Script:txtEDGrpName.ReadOnly = $True
        $Global:form.Controls.Add($Script:txtEDGrpName)
        $Script:txtEDGrpName.Add_Click({
            If ($Script:lblEDGrpName.Text -ne ".ED Group:")
            {
                $Script:lblEDGrpName.Text = ".ED Group:"
                $Script:lblEDGrpName.Visible = $False
                $Script:txtEDGrpName.Visible = $False
                $Script:chkNewEDGrp.Checked = $False
                $Script:chkNewEDGrp.Visible = $True
            }
            If (($Script:chkNewEDGrp.Checked -eq $False) -and ($Script:chkNewAUGrp.Checked -eq $False) -and ($Script:chkNewREGrp.Checked -eq $False))
            {
                $Global:OKButton.Visible = $False
            }
        })

    $Script:chkNewEDGrp = New-Object Windows.Forms.CheckBox
        $Script:chkNewEDGrp.Left = 130; $Script:chkNewEDGrp.Width = 150; $Script:chkNewEDGrp.Top = $Top
        $Script:chkNewEDGrp.Text = "Create New ED Group" 
        $Script:chkNewEDGrp.Checked = $false   # set a default value
        $Script:chkNewEDGrp.Visible = $False
        $Global:form.Controls.Add($Script:chkNewEDGrp)
        $Script:chkNewEDGrp.Add_Click({
            $Script:chkNewEDGrp.Visible = $False
            $Script:txtEDGrpName.Text = "MBX." + $Script:txtDispName.Text + ".ED"
            $Script:lblEDGrpName.Text = "New .ED Group:"
            $Script:lblEDGrpName.Visible = $True
            $Script:txtEDGrpName.Visible = $True
            $Global:OKButton.Visible = $True
        })

    $Top = $Top + 30
    ## Author Group Name
    $Script:lblAUGrpName = New-Object System.Windows.Forms.Label
        $Script:lblAUGrpName.Text = ".AU Group:" 
        $Script:lblAUGrpName.Top = $Top; $Script:lblAUGrpName.Left = 10; $Script:lblAUGrpName.Width=150; $Script:lblAUGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblAUGrpName)    # Add to Form 
        #
        $Script:txtAUGrpName = New-Object Windows.Forms.TextBox  
        $Script:txtAUGrpName.TabIndex = 0 # set Tab Order 
        $Script:txtAUGrpName.Top = $Top; $Script:txtAUGrpName.Left = 130; $Script:txtAUGrpName.Width = 220;  
        $Script:txtAUGrpName.Text = ""
        $Script:txtAUGrpName.ReadOnly = $True
        $Global:form.Controls.Add($Script:txtAUGrpName)
        $Script:txtAUGrpName.Add_Click({
            If ($Script:lblAUGrpName.Text -ne ".AU Group:")
            {
                $Script:lblAUGrpName.Text = ".AU Group:"
                $Script:lblAUGrpName.Visible = $False
                $Script:txtAUGrpName.Visible = $False
                $Script:chkNewAUGrp.Checked = $False
                $Script:chkNewAUGrp.Visible = $True
            }
            If (($Script:chkNewEDGrp.Checked -eq $False) -and ($Script:chkNewAUGrp.Checked -eq $False) -and ($Script:chkNewREGrp.Checked -eq $False))
            {
                $Global:OKButton.Visible = $False
            }
        })

    $Script:chkNewAUGrp = New-Object Windows.Forms.CheckBox
        $Script:chkNewAUGrp.Left = 130; $Script:chkNewAUGrp.Width = 150; $Script:chkNewAUGrp.Top = $Top
        $Script:chkNewAUGrp.Text = "Create New AU Group" 
        $Script:chkNewAUGrp.Checked = $false   # set a default value
        $Script:chkNewAUGrp.Visible = $False
        $Global:form.Controls.Add($Script:chkNewAUGrp)
        $Script:chkNewAUGrp.Add_Click({
            $Script:chkNewAUGrp.Visible = $False
            $Script:txtAUGrpName.Text = "MBX." + $Script:txtDispName.Text + ".AU"
            $Script:lblAUGrpName.Text = "New .AU Group:"
            $Script:lblAUGrpName.Visible = $True
            $Script:txtAUGrpName.Visible = $True
            $Global:OKButton.Visible = $True
        })

    $Top = $Top + 30
    # Reader Group Name
    $Script:lblREGrpName = New-Object System.Windows.Forms.Label
        $Script:lblREGrpName.Text = ".RE Group:" 
        $Script:lblREGrpName.Top = $Top ; $Script:lblREGrpName.Left = 10; $Script:lblREGrpName.Width=150; $Script:lblREGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblREGrpName)    # Add to Form 
        #
        $Script:txtREGrpName = New-Object Windows.Forms.TextBox  
        $Script:txtREGrpName.TabIndex = 0 # set Tab Order 
        $Script:txtREGrpName.Top = $Top; $Script:txtREGrpName.Left = 130; $Script:txtREGrpName.Width = 220;  
        $Script:txtREGrpName.Text = ""
        $Script:txtREGrpName.ReadOnly = $True
        $Global:form.Controls.Add($Script:txtREGrpName)
        $Script:txtREGrpName.Add_Click({
            If ($Script:lblREGrpName.Text -ne ".RE Group:")
            {
                $Script:lblREGrpName.Text = ".RE Group:"
                $Script:lblREGrpName.Visible = $False
                $Script:txtREGrpName.Visible = $False
                $Script:chkNewREGrp.Checked = $False
                $Script:chkNewREGrp.Visible = $True
            }
            If (($Script:chkNewEDGrp.Checked -eq $False) -and ($Script:chkNewAUGrp.Checked -eq $False) -and ($Script:chkNewREGrp.Checked -eq $False))
            {
                $Global:OKButton.Visible = $False
            }
        })

    $Script:chkNewREGrp = New-Object Windows.Forms.CheckBox
        $Script:chkNewREGrp.Left = 130; $Script:chkNewREGrp.Width = 150; $Script:chkNewREGrp.Top = $Top
        $Script:chkNewREGrp.Text = "Create New RE Group" 
        $Script:chkNewREGrp.Checked = $false   # set a default value
        $Script:chkNewREGrp.Visible = $False
        $Global:form.Controls.Add($Script:chkNewREGrp)
        $Script:chkNewREGrp.Add_Click({
            $Script:chkNewREGrp.Visible = $False
            $Script:txtREGrpName.Text = "MBX." + $Script:txtDispName.Text + ".RE"
            $Script:lblREGrpName.Text = "New .RE Group:"
            $Script:lblREGrpName.Visible = $True
            $Script:txtREGrpName.Visible = $True
            $Global:OKButton.Visible = $True
        })

    If ( $Global:form.Text -notlike "*Additional Access*")
    {
        $Top = $Top + 30
        ## Employee # to Add/Remove
            $Script:lblEmpNo = New-Object System.Windows.Forms.Label
            $Script:lblEmpNo.Text = "Emp# or Address:" 
            $Script:lblEmpNo.Top = $Top ; $Script:lblEmpNo.Left = 10; $Script:lblEmpNo.Width=150; $Script:lblEmpNo.AutoSize = $true 
            $Global:form.Controls.Add($Script:lblEmpNo)    # Add to Form 
            #
            $Script:txtEmpNo = New-Object Windows.Forms.TextBox  
            $Script:txtEmpNo.TabIndex = 0 # set Tab Order 
            $Script:txtEmpNo.Top = $Top; $Script:txtEmpNo.Left = 130; $Script:txtEmpNo.Width = 220;  
            $Script:txtEmpNo.Text = ""
            $Global:form.Controls.Add($Script:txtEmpNo)

        $Top = $Top + 30
        ## Add Owner
        $Script:chkAddOwner = New-Object Windows.Forms.RadioButton
            $Script:chkAddOwner.Left = 50; $Script:chkAddOwner.Width = 80; $Script:chkAddOwner.Top = $Top
            $Script:chkAddOwner.Text = "Add Owner" 
            $Script:chkAddOwner.Checked = $false   # set a default value 
            $Script:chkAddOwner.TabIndex = 5
            $Global:form.Controls.Add($Script:chkAddOwner) 
            # Obtain Value with: $Script:chkRemoveOwner.Checked
               
        ## Remove Owner
        $Script:chkRemoveOwner = New-Object Windows.Forms.RadioButton
            $Script:chkRemoveOwner.Left = 150; $Script:chkRemoveOwner.Width = 110; $Script:chkRemoveOwner.Top = $Top
            $Script:chkRemoveOwner.Text = "Remove Owner" 
            $Script:chkRemoveOwner.Checked = $false   # set a default value 
            $Script:chkRemoveOwner.TabIndex = 5
            $Global:form.Controls.Add($Script:chkRemoveOwner) 
            # Obtain Value with: $Script:chkRemoveOwner.Checked

        ## Replace Owner
        $Script:chkReplaceOwner = New-Object Windows.Forms.RadioButton
            $Script:chkReplaceOwner.Left = 270; $Script:chkReplaceOwner.Width = 120; $Script:chkReplaceOwner.Top = $Top
            $Script:chkReplaceOwner.Text = "Replace Owners" 
            $Script:chkReplaceOwner.Checked = $false   # set a default value 
            $Script:chkReplaceOwner.TabIndex = 5
            $Global:form.Controls.Add($Script:chkReplaceOwner) 
            # Obtain Value with: $Script:chkReplaceOwner.Checked

        $Top = $Top + 30
        ## Authorized Requestor
        $Script:chkAuthorized = New-Object Windows.Forms.checkbox 
            $Script:chkAuthorized.Left = 130; $Script:chkAuthorized.Width = 200; $Script:chkAuthorized.Top = $Top
            $Script:chkAuthorized.Text = "Authorized Requestor" 
            $Script:chkAuthorized.Checked = $false   # set a default value 
            $Script:chkAuthorized.TabIndex = 5
            $Global:form.Controls.Add($Script:chkAuthorized) 
            # Obtain Value with: $Script:chkAuthorized.Checked
    }

    $Top = $Top + 40
    $Script:lblRptFile = New-Object System.Windows.Forms.Label   
        $Script:lblRptFile.Text = "ReportFile:"  
        $Script:lblRptFile.Top = $Top; $Script:lblRptFile.Left = 10; $Script:lblRptFile.Width=60;
        $Global:form.Controls.Add($Script:lblRptFile)    # Add to Form 
        # 
        $Script:txtRptFile = New-Object Windows.Forms.TextBox  
        $Script:txtRptFile.TabIndex = 0 # set Tab Order 
        $Script:txtRptFile.Top = $Top; $Script:txtRptFile.Left = 75; $Script:txtRptFile.Width = 330;
        $Script:txtRptFile.Text = ""
        $Script:txtRptFile.ReadOnly = $True
        $Global:form.Controls.Add($Script:txtRptFile)    # Add to Form 

        Add-FormStandardButtons
}

Function Build-OwnerMailTip
{
    $Own = (Get-DistributionGroup $GrpName).ManagedBy
    $Script:Owners = "Owners: "
    Foreach ($o in $own)
    {
        $Script:Owners = $Script:Owners + ((get-mailbox $o).DisplayName -split ", ")[1] + " " + ((get-mailbox $o).DisplayName -split ", ")[0] + ", "
    }
    $Script:Owners = $Script:Owners.TrimEnd(", ")
}

Function Hide-ShrAddGrp
{
    $Global:form.Width = 520 ; $form.Height = 130
    $Script:ButGetMbx.Visible = $True
    $Script:lblTicketNo.Visible = $False
    $Script:txtTicketNo.Visible = $False
    $Script:lblGrpOwnr.Visible = $False
    $Script:txtGrpOwnr.Visible = $False
    $Script:lblEDGrpName.Visible = $False
    $Script:txtEDGrpName.Visible = $False
    $Script:txtEDGrpName.Text = ""
    $Script:lblAUGrpName.Visible = $False
    $Script:txtAUGrpName.Visible = $False
    $Script:txtAUGrpName.Text = ""
    $Script:lblREGrpName.Visible = $False
    $Script:txtREGrpName.Visible = $False
    $Script:txtREGrpName.Text = ""
    $Script:lblGrpOwnr.Visible = $False
    $Script:txtGrpOwnr.Visible = $False
    $Script:lblEDGrpName.Text = ".ED Group:"
    $Script:chkNewEDGrp.Visible = $False
    $Script:chkNewEDGrp.Checked = $False
    $Script:lblAUGrpName.Text = ".AU Group:"
    $Script:chkNewAUGrp.Visible = $False
    $Script:chkNewAUGrp.Checked = $False
    $Script:lblREGrpName.Text = ".RE Group:"
    $Script:chkNewREGrp.Visible = $False
    $Script:chkNewREGrp.Checked = $False
    $Script:lblRptFile.Visible = $False
    $Script:txtRptFile.Visible = $False
    $Global:OKButton.Visible = $False
}

Function UnHide-ShrAddGrp
{
    $Global:form.Width = 450 ; $form.Height = 315
    $Script:ButGetMbx.Visible = $False
    $Script:lblTicketNo.Visible = $True
    $Script:txtTicketNo.Visible = $True
    $Script:lblGrpOwnr.Visible = $True
    $Script:txtGrpOwnr.Visible = $True
    $Script:lblGrpOwnr.Visible = $True
    $Script:txtGrpOwnr.Visible = $True
    $Script:lblRptFile.Visible = $True
    $Script:txtRptFile.Visible = $True
}

$MbxChg = "1"
    Build-ShrMbxMenuForm
    Add-FormStandardButtons
    $Global:InputFocus = $Global:okButton
    Publish-Form

Do
{
    if ($Global:stopwatch.elapsed -ge $Global:reconnectThreshold)
    {
        # Close all sessions
        write-host "Connection Threshold Exceeded -- Reconnecting to Office365" -ForegroundColor Red
        invoke-expression -Command E:\O365AdminShared\Scripts\ConnectO365.ps1
        $Global:stopwatch = [diagnostics.stopwatch]::StartNew()
    }

    If ($Global:chkNewShrMbx.checked -eq "Checked")
    {
        invoke-expression -Command .\ShrMbxNew.ps1
        write-host "Shared Mailbox Creation Complete" -ForegroundColor Red
    }

    If ($Global:chkChgOwner.Checked -eq "Checked")
    {
		invoke-expression -Command .\ShrMbxChgOwner.ps1
    }
   
    If ($Global:chkAccessGrp.checked -eq "Checked")                 
    {
        invoke-expression -Command .\ShrMbxAddAccessGroup.ps1
    }
                
    If ($Global:chkRenShrMbx.Checked -eq "Checked")
    {
        invoke-expression -Command .\RenameShrMbx.ps1
        invoke-Expression -Command e:\O365AdminShared\EMailTemplates\SharedMailboxRenamed.oft
        write-host "Shared Mailbox Rename Complete" -ForegroundColor Red
    }

    If ($Global:chkUse.Checked -eq "Checked")
    {
        invoke-expression -Command .\ShrMbxUse.ps1
    }

    If ($Global:chkRemShrMbx.Checked -eq "Checked")
    {
        If (Test-Path "c:\temp\RemoveShrMbx.csv")
        {
            invoke-expression -Command .\ShrMbxRemove-wInputFile.ps1
        }
        else
        {
            invoke-expression -Command .\ShrMbxRemove.ps1
        }
    }

    If ($Global:chkAlias.Checked -eq "Checked")
    {

    ## Need to modify to check to see if the account is a Service Account if it is it must be modified using AD Tools
                    Write-Host "Add/Remove Additional eMail Alias to a Shared Mailbox" -ForegroundColor Magenta
                    Write-Host "Enter the name of the Shared Mailbox " -ForegroundColor Yellow -NoNewline
                    $Mbx = Read-Host
                    $ShrExists = [bool]($MbxAlias = Get-Mailbox $Mbx -ErrorAction SilentlyContinue)
                    If ($ShrExists -eq "True")
                    {
                        write-host "Current Mailbox Owners: " -ForegroundColor Yellow -NoNewline
                        $MbxGroup = "mbx." + $mbx + ".ED"
                        $MbxOwner = Get-DistributionGroup $MbxGroup -ErrorAction SilentlyContinue
                        write-host $MbxOwner.ManagedBy            
                        Write-Host "Current Email aliases on" $Mbx ": " -ForegroundColor Yellow
                        $MbxAlias.EmailAddresses
                        write-host ""
                        write-host "Do you wish to Add or Remove an Alias enter (A = Add/R = Remove)? " -ForegroundColor Yellow -NoNewline
                        $AddRem = Read-Host
                        If ($AddRem -eq "A")
                        {
                            write-host ""
                            write-host "Enter the eMail Alias you would like to add: " -ForegroundColor Yellow -NoNewline
                            $NewAlias = Read-Host
                            $AddAlias = "smtp:" + $NewAlias
                            Set-Mailbox $Mbx -EmailAddresses @{Add=$AddAlias}
                            Get-Mailbox $Mbx |ft *Addresses*
                            Write-Host "Should this be made the new Primary SMTP Address for this Mailbox (Y/N) ?" -ForegroundColor Yellow -NoNewline
                            $NewPrim = Read-Host
                            If ($NewPrim -eq "Y")
                            {
                                Set-Mailbox $Mbx -PrimarySmtpAddress $NewAlias
                            }
                        }
                        elseif ($AddRem -eq "R")
                        {
                            write-host ""
                            write-host "Enter the eMail Alias you would like to remove: " -ForegroundColor Yellow -NoNewline
                            $RemAlias = Read-Host
                            $RemAlias = "smtp:" + $RemAlias
                            Set-Mailbox $Mbx -EmailAddresses @{Remove=$RemAlias}
                            Get-Mailbox $Mbx |ft *Addresses*
                        }
                        else
                        {
                            Write-Host "No changes made to the configured aliases for this Shared Mailbox"
                        }
#                        pause
                    }
                    else
                    {
                        Write-Host ""
                        Write-Host "Shared mailbox does not exist"
                    }
                }

    If ($Global:chkStats.Checked -eq "Checked")
    {
       invoke-expression -Command .\MailboxStatistics.ps1
    }

    If ($Global:UnDelete.Checked -eq "Checked")
    {
       invoke-expression -Command .\RecoverShrMbx.ps1
    }

    Build-ShrMbxMenuForm
    Add-FormStandardButtons
    $Global:InputFocus = $Global:okButton
    Publish-Form

}While ($Global:Result -eq "OK")	