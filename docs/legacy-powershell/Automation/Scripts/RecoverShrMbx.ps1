#
# 04/10/2023 - SAG - Modified the creation of the new security group to accept messages only from Enterprise Messaging Services to prevent people from using the MBX group as distribution groups.
# 08/22/2024 - SAG - Added code to add the year that the script is run and to create the new directory if it does not exist.
#

Function Build-RecoverShrMbxForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Recover Deleted Shared Mailbox"
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(550,250) #(W,H)

    $Top = 30
    ## ListBox - Fill with Names of all Deleted Mailboxes
    $Global:lblLoc = New-Object System.Windows.Forms.Label   
        $Global:lblLoc.Text = "Deleted Shared mailboxes:"; $Global:lblLoc.Top = $Top; $Global:lblLoc.Left = 5; $Global:lblLoc.Autosize = $true  
        $Global:form.Controls.Add($Global:lblLoc)  
        # Listbox for Location Name 
        $Global:locListBox = New-Object System.Windows.Forms.ListBox  
        $Global:locListBox.Top = $Top; $locListBox.Left = 160; $locListBox.Height = 100; $LocListBox.Width = 370;
        $Global:locListBox.TabIndex = 1
        foreach ($m in $mbx)
        {
            [void] $Global:locListBox.Items.Add($m.Name.TrimStart())  # Add element to listbox 
        }
        $Global:form.Controls.Add($Global:locListBox) #Add listbox to form 
        # Obtain Value with: $Global:locListBox.SelectedItem
        Add-FormStandardButtons
}

Function Build-RecoverShrMbxDetails
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Recover Deleted Shared Mailbox Details"
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(600,450) #(W,H)

    ## Get Details to complete recovery
    $Top = 30
    ## Display Name
    $Script:lblDispName = New-Object System.Windows.Forms.Label   
        $Script:lblDispName.Text = "Mailbox Name:"
        $Script:lblDispName.Top = $Top; $Script:lblDispName.Left = 10; $Script:lblDispName.Width=150 ;$Script:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblDispName)    # Add to Form 
        # 
        $Script:txtDispName = New-Object Windows.Forms.TextBox  
        $Script:txtDispName.TabIndex = 0 # set Tab Order 
        $Script:txtDispName.Top = $Top; $Script:txtDispName.Left = 130; $Script:txtDispName.Width = 220;  
        $Script:txtDispName.Text = $SelMbx.DisplayName  # DisplayName
        $Global:form.Controls.Add($Script:txtDispName)    # Add to Form

    $Top = $Top + 30
    ## Mailbox Alias
    $Script:lblMbxAlias = New-Object System.Windows.Forms.Label   
        $Script:lblMbxAlias.Text = "Mailbox Alias"
        $Script:lblMbxAlias.Top = $Top; $Script:lblMbxAlias.Left = 10; $Script:lblMbxAlias.Width=150 ;$Script:lblMbxAlias.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblMbxAlias)    # Add to Form 
        # 
        $Script:txtMbxAlias = New-Object Windows.Forms.TextBox  
        $Script:txtMbxAlias.TabIndex = 0 # set Tab Order 
        $Script:txtMbxAlias.Top = $Top; $Script:txtMbxAlias.Left = 130; $Script:txtMbxAlias.Width = 220;  
        $Script:txtMbxAlias.Text = $SelMbx.Alias
        $Global:form.Controls.Add($Script:txtMbxAlias)    # Add to Form

    $Top = $Top + 30
    ## Mailbox Primary Address
    $Script:lblMbxAddr = New-Object System.Windows.Forms.Label   
        $Script:lblMbxAddr.Text = "Mailbox Address"
        $Script:lblMbxAddr.Top = $Top; $Script:lblMbxAddr.Left = 10; $Script:lblMbxAddr.Width=150 ;$Script:lblMbxAddr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblMbxAddr)    # Add to Form 
        # 
        $Script:txtMbxAddr = New-Object Windows.Forms.TextBox  
        $Script:txtMbxAddr.TabIndex = 0 # set Tab Order 
        $Script:txtMbxAddr.Top = $Top; $Script:txtMbxAddr.Left = 130; $Script:txtMbxAddr.Width = 220;  
        $Script:txtMbxAddr.Text = $SelMbx.PrimarySmtpAddress
        $Global:form.Controls.Add($Script:txtMbxAddr)    # Add to Form

If ($SelMbx.ForwardingAddress.Length -ne 0)
{
    $Top = $Top + 30
    ## Mailbox Forwarding Address
    $Script:lblMbxForwAddr = New-Object System.Windows.Forms.Label   
        $Script:lblMbxForwAddr.Text = "Forwarding Address"
        $Script:lblMbxForwAddr.Top = $Top; $Script:lblMbxForwAddr.Left = 10; $Script:lblMbxForwAddr.Width=150 ;$Script:lblMbxForwAddr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblMbxForwAddr)    # Add to Form 
        # 
        $Script:txtMbxForwAddr = New-Object Windows.Forms.TextBox  
        $Script:txtMbxForwAddr.TabIndex = 0 # set Tab Order 
        $Script:txtMbxForwAddr.Top = $Top; $Script:txtMbxForwAddr.Left = 130; $Script:txtMbxForwAddr.Width = 220;  
        $Script:txtMbxForwAddr.Text = $SelMbx.ForwardingAddress
        $Global:form.Controls.Add($Script:txtMbxForwAddr)    # Add to Form
}

If ($SelMbx.ForwardingSmtpAddress.Length -ne 0)
{
    $Top = $Top + 30
    ## Mailbox SMTP Forwarding Address
    $Script:lblMbxSMTPForwAddr = New-Object System.Windows.Forms.Label   
        $Script:lblMbxSMTPForwAddr.Text = "SMTP Forwarding Address"
        $Script:lblMbxSMTPForwAddr.Top = $Top; $Script:lblMbxSMTPForwAddr.Left = 10; $Script:lblMbxSMTPForwAddr.Width=150 ;$Script:lblMbxSMTPForwAddr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblMbxSMTPForwAddr)    # Add to Form 
        # 
        $Script:txtMbxSMTPForwAddr = New-Object Windows.Forms.TextBox  
        $Script:txtMbxSMTPForwAddr.TabIndex = 0 # set Tab Order 
        $Script:txtMbxSMTPForwAddr.Top = $Top; $Script:txtMbxSMTPForwAddr.Left = 130; $Script:txtMbxSMTPForwAddr.Width = 220;  
        $Script:txtMbxSMTPForwAddr.Text = $SelMbx.PrimarySmtpAddress
        $Global:form.Controls.Add($Script:txtMbxSMTPForwAddr)    # Add to Form
}

    $Top = $Top + 30
    ## Mailbox Owners
    $Script:lblMbxOwner = New-Object System.Windows.Forms.Label   
        $Script:lblMbxOwner.Text = "Mailbox Owner(s):"
        $Script:lblMbxOwner.Top = $Top; $Script:lblMbxOwner.Left = 10; $Script:lblMbxOwner.Width=150 ;$Script:lblMbxOwner.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblMbxOwner)    # Add to Form 
        # 
        $Script:txtMbxOwner = New-Object Windows.Forms.TextBox  
        $Script:txtMbxOwner.TabIndex = 0 # set Tab Order 
        $Script:txtMbxOwner.Top = $Top; $Script:txtMbxOwner.Left = 130; $Script:txtMbxOwner.Width = 220;  
        $Script:txtMbxOwner.Text = "Enter emp# separated by commas"
        $Global:form.Controls.Add($Script:txtMbxOwner)    # Add to Form

    $Top = $Top + 30
    ## Ticket Number
    $Script:lblTicketNo = New-Object System.Windows.Forms.Label   
        $Script:lblTicketNo.Text = "Ticket No.:"
        $Script:lblTicketNo.Top = $Top; $Script:lblTicketNo.Left = 10; $Script:lblTicketNo.Width=150 ;$Script:lblTicketNo.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblTicketNo)    # Add to Form 
        # 
        $Script:txtTicketNo = New-Object Windows.Forms.TextBox  
        $Script:txtTicketNo.TabIndex = 0 # set Tab Order 
        $Script:txtTicketNo.Top = $Top; $Script:txtTicketNo.Left = 130; $Script:txtTicketNo.Width = 220;  
        $Script:txtTicketNo.Text = "TASK"
        $Global:form.Controls.Add($Script:txtTicketNo)    # Add to Form

    $Top = $Top + 30
    ## Mailbox Owners
    $Script:lblGrpTitles = New-Object System.Windows.Forms.Label   
        $Script:lblGrpTitles.Text = "                        Group Name                                           Alias                                     Members"
        $Script:lblGrpTitles.Top = $Top; $Script:lblGrpTitles.Left = 10; $Script:lblGrpTitles.Width=150 ;$Script:lblGrpTitles.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblGrpTitles)    # Add to Form 

    $Top = $Top + 30
    ## ED Group
    $Script:lblEDGroup = New-Object System.Windows.Forms.Label   
        $Script:lblEDGroup.Text = "ED Group:"
        $Script:lblEDGroup.Top = $Top; $Script:lblEDGroup.Left = 10; $Script:lblEDGroup.Width=150 ;$Script:lblEDGroup.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblEDGroup)    # Add to Form 
        # 
        $Script:txtEDGroup = New-Object Windows.Forms.TextBox  
        $Script:txtEDGroup.TabIndex = 0 # set Tab Order 
        $Script:txtEDGroup.Top = $Top; $Script:txtEDGroup.Left = 70; $Script:txtEDGroup.Width = 150;  
        $Script:txtEDGroup.Text = ""
        $Global:form.Controls.Add($Script:txtEDGroup)    # Add to Form
        $Script:txtEDGroup.Add_Click{
            $Script:txtEDGroup.Text = "MBX." + $Script:txtDispName.Text + ".ED"
            $Script:txtEDAlias.Text = "MBX." + $Script:txtMbxalias.Text + ".ED"
            $Script:txtEDMem.Text = "Enter Emp# separated by commas"
            $Global:InputFocus = $Global:txtEDMem
        }

        $Script:txtEDAlias = New-Object Windows.Forms.TextBox  
        $Script:txtEDAlias.TabIndex = 0 # set Tab Order 
        $Script:txtEDAlias.Top = $Top; $Script:txtEDAlias.Left = 230; $Script:txtEDAlias.Width = 150;  
        $Script:txtEDAlias.Text = ""
        $Global:form.Controls.Add($Script:txtEDAlias)    # Add to Form

        $Script:txtEDMem = New-Object Windows.Forms.TextBox  
        $Script:txtEDMem.TabIndex = 0 # set Tab Order 
        $Script:txtEDMem.Top = $Top; $Script:txtEDMem.Left = 390; $Script:txtEDMem.Width = 150;  
        $Script:txtEDMem.Text = ""
        $Global:form.Controls.Add($Script:txtEDMem)    # Add to Form

    $Top = $Top + 30
    ## AU Group
    $Script:lblAUGroup = New-Object System.Windows.Forms.Label   
        $Script:lblAUGroup.Text = "AU Group:"
        $Script:lblAUGroup.Top = $Top; $Script:lblAUGroup.Left = 10; $Script:lblAUGroup.Width=150 ;$Script:lblAUGroup.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblAUGroup)    # Add to Form 
        # 
        $Script:txtAUGroup = New-Object Windows.Forms.TextBox  
        $Script:txtAUGroup.TabIndex = 0 # set Tab Order 
        $Script:txtAUGroup.Top = $Top; $Script:txtAUGroup.Left = 70; $Script:txtAUGroup.Width = 150;  
        $Script:txtAUGroup.Text = ""
        $Global:form.Controls.Add($Script:txtAUGroup)    # Add to Form
        $Global:InputFocus = $Script:txtAUGroup
        $Script:txtAUGroup.Add_Click{
            $Script:txtAUGroup.Text = "MBX." + $Script:txtDispName.Text + ".AU"
            $Script:txtAUAlias.Text = "MBX." + $Script:txtMbxalias.Text + ".AU"
            $Script:txtAUMem.Text = "Enter Emp# separated by commas"
            $Global:InputFocus = $Global:txtAUMem
        }

        $Script:txtAUAlias = New-Object Windows.Forms.TextBox  
        $Script:txtAUAlias.TabIndex = 0 # set Tab Order 
        $Script:txtAUAlias.Top = $Top; $Script:txtAUAlias.Left = 230; $Script:txtAUAlias.Width = 150;  
        $Script:txtAUAlias.Text = ""
        $Global:form.Controls.Add($Script:txtAUAlias)    # Add to Form
        $Global:InputFocus = $Script:txtAUAlias

        $Script:txtAUMem = New-Object Windows.Forms.TextBox  
        $Script:txtAUMem.TabIndex = 0 # set Tab Order 
        $Script:txtAUMem.Top = $Top; $Script:txtAUMem.Left = 390; $Script:txtAUMem.Width = 150;  
        $Script:txtAUMem.Text = ""
        $Global:form.Controls.Add($Script:txtAUMem)    # Add to Form

    $Top = $Top + 30
    ## RE Group
    $Script:lblREGroup = New-Object System.Windows.Forms.Label   
        $Script:lblREGroup.Text = "RE Group:"
        $Script:lblREGroup.Top = $Top; $Script:lblREGroup.Left = 10; $Script:lblREGroup.Width=150 ;$Script:lblREGroup.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblREGroup)    # Add to Form 
        # 
        $Script:txtREGroup = New-Object Windows.Forms.TextBox  
        $Script:txtREGroup.TabIndex = 0 # set Tab Order 
        $Script:txtREGroup.Top = $Top; $Script:txtREGroup.Left = 70; $Script:txtREGroup.Width = 150;  
        $Script:txtREGroup.Text = ""
        $Global:form.Controls.Add($Script:txtREGroup)    # Add to Form
        $Global:InputFocus = $Script:txtREGroup
        $Script:txtREGroup.Add_Click{
            $Script:txtREGroup.Text = "MBX." + $Script:txtDispName.Text + ".RE"
            $Script:txtREAlias.Text = "MBX." + $Script:txtMbxalias.Text + ".RE"
            $Script:txtREMem.Text = "Enter Emp# separated by commas"
            $Global:InputFocus = $Global:txtREMem
        }

        $Script:txtREAlias = New-Object Windows.Forms.TextBox  
        $Script:txtREAlias.TabIndex = 0 # set Tab Order 
        $Script:txtREAlias.Top = $Top; $Script:txtREAlias.Left = 230; $Script:txtREAlias.Width = 150;  
        $Script:txtREAlias.Text = ""
        $Global:form.Controls.Add($Script:txtREAlias)    # Add to Form
        $Global:InputFocus = $Script:txtREAlias

        $Script:txtREMem = New-Object Windows.Forms.TextBox  
        $Script:txtREMem.TabIndex = 0 # set Tab Order 
        $Script:txtREMem.Top = $Top; $Script:txtREMem.Left = 390; $Script:txtREMem.Width = 150;  
        $Script:txtREMem.Text = ""
        $Global:form.Controls.Add($Script:txtREMem)    # Add to Form

    $Top = $Top + 40
    ## Report File
    $Script:lblRptFile = New-Object System.Windows.Forms.Label   
        $Script:lblRptFile.Text = "Report File:" 
        $Script:lblRptFile.Top = $Top ; $Script:lblRptFile.Left = 10; $Script:lblRptFile.Width=120;
        $form.Controls.Add($Script:lblRptFile)    # Add to Form 
        # 
        $Script:txtRptFile = New-Object Windows.Forms.TextBox
        $Script:txtRptFile.Top = $Top; $Script:txtRptFile.Left = 130; $Script:txtRptFile.Width = 400; $Script:txtRptFile.Height = 60
        $Script:txtRptFile.ReadOnly = $True; $Script:txtRptFile.TabStop = $False
        $Script:txtRptFile.Text = $Path + "\RecoverShrMbx-" + ($Script:txtMbxAlias.Text -Replace("['.()/\- ]","")) + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
#        $Script:txtRptFile.Text = "e:\Automation\RecoverShrMbx\Report\RecoverShrMbx-" + ($Script:txtMbxAlias.Text -Replace("['.()/\- ]","")) + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
        $Global:form.Controls.Add($Script:txtRptFile)    # Add to Form

        Add-FormStandardButtons
        $Global:okButton.Text = "Recover"
}

Function Recreate-AccessGroups
{
    # Create the USG
	New-DistributionGroup -Name $GrpName `
	    -PrimarySmtpAddress $GrpAddr `
		-Alias $Script:txtMbxAlias.Text `
		-ManagedBy $DGManagedByMembers `
		-RequireSenderAuthenticationEnabled $TRUE `
		-Type Security | Out-Null

    Set-Distributiongroup $GrpName -AcceptMessagesOnlyFrom "EnterpriseMessagingServices@ul.com"

    $Owners = "Owners: " + ((Get-DistributionGroup $GrpName).ManagedBy -join (", "))
    Set-Group -identity $GrpName `
        -Notes ($Owners + " - Per: " + $Task)
 
    # Add USG members
    write-host ("Adding Members to USG: " + $GrpName + " (" + $GrpAddr + ")")
    $LineToWrite = "MEMBER" + "`t" + "Adding Members to Distribution Group " + $GrpName
    WriteReportEvent
		
    $addMember = $GrpMem.split(",")
    if ($? -eq $true)
    {
        ForEach ($member in $addMember)
        {
            $member = $member.Trim()
			if (Get-Mailbox $member)
            {
				Add-DistributionGroupMember $GrpAddr -Member $member -BypassSecurityGroupManagerCheck
				write-host "Added member:" $member
				$LineToWrite = "PASS" + "`t" + "Added " + $member
			}
			else
            {
				write-host "ERROR finding member: " $member
				$LineToWrite = "FAIL" + "`t" + "Failed adding " + $member
			}
            WriteReportEvent
		}
	}	
	else
    {
		write-host "No members to add"
	    $LineToWrite = "FAIL" + "`t" + "No Members to Add"
        WriteReportEvent
    }                        
}

$Year = (get-date).ToString("yyyy")
$Path = "e:\Automation\RecoverShrMbx\Report\" + $Year
If (Test-Path $path) {} else {New-Item -Path $Path -ItemType Directory}

write-host "Starting script to recover a shared mailbox that is available for recovery" -ForegroundColor Green

$mbx = get-mailbox -SoftDeletedMailbox -SortBy Name |Where-Object {$_.RecipientTypeDetails -eq "SharedMailbox"}
write-host "Number of maibloxes: " $mbx.count
If ($mbx.count -gt 0)
{
    Build-RecoverShrMbxForm
    Publish-Form
    If ($Global:Result -eq "OK")
    {
        $SelMbx = get-mailbox $Global:locListBox.SelectedItem -SoftDeletedMailbox
        Build-RecoverShrMbxDetails
        Publish-Form

        Do
        {
            $Refresh = "No"
            write-host "Checking input details provided" -ForegroundColor Cyan
            If ($Global:Result -eq "OK")
            {
                If ($Script:txtTicketNo.Text -eq "TASK")
                {
                    write-host "   You must enter the ticket number for this request" -ForegroundColor Red
                    $Refresh = "Yes"
                }

                #If Owners is not entered stop
                If ($Script:txtMbxOwner.Text -like "Enter*")
                {
                    write-host "   You must provide the employee number of the employee owner.  If there are multiple separate with commas" -ForegroundColor Red
                    $Refresh = "Yes"
                }

                #If members are not entered stop
                If (($Script:txtEDGroup.Text -like "MBX*") -and ($Script:txtEDMem.Text -like "Enter*"))
                {
                    write-host "   You must provide the employee number of editor group members.  If there are multiple separate with commas" -ForegroundColor Red
                    $Refresh = "Yes"
                }

                If (($Script:txtAUGroup.Text -like "MBX*") -and ($Script:txtAUMem.Text -like "Enter*"))
                {
                    write-host "   You must provide the employee number of the author group members.  If there are multiple separate with commas" -ForegroundColor Red
                    $Refresh = "Yes"
                }

                If (($Script:txtREGroup.Text -like "MBX*") -and ($Script:txtREMem.Text -like "Enter*"))
                {
                    write-host "   You must provide the employee number of the read group members.  If there are multiple separate with commas" -ForegroundColor Red
                    $Refresh = "Yes"
                }

                If ($Refresh -eq "Yes")
                {
                    Publish-Form
                }
                else
                {
                    $Reportfile = $Script:txtRptFile.Text
                    $LineToWrite = $RecordEvent + "STAR" + "`t" + "Restore Deleted Shard Maibox Script"
                    WriteReportEvent
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI
                    WriteReportEvent
 	                $LineToWrite = $RecordEvent + "INFO" + "`t" + "Script Last Change Date: 4/4/2023" + "`n"
                    WriteReportEvent
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
                    WriteReportEvent

                    write-host "Start Recovery" -ForegroundColor Yellow
				    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Restoring SoftDeleted Mailbox: " + $Script:txtDispName.Text
				    WriteReportEvent
				    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Ticket Number: " + $Script:txtTicketNo.Text
				    WriteReportEvent

                    Undo-SoftDeletedMailbox $Script:txtDispName.Text

                    If ($Script:txtEDGroup.Text -ne "")
                    {
    				    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Recreating the Editors Group" + $Script:txtEDGroup.Text
	    			    WriteReportEvent
                        $DGManagedByMembers = ($Script:txtMbxOwner.Text.Split(",")).Trim()
                        $GrpName = $Script:txtEDGroup.Text
                        $GrpAddr = $Script:txtEDAlias.Text + "@ul.com"
                        $GrpMem = $Script:txtEDMem.Text
                        $Task = $Script:txtTicketNo.Text
                        Recreate-AccessGroups
                        $Perm = [bool](get-mailboxpermission $Script:txtDispName.Text |Where-Object {$_.User -eq $Script:txtEDGroup.Text})
                        If ($Perm -eq $False)
                        {
				            $LineToWrite = $RecordEvent + "ADD " + "`t" + "Re-Adding ED Group Permission to Mailbox"
				            WriteReportEvent
                            Add-MailboxPermission $Script:txtDispName.Text -User $Script:txtEDGroup.Text -AccessRights FullAccess
				            $LineToWrite = $RecordEvent + "ADD " + "`t" + "Re-Adding SendAs Permission to Mailbox"
				            WriteReportEvent
                            Add-RecipientPermission $Script:txtDispName.Text -Trustee $Script:txtEDGroup.Text -AccessRights SendAs -Confirm:$False
                        }
                    }

                    If ($Script:txtAUGroup.Text -ne "")
                    {
    				    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Recreating the Authors Group" + $Script:txtAUGroup.Text
	    			    WriteReportEvent
                        $DGManagedByMembers = ($Script:txtMbxOwner.Text.Split(",")).Trim()
                        $GrpName = $Script:txtAUGroup.Text
                        $GrpAddr = $Script:txtAUAlias.Text + "@ul.com"
                        $GrpMem = $Script:txtAUMem.Text
                        $Task = $Script:txtTicketNo.Text
                        Recreate-AccessGroups
                    }

                    If ($Script:txtREGroup.Text -ne "")
                    {
    				    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Recreating the Readers Group" + $Script:txtREGroup.Text
	    			    WriteReportEvent
                        $DGManagedByMembers = ($Script:txtMbxOwner.Text.Split(",")).Trim()
                        $GrpName = $Script:txtREGroup.Text
                        $GrpAddr = $Script:txtREAlias.Text + "@ul.com"
                        $GrpMem = $Script:txtREMem.Text
                        $Task = $Script:txtTicketNo.Text
                        Recreate-AccessGroups
                    }

                    $Global:UPN = $Script:txtDispName.Text
                    RetentionPolicy
                }
            }
            else
            {
                 write-host "Restore process cancelled" -ForegroundColor Red
            }
        }While ($Refresh -eq "Yes")
    }
    else
    {
        write-host "Restore process cancelled" -ForegroundColor Red
    }
}