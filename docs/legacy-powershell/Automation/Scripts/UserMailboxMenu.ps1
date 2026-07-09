<####
#### User Mailbox Menu
####
#
#  Called by O365Adminmenu.ps1
#
#  01/04/2018 - SAG  Added Option to purge Organized Calendar Events for an individual.
#  06/11/2018 - SAG  Added Option 8 for configuring Gold/Silver mailbox Access.
#  09/06/2018 - SAG  Added features to look at rules that forward email and allow you to remove them
#  12/13/2018 - SAG  Added "ForwardTo" field to be displayed in the Forwarding Rules
#  08/03/2021 - SAG  Modified to use GUI Menu Form
#>

function Build-UsrMenuForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "User Mailbox Admin Menu"
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 540 ; $form.Height = 400  # Make the form wider 
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 

    $TopLoc = 20 

    ## User mailbox Info
    $TopLoc = $TopLoc + 20
    $Global:chkMbxInfo = New-Object Windows.Forms.RadioButton 
        $Global:chkMbxInfo.Left = 100; $Global:chkMbxInfo.Width = 450; $Global:chkMbxInfo.Top = $TopLoc  
        $Global:chkMbxInfo.Text = "User Mailbox Information" 
        $Global:chkMbxInfo.Checked = $false   # set a default value 
        $Global:chkMbxInfo.TabIndex = 1
        $Global:form.Controls.Add($Global:chkMbxInfo) 
        # Obtain Value with: $Global:chkMbxInfo.Checked
        $Global:InputFocus = $Global:chkMbxInfo
        
    ## Grant Full Permission
    $TopLoc = $TopLoc + 20
    $Global:chkAddPerm = New-Object Windows.Forms.RadioButton 
        $Global:chkAddPerm.Left = 100; $Global:chkAddPerm.Width = 450; $Global:chkAddPerm.Top = $TopLoc
        $Global:chkAddPerm.Text = "Grant User Full Permissions to Another Users Mailbox" 
        $Global:chkAddPerm.Checked = $false   # set a default value 
        $Global:chkAddPerm.TabIndex = 2
        $Global:form.Controls.Add($Global:chkAddPerm) 
        # Obtain Value with: $Global:chkAddPerm.Checked
		
    ## Remove Permission
    $TopLoc = $TopLoc + 20
    $Global:chkRemPerm = New-Object Windows.Forms.RadioButton 
        $Global:chkRemPerm.Left = 100; $Global:chkRemPerm.Width = 450; $Global:chkRemPerm.Top = $TopLoc  
        $Global:chkRemPerm.Text = "Remove User's Full Permissions to another Users Mailbox"
        $Global:chkRemPerm.Checked = $false   # set a default value 
        $Global:chkRemPerm.TabIndex = 3
        $Global:form.Controls.Add($Global:chkRemPerm) 
        # Obtain Value with: $Global:chkRemPerm.Checked	

    ## Remove Forwarder
    $TopLoc = $TopLoc + 20
    $Script:chkRemForw = New-Object Windows.Forms.RadioButton 
        $Script:chkRemForw.Left = 100; $Script:chkRemForw.Width = 450; $Script:chkRemForw.Top = $TopLoc  
        $Script:chkRemForw.Text = "Remove Forwarder from a User Mailbox"
        $Script:chkRemForw.Checked = $false   # set a default value 
        $Script:chkRemForw.TabIndex = 3
        $Global:form.Controls.Add($Script:chkRemForw) 
        # Obtain Value with: $Script:chkRemForw.Checked

    ## Reset Retention Policy
    $TopLoc = $TopLoc + 20
    $Global:chkRetPolicy = New-Object Windows.Forms.RadioButton 
        $Global:chkRetPolicy.Left = 100; $Global:chkRetPolicy.Width = 450; $Global:chkRetPolicy.Top = $TopLoc  
        $Global:chkRetPolicy.Text = "Reset Mailbox Retention Policy"
        $Global:chkRetPolicy.Checked = $false   # set a default value 
        $Global:chkRetPolicy.TabIndex = 3
        $Global:form.Controls.Add($Global:chkRetPolicy) 
        # Obtain Value with: $Global:chkRetPolicy.Checked

    ## View Folder Stats
    $TopLoc = $TopLoc + 20
    $Global:chkViewStat = New-Object Windows.Forms.RadioButton 
        $Global:chkViewStat.Left = 100; $Global:chkViewStat.Width = 450; $Global:chkViewStat.Top = $TopLoc  
        $Global:chkViewStat.Text = "View Mailbox Folder Statistics" 
        $Global:chkViewStat.Checked = $Global:chkViewStat.Checked   # set a default value 
        $Global:chkViewStat.TabIndex = 4
        $Global:form.Controls.Add($Global:chkViewStat) 
        # Obtain Value with: $Global:chkViewStat.Checked

    ## Find Address
    $TopLoc = $TopLoc + 20
    $Global:chkFindObj = New-Object Windows.Forms.RadioButton 
        $Global:chkFindObj.Left = 100; $Global:chkFindObj.Width = 450; $Global:chkFindObj.Top = $TopLoc  
        $Global:chkFindObj.Text = "Find Object Associated with Internet Address" 
        $Global:chkFindObj.Checked = $false   # set a default value 
        $Global:chkFindObj.TabIndex = 5
        $Global:form.Controls.Add($Global:chkFindObj) 
        # Obtain Value with: $Global:chkFindObj.Checked

    ## Cancel Meetings
    $TopLoc = $TopLoc + 20
    $Global:chkCancelMeet = New-Object Windows.Forms.RadioButton 
        $Global:chkCancelMeet.Left = 100; $Global:chkCancelMeet.Width = 450; $Global:chkCancelMeet.Top = $TopLoc  
        $Global:chkCancelMeet.Text = "Cancel All Meetings Organized for Terminated Staff" 
        $Global:chkCancelMeet.Checked = $false   # set a default value 
        $Global:chkCancelMeet.TabIndex = 6
        $Global:form.Controls.Add($Global:chkCancelMeet) 
        # Obtain Value with: $Global:chkCancelMeet.Checked

    ## Gold/Silver Setup
    $TopLoc = $TopLoc + 20
    $Global:chkGoldSilv = New-Object Windows.Forms.RadioButton 
        $Global:chkGoldSilv.Left = 100; $Global:chkGoldSilv.Width = 450; $Global:chkGoldSilv.Top = $TopLoc  
        $Global:chkGoldSilv.Text = "Set-up/Modify Gold/Silver Employee Mailbox Access" 
        $Global:chkGoldSilv.Checked = $false   # set a default value 
        $Global:chkGoldSilv.TabIndex = 7 
        $Global:form.Controls.Add($Global:chkGoldSilv) 
        # Obtain Value with: $Global:chkGoldSilv.Checked

    ## Show Rules
    $TopLoc = $TopLoc + 20
    $Global:chkShRules = New-Object Windows.Forms.RadioButton 
        $Global:chkShRules.Left = 100; $Global:chkShRules.Width = 450; $Global:chkShRules.Top = $TopLoc  
        $Global:chkShRules.Text = "Show Rules On a Mailbox" 
        $Global:chkShRules.Checked = $false   # set a default value 
        $Global:chkShRules.TabIndex = 8
        $Global:form.Controls.Add($Global:chkShRules) 
        # Obtain Value with: $Global:chkShRules.Checked

    Add-FormStandardButtons
}

Do
{
    Build-UsrMenuForm
    Publish-Form
    Check-Reconnect
    
    If ($Global:Result -eq "OK")
    {
        If ($Global:chkMbxInfo.Checked -eq "Checked")
        {
            Enter-EmpNoInputForm
            Publish-Form
    #        write-host "Enter Employee Number or Email Address " -ForegroundColor Yellow -NoNewline
    #        $ENo = Read-Host
            if (($Global:txtInpEmpNo.Text -notlike "*@*") -and ($Global:txtInpEmpNo.Text.Length -gt 6))
            {
                write-host "This is not a valid employee number or email address"
            }
            elseif (($Global:txtInpEmpNo.Text.length -eq 5) -or ($Global:txtInpEmpNo.Text.length -eq 6)) 
            {
                $ENo = $Global:txtInpEmpNo.Text + "@global.ul.com"
            }
            if ($ENo -like "*@*")
            {
                Get-Mailbox $ENo |fl DisplayName,Alias,custom*,*forw*
                write-host ""
            }
            else
            {
                write-host "Mailbox Not Valid: " $ENo
            }
        }

        If ($Global:chkAddPerm.Checked -eq "Checked")
        {
            Enter-EmpNoInputForm
    #        write-host "Enter Employee Number of mailbox to grant access to: " -ForegroundColor Yellow -NoNewline
    #        $ENo = Read-Host
            Get-MailboxPermission $Global:txtInpEmpNo.Text | fl *user*
            Write-Host ""
            Write-Host "Enter Employee Number of the individual to be granted access: " -ForegroundColor Yellow -NoNewline
            $ENoDele = Read-Host
            Write-Host "Granting " $ENoDele " to " $ENo "Mailbox"
            Add-MailboxPermission $Global:txtInpEmpNo.Text -AccessRights FullAccess -User $ENoDele -Automapping:$false
        }

        If ($Global:chkRemPerm.Checked -eq "Checked")
        {
            Enter-EmpNoInputForm
            #write-host "Enter Employee Number of mailbox to remove access from: " -ForegroundColor Yellow -NoNewline
            #$ENo = Read-Host
            Get-MailboxPermission $Global:txtInpEmpNo.Text | fl *user*
            Write-Host ""
            Write-Host "Enter Employee Number of the individual whose permissions will be revoked: " -ForegroundColor Yellow -NoNewline
            $ENoDele = Read-Host
            Write-Host "Revoking " $ENoDele " from " $ENo "Mailbox" -ForegroundColor Red
            Remove-MailboxPermission $ENo -AccessRights FullAccess -User $ENoDele
        }
		
	    If ($Global:chkRetPolicy.Checked -eq "Checked")
        {
            Enter-EmpNoInputForm
			invoke-expression -Command .\O365-ResetRetention.ps1
        }	

        If ($Script:chkRemForw.Checked -eq "Checked")
        {
            $Script:chkRemForw.Checked = $False
            Enter-EmpNoInputForm
            Publish-Form
            If ($Global:Result -eq "OK")
            {
                $Exists = [bool]($mbx=Get-Mailbox $Global:txtInpEmpNo.Text)
                If ($Exists -eq $True)
                {
                    If (($mbx.ForwardingSmtpAddress -ne $null) -or ($mbx.ForwardingAddress -ne $null))
                    {
                        Get-Mailbox $Global:txtInpEmpNo.Text |fl DisplayName,*forwarding*
                        Write-Host "Removing forwarding address from employee" $Global:txtInpEmpNo.Text "Mailbox" -ForegroundColor Green
                        Set-Mailbox $Global:txtInpEmpNo.Text-ForwardingSmtpAddress $Null -ForwardingAddress $null
                    }
                    else
                    {
                        write-host "No forwarding addresses configured on this mailbox" -ForegroundColor Red
                        start-sleep -Seconds 5
                    }
                }
                else
                {
                    write-host "Maibox not found for: $($Global:txtInpEmpNo.Text)" -ForegroundColor Red
                }
            }
            else
            {
                write-host "Removal of forwarding address Cancelled" -ForegroundColor Red
            }
            $Script:chkRemForw.Checked = $False
        }

        If ($Global:chkViewStat.Checked -eq "Checked")
        {
            invoke-expression -Command .\MailboxStatistics.ps1
        }

        If ($Global:chkFindObj.Checked -eq "Checked")
        {
            write-host "This process allows you to find what objects in O365 have a particular internet email address or a portion"
		    write-host "of the internet address associated with it.  For example I can search for all objects with ""us.ul.com"" or"
		    write-host "all objects with ""Sandi"" in it.  Since this command is running against all O365 objects it will take some"
		    write-host "time to display the results.  If you do not wish to proceed hit enter and no command will be executed."
		    write-host
		    write-host "Enter the email address or portion of address you would like to find that match your query " -ForegroundColor Green -NoNewline
		    $FindAddr = Read-Host
		    if ($FindAddr -ne "")
		    {
		        Get-Recipient -ResultSize Unlimited | where {$_.emailaddresses -match $FindAddr} |fl Name,EmailAddresses
		    }
	    }

        If ($Global:chkCancelMeet.Checked -eq "Checked")
        {
	        write-host "This process will cancel all meetings organized by an individual.  This process is executed during the account"
		    write-host "purge process.  If the individual is on Legal Hold we should check to make sure it is OK to purge this information"
		    write-host
		    write-host "Enter the employee number of the individual to Cancel Organized Meetings for: " -ForegroundColor Green -NoNewline
		    $ENo = Read-Host
		    $mbx = get-mailbox $ENo
					
		    if ($ENo.CustomAttribute1 -like "*Ex-*")
		    {
		        $CalEven = Remove-CalendarEvents -Identity $ENo -CancelOrganizedMeetings -QueryWindowInDays 1825 -PreviewOnly -Confirm:$False
    #            $CalEven = Remove-CalendarEvents -Identity $ENo -CancelOrganizedMeetings -PreviewOnly -Confirm:$False
			    if ($CalEven.count -gt 0)
			    {
			        Write-Host "Removing Calendar Event this user is the organizer of: " -ForegroundColor Cyan
				    $CalEven
				    Remove-CalendarEvents -Identity $ENo -CancelOrganizedMeetings -QueryWindowInDays 1825 -Confirm:$False
			    }
			    else
			    {
				    Write-Host "This individual did not Organize Any Calendar Events" -ForegroundColor Red
				    $LineToWrite = $RecordEvent + "INFO" + "`t" + "This individual did not Organize Any Calendar Events" + "`n"
				    WriteReportEvent
			    }
		    }
		    else
		    {
			    write-host "This user is a active user - no meetings will be cancelled!" -ForegroundColor Red
		    }
	    }

        If ($Global:chkGoldSilv.Checked -eq "Checked")
        {
 		    write-host "This process will create the security group or modify the existing membership that grants executive admins to access Gold/Silver employee mailboxes"
		    write-host
		    write-host "Enter the employee number of Gold/Silver Employee: " -ForegroundColor Green -NoNewline
		    $ExecENo = Read-Host
		    $ExecMbxExists = [bool](get-mailbox $ExecENo -ErrorAction SilentlyContinue)
		
            If ($ExecMbxExists -eq $True)
            {
                $ExecMbxName = (get-mailbox $ExecENo).Name + "'s"
                write-host "Create/Update Executive Admin Access to" $ExecMbxName "mailbox"
                $SecGroup = "ACL.UL.FullAccess-" + $ExecENo
                $GrpExits = [bool](Get-DistributionGroup $SecGroup -ErrorAction SilentlyContinue)

                If ($GrpExits -eq $True)
                {
                    write-host "Group for this Executive already exists, below are the individuals who have access to this executives mailbox: "
                    Get-DistributionGroupMember $SecGroup |ft Name
                }

                write-host "Enter the employee number of Executive Admin: " -ForegroundColor Green -NoNewline
			    $AdminENo = Read-Host

                if ($GrpExits -eq $True)
                {
                    write-host "Do you wish to add this admin to the membership of the group (Y/N)? " -ForegroundColor Cyan -NoNewline
                    $UpdGrp = Read-Host
                    if ($updGrp -eq "Y")
                    {
                        write-host "Adding Admin to the Security Group"
                        Add-DistributionGroupMember $SecGroup -Member $AdminENo -BypassSecurityGroupManagerCheck
                    }
                    else
                    {
                        write-host "Removing the Admin from the Security Group"
                        Remove-DistributionGroupMember $SecGroup -Member $AdminENo -BypassSecurityGroupManagerCheck -confirm:$False -ErrorAction SilentlyContinue
                        Get-DistributionGroupMember $SecGroup |ft Name
                    }
                        
                }
                else
                {
                    write-host "Creating group " $SecGroup
                    $SecGroupAddr = $SecGroup + "@ul.onmicrosoft.com"
			        New-DistributionGroup -Name $SecGroup `
				        -PrimarySmtpAddress $SecGroupAddr `
				        -Alias $SecGroup `
				        -ManagedBy $ExecENo `
					    -RequireSenderAuthenticationEnabled $TRUE `
				        -Type Security `
				        | Out-Null

                    Set-DistributionGroup $SecGroup -HiddenFromAddressListsEnabled $True

                    Set-Group -identity $SecGroup `
		                -Notes ("Access to " + $ExecMbxName + " mailbox")

                    Add-DistributionGroupMember $SecGroup -Member $AdminENo -BypassSecurityGroupManagerCheck

                    Write-Host "Applying Mailbox Permissions to executive mailbox: "(get-mailbox $ExecENo).Name -ForegroundColor Cyan
                    Add-MailboxPermission $ExecENo -AccessRights FullAccess -User $SecGroup
                    Add-RecipientPermission $ExecENo -Trustee $SecGroup –AccessRights SendAs -Confirm:$false
                }
            }
        }

        If ($Global:chkShRules.Checked -eq "Checked")
        {
            write-host "Enter the Employee Number or Email Address: " -ForegroundColor Green -NoNewline
            $ENo = Read-Host
            $ENo = $ENo.Trim()

            write-host "All Mailbox Rules: " -ForegroundColor Yellow
            Get-InboxRule -Mailbox $ENo |ft Name,Enabled,Priority,RuleIdentity |Out-Host

            write-host "Do you want to show only rules that forward email to external addresses (Y/N)? " -ForegroundColor Green -NoNewline
            $FwdRules = Read-Host

            If ($FwdRules -eq "Y")
            {
                write-host "All rules that forward: " -ForegroundColor Yellow
                Get-InboxRule -Mailbox $ENo |Where-Object{($_.Description -like "*forward*") -or ($_.Description -like "*redirect*")} |fl Name,Priority,*desc*
                        
                write-host "Rules that have an @ in the ForwardTo field: " -ForegroundColor Yellow
			    $Rules = Get-InboxRule -Mailbox $ENo |Where-Object{((($_.ForwardTo -like "*@*") -or ($_.RedirectTo -like "*@*")) -and (($_.ForwardTo -notlike "*@ul.*") -or ($_.RedirectTo -notlike "*@ul.*")))} |fl Name,Priority,*desc*,ForwardTo
                If ($Rules.count -gt 0)
                {
                    $Rules
                    write-host "Would you like to remove any of these rules (Y/N)?: " -ForegroundColor red -NoNewline
                    $RemRule = read-host
                    Do 
                    {
                        if ($RemRule -eq "Y")
                        {
                            write-host "Enter the Priority of the rule: " -ForegroundColor Cyan -NoNewline
                            $RulePri = Read-Host
                            $DelRule = Get-InboxRule -Mailbox $ENo |Where-Object{$_.Priority -eq $RulePri}

                            write-host "This rule will be removed: "
                            $DelRule |fl Name,Priority,Description

                            write-host "Is this the correct rule (Y/N)?: " -ForegroundColor Cyan -NoNewline
                            $CorrRule = Read-Host
                            If ($CorrRule -eq "Y")
                            {
                                Remove-InboxRule -Identity $DelRule.Identity -Confirm:$False
                                write-host "Remaining Rules that Forward Outside of UL: " -ForegroundColor Yellow
                                $Rule = Get-InboxRule -Mailbox $ENo |Where-Object{(($_.Description -like "*forward*") -or ($_.Description -like "*redirect*")) -and ($_.ForwardTo -like "*@*")} |fl Name,Priority,*desc*
						        If ($Rule.Count -ne 0)
						        {
						            $Rule
							        write-host "Would you like to remove any of these rules (Y/N)?: " -ForegroundColor red -NoNewline
						        }
                            }

                            If ($Rule.Count -ne 0)
					        {
					            write-host "Do you need to remove another rule (Y/N)?: " -ForegroundColor red -NoNewline
						        $RemRule = read-host
				            }
					        else
					        {
						        $RemRule = "N"
					        }
                        }
                    }While ($RemRule -eq "Y")

                    write-host "If there are no rules forwarding externally use the below text in the Resolve field for the Alert received"
                    write-host "Rule found that forwards to another internal UL mailbox no rules that forward to external addresses found." -ForegroundColor Cyan
                }
                else
                {
                    write-host "There are no rules that contain an external forwarding address" -ForegroundColor Red
                }
            }
            else
            {
                write-host "Would you like to see the description on any of these rules (Enter All, Rule Priority or None): " -ForegroundColor Green -NoNewline
                $ViewRule = Read-Host

                If ($ViewRule -eq "All")
                {
                    write-host "Hit Return to see the Description on All Rules"
                    $Ret = Read-Host
                    Get-InboxRule -Mailbox $ENo |fl *desc*
                }
                else
                {
                    if ($ViewRule -ne "None")
                    {
                        Get-InboxRule -Mailbox $ENo |Where-Object{$_.Priority -eq $ViewRule} |fl *desc*
                    }
                }
            }
 
        }
    }
}while ($Global:Result -eq "OK")	