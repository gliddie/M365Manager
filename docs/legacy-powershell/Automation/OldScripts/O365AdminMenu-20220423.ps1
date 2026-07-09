<####################################################################################
#
#  This script contains all the common functions used by the O365 Team
#ad
#  Created: 5/13/2020 - S.Glazebrook
#  06/11/2020 - SAG - Modifed the Common Data Service and Bookings features in the E3EnabledFeatures
#  09/01/2020 - SAG - Modifed Meeting Room to Teams Rooms Standard as name changed at Microsoft
#  11/18/2020 - SAG - Moved the check/removal of the P2 license before the assignment of the E3 license
#  01/06/2021 - SAG - Added the Phone System License as a part of the standard license set
#  07/06/2021 - SAG - Added VisioPlan1 and Visio Plan2 licenses
#  07/15/2021 - SAG - Fixed the PrtMBxAccess function to account for MBX groups having been granted access
#  08/03/2021 - SAG - Added Power Automate per User Plan license
#  10/28/2021 - SAG - Fixed the Project Plan3 AccountSKU
#  04/22/2021 - SAG - Updating for E5 licenses
####################################################################################>

#  Checks the that LogFile directory exists for the given menu item
function CheckLogFiles
{
	# Create Files folder $LogDirectory if it's not present
	if (Test-Path $LogDirectory)
		{
		# the directory is present
		}
	else
		{ 
		mkdir $LogDirectory
		}
} #end CheckLogFiles

#Default form buttons
function Add-FormStandardButtons
{
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    $Global:buttonPanel = New-Object Windows.Forms.Panel  
    $buttonPanel.Size = New-Object Drawing.Size @(400,40) 
    $buttonPanel.Dock = "Bottom"    
    $Global:cancelButton = New-Object Windows.Forms.Button  
        $Global:cancelButton.Top = $buttonPanel.Height - $Global:cancelButton.Height - 10; $Global:cancelButton.Left = $buttonPanel.Width - $Global:cancelButton.Width - 10 
        $Global:cancelButton.TabIndex = 98
        $Global:cancelButton.Text = "Cancel" 
        $Global:cancelButton.DialogResult = "Cancel" 
        $Global:cancelButton.Anchor = "Right"
    ## Create the OK button, which will anchor to the left of Cancel 
    $Global:okButton = New-Object Windows.Forms.Button   
        $Global:okButton.Top = $cancelButton.Top ; $Global:okButton.Left = $cancelButton.Left - $Global:okButton.Width - 10
        $Global:okButton.TabIndex = 97
        $Global:okButton.Text = $Action
        If ($Global:OKDetails -ne "")
        {
            $Global:okButton.Text = $Global:OKDetails
            $Global:OKDetails = ""
        }
        else
        {
            $Global:okButton.Text = "Continue"
        }
        $Global:okButton.DialogResult = "OK" 
        $Global:okButton.Anchor = "Right"
    ## Add the buttons to the button panel 
    $Global:buttonPanel.Controls.Add($Global:okButton) 
    $Global:buttonPanel.Controls.Add($Global:cancelButton) 
    ## Add the button panel to the form 
    $Global:form.Controls.Add($buttonPanel)
    ## Set Default actions for the buttons 
    $Global:form.AcceptButton = $Global:okButton          # ENTER = ok 
    $Global:form.CancelButton = $Global:cancelButton      # ESCAPE = Cancel
}

function AddMembers
{
    #  04/22/20201 - Modified the If statement to be a [bool] test
    If ([bool](Get-DistributionGroup $Global:txtInpName.Text -ErrorAction SilentlyContinue))
    {
        $CurrentMembers = Get-DistributionGroupMember $Global:txtInpName.Text -ResultSize Unlimited
        write-host "Adding Members to Distribution Group: " (Get-DistributionGroupMember $Global:txtInpName.Text).DisplayName
	    $LineToWrite = (Get-DistributionGroupMember $Global:txtInpName.Text).DisplayName + " - Adding Members to Distribution Group."
	    writeReportEvent
					
    	$addMember  = $Global:txtGrpMbr.Text.split(",")
	    if ($? -eq $true)
        {
            ForEach ($mail in $addMember)
            {
                if ($mail.contains("@"))
                {
                    if ($NewMember = Get-Recipient $mail)
                    {
					    if ($CurrentMembers -match $NewMember.Name)
                        {
                            write-host $mail "already a member."
    		    			$LineToWrite = "`t" + "Warn" + "`t" + $mail + "- is already a member of Distribution Group. " + $Global:txtInpName.Text
	    		    		WriteReportEvent
		    		    }
			    		else
                        {
					        Add-DistributionGroupMember $Global:txtInpName.Text -Member $mail
						   	write-host $mail "was added."
    						$LineToWrite = "`t" + "Success" + "`t" + $mail + " - added to Distribution Group. " + $Global:txtInpName.Text
    	    				writeReportEvent
	    	    		}
		    	   	}
		        	else
                    {
                        write-host $mail "- ERROR finding recipient."
					    $LineToWrite = "`t" + "ERROR" + "`t" + $mail + " - ERROR finding recipient. " + $Global:txtInpName.Text
       					WriteReportEvent
        			}
                }
		   	    else
                {
                    write-host "Nothing to add"
          		}
        	}
	   	}
    }
    else
    {
	    Write-Host $Global:txtInpName.Text "- ERROR finding Distribution Group."
	    $LineToWrite = "ERROR" + "`t" + $Global:txtInpName.Text + " - ERROR finding Distribution Group."
        WriteReportEvent
	}

    Write-host "Post Group membership count:  " (Get-DistributionGroupMember -ResultSize Unlimited $Global:txtInpName.Text).Count
    $LineToWrite = "INFO" + "`t" + $Global:txtInpName.Text + "`t" + "Group Membership Updated" + "`n"
    WriteReportEvent
    write-host ""
}

# Apply Mailbox Folder Permissions #Added 6/5/2021 copied from ShrMbxNew.ps1
function Add-FolderPermissions
{
    if ($GrpAddr.contains("@"))
	{
		if ((Get-Mailbox $Global:txtMbxAddr.Text).MailTip -eq $null)
        {
            $NewTip = ("Owners: " + (((Get-DistributionGroup $GrpAddr).Managedby) -join ", "))
            if ($NewTip.Length -gt 175)
            {
                write-host "Maximum mail tip length exceeded truncating to 175 characters." -foregroundcolor Red
                $NewTip = $NewTip.Substring(0,175)
                $OldTip = (Get-Mailbox $Global:txtMbxAddr.Text).MailTip
                write-host "Replacing old MailTip: " $OldTip.Substring(16,$OldTip.Length-36) -foregroundcolor Cyan
                write-host "                 With: " $NewTip
		        $LineToWrite = "REPL  " + "`t" + "Replacing old MailTip with " + $NewTip
		        WriteReportEvent
            }
            Set-Mailbox $Global:txtMbxAddr.Text -MailTip $NewTip
        }

		# 3/23/2020 is the date when all new mailboxes started getting the 3 yr retention policy
        if (((Get-Mailbox $Global:txtMbxAddr.Text).WhenCreated -gt (Get-Date -Year 2020 -Month 3 -Day 23)) -and ((Get-Mailbox $Global:txtMbxAddr.Text).RetentionPolicy -notlike "*3 yr*"))
        {
            Set-MailBox $Global:txtMbxAddr.Text -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL MRM Policy - 3 yr Delete" -RetentionHoldEnabled $false
    	    Set-Mailbox $Global:txtMbxAddr.Text -UseDatabaseQuotaDefaults $False -IssueWarningQuota 45GB -ProhibitSendQuota 49.75GB -ProhibitSendReceiveQuota 50GB
            write-host "3 yr Retention Policy Applied to" $Global:txtMbxAddr.Text
		    $LineToWrite = "UPD  " + "`t" + "Retention Policy Set " + $Global:txtMbxAddr.Text
	        WriteReportEvent
		# Check that IMAP and POP3 are disabled
    	    $ProtocolCheck = (get-CASMailbox $Global:txtMbxAddr.Text) 
    		if ($ProtocolCheck.ImapEnabled -eq $true)
	    	{
		        Set-CASMailBox $Global:txtMbxAddr.Text -ImapEnabled $false
    		    $LineToWrite = "INFO" + "`t" + $Global:txtDispName.Text + "`t" + "IMAP Protocol Disabled"
	    	    WriteReportEvent
		    }
    		if ($ProtocolCheck.PopEnabled -eq $true)
	    	{
		        Set-CASMailBox $Global:txtMbxAddr.Text -PopEnabled $false
    		    $LineToWrite = "INFO" + "`t" + $Global:txtDispName.Text + "`t" + "POP3 Protocol Disabled"
	    	    WriteReportEvent
		    }
        }

        $Permission = ""
        if ($GrpAddr.contains(".ED@"))
        {
            $Permission = "Editor"
       		write-host "Setting permissions on: " $Global:txtMbxAddr.Text " <-- " $GrpAddr " (Editor)" -ForegroundColor Cyan
	    	$LineToWrite = "INFO" + "`t" + "Setting Editor Permissions for " + $Global:txtMbxAddr.Text +  " for " + $GrpAddr
		   	WriteReportEvent
            Write-Host "GrantSendOnBehalf to Delegate Group:  " $GrpAddr -ForegroundColor Cyan
            Set-Mailbox -Identity $Global:txtMbxAddr.Text -GrantSendOnBehalfTo ((Get-Mailbox -Identity $Global:txtMbxAddr.Text).GrantSendOnBehalfTo += $GrpAddr)
            Add-MailboxPermission $Global:txtMbxAddr.Text –User $GrpAddr –AccessRights FullAccess
            Add-RecipientPermission $Global:txtMbxAddr.Text -Trustee $GrpAddr –AccessRights SendAs -Confirm:$false
        }
        elseif ($GrpAddr.contains(".AU@"))
        {
            $Permission = "PublishingAuthor"
            write-host "Setting permissions on: " $Global:txtMbxAddr.Textr " <-- " $GrpAddr " (PublishingAuthor)" -ForegroundColor Cyan
	    	$LineToWrite = "INFO" + "`t" + "Setting Author Permissions for " + $Global:txtMbxAddr.Text +  " for " + $GrpAddr
		   	WriteReportEvent
            Write-Host "GrantSendOnBehalf to Delegate Group:  " $GrpAddr -ForegroundColor Cyan
            Set-Mailbox -Identity $Global:txtMbxAddr.Text -GrantSendOnBehalfTo ((Get-Mailbox -Identity $Global:txtMbxAddr.Text).GrantSendOnBehalfTo += $GrpAddr)
        }
        elseif ($GrpAddr.contains(".RE@"))
        {
            $Permission = "Reviewer"
            write-host "Setting permissions on: " $Global:txtMbxAddr.Text " <-- " $GrpAddr " (Reviewer)" -ForegroundColor Cyan
	    	$LineToWrite = "INFO" + "`t" + "Setting Read Permissions on: " + $Global:txtMbxAddr.Text + " for " + $GrpAddr
		   	WriteReportEvent
        }
        else
        {
            write-host "Group name does not follow naming stand there is no .ED, .AU or .RE found." -ForegroundColor Red
            write-host "Group Name: " $GrpAddr
            $LineToWrite = "INFO" + "`t" + "Group name does not follow naming stand there is no .ED, .AU or .RE found " + $Global:txtMbxAddr.Text
		   	WriteReportEvent
        }
        
        If ($Permission -ne "")
        {
            
            Do
            {
                Start-Sleep -Seconds 2
                $MBXFoldersExist = [bool](Get-MailboxFolderStatistics -Identity $Global:txtMbxAddr.Text -ErrorAction SilentlyContinue)
            } while ($MBXFoldersExist -ne "True")

            $MBXFolders = Get-MailboxFolderStatistics -Identity $Global:txtMbxAddr.Text | % {$_.Identity.ToString().Split("\")[1..100] -join "\"}

            ForEach ($Folder in $MBXFolders)
            {
		        If ($Folder.Equals("Recoverable Items") -or $Folder.Equals("Calendar Logging") -or $Folder.Equals("Deletions") -or $Folder.Equals("Purges") -or $Folder.Equals("Versions"))
                {
				    # Ignore folder
				}
				else
                {
				    if ($Folder.Equals("Top of Information Store"))
                    {
                        Add-MailboxFolderPermission -Identity ($Global:txtMbxAddr.Text + ":\") -User $GrpAddr -AccessRights $Permission
                        $LineToWrite = "ADD  " + "`t" + $Permission + " permission to: " + $Global:txtMbxAddr.Text + ":\"
                    }
					else
                    {
					    Add-MailboxFolderPermission -Identity ($Global:txtMbxAddr.Text + ":\" + $Folder) -User $GrpAddr -AccessRights $Permission
                        $LineToWrite = "ADD  " + "`t" + $Permission + " permission to: " + $Global:txtMbxAddr.Text + ":\" + $Folder
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

<#
function Assign-ATPDefLic
{
    if ($Global:HasATPDef -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Windows Defender Advanced Threat Protection license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Windows Defender Advanced Threat Protection License Already Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Assigning Windows Defender Advanced Threat Protection License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:WIN_DEF_ATP"
    }
    WriteLogEvent
}

function Assign-ATPPDefMACLic
{
    if ($Global:HasATPDefMAC -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Advanced Threat Protection for MAC license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Advanced Threat Protection for MAC License Already Assigned to " + $Global:UPN
    }
    else
    {
        $LineToWrite = "REVI" + "`t" + "Assigning Advanced Threat Protection for MAC License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MDATP_XPLAT"
    }
    WriteLogEvent
}

function Assign-ATPP1Lic
{
    if ($Global:HasATPP1 -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Advanced Threat Protection Plan1 license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Advanced Threat Protection Plan1 License Already Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning Advanced Threat Protection Plan1 License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:ATP_ENTERPRISE"
    }
    WriteLogEvent
}

Function Assign-AudioConferencingLic
{
    if ($Global:HasAudioConf -eq "True")
    {
        $Output = $wshell.Popup("This individual already has an Audio Conferencing license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Audio Conferencing License Already Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning Audio Conferencing License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MCOMEETADV"
    }
    WriteLogEvent
}

Function Assign-CommonAreaPhoneLic
{
    if ($Global:HasCAP -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Common Area Phone license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Common Area Phone License Already Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning Common Area Phone License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MCOCAP"
    }
    WriteLogEvent
}
#>

Function Assign-License
{
    if ($HasLicense -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a " + $LicName + " license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + $LicName + " already assigned to " + $Global:UPN
    }
    else
    {
        $AvailLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq $LicSku}
        If ($AvailLic.ActiveUnits -gt $AvailLic.ConsumedUnits)
        {
           	$LineToWrite = "REVI" + "`t" + "Assigning " + $LicName + " License to " + $Global:UPN
            Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense $LicSKU
        }
        else
        {
           $Output = $wshell.Popup("There are no available " + $LicName + " licenses to assign.",0,"No Licenses Available",0+32)
           $LineToWrite = "REVI" + "`t" + "There are no available " + $LicName + " licenses to assign to " + $Global:UPN
        }
    }
    WriteLogEvent
}

Function UnAssign-License
{
    if ($HasLicense -eq $True)
    {
       	$LineToWrite = "REVI" + "`t" + "Removing  " + $LicName + " license from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense $LicSKU
    }
    else
    {
        $Output = $wshell.Popup("This individual does not have a " + $LicName + " license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No  " + $LicName + " license assigned to " + $Global:UPN
    }
    WriteLogEvent
}

Function Assign-E3Lic
{
    If ($HasExP2 -eq "True")
    {
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:EXCHANGEENTERPRISE"
        write-host "`nRemoving Exchange Online P2 License and reassigning to a " $LicType "License"
        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing P2 License from " + $Global:UPN
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }

    $sskid = "ul:ENTERPRISEPACK"
    $LicType = "Enterprise E3"

    If ($ADUser.ExtensionAttribute1 -like "Employee*")
    {
        $DisPlan = $Global:DisPlanEmpE3
    }
    else
    {
        $DisPlan = $Global:DisPlanNonEmpE3
    }
	
	$LicDet = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"} #This is used to be able to check active/consumed units
	
    If ($HasE3 -eq $True)
    {
        If ($Global:ADUser.ExtensionAttribute1 -like "Employee*")
        {
		    $Output = $wshell.Popup("Resetting " + $LicType + " licenses to standard employee licenses.",0,"Reset Options",0+32)
    	}
	    else
        {
	        $Output = $wshell.Popup("Resetting " + $LicType + " licenses to standard non-employee licenses.",0,"Reset Options",0+32)
		}
        EnabledE3Feature
#        write-host "`nEnterprise E3 License Already Assigned to " $LicType "License"
#        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Enterprise E3 License Already Assigned to " + $Global:UPN
#	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
    else
    {
        If ($ADUser.ExtensionAttribute1 -like "Employee*")
	    {
		    write-host "`nAssigning " $Global:UPN "an " $LicType "License with Employee Features Enabled" -ForegroundColor Yellow
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Assigning " + $LicType + "with Employee Features to " + $Global:UPN
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	    }
	    else
	    {
		    write-host "`nAssigning " $Global:UPN "an " $LicType "License with Non-Employee Features Enabled" -ForegroundColor Yellow
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Assigning " + $LicType + "with Non-Employee Features to " + $Global:UPN
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	    }

        $DLO = ($DisPlan.Split(“,”))
        $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense $SSKID –LicenseOptions $MyO365Sku
        write-host "Waiting for mailbox to be provisioned.." -ForegroundColor Red -NoNewline
        Do{
            Start-Sleep -Seconds 5
            write-host ".." -ForegroundColor Red -NoNewline
        }while ([bool](get-mailbox $Global:UPN -ErrorAction Silentlycontinue) -eq $False)
        write-host ""
        RetentPolicy
    }
}

Function Assign-E5Lic
{
    If ($HasExP2 -eq "True")
    {
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:EXCHANGEENTERPRISE"
        write-host "`nRemoving Exchange Online P2 License and reassigning to a " $LicType "License"
        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing P2 License from " + $Global:UPN
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }    
    
    $E5LicDet = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPREMIUM"} #This is used to be able to check active/consumed units

    If ($E5LicDet.ConsumedUnits -le $E5LicDet.ActiveUnits)
    {
        If ($Global:HasE5 -eq $True)
        {
            $usr = get-msoluser -UserPrincipalName $Global:UPN
            write-host "Getting O365 Licening Group Membership please be patient" -ForegroundColor Red
            $EmpGrp = [bool](Get-AzureADGroupMember -ObjectId 2aab81de-e54a-4c58-909c-b15273c54dff -all $true|Where-Object {$_.ObjectID -eq $usr.ObjectID})
            $NonEmpGrp = [bool](Get-AzureADGroupMember -ObjectId 40dff561-964a-4f7b-9794-dba3542b63c4 -all $true |Where-Object {$_.ObjectID -eq $usr.ObjectID})
            If ($ADUser.ExtensionAttribute1 -like "Employee*")
            {
                #Add/Remove to/from LIC.O365.E5Employee Group
                If ($EmpGrp -eq $False)
                {
                    Add-AzureADGroupMember -ObjectId 2aab81de-e54a-4c58-909c-b15273c54dff -RefObjectId $usr.ObjectID
    		        write-host "`nAdding " $Global:UPN "to Employee Enabled Features Group" -ForegroundColor Yellow
                    $LineToWrite = $RecordEvent + "ADD " + "`t" + "Adding to Employee Enabled Features Group " + $Global:UPN
	                Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                If ($NonEmpGrp -eq $True)
                {
                    Remove-AzureADGroupMember -ObjectId 40dff561-964a-4f7b-9794-dba3542b63c4 -MemberID $usr.ObjectID
    		        write-host "`nRemoving " $Global:UPN "from Non-Employee Enabled Features Group" -ForegroundColor Yellow
                    $LineToWrite = $RecordEvent + "REMO" + "`t" + "Removing from Non-Employee Enabled Features Group " + $Global:UPN
	                Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }
            else
            {
                #Add/Remove to/from LIC.O365.E5NonEmployee Group
                If ($NonEmpGrp -eq $False)
                {
                    Add-AzureADGroupMember -ObjectId 40dff561-964a-4f7b-9794-dba3542b63c4 -RefObjectId $usr.ObjectID
    		        write-host "`nAdding " $Global:UPN "to Non-Employee Enabled Features Group" -ForegroundColor Yellow
                    $LineToWrite = $RecordEvent + "ADD " + "`t" + "Adding to Non-Employee Enabled Features Group " + $Global:UPN
	                Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                If ($EmpGrp -eq $True)
                {
                    Remove-AzureADGroupMember -ObjectId 2aab81de-e54a-4c58-909c-b15273c54dff -MemberID $usr.ObjectID
    		        write-host "`nRemoving " $Global:UPN "from Employee Enabled Features Group" -ForegroundColor Yellow
                    $LineToWrite = $RecordEvent + "REMO" + "`t" + "Removing from Employee Enabled Features Group " + $Global:UPN
	                Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }
        }
        else
        {
            If ($ADUser.ExtensionAttribute1 -like "Employee*")
            {
                Add-AzureADGroupMember -ObjectId 2aab81de-e54a-4c58-909c-b15273c54dff -RefObjectId $usr.ObjectID
 		        write-host "`nAdded " $Global:UPN "to Employee Enabled Features Group" -ForegroundColor Yellow
                $LineToWrite = $RecordEvent + "ADD " + "`t" + "Adding to Employee Enabled Features Group " + $Global:UPN
                Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
            }
            else
            {
                Add-AzureADGroupMember -ObjectId 40dff561-964a-4f7b-9794-dba3542b63c4 -RefObjectId $usr.ObjectID
  		        write-host "`nAdding " $Global:UPN "to Non-Employee Enabled Features Group" -ForegroundColor Yellow
                $LineToWrite = $RecordEvent + "ADD " + "`t" + "Adding to Non-Employee Enabled Features Group " + $Global:UPN
	            Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
            }
        }
<#
        $DLO = ($DisPlan.Split(“,”))
        $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense $SSKID –LicenseOptions $MyO365Sku
        write-host "Waiting for mailbox to be provisioned.." -ForegroundColor Red -NoNewline
        Do{
            Start-Sleep -Seconds 5
            write-host ".." -ForegroundColor Red -NoNewline
        }while ([bool](get-mailbox $Global:UPN -ErrorAction Silentlycontinue) -eq $False)
        write-host ""
#>
        RetentPolicy
    }
    else
    {
        write-host "`nNo Available " $LicType " - no License assigned" -ForegroundColor Yellow
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + "No Available " + $LicType + " - no License assigned to " + $Global:UPN
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
<#	
    If ($HasE5 -eq $True)
    {
        If ($Global:ADUser.ExtensionAttribute1 -like "Employee*")
        {
		    $Output = $wshell.Popup("Resetting " + $LicType + " licenses to standard employee licenses.",0,"Reset Options",0+32)
    	}
	    else
        {
	        $Output = $wshell.Popup("Resetting " + $LicType + " licenses to standard non-employee licenses.",0,"Reset Options",0+32)
		}
        EnabledE5Feature
#        write-host "`nEnterprise E3 License Already Assigned to " $LicType "License"
#        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Enterprise E3 License Already Assigned to " + $Global:UPN
#	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
    else
    {
        If ($ADUser.ExtensionAttribute1 -like "Employee*")
	    {
		    write-host "`nAssigning " $Global:UPN "an " $LicType "License with Employee Features Enabled" -ForegroundColor Yellow
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Assigning " + $LicType + "with Employee Features to " + $Global:UPN
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	    }
	    else
	    {
		    write-host "`nAssigning " $Global:UPN "an " $LicType "License with Non-Employee Features Enabled" -ForegroundColor Yellow
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Assigning " + $LicType + "with Non-Employee Features to " + $Global:UPN
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	    }

        $DLO = ($DisPlan.Split(“,”))
        $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense $SSKID –LicenseOptions $MyO365Sku
        write-host "Waiting for mailbox to be provisioned.." -ForegroundColor Red -NoNewline
        Do{
            Start-Sleep -Seconds 5
            write-host ".." -ForegroundColor Red -NoNewline
        }while ([bool](get-mailbox $Global:UPN -ErrorAction Silentlycontinue) -eq $False)
        write-host ""
        RetentPolicy
    }
#>
}


function Assign-P2Lic
{
    $Lic = $Global:UserLicense = (Get-MsolUser -UserPrincipalName $Global:UPN).Licenses
    if ($Global:HasExP2 -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Exchange Online P2 license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Exchange Online P2 License Already Assigned to " + $Global:UPN
    }
    else
    {
        If ($Global:UsageLoc -eq "US")
        {
            $LineToWrite = "REVI" + "`t" + "Assigning Exchange Online P2 License to " + $Global:UPN
            Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:EXCHANGEENTERPRISE"
        }
        else
        {
            write-host "No license was ever assigned to this account"
            $LineToWrite = "REVI" + "`t" + "Usage Location Not Set and therefore no license was ever assigned to " + $Global:UPN
        }
    }
    WriteLogEvent

    If ($Lic.count -ge 1)
    {
        $Output = $wshell.Popup("This individual has " + $Global:LicAssigned + " licenses assigned do you want to remove all assigned licenses from this individual?",4,"Remove Standrd Licenses",4+32)
        If ($Output -eq 6)
        {
            Foreach ($Lic in $Lic)
            {
                Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense $Lic.AccountSkuID
                $LineToWrite = "REVI"  + "`t" + "Removing " + $Lic.AccountSkuID + " from " + $Global:UPN
                WriteLogEvent
            }
        }
        else
        {
            If ($Global:HasE3 -eq "True")
            {
                $LineToWrite = "REVI"  + "`t" + "Removing EnterpriseE3 License from " + $Global:UPN
                WriteLogEvent
                Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:ENTERPRISEPACK"
            }
        }
    }
}

function Assign-PowerBIFLic
{
    if ($Global:HasBIFree -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a PowerBI Free license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "PowerBI Free License Already Assigned to " + $Global:UPN
        WriteLogEvent
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning PowerBI Free License to " + $Global:UPN
        WriteLogEvent
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:POWER_BI_STANDARD"

        If ($Global:HasBIPro -eq "True")
        {
        	$LineToWrite = "REVI" + "`t" + "Removing PowerBI Pro License from" + $Global:UPN
	      	WriteLogEvent
            Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:POWER_BI_PRO"
        }
    }  
}
<#
function Assign-PowerBIProLic
{
    if ($Global:HasBIPro -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a PowerBI Pro license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "PowerBI Pro License Already Assigned to " + $Global:UPN
        WriteLogEvent
        If ($Global:HasBIFree -eq $True)
        {
        	$LineToWrite = "REVI" + "`t" + "Removing PowerBI (free) License from" + $Global:UPN
	      	WriteLogEvent
            Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:POWER_BI_STANDARD"
        }
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning PowerBI Pro License to " + $Global:UPN
        WriteLogEvent
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:POWER_BI_PRO"
        If ($Global:HasBIFree -eq $True)
        {
        	$LineToWrite = "REVI" + "`t" + "Removing PowerBI (free) License from" + $Global:UPN
	      	WriteLogEvent
            Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:POWER_BI_STANDARD"
        }
    }  
}

function Assign-EMSLic
{
    If ($Global:HasEMS -eq "True")
	{
        $Output = $wshell.Popup("This individual already has a EMS license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Enterprise Mobility Suite/Intune (EMS) License Already Assigned to " + $Global:UPN
    }
	else
    {
		Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:EMS"
        $LineToWrite = "REVI" + "`t" + "Assigning Enterprise Mobility Suite/Intune (EMS) License to " + $Global:UPN
    }
    WriteLogEvent
}

function Assign-MeetingLic
{
    $LicName = "Microsoft Meeting Room"
    if ($Global:HasMeeting -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a " + $LicName + " license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + $LicName + " License Already Assigned to " + $Global:UPN
    }
    else
    {
        $AvailLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MEETING_ROOM"}
        If ($AvailLic.ActiveUnits -gt $AvailLic.ConsumedUnits)
        {
           	$LineToWrite = "REVI" + "`t" + "Assigning " + $LicName + " License to " + $Global:UPN
            Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MEETING_ROOM"
        }
        else
        {
           $Output = $wshell.Popup("There are no " + $LicName + " Licenses Available to Assign.",0,"No Licenses Available",0+32)
           $LineToWrite = "REVI" + "`t" + "There are no " + $LicName + " Licenses Available to Assign to " + $Global:UPN
        }
    }
    WriteLogEvent
}


Function Assign-PAppsP2Lic
{
    if ($Global:HasPAppsP2 -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a PowerApps P2 license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "PowerApps Plan 2 License Already Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Assigning PowerApps Plan 2 License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:POWERFLOW_P2"
    }
    WriteLogEvent
}

Function Assign-AppsPUsrLic
{
    if ($Global:HasPAppsP2 -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a PowerApps p/User license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "PowerApps p/User License Already Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Assigning PowerApps p/User License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:POWERAPPS_PER_USER"
    }
    WriteLogEvent
}

Function Assign-AutoPUsrLic
{
    if ($Global:HasAutoPUsr -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Power Automate per User Plan license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Power Automate per User Plan License Already Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Assigning Power Automate per User Plan License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:POWERAUTOMATE_ATTENDED_RPA"
    }
    WriteLogEvent
}

Function Assign-AutoPUsrRPALic
{
    if ($Global:HasAutoPUsrRPA -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Power Automate per User w/RPA Plan license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Power Automate per User w/RPA Plan License Already Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Assigning Power Automate per User w/RPA Plan License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:FLOW_PER_USER"
    }
    WriteLogEvent
}

Function Assign-PhoneSystemLic
{
    $LicName = "Phone System"
    if ($Global:HasPhone -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a " + $LicName + " license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + $LicName + " License Already Assigned to " + $Global:UPN
    }
    else
    {
        $AvailLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOEV"}
        If ($AvailLic.ActiveUnits -gt $AvailLic.ConsumedUnits)
        {
     	    $LineToWrite = "REVI" + "`t" + "Assigning " + $LicName + " License to " + $Global:UPN
            Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MCOEV"
        }
        else
        {
           $Output = $wshell.Popup("There are no " + $LicName + " Licenses Available to assign.",0,"No Licenses Available",0+32)
           $LineToWrite = "REVI" + "`t" + "There are no " + $LicName + " Licenses Available to assign to " + $Global:UPN
        }
    }
    WriteLogEvent
}

function Assign-ProjP3Lic
{
    if ($Global:HasProjP3 -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Project Plan 3 license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Project Plan 3 License Already Assigned to " + $Global:UPN
        WriteLogEvent
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning Project Plan 3 License to " + $Global:UPN
        WriteLogEvent
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:PROJECTPROFESSIONAL"
    }  
}

function Assign-ProjP5Lic
{
    if ($Global:HasProjP5 -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Project Plan 5 license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Project Plan 5 License Already Assigned to " + $Global:UPN
        WriteLogEvent
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning Project Plan 5 License to " + $Global:UPN
        WriteLogEvent
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:PROJECTPREMIUM"
    }  
}

function Assign-RemoteAsstLic
{
    if ($Global:HasRmtAssist -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Microsoft Remote Assist license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Microsoft Remote Assistg License Already Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning Microsoft Remote Assist License to " + $Global:UPN
        If ($null -eq (get-msoluser -UserPrincipalName $global:UPN).UsageLocation)
        {
            Set-MsolUser -UserPrincipalName $Global:UPN -UsageLocation "US" -ErrorAction Silentlycontinue
        }
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MICROSOFT_REMOTE_ASSIST"
    }
    WriteLogEvent
    If ($ADuser.ExtensionAttribute1 -like "Employee*")
    {
        Assign-AudioConferencingLic
        Assign-PhoneSystemLic
    }
}

function Assign-VisioP1Lic
{
    If ($Global:HasVisioP1 -eq "True")
    {
		$Output = $wshell.Popup("This individual already has a Visio Plan 1 license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Visio Plan 1 License Already Assigned to " + $Global:UPN
    }
    else
    {
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:VISIOONLINE_PLAN1"
        $LineToWrite = "REVI" + "`t" + "Assigning Visio Plan 1 License to " + $Global:UPN    }
    WriteLogEvent
}

function Assign-VisioP2Lic
{
    If ($Global:HasVisioP2 -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Visio Plan 2 license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Visio Plan 2 License Already Assigned to " + $Global:UPN
    }
    else
    {
		Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:VISIOCLIENT"
        $LineToWrite = "DELE" + "`t" + "Assigning Visio Plan 2 License to " + $Global:UPN
    }
    WriteLogEvent
}

#>

<#Function Check-Reconnect
{
#    if (($stopwatch.elapsed -ge $reconnectThreshold) -and ($O365Act -ne 0))
#write-host "Elapsed Time = " $Script:StopWatch.elapsed
#write-host "Threshold = " $Script:reconnectThreshold
    if ($Script:stopwatch.elapsed -ge $Script:reconnectThreshold)
    {
        # Close all sessions
        write-host "Connection Threshold Exceeded -- Reconnecting to Office365" -ForegroundColor Red
        get-pssession | Remove-PSSession -Confirm:$false
        invoke-expression -Command E:\O365AdminShared\Scripts\ConnectO365.ps1
        $Script:stopwatch = [diagnostics.stopwatch]::StartNew()
    }    
}
#>
#This function creates New UserGroups
#   Used by NewSharedMailboxScripts
function CreateUSG
{				
    $DstExists = [bool](Get-DistributionGroup $GrpAddr -ResultSize Unlimited -ErrorAction SilentlyContinue)
    if ($DstExists -eq $false)
    {
	    if ($GrpAddr.contains("@"))
        {
            $atMail = $GrpAddr.indexOf("@")
		    $DGAlias = $GrpAddr.substring(0,$atMail)
            $DGManagedByMembers = ($Global:txtGrpOwnr.Text.Split(",")).Trim()
            $Task = $Global:txtInpTaskNo.Text
            #$DGManagedByMembers = ($Global:txtEDGrpMbr.Text.Split(",")).Trim()

            $LineToWrite = ""
            WriteReportEvent
        			
		    # Create the USG
		    New-DistributionGroup -Name $GrpName `
		        -PrimarySmtpAddress $GrpAddr `
			    -Alias $DGAlias `
			    -ManagedBy $DGManagedByMembers `
			    -RequireSenderAuthenticationEnabled $TRUE `
			    -Type Security | Out-Null
					
            $Owners = "Owners: " + ((Get-DistributionGroup $GrpName).ManagedBy -join (", "))
        
            Set-Group -identity $GrpName `
		        -Notes ($Owners + " - Per: " + $Task)

            If ($Owners.Length -gt 175)
            {
                Set-DistributionGroup $GrpName -MailTip $Owners.Substring(0,175)
            }
            else
            {
                Set-DistributionGroup $GrpName -MailTip $Owners
            }
        
            if (Get-DistributionGroup $GrpAddr)
            {
		        write-host ""
                write-host "USG created: " $GrpName " (" $GrpAddr ")" -ForegroundColor Cyan
		        $LineToWrite = "CREATE" + "`t" + $GrpName + "`t" + $GrpAddr + " USG Created"
		        WriteReportEvent
			
			    write-host "USG Ownner: " $Global:txtGrpOwnr.Text -ForegroundColor Cyan
		        $LineToWrite = "OWNER" + "`t" + $Global:txtGrpOwnr.Text + " USG Owner"
		        WriteReportEvent
				
		        $USG = Get-DistributionGroup $GrpAddr
		        Set-DistributionGroup $GrpAddr `
			        -RequireSenderAuthenticationEnabled $True `
		 		    -BypassSecurityGroupManagerCheck `
				    -CustomAttribute15 ($FileName + " PS Date: " + (get-date -uformat %D) + " PS Time: " + (get-date -uformat %T))
				
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
		    else
            {
		        Write-Host "USG not created: " $GrpName " (" $GrpAddr ")" -ForegroundColor Red
			    $LineToWrite = "FAIL" + "`t" + $GrpName + "`t" + $GrpAddr + "`t" + "USG not created"
			    WriteReportEvent
	        }
        }
    }
    else
    {
        Write-Host "User Security Group Already Exists - No Changes Made" -ForegroundColor Red
        pause
    }
}

Function EnabledE3Feature
{
    write-host "`nDetails for" $Global:UserDet -ForegroundColor Cyan

    $SSKID = "ul:ENTERPRISEPACK"
    $LicType = "Enterprise E3"

    $text = "`t{0}`t`t{1}`t`t`t{2}" -f "DefaultAssignment","Status ","SubLicenseName"
    $text
    $text = "`t{0}`t`t{1}`t`t`t{2}" -f "-----------------","------ ","--------------"
    $text

    $LineToWrite = $RecordEvent + "UPDA" + "`t" + $LicType + " SubLicenses Enabled/Status for " + $Global:UPN
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    $SubLic = ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq $SSKID}).ServiceStatus
	foreach ($SubLic in $SubLic)
	{
        $AssignWho = "Enabled All Types"

        switch ($subLic.ServicePlan.ServiceName)
        {
            "VIVA_LEARNING_SEEDED"
            {
                $LicName = "Viva Learning Seeded"
            }
            "Nucleus"
            {
                $LicName = "Nucleus"
            }              
            "ContentExplorer_Standard"
            {
                $LicName = "Information Protection and Governance Analytics - Standard"
            }
            "POWER_VIRTUAL_AGENTS_O365_P2"
            {
                $LicName = "Power Virtual Agents"
                $AssignWho = "Disabled by Default"
            }
            "CDS_O365_P2"
            {
                $LicName = "Common Data Service for Teams"
                $AssignWho = "Disabled by Default"
            }
            "PROJECT_O365_P2"
            {
                $LicName = "Project for Office (Plan E3)"
            }
            "DYN365_CDS_O365_P2"
            {
                $LicName = "Common Data Service"
                $AssignWho = "Disabled by Default"
            }
            "MICROSOFTBOOKINGS"
            {
                $LicName = "Microsoft Bookings"
            }
            "KAIZALA_O365_P3"
            {
                $LicName = "Microsoft Kaizala Pro"
            }
            "MICROSOFT_SEARCH"
            {
                $LicName = "Microsoft Search"
            }
            "WHITEBOARD_PLAN2"
            {
                $LicName = "Whiteboard (Plan 2)"
            }
            "MIP_S_CLP1"
            {
                $LicName = "Information Protection for Office 365 - Standard"
            }
            "MYANALYTICS_P2"
            {
                $LicName = "Insights by MyAnalytics"
            }
            "BPOS_S_TODO_2"
            {
                $LicName = "To-Do (Plan 2)"
            }
            "FORMS_PLAN_E3"
            {
                $LicName = "Microsoft Forms (Plan E3)"
            }
            "STREAM_O365_E3"
            {
                $LicName = "Microsoft Stream for O365 E3 SKU"
            }
            "Deskless"
            {
                $LicName = "Microsoft StaffHub"
                $AssignWho = "Disabled by Default"
            }
            "FLOW_O365_P2"
            {
                $LicName = "Power Automate for O365 "
            }
            "POWERAPPS_O365_P2"
            {
                $LicName = "PowerApps for Office 365"
            }
            "TEAMS1"
            {
                $LicName = "Microsoft Teams"
            }
            "PROJECTWORKMANAGEMENT"
            {
                $LicName = "Microsoft Planner"
            }
            "SWAY"
            {
                $LicName = "Sway"
            }
            "INTUNE_O365"
            {
                $LicName = "Mobile Device Management (MDM)"
            }
            "YAMMER_ENTERPRISE"
            {
                $LicName = "Yammer Enterprise"
                $AssignWho = "Enabled Empl Only"
            }
            "RMS_S_ENTERPRISE"
            {
                $LicName = "Azure Rights Management"
            }
            "OFFICESUBSCRIPTION"
            {
                $LicName = "M365 Apps for Enterprise"
            }
            "MCOSTANDARD"
            {
                $LicName = "Skype for Business Online (Plan 2)"
                $AssignWho = "Enabled Empl Only"
            }
            "SHAREPOINTWAC"
            {
                $LicName = "Office for the web"
                $AssignWho = "Enabled Empl Only"
            }
            "SHAREPOINTENTERPRISE"
            {
                $LicName = "SharePoint Online (Plan 2)"
                $AssignWho = "Enabled Empl Only"
            }
            "EXCHANGE_S_ENTERPRISE"
            {
                $LicName = "Exchange Online (Plan 2)"
            }
        }

        $LicStat = $SubLic.ProvisioningStatus
        If ($LicStat -like "*Pending*")
        {
            $LicStat = "Pending "
        }
        If ($LicStat -eq "Success")
        {
            $LicStat = "Enabled "
        }

        $text = "`t{0}`t`t{1}`t`t{2}" -f $AssignWho,$LicStat,$LicName
        $text
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + $text
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}
}

Function EnabledE5Feature
{
    write-host "`nDetails for" $Global:UserDet -ForegroundColor Cyan

    $SSKID = "ul:ENTERPRISEPREMIUM"
    $LicType = "Enterprise E5"

    $text = "`t{0}`t`t{1}`t`t`t{2}" -f "DefaultAssignment","Status ","SubLicenseName"
    $text
    $text = "`t{0}`t`t{1}`t`t`t{2}" -f "-----------------","------ ","--------------"
    $text

    $LineToWrite = $RecordEvent + "UPDA" + "`t" + $LicType + " SubLicenses Enabled/Status for " + $Global:UPN
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    $SubLic = ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq $SSKID}).ServiceStatus
	foreach ($SubLic in $SubLic)
	{
        $AssignWho = "Enabled All Types"

        switch ($subLic.ServicePlan.ServiceName)
        {
             "RMS_S_ENTERPRISE"
            {
                $LicName = "Azure Rights Management"
            }
            "DYN365_CDS_O365_P3"
            {
                $LicName = "Common Data Service"
                $AssignWho = "Disabled by Default"
            }
            "CDS_O365_P3"
            {
                $LicName = "Common Data Service for Teams"
                $AssignWho = "Disabled by Default"
            }
			"LOCKBOX_ENTERPRISE"
		    {
                $LicName = "Customer Lockbox"
                $AssignWho = "Disabled by Default"
            }
			"MIP_S_Exchange"
			{
                $LicName = "Data Classification in M365"
                $AssignWho = "Disabled by Default"	
			}			
            "EXCHANGE_S_ENTERPRISE"
            {
                $LicName = "Exchange Online (Plan 2)"
            }			
			"GRAPH_CONNECTORS_SEARCH_INDEX"
			{
                $LicName = "Graph Connectors Search wth Index"
			}
			"INFORMATION_BARRIERS"
            {
                $LicName = "Informaton Barriers"
                $AssignWho = "Disabled by Default"
            }
			"Content_Explorer"
			{
                $LicName = "Information Protection and Governance Analytics - Premium"
				$AssignWho = "Managed at Org Level"
			}
			"ContentExplorer_Standard"
            {
                $LicName = "Information Protection and Governance Analytics - Standard"
				$AssignWho = "Managed at Org Level"
            }			
            "MIP_S_CLP2"
            {
                $LicName = "Information Protection for O365 - Premium"
				$AssignWho = "Disabled by Default"
            }
            "MIP_S_CLP1"
            {
                $LicName = "Information Protection for O365 - Standard"
            }
            "MYANALYTICS_P2"
            {
                $LicName = "Insights by MyAnalytics"
            }			
			"M365_ADVANCED_AUDITING"
            {
                $LicName = "M365 Advanced Auditing"
                $AssignWho = "Disabled by Default"
            }			
            "OFFICESUBSCRIPTION"
            {
                $LicName = "M365 Apps for Enterprise"
            }
			"MCOMEETADV"
		    {
                $LicName = "M365 Audio Conferencing"
                $AssignWho = "Disabled by Default"
            }			
			"MICROSOFT_COMMUNICATION_COMPLIANCE"
			{
                $LicName = "M365 Communication Compliance"
                $AssignWho = "Disabled by Default"	
			}
			"MTP"
			{
				$LicName = "M365 Defender"
			}
			"MCOEV"
            {
                $LicName = "M365 Phone System"
            }
            "MICROSOFTBOOKINGS"
            {
                $LicName = "Microsoft Bookings"
            }
			"COMMUNICATIONS_DLP"
            {
                $LicName = "Communications DLP"
                $AssignWho = "Disabled by Default"
            }
			"CUSTOMER_KEY"
            {
                $LicName = "Customer Key"
                $AssignWho = "Disabled by Default"
            }
			"DATA_INVESTIGATIONS"
            {
                $LicName = "Data Investigations"
                $AssignWho = "Disabled by Default"
            }
            "ATP_ENTERPRISE"
            {
                $LicName = "Defender for O365 (Plan 1)"
				$AssignWho = "Managed at Org Level"
            }
			"THREAT_INTELLIGENCE"
            {
                $LicName = "Defender for O365 (Plan 2)"
				$AssignWho = "Disabled by Default"
            }
			"EXCEL_PREMIUM"
			{
				$LicName = "Excel Advanced Analytics"
				$AssignWho = "Disabled by Default"
			}
            "FORMS_PLAN_E5"
            {
                $LicName = "Microsoft Forms(Plan E5)"
            }
			"INFO_GOVERNANCE"
            {
                $LicName = "Information Governance"
                $AssignWho = "Disabled by Default"
            }
            "KAIZALA_STANDALONE"
            {
                $LicName = "Microsoft Kaizala Pro"
            }
			"EXCHANGE_ANALYTICS"
		    {
                $LicName = "MyAnalytics (Full)"
                $AssignWho = "Disabled by Default"
            }			
            "PROJECTWORKMANAGEMENT"
            {
                $LicName = "Microsoft Planner"
            }
			"RECORDS_MANAGEMENT"
            {
                $LicName = "Records Management"
                $AssignWho = "Disabled by Default"
            }
            "MICROSOFT_SEARCH"
            {
                $LicName = "Microsoft Search"
				$AssignWho = "Managed at Org Level"
            }
            "Deskless"
            {
                $LicName = "Microsoft StaffHub"
                $AssignWho = "Disabled by Default"
            }			
            "STREAM_O365_E5"
            {
                $LicName = "Microsoft Stream for O365 E5"
            }
            "TEAMS1"
            {
                $LicName = "Microsoft Teams"
            }
            "INTUNE_O365"
            {
                $LicName = "Mobile Device Management (MDM)"
				$AssignWho = "Managed at Org Level"
            }			
            "Nucleus"
            {
                $LicName = "Nucleus"
				$AssignWho = "Managed at Org Level"
			}
			"EQUIVIO_ANALYTICS"
		    {
                $LicName = "O365 Advanced Discovery"
                $AssignWho = "Disabled by Default"
            }	
			"ADALLOM_S_O365"
            {
                $LicName = "O365 Cloud App Security"
                $AssignWho = "Disabled by Default"
            }
			"PAM_ENTERPRISE"
            {
                $LicName = "Priviledged Access Management"
				$AssignWho = "Disabled by Default"
            }
            "SHAREPOINTWAC"
            {
                $LicName = "Office for the web"
                $AssignWho = "Enabled Empl Only"
            }
            "POWERAPPS_O365_P3"
            {
                $LicName = "PowerApps for O365 (Plan 3)"
            }
            "FLOW_O365_P3"
            {
                $LicName = "Power Automate for O365 "
            }
			"BI_AZURE_P2"
			{
                $LicName = "Power BI Pro"
                $AssignWho = "Disabled by Default"
            }
            "POWER_VIRTUAL_AGENTS_O365_P3"
            {
                $LicName = "Power Virtual Agents for O365"
                $AssignWho = "Disabled by Default"
            }
			"PREMIUM_ENCRYPTION"
			{
				$LicName = "Premium Encryption for O365"
				$AssignWho = "Disabled by Default"
			}
            "PROJECT_O365_P3"
            {
                $LicName = "Project for Office (Plan E5)"
            }
			"COMMUNICATIONS_COMPLIANCE"
            {
                $LicName = "RETIRED - Microsoft Communication Compliance"
                $AssignWho = "Disabled by Default"
            }
            "SHAREPOINTENTERPRISE"
            {
                $LicName = "SharePoint Online (Plan 2)"
                $AssignWho = "Enabled Empl Only"
            }
            "MCOSTANDARD"
            {
                $LicName = "Skype for Business Online (Plan 2)"
                $AssignWho = "Enabled Empl Only"
            }
            "SWAY"
            {
                $LicName = "Sway"
            }
            "BPOS_S_TODO_3"
            {
                $LicName = "To-Do (Plan 3)"
            }
            "VIVA_LEARNING_SEEDED"
            {
                $LicName = "Viva Learning Seeded"
            }
            "WHITEBOARD_PLAN3"
            {
                $LicName = "Whiteboard (Plan 3)"
            }
            "YAMMER_ENTERPRISE"
            {
                $LicName = "Yammer Enterprise"
                $AssignWho = "Enabled Empl Only"
            }
        }

        $LicStat = $SubLic.ProvisioningStatus
        If ($LicStat -like "*Pending*")
        {
            $LicStat = "Pending "
        }
        If ($LicStat -eq "Success")
        {
            $LicStat = "Enabled "
        }

        $text = "`t{0}`t`t{1}`t`t{2}" -f $AssignWho,$LicStat,$LicName
        $text
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + $text
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}
}

Function E5NonStandardFeatures
{
    $SubLic = ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:ENTERPRISEPREMIUM"}).ServiceStatus
	foreach ($SubLic in $SubLic)
	{
        switch ($subLic.ServicePlan.ServiceName)
        {
			"BI_AZURE_P2"
			{
                If ($subLic.ProvisioningStatus -eq "Success")
                {
                    If ($text -like "*w/*")
                    {
                        $text = $text + ","
                    }
                    $text = $text + " (w/PowerBI Pro)"
                }
            }
			"MCOEV"
            {
                If ($subLic.ProvisioningStatus -eq "Success")
                {
                    If ($text -like "*w/*")
                    {
                        $text = $text + ","
                    }
                    $text = $text + " (w/Phone System)"
                }
            }
			"MCOMEETADV"
		    {
                If ($subLic.ProvisioningStatus -eq "Success")
                {
                    If ($text -like "*w/*")
                    {
                        $text = $text + ","
                    }
                    $text = $text + " (w/Audio Conferencing)"
                }
            }
        }
 	}
    write-host $text
}

Function EnabledMRSSFeature
{
    write-host "`n`tEnabled SubLicense Options:" -ForegroundColor Magenta

    $text = "`t{0}`t`t{1}`t`t`t{2}" -f "DefaultAssignment","Status ","SubLicenseName"
    $text
    $text = "`t{0}`t`t{1}`t`t`t{2}" -f "-----------------","------ ","--------------"
    $text

    $LineToWrite = $RecordEvent + "UPDA" + "`t" + $LicType + " SubLicenses Enabled/Status for " + $Global:UPN
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    $SubLic = ($UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq $SSKID}).ServiceStatus
	foreach ($SubLic in $SubLic)
	{
        $AssignWho = "Disabled All Types"

        switch ($subLic.ServicePlan.ServiceName)
        {
           "EXCHANGE_S_FOUNDATION"
           {
               $LicName = "No Details in O365 Portal"
           }
            "SHAREPOINTENTERPRISE"
            {
                $LicName = "SharePoint Online (Plan 2)"
            }
            "PROJECT_ESSENTIALS"
            {
                $LicName = "Project Online Essentials"
            }
            "POWERAPPS_DYN_APPS"
            {
                $LicName = "PowerApps for Dynamics 365"
            }
            "SHAREPOINTWAC"
            {
                $LicName = "Office Online"
            }
            "NBENTERPRISE"
            {
                $LicName = "Microsoft Social Engagement Enterprise"
            }
            "FLOW_DYN_APPS"
            {
                $LicName = "Flow for Dynamics 365"
            }
            "DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
            {
                $LicName = "Microsoft Relationship Sales solution"
                $AssignWho = "Enabled by Default"
            }
        }

        $LicStat = $SubLic.ProvisioningStatus
        If (($LicStat -eq "PendingActivation") -or ($LicStat -eq "PendingProvisioning"))
        {
            $LicStat = "Pending "
        }
        If ($LicStat -eq "Success")
        {
            $LicStat = "Enabled "
        }
        $text = "`t{0}`t`t{1}`t`t{2}" -f $AssignWho,$LicStat,$LicName
        $text
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + $text
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}
}

function Enter-DLNameInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow"
    $Global:form.Text = "Distribution List Name"
    $Global:form.Size = New-Object System.Drawing.Size(450,175) #(W,H)
    $Global:form.StartPosition = "CenterScreen"
    Add-FormStandardButtons
    ## Employee Number
    $Global:lblDLName = New-Object System.Windows.Forms.Label   
        $Global:lblDLName.Text = "Distribution List Name:"  
        $Global:lblDLName.Top = 30 ; $Global:lblDLName.Left = 10; $Global:lblDLName.Width=120 ; $Global:lblDLName.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLName)    # Add to Form 
        # 
        $Global:txtInpDLName = New-Object Windows.Forms.TextBox  
        $Global:txtInpDLName.Top = 30; $Global:txtInpDLName.Left = 140; $Global:txtInpDLName.Width = 200;  
        $Global:txtInpDLName.Text = ""
        $Global:form.Controls.Add($Global:txtInpDLName)    # Add to Form
        $Global:InputFocus = $Global:txtInpDLName
}

function Enter-EmpNoInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow"
    $Global:form.Text = "Employee No."
    $Global:form.Size = New-Object System.Drawing.Size(450,175) #(W,H)
    $Global:form.StartPosition = "CenterScreen"
    Add-FormStandardButtons
    ## Employee Number
    $Global:lblEmpNo = New-Object System.Windows.Forms.Label   
        $Global:lblEmpNo.Text = "Employee No.:"  
        $Global:lblEmpNo.Top = 30 ; $Global:lblEmpNo.Left = 10; $Global:lblEmpNo.Width=120 ; $Global:lblEmpNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblEmpNo)    # Add to Form 
        # 
        $Global:txtInpEmpNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpEmpNo.Top = 30; $Global:txtInpEmpNo.Left = 140; $Global:txtInpEmpNo.Width = 200;  
        $Global:txtInpEmpNo.Text = ""
        $Global:form.Controls.Add($Global:txtInpEmpNo)    # Add to Form
        $Global:InputFocus = $Global:txtInpEmpNo

    If (($Global:chkPurgeAct.Checked -eq "Checked") -or ($Global:chkTermination.Checked -eq "Checked") -or ($Global:chkEnabUsr.Checked -eq "Checked"))
    {
        ## Ticket Number
        $Global:lblTicketNo = New-Object System.Windows.Forms.Label   
            $Global:lblTicketNo.Text = "Ticket No.:"  
            $Global:lblTicketNo.Top = 60 ; $Global:lblTicketNo.Left = 10; $Global:lblTicketNo.Width=120 ; $Global:lblTicketNo.AutoSize = $true
            $Global:form.Controls.Add($Global:lblTicketNo)    # Add to Form 
            # 
            $Global:txtInpTicketNo = New-Object Windows.Forms.TextBox  
            $Global:txtInpTicketNo.Top = 60; $Global:txtInpTicketNo.Left = 140; $Global:txtInpTicketNo.Width = 200;  
            $Global:txtInpTicketNo.Text = ""
            $Global:form.Controls.Add($Global:txtInpTicketNo)    # Add to Form
    }

    If ($Global:chkEnabUsr.Checked -eq "Checked")
    {
        $Global:form.Height = $Global:form.Height + 30
        # StartDate
        $Global:lblDatePicker = New-Object System.Windows.Forms.Label
        $Global:lblDatePicker.Text = "Start Date:"
        $Global:lblDatePicker.Top = 90 ; $Global:lblDatePicker.Left = 10; $Global:lblDatePicker.Width=120 ; $Global:lblDatePicker.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDatePicker)

        # DatePicker
        $Global:txtDatePicker = New-Object System.Windows.Forms.DateTimePicker
        $Global:txtDatePicker.Top = 90; $Global:txtDatePicker.Left = 140; $Global:txtDatePicker.Width = 200; 
        $Global:txtDatePicker.Format = [windows.forms.datetimepickerFormat]::custom
        $Global:txtDatePicker.CustomFormat = "dd/MMM/yyyy"
        $Global:txtDatePicker.Text = (get-date)
        $Global:form.Controls.Add($Global:txtDatePicker)

        If ($Global:txtDatePicker.Text -gt ((get-date).AddDays(14)))
        {
            write-host "The start date for this user is more than 2 weeks out; Enablement process is being cancelled"
        }
    }

    $Global:chkEmerTerm = ""
    If ($Global:chkTermination.Checked -eq "Checked")
    {
        $Global:form.Height = $Global:form.Height + 30
        #Add Emergency Termination checkbox
        $Global:chkEmerTerm = New-Object Windows.Forms.checkbox 
        $Global:chkEmerTerm.Left = 140; $Global:chkEmerTerm.Width = 250; $Global:chkEmerTerm.Top = 90
        $Global:chkEmerTerm.Text = "Emergency Termination" 
        $Global:chkEmerTerm.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkEmerTerm) 
        # Obtain Value with: $Global:chkEmerTerm.Checked
    }

    If ($Global:chkOOOMsg.Checked -eq "Checked")
    {
        $Global:form.Height = $Global:form.Height + 30
        #Add Radio Box for Standard OOO or custom OOO
        $Global:chkOOOStd = New-Object Windows.Forms.RadioButton
        $Global:chkOOOStd.Left = 140; $Global:chkOOOStd.Width = 250; $Global:chkOOOStd.Top = 60
        $Global:chkOOOStd.Text = "Standard Out Of Office"
        $Global:chkOOOStd.Checked = $true   # set a default value 
        $Global:form.Controls.Add($Global:chkOOOStd) 
        # Obtain Value with: $Global:chkOOOStd.Checked
        $Global:chkoooStd.Add_MouseClick(
        {
#            $Global:form.Controls.Remove($Global:lblOOOCustAddr)
            $Global:form.Controls.Remove($Global:txtOOOCustAddr)
        })

        $Global:form.Height = $Global:form.Height + 30
        #Add Radio Box for Standard OOO or custom OOO
        $Global:chkOOOCust = New-Object Windows.Forms.RadioButton
        $Global:chkOOOCust.Left = 140; $Global:chkOOOCust.Width = 250; $Global:chkOOOCust.Top = 80
        $Global:chkOOOCust.Text = "Custom Out Of Office"
        $Global:chkOOOCust.Checked = $false   # set a default value 
        $Global:form.Controls.Add($Global:chkOOOCust)
        $Global:chkooocust.Add_MouseClick(
        {
#            $Global:lblOOOCustAddr = New-Object System.Windows.Forms.Label   
#            $Global:lblOOOCustAddr.Text = "Custom Address:"  
#            $Global:lblOOOCustAddr.Top = 110 ; $Global:lblOOOCustAddr.Left = 50; $Global:lblOOOCustAddr.Width=120 ; $Global:lblOOOCustAddr.AutoSize = $true
#            $Global:form.Controls.Add($Global:lblOOOCustAddr)    # Add to Form 
            # 
            $Global:txtOOOCustAddr = New-Object Windows.Forms.TextBox  
            $Global:txtOOOCustAddr.Top = 110; $Global:txtOOOCustAddr.Left = 140; $Global:txtOOOCustAddr.Width = 200;
            $Global:txtoooCustAddr.MaxLength = 1000  
            $Global:txtOOOCustAddr.Text = "(Enter Address)"
            $Global:form.Controls.Add($Global:txtOOOCustAddr)    # Add to Form
            $Global:InputFocus = $Global:txtOOOCustAddr
        })
        # Obtain Value with: $Global:chkchkOOOCust.Checked
    }
}

Function Execute-HideMbx
{
    write-host "Setting user to be Hidden from the Address Book..." -ForegroundColor Yellow
    $Global:u.msExchHideFromAddressLists.value = $True
    $Global:u.CommitChanges()
    $LineToWrite = $RecordEvent + "CHG" + "`t" + "Setting user to be Hidden from the Address Book"
    WriteReportEvent
}

# This function connects to Active Directory and gets the record for the user 
Function GetUserDN($strUID)
{
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = "LDAP://DC=global,DC=ul,DC=com"
    $objSearcher.Filter = [string]::format("(sAMAccountName={0})",$strUID)
    $ux = $null
    $ux = $objSearcher.FindOne()
    
    if ($null -eq $ux)
    {
        return $null
    } 
    else
    {
        return $ux.Properties.distinguishedname
    }
}

# Gets AD ACcount Properties
Function GetAcctInfo($ENo)
{
    $Global:ADCmt = ""
    $Global:u = ""
    $ADExists = [bool]($strDN = GetUserDN $ENo -ErrorAction SilentlyContinue)
    If ($ADExists -eq "True")
    {
        $ObjExists = [bool]($Global:UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion -ErrorAction silentlyContinue)
        If ($ObjExists -eq $True)
        {
            $Global:UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion -ErrorAction silentlyContinue
        }
        else
        {
            $cnt = 0
            Do
            {
                start-sleep -Seconds 5
                $strDN = GetUserDN $ENo -ErrorAction SilentlyContinue
                $ObjExists = [bool]($Global:UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion -ErrorAction silentlyContinue)

                If ($ObjExists -eq $True)
                {
                    $Global:UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion -ErrorAction silentlyContinue
                }
                else
                {
                    $Cnt++
                }
            } while (($Cnt -le 5) -and ($ObjExists -eq $False))

            If ($Cnt -ge 5)
            {
                write-host "AD Record not found"
                $LineToWrite = $WhoAmI + "`t" + "AD Record Not Found!"
                WriteReportEvent
            }
        }
     }

    $MbxExists = [bool](get-mailbox -identity $ENo -ErrorAction SilentlyContinue)
    If ($MbxExists -eq "True")
    {
        $Global:inf = get-mailbox $ENo
        If ([bool](get-ADUser -Filter {SamAccountName -eq $ENo} -ErrorAction SilentlyContinue))
        {
            $Global:strUserPath = [string]::format("LDAP://{0}", $strDN)
            $Global:u = new-object System.DirectoryServices.DirectoryEntry($Global:strUserPath)
            $Global:ADCmt = $Global:u.ExtensionAttribute14.value
        }
        else
        {
            $Global:ADCmt = "*****No Active Directory Account For This User*****"
            $Global:strUserPath = "No Active Directory Account for this User"            
        }

        If ($Global:FormRefresh -ne "Y")
        {
            $LineToWrite = $WhoAmI + "`t" + "DisplayName                    :" + "`t" + $Global:inf.DisplayName
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "Employee Type                  :" + "`t" + $Global:inf.CustomAttribute1
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "AD Container                   :" + "`t" + $Global:strUserPath
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "LitigationHoldEnabled          :" + "`t" + $Global:inf.LitigationHoldEnabled
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "LitigationHoldDate             :" + "`t" + $Global:inf.LitigationHoldDate
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "LigitationHoldOwner            :" + "`t" + $Global:inf.LitigationHoldOwner
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "Active InPlace Holds           :" + "`t" + ($Global:inf.InPlaceHolds -join ",")
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "Hidden from Address Book O365  :" + "`t" + $Global:inf.HiddenFromAddressListsEnabled
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "Hidden from Address Book AD    :" + "`t" + $Global:u.msExchHideFromAddressLists.value
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "O365 Retention Comment         :" + "`t" + $Global:inf.RetentionComment
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "O365 Retention Comment         :" + "`t" + $Global:inf.ExtensionAttribute14
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "AD Retention Comment           :" + "`t" + $Global:ADCmt
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "AD Object Protected            :" + "`t" + $Global:UsrDetails.ProtectedFromAccidentalDeletion
            WriteReportEvent
        }

        PrtMBxAccess($Global:MbxAccess)

        $LineToWrite = "`n"
        WriteReportEvent
    }
    else
    {
        write-host""
        write-host "Mailbox does not exist for Emp#" $ENo -ForegroundColor Red
    }
}

Function Get-ENo
{
    If ($Global:ENo -ne "0")
    {
        $Global:MbxExists = [bool](Get-EXOMailbox $Global:ENo -ErrorAction SilentlyContinue)
        $ErrorActionPreference = "SilentlyContinue"
        $Global:ADExists = [bool](get-ADUser $Global:ENo)
        $ErrorActionPreference = "Continue"
        Do
        {
            If (($Global:MbxExists -eq $False) -or ($Global:ADExists -ne $True))
            {
                
                If ($Global:ADExists -ne $True)
                {
                    $Output = $wshell.Popup("No AD Account Exists For this Individual there is No AD Object to Protect.",0,"No Account Found",0+32)
                }
                If ($Global:MbxExists -eq $False)
                {
                    $Output = $wshell.Popup("No Mailbox Exists for this individual.",0,"No Mailbox Found",0+32)
                    If ($Script:ExectueLegalHold -ne "Y")
                    {
                        write-host "Enter Employee Number : " -ForegroundColor Red -NoNewline
                        $Global:ENo = Read-Host
                        $Global:MbxExists = [bool](Get-EXOMailbox $Global:ENo -ErrorAction SilentlyContinue)
                    }
                    else
                    {
                        $MbxExists=$True
                        $ENo = 0
                        $Global:txtHost.ReadOnly = $false
                        $ButGetENo.visible = $true
                        $Global:txtHost.Text = "Invalid"
                        Remove-Item $Global:ReportFile
                    }
                }
            }
        } while (($Global:MbxExists -eq $False) -and ($ENo -ne "0"))

        If ($ENo -ne "0")
        {
            GetAcctInfo($Global:ENo)
            If ($Global:ADExists -ne $True)
            {
                If ((get-mailbox $Global:ENo).RetentionComment.Length -ne 0)
                {
                #Must use the retention comment on the mailbox
                    $Global:StoredCmt = (Get-Mailbox $Global:ENo).RetentionComment
                }
                else
                {
                    $Global:StoredCmt = (Get-Mailbox $Global:ENo).CustomAttribute14
                }
            }
            else
            {
                $Global:StoredCmt = $Global:u.ExtensionAttribute14.value
            }
        }
    }
}

Function O365HasLicenses
{
    $Global:UsageLoc = (Get-MsolUser -UserPrincipalName $Global:UPN).UsageLocation
    $Global:HasBIFree = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:POWER_BI_STANDARD"})
    $Global:HasBIPro = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:POWER_BI_PRO"})
    $Global:HasCRM = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:DYN365_ENTERPRISE_PLAN1"})
    $Global:HasPhone = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:MCOEV"})
    $Global:HasCAP = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:MCOCAP"})
    $Global:HasEMS = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:EMSPREMIUM"})
    $Global:HasExP2 = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:EXCHANGEENTERPRISE"})
    $Global:HasE5 = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:ENTERPRISEPREMIUM"})
    $Global:HasE3 = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:ENTERPRISEPACK"})
    $Global:HasMRSS = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"})
    $Global:HasFlowP2 = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:FLOW_P2"})
    $Global:HasATPDef = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:WIN_DEF_ATP"})
    $Global:HasATPDefMAC = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MDATP_XPLAT"})
    $Global:HasATPP1 = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:ATP_ENTERPRISE"})
    $Global:HasMeeting = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MEETING_ROOM"})
    $Global:HasPAppsP2 = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:POWERFLOW_P2"})
    $Global:HasAutoPUsr = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:FLOW_PER_USER"})
    $Global:HasAutoPUsrPRA = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:POWERAUTOMATE_ATTENDED_RPA"})
    $Global:HasAppsPUsr = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:POWERAPPS_PER_USER"})
    $Global:HasAudioConf = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MCOMEETADV"})
    $Global:HasRmtAssist = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MICROSOFT_REMOTE_ASSIST"})
    $Global:HasInfoBarr = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:M365_INSIDER_RISK_MANAGEMENT"})
    $Global:HasProjP3 = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:PROJECT_PLAN3_DEPT"})
    $Global:HasProjP5 = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:PROJECTPREMIUM"})
    $Global:HasVisioP1 = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:VISIOONLINE_PLAN1"})
    $Global:HasVisioP2 = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:VISIOCLIENT"})
}

Function O365Licenses
{
    $Global:O365Lic = (Get-MsolUser -UserPrincipalName $Global:UPN).Licenses
	$Global:UserDet = $Global:UPN + " - " + $Global:UserLicense.DisplayName + " (" + $Global:ADUser.ExtensionAttribute1 + ")"

    If ($O365Lic.count -gt 0)
    {
        $Global:LicAssigned = ""
        $LineToWrite = "Reviewing O365 Licenses Assigned to " + $Global:UPN
	    WriteLogEvent

        foreach ($Global:O365Lic in $Global:O365Lic)
        {
            switch ($Global:O365Lic.AccountSkuID)
            {
                "ul:ENTERPRISEPREMIUM"
                {
                    $text = "Enterprise E5"
                    $SubLic = ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:ENTERPRISEPREMIUM"}).ServiceStatus
	                foreach ($SubLic in $SubLic)
	                {
                        switch ($subLic.ServicePlan.ServiceName)
                        {
			                "BI_AZURE_P2"
			                {
                                If ($subLic.ProvisioningStatus -eq "Success")
                                {
                                    If ($text -like "*w/*")
                                    {
                                        $text = $text + ";"
                                    }
                                    $text = $text + "(w/PowerBI Pro)"
                                }
                            }
			                "MCOEV"
                            {
                                If ($subLic.ProvisioningStatus -eq "Success")
                                {
                                    If ($text -like "*w/*")
                                    {
                                        $text = $text + ";"
                                    }
                                    $text = $text + "(w/Phone System)"
                                }
                            }
			                "MCOMEETADV"
		                    {
                                If ($subLic.ProvisioningStatus -eq "Success")
                                {
                                    If ($text -like "*w/*")
                                    {
                                        $text = $text + ";"
                                    }
                                    $text = $text + "(w/Audio Conferencing)"
                                }
                            }
                        }
 	                }
                }
				"ul:ENTERPRISEPACK"
                {
                    $text = "Enterprise E3"
                }
                "ul:AAD_PREMIUM_P2"
                {
                    $text = "Azure Active Directory Premium P2"
                }				
                "ul:EXCHANGEENTERPRISE"
                {
                    $text = "Exchange Online Plan2"
                }
                "ul:POWER_BI_STANDARD"
                {
                    $text = "PowerBI (Free)"
                }
                "ul:POWER_BI_PRO"
                {
                    $text = "PowerBI Pro" 
                }
                "ul:DYN365_ENTERPRISE_PLAN1"
                {
                    $text = "Dynamics 365 Customer Engagement Plan"
                }
                "ul:EMS"
                {
                    $text = "Enterprise Mobility Suite/Intune (EMS)"
                }
                "ul:EMSPREMIUM"
                {
                    $text = "Enterprise Mobility + Security E5"
                }				
                "ul:POWER_BI_INDIVIDUAL_USER"
                {
                    $text = "Power BI for O365"
                }
                "ul:MCOIMP"
                {
                   $text = "Skype for Business Online (Plan 1)"
                }
                "ul:SHAREPOINTSTANDARD"
                {
                    $text = "SharePoint Online (Plan 1)"
                }
                "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
                {
                    $text = "Microsoft Relationship Sales Solution"
                }
                "ul:TEAMS_COMMERCIAL_TRIAL"
                {
                    $text = "Microsoft Teams Commercial Cloud (User Initiated)"
                }
                "ul:ATP_ENTERPRISE"
                {
                    $text ="Advanced Threat Protection Plan1"
                }
                "ul:MDATP_XPLAT"
                {
                    $text ="Advanced Threat Protection for MAC"
                }
                "ul:WIN_DEF_ATP"
                {
                    $text ="Defender Advanced Threat Protection"
                }
                "ul:FLOW_P2"
                {
                    $text = "Microsoft Flow Plan 2"
                }
                "ul:MEETING_ROOM"
                {
                    $text = "Teams Rooms Standard"
                }
                "ul:POWERFLOW_P2"
                {
                    $text = "PowerApps Plan2"
                }
                "ul:MCOEV"
                {
                    $text = "Phone System"
                }
                "ul:MCOMEETADV"
                {
                    $text = "Audio Conferencing"
                }
                "ul:MICROSOFT_REMOTE_ASSIST"
                {
                    $text = "Microsoft Remote Assist"
                }
                "ul:MCOCAP"
                {
                    $text = "Common Area Phone"
                }
                "ul:PROJECTPROFESSIONAL"
                {
                    $text = "Project Plan 3"
                }
                "ul:PROJECTPREMIUM"
                {
                    $text = "Project Plan 5"
                }
                "ul:FLOW_PER_USER"
                {
                    $text = "Power Automate p/User Plan"
                }
                "ul:POWERAPPS_PER_USER"
                {
                    $text = "Power Apps p/User Plan"
                }
                "ul:POWERAUTOMATE_ATTENDED_RPA"
                {
                    $text = "PowerAutomate p/User with RPA"
                }
                "ul:VISIOONLINE_PLAN1"
                {
                    $text = "Visio Plan 1"
                }
                "ul:VISIOCLIENT"
                {
                    $text = "Visio Plan 2"
                }
            }

            If ($Global:LicAssigned -eq "")
            {
                $Global:LicAssigned = $text
            }
            else
            {
                $Global:LicAssigned = $Global:LicAssigned + ", " + $text
            }
            $LineToWrite = "UPDA" + "`t" + $text
	        WriteLogEvent
        }
    }
    else
    {
        $Global:LicAssigned = "No licenses assigned to this account"
        $LineToWrite = "No O365 Licenses Assigned to " + $Global:UPN
	    WriteLogEvent
    }
}


# This prints details regarding mailbox access to the console and report file
Function PrtMBxAccess($Global:MbxAccess)
{
    $Global:MbxAccess = Get-MailboxPermission $ENo |Where-Object {$_.User -like "*global.ul.com"}
    $Global:MbxFldrAccess = Get-MailboxFolderPermission $ENo |Where-Object {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous")}
    $Global:MbxAccessForm = "(None)"
    foreach ($Global:MbxAccess in $Global:MbxAccess)
    {
#       write-host "Accounts Granted Access       : " $Global:MbxAccess.User -ForegroundColor Yellow
        $UsrExists = [bool](Get-mailbox $Global:MbxAccess.User -ErrorAction SilentlyContinue)
        If ($Global:MbxAccessForm -eq "(None)")
        {
            If ($UsrExists -eq $True)
            {
                $Global:MbxAccessForm = (get-mailbox $Global:MbxAccess.User).PrimarySMTPAddress + " (" + $Global:MbxAccess.AccessRights + ")"
                $SupMbx = (get-mailbox $Global:MbxAccess.User).PrimarySMTPAddress
                $LineToWrite = $WhoAmI + "`t" + "User Granted Mailbox Permission:" + "`t" + $(get-mailbox $Global:MbxAccess.User).PrimarySMTPAddress + "- PermissionLevel: " + $Global:MbxAccess.AccessRights
                WriteReportEvent
            }
            else
            {
                $Global:MbxAccessForm = $Global:MbxAccess.User + " (" + $Global:MbxAccess.AccessRights + ")"
            }
        }
        else
        {
            $Global:MbxAccessForm = $Global:MbxAccessForm + ", " + (get-mailbox $Global:MbxAccess.User).PrimarySMTPAddress + " (" + $Global:MbxAccess.AccessRights + ")"
            $SupMbx = $SupMbx + ", " + (get-mailbox $Global:MbxAccess.User).PrimarySMTPAddress
            $LineToWrite = $WhoAmI + "`t" + "User Granted Mailbox Permission:" + "`t" + $(get-mailbox $Global:MbxAccess.User).PrimarySMTPAddress + "- PermissionLevel: " + $Global:MbxAccess.AccessRights
            WriteReportEvent
        }
    }

    foreach ($Global:MbxFldrAccess in $Global:MbxFldrAccess)
    {
        If ($Global:MbxAccessForm -eq "(None)")
        {
            If ($Global:MbxFldrAccess.User.DisplayName -like "MBX*")
            {
                $Global:MbxAccessForm = (get-distributiongroup $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + " (" + $Global:MbxFldrAccess.AccessRights + ")"
                $SupMbx = (get-distributiongroup $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress
                $LineToWrite = $WhoAmI + "`t" + "User Granted Folder Permission :" + "`t" + $(get-distributiongroup $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + "- PermissionLevel: " + $Global:MbxFldrAccess.AccessRights
            }
            else
            {
                $UsrExists = [bool](Get-mailbox $Global:MbxFldrAccess.User.DisplayName -ErrorAction SilentlyContinue)
                If ($UsrExists -eq $True)
                {
                    $Global:MbxAccessForm = $Global:MbxAccessForm + (get-mailbox $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + " (" + $Global:MbxFldrAccess.AccessRights + ")"
                    $SupMbx = (get-mailbox $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress
                }
                else
                {
                    $Global:MbxAccessForm = $Global:MbxFldrAccess.User.DisplayName + " (" + $Global:MbxFldrAccess.AccessRights + ")"
                }
                $LineToWrite = $WhoAmI + "`t" + "User Granted Folder Permission :" + "`t" + $(get-mailbox $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + "- PermissionLevel: " + $Global:MbxFldrAccess.AccessRights
            }
            WriteReportEvent
        }
        else
        {
            If ($Global:MbxFldrAccess.User.DisplayName -like "MBX*")
            {
                $Global:MbxAccessForm = $Global:MbxAccessForm + (get-distributiongroup $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + " (" + $Global:MbxFldrAccess.AccessRights + ")"
                $SupMbx = (get-distributiongroup $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress
                $LineToWrite = $WhoAmI + "`t" + "User Granted Folder Permission :" + "`t" + $(get-distributiongroup $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + "- PermissionLevel: " + $Global:MbxFldrAccess.AccessRights
            }
            else
            {
                $Global:MbxAccessForm = $Global:MbxAccessForm + ", " + (get-mailbox $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + " (" + $Global:MbxFldrAccess.AccessRights + ")"
                $SupMbx = $SupMbx + ", " + (get-mailbox $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress
                $LineToWrite = $WhoAmI + "`t" + "User Granted Folder Permission :" + "`t" + $(get-mailbox $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + "- PermissionLevel: " + $Global:MbxFldrAccess.AccessRights
            }
            WriteReportEvent
        }
    }

    If ($Global:MbxAccessForm -ne "(None)")
    {
        if (((get-date).AddHours(-1) -le $Global:inf.litigationholddate) -and ($Global:inf.CustomAttribute1 -eq "T"))
        {
            $SendTo = $Global:LegalTeamMsgs
            $SendCC = "Sandi.Glazebrook@ul.com"
            if ($Mbx.CustomAttribute1 -like "*Employee*")
            {
                $MessageSubject = "Access Previously Granted to Terminated Employee Mailbox and OneDrive"
                $MessageBody = "While placing this terminated account on Legal Hold access was previously granted to the Mailbox and OneDrive information.  The details are provided below.<ul type=""disc""><li>Terminated " + $Global:inf.CustomAttribute1 + " details " + $Mbx.Alias + " - " + $Global:inf.DisplayName + "</li><li>Individual given access to this information " + $SupMbx.Alias + " - " + $SupMbx.DisplayName + "</li></ul>UL Account Provisioning Team</li></ul>"
            }
            else
            {
                $MessageSubject = "Access Previously Granted to Terminated Non-Employee Mailbox"
                $MessageBody = "While placing this terminated account on Legal Hold access was previously granted to thes Mailbox information.<ul type=""disc""><li>Terminated " + $Global:inf.CustomAttribute1 + " details " + $Global:inf.Alias + " - " + $Global:inf.DisplayName + "</li><li>Individual given access to this information " + $SupMbx.Alias + " - " + $SupMbx.DisplayName + "</li></ul>UL Account Provisioning Team</li></ul>"
            }
            invoke-expression -Command .\SendSMTPMessage.ps1
        }
    }
}

Function Publish-Form
{
    If ($Global:InputFocus -eq $null)
    {
        $Global:InputFocus = $Global:okButton
    }
    ## Finalize Form and Show Dialog
    $Global:form.Add_Shown( { $form.Activate(); $Global:InputFocus.Focus() } )  #Activate and Set Focus 
#    $Global:form.Add_Shown( { $form.Activate(); $Global:okButton.Focus() } )  #Activate and Set Focus 
    $Global:Result = $Global:form.ShowDialog()          ## Show the form, and wait for the response
}

#  Removes access to a mailbox
Function RemoveAccess($ENo)
{
    Do
    {
        write-host "Enter Employee Number for Access to Removed From" $ENo "? " -ForegroundColor Yellow -NoNewline
        $RemENo = Read-Host
        Remove-MailboxPermission $ENo -AccessRights FullAccess -User $RemENo -Confirm:$false
        $LineToWrite = $WhoAmI + "`t" + "Removed " + $RemENo + "Access to " + $ENo + "Mailbox" + "`n"
        WriteReportEvent
        Write-Host "Removed" $RemENo "Access to" $ENo "Mailbox" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Remove Another Users access to this mailbox (Y/N)? " -ForegroundColor Yellow -NoNewline
        $RemAccess = Read-Host
    } while ($RemAccess -eq "Y")
}

Function RetentPolicy
{
    $MbxCreated = [bool](get-mailbox $Global:UPN -ErrorAction SilentlyContinue)

    If ($MbxCreated -eq "True")
    {
        Write-host "Setting Retetnion Policy on mailbox" -ForegroundColor Green
        If (($Global:ADUser.extensionattribute4 -eq "IT") -or ((get-mailbox $Global:UPN).WhenCreated -gt "03/23/2020"))
        {
            Set-MailBox $Global:UPN -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL MRM Policy - 3 yr Delete"
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Updating Retention Policy to IT Policy for " + $Global:UPN
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        }
        else
        {
            Set-MailBox $Global:UPN -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 365 Day Delete"
            Set-Mailbox $Global:UPN -RetentionHoldEnabled $true -StartDateForRetentionHold 04/01/2011
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Updating Retention Policy to UL Default Policy for " + $Global:UPN
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        }
	    Set-CASMailBox $Global:UPN -ImapEnabled $false -PopEnabled $false -ActiveSyncEnabled $false
    }
    else
    {
        write-host "Mailbox does not exist or is in the process of being created.  If the license was just assigned"
        write-host "please allow 3-5 minutes for the mailbox to be created and the select Option 10 from the menu"
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Mailbox Does Not Exist " + $Global:UPN
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
}

#  Diplays the account details to the console
Function ShowDetails
{
    write-host "Employee Number               : " $Global:inf.Alias -ForegroundColor Green
    write-host "DisplayName                   : " $Global:inf.DisplayName -ForegroundColor Green
    write-host "Employee Type                 : " $Global:inf.CustomAttribute1 -ForegroundColor Green
    write-host "AD Container                  : " $Global:strUserPath -ForegroundColor Green
    write-host "LitigationHoldEnabled         : " $Global:inf.LitigationHoldEnabled -ForegroundColor Green
    write-host "LitigationHoldDate            : " $Global:inf.LitigationHoldDate -ForegroundColor Green
    write-host "LigitationHoldOwner           : " $Global:inf.LitigationHoldOwner -ForegroundColor Green
    write-host "Active InPlace Holds          : " ($Global.inf.InPlaceHolds -join ",") -ForegroundColor Green
    write-host "Hidden from Address Book O365 : " $Global:inf.HiddenFromAddressListsEnabled -Foregroundcolor Green
    write-host "Hidden from Address Book AD   : " $Global:u.msExchHideFromAddressLists.value -Foregroundcolor Green
    write-host "O365 Retention Comment        : " $Global:inf.RetentionComment -ForegroundColor Green
    write-host "AD Retention Comment          : " $Global:ADCmt -ForegroundColor Green
    write-host "AD Object Protected           : " $Global:UsrDetails.ProtectedFromAccidentalDeletion -ForegroundColor Green
    PrtMBxAccess($Global:MbxAccess)
}

# Assign Standard license set
Function StandardLicenses
{
    If ($Global:HasEMS -ne "True")
	{
	    Write-Host "`nAssigning Enterprise Mobility Suite/Intune (EMS) License to" $Global:UPN -ForegroundColor Yellow
#		Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:EMS" -UsageLocation "US"
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:EMS"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Enterprise Mobility Suite/Intune (EMS) License to " + $Global:UPN
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}  

    If (($Global:HasBIFree -eq $False) -and ($Global:HasBIPro -eq $False))
    {
	    Write-Host "`nAssigning PowerBI Free License to" $Global:UPN -ForegroundColor Yellow
#	    Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:POWER_BI_STANDARD" -UsageLocation "US"
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:POWER_BI_STANDARD"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning PowerBI Free License to " + $Global:UPN
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}

    If ($Global:HasATPP1 -eq $False)
    {
	    Write-Host "`nAssigning Advanced Threat Protection Plan1 License to" $Global:UPN -ForegroundColor Yellow
	    Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:ATP_ENTERPRISE"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Advanced Threat Protection Plan1 License to " + $Global:UPN
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}

    If ($Global:HasPhone -eq $False)
    {
	    Write-Host "`nAssigning Phone System License to" $Global:UPN -ForegroundColor Yellow
#	    Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MCOEV" -UsageLocation "US"
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MCOEV"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Phone System License to " + $Global:UPN
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}
}

<#
function UnAssign-ATPDefLic
{
    if ($Global:HasATPDef -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Windows Defender Advanced Threat Protection license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Windows Defender Advanced Threat Protection License Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Removing Windows Defender Advanced Threat Protection License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:WIN_DEF_ATP"
    }
    WriteLogEvent
}

function UnAssign-ATPPDefMACLic
{
    if ($Global:HasATPDefMAC -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Advanced Threat Protection for MAC license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Advanced Threat Protection for MAC License Assigned to " + $Global:UPN
    }
    else
    {
        $LineToWrite = "REVI" + "`t" + "Removing Advanced Threat Protection for MAC License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:MDATP_XPLAT"
    }
    WriteLogEvent
}

function UnAssign-ATPP1Lic
{
    if ($Global:HasATPP1 -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Advanced Threat Protection Plan1 license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Advanced Threat Protection Plan1 License Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Removing Advanced Threat Protection Plan1 License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:ATP_ENTERPRISE"
    }
    WriteLogEvent
}

Function UnAssign-AudioConferencingLic
{
    if ($Global:HasAudioConf -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have an Audio Conferencing license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Audio Conferencing License Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Removing Audio Conferencing License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:MCOMEETADV"
    }
    WriteLogEvent
}

Function UnAssign-CommonAreaPhoneLic
{
    if ($Global:HasCAP -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Common Area Phone license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Common Area Phone License Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Removing Common Area Phone License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:MCOCAP"
    }
    WriteLogEvent
}
    
function UnAssign-CRMLic
{
    If ($Global:HasCRM -ne "True")
    {
      	$Output = $wshell.Popup("This individual does not have a Dynamics 365 Customer Engagement Plan license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Dynamics 365 Customer Engagement Plan License Assigned to " + $Global:UPN
    }
	else
	{
        $LineToWrite = "REVI" + "`t" + "Removing Dynamics 365 Customer Engagement Plan License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:DYN365_ENTERPRISE_PLAN1"
    }    
    WriteLogEvent
}

function UnAssign-E3Lic
{
    If ($Global:HasE3 -ne "True")
    {
      	$Output = $wshell.Popup("This individual does not have a Enterpris E3 license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Dynamics 365 Customer Engagement Plan License Assigned to " + $Global:UPN
    }
    else
    {
        $LineToWrite = "REVI" + "`t" + "Removing Enterprise E3 Plan License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:ENTERPRISEPACK"
    }
    WriteLogEvent
}

function UnAssign-EMSLic
{
    If ($Global:HasEMS -ne "True")
	{
        $Output = $wshell.Popup("This individual does not have a EMS license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Enterprise Mobility Suite/Intune (EMS) License Assigned to " + $Global:UPN
    }
	else
    {
		Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:EMS"
        $LineToWrite = "REVI" + "`t" + "Removing Enterprise Mobility Suite/Intune (EMS) License from " + $Global:UPN
    }
    WriteLogEvent
}

function UnAssign-MeetingLic
{
    if ($Global:HasMeeting -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Teams Rooms Standard license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Teams Rooms Standard License Already Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Removing Teams Rooms Standard License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:MEETING_ROOM"
    }
    WriteLogEvent
}

Function UnAssign-MRSSLic
{
    if ($Global:HasMRSS -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Microsoft Relationship Sales Solution license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Microsoft Relationship Sales Solution License Assigned to " + $Global:UPN
    }
    else
    {
        $LineToWrite = "REVI" + "`t" + "Removing Microsoft Relationship Sales Solution License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
    }
	WriteLogEvent 
}

function UnAssign-P2Lic
{
    if ($Global:HasExP2 -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Exchange Online P2 license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Exchange Online P2 License Assigned to " + $Global:UPN
    }
    else
    {
        $LineToWrite = "REVI" + "`t" + "Removing Exchange Online P2 License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:EXCHANGEENTERPRISE"
    }
    WriteLogEvent
}

Function UnAssign-PAppsP2Lic
{
    if ($Global:HasPAppsP2 -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a PowerApps P2 license assigned.",0,"License Noty Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No PowerApps Plan 2 License Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Removing PowerApps Plan 2 License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:POWERFLOW_P2"
    }
    WriteLogEvent
}

Function UnAssign-AppsPUsrLic
{
    if ($Global:HasAppsPUsr -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Power Apps per User Plan license assigned.",0,"License Noty Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Power Apps per User Plan License Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Removing Power Apps per User Plan License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:POWERAPPS_PER_USER"
    }
    WriteLogEvent
}

Function UnAssign-AutoPUsrLic
{
    if ($Global:HasAutoPUsr -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Power Automate per User Plan license assigned.",0,"License Noty Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Power Automate per User Plan License Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Removing Power Automate per User Plan License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:FLOW_PER_USER"
    }
    WriteLogEvent
}

Function UnAssign-AutoPUsrRPALic
{
    if ($Global:HasAutoPUsrRPA -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Power Automate per User W/RPA Plan license assigned.",0,"License Noty Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Power Automate per User W/RPA Plan License Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Removing Power Automate per User w/RPA Plan License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:POWERAUTOMATE_ATTENDED_RPA"
    }
    WriteLogEvent
}

Function UnAssign-PhoneSystemLic
{
    if ($Global:HasPhone -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Microsoft Phone System license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Phone System License Assigned to " + $Global:UPN
    }
    else
    {
     	$LineToWrite = "REVI" + "`t" + "Removing Phone System License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:MCOEV"
    }
    WriteLogEvent
}

function UnAssign-PowerBIProLic
{
    if ($Global:HasBIPro -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a PowerBI Pro license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No PowerBI Pro License Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Removing PowerBI Pro License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:POWER_BI_PRO"
    }
    WriteLogEvent
    if ($Global:HasBIFree -ne "True")
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning PowerBI Free License to " + $Global:UPN
        WriteLogEvent
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:POWER_BI_STANDARD"
    }
}

function UnAssign-ProjP3Lic
{
    if ($Global:HasProjP3 -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Project Plan 3 license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Project Plan 3 License Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Removing Project Plan 3 License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:PROJECTPROFESSIONAL"
    }
    WriteLogEvent
}

function UnAssign-ProjP5Lic
{
    if ($Global:HasProjP5 -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Project Plan 5 license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Project Plan 5 License Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Removing Project Plan 5 License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:PROJECTPREMIUM"
    }
    WriteLogEvent
}

function UnAssign-RemoteAsstLic
{
    If ($Global:HasRmtAssist -eq "True")
    {
		Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:MICROSOFT_REMOTE_ASSIST"
        $LineToWrite = "DELE" + "`t" + "Removing Microsoft Remote Assist License from " + $Global:UPN
    }
    else
    {
        $Output = $wshell.Popup("This individual does not have a Microsoft Remote Assist license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "DELE" + "`t" + "No Microsoft Remote Assist License assigned to " + $Global:UPN
    }
    WriteLogEvent
}

function UnAssign-VisioP1Lic
{
    If ($Global:HasVisioP1 -eq "True")
    {
		Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:VISIOONLINE_PLAN1"
        $LineToWrite = "DELE" + "`t" + "Removing Visio Plan 1 License from " + $Global:UPN
    }
    else
    {
        $Output = $wshell.Popup("This individual does not have a Visio Plan 1 license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "DELE" + "`t" + "No Visio Plan 1 License assigned to " + $Global:UPN
    }
    WriteLogEvent
}

function UnAssign-VisioP2Lic
{
    If ($Global:HasVisioP2 -eq "True")
    {
		Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:VISIOCLIENT"
        $LineToWrite = "DELE" + "`t" + "Removing Visio Plan 2 License from " + $Global:UPN
    }
    else
    {
        $Output = $wshell.Popup("This individual does not have a Visio Plan 2 license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "DELE" + "`t" + "No Visio Plan 2 License assigned to " + $Global:UPN
    }
    WriteLogEvent
}
#>

#  Writes events to the Log File
function WriteLogEvent
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LocalMachineName + "`t" + $LineToWrite
	Out-File -filepath $LogFile -append -noClobber -inputObject $RecordEvent
} #end WriteLogEvent

#  Writes events to the Report File
function WriteReportEvent
{
    If (($null -eq $ReportFile) -or ($ReportFile -eq ""))
    {
        $ReportFile = $Global:ReportFile
    }
    If ($Global:ReportFile -eq $null)
    {
    }
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LineToWrite
	Out-File -filepath $ReportFile -append -noClobber -inputObject $RecordEvent
} #end WriteReportEvent

################

# Retrieve the user name
	$WhoAmI			= WhoAmI
		
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name

$wshell = New-Object -ComObject Wscript.Shell
$Global:OKDetails = ""

$session = get-pssession -ErrorAction SilentlyContinue
If ($session)
{
    Remove-PSsession (get-Pssession)
}
Invoke-Expression -Command e:\Automation\Scripts\O365MainMenu.ps1