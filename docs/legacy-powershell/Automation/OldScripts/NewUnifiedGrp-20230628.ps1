<#   
================================================================================ 
 Name: Used SampleForm.ps1 from Dan Stolts "ITProGuru" at http://ITProGuru.com/Scripts as a template
 
 New Unified Group
 ================================================================================ 
 #
 #  05/17/2020 - SAG:Added new process using forms to perform changes$ConfDisable
 #  05/29/2020 - SAG:Fixed issues with enabling NewHolds
 #  06/02/2020 - SAG:Changed Get-Mailbox to Get-EXOMailbox
 #  07/01/2020 - SAG:Added additional buttons for and details (InPlaceHolds,Hide/Unhide,RemovingMBXAccess)
 #  07/02/2020 - SAG:Added Reporting Features
 #  10/25/2020 - SAG:Modified code to use a form
 #  12/09/2020 - SAG:Modified code so that PowerBI requests are now supported by CT Datahub Support team and these groups are no longer created 
 #  06/21/2021 - PW:Modified code to fix an issue where the Get-Team command would find more than one entry if the display name was the same as the start of another
 #  08/05/2021 - SAG:Modified the code to handle a separate input field for each app type that is requested.
 #  08/05/2021 - SAG:Modified code that checks if this is a licensed mailbox.  If the length of $DG.u_owner_empid is not 5 characters it will add leading "0" to the number before checking if the user is licensed.
 #  08/25/2021 - SAG:Made modifications to the group changes made by PW as it was not resetting owners correctly.  Also the checks if the group exists were not working properly so made modifications to that code.
 #  09/02/2021 - PW: Edited Line 622, not sure when it changed but the SharePoint Site Check was always failing.
 #  02/02/2022 - SAG:Added code to hide new groups from the address book
 #  03/28/2022 - SAG:Modified line 637 to use $Global:txtGrpName.Text in the $SiteName rather than $DGAlias
 #  02/28/2023 - SAG:Modified to disable external access for groups
 #  MM/DD/YYYY
#>  

#Forms Functions
function Add-AppsChkBoxes
{
    $Top = 80
    $Left = 500

    ## Apps Used With
    $Global:lblUsedWith = New-Object System.Windows.Forms.Label   
        $Global:lblUsedWith.Text = "Apps Used With:"  
        $Global:lblUsedWith.Top = $Top; $Global:lblUsedWith.Left = $Left; $Global:lblUsedWith.Width=150 ;$Global:lblUsedWith.AutoSize = $true 
        $form.Controls.Add($Global:lblUsedWith)    # Add to Form 

    $Top = $Top + 20
    ## Applications Group will be Used With        
    $Global:chkTeams = New-Object Windows.Forms.checkbox 
        $Global:chkTeams.Left = $Left; $Global:chkTeams.Width = 200; $Global:chkTeams.Top = $Top
        $Global:chkTeams.Text = "Teams"
        $Global:chkTeams.TabIndex = 6
        $Global:chkTeams.Checked = $False   # set a default value
        If ($DG."Microsoft Teams" -like "TRUE")
        {
            $Global:chkTeams.Checked = $True
        }
        $Global:form.Controls.Add($Global:chkTeams) 
        # Obtain Value with: $Global:chkTeams.Checked

    If ($Type -eq "I")
    {
        $Top = $Top + 20
        $Global:chkPlanner = New-Object Windows.Forms.checkbox 
            $Global:chkPlanner.Left = $Left; $Global:chkPlanner.Width = 200; $Global:chkPlanner.Top = $Top  
            $Global:chkPlanner.Text = "Planner" 
            $Global:chkPlanner.TabIndex = 8
            $Global:chkPlanner.Checked = $False   # set a default value
            If ($DG."Microsoft Planner" -eq "TRUE")
            {
                $Global:chkPlanner.Checked = $True
            }
            $Global:form.Controls.Add($Global:chkPlanner) 
            # Obtain Value with: $Global:chkPlanner.Checked

        $Top = $Top + 20
        $Global:chkAllApps = New-Object Windows.Forms.checkbox 
            $Global:chkAllApps.Left = $Left; $Global:chkAllApps.Width = 200; $Global:chkAllApps.Top = $Top  
            $Global:chkAllApps.Text = "All Apps" 
            $Global:chkAllApps.TabIndex = 12
            $Global:chkAllApps.Checked = $False   # set a default value
            If ($DG."All Applications" -eq "TRUE")
            {
                $Global:chkAllApps.Checked = $True
            }
            $Global:form.Controls.Add($Global:chkAllApps) 
            # Obtain Value with: $Global:chkAllApps.Checked
    }
}

function Build-DefaultForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "New Unified Group" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 740 ; $form.Height = 350  # Make the form wider 
    
    Add-FormStandardButtons

    $TopLoc = 60
    $LeftLabel = 10
    $LeftInput = 100
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"
        $Global:lblTaskNo.Top = $TopLoc; $Global:lblTaskNo.Left = $LeftLabel; $Global:lblTaskNo.Width=150 ;$Global:lblTaskNo.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtTaskNo.Top = $TopLoc; $Global:txtTaskNo.Left = $LeftInput; $Global:txtTaskNo.Width = 100; 
        $Global:txtTaskNo.Text = ""   # TaskNo
        If ($DG.u_request -ne "")
        {
            $Global:txtTaskNo.Text = $DG.u_requested_item
        } 
        $Global:form.Controls.Add($Global:txtTaskNo)    # Add to Form

    $TopLoc = $TopLoc + 30
    ## Group Name
    $Global:lblGrpName = New-Object System.Windows.Forms.Label   
        $Global:lblGrpName.Text = "Group Name:"  
        $Global:lblGrpName.Top = $TopLoc; $Global:lblGrpName.Left = $LeftLabel; $Global:lblGrpName.Width=150 ;$Global:lblGrpName.AutoSize = $true 
        $form.Controls.Add($Global:lblGrpName)    # Add to Form 
        # 
        $Global:txtGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtGrpName.TabIndex = 1 # set Tab Order 
        $Global:txtGrpName.Top = $TopLoc; $Global:txtGrpName.Left = $LeftInput; $Global:txtGrpName.Width = 375;  
        $Global:txtGrpName.Text = ""
        If ($DG.u_proposed_name -ne "")
        {
            $Global:txtGrpName.Text = $DG.u_proposed_name
        } 
        $Global:form.Controls.Add($Global:txtGrpName)    # Add to Form

    $TopLoc = $TopLoc + 30
    #Description
    $Global:lblDescription = New-Object System.Windows.Forms.Label   
        $Global:lblDescription.Text = "Description:"  
        $Global:lblDescription.Top = $TopLoc; $Global:lblDescription.Left = $LeftLabel; $Global:lblDescription.Width=150 ;$Global:lblDescription.AutoSize = $true 
        $form.Controls.Add($Global:lblDescription)    # Add to Form 
        # 
        $Global:txtDescription = New-Object Windows.Forms.TextBox  
        $Global:txtDescription.TabIndex = 2 # set Tab Order 
        $Global:txtDescription.Top = $TopLoc; $Global:txtDescription.Left = $LeftInput; $Global:txtDescription.Width = 375;
        If ($DG.u_brief_description -ne "")
        {
            $Global:txtDescription.Text = $DG.u_brief_description
        } 
        $Global:form.Controls.Add($Global:txtDescription)    # Add to Form
    
    $TopLoc = $TopLoc + 30
    #Owner Emp#
    $Global:lblOwner = New-Object System.Windows.Forms.Label   
        $Global:lblOwner.Text = "Owner Emp No:"  
        $Global:lblOwner.Top = $TopLoc; $Global:lblOwner.Left = $LeftLabel; $Global:lblOwner.Width=150 ;$Global:lblOwner.AutoSize = $true 
        $form.Controls.Add($Global:lblOwner)    # Add to Form 
        # 
        $Global:txtOwner = New-Object Windows.Forms.TextBox  
        $Global:txtOwner.TabIndex = 3 # set Tab Order 
        $Global:txtOwner.Top = $TopLoc; $Global:txtOwner.Left = $LeftInput; $Global:txtOwner.Width = 70;  
        $Global:txtOwner.Text = ""
        If ($DG.u_owner_empid -ne "")
        {
            $Global:txtOwner.Text = $DG.U_owner_empid
        } 
        $Global:form.Controls.Add($Global:txtOwner)    # Add to Form

        If (($DG.u_owner_name -ne "") -and (Test-path $InputFile))
        {
            $Global:txtOwnerName = New-Object Windows.Forms.TextBox  
            $Global:txtOwnerName.TabIndex = 3 # set Tab Order 
            $Global:txtOwnerName.Top = $TopLoc; $Global:txtOwnerName.Left = ($LeftInput+80); $Global:txtOwnerName.Width = 150;  
            $Global:txtOwnerName.Text = ""
            $Global:txtOwnerName.Text = "(" + $DG.u_group_owner + ")"
            $Global:form.Controls.Add($Global:txtOwnerName)    # Add to Form
        }

        $TopLoc = $TopLoc + 30
        #Public/Private
        $Global:chkPublic = New-Object Windows.Forms.checkbox 
            $Global:chkPublic.Left = $LeftInput; $Global:chkPublic.Width = 200; $Global:chkPublic.Top = $TopLoc
            $Global:chkPublic.Text = "Public Group"
            $Global:chkPublic.TabIndex = 4
            $Global:chkPublic.Checked = $False   # set a default value
            If ($DG.u_proposed_name -like "GRP.EXT*")
            {
                $Global:chkPublic.Visible = $False
            }
            else
            {
                If ($DG.u_group_type -eq "Public")
                {
                    $Global:chkPublic.Checked = $true
                }
            }
            $Global:form.Controls.Add($Global:chkPublic) 
            # Obtain Value with: $Global:chkPublic.Checked

    $TopLoc = $TopLoc + 30
    #VIP Training Copnfirmed
    $Global:chkVIP = New-Object Windows.Forms.checkbox 
        $Global:chkVIP.Left = $LeftInput; $Global:chkVIP.Width = 200; $Global:chkVIP.Top = $TopLoc
        $Global:chkVIP.Text = "VIP Training Complete"
        $Global:chkVIP.TabIndex = 5
        $Global:chkVIP.Checked = $Global:chkVIP.Checked   # set a default value
        If ($DG.u_vip -eq "True")
        {
            $Global:form.Controls.Add($Global:chkVIP)
        }         

    If ($DG.u_app_used_with -eq "PowerBI")
    {
        $TopLoc = $TopLoc + 60
        ## PowerBI Notice
        $Global:lblPowerBI = New-Object System.Windows.Forms.Label   
        $Global:lblPowerBI.Text = "*** PowerBI Group requests are now supported by the CT Datahub Support team ***"
        $Global:lblPowerBI.ForeColor = "Red"
        $Global:lblPowerBI.Left = $LeftInput + 20; $Global:lblPowerBI.Width = 200; $Global:lblPowerBI.Top = $TopLoc; $Global:lblPowerBI.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblPowerBI)    # Add to Form
    }
}

function Check-GroupName
{
    $SenderAuthentication = $true
    $ProposedName = $DG.u_group_name

    If ($ProposedName -match "GRP")
    {
        $ProposedName = ($ProposedName.TrimStart("grp"))
        $ProposedName = ($ProposedName.TrimStart("GRP"))
        $ProposedName = ($ProposedName.TrimStart(" ."))
    }

    If ($ProposedName -match $DG.u_group_location)
    {
        $ProposedName = $ProposedName.TrimStart($DG.u_group_location)
    }

    If ($ProposedName -like ".*")
    {
        $ProposedName = ($ProposedName.TrimStart(" ."))
        $ProposedName = ($ProposedName.TrimEnd())
    }
 
    #
    #change the GroupName to use proper case and remove GRP and Location from the 3rd element of the name
    #
    $TextInfo = (Get-Culture).TextInfo
    $ProposedName = $ProposedName.ToLower()
    $ProposedName = $TextInfo.ToTitleCase($ProposedName)
    $GrpDescription = ($DG.u_brief_description.TrimEnd()).ToLower()
    $GrpDescription = $TextInfo.TotitleCase($GrpDescription)
    
    #Check for Known acronyms
    foreach ($Acro in $Collection)
    {
        $ProposedName = $ProposedName -Replace($Acro.Acronym,$Acro.Translation)
    }
         
    $ProposedName = "GRP." + $DG.u_group_location + "." + $ProposedName
    $DG.u_proposed_name = $ProposedName
}

function Create-Complete
{
	$UfgExists = [bool](Get-UnifiedGroup $Global:txtGrpName.Text -ResultSize Unlimited -ErrorAction SilentlyContinue)
	$First = "Y"
	Do
	{
		If ($First -eq "Y")
		{
			write-host "Waiting for Group Creation to Complete..." -ForegroundColor Cyan -NoNewline
			$First = "N"
		}
		else
		{
			write-host ".." -foregroundcolor Cyan -NoNewline
		}
		start-sleep -s 15
		$UfgExists = [bool](Get-UnifiedGroup $Global:txtGrpName.Text -ResultSize Unlimited -ErrorAction SilentlyContinue)
	} while ($UfgExists -eq $False)
    $Global:SendMsg = "Send"
} #end waiting for group creation to complete

function Create-UniGroup
{
    Build-DefaultForm
    Add-AppsChkBoxes
    Publish-Form

    If ($Global:Result -eq "OK")
    {
        $Done = "Y"
        If (($DG.u_vip -eq "True") -and ($Global:chkVIP.Checked -eq $False))
        {
            $Output = $wshell.Popup("Unable to create tNhis group you must verify this VIP has completed training.",0,"Training Complete",0+32)
            $done = read-host "Confirm that the VIP completed training (Y/N) "
        }

        If ($Done -eq "Y")
        {
            $Global:message = new-object  System.Net.Mail.MailMessage $from, $to 
            $Global:message.IsBodyHtml = $true
            $DGOwnerID = $Global:txtOwner.Text + "@global.ul.com"
            $Global:SendTo = $DGOwnerID
            $Global:GrpRetry = "N"
            $DGAlias = $Global:txtGrpName.Text -Replace '[ /,$#_-]',''
            $DGAlias = $DGAlias.Replace("\","")
            $DGAlias = $DGAlias.Replace(".","")

            If ($DGAlias.Length -gt 64)
            {
                Write-Host "Alias is too long maximum length is 64 Characters" -ForegroundColor Red
                $DGAlias = $DGAlias.Substring(0,64)
            }
            $PubPriv = "Public"
            If ($Global:chkPublic.Checked -eq $False)
            {
                $PubPriv = "Private"
            }

            $INetAddress = $DGAlias.Replace(".","") + "@ul.onmicrosoft.com"
            $OwnerMbxExists = [bool](get-mailbox $DGOwnerID)
            $UfgExists = [bool](Get-UnifiedGroup $Global:txtGrpName.Text -ResultSize Unlimited -ErrorAction SilentlyContinue)
        
            If (($UfgExists -eq $True) -and ($Global:chkPowerBI.Checked -ne $True))
            {
                $CurGroup = Get-UnifiedGroup $Global:txtGrpName.Text
                write-host "Was this group previously created (Y) to send Message Only - (N) to continue - (s) to skip (Y/N/S)? " -ForegroundColor Cyan -NoNewline
                $Global:GrpRetry = Read-Host
                If ($Global:GrpRetry -eq "N")
                {
                    $Global:MsgSubject = $CurGroup.DisplayName + " Already Exists and Has Not Been Created - Per:  " + $Global:txtTaskNo.Text
                    $GrpOwner = $CurGroup.ManagedBy
                    $NoOwners = "owner"
                    if ($CurGroup.ManagedBy.count -gt 1)
                    {
                        $GrpOwner = $CurGroup.ManagedBy -join " and "
                        $NoOwners = "owners"
                    }
            #   Set message vaules that group was not created because it already exists
                    $Global:MsgBody = $msgfont + "<p>Per your request the O365 group named <font color=green>" + $CurGroup.DisplayName + "</font> was not created because a group with that name already exists.  The current group is owned and managed by <font color=green>" + $GrpOwner + "</font>.  If you have questions or would like access to this group please contact the " + $NoOwners + " directly.</p><p>If you would like to request a new group please create a new request using the <font color=blue>'Request O365 Group'</font> form in the Service Desk portal.</p><p>Thanks,<br>The Enterprise Messaging Team</p>"
                    $Global:ActionLog = $Global:ActionLog + $Global:MsgSubject + "<br>"
                    Send-Message
                }
                If ($Global:GrpRetry -eq "Y")
                    {
                        Find-AppsUsed
                        Send-Message
                    }
            }
            else
            {
                If (($Global:chkTeams.Checked -eq $False) -and ($Global:chkAllApps.Checked -eq $False))
                {
                    If ($UfgExists -eq $False)
                    {
                        write-host "Creating New Unified Group" -ForegroundColor Green
                        New-UnifiedGroup -DisplayName $Global:txtGrpName.Text `
                            -PrimarySMTPAddress $INetAddress `
                            -Alias $DGAlias `
                            -Owner $DGOwnerID `
                            -AccessType $PubPriv `
                            | Out-Null

			            Create-Complete
    		            Set-UnifiedGroup -identity $Global:txtGrpName.Text `
                            -UnifiedGroupWelcomeMessageEnabled:$false `
                            -HiddenFromAddressListsEnabled $true `
                            -Notes ($Global:txtDescription.Text + "`nRequested by: " + (get-mailbox $DGOwnerId).Name + "`nPer: " + $Global:txtTaskNo.Text)

                    }
                    else
                    {
                        write-host "Group already exists or this is a request for PowerBI"
                    }
                }
                else
                {
                    If ($UfgExists -eq $False)
                    {
                        $TeamDesc = $Global:txtDescription.Text + "`nRequested by: " + (get-mailbox $DGOwnerId).Name + "`nPer: " + $Global:txtTaskNo.Text
                                                            
                        If ($Global:GrpRetry -eq "N")
                        {
                            #$groupname = (New-Team -DisplayName $Global:txtGrpName.Text -Description $TeamDesc -MailNickname $DGAlias -Visibility $PubPriv) # | Out-Null
                            New-Team -DisplayName $Global:txtGrpName.Text -Description $TeamDesc -MailNickname $DGAlias -Visibility $PubPriv # | Out-Null
                            Create-Complete
                        }
                                                
                        $NewTeam = Get-UnifiedGroup $Global:txtGrpName.Text
                        If ($NewTeam.WelcomeMessageEnabled -eq "True")
                        {
	                        Set-UnifiedGroup $Global:txtGrpName.Text -UnifiedGroupWelcomeMessageEnabled:$false
                        }

                        If ($NewTeam.HiddenFromAddressListsEnabled -eq "False")
                        {
                            Set-UnifiedGroup $Global:txtGrpName.Text -HiddenFromAddressListsEnabled $true
                        }

                        If (($WhoAmI.Substring($WhoAmI.length-5)) -notlike $Global:txtOwner.Text)
                        {
                            SPSite-Complete
                            $TeamGrpID = ((Get-UnifiedGroup $Global:txtGrpName.Text).ExternalDirectoryObjectID) #$groupname.groupid #
                            Add-TeamUser -GroupId $TeamGrpID -User $DGOwnerID -Role Owner

                            If ([bool](((Get-TeamUser -GroupId $TeamGrpID -role owner).count) -ge 2))
                            {
                                Remove-TeamUser -GroupId $TeamGrpID -User ($WhoAmI.Substring($whoAmI.indexof("\")+1)+"@global.ul.com") -Role Owner
                                Remove-TeamUser -GroupId $TeamGrpID -User ($WhoAmI.Substring($whoAmI.indexof("\")+1)+"@global.ul.com")
                            }
                            else
                            {
                                $Output = $wshell.Popup("There is only one owner of the group " + $Global:txtGrpName.Text,0,"One Owner",0+32)
                            }
                        }
                    }
                    else
                    {
                        If ($Global:GrpRetry -ne "S")
                        {
                            $Output = $wshell.Popup("Group " + $Global:txtGrpName.Text + " already exists.",0,"Group Exists",0+32)
                        }
                    }
                }

    	        $UfgExists = [bool]($CurrGroup = Get-UnifiedGroup $Global:txtGrpName.Text -ErrorAction SilentlyContinue)
	
                Find-AppsUsed
            
                if ($UfgExists -eq "True")
                {                
                    If ($Global:txtGrpName.Text -notlike "GRP.EXT.*")
                    {
                        New-AzureADObjectSetting -TargetType Groups -TargetObjectId $CurrGroup.ExternalDirectoryObjectId -DirectorySetting $IntSettingsCopy
                        Set-SPOSite -Identity $CurrGroup.SharepointSiteURL -SharingCapability Disabled
                    }
                    else
                    {
                        New-AzureADObjectSetting -TargetType Groups -TargetObjectId $CurrGroup.ExternalDirectoryObjectId -DirectorySetting $ExtSettingsCopy
                    }

                    If ($Global:GrpRetry -ne "S")
                    {
                        $Output = $wshell.Popup("Unified Group " + $Global:txtGrpName.Text + " created for use with " + $Global:AppsUsed,0,"Group Created",0+32)
                        $Global:ActionLog = $Global:ActionLog + $Global:MsgSubject + "<br>"
                        Send-Message
                    }
                    else
                    {
                         $Output = $wshell.Popup("Unified Group " + $Global:txtGrpName.Text + " creation skipped.",0,"Group Created",0+32)
                    }
                }	
    	        else
                {
                    If ($Global:chkPowerBI.Checked -eq $True)
                    {
                        $Output = $wshell.Popup("Unified Group " + $Global:txtGrpName.Text + " not created; reassign task to the CT Datahub Support Team.",0,"Not Created",0+32)
                        Send-Message
                    }
                    else
                    {
                        $Output = $wshell.Popup("Unified Group " + $Global:txtGrpName.Text + " not created.",0,"Not Created",0+32)
                    }
                }
            }
        }
    }
    else
    {
        $Output = $wshell.Popup("Unified Group creation cancelled.",0,"Cancelled",0+32)
    }
}

function Find-AppsUsed
{
    $Global:AppsUsed = ""

    If ($Global:chkAllApps.Checked -eq "True")
    {
        If ($Global:AppsUsed -eq "")
        {
            $Global:AppsUsed = "All Applications"
        }
    }
    else
    {
        If ($Global:chkTeams.Checked -eq "True")
        {
            $Global:AppsUsed = "Microsoft Teams"
        }

        If ($Global:chkPowerBI.Checked -eq "True")
        {
            If ($Global:AppsUsed -eq "")
            {
               $Global:AppsUsed = "Microsoft PowerBI" 
            }
            else
            {
                $Global:AppsUsed = $Global:AppsUsed + ", Microsoft PowerBI"
            }
        }

        If ($Global:chkPlanner.Checked -eq "True")
        {
            If ($Global:AppsUsed -eq "")
            {
                $Global:AppsUsed = "Microsoft Planner" 
            }
            else
            {
                $Global:AppsUsed = $Global:AppsUsed + ", Microsoft Planner"
            }
        }

        If ($Global:chkPowerApps.Checked -eq "True")
        {
            If ($Global:AppsUsed -eq "")
            {
                $Global:AppsUsed = "Microsoft PowerApps"
            }
            else
            {
                $Global:AppsUsed = $Global:AppsUsed + ", Microsoft PowerApps"
            }
        }

        If ($Global:chkPowerAutomate.Checked -eq "True")
        {
            If ($Global:AppsUsed -eq "")
            {
                $Global:AppsUsed = "Microsoft PowerAutomate"
            }
            else
            {
                $Global:AppsUsed = $Global:AppsUsed + ", Microsoft PowerAutomate"
            }
        }

        If ($Global:chkStream.Checked -eq "True")
        {
            If ($Global:AppsUsed -eq "")
            {
                $Global:AppsUsed = "Microsoft Stream"
            }
            else
            {
                $Global:AppsUsed = $Global:AppsUsed + ", Microsoft Stream"
            }
        }


    }
        Success-Details
        Success-Email
}

function Send-Message
{
    Do
	{
#        If (($Global:SendMsg -ne "Send") -and ($Global:MsgSubject -notlike "*Nightly Log*"))
#        {
#            Write-Host "Has a message previously been sent regarding the creation of this group (Y/N)? " -ForegroundColor Cyan -NoNewline
#            $Global:SendMsg = Read-Host
#        }

		If (($Global:SendMsg -eq "N") -or ($Global:SendMsg -eq "Send"))
		{
			$Global:message.To.Clear()
			$Global:message.CC.Clear()
            $Global:message.Bcc.Clear()
			$Global:message.To.Add($Global:SendTo)
	        $Global:message.Bcc.Add("EnterpriseMessagingServices@ul.com")
			$Global:message.Subject = $Global:MsgSubject
			$Global:message.Body = $Global:MsgBody
            $SMTPClient = New-Object Net.Mail.SmtpClient($Server, 25)
            #$Global:message
            #pause
            $SMTPClient.Send($Global:message)
		}
    } while (($SendMsg -ne "N") -and ($SendMsg -ne "Y") -and ($SendMsg -ne "Send"))
    $Global:message = ""
}#end Send-Message

function Send-RecapMessage
{
    #Send email to O365 Team with details of nightly actions
    $Global:SendMsg = "Send"
    $Global:message = new-object  System.Net.Mail.MailMessage $from, $to 
    $Global:message.IsBodyHtml = $true
    $Global:SendTo = "LST.O365AdminTeam@ul.com"
    #$SendTo = "Sandi.Glazebrook@ul.com"
    $Global:MsgSubject = "Unified Group Nightly Log"
    $Global:MsgBody = $msgfont + $Global:ActionLog
    write-host "`nSending Recap Message to Enterprise Messaging Services"
    Send-Message
}

function SPSite-Complete
{
    $CycleCnt = 0
    $SiteName = "https://ul.sharepoint.com/sites/" + $DGAlias #(($Global:txtGrpName.Text).Replace(" ", "")) #
    get-date |Format-List DateTime
    Write-host "Waiting for SharePoint Site" $SiteName "to be provisioned." -ForegroundColor Red -NoNewline
    start-sleep -s 15
    $ErrorActionPreference = "SilentlyContinue"

    Do
    {
        write-host ".." -ForegroundColor Red -NoNewline
        $SiteCreated = Get-SPOSite $SiteName
	    start-sleep -s 5
        $CycleCnt++
        If (($CycleCnt -eq 10) -and ($SiteCreated.Status -ne "Active"))
        {
            write-host "`nSite Creation Has not completed would you like to continue try again (Yes/No/Complete)?" -ForegroundColor Cyan -NoNewline
            $SiteRetry = read-host
            If ($SiteRetry -like "Y*")
            {
                write-host "Continuing to wait for" $SiteName "creation to complete." -ForegroundColor Red -NoNewline
                $CycleCnt = 0
            }
        }
    } while (($SiteCreated.Status -ne "Active") -and ($CycleCnt -lt 10))

    write-host ""
    get-date |Format-List DateTime
    $ErrorActionPreference = "Continue"

    If (($SiteCreated.Status -ne "Active") -or ($SiteRetry -like "N*"))
    {
        write-host "Since the Sharepoint Site has not successfully completed do not send a email to the requestor" -ForegroundColor Red
    }

}#end waiting for sharepoint site creation to complete

function Success-Details
{
    $Global:TeamPara1 = "<u>To access and manage Microsoft Teams:</b></u><p>To Access and manage your Team open the <font color=blue>'Microsoft Teams App'</font> :</p>"
    $Global:TeamPara2 = "<p>Once the Team is created you will find it in the left hand navigation pane. To manage the group click on the <font color=blue>'(…)'</font> to the right of the team name and select <font color=blue>'Manage Team'</font>.  As an owner this where you can add additional owners or members or manage other attributes related to the Microsoft Team.<br></p>"
    $Global:AppPara1 = "<p>You have requested that this group be enabled for all or multiple  Applications this could include Microsoft Teams, Microsoft Planner, PowerApps, PowerAutomate and Stream.  To access applications other than Microsoft Teams open a web browser and go to <a href='http://webmail.ul.com'>webmail.ul.com </a> this will open your email file.  Click on the <font color=blue>'App Launcher'</font> in the upper left hand corner and select <font color=blue>'All Apps'</font> to find the application you would like to use.</p>"
#    $Global:AppPara1 = "<p>You have requested that this group be enabled for all or multiple  Applications this could include Microsoft Teams, PowerBI, Microsoft Planner, PowerApps, PowerAutomate and Stream.  To access applications other than Microsoft Teams open a web browser and go to <a href='http://webmail.ul.com'>webmail.ul.com </a> this will open your email file.  Click on the <font color=blue>'App Launcher'</font> in the upper left hand corner and select <font color=blue>'All Apps'</font> to find the application you would like to use.</p>"
    $Global:ClosePara = "<p>The subject ticket is now complete and will be closed.  If you have additional questions or need additional support please contact the Service Desk.</p><p>Thanks,<br>The Enterprise Messaging Team</p>"
}

function Success-Email
{
    If ($Global:chkAllApps.Checked -eq "True")

    {
#     Set Values for message that Group has been Created for use with All Applications
        $Global:MsgSubject = $Global:txtGrpName.Text + " Has Been Created for Use with All Applications - Per:  " + $Global:txtTaskNo.Text
        $Global:OpenPara = "<p>Per your request the O365 group named <font color=green>" + $Global:txtGrpName.Text + "</font> has been created per service desk request number " + $Global:txtTaskNo.Text + ".  You have been made the owner of the group.  Due to the increase in Teams usage it may take up to 24 hours for the entire provisioning process to complete before you will be able to fully manage the group.  It is your responsibility to manage access to the team by adding and removing members from the group.  There is currently an issue with the automated email notifications when you add members to your group.  Therefore, you will need to notify any individuals you add to the group until this issue is resolved.</p><p>The following training videos are available:  <a href='https://support.office.com/en-us/article/microsoft-teams-video-training-4f108e54-240b-4351-8084-b1089f0d21d7?wt.mc_id=otc_home&ui=en-US&rs=en-US&ad=US'>TeamsTraining</a><p>"
        $Global:message.Attachments.Add($att1)
        $Global:message.Attachments.Add($att2)
        $Global:message.Attachments.Add($att3)
        $Global:MsgBody = $msgfont + $OpenPara + $Global:TeamPara1 + $Img1 + $Global:TeamPara2 + $Img2 + $Global:AppPara1 + $Img3 + $Global:ClosePara
    }
    elseif ($Global:chkTeams.Checked -eq "True")
    {
#     Set Values for message that Group has been Created for use with Microsoft Teams
        $Global:MsgSubject = $Global:txtGrpName.Text + " Has Been Created for Use with " + $Global:AppsUsed + " - Per:  " + $Global:txtTaskNo.Text
        $Global:OpenPara = "<p>Per your request the O365 group named <font color=green>" + $Global:txtGrpName.Text + "</font> has been created per service desk request number " + $Global:txtTaskNo.Text + ".  You have been made the owner of the group.  Due to the increase in Teams usage it may take up to 24 hours for the entire provisioning process to complete before you will be able to fully manage the group.  It is your responsibility to manage access to the team by adding and removing members from the group.  There is currently an issue with the automated email notifications when you add members to your group.  Therefore, you will need to notify any individuals you add to the group until this issue is resolved.</p><p>The following training videos are available:  <a href='https://support.office.com/en-us/article/microsoft-teams-video-training-4f108e54-240b-4351-8084-b1089f0d21d7?wt.mc_id=otc_home&ui=en-US&rs=en-US&ad=US'>TeamsTraining</a><p>"
        $Global:message.Attachments.Add($att1)
        $Global:message.Attachments.Add($att2)
        $Global:MsgBody = $msgfont + $OpenPara + $Global:TeamPara1 + $Img1 + $Global:TeamPara2 + $Img2 + $Global:ClosePara
    }
    elseif ($Global:chkPowerBI.Checked -eq "True")
    {
#     Set Values for message that Group has been Created for use with PowerBI
        $Global:MsgSubject = $Global:txtGrpName.Text + " request reassigned to the CT Datahub Support Team - Per:  " + $Global:txtTaskNo.Text
#        $Global:MsgSubject = $Global:txtGrpName.Text + " Has Been Created for use with " + $Global:AppsUsed + " - Per: " + $Global:txtTaskNo.Text
        $Global:OpenPara = "<p>Requests for PowerBI workspaces are now handled by the CT Datahub support team.  Your request is being reassigned accordingly.<p>"
#        $Global:OpenPara = "<p>Per your request the O365 group named <font color=green>" + $Global:txtGrpName.Text + "</font> has been created per service desk request number " + $Global:txtTaskNo.Text + ".  It may take up to 24 hours for the entire provisioning process to complete before you will be able to fully manage the group.</p><p>You have been made an owner of the group and can add additional members or owners to the group using the <font color=green>Edit Workspace</font> feature in PowerBI.</p><p>If you specified that you are sharing with more than 10 individuals the group will be added to the PowerBI Premium Capacity.  Once the group has been added into the Premium Capacity you will see a <font color=green>Diamond</font> displayed after the group name.  Data published to a group that is a part of the Premium Capacity will <b>not</b> require the individuals who are viewing the information to obtain a PowerBI Pro license.</p><p>The following training videos are available:  <a href='https://support.office.com/office-training-center?redirectSourcePath=%252farticle%252fb8f02f81-ec85-4493-a39b-4c48e6bc4bfb'>O365TrainingCenter</a><p>"
#        $Global:MsgBody = $msgfont + $OpenPara + $Global:ClosePara
        $Global:MsgBody = $msgfont + $OpenPara + "<p>Thanks,<br>The Enterprise Messaging Team</p>"
    }
    elseif (($Global:chkPowerBI.Checked -ne "True") -and ($Global:chkTeams.Checked -ne "True") -and ($Global:chkAllApps.Checked -ne "True"))
    {
        $Global:MsgSubject = $Global:txtGrpName.Text + " Has Been Created for Use with " + $Global:AppsUsed + " - Per:  " + $Global:txtTaskNo.Text
        $Global:OpenPara = "<p>Per your request the O365 group named <font color=green>" + $Global:txtGrpName.Text + "</font> has been created per service desk request number " + $Global:txtTaskNo.Text + ".  You have been made the owner of the group.  It may take up to 24 hours for the entire provisioning process to complete before you will be able to fully manage the group.  It is your responsibility to manage the membship of the group by adding and removing individuals from the group.</p><p>The following training videos are available:  <a href='https://support.office.com/office-training-center?redirectSourcePath=%252farticle%252fb8f02f81-ec85-4493-a39b-4c48e6bc4bfb'>O365TrainingCenter</a><p>"
        $Global:MsgBody = $msgfont + $OpenPara + $Global:ClosePara
    }
}


############################################################
#
write-host "New Unified Group " -ForegroundColor Magenta
$DG = ""

$Type = read-host "Create internal or external teams (I/E) "

If ($Type -eq "E")
{
    $InputFile = "e:\Automation\NewUnifiedGroup\Input\Input-NewExtTeam.csv"
}
else
{
    $InputFile = "e:\Automation\NewUnifiedGroup\Input\Input-NewUnifiedGroup.csv"
}

$Global:ActionLog = "<p>O365 Unified Groups Automated Processing Details</p>"
$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server
$from = New-Object System.Net.Mail.MailAddress "DoNotReply@ul.com" , "Enterprise Messaging Services"
$to = $from 
$msgfont = "<basefont face=verdana size=2.5 color=black>"
$Tab = "<p style=""margin-left: 40px"">"

#Images
    $image1 = "e:\O365AdminShared\Data\TeamApp.png"
    $att1 = new-object Net.Mail.Attachment($Image1)
    $att1.ContentType.MediaType = “image/png”
    $att1.ContentId = “Attachment1”
    $Img1 = $Tab + "<img src='cid:$($att1.ContentId)'"
    $image2 = "e:\O365AdminShared\Data\ManageTeam.png"
    $att2 = new-object Net.Mail.Attachment($Image2)
    $att2.ContentType.MediaType = “image/png”
    $att2.ContentId = “Attachment2”
    $Img2 = $Tab + "<img src='cid:$($att2.ContentId)'"
    $image3 = "e:\O365AdminShared\Data\AppLauncher.png"
    $att3 = new-object Net.Mail.Attachment($Image3)
    $att3.ContentType.MediaType = “image/png”
    $att3.ContentId = “Attachment3”
    $Img3 = $Tab + "<img src='cid:$($att3.ContentId)'"

# Get External Groups Templates
    $template = Get-AzureADDirectorySettingTemplate | ? {$_.displayname -eq "group.unified.guest"}
    $IntSettingsCopy = $template.CreateDirectorySetting()
    $IntSettingsCopy["AllowToAddGuests"] = $False
    $ExtSettingsCopy = $template.CreateDirectorySetting()
    $ExtSettingsCopy["AllowToAddGuests"] = $True

# Begin Unified Group creation

if (Test-Path $InputFile)
{	
    # test the date of the file.....if it is more than 24 hrs old it is an old file
    $UnifDistGroup = Import-CSV $InputFile
     
    $Global:ProcCount = $UnifDistGroup.Count
    If ($Global:ProcCount -le 0)
    {
	    $Global:ProcCount = 1
	}

    write-host "Processing" $ProcCount "requests for New O365 Groups" -foregroundcolor Cyan
    $Acro = Import-Csv e:\Automation\NewUnifiedGroup\Input\Input-KnownAcronyms.csv
    $collection = $Acro.GetEnumerator()

  	ForEach ($DG in $UnifDistGroup)
    {
        $Global:message = new-object  System.Net.Mail.MailMessage $from, $to 
        $Global:message.IsBodyHtml = $true

        write-host "`nProcessing started for group " $DG.u_proposed_name
        If ($DG.u_owner_empid.length -lt 5)
        {
            Do
            {
                $DG.u_owner_empid = "0" + $DG.u_owner_empid
            } while ($DG.u_owner_empid.length -lt 5)
        }
        $DGOwnerID = $DG.u_owner_empid + "@global.ul.com"
        $OwnerMbxExists = [bool](get-mailbox $DGOwnerID -ErrorAction SilentlyContinue)
            
        If ($OwnerMbxExists -eq "True")
        {
            If (($DG.u_group_name -like "*.*") -and ($Global:chkPowerBI.Checked -eq $False))
            {
                Check-GroupName
            }
            Create-UniGroup
        }
        else
        {
            $Output = $wshell.Popup("Unified Group " + $DG.u_proposed_name + " cannot be created request owner is not a licensed mail user.",0,"Owner Not Licensed",0+32)
            $Global:ActionLog = $Global:ActionLog  + $DG.u_proposed_name + " Cannot be created Owner Not a Licensed Email User - Per:  " + $DG.u_requested_item + "<br>" 
        }
    }

    #Rename the Input file when finished processing
    if (Test-Path $InputFile)
    {
        $loop = 0
        Do
        {
            Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log") -ErrorAction SilentlyContinue
            $loop++
        } while ((Test-Path $InputFile) -and ($loop -lt 10))
        
        If ($loop -ge 10)
        {
            Write-Host "Unable to rename this file" $InputFule "please manually rename it to" ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
        }
    }
}
else
{
    Create-UniGroup
}

If ($Global:Result -eq "OK")
{
    Send-RecapMessage
}