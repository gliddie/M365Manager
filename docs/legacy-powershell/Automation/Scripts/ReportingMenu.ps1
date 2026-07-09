function Build-ReportAdminMenuForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "O365 Reporting Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 600 ; $form.Height = 400  # Make the form wider 
    
    Add-FormStandardButtons

    $TopLoc = 20 

    ## Label and TextBox  
    ## Title Line
    $Global:lblTitleLine = New-Object System.Windows.Forms.Label   
        $lblTitleLine.Text = "Select Option:"
        $lblTitleLine.Top = 15 ; $lblTitleLine.Left = 60; $lblTitleLine.Width=120 ;$lblTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine)    # Add to Form 

    ## Report on User Mailboxes
    $TopLoc = $TopLoc + 20
    $Global:chkUsrMbx = New-Object Windows.Forms.RadioButton 
        $Global:chkUsrMbx.Left = 100; $Global:chkUsrMbx.Width = 450; $Global:chkUsrMbx.Top = $TopLoc  
        $Global:chkUsrMbx.Text = "Report on All User Mailboxes" 
        $Global:chkUsrMbx.Checked = $false   # set a default value 
        $Global:chkUsrMbx.TabIndex = 1
        $Global:form.Controls.Add($Global:chkUsrMbx) 
        # Obtain Value with: $Global:chkUsrMbx.Checked

    ## Report on Shared Mailboxes
    $TopLoc = $TopLoc + 20
    $Global:chkShrmbx = New-Object Windows.Forms.RadioButton 
        $Global:chkShrmbx.Left = 100; $Global:chkShrmbx.Width = 450; $Global:chkShrmbx.Top = $TopLoc  
        $Global:chkShrmbx.Text = "Report on All Shared Mailboxes" 
        $Global:chkShrmbx.Checked = $false   # set a default value 
        $Global:chkShrmbx.TabIndex = 2
        $Global:form.Controls.Add($Global:chkShrmbx) 
        # Obtain Value with: $Global:chkShrmbx.Checked
        
    ## Report Conference Room Mailboxes
    $TopLoc = $TopLoc + 20
    $Global:chkConfRooms = New-Object Windows.Forms.RadioButton 
        $Global:chkConfRooms.Left = 100; $Global:chkConfRooms.Width = 450; $Global:chkConfRooms.Top = $TopLoc  
        $Global:chkConfRooms.Text = "Report on All Conference Room Malboxes" 
        $Global:chkConfRooms.Checked = $false   # set a default value 
        $Global:chkConfRooms.TabIndex = 3
        $Global:form.Controls.Add($Global:chkConfRooms) 
        # Obtain Value with: $Global:chkConfRooms.Checked

    ## Report Equipment Mailbox
    $TopLoc = $TopLoc + 20
    $Global:chkEquipMbx = New-Object Windows.Forms.RadioButton 
        $Global:chkEquipMbx.Left = 100; $Global:chkEquipMbx.Width = 450; $Global:chkEquipMbx.Top = $TopLoc  
        $Global:chkEquipMbx.Text = "Report on All Equipment Mailboxes" 
        $Global:chkEquipMbx.Checked = $false   # set a default value 
        $Global:chkEquipMbx.TabIndex = 4
        $Global:form.Controls.Add($Global:chkEquipMbx) 
        # Obtain Value with: $Global:chkEquipMbx.Checked
        
    ## Report Owners of DL, Security and Dynamic Groups
    $TopLoc = $TopLoc + 20
    $Global:chkDLRpt = New-Object Windows.Forms.RadioButton 
        $Global:chkDLRpt.Left = 100; $Global:chkDLRpt.Width = 450; $Global:chkDLRpt.Top = $TopLoc  
        $Global:chkDLRpt.Text = "Report Owners of Distribution, Security and Dynamic Groups" 
        $Global:chkDLRpt.Checked = $false   # set a default value 
        $Global:chkDLRpt.TabIndex = 5
        $Global:form.Controls.Add($Global:chkDLRpt) 
        # Obtain Value with: $Global:chkDLRpt.Checked

    ## NoOwners on DL, Security and Dynamic Groups
    $TopLoc = $TopLoc + 20
    $Global:chkNoOwnerRpt = New-Object Windows.Forms.RadioButton 
        $Global:chkNoOwnerRpt.Left = 100; $Global:chkNoOwnerRpt.Width = 450; $Global:chkNoOwnerRpt.Top = $TopLoc  
        $Global:chkNoOwnerRpt.Text = "Report of No Owners of Distribution, Security and Dynamic Groups" 
        $Global:chkNoOwnerRpt.Checked = $Global:chkNoOwnerRpt.Checked   # set a default value 
        $Global:chkNoOwnerRpt.TabIndex = 6
        $Global:form.Controls.Add($Global:chkNoOwnerRpt) 
        # Obtain Value with: $Global:chkNoOwnerRpt.Checked

    ## DST Group counts
    $TopLoc = $TopLoc + 20
    $Global:chkDSTCounts = New-Object Windows.Forms.RadioButton 
        $Global:chkDSTCounts.Left = 100; $Global:chkDSTCounts.Width = 450; $Global:chkDSTCounts.Top = $TopLoc  
        $Global:chkDSTCounts.Text = "Report of DST Group Membership Counts" 
        $Global:chkDSTCounts.Checked = $false   # set a default value 
        $Global:chkDSTCounts.TabIndex = 7
        $Global:form.Controls.Add($Global:chkDSTCounts) 
        # Obtain Value with: $Global:chkDSTCounts.Checked

    ## UM Extensions
    $TopLoc = $TopLoc + 20
    $Global:chkUMExt = New-Object Windows.Forms.RadioButton 
        $Global:chkUMExt.Left = 100; $Global:chkUMExt.Width = 450; $Global:chkUMExt.Top = $TopLoc  
        $Global:chkUMExt.Text = "Report of All UM Extensions" 
        $Global:chkUMExt.Checked = $false   # set a default value 
        $Global:chkUMExt.TabIndex = 8
        $Global:form.Controls.Add($Global:chkUMExt) 
        # Obtain Value with: $Global:chkUMExt.Checked

    ## User Mailbox No Default Role Assignment
    $TopLoc = $TopLoc + 20
    $Global:chkNoDefPolicy = New-Object Windows.Forms.RadioButton 
        $Global:chkNoDefPolicy.Left = 100; $Global:chkNoDefPolicy.Width = 450; $Global:chkNoDefPolicy.Top = $TopLoc
        $Global:chkNoDefPolicy.Text = "Report of User Mailboxes Not Assigned to the UL Default Role Assignment Policy" 
        $Global:chkNoDefPolicy.Checked = $false   # set a default value 
        $Global:chkNoDefPolicy.TabIndex = 9
        $Global:form.Controls.Add($Global:chkNoDefPolicy) 
        # Obtain Value with: $Global:chkNoDefPolicy.Checked

    ## User Retention Policy
    $TopLoc = $TopLoc + 20
    $Global:chkUsrRetPol = New-Object Windows.Forms.RadioButton 
        $Global:chkUsrRetPol.Left = 100; $Global:chkUsrRetPol.Width = 450; $Global:chkUsrRetPol.Top = $TopLoc  
        $Global:chkUsrRetPol.Text = "Report of User Mailboxes and Mailbox Retention Policy Settings" 
        $Global:chkUsrRetPol.Checked = $false   # set a default value 
        $Global:chkUsrRetPol.TabIndex = 10
        $Global:form.Controls.Add($Global:chkUsrRetPol) 
        # Obtain Value with: $Global:chkUsrRetPol.Checked
 
    ## Compare OED Locaitons
    $TopLoc = $TopLoc + 20
    $Global:chkOEDComp = New-Object Windows.Forms.RadioButton 
        $Global:chkOEDComp.Left = 100; $Global:chkOEDComp.Width = 450; $Global:chkOEDComp.Top = $TopLoc  
        $Global:chkOEDComp.Text = "Report to Compare OED Locations and DST.All Groups" 
        $Global:chkOEDComp.Checked = $false   # set a default value 
        $Global:chkOEDComp.TabIndex = 11
        $Global:form.Controls.Add($Global:chkOEDComp) 
        # Obtain Value with: $Global:chkOEDComp.Checked

    ## Hidden Mailboxes
    $TopLoc = $TopLoc + 20
    $Global:chkHiddenMbx = New-Object Windows.Forms.RadioButton 
        $Global:chkHiddenMbx.Left = 100; $Global:chkHiddenMbx.Width = 450; $Global:chkHiddenMbx.Top = $TopLoc  
        $Global:chkHiddenMbx.Text = "Report of Mailboxes Hidden from the Address Book" 
        $Global:chkHiddenMbx.Checked = $false   # set a default value 
        $Global:chkHiddenMbx.TabIndex = 12
        $Global:form.Controls.Add($Global:chkHiddenMbx) 
        # Obtain Value with: $Global:chkHiddenMbx.Checked
}

Function Publish-MMForm
{
    ## Finalize Form and Show Dialog 
    $Global:form.Add_Shown( { $form.Activate(); $Global:okButton.Focus() } )  #Activate and Set Focus 
    $Global:MainResult = $Global:form.ShowDialog()          ## Show the form, and wait for the response
}

#### Start of Script

Build-ReportAdminMenuForm
Publish-Form

If ($Global:Result -eq "OK")
{
    If ($Global:chkUsrMbx.Checked -eq "Checked")
    {
        invoke-Expression -Command .\MbxUserReport_wMailNotification.ps1
    }

    If ($Global:chkShrMbx.Checked -eq "Checked")
    {
        invoke-Expression -Command .\MbxSharedReport_wMailNotification.ps1
    }

    If ($Global:chkConfRooms.Checked -eq "Checked")
    {
       invoke-Expression -Command .\MbxRoomReport.ps1
    }

    If ($Global:chkEquipMbx.Checked -eq "Checked")
    {
       invoke-Expression -Command .\MbxResourceReport.ps1
    }

    If ($Global:chkDLRpt.Checked -eq "Checked")
    {
       write-host "This process complies a list of Distribution Lists and Dynamic Distribution Lists and their Owners"
       invoke-Expression -command .\DistributionGroupOwnershipReport.ps1
    }

    If ($Global:chkNoOwnerRpt.Checked -eq "Checked")
    {
        #Run Report on All Distribution Lists with No Owners
        $NoOwn = "ListName,ManagedBy,MembershipCount,ListNotes," + "`n"
		$NoOwnNoMem = "ListName,ManagedBy,MembershipCount,ListNotes" + "`n"
        $DstLstNoOwner = Get-DistributionGroup -ResultSize Unlimited | Where-Object{$_.ManagedBy.Count -lt 2}
		write-Host "Total Number of Distribution Groups with No Owners: " $DstLstNoOwner.count
		foreach ($DstLstNoOwner in $DstLstNoOwner)
		{
			$grp = get-group $DstLstNoOwner.Alias
			$mem = get-distributiongroupmember $DstLstNoOwner.Alias -ResultSize Unlimited
			$text = "{0},{1},{2},{3}" -f $DstLstNoOwner.Name, $DstLstNoOwner.ManagedBy.Count, $mem.Name.count, $grp.Notes
			If ($Mem.Count -eq 0)
			{
			    $NoOwnNoMem = $NoOwnNoMem + $text + "`n"
			}
			else
			{
				$NoOwn = $NoOwn + $text + "`n"
			}
		}
		write-host "Writing Groups with No Owners or Members to c:\temp\DistGroupsNoOwnersNoMembers.csv"
		$NoOwnNoMem > c:\temp\DistGroupsNoOwnersNoMembers.csv
		write-host ""
		write-host "Writing Groups with No Owners to c:\temp\DistGroupsNoOwners.csv"
		$NoOwn > c:\temp\DistGroupsNoOwners.csv
    	write-host "Unified Groups with No Members or Owners"
		Get-UnifiedGroup -resultsize unlimited | Where-Object {($_.Managedby.Count -eq 0) -and ($_.GroupMemberCount -eq 0)} |ft DisplayName,Managedby,AccessType
#		Report of Unified Groups with No Owner but Active Members
		write-host "Unified Groups with Members but no Owner"
		Get-UnifiedGroup -resultsize unlimited | Where-Object {($_.Managedby.Count -eq 0) -and ($_.GroupMemberCount -gt 0)} |ft DisplayName,Managedby,AccessType
    }

    If ($Global:chkDSTCounts.Checked -eq "Checked")
    {
#       Run Report on All DST Groups with No Members
		$DynOutPath = "C:\temp\O365_DynGroupCounts.csv"
		$DLOutPath = "C:\temp\O365_DLOwnershipReport.csv"

		$DLtext = "DisplayName,DLName,ManagedBy,NoOfMembers,Notes,LastChanged,DLType"
		$Dyntext = "DisplayName,DLName,NoOfMembers,DLType"

		Out-File -FilePath $DynOutPath -InputObject $Dyntext
		
        $DynDLs = Get-DynamicDistributionGroup -ResultSize Unlimited
		write-Host "Total Number of Dynamic Distribution Groups: " $DynDLs.count
		$DynDLsNotes = "Dynamic Group"
		foreach ($DynDLs in $DynDLs)
		{
			$DynDLs.Name
			$g = Get-DynamicDistributionGroup $DynDLs.Name
#			$DLMem = (Get-Recipient -RecipientPreviewFilter $g.LdapRecipientFilter -ResultSize Unlimited).count
#  			$DLMem = (Get-Recipient -RecipientPreviewFilter $g.RecipientFilter -OrganizationalUnit $g.RecipientContainer -ResultSize Unlimited).count
            $DLMem = Get-DynamicDistributionGroupMember $DynDLs.Name -ResultSize unlimited

			if (($DLMem.count -ne 0) -and ($DLMem.count -lt 1))
			{
				$DLMem = 1
			}
						
			write-host $DynDls.Name "," $DLMem.count
            $Dyntext = """{0}"",""{1}"",{2},{3}" -f $DynDLs.DisplayName,$DynDLs.Name,$DLMem.count,$DynDLs.RecipientTypeDetails
			$Dyntext
			Out-File -FilePath $DynOutPath -InputObject $Dyntext -Append -NoClobber
		}

		Out-File -FilePath $DynOutPath -InputObject $Dyntext -Append -NoClobber
				
#		$DstLstNoOwner.Name,$DstLstNoOwner.Notes
	}

    If ($Global:chkUMMbx.Checked -eq "Checked")
    {
        Write-Host "Creating Report on c:\temp\UMReport.csv" -ForegroundColor Yellow
        $strPath = "C:\temp\UMReport.csv"
        $text = "EmpEmailAddress,Extension,DialPlan"
        Out-File -FilePath $strPath -InputObject $text
        $UMMbx = Get-UMMailbox -ResultSize Unlimited
        foreach ($UMMbx in $UMMbx)
        {
            $text = "{0},{1},{2}" -f $UMMbx.PrimarySMTPAddress,$UMMbx.PhoneNumber,$UMMbx.UMDialPlan
            Out-File -FilePath $strPath -InputObject $text -Append
        }
    }

    If ($Global:chkNoDefPolicy.Checked -eq "Checked")
    {
		write-host "Gathering User Mailboxes not Assigned to the UL Default Role Assignment Policy"
		write-host "This policy determines what features the user is able to manage using Webmail and their O365 client"
		write-host "Items controled in this policy are things like not being able to add Forwarders and the ability to "
		write-host "edit distribution group membership that they own."
		write-host ""
		$MbxPolicy = get-mailbox -ResultSize Unlimited | where {($_.RoleAssignmentPolicy -ne "UL Default Role Assignment Policy") -and ($_.RecipientTypeDetails -eq "UserMailbox")}
		write-host "Number of UserMailboxes not assigned to the correct RoleAssignmentPolicy: " $MbxPolicy.count
		write-host $MbxPolicy
	}

    If ($Global:chkUsrRetPol.Checked -eq "Checked")
	{
		write-host "Gathering User Mailboxes not Assigned to the UL Default MRM Policy - 365 Day Delete Policy or"
		write-host """RetentionHoleEnabled"" is not set to True or the ""StartDate"" for the retention hold is not set" 
		write-host "to 3/31/2011 07:00:00 PM"
		write-host ""
    	write-host "These policies set the retention time and prevents email from being purged until approval is"
		write-host "obtained by Legal/Security to enable automatic purging of email"
		write-host ""
#		$MbxRetent = get-mailbox -ResultSize Unlimited | where {($_.RecipientTypeDetails -eq "UserMailbox") -and ($_.Alias -notlike "SVC*") -and (($_.RetentionPolicy -ne "UL Default MRM Policy - 365 Day Delete") -or ($_.RetentionHoldEnabled -ne "True") -or ($_.StartDateForRetentionHold -ne "3/31/2011 7:00:00 PM"))}
		$MbxRetent = get-mailbox -ResultSize Unlimited | where {($_.RecipientTypeDetails -eq "UserMailbox") -and (($_.RetentionPolicy -ne "UL Default MRM Policy - 365 Day Delete") -or ($_.RetentionHoldEnabled -ne "True") -or ($_.StartDateForRetentionHold -ne "3/31/2011 7:00:00 PM"))}
			
		write-host "Number of UserMailboxes not assigned to the correct Mailbox Retention Policy: " $MbxRetent.count
		
		pause
			
		foreach ($MbxRetent in $MbxRetent)
		{
			write-host $MbxRetent.Alias,$MbxRetent.RetentionPolicy,$MbxRetent.RetentionHoldEnabled,$MbxRetent.StartDateForRetentionHold
		}
    }
			
    If ($Global:chkOEDComp.Checked -eq "Checked")
    {
        Write-Host "Running report of OED Locations and DST.All UL Employees Only and UL Staff Groups"
        invoke-Expression -Command .\CheckOEDLocations.ps1
    }

    If ($Global:chkHiddenMbx.Checked -eq "Checked")
    {
        write-host "Running report of all Mailboxes hiddent from the Address List"
        write-host "Results of this report are written to c:\temp\HiddenFromAddressList.csv"
        Get-Mailbox -ResultSize Unlimited | Where {$_.HiddenFromAddressListsEnabled -eq $True} | Select Name, Alias, HiddenFromAddressListsEnabled | Export-Csv c:\temp\HiddenFromAddressBook.csv
    }
}