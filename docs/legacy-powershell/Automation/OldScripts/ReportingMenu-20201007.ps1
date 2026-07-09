<####
#### Reporting Menu
####

#  Called from O365 Admin Menu to perform Licensing Activites
#
#  Added Option 11 to Check the OED Locations and DST.All Groups
#>

$RepChg = "1"

Do
{
    write-host ""
    write-host "Reporting Menu" -ForegroundColor Magenta
    write-host
    Write-Host "     Enter ( 1) Run Report on All User Mailboxes"
    write-host "           ( 2) Run Report on All Shared Mailboxes"
    Write-Host "           ( 3) Run Report on All Conference Room Mailboxes"
    Write-Host "           ( 4) Run Report on All Equipment Mailboxes"
    write-host "           ( 5) Run Report Owners of Distribution, Security, Dynamic Lists"
    write-host "           ( 6) Run Report of No Owners of Distribution, Security, Dynamic Lists"
	write-host "           ( 7) Run Report of DST Groups Membership Counts"
    Write-Host "           ( 8) Run Report of all UM Extensions"
	Write-Host "           ( 9) Run Report of User Mailboxes Not Assigned to the UL Default Role Assignment Policy"
	Write-Host "           (10) Run Report of User Mailboxes and the Mailbox Retention Policy Settings"
    write-host "           (11) Run Report to Compare OED Locations and DST.All Groups"
    write-host "           (12) Run Report of Mailboxes Hidden from the Address Book"
    write-host ""
    write-Host "           ( 0) to Return to the O365 Admin Menu"
    write-Host "     Enter Option? " -ForegroundColor Red -NoNewline
    $RepChg = Read-Host
    write-host ""
    
    switch ($RepChg)
    {
        1
            {
				invoke-Expression -Command .\MbxUserReport_wMailNotification.ps1
            }
        2
            {
                invoke-Expression -Command .\MbxSharedReport_wMailNotification.ps1
            }
        3
            {
                invoke-Expression -Command .\MbxRoomReport.ps1
            }
        4
            {
                invoke-Expression -Command .\MbxResourceReport.ps1
            }
        5
            {
                write-host "This process complies a list of Distribution Lists and Dynamic Distribution Lists and their Owners"
				invoke-Expression -command .\DistributionGroupOwnershipReport.ps1
            }
        6
            {
#               Run Report on All Distribution Lists with No Owners
				$NoOwn = "ListName,ManagedBy,MembershipCount,ListNotes," + "`n"
				$NoOwnNoMem = "ListName,ManagedBy,MembershipCount,ListNotes" + "`n"
                $DstLstNoOwner = Get-DistributionGroup -ResultSize Unlimited | Where-Object{$_.ManagedBy.Count -lt 1}
				write-Host "Total Number of Distribution Groups with No Owners: " $DstLstNoOwner.count
				foreach ($DstLstNoOwner in $DstLstNoOwner)
				{
					$grp = get-group $DstLstNoOwner.Alias
					$mem = get-distributiongroupmember $DstLstNoOwner.Alias -ResultSize Unlimited
					$text = "{0},{1},{2},{3}" -f $DstLstNoOwner.Name, $DstLstNoOwner.ManagedBy, $mem.Name.count, $grp.Notes
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
#				Report of Unified Groups with No Owner but Active Members
				write-host "Unified Groups with Members but no Owner"
				Get-UnifiedGroup -resultsize unlimited | Where-Object {($_.Managedby.Count -eq 0) -and ($_.GroupMemberCount -gt 0)} |ft DisplayName,Managedby,AccessType
            }
		7
			{
#               Run Report on All DST Groups with No Members
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
					$DLMem = (Get-Recipient -RecipientPreviewFilter $g.LdapRecipientFilter -ResultSize Unlimited).count

					if (($DLMem -ne 0) -and ($DLMem -lt 1))
					{
						$DLMem = 1
					}
						
					$Dyntext = """{0}"",""{1}"",{2},{3}" -f $DynDLs.DisplayName,$DynDLs.Name,$DLMem,$DynDLs.RecipientTypeDetails
					$Dyntext
					Out-File -FilePath $DynOutPath -InputObject $Dyntext -Append -NoClobber
				}

				Out-File -FilePath $DynOutPath -InputObject $Dyntext -Append -NoClobber
				
#				$DstLstNoOwner.Name,$DstLstNoOwner.Notes
			}
		8
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
		9
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
		10
			{
				write-host "Gathering User Mailboxes not Assigned to the UL Default MRM Policy - 365 Day Delete Policy or"
				write-host """RetentionHoleEnabled"" is not set to True or the ""StartDate"" for the retention hold is not set" 
				write-host "to 3/31/2011 07:00:00 PM"
				write-host ""
				write-host "These policies set the retention time and prevents email from being purged until approval is"
				write-host "obtained by Legal/Security to enable automatic purging of email"
				write-host ""
#				$MbxRetent = get-mailbox -ResultSize Unlimited | where {($_.RecipientTypeDetails -eq "UserMailbox") -and ($_.Alias -notlike "SVC*") -and (($_.RetentionPolicy -ne "UL Default MRM Policy - 365 Day Delete") -or ($_.RetentionHoldEnabled -ne "True") -or ($_.StartDateForRetentionHold -ne "3/31/2011 7:00:00 PM"))}
				$MbxRetent = get-mailbox -ResultSize Unlimited | where {($_.RecipientTypeDetails -eq "UserMailbox") -and (($_.RetentionPolicy -ne "UL Default MRM Policy - 365 Day Delete") -or ($_.RetentionHoldEnabled -ne "True") -or ($_.StartDateForRetentionHold -ne "3/31/2011 7:00:00 PM"))}
				
				write-host "Number of UserMailboxes not assigned to the correct Mailbox Retention Policy: " $MbxRetent.count
				
				pause
				
				foreach ($MbxRetent in $MbxRetent)
				{
					write-host $MbxRetent.Alias,$MbxRetent.RetentionPolicy,$MbxRetent.RetentionHoldEnabled,$MbxRetent.StartDateForRetentionHold
				}
			
			}			
        11
            {
                Write-Host "Running report of OED Locations and DST.All UL Employees Only and UL Staff Groups"
                invoke-Expression -Command .\CheckOEDLocations.ps1
            }
        12
            {
                write-host "Running report of all Mailboxes hiddent from the Address List"
                write-host "Results of this report are written to c:\temp\HiddenFromAddressList.csv"
                Get-Mailbox -ResultSize Unlimited | Where {$_.HiddenFromAddressListsEnabled -eq $True} | Select Name, Alias, HiddenFromAddressListsEnabled | Export-Csv c:\temp\HiddenFromAddressBook.csv
            }
        }

}While ($RepChg -ne 0)