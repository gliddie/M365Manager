<#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Creates new Universal Security Groups
#    'Called By    : RoomResourceDelegateMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 09/30/2020
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 09/30/2020
#    '             :    Created from NewRoomAdmin.ps1
#    '
# ==========================================================================
#
#################################################################################>

# Begin Room Admin List creation
# Create the Room Admin List

    New-DistributionGroup -Name $Global:txtDeleName.Text `
		-PrimarySmtpAddress ($Global:txtDeleName.Text + "@ul.com") `
		-Alias $Global:txtDeleName.Text `
		-ManagedBy "MBX.RRS.Owner" `
		-Type Security `
		| Out-Null

	Set-DistributionGroup $Global:txtDeleName.Text -HiddenFromAddressListsEnabled $true

	if (Get-DistributionGroup $Global:txtDeleName.Text)
    {
		write-host "Room Admin List created: " $Global:txtDeleName.Text " (" $Global:txtDeleName.Text +  "@ul.com)"
		$RoomAdmin	= Get-DistributionGroup ($Global:txtDeleName.Text +  "@ul.com")
	
		Set-DistributionGroup $RoomAdmin.PrimarySMTPAddress `
 			-RequireSenderAuthenticationEnabled $false `
 			-BypassSecurityGroupManagerCheck `
			-CustomAttribute15 ($FileName + " PS Date: " + (get-date -uformat %D) + " PS Time: " + (get-date -uformat %T))
				
		# Add Room Admin List members
		write-host ("Adding Members to Room Admin List: " + $RoomAdmin.Name + " (" + $RoomAdmin.PrimarySMTPAddress + ")")
				
    	$addMember = $Global:txtDeleMgrs.Text.split(",")
		if ($? -eq $true)
        {
			ForEach ($member in $addMember)
            {
				if ($member.contains("@"))
                {
				    if (Get-Mailbox $member)
                    {
						Add-DistributionGroupMember $RoomAdmin.PrimarySMTPAddress -Member $member -BypassSecurityGroupManagerCheck
						write-host "Added member:" $member
	    			}
					else
                    {
						write-host "ERROR finding member: " $member
					}
				}
				else
                {
					write-host "ERROR invalid member: " $member
				}
			}
		}	
		else
        {
			write-host "No members to add"
    	}
	}	
	else
    {
		Write-Host "Room Admin List not created: " $DG.Name " (" $DG.Mail ")"
	}
# =============================================================================================================================================