<#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Creates new Universal Security Groups
#    'Called By    : RoomResourceAdminMenu.ps1
#    'Calls        :
#    'Author       : Sandi Glazebrook
#    'Date Created : 11/11/2020
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 11/11/2020 Created to use forms
#    '
# ==========================================================================
#
#################################################################################>

function Add-Members
{
	if ($? -eq $true)
    {
    	ForEach ($member in $addMember)
        {
		    if ($member.contains("@") -or ($member.length -eq 5))
            {
    			if (Get-Mailbox $member.Trim())
                {
				    Add-DistributionGroupMember $GrpName -Member $member.Trim() -BypassSecurityGroupManagerCheck
    				write-host "Added member:" $member.Trim()
	    			$LineToWrite = "ADD" + "`t" + "     Added " + $member.Trim() + " to Distribution Group " + $GrpName + "`n"
		    		WriteReportEvent
			    }
    			else
                {
				    write-host "ERROR finding member: " $member.Trim()
    				$LineToWrite = "FAIL" + "`t" + "     ERROR finding recipient " + $member.Trim() + " unable to add to" + $GrpName + "`n"
	    			WriteReportEvent
		    	}
    		}
	    	else
            {
			    write-host "ERROR invalid member: " $member.Trim()
    			$LineToWrite = "ERROR" + "`t" + "     Invalid recipient " + $member.Trim() + "`n"
	    		WriteReportEvent
		    }
    	}
    }
    else
    {
		write-host "No members to add"
	    $LineToWrite = "ADD" + "`t" + "     No members to add to Distribution Group " + $GrpName + "`n"
		WriteReportEvent
	}
}

####################################################################
#    Create Room Group, Delegates Users Groups Exist
####################################################################

#  Check to see of the Room List Exists
$Global:NewRoomList = "N"
If ($Global:chkNewGRRoom.Checked -eq "Checked")
{
    $Exists = [bool](Get-DistributionGroup $Global:txtInpRoomGrp.Text -ErrorAction SilentlyContinue)
    If ($Exists -ne $True)
    {
        New-DistributionGroup -Name $Global:txtInpRoomGrp.Text -RoomList -ManagedBy "MBX.RRS.Owner"
        $LineToWrite = "NEW" + "`t" + "     Create Room List: " + $Global:txtInpRoomGrp.Text
        WriteReportEvent
        $Global:NewRoomList = "Y"
    }
    else
    {
        $LineToWrite = "EXISTS" + "`t" + "     Room List Already Exists: " + $Global:txtInpRoomGrp.Text
        WriteReportEvent    
    }
}

#  Check to see if the Rooom Delegates group exists
$Exists = [bool](Get-DistributionGroup $Global:txtDeleName.Text -ErrorAction SilentlyContinue)
If ($Exists -ne $True)
{
   	# Create the Room Delegates Group
	New-DistributionGroup -Name $Global:txtDeleName.Text `
		-PrimarySmtpAddress (($Global:txtDeleName.Text) + "@ul.com") `
		-Alias $Global:txtDeleName.Text `
		-ManagedBy "MBX.RRS.Owner" `
		-Type Security `
		| Out-Null
    
    $LineToWrite = "NEW" + "`t" + "     Create Delegates Group: " + $Global:txtDeleName.Text
    WriteReportEvent

	Set-DistributionGroup $Global:txtDeleName.Text `
        -HiddenFromAddressListsEnabled $true `
        -RequireSenderAuthenticationEnabled $false `
		-BypassSecurityGroupManagerCheck `
		-CustomAttribute15 ($FileName + " PS Date: " + (get-date -uformat %D) + " PS Time: " + (get-date -uformat %T))

    $GrpName = $Global:txtDeleName.Text
    $addMember = $Global:txtDeleMgrs.Text.Split(",")
    Add-Members
    $DeleNotes = "Room Delegates: " + ((Get-DistributionGroupMember $Global:txtDeleName.Text) -join ", ") + " - Per: " + $Global:txtInpTaskNo.Text
    Set-group $Global:txtDeleName.Text -Notes $DeleNotes
}
else
{
    $LineToWrite = "EXISTS" + "`t" + "     Room Delegates Group Already Exists: " + $Global:txtDeleName.Text
    WriteReportEvent    
}

#  Check to see if there is a specific group of room users and if so does the group exist
If ($Global:chkResUsr.Checked -eq "Checked")
{
    $Exists = [bool](Get-DistributionGroup $Global:txtResUsr.Text -ErrorAction SilentlyContinue)
    If ($Exists -ne $True)
    {
   	# Create the Room User Group
	    New-DistributionGroup -Name $Global:txtResUsr.Text `
	        -PrimarySmtpAddress (($Global:txtResUsr.Text) + "@ul.com") `
	        -Alias $Global:txtResUsr.Text `
	        -ManagedBy "MBX.RRS.Owner" `
	        -Type Security `
	        | Out-Null

        $DeleNotes = "Room Delegates: " + ((Get-DistributionGroupMember $Global:txtDeleName.Text) -join ", ") + " - Per: " + $Global:txtInpTaskNo.Text
        Set-group $Global:txtResUsr.Text -Notes $DeleNotes
        $LineToWrite = "NEW" + "`t" + "     Create Restricted User Group: " + $Global:txtResUsr.Text
        WriteReportEvent

	    Set-DistributionGroup $Global:txtResUsr.Text -HiddenFromAddressListsEnabled $true `
	        -RequireSenderAuthenticationEnabled $false `
	        -BypassSecurityGroupManagerCheck `
	        -CustomAttribute15 ($FileName + " PS Date: " + (get-date -uformat %D) + " PS Time: " + (get-date -uformat %T))
        $GrpName = $Global:txtResUsr.Text
        $addMember = $Global:txtResUsrInfo.Text.Split(",")
        Add-Members
    }
}