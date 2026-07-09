<#
#
#  Created from Room Report by S.Glazebrook 01/22/2018
#
#>
Start-Transcript

Write-Host "Starting All Room/Resource Delegate Report......."
Write-Host ""

# Get all the Room Lists
    $RoomList = Get-DistributionGroup -ResultSize Unlimited |Where-Object {$_.RecipientTypeDetails -eq "RoomList"} |Sort-Object DisplayName
    $Resources = Get-mailbox -ResultSize Unlimited |Where-Object {$_.RecipientTypeDetails -eq "EquipmentMailbox"} |Sort-Object DisplayName

    $LineToWrite = "Room/Resource Name" + "`t" + "Delegate Group Name" + "`t" + "Delegate Group Members" + "`t" + "User Group Name" + "`t" + "Authorized Room Users"
    $LineToWrite

	foreach ($RoomList in $RoomList)
	{
        write-host "`nWriting details for " $RoomList -ForegroundColor Cyan
        $RoomListMembers = Get-DistributionGroupMember $RoomList.DisplayName

        foreach ($RoomListMembers in $RoomListMembers)
        {
            $DelegateGrpName = (Get-CalendarProcessing -Identity $RoomListMembers.DisplayName).ResourceDelegates
            $DelegateGrpMems = (Get-Group $DelegateGrpName[0]).members -join ", "

            if ($RoomList -like "*Restricted*")
            {
                $UserGrpName = (Get-CalendarProcessing -Identity $RoomListMembers.DisplayName).BookInPolicy

                If ($UserGrpName.count -gt 0)
                {
                    $RoomUsers = (Get-Group $UserGrpName[0]).members -join ", "
                    $UserGrpName = Get-Group $UserGrpName[0] |select Name
                }
                else
                {
                    $RoomUsers = "Authorized Users are the same as the Room Delegates"
                    $UserGrpName = "N/A"
                }
            }
            else
            {
                $RoomUsers = "General Use Room"
                $UserGrpName = "N/A"
            }
            $LineToWrite = $RoomListMembers.DisplayName + "`t" + $DelegateGrpName[0] + "`t" + $DelegateGrpMems + "`t" + $UserGrpName + "`t" + $RoomUsers
            $LineToWrite
        }
     }

    $Resources = Get-mailbox -ResultSize Unlimited |Where-Object {$_.RecipientTypeDetails -eq "EquipmentMailbox"} |Sort-Object DisplayName

    write-host "`nWriting details for" $Resources.Count "Configured Resources" -ForegroundColor Cyan
    foreach ($Resources in $Resources)
	{
        $DelegateGrpName = (Get-CalendarProcessing -Identity $Resources.DisplayName).ResourceDelegates

        If ($DelegateGrpName -like "*MBX.*")
        {
            $DelegateGrpMems = (Get-Group $DelegateGrpName[0]).members -join ", "
        }
        else
        {
            $DelegateGrpMems = $DelegateGrpName -join ", "
        }

        $UserGrpName = (Get-CalendarProcessing -Identity $Resources.DisplayName).BookInPolicy

        If ($UserGrpName.count -gt 0)
        {
            $ResourceUsers = (Get-Group $UserGrpName[0]).members -join ", "
            $UserGrpName = Get-Group $UserGrpName[0] |select Name
        }
        else
        {
            $RoomUsers = "General Use Resource"
            $UserGrpName = "N/A"
        }

        $LineToWrite = $RoomListMembers.DisplayName + "`t" + $DelegateGrpName[0] + "`t" + $DelegateGrpMems + "`t" + $UserGrpName + "`t" + $RoomUsers
        $LineToWrite
    }

write-host "Done"

Stop-Transcript
