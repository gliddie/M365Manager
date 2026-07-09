<#
#
#  Created from Room Report by S.Glazebrook 01/22/2018
#
#>

Write-Host "Starting All Room/Resource Delegate Report......."
Write-Host ""
$Cnt = 1

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
        }
        else
        {
            $RoomUsers = "Authorized Users are the same as the Room Delegates"
        }

        $LineToWrite = $Resources.DisplayName + "`t`t`t" + $DelegateGrpMems + "`t`t" + $RoomUsers
        $LineToWrite
    }

write-host "Done"