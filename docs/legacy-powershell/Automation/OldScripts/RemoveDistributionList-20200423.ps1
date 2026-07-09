<#
#
#  Called by:  DistributionSecurityGroupMenu.ps1
#              RoomResourceAdminMenu.ps1
#
#  06/14/2018 - Added to to delete Unified Groups
#  08/04/2018 - Added code if there are no members in a DST group to remove without confirming
#  01/28/2019 - Added -ResultSize Unlimited to the get-DistributionListMember statement
#  04/08/2019 - Added -ResultSize Unlimited to the line that initially reports how many individuals in the group
#  12/10/2019 - Added code so if the name does not follow the naming standard the requestor identifies the group type
#>

$InputDL = Read-Host "Name of Distribution List"

If (($InputDL -notlike "DST*") -and ($InputDL -notlike "LST*") -and ($InputDL -notlike "GRP*"))
{
    write-host "`n     This group does not follow our naming standard. What type of group is this?"
    write-host "         (1) Distribution/Security Group"
    write-host "         (2) Dynamic Distribution Group"
    write-host "         (3) O365 Group"
    write-host "`n     Enter type " -NoNewline
    $GrpType = read-host    
}

If (($InputDL -like "DST*") -or ($GrpType -eq 2))
{
    $DstExists = [bool](Get-DynamicDistributionGroup $InputDL -ErrorAction SilentlyContinue)
    If ($DstExists -eq "True")
    {
        $RemDL = Get-DynamicDistributionGroup $InputDL
    }
}
elseif (($InputDL -like "GRP*") -or ($GrpType -eq 3))
{
    $ModInputDL = $InputDL.Replace(" ","")
    $ModInputDL = $ModInputDL.Replace(".","")
    $DstExists = [bool](Get-UnifiedGroup $ModInputDL -ErrorAction SilentlyContinue)
    If ($DstExists -eq "True")
    {
        $RemDL = Get-UnifiedGroup $ModInputDL
    }
}
else
{
    $DstExists = [bool](Get-DistributionGroup $InputDL -ErrorAction SilentlyContinue)
    If ($DstExists -eq "True")
    {
        $RemDL = Get-DistributionGroup $InputDL
    }
}

write-host "List Owner: " $RemDL.ManagedBy
write-host "Is the Requestor an Owner of the Group or Approval has been obtained (Y/N) " -ForegroundColor Yellow -NoNewline
$Cont = Read-Host

If ($Cont -eq "Y")
{
    if ($DstExists -eq "True")
    {
        Write-host "Reason for Removal: " -ForegroundColor Cyan -NoNewline
        $Reason = Read-Host
        $Reason = "Reason for Removal: " + $Reason

        If (($InputDL -like "DST*") -or ($GrpType -eq 2))
        {
            $RemDL = get-DynamicDistributionGroup $InputDL
            write-host ""
	        $g = Get-DynamicDistributionGroup $InputDL
            $colu = Get-Recipient -RecipientPreviewFilter $g.LdapRecipientFilter -ResultSize Unlimited
            $Cont = "Y"
        }
        elseif (($InputDL -like "GRP*") -or ($GrpType -eq 3))
        {
            $RemDL = Get-UnifiedGroupLinks $ModInputDL -LinkType Owners
            write-host ""
	        Write-Host "Name of Group  : " $InputDL
            Write-Host "Owner of Group : " $RemDL.Name
        }
        else
        {
            $RemDL = get-DistributionGroup $InputDL
            write-host ""
            Write-Host "Name of Group  : " $InputDL
            Write-Host "Owner of Group : " $RemDL.ManagedBy
            write-host "Number of Group Members : " (get-distributiongroupmember $InputDL -Resultsize Unlimited).count
        }

        $OutFileName = "c:\temp\" + ($InputDL -replace "/ ","") + ".txt"

        start-transcript
        Write-Output "Removing Distribution List" > $OutFileName
        Write-Output $Reason >> $OutFileName

        if (($InputDL -like "DST*") -or ($GrpType -eq 2))
        {
            Write-Host
            Write-Host "     Number of members in this group " -ForegroundColor Yellow -NoNewline
            Write-Host $colu.count -ForegroundColor Yellow
            write-host ""
            Write-Output "Dynamic Distribution Group Information" >> $OutFileName
            Get-DynamicDistributionGroup $InputDL >> $OutFileName
            Write-Output "Dynamic Distribution Group Detailed Information" >> $OutFileName
            Get-DynamicDistributionGroup $InputDL |fl >> $OutFileName
            Write-Output "Dynamic Distribtuion Group Membership Rule" >> $OutFileName
            Get-DynamicDistributionGroup $InputDL |fl RecipientFilter,IncludedRecipients >>$OutFileName

            If ($colu.count -eq 0)
            {
                Remove-DynamicDistributionGroup $InputDL -confirm:$False
            }
            else
            {
                Remove-DynamicDistributionGroup $InputDL
            }
        }
        elseif (($InputDL -like "GRP*") -or ($GrpType -eq 3))
        {
            Write-Output "Unified Group Information" >> $OutFileName
            Get-UnifiedGroup $ModInputDL >> $OutFileName
            Write-Output "Unified Group Detailed Information" >> $OutFileName
            Get-UnifiedGroup $ModInputDL |fl >> $OutFileName
            Write-Output "Unfied Group Owners and Members " >> $OutFileName
            Get-UnifiedGroupLinks -LinkType Owners $ModInputDL |fl Name >>$OutFileName
            Get-UnifiedGroupLinks -LinkType Members $ModInputDL |fl Name >>$OutFileName
            Remove-UnifiedGroup $ModInputDL -Confirm:$False
        }
        else
        {
            Write-Output "Distribution Group Information" >> $OutFileName
            get-DistributionGroup $InputDL >> $OutFileName
            Write-Output "Distribution Group Detailed Information" >> $OutFileName
            get-DistributionGroup $InputDL |fl >> $OutFileName
            Write-Output "Distribution Group Manager and Notes" >> $OutFileName
            get-Group $InputDL |fl ManagedBy,Notes >> $OutFileName
            Write-Output "Distribution Group Membership" >> $OutFileName
            get-DistributionGroupMember $InputDL -ResultSize Unlimited|ft Alias,Name,RecipientType >> $OutFileName
            remove-DistributionGroup $InputDL -BypassSecurityGroupManagerCheck
        }
    
        write-host ""
        Write-Host "Copy file from " $OutFileName " to the Team Sharepoint Site" -ForegroundColor Yellow

        Stop-transcript
    }
    else
    {
        Write-Host "Distribution List Does Not Exist - No Changes Made" -ForegroundColor Red
    }
}
else
{
    Write-Host "Group not removed obtain appropriate approval " -ForegroundColor Red
    pause
}
