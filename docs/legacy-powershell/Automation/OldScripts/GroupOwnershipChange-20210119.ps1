<#
#  This script allows you to make changes to Distribution List or Security Group Ownership
#  Called by: DistributionSecurityGroupMenu.ps1
#             SharedMailboxAdminMenu.ps1
#
#  08/11/2019 - SAG - Modified to change the ownership on a Unified Group and/or Microsoft Team
#  10/01/2019 - SAG - Modified so that changes to the DL owner also modifies the MailTip on the distribution group.
#  10/02/2019 - SAG - Modified so to check if this change is for a shared mailbox and if so to also update the mail tip on the shared mailbox
#  11/03/2019 - SAG - Modified code that updates the mail tips on a shared mailbox
#  01/21/2020 - SAG - Modified Option 2 to update shared maiblox mailtip details
#  03/26/2020 - SAG - Modified to add a pause before performing the Set-Group process as updates were not always applying as quickly as the script executes.
#  04/10/2020 - SAG - Fixed the code in setting for shared mailbox ownership changes.
#>

Function ProcessGroup($GrpAct)
{
    $GrpOwner = Read-Host "Enter the Disply Name, Email Address or Employee Number of New Owner"
    if ($GrpOwner.length -eq 5)
    {
        $GrpOwner = $GrpOwner + "@global.ul.com"
    }

    $MbxExists = [bool](get-mailbox $GrpOwner -ErrorAction SilentlyContinue)

    If ($MbxExists -eq "True")
    {
        $mbx = get-mailbox $GrpOwner
        Switch ($GrpOpt)
        {
            1
                {
                    If ($GrpName -notlike "GRP.*")
                    {
                        Set-DistributionGroup $GrpName -ManagedBy $GrpOwner -BypassSecurityGroupManagerCheck -MailTip ("Owner: " + $mbx.Name)
                        Set-Group $GrpName -Notes ("Owner:  " + $mbx.Name + " - Per " + $SDTicketNo)
                        If (($GrpName -like "*MBX*") -and ((((get-mailbox $Global:ShrMbx).MailTip) -like "*Owner*") -or (((get-mailbox $Global:ShrMbx).MailTip).Length -eq 0)))
                        {
                            set-mailbox $ShrMbx -MailTip ("Owner: " + $mbx.Name)
                        }
                    }
                    else
                    {                        
                        $NewDesc = $TeamDesc.Description.Remove($TeamDesc.Description.IndexOf(":")) + (":  " + $mbx.Name + "`nPer: " + $SDTicketNo)
                        AddUFGOwners
                    }
                }
            2
                {
                    If ($GrpName -notlike "GRP.*")
                    {
                        Set-DistributionGroup $GrpName -ManagedBy @{add="$GrpOwner"} -BypassSecurityGroupManagerCheck
                        Start-Sleep -Seconds 10
                        $GrpInf.Notes = ("Owner: " + ((get-DistributionGroup $GrpName).ManagedBy -join ", ") + " - Per " + $SDTicketNo)
                        $MailTip = ("Owner: " + ((get-DistributionGroup $GrpName).ManagedBy -join ", "))
                        Set-DistributionGroup $GrpName -MailTip $MailTip
                        
                        If (($GrpName -like "*MBX*") -and ((((get-mailbox $Global:ShrMbx).MailTip) -like "*Owner*") -or (((get-mailbox $Global:ShrMbx).MailTip).Length -eq 0)))
                        {
                            set-mailbox $ShrMbx -MailTip $MailTip
#                            set-mailbox $ShrMbx -MailTip ("Owner: " + ((get-DistributionGroup $GrpName).ManagedBy -join ", "))
                        }
                        Set-Group $GrpName -Notes $GrpInf.Notes
                    }
                    else
                    {
                        $NewDesc = $GrpInf.Notes.Remove($GrpInf.Notes.IndexOf("Per:")) + ", " + $mbx.Name + "`nPer:  " + $SDTicketNo
                        AddUFGOwners
                    }
                }
        }
    }
    else
    {
        write-host ""
        Write-Host "User specified does not exist" -ForegroundColor Red
    }
}

Function AddUFGOwners($GrpAct)
{
    If ($TeamExists -eq "True")
    {
        $TeamGrpID = ((get-team |where-object {$_.DisplayName -eq $GrpName}).GroupID)
        Add-TeamUser -GroupId $TeamGrpID -User $DGOwnerID -Role Member
        Add-TeamUser -GroupId $TeamGrpID -User $DGOwnerID -Role Owner
    }
    else
    {
        Add-UnifiedGroupLinks $GrpName -LinkType Member $GroupOwner
        Add-UnifiedGroupLinks $GrpName -LinkType Owner $GroupOwner
    }
                       
    $NewDesc = $TeamDesc.Description.Remove($TeamDesc.Description.IndexOf(":")) + (":  " + $mbx.Name + "`nPer: " + $SDTicketNo)
    Set-Group $GrpName -Notes $NewDesc
}

write-host "Enter Name of Distribution List, Security Group or Unified Group " -ForegroundColor Yellow -NoNewline
$GrpName = Read-Host
                    
$GrpExists = [bool](Get-DistributionGroup $GrpName -ErrorAction SilentlyContinue)
If ($GrpExists -eq $False)
{
#  This will find groups that are Unified Groups/Microsoft Teams    
    write-host "Group is not a distribution list checking to see if this is a Unified group or a Microsoft Team"
    write-host "Checking if this is a unified group"
    $GrpExists = [bool](Get-UnifiedGroup $GrpName -ErrorAction SilentlyContinue)
#    If ($GrpExists -eq $False)
#    {
#        write-host "Checking if this is a Microsoft Team"
#        $TeamExists = [bool]((get-team |where-object {$_.DisplayName -eq $GrpName}).GroupID)
#    }
}

If ($GrpExists -eq "True")
{
    if ($GrpName -notlike "GRP*")
    {
        $Grp = get-distributiongroup $GrpName
        $GrpInf = get-group $GrpName
    }
    else
    {
        $Grp = get-unifiedgroup $GrpName
        $GrpInf = get-UnifiedGroup $GrpName
    }
                    
    write-host
    write-host "Group Name    " $Grp.Name
    write-host "Managed by    " $Grp.ManagedBy
    Write-Host "Group Notes   " $GrpInf.Notes
    write-host "Group MailTip " $Grp.MailTip
    write-host 

    Write-Host "Is the requestor an owner of the distribution list, security group or unified group have you obtained approval" -ForegroundColor Yellow
    write-host "from an owner to make this change (Y/N)? " -ForegroundColor Yellow -NoNewline
    $Approval = Read-Host

    If ($Approval -ne "Y")
    {
        write-host "No changes made please obtain approval before proceeding" -ForegroundColor Red
        pause
    }
    else
    {
        $MoreChgs = "Y"
        write-host "Enter Task Number for this Request: " -ForegroundColor Yellow -NoNewline
        $SDTicketNo = Read-Host
        DO
        {
            write-host "Enter (1) to Replace the Owner"
            write-host "      (2) to Add an Owner"
            write-host "      (3) to Remove an Owner"
            write-host "      (4) No changes"
            write-host "Enter Option? " -ForegroundColor Red -NoNewline
            $GrpOpt = Read-Host
            switch ($GrpOpt)
                                                                                                                                                                                        {
            1
                {
                    #Replace Group Owner
                    $GrpAct = "Replacing Distribution List, Security Group or Unified Group Owner"
                    ProcessGroup ($GrpAct)
                }
            2
                {
                    #Add Group Owner
                    $GrpAct = "Adding Distribution List, Security Group or Unified Group Owner"
                    ProcessGroup ($GrpAct)
                }
            3
                {
                    #Remove Group Owner
                    Write-Host
                    Write-Host "Enter (1) to Remove from ManagedBy, Notes and MailTip"
                    Write-Host "      (2) to Remove from Notes only"
                    Write-Host "      (3) to Remove from ManagedBy Only"
                    Write-Host "Enter Option? " -ForegroundColor Red -NoNewline
                    $RemAns = Read-Host
                                
                    if (($RemAns -eq 1) -or ($RemAns -eq 3))
                    {
                        $GrpOwner = Read-Host "Enter Email Address or Employee Number of Owner to Remove"
                        if ($GrpOwner.length -eq 5)
                        {
                            $GrpOwner = $GrpOwner + "@global.ul.com"
                        }
                    }
                                           
                    if ($RemAns -ne 3)
                    {
                        Write-Host "Setting Owners List to match who is in the ManagedBy field" -ForegroundColor Yellow
                        Start-Sleep -Seconds 10
                        $MailTip = Get-DistributionGroup $GrpName
                        $GrpInf.Notes = "Owner: " + ($MailTip.ManagedBy -join (",")) + " - Per: " + $SDTicketNo
                        If (($GrpName -like "*MBX*") -and ((((get-mailbox $Global:ShrMbx).MailTip) -like "*Owner*") -or (((get-mailbox $Global:ShrMbx).MailTip).Length -eq 0)))
                        {
                            If ($ShrMbx -eq "")
                            {
                                $ShrMbx = read-host "Enter Name of Shared Mailbox: "
                            }
                            set-mailbox $ShrMbx -MailTip ("Owner: " + ($MailTip.ManagedBy -join (", "))) 
                        }
                        set-Group $GrpName -Notes ("Owner: " + ($MailTip.ManagedBy -join (", "))) 
                    
                    }

                    if (($RemAns -eq 1) -or ($RemAns -eq 3))
                    {
                        Set-DistributionGroup $GrpName -ManagedBy @{remove="$GrpOwner"} -BypassSecurityGroupManagerCheck
                    }
                }
        }
            If ($GrpOpt -lt 4)
            {
                write-host "Updated group details"

                if ($GrpName -notlike "GRP*")
                {
                    $Grp = get-distributiongroup $GrpName
                    $GrpInf = get-group $GrpName
                }
                else
                {
                    $Grp = get-unifiedgroup $GrpName
                    $GrpInf = get-UnifiedGroup $GrpName
                }

                write-host
                write-host "Group Name "  $Grp.Name
                write-host "Managed by "  $Grp.ManagedBy
                Write-Host "Group Notes " $GrpInf.Notes
                write-host
            }
            Write-Host ""
            Write-Host "Do you have Additional Ownership Changes for this list (Y/N)? " -ForegroundColor Yellow -NoNewline
            $MoreChgs = Read-Host 
        } while ($MoreChgs -ne "N")
    }
}
else
{
    Write-Host ""
    Write-Host "Distribution Group or Security Group Does Not Exist" -ForegroundColor Red
}