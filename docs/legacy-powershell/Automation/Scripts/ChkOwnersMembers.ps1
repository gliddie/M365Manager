write-host "Enter Name of Item: " -ForegroundColor Cyan -NoNewline
$Name = read-host

Do
{
    if ($Name -like "LST*")
    {
        Write-Host "     Managers: " (Get-DistributionGroup $Name).ManagedBy -join (", ")
        Write-Host "     Members: " (Get-DistributionGroupMember $Name)

        write-host "Send NoOwners Email (Y/N)? " -ForegroundColor Red -NoNewline
        $SendMsg = Read-Host

        If ($SendMsg -eq "Y")
        {
            Invoke-Expression -Command "C:\Users\96151\Documents\EmailTemplates\DistributionListNoOwners.oft"
        }
    }
    else
    {
        get-mailbox $Name |fl *forw*
        $MbxPerm = Get-MailboxPermission $Name | Where-Object {$_.User -like "MBX*"}

        If ($MbxPerm.Count -ne 0)
        {
            Get-MailboxPermission $Name | Where-Object {$_.User -like "MBX*"} | ft User,AccessRights
            foreach ($MbxPerm in $MbxPerm)
            {
                Write-host $MbxPerm.User "Details:"
                Write-Host "     Managers: " ((Get-DistributionGroup $MbxPerm.User).ManagedBy -join(", "))
                Write-Host "     Members: " (Get-DistributionGroupMember $MbxPerm.User)
            }
        }

        $FldrPerm = Get-MailboxFolderPermission $Name | Where-Object {$_.User -like "MBX*"}

        If ($FldrPerm.Count -ne 0)
        {
            Get-MailboxFolderPermission $Name | Where-Object {$_.User -like "MBX*"} | ft User,AccessRights
            foreach ($FldrPerm in $FldrPerm)
            {
                Write-host $FldrPerm.User "Details:"
                Write-Host "     Managers: " ((Get-DistributionGroup $FldrPerm.User.DisplayName).ManagedBy -join(","))
                Write-Host "     Members: " (Get-DistributionGroupMember $FldrPerm.User.DisplayName)
            }
        }

        write-host "Send NoOwners Email (Y/N)? " -ForegroundColor Red -NoNewline
        $SendMsg = Read-Host
        
        If ($SendMsg -eq "Y")
        {
            Invoke-Expression -Command "C:\Users\96151\Documents\EmailTemplates\SharedMailboxNoOwners.oft"
        }

    }

    write-host ""
    write-host "Enter Name of Item: " -ForegroundColor Cyan -NoNewline
    $Name = read-host
} while ($Name -ne "")