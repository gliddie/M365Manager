<#
#
#  Created new by S.Glazebrook 01/22/2018
#
#>

Write-Host "Starting Report of Terminated Users on Legal Hold where mailbox access has been delegated......."

$LegalHoldTermed = Get-Mailbox -ResultSize Unlimited |Where-Object {($_.LitigationHoldEnabled -eq $true) -and ($_.CustomAttribute2 -eq "T")}

foreach ($LegalHoldTermed in $LegalHoldTermed)
{
    $MBXPerm = [bool](Get-MailboxPermission $LegalHoldTermed.Alias |Where-Object {($_.IsInherited -ne "True") -and ($_.User -ne "NT AUTHORITY\SELF")})

    If ($MBXPerm -eq "True")
    {
        $MBXPerm = Get-MailboxPermission $LegalHoldTermed.Alias |Where-Object {($_.IsInherited -ne "True") -and ($_.User -ne "NT AUTHORITY\SELF")}
    
#        foreach ($MBXPerm in $MBXPerm)
#        {
            write-host "Mail Access Delegated for Terminated Account" $LegalHoldTermed.Alias $LegalHoldTermed.Displayname "to"
            $MBXPerm
        }
#    }
    else
    {
        write-host "No Accounts have been delegated access to the Terminated Account" $LegalHoldTermed.Alias $LegalHoldTermed.Displayname
    }       
}