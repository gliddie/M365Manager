
$Mbx = Get-Recipient -Resultsize Unlimited -SortBy Alias|where {$_.RecipientTypeDetails -eq "UserMailbox"}
$mbx.count

#Do
foreach ($mbx in $mbx)
{
#    write-host "Enter Employee Number: " -ForegroundColor cyan -NoNewline
#    $uid = read-host
#    $usr = Get-ADUser $UID -Properties ProxyAddresses,mail,TargetAddress
    $usr = Get-ADUser $mbx.alias -Properties ProxyAddresses,mail,TargetAddress

    if (($usr.TargetAddress -notlike ("*" + $usr.Mail + "*")) -and (($mbx.alias).length -eq 5))
#    if (($usr.TargetAddress -notlike ("*" + $usr.Mail + "*")) -and ($UID.length -eq 5))
    {
        write-host "Updating Target Address for: " $mbx.alias
        write-host "Current Target Address: " $usr.TargetAddress
        write-host "Current Mail Address: " $Usr.mail
        Set-ADUser $mbx.alias -Replace @{targetaddress="SMTP:"+$usr.Mail}
    }
    else
    {
        write-host "Target Address and Mail Address already match or this is no a user account for: "$mbx.alias
    }
    $usr = Get-ADUser $mbx.Alias -Properties ProxyAddresses,mail,TargetAddress
#    $usr = Get-ADUser $UID -Properties ProxyAddresses,mail,TargetAddress
#    write-host "Do you have another account to review (Y/N)? " -ForegroundColor Cyan -NoNewline
#    $Cont = Read-Host
}
#}while ($cont -ne "N")