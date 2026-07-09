$name = Read-Host "Enter OLD Name"

Do {
#    $newname = $NAME -replace ("UL ","")
#    $newname = $NEWNAME -replace ("DST.All","DST.All UL COM")

    $newname = Read-Host "Enter NEW Name"
    $newnametrunc = $newname
    If ($newname.length -gt 64)
    {
        $newnametrunc = $newname.Substring(0,64)
        $newnametrunc = $newnametrunc.TrimEnd()
    }

    $grp= get-dynamicdistributiongroup $name
    $newalias = $grp.alias
    write-host "Group Alias: " $grp.Alias
    $chgalias = read-host "Change Alias (Y/N)? "
    If ($chgalias -eq "Y")
    {
        $newalias = read-host "Enter New Alias: "
    }
    
    If ($newname -like "*UL COM*")
    {
        $newpaddr = $newalias + "@ul.com"
        $notes = $grp.notes -replace ("contains staff","contains UL.COM staff")
    }
    else
    {
        $newpaddr = $newalias + "@ul.org"
        $notes = $grp.notes -replace ("contains staff","contains UL.ORG staff")
    }
#    $mailtip = $grp.MailTip -replace ("Sends to individuals","Sends to UL.COM individuals")
    
    write-host "Updating group " $name " to " $newname -ForegroundColor Green
    If ( $newname -like "*UL COM*")
    {
        Set-DynamicDistributionGroup $name -Name $newnametrunc -DisplayName $newname -Alias $newalias -ConditionalCompany "UL.COM" -PrimarySmtpAddress $newpaddr -notes $notes
    }
    else
    {
        Set-DynamicDistributionGroup $name -Name $newnametrunc -DisplayName $newname -Alias $newalias -ConditionalCompany "UL.ORG" -PrimarySmtpAddress $newpaddr -notes $notes
    }
#    Set-DynamicDistributionGroup $name -Name $newname -DisplayName $newname -Alias $newalias -ConditionalCompany "UL.COM" -PrimarySmtpAddress $newpaddr -notes $notes -mailtip $mailtip
    $name = Read-Host "Enter OLD Name (Exit to quit)"
}while ($name -ne "Exit")