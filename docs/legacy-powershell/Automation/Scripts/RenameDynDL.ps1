$name = Read-Host "Enter OLD Name"
Do {
    $newname = $NAME -replace ("UL ","")
    $newname = $NEWNAME -replace ("DST.All","DST.All UL COM")

    If ($newname.length -gt 64)
    {
        $newname = $newname.Substring(0,64)
    }

    $grp= get-dynamicdistributiongroup $name
    $newalias = ($grp.alias) -replace("DST.All","DST.ALLCOM")
    $newpaddr = $newalias + "@ul.com"
    $notes = $grp.notes -replace ("contains staff","contains UL.COM staff")
#    $mailtip = $grp.MailTip -replace ("Sends to individuals","Sends to UL.COM individuals")
    
    write-host "Updating group " $name " to " $newname -ForegroundColor Green
    Set-DynamicDistributionGroup $name -Name $newname -DisplayName $newname -Alias $newalias -ConditionalCompany "UL.COM" -PrimarySmtpAddress $newpaddr -notes $notes
#    Set-DynamicDistributionGroup $name -Name $newname -DisplayName $newname -Alias $newalias -ConditionalCompany "UL.COM" -PrimarySmtpAddress $newpaddr -notes $notes -mailtip $mailtip
    $name = Read-Host "Enter OLD Name"
}while ($name -ne "Exit")