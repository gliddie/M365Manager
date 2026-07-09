#### Directly assigned E5 licenses

#$Usr = Get-MsolUser -all | where { $_.IsLicensed -eq $True } |Sort-Object UserPrincipalName
foreach ($u in $usr)
{
#    write-host "Processing " $u.UserPrincipalName    
    foreach ($license in $u.Licenses)
    {
        if ($license.GroupsAssigningLicense[0].length -gt 0)
        {
            if ($license.GroupsAssigningLicense[0].ToString() -eq $u.ObjectId)
            {
                If (($license.AccountSkuId -eq "ul:EMSPREMIUM") -or ($license.AccountSkuId -eq "ul:ENTERPRISEPREMIUM"))
                {
                    Write-Host $u.UserPrincipalName , $u.DisplayName , $license.AccountSkuId
                }
            }
        }
    }
}

foreach ($u in $usr)
{
    If ($u.UserPrincipalName.Length -eq 19)
    {
#        If (($u.licenses.AccountSkuId -eq "ul:ENTERPRISEPREMIUM") -and ($u.licenses.AccountSkuId -eq "ul:EMSPREMIUM"))
        If ($u.licenses.AccountSkuId -eq "ul:ENTERPRISEPREMIUM")
        {
            If ($u.licenses.AccountSkuID -eq "ul:EMSPREMIUM")
            {
            }
            else
            {
                write-host "No EMS license"
                write $u.UserPrincipalName
                write $u.licenses.AccountSkuID
            }
        }
    }
}

foreach ($u in $usr)
{
    If ($u.licenses.AccountSkuId -eq "ul:EMSPREMIUM")
    {
        write-host $u.UserPrincipalName " - EMS License" -ForegroundColor Yellow
    }
    If ($u.licenses.AccountSkuID -eq "ul:ENTERPRISEPREMIUM")
    {
        write-host $u.UserPrincipalName " - E5 License" -ForegroundColor Green
    }
}
