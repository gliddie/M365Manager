$colUsers = Get-msolUser -MaxResults -1 | where {$_.IsLicensed -eq "True"} | sort-object "UserPrincipalName"
write-host "Numnber of accounts to review: " $colusers.count

foreach ($colusers in $colusers)
{
    $UserLicense = Get-MsolUser -UserPrincipalName $ColUsers.UserPrincipalName
    $HasE4 = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:ENTERPRISEWITHSCAL"})
    $HasE3 = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:EXCHANGEENTERPRISE"})
    $HasATP = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:ATP_ENTERPRISE"})
    $HasEMS = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:EMS"})
    $HasPBIF = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:POWER_BI_STANDARD"})
    $HasPBIP = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:POWER_BI_PRO"})

    if ($HasE4 -eq $True)
    {
#        write-host "Checking Standard License Assignment for: " $ColUsers.UserPrincipalName

        $mbx = get-mailbox $colUsers.UserPrincipalName
        If ($mbx.AccountDisabled -eq $True)
        {
            write-host "E4 License Assigned to Disabled Mailbox: " $colUsers.UserPrincipalName
            If ($HasATP -eq $True)
            {
                write-host "     ATP License Assigned to: " $Colusers.UserPrincipalName
            }
            
            If ($HasEMS -eq $True)
            {
                write-host "     EMS License Assigned to: " $Colusers.UserPrincipalName
            }

            If (($HasPBIP -eq $True) -or ($HasPBIF -eq $True))
            {
                write-host "     PowerBI Free or PowerBI Pro License Assigned to: " $Colusers.UserPrincipalName
            }
        }

        If ($HasATP -eq $False)
        {
            Write-host "     ATP License Assigned to: " $Colusers.UserPrincipalName
            Set-MsolUserLicense -UserPrincipalName $Colusers.UserPrincipalName -AddLicenses "ul:ATP_ENTERPRISE"
        }

        If ($HasEMS -eq $False)
        {
            Write-host "     EMS License Assigned to: " $Colusers.UserPrincipalName
            Set-MsolUserLicense -UserPrincipalName $Colusers.UserPrincipalName -AddLicenses "ul:EMS"
        }

        If (($HasPBIP -eq $False) -and ($HasPBIF -eq $False))
        {
            Write-host "     PowerBI Free License Assigned to: " $Colusers.UserPrincipalName
            Set-MsolUserLicense -UserPrincipalName $Colusers.UserPrincipalName -AddLicenses "ul:POWER_BI_STANDARD"
        }
    }
    else
    {
        If ($mbx.CustomAttribute2 -ne "T")
        {
            If ($HasE3 -eq $True)
            {
                write-host "E3 License Assigned to: " $Colusers.UserPrincipalName
            }
            else
            {
                Write-host "No E4 or E3 License Assigned to: " $Colusers.UserPrincipalName
            }
        }
    }
}