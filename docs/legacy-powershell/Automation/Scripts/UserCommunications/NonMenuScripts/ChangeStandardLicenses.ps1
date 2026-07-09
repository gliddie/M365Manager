$Global:DisPlanEmpE4    = ""
$Global:DisPlanNonEmpE4 = ""
$Global:DisPlanEmpE3 	 = ""
$Global:DisPlanNonEmpE3 = ""

Invoke-Expression -Command e:\O365AdminShared\Scripts\O365DisabledLicenseFeatures.ps1

$InpData = import-csv e:\Automation\LicenseChanges\LicensedUsers.csv

foreach ($InpData in $InpData)
{

    $UsrLic = Get-MsolUser -UserPrincipalName $InpData.UID
    $UsrLic.ServiceStatus

    $HasE3 = [bool]($UsrLic.Licenses | where {$_.AccountSkuId -eq "ul:ENTERPRISEPACK"})
    $HasE4 = [bool]($UsrLic.Licenses | where {$_.AccountSkuId -eq "ul:ENTERPRISEWITHSCAL"})

    If ($HasE3 -eq "True")
    {
        $SSKID = "ul:ENTERPRISEPACK"
    }
    else
    {
        $SSKID = "ul:ENTERPRISEWITHSCAL"
    }

    $CurLic = $UsrLic.Licenses |Where-Object {$_.AccountSkuId -eq $SSKID}
    $MbxDet = get-mailbox $InpData.UID
    $DisPlan = ""
    Write-Host "Employee Number " $InpData.UID " - " $MbxDet.CustomAttribute1 " - Enabled Licenses"
    $CurLic.ServiceStatus | ft
    
    if ($MbxDet.CustomAttribute1 -eq "Employee")
    {
    #    Assign Employee Features
        if ($HasE3 -eq "True")
        {
            $DisPlan = $Global:DisPlanEmpE3
        }
        if ($HasE4 -eq "True")
        {
            $DisPlan = $Global:DisPlanEmpE4
        }
    }
    else
    {
    #    Assign Non-Employee Features
        if ($HasE3 -eq "True")
        {
            $DisPlan = $Global:DisPlanNonEmpE4
        }
        if ($HasE4 -eq "True")
        {
            $DisPlan = $Global:DisPlanNonEmpE4
        }
    }

    Write-Host "Disabled Plans" $DisPlan
    Write-host
    $DLO = ($DisPlan.Split(“,”))
    $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
    Set-MsolUserLicense -UserPrincipalName $InpData.UID –LicenseOptions $MyO365Sku
    Write-Host ""
    Write-Host "New License Assignment for " $InpData.UID " " $UserLicenseTest.DisplayName -ForegroundColor Red
    
    $UsrLicNew = Get-MsolUser -UserPrincipalName $InpData.UID
    $UsrLicNew.ServiceStatus
    $CurLicNew = $UsrLicNew.Licenses |Where-Object {$_.AccountSkuId -eq $SSKID}
    $CurLicNew.ServiceStatus |ft

}