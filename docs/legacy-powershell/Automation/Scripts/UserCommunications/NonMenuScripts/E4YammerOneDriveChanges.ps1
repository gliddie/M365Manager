$Chg = "N"
$InpFile = import-csv C:\temp\YammerInput.csv

#write-host "Enter Employee Number of Individual " -ForegroundColor Yellow -NoNewline
#$EmpNo = Read-Host
#$EmpNo = $EmpNo + "@global.ul.com"

foreach ($InpFile in $InpFile)
{
    $SSKID = "ul:ENTERPRISEWITHSCAL"
    $UserLicenseTest = Get-MsolUser -UserPrincipalName $InpFile.UID
    $EmpType = get-mailbox $InpFile.UID

    If ($EmpType.CustomAttribute1 -eq "Employee")
    {
        Write-Host "Current License Assignment for " $InpFile.UID " " $UserLicenseTest.DisplayName
        $UserLicenseTest.ServiceStatus

        $NoLicenses = $UserLicenseTest.Licenses.Count

        $rpt = $UserLicenseTest.Licenses
        $cnt = 0
        foreach ($rpt in $rpt)
        {

            If ($rpt.AccountSkuId -eq $SSKID)
            {
                $PlanStats = $UserLicenseTest.Licenses[$Cnt].ServiceStatus
                $PlanStats

                $DisPlan = "MCOVOICECONF,MCOSTANDARD"
                If ($PlanStats[3].ProvisioningStatus -eq "Disabled")
                {
                    $DisPlan = $DisPlan + ",RMS_S_ENTERPRISE"
                }
                If ($PlanStats[5].ProvisioningStatus -eq "Disabled")
                {
                    $DisPlan = $DisPlan + ",OFFICESUBSCRIPTION"
                }
			    If ($PlanStats[7].ProvisioningStatus -eq "Disabled")
                {
                    $DisPlan = $DisPlan + ",SHAREPOINTWAC"
                }
                If ($PlanStats[8].ProvisioningStatus -eq "Disabled")
                {
                    $DisPlan = $DisPlan + ",SHAREPOINTENTERPRISE"
                }
			    If ($PlanStats[9].ProvisioningStatus -eq "Disabed")
                {
                    $DisPlan = $DisPlan + ",EXCHANGE_S_ENTERPRISE"
                }
 
                Write-Host "Disabled Plans" $DisPlan
                write-host
                $DLO = ($DisPlan.Split(“,”))
                $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
                Set-MsolUserLicense -UserPrincipalName $InpFile.UID –LicenseOptions $MyO365Sku
                Write-Host ""
                Write-Host "License Options Changed for" $InpFile.UID -ForegroundColor Red
                $UserLicenseTest = Get-MsolUser -UserPrincipalName $InpFile.UID
                $UserLicenseTest.ServiceStatus
                Write-Host ""
                Write-Host "Current License Assignment for " $InpFile.UID " " $UserLicenseTest.DisplayName -ForegroundColor Red
                $UserLicenseTest = Get-MsolUser -UserPrincipalName $InpFile.UID
                $UserLicenseTest.ServiceStatus

                $PlanStats = $UserLicenseTest.Licenses[$Cnt].ServiceStatus
                $PlanStats
          }

            $cnt = $cnt + 1
        }
    }
}