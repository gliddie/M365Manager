$Chg = "N"
$InpFile = import-csv C:\temp\UserLicense.csv

#write-host "Enter Employee Number of Individual " -ForegroundColor Yellow -NoNewline
#$EmpNo = Read-Host
#$EmpNo = $EmpNo + "@global.ul.com"

foreach ($InpFile in $InpFile)
{
#    $SSKID = "ul:ENTERPRISEWITHSCAL"
    $SSKID = "ul:ENTERPRISEPACK"
#    $UserLicenseTest = Get-MsolUser -UserPrincipalName $EmpNo
    $UserLicenseTest = Get-MsolUser -UserPrincipalName $InpFile.UID
#	 $EmpType = get-mailbox $EmpNo
    $EmpType = get-mailbox $InpFile.UID
	$PlanStats = $UserLicenseTest.Licenses | where {$_.AccountSkuId -eq $SSKID}

    If ($PlanStats -ne $null)
    {
	    Write-Host "Current License Assignment for " $InpFile.UID " " $UserLicenseTest.DisplayName -ForegroundColor Red
	    $PlanStats.ServiceStatus
	
        If ($EmpType.CustomAttribute1 -eq "Employee")
        {
#            $DisPlan = "Deskless,FLOW_O365_P2,POWERAPPS_O365_P2,TEAMS1,MCOVOICECONF,MCOSTANDARD"
            $DisPlan = "Deskless,FLOW_O365_P2,POWERAPPS_O365_P2,TEAMS1,MCOSTANDARD"
        }
	    else
	    {
#	    $DisPlan = "Deskless,FLOW_O365_P2,POWERAPPS_O365_P2,TEAMS1,YAMMER_ENTERPRISE,MCOVOICECONF,MCOSTANDARD,SHAREPOINTWAC,SHAREPOINTENTERPRISE"
        $DisPlan = "Deskless,FLOW_O365_P2,POWERAPPS_O365_P2,TEAMS1,YAMMER_ENTERPRISE,MCOSTANDARD,SHAREPOINTWAC,SHAREPOINTENTERPRISE"
	    }
	
        $DLO = ($DisPlan.Split(“,”))
        $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $InpFile.UID –LicenseOptions $MyO365Sku

        Write-Host ""
	    Write-Host "Updated License Assignment for " $InpFile.UID " " $UserLicenseTest.DisplayName -ForegroundColor Red
	    $UserLicenseTest = Get-MsolUser -UserPrincipalName $InpFile.UID
        $PlanStats = $UserLicenseTest.Licenses | where {$_.AccountSkuId -eq $SSKID}
        $PlanStats.ServiceStatus
    }
}