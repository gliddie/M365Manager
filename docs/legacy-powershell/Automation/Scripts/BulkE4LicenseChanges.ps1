$Usr = Import-csv "e:\SAG\MyScripts\PwrAppsFlow.csv"
$SSKID = "ul:ENTERPRISEWITHSCAL"
Invoke-Expression -command E:\O365AdminShared\Scripts\O365DisabledLicenseFeatures.ps1

foreach ($Usr in $Usr)
{
	$UsrLic = Get-MsolUser -UserPrincipalName $Usr.UPN
	$HasE4 = [bool] ($UsrLic.Licenses | where {$_.AccountSkuId -eq "ul:ENTERPRISEWITHSCAL"})
	
	if ($HasE4 -eq $True)
	{
		$Mbx = get-mailbox $Usr.UPN

        If ($Mbx.CustomAttribute1 -like "*Employee*")
        {
            $DisPlan = $Global:DisPlanEmpE4
        }
        else
        {
            $DisPlan = $Global:DisPlanNonEmpE4
        }

        $rpt = $UsrLic.Licenses
		foreach ($rpt in $rpt)
		{
			If ($rpt.AccountSkuId -eq $SSKID)
			{
				$PlanStats = $UsrLic.Licenses.ServiceStatus
				if (($PlanStats.ProvisioningStatus[4] -eq "Disabled") -or ($PlanStats.ProvisioningStatus[5] -eq "Disabled"))
				{
					write-host "Resetting License Options for " $Usr.UPN
				
					$DLO = ($DisPlan.Split(“,”))
					$MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
					Set-MsolUserLicense -UserPrincipalName $Usr.UPN –LicenseOptions $MyO365Sku
				}
			}
		}
	}
    else
    {
        write-host "User does not have an E4 License Assigned " $Usr.UPN
    }
}