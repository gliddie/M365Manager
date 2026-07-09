$SSKID = "ul:ENTERPRISEWITHSCAL"

$Date = get-date -Format "yyyy-MMdd"
$ReportFile = "C:\temp\E4LicenseReport-" + $Date + ".csv"

# Start-Transcript

$strPath = "C:\temp\E4LicenseReport-" + $Date + ".csv"

$colUsers = Get-msolUser -MaxResults -1 | where{$_.Licenses.AccountSkuId -eq "ul:ENTERPRISEWITHSCAL"}

$colUsers.count
$colusers

$text = "UPN,License SKU,Sway,MobileDevMgmt,Yammer Enterprise,AzureRightsMgmt,SkypeforBusPlan3,OfficeProPlus,SkypeforBusPlan2,Office Online,SharePointPlan2,Exchange_S_Enterprise"
Out-File -FilePath $strPath -InputObject $text

Foreach ($objUser in $colUsers)
{
    $UsrLicense = Get-MsolUser -UserPrincipalName $objUser.UserPrincipalName
	$UsrLicense.ServiceStatus
	$NoLicenses = $UsrLicense.Licenses.Count
	
	$rpt = $UsrLicense.Licenses
	$cnt = 0

    foreach ($rpt in $rpt)
	{

		If ($rpt.AccountSkuId -eq $SSKID)
		{
			$PlanStats = $UsrLicense.Licenses[$Cnt].ServiceStatus
			
			$OfficePlanner = $PlanStats[0].ProvisioningStatus
			$SwayStat = $PlanStats[1].ProvisioningStatus
			$MobileDevMgmt = $PlanStats[2].ProvisioningStatus
			$YammerStat = $PlanStats[3].ProvisioningStatus
			$AzureRightsMgmtStat = $PlanStats[4].ProvisioningStatus
			$SFBPlan3Stat = $PlanStats[5].ProvisioningStatus
			$officeProPlusStat = $PlanStats[6].ProvisioningStatus
			$SFBPlan2Stat = $PlanStats[7].ProvisioningStatus
			$OfficeOnlineStat = $PlanStats[8].ProvisioningStatus
			$SPEnterpStat = $PlanStats[9].ProvisioningStatus
			$ExchgStat= $PlanStats[10].ProvisioningStatus
				
			$text = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},[12]" -f $objUser.UserPrincipalName, $rpt.AccountSkuId, $PlanStats[0].ProvisioningStatus, $PlanStats[1].ProvisioningStatus, $PlanStats[2].ProvisioningStatus, $PlanStats[3].ProvisioningStatus, $PlanStats[4].ProvisioningStatus, $PlanStats[5].ProvisioningStatus, $PlanStats[6].ProvisioningStatus, $PlanStats[7].ProvisioningStatus, $PlanStats[8].ProvisioningStatus, $PlanStats[9].ProvisioningStatus, PlanStats[10].ProvisioningStatus
				
			Out-File -FilePath $strPath -InputObject $text -Append
				
			$text = ""
            $PlanStats = ""
		}
		
		$cnt = $cnt + 1
	}
}