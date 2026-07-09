$Date = get-date -Format "yyyy-MMdd"
$ReportFile = "C:\scripts\WeeklyLicenseReports\ReportHistory\O365License\O365_WeeklyLicenseReport-" + $Date + ".csv"

Start-Transcript -Path .\O365_WeeklyLicenseReport.log

$strPath = "C:\scripts\WeeklyLicenseReports\O365_WeeklyLicenseReport.csv"

Import-PSSession $Session -AllowClobber
Import-Module ActiveDirectory
Import-Module MSOnline
Connect-MsolService –Credential $LiveCred

$colUsers = Get-msolUser -MaxResults -1

$text = "UPN,SWAY,INTUNE_0365,YAMMER_ENTERPRISE,RMS_S_ENTERPRISE,MCOVOICECONF,OFFICESUBSCRIPTION,MCOSTANDARD,SHAREPOINTWAC,SHAREPOINTENTERPRISE,EXCHANGE_S_ENTERPRISE"
Out-File -FilePath $strPath -InputObject $text

$SSKID = "ul:ENTERPRISEWITHSCAL"

Foreach ($objUser in $colUsers)
{
		Foreach ($objLic in $objUser.Licenses)
	{
		$UsrLicense = Get-MsolUser -UserPrincipalName $objUser.UserPrincipalName
		$UsrLicense.ServiceStatus
		$NoLicenses = $UsrLicense.Licenses.Count
		
		$rpt = $UserLicenseTest.Licenses
		$cnt = 0
		foreach ($rpt in $rpt)
		{
		If ($rpt.AccountSkuId -eq $SSKID)
		{
			$PlanStats = $UserLicenseTest.Licenses[$Cnt].ServiceStatus

			$text = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}" -f $objUser.UserPrincipalName, $PlanStats.objUser.IsLicensed, $objUser.DisplayName, $objLic.AccountSku.SkuPartNumber, $usrmbx.CustomAttribute1, $usrmbx.CustomAttribute3, $usrmbx.CustomAttribute4, $stsmbx.LastLogonTime, $usrmbx.CustomAttribute15, $usrmbx.WhenMailboxCreated, $usrmbx.ForwardingSMTPAddress
		
		Out-File -FilePath $strPath -InputObject $text -Append
	}
}


$rpt = $UserLicenseTest.Licenses
$cnt = 0
foreach ($rpt in $rpt)
{
    If ($rpt.AccountSkuId -eq $SSKID)
    {
        $PlanStats = $UserLicenseTest.Licenses[$Cnt].ServiceStatus
        $PlanStats
        write-host
        $SCnt = 0
        $DisPlan = ""
        foreach ($PlanStats in $PlanStats)
        {
            If ($PlanStats.ProvisioningStatus -eq "Disabled")
            {
            switch($SCnt)
            {
                0
                {
                    write-host "SWAY is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $EnabSway = Read-Host
                    If ($EnabSway -ne "Y")
                    {
                        $DisPlan = "SWAY"
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
                1
                {
                    write-host "Mobile Device Management is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $EnabMDM = Read-Host
                    If ($EnabMDM -ne "Y")
                    {
                        if ($DisPlan -eq "")
                        {
                            $DisPlan = "INTUNE_0365"
                        }
                        else
                        {
                            $DisPlan = $DisPlan + ",INTUNE_0365"
                        }
                    } 
                    else
                    {
                        $Chg = "Y"
                    }
                }
                2
                {
                    Write-Host "Yammer Enterprise is Disabledwould you like to Enable this license (Y/N)? " -NoNewline
                    $EnabYam = Read-Host
                    if ($EnabYam -ne "Y")
                    {
                        if ($DisPlan -eq "")
                        {
                            $DisPlan = "YAMMER_ENTERPRISE"
                        }
                        else
                        {
                            $DisPlan = $DisPlan + ",YAMMER_ENTERPRISE"
                        }
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
                3
                {
                    Write-Host "Azure Rights Management is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $EnabRMS = Read-host
                    if ($EnabRMS -ne "Y")
                    {
                        if ($DisPlan -eq "")
                        {
                            $DisPlan = "RMS_S_ENTERPRISE"
                        }
                        else
                        {
                            $DisPlan = $DisPlan + ",RMS_S_ENTERPRISE"
                        }
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
                4
                {
                    Write-Host "Skype for Business Oneline (Plan3) is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $EnabSKY3 = Read-Host
                    if ($EnabSKY3 -ne "Y")
                    {
                        if ($DisPlan -eq "")
                        {
                            $DisPlan = "MCOVOICECONF"
                        }
                        else
                        {
                            $DisPlan = $DisPlan + ",MCOVOICECONF"
                        }
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
                5
                {
                    Write-Host "Office 365 ProPlus is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $EnabPro = Read-Host
                    if ($EnabPro -ne "Y")
                    {
                        if ($DisPlan -eq "")
                        {
                            $DisPlan = "OFFICESUBSCRIPTION"
                        }
                        else
                        {
                            $DisPlan = $DisPlan + ",OFFICESUBSCRIPTION"
                        }
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
                6
                {
                    Write-Host "Skype for Business Online (Plan2) is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $EnabSKY2 = Read-Host
                    if ($EnabSKY2 -ne "Y")
                    {
                        if ($DisPlan -eq "")
                        {
                            $DisPlan = "MCOSTANDARD"
                        }
                        else
                        {
                            $DisPlan = $DisPlan + ",MCOSTANDARD"
                        }
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
                7
                {
                    Write-Host "Office Online is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $EnabONL = Read-Host
                    if ($EnabONL -ne "Y")
                    {
                        if ($DisPlan -eq "")
                        {
                            $DisPlan = "SHAREPOINTWAC"
                        }
                        else
                        {
                            $DisPlan = $DisPlan + ",SHAREPOINTWAC"
                        }
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
                8
                {
                    Write-Host "SharePoint Oneline (Plan2)/OneDrive is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $EnabSP2 = Read-Host
                    if ($EnabSP2 -ne "Y")
                    {
                        if ($DisPlan -eq "")
                        {
                            $DisPlan = "SHAREPOINTENTERPRISE"
                        }
                        else
                        {
                            $DisPlan = $DisPlan + ",SHAREPOINTENTERPRISE"
                        }
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
                9
                {
                    Write-Host "Exchange Online Plan2 is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $EnabEX2 = Read-Host
                    if ($EnabEX2 -ne "Y")
                    {
                        if ($DisPlan -eq "")
                        {
                            $DisPlan = "EXCHANGE_S_ENTERPRISE"
                        }
                        else
                        {
                            $DisPlan = $DisPlan + ",EXCHANGE_S_ENTERPRISE"
                        }
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
            }
            }

            $Scnt = $SCnt + 1
        }

        If ($Chg -eq "Y")
        {
            Write-Host "Disabled Plans" $DisPlan
            write-host
            $DLO = ($DisPlan.Split(“,”))
            $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
            Set-MsolUserLicense -UserPrincipalName $EmpNo –LicenseOptions $MyO365Sku
            Write-Host ""
            Write-Host "License Options Changed for" $EmpNo -ForegroundColor Red
            $UserLicenseTest = Get-MsolUser -UserPrincipalName $EmpNo
            $UserLicenseTest.ServiceStatus
            Write-Host ""
            Write-Host "Current License Assignment for " $EmpNo " " $UserLicenseTest.DisplayName -ForegroundColor Red
            $UserLicenseTest = Get-MsolUser -UserPrincipalName $EmpNo
            $UserLicenseTest.ServiceStatus
            $PlanStats = $UserLicenseTest.Licenses[$Cnt].ServiceStatus
            $PlanStats
        }
        else
        {
            write-host ""
            write-host "No License Changes Made to current License Assignments for" $EmpNo -ForegroundColor Red
            write-host ""
        }
    }
    $cnt = $cnt + 1
}