$Chg = "N"
#write-host "Enter Employee Number of Individual " -ForegroundColor Yellow -NoNewline
#$EmpNo = Read-Host
#$EmpNo = $EmpNo + "@global.ul.com"
$SSKID = "ul:ENTERPRISEWITHSCAL"
$UserLicenseTest = Get-MsolUser -UserPrincipalName $EmpNo
Write-Host "Current License Assignment for " $EmpNo " " $UserLicenseTest.DisplayName
$UserLicenseTest.ServiceStatus
# $UserLicenseTest.Licenses |ft AccountSkuId

$NoLicenses = $UserLicenseTest.Licenses.Count

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
                    write-host "Flow for Office 365 is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $Enab0365PP = Read-Host
                    If ($Enab0365PP -ne "Y")
                    {
                        $DisPlan = "FLOW_O365_P2"
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
				1
                {
                    write-host "PowerApps for Office 365 is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $Enab0365PP = Read-Host
                    If ($Enab0365PP -ne "Y")
                    {
                        $DisPlan = "POWERAPPS_O365_P2"
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
				2
                {
                    write-host "Microsoft Teams is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $Enab0365PP = Read-Host
                    If ($Enab0365PP -ne "Y")
                    {
                        $DisPlan = "TEAMS1"
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
				3
                {
                    write-host "0365 Plan Preview is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $Enab0365PP = Read-Host
                    If ($Enab0365PP -ne "Y")
                    {
                        $DisPlan = "PROJECTWORKMANAGEMENT"
                    }
                    else
                    {
                        $Chg = "Y"
                    }
                }
                4
                {
                    write-host "SWAY is Disabled would you like to Enable this license (Y/N)? " -NoNewline
                    $EnabSway = Read-Host
                    If ($EnabSway -ne "Y")
                    {
                        If ($DisPlan -eq "")
                        {
                            $DisPlan = "SWAY"
                        }
                        else
                        {
                            $DisPlan = $DisPlan + ",SWAY"
                        }
                    } 
                    else
                    {
                        $Chg = "Y"
                    }
                }
                5
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
                6
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
                7
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
                8
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
                9
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
                10
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
                11
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
                12
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
                13
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