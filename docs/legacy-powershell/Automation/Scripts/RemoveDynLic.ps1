Function O365Licenses
{
    $O365Lic = (Get-MsolUser -UserPrincipalName $EmpNo).Licenses
	Write-Host "`nCurrent License Assignment for " $EmpNo " " $UserLicense.DisplayName "-" $ADUser.ExtensionAttribute1-ForegroundColor Green
    
    If ($O365Lic.count -gt 0)
    {
        foreach ($O365Lic in $O365Lic)
        {
            switch ($O365Lic.AccountSkuID)
            {
                "ul:ENTERPRISEPACK"
                {
                    $text = "`tEnterprise E3"
                }
                "ul:EXCHANGEENTERPRISE"
                {
                    $text = "`tExchange Online Plan 2"
                }
                "ul:POWER_BI_STANDARD"
                {
                    $text = "`tPowerBI (Free)"
                }
                "ul:POWER_BI_PRO"
                {
                    $text = "`tPowerBI Pro" 
                }
                "ul:DYN365_ENTERPRISE_PLAN1"
                {
                    $text = "`tDynamics 365 Customer Engagement Plan"
                }
                "ul:EMS"
                {
                    $text = "`tEnterprise Mobility Suite/Intune (EMS)"
                }
                "ul:POWER_BI_INDIVIDUAL_USER"
                {
                    $text = "`tPower BI for O365"
                }
                "ul:MCOIMP"
                {
                   $text = "`tSkype for Business Online (Plan 1)"
                }
                "ul:SHAREPOINTSTANDARD"
                {
                    $text = "`tSharePoint Online (Plan 1)"
                }
                "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
                {
                    $text = "`tMicrosoft Relationship Sales Solution"
                }
                "ul:TEAMS_COMMERCIAL_TRIAL"
                {
                    $text = "`tMicrosoft Teams Commercial Cloud (User Initiated)"
                }
                "ul:ATP_ENTERPRISE"
                {
                    $text ="`tAdvanced Threat Protection Plan1"
                }
                "ul:WIN_DEF_ATP"
                {
                    $text ="`tDefender Advanced Threat Protection"
                }
                "ul:FLOW_P2"
                {
                    $text = "`tMicrosoft Flow Plan 2"
                }
                "ul:MEETING_ROOM"
                {
                    $text = "`tMeeting Room"
                }
                "ul:POWERFLOW_P2"
                {
                    $text = "`tPowerApps Plan 2"
                }
                "ul:MCOEV"
                {
                    $text = "`tPhone System"
                }
                "ul:MCOMEETADV"
                {
                    $text = "`tAudio Conferencing"
                }
            }
            Write-Host $text
        }
    }
    else
    {
        Write-Host "`n`tNo licenses assigned to this account" -ForegroundColor Red

    }
}

$inp = import-csv "c:\temp\RemoveDynLic.csv"
$RemCRMLic = 0
$RemMRSSLic = 0
$NoLicAss = 0
$NoAccounts = 0
$NoUserAccount = 0

foreach ($inp in $inp)
{
	$NoAccounts++
    $Global:ENo = $inp.UID
	$EmpNo = $ENo + "@global.ul.com"
    
	if ([bool](Get-MsolUser -UserPrincipalName $EmpNo -ErrorAction SilentlyContinue) -eq "True")
	{
	    $UserLicense = Get-MsolUser -UserPrincipalName $EmpNo

        $HasCRM = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:DYN365_ENTERPRISE_PLAN1"})
        $HasMRSS = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"})
			
	    $LDAPFilter = "(userPrincipalName=" + $EmpNo + ")"
	    $ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,proxyaddresses,extensionattribute1,extensionattribute4
        $O365Lic = (Get-MsolUser -UserPrincipalName $EmpNo).Licenses
#        Write-Host "`nReviewing License Assignment for " $EmpNo " " $UserLicense.DisplayName "-" $ADUser.ExtensionAttribute1-ForegroundColor Green
#        O365Licenses

        if ($HasCRM -eq "True")
        {
#    		write-host "Removing CRM/Dynamics License" -ForegroundColor Yellow
#            Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:DYN365_ENTERPRISE_PLAN1"
    		write-host "CRM Lisense assigned to: " $EmpNo
            $RemCRMLic++
        }
	    elseif ($HasMRSS -eq "True")
        {
#    		write-host "Removing MRSS License" -ForegroundColor Yellow
#            Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
    		write-host "MRSS License assigned to: " $EmpNo
            $RemMRSSLic++
        }
	    else
	    {
		    write-host "No CRM or MRSS License Assigned to: " $EmpNo -ForegroundColor Cyan
            $NoLicAss++
	    }
    }
    else
    {
        write-host "No User Account exists for: " $EmpNo -ForegroundColor Red
        $NoUserAccount++
    }
}
write-host "CRM Licenses to be Removed: " $RemCRMLic
write-host "MRSS Licenses to be Removed: " $RemMRSSLic
write-host "No CRM or MRSS Licenses Assigned: " $NoLicAss
write-host "No User Account Exists: " $NoUserAccount