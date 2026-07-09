$mbx = Import-csv c:\temp\LicensedUsers.csv

$Global:DisPlanEmpE4    = ""
$Global:DisPlanNonEmpE4 = ""
$Global:DisPlanEmpE3 	 = ""
$Global:DisPlanNonEmpE3 = ""

Invoke-Expression -Command E:\O365AdminShared\Scripts\O365DisabledLicenseFeatures.ps1
			
foreach ($mbx in $mbx)
{

	#$UserDetails = get-mailbox $mbx.UID -ErrorAction SilentlyContinue
	$UserLicenseTest = Get-MsolUser -UserPrincipalName $mbx.UserPrincipalName
	$LDAPFilter = "(userPrincipalName=" + $mbx.UserPrincipalName + ")"
	$ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,proxyaddresses,extensionattribute1,extensionattribute4

#	Write-Host ""
#	Write-Host "Resetting License for " $Mbx.UserPrincipalName " - " $mbx.DisplayName "sublicense features" $ADUser.ExtensionAttribute1 -ForegroundColor Green
	$UserLicenseTest.ServiceStatus
	#$UserLicenseTest.Licenses |ft AccountSkuId

	#$LicDet = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEWITHSCAL"}

		If ($ADUser.ExtensionAttribute1 -like "Employee*")
		{
			Write-Host "Resetting E4 License to standard employee features for" $Mbx.UserPrincipalName "-" $mbx.DisplayName "sublicense features" $ADUser.ExtensionAttribute1 -ForegroundColor Green
            $DisPlan = $Global:DisPlanEmpE4
		}
		else
		{
			Write-Host "Resetting E4 License to standard non-employee features for" $Mbx.UserPrincipalName "-" $mbx.DisplayName "sublicense features" $ADUser.ExtensionAttribute1 -ForegroundColor Yellow
			$DisPlan = $Global:DisPlanEmpE3
		}

	$DLO = ($DisPlan.Split(“,”))
	$MyO365Sku = New-MsolLicenseOptions -AccountSkuId "ul:ENTERPRISEWITHSCAL" -DisabledPlans $DLO
	Set-MsolUserLicense -UserPrincipalName $mbx.UserPrincipalName –LicenseOptions $MyO365Sku
						
}