$Usr = import-csv c:\temp\DisabledAccts.csv
$usr.count
foreach ($usr in $usr) 
{
	$usr.Alias
	$usrAlias = $usr.alias + "@global.ul.com"
	$userLicenseTest = Get-MsolUser -UserPrincipalName $UsrAlias
	$userLicenseTest.ServiceStatus
	$userLicenseTest.Licenses |ft Accountskuid
	write-host ""
}