$usr = import-csv biinput.csv
$USR.COUNT
PAUSE
foreach ($usr in $usr)
{ set-msoluser -userprincipalname $usr.upn -usagelocation "US"
  write-host $usr.upn
  Set-MsolUserLicense -UserPrincipalName $usr.upn -AddLicenses "ul:POWER_BI_STANDARD"
}