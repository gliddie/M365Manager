$pattern = '[a-zA-Z0-9().,-]'
#$NameCheck = $user."FULL NAME" -replace $pattern,''
#$NameCheck = "Taurose'vičiūtė, Indré1"
$NameCheck = "Hansen-O'Glazebrook,. Sandra1 (Sandi)"
write-host "NameCheck= " $NameCheck
#$NameCheck = $User."FULL NAME"
#$NameCheck = $user."FULL NAME" -replace $pattern,''
$NameCheck = $NameCheck -replace $pattern,''
write-host "NameCheck= " $NameCheck
$NameCheck.IndexOf("'")
If ($NameCheck.IndexOf("'") -eq 0)
{
	$NameCheck = ($NameCheck.Replace("'","")).Trim()
#	write-host "NameCheck= !" $NameCheck "!"
#	write-host "NameCheckLength " $NameCheck.Length
}

If ($NameCheck.Length -le 0)
{
	write-host "NameCheck= " $NameCheck
#        write-host "There are nonprintable characters in the name" $User."FULL NAME"
#        write-host "Non-printable characters in the full name:" $User."FULL NAME" "Employee No:" $User."EMPLOYEE NUMBER"
#        $Comment = "Non-printable characters in the name"
#        $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment
#        Out-File -FilePath $strPath -InputObject $text -Append
}


