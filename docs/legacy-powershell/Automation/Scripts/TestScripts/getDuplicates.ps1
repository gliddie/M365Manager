$User = Import-CSV \\usnbks181p\itd1\SharedFiles\OED_Extract5.csv |Sort-Object "EMPLOYEE NUMBER"
#$User = Import-CSV \\usnbku134p\c$\scripts\OrgCharts\OED_ExtractDups.csv |Sort-Object "EMPLOYEE NUMBER"
$var = $User."EMPLOYEE NUMBER" |Select-object -unique

Function Get-Duplicate {
    param($array, [switch]$count)
    begin {
        $hash = @{}
    }
    process {
        $array | %{ $hash[$_] = $hash[$_] + 1 }
        if($count) {
            $hash.GetEnumerator() | ?{$_.value -gt 1} | %{
                New-Object PSObject -Property @{
                    Value = $_.key
                    Count = $_.value
                }
            }
        }
        else {
            $hash.GetEnumerator() | ?{$_.value -gt 1} | %{$_.key}
        }    
    }
}

$RecDups = Get-Duplicate $User."EMPLOYEE NUMBER" -count |sort-object "VALUE"

foreach ($RecDups in $RecDups)
{

	foreach ($UserRec in $user)
	{
		if ($RecDups.Value -eq $UserRec."EMPLOYEE NUMBER")
		{
			write-host $UserREc."EMPLOYEE NUMBER",$UserRec."FULL NAME",$UserRec."PERSON TYPE"
		}
	}
}