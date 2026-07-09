Start-Transcript -Path .\CreateoEDAudit.Log

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

Function PrtData
{
    $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment,$RespGrp
    Out-File -FilePath $strPath -InputObject $text -Append
    $RespGrp = "BU"    
}


$Date = get-date -Format "yyyy-MMdd"

$strPath = "\\usnbku134p\c$\scripts\OrgCharts\OEDAuditReportwCharaCheck.csv"
$AuditHistory = "\\usnbku134p\c$\scripts\OrgCharts\OrgChartHistory\OEDAuditReport-" + $date + ".csv"
$text = "EmpNo;EmpType;FullName;Location;Supervisor;SupEmpNo;Supervisor2;Sup2EmpNo;Supervisor3;Sup3EmpNo;Comment;ResponsibleGroup"

Out-File -FilePath $strPath -InputObject $text

$EmpNo = [string]
$SupEmpNo = [string]
$Sup2EmpNo = [string]
$Sup3EmpNo = [string]
$DupEmpNo = "N"
$DupChecked = "0"
$RespGrp = "BU"

$OEDDetails = Import-CSV \\usnbks181p\itd1\SharedFiles\OED_Extract5.csv |Sort-Object "ASSIGNMENT STATUS","EMPLOYEE NUMBER"
$DupCheck = Import-CSV \\usnbks181p\itd1\SharedFiles\OED_Extract5.csv |Sort-Object "EMPLOYEE NUMBER","ASSIGNMENT STATUS"

$var = $OEDDetails."EMPLOYEE NUMBER" |Select-object -unique

$RecDups = Get-Duplicate $OEDDetails."EMPLOYEE NUMBER" -count |sort-object "VALUE"

foreach ($RecDups in $RecDups)
{
	foreach ($User in $OEDDetails)
	{
		if ($RecDups.Value -eq $User."EMPLOYEE NUMBER")
		{
            $RespGrp = "HR"
            write-host "Duplicate Records for Employee:" $User."FULL NAME" "Employee No:" $User."EMPLOYEE NUMBER"
            $Comment = "Duplicate Records Exist"
            PrtData ($User, $text)
#            $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment,$RespGrp
#            Out-File -FilePath $strPath -InputObject $text -Append
		}
	}
}

Foreach ($user in $OEDDetails)
{

#   Check details for Active staff
    If ($User."ASSIGNMENT STATUS" -eq "A")
    {
#	 # Check if there are non-printable characters in the Users Name
		$pattern = '[a-zA-Z0-9().,-]'
		$NameCheck = $user."FULL NAME" -replace $pattern,''
	
		If ($NameCheck.IndexOf("'") -ge 0)
		{
			$NameCheck = ($NameCheck.Replace("'","")).Trim()
		}

		If ($NameCheck.Trim().Length -gt 0)
		{
#			write-host "There are nonprintable characters in the name" $User."FULL NAME"
#		    write-host "Non-printable characters in the full name:" $User."FULL NAME" "Employee No:" $User."EMPLOYEE NUMBER"
            If ($User."PERSON TYPE" -eq "Employee")
            {
                $RespGrp = "HR"
            }

		    $Comment = "Non-printable characters in the name"
		    $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment,$RespGrp
		    Out-File -FilePath $strPath -InputObject $text -Append
		}
		
#    # Check if Location is Blank
        If ($User."LOCATION".Length -lt 1)
        {
            write-host "Location is Blank for Employee:" $User."FULL NAME"
            If ($User."PERSON TYPE" -eq "Employee")
            {
                $RespGrp = "HR"
            }
            else
            $Comment = "Location is Blank"
            $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment
            Out-File -FilePath $strPath -InputObject $text -Append
        }

    # Check if Employee Name is in correct format
        If ($ENComma = $User."FULL NAME".IndexOf(",") -lt 1)
        {
            write-host "Employee Name not in corect format for Employee:" $User."FULL NAME"
            $Comment = "Employee Name Bad Format"
            $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment
            Out-File -FilePath $strPath -InputObject $text -Append
        }
    
    # Check if Supervisor Name is in correct format
        If ($ENComma = $User."SUPERVISOR".IndexOf(",") -lt 1)
        {
            If (($User."SUPERVISOR" -ne "Williams, Keith") -and ($User."FULL NAME" -ne "Williams, Keith"))
            {
                If ($User."SUPERVISOR".Length -lt 1)
                {
                    $Comment = "Supervisor Name Blank"
                    write-host "Supervisor Name Blank for Employee:" $User."FULL NAME"
                }
                else
                {
                    $Comment = "Supervisor Name Bad Format"
                    write-host "Supervisor Name not in corect format for Employee:" $User."FULL NAME" "Supervisor Assigned:" $User."SUPERVISOR"
                }
                $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment
                Out-File -FilePath $strPath -InputObject $text -Append
            }
        }

#   # Check if Supervisor Emp# is not blank
        If ($User."SUPERVISOR EMP NUMBER".Length -lt 1)
        {
            If (($User."SUPERVISOR" -ne "Williams, Keith") -and ($User."FULL NAME" -ne "Williams, Keith"))
            {
                write-host "Supervisor Employee Number Blank for Employee:" $User."FULL NAME"
                $Comment = "Supervisor Employee Number Blank"
                $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment
                Out-File -FilePath $strPath -InputObject $text -Append
            }
        }
        else
        {
#  # Check if Supervisor is Active
            Foreach ($UserDup in $DupCheck)
            {
                If ($UserDup.'ASSIGNMENT STATUS' -eq "T")
                {
                    If ($User."SUPERVISOR EMP NUMBER" -eq $UserDup.'EMPLOYEE NUMBER')
                    {
                        write-host "Assigned to Terminated Supervisor for Employee:" $User."FULL NAME" "Assigned to: " $User."SUPERVISOR EMP NUMBER"
                    }
                }
            }
        }
#  # Check if Supervisor2 Name is in correct format
        If ($ENComma = $User."SUPERVISOR_2".IndexOf(",") -lt 1)
        {
            If (($User."SUPERVISOR" -ne "Williams, Keith") -and ($User."FULL NAME" -ne "Williams, Keith"))
            {
                If ($User."SUPERVISOR_2".Length -lt 1)
                {
                    $Comment = "Supervisor2 Name Blank"
                    write-host "Supervisor2 Name Blank for Employee:" $User."FULL NAME"
                }
                else
                {
                    $Comment = "Supervisor2 Name Bad Format"
                    write-host "Supervisor2 Name not in corect format for Employee:" $User."FULL NAME" "Supervisor Listed " $User."SUPERVISOR_2"
                }
                $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment
                Out-File -FilePath $strPath -InputObject $text -Append
            }
        }

#   # Check if Supervisor2 Emp# is not blank
        If ($User."SUPERVISOR_2_NUMBER".Length -lt 1)
        {
            If (($User."SUPERVISOR" -ne "Williams, Keith") -and ($User."FULL NAME" -ne "Williams, Keith"))
            {
                write-host "Supervisor2 Employee Number Blank for Employee:" $User."FULL NAME"
                $Comment = "Supervisor2 Employee Number Blank"
                $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment
                Out-File -FilePath $strPath -InputObject $text -Append
            }
        }
        else
        {
#  # Check if Supervisor2 is Active
            Foreach ($UserDup in $DupCheck)
            {
                If ($UserDup.'ASSIGNMENT STATUS' -eq "T")
                {
                    If ($User."SUPERVISOR_2_NUMBER" -eq $UserDup.'EMPLOYEE NUMBER')
                    {
                        write-host "Assigned to Terminated Supervisor2 for Employee:" $User."FULL NAME" "Assigned to: " $User."SUPERVISOR_2_NUMBER"
                    }
                }
            }
        }

#   # Check if Supervisor3 Name is in correct format
        If ($ENComma = $User."SUPERVISOR_3".IndexOf(",") -lt 1)
        {
            If (($User."SUPERVISOR" -ne "Williams, Keith") -and ($User."FULL NAME" -ne "Williams, Keith"))
            {
                If ($User."SUPERVISOR_3".Length -lt 1)
                {
                    $Comment = "Supervisor3 Name Blank"
                    write-host "Supervisor3 Name Blank for Employee:" $User."FULL NAME"
                }
                else
                {
                    $Comment = "Supervisor3 Name Bad Format"
                    write-host "Supervisor3 Name not in corect format for Employee:" $User."FULL NAME" "Supervisor Name Listed:" $User."SUPERVISOR_3"
                }
                $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment
                Out-File -FilePath $strPath -InputObject $text -Append
            }
        }

#   # Check if Supervisor3 Emp# is not blank
        If ($User."SUPERVISOR_3_NUMBER".Length -lt 1)
        {
            If (($User."SUPERVISOR" -ne "Williams, Keith") -and ($User."FULL NAME" -ne "Williams, Keith"))
            {
                write-host "Supervisor3 Employee Number Blank for Employee:" $User."FULL NAME"
                $Comment = "Supervisor3 Employee Number Blank"
                $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment
                Out-File -FilePath $strPath -InputObject $text -Append
            }
        }
        else
        {
#  # Check if Supervisor3 is Active
            Foreach ($UserDup in $DupCheck)
            {
                If ($UserDup.'ASSIGNMENT STATUS' -eq "T")
                {
                    If ($User."SUPERVISOR_3_NUMBER" -eq $UserDup.'EMPLOYEE NUMBER')
                    {
                        write-host "Assigned to Terminated Supervisor3 for Employee:" $User."FULL NAME" "Assigned to:" $User."SUPERVISOR_3_NUMBER"
                    }
                }
            }
        }
    }
  }

copy-item $strPath -Destination $AuditHistory
	
Stop-Transcript