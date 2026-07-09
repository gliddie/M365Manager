Start-Transcript -Path .\CreateoEDAudit.Log

$strPath = "\\usnbku134p\c$\scripts\OrgCharts\OEDAuditReport.csv"
$text = "EmpNo;EmpType;FullName;Location;Supervisor;SupEmpNo;Supervisor2;Sup2EmpNo;Supervisor3;Sup3EmpNo;Comment"

Out-File -FilePath $strPath -InputObject $text

$EmpNo = [string]
$SupEmpNo = [string]
$Sup2EmpNo = [string]
$Sup3EmpNo = [string]
$DupEmpNo = "N"
$DupChecked = "0"

$OEDDetails = Import-CSV \\usnbks181p\itd1\SharedFiles\OED_Extract5.csv |Sort-Object "ASSIGNMENT STATUS","EMPLOYEE NUMBER"
$DupCheck = Import-CSV \\usnbks181p\itd1\SharedFiles\OED_Extract5.csv |Sort-Object "EMPLOYEE NUMBER","ASSIGNMENT STATUS"

$OEDDetails.count
$DupCheck.count

Foreach ($user in $OEDDetails)
{

#   Check details for Active staff

    If ($User."ASSIGNMENT STATUS" -eq "A")
    {
#   Check if this is a Duplicate
        $DupStart = "Y"
        $Cnt = 0
        If ($DupChecked -ne $User."EMPLOYEE NUMBER")
        {
            foreach ($UserDup in $DupCheck)
            {
                if ($User."EMPLOYEE NUMBER" -eq $UserDup."EMPLOYEE NUMBER")
                {
                    $DupStart = "N"
                    If ($DupStart -eq "N")
                    {
                        If ($Cnt -gt 0)
                        {
                            If ($cnt -eq 1)
                            {
                                write-host "Duplicate Records for Employee:" $User."FULL NAME" "Employee No:" $User."EMPLOYEE NUMBER"
                                $Comment = "Duplicate Records Exist"
                                $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment
                                Out-File -FilePath $strPath -InputObject $text -Append
                            }
                            else
                            {
                                write-host "Duplicate Records for Employee:" $UserDup."EMPLOYEE NUMBER" "Assignment Status:" $UserDup."ASSIGMENT STATUS"
                                $Comment = "Duplicate Records Exist"
                                $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $UserDup."EMPLOYEE NUMBER",$UserDup."PERSON TYPE",$UserDup."FULL NAME",$UserDup."LOCATION",$UserDup."SUPERVISOR",$UserDup."SUPERVISOR EMP NUMBER",$UserDup."SUPERVISOR_2",$UserDup."SUPERVISOR_2_NUMBER",$UserDup."SUPERVISOR_3",$UserDup."SUPERVISOR_3_NUMBER",$Comment
                                Out-File -FilePath $strPath -InputObject $text -Append
                                $DupChecked = $User."EMPLOYEE NUMBER"
                            }
                        }
                        $Cnt = $Cnt + 1
                    }
                }
            }
        }

#    # Check if Location is Blank
        If ($User."LOCATION".Length -lt 1)
        {
            write-host "Location is Blank for Employee:" $User."FULL NAME"
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
    else
    {
#    Check for terminated users with duplicate records for non Active Staff
        $DupChecked = ""
        $DupStart = "Y"
        $Cnt = 0
        write-host "Checking for duplicated Terminated Staff Records" -ForegroundColor Cyan

        If ($DupChecked -ne $User."EMPLOYEE NUMBER")
        {
            foreach ($UserDup in $DupCheck)
            {
                if (($User."EMPLOYEE NUMBER" -eq $UserDup."EMPLOYEE NUMBER"))
                {
                    $DupStart = "N"
                    If ($DupStart -eq "N")
                    {
                        If ($Cnt -gt 0)
                        {
                            If ($cnt -eq 1)
                            {
                                write-host "Duplicate Terminated Records for Staff for Employee:" $User."FULL NAME" "Employee No:" $User."EMPLOYEE NUMBER"
                                $Comment = "Duplicate Terminated Records Exist for Staff "
                                $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $User."EMPLOYEE NUMBER",$User."PERSON TYPE",$User."FULL NAME",$User."LOCATION",$User."SUPERVISOR",$User."SUPERVISOR EMP NUMBER",$User."SUPERVISOR_2",$User."SUPERVISOR_2_NUMBER",$User."SUPERVISOR_3",$User."SUPERVISOR_3_NUMBER",$Comment
                                Out-File -FilePath $strPath -InputObject $text -Append
                            }
                            else
                            {
                                write-host "Duplicate Terminated Records for Employee:" $UserDup."EMPLOYEE NUMBER" "Assignment Status:" $UserDup."ASSIGMENT STATUS"
                                $Comment = "Duplicate Terminated Records Exist"
                                $text = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}" -f $UserDup."EMPLOYEE NUMBER",$UserDup."PERSON TYPE",$UserDup."FULL NAME",$UserDup."LOCATION",$UserDup."SUPERVISOR",$UserDup."SUPERVISOR EMP NUMBER",$UserDup."SUPERVISOR_2",$UserDup."SUPERVISOR_2_NUMBER",$UserDup."SUPERVISOR_3",$UserDup."SUPERVISOR_3_NUMBER",$Comment
                                Out-File -FilePath $strPath -InputObject $text -Append
                                $DupChecked = $User."EMPLOYEE NUMBER"
                            }
                        }
                        $Cnt = $Cnt + 1
                    }
                }
            }
        }
    }
  }
	
Stop-Transcript