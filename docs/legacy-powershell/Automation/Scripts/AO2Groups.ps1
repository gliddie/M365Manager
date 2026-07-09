$OED1 = Import-Csv "\\usnbks140p-spa\itd1\SharedFiles\OED_Extract6.csv" |Sort-Object "Assignment Status","ALPHA_HR_ORG_LEVEL_2" -Unique
$AO2Groups = Get-AzureADMSGroup -SearchString "DSG.AO2"
$cnt = 0
Foreach ($OED in $OED1)
{
    If (($OED."ASSIGNMENT STATUS" -eq "A") -and ($OED."ALPHA_HR_ORG_LEVEL_2" -ne ""))
    {
        $bute8 = $OED."ALPHA_HR_ORG_LEVEL_2"
        $AO2 = $OED."ALPHA_HR_ORG_LEVEL_2" -replace (",","")
        $EmpGrpName = "DSG.AO2." + $AO2 + " UL Employees Only"
        $EmpGrpNick =  ("DSG.AO2." + $AO2) + "EmpOnly" -replace("[ ]","")
        $EmpGrpDesc = "All UL Employees in Alpha Org 2 " +  $OED."ALPHA_HR_ORG_LEVEL_2"
        $Found = ""
        Foreach ($chk in $AO2Groups)
        {
            If ($chk.DisplayName -eq $EmpGrpName)
            {
                $Found = "Yes"
            }
        }
        If ($Found -eq "")
        {
            Write-host "`nCreating" $EmpGrpName -ForegroundColor Green
            New-AzureADMSGroup -Description $EmpGrpDesc -DisplayName $EmpGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $EmpGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute1 -contains ""Employee"") and (user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute8 -eq ""$bute8"")" -MembershipRuleProcessingState "On"
        }
        else
        {
            Write-host "`nGroup" $EmpGrpName "already exists" -ForegroundColor Red
        }
        $StaffGrpName = "DSG.AO2." + $AO2 + " UL Staff"
        $StaffGrpNick = "DSG.AO2." + $AO2 + "Staff" -replace("[ ]","")
        $StaffGrpDesc = "All UL Staff in Alpha Org 2 " +  $OED."ALPHA_HR_ORG_LEVEL_2"
        $Found = ""
        Foreach ($chk in $AO2Groups)
        {
            If ($chk.DisplayName -eq $StaffGrpName)
            {
                $Found = "Yes"
            }
        }
        If ($Found -eq "")
        {
            Write-host "Creating" $StaffGrpName -ForegroundColor Green
            New-AzureADMSGroup -Description $StaffGrpDesc -DisplayName $StaffGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $StaffGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute8 -eq ""$bute8"")" -MembershipRuleProcessingState "On"
        }
        else
        {
            Write-host "Group" $StaffGrpName "already exists" -ForegroundColor Red
        }

    }
}