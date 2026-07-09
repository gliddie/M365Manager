#$OED1 = Import-Csv "\\usnbks140p-spa\itd1\SharedFiles\OED_Extract6.csv" |Sort-Object "Assignment Status","ALPHA_HR_ORG_LEVEL_4" -Unique
$AO4Groups = Get-AzureADMSGroup -SearchString "DSG.AO4"
$cnt = 0
Foreach ($OED in $OED1)
{
    If (($OED."ASSIGNMENT STATUS" -eq "A") -and ($OED."ALPHA_HR_ORG_LEVEL_4" -ne ""))
    {
#        write-host $OED."Assignment Status",$OED."ALPHA_HR_ORG_LEVEL_4"
        $bute4 = $OED."ALPHA_HR_ORG_LEVEL_4"
        $AO4 = $OED."ALPHA_HR_ORG_LEVEL_4" -replace (",","")
        $EmpGrpName = "DSG.AO4." + $AO4 + " UL Employees Only"
        $EmpGrpNick =  ("DSG.AO4." + $AO4) + "EmpOnly" -replace("[ ]","")
        $EmpGrpDesc = "All UL Employees in Alpha Org 4 " +  $OED."ALPHA_HR_ORG_LEVEL_4"
        $Found = ""
        Foreach ($chk in $AO4Groups)
        {
            If ($chk.DisplayName -eq $EmpGrpName)
            {
                $Found = "Yes"
            }
        }
        If ($Found -eq "")
        {
            Write-host "`nCreating" $EmpGrpName -ForegroundColor Green
            New-AzureADMSGroup -Description $EmpGrpDesc -DisplayName $EmpGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $EmpGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute1 -contains ""Employee"") and (user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute4 -eq ""$bute4"")" -MembershipRuleProcessingState "On"
        }
        else
        {
            Write-host "`nGroup" $EmpGrpName "already exists" -ForegroundColor Red
        }
        $StaffGrpName = "DSG.AO4." + $AO4 + " UL Staff"
        $StaffGrpNick = "DSG.AO4." + $AO4 + "Staff" -replace("[ ]","")
        $StaffGrpDesc = "All UL Staff in Alpha Org 4 " +  $OED."ALPHA_HR_ORG_LEVEL_4"
        $Found = ""
        Foreach ($chk in $AO4Groups)
        {
            If ($chk.DisplayName -eq $StaffGrpName)
            {
                $Found = "Yes"
            }
        }
        If ($Found -eq "")
        {
            Write-host "Creating" $StaffGrpName -ForegroundColor Green
            New-AzureADMSGroup -Description $StaffGrpDesc -DisplayName $StaffGrpName -MailEnabled $false -SecurityEnabled $true -MailNickname $StaffGrpNick -GroupTypes "DynamicMembership" -MembershipRule "(user.extensionAttribute2 -eq ""A"") and (user.extensionAttribute4 -eq ""$bute4"")" -MembershipRuleProcessingState "On"
        }
        else
        {
            Write-host "Group" $StaffGrpName "already exists" -ForegroundColor Red
        }

    }
}
# Get All Azure DST.A04 groups
#Get-AzureADMSGroup -SearchString "DSG.AO4"
#Set-AzureADMSGroup -Id $grp.id -MembershipRuleProcessingState "On"

#$val = Get-AzureADGroupMember -ObjectId 6b4f9304-d428-44df-bf80-c4cbcdf4590d #to get group membership count