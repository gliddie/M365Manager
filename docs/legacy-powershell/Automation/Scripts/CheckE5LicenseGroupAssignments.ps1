#Check Licenses

$LiveCred = Import-Clixml C:\Users\SVC.NBK.ADQueries\Documents\mySVCULONMICSFile.xml
Connect-MsolService -Credential $LiveCred
Connect-AzureAD –Credential $LiveCred

$Usr = Get-MsolUser -all | where { $_.IsLicensed -eq $True } |Sort-Object UserPrincipalName
$EmpGrpMem = Get-AzureADGroupMember -ObjectId 2aab81de-e54a-4c58-909c-b15273c54dff -all $true
$NonEmpGrpMem = Get-AzureADGroupMember -ObjectId 40dff561-964a-4f7b-9794-dba3542b63c4 -all $true

foreach ($u in $usr)
{
    $ADUsr = ""
    $EmpNo = $u.UserPrincipalName -Replace ("@global.ul.com","")
    If ($EmpNo.Length -le 6)
    {
        $ADUsrExists = ""
        $IsEmpLic = ""
        $IsNonEmpLic = ""
        $SrchVal = $EmpNo + "*"
        $ErrorActionPreference = "SilentlyContinue"
        $ADUsrExists = [bool]($ADUsr = get-ADUser $EmpNo -Properties ExtensionAttribute1)
        $ErrorActionPreference = "Continue"

        If ($ADUsrExists -eq $True)
        {
             $IsEmpLic = [bool]($EmpGrpMem.UserPrincipalName -like $SrchVal)
            $IsNonEmpLic = [bool]($NonEmpGrpMem.UserPrincipalName -like $SrchVal)

            If ($ADUsr.ExtensionAttribute1 -like "Emp*")
            {
                If ($IsNonEmpLic -eq $True)
                {
                    write-host "Processing $EmpNo -" $u.DisplayName "- [" $ADUsr.ExtensionAttribute1 "]"
                    write-host "`tAssigned Employee License $IsEmpLic" -ForegroundColor Green
                    write-host "`tAssigned Non-Employee License $IsNonEmpLic" -ForegroundColor Green
                    If ($IsEmpLic -eq $False)
                    {
                        Add-AzureADGroupMember -ObjectId 2aab81de-e54a-4c58-909c-b15273c54dff -RefObjectId $u.ObjectID
                        Write-host "`tAdded to Employee license group" -ForegroundColor Red
                    }
                    Remove-AzureADGroupMember -ObjectId 40dff561-964a-4f7b-9794-dba3542b63c4 -MemberID $u.ObjectID
                    Write-host "`tRemoved from the Non-Employee license group" -ForegroundColor Red
                }
            }
            else
            {
                If (($ADUsr.ExtensionAttribute1 -notlike "Ex-*") -and ($IsEmpLic -eq $True) -and ($IsNonEmpLic -eq $False))
                {
                    write-host "Processing $EmpNo -" $u.DisplayName "- [" $ADUsr.ExtensionAttribute1 "]"
                    write-host "`tAssigned Employee License $IsEmpLic"-ForegroundColor Cyan
                    write-host "`tAssigned Non-Employee License $IsNonEmpLic" -ForegroundColor Cyan
                    If ($IsNonEmpLic -eq $False)
                    {
                        Add-AzureADGroupMember -ObjectId 40dff561-964a-4f7b-9794-dba3542b63c4 -RefObjectId $u.ObjectID
                        Write-host "`tAdded to Non-Employee license group" -ForegroundColor Red
                    }
                    Remove-AzureADGroupMember -ObjectId 2aab81de-e54a-4c58-909c-b15273c54dff -MemberID $u.ObjectID
                    Write-host "`tRemoved from the Employee license group" -ForegroundColor Red
                }
            }
        }
        else
        {
            $mbx = get-mailbox $EmpNo
            write-host "Processing $EmpNo -" $u.DisplayName "- [" $Mbx.CustomAttribute1 "] - No AD Account found" -ForegroundColor Red
        }
    }
}
