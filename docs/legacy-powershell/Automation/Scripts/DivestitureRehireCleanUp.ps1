$EmpNo = read-host "Enter Employee Number"
Do
{
    $UPN = $EmpNo + "@global.ul.com"
    $ADGroupMem = Get-ADPrincipalGroupMembership $EmpNo -ResourceContextServer global.ul.com
    $AZGroupOwner = Get-AzureADUserOwnedObject -ObjectId (Get-MsolUser -userPrincipalName $UPN).ObjectID
    $DLMember = Get-AzureADUser -SearchString $EmpNo | Get-AzureADUserMembership | Where-Object {(($_.ObjectType -ne "Role") -and ($_.ObjectType -ne "Application"))} |Sort-Object DisplayName
    $MUsr = [bool]($DistName = (get-User $EmpNo -ErrorAction SilentlyContinue).DistinguishedName)
    If ($MUsr -eq $True)
    {
        If ($DistName -like "*'*")
        {
	        $DistName = $Distname.Replace("'","")
        }
        $DLOwner = (Get-Recipient -Filter "ManagedBy -eq '$DistName'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup | Select Name,RecipientTypeDetails)
    }
    else
    {
        write-host "This user is not enabled for Email"
        $DLOwner = ""
    }
    $transfile = $EmpNo + ".log"
    Start-transcript $TransFile

    write-host "`n" $DistName
    write-host "`nAD Group Member"
    $ADGroupMem.Name

    write-host "`nGroup Owner"
    $AZGroupOwner.DisplayName

    write-host "`nO365 Group Member"
    Foreach ($d in $DLMember)
    {
	    If (($D.DisplayName -like "MBX.*") -or ($D.DisplayName -like "LST*") -or ($D.DisplayName -like "PRM*"))
	    {
		    Remove-DistributionGroupMember $d.DisplayName -Member $upn -BypassSecurityGroupManagerCheck -confirm:$False
		    write-host "Removed from membership of: " $D.DisplayName
	    }
	    else
	    {
		    If ($D.DisplayName -like "GRP*")
		    {
			    Remove-UnifiedGroupLinks $D.ObjectID -LinkType Member -Links $EmpNo -Confirm:$False
			    write-host "Removed from membership of: " $D.DisplayName
		    }
		    else
		    {
			    write-host "Review this for manual removal: " $D.DisplayName
		    }
	    }
    }

    write-host "`nO365 Group Owner"
    $DLOwner.Name

    stop-transcript
    $EmpNo = read-host "Enter Employee Number"
} While ($EmpNo -ne "0")

