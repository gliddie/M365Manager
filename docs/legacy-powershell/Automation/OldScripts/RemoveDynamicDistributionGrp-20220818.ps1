<###  This script remove UL Employee Only and UL Staff Dynamic Distribution Lists
#
#  Called by:  DistributionnSecurityGroupMenu.ps1
#
#>

write-host "`n     Enter ( 1) to remove DST.All group"
write-host "           ( 2) to remove DST.SUP group"
write-host "           ( 3) to remove DSG.AO2 groups"
write-host "           ( 4) to remove DSG.AO4 groups"
write-host ""
write-Host "           ( 0) to Return to the Distribution Group Admin Menu"
write-Host "     Enter Option? " -ForegroundColor Red -NoNewline
$GrpType = Read-Host

switch ($GrpType)
{
	1
	{
		write-host ""
        $orggroup = "UL COM "
        $org = read-host "Is this for the UL.COM Organization (Y/N)? "
        If ($org -eq "N")
        {
            $orggroup = "UL ORG "
        }
        write-host ""
		write-host "Enter the Location of the DST.All Dynamic Distribution List to remove? " -ForegroundColor Yellow -NoNewline
		$LocName = Read-Host
		$EmpGrpName = "DST.All " + $orggroup + $LocName + " Employees Only"
		$StaffGrpName = "DST.All " + $orggroup + $LocName + " Staff"
	}
	2
	{
		write-host ""
		write-host "Enter the FirstName of the New DST.SUP Dynamic Distribution List to remove? " -ForegroundColor Yellow -NoNewline	
		$FirstName = Read-Host
		write-host "Enter the LastName of the New DST.SUP Dynamic Distribution List to remove? " -ForegroundColor Yellow -NoNewline	
		$LastName = Read-Host	
		$EmpGrpName = "DST.SUP " + $LastName + " " + $FirstName + " Employees Only"
		$StaffGrpName = "DST.SUP " + $LastName + " " + $FirstName + " Staff"
	}
	3
	{
		write-host ""
		write-host "Enter the Location of the DSG.AO2 List to remove? " -ForegroundColor Yellow -NoNewline	
		$LocName = Read-Host
		$EmpGrpName = "DSG.AO2." + $LocName + " Employees Only"
		$StaffGrpName = "DSG.AO2." + $LocName + " Staff"
	}
	4
	{
		write-host ""
		write-host "Enter the Location of the DSG.AO4 List to remove? " -ForegroundColor Yellow -NoNewline	
		$LocName = Read-Host
		$EmpGrpName = "DSG.AO4." + $LocName + " Employees Only"
		$StaffGrpName = "DSG.AO4." + $LocName + " Staff"
	}
}

write-host
write-host "Name of the group for Employees Only - " $EmpGrpName
write-host "     Name of the group for all Staff - " $StaffGrpName

If ($GrpType -le 2)
{
    $colu = Get-DynamicDistributionGroup $EmpGrpName
    $DLMem = (Get-Recipient -RecipientPreviewFilter $colu.LdapRecipientFilter -ResultSize Unlimited).count
}
else
{
    $colu = Get-AzureADGroup -SearchString $EmpGrpName
    If ($colu -ne $null)
    {
        $DLMem = (Get-AzureADGroupMember -ALL 1 -ObjectId $colu.ObjectId).count
    }
}

If (($DLMem -eq 0) -and ($colu -ne $null))
{
    $WhoAmI	= WhoAmI
    write-host "`n     Removing group " $EmpGrpName
    write-host "Report file can be found at: " $OutFileName
    $OutFileName = "e:\Automation\RemoveDynamicGroup\Report\Report-" + ($EmpGrpName -replace "[ /]","") + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
    $LineToWrite = "STAR" + "`t" + "RemoveDynDistGroup script has started"
    Write-Output $LineToWrite >> $OutFileName
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI
    Write-Output $LineToWrite >> $OutFileName
    Write-Output "" >> $OutFileName
    start-transcript
    Write-Host
    Write-Host "     Number of members in " $EmpGrpName " group " -ForegroundColor Yellow -NoNewline
    Write-Host $DLMem -ForegroundColor Yellow
    write-host ""
    Write-Output "Dynamic Distribution Group Information" >> $OutFileName

    If ($GrpType -le 2)
    {
        Get-DynamicDistributionGroup $EmpGrpName > $OutFileName
   
        Write-Output "Dynamic Distribution Group Detailed Information" >> $OutFileName

        Get-DynamicDistributionGroup $EmpGrpName |fl >> $OutFileName

        Write-Output "Dynamic Distribtuion Group Membership Rule" >> $OutFileName
        
        Get-DynamicDistributionGroup $EmpGrpName |fl RecipientFilter,IncludedRecipients >>$OutFileName

        Remove-DynamicDistributionGroup $EmpGrpName -Confirm:$False
    }
    else
    {
        Get-AzureADGroup -SearchString $EmpGrpName > $OutFileName
   
        Write-Output "AzureAD Dynamic Group Detailed Information" >> $OutFileName

        Get-AzureADGroup -SearchString $EmpGrpName |fl >> $OutFileName

        Write-Output "AzureAD Dynamic Group Membership Rule" >> $OutFileName
        
        Get-AzureADGroup -SearchString $EmpGrpName |fl MembershipRule >>$OutFileName

        Remove-AzureADGroup -ObjectID $colu.ObjectId
    }
}
else
{
    write-host "This group does not exist or there are active members in this group and therefore it will not be removed" - foregroundcolor Red
}

If ($GrpType -le 2)
{
    $colu = Get-DynamicDistributionGroup $StaffGrpName
    $DLMem = (Get-Recipient -RecipientPreviewFilter $colu.LdapRecipientFilter -ResultSize Unlimited).count
}
else
{
    $colu = Get-AzureADGroup -SearchString $StaffGrpName
    $DLMem = (Get-AzureADGroupMember -ALL 1 -ObjectId $colu.ObjectId).count
}

If (($DLMem -eq 0) -and ($colu -ne $null))
{
    write-host "     Removing group " $StaffGrpName
    $OutFileName = "e:\Automation\RemoveDynamicGroup\Report\Report-" + ($StaffGrpName -replace "[ /]","") + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
    write-host "Report file can be found at: " $OutFileName
    Write-Host
    Write-Host "     Number of members in" $StaffGrpName "group " -ForegroundColor Yellow -NoNewline
    Write-Host $DLMem -ForegroundColor Yellow
    write-host ""
    Write-Output "Dynamic Distribution Group Information" >> $OutFileName

    If ($GrpType -lt 3)
    {
        Get-DynamicDistributionGroup $StaffGrpName > $OutFileName
        
        Write-Output "Dynamic Distribution Group Detailed Information" >> $OutFileName

        Get-DynamicDistributionGroup $StaffGrpName |fl >> $OutFileName

        Write-Output "Dynamic Distribtuion Group Membership Rule" >> $OutFileName
        
        Get-DynamicDistributionGroup $StaffGrpName |fl RecipientFilter,IncludedRecipients >>$OutFileName

        Remove-DynamicDistributionGroup $StaffGrpName -Confirm:$False
    }
    else
    {
        Get-AzureADGroup -SearchString $StaffGrpName > $OutFileName
   
        Write-Output "AzureAD Dynamic Group Detailed Information" >> $OutFileName

        Get-AzureADGroup -SearchString $StaffGrpName |fl >> $OutFileName

        Write-Output "AzureAD Dynamic Group Membership Rule" >> $OutFileName
        
        Get-AzureADGroup -SearchString $StaffGrpName |fl MembershipRule >>$OutFileName

        Remove-AzureADGroup -ObjectID $colu.ObjectId
    }
}
else
{
    write-host "This group does not exist or there are active members in this group and therefore it will not be removed" - foregroundcolor Red
}
stop-transcript
