<####  This script create new UL Employee Only and UL Staff Dynamic Distribution Lists
#
# Called by:  DistributionSecurityGroupMenu.ps1
#
# 06/17/2016 - SAG - Added code to check for Displaynames longer than 64 chara and truncate them
#                  - Removed the MailUsers configuraiton as these individuals are mail enabled but not licensed
#				   - Added code to allow for creation of DST.SUP groups
# 09/28/2016 - SAG - Removed extra quotes around the executive name in the DST.SUP group code.
# 08/31/2017 - SAG - Added code to allow the configuring the Supervisor2/Supervisor3 attributes for the DST.SUP groups 
# 02/09/2018 - SAG - Added code to include MailTips
# 04/17/2021 - SAG - Modified creation for the new Alpha Org Supervisor3 now Supervisor2 and Supervisor2 now Supervisor1 also removed configuring the MailContacts on these groups
# 05/27/2021 - SAG - Modified to inlcude creation of location and supervisor people leader groups
# 11/19/2021 - SAG - Modified to add that senderauthentication is required when creating new groups
# 01/04/2022 - SAG - Modified to include creation of AO2 groups
#>

Function Create-DSTGroups
{
    write-host

    $EmpGrpNameTrunc = $EmpGrpName

    if ($EmpGrpName.Length -gt 64)
    {
    #	The group Name cannot be longer than the 64 charactrer
	    $EmpGrpNameTrunc = $EmpGrpName.substring(0,64)
    }

    If (($GrpType -eq 1) -or ($GrpType -eq 2))
    {
        write-host "Name of the group for UL Employees Only - " $EmpGrpName
        $StaffGrpNameTrunc = $StaffGrpName
        if ($StaffGrpName.Length -gt 64)
        {
        #	The group Name cannot be longer than the 64 charactrer
	        $StaffGrpNameTrunc = $StaffGrpName.substring(0,64)
        }
        write-host "     Name of the group for all UL Staff - " $StaffGrpName
    }
    If (($GrpType -eq 3) -or ($GrpType -eq 4))
    {
        write-host "Name of the group for People Leaders - " $EmpGrpName
    }


    write-host "Continue creating the groups (Y/N)? " -ForegroundColor Red -NoNewline
    $Cont = Read-Host

    If ($Cont -eq "Y")
    {
        $EmpGrpAlias = $EmpGrpName -Replace '[ (),-]',''
        $EmpGrpAlias = $EmpGrpAlias.Replace("UL","")
        $EmpGrpAlias = $EmpGrpAlias.Replace("loyees","")

        $INetAlias = $EmpGrpAlias + "@ul.com"
        If ($SupVal -eq "N")
        {
        # Individual is a Supervisor2
            Switch ($GrpType)
            {
                1
                {
                    New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -PrimarySmtpAddress $INetAlias
                }
                2
                {
                    New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute7 $SupENo -PrimarySmtpAddress $INetAlias
                }
                3
                {
                    New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $INetAlias
                }
                4
                {
                    New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute7 $SupENo -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $INetAlias
                }
            }
            $SupType = "Supervisor2"
        }
        else
        {
        # Individual is a Supervisor1
            switch($GrpType)
            {
                1
                {
                    New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -PrimarySmtpAddress $INetAlias
                }
                2
                {
                    New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute6 $SupENo -PrimarySmtpAddress $INetAlias
                }
                3
                {
                    New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute6 $SupENo -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $INetAlias
                }
                4
                {
                    New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -ConditionalCustomAttribute6 $SupENo -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $INetAlias
                }
            }
            $SupType = "Supervisor1"
        }

        switch ($GrpType)
	    {
		    1
		    {
			    $EmpGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff that have an employee type of ""Employee"" and are in the ""$LocName"" location.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                $MailTip = "Sends to individuals listed as Employee assigned to the '" + $LocName + "' site in Oracle."
		    }
		    2
		    {
			    $EmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " + $FirstName + " " + $LastName + " as " + $SupType + " in Oracle.  This group contains staff that have an employee type of ""Employee"".  Manual modification of this group is not possible."
                $MailTip = "Sends to individuals listed as Employee that report to '" + $LastName + " " + $FirstName + "' as" + $SupType + " in Oracle."
		    }
            3
            {
			    $EmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals location in the ""$LocName"" location and is a People Leader in Oracle.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                $MailTip = "Sends to individuals listed as Employee assigned to the '" + $LocName + "' site in Oracle."
            }
            4
            {
			    $EmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " + $FirstName + " " + $LastName + " as " + $SupType + " and are People Leaders in Oracle.  Manual modification of this group is not possible."
                $MailTip = "Sends to individuals listed as Employee that report to '" + $LastName + " " + $FirstName + "' as" + $SupType + " in Oracle."
            }
	    }
    #    $INetAlias = $EmpGrpAlias + "@ul.com"

        Set-DynamicDistributionGroup $EmpGrpName -Notes $EmpGrpNote -MailTip $MailTip -RequireSenderAuthenticationEnabled $True

        If ($GrpType -ne 4)
        {
            $StaffGrpAlias = $StaffGrpName -Replace '[ (),-]',''
            $StaffGrpAlias = $StaffGrpAlias.Replace("UL","")
            $INetAlias = $StaffGrpAlias + "@ul.com"
	        switch ($GrpType)
	        {
		        1
		        {
			        $StaffGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff that have an employee type of ""Employee, Consultant, Contractor, Temporary Worker or Freelance Worker"" and are in the ""$LocName"" location.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                    $StaffTip = "Sends to all Employee, Contractor, Consultant, Freelance Contractor or Contingent Worker assigned to the '" + $LocName + "' site in Oracle."
		        }
		        2
		        {
			        $StaffGrpNote = "The membership of this group is determined at the time the group and contains individuals that report up to " + $FirstName + " " + $LastName + " as " + $SupType + " in Oracle.  This group contains staff that have an employee type of ""Employee, Consultant, Contractor or Temporary Worker"".  Manual modification of this group is not possible."
                    $StaffTip = "Sends to all Employee, Contractor, Consultant, Freelance Contractor or Contingent Worker that report to '"  + $FirstName + " " + $LastName + "' as" + $SupType + " in Oracle."
		        }
	        }
        #    $INetAlias = $StaffGrpAlias + "@ul.com"
    
            If ($SupVal -eq "N")
            {
            # Individual is a Supervisor2
                New-DynamicDistributionGroup -DisplayName $StaffGrpName -Alias $StaffGrpAlias -Name $StaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -ConditionalCustomAttribute7 $SupENo -PrimarySmtpAddress $INetAlias
            }
            else
            {
            # Individual is a Supervisor1
                New-DynamicDistributionGroup -DisplayName $StaffGrpName -Alias $StaffGrpAlias -Name $StaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -ConditionalCustomAttribute6 $SupENo -PrimarySmtpAddress $INetAlias
            }

            Set-DynamicDistributionGroup $StaffGrpName -Notes $StaffGrpNote -MailTip $StaffTip -RequireSenderAuthenticationEnabled $True
        }
    }
    else
    {
        Write-Host "No Groups Created"
        pause
    }
}

Function Create-DSGAO4Groups
{
    $AO4Groups = Get-AzureADMSGroup -SearchString "DSG.AO4"
    $cnt = 0

	$StaffGrpName = "DSG.AO4." + $LocName + " UL Staff"

    $bute4 = $Location
    $AO4 = $Location -replace (",","")
    $EmpGrpName = "DSG.AO4." + $LocName + " UL Employees Only"
    $EmpGrpNick =  "DSG.AO4." + $LocName + "EmpOnly" -replace("[ ]","")
    $EmpGrpDesc = "All UL Employees in Alpha Org 4 " +  $LocName
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

    $StaffGrpName = "DSG.AO4." + $LocName + " UL Staff"
    $StaffGrpNick = "DSG.AO4." + $LocName + "Staff" -replace("[ ]","")
    $StaffGrpDesc = "All UL Staff in Alpha Org 4 " +  $LocName
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

Function Create-DSGAO2Groups
{
    $AO2Groups = Get-AzureADMSGroup -SearchString "DSG.AO2"
    $cnt = 0

	$StaffGrpName = "DSG.AO2." + $LocName + " UL Staff"

    $bute8 = $Location
    $AO2 = $Location -replace (",","")
    $EmpGrpName = "DSG.AO2." + $LocName + " UL Employees Only"
    $EmpGrpNick =  "DSG.AO2." + $LocName + "EmpOnly" -replace("[ ]","")
    $EmpGrpDesc = "All UL Employees in Alpha Org 4 " +  $LocName
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

    $StaffGrpName = "DSG.AO2." + $LocName + " UL Staff"
    $StaffGrpNick = "DSG.AO2." + $LocName + "Staff" -replace("[ ]","")
    $StaffGrpDesc = "All UL Staff in Alpha Org 4 " +  $LocName
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

write-host "     Enter ( 1) for DST.All groups"
write-host "           ( 2) for DST.SUP groups"
write-host "           ( 3) for DST.All People Leader group"
write-host "           ( 4) for DST.SUP People Leader group"
write-host "           ( 5) for DSG.AO4 groups"
write-host "           ( 6) for DSG.AO2 groups"
write-host ""
write-Host "           ( 0) to Return to the Distribution Group Admin Menu"
write-Host "     Enter Option? " -ForegroundColor Red -NoNewline
$GrpType = Read-Host

switch ($GrpType)
{
	1
	{
		write-host ""
		write-host "Enter the Location of the New DST.All Dynamic Distribution List? " -ForegroundColor Yellow -NoNewline
		$LocName = Read-Host
		$EmpGrpName = "DST.All " + $LocName + " UL Employees Only"
		$StaffGrpName = "DST.All " + $LocName + " UL Staff"
  	}
	2
	{
		write-host ""
		write-host "Enter the FirstName of the New DST.SUP Dynamic Distribution List? " -ForegroundColor Yellow -NoNewline	
		$FirstName = Read-Host
		write-host "Enter the LastName of the New DST.SUP Dynamic Distribution List? " -ForegroundColor Yellow -NoNewline	
		$LastName = Read-Host	
		$EmpGrpName = "DST.SUP " + $LastName + " " + $FirstName + " UL Employees Only"
		$StaffGrpName = "DST.SUP " + $LastName + " " + $FirstName + " UL Staff"
        write-host "Enter the Employee Number for this individual " -ForegroundColor Yellow -NoNewline
        $SupENo = Read-Host
        write-host "Is this individuals listed in the Supervisor1 field in the OED_Extract6.csv file (Y/N)? " -ForegroundColor Yellow -NoNewline
        $SupVal = Read-Host
	}
	3
	{
		write-host ""
		write-host "Enter the Location of the New DST.All People Leader Dynamic Distribution List? " -ForegroundColor Yellow -NoNewline
		$LocName = Read-Host
		$EmpGrpName = "DST.All " + $LocName + " People Leaders"
  	}
	4
	{
		write-host ""
		write-host "Enter the FirstName of the New DST.SUP People Leader Dynamic Distribution List? " -ForegroundColor Yellow -NoNewline	
		$FirstName = Read-Host
		write-host "Enter the LastName of the New DST.SUP People Leader Dynamic Distribution List? " -ForegroundColor Yellow -NoNewline	
		$LastName = Read-Host	
		$EmpGrpName = "DST.SUP " + $LastName + " " + $FirstName + " People Leaders"
        write-host "Enter the Employee Number for this individual " -ForegroundColor Yellow -NoNewline
        $SupENo = Read-Host
        write-host "Is this individuals listed in the Supervisor1 field in the OED_Extract6.csv file (Y/N)? " -ForegroundColor Yellow -NoNewline
        $SupVal = Read-Host
	}
	5
	{
		write-host ""
		write-host "Enter the Alpha HR Organization Level 4 of the New DSG.AO4 Dynamic Group? " -ForegroundColor Yellow -NoNewline
		$LocName = Read-Host
  	}

	6
	{
		write-host ""
		write-host "Enter the Alpha HR Organization Level 2 of the New DSG.AO2 Dynamic Group? " -ForegroundColor Yellow -NoNewline
		$LocName = Read-Host
  	}
}

If ($GrpType -ne 5)
{
    Create-DSTGroups
}
else
{
    Create-DSGAO4Groups
}

<#
write-host

$EmpGrpNameTrunc = $EmpGrpName

if ($EmpGrpName.Length -gt 64)
{
#	The group Name cannot be longer than the 64 charactrer
	$EmpGrpNameTrunc = $EmpGrpName.substring(0,64)
}

If (($GrpType -eq 1) -or ($GrpType -eq 2))
{
    write-host "Name of the group for UL Employees Only - " $EmpGrpName
    $StaffGrpNameTrunc = $StaffGrpName
    if ($StaffGrpName.Length -gt 64)
    {
    #	The group Name cannot be longer than the 64 charactrer
	    $StaffGrpNameTrunc = $StaffGrpName.substring(0,64)
    }
    write-host "     Name of the group for all UL Staff - " $StaffGrpName
}
If (($GrpType -eq 3) -or ($GrpType -eq 4))
{
    write-host "Name of the group for People Leaders - " $EmpGrpName
}


write-host "Continue creating the groups (Y/N)? " -ForegroundColor Red -NoNewline
$Cont = Read-Host

If ($Cont -eq "Y")
{
    $EmpGrpAlias = $EmpGrpName -Replace '[ (),-]',''
    $EmpGrpAlias = $EmpGrpAlias.Replace("UL","")
    $EmpGrpAlias = $EmpGrpAlias.Replace("loyees","")

    $INetAlias = $EmpGrpAlias + "@ul.com"
    If ($SupVal -eq "N")
    {
    # Individual is a Supervisor2
        Switch ($GrpType)
        {
            1
            {
                New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -PrimarySmtpAddress $INetAlias
            }
            2
            {
                New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute7 $SupENo -PrimarySmtpAddress $INetAlias
            }
            3
            {
                New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $INetAlias
            }
            4
            {
                New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute7 $SupENo -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $INetAlias
            }
        }
        $SupType = "Supervisor2"
    }
    else
    {
    # Individual is a Supervisor1
        switch($GrpType)
        {
            1
            {
                New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -PrimarySmtpAddress $INetAlias
            }
            2
            {
                New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute6 $SupENo -PrimarySmtpAddress $INetAlias
            }
            3
            {
                New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute1 "Employee" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute6 $SupENo -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $INetAlias
            }
            4
            {
                New-DynamicDistributionGroup -DisplayName $EmpGrpName -Alias $EmpGrpAlias -Name $EmpGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -ConditionalCustomAttribute6 $SupENo -ConditionalCustomAttribute9 "Y" -PrimarySmtpAddress $INetAlias
            }
        }
        $SupType = "Supervisor1"
    }

    switch ($GrpType)
	{
		1
		{
			$EmpGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff that have an employee type of ""Employee"" and are in the ""$LocName"" location.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
            $MailTip = "Sends to individuals listed as Employee assigned to the '" + $LocName + "' site in Oracle."
		}
		2
		{
			$EmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " + $FirstName + " " + $LastName + " as " + $SupType + " in Oracle.  This group contains staff that have an employee type of ""Employee"".  Manual modification of this group is not possible."
            $MailTip = "Sends to individuals listed as Employee that report to '" + $LastName + " " + $FirstName + "' as" + $SupType + " in Oracle."
		}
        3
        {
			$EmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals location in the ""$LocName"" location and is a People Leader in Oracle.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
            $MailTip = "Sends to individuals listed as Employee assigned to the '" + $LocName + "' site in Oracle."
        }
        4
        {
			$EmpGrpNote = "The membership of this group is determined at the time the group is used and contains individuals that report up to " + $FirstName + " " + $LastName + " as " + $SupType + " and are People Leaders in Oracle.  Manual modification of this group is not possible."
            $MailTip = "Sends to individuals listed as Employee that report to '" + $LastName + " " + $FirstName + "' as" + $SupType + " in Oracle."
        }
	}
#    $INetAlias = $EmpGrpAlias + "@ul.com"

    Set-DynamicDistributionGroup $EmpGrpName -Notes $EmpGrpNote -MailTip $MailTip

    If ($GrpType -ne 4)
    {
        $StaffGrpAlias = $StaffGrpName -Replace '[ (),-]',''
        $StaffGrpAlias = $StaffGrpAlias.Replace("UL","")
        $INetAlias = $StaffGrpAlias + "@ul.com"
	    switch ($GrpType)
	    {
		    1
		    {
			    $StaffGrpNote = "The membership of this group is determined at the time the group is used and is based upon the location an individual is assigned to in Oracle.  This group contains staff that have an employee type of ""Employee, Consultant, Contractor, Temporary Worker or Freelance Worker"" and are in the ""$LocName"" location.  To check the membership use the Global Employee Directory.  Manual modification of this group is not possible."
                $StaffTip = "Sends to all Employee, Contractor, Consultant, Freelance Contractor or Contingent Worker assigned to the '" + $LocName + "' site in Oracle."
		    }
		    2
		    {
			    $StaffGrpNote = "The membership of this group is determined at the time the group and contains individuals that report up to " + $FirstName + " " + $LastName + " as " + $SupType + " in Oracle.  This group contains staff that have an employee type of ""Employee, Consultant, Contractor or Temporary Worker"".  Manual modification of this group is not possible."
                $StaffTip = "Sends to all Employee, Contractor, Consultant, Freelance Contractor or Contingent Worker that report to '"  + $FirstName + " " + $LastName + "' as" + $SupType + " in Oracle."
		    }
	    }
    #    $INetAlias = $StaffGrpAlias + "@ul.com"
    
        If ($SupVal -eq "N")
        {
        # Individual is a Supervisor2
            New-DynamicDistributionGroup -DisplayName $StaffGrpName -Alias $StaffGrpAlias -Name $StaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -ConditionalCustomAttribute7 $SupENo -PrimarySmtpAddress $INetAlias
        }
        else
        {
        # Individual is a Supervisor1
            New-DynamicDistributionGroup -DisplayName $StaffGrpName -Alias $StaffGrpAlias -Name $StaffGrpNameTrunc -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute2 "A" -ConditionalCustomAttribute3 $LocName -ConditionalCustomAttribute6 $SupENo -PrimarySmtpAddress $INetAlias
        }

        Set-DynamicDistributionGroup $StaffGrpName -Notes $StaffGrpNote -MailTip $StaffTip
    }
}
else
{
    Write-Host "No Groups Created"
    pause
}
#>