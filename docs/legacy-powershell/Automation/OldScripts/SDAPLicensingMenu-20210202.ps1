<#
#                       Licensing Menu
#  Called from SDAP Admin Menu to perform Licensing Activites
#
# 06/21/2017 - SAG - Created from O365 Admin Team Licensing Menu
# 11/26/2017 - SAG - Added the Microsoft Relationship Sales Solution licencesigncr
#     also change numbering to provide gaps for adding of new license types
# 01/08/2018 - SAG - Added to the MRSS License to remove CRM License if Assigned
#     and to output notes if there are no MRSS licenses available
# 03/22/2018 - SAG - Added code when assigning a CRM license if there is an MRSS 
#     license assigned that will be removed.  Also removed the O365License function
#     and changed this to call the SDAPO365Licenses.ps1 script.
# 04/21/2018 - SAG - Modified code when assigning CRM licenses to disable most sublicensesb
# 06/26/2018 - SAG - Added option for assigning ATP license
# 01/14/2019 - SAG - Added code to remove E3 license if E4 is already assigned.  Also added checks
#     to make sure the other standard PowerBI, EMS and ATP licenses were assignedf
# 04/16/2019 - SAG - Modified the CRM License assignements to assign the new license type
# 04/19/2019 - SAG - Removed the E4 Options and updated to include new license features with the E3
# 06/14/2019 - SAG - Removed the ATP license assignment options and added Windows Defender/Teams Commercial and Flow P2
# 07/05/2019 - SAG - Added code to not allow overassignment of CRM or MRSS licenses
# 08/21/2019 - SAG - Added Phone System and Audio Conferencing features
# 02/06/2020 - SAG - Added Advanced Threat Protection Plan1 to the default license assignment
# 09/15/2020 - SAG - Fixed the E3 Enabled features code
# 11/18/2020 - SAG - Moved the E3 license assignment after the check/removal of the P2 license
# 12/06/2020 - SAG - Added Project Plan 5 license
#>
$WhoAmI			= WhoAmI

$LogDirectory   = "E:\SDAP\Licensing\Log"
$LogFile		= $LogDirectory + "\" + "Log-Licensing.log"
	
# Retrieve the current Date and Time for use in log files
$uDate = get-date -uformat %D
$uTime = get-date -uformat %T
$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"


Invoke-Expression -command E:\O365AdminShared\Scripts\O365DisabledLicenseFeatures.ps1

Function RetentPolicy
{
    $MbxCreated = [bool](get-mailbox $EmpNo -ErrorAction SilentlyContinue)

    If ($MbxCreated -eq "True")
    {
        If ($ADUser.extensionattribute4 -eq "IT")
        {
            Set-MailBox $EmpNo -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 3 yr Delete"
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Updating Retention Policy to IT Policy for " + $EmpNo
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        }
        else
        {
            Set-MailBox $EmpNo -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 365 Day Delete"
            Set-Mailbox $EmpNo –RetentionHoldEnabled $true –StartDateForRetentionHold 04/01/2011
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Updating Retention Policy to UL Default Policy for " + $EmpNo
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        }
	    Set-CASMailBox $EmpNo -ImapEnabled $false -PopEnabled $false
    }
    else
    {
        write-host "Mailbox does not exist or is in the process of being created.  If the license was just assigned"
        write-host "please allow 3-5 minutes for the mailbox to be created and the select Option 10 from the menu"
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Mailbox Does Not Exist " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
}

Function EnabledFeature
{
        write-host "`n`tEnabled SubLicense Options:" -ForegroundColor Magenta

    If ($HasE3 -eq "True")
    {
        $SSKID = "UL:ENTERPRISEPACK"
        $LicType = "Enterprise E3"
    }

    $text = "`t{0}`t`t{1}`t`t`t{2}" -f "DefaultAssignment","Status ","SubLicenseName"
    $text
    $text = "`t{0}`t`t{1}`t`t`t{2}" -f "-----------------","------ ","--------------"
    $text

    $LineToWrite = $RecordEvent + "UPDA" + "`t" + $LicType + " SubLicenses Enabled/Status for " + $Global:UPN
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    $SubLic = ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq $SSKID}).ServiceStatus
	foreach ($SubLic in $SubLic)
	{
        $AssignWho = "Enabled All Types"

        switch ($subLic.ServicePlan.ServiceName)
        {
            "POWER_VIRTUAL_AGENTS_O365_P2"
            {
                $LicName = "Power Virtual Agents"
                $AssignWho = "Disabled by Default"
            }
            "CDS_O365_P2"
            {
                $LicName = "Common Data Service P2"
                $AssignWho = "Disabled by Default"
            }
            "PROJECT_O365_P2"
            {
                $LicName = "O365 Project P2"
            }
            "DYN365_CDS_O365_P2"
            {
                $LicName = "Common Data Service"
                $AssignWho = "Disabled by Default"
            }
            "MICROSOFTBOOKINGS"
            {
                $LicName = "Microsoft Bookings"
            }
            "KAIZALA_O365_P3"
            {
                $LicName = "Microsoft Kaizala Pro"
            }
            "MICROSOFT_SEARCH"
            {
                $LicName = "Microsoft Search"
            }
            "WHITEBOARD_PLAN2"
            {
                $LicName = "Whiteboard (Plan 2)"
            }
            "MIP_S_CLP1"
            {
                $LicName = "Information Protection for Office 365 - Standard"
            }
            "MYANALYTICS_P2"
            {
                $LicName = "Insights by MyAnalytics"
            }
            "BPOS_S_TODO_2"
            {
                $LicName = "To-Do (Plan 2)"
            }
            "FORMS_PLAN_E3"
            {
                $LicName = "Microsoft Forms (Plan E3)"
            }
            "STREAM_O365_E3"
            {
                $LicName = "Microsoft Stream for O365 E3 SKU"
            }
            "Deskless"
            {
                $LicName = "Microsoft StaffHub"
                $AssignWho = "Disabled by Default"
            }
            "FLOW_O365_P2"
            {
                $LicName = "Flow for Office 365 "
            }
            "POWERAPPS_O365_P2"
            {
                $LicName = "PowerApps for Office 365"
            }
            "TEAMS1"
            {
                $LicName = "Microsoft Teams"
            }
            "PROJECTWORKMANAGEMENT"
            {
                $LicName = "Microsoft Planner"
            }
            "SWAY"
            {
                $LicName = "Sway"
            }
            "INTUNE_O365"
            {
                $LicName = "Intune (Uses EMS License)"
           }
            "YAMMER_ENTERPRISE"
            {
                $LicName = "Yammer Enterprise"
                $AssignWho = "Enabled Empl Only"
            }
            "RMS_S_ENTERPRISE"
            {
                $LicName = "Azure Rights Management"
            }
            "OFFICESUBSCRIPTION"
            {
                $LicName = "Office 365 ProPlus"
            }
            "MCOSTANDARD"
            {
                $LicName = "Skype for Business Online (Plan 2)"
                $AssignWho = "Enabled Empl Only"
            }
            "SHAREPOINTWAC"
            {
                $LicName = "Office Online"
                $AssignWho = "Enabled Empl Only"
            }
            "SHAREPOINTeNTERPRISE"
            {
                $LicName = "SharePoint Online (Plan 2)"
                $AssignWho = "Enabled Empl Only"
            }
            "EXCHANGE_S_ENTERPRISE"
            {
                $LicName = "Exchange Online (Plan 2)"
            }
        }

        $LicStat = $SubLic.ProvisioningStatus
        If ($LicStat -like "*Pending*")
        {
            $LicStat = "Pending "
        }
        If ($LicStat -eq "Success")
        {
            $LicStat = "Enabled "
        }
        $text = "`t{0}`t`t{1}`t`t{2}" -f $AssignWho,$LicStat,$LicName
        $text
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + $text
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}
}

Function AssignMRSS
{
	Write-host ""
    Set-MsolUser -UserPrincipalName $EmpNo -UsageLocation US -ErrorAction Silentlycontinue
    $sskid = "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
    $LicType = "Microsoft Relationship Sales Solution"
    $DisPlan = $Global:DisMRSSPlan
    	
	$LicDet = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq $SSKID}
			
	if ($LicDet.ActiveUnits -eq $LicDet.ConsumedUnits)
	{					
        write-host "`nNo " $LicType "Licenses Available to Assign"
        $LineToWrite = $RecordEvent + "DELE" + "`t" + "No Microsoft Relationship Sales Solution Licenses Available for Assignment to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
	else
	{
        If ($HasCRM -eq "True")
        {
            write-host "Removing CRM License from this account"
            Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:DYN365_ENTERPRISE_PLAN1"
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Removing Dynamics 365 Customer Engagement Plan License from account for " + $EmpNo
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        }

		write-host "`nAssigning " $EmpNo "a " $LicType "License" -ForegroundColor Yellow
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Assigning " + $LicType + "to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

        $DLO = ($DisPlan.Split(“,”))
        $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense $SSKID –LicenseOptions $MyO365Sku
    }

}

Function AssignCRM
{
	Write-host ""
    Set-MsolUser -UserPrincipalName $EmpNo -UsageLocation US -ErrorAction Silentlycontinue
    $sskid = "ul:DYN365_ENTERPRISE_PLAN1"
    $LicType = "Dynamics 365 Customer Engagement Plan"
    If($ADUser.ExtensionAttribute1 -like "Employee*")
    {
        $DisPlan = $Global:DisCRMPlanEmp
        $UsrType = "Employee"
    }
    else
    {
        $DisPlan = $Global:DisCRMPlanNonEmp
        $UsrType = "Non-Employee"
    }

	$LicDet = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq $SSKID}
			
	if ($LicDet.ActiveUnits -eq $LicDet.ConsumedUnits)
	{					
        write-host "`nNo " $LicType "Licenses Available to Assign"
        $LineToWrite = $RecordEvent + "DELE" + "`t" + "No Dynamics 365 Customer Engagement Plan Licenses Available for Assignment to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
	else
	{

        If ($HasMRSS -eq "True")
        {
            write-host "Removing Microsoft Relationship Sales Solution License from this account"
            Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
        }
		write-host "`nAssigning " $EmpNo "a " $LicType "License" -ForegroundColor Yellow
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Assigning " + $LicType + "with " + $UsrType + "Features to " + $EmpNo
        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
   	}

    $DLO = ($DisPlan.Split(“,”))
    $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense $SSKID –LicenseOptions $MyO365Sku

}

Function O365Licenses
{
    $O365Lic = (Get-MsolUser -UserPrincipalName $EmpNo).Licenses
	Write-Host "`nCurrent License Assignment for " $EmpNo " " $Global:UserLicense.DisplayName "-" $ADUser.ExtensionAttribute1-ForegroundColor Green
    
    If ($O365Lic.count -gt 0)
    {
        $LineToWrite = $RecordEvent + "DISP" + "`t" + "Reviewing O365 Licenses Assigned to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

        foreach ($O365Lic in $O365Lic)
        {
            switch ($O365Lic.AccountSkuID)
            {
                "ul:ENTERPRISEPACK"
                {
                    $text = "`tEnterprise E3"
                }
                "ul:EXCHANGEENTERPRISE"
                {
                    $text = "`tExchange Online Plan 2"
                }
                "ul:POWER_BI_STANDARD"
                {
                    $text = "`tPowerBI (Free)"
                }
                "ul:POWER_BI_PRO"
                {
                    $text = "`tPowerBI Pro" 
                }
                "ul:DYN365_ENTERPRISE_PLAN1"
                {
                    $text = "`tDynamics 365 Customer Engagement Plan"
                }
                "ul:EMS"
                {
                    $text = "`tEnterprise Mobility Suite/Intune (EMS)"
                }
                "ul:POWER_BI_INDIVIDUAL_USER"
                {
                    $text = "`tPower BI for O365"
                }
                "ul:MCOIMP"
                {
                   $text = "`tSkype for Business Online (Plan 1)"
                }
                "ul:ECAL_SERVICES"
                {
                    $text = "`tECAL Services (EOA, EOP, DLP)"
                }
                "ul:SHAREPOINTSTANDARD"
                {
                    $text = "`tSharePoint Online (Plan 1)"
                }
                "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
                {
                    $text = "`tMicrosoft Relationship Sales Solution"
                }
				"ul:ATP_ENTERPRISE"
				{
					$text = "`tAdvanced Threat Protection Plan1"
				}
                "ul:FLOW_P2"
                {
                    $text = "`tMicrosoft Flow (Plan 2)"
                }
                "ul:TEAMS_COMMERCIAL_TRIAL"
                {
                    $text = "Microsoft Teams Commercial Cloud (User Initiated)"
                }
                "ul:WIN_DEF_ATP"
                {
                    $text ="Defender Advanced Threat Protection"
                }
                "ul:MEETING_ROOM"
                {
                    $text = "`tMeeting Room"
                }
                "ul:POWERFLOW_P2"
                {
                    $text = "`tPowerApps Plan 2"
                }
                "ul:MCOEV"
                {
                    $text = "`tPhone System"
                }
                "ul:MCOMEETADV"
                {
                    $text = "`tAudio Conferencing"
                }
                "ul:PROJECTPREMIUM"
                {
                    $text = "`tProject Plan 5"
                }
            }
            Write-Host $text
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + $text
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        }
    }
    else
    {
        Write-Host "`n`tNo licenses assigned to this account" -ForegroundColor Red
        $LineToWrite = $RecordEvent + "DISP" + "`t" + "No O365 Licenses Assigned to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

    }
}

Function StandardLicenses
{
    If ($HasEMS -eq $False)
	{
	    Write-Host "`nAssigning Enterprise Mobility Suite/Intune (EMS) License to" $EmpNo -ForegroundColor Yellow
		Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:EMS"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Enterprise Mobility Suite/Intune (EMS) License to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}  

    If (($HasBIFree -eq $False) -and ($HasBIPro -eq $False))
    {
	    Write-Host "`nAssigning PowerBI Free License to" $EmpNo -ForegroundColor Yellow
	    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:POWER_BI_STANDARD"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning PowerBI Free License to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}

    If ($HasATPP1 -eq $False)
    {
	    Write-Host "`nAssigning Advanced Threat Protection Plan1 License to" $EmpNo -ForegroundColor Yellow
	    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:ATP_ENTERPRISE"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Advanced Threat Protection Plan1 License to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}

    if ($HasPhone -eq $False)
    {
        Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicenses "ul:MCODEV" -ErrorAction SilentlyContinue
	    $LineToWrite = $RecordEvent + "INFO" + "`t" + "     Assigned Phone System License to " + $User.UPN
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "     Assigned Phone System License"
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite		
    }
}

Function AssignE3License
{
	Write-host ""
    $sskid = "ul:ENTERPRISEPACK"
    $LicType = "Enterprise E3"

    If ($ADUser.ExtensionAttribute1 -like "Employee*")
    {
        $DisPlan = $Global:DisPlanEmpE3
    }
    else
    {
        $DisPlan = $Global:DisPlanNonEmpE3
    }
	
	$LicDet = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}
			
    If ($HasE3 -eq $True)
    {
        write-host "`nEnterprise E3 License Already Assigned to " $LicType "License"
        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Enterprise E3 License Already Assigned to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
    else
    {
        If ($ADUser.ExtensionAttribute1 -like "Employee*")
	    {
		    write-host "`nAssigning " $EmpNo "an " $LicType "License with Employee Features Enabled" -ForegroundColor Yellow
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Assigning " + $LicType + "with Employee Features to " + $EmpNo
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	    }
	    else
	    {
		    write-host "`nAssigning " $EmpNo "an " $LicType "License with Non-Employee Features Enabled" -ForegroundColor Yellow
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Assigning " + $LicType + "with Non-Employee Features to " + $EmpNo
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	    }

        $DLO = ($DisPlan.Split(“,”))
        $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense $SSKID –LicenseOptions $MyO365Sku
        RetentPolicy
    }

    If ($HasExP2 -eq "True")
    {
        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:EXCHANGEENTERPRISE"
        write-host "`nRemoving Exchange Online P2 License and reassigning to a " $LicType "License"
        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing P2 License from " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
}
# Create logging folder $LogDirectory if it's not present
	if (Test-Path $LogDirectory)
	{
		# the directory is present
	}
	else
	{
		mkdir $LogDirectory
	}
		
# Add start record to log file
	$LineToWrite = "`n" + $RecordEvent + "STAR" + "`t" + "Licensing script has started"
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

write-host "`nEnter Employee Number of Individual to make License Changes for or 0 to Exit: " -ForegroundColor Yellow -NoNewline
$ENo = Read-Host
$LicChg = 1
$EmpNo = $ENo + "@global.ul.com"

Do
{
    If ($LicChg -ne "0")
    {
        $LineToWrite = "`n" + $RecordEvent + "INFO" + "`t" + "Performing Licensing Review for " + $EmpNo
        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		if ([bool](Get-MsolUser -UserPrincipalName $EmpNo -ErrorAction SilentlyContinue) -eq "True")
		{

			$Global:UserLicense = Get-MsolUser -UserPrincipalName $EmpNo

            $UserDetails = [bool](get-mailbox $EmpNo -ErrorAction SilentlyContinue)
            $HasBIFree = [bool] ($Global:UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:POWER_BI_STANDARD"})
            $HasBIPro = [bool] ($Global:UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:POWER_BI_PRO"})
            $HasCRM = [bool] ($Global:UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:DYN365_ENTERPRISE_PLAN1"})
            $HasEMS = [bool] ($Global:UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:EMS"})
            $HasExP2 = [bool] ($Global:UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:EXCHANGEENTERPRISE"})
            $HasE3 = [bool] ($Global:UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:ENTERPRISEPACK"})
            $HasMRSS = [bool] ($Global:UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"})
            $HasFlowP2 = [bool] ($Global:UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:FLOW_P2"})
            $HasATPP1 = [bool] ($Global:UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:ATP_ENTERPRISE"})
            $HasATP = [bool] ($Global:UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:WIN_DEF_ATP"})
            $HasPhone = [bool] ($Global:UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:MCODEV"})
			
			$LDAPFilter = "(userPrincipalName=" + $EmpNo + ")"
		    $ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,proxyaddresses,extensionattribute1,extensionattribute4
#            invoke-expression -Command .\SDAPO365Licenses.ps1
            O365Licenses

			write-host "`n`t`t`t`tLicensing Menu" -ForegroundColor Magenta
			write-host
			write-Host "     Enter ( 1) Assign Standard E3 and Standard License Set"
            write-host "           ( 2) Assign CRM/Dynamics License"
            write-host "           ( 3) Assign PowerBI (Free) License"
            write-host "           ( 4) Assign Enterprise Mobility (EMS) License"
            write-host "           ( 5) Assign Microsoft Relationship Sales solution"
            write-host "           ( 6) Assign Advanced Threat Protection Plan1 License"
            write-host "`n           (10) Remove CRM/Dynamics License"
            write-host "           (11) Remove PowerBI (Free) License"
            write-host "           (12) Remove Enterprise Mobility Suite/Intune (EMS) License"
            write-host "           (13) Remove Microsoft Relationship Sales solution"
            write-host "           (14) Remove Advanced Threat Protection Plan1 License" 
            write-host "`n           (20) Review Enabled E3 Sub-License Features"
            Write-Host "           (21) Reset Enabled E3 Sub-License Features to Employee/Non-Employee Standard"
 
			write-Host "`n           ( 0) to Return to the SDAP Admin Menu"
			write-Host "     Enter Option? " -ForegroundColor Red -NoNewline
			$LicChg = Read-Host
            if ($LicChg.length -eq 5)
            {
                $EmpNo = $LicChg + "@global.ul.com"
            }
		}
		else
		{
			write-host "`nUser Account" $EmpNo "does not exist" -ForegroundColor Red
	      	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Mailbox Does Not Exist for " + $EmpNo
	      	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
            write-host "`nEnter Employee Number of Individual to make License Changes for or 0 to Exit: " -ForegroundColor Yellow -NoNewline
            $ENo = Read-Host
            $EmpNo = $ENo + "@global.ul.com"
			$LicChg = "99"
		}
    }
    else
    {
        $LicChg = "0"
    }

	$ContReassign = "Y"
	
    If ($LicChg -ne 0)
    {
	
		switch ($LicChg)
		{
			1
            {
                write-host "Checking that the PowerBI (Free), EMS and Advanced Threat Protection Plan1 licenses are assigned"
				StandardLicenses

#                If ($HasExP2 -eq $True)
#                {
#                    Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:EXCHANGEENTERPRISE"
#                    write-host "`nRemoving Exchange Online P2 License and reassigning to a " $LicType "License"
#                    $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing P2 License from " + $EmpNo#
#	                Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
#                }
                AssignE3License
			}

            2
            {
    			If ($HasCRM -eq "True")
   				{
					Write-Host "`nEmployee " $EmpNo "Already has a Microsoft Dynamics CRM Online License Assigned" -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "CRM/Dynamics License Already Assigned to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				}
				else
				{
                    $CRMLic = Get-MsolAccountSku |where {$_.AccountSkuID -eq "ul:DYN365_ENTERPRISE_PLAN1"}
                    IF ($CRMLic.ConsumedUnits -lt $CRMLic.ActiveUnits)
                    {
                        Write-Host "`nAssigning Microsoft Dynamics CRM Online License to" $EmpNo -ForegroundColor Yellow
        	      	    $LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Microsoft Dynamic CRM Online License to " + $EmpNo
	      	            Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                        AssignCRM
                    }
                    else
                    {
                        Write-Host "`nNo Microsoft Dynamics CRM Online License available to assign to" $EmpNo -ForegroundColor Yellow
        	      	    $LineToWrite = $RecordEvent + "REVI" + "`t" + "No Microsoft Dynamics CRM Online License available to assign to " + $EmpNo
	      	            Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    }
                            
                }
			}

            3
            {
				If ($HasBIFree -eq "True")
				{
					Write-Host "`nEmployee " $EmpNo "Already has the PowerBI (Free) License Assigned" -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "PowerBI (Free) License Already Assigned to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				}
				else
				{
					Write-Host "Assigning PowerBI (Free) to" $EmpNo -ForegroundColor Yellow
					Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:POWER_BI_STANDARD"
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning PowerBI (Free) License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				}
			}

            4
            {
				If ($HasEMS -eq "True")
				{
					Write-Host "`nEmployee " $EmpNo "Already has the Enterprise Mobility Suite/Intune (EMS) License Assigned" -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Enterprise Mobility Suite/Intune (EMS) License Already Assigned to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				}
				else
				{
					Write-Host "`nAssigning Enterprise Mobility Suite/Intune (EMS) License to" $EmpNo -ForegroundColor Yellow
					Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:EMS"
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Enterprise Mobility Suite/Intune (EMS) License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				}  
            }

            5
            {
                if ($HasMRSS -eq "True")
                {
                    write-host "`nThis individual already has a Microsoft Relationship Sales Solution License Assigned"
                    $LineToWrite = $RecordEvent + "REVI" + "`t" + "Microsoft Relationship Sales Solution License Already Assigned to " + $EmpNo
                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
                    $MRSSLic = Get-MsolAccountSku |where {$_.AccountSkuID -eq "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"}
                    If ($MRSSLIc.ConsumedUnits -lt $CRMliC.ActiveUnits)
                    {
                        Write-Host "`nAssigning Microsoft Relationship Sales Solution License to" $EmpNo -ForegroundColor Yellow
        	      	    $LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Microsoft Relationship Sales Solution License to " + $EmpNo
	      	            Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                        AssignMRSS
                    }
                    else
                    {
                        Write-Host "`nNo Microsoft Relationship Sales Solution License available to assign to " + $EmpNo
        	      	    $LineToWrite = $RecordEvent + "REVI" + "`t" + "No Microsoft Relationship Sales Solution License available to assign to " + $EmpNo
	      	            Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    }
                }
            }
			
            6
            {
				If ($HasATP1 -eq "True")
				{
					Write-Host "`nEmployee " $EmpNo "Already has the Advanced Threat Protection Plan1 License Assigned" -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Advanced Threat Protection Plan1 License Already Assigned to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				}
				else
				{
					Write-Host "`nAssigning Advanced Threat Protection Plan1 License to" $EmpNo -ForegroundColor Yellow
					Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:ATP_ENTERPRISE"
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Advanced Threat Protection Plan1 License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				}  
            }

            10
            {
				If ($HasCRM -eq "True")
				{
					Write-Host "`nRemoving Dynamics 365 Customer Engagement Plan License from" $EmpNo -ForegroundColor Yellow
					Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:DYN365_ENTERPRISE_PLAN1"
        	      	$LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing Dynamics 365 Customer Engagement Plan License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host "`nEmployee " $EmpNo "Does Not have a Dynamics 365 Customer Engagement Plan License Assigned" -ForegroundColor Yellow
         	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove Dynamics 365 Customer Engagement Plan License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
               }
            }

            11
            {
				If ($HasBIFree -eq "True")
				{
					Write-Host "`nRemoving PowerBI (Free) License from" $EmpNo -ForegroundColor Yellow
					Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:POWER_BI_STANDARD"
        	      	$LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing PowerBI (Free) License from " + $EmpNo + "`n"
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host "`nEmployee " $EmpNo "Does Not have a PowerBI (Free) License Assigned" -ForegroundColor Yellow
          	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove PowerBI (Free) License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
               }
            }

            12
            {
 				If ($HasEMS -eq "True")
 				{
					Write-Host "`nRemoving Enterprise Mobility Suite/Intune (EMS) License from" $EmpNo -ForegroundColor Yellow
					Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:EMS"
        	      	$LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing nterprise Mobility Suite/Intune (EMS) License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host "`nEmployee " $EmpNo "Does Not have a Enterprise Mobility Suite/Intune (EMS) License Assigned" -ForegroundColor Yellow
         	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove Enterprise Mobility Suite/Intune (EMS) License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            13
            {
                If ($HasMRSS -eq "True")
                {
			        Write-Host "`nRemoving Microsoft Relationship Sales Solution License from" $EmpNo -ForegroundColor Yellow
				    Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing Microsoft Relationshp Sales Solution License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host "`nEmployee " $EmpNo "Does Not have a Microsoft Relationship Sales Solution License Assigned" -ForegroundColor Yellow
         	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove Microsoft Relationship Sales Solution License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            14
            {
 				If ($HasATPP1 -eq "True")
 				{
					Write-Host "`nRemoving Advanced Threat Protection Plan1 License from" $EmpNo -ForegroundColor Yellow
					Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:ATP_ENTERPRISE"
        	      	$LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing Advanced Threat Protection Plan1 License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host "`nEmployee " $EmpNo "Does Not have a Advanced Threat Protection Plan1 License Assigned" -ForegroundColor Yellow
         	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove Advanced Threat Protection Plan1 License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            20
            {
                $sskid = "ul:ENTERPRISEPACK"
#                $Global:UserLicense = Get-MsolUser -UserPrincipalName $EmpNo -ErrorAction SilentlyContinue
                EnabledFeature
            }

            21
           	{
                if ($HasE3)
                {
                    $SSKID = "ul:ENTERPRISEPACK"
                    $LicType = "Enterprise E3"
                    $DisPlan = $DisPlanEmpE3
                    if ($ADUser.ExtensionAttribute1 -notlike "Employee*")
                    {
                        $DisPlan = $Global:DisPlanNonEmpE3
                    }
                }

                If ($ADUser.ExtensionAttribute1 -like "Employee*")
                {
					write-host "Resetting " $LicType "licenses to standard employee licenses" -ForegroundColor Yellow				
				}
				else
				{
					write-host "Resetting " $LicType "licenses to standard non-employee licenses" -ForegroundColor Yellow						
				}

                EnabledFeature

      	      	$LineToWrite = $RecordEvent + "UPDA" + "`t" + "Resetting " + $LicType + " SubLicense Options to Standard Employee/Non-Employee for " + $EmpNo
      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				$DLO = ($DisPlan.Split(“,”))
				$MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
                Set-MsolUserLicense -UserPrincipalName $EmpNo –LicenseOptions $MyO365Sku
						
    			RetentPolicy
            }
            
            99
            {
#                Do Nothing User Account/Mailbox does not exist
            }                										
    	}
    }

}While ($LicChg -ne 0)