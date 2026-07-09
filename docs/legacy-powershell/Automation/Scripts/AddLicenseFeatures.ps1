<#################################################################################
# 
# PowerShell source code
# Revision v1.05
# ==========================================================================
#    'Project      : UL Office 365 O365 Admin Team Menu
#    'Description  : Manage O365 Licences
#    'Called By    : O365AdminMenu.ps1
#    'Calls        : O365DisabledLicenseFeatures.ps1
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 04/21/2016
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '           06/17/2018 SAG To rest license options for new features
# ==========================================================================
#
#################################################################################>

$WhoAmI			= WhoAmI

	
# Retrieve the current Date and Time for use in log files
$uDate = get-date -uformat %D
$uTime = get-date -uformat %T

Function EnabledFeature
{
    write-host "`n`tEnabled SubLicense Options:" -ForegroundColor Magenta

    If ($HasE3 -eq "True")
    {
        $SSKID = "UL:ENTERPRISEPACK"
        $LicType = "Enterprise E3"
    }
    else
    {
        $LicType = "Enterprise E4"
    }

    $SubLic = ($UserLicense.Licenses | where {$_.AccountSkuId -eq $SSKID}).ServiceStatus
	foreach ($SubLic in $SubLic)
	{
        $AssignWho = "Enabled All Types"

        switch ($subLic.ServicePlan.ServiceName)
        {
            "BPOS_S_TODO_2"
            {
                $LicName = "To-Do (Plan 2)"
            }
            "FORMS_PLAN_E3"
            {
                $LicName = "Microsoft Forms (Plan E3)"
            }
            "stream_O365_E3"
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
                $AssignWho = "Disabled by Default"
            }
            "POWERAPPS_O365_P2"
            {
                $LicName = "PowerApps for Office 365"
                $AssignWho = "Disabled by Default"
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
            "MCOVOICECONF"
            {
                $LicName = "Skype for Business Online (Plan 3)"
                $AssignWho = "Disabled by Default"
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
	}
}

Function EnabledMRSSFeature
{
    write-host "`n`tEnabled SubLicense Options:" -ForegroundColor Magenta

    $text = "`t{0}`t`t{1}`t`t`t{2}" –f "DefaultAssignment","Status ","SubLicenseName"
    $text
    $text = "`t{0}`t`t{1}`t`t`t{2}" –f "-----------------","------ ","--------------"
    $text

    $LineToWrite = $RecordEvent + "UPDA" + "`t" + $LicType + " SubLicenses Enabled/Status for " + $EmpNo
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    $SubLic = ($UserLicense.Licenses | where {$_.AccountSkuId -eq $SSKID}).ServiceStatus
	foreach ($SubLic in $SubLic)
	{
        $AssignWho = "Disabled All Types"

        switch ($subLic.ServicePlan.ServiceName)
        {
           "EXCHANGE_S_FOUNDATION"
           {
               $LicName = "No Details in O365 Portal"
           }
            "SHAREPOINTENTERPRISE"
            {
                $LicName = "SharePoint Online (Plan 2)"
            }
            "PROJECT_ESSENTIALS"
            {
                $LicName = "Project Online Essentials"
            }
            "POWERAPPS_DYN_APPS"
            {
                $LicName = "PowerApps for Dynamics 365"
            }
            "SHAREPOINTWAC"
            {
                $LicName = "Office Online"
            }
            "NBENTERPRISE"
            {
                $LicName = "Microsoft Social Engagement Enterprise"
            }
            "FLOW_DYN_APPS"
            {
                $LicName = "Flow for Dynamics 365"
            }
            "DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
            {
                $LicName = "Microsoft Relationship Sales solution"
                $AssignWho = "Enabled by Default"
            }
        }

        $LicStat = $SubLic.ProvisioningStatus
        If (($LicStat -eq "PendingActivation") -or ($LicStat -eq "PendingProvisioning"))
        {
            $LicStat = "Pending "
        }
        If ($LicStat -eq "Success")
        {
            $LicStat = "Enabled "
        }
        $text = "`t{0}`t`t{1}`t`t{2}" –f $AssignWho,$LicStat,$LicName
        $text
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + $text
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}
}

Function AssignE3E4License
{
	Write-host ""
    $sskid = "ul:ENTERPRISEPACK"
    $LicType = "Enterprise E3"
    $DisPlan = $Global:DisPlanEmpE3
    $DisPlan = $Global:DisPlanNonEmpE3
	
	$LicDet = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEWITHSCAL"}
			
	if ($LicDet.ActiveUnits -gt $LicDet.ConsumedUnits)
	{					
        $SSKID = "ul:ENTERPRISEWITHSCAL"
        $LicType = "Enterprise E4"
        $DisPlan = $DisPlanEmpE4
        $DisPlan = $Global:DisPlanNonEmpE4
    }

    If ($HasE3 -eq "True")
    {
        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:ENTERPRISEPACK"
        write-host "`nRemoving Enterprise E3 License and reassigning to a " $LicType "License"
        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing E3 License from " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }

    If ($HasExP2 -eq "True")
    {
        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:EXCHANGEENTERPRISE"
        write-host "`nRemoving Exchange Online P2 License and reassigning to a " $LicType "License"
        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing P2 License from " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
				
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
							
	If ($UserDetails -eq "True")
    {
        RetentPolicy
    }
}

Function AssignMRSS
{
	Write-host "Resetting  MRSS Licenise features" -ForegroundColor Yellow
    $sskid = "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
    $LicType = "Microsoft Relationship Sales Solution"
    $DisPlan = $Global:DisMRSSPlan

	$LicDet = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq $SSKID}
			
    $DLO = ($DisPlan.Split(“,”))
    $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
    Set-MsolUserLicense -UserPrincipalName $EmpNo –LicenseOptions $MyO365Sku
}

Function O365Licenses
{
    $O365Lic = (Get-MsolUser -UserPrincipalName $EmpNo).Licenses
	Write-Host "`nCurrent License Assignment for " $EmpNo " " $UserLicense.DisplayName "-" $ADUser.ExtensionAttribute1-ForegroundColor Green
    
    If ($O365Lic.count -gt 0)
    {
        foreach ($O365Lic in $O365Lic)
        {
            switch ($O365Lic.AccountSkuID)
            {
                "ul:ENTERPRISEWITHSCAL"
                {
                    $text = "`tEnterprise E4"
                }
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
                "ul:CRMSTANDARD"
                {
                    $text = "`tCRM/Dynamics"
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
                    $text = "`tAdvanced Threat Protection"
                }
            }
            Write-Host $text
       }
    }
}


Invoke-Expression -command E:\O365AdminShared\Scripts\O365DisabledLicenseFeatures.ps1

$coluser = import-csv c:\temp\users.csv
$coluser.count

foreach ($coluser in $coluser)
{
	$EmpNo = $coluser.UserPrincipalName
	
	$UserLicense = Get-MsolUser -UserPrincipalName $EmpNo

    $UserDetails = [bool](get-mailbox $EmpNo -ErrorAction SilentlyContinue)
    $HasBIFree = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:POWER_BI_STANDARD"})
    $HasBIPro = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:POWER_BI_PRO"})
    $HasCRM = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:CRMSTANDARD"})
    $HasEMS = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:EMS"})
    $HasExP2 = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:EXCHANGEENTERPRISE"})
    $HasE4 = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:ENTERPRISEWITHSCAL"})
    $HasE3 = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:ENTERPRISEPACK"})
    $HasMRSS = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"})
    $HasATP = [bool] ($UserLicense.Licenses | where {$_.AccountSkuId -eq "ul:ATP_ENTERPRISE"})
			
	$LDAPFilter = "(userPrincipalName=" + $EmpNo + ")"
	$ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,proxyaddresses,extensionattribute1,extensionattribute4
    O365Licenses

    if ($HasE4)
    {
        $SSKID = "ul:ENTERPRISEWITHSCAL"
        $LicType = "Enterprise E4"
        $DisPlan = $Global:DisPlanEmpE4
        if ($ADUser.ExtensionAttribute1 -notlike "Employee*")
        {
            $DisPlan = $Global:DisPlanNonEmpE4
        }
    }
    else
    {
        $SSKID = "ul:ENTERPRISEPACK"
        $LicType = "Enterprise E3"
        $DisPlan = $DisPlanEmpE3
        If ($ADUser.ExtensionAttribute1 -notlike "Employee*")
        {
            $DisPlan = $Global:DisPlanNonEmpE3
        }
    }

	If ($ADUser.ExtensionAttribute1 -like "Employee*")
    {
		write-host "Resetting " $LicType "licenses to standard employee licenses" -ForegroundColor Yellow				
	}
	elseif ($ADUser.ExtensionAttribute1 -notlike "Ex-*")
    {
		write-host "Resetting " $LicType "licenses to standard non-employee licenses" -ForegroundColor Yellow						
    }

    If ($ADUser.ExtensionAttribute -notlike "Ex-*")
	{
        $DLO = ($DisPlan.Split(“,”))
	    $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $EmpNo –LicenseOptions $MyO365Sku
    }

    If ((($HasEMS -eq $False) -and ($HasP2 -eq $False)) -and (($HasE4 -eq $True) -or ($HasE3 -eq $True)))
    {
        write-host "Assigning  EMS License to " $EmpNo
        Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:EMS"
    }

    If ($HasMRSS -eq $True)
    {
        AssignMRSS
    }

    If ($HasCRM -eq $True)
    {
        write-host "Resetting  CRM License features" -ForegroundColor Yellow

        $SSKID ="ul:CRMSTANDARD"
        $DisPlan = $Global:DisCRMPlan
        $DLO = ($DisPlan.Split(“,”))
        $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $EmpNo –LicenseOptions $MyO365Sku
    }
}
