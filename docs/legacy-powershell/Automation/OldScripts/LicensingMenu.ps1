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
#    '           11/15/2017 SAG Added Code to check if there are available EMS licenses
#    '               and to see if an E3 or E4 license is already assigned
#                11/26/2017 - SAG - Added the Microsoft Relationship Sales Solution licence
#                    also change numbering to provide gaps for adding of new license types
#                11/27/2017 - SAG - Added option to assign a P2 license
#                01/16/2018 - SAG - Added code in the move from E4 to P2 to check if individual
#                    is terminated and if so to check all other assigned licenses and remove them
#                01/30/2018 - SAG - Added to to assign a PowerBI Pro license and remove the 
#                    PowerBI free license
#                03/05/2018 - SAG Changed the IT retention policy to 3 yr
#                07/10/2018 - SAG Added StandardLicense set to run for Option 10 and 11
#                08/09/2018 - SAG Changed Flow and PowerApps to "Enabled All Types"
#                10/10/2018 - SAG Added check if PowerBI Pro license is assigned when requesting a PowerBI Free license be assigned.
#                01/29/2018 - SAG Added code to make sure the UsageLocation is set if assigning a license.
#                02/08/2019 - SAG Added option to show current license allocations also fixed the process for existing the script
#                02/19/2019 - SAG Fixed so that the details are updated when entering new employee number 
#                04/16/2019 - SAG Modifed CRM to the new license type which has Employee/Non-Employee features
#                04/20/2019 - SAG Modified options to remove E4 license
#                06/14/2019 - SAG Removed the ATP license options
#                08/16/2018 - SAG Added new features that are part of the E3 license and Meeting Room and PowerApps P2 license
#                08/21/2019 - SAG Added PhoneSys and AudioConf licenses
#                05/07/2020 - SAG Added Insider Risk Management license
# ==========================================================================
#
#################################################################################>

$WhoAmI			= WhoAmI

$LogDirectory   = "E:\Automation\Licensing\Log"
$LogFile		= $LogDirectory + "\" + "Log-Licensing.log"
	
# Retrieve the current Date and Time for use in log files
$uDate = get-date -uformat %D
$uTime = get-date -uformat %T
$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"

Invoke-Expression -command E:\O365AdminShared\Scripts\O365DisabledLicenseFeatures.ps1

Function LicMenu
{
    write-host "`n`t`t`t`tLicensing Menu" -ForegroundColor Magenta
	write-host
	write-Host "     Enter (  1) Assign E3 and Standard License Set"
    write-host "           (  2) Assign Dynamics 365 Customer Engagement Plan License"
    write-host "           (  3) Assign PowerBI (Free) License"
    write-host "           (  4) Assign Enterprise Mobility (EMS) License"
    write-host "           (  5) Assign Microsoft Relationship Sales Solution License"
    write-host "           (  6) Assign Exchange Online (Plan2) License"
    write-host "           (  7) Assign PowerBI Pro License"
    write-host "           (  8) Assign Defender Advanced Threat Protection License"
    write-host "           ( 8a) Assign Advanced Threat Protection Plan1 License"
    write-host "           ( 8b) Assign Advanced Threat Protection for MAC"
    write-host "           (  9) Assign Meeting Room License"
    write-host "           ( 10) Assign PowerApps Plan 2 License"
    write-host "           ( 11) Assign Phone System License"
    write-host "           ( 12) Assign Audio Conferencing License"
    write-host "           ( 13) Assign Remote Assist License Set"
    write-host "`n           ( 20) Move from P2 to E3 and Standard License Set"
    write-host "           ( 21) Move from E3 to P2 License"
    write-host "`n           ( 30) Remove CRM/Dynamics License"
    write-host "           ( 31) Remove PowerBI (Free) License"
    write-host "           ( 32) Remove Enterprise Mobility Suite/Intune (EMS) License"
    write-host "           ( 33) Remove Microsoft Relationship Sales solution"
    write-host "           ( 34) Remove Defender Advanced Threat Protection"
    write-host "           (34a) Remove Advanced Threat Protection Plan1 License"
    write-host "           (34b) Remove Advanced Threat Protection for MAC"
    write-host "           ( 35) Remove Meeting Room License"
    write-host "           ( 36) Remove PowerApps Plan 2 License"
    write-host "           ( 37) Remove Phone System License"
    write-host "           ( 38) Remove Audio Conferencing License"
    write-host "           ( 39) Remove Advanced Threat Protection Plan1 License"
    write-host "           ( 40) Remove Remote Assist License Set"
    Write-Host "           ( 41) Remove Insider Risk Management License"
    write-host "`n           ( 50) Review Enabled E3 Sub-License Features"
    Write-Host "           ( 51) Reset Enabled E3 Sub-License Features to Employee/Non-Employee Standard"
    write-host "           ( 52) Review Enabled Microsoft Relationship Sales Solution License Features"
    write-Host "`n           ( 90) Show License Allocations"
	write-Host "`n           ( 0) to Return to the O365 Admin Menu"
	write-Host "     Enter Emp# or Option? " -ForegroundColor Red -NoNewline
	$Global:LicChg = Read-Host

    If ($Global:LicChg.Length -ge 5)
    {
        $Global:ENo = $Global:LicChg
    }
}

Function AvailLicense
{
    $E3Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}
    $P2Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EXCHANGEENTERPRISE"}
    $EMSLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EMS"}
    $ATPP1Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ATP_ENTERPRISE"}
    $CRMLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:DYN365_ENTERPRISE_PLAN1"}
    $PBIFLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWER_BI_STANDARD"}
    $PBIPLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWER_BI_PRO"}
    $MRSSLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"}
    $FP2Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:FLOW_P2"}
    $ATPDefLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:WIN_DEF_ATP"}
    $ATPDefMAC = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MDATP_XPLAT"}
    $MeetingLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MEETING_ROOM"}
    $PAppsP2Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:POWERFLOW_P2"}
    $PhoneSysLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOEV"}
    $AudioConfLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MCOMEETADV"}
    $RmtAssistLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:MICROSOFT_REMOTE_ASSIST"}
    $InfoBarrLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:M365_INSIDER_RISK_MANAGEMENT"}
        
#    Send O365 Team email that the Staff Type Changes is complete

    write-host "          Current license allocations:"
    write-host "`nActive`t InUse`t Avail`t-  License Name"
    write-host "-------------------------------------------------------------"

    write-host $E3Lic.ActiveUnits "`t" $E3Lic.ConsumedUnits "`t" ($E3Lic.ActiveUnits-$E3Lic.ConsumedUnits) "`t-  Enterprise E3 Licenses"
    write-host $P2Lic.ActiveUnits "`t" $P2Lic.ConsumedUnits "`t" ($P2Lic.ActiveUnits-$P2Lic.ConsumedUnits) "`t-  ExchangeOnline P2 Licenses"
    write-host $EMSLic.ActiveUnits "`t" $EMSLic.ConsumedUnits "`t" ($EMSLic.ActiveUnits-$EMSLic.ConsumedUnits) "`t-  Enterprise Mobility + Security E3 License"
    write-host $ATPP1Lic.ActiveUnits "`t" $ATPP1Lic.ConsumedUnits "`t" ($ATPP1Lic.ActiveUnits-$ATPP1Lic.ConsumedUnits) "`t-  Advanced Threat Protection Plan1 License"
    write-host $ATPDefLic.ActiveUnits "`t" $ATPDefLic.ConsumedUnits "`t" ($ATPDefLic.ActiveUnits-$ATPDefLic.ConsumedUnits) "`t-  Defender Advanced Threat Protection License"
    write-host $ATPDefMAC.ActiveUnits "`t" $ATPDefMAC.ConsumedUnits "`t" ($ATPDefMAC.ActiveUnits-$ATPDefMAC.ConsumedUnits) "`t-  Advanced Threat Protection for MAC License"
    write-host $CRMLic.ActiveUnits "`t" $CRMLic.ConsumedUnits "`t" ($CRMLic.ActiveUnits-$CRMLic.ConsumedUnits) "`t-  Dynamics 365 Customer Engagement Plan Licenses"
    write-host $MRSSLic.ActiveUnits "`t" $MRSSLic.ConsumedUnits "`t" ($MRSSLic.ActiveUnits-$MRSSLic.ConsumedUnits) "`t-  Microsoft Relationship Sales solution Licenses"
    write-host "N/A`t" $PBIFLic.ConsumedUnits "`t N/A`t-  PowerBI (Free) Licenses"
    write-host $PBIPLic.ActiveUnits "`t" $PBIPLic.ConsumedUnits "`t" ($PBIPLic.ActiveUnits-$PBIPLic.ConsumedUnits) "`t-  PowerBI Pro Licenses"
    write-host $FP2Lic.ActiveUnits "`t" $FP2Lic.ConsumedUnits "`t" ($FP2Lic.ActiveUnits-$FP2Lic.ConsumedUnits) "`t-  Microsoft Flow (Plan2) Licenses"
    write-host $MeetingLic.ActiveUnits "`t" $MeetingLic.ConsumedUnits "`t" ($MeetingLic.ActiveUnits-$MeetingLic.ConsumedUnits) "`t-  Meeting Room Licenses"
    write-host $PAppsP2Lic.ActiveUnits "`t" $PAppsP2Lic.ConsumedUnits "`t" ($PAppsP2Lic.ActiveUnits-$PAppsP2Lic.ConsumedUnits) "`t-  PowerApps Plan 2 Licenses"
    write-host $PhoneSysLic.ActiveUnits "`t" $PhoneSysLic.ConsumedUnits "`t" ($PhoneSysLic.ActiveUnits-$PhoneSysLic.ConsumedUnits) "`t-  Phone System Licenses"
    write-host $AudioConfLic.ActiveUnits "`t" $AudioConfLic.ConsumedUnits "`t" ($AudioConfLic.ActiveUnits-$AudioConfLic.ConsumedUnits) "`t-  Audio Conferencing Licenses"
    write-host $RmtAssistLic.ActiveUnits "`t" $RmtAssistLic.ConsumedUnits "`t" ($RmtAssistLic.ActiveUnits-$RmtAssistLic.ConsumedUnits) "`t-  Microsoft Remote Assist Licenses"
    write-host $InfoBarrLic.ActiveUnits "`t" $InfoBarrLic.ConsumedUnits "`t" ($InfoBarrLic.ActiveUnits-$InfoBarrLic.ConsumedUnits) "`t-  E5 Insider Risk Management Licenses"
}

Function RetentPolicy
{
    $MbxCreated = [bool](get-mailbox $EmpNo -ErrorAction SilentlyContinue)

    If ($MbxCreated -eq "True")
    {
        If (($ADUser.extensionattribute4 -eq "IT") -or ((get-mailbox $EmpNo).WhenCreated -gt "03/23/2020"))
        {
            Set-MailBox $EmpNo -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL MRM Policy - 3 yr Delete"
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Updating Retention Policy to IT Policy for " + $EmpNo
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        }
        else
        {
            Set-MailBox $EmpNo -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL Default MRM Policy - 365 Day Delete"
            Set-Mailbox $EmpNo -RetentionHoldEnabled $true -StartDateForRetentionHold 04/01/2011
            $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Updating Retention Policy to UL Default Policy for " + $EmpNo
	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        }
	    Set-CASMailBox $EmpNo -ImapEnabled $false -PopEnabled $false -ActiveSyncEnabled $false
    }
    else
    {
        write-host "Mailbox does not exist or is in the process of being created.  If the license was just assigned"
        write-host "please allow 3-5 minutes for the mailbox to be created and the select Option 10 from the menu"
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Mailbox Does Not Exist " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
}

Function EnabledMRSSFeature
{
    write-host "`n`tEnabled SubLicense Options:" -ForegroundColor Magenta

    $text = "`t{0}`t`t{1}`t`t`t{2}" -f "DefaultAssignment","Status ","SubLicenseName"
    $text
    $text = "`t{0}`t`t{1}`t`t`t{2}" -f "-----------------","------ ","--------------"
    $text

    $LineToWrite = $RecordEvent + "UPDA" + "`t" + $LicType + " SubLicenses Enabled/Status for " + $EmpNo
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    $SubLic = ($UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq $SSKID}).ServiceStatus
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
        $text = "`t{0}`t`t{1}`t`t{2}" -f $AssignWho,$LicStat,$LicName
        $text
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + $text
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
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

        $DLO = ($DisPlan.Split(","))
        $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense $SSKID -LicenseOptions $MyO365Sku
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

Function AssignMRSS
{
	Write-host ""
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
            write-host "Removing Dynamics 365 Customer Engagement Plan License from this account"
            Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:DYN365_ENTERPRISE_PLAN1"
        }
		write-host "`nAssigning " $EmpNo "a " $LicType "License" -ForegroundColor Yellow
        $LineToWrite = $RecordEvent + "UPDA" + "`t" + "Assigning " + $LicType + "to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	}

    $DLO = ($DisPlan.Split(","))
    $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense $SSKID -LicenseOptions $MyO365Sku

}

Function StandardLicenses
{
    If ($HasEMS -ne "True")
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
}

Function O365Licenses
{
    $O365Lic = (Get-MsolUser -UserPrincipalName $EmpNo).Licenses
	Write-Host "`nCurrent License Assignment for " $EmpNo " " $UserLicense.DisplayName "-" $ADUser.ExtensionAttribute1-ForegroundColor Green
    
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
                "ul:SHAREPOINTSTANDARD"
                {
                    $text = "`tSharePoint Online (Plan 1)"
                }
                "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
                {
                    $text = "`tMicrosoft Relationship Sales Solution"
                }
                "ul:TEAMS_COMMERCIAL_TRIAL"
                {
                    $text = "`tMicrosoft Teams Commercial Cloud (User Initiated)"
                }
                "ul:ATP_ENTERPRISE"
                {
                    $text ="`tAdvanced Threat Protection Plan1"
                }
                "ul:MDATP_XPLAT"
                {
                    $text ="`tAdvanced Threat Protection for MAC"
                }
                "ul:WIN_DEF_ATP"
                {
                    $text ="`tDefender Advanced Threat Protection"
                }
                "ul:FLOW_P2"
                {
                    $text = "`tMicrosoft Flow Plan 2"
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
                "ul:MICROSOFT_REMOTE_ASSIST"
                {
                    $text = "`tMicrosoft Remote Assist"
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

Function AudioConferencingAdd
{
    if ($HasAudioConf -eq "True")
    {
        write-host "`nThis individual already has an Audio Conferencing License Assigned"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "Audio Conferencing License Already Assigned to " + $EmpNo
        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
    else
    {
        Write-Host "`nAssigning Audio Conferencing License to" $EmpNo -ForegroundColor Yellow
      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Audio Conferencing License to " + $EmpNo
	    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:MCOMEETADV"
    }
}

Function PhoneSystemAdd
{
    if ($HasPhoneSys -eq "True")
    {
        write-host "`nThis individual already has a Phone System License Assigned"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "Phone System License Already Assigned to " + $EmpNo
        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
    else
    {
        Write-Host "`nAssigning Phone System License to" $EmpNo -ForegroundColor Yellow
     	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Phone System License to " + $EmpNo
        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:MCOEV"
    }            
}

Function InsiderRisk
{
    if ($HasInfoBarr -eq "True")
    {
        write-host "`nThis individual already has a E5 Insider Risk Management License Assigned"
        $LineToWrite = $RecordEvent + "REVI" + "`t" + "E5 Insider Risk Managemen License Already Assigned to " + $EmpNo
        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
    }
    else
    {
        Write-Host "`nAssigning E5 Insider Risk Managemen License to" $EmpNo -ForegroundColor Yellow
     	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning E5 Insider Risk Managemenm License to " + $EmpNo
        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
        Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:M365_INSIDER_RISK_MANAGEMENT"
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

AvailLicense

write-host "`nEnter Employee Number of Individual to make License Changes for or 0 to Exit: " -ForegroundColor Yellow -NoNewline
$Global:ENo = Read-Host

If($Global:ENo -ne 0)
{
    $Global:LicChg = 1
}

Do
{
    If ($Global:LicChg -ne "0")
    {

        $EmpNo = $ENo + "@global.ul.com"
        $LineToWrite = "`n" + $RecordEvent + "INFO" + "`t" + "Performing Licensing Review for " + $EmpNo
        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		if ([bool](Get-MsolUser -UserPrincipalName $EmpNo -ErrorAction SilentlyContinue) -eq "True")
		{

			$UserLicense = Get-MsolUser -UserPrincipalName $EmpNo

            $HasBIFree = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:POWER_BI_STANDARD"})
            $HasBIPro = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:POWER_BI_PRO"})
            $HasCRM = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:DYN365_ENTERPRISE_PLAN1"})
            $HasEMS = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:EMS"})
            $HasExP2 = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:EXCHANGEENTERPRISE"})
            $HasE3 = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:ENTERPRISEPACK"})
            $HasMRSS = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"})
            $HasFlowP2 = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:FLOW_P2"})
            $HasATPDef = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:WIN_DEF_ATP"})
            $HasATPDefMAC = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MDATP_XPLAT"})
            $HasATPP1 = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:ATP_ENTERPRISE"})
            $HasMeeting = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MEETING_ROOM"})
            $HasPAppsP2 = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:POWERFLOW_P2"})
            $HasPhoneSys = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MCOEV"})
            $HasAudioConf = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MCOMEETADV"})
            $HasRmtAssist = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MICROSOFT_REMOTE_ASSIST"})
            $HasInfoBarr = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:M365_INSIDER_RISK_MANAGEMENT"})
			
			$LDAPFilter = "(userPrincipalName=" + $EmpNo + ")"
		    $ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,proxyaddresses,extensionattribute1,extensionattribute4
            O365Licenses
#            Invoke-Expression E:\O365AdminShared\Scripts\O365LicensesAssignment.ps1

			LicMenu
		}
		else
		{
            If ($Global:ENo -eq 0)
            {
                $Global:LicChg = "0"
            }
            else
            {
    			write-host "`nUser Account" $EmpNo "does not exist" -ForegroundColor Red
	          	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Mailbox Does Not Exist for " + $EmpNo
	          	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                write-host "`nEnter Employee Number of Individual to make License Changes for or 0 to Exit: " -ForegroundColor Yellow -NoNewline
                $Global:ENo = Read-Host
                $EmpNo = $Global:ENo + "@global.ul.com"
	    		$Global:LicChg = "99"
            }
		}
    }
    else
    {
        $Global:LicChg = "0"
    }

	$ContReassign = "Y"
	
    If (($Global:LicChg -ne 0) -and ($Global:LicChg.Length -le 3))
    {
        If ($UserLicense.IsLicensed -eq $fALSE)
        {
            Set-MsolUser -UserPrincipalName $EmpNo -UsageLocation "US"
        }
	
		switch ($Global:LicChg)
		{
			1
            {
                If ($HasE3 -eq "True")
                {
                    $SSKID = "ul:ENTERPRISEPACK"
                    $LicType = "Enterprise E3"
                    $DisPlan = $Global:DisPlanEmpE3
                    if ($ADUser.ExtensionAttribute1 -notlike "Employee*")
                    {
                        $DisPlan = $Global:DisPlanNonEmpE3
                    }

                    If ($ADUser.ExtensionAttribute1 -like "Employee*")
                    {
		    			write-host "Resetting " $LicType "licenses to standard employee licenses" -ForegroundColor Yellow				
			    	}
				    else
    				{
	    				write-host "Resetting " $LicType "licenses to standard non-employee licenses" -ForegroundColor Yellow						
		    		}

                    EnabledE3Feature

          	      	$LineToWrite = $RecordEvent + "UPDA" + "`t" + "Resetting " + $LicType + " SubLicense Options to Standard Employee/Non-Employee for " + $EmpNo
          	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		    		$DLO = ($DisPlan.Split(","))
			    	$MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
                    Set-MsolUserLicense -UserPrincipalName $EmpNo -LicenseOptions $MyO365Sku
                    RetentPolicy
                }
                else
                {
                    AssignE3License
                }
				write-host "Checking the Retention Policy Configuration and that the PowerBI and EMS licenses are assigned"
            	StandardLicenses
            }

            2
            {
    			If ($HasCRM -eq "True")
   				{
					Write-Host "`nEmployee" $EmpNo "Already has a Dynamics 365 Customer Engagement Plan License Assigned" -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Dynamics 365 Customer Engagement Plan License Already Assigned to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				}
				else
				{
                    If ($HasMRSS -eq "True")
                    {
                        write-host "Removing Microsoft Relationship Sales Solution License from this account"
                        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
                    }
                    
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

                    Write-Host "`nAssigning Dynamics 365 Customer Engagement Plan License to" $EmpNo -ForegroundColor Yellow
                    $DLO = ($DisPlan.Split(","))
                    $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
                    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense $SSKID -LicenseOptions $MyO365Sku
           	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Dynamics 365 Customer Engagement Plan License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				}
			}

            3
            {
				If ($HasBIPro -eq $True)
                {
                    $ProToFreeLicense = "N"
                    write-host "`nEmployee" $EmpNo "Has a PowerBI Pro License Assigned" -ForegroundColor Red
                    $LineToWrite = $RecordEvent + "REVI" + "`t" + "PowerBI (Pro) License Assigned to " + $EmpNo
                    write-host "`nDo you want to remove the Pro License and Reassign a Free License (Y/N)? " -ForegroundColor Cyan -NoNewline
                    $ProToFreeLicense = Read-Host

                    If ($ProToFreeLicense -eq "Y")
                    {
                            Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicense "ul:POWER_BI_PRO"
                            write-host "`nEmployee" $EmpNo "PowerBI Pro License Removed" -ForegroundColor Red
                            $LineToWrite = $RecordEvent + "REVI" + "`t" + "PowerBI Pro License Removed from " + $EmpNo
                    }
                    $HasBIPro = [bool] ($UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:POWER_BI_PRO"})
                }
                                
                If (($HasBIFree -eq $True) -and ($HasPBIPro -eq $False))
				{
					Write-Host "`nEmployee" $EmpNo "Already has the PowerBI (Free) License Assigned" -ForegroundColor Yellow
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
					Write-Host"`nEmployee" $EmpNo "Already has the Enterprise Mobility Suite/Intune (EMS) License Assigned" -ForegroundColor Yellow
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
                    Write-Host "`nAssigning Microsoft Relationship Sales Solution License to" $EmpNo -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Microsoft Relationship Sales Solution License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    AssignMRSS
                }
            }

            6
            {
                if ($HasExP2 -eq "True")
                {
                    write-host "`nThis individual already has a Exchange Online P2 License Assigned"
                    $LineToWrite = $RecordEvent + "REVI" + "`t" + "Exchange Online P2 License Already Assigned to " + $EmpNo
                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
                    Write-Host "`nAssigning Exchange Online P2 License to" $EmpNo -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Exchange Online P2 License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:EXCHANGEENTERPRISE"
                }
            }

            7
            {
                if ($HasBIPro -eq "True")
                {
                    write-host "`nThis individual already has a PowerBI Pro License Assigned"
                    $LineToWrite = $RecordEvent + "REVI" + "`t" + "PowerBI Pro License Already Assigned to " + $EmpNo
                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
                    Write-Host "`nAssigning PowerBI Pro License to" $EmpNo -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning PowerBI Pro License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:POWER_BI_PRO"

                    If ($HasBIFree -eq "True")
                    {
                        Write-Host "`nRemoving PowerBI (Free) License from" $EmpNo -ForegroundColor Yellow
        	      	    $LineToWrite = $RecordEvent + "REVI" + "`t" + "Removing PowerBI (free) License from" + $EmpNo
	      	            Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicense "ul:POWER_BI_STANDARD"
                    }
                }                
            }

            8
            {
                if ($HasATPDef -eq "True")
                {
                    write-host "`nThis individual already has a Windows Defender Advanced Threat Protection License Assigned"
                    $LineToWrite = $RecordEvent + "REVI" + "`t" + "Windows Defender Advanced Threat Protection License Already Assigned to " + $EmpNo
                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
                    Write-Host "`nAssigning Windows Defender Advanced Threat Protection License to" $EmpNo -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Windows Defender Advanced Threat Protection License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:WIN_DEF_ATP"
                }
            }

            8a
            {
                if ($HasATPP1 -eq "True")
                {
                    write-host "`nThis individual already has a Advanced Threat Protection Plan1 License Assigned"
                    $LineToWrite = $RecordEvent + "REVI" + "`t" + "Advanced Threat Protection Plan1 License Already Assigned to " + $EmpNo
                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
                    Write-Host "`nAssigning Advanced Threat Protection Plan1 License to" $EmpNo -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Advanced Threat Protection Plan1 License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:ATP_ENTERPRISE"
                }
            }
 
            8b
            {
                if ($HasATPDefMAC -eq "True")
                {
                    write-host "`nThis individual already has a Advanced Threat Protection for MAC License Assigned"
                    $LineToWrite = $RecordEvent + "REVI" + "`t" + "Advanced Threat Protection for MAC License Already Assigned to " + $EmpNo
                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
                    Write-Host "`nAssigning Advanced Threat Protection for MAC License to" $EmpNo -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Advanced Threat Protection for MAC License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:MDATP_XPLAT"
                }
            }

            9
            {
                if ($HasMeeting -eq "True")
                {
                    write-host "`nThis individual already has a Meeting Room License Assigned"
                    $LineToWrite = $RecordEvent + "REVI" + "`t" + "Meeting Room License Already Assigned to " + $EmpNo
                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
                    Write-Host "`nAssigning Meeting Room License to" $EmpNo -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Meeting Room License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:MEETING_ROOM"
                }
            }
 
            10
            {
                if ($HasPAppsP2 -eq "True")
                {
                    write-host "`nThis individual already has a PowerApps Plan 2 License Assigned"
                    $LineToWrite = $RecordEvent + "REVI" + "`t" + "PowerApps Plan 2 License Already Assigned to " + $EmpNo
                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
                    Write-Host "`nAssigning PowerApps Plan 2 License to" $EmpNo -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning PowerApps Plan 2 License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:POWERFLOW_P2"
                }
            }

            11
            {
                PhoneSystemAdd            
            }

            12
            {
                AudioConferencingAdd
            }

            13
            {
                if ($HasRmtAssist -eq "True")
                {
                    write-host "`nThis individual already has a Microsoft Remote Assist License Assigned"
                    $LineToWrite = $RecordEvent + "REVI" + "`t" + "Microsoft Remote Assistg License Already Assigned to " + $EmpNo
                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
                    Write-Host "`nAssigning Microsoft Remote Assist License to" $EmpNo -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Assigning Microsoft Remote Assist License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:MICROSOFT_REMOTE_ASSIST"
                }
                AudioConferencingAdd
                PhoneSystemAdd
#                InsiderRisk   
            }
 
            20
            {
                $SSKID = "ul:ENTERPRISEPACK"
                $LicDet = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq $SSKID}
						
				if ($LicDet.ActiveUnits -gt $LicDet.ConsumedUnits)
				{
				    Write-host ""
					write-host "Moving " $EmpNo "from Exchange Online P2 License to an Enterprise E3 License" -ForegroundColor Yellow
		 
					Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:EXCHANGEENTERPRISE"
					
                    if ($ADUser.ExtensionAttribute1 -notlike "Employee*")
                    {
                        $DisPlan = $Global:DisPlanNonEmpE3
                    }
                    else
                    {
                        $DisPlan = $Global:DisPlanEmpE3
                    }
                    
					$DLO = ($DisPlan.Split(","))
					$MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
					Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense $SSKID -LicenseOptions $MyO365Sku
					RetentPolicy
				}
				else
				{
					Write-Host "There are No Enterprise E3 Licenses available for assignment" -ForegroundColor Red
				}
                StandardLicenses
                write-host ""
            }
            
            21
            {
                $LicDet = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EXCHANGEENTERPRISE"}
						
				if (($ADUser.ExtensionAttribute1 -like "Ex-*") -or ($ADUser.ExtensionAttribute1 -like "*Service*"))
                {
                    if ($LicDet.ActiveUnits -gt $LicDet.ConsumedUnits)
				    {
                        Write-host ""
					    write-host "Moving " $EmpNo "from Enterprise E3 License to Exchange Online P2 License" -ForegroundColor Yellow
		 
                        write-host "Removing Enterprise E4 License" -ForegroundColor Yellow
                        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:ENTERPRISEPACK"
                        write-host "Assigning Exchange Online Plan 2 License" -ForegroundColor Yellow
					    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicense "ul:EXCHANGEENTERPRISE"
				    }
				    else
				    {
					    Write-Host "There are No Exchange Online P2 Licenses available for assignment E3 License not Removed" -ForegroundColor Red
				    }

				    If (($HasBIPro -eq "True") -and ($ADUser.ExtensionAttribute1 -like "Ex-*"))
                    {
                        write-host "Removing PowerBI Pro License" -ForegroundColor Yellow
                        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:POWER_BI_PRO"
                    }
                    If (($HasBIFree -eq "True") -and ($ADUser.ExtensionAttribute1 -like "Ex-*"))
                    {
                        write-host "Removing PowerBI (Free) License" -ForegroundColor Yellow
                        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:POWER_BI_STANDARD"
                    }
                    if (($HasEMS -eq "True") -and ($ADUser.ExtensionAttribute1 -like "Ex-*"))
                    {
                        write-host "Removing Enterprise Mobility Suite/Intune (EMS) License" -ForegroundColor Yellow
                        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:EMS"
                    }
                    if (($HasCRM -eq "True") -and ($ADUser.ExtensionAttribute1 -like "Ex-*"))
                    {
                        write-host "Removing CRM/Dynamics License" -ForegroundColor Yellow
                        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:DYN365_ENTERPRISE_PLAN1"
                    }
                    if (($HasMRSS -eq "True") -and ($ADUser.ExtensionAttribute1 -like "Ex-*"))
                    {
                        write-host "Removing Microsoft Relationship Sales Solutin License" -ForegroundColor Yellow
                        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
                    }
                }
                else
                {
                    write-host "Exchange Online P2 licenses are only assigned to Terminated Staff or Service Accounts" -ForegroundColor Red
                }
				write-host ""
            }

            30
            {
				If ($HasCRM -eq "True")
				{
					Write-Host "`nRemoving CRM/Dynamics License from" $EmpNo -ForegroundColor Yellow
					Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:DYN365_ENTERPRISE_PLAN1"
        	      	$LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing Dynamics 365 Customer Engagement Plan License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host"`nEmployee" $EmpNo "Does Not have a Dynamics 365 Customer Engagement Plan License Assigned" -ForegroundColor Yellow
         	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove Dynamics 365 Customer Engagement Plan License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
               }
            }

            31
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
					Write-Host"`nEmployee" $EmpNo "Does Not have a PowerBI (Free) License Assigned" -ForegroundColor Yellow
          	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove PowerBI (Free) License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
               }
            }

            32
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
					Write-Host"`nEmployee" $EmpNo "Does Not have a Enterprise Mobility Suite/Intune (EMS) License Assigned" -ForegroundColor Yellow
         	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove Enterprise Mobility Suite/Intune (EMS) License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            33
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
					Write-Host"`nEmployee" $EmpNo "Does Not have a Microsoft Relationship Sales Solution License Assigned" -ForegroundColor Yellow
         	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove Microsoft Relationship Sales Solution License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            34
            {
               If ($HasATPDef -eq "True")
                {
			        Write-Host "`nRemoving Windows Defender Advanced Threat Protection License from" $EmpNo -ForegroundColor Yellow
				    Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:WIN_DEF_ATP"
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing Windows Defender Advanced Threat Protection License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host"`nEmployee" $EmpNo "Does Not have a Windows Defender Advanced Threat Protection Assigned" -ForegroundColor Yellow
         	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove Windows Defender Advanced Threat Protection License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            "34a"
            {
                if ($HasATPP1 -eq "True")
                {
                    write-host "`nRemoving Advanced Threat Protection Plan1 License from" $EmpNo -ForegroundColor Yellow
                    Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:ATP_ENTERPRISE"
                    $LineToWrite = $RecordEvent + "REVI" + "`t" + "Removing Advanced Threat Protection Plan1 License from " + $EmpNo
                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
                    Write-Host"`nEmployee" $EmpNo "Does Not have a Advanced Threat Protection Plan1 Assigned" -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove Advanced Threat Protection Plan1 License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            "34b"
            {
                if ($HasATPDefMAC -eq "True")
                {
                    write-host "`nRemoving Advanced Threat Protection for MAC License from" $EmpNo -ForegroundColor Yellow
                    Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:MDATP_XPLAT"
                    $LineToWrite = $RecordEvent + "REVI" + "`t" + "Removing Advanced Threat Protection for MAC License from " + $EmpNo
                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
                    Write-Host"`nEmployee" $EmpNo "Does Not have a Advanced Threat Protection for MAC Assigned" -ForegroundColor Yellow
        	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove Advanced Threat Protection for MAC License to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            35
            {
                If ($HasMeeting -eq "True")
                {
			        Write-Host "`nRemoving Meeting Room License from" $EmpNo -ForegroundColor Yellow
				    Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:MEETING_ROOM"
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing Meeting Room License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host"`nEmployee" $EmpNo "Does Not have a  Meeting Room License Assigned" -ForegroundColor Yellow
         	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove  Meeting Room License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            36
            {
                If ($HasPAppP2 -eq "True")
                {
			        Write-Host "`nRemoving PowerApps Plan 2 License from" $EmpNo -ForegroundColor Yellow
				    Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:POWERFLOW_P2"
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing PowerApps Plan 2 License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host"`nEmployee" $EmpNo "Does Not have a PowerApps Plan 2 License Assigned" -ForegroundColor Yellow
         	      	$LineToWrite = $RecordEvent + "REVI" + "`t" + "Cannot Remove PowerApps Plan 2 License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            37
            {
                If ($HasPhoneSys -eq "True")
                {
			        Write-Host "`nRemoving Phone System License from" $EmpNo -ForegroundColor Yellow
				    Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:MCOEV"
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing Phone System License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host"`nEmployee" $EmpNo "Does Not have a Phone System License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            38
            {
                If ($HasAudioConf -eq "True")
                {
			        Write-Host "`nRemoving Audio Conferencing License from" $EmpNo -ForegroundColor Yellow
				    Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:MCOMEETADV"
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing Audio Conferencing License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host"`nEmployee" $EmpNo "Does Not have a Audio Conferencing License"
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "No Audio Conferencing License assigned to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            39
            {
                If ($HasATPP1 -eq "True")
                {
			        Write-Host "`nRemoving Advanced Theat Protection Plan1 License from" $EmpNo -ForegroundColor Yellow
				    Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:ATP_ENTERPRISE"
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing Advanced Theat Protection Plan1 License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host"`nEmployee" $EmpNo "Does Not have a Advanced Theat Protection Plan1 License from " + $EmpNo
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "No Advanced Theat Protection Plan1 License assigned to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }

            40
            {
                If ($HasRmtAssist -eq "True")
                {
			        Write-Host "`nRemoving Microsoft Remote Assist License from" $EmpNo -ForegroundColor Yellow
				    Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:MICROSOFT_REMOTE_ASSIST"
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing Microsoft Remote Assist License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host "`nEmployee" $EmpNo "Does Not have a Microsoft Remote Assist License"
                    $LineToWrite = $RecordEvent + "DELE" + "`t" + "No Microsoft Remote Assist License assigned to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }

                If ($HasPhoneSys -eq "True")
                {
                    $Remove = "N"
                    write-host "Should the Phone System License be removed from this user (Y/N)? " -ForegroundColor Cyan -NoNewLine
                    $Remove = Read-Host
                    If ($Remove -eq "Y")
                    {
    	                Write-Host "`nRemoving Phone System License from" $EmpNo -ForegroundColor Yellow
			            Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:MCOEV"
       	                $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing Phone System License from " + $EmpNo
	   	                Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    }
                    else
                    {
					    Write-Host "`nEmployee" $EmpNo "Phone System License not removed from " + $EmpNo
       	                $LineToWrite = $RecordEvent + "DELE" + "`t" + "Phone System License not removed from " + $EmpNo
	      	            Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    }
                }
                else
                {
					Write-Host "`nEmployee" $EmpNo "Does Not have a Phone System License"
                    $LineToWrite = $RecordEvent + "DELE" + "`t" + "No Phone Phone System License assigned to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }

                If ($HasAudioConf -eq "True")
                {
                    $Remove = "N"
                    write-host "Should the Audio Conferencing License be removed from this user (Y/N)? " -ForegroundColor Cyan -NoNewLine
                    $Remove = Read-Host
                    If ($Remove -eq "Y")
                    {
    	                Write-Host "`nRemoving Audio Conferencing License from" $EmpNo -ForegroundColor Yellow
			            Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:MCOMEETADV"
       	                $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing Audio Conferencing License from " + $EmpNo
	   	                Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    }
                    else
                    {
					    Write-Host "`nEmployee" $EmpNo "Audio Conferencing License not removed from " + $EmpNo
       	                $LineToWrite = $RecordEvent + "DELE" + "`t" + "Audio Conferencing License not removed from " + $EmpNo
	      	            Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    }
                }
                else
                {
					Write-Host "`nEmployee" $EmpNo "Does Not have an Audio Conferencing License"
                    $LineToWrite = $RecordEvent + "DELE" + "`t" + "No Audio Conferencing License assigned to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }

                If ($HasInfoBarr -eq "True")
                {
                    $Remove = "N"
                    write-host "Should the Audio Conferencing License be removed from this user (Y/N)? " -ForegroundColor Cyan -NoNewLine
                    $Remove = Read-Host
                    If ($Remove -eq "Y")
                    {
                        Write-Host "`nRemoving E5 Insider Risk Management License from" $EmpNo -ForegroundColor Yellow
                        Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:M365_INSIDER_RISK_MANAGEMENT"
                        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing E5 Insider Risk Management License from " + $EmpNo
                        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    }
                    else
                    {
                        Write-Host "`nEmployee" $EmpNo "E5 Insider Risk Management License not removed from " + $EmpNo
                        $LineToWrite = $RecordEvent + "DELE" + "`t" + "E5 Insider Risk Management License not removed from " + $EmpNo
                        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    }
                }
                else
                {
					Write-Host "`nEmployee" $EmpNo "Does Not have a E5 Insider Risk Management License"
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "No E5 Insider Risk Management License assigned to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }                
            }

            41
            {
                If ($HasInfoBarr -eq "True")
                {
			        Write-Host "`nRemoving E5 Insider Risk Management License from" $EmpNo -ForegroundColor Yellow
				    Set-MsolUserLicense -UserPrincipalName $EmpNo -RemoveLicenses "ul:M365_INSIDER_RISK_MANAGEMENT"
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "Removing E5 Insider Risk Management License from " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
					Write-Host "`nEmployee" $EmpNo "Does Not have a E5 Insider Risk Management License"
        	        $LineToWrite = $RecordEvent + "DELE" + "`t" + "No E5 Insider Risk Management License assigned to " + $EmpNo
	      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }
            }  

            50
            {
                $sskid = "ul:ENTERPRISEPACK"
                EnabledE3Feature
            }

            51
           	{
                $SSKID = "ul:ENTERPRISEPACK"
                $LicType = "Enterprise E3"
                $DisPlan = $DisPlanEmpE3
                if ($ADUser.ExtensionAttribute1 -notlike "Employee*")
                {
                   $DisPlan = $Global:DisPlanNonEmpE3
                }

                If ($ADUser.ExtensionAttribute1 -like "Employee*")
                {
					write-host "Resetting " $LicType "licenses to standard employee licenses" -ForegroundColor Yellow				
				}
				else
				{
					write-host "Resetting " $LicType "licenses to standard non-employee licenses" -ForegroundColor Yellow						
				}

                EnabledE3Feature

      	      	$LineToWrite = $RecordEvent + "UPDA" + "`t" + "Resetting " + $LicType + " SubLicense Options to Standard Employee/Non-Employee for " + $EmpNo
      	        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				$DLO = ($DisPlan.Split(","))
				$MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
                Set-MsolUserLicense -UserPrincipalName $EmpNo -LicenseOptions $MyO365Sku
                StandardLicenses
						
    			RetentPolicy
            }

            52
            {
                $sskid = "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
                EnabledMRSSFeature
            }

            90
            {
                AvailLicense
            }

            99
            {
#                Do Nothing User Account/Mailbox does not exist
            }                										
    	}
    }

}While  ($Global:LicChg -ne 0)