#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Description  : Resets values Disabled for Users that are Rehired
#    'Project      : UL Office 365 Exchange
#    'Called By    : SDAP Menu
#    'Calls        : GetADUserInfo.ps1
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 11/3/2017
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 11/03/2017 Created from the Emergency Disable script
#    '             : 07/05/2018 Removed -ActiveSyncEnabled $true from the Set-CASMailbox command for staff in Americas
#    '             : 07/09/2018 Added the "EmpNo" into the Report file name
#    '             : 02/12/2021 Changed the URL to the NewHire sfbhelper line
#    '             : 04/08/2021 Modified the enablement of OWA to check to see if the user is Active
#    '             : 05/10/2021 Modified to make sure the new SIP values are set since Lync is no longer in user
#    '             : 07/13/2021 Modified the check if account is on legal hold not to overwrite ExtensionAttribute14
#    '             : 08/30/2021 Added code to disabled OOO message if enabled
#    '             : 09/28/2022 Added code to reset MaxReceiveSize to 35 MB
#    '             : 12/02/2022 Modifed code to assign E5 ilcense and not an E3 license
# ==========================================================================
#
#################################################################################

# Add Code to Remove access granted to others
function Check-LicenseAssignment
{
    $usr = get-msoluser -UserPrincipalName $EmpNo -ErrorAction SilentlyContinue
    $EmpGrp = [bool](Get-AzureADGroupMember -ObjectId 2aab81de-e54a-4c58-909c-b15273c54dff -Top 200000 |Where-Object {$_.ObjectID -eq $usr.ObjectID})
    $NonEmpGrp = [bool](Get-AzureADGroupMember -ObjectId 40dff561-964a-4f7b-9794-dba3542b63c4 -Top 200000 |Where-Object {$_.ObjectID -eq $usr.ObjectID})
    
#    $E5Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPREMIUM"}
    $E5Free = [bool] ($E5Lic.ConsumedUnits -lt $E5Lic.ActiveUnits)
	$HasExP2 = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:EXCHANGEENTERPRISE"})
    $HasE5 = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPREMIUM"})
	$HasPBIF = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:Power_BI_Standard"})

<#    $HasEMS = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:EMS"})
	$HasPBIF = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:Power_BI_Standard"})
    $HasATP = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:ATP_ENTERPRISE"})
    $HasPhone = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:MCOEV"})

    #Set SSKID to the E3 license and test to see if there are available licenses
   	$SSKID = "ul:ENTERPRISEPACK"
	$LicType = "Enterprise E3"
   	$DisPlanEmp = $Global:DisPlanEmpE3
    $DisPlanNonEmp = $Global:DisPlanNonEmpE3

    $E3Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}
    $EMSLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EMS"}

	If ($E3Lic.ConsumedUnits -ge $E3Lic.ActiveUnits)
	{
        $LineToWrite = $WhoAmI + "`t" + "INFO" + "`t" + "No E3 Licenses Available"
        WriteReportEvent
        WriteLogEvent
        $LineToWrite = "E3 Licenses Assigned " + (Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}).ConsumedUnits + " E3 License Allocation " +(Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}).ActiveUnits
        WriteReportEvent
        WriteLogEvent
    }
    else
    {
        If ($ADUser.ExtensionAttribute1 -like "Employee*")
        {
           write-host "Assigning standard License Set for individual who are Employees" -ForegroundColor Green
        }
        else
        {
            write-host "Assigning standard license set for individuals who are Non-Employees" -ForegroundColor Green
        }
    }

    Invoke-Expression -command E:\O365AdminShared\Scripts\O365DisabledLicenseFeatures.ps1
    If ($ADUser.ExtensionAttribute1 -like "Employee*")
    {
        $DLO = ($Global:DisPlanEmpE3.Split(“,”))
    }
    else
    {
        $DLO = ($Global:DisPlanNonEmpE3.Split(“,”))
    }
    
    If ($HasE5 -ne "True")
    {
        #Assign the license		
	    $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicenses $sskid –LicenseOptions $MyO365Sku
        $HasE5 = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"})
        If ($HasE5 -eq "True")
        {
            $LineToWrite = $WhoAmI + "`t" + "ASSG" + "`t" + "     Assigned " + $LicType + " License"
            WriteLogEvent
            WriteReportEvent
        }
        else
        {
            $LineToWrite = $WhoAmI + "`t" + "INFO" + "`t" + "     Unable to assign " + $LicType + " License"
	        WriteLogEvent
            WriteReportEvent
        }
    }
    else
    {
        write-host "E3 License Already assigned to this account resetting enabled features" -ForegroundColor Red
        $LineToWrite = $WhoAmI + "`t" + "INFO" + "`t" + "     " + $LicType + " License Already Assigned resetting enabled features "
		WriteLogEvent
#    	$DLO = ($DisPlan.Split(“,”))
		$MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $EmpNo –LicenseOptions $MyO365Sku
						
        RetentPolicy
    }
	
    if ($HasEMS -eq "True")
	{
        $LineToWrite = $WhoAmI + "`t" + "INFO" + "`t" + "     Enterprise Mobility Suite (EMS) License Already Assigned"
		WriteLogEvent
		WriteReportEvent	            
    }
	else
	{
	    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicenses "ul:EMS" -ErrorAction SilentlyContinue
        $HasEMS = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:EMS"})
        If ($HasEMS -eq "True")
        {
            $LineToWrite = $WhoAmI + "`t" + "ASSG" + "`t" + "     Assigned Enterprise Mobility Suite (EMS) License"
        }
        else
        {
    		$LineToWrite = $WhoAmI + "`t" + "ASSG" + "`t" + "     Unable to Assign Enterprise Mobility Suite (EMS) License"
            Write-Host "Unable to Assign Enterprise Mobility Suite (EMS) License" -ForegroundColor Red
        }
	    WriteLogEvent
	    WriteReportEvent		
    }

    if ($HasATP -eq "True")
    {
        $LineToWrite = $WhoAmI + "`t" + "INFO" + "`t" + "     Advanced Threat Protection Plan 1 (ATP) License Already Assigned"
	    WriteLogEvent
	    WriteReportEvent	
	}
	else
	{
	    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicenses "ul:ATP_ENTERPRISE" -ErrorAction SilentlyContinue
        $HasATP = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:ATP_ENTERPRISE"})
        If ($HasATP -eq "True")
        {
		    $LineToWrite = $WhoAmI + "`t" + "ASSG" + "`t" + "     Assigned Advanced Threat Protection Plan 1 (ATP) License"
        }
        else
        {
            $LineToWrite = $WhoAmI + "`t" + "ASSG" + "`t" + "     Unable to Advanced Threat Protection Plan 1 (ATP) License"
            Write-Host "Unable to Assign Advanced Threat Protection Plan 1 (ATP) License" -ForegroundColor Red
        }
	    WriteLogEvent
	    WriteReportEvent		
	}

    if ($HasPhone -eq "True")
	{
	    $LineToWrite = $WhoAmI + "`t" + "INFO" + "`t" + "     Phone System License Already Assigned"
	    WriteLogEvent
	    WriteReportEvent	
	}
	else
	{
	    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicenses "ul:MCOEV" -ErrorAction SilentlyContinue
        $HasPhone = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:MCOEV"})
        If ($HasPhone -eq "True")
        {
    	    $LineToWrite = $WhoAmI + "`t" + "INFO" + "`t" + "     Assigned Phone System License"
        }
        else
        {
    	    $LineToWrite = $WhoAmI + "`t" + "ASSG" + "`t" + "     Unable to Phone System License"
            Write-Host "Unable to Assign Phone system License" -ForegroundColor Red
        }
		WriteLogEvent
		WriteReportEvent
    }		
#>             

#	if ($HasE5 -eq "False")
#	{
		Assign-E5Lic
#	}
#	else
#	{
#		write-host "No availalbe E5 licenses" -foregroundcolor Red
#		$LineToWrite = $WhoAmI + "`t" + "INFO" + "`t" + "     No E5 licenses Available"	
#		WriteLogEvent	
#	}

    if ($HasPBIF -eq "True")
    {
	    $LineToWrite = $WhoAmI + "`t" + "INFO" + "`t" + "     PowerBI Free License Already Assigned"
        WriteLogEvent
        WriteReportEvent
    }
	else
	{
	    Set-MsolUserLicense -UserPrincipalName $EmpNo -AddLicenses "ul:POWER_BI_STANDARD" -erroraction SilentlyContinue
        $HasPBIF = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:Power_BI_Standard"})
        If ($HasPBIF -eq "True")
        {
            $LineToWrite = $WhoAmI + "`t" + "ASSG" + "`t" + "     Assigned PowerBI Free License"
        }
        else
        {
            $LineToWrite = $WhoAmI + "`t" + "ASSG" + "`t" + "     Unable to Assign PowerBI Free License"
            Write-Host "Unable to PowerBI Free License" -ForegroundColor Red
        }
		WriteLogEvent
		WriteReportEvent	            
    }
}

# Declare Drive | Folders | and Files
	$FileName		= "EmergencyRehire"
	$LogDrive		= "E:"
	$LogPath		= "\SDAP"
	$LogFolder		= "\" + $FileNAme
	$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
#	$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	$InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
	$ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	$LogFile        = $ReportFile
#################################################################################

$Global:strUserPath = ""
$Global:u = ""
$Global:inf = ""
$Global:MBXAccess = ""
$strDN = 0
$Cont = "N"

write-host "Starting Rehire Script " -ForegroundColor Magenta
write-host
write-host
write-Host "Enter Employee Number of Rehired Staff : " -ForegroundColor Cyan -NoNewline
$ENo = Read-Host
$EmpNo = $ENo + "@global.ul.com"
$ADU = get-ADUser -Filter {SamAccountName -eq $ENo} -ErrorAction SilentlyContinue
$ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-EmpNo" + $ENo + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"

# Setup Folders and Files	
	CheckLogFiles
	$LineToWrite = "STAR" + "`t" + $FileName + " script has started"
	WriteLogEvent
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	WriteLogEvent
    WriteReportEvent

if ($adu -ne $null)
{
    write-Host "Enter the Request Number : " -ForegroundColor Cyan -NoNewline
    $RequestNo = Read-Host
    $DescriptionComment = "Rehire Request - " + $RequestNo

	$LineToWrite = $WhoAmI + "`t" + "INFO" + "`t" + "Rehire for: " + $ENo
    WriteReportEvent
	$LineToWrite = $WhoAmI + "`t" + "INFO" + "`t" + "Request Number: " + $RequestNo
    WriteReportEvent

    $Usr = Get-MsolUser -UserPrincipalName $EmpNo -ErrorAction SilentlyContinue
    $CASMbx = Get-CASMailbox $EmpNo -ErrorAction SilentlyContinue

#    Check-LicenseAssignment
#    $UMMbx = Get-UMMailbox $EmpNo -ErrorAction SilentlyContinue

    invoke-expression -Command .\GetADUserInfo.ps1

    Write-Host "Do you want to continue with re-enabling this account (Y/N)? " -ForegroundColor Cyan -NoNewline
    $Cont = Read-Host

    If ($Cont -eq "Y")
    {
        write-host""
        write-host""
        write-host "Re-enabling AD Account..." -ForegroundColor Cyan
        Enable-ADAccount -Identity $ENo

        write-host "Resetting AD Description to" $DescriptionComment "..." -ForegroundColor Green
	    $Global:u.Description.value = $ENo
        $Global:u.CommitChanges()	

        $Pwd = Get-Random
        write-host "Resetting AD Password to" $Pwd "..." -ForegroundColor Green
        Set-ADAccountPassword -Identity $ENo -NewPassword (ConvertTo-SecureString -AsPlainText [string]$PWD -Force)

        if (($Global:MbxAccess -ne $null) -or ($Global:u:ExtenstionAttribute2 -ne "T"))
        {
            write-host "Enabling O365 Mailbox..." -ForegroundColor Green
            Set-Mailbox $ENo -AccountDisabled:$False

            write-host "Checking Stanard License Set Assigned...." -ForegroundColor Green
            Set-MsolUser -UserPrincipalName $EmpNo -UsageLocation US -ErrorAction Silentlycontinue
            $HasExP2 = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:EXCHANGEENTERPRISE"})
            $HasE5 = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPREMIUM"})
            $HasEMS = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:EMS"})
	        $HasPBIFree = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:Power_BI_Standard"})
            $HasPhone = [bool]((Get-MsolUser -UserPrincipalName $EmpNo).Licenses |Where-Object {$_.AccountSkuID -eq "ul:MCOEV"})

            Assign-E5Lic
#            write-host "Setting Mailbox Retention Policy...." -ForegroundColor Green
#            RetentPolicy
#            StandardLicenses
        
            write-host "Enabling Mailbox Protocols for OWA, OWAForDevices, ActiveSync and AllowMAC's..." -ForegroundColor Green

            if (($Global:u.ExtensionAttribute5 -ne "Europe") -and ($Global:u.ExtensionAttribute5 -ne "Asia Pacific"))
            {
                Set-CASMailbox -Identity $ENo -OwaEnabled $true -OwaForDevicesEnabled $true -ActiveSyncEnabled $true -EwsAllowMacOutlook $true
            }
            else
            {
                Set-CASMailbox -Identity $ENo -OwaEnabled $true -OwaForDevicesEnabled $true -EwsAllowMacOutlook $true
            }

            If ($ADUser.msExchHideFromAddressLists -eq "True")
            {
                write-host "Setting user to be Hidden from the Address Book" -ForegroundColor Red
                write-host
                $strUserPath = [string]::format("LDAP://{0}", $ADUser.DistinguishedName)
                $u = new-object System.DirectoryServices.DirectoryEntry($strUserPath) -ErrorAction SilentlyContinue
                $u.msExchHideFromAddressLists.value = $False
                $U.CommitChanges()
            }

            $mbx = get-mailbox $EmpNo
            If ($mbx.MaxReceiveSize -like "0*")
            {
                write-host "Resetting mailbox to allow receipt of new messages" -ForegroundColor Green
                set-mailbox $EmpNo -MaxReceiveSize 35MB
                $LineToWrite = $RecordEvent + "SET " + "`t" + "Resetting MaxReceiveSize to 35MB."
                WriteReportEvent
            }

            $IntuneMem = "off"
            $ADGroupMem = Get-ADPrincipalGroupMembership $ENo -ResourceContextServer global.ul.com
            foreach ($ADGRoupMem in $ADGroupmem)
            {
                if ($ADGroupMem.SamAccountName -like "*IntuneEnfor*")
                {
                    $IntuneMem = "1"
                    $IntuneGroup = "Previously Assigned"
                }
            }
            If ($IntuneMem -eq "off")
            {
                If ($ADUser.ExtensionAttribute5 -eq "Europe")
                {
                    Add-ADGroupMember -Identity "ACL.EU.IntuneEnforced" -Members $ENo -Confirm:$False
                    write-host  "Adding to the ACL.EU.IntuneEnforced group" -ForegroundColor Green
                    $LineToWrite = $WhoAmI + "`t" + "ADD " + "`t" + "     Adding to the ACL.EU.IntuneEnforced group"
                    WriteLogEvent
		            WriteReportEvent
                    $IntuneGroup = "ACL.EU.IntuneEnforced"
                }
                elseif (($ADUser.ExtensionAttribute5 -eq "Asia Pacific") -or ($ADUser.ExtensionAttribute5 -eq "APAC & MEA"))
                {
                    Add-ADGroupMember -Identity "ACL.AP.IntuneEnforced" -Members $ENo -Confirm:$False
                    write-host  "Adding to the ACL.AP.IntuneEnforced group" -ForegroundColor Green
                    $LineToWrite = $WhoAmI + "`t" + "ADD " + "`t" + "     Adding to the ACL.AP.IntuneEnforced group"
		            WriteLogEvent
		            WriteReportEvent
                    $IntuneGroup = "ACL.AP.IntuneEnforced"
                }
                else
                {
                    Add-ADGroupMember -Identity "ACL.UL.IntuneEnforced" -Members $ENo -Confirm:$False
                    write-host  "Adding to the ACL.UL.IntuneEnforced group" -ForegroundColor Green
                    $LineToWrite = $WhoAmI + "`t" + "ADD " + "`t" + "     Adding to the ACL.UL.IntuneEnforced group"
                    WriteLogEvent
                    WriteReportEvent
                    $IntuneGroup = "ACL.UL.IntuneEnforced"
                }
            }
            else
	        {
			    $LineToWrite = $WhoAmI + "`t" + "INFO" + "`t" + "     Already a member an IntuneEnforced group"
		        WriteLogEvent
				WriteReportEvent				
			}

            write-host "Checking to see if this individual is already a member of the MFA_Enabled Group (this may take a few minutes)" -ForegroundColor Cyan
            $MFAMem = [bool](get-AzureADGroupMember -ObjectId abf2ba7d-50e2-4ca6-99c8-2e9c245cfed1 -all $true |Where-Object {$_.UserPrincipalName -eq $EmpNo})
            If ($MFAMem -eq $False)
            {
                write-host "Adding user to the MFA_Enabled Group" -ForegroundColor Green
                Add-AzureADGroupMember -ObjectId abf2ba7d-50e2-4ca6-99c8-2e9c245cfed1 -RefObjectId ((get-msoluser -UserPrincipalName $EmpNo).objectid)
                $LineToWrite = $WhoAmI + "`t" + "ADDMEM" + "`t" + "     Added to MFA_Enabled Group"
        	    WriteReportEvent
            }
            else
            {
                write-host "Already a member of the MFA_Enabled Group" -ForegroundColor Red
                $LineToWrite = $WhoAmI + "`t" + "NOCHG" + "`t" + "     Already MFA_Enabled Group Member"
        	    WriteReportEvent
            }

            write-host "Disabling RemotePowerShell" -ForegroundColor Green
            Set-User $ENo -RemotePowerShellEnabled $False -Confir:$False
            $LineToWrite = $WhoAmI + "`t" + "DISA" + "`t" + "     Disabled Remote PowerShell"
            WriteReportEvent

            write-host "Setting SIP Addresses" -ForegroundColor Green
            Set-ADUser $ENo -Replace @{'msRTCSIP-PrimaryUserAddress'= ("sip:" + $CASMbx.PrimarySMTPAddress)}
            Set-ADUser $ENo -Replace @{'msRTCSIP-DeploymentLocator'="sipfed.online.lync.com"}
            Set-ADUser $ENo -Replace @{'msRTCSIP-FederationEnabled'="TRUE"}
            Set-ADUser $ENo -Replace @{'msRTCSIP-InternetAccessEnabled'="TRUE"}
            Set-ADUser $ENo -Replace @{'msRTCSIP-UserEnabled'="TRUE"}
        }
        else
        {
            write-host "No O365 Mailbox is configured for this account or this user is still marked as a Terminated user..." -ForegroundColor Red
        }

        $UsrVal = $Global:strUserPath.Trim("LDAP://")
        $Global:UsrDetails = Get-ADObject -Identity  $UsrVal -Properties ProtectedFromAccidentalDeletion
        Write-Host "What is the 3 letter site code for this user?" -ForegroundColor Cyan -NoNewline
        $sitecode = Read-Host
        $NewLoc = ""
        switch ($Global:u.ExtensionAttribute5)
        {
            "United States"
            {
                $NewLoc = "OU=Users,OU=" + $SiteCode + "Win7,OU=NA,DC=global,DC=ul,DC=com"
            }
            "Canada"
            {
                $NewLoc = "OU=Users,OU=" + $SiteCode + "Win7,OU=CA,DC=global,DC=ul,DC=com"
            }
            "Asia Pacific"
            {
                $NewLoc = "OU=Users,OU=" + $SiteCode + "Win7,OU=ASIA,DC=global,DC=ul,DC=com"
            }
            "Europe"
            {
                $NewLoc = "OU=Users,OU=" + $SiteCode + "Win7,OU=EULA,DC=global,DC=ul,DC=com"
            }
            "Latin America"
            {
                $NewLoc = "OU=Users,OU=" + $SiteCode + "Win7,OU=EULA,DC=global,DC=ul,DC=com"
            }
        }

        if ($NewLoc -notlike "*users*")
        {
            write-host "No Region defined for this user please enter 2 letter code for region (NA/CA/AP/EU/LA):" -ForegroundColor Cyan -NoNewline
            $Region = Read-Host
            $NewLoc = "OU=Users,OU=" + $SiteCode + "Win7,OU=" + $Region + ",DC=global,DC=ul,DC=com"
        }

        if ($UsrDetails.ProtectedFromAccidentalDeletion -ne $True)
        {
            write-host "Moving AD Object to the " $SiteCode " Users OU..." -ForegroundColor Green
            Move-ADObject ($Global:strUserPath -replace("LDAP://","")) $NewLoc
        }
        else
        {
            write-host "The AD Account is protected and cannot be moved an email is being sent to the Enterprise Email Team..." -ForegroundColor Red
            $server = "smtp-relay.ul.com"
            $client = new-object system.net.mail.smtpclient $server
            $from = New-Object System.Net.Mail.MailAddress "DoNotReply@ul.com" , "Account Provisioning Team"
            $to = $from 
            $message = new-object  System.Net.Mail.MailMessage $from, $to 
            $message.IsBodyHtml = $true
            $msgfont = "<basefont face=verdana size=2.5 color=black>"
    
            $DGOwnerID = "EnterpriseMessagingServices@ul.com"
            $MsgSubject = "Emergency ReHire - Move Protected AD Account to Active OU"
            $MsgBody = "The Active Directory Account for " + $ENo + " is current protected an needs to be moved to the " + $NewLoc + " OU"
            $message.CC.Clear()
            $message.To.Add($DGOwnerID)
            $message.Subject = $MsgSubject
            $message.Body = $MsgBody
            $client.Send($message)
        }

        if ($Global:u.msExchHideFromAddressLists.value -eq $True)
        {
            write-host "Unhiding User from the Address Book..." -ForegroundColor Green
            $Global:u.msExchHideFromAddressLists.value = $False
            $Global:u.CommitChanges()	
        }
        else
        {
            write-host "User Not Hidden from the Address Book..." -ForegroundColor Red
        }

        $MbxOOO = Get-MailboxAutoReplyConfiguration $ENo
        If ($MbxOOO.AutoReplyState -ne "Disabled")
        {
            write-host "Disabling OOO message..." -ForegroundColor Green
            Set-MailboxAutoReplyConfiguration $ENo –AutoReplyState Disabled
            $LineToWrite = $WhoAmI + "`t" + "         Disabled OOO Message:  #####################################"
            WriteReportEvent
        }

        If ($Global:inf.LitigationHoldEnabled -eq $false)
#        if (($Global:u.ExtensionAttribute14.value -notlike "*Preservation*") -or ($Global:u.ExtensionAttribute14.value -notlike "*Legal*"))
        {
            write-host "Updating ExtensionAttribute14 to Emergency Rehire..." -ForegroundColor Green
            $Global:u.ExtensionAttribute14.value = $DescriptionComment
            $Global:u.CommitChanges()					
        }
        else
        {
            write-host "ExtensionAttribute14 already contains a value for legal hold not overwriting...." -ForegroundColor Red
        }

        if ($Global:u.userCertificate.value -eq $null)
        {
            write-host "There are no userCertificates saved for this account...." -ForegroundColor Green
        }


        Write-host "Rehire Process Complete..." -ForegroundColor Cyan
        Write-Host ""
        write-host ""
        $LineToWrite = $WhoAmI + "`t" + "Rehire Process Completed!!"
        WriteReportEvent
    }
    else
    {
        Write-host "Disablement Process Aborted..." -ForegroundColor Red
        $LineToWrite = $WhoAmI + "`t" + "Disablement Process Aborted!!"
        WriteReportEvent
    }
}
else
{
    write-host "There is no account for" $Eno
    $LineToWrite = $WhoAmI + "`t" + "There is No Account for Employee Number " + $ENo
    WriteReportEvent
}