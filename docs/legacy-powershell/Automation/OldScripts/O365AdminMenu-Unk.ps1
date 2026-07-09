<####################################################################################
#
#  This script contains all the common functions used by the O365 Team
#
#  Created: 5/13/2020 - S.Glazebrook
#  06/11/2020 - SAG - Modifed the Common Data Service and Bookings features in the E3EnabledFeatures
#
####################################################################################>

#  Checks the that LogFile directory exists for the given menu item
function CheckLogFiles
{
	# Create Files folder $LogDirectory if it's not present
	if (Test-Path $LogDirectory)
		{
		# the directory is present
		}
	else
		{ 
		mkdir $LogDirectory
		}
} #end CheckLogFiles

#Default form buttons
function Add-FormStandardButtons
{
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    $Global:buttonPanel = New-Object Windows.Forms.Panel  
    $buttonPanel.Size = New-Object Drawing.Size @(400,40) 
    $buttonPanel.Dock = "Bottom"    
    $Global:cancelButton = New-Object Windows.Forms.Button  
        $Global:cancelButton.Top = $buttonPanel.Height - $Global:cancelButton.Height - 10; $Global:cancelButton.Left = $buttonPanel.Width - $Global:cancelButton.Width - 10 
        $Global:cancelButton.Text = "Cancel" 
        $Global:cancelButton.DialogResult = "Cancel" 
        $Global:cancelButton.Anchor = "Right"
    ## Create the OK button, which will anchor to the left of Cancel 
    $Global:okButton = New-Object Windows.Forms.Button   
        $Global:okButton.Top = $cancelButton.Top ; $Global:okButton.Left = $cancelButton.Left - $Global:okButton.Width - 10
        $Global:okButton.Text = $Action
        If ($BldDetails -eq "N")
        {
            $Global:okButton.Text = "Continue"
        }
        $Global:okButton.DialogResult = "OK" 
        $Global:okButton.Anchor = "Right"
    ## Add the buttons to the button panel 
    $Global:buttonPanel.Controls.Add($Global:okButton) 
    $Global:buttonPanel.Controls.Add($Global:cancelButton) 
    ## Add the button panel to the form 
    $Global:form.Controls.Add($buttonPanel) 
    ## Set Default actions for the buttons 
    $Global:form.AcceptButton = $Global:okButton          # ENTER = ok 
    $Global:form.CancelButton = $Global:cancelButton      # ESCAPE = Cancel
}

function Assign-ATPDefLic
{
    if ($HasATPDef -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Windows Defender Advanced Threat Protection license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Windows Defender Advanced Threat Protection License Already Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Assigning Windows Defender Advanced Threat Protection License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:WIN_DEF_ATP"
    }
    WriteLogEvent
}

function Assign-ATPPDefMACLic
{
    if ($HasATPDefMAC -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Advanced Threat Protection for MAC license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Advanced Threat Protection for MAC License Already Assigned to " + $Global:UPN
    }
    else
    {
        $LineToWrite = "REVI" + "`t" + "Assigning Advanced Threat Protection for MAC License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MDATP_XPLAT"
    }
    WriteLogEvent
}

function Assign-ATPP1Lic
{
    if ($HasATPP1 -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Advanced Threat Protection Plan1 license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Advanced Threat Protection Plan1 License Already Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning Advanced Threat Protection Plan1 License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:ATP_ENTERPRISE"
    }
    WriteLogEvent
}

Function Assign-AudioConferencingLic
{
    if ($HasAudioConf -eq "True")
    {
        $Output = $wshell.Popup("This individual already has an Audio Conferencing license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Audio Conferencing License Already Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning Audio Conferencing License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MCOMEETADV"
    }
    WriteLogEvent
}

Function Assign-CommonAreaPhoneLic
{
    if ($HasCAP -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Common Area Phone license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Common Area Phone License Already Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning Common Area Phone License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MCOCAP"
    }
    WriteLogEvent
}
    
function Assign-CRMLic
{
    If ($HasCRM -eq "True")
    {
      	$Output = $wshell.Popup("This individual already has a Microsoft Phone System license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Dynamics 365 Customer Engagement Plan License Already Assigned to " + $Global:UPN
    }
	else
	{
        If ($HasMRSS -eq "True")
        {
            Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
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

        $DLO = ($DisPlan.Split(","))
        $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense $SSKID -LicenseOptions $MyO365Sku
        $LineToWrite = "REVI" + "`t" + "Assigning Dynamics 365 Customer Engagement Plan Lse to " + $Global:UPN
    }
    WriteLogEvent
}

function Assign-E3Lic
{
    If ($HasE3 -eq "True")
    {
        $SSKID = "ul:ENTERPRISEPACK"
        $LicType = "Enterprise E3"
        $DisPlan = $Global:DisPlanEmpE3
        if ($Global:ADUser.ExtensionAttribute1 -notlike "Employee*")
        {
            $DisPlan = $Global:DisPlanNonEmpE3
        }

        If ($Global:ADUser.ExtensionAttribute1 -like "Employee*")
        {
		    $Output = $wshell.Popup("Resetting " + $LicType + " licenses to standard employee licenses.",0,"Reset Options",0+32)
    	}
	    else
        {
	        $Output = $wshell.Popup("Resetting " + $LicType + " licenses to standard non-employee licenses.",0,"Reset Options",0+32)
		}

        EnabledE3Feature

        $LineToWrite = "UPDA" + "`t" + "Resetting " + $LicType + " SubLicense Options to Standard Employee/Non-Employee for " + $Global:UPN
        WriteLogEvent
		$DLO = ($DisPlan.Split(","))
		$MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -LicenseOptions $MyO365Sku

        RetentPolicy
    }
    else
    {
        AssignE3License
    }

    If ($HasExP2 -eq "True")
    {
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:EXCHANGEENTERPRISE"
        $Output = $wshell.Popup("Remvoing Exchange Online P2 license and reassigning to a " + $LiceType + " license",0,"Removing License",0+32)
        $LineToWrite = "DELE" + "`t" + "Removing P2 License from " + $Global:UPN
        WriteLogEvent
    }

    write-host "Checking the Retention Policy Configuration and that the PowerBI and EMS licenses are assigned"
    StandardLicenses
}

function Assign-EMSLic
{
    If ($HasEMS -eq "True")
	{
        $Output = $wshell.Popup("This individual already has a EMS license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Enterprise Mobility Suite/Intune (EMS) License Already Assigned to " + $Global:UPN
    }
	else
    {
		Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:EMS"
        $LineToWrite = "REVI" + "`t" + "Assigning Enterprise Mobility Suite/Intune (EMS) License to " + $Global:UPN
    }
    WriteLogEvent
}

function Assign-MeetingLic
{
    if ($HasMeeting -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Meeting Room license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Meeting Room License Already Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Assigning Meeting Room License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MEETING_ROOM"
    }
    WriteLogEvent
}

Function Assign-MRSSLic
{
    if ($HasMRSS -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Microsoft Relationship Sales Solution license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Microsoft Relationship Sales Solution License Already Assigned to " + $EmpNo
        WriteLogEvent
    }
    else
    {
        Write-Host "`nAssigning Microsoft Relationship Sales Solution License to" $EmpNo -ForegroundColor Yellow
        $LineToWrite = "REVI" + "`t" + "Assigning Microsoft Relationship Sales Solution License to " + $EmpNo
	    WriteLogEvent 
     
        $sskid = "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
        $LicType = "Microsoft Relationship Sales Solution"
        $DisPlan = $Global:DisMRSSPlan

	    $LicDet = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq $SSKID}
			
	    if ($LicDet.ActiveUnits -eq $LicDet.ConsumedUnits)
	    {					
            write-host "`nNo " $LicType "Licenses Available to Assign"
            $LineToWrite = "DELE" + "`t" + "No Microsoft Relationship Sales Solution Licenses Available for Assignment to " + $Global:UPN
	        WriteLogEvent
        }
	    else
	    {

            If ($HasCRM -eq "True")
            {
                write-host "Removing Dynamics 365 Customer Engagement Plan License from this account"
                Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:DYN365_ENTERPRISE_PLAN1"
            }
		    write-host "`nAssigning " $EmpNo "a " $LicType "License" -ForegroundColor Yellow
            $LineToWrite = "UPDA" + "`t" + "Assigning " + $LicType + "to " + $Global:UPN
	        WriteLogEvent
	    }

        $DLO = ($DisPlan.Split(","))
        $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense $SSKID -LicenseOptions $MyO365Sku
    }                       
}

function Assign-P2Lic
{
    $Lic = $Global:UserLicense = (Get-MsolUser -UserPrincipalName 96151@global.ul.com).Licenses
    if ($HasExP2 -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Exchange Online P2 license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Exchange Online P2 License Already Assigned to " + $Global:UPN
    }
    else
    {
        $LineToWrite = "REVI" + "`t" + "Assigning Exchange Online P2 License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:EXCHANGEENTERPRISE"
    }
    WriteLogEvent

    If ($Lic.count -ge 1)
    {
        $Output = $wshell.Popup("This individual has " + $Global:LicAssigned + " licenses assigned do you want to remove all assigned licenses from this individual?",4,"Remove Standrd Licenses",4+32)
        If ($Output -eq 6)
        {
            Foreach ($Lic in $Lic)
            {
                Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense $Lic.AccountSkuID
                $LineToWrite = "REVI"  + "`t" + "Removing " + $Lic.AccountSkuID + " from " + $Global:UPN
                WriteLogEvent
            }
        }
        else
        {
            (If $HasE3 -eq "True")
            {
                $LineToWrite = "REVI"  + "`t" + "Removing EnterpriseE3 License from " + $Global:UPN
                WriteLogEvent
                Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:ENTERPRISEPACK"
            }
        }
    }
}

Function Assign-PAppsP2Lic
{
    if ($HasPAppsP2 -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a PowerApps P2 license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "PowerApps Plan 2 License Already Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Assigning PowerApps Plan 2 License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:POWERFLOW_P2"
    }
    WriteLogEvent
}

Function Assign-PhoneSystemLic
{
    if ($HasPhoneSys -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Microsoft Phone System license assigned.",0,"License Already Assigned",0+32)
        write-host "`nThis individual already has a Phone System License Assigned"
        $LineToWrite = "REVI" + "`t" + "Phone System License Already Assigned to " + $Global:UPN
    }
    else
    {
     	$LineToWrite = "REVI" + "`t" + "Assigning Phone System License to " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MCOEV"
    }
    WriteLogEvent
}

function Assign-PowerBIProLic
{
    if ($HasBIPro -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a PowerBI Pro license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "PowerBI Pro License Already Assigned to " + $Global:UPN
        WriteLogEvent
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning PowerBI Pro License to " + $Global:UPN
        WriteLogEvent
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:POWER_BI_PRO"

        If ($HasBIFree -eq "True")
        {
        	$LineToWrite = "REVI" + "`t" + "Removing PowerBI (free) License from" + $Global:UPN
	      	WriteLogEvent
            Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:POWER_BI_STANDARD"
        }
    }  
}

function Assign-RemoteAsstLic
{
    if ($HasRmtAssist -eq "True")
    {
        $Output = $wshell.Popup("This individual already has a Microsoft Remote Assist license assigned.",0,"License Already Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "Microsoft Remote Assistg License Already Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Assigning Microsoft Remote Assist License to " + $Global:UPN
        If ((get-msoluser -UserPrincipalName $global:UPN).UsageLocation -eq $null)
        {
            Set-MsolUser -UserPrincipalName $Global:UPN -UsageLocation "US" -ErrorAction Silentlycontinue
        }
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:MICROSOFT_REMOTE_ASSIST"
    }
    WriteLogEvent
    If ($ADuser.ExtensionAttribute1 -like "Employee*")
    {
        Assign-AudioConferencingLic
        Assign-PhoneSystemLic
    }
}

#This function creates New UserGroups
#   Used by NewSharedMailboxScripts
function CreateUSG
{				
    $DstExists = [bool](Get-DistributionGroup $GrpAddr -ResultSize Unlimited -ErrorAction SilentlyContinue)
    if ($DstExists -eq $false)
    {
	    if ($GrpAddr.contains("@"))
        {
            $atMail = $GrpAddr.indexOf("@")
		    $DGAlias = $GrpAddr.substring(0,$atMail)
		    $DGManagedByMembers = ($Global:txtGrpOwnr.Text.Split(",")).Trim()
            #$DGManagedByMembers = ($Global:txtEDGrpMbr.Text.Split(",")).Trim()

            $LineToWrite = ""
            WriteReportEvent
        			
		    # Create the USG
		    New-DistributionGroup -Name $GrpName `
		        -PrimarySmtpAddress $GrpAddr `
			    -Alias $DGAlias `
			    -ManagedBy $DGManagedByMembers `
			    -RequireSenderAuthenticationEnabled $TRUE `
			    -Type Security | Out-Null
					
            $Owners = "Owners: " + ((Get-DistributionGroup $GrpName).ManagedBy -join (", "))
        
            Set-Group -identity $GrpName `
		        -Notes ($Owners + " - Per: " + $Global:txtInpTaskNo.Text)

            If ($Owners.Length -gt 175)
            {
                Set-DistributionGroup $GrpName -MailTip $Owners.Substring(0,175)
            }
            else
            {
                Set-DistributionGroup $GrpName -MailTip $Owners
            }
        
            if (Get-DistributionGroup $GrpAddr)
            {
		        write-host ""
                write-host "USG created: " $GrpName " (" $GrpAddr ")" -ForegroundColor Cyan
		        $LineToWrite = "CREATE" + "`t" + $GrpName + "`t" + $GrpAddr + " USG Created"
		        WriteReportEvent
			
			    write-host "USG Ownner: " $Global:txtGrpOwnr.Text -ForegroundColor Cyan
		        $LineToWrite = "OWNER" + "`t" + $Global:txtGrpOwnr.Text + " USG Owner"
		        WriteReportEvent
				
		        $USG	= Get-DistributionGroup $GrpAddr
		        Set-DistributionGroup $GrpAddr `
			        -RequireSenderAuthenticationEnabled $True `
		 		    -BypassSecurityGroupManagerCheck `
				    -CustomAttribute15 ($FileName + " PS Date: " + (get-date -uformat %D) + " PS Time: " + (get-date -uformat %T))
				
		       # Add USG members
                write-host ("Adding Members to USG: " + $GrpName + " (" + $GrpAddr + ")")
		        $LineToWrite = "MEMBER" + "`t" + "Adding Members to Distribution Group " + $GrpName
		        WriteReportEvent
		
                $addMember = $GrpMem.split(",")
		        if ($? -eq $true)
                {
			        ForEach ($member in $addMember)
                    {
                        $member = $member.Trim()
				        if (Get-Mailbox $member)
                        {
						    Add-DistributionGroupMember $GrpAddr -Member $member -BypassSecurityGroupManagerCheck
						    write-host "Added member:" $member
						    $LineToWrite = "PASS" + "`t" + "Added " + $member
					    }
					    else
                        {
						    write-host "ERROR finding member: " $member
						    $LineToWrite = "FAIL" + "`t" + "Failed adding " + $member
					    }
				    }
                    WriteReportEvent
		        }	
		        else
                {
			        write-host "No members to add"
			        $LineToWrite = "FAIL" + "`t" + "No Members to Add"
                    WriteReportEvent
		        }
            }	
		    else
            {
		        Write-Host "USG not created: " $GrpName " (" $GrpAddr ")" -ForegroundColor Red
			    $LineToWrite = "FAIL" + "`t" + $GrpName + "`t" + $GrpAddr + "`t" + "USG not created"
			    WriteReportEvent
	        }
        }
    }
    else
    {
        Write-Host "User Security Group Already Exists - No Changes Made" -ForegroundColor Red
        pause
    }
}

Function EnabledE3Feature
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
    $SubLic = ($UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq $SSKID}).ServiceStatus
	foreach ($SubLic in $SubLic)
	{
        $AssignWho = "Enabled All Types"

        switch ($subLic.ServicePlan.ServiceName)
        {
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

# This function connects to Active Directory and gets the record for the user 
Function GetUserDN($strUID)
{
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = "LDAP://DC=global,DC=ul,DC=com"
    $objSearcher.Filter = [string]::format("(sAMAccountName={0})",$strUID)
    $ux = $null
    $ux = $objSearcher.FindOne()
    
    if ($null -eq $ux)
    {
        return $null
    } 
    else
    {
        return $ux.Properties.distinguishedname
    }
}

# Gets AD ACcount Properties
Function GetAcctInfo($ENo)
{
    $Global:ADCmt = ""
    $ADExists = [bool]($strDN = GetUserDN $ENo -ErrorAction SilentlyContinue)
    If ($ADExists -eq "True")
    {
       $Global:UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion
     }

    $MbxExists = [bool](get-mailbox -identity $ENo -ErrorAction SilentlyContinue)
    If ($MbxExists -eq "True")
    {
        $Global:inf = get-mailbox $ENo
        If ([bool](get-ADUser -Filter {SamAccountName -eq $ENo} -ErrorAction SilentlyContinue))
        {
            $Global:strUserPath = [string]::format("LDAP://{0}", $strDN)
            $Global:u = new-object System.DirectoryServices.DirectoryEntry($Global:strUserPath)
            $Global:ADCmt = $Global:u.ExtensionAttribute14.value
        }
        else
        {
            $Global:ADCmt = "*****No Active Directory Account For This User*****"
            $Global:strUserPath = "No Active Directory Account for this User"            
        }

        If ($Global:FormRefresh -ne "Y")
        {
            $LineToWrite = $WhoAmI + "`t" + "DisplayName                    :" + "`t" + $Global:inf.DisplayName
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "Employee Type                  :" + "`t" + $Global:inf.CustomAttribute1
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "AD Container                   :" + "`t" + $Global:strUserPath
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "LitigationHoldEnabled          :" + "`t" + $Global:inf.LitigationHoldEnabled
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "LitigationHoldDate             :" + "`t" + $Global:inf.LitigationHoldDate
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "LigitationHoldOwner            :" + "`t" + $Global:inf.LitigationHoldOwner
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "Active InPlace Holds           :" + "`t" + ($Global:inf.InPlaceHolds -join ",")
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "Hidden from Address Book O365  :" + "`t" + $Global:inf.HiddenFromAddressListsEnabled
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "Hidden from Address Book AD    :" + "`t" + $Global:u.msExchHideFromAddressLists.value
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "O365 Retention Comment         :" + "`t" + $Global:inf.RetentionComment
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "O365 Retention Comment         :" + "`t" + $Global:inf.ExtensionAttribute14
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "AD Retention Comment           :" + "`t" + $Global:ADCmt
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "AD Object Protected            :" + "`t" + $Global:UsrDetails.ProtectedFromAccidentalDeletion
            WriteReportEvent
        }

        PrtMBxAccess($Global:MbxAccess)

        $LineToWrite = "`n"
        WriteReportEvent
    }
    else
    {
        write-host""
        write-host "Mailbox does not exist for Emp#" $ENo -ForegroundColor Red
    }
}

Function O365HasLicenses
{
    $Global:HasBIFree = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:POWER_BI_STANDARD"})
    $Global:HasBIPro = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:POWER_BI_PRO"})
    $Global:HasCRM = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:DYN365_ENTERPRISE_PLAN1"})
    $Global:HasCAP = [bool]($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:MCOCAP"})
    $Global:HasEMS = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:EMS"})
    $Global:HasExP2 = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:EXCHANGEENTERPRISE"})
    $Global:HasE3 = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:ENTERPRISEPACK"})
    $Global:HasMRSS = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"})
    $Global:HasFlowP2 = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuId -eq "ul:FLOW_P2"})
    $Global:HasATPDef = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:WIN_DEF_ATP"})
    $Global:HasATPDefMAC = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MDATP_XPLAT"})
    $Global:HasATPP1 = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:ATP_ENTERPRISE"})
    $Global:HasMeeting = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MEETING_ROOM"})
    $Global:HasPAppsP2 = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:POWERFLOW_P2"})
    $Global:HasPhoneSys = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MCOEV"})
    $Global:HasAudioConf = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MCOMEETADV"})
    $Global:HasRmtAssist = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:MICROSOFT_REMOTE_ASSIST"})
    $Global:HasInfoBarr = [bool] ($Global:UserLicense.Licenses |Where-Object {$_.AccountSkuID -eq "ul:M365_INSIDER_RISK_MANAGEMENT"})
}

Function O365Licenses
{
    $Global:O365Lic = (Get-MsolUser -UserPrincipalName $Global:UPN).Licenses
	$Global:UserDet = $Global:UPN + " - " + $Global:UserLicense.DisplayName + " (" + $Global:ADUser.ExtensionAttribute1 + ")"

    If ($O365Lic.count -gt 0)
    {
        $Global:LicAssigned = ""
        $LineToWrite = "Reviewing O365 Licenses Assigned to " + $Global:UPN
	    WriteLogEvent

        foreach ($Global:O365Lic in $Global:O365Lic)
        {
            switch ($Global:O365Lic.AccountSkuID)
            {
                "ul:ENTERPRISEPACK"
                {
                    $text = "Enterprise E3"
                }
                "ul:EXCHANGEENTERPRISE"
                {
                    $text = "Exchange Online Plan 2"
                }
                "ul:POWER_BI_STANDARD"
                {
                    $text = "PowerBI (Free)"
                }
                "ul:POWER_BI_PRO"
                {
                    $text = "PowerBI Pro" 
                }
                "ul:DYN365_ENTERPRISE_PLAN1"
                {
                    $text = "Dynamics 365 Customer Engagement Plan"
                }
                "ul:EMS"
                {
                    $text = "Enterprise Mobility Suite/Intune (EMS)"
                }
                "ul:POWER_BI_INDIVIDUAL_USER"
                {
                    $text = "Power BI for O365"
                }
                "ul:MCOIMP"
                {
                   $text = "Skype for Business Online (Plan 1)"
                }
                "ul:SHAREPOINTSTANDARD"
                {
                    $text = "SharePoint Online (Plan 1)"
                }
                "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
                {
                    $text = "Microsoft Relationship Sales Solution"
                }
                "ul:TEAMS_COMMERCIAL_TRIAL"
                {
                    $text = "Microsoft Teams Commercial Cloud (User Initiated)"
                }
                "ul:ATP_ENTERPRISE"
                {
                    $text ="Advanced Threat Protection Plan1"
                }
                "ul:MDATP_XPLAT"
                {
                    $text ="Advanced Threat Protection for MAC"
                }
                "ul:WIN_DEF_ATP"
                {
                    $text ="Defender Advanced Threat Protection"
                }
                "ul:FLOW_P2"
                {
                    $text = "Microsoft Flow Plan 2"
                }
                "ul:MEETING_ROOM"
                {
                    $text = "Meeting Room"
                }
                "ul:POWERFLOW_P2"
                {
                    $text = "PowerApps Plan 2"
                }
                "ul:MCOEV"
                {
                    $text = "Phone System"
                }
                "ul:MCOMEETADV"
                {
                    $text = "Audio Conferencing"
                }
                "ul:MICROSOFT_REMOTE_ASSIST"
                {
                    $text = "Microsoft Remote Assist"
                }
                "ul:MCOCAP"
                {
                    $text = "Common Area Phone"
                }
            }

            If ($Global:LicAssigned -eq "")
            {
                $Global:LicAssigned = $Text
            }
            else
            {
                $Global:LicAssigned = $Global:LicAssigned + ", " + $Text
            }
            $LineToWrite = "UPDA" + "`t" + $text
	        WriteLogEvent
        }
    }
    else
    {
        $Global:LicAssigned = "No licenses assigned to this account"
        $LineToWrite = "No O365 Licenses Assigned to " + $Global:UPN
	    WriteLogEvent
    }
}

# This prints details regarding mailbox access to the console and report file
Function PrtMBxAccess($Global:MbxAccess)
{
    $Global:MbxAccess = Get-MailboxPermission $ENo |Where-Object {$_.User -like "*global.ul.com"}
    $Global:MbxFldrAccess = Get-MailboxFolderPermission $ENo |Where-Object {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous")}
    $Global:MbxAccessForm = "None"
    foreach ($Global:MbxAccess in $Global:MbxAccess)
    {
#       write-host "Accounts Granted Access       : " $Global:MbxAccess.User -ForegroundColor Yellow
        If ($Global:MbxAccessForm -eq "None")
        {
            $Global:MbxAccessForm = (get-mailbox $Global:MbxAccess.User).PrimarySMTPAddress + " (" + $Global:MbxAccess.AccessRights + ")"
            $SupMbx = (get-mailbox $Global:MbxAccess.User).PrimarySMTPAddress
            $LineToWrite = $WhoAmI + "`t" + "User Granted Mailbox Permission:" + "`t" + $(get-mailbox $Global:MbxAccess.User).PrimarySMTPAddress + "- PermissionLevel: " + $Global:MbxAccess.AccessRights
            WriteReportEvent
        }
        else
        {
            $Global:MbxAccessForm = $Global:MbxAccessForm + ", " + (get-mailbox $Global:MbxAccess.User).PrimarySMTPAddress + " (" + $Global:MbxAccess.AccessRights + ")"
            $SupMbx = $SupMbx + ", " + (get-mailbox $Global:MbxAccess.User).PrimarySMTPAddress
            $LineToWrite = $WhoAmI + "`t" + "User Granted Mailbox Permission:" + "`t" + $(get-mailbox $Global:MbxAccess.User).PrimarySMTPAddress + "- PermissionLevel: " + $Global:MbxAccess.AccessRights
            WriteReportEvent
        }
    }

    foreach ($Global:MbxFldrAccess in $Global:MbxFldrAccess)
    {
        If ($Global:MbxAccessForm -eq "None")
        {
            $Global:MbxAccessForm = (get-mailbox $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + " (" + $Global:MbxFldrAccess.AccessRights + ")"
            $SupMbx = (get-mailbox $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress
            $LineToWrite = $WhoAmI + "`t" + "User Granted Folder Permission :" + "`t" + $(get-mailbox $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + "- PermissionLevel: " + $Global:MbxFldrAccess.AccessRights
            WriteReportEvent
        }
        else
        {
            $Global:MbxAccessForm = $Global:MbxAccessForm + ", " + (get-mailbox $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + " (" + $Global:MbxFldrAccess.AccessRights + ")"
            $SupMbx = $SupMbx + ", " + (get-mailbox $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress
            $LineToWrite = $WhoAmI + "`t" + "User Granted Folder Permission :" + "`t" + $(get-mailbox $Global:MbxFldrAccess.User.DisplayName).PrimarySMTPAddress + "- PermissionLevel: " + $Global:MbxFldrAccess.AccessRights
            WriteReportEvent
        }
    }

    If ($mbxAccessForm -ne "None")
    {
        if (((get-date).AddHours(-1) -le $Global:inf.litigationholddate) -and ($Global:inf.CustomAttribute1 -eq "T"))
        {
            $SendTo = $Global:LegalTeamMsgs
            $SendCC = "Sandi.Glazebrook@ul.com"
            if ($Mbx.CustomAttribute1 -like "*Employee*")
            {
                $MessageSubject = "Access Previously Granted to Terminated Employee Mailbox and OneDrive"
                $MessageBody = "While placing this terminated account on Legal Hold access was previously granted to the Mailbox and OneDrive information.  The details are provided below.<ul type=""disc""><li>Terminated " + $Global:inf.CustomAttribute1 + " details " + $Mbx.Alias + " - " + $Global:inf.DisplayName + "</li><li>Individual given access to this information " + $SupMbx.Alias + " - " + $SupMbx.DisplayName + "</li></ul>UL Account Provisioning Team</li></ul>"
            }
            else
            {
                $MessageSubject = "Access Previously Granted to Terminated Non-Employee Mailbox"
                $MessageBody = "While placing this terminated account on Legal Hold access was previously granted to thes Mailbox information.<ul type=""disc""><li>Terminated " + $Global:inf.CustomAttribute1 + " details " + $Global:inf.Alias + " - " + $Global:inf.DisplayName + "</li><li>Individual given access to this information " + $SupMbx.Alias + " - " + $SupMbx.DisplayName + "</li></ul>UL Account Provisioning Team</li></ul>"
            }
            invoke-expression -Command .\SendSMTPMessage.ps1
        }
    }
}

#  Removes access to a mailbox
Function RemoveAccess($ENo)
{
    Do
    {
        write-host "Enter Employee Number for Access to Removed From" $ENo "? " -ForegroundColor Yellow -NoNewline
        $RemENo = Read-Host
        Remove-MailboxPermission $ENo -AccessRights FullAccess -User $RemENo -Confirm:$false
        $LineToWrite = $WhoAmI + "`t" + "Removed " + $RemENo + "Access to " + $ENo + "Mailbox" + "`n"
        WriteReportEvent
        Write-Host "Removed" $RemENo "Access to" $ENo "Mailbox" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Remove Another Users access to this mailbox (Y/N)? " -ForegroundColor Yellow -NoNewline
        $RemAccess = Read-Host
    } while ($RemAccess -eq "Y")
}

Function RetentPolicy
{
    $MbxCreated = [bool](get-mailbox $EmpNo -ErrorAction SilentlyContinue)

    If ($MbxCreated -eq "True")
    {
        If (($Global:ADUser.extensionattribute4 -eq "IT") -or ((get-mailbox $EmpNo).WhenCreated -gt "03/23/2020"))
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

#  Diplays the account details to the console
Function ShowDetails
{
    write-host "Employee Number               : " $Global:inf.Alias -ForegroundColor Green
    write-host "DisplayName                   : " $Global:inf.DisplayName -ForegroundColor Green
    write-host "Employee Type                 : " $Global:inf.CustomAttribute1 -ForegroundColor Green
    write-host "AD Container                  : " $Global:strUserPath -ForegroundColor Green
    write-host "LitigationHoldEnabled         : " $Global:inf.LitigationHoldEnabled -ForegroundColor Green
    write-host "LitigationHoldDate            : " $Global:inf.LitigationHoldDate -ForegroundColor Green
    write-host "LigitationHoldOwner           : " $Global:inf.LitigationHoldOwner -ForegroundColor Green
    write-host "Active InPlace Holds          : " ($Global.inf.InPlaceHolds -join ",") -ForegroundColor Green
    write-host "Hidden from Address Book O365 : " $Global:inf.HiddenFromAddressListsEnabled -Foregroundcolor Green
    write-host "Hidden from Address Book AD   : " $Global:u.msExchHideFromAddressLists.value -Foregroundcolor Green
    write-host "O365 Retention Comment        : " $Global:inf.RetentionComment -ForegroundColor Green
    write-host "AD Retention Comment          : " $Global:ADCmt -ForegroundColor Green
    write-host "AD Object Protected           : " $Global:UsrDetails.ProtectedFromAccidentalDeletion -ForegroundColor Green
    PrtMBxAccess($Global:MbxAccess)
}

function UnAssign-ATPDefLic
{
    if ($HasATPDef -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Windows Defender Advanced Threat Protection license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Windows Defender Advanced Threat Protection License Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Removing Windows Defender Advanced Threat Protection License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:WIN_DEF_ATP"
    }
    WriteLogEvent
}

function UnAssign-ATPPDefMACLic
{
    if ($HasATPDefMAC -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Advanced Threat Protection for MAC license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Advanced Threat Protection for MAC License Assigned to " + $Global:UPN
    }
    else
    {
        $LineToWrite = "REVI" + "`t" + "Removing Advanced Threat Protection for MAC License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:MDATP_XPLAT"
    }
    WriteLogEvent
}

function UnAssign-ATPP1Lic
{
    if ($HasATPP1 -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Advanced Threat Protection Plan1 license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Advanced Threat Protection Plan1 License Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Removing Advanced Threat Protection Plan1 License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:ATP_ENTERPRISE"
    }
    WriteLogEvent
}

Function UnAssign-AudioConferencingLic
{
    if ($HasAudioConf -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have an Audio Conferencing license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Audio Conferencing License Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Removing Audio Conferencing License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:MCOMEETADV"
    }
    WriteLogEvent
}

Function UnAssign-CommonAreaPhoneLic
{
    if ($HasCAP -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Common Area Phone license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Common Area Phone License Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Removing Common Area Phone License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:MCOCAP"
    }
    WriteLogEvent
}
    
function UnAssign-CRMLic
{
    If ($HasCRM -ne "True")
    {
      	$Output = $wshell.Popup("This individual does not have a Microsoft Phone System license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Dynamics 365 Customer Engagement Plan License Assigned to " + $Global:UPN
    }
	else
	{
        $LineToWrite = "REVI" + "`t" + "Removing Dynamics 365 Customer Engagement Plan License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
    }    
    WriteLogEvent
}

function UnAssign-E3Lic
{
    If ($HasE3 -ne "True")
    {
      	$Output = $wshell.Popup("This individual does not have a Enterpris E3 license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Dynamics 365 Customer Engagement Plan License Assigned to " + $Global:UPN
    }
    else
    {
        $LineToWrite = "REVI" + "`t" + "Removing Enterprise E3 Plan License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:ENTERPRISEPACK"
    }
    WriteLogEvent
}

function UnAssign-EMSLic
{
    If ($HasEMS -ne "True")
	{
        $Output = $wshell.Popup("This individual does not have a EMS license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Enterprise Mobility Suite/Intune (EMS) License Assigned to " + $Global:UPN
    }
	else
    {
		Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:EMS"
        $LineToWrite = "REVI" + "`t" + "Removing Enterprise Mobility Suite/Intune (EMS) License from " + $Global:UPN
    }
    WriteLogEvent
}

function UnAssign-MeetingLic
{
    if ($HasMeeting -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Meeting Room license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Meeting Room License Already Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Removing Meeting Room License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:MEETING_ROOM"
    }
    WriteLogEvent
}

Function UnAssign-MRSSLic
{
    if ($HasMRSS -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Microsoft Relationship Sales Solution license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Microsoft Relationship Sales Solution License Assigned to " + $Global:UPN
    }
    else
    {
        $LineToWrite = "REVI" + "`t" + "Removing Microsoft Relationship Sales Solution License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:DYN365_ ENTERPRISE _RELATIONSHIP_SALES"
    }
	WriteLogEvent 
}

function UnAssign-P2Lic
{
    if ($HasExP2 -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Exchange Online P2 license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Exchange Online P2 License Assigned to " + $Global:UPN
    }
    else
    {
        $LineToWrite = "REVI" + "`t" + "Removing Exchange Online P2 License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:EXCHANGEENTERPRISE"
    }
    WriteLogEvent
}

Function UnAssign-PAppsP2Lic
{
    if ($HasPAppsP2 -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a PowerApps P2 license assigned.",0,"License Noty Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No PowerApps Plan 2 License Assigned to " + $Global:UPN
    }
    else
    {
       	$LineToWrite = "REVI" + "`t" + "Removing PowerApps Plan 2 License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:POWERFLOW_P2"
    }
    WriteLogEvent
}

Function UnAssign-PhoneSystemLic
{
    if ($HasPhoneSys -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a Microsoft Phone System license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No Phone System License Assigned to " + $Global:UPN
    }
    else
    {
     	$LineToWrite = "REVI" + "`t" + "Removing Phone System License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicense "ul:MCOEV"
    }
    WriteLogEvent
}

function UnAssign-PowerBIProLic
{
    if ($HasBIPro -ne "True")
    {
        $Output = $wshell.Popup("This individual does not have a PowerBI Pro license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "REVI" + "`t" + "No PowerBI Pro License Assigned to " + $Global:UPN
    }
    else
    {
      	$LineToWrite = "REVI" + "`t" + "Removing PowerBI Pro License from " + $Global:UPN
        Set-MsolUserLicense -UserPrincipalName $Global:UPN -AddLicense "ul:POWER_BI_PRO"
    }
    WriteLogEvent
}

function UnAssign-RemoteAsstLic
{
    If ($HasRmtAssist -eq "True")
    {
		Set-MsolUserLicense -UserPrincipalName $Global:UPN -RemoveLicenses "ul:MICROSOFT_REMOTE_ASSIST"
        $LineToWrite = "DELE" + "`t" + "Removing Microsoft Remote Assist License from " + $Global:UPN
    }
    else
    {
        $Output = $wshell.Popup("This individual does not have a Microsoft Remote Assist license assigned.",0,"License Not Assigned",0+32)
        $LineToWrite = "DELE" + "`t" + "No Microsoft Remote Assist License assigned to " + $Global:UPN
    }
    WriteLogEvent
}

#  Writes events to the Log File
function WriteLogEvent
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LocalMachineName + "`t" + $LineToWrite
	Out-File -filepath $LogFile -append -noClobber -inputObject $RecordEvent
} #end WriteLogEvent


#  Writes events to the Report File
function WriteReportEvent
{
    If ($ReportFile -eq $null)
    {
        $ReportFile = $Global:ReportFile
    }
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LineToWrite
	Out-File -filepath $ReportFile -append -noClobber -inputObject $RecordEvent
} #end WriteReportEvent

################

# Retrieve the user name
	$WhoAmI			= WhoAmI
		
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name

$wshell = New-Object -ComObject Wscript.Shell

Invoke-Expression -Command e:\Automation\Scripts\O365MainMenu.ps1