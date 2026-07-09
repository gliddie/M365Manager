#################################################################################
# 
# PowerShell source code
# Revision v1.05
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Enable MailUSer and Assign O365 License for existing AD User account
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 04/21/2016
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '    04/21/2016 SAG Created Script
#    '    11/14/2016 SAG Added code to disable the TeamSites
#    '    12/13/2016 SAG Added code so that if no E4 licenses are available an E3 license will be assigned and the O365 team will be notified that no E4 licenses are available
#    '    02/06/2017 SAG Added code so that if not connect to the MSOLService ueser is prompted again for the password
#    '	  03/15/2017 SAG Added Deskless to the disabled plans as this is a new sub-license type
#    '    03/29/2017 SAG Added code to force connection to the MSOLservice to use the windows credential dialogue box.
#    '    05/09/2017 SAG Commented out the assigning of the PowerBI free license
#    '    05/25/2017 SAG Added line to enable the EMS License for all
#    '    05/31/2017 SAG Added like to update the number of E3 Licenses assigned to include current count into email
#    '    06/19/2017 SAG New disabling of new FORMS license types
#    '    07/03/2017 SAG Removing TEAMS form the disabled list
#    '    07/25/2016 SAG Re-enabled the setting of the PowerBI Free license
#    '    11/15/2017 SAG Added Code to check if there are available EMS licenses and to see an E3 or E4 license is already assigned
#    '    01/08/2018 SAG Fixed code where EMS license is assigned but text prints indicating no available licenses and fixed text printed to ReportFile
#    '    04/28/2018 SAG Added code to add to ACL.UL.IntuneEnforced group for NA/CA/LA staff
#    '    06/14/2018 SAG Added code to trim spaces from the start of the email address field
#    '    06/25/2018 SAG Added code to assign ATP license to all staff
#    '    07/09/2018 SAG Added reporting for the assignment of licenses and adding users to the ACL.UL.IntuneEnforced Group for NA/CA/LA staff
#    '    10/04/2018 SAG Added code so that the SDAP staff can confirm that all values have been set
#    '    01/08/2019 SAG Added line to see who is running the script when the No Licenses report is generated
#    '    01/22/2019 SAG Fixed the WhoAmI variable name in the line to print who executed the script.
#    '    02/19/2019 SAG Changed email receipients removed Thom Staples, added BJ Stone and Joe Park
#    '    03/12/2019 SAG Changed to assign E3 licenses to new staff
#    '    03/25/2019 SAG Added setting Intune to the Cambridge Staff
#    '    05/28/2019 SAG Modified code to allow overallocation of ATP and EMS licenses
#    '    06/14/2019 SAG Commented out the assignment of the ATP license
#    '    08/05/2019 SAG Fixed line where there was no # as the first character and changed usnbkd300p to usnbkadds001p
#    '    08/11/2019 SAG Added code to Enable New Users in the Not for Profit Org with @ul.org addresses
#    '    08/28/2019 SAG Added code to skip enabling accounts where the ExtensionAttribute8 does not have a value
#    '    09/09/2019 SAG Added code so that the Intune enforced configuration is enabled for all new staff
#    '    09/18/2019 SAG Added code to document the selection entered by the agent if the Org2 field is blank
#                        Added code to remove the Teams Commercial License as it is causing licensing issues and granting access to items some individuals shoul not have access to
#    '    09/19/2019 SAG Removed Joe Park from license report and removed extra line feeds for the report file.
#==========================================================================
#
#################################################################################
$WhoAmI			= WhoAmI
$DC             = "usnbkadds004d.global.uldev.com"

$LogDirectory	= "E:\Automation\EnableDEVUserAssignLicense\Log"
$LogFile		= $LogDirectory + "\" + "Log-EnableUserAssignLicense.log"

$InputDirectory	= "E:\Automation\EnableDEVUserAssignLicense\Input"
$InputFile		= $InputDirectory + "\" + "Input-EnableUserAssignLicense.csv"

$ReportDirectory	= "E:\Automation\EnableDEVUserAssignLicense\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-EnableUserAssignLicense"

$RoutingDomain	= "@global.uldev.com"

#	Retrieve the local server name
		$Machine = get-wmiobject "Win32_ComputerSystem"
		$LocalMachineName = $Machine.Name

#	Retrieve the current Date and Time for use in log files
		$uDate = get-date -uformat %D
		$uTime = get-date -uformat %T
				
		$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"

#	Convert the current Date and Time to format MM-DD-YY and HHMMSS for use in file names	
		$Date  = $uDate.Replace("/", "-")
		$Time  = $uTime.Replace(":", "")
		
		$ReportFile = $ReportFile + "-" + "Date" + $Date + "Time" + $Time + ".Log"	
	
#	Location to write file if there are no E4 licenses available
		$NoLicRpt = "E:\Automation\EnableDEVUserAssignLicense\Log\NoE4Licenses" + "-" + "Date" + $Date + ".Log"
		
#	Create logging folder $LogDirectory if it's not present
		if (Test-Path $LogDirectory)
		{
			# the directory is present
		}
		else
		{
			mkdir $LogDirectory
		}
		
#	Add start record to log file
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "Enable MailUser/EnableUserAssignLicense script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

# =============================================================================================================================================

Invoke-Expression -command E:\O365AdminShared\Scripts\O365DisabledLicenseFeatures.ps1

if (Test-Path $InputFile)
{					
	$UserAccounts = Import-CSV $InputFile

	ForEach ($User in $UserAccounts)
    {
	   if ($User.UPN.contains("@"))
	   {
            $atUPN  = $User.UPN.indexOf("@")
		    $atMAIL = $User.MAIL.indexOf("@")
		    $empID   = $User.UPN.substring(0,$atUPN)
		    $Name    = $User.MAIL.substring(0,$atMAIL).TrimStart("")
		    write-host "`nStarting Enablement Process for: " $User.UPN " " $Name
            $LDAPFilter = "(userPrincipalName=" + $User.UPN + ")"
		    $ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,targetaddress,proxyaddresses,extensionattribute1,extensionattribute5,extensionattribute8,extensionattribute15,msExchHideFromAddressLists
            $PrimaryMail = "@uldev.com"
            Out-File -filepath $ReportFile -append -noClobber -inputObject ""
			
            If ($AdUser.ExtensionAttribute8 -ne "Not for Profit")
            {
                If (($AdUser.ExtensionAttribute8 -eq "") -or ($AdUser.ExtensionAttribute8 -eq $null))
                {
                    $NFPBU = "X"
                    $PrimaryMail = "CannotRegister"
                    write-host "`nThe value for the Business Unit is not yet set. `nBefore continuing with the Enablement of this account you must provide the Business Unit for this individual." -ForegroundColor Yellow
                    write-host "Select <" -ForegroundColor Cyan -NoNewline
                    write-host "Y" -ForegroundColor Red -NoNewline
                    write-host "es> if you have confirmed they are in the 'Not for Profit' business unit`n       <" -ForegroundColor Cyan -NoNewline
                    write-host "N" -ForegroundColor Red -NoNewline
                    write-host "o> if you have confirmed they are in one of the other busness units`n       <" -ForegroundColor Cyan -NoNewline
                    write-host "U" -ForegroundColor Red -NoNewline
                    Write-Host "nsure> if you do not know what the correct businnes unit is " -ForegroundColor Cyan -NoNewline
                    write-host "`nIs this individual in the " -ForegroundColor Cyan -NoNewline
                    write-host "Not for Profit " -ForegroundColor Red -NoNewline
                    write-host "business unit (Y/N/U) ?" -ForegroundColor Cyan -NoNewline
                    $NFPBU = read-host
                    
                    switch -Wildcard ($NFPBU)
                    {
                        "Y*"
                        {
                            $PrimaryMail = "@uldev.org"
			                $LineToWrite = $RecordEvent + "INFO" + "`t" + "Agent confirmed to continue with UL.ORG registration for " + $User.UPN
		                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                        }
                        "N*"
                        {
			                $PrimaryMail = "@uldev.com"
                            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Agent confirmed a for UL.COM business unit for " + $User.UPN
		                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                        }
                        "U*"
                        {
                            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Agent confirmed the business unit is unknown and registration cannot continue for " + $User.UPN
		                    Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                        }
                    }
                }
            }
            else
            {
                $PrimaryMail = "@uldev.org"
            }

		    If ($PrimaryMail -ne "CannotRegister")
            {
                write-host "This user will be assigned a" $PrimaryMail "email address" $Name -ForegroundColor Green

		        $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + $Name + $PrimaryMail + "`t" + $User.LegacyMail + "`t" + $DC
                Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

		        $LineToWrite = $RecordEvent + "INFO" + "`t" + $empID + "`t" + $Name
		        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

                write-host "Enabling MailUser and Assigning Licenses for " $Name -ForegroundColor Green
			
			    write-host "Prior to execution AD account values are set as follows:"
			    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Prior to execution AD account values are set as follows:"
			    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
			    write-host "     User Mail Address:     " $ADUser.mail
			    $LineToWrite = $RecordEvent + "INFO" + "`t" + "     User Mail Address:     " + $Name + $PrimaryMail
			    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
			    write-host "     User Target Address:   " $ADUser.targetaddress
			    $LineToWrite = $RecordEvent + "INFO" + "`t" + "     User Target Address:   " + $ADUser.targetaddress
			    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
			    write-host "     User Proxy Addresses:  " $ADUser.proxyaddresses
			    $LineToWrite = $RecordEvent + "INFO" + "`t" + "     User Proxy Addresses:  " + $ADUser.proxyaddresses
			    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
			    write-host "     ExtensionAttibute15:   " $ADUser.extensionattribute15
			    $LineToWrite = $RecordEvent + "INFO" + "`t" + "     ExtensionAttibute15:   " + $ADUser.extensionattribute15
			    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

                Enable-MailUser $User.UPN -Alias $empID -ExternalEmailaddress ($Name + $PrimaryMail) -PrimarySmtpAddress ($Name + $PrimaryMail) -DomainController:$DC
			
    #	Start Check for Internet Mail = @ul.com or @ul.org
                Set-MsolUserLicense -UserPrincipalName $User.UPN -RemoveLicenses "ul:TEAMS_COMMERCIAL_TRIAL" -erroraction SilentlyContinue
                if ($User.Mail -like "*@ul.*")
                {
			      Set-MailUser $User.UPN -EmailAddresses (((Get-MailUser $User.UPN -DomainController:$DC).EmailAddresses)+=($Name + $RoutingDomain)) -CustomAttribute15 "EnableMailUser PS Date: $Date PS Time: $Time" -DomainController:$DC
                }
                else
                {
			      Set-MailUser $User.UPN -EmailAddresses (((Get-MailUser $User.UPN -DomainController:$DC).EmailAddresses)+=$User.MAIL,($Name + $RoutingDomain)) -CustomAttribute15 "EnableMailUser PS Date: $Date PS Time: $Time" -DomainController:$DC
                }

		        If ($ADUser.displayName.contains(",") -eq $false)
		        {
			        If ($ADUser.initials.length -gt 0)
			        {
				        $ADName = $ADUser.sn + ", " + $ADUser.givenName + " " + $ADUser.initials.substring(0,1) + "."
			        }
                    else
                    {
				        $ADName = $ADUser.sn + ", " + $ADUser.givenName
			        }
		            Set-ADUser $ADuser -DisplayName $ADName
			        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Changed displayName to: " + $ADName
			        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		        }

    #	Set UsageLocation, Assign Free PowerBI License and enable Standard Email License for this Employee Type

                Set-MsolUser -UserPrincipalName $User.UPN -UsageLocation US -ErrorAction Silentlycontinue
		
                $HasE3 = [bool]((Get-MsolUser -UserPrincipalName $User.UPN).Licenses |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"})
		        $HasEMS = [bool]((Get-MsolUser -UserPrincipalName $User.UPN).Licenses |Where-Object {$_.AccountSkuID -eq "ul:EMS"})
		        $HasPBIF = [bool]((Get-MsolUser -UserPrincipalName $User.UPN).Licenses |Where-Object {$_.AccountSkuID -eq "ul:Power_BI_Standard"})

                if ($HasPBIF -eq "True")
		        {
			        $LineToWrite = $RecordEvent + "INFO" + "`t" + "PowerBI Free License Already Assigned to " + $User.UPN
		            Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "PowerBI Free License Already Assigned"
		            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite	
		        }
		        else
		        {
			        Set-MsolUserLicense -UserPrincipalName $User.UPN -AddLicenses "ul:POWER_BI_STANDARD" -erroraction SilentlyContinue
			        $LineToWrite = $RecordEvent + "INFO" + "`t" + "PowerBI Free License to " + $User.UPN
			        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
			        $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "Assigned PowerBI Free License"
			        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite	            
		        }
		
                if ($HasEMS -eq "True")
		        {
			        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Enterprise Mobility Suite License Already Assigned to " + $User.UPN
		            Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "(EMS) Enterprise Mobility Suite License Already Assigned"
		            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite	
		        }
		        else
		        {
				        Set-MsolUserLicense -UserPrincipalName $User.UPN -AddLicenses "ul:EMS" -ErrorAction SilentlyContinue
				        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Assigned Enterprise Mobility Suite License to " + $User.UPN
				        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				        $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "Assigned (EMS) Enterprise Mobility Suite License"
				        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite		
		        }
		
                If ($Aduser.msExchHideFromAddressLists -eq "True")
                {
                    write-host "Setting user to be Hidden from the Address Book for" $EmpNo
                    $ADUser.msExchHideFromAddressLists = $False
                    $ADUser.CommitChanges()	
                }

                If ((($Aduser.extensionattribute5 -ne "Europe") -and ($Aduser.extensionattribute5 -ne "Asia Pacific")) -and ($Aduser.extensionattribute3 -like "*Cambridge*"))
                {
                    $IntuneMem = "off"
                    $ADGroupMem = Get-ADPrincipalGroupMembership $EmpID
                    foreach ($ADGRoupMem in $ADGroupmem)
                    {
                        if ($ADGroupMem.SamAccountName -like "*IntuneEnfor*")
                        {
                            $IntuneMem = "1"
                        }
                    }
                    If ($IntuneMem -eq "off")
                    {
                        Switch ($Aduser.extensionattribute5)
                        {
                            "Europe"
                            {
                                Add-ADGroupMember -Identity "ACL.EU.IntuneEnforced" -Members $EmpID -Confirm:$False
                                write-host  "Adding to the ACL.EU.IntuneEnforced group"
			                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Adding to the ACL.EU.IntuneEnforced group for " + $User.UPN
		                        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				                $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "Added to ACL.EU.IntuneEnforced Group"
				                Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                            }

                            "Asia Pacific"
                            {
                                Add-ADGroupMember -Identity "ACL.AP.IntuneEnforced" -Members $EmpID -Confirm:$False
                                write-host  "Adding to the ACL.AP.IntuneEnforced group"
			                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Adding to the ACL.AP.IntuneEnforced group for " + $User.UPN
		                        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				                $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "Added to ACL.AP.IntuneEnforced Group"
				                Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                            }

                            default
                            {
                                Add-ADGroupMember -Identity "ACL.UL.IntuneEnforced" -Members $EmpID -Confirm:$False
                                write-host  "Adding to the ACL.UL.IntuneEnforced group"
			                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Adding to the ACL.UL.IntuneEnforced group for " + $User.UPN
		                        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				                $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "Added to ACL.UL.IntuneEnforced Group"
				                Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                            }
                        }
                    }
			        else
			        {
			            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Already a member an IntuneEnforced group for " + $User.UPN
		                Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
				        $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "Already a member of an L.IntuneEnforced group"
				        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite				
			        }
                }			
        #	Set SSKID to the E3 license and test to see if there are available licenses
    	        $SSKID = "ul:ENTERPRISEPACK"
		        $LicType = "Enterprise E3"
    	        $DisPlanEmp = $Global:DisPlanEmpE3
	            $DisPlanNonEmp = $Global:DisPlanNonEmpE3

                $E3Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}
                $EMSLic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:EMS"}

		        if ($E3Lic.ConsumedUnits -ge $E3Lic.ActiveUnits)
		        {
			        If (Test-Path $NoLicRpt)
			        {
        #	Do Nothing File Exists and an email was already sent to the team
				        $LineToWrite = "Assigned " + $LicType + " license to " + $EmpID + " - " + $Name
				        Out-File -FilePath $NoLicRpt -InputObject $LineToWrite -Append -NoClobber
				        $LineToWrite = "E3 Licenses Assigned " + (Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}).ConsumedUnits + " E3 License Allocation " +(Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}).ActiveUnits
				        Out-File -FilePath $NoLicRpt -InputObject $LineToWrite -Append -NoClobber
			        }
			        else
			        {
        #	Location to write file for tracking the frequency that no E4 licenses are available
        #	Also prevents this message from being sent more than 1 time per day
				        Out-File -FilePath $NoLicRpt -InputObject "No E3 Licenses Available"
				        $LineToWrite = ""
				        Out-File -FilePath $NoLicRpt -InputObject $LineToWrite -Append -NoClobber
				        $LineToWrite = "Assigned " + $LicType + " license to " + $EmpID + " - " + $Name
				        Out-File -FilePath $NoLicRpt -InputObject $LineToWrite -Append -NoClobber
				        $LineToWrite = "E3 Licenses Assigned " + (Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}).ConsumedUnits + " E3 License Allocation " + (Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}).ActiveUnits
				        Out-File -FilePath $NoLicRpt -InputObject $LineToWrite -Append -NoClobber
        #	Send email to notify appropriate individuals that no E3 licenses are available
                        $E3Lic = Get-MsolAccountSku |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"}				
				        $server = "smtp-relay.ul.com"
				        $client = new-object system.net.mail.smtpclient $server
				        $from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "UL Account Provisioning Process"           
				        $to = $from 
				        $SendTo = "LST.O365AdminTeam@ul.com,BJ.Stone@ul.com,Ginger.M.Tucibat@ul.com"
				        $message = new-object  System.Net.Mail.MailMessage $from, $to 
				        $message.IsBodyHtml = $true
				        $msgfont = "<basefont face=verdana size=2.5 color=black>"
				        $message.Subject = "No Enterprise E3 Licenses Available"
				        $message.To.Clear()
				        $message.Body = $msgfont + "<p>Person executing script: " + $whoami + "<p>Against user account:  " + $User.UPN + "<p>You are receiving this message because there are no Enterprise E3 licenses available.  <ul type=""disc""><li>E3 Licenses Active/Consumed/Available:  " + $E3Lic.ActiveUnits + "/" + $E3Lic.ConsumedUnits + "/" + ($E3Lic.ActiveUnits-$E3Lic.ConsumedUnits)
				        $message.To.Add($SendTo)
				        $client.Send($message)				
			        }
		        }
 
        #        Write-Host "Assigning" $LicType "License to" $EmpNo "-" $Name -ForegroundColor Green
                IF ($HasE3 -ne "True")
                {
		            If (($ADUser.ExtensionAttribute1 -like "*Employee*") -and ($HasE3 -ne "True"))
                    {
			            Write-Host "Assigning" $LicType "License with Employee Type Licenses Enabled to" $Name -ForegroundColor Green
                        $DLO = ($DisPlanEmp.Split(“,”))
		            }
                    else
                    {
			            Write-Host "Assigning" $LicType "License with Non-Employee Type Licenses Enabled to" $Name -ForegroundColor Green
			            $DLO = ($DisPlanNonEmp.Split(“,”))
                    }

            #  Assign the license		
		            $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
                    Set-MsolUserLicense -UserPrincipalName $User.UPN -AddLicenses $sskid –LicenseOptions $MyO365Sku

                    $HasE3 = [bool]((Get-MsolUser -UserPrincipalName $User.UPN).Licenses |Where-Object {$_.AccountSkuID -eq "ul:ENTERPRISEPACK"})
                    if ($HasE3 -eq "True")
                    {
                        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Assigned " + $LicType + " Licenses to " + $User.UPN
		                Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                        $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "Assigned " + $LicType + " Licenses"
		                Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                    }
                    else
                    {
                        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unable to assign " + $LicType + " Licenses to " + $User.UPN
		                Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                        $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "Unable to assign " + $LicType + " Licenses"
		                Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                    }
                }
                else
                {
                    write-host "E3 License Already assigned to this account"
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + $LicType + " License Already Assigned "
		            Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                }

		        Start-sleep 30
		        $ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,targetaddress,proxyaddresses,extensionattribute1,extensionattribute5,extensionattribute15
		        write-host "After execution AD account values are set as follows:"
		        $LineToWrite = $RecordEvent + "INFO" + "`t" + "After execution AD account values are set as follows:"
		        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite		
		        write-host "     User Mail Address:     " $ADUser.mail
		        $LineToWrite = $RecordEvent + "INFO" + "`t" + "     User Mail Address:     " + $ADUser.mail
		        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		        write-host "     User Target Address:   " $ADUser.targetaddress
		        $LineToWrite = $RecordEvent + "INFO" + "`t" + "     User Target Address:   " + $ADUser.targetaddress
		        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		        write-host "     User Proxy Addresses:  " $ADUser.proxyaddresses
		        $LineToWrite = $RecordEvent + "INFO" + "`t" + "     User Proxy Addresses:  " + $ADUser.proxyaddresses
		        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		        write-host "     ExtensionAttibute15:   " $ADUser.extensionattribute15
		        $LineToWrite = $RecordEvent + "INFO" + "`t" + "     ExtensionAttibute15:   " + $ADUser.extensionattribute15
		        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            }
            else
            {
                write-host "`nThis individual cannot be configured until the Business Unit can be identified." -ForegroundColor Red
                write-host "Skipping configuration of this user account" -ForegroundColor Red
                $LineToWrite = $RecordEvent + "INFO" + "`t" + "Business Unit is Not Identified for this Inidividual " + $User.UPN
		        Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "Business Unit is Not Identified for this Inidividual Skipping Configuration"
		        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite	
                pause
            }
        }
    }
    #	Rename the input file for future reference
	    Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
}
else
{
#   This executes if there is no input file
	write-host "Could not find input file :^) " $InputFile
	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Could not find input file "
	Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
}

# =============================================================================================================================================

	
#	Retrieve the current Date and Time for use in log files
		$uDate = get-date -uformat %D
		$uTime = get-date -uformat %T
		
		$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"

#	Write end record to the log file
		$LineToWrite = $RecordEvent + "STOP" + "`t" + "This instance is stopping."
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

# 	End of script #