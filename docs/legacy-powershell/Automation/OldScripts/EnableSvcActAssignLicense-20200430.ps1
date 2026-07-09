#################################################################################
# 
# PowerShell source code
# Revision v1.05
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Enable MailUSer and Assign O365 License for existing AD User account
#    'Called By    : EUMMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 04/21/2016
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '               08/05/2016 SAG Created Script
#    '               09/14/2016 SAG Modified to add setting of the AD ExtentionAttribute,
#    '                              Allowed Workstations and the Info field
#    '               11/15/2016 SAG Added code to so when you are configuring accounts
#    '                              that are not UM mailbox it prompts for the Account
#    '                              owner and sets the allowed computers to the ADFS servers
#    '               05/04/2017 SAG Added code for Oracle OFR and KFI Service Accounts
#    '               08/05/2019 SAG Changed USNBKD300P to usnbkadds001p
# ==========================================================================
#
#################################################################################

$WhoAmI			= WhoAmI
$DC			= "usnbkadds001p.global.ul.com"

$LogDirectory		= "E:\Automation\EnableSvcActAssignLicense\Log"
$LogFile		= $LogDirectory + "\" + "Log-EnableSvcActAssignLicense.log"

$InputDirectory		= "E:\Automation\EnableSvcActAssignLicense\Input"
$InputFile		= $InputDirectory + "\" + "Input-EnableSvcActAssignLicense.csv"

$ReportDirectory	= "E:\Automation\EnableSvcActAssignLicense\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-EnableSvcActAssignLicense.log"

$PrimaryMail		= "@ul.com"
$RoutingDomain		= "@global.ul.com"

#$session = get-PSsession
#Remove-PSSession $session

#		invoke-expression -Command .\ConnectO365.ps1
	# Add the Exchange PowerShell Snap-in
		Add-PSSnapin *Exchange* -erroraction SilentlyContinue

	# Add the ActiveDirectory modules.
   		Import-Module ActiveDirectory

	# Connect to the MsolService.
		Import-Module MSOnline
        Connect-MsolService -Credential $Global:LiveCred

#    Import-PSsession (get-PSSession)
		
	# Retrieve the local server name
		$Machine = get-wmiobject "Win32_ComputerSystem"
		$LocalMachineName = $Machine.Name

	# Retrieve the current Date and Time for use in log files
		$uDate = get-date -uformat %D
		$uTime = get-date -uformat %T
				
		$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"

	# Convert the current Date and Time to format MM-DD-YY and HHMMSS for use in file names	
		$Date  = $uDate.Replace("/", "-")
		$Time  = $uTime.Replace(":", "")
		
		$ReportFile = $ReportFile + "-" + "Date" + $Date + "Time" + $Time + ".Log"	
	
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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "Enable MailUser/EnableUserAssignLicense script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

# =============================================================================================================================================

if (Test-Path $InputFile)
{					
	$UserAccounts = Import-CSV $InputFile
	$Info = "Service Account Owner:  Enterprise Lync/Telecom Team`r`n`r`nThis service account will never be logged into it is created in order to assign an O365 License and assign a UM extension to allow the business to manage voice mail messages to a given number.`r`n`r`nIndividuals will be given access to the mailbox by the Mailbox Owner and they will use their personal credentials in order to access to the mailbox.  To determine who the owner of the mailbox is check the Properties of the group MBX." + $ActDisName  + ".ED in the Outlook Address Book."

	ForEach ($User in $UserAccounts)
    {
	   if ($User.UPN.contains("@"))
	   {
		    $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + $User.Mail + "`t" + $User.LegacyMail + "`t" + $DC + "`n"
		    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

		    $atUPN  = $User.UPN.indexOf("@")
		    $atMAIL = $User.MAIL.indexOf("@")
		    $empID   = $User.UPN.substring(0,$atUPN)
		    $Name    = $User.MAIL.substring(0,$atMAIL)

		    $LineToWrite = $RecordEvent + "INFO" + "`t" + $empID + "`t" + $Name + "`n"
		    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            
            If (([bool](get-Mailbox $User.UPN -ErrorAction SilentlyContinue)) -ne "True")
            {
                write-host "Enabling MailUser and Assigning Licenses for " $Name -ForegroundColor Green

	       	    If ([bool](Get-MailUser $User.UPN) -eq "False")
                {
                    Enable-MailUser $User.UPN -Alias $empID -ExternalEmailaddress ($Name + $PrimaryMail) -PrimarySmtpAddress ($Name + $PrimaryMail) -DomainController:$DC
    			    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Mail Enabled the Account for: " + $ADName + "`n"
	    		    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                }
                else
                {
                    write-host "This account is already mail enabled" -ForegroundColor Red
    			    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Account is Already Mail Enabled for: " + $ADName + "`n"
	    		    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                }
			
                # Start Check for Internet Mail = @ul.com
		          if ($User.Mail.contains($PrimaryMail))
        		    {
        			  Set-MailUser $User.UPN -EmailAddresses (((Get-MailUser $User.UPN -DomainController:$DC).EmailAddresses)+=($Name + $RoutingDomain)) -CustomAttribute1 "ServiceAccount" -CustomAttribute15 "EnableServiceAcct PS Date: $Date PS Time: $Time" -DomainController:$DC
        		    }
        		   else
        		    {
        			  Set-MailUser $User.UPN -EmailAddresses (((Get-MailUser $User.UPN -DomainController:$DC).EmailAddresses)+=$User.MAIL,($Name + $RoutingDomain)) -CustomAttribute1 "ServiceAccount" -CustomAttribute15 "EnableServcieAcctr PS Date: $Date PS Time: $Time" -DomainController:$DC
        		    }
             }

		    $LDAPFilter = "(userPrincipalName=" + $User.UPN + ")"
		    $ADUser = Get-ADUser -LDAPFilter $LDAPFilter -Properties sn,givenName,initials,displayName,mail,proxyaddresses,info,userWorkstations
            
            $CheckDispName = Read-Host "Is this the correct DisplayName" $ADUser.displayname "(Y/N)? "
                        
            If ($CheckDispName -eq "N")
		    {
			    $ActDisName = read-host "Enter the DisplayName for the account" $User.UPN
		        Set-ADUser $ADuser -DisplayName $ActDisName
			    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Changed displayName to: " + $ADName + "`n"
			    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		    }
            else
            {
                $ActDisName = $ADUser.DisplayName
            }

			write-host "Is this account for a UM Mailbox (Y/N)? " -foregroundcolor cyan -nonewline
			$CheckUM = read-host
			
			if ($CheckUM -eq "Y")
			{
				Set-ADUser -Identity $ADUser -Replace @{info=$Info;userWorkstations="NotAllowed"}
				Move-ADObject -Identity $ADUser.ObjectGUID -TargetPath "OU=UM Mailboxes,OU=O365 Licensed,OU=ServiceAccounts,OU=Enterprise,DC=global,DC=ul,DC=com"
			}
			else
			{
				$ImagingProcess = "N"
                Write-Host "Who is the owner of this Service Account? " -foregroundcolor cyan -nonewline
				$AcctOwner = read-host
                $AcctOwner = "Account Owner:  " + $AcctOwner
                write-host "Is this account for the Oracle OFR or KFI Imaging Process (Y/N)? " -ForegroundColor cyan -NoNewline
                $ImagingProcess = read-host
                If ($ImagingProcess -eq "Y")
                {
                    Set-ADUser -Identity $ADUser -Replace @{info=$AcctOwner;userWorkstations="usnbka133d,usnbka397t,usnbka398p,usnbka397p,usnbkd521p,usnbkd522p"}
                    write-host "This account owner has been (1) set to" $AcctOwner "(2) to allow only the ADFS servers and Oracle Servers to login to this account and (3) moved to the O365 licensed OU" -foregroundcolor yellow
                }
                else
                {
    				Set-ADUser -Identity $ADUser -Replace @{info=$AcctOwner;userWorkstations="usnbkd521p,usnbkd522p"}
    				write-host "This account owner has been (1) set to" $AcctOwner "(2) to allow only the ADFS servers to login to this account and (3) moved to the O365 licensed OU" -foregroundcolor yellow
                }

				Move-ADObject -Identity $ADUser.ObjectGUID -TargetPath "OU=O365 Licensed,OU=ServiceAccounts,OU=Enterprise,DC=global,DC=ul,DC=com"
                write-host "Please manually move the Service Account to the appropriate OU in AD" -foregroundcolor yellow
			}
        }
        else
        {
# UPN is not valid
	      	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Invalid UPN in input file " + "`n"
	      	Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
        }

# Set UsageLocation, Assign Free PowerBI License and enable Standard Email License for this Employee Type

            Set-MsolUser -UserPrincipalName $User.UPN -UsageLocation US
#		    Set-MsolUserLicense -UserPrincipalName $User.UPN -AddLicenses "ul:POWER_BI_STANDARD" -erroraction SilentlyContinue
            
            write-host "Does this account need a email (Y/N)? " -ForegroundColor Cyan NoNewLine
            $MLic = Read-Host
            If ($MLic -eq "Y")
            {
                $SSKID = "ul:EXCHANGEENTERPRISE"
                $LicType = "ExchangeOnline Plan2"
                Write-Host "Assigning " $LicType "Licenses to" $Name -ForegroundColor Green
                Set-MsolUserLicense -UserPrincipalName $User.UPN -AddLicenses $SSKID
                $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "Assigned " + $SSKID + " License" + "`n"
                Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            }

            write-host "Will this account be used for remote assist (Y/N)? " -ForegroundColor Cyan NoNewLine
            $RmAssLic = Read-Host
            If ($RmAssLic -eq "Y")
            {
                $SSKID = "ul:EXCHANGEENTERPRISE"
                $LicType = "Remote Assist"
                Write-Host "Assigning " $LicType "Licenses to" $Name -ForegroundColor Green
                Set-MsolUserLicense -UserPrincipalName $User.UPN -AddLicenses $SSKID
                $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`t" + "Assigned " + $SSKID + " License" + "`n"
                Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            }

#            else
#            {
#                Write-Host "There are No" $LicType "Licenses available for assignment" -ForegroundColor Red
#                $LineToWrite = $RecordEvent + "INFO" + "`t" + $User.UPN + "`n"
#			    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
#            }
    }
	
	# Rename the input file for future reference
	Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log") -ErrorAction SilentlyContinue
    if (Test-Path $InputFile)
    {
        $loop = 0
        Do
        {
            write-host "The file" $InputFile "is open by another process please close the file and hit returnt to continue" -ForegroundColor Red -NoNewline
            $Cont = read-host
            Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log") -ErrorAction SilentlyContinue
            $loop++
        } while ((Test-Path $InputFile) -and ($loop -lt 5))
        
        If ($loop -ge 10)
        {
            Write-Host "Unable to rename this file" $InputFule "please manually rename it to" ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
        }
    }
}
else
{
	write-host "Could not find input file :^) " $InputFile
	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Could not find input file " + "`n"
	Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
}

# =============================================================================================================================================

	# Remove the Exchange PowerShell Snap-in
		Remove-PSSnapin *Exchange* -ErrorAction SilentlyContinue
		
	# Retrieve the current Date and Time for use in log files
		$uDate = get-date -uformat %D
		$uTime = get-date -uformat %T
		
		$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"

	# Write end record to the log file
		$LineToWrite = $RecordEvent + "STOP" + "`t" + "This instance is stopping." + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

# End of script #