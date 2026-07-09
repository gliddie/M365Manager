#################################################################################
# 
# PowerShell source code
# Revision v1.0f
# ==========================================================================
#    'Project      : SDAP Menu Development
#    'Description  : Move Accounts to/from UL.COM and UL.ORG email addresses
#    'Called By    : SDAPAdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 08/28/2019
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 08/28/2019 Created from ConvertULORG.ps1 Script

# Declare Drive | Folders | and Files
	$FileName		= "EmailDomainChange"
	$LogDrive		= "E:"
	$LogPath		= "\SDAP"
	$LogFolder		= "\" + $FileNAme
	$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
	$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	$ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
# Retrieve the user name
	$WhoAmI			= WhoAmI
		
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name

function ShowProxy
{
    $cnt = 0
    foreach ($PrxyAddr in $usr.ProxyAddresses)
    {
        if ($cnt -eq 0)
        {
            write-host "        All Proxy Addresses: " $PrxyAddr
            $LineToWrite = "INFO" + "`t" + "   All Proxy Addresses:  " + $PrxyAddr
            $cnt++
        }
        else
        {
            write-host "                             " $PrxyAddr
            $LineToWrite = "INFO" + "`t" + "                        " + $PrxyAddr
        }
        invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1
    }
}

write-host "Running Convert Email Domain....."

# Setup Folders and Files	
	invoke-expression -Command E:\O365AdminShared\Scripts\CheckLogFiles.ps1
	$LineToWrite = "STAR" + "`t" + $FileName + " script has started"
	invoke-expression -Command E:\O365AdminShared\Scripts\WriteEventLog.ps1
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	invoke-expression -Command E:\O365AdminShared\Scripts\WriteEventLog.ps1
    
write-host "Enter Employee Number for email domain change: " -ForegroundColor cyan -NoNewline
$uid = read-host

Do
{
    $ReportFile	= $LogDirectory + "Report\Report-" + $FileName + "-EmpNo" + $uid + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
    $LineToWrite = "INFO" + "`t" + "Converting Email Addresses for: " + $UID + "`n"
    invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1
    $EmpID = $uid + "@global.ul.com"
    $usr = Get-ADUser $UID -Properties ProxyAddresses,mail,TargetAddress

    If ($usr.mail -like "*ul.com*")
    {
        write-host "This user will be moved from the UL.COM email domain to the UL.ORG mail domain." -ForegroundColor Green
        $LineToWrite = "INFO" + "`t" + "This user will be moved from the UL.COM email domain to the UL.ORG mail domain "
        invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1
    }
    else
    {
        write-host "This user will be moved from the UL.ORG email domain to the UL.COM mail domain." -ForegroundColor Magenta
        $LineToWrite = "INFO" + "`t" + "This user will be moved from the UL.ORG email domain to the UL.COM mail domain "
        invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1
    }

    write-host "              Employee Name: " $usr.Name
    $LineToWrite = "INFO" + "`t" + "         Employee Name: " + $usr.Name
    invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1
    write-host "       Current Mail Address: " $Usr.mail
    $LineToWrite = "INFO" + "`t" + "  Current Mail Address: " + $Usr.mail
    invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1

	switch -Wildcard ($usr.mail)
    {
        "*ul.com"
        {
            $NewAddr = $usr.mail.Replace("ul.com","ul.org")
            $TarAddr = $usr.TargetAddress.Replace("ul.com","ul.org")
            write-host "         New UL.ORG Address: " $NewAddr
            $LineToWrite = "INFO" + "`t" + "    New UL.ORG Address: " + $NewAddr
        }
        "*ul.org"
        {
            $NewAddr = $usr.mail.Replace("ul.org","ul.com")
            $TarAddr = $usr.TargetAddress.Replace("ul.org","ul.com")
            write-host "         New UL.COM Address: " $NewAddr
            $LineToWrite = "INFO" + "`t" + "    New UL.COM Address: " + $NewAddr
        }
    }
	if ($PrxyAddr -clike "sip:*")
	{
		$SIPAddr = $TarAddr.Replace("SMTP:","sip:")
	}
	else
	{
		$SIPAddr = $TarAddr.Replace("SMTP:","SIP:")
	}
    invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1
    write-host "     Current Target Address: " $usr.TargetAddress
    $LineToWrite = "INFO" + "`t" + "Current Target Address: " + $usr.TargetAddress
    invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1
    write-host "         New Target Address: " $TarAddr
    $LineToWrite = "INFO" + "`t" + "    New Target Address: " + $TarAddr
    invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1
    ShowProxy

    write-host "Is the the correct user (Y/N)? " -ForegroundColor yellow -NoNewline
    $cont = read-host

    If ($Cont -eq "Y")
    {
        write-host "`nAddress Details After Changes" -ForegroundColor Green
        $LineToWrite = "`nINFO" + "`t" + "Address Details After Changes "
        invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1

        Set-ADUser -Identity $usr -EMailAddress $NewAddr
	    $AddAddr = "SMTP:"+$NewAddr
        Set-ADUser -Identity $usr -Add @{ProxyAddresses=$AddAddr}
	    Set-AdUser -Identity $usr -Remove @{ProxyAddresses=$usr.TargetAddress}
	    Set-ADUser -Identity $usr -Replace @{targetaddress=$TarAddr}
        $OldAddr = $usr.TargetAddress.Replace("SMTP:","smtp:")
        Set-ADUser -Identity $usr -Add @{ProxyAddresses=$OldAddr} -ErrorAction SilentlyContinue
        sleep -seconds 5
		write-host "Would you like to change the SIP Address now (Y/N)? " -ForegroundColor yellow -NoNewline
		$ChgSIP = read-host
		If ($ChgSIP -eq "Y")
		{
			Set-CSUser –Identity $usr –SIPAddress $SIPAddr
			write-host "`n***  SIP Address updated as requested. ****" -ForegroundColor Red
			$LineToWrite = "`nINFO" + "`t" + "   SIP Address updated as requested"
		}
		else
		{
			write-host "`n***  SIP Address Must be manually changed to the new domain.  Do this once you are in contact with the user. ****" -ForegroundColor Red
			$LineToWrite = "`nINFO" + "`t" + "   SIP Address Must be changed when agent is ready to transition user"
		}
		invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1
        $usr = Get-ADUser $UID -Properties ProxyAddresses,mail,TargetAddress
        write-host "`n        New Primary Address: " $usr.mail
        $LineToWrite = "INFO" + "`t" + "   New Primary Address: " + $usr.mail
        invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1
		sleep -Seconds 5
        ShowProxy		
    }
    else
    {
            write-host "No changes being made for this individual"
            $LineToWrite = "INFO" + "`t" + "Agent cancelled changes for this individual"
            invoke-expression -Command E:\O365AdminShared\Scripts\WriteReportEvent.ps1
    }
    write-host "Enter Employee Number of individual for email domain change (0) to exit: " -ForegroundColor cyan -NoNewline
    $uid = read-host
}while ($uid -ne "0")