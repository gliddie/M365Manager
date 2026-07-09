#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange
#    'Description  : Remove the Email Address set for User with Mail Contact
#    'Called By    : O365AdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 5/26/2016
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 05/26/2016 Created Script        
#    '             : 
#    '
# ==========================================================================
#
#################################################################################

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
}

function WriteLogEvent
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LocalMachineName + "`t" + $LineToWrite
	Out-File -filepath $LogFile -append -noClobber -inputObject $RecordEvent
}

function WriteReportEvent 
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LineToWrite
	Out-File -filepath $ReportFile -append -noClobber -inputObject $RecordEvent
}

Function GetUserDN($strUID)
{
#####   This function connects to Active Directory and gets the record for the user 
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = "LDAP://DC=global,DC=ul,DC=com"
    $objSearcher.Filter = [string]::format("(sAMAccountName={0})",$strUID)
    $ux = $null
    $ux = $objSearcher.FindOne()
    
    if ($ux -eq $null)
    {
        return $null
    } 
    else
    {
        return $ux.Properties.distinguishedname
    }
}

Function GetAcctInfo($ENo)
{
    $ADExists = [bool]($strDN = GetUserDN $ENo -ErrorAction SilentlyContinue)
    If ($ADExists -eq "True")
    {
       $strDN = GetUserDN $ENo
       $strUserPath = [string]::format("LDAP://{0}", $strDN)
       $u = new-object System.DirectoryServices.DirectoryEntry($strUserPath) 
    }
    else
    {
        Write-Host ""
        Write-Host "Active Directory Account Does Not Exist for Emp#" $ENo -ForegroundColor Red
    }
}

Function RemoveEmail($ENo)
{
#####  This function handles removing the Email Address from the AD Account  #######

        if ($u.mail.value -ne $null)
        {
            $u.mail.Clear()
			$u.CommitChanges()
		}
 }

#################################################################################
# Declare Drive | Folders | and Files
	$FileName		= "RemoveEmailforMailContacts"
	$LogDrive		= "E:"
	$LogPath		= "\Automation"
	$LogFolder		= "\" + $FileNAme
	$LogDirectory	= $LogDrive + $LogPath + $LogFolder + "\"
	$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	$InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
	$ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
# Retrieve the user name
	$WhoAmI			= WhoAmI
		
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name

# Setup Folders and Files	
	CheckLogFiles
	$LineToWrite = "STAR" + "`t" + $FileName + " script has started"
	WriteLogEvent
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	WriteLogEvent
    WriteReportEvent
	
invoke-expression -Command .\ConnectO365.ps1

write-host "Starting Removing Email from Individuals Created with Mail Contact Records " -ForegroundColor Magenta
write-host
write-host
write-Host "Enter Employee Number (or 0 to Exit) : " -ForegroundColor Green -NoNewline
$ENo = Read-Host
$Act = 1

	while ($Act -gt 0)
	{
        $ADExists = [bool]($strDN = GetUserDN $ENo -ErrorAction SilentlyContinue)
        If ($ADExists -eq "True")
        {
           $strDN = GetUserDN $ENo 
        }
        else
        {
            Write-Host ""
            Write-Host "Active Directory Account Does Not Exist for Emp#" $ENo -ForegroundColor Red
        }
        
		$strUserPath = [string]::format("LDAP://{0}", $strDN)
        $u = new-object System.DirectoryServices.DirectoryEntry($strUserPath)
        If ($strDN -eq $null)
        {
            $ADCmt = "*****No Active Directory Account For This User*****"
            $strUserPath = "No Active Directory Account for this User"
        }
        else
		{
            $ADMail = $u.Mail.value
        }

 		write-host ""
		Write-Host "     Remove" $u.mail.value " from account " $strDN "(Y/N)? " -ForegroundColor Red -NoNewline
		$Act = Read-Host
		write-host ""
        		
		switch ($act)
		{
            "Y"
                {
                    write-host "Removing Email Address from the AD Account"

					$u.mail.value = $null
                    $u.CommitChanges()
				}
        }
}	