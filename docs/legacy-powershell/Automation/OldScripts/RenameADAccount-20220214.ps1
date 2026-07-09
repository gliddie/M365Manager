#################################################################################
# 
# PowerShell source code
# Revision v1.0f
# ==========================================================================
#    'Project      : SDAP Menu Development
#    'Description  : Rename the AD Account
#    'Called By    : SDAPAdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 11/26/2017
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 11/26/2017 Created from the Script
#    '             : 11/28/2017 - Check for forwarding and prompt to change; change 
#    '                  SIP address on the Skype Servers if address in use apply our
#    '                  naming rules and made SearchAddr a function
#    '             : 01/10/2018 - Corrected issue with updating SIP Address
#    '             : 03/20/2018 - Added code to allow changes to the AD DisplayName Only
#    '             : 03/23/2018 - Added code to remove spaces from the SMTP and SIP addresses
#    '             : 06/19/2018 - Added code to trim leading and trailing spaces from the name input values and
#    '             :    to make sure the $NewAddr does not have SIP: in front of it.
#    '             : 10/30/2018 - Fixed line $Global:ProposedAddress = $Global:ProposedAddress.Replace(" ","") there
#    '             :    there was no ": in the second instance of the variable name.  Also removed the "SIP:" + when
#    '             :    adding the new SIP address tot he AD account.
#    '             : 07/28/2019 - Added code to update the Target Email address to match the new primary address
#    '             : 10/02/2019 - Added code to set the routing domain if ExtensionAttribute8 is NFP then routing domain is @ul.org
#    '             : 10/29/2919 - Modified report file name to include the individuals employee #
#    '             : 11/03/2020 - Added code so if the agent enters "cr" or "<cr>" it sill be set to blank
#    '             : 04/16/2021 - Modified code for the new NFP org names
#    '             : 05/10/2021 - Modified to replace the old SIP assignment process
#    '             : 06/14/2021 - Modified SIP address code as it was putting in sip:SMTP:Address@ul.com
#    '             : 07/22/2021 - Fixed the change to the SIP address it was still using the old primary email address
#    '             : 07/29/2021 - Modified the lookup to see if the addressis in use
#    '             : 01/27/2022 - Modified code that checks for what groups are included as UL.ORG business groups
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "RenameADAccount"
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

$Global:EAddr = ""
$Global:SrchAddr = ""
$Global:MatchFound = ""
$Global:AddrDet = ""
$Global:PrimAddr = ""

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

Function BuildAddress
{
    If ($SearchMiddle -eq "")
    {
        $Global:ProposedAddress = $NewGivenName + "." + $NewSurName + $RoutingDomain
    }
    else
    {
        $Global:ProposedAddress = $NewGivenName + "." + $SearchMiddle + "." + $NewSurName + $RoutingDomain
    }

    $Global:ProposedAddress = $Global:ProposedAddress.Replace(" ","")
        #$Global:NewAddr = $GlobalProposedAddress
    switch ($Global:SrchCnt)
    {
        "1"
        {
            $Global:ActSrch1 = $Global:ProposedAddress
        }
        "2"
        {
            $Global:ActSrch2 = $Global:ProposedAddress
        }
        "3"
        {
            $Global:ActSrch3 = $Global:ProposedAddress
        }
        "4"
        {
            $Global:ActSrch4 = $Global:ProposedAddress
        }
        "5"
        {
            $Global:ActSrch5 = $Global:ProposedAddress
        }
    }
}

function SearchForAddress
{
    $Global:NewAddr = $Global:ProposedAddress.Replace(" ","")
    write-host "Checking to see if" $Global:NewAddr "is in use" -ForegroundColor Cyan
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Checking to see if " + $Global:NewAddr + " is in use for Emp#" + $UID
    WriteReportEvent
    $Global:Matchfound = "N"
#    $collection = $Global:EAddr.GetEnumerator()
#    $Global:SrchAddr = "*" + $NewAddr + "*"
#    $Global:SrchPrimAddr = "SMTP:" + $NewAddr
#    $Global:SrchAddlAddr = "smtp:" + $NewAddr
#    foreach ($ExistAddr in $ExistAddr)
#    {
        $AddrInUse = [bool](get-mailbox $Global:NewAddr -ErrorAction Silentlycontinue)
        If ($AddrInUse -eq $True)
#        if (($ExistAddr.EmailAddresses -eq $Global:SrchPrimAddr) -or ($ExistAddr.EmailAddresses -eq $Global:SrchAddlAddr))
        {
            If ($Global:MatchFound -eq "N")
            {
                $Global:MatchFound = "Y"
                $Global:AddrDet = get-mailbox $Global:NewAddr -ErrorAction Silentlycontinue
                write-host "The address" $Global:NewAddr "is use by Emp#" $Global:AddrDet.Alias "-" $Global:AddrDet.Name -ForegroundColor Red
                $LineToWrite = $RecordEvent + "INFO" + "`t" + "The Address " + $Global:NewAddr + " is in use by Emp#" + $Global:AddrDet.Alias + " - " + $Global:AddrDet.Name
                WriteReportEvent
            }
        }
#    }
#    $collection.reset()
    $Global:SrchCnt++

    If ($Global:MatchFound -eq "N")
    {
        write-host "The address" $Global:NewAddr "is not in use" -ForegroundColor Yellow
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "The address " + $Global:NewAddr + " is not in use"
        WriteReportEvent
    }
}

function MatchFound
{
    $Global:EAddr = $Usr.ProxyAddresses
    $Global:SIPAddr = ""
    $collection = $Global:EAddr.GetEnumerator()
    foreach ($Global:EAddr in $collection)
    {
        if ($Global:EAddr -clike "*SMTP:*")
        {
            $Global:PrimAddr = $Global:EAddr
        }
        If ($Global:EAddr -like "*sip:*")
        {
            $Global:SIPAddr = $Global:EAddr
        }
    }
    $collection.reset()
}

function NewMiddle
{
    If ($NewMiddle -eq "")
    {
        $Global:NewCN = $NewGivenName + " " + $NewSurName
        $Global:NewAddr = $NewGivenName.Replace(" ","") + "." + $NewSurName.Replace(" ","") + $RoutingDomain
        $Global:NewDisplayName = $NewSurName + ", " + $NewGivenName
    }
    else
    {
        $Global:NewCN = $NewGivenName + " " + $NewMiddle + ". " + $NewSurName
        $Global:NewAddr = $NewGivenName.Replace(" ","") + "." + $NewMiddle.Replace(" ","") + "." + $NewSurName.Replace(" ","") + $RoutingDomain
        $Global:NewDisplayName = $NewSurName + ", " + $NewGivenName + " " + $NewMiddle + "."
    }
}

Function ShowValues
{
    write-host "`n               SwithcMiddle: " $SwitchMiddle
    write-host "                      First: " $NewFirst
    write-host "                Curr Middle: " $CurrMiddle
    write-host "                 New Middle: " $NewMiddle
    Write-Host "                       Last: " $NewSurName
    write-host "               SearchMiddle: " $SearchMiddle
    write-host "                SearchCount: " $Global:SrchCnt " count: (" $MCnt ")"
    write-host "                Act Search1: " $Global:ActSrch1
    write-host "                Act Search2: " $Global:ActSrch2
    write-host "                Act Search3: " $Global:ActSrch3
    write-host "                Act Search4: " $Global:ActSrch4
    write-host "                Act Search5: " $Global:ActSrch5
    write-host "Proposed Search For Address: " $Global:ProposedAddress
}

$DispOnly = "N"
write-host "Running Rename AD Account....."

$Global:ActSrch1 = ""
$Global:ActSrch2 = ""
$Global:ActSrch3 = ""
$Global:ActSrch4 = ""
$Global:ActSrch5 = ""
$Global:SrchCnt = 1
$CurrMiddle = ""
$NewMiddle = ""
$MCnt = 0

write-host "Starting Checks"

$AddrFile = "E:\O365AdminShared\Data\UsrAddresses.csv"

write-host "`nEnter Employee Number of Account to Rename: " -ForegroundColor Cyan -NoNewline
$UID = Read-Host

$ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-EmpNo" + $UID + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"

# Setup Folders and Files	
	CheckLogFiles
	$LineToWrite = "STAR" + "`t" + $FileName + " script has started"
	WriteLogEvent
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	WriteLogEvent
    WriteReportEvent

If ([bool](Get-ADUser -Filter {SamAccountName -eq $UID} -ErrorAction SilentlyContinue))
{
    $ExistAddr = Import-Csv $AddrFile
    $usr = Get-ADUser $UID -Properties GivenName,SurName,Initials,DisplayName,ProxyAddresses,mail,ExtensionAttribute8,msRTCSIP-PrimaryUserAddress
#   $Global:MatchFound = "N"
    $Cont = "Y"

    If (($Usr.ExtensionAttribute8 -like "NFP *") -or ($Usr.ExtensionAttribute8 -eq "Standards") -or ($Usr.ExtensionAttribute8 -eq "Research") -or ($Usr.ExtensionAttribute8 -eq "Support"))
    {
        $RoutingDomain = "@ul.org"
    }
    else
    {
        $RoutingDomain = "@ul.com"
    }
 
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Renaming AD Account for: " + $UID
    WriteReportEvent

    write-host ""
    write-host "Current AD Account DisplayName: " $usr.DisplayName -ForegroundColor Green
    write-host ""
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Current AD Account DisplayName: " + $usr.DisplayName
    WriteReportEvent

    write-host "    Enter New First Name or <cr> for no Change: " -ForegroundColor Cyan -NoNewline
    $NewGivenName = read-host

    If (($NewGivenName -eq "cr") -or ($NewGivenName -eq "<cr>"))
    {
        $NewGivenName = ""
    }

    $NewGivenName = $NewGivenName.Trim(" ") 
    If ($NewGivenName -eq "")
    {
        $NewGivenName = $usr.GivenName
    }
    else
    {
        $NewGivenName = $NewGivenName.Trim(" ")
    }

    write-host "Enter New Middle Initial or <cr> for no Change: " -ForegroundColor Cyan -NoNewline
    $NewMiddle = read-host

    If (($NewMiddle -eq "cr") -or ($NewMiddle -eq "<cr>"))
    {
        $NewMiddle = ""
    }

    If (($NewMiddle -eq "") -and ($usr.Initials -ne $null))
    {
        $RemInit = "N"
        If (($usr.Initials -ne $null) -and ($NewMiddle -ne $null))
        {
            write-host "Do you wish to remove the Middle Initial of" $usr.Initials "from this Account (Y/N)? " -ForegroundColor Red -NoNewline
            $RemInit = read-host
            if ($RemInit -eq "N")
            {
                $NewMiddle = $usr.Initials.substring(0,1)
            }
        }
    }
    else
    {
    #    Middle Initial should not have a period - removing this if one was entered
        $NewMiddle = $NewMiddle.Trim(" ")
        $NewMiddle = $NewMiddle.TrimEnd(".")
    }
    $SearchMiddle = $NewMiddle
#    $Global:AddrDet = get-mailbox $UID -ErrorAction SilentlyContinue

    write-host "     Enter new Last Name or <cr> for no Change: " -ForegroundColor Cyan -NoNewline
    $NewSurName = read-host

    If (($NewSurName -eq "cr") -or ($NewSurName -eq "<cr>"))
    {
        $NewSurName = ""
    }
    
    If ($NewSurName -eq "")
    {
        $NewSurName = $usr.SurName
    }
    else
    {
        $NewSurName = $NewSurName.Trim(" ")
    }

    If (($CurrMiddle.length -gt 0) -and ($CurrMiddle.substring(0,1) -ne "X") -and ($NewMiddle -ne ""))
    {
        $SwitchMiddle = "OtherNewMI"
    }
    else
    {
        If (($CurrMiddle.length -gt 0) -and ($NewMiddle -eq ""))
        {
            $SwitchMiddle = "OtherNoMI"
        }
        else
        {
            If ($CurrMiddle -eq "X")
            {
                $SwitchMiddle = "X"
            }
            else
            {
                $SwitchMiddle = "NoMI"
            }
        }
    }

    NewMiddle

    write-host "Do you need to modify the DisplayName Only (Y/N)? " -ForegroundColor Red -NoNewline
    $DispOnly = Read-Host

    If ($DispOnly -eq "N")
    {
        write-host      
        Write-host "            Proposed New Primary Email Address:" $Global:NewAddr -ForegroundColor Green
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Proposed New Primary Email Address: " + $Global:NewAddr
        WriteReportEvent
        BuildAddress
        SearchForAddress
    
        If ($MatchFound -eq "Y")
        {
            Do
            {
                switch ($SwitchMiddle)
                {
                    "NoMI"
                    {
                        If ($NewMiddle -eq "")
                        {
                            $NewMiddle = "X"
                            $SearchMiddle = $NewMiddle
                        }
                        else
                        {
                            $MCnt++
                            $SearchMiddle = $NewMiddle + $MCnt
                        }
                    }
                    "X"
                    {
                        if ($NewMiddle -eq "")
                        {
                            ""
                        }
                        else
                        {
                            $MCnt++
                            $SearchMiddle = $NewMiddle + $MCnt
                        }
                    }
                    "OtherNewMI"
                    {
                        If (($SearchMiddle -ne "") -and ($Global:SrchCnt -eq 2))
                        {
                            $SearchMiddle = ""
                        }
                        else
                        {
                            If (($NewMiddle -ne "") -and ($CurrMiddle -ne $SearchMiddle) -and ($Global:SrchCnt -gt 2) -and ($MCnt -eq 0))
                            {
                                $NewMiddle = $CurrMiddle
                                $SearchMiddle = $CurrMiddle
                            }
                            else
                            {
                                If ($NewMiddle -eq $CurrMiddle)
                                {
                                    $MCnt++
                                    $SearchMiddle = $NewMiddle + $MCnt
                                }
                            }
                        }
                    }
                    "OtherNoMI"
                    {
                        If ($NewMiddle -eq "")
                        {
                            $NewMiddle = $CurrMiddle
                            $SearchMiddle = $NewMiddle
                        }
                        else
                        {
                            $MCnt++
                            $SearchMiddle = $NewMiddle + $MCnt
                        }
                    }
                }
        
                If ($Global:MatchFound -eq "Y")
                {
                    BuildAddress
    #                ShowValues
                    $Global:EAddr = "*" + $Global:NewAddr + "*"
                    SearchForAddress
                    $Cont = "N"
                    If ($Global:MatchFound -eq "Y")
                    {
                        write-host "The email address is already used by another account applying naming rules to determine an alternate address to use" -ForegroundColor Red
                    }
                }
    
            } while ($Global:Matchfound -eq "Y")
        }

        write-host "`nDo you want to continue Renaming this Account with this address " $Global:NewAddr " (Y/N)? " -ForegroundColor red -NoNewline
        $Cont = Read-Host
    }
    else
    {
        write-host "Only the DisplayName will be modified as requested." $Global:NewAddr -Foregroundcolor Yellow
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Only the DisplayName will be modified as requested. "
        WriteReportEvent
    }

    If ($Cont -eq "Y")
    {
        NewMiddle
        write-host
        If ($DispOnly -ne "Y")
        {
            write-host "Renaming Account this will be the new Primary Email Address: " $Global:NewAddr -Foregroundcolor Yellow
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Renaming Account this will be the new Primary Email Address: " + $Global:NewAddr
        }
        else
        {
            write-host "Renaming Account this will be the new DisplayName: " $Global:NewDisplayName -Foregroundcolor Yellow
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Renaming Account this will be the new DisplayName: " + $Global:NewDisplayName
        }
        WriteReportEvent

        if ($newSurName -ne "")
        {
            Set-ADUser -Identity $usr -Surname $NewSurName
        }

        if ($newGivenName -ne "")
        {
            Set-ADUser -Identity $Usr -GivenName $newGivenName
        }

        if (($NewMiddle -eq "") -and ($usr.Initials -ne ""))
        {
            Set-ADUser -Identity $usr -Initials $null
        }
        else
        {
            If ($NewMiddle -ne $usr.Initials)
            {
                Set-ADUser -Identity $usr -Initials $NewMiddle
            }
        }

        if ($usr.DisplayName -ne $Global:NewDisplayname)
        {
            Set-ADUser -Identity $usr -DisplayName $Global:NewDisplayName
        }

        write-host "                    New AD Account DisplayName:" $Global:NewDisplayName -ForegroundColor Green
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "New AD Account DisplayName: " + $Global:NewDisplayName
        WriteReportEvent

        If (($Global:MatchFound -eq "N") -and ($DispOnly -ne "Y"))
        {
            MatchFound
            
            Set-ADUser -Identity $Usr -EMailAddress $Global:NewAddr
            write-host "                    New Email Address Set to:" $Global:NewAddr -ForegroundColor Green
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "New Email Address Set to: " + $Global:NewAddr
            WriteReportEvent
            $DemoteAddr = $Global:PrimEAddr -replace ("SMTP:","smtp:")
			
			Set-ADUser -Identity $usr -Replace @{targetaddress=$Global:NewAddr}
            write-host "                   New Target Address Set to:" $Global:NewAddr -ForegroundColor Green
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "New Target Address Set to: " + $Global:NewAddr
            WriteReportEvent			
        
            Set-ADUser -Identity $usr -Add @{ProxyAddresses="SMTP:" + $Global:NewAddr}
            write-host "            New Primary Email Address Set to:" $Global:NewAddr -ForegroundColor Green
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "New Primary Email Address Set to: " + $Global:NewAddr
            WriteReportEvent
        
            Set-AdUser -Identity $usr -Remove @{ProxyAddresses=$Global:PrimAddr}
            write-host "    Demoting old Primary Address to an alias:" $Global:PrimAddr -ForegroundColor Green
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Demoting old Primary Address to an additional Alias: " + $Global:PrimAddr
            WriteReportEvent
        
            $DemoteAddr="smtp:" + $usr.mail            
            Set-AdUser -Identity $usr -Add @{ProxyAddresses=$DemoteAddr}
        
            if ($Usr."msRTCSIP-PrimaryUserAddress" -ne "")
            {
                Set-ADUser $Usr -Replace @{'msRTCSIP-PrimaryUserAddress'="sip:" + $Global:NewAddr}
                write-host "                     Current SIP Address: $($Usr."msRTCSIP-PrimaryUserAddress")" -ForegroundColor Green
                $LineToWrite = $RecordEvent + "INFO" + "`t" + "Current SIP Address:" + $Usr."msRTCSIP-PrimaryUserAddress"
                WriteReportEvent
                write-host "                    Updating SIP Address: $("sip:" + $Global:NewAddr)" -ForegroundColor Green
                $LineToWrite = $RecordEvent + "INFO" + "`t" + "Updating SIP Address:" + $("sip:" + $Global:NewAddr)
                WriteReportEvent
            }

            If ($SIPAddr -ne "")
            {
                write-host "Removing old SIP Address from Proxy Addr:" $SIPAddr -ForegroundColor Green
                $LineToWrite = $RecordEvent + "INFO" + "`t" + "Removing old SIP Address from Proxy Addr: " + $SIPAddr
                WriteReportEvent
            }
          
            $Global:AddrDet = get-mailbox $UID -ErrorAction SilentlyContinue

            if ($Global:AddrDet.ForwardingSMTPAddress -ne $null)
            {
                $ForwAddr = ""
                write-host "Found a forwarding address of " $Global:AddrDet.ForwardingSMTPAddress " configured on this account enter (C) to Change and (R) to Remove? " -ForegroundColor Cyan -NoNewline
                $ForwAddr = read-host

                switch ($ForwAddr)
                {
                    "C"
                    {
                        write-host "Enter the new forwarding address? " -ForegroundColor Cyan -NoNewline
                        $NewForwAddr = Read-Host
                        set-mailbox $Global.AddrDet.ForwardingSMTPAddress $NewForwAddr
                    }
                    "R"
                    {
                        set-mailbox $Global.AddrDet.ForwardingSMTPAddress $null
                    }
                }
            }
            Add-Content $AddrFile ("SMTP:" + $Global:NewAddr)
        }
        else
        {
            write-host "No changes to the SIP or SMTP addressses this request was to modify the DisplayName for this user has been modified." -ForegroundColor Cyan
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "No changes to the SIP or SMTP addressses this request was to modify the DisplayName for this user has been modified."
            WriteReportEvent
        }

        ## This performs the final step in the process of renaming the AD user account

        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Completing AD Account Rename modifying Distinguished Name from: " + $Usr.DistinguishedName
        WriteReportEvent
        Rename-ADObject $Usr.DistinguishedName -NewName $Global:NewCN

        write-host "AD Account Rename Complete"
    }
    else
    {
        write-host "Rename Process Aborted as requested by Agent"
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Rename Process Aborted as requested by Agent for: " + $UID
        WriteReportEvent
    }
}
else
{
    write-host "This User does not Exist: " $UID -ForegroundColor Red
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "This User does not Exist: " + $UID
    WriteReportEvent
}

# Write end record to the log file
    $LineToWrite = $RecordEvent + "STOP" + "`t" + "This instance is stopping."
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

# End of script #  
