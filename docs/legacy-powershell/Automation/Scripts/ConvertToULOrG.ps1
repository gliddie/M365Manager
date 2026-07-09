#
#   This script moves individuals from ul.com to ul.org addresses
#

write-host "Enter Employee Number to Move to UL.ORG: " -ForegroundColor cyan -NoNewline
$uid = read-host

Do
{
    $EmpID = $uid + "@global.ul.com"
    $usr = Get-ADUser $UID -Properties ProxyAddresses,mail,TargetAddress

	$NewAddr = $usr.mail.Replace("ul.com","ul.org")
    $TarAddr = $usr.TargetAddress.Replace("ul.com","ul.org")
	if ($usr.ProxyAddresses -clike "sip:*")
    {
        $SIPAddr = $TarAddr.Replace("SMTP:","sip:")
    }
    else
    {
        $SIPAddr = $TarAddr.Replace("SMTP:","SIP:")
    }
	
    write-host "                Employee ID: " $EmpID
    write-host "              Employee Name: " $usr.Name
    write-host "       Current Mail Address: " $Usr.mail
    write-host "         New UL.ORG Address: " $NewAddr
    write-host "     Current Target Address: " $usr.TargetAddress
    write-host "         New Target Address: " $TarAddr
    write-host "            New SIP Address: " $SIPAddr
    write-host "        All Proxy Addresses: " $usr.ProxyAddresses
    
    write-host "Is the the correct user (Y/N)? " -ForegroundColor yellow -NoNewline
    $cont = read-host

    If (($usr.mail -notlike "*ul.org*") -and ($cont-eq "Y"))
    {
        Set-ADUser -Identity $usr -EMailAddress $NewAddr
	    $AddAddr = "SMTP:"+$NewAddr
        Set-ADUser -Identity $usr -Add @{ProxyAddresses=$AddAddr}
	    Set-AdUser -Identity $usr -Remove @{ProxyAddresses=$usr.TargetAddress}
	    Set-ADUser -Identity $usr -Replace @{targetaddress=$TarAddr}
        $OldAddr = $usr.TargetAddress.Replace("SMTP:","smtp:")
        Set-ADUser -Identity $usr -Add @{ProxyAddresses=$OldAddr}
        sleep -seconds 5
        Set-CSUser –Identity $EmpID –SIPAddress $SIPAddr
        sleep -Seconds 5
        $usr = Get-ADUser $UID -Properties ProxyAddresses,mail,TargetAddress
                write-host "`nNew Address: " $usr.mail
        write-host "New Proxy Addresses: " $usr.ProxyAddresses
        write-host " New Target Address: " $usr.TargetAddress
    }
    else
    {
        If ($cont -eq "N")
        {
            write-host "Cancelling request as requested"
        }
        else
        {
            write-host "This user already has ul.org as the primary email address"
        }
    }
    write-host "Enter Employee Number to Move to UL.ORG (0) to exit: " -ForegroundColor cyan -NoNewline
    $uid = read-host
}while ($uid -ne "0")