Function GetUserDN($strDN)
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

Function GetAcctInfo($strUID)
{
    $ADExists = [bool]($strDN = GetUserDN $Contact.EmployeeID -ErrorAction SilentlyContinue)
    If ($ADExists -eq "True")
    {
       $strDN = GetUserDN $Contact.EmployeeID
       $strUserPath = [string]::format("LDAP://{0}", $strDN)
       $u = new-object System.DirectoryServices.DirectoryEntry($strUserPath) 
    }
    else
    {
        Write-Host ""
        Write-Host "Active Directory Account Does Not Exist for Emp#" $Contact.EmployeeID -ForegroundColor Red
    }
}            

            $Contact = "42657"
 #           $strUID = get-ADUser $Contact -Properties sn,givenName,initials,displayName,mail,proxyaddresses
            GetAcctInfo

	        If ($ux.mail.value -ne "")
		    {
                write-host "Removing Email Address" $ux.mail.value "from the AD Account" $Contact
                $ux.mail.value = $null
                $ux.CommitChanges()
#                $LineToWrite = $RecordEvent + "INFO" + "`t" + $Name + "`t" + "Removing " + $strUID.Mail + " from " + $Contact + "`n"
#    			Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            }
            else
            {
                write-host "There is no value in the AD mail string"
            }