################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange
#    'Description  : Get AD User Account Info
#    'Called By    : EmergencyDisablement.ps1, StandardTermination.ps1,
#                    EmergencyReHire.ps1
#    'Calls        :  
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 11/8/2017
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 11/08/2017 Created from code that was duplicated in these scripts
#    '             : 10/31/2018 Added logging of CustomAttribute2 (Status Type)
#    '             : 04/08/2020 Added details of who the users manager is
#    '             : 06/23/2020 Changed Mailbox Status to Mailbox Disabled
#    '             : 09/21/2021 Commented out the UMMailbox process that is no longer valid.
# ==========================================================================
#
#################################################################################

Function GetUserDN($Global:strDN)
{
#####   This function connects to Active Directory and gets the record for the user 
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = "LDAP://DC=global,DC=ul,DC=com"
    $objSearcher.Filter = [string]::format("(sAMAccountName={0})",$Global:strDN)
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

Function GetAcctInfo($Global:ENo)
{
    $ADExists = [bool]($strDN = GetUserDN $Global:ENo -ErrorAction SilentlyContinue)
    If ($ADExists -eq "True")
    {
       $UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion
     }
    else
    {
        Write-Host ""
        Write-Host "Active Directory Account Does Not Exist for Emp#" $Global:ENo -ForegroundColor Red
    }

    $Global:inf = get-mailbox $Global:ENo -ErrorAction SilentlyContinue
    $Global:cont = get-mailcontact $Global:ENo -ErrorAction SilentlyContinue
    $Global:strUserPath = [string]::format("LDAP://{0}", $strDN)
    $Global:u = new-object System.DirectoryServices.DirectoryEntry($Global:strUserPath)
    If ($strDN -eq $null)
    {
        $ADCmt = "*****No Active Directory Account For This User*****"
        $Global:strUserPath = "No Active Directory Account for this User"
    }
    else
    {
        $ADCmt = $Global:u.ExtensionAttribute14.value
    }
    ShowDetails($Global:ENo)
}

Function PrtMbxAccess($Global:MbxAccess)
{
    foreach ($Global:MbxAccess in $Global:MbxAccess)
    {
        if ($Global:MbxAccess.User -like "*@global*")
        {
            write-host "Accounts Granted Access       : " $Global:MbxAccess.User "- granted" $Global:MbxAccess.AccessRights -ForegroundColor Yellow
            $LineToWrite = $WhoAmI + "`t" + "Accounts Granted Access        :" + "`t" + $Global:MbxAccess.User + " - granted " + $Global:MbxAccess.AccessRights
            WriteReportEvent
        }
    }
}

Function ShowDetails
{
    $CASMbx = Get-CASMailbox $Global:ENo
    $Usr = $Usr = Get-MsolUser -UserPrincipalName ($Global:ENo + "@global.ul.com") -ErrorAction SilentlyContinue
#    $UMMbx = Get-UMMailbox -Identity $ENo -ErrorAction SilentlyContinue

    write-host "DisplayName                   : " $Global:u.CN -ForegroundColor Green
    write-host "Employee Type                 : " $Global:u.ExtensionAttribute1.value -ForegroundColor Green
    write-host "Employee Region               : " $Global:u.ExtensionAttribute5.value -ForegroundColor Green
    write-host "AD Account Description        : " $Global:u.Description.value -ForegroundColor Green
    write-host "AD Container                  : " $Global:strUserPath -ForegroundColor Green
    write-host "AD Account Enabled            : " $ADU.Enabled -ForegroundColor Green
    write-host "AD User Manager               : " $Global:u.Manager.Value -ForegroundColor Green
	write-host "AD ExtensionAttribute2        : " $Global:u.extensionAttribute2.value -Foregroundcolor Green
    write-host "AD ExtensionAttribute14       : " $Global:u.extensionAttribute14.value -Foregroundcolor Green
    write-host "AD Object Protected           : " $UsrDetails.ProtectedFromAccidentalDeletion -ForegroundColor Green
    write-host "AD Hidden from Address Book   : " $Global:u.msExchHideFromAddressLists.value -ForegroundColor Green
    if ($Global:inf -ne "")
    {
        write-host "Mailbox Status                : " $Global:inf.AccountDisabled -ForegroundColor Green
        write-host "Hidden from O365 Address Book : " $Global:inf.HiddenFromAddressListsEnabled -ForegroundColor Green
        write-host "LitigationHoldEnabled         : " $Global:inf.LitigationHoldEnabled -ForegroundColor Green
        write-host "OWA Enabled                   : " $CASMbx.OWAEnabled -ForegroundColor Green
        write-host "OWA For Devices Enabled       : " $CASMbx.OWAforDevicesEnabled -ForegroundColor Green
        write-host "ActiveSync Enabled            : " $CASMbx.ActiveSyncEnabled -ForegroundColor Green
        write-host "OutlookMAC Enabled            : " $CASMbx.EwsAllowMacOutlook -ForegroundColor Green
#        write-host "UM Mailbox                    : " $UMMbx.Extensions -ForegroundColor Green
        $Global:MbxAccess = Get-MailboxPermission $Global:ENo -ErrorAction SilentlyContinue| where {$_.User -like "*ul.com*"}
    }
    if ($Global:cont -ne "")
    {
        write-host "Mailbox Contact Forwarder     : " $Global:cont.ExternalEmailAddress -ForegroundColor Green
    }
    write-host "STS Last Token Refresh Date   : " $Usr.StsRefreshTokensValidFrom -ForegroundColor Green

    WriteDetails

    If ($Global:MbxAccess -ne $null)
    {
        PrtMbxAccess($Global:MbxAccess)
    }
}

Function WriteDetails
{
    $LineToWrite = $WhoAmI + "`t" + "DisplayName                    :" + "`t" + $Global:u.CN
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "Employee Type                  :" + "`t" + $Global:u.ExtensionAttribute1.value
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "Status Type                    :" + "`t" + $Global:u.extensionAttribute2.value
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "Employee Region                :" + "`t" + $Global:u.ExtensionAttribute5.value
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "AD Account Description         :" + "`t" + $Global:u.Description.value
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "AD Container                   :" + "`t" + $Global:strUserPath
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "AD Account Enabled             :" + "`t" + $ADU.Enabled
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "AD User Manager                :" + "`t" + $Global:u.Manager.value
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "AD ExtensionAttribute14        :" + "`t" + $Global:u.extensionAttribute14.value
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "AD Object Protected            :" + "`t" + $UsrDetails.ProtectedFromAccidentalDeletion
    WriteReportEvent
    $LineToWrite = $WhoAmI + "`t" + "AD Hidden from Address Book    :" + "`t" + $Global:u.msExchHideFromAddressLists.value
    WriteReportEvent

    if ($Global:inf -ne "")
    {
        $LineToWrite = $WhoAmI + "`t" + "Mailbox Disabled               :" + "`t" + $Global:inf.AccountDisabled
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "Hidden from O365 Address Book  :" + "`t" + $Global:inf.HiddenFromAddressListsEnabled
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "LitigationHoldEnabled          :" + "`t" + $Global:inf.LitigationHoldEnabled
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "OWA Enabled                    :" + "`t" + $CASMbx.OWAEnabled
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "OWA For Devices Enabled        :" + "`t" + $CASMbx.OWAforDevicesEnabled
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "ActiveSync Enabled             :" + "`t" + $CASMbx.ActiveSyncEnabled
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "OutlookMAC Enabled             :" + "`t" + $CASMbx.EwsAllowMacOutlook
        WriteReportEvent
#        $LineToWrite = $WhoAmI + "`t" + "UM Mailbox                     :" + "`t" + $UMMbx.Extensions
#        WriteReportEvent
    }
    If ($Global:cont -ne "")
    {
        $LineToWrite = $WhoAmI + "`t" + "Mailbox Contact Forwarder      :" + "`t" + $Global:cont.ExternalEmailAddress
        WriteReportEvent
    }
    $LineToWrite = $WhoAmI + "`t" + "STS Last Token Refresh Date    :" + "`t" + $Usr.StsRefreshTokensValidFrom
    WriteReportEvent
}

    GetAcctInfo($ENo)