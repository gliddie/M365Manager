#################################################################################
# 
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange
#    'Description  : Checks Legal Hold Status and allow to Enable/Disable
#    'Called By    : LegalHoldCheck.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 01/23/2018
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 01/23/2018 Created from Function Call        
#
# ==========================================================================
#
#################################################################################
Function ShowDetails
{
    write-host "Employee Number               : " $Global:inf.Alias -ForegroundColor Green
    write-host "DisplayName                   : " $Global:inf.DisplayName -ForegroundColor Green
    write-host "Employee Type                 : " $Global:inf.CustomAttribute1 -ForegroundColor Green
    write-host "AD Container                  : " $Global:strUserPath -ForegroundColor Green
    write-host "LitigationHoldEnabled         : " $Global:inf.LitigationHoldEnabled -ForegroundColor Green
    write-host "LitigationHoldDate            : " $Global:inf.LitigationHoldDate -ForegroundColor Green
    write-host "LigitationHoldOwner           : " $Global:inf.LitigationHoldOwner -ForegroundColor Green
    write-host "Hidden from Address Book O365 : " $Global:inf.HiddenFromAddressListsEnabled -Foregroundcolor Green
    write-host "Hidden from Address Book AD   : " $Global:u.msExchHideFromAddressLists.value -Foregroundcolor Green
    write-host "O365 Retention Comment        : " $Global:inf.RetentionComment -ForegroundColor Green
    write-host "AD Retention Comment          : " $ADCmt -ForegroundColor Green
    write-host "AD Object Protected           : " $UsrDetails.ProtectedFromAccidentalDeletion -ForegroundColor Green
    $Global:MbxAccess = Get-MailboxPermission $ENo
    PrtMBxAccess($Global:MbxAccess)
}


    $ADExists = [bool]($strDN = GetUserDN $ENo -ErrorAction SilentlyContinue)
    If ($ADExists -eq "True")
    {
       $UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion
     }
    else
    {
        Write-Host ""
        Write-Host "Active Directory Account Does Not Exist for Emp#" $ENo -ForegroundColor Red
    }
    $MbxExists = [bool](get-mailbox -identity $ENo -ErrorAction SilentlyContinue)
    If ($MbxExists -eq "True")
    {
        $Global:inf = get-mailbox $ENo
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
        ShowDetails($ENo)

        $LineToWrite = $WhoAmI + "`t" + "EmployeeNumber                 :" + "`t" + $Global:inf.Alias
        WriteReportEvent
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
        $LineToWrite = $WhoAmI + "`t" + "Hidden from Address Book O365  :" + "`t" + $Global:inf.HiddenFromAddressListsEnabled
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "Hidden from Address Book AD    :" + "`t" + $Global:u.msExchHideFromAddressLists.value
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "O365 Retention Comment         :" + "`t" + $Global:inf.RetnetionComment
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "AD Retention Comment           :" + "`t" + $ADCmt
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "AD Object Protected            :" + "`t" + $UsrDetails.ProtectedFromAccidentalDeletion
        WriteReportEvent
    }
    else
    {
        write-host""
        write-host "Mailbox does not exist for Emp#" $ENo -ForegroundColor Red
    }
