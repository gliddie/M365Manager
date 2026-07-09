#
# PowerShell source code
# Revision v1.0
# ==========================================================================
#    'Project      : UL Office 365 Exchange
#    'Description  : Creates a report of all mailboxes for a preservatione
#    'Called By    : LegalHoldCheck.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 01/23/2018
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 08/07/2015 Integrated to be called from O365AdminMenu
#                  : 02/19/2018 Removed the "AD Account Does not Exist" so it only
#                       prints if the account is on that preservation activity        
#
# ==========================================================================
#
#################################################################################
#
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

$ReportFile		= "c:\temp\LegalHoldReport.txt"

write-host "Enter Name of Preservation Activity: " -ForegroundColor Green -NoNewline
$PresvName = Read-Host
$ReportFile		= "c:\temp\" + $PresvName + "LegalHoldReport.txt"
$PresvName = "*" + $PresvName + "*"

$PreservationMbx = get-mailbox -ResultSize Unlimited |Where-Object {($_.LitigationHoldEnabled -eq "True")} |Sort-Object Alias

write-host "Number of Mailboxes on Legal Hold: " $PreservationMbx.count

$cnt = 0

foreach ($PreservationMbx in $PreservationMbx)
{
    $ENo = $PreservationMbx.Alias

    $ADExists = [bool]($strDN = GetUserDN $ENo -ErrorAction SilentlyContinue)
    If ($ADExists -eq "True")
    {
       $UsrDetails = Get-ADObject -Identity  $strDN -Properties ProtectedFromAccidentalDeletion
     }
#    else
#    {
#        Write-Host ""
#        Write-Host "Active Directory Account Does Not Exist for Emp#" $ENo -ForegroundColor Red
#    }
    $MbxExists = [bool](get-mailbox -identity $ENo -ErrorAction SilentlyContinue)
    If ($MbxExists -eq "True")
    {
        $Global:inf = get-mailbox $ENo
        if ($ADExiSts -ne "True")
        {
            $ADCmt = "*****No Active Directory Account For This User*****"
            $Global:strUserPath = "*****No Active Directory Account for this User*****"
        }
        else
        {
            $Global:strUserPath = [string]::format("LDAP://{0}", $strDN)
            $Global:u = new-object System.DirectoryServices.DirectoryEntry($Global:strUserPath)
            $ADCmt = $Global:u.ExtensionAttribute14.value
        }

        If (($ADCmt -like $PresvName) -or ($Global:inf.RetentionComment -like $PresvName))
        {
            write-host "`nEmployee Number               : " $Global:inf.Alias -ForegroundColor Green
            $LineToWrite = "`nEmployee Number                :" + "`t" + $Global:inf.Alias
            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            write-host "DisplayName                   : " $Global:inf.DisplayName -ForegroundColor Green
            $LineToWrite = "DisplayName                    :" + "`t" + $Global:inf.DisplayName
            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            write-host "Employee Type                 : " $Global:inf.CustomAttribute1 -ForegroundColor Green
            $LineToWrite = "Employee Type                  :" + "`t" + $Global:inf.CustomAttribute1
            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            write-host "AD Container                  : " $Global:strUserPath -ForegroundColor Green
            $LineToWrite = "AD Container                   :" + "`t" + $Global:strUserPath
            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            write-host "LitigationHoldEnabled         : " $Global:inf.LitigationHoldEnabled -ForegroundColor Green
            $LineToWrite = "LitigationHoldEnabled          :" + "`t" + $Global:inf.LitigationHoldEnabled
            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            write-host "LitigationHoldDate            : " $Global:inf.LitigationHoldDate -ForegroundColor Green
            $LineToWrite = "LitigationHoldDate             :" + "`t" + $Global:inf.LitigationHoldDate
            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            write-host "LigitationHoldOwner           : " $Global:inf.LitigationHoldOwner -ForegroundColor Green
            $LineToWrite = "LigitationHoldOwner            :" + "`t" + $Global:inf.LitigationHoldOwner
            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            write-host "Hidden from Address Book O365 : " $Global:inf.HiddenFromAddressListsEnabled -Foregroundcolor Green
            $LineToWrite = "Hidden from Address Book O365  :" + "`t" + $Global:inf.HiddenFromAddressListsEnabled
            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            write-host "Hidden from Address Book AD   : " $Global:u.msExchHideFromAddressLists.value -Foregroundcolor Green
            $LineToWrite = "Hidden from Address Book AD    :" + "`t" + $Global:u.msExchHideFromAddressLists.value
            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            write-host "O365 Retention Comment        : " $Global:inf.RetentionComment -ForegroundColor Green
            $LineToWrite = "O365 Retention Comment         :" + "`t" + $Global:inf.RetentionComment
            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            write-host "AD Retention Comment          : " $ADCmt -ForegroundColor Green
            $LineToWrite = "AD Retention Comment           :" + "`t" + $ADCmt
            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            write-host "AD Object Protected           : " $UsrDetails.ProtectedFromAccidentalDeletion -ForegroundColor Green
            $LineToWrite = "AD Object Protected            :" + "`t" + $UsrDetails.ProtectedFromAccidentalDeletion
            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
            $Global:MbxAccess = Get-MailboxPermission $ENo
            foreach ($Global:MbxAccess in $Global:MbxAccess)
            {
                if ($Global:MbxAccess.User -like "*@global*")
                {
                    write-host "Accounts Granted Access       : " $Global:MbxAccess.User -ForegroundColor Yellow
                    $LineToWrite = "Accounts Granted Access       : " + "`t" + $Global:MbxAccess.User
                    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                }
            }
            Out-File -filepath $ReportFile -append -noClobber -inputObject "`n______________________________________________________________"
            $Cnt++
            }
    }
    else
    {
        write-host""
        write-host "Mailbox does not exist for Emp#" $ENo -ForegroundColor Red
    }

}

write-host "`Number of mailboxes in this Preservation Activity: " $Cnt
$LineToWrite = "Number of mailboxes in this Preservation Activity:" + "`t" + $Cnt
Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite