#
#  Scripts to immediately remove access for Terminated Staff
#
#  Cancel all future meetings organized by the Terminated Staff's mailbox
#
    $EmpID = Read-Host "Enter the Employee Number of the Terminated Individual"
    $EmpID = $EmpID + "@global.ul.com"
    $Usr = Get-MsolUser -UserPrincipalName $EmpID
    Revoke-AzureADUserAllRefreshToken -ObjectID $Usr.ObjectID

#    Remove-CalendarEvents -Identity $EmpID -CancelOrganizedMeetings
#
#  Disable All Mailbox Protocols for terminated user
#
    Set-Mailbox $EmpNo -AccountDisabled:$True
    Set-CASMailbox -Identity $EmpNo -OwaEnabled $false -ActiveSyncEnabled $false -EwsAllowMacOutlook $false
  