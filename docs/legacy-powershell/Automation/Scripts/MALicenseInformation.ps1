<#
#
#  Called by:  MAActivitiesMenu.ps1
#
#>

write-host "Executing Report for M&A Licensing....." -ForegroundColor Yellow

$MAUsers = Import-Csv e:\Automation\MAActivities\Input\MAActivities.csv

$text = "UserPrincipalName,IsLicensed,DisplayName,SkuPartNumber,EmployeeType,Location,Organization3,LastLogonTime,MailboxEnabled,MailboxCreationDate,ForwardingAddress"
Out-File -FilePath c:\temp\MALicenseInformation.csv -InputObject $text

Foreach ($MAUsers in $MAUsers)
{
    $Usr = $MAUsers.UPN + "@global.ul.com"
    $UsrExists = [bool](get-MsolUser -UserPrincipalName $Usr -ErrorAction SilentlyContinue)
    If ($usrExists -eq "True")
    {
        $Coluser = Get-MsolUser -UserPrincipalName $Usr
        Foreach ($objLic in $ColUser.Licenses)
	    {
            
	        $usrmbx = get-mailbox $ColUser.UserPrincipalName
            $stsmbx = get-mailboxstatistics $ColUser.UserPrincipalName
		    $text = "{0},{1},""{2}"",{3},{4},""{5}"",""{6}"",{7},{8},{9},{10}" -f $ColUser.UserPrincipalName, $ColUser.IsLicensed, $ColUser.DisplayName, $objLic.AccountSku.SkuPartNumber, $usrmbx.CustomAttribute1, $usrmbx.CustomAttribute3, $usrmbx.CustomAttribute4, $stsmbx.LastLogonTime, $usrmbx.CustomAttribute15, $usrmbx.WhenMailboxCreated, $usrmbx.ForwardingSMTPAddress
	
		    Out-File -FilePath c:\temp\MALicenseInformation.csv -InputObject $text -Append
        }
    }
    else
    {
        Write-Host "User " $MAUsers.UPN " does not exist"
        $text = "{0},{1},{2}" -f $Usr, "False", "User Does Not Exist"

        Out-File -FilePath c:\temp\MALicenseInformation.csv -InputObject $text -Append

    }
}
Write-Host "Report file generated at c:\temp\MALicenseInformation.csv.  If you would like to maintain the input CSV file please rename it manually." -ForegroundColor Red
