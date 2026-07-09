####
#### User Mailbox Menu
####
write-host ""
invoke-expression -Command .\ConnectO365.ps1
$UsrMbxAct = "1"

Do
{
    write-host ""
    write-host "User Mailbox Admin Menu" -ForegroundColor Magenta
    write-host
    Write-Host "     Enter ( 1) User Mailbox Informtion"
    write-host "           ( 2) Grant User Full Permissions to Another Users Mailbox"
    write-host "           ( 3) Remove User's Full Permissions to another Users Mailbox"
    write-host "           ( 4) Remove Forwarder from a User Mailbox"
    write-host "           ( 5) View Mailbox Folder Statistics"
	write-host "           ( 6) Find Object Associated with Internet Address"
	write-host "           ( 7) Set-up Gold/Silver Employee Mailbox Access (Under Development)"
    write-host ""
    write-Host "           ( 0) No Changes and Exit"
    write-Host "     Enter Option Above? " -ForegroundColor Red -NoNewline
    $UsrMbxAct = Read-Host
    write-host ""

    #invoke-Expression -Command .\ProcessInstructions.ps1
        		
		switch ($UsrMbxAct)
		{
                1
                {
                    write-host "Enter Employee Number or Email Address " -ForegroundColor Yellow -NoNewline
                    $ENo = Read-Host
                    if (($Eno -notlike "*@*") -and ($ENo.Length -gt 5))
                    {
                        write-host "This is not a valid employee number or email address"
                    }
                    elseif ($ENo.length -eq 5)
                    {
                        $ENo = $ENo + "@global.ul.com"
                    }
                    if ($ENo -like "*@*")
                    {
                        Get-Mailbox $ENo |fl DisplayName,Alias,custom*,*forw*
                        write-host ""
                    }
                    pause
                    
                }
                2
                {
                    write-host "Enter Employee Number of mailbox to grant access to: " -ForegroundColor Yellow -NoNewline
                    $ENo = Read-Host
                    Get-MailboxPermission $ENo | fl *user*
                    Write-Host ""
                    Write-Host "Enter Employee Number of the individual to be granted access: " -ForegroundColor Yellow -NoNewline
                    $ENoDele = Read-Host
                    Write-Host "Granting " $ENoDele " to " $ENo "Mailbox"
                    Add-MailboxPermission $ENo -AccessRights FullAccess -User $ENoDele
                    pause
                    
                }
                3
                {
                    write-host "Enter Employee Number of mailbox to remove access from: " -ForegroundColor Yellow -NoNewline
                    $ENo = Read-Host
                    Get-MailboxPermission $ENo | fl *user*
                    Write-Host ""
                    Write-Host "Enter Employee Number of the individual whose permissions will be revoked: " -ForegroundColor Yellow -NoNewline
                    $ENoDele = Read-Host
                    Write-Host "Revoking " $ENoDele " from " $ENo "Mailbox" -ForegroundColor Red
                    Remove-MailboxPermission $ENo -AccessRights FullAccess -User $ENoDele
                    pause
                }
                4
                {
                    write-host "Enter Employee Number of mailbox to remove forwarder from: " -ForegroundColor Yellow -NoNewline
                    $ENo = Read-Host
                    Get-Mailbox $ENo |fl DisplayName,*forwarding*
                    Write-Host "Removing forwarding address from employee" $ENo "Mailbox" -ForegroundColor Red
                    Set-Mailbox $ENo -ForwardingSmtpAddress $Null -ForwardingAddress $null
                    pause
                }
                5
                {
                    invoke-expression -Command .\MailboxStatistics.ps1
                }
				6
				{
					write-host "This process allows you to find what objects in O365 have a particular internet email address or a portion"
					write-host "of the internet address associated with it.  For example I can search for all objects with ""us.ul.com"" or"
					write-host "all objects with ""Sandi"" in it.  Since this command is running against all O365 objects it will take some"
					write-host "time to display the results.  If you do not wish to proceed hit enter and no command will be executed."
					write-host
					write-host "Enter the email address or portion of address you would like to find that match your query " -ForegroundColor Green -NoNewline
					$FindAddr = Read-Host
					if ($FindAddr -ne "")
					{
						Get-Recipient -ResultSize Unlimited | where {$_.emailaddresses -match $FindAddr} |fl Name,EmailAddresses
					}
				}
        }
}While ($UsrMbxAct -ne 0)	