<####
#### Unified Message Menu
####
#
#  Called by:  O365AdminMenu.ps1
#
#>

Do {
    write-host ""
    write-host "Unified Messaging Menu" -ForegroundColor Magenta
    write-host
    Write-Host "     Enter ( 1) Enable Unified Messaging"
    write-host "           ( 2) Reset EUM PIN"
    write-host "           ( 3) Check Assigned Extension and DialPlan"
	write-host "           ( 4) Enable and License UM Mailbox"
    write-host "           ( 5) Configure UM Mailbox"
    write-host ""
    write-Host "           ( 0) No Changes and Exit"
    write-Host "     Enter Option Above? " -ForegroundColor Red -NoNewline
    $MbxInfo = Read-Host

    If ($MbxInfo -ne 0)
    {
        invoke-Expression -Command .\UMMailboxProcessInstructions.ps1
    }
           		
    switch ($MbxInfo)
    {
        1
        {
            invoke-expression -Command .\EnableUM.ps1
        }

        2
        {
            write-host ""
            write-host "Enter Employee Number or Email Address " -ForegroundColor Yellow -NoNewline
            $ENo = Read-Host

            if (($ENo -notlike "*@*") -and ($ENo.Length -gt 5))
            {
                write-host "This is not a valid employee number or email address"
            }
            elseif ($ENo.length -eq 5)
            {
                $ENo = $ENo + "@global.ul.com"
            }
           
            If (([bool](get-UMMailbox $ENo -ErrorAction SilentlyContinue)) -ne "True")
			{
				write-host "This user does not have Unififed Messaging Enabled"
			}
			else
			{
				Set-UMMailboxPIN -Identity $ENo -SendEmail $true -PinExpired $true
				write-host "Um Mailbox PIN reset email message sent to users mailbox"
			}
			pause
        }
        
        3
        {
            write-host ""
            write-host "Enter Employee Number or Email Address " -ForegroundColor Yellow -NoNewline
            $ENo = Read-Host

            if (($ENo -notlike "*@*") -and ($ENo.Length -gt 5))
            {
                write-host "This is not a valid employee number or email address"
            }
            elseif ($ENo.length -eq 5)
            {
                $ENo = $ENo + "@global.ul.com"
            }
            
            Get-UMMailbox $ENo |fl EmailAddresses
        }
		4		
        {
#           Step 1 - Mail Enable Account and Assign P2 License


            write-host ""
            write-host "To Mail Enalbe Accounts you must have the Exchange Modules on your workstation.  If you"
            write-host "do not have these modules then run this part of the process from USNBKM100P.  Also this"
            write-host "you must be running the Active Directory Powershell Module for this to run correctly."
            write-host ""
            write-host "Enter (Y) to run the EnableMailUserAssignLicense Scripts and (N) to exit this step: " -foregroundcolor cyan -nonewline
            $RunEnable = read-host

            If ($RunEnable -eq "Y")
            {
                invoke-expression -Command .\EnableSvcActAssignLicense.ps1

                write-host ""
                write-host "Updated Notes field on the Telephones tab of the AD Account with Account Ownership details" -ForegroundColor Red
                write-host ""
                write-host "Set AD Account to prohibit logins to the Service Account" -ForegroundColor Red
                write-host ""
                write-host "Moved AD Account to the Enterprise\O365 Licensed\UM Mailboxes container in Active Directory" -ForegroundColor Red
                write-host ""
            }
            else
            {

                write-host "No changes made to account....." -ForegroundColor Red -NoNewLine
                $Cont = read-host
            }
        }
        5
        {

#           Step 1 - Create security Groups to manage access"
            write-host ""
            write-host "Executing the Creation of the New Security Groups hit return when ready" -ForegroundColor Yellow -NoNewline
            $Cont = read-host
            $ExeOK = "Y"
            Do {
                invoke-expression -Command .\NewUSG.ps1
                Write-Host "Do you need to run this script again due to errors (Y/N)?" -ForegroundColor Yellow -NoNewline
                $ExeOK = Read-Host
            } while ($ExeOK -eq "Y")

#           Step 2 - Applying Mailbox Permissions"
            write-host ""
            write-host "Applying Mailbox Permissions to the New UM Mailboxes - make sure the mailbox creation has completed before continuing" -ForegroundColor Yellow -NoNewline
            $Cont = read-host
            $ExeOK = "Y"
            Do {
                invoke-expression -Command .\AddMailboxPermissionSharedMailbox.ps1
                Write-Host "Do you need to run this script again due to errors (Y/N)?" -ForegroundColor Yellow -NoNewline
                $ExeOK = Read-Host
            } while ($ExeOK -eq "Y")

#           Step 3 - Applying Mailbox Folder Permissions"
            write-host ""
            write-host "Applying Folder Permissions to the UM Mailboxes hit return when ready" -ForegroundColor Yellow -NoNewline
            $Cont = read-host
            $ExeOK = "Y"
            Do {
                invoke-expression -Command .\AddFolderPermissions.ps1

                Write-Host "Do you need to run this script again due to errors (Y/N)?" -ForegroundColor Yellow -NoNewline
                $ExeOK = Read-Host
            } while ($ExeOK -eq "Y")

#           Step 4 - Applying Retention Policy"
            write-host ""
            write-host "Applying Retention Policy to the New Mailboxes hit return when ready" -ForegroundColor Yellow -NoNewline
            $Cont = read-host
            $ExeOK = "Y"
            Do {
                invoke-expression -Command .\ApplyRetentionPolicy.ps1
                Write-Host "Do you need to run this script again due to errors (Y/N)?" -ForegroundColor Yellow -NoNewline
                $ExeOK = Read-Host
            } while ($ExeOK -eq "Y")

#           Step 5 - Assign Extension to Mailbox"			
            write-host ""
            write-host "Configure UM Extension to UM Mailboxes hit return when ready" -ForegroundColor Yellow -NoNewline
            $Cont = read-host
            $ExeOK = "Y"
            Do {
                invoke-expression -Command .\EnableUM.ps1
                Write-Host "Do you need to run this script again due to errors (Y/N)?" -ForegroundColor Yellow -NoNewline
                $ExeOK = Read-Host
            } while ($ExeOK -eq "Y")			

            write-host ""
            write-host "UM Mailbox Set-up and Configuration Complete" -ForegroundColor Magenta
            write-host ""
            
            write-host "Hit any key to continue....." -ForegroundColor Cyan -NoNewLine
            $Cont = read-host
        }		
   
    }
}While ($MbxInfo -ne 0)