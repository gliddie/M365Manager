#
#	Enable/Disable Compromised Account
#

Param
(
  [string]$ENo,
  [string]$TASKNo
)

Connect to AD and O365

Write-Host "Disabling AD Account"
Disable-ADAccount -Identity $ENo
$EmpNo = $ENo + "@global.ul.com"

Write-Host "Resetting AD Password to Random Password"
$Pwd = get-random
Set-ADAccountPassword -Identity $ENo -NewPassword (ConvertTo-SecureString -AsPlainText [string]$PWD -Force)

Write-Host "Revoking AzureADTokens"
Revoke-AzureADUserAllRefreshToken -ObjectID (get-MSOLUser -UserPrincipalName $EmpNo).ObjectID

Write-Host "Disabling Mailbox and Mailbox Protocols" -foregroundcolor Red
Set-Mailbox $ENo -AccountDisabled:$True

Write-Host "Disabling mailbox protocols"
Set-CASMailbox -Identity $ENo -OwaEnabled $false -OWAforDevicesEnabled $false -ActiveSyncEnabled $false -EwsAllowMacOutlook $false

Write-Host "Checking to see if the messages are forwarded at the mailbox level.  Unless this user is from an M&A that has not migrated"
Write-Host "there should be no forwarding set for the mailbox"
get-mailbox $ENo |fl *forw*
$mbx = get-mailbox $EmpNo
if (($mbx.ForwardingAddress -ne $null) -or ($mbx.ForwardingSmtpAddress -ne $null))
{
    write-host "Forwarding Exists on Mailbox: "
    write-host "    Forwarding Address: " $mbx.ForwardingAddress
    write-host "    Forwading SMTP Addres: "$mbx.ForwardingSMTPAddress
}
            
get-InboxRule -Mailbox $EmpNo |select Name,Enabled,Priority,RuleIdentity |ft
#    write-host "Enter the Rule Name: " -foregroundcolor Cyan -NoNewline
#    $RuleName = Read-Host
#    get-InboxRule -Mailbox $EmpNo -Identity $RuleName |fl *Description*
#    Disable-InboxRule -Mailbox $EmpNo -Identity $RuleName

#Send email to LST.Global.DigitalSecurity@ul.com with transcript file


x1.	Disable Mailbox (not sure if this is required) Yes, Disable the AD account so the affected user or hacker cannot log in.
x2.	Reset AD Password Yes, Correct.
x3.	Revoke AzureTokens Yes, Correct.
x4.	Disable OWA Access, OWA for Devices, MAC Access (confirm if this is required) Yes, Correct.
5.	Execute command to get rules on mailbox – where should the output of this be written to  so that ServiceNow can determine if a ticket needs to be created for the Enterprise Email Team to investigate further? Don’t know. Let’s ask Bob. – Bob?
6.	Add Comment in AD account (what field did we determine this should be placed in the Description?) Yes, use the description field. Use verbiage “Compromised Account Task  - TASK#####”. How does is the TASK# passed to the script?
7.	What account will this use to run as it will need to be given privileges in our environment to execute these tasks.  Use the same service account as the workstation admin rights. Who has access to this account as I’m not comfortable giving an account O365 Admin rights to an account where the password of the account is known to others and is used to administer workstations.
8.	Do you want an email sent after this executes?  If so who should the email be sent to?  LST.Global.Digital Security
9.	Should we include removing Mailbox delegates? -  I would recommend we report of there are delegates (maybe not remove them)
10.	Disable Skype for Business (per Christian) – if we do this then their extension and everything is removed I’m not sure that this is a good idea.


$MenuOpt = 1

Do
{
    write-host "Connecting to AzureAD...... " -ForegroundColor Cyan
    Connect-AzureAD -Credential $Global:LiveCred | Out-Null

            write-host "`nEnter Employee Number: " -foregroundcolor cyan -NoNewline
	        $ENo = Read-Host
	        $EmpNo = $ENo + "@global.ul.com"
            $mbx =  get-mailbox $EmpNo
            Write-Host "Changing Account Settings for: " -ForegroundColor Green -NoNewline
            write-host $ENo,$Mbx.Name
            write-host "`nIs this the correct account (Y/N)? " -ForegroundColor Red -NoNewline
            $Cont = Read-Host
        }while($Cont -eq "N")
    }

    switch ($MenuOpt)
    {
	    1
	    {
	    #
	    #   Disable Compromised Account
	    #
            Write-Host "Disabling AD Account" -foregroundcolor Red
    		Disable-ADAccount -Identity $ENo
    		Write-Host "Revoking AzureADTokens" -foregroundcolor Red
    		Revoke-AzureADUserAllRefreshToken -ObjectID (get-MSOLUser -UserPrincipalName $EmpNo).ObjectID
		    Write-Host "Disabling Mailbox and Mailbox Protocols" -foregroundcolor Red
    		Set-Mailbox $EmpNo -AccountDisabled:$True
		    Set-CASMailbox -Identity $EmpNo -OwaEnabled $false -OWAforDevicesEnabled $false -ActiveSyncEnabled $false -EwsAllowMacOutlook $false
            Write-Host "Checking to see if the messages are forwarded at the mailbox level.  Unless this user is from an M&A that has not migrated"
            Write-Host "there should be no forwarding set for the mailbox"
            get-mailbox $EmpNo |fl *forw*
            $mbx = get-mailbox $EmpNo
            if (($mbx.ForwardingAddress -ne $null) -or ($mbx.ForwardingSmtpAddress -ne $null))
            {
                write-host "Does Forwarding need to be disabled for this account (Y/N?)" -ForegroundColor Cyan -NoNewline
                $DisForw = Read-Host
                If ($DisForw -eq "Y")
                {
                    set-mailbox $EmpNo -ForwrdingAddress $null -ForwardingSmtpAddress $null -DeliverToMailboxAndForward:$False
                    write-host "Forwarding removed for this account" -ForegroundColor Cyan
                    get-mailbox $EmpNo |fl *forw*
                }
                else
                {
                    write-host "Forwarding Not removed for this account" -ForegroundColor Red
                }
            }
            
            
            get-InboxRule -Mailbox $EmpNo |select Name,Enabled,Priority,RuleIdentity |ft
		    write-host "Do you need to Disable or see more details about a rule (Y/N)? " -foregroundcolor cyan -NoNewline
		    $Rules = read-host
		    Do
		    {
                If ($Rules -eq "Y")
                {
    			    write-host "Enter the Rule Name: " -foregroundcolor Cyan -NoNewline
	    		    $RuleName = Read-Host
		    	    get-InboxRule -Mailbox $EmpNo -Identity $RuleName |fl *Description*
			        write-host "Do you want to disable this rule? " -foregroundcolor cyan -NoNewline
    			    $DisRule = Read-Host
	    		    If ($DisRule -eq "Y")
		    	    {
			    	    Disable-InboxRule -Mailbox $EmpNo -Identity $RuleName
    			    }
                    get-InboxRule -Mailbox $EmpNo |select Name,Enabled,Priority,RuleIdentity |ft
	    		    write-host "Disable or see more details on another rule (Y/N)? " -foregroundcolor cyan -Nonewline
		    	    $Rules = read-host	
                }	
		    }while($Rules -eq "Y")
         }
	
	    2
	    {
	    #
	    #   Enable Compromised Account
	    #
		    Write-Host "Enabling AD Account" -foregroundcolor Red
    		Enable-ADAccount -Identity $ENo
		    Write-Host "Enable Mailbox and Mailbox Protocols"
    		Set-Mailbox $EmpNo -AccountDisabled:$False
    		Set-CASMailbox -Identity $EmpNo -OwaEnabled $true -OWAforDevicesEnabled $true -ActiveSyncEnabled $true -EwsAllowMacOutlook $true
		    get-InboxRule -Mailbox $EmpNo |select Name,Enabled,Priority,RuleIdentity |ft
		    write-host "Do you need to Disable or see more details about a rule (Y/N)? " -foregroundcolor cyan -NoNewline
		    $Rules = read-host
		    do
		    {
			    If ($Rules -ne "N")
                {
                    write-host "Enter the Rule Name: " -foregroundcolor Cyan -NoNewline
    			    $RuleName = Read-Host
	    		    get-InboxRule -Mailbox $EmpNo -Identity $RuleName |fl *Description*
		    	    write-host "Do you want to disable this rule? " -foregroundcolor cyan
			        $DisRule = Read-Host
    			    If ($DisRule -eq "Y")
	    		    {
		    		    Disable-InboxRule -Mailbox $EmpNo -Identity $RuleName
			        }
    	  		    get-InboxRule -Mailbox $EmpNo |select Name,Enabled,Priority,RuleIdentity |ft
                    write-host "Disable or see more details on another rule (Y/N)? " -foregroundcolor cyan
			        $Rules = read-host
                }		
		    }while($Rules -eq "Y")

            write-host "Check in the O365 Exchange Admin Portal to see if outbound email for this user has been blocked.  This is found in the EAC under Protection-->ActionCenter"
        }
    }
}while($MenuOpt -ne "0")