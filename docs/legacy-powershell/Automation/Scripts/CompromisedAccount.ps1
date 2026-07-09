#
#	Enable/Disable Compromised Account
#

$MenuOpt = 1

Do
{
    write-host "Connecting to AzureAD...... " -ForegroundColor Cyan
    Connect-AzureAD -Credential $Global:LiveCred | Out-Null

    write-host "`n`n`n   Enter (1) Disable Compromised Account"
    write-host "         (2) Enable Compromised Account"
    write-host "`n         (0) to Exit"
    write-host "   Enter Option Above? " -foregroundcolor Red -nonewline
    $MenuOpt = Read-Host

    If ($MenuOpt -ne "0")
    {
    	Do
        {
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