#				$loopWaitForMBX = [bool](get-mailbox $96151@global.ul.com -ErrorAction SilentlyContinue)
#				while($loopWaitForMbx -ne $True)
#				{		
#					sleep 10
#					"Waiting for Mailbox to be created for "
#					$MbxExists = [bool](get-mailbox $96151@global.ul.com -ErrorAction SilentlyContinue)
#				
#					If ($MbxExists = $True)
#					{
#						write-host "configuring forwarding Address"
#						set-mailbox $User.UPN -ForwardingSMTPAddress $User.LegacyForwarding
#						$loopWaitForMbx = $false
#					}
#				}
$USERUPN = "96151@GLOBAL.UL.COM"	
		Do
		{
			"Waiting for Mailbox to be created for $UserUPN"
			sleep 10
		} while (([bool](get-mailbox 96151@global.ul.com -ErrorAction SilentlyContinue)) -ne $True)
		
		write-host "configuring forwarding Address"
#		set-mailbox $User.UPN -ForwardingSMTPAddress $User.LegacyForwarding