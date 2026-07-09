	$mbx = get-mailbox -ResultSize Unlimited | where-object {$_.RecipientTypeDetails -eq "SharedMailbox"}
	
	foreach ($mbx in $mbx)
	{
		write-host "Shared Mailbox: " $Mbx.DisplayName
		set-mailbox $Mbx.Alias -MessageCopyForSendOnBehalfEnabled $true -MessageCopyForSentAsEnabled $true
	}