foreach ($i in (get-mailbox -ResultSize Unlimited))
{
    write-host $i.DisplayName
    If ($i.RecipientTypeDetails -eq "UserMailbox")
    {
        write-host ""
        Get-InboxRule -Mailbox $i.DistinguishedName | where {$_.ForwardTo -ne $null} |fl Identity,Name,ForwardTo
   }
}