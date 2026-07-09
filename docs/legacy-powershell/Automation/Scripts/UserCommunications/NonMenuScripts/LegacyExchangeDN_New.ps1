$mbx = get-mailbox -ResultSize Unlimited
Foreach ($mbx in $mbx)
{
    $mbx.LegacyExchangeDN.length
    ($mbx.LegacyExchangeDN.trimend(" ")).length
    write-host ""
    If ($mbx.LegacyExchangeDN.length -ne ($mbx.LegacyExchangeDN.trim(" ")).length)
    {
        write-host "This mailbox has a trailing space" ($mbx.LegacyExchangeDN.trimend(" "))
        $NewAddr = "X500:" + $mbx.LegacyExchangeDN.trimend(" ")
        set-mailbox $mbx.Alias -EmailAddresses @{add=$NewAddr}
        pause
    }
}