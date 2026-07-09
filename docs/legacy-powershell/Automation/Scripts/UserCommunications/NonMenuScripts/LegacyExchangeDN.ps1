$det = get-mailbox "FRK PQA SN-007 IDS 290"
write-host "'"$det.LegacyExchangeDN"'"
$det.LegacyExchangeDN.length
($det.LegacyExchangeDN.TrimEnd(" ")).Length