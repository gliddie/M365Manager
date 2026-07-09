$mbxstats = get-mailboxstatistics 96151@global.ul.com
$less180 = (get-date).AddDays(-180)

write-host "Last Logoff Time" $mbxstats.LastLogoffTime
write-host "180 days ago" $less180

if ($mbxstats.LastLogoffTime -gt $less180)
{
    write-host "Last Logoff occurred within the last 180 days" $mbxstats.LastLogofftime,$less180
}
else
{
    write-host "Last Logoff has been more than 180 days ago" $mbxstats.LastLogofftime,$less180
}
