$Mbx = Import-csv e:\automation\maactivities\maactivities.csv
Foreach ($Mbx in $Mbx)
{
    $FldrCnt = Get-MailboxFolderStatistics $MbxName | Where-Object{$_.FolderType -eq "User Created"}
}