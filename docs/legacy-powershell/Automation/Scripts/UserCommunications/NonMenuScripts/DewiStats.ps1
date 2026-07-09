    $Mbx = import-csv e:\automation\maactivities\input\maactivities.csv
    
    foreach ($mbx in $mbx)
    {
        $FldrCnt = Get-MailboxFolderStatistics $Mbx.UPN  | Where-Object{$_.FolderType -eq "User Created"}
        $AllFldrCnt = Get-MailboxFolderStatistics $Mbx.UPN
         write-host $mbx.upn "has" $FldrCnt.Count "User Created Folders and" $AllFldrCnt.Count "All Folders"

    }