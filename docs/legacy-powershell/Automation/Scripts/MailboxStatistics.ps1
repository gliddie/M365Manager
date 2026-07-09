<#
#
#  Called by:  UserMailboxMenu.ps1
#              SharedMailboxAdminMenu.ps1
#
#>
    write-host ""
    write-host "User Mailbox Statistics Menu" -ForegroundColor Magenta
    write-host ""
    write-host "     Enter ( 1) to Show Mailbox Total Size and User Created Folder Counts"
    WRITE-HOST "           ( 2) to Display Folder Statistics to your display"
    write-host "           ( 3) to Write Folder Statistics to an output file"
    write-host ""
    write-host "           ( 0) to Return to the Previous Menu"
    write-host "     Enter Option? " -ForegroundColor Red -NoNewline
    $DispInf = Read-Host

    if ($DispInf -ne 0)
    {
        write-host""
        write-host "Enter Shared Mailbox Name or Employee Number to obtain Folder Statistics for: " -ForegroundColor Yellow -NoNewline
        $MBXName = Read-Host
    }
    
    switch ($DispInf)
    {
        1
        {
            $FldrCnt = Get-MailboxFolderStatistics $MbxName | Where-Object{$_.FolderType -eq "User Created"}
            $MbxStats = Get-MailboxStatistics $MBXName
            Write-Host ""
            Write-Host $MbxName "has" $FldrCnt.Count "User Created Folders and is" $MbxStats.TotalItemSize -ForegroundColor Cyan
            Pause
        }
        2
        {
            Get-MailboxFolderStatistics $MBXName |Out-GridView
        }
        3
        {
            Get-MailboxFolderStatistics $MBXName |ft FolderPath,FolderSize,FolderType,ItemsInFolderandSubfolders,FolderandSubfolderSize >c:\temp\$MbxName.csv
        }
    }