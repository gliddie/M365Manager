####
#### Mailbox Information Menu
####

Do {
    write-host ""
    write-host "Mailbox Information Menu" -ForegroundColor Magenta
    write-host
    Write-Host "     Enter ( 1) Check Attibutes for a User"
    write-host "           ( 2) Show Folder Message Counts"
    write-host ""
    write-Host "           ( 0) No Changes and Exit"
    write-Host "     Enter Option Above? " -ForegroundColor Red -NoNewline
    $MbxInfo = Read-Host

    If ($MbxInfo -ne 0)
    {
        write-host ""
        write-host "Enter Employee Number or Email Address " -ForegroundColor Yellow -NoNewline
        $ENo = Read-Host

        if (($ENo -notlike "*@*") -and ($ENo.Length -gt 5))
            {
                write-host "This is not a valid employee number or email address"
            }
            elseif ($ENo.length -eq 5)
            {
               $ENo = $ENo + "@global.ul.com"
            }
    }
            		
    switch ($MbxInfo)
    {
        1
        {
            Get-Mailbox $ENo |fl DisplayName,Alias,custom*
            pause
        }

        2
        {
            Get-MailboxFolderStatistics $ENo |out-gridview
            pause
        }             
   
    }
}While ($MbxInfo -ne 0)