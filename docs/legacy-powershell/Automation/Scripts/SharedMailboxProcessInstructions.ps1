#### Process Instructions Script 
$DispInstr = "N"

write-host "Do you want to display the instructions for this process (Y/N)? " -ForegroundColor Red -NoNewline
$DispInstr = Read-Host
If ($DispInstr -eq "Y")
{
    switch ($MbxChg)
    {
        1
            {
                Invoke-Item '.\Instructions\O365 Creating New Shared Mailbox Instructions.docx'
                write-host "When input files are complete hit return to execute the script....." -ForegroundColor Red -NoNewline
                $cont = read-host
            }
        2
            {
                Invoke-Item '.\Instructions\O365 Shared Mailbox Ownership Change.docx'
            }
        3
            {
               Invoke-Item '.\Instructions\O365 Adding Access to Shared Mailbox Instructions.docx'
            }
    }
}  