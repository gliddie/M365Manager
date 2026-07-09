#### Process Instructions Script 
$DispInstr = "N"

write-host "Do you want to display the instructions for this process (Y/N)? " -ForegroundColor Red -NoNewline
$DispInstr = Read-Host
If ($DispInstr -eq "Y")
{
    switch ($MbxInfo)
    {
        4
            {
                Invoke-Item '.\Instructions\O365 Creating New UM Shared Mailbox Instructions.docx'
                write-host "When input files are complete hit return to execute the script....." -ForegroundColor Red -NoNewline
                $cont = read-host
            }
        5
            {

                Invoke-Item '.\Instructions\O365 Creating New UM Shared Mailbox Instructions.docx'
                write-host "When input files are complete hit return to execute the script....." -ForegroundColor Red -NoNewline
                $cont = read-host
            }
    }
}  