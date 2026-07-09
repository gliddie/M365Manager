#### Process Instructions Script for Group Changes
$DispInstr = "N"

write-host "Do you want to display the instructions for this process (Y/N)? " -ForegroundColor Red -NoNewline

$DispInstr = Read-Host
If ($DispInstr -eq "Y")
{
    switch ($GrpChg)
    {
        1
            {
                Invoke-Item '.\Instructions\O365 Creating New Distribution List Instructions.docx'
                write-host ""
                write-host "When input file is complete hit return to execute the script....." -ForegroundColor Red -NoNewline
                $cont = read-host
            }
        2
            {
                write-host ""
                write-host "     1.  You will be prompted to provide the existing distribution list name" -ForegroundColor Cyan
                write-host "     2.  Based on the name provided the current list owner will be displayed" -ForegroundColor Cyan
                write-host "     3.  If the current ManagedBy field is blank then the owner has left UL and the request can be completed" -ForegroundColor Cyan
                write-host "     4.  If the current Managedby field is not blank the request must come from one of the current owners or approval from a " -ForegroundColor Cyan
                write-host "           current owner must be provided or obtained before proceeding" -ForegroundColor Cyan
                write-host "     5.  If the requestor did not attach approval from the Enterprise Messaging Services shared mailbox send an email to the " -ForegroundColor Cyan
                write-host "           current owners using the ""Distribution List Rename Request Approval Needed"" form" -ForegroundColor Cyan
                write-host "     6.  If the appropriate approval is provided you will be prompted for the new list name and the Service Desk Task number" -ForegroundColor Cyan
                write-host "     7.  The New Name, Alias and Internet Address will be displayed" -ForegroundColor Cyan
                write-host "     8.  If all this information looks correct enter Y to complete the rename" -ForegroundColor Cyan
                write-host ""
             }
        3
            {
                Invoke-Item '.\Instructions\O365 Group Ownership Change Instructions.docx'
                write-host ""
            }
        4
            {
                #Delete Distribution or Security Group
            }
		10
			{
			    Invoke-Item '.\Instructions\O365 Creating New Dynamic Distribution List Instructions.docx'
                write-host ""
                		
			}
		11
			{
				write-host ""
				write-host "     You will be requested to proivde the Location form a DST.All group or the LastName FirstName for a DST.SUP group" -ForegroundColor Cyan
				write-host ""
			}
    
    }
}  