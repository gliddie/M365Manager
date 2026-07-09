<#
#
#  Called by:  RoomResourceAdminMenu.ps1
#
#
#  Changed the location of where the report details are written to.
#>

$InputMBX = Read-Host "Name of Room or Resource"
$Ticket = Read-Host "Enter Ticket Number" 

$OutFileName = $InputMBX
$OutFileName = $OutFileName.Replace("-","")
$OutFileName = $OutFileName.Replace("(","")
$OutFileName = $OutFileName.Replace(")","")
$OutFileName = $OutFileName.Replace("\","")
$OutFileName = $OutFileName.Replace("/","")
$OutFileName = $OutFileName.Replace(" ","")

$Date = get-date -Format "yyyy-MMdd"
$Year = (get-date).ToString("yyyy")
$Path = "e:\Automation\RemoveRoomResource\Report\" + $Year
If (Test-Path $path) {} else {New-Item -Path $Path -ItemType Directory}

$OutFileName = $Path + "\Remove-" + $OutFileName + "-Date" + $Date + ".log"

$Exists = [bool]($mbx = get-mailbox $InputMBX -ErrorAction SilentlyContinue)

If ($Exists -eq $True)
{
    Write-Output "Room or Resource Information" >> $OutFileName
    Write-Output "Room/Resource Name:  "  $InputMBX >> $OutFileName
    Write-Output "Ticket Number: " $Ticket >> $OutFileName
    write-Output "Run by: "  whoami >> $OutFileName

    get-Mailbox $InputMBX > $OutFileName

    Write-Output "Room or Resource Detailed Information" >> $OutFileName

    get-Mailbox $InputMBX |fl >> $OutFileName

    Write-Output "Mailbox Message Statistics - Number of Messages in Mailbox" >> $OutFileName

    Get-MailboxStatistics $InputMBX |ft >> $OutFileName

    Write-Output "Mailbox Message Folder Statistics - Number of Messages in Each Folder" >> $OutFileName

    Get-MailboxFolderStatistics $InputMBX |ft Name,ItemsInFolder >> $OutFileName

    Write-Output "Mailbox Permissions" >> $OutFileName

    Get-MailboxPermission $InputMBX |ft User,AccessRights >> $OutFileName

    Write-Output "Room or Resource Calendar Notification Information" >> $OutFileName

#    get-CalendarNotification $InputMBX >> $OutFileName
    Get-EventsFromEmailConfiguration $mbx.PrimarySMTPAddress

    Write-Output "Room or Resource Calendar Processing Information" >> $OutFileName

    get-CalendarProcessing $InputMBX |fl >> $OutFileName

    Write-Output "Room or Resource Calendar Configuration Information" >> $OutFileName
    write-host "Ignore Warning about deprecated command the new command does not provide the calendar configuration"

    get-MailboxCalendarConfiguration $InputMBX |fl >> $OutFileName

    Write-Output "Room or Resource Regional Configuration Information" >> $OutFileName

    get-MailboxRegionalConfiguration $InputMBX |fl >> $OutFileName

    remove-Mailbox $InputMBX -confirm:$False

    #stop-transcript
    write-host "Report file written to: " $OutFileName
}
else
{
    write-host "Mailbox not found" -ForegroundColor Red
}