<#
#
#  Called by:  RoomResourceAdminMenu.ps1
#
#>

$InputMBX = Read-Host "Name of Room or Resource"

$OutFileName = $InputMBX
$OutFileName = $OutFileName.Replace("-","")
$OutFileName = $OutFileName.Replace("(","")
$OutFileName = $OutFileName.Replace(")","")
$OutFileName = $OutFileName.Replace("\","")
$OutFileName = $OutFileName.Replace("/","")

$OutFileName = "c:\temp\" + $OutFileName + ".txt"

start-transcript

Write-Output "Room or Resource Information" >> $OutFileName

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

get-CalendarNotification $InputMBX >> $OutFileName

Write-Output "Room or Resource Calendar Processing Information" >> $OutFileName

get-CalendarProcessing $InputMBX |fl >> $OutFileName

Write-Output "Room or Resource Calendar Configuration Information" >> $OutFileName

get-MailboxCalendarConfiguration $InputMBX |fl >> $OutFileName

Write-Output "Room or Resource Regional Configuration Information" >> $OutFileName

get-MailboxRegionalConfiguration $InputMBX |fl >> $OutFileName

remove-Mailbox $InputMBX

Stop-transcript