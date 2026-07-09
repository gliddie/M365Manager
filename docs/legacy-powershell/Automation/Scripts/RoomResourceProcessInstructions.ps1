#### Process Instructions Script 
$DispInstr = "N"

write-host "Do you want to display the instructions for this process (Y/N)? " -ForegroundColor Red -NoNewline
$DispInstr = Read-Host
If ($DispInstr -eq "Y")
{
    switch ($ResourceAct)
    {
        1
            {
                #Create New Rooms
                Write-Host "Instructions for Creating New Conference Rooms" -ForegroundColor Magenta
                Write-Host ""
                Write-Host " 1.  To create new Conference Rooms there are multiple input files required they are:"
                write-host "     a.  e:\automation\NewConferenceRoom\Input"
                write-host "     b.  e:\automation\NewUSG\RoomAdmin\Input"
                write-host "     c.  e:\automation\RoomList\Input"
                write-host "     d.  e:\automation\ConfigConferenceroom\Input"
                write-host " 2.  The NewConferenceRoom and ConfigConferenceRoom scripts are the same when editing these files use Excel"
                write-host "     rather than Notepad"
                write-host " 3.  When saving the NewConferenceRoom and ConfigConferenceRoom files perform a ""Save-As"" select ""CSV Comma"
                write-host "     delimited (*.csv)"""
                write-host "     in the name name enter ""Input-ConfigConferenceRoom.csv"""
                write-host " 4.  The first line in each input file is the variable names, keep this line and delete all others"
                write-host ""
                write-host "Naming Conventions - always look in the address book for how the name is constructed for a given site" -ForegroundColor Yellow
                write-host " 5.  The naming convention for Conference Rooms is:"
                write-host "     a.  XXX Name of Conference Room where XXX = the 3 letter site code"
                write-host "     b.  Typically the name is contains ""Conference Room"" , ""Meeting Room"" or ""Room"""
                write-host " 6.  The internet address for a Conference Room is XXX.Location.NameofMailbox.Capacity@ul.com"
                write-host "     example:  NBK.Flr1.ConferenceRoomA.10@ul.com"
                write-host " 7.  The Room Admin (i.e. delegate group) is named:  MBX.XXX.RRS.OutOfPolicy.DE where XXX = 3 letter site code"
                write-host " 8.  The Room List is named XXX Conference Rooms where XXX = 3 letter site code"
                write-host ""
                write-host "Input File Variables " -ForegroundColor Yellow      
                write-host " 9.  Input file variables are:"
                write-host "     a.  NewConferenceRoom and ConfigConferenceRoom Input file:"
                write-host "             i.  Name - Name of the Conference Room (ex: XXX Conference Room Name)"
                write-host "            ii.  Location - Building Name and or Floor"
                write-host "           iii.  Capacity - This is the number of people that the room will accommodate"
                write-host "            iv.  Mail - This is the internet address of the room"
                write-host "             v.  ProcessExternalMeetingMessages - This is always FALSE"
                write-host "            vi.  AutomateProcessing - This is always AutomateProcessing"
                write-host "           vii.  BookingWindowInDays - This is always 90 - this is the number of days into the future a room"
                write-host "                 can be requested without requiring delegate approval"
                write-host "          viii.  BookInPolicy - This is always TRUE"
                write-host "            ix.  DGOwner - This is always MBX.RRS.Owner"
                write-host "             x.  BookOutPolicy - This is the name of the delegates group so MBX.XXX.RRS.OutOfPolicy.DE where"
                write-host "                 XXX = 3 letter site code"
                write-host "            xi.  RRSRegionalOwner - This is the group of IT staff for the region group name DBS.XX.RRS.Admins"
                write-host "                 where XX = 2 letter Region identifier"
                write-host "           xii.  Type - This is always ""Room"""
                write-host "          xiii.  RoomListName - This is the group of all rooms for a site the name is XXX Conference Rooms"
                write-host "                 where XXX = 3 letter site code"
                write-host "           xiv.  WorkHours - This is used if they want to define hours that the room can be booked so only"
                write-host "                 during working hours"
                write-host "            xv.  WorkDays - Days that appear in the calendar as work days in Microsoft Office Outlook Web App"
                write-host "                 forexample WEEKDAYS"
                write-host "           xvi.  TimeZone - specifies the time zone the Nem of the TimeZone can be found in here:"
                write-host "                 https://msdn.microsoft.com/en-us/library/ms912391(v=winembedded.11).aspx"
                write-host "          xvii.  StartTime - Time that the calendar work day starts most often this is left blank if a time"
                write-host "                 is provided this is entered as 8:00 AM use 08:00:00"
                write-host "         xviii.  EndTime - Time that the calendar work day starts most often this is left blank if a time"
                write-host "                 is provided this is entered as 5:00 PM use 17:00:00"
                write-host "           xix.  SDTicket - This is the ticket number for the request"
                write-host "     b.  NewRoomAdmin - creates the OutOfPolicy Delegates group this is only needed if this is a new site"
                write-host "             i.  Name - his is Room Delegate group so MBX.XXX.RRS.OutOfPolicy.DE where XXX = 3 letter site code"
                write-host "            ii.  Mail - this is the internet address of the delegate group MBX.XXX.RRS.OutOfPolicy.DE@ul.com"
                write-host "           iii.  ManagedBy - this is always ""MBX.RRS.Owner"""
                write-host "            iv.  Members - this is the email address or emp#@global.ul.com of the individuals who are listed to"
                write-host "                 be the Room Delegate(s)"
                write-host "     c.  NewRoomList - creates the XXX Conference Rooms group and is only needed if this is a new site"
                write-host "             i.  Name - this is the name of the group that will list conference rooms for the site the naming"
                write-host "                 standard is XXX Conference Rooms"
                write-host "            ii.  ManageBy - this is Room Delegate group so MBX.XXX.RRS.OutOfPolicy.DE where XXX = 3 letter site"
                write-host "                 code"
                write-host "     d.  ConfConferenceRoom is identical to the NewConferenceRoom Input file:"
                write-host "10.  You can create condference rooms using one input file each new line will be for conference room"
                write-host "11.  Save the input files along the way but make sure to perform a final Save before proceeding with executing"
                write-host "     the script"
                write-host "12.  Send a ""New Conference Room"" from the O365 Admin Templates - message is sent to the requestor and room"
                write-host "     delegate(s)"
                write-host "13.  Enter ""Conference Room created attached email sent with details."" in the work history, attach a copy of"
                write-host "     the email sent and close the ticket"
                write-host ""
                write-host "When input files are complete hit return to execute the script....." -ForegroundColor Red -NoNewline
                $cont = read-host
            }
 
    }
}