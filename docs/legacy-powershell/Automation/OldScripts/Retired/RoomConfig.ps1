##########################################################
#   Configure Room
##########################################################

$text = get-content e:\automation\scripts\RoomResourcePolicy.htm

If ($Global:chkStdDele.Checked -eq "Checked")
{
    If ($Global:chkNewGRRoom.Checked -eq "Checked")
    {
        Set-Mailbox $Global:txtMbxAddr.Text -CustomAttribute15 "ConfigConferenceRoom PS Date: $Date PS Time: $Time"
    }
    else
    {
        Set-Mailbox $Global:txtMbxAddr.Text -CustomAttribute15 "ConfigResourceEquipment PS Date: $Date PS Time: $Time"
    }
}
else
{
    If ($Global:chkNewGRRoom.Checked -eq "Checked")
    {
        Set-Mailbox $Global:txtMbxAddr.Text -CustomAttribute15 "ConfigRestrictedConferenceRoom PS Date: $Date PS Time: $Time"
    }
    else
    {
       Set-Mailbox $Global:txtMbxAddr.Text -CustomAttribute15 "ConfigResourceEquipment PS Date: $Date PS Time: $Time"
    }
}

If ($Global:chkStdDele.Checked -eq "Checked")
{
    Set-CalendarProcessing $Global:txtMbxAddr.Text -AllRequestOutOfPolicy $True -ResourceDelegates $Global:txtDeleName.Text -ProcessExternalMeetingMessages $false -BookingWindowInDays 90 -AutomateProcessing AutoAccept -AllowConflicts $false -MaximumConflictInstances 3 -ConflictPercentageAllowed 20 -AddAdditionalResponse $true -AdditionalResponse $text -DeleteComments $false -DeleteSubject $False
    $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to Delegates to: " + $Global:txtDeleName.Text
    WriteReportEvent
    $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to AutoAccept"
    WriteReportEvent        
    $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to AllRequestOutOfPolicy to True"
    WriteReportEvent
    $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to Standard Booking Policy: 90 Days, 1440 Minutes"
    WriteReportEvent
    $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to AutoResponse Text, Do Not Delete Commnets and Do Not Delete Subject"
    WriteReportEvent
}
else
{
    # If the restricted users is not check set users to the delegates group
    If ($Global:chkResUsr.Checked -eq "Checked")
    {
        Set-CalendarProcessing $LeftName -AllRequestOutOfPolicy $false -AllBookInPolicy $false -RequestOutOfPolicy $Global:txtResUsr.Text -RequestInPolicy $Global:txtResUsr.Text -AllRequestInPolicy $false -BookInPolicy $Global:txtResUsr.Text -ResourceDelegates $Global:txtDeleName.Text -ProcessExternalMeetingMessages $false -BookingWindowInDays 90 -AutomateProcessing AutoAccept -AllowConflicts $false -MaximumDurationInMinutes 1440 -MaximumConflictInstances 3 -ConflictPercentageAllowed 20 -AddAdditionalResponse $true -AdditionalResponse $text -DeleteComments $false -DeleteSubject $False
        If ($Global:chkNewGRRoom.Checked -eq "Checked")
        {
            $LineToWrite = "CFG" + "`t" + "     Setting Room Delegates to: " + $Global:txtDeleName.Text
            WriteReportEvent
            $LineToWrite = "CFG" + "`t" + "     Setting Room Users to: " + $Global:txtResUsr.Text
        }
        else
        {
            $LineToWrite = "CFG" + "`t" + "     Setting Resource Delegates to: " + $Global:txtDeleName.Text
            WriteReportEvent
            $LineToWrite = "CFG" + "`t" + "     Setting Resource Users to: " + $Global:txtResUsr.Text
        }
        WriteReportEvent
        $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to AutoAccept"
        WriteReportEvent        
        $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to AllRequestOutOfPolicy to False"
        WriteReportEvent
        $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to AllBookInPolicy and AllRequestInPolicy to False"
        WriteReportEvent
        $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to Standard Booking Policy: 90 Days, 1440 Minutes"
        WriteReportEvent
        $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to AutoResponse Text, Do Not Delete Commnets and Do Not Delete Subject"
        WriteReportEvent
    }
    else
    {	
  		Set-CalendarProcessing $LeftName -AllRequestOutOfPolicy $false -AllBookInPolicy $false -RequestOutOfPolicy $Global:txtDeleMgrs.Text -RequestInPolicy $Global:txtDeleMgrs.Text -AllRequestInPolicy $false -BookInPolicy $Global:txtDeleMgrs.Text -ResourceDelegates $Global:txtDeleMgrs.Text -ProcessExternalMeetingMessages $false -BookingWindowInDays 90 -AutomateProcessing AutoAccept -AllowConflicts $false -MaximumDurationInMinutes 1440 -MaximumConflictInstances 3 -ConflictPercentageAllowed 20 -AddAdditionalResponse $true -AdditionalResponse $text -DeleteComments $false -DeleteSubject $False
        $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to Room Delegates to: " + $Global:txtDeleName.Text
        WriteReportEvent
        $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to Room Users to: " + $Global:txtDeleName.Text
        WriteReportEvent
        $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to AutoAccept"
        WriteReportEvent        
        $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to AllRequestOutOfPolicy to False"
        WriteReportEvent
        $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to AllBookInPolicy and AllRequestInPolicy to False"
        WriteReportEvent
        $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to Standard Booking Policy: 90 Days, 1440 Minutes"
        WriteReportEvent
        $LineToWrite = "CFG" + "`t" + "     Setting CalendarProcessing to AutoResponse Text, Do Not Delete Commnets and Do Not Delete Subject"
        WriteReportEvent
    }
}

write-host "Granting Global RRS Admins Full Access to room"
Add-MailboxPermission $Global:txtMbxAddr.Text -User "DBS.CRP.RRS.Admins" -AccessRights FullAccess
$LineToWrite = "CFG" + "`t" + "     Granting O365 Admins FullAccess to Calendar"
WriteReportEvent
write-host "Granting Room Delegates Full Access to room"
Add-MailboxPermission $Global:txtMbxAddr.Text -User $Global:txtDeleName.Text -AccessRights FullAccess
$LineToWrite = "CFG" + "`t" + "     Granting Room Delegates FullAccess to Calendar"
WriteReportEvent
write-host "Granting Regional RRS Admins Full Access to room"
Add-MailboxPermission $Global:txtMbxAddr.Text -User $Global:txtRegMgr.Text -AccessRights FullAccess
$LineToWrite = "CFG" + "`t" + "     Granting Regional RRS Admins FullAccess to Calendar"
WriteReportEvent
write-host "Granting All Staff Read Access to room"
Set-MailboxFolderPermission ($Global:txtMbxAddr.Text + $Folder) -User Default -AccessRights Reviewer
$LineToWrite = "CFG" + "`t" + "     Granting All Staff Read Access to Calendar"
WriteReportEvent

If ($Global:txtTimeZone.Text -notlike "")
{
    Set-MailboxCalendarConfiguration $Global:txtDispName.Text -WorkingHoursTimeZone $Global:txtTimeZone.Text
    Set-MailboxRegionalConfiguration $Global:txtDispName.Text -TimeZone $Global:txtTimeZone.Text
    $LineToWrite = "CFG" + "`t" + "     Setting MailboxCalendarConfiguation and MailboxRegionalConfiguration TimeZone to: " + $Global:txtTimeZone.Text
    WriteReportEvent
}

Add-DistributionGroupMember $Global:txtInpRoomGrp.Text -Member $Global:txtDispName.Text
write-host "Adding " $Global:txtDispName.Text "to " $Global:txtInpRoomGrp.Text -ForegroundColor Yellow
$LineToWrite = "CFG" + "`t" + "     Adding Room to RoomList: " + $Global:txtInpRoomGrp.Text
WriteReportEvent

#   Use this code if working hours are added to the process.  This would allow booking only from 8-5 (or what ever hours are identified anythng outisde of this would require delegate approval)
#	If ($Room.WorkHours -eq '$true')
#	{
#	    Set-CalendarProcessing $Room.Mail -ScheduleOnlyDuringWorkHours $true
#		Set-MailboxCalendarConfiguration $Room.Mail -WorkDays $Room.WorkDays -WorkingHoursStartTime $Room.StartTime -WorkingHoursEndTime $Room.EndTime
#       Write-Host "Left name is = " $Leftname
#		Write-Host "Start time   = " $Room.StartTime
#		Write-Host "End time     = " $Room.EndTime
#   }