##########################################################
#   Create New Conference Room or Equipment/Resource
##########################################################

If ($Global:chkNewGRRoom.Checked -eq "Checked")
{
    If ($Global:txtBldgName.Text -ne "")
    {
        If ($Global:txtFlrNo.Text -eq "0")
        {
            $Location = "Building " + $Global:txtBldgName.Text + ", Ground Floor"
        }
        else
        {
            If ($Global:txtFlrNo.Text.Length -ne 0)
            {
                $Location = "Building " + $Global:txtBldgName.Text + ", Floor " + $Global:txtFlrNo.Text
            }
        }
    }
    else
    {
        If ($Global:txtFlrNo.Text -eq "0")
        {
            $Location = "Ground Floor"
        }
        else
        {
            If ($Global:txtFlrNo.Text.Length -ne 0)
            {
                $Location = "Floor " + $Global:txtFlrNo.Text
            }
        }
    }
    New-Mailbox -Name $Global:txtDispName.Text -Room -DisplayName $Global:txtDispName.Text -Alias $LeftName -PrimarySmtpAddress $Global:txtMbxAddr.Text -ResourceCapacity $Global:txtRoomCap.Text -Confirm:$False
}
else
{
    New-Mailbox -Name $Global:txtDispName.Text -Equipment -DisplayName $Global:txtDispName.Text -Alias $LeftName -PrimarySmtpAddress $Global:txtMbxAddr.Text -Confirm:$False
}

write-host "Pausing script to allow the mailbox creation to complete in O365 environment" -ForegroundColor Yellow
Do
{
    $Exists = [bool](Get-Mailbox $Global:txtDispName.Text -ErrorAction Silentlycontinue)
} while ($Exists -ne $True)

Set-Mailbox $Global:txtDispName.Text -IssueWarningQuota 0.5GB -ProhibitSendQuota 0.75GB -ProhibitSendReceiveQuota 1.0GB
   
$LineToWrite = "NEW" + "`t" + "     Task Number: " + $Global:txtInpTaskNo.Text
WriteReportEvent
If ($Global:chkNewGRRoom.Checked -eq "Checked")
{
    $LineToWrite = "NEW" + "`t" + "     New Room Name: " + $Global:txtDispName.Text
}
else
{
    $LineToWrite = "NEW" + "`t" + " New Resource Name: " + $Global:txtDispName.Text
}
WriteReportEvent
$LineToWrite = "NEW" + "`t" + "     Email Address: " + $Global:txtMbxAddr.Text
WriteReportEvent
If ($Global:chkNewGRRoom.Checked -eq "Checked")
{
    $LineToWrite = "NEW" + "`t" + "     Room Group Name: " + $Global:txtInpRoomGrp.Text
    WriteReportEvent
    $LineToWrite = "NEW" + "`t" + "     Room Building Location: " + $Global:txtBldgName.Text
    WriteReportEvent
    $LineToWrite = "NEW" + "`t" + "     Room Floor Number: " + $Global:txtFlrNo.Text
    WriteReportEvent
    $LineToWrite = "NEW" + "`t" + "     Room Capacity: " + $Global:txtRoomCap.Text
    WriteReportEvent
}

If ($Global:chkStdDele.Checked -eq "Checked")
{
    If ($Global:chkNewGRRoom.Checked -eq "Checked")
    {
	    Set-Mailbox $LeftName -CustomAttribute15 "NewConferenceRoom PS Date: $Date PS Time: $Time"
    }
    else
    {
        Set-Mailbox $LeftName -CustomAttribute15 "NewResourceEquipment PS Date: $Date PS Time: $Time"
    }
    $LineToWrite = "NEW" + "`t" + "     Room Delegates Group: " + $Global:txtDeleName.Text
    WriteReportEvent
}
else
{
    If ($Global:chkNewGRRoom.Checked -eq "Checked")
    {
    	Set-Mailbox $LeftName -CustomAttribute15 "NewRestrictedConferenceRoom PS Date: $Date PS Time: $Time"
    }
    else
    {
        Set-Mailbox $LeftName -CustomAttribute15 "NewRestrictedEquipmentResource PS Date: $Date PS Time: $Time"
    }
    $LineToWrite = "NEW" + "`t" + "     Restricted Delegates Group: " + $Global:txtResUsr.Text
    WriteReportEvent
    $LineToWrite = "NEW" + "`t" + "     Restricted Users: " + $Global:txtResUsrInfo.Text
    WriteReportEvent
}

$LineToWrite = "NEW" + "`t" + "     Room TimeZone: " + $Global:txtTimeZone.Text
WriteReportEvent
$LineToWrite = "NEW" + "`t" + "     Room Regional Managers: " + $Global:txtRegMgr.Text
WriteReportEvenT