$EmpNo = "0"
write-host "Enter Employee Number to Report on Individual User or ""0"" Return to Report on All Mailboxes: " -ForegroundColor Yellow -NoNewline
$EmpNo = read-host

If ($EmpNo -eq "0")
{
    write-host "Reporting on All Mailboxes Report can be found in c:\temp\ActiveSyncReport.csv"
    $MBLDevices = Get-MobileDevice
}
else
{
    $MBLDevices = Get-MobileDevice -Mailbox $EmpNo
	Write-Host "UserName                 DeviceModel          DeviceOS          DeviceType"
	Write-Host "--------                 -----------          --------          ----------"
}

$strPath = "C:\temp\ActiveSyncReport.csv"

If ($EmpNo -eq "0")
{
    $text = "User;DeviceModel;DeviceOS;DeviceType"
    out-file -FilePath $strPath -InputObject $text
}

foreach ($MBLDevices in $MBLDevices)
{
    Write-Host $MBLDevices.Identity.Remove($MBLDevices.Identity.IndexOf("\")) `t $MBLDevices.DeviceModel `t $MBLDevices.DeviceOS `t $MBLDevices.DeviceType
    If ($EmpNo -eq "0")
    {
        $text = """{0}"";""{1}"";""{2}"";""{3}""" -f $MBLDevices.Identity.Remove($MBLDevices.Identity.IndexOf("\")),$MBLDevices.DeviceModel,$MBLDevices.DeviceOS,$MBLDevices.DeviceType
        out-file -FilePath $strPath -InputObject $text -Append
    }
}