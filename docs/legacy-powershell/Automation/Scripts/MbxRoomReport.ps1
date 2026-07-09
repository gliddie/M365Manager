<#
#
#  Called by: ReportingMenu.ps1
#
#>

$Date = get-date -Format "yyyyMM"

Write-Host "Starting All Room Mailbox Report.  This Will take some time to complete......"
Write-Host ""
write-host "Are you restarting this script due to a previous failure (Y/N)? " -ForegroundColor Yellow -NoNewLine
$Rest = Read-Host

$Cnt = 1

If ($Rest -eq "N")
{
    $Mbx = Get-Recipient -Resultsize Unlimited |where {$_.RecipientTypeDetails -eq "EquipmentMailbox"}
    Write-Host "Total Number of Room Mailboxes: " $Mbx.Count
	out-file -FilePath c:\temp\Input-RoomMbxUPN.csv -InputObject "Alias,PrimarySMTPAddress"
	foreach ($Mbx in $Mbx)
	{
		$Detail = $Mbx.Alias + "," + $Mbx.PrimarySMTPAddress
		out-file -FilePath c:\temp\Input-RoomMbxUPN.csv -InputObject $Detail -Append
	}
	write-host "Do you want to create a report file?  This will take some time to complete. (Y/N)? " -foregroundColor Yellow -Nonewline
    $Rpt = read-host
	
	Write-Host "Creating Report on c:\temp\AllRoomMbxDetails.csv" -ForegroundColor Yellow
	$strAllDetails = "C:\temp\AllRoomMbxDetails.csv"
    $strTooManyFolders = "C:\temp\AllRoomMbxTooManyFldrsDetails.csv"
    $strOver30GB = "C:\temp\AllRoomMbxOvr30GBDetails.csv"

    $headers = "Emp#,PrimarySMTPAddress,DisplayName,MailboxType,WhenCreated,WhenChanged,RetentionPolicy,UserCreatedFolders,TotalSizeInBytes,TotalSize(GB)"
    Out-File -FilePath $strAllDetails -InputObject $headers
	Out-File -FilePath $strTooManyFolders -InputObject $headers
	Out-File -FilePath $strOver30GB -InputObject $headers
}
else
{
	write-host "Enter the last Room Maibox Processed: " -ForegroundColor Yellow -NoNewLine
	$LastProc = Read-Host
}

	$Mbx = Import-csv c:\temp\Input-RoomMbxUPN.csv
	Write-Host "Total Number of Room Mailboxes: " $Mbx.Count
	
	foreach ($Mbx in $Mbx)
	{
		If ($Rest -eq "Y")
		{
			If ($LastProc -ne $Mbx.alias)
			{
				$cnt = $cnt + 1
			}
			else
			{
				$Rest = "N"
				$cnt = $cnt + 1
			}
			write-host $cnt
		}
		else
		{	
			if (([bool](get-mailbox $mbx.alias -ErrorAction SilentlyContinue) -eq "True"))
			{
				write-host $cnt,$mbx.alias
				$MbxDetails = Get-Mailbox $mbx.alias
				$MbxFldr = Get-MailboxFolderStatistics $mbx.Alias
				$MbxStats = Get-MailboxStatistics -Identity $mbx.Alias
				$MbxStats | Add-Member -MemberType ScriptProperty -Name TotalItemSizeInBytes -Value {$this.TotalItemSize -replace "(.*\()|,| [a-z]*\)", ""}
				$SizeGB = $MbxStats | Select-Object @{Name="TotalItemSize(GB)"; Expression={[math]::Round($_.TotalItemSizeInBytes/1GB,2)}}

                $text = """{0}"",{1},""{2}"",{3},{4},{5},{6},{7},{8},{9}" -f $MbxDetails.Alias,$MbxDetails.PrimarySMTPAddress,$MbxDetails.Identity,$MbxDetails.RecipientTypeDetails,$MbxDetails.WhenCreated,$MbxDetails.WhenChanged,$MbxDetails.RetentionPolicy,$MbxFldr.Count,$MbxStats.TotalItemSizeInBytes,$SizeGB.'TotalItemSize(GB)'
                Out-File -FilePath $strAllDetails -InputObject $text -Append

				If ($MbxFldr.Count -gt 949)
				{
					write-host "Mailbox with 950 Folders or More" $Mbx.Alias
					Out-File -FilePath $strTooManyFolders -InputObject $text -Append
				}

				If ($SizeGB.'TotalItemSize(GB)' -ge "30")
				{
					write-host "Mailbox 30GB or Larger" $Mbx.Alias
					Out-File -FilePath $strOver30GB -InputObject $text -Append
				}
			}
			else
			{
				write-host $cnt,$mbx.alias,"No Mailbox"
				$text = """{0}"",{1}" -f $Mbx.Alias,"No Mailbox"
				Out-File -FilePath $strAllDetails -InputObject $text -Append
			}
		$Cnt = $Cnt + 1
		}
	}
	write-host "Done"

    $InpReportfile = "\\usnbku134p\c$\ReportHistory\RoomMbxs\AllInputRoomMbxUPN_" + $Date + ".csv"
    $OutReportFile1 = "\\usnbku134p\c$\ReportHistory\RoomMbxs\AllRoomMbxDetails_" + $Date + ".csv"
    $OutReportFile2 = "\\usnbku134p\c$\ReportHistory\RoomMbxs\AllRoomMbxOvr30GBDetails_" + $Date + ".csv"
    $OutReportFile3 = "\\usnbku134p\c$\ReportHistory\RoomMbxs\AllRoomMbxOvr30GBDetails_" + $Date + ".csv"

    Move-Item c:\temp\Input-RoomMbxUPN.csv $InpReportFile
    Move-Item $strAllDetails $OutReportFile1
    Move-Item $strOver30GB $OutReportFile2
    Move-Item $strTooManyFolders $OutReportFile3
