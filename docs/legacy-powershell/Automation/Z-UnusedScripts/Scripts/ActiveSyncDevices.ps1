#Retrieve all mailboxes in the Exchange organization 
$mailboxes = Get-Mailbox -ResultSize unlimited 
 
#Output file
$OutputFile = "ActiveSyncDevices.csv"

#Setup headers
Out-File -FilePath $OutputFile -InputObject "Alias,Name,CustomAttribute5,DeviceType,DeviceID,DeviceModel,DeviceFriendlyName,DeviceOS,LastSuccessSync" -Encoding UTF8
 
#Loop through each mailbox 
foreach ($mailbox in $mailboxes) { 
	$devices = Get-ActiveSyncDeviceStatistics -Mailbox $mailbox.samaccountname 
	 
	#If the current mailbox has an ActiveSync device associated, loop through each device 
	if ($devices) { 
		foreach ($device in $devices){ 
		  
			#Create a new object and add custom note properties for each device.  Comment out the ones you don't need
			$deviceobj = New-Object -TypeName psobject
			$deviceobj | Add-Member -Name Alias -Value $mailbox.Alias -MemberType NoteProperty
			$deviceobj | Add-Member -Name Name -Value $mailbox.Name -MemberType NoteProperty 
			$deviceobj | Add-Member -Name CustomAttribute5 -Value $mailbox.CustomAttribute5 -MemberType NoteProperty 
			$deviceobj | Add-Member -Name DeviceType -Value $device.DeviceType -MemberType NoteProperty
			$deviceobj | Add-Member -Name DeviceID -Value $device.DeviceID -MemberType NoteProperty
			$deviceobj | Add-Member -Name DeviceModel -Value $device.DeviceModel -MemberType NoteProperty
			$deviceobj | Add-Member -Name DeviceFriendlyName -Value $device.DeviceFriendlyName -MemberType NoteProperty 
			$deviceobj | Add-Member -Name DeviceOS -Value $device.DeviceOS -MemberType NoteProperty 
			$deviceobj | Add-Member -Name LastSuccessSync -Value ($device.LastSuccessSync).ToString("yyyy-MM-dd HH:mm:ss") -MemberType NoteProperty 
															
			#Write the custom object to the pipeline 
			Write-Output -InputObject $deviceobj 
			
			#Write line to file
			Out-File -FilePath $OutputFile -InputObject "$($deviceobj.Alias),$($deviceobj.Name),$($deviceobj.CustomAttribute5),$($deviceobj.DeviceType),$($deviceobj.DeviceID),$($deviceobj.DeviceModel),$($deviceobj.DeviceFriendlyName),$($deviceobj.DeviceOS),$($deviceobj.LastSuccessSync)" -Encoding UTF8 -append
			
		} 
	 
	} 
 
}
