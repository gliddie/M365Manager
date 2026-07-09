#$cred = Get-Credential
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection
#$ImportCmd = Import-PSSession $Session

$userToFind = Read-Host -Prompt "Enter user to find (leave blank for all)"

$params = @{}
if([string]::IsNullOrEmpty($userToFind) -eq $false)
{$params = @{Identity = $userToFind}
}

#$UserMailboxStats = Get-Mailbox -ResultSize Unlimited @Params | Get-MailboxStatistics
$UserMailboxStats = Get-MailboxStatistics -Identity $UserToFind
$UserMailboxStats | Add-Member -MemberType ScriptProperty -Name TotalItemSizeInBytes -Value {$this.TotalItemSize -replace "(.*\()|,| [a-z]*\)", ""}
$UserMailboxStats | Select-Object DisplayName,TotalItemSizeInBytes,@{Name="TotalItemSize(GB)"; Expression={[math]::Round($_.TotalItemSizeInBytes/1GB,2)}}
write-host ""
$SizeGB = $UserMailboxStats | Select-Object @{Name="TotalItemSize(GB)"; Expression={[math]::Round($_.TotalItemSizeInBytes/1GB,2)}}
write-host "Mailbox size in GB = " $SizeGB.'TotalItemSize(GB)'
