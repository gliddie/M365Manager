#
#	Connect to SharePoint Online O365 Services
#	
#	Check to see if there is a c:\temp directory if not create one
#	Also check to see if there is a Credential file for this individual if
#	  if not create one.
#

$me = whoami
$dir = "c:\users\" + $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)) + "\documents\"
$File = "my" + ($me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))).replace(".","") + "File.xml"
$moveCredFile = "c:\temp\" + $File
$CredFile = $dir + $File

If (Test-Path $moveCredFile)
{
	move-item $moveCredFile -destination $CredFile
}

write-host "Connecting to SPO Office 365....." -ForegroundColor Cyan
If ((Test-path $CredFile) -ne "True")
{
    Get-Credential | Export-Clixml $CredFile
}
$LiveCred = Import-Clixml $CredFile
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $LiveCred -Authentication Basic -AllowRedirection
Import-PSsession $Session -AllowClobber