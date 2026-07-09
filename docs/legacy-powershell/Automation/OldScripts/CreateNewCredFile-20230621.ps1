<#
#    This script create a new credential file that is used to connect to O365
#    Created by Sandi Glazebrook 3/28/2017
#
#   Called by:  O365AdminMenu.ps1
#
#>

$me = whoami
$CredENo = $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)).replace(".","")
$UsrName = $CredENo + "@global.ul.com"
$dir = "c:\users\" + $CredENo + "\documents\"
$File = "my" + $CredENo + "File.xml"
$CredFile = $dir + $File
Get-Credential -UserName $UsrName -Message 'Enter Password' | Export-Clixml $CredFile

write-host "New Credential File Created for: " $me