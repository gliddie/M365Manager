#
#    This script create a new credential file that is used to connect to O365
#    Created by Sandi Glazebrook 3/28/2017
#    03/07/2019 - added text that the username must be entered as emp#@global.ul.com

$me = whoami
$dir = "c:\users\" + $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)) + "\documents\"
$File = "my" + ($me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))).replace(".","") + "File.xml"
$CredFile = $dir + $File
write-host "********************************************************" -ForegroundColor Red
write-host "** You MUST enter your username as emp#@global.ul.com **" -ForegroundColor Red
write-host "********************************************************" -ForegroundColor Red

Get-Credential | Export-Clixml $CredFile

write-host "New Credential File Created for: " $me