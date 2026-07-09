<#
#    This script create a new credential file that is used to connect to O365
#    Created by Sandi Glazebrook 3/28/2017
#
#   Called by:  O365AdminMenu.ps1
#
#>

write-host "Select (1) to create cred file for non-A account"
write-host "       (2) to create for A account"
write-host "       (3) to create for both accounts"
$Selection = Read-host "Choice"

If (($Selection -eq 1) -or ($Selection -eq 3))
{
    $me = whoami
    $CredENo = $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)).replace(".","")
    $UsrName = $CredENo + "@global.ul.com"
    $dir = "c:\users\" + $CredENo + "\documents\"
    $File = "my" + $CredENo + "File.xml"
    $CredFile = $dir + $File
    Get-Credential -UserName $UsrName -Message 'Enter Password' | Export-Clixml $CredFile

    write-host "New Credential File Created for: " $me
}

If (($Selection -eq 2) -or ($Selection -eq 3))
{
    $me = whoami
    $CredENo = $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)).replace(".","")
    $ACredENo = "A" + $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)).replace(".","")
    $UsrName = $ACredENo + "@global.ul.com"
    $dir = "c:\users\" + $CredENo + "\documents\"
    $File = "my" + $ACredENo + "File.xml"
    $CredFile = $dir + $File
    Get-Credential -UserName $UsrName -Message 'Enter Password' | Export-Clixml $CredFile

    write-host "New Credential File Created for: " $me
}