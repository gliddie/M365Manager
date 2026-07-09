$me = whoami
$CredENo = $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))
$AAct = "A" + $CredENo
$Pwd = read-host "Enter password for" $ADOnPrem "account" -AsSecureString
$Server = "usnbkadds006p.global.ul.com"
$cred= New-Object System.Management.Automation.PSCredential($AAct,$Pwd)

$ENo = read-host "Enter employee number of new SDAP Team Member"

Set-User $ENo -RemotePowerShellEnabled $True -Confirm:$False

Add-MsolRoleMember -RoleMemberEmailAddress $upn -RoleName "User Administrator"
Add-RoleGroupMember -Identity "UL Account Provisioning" -Member $upn
Add-ADGroupMember ACL.UL.IntuneServiceDesk -Members $ENo

#These three commands need to be run with Admin Creds
Add-ADGroupMember ADMIN.UL.SDAdmin -Members $ENo -Credential $Cred -Server $Server
Add-ADGroupMember ADMIN.UL.ServiceDeskAccountProvisioning -Members $Cred -Cred $ADOnPrem -Server $Server
Add-ADGroupMember ADMIN.UL.UNISYS.LYNCPROVISIONER -Members $ENo -Cred $Cred -Server $Server