$mbx = import-csv "c:\temp\OrchidMbx.csv"
foreach ($m in $mbx)
{
    $grp = ""
    write-host "Processing mailbox: " $m.mbx
    $perm = get-mailboxPermission $m.mbx |where-object {$_.User -ne "NT AUTHORITY\SELF"}
    if ($Perm.count -ne 0)
    {
        foreach ($p in $perm)
        {
            $newgrp = "*"+ $p.User + "*"
            If (($p.User -like "mbx*") -and ($grp -notlike $newgrp))
            {
                write-host "Individuals with" $p.AccessRights "to this mailbox"
                Get-DistributionGroupMember $p.User |ft PrimarySMTPAddress -ErrorAction SilentlyContinue
                $grp = $grp + $p.User + ","
            }
        }
    }
    $fldr = get-mailboxfolderPermission ($m.mbx+":\Inbox") |Where-Object {($_.AccessRights -notlike "None")}
    foreach ($f in $fldr)
    {
        If ($grp -notlike ("*" + $f.User.DisplayName + "*"))
        {
            If ($f.User.DisplayName -like "MBX*")
            {
                $grp = $grp + $f.User.DisplayName + ","
                write-host "Inidividuals with" $f.AccessRights "to this mailbox"
                Get-DistributionGroupMember $f.User.DisplayName |ft PrimarySMTPAddress -ErrorAction SilentlyContinue
            }
            else
            {
                If (($f.User.DisplayName -ne "Anonymous") -and ($f.User.DisplayName -ne "Default"))
                {
                    write-host $f.User.DisplayName "has" $f.AccessRights "to this mailbox"
                }
            }
        }
    }

    write-host "____________________________________________"
}


####This is the same but check for MBX group types

$mbx = import-csv "c:\temp\OrchidMbx.csv"
foreach ($m in $mbx)
{
	write-host "Processing Mailbox: " $m.mbx
	$root = Get-MailboxPermission $m.mbx.Trim() | Where-Object {$_.User -notlike "NT AUTHORITY\SELF"}
#	Get-MailboxPermission $m.mbx.Trim() | Where-Object {$_.User -notlike "NT AUTHORITY\SELF"} |FT User,AccessRights
	write-host "Mailbox Folder Permissions:"
	$Fldr = Get-MailboxfolderPermission ($m.mbx.Trim() + ":\Inbox") | Where-Object {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous")}
    $grp = ""
	
	Foreach ($r in $root)
	{
        $grp = $grp + $r + ","
        If (("*" + $r + "*") -notlike $grp)
        {
		    If ($r.User -like "MBX*")
			{
			    If ($r.user -like "*.ED")
				{
				    write-host "Individuals wiith FullAccess to " $m.mbx
				}
				If ($r.user -like "*.AU")
				{
				    write-host "Individuals wiith PublishingAuthor Access to " $m.mbx
				}
				If ($r.user -like "*.RE")
				{
				    write-host "Individuals wiith ReadOnly Access to " $m.mbx
				}				
				Get-DistributionGroupmember $r.user |ft PrimarySMTPAddress
		    }
        }
    }
	
	Foreach ($f in $fldr)
	{
        $grp = $grp + $f + ","
        If (("*" + $f + "*") -notlike $grp)
        {
#			write-host $f.user.DisplayName
			If (($f.user.DisplayName -like "MBX*") -and (($fldr.user.DisplayName -notlike $root.user)))
			{
				If ($f.user.DisplayName -like "*.ED")
				{
					write-host "Individuals wiith FullAccess to " $m.mbx
				}
				If ($f.user.DisplayName -like "*.AU")
				{
					write-host "Individuals wiith PublishingAuthor Access to " $m.mbx
				}
				If ($f.user.DisplayName -like "*.RE")
				{
					write-host "Individuals wiith ReadOnly Access to " $m.mbx
				}				
				Get-DistributionGroupmember $f.user.DisplayName |ft PrimarySMTPAddress
			}
            else
            {
                write-host $f.User.DisplayName "has" $f.AccessRights "to this mailbox"
            }
        }

	}
    write-host "____________________________________________"
}