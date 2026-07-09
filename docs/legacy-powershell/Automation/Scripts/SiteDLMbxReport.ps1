$DLOutPath = "C:\temp\O365_DLOwnershipReport.csv"
$DLtext = "DisplayName,DLName,ManagedBy,NoOfMembers,Notes,LastChanged,DLType"
Out-File -FilePath $DLOutPath -InputObject $DLtext

write-host "Enter 3 letter site code: " -ForegroundColor Cyan -nonewline
$SiteCode = read-host 

$Loc = "LST." + $SiteCode + "*"

write-host "Compiling list of all Distribution Lists"
$DLs = get-DistributionGroup -ResultSize Unlimited |where-object {$_.DisplayName -like "$Loc"}
$DLs.count
				
foreach ($DLs in $DLs)
{
    If ($DLs.RecipientTypeDetails -notlike "RoleGroup")
    {
        $DLMem = (Get-DistributionGroupMember $DLs.Name -ResultSize Unlimited).count
        If (($DLMem -ne 0) -and ($DLMem -lt 1))
        {
            $DLMem = 1
        }
    }
    else
    {
        $DLMem = 0
    }

    $DLs.Name
    $text = """{0}"",""{1}"",""{2}"",{3},""{4}"",""{5}"",{6}" -f $DLs.DisplayName,$DLs.Name,$DLs.ManagedBy,$DLMem,$DLs.Notes,$DLs.WhenChanged,$DLs.RecipientTypeDetails
    out-file -FilePath $DLOutPath -InputObject $text -Append
}

$DLOutPath = "C:\temp\O365_ShareMbxRoomResourceOwnershipReport.csv"
$DLtext = "DisplayName,MbxName,ManagedBy,NoOfMembers,Notes,LastChanged,MbxType"
Out-File -FilePath $DLOutPath -InputObject $DLtext

$ShrMbx = $Loc + "*"
$Mbxs = get-mailbox -ResultSize Unlimited | Where-Object {(($_.RecipientTypeDetails -eq "SharedMailbox") -and ($_.DisplayName -like "$ShrMbx"))}
$Mbxs.count

foreach ($Mbxs in $Mbxs)
{
    $MbxAccess = Get-MailboxPermission $Mbxs.DisplayNAME |Where-Object {($_.User -like "*global.ul.com*")}
}

 select DisplayName,PrimarySMTPAddress,ForwardingAddress,ForwardingSMTPAddress | Export-csv c:\temp\SharedMbxAudit.csv