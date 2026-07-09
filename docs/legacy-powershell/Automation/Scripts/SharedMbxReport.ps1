<#start-transcript
#
#  Called by:  ReportingMenu.ps1
#
#>

$date =get-date -Format yyyyMMdd

$ShrMbxOutPath = "C:\temp\O365_SharedMbxReport" + $Date + ".csv"

$MbxText = "MbxName,MbxAddr,MbxForwarder,ManagedBy,NoofMgrs,EditorsGrp,NoMembers,AuthGrp,NoMembers,ReadGrp,NoMembers,Notes"
Out-File -FilePath $MbxOutPath -InputObject $Mbxtext

write-host "Compiling list of all Shared Mailboxe
$ShrMbx = get-mailbox -ResultSize Unlimited |Where-Object {$_.RecipientTypeDetails -like "SharedMailbox"} -SortBy DisplayName

write-Host "        Total Number of Shared Mailboxes: " $ShrMbx.count
				
foreach ($ShrMbx in $ShrMbx)
{ 
    write-host "Processing " $ShrMbx.Name "...."
    
    $EDGroup = ""
    $AUGroup = ""
    $REGroup = ""
    $MbxOwners = ""
    $MbxNotes = ""
    $EDMemCount = ""
    $ADMemCount = ""
    $REMencount = ""
    
    $ShrMbxDetails = Get-Mailbox $ShrMbx.Name
    $ShrMbxGrps = Get-MailboxFolderPermission $ShrMbx.Name |Where-Object {$_.User -like "*MBX*"}
    $EDGrpExists = [bool](Get-MailboxFolderPermission $ShrMbx.Name |Where-Object {$_.User -like "*.ED"})
        
    $MbxAddr = $ShrMbxDetails.EmailAddresses -join ';'
    $MbxFwd = $ShrMbxDetails.ForwardingSmtpAddress
    $MbxFwd1 = $ShrMbxDetails.ForwardingAddress

    Foreach ($ShrMbxGrps in $ShrMbxGrps)
    {

        $GrpType = SShrMbxGrps.name.Substring(($grp.name.length-3),3)

        select $GrpType
        {        
            ".ED"
            {
                $EDGroup = $ShrMbxGrps.Name
                $EDMemCnt = (Get-DistributionGroupMember $ShrMbxGrps.Name).count
                $MbxOwners = Get-Group $ShrMbxGrps.ManagedBy
                $MbxNotes = Get-Group $ShrMbxGrps.Notes
            }

            "*.AU"
            {
                $AUGroup = $ShrMbxGrps.Name
                $AUMemCnt = (Get-DistributionGroupMember $ShrMbxGrps.Name).count
            }

            "*.RE"
            {
                $REGroup = $ShrMbxGrps.Name
                $REMemCnt = (Get-DistributionGroupMember $ShrMbxGrps.Name).count
            }
        }
        
        If ($EDGrpExists -eq $False)
        {
            $MbxOwners = Get-Group $ShrMbxGrps.ManagedBy
            $MbxNotes = Get-Group $ShrMbxGrps.Notes
        }


    }
}

        $text = "{0},{1},{2},{3},{4},{5},{6},""{7}""" -f $DLs.DisplayName,$DLs.Name,$DLs.ManagedBy.Count,$DLManagedBy,$DLMem,$DLs.WhenChanged,$DLs.RecipientTypeDetails,$DLs.Notes
        out-file -FilePath $DLOutPath -InputObject $text -Append