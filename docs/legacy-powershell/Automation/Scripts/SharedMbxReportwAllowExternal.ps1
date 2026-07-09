<#start-transcript
#
#  Called by:  ReportingMenu.ps1
#
# 2022-06-16 - SAG - Added details if an external senders are allowed
#>	

$date =get-date -Format yyyyMMdd

$ShrMbxOutPath = "C:\temp\O365_SharedMbxReport" + $Date + ".csv"

$MbxText = "MbxName,MbxAddr,MbxForwarder,ManagedBy,EDGroup,NoEDMembers,AuthGrp,NoAUMembers,ReadGrp,NoREMembers,InternalOnly,WhenCreated"
Out-File -FilePath $ShrMbxOutPath -InputObject $Mbxtext

write-host "Compiling list of all Shared Mailboxes....."
$ShrMbx = get-mailbox -ResultSize Unlimited -SortBy DisplayName|Where-Object {$_.RecipientTypeDetails -like "SharedMailbox"}
#$ShrMbx = get-mailbox "3D Printing - NBK"

write-Host "        Total Number of Shared Mailboxes: " $ShrMbx.count
			
foreach ($m in $ShrMbx)
{ 
    write-host "Processing " $m.Name "...."
    
    $Text = ""
    $MbxOwners = ""
    $MbxNotes = ""
    $FwdAddr = ""
    $EDGroup = ""
    $AUGroup = ""
    $REGroup = ""
    $EDMemCnt = ""
    $AUMemCnt = ""
    $REMemCnt = ""
    $GrpType = ""
    
    $ShrMbxDetails = Get-Mailbox $m.Name
    $ShrMbxGrps = Get-MailboxFolderPermission $m.Name |Where-Object {$_.User -like "*MBX*"}
    $EDGroup = (Get-MailboxPermission $m.Name |Where-Object {$_.User -like ".ED"}).Name
    If ($ShrMbxGrps.Length -eq 0)
    {
        $ShrMbxGrps = Get-MailboxPermission $m.Name |Where-Object {$_.User -like "*MBX*"}
        $EDGroup = (Get-MailboxPermission $m.Name |Where-Object {$_.User -like "*.ED"}).User
        $GrpType = $EDGroup.Substring(($EDGroup.length-3),3)
    }
    else
    {
        $ShrMbxGrps = Get-MailboxFolderPermission $m.Name |Where-Object {$_.User -like ".ED"}
        $EDGroup = $ShrMbxGrps.User.DisplayName
    }
        
    $MbxAddr = $ShrMbxDetails.EmailAddresses -join ';'
    $MbxFwd1 = $ShrMbxDetails.ForwardingSmtpAddress
    $MbxFwd2 = $ShrMbxDetails.ForwardingAddress
    If ($MbxFwd1.length -gt 0)
    {
        $FwdAddr = $MbxFwd1
    }
    else
    {
        If ($MbxFwd2.length -gt 0)
        {
            $FwdAddr = $MbxFwd2
        }
    }

    Foreach ($ShrMbxGrps in $ShrMbxGrps)
    {

        If ($GrpType.Length -eq 0)
        {
            $GrpType = $ShrMbxGrps.User.DisplayName.Substring(($ShrMbxGrps.User.DisplayName.length-3),3)
        }

        switch ($GrpType)
        {        
            ".ED"
            {
#                $EDGroup = $ShrMbxGrps.User.DisplayName
                $EDMemCnt = (Get-DistributionGroupMember $EDGroup).count
                $MbxOwners = ((Get-Group $EDGroup).ManagedBy) -join ", "
                $MbxNotes = (Get-Group $EDGroup).Notes
                If (($EDMemCnt -ne 0) -and ($EDMemCnt -lt 2))
                {
                    $EDMemCnt = 1
                }
            }

            ".AU"
            {
                $AUGroup = $ShrMbxGrps.User.DisplayName
                $AUMemCnt = (Get-DistributionGroupMember $ShrMbxGrps.User.DisplayName).count
                If ($MbxOwners.length -eq 0)
                {
                    $MbxOwners = ((Get-Group $AUGroup).ManagedBy) -join ", "
                }
                If (($AUMemCnt -ne 0) -and ($AUMemCnt -lt 2))
                {
                    $AUMemCnt = 1
                }
            }

            ".RE"
            {
                $REGroup = $ShrMbxGrps.User.DisplayName
                $REMemCnt = (Get-DistributionGroupMember $ShrMbxGrps.User.DisplayName).count
                If ($MbxOwners.length -eq 0)
                {
                    $MbxOwners = ((Get-Group $REGroup).ManagedBy) -join ", "
                }
                If (($REMemCnt -ne 0) -and ($REMemCnt -lt 2))
                {
                    $REMemCnt = 1
                }
            }
        }
    }
    $text = "{0},""{1}"",{2},""{3}"",{4},{5},{6},{7},{8},{9},{10},""{11}""" -f $m.DisplayName,$MbxAddr,$FwdAddr,$MbxOwners,$EDGroup,$EDMemCnt,$auGroup,$AUMemCnt,$REGroup,$REMemCnt,$m.RequireSenderAuthenticationEnabled,$m.WhenCreated
    out-file -FilePath $ShrMbxOutPath -InputObject $text -Append
}