#MessageCounts for MBX groups

$End = get-date
$Strt = $End.AddDays(-10)
$grp = get-DistributionGroup -ResultSize Unlimited |Where-Object {$_.Name -like "MBX*"}
$File = "c:\temp\MBXGroupAudit-" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
write-host "GrpName,Address,MsgRcvd,NoMbrs,NoMgrs"
Start-Transcript $File

Foreach ($v in $grp)
{
    If ($v.Name -notlike "*RRS*")
    {
        $msg = 0
        $msg = Get-MessageTrace -RecipientAddress $v.PrimarySmtpAddress -StartDate $Strt -EndDate $end
        $mbrs = (get-distributiongroupmember $v.PrimarySmtpAddress -ResultSize Unlimited).count
        $noMgrs = (get-distribtuiongroup $v.PrimarySmtpAddress).count
        write-host $v.DisplayName " ," $v.PrimarySMTPAddress " , "  $msg.count " , " $mbrs " , "  $noMgrs

        If ($msg.count -ne 0)
        {
            $mbxmsg = ""
            $mbx = $v.Alias.substring(4,$v.Alias.length-7) + "@ul.com"
            $Exists = [bool](get-mailbox $mbx -ErrorAction SilentlyContinue)
            If ($Exists -eq $True)
            {
                $msgRec = Get-MessageTrace -RecipientAddress $mbx -StartDate $Strt -EndDate $end
                $msgCnt = get-mailboxstatistics $mbx
                write-host $mbx " , "  $mbxRec.count
                Get-MailboxFolderStatistics $mbx -folderScope "Inbox" -IncludeOldestAndNewestItems |ft Identity,OldestItemReceivedDate,NewestItemReceivedDate
                Get-MailboxFolderStatistics $mbx -folderScope "SentItems" -IncludeOldestAndNewestItems |ft Identity,OldestItemReceivedDate,NewestItemReceivedDate
            }
            else
            {
                write-host $mbx ", Mailbox Not found"
            }
        }
    }
}
Stop-Transcript

#MessageCounts for LST groups

$End = get-date
$Strt = $End.AddDays(-10)
$grp = get-DistributionGroup -ResultSize Unlimited |Where-Object {$_.Name -like "LST*"}
$File = "c:\temp\LSTGroupAudit-" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
Start-Transcript $File
write-host "GrpName,Address,MsgRcvd,NoMbrs,NoMgrs"

Foreach ($v in $grp)
{
    If ($v.Name -notlike "*RRS*")
    {
        $msg = 0
        $msg = Get-MessageTrace -RecipientAddress $v.PrimarySmtpAddress -StartDate $Strt -EndDate $end
        $NoMbrs = (Get-DistributionGroupMember $v.PrimarySmtpAddress -ResultSize Unlimited).count
        write-host $v.DisplayName " ," $v.PrimarySMTPAddress " , "  $msg.count " , " $NoMbrs " , " $v.ManagedBy.Count
    }
}
Stop-Transcript

#Oldest/Newest Messages for Shared Mbx

$File = "c:\temp\ShrMBXAudit-" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
Start-Transcript $File
$End = get-date
$Strt = $End.AddDays(-10)
$mbx = get-mailbox -ResultSize Unlimited |Where-Object {$_.RecipientTypeDetails -eq "SharedMailbox"}
write-host $mbx.count

write-host "Mailbox,whenCreated,Forwards,MsgReceived,MsgSent,InboxOldestReceived,InboxNewestReceived,SentOldestSent,SentNewestSent"

Foreach ($v in $mbx)
{
    $msgs = ""
    $SntMsgs = ""
    $InOldRcv = ""
    $InNewRcv = ""
    $InOldSnt = ""
    $InNewSnt = ""
    $Forwards = "False"
    $mbx = get-mailbox $v.PrimarySmtpAddress
    $msgrcv = Get-MessageTrace -RecipientAddress $v.PrimarySmtpAddress -StartDate $Strt -EndDate $end
    $msgsnt = Get-MessageTrace -SenderAddress $v.PrimarySmtpAddress -StartDate $Strt -EndDate $end

    $msgs = Get-MailboxFolderStatistics $v.PrimarySMTPAddress -folderScope "Inbox" -IncludeOldestAndNewestItems
    Foreach ($m in $msgs)
    {
        If ($m.Identity -like "*Inbox")
        {
            If ($m.OldestItemReceivedDate -ne $null)
            {
                $InOldRcv = $m.OldestItemReceivedDate.Tostring("yyyy-MM-dd")
            }
            If ($m.NewestItemReceivedDate -ne $null)
            {
                $InNewRcv = $m.NewestItemReceivedDate.Tostring("yyyy-MM-dd")
            }
        }
    }
    If (($mbx.FowardingAddress -ne $null) -or ($mbx.ForwardingSmtpAddress -ne $null))
    {
        If (($mbx.FowardingAddress -like "*salesforce*") -or ($mbx.ForwardingSmtpAddress -like "*salesforce*"))
        {
            $Forwards = "SalesForce"
        }
        else
        {
            $Forwards = "True"
        }
    }
    $SntMsgs = Get-MailboxFolderStatistics $v.PrimarySMTPAddress -folderScope "SentItems" -IncludeOldestAndNewestItems
    If ($SntMsgs.OldestItemReceivedDate -ne $null)
    {
        $InOldSnt = $SntMsgs.OldestItemReceivedDate.Tostring("yyyy-MM-dd")
    }
    If ($SntMsgs.NewestItemReceivedDate -ne $null)
    {
        $InNewSnt = $m.NewestItemReceivedDate.Tostring("yyyy-MM-dd")
    }
    write-host $SntMsgs.Identity " , " $SntMsgs.OldestItemReceivedDate " , " $SntMsgs.NewestItemReceivedDate
    write-host $v.PrimarySMTPAddress " , " $mbx.whenCreated " ," $Forwards " , " $msgrcv.count " , " $msgsnt.count " , " $InOldRcv " , " $InNewRcv " , " $InOldSnt " , " $InNewSnt
}
Stop-Transcript

#Audit MBX group for a shared mailbox
$End = get-date
$Strt = $End.AddDays(-10)
$grp = get-DistributionGroup -ResultSize Unlimited |Where-Object {$_.Name -like "MBX*"}
$File = "c:\temp\ShrMBXGroupMatchAudit-" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
write-host "GrpName,MailboxAddress,Status,MailTip"
Start-Transcript $File

foreach ($g in $grp)
{
    If (($g.Name -notlike "*RRS*") -and ($g.Name -notlike "*.RGS*"))
    {
        $Marked = ""
        $mbx = $g.Alias.substring(4,$g.Alias.length-7) + "@ul.com"
        $Exists = [bool](get-mailbox $mbx -ErrorAction SilentlyContinue)
        If ($Exists -eq $True)
        {
            If ($g.MailTip -like "*Message sent*")
            {
                $Marked = "Requested Owner"
            }
            write-host $g.Name "`t" $mbx "`t" "Mailbox Exists" "`t" $Marked
        }
        else
        {
            write-host $g.Name "`t" $mbx "`t" "Mailbox Not Found" "`t" $Marked
        }
    }
}