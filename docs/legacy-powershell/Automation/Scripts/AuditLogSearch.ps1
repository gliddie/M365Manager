#  Get Mailbox Audit Log Information:

$MbxAddr = read-host "`nEnter email address of mailbox"
$NoDays = read-host "Enter the number of days to search back"
$NoDaysBack = "-" + $NoDays

$Results = Search-MailboxAuditLog -Identity $MbxAddr -LogonTypes delegate -ShowDetails -StartDate (get-date).AddDays($NoDaysBack) -EndDate $(Get-Date)
$Results | select Operation, OperationResult, SourceItemSubjectsList, FolderPathName, MailboxOwnerUPN, LogonUserDisplayName, SourceItemFolderPathNamesList, LastAccessed | ogv  
 
pause