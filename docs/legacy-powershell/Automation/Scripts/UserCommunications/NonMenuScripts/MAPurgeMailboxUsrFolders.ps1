#
#  Script to purge contents of user mailbox including deleting of User Created Folders
#

    Write-Host "Enter Employee Number for mailbox to clear: " -ForegroundColor Yellow -NoNewline
    $Usr = Read-Host
                    
    Get-MailboxStatistics $usr |fl DisplayName,ItemCount,TotalItemSize
                    
# http://kb.cloudiway.com/how-to-delete-a-mailbox-content-in-office-365-or-exchange/
#    Search-Mailbox -Identity $Usr -DeleteContent -Force


# Deleted folders
    get-mailboxfolderstatistics 38200 | Where-Object{$_.FolderType -eq "User Created" -band $_.ItemsInFolderAndSubFolders -eq 0} | ForEach-Object{
     
        "Deleting Folder " + $_.FolderPath
#        try
#        {  
            $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId((Convertid $_.FolderId))     
            $ewsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
            write-host $ewsFolder
            if($ewsFolder.TotalCount -eq 0)
            {
                $ewsFolder.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
                write-host $ewsFolder " - Folder Deleted"
            }
 #       }
 #       catch
 #           {
 #           }
    }
write-host "Removal of Messages and User Created Folders Complete for" $Usr
Get-MailboxStatistics $usr |fl DisplayName,ItemCount,TotalItemSize