                Write-Host "Enter Employee Number for mailbox to clear" -ForegroundColor Yellow -NoNewline
                    $Usr = Read-Host
                    
                    Get-MailboxStatistics $usr |ft DisplayName,ItemCount
                    
# http://kb.cloudiway.com/how-to-delete-a-mailbox-content-in-office-365-or-exchange/
                    Search-Mailbox -Identity $Usr -DeleteContent -Force


# Deleted folders
# http://gsexdev.blogspot.in/2013/01/cleaning-out-skeletons-in-your-mailbox.html

                    $folders = get-mailboxfolderstatistics $MailboxName | Where-Object{$_.FolderType -eq "User Created" -band $_.ItemsInFolderAndSubFolders -eq 0} | ForEach-Object{  
# Bind to the Inbox Folder  
                    "Deleting Folder " + $_.FolderPath    
                    try
                    {  
                        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId((Convertid $_.FolderId))     
                        $ewsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
                        write-host $ewsFolder       
                        if($ewsFolder.TotalCount -eq 0)
                        {  
#                            $ewsFolder.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
                            write-host $ewsFolder " - Folder Deleted"
                        }
                    }
                    catch {}
                }