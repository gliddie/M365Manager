#                $Mbx = Get-Recipient -Resultsize Unlimited |where {$_.RecipientTypeDetails -eq "UserMailbox"}
                $mbx = get-recipient "61177"
#                Write-Host "Total Number of Shared Mailboxes: " $Mbx.Count
            
#                write-host "Do you want to create a report file?  This will take some time to complete. (Y/N)? " -foregroundColor Yellow -Nonewline
#                $Rpt = read-host
                
#                If ($Rpt -eq "Y")
#                {

#                   Write-Host "Creating Report on c:\temp\AllUserMbxDetailsTst.csv" -ForegroundColor Yellow
#                   $strPath = "C:\temp\AllUserMbxDetailsTst.csv"
#                    Out-File -FilePath $strPath -InputObject $text

                    $text = "Emp#,DisplayName,MailboxType,WhenCreated,WhenChanged,RetentionPolicy,UserCreatedFolders,TotalSizeInBytes,TotalSize(GB)"
                    Out-File -FilePath c:\temp\AllUserMbxDetailsTst.csv -InputObject $text
                    Out-File -FilePath c:\temp\AllUserMbxTooManyFldrsDetailsTst.csv -InputObject $text
                    Out-File -FilePath c:\temp\AllUserMbxOvr30GBDetailsTst.csv -InputObject $text
                    
#                    foreach ($Mbx in $Mbx)
#                    {
                        $MbxFldr = Get-MailboxFolderStatistics $mbx.alias
                        $MbxStats = Get-MailboxStatistics -Identity $mbx.alias
                        $MbxStats | Add-Member -MemberType ScriptProperty -Name TotalItemSizeInBytes -Value {$this.TotalItemSize -replace "(.*\()|,| [a-z]*\)", ""}
                        $SizeGB = $MbxStats | Select-Object @{Name="TotalItemSize(GB)"; Expression={[math]::Round($_.TotalItemSizeInBytes/1GB,2)}}
#                        $MbxB = $MbxStats |  Select-Object TotalItemSizeInBytes
#                        $MbxGB = $MbxStats | Select-Object @{Name="TotalItemSize(GB)"; Expression={[math]::Round($_.TotalItemSizeInBytes/1GB,2)}}

                        $text = """{0}"",{1},{2},{3},{4},{5},{6},{7},{8}" -f $Mbx.Alias,$Mbx.Identity,$Mbx.RecipientTypeDetails,$Mbx.WhenCreated,$Mbx.WhenChanged,$Mbx.RetentionPolicy,$MbxFldr.Count,$MbxStats.TotalItemSizeInBytes,$SizeGB.'TotalItemSize(GB)'
                        Out-File -FilePath c:\temp\AllUserMbxDetailsTst.csv -InputObject $text -Append

                        $text
                        pause

                        If ($MbxFldr.Count -gt 949)
                        {
                             Out-File -FilePath c:\temp\AllUserMbxTooManyFldrsDetailsTst.csv -InputObject $text -Append
                        }

#                        $MbxSize = get-mailboxstatistics "61177" |select @{name=”TotalItemSize (GB)”; expression={[math]::Round(($_.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)}}

#                        If ($MbxStats.TotalItemSizeInBytes -ge "30000000000")
                        If ($SizeGB.'TotalItemSize(GB)' -ge "30")
                        {
                            write-host "Mailbox is larger than 30 GB"
                             Out-File -FilePath c:\temp\AllUserMbxOvr30GBDetailsTst.csv -InputObject $text -Append
                        }
#                    }
#               }