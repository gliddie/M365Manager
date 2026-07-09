                Write-Host "Starting All Discovery Mailbox Report This Will take some time......"
                Write-Host ""

                $FldrRpt = 0
                $SizeRpt = 0
				
                $Mbx = Get-Recipient -Resultsize Unlimited |where {$_.RecipientTypeDetails -eq "DiscoveryMailbox"}
                Write-Host "Total Number of Discovery Mailboxes: " $Mbx.Count
				
                write-host "Do you want to create a report file?  This will take some time to complete. (Y/N)? " -foregroundColor Yellow -Nonewline
                $Rpt = read-host
                
                If ($Rpt -eq "Y")
                {
                   Write-Host "Creating Report on c:\temp\AllDiscoveryMbxDetails.csv" -ForegroundColor Yellow
                   $strAllDetails = "C:\temp\AllDiscoveryMbxDetails.csv"
                   $strTooManyFolders = "C:\temp\AllDiscoveryTooManyFldrsDetails.csv"
                   $strOver30GB = "C:\temp\AllDiscoveryMbxOvr30GBDetails.csv"
				   
                   $headers = "Emp#,PrimarySMTPAddress,DisplayName,MailboxType,WhenCreated,WhenChanged,RetentionPolicy,UserCreatedFolders,TotalSizeInBytes,TotalSize(GB)"
                   Out-File -FilePath $strAllDetails -InputObject $headers
                    
                   foreach ($Mbx in $Mbx)
                   {
                        $MbxFldr = Get-MailboxFolderStatistics $mbx.alias
                        $MbxStats = Get-MailboxStatistics -Identity $mbx.alias
                        $MbxStats | Add-Member -MemberType ScriptProperty -Name TotalItemSizeInBytes -Value {$this.TotalItemSize -replace "(.*\()|,| [a-z]*\)", ""}
                        $SizeGB = $MbxStats | Select-Object @{Name="TotalItemSize(GB)"; Expression={[math]::Round($_.TotalItemSizeInBytes/1GB,2)}}

                        $text = """{0}"",{1},""{2}"",{3},{4},{5},{6},{7},{8},{9}" -f $Mbx.Alias,$Mbx.PrimarySMTPAddress,$Mbx.Identity,$Mbx.RecipientTypeDetails,$Mbx.WhenCreated,$Mbx.WhenChanged,$Mbx.RetentionPolicy,$MbxFldr.Count,$MbxStats.TotalItemSizeInBytes,$SizeGB.'TotalItemSize(GB)'
                        Out-File -FilePath $strAllDetails -InputObject $text -Append

                        If ($MbxFldr.Count -gt 949)
                        {
							If ($FldrRpt -eq 0)
							{
								Out-File -FilePath $strTooManyFolders -InputObject $headers
								$FldrRpt = 1
							}
							write-host "Mailbox with 950 Folders or More" $Mbx.Alias
							Out-File -FilePath $strTooManyFolders -InputObject $text -Append
                        }

                        If ($SizeGB.'TotalItemSize(GB)' -ge "30")
                        {
                            If ($SizeRpt -eq 0)
                            {
                                Out-File -FilePath $strOver30GB -InputObject $headers
                                $SizeRpt = 1
                            }
							write-host "Mailbox 30GB or Larger" $Mbx.Alias
							Out-File -FilePath $strOver30GB -InputObject $text -Append
                        }
                    }
               }