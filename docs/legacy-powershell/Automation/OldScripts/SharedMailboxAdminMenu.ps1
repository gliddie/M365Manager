<####
#### Shared Mailbox Admin Menu
#
#  Called by:  O365AdminMenu.ps1
#
#  04/15/2020 - SAG - Combined the AddMailboxPermissionSharedMailbox and ApplyRetentionPolicy scripts into the AddFolderPermission script for the Option 1 and 3
#
####>


$MbxChg = "1"

Do
{
    write-host ""
    write-host "Shared Mailbox Admin Menu" -ForegroundColor Magenta
    write-host
    Write-Host "     Enter ( 1) Create a New Shared Mailbox"
    write-host "           ( 2) Change Shared Mailbox Ownership"
    write-host "           ( 3) Add Additional Groups (.ED, .AU or .RE) for Mailbox Access"
    write-host "           ( 4) Rename Shared Mailbox"
    write-host "           ( 5) Remove Shared Mailbox"
    write-host "           ( 6) Add/Removed Email Alias from Shared Mailbox"
    write-host "           ( 7) View Mailbox Folder Statistics"
    write-host ""
    write-Host "           ( 0) to Return to the O365 Admin Menu"
    write-Host "     Enter Option? " -ForegroundColor Red -NoNewline
    $MbxChg = Read-Host
    write-host ""

    If ($MbxChg -ne 0)
    {
        invoke-Expression -Command .\SharedMailboxProcessInstructions.ps1
    }
        		
		switch ($MbxChg)
		{

           1
                {
                    write-host "Executing the Creation of the New Mailboxes hit return when ready" -ForegroundColor Yellow -NoNewline
                    $Cont = read-host
                    $ExeOK = "Y"
                    Do {
                        invoke-expression -Command .\NewSharedMailbox.ps1
                        Write-Host "Do you need to run this script again due to errors (Y/N)?" -ForegroundColor Yellow -NoNewline
                        $ExeOK = Read-Host
                    } while ($ExeOK -eq "Y")

                    write-host ""
                    write-host "Executing the Creation of the New Security Groups hit return when ready" -ForegroundColor Yellow -NoNewline
                    $Cont = read-host
                    $ExeOK = "Y"
                    Do {
                        invoke-expression -Command .\NewUSG.ps1
                        Write-Host "Do you need to run this script again due to errors (Y/N)?" -ForegroundColor Yellow -NoNewline
                        $ExeOK = Read-Host
                    } while ($ExeOK -eq "Y")

                    write-host ""
#                    write-host "Applying Mailbox Permissions to the New Mailboxes hit return when ready" -ForegroundColor Yellow -NoNewline
#                    $Cont = read-host
#                    $ExeOK = "Y"
#                    Do {
#                        invoke-expression -Command .\AddMailboxPermissionSharedMailbox.ps1
#                        Write-Host "Do you need to run this script again due to errors (Y/N)?" -ForegroundColor Yellow -NoNewline
#                        $ExeOK = Read-Host
#                    } while ($ExeOK -eq "Y")

                    write-host ""
#                    write-host "Applying Folder Permissions to the New Mailboxes hit return when ready" -ForegroundColor Yellow -NoNewline
                    write-host "Applying Permissions to the New Mailboxes hit return when ready" -ForegroundColor Yellow -NoNewline
                    $Cont = read-host
                    $ExeOK = "Y"
                    Do {
                        invoke-expression -Command .\AddFolderPermissions.ps1
                        Write-Host "Do you need to run this script again due to errors (Y/N)?" -ForegroundColor Yellow -NoNewline
                        $ExeOK = Read-Host
                    } while ($ExeOK -eq "Y")

#                    write-host ""
#                    write-host "Applying Retention Policy to the New Mailboxes hit return when ready" -ForegroundColor Yellow -NoNewline
#                    $Cont = read-host
#                    $ExeOK = "Y"
#                    Do {
#                        invoke-expression -Command .\ApplyRetentionPolicy.ps1
#                        Write-Host "Do you need to run this script again due to errors (Y/N)?" -ForegroundColor Yellow -NoNewline
#                        $ExeOK = Read-Host
#                        } while ($ExeOK -eq "Y")

                    write-host ""
                    write-host "Shared Mailbox Creation and Set-up Complete" -ForegroundColor Magenta
                    pause
                }
            2
                {
                    write-host "Changed Shared Mailbox Ownership"

                    Write-Host "Enter the Name of the Shared Mailbox: " -ForegroundColor Yellow -NoNewline
                    $Global:ShrMbx = Read-Host
                    $ShrExists = [bool](Get-Mailbox $Global:ShrMbx -ErrorAction SilentlyContinue)
                    If ($ShrExists -eq "True")
                    {
                        write-host "Mailbox owners are: " (get-mailbox $Global:ShrMbx).MailTipTranslations
                        write-host "Below are the individuals and groups that have access to the mailbox: " -ForegroundColor Yellow
                        $MoreGrps = "Y"
                        Get-mailboxpermission $Global:ShrMbx | where {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous") -and ($_.IsInherited -notlike "True")}
                        Get-MailboxFolderPermission $Global:ShrMbx| where {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous")} | ft
                        Do
                        {
                            invoke-expression -Command .\GroupOwnershipChanges.ps1
                            Write-Host "Do you need to modify the ownership for more access groups (Y/N)? " -ForegroundColor Yellow -NoNewline
                            $MoreGrps = Read-Host
                        } while ($MoreGrps -ne "N")
                        Write-Host 
                    }
                    else
                    {
                        Write-Host ""
                        Write-Host "Shared Mailbox does not Exist" -ForegroundColor Red
                    }
                }
            3
                {
                    write-host "Add Addtional Shared Mailbox Security Groups"
                    
                    write-host ""
                    write-host "Executing the Creation of the New Security Groups hit return when ready" -ForegroundColor Yellow -NoNewline
                    $Cont = read-host
                    $ExeOK = "Y"
                    Do {
                        invoke-expression -Command .\NewUSG.ps1
                        Write-Host "Do you need to run this script again due to errors (Y/N)? " -ForegroundColor Yellow -NoNewline
                        $ExeOK = Read-Host
                    } while ($ExeOK -eq "Y")

#                    write-host ""
#                    Write-Host "Do any of the new groups grant Editor Access to the Shared Mailbox (Y/N)? " -ForegroundColor Yellow -NoNewline
#                    $EdiAccess = Read-Host

#                    If ($EdiAccess -eq "Y")
#                    {
#                        write-host ""
#                        write-host "Applying Mailbox Permissions to the New Mailboxes hit return when ready" -ForegroundColor Yellow -NoNewline
#                        $Cont = read-host
#                        $ExeOK = "Y"
#                        Do {
#                            invoke-expression -Command .\AddMailboxPermissionSharedMailbox.ps1
#                            Write-Host "Do you need to run this script again due to errors (Y/N)? " -ForegroundColor Yellow -NoNewline
#                            $ExeOK = Read-Host
#                        } while ($ExeOK -eq "Y")
#                    }

                    write-host ""
#                    write-host "Applying Folder Permissions to the New Mailboxes hit return when ready" -ForegroundColor Yellow -NoNewline
                    write-host "Applying Permissions to the New Mailboxes hit return when ready" -ForegroundColor Yellow -NoNewline
                    $Cont = read-host
                    $ExeOK = "Y"
                    Do {
                        invoke-expression -Command .\AddFolderPermissions.ps1
                        Write-Host "Do you need to run this script again due to errors (Y/N)? " -ForegroundColor Yellow -NoNewline
                        $ExeOK = Read-Host
                    } while ($ExeOK -eq "Y")

                    Write-Host
                }
            4
                {
                    write-host "Rename Shared Mailbox"
                    Write-Host "Enter Existing Name of the Shared Mailbox to Rename: " -ForegroundColor Yellow -NoNewline
                    $ShrMbx = Read-Host
                    $ShrExists = [bool](Get-Mailbox $ShrMbx -ErrorAction SilentlyContinue)
                    If ($ShrExists -eq "True")
                    {
                        $ShrMbxInf = Get-Mailbox $ShrMbx
                        Write-Host "Enter New Shared Mailbox Name: " -ForegroundColor Yellow -NoNewline
                        $NewName = Read-Host
                        If ($NewName.IndexOf(" ") -eq 3)
                        {
                            $NewAlias = ($NewName.Insert($NewName.IndexOf(" "),"."))
                            $NewAlias = $NewAlias -replace '\s',''
                            if ($NewAlias -like "*&*")
                            {
                                $NewAlias = $NewAlias.Remove("&")
                            }
                            if ($NewAlias -like "*,*")
                            {
                                $NewAlias = $NewAlias.Remove(",")
                            }
                            if ($NewAlias -like "*-*")
                            {
                                $NewAlias = $NewAlias.Remove("-")
                            }
                            if ($NewAlias -like "*/*")
                            {
                                $NewAlias = $NewAlias.Remove("/")
                            }

                            $NewInetAddr = $NewAlias + "@ul.com"
                        }
                        else
                        {
                            write-host "Enter New Shared Mailbox Internet Adress: " -ForegroundColor Yellow -NoNewline
                            $NewInetAddr = Read-Host
                            pause
                            $NewAlias = $NewInetAddr.Remove($NewInetAddr.IndexOf("@"))
                        }

                        write-host "List of Who Has Access to the Mailbox: " -ForegroundColor Yellow
                        Get-MailboxFolderPermission $ShrMbx | where {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous") -and ($_.IsInherited -notlike "True")}

                        $NewPrimaryAlias = $ShrMbxInf.EmailAddresses += "SMTP:" + $NewInetAddr

                        write-host "Renaming Mailbox from" $ShrMbx "to" $NewName
                        Set-Mailbox $ShrMbx -Name $NewName -DisplayName $NewName -Alias $NewAlias -EmailAddresses $NewPrimaryAlias

                        write-host "Do you have access groups that need to be renamed (Y/N)? " -ForegroundColor Yellow -NoNewline
                        $RenAccessGrps = Read-Host

                        Do {
                            Write-Host "Enter the Name of the Access Group to be renamed " -ForegroundColor Yellow -NoNewline
                            $DSTGrp = Read-Host
                            Write-Host "Enter the New Access Group Name " -ForegroundColor Yellow -NoNewline
                            $NewDSTGrp = Read-Host
                            write-host "Enter Service Desk Ticket Number " -ForegroundColor Yellow -NoNewline
                            $SDTicketNo = Read-Host
                
                            $DSTGrpDetails = Get-DistributionGroup $DSTGrp
                            $GrpInf = get-Group $DSTGrp
                            $NewDSTAlias = $NewDSTGrp.Replace(" ","")
                            $NewDSTAlias = $NewDSTAlias.Replace("&","")
                            $NewDSTAlias = $NewDSTAlias.Replace("-","")
                            $NewDSTInetAddr = $NewDSTAlias + "@ul.com"

                            Write-Host ""
                            Write-Host "Renaming" $DSTGrp " to" $NewDSTGrp -ForegroundColor Yellow
                            write-host "            New Access Group Name: " $NewDSTGrp
                            Write-Host "           New Access Group Alias: " $NewDSTAlias
                            Write-Host "New Access Group Internet Address: " $NewDSTInetAddr

                            write-host "Rename this Access Group (Y/N)? " -ForegroundColor Yellow -NoNewline
                            $ContGrp = Read-Host
                            If ($ContGrp -eq "Y")
                                                                                                                                            {
                        If ($GrpInf.Notes -like "*Per*")
                        {
                            $GrpInf.Notes = $GrpInf.Notes.Remove($GrpInf.Notes.IndexOf("Per")) + "Per " + $SDTicketNo
                        }
                        else
                        {
                            write-host "The notes on this group are not in the standard format. The current Notes contains " $GrpInf.Notes ":"
                            $GrpOwners = write-host "Enter the group owners: " -ForegroundColor Yellow
                            $GrpInf.Notes = "Owner: " + $GrpOwners + " - Per " + $SDTicketNo
                        }

                        Set-Group $DSTGrp -Notes $GrpInf.Notes
                        Set-DistributionGroup $DSTGrp -DisplayName $NewDSTGrp -Name $NewDSTGrp -Alias $NewDSTAlias -EmailAddresses @{add=$NewDSTInetAddr}
                        Set-DistributionGroup $NewDSTGrp -PrimarySmtpAddress $NewDSTInetAddr
                    }
                            else
                            {
                                Write-Host "No Changes Made to this Group " $MbxAccess.User
                            }
                            $ContGrp = ""
                            Write-Host "Do you have more access groups that need to be renamed (Y/N)? " -ForegroundColor Yellow -NoNewline
                            $RenAccessGrps = Read-Host
                        } while ($RenAccessGrps -eq "Y")
                    }
                    else
                    {
                        write-host ""
                        Write-Host "Shared mailbox does not exist" -ForegroundColor Red
                    }

                }
            5
                {
                    invoke-expression -Command .\RemoveSharedMailbox.ps1
                }
            6
                {
                    Write-Host "Add/Remove Additional eMail Alias to a Shared Mailbox" -ForegroundColor Magenta
                    Write-Host "Enter the name of the Shared Mailbox " -ForegroundColor Yellow -NoNewline
                    $Mbx = Read-Host
                    $ShrExists = [bool](Get-Mailbox $Mbx -ErrorAction SilentlyContinue)
                    If ($ShrExists -eq "True")
                    {
                        $MbxAlias = Get-Mailbox $Mbx
                        write-host "Current Mailbox Owners: " -ForegroundColor Yellow -NoNewline
                        $MbxGroup = "mbx." + $mbx + ".ED"
                        $MbxOwner = Get-DistributionGroup $MbxGroup
                        write-host $MbxOwner.ManagedBy            
                        Write-Host "Current Email aliases on" $Mbx ": " -ForegroundColor Yellow -NoNewline
                        write-host $MbxAlias.EmailAddresses
                        write-host ""
                        write-host "Do you wish to Add or Remove an Alias enter (A = Add/R = Remove)? " -ForegroundColor Yellow -NoNewline
                        $AddRem = Read-Host
                        If ($AddRem -eq "A")
                        {
                            write-host ""
                            write-host "Enter the eMail Alias you would like to add: " -ForegroundColor Yellow -NoNewline
                            $NewAlias = Read-Host
                            $AddAlias = "smtp:" + $NewAlias
                            Set-Mailbox $Mbx -EmailAddresses @{Add=$AddAlias}
                            Get-Mailbox $Mbx |ft *Addresses*
                            Write-Host "Should this be made the new Primary SMTP Address for this Mailbox (Y/N) ?" -ForegroundColor Yellow -NoNewline
                            $NewPrim = Read-Host
                            If ($NewPrim -eq "Y")
                            {
                                Set-Mailbox $Mbx -PrimarySmtpAddress $NewAlias
                            }
                        }
                        elseif ($AddRem -eq "R")
                        {
                            write-host ""
                            write-host "Enter the eMail Alias you would like to remove: " -ForegroundColor Yellow -NoNewline
                            $RemAlias = Read-Host
                            $RemAlias = "smtp:" + $RemAlias
                            Set-Mailbox $Mbx -EmailAddresses @{Remove=$RemAlias}
                            Get-Mailbox $Mbx |ft *Addresses*
                        }
                        else
                        {
                            Write-Host "No changes made to the configured aliases for this Distribution list"
                        }
                        pause
                    }
                    else
                    {
                        Write-Host ""
                        Write-Host "Shared mailbox does not exist"
                    }
                }
            7
                {
                    invoke-expression -Command .\MailboxStatistics.ps1
                }
         }

}While ($MbxChg -ne 0)	