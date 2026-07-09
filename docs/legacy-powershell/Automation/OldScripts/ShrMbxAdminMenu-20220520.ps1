<####
#### Shared Mailbox Admin Menu
#
#  Called by:  O365AdminMenu.ps1
#
#  04/15/2020 - SAG - Combined the AddMailboxPermissionSharedMailbox and ApplyRetentionPolicy scripts into the AddFolderPermission script for the Option 1 and 3
#  05/20/2022 - SAG - Added Shared Mailbox Removal with Input File
#
####>

function Build-ShrMbxMenuForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Shared Mailbox Admin Admin Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 550 ; $form.Height = 300  # Make the form wider

    $TopLoc = 20 
    ## Label and TextBox  
    ## Title Line
    $Global:lblDLSecTitleLine = New-Object System.Windows.Forms.Label   
        $lblDLSecTitleLine.Text = "Select Option:"
        $lblDLSecTitleLine.Top = $TopLoc ; $lblDLSecTitleLine.Left = 60; $lblDLSecTitleLine.Width=220 ;$lblDLSecTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($lblDLSecTitleLine)    # Add to Form 

    ## New ShrMbx
    $TopLoc = $TopLoc + 20
    $Global:chkNewShrMbx = New-Object Windows.Forms.RadioButton
        $Global:chkNewShrMbx.Left = 100; $Global:chkNewShrMbx.Width = 450; $Global:chkNewShrMbx.Top = $TopLoc
        $Global:chkNewShrMbx.Text = "Create a New Shared Mailbox" 
        $Global:chkNewShrMbx.Checked = $false   # set a default value 
        $Global:chkNewShrMbx.TabIndex = 1
        $Global:form.Controls.Add($Global:chkNewShrMbx) 
        # Obtain Value with: $Global:chkNewShrMbx.Checked
#        $Global:InputFocus = $Global:chkNewShrMbx

    ## Change Ownership
    $TopLoc = $TopLoc + 20
    $Global:chkChgOwner = New-Object Windows.Forms.RadioButton
        $Global:chkChgOwner.Left = 100; $Global:chkChgOwner.Width = 450; $Global:chkChgOwner.Top = $TopLoc  
        $Global:chkChgOwner.Text = "Change Shared Mailbox Ownership" 
        $Global:chkChgOwner.Checked = $false   # set a default value 
        $Global:chkChgOwner.TabIndex = 2
        $Global:form.Controls.Add($Global:chkChgOwner) 
        # Obtain Value with: $Global:chkChgOwner.Checked

    ## Add Access Group
    $TopLoc = $TopLoc + 20
    $Global:chkAccessGrp = New-Object Windows.Forms.RadioButton
        $Global:chkAccessGrp.Left = 100; $Global:chkAccessGrp.Width = 450; $Global:chkAccessGrp.Top = $TopLoc  
        $Global:chkAccessGrp.Text = "Add Access Group (.ED, .AU or .RE)" 
        $Global:chkAccessGrp.Checked = $false   # set a default value 
        $Global:chkAccessGrp.TabIndex = 3
        $Global:form.Controls.Add($Global:chkAccessGrp) 
        # Obtain Value with: $Global:chkAccessGrp.Checked

    ## Rename Shared Mailbox
    $TopLoc = $TopLoc + 20
    $Global:chkRenShrMbx = New-Object Windows.Forms.RadioButton
        $Global:chkRenShrMbx.Left = 100; $Global:chkRenShrMbx.Width = 450; $Global:chkRenShrMbx.Top = $TopLoc  
        $Global:chkRenShrMbx.Text = "Rename Shared Mailbox" 
        $Global:chkRenShrMbx.Checked = $false   # set a default value 
        $Global:chkRenShrMbx.TabIndex = 4 
        $Global:form.Controls.Add($Global:chkRenShrMbx) 
        # Obtain Value with: $Global:chkRenShrMbx.Checked

    ## Remove Shared Mailbox
    $TopLoc = $TopLoc + 20
    $Global:chkRemShrMbx = New-Object Windows.Forms.RadioButton
        $Global:chkRemShrMbx.Left = 100; $Global:chkRemShrMbx.Width = 450; $Global:chkRemShrMbx.Top = $TopLoc  
        $Global:chkRemShrMbx.Text = "Remove Shared Mailbox"
        $Global:chkRemShrMbx.TabIndex = 5
        $Global:form.Controls.Add($Global:chkRemShrMbx) 
        # Obtain Value with: $Global:chkRemShrMbx.Checked

    ## Add/Remove Alias
    $TopLoc = $TopLoc + 20
    $Global:chkAlias = New-Object Windows.Forms.RadioButton
        $Global:chkAlias.Left = 100; $Global:chkAlias.Width = 450; $Global:chkAlias.Top = $TopLoc  
        $Global:chkAlias.Text = "Add/Remove Email Alias from Shared Mailbox" 
        $Global:chkAlias.Checked = $Global:chkAlias.Checked   # set a default value 
        $Global:chkAlias.TabIndex = 6
        $Global:form.Controls.Add($Global:chkAlias) 
        # Obtain Value with: $Global:chkAlias.Checked

    ## View mailbox statistics
    $TopLoc = $TopLoc + 20
    $Global:chkStats = New-Object Windows.Forms.RadioButton
        $Global:chkStats.Left = 100; $Global:chkStats.Width = 450; $Global:chkStats.Top = $TopLoc  
        $Global:chkStats.Text = "Grant/Deny External Addresses Access to Distribution List or Security Group" 
        $Global:chkStats.Checked = $false   # set a default value 
        $Global:chkStats.TabIndex = 7
        $Global:form.Controls.Add($Global:chkStats) 
        # Obtain Value with: $Global:chkStats.Checked
}
<#
Function Publish-ShrForm
{
    ## Finalize Form and Show Dialog 
    $Global:form.Add_Shown( { $form.Activate(); $Global:okButton.Focus() } )  #Activate and Set Focus 
    $Global:ShrResult = $Global:form.ShowDialog()          ## Show the form, and wait for the response
}
#>

$MbxChg = "1"

Do
{
    Build-ShrMbxMenuForm
    Add-FormStandardButtons
    $Global:InputFocus = $Global:okButton
    Publish-Form

    If ($Global:chkNewShrMbx.checked -eq "Checked")
    {
        invoke-expression -Command .\ShrMbxNew.ps1
        write-host "Shared Mailbox Creation Complete" -ForegroundColor Red
    }

    If ($Global:chkChgOwner.Checked -eq "Checked")
    {
		invoke-expression -Command .\ShrMbxChgOwner.ps1
    }
   
    If ($Global:chkAccessGrp.checked -eq "Checked")                 
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
                
    If ($Global:chkRenShrMbx.Checked -eq "Checked")
    {
        invoke-expression -Command .\RenameShrMbx.ps1
        invoke-Expression -Command e:\O365AdminShared\EMailTemplates\SharedMailboxRenamed.oft
        write-host "Shared Mailbox Rename Complete" -ForegroundColor Red
    }

    If ($Global:chkRemShrMbx.Checked -eq "Checked")
    {
        If (Test-Path "c:\temp\RemoveShrMbx.csv")
        {
            invoke-expression -Command .\ShrMbxRemove-wInputFile.ps1
        }
        else
        {
            invoke-expression -Command .\ShrMbxRemove.ps1
        }
        write-host "Shared Mailbox Removal Complete" -ForegroundColor Red
    }

    If ($Global:chkAlias.Checked -eq "Checked")
    {

    ## Need to modify to check to see if the account is a Service Account if it is it must be modified using AD Tools
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

    If ($Global:chkStats.Checked -eq "Checked")
    {
       invoke-expression -Command .\MailboxStatistics.ps1
    }
}While ($Global:Result -eq "OK")	