<####
#### Room and Resource Admin Menu
####
#
# Called by:  RoomResourceAdminMenu.ps1
#
#   Modified 02/06/2018 - Added necessary code to properly abort the Room Rename#
#   Modified 10/03/2019 - Added pause points to make sure the files needed to continue are ready
#>

function Build-RRAdminMenuForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Room and Resource Admin Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 540 ; $form.Height = 250  # Make the form wider 
    Add-FormStandardButtons

    $TopLoc = 20 
    ## Title Line
    $Global:lblTitleLine = New-Object System.Windows.Forms.Label   
        $lblTitleLine.Text = "Select Option:"
        $lblTitleLine.Top = 15 ; $lblTitleLine.Left = 60; $lblTitleLine.Width=120 ;$lblTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine)    # Add to Form 

    ## New General Use or Restricted Room
    $TopLoc = $TopLoc + 20
    $Global:chkNewGRRoom = New-Object Windows.Forms.RadioButton 
        $Global:chkNewGRRoom.Left = 100; $Global:chkNewGRRoom.Width = 450; $Global:chkNewGRRoom.Top = $TopLoc  
        $Global:chkNewGRRoom.Text = "Create a New General Use or Restricted Room" 
        $Global:chkNewGRRoom.Checked = $false   # set a default value 
        $Global:chkNewGRRoom.TabIndex = 1
        $Global:form.Controls.Add($Global:chkNewGRRoom) 
        # Obtain Value with: $Global:chkNewGRRoom.Checked
        $Global:InputFocus = $Global:chkNewGRRoom

    ## New General Use or Restricted Equipment
    $TopLoc = $TopLoc + 20
    $Global:chkNewGREquip = New-Object Windows.Forms.RadioButton 
        $Global:chkNewGREquip.Left = 100; $Global:chkNewGREquip.Width = 450; $Global:chkNewGREquip.Top = $TopLoc  
        $Global:chkNewGREquip.Text = "Create a New General Use or Restricted Resource" 
        $Global:chkNewGREquip.Checked = $false   # set a default value 
        $Global:chkNewGREquip.TabIndex = 2
        $Global:form.Controls.Add($Global:chkNewGREquip) 
        # Obtain Value with: $Global:chkNewGREquip.Checked


    ## Review OutofPolicy Management Groups
    $TopLoc = $TopLoc + 20
    $Script:chkOOPReview = New-Object Windows.Forms.RadioButton 
        $Script:chkOOPReview.Left = 100; $Script:chkOOPReview.Width = 450; $Script:chkOOPReview.Top = $TopLoc  
        $Script:chkOOPReview.Text = "Change/Review Out of Policy Configuration for All Room Types" 
        $Script:chkOOPReview.Checked = $false   # set a default value 
        $Script:chkOOPReview.TabIndex = 5
        $Global:form.Controls.Add($Script:chkOOPReview) 
        # Obtain Value with: $Script:chkOOPReview.Checked

 <#
    ## Grant/Remove Access to Restricted Room
    $TopLoc = $TopLoc + 20
    $Global:chkChgAccess = New-Object Windows.Forms.RadioButton 
        $Global:chkChgAccess.Left = 100; $Global:chkChgAccess.Width = 450; $Global:chkChgAccess.Top = $TopLoc  
        $Global:chkChgAccess.Text = "Grant/Remove Access to Restricted Room" 
        $Global:chkChgAccess.Checked = $false   # set a default value 
        $Global:chkChgAccess.TabIndex = 5
        $Global:form.Controls.Add($Global:chkChgAccess) 
        # Obtain Value with: $Global:chkChgAccess.Checked
#>
    ## Rename Room and/or Change Details
    $TopLoc = $TopLoc + 20
    $Script:chkRenameChgRR = New-Object Windows.Forms.RadioButton 
        $Script:chkRenameChgRR.Left = 100; $Script:chkRenameChgRR.Width = 450; $Script:chkRenameChgRR.Top = $TopLoc  
        $Script:chkRenameChgRR.Text = "Rename Room and Change Room Details" 
        $Script:chkRenameChgRR.Checked = $Script:chkRenameChgRR.Checked   # set a default value 
        $Script:chkRenameChgRR.TabIndex = 6
        $Global:form.Controls.Add($Script:chkRenameChgRR)

    ## Remove Room/Resource
    $TopLoc = $TopLoc + 20
    $Global:chkRemoveRR = New-Object Windows.Forms.RadioButton 
        $Global:chkRemoveRR.Left = 100; $Global:chkRemoveRR.Width = 450; $Global:chkRemoveRR.Top = $TopLoc  
        $Global:chkRemoveRR.Text = "Remove Room or Resoure" 
        $Global:chkRemoveRR.Checked = $false   # set a default value 
        $Global:chkRemoveRR.TabIndex = 7
        $Global:form.Controls.Add($Global:chkRemoveRR) 
        # Obtain Value with: $Global:chkRemoveRR.Checked
}

Do
{
    
    Build-RRAdminMenuForm
    Publish-Form

    If ($Global:Result -eq "OK")
    {
        If (($Global:chkNewGRRoom.Checked -eq "Checked") -or ($Global:chkNewGREquip.Checked))
        {
            invoke-expression -Command .\RoomResourceNewForm.ps1
        }

        If ($Script:chkOOPReview.Checked -eq "Checked")
        {
            invoke-expression -Command .\RoomResourceOOPChanges.ps1
            $Global:Result = "OK"
        }

    <#    If ($Global:chkNewGREquip.Checked -eq "Checked")
        {
                write-host "Executing the Creation of New Equipment hit return when ready" -ForegroundColor Yellow -NoNewline
                $Cont = read-host
                $ExeOK = "Y"
                Do {
                    invoke-expression -Command .\NewEquipment.ps1
                    Write-Host "Do you need to run this script again due to errors (Y/N)?" -ForegroundColor Yellow -NoNewline
                    $ExeOK = Read-Host
                    write-host ""
                } while ($ExeOK -eq "Y")

                write-host "Executing the Configure New Equipment script" -ForegroundColor Yellow
                write-host "Please wait 1-2 minutes betwen creating and starting the configuraiton process to allow time for the equipment" -foreground Red
                write-host "mailbox to be created and to be ready for configuration.  Starting this too soon will result in errors" -ForegroundColor Red
                write-host "Hit return when ready" -ForegroundColor Yellow -NoNewline
                $Cont = read-host
                $ExeOK = "Y"
                Do {
                    invoke-expression -Command .\ConfigEquipment.ps1
                    Write-Host "Do you need to run this script again due to errors (Y/N)?" -ForegroundColor Yellow -NoNewline
                    $ExeOK = Read-Host
                    write-host ""
                } while ($ExeOK -eq "Y")

            }

        If ($Global:chkChgAccess.Checked -eq "Checked")
        {
            write-host "     Enter (1) to Grant Access to Book a Restricted Room"
            write-host "           (2) to Remove Access to Book a Restricted Room"
            write-host "           (3) to Add an Additional Room Delegate"
            write-host "           (4) to Remove a Room Delegate"
            write-host "     Enter Option or 0 Exist with No Changes? "-ForegroundColor Red -NoNewline
            $RoomOpt = Read-Host
            write-host ""

            write-host "Do you have approval for this change from one of the existing Room Delegates (Y/N)? " -ForegroundColor Yellow -NoNewline
            $RmApproval = Read-Host

            If ($RmApproval -eq "Y")
            {
               If ($RoomOpt -ne 0)
               {
                   write-host "Do you have approval from one of the existing Room Delegates (Y/N)? "
                   $RmApproval = Read-Host
                   Write-Host "Enter the name of the Room Delegate or Room User Group " -ForegroundColor Yellow -NoNewline
                   $Grp = Read-Host
                   Write-Host "Enter the Name or Employee Number of the individual to add or remove from the membership " -ForegroundColor Yellow -NoNewline
                   $Newmem = Read-Host
                }

                If (($RoomOpt -eq 1) -or ($RoomOpt -eq 3))
                {
                    Add-DistributionGroupMember $grp -Member $newmem -BypassSecurityGroupManagerCheck

                    Write-Host $Newmem "added to " $Grp
                    Write-Host ""
                }
                else
                {
                    Remove-DistributionGroupMember $grp -Member $newmem -BypassSecurityGroupManagerCheck

                    Write-Host $Newmem "removed from " $Grp
                    Write-Host ""
                }
            }
            else
            {
                write-host "Please obtain approval from one of the existing room delegates before making this change" -fore Red
                pause
            }
        }
    #>

        If ($Script:chkRenameChgRR.Checked -eq "Checked")
        {
            If ($whoami -like "*96151")
            {
                invoke-expression -Command .\RoomResourceRename.ps1
            }
            else
            {
                #change room Details
                write-host "Enter the name of the Room: " -ForegroundColor Yellow -NoNewline
                $ChgRoom = Read-Host
                $RoomDetails = Get-Mailbox $ChgRoom

                If ($RoomDetails.ResourceType -eq "Room")
                {
			        write-host "    Room Name: " $RoomDetails.Name
			        write-host "Room Capacity: " $RoomDetails.ResourceCapacity
			        write-host "Room Location: " $RoomDetails.Office
			        write-host "   Room Alias: " $RoomDetails.Alias
			        write-host "Email Address: " $RoomDetails.PrimarySmtpAddress
			        write-host
		        }
		        else
		        {
			        write-host " Resource Name: " $RoomDetails.Name
			        write-host "Resource Alias: " $RoomDetails.Alias
			        write-host " Email Address: " $RoomDetails.PrimarySmtpAddress				
		        }

                write-host "Enter the New Room Name <CR> for no change: " -ForegroundColor Yellow -NoNewline
                $NewName = Read-Host
				
		        If ($RoomDetails.ResourceType -eq "Room")
		        {
			        write-host "Enter the New Room Capacity <CR> for no change: "  -ForegroundColor Yellow -NoNewline
			        $NewCapacity = Read-Host
			        write-host "Enter the New Room Location <CR> for no change: " -ForegroundColor Yellow -NoNewline
			        $NewLocation = Read-Host
		        }

                If ($NewName -eq "")
                {
                    $NewName = $RoomDetails.Name
                    $NewDispName = $RoomDetails.DisplayName
                    $NewAlias = $RoomDetails.Alias
                    $NewSMTPAddr = $RoomDetails.PrimarySmtpAddress
                }
                else
                {
                     $NewAlias = $RoomDetails.Alias
                }

                If (($NewCapacity -ne "") -and ($RoomDetails.ResourceType -eq "Room"))
                {
                    $NewAlias = $RoomDetails.Alias.Replace($RoomDetails.ResourceCapacity,$NewCapacity)
                    $NewSMTPAlias = $NewAlias + "@ul.com"
                }
                else
                {
                    $NewCapacity = $RoomDetails.ResourceCapacity
                }
                
                If (($NewLocation -eq "") -and ($RoomDetails.ResourceType -eq "Room"))
                {
                    $NewLocation = $RoomDetails.Office
                }
                
                if ($NewName -ne "")
                {
                    write-host ""
                    write-host "Existing Room Alias: " $RoomDetails.Alias
                    write-host "New Room Alias:      " $RoomDetails.Alias.Replace($RoomDetails.ResourceCapacity,$NewCapacity)
                    write-host ""
                    write-host "Enter in the Room Alias <CR> for no change (this is the information to the left of the @ in the internet address, no spaces or special characters) " -ForegroundColor Yellow -NoNewline
                    $UpdAlias = Read-Host
                    If ($UpdAlias -eq "")
                    {
                        $NewAlias = $RoomDetails.Alias.Replace($RoomDetails.ResourceCapacity,$NewCapacity)
                        $NewSMTPAlias = $NewAlias + "@ul.com"
                    }
                    else
                    {
                        $NewAlias = $UpdAlias
                    }
                    $NewSMTPAlias = $NewAlias + "@ul.com"
                }
               
                write-host
                write-host "       New Room Name: " $NewName
                write-host "   New Room Capacity: " $NewCapacity
                If (($NewLocation -ne "*Building*") -or ($NewLocation -ne "*Bldg*"))
                {
                    $NewLocation = "Floor " + $NewLocation
                }
                write-host "   New Room Location: " $NewLocation
                write-host "      New Room Alias: " $NewAlias
                write-host "New Internet Address: " $NewSMTPAlias
                write-host
                
                $NewPrimaryAlias = $RoomDetails.EmailAddresses += "SMTP:" + $NewSMTPAlias
      	        $Cont = "N"
                write-host "Continue with Renaming this object (Y/N) ? "-ForegroundColor Yellow -NoNewline
                $Cont = Read-Host

   		        If ($Cont -eq "Y")
		        {
    		        If ($NewName -ne $RoomDetails.Name)
			        {
				        Set-Mailbox $RoomDetails.Name -Name $NewName -Alias $NewAlias -DisplayName $NewName -ResourceCapacity $NewCapacity -Office $NewLocation -EmailAddresses $NewPrimaryAlias
			        }
			        else
			        {
				        Set-Mailbox $NewName -Alias $NewAlias -ResourceCapacity $NewCapacity -EmailAddresses $NewPrimaryAlias -Office $NewLocation
			        }
			        Write-Host
			        Write-Host "Room/Resource " $RoomDetails.Name " renamed to " $NewName
			        Write-Host "If this is a restricted room please use the Rename DistributionGroup script to rename the groups used to configure the room delegates and who is allowed to book the room"
                }
		        else
	            {
			        Write-Host "Room not renamed - process has been cancelled"
		        }
		
		        pause
            }
        }

        If ($Global:chkRemoveRR.Checked -eq "Checked")
        {
            invoke-expression -Command .\RemoveRoomResource.ps1
      	    write-host "Is This a Room that was part of a Restricted Room group (Y/N)? " -foreground Yellow -NoNewline
		    $ResRes = Read-Host
		    If ($ResRes -eq "Y")
		    {
			    Invoke-expression .\RemoveDistributionList.ps1
		    }
        }
    $Global:Result ="OK"
    }
}While ($Global:Result -eq "OK")	