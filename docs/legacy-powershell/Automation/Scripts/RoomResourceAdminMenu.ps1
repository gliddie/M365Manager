<####
#### Room and Resource Admin Menu
####
#
# Called by:  RoomResourceAdminMenu.ps1
#
#   Modified 02/06/2018 - Added necessary code to properly abort the Room Rename#
#   Modified 10/03/2019 - Added pause points to make sure the files needed to continue are ready
#   Modified 09/26/2022 - Modified the Rename room process to use a GUI interface.
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
        }

        If ($Script:chkRenameChgRR.Checked -eq "Checked")
        {
            invoke-expression -Command .\RoomResourceRename.ps1
        }

        If ($Global:chkRemoveRR.Checked -eq "Checked")
        {
            invoke-expression -Command .\RoomResourceRemove.ps1
#            invoke-expression -Command .\RemoveRoomResource.ps1
#      	    write-host "Is This a Room that was part of a Restricted Room group (Y/N)? " -foreground Yellow -NoNewline#
#		    $ResRes = Read-Host
#		    If ($ResRes -eq "Y")
#		    {
#			    Invoke-expression .\RemoveDistributionList.ps1
#		    }
        }
    $Global:Result ="OK"
    }
}While ($Global:Result -eq "OK")	