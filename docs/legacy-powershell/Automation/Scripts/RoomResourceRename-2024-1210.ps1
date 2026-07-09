<####
#### Room and Resource Admin Menu
####
#
# Called by:  RoomResourceAdminMenu.ps1
#
#   Modified 02/06/2018 - Added necessary code to properly abort the Room Rename#
#   Modified 10/03/2019 - Added pause points to make sure the files needed to continue are ready
#   Modified 02/20/2023 - Modified regional manager setting (mbx.XX.rrs.admin) and to use the get-mailboxRegionalConfiguration from get-mailboxCalendarConfiguration
#   Modified 10/24/2023 - Added pause to allow the rename to sync to whereever in O365 as it is reporting errors of the new object not being found.
#>

function Build-DefaultForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Rename Conference Room or Resource/Equipment" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(600,200) #(W,H)

#    $Global:okButton.Text = "Update"
#    $Global:OKDetails = "Update"
    Add-FormStandardButtons
    
    $Top = 10 
    $InpLeft = 110
    $InpRight = 370

    ## Ticket Number
    $Script:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Script:lblTaskNo.Text = "Ticket Number:"
        $Script:lblTaskNo.Top = $Top; $Script:lblTaskNo.Left = $InpRight-100; $Script:lblTaskNo.Width=150 ; $Script:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Script:lblTaskNo)    # Add to Form 
        # 
        $Script:txtTaskNo = New-Object Windows.Forms.TextBox  
        $Script:txtTaskNo.Top = $Top; $Script:txtTaskNo.Left = $InpRight; $Script:txtTaskNo.Width = 120;  
        $Script:TxtTaskNo.TabIndex = 0
        $Script:txtTaskNo.Text = "TASK"   # Enter ticket number 
        $Global:form.Controls.Add($Script:txtTaskNo)    # Add to Form

    ## Column Names
    $Script:lblTitles = New-Object System.Windows.Forms.Label   
        $Script:lblTitles.Text = "Original Room Details                                         New Room Details"
        $Script:lblTitles.Top = $Top+30 ; $Script:lblTitles.Left = 165; $Script:lblTitles.Width=500 ;$Script:lblTitles.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblTitles)    # Add to Form 

    $Top = $Top + 50
    ## Display Name
    $Script:lblRoomName = New-Object System.Windows.Forms.Label   
        $lblRoomName.Text = "Room Name:"
        $lblRoomName.Top = $Top ; $lblRoomName.Left = 10; $lblRoomName.Width=120 ;$lblRoomName.AutoSize = $true 
        $Global:form.Controls.Add($lblRoomName)    # Add to Form 
        # 
        $Script:txtRoomName = New-Object Windows.Forms.TextBox
        $Script:txtRoomName.Top = $Top; $Script:txtRoomName.Left = $InpLeft; $Script:txtRoomName.Width = 300;  
        $Script:txtRoomName.Text = ""  # DisplayName
        $Script:txtRoomName.TabStop = $False
        $Global:form.Controls.Add($Script:txtRoomName)    # Add to Form
        $Script:txtRoomName.Add_Click({
            Hide-Fields
            $Script:ButGetGrp.visible = $True
        })

        $Script:txtNewRoomName = New-Object Windows.Forms.TextBox
        $Script:txtNewRoomName.Top = $Top; $Script:txtNewRoomName.Left = $InpRight; $Script:txtNewRoomName.Width = 250;  
        $Script:txtNewRoomName.TabIndex = 2
        $Global:form.Controls.Add($Script:txtNewRoomName)    # Add to Form
        $Script:txtNewRoomName.Add_Click({
            Update-RoomDetails
        })

    #Get Group Details Button
    $Script:ButGetGrp = New-Object Windows.Forms.Button
    $Script:ButGetGrp.Location = New-object System.Drawing.Size(430,40)
    $Script:ButGetGrp.Size = new-Object System.Drawing.Size(110,20)
    $Script:ButGetGrp.Text = "Get Room Details"
    $Global:form.Controls.Add($Script:ButGetGrp)
    $Script:ButGetGrp.Add_Click({
        $Script:Room = $Script:txtRoomName.Text.Trim(" ")                
        $Exists = [bool]($Script:RoomDetails = get-Mailbox $Script:Room -ErrorAction SilentlyContinue)

        If ($Exists -eq $True)
        {
            $Script:ButGetGrp.visible = $false
            $Script:Room = $Script:RoomDetails.DisplayName
            $Script:txtRoomName.Text = $Script:Room
            $Global:InputFocus = $Script:txtTaskNo.Text
            write-host "Retreiving Group Details for: $Script:Room" -ForegroundColor Cyan
            Complete-Form
        }
        else
        {
            $Script:txtRoomName.Text = "Invalid"
            $Script:ButGetGrp.Visible = $True
            $Global:form.Size = New-Object System.Drawing.Size(600,200) #(W,H)
        }
    })

    $Top = $Top + 30
    #Room Job Title
    $Script:lblJobTitle = New-Object System.Windows.Forms.Label   
        $Script:lblJobTitle.Text = "Room Title:"
        $Script:lblJobTitle.Top = $Top; $Script:lblJobTitle.Left = 10; $Script:lblJobTitle.Width=150 ; $Script:lblJobTitle.AutoSize = $true
        $Global:form.Controls.Add($Script:lblJobTitle)    # Add to Form 
        # 
        $Script:txtJobTitle = New-Object Windows.Forms.TextBox  
        $Script:txtJobTitle.Top = $Top; $Script:txtJobTitle.Left = $InpLeft; $Script:txtJobTitle.Width = 250;
        $Script:txtJobTitle.ReadOnly = $True
        $Script:txtJobTitle.TabStop = $False
        $Script:txtJobTitle.Text = ""
        $Global:form.Controls.Add($Script:txtJobTitle)    # Add to Form

        $Script:txtNewJobTitle = New-Object Windows.Forms.TextBox  
        $Script:txtNewJobTitle.Top = $Top; $Script:txtNewJobTitle.Left = $InpRight; $Script:txtNewJobTitle.Width = 250;
        $Script:txtNewJobTitle.TabIndex = 3
        $Script:txtNewJobTitle.Text = ""
        $Global:form.Controls.Add($Script:txtNewJobTitle)    # Add to Form
        $Script:txtNewJobTitle.Add_Click({
            $Script:txtnewJobTitle.Text = ""
            Update-RoomDetails
        })

    $Top = $Top + 30
    #Room Building
    $Script:lblBldNo = New-Object System.Windows.Forms.Label   
        $Script:lblBldNo.Text = "Building Name/No:"
        $Script:lblBldNo.Top = $Top; $Script:lblBldNo.Left = 10; $Script:lblBldNo.Width=150 ; $Script:lblBldNo.AutoSize = $true
        $Global:form.Controls.Add($Script:lblBldNo)    # Add to Form 
        # 
        $Script:txtBldNo = New-Object Windows.Forms.TextBox  
        $Script:txtBldNo.Top = $Top; $Script:txtBldNo.Left = $InpLeft; $Script:txtBldNo.Width = 110;
        $Script:txtBldNo.ReadOnly = $True
        $Script:txtBldNo.TabStop = $False
        $Script:txtBldNo.Text = ""
        $Global:form.Controls.Add($Script:txtBldNo)    # Add to Form

        $Script:txtNewBldNo = New-Object Windows.Forms.TextBox  
        $Script:txtNewBldNo.Top = $Top; $Script:txtNewBldNo.Left = $InpRight; $Script:txtNewBldNo.Width = 110;
        $Script:txtNewBldNo.TabIndex = 3
        $Script:txtNewBldNo.Text = ""
        $Global:form.Controls.Add($Script:txtNewBldNo)    # Add to Form
        $Script:txtNewBldNo.Add_Click({
            Update-RoomDetails
        })

    $Top = $Top + 30 
    #Room Floor
    $Script:lblFloor = New-Object System.Windows.Forms.Label   
        $Script:lblFloor.Text = "Floor Number:"
        $Script:lblFloor.Top = $Top; $Script:lblFloor.Left = 10; $Script:lblFloor.Width=150 ; $Script:lblFloor.AutoSize = $true
        $Global:form.Controls.Add($Script:lblFloor)    # Add to Form
        # 
        $Script:txtFloor = New-Object Windows.Forms.TextBox  
        $Script:txtFloor.Top = $Top; $Script:txtFloor.Left = $InpLeft; $Script:txtFloor.Width = 110;
        $Script:txtFloor.ReadOnly = $True
        $Script:txtFloor.TabStop = $False 
        $Script:txtFloor.Text = ""
        $Global:form.Controls.Add($Script:txtFloor)    # Add to Form         
        
        $Script:txtNewFloor = New-Object Windows.Forms.TextBox  
        $Script:txtNewFloor.Top = $Top; $Script:txtNewFloor.Left = $InpRight; $Script:txtNewFloor.Width = 110;
        $Script:txtNewFloor.TabIndex = 4
        $Script:txtNewFloor.Text = ""
        $Global:form.Controls.Add($Script:txtNewFloor)    # Add to Form
        $Script:txtNewFloor.Add_Click({
            Update-RoomDetails
        })                  

    $Top = $Top + 30
    #Room Capacity
    $Script:lblCapacity = New-Object System.Windows.Forms.Label   
        $Script:lblCapacity.Text = "Room Capacity:"
        $Script:lblCapacity.Top = $Top; $Script:lblCapacity.Left = 10; $Script:lblCapacity.Width=150 ; $Script:lblCapacity.AutoSize = $true
        $Global:form.Controls.Add($Script:lblCapacity)    # Add to Form 
        # 
        $Script:txtCapacity = New-Object Windows.Forms.TextBox  
        $Script:txtCapacity.Top = $Top; $Script:txtCapacity.Left = $InpLeft; $Script:txtCapacity.Width = 110;  
        $Script:txtCapacity.ReadOnly = $True
        $Script:txtCapacity.TabStop = $False
        $Script:txtCapacity.Text = ""
        $Global:form.Controls.Add($Script:txtCapacity)    # Add to Form
        
        $Script:txtNewCapacity = New-Object Windows.Forms.TextBox  
        $Script:txtNewCapacity.Top = $Top; $Script:txtNewCapacity.Left = $InpRight; $Script:txtNewCapacity.Width = 110;
        $Script:txtNewCapacity.TabIndex = 5
        $Script:txtNewCapacity.Text = ""
        $Global:form.Controls.Add($Script:txtNewCapacity)    # Add to Form 
        $Script:txtNewCapacity.Add_Click({
            Update-RoomDetails
        })
 
    $Top = $Top + 30
    ## Do Not Fix Case
    $Script:chkAdjCase = New-Object Windows.Forms.checkbox 
        $Script:chkAdjCase.Left = $InpRight; $Script:chkAdjCase.Width = 250; $Script:chkAdjCase.Top = $Top
        $Script:chkAdjCase.Text = "Do Not Adjust Room Name Case" 
        $Script:chkAdjCase.TabIndex = 6
        $Script:chkAdjCase.Checked = $false   # set a default value
        $Global:form.Controls.Add($Script:chkAdjCase)

    ## EmailAddress
    $Script:lblRoomAddr = New-Object System.Windows.Forms.Label   
        $Script:lblRoomAddr.Text = "Email Address:"  
        $Script:lblRoomAddr.Top = $Top ; $Script:lblRoomAddr.Left = 10; $Script:lblRoomAddr.Width=150 ;$Script:lblRoomAddr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblRoomAddr)    # Add to Form 
        # 
        $Script:txtRoomAddr = New-Object Windows.Forms.TextBox  
        $Script:txtRoomAddr.Top = $Top; $Script:txtRoomAddr.Left = $InpLeft; $Script:txtRoomAddr.Width = 250;
        $Script:txtRoomAddr.ReadOnly = $true
        $Script:txtRoomAddr.TabStop = $False
        $Global:form.Controls.Add($Script:txtRoomAddr)    # Add to Form
        #
        $Script:txtNewRoomAddr = New-Object Windows.Forms.TextBox  
        $Script:txtNewRoomAddr.Top = $Top; $Script:txtNewRoomAddr.Left = $InpRight; $Script:txtNewRoomAddr.Width = 250;
        $Script:txtNewRoomAddr.ReadOnly = $True
        $Script:txtNewRoomAddr.TabStop = $False
        $Global:form.Controls.Add($Script:txtNewRoomAddr)    # Add to Form
        $Script:txtNewRoomAddr.Add_Click({
            $Script:txtNewRoomAddr.ReadOnly = $False
            $Script:txtNewAlias.ReadOnly = $False
        })

    $Top = $Top + 30 
    #Room Alias
    $Script:lblAlias = New-Object System.Windows.Forms.Label   
        $Script:lblAlias.Text = "Room Alias:"
        $Script:lblAlias.Top = $Top; $Script:lblAlias.Left = 10; $Script:lblAlias.Width=150 ; $Script:lblAlias.AutoSize = $true
        $Global:form.Controls.Add($Script:lblAlias)    # Add to Form 
        # 
        $Script:txtAlias = New-Object Windows.Forms.TextBox  
        $Script:txtAlias.Top = $Top; $Script:txtAlias.Left = $InpLeft; $Script:txtAlias.Width = 250;
        $Script:txtAlias.ReadOnly = $True
        $Script:txtAlias.TabStop = $False 
        $Script:txtAlias.Text = ""
        $Global:form.Controls.Add($Script:txtAlias)    # Add to Form
        # 
        $Script:txtNewAlias = New-Object Windows.Forms.TextBox  
        $Script:txtNewAlias.Top = $Top; $Script:txtNewAlias.Left = $InpRight; $Script:txtNewAlias.Width = 250;
        $Script:txtNewAlias.ReadOnly = $True
        $Script:txtNewAlias.TabStop = $False
        $Script:txtNewAlias.Text = ""
        $Global:form.Controls.Add($Script:txtNewAlias)    # Add to Form
        $Script:txtNewAlias.Add_Click({
            $Script:txtNewAlias.ReadOnly = $False
            $Script:txtNewRoomAddr.ReadOnly = $False
        })                        

    #Build New Alias,EmailAddress
    $Script:ButBldAddr = New-Object Windows.Forms.Button
        $Script:ButBldAddr.Location = New-object System.Drawing.Size(($InpRight+40),$Top)
        $Script:ButBldAddr.Size = new-Object System.Drawing.Size(120,20)
        $Script:ButBldAddr.Text = "Get Address Details"
        $Script:ButBldAddr.TabIndex = 7
        $Global:form.Controls.Add($Script:ButBldAddr)
        $Script:ButBldAddr.Add_Click({
            Rebuild-AliasAddr
            $Global:okButton.Text = "Rename"
            $Global:OkButton.Visible = $True
        })

    $Top = $Top + 30
    ## Room Delegates
    $Script:lblRoomDele = New-Object System.Windows.Forms.Label   
        $Script:lblRoomDele.Text = "Room Delegates:"  
        $Script:lblRoomDele.Top = $Top; $Script:lblRoomDele.Left = 10; $Script:lblRoomDele.Width=150 ;$Script:lblRoomDele.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblRoomDele)    # Add to Form 
        # 
        $Script:txtRoomDele = New-Object Windows.Forms.TextBox  
        $Script:txtRoomDele.Top = $Top; $Script:txtRoomDele.Left = $InpLeft; $Script:txtRoomDele.Width = 250;  
        $Script:txtRoomDele.Text = "Retrieving Details...."  # Legacy Address
        $Script:txtRoomDele.TabStop = $False
        $Script:txtRoomDele.ReadOnly = $true
        $Global:form.Controls.Add($Script:txtRoomDele)    # Add to Form

    $Top = $Top + 30
    ## Room Users (Restricted Rooms only)
    $Script:lblRoomUsers = New-Object System.Windows.Forms.Label   
        $Script:lblRoomUsers.Text = "Authorized Users:"  
        $Script:lblRoomUsers.Top = $Top; $Script:lblRoomUsers.Left = 10; $Script:lblRoomUsers.Width=150 ;$Script:lblRoomUsers.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblRoomUsers)    # Add to Form 
        # 
        $Script:txtRoomUsers = New-Object Windows.Forms.TextBox  
        $Script:txtRoomUsers.Top = $Top; $Script:txtRoomUsers.Left = $InpLeft; $Script:txtRoomUsers.Width = 250;  
        $Script:txtRoomUsers.Text = "General Use Room"  # Legacy Address
        $Script:txtRoomusers.TabStop = $False
        $Script:txtRoomUsers.ReadOnly = $true
        $Global:form.Controls.Add($Script:txtRoomUsers)    # Add to Form

    $Top = $Top + 30
    ## Not for Display Box to Not Display the email template
    $Script:chkTemplate = New-Object Windows.Forms.checkbox 
        $Script:chkTemplate.Left = $InpLeft+30; $Script:chkTemplate.Width = 350; $Script:chkTemplate.Top = $Top
        $Script:chkTemplate.Text = "Display Email Template for Room/Equipment Booking Changes" 
        $Script:chkTemplate.Checked = $true   # set a default value
        $Global:form.Controls.Add($Script:chkTemplate)

    #Report Details
    $Top = $Top + 30
    $Script:lblRptFile = New-Object System.Windows.Forms.Label   
        $Script:lblRptFile.Text = "Report File:"  
        $Script:lblRptFile.Top = $Top ; $Script:lblRptFile.Left = 10; $Script:lblRptFile.Width=150 ;$Script:lblRptFile.AutoSize = $true 
        $form.Controls.Add($Script:lblRptFile)    # Add to Form 
        # 
        $Script:txtRptFile = New-Object Windows.Forms.TextBox
        $Script:txtRptFile.ReadOnly = $true; 
        $Script:txtRptFile.Top = $Top; $Script:txtRptFile.Left = 80; $Script:txtRptFile.Width = 525; 
        $Script:txtRptFile.Text = $ReportFile
        $Script:txtRptFile.TabStop = $False
        $Global:form.Controls.Add($Script:txtRptFile)    # Add to Form

    Hide-Fields
}

Function Hide-Fields
{
    $Global:form.Text = "Rename Conference Room or Resource/Equipment" 
    $Script:lblTaskNo.Visible = $False
    $Script:txtTaskNo.Visible = $False
    $Script:lblTitles.Visible = $False
    $Script:txtRoomName.Width = 300; $Script:lblRoomName.Top = 40; $Script:txtRoomName.Top = 40
    $Script:txtRoomName.ReadOnly = $False
    $Script:txtNewRoomName.Visible = $False
    $Script:lblRoomAddr.Visible = $False
    $Script:txtRoomAddr.Visible = $False
    $Script:txtNewRoomAddr.Visible = $False
    $Script:lblRoomDele.Visible = $False
    $Script:txtRoomDele.Visible = $False
    $Script:lblRoomUsers.Visible = $False
    $Script:txtRoomUsers.Visible = $False
    $Script:lblRptFile.Visible = $False
    $Script:txtRptFile.Visible = $False
    $Script:chkTemplate.Visible = $False
    $Global:okButton.Visible = $False
    $Script:txtRoomDele.Text = "Retrieving Details...."
    $Script:lblAlias.Visible = $False
    $Script:txtAlias.Visible =  $False
    $Script:txtNewAlias.Visible =  $False
    $Script:lblBldNo.Visible = $False
    $Script:txtBldNo.Visible =  $False
    $Script:lblJobTitle.Visible = $False
    $Script:txtJobTitle.Visible = $False
    $Script:txtNewJobTitle.Visible = $False
    $Script:lblFloor.Visible = $False
    $Script:txtFloor.Visible =  $False
    $Script:txtNewBldNo.Visible =  $False
    $Script:txtNewFloor.Visible =  $False
    $Script:lblCapacity.Visible = $False
    $Script:txtCapacity.Visible =  $False
    $Script:ButBldAddr.Visible = $False
    $Script:txtNewCapacity.Visible =  $False
    $Global:form.Size = New-Object System.Drawing.Size(600,200) #(W,H)
    $Script:txtNewRoomName.Text = ""
    $Script:txtBldNo.Text = ""
    $Script:txtFloor.Text = ""
    $Script:txtCapacity.Text = ""
    $Script:chkAdjCase.Checked = $False
    $Script:chkAdjCase.Visible = $False
}

Function Update-RoomDetails
{
    $Script:ButBldAddr.Visible = $True
    $Script:chkAdjCase.Visible = $True
    $Script:lblJobTitle.Visible = $True
    $Script:txtJobTitle.Visible = $True
    $Script:txtNewJobTitle.Visible = $True
    $Script:txtNewRoomAddr.Visible = $False
    $Script:txtNewAlias.Visible = $False
    $Global:OkButton.Visible = $False
}

Function Complete-Form
{
    $Script:txtRoomName.Width = 250;$Script:lblRoomName.Top = 60; $Script:txtRoomName.Top = 60
    $Script:lblTaskNo.Visible = $True
    $Script:txtTaskNo.Visible = $True
    $Script:lblTitles.Visible = $True
    $Script:txtRoomName.ReadOnly = $True
    $Script:txtNewRoomName.Visible = $True
    $Script:txtRoomName.Text = ($Script:txtRoomName.Text).Trim()
    $Script:txtRoom.Refresh
    $Global:form.Size = New-Object System.Drawing.Size(640,430) #(W,H)
    $Script:lblRoomAddr.Visible = $True
    $Script:txtRoomAddr.Visible = $True
    $Script:lblRoomDele.Visible = $True
    $Script:txtRoomDele.Visible = $True
    $Script:lblAlias.Visible = $True
    $Script:txtAlias.Visible = $True
    $Script:lblJobTitle.Visible = $True
    $Script:txtJobTitle.Visible = $True
    $Script:txtNewJobTitle.Visible = $True
    $Script:lblBldNo.Visible = $True
    $Script:txtBldNo.Visible = $True
    $Script:lblFloor.Visible = $True
    $Script:txtFloor.Visible = $True
    $Script:txtNewBldNo.Visible = $True
    $Script:txtNewFloor.Visible =  $True
    $Script:lblCapacity.Visible = $True
    $Script:txtCapacity.Visible = $True
    $Script:txtNewCapacity.Visible = $True
    $Script:chkTemplate.Visible = $True
    $Script:lblRptFile.Visible = $True
    $Script:txtRptFile.Visible = $True
    $Script:chkAdjCase.Visible = $True
    $Script:txtJobTitle.Text = (Get-MsolUser -UserPrincipalName $Script:RoomDetails.UserPrincipalName).Title
    $Script:txtNewJobTitle.Text = $Script:txtJobTitle.Text
    $Script:txtRoomAddr.Text = $Script:RoomDetails.PrimarySmtpAddress
    $Script:RoomDele = (Get-CalendarProcessing $Script:txtRoomName.Text).ResourceDelegates
    If ($Script:RoomDele.length -gt 0)
    {
        $Script:txtRoomDele.Text = ((Get-DistributionGroupMember -ResultSize Unlimited $Script:RoomDele[0]) | get-mailbox) -join ("; ")
    }
    else
    {
        $Script:txtRoomDele.Text = "**Not Configured - Plesase Review**"
    }
    $Script:RoomPolicy = Get-CalendarProcessing $Script:txtRoomName.Text
    If ($RoomPolicy.BookinPolicy[0].length -ne 0)
    {
        $Script:OldUsrGrp = (Get-DistributionGroup $Script:RoomPolicy.BookinPolicy[0]).DisplayName
        If ($Script:OldUsrGrp.length -gt 0)
        {
            $Script:txtRoomUsers.Text = $Script:OldUsrGrp
            $Script:lblRoomUsers.Visible = $True
            $script:txtRoomUsers.Visible = $True
        }
    }
    $Script:txtNewRoomName.Text = $Script:txtRoomName.Text
    $Script:RoomDetails = Get-Mailbox $Script:txtRoomName.Text
    $Global:form.Text = "Rename Conference Room"
    $Script:lblTitles.Text = "Original Room Details                                         New Room Details"
    If ($Script:RoomDetails.ResourceType -eq "Equipment")
    {
        $Global:form.Text = "Rename Equipment or Resouce"
        $Script:lblTitles.Text = "Original Equipment Details                                  New Equipment Details"      
        $Script:txtNewBldNo.Text = ""
        $Script:txtNewFloor.Text = ""
        $Script:txtNewCapacity.Text = ""
        $Script:txtNewBldNo.Visible = $False
        $Script:txtNewFloor.Visible = $False
        $Script:txtNewCapacity.Visible = $False
        $Script:chkAdjCase.Checked = $True
    }
    If (($Script:RoomDetails.Office -like "Building*") -or ($Script:RoomDetails.Office -like "Bldg*"))
    {
        If ($Script:RoomDetails.Office -like "Building*")
        {
            $Script:txtNewBldNo.Text = (($Script:RoomDetails.Office).Substring(0,(($Script:RoomDetails.Office).IndexOf("-")-1)))
            $Script:txtNewBldNo.Text = $Script:txtNewBldNo.Text -replace("Building","")
        }
        else
        {
            If ($Script:RoomDetails.Office -like "Bldg*")
            {
                $Script:txtNewBldNo.Text = (($Script:RoomDetails.Office).Substring(0,(($Script:RoomDetails.Office).IndexOf("-")-1)))
                $Script:txtNewBldNo.Text = $Script:txtBldNo.Text -replace("Bldg","")
            }
        }
        $Script:txtNewBldNo.Text = $Script:txtNewBldNo.Text.Trim()
        $Script:txtBldNo.Text = ($Script:txtNewBldNo.Text).trim()
    }
    Build-BldgFloor
    $Script:txtAlias.Text = $Script:RoomDetails.Alias
    $Script:txtCapacity.Text = $Script:RoomDetails.ResourceCapacity
    $Script:txtNewCapacity.Text = $Script:RoomDetails.ResourceCapacity
    $Script:ButBldAddr.Visible = $True
    $Script:ReportFile = $ReportDirectory + ($Script:Room -replace(" ","")) + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "Time" + ((get-date -uformat %T).Replace(":", "")) + ".Log"
    $Script:txtRptFile.Text = $Script:ReportFile
    $Script:txtTaskNo.Focus()
}

Function Build-BldgFloor
{
    If (($Script:RoomDetails.Office -like "*Floor*") -or ($Script:RoomDetails.Office -like "*Flr*"))
    {
        If ($Script:RoomDetails.Office -like "*Floor*")
        {
            If ($Script:RoomDetails.Office -like "*Ground Floor*")
            {
                $Script:txtNewFloor.Text = "Ground"
            }
            else
            {
                If ($Script:RoomDetails.Office -like "*-*")
                {
                    $hyph = (($Script:RoomDetails.Office).IndexOf("-")+1)
                    $Len = ($Script:RoomDetails.Office).Length
                    $Script:txtNewFloor.Text = ($Script:RoomDetails.Office).Substring(($hyph+1),($Len-($hyph+1)))
                    $Script:txtNewFloor.Text = $Script:txtNewFloor.Text -replace ("Floor ","")
                }
                else
                {
                    $Script:txtNewFloor.Text = $Script:RoomDetails.Office -replace ("Floor ","")
                }
            }
        }
        else
        {
            If ($Script:RoomDetails.Office -like "*Flr*")
            {
                $hyph = (($Script:RoomDetails.Office).IndexOf("-")+1)
                $Len = ($Script:RoomDetails.Office).Length
                $Script:txtNewFloor.Text = ($Script:RoomDetails.Office).Substring(($hyph+1),($Len-($hyph+1)))
                $Script:txtNewFloor.Text = $Script:txtNewFloor.Text -replace("FLR ","")
            }
        } 
    }
    If ((($Script:txtNewFloor.Text).IndexOf("F")) -gt 1)
    {
        $flr = (($Script:RoomDetails.Office).IndexOf("F"))
        $Len = ($Script:RoomDetails.Office).Length
        $Script:txtNewFloor.Text = ($Script:txtNewFloor.Text).Substring($flr,$len-$flr)
        $Script:txtNewFloor.Text = $Script:txtNewFloor.Text -replace ("Floor ","")
    }
    $Script:txtFloor.Text = ($Script:txtNewFloor.Text).Trim()
}

Function Rebuild-AliasAddr
{
    #Adjust name to proper case and replace acronyms
    $Script:chkAdjCase.Visible = $False
    $Script:Name = $Script:txtNewRoomName.Text
    If ($Script:chkAdjCase.Checked -eq $False)
    {
        $Script:Name = (Get-Culture).textinfo.totitlecase($Script:txtNewRoomName.Text.ToLower())
    }
    $Space = $Script:Name.IndexOf(" ")
    $NameLen = $Script:Name.Length
    $Script:MbxLoc = $Script:Name.Substring(0,$Script:Name.IndexOf(" "))
    $Script:MbxName = $Script:Name.Substring($Space+1,($NameLen-($Space+1)))
    $Script:MbxName = $Script:MbxName.TrimStart()
    $Script:RoomName = $Script:MbxName -replace(" ","")
    If ($Script:chkAdjCase.Checked -eq $False)
    {
        $Script:RoomName = ((Get-Culture).textinfo.totitlecase($Script:MbxName.ToLower())) -replace(" ","")
    }
    $Script:RoomName = $Script:RoomName -replace ("Room","")
    $Script:RoomGrp = $Script:MbxLoc + "Conference Rooms"  #Is this needed?
    
    foreach ($Acro in $Acro)
    {
        $Script:MbxName = $Script:MbxName -Replace($Acro.Acronym,$Acro.Translation)
    }

    If ($Script:RoomDetails.ResourceType -eq "Room")
    {
        $Global:form.Text = "Rename Conference Room"
        $Script:lblTitles.Text = "Original Room Details                                         New Room Details"
        $Script:RoomAddr = $Script:MbxLoc.ToUpper()
        If ($Script:txtNewBldNo.Text.Length -ne 0)
        {
            $Script:txtNewBldNo.Text = ((Get-Culture).textinfo.totitlecase($Script:txtNewBldNo.Text.ToLower())) -replace(" ","")
            $BldNo = $Script:txtNewBldNo.Text -replace ("Building ","")
            $Script:RoomAddr = $Script:RoomAddr + "Bldg" + $BldNo + "."
        }
        else
        {
            $Script:RoomAddr = $Script:RoomAddr + "."
        }

        If ($Script:txtNewFloor.Text.Length -ne 0)
        {
            $FlrNo = $Script:txtNewFloor.Text -replace ("Floor ","")
            If ($Script:txtNewFloor.Text -eq "Ground")
            {
                $Script:RoomAddr = $Script:RoomAddr + "FLRGrnd"
            }
            else
            {
                If ($Script:txtNewFloor.Text -eq "Basement")
                {
                    $Script:RoomAddr = $Script:RoomAddr + "FLRBase"
                }
                else
                {
                    $Script:RoomAddr = $Script:RoomAddr + "FLR" + $FlrNo
                }
            }
        }
        $Script:RoomAddr = ($Script:RoomAddr + "." + $Script:RoomName + "." + $Script:txtNewCapacity.Text + "@ul.com").Replace("..",".")
    }
    else
    {
        $Global:form.Text = "Rename Equipment or Resouce"
        $Script:lblTitles.Text = "Original Equipment Details                                  New Equipment Details" 
        $Script:RoomAddr = "RES." + $Script:MbxLoc.ToUpper()
        $Script:RoomAddr = ($Script:RoomAddr + "." + $Script:RoomName + "@ul.com").Replace("..",".")
        $Script:txtNewBldNo.Text = "N/A"
        $Script:txtNewFloor.Text = "N/A"
        $Script:txtNewCapacity.Text = "0"
    }

    $Script:EnteredMbxName = $Script:txtNewRoomName.Text
    $Script:txtNewRoomName.Text = ($Script:MbxLoc).ToUpper() + " "  + $Script:MbxName
    $Script:RoomAddr = $Script:RoomAddr -replace (" Room","")
    $Script:txtNewRoomAddr.Text = $Script:RoomAddr -replace ("()","")
    $Script:txtNewAlias.Text = $Script:RoomAddr -replace ("@ul.com","")
    $Script:MbxNewLoc = $Script:txtNewRoomName.Text.Substring(0,3)
    $Script:MbxOldLoc = $Script:txtRoomName.Text.Substring(0,3)
#    If ($Script:MbxNewLoc -ne $Script:MbxOldLoc)
#    {
        Check-TimeZone
#    }
    $Script:ButBldAddr.Visible = $False
    $Script:txtNewRoomAddr.Visible = $True
    $Script:txtNewAlias.Visible = $True
}

Function Check-TimeZone
{
    $Script:TimeZone = ""
    $Script:RegMgr = ""
    $TimeZones = Import-Csv e:\O365AdminShared\Data\RoomTimeZones.csv
    foreach ($Zone in $TimeZones)
    {
        If ($Zone.Code -eq $Script:MbxLoc)
        {
            $Script:TimeZone = $Zone.Zone
            $Script:RegMgr = $Zone.Admins
        }
    }
}

########################################

Check-Reconnect
$FileName		= "RoomResourceRename"
$Year = (get-date).ToString("yyyy")
$Path = "E:\Automation\RenameRoomResource\Report\" + $Year
If (Test-Path $path) {} else {New-Item -Path $Path -ItemType Directory}
$ReportDirectory	= $Path + "\Report-RenameRoomResource-"
$Acro = Import-Csv e:\O365AdminShared\Data\KnownAcronyms.csv

$Global:form = ""

Build-DefaultForm
$Global:InputFocus = $Script:txtRoomName
Publish-Form

If ($Global:Result -eq "OK")
{
    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
    WriteReportEvent
    $LineToWrite = "Room Settings Prior to Changes"
    WriteReportEvent
    $LineToWrite = "   Ticket Number              :" + "`t" + $Script:txtTaskNo.Text
    WriteReportEvent
    $LineToWrite = "   Room Name                  :" + "`t" + $Script:txtRoomName.Text
    WriteReportEvent
    $LineToWrite = "   Office Location            :" + "`t" + $Script:RoomDetails.Office
    WriteReportEvent
    $LineToWrite = "   Building Name/No           :" + "`t" + $Script:txtBldNo.Text
    WriteReportEvent
    $LineToWrite = "   Floor Number               :" + "`t" + $Script:txtFloor.Text
    WriteReportEvent
    $LineToWrite = "   Room Capacity              :" + "`t" + $Script:txtCapacity.Text
    WriteReportEvent
    $LineToWrite = "   Primary Email Address      :" + "`t" + $Script:txtRoomAddr.Text
    WriteReportEvent
    $LineToWrite = "   Additional Email Addresses :" + "`t" + ($Script:RoomDetails.EmailAddresses -replace ("smtp:","")) -join ", "
    WriteReportEvent
    $LineToWrite = "   Room Alias                 :" + "`t" + $Script:txtAlias.Text
    WriteReportEvent
    $LineToWrite = "   Room Delegates Group       :" + "`t" + $Script:RoomDele -join ", "
    WriteReportEvent  
    $LineToWrite = "   Room Delegate Members      :" + "`t" + $Script:txtRoomDele.Text
    WriteReportEvent     
    $LineToWrite = "   Authorized Users           :" + "`t" + $Script:txtRoomUsers.Text
    WriteReportEvent
    If ($Script:txtRoomUsers.Text -like "MBX.*")
    {
        $LineToWrite = "   Authorized User Members    :" + "`t" + (Get-DistributionGroupMember $Script:txtRoomUsers.Text).DisplayName -join "; "
#        $LineToWrite = "   Authorized User Members    :" + "`t" + ((Get-DistributionGroupMember $Script:txtRoomUsers.Text) | get-mailbox) -join ("; ")
    }
    WriteReportEvent
    $LineToWrite = "   Calendar Time Zone         :" + "`t" + (Get-MailboxRegionalConfiguration $Script:txtRoomName.Text).TimeZone + "`n"
    WriteReportEvent               
    $LineToWrite = "Room Requsted to Changes"
    WriteReportEvent
    $LineToWrite = "   Ticket Number              :" + "`t" + $Script:txtTaskNo.Text
    WriteReportEvent
    $LineToWrite = "   Room Name Entered          :" + "`t" + $Script:EnteredMbxName
    WriteReportEvent
    $LineToWrite = "   Mbx Location               :" + "`t" + $Script:MbxLoc.ToUpper()
    WriteReportEvent
    $LineToWrite = "   Room Name                  :" + "`t" + $Script:txtNewRoomName.Text
    WriteReportEvent
    $LineToWrite = "   Building Name/No           :" + "`t" + $Script:txtNewBldNo.Text
    WriteReportEvent
    $LineToWrite = "   Floor Number               :" + "`t" + $Script:txtNewFloor.Text
    WriteReportEvent
    $LineToWrite = "   Room Capacity              :" + "`t" + $Script:txtNewCapacity.Text
    WriteReportEvent
    $LineToWrite = "   Primary Email Address      :" + "`t" + $Script:txtNewRoomAddr.Text
    WriteReportEvent
    $LineToWrite = "   Room Alias                 :" + "`t" + $Script:txtNewAlias.Text
    WriteReportEvent
    $LineToWrite = "   Calendar TimeZone          :" + "`t" + $Script:TimeZone
    WriteReportEvent
    $LineToWrite = "   Room Regional Managers     :" + "`t" + $Script:RegMgr + "`n"
    WriteReportEvent
    $LineToWrite = "Processing Changes"
    WriteReportEvent       
    
    $AddrExists = [bool](get-mailbox $Script:txtNewRoomAddr.Text -ErrorAction SilentlyContinue)
    $AliasExists = [bool](get-mailbox $Script:txtNewAlias.Text -ErrorAction SilentlyContinue)

    If (($AddrExists -eq $True) -or ($AliasExists -eq $True))
    {
        #Check to see if the Title (Job Title in Azure) has changed
        If ($Script:txtJobTitle.Text -ne $Script:txtNewJobTitle.Text)
        {
            Set-MsolUser -UserPrincipalName $Script:RoomDetails.UserPrincipalName -Title $Script:txtNewJobTitle.Text
            Write-Host "Updating the Title of the Room/Equipment: " $Script:txtNewJobTitle.Text
            $LineToWrite = $RecordEvent + "REMO" + "`t" + "Current Title for Room/Resuorce            : "+ $Script:txtJobTitle.Text
            WriteReportEvent
            $LineToWrite = $RecordEvent + "REMO" + "`t" + "New Title to for Room/Resuorce            : "+ $Script:txtNewJobTitle.Text
            WriteReportEvent
        }
        else
        {
            write-host "Cannot Perform Rename this change the Room/Equipment already exists" -ForegroundColor Red
        }
    }
    else
    {
        $NewLocation = ""
        If ($Script:txtNewBldNo.Text.length -gt 0)
        {
            $NewLocation = "Building " + $Script:txtNewBldNo.Text
        }
        If ($Script:txtNewFloor.Text.length -gt 0)
        {
            If ($NewLocation.Length -gt -0)
            {
                $NewLocation = $NewLocation + " - Floor " + $Script:txtNewFloor.Text
            }
            else
            {
                $NewLocation = "Floor " + $Script:txtNewFloor.Text
            }
        }

        If ($Script:RoomPolicy.BookInPolicy.Length -lt 1)
        {
            $OldRoomGrp = $Script:MbxOldLoc + " Conference Rooms"
            $NewRoomGrp = $Script:MbxNewLoc + " Conference Rooms"
            $DeleGrp = "MBX." + $Script:MbxNewLoc + ".RRS.OutOfPolicy.DE"
            $Restricted = "No"
        }
        else
        {
            $OldRoomGrp = $Script:MbxOldLoc + " Restricted Rooms"
            $NewRoomGrp = $Script:MbxNewLoc + " Restricted Rooms"
            $DeleGrp = ("MBX." + $Script:MbxNewLoc + ".RRS.OutOfPolicy." + ($Script:MbxName -replace ("Room","")) + ".DE") -replace (" ","")
            $Restricted = "Yes"
        }

        #check to see if the rooms is a member of the old room list and if the new room list exists
        $LstMbrExists = [bool](Get-DistributionGroupMember $OldRoomGrp -ErrorAction SilentlyContinue |Where-Object {$_.Name -eq $Script:txtRoomName.Text})
        $NewLstExists = [bool](Get-DistributionGroup $NewRoomGrp -ResultSize Unlimited -ErrorAction SilentlyContinue)
                
        $AddrList = $Script:RoomDetails.EmailAddresses += "SMTP:" + $Script:txtNewRoomAddr.Text
        Set-Mailbox $Script:txtRoomName.Text -Name $Script:txtNewRoomName.Text -Alias $Script:txtNewAlias.Text -DisplayName $Script:txtNewRoomName.Text -ResourceCapacity $Script:txtNewCapacity.Text.Trim() -Office $NewLocation -EmailAddresses $AddrList
        Write-Host "Room/Resource " $Script:RoomDetails.Name " renamed to " $Script:txtNewRoomName.Text
        $LineToWrite = $RecordEvent + "RENA" + "`t" + "Renamed Room/Resuorce                 : "+ $Script:RoomDetails.Name + " to " + $Script:txtNewRoomName.Text
        WriteReportEvent

        write-host "Pausing script to allow the rename to complete in O365 environment" -ForegroundColor Yellow
        Start-Sleep -Seconds 15

        #Check to see if this is in a new TimeZone
        $CurrZone = Get-MailboxRegionalConfiguration $Script:txtNewRoomName.Text
        If ($CurrZone.TimeZone -ne $Script:TimeZone)
        {
            Set-MailboxRegionalConfiguration $Script:txtNewRoomName.Text -TimeZone $Script:TimeZone
            $LineToWrite = $RecordEvent + "RENA" + "`t" + "Updated Room Time Zone Setting        : "+ $CurrZone.TimeZone + " to " + $Script:TimeZone
            WriteReportEvent
        }

        #This will move the room to a general use or restricted room list if the old and new locations don't match
        If (($Script:MbxOldLoc -ne $Script:MbxNewLoc) -and ($Script:RoomDetails.ResourceType -eq "Room"))
        {
            #Remove from old room list and add to new
            Write-Host "Removing from old Room Group: " $OldRoomGrp
            Remove-DistributionGroupMember $OldRoomGrp -Member $Script:txtNewRoomName.Text -Confirm:$False
            $LineToWrite = $RecordEvent + "REMO" + "`t" + "Removed from Room/Resuorce            : "+ $OldRoomGrp
            WriteReportEvent

            If ($NewLstExists -eq $False)
            {
                New-DistributionGroup -Name $NewRoomGrp -RoomList -ManagedBy "MBX.RRS.Owner"
                Write-host "The new room group created: " $NewRoomGrp
                $LineToWrite = $RecordEvent + "CREA" + "`t" + "Created new Room List                 : "+ $NewRoomGrp
                WriteReportEvent
            }

            Add-DistributionGroupMember $NewRoomGrp -Member $Script:txtNewRoomName.Text
            Write-Host "Adding to the New Room Group: " $NewRoomGrp
            $LineToWrite = $RecordEvent + "ADD" + "`t" + "Added Room to Room List               : "+ $NewRoomGrp
            WriteReportEvent

            #Change the to the new room delegates group if this is a general use room
        }
        else
        {
            $LineToWrite = $RecordEvent + "NOCH" + "`t" + "No Room List Changes Needed this item is not a Room"
            WriteReportEvent
        }

        # Rename the delegates and users groups for restricted rooms
        If ($Restricted -eq "Yes")
        {
            $LineToWrite = $RecordEvent + "INFO" + "`t" + "This is a Restricted Room/Resource"
            WriteReportEvent
            $Script:OldUsrGrp = (Get-DistributionGroup $Script:RoomPolicy.BookinPolicy[0]).DisplayName
            $NewUsrGrp = ("MBX." + $Script:MbxNewLoc + ".RRS.OutOfPolicy." + ($Script:MbxName -replace ("Room","")) + ".US") -replace (" ","")

            If ($Script:MbxOldLoc -ne $Script:MbxNewLoc)
            {
                #Rename the restricted room delegate group name
                $chk = $Script:RoomPolicy.ResourceDelegates[0]
                $Exists = [bool](Get-DistributionGroup $chk -ResultSize Unlimited -ErrorAction SilentlyContinue)
                If ($Exists -eq $True)
                {
                    write-host "Renamed Restricted Delegate Group" $Script:RoomPolicy.ResourceDelegates "to" $DeleGrp
                    Set-DistributionGroup $chk -Name $DeleGrp -DisplayName $DeleGrp -EmailAddresses @{add=($DeleGrp + "@ul.com")}
                    $LineToWrite = $RecordEvent + "RENA" + "`t" + "Renamed Delegated Group"
                    WriteReportEvent
                    $LineToWrite = $RecordEvent + "RENA" + "`t" + "    Old Name:        : " + $chk
                    WriteReportEvent
                    $LineToWrite = $RecordEvent + "RENA" + "`t" + "    New Name:        : " + $DeleGrp
                    WriteReportEvent
                }
                else
                {
                    write-host "The restricted delegate group does not exist: " $Script:RoomPolicy.ResourceDelegates
                    $LineToWrite = $RecordEvent + "FAIL" + "`t" + "Restricted Delegate Group Not Found   : " + $chk
                    WriteReportEvent
                }
            }
            else
            {
                $LineToWrite = $RecordEvent + "NOCH" + "`t" + "No Site Code Changes made"
                WriteReportEvent
            }

            If ($Script:OldUsrGrp -ne $NewUsrGrp)
            {
                $Exists = [bool](Get-DistributionGroup $NewUsrGrp -ErrorAction SilentlyContinue)
                If ($Exists -eq $False)
                {
                    write-host "Renamed User Group" $Script:OldUsrGrp "to" $NewUsrGrp
                    Set-DistributionGroup $Script:OldUsrGrp -Name $NewUsrGrp -DisplayName $NewUsrGrp -EmailAddresses @{add=($NewUsrGrp + "@ul.com")}
                    $LineToWrite = $RecordEvent + "RENA" + "`t" + "Renamed Authorized User Group"
                    WriteReportEvent
                    $LineToWrite = $RecordEvent + "RENA" + "`t" + "    Old Name:        : " + $Script:OldUsrGrp
                    WriteReportEvent
                    $LineToWrite = $RecordEvent + "RENA" + "`t" + "    New Name:        : " + $NewUsrGrp
                    WriteReportEvent
                }
                else
                {
                    write-host "The current Restricted User user group does not exist: " $Script:OldUsrGrp
                    $LineToWrite = $RecordEvent + "FAIL" + "`t" + "Restricted User Group Not Found       : " + $Script:OldUsrGrp
                    WriteReportEvent
                }
            }
            else
            {
                $LineToWrite = $RecordEvent + "FAIL" + "`t" + "Authorized User Group Did Not Change      : " + $Script:OldUsrGrp
                WriteReportEvent
            }
        }
        else
        {
            #This section with update the room permissions on general use rooms only as the groups for restricted rooms are simply renamed
            $Exists = [bool](Get-MailboxPermission $Script:txtNewRoomName.Text |Where-Object {$_.User -eq $Script:RoomDele})
            If ($Script:MbxOldLoc -ne $Script:MbxNewLoc)
            {
                $Exists = [bool](Get-MailboxPermission $Script:txtNewRoomName.Text |Where-Object {$_.User -eq $Script:RoomDele})
                If ($Exists -eq $True)
                {
                    Remove-MailboxPermission $Script:txtNewRoomName.Text -User $Script:RoomDele -AccessRights FullAccess -Confirm:$FALSE
                    $LineToWrite = $RecordEvent + "REMO" + "`t" + "Removed Permission from old Group    : "+ $Script:RoomDele
                    WriteReportEvent
                    $NewDele = $Script:RoomDele -replace ($Script:MbxOldLoc,$Script:MbxNewLoc)
                    Add-MailboxPermission $Script:txtNewRoomName.Text -User $NewDele -AccessRights FullAccess -Confirm:$FALSE
                    $LineToWrite = $RecordEvent + "ADD " + "`t" + "Added Permission to new Group        : "+ $NewDele
                    WriteReportEvent
                }
            }
        }

        # Change the regional rrs admins for this mailbox
#        $NewRegAdm = "MBX." + $Script:RegMgr + "RRS.Admins"
        $Exists = [bool](Get-MailboxPermission $Script:txtNewRoomName.Text |Where-Object {$_.User -eq $Script:RegMgr})
        If ($Exists -eq $False)
        {
            $RegAdmin = (Get-MailboxPermission $Script:txtNewRoomName.Text |Where-Object {$_.User -like "*RRS.Admins"}).User
            foreach ($Adm in $RegAdmin)
            {
                If (($Adm -like "MBX.*") -and ($Adm -like "*RRS.Admins") -and ($Adm -ne $Script:RegMgr))
                {
                    Remove-MailboxPermission $Script:txtNewRoomName.Text -User $Adm -AccessRights FullAccess -Confirm:$FALSE
                    $LineToWrite = $RecordEvent + "REMO" + "`t" + "Removed Old Regional Permission Group: "+ $Adm
                    WriteReportEvent
                }
                Add-MailboxPermission $Script:txtNewRoomName.Text -User $Script:RegMgr -AccessRights FullAccess -Confirm:$FALSE
                $LineToWrite = $RecordEvent + "ADD " + "`t" + "Added New Regional Permission Group  : "+ $Script:RegMgr
                WriteReportEvent
            }
        }
    }
    Write-Host "Room/Resource Rename Changes Complete"
    $LineToWrite = "Room/Resource Rename Process Complete"
    WriteReportEvent

    If ($Script:chkTemplate.Checked -eq $True)
    {
        Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\RoomorResourceRename.oft
    }
}
else
{
    Write-Host "Room/Resource Rename Changes Cancelled" -ForegroundColor Red
}