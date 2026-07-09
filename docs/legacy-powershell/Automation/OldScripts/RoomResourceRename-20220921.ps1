<####
#### Room and Resource Admin Menu
####
#
# Called by:  RoomResourceAdminMenu.ps1
#
#   Modified 02/06/2018 - Added necessary code to properly abort the Room Rename#
#   Modified 10/03/2019 - Added pause points to make sure the files needed to continue are ready
#>

function Build-DefaultForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Rename Conference Room" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(600,200) #(W,H)

    $Global:OKDetails = "Update"
    Add-FormStandardButtons
    
    $Top = 10 
    $InpLeft = 110
    $InpRight = 370

    ## Ticket Number
    $Script:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Script:lblTaskNo.Text = "Ticket Number:"
        $Script:lblTaskNo.Top = $Top; $Script:lblTaskNo.Left = 10; $Script:lblTaskNo.Width=150 ; $Script:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Script:lblTaskNo)    # Add to Form 
        # 
        $Script:txtTaskNo = New-Object Windows.Forms.TextBox  
        $Script:txtTaskNo.Top = $Top; $Script:txtTaskNo.Left = $InpLeft; $Script:txtTaskNo.Width = 120;  
        $Script:TxtTaskNo.TabIndex = 1
        $Script:txtTaskNo.Text = "TASK"   # Enter ticket number 
        $Global:form.Controls.Add($Script:txtTaskNo)    # Add to Form

    ## Column Names
    $Script:lblTitles = New-Object System.Windows.Forms.Label   
        $Script:lblTitles.Text = "Original Room Details                                       New Room Details"
        $Script:lblTitles.Top = $Top+30 ; $Script:lblTitles.Left = 170; $Script:lblTitles.Width=500 ;$Script:lblTitles.AutoSize = $true 
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
        $Script:txtRoomName.TabIndex=10
        $Global:form.Controls.Add($Script:txtRoomName)    # Add to Form
        $Global:InputFocus = $Script:txtRoomName
        $Script:txtRoomName.Add_Click({
            Hide-Fields
            $Script:ButGetGrp.visible = $True
        })

        $Script:txtNewRoomName = New-Object Windows.Forms.TextBox
        $Script:txtNewRoomName.Top = $Top; $Script:txtNewRoomName.Left = $InpRight; $Script:txtNewRoomName.Width = 250;  
        $Script:txtNewRoomName.TabIndex = 2
        $Global:form.Controls.Add($Script:txtNewRoomName)    # Add to Form
        $Script:txtNewRoomName.Add_Click({$Script:txtNewRoomName.Text = ""})

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
    #Room Location (Buidling/Floor)
    $Script:lblLocation = New-Object System.Windows.Forms.Label   
        $Script:lblLocation.Text = "Room Location:"
        $Script:lblLocation.Top = $Top; $Script:lblLocation.Left = 10; $Script:lblLocation.Width=150 ; $Script:lblLocation.AutoSize = $true
        $Global:form.Controls.Add($Script:lblLocation)    # Add to Form 
        # 
        $Script:txtLocation = New-Object Windows.Forms.TextBox  
        $Script:txtLocation.Top = $Top; $Script:txtLocation.Left = $InpLeft; $Script:txtLocation.Width = 250;
        $Script:txtLocation.ReadOnly = $True 
        $Script:txtLocation.Text = ""
        $Global:form.Controls.Add($Script:txtLocation)    # Add to Form

        $Script:txtNewBldNo = New-Object Windows.Forms.TextBox  
        $Script:txtNewBldNo.Top = $Top; $Script:txtNewBldNo.Left = $InpRight; $Script:txtNewBldNo.Width = 140;
        $Script:txtNewBldNo.TabIndex = 3
        $Script:txtNewBldNo.Text = "Building "
        $Global:form.Controls.Add($Script:txtNewBldNo)    # Add to Form
        $Script:txtNewBldNo.Add_Click({$Script:txtNewBldNo.Text = ""})
        
        $Script:txtNewFloor = New-Object Windows.Forms.TextBox  
        $Script:txtNewFloor.Top = $Top; $Script:txtNewFloor.Left = $InpRight+145; $Script:txtNewFloor.Width = 105;
        $Script:txtNewFloor.TabIndex = 4
        $Script:txtNewFloor.Text = "Floor "
        $Global:form.Controls.Add($Script:txtNewFloor)    # Add to Form
        $Script:txtNewFloor.Add_Click({$Script:txtNewFloor.Text = ""})                  

    $Top = $Top + 30
    #Rooom Capacity
    $Script:lblCapacity = New-Object System.Windows.Forms.Label   
        $Script:lblCapacity.Text = "Room Capacity:"
        $Script:lblCapacity.Top = $Top; $Script:lblCapacity.Left = 10; $Script:lblCapacity.Width=150 ; $Script:lblCapacity.AutoSize = $true
        $Global:form.Controls.Add($Script:lblCapacity)    # Add to Form 
        # 
        $Script:txtCapacity = New-Object Windows.Forms.TextBox  
        $Script:txtCapacity.Top = $Top; $Script:txtCapacity.Left = $InpLeft; $Script:txtCapacity.Width = 250;  
        $Script:txtCapacity.ReadOnly = $True
        $Script:txtCapacity.Text = ""
        $Global:form.Controls.Add($Script:txtCapacity)    # Add to Form
        
        $Script:txtNewCapacity = New-Object Windows.Forms.TextBox  
        $Script:txtNewCapacity.Top = $Top; $Script:txtNewCapacity.Left = $InpRight; $Script:txtNewCapacity.Width = 250;
        $Script:txtNewCapacity.TabIndex = 5
        $Script:txtNewCapacity.Text = ""
        $Global:form.Controls.Add($Script:txtNewCapacity)    # Add to Form 
        $Script:txtNewCapacity.Add_Click({$Script:txtNewCapacity.Text = ""})
 
    $Top = $Top + 30
    #Build New Alias,EmailAddress
    $Script:ButBldAddr = New-Object Windows.Forms.Button
    $Script:ButBldAddr.Location = New-object System.Drawing.Size(($InpRight+40),$Top)
    $Script:ButBldAddr.Size = new-Object System.Drawing.Size(110,20)
    $Script:ButBldAddr.Text = "Get Room Details"
    $Script:ButBldAddr.TabIndex = 6
    $Global:form.Controls.Add($Script:ButBldAddr)
    $Script:ButBldAddr.Add_Click({
        Rebuild-AliasAddr
    })

    ## EmailAddress
    $Script:lblRoomAddr = New-Object System.Windows.Forms.Label   
        $Script:lblRoomAddr.Text = "Email Address:"  
        $Script:lblRoomAddr.Top = $Top ; $Script:lblRoomAddr.Left = 10; $Script:lblRoomAddr.Width=150 ;$Script:lblRoomAddr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblRoomAddr)    # Add to Form 
        # 
        $Script:txtRoomAddr = New-Object Windows.Forms.TextBox  
        $Script:txtRoomAddr.Top = $Top; $Script:txtRoomAddr.Left = $InpLeft; $Script:txtRoomAddr.Width = 250;
        $Script:txtRoomAddr.ReadOnly = $true
        $Global:form.Controls.Add($Script:txtRoomAddr)    # Add to Form

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
        $Script:txtAlias.Text = ""
        $Global:form.Controls.Add($Script:txtAlias)    # Add to Form      

    $Top = $Top + 30
    ## Room Delegates
    $Script:lblRoomDele = New-Object System.Windows.Forms.Label   
        $lblRoomDele.Text = "Room Delegates:"  
        $lblRoomDele.Top = $Top; $lblRoomDele.Left = 10; $lblRoomDele.Width=150 ;$lblRoomDele.AutoSize = $true 
        $Global:form.Controls.Add($lblRoomDele)    # Add to Form 
        # 
        $Script:txtRoomDele = New-Object Windows.Forms.TextBox  
        $Script:txtRoomDele.Top = $Top; $Script:txtRoomDele.Left = $InpLeft; $Script:txtRoomDele.Width = 250;  
        $Script:txtRoomDele.Text = "Retrieving Details...."  # Legacy Address
        $Script:txtRoomDele.ReadOnly = $true
        $Global:form.Controls.Add($Script:txtRoomDele)    # Add to Form

    $Top = $Top + 30
    ## Not for Display Box to Not Display the email template
    $Script:chkTemplate = New-Object Windows.Forms.checkbox 
        $Script:chkTemplate.Left = $InpLeft; $Script:chkTemplate.Width = 350; $Script:chkTemplate.Top = $Top
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
        $Script:txtRptFile.TabIndex = 90 # set Tab Order
        $Script:txtRptFile.Top = $Top; $Script:txtRptFile.Left = 80; $Script:txtRptFile.Width = 525; 
        $Script:txtRptFile.Text = $ReportFile
        $Global:form.Controls.Add($Script:txtRptFile)    # Add to Form

    Hide-Fields
}

Function Hide-Fields
{
    $Script:lblTaskNo.Visible = $False
    $Script:txtTaskNo.Visible = $False
    $Script:lblTitles.Visible = $False
    $Script:txtRoomName.Width = 300; $Script:lblRoomName.Top = 40; $Script:txtRoomName.Top = 40
    $Script:txtRoomName.ReadOnly = $False
    $Script:txtNewRoomName.Visible = $False
    $Script:lblRoomAddr.Visible = $False
    $Script:txtRoomAddr.Visible = $False
    $Script:lblRoomDele.Visible = $False
    $Script:txtRoomDele.Visible = $False
    $Script:lblRptFile.Visible = $False
    $Script:txtRptFile.Visible = $False
    $Script:chkTemplate.Visible = $False
    $Global:okButton.Visible = $False
    $Script:txtRoomDele.Text = "Retrieving Details...."
    $Script:lblAlias.Visible = $False
    $Script:txtAlias.Visible =  $False
    $Script:lblLocation.Visible = $False
    $Script:txtLocation.Visible =  $False
    $Script:txtNewBldNo.Visible =  $False
    $Script:txtNewFloor.Visible =  $False
    $Script:lblCapacity.Visible = $False
    $Script:txtCapacity.Visible =  $False
    $Script:ButBldAddr.Visible = $False
    $Script:txtNewCapacity.Visible =  $False
    $Global:form.Size = New-Object System.Drawing.Size(600,200) #(W,H)
    $Script:txtNewRoomName.Text = ""
    $Script:txtNewBldNo.Text = "Building "
    $Script:txtNewFloor.Text = "Floor "
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
    $Script:lblLocation.Visible = $True
    $Script:txtLocation.Visible = $True
    $Script:txtNewBldNo.Visible = $True
    $Script:txtNewFloor.Visible =  $True
    $Script:lblCapacity.Visible = $True
    $Script:txtCapacity.Visible = $True
    $Script:txtNewCapacity.Visible = $True
    $Script:chkTemplate.Visible = $True
    $Script:lblRptFile.Visible = $True
    $Script:txtRptFile.Visible = $True
    $Script:txtRoomAddr.Text = $Script:RoomDetails.PrimarySmtpAddress
    $Script:RoomDele = (get-CalendarProcessing $Script:txtRoomName.Text).ResourceDelegates
    $Script:txtRoomDele.Text = (Get-DistributionGroupMember -ResultSize Unlimited $Script:RoomDele[0]) -join (", ")
    $Script:txtNewRoomName.Text = $Script:txtRoomName.Text
    $Script:txtLocation.Text = $Script:RoomDetails.Office
    $Script:RoomDetails = Get-Mailbox $Script:txtRoomName.Text
    $Global:form.Text = "Rename Conference Room" 
    If ($Script:RoomDetails.ResourceType -eq "Equipment")
    {
        $Global:form.Text = "Rename Equipment or Resouce" 
    }
    If (($Script:RoomDetails.Office -like "Building*") -or ($Script:RoomDetails.Office -like "Bldg*"))
    {
        If ($Script:RoomDetails.Office -like "Building*")
        {
            $Script:txtNewBldNo.Text = (($Script:RoomDetails.Office).Substring(0,(($Script:RoomDetails.Office).IndexOf("-")-1)))
        }
        else
        {
            If ($Script:RoomDetails.Office -like "Bldg*")
            {
                $Script:txtNewBldNo.Text = (($Script:RoomDetails.Office).Substring(0,(($Script:RoomDetails.Office).IndexOf("-")-1)))
                $Script:txtNewBldNo.Text = $Script:txtBldNo.Text -replace("Bldg","Building")
            }
        }
    }
    Build-NewFloor   
    $Script:txtAlias.Text = $Script:RoomDetails.Alias
    $Script:txtLocation.Text = $Script:RoomDetails.Office
    $Script:txtCapacity.Text = $Script:RoomDetails.ResourceCapacity
    $Script:txtNewCapacity.Text = $Script:RoomDetails.ResourceCapacity
    $Script:ButBldAddr.Visible = $True
    $Script:ReportFile = $ReportDirectory + ($Script:Room -replace(" ","")) + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "Time" + ((get-date -uformat %T).Replace(":", "")) + ".Log"
    $Script:txtRptFile.Text = $Script:ReportFile
    $Global:InputFocus = $Script:txtNewRoomName
    $Global:InputFocus.Refresh()
}

Function Build-NewFloor
{
    If (($Script:RoomDetails.Office -like "*Floor*") -or ($Script:RoomDetails.Office -like "*Flr*"))
    {
        If ($Script:RoomDetails.Office -like "*Floor*")
        {
            If ($Script:RoomDetails.Office -like "*Ground Floor*")
            {
                $Script:txtNewFloor.Text = "Ground Floor"
            }
            else
            {
                If ($Script:RoomDetails.Office -like "*-*")
                {
                    $hyph = (($Script:RoomDetails.Office).IndexOf("-")+1)
                    $Len = ($Script:RoomDetails.Office).Length
                    $Script:txtNewFloor.Text = ($Script:RoomDetails.Office).Substring(($hyph+1),($Len-($hyph+1)))
                }
                else
                {
                    $Script:txtNewFloor.Text = $Script:RoomDetails.Office
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
                $Script:txtNewFloor.Text = $Script:txtNewFloor.Text -replace("FLR","Floor")
            }
        }      
    }
    If ((($Script:txtNewFloor.Text).IndexOf("F")) -gt 1)
    {
        $flr = (($Script:RoomDetails.Office).IndexOf("F"))
        $Len = ($Script:RoomDetails.Office).Length
        $Script:txtNewFloor.Text = ($Script:txtNewFloor.Text).Substring($flr,$len-$flr)
    }
}

<#
Function Build-NewDetails
{
#make sure the building and floor are preceeded with this label info
#            If ($NewName -eq "")
#            {
#                $NewName = $Script:RoomDetails.Name
#                $NewDispName = $Script:RoomDetails.DisplayName
#                $NewAlias = $Script:RoomDetails.Alias
#                $NewSMTPAddr = $Script:RoomDetails.PrimarySmtpAddress
#            }
#            else
#            {
#                 $NewAlias = $Script:RoomDetails.Alias
#            }
    If ($Script:RoomDetails.ResourceType -eq "Room")
    {
         If (($Script:txtNewCapacity.Text -ne "") -and (($Script:txtNewCapacity.Text).Trim()) -ne "Floor")
         {
            $NewAlias = $Script:RoomDetails.Alias.Replace($Script:RoomDetails.ResourceCapacity,$NewCapacity)
            $NewSMTPAlias = $NewAlias + "@ul.com"
         }
    }

            If (($Script:txtNewCapacity.Text -ne "") -and ($Script:RoomDetails.ResourceType -eq "Room"))
            {
                $NewAlias = $Script:RoomDetails.Alias.Replace($Script:RoomDetails.ResourceCapacity,$NewCapacity)
                $NewSMTPAlias = $NewAlias + "@ul.com"
            }
            else
            {
                $NewCapacity = $Script:RoomDetails.ResourceCapacity
            }
                
            If (($NewLocation -eq "") -and ($Script:RoomDetails.ResourceType -eq "Room"))
            {
                $NewLocation = $Script:RoomDetails.Office
            }
                
            if ($NewName -ne "")
            {
                write-host ""
                write-host "Existing Room Alias: " $Script:RoomDetails.Alias
                write-host "New Room Alias:      " $Script:RoomDetails.Alias.Replace($Script:RoomDetails.ResourceCapacity,$NewCapacity)
                write-host ""
                write-host "Enter in the Room Alias <CR> for no change (this is the information to the left of the @ in the internet address, no spaces or special characters) " -ForegroundColor Yellow -NoNewline
                $UpdAlias = Read-Host
                If ($UpdAlias -eq "")
                {
                    $NewAlias = $Script:RoomDetails.Alias.Replace($Script:RoomDetails.ResourceCapacity,$NewCapacity)
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
                
            $NewPrimaryAlias = $Script:RoomDetails.EmailAddresses += "SMTP:" + $NewSMTPAlias
}
#>
Function Rebuild-AliasAddr
{
    #Adjust name to proper case ann replace acronyms
    $Script:Name = (Get-Culture).textinfo.totitlecase($Script:txtNewRoomName.Text.ToLower())
    $Space = $Script:Name.IndexOf(" ")
    $NameLen = $Script:Name.Length
    $Script:MbxLoc = $Script:Name.Substring(0,$Script:Name.IndexOf(" "))
    $MbxName = $Script:Name.Substring($Space+1,($NameLen-($Space+1)))
    $RoomName = ((Get-Culture).textinfo.totitlecase($MbxName.ToLower())) -replace(" ","")
    $Script:RoomGrp = $Script:MbxLoc + "Conference Rooms"  #Is this needed?
    
    foreach ($Acro in $Acro)
    {
        $MbxName = $MbxName -Replace($Acro.Acronym,$Acro.Translation)
    }

    If ($Script:RoomDetails.ResourceType -eq "Room")
    {
        $Script:RoomAddr = $Script:MbxLoc.ToUpper()
    }
    else
    {
        $Script:RoomAddr = "RES." + $Script:MbxLoc.ToUpper()
    }

    If ($Script:txtNewBldNo.Text.Length -ne 0)
    {
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
            $Script:RoomAddr = $Script:RoomAddr + "FLR" + $FlrNo
        }
    }
    
    If ($Script:RoomDetails.ResourceType -eq "Room")
    {
        $Script:RoomAddr = ($Script:RoomAddr + "." + $RoomName + "@ul.com").Replace("..",".")
    }
    else
    {
        $Script:RoomAddr = ($Script:RoomAddr + "." + $RoomName + "." + $Global:txtRoomCap.Text + "@ul.com").Replace("..",".")        
    }

    $Script:txtNewRoomName.Text = ($Script:MbxLoc).ToUpper() + " "  + $MbxName
}

########################################

Check-Reconnect
$FileName		= "RoomResourceRename"
$ReportDirectory	= "E:\Automation\RoomResourceRename\Report\Report-RoomResourceRename-"
$Acro = Import-Csv e:\O365AdminShared\Data\KnownAcronyms.csv

$Global:form = ""

Do
{
    
    Build-DefaultForm
    Publish-Form

    If ($Global:Result -eq "OK")
    {
        If ($Script:chkRenameChgRR.Checked -eq "Checked")
        {
#Removed code and place in the Build-NewDetails function
      	    $Cont = "N"
            write-host "Continue with Renaming this object (Y/N) ? "-ForegroundColor Yellow -NoNewline
            $Cont = Read-Host

   		    If ($Cont -eq "Y")
		    {
    		    If ($NewName -ne $Script:RoomDetails.Name)
			    {
				    Set-Mailbox $Script:RoomDetails.Name -Name $NewName -Alias $NewAlias -DisplayName $NewName -ResourceCapacity $NewCapacity -Office $NewLocation -EmailAddresses $NewPrimaryAlias
			    }
			    else
			    {
				    Set-Mailbox $NewName -Alias $NewAlias -ResourceCapacity $NewCapacity -EmailAddresses $NewPrimaryAlias -Office $NewLocation
			    }
			    Write-Host
			    Write-Host "Room/Resource " $Script:RoomDetails.Name " renamed to " $NewName
			    Write-Host "If this is a restricted room please use the Rename DistributionGroup script to rename the groups used to configure the room delegates and who is allowed to book the room"
            }
		    else
	        {
			    Write-Host "Room not renamed - process has been cancelled"
		    }
		
		    pause
        }

        $Global:Result = "OK"
    }
}While ($Global:Result -eq "OK")	