#####################################################################################
#
#   Display User Mailbox Information
#
#   11/22/2020 - SAG - Created New Script from originals SDAPAdminMenu details
#
#####################################################################################

Function Build-RoomResourceInfoForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Room/Resource Mailbox Information" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 740 ; $Global:form.Height = 460  # Make the form wider 
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 
    
    $Tab = 0
    $Top = 10
    ## DisplayName
    $Global:lblName = New-Object System.Windows.Forms.Label   
        $Global:lblName.Text = "Display Name:"  
        $Global:lblName.Top = $Top ; $Global:lblName.Left = 5; $Global:lblName.Width=120 ;$Global:lblName.AutoSize = $true 
        $form.Controls.Add($Global:lblName)    # Add to Form 
        # 
        $Global:txtName = New-Object Windows.Forms.TextBox
        $Global:txtName.ReadOnly = $true;
        $Global:txtName.TabIndex = $Tab++ # set Tab Order  
        $Global:txtName.Top = $Top; $Global:txtName.Left = 130; $Global:txtName.Width = 200;  
        $Global:txtName.Text = $Room.DisplayName
        $Global:form.Controls.Add($Global:txtName)    # Add to Form 

    $Top = $Top + 30
    ##ResourceType
    $Global:lblRRType = New-Object System.Windows.Forms.Label   
        $Global:lblRRType.Text = "Resource Type:"  
        $Global:lblRRType.Top = $Top ; $Global:lblRRType.Left = 5; $Global:lblRRType.Width=120 ;$Global:lblRRType.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblRRType)    # Add to Form 
        # 
        $Global:txtRRType = New-Object Windows.Forms.TextBox
        $Global:txtRRType.ReadOnly = $true;  
        $Global:txtRRType.TabIndex = $Tab++ # set Tab Order 
        $Global:txtRRType.Top = $Top; $Global:txtRRType.Left = 130; $Global:txtRRType.Width = 200;
        $Global:txtRRType.Text = $Room.ResourceType
        $Global:form.Controls.Add($Global:txtRRType)    # Add to Form

    ##RoomList
    If ($Global:txtRRType.Text -eq "Room")
    {
        $Top = $Top + 30
        $Global:lblRoomList = New-Object System.Windows.Forms.Label   
            $Global:lblRoomList.Text = "Room List:"  
            $Global:lblRoomList.Top = $Top ; $Global:lblRoomList.Left = 5; $Global:lblRoomList.Width=120 ;$Global:lblRoomList.AutoSize = $true 
            $Global:form.Controls.Add($Global:lblRoomList)    # Add to Form 
            # 
            $Global:txtRoomList = New-Object Windows.Forms.TextBox
            $Global:txtRoomList.ReadOnly = $true;  
            $Global:txtRoomList.TabIndex = $Tab++ # set Tab Order 
            $Global:txtRoomList.Top = $Top; $Global:txtRoomList.Left = 130; $Global:txtRoomList.Width = 200;
            $Global:txtRoomList.Text = $RoomList
            $Global:form.Controls.Add($Global:txtRoomList)    # Add to Form
    }

    $Top = $Top + 30
    ##RoomUse
    $Global:lblRoomUse = New-Object System.Windows.Forms.Label   
        $Global:lblRoomUse.Text = "Usage Type:"  
        $Global:lblRoomUse.Top = $Top ; $Global:lblRoomUse.Left = 5; $Global:lblRoomUse.Width=120 ;$Global:lblRoomUse.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblRoomUse)    # Add to Form 
        # 
        $Global:txtRoomUse = New-Object Windows.Forms.TextBox
        $Global:txtRoomUse.ReadOnly = $true;  
        $Global:txtRoomUse.TabIndex = $Tab++ # set Tab Order 
        $Global:txtRoomUse.Top = $Top; $Global:txtRoomUse.Left = 130; $Global:txtRoomUse.Width = 200;
        $Global:txtRoomUse.Text = "General Use"
        If ($CalPro.ResourceDelegates -notlike "*OutOfPolicy.DE")
        {
            $Global:txtRoomUse.Text = "Restricted"
        }
        $Global:form.Controls.Add($Global:txtRoomUse)    # Add to Form

    If ($Global:txtRRType.Text -eq "Room")
        {
        $Top = $Top + 30
        ##RoomLocation
        $Global:lblLocation = New-Object System.Windows.Forms.Label   
            $Global:lblLocation.Text = "Room Location:"
            $Global:lblLocation.Top = $Top ; $Global:lblLocation.Left = 5; $Global:lblLocation.Width=120 ;$Global:lblLocation.AutoSize = $true 
            $Global:form.Controls.Add($Global:lblLocation)    # Add to Form 
            # 
            $Global:txtLocation = New-Object Windows.Forms.TextBox
            $Global:txtLocation.ReadOnly = $true; 
            $Global:txtLocation.TabIndex = $Tab++ # set Tab Order 
            $Global:txtLocation.Top = $Top; $Global:txtLocation.Left = 130; $Global:txtLocation.Height = 50; $Global:txtLocation.Width = 200;
            $Global:txtLocation.Text = $Room.Office
            $Global:form.Controls.Add($Global:txtLocation)    # Add to Form

        $Top = $Top + 30
        ##RoomCapacity
        $Global:lblCapacity = New-Object System.Windows.Forms.Label   
            $Global:lblCapacity.Text = "Room Capacity:"
            $Global:lblCapacity.Top = $Top ; $Global:lblCapacity.Left = 5; $Global:lblCapacity.Width=120 ;$Global:lblCapacity.AutoSize = $true 
            $Global:form.Controls.Add($Global:lblCapacity)    # Add to Form 
            # 
            $Global:txtCapacity = New-Object Windows.Forms.TextBox
            $Global:txtCapacity.ReadOnly = $true; 
            $Global:txtCapacity.TabIndex = $Tab++ # set Tab Order 
            $Global:txtCapacity.Top = $Top; $Global:txtCapacity.Left = 130; $Global:txtCapacity.Height = 50; $Global:txtCapacity.Width = 50;
            $Global:txtCapacity.Text = $Room.ResourceCapacity
            $Global:form.Controls.Add($Global:txtCapacity)    # Add to Form
    }

    $Top = $Top + 30
    ##RoomPolicy
    $Global:lblPolicy = New-Object System.Windows.Forms.Label   
        $Global:lblPolicy.Text = "Room Booking Policy:"
        $Global:lblPolicy.Top = $Top ; $Global:lblPolicy.Left = 5; $Global:lblPolicy.Width=120 ;$Global:lblPolicy.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblPolicy)    # Add to Form 
        # 
        $Global:txtDays = New-Object Windows.Forms.TextBox
        $Global:txtDays.ReadOnly = $true; 
        $Global:txtDays.TabIndex = $Tab++ # set Tab Order 
        $Global:txtDays.Top = $Top; $Global:txtDays.Left = 130; $Global:txtDays.Height = 50; $Global:txtDays.Width = 30;
        $Global:txtDays.Text = $CalPro.BookingWindowInDays
        $Global:form.Controls.Add($Global:txtDays)    # Add to Form
        $Global:lblDays = New-Object System.Windows.Forms.Label   
        $Global:lblDays.Text = "Days into the Future:"
        $Global:lblDays.Top = $Top + 3 ; $Global:lblDays.Left = 160; $Global:lblDays.Width=100 ;$Global:lblDays.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDays)    # Add to Form
        $Global:txtHours = New-Object Windows.Forms.TextBox
        $Global:txtHours.ReadOnly = $true; 
        $Global:txtHours.TabIndex = $Tab++ # set Tab Order 
        $Global:txtHours.Top = $Top; $Global:txtHours.Left = 280; $Global:txtHours.Height = 50; $Global:txtHours.Width = 30;
        $Global:txtHours.Text = ($CalPro.MaximumDurationInMinutes/60)
        $Global:form.Controls.Add($Global:txtHours)    # Add to Form
        $Global:lblHours = New-Object System.Windows.Forms.Label   
        $Global:lblHours.Text = "Maximum Length in Hours:"
        $Global:lblHours.Top = $Top + 3 ; $Global:lblHours.Left = 310; $Global:lblHours.Width=100 ;$Global:lblHours.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblHours)    # Add to Form   

    $Top = $Top + 30
    ##RoomTimeZone
    $Global:lblTimeZone = New-Object System.Windows.Forms.Label   
        $Global:lblTimeZone.Text = "Room TimeZone:"
        $Global:lblTimeZone.Top = $Top ; $Global:lblTimeZone.Left = 5; $Global:lblTimeZone.Width=120 ;$Global:lblTimeZone.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblTimeZone)    # Add to Form 
        # 
        $Global:txtTimeZone = New-Object Windows.Forms.TextBox
        $Global:txtTimeZone.ReadOnly = $true; 
        $Global:txtTimeZone.TabIndex = $Tab++ # set Tab Order 
        $Global:txtTimeZone.Top = $Top; $Global:txtTimeZone.Left = 130; $Global:txtTimeZone.Height = 50; $Global:txtTimeZone.Width = 540;
        $Global:txtTimeZone.Text = $RegConf.TimeZone
        $Global:form.Controls.Add($Global:txtTimeZone)    # Add to Form

    $Top = $Top + 30
    ##RoomDelegates
    $Global:lblDelegates = New-Object System.Windows.Forms.Label   
        $Global:lblDelegates.Text = "Delegate Group:"
        $Global:lblDelegates.Top = $Top ; $Global:lblDelegates.Left = 5; $Global:lblDelegates.Width=120 ;$Global:lblDelegates.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDelegates)    # Add to Form 
        # 
        $Global:txtDelegates = New-Object Windows.Forms.TextBox
        $Global:txtDelegates.ReadOnly = $true; 
        $Global:txtDelegates.TabIndex = $Tab++ # set Tab Order 
        $Global:txtDelegates.Top = $Top; $Global:txtDelegates.Left = 130; $Global:txtDelegates.Height = 50; $Global:txtDelegates.Width = 540;
        $Global:txtDelegates.Text = $CalPro.ResourceDelegates
        $Global:form.Controls.Add($Global:txtDelegates)    # Add to Form

    $Top = $Top + 30
    ##RoomDelegate Members
    $Global:lblDelegateMem = New-Object System.Windows.Forms.Label   
        $Global:lblDelegateMem.Text = "Delegate Members:"
        $Global:lblDelegateMem.Top = $Top ; $Global:lblDelegateMem.Left = 5; $Global:lblDelegateMem.Width=120 ;$Global:lblDelegateMem.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDelegateMem)    # Add to Form 
        # 
        $Global:txtDelegateMem = New-Object Windows.Forms.TextBox
        $Global:txtDelegateMem.ReadOnly = $true;
        $Global:txtDelegateMem.TabIndex = $Tab++ # set Tab Order 
        $Global:txtDelegateMem.Top = $Top; $Global:txtDelegateMem.Left = 130; $Global:txtDelegateMem.Height = 50; $Global:txtDelegateMem.Width = 540;
        $Global:txtDelegateMem.Text = $DelegateMem
        $Global:form.Controls.Add($Global:txtDelegateMem)    # Add to Form

    If ($AuthUS.Name -like "*OutOfPolicy*")
    {
        $Global:form.Height = 530  # Make the form wider 
        $Top = $Top + 30
        ##RestrictedRoomUsers
        $Global:lblRestrictedUsrGrp = New-Object System.Windows.Forms.Label   
            $Global:lblRestrictedUsrGrp.Text = "Restricted User Group:"
            $Global:lblRestrictedUsrGrp.Top = $Top ; $Global:lblRestrictedUsrGrp.Left = 5; $Global:lblRestrictedUsrGrp.Width=120 ;$Global:lblRestrictedUsrGrp.AutoSize = $true 
            $Global:form.Controls.Add($Global:lblRestrictedUsrGrp)    # Add to Form 
            # 
            $Global:txtRestrictedUsrGrp = New-Object Windows.Forms.TextBox
            $Global:txtRestrictedUsrGrp.ReadOnly = $true; 
            $Global:txtRestrictedUsrGrp.TabIndex = $Tab++ # set Tab Order 
            $Global:txtRestrictedUsrGrp.Top = $Top; $Global:txtRestrictedUsrGrp.Left = 130; $Global:txtRestrictedUsrGrp.Height = 50; $Global:txtRestrictedUsrGrp.Width = 540;
            $Global:txtRestrictedUsrGrp.Text = $AuthUs
            $Global:form.Controls.Add($Global:txtRestrictedUsrGrp)    # Add to Form
    
        $Top = $Top + 30
        ##RestrictedRoomUser Members
        $Global:lblUserMem = New-Object System.Windows.Forms.Label   
            $Global:lblUserMem.Text = "Restricted Users:"
            $Global:lblUserMem.Top = $Top ; $Global:lblUserMem.Left = 5; $Global:lblUserMem.Width=120 ;$Global:lblUserMem.AutoSize = $true
            $Global:form.Controls.Add($Global:lblUserMem)    # Add to Form 
            # 
            $Global:txtUserMem = New-Object Windows.Forms.TextBox
            $Global:txtUserMem.ReadOnly = $true;
            $Global:txtUserMem.TabIndex = $Tab++ # set Tab Order 
            $Global:txtUserMem.Top = $Top; $Global:txtUserMem.Left = 130; $Global:txtUserMem.Height = 50; $Global:txtUserMem.Width = 540;
            $Global:txtUserMem.Text = $AuthUsMem
            $Global:form.Controls.Add($Global:txtUserMem)    # Add to Form
    }

    $Top = $Top + 60
    ##ReportDetails
    $Global:lblRptFile = New-Object System.Windows.Forms.Label   
        $Global:lblRptFile.Text = "Report Details:"  
        $Global:lblRptFile.Top = $Top ; $Global:lblRptFile.Left = 5; $Global:lblRptFile.Width=100 ;$Global:lblRptFile.AutoSize = $true 
        $form.Controls.Add($Global:lblRptFile)    # Add to Form 
        # 
        $Global:txtRptFile = New-Object Windows.Forms.TextBox
        $Global:txtRptFile.ReadOnly = $true; 
        $Global:txtRptFile.TabIndex = $Tab++ # set Tab Order
        $Global:txtRptFile.Top = $Top; $Global:txtRptFile.Left = 130; $Global:txtRptFile.Width = 540; 
        $Global:txtRptFile.Text = $ReportFile
        $Global:form.Controls.Add($Global:txtRptFile)    # Add to Form

    Add-FormInfoOnlyButtons
}

$Filename       = "RoomInfo"
$LogFile		= "E:\SDAP\RoomInfo\Log\Log-" + $FileName + ".log"
$AuthUs         = ""
$AuthUsMem      = ""
$DelegateMem    = ""

Enter-MbxInputForm
Publish-form

If ($Global:Result -eq "OK")
{
    $RoomExists = [bool](get-mailbox $Global:txtInpMailbox.Text -ErrorAction SilentlyContinue)
    If ($RoomExists -eq "True")
    {
        write-host "Obtaining room/resource mailbox information....." -ForegroundColor Cyan
       
        $Room = get-mailbox $Global:txtInpMailbox.Text
        write-host "Obtaining room/resource calendar processing configuration....." -ForegroundColor Cyan
        $CalPro = get-calendarprocessing $Global:txtInpMailbox.Text -ErrorAction SilentlyContinue
        $DelegateMem = ((get-distributiongroupmember $CalPro.ResourceDelegates[0]).Name -join ", ")
        $RegConf = Get-MailboxRegionalConfiguration $Global:txtInpMailbox.Text
                            
        If ($CalPro.BookInPolicy[0] -like "*OutOfPolicy*")
        {
            $AuthUs = (get-DistributionGroup $CalPro.BookInPolicy[0]).DisplayName
            $AuthUsMem = ((Get-DistributionGroupMember $AuthUs) -join ", ")
        }

        If ($Room.ResourceType -eq "Room")
        {
            write-host "Obtaining the Room List..." -ForegroundColor Cyan
            $RoomList = Get-DistributionGroup (($Room.Name).Substring(0,3) + " Conference Rooms")
            If ($CalPro.ResourceDelegates -notlike "*OutOfPolicy.DE")
            {
                $RoomList = Get-DistributionGroup (($Room.Name).Substring(0,3) + " Restricted Rooms")
            }

            If ([bool](Get-DistributionGroupMember $RoomList.Name |where {$_.Name -eq $Room.DisplayName}))
            {
                write-host "Obtaining the Room List Members..." -ForegroundColor Cyan
                $RoomListMem = Get-DistributionGroupMember $RoomList.Name
            }
            else
            {
                write-host "Obtaining the Room List that this Room is a member of -- this may take some time to complete..." -ForegroundColor Yellow
                $RoomList = Get-Group -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "RoomList"} |Where-Object -FilterScript {$_.Members -contains $Room.DisplayName}
                $RoomListMem = Get-DistributionGroupMember $RoomList.Name
            }
        }

        If ($CalPro.BookInPolicy[0].Length -ne 0)
        {
            If ((get-DistributionGroup $CalPro.BookInPolicy[0]) -like "*OutOfPolicy*")
            {
                $AuthUs = get-DistributionGroup $CalPro.BookInPolicy[0]
                $AuthUsMem = ((Get-DistributionGroupMember $CalPro.BookInPolicy[0]) -join ", ")
            }
        }

        Build-RoomResourceInfoForm
        Publish-Form

        #write details to Report File
        $ReportFile	= "E:\SDAP\RoomInfo\Report\Report-RoomResource-" + $Room.DisplayName.Replace(" ","") + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
        $LineToWrite = "STAR" + "`t" + $FileName + " script has started"
        WriteLogEvent
        $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
        WriteLogEvent
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "Room/Resource Display Name:  "  + $Room.DisplayName
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "             Resource Type:  "  + $Room.ResourceType
        WriteReportEvent
        If ($Room.ResourceType -like "Room")
        {
            $LineToWrite = $WhoAmI + "`t" + "                 Room List:  "  + $RoomList
            WriteReportEvent
        }
        $LineToWrite = $WhoAmI + "`t" + "                Usage Type:  "  + $Global:txtRoomUse.Text
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "             Room Location:  "  + $Room.Office
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "             Room Capacity:  "  + $Room.ResourceCapacity
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "            Booking Policy:  "  + $Global:txtDays.Text + " Days into the Future, " + $Global:txtHours.Text + " Maximum Length in Hours"
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "                  TimeZone:  "  + $Global:txtTimeZone.Text
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "            Delegate Group:  "  + $CalPro.ResourceDelegates
        WriteReportEvent
        $LineToWrite = $WhoAmI + "`t" + "            Delegate Users:  "  + $DelegateMem
        WriteReportEvent
        If ($AuthUS.Name -like "*OutOfPolicy*")
        {
            $LineToWrite = $WhoAmI + "`t" + "     Restricted User Group:  "  + $AuthUs
            WriteReportEvent
            $LineToWrite = $WhoAmI + "`t" + "          Restricted Users:  "  + $AuthUsMem
            WriteReportEvent
        }
    }
    else
    {
       $Output = $wshell.Popup("Mailbox " + $Global:txtInpMailbox.Text + " does not exist.",0,"No Mailbox",0+32)
    }
}
else
{
    $Output = $wshell.Popup("Request cancelled.",0,"Cancelled",0+32)
}
write-host "Room/resource mailbox information complete.....`n" -ForegroundColor Red