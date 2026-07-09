#################################################################################
# 
# PowerShell source code
# Revision v1.1
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Create new Conference Rooms in Exchange Online
#    'Called By    : RoomResourceMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 09/16/2022
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      : 
#    '             : 09/16/2022 - Created from the New Conference Room Script incorporating elements of the modifying access to Room Delegates/Restricted Rooms
#    '
# ==========================================================================
#
#################################################################################

Function Build-DefaultForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Room/Equipment Booking Configuration" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(600,200) #(W,H)

    $Global:OKDetails = "Update"
    Add-FormStandardButtons
    
    $Top = 10 
    ## Label and TextBox  
    ## Display Name
    $Script:lblRoomName = New-Object System.Windows.Forms.Label   
        $lblRoomName.Text = "Room Name:"
        $lblRoomName.Top = $Top ; $lblRoomName.Left = 10; $lblRoomName.Width=120 ;$lblRoomName.AutoSize = $true 
        $Global:form.Controls.Add($lblRoomName)    # Add to Form 
        # 
        $Script:txtRoomName = New-Object Windows.Forms.TextBox
        $Script:txtRoomName.TabIndex = 1
        $Script:txtRoomName.Top = 10; $txtRoomName.Left = 130; $txtRoomName.Width = 300;  
        $Script:txtRoomName.Text = ""  # DisplayName
        $Global:form.Controls.Add($Script:txtRoomName)    # Add to Form
        $Global:InputFocus = $Script:txtRoomName
        $Script:txtRoomName.Add_Click({
            Hide-Fields
            $Script:txtRoomName.ReadOnly = $False
            $Script:ButGetGrp.visible = $True
        })

    $Top = $Top + 30
    ## Ticket Number
    $Script:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Script:lblTaskNo.Text = "Ticket Number:"
        $Script:lblTaskNo.Top = $Top; $Script:lblTaskNo.Left = 10; $Script:lblTaskNo.Width=150 ; $Script:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Script:lblTaskNo)    # Add to Form 
        # 
        $Script:txtTaskNo = New-Object Windows.Forms.TextBox  
        $Script:txtTaskNo.TabIndex = 2
        $Script:txtTaskNo.Top = $Top; $Script:txtTaskNo.Left = 130; $Script:txtTaskNo.Width = 120;  
        $Script:txtTaskNo.Text = "TASK"   # Enter ticket number 
        $Global:form.Controls.Add($Script:txtTaskNo)    # Add to Form 

    #Get Group Details Button
    $Script:ButGetGrp = New-Object Windows.Forms.Button
    $Script:ButGetGrp.Location = New-object System.Drawing.Size(430,10)
    $Script:ButGetGrp.Size = new-Object System.Drawing.Size(110,20)
    $Script:ButGetGrp.Text = "Get Room Details"
    $Script:ButGetGrp.TabIndex = 3
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
    ## EmailAddress
    $Script:lblRoomAddr = New-Object System.Windows.Forms.Label   
        $Script:lblRoomAddr.Text = "Email Address:"  
        $Script:lblRoomAddr.Top = $Top ; $Script:lblRoomAddr.Left = 10; $Script:lblRoomAddr.Width=150 ;$Script:lblRoomAddr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblRoomAddr)    # Add to Form 
        # 
        $Script:txtRoomAddr = New-Object Windows.Forms.TextBox  
        $Script:txtRoomAddr.Top = $Top; $Script:txtRoomAddr.Left = 130; $Script:txtRoomAddr.Width = 400;
        $Script:txtRoomAddr.TabStop = $False
        $Script:txtRoomAddr.ReadOnly = $true
        $Global:form.Controls.Add($Script:txtRoomAddr)    # Add to Form

    $Top = $Top + 30
    ## Room Delegates
    $Script:lblRoomDele = New-Object System.Windows.Forms.Label   
        $lblRoomDele.Text = "Room Delegates:"  
        $lblRoomDele.Top = $Top; $lblRoomDele.Left = 10; $lblRoomDele.Width=150 ;$lblRoomDele.AutoSize = $true 
        $Global:form.Controls.Add($lblRoomDele)    # Add to Form 
        # 
        $Script:txtRoomDele = New-Object Windows.Forms.TextBox  
        $Script:txtRoomDele.Top = $Top; $Script:txtRoomDele.Left = 130; $Script:txtRoomDele.Width = 400;  
        $Script:txtRoomDele.Text = "Retrieving Details...."  # Legacy Address
        $Script:txtRoomDele.TabStop = $False
        $Script:txtRoomDele.ReadOnly = $true
        $Global:form.Controls.Add($Script:txtRoomDele)    # Add to Form 

    $Top = $Top + 30
    ## Room Users
    $Script:lblRoomUsers = New-Object System.Windows.Forms.Label   
        $Script:lblRoomUsers.Text = "Room Users:"  
        $Script:lblRoomUsers.Top = $Top; $Script:lblRoomUsers.Left = 10; $Script:lblRoomUsers.Width=150; $Script:lblRoomUsers.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblRoomUsers)    # Add to Form 
        # 
        $Script:txtRoomUsers = New-Object Windows.Forms.TextBox  
        $Script:txtRoomUsers.Top = $Top; $Script:txtRoomUsers.Left = 130; $Script:txtRoomUsers.Width = 400;  
        $Script:txtRoomUsers.Text = "Retrieving Details...."   # Use Corrent computer name as default
        $Script:txtRoomUsers.TabStop = $False
        $Script:txtRoomUsers.ReadOnly = $true
        $Global:form.Controls.Add($Script:txtRoomUsers)    # Add to Form 

    $Top = $Top + 30
    ## Room Delegate Group Membership
    $Script:lblDeleGrpMbr = New-Object System.Windows.Forms.Label   
        $Script:lblDeleGrpMbr.Text = "Delegate Group Member(s):"  
        $Script:lblDeleGrpMbr.Top = $Top ; $Script:lblDeleGrpMbr.Left = 10; $Script:lblDeleGrpMbr.Width=200; $Script:lblDeleGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblDeleGrpMbr)    # Add to Form 
    ## Room User Group Membership
    $Script:lblUserGrpMbr = New-Object System.Windows.Forms.Label   
        $Script:lblUserGrpMbr.Text = "User Group Member(s):"  
        $Script:lblUserGrpMbr.Top = $Top ; $Script:lblUserGrpMbr.Left = 310; $Script:lblUserGrpMbr.Width=150; $Script:lblUserGrpMbr.AutoSize = $true
        $Global:form.Controls.Add($Script:lblUserGrpMbr)    # Add to Form 
        # 
    $Top = $Top + 20
    $Script:txtDeleGrpMbr = New-Object Windows.Forms.TextBox
        $Script:txtDeleGrpMbr.MaxLength = 200000
        $Script:txtDeleGrpMbr.Location = New-Object System.Drawing.Size(10,$Top)
        $Script:txtDeleGrpMbr.Size = New-Object system.Drawing.Size(300,60)
        $Script:txtDeleGrpMbr.Text = "Retrieving Details...." 
        $Script:txtDeleGrpMbr.MultiLine = $true
        $Script:txtDeleGrpMbr.ScrollBars = 'Both'
        $Script:txtDeleGrpMbr.ReadOnly = $true
        $Script:txtDeleGrpMbr.TabStop = $False  
        $Global:form.Controls.Add($Script:txtDeleGrpMbr)    # Add to Form 

   $Script:txtUserGrpMbr = New-Object Windows.Forms.TextBox
        $Script:txtUserGrpMbr.MaxLength = 200000
        $Script:txtUserGrpMbr.Location = New-Object System.Drawing.Size(310,$Top)
        $Script:txtUserGrpMbr.Size = New-Object system.Drawing.Size(300,60)
        $Script:txtUserGrpMbr.Text = "Retrieving Details...." 
        $Script:txtUserGrpMbr.MultiLine = $true
        $Script:txtUserGrpMbr.ScrollBars = 'Both'
        $Script:txtUserGrpMbr.ReadOnly = $true
        $Script:txtUserGrpMbr.TabStop = $False
        $Global:form.Controls.Add($Script:txtUserGrpMbr)    # Add to Form 
       
    $Top = $Top + 70
    ## Add Room Delegates
    $Script:chkAddDele = New-Object Windows.Forms.checkbox 
        $Script:chkAddDele.Left = 30; $Script:chkAddDele.Width = 250; $Script:chkAddDele.Top = $Top
        $Script:chkAddDele.Text = "Add Room Delegates"
        $Script:chkAddDele.TabIndex=4
        $Script:chkAddDele.Checked = $false   # set a default value
        $Global:form.Controls.Add($Script:chkAddDele)
        $Script:chkAddDele.Add_Click({
            $Script:lblDeleChg.Text = "Add Room Delegates:"
            $Global:okButton.Text = "Add"
            If($Script:chkRemUser.Checked -eq $True)
            {
                $Global:okButton.Text = "Remove"
            }
            Display-AuthDele
        })

    ## Add Room Users
    $Script:chkAddUser = New-Object Windows.Forms.checkbox 
        $Script:chkAddUser.Left = 330; $Script:chkAddUser.Width = 250; $Script:chkAddUser.Top = $Top
        $Script:chkAddUser.Text = "Add Room Users"
        $Script:chkAddUser.TabIndex=6
        $Script:chkAddUser.Checked = $false   # set a default value
        $Global:form.Controls.Add($Script:chkAddUser)
        $Script:chkAddUser.Add_Click({
            $Script:lblUserChg.Text = "Add Room Users:"
            $Global:okButton.Text = "Add"
            If($Script:chkRemDele.Checked -eq $True)
            {
                If ($Script:chkAddUser.Checked -eq $True)
                {
                    $Global:okButton.Text = "Add/Remove"
                    $Global:okButton.Refresh
                }
            }
            Display-AuthUser
        })

    ## Enter Room Delegate Changes
    $Script:lblDeleChg = New-Object System.Windows.Forms.Label   
        $Script:lblDeleChg.Top = $Top ; $Script:lblDeleChg.Left = 10; $Script:lblDeleChg.Width=200; $Script:lblDeleChg.AutoSize = $true
        $Script:lblDeleChg.Text = "Room Delegate Changes:"
        $Global:form.Controls.Add($Script:lblDeleChg)    # Add to Form 
        # 
        $Script:txtDeleChg = New-Object Windows.Forms.TextBox
        $Script:txtDeleChg.Top = $Top+20; $Script:txtDeleChg.Left = 10; $Script:txtDeleChg.Width = 290;
        $Script:txtDeleChg.Text = "Enter Address or Employee# separate with commas...." 
        $Script:txtDeleChg.TabIndex=10
        $Global:form.Controls.Add($Script:txtDeleChg)    # Add to Form 
        $Script:txtDeleChg.Add_Click({
            If ($Script:txtDeleChg.Text -like "Enter Address*")
            {
                $Script:txtDeleChg.Text = ""
            }
            $Global:OKButton.visible = $True
            $Global:CancelButton.visible = $True 
        })

    ## Enter Room User Changes
    $Script:lblUserChg = New-Object System.Windows.Forms.Label   
        $Script:lblUserChg.Top = $Top ; $Script:lblUserChg.Left = 310; $Script:lblUserChg.Width=200; $Script:lblUserChg.AutoSize = $true
        $Script:lblUserChg.Text = "Room User Changes"
        $Global:form.Controls.Add($Script:lblUserChg)    # Add to Form 
        # 
        $Script:txtUserChg = New-Object Windows.Forms.TextBox
        $Script:txtUserChg.Top = $Top+20; $Script:txtUserChg.Left = 310; $Script:txtUserChg.Width = 290;
        $Script:txtUserChg.Text = "Enter Address or Employee# separate with commas...."
        $Script:txtUserChg.TabIndex=11 
        $Global:form.Controls.Add($Script:txtUserChg)    # Add to Form 
        $Script:txtUserChg.Add_Click({
            If ($Script:txtUserChg.Text -like "Enter Address*")
            {
                $Script:txtUserChg.Text = ""
            }
            $Global:OKButton.visible = $True
            $Global:CancelButton.visible = $True  
        })

    $Top = $Top + 20
    ## Remove Room Delegates
    $Script:chkRemDele = New-Object Windows.Forms.checkbox 
        $Script:chkRemDele.Left = 30; $Script:chkRemDele.Width = 250; $Script:chkRemDele.Top = $Top
        $Script:chkRemDele.Text = "Remove Room Delegates"
        $Script:chkRemDele.TabIndex=5 
        $Script:chkRemDele.Checked = $false   # set a default value
        $Global:form.Controls.Add($Script:chkRemDele)
        $Script:chkRemDele.Add_Click({
            $Script:lblDeleChg.Text = "Remove Room Delegates:"
            $Global:okButton.Text = "Remove"
            Display-AuthDele
        })
        
    ## Remove Room Users
    $Script:chkRemUser = New-Object Windows.Forms.checkbox 
        $Script:chkRemUser.Left = 330; $Script:chkRemUser.Width = 250; $Script:chkRemUser.Top = $Top
        $Script:chkRemUser.Text = "Remove Room Users"
        $Script:chkRemUser.TabIndex=7
        $Script:chkRemUser.Checked = $false   # set a default value
        $Global:form.Controls.Add($Script:chkRemUser)
        $Script:chkRemUser.Add_Click({
            $Script:lblUserChg.Text = "Remove Room Users:"
            $Global:okButton.Text = "Remove"
            Display-AuthUser
        })                                   

    $Top = $Top + 30
    ## Not for Display Box to Not Display the email template
    $Script:chkTemplate = New-Object Windows.Forms.checkbox 
        $Script:chkTemplate.Left = 130; $Script:chkTemplate.Width = 350; $Script:chkTemplate.Top = $Top
        $Script:chkTemplate.Text = "Display Email Template for Room/Equipment Booking Changes"
        $Script:chkTemplate.TabIndex=15 
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
        $Script:txtRptFile.TabStop = $False
        $Script:txtRptFile.Text = $ReportFile
        $Global:form.Controls.Add($Script:txtRptFile)    # Add to Form

    Hide-Fields
}

Function Hide-Fields
{
    $Script:lblRoomAddr.Visible = $False
    $Script:txtRoomAddr.Visible = $False
    $Script:lblRoomDele.Visible = $False
    $Script:txtRoomDele.Visible = $False
    $Script:lblRoomUsers.Visible = $False
    $Script:txtRoomUsers.Visible = $False
    $Script:lblDeleGrpMbr.Visible = $False
    $Script:txtDeleGrpMbr.Visible = $False
    $Script:lblRoomUsers.Visible = $False
    $Script:txtRoomUsers.Visible = $False
    $Script:lblUserGrpMbr.Visible = $False
    $Script:txtUserGrpMbr.Visible = $False
    $Script:lblDeleChg.Visible = $False
    $Script:txtDeleChg.Visible = $False
    $Script:lblUserChg.Visible = $False
    $Script:txtUserChg.Visible = $False
    $Script:lblRptFile.Visible = $False
    $Script:txtRptFile.Visible = $False
    $Script:chkTemplate.Visible = $False
    $Global:okButton.Visible = $False
    $Global:CancelButton.visible = $False 

    $Script:chkAddDele.Checked = $False
    $Script:chkAddDele.visible = $False
    $Script:chkAddUser.Checked = $False
    $Script:chkAddUser.visible = $False

    $Script:chkRemDele.Visible = $False
    $Script:chkRemDele.Checked = $False
    $Script:chkRemUser.Checked = $False
    $Script:chkRemUser.Visible = $False
    $Script:txtRoomUsers.Visible = $False
    $Script:txtRoomUsers.Text = "Retrieving Details...."
    $Script:txtUserGrpMbr.Text = "Retrieving Details...."
    $Script:txtDeleChg.Text = "Enter Address or Employee# separate with commas...."
    $Script:txtUserChg.Text = "Enter Address or Employee# separate with commas...."
    $Global:form.Size = New-Object System.Drawing.Size(600,200) #(W,H)
}

Function Complete-Form
{
    $Script:txtRoomName.Text = ($Script:txtRoomName.Text).Trim()
#    $Script:txtRoom.Refresh
    $Global:form.Size = New-Object System.Drawing.Size(640,430) #(W,H)
    $Script:lblRoomAddr.Visible = $True
    $Script:txtRoomAddr.Visible = $True
    $Script:lblRoomDele.Visible = $True
    $Script:txtRoomDele.Visible = $True
    $Script:lblRoomUsers.Visible = $True
    $Script:txtRoomUsers.Visible = $True
    $Script:lblDeleGrpMbr.Visible = $True
    $Script:txtDeleGrpMbr.Visible = $True
    $Script:lblUserGrpMbr.Visible = $True
    $Script:txtUserGrpMbr.Visible = $True
    $Script:chkTemplate.Visible = $True
    $Script:lblRptFile.Visible = $True
    $Script:txtRptFile.Visible = $True
    $Script:txtRoomAddr.Text = $Script:RoomDetails.PrimarySmtpAddress
    $Script:txtRoomDele.Text = (get-CalendarProcessing $Script:txtRoomName.Text).ResourceDelegates
    $Script:chkRemDele.Visible = $True
    $Script:chkAddDele.visible = $True
    $GenUse = (Get-CalendarProcessing $Script:txtRoomName.Text).BookinPolicy
    If ($GenUse -ne $null)
    {
        $Script:txtRoomUsers.Text = (Get-DistributionGroup ((Get-CalendarProcessing $Script:txtRoomName.Text).BookinPolicy)[0]).DisplayName
        $Script:txtUserGrpMbr.Text = (Get-DistributionGroupMember -ResultSize Unlimited $Script:txtRoomUsers.Text).DisplayName -join "; "
        $Script:chkAddUser.visible = $True
        $Script:chkRemUser.visible = $True
    }
    else
    {
        $Script:txtRoomUsers.Text = "General Use Room - All Users Can Book"
        $Script:txtUserGrpMbr.Text = "General Use Room - All Users Can Book"
    }
    $Script:txtDeleGrpMbr.Text = (Get-DistributionGroupMember -ResultSize Unlimited $Script:txtRoomDele.Text).DisplayName -join "; "

    $Script:ReportFile = $ReportDirectory + $FileName + ($Script:Room -replace(" ","")) + "-Date" + $Date + "Time" + $Time + ".Log"
    $Script:txtRptFile.Text = $Script:ReportFile
    $Script:txtRoomName.ReadOnly = $True
    $Global:CancelButton.visible = $True
    $Script:txtTaskNo.Focus()
}

Function Display-AuthDele
{
    $Script:lblDeleChg.Text = "Add Room Delegates:"
    If($Script:chkRemDele.Checked -eq $True)
    {
        $Script:lblDeleChg.Text = "Remove Room Delegates:"
    }
    If ((($Script:chkAddUser.Checked -eq $True) -and ($Script:chkRemDele.Checked -eq $True)) -or (($Script:chkRemUser.Checked -eq $True) -and ($Script:chkAddDele.Checked -eq $True)))
    {
        $Global:okButton.Text = "Add-Remove"
        $Global:okButton.Width = 80
    }
    $Script:chkAddDele.Visible = $False
    $Script:chkRemDele.Visible = $False
    $Script:lblDeleChg.Visible = $True
    $Script:txtDeleChg.Visible = $True
}

Function Display-AuthUser
{
    $Script:lblUserChg.Text = "Add Room Users:"
    If($Script:chkRemUser.Checked -eq $True)
    {
        $Script:lblUserChg.Text = "Remove Room Users:"
    }
    If ((($Script:chkAddUser.Checked -eq $True) -and ($Script:chkRemDele.Checked -eq $True)) -or (($Script:chkRemUser.Checked -eq $True) -and ($Script:chkAddDele.Checked -eq $True)))
    {
        $Global:okButton.Text = "Add-Remove"
        $Global:okButton.Width = 80
    }
    $Script:chkAddUser.Visible = $False
    $Script:chkRemUser.Visible = $False
    $Script:lblUserChg.Visible = $True
    $Script:txtUserChg.Visible = $True
}

#######################################################################################

$WhoAmI			= WhoAmI

# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name

# Retrieve the current Date and Time for use in log files
	$uDate = get-date -uformat %D
	$uTime = get-date -uformat %T
	$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"
# Convert the current Date and Time to format MM-DD-YY and HHMMSS for use in file names	
	$Date  = $uDate.Replace("/", "-")
	$Time  = $uTime.Replace(":", "")

# =============================================================================================================================================

Check-Reconnect
$FileName		= "RROOPChange-"
$ReportDirectory	= "E:\Automation\RROOPChanges\Report\"
$Global:form = ""
Add-Type -Assembly System.Windows.Forms     ## Load the Windows Forms assembly 

Build-DefaultForm
Publish-Form

If ($Global:Result -eq "OK")
{
	$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI
	Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
	$LineToWrite = $RecordEvent + "STAR" + "`t" + "RoomResourceOOPChange script has started" + "`n"
	Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

    Do{
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Ticket Number                         : " + $Script:txtTaskNo.Text
        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Room/Resource Name                    : " + $Script:Room
	    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Room/Resource Address                 : " + $Script:txtRoomAddr.Text
	    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Room/Resource Delegates Group         : " + $Script:txtRoomDele.Text
	    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Room/Resource Delegate Members        : " + $Script:txtDeleGrpMbr.Text
	    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Room/Resource Users Group             : " + $Script:txtRoomUsers.Text
	    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Room/Resource User Members            : " + $Script:txtUserGrpMbr.Text
	    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

        If (($Script:chkAddUser.Checked -eq $True) -or ($Script:chkRemUser.Checked -eq $True))
        {
            $Script:txtUserChg.Text = ($Script:txtUserChg.Text) -replace (" ","")
            $UsrChg = $Script:txtUserChg.Text -split (",")
            If ($Script:chkAddUser.Checked -eq $True)
            {
                #This section adds additional authorized users
                foreach ($usr in $Usrchg)
                {
                    If ($Usr -notlike "ACL*")
                    {
                        $Exists = [bool]($obj = get-mailbox $usr -ErrorAction SilentlyContinue)
                    }
                    else
                    {
                        $Exists = [bool]($obj = get-DistributionGroup $usr -ErrorAction SilentlyContinue)
                    }
                    If ($Exists -eq $True)
                    {
                        If ($Script:txtRoomUsers.Text -notlike "General*")
                        {
                            Add-DistributionGroupMember $Script:txtRoomUsers.Text -Member $obj.alias -BypassSecurityGroupManagerCheck
                            write-host "Added Resource User: $usr [" $obj.DisplayName "]" -ForegroundColor Green
                            $LineToWrite = $RecordEvent + "ADDUSR" + "`t" + "Added Room User                       : "+ $USR + " [" + $obj.DisplayName + "]"
                        }
                        else
                        {
                            write-host "This is a general use room no changes to the room users made" -ForegroundColor Red
                            $LineToWrite = $RecordEvent + "FAIL" + "`t" + "General Use Room                      : " + $usr
                        }
                    }
                    else
                    {
                        write-host "Unable to add user, no User Account or Group found for $Usr" -ForegroundColor Red
                        $LineToWrite = $RecordEvent + "FAIL" + "`t" + "No User or Group Exists               : " + $usr
                    }
                    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                }
            }
            else
            {
                #This section removes authorized users
                foreach ($usr in $Usrchg)
                {
                    If ($Usr -notlike "ACL*")
                    {
                        $Exists = [bool]($obj = get-mailbox $usr -ErrorAction SilentlyContinue)
                    }
                    else
                    {
                        $Exists = [bool]($obj = get-DistributionGroup $usr -ErrorAction SilentlyContinue)
                    }
                    If ($Exists -eq $True)
                    {
                        Remove-DistributionGroupMember $Script:txtRoomUsers.Text -Member $obj.alias -BypassSecurityGroupManagerCheck -confirm:$false
                        write-host "Removed Resource User: $usr [" $obj.DisplayName "]" -ForegroundColor Green
                        $LineToWrite = $RecordEvent + "REMUSR" + "`t" + "Removed Room User                     : "+ $USR + " [" + $obj.DisplayName + "]"
                    }
                    else
                    {
                        write-host "Unable to remove user, no User Account or Group found for $Usr" -ForegroundColor Red
                        $LineToWrite = $RecordEvent + "FAIL" + "`t" + "No User or Group Exists               : " + $usr
                    }
                    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                }
            }
        }

        If (($Script:chkAddDele.Checked -eq $True) -or ($Script:chkRemDele.Checked -eq $True))
        {
            $Script:txtDeleChg.Text = ($Script:txtDeleChg.Text) -replace (" ","")
            $DeleChg = $Script:txtDeleChg.Text -split (",")
            If ($Script:chkAddDele.Checked -eq $True)
            {
                #This section adds additional room/resource delegates
                foreach ($usr in $DeleChg)
                {
                    If ($Usr -notlike "ACL*")
                    {
                        $Exists = [bool]($obj = get-mailbox $usr -ErrorAction SilentlyContinue)
                    }
                    else
                    {
                        $Exists = [bool]($obj = get-DistributionGroup $usr -ErrorAction SilentlyContinue)
                    }
                    If ($Exists -eq $True)
                    {
                        Add-DistributionGroupMember $Script:txtRoomDele.Text -Member $obj.alias -BypassSecurityGroupManagerCheck
                        write-host "Added Resource Delegate: $usr [" $obj.DisplayName "]" -ForegroundColor Green
                        $LineToWrite = $RecordEvent + "ADDDEL" + "`t" + "Added Room Delegate                   : " + $usr + "[" + $obj.DisplayName+ "]"
                    }
                    else
                    {
                        write-host "Unable to add delegate, no User Account or Group found for $Usr" -ForegroundColor Red
                        $LineToWrite = $RecordEvent + "FAIL" + "`t" + "No User or Group Exists               : " + $usr
                    }
                    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                }
            }
            else
            {
                #This section removes room/resource delegates
                foreach ($usr in $DeleChg)
                {
                    If ($Usr -notlike "ACL*")
                    {
                        $Exists = [bool]($obj = get-mailbox $usr -ErrorAction SilentlyContinue)
                    }
                    else
                    {
                        $Exists = [bool]($obj = get-DistributionGroup $usr -ErrorAction SilentlyContinue)
                    }
                    If ($Exists -eq $True)
                    {
                        Remove-DistributionGroupMember $Script:txtRoomDele.Text -Member $obj.alias -BypassSecurityGroupManagerCheck -confirm:$false
                        write-host "Removed Resource Delegate: $usr [" $obj.DisplayName "]" -ForegroundColor Green
                        $LineToWrite = $RecordEvent + "REMDEL" + "`t" + "Removed Room Delegate                 : " + $usr + "[" + $obj.DisplayName+ "]"
                    }
                    else
                    {
                        write-host "Unable to remove delegate, no User Account or Group found for $Usr" -ForegroundColor Red
                        $LineToWrite = $RecordEvent + "FAIL" + "`t" + "No User or Group Exists               : " + $usr
                    }
                    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                }
            }
        }

        $Script:txtRoomDele.Text = (Get-DistributionGroupMember -ResultSize Unlimited $Script:txtRoomDele.Text).DisplayName -join "; "
        If ($Script:txtRoomUsers.Text -notlike "General*")
        {
            $Script:txtRoomUsers.Text = (Get-DistributionGroupMember -ResultSize Unlimited $Script:txtRoomUsers.Text).DisplayName -join "; "
        }

        Hide-Fields
        Complete-Form
        Publish-Form
    	$LineToWrite = $RecordEvent + "`n"
    	Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
    }While ($Global:Result -eq "OK")

    write-host "Report file can be found at: " $ReportFile
    write-host "Room Update Complete" -ForegroundColor Cyan
    $LineToWrite = "Room/Resource Delegate/Authorized User Changes Complete"
    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
    
    If ($Script:chkTemplate.Checked -eq $True)
    {
        Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\RROOPChanges.oft
    }
}