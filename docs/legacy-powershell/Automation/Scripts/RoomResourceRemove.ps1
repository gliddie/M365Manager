<####
#### Room and Resource Admin Menu
####
#
# Called by:  RoomResourceAdminMenu.ps1
#
#   Created  11/06/2024 - Adapted from the Rename Room/Resource and the original Remove RoomResource script
#>

function Build-DefaultForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Rename Conference Room or Resource/Equipment" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(600,200) #(W,H)
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
    $Script:txtTaskNo.Add_Click({
            $Global:okButton.Text = "Remove"
            $Global:OkButton.Visible = $True
        })

    ## Column Names
    $Script:lblTitles = New-Object System.Windows.Forms.Label   
        $Script:lblTitles.Text = "        Room Details"
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

    $Top = $Top + 30
    #Rooom Capacity
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
 
    $Top = $Top + 30
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
        $Script:chkTemplate.Text = "Display Email Template for Room/Equipment Removed" 
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
    $Global:form.Text = "Remove Conference Room or Resource/Equipment" 
    $Script:lblTaskNo.Visible = $False
    $Script:txtTaskNo.Visible = $False
    $Script:lblTitles.Visible = $False
    $Script:txtRoomName.Width = 300; $Script:lblRoomName.Top = 40; $Script:txtRoomName.Top = 40
    $Script:txtRoomName.ReadOnly = $False
    $Script:lblRoomAddr.Visible = $False
    $Script:txtRoomAddr.Visible = $False
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
    $Script:lblBldNo.Visible = $False
    $Script:txtBldNo.Visible =  $False
    $Script:lblFloor.Visible = $False
    $Script:txtFloor.Visible =  $False
    $Script:lblCapacity.Visible = $False
    $Script:txtCapacity.Visible =  $False
    $Global:form.Size = New-Object System.Drawing.Size(600,200) #(W,H)
    $Script:txtBldNo.Text = ""
    $Script:txtFloor.Text = ""
    $Script:txtCapacity.Text = ""

}

Function Complete-Form
{
    $Script:txtRoomName.Width = 250;$Script:lblRoomName.Top = 60; $Script:txtRoomName.Top = 60
    $Script:lblTaskNo.Visible = $True
    $Script:txtTaskNo.Visible = $True
    $Script:lblTitles.Visible = $True
    $Script:txtRoomName.ReadOnly = $True
    $Script:txtRoomName.Text = ($Script:txtRoomName.Text).Trim()
    $Script:txtRoom.Refresh
    $Global:form.Size = New-Object System.Drawing.Size(640,430) #(W,H)
    $Script:lblRoomAddr.Visible = $True
    $Script:txtRoomAddr.Visible = $True
    $Script:lblRoomDele.Visible = $True
    $Script:txtRoomDele.Visible = $True
    $Script:lblAlias.Visible = $True
    $Script:txtAlias.Visible = $True
    $Script:lblBldNo.Visible = $True
    $Script:txtBldNo.Visible = $True
    $Script:lblFloor.Visible = $True
    $Script:txtFloor.Visible = $True
    $Script:lblCapacity.Visible = $True
    $Script:txtCapacity.Visible = $True
    $Script:chkTemplate.Visible = $True
    $Script:lblRptFile.Visible = $True
    $Script:txtRptFile.Visible = $True
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
    $Script:RoomDetails = Get-Mailbox $Script:txtRoomName.Text

    If ($Script:RoomDetails.ResourceType -eq "Equipment")
    {
        $Global:form.Text = "Remove Equipment or Resouce"
        $Script:lblTitles.Text = "         Equipment Details"
    }

    If (($Script:RoomDetails.Office -like "Building*") -or ($Script:RoomDetails.Office -like "Bldg*"))
    {
        If ($Script:RoomDetails.Office -like "Building*")
        {
            $Script:txtBldNo.Text = (($Script:RoomDetails.Office).Substring(0,(($Script:RoomDetails.Office).IndexOf("-")-1)))
            $Script:txtBldNo.Text = $Script:txtBldNo.Text -replace("Building","")
        }
        $Script:txtBldNo.Text = ($Script:txtBldNo.Text).trim()
    }

    Build-BldgFloor
    $Script:txtAlias.Text = $Script:RoomDetails.Alias
    $Script:txtCapacity.Text = $Script:RoomDetails.ResourceCapacity
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
                $Script:txtFloor.Text = "Ground"
            }
            else
            {
                If ($Script:RoomDetails.Office -like "*-*")
                {
                    $hyph = (($Script:RoomDetails.Office).IndexOf("-")+1)
                    $Len = ($Script:RoomDetails.Office).Length
                    $Script:txtFloor.Text = ($Script:RoomDetails.Office).Substring(($hyph+1),($Len-($hyph+1)))
                    $Script:txtFloor.Text = $Script:txtFloor.Text -replace ("Floor ","")
                }
                else
                {
                    $Script:txtFloor.Text = $Script:RoomDetails.Office -replace ("Floor ","")
                }
            }
        }
        else
        {
            If ($Script:RoomDetails.Office -like "*Flr*")
            {
                $hyph = (($Script:RoomDetails.Office).IndexOf("-")+1)
                $Len = ($Script:RoomDetails.Office).Length
                $Script:txtFloor.Text = ($Script:RoomDetails.Office).Substring(($hyph+1),($Len-($hyph+1)))
                $Script:txtFloor.Text = $Script:txtFloor.Text -replace("FLR ","")
            }
        } 
    }
    $Script:txtFloor.Text = ($Script:txtFloor.Text).Trim()
}

Function Remove-Item
{
    Write-Output "Room or Resource Detailed Information" >> $ReportFile

    get-Mailbox $Script:txtRoomName.Text |fl >> $ReportFile

    Write-Output "Mailbox Message Statistics - Number of Messages in Mailbox" >> $ReportFile

    Get-MailboxStatistics $Script:txtRoomName.Text |ft >> $ReportFile

    Write-Output "Mailbox Message Folder Statistics - Number of Messages in Each Folder" >> $ReportFile

    Get-MailboxFolderStatistics $Script:txtRoomName.Text |ft Name,ItemsInFolder >> $ReportFile

    Write-Output "Mailbox Permissions" >> $ReportFile

    Get-MailboxPermission $Script:txtRoomName.Text |ft User,AccessRights >> $ReportFile

#    Write-Output "Room or Resource Calendar Notification Information" >> $ReportFile
#    get-CalendarNotification $Script:txtRoomName.Text >> $ReportFile
#    Get-EventsFromEmailConfiguration $mbx.PrimarySMTPAddress

    Write-Output "Room or Resource Calendar Processing Information" >> $ReportFile

    get-CalendarProcessing $Script:txtRoomName.Text |fl >> $ReportFile

    Write-Output "Room or Resource Calendar Configuration Information" >> $ReportFile
#    write-host "Ignore Warning about deprecated command the new command does not provide the calendar configuration"

    get-MailboxCalendarConfiguration $Script:txtRoomName.Text |fl >> $ReportFile

    Write-Output "Room or Resource Regional Configuration Information" >> $ReportFile

    get-MailboxRegionalConfiguration $Script:txtRoomName.Text |fl >> $ReportFile

    remove-Mailbox $Script:txtRoomName.Text -confirm:$False

    write-host "Report file written to: " $ReportFile
}

########################################

#Check-Reconnect
$Year = (get-date).ToString("yyyy")
$Path = "E:\Automation\RemoveRoomResource\Report\" + $Year
If (Test-Path $path) {} else {New-Item -Path $Path -ItemType Directory}
$ReportDirectory	= $Path + "\Report-RemoveRoomResource-"

$Global:form = ""
Build-DefaultForm
$Global:InputFocus = $Script:txtRoomName
Publish-Form

If ($Global:Result -eq "OK")
{
    Write-Output "Remove Room or Resource Script Started" >> $ReportFile
    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
    WriteReportEvent
    $LineToWrite = "Room Settings"
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
    }
    WriteReportEvent
    $LineToWrite = "   Calendar Time Zone         :" + "`t" + (Get-MailboxRegionalConfiguration $Script:txtRoomName.Text).TimeZone + "`n"
    WriteReportEvent
   
    Remove-Item               

    Write-Host "Room/Resource Removal Complete"
    $LineToWrite = "Room/Resource Removeal Process Complete"
    WriteReportEvent

    If ($Script:chkTemplate.Checked -eq $True)
    {
        Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\RoomResourceRemoval.oft
    }
}
else
{
    Write-Host "Room/Resource Rename Changes Cancelled" -ForegroundColor Red
}