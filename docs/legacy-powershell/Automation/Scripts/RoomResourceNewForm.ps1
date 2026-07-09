<#   
================================================================================ 
 Name: New Room/Resource Mailbox Form
 ================================================================================ 

 #  03/11/2025 - SAG - Changd the email address from @ul.com to @ul.onmicrosoft.com
#>  

function Build-RRMbxInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow"
    $Global:form.Text = "New Conference Room Mailbox"
    $Global:form.Size = New-Object System.Drawing.Size(450,380) #(W,H)
    If ($Global:chkNewGREquip.Checked -eq "Checked")
    {
        $Global:form.Text = "New Equipment/Resource Mailbox"
        $Global:form.Size = New-Object System.Drawing.Size(450,250) #(W,H)   
    }          
    $Global:form.StartPosition = "CenterScreen"

    $Top = 10
    ## Get Details used to create
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = $Top ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtInpTaskNo.Top = $Top; $Global:txtInpTaskNo.Left = 140; $Global:txtInpTaskNo.Width = 200;  
        $Global:txtInpTaskNo.Text = "TASK"   # Enter ticket number 
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form 
    
    $Top = $Top + 30
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label
        $Global:lblDispName.Text = "Room Name:"
        If ($Global:chkNewGREquip.Checked -eq "Checked")
        {   
            $Global:lblDispName.Text = "Equip/Resource Name:"
        }
        $Global:lblDispName.Top = $Top ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtInpName = New-Object Windows.Forms.TextBox  
        $Global:txtInpName.TabIndex = 0 # set Tab Order 
        $Global:txtInpName.Top = $Top; $Global:txtInpName.Left = 140; $Global:txtInpName.Width = 200;  
        $Global:txtInpName.Text = ""   # DisplayName
        $Global:form.Controls.Add($Global:txtInpName)    # Add to Form 
       # Obtain Value with: $Global:txtInpName.Text

    If ($Global:chkNewGRRoom.Checked -eq "Checked")
    {
        $Top = $Top + 30
        ## Building Name
        $Global:lblBldgName = New-Object System.Windows.Forms.Label   
            $Global:lblBldgName.Text = "Building Name/Number:"
            $Global:lblBldgName.Top = $Top ; $Global:lblBldgName.Left = 10; $Global:lblBldgName.Width=120 ;$Global:lblBldgName.AutoSize = $true 
            $Global:form.Controls.Add($Global:lblBldgName)    # Add to Form 
            # 
            $Global:txtBldgName = New-Object Windows.Forms.TextBox  
            $Global:txtBldgName.TabIndex = 0 # set Tab Order 
            $Global:txtBldgName.Top = $Top; $Global:txtBldgName.Left = 140; $Global:txtBldgName.Width = 200;  
            $Global:txtBldgName.Text = ""   # BuildingName
            $Global:txtBldgName.TabIndex = 0
            $Global:form.Controls.Add($Global:txtBldgName)    # Add to Form

        $Top = $Top + 30
        ## Floor Number
        $Global:lblFlrNo = New-Object System.Windows.Forms.Label   
            $Global:lblFlrNo.Text = "Floor No.:"
            $Global:lblFlrNo.Top = $Top ; $Global:lblFlrNo.Left = 10; $Global:lblFlrNo.Width=120 ;$Global:lblFlrNo.AutoSize = $true 
            $Global:form.Controls.Add($Global:lblFlrNo)    # Add to Form 
            # 
            $Global:txtFlrNo = New-Object Windows.Forms.TextBox  
            $Global:txtFlrNo.TabIndex = 0 # set Tab Order 
            $Global:txtFlrNo.Top = $Top; $Global:txtFlrNo.Left = 140; $Global:txtFlrNo.Width = 30;  
            $Global:txtFlrNo.Text = ""   # DisplayName
            $Global:txtFlrNo.TabIndex = 0
            $Global:form.Controls.Add($Global:txtFlrNo)    # Add to Form

        $Top = $Top + 30
        ## Room Capacity
        $Global:lblRoomCap = New-Object System.Windows.Forms.Label   
            $Global:lblRoomCap.Text = "*Room Capacity:"
            $Global:lblRoomCap.Top = $Top ; $Global:lblRoomCap.Left = 10; $Global:lblRoomCap.Width=120 ;$Global:lblRoomCap.AutoSize = $true 
            $Global:form.Controls.Add($Global:lblRoomCap)    # Add to Form 
            # 
            $Global:txtRoomCap = New-Object Windows.Forms.TextBox  
            $Global:txtRoomCap.TabIndex = 0 # set Tab Order 
            $Global:txtRoomCap.Top = $Top; $Global:txtRoomCap.Left = 140; $Global:txtRoomCap.Width = 30;  
            $Global:txtRoomCap.Text = ""   # DisplayName
            $Global:txtRoomCap.TabIndex = 0
            $Global:form.Controls.Add($Global:txtRoomCap)    # Add to Form
    }

    $Top = $Top + 30
    ## Standard Site Delegates
    $Global:chkStdDele = New-Object Windows.Forms.checkbox
        $Global:chkStdDele.Left = 140; $Global:chkStdDele.Width = 160; $Global:chkStdDele.Top = $Top
        $Global:chkStdDele.Text = "*Standrd Site Delegates" 
        $Global:chkStdDele.Checked = $true   # set a default value 
        $Global:chkStdDele.TabIndex = 2
        $Global:form.Controls.Add($Global:chkStdDele)

    $Top = $Top + 30
    ## General Use Room with Custom Delegates
    $Global:chkGenUseCustDele = New-Object Windows.Forms.checkbox 
        $Global:chkGenUseCustDele.Left = 140; $Global:chkGenUseCustDele.Width = 290; $Global:chkGenUseCustDele.Top = $Top
        $Global:chkGenUseCustDele.Text = "General Use Room with Custom Delegates"
        $Global:chkGenUseCustDele.Checked = $false   # set a default value 
        $Global:chkGenUseCustDele.TabIndex = 2
        $Global:form.Controls.Add($Global:chkGenUseCustDele)

    $Top = $Top + 20
    ## Restricted Room Delegates
    $Global:chkResDele = New-Object Windows.Forms.checkbox 
        $Global:chkResDele.Left = 140; $Global:chkResDele.Width = 190; $Global:chkResDele.Top = $Top
        $Global:chkResDele.Text = "Restricted Room Delegates"
        If ($Global:chkNewGREquip.Checked -eq "Checked")
        {
            $Global:chkResDele.Text = "Restricted Resource Delegates"
        }
        $Global:chkResDele.Checked = $false   # set a default value 
        $Global:chkResDele.TabIndex = 2
        $Global:form.Controls.Add($Global:chkResDele)

    $Top = $Top + 20
    ## Restricted Room Users
    $Global:chkResUsr = New-Object Windows.Forms.checkbox 
        $Global:chkResUsr.Left = 140; $Global:chkResUsr.Width = 350; $Global:chkResUsr.Top = $Top
        $Global:chkResUsr.Text = "Restricted Room Users"
        If ($Global:chkNewGREquip.Checked -eq "Checked")
        {
            $Global:chkResUsr.Text = "Restricted Resource Users"
        }         
        $Global:chkResUsr.Checked = $false   # set a default value 
        $Global:chkResUsr.TabIndex = 2
        $Global:form.Controls.Add($Global:chkResUsr)

    Add-FormStandardButtons
}

function Build-RRMbxDetailsForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow"
    If ($Global:chkNewGRRoom.Checked -eq "Checked")
    { 
        $Global:form.Text = "New Conferemce Room" 
    }
    else
    {
        $Global:form.Text = "New Equipment/Resource" 
    }
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(450,400) #(W,H)

    $Top = 10
    ## Get Details to complete creation
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = $Top; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=150 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtInpTaskNo.Top = $Top; $Global:txtInpTaskNo.Left = 140; $Global:txtInpTaskNo.Width = 220; 
        $Global:txtInpTaskNo.Text = $Global:TaskNo 
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form
    $Top = $Top + 30

    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label
        $Global:lblDispName.Text = "Room Name:"
        If ($Global:chkNewGREquip.Checked -eq "Checked")
        {
            $Global:lblDispName.Text = "Resource Name:"
        }
        $Global:lblDispName.Top = $Top; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=150 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtDispName = New-Object Windows.Forms.TextBox  
        $Global:txtDispName.TabIndex = 0 # set Tab Order 
        $Global:txtDispName.Top = $Top; $Global:txtDispName.Left = 140; $Global:txtDispName.Width = 220;  
        $Global:txtDispName.Text = $Global:DispName   # DisplayName
        $Global:form.Controls.Add($Global:txtDispName)    # Add to Form 
    $Top = $Top + 30

    ## EmailAddress
    $Global:lblMbxAddr = New-Object System.Windows.Forms.Label   
        $Global:lblMbxAddr.Text = "Email Address:"  
        $Global:lblMbxAddr.Top = $Top; $Global:lblMbxAddr.Left = 10; $Global:lblMbxAddr.Width=120 ;$Global:lblMbxAddr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblMbxAddr)    # Add to Form 
        # 
        $Global:txtMbxAddr = New-Object Windows.Forms.TextBox  
        $Global:txtMbxAddr.TabIndex = 0 # set Tab Order 
        $Global:txtMbxAddr.Top = $Top; $Global:txtMbxAddr.Left = 140; $Global:txtMbxAddr.Width = 220;
        $Global:txtMbxAddr.Text = $Global:RoomAddr   # Email Address
        $Global:form.Controls.Add($Global:txtMbxAddr)    # Add to Form
    $Top = $Top + 30

    $Global:RoomGrp = ""
    If ($Global:chkNewGRRoom.Checked -eq "Checked")
    {
        ## Rooms Group
        $Global:lblRoomGrp = New-Object System.Windows.Forms.Label   
            $Global:lblRoomGrp.Text = "Room List Name:"  
            $Global:lblRoomGrp.Top = $Top; $Global:lblRoomGrp.Left = 10; $Global:lblRoomGrp.Width=120 ; $Global:lblRoomGrp.AutoSize = $true
            $Global:form.Controls.Add($Global:lblRoomGrp)    # Add to Form 
            # 
            $Global:txtInpRoomGrp = New-Object Windows.Forms.TextBox  
            $Global:txtInpRoomGrp.TabIndex = 0 # set Tab Order 
            $Global:txtInpRoomGrp.Top = $Top; $Global:txtInpRoomGrp.Left = 140; $Global:txtInpRoomGrp.Width = 220;  
            $Global:txtInpRoomGrp.Text = $Global:RoomGrp   # Enter ticket number
            If (($Global:chkStdDele.Checked -eq "Checked") -or ($Global:chkGenUseCustDele.Checked -eq "Checked"))
            {  
                #Standard Site Group
                $Global:txtInpRoomGrp.Text = ($Global:DispName.Substring(0,3)) + " Conference Rooms"
            }
            else
            {
                #Restricted Room Group
                $Global:txtInpRoomGrp.Text = ($Global:DispName.Substring(0,3)) + " Restricted Rooms"
            }
            $Global:form.Controls.Add($Global:txtInpRoomGrp)    # Add to Form
            $Top = $Top + 30
    }

    ## Delegates
    $Global:lblDeleName = New-Object System.Windows.Forms.Label
        $Global:lblDeleName.Top = $Top ; $Global:lblDeleName.Left = 10; $Global:lblDeleName.Width=150; $Global:lblDeleName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDeleName)    # Add to Form 
        #
        $Global:txtDeLeName = New-Object Windows.Forms.TextBox  
        $Global:txtDeleName.TabIndex = 0 # set Tab Order 
        $Global:txtDeleName.Top = $Top; $Global:txtDeleName.Left = 140; $Global:txtDeleName.Width = 220;
    If ($Global:chkStdDele.Checked -eq "Checked")
    {
        #Standard Site Delegates
        $Global:lblDeleName.Text = "Site Delegates:"
        $Global:txtDeleName.Text = ("MBX." + ($Global:DispName.Substring(0,3)) + ".RRS.OutOfPolicy.DE")
    }
    else
    {
        #Restricted Room Delegates or Custom General User Room Delegates
        $Global:lblDeleName.Text = "Room Delegates Group:"
        $Global:txtDeleName.Text = ("MBX." + ($Global:DispName.Substring(0,3)) + ".RRS.OutOfPolicy." + ($Global:DispName.substring(4,$Global:DispName.IndexOf(" ")+1)) + ".DE")
    }
    $Global:form.Controls.Add($Global:txtDeleName)    # Add to Form
    $Top = $Top + 30

    ## Room Delegate Info
    $Exists = [bool](Get-DistributionGroup $Global:txtDeleName.Text -ErrorAction SilentlyContinue)
    if ($Exists -eq $False)
    {
        $Global:lblDeleMgrs = New-Object System.Windows.Forms.Label
        $Global:lblDeleMgrs.Text = "Delegate Emp Nos.:" 
        $Global:lblDeleMgrs.Top = $Top ; $Global:lblDeleMgrs.Left = 10; $Global:lblDeleMgrs.Width=150; $Global:lblDeleMgrs.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDeleMgrs)    # Add to Form 
        #
        $Global:txtDeLeMgrs = New-Object Windows.Forms.TextBox  
        $Global:txtDeleMgrs.TabIndex = 0 # set Tab Order 
        $Global:txtDeleMgrs.Top = $Top; $Global:txtDeleMgrs.Left = 140; $Global:txtDeleMgrs.Width = 220;
        $Global:form.Controls.Add($Global:txtDeleMgrs)    # Add to Form
        $Top = $Top + 30
    }

    ## Restricted Users
    If ($Global:chkResUsr.Checked -eq "Checked")
    {
        $Global:lblResUsr = New-Object System.Windows.Forms.Label
        $Global:lblResUsr.Text = "Room Users Group:" 
        $Global:lblResUsr.Top = $Top ; $Global:lblResUsr.Left = 10; $Global:lblResUsr.Width=150; $Global:lblResUsr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblResUsr)    # Add to Form 
        #
        $Global:txtResUsr = New-Object Windows.Forms.TextBox  
        $Global:txtResUsr.TabIndex = 0 # set Tab Order 
        $Global:txtResUsr.Top = $Top; $Global:txtResUsr.Left = 140; $Global:txtResUsr.Width = 220;
        $Global:txtResUsr.Text = ("MBX." + ($Global:DispName.Substring(0,3)) + ".RRS.OutofPolicy." + ($Global:DispName.substring(4,$Global:DispName.IndexOf(" ")+1)) + ".US")
        $Global:form.Controls.Add($Global:txtResUsr)    # Add to Form
        $Top = $Top + 30
    }

    If ($Global:chkResUsr.Checked -eq "Checked")
    {
    ## Restricted Users Info
        $Exists = [bool](Get-DistributionGroup $Global:txtResUsr.Text -ErrorAction SilentlyContinue)
        if ($Exists -eq $False)
        {
            $Global:lblResUsrInfo = New-Object System.Windows.Forms.Label
            $Global:lblResUsrInfo.Text = "Users Emp Nos.:" 
            $Global:lblResUsrInfo.Top = $Top ; $Global:lblResUsrInfo.Left = 10; $Global:lblResUsrInfo.Width=150; $Global:lblResUsrInfo.AutoSize = $true 
            $Global:form.Controls.Add($Global:lblResUsrInfo)    # Add to Form 
            #
            $Global:txtResUsrInfo = New-Object Windows.Forms.TextBox  
            $Global:txtResUsrInfo.TabIndex = 0 # set Tab Order 
            $Global:txtResUsrInfo.Top = $Top; $Global:txtResUsrInfo.Left = 140; $Global:txtResUsrInfo.Width = 220;
            $Global:form.Controls.Add($Global:txtResUsrInfo)    # Add to Form
            $Top = $Top + 30
        }
    }

    ## TimeZone
    $Global:lblTimeZone = New-Object System.Windows.Forms.Label
        $Global:lblTimeZone.Text = "Room Time Zone:" 
        $Global:lblTimeZone.Top = $Top ; $Global:lblTimeZone.Left = 10; $Global:lblTimeZone.Width=150; $Global:lblTimeZone.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblTimeZone)    # Add to Form 
        #
        $Global:txtTimeZone = New-Object Windows.Forms.TextBox  
        $Global:txtTimeZone.TabIndex = 0 # set Tab Order 
        $Global:txtTimeZone.Top = $Top; $Global:txtTimeZone.Left = 140; $Global:txtTimeZone.Width = 220;
        $Global:txtTimeZone.Text = $Global:TimeZone
        $Global:form.Controls.Add($Global:txtTimeZone)    # Add to Form
    $Top = $Top + 30

    ## Regional Managers
    $Global:lblRegMgr = New-Object System.Windows.Forms.Label
        $Global:lblRegMgr.Text = "Regional Managers:" 
        $Global:lblRegMgr.Top = $Top ; $Global:lblRegMgr.Left = 10; $Global:lblRegMgr.Width=150; $Global:lblRegMgr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblRegMgr)    # Add to Form 
        #
        $Global:txtRegMgr = New-Object Windows.Forms.TextBox  
        $Global:txtRegMgr.TabIndex = 0 # set Tab Order 
        $Global:txtRegMgr.Top = $Top; $Global:txtRegMgr.Left = 140; $Global:txtRegMgr.Width = 220;
        $Global:txtRegMgr.Text = $Global:RegMgr
        $Global:form.Controls.Add($Global:txtRegMgr)    # Add to Form

    Add-FormStandardButtons
}

function Get-NewSiteInfo
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "New Site Details" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(450,380) #(W,H)

    ## Get Details used to create
    ## 3 Letter Site Code
    $Global:lblOfficeCode = New-Object System.Windows.Forms.Label   
        $Global:lblOfficeCode.Text = "Site Code:"  
        $Global:lblOfficeCode.Top = 10 ; $Global:lblOfficeCode.Left = 10; $Global:lblOfficeCode.Width=120 ; $Global:lblOfficeCode.AutoSize = $true
        $Global:form.Controls.Add($Global:lblOfficeCode)    # Add to Form 
        # 
        $Global:txtInpOfficeCode = New-Object Windows.Forms.TextBox  
        $Global:txtInpOfficeCode.TabIndex = 0 # set Tab Order 
        $Global:txtInpOfficeCode.Top = 10; $Global:txtInpOfficeCode.Left = 140; $Global:txtInpOfficeCode.Width = 200;  
        $Global:txtInpOfficeCode.Text = $Global:MbxLoc.ToUpper($Global:MbxLoc)
        $Global:form.Controls.Add($Global:txtInpOfficeCode)   # Add to Form 

    ## Time Zone
    $Global:lblNewTimeZone = New-Object System.Windows.Forms.Label   
        $Global:lblNewTimeZone.Text = "Select Time Zone:"
        $Global:lblNewTimeZone.Top = 40 ; $Global:lblNewTimeZone.Left = 10; $Global:lblNewTimeZone.Width=120 ;$Global:lblNewTimeZone.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblNewTimeZone)    # Add to Form 
        # 
        $Global:txtNewTimeZone = New-Object Windows.Forms.ListBox  
        $Global:txtNewTimeZone.TabIndex = 0 # set Tab Order 
        $Global:txtNewTimeZone.Top = 40; $Global:txtNewTimeZone.Left = 140; $Global:txtNewTimeZone.Width = 200; $Global:txtNewTimeZone.Height = 70
        $Zones = import-csv "e:\O365AdminShared\Data\UniqueTimeZones.txt"
        Foreach ($Zones in $Zones)
        {
            [void] $Global:txtNewTimeZone.Items.Add($Zones.TimeZone)
        }
        $Global:form.Controls.Add($Global:txtNewTimeZone)    # Add to Form 
        $Global:form.Topmost = $true

    ## Regional Location
    $Global:lblRegionLoc = New-Object System.Windows.Forms.Label   
        $Global:lblRegionLoc.Text = "Select Region:"
        $Global:lblRegionLoc.Top = 130 ; $Global:lblRegionLoc.Left = 10; $Global:lblRegionLoc.Width=120 ;$Global:lblRegionLoc.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblRegionLoc)    # Add to Form 
        # 
        $Global:txtRegionLoc = New-Object System.Windows.Forms.ListBox
        $Global:txtRegionLoc.TabIndex = 0 # set Tab Order 
        $Global:txtRegionLoc.Top = 130; $Global:txtRegionLoc.Left = 140; $Global:txtRegionLoc.Width = 200; $Global:txtRegionLoc.Height = 70;
        [void] $Global:txtRegionLoc.Items.Add('AP')
        [void] $Global:txtRegionLoc.Items.Add('CA')
        [void] $Global:txtRegionLoc.Items.Add('EU')
        [void] $Global:txtRegionLoc.Items.Add('LA')
        [void] $Global:txtRegionLoc.Items.Add('US')
        $Global:txtRegionLoc.TabIndex = 0
        $Global:form.Controls.Add($Global:txtRegionLoc)    # Add to Form
        $Global:form.Topmost = $true

    Add-FormStandardButtons
}

function Build-MbxName
{
    #Adjust name to proper case ann replace acronyms
    $Global:Name = (Get-Culture).textinfo.totitlecase($Global:txtInpName.Text.ToLower())
    $Space = $Global:Name.IndexOf(" ")
    $NameLen = $Global:Name.Length
    $Global:MbxLoc = $Global:Name.Substring(0,$Global:Name.IndexOf(" "))
    $MbxName = $Global:Name.Substring($Space+1,($NameLen-($Space+1)))
    $RoomName = ((Get-Culture).textinfo.totitlecase($MbxName.ToLower())) -replace(" ","")
    $Global:RoomGrp = $Global:MbxLoc + "Conference Rooms"
    
    foreach ($Acro in $Acro)
    {
        $MbxName = $MbxName -Replace($Acro.Acronym,$Acro.Translation)
    }

    If ($Global:chkNewGRRoom.Checked -eq "Checked")
    {
        $Global:RoomAddr = $Global:MbxLoc.ToUpper()
    }
    else
    {
        $Global:RoomAddr = "RES." + $Global:MbxLoc.ToUpper()
    }

    If ($Global:txtBldgName.Text.Length -ne 0)
    {
        $Global:RoomAddr = $Global:RoomAddr + "Bldg" + $Global:txtBldgName.Text + "."
    }
    else
    {
        $Global:RoomAddr = $Global:RoomAddr + "."
    }
        
    If ($Global:txtFlrNo.Text.Length -ne 0)
    {
        If ($Global:txtFlrNo.Text -eq "Ground")
        {
            $Global:RoomAddr = $Global:RoomAddr + "FLRGrnd"
        }
        else
        {
            $Global:RoomAddr = $Global:RoomAddr + "FLR" + $Global:txtFlrNo.Text
        }
    }
    
    If ($Global:chkNewGREquip.Checked -eq "Checked")
    {
        $Global:RoomAddr = ($Global:RoomAddr + "." + $RoomName + "@ul.onmicrosoft.com").Replace("..",".")
    }
    else
    {
        $Global:RoomAddr = ($Global:RoomAddr + "." + $RoomName + "." + $Global:txtRoomCap.Text + "@ul.onmicrosoft.com").Replace("..",".")        
    }

    $Global:DispName = ($Global:MbxLoc).ToUpper() + " "  + $MbxName
}

$Global:OKDetails = ""
$Global:form = ""
$Acro = Import-Csv e:\O365AdminShared\Data\KnownAcronyms.csv

Add-Type -Assembly System.Windows.Forms     ## Load the Windows Forms assembly 
## Create the main form

$Global:OKDetails = "Continue"
$Global:txtRoomCap = ""
$Global:TimeZone = ""
$Global:RegMgr = ""
$siteDet = import-csv "e:\O365AdminShared\Data\RoomTimeZones.csv"
Build-RRMbxInputForm
#Publish-Form

If($Global:chkNewGRRoom.Checked -eq "Checked")
{
    Do
    {
        Publish-Form
    } while (($Global:txtRoomCap.Text.Length -eq 0) -and ($Global:Result -eq "OK"))
}
else
{
    Publish-Form
}

If ($Global:Result -eq "OK")
{
    #Adjust name to proper case ann replace acronyms
    $Global:Name = (Get-Culture).textinfo.totitlecase($Global:txtInpName.Text.ToLower())
    $Global:Addr = ((Get-Culture).textinfo.totitlecase($Global:txtInpName.Text)).replace(" ","")
    $Space = $Global:Name.IndexOf(" ")
    $NameLen = $Global:Name.Length
    $Global:MbxLoc = $Global:Name.Substring(0,$Global:Name.IndexOf(" "))
    $MbxName = $Global:Name.Substring($Space+1,($NameLen-($Space+1)))
    $MbxName = ((Get-Culture).textinfo.totitlecase($MbxName.ToLower()))
    $Global:TaskNo = $Global:txtInpTaskNo.Text

    $BldDetails = "Y"
    Build-MbxName

    foreach ($SiteDet in $SiteDet)
    {
        If ($SiteDet.Code -eq $Global:MbxLoc)
        {
            $Global:TimeZone = $SiteDet.Zone
            $Global:RegMgr = $SiteDet.Admins
            $Found = "Y"
        }
    }

    If ($Global:TimeZone -eq "")
    {
        Get-NewSiteInfo
        Publish-Form

        If ($Global:Result -eq "OK")
        {
            $Global:RegMgr = "MBX." + $Global:txtRegionLoc.SelectedItem + ".RRS.Admins"

            If ($Global:txtNewTimeZone.SelectedItem -ne "Not Listed" )
            {
                $Global:TimeZone = $Global:txtNewTimeZone.SelectedItem
            }
            else
            {
                $Global:form = New-Object Windows.Forms.Form 
                $Global:form.FormBorderStyle = "FixedToolWindow" 
                $Global:form.Text = "New TimeZone" 
                $Global:form.StartPosition = "CenterScreen"
                $Global:form.Size = New-Object System.Drawing.Size(450,150) #(W,H)

                ## Get Details used to create
                ## 3 Letter Site Code
                $Global:lblNewZone = New-Object System.Windows.Forms.Label   
                $Global:lblOfficeCode.Text = "Time Zone:"  
                $Global:lblNewZone.Top = 30 ; $Global:lblNewZone.Left = 10; $Global:lblNewZone.Width=120 ; $Global:lblNewZone.AutoSize = $true
                $Global:form.Controls.Add($Global:lblNewZone)    # Add to Form 
                # 
                $Global:txtInpNewZone = New-Object Windows.Forms.TextBox  
                $Global:txtInpNewZone.TabIndex = 0 # set Tab Order 
                $Global:txtInpNewZone.Top = 30; $Global:txtInpNewZone.Left = 140; $Global:txtInpNewZone.Width = 200;  
                $Global:txtInpNewZone.Text = ""
                $Global:form.Controls.Add($Global:txtInpNewZone)   # Add to Form

                Add-FormStandardButtons
                Publish-Form
                $Global:TimeZone = $Global:txtInpNewZone.Text
            }

            #Add new location to the RoomTimeZone file
            $LineToWrite = "{0},""{1}"",{2}" -f ($Global:MbxLoc.ToUpper($Global:MbxLoc)),$Global:TimeZone,$Global:RegMgr
            Out-File -FilePath "E:\O365AdminShared\Data\RoomTimeZones.csv" -InputObject $LineToWrite -Append
            write-host "Please inform Enterprise Messaging Services Teams so this new Site Location/TimeZone can be added to the pick list" -ForegroundColor Red
        }
    }

    If ($Global:Result -eq "OK")
    {
        $Global:TaskNo = $Global:txtInpTaskNo.Text
        $Global:OKDetails = "Create"
        Build-RRMbxDetailsForm
        Publish-Form
    }
}

If ($Global:Result -eq "OK")
{
    $uDate = get-date -uformat %D
    $uTime = get-date -uformat %T
    $Date  = $uDate.Replace("/", "-")
	$Time  = $uTime.Replace(":", "")
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name
    $RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"

	$atMAIL   = $Global:txtMbxAddr.Text.indexOf("@")
    $LeftName = $Global:txtMbxAddr.Text.substring(0,$atMAIL)
    
    If ($Global:chkStdDele.Checked -eq "Checked")
    {
        $ReportFile	= "E:\Automation\NewConferenceRoom\Report\Report-NewConferenceRoom-Date" + $Date + "Time" + $Time + ".log"
        If ($Global:chkNewGREquip.Checked -eq "Checked")
        {
            $ReportFile	= "E:\Automation\NewEquipment\Report\Report-NewEquipment-Date" + $Date + "Time" + $Time + ".log"
        }
    }
    else
    {
        $ReportFile	= "E:\Automation\NewConferenceRoom\Report\Report-NewRestrictedConferenceRoom-Date" + $Date + "Time" + $Time + ".log"
        If ($Global:chkNewGREquip.Checked -eq "Checked")
        {
            $ReportFile	= "E:\Automation\NewEquipment\Report\Report-NewRestrictedEquipment-Date" + $Date + "Time" + $Time + ".log"
        }
    }

    $LineToWrite = "START" + "`t" + "NewRoomEquipment script has started"
 	WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
    WriteReportEvent  

    If (Get-Mailbox $Global:txtDispName.Text -ErrorAction SilentlyContinue)
    {
        $Output = $wshell.Popup("This room/resource already exists",0,"Room/Resource Exists",0+32)
        $LineToWrite = "STAR" + "`t" + "Room/Resource Already Exists: " + $Global:txtDispName.Text
    	WriteReportEvent
    }
    else
    {
        Invoke-Expression -Command e:\Automation\Scripts\RoomEquipNew.ps1
    }

#  Check to see if the Delegates and Users Groups exist

    Invoke-Expression -Command e:\Automation\Scripts\RoomEquipGroups

## Configure Room

    Invoke-Expression -Command e:\Automation\Scripts\RoomEquipConfig.ps1

    If ($Global:chkNewGRRoom.Checked -eq "Checked")
    {
        write-host "Room creation complete"
        $Output = $wshell.Popup("Room " + $Global:txtDispName.Text + " creation complete.",0,"Room Created",0+32)
    }
    else
    {
        write-host "Resource creation complete"
        $Output = $wshell.Popup("Resource " + $Global:txtDispName.Text + " creation complete.",0,"Resource Created",0+32)
    }

    If ($Global:NewRoomList -eq "Y")
    {
        If (($Global:chkStdDele.Checked -eq "Checked") -or($Global:chkGenUseCustDele.Checked -eq "Checked"))
        {
            Invoke-Expression -Command E:\O365AdminShared\EmailTemplates\RoomorResourceNewSite.oft
        }
        else
        {
            Invoke-Expression -Command E:\O365AdminShared\EmailTemplates\RestrictedRoomsNewSite.oft
        }
    }
    else
    {
        If (($Global:chkStdDele.Checked -eq "Checked") -or($Global:chkGenUseCustDele.Checked -eq "Checked"))
        {
            Invoke-Expression -Command E:\O365AdminShared\EmailTemplates\RoomorResourceAdditions.oft
        }
        else
        {
            Invoke-Expression -Command E:\O365AdminShared\EmailTemplates\RestrictedRoomAdditions.oft
        }
    }
}
else
{
    If ($Global:chkNewGRRoom.Checked -eq "Checked")
    {
        write-host "Room creation cancelled"
        $Output = $wshell.Popup("Room " + $Global:txtDispName.Text + " creation cancelled.",0,"Room Cancelled",0+32)
    }
    else
    {
        write-host "Resource creation cancelled"
        $Output = $wshell.Popup("Resource " + $Global:txtDispName.Text + " creation cancelled.",0,"Resource Cancelled",0+32)
    }
}