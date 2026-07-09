<#   
================================================================================ 
 Name: New Room/Resource Mailbox Form
 ================================================================================ 
#>  

function Build-RRMbxInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow"
    $Global:form.Text = "New Conference Room Mailbox"
    $Global:form.Size = New-Object System.Drawing.Size(450,380) #(W,H)
    If ($Script:txtNewGREquip.Checked -eq "Checked")
    {
        $Global:form.Text = "New Equipment/Resource Mailbox"
        $Global:form.Size = New-Object System.Drawing.Size(450,250) #(W,H)   
    }          
    $Global:form.StartPosition = "CenterScreen"

    $Top = 10
    ## Get Details used to create
    ## Ticket Number
    $Script:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Script:lblTaskNo.Text = "Ticket Number:"  
        $Script:lblTaskNo.Top = $Top ; $Script:lblTaskNo.Left = 10; $Script:lblTaskNo.Width=120 ; $Script:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Script:lblTaskNo)    # Add to Form 
        # 
        $Script:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Script:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Script:txtInpTaskNo.Top = $Top; $Script:txtInpTaskNo.Left = 140; $Script:txtInpTaskNo.Width = 200;  
        $Script:txtInpTaskNo.Text = "TASK"   # Enter ticket number 
        $Global:form.Controls.Add($Script:txtInpTaskNo)    # Add to Form 

    
    $Top = $Top + 30
    ## Display Name
    $Script:lblDispName = New-Object System.Windows.Forms.Label
        $Script:lblDispName.Text = "Room Name:"
        If ($Script:txtNewGREquip.Checked -eq "Checked")
        {   
            $Script:lblDispName.Text = "Equip/Resource Name:"
        }
        $Script:lblDispName.Top = $Top ; $Script:lblDispName.Left = 10; $Script:lblDispName.Width=120 ;$Script:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblDispName)    # Add to Form 
        # 
        $Script:txtInpName = New-Object Windows.Forms.TextBox  
        $Script:txtInpName.TabIndex = 0 # set Tab Order 
        $Script:txtInpName.Top = $Top; $Script:txtInpName.Left = 140; $Script:txtInpName.Width = 200;  
        $Script:txtInpName.Text = ""   # DisplayName
        $Global:form.Controls.Add($Script:txtInpName)    # Add to Form 
       # Obtain Value with: $Script:txtInpName.Text

    If ($Script:txtNewFloor.Checked -eq "Checked")
    {
        $Top = $Top + 30
        ## Building Name
        $Script:lblBldgName = New-Object System.Windows.Forms.Label   
            $Script:lblBldgName.Text = "Building Name/Number:"
            $Script:lblBldgName.Top = $Top ; $Script:lblBldgName.Left = 10; $Script:lblBldgName.Width=120 ;$Script:lblBldgName.AutoSize = $true 
            $Global:form.Controls.Add($Script:lblBldgName)    # Add to Form 
            # 
            $Script:txtBldgName = New-Object Windows.Forms.TextBox  
            $Script:txtBldgName.TabIndex = 0 # set Tab Order 
            $Script:txtBldgName.Top = $Top; $Script:txtBldgName.Left = 140; $Script:txtBldgName.Width = 200;  
            $Script:txtBldgName.Text = ""   # BuildingName
            $Script:txtBldgName.TabIndex = 0
            $Global:form.Controls.Add($Script:txtBldgName)    # Add to Form

        $Top = $Top + 30
        ## Floor Number
        $Script:lblFlrNo = New-Object System.Windows.Forms.Label   
            $Script:lblFlrNo.Text = "Floor No.:"
            $Script:lblFlrNo.Top = $Top ; $Script:lblFlrNo.Left = 10; $Script:lblFlrNo.Width=120 ;$Script:lblFlrNo.AutoSize = $true 
            $Global:form.Controls.Add($Script:lblFlrNo)    # Add to Form 
            # 
            $Script:txtFlrNo = New-Object Windows.Forms.TextBox  
            $Script:txtFlrNo.TabIndex = 0 # set Tab Order 
            $Script:txtFlrNo.Top = $Top; $Script:txtFlrNo.Left = 140; $Script:txtFlrNo.Width = 30;  
            $Script:txtFlrNo.Text = ""   # DisplayName
            $Script:txtFlrNo.TabIndex = 0
            $Global:form.Controls.Add($Script:txtFlrNo)    # Add to Form

        $Top = $Top + 30
        ## Room Capacity
        $Script:lblRoomCap = New-Object System.Windows.Forms.Label   
            $Script:lblRoomCap.Text = "*Room Capacity:"
            $Script:lblRoomCap.Top = $Top ; $Script:lblRoomCap.Left = 10; $Script:lblRoomCap.Width=120 ;$Script:lblRoomCap.AutoSize = $true 
            $Global:form.Controls.Add($Script:lblRoomCap)    # Add to Form 
            # 
            $Script:txtRoomCap = New-Object Windows.Forms.TextBox  
            $Script:txtRoomCap.TabIndex = 0 # set Tab Order 
            $Script:txtRoomCap.Top = $Top; $Script:txtRoomCap.Left = 140; $Script:txtRoomCap.Width = 30;  
            $Script:txtRoomCap.Text = ""   # DisplayName
            $Script:txtRoomCap.TabIndex = 0
            $Global:form.Controls.Add($Script:txtRoomCap)    # Add to Form
    }

    $Top = $Top + 30
    ## Standard Site Delegates
    $Script:txtStdDele = New-Object Windows.Forms.checkbox
        $Script:txtStdDele.Left = 140; $Script:txtStdDele.Width = 160; $Script:txtStdDele.Top = $Top
        $Script:txtStdDele.Text = "*Standrd Site Delegates" 
        $Script:txtStdDele.Checked = $true   # set a default value 
        $Script:txtStdDele.TabIndex = 2
        $Global:form.Controls.Add($Script:txtStdDele)

    $Top = $Top + 30
    ## General Use Room with Custom Delegates
    $Script:txtGenUseCustDele = New-Object Windows.Forms.checkbox 
        $Script:txtGenUseCustDele.Left = 140; $Script:txtGenUseCustDele.Width = 290; $Script:txtGenUseCustDele.Top = $Top
        $Script:txtGenUseCustDele.Text = "General Use Room with Custom Delegates"
        $Script:txtGenUseCustDele.Checked = $false   # set a default value 
        $Script:txtGenUseCustDele.TabIndex = 2
        $Global:form.Controls.Add($Script:txtGenUseCustDele)

    $Top = $Top + 20
    ## Restricted Room Delegates
    $Script:txtResDele = New-Object Windows.Forms.checkbox 
        $Script:txtResDele.Left = 140; $Script:txtResDele.Width = 190; $Script:txtResDele.Top = $Top
        $Script:txtResDele.Text = "Restricted Room Delegates"
        If ($Script:txtNewGREquip.Checked -eq "Checked")
        {
            $Script:txtResDele.Text = "Restricted Resource Delegates"
        }
        $Script:txtResDele.Checked = $false   # set a default value 
        $Script:txtResDele.TabIndex = 2
        $Global:form.Controls.Add($Script:txtResDele)

    $Top = $Top + 20
    ## Restricted Room Users
    $Script:txtResUsr = New-Object Windows.Forms.checkbox 
        $Script:txtResUsr.Left = 140; $Script:txtResUsr.Width = 350; $Script:txtResUsr.Top = $Top
        $Script:txtResUsr.Text = "Restricted Room Users"
        If ($Script:txtNewGREquip.Checked -eq "Checked")
        {
            $Script:txtResUsr.Text = "Restricted Resource Users"
        }         
        $Script:txtResUsr.Checked = $false   # set a default value 
        $Script:txtResUsr.TabIndex = 2
        $Global:form.Controls.Add($Script:txtResUsr)

    Add-FormStandardButtons
}

function Build-RRMbxDetailsForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow"
    If ($Script:txtNewFloor.Checked -eq "Checked")
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
    $Script:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Script:lblTaskNo.Text = "Ticket Number:"  
        $Script:lblTaskNo.Top = $Top; $Script:lblTaskNo.Left = 10; $Script:lblTaskNo.Width=150 ; $Script:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Script:lblTaskNo)    # Add to Form 
        # 
        $Script:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Script:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Script:txtInpTaskNo.Top = $Top; $Script:txtInpTaskNo.Left = 140; $Script:txtInpTaskNo.Width = 220; 
        $Script:txtInpTaskNo.Text = $Script:TaskNo 
        $Global:form.Controls.Add($Script:txtInpTaskNo)    # Add to Form
    $Top = $Top + 30

    ## Display Name
    $Script:lblDispName = New-Object System.Windows.Forms.Label
        $Script:lblDispName.Text = "Room Name:"
        If ($Script:txtNewGREquip.Checked -eq "Checked")
        {
            $Script:lblDispName.Text = "Resource Name:"
        }
        $Script:lblDispName.Top = $Top; $Script:lblDispName.Left = 10; $Script:lblDispName.Width=150 ;$Script:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblDispName)    # Add to Form 
        # 
        $Script:txtDispName = New-Object Windows.Forms.TextBox  
        $Script:txtDispName.TabIndex = 0 # set Tab Order 
        $Script:txtDispName.Top = $Top; $Script:txtDispName.Left = 140; $Script:txtDispName.Width = 220;  
        $Script:txtDispName.Text = $Script:DispName   # DisplayName
        $Global:form.Controls.Add($Script:txtDispName)    # Add to Form 
    $Top = $Top + 30

    ## EmailAddress
    $Script:lblMbxAddr = New-Object System.Windows.Forms.Label   
        $Script:lblMbxAddr.Text = "Email Address:"  
        $Script:lblMbxAddr.Top = $Top; $Script:lblMbxAddr.Left = 10; $Script:lblMbxAddr.Width=120 ;$Script:lblMbxAddr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblMbxAddr)    # Add to Form 
        # 
        $Script:txtMbxAddr = New-Object Windows.Forms.TextBox  
        $Script:txtMbxAddr.TabIndex = 0 # set Tab Order 
        $Script:txtMbxAddr.Top = $Top; $Script:txtMbxAddr.Left = 140; $Script:txtMbxAddr.Width = 220;
        $Script:txtMbxAddr.Text = $Script:RoomAddr   # Email Address
        $Global:form.Controls.Add($Script:txtMbxAddr)    # Add to Form
    $Top = $Top + 30

    $Script:RoomGrp = ""
    If ($Script:txtNewFloor.Checked -eq "Checked")
    {
        ## Rooms Group
        $Script:lblRoomGrp = New-Object System.Windows.Forms.Label   
            $Script:lblRoomGrp.Text = "Room List Name:"  
            $Script:lblRoomGrp.Top = $Top; $Script:lblRoomGrp.Left = 10; $Script:lblRoomGrp.Width=120 ; $Script:lblRoomGrp.AutoSize = $true
            $Global:form.Controls.Add($Script:lblRoomGrp)    # Add to Form 
            # 
            $Script:txtInpRoomGrp = New-Object Windows.Forms.TextBox  
            $Script:txtInpRoomGrp.TabIndex = 0 # set Tab Order 
            $Script:txtInpRoomGrp.Top = $Top; $Script:txtInpRoomGrp.Left = 140; $Script:txtInpRoomGrp.Width = 220;  
            $Script:txtInpRoomGrp.Text = $Script:RoomGrp   # Enter ticket number
            If (($Script:txtStdDele.Checked -eq "Checked") -or ($Script:txtGenUseCustDele.Checked -eq "Checked"))
            {  
                #Standard Site Group
                $Script:txtInpRoomGrp.Text = ($Script:DispName.Substring(0,3)) + " Conference Rooms"
            }
            else
            {
                #Restricted Room Group
                $Script:txtInpRoomGrp.Text = ($Script:DispName.Substring(0,3)) + " Restricted Rooms"
            }
            $Global:form.Controls.Add($Script:txtInpRoomGrp)    # Add to Form
            $Top = $Top + 30
    }

    ## Delegates
    $Script:lblDeleName = New-Object System.Windows.Forms.Label
        $Script:lblDeleName.Top = $Top ; $Script:lblDeleName.Left = 10; $Script:lblDeleName.Width=150; $Script:lblDeleName.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblDeleName)    # Add to Form 
        #
        $Script:txtDeLeName = New-Object Windows.Forms.TextBox  
        $Script:txtDeleName.TabIndex = 0 # set Tab Order 
        $Script:txtDeleName.Top = $Top; $Script:txtDeleName.Left = 140; $Script:txtDeleName.Width = 220;
    If ($Script:txtStdDele.Checked -eq "Checked")
    {
        #Standard Site Delegates
        $Script:lblDeleName.Text = "Site Delegates:"
        $Script:txtDeleName.Text = ("MBX." + ($Script:DispName.Substring(0,3)) + ".RRS.OutOfPolicy.DE")
    }
    else
    {
        #Restricted Room Delegates or Custom General User Room Delegates
        $Script:lblDeleName.Text = "Room Delegates Group:"
        $Script:txtDeleName.Text = ("MBX." + ($Script:DispName.Substring(0,3)) + ".RRS.OutOfPolicy." + ($Script:DispName.substring(4,$Script:DispName.IndexOf(" ")+1)) + ".DE")
    }
    $Global:form.Controls.Add($Script:txtDeleName)    # Add to Form
    $Top = $Top + 30

    ## Room Delegate Info
    $Exists = [bool](Get-DistributionGroup $Script:txtDeleName.Text -ErrorAction SilentlyContinue)
    if ($Exists -eq $False)
    {
        $Script:lblDeleMgrs = New-Object System.Windows.Forms.Label
        $Script:lblDeleMgrs.Text = "Delegate Emp Nos.:" 
        $Script:lblDeleMgrs.Top = $Top ; $Script:lblDeleMgrs.Left = 10; $Script:lblDeleMgrs.Width=150; $Script:lblDeleMgrs.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblDeleMgrs)    # Add to Form 
        #
        $Script:txtDeLeMgrs = New-Object Windows.Forms.TextBox  
        $Script:txtDeleMgrs.TabIndex = 0 # set Tab Order 
        $Script:txtDeleMgrs.Top = $Top; $Script:txtDeleMgrs.Left = 140; $Script:txtDeleMgrs.Width = 220;
        $Global:form.Controls.Add($Script:txtDeleMgrs)    # Add to Form
        $Top = $Top + 30
    }

    ## Restricted Users
    If ($Script:txtResUsr.Checked -eq "Checked")
    {
        $Script:lblResUsr = New-Object System.Windows.Forms.Label
        $Script:lblResUsr.Text = "Room Users Group:" 
        $Script:lblResUsr.Top = $Top ; $Script:lblResUsr.Left = 10; $Script:lblResUsr.Width=150; $Script:lblResUsr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblResUsr)    # Add to Form 
        #
        $Script:txtResUsr = New-Object Windows.Forms.TextBox  
        $Script:txtResUsr.TabIndex = 0 # set Tab Order 
        $Script:txtResUsr.Top = $Top; $Script:txtResUsr.Left = 140; $Script:txtResUsr.Width = 220;
        $Script:txtResUsr.Text = ("MBX." + ($Script:DispName.Substring(0,3)) + ".RRS.OutofPolicy." + ($Script:DispName.substring(4,$Script:DispName.IndexOf(" ")+1)) + ".US")
        $Global:form.Controls.Add($Script:txtResUsr)    # Add to Form
        $Top = $Top + 30
    }

    If ($Script:txtResUsr.Checked -eq "Checked")
    {
    ## Restricted Users Info
        $Exists = [bool](Get-DistributionGroup $Script:txtResUsr.Text -ErrorAction SilentlyContinue)
        if ($Exists -eq $False)
        {
            $Script:lblResUsrInfo = New-Object System.Windows.Forms.Label
            $Script:lblResUsrInfo.Text = "Users Emp Nos.:" 
            $Script:lblResUsrInfo.Top = $Top ; $Script:lblResUsrInfo.Left = 10; $Script:lblResUsrInfo.Width=150; $Script:lblResUsrInfo.AutoSize = $true 
            $Global:form.Controls.Add($Script:lblResUsrInfo)    # Add to Form 
            #
            $Script:txtResUsrInfo = New-Object Windows.Forms.TextBox  
            $Script:txtResUsrInfo.TabIndex = 0 # set Tab Order 
            $Script:txtResUsrInfo.Top = $Top; $Script:txtResUsrInfo.Left = 140; $Script:txtResUsrInfo.Width = 220;
            $Global:form.Controls.Add($Script:txtResUsrInfo)    # Add to Form
            $Top = $Top + 30
        }
    }

    ## TimeZone
    $Script:lblTimeZone = New-Object System.Windows.Forms.Label
        $Script:lblTimeZone.Text = "Room Time Zone:" 
        $Script:lblTimeZone.Top = $Top ; $Script:lblTimeZone.Left = 10; $Script:lblTimeZone.Width=150; $Script:lblTimeZone.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblTimeZone)    # Add to Form 
        #
        $Script:txtTimeZone = New-Object Windows.Forms.TextBox  
        $Script:txtTimeZone.TabIndex = 0 # set Tab Order 
        $Script:txtTimeZone.Top = $Top; $Script:txtTimeZone.Left = 140; $Script:txtTimeZone.Width = 220;
        $Script:txtTimeZone.Text = $Script:TimeZone
        $Global:form.Controls.Add($Script:txtTimeZone)    # Add to Form
    $Top = $Top + 30

    ## Regional Managers
    $Script:lblRegMgr = New-Object System.Windows.Forms.Label
        $Script:lblRegMgr.Text = "Regional Managers:" 
        $Script:lblRegMgr.Top = $Top ; $Script:lblRegMgr.Left = 10; $Script:lblRegMgr.Width=150; $Script:lblRegMgr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblRegMgr)    # Add to Form 
        #
        $Script:txtRegMgr = New-Object Windows.Forms.TextBox  
        $Script:txtRegMgr.TabIndex = 0 # set Tab Order 
        $Script:txtRegMgr.Top = $Top; $Script:txtRegMgr.Left = 140; $Script:txtRegMgr.Width = 220;
        $Script:txtRegMgr.Text = $Script:RegMgr
        $Global:form.Controls.Add($Script:txtRegMgr)    # Add to Form

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
    $Script:lblOfficeCode = New-Object System.Windows.Forms.Label   
        $Script:lblOfficeCode.Text = "Site Code:"  
        $Script:lblOfficeCode.Top = 10 ; $Script:lblOfficeCode.Left = 10; $Script:lblOfficeCode.Width=120 ; $Script:lblOfficeCode.AutoSize = $true
        $Global:form.Controls.Add($Script:lblOfficeCode)    # Add to Form 
        # 
        $Script:txtInpOfficeCode = New-Object Windows.Forms.TextBox  
        $Script:txtInpOfficeCode.TabIndex = 0 # set Tab Order 
        $Script:txtInpOfficeCode.Top = 10; $Script:txtInpOfficeCode.Left = 140; $Script:txtInpOfficeCode.Width = 200;  
        $Script:txtInpOfficeCode.Text = $Script:MbxLoc.ToUpper($Script:MbxLoc)
        $Global:form.Controls.Add($Script:txtInpOfficeCode)   # Add to Form 

    ## Time Zone
    $Script:lblNewTimeZone = New-Object System.Windows.Forms.Label   
        $Script:lblNewTimeZone.Text = "Select Time Zone:"
        $Script:lblNewTimeZone.Top = 40 ; $Script:lblNewTimeZone.Left = 10; $Script:lblNewTimeZone.Width=120 ;$Script:lblNewTimeZone.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblNewTimeZone)    # Add to Form 
        # 
        $Script:txtNewTimeZone = New-Object Windows.Forms.ListBox  
        $Script:txtNewTimeZone.TabIndex = 0 # set Tab Order 
        $Script:txtNewTimeZone.Top = 40; $Script:txtNewTimeZone.Left = 140; $Script:txtNewTimeZone.Width = 200; $Script:txtNewTimeZone.Height = 70
        $Zones = import-csv "e:\O365AdminShared\Data\UniqueTimeZones.txt"
        Foreach ($Zones in $Zones)
        {
            [void] $Script:txtNewTimeZone.Items.Add($Zones.TimeZone)
        }
        $Global:form.Controls.Add($Script:txtNewTimeZone)    # Add to Form 
        $Global:form.Topmost = $true

    ## Regional Location
    $Script:lblRegionLoc = New-Object System.Windows.Forms.Label   
        $Script:lblRegionLoc.Text = "Select Region:"
        $Script:lblRegionLoc.Top = 130 ; $Script:lblRegionLoc.Left = 10; $Script:lblRegionLoc.Width=120 ;$Script:lblRegionLoc.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblRegionLoc)    # Add to Form 
        # 
        $Script:txtRegionLoc = New-Object System.Windows.Forms.ListBox
        $Script:txtRegionLoc.TabIndex = 0 # set Tab Order 
        $Script:txtRegionLoc.Top = 130; $Script:txtRegionLoc.Left = 140; $Script:txtRegionLoc.Width = 200; $Script:txtRegionLoc.Height = 70;
        [void] $Script:txtRegionLoc.Items.Add('AP')
        [void] $Script:txtRegionLoc.Items.Add('CA')
        [void] $Script:txtRegionLoc.Items.Add('EU')
        [void] $Script:txtRegionLoc.Items.Add('LA')
        [void] $Script:txtRegionLoc.Items.Add('US')
        $Script:txtRegionLoc.TabIndex = 0
        $Global:form.Controls.Add($Script:txtRegionLoc)    # Add to Form
        $Global:form.Topmost = $true

    Add-FormStandardButtons
}

function Build-MbxName
{
    #Adjust name to proper case ann replace acronyms
    $Script:Name = (Get-Culture).textinfo.totitlecase($Script:txtInpName.Text.ToLower())
    $Space = $Script:Name.IndexOf(" ")
    $NameLen = $Script:Name.Length
    $Script:MbxLoc = $Script:Name.Substring(0,$Script:Name.IndexOf(" "))
    $MbxName = $Script:Name.Substring($Space+1,($NameLen-($Space+1)))
    $RoomName = ((Get-Culture).textinfo.totitlecase($MbxName.ToLower())) -replace(" ","")
    $Script:RoomGrp = $Script:MbxLoc + "Conference Rooms"
    
    foreach ($Acro in $Acro)
    {
        $MbxName = $MbxName -Replace($Acro.Acronym,$Acro.Translation)
    }

    If ($Script:txtNewFloor.Checked -eq "Checked")
    {
        $Script:RoomAddr = $Script:MbxLoc.ToUpper()
    }
    else
    {
        $Script:RoomAddr = "RES." + $Script:MbxLoc.ToUpper()
    }

    If ($Script:txtBldgName.Text.Length -ne 0)
    {
        $Script:RoomAddr = $Script:RoomAddr + "Bldg" + $Script:txtBldgName.Text + "."
    }
    else
    {
        $Script:RoomAddr = $Script:RoomAddr + "."
    }
        
    If ($Script:txtFlrNo.Text.Length -ne 0)
    {
        If ($Script:txtFlrNo.Text -eq "Ground")
        {
            $Script:RoomAddr = $Script:RoomAddr + "FLRGrnd"
        }
        else
        {
            $Script:RoomAddr = $Script:RoomAddr + "FLR" + $Script:txtFlrNo.Text
        }
    }
    
    If ($Script:txtNewGREquip.Checked -eq "Checked")
    {
        $Script:RoomAddr = ($Script:RoomAddr + "." + $RoomName + "@ul.com").Replace("..",".")
    }
    else
    {
        $Script:RoomAddr = ($Script:RoomAddr + "." + $RoomName + "." + $Script:txtRoomCap.Text + "@ul.com").Replace("..",".")        
    }

    $Script:DispName = ($Script:MbxLoc).ToUpper() + " "  + $MbxName
}

$Global:OKDetails = ""
$Global:form = ""
$Acro = Import-Csv e:\O365AdminShared\Data\KnownAcronyms.csv

Add-Type -Assembly System.Windows.Forms     ## Load the Windows Forms assembly 
## Create the main form

$Global:OKDetails = "Continue"
$Script:txtRoomCap = ""
$Script:TimeZone = ""
$Script:RegMgr = ""
$siteDet = import-csv "e:\O365AdminShared\Data\RoomTimeZones.csv"
Build-RRMbxInputForm
#Publish-Form

If($Script:txtNewFloor.Checked -eq "Checked")
{
    Do
    {
        Publish-Form
    } while (($Script:txtRoomCap.Text.Length -eq 0) -and ($Global:Result -eq "OK"))
}
else
{
    Publish-Form
}

If ($Global:Result -eq "OK")
{
    #Adjust name to proper case ann replace acronyms
    $Script:Name = (Get-Culture).textinfo.totitlecase($Script:txtInpName.Text.ToLower())
    $Script:Addr = ((Get-Culture).textinfo.totitlecase($Script:txtInpName.Text)).replace(" ","")
    $Space = $Script:Name.IndexOf(" ")
    $NameLen = $Script:Name.Length
    $Script:MbxLoc = $Script:Name.Substring(0,$Script:Name.IndexOf(" "))
    $MbxName = $Script:Name.Substring($Space+1,($NameLen-($Space+1)))
    $MbxName = ((Get-Culture).textinfo.totitlecase($MbxName.ToLower()))
    $Script:TaskNo = $Script:txtInpTaskNo.Text

    $BldDetails = "Y"
    Build-MbxName

    foreach ($SiteDet in $SiteDet)
    {
        If ($SiteDet.Code -eq $Script:MbxLoc)
        {
            $Script:TimeZone = $SiteDet.Zone
            $Script:RegMgr = $SiteDet.Admins
            $Found = "Y"
        }
    }

    If ($Script:TimeZone -eq "")
    {
        Get-NewSiteInfo
        Publish-Form

        If ($Global:Result -eq "OK")
        {
            $Script:RegMgr = "MBX." + $Script:txtRegionLoc.SelectedItem + ".RRS.Admins"

            If ($Script:txtNewTimeZone.SelectedItem -ne "Not Listed" )
            {
                $Script:TimeZone = $Script:txtNewTimeZone.SelectedItem
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
                $Script:lblNewZone = New-Object System.Windows.Forms.Label   
                $Script:lblOfficeCode.Text = "Time Zone:"  
                $Script:lblNewZone.Top = 30 ; $Script:lblNewZone.Left = 10; $Script:lblNewZone.Width=120 ; $Script:lblNewZone.AutoSize = $true
                $Global:form.Controls.Add($Script:lblNewZone)    # Add to Form 
                # 
                $Script:txtInpNewZone = New-Object Windows.Forms.TextBox  
                $Script:txtInpNewZone.TabIndex = 0 # set Tab Order 
                $Script:txtInpNewZone.Top = 30; $Script:txtInpNewZone.Left = 140; $Script:txtInpNewZone.Width = 200;  
                $Script:txtInpNewZone.Text = ""
                $Global:form.Controls.Add($Script:txtInpNewZone)   # Add to Form

                Add-FormStandardButtons
                Publish-Form
                $Script:TimeZone = $Script:txtInpNewZone.Text
            }

            #Add new location to the RoomTimeZone file
            $LineToWrite = "{0},""{1}"",{2}" -f ($Script:MbxLoc.ToUpper($Script:MbxLoc)),$Script:TimeZone,$Script:RegMgr
            Out-File -FilePath "E:\O365AdminShared\Data\RoomTimeZones.csv" -InputObject $LineToWrite -Append
            write-host "Please inform Enterprise Messaging Services Teams so this new Site Location/TimeZone can be added to the pick list" -ForegroundColor Red
        }
    }

    If ($Global:Result -eq "OK")
    {
        $Script:TaskNo = $Script:txtInpTaskNo.Text
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

	$atMAIL   = $Script:txtMbxAddr.Text.indexOf("@")
    $LeftName = $Script:txtMbxAddr.Text.substring(0,$atMAIL)
    
    If ($Script:txtStdDele.Checked -eq "Checked")
    {
        $ReportFile	= "E:\Automation\NewConferenceRoom\Report\Report-NewConferenceRoom-Date" + $Date + "Time" + $Time + ".log"
        If ($Script:txtNewGREquip.Checked -eq "Checked")
        {
            $ReportFile	= "E:\Automation\NewEquipment\Report\Report-NewEquipment-Date" + $Date + "Time" + $Time + ".log"
        }
    }
    else
    {
        $ReportFile	= "E:\Automation\NewConferenceRoom\Report\Report-NewRestrictedConferenceRoom-Date" + $Date + "Time" + $Time + ".log"
        If ($Script:txtNewGREquip.Checked -eq "Checked")
        {
            $ReportFile	= "E:\Automation\NewEquipment\Report\Report-NewRestrictedEquipment-Date" + $Date + "Time" + $Time + ".log"
        }
    }

    $LineToWrite = "START" + "`t" + "NewRoomEquipment script has started"
 	WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
    WriteReportEvent  

    If (Get-Mailbox $Script:txtDispName.Text -ErrorAction SilentlyContinue)
    {
        $Output = $wshell.Popup("This room/resource already exists",0,"Room/Resource Exists",0+32)
        $LineToWrite = "STAR" + "`t" + "Room/Resource Already Exists: " + $Script:txtDispName.Text
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

    If ($Script:txtNewFloor.Checked -eq "Checked")
    {
        write-host "Room creation complete"
        $Output = $wshell.Popup("Room " + $Script:txtDispName.Text + " creation complete.",0,"Room Created",0+32)
    }
    else
    {
        write-host "Resource creation complete"
        $Output = $wshell.Popup("Resource " + $Script:txtDispName.Text + " creation complete.",0,"Resource Created",0+32)
    }

    If ($Global:NewRoomList -eq "Y")
    {
        If (($Script:txtStdDele.Checked -eq "Checked") -or($Script:txtGenUseCustDele.Checked -eq "Checked"))
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
        If (($Script:txtStdDele.Checked -eq "Checked") -or($Script:txtGenUseCustDele.Checked -eq "Checked"))
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
    If ($Script:txtNewFloor.Checked -eq "Checked")
    {
        write-host "Room creation cancelled"
        $Output = $wshell.Popup("Room " + $Script:txtDispName.Text + " creation cancelled.",0,"Room Cancelled",0+32)
    }
    else
    {
        write-host "Resource creation cancelled"
        $Output = $wshell.Popup("Resource " + $Script:txtDispName.Text + " creation cancelled.",0,"Resource Cancelled",0+32)
    }
}