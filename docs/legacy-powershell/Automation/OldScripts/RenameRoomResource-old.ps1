<#   
================================================================================ 
 Name: Rename Shared Mailbox Form
 ================================================================================ 
#>  

function Build-RenRRInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Rename Room or Resource Mailbox" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(450,250) #(W,H)
    $BldDetails = "N"

    ## Get Details used to create
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Display Name:"
        $Global:lblDispName.Top = 10 ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtInpName = New-Object Windows.Forms.TextBox  
        $Global:txtInpName.TabIndex = 0 # set Tab Order 
        $Global:txtInpName.Top = 10; $Global:txtInpName.Left = 130; $Global:txtInpName.Width = 280;  
        $Global:txtInpName.Text = ""   # DisplayName
        $Global:txtInpName.TabIndex = 0
        $Global:form.Controls.Add($Global:txtInpName)    # Add to Form

    ## Building Name
    $Global:lblBldgName = New-Object System.Windows.Forms.Label   
        $Global:lblBldgName.Text = "Building Name:"
        $Global:lblBldgName.Top = 40 ; $Global:lblBldgName.Left = 10; $Global:lblBldgName.Width=120 ;$Global:lblBldgName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblBldgName)    # Add to Form 
        # 
        $Global:txtBldgName = New-Object Windows.Forms.TextBox  
        $Global:txtBldgName.TabIndex = 0 # set Tab Order 
        $Global:txtBldgName.Top = 40; $Global:txtBldgName.Left = 130; $Global:txtBldgName.Width = 280;  
        $Global:txtBldgName.Text = ""   # BuildingName
        $Global:txtBldgName.TabIndex = 0
        $Global:form.Controls.Add($Global:txtBldgName)    # Add to Form

    ## Floor Number
    $Global:lblFlrNo = New-Object System.Windows.Forms.Label   
        $Global:lblFlrNo.Text = "Floor No.:"
        $Global:lblFlrNo.Top = 70 ; $Global:lblFlrNo.Left = 10; $Global:lblFlrNo.Width=120 ;$Global:lblFlrNo.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblFlrNo)    # Add to Form 
        # 
        $Global:txtFlrNo = New-Object Windows.Forms.TextBox  
        $Global:txtFlrNo.TabIndex = 0 # set Tab Order 
        $Global:txtFlrNo.Top = 70; $Global:txtFlrNo.Left = 130; $Global:txtFlrNo.Width = 280;  
        $Global:txtFlrNo.Text = ""   # DisplayName
        $Global:txtFlrNo.TabIndex = 0
        $Global:form.Controls.Add($Global:txtFlrNo)    # Add to Form

    ## Room Capacity
    $Global:lblRoomCap = New-Object System.Windows.Forms.Label   
        $Global:lblRoomCap.Text = "Room Capacity:"
        $Global:lblRoomCap.Top = 100 ; $Global:lblRoomCap.Left = 10; $Global:lblRoomCap.Width=120 ;$Global:lblRoomCap.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblRoomCap)    # Add to Form 
        # 
        $Global:txtRoomCap = New-Object Windows.Forms.TextBox  
        $Global:txtRoomCap.TabIndex = 0 # set Tab Order 
        $Global:txtRoomCap.Top = 100; $Global:txtRoomCap.Left = 130; $Global:txtRoomCap.Width = 280;  
        $Global:txtRoomCap.Text = ""   # DisplayName
        $Global:txtRoomCap.TabIndex = 0
        $Global:form.Controls.Add($Global:txtRoomCap)    # Add to Form

    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 130 ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtInpTaskNo.Top = 130; $Global:txtInpTaskNo.Left = 130; $Global:txtInpTaskNo.Width = 120;  
        $Global:txtInpTaskNo.Text = "TASK"   # Enter ticket number
        $Global:txtInpTaskNo.TabIndex = 1
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form

        Add-FormStandardButtons
}

function Build-RenMbxDetailsForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Rename Room or Resource Mailbox" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(740,640) #(W,H)

    Add-FormStandardButtons

    ## Get Details to complete creation
    $TopLoc = 10
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Old Room/Resource Name:"
        $Global:lblDispName.Top = $TopLoc; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=150 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtDispName = New-Object Windows.Forms.TextBox  
        $Global:txtDispName.TabIndex = 0 # set Tab Order 
        $Global:txtDispName.Top = $TopLoc; $Global:txtDispName.Left = 120; $Global:txtDispName.Width = 220;  
        $Global:txtDispName.Text = $Global:txtInpName.Text   # DisplayName
        $Global:form.Controls.Add($Global:txtDispName)    # Add to Form

    $TopLoc = $TopLoc + 30
    ## Ticket Number
    $Global:lblTaskNoDet = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNoDet.Text = "Ticket Number:"
        $Global:lblTaskNoDet.Top = $TopLoc; $Global:lblTaskNoDet.Left = 10; $Global:lblTaskNoDet.Width=150 ;$Global:lblTaskNoDet.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblTaskNoDet)    # Add to Form 
        # 
        $Global:txtTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtTaskNo.Top = $TopLoc; $Global:txtTaskNo.Left = 120; $Global:txtTaskNo.Width = 220;  
        $Global:txtTaskNo.Text = $Global:txtInpTaskNo.Text   # DisplayName
        $Global:form.Controls.Add($Global:txtTaskNo)    # Add to Form

    $TopLoc = $TopLoc + 30
    ## New Display Name
    $Global:lblNewDispName = New-Object System.Windows.Forms.Label   
        $Global:lblNewDispName.Text = "New Room/Resource Name:"
        $Global:lblNewDispName.Top = $TopLoc; $Global:lblNewDispName.Left = 10; $Global:lblNewDispName.Width=150 ;$Global:lblNewDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblNewDispName)    # Add to Form 
        # 
        $Global:txtNewDispName = New-Object Windows.Forms.TextBox  
        $Global:txtNewDispName.TabIndex = 0 # set Tab Order 
        $Global:txtNewDispName.Top = $TopLoc; $Global:txtNewDispName.Left = 120; $Global:txtNewDispName.Width = 220;  
        $Global:txtNewDispName.Text = $Global:DispName   # DisplayName
        $Global:form.Controls.Add($Global:txtNewDispName)    # Add to Form 

    $TopLoc = $TopLoc + 30
    ## New Mailbox Alias
    $Global:lblNewMbxAlias = New-Object System.Windows.Forms.Label   
        $Global:lblNewMbxAlias.Text = "New Mailbox Alias:"
        $Global:lblNewMbxAlias.Top = $TopLoc; $Global:lblNewMbxAlias.Left = 10; $Global:lblNewMbxAlias.Width=150 ;$Global:lblNewMbxAlias.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblNewMbxAlias)    # Add to Form 
        # 
        $Global:txtNewMbxAlias = New-Object Windows.Forms.TextBox  
        $Global:txtNewMbxAlias.TabIndex = 0 # set Tab Order 
        $Global:txtNewMbxAlias.Top = $TopLoc; $Global:txtNewMbxAlias.Left = 120; $Global:txtNewMbxAlias.Width = 220;  
        $Global:txtNewMbxAlias.Text = $Global:GrpAddr
        $Global:form.Controls.Add($Global:txtNewMbxAlias)    # Add to Form 

    $TopLoc = $TopLoc + 30
    ## EmailAddress
    $Global:lblMbxAddr = New-Object System.Windows.Forms.Label   
        $Global:lblMbxAddr.Text = "New Email Address:"  
        $Global:lblMbxAddr.Top = $TopLoc; $Global:lblMbxAddr.Left = 10; $Global:lblMbxAddr.Width=150 ;$Global:lblMbxAddr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblMbxAddr)    # Add to Form 
        # 
        $Global:txtMbxAddr = New-Object Windows.Forms.TextBox  
        $Global:txtMbxAddr.TabIndex = 0 # set Tab Order 
        $Global:txtMbxAddr.Top = $TopLoc; $Global:txtMbxAddr.Left = 120; $Global:txtMbxAddr.Width = 220;  
        $Global:txtMbxAddr.Text = $Global:Address   # Email Address
        $Global:form.Controls.Add($Global:txtMbxAddr)    # Add to Form 

    $TopLoc = $TopLoc + 30
    ## Legacyddress
    $Global:lblLegAddr = New-Object System.Windows.Forms.Label   
        $Global:lblLegAddr.Text = "Additional Aliases:"  
        $Global:lblLegAddr.Top = $TopLoc; $Global:lblLegAddr.Left = 10; $Global:lblLegAddr.Width=150 ;$Global:lblLegAddr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblLegAddr)    # Add to Form 
        # 
        $Global:txtLegAddr = New-Object Windows.Forms.TextBox  
        $Global:txtLegAddr.TabIndex = 0 # set Tab Order 
        $Global:txtLegAddr.Top = $TopLoc; $Global:txtLegAddr.Left = 120; $Global:txtLegAddr.Width = 520;  
        $Global:txtLegAddr.Text = ((get-mailbox $Global:txtDispName.Text).EMailAddresses -replace "smtp:","") -join ", "
        $Global:form.Controls.Add($Global:txtLegAddr)    # Add to Form 

    $TopLoc = $TopLoc + 30
    ## Mailbox Owner
    $Global:lblGrpOwnr = New-Object System.Windows.Forms.Label   
        $Global:lblGrpOwnr.Text = "Mailbox Owner(s):"  
        $Global:lblGrpOwnr.Top = $TopLoc; $Global:lblGrpOwnr.Left = 10; $Global:lblGrpOwnr.Width=150; $Global:lblGrpOwnr.AutoSize = $true 
        $Global:form.Controls.Add($lblGrpOwnr)    # Add to Form 
        # 
        $Global:txtGrpOwnr = New-Object Windows.Forms.TextBox  
        $Global:txtGrpOwnr.TabIndex = 0 # set Tab Order 
        $Global:txtGrpOwnr.Top = $TopLoc; $Global:txtGrpOwnr.Left = 120; $Global:txtGrpOwnr.Width = 520;  
        $Global:txtGrpOwnr.Text = $MbxOwner
        $Global:form.Controls.Add($Global:txtGrpOwnr)    # Add to Form 

    $TopLoc = 220
    If ($Global:EDAccess -ne "")
    {
        ## Editor Old Group Name
        $Global:lblEDGrpName = New-Object System.Windows.Forms.Label
        $Global:lblEDGrpName.Text = ".ED Old Group:"  
        $Global:lblEDGrpName.Top = $TopLoc ; $Global:lblEDGrpName.Left = 10; $Global:lblEDGrpName.Width=150; $Global:lblEDGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblEDGrpName)    # Add to Form 
        #
        $Global:txtEDGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtEDGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtEDGrpName.Top = $TopLoc; $Global:txtEDGrpName.Left = 120; $Global:txtEDGrpName.Width = 180;  
        $Global:txtEDGrpName.Text = $Global:EDAccess
        $Global:form.Controls.Add($Global:txtEDGrpName)    # Add to Form 

        ## Editor New Group Name
        $Global:lblNewEDGrpName = New-Object System.Windows.Forms.Label
        $Global:lblNewEDGrpName.Text = ".ED New Group:"  
        $Global:lblNewEDGrpName.Top = $TopLoc ; $Global:lblNewEDGrpName.Left = 310; $Global:lblNewEDGrpName.Width=150; $Global:lblNewEDGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblNewEDGrpName)    # Add to Form 
        #
        $Global:txtNewEDGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtNewEDGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtNewEDGrpName.Top = $TopLoc; $Global:txtNewEDGrpName.Left = 430; $Global:txtNewEDGrpName.Width = 210;  
        $Global:txtNewEDGrpName.Text = ("MBX." + $Global:txtNewDispName.Text + ".ED")
        $Global:form.Controls.Add($Global:txtNewEDGrpName)    # Add to Form 

        $TopLoc = $TopLoc + 30
        ## Editor Group Members
        $Global:lblEDGrpMbr = New-Object System.Windows.Forms.Label   
        $Global:lblEDGrpMbr.Text = ".ED Member(s):"  
        $Global:lblEDGrpMbr.Top = $TopLoc; $Global:lblEDGrpMbr.Left = 10; $Global:lblEDGrpMbr.Width=150; $Global:lblEDGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblEDGrpMbr)    # Add to Form 
        # 
        $Global:txtEDGrpMbr = New-Object Windows.Forms.TextBox
        $Global:txtEDGrpMbr.TabIndex = 0 # set Tab Order
        $Global:txtEDGrpMbr.Location = New-Object System.Drawing.Size(120,$TopLoc)
        $Global:txtEDGrpMbr.Size = New-Object system.Drawing.Size(520,60)
        $Global:txtEDGrpMbr.MultiLine = $true
        $Global:txtEDGrpMbr.ScrollBars = 'Both'
        $Global:txtEDGrpMbr.Text = $Global:EDMembers 
        $Global:form.Controls.Add($Global:txtEDGrpMbr)    # Add to Form 

        $TopLoc = $TopLoc + 70
    }

    If ($AUAccess -ne "")
    {
        ## Author Group Name
        $Global:lblAUGrpName = New-Object System.Windows.Forms.Label
        $Global:lblAUGrpName.Text = ".AU Old Group:" 
        $Global:lblAUGrpName.Top = $TopLoc; $Global:lblAUGrpName.Left = 10; $Global:lblAUGrpName.Width=150; $Global:lblAUGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAUGrpName)    # Add to Form 
        #
        $Global:txtAUGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtAUGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtAUGrpName.Top = $TopLoc; $Global:txtAUGrpName.Left = 120; $Global:txtAUGrpName.Width = 180;  
        $Global:txtAUGrpName.Text = $Global:AUAccess   # Use Corrent computer name as default 
        $Global:form.Controls.Add($Global:txtAUGrpName)    # Add to Form

        ## Author New Group Name
        $Global:lblNewAUGrpName = New-Object System.Windows.Forms.Label
        $Global:lblNewAUGrpName.Text = ".AU New Group:"  
        $Global:lblNewAUGrpName.Top = $TopLoc ; $Global:lblNewAUGrpName.Left = 310; $Global:lblNewAUGrpName.Width=150; $Global:lblNewAUGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblNewAUGrpName)    # Add to Form 
        #
        $Global:txtNewAUGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtNewAUGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtNewAUGrpName.Top = $TopLoc; $Global:txtNewAUGrpName.Left = 430; $Global:txtNewAUGrpName.Width = 210;  
        $Global:txtNewAUGrpName.Text = ("MBX." + $Global:txtNewDispName.Text + ".AU")
        $Global:form.Controls.Add($Global:txtNewAUGrpName)    # Add to Form 

        $TopLoc = $TopLoc + 30
        ## Author Group Members
        $Global:lblAUGrpMbr = New-Object System.Windows.Forms.Label   
        $Global:lblAUGrpMbr.Text = ".AU Member(s):"  
        $Global:lblAUGrpMbr.Top = $TopLoc; $Global:lblAUGrpMbr.Left = 10; $Global:lblAUGrpMbr.Width=150; $Global:lblAUGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAUGrpMbr)    # Add to Form 
        # 
        $Global:txtAUGrpMbr = New-Object Windows.Forms.TextBox
        $Global:txtAUGrpMbr.TabIndex = 0 # set Tab Order 
        $Global:txtAUGrpMbr.Text = $Global:AUMembers
        $Global:txtAUGrpMbr.Location = New-Object System.Drawing.Size(120,$TopLoc)
        $Global:txtAUGrpMbr.Size = New-Object system.Drawing.Size(520,60)
        $Global:txtAUGrpMbr.MultiLine = $true
        $Global:txtAUGrpMbr.ScrollBars = 'Both'  
        $Global:form.Controls.Add($Global:txtAUGrpMbr)    # Add to Form 

        $TopLoc = $TopLoc + 70
    }

    If ($REAccess -ne "")
    {
        # Reader Group Name
        $Global:lblREGrpName = New-Object System.Windows.Forms.Label
        $Global:lblREGrpName.Text = ".RE Old Group:" 
        $Global:lblREGrpName.Top = $TopLoc ; $Global:lblREGrpName.Left = 10; $Global:lblREGrpName.Width=150; $Global:lblREGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblREGrpName)    # Add to Form 
        #
        $Global:txtREGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtREGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtREGrpName.Top = $TopLoc; $Global:txtREGrpName.Left = 120; $Global:txtREGrpName.Width = 180;  
        $Global:txtREGrpName.Text = $Global:REAccess
        $Global:form.Controls.Add($Global:txtREGrpName)    # Add to Form 
       # Obtain Value with: $txtGrpOwnr.Text

        ## Author New Group Name
        $Global:lblNewREGrpName = New-Object System.Windows.Forms.Label
        $Global:lblNewREGrpName.Text = ".RE New Group:"  
        $Global:lblNewREGrpName.Top = $TopLoc ; $Global:lblNewREGrpName.Left = 310; $Global:lblNewREGrpName.Width=150; $Global:lblNewREGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblNewREGrpName)    # Add to Form 
        #
        $Global:txtNewREGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtNewREGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtNewREGrpName.Top = $TopLoc; $Global:txtNewREGrpName.Left = 430; $Global:txtNewREGrpName.Width = 210;  
        $Global:txtNewREGrpName.Text = ("MBX." + $Global:txtNewDispName.Text + ".RE")
        $Global:form.Controls.Add($Global:txtNewREGrpName)    # Add to Form 

        $TopLoc = $TopLoc + 30
        ## Reader Group Members
        $Global:lblREGrpMbr = New-Object System.Windows.Forms.Label   
        $Global:lblREGrpMbr.Text = ".RE Member(s):"  
        $Global:lblREGrpMbr.Top = $TopLoc; $Global:lblREGrpMbr.Left = 10; $Global:lblREGrpMbr.Width=150; $Global:lblREGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblREGrpMbr)    # Add to Form 
        # 
        $Global:txtREGrpMbr = New-Object Windows.Forms.TextBox
        $Global:txtREGrpMbr.TabIndex = 0 # set Tab Order 
        $Global:txtREGrpMbr.Text = $Global:REMembers
        $Global:txtREGrpMbr.Location = New-Object System.Drawing.Size(120,$TopLoc)
        $Global:txtREGrpMbr.Size = New-Object system.Drawing.Size(520,60)
        $Global:txtREGrpMbr.MultiLine = $true
        $Global:txtREGrpMbr.ScrollBars = 'Both'  
        $Global:form.Controls.Add($Global:txtREGrpMbr)    # Add to Form 

        $TopLoc = $TopLoc + 70
    }

    ## Authorized Requestor
    $Global:chkAuthorized = New-Object Windows.Forms.checkbox 
        $Global:chkAuthorized.Left = 120; $Global:chkAuthorized.Width = 200; $Global:chkAuthorized.Top = $TopLoc
        $Global:chkAuthorized.Text = "Authorized Requestor" 
        $Global:chkAuthorized.Checked = $false   # set a default value 
        $Global:chkAuthorized.TabIndex = 5
        $Global:form.Controls.Add($Global:chkAuthorized) 
        # Obtain Value with: $Global:chkAuthorized.Checked
}

function Publish-Form
{
    ## Finalize Form and Show Dialog 
    $Global:form.Add_Shown( { $form.Activate(); $okButton.Focus() } )  #Activate and Set Focus 
    $Global:ShrResult = $Global:form.ShowDialog()          ## Show the form, and wait for the response 
}

function Build-MbxName
{
    #Adjust new name to proper case and replace acronyms 
    $Global:Name = (Get-Culture).textinfo.totitlecase($Global:txtInpNewName.Text.ToLower())
    $Space = $Global:Name.IndexOf(" ")
    $NameLen = $Global:Name.Length
    $MbxLoc = $Global:Name.Substring(0,$Global:Name.IndexOf(" "))
    $MbxName = $Global:Name.Substring($Space+1,($NameLen-($Space+1)))
    $MbxName = ((Get-Culture).textinfo.totitlecase($MbxName.ToLower()))

    foreach ($Acro in $Acro)
    {
        $MbxName = $MbxName -Replace($Acro.Acronym,$Acro.Translation)
    }

    $NewAlias = $MbxName -replace '\s',''
    if ($NewAlias -like "*&*")
    {
        $NewAlias = $NewAlias.Replace("&","")
    }
    if ($NewAlias -like "*,*")
    {
        $NewAlias = $NewAlias.Replace(",","")
    }
    if ($NewAlias -like "*-*")
    {
        $NewAlias = $NewAlias.Replace("-","")
    }
    if ($NewAlias -like "*/*")
    {
        $NewAlias = $NewAlias.Replace("/","")
    }

    If ($MbxLoc -notlike "Global*")
    {
        $Global:DispName = ($MbxLoc).ToUpper() + " "  + $MbxName
        $Global:Address = ($MbxLoc).ToUpper() + "." + $NewAlias.Replace(" ","") + "@ul.com"
        $Global:GrpAddr = ($MbxLoc).ToUpper() + "." + $NewAlias.Replace(" ","")
    }
    else
    {
        $Global:DispName = ((Get-Culture).textinfo.totitlecase($MbxLoc.ToLower())) + " "  + $MbxName
#        $Global:Address = ((Get-Culture).textinfo.totitlecase($NewAlias.ToLower())) + "." + $MbxName.Replace(" ","") + "@ul.com"
#        $Global:GrpAddr = ((Get-Culture).textinfo.totitlecase($NewAlias.ToLower())) + "." + $MbxName.Replace(" ","")
        $Global:Address = ((Get-Culture).textinfo.totitlecase($MbxLoc.ToLower())) + "." + $MbxName.Replace(" ","") + "@ul.com"
        $Global:GrpAddr = ((Get-Culture).textinfo.totitlecase($MbxLoc.ToLower())) + "." + $MbxName.Replace(" ","")
    }

    #Get Mailbox Owner and Access Groups
    $Global:MbxOwner = ""
    $Global:EDAccess = ""
    $Global:EDMembers = ""
    $Global:AUAccess = ""
    $Global:AUMembers = ""
    $Global:REAccess = ""
    $Global:REMembers = ""
    $Global:okButton.Text = "Continue"
    $MbxPerm = get-mailboxFolderPermission $Global:txtInpName.Text |Where-Object {$_.User -like "MBX*"}

    foreach ($MbxPerm in $MbxPerm)
    {
        If (($MbxPerm.User.Displayname -like "*ED") -or ($MbxPerm.AccessRights -eq "Editor"))
        {
            $Global:EDAccess = $MbxPerm.User.DisplayName
            $Global:EDMembers = (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ","
            $Global:MbxOwner = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
        }
        If (($MbxPerm.User.Displayname -like "*AU") -or ($MbxPerm.AccessRights -eq "PublishingAuthor"))
        {
            $Global:AUAccess = $MbxPerm.User.DisplayName
            $Global:AUMembers = (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ","
            If ($Global:MbxOwner -eq "")
            {
                $Global:MbxOwner = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
            }
        }
        If (($MbxPerm.User.Displayname -like "*RE") -or ($MbxPerm.AccessRights -eq "Reviewer"))
        {
            $Global:REAccess = $MbxPerm.User.DisplayName
            $Global:REMembers = (Get-DistributionGroupMember $MbxPerm.User.Displayname) -join ","
            If ($Global:MbxOwner -eq "")
            {
                $Global:MbxOwner = (Get-DistributionGroup $mbxPerm.User.DisplayName).ManagedBy -join ", "
            }
        }
    }
}

function Rename-AccessGroup
{
    $GrpInf = Get-DistributionGroup $Global:DSTGrp
    If (($GrpInfo.Notes -notlike $GrpNotes) -and ($GrpInfo.Notes -like "*Owner*"))
    {
        write-host "Setting Group Notes"
        Set-Group $Global:DSTGrp -Notes $GrpNotes
    }
    write-host "Changing distribution Group Name"
    Set-DistributionGroup $Global:DSTGrp -DisplayName $Global:NewDSTGrp -Name $NewDSTGrp -Alias $Global:NewDSTGrpAlias -EmailAddresses @{add=$Global:NewDSTIntAddr}
    write-host "Setting Distribution Group Primary Address"
    Set-DistributionGroup $Global:NewDSTGrp -PrimarySmtpAddress $Global:NewDSTIntAddr
}

$Acro = Import-Csv e:\O365AdminShared\Data\KnownAcronyms.csv
Add-Type -Assembly System.Windows.Forms     ## Load the Windows Forms assembly 

## Create the main form

$BldDetails = "N"
$Global:OKDetails = "Continue"
Build-RenRRInputForm
Publish-Form

If ($Global:ShrResult -eq "OK")
{
    $RRExists = [bool](Get-Mailbox $Global:txtInpName.Text -ErrorAction SilentlyContinue)
    If ($RRExists -eq "True")
    {
        #Adjust name to proper case ann replace acronyms
        $Global:Name = (Get-Culture).textinfo.totitlecase($Global:txtInpName.Text.ToLower())
        $Global:Addr = ((Get-Culture).textinfo.totitlecase($Global:txtInpName.Text)).replace(" ","")
        $Space = $Global:Name.IndexOf(" ")
        $NameLen = $Global:Name.Length
        $MbxLoc = $Global:Name.Substring(0,$Global:Name.IndexOf(" "))
        $MbxName = $Global:Name.Substring($Space+1,($NameLen-($Space+1)))
        $MbxName = ((Get-Culture).textinfo.totitlecase($MbxName.ToLower()))

        If ($Global:txtInpNewName.Text.IndexOf(" ") -eq 3)
        {
            $NewAlias = ($Global:txtInpNewName.Text.Insert($Global:txtInpNewName.Text.IndexOf(" "),"."))
            $NewAlias = $NewAlias -replace '\s',''
            if ($NewAlias -like "*&*")
            {
                $NewAlias = $NewAlias.Remove("&")
            }
            if ($NewAlias -like "*,*")
            {
                $NewAlias = $NewAlias.Remove(",")
            }
            if ($NewAlias -like "*-*")
            {
                $NewAlias = $NewAlias.Remove("-")
            }
            if ($NewAlias -like "*/*")
            {
                $NewAlias = $NewAlias.Remove("/")
            }
             $Global:NewInetAddr = $NewAlias + "@ul.com"
        }

        $BldDetails = "Y"
        Build-MbxName
        Build-RenMbxDetailsForm
        Publish-Form

        If (($Global:chkAuthorized.Checked -eq "Checked") -and ($Global:ShrResult -eq "OK"))
        {
            $GrpNotes = "*Owner: " + $Global:txtGrpOwnr.Text + "Per " + $Global:txtInpTaskNo.Text + "*"
            $MbxTip = "Owner: " + $Global:txtGrpOwnr.Text
            $NewAlias = $Global:txtMbxAddr.Text.Remove($Global:txtMbxAddr.Text.IndexOf("@ul"))
            $ShrMbxInf = get-mailbox $Global:txtInpName.Text
            $NewPrimaryAlias = $ShrMbxInf.EmailAddresses += "SMTP:" + $Global:txtMbxAddr.Text
            If (($ShrMbxInf.MailTip -ne $MbxTip) -and ($ShrMbxInf.MailTip -like "*Owner*"))
            {
                write-host "Renaming Mailbox"
                Set-Mailbox $Global:txtInpName.Text -MailTip $MbxTip
            }
            write-host "Setting Other Mailbox Details"
            Set-Mailbox $Global:txtInpName.Text -Name $Global:txtNewDispName.Text -DisplayName $Global:txtNewDispName.Text -Alias $NewAlias -EmailAddresses $NewPrimaryAlias

            If (($Global:EDAccess -ne "") -and ($Global:EDAccess -ne $Global:txtNewEDGrpName.Text))
            {
                $Global:DSTGrp = $Global:EDAccess
                $Global:NewDSTGrp = $Global:txtNewEDGrpName.Text
                $Global:NewDSTGrpAlias = "MBX." + $Global:txtNewMbxAlias.Text + ".ED"
                $Global:NewDSTIntAddr = $Global:NewDSTGrpAlias + "@ul.com"
                write-host "Renaming Editors Group"
                Rename-AccessGroup
            }

            If (($Global:AUAccess -ne "") -and ($Global:AUAccess -ne $Global:txtNewAUGrpName.Text))
            {
                $Global:DSTGrp = $Global:AUAccess
                $Global:NewDSTGrp = $Global:txtNewAUGrpName.Text
                $Global:NewDSTGrpAlias = "MBX." + $Global:txtNewMbxAlias.Text + ".AU"
                $Global:NewDSTIntAddr = $Global:NewDSTGrpAlias + "@ul.com"
                write-host "Renaming Authors Group"
                Rename-AccessGroup
            }

            If (($Global:AUAccess -ne "") -and ($Global:REAccess -ne $Global:txtNewREGrpName.Text))
            {
                $Global:DSTGrp = $Global:REAccess
                $Global:NewDSTGrp = $Global:txtNewREGrpName.Text
                $Global:NewDSTGrpAlias = "MBX." + $Global:txtNewMbxAlias.Text + ".RE"
                $Global:NewDSTIntAddr = $Global:NewDSTGrpAlias + "@ul.com"
                write-host "Renaming Readers Group"
                Rename-AccessGroup
            }
        }
        else
        {
            If ($Global:ShrResult -eq "OK")
            {
                $Output = $wshell.Popup("No changes made obtain authorization from an existing mailbox owner of " + $Global:txtInpName.Text + " mailbox",0,"Not Found",0+32)
            }
            else
            {
                $Output = $wshell.Popup("No changes made rename request cancelled." + $Global:txtInpName.Text + "Request cancelled",0,"Not Found",0+32)
            }
        }
    }
    else
    {
        $Output = $wshell.Popup("Shared Mailbox " + $Global:txtInpName.Text + " Does Not Exist - No Changes Made",0,"Not Found",0+32)
    }
}