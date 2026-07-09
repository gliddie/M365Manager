<#   
================================================================================ 
 Name: New Shared Mailbox Form
 ================================================================================ 
#>  

function Build-ShrMbxInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "New Shared Mailbox" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(400,280) #(W,H)

    Add-FormStandardButtons

    ## Get Details used to create
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Mailbox Name:"
        $Global:lblDispName.Top = 10 ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtInpName = New-Object Windows.Forms.TextBox  
        $Global:txtInpName.TabIndex = 0 # set Tab Order 
        $Global:txtInpName.Top = 10; $Global:txtInpName.Left = 120; $Global:txtInpName.Width = 200;  
        $Global:txtInpName.Text = ""   # DisplayName
        $Global:form.Controls.Add($Global:txtInpName)    # Add to Form
        $Global:InputFocus = $Global:txtInpName 
       # Obtain Value with: $Global:txtInpName.Text
 
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 40 ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtInpTaskNo.Top = 40; $Global:txtInpTaskNo.Left = 120; $Global:txtInpTaskNo.Width = 120;  
        $Global:txtInpTaskNo.Text = "TASK"   # Enter ticket number 
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to ForM

    ## Not for Profit Group CheckBox
    $Global:chkNFP = New-Object Windows.Forms.checkbox 
        $Global:chkNFP.Left = 120; $Global:chkNFP.Width = 280; $Global:chkNFP.Top = 70
        $Global:chkNFP.Text = "Requested by Not For Profit Group" 
        $Global:chkNFP.Checked = $false   # set a default value
        $Global:chkNFP.TabIndex = 2
        $Global:form.Controls.Add($Global:chkNFP)        
         
    ## Editors Group CheckBox
    $Global:chkEditors = New-Object Windows.Forms.checkbox 
        $Global:chkEditors.Left = 120; $Global:chkEditors.Width = 280; $Global:chkEditors.Top = 100  
        $Global:chkEditors.Text = "Create (.ED) Editors Group" 
        $Global:chkEditors.Checked = $false   # set a default value 
        $Global:chkEditors.TabIndex = 2
        $Global:form.Controls.Add($Global:chkEditors) 

    ## Authors Group CheckBox
    $Global:chkAuthors = New-Object Windows.Forms.checkbox 
        $Global:chkAuthors.Left = 120; $Global:chkAuthors.Width = 280; $Global:chkAuthors.Top = 130
        $Global:chkAuthors.Text = "Create (.AU) Authors Group" 
        $Global:chkAuthors.Checked = $false   # set a default value 
        $Global:chkAuthors.TabIndex = 2
        $Global:form.Controls.Add($Global:chkAuthors) 
 
     ## Readers Group CheckBox
    $Global:chkReaders = New-Object Windows.Forms.checkbox 
        $Global:chkReaders.Left = 120; $Global:chkReaders.Width = 280; $Global:chkReaders.Top = 160  
        $Global:chkReaders.Text = "Create (.RE) Readers Group" 
        $Global:chkReaders.Checked = $false   # set a default value 
        $Global:chkReaders.TabIndex = 2
        $Global:form.Controls.Add($Global:chkReaders) 
}

function Build-MbxDetailsForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "New Shared Mailbox" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(740,550) #(W,H)

    Add-FormStandardButtons

    ## Get Details to complete creation
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Mailbox Name:"
        $Global:lblDispName.Top = 10; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=150 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtDispName = New-Object Windows.Forms.TextBox  
        $Global:txtDispName.TabIndex = 0 # set Tab Order 
        $Global:txtDispName.Top = 10; $Global:txtDispName.Left = 120; $Global:txtDispName.Width = 220;  
        $Global:txtDispName.Text = $Global:DispName   # DisplayName
        $Global:form.Controls.Add($Global:txtDispName)    # Add to Form 
        $Global:InputFocus = $Global:txtDispName 

#    $Name = (Get-Culture).textinfo.totitlecase($Global:txtInpName.Text)
#    $Addr = ((Get-Culture).textinfo.totitlecase($Global:txtInpName.Text)).replace(" ","")
    ## EmailAddress
    $Global:lblMbxAddr = New-Object System.Windows.Forms.Label   
        $Global:lblMbxAddr.Text = "Email Address:"  
        $Global:lblMbxAddr.Top = 40; $Global:lblMbxAddr.Left = 10; $Global:lblMbxAddr.Width=150 ;$Global:lblMbxAddr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblMbxAddr)    # Add to Form 
        # 
        $Global:txtMbxAddr = New-Object Windows.Forms.TextBox  
        $Global:txtMbxAddr.TabIndex = 0 # set Tab Order 
        $Global:txtMbxAddr.Top = 40; $Global:txtMbxAddr.Left = 120; $Global:txtMbxAddr.Width = 220;  
        $Global:txtMbxAddr.Text = $Global:Address   # Email Address
        $Global:form.Controls.Add($Global:txtMbxAddr)    # Add to Form 

    ## Legacyddress
    $Global:lblLegAddr = New-Object System.Windows.Forms.Label   
        $Global:lblLegAddr.Text = "Additional Aliases:"  
        $Global:lblLegAddr.Top = 70; $Global:lblLegAddr.Left = 10; $Global:lblLegAddr.Width=150 ;$Global:lblLegAddr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblLegAddr)    # Add to Form 
        # 
        $Global:txtLegAddr = New-Object Windows.Forms.TextBox  
        $Global:txtLegAddr.TabIndex = 0 # set Tab Order 
        $Global:txtLegAddr.Top = 70; $Global:txtLegAddr.Left = 120; $Global:txtLegAddr.Width = 520;  
        $Global:txtLegAddr.Text = ""   # Legacy Address
        $Global:form.Controls.Add($Global:txtLegAddr)    # Add to Form 

    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 100; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtInpTaskNo.Top = 100; $Global:txtInpTaskNo.Left = 120; $Global:txtInpTaskNo.Width = 120;  
        $Global:txtInpTaskNo.Text = $Global:TaskNo   # Enter ticket number
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form

    ## Mailbox Owner
    $Global:lblGrpOwnr = New-Object System.Windows.Forms.Label   
        $Global:lblGrpOwnr.Text = "Mailbox Owner(s):"  
        $Global:lblGrpOwnr.Top = 130; $Global:lblGrpOwnr.Left = 10; $Global:lblGrpOwnr.Width=150; $Global:lblGrpOwnr.AutoSize = $true 
        $Global:form.Controls.Add($lblGrpOwnr)    # Add to Form 
        # 
        $Global:txtGrpOwnr = New-Object Windows.Forms.TextBox  
        $Global:txtGrpOwnr.TabIndex = 0 # set Tab Order 
        $Global:txtGrpOwnr.Top = 130; $Global:txtGrpOwnr.Left = 120; $Global:txtGrpOwnr.Width = 520;  
        $Global:txtGrpOwnr.Text = ""   # Use Corrent computer name as default 
        $Global:form.Controls.Add($Global:txtGrpOwnr)    # Add to Form 

       $Top = 160

   If ($Global:chkEditors.Checked -eq "Checked")
    {
        ## Editor Group Name
        $Global:lblEDGrpName = New-Object System.Windows.Forms.Label
        $Global:lblEDGrpName.Text = ".ED Group Name:"  
        $Global:lblEDGrpName.Top = $Top ; $Global:lblEDGrpName.Left = 10; $Global:lblEDGrpName.Width=150; $Global:lblEDGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblEDGrpName)    # Add to Form 
        #
        $Global:txtEDGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtEDGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtEDGrpName.Top = $Top; $Global:txtEDGrpName.Left = 120; $Global:txtEDGrpName.Width = 180;  
        $Global:txtEDGrpName.Text = ("MBX." + $Global:DispName + ".ED")   # Use Corrent computer name as default 
        $Global:form.Controls.Add($Global:txtEDGrpName)    # Add to Form 

        ## Editor Group Address
        $Global:lblEDGrpAddr = New-Object System.Windows.Forms.Label   
        $Global:lblEDGrpAddr.Text = ".ED Group Address:"  
        $Global:lblEDGrpAddr.Top = $Top; $Global:lblEDGrpAddr.Left = 310; $Global:lblEDGrpAddr.Width=150; $Global:lblEDGrpAddr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblEDGrpAddr)    # Add to Form 
        #
        $Global:txtEDGrpAddr = New-Object Windows.Forms.TextBox  
        $Global:txtEDGrpAddr.TabIndex = 0 # set Tab Order
        $Global:txtEDGrpAddr.Top = $Top; $Global:txtEDGrpAddr.Left = 430; $Global:txtEDGrpAddr.Width = 210;
        $Global:txtEDGrpAddr.Text = ("MBX." + $Global:GrpAddr + ".ED@ul.com")
#        If ($Global:chkNFP.checked -eq "Checked")
#        {
#            $Global:txtEDGrpAddr.Text = ("MBX." + $Global:GrpAddr + ".ED@ul.org")
#        }
        $Global:form.Controls.Add($Global:txtEDGrpAddr)    # Add to Form 

        $Top = $Top + 30
        ## Editor Group Members
        $Global:lblEDGrpMbr = New-Object System.Windows.Forms.Label   
        $Global:lblEDGrpMbr.Text = ".ED Member(s):"  
        $Global:lblEDGrpMbr.Top = $Top; $Global:lblEDGrpMbr.Left = 10; $Global:lblEDGrpMbr.Width=150; $Global:lblEDGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblEDGrpMbr)    # Add to Form 
        # 
        $Global:txtEDGrpMbr = New-Object Windows.Forms.TextBox
        $Global:txtEDGrpMbr.TabIndex = 0 # set Tab Order 
        $Global:txtEDGrpMbr.Location = New-Object System.Drawing.Size(120,$Top)
        $Global:txtEDGrpMbr.Size = New-Object system.Drawing.Size(520,60)
        $Global:txtEDGrpMbr.MultiLine = $true
        $Global:txtEDGrpMbr.ScrollBars = 'Both'  
        $Global:form.Controls.Add($Global:txtEDGrpMbr)    # Add to Form 

        $Top = $Top + 70
    }

    If ($Global:chkAuthors.Checked -eq "Checked")
    {
        ## Author Group Name
        $Global:lblAUGrpName = New-Object System.Windows.Forms.Label
        $Global:lblAUGrpName.Text = ".AU Group Name:" 
        $Global:lblAUGrpName.Top = $Top; $Global:lblAUGrpName.Left = 10; $Global:lblAUGrpName.Width=150; $Global:lblAUGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAUGrpName)    # Add to Form 
        #
        $Global:txtAUGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtAUGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtAUGrpName.Top = $Top; $Global:txtAUGrpName.Left = 120; $Global:txtAUGrpName.Width = 180;  
        $Global:txtAUGrpName.Text = ("MBX." + $Global:DispName + ".AU")   # Use Corrent computer name as default 
        $Global:form.Controls.Add($Global:txtAUGrpName)    # Add to Form 

        ## Author Group Address
        $Global:lblAUGrpAddr = New-Object System.Windows.Forms.Label   
        $Global:lblAUGrpAddr.Text = ".AU Group Address:"  
        $Global:lblAUGrpAddr.Top = $Top; $Global:lblAUGrpAddr.Left = 310; $Global:lblAUGrpAddr.Width=150; $Global:lblAUGrpAddr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAUGrpAddr)    # Add to Form 
        #
        $Global:txtAUGrpAddr = New-Object Windows.Forms.TextBox  
        $Global:txtAUGrpAddr.TabIndex = 0 # set Tab Order
        $Global:txtAUGrpAddr.Top = $Top; $Global:txtAUGrpAddr.Left = 430; $Global:txtAUGrpAddr.Width = 210;
        $Global:txtAUGrpAddr.Text = ("MBX." + $Global:GrpAddr + ".AU@ul.com")
#        If ($Global:chkNFP.checked -eq "Checked")
#        {
#            $Global:txtAUGrpAddr.Text = ("MBX." + $Global:GrpAddr + ".AU@ul.org")
#        }
        $Global:form.Controls.Add($Global:txtAUGrpAddr)    # Add to Form 

        $Top = $Top + 30
        ## Author Group Members
        $Global:lblAUGrpMbr = New-Object System.Windows.Forms.Label   
        $Global:lblAUGrpMbr.Text = ".AU Member(s):"  
        $Global:lblAUGrpMbr.Top = $Top; $Global:lblAUGrpMbr.Left = 10; $Global:lblAUGrpMbr.Width=150; $Global:lblAUGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblAUGrpMbr)    # Add to Form 
        # 
        $Global:txtAUGrpMbr = New-Object Windows.Forms.TextBox
        $Global:txtAUGrpMbr.TabIndex = 0 # set Tab Order 
        $Global:txtAUGrpMbr.Location = New-Object System.Drawing.Size(120,$Top)
        $Global:txtAUGrpMbr.Size = New-Object system.Drawing.Size(520,60)
        $Global:txtAUGrpMbr.MultiLine = $true
        $Global:txtAUGrpMbr.ScrollBars = 'Both'  
        $Global:form.Controls.Add($Global:txtAUGrpMbr)    # Add to Form 

        $Top = $Top + 70
    }

    If ($Global:chkReaders.Checked -eq "Checked")
    {
        ## Reader Group Name
        $Global:lblREGrpName = New-Object System.Windows.Forms.Label
        $Global:lblREGrpName.Text = ".RE Group Name:" 
        $Global:lblREGrpName.Top = $Top ; $Global:lblREGrpName.Left = 10; $Global:lblREGrpName.Width=150; $Global:lblREGrpName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblREGrpName)    # Add to Form 
        #
        $Global:txtREGrpName = New-Object Windows.Forms.TextBox  
        $Global:txtREGrpName.TabIndex = 0 # set Tab Order 
        $Global:txtREGrpName.Top = $Top; $Global:txtREGrpName.Left = 120; $Global:txtREGrpName.Width = 180;  
        $Global:txtREGrpName.Text = ("MBX." + $Global:DispName + ".RE")
        $Global:form.Controls.Add($Global:txtREGrpName)    # Add to Form 
       # Obtain Value with: $txtGrpOwnr.Text

        ## Reader Group Address
        $Global:lblREGrpAddr = New-Object System.Windows.Forms.Label   
        $Global:lblREGrpAddr.Text = ".RE Group Address:"  
        $Global:lblREGrpAddr.Top = $Top; $Global:lblREGrpAddr.Left = 310; $Global:lblREGrpAddr.Width=150; $Global:lblREGrpAddr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblREGrpAddr)    # Add to Form 
        #
        $Global:txtREGrpAddr = New-Object Windows.Forms.TextBox  
        $Global:txtREGrpAddr.TabIndex = 0 # set Tab Order
        $Global:txtREGrpAddr.Top = $Top; $Global:txtREGrpAddr.Left = 430; $Global:txtREGrpAddr.Width = 210;
        $Global:txtREGrpAddr.Text = ("MBX." + $Global:GrpAddr + ".RE@ul.com")
#        If ($Global:chkNFP.checked -eq "Checked")
#        {
#            $Global:txtREGrpAddr.Text = ("MBX." + $Global:GrpAddr + ".RE@ul.org")
#        }
        $Global:form.Controls.Add($Global:txtREGrpAddr)    # Add to Form 

        $Top = $Top + 30
        ## Reader Group Members
        $Global:lblREGrpMbr = New-Object System.Windows.Forms.Label   
        $Global:lblREGrpMbr.Text = ".RE Member(s):"  
        $Global:lblREGrpMbr.Top = $Top; $Global:lblREGrpMbr.Left = 10; $Global:lblREGrpMbr.Width=150; $Global:lblREGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblREGrpMbr)    # Add to Form 
        # 
        $Global:txtREGrpMbr = New-Object Windows.Forms.TextBox
        $Global:txtREGrpMbr.TabIndex = 0 # set Tab Order 
        $Global:txtREGrpMbr.Location = New-Object System.Drawing.Size(120,$Top)
        $Global:txtREGrpMbr.Size = New-Object system.Drawing.Size(520,60)
        $Global:txtREGrpMbr.MultiLine = $true
        $Global:txtREGrpMbr.ScrollBars = 'Both'  
        $Global:form.Controls.Add($Global:txtREGrpMbr)    # Add to Form 
    }
}

function Build-MbxName
{
    #Adjust name to proper case ann replace acronyms
    $Global:Name = (Get-Culture).textinfo.totitlecase($Global:txtInpName.Text.ToLower())
    $Space = $Global:Name.IndexOf(" ")
    $NameLen = $Global:Name.Length
    $MbxLoc = $Global:Name.Substring(0,$Global:Name.IndexOf(" "))
    $MbxName = $Global:Name.Substring($Space+1,($NameLen-($Space+1)))
    $MbxName = ((Get-Culture).textinfo.totitlecase($MbxName.ToLower()))

    foreach ($Acro in $Acro)
    {
        $MbxName = $MbxName -Replace($Acro.Acronym,$Acro.Translation)
    }

    If ($MbxLoc -notlike "Global*")
    {
        $Global:DispName = ($MbxLoc).ToUpper() + " "  + $MbxName
        $Global:Address = ($MbxLoc).ToUpper() + "." + $MbxName.Replace(" ","") + "@ul.com"
        If ($Global:chkNFP.checked -eq "Checked")
        {
            $Global:Address = ($MbxLoc).ToUpper() + "." + $MbxName.Replace(" ","") + "@ul.org"
        }
        $Global:GrpAddr = ($MbxLoc).ToUpper() + "." + $MbxName.Replace(" ","")
    }
    else
    {
        $Global:DispName = ((Get-Culture).textinfo.totitlecase($MbxLoc.ToLower())) + " "  + $MbxName
        $Global:Address = ((Get-Culture).textinfo.totitlecase($MbxLoc.ToLower())) + "." + $MbxName.Replace(" ","") + "@ul.com"
        If ($Global:chkNFP.checked -eq "Checked")
        {
            $Global:Address = ((Get-Culture).textinfo.totitlecase($MbxLoc.ToLower())) + "." + $MbxName.Replace(" ","") + "@ul.org"
        }
        $Global:GrpAddr = ((Get-Culture).textinfo.totitlecase($MbxLoc.ToLower())) + "." + $MbxName.Replace(" ","")
    }
}

$Global:OKDetails = ""
$Global:form = ""
$Global:txtEDGrpName = ""
$Global:txtAUGrpName = ""
$Global:txtREGrpName = ""
$Acro = Import-Csv e:\O365AdminShared\Data\KnownAcronyms.csv

Add-Type -Assembly System.Windows.Forms     ## Load the Windows Forms assembly 
## Create the main form

Build-ShrMbxInputForm
Publish-Form

If ($Global:Result -eq "OK")
{
    #Adjust name to proper case ann replace acronyms
    $Global:Name = (Get-Culture).textinfo.totitlecase($Global:txtInpName.Text.ToLower())
#    $Global:Name = (Get-Culture).textinfo.totitlecase($Global:txtDispName.Text.ToLower())
    $Global:Addr = ((Get-Culture).textinfo.totitlecase($Global:txtInpName.Text)).replace(" ","")
#    $Global:Addr = ((Get-Culture).textinfo.totitlecase($Global:txtDispName.Text)).replace(" ","")
    $Space = $Global:Name.IndexOf(" ")
    $NameLen = $Global:Name.Length
    $MbxLoc = $Global:Name.Substring(0,$Space)
    $MbxName = $Global:Name.Substring($Space+1,($NameLen-($Space+1)))
    $MbxName = ((Get-Culture).textinfo.totitlecase($MbxName.ToLower()))

    $Global:OKDetails = "Create"
    Build-MbxName

    $Global:TaskNo = $Global:txtInpTaskNo.Text
    Build-MbxDetailsForm
    Publish-Form
}
