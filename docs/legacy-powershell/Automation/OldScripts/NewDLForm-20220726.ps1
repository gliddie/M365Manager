<#   
================================================================================ 
 Name: NewDistributionList Form
 ================================================================================

 11/11/2020 - SAG - Added code so the the UL.RS and UL.IMS locations have "." only after the "UL" portion of the name
 01/13/2021 - SAG - In the DefaultForm change the addition of the buttonts to use the Add-FormStandardButtons function
 03/18/2022 - SAG - Change the Remove Date to use a DatePicker
#>  

function DefaultForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "New Distribution List" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(740,430) #(W,H)

    $Global:OKDetails = "Create"
    Add-FormStandardButtons
     
    ## Label and TextBox  
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $lblDispName.Text = "Display Name:"
        $lblDispName.Top = 10 ; $lblDispName.Left = 10; $lblDispName.Width=120 ;$lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($lblDispName)    # Add to Form 
        # 
        $Global:txtDispName = New-Object Windows.Forms.TextBox  
        $Global:txtDispName.TabIndex = 0 # set Tab Order 
        $Global:txtDispName.Top = 10; $txtDispName.Left = 120; $txtDispName.Width = 500;  
        $Global:txtDispName.Text = $DLName   # DisplayName
        $Global:form.Controls.Add($Global:txtDispName)    # Add to Form
        $Global:InputFocus = $Global:txtDispName
       # Obtain Value with: $txtDispName.Text
 
    ## EmailAddress
    $Global:lblGrpAddr = New-Object System.Windows.Forms.Label   
        $Global:lblGrpAddr.Text = "Email Address:"  
        $Global:lblGrpAddr.Top = 40 ; $Global:lblGrpAddr.Left = 10; $Global:lblGrpAddr.Width=150 ;$Global:lblGrpAddr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblGrpAddr)    # Add to Form 
        # 
        $Global:txtGrpAddr = New-Object Windows.Forms.TextBox  
        $Global:txtGrpAddr.TabIndex = 0 # set Tab Order 
        $Global:txtGrpAddr.Top = 40; $Global:txtGrpAddr.Left = 120; $Global:txtGrpAddr.Width = 500;
        $DLAddr = $DLName -Replace '[ &/,$#_-]',''
        If ($Global:chkNFP.Checked -ne "Checked")
        {
            $Global:txtGrpAddr.Text = $DLAddr + "@ul.com"   # Email Address
        }
        else
        {
            $Global:txtGrpAddr.Text = $DLAddr + "@ul.org"   # Email Address
        }
#        $Global:txtGrpAddr.Text = (((Get-Culture).textinfo.totitlecase($DLName)) -Replace '[ &/,$#_-]','') + "@ul.com"   # Email Address
        $Global:form.Controls.Add($Global:txtGrpAddr)    # Add to Form 
       # Obtain Value with: $Global:txtGrpAddr.Text 

    ## Legacyddress
    $Global:lblLegAddr = New-Object System.Windows.Forms.Label   
        $lblLegAddr.Text = "Additional Aliases:"  
        $lblLegAddr.Top = 70 ; $lblLegAddr.Left = 10; $lblLegAddr.Width=150 ;$lblLegAddr.AutoSize = $true 
        $Global:form.Controls.Add($lblLegAddr)    # Add to Form 
        # 
        $Global:txtLegAddr = New-Object Windows.Forms.TextBox  
        $txtLegAddr.TabIndex = 0 # set Tab Order 
        $txtLegAddr.Top = 70; $txtLegAddr.Left = 120; $txtLegAddr.Width = 500;  
        $txtLegAddr.Text = ""   # Legacy Address
        $Global:form.Controls.Add($txtLegAddr)    # Add to Form 
       # Obtain Value with: $txtLegAddr.Text 

    ## Group Owner
    $Global:lblGrpOwnr = New-Object System.Windows.Forms.Label   
        $Global:lblGrpOwnr.Text = "Group Owner(s):"  
        $Global:lblGrpOwnr.Top = 100 ; $Global:lblGrpOwnr.Left = 10; $Global:lblGrpOwnr.Width=150; $Global:lblGrpOwnr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblGrpOwnr)    # Add to Form 
        # 
        $Global:txtGrpOwnr = New-Object Windows.Forms.TextBox  
        $Global:txtGrpOwnr.TabIndex = 0 # set Tab Order 
        $Global:txtGrpOwnr.Top = 100; $Global:txtGrpOwnr.Left = 120; $Global:txtGrpOwnr.Width = 500;  
        $Global:txtGrpOwnr.Text = ""   # Use Corrent computer name as default 
        $Global:form.Controls.Add($Global:txtGrpOwnr)    # Add to Form 
       # Obtain Value with: $Global:txtGrpOwnr.Text        

    ## Group Members
    $Global:lblGrpMbr = New-Object System.Windows.Forms.Label   
        $Global:lblGrpMbr.Text = "Group Member(s):"  
        $Global:lblGrpMbr.Top = 130 ; $Global:lblGrpMbr.Left = 10; $Global:lblGrpMbr.Width=150; $Global:lblGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblGrpMbr)    # Add to Form 
        # 
        $Global:txtGrpMbr = New-Object Windows.Forms.TextBox
        $Global:txtGrpMbr.MaxLength = 2000000
        $Global:txtGrpMbr.TabIndex = 0 # set Tab Order 
        $Global:txtGrpMbr.Location = New-Object System.Drawing.Size(120,130)
        $Global:txtGrpMbr.Size = New-Object system.Drawing.Size(500,60)
        $Global:txtGrpMbr.MultiLine = $true
        $Global:txtGrpMbr.ScrollBars = 'Both'  
        $Global:form.Controls.Add($Global:txtGrpMbr)    # Add to Form 
       # Obtain Value with: $Global:txtGrpMbr.Text  

    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 200 ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=150 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtTaskNo.Top = 200; $Global:txtTaskNo.Left = 120; $Global:txtTaskNo.Width = 120;  
        $Global:txtTaskNo.Text = $Global:txtInpTaskNo.Text   # Enter ticket number 
        $Global:form.Controls.Add($Global:txtTaskNo)    # Add to Form 
       # Obtain Value with: $txtTaskNo.Text

    $Top = 230
    If (($Global:chkTemp.Checked -eq "Checked") -or ($Global:txtInpName.Text -like "*.TMP.*"))
    {
        ## Date to remove
        $Global:lblRemoveDate = New-Object System.Windows.Forms.Label   
            $Global:lblRemoveDate.Text = "Remove On:"  
            $Global:lblRemoveDate.Top = $Top ; $lblRemoveDate.Left = 10; $lblRemoveDate.Width=150 ; $lblRemoveDate.AutoSize = $true
            $Global:form.Controls.Add($lblRemoveDate)    # Add to Form 
            # 
            $Global:txtRemoveDate = New-Object Windows.Forms.DateTimePicker
            $Global:txtRemoveDate.TabIndex = 0 # set Tab Order 
            $Global:txtRemoveDate.Top = $Top; $txtRemoveDate.Left = 120; $txtRemoveDate.Width = 120;
            $Global:txtRemoveDate.Format = [windows.forms.datetimepickerFormat]::custom
            $Global:txtRemoveDate.CustomFormat = "MM/dd/yyyy"
            $Global:txtRemoveDate.Text = (get-date)
            $Global:form.Controls.Add($Global:txtRemoveDate)    # Add to Form 
           # Obtain Value with: $txtRemoveDate.Text
        $Top = 260
    }   

    ## Allow External Senders CheckBox
    $Global:chkExtSdrs = New-Object Windows.Forms.checkbox 
        $Global:chkExtSdrs.Left = 120; $chkExtSdrs.Width = 280; $chkExtSdrs.Top = $Top  
        $Global:chkExtSdrs.Text = "Allow Use by External Senders" 
        $Global:chkExtSdrs.Checked = $false   # set a default value 
        $Global:chkExtSdrs.TabIndex = 2
        $Global:form.Controls.Add($chkExtSdrs) 
        # Obtain Value with: $chkExtSdrs.Checked

    ## Not for Display Box to Not Display the email template
    $Global:chkTemplate = New-Object Windows.Forms.checkbox 
        $Global:chkTemplate.Left = 120; $Global:chkTemplate.Width = 300; $Global:chkTemplate.Top = $Top + 30
        $Global:chkTemplate.Text = "Display Email Template for New Distribution Group" 
        $Global:chkTemplate.Checked = $true   # set a default value
        $Global:form.Controls.Add($Global:chkTemplate)
}

function Build-DLInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "New Distribution Group" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(580,250) #(W,H)
#    $Global:okButton.Text = "Continue"

    Add-FormStandardButtons

    ## Get Details used to create (O365 Name max length 64 characters)
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Distribution List Name:"
        $Global:lblDispName.Top = 10 ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtInpName = New-Object Windows.Forms.TextBox  
        $Global:txtInpName.TabIndex = 0 # set Tab Order 
        $Global:txtInpName.Top = 10; $Global:txtInpName.Left = 130; $Global:txtInpName.Width = 350;
        $Global:txtInpName.MaxLength = 64
        $Global:txtInpName.Text = ""   # DisplayName
        $Global:form.Controls.Add($Global:txtInpName)    # Add to Form 
        $Global:InputFocus = $Global:txtInpName
 
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 40 ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtInpTaskNo.Top = 40; $Global:txtInpTaskNo.Left = 130; $Global:txtInpTaskNo.Width = 120;  
        $Global:txtInpTaskNo.Text = "TASK"   # Enter ticket number 
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form 

    ## Temporary Group CheckBox
    $Global:chkTemp = New-Object Windows.Forms.checkbox 
        $Global:chkTemp.Left = 130; $Global:chkTemp.Width = 200; $Global:chkTemp.Top = 70  
        $Global:chkTemp.Text = "Temporary Group" 
        $Global:chkTemp.Checked = $false   # set a default value
        $Global:chkTemp.TabIndex = 2
        $Global:form.Controls.Add($Global:chkTemp)

    ## Not for Profit Group CheckBox
    $Global:chkNFP = New-Object Windows.Forms.checkbox 
        $Global:chkNFP.Left = 130; $Global:chkNFP.Width = 200; $Global:chkNFP.Top = 100  
        $Global:chkNFP.Text = "Requested by Not For Profit Group" 
        $Global:chkNFP.Checked = $false   # set a default value
        $Global:chkNFP.TabIndex = 2
        $Global:form.Controls.Add($Global:chkNFP)
        
    ## Remove Special Character CheckBox
    $Global:chkSpecChara = New-Object Windows.Forms.checkbox 
        $Global:chkSpecChara.Left = 130; $Global:chkSpecChara.Width = 230; $Global:chkSpecChara.Top = 130  
        $Global:chkSpecChara.Text = "Remove Special Character from Group: " 
        $Global:chkSpecChara.Checked = $false   # set a default value
        $Global:chkSpecChara.TabIndex = 3
        $Global:form.Controls.Add($Global:chkSpecChara) 
        $Global:txtSpecChara = New-Object Windows.Forms.TextBox  
        $Global:txtSpecChara.TabIndex = 4 # set Tab Order 
        $Global:txtSpecChara.Top = 130; $Global:txtSpecChara.Left = 360; $Global:txtSpecChara.Width = 20;
#        $Global:txtSpecChara.MaxLength = 64
        $Global:txtSpecChara.Text = ""   # DisplayName
        $Global:form.Controls.Add($Global:txtSpecChara)
}

$Global:form = ""
$Global:lblLegAddr = ""
$Global:txtLegAddr = ""
$Global:lblGrpOwnr = ""
$Global:txtGrpOwnr = ""
$Global:lblGrpMbr = ""
$Global:txtGrpMbr = ""
$Global:lblTaskNo = ""
$Global:txtTaskNo = ""
$Global:chkExtSdrs = ""

Add-Type -Assembly System.Windows.Forms     ## Load the Windows Forms assembly 
## Create the main form

Build-DLInputForm
Publish-Form

If ($Global:Result -eq "OK")
{
    $DLName = ((Get-Culture).textinfo.totitlecase(($Global:txtInpName.Text).ToLower()))

    If (($DLName -like "LST.UL.RS.*") -or ($DLName -like "LST.UL.IMS.*") -or ($DLName -like "LST.Commercial.Operations*"))
    {
        If ($DLName -like "*UL.RS.*")
        {
            $DLName = $DLName -Replace("UL.RS.","UL.RS ")
        }
        elseif ($DLName -like "*UL.IMS.*")
        {
            $DLName = $DLName -Replace("UL.IMS.","UL.IMS ")
        }
        else
        {
            $DLName = $DLName -Replace("LST.COMMERCIAL.OPERATIONS","LST.CommercialOperations")
        }
    }

    $DLName1 = $DLName.Split(".")
    $Name = $DLName.Split(".")
    $cnt = 0
    Foreach ($DLName1 in $DLName1)
    {
        If ($cnt -eq "0")
        {
            $Name[0] = $DLName1.ToUpper()
        }
        elseif ($cnt -eq 1)
        {
            If (($Global:chkTemp.Checked -eq "Checked") -and ($DLName1 -ne "TMP"))
            {
                If ($DLName1.length -eq 3)
                {
                    $DLName1 = $DLName1.ToUpper()
                }
                $Name[2] = $DLName1 + " " + $Name[2]
                $Name[1] = "TMP"
            }
            else
            {
                If ($DLName1 -like "EMEA*")
                {
                    $Name[1] = $DLName1.ToUpper()
                }
                else
                {            
                    If ($DLName1 -notlike "*CommercialOperations*")
                    {
                        $Name[1] = If ($DLName1.Length -le 3) {$DLName1.ToUpper()} else {$DLName1}
                    }
                    else
                    {
                        $Name[1] = "CommercialOperations"
                    }
                }
            }
        }
        else
        {
            If (($Global:chkTemp.Checked -ne "Checked") -or ($cnt -gt 2))
            {
                $Name[$cnt] = $DLName1
            }
            #Check for Known acronyms

            foreach ($Acro in $Collection)
            {
                $Name[$cnt] = $Name[$cnt] -Replace($Acro.Acronym,$Acro.Translation)
            }
        }
        $Cnt++
    }

    $cnt = 0
    foreach ($Name in $Name)
    {
        if ($cnt -eq 0)
        {
            $DLName = $Name
        }

        if ($cnt -eq 1)
        {
            $DLName = $DLName + "." + $Name
        }

        if ($Cnt -eq 2)
        {
            If ($Global:chkSpecChara.Checked -eq "Checked")
            {
                $Name = ($Name -Replace($Global:txtSpecChara.Text," "))
            }
            $DLName = $DLName + "." + $Name
        }

        If ($cnt -gt 2)
        {
            $DLName = $DLName + " " + $Name
        }
        $Cnt++
    }

    DefaultForm
    Publish-Form
}
