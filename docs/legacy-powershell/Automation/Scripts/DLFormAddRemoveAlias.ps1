<#   
================================================================================ 
 Name: AddRemoveAlias Form
 ================================================================================ 
#>  

function Build-DLFormDetails
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Add/Remove Distribution Group Alias" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(740,380) #(W,H)

    Add-FormStandardButtons
     
    $TopLoc = 10
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Display Name:"
        $Global:lblDispName.Top = $TopLoc; $lblDispName.Left = 10; $lblDispName.Width=120 ;$lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtDispName = New-Object Windows.Forms.TextBox 
        $Global:txtDispName.ReadOnly = "true";
        $Global:txtDispName.TabIndex = 99
        $Global:txtDispName.Top = $TopLoc; $txtDispName.Left = 120; $txtDispName.Width = 500;  
        $Global:txtDispName.Text = $Global:DLInf.DisplayName
        $Global:form.Controls.Add($Global:txtDispName)    # Add to Form 
       # Obtain Value with: $txtDispName.Text
 
    $TopLoc = $TopLoc + 30
    ## EmailAddress
     $Global:lblGrpAddr = New-Object System.Windows.Forms.Label   
        $Global:lblGrpAddr.Text = "Primary Address:"
        $Global:lblGrpAddr.Top = $TopLoc; $Global:lblGrpAddr.Left = 10; $Global:lblGrpAddr.Width=150 ;$Global:lblGrpAddr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblGrpAddr)    # Add to Form 
        # 
        $Global:txtGrpAddr = New-Object Windows.Forms.TextBox  
        $Global:txtGrpAddr.ReadOnly = "true";
        $Global:txtGrpAddr.TabIndex = 99  
        $Global:txtGrpAddr.Top = $TopLoc; $Global:txtGrpAddr.Left = 120; $Global:txtGrpAddr.Width = 500;
        $Global:txtGrpAddr.Text = $Global:DLInf.PrimarySmtpAddress
        $Global:form.Controls.Add($Global:txtGrpAddr)    # Add to Form 
       # Obtain Value with: $Global:txtGrpAddr.Text 

    $TopLoc = $TopLoc + 30
    ## Legacyddress
    $Global:lblLegAddr = New-Object System.Windows.Forms.Label   
        $Global:lblLegAddr.Text = "Other Aliases:"  
        $Global:lblLegAddr.Top = $TopLoc; $Global:lblLegAddr.Left = 10; $Global:lblLegAddr.Width=150 ;$Global:lblLegAddr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblLegAddr)    # Add to Form 
        # 
        $Global:txtLegAddr = New-Object Windows.Forms.TextBox
        $Global:txtLegAddr.ReadOnly = "true";
        $Global:txtLegAddr.TabIndex = 99  
        $Global:txtLegAddr.Top = $TopLoc; $Global:txtLegAddr.Left = 120; $Global:txtLegAddr.Width = 500;
        $Aliases = "(None)"
        If ($Global:DLInf.EmailAddresses.count -gt 1)
        {
            Foreach ($Addr in $Global:DLInf.EmailAddresses)
            {
                If ($Addr -notlike "*" + $Global:DLInf.PrimarySmtpAddress)
                {
                    If ($Aliases -eq "(None)")
                    {
                        $Aliases = $Addr
                    }
                    else
                    {
                        $Aliases = $Aliases + ", " + $Addr
                    }
                }
            }
        }
        $Global:txtLegAddr.Text = $Aliases  # Legacy Address
        $Global:form.Controls.Add($Global:txtLegAddr)    # Add to Form 
       # Obtain Value with: $Global:txtLegAddr.Text 

    $TopLoc = $TopLoc + 30
    ## Group Owner
    $Global:lblGrpOwnr = New-Object System.Windows.Forms.Label   
        $Global:lblGrpOwnr.Text = "Group Owner(s):"  
        $Global:lblGrpOwnr.Top = $TopLoc; $Global:lblGrpOwnr.Left = 10; $Global:lblGrpOwnr.Width=150; $Global:lblGrpOwnr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblGrpOwnr)    # Add to Form 
        # 
        $Global:txtGrpOwnr = New-Object Windows.Forms.TextBox 
        $Global:txtGrpOwnr.ReadOnly = "true";
        $Global:txtGrpOwnr.TabIndex = 99
        $Global:txtGrpOwnr.Top = $TopLoc; $Global:txtGrpOwnr.Left = 120; $Global:txtGrpOwnr.Width = 500;
        $Global:txtGrpOwnr.Text = $Global:DLInf.ManagedBy -join ", "   # Use Corrent computer name as default 
        $Global:form.Controls.Add($Global:txtGrpOwnr)    # Add to Form 
       # Obtain Value with: $Global:txtGrpOwnr.Text        

    $TopLoc = $TopLoc + 30
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = $TopLoc; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=150 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtTaskNo = New-Object Windows.Forms.TextBox
        $Global:txtTaskNo.TabIndex = 98
        $Global:txtTaskNo.Top = $TopLoc; $Global:txtTaskNo.Left = 120; $Global:txtTaskNo.Width = 120;  
        $Global:txtTaskNo.Text = $Global:txtInpTaskNo.Text   # Enter ticket number 
        $Global:form.Controls.Add($Global:txtTaskNo)    # Add to Form 
       # Obtain Value with: $txtTaskNo.Text

    $TopLoc = $TopLoc + 40
    ## Authorized Requestor
    $Global:chkAuthorized = New-Object Windows.Forms.checkbox 
        $Global:chkAuthorized.Left = 120; $Global:chkAuthorized.Width = 200; $Global:chkAuthorized.Top = $TopLoc
        $Global:chkAuthorized.Text = "Authorized Requestor" 
        $Global:chkAuthorized.Checked = $false   # set a default value 
        $Global:chkAuthorized.TabIndex = 0
        $Global:form.Controls.Add($Global:chkAuthorized) 
        # Obtain Value with: $Global:chkAuthorized.Checked

    $TopLoc = $TopLoc + 30

    If ($Global:chkAddAlias.Checked -eq "Checked")
    {
        ## New Alias
        $Global:lblNewAlias = New-Object System.Windows.Forms.Label   
            $Global:lblNewAlias.Text = "New Email Alias:"  
            $Global:lblNewAlias.Top = $TopLoc; $Global:lblNewAlias.Left = 10; $Global:lblNewAlias.Width=150 ; $Global:lblNewAlias.AutoSize = $true
            $Global:form.Controls.Add($Global:lblNewAlias)    # Add to Form 
            # 
            $Global:txtNewAlias = New-Object Windows.Forms.TextBox  
            $Global:txtNewAlias.TabIndex = 2 # set Tab Order 
            $Global:txtNewAlias.Top = $TopLoc; $Global:txtNewAlias.Left = 120; $Global:txtNewAlias.Width = 350;  
            $Global:txtNewAlias.Text = ""   # Enter ticket number 
            $Global:form.Controls.Add($Global:txtNewAlias)    # Add to Form 
           # Obtain Value with: $txtNewAlias.Text
        $Global:chkPrimary = New-Object Windows.Forms.checkbox 
            $Global:chkPrimary.Left = 480; $Global:chkPrimary.Width = 150; $Global:chkPrimary.Top = ($TopLoc - 2)
            $Global:chkPrimary.Text = "Set as Primary Address" 
            $Global:chkPrimary.Checked = $false   # set a default value 
            $Global:chkPrimary.TabIndex = 3
            $Global:form.Controls.Add($Global:chkPrimary) 
            # Obtain Value with: $Global:chkPrimary.Checked
    }

    If ($Global:chkRemoveAlias.Checked -eq "Checked")
    {
        ## Remove Alias
        $Global:lblRemoveAlias = New-Object System.Windows.Forms.Label   
            $Global:lblRemoveAlias.Text = "Alias to Remove:"  
            $Global:lblRemoveAlias.Top = $TopLoc; $Global:lblRemoveAlias.Left = 10; $Global:lblRemoveAlias.Width=150 ; $Global:lblRemoveAlias.AutoSize = $true
            $Global:form.Controls.Add($Global:lblRemoveAlias)    # Add to Form 
            # 
            $Global:txtRemoveAlias = New-Object Windows.Forms.TextBox  
            $Global:txtRemoveAlias.TabIndex = 2 # set Tab Order 
            $Global:txtRemoveAlias.Top = $TopLoc; $Global:txtRemoveAlias.Left = 120; $Global:txtRemoveAlias.Width = 350;  
            $Global:txtRemoveAlias.Text = ""   # Enter ticket number 
            $Global:form.Controls.Add($Global:txtRemoveAlias)    # Add to Form 
           # Obtain Value with: $txtRemoveAlias.Text
    }
}

function Build-DLInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Add/Remove Distribution Group Alias" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(400,250) #(W,H)
    $BldDetails = "N"

    Add-FormStandardButtons

    ## Get Details used to create
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Distribution List Name:"
        $Global:lblDispName.Top = 20 ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtInpName = New-Object Windows.Forms.TextBox  
        $Global:txtInpName.TabIndex = 1 # set Tab Order 
        $Global:txtInpName.Top = 20; $Global:txtInpName.Left = 130; $Global:txtInpName.Width = 200;  
        $Global:txtInpName.Text = ""   # DisplayName
        $Global:form.Controls.Add($Global:txtInpName)    # Add to Form 
 
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 50 ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 2 # set Tab Order 
        $Global:txtInpTaskNo.Top = 50; $Global:txtInpTaskNo.Left = 130; $Global:txtInpTaskNo.Width = 120;  
        $Global:txtInpTaskNo.Text = "TASK"   # Enter ticket number 
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form 

    ## Add Alias Checkbox
    $Global:chkAddAlias = New-Object Windows.Forms.checkbox 
        $Global:chkAddAlias.Left = 130; $Global:chkAddAlias.Width = 150; $Global:chkAddAlias.Top = 80
        $Global:chkAddAlias.Text = "Add Alias" 
        $Global:chkAddAlias.Checked = $false   # set a default value 
        $Global:chkAddAlias.TabIndex = 3
        $Global:form.Controls.Add($Global:chkAddAlias) 
        # Obtain Value with: $Global:chkAddAlias.Checked

    ## Remove Alias Checkbox
    $Global:chkRemoveAlias = New-Object Windows.Forms.checkbox 
        $Global:chkRemoveAlias.Left = 130; $Global:chkRemoveAlias.Width = 150; $Global:chkRemoveAlias.Top = 100
        $Global:chkRemoveAlias.Text = "Remove Alias" 
        $Global:chkRemoveAlias.Checked = $false   # set a default value 
        $Global:chkRemoveAlias.TabIndex = 4
        $Global:form.Controls.Add($Global:chkRemoveAlias) 
        # Obtain Value with: $Global:chkRemoveAlias.Checked

}

$Global:form = ""
$Global:DLInf = ""
$Global:txtLegAddr = ""
$Global:txtGrpOwnr = ""
$Global:txtTaskNo = ""

Add-Type -Assembly System.Windows.Forms     ## Load the Windows Forms assembly 
## Create the main form

Build-DLInputForm
Publish-Form

If ($Global:Result -eq "OK")
{
    $DLExists = [bool](Get-DistributionGroup $Global:txtInpName.Text)
    If ($DLExists -eq $True)
    {
        $Global:DLInf = Get-DistributionGroup $Global:txtInpName.Text
    }

    If (($Global:DLInf.EmailAddresses.count -le 1) -and ($Global:chkRemoveAlias.Checked -eq "Checked"))
    {
        $Output = $wshell.Popup("Cannot remove an alias as there are no secondary aliases assigned to " +  $Global:txtInpName.Text,0,"No Additional Aliases",0+32) 
    }
    else
    {
        Build-DLFormDetails
        Publish-Form
    }
}

<#
    $Grp = $Global:txtInpName.Text
    $DstExists = [bool](Get-DistributionGroup $Grp -ErrorAction SilentlyContinue)
    if ($DstExists -eq "True")
    {
        $GrpAlias = Get-DistributionGroup $Grp                    
        write-host "Distribution List Owners on" $Grp ": " $GrpAlias.ManagedBy -ForegroundColor Yellow
            Write-Host "Current Email aliases on" $Grp "   : " -ForegroundColor Yellow -NoNewline
            write-host $GrpAlias.EmailAddresses 
            write-host ""
            write-host "Is the requestor an owner of this distribution list or have you obtained owner approval (Y/N)? " -ForegroundColor Red -NoNewline
            $OwnrAns = Read-Host
            If ($OwnrAns -eq "Y")
            {
                write-host "Do you wish to Add or Remove an Alias enter (A = Add/R = Remove)? " -ForegroundColor Yellow -NoNewline
                $AddRem = Read-Host
                If ($AddRem -eq "A")
                {
                    write-host ""
                    write-host "Enter the eMail Alias you would like to add: " -ForegroundColor Yellow -NoNewline
                    $NewAlias = Read-Host
                    $AddAlias = "smtp:" + $NewAlias
                    Set-DistributionGroup $Grp -EmailAddresses @{Add=$AddAlias}
                    Get-DistributionGroup $Grp |ft *Addresses*
                    Write-Host "Should this be made the new Primary SMTP Address for this Mailbox (Y/N) ?" -ForegroundColor Yellow -NoNewline
                    $NewPrim = Read-Host
                    If ($NewPrim -eq "Y")
                    {
                        Set-DistributionGroup $Grp -PrimarySmtpAddress $NewAlias
                    }
                }
                elseif ($AddRem -eq "R")
                {
                    write-host ""
                    write-host "Enter the eMail Alias you would like to remove: " -ForegroundColor Yellow -NoNewline
                    $RemAlias = Read-Host
                    $GrpDet = Get-DistributionGroup $Grp
                    If ($RemAlias -eq $GrpDet.PrimarySMTPAddress)
                    {
                        write-host "This is the Primary SMTP Alias for this group you must assign a new Primary Alias.  Enter New Primary Alias: " -ForegroundColor Yellow -NoNewline
                        $NewAlias = Read-Host
                        Set-DistributionGroup $Grp -PrimarySmtpAddress $NewAlias                                                       
                    }
                    $RemAlias = "smtp:" + $RemAlias
                    Set-DistributionGroup $Grp -EmailAddresses @{Remove=$RemAlias}
                }
                else
                {
                    Write-Host "No changes made to the configured aliases for this Distribution list"
                }
                Get-DistributionGroup $Grp |ft EmailAddresses
                pause
            }
            else
            {
                Write-Host "Please obtain approval to make this change" -BackgroundColor Red
                pause
            }
        }
        else
        {
             Write-Host "Distribution List Does Not Exist - No Changes Made" -ForegroundColor Red
        }
    }
#>