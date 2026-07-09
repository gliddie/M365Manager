<#   
================================================================================ 
 Name: Modify DL Authorized Users
 ================================================================================

 09/09/2022 - Created from the New DL script
 09/15/2022 - Added code to include changes to DST groups and added display of the email template
#>  

function Build-DefaultForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Modify Distribution List Authorized Users" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(600,200) #(W,H)

    $Global:OKDetails = "Update"
    Add-FormStandardButtons
    
    $Top = 10 
    ## Label and TextBox  
    ## Display Name
    $Script:lblDispName = New-Object System.Windows.Forms.Label   
        $lblDispName.Text = "Display Name:"
        $lblDispName.Top = $Top ; $lblDispName.Left = 10; $lblDispName.Width=120 ;$lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($lblDispName)    # Add to Form 
        # 
        $Script:txtDispName = New-Object Windows.Forms.TextBox  
        $Script:txtDispName.TabIndex = 0 # set Tab Order 
        $Script:txtDispName.Top = 10; $txtDispName.Left = 130; $txtDispName.Width = 300;  
        $Script:txtDispName.Text = ""  # DisplayName
        $Global:form.Controls.Add($Script:txtDispName)    # Add to Form
        $Global:InputFocus = $Script:txtDispName
        $Script:txtDispName.Add_Click({
            Hide-Fields
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
        $Script:txtTaskNo.TabIndex = 0 # set Tab Order
        $Script:txtTaskNo.Top = $Top; $Script:txtTaskNo.Left = 130; $Script:txtTaskNo.Width = 120;  
        $Script:txtTaskNo.Text = "TASK"   # Enter ticket number 
        $Global:form.Controls.Add($Script:txtTaskNo)    # Add to Form 

    #Get Group Details Button
    $Script:ButGetGrp = New-Object Windows.Forms.Button
    $Script:ButGetGrp.Location = New-object System.Drawing.Size(430,10)
    $Script:ButGetGrp.Size = new-Object System.Drawing.Size(110,20)
    $Script:ButGetGrp.Text = "Get Group Details"
    $Script:ButGetGrp.TabIndex = 1
    $Global:form.Controls.Add($Script:ButGetGrp)
    $Script:ButGetGrp.Add_Click({                
        $Script:ButGetGrp.visible = $false
        write-host "Retreiving Group Details for: $Script:Grp" -ForegroundColor Cyan
        $Script:Grp = $Script:txtDispName.Text.Trim(" ")
        $Script:txtReleaseby = ""
        $Script:Republish = "Y"
        If ($Script:Grp -notlike "DST*")
        {
            $Exists = [bool]($Script:GrpDetails = get-DistributionGroup $Script:Grp -ErrorAction SilentlyContinue)
        }
        else
        {
            $Exists = [bool]($Script:GrpDetails = get-DynamicDistributionGroup $Script:Grp -ErrorAction SilentlyContinue)
        }
        If ($Exists -eq $True)
        {
            $Script:txtDispName.Text = $Script:GrpDetails.DisplayName
            $Script:Grp = $Script:txtDispName.Text.Trim(" ")
            Complete-Form
        }
        else
        {
            $Script:txtDispName.Text = "Invalid"
            $Script:ButGetGrp.Visible = $True
            $Global:form.Size = New-Object System.Drawing.Size(600,200) #(W,H)
        }
    })
 
    $Top = $Top + 30
    ## EmailAddress
    $Script:lblGrpAddr = New-Object System.Windows.Forms.Label   
        $Script:lblGrpAddr.Text = "Email Address:"  
        $Script:lblGrpAddr.Top = $Top ; $Script:lblGrpAddr.Left = 10; $Script:lblGrpAddr.Width=150 ;$Script:lblGrpAddr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblGrpAddr)    # Add to Form 
        # 
        $Script:txtGrpAddr = New-Object Windows.Forms.TextBox  
        $Script:txtGrpAddr.Top = $Top; $Script:txtGrpAddr.Left = 130; $Script:txtGrpAddr.Width = 400;
        $Script:txtGrpAddr.ReadOnly = $true
        $Global:form.Controls.Add($Script:txtGrpAddr)    # Add to Form

    $Top = $Top + 30
    ## Legacyddress
    $Script:lblLegAddr = New-Object System.Windows.Forms.Label   
        $lblLegAddr.Text = "Additional Aliases:"  
        $lblLegAddr.Top = $Top; $lblLegAddr.Left = 10; $lblLegAddr.Width=150 ;$lblLegAddr.AutoSize = $true 
        $Global:form.Controls.Add($lblLegAddr)    # Add to Form 
        # 
        $Script:txtLegAddr = New-Object Windows.Forms.TextBox  
        $Script:txtLegAddr.Top = $Top; $Script:txtLegAddr.Left = 130; $Script:txtLegAddr.Width = 400;  
        $Script:txtLegAddr.Text = "Retrieving Details...."  # Legacy Address
        $Script:txtLegAddr.ReadOnly = $true
        $Global:form.Controls.Add($Script:txtLegAddr)    # Add to Form 

    $Top = $Top + 30
    ## Group Owner
    $Script:lblGrpOwnr = New-Object System.Windows.Forms.Label   
        $Script:lblGrpOwnr.Text = "Group Owner(s):"  
        $Script:lblGrpOwnr.Top = $Top; $Script:lblGrpOwnr.Left = 10; $Script:lblGrpOwnr.Width=150; $Script:lblGrpOwnr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblGrpOwnr)    # Add to Form 
        # 
        $Script:txtGrpOwnr = New-Object Windows.Forms.TextBox  
        $Script:txtGrpOwnr.Top = $Top; $Script:txtGrpOwnr.Left = 130; $Script:txtGrpOwnr.Width = 400;  
        $Script:txtGrpOwnr.Text = "Retrieving Details...."   # Use Corrent computer name as default
        $Script:txtGrpOwnr.ReadOnly = $true
        $Global:form.Controls.Add($Script:txtGrpOwnr)    # Add to Form 

    $Top = $Top + 30
    ## Group Members
    $Script:lblGrpMbr = New-Object System.Windows.Forms.Label   
        $Script:lblGrpMbr.Text = "Group Member(s):"  
        $Script:lblGrpMbr.Top = $Top ; $Script:lblGrpMbr.Left = 10; $Script:lblGrpMbr.Width=150; $Script:lblGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblGrpMbr)    # Add to Form 
        # 
        $Script:txtGrpMbr = New-Object Windows.Forms.TextBox
        $Script:txtGrpMbr.MaxLength = 200000000
        $Script:txtGrpMbr.Location = New-Object System.Drawing.Size(130,$Top)
        $Script:txtGrpMbr.Size = New-Object system.Drawing.Size(400,60)
        $Script:txtGrpMbr.Text = "Retrieving Details...." 
        $Script:txtGrpMbr.MultiLine = $true
        $Script:txtGrpMbr.ScrollBars = 'Both'
        $Script:txtGrpMbr.ReadOnly = $true  
        $Global:form.Controls.Add($Script:txtGrpMbr)    # Add to Form 
       
    $Top = $Top + 70
    ## Authorized Users
    $Script:lblAuthUser = New-Object System.Windows.Forms.Label   
        $Script:lblAuthUser.Text = "Authorized Users:"  
        $Script:lblAuthUser.Top = $Top; $Script:lblAuthUser.Left = 10; $Script:lblAuthUser.Width=150 ; $Script:lblAuthUser.AutoSize = $true
        $Global:form.Controls.Add($Script:lblAuthUser)    # Add to Form 
        # 
        $Script:txtAuthUser = New-Object Windows.Forms.TextBox
        $Script:txtAuthUser.Top = $Top; $Script:txtAuthUser.Left = 130; $Script:txtAuthUser.Width = 400;  
        $Script:txtAuthUser.Text = "Retrieving Details...."    # Enter ticket number
        $Script:txtAuthUser.ReadOnly = $true 
        $Global:form.Controls.Add($Script:txtAuthUser)    # Add to Form
       
    $Top = $Top + 30
    ## Add Authorized User
    $Script:chkAddUsr = New-Object Windows.Forms.checkbox 
        $Script:chkAddUsr.Left = 130; $Script:chkAddUsr.Width = 300; $Script:chkAddUsr.Top = $Top
        $Script:chkAddUsr.Text = "Add Authorized Users" 
        $Script:chkAddUsr.Checked = $false   # set a default value
        $Global:form.Controls.Add($Script:chkAddUsr)
        $Script:chkAddUsr.Add_Click({
            $Script:lblUsrChg.Text = "Add Auth Users:"
            $Global:okButton.Text = "Add"
            Display-AuthUsers
        })

    ## Enter Users
    $Script:lblUsrChg = New-Object System.Windows.Forms.Label   
        $Script:lblUsrChg.Top = $Top ; $Script:lblUsrChg.Left = 10; $Script:lblUsrChg.Width=150; $Script:lblUsrChg.AutoSize = $true 
        $Global:form.Controls.Add($Script:lblUsrChg)    # Add to Form 
        # 
        $Script:txtUsrChg = New-Object Windows.Forms.TextBox
        $Script:txtUsrChg.Top = $Top; $Script:txtUsrChg.Left = 130; $Script:txtUsrChg.Width = 400;
        $Script:txtUsrChg.Text = "Enter Address or Employee# separate with commas...." 
        $Global:form.Controls.Add($Script:txtUsrChg)    # Add to Form 
        $Script:txtUsrChg.Add_Click({
            $Script:txtUsrChg.Text = ""
            $Global:OKButton.visible = $True 
        })

    $Top = $Top + 30
    ## Remove Authorized User
    $Script:chkRemUsr = New-Object Windows.Forms.checkbox 
        $Script:chkRemUsr.Left = 130; $Script:chkRemUsr.Width = 300; $Script:chkRemUsr.Top = $Top
        $Script:chkRemUsr.Text = "Remove Authorized Users" 
        $Script:chkRemUsr.Checked = $false   # set a default value
        $Global:form.Controls.Add($Script:chkRemUsr)
        $Script:chkRemUsr.Add_Click({
            $Script:lblUsrChg.Text = "Remove Auth Users:"
            $Global:okButton.Text = "Remove"
            Display-AuthUsers
        })                  

    $Top = $Top + 30
    ## Not for Display Box to Not Display the email template
    $Script:chkTemplate = New-Object Windows.Forms.checkbox 
        $Script:chkTemplate.Left = 130; $Script:chkTemplate.Width = 300; $Script:chkTemplate.Top = $Top
        $Script:chkTemplate.Text = "Display Email Template for Modifying Authorized Users" 
        $Script:chkTemplate.Checked = $true   # set a default value
        $Global:form.Controls.Add($Script:chkTemplate)

    Hide-Fields
}

Function Complete-Form
{
    $Script:txtDispName.Text = ($Script:txtDispName.Text).Trim()
    $Script:txtDispName.Refresh
    $Global:form.Size = New-Object System.Drawing.Size(640,430) #(W,H)
    $Script:lblGrpAddr.Visible = $True
    $Script:txtGrpAddr.Visible = $True
    $Script:lblLegAddr.Visible = $True
    $Script:txtLegAddr.Visible = $True
    $Script:lblGrpOwnr.Visible = $True
    $Script:txtGrpOwnr.Visible = $True
    $Script:lblGrpMbr.Visible = $True
    $Script:txtGrpMbr.Visible = $True
    $Script:lblAuthUser.Visible = $True
    $Script:txtAuthUser.Visible = $True
    $Script:chkTemplate.Visible = $True
    $Script:txtGrpAddr.Text = $Script:GrpDetails.PrimarySmtpAddress
    $Script:txtLegAddr.Text = ($Script:GrpDetails.EmailAddresses) -join(", ")
    $Script:txtGrpOwnr.Text = ($Script:GrpDetails.ManagedBy) -join (", ")
    $Script:txtAuthUser.Text = ($Script:GrpDetails.AcceptMessagesOnlyFromSendersOrMembers) -join (", ")
    If ($Script:Grp -notlike "DST*")
    {
        $Script:txtGrpMbr.Text = (Get-DistributionGroupMember -ResultSize Unlimited $Script:Grp) -join (", ")
    }
    else
    {
        $Script:txtGrpMbr.Text = (Get-DynamicDistributionGroupMember -ResultSize Unlimited $Script:Grp) -join (", ")
    }
    $Script:chkAddUsr.visible = $True
    $Script:chkRemUsr.visible = $True
}

Function Display-AuthUsers
{
    $Script:chkAddUsr.Visible = $False
    $Script:chkRemUsr.Visible = $False
    $Script:lblUsrChg.Visible = $True
    $Script:txtUsrChg.Visible = $True
}

Function Hide-Fields
{
    $Script:lblGrpAddr.Visible = $False
    $Script:txtGrpAddr.Visible = $False
    $Script:lblLegAddr.Visible = $False
    $Script:txtLegAddr.Visible = $False
    $Script:lblGrpOwnr.Visible = $False
    $Script:txtGrpOwnr.Visible = $False
    $Script:lblGrpMbr.Visible = $False
    $Script:txtGrpMbr.Visible = $False
    $Script:lblAuthUser.Visible = $False
    $Script:txtAuthUser.Visible = $False
    $Script:chkTemplate.Visible = $False
    $Global:okButton.Visible = $False
    $Script:chkAddUsr.visible = $False
    $Script:chkRemUsr.visible = $False
    $Script:lblUsrChg.Visible = $False
    $Script:txtUsrChg.Visible = $False
    $Script:chkAddUsr.Checked = $False
    $Script:chkRemUsr.Checked = $False
    $Script:txtUsrChg.Text = "Enter Address or Employee# separate with commas...." 
    $Global:form.Size = New-Object System.Drawing.Size(600,200) #(W,H)
}

Check-Reconnect
$FileName		= "ModifyAuthorizedUsers"
$ReportDirectory	= "E:\Automation\ModifyAuthorizedUsers\Report\"

$Global:form = ""

Add-Type -Assembly System.Windows.Forms     ## Load the Windows Forms assembly 
## Create the main form

Build-DefaultForm
Publish-Form

If ($Global:Result -eq "OK")
{
    $Script:ReportFile = $ReportDirectory + "Report-" + $FileName + "-Group" + ($Script:Grp -replace (" ","")) + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
    WriteReportEvent
    $LineToWrite = "Group Name                 :" + "`t" + $Script:txtDispName.Text
    WriteReportEvent
    $LineToWrite = "Ticket Number              :" + "`t" + $Script:txtTaskNo.Text
    WriteReportEvent
    $LineToWrite = "Email Address              :" + "`t" + $Script:txtGrpAddr.Text
    WriteReportEvent
    $LineToWrite = "Additional Aliases         :" + "`t" + $Script:txtLegAddr.Text
    WriteReportEvent
    $LineToWrite = "Group Owners               :" + "`t" + $Script:txtGrpOwnr.Text
    WriteReportEvent
    $LineToWrite = "Authorized Users           :" + "`t" + $Script:txtAuthUser.Text + "`n"
    WriteReportEvent

    $Script:txtUsrChg.Text = ($Script:txtUsrChg.Text) -replace (" ","")
    $AuthUsr = $Script:txtUsrChg.Text -split (",")

    If ($Script:chkAddUsr.Checked -eq $True)
    {
        $LineToWrite = "Starting Authorized Group Additions"
        WriteReportEvent
        Foreach ($Usr in $AuthUsr)
        {
            If ($usr -notlike "ACL*")
            {
                $Exists = [bool]($mbx = get-mailbox $usr -ErrorAction SilentlyContinue)
            }
            else
            {
                $Exists = [bool]($mbx = get-DistributionGroup $usr -ErrorAction SilentlyContinue)
            }

            If ($Exists -eq $True)
            {
                If ($Script:Grp -notlike "DST*")
                {
                    Set-DistributionGroup $Script:GrpDetails.PrimarySMTPAddress -AcceptMessagesOnlyFromSendersOrMembers @{add=$mbx.alias}
                }
                else
                {
                    Set-DynamicDistributionGroup $Script:GrpDetails.PrimarySMTPAddress -AcceptMessagesOnlyFromSendersOrMembers @{add=$mbx.alias}
                }
                write-host "Added Auth User:  " $Usr
                $LineToWrite = "ADDUSR" + "`t`t" +  "Added Authorized User: " + $Usr + " [" + $mbx.DisplayName + "]"
            }
            else
            {
                write-host "ERROR Finding" $Usr
                $LineToWrite = "FAIL" + "`t`tAccount Not Found: " + $Usr
            }
            WriteReportEvent
        }
    }

    If ($Script:chkRemUsr.Checked -eq $True)
    {
        $LineToWrite = "Starting Authorized Group Removal"
        WriteReportEvent
        Foreach ($Usr in $AuthUsr)
        {
            If ($usr -notlike "ACL*")
            {
                $Exists = [bool]($mbx = get-mailbox $usr -ErrorAction SilentlyContinue)
            }
            else
            {
                $Exists = [bool]($mbx = get-DistributionGroup $usr -ErrorAction SilentlyContinue)
            }

            If ($Exists -eq $True)
            {
                If ($Script:Grp -notlike "DST*")
                {
                    Set-DistributionGroup $Script:GrpDetails.PrimarySMTPAddress –AcceptMessagesOnlyFromSendersOrMembers @{remove=$mbx.alias}
                }
                else
                {
                    Set-DynamicDistributionGroup $Script:GrpDetails.PrimarySMTPAddress –AcceptMessagesOnlyFromSendersOrMembers @{remove=$mbx.alias}
                }
                write-host "Removed Auth User:  " $Usr
                $LineToWrite = "REMUSR" + "`t`t" +  "Removed Authorized User: " + $Usr + " [" + $mbx.DisplayName + "]"
            }
            else
            {
                write-host "ERROR Finding" $Usr
                $LineToWrite = "FAIL" + "`t`tAccount Not Found: " + $Usr
            }
            WriteReportEvent
        }
    }

    $LineToWrite = "Authorized Group Use Changes Complete"
    WriteReportEvent

    If ($Script:chkTemplate.Checked -eq $True)
    {
        Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\DLRestrictedAccessChange.oft
    }
}