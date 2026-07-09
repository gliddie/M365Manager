<#
#
#  Called by:  DistributionSecurityGroupMenu.ps1
#
#  08/08/2020 - Created for new GUI Interface
#  11/16/2020 - Added code to set owner/ticket info on RRS groups and also to not perform the "Get-group/Set-Group" information for RoomLists.
#>

function Build-DLInputForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Rename Distribution Group" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(450,250) #(W,H)
    $BldDetails = "N"

    ## Get Details used to create
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "DL Name:"
        $Global:lblDispName.Top = 10 ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtInpName = New-Object Windows.Forms.TextBox  
        $Global:txtInpName.TabIndex = 0 # set Tab Order 
        $Global:txtInpName.Top = 10; $Global:txtInpName.Left = 130; $Global:txtInpName.Width = 280;  
        $Global:txtInpName.Text = ""   # DisplayName
        $Global:txtInpName.TabIndex = 0
        $Global:form.Controls.Add($Global:txtInpName)    # Add to Form

    ## New Display Name
    $Global:lblNewDispName = New-Object System.Windows.Forms.Label   
        $Global:lblNewDispName.Text = "New DL Name:"
        $Global:lblNewDispName.Top = 40 ; $Global:lblNewDispName.Left = 10; $Global:lblNewDispName.Width=120 ;$Global:lblNewDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblNewDispName)    # Add to Form 
        # 
        $Global:txtInpNewName = New-Object Windows.Forms.TextBox  
        $Global:txtInpNewName.TabIndex = 0 # set Tab Order 
        $Global:txtInpNewName.Top = 40; $Global:txtInpNewName.Left = 130; $Global:txtInpNewName.Width = 280;  
        $Global:txtInpNewName.Text = ""   # DisplayName
        $Global:txtInpNewName.TabIndex = 0
        $Global:form.Controls.Add($Global:txtInpNewName)    # Add to Form 
 
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 70 ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtInpTaskNo.Top = 70; $Global:txtInpTaskNo.Left = 130; $Global:txtInpTaskNo.Width = 120;  
        $Global:txtInpTaskNo.Text = "TASK"   # Enter ticket number
        $Global:txtInpTaskNo.TabIndex = 1
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form

        Add-FormStandardButtons
}

function Build-DLDetailsForm
{
    $BldDetails = "N"
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Rename Distribution Group Details" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(750,450) #(W,H)

    ## Get Details used to create
    ## Display Name
    $LocTop = 10
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "DL Name:"
        $Global:lblDispName.Top = $LocTop ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtDLDet = New-Object Windows.Forms.TextBox  
        $Global:txtDLDet.TabIndex = 0 # set Tab Order
        $Global:txtDLDet.Top = $LocTop; $Global:txtDLDet.Left = 130; $Global:txtDLDet.Width = 400;  
        $Global:txtDLDet.Text = $Global:txtInpName.Text
        $Global:txtDLDet.TabIndex = 0
        $Global:form.Controls.Add($Global:txtDLDet)    # Add to Form 
 
    $LocTop = $LocTop + 30
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = $LocTop ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtDetTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtDetTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtDetTaskNo.Top = $LocTop; $Global:txtDetTaskNo.Left = 130; $Global:txtDetTaskNo.Width = 400;
        $Global:txtDetTaskNo.Text = $Global:txtInpTaskNo.Text 
        $Global:txtDetTaskNo.TabIndex = 1
        $Global:form.Controls.Add($Global:txtDetTaskNo)    # Add to Form

    $LocTop = $LocTop + 30
    ## DL Notes
    $Global:lblDLNotes = New-Object System.Windows.Forms.Label   
        $Global:lblDLNotes.Text = "Notes:"  
        $Global:lblDLNotes.Top = $LocTop ; $Global:lblDLNotes.Left = 10; $Global:lblDLNotes.Width=120 ; $Global:lblDLNotes.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLNotes)    # Add to Form 
        # 
        $Global:txtInpDLNotes = New-Object Windows.Forms.TextBox  
        $Global:txtInpDLNotes.TabIndex = 0 # set Tab Order 
        $Global:txtInpDLNotes.Top = $LocTop; $Global:txtInpDLNotes.Left = 130; $Global:txtInpDLNotes.Width = 400;  
        $Global:txtInpDLNotes.TabIndex = 3
        $Global:txtInpDLNotes.Text = "(Not Configurable on This Group Type)"
        If ($DSTGrpDetails.RecipientTypeDetails -ne "RoomList")
        {
            $Global:txtInpDLNotes.Text = (get-group $Global:txtInpName.Text).Notes
        }
        $Global:form.Controls.Add($Global:txtInpDLNotes)    # Add to Form

    $LocTop = $LocTop + 30
    ## DL Owners
    $Global:lblDLOwner = New-Object System.Windows.Forms.Label   
        $Global:lblDLOwner.Text = "Group Owner(s):"  
        $Global:lblDLOwner.Top = $LocTop ; $Global:lblDLOwner.Left = 10; $Global:lblDLOwner.Width=120 ; $Global:lblDLOwner.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLOwner)    # Add to Form 
        # 
        $Global:txtInpDLOwner = New-Object Windows.Forms.TextBox  
        $Global:txtInpDLOwner.TabIndex = 0 # set Tab Order 
        $Global:txtInpDLOwner.Top = $LocTop; $Global:txtInpDLOwner.Left = 130; $Global:txtInpDLOwner.Width = 400;  
        $Global:txtInpDLOwner.TabIndex = 2
        $Global:txtInpDLOwner.Text = ((Get-DistributionGroup $Global:txtDLDet.Text).ManagedBy -join ", ")
        $Global:form.Controls.Add($Global:txtInpDLOwner)    # Add to Form

    $LocTop = $LocTop + 30
    ## DL Members Count
    $Global:lblDLMemCnt = New-Object System.Windows.Forms.Label   
        $Global:lblDLMemCnt.Text = "Number of Member(s):"  
        $Global:lblDLMemCnt.Top = $LocTop ; $Global:lblDLMemCnt.Left = 10; $Global:lblDLMemCnt.Width=120 ; $Global:lblDLMemCnt.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLMemCnt)    # Add to Form 
        # 
        $Global:txtInpDLMemCnt = New-Object Windows.Forms.TextBox  
        $Global:txtInpDLMemCnt.TabIndex = 0 # set Tab Order 
        $Global:txtInpDLMemCnt.Top = $LocTop; $Global:txtInpDLMemCnt.Left = 130; $Global:txtInpDLMemCnt.Width = 400;  
        $Global:txtInpDLMemCnt.TabIndex = 3
        $Global:txtInpDLMemCnt.Text = (Get-DistributionGroupMember $Global:txtInpName.Text).count
        $Global:form.Controls.Add($Global:txtInpDLMemCnt)    # Add to Form

    $LocTop = $LocTop + 30
    ## DL Members
    $Global:lblDLMem = New-Object System.Windows.Forms.Label   
        $Global:lblDLMem.Text = "Group Member(s):"  
        $Global:lblDLMem.Top = $LocTop ; $Global:lblDLMem.Left = 10; $Global:lblDLMem.Width=120 ; $Global:lblDLMem.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLMem)    # Add to Form 
        # 
        $Global:txtInpDLMem = New-Object Windows.Forms.TextBox  
        $Global:txtInpDLMem.TabIndex = 0 # set Tab Order 
        $Global:txtInpDLMem.Location = New-Object System.Drawing.Size(130,$LocTop)
        $Global:txtInpDLMem.Size = New-Object system.Drawing.Size(400,120)
        $Global:txtInpDLMem.MultiLine = $true
        $Global:txtInpDLMem.ScrollBars = 'Both'
        $Global:txtInpDLMem.TabIndex = 4
        $Global:txtInpDLMem.Text = (Get-DistributionGroupMember $Global:txtInpName.Text).PrimarySMTPAddress -join "; "
        $Global:form.Controls.Add($Global:txtInpDLMem)    # Add to Form

    ## New Display Name
    $LocTop = $LocTop + 140
    $Global:lblNewDispName = New-Object System.Windows.Forms.Label   
        $Global:lblNewDispName.Text = "New DL Name:"
        $Global:lblNewDispName.Top = $LocTop ; $Global:lblNewDispName.Left = 10; $Global:lblNewDispName.Width=120 ;$Global:lblNewDispName.AutoSize = $true
        $Global:form.Controls.Add($Global:lblNewDispName)    # Add to Form 
        # 
        $Global:txtNewDLDet = New-Object Windows.Forms.TextBox  
        $Global:txtNewDLDet.TabIndex = 0 # set Tab Order
        $Global:txtNewDLDet.Top = $LocTop; $Global:txtNewDLDet.Left = 130; $Global:txtNewDLDet.Width = 400;  
        $Global:txtNewDLDet.Text = $Global:txtInpNewName.Text
        $Global:txtNewDLDet.TabIndex = 0
        $Global:form.Controls.Add($Global:txtNewDLDet)    # Add to Form 

    ## New Address
    $LocTop = $LocTop + 30
    $Global:lblDLAddr = New-Object System.Windows.Forms.Label   
        $Global:lblDLAddr.Text = "Email Address:"
        $Global:lblDLAddr.Top = $LocTop ; $Global:lblDLAddr.Left = 10; $Global:lblDLAddr.Width=120 ;$Global:lblDLAddr.AutoSize = $true
        $Global:form.Controls.Add($Global:lblDLAddr)    # Add to Form 
        # 
        $Global:txtDLAddr = New-Object Windows.Forms.TextBox  
        $Global:txtDLAddr.TabIndex = 0 # set Tab Order
        $Global:txtDLAddr.Top = $LocTop; $Global:txtDLAddr.Left = 130; $Global:txtDLAddr.Width = 400;
#        $Global:DLAddr = ((Get-Culture).textinfo.totitlecase($Global:txtInpNewName.Text))
        $Global:txtDLAddr.Text = ($Global:txtInpNewName.Text -Replace '[ &/,$#_-]','') + "@ul.com"
        $Global:txtDLAddr.TabIndex = 0
        $Global:form.Controls.Add($Global:txtDLAddr)    # Add to Form 

    ## Authorized Requestor
    $LocTop = $LocTop + 30
    $Global:chkAuthorized = New-Object Windows.Forms.checkbox 
        $Global:chkAuthorized.Left = 130; $Global:chkAuthorized.Width = 200; $Global:chkAuthorized.Top = $LocTop  
        $Global:chkAuthorized.Text = "Authorized Requestor" 
        $Global:chkAuthorized.Checked = $false   # set a default value 
        $Global:chkAuthorized.TabIndex = 5
        $Global:form.Controls.Add($Global:chkAuthorized) 
        # Obtain Value with: $Global:chkAuthorized.Checked

        Add-FormStandardButtons
}

$Global:GrpName = $Global:txtDispName.Text
$Global:EMailAddr = $Global:txtGrpAddr.Text

Build-DLInputForm
Publish-Form

$DstExists = [bool](Get-DistributionGroup $Global:txtInpName.Text -ErrorAction SilentlyContinue)
if ($DstExists -eq "True")
{

    $DSTGrpDetails = Get-DistributionGroup $Global:txtInpName.Text
#    $DSTGrpDetails = Get-DistributionGroup $Global:txtDLDet.Text
    Build-DLDetailsForm
    Publish-Form

    If ($Global:Result -eq "OK")
    {
        If ($Global:chkAuthorized.Checked -eq "Checked")
        {
            If ($DSTGrpDetails.RecipientTypeDetails -ne "RoomList")
            {
                $GrpInf = get-Group $Global:txtDLDet.text
            }

            If ($GrpInf.Notes -like "*Per*")
            {
                $GrpInf.Notes = $GrpInf.Notes.Remove($GrpInf.Notes.IndexOf("Per")) + "Per " + $Global:txtDetTaskNo.Text
            }
            else
            {
                If ($Global:txtDLDet.Text -notlike "*outOfPolicy*")
                {
                    $LSTTag = "Owners: "
                }
                else
                {
                    $LSTTag = "Room Delegates: "
                }
                 
                $GrpInf.Notes = $LSTTag + ((Get-DistributionGroup $Global:txtDLDet.Text).ManagedBy -join (", ")) + " - Per " + $Global:txtDetTaskNo.Text
                $MailTip = $LSTTag + ((Get-DistributionGroup $Global:txtDLDet.Text).ManagedBy -join (", "))
            }

            If ($DSTGrpDetails.RecipientTypeDetails -ne "RoomList")
            {
                Set-Group $Global:txtDLDet.Text -Notes $GrpInf.Notes
            }
            write-host "Resetting distribution group values"
            Set-DistributionGroup $Global:txtDLDet.Text -DisplayName $Global:txtNewDLDet.Text -Name $Global:txtNewDLDet.Text -Alias ($Global:txtNewDLDet.Text -Replace '[ &/,$#_-]','') -EmailAddresses @{add=$Global:txtDLAddr.Text}
            write-host "Resetting distribution group primary email address and mailtip value"
            Set-DistributionGroup $Global:txtNewDLDet.Text -PrimarySmtpAddress $Global:txtDLAddr.Text -MailTip $MailTip
        }
        else
        {
            $Output = $wshell.Popup("Obtain approval from authorized requestor - No Changes Made",0,"Not Authorized",0+32)
        }
        Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\DLRenamed.oft
    }
    else
    {
        $Output = $wshell.Popup("Distribution Group rename cancelled.",0,"Request Cancelled",0+32)
    }
}
else
{
    $Output = $wshell.Popup("Distribution List Does Not Exist - No Changes Made",0,"Not Found",0+32)
}
