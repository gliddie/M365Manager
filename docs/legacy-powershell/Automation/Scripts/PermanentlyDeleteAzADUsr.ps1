<##
#  Created to Permanently Delete AzureAD User
#
#  10/25/2024 - SAG - Added code in order to create a report file for the account removal
#  05/20/2025 - SAG - Changed over code using the MSOLService commands to use Graph
#
#>

Function Build-Form
{

    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Permanently Remove AzureAD User" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(560,230) #(W,H)

    $Top = 20
    $TxtLeft = 120
    ## Enter User
    $Script:lblRemUsr = New-Object System.Windows.Forms.Label   
    $Script:lblRemUsr.Text = "Account:"
    $Script:lblRemUsr.Top = $Top ; $Script:lblRemUsr.Left = 10; $Script:lblRemUsr.Width=120 ;$Script:lblRemUsr.AutoSize = $true 
    $Global:form.Controls.Add($Script:lblRemUsr)    # Add to Form 
    $Script:txtRemUsr = New-Object Windows.Forms.TextBox  
    $Script:txtRemUsr.TabIndex = 0 # set Tab Order 
    $Script:txtRemUsr.Top = $Top; $Script:txtRemUsr.Left = $TxtLeft; $Script:txtRemUsr.Width = 250;  
    $Script:txtRemUsr.Text = ""
    $Global:form.Controls.Add($Script:txtRemUsr)    # Add to Form
    $Global:InputFocus = $Script:txtRemUsr
    $Script:txtRemUsr.Add_Click({
        $Script:txtRemUsr.Text = ""
        $Script:txtUPN.Text = ""
        $Script:lblUPN.Visible = $False
        $Script:txtUPN.Visible = $False
        $Script:txtDispName.Text = ""
        $Script:lblDispName.visible = $False
        $Script:txtDispName.visible = $False
        $Script:ButGetENo.Visible = $True
        $Script:txtUPN.Visible = $False
        $Script:txtDispName.Visible = $False
        $Script:lblTicketNo.Visible = $False
        $Script:txtTicketNo.Text = ""
        $Script:txtTicketNo.Visible = $False
        $Global:OKButton.Visible = $False
    })

    $Script:ButGetENo = New-Object Windows.Forms.Button
        $Script:ButGetENo.Location = New-object System.Drawing.Size(380,$Top)
        $Script:ButGetENo.Size = new-Object System.Drawing.Size(150,20)
        $Script:ButGetENo.Text = "Get Account Details"
        $Global:form.Controls.Add($Script:ButGetENo)
        $Script:ButGetENo.Add_Click({
            $Exists = [bool]($usr = Get-AzureADUser -SearchString $Script:txtRemUsr.Text)
            If ($Exists -eq $True)
            {
                $Script:txtUPN.Text = $usr.UserPrincipalName
                $Script:txtDispName.Text = $usr.DisplayName
                $Script:lblUPN.Visible = $True
                $Script:txtUPN.Visible = $True
                $Script:lblDispName.visible = $True
                $Script:txtDispName.visible = $True
                $Script:ButGetENo.Visible = $False
                $Script:lblTicketNo.Visible = $True
                $Script:txtTicketNo.Visible = $True
            }
            else
            {
                $Act = $Script:txtRemUsr.Text
                Add-Type -AssemblyName PresentationCore,PresentationFramework
                $ButtonType = [System.Windows.MessageBoxButton]::OK
                $MessageIcon = [System.Windows.MessageBoxImage]::Error
                $MessageBody = "Account Not Found for: $Act"
                $MessageTitle = "Not Found"
                $Script:Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
                $Script:txtRemUsr.Text = ""
                $Global:InputFocus = $Script:txtRemUsr
            }
        })

    $Top = $Top + 30
    ##Ticket Number
    $Script:lblTicketNo = New-Object System.Windows.Forms.Label
    $Script:lblTicketNo.Text = "Ticket Number:"
    $Script:lblTicketNo.Top = $Top ; $Script:lblTicketNo.Left = 10; $Script:lblTicketNo.Width=120 ;$Script:lblTicketNo.AutoSize = $true 
    $form.Controls.Add($Script:lblTicketNo)    # Add to Form 
    # 
    $Script:txtTicketNo = New-Object Windows.Forms.TextBox
    $Script:txtTicketNo.TabIndex = $Tab++ # set Tab Order 
    $Script:txtTicketNo.Top = $Top; $Script:txtTicketNo.Left = $TxtLeft; $Script:txtTicketNo.Width = 250;
    $Global:form.Controls.Add($Script:txtTicketNo)    # Add to Form
    $Script:lblTicketNo.Visible = $False
    $Script:txtTicketNo.Visible = $False
    $Script:txtTicketNo.Add_Click({
        $Global:OKButton.Visible = $True
    })

    $Top = $Top + 30    
    ## Account UPN
    $Script:lblUPN = New-Object System.Windows.Forms.Label   
    $Script:lblUPN.Text = "UserPrincipalName:"
    $Script:lblUPN.Top = $Top ; $Script:lblUPN.Left = 10; $Script:lblUPN.Width=120 ;$Script:lblUPN.AutoSize = $true
    $Script:lblUPN.Visible = $False
    $Global:form.Controls.Add($Script:lblUPN)    # Add to Form 
    $Script:txtUPN = New-Object Windows.Forms.TextBox  
    $Script:txtUPN.TabIndex = 0 # set Tab Order 
    $Script:txtUPN.Top = $Top; $Script:txtUPN.Left = $TxtLeft; $Script:txtUPN.Width = 250;  
    $Script:txtUPN.Text = ""
    $Script:txtUPN.Visible = $False
    $Script:txtUPN.ReadOnly = $True
    $Global:form.Controls.Add($Script:txtUPN)    # Add to Form

    $Top = $Top + 30      
    ## Account DisplayName
    $Script:lblDispName = New-Object System.Windows.Forms.Label   
    $Script:lblDispName.Text = "DisplayName:"
    $Script:lblDispName.Top = $Top ; $Script:lblDispName.Left = 10; $Script:lblDispName.Width=120 ;$Script:lblDispName.AutoSize = $true
    $Script:lblDispName.visible = $False
    $Global:form.Controls.Add($Script:lblDispName)    # Add to Form 
    $Script:txtDispName = New-Object Windows.Forms.TextBox  
    $Script:txtDispName.TabIndex = 0 # set Tab Order 
    $Script:txtDispName.Top = $Top; $Script:txtDispName.Left = $TxtLeft; $Script:txtDispName.Width = 250;  
    $Script:txtDispName.Text = ""
    $Script:txtDispName.Visible = $False
    $Script:txtDispName.ReadOnly = $True
    $Global:form.Controls.Add($Script:txtDispName)    # Add to Form
}

Build-Form
Add-FormStandardButtons
$Global:OKButton.Visible = $False
Publish-Form

If ($Global:Result -eq "OK")
{
    $Filename = "RemoveAzureAccount"
    $Year = (get-date).ToString("yyyy")
    $Path = "e:\Automation\RemoveAzureAccount\Report\" + $Year + "\"
    If (Test-Path $path) {} else {New-Item -Path $Path -ItemType Directory}
    $ReportFile = $Path + "Report-RemoveAzureAccount" + "-Acct" + $Script:txtRemUsr.Text  + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"

    $LineToWrite = $RecordEvent + "STAR" + "`tPermanently Remove Azure User script has started"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`tLaunched by: " + $WhoAmI + "`n"
    WriteReportEvent

    write-host "Removing account: " $Script:txtRemUsr.Text
    $usr = Get-AzureADUser -SearchString $Script:txtRemUsr.Text
    $inf= $usr.UserPrincipalName + ", " + $usr.displayName

    $LineToWrite = $RecordEvent + "INFO" + "`t            Account Details Entered: " + $Script:txtRemUsr.Text
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t                Account DisplayName: " + $usr.displayName
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t                   Account ObjectID: " + $usr.ObjectId
    WriteReportEvent

    Remove-AzureADUser -ObjectId $usr.ObjectId
    $Searchval = $usr.UserPrincipalName

    write-host "Removing the SoftDeleted Account" $SearchVal
    $LineToWrite = $RecordEvent + "INFO" + "`t       Removing Softdeleted Account: " + $SearchVal
    WriteReportEvent
    $Cnt = 0
    $DispName = $usr.displayName
    Do {
        Start-Sleep -Seconds 2
        $Found = [bool]($val = get-MgDirectoryDeletedItemAsUser -All |Where-Object {$_.DisplayName -like $DispName})
#        $Found = [bool]($val = get-MsolUser -ReturnDeletedUsers -all |Where-Object {$_.UserPrincipalName -like $Searchval})
        $cnt++
    }While (($Found -ne $True) -and ($cnt -le 10))

    $Exists = ""
    If ($Found -eq $True)
    {
        Remove-AzureADMSDeletedDirectoryObject -Id $val.ID
        start-sleep -Seconds 2
        $Exists = [bool]($val = get-MgDirectoryDeletedItemAsUser -All |Where-Object {$_.DisplayName -like $DispName})
        If ($Exists -eq $True)
        {
            $LineToWrite = $RecordEvent + "INFO" + "`tUnable to Permanently Remove AzureAD Account: " + $Usr.UserPrincipalName + "," + $Usr.DisplayName
            $FinalStatus = "`tUnable to Permanently Remove AzureAD Account"
            $Status = "Unable to Permanently Remove AzureAD Account: " + $Usr.UserPrincipalName + "," + $Usr.DisplayName
        }
        else
        {
            $LineToWrite = $RecordEvent + "INFO" + "`tPermanently Removed AzureAD Account: " + $Usr.UserPrincipalName
            $FinalStatus = "`tAzureAD Account Removal complete"
            $Status = "Permanently Removed User from AzureAD: " + $Usr.UserPrincipalName
        }
    }
    else
    {
        $LineToWrite = $RecordEvent + "INFO" + "`tUnable to Find AzureAD User Account: " + $Usr.UserPrincipalName
        $FinalStatus = "`tUnable to Find AzureAD User Account"
        $Status = "Unable to Find AzureAD User Account: " + $Usr.UserPrincipalName
    }
    WriteReportEvent
    write-host $Status
    $LineToWrite = $RecordEvent + "STAR" + $FinalStatus
    WriteReportEvent
}
