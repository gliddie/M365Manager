#Set forwards on User or Shared Mailboxes
#
#    2024/04/18 - SAG - Created Script to move accounts to another AD Container
#    2026/03/25 - SAG - Modified to use the Graph commands also to modify how users are removed from licensing groups
#
#####################################################

Function Build-MAMoveLicense
{
    $Script:form = New-Object Windows.Forms.Form 
    $Script:form.FormBorderStyle = "FixedToolWindow" 
    $Script:form.Text = "Move to New License Group"
    $Script:form.StartPosition = "CenterScreen" 
    $Script:form.Width = 500 ; $Script:form.Height = 350  # Make the form wider 

    $Script:Top = 30
    ## Project Name
    $Script:lblProject = New-Object System.Windows.Forms.Label   
        $Script:lblProject.Text = "Project Name:"  
        $Script:lblProject.Top = $Script:Top ; $Script:lblProject.Left = 10; $Script:lblProject.Width=150; $Script:lblProject.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblProject)    # Add to Form 
        #
        $Script:txtProject = New-Object Windows.Forms.TextBox
        $Script:txtProject.TabIndex = 0 # set Tab Order 
        $Script:txtProject.Top = $Script:Top ; $Script:txtProject.Left = 140; $Script:txtProject.Width = 150; $Script:txtProject.AutoSize = $true
        $Script:txtProject.Location = New-Object System.Drawing.Size(140,$Script:Top)
        $Script:txtProject.Size = New-Object system.Drawing.Size(300,40)
        $Global:InputFocus = $Global:txtProject
        $Script:form.Controls.Add($Script:txtProject)    # Add to Form 

    $Script:Top = $Script:Top + 30
    ## Select New Container
    $Script:lblLicense = New-Object System.Windows.Forms.Label   
        $Script:lblLicense.Text = "New License Group:"  
        $Script:lblLicense.Top = $Script:Top ; $Script:lblLicense.Left = 10; $Script:lblLicense.Width=150; $Script:lblLicense.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblLicense)    # Add to Form 
        # 
        $Script:txtLicense = New-Object Windows.Forms.ComboBox
        $Script:txtLicense.TabIndex = 0 # set Tab Order 
        $Script:txtLicense.Location = New-Object System.Drawing.Size(140,$Script:Top)
        $Script:txtLicense.Size = New-Object system.Drawing.Size(300,40)
#        $Details = Get-AzureADGroup -SearchString "LIC.O365"
        $Script:txtLicense.Items.Clear()
        Foreach ($Rec in $LicGrps)
        {
            [void] $Script:txtLicense.Items.Add($Rec.Name)
        }
        If ($details.count -gt 0)
        {
            $Script:txtLicense.SelectedItem = $LicGrps[0].Name
        }
        else
        {
            $Script:txtLicense.SelectedItem = $LicGrps.Name
        }
        $Script:form.Controls.Add($Script:txtLicense)    # Add to Form 

    $Script:Top = $Script:Top + 30
    ## Enter Employee #s
    $Script:lblEmpID = New-Object System.Windows.Forms.Label   
        $Script:lblEmpID.Text = "Employee Number(s):"  
        $Script:lblEmpID.Top = $Top ; $Script:lblEmpID.Left = 10; $Script:lblEmpID.Width=150; $Script:lblEmpID.AutoSize = $true 
        $Script:form.Controls.Add($Script:lblEmpID)    # Add to Form 
        # 
        $Script:txtEmpID = New-Object Windows.Forms.TextBox
        $Script:txtEmpID.MaxLength = 2000000
        $Script:txtEmpID.TabIndex = 0 # set Tab Order 
        $Script:txtEmpID.Location = New-Object System.Drawing.Size(140,$Script:Top)
        $Script:txtEmpID.Size = New-Object system.Drawing.Size(300,150)
        $Script:txtEmpID.MultiLine = $true
        $Script:txtEmpID.ScrollBars = 'Both'  
        $Script:form.Controls.Add($Script:txtEmpID)    # Add to Form 
}

Function Start-Report
{
    $LineToWrite = $RecordEvent + "STAR" + "`t" + "Modify Licensing Group script has started"
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + (whoami)
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "User TimeZone: " + (Get-TimeZone)
    WriteReportEvent
    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Number of Users to Process: " + $addMember.count
    WriteReportEvent
    $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
    WriteReportEvent
}

$me = whoami
$CredENo = $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))
$UsrName = $CredENo + "@global.ul.com"
$dir = "c:\users\" + $CredENo + "\documents\"
$File = "my" + $CredENo + "File.xml"
$Global:CredFile = $dir + $File
#AdminCredFile
$AFile = "myA" + ($me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))).replace(".","") + "File.xml"
$Global:ACredFile = $dir + $AFile
If (Test-Path $Global:ACredFile)
{
    $Global:AdmLiveCred = Import-Clixml $Global:ACredFile
}

$wshell = New-Object -ComObject Wscript.Shell
write-host "Gathering All Licensing Groups"
$LicGrps = import-csv E:\O365AdminShared\Data\LicenseGroups.csv
Build-MAMoveLicense
Add-FormStandardButtons
$Script:OKButton.Text = "Continue"
Publish-Form

Do
{
    If ($Script:Result -eq "OK")
    {
        If ($Script:txtEmpID.Text.Length -gt 0)
        {
            $ReportFile = "E:\Automation\MAActivities\Reports\ModifyLicenses\ModifyLicenses-" + $Script:txtProject.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            write-host $ReportFile
            Start-Report
            $addMember = (($Script:txtEmpID.text -replace ("`n",",")).split(",") -replace(" ","")).Trim()
            write-Host "Number of Accounts to Process: " $addMember.Count
            $NewLic = (Get-MgGroup -Filter "DisplayName eq '$Script:txtLicense.SelectedItem'")

            foreach ($u in $addMember)
            {
                write-host "`nProcessing UserID: " $u -ForegroundColor Cyan
                $LineToWrite = $RecordEvent + "CURR" + "`t" + "                  Employee ID: " + $u
                WriteReportEvent
                $ErrorActionPreference = "SilentlyContinue"
                $ADExists = $False
                $ADExists = [bool](get-aduser $u)
                $ErrorActionPreference = "Continue"

                If ($ADExists -ne $False)
                {
                    If ($u -notlike "*@*")
                    {
                       $UPN = $u + "@global.ul.com"
                       $MgEmpNo = $u + "@"
                       $O365Lic = Get-MgUserLicenseDetail -UserId ($u + "@global.ul.com")
                    }
                    else
                    {
                        $UPN = $u
                        $MgEmpNo = $u + "@"
                        $O365Lic = Get-MgUserLicenseDetail -UserId $u
                    }
                    $MgUsr = Get-MgUser -Filter "startsWith(UserPrincipalName, '$MgEmpNo')"
                    $DLMember = Get-MgUserMemberOf -UserId ($u + "@global.ul.com") -All
                    write-host "`tNumber of licenses assigned to this individual: " $O365Lic.Count
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "     Number licenses assigned: " + $O365Lic.Count
                    WriteReportEvent
               
               #Add user to the new license group
               
                    New-MgGroupMember -GroupId 4d2c6448-840b-41d7-b05d-60ac618062a6 -DirectoryObjectId $MgUsr.Id
#                    New-MgGroupMember -GroupId $NewLic.Id -DirectoryObjectId $MgUsr.Id
                    Write-Host "`tAdding to the membership of: " $Script:txtLicense.SelectedItem
                    $LineToWrite = $RecordEvent + "ADD " + "`t" + "  Adding to New License Group: " + $Script:txtLicense.SelectedItem
                    WriteReportEvent
                    
                    Foreach ($l in $O365Lic)
                    {
                        If (($l.SkuPartNumber -ne "M365_E5_SUITE_COMPONENTS") -and ($l.SkuPartNumber -ne "ENTERPRISEPREMIUM") -and ($l.SkuPartNumber -ne "EMSPREMIUM") -and ($l.SkuPartNumber -ne "Microsoft_365_Copilot"))
                        {
<#                            If ($l.SkuPartNumber -ne "Microsoft_365_Copilot")
                            {
                                $CoPwTGrp = [bool](Get-MgUserMemberOf -UserId $UPN -All | Where-Object {($_.Id -eq "11c2475a-7daa-4962-b489-9d4bfccc7597")})
                                $CoPwoTGrp = [bool](Get-MgUserMemberOf -UserId $UPN -All | Where-Object {($_.Id -eq "6047997a-6204-48e1-b5c3-96f5f39637ec")})
                                If ($CoPwTGrp -eq $True)
                                {
                                    Remove-MgGroupMemberByRef -GroupId 11c2475a-7daa-4962-b489-9d4bfccc7597 -DirectoryObjectId $MgUsr.Id
                                    write-host "`tRemoving from CoPilot w/o Transcription Group License: " $l.SkuPartNumber
                                    $LineToWrite = $RecordEvent + "REMV" + "`t" + "  Removing CoPilot w/o Transcription Group License: " + $l.SkuPartNumber
                                    WriteReportEvent
                                }
                                If ($CoPwoTGrp -eq $True)
                                {
                                    Remove-MgGroupMemberByRef -GroupId 6047997a-6204-48e1-b5c3-96f5f39637ec -DirectoryObjectId $MgUsr.Id
                                    write-host "`tRemoving from CoPilot w/Transcription Group License: " $l.SkuPartNumber
                                    $LineToWrite = $RecordEvent + "REMV" + "`t" + "  Removing CoPilot w/Transcription Group License: " + $l.SkuPartNumber
                                    WriteReportEvent
                                }
                            }
                            else
#>
#                            {
                                write-host "`tRemoving Additional License: " $l.SkuPartNumber
                                $LineToWrite = $RecordEvent + "REMV" + "`t" + "  Removing Additional License: " + $l.SkuPartNumber
                                WriteReportEvent
                                $LicSkuId = $l.SkuId
                                Set-MgUserLicense -UserId $MgUsr.Id -AddLicenses @{} -RemoveLicenses @($l.SkuId)
                            }
#                        }
                    }
<#
               #Remove Original E5 License Group Membership Last so that the Membership into the new License Group is completed before removing original groups.
                    $EmpGrp = [bool](Get-MgUserMemberOf -UserId $UPN -All | Where-Object {($_.Id -eq "2aab81de-e54a-4c58-909c-b15273c54dff")})
                    $NonEmpGrp = [bool](Get-MgUserMemberOf -UserId $UPN -All | Where-Object {($_.Id -eq "40dff561-964a-4f7b-9794-dba3542b63c4")})
                    If ($EmpGrp -eq $True)
                    {
                        Remove-MgGroupMemberByRef -GroupId 2aab81de-e54a-4c58-909c-b15273c54dff -DirectoryObjectId $MgUsr.Id
                        write-host "`tRemoving from Employee Group License: LIC.O365.E5Employee"
                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "  Removing Employee Group License: LIC.O365.E5Employee"
                        WriteReportEvent
                    }

                    If ($NonEmpGrp -eq $True)
                    {
                        Remove-MgGroupMemberByRef -GroupId 40dff561-964a-4f7b-9794-dba3542b63c4 -DirectoryObjectId $MgUsr.Id
                        write-host "`tRemoving from Employee Group License: LIC.O365.E5NonEmployee"
                        $LineToWrite = $RecordEvent + "REMV" + "`t" + "  Removing Employee Group License: LIC.O365.E5NonEmployee"
                        WriteReportEvent
                    }
#>
                    Foreach ($m in $LicGrps)
                    {
                        If ($m.Name -ne $Script:txtLicense.SelectedItem)
                        {
                            $GrpMem = ""
                            $GrpMem = [bool](Get-MgUserMemberOf -UserId $UPN -All | Where-Object {($_.Id -eq $m.Id)})
                            If ($GrpMem -eq $True)
                            {
                                Remove-MgGroupMemberByRef -GroupId $m.Id -DirectoryObjectId $MgUsr.Id
    #                            Remove-AzureADGroupMember -ObjectId $m.ObjectID -MemberId $MgUsr.Id
                                write-host "`tRemoving from Standard License group: " $m.Name
                                $LineToWrite = $RecordEvent + "REMV" + "`t" + "  Removing from License Group: " + $m.Name
                                WriteReportEvent
                            }
                        }
                    }
                }
                else
                {
                    write-host "`tNo Account found for: " $u -ForegroundColor Red
                    $LineToWrite = $RecordEvent + "NOACT" + "`t" + "         No Account found for: " + $u
                    WriteReportEvent
                }
                $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
                WriteReportEvent
            }
        }
        else
        {
            $output = $wshell.Popup("No employee number or email addresses entered try again.",0,"No Employee Nos.",0+32)
        }
        $Script:txtEmpID.Text = ""
        Publish-Form
    }
}While ($Script:Result -eq "OK")
