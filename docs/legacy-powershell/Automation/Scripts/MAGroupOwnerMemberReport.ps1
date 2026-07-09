#Set forwards on User or Shared Mailboxes
#
#    2025/16/18 - SAG - Report of Groups an individual is a member or owner of
#
#####################################################

Function Build-MAGrpMbrOwn
{
    $Script:form = New-Object Windows.Forms.Form 
    $Script:form.FormBorderStyle = "FixedToolWindow" 
    $Script:form.Text = "Report Group Membership/Ownership"
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
    $LineToWrite = $RecordEvent + "STAR" + "`t" + "Generating Report of Group Ownership/Membership has started"
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
Build-MAGrpMbrOwn
Add-FormStandardButtons
$Script:OKButton.Text = "Continue"
Publish-Form

Do
{
    If ($Script:Result -eq "OK")
    {
        If ($Script:txtEmpID.Text.Length -gt 0)
        {
            $ReportFile = "E:\Automation\MAActivities\Reports\MoveADOU\MoveADOU-" + $Script:txtProject.Text + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
            write-host $ReportFile
            $addMember = (($Script:txtEmpID.text -replace ("`n",",")).split(",") -replace(" ","")).Trim()
            Start-Report
            write-Host "Number of Accounts to Process: " $addMember.Count
            foreach ($u in $addMember)
            {
                If ($u.length -gt 0)
                {
                    write-host "Processing UserID: " $u
                    $LineToWrite = $RecordEvent + "CURR" + "`t" + "                  Employee ID: " + $u
                    WriteReportEvent
                    $ErrorActionPreference = "SilentlyContinue"
                    $ADExists = ""
                    $ADExists = [bool]($strDN = (Get-ADuser $u -Properties DistinguishedName).DistinguishedName)
                    $ErrorActionPreference = "Continue"
                    If ($ADExists -eq "True")
                    {
                        If ($strdn -notlike ("*"+$Script:txtContainer.Text))
                        {
                            $Protected = Get-ADObject -Identity $strDN -Properties ProtectedFromAccidentalDeletion
                            If ($Protected.ProtectedFromAccidentalDeletion -eq $True)
                            {
                                write-host "`tDisablng Object Protection"
                                set-ADObject -identity $strDN -ProtectedFromAccidentalDeletion $false
                                $LineToWrite = $RecordEvent + "CURR" + "`t" + "Disabling ADObject Protection: " + $Protected.ProtectedFromAccidentalDeletion
                                WriteReportEvent
                                $LineToWrite = $RecordEvent + "CURR" + "`t" + "   Current ADObject Container: " + $strDN
                                WriteReportEvent
                                write-host "`tMoving User Object"
                                Move-ADObject $strDN $Script:txtContainer.Text
                                $LineToWrite = $RecordEvent + "CURR" + "`t" + "       New ADObject Container: " + $Script:txtContainer.Text
                                WriteReportEvent
                                Start-Sleep -Seconds 1
                                $strDN = (Get-ADuser $u -Properties DistinguishedName).DistinguishedName
                                write-host "Resetting Object Protection: " $strDN
                                set-ADObject -identity $strDN -ProtectedFromAccidentalDeletion $True -Credential $Global:AdmLiveCred
                                $LineToWrite = $RecordEvent + "CURR" + "`t" + " Enabling ADObject Protection: " + (Get-ADObject -Identity $strDN -Properties ProtectedFromAccidentalDeletion)
                                WriteReportEvent
                            }
                            else
                            {
                                $LineToWrite = $RecordEvent + "CURR" + "`t" + "   Current ADObject Container: " + $strDN
                                WriteReportEvent
                                write-host "`tMoving User Object"
                                Move-ADObject $strDN $Script:txtContainer.Text
                                $LineToWrite = $RecordEvent + "CURR" + "`t" + "       New ADObject Container: " + $Script:txtContainer.Text
                                WriteReportEvent
                            }
                            $LineToWrite = $RecordEvent + "----" + "`t" + "---------------------------------------------------------------"
                            WriteReportEvent
                        }
                        else
                        {
                            write-host "Account is already in the new container"
                            $LineToWrite = $RecordEvent + "CURR" + "`t" + "      Account is already in the new container for: " + $u
                            WriteReportEvent
                        }
                    }
                    else
                    {
                        write-host "AD Account does not exist"
                        $LineToWrite = $RecordEvent + "CURR" + "`t" + "      No AD Account Found for: " + $u
                        WriteReportEvent
                    }
                }
            }
            $LineToWrite = $RecordEvent + "DONE" + "`t" + "Move AD Accounts to New Container Complete"
            WriteReportEvent
        }
        else
        {
            $output = $wshell.Popup("No employee number or email addresses entered try again.",0,"No Employee Nos.",0+32)
        }
        $Script:txtEmpID.Text = ""
        Publish-Form
    }
}While ($Script:Result -eq "OK")
