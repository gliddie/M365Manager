<####  This script create new UL Employee Only and UL Staff Dynamic Distribution Lists
#
# Called by:  DistributionSecurityGroupMenu.ps1
#
# 06/17/2016 - SAG - Added code to check for Displaynames longer than 64 chara and truncate them
#                  - Removed the MailUsers configuraiton as these individuals are mail enabled but not licensed
#				   - Added code to allow for creation of DST.SUP groups
# 09/28/2016 - SAG - Removed extra quotes around the executive name in the DST.SUP group code
# 08/31/2017 - SAG - Added code to allow the configuring the Supervisor2/Supervisor3 attributes for the DST.SUP groups 
# 02/09/2018 - SAG - Added code to include MailTips
# 04/17/2021 - SAG - Modified creation for the new Alpha Org Supervisor3 now Supervisor2 and Supervisor2 now Supervisor1 also removed configuring the MailContacts on these groups
# 05/27/2021 - SAG - Modified to inlcude creation of location and supervisor people leader groups
# 11/19/2021 - SAG - Modified to add that senderauthentication is required when creating new groups
# 01/04/2022 - SAG - Modified to include creation of AO2 groups
# 08/11/2022 - SAG - Added configuring restriction to the use of the SUP group
# 08/24/2022 - SAG - Modifed to use Menu and Forms
# 10/14/2024 - SAG - Added code to test and create a folder for the current year if it does not exist.
# 11/25/2024 - SAG - Added code to notify Sharepoint and Fabric admins when DSG groups are removed
# 02/24/2026 - SAG - Modified to resolve issues removing DSG groups since MS changed to require MSGraph
#>

Function Build-RemDynGroupsMenu
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Remove DST/DSG Groups Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 480 ; $form.Height = 370  # Make the form wider 
    
    Add-FormStandardButtons
    
    $Left = 140
    $LeftBox = 120
    $TopLoc = 20 
    $Script:lblTitleLine = New-Object System.Windows.Forms.Label   
        $lblTitleLine.Text = "Select Option:"
        $lblTitleLine.Top = 15 ; $lblTitleLine.Left = 100; $lblTitleLine.Width=120 ;$lblTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($lblTitleLine)    # Add to Form 

    ## DST.All Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSTAllGrp = New-Object Windows.Forms.RadioButton 
        $Script:chkDSTAllGrp.Left = $Left; $Script:chkDSTAllGrp.Width = 450; $Script:chkDSTAllGrp.Top = $TopLoc  
        $Script:chkDSTAllGrp.Text = "Remove DST.All Groups" 
        $Script:chkDSTAllGrp.Checked = $false   # set a default value 
        $Script:chkDSTAllGrp.TabIndex = 1
        $Global:form.Controls.Add($Script:chkDSTAllGrp)
        $Script:chkDSTAllGrp.Add_Click({
            $Script:lblGrpLocName.Text = "Enter DST.All Location Name:"
            DSTAllDetails
        })

    ## DST.SUP Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSTSupGrp = New-Object Windows.Forms.RadioButton 
        $Script:chkDSTSupGrp.Left = $Left; $Script:chkDSTSupGrp.Width = 450; $Script:chkDSTSupGrp.Top = $TopLoc  
        $Script:chkDSTSupGrp.Text = "Remove DST.SUP Groups" 
        $Script:chkDSTSupGrp.Checked = $false   # set a default value 
        $Script:chkDSTSupGrp.TabIndex = 2
        $Global:form.Controls.Add($Script:chkDSTSupGrp)
        $Script:chkDSTSupGrp.Add_Click({
            $Script:lblGrpLocName.Text = "Enter DST.SUP Lastname Details:"
            DSTAllDetails
            $Script:ButBldGroups.visible = $false
        })
        
    ## DSG.AO2 Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSGAO2 = New-Object Windows.Forms.RadioButton 
        $Script:chkDSGAO2.Left = $Left; $Script:chkDSGAO2.Width = 450; $Script:chkDSGAO2.Top = $TopLoc  
        $Script:chkDSGAO2.Text = "Remove DSG.AO2 Groups" 
        $Script:chkDSGAO2.Checked = $Script:chkDSGAO2.Checked   # set a default value 
        $Script:chkDSGAO2.TabIndex = 6
        $Global:form.Controls.Add($Script:chkDSGAO2)
        $Script:chkDSGAO2.Add_Click({
            $Script:lblGrpLocName.Text = "Enter DSG.AO2 Alpha Org Level 2 Name:"
            $Script:Subject = "Removal of DSG.AO2 Groups"
            DSTAllDetails
        })         
               
    ## DSG.AO4 Groups
    $TopLoc = $TopLoc + 20
    $Script:chkDSGAO4 = New-Object Windows.Forms.RadioButton 
        $Script:chkDSGAO4.Left = $Left; $Script:chkDSGAO4.Width = 450; $Script:chkDSGAO4.Top = $TopLoc  
        $Script:chkDSGAO4.Text = "Remove DSG.AO4 Groups" 
        $Script:chkDSGAO4.Checked = $false   # set a default value 
        $Script:chkDSGAO4.TabIndex = 5
        $Global:form.Controls.Add($Script:chkDSGAO4)
        $Script:chkDSGAO4.Add_Click({
            $Script:lblGrpLocName.Text = "Enter DSG.AO4 Alpha Org Level 4 Name:"
            $Script:Subject = "Removal of DSG.AO4 Groups"
            DSTAllDetails
        }) 

    $TopLoc = $TopLoc + 20
        $Script:lblSeperatorLine = New-Object System.Windows.Forms.Label   
        $lblSeperatorLine.Text = "____________________________________________________________________"
        $lblSeperatorLine.Top = $TopLoc ; $lblSeperatorLine.Left = 1; $lblSeperatorLine.Width=200 ;$lblSeperatorLine.AutoSize = $true
        $lblSeperatorLine.Visible = $false 
        $Global:form.Controls.Add($lblSeperatorLine)    # Add to Form 

    $TopLoc = $TopLoc + 20
    ## Details to Add Location Information
    $Script:lblGrpLocName = New-Object System.Windows.Forms.Label   
        $Script:lblGrpLocName.Top = $TopLoc ; $Script:lblGrpLocName.Left = $LeftBox; $Script:lblGrpLocName.Width=150 ;$Script:lblGrpLocName.AutoSize = $true
        $Script:lblGrpLocName.visible = $false
        $Global:form.Controls.Add($Script:lblGrpLocName)    # Add to Form 
        # 
        $Script:txtGrpLocName = New-Object Windows.Forms.TextBox  
        $Script:txtGrpLocName.Top = $TopLoc + 20; $Script:txtGrpLocName.Left = $LeftBox; $Script:txtGrpLocName.Width = 300;
        $Script:txtGrpLocName.Text = ""
        $Script:txtGrpLocName.visible = $false
        $Global:form.Controls.Add($Script:txtGrpLocName)    # Add to Form 
        $Script:txtGrpLocName.Add_Click({
            $Script:ButBldGroups.visible = $true
            UpdateLocation
        })

    ## Groups Found
    $Script:lblGrpFound = New-Object System.Windows.Forms.Label   
        $Script:lblGrpFound.Top = 200; $Script:lblGrpFound.Left = 10; $Script:lblGrpFound.Width=110
        $Script:lblGrpFound.Text = "Groups Found:"
        $Script:lblGrpFound.visible = $false
        $Global:form.Controls.Add($Script:lblGrpFound)    # Add to Form 
        # 
        $Script:txtGrpFound = New-Object System.Windows.Forms.ListBox
        $Script:txtGrpFound.Top = 200; $Script:txtGrpFound.Left = $LeftBox; $Script:txtGrpFound.Height = 80; $Script:txtGrpFound.Width = 300;
        $Script:txtGrpFound.TabIndex = 1
        $Script:txtGrpFound.visible = $false
        $Global:form.Controls.Add($Script:txtGrpFound)    # Add to Form

    $Script:ButBldGroups = New-Object Windows.Forms.Button
        $Script:ButBldGroups.Location = New-object System.Drawing.Size(120,200)
        $Script:ButBldGroups.Size = new-Object System.Drawing.Size(150,20)
        $Script:ButBldGroups.Text = "Find Groups"
        $Script:ButBldGroups.visible = $false
        $Global:form.Controls.Add($Script:ButBldGroups)
        $Script:ButBldGroups.Add_Click({
            BuildGroupNames
            $Script:lblGrpFound.visible = $true
            $Script:txtGrpFound.visible = $true
            $Script:ButBldGroups.visible = $false
        })
}

Function DSTAllDetails
{
    $Global:form.Height = 400
    $Script:lblGrpLocName.visible = $true
    $Script:txtGrpLocName.visible = $true
    $Script:ButBldGroups.visible = $true
    $Script:txtGrpLocName.Focus()

    $Script:lblGrpFound.Top = 200
    $Script:lblGrpFound.visible = $false
    $Script:txtGrpFound.Top = 200
    $Script:txtGrpFound.visible = $false
    
    $Script:ButBldGroups.Location = New-object System.Drawing.Size(120,200)
    $Script:ButBldGroups.visible = $true
}

Function UpdateLocation
{
    $Script:lblGrpFound.visible = $false
    $Script:txtGrpFound.visible = $false
    $Script:ButBldGroups.visible = $true
}

Function BuildGroupNames
{
    $Script:txtGrpFound.Items.Clear()
    If (($Script:chkDSTAllGrp.Checked -eq $True) -or ($Script:chkDSTSupGrp.Checked -eq $True))
    {
        $SearchVal = "*" + $Script:txtGrpLocName.Text + "*"
        $Grps = Get-DynamicDistributionGroup -ResultSize Unlimited |Where-Object {$_.DisplayName -like $SearchVal}
    }

    If (($Script:chkDSGAO2.Checked -eq $True) -or ($Script:chkDSGAO4.Checked -eq $True))
    {
        If ($Script:chkDSGAO2.Checked -eq $True)
        {
            $SearchVal = "DSG.AO2." + $Script:txtGrpLocName.Text
        }

        If ($Script:chkDSGAO4.Checked -eq $True)
        {
            $SearchVal = "DSG.AO4." + $Script:txtGrpLocName.Text
        }
        Write-Host $SearchVal
        $Grps = Get-AzureADMSGroup -SearchString $SearchVal
    }
    $Script:lblGrpFound.visible = $true
    $Script:txtGrpFound.visible = $true

    If ($Grps.count -ne 0)
    {
        foreach ($element in $Grps)
        {
            $Name = $element.DisplayName
            [void] $Script:txtGrpFound.Items.Add($Name.TrimStart())
        }
    }
    else
    {
        [void] $Script:txtGrpFound.Items.Add("No Groups Found".TrimStart())
    }
    $Script:txtGrpFound.Refresh()
}

Function Remove-DSTAllGroups
{
    If ($Script:SelGrp -like "DSG*")
    {
        $Script:Subject = "Removal of DSG Group from Azure"
        $Script:Body = "As a result of our weekly Alpha Org2 and Alpha Org4 review the following group was removed from the system:<br><br>"
    }    
    If ($Script:txtGrpFound.SelectedItem -ne $null)
    {
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $ButtonType = [System.Windows.MessageBoxButton]::YesNo
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        $MessageBody = "Delete Only " + $Script:txtGrpFound.SelectedItem + " group?"
        $MessageTitle = "Delete Selected"
        $Script:Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
        If ($Script:Result -eq "Yes")
        {
            write-host "Removing Selected Item" $Script:SelGrp -ForegroundColor Cyan
            $Script:SelGrp = $Script:txtGrpFound.SelectedItem
            Exec-Deletion
        }
        else
        {
            write-host "Request Cancelled" -ForegroundColor Red
        }
    }
    else
    {
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $ButtonType = [System.Windows.MessageBoxButton]::YesNo
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        $MessageBody = "Delete All Groups Listed?"
        $MessageTitle = "Delete All Groups"
        $Script:Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
        If ($Script:Result -eq "Yes")
        {
            foreach ($Script:SelGrp in $Script:txtGrpFound.Items)
            {
                write-host "Removing Group Name: " $Script:SelGrp -ForegroundColor Cyan
                Exec-Deletion
            }
        }
        else
        {
            write-host "Request Cancelled" -ForegroundColor Red
        }
    }
}

Function Exec-Deletion
{
    If (($Script:chkDSTAllGrp.Checked -eq $True) -or ($Script:chkDSTSupGrp.Checked -eq $True))
    {
        $colu = Get-DynamicDistributionGroup $Script:SelGrp
        $DLMem = (Get-Recipient -RecipientPreviewFilter $colu.LdapRecipientFilter -ResultSize Unlimited).count
    }
    else
    {
#        $colu = Get-AzureADGroup -SearchString $Script:SelGrp
        $colu = (Get-AzureADMSGroup -SearchString $Script:SelGrp).Id
#        $colu = Get-MgGroup -Filter "DisplayName eq $Script:SelGrp"
        If ($colu -ne $null)
        {
            $DLMem = (Get-MgGroupMember -All -GroupId $colu).Count
#            $DLMem = (Get-AzureADGroupMember -ALL 1 -ObjectId $colu.ObjectId).count
        }
    }
    
    If ((($DLMem -eq 0) -and ($colu -ne $null)) -or (($AO2Recs."ALPHA_HR_ORG_LEVEL_2" -eq $Script:txtGrpLocName.Text) -or ($AO4Recs."ALPHA_HR_ORG_LEVEL_4" -eq $Script:txtGrpLocName.Text)))
#    If (($DLMem -eq 0) -and ($colu -ne $null))
    {
        $WhoAmI	= WhoAmI
        write-host "`n     Removing group " $Script:SelGrp
        $OutFileName = $Path + "Report-" + ($Script:SelGrp -replace "[ /]","") + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
        write-host "Report file can be found at: " $OutFileName
        $LineToWrite = "STAR" + "`t" + "RemoveDynDistGroup script has started"
        Write-Output $LineToWrite >> $OutFileName
	    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI
        Write-Output $LineToWrite >> $OutFileName
        Write-Output "" >> $OutFileName
        start-transcript
        Write-Host
        Write-Host "     Number of members in " $Script:SelGrp " group " -ForegroundColor Yellow -NoNewline
        Write-Host $DLMem -ForegroundColor Yellow
        write-host ""
        Write-Output "Dynamic Distribution Group Information" >> $OutFileName

        If (($Script:chkDSTAllGrp.Checked -eq $True) -or ($Script:chkDSTSupGrp.Checked -eq $True))
        {
            Get-DynamicDistributionGroup $Script:SelGrp > $OutFileName
            Write-Output "Dynamic Distribution Group Detailed Information" >> $OutFileName
            Get-DynamicDistributionGroup $Script:SelGrp |fl >> $OutFileName
            Write-Output "Dynamic Distribtuion Group Membership Rule" >> $OutFileName
            Get-DynamicDistributionGroup $Script:SelGrp |fl RecipientFilter,IncludedRecipients >>$OutFileName
            Remove-DynamicDistributionGroup $Script:SelGrp -Confirm:$False
        }
        else
        {
            Get-AzureADMSGroup -SearchString $Script:SelGrp >> $OutFileName
#            Get-AzureADGroup -SearchString $Script:SelGrp > $OutFileName
            Write-Output "AzureAD Dynamic Group Detailed Information" >> $OutFileName
            Get-AzureADMSGroup -SearchString $Script:SelGrp |fl >> $OutFileName
#            Get-AzureADGroup -SearchString $Script:SelGrp |fl >> $OutFileName
            Write-Output "AzureAD Dynamic Group Membership Rule" >> $OutFileName
#            Get-AzureADGroup -SearchString $Script:SelGrp |fl MembershipRule >> $OutFileName
            Get-AzureADMSGroup -SearchString $Script:SelGrp |fl MembershipRule >> $OutFileName
            Remove-AzureADMSGroup -ID $colu
            $Script:Body = $Script:Body + "<li>$Script:SelGrp</li>"
        }
    }
    else
    {
        If ($DLMem -ne 0)
        {
            write-host "DLMem Count: " $dlmem
            write-host "Name of Group: " $colu
            write-host "There are still active members of " $Script:SelGrp " group and therefore it will not be removed." -foregroundcolor Red
            write-host "Number of Members: " $mem.count
        }
        else
        {
            write-host "The group " $Script:SelGrp " does not exist." -foregroundcolor Red
        }
    }
}

Function Send-Email
{
    write-host "Sending Emai Message"
    $server = "smtp-relay.ul.com"
    $client = new-object system.net.mail.smtpclient $server
    $from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"
    $to = $from 
    $message = new-object  System.Net.Mail.MailMessage $from, $to 
    $message.IsBodyHtml = $true
    $msgfont = "<basefont face=verdana size=2.5 color=black>"

    $message.Subject = $Script:Subject
    $message.To.Clear()
    $message.Body = $msgfont + $Script:Body
    $message.To.Add("Lucja.Glodny@ul.com,Brandon.Richter@ul.com,Vinay.M.Puligundla@ul.com,Abhishek.Jain@ul.com")
    $message.cc.Add("BJ.Stone@ul.com,Gary.Herbold@ul.com,Sandi.Glazebrook@ul.com")
    $client.Send($message)
}


write-host "Starting.... " -ForegroundColor Green
Build-RemDynGroupsMenu
Publish-Form

If ($Global:Result -eq "OK")
{
    Write-host "Getting AO2 and AO4 Records from the OED7 File"
    $AO2Recs = Import-csv \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract7.csv | Where-Object{($_.ALPHA_HR_ORG_LEVEL_1 -ne "UL Inc.")} | select -unique ALPHA_HR_ORG_LEVEL_2 | Sort-Object ALPHA_HR_ORG_LEVEL_2
    $AO4Recs = Import-csv \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract7.csv | Where-Object{($_.ALPHA_HR_ORG_LEVEL_1 -ne "UL Inc.")} | select -unique ALPHA_HR_ORG_LEVEL_4 | Sort-Object ALPHA_HR_ORG_LEVEL_4
    $Year = (get-date).ToString("yyyy")
    $Path = "e:\Automation\RemoveDynamicGroup\Report\" + $Year + "\"
    If (Test-Path $path) {} else {New-Item -Path $Path -ItemType Directory}

    Remove-DSTAllGroups

    write-host "Script:Body Length: " $Script:Body.Length
    If ($Script:Body.Length -gt 115)
    {
        Send-EMail
    }
}
write-host "`nRemove Dynamic Distribution Group Removal Complete" -ForegroundColor Cyan