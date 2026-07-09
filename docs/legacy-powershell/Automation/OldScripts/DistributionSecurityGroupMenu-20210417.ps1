<#
#
#  Called by:  O365MainMenu.ps1
#
#  08/09/2020 - Changed over to a GUI Interface
#  04/17/2021 - Commented out the connect to O365
#>

function Build-DLMenuForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Distribution List / Security and O365 Group Admin Menu" 
    $Global:form.StartPosition = "CenterScreen" 
    $Global:form.Width = 640 ; $form.Height = 650  # Make the form wider 
    #Add Buttons- ## Create the button panel to hold the OK and Cancel buttons 

    $TopLoc = 20 

    ## Label and TextBox  
    ## Title Line
    $Global:lblDLSecTitleLine = New-Object System.Windows.Forms.Label   
        $lblDLSecTitleLine.Text = "Distribution List/Security Group Action"
        $lblDLSecTitleLine.Top = 15 ; $lblDLSecTitleLine.Left = 60; $lblDLSecTitleLine.Width=220 ;$lblDLSecTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($lblDLSecTitleLine)    # Add to Form 

    ## Add/Remove Aliases
    $TopLoc = $TopLoc + 20
    $Global:chkAddRemAlias = New-Object Windows.Forms.checkbox 
        $Global:chkAddRemAlias.Left = 100; $Global:chkAddRemAlias.Width = 450; $Global:chkAddRemAlias.Top = $TopLoc  
        $Global:chkAddRemAlias.Text = "Add/Remove Additional eMail Alias to Distribution List" 
        $Global:chkAddRemAlias.Checked = $false   # set a default value 
        $Global:chkAddRemAlias.TabIndex = 1
        $Global:form.Controls.Add($Global:chkAddRemAlias) 
        # Obtain Value with: $Global:chkAddRemAlias.Checked
        
    ## New DL
    $TopLoc = $TopLoc + 20
    $Global:chkNewDL = New-Object Windows.Forms.checkbox 
        $Global:chkNewDL.Left = 100; $Global:chkNewDL.Width = 450; $Global:chkNewDL.Top = $TopLoc
        $Global:chkNewDL.Text = "Create a New Distribtuion Group" 
        $Global:chkNewDL.Checked = $false   # set a default value 
        $Global:chkNewDL.TabIndex = 2
        $Global:form.Controls.Add($Global:chkNewDL) 
        # Obtain Value with: $Global:chkNewDL.Checked
        $Global:InputFocus = $Global:chkNewDL

    ## Change DL/Sec/O365 Group
    $TopLoc = $TopLoc + 20
    $Global:chkChgDLSec = New-Object Windows.Forms.checkbox 
        $Global:chkChgDLSec.Left = 100; $Global:chkChgDLSec.Width = 450; $Global:chkChgDLSec.Top = $TopLoc  
        $Global:chkChgDLSec.Text = "Change Distribution List/Security Group/Unified Group Ownership" 
        $Global:chkChgDLSec.Checked = $false   # set a default value 
        $Global:chkChgDLSec.TabIndex = 3
        $Global:form.Controls.Add($Global:chkChgDLSec) 
        # Obtain Value with: $Global:chkChgDLSec.Checked

    ## Change DL Restriced Access
    $TopLoc = $TopLoc + 20
    $Global:chkDLRestrict = New-Object Windows.Forms.checkbox 
        $Global:chkDLRestrict.Left = 100; $Global:chkDLRestrict.Width = 450; $Global:chkDLRestrict.Top = $TopLoc  
        $Global:chkDLRestrict.Text = "Change Distribution List Restricted Access"
        $Global:chkDLRestrict.Checked = $false   # set a default value 
        $Global:chkDLRestrict.TabIndex = 3
        $Global:form.Controls.Add($Global:chkDLRestrict) 
        # Obtain Value with: $Global:chkDLRestrict.Checked

    ## DL/Sec Group Membership
    $TopLoc = $TopLoc + 20
    $Global:chkDLMem = New-Object Windows.Forms.checkbox 
        $Global:chkDLMem.Left = 100; $Global:chkDLMem.Width = 450; $Global:chkDLMem.Top = $TopLoc  
        $Global:chkDLMem.Text = "Count of Distribution/Security Group Membership" 
        $Global:chkDLMem.Checked = $Global:chkDLMem.Checked   # set a default value 
        $Global:chkDLMem.TabIndex = 4
        $Global:form.Controls.Add($Global:chkDLMem) 
        # Obtain Value with: $Global:chkDLMem.Checked

    ## External Access
    $TopLoc = $TopLoc + 20
    $Global:chkExtAccess = New-Object Windows.Forms.checkbox 
        $Global:chkExtAccess.Left = 100; $Global:chkExtAccess.Width = 450; $Global:chkExtAccess.Top = $TopLoc  
        $Global:chkExtAccess.Text = "Grant/Deny External Addresses Access to Distribution List or Security Group" 
        $Global:chkExtAccess.Checked = $false   # set a default value 
        $Global:chkExtAccess.TabIndex = 5
        $Global:form.Controls.Add($Global:chkExtAccess) 
        # Obtain Value with: $Global:chkExtAccess.Checked

    ## Add Members
    $TopLoc = $TopLoc + 20
    $Global:chkAddMem = New-Object Windows.Forms.checkbox 
        $Global:chkAddMem.Left = 100; $Global:chkAddMem.Width = 450; $Global:chkAddMem.Top = $TopLoc  
        $Global:chkAddMem.Text = "M&A Activities Menu" 
        $Global:chkAddMem.Checked = $false   # set a default value 
        $Global:chkAddMem.TabIndex = 6
        $Global:form.Controls.Add($Global:chkAddMem) 
        # Obtain Value with: $Global:chkAddMem.Checked

    ## Remove DL/Security Group 
    $TopLoc = $TopLoc + 20
    $Global:chkRemDLSecGrp = New-Object Windows.Forms.checkbox 
        $Global:chkRemDLSecGrp.Left = 100; $Global:chkRemDLSecGrp.Width = 450; $Global:chkRemDLSecGrp.Top = $TopLoc  
        $Global:chkRemDLSecGrp.Text = "Remove Distribution List or Security Group" 
        $Global:chkRemDLSecGrp.Checked = $false   # set a default value 
        $Global:chkRemDLSecGrp.TabIndex = 7 
        $Global:form.Controls.Add($Global:chkRemDLSecGrp) 
        # Obtain Value with: $Global:chkRemDLSecGrp.Checked

    ## Remove Members
    $TopLoc = $TopLoc + 20
    $Global:chkRemMem = New-Object Windows.Forms.checkbox 
        $Global:chkRemMem.Left = 100; $Global:chkRemMem.Width = 450; $Global:chkRemMem.Top = $TopLoc  
        $Global:chkRemMem.Text = "Remove Member(s) from a Distribution List/Security or Unified Group" 
        $Global:chkRemMem.Checked = $false   # set a default value 
        $Global:chkRemMem.TabIndex = 8
        $Global:form.Controls.Add($Global:chkRemMem) 
        # Obtain Value with: $Global:chkRemMem.Checked

    ## Rename DL/Security Group
    $TopLoc = $TopLoc + 20
    $Global:chkRenDL = New-Object Windows.Forms.checkbox 
        $Global:chkRenDL.Left = 100; $Global:chkRenDL.Width = 450; $Global:chkRenDL.Top = $TopLoc  
        $Global:chkRenDL.Text = "Rename Distribution List or Security Group" 
        $Global:chkRenDL.Checked = $false   # set a default value 
        $Global:chkRenDL.TabIndex = 9
        $Global:form.Controls.Add($Global:chkRenDL) 
        # Obtain Value with: $Global:chkRenDL.Checked

    ## Replace DL/Security Group Membership
    $TopLoc = $TopLoc + 20
    $Global:chkReplMem = New-Object Windows.Forms.checkbox 
        $Global:chkReplMem.Left = 100; $Global:chkReplMem.Width = 450; $Global:chkReplMem.Top = $TopLoc  
        $Global:chkReplMem.Text = "Replace Distribution List or Security Group Membership" 
        $Global:chkReplMem.Checked = $false   # set a default value 
        $Global:chkReplMem.TabIndex = 10
        $Global:form.Controls.Add($Global:chkReplMem) 
        # Obtain Value with: $Global:chkReplMem.Checked

    ## Update DL/Security Group
    $TopLoc = $TopLoc + 20
    $Global:chkUpdDLSecGrp = New-Object Windows.Forms.checkbox 
        $Global:chkUpdDLSecGrp.Left = 100; $Global:chkUpdDLSecGrp.Width = 450; $Global:chkUpdDLSecGrp.Top = $TopLoc  
        $Global:chkUpdDLSecGrp.Text = "Update (Add/Remove) Distribution List or Security Group Membership" 
        $Global:chkUpdDLSecGrp.Checked = $false   # set a default value 
        $Global:chkUpdDLSecGrp.TabIndex = 11
        $Global:form.Controls.Add($Global:chkUpdDLSecGrp) 
        # Obtain Value with: $Global:chkUpdDLSecGrp.Checked
 
    ## Title Line
    $TopLoc = $TopLoc + 40
    $Global:lblDynTitleLine = New-Object System.Windows.Forms.Label   
        $Global:lblDynTitleLine.Text = "Dynamic Distribution Group Actions"
        $Global:lblDynTitleLine.Top = $TopLoc ; $Global:lblDynTitleLine.Left = 60; $Global:lblDynTitleLine.Width=120 ;$Global:lblDynTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDynTitleLine)    # Add to Form 
    
    
    ## New DST Group
    $TopLoc = $TopLoc + 20
    $Global:chkNewDST = New-Object Windows.Forms.checkbox 
        $Global:chkNewDST.Left = 100; $Global:chkNewDST.Width = 450; $Global:chkNewDST.Top = $TopLoc  
        $Global:chkNewDST.Text = "Create New UL Employee Only and UL Staff Dynamic Distribution Groups" 
        $Global:chkNewDST.Checked = $false   # set a default value 
        $Global:chkNewDST.TabIndex = 12
        $Global:form.Controls.Add($Global:chkNewDST) 
        # Obtain Value with: $Global:chkNewDST.Checked

    ## DST Members
    $TopLoc = $TopLoc + 20
    $Global:chkDSTMem = New-Object Windows.Forms.checkbox 
        $Global:chkDSTMem.Left = 100; $Global:chkDSTMem.Width = 450; $Global:chkDSTMem.Top = $TopLoc  
        $Global:chkDSTMem.Text = "List of DST Membership" 
        $Global:chkDSTMem.Checked = $false   # set a default value 
        $Global:chkDSTMem.TabIndex = 13
        $Global:form.Controls.Add($Global:chkDSTMem) 
        # Obtain Value with: $Global:chkDSTMem.Checked

    ## Remove DST Group
    $TopLoc = $TopLoc + 20
    $Global:chkRemDST = New-Object Windows.Forms.checkbox 
        $Global:chkRemDST.Left = 100; $Global:chkRemDST.Width = 450; $Global:chkRemDST.Top = $TopLoc  
        $Global:chkRemDST.Text = "Remove Dynamic Distribution List" 
        $Global:chkRemDST.Checked = $false   # set a default value 
        $Global:chkRemDST.TabIndex = 14
        $Global:form.Controls.Add($Global:chkRemDST) 
        # Obtain Value with: $Global:chkRemDST.Checked

    ## Title Line
    $TopLoc = $TopLoc + 40
    $Global:lblUniGrpTitleLine = New-Object System.Windows.Forms.Label   
        $Global:lblUniGrpTitleLine.Text = "O365 Groups/Teams Actions"
        $Global:lblUniGrpTitleLine.Top = $TopLoc ; $Global:lblUniGrpTitleLine.Left = 60; $Global:lblUniGrpTitleLine.Width=120 ;$Global:lblUniGrpTitleLine.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblUniGrpTitleLine)    # Add to Form 

    ## Check Team/Sharepoint Site
    $TopLoc = $TopLoc + 20
    $Global:chkTeamShrPtSite = New-Object Windows.Forms.checkbox 
        $Global:chkTeamShrPtSite.Left = 100; $Global:chkTeamShrPtSite.Width = 450; $Global:chkTeamShrPtSite.Top = $TopLoc  
        $Global:chkTeamShrPtSite.Text = "Check for Team SharePoint Site" 
        $Global:chkTeamShrPtSite.Checked = $false   # set a default value 
        $Global:chkTeamShrPtSite.TabIndex = 15
        $Global:form.Controls.Add($Global:chkTeamShrPtSite) 
        # Obtain Value with: $Global:chkTeamShrPtSite.Checked

    ## SoftDeleted Site
    $TopLoc = $TopLoc + 20
    $Global:chkSoftDelObj = New-Object Windows.Forms.checkbox 
        $Global:chkSoftDelObj.Left = 100; $Global:chkSoftDelObj.Width = 450; $Global:chkSoftDelObj.Top = $TopLoc  
        $Global:chkSoftDelObj.Text = "Check for SoftDeleted Team SharePoint Sites" 
        $Global:chkSoftDelObj.Checked = $false   # set a default value 
        $Global:chkSoftDelObj.TabIndex = 16
        $Global:form.Controls.Add($Global:chkSoftDelObj) 
        # Obtain Value with: $Global:chkSoftDelObj.Checked

    ## Create O365/Team
    $TopLoc = $TopLoc + 20
    $Global:chkCreUniGrp = New-Object Windows.Forms.checkbox 
        $Global:chkCreUniGrp.Left = 100; $Global:chkCreUniGrp.Width = 450; $Global:chkCreUniGrp.Top = $TopLoc  
        $Global:chkCreUniGrp.Text = "Create O365 Groups (Teams/Unified Groups)" 
        $Global:chkCreUniGrp.Checked = $Global:chkCreUniGrp.Checked   # set a default value 
        $Global:chkCreUniGrp.TabIndex = 17
        $Global:form.Controls.Add($Global:chkCreUniGrp) 
        # Obtain Value with: $Global:chkCreUniGrp.Checked

    ## Remove Uni Mem
    $TopLoc = $TopLoc + 20
    $Global:chkUniGrpMem = New-Object Windows.Forms.checkbox 
        $Global:chkUniGrpMem.Left = 100; $Global:chkUniGrpMem.Width = 500; $Global:chkUniGrpMem.Top = $TopLoc  
        $Global:chkUniGrpMem.Text = "Remove Member(s) from a Distribution List/Security or Unified Group" 
        $Global:chkUniGrpMem.Checked = $false   # set a default value 
        $Global:chkUniGrpMem.TabIndex = 18
        $Global:form.Controls.Add($Global:chkUniGrpMem) 
        # Obtain Value with: $Global:chkUniGrpMem.Checked

    ## Misc Powershell Commands
    $TopLoc = $TopLoc + 20
    $Global:chkO356GrpTeams = New-Object Windows.Forms.checkbox 
        $Global:chkO356GrpTeams.Left = 100; $Global:chkO356GrpTeams.Width = 450; $Global:chkO356GrpTeams.Top = $TopLoc  
        $Global:chkO356GrpTeams.Text = "View All O365 Groups Enabled for Teams" 
        $Global:chkO356GrpTeams.Checked = $false   # set a default value 
        $Global:chkO356GrpTeams.TabIndex = 19
        $Global:form.Controls.Add($Global:chkO356GrpTeams) 
        # Obtain Value with: $Global:chkO356GrpTeams.Checked

    Add-FormStandardButtons
}

#### Start of Script
write-host ""

Do
{
    Build-DLMenuForm
    Publish-Form
        		
    If ($Global:chkNewDL.Checked -eq "Checked")
    {
        Invoke-Expression -Command .\NewDLGroup.ps1
    }

    If ($Global:chkReplMem.Checked -eq "Checked")
    {
        Invoke-Expression -Command .\ReplGrpMembers.ps1
    }

    If ($Global:chkRenDL.Checked -eq "Checked")
    {
        Invoke-Expression -Command .\RenameDLGroup.ps1
    }

    If ($Global:chkChgDLSec.Checked -eq "Checked")
    {
        $AnoGrp = "Y"
        DO
        {
            write-host ""
            write-host "     Enter ( 1) Change Ownership for a Single Group"
            write-host "           ( 2) Change Ownership for Multiple Groups"
            write-host ""
            write-host "           ( 0) to Return to the Distribution List and Security Group Admin Menu"
            write-Host "     Enter Option? " -ForegroundColor Red -NoNewline
            $NoGrps = Read-Host
            switch ($NoGrps)
            {
                1
                {
                    invoke-expression -Command .\GroupOwnershipChanges.ps1
                }
                2
                {
                    invoke-expression -Command .\GroupOwnershipChangeMultiple.ps1
                }
            }
            write-host "Change Membership on more Distribution Lists (Y/N)? " -ForegroundColor Yellow -NoNewline
            $AnoGrp = Read-Host
        } while ($anoGrp -eq "Y")
    }

    If ($Global:chkDLRestrict.Checked -eq "Checked")
    {
        Invoke-Expression -Command .\RestrictedDLAccess.ps1
        Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\DLRestrictedAccessGranted.oft        
    }

    If ($Global:chkRemDLSecGrp.Checked -eq "Checked")
    {
        invoke-expression -Command .\RemoveDLGroup.ps1
    }

    If ($Global:chkUpdDLSecGrp.Checked -eq "Checked")
    {
        invoke-expression -command .\UpdateGroupMembership.ps1
        write-host""
        write-host "List Updating Complete" -ForegroundColor Magenta
    }

    If ($Global:chkDLMem.Checked -eq "Checked")
    {
        Write-Host "Enter the Name of the group that you would like the get the membership for " -ForegroundColor Yellow -NoNewline
        $Grp = Read-Host
    	$DynDstExits = [bool](Get-DynamicDistributionGroup $Grp -ErrorAction SilentlyContinue)
        if ($DynDstExits = "True")
        {
            $colu = Get-DistributionGroupMember $Grp -ResultSize Unlimited
            Write-Host "Number of members in the group " -ForegroundColor Yellow -NoNewline
            Write-Host $Grp -ForegroundColor Red -NoNewline
            Write-Host " has " -ForegroundColor Yellow -NoNewline
            write-host $colu.count "members" -ForegroundColor Red
            pause
        }
        else
        {
             Write-Host "Dynamic Distribution List Does Not Exist - No Changes Made" -ForegroundColor Red
        }
    }

    If ($Global:chkExtAccess.Checked -eq "Checked")
    {
        Write-Host "Enter the Name of the group that you would like Grant or Deny External Addresses to use: " -ForegroundColor Yellow -NoNewline
        $Grp = Read-Host
        
        $DstExists = [bool](Get-DistributionGroup $Grp -ErrorAction SilentlyContinue)
        if ($DstExists -eq "True")
        {
            $GrpExt = Get-DistributionGroup $Grp
            write-host "Group Owners:  " $GrpExt.ManagedBy

            if ($GrpExt.RequireSenderAuthenticationEnabled -eq "True")
            {
                write-host "Is the requestor an owner of this distribution list? (Y/N) " -ForegroundColor Yellow -NoNewline
                $Owner = Read-host

                If ($Owner -eq"Y")
                {
                    Write-Host "External addresses are not allowed to send to this group would you like to allow this (Y/N) " -ForegroundColor Yellow -NoNewline
                    $ExtChg = Read-Host
                    if ($ExtChg -eq "Y")
                    {
                        Set-DistributionGroup $Grp -RequireSenderAuthenticationEnabled $false
                        Write-Host "Setting changed to allow external addresses to send to the group it may take 15-20 minutes for this setting to synchronize in the O365 environment" -ForegroundColor Yellow
                    }
                    else
                    {
                        Write-Host "No changes made to the settings for allowing external addresses to send to the group"
                    }
                    pause
                }
                else
                {
                    write-host "No changes made without a curent owner approving the changes" -ForegroundColor Red
                }
            }
            else
            {
                write-host "Is the requestor an owner of this distribution list? (Y/N) " -ForegroundColor Yellow -NoNewline
                $Owner = Read-host

                If ($Owner -eq "Y")
                {
                    Write-Host "External addresses are allowed to send to this group would you like to disallow this (Y/N) " -ForegroundColor Yellow -NoNewline
                    $ExtChg = Read-Host
                    if ($ExtChg -eq "Y")
                    {
                        Set-DistributionGroup $Grp -RequireSenderAuthenticationEnabled $true
                        Write-Host "Setting changed to not allow external addresses to send to the group it may take 15-20 minutes for this setting to synchronize in the O365 environment" -ForegroundColor Yellow
                    }
                    else
                    {
                        Write-Host "No changes made to the settings for allowing external addresses to send to the group"
                    }
                    pause
                }
                else
                {
                    write-host "No changes made without a curent owner approving the changes" -ForegroundColor Red
                }
            }
        }
        else
        {
            Write-Host "Distribution List Does Not Exist - No Changes Made" -ForegroundColor Red
        }                 
    }

    If ($Global:chkAddRemAlias.Checked -eq "Checked")
    {
        Write-Host "Add/Remove Additional eMail Alias to Distribution Group" -ForegroundColor Magenta
        Write-Host "Enter the name of the Distribution List " -ForegroundColor Yellow -NoNewline
        $Grp = Read-Host
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

    If ($Global:chkAddMem.Checked -eq "Checked")
    {
        #  Declare Drive | Folders | and Files
	    $FileName		= "AddDgMembers"
	    $LogDrive		= "e:\Automation"
	    $LogFolder		= "\" + $FileNAme
	    $LogDirectory	= $LogDrive + $LogFolder + "\"
	    $LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	    $InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
	    $ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"

        if (Test-Path $InputFile)
        {
            write-host "Importing the e:\Automation\AddDGMembers\Input-AddDgMembers.csv Input file" -ForegroundColor Yellow
            $InpFile = Import-Csv $InputFile
            $DstExists = [bool](Get-DistributionGroup $InpFile.DgMail -ErrorAction SilentlyContinue)
            $UfgExists = [bool](Get-UnifiedGroup $InpFile.DgMail -ErrorAction SilentlyContinue)
            If (($DstExists -eq "True") -or ($UfgExists -eq "True"))
            {
                $Members = $InpFile.MemberMail.Replace(" ","")
    	        $addMember = $Members.split(",")
                write-host "Number of Members Listed in Input File: " $addmember.count
                if ($? -eq $true)
                {
                    $CurrentMembers = Get-DistributionGroupMember $InpFile.DgMail -ResultSize Unlimited
		            ForEach ($member in $addMember)
                    {
                        if ($member.contains("@"))
                        {
	                        if (Get-Mailbox $member -ErrorAction SilentlyContinue)
                            {
                                if ($CurrentMembers -match (get-mailbox $Member).Name)
                                {
                                    write-host $member " already a member."
           				            $LineToWrite = "`t" + "Warn" + "`t" + $member + " - is already a member of Distribution Group. " + $Grp.GrpName
                                }
                                else
                                {
                                    If ($DstExists -eq "True")
                                    {
                                        Add-DistributionGroupMember $InpFile.DgMail -Member $Member -BypassSecurityGroupManagerCheck
    	  				                Write-Host "Added member:" $member
		    			                $LineToWrite = "`t" + "PASS" + "`t" + $member + "`t" + "Added to Distribution/Security Group" + "`t" + $InpFile.DgDisplayName + "`n"
                                    }
                                    else
                                    {
                                        Add-UnifiedGroupLinks -Identity $InpFile.DgMail -LinkType Member -links $Member -Confirm:$false								    
    					                Write-Host "Added member:" $member
						                $LineToWrite = "`t" + "PASS" + "`t" + $member + "`t" + "Added to Unified Group" + "`t" + $InpFile.DgDisplayName + "`n"
                                    }
                                }
					        }
		    		        else
                            {
						        write-host "ERROR finding member: " $member
						        $LineToWrite = "`t" + "FAIL" + "`t" + $member + "`t" + "ERROR finding recipient" + "`t" + $InpFile.DgDisplayName + "`n"
					        }
				        }
				        else
                        {
                            write-host "ERROR invalid member: " $member
					        $LineToWrite = "`t" + "FAIL" + "`t" + $member + "`t" + "ERROR invalid recipient" + "`t" + $InpFile.DgDisplayName + "`n"
		                }
                        WriteReportEvent
                    }
                    Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")
                }
            }	
		    else
            {
                write-host "Distribution/Security or Unified Group Does Not Exist"
		    }
        }
        else
        {
            Write-Host "Enter the name of the Security Group, Distribution List or Unified Group " -ForegroundColor Yellow -NoNewline
            $Grp = Read-Host
            $DstExists = [bool](Get-DistributionGroup $Grp -ErrorAction SilentlyContinue)
            $UfgExists = [bool](Get-UnifiedGroup $Grp -ErrorAction SilentlyContinue)
            If (($DstExists -eq "True") -or ($UfgExists -eq "True"))
            {
                If ($DstExists -eq "True")
                {
                    $GrpDet = Get-DistributionGroup $Grp
                }
                else
                {
                    $GrpDet = Get-UnifiedGroup $Grp
                }
                write-host ""
                write-host "Distribution/Security/Unified Group Owner(s): " $GrpDet.ManagedBy 
                write-host ""
                write-host "Is the requestor an owner of this security or distribution list/unified group or have you obtained owner approval (Y/N)? " -ForegroundColor Red -NoNewline
                $OwnrAns = Read-Host
                
                If ($OwnrAns -eq "Y")
                {
                    Do
                    {
                        Write-Host "Enter the Employee Number of the individual to add to the membership " -ForegroundColor Yellow -NoNewline
                        $Member = Read-Host
                        $Member = $Member + "@global.ul.com"
                        
                        If ($DstExists -eq "True")
                        {
                            Add-DistributionGroupMember $grp -Member $Member -BypassSecurityGroupManagerCheck
                        }
                        else
                        {
                            Add-UnifiedGroupLinks -Identity $Grp -LinkType Members -links $Member -Confirm:$false
                        }

                        Write-Host $Member "added to " $Grp
                        Write-Host ""
                        Write-Host "Do you have more members to add to this group (Y/N)? " -ForegroundColor Yellow -NoNewline
                        $OwnrAns = Read-Host
                    } while  ($OwnrAns -eq "Y")
                }
           }
           else
           {
                Write-Host "Distribution/Security or Unified Group Does Not Exist" -ForegroundColor Red
           }
        }
        else
        {
            write-host "No changes made without owner approval"
        }
    }
    
    If (($Global:chkRemMem.Checked -eq "Checked") -or ($Global:chkUniGrpMem.Checked -eq "Checked"))
    {
        #  Declare Drive | Folders | and Files
        $FileName		= "RemoveDgMembers"
	    $LogDrive		= "e:\Automation"
	    $LogFolder		= "\" + $FileNAme
	    $LogDirectory	= $LogDrive + $LogFolder + "\"
	    $LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	    $InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
	    $ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"

        if (Test-Path $InputFile)
        {
            write-host "Importing the e:\Automation\RemoveDGMembers\Input-RemoveDgMembers.csv Input file" -ForegroundColor Yellow
            $InpFile = Import-Csv $InputFile
            $DstExists = [bool](Get-DistributionGroup $InpFile.DgMail -ErrorAction SilentlyContinue)
            $UfgExists = [bool](Get-UnifiedGroup $InpFile.DgMail -ErrorAction SilentlyContinue)
            If (($DstExists -eq "True") -or ($UfgExists -eq "True"))
            {
                $Members = $InpFile.MemberMail.Replace(" ","")
                $removeMember = $Members.split(",")
                write-host "Number of Members Listed in Input File: " $removemember.count
	            if ($? -eq $true)
                {
                    $CurrentMembers = Get-DistributionGroupMember $InpFile.DgMail -ResultSize Unlimited
	                ForEach ($member in $removeMember)
                    {
		                if ($member.contains("@"))
                        {
	                        if (Get-Mailbox $member -ErrorAction SilentlyContinue)
                            {
                                if ($CurrentMembers -match (get-mailbox $Member).Name)
                                {
                                    write-host $member " already a member."
        		                    $LineToWrite = "`t" + "Warn" + "`t" + $member + " - is already a member of Distribution Group. " + $Grp.GrpName
                                }
                                else
                                {
                                    If ($DstExists -eq "True")
                                    {
                                        Remove-DistributionGroupMember $InpFile.DgMail -Member $Member -BypassSecurityGroupManagerCheck -Confirm:$false
    	    		                    Write-Host "Added member:" $member
		        	                    $LineToWrite = "`t" + "PASS" + "`t" + $member + "`t" + "Removed from Distribution/Security Group" + "`t" + $InpFile.DgDisplayName + "`n"
                                    }
                                    else
                                    {
                                        Remove-UnifiedGroupLinks -Identity $InpFile.DgMail -LinkType Member -links $Member -Confirm:$false								    
    			                        Write-Host "Added member:" $member
					                    $LineToWrite = "`t" + "PASS" + "`t" + $member + "`t" + "Removed from Unified Group" + "`t" + $InpFile.DgDisplayName + "`n"
                                    }
                                }
					        }
		    		        else
                            {
                                write-host "ERROR finding member: " $member
						        $LineToWrite = "`t" + "FAIL" + "`t" + $member + "`t" + "ERROR finding recipient" + "`t" + $InpFile.DgDisplayName + "`n"
					        }
				        }
	                    else
                        {
		                    write-host "ERROR invalid member: " $member
			                $LineToWrite = "`t" + "FAIL" + "`t" + $member + "`t" + "ERROR invalid recipient" + "`t" + $InpFile.DgDisplayName + "`n"
			            }
                        WriteReportEvent
			        }
                    Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")
                }
            }	
            else
            {
                write-host "Distrubtion/Security or Unified Group Does Not Exist"
	        }
        }
        else
        {
            Write-Host "Enter the name of the Security Group, Distribution List or Unified Group " -ForegroundColor Yellow -NoNewline
            $Grp = Read-Host
            $DstExists = [bool](Get-DistributionGroup $Grp -ErrorAction SilentlyContinue)
            $UfgExists = [bool](Get-UnifiedGroup $Grp -ErrorAction SilentlyContinue)
            If (($DstExists -eq "True") -or ($UfgExists -eq "True"))
            {
                If ($DstExists -eq "True")
                {
                    $GrpDet = Get-DistributionGroup $Grp
                }
                else
                {
                    $GrpDet = Get-UnifiedGroup $Grp
                }
                write-host ""
                write-host "Distribution/Security/Unified Group Owner(s): " $GrpDet.ManagedBy 
                write-host ""
                write-host "Is the requestor an owner of this security or distribution list/unified group or have you obtained owner approval (Y/N)? " -ForegroundColor Red -NoNewline
                $OwnrAns = Read-Host
                
                If ($OwnrAns -eq "Y")
                {
                    Do
                    {
                        Write-Host "Enter the Employee Number of the individual to remove to the membership " -ForegroundColor Yellow -NoNewline
                        $Member = Read-Host
                        $Member = $Member + "@global.ul.com"
                        
                        If ($DstExists -eq "True")
                        {
                            Remove-DistributionGroupMember $grp -Member $Member -BypassSecurityGroupManagerCheck -Confirm:$false
                        }
                        else
                        {
                            Remove-UnifiedGroupLinks -Identity $Grp -LinkType Members -links $Member -Confirm:$false
                        }

                        Write-Host $Member "removed to " $Grp
                        Write-Host ""
                        Write-Host "Do you have more members to remove to this group (Y/N)? " -ForegroundColor Yellow -NoNewline
                        $OwnrAns = Read-Host
                    } while  ($OwnrAns -eq "Y")
                }
            }
            else
            {
                
                Write-Host "Distribution/Security or Unified Group Does Not Exist" -ForegroundColor Red
            }
        }
        else
        {
            write-host "No changes made without owner approval"
        }
    }

    If ($Global:chkDSTMem.Checked -eq "Checked")
    {
        Write-Host "Enter the Name of the DST group that you would like the membership for " -ForegroundColor Yellow -NoNewline
        $DSTGrp = Read-Host
        $DstExists = [bool](Get-DynamicDistributionGroup $DSTGrp -ErrorAction SilentlyContinue)
        if ($DstExists -eq "True")
        {
            $g = Get-DynamicDistributionGroup $DSTGrp
            $colu = Get-Recipient -RecipientPreviewFilter $g.LdapRecipientFilter -ResultSize Unlimited
            Write-Host
            Write-Host "     Number of members in this group " -ForegroundColor Yellow -NoNewline
            Write-Host $colu.count -ForegroundColor Yellow
            Write-Host
        	Write-Host "Writing list of group members to c:\temp\DSTMembership.csv"
			$colu.PrimarySMTPAddress | Out-File c:\temp\DSTMembership.csv
            Write-Host "Request complete hit return to continue" -NoNewline
            $Cont = Read-Host 
        }
        else
        {
            Write-Host "Distribution/Security List Does Not Exist - No Changes Made" -ForegroundColor Red
        }
    }

    If ($Global:chkNewDST.Checked -eq "Checked") 
    {
        Invoke-Expression .\NewDynamicDistributionGroup.ps1
    }

    If ($Global:chkRemDST.Checked -eq "Checked") 
    {
        Invoke-Expression .\RemoveDynamicDistributionGrp.ps1
    }

    If ($Global:chkCreUniGrp.Checked -eq "Checked") 
    {
        Invoke-Expression .\NewUnifiedGrp.ps1
    }

    If ($Global:chkTeamShrPtSite.Checked -eq "Checked")
    {
        Write-Host "Enter the Name of the O365 Group to check for an active Sharepoint Site: " -ForegroundColor Yellow -NoNewline
        $Grp = Read-Host
        $O365Exists = [bool](Get-UnifiedGroup $Grp -ErrorAction SilentlyContinue)
        if ($O365Exists -eq "True")
        {
            $O365Grp = Get-UnifiedGroup $Grp
            $SPSiteExists = [bool](get-sposite $O365Grp.SharePointSiteUrl)
            If ($SPSiteExists -eq "True")
            {
                write-host "SharePointSite" $0365Grp.SharePointSiteUrl "Exists" -ForegroundColor Green
            }
            else
            {
                write-host "No SharepointSite " $0365Grp.SharePointSiteUrl "Exists" -ForegroundColor Red
            }
        }
        else
        {
            write-host "The O365 Group" $O365Grp "does not exist" -ForegroundColor Red
        }
    }

    If ($Global:chkSoftDelObj.Checked -eq "Checked")
    {
        write-host "Connecting to AzureAD...... " -ForegroundColor Cyan
     	Connect-AzureAD -Credential $Global:LiveCred | Out-Null
        Write-Host "Enter the Name of the O365 Group to check for a soft deleted Sharepoint Site: " -ForegroundColor Cyan -NoNewline
        $Grp = Read-Host
        $O365Exists = [bool](Get-AzureADMSDeletedGroup -SearchString $Grp -ErrorAction SilentlyContinue)
        if ($O365Exists -eq "True")
        {
            $O365Grp = Get-AzureADMSDeletedGroup -SearchString $Grp
            $SPSiteExists = [bool](get-SPODeletedSite $O365Grp.SharePointSiteUrl)
            If ($SPSiteExists -eq "True")
            {
                write-host "SharePointSite" $0365Grp.SharePointSiteUrl "Exists. Do you want to permanently Delete it (Y/N)?: " -ForegroundColor Green -NoNewline
                $SPODel = Read-Host
                If ($SPODel -eq "Y")
                {
                    Write-Host "Permanently removing the soft deleted sharepoint site" -ForegroundColor Yellow
                    Remove-SPODeletedSite $O365Grp.SharePointSiteUrl -Confirm:$False
                }
                else
                {
                    write-host "Not removing the soft deleted sharepoint site" -ForegroundColor Red
                }
            }
            else
            {
                write-host "No SharepointSite " $0365Grp.SharePointSiteUrl "Exists" -ForegroundColor Red
            }
        }
        else
        {
            write-host "A deleted O365 Group named" $Grp "does not exist" -ForegroundColor Red
        }
#      	Remove-AzureADMSDeletedDirectoryObject -Id $O365Grp.Id
        pause
    }

    If ($Global:chkMiscCmd.Checked -eq "Checked")
    {
        write-host "Enter (1) to Show All O365 Groups or (2) to Show a Specific Group: " -ForegroundColor Cyan -NoNewline
        $ShwGrps = Read-Host
               
        switch ($ShwGrps)
        {

            1
            {
                $o365groups = Get-UnifiedGroup
                foreach ($o365group in $o365groups)
                {
                    $TeamEna = "Enabled"
                    try
                    {
                        $teamschannels = Get-TeamChannel -GroupId $o365group.ExternalDirectoryObjectId
                        [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; TeamsEnabled = $true}
                    }
                    catch
                    {
                        $ErrorCode = $_.Exception.ErrorCode
                        switch ($ErrorCode)
                        {
                            "404"
                            {
                               [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; TeamsEnabled = $false}
                               $TeamEna = "NotEnabled"
                               break;
                            }
                            "403"
                            {
                                [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; TeamsEnabled = $true}
                                break;
                            }
                            default
                            {
                                Write-Error ("Unknown ErrorCode trying to 'Get-TeamChannel -GroupId {0}' :: {1}" -f $o365group, $ErrorCode)
                                $TeamEna = "Unknown"
                            }
                        }
                    }
                }
            }
            2
            {
                Write-host "Enter the Name of the Group: " -ForegroundColor Cyan -NoNewline
                $Grp = Read-Host
                $o365groups = Get-UnifiedGroup $Grp
                foreach ($o365group in $o365groups)
                {
                    $TeamEna = "Enabled"
                    try
                    {
                        $teamschannels = Get-TeamChannel -GroupId $o365group.ExternalDirectoryObjectId
                        [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; TeamsEnabled = $true}
                    }
                    catch
                    {
                        $ErrorCode = $_.Exception.ErrorCode
                        switch ($ErrorCode)
                        {
                            "404"
                            {
                                [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; TeamsEnabled = $false}
                                $TeamEna = "NotEnabled"
                                break;
                            }
                            "403"
                            {
#                                [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; GroupName = $o365group.DisplayName; TeamsEnabled = $true}
                                break;
                            }
                            default
                            {
                                Write-Error ("Unknown ErrorCode trying to 'Get-TeamChannel -GroupId {0}' :: {1}" -f $o365group, $ErrorCode)
                                $TeamEna = "Unknown"
                            }
                        }
                    }
                    write-host "Status:    " $TeamEna
                    write-host "GroupID:   " $O365group.ExternalDirectoryObjectId
                    write-host "GroupName: " $o365group.DisplayName
                }
                pause
            }
        }
    }

}While ($Global:DLResult -eq "OK")	