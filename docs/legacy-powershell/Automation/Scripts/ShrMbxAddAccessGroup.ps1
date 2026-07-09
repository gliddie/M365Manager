<#
#
#  Called by:  O365MainMenu.ps1
#              ShrMbxAdminMenu.ps1
#
#  11/03/2022 - Adapated from the Change Shared Mailbox Ownership Script so this uses a GUI interface
#>

Function Create-NewAccessGrp #This can replace the NEWUSG and should be moved to the SharedMbxAdminMenu.ps1
{
#This function creates New UserGroups
#   Used by NewSharedMailboxScripts				
    $DSTExists = [bool](Get-DistributionGroup $GrpAddr -ResultSize Unlimited -ErrorAction SilentlyContinue)
    if ($DSTExists -eq $false)
    {
#       write-host $GrpAddr, $GrpName
#        pause

	    if ($GrpAddr.contains("@"))
        {
            # Create the USG
            If ($Script:chkNewEDGrp.Checked -eq $False)
            {
                $DGManagedByMembers = (((Get-DistributionGroup $Script:txtEDGrpName.Text).ManagedBy).Split(",")).Trim()
            }
            else
            {
                $DGManagedByMembers = (($Script:txtGrpOwnr.Text).Split(",")).Trim()
            }
#            $DGManagedByMembers = ($Script:txtGrpOwnr.Text.Split(",")).Trim()
		    New-DistributionGroup -Name $GrpName `
		        -PrimarySmtpAddress $GrpAddr `
			    -Alias $DGAlias `
			    -ManagedBy $DGManagedByMembers `
			    -RequireSenderAuthenticationEnabled $TRUE `
			    -Type Security | Out-Null

            write-host "Pausing process to allow for O365 to complete background processing for the creation of the new group" -ForegroundColor Yellow
            Start-Sleep -Seconds 15

            $Own = (Get-DistributionGroup $GrpName).ManagedBy
            $Script:Owners = "Owners: "
            Foreach ($o in $own)
            {
                $Script:Owners = $Script:Owners + ((get-mailbox $o).DisplayName -split ", ")[1] + " " + ((get-mailbox $o).DisplayName -split ", ")[0] + ", "
            }
            $Script:Owners = $Script:Owners.TrimEnd(", ")
					
            Build-OwnerMailTip
            $Tip =  "Owners: " + $Script:Owners + " - Per: " + $Script:txtTicketNo.Text

#            Set-Group -identity $GrpName -Notes $Tip

            If ($Tip.Length -gt 175)
            {
                Set-DistributionGroup $GrpName -MailTip $Tip.Substring(0,175)
            }
            else
            {
                Set-DistributionGroup $GrpName -MailTip $Tip
            }

            $again = 0
            $ErrorActionPreference = "SilentlyContinue"
            Do
            {
                Set-Group -identity $GrpName -Notes $Tip
                $Verify = get-Group -identity $GrpName
                Start-Sleep -Seconds 3
                $again++
            } while (($Verify.Notes -ne $Tip) -and ($again -lt 10))
            $ErrorActionPreference = "Continue"

            If ($again -ge 10)
            {
                write-host "Unable to update the Notes on th set-group command.  This should match the mailtip on the group" -ForegroundColor Red
            }
        
            if (Get-DistributionGroup $GrpAddr -ErrorAction SilentlyContinue)
            {
		        write-host ""
                write-host "USG created: " $GrpName " (" $GrpAddr ")" -ForegroundColor Cyan
		        $LineToWrite = "CREATE" + "`t" + $GrpName + "`t" + $GrpAddr + " USG Created"
		        WriteReportEvent
			
			    write-host "USG Ownner: " ($DGManagedByMembers -join (", ")) -ForegroundColor Cyan
		        $LineToWrite = "OWNR" + "`t" + ($DGManagedByMembers -join (", "))
		        WriteReportEvent
				
                Set-DistributionGroup $GrpAddr `
			        -RequireSenderAuthenticationEnabled $True `
		 		    -BypassSecurityGroupManagerCheck `
				    -CustomAttribute15 ($FileName + " PS Date: " + (get-date -uformat %D) + " PS Time: " + (get-date -uformat %T))

                Set-DistributionGroup $GrpAddr -MailTip $Tip
            }	
		    else
            {
		        Write-Host "USG not created: " $GrpName " (" $GrpAddr ")" -ForegroundColor Red
			    $LineToWrite = "FAIL" + "`t" + $GrpName + "`t" + $GrpAddr + "`t" + "USG not created"
			    WriteReportEvent
	        }
        }
    }
    else
    {
        Write-Host "User Security Group Already Exists - No Changes Made" -ForegroundColor Red
        pause
    }
}

$Global:OKDetails = "Continue"
$Global:InputFocus = $Script:txtDispName

Build-ShrMbxChgForm
Hide-ShrAddGrp
$Global:form.Text = "Shared Mailbox Additional Access Group"
Publish-Form

If ($Global:Result -eq "OK")
{
	$EDGrpExists = [bool](Get-DistributionGroup $Script:txtEDGrpName.Text -ErrorAction SilentlyContinue)
	$AUGrpExists = [bool](Get-DistributionGroup $Script:txtAUGrpName.Text -ErrorAction SilentlyContinue)
	$REGrpExists = [bool](Get-DistributionGroup $Script:txtREGrpName.Text -ErrorAction SilentlyContinue)
	$Script:GetAlias = ""
	If ($EDGrpExists -eq $True)
    {
        $Script:GetAlias = (Get-DistributionGroup $Script:txtEDGrpName.Text).Alias
    }
    else
    {
        If ($AUGrpExists -eq $True)
        {
            $Script:GetAlias = (Get-DistributionGroup $Script:txtAUGrpName.Text).Alias
        }
        else
        {
			If ($REGrpExists -eq $True)
			{
				$Script:GetAlias = (Get-DistributionGroup $Script:txtREGrpName.Text).Alias
			}
		}       
    }
   
    $ReportFile = $Script:txtRptFile.Text
    $LineToWrite = "Starting New Access Group Creation for Shared Mailbox"
    WriteReportEvent
    $LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
    WriteReportEvent
    $LineToWrite = "INFO" + "`tAdding Access Group to Shared Mailbox"
    WriteReportEvent
    $LineToWrite = "INFO" + "`tTicket Number              :" + "`t" + $Script:txtTicketNo.Text
    WriteReportEvent
    $LineToWrite = "INFO" + "`tShared Mailbox Name        :" + "`t" + $Script:txtDispName.Text
    WriteReportEvent
    $LineToWrite = "INFO" + "`tMailbox Owner(s)           :" + "`t" + $Script:txtGrpOwnr.Text
    WriteReportEvent
    $LineToWrite = "INFO" + "`tMailbox Retention Policy   :" + "`t" + (get-mailbox $Script:txtDispName.Text).RetentionPolicy
    WriteReportEvent
    $LineToWrite = "INFO" + "`tMailbox Retention Enabled  :" + "`t" + (get-mailbox $Script:txtDispName.Text).RetentionHoldEnabled
    WriteReportEvent
    If ($Script:chkNewEDGrp.Checked -eq $True)
    {
        $LineToWrite = "INFO" + "`t.ED Group Name (New)       :" + "`t" + $Script:txtEDGrpName.Text
    }
    else
    {
        $LineToWrite = "INFO" + "`t.ED Group Name             :" + "`t" + $Script:txtEDGrpName.Text
    }
    WriteReportEvent
    If ($Script:chkNewAUGrp.Checked -eq $True)
    {
        $LineToWrite = "INFO" + "`t.AU Group Name (New)       :" + "`t" + $Script:txtAUGrpName.Text
    }
    else
    {
        $LineToWrite = "INFO" + "`t.AU Group Name             :" + "`t" + $Script:txtAUGrpName.Text
    }
    WriteReportEvent
    If ($Script:chkNewAUGrp.Checked -eq $True)
    {
        $LineToWrite = "INFO" + "`t.RE Group Name (New)       :" + "`t" + $Script:txtREGrpName.Text
    }
    else
    {
        $LineToWrite = "INFO" + "`t.RE GRoup Name             :" + "`t" + $Script:txtREGrpName.Text + "`n"
    }
    WriteReportEvent

    $BldOwnr = ($Script:txtGrpOwnr.Text.Split(", "))
	IF ($Script:GetAlias.Length -gt 0)
	{
		$DGManagedByMembers = (Get-DistributionGroup $Script:GetAlias -ErrorAction SilentlyContinue).ManagedBy
		$Script:GrpMem = (Get-DistributionGroup $Script:GetAlias -ErrorAction SilentlyContinue).ManagedBy
	}
	else
	{
		$DGManagedByMembers = $Script:txtGrpOwnr.Text
		$Script:GrpMem = ""
	}
    $Task = $Script:txtTicketNo.Text
    $MbxAddr = (Get-Mailbox $Script:txtDispName.Text).PrimarySMTPAddress
  
    If ($Script:chkNewEDGrp.Checked -eq $True)
    {
        $Script:GrpName = $Script:txtEDGrpName.Text
        $Script:GrpAddr = $Script:txtEDGrpName.Text -replace ("[ -]","")
        $Script:GrpAddr = ($Script:GrpAddr.substring(0,$Script:GrpAddr.Length-2)) + "ED@ul.com"
        $atMail = $Script:GrpAddr.indexOf("@")
        $DGAlias = $Script:GrpAddr.substring(0,$atMail)
        Create-NewAccessGrp
        Add-FolderPermissions
        $LineToWrite = "INFO" + "`tCompleted creation and addition of new Editor Group`n"
        WriteReportEvent
    }

    If ($Script:chkNewAUGrp.Checked -eq $True)
    {
        $Script:GrpName = $Script:txtAUGrpName.Text
        $Script:GrpAddr = $Script:txtAUGrpName.Text -replace ("[ -]","")
        $Script:GrpAddr = ($Script:GrpAddr.substring(0,$Script:GrpAddr.Length-2)) + "AU@ul.com"
        $atMail = $Script:GrpAddr.indexOf("@")
        $DGAlias = $Script:GrpAddr.substring(0,$atMail)
        Create-NewAccessGrp
        Add-FolderPermissions
        $LineToWrite = "INFO" + "`tCompleted creation and addition of new Authors Group`n"
        WriteReportEvent
    }

    If ($Script:chkNewREGrp.Checked -eq $True)
    {
        $Script:GrpName = $Script:txtREGrpName.Text
        $Script:GrpAddr = $Script:txtREGrpName.Text -replace ("[ -]","")
        $Script:GrpAddr = ($Script:GrpAddr.substring(0,$Script:GrpAddr.Length-2)) + "RE@ul.com"
        $atMail = $Script:GrpAddr.indexOf("@")
        $DGAlias = $Script:GrpAddr.substring(0,$atMail)
        Create-NewAccessGrp
        Add-FolderPermissions
        $LineToWrite = "INFO" + "`tCompleted creation and addition of new Readers Group`n"
        WriteReportEvent
    }

    $LineToWrite = "New Acess Group Creation Process Completed"
    WriteReportEvent
}
else
{
    $Output = $wshell.Popup("Shared mailbox ownership change request cancelled.",0,"Cancelled",0+32)
}
