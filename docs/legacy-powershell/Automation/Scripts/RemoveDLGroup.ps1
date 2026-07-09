<#
#
#  Called by:  DistributionSecurityGroupMenu.ps1
#              RoomResourceAdminMenu.ps1
#
#  06/14/2018 - Added to to delete Unified Groups
#  08/04/2018 - Added code if there are no members in a DST group to remove without confirming
#  01/28/2019 - Added -ResultSize Unlimited to the get-DistributionListMember statement
#  04/08/2019 - Added -ResultSize Unlimited to the line that initially reports how many individuals in the group
#  12/10/2019 - Added code so if the name does not follow the naming standard the requestor identifies the group type
#  08/08/2020 - Changed over to a GUI Interface
#  05/19/2022 - Added option to not display the template after removing a group.
#  06/01/2022 - Fixed the location where the option to display the template appears.
#  11/02/2022 - Modified to use the default DL Group forms
#  04/19/2024 - Modified the Unified group to check to see if this is a Microsoft Team and if so to add aditional reporting and to use the Remove-Team command
#>

Build-DefaultForm
$Global:form.Text = "Remove Distribution/Security/O365 Group"
$Global:okButton.Text = "Remove" 
Add-ActionBoxes
Add-FormStandardButtons
Hide-Details
If (($Script:txtCurrOwner.Text.Length -le 0) -and ($Script:txtCurrMbr.Text.Length -le 0))
{
    $Script:txtTicketNo.Text = "NoRemainingOwnersOrMembers"
}  
Publish-Form


If ($Global:Result -eq "OK")
{
    $Reason = "Reason for Removal: " + $Script:txtTicketNo.Text
    $Year = (get-date).ToString("yyyy")
    $Path = "e:\Automation\RemoveDLGroup\Report\" + $Year
    If (Test-Path $path) {} else {New-Item -Path $Path -ItemType Directory}

    $ModInputDL = ($Script:txtDLName.Text.Replace(".","")).Trim()
    $OutFileName = $Path + "\Report-RemoveDLGroup-" + ($Script:txtDLName.Text -replace "[. ]","") + ".txt"

    If ($Script:DLExists -eq "True")
    {
        Write-Output "Removing Distribution List" > $OutFileName
        Write-Output $Reason >> $OutFileName
        Write-Output ("Launched by: " + $WhoAmI + "`n") >> $OutFileName
        Write-Output "Distribution Group Information" >> $OutFileName
        get-DistributionGroup $Script:txtDLName.Text |ft Name,DisplayName,GroupType,PrimarySMTPPAddress >> $OutFileName
        Write-Output "Distribution Group Detailed Information" >> $OutFileName
        Write-Output (get-DistributionGroup $Script:txtDLName.Text) |fl >> $OutFileName
#        get-DistributionGroup $Script:txtDLName.Text |fl >> $OutFileName
        Write-Output "Distribution Group Manager and Notes" >> $OutFileName
        Write-Output (get-DistributionGroup $Script:txtDLName.Text) |fl ManagedBy,Notes >> $OutFileName
#        get-Group $Script:txtDLName.Text |fl ManagedBy,Notes >> $OutFileName
        $mem = (get-DistributionGroupMember $Script:txtDLName.Text -ResultSize Unlimited) |Sort-Object {$_.PrimarySMTPAddress}
        Write-Output ("Distribution Group Membership Count: " +  $mem.count) >> $OutFileName
        Write-Output "Distribution Group Membership" >> $OutFileName
        Write-Output $mem |ft Alias,PrimarySMTPAddress,RecipientType >> $OutFileName
#        Write-Output (get-DistributionGroupMember $Script:txtDLName.Text -ResultSize Unlimited|ft Alias,PrimarySMTPAddress,RecipientType) >> $OutFileName
        remove-DistributionGroup $Script:txtDLName.Text -BypassSecurityGroupManagerCheck -Confirm:$False
        write-host "Distribtuion Group Removal Complete" -ForegroundColor Red
    }

    If ($Script:UniExists -eq "True")
    {
        $check = Get-UnifiedGroup $Script:txtDLName.Text
        If ($check.ResourceProvisioningOptions -contains "Team")
#        If ([bool](Get-Team -GroupID $check.ExternalDirectoryObjectId))
        {
            $Type = "Team "
            Write-Output "Removing Microsoft Team Group" > OutFileName
        }
        else
        {
            $Type = "Unified "
            Write-Output "Removing Unified Group" > $OutFileName
        }
        Write-Output $Reason >> $OutFileName
        Write-Output ("Launched by: " + $WhoAmI + "`n") >> $OutFileName
        Write-Output ($Type + "Group Information") >> $OutFileName
        Write-Output (get-UnifiedGroup $Script:txtDLName.Text) |ft Name,DisplayName,GroupType,PrimarySMTPAddress >> $OutFileName
        Write-Output ($Type + "Group Detailed Information") >> $OutFileName
        Write-Output (get-UnifiedGroup $Script:txtDLName.Text) |fl >> $OutFileName
        Write-Output ($Type + "Group Owners") >> $OutFileName
        Write-Output (Get-UnifiedGroupLinks $Script:txtDLName.Text -LinkType Owners).PrimarySMTPAddress |ft Name >>$OutFileName
        Write-Output ("`n" + $Type + "Group Members ") >> $OutFileName
        Write-Output (Get-UnifiedGroupLinks $Script:txtDLName.Text -LinkType Members).PrimarySMTPAddress |ft Name >>$OutFileName
<#
        If ($check.ResourceProvisioningOptions -contains "Team")
#        If ([bool](Get-Team -GroupID $check.ExternalDirectoryObjectId))
        {
            Write-Output "Microsoft Team Group Information" >> $OutFileName
            Write-Output (Get-Team -GroupID $check.ExternalDirectoryObjectId) |ft Name,DisplayName,GroupType,PrimarySMTPAddress >> $OutFileName
            Write-Output ($Type + "Detailed Information") >> $OutFileName
            Write-Output (Get-Team -GroupID $check.ExternalDirectoryObjectId) |ft DisplayName,MailNickName,PrimarySmtpAddress,Visibility >> $OutFileName
        }
        else
        {
            Write-Output ($Type + "Group Information") >> $OutFileName
            Write-Output (get-UnifiedGroup $Script:txtDLName.Text) |ft Name,DisplayName,GroupType,PrimarySMTPAddress >> $OutFileName
            Write-Output ($Type + "Group Detailed Information") >> $OutFileName
            Write-Output (get-UnifiedGroup $Script:txtDLName.Text) |fl >> $OutFileName
        }
        
        Write-Output ($Type + "Group Owners") >> $OutFileName
        Write-Output (Get-UnifiedGroupLinks $Script:txtDLName.Text -LinkType Owners).PrimarySMTPAddress |ft Name >>$OutFileName
        Write-Output ("`n" + $Type + "Group Members ") >> $OutFileName
        Write-Output (Get-UnifiedGroupLinks $Script:txtDLName.Text -LinkType Members).PrimarySMTPAddress |ft Name >>$OutFileName
        write-host "Removing Object"
#>
        If ($check.ResourceProvisioningOptions -contains "Team")
        {
#           Remove-Team -GroupId $check.ExternalDirectoryObjectId -confirm:$False
            Remove-Team -GroupId $check.ExternalDirectoryObjectId
        }
        else
        {
#            Remove-UnifiedGroup -Identity $Script:txtDLName.Text -confirm:$False
            Remove-UnifiedGroup -Identity $Script:txtDLName.Text
        }
        write-host "Object Removal Complete"
        write-host $Type "Group Removal Complete" -ForegroundColor Green
    }

    If ($Global:chkTemplate.Checked -eq $True)
    {
        Invoke-Expression -Command e:\O365AdminShared\EMailTemplates\DLRemoval.oft
    }

    Write-Output ("`n" + $Type + "Group Removal Complete") >> $OutFileName
}
else
{
    $Output = $wshell.Popup("As requested group removal has been cancelled",0,"Do Not Continue",0+32)
}
