<#################################################################################
# 
# PowerShell source code
# Revision v1.05
# ==========================================================================
#    'Project      : UL O365 Admin Team Menu
#    'Description  : O365 Admin Team Main Menu
#    'Called By    :
#    'Calls        : 
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 07/28/2015
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '            11/15/2017 - SAG - Added code to provide a credential file 
#                    and to update the credential file when passwords change
#                    Modified to call the ConnectO365.ps1 from a shared space
#    '            03/07/2018 - SAG - Added a reconnect threshold of 60 minutes
#    '            08/29/2019 - SAG - Adding menu items for configuring surface hubs
#    '            10/04/2019 - SAG - Added feature to ask if more devices need to be configured for Service Accounts
#    '            04/30/2020 - SAG - Added code to create groups and configure mailbox permissions for Service Accounts
#    '            05/12/2020 - SAG - Started to add in the CommonFunctions script
# ==========================================================================
#
#################################################################################>

$Global:LiveCred = ""
$Global:CredFile = ""
$File = "e:\Automation\scripts\O365AdminMenu.ps1"

#invoke-expression -Command E:\O365AdminShared\Scripts\CommonFunctions.ps1

if (test-path $file)
{
#   Do nothing E" drive exists
}
else
{
    New-PSDrive -Name E -PSProvider FileSystem -Root \\USNBKM100P\e$
}

Set-Location E:\Automation\Scripts
write-host ""

$reconnectThreshold = New-TimeSpan -Minutes 60
invoke-expression -Command E:\O365AdminShared\Scripts\ConnectO365.ps1
$stopwatch = [diagnostics.stopwatch]::StartNew()
$O365Act = "1"

Do
{
    write-host ""
    write-host "O365 Admin Menu" -ForegroundColor Magenta
    write-host
    Write-Host "     Enter (  1 ) Legal Hold Admin Menu"
    write-host "           (  2 ) Unified Messaging Menu"
    write-host "           (  3 ) User Mailbox Menu"
    write-host "           (  4 ) Distribution List/O365 Group/Security Group Admin Menu"
    write-host "           (  5 ) Shared Mailbox Admin Menu"
    write-host "           (  6 ) Room or Resource Admin"
    write-host "           (  7 ) Create Discovery Search Mailbox"
    write-host "           (  8 ) Get LockedPC Information"
    write-host "           (  9 ) M&A Activities Menu"
    write-host "           ( 10 ) Licensing Menu"
    write-host "           ( 11 ) Reporting"
    write-host "           ( 12 ) Add a New Domain to the Tenant"
    write-host "           ( 12a) Add/Enalbe DKIM Keys for New Domain"
    write-host "           ( 13 ) Enable/Disable Compromised Account"
    write-host "           ( 14 ) Configure Service Accounts/Surface Hub Devices"
	write-host "           ( 15 ) Get Users OneDrive Space Usage"
    write-host ""
    write-host "           (20) Create new credential file"
    write-host "           (21) Open miscellaneous powershell commands document"
    write-host ""
    write-Host "           ( 0) No Changes and Exit"
    write-Host "     Enter Option Above? " -ForegroundColor Red -NoNewline
    $O365Act = Read-Host
    write-host ""

    if (($stopwatch.elapsed -ge $reconnectThreshold) -and ($O365Act -ne 0))
    {
        # Close all sessions
        write-host "Connection Threshold Exceeded -- Reconnecting to Office365" -ForegroundColor Red
        get-pssession | Remove-PSSession -Confirm:$false
        invoke-expression -Command E:\O365AdminShared\Scripts\ConnectO365.ps1
        $stopwatch = [diagnostics.stopwatch]::StartNew()
    }
        		
		switch ($O365Act)
		{
            1
                {
                    invoke-expression -Command .\LegalHoldCheck.ps1
                }
            2
                {
                    invoke-expression -Command .\EUMMenu.ps1
                }
            3
                {
                    invoke-expression -Command .\UserMailboxMenu.ps1
                }
            4
                {
                    invoke-expression -Command .\DistributionSecurityGroupMenu.ps1
                }
            5
                {
                    invoke-expression -Command .\SharedMailboxAdminMenu.ps1
                }
            6
                {
                    invoke-expression -Command .\RoomResourceAdminMenu.ps1
                }
            7
                {
                    Write-Host "Enter Name of Discovery Mailbox:  " -ForegroundColor Yellow -NoNewline
                    $DiscMbx = Read-Host
                    $DiscMbxAddr = ($DiscMbx -replace '\s','') + "@ul.onmicrosoft.com"
                    
                    New-mailbox $DiscMbx –Discovery –PrimarySMTPAddress $DiscMbxAddr 
                    Add-MailboxPermission $DiscMbx –AccessRights FullAccess –User “Discovery Management{8f1d3a49ccc740dcb9d8c3cd06b95011}”

                    Write-Host "Discovery Mailbox Created:  " $DiscMbx
                    pause

                }
            8
                {
                    Invoke-Expression -Command .\Get-UserLockedPCName.ps1
                }
            9
                {
                    Invoke-Expression -Command .\MAActivitesMenu.ps1
                }              
            10
                {
                    Invoke-Expression -Command .\LicensingMenu.ps1
                }
            11
                {
                    Invoke-Expression -Command .\ReportingMenu.ps1
                }
            12
                {
                    write-host "Enter the Domain to add: " -ForegroundColor Cyan -NoNewline
                    $NewDom = Read-Host
                    $NewDom = $NewDom.TrimStart("@")
                    If (($NewDom -like "*ul.com*") -or ($NewDom -like "*ul.org*"))
                    {
                        New-MsolDomain -Name $NewDom -Authentication Federated
                        $RcvMail = "N"
                        write-host "Will Email to this domain be received by a mailbox in the UL environment (Y/N)? " -ForegroundColor Cyan -NoNewline
                        $RcvMail = Read-Host
                        If ($RcvMail -eq "Y")
                        {
                            write-host "Setting the Domain for O365 to be Authoritative" -ForegroundColor Yellow
                            Set-AcceptedDomain $NewDom -DomainType Authoritative
                        }
                    }
                    else
                    {
                        New-MsolDomain -Name $NewDom
                        write-host "Setting the Domain to be an InternalRelay" -ForegroundColor Yellow
                        Set-AcceptedDomain $NewDonm -DomainType InternalRelay
                    }
                }
            12a
                {
                    write-host "Enter the Domain to add: " -ForegroundColor Cyan -NoNewline
                    $NewDom = Read-Host
                    $NewDom = $NewDom.TrimStart("@")
                    If (($NewDom -like "*ul.com*") -or ($NewDom -like "*ul.org*"))
                    {
                        New-DkimSigningConfig -DomainName $NewDom -Enabled $true 
                        write-host "DKIM Configured for: " $NewDom 
                    }
                }
            13
                {
                    Invoke-Expression -Command .\CompromisedAccount.ps1
                }
            14
                {
                    write-host ""
                    Write-Host "`n     Enter ( 1) Enable/Configure UM Mailbox"
                    write-host "           ( 2) Enable/Configure Oracle OFR/KFI Mailbox"
                    write-host "           ( 3) Enable/Configure Surface Hub Device"
                    write-host "           ( 4) Enable/Configure Remote Assist Account"
                    write-host ""
                    write-Host "           ( 0) No Changes and Exit"
                    write-Host "     Enter Option Above? " -ForegroundColor Red -NoNewline
                    $EnabMbx = Read-Host
                    $Repeat = "N"
                    Do
                    {
                        switch ($EnabMbx)
	                    {
                        
                            2
                            {
                                Remove-PSSession (Get-PSSession)
                                Invoke-Expression -Command .\EnableSvcActAssignLicense.ps1
                                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
                                Import-PSSession $session -AllowClobber
#                           Create USG group to control access
                                Write-Host "Creating groups to control access to the mailbox.  Hit return when the NewUSG input file has ready" -ForegroundColor Red -NoNewline
                                $Ret = Read-Host
                                Invoke-Expression -Command .\NewUSG.ps1
#                           Add Mailbox/Folder Permissions and Apply Retention Policy
                                Write-Host "Applying folder permissions and retention policy to the mailbox.  Hit return when the AddFolderPermissionsinput file has ready and the mailbox has been provisioned." -ForegroundColor Red -NoNewline
                                $Ret = Read-Host
                                Invoke-Expression -Command .\AddFolderPermissions.ps1
                           }
                        
                            3
                            {
                                Remove-PSSession (Get-PSSession)
                                Invoke-Expression -Command .\SetupSurfaceHubDevice.ps1
                                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
                                Import-PSSession $session -AllowClobber
                            }

                            Default
                            {
                                Remove-PSSession (Get-PSSession)
                                Invoke-Expression -Command .\EnableSvcActAssignLicense.ps1
                                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
                                Import-PSSession $session -AllowClobber
#                           Create USG group to control access
                                Write-Host "Creating groups to control access to the mailbox.  Hit return when the NewUSG input file has ready" -ForegroundColor Red -NoNewline
                                $Ret = Read-Host
                                Invoke-Expression -Command .\NewUSG.ps1
#                           Add Mailbox/Folder Permissions and Apply Retention Policy
                                Write-Host "Applying folder permissions and retention policy to the mailbox.  Hit return when the AddFolderPermissions input file has ready" -ForegroundColor Red -NoNewline
                                $Ret = Read-Host
                                Invoke-Expression -Command .\AddFolderPermissions.ps1
                            }
                        }
                        write-host "`nDo you need to enable another device of this type (Y/N)? " -ForegroundColor Cyan -NoNewline
                        $Repeat = Read-Host
                    }while($Repeat -eq "Y")
                }
			15
			{
				write-host "Enter Employee Number to get usage details for: "
				$EmpID = read-host
				$EmpID = $EmpID + "_global_ul_com"
				write-host "Connecting to SharePoint"
				invoke-expression -Command .\ConnectO365SPO.ps1
				$Site = Get-SPOSite ("https://ul-my.sharepoint.com/personal/"+ $EmpID)
				write-host "Current Site Usage: " $site.StorageUsageCurrent
			}
            20
                {
                    Invoke-Expression -Command .\CreateNewCredFile.ps1
                }
            21
                {
                    invoke-item '.\Instructions\O365 powershell commands.docx'
                }               
   
        }
}While ($O365Act -ne 0)	