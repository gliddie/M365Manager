#################################################################################
# 
# PowerShell source code
# Revision DRAFT v1.00
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online
#    'Description  : Modify Default ActiveSyncMailboxPolicy to UL settings
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 04/06/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '               
#    '
# ==========================================================================
#
#################################################################################

$LiveCred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
Import-PSSession $Session

# Modify Default ActiveSyncMailboxPolicy to UL settings
Set-ActiveSyncMailboxPolicy `
	-Identity:Default `
	-AllowPOPIMAPEmail:$false `
	-MaxAttachmentSize:26MB `
	-MaxEmailBodyTruncationSize:50KB `
	-MaxEmailHTMLBodyTruncationSize:100KB `
	-AllowNonProvisionableDevices:$false `
	-AllowUnsignedApplications:$false `
	-AllowUnsignedInstallationPackages:$false `
	-DevicePolicyRefreshInterval:20 `
	-RequireManualSyncWhenRoaming:$true `
	-UNCAccessEnabled:$false `
	-DeviceEncryptionEnabled:$true `
	-RequireStorageCardEncryption:$false `
	-DevicePasswordEnabled:$true `
	-MaxDevicePasswordFailedAttempts:16 `
	-MaxInactivityTimeDeviceLock:01:00:00 `
	-MinDevicePasswordComplexCharacters:1 `
	-MinDevicePasswordLength:4 `
	-PasswordRecoveryEnabled:$true `
	-RequireDeviceEncryption:$true

