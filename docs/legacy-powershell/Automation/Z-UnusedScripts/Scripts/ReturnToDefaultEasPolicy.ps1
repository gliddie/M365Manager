#################################################################################
# 
# PowerShell source code
# Revision v1.00
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online
#    'Description  : Return Default ActiveSyncMailboxPolicy to default settings
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


Set-ActiveSyncMailboxPolicy `
	-Identity:Default `
	-AllowPOPIMAPEmail:$true `
	-MaxAttachmentSize:unlimited `
	-MaxEmailBodyTruncationSize:unlimited `
	-MaxEmailHTMLBodyTruncationSize:unlimited `
	-AllowNonProvisionableDevices:$true `
	-AllowUnsignedApplications:$true `
	-AllowUnsignedInstallationPackages:$true `
	-DevicePolicyRefreshInterval:unlimited `
	-RequireManualSyncWhenRoaming:$false `
	-UNCAccessEnabled:$true `
	-DeviceEncryptionEnabled:$false `
	-RequireStorageCardEncryption:$true `
	-DevicePasswordEnabled:$false `
	-MaxDevicePasswordFailedAttempts:unlimited `
	-MaxInactivityTimeDeviceLock:unlimited `
	-MinDevicePasswordComplexCharacters:1 `
	-MinDevicePasswordLength:$null `
	-PasswordRecoveryEnabled:$false `
	-RequireDeviceEncryption:$false

