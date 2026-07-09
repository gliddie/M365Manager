#
#	Connect to O365 Services
#	
#	Check to see if there is a c:\temp directory if not create one
#	Also check to see if there is a Credential file for this individual if
#	  if not create one.
#
#################################################################################
# 
# PowerShell source code
# Revision v1.01
# ==========================================================================
#    'Project      : SDAP Menu Creation
#    'Description  : Connect to O365 Environment
#    'Called By    : SDAPAdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 1/11/2015
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '             : SAG - 11/16/2017 Changed the $LiveCred and $CredFile to be global parameters
#    '             : SAG - 02/15/2018 Added Command to connect to MicrosoftTeams
#    '             : SAG - 02/19/2018 Added code to check with the credential file was last updated
#    '             : SAG - 07/03/2020 Added connection to the Security & Compliance Center
#    '             : SAG - 09/14/2021 Added code to check if the Exchange Management Tools are installed and report this to the user
#
#################################################################################

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

Import-Module ActiveDirectory
$ADUser = Get-ADUser $CredENo -Properties PasswordLastSet
$PwdLastSet = $ADUser.PasswordLastSet
$CredFileDate = "0"

If (Test-Path $Global:CredFile)
{
    $CredFileDate = (Get-ChildItem $Global:CredFile).LastWriteTime
}

If ($Global:CredFileDate -ge $PwdLastSet)
{
    write-host "Credential File is Up to Date." -ForegroundColor Green
}
else
{
   write-host "`nYour password was recently changed and your Credential file needs to be updated." -ForegroundColor Red
   write-host "Please make sure to enter your username as emp#@global.ul.com" -ForegroundColor Red
   Invoke-Expression -Command .\CreateNewCredFile.ps1
}

<#$Session = get-PSSession

if ($Session -eq $null)
{
#>
    write-host "Connecting to Office 365....."
    If ((Test-path $Global:CredFile) -ne "True")
    {
        Get-Credential | Export-Clixml $Global:CredFile
    }

    $Global:LiveCred = Import-Clixml $Global:CredFile
#    Import-Module ExchangeOnlineManagement
    write-host "Connecting to ExchangeOnlineService...... " -ForegroundColor Cyan
    Connect-ExchangeOnline -Credential $Global:LiveCred
#    Connect-ExchangeOnline -ConnectionUri "https://outlook.office365.com/powershell-liveid?SerializationLevel=Full" -Credential $Global:LiveCred
    Import-Module MSOnline
    Import-Module AzureADPreview
    write-host "Connecting to SharePointOnlineService...... " -ForegroundColor Cyan
    Connect-SPOService -url "https://ul-admin.sharepoint.com" -credential $Global:LiveCred
    write-host "Connecting to MSOLService...... " -ForegroundColor Cyan
    Connect-MsolService -Credential $Global:LiveCred
    write-host "Connecting to MicrosoftTeams...... " -ForegroundColor Cyan
    Connect-MicrosoftTeams -Credential $Global:LiveCred |Out-Null
    Add-PSSnapin *Exchange* -erroraction SilentlyContinue
    $Script:ExSnap = [bool](Get-PSSnapin *Exchange* -ErrorAction SilentlyContinue)
    If ($Script:ExSnap -eq $True)
    {
#        Remove-PSSnapin *Exchange*
#        Add-PSSnapin *Exchange*
        write-host "Connected to Exchange Management Snap-In...... " -ForegroundColor Cyan
    }
    else
    {
        write-host "Exchange Management Tools are not installed you cannot enable accounts for email on this device...... " -ForegroundColor Red
    }
	write-host "Connecting to AzureAD...... " -ForegroundColor Cyan
	Connect-AzureAD -Credential $Global:LiveCred | Out-Null
#    write-host "Connecting to Security and Compliance Center...... " -ForegroundColor Cyan
#    Connect-SecurityCompliance -Credential $Global:LiveCred
<#
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $LiveCred -Authentication Basic -AllowRedirection
    write-host "Connecting to Exchange Online Powershell Commands...... " -ForegroundColor Cyan
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
    Import-PSsession $Session -AllowClobber
    Import-PSSession $Session -DisableNameChecking -AllowClobber
}
else
{
	write-host "Session with Office 365 already exists." -ForegroundColor Yellow
	write-host ""
}
#>

