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
#                  : SAG - 09/14/2021 Added code so if Exchange Management Tools are not installed this is reported back to the user
#
#################################################################################

$me = whoami
$dir = "c:\users\" + $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)) + "\documents\"
$File = "my" + ($me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))).replace(".","") + "File.xml"
$moveCredFile = "c:\temp\" + $File
$Global:CredFile = $dir + $File

If (Test-Path $moveCredFile)
{
	move-item $moveCredFile -destination $Global.CredFile
}

$Session = get-PSSession

If ($Session.State -eq "Broken")
{
    Remove-PSSession $Session
    $Session = get-PSSession
}

if ($Session -eq $null)
{
    write-host "Connecting to Office 365....."
    If ((Test-path $Global:CredFile) -ne "True")
    {
        Get-Credential | Export-Clixml $Global:CredFile
    }

    $Global:LiveCred = Import-Clixml $Global:CredFile

    Connect-ExchangeOnline -ConnectionUri "https://outlook.office365.com/powershell-liveid?SerializationLevel=Full" -Credential $Global:LiveCred
	$Script:ExSnap = [bool](Add-PSSnapin *Exchange* -erroraction SilentlyContinue)
    If ($ExSnap -eq $True)
    {
        Add-PSSnapin *Exchange*
    }
    else
    {
        write-host "Exchange Management Tools are not installed on this device, you cannot enable new account for email on this machine" -ForegroundColor Red
    }
    Import-Module ActiveDirectory
    Import-Module MSOnline
    Import-Module AzureADPreview
    write-host "Connecting to SharePointOnlineService...... " -ForegroundColor Cyan
    Connect-SPOService -url "https://ul-admin.sharepoint.com" -credential $Global:LiveCred
    write-host "Connecting to MSOLService...... " -ForegroundColor Cyan
    Connect-MsolService -Credential $Global:LiveCred
    write-host "Connecting to MicrosoftTeams...... " -ForegroundColor Cyan
    Connect-MicrosoftTeams -Credential $Global:LiveCred
	write-host "Connecting to AzureAD...... " -ForegroundColor Cyan
    Remove-PSSnapin *Exchange*
    Add-PSSnapin *Exchange*	
	Connect-AzureAD -Credential $Global:LiveCred | Out-Null
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection
    Import-PSsession $Session -AllowClobber
}
else
{
	write-host "Session with Office 365 already exists." -ForegroundColor Yellow
	write-host ""
}
