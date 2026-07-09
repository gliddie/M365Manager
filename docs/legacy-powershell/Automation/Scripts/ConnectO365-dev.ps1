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
#
#################################################################################

$me = whoami
$dir = "c:\temp\my96151File-Dev.xml"
#" + $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)) + "\documents\"
#$File = "my" + ($me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1))).replace(".","") + "File.xml"
#$moveCredFile = "c:\temp\" + $File
$Global:CredFile = "C:\temp\my96151File-Dev.xml"
#$dir + $File

#If (Test-Path $moveCredFile)
#{
#	move-item $moveCredFile -destination $Global:CredFile
#}

$Session = get-PSSession

if ($Session -eq $null)
{
    write-host "Connecting to Office 365 DEV Environment....."

    If ((Test-path $Global:CredFile) -ne "True")
    {
		Get-Credential | Export-Clixml "C:\temp\my96151File-Dev.xml"
    }

	$Global:LiveCred = Import-Clixml $Global:CredFile
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Global:LiveCred -Authentication Basic -AllowRedirection

    Add-PSSnapin *Exchange* -erroraction SilentlyContinue
    Import-Module ActiveDirectory
    Import-Module MSOnline
    Import-Module AzureADPreview
    write-host "Connecting to SharePointOnlineService...... " -ForegroundColor Cyan
    Connect-SPOService -url "https://uldev-admin.sharepoint.com" -credential $Global:LiveCred
    write-host "Connecting to MSOLService...... " -ForegroundColor Cyan
    Connect-MsolService -Credential $Global:LiveCred
    write-host "Connecting to MicrosoftTeams...... " -ForegroundColor Cyan
    Connect-MicrosoftTeams -Credential $Global:LiveCred
#	write-host "Connecting to AzureAD...... " -ForegroundColor Cyan
	Connect-AzureAD -Credential $Global:LiveCred | Out-Null
    Import-PSsession $Session -AllowClobber
}
else
{
	write-host "Session with Office 365 already exists." -ForegroundColor Yellow
	write-host ""
}
