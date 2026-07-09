#
#   Give Manager Access to OneDrive
#
#   03/06/2019 - SAG - Added code so that the manager is given read-only access to OneDrive
#

    invoke-expression -Command .\ConnectO365SPO.ps1
#    $AdminURL = "https://ul-admin.sharepoint.com"
#    $AdminName = "96151@global.ul.com"
#    $Password = Read-host -assecurestring "Enter Password for $AdminName"
#    $Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Password
#    Connect-SPOService -url $AdminURL -credential $Credential

	write-Host "Enter Terminated Staff Employee Number: " -ForegroundColor Green -NoNewline
	$EmpID = Read-host
    $EmpID = $EmpID + "_global_ul_com"
    $ManagerID = Read-Host "Enter the Manager Request Access Employee Number: " -ForegroundColor Green -NoNewline
    $LoginName = ($ManagerID + "@global.ul.com")

    $Site = Get-SPOSite ("https://ul-my.sharepoint.com/personal/"+ $EmpId)
#	Set the individual running the script admin access
	Set-SPOUser -site $Site -LoginName $LoginName -IsSiteCollectionAdmin $True 

#   Create a Readers Group and add the individual into the readers group
	If ([bool](get-spositegroup -site $site | Where-Object {$_.LoginName -eq "Readers"}))
	{
#       Do nothing a Readers Group already exists
	}
	else
	{
#       Create the Readers Group 		
		New-SPOSiteGroup -Group "Readers" -PermissionLevels "Read" -Site $Site
	}
	
#   Add manager to the Readers group	
	Add-SPOUser -Group "Readers" -LoginName $LoginName -Site $site
	
#   Remove individual running the script admin access
	Set-SPOUser -site $Site -LoginName $LoginName -IsSiteCollectionAdmin $False
