#
#   Give Manager Access to OneDrive
#

    invoke-expression -Command .\ConnectO365SPO.ps1
#    $AdminURL = "https://ul-admin.sharepoint.com"
#    $AdminName = "96151@global.ul.com"
#    $Password = Read-host -assecurestring "Enter Password for $AdminName"
#    $Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Password
#    Connect-SPOService -url $AdminURL -credential $Credential

    $EmpID = Read-Host "Enter Terminated Staff Employee Number: " -ForegroundColor Green -NoNewline
    $EmpID = $EmpID + "_global_ul_com"
    $ManagerID = Read-Host "Enter the Manager Request Access Employee Number: " -ForegroundColor Green -NoNewline
    $LoginName = ($ManagerID + "@global.ul.com")

    $Site = Get-SPOSite ("https://ul-my.sharepoint.com/personal/"+ $EmpId)
    Set-SPOUser -site $Site -LoginName $LoginName -IsSiteCollectionAdmin $True 
