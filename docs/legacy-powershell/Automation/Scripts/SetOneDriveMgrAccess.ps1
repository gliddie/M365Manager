#connection
$AdminURL = "https://ul-admin.sharepoint.com"
$AdminName = "96151@global.ul.com"
  
#creds
#$Password = Read-host -assecurestring "Enter Password for $AdminName"
#$Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Password
#$SPOCred = new-object -typename System.Management.Automation.PSCredential

Connect-SPOService -url $AdminURL -credential $Global:LiveCred 
 
$EmpId = Read-Host "Enter employee ID"
$ManagerID = Read-Host "Enter the manager's ID"
$LoginName = ($ManagerID + "@global.ul.com")

$Site = Get-SPOSite ("https://ul-my.sharepoint.com/personal/"+ $empId +"_global_ul_com")


Set-SPOUser -site $Site -LoginName $LoginName -IsSiteCollectionAdmin $False
 
