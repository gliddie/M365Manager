<#################################################################################
# 
# PowerShell source code
# Revision DRAFT v.01
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Grant Access to OneDrive
#    'Called By    : AccessRights.ps1, StandardTermination.ps1, EmergencyDisablement.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Larry Einhorn
#    'Date Created : 09/06/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :  12/04/2017 - S.Glazebrook - Created Callable Module
#    '  
# ==========================================================================
#
#################################################################################>
	write-Host "Enter Terminated Staff Employee Number: " -ForegroundColor Green -NoNewline
	$EmpID = Read-host
    $EmpID = $EmpID + "_global_ul_com"
	write-host "Enter the Manager Request Access Employee Number: " -ForegroundColor Green -NoNewline
	$ManagerID = Read-Host 
    $LoginName = ($ManagerID + "@global.ul.com")

    $Site = Get-SPOSite ("https://ul-my.sharepoint.com/personal/"+ $EmpId)
#	Set the individual running the script admin access
	Set-SPOUser -site $Site -LoginName $LoginName -IsSiteCollectionAdmin $True 

    $Site = Get-SPOSite ("https://ul-my.sharepoint.com/personal/"+ $EmpID)
    $OneDrAccess = [bool](Set-SPOUser -site $Site -LoginName $ManagerID -IsSiteCollectionAdmin $True)
    If ($OneDrAccess -eq $True)
    {
        write-host "Granting " $ManagerID "access to OneDrive" -ForegroundColor Yellow
#        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Removed " + $ENo + " Supervisor Access to OneDrive" + "`n"
#	    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
        Set-SPOUser -site $Site -LoginName $ManagerID -IsSiteCollectionAdmin $True 
    }
    else
    {
        write-host "This individual " $ManagerID "already has access to users OneDrive"
#        $LineToWrite = $RecordEvent + "INFO" + "`t" + "This individual " + $ENo + " already has access to users OneDrive" + "`n"
#	    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
    }

# End of Script