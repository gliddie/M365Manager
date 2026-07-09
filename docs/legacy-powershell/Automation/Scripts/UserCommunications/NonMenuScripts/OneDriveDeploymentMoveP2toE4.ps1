##  This script is to modify the enabled bits of the E4 licenses for Employees
##  This uses an input file to make changes the column titile is "UID" and each
##  line data is the emp#@global.ul.com

##  Current default license items to be enbaled for all users is:
##	   Sway, Yammer, Azure Rights Management, Sharepoint(OneDrive), Exchange Online(Plan2)

Start-transcript

$Chg = "N"
$InpFile = import-csv C:\temp\OneDriveEmpWP2.csv
$SSKID = "ul:EXCHANGEENTERPRISE"
$E4SSKID = "ul:ENTERPRISEWITHSCAL"

foreach ($InpFile in $InpFile)
{
    If ($InpFile.UID -notlike "*@*")
    {
        $InpFile.UID = $InpFile.UID + "@global.ul.com"
    }

    if ([bool](Get-MsolUser -UserPrincipalName $InpFile.UID -ErrorAction SilentlyContinue) -eq "True")
	{
		$UserLicenseTest = Get-MsolUser -UserPrincipalName $InpFile.UID
	}
	
	$MbxExists = [bool](Get-Mailbox $InpFile.UID -ErrorAction SilentlyContinue)
	
	If ($MbxExists -eq "True")
	{
			$EmpType = get-mailbox $InpFile.UID

		If ($EmpType.CustomAttribute1 -eq "Employee")
		{
			Write-host ""
            write-host "Moving " $InpFile.UID "from Exchange Online P2 License to an Enterprise E4 License" -ForegroundColor Yellow

			Set-MsolUserLicense -UserPrincipalName $InpFile.UID -RemoveLicenses "ul:EXCHANGEENTERPRISE"
					
			$DisPlan = "MCOVOICECONF,OFFICESUBSCRIPTION,MCOSTANDARD,SHAREPOINTWAC"
            $DLO = ($DisPlan.Split(“,”))
            $MyO365Sku = New-MsolLicenseOptions -AccountSkuId "ul:ENTERPRISEWITHSCAL" -DisabledPlans $DLO
            Set-MsolUserLicense -UserPrincipalName $InpFile.UID -AddLicense "ul:ENTERPRISEWITHSCAL" –LicenseOptions $MyO365Sku
			
			$UserLicenseTest = Get-MsolUser -UserPrincipalName $InpFile.UID
			$UserLicenseTest.ServiceStatus
			$NewE4License = $UserLicenseTest.Licenses | Where-Object {$_.AccountSkuId -eq "ul:ENTERPRISEWITHSCAL"}
			$NewE4License.ServiceStatus
		}
		else
		{
			$LineToWrite = "This individual " + $InpFile.UID + " - " + $UserLicenseTest.DisplayName + " is not an Employee and not entitled to use OneDrive"
#			Write-Host $LineToWrite -ForegroundColor Red
			write-output $LineToWrite
		}
	}
	else
	{
		$LineToWrite = "This individual " + $InpFile.UID + " - " + $UserLicenseTest.DisplayName + " does not have a mailbox, user may be terminated"
#		Write-Host $LineToWrite -ForegroundColor Red
		write-output $LineToWrite
	}
write-output ""
}
stop-transcript