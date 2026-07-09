##  This script is to modify the enabled bits of the E4 licenses for Employees
##  This uses an input file to make changes the column titile is "UID" and each
##  line data is the emp#@global.ul.com

##  Current default license items to be enbaled for all users is:
##	   Sway, Yammer, Azure Rights Management, Sharepoint(OneDrive), Exchange Online(Plan2)

Start-transcript

$Chg = "N"
$InpFile = import-csv C:\temp\OneDrive.csv
$SSKID = "ul:ENTERPRISEWITHSCAL"

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
			$UserLicenseTest.ServiceStatus

			$HasE4Lic = $UserLicenseTest.Licenses | Where-Object {$_.AccountSkuId -eq $SSKID}

			If ($HasE4Lic.AccountSkuId -eq $SSKID)
			{
			
                write-host ""
                write-host ""
                $LineToWrite = ("Current License Assignment for " + $InpFile.UID + " - " + $UserLicenseTest.DisplayName)
#				Write-Host $LineToWrite -ForegroundColor Green
				write-output $LineToWrite

				$PlanStats = $HasE4Lic.ServiceStatus
				$PlanStats
                write-host ""

	# PlanStats[0]=PROJECTWORKMANAGEMENT, PlanStats[1]=Sway, PlanStats[2]=INTUNE_0365, PlanStats[3]=YAMMER_ENTERPRISE, PlanStats[4]=RMS_S_ENTERPRISE, PlanStats[5]=MCOVOICECONF,
	# PlanStats[6]=OFFICESUBSCRIPTION, PlanStats[7]=MCOSTANDARD, PlanStats[8]=SHAREPOINTWAC, PlanStats[9]=SHAREPOINTENTERPRISE, PlanStats[10]=EXCHANGE_S_ENTERPRISE
			    If (($PlanStats.ProvisioningStatus[0] -eq "Success") -and ($PlanStats.ProvisioningStatus[1] -eq "Success") -and ($PlanStats.ProvisioningStatus[3] -eq "Success") -and ($PlanStats.ProvisioningStatus[4] -and "Success") -and ($PlanStats.ProvisioningStatus[9] -eq "Success"))
			    {
				    $LineToWrite = "No changes made to license assignment User already has O365 Project Planner, Sway, Yammer, Rights Management & OneDrive Enabled"
#					Write-Host $LineToWrite
				    write-output $LineToWrite
			    }
			    else
			    {
				    $DisPlan = "MCOVOICECONF,MCOSTANDARD"
				    If (($PlanStats[6].ServicePlan.ServiceName -eq "OFFICESUBSCRIPTION") -and ($PlanStats[6].ProvisioningStatus -eq "Disabled"))
				    {
					    $DisPlan = $DisPlan + ",OFFICESUBSCRIPTION"
				    }
				    If (($PlanStats[8].ServicePlan.ServiceName -eq "SHAREPOINTWAC") -and ($PlanStats[8].ProvisioningStatus -eq "Disabled"))
				    {
					    $DisPlan = $DisPlan + ",SHAREPOINTWAC"
				    }
				    If (($PlanStats[10].ServicePlan.ServiceName -eq "EXCHANGE_S_ENTERPRISE") -and ($PlanStats[10].ProvisioningStatus -eq "Disabed"))
				    {
					    $DisPlan = $DisPlan + ",EXCHANGE_S_ENTERPRISE"
				    }

				    $DLO = ($DisPlan.Split(“,”))
				    $MyO365Sku = New-MsolLicenseOptions -AccountSkuId $SSKID -DisabledPlans $DLO
				    Set-MsolUserLicense -UserPrincipalName $InpFile.UID –LicenseOptions $MyO365Sku
				    $LineToWrite = "New License Assignment for " + $InpFile.UID
#					Write-host $LineToWrite -ForegroundColor Magenta
				    write-output $LineToWrite
				    $UserLicenseTest = Get-MsolUser -UserPrincipalName $InpFile.UID
				    $UserLicenseTest.ServiceStatus
					
				    $NewE4License = $UserLicenseTest.Licenses | Where-Object {$_.AccountSkuId -eq $SSKID}
				    $NewE4License.ServiceStatus
#					write-output ""
			    }
			}
			else
			{
				$LineToWrite = "This individual " + $InpFile.UID + " - " + $UserLicenseTest.DisplayName + " does not have an Enterprise E4 license"
#				Write-Host $LineToWrite -ForegroundColor Red
				write-output $LineToWrite
			}
        pause
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