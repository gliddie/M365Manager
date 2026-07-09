<####################################################################################
#
#  This script contains all the common functions used by the O365 Team
#
#  Created: 5/13/2020 - S.Glazebrook
#  06/11/2020 - SAG - Modifed the Common Data Service and Bookings features in the E3EnabledFeatures
#  09/01/2020 - SAG - Modifed Meeting Room to Teams Rooms Standard as name changed at Microsoft
#  11/18/2020 - SAG - Moved the check/removal of the P2 license before the assignment of the E3 license
#  01/06/2021 - SAG - Added the Phone System License as a part of the standard license set
####################################################################################>

#Moved these to the O365AdminCommonFunctions.ps1

################

# Retrieve the user name
	$WhoAmI			= WhoAmI
		
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name

$wshell = New-Object -ComObject Wscript.Shell
$Global:OKDetails = ""

Invoke-Expression -Command e:\Scripts\Shared\O365AdminCommonFunctions.ps1