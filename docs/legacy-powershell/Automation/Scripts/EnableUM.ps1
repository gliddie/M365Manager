<#
#################################################################################
# 
# PowerShell source code
# Revision v1.05
# ==========================================================================
#    'Description  : Enable Unified Messaging or existing AD User account
#    'Called By    :  EUMMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Sandi Glazebrook
#    'Date Created : 05/06/2014
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '               05/06/2014 - SAG  - Created script from EnableMailUser2.ps1
                                         Removed (commented out) checks for employee type
                     04/07/2015 - Josh - Changed line that checks if extension is already in use to 
                                         filter on Cloud side to help speed up the script.  
                                         I commented out the old line if the new line doesn’t work for whatever reason. 
                     23/06/2015 - Paul - Added in an option to apply the relevant dial plan
                     23/07/2015 - Paul - Updated and redesigned the script to prompt for an action if an error is encountered,
                                         such as the extension already being in use or a different dial plan being assigned.

# ==========================================================================
#
#################################################################################
#>

$WhoAmI			= WhoAmI
$DC			= "usnbkd300p.global.ul.com"

$LogDirectory		= "E:\Automation\EnableUM\Log"
$LogFile		= $LogDirectory + "\" + "Log-EnableUM.log"

$InputDirectory		= "E:\Automation\EnableUM\Input"
$InputFile		= $InputDirectory + "\" + "Input-EnableUM.csv"

$ReportDirectory	= "E:\Automation\EnableUM\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-EnableUM"

$PrimaryMail		= "@ul.com"
$RoutingDomain		= "@global.ul.com"

invoke-expression -Command .\ConnectO365.ps1

	# Retrieve the local server name
		$Machine = get-wmiobject "Win32_ComputerSystem"
		$LocalMachineName = $Machine.Name


	# Retrieve the current Date and Time for use in log files
		$uDate = get-date -uformat %D
		$uTime = get-date -uformat %T
				
		$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"

	# Convert the current Date and Time to format MM-DD-YY and HHMMSS for use in file names	
		$Date  = $uDate.Replace("/", "-")
		$Time  = $uTime.Replace(":", "")
		
		$ReportFile = $ReportFile + "-" + "Date" + $Date + "Time" + $Time + ".Log"	
	
	# Create logging folder $LogDirectory if it's not present
		if (Test-Path $LogDirectory)
		{
			# the directory is present
		}
		else
		{
			mkdir $LogDirectory
		}
			
	# Add start record to log file
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "Enable MailUser script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	
# =============================================================================================================================================

if (Test-Path $InputFile)
{					
	$UserAccounts = Import-CSV $InputFile

	ForEach ($User in $UserAccounts)
	{
	   if ($User.EmpNo.contains("@"))
	    {
			$empID = $User.EmpNo.substring(0,($User.EmpNo.indexOf("@")))
			#$empName = $User.Address.substring(0,($User.Address.indexOf("@")))
            $dialplan = $User.DialPlan

            #check to make sure this is a valid mailbox
            $mbx = Get-Mailbox $user.empno -ErrorAction silentlyContinue
        }
            if ($mbx.Name -ne $null)
            {
                $empName = $mbx.PrimarySmtpAddress
                #check to see if Unified Messaging is Already Enabled for this user
                $umSet = get-UMMailbox $user.empno -ErrorAction SilentlyContinue
                
                #check to see if this Extension is already assigned to someone Else
                $filter = "EmailAddresses -Like 'eum:" + $user.Extension + "*'"
                $umAssigned = Get-UMMailbox -Filter $filter
                
                switch ($dialplan) 
                    { 
                        1 {
                            $dialplantext = "DefaultDialPlan Default Policy";
                            $functiontext = @{UMMailboxPolicy = $dialplantext
                            Extensions = $User.Extension
                            SIPResourceIdentifier = $empName
                            PinExpired = $true
                            }
                            break
                            }          
                        2 {
                            $dialplantext = "UL-NBK-Avaya";
                            $functiontext = @{UMMailboxPolicy = $dialplantext
                            Extensions = $User.Extension
                            PinExpired = $true
                            }
                            break
                            } 
                        default {
                            $dialplantext = "DefaultDialPlan Default Policy";
                            $functiontext = @{UMMailboxPolicy = $dialplantext
                            Extensions = $User.Extension
                            SIPResourceIdentifier = $empName
                            PinExpired = $true
                            }
                            break
                            } 
                        }
                If (($umSet.Name -eq $Null) -and ($umAssigned -eq $Null))
                {
                    Enable-UMMailBox $User.EmpNo @functiontext
                    
                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unified Messaging Enabled for: " + $empID + ", " + $empName + "`t" + "Ext: " + $User.Extension + " for Dial Plan: " + $dialplantext + "`n"
			        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

                    Write-Host "Unified Messaging Enabled for: " $empName "    Ext: " $user.extension  "    DialPlan: " $dialplantext -BackgroundColor Green -ForegroundColor Black
                }

                If ($umSet.Name -ne $Null) #Unified Messaging Already Enabled for User - is it for the same dial plan and number,?
					{
					    If ($dialplantext -eq $umSet.UMMailboxPolicy) #Same Dial Plan, No Action Needed
                            {
                                If ($umSet.PhoneNumber -eq $user.extension) #Check Extension
                                {
                                        $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unified Messaging Already Enabled for: " + $empID + ", " + $empName + ", " + $umSet.UMMailboxPolicy
                                        Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
                                        Write-Host "Unified Messaging Already Enabled for: " $empName "    EmpID: "$empID "    Dial Plan: " $umSet.UMMailboxPolicy -BackgroundColor Red
                                    }
                                Else
                            {
                                $title = "Unified Messaging Already Enabled for: " + $empName + " EmpID: " + $empID + " on Extension: " + $umSet.PhoneNumber
                                $message = "Do you want change the extension for " + $empName + " to " + $user.extension + "?"

                                $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
                                    "Changes the Extension."

                                $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
                                    "No Changes will be made."

                                $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

                                $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

                                switch ($result)
                                    {
                                        0 {
                                            Disable-UMMailbox $User.EmpNo -Confirm:$false;
                                            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unified Messaging Disabled for: " + $empID + ", " + $empName + "`t" + "Ext: " + $umSet.PhoneNumber + " for Dial Plan: " + $umSet.UMMailboxPolicy + "`n";
                                            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite;

                                            Enable-UMMailBox $User.EmpNo @functiontext
                                            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unified Messaging Enabled for: " + $empID + ", " + $empName + "`t" + "Ext: " + $User.Extension + " for Dial Plan: " + $dialplantext + "`n";
                                            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite;

                                            Write-Host "Unified Messaging Enabled for: " $empName "    Ext: " $user.extension  "    DialPlan: " $dialplantext -BackgroundColor Green -ForegroundColor Black
                                            }
                                        1 {
                                            "No Changes Made.";
                                            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unified Messaging Already Enabled for: " + $empID + ", " + $empName;
						                    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite;
						                    Write-Host "Unified Messaging Already Enabled for: " $empName "    EmpID:  "$empID -BackgroundColor Red
                                       }
                            }
                   
                    }  

                            }
                        Else
                            {
                                $title = "Unified Messaging Already Enabled for: " + $empName + " EmpID: " + $empID + " Dial Plan: " + $umSet.UMMailboxPolicy
                                $message = "Do you want change the dial plan for to " + $dialplantext + "?"

                                $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
                                    "Changes the Dial Plan."

                                $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
                                    "No Changes will be made."

                                $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

                                $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

                                switch ($result)
                                    {
                                        0 {
                                            Disable-UMMailbox $User.EmpNo -Confirm:$false;
                                            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unified Messaging Disabled for: " + $empID + ", " + $empName + "`t" + "Ext: " + $User.Extension + " for Dial Plan: " + $umSet.UMMailboxPolicy + "`n";
                                            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite;

                                            Enable-UMMailBox $User.EmpNo @functiontext
                                            Set-UMMailboxPIN $User.EmpNo -PinExpired $true
                                            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unified Messaging Enabled for: " + $empID + ", " + $empName + "`t" + "Ext: " + $User.Extension + " for Dial Plan: " + $dialplantext + "`n";
                                            Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite;

                                            Write-Host "Unified Messaging Enabled for: " $empName "    Ext: " $user.extension  "    DialPlan: " $dialplantext -BackgroundColor Green -ForegroundColor Black
                                            }
                                        1 {
                                            "No Changes Made.";
                                            $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unified Messaging Already Enabled for: " + $empID + ", " + $empName;
						                    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite;
						                    Write-Host "Unified Messaging Already Enabled for: " $empName "    EmpID:  "$empID -BackgroundColor Red
                                       }
                            }
                   
                    }  
                    }
                If (($umAssigned -ne $Null) -and ($mbx.PrimarySmtpAddress -ne $umAssigned.PrimarySmtpAddress)) #Unified Messaging Already Enabled for the Extension, but not the user
                    {

                        $title = "Extension " + $user.Extension + " Already Assigned to: " + $umAssigned.PrimarySMTPAddress + " unable to assign to " + $empName
                        $message = "Do you want move the extension to " + $empName + "?"

                        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
                            "Moves the extension from the existing user to the new user."

                        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
                            "No Changes will be made."

                        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

                        $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

                        switch ($result)
                            {
                                0 {
                                    Disable-UMMailbox $umAssigned.PrimarySmtpAddress -Confirm:$false
                                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unified Messaging Disabled for: " + $umAssigned.PrimarySmtpAddress + ", " + $umAssigned.Name + "`t" + "Ext: " + $User.Extension + "`n"
                                    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

                                    Enable-UMMailBox $User.EmpNo @functiontext
                                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Unified Messaging Enabled for: " + $empID + ", " + $empName + "`t" + "Ext: " + $User.Extension + " for Dial Plan: " + $dialplantext + "`n"
                                    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

                                    Write-Host "Unified Messaging Enabled for: " $empName "    Ext: " $user.extension  "    DialPlan: " $dialplantext -BackgroundColor Green -ForegroundColor Black
                                    }
                                1 {
                                    $LineToWrite = $RecordEvent + "INFO" + "`t" + "Extension " + $user.Extension + " Already Assigned to: " + $umAssigned.PrimarySMTPAddress + " unable to assign to " + $empName;
                	                Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite;
                                    Write-Host "Extension" $user.Extension "Already Assigned to: " $umAssigned.PrimarySMTPAddress "unable to assign to" $empName -BackgroundColor Red
                                }
                            }
                    } 

			}
    
    }
	# Rename the input file for future reference
		Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
}

else
{
	write-host "Could not find input file :^) " $InputFile
	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Could not find input file " + "`n"
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
}

# =============================================================================================================================================

	# Remove the Exchange PowerShell Snap-in
		Remove-PSSnapin *Exchange* -ErrorAction SilentlyContinue
		
	# Retrieve the current Date and Time for use in log files
		$uDate = get-date -uformat %D
		$uTime = get-date -uformat %T
		
		$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"

	# Write end record to the log file
		$LineToWrite = $RecordEvent + "STOP" + "`t" + "This instance is stopping." + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

# End of script #
