<#################################################################################
# 
# PowerShell source code
# Revision v1.00
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Create New Room List in Office 365
#    'Called By    : RoomResourceAdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 08/09/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '             : 
#    '               
#    '
# ==========================================================================
#
#################################################################################>
 
$WhoAmI			= WhoAmI
 
$LogDirectory		= "E:\Automation\NewRoomList\Log"
$LogFile		= $LogDirectory + "\" + "Log-NewRoomList.log"
 
$InputDirectory		= "E:\Automation\NewRoomList\Input"
$InputFile		= $InputDirectory + "\" + "Input-NewRoomList.csv"
 
$ReportDirectory	= "E:\Automation\NewRoomList\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-NewRoomList"
 
 
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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "NewRoomList script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                                
 
# =============================================================================================================================================

# Connect to Office 365
#	  $LiveCred = Get-Credential
#	  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
#	  Import-PSSession $Session

if (Test-Path $InputFile)
  {

	$RoomList = Import-CSV $InputFile

	ForEach ($Room in $RoomList)
	  {

New-DistributionGroup -Name $Room.Name -RoomList -ManagedBy $Room.ManagedBy

	  }

	# Rename the input file for future reference & Remove PS Session
		Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
#		Remove-PSSession $Session

  }

else

  {
	write-host "Could not find input file :^) " $InputFile
	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Could not find input file " + "`n"
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
#	Remove-PSSession $Session
  }



# =============================================================================================================================================


# Retrieve the current Date and Time for use in log files
	$uDate = get-date -uformat %D
	$uTime = get-date -uformat %T

	$RecordEvent = $uDate + "`t" + $uTime + "`t" + $LocalMachineName + "`t"

# Write end record to the log file
	$LineToWrite = $RecordEvent + "INFO" + "`t" + "This instance is stopping." + "`n"
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite

# End of script #
