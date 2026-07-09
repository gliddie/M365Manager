<#################################################################################
# 
# PowerShell source code
# Revision v1.01
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online Migration
#    'Description  : Create New Shared Mailbox in Office 365
#    'Called By    : SharedMailboxAdminMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Mitch Skrove (HP)
#    'Date Created : 05/27/2011 10:00:00 AM
#    'Comments     :  
#    '               Set-ExecutionPolicy Unrestricted
#    '
#    'History      :         
#    '               06/03/2011 MDS Added Retention Policy and User Role Assignment 
#    '               11/30/2011 SAG Remove SMTP Forwarding and Added Quota Settings
#    '               06/11/2012 SAG Added -RetentionHoldEnabled $true
#    '               03/28/2013 SAG Commented out adding full permission for 
#    '                  ACL.UL.FullMailboxRight
#    '               04/22/2013 SAG Modified quota setting to 25GB per change in 
#    '                  O365 Service Description
#    '               10/17/2013 SAG Extracted Retention Policy Setting and created
#    '                  separate Script as Wave15 upgrade errors when setting within
#    '                  this script
#    '               05/04/2018 SAG Added code to set the MessageCopyForSendOnBehalf and 
#    '                  MessageCopyForSendOnBehalf to enbled.  This ensures messages sent
#    '                  from the mailbox are stored the the Sent folder of the shared mailbox
#    '               04/29/2019 SAG Modified the code so that mailboxes can be created with the @ul.org address
#    '               03/23/2020 SAG Changed the retention policy to an active 3 yr policy.
# ==========================================================================
#
#################################################################################>
 
$WhoAmI			= WhoAmI
$PrimaryMailDomain	= "@ul.com"
$RoutingDomain		= "@global.ul.com"

 
$LogDirectory		= "E:\Automation\NewSharedMailbox\Log"
$LogFile		= $LogDirectory + "\" + "Log-NewSharedMailbox.log"
 
$InputDirectory		= "E:\Automation\NewSharedMailbox\Input"
$InputFile		= $InputDirectory + "\" + "Input-NewSharedMailbox.csv"
 
$ReportDirectory	= "E:\Automation\NewSharedMailbox\Report"
$ReportFile		= $ReportDirectory + "\" + "Report-NewSharedMailbox"
 
 
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
		$LineToWrite = $RecordEvent + "STAR" + "`t" + "NewSharedMailbox script has started"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
		$LineToWrite = $RecordEvent + "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
		Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
                                
#	
# =============================================================================================================================================
#
# Connect to Office 365
#
    invoke-expression -Command .\ConnectO365.ps1
#
# =============================================================================================================================================
#
if (Test-Path $InputFile)
  {

	$SharedMailboxes = Import-CSV $InputFile

	ForEach ($Mailbox in $SharedMailboxes)
	  {
		If ($mailbox.Mail -like "*ul.org*")
        {
            $PrimaryMailDomain = "@ul.org"
        }
        else
        {
            $PrimaryMailDomain = "@ul.com"
        }
        $atMAIL   = $Mailbox.MAIL.indexOf("@")
		$LeftName = $Mailbox.MAIL.substring(0,$atMAIL)
		$addMail  = $Mailbox.LegacyMail.split(",")

		New-Mailbox -Name $Mailbox.Name -shared -Alias $LeftName -PrimarySmtpAddress ($LeftName + $PrimaryMailDomain)
		Set-CASMailbox $LeftName -ImapEnabled $false -PopEnabled $false
#		set-Mailbox $Mailbox.Name -IssueWarningQuota 24.5GB -ProhibitSendQuota 24.75GB -ProhibitSendReceiveQuota 25GB
		set-mailbox $Mailbox.Name -MessageCopyForSendOnBehalfEnabled $true -MessageCopyForSentAsEnabled $true
#		Set-Mailbox $Mailbox.Name –RetentionHoldEnabled $true –StartDateForRetentionHold 04/01/2011
		
		if ($Mailbox.Mail.contains($PrimaryMailDomain))
        {
		    Set-Mailbox $LeftName -EmailAddresses (((Get-Mailbox $Mailbox.Name).EmailAddresses)+=($LeftName + $RoutingDomain)) -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL MRM Policy - 3 yr Delete" -RetentionHoldEnabled $false -CustomAttribute15 "NewSharedMailbox PS Date: $Date PS Time: $Time"
       	}
		else
		{
		    Set-Mailbox $LeftName -EmailAddresses (((Get-Mailbox $Mailbox.Name).EmailAddresses)+=$Mailbox.MAIL,($LeftName + $RoutingDomain)) -RoleAssignmentPolicy "UL Default Role Assignment Policy" -RetentionPolicy "UL MRM Policy - 3 yr Delete" -RetentionHoldEnabled $false -CustomAttribute15 "NewSharedMailbox PS Date: $Date PS Time: $Time"
		}

		ForEach ($mail in $addMail)
		{
		    $add = $mail

		    if($add.contains("@"))
		    {
		        Set-Mailbox $LeftName -EmailAddresses (((Get-Mailbox $LeftName).EmailAddresses)+=$add)
		        #write-host $add " was added"
			
			$LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Name + "`t" + $add + "`t" + " proxyAddress added" + "`n"
			Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		    }
            else
		    {
		         #write-host "Nothing to add"
		    	$LineToWrite = $RecordEvent + "INFO" + "`t" + $Mailbox.Name + "`t" + " no legacyAddress to add" + "`n"
			    Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite
		    }
		}
	}

	# Rename the input file for future reference & Remove PS Session
    Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log") -ErrorAction SilentlyContinue
    if (Test-Path $InputFile)
    {
        $loop = 0
        Do
        {
            write-host "The file" $InputFile "is open by another process please close the file and hit returnt to continue" -ForegroundColor Red -NoNewline
            $Cont = read-host
            Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log") -ErrorAction SilentlyContinue
            $loop++
        } while ((Test-Path $InputFile) -and ($loop -lt 10))
        
        If ($loop -ge 10)
        {
            Write-Host "Unable to rename this file" $InputFule "please manually rename it to" ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
        }
    }
#		Remove-PSSession $Session

  }

else

  {
	write-host "Could not find input file :^) " $InputFile
	$LineToWrite = $RecordEvent + "STOP" + "`t" + "Could not find input file " + "`n"
	Out-File -filepath $LogFile -append -noClobber -inputObject $LineToWrite
	Remove-PSSession $Session
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