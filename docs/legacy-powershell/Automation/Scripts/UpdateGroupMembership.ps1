#  Update Group Membership
#################################################################################
# 
# PowerShell source code
# Revision v1.2
# ==========================================================================
#    'Project      : UL Office 365 Exchange Online
#    'Description  : Updates Distribution Group Membership
#    'Called By    : DistributionGroupMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : by Sandi Glazebrook
#    'Date Created : 10/11/2016 04:00:00 PM
#    'Modified     : 12/06/2017 - Added Menu Option 3 and the related code.  Moved the 
#                  :    membership adds into a Function called AddMembers
#                  : 02/14/2023 - Changed the .Name to .DisplayName due to changes by MS where name for new users is a GUID
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "UpdateGroupMembership"
	$LogDrive		= "e:\Automation"
	$LogFolder		= "\" + $FileName
	$LogDirectory	= $LogDrive + $LogFolder + "\"
	$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	$InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
	$ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
# Retrieve the user name
	$WhoAmI			= WhoAmI
		
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name
	
# =============================================================================================================================================

function CheckLogFiles
{
	# Create Files folder $LogDirectory if it's not present
	if (Test-Path $LogDirectory)
		{
		# the directory is present
		}
	else
		{ 
		mkdir $LogDirectory
		}
} #end CheckLogFiles

function WriteLogEvent
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LocalMachineName + "`t" + $LineToWrite
	Out-File -filepath $LogFile -append -noClobber -inputObject $RecordEvent
} #end WriteLogEvent

function WriteReportEvent
{
	$RecordEvent = (get-date -uformat %D) + "`t" + (get-date -uformat %T) + "`t" + $LineToWrite
	Out-File -filepath $ReportFile -append -noClobber -inputObject $RecordEvent
} #end WriteReportEvent

function AddMembers
{
    ForEach ($Grp in $Grp)
    {
        If ($TestGroup = Get-DistributionGroup $Grp.GrpName)
        {
            $CurrentMembers = (Get-DistributionGroupMember $Grp.GrpName -ResultSize Unlimited | Sort-Object {$_.DisplayName})
        	write-host "Adding Members to Distribution Group: " $Grp.DgDisplayName
	        $LineToWrite = $Grp.DgDisplayName + " - Adding Members to Distribution Group."
		    WriteReportEvent
					
    		$addMember  = $Grp.NewMembers.split(",")
	    	if ($? -eq $true)
            {
                ForEach ($mail in $addMember)
                {
	    		    if ($mail.contains("@"))
                    {
			    	    if ($NewMember = Get-Recipient $mail)
                        {
					        if ($CurrentMembers -match $NewMember.DisplayName)
                            {
                                write-host $mail " already a member."
    		    				$LineToWrite = "`t" + "Warn" + "`t" + $mail + " - is already a member of Distribution Group. " + $Grp.GrpName
	    		    			WriteReportEvent
		    		        }
			    		    else
                            {
					    	    Add-DistributionGroupMember $Grp.GrpName -Member $mail
						    	write-host $mail " was added."
    						    $LineToWrite = "`t" + "Success" + "`t" + $mail + " - added to Distribution Group. " + $Grp.GrpName
    	    					WriteReportEvent
	    	    			}
		    	   		}
		        		else
                        {
                            write-host $mail " - ERROR finding recipient."
						    $LineToWrite = "`t" + "ERROR" + "`t" + $mail + " - ERROR finding recipient. " + $Grp.GrpName
        					WriteReportEvent
	        			}
                    }
		   	    	else
                    {
                        write-host "Nothing to add"
          			}
        		}
	   	    }
        }
        else
        {
	        Write-Host $Grp.GrpName " - ERROR finding Distribution Group."
		    $LineToWrite = "ERROR" + "`t" + $Grp.GrpName + " - ERROR finding Distribution Group."
        	WriteReportEvent
	    }
    }
    Write-host "Post Group membership count:  " (Get-DistributionGroupMember -ResultSize Unlimited $grp.GrpName).Count
    $LineToWrite = "INFO" + "`t" + $Grp.GrpName + "`t" + "Group Membership Updated" + "`n"
    WriteReportEvent
    write-host ""
}

# Setup Folders and Files	
	CheckLogFiles
	$LineToWrite = "STAR" + "`t" + $FileName + " script has started"
	WriteLogEvent
	$LineToWrite = "INFO" + "`t" + "Launched by: " + $WhoAmI + "`n"
	WriteLogEvent
		
#	
# =============================================================================================================================================
#
# Connect to Office 365
#
invoke-expression -Command .\ConnectO365.ps1

# =============================================================================================================================================

write-host "     Enter ( 1) Update Membership Using an Input File"
write-host "           ( 2) Move Membership from one Group to Another Group"
write-host "           ( 3) Add to Group Membership Using an Input File"
write-host ""
write-Host "           ( 0) to Return to the O365 Admin Menu"
write-Host "     Enter Option? " -ForegroundColor Red -NoNewline
$MemChg = Read-Host
write-host ""

switch ($MemChg)
{

    1
    {
        write-host "This feature will replace the membership of an existing distribution list with the membership"
        write-host "provided in the Input file that was created.  You can enter details for multiple groups into the"
        write-host "input file"
        write-host ""
        pause

        if (Test-Path $InputFile)
        {					
	        $Grp = Import-CSV $InputFile
  	        ForEach ($Grp in $Grp)
            {

                write-host "Writing existing membership of the" $Grp.GrpName "to the location:" $LogDirectory
    	        $LineToWrite = $Grp.DgDisplayName + " - Current membership"
	    	    WriteReportEvent
#                $OutFileName = $LogDirectory + ($Grp.GrpName -replace " ","") + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".txt"
                (Get-DistributionGroupMember $Grp.GrpName -Resultsize Unlimited |Sort-Object {$_.DisplayName}) > $ReportFile
                write-host "Current Group membership count:  " (Get-DistributionGroupMember -Resultsize Unlimited $grp.GrpName).Count
    	        $LineToWrite = $Grp.DgDisplayName + " - membership count" + (Get-DistributionGroupMember -Resultsize Unlimited $grp.GrpName).Count
	    	    WriteReportEvent
				write-host "Updating group with new membership list"
    	        $LineToWrite = $Grp.DgDisplayName + " - Removing existing group membership"
	    	    WriteReportEvent
                Update-DistributionGroupMember -identity $Grp.GrpName -members $null -Confirm:$False -BypassSecurityGroupManagerCheck
			    write-host "Old Membership Removed Starting Adding Mew Members...." -ForegroundColor Cyan
    	        $LineToWrite = $Grp.DgDisplayName + " - Adding New group members"
	    	    WriteReportEvent
                AddMembers
    	        $LineToWrite = $Grp.DgDisplayName + " - Replace Membership complete"
	    	    WriteReportEvent
	        }
            Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")
        }
        else
        {
            write-host "Input file not found."
        }
    }

    2
    {

        write-host "This feature will replace the membership of an existing distribution list with the membership"
        write-host "from another distribution list.  The membership of each group prior to being updated is written"
        write-host "to the \Logs directory.  No input file is needed as you will be prompted for the deatils to"
        write-host "complete this request."
        write-host ""
        pause

        $NewMembership = $null
        $first = "on"

        write-host "Enter the Name of the Distribution List Membership you want to move to Another Group: " -ForegroundColor Yellow -nonewline
        $GrpName = read-host

        $GrpMem = (Get-DistributionGroupMember $GrpName |Sort-Object {$_.DisplayName})

        if ($GrpMem.count -gt 0)
        {
            foreach ($GrpMem in $GrpMem)
            {
                $mbx = get-mailbox $GrpMem.DisplayName
                if ($first -eq "on")
                {
                    $NewMembership = $mbx.PrimarySmtpAddress
                    $first = "off"
                }
                else
                {
                    $NewMembership = $mbx.PrimarySmtpAddress + "," + $NewMembership
                }
            }

            write-host "Enter the Name of the Distribution List to apply this New Membership to: " -ForegroundColor Yellow -nonewline
            $UpdGrp = read-host

            write-host "Writing existing membership of the" $UpdGrp "to the location:" $LogDirectory

            $OutFileName = $LogDirectory + ($UpdGrp -replace " ","") + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".txt"
            (Get-DistributionGroupMember $UpdGrp | Sort-Object {$_.DisplayName}) > $OutFileName

            Update-DistributionGroupMember -identity $UpdGrp -members $NewMembership.split(",") -BypassSecurityGroupManagerCheck
            $LineToWrite = "INFO" + "`t" + $UpdGrp + "`t" + "Group Membership Replaced by" + $GrpName + "Membership" + "`n"
			WriteReportEvent
        }
        else
        {
            write-host "There are no members in the group" $GrpName
        }
    }

    3
    {
        write-host "This feature will add to the membership of an existing distribution list"
        write-host "provided in the Input file that was created.  You can enter details for multiple groups into the"
        write-host "input file"
        write-host ""
        pause

       if (Test-Path $InputFile)
        {					
	        $Grp = Import-CSV $InputFile
  	        AddMembers
            Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")
        }
        else
        {
            write-host "Input file not found."
        }
    }
}
$LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
WriteLogEvent