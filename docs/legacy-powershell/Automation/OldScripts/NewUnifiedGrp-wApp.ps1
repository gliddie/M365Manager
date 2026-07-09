#################################################################################
# 
# PowerShell source code
# Revision v1.2
# ==========================================================================
#    'Project      : Deploy Microsoft Teams
#    'Description  : Creates New Unified Groups
#    'Called By    :
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Created from the NewDistributionGoup Script by Sandi Glazebrook
#    'Date Created : 6/26/2017 08:00:00 AM
#
#    'Comments     :  
#    '               6/28/2017 Image Files are written to: 
#    '                  http://sharepoint.ul.com/Mail/Forms/Thumbnails.aspx
#    '               7/17/2017 Fixed the variable name when populating the Notes
#	 '  				with the group owner information
#    '               8/2/2017 Modified the code to detect and remove the "GRP.XXX"
#    '                  identifiers if they were included in the 3rd element of the
#    '                  group name.  Also added verbiage for groups to be  used with
#    '                  PowerBI
#    '               8/15/2017 Adding code to modify email based on App the group will
#    '                  be used witn
# ==========================================================================
#
#################################################################################

# Declare Drive | Folders | and Files
	$FileName		= "NewUnifiedGroup"
	$LogDrive		= "e:\Automation"
	$LogFolder		= "\" + $FileNAme
	$LogDirectory	= $LogDrive + $LogFolder + "\"
	$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
	$InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
    $HistoryFile    = $LogDirectory + "History\History-" + $FileName  + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".csv"
	$ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
# =============================================================================================================================================

# Retrieve the user name
	$WhoAmI			= WhoAmI
		
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name

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

function SendMessage
{
    $message.To.Clear()
    $message.CC.Clear()
    $message.To.Add($SendTo)
#    $message.CC.Add("EnterpriseMessagingServices@ul.com")
    $message.Subject = $MsgSubject
    $message.Body = $MsgBody
    $client.Send($message)
}#end SendMessage

$ActionLog = "<p>O365 Unified Groups Automaed Processing Details</p>"
$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server
$from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"
$to = $from 
$message = new-object  System.Net.Mail.MailMessage $from, $to 
$message.IsBodyHtml = $true
$msgfont = "<basefont face=verdana size=2.5 color=black>"

$tab = "<p style=""margin-left: 40px"">"

$image1 = "http://sharepoint.ul.com/Mail/image018.jpg"
$image2 = "http://sharepoint.ul.com/Mail/image019.jpg"
#$image3 = "http://sharepoint.ul.com/Mail/image020a.jpg"
$image3 = "http://sharepoint.ul.com/Mail/image20b.jpg"
#$image4 = "http://sharepoint.ul.com/Mail/image021.jpg"
$image5 = "http://sharepoint.ul.com/Mail/image022.jpg"
$image6 = "http://sharepoint.ul.com/Mail/image023.jpg"
$image7 = "http://sharepoint.ul.com/Mail/image024.jpg"
$image8 = "http://sharepoint.ul.com/Mail/image025.jpg"
$image9 = "http://sharepoint.ul.com/Mail/image017.png"
$image10 = "http://sharepoint.ulcom/Mail/PowerBIDiamond.jpg"

#$secndline = "<u><b>Adding Microsoft Teams to O365 Group:</b></u><br>Open your web browser and go to <a href='http://webmail.ul.com'>webmail.ul.com </a> this will open your email file.  Click on the waffle <img src=$image1>&nbsp in the upper left hand corner to display other features that you have access to and select the <font color=blue>'Teams'</font> tile.<br>"
$secndline = "<u><b>To access and manage Microsoft Teams:</b></u><br>Open your web browser and go to <a href='http://webmail.ul.com'>webmail.ul.com </a> this will open your email file.  Click on the waffle <img src=$image1>&nbsp in the upper left hand corner to display other features that you have access to and select the <font color=blue>'Teams'</font> tile.<br>"
#$thirdline = "<p>This will bring you into the Microsoft Teams work space.  If you do not see the <font color=blue>'Create Team'</font> button you will need to select the <font color=blue>'Add Team'</font> which can be found at the bottom of the left hand navigation pane:</p>"
$thirdline = "<p>This will bring you into the Microsoft Teams work space.</p><p>Once the Teams Functionality is added to the group you will see this group in the left hand navigation pane.  To add individuals into the Microsoft Teams space use the <font color=blue>'Add more people'</font> feature.</p>"
#$forthline = "<p>This will change the icons in the right hand navigation pane and you will select <font color=blue>'Create Team'</font>:</p>"
#$fifthline = "<p>At the bottom of the right hand display window select the <font color=blue>'Yes add Microsoft Teams functionality'</font></p>"
#$sixthline = "<p>This will then show you a list of groups that you are the owner of select the group(s) you would like to add the Microsoft Teams functionality for and click on <font color=blue>'Choose Team'</font></p>"
#$svnthline = "<p>This will complete the creation of the Microsoft Teams workspace.  To add individuals into the Microsoft Teams space use the <font color=blue>'Add more people'</font> feature.</p>"
#$eigthline = "<p>Each person you add will receive a message informing them that they have been added in Microsoft Teams as shown below:</p>"
$ninthline = "<p>The subject ticket is now complete and will be closed.  If you have additional questions or need additional support please contact the Service Desk.</p><p>Thanks,<br>The Enterprise Messaging Team<Align=left></p>"

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
# =============================================================================================================================================
# Begin Unified Group creation
    if (Test-Path $InputFile)
    {	
    # test the date of the file.....if it is more than 24 hrs old it is an old file
        invoke-expression -Command .\ConnectO365.ps1
        $UnifDistGroup = Import-CSV $InputFile
        $Acro = Import-Csv e:\Automation\NewUnifiedGroup\Input\Input-KnownAcronyms.csv
        $collection = $Acro.GetEnumerator()

  	    ForEach ($DG in $UnifDistGroup)
        {
            $DGOwnerID = $DG.u_owner_empid + "@global.ul.com"
            $OwnerMbxExists = [bool](get-mailbox $DGOwnerID -ErrorAction SilentlyContinue)
            
            If ($OwnerMbxExists -eq "True")
            {
                $SendTo = $DGOwnerID		
                $DGGrpNameSplit = $DG.u_proposed_name.split(".")
                $SenderAuthentication = $true

                #
                #change the GroupName to use proper case and remove GRP and Location from the 3rd element of the name
                #
                $TextInfo = (Get-Culture).TextInfo
                $ProposedName = $DG.u_group_name.ToLower()
                $GrpName2 = $DGGrpNameSplit[1].ToLower() + "."
                $GrpName2a = $DGGrpNameSplit[1].ToLower() + " "
                $ProposedName = ($ProposedName.TrimStart("grp")).TrimStart(". ")
                $ProposedName = ($ProposedName.TrimStart($GrpName2))
                $ProposedName = ($ProposedName.TrimStart($GrpName2a))
                $ProposedName = $TextInfo.ToTitleCase($ProposedName)

                $GrpDescription = $DG.u_brief_description.ToLower()
                $GrpDescription = $TextInfo.TotitleCase($GrpDescription)

                #Check for Known acronyms
				write-host "3rd element of GroupName prior to running througj Known Acronyms: " $ProposedName
				
                foreach ($Acro in $Collection)
                {
                    $ProposedName = $ProposedName -Replace($Acro.Acronym,$Acro.Translation)
                }

                write-host "3rd element of the GroupName: " $ProposedName
                write-host "Do any additional changes need to be made to this name (Y/N)? " -ForegroundColor green -NoNewline
                $NameOK = Read-Host

                If ($NameOK -ne "N")
                {
                    write-host "Enter the 3rd element of the GroupName: " -ForegroundColor green -NoNewline
                    $ProposedName = Read-Host
                    $NameOK = "N"
                } 

                $collection.reset()
                $GrpName = "GRP." + $DGGrpNameSplit[1] + "." + $ProposedName
                $DGAlias = $GrpName -Replace '[ /,$#_-]',''
                $DGAlias = $DGAlias.Replace("\","")
                $INetAddress = $DGAlias.Replace(".","") + "@ul.onmicrosoft.com"
                $OwnerMbxExists = [bool](get-mailbox $DGOwnerID)

                $UfgExists = [bool](Get-UnifiedGroup $GrpName -ResultSize Unlimited -ErrorAction SilentlyContinue)

                If ($UfgExists -eq "True")
                {
                    $CurGroup = Get-UnifiedGroup $GrpName
                    $MsgSubject = $CurGroup.DisplayName + " Already Exists and Has Not Been Created - Per:  " + $DG.u_requested_item
                    $GrpOwner = $CurGroup.ManagedBy
                    $NoOwners = "owner"
                    if ($CurGroup.ManagedBy.count -gt 1)
                    {
                        $GrpOwner = $CurGroup.ManagedBy -join " and "
                        $NoOwners = "owners"
                    }
#   Set message vaules that group was not created because it already exists
                    $MsgBody = $msgfont + "<p>Per your request the O365 group named <font color=green>" + $CurGroup.DisplayName + "</font> was not created because a group with that name already exists.  The current group is owned and managed by <font color=green>" + $GrpOwner + "</font>.  If you have questions or would like access to this group please contact the " + $NoOwners + " directly.</p><p>If you would like to request a new group please create a new request using the <font color=blue>'Request O365 Group'</font> form in the Service Desk portal.</p><p>Thanks,<br>The Enterprise Messaging Team</p>"
                }
                else
                {
                    # Create the Unified Group an remove the owner as a member as by default the owner is added as a member
			        New-UnifiedGroup -DisplayName $GrpName `
                        -PrimarySMTPAddress $INetAddress `
                        -Alias $DGAlias `
	                    -Owner $DGOwnerID `
                        -AccessType $DG.u_group_type `
                        | Out-Null
                   Set-UnifiedGroup -identity $GrpName `
                        -UnifiedGroupWelcomeMessageEnabled:$false `
      		            -Notes ($DG.u_brief_description + "`nRequested by: " + $DG.u_group_owner + "`nPer: " + $DG.u_requested_item)
        #                Remove-UnifiedGroupLinks -Identity $GrpName -LinkType members -links $DGOwnerID
    
                    if (Get-UnifiedGroup $GrpName)
                    {
                        write-host " "
			            write-host "Unified Group created: " $GrpName  -ForegroundColor Green
			            $LineToWrite = "INFO" + "`t" + $GrpName + "`t UnifiedGroup Created" + "`n"
			            WriteReportEvent
		            }	
		            else
                    {
			            Write-Host "UnifiedGroup not created: " $GrpName
			            $LineToWrite = "FAIL" + "`t" + $GrpName + "`t" + "UnifiedGroup not created" + "`n"
			            WriteReportEvent
		            }
					
					if ($DG.u_app_used_with -eq "Microsoft Teams")
					{

        #            Set Values for message that Group has been Created
						$MsgSubject = $GrpName + " Has Been Created for Use with Microsoft Teams - Per:  " + $DG.u_requested_item
        #            $frstline = "<p>Per your request the O365 group named <font color=green>" + $GrpName + "</font> has been created per service desk request number " + $DG.u_requested_item + ". You have been made the owner of the group and it is your responsibility to manage the group membership.  There is currently an issue with the automated email notifiations when you add members to your group.  Therfore, you will need to notify any individuals you add to the group until this issue is resolved.</p><p>The Microsoft Teams functionality is being enabled for this group.  You will receive a separate email when these features are enabled.</p>"
        #            $frstline = "<p>Per your request the O365 group named <font color=green>" + $GrpName + "</font> has been created per service desk request number " + $DG.u_requested_item + ".  You have been made the owner of the group.  It is your responsibility to manage access to the team by adding and removing members from the group.  If you plan to use this group in the Microsoft Teams environment you will need to add the Microsoft Teams functionality to the group using the instructions provided below.</p>"
						$frstline = "<p>Per your request the O365 group named <font color=green>" + $GrpName + "</font> has been created per service desk request number " + $DG.u_requested_item + ".  You have been made the owner of the group.  It is your responsibility to manage access to the team by adding and removing members from the group.  There is currently an issue with the automated email notifiations when you add members to your group.  Therfore, you will need to notify any individuals you add to the group until this issue is resolved.</p><p>The Microsoft Teams functionality is being enabled for this group.  You will receive a separate email when these features are enabled.</p>"
						$MsgBody = $msgfont + $frstline + $secndline + $thirdline + "<p style=""margin-left: 40px""><img src=$image3></p>" + $ninthline
        #            $MsgBody = $msgfont + $frstline + $secndline + "<p style=""margin-left: 40px""><img src=$image2></p>" + $thirdline + $tab + "<img src=$image3>" + "  " + "<img src=$image4></p>" + $forthline + $tab + "<img src=$image5></p>" + $fifthline + $tab + "<img src=$image6></p>" + $sixthline + $tab + "<img src=$image7></p>" + $svnthline + $tab + "<img src=$image8></p>" + $eigthline + $tab + "<img src=$image9></p>" + $ninthline
					}
					if ($DG.u_app_used_with -eq "PowerBI")
					{

        #            Set Values for message that Group has been Created
						$MsgSubject = $GrpName + " Has Been Created for use with PowerBI - Per:  " + $DG.u_requested_item
						$frstline = "<p>Per your request the O365 group named <font color=green>" + $GrpName + "</font> has been created per service desk request number " + $DG.u_requested_item + ".  
						You have been made the owner of the group.  As an owner you can add additional members or owners to the group using the <font color=green>Edit Workspace Feature</font> in PowerBI.  If your group has been added to our Premium Capacity you will see a <img src=$image10>&nbsp after the group and any information that is published to the group will </bold>not</b> require the individuals who are viewing the inforamtion to obtain a PowerBI Pro license.</p><p>The PowerBI Premium Capacity is being enabled for this group and once complete the diamond will be displayed after the name of the group.</p>"
						$MsgBody = $msgfont + $frstline + $secndline + $thirdline + "<p style=""margin-left: 40px""><img src=$image3></p>" + $ninthline
        #            $MsgBody = $msgfont + $frstline + $secndline + "<p style=""margin-left: 40px""><img src=$image2></p>" + $thirdline + $tab + "<img src=$image3>" + "  " + "<img src=$image4></p>" + $forthline + $tab + "<img src=$image5></p>" + $fifthline + $tab + "<img src=$image6></p>" + $sixthline + $tab + "<img src=$image7></p>" + $svnthline + $tab + "<img src=$image8></p>" + $eigthline + $tab + "<img src=$image9></p>" + $ninthline
					}
					
                }
            }
            else
            {
                $MsgSubject = $Dg.u_proposed_name + " Cannot be created Owner Not a Licensed Email User - Per:  " + $DG.u_requested_item
                $MsgBody = $msgfont + "<p>Per your request the O365 group named <font color=green>" + $Dg.u_proposed_name + "</font> was not created because " + $DG.u_group_owner + " is not a Licensed email User.</p><p>Thanks,<br>The Enterprise Messaging Team</p>"
            }

            $ActionLog = $ActionLog + $MsgSubject + "<br>"
            SendMessage
        }
    }
    else
    {
        $ActionLog = $ActionLog + "Input file not found."
    }

 # =============================================================================================================================================

 #Send email to O365 Team with details of nightly actions
 $SendTo = "LST.O365AdminTeam@ul.com"
 #$SendTo = "Sandi.Glazebrook@ul.com"
 $MsgSubject = "Unified Group Nightly Log"
 $MsgBody = $msgfont + $ActionLog
 SendMessage

# Rename the input file for future reference & Remove PS Session
    Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")

# End of script #
#	Remove-PSSession $Session
#	$Session = $null
  $LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
 WriteLogEvent
 # =============================================================================================================================================