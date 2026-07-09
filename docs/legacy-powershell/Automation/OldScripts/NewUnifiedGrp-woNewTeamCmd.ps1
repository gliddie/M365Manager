#################################################################################
# 
# PowerShell source code
# Revision v1.2
# ==========================================================================
#    'Project      : Deploy Microsoft Teams
#    'Description  : Creates New Unified Groups
#    'Called By    : DistributionSecurityGroupMenu.ps1
#    'Calls        :
#    'Parameters   :
#    'Returns      :
#    'Author       : Created from the NewDistributionGoup Script by Sandi Glazebrook
#    'Date Created : 6/26/2017 08:00:00 AM
#
#    'Comments     :  
#    '     06/28/2017 Image Files are written to: http://sharepoint.ul.com/Mail/Forms/Thumbnails.aspx
#    '     07/17/2017 Fixed the variable name when populating the Notes with the group owner information
#    '     08/02/2017 Modified the code to detect and remove the "GRP.XXX" identifiers if they were 
#    '         included in the 3rd element of the group name.  Also added verbiage for groups to be
#    '         used with PowerBI
#    '     10/17/2017 Modified code so that PowerBI groups that already exit report correctly to the
#    '         $ActionLog email sent to the admins and modified the check so that the ReportFile logs
#    '         correctly for groups that exists or cannot be created
#    '     01/30/2018 Added code to trim white space from the end of the ProposedName variable
#    '     02/13/2018 Added line to provide details of how many new group requests there are to process
#    '     02/15/2018 Added code to use the New-Team commands
#    '     03/26/2018 Added "sleep" code after the creation and before any changes to address errors that
#    '         started occuring
#    '     04/25/2018 Added TrimEnd code on the description field.
#    '
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
    Write-Host "Has a message previously be sent regarding the creation of this group (Y/N)? " -ForegroundColor Cyan -NoNewline
    $SendMsg = Read-Host
    If ($SendMsg -eq "N")
    {
        $message.To.Clear()
        $message.CC.Clear()
        $message.To.Add($SendTo)
#        $message.CC.Add("EnterpriseMessagingServices@ul.com")
        $message.Subject = $MsgSubject
        $message.Body = $MsgBody
        $client.Send($message)
    }
}#end SendMessage

function CreateComplete
{
	$UfgExists = [bool](Get-UnifiedGroup $GrpName -ResultSize Unlimited -ErrorAction SilentlyContinue)
	$First = "Y"
    get-date -uformat %T

	Do
	{
		If ($First -eq "Y")
		{
			write-host "Waiting for Group" $GrpName "Creation to Complete..." -ForegroundColor Cyan -NoNewline
			$First = "N"
		}
		else
		{
			write-host ".." -foregroundcolor Cyan -NoNewline
		}
		start-sleep -s 30
		$UfgExists = [bool](Get-UnifiedGroup $GrpName -ResultSize Unlimited -ErrorAction SilentlyContinue)
	} while ($UfgExists -eq $False)
    write-host "`nUnified Group Creation Complete " -NoNewline
    get-date -uformat %T
} #end waiting for group creation to complete

function SPSiteComplete
{
    $SiteName = "https://ul.sharepoint.com/sites/" + $DGAlias
    Write-Host "Hit Return After Creating New Team from O365 Group in the Teams App" -ForegroundColor Red
    Read-Host
    Write-host "Waiting for SharePoint Site" $SiteName "to be provisioned.  Disregard Errors the SilentlyContinue ErrorAction does not work on this command" -ForegroundColor Red -NoNewline
    get-date -uformat %T
    $SiteName = "https://ul.sharepoint.com/sites/" + $DGAlias
	start-sleep -s 15
    Do
    {
		write-host ".." -ForegroundColor Red -NoNewline
	    $SiteCreated = Get-SPOSite $SiteName
	    start-sleep -s 45
    } while ($SiteCreated.Status -ne "Active")
    write-host "Sharepoint Site has been Provisioned and is in an Active State" -ForegroundColor Green
    get-date -uformat %T
}#end waiting for sharepoint site creation to complete

$ActionLog = "<p>O365 Unified Groups Automated Processing Details</p>"
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
$image3 = "http://sharepoint.ul.com/Mail/image020b.jpg"
#$image4 = "http://sharepoint.ul.com/Mail/image021.jpg"
$image5 = "http://sharepoint.ul.com/Mail/image022.jpg"
$image6 = "http://sharepoint.ul.com/Mail/image023.jpg"
$image7 = "http://sharepoint.ul.com/Mail/image024.jpg"
$image8 = "http://sharepoint.ul.com/Mail/image025.jpg"
$image9 = "http://sharepoint.ul.com/Mail/image017.png"

#$secndline = "<u><b>Adding Microsoft Teams to O365 Group:</b></u><br>Open your web browser and go to <a href='http://webmail.ul.com'>webmail.ul.com </a> this will open your email file.  Click on the waffle <img src=$image1>&nbsp in the upper left hand corner to display other features that you have access to and select the <font color=blue>'Teams'</font> tile.  This will open the Microsoft Teams workspace.<br>"
$secndline = "<u><b>To access and manage Microsoft Teams:</b></u><br>Open your web browser and go to <a href='http://webmail.ul.com'>webmail.ul.com </a> this will open your email file.  Click on the waffle <img src=$image1>&nbsp in the upper left hand corner to display other features that you have access to and select the <font color=blue>'Teams'</font> tile.<br>"
#$thirdline = "<p>This will bring you into the Microsoft Teams work space.  If you do not see the <font color=blue>'Create Team'</font> button you will need to select the <font color=blue>'Add Team'</font> which can be found at the bottom of the left hand navigation pane:</p>"
$thirdline = "<p>Once the Teams Functionality is added to the group you will see this group in the left hand navigation pane.  To add additional owners or members into the Microsoft Teams space use the <font color=blue>'Add more people'</font> feature.</p>"
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
		$ProcCount = $UnifDistGroup.Count
		If ($ProcCount -le 0)
		{
			$ProcCount = 1
		}
        write-host "Processing" $ProcCount "requests for New O365 Groups" -foregroundcolor Cyan
        $Acro = Import-Csv e:\Automation\NewUnifiedGroup\Input\Input-KnownAcronyms.csv
        $collection = $Acro.GetEnumerator()

  	    ForEach ($DG in $UnifDistGroup)
        {
			write-host "`nStarting Processing for group " $DG.u_proposed_name
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
                $ProposedName = (($DG.u_group_name.ToLower()).Trimstart(" ."))
                If ($ProposedName -match "GRP")
                {
                    $ProposedName = ($ProposedName.TrimStart("grp"))
                    $ProposedName = ($ProposedName.TrimStart(" ."))
                    $ProposedName = ($ProposedName.TrimEnd())
                }
                If ($ProposedName -match $DGGrpNameSplit[1])
                {
                    $ProposedName = ($ProposedName.TrimStart(($DGGrpNameSplit[1]).ToLower()))
                }
                $ProposedName = ($ProposedName.TrimStart(" ."))
                $ProposedName = $TextInfo.ToTitleCase($ProposedName)
                $GrpDescription = ($DG.u_brief_description.TrimEnd()).ToLower()
                $GrpDescription = $TextInfo.TotitleCase($GrpDescription)

                #Check for Known acronyms

                foreach ($Acro in $Collection)
                {
                    $ProposedName = $ProposedName -Replace($Acro.Acronym,$Acro.Translation)
                }

                write-host "             To be used with: " $DG.u_app_used_with -ForegroundColor Yellow
                if ($DG.u_app_used_with -eq "PowerBI")
                {
                    write-host "           To be shared with: " $DG.u_how_many "individuals" -ForegroundColor Yellow
                }
                write-host "3rd element of the GroupName: " $ProposedName -ForegroundColor Yellow
                write-host "`nDo any additional changes need to be made to this name (Y/N)? " -ForegroundColor Cyan -NoNewline
                $NameOK = Read-Host

                If ($NameOK -ne "N")
                {
                    write-host "Enter the 3rd element of the GroupName: " -ForegroundColor green -NoNewline
                    $ProposedName = Read-Host
                    $NameOK = "N"
                } 

                $collection.reset()
                $GrpName = "GRP." + $DGGrpNameSplit[1] + "." + $ProposedName.TrimEnd()
                $DGAlias = $GrpName -Replace '[ /,$#_-]',''
                $DGAlias = $DGAlias.Replace("\","")
                $DGAlias = $DGAlias.Replace(".","")
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
					if ($DG.u_app_used_with -eq "Microsoft Teams")
					{
        #            Set Values for message that Group has been Created for use with Microsoft Teams
						$MsgSubject = $GrpName + " Has Been Created for Use with Microsoft Teams - Per:  " + $DG.u_requested_item
#						$frstline = "<p>Per your request the O365 group named <font color=green>" + $GrpName + "</font> has been created per service desk request number " + $DG.u_requested_item + ".  You have been made the owner of the group.  It is your responsibility to manage access to the team by adding and removing members from the group.  There is currently an issue with the automated email notifications when you add members to your group.  Therefore, you will need to notify any individuals you add to the group until this issue is resolved.</p><p>The Microsoft Teams functionality is being enabled separately for this group, once complete you will see the group listed in the left navigator pane in the Teams workspace.</p>"
						$frstline = "<p>Per your request the O365 group named <font color=green>" + $GrpName + "</font> has been created per service desk request number " + $DG.u_requested_item + ".  You have been made the owner of the group.  It is your responsibility to manage access to the team by adding and removing members from the group.  There is currently an issue with the automated email notifications when you add members to your group.  Therefore, you will need to notify any individuals you add to the group until this issue is resolved.</p><p>The following training videos are available:  <a href='https://support.office.com/en-us/article/microsoft-teams-video-training-4f108e54-240b-4351-8084-b1089f0d21d7?wt.mc_id=otc_home&ui=en-US&rs=en-US&ad=US'>TeamsTraining</a><p>"
						$MsgBody = $msgfont + $frstline + $secndline + $thirdline + "<p style=""margin-left: 40px""><img src=$image3></p>" + $ninthline
					}
					if ($DG.u_app_used_with -eq "PowerBI")
					{
        #            Set Values for message that Group has been Created for use with PowerBI
						$MsgSubject = $GrpName + " Has Been Created for use with PowerBI - Per: " + $DG.u_requested_item
						$frstline = "<p>Per your request the O365 group named <font color=green>" + $GrpName + "</font> has been created per service desk request number " + $DG.u_requested_item + ".  You have been made the owner of the group.  As an owner you can add additional members or owners to the group using the <font color=green>Edit Workspace</font> feature in PowerBI.  The PowerBI Premium Capacity is being enabled separately for this group.  Once complete you will see a <img src=$image10>&nbsp displayed after the group name.  Data published to a group that is a part of the Premium Capacity will <b>not</b> require the individuals who are viewing the information to obtain a PowerBI Pro license.</p>"
						$MsgBody = $msgfont + $frstline + $ninthline
					}
        #            Set Values for message that Group has been Created for use with all other group types
                    if (($DG.u_app_used_with -ne "PowerBI") -and ($DG.u_app_used_with -ne "Microsoft Teams"))
					{
						$MsgSubject = $GrpName + " Has Been Created for Use with " + $DG.u_app_used_with + " - Per:  " + $DG.u_requested_item
						$frstline = "<p>Per your request the O365 group named <font color=green>" + $GrpName + "</font> has been created per service desk request number " + $DG.u_requested_item + ".  You have been made the owner of the group.  It is your responsibility to manage the membship of the gorup by adding and removing individuals from the group.</p>"
						$MsgBody = $msgfont + $frstline + $ninthline
                    }
        #            $MsgBody = $msgfont + $frstline + $secndline + "<p style=""margin-left: 40px""><img src=$image2></p>" + $thirdline + $tab + "<img src=$image3>" + "  " + "<img src=$image4></p>" + $forthline + $tab + "<img src=$image5></p>" + $fifthline + $tab + "<img src=$image6></p>" + $sixthline + $tab + "<img src=$image7></p>" + $svnthline + $tab + "<img src=$image8></p>" + $eigthline + $tab + "<img src=$image9></p>" + $ninthline

                    # Create the Unified Group an remove the owner as a member as by default the owner is added as a member

                    If ($DG.u_app_used_with -ne "Microsoft Teams")
                    {
                        write-host "Unified Group Creation Started at: " -ForegroundColor Green -NoNewline
			            New-UnifiedGroup -DisplayName $GrpName `
                            -PrimarySMTPAddress $INetAddress `
                            -Alias $DGAlias `
	                        -Owner $DGOwnerID `
                            -AccessType $DG.u_group_type `
                            | Out-Null

						CreateComplete

						Set-UnifiedGroup -identity $GrpName -UnifiedGroupWelcomeMessageEnabled:$false
      		            Set-UnifiedGroup -identity $GrpName -Notes ($DG.u_brief_description + "`nRequested by: " + $DG.u_group_owner + "`nPer: " + $DG.u_requested_item)
                        if (($DG.u_app_used_with -eq "PowerBI") -and ($DG.u_how_many -gt 10))
                        {
                            start-sleep -s 30
							Add-UnifiedGroupLinks -Identity $GrpName -LinkType Members -links $WhoAmI -Confirm:$false
                            Add-UnifiedGroupLinks -Identity $GrpName -LinkType Owners -links $WhoAmI -Confirm:$false
                        }
                    }
                    else
                    {
#                        New-Team -DisplayName $GrpName `
#                            -Description ($DG.u_brief_description + "`nRequested by: " + $DG.u_group_owner + "`nPer: " + $DG.u_requested_item) `
#                            -Alias $DGAlias `
#                            -AccessType $DG.u_group_type `
#                            | Out-Null

#                        CreateComplete
##################
			            New-UnifiedGroup -DisplayName $GrpName `
                            -PrimarySMTPAddress $INetAddress `
                            -Alias $DGAlias `
	                        -Owner $DGOwnerID `
                            -AccessType $DG.u_group_type `
                            | Out-Null

						CreateComplete

                        Start-Sleep -s 15
						Set-UnifiedGroup -identity $GrpName -UnifiedGroupWelcomeMessageEnabled:$false
						Set-UnifiedGroup -identity $GrpName -Notes ($DG.u_brief_description + "`nRequested by: " + $DG.u_group_owner + "`nPer: " + $DG.u_requested_item)
                        
                        Add-UnifiedGroupLinks -Identity $GrpName -LinkType Members -links $WhoAmI -Confirm:$false
                        Add-UnifiedGroupLinks -Identity $GrpName -LinkType Owners -links $WhoAmI -Confirm:$false
##################                        
                        $WelMsg = Get-UnifiedGroup $GrpName
                        If ($WelMsg.WelcomeMessageEnabled -eq "True")
                        {
    						Set-UnifiedGroup $GrpName -UnifiedGroupWelcomeMessageEnabled:$false
                        }

### Added the next 4 Lines as part of the Teams Space workaround
                        SPSiteComplete
                        write-host "`nCheck that SharePoint site is created once complete hit return" -ForegroundColor Red -NoNewline
                        Read-Host
                        get-date -uformat %T
                        Remove-UnifiedGroupLinks -Identity $GrpName -LinkType Owners -links $WhoAmI -Confirm:$false
                        Remove-UnifiedGroupLinks -Identity $GrpName -LinkType Members -links $WhoAmI -Confirm:$false


#### Commented Out below Section as part of the Teams Space workaround
#                        If (($WhoAmI.Substring($WhoAmI.length-5)) -notlike $DG.u_owner_empid)
#                        {
#                            write-host "Hit Return Once the Sharepoint Site is Created" -ForegroundColor Red
#                            read-host
#                            Add-UnifiedGroupLinks -Identity $GrpName -LinkType Members -links $DGOwnerID -Confirm:$false
#                            Add-UnifiedGroupLinks -Identity $GrpName -LinkType Owners -links $DGOwnerID -Confirm:$false

#                            If ([bool] ($NoOwners = get-UnifiedGroupLinks $GrpName -LinkType Owner | where {$NoOnwers.count -le 0}))
#                            {
#                                Do
#                                {
#                                    Add-UnifiedGroupLinks -Identity $GrpName -LinkType Members -links $DGOwnerID -Confirm:$false
#                                    Add-UnifiedGroupLinks -Identity $GrpName -LinkType Owners -links $DGOwnerID -Confirm:$false
#                                    $NoOwners = get-UnifiedGroupLinks $GrpName -LinkType Owner
#                                } while ($NoOwners.count -le 0)
#                            }
#  Temporarily removing the removal of (the below 2 lines may need to stay excluded depends on how membership issues play out)
#                            Remove-UnifiedGroupLinks -Identity $GrpName -LinkType Owners -links $WhoAmI -Confirm:$false
#                            Remove-UnifiedGroupLinks -Identity $GrpName -LinkType Members -links $WhoAmI -Confirm:$false
#                        }
                    }

					$UfgExists = [bool](Get-UnifiedGroup $GrpName -ErrorAction SilentlyContinue)
					
                    if ($UfgExists -eq "True")
                    {
                        write-host " "
                        if ($DG.u_app_used_with -eq "PowerBI")
                        {
			                write-host "Unified Group created: " $GrpName "for use with" $DG.u_app_used_with "to be shared with" $DG.u_how_many -ForegroundColor Green
                        }
                        else
                        {
                            write-host "Unified Group created: " $GrpName "for use with" $DG.u_app_used_with -ForegroundColor Green
                        }
			            $LineToWrite = "INFO" + "`t" + $GrpName + "`t UnifiedGroup Created" + "`n"
			            WriteReportEvent
		            }	
		            else
                    {
			            Write-Host "UnifiedGroup not created: " $GrpName
			            $LineToWrite = "FAIL" + "`t" + $GrpName + "`t" + "UnifiedGroup already exists or unable to create" + "`n"
			            WriteReportEvent
		            }
                }
            }
            else
            {
                $MsgSubject = $Dg.u_proposed_name + " Cannot be created Owner Not a Licensed Email User - Per:  " + $DG.u_requested_item
                $MsgBody = $msgfont + "<p>Per your request the O365 group named <font color=green>" + $Dg.u_proposed_name + "</font> was not created because " + $DG.u_group_owner + " is not a Licensed email User.</p><p>Thanks,<br>The Enterprise Messaging Team</p>"
            }

			If (($DG.u_app_used_with -eq "PowerBI") -and ($UfgExists -ne "True"))
            {
                $ActionLog = $ActionLog + $GrpName + " has Been Created for use with PowerBI to be shared with " + $DG.u_how_many + " individuals - Per: " + $DG.u_requested_item + "<br>"
            }
            else
            {
                $ActionLog = $ActionLog + $MsgSubject + "<br>"
            }
            write-host "Sending message to requestor"
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
 write-host "Sending Recap Message"
 SendMessage

# Rename the input file for future reference & Remove PS Session
    Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log")

# End of script #
#	Remove-PSSession $Session
#	$Session = $null
  $LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
 WriteLogEvent
 # =============================================================================================================================================

# Group created using Microsoft workaround -- cannot use New-Team must use New-UnifiedGroup