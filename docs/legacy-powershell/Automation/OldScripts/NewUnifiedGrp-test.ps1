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
#    '     05/23/2018 Added Code to check $DGAlias length maximum is 64 Characters
#    '     06/13/2018 Added Code for new "All Applications" option in the request forms and to allow you to
#    '          the creation of a group
#    '     07/13/2018 Added code so that if $UFGExists if false when checking DGAlias it also checks the GrpName
#    '     08/31/2018 Added pause to make wait until Team is created before removing creator ownership/membership
#    '     11/12/2018 Fixing code where the creator is removed from team ownership/membership
#    '     01/19/2019 Added code so that when the SP Site has issues it allows you to exit that code
#    '     03/01/2019 Fixed code when group already exists and when an email was not previously sent to requestor
#    '     06/10/2019 Commented out the line to add a Team Owner as a member first
#    '     06/19/2019 Moved the Add requestor as team member to be after they are added as the owner
#    '     07/09/2019 Adding additional Check to see if the group name was entered as upper/lowercase.
#    '     08/07/2019 Made changes to New-Teams command as MS changed the New-Teams switches (-Alias changed to -MailNickname and -AccessType changed to -Visibility)
#    '     08/12/2019 Added code to prevent the question from appearing for sending the Recap email
#    '     09/26/2019 Added information to the email message that it may take several hours for the entire provisioning process to complete
#    '     02/21/2020 Added $image10 as the PowerBI Diamond was missing from the configuration
#    '     04/06/2020 Modified the email to state it may take up to 24 hrs for the provisioning process to complete
#    '     04/10/2020 Commented out adding the requestor as a team member looks like you can now add as an owner without them being a member first
#    '     04/17/2020 Modified to revamp email and fix issues with the images not being displayed.
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
    Do
	{
		$SendMsg = "N"
        If (($GrpRetry -eq "N") -and ($MsgSubject -notlike "*Nightly Log*"))
        {
            Write-Host "Has a message previously be sent regarding the creation of this group (Y/N)? " -ForegroundColor Cyan -NoNewline
            $SendMsg = Read-Host
        }
		If ($SendMsg -eq "N")
		{
			write-host "Sending email " $MsgSubject
			$message.To.Clear()
			$message.CC.Clear()
            $message.Bcc.Clear()
			$message.To.Add($SendTo)
	        $message.Bcc.Add("EnterpriseMessagingServices@ul.com")
			$message.Subject = $MsgSubject
			$message.Body = $MsgBody
            $SMTPClient = New-Object Net.Mail.SmtpClient($Server, 25)
            $SMTPClient.Send($message)
		}
    } while (($SendMsg -ne "N") -and ($SendMsg -ne "Y"))
}#end SendMessage

function CreateComplete
{
	$UfgExists = [bool](Get-UnifiedGroup $DGAlias -ResultSize Unlimited -ErrorAction SilentlyContinue)
	$First = "Y"
	Do
	{
		If ($First -eq "Y")
		{
			write-host "Waiting for Group Creation to Complete..." -ForegroundColor Cyan -NoNewline
			$First = "N"
		}
		else
		{
			write-host ".." -foregroundcolor Cyan -NoNewline
		}
		start-sleep -s 15
		$UfgExists = [bool](Get-UnifiedGroup $DGAlias -ResultSize Unlimited -ErrorAction SilentlyContinue)
	} while ($UfgExists -eq $False)
} #end waiting for group creation to complete

function SPSiteComplete
{
    $CycleCnt = 0
    $SiteName = "https://ul.sharepoint.com/sites/" + $DGAlias
#    Write-host "Waiting for SharePoint Site" $SiteName "to be provisioned.  Disregard Errors the SilentlyContinue ErrorAction does not work on this command" -ForegroundColor Red -NoNewline
    get-date |fl DateTime
    Write-host "Waiting for SharePoint Site" $SiteName "to be provisioned." -ForegroundColor Red -NoNewline
    start-sleep -s 15
    $ErrorActionPreference = "SilentlyContinue"

    Do
    {
        write-host ".." -ForegroundColor Red -NoNewline
        $SiteCreated = Get-SPOSite $SiteName
	    start-sleep -s 5
        $CycleCnt++
        If (($CycleCnt -eq 10) -and ($SiteCreated.Status -ne "Active"))
        {
            write-host "`nSite Creation Has not completed would you like to continue try again (Y/N)?" -ForegroundColor Cyan -NoNewline
            $SiteRetry = read-host
            If ($SiteRetry -eq "Y")
            {
                write-host "Continuing to wait for" $SiteName "creation to complete." -ForegroundColor Red -NoNewline
                $CycleCnt = 0
            }
        }
    } while (($SiteCreated.Status -ne "Active") -and ($CycleCnt -lt 10))

    write-host ""
    get-date |fl DateTime
    $ErrorActionPreference = "Continue"
    If ($SiteCreated.Status -eq "Active")
    {
        write-host "`nSharepoint Site has been Provisioned and is in an Active State" -ForegroundColor Green
    }
    else
    {
        write-host "Since the Sharepoint Site has not successfully completed do not send a email to the requestor" -ForegroundColor Red
    }
}#end waiting for sharepoint site creation to complete

function SuccessDetails
{
    $Global:TeamPara1 = "<u>To access and manage Microsoft Teams:</b></u><p>To Access and manage your Team open the <font color=blue>'Microsoft Teams App'</font> :</p>"
    $Global:TeamPara2 = "<p>Once the Team is created you will find it in the left hand navigation pane. To manage the group click on the <font color=blue>'(…)'</font> to the right of the team name and select <font color=blue>'Manage Team'</font>.  As an owner this where you can add additional owners or members or manage other attributes related to the Microsoft Team.<br></p>"
    $Global:AppPara1 = "<p>You have requested that this group be enabled for All Applications this includes Microsoft Teams, PowerBI, Microsoft Planner, PowerApps, Flow, Sway and Stream.  To access applications other than Microsoft Teams open a web browser and go to <a href='http://webmail.ul.com'>webmail.ul.com </a> this will open your email file.  Click on the <font color=blue>'App Launcher'</font> in the upper left hand corner and select <font color=blue>'All Apps'</font> to find the application you would like to use.</p>"
    $Global:ClosePara = "<p>The subject ticket is now complete and will be closed.  If you have additional questions or need additional support please contact the Service Desk.</p><p>Thanks,<br>The Enterprise Messaging Team</p>"
}

$ActionLog = "<p>O365 Unified Groups Automated Processing Details</p>"
$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server
$from = New-Object System.Net.Mail.MailAddress "DoNotReply@ul.com" , "Enterprise Messaging Services"
$to = $from 
$msgfont = "<basefont face=verdana size=2.5 color=black>"
$Tab = "<p style=""margin-left: 40px"">"

#Images
    $image1 = "e:\O365AdminShared\Data\TeamApp.png"
    $att1 = new-object Net.Mail.Attachment($Image1)
    $att1.ContentType.MediaType = “image/png”
    $att1.ContentId = “Attachment1”
    $Img1 = $Tab + "<img src='cid:$($att1.ContentId)'"
    $image2 = "e:\O365AdminShared\Data\ManageTeam.png"
    $att2 = new-object Net.Mail.Attachment($Image2)
    $att2.ContentType.MediaType = “image/png”
    $att2.ContentId = “Attachment2”
    $Img2 = $Tab + "<img src='cid:$($att2.ContentId)'"
    $image3 = "e:\O365AdminShared\Data\AppLauncher.png"
    $att3 = new-object Net.Mail.Attachment($Image3)
    $att3.ContentType.MediaType = “image/png”
    $att3.ContentId = “Attachment3”
    $Img3 = $Tab + "<img src='cid:$($att3.ContentId)'"

SuccessDetails

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
 #       invoke-expression -Command .\ConnectO365.ps1
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
            $message = new-object  System.Net.Mail.MailMessage $from, $to 
            $message.IsBodyHtml = $true
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

                if ($DG.u_group_name -cmatch "[A-Z]")
                {
                    $ProposedName = (($DG.u_group_name).Trimstart(" ."))
                }
                else
                {
                    $ProposedName = (($DG.u_group_name.ToLower()).Trimstart(" ."))
                }

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
#                $ProposedName = $TextInfo.ToTitleCase($ProposedName)
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
                write-host "              Group Location: " $DG.u_group_location -ForegroundColor Yellow
                write-host "    3rd element as Submitted: " $DG.u_group_name -ForegroundColor Yellow
                write-host "3rd element of the GroupName: " $ProposedName -ForegroundColor Yellow
                write-host "         Name of Group Owner: " $DG.u_group_owner -ForegroundColor Yellow
                write-host "            Name of Requstor: " $DG.requested_for -ForegroundColor Yellow
                write-host "             Use Description: " $DG.u_brief_description -ForegroundColor Yellow
                write-host "      Service Desk Ticket No: " $DG.u_requested_item -ForegroundColor Yellow
                write-host ""
                $GrpName = "GRP." + $DGGrpNameSplit[1] + "." + $ProposedName.TrimEnd()
                write-host "          Create group using: " $GrpName -ForegroundColor Red
                write-host "`nDo any additional changes need to be made to this name (Yes/No/Skip)? " -ForegroundColor Cyan -NoNewline
                $NameOK = Read-Host

                If ($NameOK -notlike "S*")
                {
                    If ($NameOK -notlike "N*")
                    {
                        Do
                        {
                            If ($NameOK -notlike "N*")
                            {
                                write-host "Enter the 3rd element of the GroupName: " -ForegroundColor green -NoNewline
                                $ProposedName = Read-Host
                                $GrpName = "GRP." + $DGGrpNameSplit[1] + "." + $ProposedName.TrimEnd()
                                write-host "          Creating group as: " $GrpName -ForegroundColor Red
                                write-host "`Do you need to make additional changes (Yes/No)? " -ForegroundColor Cyan -NoNewline
                                $NameOK = Read-Host
                            }
                        } while ($NameOK -notlike "N*")
                    } 

                    $collection.reset()
                    $GrpName = "GRP." + $DGGrpNameSplit[1] + "." + $ProposedName.TrimEnd()
                    $DGAlias = $GrpName -Replace '[ /,$#_-]',''
                    $DGAlias = $DGAlias.Replace("\","")
                    $DGAlias = $DGAlias.Replace(".","")
                    If ($DGAlias.Length -gt 64)
                    {
                        Write-Host "Alias is too long maximum length is 64 Characters" -ForegroundColor Red
                        $DGAlias = $DGAlias.Substring(0,64)
                    }
                    $INetAddress = $DGAlias.Replace(".","") + "@ul.onmicrosoft.com"
                    $OwnerMbxExists = [bool](get-mailbox $DGOwnerID)

                    $UfgExists = [bool](Get-UnifiedGroup $DGAlias -ResultSize Unlimited -ErrorAction SilentlyContinue)

                    If (($UfgExists -ne "True") -and ($UfgExists -ne "False"))
                    {
                        $UfgExists = [bool](Get-UnifiedGroup $GrpName -ResultSize Unlimited -ErrorAction SilentlyContinue)
                        If ($UfgExists -eq $True)
                        {
                            $CurGroup = Get-UnifiedGroup $GrpName
                        }
                    }
                    else
                    {
                        $CurGroup = Get-UnifiedGroup $DGAlias
                    }

                    $GrpRetry = "N"
                    If ($UfgExists -eq $True)
                    {
#                        $CurGroup = Get-UnifiedGroup $GrpName
                        write-host "Was this group previously created without an email being sent (Y/N)? " -ForegroundColor Cyan -NoNewline
                        $GrpRetry = Read-Host
                        If ($GrpRetry -eq "N")
                        {
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
                    }
                    else
                    {
				        if ($DG.u_app_used_with -eq "All Applications")
				        {
            #            Set Values for message that Group has been Created for use with All Applications
					        $MsgSubject = $GrpName + " Has Been Created for Use with All Applications - Per:  " + $DG.u_requested_item
                            $OpenPara = "<p>Per your request the O365 group named <font color=green>" + $GrpName + "</font> has been created per service desk request number " + $DG.u_requested_item + ".  You have been made the owner of the group.  Due to the increase in Teams usage it may take up to 24 hours for the entire provisioning process to complete before you will be able to fully manage the group.  It is your responsibility to manage access to the team by adding and removing members from the group.  There is currently an issue with the automated email notifications when you add members to your group.  Therefore, you will need to notify any individuals you add to the group until this issue is resolved.</p><p>The following training videos are available:  <a href='https://support.office.com/en-us/article/microsoft-teams-video-training-4f108e54-240b-4351-8084-b1089f0d21d7?wt.mc_id=otc_home&ui=en-US&rs=en-US&ad=US'>TeamsTraining</a><p>"
                            $message.Attachments.Add($att1)
                            $message.Attachments.Add($att2)
                            $message.Attachments.Add($att3)
                            $MsgBody = $msgfont + $OpenPara + $Global:TeamPara1 + $Img1 + $Global:TeamPara2 + $Img2 + $Global:AppPara1 + $Img3 + $Global:ClosePara
				        }
				        elseif ($DG.u_app_used_with -eq "Microsoft Teams")
				        {
            #            Set Values for message that Group has been Created for use with Microsoft Teams
					        $MsgSubject = $GrpName + " Has Been Created for Use with Microsoft Teams - Per:  " + $DG.u_requested_item
                            $OpenPara = "<p>Per your request the O365 group named <font color=green>" + $GrpName + "</font> has been created per service desk request number " + $DG.u_requested_item + ".  You have been made the owner of the group.  Due to the increase in Teams usage it may take up to 24 hours for the entire provisioning process to complete before you will be able to fully manage the group.  It is your responsibility to manage access to the team by adding and removing members from the group.  There is currently an issue with the automated email notifications when you add members to your group.  Therefore, you will need to notify any individuals you add to the group until this issue is resolved.</p><p>The following training videos are available:  <a href='https://support.office.com/en-us/article/microsoft-teams-video-training-4f108e54-240b-4351-8084-b1089f0d21d7?wt.mc_id=otc_home&ui=en-US&rs=en-US&ad=US'>TeamsTraining</a><p>"
                            $message.Attachments.Add($att1)
                            $message.Attachments.Add($att2)
                            $MsgBody = $msgfont + $OpenPara + $Global:TeamPara1 + $Img1 + $Global:TeamPara2 + $Img2 + $Global:ClosePara
                        				        }
				        elseif ($DG.u_app_used_with -eq "PowerBI")
				        {
            #            Set Values for message that Group has been Created for use with PowerBI
					        $MsgSubject = $GrpName + " Has Been Created for use with PowerBI - Per: " + $DG.u_requested_item
					        $OpenPara = "<p>Per your request the O365 group named <font color=green>" + $GrpName + "</font> has been created per service desk request number " + $DG.u_requested_item + ".  It may take up to 24 hours for the entire provisioning process to complete before you will be able to fully manage the group.</p><p>You have been made an owner of the group and can add additional members or owners to the group using the <font color=green>Edit Workspace</font> feature in PowerBI.</p><p>If you specified that you are sharing with more than 10 individuals the group will be added to the PowerBI Premium Capacity.  Once the group has been added into the Premium Capacity you will see a <font color=green>Diamond</font> displayed after the group name.  Data published to a group that is a part of the Premium Capacity will <b>not</b> require the individuals who are viewing the information to obtain a PowerBI Pro license.</p><p>The following training videos are available:  <a href='https://support.office.com/office-training-center?redirectSourcePath=%252farticle%252fb8f02f81-ec85-4493-a39b-4c48e6bc4bfb'>O365TrainingCenter</a><p>"
					        $MsgBody = $msgfont + $OpenPara + $Global:ClosePara
				        }
            #            Set Values for message that Group has been Created for use with all other group types
                        elseif (($DG.u_app_used_with -ne "PowerBI") -and ($DG.u_app_used_with -ne "Microsoft Teams") -and ($DG.u_app_used_with -ne "All Applications"))
				        {
					        $MsgSubject = $GrpName + " Has Been Created for Use with " + $DG.u_app_used_with + " - Per:  " + $DG.u_requested_item
					        $OpenPara = "<p>Per your request the O365 group named <font color=green>" + $GrpName + "</font> has been created per service desk request number " + $DG.u_requested_item + ".  You have been made the owner of the group.  It may take up to 24 hours for the entire provisioning process to complete before you will be able to fully manage the group.  It is your responsibility to manage the membship of the group by adding and removing individuals from the group.</p><p>The following training videos are available:  <a href='https://support.office.com/office-training-center?redirectSourcePath=%252farticle%252fb8f02f81-ec85-4493-a39b-4c48e6bc4bfb'>O365TrainingCenter</a><p>"
					        $MsgBody = $msgfont + $OpenPara + $Global:ClosePara
                        }
                    }
       # Create the Unified Group an remove the owner as a member as by default the owner is added as a member

                    If (($DG.u_app_used_with -ne "Microsoft Teams") -and ($DG.u_app_used_with -ne "All Applications"))
                    {
                        If ($UfgExists -eq $False)
                        {
                            write-host "Creating New Unified Group" -ForegroundColor Green
			                New-UnifiedGroup -DisplayName $GrpName `
                                -PrimarySMTPAddress $INetAddress `
                                -Alias $DGAlias `
	                            -Owner $DGOwnerID `
                                -AccessType $DG.u_group_type `
                                | Out-Null

					        CreateComplete
                        }
					    Set-UnifiedGroup -identity $GrpName `
                            -UnifiedGroupWelcomeMessageEnabled:$false `
      		                -Notes ($DG.u_brief_description + "`nRequested by: " + $DG.u_group_owner + "`nPer: " + $DG.u_requested_item)
                        if (($DG.u_app_used_with -eq "PowerBI") -and ($DG.u_how_many -gt 10))
                        {
                            Add-UnifiedGroupLinks -Identity $GrpName -LinkType Members -links $WhoAmI -Confirm:$false
                            Add-UnifiedGroupLinks -Identity $GrpName -LinkType Owners -links $WhoAmI -Confirm:$false
                        }
                    }
                    else
                    {
                        If ($UfgExists -eq $False)
                        {
                            $TeamDesc = $DG.u_brief_description + "`nRequested by: " + $DG.u_group_owner + "`nPer: " + $DG.u_requested_item
                                                            
                            If ($GrpRetry -eq "N")
                            {
                                New-Team -DisplayName $GrpName -Description $TeamDesc -MailNickname $DGAlias -Visibility $DG.u_group_type | Out-Null
                                CreateComplete
                            }
                                                
                            $WelMsg = Get-UnifiedGroup $GrpName
                            If ($WelMsg.WelcomeMessageEnabled -eq "True")
                            {
    			                Set-UnifiedGroup $GrpName -UnifiedGroupWelcomeMessageEnabled:$false
                            }

                            If (($WhoAmI.Substring($WhoAmI.length-5)) -notlike $DG.u_owner_empid)
                            {
                                SPSiteComplete
                                $TeamGrpID = ((get-team |where-object {$_.DisplayName -eq $GrpName}).GroupID)
                                If ($GrpRetry -eq "N")
                                {
                                    write-host "Hit return when the Team has been created" -ForegroundColor Red -NoNewline
                                    $ret = read-host
                                }
                                Add-TeamUser -GroupId $TeamGrpID -User $DGOwnerID -Role Owner

                                If ([bool](((Get-TeamUser -GroupId $TeamGrpID -role owner).count) -ge 2))
                                {
                                    Remove-TeamUser -GroupId $TeamGrpID -User ($WhoAmI.Substring($whoAmI.indexof("\")+1)+"@global.ul.com") -Role Owner
                                    Remove-TeamUser -GroupId $TeamGrpID -User ($WhoAmI.Substring($whoAmI.indexof("\")+1)+"@global.ul.com")
                                }
                                else
                                {
                                    write-host "There is only one owner of this group"
                                }
                            }
                        }
                        else
                        {
                            write-host "Group already Exists"
                        }
                    }

					$UfgExists = [bool](Get-UnifiedGroup $DGAlias -ErrorAction SilentlyContinue)
					
                    if ($UfgExists -eq "True")
                    {
                        If ($GrpRetry -eq "Y")
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
                    write-host "Skipping the Creation of Group: " $DG.u_proposed_name -ForegroundColor Red
                }

            }
            else
            {
                $MsgSubject = $Dg.u_proposed_name + " Cannot be created Owner Not a Licensed Email User - Per:  " + $DG.u_requested_item
                $MsgBody = $msgfont + "<p>Per your request the O365 group named <font color=green>" + $Dg.u_proposed_name + "</font> was not created because " + $DG.u_group_owner + " is not a Licensed email User.</p><p>Thanks,<br>The Enterprise Messaging Team</p>"
            }

			If ($NameOK -notlike "S*")
            {
                If (($DG.u_app_used_with -eq "PowerBI") -and ($UfgExists -ne "True"))
                {
                    $ActionLog = $ActionLog + $GrpName + " has Been Created for use with PowerBI to be shared with " + $DG.u_how_many + " individuals - Per: " + $DG.u_requested_item + "<br>"
                }
                else
                {
                    $ActionLog = $ActionLog + $MsgSubject + "<br>"
                }
			    write-host "Sending Group Created Message to Requestor with Subject of: " $MsgSubject
                SendMessage
            }
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
 write-host "Sending Recap Message to Enterprise Messaging Services"
 SendMessage

# Rename the input file for future reference & Remove PS Session
    Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log") -ErrorAction SilentlyContinue
    if (Test-Path $InputFile)
    {
        $loop = 0
        Do
        {
            write-host "The file" $InputFile "is open by another process please close the file and hit returnt to continue" -ForegroundColor Red -NoNewline
            $Cont = read-host
            Rename-Item $InputFile ($InputFile + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log") -ErrorAction SilentlyContinue
            $loop++
        } while ((Test-Path $InputFile) -and ($loop -lt 10))
        
        If ($loop -ge 10)
        {
            Write-Host "Unable to rename this file" $InputFule "please manually rename it to" ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
        }
    }

# End of script #
#	Remove-PSSession $Session
#	$Session = $null
  $LineToWrite = "STOP" + "`t" + "This instance is stopping." + "`n"
 WriteLogEvent
 # =============================================================================================================================================