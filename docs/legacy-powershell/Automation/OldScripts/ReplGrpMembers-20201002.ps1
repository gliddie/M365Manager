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
#    'Modified     : 12/06/2017 - Added Menu Option 3 and the related code.  Moved the membership adds into a Function called AddMembers
#                  : 10/02/2020 - Modified the add members to replace newline/carriage returns "`n" with a ","
# ==========================================================================
#
#################################################################################

# =============================================================================================================================================

function Build-UpdGrpMemberForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Replace Group Membership" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(650,250) #(W,H)
    $BldDetails = "N"

    ## Get Details used to create
    ## Display Name
    $Global:lblDispName = New-Object System.Windows.Forms.Label   
        $Global:lblDispName.Text = "Display Name:"
        $Global:lblDispName.Top = 10 ; $Global:lblDispName.Left = 10; $Global:lblDispName.Width=120 ;$Global:lblDispName.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblDispName)    # Add to Form 
        # 
        $Global:txtInpName = New-Object Windows.Forms.TextBox  
        $Global:txtInpName.TabIndex = 0 # set Tab Order 
        $Global:txtInpName.Top = 10; $Global:txtInpName.Left = 130; $Global:txtInpName.Width = 280;  
        $Global:txtInpName.Text = ""   # DisplayName
        $Global:txtInpName.TabIndex = 1
        $Global:form.Controls.Add($Global:txtInpName)    # Add to Form
 
    ## Ticket Number
    $Global:lblTaskNo = New-Object System.Windows.Forms.Label   
        $Global:lblTaskNo.Text = "Ticket Number:"  
        $Global:lblTaskNo.Top = 40 ; $Global:lblTaskNo.Left = 10; $Global:lblTaskNo.Width=120 ; $Global:lblTaskNo.AutoSize = $true
        $Global:form.Controls.Add($Global:lblTaskNo)    # Add to Form 
        # 
        $Global:txtInpTaskNo = New-Object Windows.Forms.TextBox  
        $Global:txtInpTaskNo.TabIndex = 0 # set Tab Order 
        $Global:txtInpTaskNo.Top = 40; $Global:txtInpTaskNo.Left = 130; $Global:txtInpTaskNo.Width = 120;  
        $Global:txtInpTaskNo.Text = "TASK"   # Enter ticket number
        $Global:txtInpTaskNo.TabIndex = 2
        $Global:form.Controls.Add($Global:txtInpTaskNo)    # Add to Form

    ## New Members
    $Global:lblGrpMbr = New-Object System.Windows.Forms.Label   
        $Global:lblGrpMbr.Text = "Group Member(s):"
        $Global:lblGrpMbr.Top = 70 ; $Global:lblGrpMbr.Left = 10; $Global:lblGrpMbr.Width=150; $Global:lblGrpMbr.AutoSize = $true 
        $Global:form.Controls.Add($Global:lblGrpMbr)    # Add to Form 
        # 
        $Global:txtGrpMbr = New-Object Windows.Forms.TextBox
        $Global:txtGrpMbr.MaxLength = 2000000
        $Global:txtGrpMbr.TabIndex = 3 # set Tab Order 
        $Global:txtGrpMbr.Location = New-Object System.Drawing.Size(130,70)
        $Global:txtGrpMbr.Size = New-Object system.Drawing.Size(500,90)
        $Global:txtGrpMbr.MultiLine = $true
        $Global:txtGrpMbr.ScrollBars = 'Both'  
        $Global:form.Controls.Add($Global:txtGrpMbr)    # Add to Form 
       # Obtain Value with: $Global:txtGrpMbr.Text

        Add-FormStandardButtons
}

function AddMembers
{
    If ($TestGroup = Get-DistributionGroup $Global:txtInpName.Text)
    {
        $CurrentMembers = Get-DistributionGroupMember $Global:txtInpName.Text -ResultSize Unlimited
        write-host "Adding Members to Distribution Group: " (Get-DistributionGroupMember $Global:txtInpName.Text).DisplayName
	    $LineToWrite = (Get-DistributionGroupMember $Global:txtInpName.Text).DisplayName + " - Adding Members to Distribution Group."
	    writeReportEvent
					
        $addMember  = ($Global:txtGrpMbr.Text -replace("`n",",")).split(",") -replace(" ","")
	    if ($? -eq $true)
        {
            ForEach ($mail in $addMember)
            {
#                if ($mail.contains("@"))
#                {
                if ($NewMember = Get-Recipient $mail)
                {
            	    if ($CurrentMembers -match $NewMember.Name)
                    {
                        write-host $mail "already a member."
    		  			$LineToWrite = "`t" + "Warn" + "`t" + $mail + "- is already a member of Distribution Group. " + $Global:txtInpName.Text
	    	    		WriteReportEvent
		    	    }
			   		else
                    {
				        Add-DistributionGroupMember $Global:txtInpName.Text -Member $mail
					   	write-host $mail "was added."
    					$LineToWrite = "`t" + "Success" + "`t" + $mail + " - added to Distribution Group. " + $Global:txtInpName.Text
    	    			writeReportEvent
	    	   		}
		       	}
		       	else
                {
                    write-host $mail "- ERROR finding recipient."
			        $LineToWrite = "`t" + "ERROR" + "`t" + $mail + " - ERROR finding recipient. " + $Global:txtInpName.Text
       				WriteReportEvent
        		}
            }
#		   	 else
#            {
#                write-host "Nothing to add"
#          	 }
#        }
	   	}
    }
    else
    {
	    Write-Host $Global:txtInpName.Text "- ERROR finding Distribution Group."
	    $LineToWrite = "ERROR" + "`t" + $Global:txtInpName.Text + " - ERROR finding Distribution Group."
        WriteReportEvent
	}

    Write-host "Post Group membership count:  " (Get-DistributionGroupMember -ResultSize Unlimited $Global:txtInpName.Text).Count
    $LineToWrite = "INFO" + "`t" + $Global:txtInpName.Text + "`t" + "Group Membership Updated" + "`n"
    WriteReportEvent
    write-host ""
}

# Declare Drive | Folders | and Files
	$FileName		= "UpdateGroupMembership"
	$LogDrive		= "e:\Automation"
	$LogFolder		= "\" + $FileName
	$LogDirectory	= $LogDrive + $LogFolder + "\"
	$LogFile		= $LogDirectory + "Log\Log-" + $FileName + ".log"
#	$InputFile		= $LogDirectory + "Input\Input-" + $FileName + ".csv"
	$ReportFile		= $LogDirectory + "Report\Report-" + $FileName + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".log"
	
# Retrieve the user name
	$WhoAmI			= WhoAmI
		
# Retrieve the local server name
	$Machine = get-wmiobject "Win32_ComputerSystem"
	$LocalMachineName = $Machine.Name
	
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

Build-UpdGrpMemberForm
Publish-Form

write-host "Writing existing membership of the" $Global:txtInpName.Text "to the location:" $LogDirectory
$OutFileName = $LogDirectory + ($Global:txtInpName.Text -replace " ","") + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".txt"
Get-DistributionGroupMember $Global:txtInpName.Text -Resultsize Unlimited > $OutFileName
write-host "Current Group membership count:  " (Get-DistributionGroupMember -Resultsize Unlimited $Global:txtInpName.Text).Count
write-host "Updating group with new membership list"
Update-DistributionGroupMember -identity $Global:txtInpName.Text -members $null -Confirm:$False -BypassSecurityGroupManagerCheck
write-host "Old Membership Removed Starting Adding Mew Members...." -ForegroundColor Cyan
AddMembers


<# Original Code

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
                $OutFileName = $LogDirectory + ($Grp.GrpName -replace " ","") + "-Date" + ((get-date -uformat %D).Replace("/", "-")) + "-Time" + ((get-date -uformat %T).Replace(":", "")) + ".txt"
                Get-DistributionGroupMember $Grp.GrpName -Resultsize Unlimited > $OutFileName
                write-host "Current Group membership count:  " (Get-DistributionGroupMember -Resultsize Unlimited $grp.GrpName).Count
				write-host "Updating group with new membership list"
                Update-DistributionGroupMember -identity $Grp.GrpName -members $null -Confirm:$False -BypassSecurityGroupManagerCheck
			    write-host "Old Membership Removed Starting Adding Mew Members...." -ForegroundColor Cyan
                AddMembers
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

        $GrpMem = Get-DistributionGroupMember $GrpName

        if ($GrpMem.count -gt 0)
        {
            foreach ($GrpMem in $GrpMem)
            {
                $mbx = get-mailbox $GrpMem.Name
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
            Get-DistributionGroupMember $UpdGrp > $OutFileName

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
#>