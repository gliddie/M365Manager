<#
The sample scripts are not supported under any Microsoft standard support 
program or service. The sample scripts are provided AS IS without warranty  
of any kind. Microsoft further disclaims all implied warranties including,  
without limitation, any implied warranties of merchantability or of fitness for 
a particular purpose. The entire risk arising out of the use or performance of  
the sample scripts and documentation remains with you. In no event shall 
Microsoft, its authors, or anyone else involved in the creation, production, or 
delivery of the scripts be liable for any damages whatsoever (including, 
without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use 
of or inability to use the sample scripts or documentation, even if Microsoft 
has been advised of the possibility of such damages.
#>

#requries -Version 2.0

<#
 	.SYNOPSIS
        This script is used to generate a report of Exchange Active Sync devices which connected to organization’s mailboxes, group mailboxes or a single mailbox 
    .DESCRIPTION
        This script is used to generate a report of Exchange Active Sync devices which connected to organization’s mailboxes, group mailboxes or a single mailbox
    .PARAMETER  Org
		The report will be contains all Exchange Active devices of all mailboxes in the organization. The default range of mailboxes is organization wide
    .PARAMETER  List
        The report will be contains all Exchange Active devices of the List's mailbox
    .PARAMETER  Identity
        The report will be contains Exchange Active devices of the specific mailbox
    .PARAMETER  Age
        The report will be contains Exchange Active devices which is synced in last 'Age' days
    .PARAMETER  EXOService
        This parameter is used to indicate whether the exchange service is on Exchange Online
    .PARAMETER  Path
        Specifies the path which the report will be stored in. If this parameter is empty, the report will be stored in current directory
    .PARAMETER  FileName
        Specifies the file name which the report will be. If this parameter is empty, the default name is ExchangeActiveSyncDeviceReport.csv
    .EXAMPLE
        GetActiveSyncDeviceReport.ps1 -EXOService 
        Generate a Exchagne Active Sync devices report of all mailboxes in Exchange Online
    .EXAMPLE
        GetActiveSyncDeviceReport.ps1 -List $List -Path 'D:\ExchangeReport' -FileName EASstatics.csc
        #Generate a Exchagne Active Sync devices report of a group mailboxes, and this report will be stored in D:\ExchangeReport and the file name will be EASstatics.csc
    .EXAMPLE
        GetActiveSyncDeviceReport.ps1 -Identity 'Joe White' -Age 20
        Generate a Exchagne Active Sync devices report of Joe White and this device is synce in last 20 days
#>

[CmdletBinding(DefaultParametersetName="Org")] 
Param
(
     [parameter(ParameterSetName="Single",Mandatory=$false)]
     [String]$Identity=$null,
     [parameter(ParameterSetName="Group",Mandatory=$false)]
     [Array]$List=@(),
     [parameter(ParameterSetName="Org",Mandatory=$false)]
     [switch]$Org=$true,
     [Parameter(Mandatory=$false)]
     [int]$Age = 0,
     [Parameter(Mandatory=$false)]
     [switch]$EXOService,
     [Parameter(Mandatory=$false)][ValidateScript({Test-Path $_ -PathType 'Container'})] 
     [string] $Path=$null,
     [Parameter(Mandatory=$false)]
     [string]$FileName = "ExchangeActiveSyncDeviceReport.csv"
)
Begin
{
    if($Path -eq "")
    {
        $Path = Split-Path -parent $MyInvocation.MyCommand.Definition
    }
    $Today = Get-Date
    $reportFileInfo = $Path +'\' + $FileName
    if($Identity -or ($List.Count -gt 0))
    {
        $Org = $false
    }
    if($EXOService) #This is used for EXO service
    {
        $exitingSnaping = Get-PSSnapin -Verbose:$false | Where-Object {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}
        $existingSession = Get-PSSession -Verbose:$false | Where-Object {(($_.ConfigurationName -eq "Microsoft.Exchange") -and ($_.ComputerName -notlike "*outlook.com" ))}
        if(($exitingSnaping -ne $null) -or ($existingSession -ne $null) )
        {
            Write-Error "Please run a PowerShell instance instead of running Exchange Management Shell"
            Exit
        }
        Try
		{
			#If the remote powershell session does not exist, create a new session.
			$existingSession = Get-PSSession -Verbose:$false | Where-Object {($_.ConfigurationName -eq "Microsoft.Exchange") -and ($_.ComputerName -like "*outlook.com" )}
			if ($existingSession -eq $null) 
            {
				$verboseMsg = "Creating a new session to https://ps.outlook.com/powershell."
				$pscmdlet.WriteVerbose($verboseMsg)
				$O365Session = New-PSSession -ConfigurationName Microsoft.Exchange `
				-ConnectionUri "https://ps.outlook.com/powershell" -Credential $Credential `
				-Authentication Basic -AllowRedirection
				#If session is newly created, import the session.
				Import-PSSession -Session $O365Session -Verbose:$false
				$existingSession = $O365Session
			} 
            else 
            {
				$verboseMsg = "Found existing session, new session creation is skipped."
				$pscmdlet.WriteVerbose($verboseMsg)
			}
        }
        Catch
		{
		    write-error $Error[0]
            exit
		}
    }
    else
    {
        Try
        {
            $exitingSnaping = Get-PSSnapin -Verbose:$false | Where-Object {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}
            $existingSession = Get-PSSession -Verbose:$false | Where-Object {($_.ConfigurationName -eq "Microsoft.Exchange") -and ($_.ComputerName -notcontains "outlook.com" )}
        if(($exitingSnaping -eq $null) -and ($existingSession -eq $null) )
            {
		        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
                . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	            Connect-ExchangeServer -auto -AllowClobber
	        }
        }
	    catch
	    {
		    Write-Error "No Exchange Server connected"
            exit
	    }
    }
}
    
Process
{
    Write-Host "Getting the list of mailboxes which has Active Sync devices connected"
    $MailboxesHasEASConnected = @()
    if($Org -eq $true)
    {
        $MailboxesHasEASConnected += Get-CASMailbox -Resultsize Unlimited | Where-Object {$_.HasActiveSyncDevicePartnership -eq $true}
        Write-Host "There are $($MailboxesHasEASConnected.Count) mailboxes which has Active Sync devices connected"
    }
    elseif($Identity -ne "")
    {
       $MailboxesHasEASConnected = Get-CASMailbox -identity $Identity -ErrorAction stop | Where-Object {$_.HasActiveSyncDevicePartnership -eq $true}
       if($MailboxesHasEASConnected -eq $null)
       {
            Write-Error "This mailbox doesn't have EAS device connected"
            Exit
       }
    }
    else
    {
        
        foreach($mailboxIdentity in $List)
        {
            $MailboxesHasEASConnected += Get-CASMailbox -identity $mailboxIdentity -ErrorAction stop | Where-Object {$_.HasActiveSyncDevicePartnership -eq $true}
        }
        if($MailboxesHasEASConnected -eq $null)
        {
            Write-Error "The mailboxes in this list don't have EAS device connected"
            Exit
        }
    }

    $strPath = "C:\temp\ActiveSyncReport.csv"

    $text = "User, DeviceModel,DeviceOS"
    out-file -FilePath $strPath -InputObject $text
    $cnt = 1

    $devicesReport = @()
    Foreach ($Mailbox in $MailboxesHasEASConnected)
    {
        $EASDeviceStats = @()
        $EASDeviceStats += Get-MobileDeviceStatistics -Mailbox $Mailbox.Identity

#        $text = "{0}, {1}, {2}" -f $EASDeviceStats.Displayname,$EASDeviceStats.DeviceModel,$EASDeviceStats.DeviceOS
#        out-file -FilePath $strPath -InputObject $text -Append

    	Write-Host "$cnt,$($Mailbox.Identity) has $($EASDeviceStats.Count) device(s)"
        $cnt = $cnt + 1
        get-mobiledevice -mailbox $mailbox.identity |ft Identity,DeviceModel,DeviceOS
    	write-host 

        $MailboxInfo = Get-Mailbox $Mailbox.Identity | Select DisplayName,PrimarySMTPAddress
        

        Foreach ($EASDevice in $EASDeviceStats)
        {
            $LastSyncAttempt = $EASDevice.LastSyncAttemptTime

            if ($LastSyncAttempt -eq $null)
            {
                $syncAge = "Never"
            }
            else
            {
                $syncAge = ($Today - $LastSyncAttempt).Days
            }
           
            if ($syncAge -ge $Age -or $syncAge -eq "Never")
            { 
                $deviceObj = New-Object PSObject
                $deviceObj | Add-Member -membertype NoteProperty -Name "Display Name" -Value $MailboxInfo.DisplayName
                $deviceObj | Add-Member -membertype NoteProperty -Name "Email Address" -Value $MailboxInfo.PrimarySMTPAddress
                $deviceObj | Add-Member -membertype NoteProperty -Name "Last Sync Days" -Value $syncAge
                $deviceObj | Add-Member -membertype NoteProperty -Name "DeviceOS" -Value $EASDevice.DeviceOS
                $deviceObj | Add-Member -membertype NoteProperty -Name "DeviceModel" -Value $(If($EASDevice.DeviceModel -eq ""){"NULL"}Else{$EASDevice.DeviceModel})
                $deviceObj | Add-Member -membertype NoteProperty -Name "DeviceID" -Value $(if($EASDevice.DeviceID -eq ""){"NULL"}Else{$EASDevice.DeviceID})
                $deviceObj | Add-Member -membertype NoteProperty -Name "Status" -Value $(if($EASDevice.Status -eq ""){"NULL"}Else{$EASDevice.Status})
                $deviceObj | Add-Member -membertype NoteProperty -Name "FirstSyncTime" -Value $(if($EASDevice.FirstSyncTime -eq ""){"NULL"}Else{$EASDevice.FirstSyncTime})
                $deviceObj | Add-Member -membertype NoteProperty -Name "LastSyncAttemptTime" -Value $(if($EASDevice.LastSyncAttemptTime -eq ""){"NULL"}Else{$EASDevice.LastSyncAttemptTime})
                $deviceObj | Add-Member -membertype NoteProperty -Name "LastSuccessSync" -Value $(if($EASDevice.LastSuccessSync -eq ""){"NULL"}Else{$EASDevice.LastSuccessSync})#>

                $devicesReport += $deviceObj
            }      
        }
    }
}

End
{
    $devicesReport | Export-Csv -NoTypeInformation $reportFileInfo
}

