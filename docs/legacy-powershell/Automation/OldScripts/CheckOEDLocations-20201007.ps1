<# Check to see if the DST.All Groups Created have active staff in Oracle
#
#  Called by:  ReportingMenu.ps1
#
#  04/06/2020 - Added code to bypass the UL LLC and UL INC groups
#>

$RemoveGroups = 0
$PrevName = ""

$DSTGroups = Get-DynamicDistributionGroup | Where-Object {($_.Name -like "DST.All*") -and ($_.Name -like "*UL Em*")} | select DisplayName | Sort-Object DisplayName
$DSTStaffGroups = Get-DynamicDistributionGroup | Where-Object {($_.Name -like "DST.All*") -and ($_.Name -like "*UL St*")} | select DisplayName | Sort-Object DisplayName
$DSTAllGroups = Get-DynamicDistributionGroup | Where-Object {($_.Name -like "DST.All*") -and (($_.Name -like "*UL Em*") -or ($_.Name -like "*UL St*"))} | select DisplayName | Sort-Object DisplayName
$ActiveRecs = Import-csv  \\usnbkutil100p\d$\scripts\DistributionList\OED_Extract5.csv | Where-Object{(($_."Assignment Status" -eq "A") -or ($_."Assignment Status" -eq "F"))} | select -unique Location | Sort-Object Location

$NoOfRecs = ($ActiveRecs.count - 1)
$NoOfDSTGrps = $DSTGroups.count
$collection = $ActiveRecs.GetEnumerator()
$CreateNewLine = ""
$DeleteOldLine = ""

$ActiveRecs.count
$DSTAllGroups.count

write-host "Checking Existing DST groups for active staff at each location" -ForegroundColor Cyan
write-host

foreach ($DSTGroups in $DSTGroups)
{
    $Name = $DSTGroups.DisplayName -replace "DST.All ",""
    If (($name -eq "UL Employees Only") -or ($name -eq "UL Staff"))
    {
        switch ($Name)
        {
            "UL Employees Only"
                {
                    $Name = $Name.Substring(0,$name.IndexOf(" Emp"))
                }
            "UL Staff"
                {
                    $Name = $Name.Substring(0,$name.IndexOf(" Staff"))
                }
        }
    }
    else
    {
        $Name = $Name.Substring(0,$Name.IndexOf(" UL"))
    }
    $Locfound = "No"

    If (($Name -ne "Asia Pacific") -and ($Name -ne "Canada") -and ($Name -ne "Europe") -and ($Name -ne "Latin America") -and ($Name -ne "United States") -and ($Name -ne "UL") -and ($Name -notlike "*UL COM*") -and ($Name -notlike "*UL ORG*") -and ($Name -ne $PrevName))
    {
        Foreach ($ActiveRecs in $Collection)
        {
            $LocExists = [bool] ($Name -eq ($ActiveRecs.Location -replace '[.]',''))
            #    $LocExists = [bool] ($ActiveRecs -contains $Name)
            If ($LocExists -eq $True)
            {
                write-host "Dynamic Distribtuion Lists Exists for" $Name
                $LocFound = "Yes"
            }
        }
    
        If ($LocFound -eq "No")
        {
           write-host "No Remaining Active Staff in" $Name -ForegroundColor Red
           write-host "Remove Dynamic Distribution Lists" $DSTGroups.DisplayName -ForegroundColor Red
           write-host "Remove Dynamic Distribution Lists" "DST.All" $Name "UL Staff" -ForegroundColor Red
           $RemoveGroups++
           $DeleteOldLine = $DeleteOldLine + "<li>" + $DSTGroups.DisplayName + "</li>"
           $DeleteOldLine = $DeleteOldLine + "<li>" + "DST.All " + $Name + " UL Staff </li>"

        }
    }
    $PrevName = $Name
    $collection.reset()
}

$NoGroup = 0
$HasGroup = 0

write-host
write-host "Checking OED Locations to determine if Additional DST Groups need to be Created" -ForegroundColor Cyan
write-host

#Check to see if there is a DST.All Group for the location

foreach ($ActiveRecs in $Collection)
{
    If ($ActiveRecs.Location -ne "")
    {
        $DLName = "DST.All " + ($ActiveRecs.Location -replace '[.]','') + " UL Staff"
        $DLExists = [bool] (Get-DynamicDistributionGroup $DLName -ErrorAction SilentlyContinue)

        If ($DLExists -eq $True)
        {
            $HasGroup++
            write-host "Dynamic Distribution List Exists for" ($ActiveRecs.Location -replace '[.]','')
        }
        else
        {
            write-host "Dynamic Distribution List Does Not Exist for" ($ActiveRecs.Location -replace '[.]','') -ForegroundColor Red
            $NoGroup++
            $CreateNewLine = $CreateNewLine + "<li>" + $ActiveRecs.Location + "</li>"
        }
    }
}
If ($RemoveGroups -gt 0)
{
    $DeleteOldLine = "<ul type=""disc"">" + $DeleteOldLine + "</ul>"
}
else
{
    $DeleteOldLine = "<ul type=""disc""><li>No DST.All Groups need to be removed</li></ul>"
}

If ($NoGroup -gt 0)
{
    $CreateNewLine = "<ul type=""disc"">" + $CreateNewLine + "</ul>"
}
else
{
    $CreateNewLine = "<ul type=""disc""><li>No DST.All Groups need to be created</li></ul>"
}

$server = "smtp-relay.ul.com"
$client = new-object system.net.mail.smtpclient $server
$from = New-Object System.Net.Mail.MailAddress "EnterpriseMessagingServices@ul.com" , "Enterprise Messaging Services"
$to = $from 
$message = new-object  System.Net.Mail.MailMessage $from, $to 
$message.IsBodyHtml = $true
$msgfont = "<basefont face=verdana size=2.5 color=black>"
$frstline = "<p>DST.All Groups that No Longer Have Active Staff:<br>"
$thrdline = "<p>OED Locations with Active Staff but no DST.All Groups:<br>"
$SendTo = "EnterpriseMessagingServices@ul.com"
$SendCC = "Sandi.Glazebrook@ul.com,RJ.Borja@ul.com"

$message.Subject = "OED Locations Review"
#$message.Attachments.Add($attach)
$message.Body = $msgfont + $frstline + $DeleteOldLine + $thrdline + $CreateNewLine
$message.Body = $message.Body + "             Total Number of DST.All Employee Only or Staff Groups:  " + $DSTAllGroups.count + "<br>"
$message.Body = $message.Body + "Number of Locations that have DST.All Employee Only Groups Created:  " + $NoOfDSTGrps + " - (2 Unisys & 1 Infosys Locations Non-Employees Only)<br>"
$message.Body = $message.Body + "        Number of Locations that have DST.All Staff Groups Created:  " + $DSTStaffGroups.count + "<br>"
$message.Body = $message.Body + "                                     Total Number of OED Locations:  " + $NoOfRecs + "<br>"
$message.Body = $message.Body + "              Number of Locations that have DST.All Groups Created:  " + $HasGroup + "<br>"
$message.Body = $message.Body + "              Number of Locations that need DST.All Groups Created:  " + $NoGroup + "<br>"
$message.Body = $message.Body + "              Number of Locations that need DST.All Groups Deleted:  " + $RemoveGroups + "<br>"

$message.To.Clear()
	
$message.Body = $msgfont + $message.body
$message.To.Add($SendTo)
$message.CC.Add($SendCC)
$client.Send($message)