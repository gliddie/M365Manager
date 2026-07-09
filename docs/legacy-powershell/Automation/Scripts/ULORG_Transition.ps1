
8888882@global.ul.com (Robert Plant)

#Remove Primary email and reassign the emp#global.ul.com as the new primary addr
#Hide from Address Book
#Remove SIp address and disable the Teams settings
$Usr = import-csv "c:\temp\ulorg_ursmbxs.csv" |Sort-Object mbx
Foreach ($u in $Usr)
{
    write-host "Processing " $u.mbx
    $ADUser = Get-ADUser $u.mbx -Properties *
    If ($ADUser.Company -eq "UL.ORG")
    {
        If ($ADUser.TargetAddress -like "*ul.org*")
        {
            write-host "Current Values"
            write-host "`tAccount UPN: " $ADUser.UserPrincipalName
            write-host "`tPrimarySMTPAddress: " $ADUser.TargetAddress
            write-host "`tProxyAddresses: " $ADUser.ProxyAddresses
            write-host "`tTargetAddress: " $ADUser.TargetAddress
            write-host "`tMailAddress: " $ADUser.Mail
            write-host "`tSIP Primary SIP Address: " $ADUser.'msRTCSIP-PrimaryUserAddress'
            write-host "`tSIP Enabled: " $ADUser.'msrtcsip-userenabled'
            write-host "`tSIP Deployment Locator: " $ADUser.'msRTCSIP-DeploymentLocator'
            write-host "`tSIP Federatopm Enabled: " $ADUser.'msRTCSIP-FederationEnabled'
            write-host "`tSIP Line: " $ADUser.'msRTCSIP-Line'
            write-host "`tSIP Option Flags: " $ADUser.'msRTCSIP-OptionFlags'
            write-host "`tSIP Home Server: " $ADUser.'msRTCSIP-PrimaryHomeServer'
            write-host "`tSIP User Policies: " $ADUser.'msRTCSIP-UserPolicies'

            write-host "`tClearing Teams (SIP) Settings"
    #        Set-ADUser $u.mbx -Clear msrtcsip-userenabled,msRTCSIP-DeploymentLocator,msRTCSIP-FederationEnabled,msRTCSIP-Line,msRTCSIP-OptionFlags,msRTCSIP-PrimaryHomeServer,msRTCSIP-PrimaryUserAddress,msRTCSIP-UserPolicies
            write-host "`tUpdating mail address with UPN " $ADUser.UserPrincipalName
    #        Set-ADUser $u.mbx -Replace @{mail=$ADUser.UserPrincipalName}
            write-host "`tUpdating mail targetaddress with UPN " $ADUser.TargetAddress
    #        Set-ADUser $u.mbx -Replace @{targetAddress=("SMTP:*"+$ADUser.UserPrincipalName)}
            If ($ADUser.ProxyAddresses -clike ("smtp:*"+$ADUser.UserPrincipalName))
            {
                $RAddr = "smtp:*" + $ADUser.UserPrincipalName
                write-host "`tReplacing UPN in Proxy with " $RAddr
    #            Set-ADUser $u.mbx -Remove @{ProxyAddresses=$RAddr}
            }
            $NewP = "SMTP:" + $ADUser.UserPrincipalName
            write-host "`tAdding New Primary Address " $NewP
    #        Set-AdUser $u.mbx -Add @{ProxyAddresses=$NewP} #this address should already be there
            write-host "`tRemoving targtAddress " $ADUser.TargetAddress
    #        Set-AdUser $u.mbx -Remove @{ProxyAddresses=$ADUser.TargetAddress}
            write-host "`tHiding mailbox from Exchange Address Book"
    #        Set-ADUser $u.mbx -Replace @{'msExchHideFromAddressLists'=$True}
        }
    }

    $AftUser = Get-ADUser $u.mbx -Properties *
    If ($AftUser.ProxyAddresses -like "*ul.org")
    {
        Foreach ($a in $AftUser.ProxyAddresses)
        {
            If ($AftUser.Company -eq "UL.ORG")
            {
                If ($a -like "*@ul.com")
                {
                    write-host "`tRemoving $a from " $u.mbx
#                   Set-AdUser $u.mbx -Remove @{ProxyAddresses=$a}
                }
            }
            If ($a -like "*ul.org")
            {
                write-host "`tRemoving $a from " $u.mbx
#                Set-AdUser $u.mbx -Remove @{ProxyAddresses=$a}
            }
        }
    }
}

#$NewU = get-mailbox -ResultSize Unlimited |Where-Object {($_.EmailAddresses -like "*ul.org") -and ($_.RecipientTypeDetails -eq "UserMailbox")}
#write-host "Number of Users with ul.org Address: " $NewD.count
#$NewU.Alias

##############################################################################################################
#Remove ul.org addresses from Distribution groups and hide all from the address book
$grp = import-csv "c:\temp\ulorgdl.csv" |Sort-Object grp
Foreach ($g in $grp)
{
    $DLGrp = get-DistributionGroup $g.Grp
    Write-host "`nProcessing " $g.Grp
            write-host "Current Values"
            write-host "`tPrimarySMTPAddress: " $DLGrp.PrimarySmtpAddress
            write-host "`tProxyAddresses: " $DLGrp.EmailAddresses

    #Check the Primary SMTP Address if UL.ORG replace with the @ul.onmicrosoft.com address
    If ($DLGrp.PrimarySMTPAddress -like "*ul.org")
    {
        $newPAddr = "SMTP:" + ($DLGrp.emailaddresses |Where-Object {$_ -like "*ul.onmicro*"})
        write-host "`tSetting New Primary Address " $newPAddr
#        Set-DistributionGroup $g.Grp -EmailAddresses @{remove=$DLGrp.PrimarySMTPAddress' Add=$newPAddr}
#        Set-DistributionGroup $g.Grp -HiddenFromAddressListsEnabled $True
    }

    $addr = $DLGrp.EmailAddresses
    If (($addr -clike "smtp*") -and ($addr -like "*ul.org*"))
    {
        Foreach ($a in $addr)
        {
            If (($a -clike "smtp*") -and ($a -like "*ul.org*"))
            {
#               Removes any other UL.ORG Aliases
                write-host "`tRemoving $a from " $DLGroup.DisplayName
#                Set-DistributionGroup $g.Grp -EmailAddresses @{remove=$addr}
            }
        }
    }
}

$NewD = get-distributiongroup -ResultSize Unlimited |Where-Object {$_.EmailAddresses -like "*ul.org"}
write-host "Number of Distribution Groups with ul.org Address: " $NewD.count
If ($NewD.Count -gt 0)
{
    write-host $NewD.displayName
}

#############################################################################################################
#Remove ul.org addresses from Shared mailboxes and hide all from the address book
$OMbx = import-csv "c:\temp\ulorg_shrmbxs.csv" |Sort-Object Mbx
Foreach ($m in $OMbx)
{
    write-host "`nProcessing shared mailbox: " $m.Mbx
    $Mbx = get-Mailbox ($m.Mbx).TrimEnd()
            write-host "Current Values"
            write-host "`tPrimarySMTPAddress: " $Mbx.PrimarySmtpAddress
            write-host "`tProxyAddresses: " $Mbx.EmailAddresses

    #Check the Primary SMTP Address if UL.ORG replace with the @ul.onmicrosoft.com address
    If ($Mbx.PrimarySMTPAddress -like "*ul.org")
    {
        $newPAddr = $Mbx.emailaddresses |Where-Object {$_ -like "*ul.onmicro*"}
        write-host "`tNew Primary Address: " $newPAddr
#        Set-Mailbox $m.Mbx -EmailAddresses @{remove=$Mbx.PrimarySMTPAddress' Add=$newPAddr}
#        Set-Mailbox $m.Mbx -HiddenFromAddressListsEnabled $True

#Find groups that manage access to this mailbox
    }

    $addr = $Mbx.EmailAddresses
    If (($addr -clike "smtp*") -and ($addr -like "*ul.org*"))
    {
        Foreach ($a in $addr)
        {
            If (($a -clike "smtp*") -and ($a -like "*ul.org*"))
            {
#               Removes any other UL.ORG Aliases
                write-host "Removing $a from " $Mbx.DisplayName
#                Set-Mailbox $m.Mbx -EmailAddresses @{remove=$addr}
            }
        }
    }
}

#$NewM = get-mailbox -ResultSize Unlimited |Where-Object {($_.EmailAddresses -like "*ul.org") -and ($_.RecipientType -eq "SharedMailbox")}
$NewM = get-mailbox -ResultSize Unlimited |Where-Object {($_.EmailAddresses -like "*ul.org")}
write-host "Number of Mailboxes with ul.org Address: " $NewM.count
If ($NewM.Count -gt 0)
{
    write-host $NewM.DisplayName
}

#Remove ul.org addresses from any Teams groups and hide from address book


#chemicalinsights.ul.org
#connect.ul.org
#convene.ul.org
#mail.myinfo.ul.org
#standards.ul.org
#ul.org


foreach ($m in $mbx)
{
    $det = get-mailbox $m.mbx
    write-host $m.mbx,$det.PrimarySMTPAddress,$det.EmailAddresses
}


$inp = Get-ADUser -Filter * -SearchBase "DC=global,DC=ul,DC=com" -Properties proxyAddresses |Sort-Object "SamAccountName"
$cnt = 0
foreach ($i in $inp)
{
    if ($i.proxyAddresses -like "*ul.org*")
    {
        write-host $i.UserPrincipalName "," $i.proxyAddresses
        $cnt++
    }
}
write-host "Number of Accounts: " $cnt