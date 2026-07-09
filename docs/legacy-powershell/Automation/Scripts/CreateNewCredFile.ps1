<#
#    This script create a new credential file that is used to connect to O365
#    Created by Sandi Glazebrook 3/28/2017
#
#   Called by:  O365AdminMenu.ps1
#
#>

function Build-NewCredForm
{
    $Global:form = New-Object Windows.Forms.Form 
    $Global:form.FormBorderStyle = "FixedToolWindow" 
    $Global:form.Text = "Create New Credential File" 
    $Global:form.StartPosition = "CenterScreen"
    $Global:form.Size = New-Object System.Drawing.Size(300,200) #(W,H)
    
    $Top = 30 
    $InpLeft = 40
    
    ## Create for Standard Account
    $Script:chkStdAcct = New-Object Windows.Forms.checkbox
    $Script:chkStdAcct.Left = $InpLeft; $Script:chkStdAcct.Width = 200; $Script:chkStdAcct.Top = $Top
    $Script:chkStdAcct.Text = "New File for Standard Account"
    $Script:chkStdAcct.Checked = $false   # set a default value
    $Global:form.Controls.Add($Script:chkStdAcct)

    $Top = $Top + 20
    ## Create for Admin Account
    $Script:chkAdmAcct = New-Object Windows.Forms.checkbox
    $Script:chkAdmAcct.Left = $InpLeft; $Script:chkAdmAcct.Width = 200; $Script:chkAdmAcct.Top = $Top
    $Script:chkAdmAcct.Text = "New File for Admin 'A' Account"
    $Script:chkAdmAcct.Checked = $false   # set a default value
    $Global:form.Controls.Add($Script:chkAdmAcct)

    $Top = $Top + 20
    ## Create for Both Accounts
    $Script:chkBothAcct = New-Object Windows.Forms.checkbox
    $Script:chkBothAcct.Left = $InpLeft; $Script:chkBothAcct.Width = 200; $Script:chkBothAcct.Top = $Top
    $Script:chkBothAcct.Text = "New File for Both Accounts"
    $Script:chkBothAcct.Checked = $false   # set a default value
    $Global:form.Controls.Add($Script:chkBothAcct)

    Add-FormStandardButtons
}

#write-host "Select (1) to create cred file for non-A account"
#write-host "       (2) to create for A account"
#write-host "       (3) to create for both accounts"
#write-host "`n       (0) to Exit"
#$Selection = Read-host "Choice"

Build-NewCredForm
Publish-Form

If ($Global:Result -eq "OK")
{

$Dev = $env:ComputerName
$me = whoami
$CredENo = $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)).replace(".","")

$UsrName = $CredENo + "@global.ul.com"
$dir = "c:\users\" + $CredENo + "\documents\"
$File = "my" + $CredENo + "File.xml"
$Global:CredFile = $dir + $File

If ($Dev -ne "USNBKEMES100P")
{
#    If ((get-mailboxstatistics $CredENo |Out-Gridview -ErrorAction SilentlyContinue) -eq $False)
#    {
        #New credentials ill need to be manually entered here to proceed.
#        Connect-ExchangeOnline
#        get-mailboxstatistics $CredENo |Out-Gridview
        get-netadapter |Out-GridView
#    }

    $wshell = New-Object -ComObject Wscript.Shell
    $Output = $wshell.Popup("To resolve an error when updating credentials that occurs on workstations the statistics of your mailbox will be opened in another windows.  Once you have entered your credentials you can close this window",0,"Notice",0+64)
}

If (($Script:chkStdAcct.Checked -eq $True) -or ($Script:chkBothAcct.Checked -eq $True))
{
    Get-Credential -UserName $UsrName -Message 'Enter Password' | Export-Clixml $CredFile
    write-host "New Credential File Created for: " $me
}

If (($Script:chkAdmAcct.Checked -eq $True) -or ($Script:chkBothAcct.Checked -eq $True))
{
    $ACredENo = "A" + $me.Substring(($me.IndexOf("\")+1),$me.length-($me.IndexOf("\")+1)).replace(".","")
    $UsrName = $ACredENo + "@global.ul.com"
    $dir = "c:\users\" + $CredENo + "\documents\"
    $File = "my" + $ACredENo + "File.xml"
    $CredFile = $dir + $File
    Get-Credential -UserName $UsrName -Message 'Enter Password' | Export-Clixml $CredFile

    write-host "New Credential File Created for: " ("global\" + $ACredENo)
}

}