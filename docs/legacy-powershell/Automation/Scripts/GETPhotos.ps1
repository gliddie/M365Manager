#Import-Module MSOnline 
#$365Cred = Get-Credential 
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $365Cred -Authentication Basic -AllowRedirection 
#Import-PSSession $Session 
#Connect-MsolService -Credential $365Cred 

# Set-UserPhoto -Identity [user name] -PictureData ([System.IO.File]::ReadAllBytes("C:\export\Photo$($User.Identity).jpg")) 

$Cont ="N"
write-host "Enter Last Employee Number processed (Enter 'New' for first time run): " -ForegroundColor Cyan -NoNewline
$LastEmp = read-host

$objUsers = get-mailbox -ResultSize Unlimited -SortBy Alias |Where-Object{$_.RecipientTypeDetails -eq "UserMailbox"}| select UserPrincipalName,Alias
$objusers.count

$ErrorActionPreference = "SilentlyContinue"

Foreach ($objUser in $objUsers) 
{ 
    If (($objuser.alias -eq $LastEmp) -or ($LastEmp -eq "New"))
    {
        $Cont = "Y"
    }

    If ($Cont -eq "Y")
    {
        if ([bool](Get-UserPhoto $objUser.UserPrincipalName)) 
        {
            $user = Get-UserPhoto $objUser.UserPrincipalName 
            $user.PictureData |Set-Content "C:\Export\Photo\$($ObjUser.Alias).jpg" -Encoding byte 
        }
        else
        {
            #Never uploaded a photo 
            write-host "No Photo Exists for: " $objUser.UserPrincipalName 
        }
    }
} 
