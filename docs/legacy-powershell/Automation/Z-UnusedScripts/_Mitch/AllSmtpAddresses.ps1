$ReportFile = "e:\automation\ulMail.txt"

$users = Get-Recipient -ResultSize Unlimited

foreach ($user in $users)

{

$addmail = $user.EmailAddresses

foreach ($address in $addMail)

  {

       Write-Host $address

       $LineToWrite = $User.DisplayName + "`t" + $address

       Out-File -filepath $ReportFile -append -noClobber -inputObject $LineToWrite

  }

}