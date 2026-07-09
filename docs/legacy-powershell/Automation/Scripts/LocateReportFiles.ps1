#get the location of a file with a given string
$startPath = 'e:\SDAP'
$ENo = Read-Host "`n`nEnter Employee Number " 
$magicWord = "No" + $ENo
$RecsFound = "No"

Write-host "Searching through the SDAP directory for report files for actions taken..."  -ForegroundColor Cyan

foreach ($file in (Get-ChildItem $startPath -Recurse -File |Where-Object {$_.Name -like "Report*"}))
{
    If ($File.DirectoryName -like "*Report*")
#   If (($File.DirectoryName -like "*Report*") -and (($File.DirectoryName -like "*Termination*") -or ($File.DirectoryName -like "*Emergency*") -or ($File.DirectoryName -like "*ConfigMbx*") -or ($File.DirectoryName -like "*Purge*") -or ($File.DirectoryName -like "*Phone*") -or ($File.DirectoryName -like "*Permission*") -or ($File.DirectoryName -like "*Rename*") -or ($File.DirectoryName -like "*Access*") -or ($File.DirectoryName -like "*ADAccount*")  ))
    {
        if($file.Name -match $magicWord)
        {
            If ($RecsFound -eq "No")
            {
                write-host "`nDirectory`t`t`t`t`t`t FileName" -ForegroundColor Yellow
                $RecsFound = "Yes"
            }
            write-host $File.DirectoryName,"`t",$file.Name
        }
    }
}
