#Rename files in the Reports directories to have the EmpNo in the name
$Path = Read-Host "Enter the directory path including (ex: E:\SDAP\Details\)"
$LastChar = $Dir.substring(($Dir.length-1),1)
If ($lastchar -ne "\")
{
    $Path = $Path + "\"
}
#$path = "E:\SDAP\StandardTermination\Report\2018\"
$TxtPattern = Read-Host "Enter the text to search for the EmpNo:"
$files = Get-ChildItem -Path $Path
write-host "Number of files in this directory: " $Files.count
foreach ($f in $files)
{
    $ENo = ""
    $LocColon = ""
    $ENoLength = ""
    $rec = (Select-String -Path ($Path+$f.Name) -Pattern $TxtPattern).ToString()
#    $rec = (Select-String -Path ($Path+$f.Name) -Pattern 'Starting Standard Termination for:').ToString()
    $LocColon = ($rec.IndexOf("for:") + 4)
    $ENoLength = ($rec.length - ($LocColon))
    $ErrorActionPreference = "SilentlyContinue"
    [Int]$ENo = ($rec.substring(($rec.length-6),6)).Trim()
    $ErrorActionPreference = "Continue"
    If (($ENoLength -ge 5) -and ($rec -notlike "*-EmpNo*"))
    {
        $Log = $rec.IndexOf(".log")
        $Filename = ($rec.Substring(0,($Log+4))).Trim()
        $AddDetails = "-EmpNo" + $ENo + "-Date"
        $newFilename = $Filename.Replace("-Date",$AddDetails)
        write-host "`nRenaming $Filename"
        write-host "`tto $newFileName"
        Rename-Item $Filename $newFilename
    }
    else
    {
        #This section moves the EmpNo notation in the filename
        $rec = (Select-String -Path ($Path+$f.Name) -Pattern 'Starting Standard Termination for:').ToString()
        $LocTime = $rec.IndexOf("-Time")
        $LocEmpNo = $rec.IndexOf("-EmpNo")
#        If ($LocEmpNo -gt $LocTime)
        If (($LocEmpNo -gt $LocTime) -and ($ENoLength -ge 5) -and ($ENoLength -le 6))
        {
            $Log = $rec.IndexOf(".log")
            $Filename = $rec.Substring(0,($Log+4))
            $File = $rec.Substring(0,($rec.Indexof("-EmpNo")))
            If (($File.Length -gt 0) -and ($ENo -is [int]))
            {
                $AddDetails = "-EmpNo" + $ENo + "-Date"
                $newFilename = ($File.Replace("-Date",$AddDetails)) + ".log"
                write-host "`nRenaming $Filename"
                write-host "`tto $newFileName"
                Rename-Item $Filename $newFilename
            }
        }
    }
}
