	# Rename the input file for future reference & Remove PS Session
    $uDate = get-date -uformat %D
	$uTime = get-date -uformat %T
	$Date  = $uDate.Replace("/", "-")
	$Time  = $uTime.Replace(":", "")
    $InputFile = "c:\temp\Input-ConfigMailbox.csv"

# Rename the input file for future reference & Remove PS Session
    $Error = ""

#    '    03/25/2020 SAG Added code so if the input file is not renamed the agent will be told the file is in use.  It will attempt to rename 10 times and then print a message for the agent to do it manually    Rename-Item $InputFile "c:\temp\FileRename.log" -ErrorAction SilentlyContinue
    if (Test-Path $InputFile)
    {
        $loop = 0
        Do
        {
            write-host "The file" $InputFile "is open by another process please close the file and hit returnt to continue" -ForegroundColor Red -NoNewline
            $Cont = read-host
            Rename-Item $InputFile ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log") -ErrorAction SilentlyContinue
            $loop++
        } while ((Test-Path $InputFile) -and ($loop -lt 10))
        
        If ($loop -ge 10)
        {
            Write-Host "Unable to rename this file" $InputFule "please manually rename it to" ($InputFile + "-" + "Date" + $Date + "Time" + $Time + ".log")
        }
    }
