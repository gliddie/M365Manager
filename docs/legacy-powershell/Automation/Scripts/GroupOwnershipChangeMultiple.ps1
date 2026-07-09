<#
#  This script allows you to make the same changes to Multiple Distribution List or Security Group Ownership
#  Called by:  DistributionSecurityGroupMenu.ps1
#
#  Create the input file where you'd like the format is a single column titled "name" and then you place the names of the distribution lists to modify
#>

write-host "Enter Name and Location of Input File with the List Distribution Lists or Security Groups to Change " -ForegroundColor Yellow -NoNewline
$GrpFileName = Read-Host

$InpFile = import-csv $GrpFileName

write-host "Enter Task Number for this Request: " -ForegroundColor Yellow -NoNewline
$SDTicketNo = Read-Host

        write-host "Enter (1) to Replace the Owner"
        write-host "      (2) to Add an Owner"
        write-host "      (3) to Remove an Owner"
        write-host "      (4) No changes"
        write-host "Enter Option? " -ForegroundColor Red -NoNewline
        $GrpOpt = Read-Host
        If (($GrpOpt -eq 1) -or ($GrpOpt -eq 2))
        {
        Do {
           $GrpOwner = Read-Host "Enter the Disply Name, Email Address or Employee Number of New Owner"
           
           if ($GrpOwner.length -eq 5)
           {
                $GrpOwner = $GrpOwner + "@global.ul.com"
           }

           $MbxExists = [bool](get-mailbox $GrpOwner -ErrorAction SilentlyContinue)

           If ($MbxExists -ne "True")
           {
               write-host ""
               Write-Host "User specified does not exist" -ForegroundColor Red
           }
        } while ($MbxExists -ne "True")

        $mbx = get-mailbox $GrpOwner

        }
        
        If ($GrpOpt -eq 3)
        {
            #Remove Group Owner
            Write-Host
            Write-Host "Enter (1) to Remove from ManagedBy and Notes"
            Write-Host "      (2) to Remove from Notes only"
            Write-Host "      (3) to Remove from ManagedBy Only"
            Write-Host "Enter Option? " -ForegroundColor Red -NoNewline
            $RemAns = Read-Host
                             
            if (($RemAns -eq 1) -or ($RemAns -eq 3))
            {
                $GrpOwner = Read-Host "Enter Email Address or Employee Number of Owner to Remove"
                if ($GrpOwner.length -eq 5)
                {
                    $GrpOwner = $GrpOwner + "@global.ul.com"
                }
            }
                                       
            if ($RemAns -ne 3)
            {
                Write-host "Enter details to be removed from Notes field (include trailing spaces and commas)? " -ForegroundColor Yellow -NoNewline
                $CmtName = Read-Host
            }
        }

foreach ($InpFile in $InpFile)
{
    If ($FirstRun -eq "Yes")
    {
        Do
        {
    #        $GrpOwner = Read-Host "Enter the Disply Name, Email Address or Employee Number of New Owner"
            if ($GrpOwner.length -eq 5)
            {
                $GrpOwner = $GrpOwner + "@global.ul.com"
            }

            $MbxExists = [bool](get-mailbox $GrpOwner -ErrorAction SilentlyContinue)

            If ($MbxExists -ne "True")
            {
                write-host ""
                Write-Host "User specified does not exist" -ForegroundColor Red
            }
        } while ($MbxExists -ne "True")

        $mbx = get-mailbox $GrpOwner
        }
                    
    $GrpExists = [bool](Get-DistributionGroup $InpFile.Name -ErrorAction SilentlyContinue)
    If ($GrpExists -eq "True")
    {
        $Grp = get-distributiongroup $InpFile.Name
        $GrpInf = get-group $InpFile.Name

        write-host "Group Name "  $InpFile.Name
        write-host "Managed by "  $Grp.ManagedBy
        Write-Host "Group Notes " $GrpInf.Notes
        write-host

        switch ($GrpOpt)
        {
            1
                {
                    #Replace Group Owner
                    $mbx = get-mailbox $GrpOwner
                    Set-DistributionGroup $InpFile.Name -ManagedBy $GrpOwner -BypassSecurityGroupManagerCheck
                    Set-Group $InpFile.Name -Notes ("Owner:  " + $mbx.Name + " - Per " + $SDTicketNo)
                }
            2
                {
                    #Add Group Owner
                    $GrpInf.Notes = $GrpInf.Notes.Remove($GrpInf.Notes.IndexOf(" - Per")) + ", " + $mbx.Name + " - Per " + $SDTicketNo
                    Set-DistributionGroup $InpFile.Name -ManagedBy @{add="$GrpOwner"} -BypassSecurityGroupManagerCheck
                    Set-Group $InpFile.Name -Notes $GrpInf.Notes
                }
            3
                {
                    #Remove Group Owner
                    if ($RemAns -ne 3)
                    {
                        $GrpInf.Notes = $GrpInf.Notes.Remove($GrpInf.Notes.IndexOf("TASK")) + $SDTicketNo
                        set-Group $InpFile.Name -Notes $GrpInf.Notes.Replace($CmtName,"")  
                    }
                    if (($RemAns -eq 1) -or ($RemAns -eq 3))
                    {
                        Set-DistributionGroup $InpFile.Name -ManagedBy @{remove="$GrpOwner"} -BypassSecurityGroupManagerCheck
                    }
                }
        }

        If ($GrpOpt -lt 4)
        {
            write-host "Updated group details"
            $Grp = get-DistributionGroup $InpFile.Name
            $GrpInf = get-group $InpFile.Name
            write-host
            write-host "Updated Managed by "  $Grp.ManagedBy
            Write-Host "Updated Group Notes " $GrpInf.Notes
            write-host "***********************************************" -ForegroundColor Cyan
        }
    }
    else
    {
        Write-Host ""
        Write-Host "Distribution Group or Security Group Does Not Exist" $InpFile.Name -ForegroundColor Red
    }
}
