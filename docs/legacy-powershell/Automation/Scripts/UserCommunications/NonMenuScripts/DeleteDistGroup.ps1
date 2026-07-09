$InputDL = Read-Host "Name of Distribution List"

start-transcript

get-DistributionGroup $InputDL

get-DistributionGroup $InputDL |fl

get-Group $InputDL |fl

get-DistributionGroupMember $InputDL |ft Alias,Name,RecipientType

remove-DistributionGroup $InputDL -BypassSecurityGroupManagerCheck

Stop-transcript