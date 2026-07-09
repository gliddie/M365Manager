$User = read-host -Prompt "Enter User" 

$User + " is a member of these groups:" 

$Group = get-distributiongroup -ResultSize Unlimited

ForEach ($Group in $group) 
{ 
   ForEach ($Member in Get-DistributionGroupMember $Group | Where { $_.Alias –eq $User }) 
   { 
      $Group.name 
   } 
}