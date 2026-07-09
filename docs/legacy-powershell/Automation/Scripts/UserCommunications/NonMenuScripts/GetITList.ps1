Write-Host "Enter Name of Distribution List: " -ForegroundColor Yellow -NoNewline
$DLLookup = Read-Host

$MEM = get-DistributionGroupMember $DLLookup -ResultSize Unlimited | select Alias,RecipientType
start-transcript
foreach ($mem in $mem)
	{
		IF ($MEM.RecipientType -eq "MailUser")
		{
			get-MailUser $mem.alias |ft Alias,PrimarySMTPAddress,DisplayName
		}
		else
		{
			get-Mailbox $mem.alias |ft Alias,PrimarySMTPAddress,DisplayName
		}
	}
stop-transcript	
	