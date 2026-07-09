$mailbox = Get-Mailbox -id 96151@global.ul.com
$mailAddress = $mailbox.PrimarySmtpAddress.ToString();
 
[Reflection.Assembly]::LoadFile("C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll") | Out-Null
$s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
$s.AutodiscoverUrl($mailAddress);
 
#Bind Root of Mailbox
$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$mailAddress)
$MailboxRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($s,$folderid)
 
# Function to Get Folder ID from Folder Path
Function GetFolder()
{
	# Return a reference to a folder specified by path
 
	$RootFolder, $FolderPath = $args[0];
 
	$Folder = $RootFolder;
	if ($FolderPath -ne '\')
	{
		$PathElements = $FolderPath -split '\\';
		For ($i=0; $i -lt $PathElements.Count; $i++)
		{
			if ($PathElements[$i])
			{
				$View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0);
				$View.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;
				$View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly;
 
				$SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i]);
 
				$FolderResults = $Folder.FindFolders($SearchFilter, $View);
				if ($FolderResults.TotalCount -ne 1)
				{
					# We have either none or more than one folder returned... Either way, we can't continue
					$Folder = $null;
					Write-Host "Failed to find " $PathElements[$i];
					Write-Host "Requested folder path: " $FolderPath;
					break;
				}
 
				$Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($s, $FolderResults.Folders[0].Id)
			}
		}
	}
 
	$Folder;
}
 
#Bind Inbox of Mailbox
$Inboxid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$mailAddress)
$InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($s,$Inboxid)
 
# Get x Items in Inbox
$mails = $InboxFolder.FindItems(5) 
 
    foreach($Item in $mails.Items)
	{  
	# Get list of Reciepients from Message ID that has been BCC'ed 
	$mailrecipeients = Get-MessageTrackingLog -MessageID $item.InternetMessageId | where{$_.RecipientStatus -eq "Bcc"} |% {$_.recipients}
 
		if ($mailrecipeients -match "Test@Test.com") 
		{
			$FolderObject = GetFolder($MailboxRoot, "\Moved\Folder");
			try
			{
				$Item.Move($FolderObject.Id) | out-null;
			}
				catch
			{
				Write-Host "Failed to move item", $Item.Id.UniqueId
		}
 
		}
 
 
 
	}