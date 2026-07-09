#### Process Instructions Script 
cls
#
switch ($O365Act)
{
    2
        {
            #Group Ownership Changes
            Write-Host " Instructions for Implementing changes to Distribution List or Security Group Ownership" -ForegroundColor Magenta
            Write-Host ""
            Write-Host " 1.  In the request check the name of the distribution list.  It should start with LST.XXX."
            Write-Host " 2.  If the name provided starts with MBX then the requestor is actually requesting a change"
            Write-Host "     to a Shared mailbox"
            Write-Host " 3.  For either change to requestor must be an owner of the distribution list or shared mailbox"
            Write-Host "     in order to process the request.  The only reasons that changes are made when requested by"
            Write-Host "     a non-owner would be if the current owner is no longer with UL or the current owner is out"
            Write-Host "     on an extended leave"
            Write-Host " 4.  If the requestor is not one of the current owners and (1) current owner is no longer with"
            Write-Host "     UL or (2) the current owner is out on an extended leave or vacation then we need to seek"
            Write-Host "     approval for the change from one of the existing owners"
            write-host " 5.  If the current owner is no longer with UL and they have been deleted fronm O365 their name"
            write-host "     will only appear in the Notes field they will not be in the ManagedBy field"
            Write-Host " 6.  If approval is needed use the template ""Distribution List Ownership Change Approval Needed"""
            Write-Host "     to request approval to make the change.  The message is sent to the current owner(s) with a"
            Write-Host "     CC: to the requestor"
            Write-Host " 7.  If approval is received or the requestor is an owner it is OK to proceed with the requested"
            Write-Host "     changes.  This may be (1) adding an additional owner, (2) replacing an existing owner, or "
            Write-Host "     (3) just removing an existing owner."
            Write-Host " 8.  This script will prompt for (1) name of the distribution list, (2) email address or employee"
            write-host "     number of the new owner, (3) the TASK of the request, (4) if removing or replacing an owner"
            write-host "      of the reqeust."
            Write-Host " 9.  Once the requested changes have been made use the template ""Distribution List Ownership Change"""
            Write-Host "     and send an email to the requestor, new owner and any owners that may have been removed from"
            Write-Host "     from the Ownership"
            write-host "10.  Enter ""Distribution List ownership changed attached email sent with details."" in the work"
            write-host "     history, attach a copy of the email sent and close the ticket"
            write-host ""
            write-host "When input file is complete hit return to execute the script....." -ForegroundColor Red -NoNewline
            $cont = read-host
        }
    3
        {
            #Shared Mailbox Ownership Changes
            Write-Host " Instructions for Implementing changes to Shared Mailbox Ownership" -ForegroundColor Magenta
            Write-Host ""
            Write-Host " 1.  In the request check the name of the Shared Mailbox.  The name should be XXX Name or "
            Write-Host "     MBX.Mailbox Name.XX."
            Write-Host " 2.  If the name provided starts with LST then the requestor is actually requesting a change"
            Write-Host "     to a Distribution List"
            Write-Host " 3.  For either change to requestor must be an owner of the shaed mailbox or distribution list"
            Write-Host "     in order to process the request.  The only reasons that changes are made when requested by"
            Write-Host "     a non-owner would be if the current owner is no longer with UL or the current owner is out"
            Write-Host "     on an extended leave"
            Write-Host " 4.  If the requestor is not one of the current owners and (1) current owner is no longer with"
            Write-Host "     UL or (2) the current owner is out on an extended leave or vacation then we need to seek"
            Write-Host "     approval for the change from one of the existing owners"
            Write-Host " 5.  If approval is needed use the template ""Shared Mailbox Ownership Change Approval Needed"""
            Write-Host "     to request approval to make the change.  The message is sent to the current owner(s) with a"
            Write-Host "     CC: to the requestor"
            Write-Host " 6.  If approval is received or the requestor is an owner it is OK to proceed with the requested"
            Write-Host "     changes.  This may be (1) adding an additional owner, (2) replacing an existing owner, or "
            Write-Host "     (3) just removing an existing owner."
            write-host " 7.  Once the requested changes have been made use the template ""Shared Mailbox Ownership Change"""
            Write-Host "     and send an email to the requestor, new owner and any owners that may have been removed from"
            Write-Host "     from the Ownership"
            write-host " 8.  Enter ""Shared Mailbox ownership changed attached email sent with details."" in the work"
            write-host "     history, attach a copy of the email sent and close the ticket"
            write-host ""
            write-host "When input file is complete hit return to execute the script....." -ForegroundColor Red -NoNewline
            $cont = read-host
        }
    4
        {
            # Creating New Distribution Lists
            Write-Host " Instructions for Creating New Distribution Lists" -ForegroundColor Magenta
            Write-Host ""
            Write-Host " 1.  Prepare the input file"
            write-host " 2.  In the directory e:\automation\NewDistributionGroup\Input edit the last log file created"
            write-host " 3.  Perform a Save-As and strip off all information after Input-NewDistributionGroup.csv"
            write-host " 4.  The first line is the variable names, keep this line and delete all others"
            write-host " 5.  The naming convention for Distribution Lists is:"
            write-host "     a.  LST.XXX.Name of Group where XXX = the 3 letter site code"
            write-host "     b.  LST.XX.Name of Group where XX = the 2 letter region code for the group"
            write-host "     c.  Groups that are corporate use CRP as the XXX (ex: LST.CRP.CDC)"
            write-host "     d.  Groups that are global groups use ""Global"" in the XXX (ex: LST.Global.CI Marketing)"
            write-host " 6.  Input file variables are:"
            write-host "     a.  SenderAuthentication - set to TRUE if no external sources or custom applications will use the list"
            write-host "                              - set to FALSE if external sources or custom applications will use the list"
            write-host "     b.  DisplayName - this is the display name of the distribution list; make sure to properly use"
            write-host "                       upper and lower case letters"
            write-host "     c.  Mail - this is the internet address of the list copy the name and remove any spaces"
            write-host "     d.  SDTicketNo - this is the ticket number of this request"
            write-host "     e.  OwnerName - this is entered into the Notes field so it is the friendly name not the internet address"
            write-host "                     of the owner(s)"
            write-host "     f.  ManagedBy - this is the internet address of the owners multiple entries are separated by commas"
            write-host "     g.  Members - this is the list of members for the group; multiple entries are separated by commas and only"
            write-host "                   internal UL addresses can be members of a group"
            write-host "     h.  LegacyAddress - this is used for groups that may be created as part of an M&A or groups that need to"
            write-host "                         have Lotus Notes resolve them"
            write-host " 7.  You can create multiple groups using one input file each new line will be for another group"
            write-host " 8.  Save the input file along the way but make sure to perform a final Save before proceeding with executing"
            write-host "     the script"
            write-host " 9.  Send a ""Distribution List New"" from the O365 Admin Templates - message is sent to the requestor and"
            write-host "     owner/owners"
            write-host "10.  Enter ""Distribution List created attached email sent with details."" in the work history, attach a copy of the"
            write-host "     email sent and close the ticket"
            write-host ""
            write-host "When input file is complete hit return to execute the script....." -ForegroundColor Red -NoNewline
            $cont = read-host
        }
    5
        {
            #Created New shared Mailboxes
            Write-Host "Instructions for Creating New Shared Mailboxes" -ForegroundColor Magenta
            Write-Host ""
            Write-Host " 1.  To create a Shared Mailbox there are multiple input files required they are:"
            write-host "     a.  e:\automation\NewSharedMailbox\Input"
            write-host "     b.  e:\automation\NewUSG\Input"
            write-host "     c.  e:\automation\AddFolderPermissions\Input"
            write-host "     d.  e:\automation\AddMailboxPermissionSharedMailbox\Input"
            write-host "     e.  e:\automation\ApplyRetentionPolicy\Input"
            write-host " 2.  For each input file edit the last log file created"
            write-host " 3.  Perform a Save-As and strip off all information after the "".csv"""
            write-host " 4.  The first line in each input file is the variable names, keep this line and delete all others"
            write-host ""
            write-host "Naming Conventions - always look in the address book for how the name is constructed for a given site" -ForegroundColor Yellow
            write-host " 5.  The naming convention for Shared Mailboxes is:"
            write-host "     a.  XXX Name of Mailbox where XXX = the 3 letter site code"
            write-host "     b.  XX Name of Maibox where XX = the 2 letter region code"
            write-host "     c.  In some instances we have maiboxes that start with ""UL"" or may be custom for integration of an M&A"
            write-host " 6.  The internet address for a shared mailbox is XXX.NameofMailbox@ul.com"
            write-host " 7.  Naming convention for the groups that grant access is MBX.Name of Mailbox.XX"
            write-host "     a.  where XX = ED for Editor Access"
            write-host "     b.  where XX = AU for Author Access"
            write-host "     c.  where XX = RE for Reader Access"
            write-host ""
            write-host "Input File Variables " -ForegroundColor Yellow      
            write-host " 8.  Input file variables are:"
            write-host "     a.  SharedMailbox and ApplyRetentionPolicy Input file:"
            write-host "           i.  Name - Name of the Shared Mailbox (ex: XXX Name of Mailbox)"
            write-host "          ii.  Mail - Internet address for mailbox (ex: XXX.NameofMailbox@ul.com)"
            write-host "         iii.  LegacyMail - Legacy address (ex: OldName@us.ul.com  or OldName@dewi.es)"
            write-host "     b.  NewUSG Input file - create a group for each level of access requested - .ED, .AU or .RE groups"
            write-host "           i.  DisplayName - this is the display name of the shared mailbox; make sure to properly use"
            write-host "               upper and lower case letters"
            write-host "          ii.  Mail  - this is the internet address of the mailbox copy the name add a ""."" after the XXX"
            write-host "               and remove any spaces"
            write-host "         iii.  SDTicketNo - this is the ticket number of this request"
            write-host "          iv.  OwnerName - this is entered into the Notes field so it is the friendly name not the"
            write-host "               internet address of the owner(s)"
            write-host "           v.  ManagedBy - this is the internet address of the owners multiple entries are separated by commas"
            write-host "          vi.  Members - this is the list of members for the group; multiple entries are separated by commas"
            write-host "               and only internal UL addresses can be members of a group"
            write-host "     c.  AddFolderPermissions Input file - this file is used to apply folder permissions"
            write-host "           i.  Mail - this is the internet address of the Shared Mailbox"
            write-host "          ii.  DelegateToAdd - this is the internet address of the access group"
            write-host "         iii.  Permission - this is the permission granted:"
            write-host "               -  for .ED enter ""Editor"""
            write-host "               -  for .AU enter ""PublishingAuthor"""
            write-host "               -  for .RE enter ""Reviewer"""
            write-host "     d.  AddMailboxPermissionSharedMailbox Input file - only groups that are being given Editor Access should"
            write-host "         be included in this file"
            write-host "           i.  MbxName - this is the internet address of the Shared Mailbox"
            write-host "          ii.  AccessGroup - this is the internet address of the access group (edior groups only)"
            write-host "     e.  ApplyRetentionPolicy Input file can just be created by doing a SaveAs from the SharedMailbox"
            write-host "         input file."
            write-host " 9.  You can create multiple mailboxes using one input file each new line will be for another mailbox"
            write-host "10.  Save the input files along the way but make sure to perform a final Save before proceeding with executing"
            write-host "     the script"
            write-host "11.  Send a ""Shared Mailbox New"" from the O365 Admin Templates - message is sent to the requestor, owner(s)"
            write-host "     and groups that were created for access"
            write-host "12.  Enter ""Mailbox created attached email sent with details."" in the work history, attach a copy of the"
            write-host "     email sent and close the ticket"
            write-host ""
            write-host "When input files are complete hit return to execute the script....." -ForegroundColor Red -NoNewline
            $cont = read-host
        }
    6
        {
            #New Room
        }
    7
        {
            #New Equipment
        }
    8
        {
            #New Restricted Room
        }
    9
        {
            #Enable Unified Messaging
        }
    10
        {
            #Display User Attributes
        }
}
    
