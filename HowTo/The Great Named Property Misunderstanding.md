# The Great Named Property Misunderstanding

> **Note**: There is now a much easier way to check and fix this issue for Exchange Online.  Please see https://learn.microsoft.com/en-us/troubleshoot/exchange/outlook-issues/named-property-quota-exceeded.

Ever since they were introduced, it seems that developers new to Exchange programming have misunderstood named properties.  This can go unnoticed for a long time, but there are a couple of specific common issues that can block mailbox migrations (particularly from on-premises environments to Office 365).

The two issues are these:
* The developer uses a single name for the named property (more specifically, the first ten characters are the same), but a different Guid each time they create the property on an item.
* The developer uses a single Guid to identify their named properties (which in itself is recommended), but creates too many properties under this Guid (the hard limit that will prevent migration is 2000, but I would suggest that any application that is creating more than ten properties should seriously reconsider the architecture e.g. just use a single property to store all your application data in a blob).

Note that there are other ways to create too many named properties that may not match the patterns described above, but the scripts described here are likely to be of limited (if any) help in those cases.

Named properties have their own table within a mailbox, and for performance reasons the size of this table is limited.  Too many named properties can exhaust this table, as the table contains all named properties that have been defined in that mailbox.  When the table is exhausted, no new named properties can be created and you may start to receive errors when accessing existing items in the mailbox.

Ever since named properties were introduced, we've intermittently had issues of the named property table being exhausted (due to misbehaving applications), and have published the following advice to clean the mailbox:
* List all the named properties in the mailbox (using a tool such as MFCMapi).
* Go through the mailbox (manually) deleting the items or properties from the items until the named property count is within reasonable limits (usually this involves deleting all named properties that were created by a particular application).
* Purge retention (as otherwise the deleted items, with the offending properties, will still be in the mailbox).
* Move the mailbox to another database using the `-DoNotPreserveMailboxSignature` parameter (this moves the mailbox but does not copy the named property table, which is instead rebuilt).

The process above works, but it can be incredibly time consuming to do the first two steps.

# Automatically repairing the mailbox

**Note that this process only works against an on-premises mailbox. The Check-NamedProps script cannot be run against an Exchange Online mailbox.**  You need to ensure that any invalid named properties are removed from the mailbox before migration to Office 365.

We regularly receive cases where customers have been in the process of migrating to Office 365 and the named properties issue has blocked them.  Given that there may be many mailboxes affected, and hundreds of named properties per mailbox, the manual process to clear the properties is simply not feasible.  This guide outlines an automated process for clearing the named properties in such a scenario.

It is the first two steps that are the tricky ones, but so long as you know the name or Guid of the named properties that need to be deleted, then the process to remove them can be scripted.

I wrote two scripts to do this.  The first script is run against the mailbox (or mailbox database, or whole organisation if you like) and it will search for the offending named properties and create a list of items that contain them (including which properties are on those items).  The second script runs against the mailbox to remove the named properties (no need to delete the items anymore, the script targets the properties specifically and leaves the rest of the item alone).

Check-NamedProps: https://github.com/David-Barrett-MS/PowerShell-EWS-Scripts/blob/master/Legacy/Check-NamedProps.ps1

Delete-ByEntryId: https://github.com/David-Barrett-MS/PowerShell-EWS-Scripts/blob/master/Legacy/Delete-ByEntryId.ps1

## Example repair of a single mailbox

To repair a single mailbox, the process is as follows:

* Download and install the EWS Managed API on the machine being used to run the scripts (which don't need to be run on Exchange, any client machine with PowerShell available will do).  The current version of the EWS API is available [from Github](https://github.com/officedev/ews-managed-api).
* From your Exchange server , copy the ManagedStoreDiagnosticFunctions.ps1 script into your scripts folder (which should also contain Check-NamedProps.ps1 and Delete-ByEntryId.ps1).
* Optionally create a folder for storing the log files of Check-NamedProps (the script dumps the named properties found, and the EntryIds of objects containing them).  In the following examples the files will be created in the same folder as the script.
* Load Check-NamedProps.ps1 to make the Check-NamedProps function available (the script needs to be dot dot loaded).
* If you don't know which properties are causing the issue, then first of all call Check-NamedProps against the mailbox to retrieve the named property table: `Check-NamedProps -Mailbox "badnamedprops@e19.local" -powershellurl http://e1.e19.local/PowerShell/`.  Once run, you can check the exported .namedprops file (which is Xml) to identify the name properties to search for in the next step.
* Call Check-NamedProps to dump the entry Ids of any items that have one of the named properties to be removed (note that this can take a long time to run).  The script supports wild card name searches, so if you want to find all named properties that start with 'badproperty', you can search for 'badproperty*': `Check-NamedProps -Mailbox "badnamedprops@e19.local" -powershellurl http://e1.e19.local/PowerShell/ -SearchNamedProp "badnamedprop*" -DumpEntryIds`
* Check the output folder to see the filename for the list of EntryIds that the above creates (the script uses the mailbox Guid for the dump files, and each EntryId file will have the SMTP address of the mailbox as the first line).  The next step is to use Delete-ByEntryId.ps1 to delete the properties from each of the listed items.  This process uses EWS, and by default uses ApplicationImpersonation to access the mailbox (so the authenticating account must be granted the relevant permissions).  Optionally, the whole item can be deleted instead by specifying -DeleteItems switch.  Example call: `.\Delete-ByEntryId.ps1 -EntryIds "C:\Scripts\60585fa4-0ef6-49fd-8dee-56f3f00f79c8.EntryIds.txt" -EwsUrl "https://e1.e19.local/EWS/Exchange.asmx" -Credentials ($ewsAuth) -LogFile "log.txt"`
* After the properties have been deleted, they will still be in the named properties table of the mailbox, so the final step is to move the mailbox to another database without preserving that table (which means it is rebuilt, and therefore won't contain the now deleted named properties).  Do this using [New-MoveRequest](https://docs.microsoft.com/en-us/powershell/module/exchange/new-moverequest?view=exchange-ps) and ensure that `-DoNotPreserveMailboxSignature` is specified.
* Once the mailbox move is complete, the bad named properties should be gone.  You can confirm this by running Check-NamedProps against it with the same search parameters (it should return 0 named properties after the move).


