# Creating Welcome messages automatically for Exchange

A question that we are regularly asked is how to automatically create welcome messages when new mailboxes are created in an Exchange environment.  There are various approaches that could be used, and previously I have published examples of how to do this with Exchange transport agents (which will trigger as soon as the mailbox is created).  Transport agents can't be used with Exchange Online, so a different technique is needed.

The easiest way to automate this is to schedule a program or script that runs daily (or however often you wish) and retrieves details of any new mailboxes created since it was last run.  The script can then send a welcome email to these new accounts.  This technique works well with both Exchange Online and Exchange on-premises.

Obtaining the list of mailboxes created since a particular date can be done with `Get-Mailbox -Filter WhenMailboxCreated -gt <date>`.  Sending the welcome message can be achieved in several ways.  For Exchange Online the easiest option is to use Microsoft Graph, while for on-premises, the easiest options are SMTP or EWS (both could also work with EXO).

I have written a sample script that shows how to implement the above process and can send a message using SMTP or Graph.  The script shows how a message template (in HTML) can be used with custom fields so that the message can be personalised (e.g. using the recipient's first name), and an embedded image attached.

You can find the script and the sample files [here](Code/).

## Sample Walk-through for Exchange Online

For Exchange Online, Graph is used to send the welcome message and so an application first of all needs to be registered in Azure Active Directory.

1. [Register the application](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app).  You will need to grant and consent to Mail.Send permission, and then make a note of the tenant Id, application Id, and secret key to provide as parameters to the script.

2. Prepare the email template files.  I've included a sample welcome.html with an embedded image to show how this can be done.  When the script attaches the image to the message, it will add a Content-Id as the image name so that it can be referenced in the HTML (e.g. `<img width=129 height=43 src="cid:image001.png" alt="MSFT_logo" v:shapes="Picture_x0020_1">`, as in the sample).  I've also included a single custom field `#UserFirstName#` that the script will replace with the user's first name (as returned by `Get-User`).

3. Create a test mailbox.  The first time the script is run, it will check for any mailboxes created that day (since 00:00:00).  When it is run on subsequent occasions, it will check for mailboxes created since it was last run (it stores the last run time in a config file).  Note that if any mailboxes were created other than the test mailbox, the script will also send emails to those, so it is best to not test this in a production environment (as with all sample scripts).

4. Run the script using appropriate parameters: `.\Send-WelcomeEmails.ps1 -MessageSubject "Test Welcome" -Office365 -MessageSender welcome@domain.com -AppId "e1e6613d-0ca0-43b2-b702-a327b110ddc0" -TenantId "fc69f6a8-90cd-4047-977d-0c768925b8ec" -AppSecretKey "1234"`

If the script is run from a PowerShell console where the Exchange Online module is not available, it will attempt to automatically connect.  This will trigger an auth prompt.  This is an example of a successful script run where one new mailbox was found:

![PowerShell Console screenshot showing successful script execution against Exchange Online](Images/EXOPSTest.png?raw=true)

And here is the message that was received:

![Welcome Message displayed in OWA](Images/EXOSampleMessage.png?raw=true)

## Sample Walk-through for Exchange 2019 (on-premises)

Graph is not available for Exchange 2019, so SMTP (via System.Net.Mail) is used to send the email in this environment.

1. Prepare the email template files.  I've included a sample welcome.html with an embedded image to show how this can be done.  When the script attaches the image to the message, it will add a Content-Id as the image name so that it can be referenced in the HTML (e.g. `<img width=129 height=43 src="cid:image001.png" alt="MSFT_logo" v:shapes="Picture_x0020_1">`, as in the sample).  I've also included a single custom field `#UserFirstName#` that the script will replace with the user's first name (as returned by `Get-User`).

2. Create a test mailbox.  The first time the script is run, it will check for any mailboxes created that day (since 00:00:00).  When it is run on subsequent occasions, it will check for mailboxes created since it was last run (it stores the last run time in a config file).  Note that if any mailboxes were created other than the test mailbox, the script will also send emails to those, so it is best to not test this in a production environment (as with all sample scripts).

3. To successfully send the welcome message, the user creating the PowerShell session must have Send As rights for the mailbox being used to send the welcome message.  The script use default auth for SMTP (i.e. the logged on user's Windows credentials), and if the user does not have permission to send from the specified email address then it will fail.

4. Run the script using appropriate parameters: `.\Send-WelcomeEmails.ps1 -MessageSubject "Test Welcome" -MessageSender welcome@e19.local -SMTPServer e1.e19.local -PowerShellUrl "http://e1.e19.local/PowerShell" -ExchangeCredential (Get-Credential)`.  Note that this example will prompt for credentials to connect to Exchange PowerShell.  In the screenshot below I use an alternative of assigning the credentials to a variable before calling the script.

![PowerShell Console screenshot showing successful script execution against Exchange 2019](Images/E19PSTest.png?raw=true)

![Welcome Message displayed in OWA](Images/E19SampleMessage.png?raw=true)


## Notes

* The script would need modification to run unattended (due to the current Exchange PowerShell implementation).  This is commented in the script with a link to the appropriate docs.
* This is a sample script and should only be used for testing purposes or as a basis for further development.  It is not resilient at all (e.g. if the message send fails, it isn't reattempted).
* The script only supports .png and .jpg attachments, and they are always attached as inline images.


## Send-WelcomeMessage.ps1 Parameters

`-MessageSender`: Email address from which the welcome message will be sent.

`-MessageTemplate`: Filename of the welcome message (message is required to be in HTML).

`-MessageSubject`: Subject of the welcome message.

`-MessageAttachments`: A list of the files that must be attached to the message (e.g. images, etc.).

`-SMTPServer`: SMTP server that will be used to send the welcome message (if not provided, Graph will be used to send).

`-SMTPCredential`: Credentials that will be used to authenticate with the SMTP server.  If not specified, no auth (anonymous) is used.

`-PowerShellUrl`: Exchange PowerShell URL (so that we can connect to Exchange).  If using Exchange Online, use -Office365 switch instead.

`-ExchangeCredential`: Credentials used to authenticate with Exchange (only required when PowerShellUrl specified).

`-Office365`: If set, connect to Office 365 PowerShell as necessary (which will be if a session is not already available).

`-TenantId`: Tenant Id.  Required for Exchange Online.

`-AppId`: Application Id (from application registered in Azure AD).  Required for Exchange Online.

`-AppSecretKey`: Application secret key (from Azure AD).  Required for Exchange Online.

`-ConfigFolder`: Folder in which the configuration (including welcome message and any related files) is located.  If missing, current folder is used.

`-LogFile`: Logs script activity to the specified file.
