#
# Send-WelcomeEmails.ps1
#
# By David Barrett, Microsoft Ltd. 2022. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
 

<#
.SYNOPSIS
Send welcome emails to newly created mailboxes.

.DESCRIPTION
This script demonstrates how to check for new mailboxes (created since the script was last run) and if any are found, send a welcome email (which can be personalised).  This can work for both Exchange Online and on-premises.

.EXAMPLE
.\Send-WelcomeEmails.ps1 -MessageSubject "Test Welcome" -Office365 -MessageSender WelcomeBot@domain.com -AppId "e1e6613d-0ca0-43b2-b702-a327b110ddc0" -TenantId "fc69f6a8-90cd-4047-977d-0c768925b8ec" -AppSecretKey "" -Verbose

This will run the script against Exchange Online and send new emails using Graph sendMail (an application must be registered in Azure AD with the appropriate permissions).  Note that the script uses application permissions when
sending the message, delegate permissions won't work.

#>

param (
	[Parameter(Mandatory=$True,HelpMessage="Email address from which the welcome message will be sent.")]
	[ValidateNotNullOrEmpty()]
	[string]$MessageSender,

	[Parameter(Mandatory=$False,HelpMessage="Filename of the welcome message (message is required to be in HTML).")]
	[ValidateNotNullOrEmpty()]
	[string]$MessageTemplate = "Welcome.html",

	[Parameter(Mandatory=$False,HelpMessage="Subject of the welcome message.")]
	[ValidateNotNullOrEmpty()]
	[string]$MessageSubject = "Welcome",

	[Parameter(Mandatory=$False,HelpMessage="A list of the files that must be attached to the message (e.g. images, etc.).")]
	$MessageAttachments = @("image001.png"),

	[Parameter(Mandatory=$False,HelpMessage="SMTP server that will be used to send the welcome message (if not provided, Graph will be used to send).")]
	[ValidateNotNullOrEmpty()]
	[string]$SMTPServer = "",

	[Parameter(Mandatory=$False,HelpMessage="Credentials that will be used to authenticate with the SMTP server.  If not specified, default auth (current user) will be attempted.")]
	[ValidateNotNullOrEmpty()]
	[PSCredential]$SMTPCredential,

	[Parameter(Mandatory=$False,HelpMessage="Exchange PowerShell URL (so that we can connect to Exchange).  If using Exchange Online, use -Office365 switch instead.")]
	[ValidateNotNullOrEmpty()]
	[string]$PowerShellUrl,

	[Parameter(Mandatory=$False,HelpMessage="Credentials used to authenticate with Exchange (only required when PowerShellUrl specified).")]
    [PSCredential]$ExchangeCredential,

	[Parameter(Mandatory=$False,HelpMessage="If set, connect to Office 365 PowerShell as necessary (which will be if a session is not already available).")]
	[switch]$Office365,

	[Parameter(Mandatory=$False,HelpMessage="Tenant Id.  Required for Exchange Online.")]
	[ValidateNotNullOrEmpty()]
	[string]$TenantId,

	[Parameter(Mandatory=$False,HelpMessage="Application Id (from application registered in Azure AD).  Required for Exchange Online.")]
	[ValidateNotNullOrEmpty()]
	[string]$AppId,

	[Parameter(Mandatory=$False,HelpMessage="Application secret key (from Azure AD).  Required for Exchange Online.")]
	[ValidateNotNullOrEmpty()]
	[string]$AppSecretKey,

	[Parameter(Mandatory=$False,HelpMessage="Folder in which the configuration (including welcome message and any related files) is located.  If missing, current folder is used.")]
	[ValidateNotNullOrEmpty()]
	[string]$ConfigFolder,

	[Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified.")]
	[string]$LogFile = ""
)

$script:ScriptVersion = "1.0.2"

# Default config.  First time the script is run it will check for any new mailboxes created today.  On subsequent runs, it checks from the last check date (which is saved in the config file).
$script:config = @{ "LastDateCheck" = [DateTime]::Today }

Function LogToFile([string]$Details)
{
	if ( [String]::IsNullOrEmpty($LogFile) ) { return }
    $logInfo = "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details"
    if ($FastFileLogging)
    {
        if (!$script:logFileStream)
        {
            # Open a filestream to write to our log
            $script:logFileStream = New-Object IO.FileStream($LogFile, ([System.IO.FileMode]::Append), ([IO.FileAccess]::Write), ([IO.FileShare]::Read) )
            ReportError "Opening log file"
        }
        if ($script:logFileStream)
        {
            $streamWriter = New-Object System.IO.StreamWriter($script:logFileStream)
            $streamWriter.WriteLine($logInfo)
            $streamWriter.Dispose()
            if ( ErrorReported("Writing log file") )
            {
                $FastFileLogging = $false
            }
            else
            {
                return
            }
        }
    }
	$logInfo | Out-File $LogFile -Append
}

Function Log([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
    Write-Host $Details -ForegroundColor $Colour
    LogToFile $Details
}
Log "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) started" Green

Function LogVerbose([string]$Details)
{
    Write-Verbose $Details
    if ($VerbosePreference -eq "SilentlyContinue") { return }
    LogToFile $Details
}

Function LogDebug([string]$Details)
{
    Write-Debug $Details
    if ($DebugPreference -eq "SilentlyContinue") { return }
    LogToFile $Details
}

$script:LastError = $Error[0]
Function ErrorReported($Context)
{
    # Check for any error, and return the result ($true means a new error has been detected)

    # We check for errors using $Error variable, as try...catch isn't reliable when remoting
    if ([String]::IsNullOrEmpty($Error[0])) { return $false }

    # We have an error, have we already reported it?
    if ($Error[0] -eq $script:LastError) { return $false }

    # New error, so log it and return $true
    $script:LastError = $Error[0]
    if ($Context)
    {
        Log "Error ($Context): $($Error[0])" Red
    }
    else
    {
        Log "Error: $($Error[0])" Red
    }
    return $true
}

Function ReportError($Context)
{
    # Reports error without returning the result
    ErrorReported $Context | Out-Null
}

Function CmdletsAvailable()
{
    # Check if the specified cmdlets are available in the current PowerShell session
    param (
        $RequiredCmdlets,
        $Silent = $False,
        $PSSession = $null
    )

    $cmdletsAvailable = $True
    foreach ($cmdlet in $RequiredCmdlets)
    {
        $cmdletExists = $false
        if ($PSSession)
        {
            $cmdletExists = $(Invoke-Command -Session $PSSession -ScriptBlock { Get-Command $Using:cmdlet -ErrorAction Ignore })
        }
        else
        {
            $cmdletExists = $(Get-Command $cmdlet -ErrorAction Ignore)
        }
        if (!$cmdletExists)
        {
            if (!$Silent) { Log "Required cmdlet $cmdlet is not available" Red }
            $cmdletsAvailable = $False
        }
    }

    return $cmdletsAvailable
}

Function CheckEnvironment($PowerShellUrl)
{
    # Check that we have the required Exchange session available

    # We need two cmdlets, so we check specifically that these are available
    if ($(CmdletsAvailable @("Get-User", "Get-Mailbox") $True )) {
        return
    }

    # Cmdlets are not available, so attempt to connect to Exchange runspace

    if ($Office365)
    {
        # For Office 365, we connect using the EXO v2 management module (though we still use v1 cmdlets so we can use the same code for on-prem and EXO)
        Import-Module ExchangeOnlineManagement

        # Note that the following will trigger an interactive log-on.  For production implementation, this script would want to be scheduled and also updated
        # so as not to require user interaction here.  See https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps
        Connect-ExchangeOnline
        if ( !$(CmdletsAvailable @("Get-User", "Get-Mailbox") $True ) )
        {
            # Failed to connect to Exchange Online
            Log "Connection to Exchange Online failed, cannot continue" Red
            exit
        }
        return
    }

    if ( ![String]::IsNullOrEmpty($PowerShellUrl) )
    {
        # Try to connect and import a session
        Log "Connecting to Exchange using PowerShell Url: $PowerShellUrl"
        $script:ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $PowerShellUrl -Credential $ExchangeCredential -Authentication Kerberos  -WarningAction 'SilentlyContinue'
        ReportError "New-PSSession"

        # If we don't have all the cmdlets available, we can't go any further
        if ( !$(CmdletsAvailable @("Get-User", "Get-Mailbox") $False $script:ExchangeSession) )
        {
            Log "Required Exchange cmdlet(s) are missing, cannot continue" Red
            exit
        }

        # Import the session
        Import-PSSession $script:ExchangeSession -AllowClobber -WarningAction 'SilentlyContinue' -CommandType All -DisableNameChecking
        ReportError "Import-PSSession"
    }
}

Function SendWithSystemNetMail
{
    param (
        $messageBody,
        $recipient
    )

	# Create the objects we need to create the message and send the mail
    LogVerbose "Creating welcome message for $($recipient)"
    $message = New-Object System.Net.Mail.MailMessage
    LogVerbose "Using SMTP server: $SMTPServer"
    $smtpClient = New-Object System.Net.Mail.SmtpClient($SMTPServer)

    if ($SMTPCredential)
    {
        LogVerbose "Using authenticated SMTP"
        $smtpClient.Credentials = $SMTPCredential
    }


	# Create the HTML view for this message
	$view = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($messageBody, $null, "text/html")

	# Add any linked resources (e.g. images)
	foreach ($linkedResource in $MessageAttachments)
	{
        $image = $null
        $imageFile = Get-Item $ConfigFolder$linkedResource
		$image = New-Object System.Net.Mail.LinkedResource("$($imageFile.VersionInfo.FileName)")
        if ($image -ne $null)
        {
		    $image.ContentId = $linkedResource
		    $image.ContentType = "image/png"
		    $view.LinkedResources.Add($image)
            LogVerbose "Added image to message: $linkedResource"
        }
        else
        {
            Log "FAILED to add image to message: $linkedResource" Red
        }
	}

	# Create the message
	$message.From = $MessageSender
	$message.To.Add($recipient)
	$message.Subject = $MessageSubject
	$message.AlternateViews.Add($view)
	$message.IsBodyHtml = $true

	# Send the message
	
    try
    {
	    $smtpClient.Send($message)
        Log "Successfully sent welcome email to $recipient"
    }
    catch
    {
        Log "Error occurred when sending welcome email to $($recipient): $($Error[0])" Red
        return $false
    }
}

Function GetToken
{
    # Obtain an OAuth token to use Graph
    # For simplicity, this does not cover token renewal (that would only be required if the script took longer than an hour to run, which is unlikely anyway)

    if ($script:authHeader) {
        return
    }

    # Acquire token for application permissions
    $body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$AppId;client_secret=$AppSecretKey}
    try
    {
        $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
    }
    catch
    {
        Write-Host "Failed to obtain OAuth token" -ForegroundColor Red
        exit # Failed to obtain a token
    }
    $script:authHeader = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
    LogVerbose "Successfully obtained OAuth token" Green
}

Function SendWithGraph
{
    # Send the welcome message using Graph sendMail
    # https://docs.microsoft.com/en-us/graph/api/user-sendmail
    param (
        $messageBody,
        $recipient
    )

    GetToken
    $graphUrl = "https://graph.microsoft.com/v1.0/users/$MessageSender/"
    $sendMailUrl = "https://graph.microsoft.com/v1.0/users/$MessageSender/sendMail"

    $attachmentJson = ""    
    if ($MessageAttachments.Count -gt 0)
    {
        $attachmentJson = ",
        ""attachments"": ["
        $attachmentsAdded = 0

        foreach ($AttachPicture in $MessageAttachments)
        {
            if (-not [String]::IsNullOrEmpty($AttachPicture))
            {
                # Add picture attachments (currently only jpg and png are supported)
                $pictureType = $AttachPicture.Substring($AttachPicture.Length-3).ToLower()
                if ($pictureType -ne "jpg" -and $pictureType -ne "png")
                {
                    Write-Host "Attachment must be jpg or png." -ForegroundColor Red
                }
                else
                {
                    $pictureFile = Get-Item $AttachPicture
                    if (!$pictureFile)
                    {
                        Write-Host "Failed to read picture: $AttachPicture" -ForegroundColor Red
                    }
                    else
                    {
                        # Read the byte data for the picture
                        LogVerbose "Adding attachment: $($pictureFile.VersionInfo.FileName)"
                        $fileStream = New-Object -TypeName System.IO.FileStream -ArgumentList ($pictureFile.VersionInfo.FileName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
                        $fileReader = New-Object -TypeName System.IO.BinaryReader -ArgumentList $fileStream
                        if (!$fileReader) { exit }

                        $pictureBytes = $fileReader.ReadBytes($pictureFile.Length)
                        $fileReader.Dispose()
                        $fileStream.Dispose()

                        # Convert the data into a Base64 string
                        $attachBytesBase64 = [Convert]::ToBase64String($pictureBytes)

                        # Add the attachment JSON
                        if ($attachmentsAdded -gt 0) {
                            $attachmentJson = "$attachmentJson,"
                        }

                        $attachmentJson = "$attachmentJson
            {
            ""@odata.type"": ""#microsoft.graph.fileAttachment"",
            ""contentBytes"": ""$attachBytesBase64"",
            ""contentId"": ""$($pictureFile.Name)"",
            ""contentType"": ""image/$pictureType"",
            ""isInline"": true,
            ""name"": ""$($pictureFile.Name)"",            
            }"
                        $attachmentsAdded++
                    }
                }
            }
        }
        $attachmentJson = "$attachmentJson
        ]"
    }

    $bodyJson = """body"":{
            ""contentType"":""HTML"",
            ""content"":$(ConvertTo-Json $messageBody)
        }"

    $sendMessageJson = "{
    ""message"": {
        ""subject"":""$MessageSubject"",
        ""importance"":""Low"",
        $bodyJson,
        ""toRecipients"":[
            {
                ""emailAddress"":{
                    ""address"":""$recipient""
                }
            }
        ]"

    $sendMessageJson = "$sendMessageJson$attachmentJson"

    $sendMessageJson = "$sendMessageJson
    },
    ""saveToSentItems"": ""true""    
}"

    $global:messageDebug = $sendMessageJson 

    try
    {
        LogVerbose "Sending request to: $sendMailUrl"
        $sendMessageResults = Invoke-RestMethod -Method Post -Uri $sendMailUrl -Headers $authHeader -Body $sendMessageJson -ContentType "application/json"
        Log "Successfully submitted message for $recipient"
    }
    catch
    {
        ReportError "Failed to send message to $recipient"
    }

    # The maximum message rate for Exchange Online is 30 messages per minute: https://docs.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits?msclkid=f574c61bd06d11ec9fc5c17d2fe87fb0#sending-limits-1
    # So we add a two second sleep here to ensure we never exceed that.  We don't check for other limits (they are unlikely to be an issue).
    Start-Sleep -Milliseconds 2000
}

Function CreateMessageBody
{
    # Create the personalised welcome message by reading the message template and replacing any custom fields
    param (
        $mailbox,
        $user
    )

	# Read the HTML template
    if ([string]::IsNullOrEmpty($ConfigFolder))
    {
	    $messageBody = [string](Get-Content $MessageTemplate)
    }
    else
    {
        $messageBody = [string](Get-Content $ConfigFolder$MessageTemplate)
    }

    # Replace any fields within the template.  Further fields can be added here as required
	$messageBody = $messageBody.Replace("#UserFirstName#", $user.FirstName) # #UserFirstName# : user's first name

    return $messageBody
}

Function SendEmail
{
	param (
		$mailbox
	)
		
	# Mailbox does not have first name property, so we need to get the user account
    $usr = $null
	$usr = Get-User $mailbox.PrimarySmtpAddress
    if ($mailbox -eq $null)
    {
        Log "Unable to locate user (Get-User): $($mailbox.PrimarySmtpAddress)" Red
        return $False
    }

	# Get the HTML body for the message (this can be customised per user)
	$messageBody = CreateMessageBody $mailbox $usr

    # Send the message
    if ([String]::IsNullOrEmpty($SMTPServer))
    {
        SendWithGraph $messageBody $mailbox.PrimarySmtpAddress
    }
    else
    {
        SendWithSystemNetMail $messageBody $mailbox.PrimarySmtpAddress
    }
    
    return $true
}



# Script starts here

# We need an Exchange Management Shell, so check we have one (or can get one)
CheckEnvironment $PowerShellUrl

# Check if we have a saved configuration file, and if so, load it
if (Test-Path -Path "$ConfigFolder$($MyInvocation.MyCommand.Name).config") {
    LogVerbose "Restoring configuration from file: $ConfigFolder$($MyInvocation.MyCommand.Name).config"
    $json = Get-Content "$ConfigFolder$($MyInvocation.MyCommand.Name).config" | Out-String
    (ConvertFrom-Json $json) | Foreach {
        LogVerbose "$($_.Name): $($_.Value)"
        if ($script:config.ContainsKey($_.Name)) {
            $script:config[$_.Name] = $_.Value
        }
    }
}

# Retrieve the list of mailboxes that have been created since we last checked
$filter = "WhenMailboxCreated -gt '$($script:config["LastDateCheck"].ToLocalTime())'"
$script:config["LastDateCheck"] = [DateTime]::UtcNow
LogVerbose "Get-Mailbox -Filter $filter"
$newMailboxes = Get-Mailbox -Filter $filter

if (!$newMailboxes)
{
    Log "No new mailboxes found."
}
else
{
    # Send a welcome email to each of the new mailboxes
    foreach ($newMailbox in $newMailboxes)
    {
        SendEmail $newMailbox | out-null
    }
}

# Save our updated config
$script:config.GetEnumerator() | ConvertTo-Json | Out-File "$ConfigFolder$($MyInvocation.MyCommand.Name).config"

Log "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) finished" Green