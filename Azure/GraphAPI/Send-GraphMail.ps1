#region Module Imports
Import-Module -Name Microsoft.Graph.Authentication
Import-Module -Name Microsoft.Graph.Users.Actions
#endregion

<#
.Synopsis

   This function enables you to send mails through an existing App Registration.

.DESCRIPTION

   Please create an App Registration with the following API Application permissions:
        - User.Read.All
        - Mail.Send

    Authenticate to Microsoft Graph before you call the function.

    Import the following Modules:
        Import-Module -Name Microsoft.Graph.Authentication
        Import-Module -Name Microsoft.Graph.Users.Actions

.EXAMPLE
   
   Send-GraphMail -mailFrom "example@example.com" -mailTo "anyone@anyone.com" -mailSubject "Example" -mailBodyAsHTML $true -mailBody "This is an example!" -mailAttachments "FILEPATH1","FILEPATH2"

.EXAMPLE
   
   Send-GraphMail -mailFrom "example@example.com" -mailTo "anyone@anyone.com" -mailSubject "Example" -mailBodyAsHTML $false -mailBody "This is an example!" -mailAttachments "FILEPATH1","FILEPATH2"

.EXAMPLE

   Send-GraphMail -mailFrom "example@example.com" -mailTo "anyone@anyone.com" -mailSubject "Example" -mailBodyAsHTML $true -mailBody "This is an example!"
#>
function Send-GraphMail
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true)] $mailFrom,
        [Parameter(Mandatory=$true)] $mailTo,
        [Parameter(Mandatory=$true)] $mailSubject,
        [Parameter(Mandatory=$true)] [boolean]$mailBodyAsHTML,
        [Parameter(Mandatory=$true)] $mailBody,
        [Parameter(Mandatory=$true)] [boolean]$saveToSentItems,
        $mailAttachments,
        $ccRecipients,
        $bccRecipients    
    )
    
    if(!(Get-MgContext))
    {
        Write-Host "No open Microsoft Graph session was found."
        return
    }


    if((!(Get-MgContext).Scopes.Contains('Mail.Send')) -and (!(Get-MgContext).Scopes.Contains('User.Read.All')))
    {
        Write-Host "Current sessions does not have the required API Permissions"
        return
    }

    switch($mailBodyAsHTML){
        $true {
            $type = "HTML"
        }
        $false {
            $type = "Text"
        }
    }

    $params = @{
        Message = @{
            Subject = $mailSubject
            Body = @{
                ContentType = $type
                Content     = $mailBody
            }
            Attachments = @(
            )

            ToRecipients  = @(
            )

            CcRecipients  = @(
            )
            
            BccRecipients = @(
            )
        }
        SaveToSentItems = $saveToSentItems
    }

    $totalBytes = 0

    If(!($null -eq $mailAttachments))
    {
        Foreach($attachment in $mailAttachments)
        {
            $totalBytes += (Get-Item -Path $attachment).Length
            $hashObject = @{ "@odata.type" = "#microsoft.graph.fileAttachment"
                             "name"        = $attachment.Substring($attachment.LastIndexOf('\') +1)
                             "contentBytes" = $([convert]::ToBase64String((Get-Content $attachment -Raw -Encoding Byte)))
            }
            $params.Message.Attachments += $hashObject
        }
    }

    Foreach($recipient in $mailTo)
    {
        $hashObject = @{ EmailAddress = @{
            Address = $recipient
            }
        }
        $params.Message.ToRecipients += $hashObject
    }

    if(!($null -eq $ccRecipients))
    {
        Foreach($recipient in $ccRecipients)
        {
            $hashObject = @{ EmailAddress = @{
                Address = $recipient
                }
            }
            $params.Message.CcRecipients += $hashObject
        }
    }

    if(!($null -eq $bccRecipients))
    {
        Foreach($recipient in $bccRecipients)
        {
            $hashObject = @{ EmailAddress = @{
                Address = $recipient
                }
            }
            $params.Message.BccRecipients += $hashObject
        }
    }

    If($totalBytes /1024 /1024 -gt 25)
    {
        Write-Host "WARNING: Attachment size(s) exceed the default message size limit in Exchange Online" -ForegroundColor DarkYellow
    }

    Send-MgUserMail -UserId $mailFrom -BodyParameter $params
}
