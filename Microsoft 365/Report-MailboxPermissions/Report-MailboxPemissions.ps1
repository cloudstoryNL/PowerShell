<#
Create an App Registration with the following API Permissions:
    Office 365 Exchange Online:
    - Exchange.ManageAsApp
    - Mail.ReadWrite
    - User.Read.All

Create a certificate on your management server under the service account that you will use:

New-SelfSignedCertificate -DnsName "yourtenant.onmicrosoft.com" -CertStoreLocation "cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(2) -KeySpec KeyExchange | Export-PfxCertificate -FilePath mycert.pfx -Password (Read-Host -AsSecureString -Prompt "Enter a Password for PFX File")
(get-childitem Cert:\CurrentUser\my) | Where-Object {$_.Subject -eq "cn=yourtenant.onmicrosoft.com"} | Export-Certificate -FilePath mycert.cer

Upload the certificate file to the newly created App Registration
#>

#region Setting up Modules
$installedModules = Get-Module -ListAvailable
if(!($installedModules | ? { $_.Name -eq 'ExchangeOnlineManagement' }))
{
    Install-Module -Name ExchangeOnlineManagement
}
if(!($installedModules | ? { $_.Name -eq 'ImportExcel' }))
{
    Install-Module -Name ImportExcel
}
#endregion
#region Module Imports
Import-Module -Name ImportExcel
Import-Module -Name ExchangeOnlineManagement
#endregion


function Connect_ExchangeOnline
{
    $certficateThumbprint = ""
    $appID = ""
    $organizationURL = "yourtenant.onmicrosoft.com"
    Connect-ExchangeOnline -AppId $appID -Organization $organizationURL -CertificateThumbprint $certficateThumbprint -ShowBanner:$false
}

# Calling a function to connect to Exchange Online using the App Registration
Connect_ExchangeOnline

# Deleting Mailbox Permissions exports that are older than 2 days.
Get-ChildItem -Path "$($env:TEMP)\" -Filter "*Mailbox Permissions.xlsx" | ? { $_.CreationTime -lt (Get-Date).AddDays(-2) } | Remove-Item -Force

# Collecting all mailboxes
$mailboxCollection = Get-Mailbox -ResultSize Unlimited | Select-Object Name, PrimarySMTPAddress, isShared

$outputList = New-Object System.Collections.Generic.List[PSObject]

$counter = 0
Foreach($mailbox in $mailboxCollection)
{
    'Processing {0} out of {1}' -f $counter, $mailboxCollection.Count | Write-Host -ForegroundColor Yellow
    $counter++
    $mailboxPermissions = Get-MailboxPermission -Identity $mailbox.PrimarySmtpAddress | ? { ($_.User -notmatch "NT AUTHORITY") -and ($_.User -ne $mailbox.PrimarySmtpAddress) -and ($_.User -notlike 'S-1-5-21-*')} 
    $sendasPermissions = Get-RecipientPermission -Identity $mailbox.PrimarySmtpAddress | ? { ($_.Trustee -notmatch "NT AUTHORITY") -and ($_.Trustee -ne $mailbox.PrimarySmtpAddress) -and ($_.Trustee -notlike 'S-1-5-21-*')}  
    
    Foreach($permission in $mailboxPermissions) {
        $accessRights = ""
        $permission.AccessRights | % { $accessRights += $_ }
        $hashObject = @{PrimarySMTP = $mailbox.PrimarySMTPAddress ; DelegatedUser = $permission.User; AccessRights = $accessRights; SharedMailbox = $mailbox.IsShared}
        $outputObject = New-Object -TypeName PSObject -Property $hashObject
        $outputList.Add($outputObject)
    }

    Foreach($permission in $sendasPermissions)
    {
        $accessRights = ""
        $permission.AccessRights | % { $accessRights += $_ }
        $hashObject = @{PrimarySMTP = $mailbox.PrimarySMTPAddress ; DelegatedUser = $permission.Trustee; AccessRights = $accessRights; SharedMailbox = $mailbox.IsShared}
        $outputObject = New-Object -TypeName PSObject -Property $hashObject
        $outputList.Add($outputObject)
    }
}

$outputFile = "$($env:TEMP)\{0} - Mailbox Permissions.xlsx" -f (Get-Date -Format "dd-MM-yyyy HH-mm")

$outputList | Export-Excel -Path $outputFile -AutoSize -AutoFilter -FreezeTopRow