#region Start Variables
    $outputPath = 'C:\temp\'
    $completeList = New-Object System.Collections.Generic.List[PSObject]
    $expiringList = New-Object System.Collections.Generic.List[PSObject]
#endregion
#region Module Check
    $installedModules          = Get-Module -ListAvailable
    $graphAuthenticationModule = $installedModules | ? { $_.Name -eq 'Microsoft.Graph.Authentication' }
    $graphApplicationsModule   = $installedModules | ? { $_.Name -eq 'Microsoft.Graph.Applications' }
    $excelImportModule         = $installedModules | ? { $_.Name -eq 'ImportExcel' }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if($null -eq $graphAuthenticationModule)
    {
        Install-Module -Name Microsoft.Graph.Authentication -RequiredVersion 2.0.0-preview2 -AllowPrerelease -Force -AllowClobber -Confirm:$false
    }

    if($null -eq $graphApplicationsModule)
    {
        Install-Module -Name Microsoft.Graph.Applications -RequiredVersion 2.0.0-preview2 -AllowPrerelease -Force -AllowClobber -Confirm:$false
    }

    if($null -eq $excelImportModule)
    {
        Install-Module -Name ImportExcel -Force -AllowClobber -Confirm:$false
    }

    Import-Module -Name Microsoft.Graph.Authentication
    Import-Module -Name Microsoft.Graph.Applications
    Import-Module -Name ImportExcel
#endregion
#region Graph Authentication
    $clientID = ''
    $tenantID = ''
    $secretID = Get-Content -Path "C:\scripts\Azure\encryptedSecret.txt" | ConvertTo-SecureString -Force

    $ClientSecretCredential = New-object -TypeName System.Management.Automation.PSCredential -ArgumentList $clientID, $secretID
    Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential
#endregion



$collectedApplications = Get-MgApplication -All

Foreach($application in $collectedApplications)
{
    If($application.PasswordCredentials.Count -ne 0)
    {
        Foreach($secret in $application.PasswordCredentials)
        {
            $expiresInDays = ($secret.EndDateTime - (Get-Date)).Days
            $isExpired = $expiresInDays -lt 1
            $hashObject = [ordered]@{ ApplicationName = $application.DisplayName; Type = 'Secret'; 
                SecretName = $secret.DisplayName; SecretStartTime = $secret.StartDateTime; SecretEndTime = $secret.EndDateTime; ExpiresInDays = $expiresInDays;
                Expired = $isExpired }
            $outputObject = New-Object PSObject -Property $hashObject
            $completeList.Add($outputObject)
            if($expiresInDays -lt 30)
            {
                $expiringList.Add($outputObject)
            }
        }
    }
    If($application.KeyCredentials.Count -ne 0)
    {
        Foreach($certificate in $application.KeyCredentials)
        {
            $expiresInDays = ($certificate.EndDateTime - (Get-Date)).Days
            $isExpired = $expiresInDays -lt 1
            $hashObject = [ordered]@{ ApplicationName = $application.DisplayName; Type = 'Certificate'; 
                SecretName = $certificate.DisplayName; SecretStartTime = $certificate.StartDateTime; SecretEndTime = $certificate.EndDateTime; ExpiresInDays = $expiresInDays;
                Expired = $isExpired}
            $outputObject = New-Object PSObject -Property $hashObject
            $completeList.Add($outputObject)
            if($expiresInDays -lt 30)
            {
                $expiringList.Add($outputObject)
            }
        }
    }
}

$completeList | Export-Excel -Path "$($outputPath)app_registration_complete.xlsx" -FreezeTopRow -TitleBold -AutoFilter -AutoSize
$expiringList | Export-Excel -Path "$($outputPath)app_registration_expiring.xlsx" -FreezeTopRow -TitleBold -AutoFilter -AutoSize

$mailTo = ''
$mailFrom = ''
$mailSubject = ''
$mailSMTP = ''

Send-MailMessage -SmtpServer $mailSMTP -From $mailFrom -To $To -Subject "$($mailSubject) - $(Get-Date -Format "dd-MM-yyyy")" -Body "$($expiringList.Count) secrets/certificates have less than 30 days before they expire." -Attachments "$($outputPath)app_registration_complete.xlsx", "$($outputPath)app_registration_expiring.xlsx"

Disconnect-MgGraph