#https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/replace-an-expiring-client-secret-in-a-sharepoint-add-in

$appPrincipal = Get-MgServicePrincipal -ServicePrincipalId 1c55990d-1270-4fc4-ad30-32f58ffcd611 
$params = @{
    PasswordCredential = @{
        DisplayName = "NewSecret" # Replace with a friendly name.
    }
}
$result = Add-MgServicePrincipalPassword -ServicePrincipalId $appPrincipal.Id -BodyParameter $params    # Update the secret
$base64Secret = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($result.SecretText)) # Convert to base64 string.
$app = Get-MgServicePrincipal -ServicePrincipalId $appPrincipal.Id # get existing app information
$existingKeyCredentials = $app.KeyCredentials # read existing credentials
$dtStart = [System.DateTime]::Now # Start date
$dtEnd = $dtStart.AddYears(2) # End date (equals to secret end date)
$keyCredentials = @( # construct keys
    @{
        Type = "Symmetric"
        Usage = "Verify"
        Key = [System.Text.Encoding]::ASCII.GetBytes($result.SecretText)
        StartDateTime = $dtStart
        EndDateTIme = $dtEnd
    },
    @{
        type = "Symmetric"
        usage = "Sign"
        key = [System.Text.Encoding]::ASCII.GetBytes($result.SecretText)
        StartDateTime = $dtStart
        EndDateTIme = $dtEnd
    }
) + $existingKeyCredentials # combine with existing
Update-MgServicePrincipal -ServicePrincipalId $appPrincipal.Id -KeyCredentials $keyCredentials # Update keys
$base64Secret # Print base64 secret
$result.EndDateTime # Print the end date.