function Get-AuthenticationToken
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)] $tenantID,
        [Parameter(Mandatory=$true)] $appID,
        [Parameter(Mandatory=$true)] $appSecret,
        [Parameter(Mandatory=$true)] $apiUri)

        $OAuthUri = "https://login.microsoftonline.com/$tenantID/oauth2/token"

        $authBody = [Ordered] @{
            resource = "$apiUri"
            client_id = "$appID"
            client_secret = "$appSecret"
            grant_type = 'client_credentials'
        }

        Try{
            $authResponse = Invoke-RestMethod -Method Post -Uri $OAuthUri -Body $authBody -ErrorAction Stop
            $authToken    = $authResponse.access_token

            return $authToken
        } Catch {
            Write-Host $error[0].Exception
        }        
}