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
#region Connections
Connect-MgGraph
Connect-ExchangeOnline
#endregion
#region Microsoft.Graph information collection
$appRegistrations = Get-MgApplication -All
$appAccessPolicies = (Get-ApplicationAccessPolicy).Identity
#endregion
#region Policy Scopes collection
$appAccessResponse = Invoke-WebRequest -Uri "https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access"
$appAccessPolicyScopes = @() 

[regex]::Matches($appAccessResponse, '<li><code>(.*?)</code></li>') | % {
    $match = $_.Value
    $match = $match.Replace('<li><code>', '')
    $match = $match.Replace('</code></li>', '')
    $appAccessPolicyScopes += $match
}
#endregion
#region Output
$outputPath = "$($env:TEMP)\Unsafe_AppRegistrations.xlsx"
$outputList = New-Object System.Collections.Generic.List[PSObject]
#endregion

Foreach($appRegistration in $appRegistrations)
{
    Foreach($resource in $appRegistration.RequiredResourceAccess)
    {
        $resourceID = $resource.ResourceAppID
        $resourceSP = Get-MgServicePrincipal -Filter "AppId eq '$resourceID'"
        Foreach($permission in $resource.ResourceAccess)
        {
            if($permission.Type -eq 'Role')
            {
                $appRoleInfo = $resourceSP.AppRoles | Where-Object { $_.ID -eq $permission.ID }
                if(($appAccessPolicyScopes -contains $appRoleInfo.Value) -and ($appRoleInfo.Origin -eq 'Application'))
                {
                    $result = $null
                    $result = $appAccessPolicies -like "*$($appRegistration.AppID)*"
                    if("" -eq $result)
                    {
                        
                        $hashObject = @{
                            AppName        = $appRegistration.DisplayName;
                            AppID          = $appRegistration.AppId;
                            ObjectID       = $appRegistration.Id;
                            PermissionType = "Application";
                            Scope          = $appRoleInfo.Value;
                        }
                        $outputObject = New-Object -TypeName PSObject -Property $hashObject
                        $outputList.Add($outputObject)
                    }
                }
            } 
        }
    }
}

$outputList | Export-Excel -Path $outputPath -AutoFilter -AutoSize -FreezeTopRow