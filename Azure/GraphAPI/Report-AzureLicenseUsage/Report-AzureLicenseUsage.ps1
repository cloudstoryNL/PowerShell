#region Setting up Modules
$installedModules = Get-Module -ListAvailable
if(!($installedModules | ? { $_.Name -eq 'Microsoft.Graph.Authentication' }))
{
    Install-Module -Name Microsoft.Graph.Authentication
}
if(!($installedModules | ? { $_.Name -eq 'Microsoft.Graph.Users' }))
{
    Install-Module -Name Microsoft.Graph.Users
}
if(!($installedModules | ? { $_.Name -eq 'ImportExcel' }))
{
    Install-Module -Name ImportExcel
}
#endregion
#region Module Imports
Import-Module -Name Microsoft.Graph.Authentication
Import-Module -Name Microsoft.Graph.Users
Import-Module -Name ImportExcel
#endregion
#region Collecting Azure Information
# Collecting all the SKU's from the licenses overview in Azure
$consumedLicenses = Get-MgSubscribedSku | ? { ($_.ConsumedUnits -gt 1) -and ($_.CapabilityStatus -eq 'Enabled') }
# Collecting all the users from EntraID
$userCollection = Get-MgUser -All
#endregion
#region licensing-service-plan-reference
if(!(Test-Path -Path "$($env:TEMP)\Azure-Licensing-Plans.csv"))
{
    $servicePlanPage = Invoke-WebRequest -Uri "https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference" -Method Get
    $servicePlanDownloadURI = [regex]::Matches($servicePlanPage, '(https://.*?\.csv)')
    Invoke-WebRequest -Uri $servicePlanDownloadURI[0].Value -OutFile "$($env:TEMP)\Azure-Licensing-Plans.csv"
}
$licensingPlans = Import-Csv -Path "$($env:TEMP)\Azure-Licensing-Plans.csv"
#endregion
#region Dynamic Variables
$licenseAssignments = @()
$licenseOverview    = @()
#endregion
#region Output Locations
$outputBase = "$($env:TEMP)\"
$outputLicenseAssignments = '{0}\License Assignment Overview.xlsx' -f $outputBase
$outputLicenseOverview    = '{0}\License Usage Overview.xlsx' -f $outputBase
#endregion


Foreach($user in $userCollection)
{
    $userLicense = Get-MgUserLicenseDetail -UserId $user.UserPrincipalname
    $userType    = "Internal"

    if ($user.UserPrincipalName -like "*#EXT#*")
	{
        $companyName = $user.UserPrincipalName.Substring($user.UserPrincipalName.LastIndexOf("_") + 1, $user.UserPrincipalName.IndexOf("#") - $user.UserPrincipalName.LastIndexOf("_") -1)
		$userType = "External Account Supplier - $companyName"
	}

    $licenseObject = New-Object -TypeName PSObject
    $licenseObject | Add-Member -MemberType NoteProperty -Name UserPrincipalname -Value $user.UserPrincipalname
    $licenseObject | Add-Member -MemberType NoteProperty -Name DisplayName -Value $user.DisplayName
    $licenseObject | Add-Member -MemberType NoteProperty -Name Domain -Value $user.UserPrincipalname.Split('@')[1]
    $licenseObject | Add-Member -MemberType NoteProperty -Name UserType -Value $userType

    Foreach($consumedLicense in $consumedLicenses)
    {
        $friendlyLicenseName = ($licensingPlans | ? { $_ -match $consumedLicense.SkuID })

        if($friendlyLicenseName.Count -gt 0)
        {
            $friendlyLicenseName = $friendlyLicenseName[0].Product_Display_Name
        } else {
            $friendlyLicenseName = $friendlyLicenseName.Product_Display_Name
        }

        if($consumedLicense.SkuId -eq 'e0dfc8b9-9531-4ec8-94b4-9fec23b05fc8')
        {
            $friendlyLicenseName = "Microsoft Teams Exploratory"
        }
        

        if($userLicense.skuId -contains $consumedLicense.SkuID)
        {
            $licenseObject | Add-Member -MemberType NoteProperty -Name $friendlyLicenseName -Value "Yes"
        } else {
            $licenseObject | Add-Member -MemberType NoteProperty -Name $friendlyLicenseName -Value "No"
        }
    }

    $licenseAssignments += $licenseObject
}

Foreach($license in $consumedLicenses)
{
    $friendlyLicenseName = ($licensingPlans | ? { $_ -match $license.SkuID })
    if($friendlyLicense.Count -gt 0)
    {
        $friendlyLicenseName = $friendlyLicenseName[0].Product_Display_Name
    } else {
        $friendlyLicenseName = $friendlyLicenseName.Product_Display_Name
    }

    if($license.SkuId -eq 'e0dfc8b9-9531-4ec8-94b4-9fec23b05fc8')
    {
        $friendlyLicenseName = "Microsoft Teams Exploratory"
    }

    $hashObject = [ordered]@{ LicenseName = $friendlyLicenseName; AccountSkuID = $license.SkuID; ActiveUnits = $license.PrepaidUnits.Enabled; WarningUnits = $license.PrepaidUnits.Warning; ConsumedUnits = $license.ConsumedUnits}
    $outputObject = New-Object -TypeName PSObject -Property $hashObject
    $licenseOverview += $outputObject
}

$licenseAssignments | Export-Excel $outputLicenseAssignments -FreezeTopRow -TitleBold -AutoFilter -AutoSize
$licenseOverview    | Export-Excel $outputLicenseOverview -FreezeTopRow -TitleBold -AutoFilter -AutoSize