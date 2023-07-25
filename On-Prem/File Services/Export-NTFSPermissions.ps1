Param(
    [Parameter(Mandatory=$true)]
    [string]$OutputPath
)

#region Verifying Output Path
    Do{
        if((Test-Path -Path $OutputPath) -eq $false)
        {
            $NTFSPermissionsOutputPath = Read-Host -Prompt "Please enter a valid path for the NTFS Permissions output"
        }
    } while (((Test-Path -Path $OutputPath) -eq $false))

    if(($OutputPath.EndsWith('\')) -eq $false)
    {
        $OutputPath += '\'
    }
#endregion
#region Modules check
    $foundModules = Get-Module -ListAvailable
    $NTFSModule = $foundModules | ? { $_.Name -match 'NTFSSecurity' }
    $ExcelModule = $foundModules | ? { $_.Name -match 'ImportExcel' }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if($null -eq $ExcelModule)
    {
        Install-Module -Name ImportExcel -Force -AllowClobber -Confirm:$false
    }

    If($null -eq $NTFSModule)
    { 
        Install-Module -Name NTFSSecurity -RequiredVersion 4.2.4 -Force -AllowClobber -Confirm:$false
    }

    Import-Module -Name NTFSSecurity, ImportExcel
#endregion
#region Starting Variables
    $accessCollection = New-Object System.Collections.Generic.List[PSObject]
    $targetDirectory = 'C:\'
#endregion

$childFolders = Get-ChildItem2 -Path $targetDirectory -Directory -Recurse -ErrorAction SilentlyContinue

$progressCount = 0
Foreach($folder in $childFolders)
{
    $progressCount++; 'Processing {0} out of {1}' -f $progressCount, $childFolders.Count | Write-Host
    if((Test-Path2 $folder.FullName -ErrorAction SilentlyContinue) -eq $true)
    {
        if($folder.FullName.Length -lt 250)
        {
            $ACL = Get-ACL $folder.FullName
            Foreach($access in $ACL.Access)
            {
                if(($access.IsInherited -eq $false) -and ($access.IdentityReference -notmatch 'S-1-15-') -and ($access.IdentityReference -notmatch 'S-1-5-'))
                {
                    $hashObject = [ordered]@{FullName = $folder.FullName; PathName = $folder.Name; Identity = $access.IdentityReference; Permission = $access.FileSystemRights; AccessControlType = $access.AccessControlType; InheritanceFlags = $access.InheritanceFlags}
                    $outputObject = New-Object -TypeName PSObject -Property $hashObject
                    $accessCollection.Add($outputObject)
                }
            }
        }
    }
    
}

$groupMemberCollection = New-Object System.Collections.Generic.List[PSObject]
$uniqueGroupList = $accessCollection.identity | Sort-Object -Unique
Foreach($entry in $uniqueGroupList)
{
    Try{
        $ADGroup = $null
        $groupName = $entry.Value.Split('\')[1]
        $ADGroup = Get-ADGroup -Identity $groupName -ErrorAction Stop
        If($ADGroup)
        {
            $ADGroupMembers = Get-ADGroupMember -Identity $groupName
            $ADGroupMembers | ? { $_.objectClass -eq 'user' } | ForEach-Object {
                $ADUser = Get-ADuser -Identity $_.SamAccountName -Properties UserPrincipalName, EmailAddress, Enabled
                if($ADUser.EmailAddress.Length -gt 1)
                {
                    $hashObject = [ordered]@{ GroupName = $groupName; EmailAddress = $ADUser.EmailAddress}
                    $outputObject = New-Object -TypeName PSObject -Property $hashObject
                    $groupMemberCollection.Add($outputObject)
                }
            }
        }
    } catch{
    }
}

$accessCollection | Export-Excel -Path "$($OutputPath)_NTFSPermissions.xlsx" -FreezeTopRow -TitleBold -AutoFilter -AutoSize
$groupMemberCollection | Export-Excel -Path "$($OutputPath)_Groupmembers.xlsx" -FreezeTopRow -TitleBold -AutoFilter -AutoSize