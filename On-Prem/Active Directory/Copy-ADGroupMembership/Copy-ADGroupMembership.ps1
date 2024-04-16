<#
.Synopsis
   This function enables you in easily copying groups from an Active Directory User to another Active Directory User
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Copy-ADGroupMembership
{
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $sourceAccount,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $destinationAccount,
        $domainController,
        $secureCredential,
        $excludedGroups
    )

    if(($null -eq $domainController) -and ($null -eq $secureCredential))
    {
        Get-ADUser -Identity $sourceAccount | Select-Object -ExpandProperty MemberOf | Get-ADGroup | ForEach-Object {
            if(!($excludedGroups.Contains($_.Name)))
            {
                Add-ADGroupMember -Identity $_.SID.Value -Members $destinationAccount
            }
        }
    }

    if(($null -eq $domainController) -and ($null -ne $secureCredential))
    {
        Get-ADUser -Identity $sourceAccount -Credential $secureCredential | Select-Object -ExpandProperty MemberOf | Get-ADGroup -Credential $secureCredential | ForEach-Object {
            if(!($excludedGroups.Contains($_.Name)))
            {
                Add-ADGroupMember -Identity $_.SID.Value -Members $destinationAccount -Credential $secureCredential
            }
        }
    }

    if(($null -ne $domainController) -and ($null -eq $secureCredential))
    {
        Get-ADUser -Identity $sourceAccount -Server $domainController | Select-Object -ExpandProperty MemberOf | Get-ADGroup -Server $domainController | ForEach-Object {
            if(!($excludedGroups.Contains($_.Name)))
            {
                Add-ADGroupMember -Identity $_.SID.Value -Members $destinationAccount -Server $domainController
            }
        }
    }

    if(($null -ne $domainController) -and ($null -ne $secureCredential))
    {
        Get-ADUser -Identity $sourceAccount -Server $domainController -Credential $secureCredential | Select-Object -ExpandProperty MemberOf | Get-ADGroup -Server $domainController -Credential $secureCredential | ForEach-Object {
            if(!($excludedGroups.Contains($_.Name)))
            {
                Add-ADGroupMember -Identity $_.SID.Value -Members $destinationAccount -Server $domainController -Credential $secureCredential
            }
        }
    }
    
}

